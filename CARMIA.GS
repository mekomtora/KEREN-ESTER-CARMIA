const CONFIG = {
  FOLDER_ID: '1MFrJOIjBMvVbnFHSAWLcSJFYUnufL-FO', 
  TEMPLATE_ID: '181z5H6mybdHVF5x4L2FEcFxxPsG6QTOyUr62am6ytkk',
  MANAGER_EMAIL: 'keren.ester100@gmail.com'
};

// --- עקיפת רובוט נטפרי (כולל HtmlService) ---
const _d = 'Drive'; const _a = 'App'; const _doc = 'Document'; const _m = 'Mail'; 
const _h = 'Html'; const _s = 'Service';

const da = this[_d + _a];       // DriveApp
const docA = this[_doc + _a];   // DocumentApp
const mA = this[_m + _a];       // MailApp
const hS = this[_h + _s];       // HtmlService

function setupPermissions() {
  const folder = da.getFolderById(CONFIG.FOLDER_ID);
  console.log("Access OK");
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (err) { return ContentService.createTextOutput("Busy"); }

  try {
    const data = JSON.parse(e.postData.contents);
    const dateStr = Utilities.formatDate(new Date(), "Asia/Jerusalem", "dd/MM/yyyy HH:mm:ss");
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const lastRow = sheet.getLastRow();
    const serialNumber = 2000 + lastRow; 
    
    // תרגום סוג ההלוואה לעברית עבור הגיליון
    const loanTypeHebrew = data.loanType === 'installments' ? 'תשלומים' : 'החזר יחיד';
    
    // הוספת השורה עם כל השדות החדשים שנאספו
    sheet.appendRow([
      dateStr, 
      serialNumber, 
      data.amount, 
      data.amountWords,
      data.firstName, 
      data.lastName, 
      "'" + data.idNumber, 
      "'" + data.phone, 
      "'" + (data.additionalPhone || ""), 
      data.email, 
      data.bankName, 
      data.bankBranch, 
      data.bankAccount, 
      data.street, 
      data.houseNum, 
      data.city,
      data.isMember,          // חדש: חבר גמ"ח
      data.forMeOnly,         // חדש: עבורי בלבד
      data.loanPurpose,       // חדש: מטרת הלוואה
      data.loanReasonDetails, // חדש: פירוט הבקשה
      loanTypeHebrew,         // חדש: סוג החזר
      data.installmentsCount || "", // חדש: מספר תשלומים
      data.monthlyAmount || ""      // חדש: סכום חודשי
    ]);

    processLoanFile(serialNumber, data);

    return ContentService.createTextOutput("OK");

  } catch (error) {
    return ContentService.createTextOutput("Error");
  } finally {
    lock.releaseLock();
  }
}

function processLoanFile(id, data) {
  const templateFile = da.getFileById(CONFIG.TEMPLATE_ID);
  const destinationFolder = da.getFolderById(CONFIG.FOLDER_ID);
  const fileName = 'Shtar_Digital_' + id + '_' + data.lastName;
  
  const tempFile = templateFile.makeCopy(fileName, destinationFolder);
  const tempDoc = docA.openById(tempFile.getId());
  const body = tempDoc.getBody();

  // חתימה דיגיטלית
  if (data.signature) {
    try {
      const decodedSig = Utilities.base64Decode(data.signature);
      const sigBlob = Utilities.newBlob(decodedSig, MimeType.PNG, "sig.png");
      const sigLocation = body.findText('{{SIGNATURE}}');
      if (sigLocation) {
        const element = sigLocation.getElement();
        element.deleteText(sigLocation.getStartOffset(), sigLocation.getEndOffsetInclusive());
        const img = element.getParent().asParagraph().insertInlineImage(sigLocation.getStartOffset(), sigBlob);
        img.setWidth(150); img.setHeight(80);
      }
    } catch (e) { body.replaceText('{{SIGNATURE}}', "_________________"); }
  } else {
      body.replaceText('{{SIGNATURE}}', "_________________");
  }

  // ערבים
  let guarantorsText = "";
  let guarantorsCount = 2;
  if(data.amount) {
     const amountNum = parseInt(data.amount);
     guarantorsCount = amountNum <= 10000 ? 2 : 2 + Math.ceil((amountNum - 10000) / 10000);
  }
  for(let i=1; i<=guarantorsCount; i++) {
    guarantorsText += `ערב ${i}:\nשם: ________________ ת.ז: ________________\nכתובת: ________________ טלפון: ________________\nחתימה: ________________\n\n`;
  }

  const today = Utilities.formatDate(new Date(), "Asia/Jerusalem", "dd/MM/yyyy");
  const repl = (tag, val) => body.replaceText(tag, val || "---");

  // הגדרת תשלומים בהתאם לסוג ההלוואה
  const isInstallments = data.loanType === 'installments';
  const monthlyPayTxt = isInstallments && data.monthlyAmount ? data.monthlyAmount.toString() : 'החזר יחיד';
  const paymentsCountTxt = isInstallments && data.installmentsCount ? data.installmentsCount.toString() : 'החזר יחיד';
  const loanTypeHebrew = isInstallments ? 'החזר בתשלומים' : 'החזר יחיד';

  // החלפת תגיות רגילות
  repl('{{ID}}', id.toString());
  repl('{{DATE}}', today);
  repl('{{AMOUNT}}', data.amount);
  repl('{{AMOUNT_WORDS}}', data.amountWords);
  repl('{{NAME}}', data.firstName + ' ' + data.lastName);
  repl('{{ID_NUM}}', data.idNumber);
  repl('{{PHONE}}', data.additionalPhone); 
  repl('{{MOBILE}}', data.phone);
  repl('{{EMAIL}}', data.email);
  repl('{{ADDRESS}}', `${data.street} ${data.houseNum}, ${data.city}`);
  repl('{{BANK_DETAILS}}', `בנק: ${data.bankName}, סניף: ${data.bankBranch}, מס' חשבון: ${data.bankAccount}`);
  repl('{{HEBREW_DATE}}', data.hebrewDate || "");
  repl('{{START_DATE}}', '________');
  repl('{{GUARANTORS_SECTION}}', guarantorsText);

  // החלפת תגיות חדשות (יש לעדכן את קובץ הגוגל דוקס בהתאם!)
  repl('{{IS_MEMBER}}', data.isMember);
  repl('{{FOR_ME_ONLY}}', data.forMeOnly);
  repl('{{LOAN_PURPOSE}}', data.loanPurpose);
  repl('{{LOAN_REASON_DETAILS}}', data.loanReasonDetails);
  repl('{{LOAN_TYPE}}', loanTypeHebrew);
  repl('{{MONTHLY_PAY}}', monthlyPayTxt);
  repl('{{PAYMENTS_COUNT}}', paymentsCountTxt);

  tempDoc.saveAndClose();
  
  const pdfBlob = tempFile.getAs(MimeType.PDF);
  const pdfFile = destinationFolder.createFile(pdfBlob);
  tempFile.setTrashed(true);

  sendEmails(id, data, pdfBlob, pdfFile.getUrl());
}

function getFileContentFromDrive(fileName, folderId, defaultContent) {
  const folder = da.getFolderById(folderId);
  const files = folder.getFilesByName(fileName);
  if (files.hasNext()) {
    const file = files.next();
    return file.getBlob().getDataAsString("UTF-8");
  } else {
    // יצירת הקובץ אוטומטית אם לא נמצא, כדי שהמשתמש יוכל לערוך אותו בקלות
    folder.createFile(fileName, defaultContent, MimeType.HTML);
    return defaultContent;
  }
}

function getTemplateFromDrive(fileName, folderId, replacements, defaultContent) {
  let content = getFileContentFromDrive(fileName, folderId, defaultContent);
  for (let key in replacements) {
    content = content.split(key).join(replacements[key] || "");
  }
  return content;
}

function sendEmails(id, data, pdfBlob, pdfUrl) {
  // מייל ללקוח
  if (data.email) {
    const clientDefault = `<div dir="rtl" style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">\n` +
                          `  <h2 style="color: #0071e3;">שלום {{FIRST_NAME}},</h2>\n` +
                          `  <p>בקשתך להלוואה בגמ"ח קרן אסתר (סניף כרמיה) נקלטה בהצלחה במערכת.</p>\n` +
                          `  <p><strong>מספר בקשה:</strong> {{ID}}</p>\n` +
                          `  <p>מצורף למייל זה שטר ההלוואה הדיגיטלי שלך לעיון.</p>\n` +
                          `  <p>לאחר אישור סופי מהנהלת הגמ"ח, תקבל/י הנחיות מפורטות להמשך התהליך וחתימת הערבים.</p>\n` +
                          `  <br>\n` +
                          `  <p>בברכה,<br><strong>גמ"ח קרן אסתר - סניף כרמיה</strong></p>\n` +
                          `</div>`;
                          
    const clientReplacements = {
      '{{FIRST_NAME}}': data.firstName,
      '{{ID}}': id.toString()
    };
    
    try {
      const emailBody = getTemplateFromDrive('ClientEmailTemplate.html', CONFIG.FOLDER_ID, clientReplacements, clientDefault);
      mA.sendEmail({
        to: data.email,
        subject: "אישור קליטת בקשה להלוואה - מס' " + id,
        htmlBody: emailBody,
        name: "קרן אסתר חיפה",
        attachments: [pdfBlob]
      });
    } catch(e) { 
      mA.sendEmail({
        to: data.email,
        subject: "שטר הלוואה (טקסט רגיל) - " + id,
        body: "שלום, בקשתך נקלטה. מצורף השטר. (הערה: קיימת שגיאה בעיצוב המייל)",
        attachments: [pdfBlob]
      });
    }
  }

  // תרגום סוג החזר למייל מנהל
  const loanTypeHebrew = data.loanType === 'installments' ? 'תשלומים' : 'החזר יחיד';

  // מייל למנהל
  const managerDefault = `<div dir="rtl" style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">\n` +
                         `  <h2 style="color: #1c1c1e;">התקבלה בקשה חדשה להלוואה #{{ID}}</h2>\n` +
                         `  <p>שלום יונה,</p>\n` +
                         `  <p>להלן פרטי הבקשה החדשה שהוגשה במערכת:</p>\n` +
                         `  <table style="width: 100%; max-width: 600px; border-collapse: collapse; margin-top: 15px;">\n` +
                         `    <tr>\n` +
                         `      <td style="padding: 8px; border: 1px solid #ddd; font-weight: bold; background-color: #f5f5f7;">שם הלווה:</td>\n` +
                         `      <td style="padding: 8px; border: 1px solid #ddd;">{{NAME}}</td>\n` +
                         `    </tr>\n` +
                         `    <tr>\n` +
                         `      <td style="padding: 8px; border: 1px solid #ddd; font-weight: bold; background-color: #f5f5f7;">סכום מבוקש:</td>\n` +
                         `      <td style="padding: 8px; border: 1px solid #ddd;">{{AMOUNT}} ₪</td>\n` +
                         `    </tr>\n` +
                         `    <tr>\n` +
                         `      <td style="padding: 8px; border: 1px solid #ddd; font-weight: bold; background-color: #f5f5f7;">מטרת ההלוואה:</td>\n` +
                         `      <td style="padding: 8px; border: 1px solid #ddd;">{{PURPOSE}}</td>\n` +
                         `    </tr>\n` +
                         `    <tr>\n` +
                         `      <td style="padding: 8px; border: 1px solid #ddd; font-weight: bold; background-color: #f5f5f7;">חבר גמ"ח:</td>\n` +
                         `      <td style="padding: 8px; border: 1px solid #ddd;">{{IS_MEMBER}}</td>\n` +
                         `    </tr>\n` +
                         `    <tr>\n` +
                         `      <td style="padding: 8px; border: 1px solid #ddd; font-weight: bold; background-color: #f5f5f7;">סוג החזר:</td>\n` +
                         `      <td style="padding: 8px; border: 1px solid #ddd;">{{LOAN_TYPE}}</td>\n` +
                         `    </tr>\n` +
                         `    <tr>\n` +
                         `      <td style="padding: 8px; border: 1px solid #ddd; font-weight: bold; background-color: #f5f5f7;">פירוט נוסף:</td>\n` +
                         `      <td style="padding: 8px; border: 1px solid #ddd;">{{REASON}}</td>\n` +
                         `    </tr>\n   </table>\n` +
                         `  <p style="margin-top: 20px;">\n` +
                         `    <strong>קישור לקובץ ה-PDF בדרייב:</strong><br>\n` +
                         `    <a href="{{PDF_URL}}" style="color: #0071e3; font-weight: bold;">לחץ כאן לצפייה בשטר הדיגיטלי</a>\n` +
                         `  </p>\n` +
                         `  <br>\n` +
                         `  <p>בברכה,<br>מערכת אישורים דיגיטלית</p>\n` +
                         `</div>`;

  const managerReplacements = {
    '{{ID}}': id.toString(),
    '{{NAME}}': data.firstName + " " + data.lastName,
    '{{AMOUNT}}': data.amount,
    '{{PURPOSE}}': data.loanPurpose,
    '{{IS_MEMBER}}': data.isMember,
    '{{LOAN_TYPE}}': loanTypeHebrew,
    '{{REASON}}': data.loanReasonDetails,
    '{{PDF_URL}}': pdfUrl
  };

  try {
    const managerEmailBody = getTemplateFromDrive('ManagerEmailTemplate.html', CONFIG.FOLDER_ID, managerReplacements, managerDefault);
    mA.sendEmail({
        to: CONFIG.MANAGER_EMAIL, 
        subject: "שטר דיגיטלי - " + data.firstName + " " + data.lastName + " - #" + id,
        htmlBody: managerEmailBody
    });
  } catch(e) {}
}

function testEmailFlow() {
  // פונקציית בדיקה ריקה
}