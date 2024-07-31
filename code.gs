//Nama-Nama Kolom 
const RECIPIENT_COL = "Email";
const EMAIL_SENT_COL = "Email Sent";
const STATUS_COL = "Email Track Status";
const SENDER_NAME_COL = "Sender Name";

//API_URL dari deployment Track.gs
const TRACK_API_URL = "https://script.google.com/macros/s/AKfycbx2E8sgeOoEDxtvZNFTkwI0jLlCpXQkENP3wxit8IHozISEpoqtyTQUM7STjpr9_uFY/exec";

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Mail Merge by OMI')
    .addItem('Blast Emails', 'sendEmails')
    .addToUi();
}

function sendEmails(subjectLine, sheet = SpreadsheetApp.getActiveSheet()) {
  if (!subjectLine) {
    subjectLine = Browser.inputBox("Mail Merge by O. Midiyanto", 
                                      "Enter the subject draft from the Gmail: " ,
                                      Browser.Buttons.OK_CANCEL);
                                      
    if (subjectLine === "cancel" || subjectLine == "") { 
      return;
    }
  }
  // Bersihkan kolom "Email Sent" dan "Status"
  clearEmailSentColumn(sheet);
  clearStatusColumn(sheet);
  
  const emailTemplate = getGmailTemplateFromDrafts_(subjectLine);
  
  const dataRange = sheet.getDataRange();
  const data = dataRange.getDisplayValues();

  const heads = data.shift(); 
  
  const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);
  const recipientColIdx = heads.indexOf(RECIPIENT_COL);
  const statusColIdx = heads.indexOf(STATUS_COL);
  const senderNameColIdx = heads.indexOf(SENDER_NAME_COL);
  
  const obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));

  const out = [];
  
  obj.forEach(function(row, rowIdx){
    if (row[EMAIL_SENT_COL] == ''){
      try {
        const msgObj = fillInTemplateFromObject_(emailTemplate.message, row);

        // Embed pixel tracking
        const pixelUrl = `${TRACK_API_URL}?recipient=${encodeURIComponent(row[RECIPIENT_COL])}`;
        const trackedHtmlBody = `${msgObj.html}<img src="${pixelUrl}" style="display:none;">`;
        const senderName = sheet.getRange(rowIdx + 2, senderNameColIdx + 1).getValue(); // Get sender name from the row
        const msgOptions = {
          htmlBody: trackedHtmlBody,
          attachments: emailTemplate.attachments,
          inlineImages: emailTemplate.inlineImages,
          name: senderName // Menggunakan sender name dari kolom yang sesuai
        };

        GmailApp.sendEmail(row[RECIPIENT_COL], msgObj.subject, msgObj.text, msgOptions);
        out.push([new Date()]);
        
        // Tandai timestamp pengiriman email
        var timestamp = new Date();
        sheet.getRange(rowIdx + 2, emailSentColIdx + 1).setValue(timestamp);
        
      } catch(e) {
        out.push([e.message]);
      }
    } else {
      out.push([row[EMAIL_SENT_COL]]);
    }
  });
  
  sheet.getRange(2, emailSentColIdx + 1, out.length).setValues(out);
}

function clearEmailSentColumn(sheet) {
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  const heads = data[0];
  const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL) + 1; // Menambahkan 1 karena getRange menggunakan 1-based index
  
  if (emailSentColIdx > 0) {
    const range = sheet.getRange(2, emailSentColIdx, sheet.getLastRow() - 1);
    range.clearContent();
  }
}

function clearStatusColumn(sheet) {
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  const heads = data[0];
  const statusColIdx = heads.indexOf(STATUS_COL) + 1; // Menambahkan 1 karena getRange menggunakan 1-based index
  
  if (statusColIdx > 0) {
    const range = sheet.getRange(2, statusColIdx, sheet.getLastRow() - 1);
    range.clearContent();
  }
}

function getGmailTemplateFromDrafts_(subject_line){
  try {
    const drafts = GmailApp.getDrafts();
    const draft = drafts.filter(subjectFilter_(subject_line))[0];
    const msg = draft.getMessage();

    const allInlineImages = draft.getMessage().getAttachments({includeInlineImages: true,includeAttachments:false});
    const attachments = draft.getMessage().getAttachments({includeInlineImages: false});
    const htmlBody = msg.getBody(); 

    const img_obj = allInlineImages.reduce((obj, i) => (obj[i.getName()] = i, obj) ,{});

    const imgexp = RegExp('<img.*?src="cid:(.*?)".*?alt="(.*?)"[^\>]+>', 'g');
    const matches = [...htmlBody.matchAll(imgexp)];

    const inlineImagesObj = {};
    matches.forEach(match => inlineImagesObj[match[1]] = img_obj[match[2]]);

    return {message: {subject: subject_line, text: msg.getPlainBody(), html:htmlBody}, 
            attachments: attachments, inlineImages: inlineImagesObj };
  } catch(e) {
    throw new Error("Oops - can't find Gmail draft");
  }
}

function subjectFilter_(subject_line){
  return function(element) {
    if (element.getMessage().getSubject() === subject_line) {
      return element;
    }
  }
}

function fillInTemplateFromObject_(template, data) {
  let template_string = JSON.stringify(template);
  template_string = template_string.replace(/{{[^{}]+}}/g, key => {
    return escapeData_(data[key.replace(/[{}]+/g, "")] || "");
  });
  return  JSON.parse(template_string);
}

function escapeData_(str) {
  return str
    .replace(/[\\]/g, '\\\\')
    .replace(/[\"]/g, '\\\"')
    .replace(/[\/]/g, '\\/')
    .replace(/[\b]/g, '\\b')
    .replace(/[\f]/g, '\\f')
    .replace(/[\n]/g, '\\n')
    .replace(/[\r]/g, '\\r')
    .replace(/[\t]/g, '\\t');
}
