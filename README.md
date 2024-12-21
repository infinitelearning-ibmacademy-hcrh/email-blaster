# TUTORIAL PENGGUNAAN SCRIPT BLAST EMAIL VIA GOOGLE SHEETS
## Limit/Batasan
1. **WAJIB** menggunakan Gmail sebagai backend untuk blast nya
2. Sekali pakai script hanya bisa blast 70-90 email per harinya

## Persiapan
1. Siapkan akun Gmail aktif (gunakan akun dengan **Foto Profil** dan **Nama Akun** yang professional)
2. Data yang mau diblast pada Google Spreadsheets **minimal** harus memiliki kolom-kolom berikut:
   - Nama Peserta
   - Email
   - Sesuatu yang mau diblast, misalnya seperti **Link Drive Sertifikat** Peserta

## Contoh Studi Kasus Penggunaan Beserta Tahapannya
**STUDI KASUS** : Blast Sertifikat Ke Peserta Melalui Gmail

1. Siapkan data di google sheets dengan fields **minimal** memiliki kolom **"Nama, Email, dan Link Drive Sertifikat"**
![image](https://github.com/user-attachments/assets/5f0e440b-96d8-4bf3-8e02-2b6573acb961)

2. Dalam script ini, juga memiliki fitur **tracking**, oleh karena itu buatkan dua kolom baru yaitu **"Email Sent Date-Time"** dan **"Email Track Status"**. 
![image](https://github.com/user-attachments/assets/6201c48e-9d0e-4052-95ff-6abd41e25b47)

3. Setelahnya, buatkan kolom baru lagi yaitu **"Sender Name"**. Kolom ini bertujuan untuk menentukan nama Sender/Pengirim atau pihak yang melakukan Blast email ini. Sebagai contoh disini diisi dengan nama organisasi yaitu **"Infinite Learning Indonesia"**, sehingga nantinya penerima email akan mengetahui siapa pengirim emailnya. 
![image](https://github.com/user-attachments/assets/7a97f052-3ad1-465e-b631-8c04185a6d4c)

4. Selanjutnya buka tab **Extensions**, lalu pilih **App Script**
![image](https://github.com/user-attachments/assets/160f3ee2-54f6-4d2e-bd7c-b2189d5b73b6)

5. Setelah terbuka, rename nama **"Untitled project"** menjadi **"Blast Email"**. Lalu rename **code.gs** menjadi **track.gs**
![image](https://github.com/user-attachments/assets/aeb72094-072c-44e0-8749-29e6747ee7fa)

6. Isi **track.gs** dengan copy paste kode berikut:
```
const SHEETS_NAME = "Sheet1";
const SPREADSHEET_ID = "1Nxxqz7qRCd1cqm3IGoZZNwIOgUxtKCpcZ3lMjJN0GLg";

function doGet(e) {
  var recipient = e.parameter.recipient;
  
  // Update status "Opened" di Google Sheets jika email penerima membuka email
  if (recipient) {
    updateStatusIfEmailOpened(recipient);
  }
  
  // Memberikan respons pixel (gambar 1x1 piksel transparan)
  var pixelUrl = "data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7";
  return ContentService.createTextOutput(pixelUrl).setMimeType(ContentService.MimeType.GIF);
}

function updateStatusIfEmailOpened(recipient) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEETS_NAME);
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  
  // Get headers to find the column index of "Email"
  var headers = values[0];
  var emailColumnIndex = headers.indexOf(RECIPIENT_COL);

  // Cari baris dengan email penerima yang sesuai
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var emailValue = row[emailColumnIndex];
    if (emailValue && emailValue.toLowerCase() === recipient.toLowerCase()) {
      var columnIndex = getColumnIndexByName(sheet, STATUS_COL);
      if (columnIndex !== -1) {
        sheet.getRange(i + 1, columnIndex + 1).setValue("Opened");
        break;
      }
    }
  }
}

// Function to get column index by name
function getColumnIndexByName(sheet, columnName) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var columnIndex = headers.indexOf(columnName);
  return columnIndex;
}
```

7. Ubah **SHEETS_NAME** dan **SPREADSHEET_ID** sesuai dengan yang anda miliki di data google sheets anda.<br>
    Misal di google sheets saya, yaitu:
    ![Image](https://github.com/user-attachments/assets/043bfcf3-3adf-4511-9e1a-f073c3332004)
    <br>
    Maka, ubah code nya menjadi :
    ```
    const SHEETS_NAME = "Sheet1";
    const SPREADSHEET_ID = "128vEyKx9KnHcGgPPC58fQhd9h9y5XWXtKdwzua7GAP8";
    ```

8. Save code nya, lalu klik Deploy dan pilih **New Deployment**
    ![image](https://github.com/user-attachments/assets/3970f410-5f2f-40d1-9482-091f3b01a8fd)
    <br>

9. Pilih type nya menjadi **Web app**
    <br>
    ![image](https://github.com/user-attachments/assets/a1f6104e-3441-4489-bddc-41de18cbf094)

10. Ubah **Who has access** menjadi **Anyone**. Untuk Description bisa diisi bebas. Lalu klik **Deploy**
    ![image](https://github.com/user-attachments/assets/91326eac-4237-424d-a164-633ee894e1ca)

11. Jika diminta **Authorize access**, klik saja dan beri izin untuk deploy script ini 
    ![image](https://github.com/user-attachments/assets/ef006808-f2f8-4c86-a3ad-3b2d6f1d8edb)
    <br>
    Pilih **Advanced** lalu klik **Go to Blast Email (unsafe)**
    <br>
    ![image](https://github.com/user-attachments/assets/bfeb5703-9117-44bc-8d81-eb5d9e4127a6)

12. Jika sudah, maka **Copy** dan **Simpan** terlebih dahulu **URL** deploymentnya, karena nanti akan digunakan. Jika sudah klik **Done**
![image](https://github.com/user-attachments/assets/06fe66b5-7dc3-487a-9f3b-e0a03da61fe2)

13. Lalu buat **Script** baru dan beri nama file nya **Code**
![image](https://github.com/user-attachments/assets/a6e32054-76ed-44c9-b73d-0a500b385e0d)

14. Isi dengan code berikut:
```
//Nama-Nama Kolom 
const RECIPIENT_COL = "Email";
const EMAIL_SENT_COL = "Email Sent";
const STATUS_COL = "Email Track Status";
const SENDER_NAME_COL = "Sender Name";

//API_URL dari deployment Track.gs
const TRACK_API_URL = "https://script.google.com/macros/saaaaxxxxsaaaassssbbgyqggbasjdha/exec";

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
```

15. Ubah beberapa variabel menyesuaikan dengan data yang kita buat di google sheets
    - Ubah variabel **TRACK_API_URL** menggunakan **URL** deployment yang sebelumnya kita simpan pada **langkah 12**
    ```
    //API_URL dari deployment Track.gs
    const TRACK_API_URL = "https://script.google.com/macros/s/AKfycbwoj6tRwutvsc2zy-mT9Z8w0Q_3VJQfocciToQ-quPaFIW1U_s8e9rLgdX-eR9Qdvcg/exec";
    ```
    - Ubah variabel **RECIPIENT_COL**, **EMAIL_SENT_COL**, **STATUS_COL**, dan **SENDER_NAME_COL** sesuai dengan tabel di google sheets kita. <br>
    Misal di sheets kita seperti ini:
    ![image](https://github.com/user-attachments/assets/5f03d7d7-9229-4ec9-ba01-80e90c138442)
    Maka ubah variabelnya menjadi:
    ```
    //Nama-Nama Kolom 
    const RECIPIENT_COL = "Email Peserta";
    const EMAIL_SENT_COL = "Email Sent Date-Time";
    const STATUS_COL = "Email Track Status";
    const SENDER_NAME_COL = "Sender Name";
    ```


16. Jika sudah, maka save code nya. Kembali ke google sheets, lalu **refresh page** nya. Seharusnya akan muncul menu baru yaitu **Mail Merge by OMI**.
![image](https://github.com/user-attachments/assets/241d0624-21b6-4999-9018-33862216180a)


17. Sebelum menggunakan scriptnya, buka **GMail** terlebih dahulu. Lalu klik **Compose** untuk membuat pesan baru
![image](https://github.com/user-attachments/assets/ca9d9b86-80d2-44e9-ba44-b806e5740eef)

18. Buatlah **Template** email yang akan diblast, sebagai contoh saya ingin blast pesan seperti berikut:
    ![image](https://github.com/user-attachments/assets/61bd758d-b4e4-482e-bb36-0de0ea616d43)
    <br>
    - Lalu **Save Sebagai Draft** (tidak perlu di send), indikasinya ada tertulis **Draft Saved**. 
    - Catat nama **Subject** yang dibuat karena nanti akan digunakan, sebagai contoh disini dibuat dengan **Sertifikat Bootcamp Infinite Learning Indonesia**
    - Untuk format penulisan variabel seperti {{Nama Peserta}}, {{Link Sertifikat}}, dan {{Sender Name}} ini merupakan **nama kolom** yang telah dibuat dari awal pada data google sheets nya.

    **NOTE**: Pesan email ini formatnya ialah HTML, sehingga bisa dibuat menjadi lebih menarik jika diinginkan

19. Kembali ke Google Sheets, lalu jalankan scriptnya dengan klik **Blast Emails**
![image](https://github.com/user-attachments/assets/8a477422-d514-46e7-8d2a-e8ec934859c3)

20. Jika diminta izin akses seperti ini, beri saja izinnya
![image](https://github.com/user-attachments/assets/d177de14-7b5c-44ab-bfaf-ceb3083f3ccb)

21. Isi Subject Email yang diminta dengan Subject yang telah kita buat di **langkah 18**
![image](https://github.com/user-attachments/assets/24b4ba10-71a0-4baa-89e4-c28572cde34d)

22. Klik **OK** untuk start blast email nya.

23. Jika kolom **Email Sent Date-Time** sudah terisi, maka email telah berhasil dikirim, dan seharusnya sudah diterima oleh peserta.
![image](https://github.com/user-attachments/assets/4e942e4c-4534-4562-bebd-91db4c1ab5c7)

24. Contoh email yang diterima:
![image](https://github.com/user-attachments/assets/4fd7c121-8b2f-4522-9d5b-85a31828f843)
![image](https://github.com/user-attachments/assets/d19f3396-b9fc-453b-a4d4-54d9e9fd7934)

**NOTE** Arahkan ke peserta, jika ada pesan "To protect your privacy remote resources have been blocked." maka klik **Allow** agar tidak masuk **SPAM**


## BUGS
- (21-12-2024) Fitur Tracking Tidak Bekerja Semestinya (perlu dilakukan riset tambahan untuk perbaikan)
