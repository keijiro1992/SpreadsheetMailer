function sendEmails() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("顧客リスト");
    var resultSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("送信結果");
    var headers = sheet.getDataRange().getValues()[0];
    var data = sheet.getDataRange().getValues().slice(1);
  
    var columnIndex = {
      name: headers.indexOf("名前"),
      email: headers.indexOf("メールアドレス"),
      subject: headers.indexOf("件名"),
      senderAddress: headers.indexOf("送信元アドレス"),
      body: headers.indexOf("メール本文"),
      flag: headers.indexOf("フラグ")
    };
  
    data.forEach(function(row) {
      if (row[columnIndex.flag] == true) {
        try {
          MailApp.sendEmail({
            to: row[columnIndex.email],
            subject: row[columnIndex.subject],
            body: row[columnIndex.body],
            name: "送信者名",
            from: row[columnIndex.senderAddress]
          });
          resultSheet.appendRow([row[columnIndex.name], row[columnIndex.email], "成功"]);
        } catch(e) {
          resultSheet.appendRow([row[columnIndex.name], row[columnIndex.email], "失敗: " + e.message]);
        }
      }
    });
  }