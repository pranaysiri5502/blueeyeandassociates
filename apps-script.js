function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = JSON.parse(e.postData.contents);
  var timestamp = new Date();

  if (data.formType === "consultation") {
    var sheet = ss.getSheetByName("Consultations");
    if (!sheet) sheet = ss.insertSheet("Consultations");
    if (sheet.getRange(1,1).getValue() === "") {
      sheet.appendRow(["Timestamp","Name","Phone","Date","Time Slot","Service","Description","Status","Notes"]);
    }
    sheet.appendRow([
      timestamp,
      data.name || "",
      data.phone || "",
      data.date || "",
      data.timeSlot || "",
      data.service || "",
      data.description || "",
      "New",
      ""
    ]);
  } else if (data.formType === "enquiry") {
    var sheet = ss.getSheetByName("Enquiries");
    if (!sheet) sheet = ss.insertSheet("Enquiries");
    if (sheet.getRange(1,1).getValue() === "") {
      sheet.appendRow(["Timestamp","Name","Phone","Email","Interested In","Message","Status","Notes"]);
    }
    sheet.appendRow([
      timestamp,
      data.name || "",
      data.phone || "",
      data.email || "",
      data.interestedIn || "",
      data.message || "",
      "New",
      ""
    ]);
  }

  return ContentService.createTextOutput(JSON.stringify({ result: "success" }))
    .setMimeType(ContentService.MimeType.JSON);
}
