const RECIPIENT_EMAIL = "nepalt12345@gmail.com";
const MAX_DAILY_SENDS = 10;
const MAX_ATTACHMENT_BASE64_LENGTH = 20 * 1024 * 1024;

function doPost(event) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    const payload = JSON.parse(event.parameter.payload || "{}");
    const sender = Session.getActiveUser().getEmail().trim().toLowerCase();
    if (!sender) {
      throw new Error("Sign in with an authorized Google account to send attendance reports.");
    }

    const shift = String(payload.shift || "");
    const reportDate = String(payload.reportDate || "");
    const fileName = String(payload.fileName || "");
    const fileBase64 = String(payload.fileBase64 || "");

    if (!/^[12]$/.test(shift) || !/^\d{4}-\d{2}-\d{2}$/.test(reportDate)) {
      throw new Error("The attendance report details are invalid.");
    }
    if (!new RegExp(`^irts_${shift}_${reportDate}\\.xlsx$`).test(fileName)) {
      throw new Error("The workbook filename is invalid.");
    }
    if (!fileBase64 || fileBase64.length > MAX_ATTACHMENT_BASE64_LENGTH || !/^[A-Za-z0-9+/]+={0,2}$/.test(fileBase64)) {
      throw new Error("The workbook attachment is missing or too large.");
    }

    const properties = PropertiesService.getUserProperties();
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    const dayKey = `attendance-email-count-${today}`;
    const sentToday = Number(properties.getProperty(dayKey) || 0);
    if (sentToday >= MAX_DAILY_SENDS) {
      throw new Error("The daily attendance email limit has been reached.");
    }

    const attachment = Utilities.newBlob(
      Utilities.base64Decode(fileBase64),
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      fileName
    );
    MailApp.sendEmail({
      to: RECIPIENT_EMAIL,
      subject: `${shift}-р ээлжийн ирцийн тайлан - ${reportDate}`,
      body: "",
      attachments: [attachment]
    });

    properties.setProperty(dayKey, String(sentToday + 1));
    return responseText(`Sent successfully from ${sender} to ${RECIPIENT_EMAIL}.`);
  } catch (error) {
    return responseText(`Could not send the attendance workbook: ${error.message || error}`);
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function responseText(message) {
  return ContentService.createTextOutput(message).setMimeType(ContentService.MimeType.TEXT);
}
