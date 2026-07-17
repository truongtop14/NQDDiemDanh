const SHEET_NAME = "Form Responses 1";

const START_HOUR = 7;
const START_MINUTE = 0;

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile("Index")
    .setTitle("Dashboard Điểm Danh");
}

function getAttendance() {

  const sh = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(SHEET_NAME);

  const data = sh.getDataRange().getValues();

  // Bỏ dòng tiêu đề
  data.shift();

  // Sắp xếp Timestamp mới nhất -> cũ nhất
  data.sort((a, b) => new Date(b[0]) - new Date(a[0]));

  let result = [];

  let onTime = 0;
  let late = 0;

  const limit = START_HOUR * 60 + START_MINUTE;

  data.forEach(r => {

    const t = new Date(r[0]);

    const minute = t.getHours() * 60 + t.getMinutes();

    let status = "Đúng giờ";
    let lateMinute = 0;

    if (minute > limit) {
      status = "Đi trễ";
      lateMinute = minute - limit;
      late++;
    } else {
      onTime++;
    }

    result.push({

      timestamp: t.getTime(),

      date: Utilities.formatDate(
        t,
        Session.getScriptTimeZone(),
        "dd/MM/yyyy"
      ),

      time: Utilities.formatDate(
        t,
        Session.getScriptTimeZone(),
        "HH:mm:ss"
      ),

      name: r[2],
      className: r[3],

      status: status,
      lateMinute: lateMinute

    });

  });

  return {

    total: data.length,

    onTime: onTime,

    late: late,

    list: result

  };

}