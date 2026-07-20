const SHEET_NAME = "Form Responses 1";

const START_HOUR = 7;
const START_MINUTE = 0;

//====================================================
// WEB APP
//====================================================

function doGet() {

  return HtmlService
    .createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Dashboard Điểm Danh");

}

//====================================================
// INCLUDE HTML
//====================================================

function include(filename) {

  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();

}

//====================================================
// GET ATTENDANCE
//====================================================

function getAttendance() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(SHEET_NAME);

  if (!sheet) {

    return {
      total: 0,
      onTime: 0,
      latePermission: 0,
      lateNoPermission: 0,
      absentPermission: 0,
      absentNoPermission: 0,
      list: [],
      classSummary: []
    };

  }

  let data = sheet
    .getDataRange()
    .getValues();

  if (data.length <= 1) {

    return {
      total: 0,
      onTime: 0,
      latePermission: 0,
      lateNoPermission: 0,
      absentPermission: 0,
      absentNoPermission: 0,
      list: [],
      classSummary: []
    };

  }

  //=========================================
  // Bỏ dòng tiêu đề
  //=========================================

  data.shift();

  const tz = Session.getScriptTimeZone();

  const today = Utilities.formatDate(

    new Date(),

    tz,

    "yyyyMMdd"

  );

  //=========================================
  // Chỉ lấy dữ liệu hôm nay
  //=========================================

  data = data.filter(function(row){

    if(!row[0]) return false;

    return Utilities.formatDate(

      new Date(row[0]),

      tz,

      "yyyyMMdd"

    ) === today;

  });

  //=========================================
  // Mới nhất -> cũ nhất
  //=========================================

  data.sort(function(a,b){

    return new Date(b[0]) - new Date(a[0]);

  });

  const limit =
    START_HOUR * 60 + START_MINUTE;

  //=========================================
  // Biến thống kê
  //=========================================

  let attendanceList = [];

  let classStats = {};

  let onTime = 0;

  let latePermission = 0;

  let lateNoPermission = 0;

  let absentPermission = 0;

  let absentNoPermission = 0;

  //=========================================
  // Duyệt từng bản ghi
  //=========================================

  data.forEach(function(row){

    const time = new Date(row[0]);

    const minute =
      time.getHours() * 60 +
      time.getMinutes();

    const name =
      String(row[2]).trim();

    const className =
      String(row[3]).trim() +
      " " +
      String(row[4]).trim();

    const lateText =
      String(row[5] || "").trim();

    const absentText =
      String(row[6] || "").trim();

    //======================================
    // Khởi tạo thống kê lớp
    //======================================

    if(!classStats[className]){

      classStats[className]={

        className:className,

        onTime:0,

        latePermission:0,

        lateNoPermission:0,

        absentPermission:0,

        absentNoPermission:0

      };

    }

    let status = "";

    let lateMinute = 0;

        //======================================
    // Xác định trạng thái điểm danh
    //======================================

    // 1. Vắng có phép
    if (absentText === "Có") {

      status = "Vắng có phép";

      absentPermission++;

      classStats[className].absentPermission++;

    }

    // 2. Vắng không phép
    else if (absentText === "Không") {

      status = "Vắng không phép";

      absentNoPermission++;

      classStats[className].absentNoPermission++;

    }

    //======================================
    // Có mặt
    //======================================

    else {

      // Đi trễ
      if (minute > limit) {

        lateMinute = minute - limit;

        // Trễ có phép
        if (lateText === "Có") {

          status = "Trễ có phép";

          latePermission++;

          classStats[className].latePermission++;

        }

        // Trễ không phép
        else {

          status = "Trễ không phép";

          lateNoPermission++;

          classStats[className].lateNoPermission++;

        }

      }

      // Đúng giờ
      else {

        status = "Đúng giờ";

        onTime++;

        classStats[className].onTime++;

      }

    }

    //======================================
    // Lưu danh sách hiển thị
    //======================================

    attendanceList.push({

      timestamp: time.getTime(),

      date: Utilities.formatDate(
        time,
        tz,
        "dd/MM/yyyy"
      ),

      time: Utilities.formatDate(
        time,
        tz,
        "HH:mm:ss"
      ),

      name: name,

      className: className,

      status: status,

      lateMinute: lateMinute,

      latePermission: lateText,

      absentPermission: absentText

    });

  });

  //======================================
  // Chuyển thống kê lớp sang mảng
  //======================================

  const classSummary = Object
    .values(classStats)
    .sort(function(a, b) {

      return a.className.localeCompare(
        b.className,
        "vi"
      );

    });

      //======================================
  // Kết quả trả về
  //======================================

  const result = {

    // Tổng số lượt điểm danh hôm nay
    total: attendanceList.length,

    // Thống kê chung
    onTime: onTime,

    latePermission: latePermission,

    lateNoPermission: lateNoPermission,

    absentPermission: absentPermission,

    absentNoPermission: absentNoPermission,

    // Danh sách điểm danh
    list: attendanceList,

    // Thống kê theo lớp
    classSummary: classSummary

  };

  return JSON.parse(JSON.stringify(result));

}