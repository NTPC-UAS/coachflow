window.APP_CONFIG = {
  mode: "cloud",
  appsScriptUrl: "https://script.google.com/macros/s/AKfycbxJV_y__nD9nD-6AepZFENxrkQbJv_h3cq59pWaCk2pudGaY5ew5v5WJ_N9zDaj_B7WVg/exec",
  leaveAppsScriptUrl: "https://script.google.com/macros/s/AKfycbxJV_y__nD9nD-6AepZFENxrkQbJv_h3cq59pWaCk2pudGaY5ew5v5WJ_N9zDaj_B7WVg/exec",
  coachflowAppsScriptUrl: "https://script.google.com/macros/s/AKfycbxJV_y__nD9nD-6AepZFENxrkQbJv_h3cq59pWaCk2pudGaY5ew5v5WJ_N9zDaj_B7WVg/exec",
  defaultNotifyEmail: "hsnu115023@gmail.com",
  coachCalendarIds: {
    MO001: "tp15vijlf0p2se3o6j9c4qdvks@group.calendar.google.com"
  },
  productionCoachAccessCodes: ["MO001"],
  calendarStudentAliases: {
    "宇含": "ST013",
    "宇涵": "ST013",
    "宇菡": "ST013",
    "建銘": "ST010",
    "建名": "ST010",
    "亭昀": "ST014",
    "亭勻": "ST014",
    "學姐": "ZO006",
    "zoe": "ZO006",
    "Zoe": "ZO006",
    "妹妹": "ST007",
    "阿尼": "ST007"
  },
  hiddenStudentNames: [
    "A學生",
    "B學生",
    "C學生",
    "張喆",
    "學姐",
    "妹妹",
    "建銘",
    "宇涵",
    "亭昀",
    "大叔",
    "宇含"
  ],
  publicBaseUrl: "https://ntpc-uas.github.io/coachflow/",
  // Google Apps Script 偶爾需要 20~30 秒才能讀完整課表；12 秒會把仍在
  // 正常處理的請求誤判為失敗，導致教練端顯示「雲端更新失敗」。
  requestTimeoutMs: 60000,
  leaveSandbox: {
    enabled: true,
    coachPage: "leave-coach-sandbox.html",
    studentPage: "leave-student-sandbox.html"
  }
};
