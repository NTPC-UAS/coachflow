(function () {
  "use strict";

  const STORAGE_KEY = "coachflow-leave-sandbox-v1";
  const SCHEMA_VERSION = 4;
  const TZ = "Asia/Taipei";
  const PENDING_EXPIRE_HOURS = 24;
  const DEFAULT_REJECT_REASON = "時段已被其他學生選走。";
  const CHARGE_REMINDER_STEP = 4;
  const MAX_CHARGE_REMINDER_LOGS = 24;
  const CALENDAR_REMOVED_PREVIEW_COUNT = 5;
  const LEAVE_PREFILL_STORAGE_KEY = "coachflow-leave-prefill";
  const LEAVE_PREFILL_MAX_AGE_MS = 10 * 60 * 1000;
  const READ_ONLY_ACCOUNT_CODE = "READONLY";
  const READ_ONLY_DEFAULT_COACH_CODE = "MO001";
  const READ_ONLY_DEFAULT_COACH_NAME = "Monster Chang";
  const READ_ONLY_MONTH_LOOKAHEAD = 1;
  const PAYMENT_STATUS_LABELS = {
    unknown: "未註記",
    paid: "已繳費",
    unpaid: "未繳費"
  };
  const el = {
    uiStatusText: document.getElementById("ui-status-text"),
    studentCode: document.getElementById("student-code"),
    studentCoachCode: document.getElementById("student-coach-code"),
    studentLoginBtn: document.getElementById("student-login-btn"),
    studentResetBtn: document.getElementById("student-reset-btn"),
    coachCode: document.getElementById("coach-code"),
    coachLoginBtn: document.getElementById("coach-login-btn"),
    studentSessionText: document.getElementById("student-session-text"),
    studentCalendarMonthLabel: document.getElementById("student-calendar-month-label"),
    studentCalendarPrevBtn: document.getElementById("student-calendar-prev-btn"),
    studentCalendarNextBtn: document.getElementById("student-calendar-next-btn"),
    studentCalendarGrid: document.getElementById("student-calendar-grid"),
    studentDayDetail: document.getElementById("student-day-detail"),
    studentDayModal: document.getElementById("student-day-modal"),
    studentDayModalCloseBtn: document.getElementById("student-day-modal-close"),
    coachSessionText: document.getElementById("coach-session-text"),
    coachLeaveDate: document.getElementById("coach-leave-date"),
    coachLeaveTime: document.getElementById("coach-leave-time"),
    coachLeaveEndTime: document.getElementById("coach-leave-end-time"),
    coachLeaveReason: document.getElementById("coach-leave-reason"),
    coachLeaveAddBtn: document.getElementById("coach-leave-add-btn"),
    coachLeaveTable: document.getElementById("coach-leave-table"),
    coachCalendarMonthLabel: document.getElementById("coach-calendar-month-label"),
    coachCalendarPrevBtn: document.getElementById("coach-calendar-prev-btn"),
    coachCalendarNextBtn: document.getElementById("coach-calendar-next-btn"),
    coachCalendarSyncBtn: document.getElementById("coach-calendar-sync-btn"),
    coachCalendarSyncText: document.getElementById("coach-calendar-sync-text"),
    coachCalendarGrid: document.getElementById("coach-calendar-grid"),
    calendarRemovedSummary: document.getElementById("calendar-removed-summary"),
    calendarRemovedRefreshBtn: document.getElementById("calendar-removed-refresh-btn"),
    calendarRemovedRestoreAllBtn: document.getElementById("calendar-removed-restore-all-btn"),
    calendarRemovedToggleBtn: document.getElementById("calendar-removed-toggle-btn"),
    calendarRemovedTable: document.getElementById("calendar-removed-table"),
    coachDayDetail: document.getElementById("coach-day-detail"),
    coachDayModal: document.getElementById("coach-day-modal"),
    coachDayModalCloseBtn: document.getElementById("coach-day-modal-close"),
    studentLessonsTable: document.getElementById("student-lessons-table"),
    makeupLeaveSelect: document.getElementById("makeup-leave-select"),
    makeupSlotSelect: document.getElementById("makeup-slot-select"),
    loadSlotsBtn: document.getElementById("load-slots-btn"),
    submitMakeupBtn: document.getElementById("submit-makeup-btn"),
    slotWindowHint: document.getElementById("slot-window-hint"),
    studentMakeupTable: document.getElementById("student-makeup-table"),
    coachPendingTable: document.getElementById("coach-pending-table"),
    coachReviewSummary: document.getElementById("coach-review-summary"),
    coachReviewUrgent: document.getElementById("coach-review-urgent"),
    coachReviewStudentFilter: document.getElementById("coach-review-student-filter"),
    coachReviewRefreshBtn: document.getElementById("coach-review-refresh-btn"),
    coachReviewModal: document.getElementById("coach-review-modal"),
    coachReviewModalTitle: document.getElementById("coach-review-modal-title"),
    coachReviewModalMessage: document.getElementById("coach-review-modal-message"),
    coachReviewReasonWrap: document.getElementById("coach-review-reason-wrap"),
    coachReviewReasonInput: document.getElementById("coach-review-reason-input"),
    coachReviewModalConfirmBtn: document.getElementById("coach-review-modal-confirm"),
    coachReviewModalCancelBtn: document.getElementById("coach-review-modal-cancel"),
    coachReviewModalCloseBtn: document.getElementById("coach-review-modal-close"),
    compensationSummary: document.getElementById("compensation-summary"),
    compensationRefreshBtn: document.getElementById("compensation-refresh-btn"),
    compensationRetryAllBtn: document.getElementById("compensation-retry-all-btn"),
    compensationTable: document.getElementById("compensation-table"),
    eventLog: document.getElementById("event-log"),
    chargeStudentSelect: document.getElementById("charge-student-select"),
    chargeStudentEmailInput: document.getElementById("charge-student-email-input"),
    chargeStudentEmailSaveBtn: document.getElementById("charge-student-email-save-btn"),
    chargeEmailSendBtn: document.getElementById("charge-email-send-btn"),
    chargeStudentEmailMeta: document.getElementById("charge-student-email-meta"),
    chargeBaseCountInput: document.getElementById("charge-base-count-input"),
    chargeBaseCountSaveBtn: document.getElementById("charge-base-count-save-btn"),
    chargePaymentStatusSelect: document.getElementById("charge-payment-status-select"),
    chargePaymentNoteInput: document.getElementById("charge-payment-note-input"),
    chargePaymentSaveBtn: document.getElementById("charge-payment-save-btn"),
    chargePaymentMeta: document.getElementById("charge-payment-meta"),
    chargeReminderSummary: document.getElementById("charge-reminder-summary"),
    chargeReminderTable: document.getElementById("charge-reminder-table"),
    chargeMetricsBox: document.getElementById("charge-metrics-box"),
    chargeLedgerTable: document.getElementById("charge-ledger-table"),
    studentOverviewTable: document.getElementById("student-overview-table")
  };

  let state = loadState();
  let activeStudentCode = "";
  let activeCoachCode = "";
  let selectedChargeStudentCode = "";
  let studentCalendarMonthStart = getMonthStart(makeTaipeiDateTime(getDateKeyInTaipei(new Date()), "12:00"));
  let selectedStudentDateKey = getDateKeyInTaipei(new Date());
  let selectedStudentLessonId = "";
  let coachCalendarMonthStart = getMonthStart(makeTaipeiDateTime(getDateKeyInTaipei(new Date()), "12:00"));
  let selectedCoachDateKey = getDateKeyInTaipei(new Date());
  let selectedCoachLessonId = "";
  let coachReviewFilterStudent = "ALL";
  let coachReviewPendingRequestId = "";
  let coachReviewPendingAction = "";
  let coachRejectReasonPreset = DEFAULT_REJECT_REASON;
  let activeCoachReadOnly = false;
  const sendingChargeReminderKeys = new Set();
  let isCalendarRemovedExpanded = false;
  let studentLoginPromptHidden = false;
  let coachLoginPromptHidden = false;

  function loadState() {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      return seedState();
    }
    try {
      const parsed = JSON.parse(raw);
      if (!parsed || typeof parsed !== "object") {
        return seedState();
      }
      if (Number(parsed.version) !== SCHEMA_VERSION) {
        return seedState();
      }
      return parsed;
    } catch (error) {
      console.warn("Failed to parse sandbox state:", error);
      return seedState();
    }
  }

  function seedState() {
    const today = getDateKeyInTaipei(new Date());
    const cycleSat = getNextDateKeyByWeekday(today, "sat", true);
    const cycleSun = addDays(cycleSat, 1);
    const cycleMon = addDays(cycleSat, 2);
    const prevCycleSun = addDays(cycleSun, -7);
    const week1Sun = addDays(cycleSun, 7);
    const week1Sat = addDays(cycleSat, 7);
    const week1Mon = addDays(cycleMon, 7);
    const week2Sun = addDays(cycleSun, 14);
    const week2Sat = addDays(cycleSat, 14);
    const week2Mon = addDays(cycleMon, 14);
    const week3Sun = addDays(cycleSun, 21);
    const week3Sat = addDays(cycleSat, 21);
    const week3Mon = addDays(cycleMon, 21);
    const week4Sun = addDays(cycleSun, 28);
    const week5Sun = addDays(cycleSun, 35);
    const week6Sun = addDays(cycleSun, 42);
    const now = new Date();

    const defaultNotifyEmail = String(window.APP_CONFIG?.defaultNotifyEmail || "").trim();
    const students = [
      {
        code: "STU001",
        name: "A學生",
        coachCode: "CH001",
        email: defaultNotifyEmail,
        chargeStartCount: 0,
        paidThroughCount: 0,
        paymentStatus: "unknown",
        paymentNote: "",
        paymentConfirmedAt: "",
        paymentConfirmedBy: "",
        chargeReminderLogs: []
      },
      {
        code: "STU002",
        name: "B學生",
        coachCode: "CH001",
        email: defaultNotifyEmail,
        chargeStartCount: 2,
        paidThroughCount: 0,
        paymentStatus: "unknown",
        paymentNote: "",
        paymentConfirmedAt: "",
        paymentConfirmedBy: "",
        chargeReminderLogs: []
      },
      {
        code: "STU003",
        name: "C學生",
        coachCode: "CH001",
        email: defaultNotifyEmail,
        chargeStartCount: 1,
        paidThroughCount: 0,
        paymentStatus: "unknown",
        paymentNote: "",
        paymentConfirmedAt: "",
        paymentConfirmedBy: "",
        chargeReminderLogs: []
      }
    ];
    const coaches = [{ code: "CH001", name: "王教練", email: defaultNotifyEmail }];

    const lessons = [
      makeLesson("L000", "STU001", "CH001", prevCycleSun, "10:00"),
      makeLesson("L001", "STU001", "CH001", cycleSun, "10:00"),
      makeLesson("L002", "STU001", "CH001", week1Sun, "10:00"),
      makeLesson("L003", "STU001", "CH001", week2Sun, "10:00"),
      makeLesson("L004", "STU001", "CH001", week3Sun, "10:00"),
      makeLesson("L005", "STU001", "CH001", week4Sun, "10:00"),
      makeLesson("L006", "STU001", "CH001", week5Sun, "10:00"),
      makeLesson("L007", "STU001", "CH001", week6Sun, "10:00"),

      makeLesson("L100", "STU002", "CH001", prevCycleSun, "11:00"),
      makeLesson("L101", "STU002", "CH001", cycleSun, "11:00"),
      makeLesson("L102", "STU002", "CH001", week1Sun, "11:00"),
      makeLesson("L103", "STU002", "CH001", week2Sun, "11:00"),
      makeLesson("L104", "STU002", "CH001", week3Sun, "11:00"),
      makeLesson("L105", "STU002", "CH001", week4Sun, "11:00"),
      makeLesson("L106", "STU002", "CH001", week5Sun, "11:00"),
      makeLesson("L107", "STU002", "CH001", week6Sun, "11:00"),

      makeLesson("L200", "STU003", "CH001", prevCycleSun, "12:00"),
      makeLesson("L201", "STU003", "CH001", cycleSun, "12:00"),
      makeLesson("L202", "STU003", "CH001", week1Sun, "12:00"),
      makeLesson("L203", "STU003", "CH001", week2Sun, "12:00"),
      makeLesson("L204", "STU003", "CH001", week3Sun, "12:00"),
      makeLesson("L205", "STU003", "CH001", week4Sun, "12:00"),
      makeLesson("L206", "STU003", "CH001", week5Sun, "12:00"),
      makeLesson("L207", "STU003", "CH001", week6Sun, "12:00")
    ];

    const leaveRequests = [
      makeLeave("LEAVE_A_OPEN_SLOT", "L002", "STU001", "CH001", hoursAgoIso(now, 6)),
      makeLeave("LEAVE_B_VERIFY_SLOT", "L102", "STU002", "CH001", hoursAgoIso(now, 4)),
      makeLeave("LEAVE_A_APPROVED", "L003", "STU001", "CH001", hoursAgoIso(now, 40)),
      makeLeave("LEAVE_A_REJECTED", "L004", "STU001", "CH001", hoursAgoIso(now, 10)),
      makeLeave("LEAVE_B_EXPIRED", "L103", "STU002", "CH001", hoursAgoIso(now, 60)),
      makeLeave("LEAVE_C_CANCELLED", "L203", "STU003", "CH001", hoursAgoIso(now, 30)),
      makeLeave("LEAVE_C_PENDING", "L204", "STU003", "CH001", hoursAgoIso(now, 3))
    ];

    lessons.forEach((lesson) => {
      if (leaveRequests.some((leave) => leave.lessonId === lesson.id)) {
        lesson.calendarOccupied = false;
        lesson.attendanceStatus = "leave-normal";
      }
    });

    const makeupRequests = [
      {
        id: "MAKEUP_P_A",
        leaveId: "LEAVE_A_OPEN_SLOT",
        lessonId: "L002",
        studentCode: "STU001",
        coachCode: "CH001",
        status: "pending",
        pendingAt: hoursAgoIso(now, 1),
        resolvedAt: "",
        startAt: makeTaipeiDateTime(week1Mon, "18:30").toISOString(),
        rejectReason: ""
      },
      {
        id: "MAKEUP_A_A",
        leaveId: "LEAVE_A_APPROVED",
        lessonId: "L003",
        studentCode: "STU001",
        coachCode: "CH001",
        status: "approved",
        pendingAt: hoursAgoIso(now, 30),
        resolvedAt: hoursAgoIso(now, 26),
        startAt: makeTaipeiDateTime(week2Mon, "18:30").toISOString(),
        rejectReason: ""
      },
      {
        id: "MAKEUP_R_A",
        leaveId: "LEAVE_A_REJECTED",
        lessonId: "L004",
        studentCode: "STU001",
        coachCode: "CH001",
        status: "rejected",
        pendingAt: hoursAgoIso(now, 10),
        resolvedAt: hoursAgoIso(now, 9),
        startAt: makeTaipeiDateTime(week3Sat, "09:00").toISOString(),
        rejectReason: "時段已被其他學生選走。"
      },
      {
        id: "MAKEUP_E_B",
        leaveId: "LEAVE_B_EXPIRED",
        lessonId: "L103",
        studentCode: "STU002",
        coachCode: "CH001",
        status: "expired",
        pendingAt: hoursAgoIso(now, 40),
        resolvedAt: hoursAgoIso(now, 15),
        startAt: makeTaipeiDateTime(week1Mon, "19:30").toISOString(),
        rejectReason: ""
      },
      {
        id: "MAKEUP_C_C",
        leaveId: "LEAVE_C_CANCELLED",
        lessonId: "L203",
        studentCode: "STU003",
        coachCode: "CH001",
        status: "cancelled",
        pendingAt: hoursAgoIso(now, 8),
        resolvedAt: hoursAgoIso(now, 7),
        startAt: makeTaipeiDateTime(week2Sat, "10:00").toISOString(),
        rejectReason: ""
      },
      {
        id: "MAKEUP_P_C",
        leaveId: "LEAVE_C_PENDING",
        lessonId: "L204",
        studentCode: "STU003",
        coachCode: "CH001",
        status: "pending",
        pendingAt: hoursAgoIso(now, 2),
        resolvedAt: "",
        startAt: makeTaipeiDateTime(week3Mon, "19:30").toISOString(),
        rejectReason: ""
      }
    ];

    lessons.push(
      {
        id: "MAKEUP_LESSON_1",
        studentCode: "STU001",
        coachCode: "CH001",
        startAt: makeTaipeiDateTime(week2Mon, "18:30").toISOString(),
        sourceType: "MAKEUP",
        calendarOccupied: true,
        attendanceStatus: "scheduled",
        charged: false
      }
    );

    markLesson(lessons, "L000", "no-show", true);
    markLesson(lessons, "L100", "temporary-leave", true);
    markLesson(lessons, "L200", "major-case", false);

    return {
      version: SCHEMA_VERSION,
      students,
      coaches,
      lessons,
      leaveRequests,
      makeupRequests,
      coachBlocks: [
        {
          id: "BLOCK_DEMO_1",
          coachCode: "CH001",
          startAt: makeTaipeiDateTime(week2Mon, "19:30").toISOString(),
          endAt: makeTaipeiDateTime(week2Mon, "20:30").toISOString(),
          reason: "教練外出"
        }
      ],
      eventLog: [
        {
          id: "LOG_BOOT_1",
          at: now.toISOString(),
          message: "已載入模擬資料：待審／核准／退回／逾時／取消。"
        },
        {
          id: "LOG_BOOT_2",
          at: now.toISOString(),
          message: "驗證情境：STU001 請假後，STU002 可看到釋出的空堂。"
        }
      ]
    };
  }

  function makeLeave(id, lessonId, studentCode, coachCode, submittedAt) {
    return {
      id,
      lessonId,
      studentCode,
      coachCode,
      type: "normal",
      submittedAt,
      makeupEligible: true
    };
  }

  function markLesson(lessons, lessonId, attendanceStatus, charged) {
    const lesson = lessons.find((item) => item.id === lessonId);
    if (!lesson) {
      return;
    }
    lesson.attendanceStatus = attendanceStatus;
    lesson.charged = charged;
  }

  function hoursAgoIso(baseDate, hours) {
    return new Date(baseDate.getTime() - hours * 3600 * 1000).toISOString();
  }

  function makeLesson(id, studentCode, coachCode, dateKey, timeText) {
    const startAt = makeTaipeiDateTime(dateKey, timeText).toISOString();
    return {
      id,
      studentCode,
      coachCode,
      startAt,
      sourceType: "REGULAR",
      calendarEventId: `GCAL_${id}`,
      calendarOccupied: true,
      attendanceStatus: "scheduled",
      charged: false
    };
  }

  function saveState() {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  }

  function addLog(message) {
    state.eventLog.unshift({
      id: newId("LOG"),
      at: new Date().toISOString(),
      message
    });
    state.eventLog = state.eventLog.slice(0, 120);
  }

  function newId(prefix) {
    return `${prefix}_${Math.random().toString(36).slice(2, 8)}_${Date.now().toString(36)}`;
  }

  function getDateTimePartsInTaipei(dateValue) {
    const parts = new Intl.DateTimeFormat("en-CA", {
      timeZone: TZ,
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
      hour12: false
    }).formatToParts(dateValue instanceof Date ? dateValue : new Date(dateValue));
    return {
      year: parts.find((p) => p.type === "year")?.value || "0000",
      month: parts.find((p) => p.type === "month")?.value || "00",
      day: parts.find((p) => p.type === "day")?.value || "00",
      hour: parts.find((p) => p.type === "hour")?.value || "00",
      minute: parts.find((p) => p.type === "minute")?.value || "00"
    };
  }

  function getDateKeyFromIso(isoText) {
    return getDateKeyInTaipei(new Date(isoText));
  }

  function isTimeWithinRange(targetIso, startIso, endIso) {
    const target = new Date(targetIso).getTime();
    const start = new Date(startIso).getTime();
    const end = new Date(endIso).getTime();
    return target >= start && target < end;
  }

  function isRangeOverlap(startA, endA, startB, endB) {
    const a1 = new Date(startA).getTime();
    const a2 = new Date(endA).getTime();
    const b1 = new Date(startB).getTime();
    const b2 = new Date(endB).getTime();
    return a1 < b2 && b1 < a2;
  }

  function buildMakeupCode(fromDate) {
    const p = getDateTimePartsInTaipei(fromDate);
    return `MP${p.year}${p.month}${p.day}${p.hour}${p.minute}`;
  }

  function ensureMakeupCodes() {
    const used = new Set();
    state.makeupRequests.forEach((request) => {
      const base = request.code || buildMakeupCode(request.pendingAt || request.startAt || new Date());
      let code = base;
      let seq = 1;
      while (used.has(code)) {
        seq += 1;
        code = `${base}-${String(seq).padStart(2, "0")}`;
      }
      request.code = code;
      used.add(code);
    });
  }

  function getDateKeyInTaipei(date) {
    const parts = new Intl.DateTimeFormat("en-CA", {
      timeZone: TZ,
      year: "numeric",
      month: "2-digit",
      day: "2-digit"
    }).formatToParts(date);
    return `${parts.find((item) => item.type === "year").value}-${parts.find((item) => item.type === "month").value}-${parts.find((item) => item.type === "day").value}`;
  }

  function makeTaipeiDateTime(dateKey, timeText) {
    return new Date(`${dateKey}T${timeText}:00+08:00`);
  }

  function normalizeDateInput(inputText) {
    const raw = String(inputText || "").trim();
    const matched = raw.match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/);
    if (!matched) {
      return "";
    }
    const year = matched[1];
    const month = String(Number(matched[2])).padStart(2, "0");
    const day = String(Number(matched[3])).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }

  function normalizeTimeInput(inputText) {
    const raw = String(inputText || "").trim();
    const matched = raw.match(/^(\d{1,2}):(\d{1,2})$/);
    if (!matched) {
      return "";
    }
    const hour = Number(matched[1]);
    const minute = Number(matched[2]);
    if (hour < 0 || hour > 23 || minute < 0 || minute > 59) {
      return "";
    }
    return `${String(hour).padStart(2, "0")}:${String(minute).padStart(2, "0")}`;
  }

  function formatDateTime(isoText) {
    return new Intl.DateTimeFormat("zh-TW", {
      timeZone: TZ,
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
      hour12: false
    }).format(new Date(isoText));
  }

  function addDays(dateKey, days) {
    const base = makeTaipeiDateTime(dateKey, "12:00");
    base.setUTCDate(base.getUTCDate() + days);
    return getDateKeyInTaipei(base);
  }

  function getMonthStart(date) {
    const key = getDateKeyInTaipei(date);
    const [yearText, monthText] = key.split("-");
    return makeTaipeiDateTime(`${yearText}-${monthText}-01`, "00:00");
  }

  function getMonthMeta(date) {
    const key = getDateKeyInTaipei(date);
    const [yearText, monthText] = key.split("-");
    const year = Number(yearText);
    const month = Number(monthText);
    const firstDayKey = `${yearText}-${monthText}-01`;
    const firstWeekday = getWeekdayIndex(firstDayKey);
    const daysInMonth = new Date(year, month, 0).getDate();
    return { year, month, firstWeekday, daysInMonth, firstDayKey };
  }

  function getWeekdayIndex(dateKey) {
    return makeTaipeiDateTime(dateKey, "12:00").getUTCDay();
  }

  function shiftMonth(date, diff) {
    const key = getDateKeyInTaipei(date);
    const [yearText, monthText] = key.split("-");
    let year = Number(yearText);
    let month = Number(monthText) + diff;
    while (month > 12) {
      month -= 12;
      year += 1;
    }
    while (month < 1) {
      month += 12;
      year -= 1;
    }
    return makeTaipeiDateTime(`${String(year).padStart(4, "0")}-${String(month).padStart(2, "0")}-01`, "00:00");
  }

  function getMonthStartTime(date) {
    return getMonthStart(date).getTime();
  }

  function getReadOnlyMonthBounds() {
    const minMonth = getMonthStart(new Date());
    return {
      minMonth,
      maxMonth: shiftMonth(minMonth, READ_ONLY_MONTH_LOOKAHEAD)
    };
  }

  function clampReadOnlyCoachMonth() {
    if (!isCoachReadOnlyMode()) {
      return false;
    }
    const { minMonth, maxMonth } = getReadOnlyMonthBounds();
    const currentTime = getMonthStartTime(coachCalendarMonthStart);
    if (currentTime < minMonth.getTime()) {
      coachCalendarMonthStart = minMonth;
      return true;
    }
    if (currentTime > maxMonth.getTime()) {
      coachCalendarMonthStart = maxMonth;
      return true;
    }
    return false;
  }

  function getStudentLessonsForDate(studentCode, dateKey) {
    return state.lessons
      .filter(
        (lesson) =>
          lesson.studentCode === studentCode &&
          lesson.attendanceStatus !== "coach-leave" &&
          lesson.attendanceStatus !== "calendar-removed" &&
          getDateKeyInTaipei(new Date(lesson.startAt)) === dateKey
      )
      .sort((a, b) => new Date(a.startAt) - new Date(b.startAt));
  }

  function getCoachLessonsForDate(coachCode, dateKey) {
    return state.lessons
      .filter(
        (lesson) =>
          lesson.coachCode === coachCode &&
          lesson.calendarOccupied &&
          getDateKeyInTaipei(new Date(lesson.startAt)) === dateKey
      )
      .sort((a, b) => new Date(a.startAt) - new Date(b.startAt));
  }

  function getCoachPendingForDate(coachCode, dateKey) {
    return state.makeupRequests
      .filter(
        (request) =>
          request.coachCode === coachCode &&
          request.status === "pending" &&
          getDateKeyInTaipei(new Date(request.startAt)) === dateKey
      )
      .sort((a, b) => new Date(a.pendingAt) - new Date(b.pendingAt));
  }

  function getCoachBlocksForDate(coachCode, dateKey) {
    return (state.coachBlocks || [])
      .filter((block) => block.coachCode === coachCode && getDateKeyInTaipei(new Date(block.startAt)) === dateKey)
      .sort((a, b) => new Date(a.startAt) - new Date(b.startAt));
  }

  function getStudentCoachBlocksForDate(studentCode, dateKey) {
    const student = getStudentByCode(studentCode);
    if (!student) {
      return [];
    }
    return getCoachBlocksForDate(student.coachCode, dateKey);
  }

  function getPendingHoursLeft(request) {
    const passedHours = (new Date() - new Date(request.pendingAt)) / 36e5;
    return Math.max(0, PENDING_EXPIRE_HOURS - passedHours);
  }

  function getTimeText(isoText) {
    const formatted = formatDateTime(isoText);
    const parts = formatted.split(" ");
    return parts[1] || "";
  }

  function setUiStatus(message) {
    if (el.uiStatusText) {
      el.uiStatusText.textContent = String(message || "");
    }
  }

  function getEventElement(event) {
    const target = event?.target;
    if (!target) {
      return null;
    }
    if (target.nodeType === 1 && typeof target.closest === "function") {
      return target;
    }
    if (target.parentElement && typeof target.parentElement.closest === "function") {
      return target.parentElement;
    }
    return null;
  }

  function closestFromEvent(event, selector) {
    const targetEl = getEventElement(event);
    return targetEl ? targetEl.closest(selector) : null;
  }

  function isReadonlyParamEnabled() {
    const params = new URLSearchParams(window.location.search);
    const values = [
      params.get("readonly"),
      params.get("readOnly"),
      params.get("mode"),
      params.get("role"),
      params.get("permission")
    ].map((value) => String(value || "").trim().toLowerCase());
    return values.some((value) => ["1", "true", "yes", "readonly", "read-only", "viewer", "view"].includes(value));
  }

  function isReadonlyCoachRole(coach) {
    const role = String(coach?.role || "").trim().toLowerCase();
    return ["readonly", "read-only", "viewer", "view"].includes(role);
  }

  function isCoachReadOnlyMode() {
    return Boolean(activeCoachReadOnly);
  }

  function isReadOnlyVisibleCoachLesson(lesson) {
    if (!isCoachReadOnlyMode()) {
      return true;
    }
    return !isGeneratedLocalCalendarEventId(lesson?.calendarEventId);
  }

  function isReadOnlyLoginCode(code) {
    return normalizeParticipantCode(code).replace(/[-_\s]/g, "") === READ_ONLY_ACCOUNT_CODE;
  }

  function resolveReadOnlyCoachCode() {
    const configuredCoach = getCoachByCode(READ_ONLY_DEFAULT_COACH_CODE);
    if (configuredCoach) {
      return READ_ONLY_DEFAULT_COACH_CODE;
    }
    const monsterCoach = state.coaches.find((coach) => {
      const code = normalizeParticipantCode(coach?.code);
      const name = String(coach?.name || "").trim().toLowerCase();
      return code.startsWith("MO") || name.includes("monster");
    });
    if (monsterCoach?.code) {
      return normalizeParticipantCode(monsterCoach.code);
    }
    ensureCoachProfile(READ_ONLY_DEFAULT_COACH_CODE, READ_ONLY_DEFAULT_COACH_NAME);
    return READ_ONLY_DEFAULT_COACH_CODE;
  }

  function requireCoachWriteAccess() {
    if (!isCoachReadOnlyMode()) {
      return true;
    }
    alert("此帳號為唯讀權限，只能查看教練月曆，不能修改資料。");
    return false;
  }

  function setSectionHiddenByElement(target, hidden) {
    const section = target?.closest?.("section, article");
    if (section) {
      section.hidden = Boolean(hidden);
    }
  }

  function updateCoachReadOnlyUi() {
    const hasCoachSession = Boolean(activeCoachCode);
    const readOnly = isCoachReadOnlyMode();
    document.body.classList.toggle("coach-readonly-mode", readOnly);
    document.body.classList.toggle("coach-session-active", hasCoachSession);
    [
      el.coachLeaveAddBtn,
      el.coachLeaveTable,
      el.calendarRemovedTable,
      el.coachPendingTable,
      el.chargeStudentSelect,
      el.studentOverviewTable,
      el.compensationTable,
      el.eventLog
    ].forEach((target) => setSectionHiddenByElement(target, !hasCoachSession || readOnly));
    setSectionHiddenByElement(el.coachCalendarGrid, !hasCoachSession);
    if (el.coachCalendarSyncBtn) {
      const syncRow = el.coachCalendarSyncBtn.closest(".btn-row");
      if (syncRow) {
        syncRow.hidden = !hasCoachSession || readOnly;
      }
      el.coachCalendarSyncBtn.disabled = !hasCoachSession || readOnly;
    }
    if (el.coachCalendarSyncText && readOnly) {
      el.coachCalendarSyncText.textContent = "唯讀模式：只能查看教練月曆，不能同步或修改資料。";
    }
    if (el.coachCalendarGrid && readOnly) {
      closeCoachDayModal();
    }
  }

  function updateModalBodyState() {
    const hasOpenModal = [el.studentDayModal, el.coachDayModal, el.coachReviewModal].some((modal) => modal && !modal.hidden);
    document.body.classList.toggle("modal-open", hasOpenModal);
  }

  function openStudentDayModal() {
    if (!el.studentDayModal) {
      return;
    }
    el.studentDayModal.hidden = false;
    updateModalBodyState();
  }

  function closeStudentDayModal() {
    if (!el.studentDayModal) {
      return;
    }
    el.studentDayModal.hidden = true;
    updateModalBodyState();
  }

  function openCoachDayModal() {
    if (!el.coachDayModal) {
      return;
    }
    el.coachDayModal.hidden = false;
    updateModalBodyState();
  }

  function closeCoachDayModal() {
    if (!el.coachDayModal) {
      return;
    }
    el.coachDayModal.hidden = true;
    updateModalBodyState();
  }

  function clearStudentCalendarSelection() {
    if (!selectedStudentDateKey && !selectedStudentLessonId) {
      return;
    }
    selectedStudentDateKey = "";
    selectedStudentLessonId = "";
    closeStudentDayModal();
    renderStudentCalendar();
    setUiStatus("已清除學生月曆選取。");
  }

  function clearCoachCalendarSelection() {
    if (!selectedCoachDateKey && !selectedCoachLessonId) {
      return;
    }
    selectedCoachDateKey = "";
    selectedCoachLessonId = "";
    closeCoachDayModal();
    renderCoachCalendar();
    setUiStatus("已清除教練月曆選取。");
  }

  function activateStudentSession(studentCodeInput, coachCodeInput, silentMode) {
    const studentCode = normalizeParticipantCode(studentCodeInput);
    const coachCode = normalizeParticipantCode(coachCodeInput);
    const student = getStudentByCode(studentCode);
    if (!student) {
      if (!silentMode) {
        alert("找不到學生代碼。請先回 CoachFlow 載入學生資料，或稍後重試同步。");
      }
      return false;
    }
    const resolvedCoachCode = coachCode || student.coachCode;
    if (student.coachCode !== resolvedCoachCode) {
      if (!silentMode) {
        alert("教練代碼與學生不相符。");
      }
      return false;
    }
    activeStudentCode = studentCode;
    activeCoachCode = resolvedCoachCode;
    activeCoachReadOnly = false;
    selectedChargeStudentCode = studentCode;
    if (el.studentCoachCode) {
      el.studentCoachCode.value = resolvedCoachCode;
    }
    if (el.coachCode) {
      el.coachCode.value = resolvedCoachCode;
    }
    selectedStudentLessonId = "";
    closeStudentDayModal();
    closeCoachDayModal();
    closeCoachReviewModal();
    const firstLesson = state.lessons
      .filter((lesson) => lesson.studentCode === studentCode)
      .sort((a, b) => new Date(a.startAt) - new Date(b.startAt))[0];
    if (firstLesson) {
      const focusDate = new Date(firstLesson.startAt);
      studentCalendarMonthStart = getMonthStart(focusDate);
      selectedStudentDateKey = getDateKeyInTaipei(focusDate);
    }
    renderAll();
    setStudentLoginPromptVisibility(true);
    setUiStatus(`學生 ${student.name || studentCode} 已載入，可直接點月曆。`);
    return true;
  }

  function activateCoachSession(coachCodeInput, silentMode) {
    const isReadOnlyLogin = isReadOnlyLoginCode(coachCodeInput);
    const coachCode = isReadOnlyLogin ? resolveReadOnlyCoachCode() : normalizeParticipantCode(coachCodeInput);
    const coach = getCoachByCode(coachCode);
    if (!coach) {
      if (!silentMode) {
        alert("找不到教練代碼。請先回 CoachFlow 載入教練資料，或稍後重試同步。");
      }
      return false;
    }
    activeCoachCode = coachCode;
    activeCoachReadOnly = isReadOnlyLogin || isReadonlyParamEnabled() || isReadonlyCoachRole(coach);
    const coachPending = state.makeupRequests
      .filter((request) => request.coachCode === coachCode && request.status === "pending")
      .sort((a, b) => new Date(a.pendingAt) - new Date(b.pendingAt))[0];
    const coachLesson = state.lessons
      .filter((lesson) => lesson.coachCode === coachCode)
      .sort((a, b) => new Date(a.startAt) - new Date(b.startAt))[0];
    const focusDate = activeCoachReadOnly
      ? new Date()
      : coachPending
        ? new Date(coachPending.startAt)
        : coachLesson
          ? new Date(coachLesson.startAt)
          : new Date();
    coachCalendarMonthStart = getMonthStart(focusDate);
    clampReadOnlyCoachMonth();
    selectedCoachDateKey = getDateKeyInTaipei(focusDate);
    selectedCoachLessonId = "";
    closeStudentDayModal();
    closeCoachDayModal();
    closeCoachReviewModal();
    renderAll();
    updateCoachReadOnlyUi();
    setCoachLoginPromptVisibility(true);
    setUiStatus(`教練 ${coach.name || coachCode} 已載入${activeCoachReadOnly ? "（唯讀）" : ""}。`);
    return true;
  }

  function openCoachReviewModal(requestId, action) {
    if (!el.coachReviewModal || !el.coachReviewModalTitle || !el.coachReviewModalMessage || !el.coachReviewModalConfirmBtn) {
      return;
    }
    const request = state.makeupRequests.find((item) => item.id === requestId);
    if (!request || request.coachCode !== activeCoachCode || request.status !== "pending") {
      alert("找不到待審申請，或你沒有權限。");
      return;
    }
    coachReviewPendingRequestId = requestId;
    coachReviewPendingAction = action;
    const actionText = action === "approve" ? "核准" : "退回";
    const message = [
      `申請代碼：${request.code || request.id}`,
      `學生：${getStudentDisplayName(request.studentCode)}`,
      `補課時段：${formatDateTime(request.startAt)}`,
      "",
      `是否確認要「${actionText}」這筆申請？`
    ].join("\n");
    el.coachReviewModalTitle.textContent = `${actionText}補課申請`;
    el.coachReviewModalMessage.textContent = message;
    el.coachReviewModalConfirmBtn.textContent = `確認${actionText}`;
    el.coachReviewModalConfirmBtn.classList.toggle("danger", action !== "approve");
    if (el.coachReviewReasonWrap && el.coachReviewReasonInput) {
      const isReject = action === "reject";
      el.coachReviewReasonWrap.hidden = !isReject;
      if (isReject) {
        el.coachReviewReasonInput.value = coachRejectReasonPreset;
      }
    }
    el.coachReviewModal.hidden = false;
    updateModalBodyState();
    if (action === "reject" && el.coachReviewReasonInput) {
      el.coachReviewReasonInput.focus();
      el.coachReviewReasonInput.select();
    }
  }

  function closeCoachReviewModal() {
    if (!el.coachReviewModal) {
      return;
    }
    el.coachReviewModal.hidden = true;
    coachReviewPendingRequestId = "";
    coachReviewPendingAction = "";
    if (el.coachReviewReasonWrap) {
      el.coachReviewReasonWrap.hidden = true;
    }
    updateModalBodyState();
  }

  function getAppsScriptUrl() {
    return String(window.APP_CONFIG?.leaveAppsScriptUrl || window.APP_CONFIG?.appsScriptUrl || "").trim();
  }

  function getCoachflowAppsScriptUrl() {
    return String(window.APP_CONFIG?.coachflowAppsScriptUrl || "").trim();
  }

  async function callAppsScriptApi(action, payload = {}, method = "POST") {
    const url = getAppsScriptUrl();
    if (!url) {
      return {
        ok: true,
        mock: true,
        action,
        eventId: payload.eventId || payload.calendarEventId || newId("GCAL")
      };
    }

    const controller = new AbortController();
    const timeoutMs = Number(window.APP_CONFIG?.requestTimeoutMs || 12000);
    const timeout = window.setTimeout(() => controller.abort(), timeoutMs);
    try {
      const endpoint = new URL(url);
      let response;
      if (String(method).toUpperCase() === "GET") {
        endpoint.searchParams.set("action", action);
        endpoint.searchParams.set("_ts", String(Date.now()));
        Object.entries(payload || {}).forEach(([key, value]) => {
          if (value !== undefined && value !== null) {
            endpoint.searchParams.set(key, String(value));
          }
        });
        response = await fetch(endpoint.toString(), {
          method: "GET",
          cache: "no-store",
          signal: controller.signal
        });
      } else {
        response = await fetch(endpoint.toString(), {
          method: "POST",
          cache: "no-store",
          headers: {
            "Content-Type": "text/plain;charset=UTF-8"
          },
          body: JSON.stringify({ action, ...payload, _ts: Date.now() }),
          signal: controller.signal
        });
      }
      const json = await response.json();
      if (!response.ok || json?.ok === false) {
        throw new Error(json?.message || `${action} 執行失敗`);
      }
      return json;
    } catch (error) {
      if (error?.name === "AbortError") {
        throw new Error(`${action} 逾時`);
      }
      throw error;
    } finally {
      window.clearTimeout(timeout);
    }
  }

  async function callCoachflowApi(action, payload = {}, method = "POST") {
    const url = getCoachflowAppsScriptUrl();
    if (!url) {
      throw new Error("未設定 coachflowAppsScriptUrl。");
    }

    const controller = new AbortController();
    const timeoutMs = Number(window.APP_CONFIG?.requestTimeoutMs || 12000);
    const timeout = window.setTimeout(() => controller.abort(), timeoutMs);
    try {
      const endpoint = new URL(url);
      let response;
      if (String(method).toUpperCase() === "GET") {
        endpoint.searchParams.set("action", action);
        endpoint.searchParams.set("_ts", String(Date.now()));
        Object.entries(payload || {}).forEach(([key, value]) => {
          if (value !== undefined && value !== null) {
            endpoint.searchParams.set(key, String(value));
          }
        });
        response = await fetch(endpoint.toString(), {
          method: "GET",
          cache: "no-store",
          signal: controller.signal
        });
      } else {
        response = await fetch(endpoint.toString(), {
          method: "POST",
          cache: "no-store",
          headers: {
            "Content-Type": "text/plain;charset=UTF-8"
          },
          body: JSON.stringify({ action, ...payload, _ts: Date.now() }),
          signal: controller.signal
        });
      }
      const json = await response.json();
      if (!response.ok || json?.ok === false) {
        throw new Error(json?.message || `${action} 執行失敗`);
      }
      return json;
    } catch (error) {
      if (error?.name === "AbortError") {
        throw new Error(`${action} 逾時`);
      }
      throw error;
    } finally {
      window.clearTimeout(timeout);
    }
  }

  function enqueueCompensationTask(type, payload, reason) {
    state.compensationTasks = Array.isArray(state.compensationTasks) ? state.compensationTasks : [];
    state.compensationTasks.unshift({
      id: newId("COMP"),
      type,
      payload,
      reason,
      status: "pending",
      createdAt: new Date().toISOString()
    });
    state.compensationTasks = state.compensationTasks.slice(0, 80);
    addLog(`[補償任務] ${type} 已建立：${reason}`);
    saveState();
  }

  function getLessonCloudMatchKey(lesson) {
    if (!lesson) {
      return "";
    }
    return [
      normalizeParticipantCode(lesson.studentCode),
      normalizeParticipantCode(lesson.coachCode),
      new Date(lesson.startAt || "").toISOString()
    ].join("|");
  }

  function buildCloudLeaveRecord(leave, lesson) {
    const targetLesson = lesson || getLessonById(leave?.lessonId);
    if (!leave || !targetLesson) {
      return null;
    }
    return {
      id: leave.id,
      lessonId: leave.lessonId,
      lessonKey: getLessonCloudMatchKey(targetLesson),
      calendarEventId: targetLesson.calendarEventId || "",
      studentCode: normalizeParticipantCode(leave.studentCode || targetLesson.studentCode),
      coachCode: normalizeParticipantCode(leave.coachCode || targetLesson.coachCode),
      lessonStartAt: targetLesson.startAt,
      type: leave.type || "normal",
      submittedAt: leave.submittedAt || new Date().toISOString(),
      makeupEligible: leave.makeupEligible !== false,
      revokedAt: leave.revokedAt || "",
      revokedBy: leave.revokedBy || ""
    };
  }

  function findLessonForCloudLeaveRecord(record) {
    if (!record) {
      return null;
    }
    const lessonId = String(record.lessonId || "").trim();
    if (lessonId) {
      const byId = getLessonById(lessonId);
      if (byId) {
        return byId;
      }
    }
    const recordEventId = normalizeCalendarEventId(record.calendarEventId);
    if (recordEventId) {
      const byEventId = state.lessons.find((lesson) => normalizeCalendarEventId(lesson.calendarEventId) === recordEventId);
      if (byEventId) {
        return byEventId;
      }
    }
    const studentCode = normalizeParticipantCode(record.studentCode);
    const coachCode = normalizeParticipantCode(record.coachCode);
    const startTime = new Date(record.lessonStartAt || "").getTime();
    if (!studentCode || !coachCode || !Number.isFinite(startTime)) {
      return null;
    }
    return state.lessons
      .filter((lesson) => (
        normalizeParticipantCode(lesson.studentCode) === studentCode &&
        normalizeParticipantCode(lesson.coachCode) === coachCode &&
        Math.abs(new Date(lesson.startAt).getTime() - startTime) <= 2 * 60 * 1000
      ))
      .sort((a, b) => Math.abs(new Date(a.startAt).getTime() - startTime) - Math.abs(new Date(b.startAt).getTime() - startTime))[0] || null;
  }

  function applyCloudLeaveRecord(record) {
    if (!record || String(record.type || "normal") !== "normal") {
      return false;
    }
    const lesson = findLessonForCloudLeaveRecord(record);
    if (!lesson) {
      return false;
    }
    const existing = state.leaveRequests.find((leave) => leave.id === record.id) ||
      state.leaveRequests.find((leave) => leave.lessonId === lesson.id && leave.studentCode === record.studentCode && !leave.revokedAt);
    let changed = false;
    const revokedAt = String(record.revokedAt || "").trim();
    if (revokedAt) {
      if (existing && !existing.revokedAt) {
        existing.revokedAt = revokedAt;
        existing.revokedBy = record.revokedBy || "cloud";
        changed = true;
      }
      if (lesson.attendanceStatus === "leave-normal") {
        lesson.attendanceStatus = "scheduled";
        lesson.calendarOccupied = true;
        changed = true;
      }
      return changed;
    }

    if (!existing) {
      state.leaveRequests.push({
        id: record.id || newId("LEAVE"),
        lessonId: lesson.id,
        studentCode: record.studentCode,
        coachCode: record.coachCode,
        type: "normal",
        submittedAt: record.submittedAt || new Date().toISOString(),
        makeupEligible: record.makeupEligible !== false
      });
      changed = true;
    }
    if (lesson.attendanceStatus !== "leave-normal" || lesson.calendarOccupied !== false) {
      lesson.attendanceStatus = "leave-normal";
      lesson.calendarOccupied = false;
      changed = true;
    }
    return changed;
  }

  async function pushCloudLeaveRecord(leave, lesson) {
    if (!getAppsScriptUrl()) {
      return false;
    }
    const record = buildCloudLeaveRecord(leave, lesson);
    if (!record) {
      return false;
    }
    const result = await callAppsScriptApi("saveLeaveRecord", { record });
    return result?.ok !== false;
  }

  async function syncCloudLeaveRecords(scope = {}) {
    if (!getAppsScriptUrl()) {
      return false;
    }
    const payload = {
      coachCode: normalizeParticipantCode(scope.coachCode || activeCoachCode || ""),
      studentCode: normalizeParticipantCode(scope.studentCode || activeStudentCode || "")
    };
    const result = await callAppsScriptApi("listLeaveRecords", payload, "GET");
    const records = Array.isArray(result?.records) ? result.records : [];
    let changed = false;
    records.forEach((record) => {
      if (applyCloudLeaveRecord(record)) {
        changed = true;
      }
    });
    if (changed) {
      addLog(`[雲端請假] 已同步 ${records.length} 筆請假紀錄。`);
      saveState();
    }
    return changed;
  }

  function getCoachByCode(coachCode) {
    return state.coaches.find((coach) => coach.code === coachCode);
  }

  function getStudentDisplayName(studentCode) {
    const student = getStudentByCode(studentCode);
    const name = String(student?.name || "").trim();
    return name || String(studentCode || "-");
  }

  function getCoachDisplayName(coachCode) {
    const coach = getCoachByCode(coachCode);
    const name = String(coach?.name || "").trim();
    return name || String(coachCode || "-");
  }

  function normalizeParticipantCode(code) {
    return String(code || "").trim().toUpperCase();
  }

  function ensureCoachProfile(coachCode, coachName) {
    const normalizedCode = normalizeParticipantCode(coachCode);
    if (!normalizedCode) {
      return null;
    }
    const existed = getCoachByCode(normalizedCode);
    if (existed) {
      const nextName = String(coachName || "").trim();
      if (nextName && existed.name !== nextName) {
        existed.name = nextName;
      }
      return existed;
    }
    const fallbackEmail = String(window.APP_CONFIG?.defaultNotifyEmail || "").trim();
    const created = {
      code: normalizedCode,
      name: String(coachName || "").trim() || `${normalizedCode} 教練`,
      email: fallbackEmail
    };
    state.coaches.push(created);
    addLog(`[同步] 已建立教練代碼 ${normalizedCode}（來源資料）。`);
    return created;
  }

  function pickDefaultLessonTimeForCoach(coachCode) {
    const candidates = ["10:00", "11:00", "12:00", "18:30", "19:30", "20:30"];
    const counts = new Map(candidates.map((time) => [time, 0]));
    state.lessons
      .filter((lesson) => lesson.coachCode === coachCode && lesson.sourceType === "REGULAR")
      .forEach((lesson) => {
        const lessonTime = getTimeText(lesson.startAt);
        if (counts.has(lessonTime)) {
          counts.set(lessonTime, counts.get(lessonTime) + 1);
        }
      });
    return [...candidates].sort((a, b) => {
      const countDiff = (counts.get(a) || 0) - (counts.get(b) || 0);
      if (countDiff !== 0) {
        return countDiff;
      }
      return candidates.indexOf(a) - candidates.indexOf(b);
    })[0];
  }

  function pickDefaultLessonWeekdayForCoach(coachCode) {
    const weekdays = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"];
    const counts = new Map(weekdays.map((day) => [day, 0]));
    state.lessons
      .filter((lesson) => lesson.coachCode === coachCode && lesson.sourceType === "REGULAR")
      .forEach((lesson) => {
        const dateKey = getDateKeyInTaipei(new Date(lesson.startAt));
        const weekday = getWeekdayFromDateKey(dateKey);
        counts.set(weekday, (counts.get(weekday) || 0) + 1);
      });
    const sorted = [...weekdays].sort((a, b) => {
      const diff = (counts.get(b) || 0) - (counts.get(a) || 0);
      if (diff !== 0) {
        return diff;
      }
      return weekdays.indexOf(a) - weekdays.indexOf(b);
    });
    return (counts.get(sorted[0]) || 0) > 0 ? sorted[0] : "mon";
  }

  function seedWeeklyLessonsForStudent(studentCode, coachCode) {
    const todayKey = getDateKeyInTaipei(new Date());
    const defaultWeekday = pickDefaultLessonWeekdayForCoach(coachCode);
    const firstCycleDay = getNextDateKeyByWeekday(todayKey, defaultWeekday, true);
    const defaultTime = pickDefaultLessonTimeForCoach(coachCode);
    const dayOffsets = [-7, 0, 7, 14, 21, 28, 35, 42];

    dayOffsets.forEach((offset) => {
      const targetDateKey = addDays(firstCycleDay, offset);
      const existed = state.lessons.some(
        (lesson) =>
          lesson.studentCode === studentCode &&
          lesson.sourceType === "REGULAR" &&
          getDateKeyInTaipei(new Date(lesson.startAt)) === targetDateKey
      );
      if (existed) {
        return;
      }
      state.lessons.push(makeLesson(newId("L"), studentCode, coachCode, targetDateKey, defaultTime));
    });
  }

  function ensureStudentProfile(studentCode, coachCode, options = {}) {
    const { studentName = "", studentEmail = "", silentMode = true } = options;
    const normalizedStudentCode = normalizeParticipantCode(studentCode);
    if (!normalizedStudentCode) {
      return null;
    }
    const existed = getStudentByCode(normalizedStudentCode);
    if (existed) {
      const nextName = String(studentName || "").trim();
      if (nextName && existed.name !== nextName) {
        existed.name = nextName;
      }
      const nextCoachCode = normalizeParticipantCode(coachCode) || existed.coachCode;
      if (nextCoachCode && existed.coachCode !== nextCoachCode) {
        existed.coachCode = nextCoachCode;
      }
      const nextEmail = normalizeEmailList(studentEmail).join(", ");
      if (nextEmail && existed.email !== nextEmail) {
        existed.email = nextEmail;
      }
      return existed;
    }

    const normalizedCoachCode = normalizeParticipantCode(coachCode) || "CH001";
    ensureCoachProfile(normalizedCoachCode);

    const fallbackEmail = String(window.APP_CONFIG?.defaultNotifyEmail || "").trim();
    const normalizedEmail = normalizeEmailList(studentEmail).join(", ");
    const created = {
      code: normalizedStudentCode,
      name: String(studentName || "").trim() || `${normalizedStudentCode} 學生`,
      coachCode: normalizedCoachCode,
      email: normalizedEmail || fallbackEmail,
      chargeStartCount: 0,
      paidThroughCount: 0,
      paymentStatus: "unknown",
      paymentNote: "",
      paymentConfirmedAt: "",
      paymentConfirmedBy: "",
      chargeReminderLogs: []
    };
    state.students.push(created);
    seedWeeklyLessonsForStudent(created.code, created.coachCode);
    ensureLessonCalendarEventIds();
    ensureParticipantEmails();
    ensureStudentBillingProfiles();
    addLog(`[同步] 已建立學生代碼 ${created.code}（教練 ${created.coachCode}）來源資料。`);
    saveState();
    return created;
  }

  function syncCoachflowRosterFromPayload(rawPayload, sourceLabel) {
    const payload = rawPayload && typeof rawPayload === "object" && rawPayload.data && typeof rawPayload.data === "object"
      ? rawPayload.data
      : rawPayload;
    const sourceStudents = Array.isArray(payload?.students) ? payload.students : [];
    const sourceCoaches = Array.isArray(payload?.coaches) ? payload.coaches : [];
    if (!sourceStudents.length && !sourceCoaches.length) {
      return false;
    }

    const coachCodeById = new Map();
    sourceCoaches.forEach((coach) => {
      const status = String(coach?.status || "active").toLowerCase();
      if (status === "inactive") {
        return;
      }
      const code = normalizeParticipantCode(
        coach?.accessCode || coach?.access_code || coach?.code || coach?.coachCode || coach?.coach_code
      );
      if (!code) {
        return;
      }
      const id = String(coach?.id || coach?.coach_id || "").trim();
      if (id) {
        coachCodeById.set(id, code);
      }
      const coachName = String(coach?.name || coach?.coach_name || coach?.coachName || `${code} 教練`).trim();
      ensureCoachProfile(code, coachName);
    });

    let changed = false;
    sourceStudents.forEach((student) => {
      const status = String(student?.status || "active").toLowerCase();
      if (status === "inactive") {
        return;
      }
      const studentCode = normalizeParticipantCode(
        student?.accessCode || student?.access_code || student?.code || student?.studentCode || student?.student_code
      );
      if (!studentCode) {
        return;
      }
      const studentName = String(
        student?.name || student?.student_name || student?.studentName || `${studentCode} 學生`
      ).trim();
      const studentEmail = String(
        student?.email ||
        student?.studentEmail ||
        student?.student_email ||
        student?.notifyEmail ||
        student?.notify_email ||
        student?.parentEmail ||
        student?.parent_email ||
        ""
      ).trim();
      const primaryCoachId = String(
        student?.primaryCoachId || student?.primary_coach_id || student?.coachId || student?.coach_id || ""
      ).trim();
      const directCoachCode = normalizeParticipantCode(
        student?.primaryCoachCode ||
        student?.primary_coach_code ||
        student?.coachCode ||
        student?.coach_code ||
        student?.primaryCoachAccessCode ||
        student?.primary_coach_access_code
      );
      const mappedCoachCode = directCoachCode || coachCodeById.get(primaryCoachId) || "";
      const fallbackCoachCode = mappedCoachCode || normalizeParticipantCode(state.coaches[0]?.code) || "CH001";
      const before = JSON.stringify(getStudentByCode(studentCode) || {});
      ensureCoachProfile(fallbackCoachCode);
      ensureStudentProfile(studentCode, fallbackCoachCode, { studentName, studentEmail, silentMode: true });
      const after = JSON.stringify(getStudentByCode(studentCode) || {});
      if (before !== after) {
        changed = true;
      }
      if (!state.lessons.some((lesson) => lesson.studentCode === studentCode && lesson.sourceType === "REGULAR")) {
        seedWeeklyLessonsForStudent(studentCode, fallbackCoachCode);
        changed = true;
      }
    });

    if (changed) {
      ensureLessonCalendarEventIds();
      ensureParticipantEmails();
      ensureStudentBillingProfiles();
      saveState();
      addLog(`[同步] 已從 ${sourceLabel} 同步 coachflow 學生/教練資料。`);
    }
    return changed;
  }

  function syncCoachflowRosterFromLocalState() {
    try {
      const raw = localStorage.getItem("coachflow-v2-state");
      if (!raw) {
        return false;
      }
      const parsed = JSON.parse(raw);
      if (!parsed || typeof parsed !== "object") {
        return false;
      }
      return syncCoachflowRosterFromPayload(parsed, "本機 CoachFlow");
    } catch (error) {
      console.warn("Failed to sync roster from local CoachFlow state:", error);
      return false;
    }
  }

  async function syncCoachflowRosterFromCloud() {
    if (!getCoachflowAppsScriptUrl()) {
      return false;
    }
    const attempts = [
      { action: "bootstrap", method: "GET" },
      { action: "bootstrapAdmin", method: "GET" },
      { action: "bootstrap", method: "POST" },
      { action: "bootstrapAdmin", method: "POST" }
    ];
    for (const attempt of attempts) {
      try {
        const payload = await callCoachflowApi(attempt.action, {}, attempt.method);
        if (syncCoachflowRosterFromPayload(payload, "雲端 CoachFlow")) {
          return true;
        }
      } catch (error) {
        console.warn(`CoachFlow roster sync failed via ${attempt.action}:`, error);
      }
    }
    return false;
  }

  function buildLessonEventPayload(lesson, extra = {}) {
    const student = getStudentByCode(lesson.studentCode);
    const coach = getCoachByCode(lesson.coachCode);
    const startAt = lesson.startAt;
    const endAt = new Date(new Date(startAt).getTime() + 60 * 60 * 1000).toISOString();
    const calendarPayload = getCalendarPayloadForCoach(lesson.coachCode);
    return {
      lessonId: lesson.id,
      eventId: lesson.calendarEventId || "",
      title: `${lesson.sourceType === "MAKEUP" ? "補課" : "課程"} ${student?.name || lesson.studentCode} / ${coach?.name || lesson.coachCode}`,
      startAt,
      endAt,
      studentCode: lesson.studentCode,
      studentName: student?.name || "",
      coachCode: lesson.coachCode,
      coachName: coach?.name || "",
      sourceType: lesson.sourceType,
      ...calendarPayload,
      ...extra
    };
  }

  function addSingleEventDeleteGuards(payload) {
    const guarded = { ...(payload || {}) };
    const startAt = String(guarded.occurrenceStartAt || guarded.lessonStartAt || guarded.startAt || "").trim();
    const endAt = String(guarded.occurrenceEndAt || guarded.lessonEndAt || guarded.endAt || "").trim();
    const startDate = new Date(startAt);
    const dateKey = !Number.isNaN(startDate.getTime()) ? getDateKeyInTaipei(startDate) : String(guarded.dateKey || guarded.occurrenceDate || "").trim();
    return {
      ...guarded,
      calendarEventId: guarded.calendarEventId || guarded.eventId || "",
      eventId: guarded.eventId || guarded.calendarEventId || "",
      deleteScope: "single",
      deleteMode: "single",
      singleEventOnly: true,
      strictSingleOccurrence: true,
      occurrenceStartAt: startAt,
      lessonStartAt: startAt,
      occurrenceEndAt: endAt,
      lessonEndAt: endAt,
      occurrenceDate: dateKey,
      dateKey
    };
  }

  function buildSingleLessonDeletePayload(lesson, extra = {}) {
    const basePayload = buildLessonEventPayload(lesson, extra);
    return addSingleEventDeleteGuards(basePayload);
  }

  function isDeleteSingleEventEndpointUnsupported(message) {
    return /unsupported action|invalid action|deleteSingleEvent/i.test(String(message || ""));
  }

  function queueSingleEventDeleteCompensation(lesson, payload, message) {
    enqueueCompensationTask("deleteSingleEvent", payload, message);
    addLog(`[Google日曆] 單堂刪除暫未完成，已排入補償任務：${lesson.id}（${message}）`);
    saveState();
  }

  function isNormalLeaveCalendarCreateBlocked(payload = {}) {
    const reason = String(payload.reason || "").trim().toLowerCase();
    const attendanceStatus = String(payload.attendanceStatus || "").trim().toLowerCase();
    const leaveType = String(payload.leaveType || payload.type || "").trim().toLowerCase();
    return reason === "student_normal_leave" ||
      attendanceStatus === "leave-normal" ||
      leaveType === "normal_leave" ||
      leaveType === "student_normal_leave";
  }

  function getCalendarPayloadForCoach(coachCodeInput) {
    const coachCode = normalizeParticipantCode(coachCodeInput || activeCoachCode);
    const coach = getCoachByCode(coachCode) || {};
    const configuredByCoach = window.APP_CONFIG?.coachCalendarIds && typeof window.APP_CONFIG.coachCalendarIds === "object"
      ? window.APP_CONFIG.coachCalendarIds[coachCode]
      : "";
    const params = new URLSearchParams(window.location.search);
    const calendarId = String(
      params.get("calendarId") ||
      params.get("coachCalendarId") ||
      coach.calendarId ||
      coach.calendar_id ||
      coach.coachCalendarId ||
      coach.coach_calendar_id ||
      coach.googleCalendarId ||
      coach.google_calendar_id ||
      configuredByCoach ||
      window.APP_CONFIG?.coachCalendarId ||
      window.APP_CONFIG?.googleCalendarId ||
      window.APP_CONFIG?.calendarId ||
      ""
    ).trim();
    return calendarId
      ? {
        calendarId,
        coachCalendarId: calendarId
      }
      : {};
  }

  function normalizeCalendarEventId(eventId) {
    const raw = String(eventId || "").trim();
    if (!raw) {
      return "";
    }
    return raw.endsWith("@google.com") ? raw.slice(0, raw.length - "@google.com".length) : raw;
  }

  function isGeneratedLocalCalendarEventId(eventId) {
    return /^GCAL_/i.test(String(eventId || "").trim());
  }

  function getDaySyncRange(dateKey) {
    return {
      startAt: makeTaipeiDateTime(dateKey, "00:00").toISOString(),
      endAt: makeTaipeiDateTime(addDays(dateKey, 1), "00:00").toISOString()
    };
  }

  function getSyncRangeFromLessons(lessons) {
    if (!Array.isArray(lessons) || !lessons.length) {
      return null;
    }
    const sorted = [...lessons].sort((a, b) => new Date(a.startAt) - new Date(b.startAt));
    const firstDateKey = getDateKeyInTaipei(new Date(sorted[0].startAt));
    const lastDateKey = getDateKeyInTaipei(new Date(sorted[sorted.length - 1].startAt));
    return {
      startAt: makeTaipeiDateTime(firstDateKey, "00:00").toISOString(),
      endAt: makeTaipeiDateTime(addDays(lastDateKey, 1), "00:00").toISOString()
    };
  }

  function getMonthSyncRange(monthStartDate) {
    const monthStart = getMonthStart(monthStartDate || new Date());
    const nextMonthStart = shiftMonth(monthStart, 1);
    return {
      startAt: monthStart.toISOString(),
      endAt: nextMonthStart.toISOString()
    };
  }

  function normalizeLooseText(value) {
    return String(value || "")
      .toLowerCase()
      .replace(/[\s\-_/|｜·.,，。;；:：'"`~!@#$%^&*()+={}[\]<>?！￥…（）【】「」『』]/g, "");
  }

  function scoreLooseNameMatch(eventTitle, studentName) {
    const title = normalizeLooseText(eventTitle);
    const name = normalizeLooseText(studentName);
    if (!title || !name) {
      return 0;
    }
    if (title === name) {
      return 4;
    }
    if (title.includes(name) || name.includes(title)) {
      return 3;
    }
    const nameChars = [...new Set(Array.from(name))];
    let overlap = 0;
    nameChars.forEach((ch) => {
      if (title.includes(ch)) {
        overlap += 1;
      }
    });
    if (overlap >= Math.min(2, nameChars.length)) {
      return 2;
    }
    return 0;
  }

  function hashTextForCode(value) {
    const text = String(value || "").trim();
    let hash = 0;
    for (let index = 0; index < text.length; index += 1) {
      hash = ((hash << 5) - hash + text.charCodeAt(index)) >>> 0;
    }
    return hash.toString(36).toUpperCase().padStart(5, "0").slice(0, 5);
  }

  function readCalendarDateTime(value) {
    if (!value) {
      return "";
    }
    if (typeof value === "string") {
      return value;
    }
    if (typeof value === "object") {
      return String(value.dateTime || value.date || value.value || "").trim();
    }
    return "";
  }

  function normalizeGoogleCalendarEvent(rawEvent) {
    if (!rawEvent || typeof rawEvent !== "object") {
      return null;
    }
    const startAtText = readCalendarDateTime(
      rawEvent.startAt ||
      rawEvent.startTime ||
      rawEvent.start ||
      rawEvent.startDateTime
    );
    const startDate = new Date(startAtText);
    if (!startAtText || Number.isNaN(startDate.getTime())) {
      return null;
    }
    const endAtText = readCalendarDateTime(
      rawEvent.endAt ||
      rawEvent.endTime ||
      rawEvent.end ||
      rawEvent.endDateTime
    );
    const endDate = new Date(endAtText);
    const fallbackEndDate = new Date(startDate.getTime() + 60 * 60 * 1000);
    const eventId = String(
      rawEvent.eventId ||
      rawEvent.calendarEventId ||
      rawEvent.id ||
      rawEvent.iCalUID ||
      ""
    ).trim();
    const title = String(rawEvent.title || rawEvent.summary || rawEvent.name || rawEvent.subject || "").trim();
    return {
      ...rawEvent,
      eventId,
      id: eventId || rawEvent.id || "",
      title,
      description: String(rawEvent.description || rawEvent.notes || "").trim(),
      location: String(rawEvent.location || "").trim(),
      startAt: startDate.toISOString(),
      endAt: Number.isNaN(endDate.getTime()) ? fallbackEndDate.toISOString() : endDate.toISOString()
    };
  }

  function ensureGoogleEventStudentProfile(event, coachCode) {
    const title = String(event?.title || "").trim();
    const fallbackKey = normalizeCalendarEventId(event?.eventId || event?.id || "") || `${event?.startAt || ""}-${title}`;
    const baseCode = `G${hashTextForCode(normalizeLooseText(title) || fallbackKey)}`;
    let code = baseCode;
    let suffix = 1;
    while (state.students.some((student) => student.code === code && student.coachCode !== coachCode)) {
      suffix += 1;
      code = `${baseCode}${suffix}`;
    }
    const existed = getStudentByCode(code);
    if (existed) {
      if (title && existed.name !== title) {
        existed.name = title;
      }
      existed.coachCode = coachCode;
      existed.sourceType = existed.sourceType || "GOOGLE_CALENDAR";
      return existed;
    }
    const created = {
      code,
      name: title || "Google 日曆課程",
      coachCode,
      email: "",
      chargeStartCount: 0,
      paidThroughCount: 0,
      paymentStatus: "unknown",
      paymentNote: "",
      paymentConfirmedAt: "",
      paymentConfirmedBy: "",
      chargeReminderLogs: [],
      sourceType: "GOOGLE_CALENDAR"
    };
    state.students.push(created);
    addLog(`[日曆同步] 已建立 Google 日曆來源學生 ${created.name}（${created.code}）。`);
    return created;
  }

  function resolveStudentCodeFromCalendarEvent(event, coachCode) {
    const searchable = `${event?.title || ""}\n${event?.description || ""}\n${event?.location || ""}`.toLowerCase();
    const eventTitle = String(event?.title || "");
    const students = state.students.filter((student) => student.coachCode === coachCode);
    const aliases = window.APP_CONFIG?.calendarStudentAliases && typeof window.APP_CONFIG.calendarStudentAliases === "object"
      ? window.APP_CONFIG.calendarStudentAliases
      : {};
    const aliasKey = normalizeLooseText(eventTitle);
    const aliasTarget = Object.entries(aliases).find(([alias]) => normalizeLooseText(alias) === aliasKey)?.[1];
    if (aliasTarget) {
      const targetText = String(aliasTarget || "").trim();
      const targetLoose = normalizeLooseText(targetText);
      const matchedByAlias = students.find((student) => (
        String(student.code || "").trim().toLowerCase() === targetText.toLowerCase() ||
        normalizeLooseText(student.name) === targetLoose
      ));
      if (matchedByAlias) {
        return matchedByAlias.code;
      }
    }
    let bestCode = "";
    let bestScore = 0;
    let hasTie = false;
    students.forEach((student) => {
      const code = String(student.code || "").trim().toLowerCase();
      const name = String(student.name || "").trim().toLowerCase();
      let score = 0;
      if (code && searchable.includes(code)) {
        score += 4;
      }
      if (name && searchable.includes(name)) {
        score += 3;
      }
      if (name) {
        score += scoreLooseNameMatch(eventTitle, name);
      }
      if (score > bestScore) {
        bestScore = score;
        bestCode = student.code;
        hasTie = false;
      } else if (score > 0 && score === bestScore) {
        hasTie = true;
      }
    });
    if (bestScore <= 0 || hasTie) {
      return "";
    }
    return bestCode;
  }

  function isGoogleSyncLesson(lesson) {
    return lesson?.sourceType === "REGULAR" || lesson?.sourceType === "GOOGLE_CALENDAR";
  }

  function isCalendarEventLikelySameLesson(event, lesson, requestedEventId) {
    const normalizedRequestedId = normalizeCalendarEventId(requestedEventId || lesson?.calendarEventId || "");
    const normalizedEventId = normalizeCalendarEventId(event?.eventId || event?.id || "");
    const lessonStartTime = new Date(lesson?.startAt || "").getTime();
    const eventStartTime = new Date(event?.startAt || "").getTime();
    if (!Number.isFinite(lessonStartTime) || !Number.isFinite(eventStartTime)) {
      return { matched: false, score: 0 };
    }

    const diffMs = Math.abs(eventStartTime - lessonStartTime);
    const sameEventId = normalizedRequestedId && normalizedEventId && normalizedRequestedId === normalizedEventId;
    const student = getStudentByCode(lesson.studentCode);
    const searchable = `${event?.title || ""}\n${event?.description || ""}\n${event?.location || ""}`;
    const looseSearchable = normalizeLooseText(searchable);
    const studentCode = normalizeLooseText(lesson.studentCode);
    const studentName = normalizeLooseText(student?.name || "");
    const hasStudentText = Boolean(
      (studentCode && looseSearchable.includes(studentCode)) ||
      (studentName && looseSearchable.includes(studentName)) ||
      (student?.name && scoreLooseNameMatch(event?.title || "", student.name) > 0)
    );

    let score = 0;
    if (sameEventId) {
      score += 10;
    }
    if (diffMs <= 2 * 60 * 1000) {
      score += 8;
    } else if (diffMs <= 15 * 60 * 1000) {
      score += 4;
    }
    if (hasStudentText) {
      score += 4;
    }

    return {
      matched: (sameEventId && diffMs <= 15 * 60 * 1000) || (hasStudentText && diffMs <= 2 * 60 * 1000),
      score
    };
  }

  async function resolveSingleCalendarEventForLesson(lesson, requestedEventId) {
    if (!lesson?.startAt || !getAppsScriptUrl()) {
      return null;
    }
    const startDate = new Date(lesson.startAt);
    if (Number.isNaN(startDate.getTime())) {
      return null;
    }
    const calendarPayload = getCalendarPayloadForCoach(lesson.coachCode);
    if (!calendarPayload.calendarId && !calendarPayload.coachCalendarId) {
      return null;
    }

    const dateKey = getDateKeyInTaipei(startDate);
    const listResult = await callAppsScriptApi("listEvents", {
      ...getDaySyncRange(dateKey),
      ...calendarPayload,
      coachCode: lesson.coachCode,
      dateKey,
      occurrenceDate: dateKey
    });
    const candidates = (Array.isArray(listResult?.events) ? listResult.events : [])
      .map(normalizeGoogleCalendarEvent)
      .filter(Boolean)
      .map((event) => ({
        event,
        match: isCalendarEventLikelySameLesson(event, lesson, requestedEventId)
      }))
      .filter((item) => item.match.matched)
      .sort((a, b) => b.match.score - a.match.score);

    if (!candidates.length) {
      return null;
    }
    if (candidates.length > 1 && candidates[0].match.score === candidates[1].match.score) {
      return null;
    }
    return candidates[0].event;
  }

  function pickClosestRegularLessonForStudent(studentCode, coachCode, usedLessonIds, targetStartAt, maxDistanceMs = Infinity) {
    const targetTime = new Date(targetStartAt).getTime();
    const candidates = state.lessons
      .filter(
        (lesson) =>
          lesson.studentCode === studentCode &&
          lesson.coachCode === coachCode &&
          isGoogleSyncLesson(lesson) &&
          !usedLessonIds.has(lesson.id)
      )
      .sort((a, b) => Math.abs(new Date(a.startAt).getTime() - targetTime) - Math.abs(new Date(b.startAt).getTime() - targetTime));
    const closest = candidates[0] || null;
    if (!closest) {
      return null;
    }
    const distance = Math.abs(new Date(closest.startAt).getTime() - targetTime);
    return distance <= maxDistanceMs ? closest : null;
  }

  function alignCoachLessonsWithGoogleEvents(coachCode, events) {
    const usedLessonIds = new Set();
    const eventIdToLessons = new Map();
    state.lessons.forEach((lesson) => {
      if (lesson.coachCode !== coachCode || !isGoogleSyncLesson(lesson)) {
        return;
      }
      const key = normalizeCalendarEventId(lesson.calendarEventId);
      if (key) {
        if (!eventIdToLessons.has(key)) {
          eventIdToLessons.set(key, []);
        }
        eventIdToLessons.get(key).push(lesson);
      }
    });

    let matchedEvents = 0;
    let updatedStart = 0;
    let relinkedEventId = 0;
    let createdLessons = 0;
    let createdGoogleStudents = 0;
    let skippedUnmatchedStudent = 0;
    let skippedNonLessonEvent = 0;

    const maxRelinkDistanceMs = Math.max(1, Number(window.APP_CONFIG?.googleRelinkMaxHours || 12)) * 60 * 60 * 1000;
    const sortedEvents = [...events]
      .map(normalizeGoogleCalendarEvent)
      .filter(Boolean)
      .sort((a, b) => new Date(a.startAt) - new Date(b.startAt));
    sortedEvents.forEach((event) => {
      const startAt = String(event?.startAt || "").trim();
      const eventId = String(event?.eventId || "").trim();
      if (!startAt || !eventId) {
        skippedNonLessonEvent += 1;
        return;
      }
      const durationMs = new Date(String(event?.endAt || "")).getTime() - new Date(startAt).getTime();
      if (!Number.isFinite(durationMs) || durationMs <= 0 || durationMs > 4 * 60 * 60 * 1000) {
        skippedNonLessonEvent += 1;
        return;
      }

      let studentCode = resolveStudentCodeFromCalendarEvent(event, coachCode);
      if (!studentCode) {
        const googleStudent = ensureGoogleEventStudentProfile(event, coachCode);
        studentCode = googleStudent?.code || "";
        if (studentCode) {
          createdGoogleStudents += 1;
        } else {
          skippedUnmatchedStudent += 1;
          return;
        }
      }

      const normalizedEventId = normalizeCalendarEventId(eventId);
      const sameEventLessons = eventIdToLessons.get(normalizedEventId) || [];
      let lesson = sameEventLessons
        .filter((item) => (
          !usedLessonIds.has(item.id) &&
          Math.abs(new Date(item.startAt).getTime() - new Date(startAt).getTime()) <= maxRelinkDistanceMs
        ))
        .sort((a, b) => Math.abs(new Date(a.startAt).getTime() - new Date(startAt).getTime()) - Math.abs(new Date(b.startAt).getTime() - new Date(startAt).getTime()))[0] || null;
      if (!lesson) {
        lesson = pickClosestRegularLessonForStudent(studentCode, coachCode, usedLessonIds, startAt, maxRelinkDistanceMs);
      }
      if (!lesson) {
        const dateKey = getDateKeyInTaipei(new Date(startAt));
        const timeText = getTimeText(startAt);
        lesson = makeLesson(newId("L"), studentCode, coachCode, dateKey, timeText);
        lesson.sourceType = "GOOGLE_CALENDAR";
        state.lessons.push(lesson);
        createdLessons += 1;
      }

      usedLessonIds.add(lesson.id);
      const beforeStartAt = String(lesson.startAt || "");
      const beforeEventId = normalizeCalendarEventId(lesson.calendarEventId);
      lesson.startAt = new Date(startAt).toISOString();
      lesson.coachCode = coachCode;
      lesson.studentCode = studentCode;
      if (lesson.attendanceStatus !== "leave-normal") {
        lesson.calendarOccupied = true;
      }
      if (lesson.attendanceStatus === "calendar-removed") {
        const previous = lesson.beforeCalendarRemoved || {};
        lesson.attendanceStatus = previous.attendanceStatus || "scheduled";
        delete lesson.calendarRemovedAt;
        delete lesson.calendarRemovedReason;
      }
      lesson.calendarEventId = eventId;

      if (beforeStartAt !== lesson.startAt) {
        updatedStart += 1;
      }
      if (beforeEventId !== normalizeCalendarEventId(lesson.calendarEventId)) {
        relinkedEventId += 1;
      }
      matchedEvents += 1;
    });

    return {
      totalEvents: sortedEvents.length,
      matchedEvents,
      updatedStart,
      relinkedEventId,
      createdLessons,
      createdGoogleStudents,
      skippedUnmatchedStudent,
      skippedNonLessonEvent,
      matchedLessonIds: Array.from(usedLessonIds)
    };
  }

  async function tryDeleteCalendarEventForLesson(lesson, reason, strictMode) {
    const eventId = String(lesson.calendarEventId || "").trim();
    if (!eventId) {
      return true;
    }
    if (isGeneratedLocalCalendarEventId(eventId)) {
      lesson.calendarEventId = "";
      addLog(`[Google日曆] ${lesson.id} 只有本地暫存事件ID，已略過 Google 刪除。`);
      saveState();
      return true;
    }

    let deletePayload = buildSingleLessonDeletePayload(lesson, { reason });
    try {
      const resolvedEvent = await resolveSingleCalendarEventForLesson(lesson, eventId);
      if (!resolvedEvent) {
        const safeMessage = "無法在 Google 日曆確認這堂課的單一事件，已停止刪除以避免誤刪整個週期課程。請先同步 Google 日曆後再請假。";
        queueSingleEventDeleteCompensation(lesson, deletePayload, safeMessage);
        return Boolean(strictMode);
      }
      deletePayload = buildSingleLessonDeletePayload(lesson, {
        reason,
        eventId: resolvedEvent.eventId || eventId,
        calendarEventId: resolvedEvent.eventId || eventId,
        matchedEventStartAt: resolvedEvent.startAt || "",
        matchedEventEndAt: resolvedEvent.endAt || ""
      });
    } catch (error) {
      const message = String(error?.message || "未知錯誤");
      queueSingleEventDeleteCompensation(lesson, deletePayload, `單堂刪除前檢查失敗：${message}`);
      return Boolean(strictMode);
    }

    try {
      const result = await callAppsScriptApi("deleteSingleEvent", deletePayload);
      lesson.calendarEventId = "";
      addLog(`[Google日曆] 已刪除事件 ${eventId}${result.mock ? "（模擬）" : ""}`);
      saveState();
      return true;
    } catch (error) {
      const message = String(error?.message || "未知錯誤");
      const alreadyDeleted = /event not found|already be deleted|already deleted/i.test(message);
      if (alreadyDeleted) {
        lesson.calendarEventId = "";
        addLog(`[Google日曆] 事件 ${eventId} 已不存在，視為刪除完成。`);
        saveState();
        return true;
      }
      if (strictMode && !isDeleteSingleEventEndpointUnsupported(message)) {
        throw new Error(`Google 日曆單堂刪除失敗：${message}`);
      }
      queueSingleEventDeleteCompensation(lesson, deletePayload, message);
      return Boolean(strictMode);
    }
  }

  async function checkCalendarEventExists(lesson) {
    const eventId = String(lesson.calendarEventId || "").trim();
    if (!eventId) {
      return { exists: true, missingEventId: true };
    }
    if (isGeneratedLocalCalendarEventId(eventId)) {
      return { exists: true, generatedLocalId: true };
    }
    try {
      const result = await callAppsScriptApi("checkEvent", buildLessonEventPayload(lesson, { eventId }));
      if (result?.exists === true) {
        return { exists: true };
      }
      if (result?.exists === false) {
        return { exists: false, explicitNotFound: true };
      }
      return { exists: true, error: true, message: "checkEvent 回傳格式缺少 exists，已停止自動刪除。" };
    } catch (error) {
      const message = String(error?.message || "未知錯誤");
      if (/unsupported|invalid action/i.test(message)) {
        return { exists: true, unsupported: true, message };
      }
      return { exists: true, error: true, message };
    }
  }

  async function syncCoachCalendarEvents() {
    if (!requireCoachWriteAccess()) {
      return;
    }
    if (!activeCoachCode) {
      alert("請先登入教練。");
      return;
    }
    const meta = getMonthMeta(coachCalendarMonthStart);
    const monthPrefix = `${String(meta.year).padStart(4, "0")}-${String(meta.month).padStart(2, "0")}-`;
    const monthRange = getMonthSyncRange(coachCalendarMonthStart);
    const collectMonthLessons = () => state.lessons
      .filter(
        (lesson) =>
          lesson.coachCode === activeCoachCode &&
          lesson.calendarOccupied &&
          getDateKeyInTaipei(new Date(lesson.startAt)).startsWith(monthPrefix)
      )
      .sort((a, b) => new Date(a.startAt) - new Date(b.startAt));

    let lessons = collectMonthLessons();
    const syncScopeText = `${String(meta.year).padStart(4, "0")}/${String(meta.month).padStart(2, "0")} 當月`;
    const calendarPayload = getCalendarPayloadForCoach(activeCoachCode);

    const confirmed = window.confirm(
      `本次將以 Google 日曆同步 ${syncScopeText}。\n會先讀取 Google 當月事件並更新請假系統課程；若本地課程在 Google 當月找不到對應事件，會列入同步移除確認。\n是否繼續？`
    );
    if (!confirmed) {
      if (el.coachCalendarSyncText) {
        el.coachCalendarSyncText.textContent = "已取消同步。";
      }
      return;
    }

    if (el.coachCalendarSyncBtn) {
      el.coachCalendarSyncBtn.disabled = true;
    }
    if (el.coachCalendarSyncText) {
      el.coachCalendarSyncText.textContent = `同步中... 讀取 Google ${syncScopeText}`;
    }

    let checkedCount = 0;
    let removedCount = 0;
    let alreadyOkCount = 0;
    let unsupportedCount = 0;
    let missingIdCount = 0;
    let generatedLocalIdCount = 0;
    let generatedLocalPendingRemoveCount = 0;
    let explicitNotFoundCount = 0;
    let pastMissingKeptCount = 0;
    let errorCount = 0;
    const pendingRemoveLessons = [];
    let alignedSummary = "";
    let alignedLessonIdSet = null;
    let listEventsLoaded = false;
    let googleEventCount = 0;
    let syncedCalendarId = "";

    const batchSize = Math.max(1, Math.min(8, Number(window.APP_CONFIG?.calendarSyncBatchSize || 4)));
    const updateProgressText = () => {
      if (el.coachCalendarSyncText) {
        el.coachCalendarSyncText.textContent = `同步中... ${checkedCount}/${lessons.length}`;
      }
    };

    try {
      try {
        const listResult = await callAppsScriptApi("listEvents", {
          ...monthRange,
          ...calendarPayload,
          coachCode: activeCoachCode
        });
        const events = Array.isArray(listResult?.events) ? listResult.events : [];
        listEventsLoaded = true;
        googleEventCount = events.length;
        syncedCalendarId = String(listResult?.calendarId || calendarPayload.calendarId || "").trim();
        const alignStats = alignCoachLessonsWithGoogleEvents(activeCoachCode, events);
        alignedLessonIdSet = new Set(Array.isArray(alignStats.matchedLessonIds) ? alignStats.matchedLessonIds : []);
        if (alignStats.totalEvents > 0) {
          alignedSummary = `Google對齊：calendar ${syncedCalendarId || "-"}，抓到 ${alignStats.totalEvents} 筆，匹配 ${alignStats.matchedEvents} 筆，調整時間 ${alignStats.updatedStart} 筆，重綁事件ID ${alignStats.relinkedEventId} 筆，新增課程 ${alignStats.createdLessons} 筆，Google來源學生 ${alignStats.createdGoogleStudents} 筆，未匹配學生 ${alignStats.skippedUnmatchedStudent} 筆，非課程事件 ${alignStats.skippedNonLessonEvent} 筆`;
          addLog(`[日曆同步] ${alignedSummary}`);
        } else {
          alignedSummary = `Google對齊：calendar ${syncedCalendarId || "-"} 查無事件。`;
          addLog(`[日曆同步] ${alignedSummary}`);
        }
        lessons = collectMonthLessons();
      } catch (error) {
        const message = String(error?.message || "未知錯誤");
        addLog(`[日曆同步] Google 對齊略過：${message}`);
      }

      if (listEventsLoaded && alignedLessonIdSet) {
        for (const lesson of lessons) {
          checkedCount += 1;
          if (alignedLessonIdSet.has(lesson.id)) {
            alreadyOkCount += 1;
            continue;
          }
          const hasGeneratedLocalEventId = isGeneratedLocalCalendarEventId(lesson.calendarEventId);
          if (hasGeneratedLocalEventId) {
            generatedLocalIdCount += 1;
          }
          if (!String(lesson.calendarEventId || "").trim()) {
            missingIdCount += 1;
          }
          if (!hasGeneratedLocalEventId) {
            explicitNotFoundCount += 1;
          }
          pendingRemoveLessons.push(lesson);
          if (hasGeneratedLocalEventId) {
            generatedLocalPendingRemoveCount += 1;
          }
        }
        updateProgressText();
      } else if (lessons.length) {
        for (let index = 0; index < lessons.length; index += batchSize) {
          const batch = lessons.slice(index, index + batchSize);
          const batchResults = await Promise.allSettled(batch.map((lesson) => checkCalendarEventExists(lesson)));
          for (let offset = 0; offset < batch.length; offset += 1) {
            checkedCount += 1;
            const lesson = batch[offset];
            const result = batchResults[offset];
            if (result.status === "rejected") {
              errorCount += 1;
              const message = String(result.reason?.message || result.reason || "未知錯誤");
              addLog(`[日曆同步] 檢查失敗 ${lesson.id}：${message}`);
              continue;
            }
            const check = result.value || {};
            if (check.unsupported) {
              unsupportedCount += 1;
              continue;
            }
            if (check.generatedLocalId) {
              generatedLocalIdCount += 1;
              continue;
            }
            if (check.error) {
              errorCount += 1;
              addLog(`[日曆同步] 檢查失敗 ${lesson.id}：${check.message}`);
              continue;
            }
            if (check.exists === true) {
              alreadyOkCount += 1;
              continue;
            }
            if (check.missingEventId) {
              missingIdCount += 1;
              continue;
            }
            if (!check.explicitNotFound) {
              errorCount += 1;
              addLog(`[日曆同步] 檢查結果不明，已略過 ${lesson.id}`);
              continue;
            }
            explicitNotFoundCount += 1;
            pendingRemoveLessons.push(lesson);
          }
          updateProgressText();
        }
      }

      if (pendingRemoveLessons.length) {
        const generatedHint = generatedLocalPendingRemoveCount
          ? `\n其中 ${generatedLocalPendingRemoveCount} 堂是本機暫存課（無 Google 事件ID），將一併移除。`
          : "";
        const explicitHint = explicitNotFoundCount
          ? `\n另有 ${explicitNotFoundCount} 堂為 Google 明確回報不存在。`
          : "";
        const confirmRemoval = window.confirm(
          `本次有 ${pendingRemoveLessons.length} 堂課在 Google ${syncScopeText} 找不到對應事件。${generatedHint}${explicitHint}\n是否要同步移除這 ${pendingRemoveLessons.length} 堂課？`
        );
        if (confirmRemoval) {
          pendingRemoveLessons.forEach((lesson) => {
            markLessonRemovedByCalendar(lesson, "google_calendar_deleted");
            removedCount += 1;
          });
        } else {
          addLog(`[日曆同步] 已取消本次移除（待移除 ${pendingRemoveLessons.length} 堂）。`);
        }
      }

      const googleSummary = listEventsLoaded ? `，Google事件 ${googleEventCount} 筆` : "";
      const summary = `${syncScopeText}：已檢查 ${checkedCount} 堂${googleSummary}，已同步移除 ${removedCount} 堂，正常 ${alreadyOkCount} 堂，錯誤 ${errorCount} 堂${unsupportedCount ? `，端點未支援 ${unsupportedCount} 堂` : ""}${generatedLocalIdCount ? `，本機暫存事件ID ${generatedLocalIdCount} 堂` : ""}${generatedLocalPendingRemoveCount ? `（待移除 ${generatedLocalPendingRemoveCount} 堂）` : generatedLocalIdCount ? "（已略過）" : ""}${missingIdCount ? `，缺事件ID ${missingIdCount} 堂` : ""}${pastMissingKeptCount ? `，歷史課程缺事件 ${pastMissingKeptCount} 堂（保留）` : ""}${pendingRemoveLessons.length && removedCount !== pendingRemoveLessons.length ? `，取消移除 ${pendingRemoveLessons.length - removedCount} 堂` : ""}${alignedSummary ? `，${alignedSummary}` : ""}`;
      addLog(`[日曆同步] ${summary}`);
      if (el.coachCalendarSyncText) {
        el.coachCalendarSyncText.textContent = summary;
      }
      saveState();
      renderAll();
    } catch (error) {
      const message = String(error?.message || "未知錯誤");
      addLog(`[日曆同步] 同步失敗：${message}`);
      if (el.coachCalendarSyncText) {
        el.coachCalendarSyncText.textContent = `同步失敗：${message}`;
      }
    } finally {
      if (el.coachCalendarSyncBtn) {
        el.coachCalendarSyncBtn.disabled = false;
      }
    }
  }

  async function syncStudentCalendarEventsFromGoogle() {
    if (!activeStudentCode || !activeCoachCode || !getAppsScriptUrl()) {
      return false;
    }
    const calendarPayload = getCalendarPayloadForCoach(activeCoachCode);
    if (!calendarPayload.calendarId && !calendarPayload.coachCalendarId) {
      return false;
    }
    const monthStarts = [
      getMonthStart(studentCalendarMonthStart || new Date()),
      shiftMonth(studentCalendarMonthStart || new Date(), 1),
      shiftMonth(studentCalendarMonthStart || new Date(), 2)
    ];
    let totalEvents = 0;
    let totalMatched = 0;
    let totalCreated = 0;
    let totalRemovedPlaceholders = 0;
    let syncedAny = false;
    for (const monthStart of monthStarts) {
      const monthRange = getMonthSyncRange(monthStart);
      try {
        const listResult = await callAppsScriptApi("listEvents", {
          ...monthRange,
          ...calendarPayload,
          coachCode: activeCoachCode
        });
        const events = Array.isArray(listResult?.events) ? listResult.events : [];
        const stats = alignCoachLessonsWithGoogleEvents(activeCoachCode, events);
        const matchedLessonIds = new Set(Array.isArray(stats.matchedLessonIds) ? stats.matchedLessonIds : []);
        const meta = getMonthMeta(monthStart);
        const monthPrefix = `${String(meta.year).padStart(4, "0")}-${String(meta.month).padStart(2, "0")}-`;
        state.lessons
          .filter((lesson) => (
            lesson.coachCode === activeCoachCode &&
            lesson.calendarOccupied &&
            getDateKeyInTaipei(new Date(lesson.startAt)).startsWith(monthPrefix) &&
            !matchedLessonIds.has(lesson.id) &&
            isGeneratedLocalCalendarEventId(lesson.calendarEventId)
          ))
          .forEach((lesson) => {
            markLessonRemovedByCalendar(lesson, "student_google_auto_sync");
            totalRemovedPlaceholders += 1;
          });
        totalEvents += stats.totalEvents;
        totalMatched += stats.matchedEvents;
        totalCreated += stats.createdLessons;
        syncedAny = true;
      } catch (error) {
        const message = String(error?.message || "未知錯誤");
        addLog(`[日曆同步] 學生端自動同步略過：${message}`);
      }
    }
    if (syncedAny) {
      ensureLessonCalendarEventIds();
      ensureParticipantEmails();
      ensureStudentBillingProfiles();
      saveState();
      addLog(`[日曆同步] 學生端已自動同步 Google 課程：抓到 ${totalEvents} 筆，匹配 ${totalMatched} 筆，新增 ${totalCreated} 筆，移除暫存課 ${totalRemovedPlaceholders} 筆。`);
    }
    return syncedAny;
  }

  async function syncReadOnlyCoachCalendarFromGoogle(coachCodeInput) {
    const coachCode = normalizeParticipantCode(coachCodeInput || activeCoachCode);
    if (!coachCode || !getAppsScriptUrl()) {
      return false;
    }
    const calendarPayload = getCalendarPayloadForCoach(coachCode);
    if (!calendarPayload.calendarId && !calendarPayload.coachCalendarId) {
      return false;
    }
    const { minMonth, maxMonth } = getReadOnlyMonthBounds();
    const monthStarts = [minMonth, maxMonth];
    let totalEvents = 0;
    let totalMatched = 0;
    let totalCreated = 0;
    let removedPlaceholders = 0;
    let syncedAny = false;
    for (const monthStart of monthStarts) {
      const monthRange = getMonthSyncRange(monthStart);
      const listResult = await callAppsScriptApi("listEvents", {
        ...monthRange,
        ...calendarPayload,
        coachCode
      });
      const events = Array.isArray(listResult?.events) ? listResult.events : [];
      const stats = alignCoachLessonsWithGoogleEvents(coachCode, events);
      const matchedLessonIds = new Set(Array.isArray(stats.matchedLessonIds) ? stats.matchedLessonIds : []);
      const meta = getMonthMeta(monthStart);
      const monthPrefix = `${String(meta.year).padStart(4, "0")}-${String(meta.month).padStart(2, "0")}-`;
      state.lessons
        .filter((lesson) => (
          lesson.coachCode === coachCode &&
          lesson.calendarOccupied &&
          getDateKeyInTaipei(new Date(lesson.startAt)).startsWith(monthPrefix) &&
          !matchedLessonIds.has(lesson.id) &&
          isGeneratedLocalCalendarEventId(lesson.calendarEventId)
        ))
        .forEach((lesson) => {
          markLessonRemovedByCalendar(lesson, "readonly_google_auto_sync");
          removedPlaceholders += 1;
        });
      totalEvents += stats.totalEvents;
      totalMatched += stats.matchedEvents;
      totalCreated += stats.createdLessons;
      syncedAny = true;
    }
    if (syncedAny) {
      ensureLessonCalendarEventIds();
      ensureParticipantEmails();
      ensureStudentBillingProfiles();
      saveState();
      addLog(`[日曆同步] 唯讀月曆已同步 Google：抓到 ${totalEvents} 筆，匹配 ${totalMatched} 筆，新增 ${totalCreated} 筆，移除暫存課 ${removedPlaceholders} 筆。`);
      return true;
    }
    return false;
  }

  function markLessonRemovedByCalendar(lesson, reason) {
    if (!lesson) {
      return;
    }
    if (!lesson.beforeCalendarRemoved) {
      lesson.beforeCalendarRemoved = {
        attendanceStatus: lesson.attendanceStatus,
        calendarOccupied: lesson.calendarOccupied,
        charged: lesson.charged,
        calendarEventId: lesson.calendarEventId || ""
      };
    }
    lesson.calendarOccupied = false;
    lesson.attendanceStatus = "calendar-removed";
    lesson.charged = false;
    lesson.calendarEventId = "";
    lesson.calendarRemovedAt = new Date().toISOString();
    lesson.calendarRemovedReason = reason || "google_calendar_deleted";
    addLog(`[日曆同步] 已依 Google 日曆移除課程 ${lesson.id}（${formatDateTime(lesson.startAt)}）。`);
  }

  async function tryCreateCalendarEventForLesson(lesson, reason, strictMode) {
    if (isNormalLeaveCalendarCreateBlocked({
      reason,
      attendanceStatus: lesson?.attendanceStatus,
      type: lesson?.type
    })) {
      addLog(`[Google日曆] 已阻擋正常請假建立事件：${lesson?.id || "-"}`);
      return "";
    }
    try {
      const result = await callAppsScriptApi("createEvent", buildLessonEventPayload(lesson, { reason }));
      const createdEventId = String(result?.eventId || result?.calendarEventId || lesson.calendarEventId || newId("GCAL")).trim();
      lesson.calendarEventId = createdEventId;
      addLog(`[Google日曆] 已建立事件 ${createdEventId}${result.mock ? "（模擬）" : ""}`);
      saveState();
      return createdEventId;
    } catch (error) {
      const message = String(error?.message || "未知錯誤");
      const unsupported = /unsupported|找不到|停用|invalid action|not found/i.test(message);
      if (strictMode) {
        if (!unsupported) {
          throw new Error(`Google 日曆建立失敗：${message}`);
        }
        lesson.calendarEventId = lesson.calendarEventId || `GCAL_${lesson.id}`;
        enqueueCompensationTask("createEvent", buildLessonEventPayload(lesson, { reason }), message);
        addLog(`[Google日曆] 端點未就緒，已改記補償任務：${lesson.id}`);
        saveState();
        return lesson.calendarEventId;
      }
      enqueueCompensationTask("createEvent", buildLessonEventPayload(lesson, { reason }), message);
      return "";
    }
  }

  function safeFormatNoticeDateTime(value) {
    const raw = String(value || "").trim();
    if (!raw) {
      return "-";
    }
    try {
      const date = new Date(raw);
      if (Number.isNaN(date.getTime())) {
        return raw;
      }
      return formatDateTime(date.toISOString());
    } catch (error) {
      return raw;
    }
  }

  function getNoticeStudentText(payload = {}) {
    const code = String(payload.studentCode || "").trim();
    const fallbackName = String(payload.studentName || "").trim();
    const name = code ? getStudentDisplayName(code) : fallbackName;
    if (name && code && name !== code) {
      return `${name}（${code}）`;
    }
    return name || code || "-";
  }

  function getNoticeCoachText(payload = {}) {
    const code = String(payload.coachCode || "").trim();
    const fallbackName = String(payload.coachName || "").trim();
    const name = code ? getCoachDisplayName(code) : fallbackName;
    if (name && code && name !== code) {
      return `${name}（${code}）`;
    }
    return name || code || "-";
  }

  function buildNoticeEmailContent(template, payload = {}) {
    const studentText = getNoticeStudentText(payload);
    const coachText = getNoticeCoachText(payload);
    const lessonTimeText = safeFormatNoticeDateTime(payload.lessonStartAt || payload.startAt || payload.when);
    const requestCode = String(payload.requestCode || payload.requestId || "-").trim() || "-";
    const leaveCode = String(payload.leaveId || payload.lessonId || "-").trim() || "-";
    const rejectReason = String(payload.rejectReason || payload.reason || "無").trim() || "無";
    const milestone = Number(payload.milestone || payload.totalChargedCount || 0);

    switch (String(template || "").trim()) {
      case "billing_reminder": {
        return {
          subject: `[CoachFlow] 繳費提醒｜${studentText}`,
          body: [
            `${studentText} 您好：`,
            "",
            `系統顯示目前累積扣堂已達 ${milestone || "-"} 堂，請協助完成本期繳費。`,
            `學生：${studentText}`,
            `教練：${coachText}`,
            `累積扣堂：${payload.totalChargedCount ?? "-"}`,
            `起始堂數：${payload.baseChargedCount ?? "-"}`,
            `本次系統扣堂：${payload.systemChargedCount ?? "-"}`,
            `已繳到堂數：${payload.paidThroughCount ?? "-"}`,
            `本期已扣：${payload.currentCycleChargedCount ?? "-"} / ${CHARGE_REMINDER_STEP}`,
            `下次應繳門檻：第 ${payload.nextPaymentDueCount ?? "-"} 堂`,
            "",
            "此信件由 CoachFlow 請假系統自動發送。"
          ].join("\n")
        };
      }
      case "billing_summary": {
        return {
          subject: `[CoachFlow] 計費摘要｜${studentText}`,
          body: [
            `${studentText} 您好：`,
            "",
            "以下為目前 CoachFlow 計費摘要：",
            `學生：${studentText}`,
            `教練：${coachText}`,
            `累積扣堂：${payload.totalChargedCount ?? "-"}`,
            `起始堂數：${payload.baseChargedCount ?? "-"}`,
            `本次系統扣堂：${payload.systemChargedCount ?? "-"}`,
            `已繳到堂數：${payload.paidThroughCount ?? "-"}`,
            `本期已扣：${payload.currentCycleChargedCount ?? "-"} / ${CHARGE_REMINDER_STEP}`,
            `下次應繳門檻：第 ${payload.nextPaymentDueCount ?? "-"} 堂`,
            `繳費狀態：${payload.paymentStatusLabel || "-"}`,
            `繳費註記：${payload.paymentNote || "無"}`,
            "",
            "此信件由 CoachFlow 請假系統手動發送。"
          ].join("\n")
        };
      }
      case "billing_payment_confirmed": {
        return {
          subject: `[CoachFlow] 匯款確認｜${studentText}`,
          body: [
            "您好：",
            "",
            "教練已在 CoachFlow 確認本期匯款，資訊如下：",
            `學生：${studentText}`,
            `教練：${coachText}`,
            `確認時間：${safeFormatNoticeDateTime(payload.confirmedAt)}`,
            `確認人：${payload.confirmedBy || "-"}`,
            `累積扣堂：${payload.totalChargedCount ?? "-"}`,
            `起始堂數：${payload.baseChargedCount ?? "-"}`,
            `本次系統扣堂：${payload.systemChargedCount ?? "-"}`,
            `已結算到第 ${payload.paidThroughCount ?? "-"} 堂`,
            `繳費狀態：${payload.paymentStatusLabel || "-"}`,
            `繳費註記：${payload.paymentNote || "無"}`,
            "",
            "此信件由 CoachFlow 請假系統自動發送。"
          ].join("\n")
        };
      }
      case "leave_submitted": {
        return {
          subject: `[CoachFlow] 已收到請假申請｜${studentText}`,
          body: [
            "您好：",
            "",
            "系統已收到請假申請，資訊如下：",
            `學生：${studentText}`,
            `教練：${coachText}`,
            `原上課時間：${lessonTimeText}`,
            `請假單號：${leaveCode}`,
            "",
            "如需補課，請至系統送出補課申請。",
            "此信件由 CoachFlow 請假系統自動發送。"
          ].join("\n")
        };
      }
      case "leave_cancelled_by_student": {
        return {
          subject: `[CoachFlow] 請假已取消｜${studentText}`,
          body: [
            "您好：",
            "",
            "學生已取消先前請假，原課程恢復排課。",
            `學生：${studentText}`,
            `教練：${coachText}`,
            `課程時間：${lessonTimeText}`,
            `課程編號：${leaveCode}`,
            "",
            "此信件由 CoachFlow 請假系統自動發送。"
          ].join("\n")
        };
      }
      case "coach_leave_blocked_lesson": {
        return {
          subject: `[CoachFlow] 教練請假異動｜${studentText}`,
          body: [
            "您好：",
            "",
            "教練請假造成課程異動，該堂課已取消排課。",
            `學生：${studentText}`,
            `教練：${coachText}`,
            `原課程時間：${lessonTimeText}`,
            `原因：${rejectReason}`,
            "",
            "請至系統查看後續補課安排。",
            "此信件由 CoachFlow 請假系統自動發送。"
          ].join("\n")
        };
      }
      case "makeup_auto_rejected_by_coach_leave": {
        return {
          subject: `[CoachFlow] 補課申請已自動退回｜${requestCode}`,
          body: [
            "您好：",
            "",
            "補課申請因教練請假時段重疊，系統已自動退回。",
            `申請代碼：${requestCode}`,
            `學生：${studentText}`,
            `教練：${coachText}`,
            `補課時段：${lessonTimeText}`,
            `退回原因：${rejectReason}`,
            "",
            "請重新選擇可用時段再送出申請。",
            "此信件由 CoachFlow 請假系統自動發送。"
          ].join("\n")
        };
      }
      case "coach_leave_restored_lesson": {
        return {
          subject: `[CoachFlow] 課程已恢復｜${studentText}`,
          body: [
            "您好：",
            "",
            "教練請假已取消，原課程已恢復。",
            `學生：${studentText}`,
            `教練：${coachText}`,
            `課程時間：${lessonTimeText}`,
            `課程編號：${leaveCode}`,
            "",
            "此信件由 CoachFlow 請假系統自動發送。"
          ].join("\n")
        };
      }
      case "makeup_pending": {
        return {
          subject: `[CoachFlow] 補課申請待審｜${requestCode}`,
          body: [
            "您好：",
            "",
            "有一筆新的補課申請待審核。",
            `申請代碼：${requestCode}`,
            `學生：${studentText}`,
            `教練：${coachText}`,
            `申請時段：${lessonTimeText}`,
            "",
            "請至教練端審核清單確認。",
            "此信件由 CoachFlow 請假系統自動發送。"
          ].join("\n")
        };
      }
      case "makeup_cancelled": {
        return {
          subject: `[CoachFlow] 補課申請已取消｜${requestCode}`,
          body: [
            "您好：",
            "",
            "學生已取消補課申請。",
            `申請代碼：${requestCode}`,
            `學生：${studentText}`,
            `教練：${coachText}`,
            `原申請時段：${lessonTimeText}`,
            "",
            "此信件由 CoachFlow 請假系統自動發送。"
          ].join("\n")
        };
      }
      case "leave_revoked_by_coach": {
        return {
          subject: `[CoachFlow] 教練已取消請假｜${studentText}`,
          body: [
            "您好：",
            "",
            "教練已取消此堂請假，課程恢復為可上課狀態。",
            `學生：${studentText}`,
            `教練：${coachText}`,
            `課程時間：${lessonTimeText}`,
            `課程編號：${leaveCode}`,
            "",
            "此信件由 CoachFlow 請假系統自動發送。"
          ].join("\n")
        };
      }
      case "makeup_approved": {
        return {
          subject: `[CoachFlow] 補課申請已核准｜${requestCode}`,
          body: [
            "您好：",
            "",
            "補課申請已核准，請依以下時間上課。",
            `申請代碼：${requestCode}`,
            `學生：${studentText}`,
            `教練：${coachText}`,
            `補課時間：${lessonTimeText}`,
            `補課課程編號：${payload.lessonId || "-"}`,
            "",
            "此信件由 CoachFlow 請假系統自動發送。"
          ].join("\n")
        };
      }
      case "makeup_rejected": {
        return {
          subject: `[CoachFlow] 補課申請未核准｜${requestCode}`,
          body: [
            "您好：",
            "",
            "補課申請未核准，請改選其他時段再申請。",
            `申請代碼：${requestCode}`,
            `學生：${studentText}`,
            `教練：${coachText}`,
            `申請時段：${lessonTimeText}`,
            `退回原因：${rejectReason}`,
            "",
            "此信件由 CoachFlow 請假系統自動發送。"
          ].join("\n")
        };
      }
      default: {
        return {
          subject: `[CoachFlow] 系統通知｜${String(template || "notification")}`,
          body: [
            "您好：",
            "",
            "這是 CoachFlow 請假系統通知。",
            `通知類型：${String(template || "-")}`,
            `學生：${studentText}`,
            `教練：${coachText}`,
            `時間：${lessonTimeText}`,
            "",
            "此信件由 CoachFlow 請假系統自動發送。"
          ].join("\n")
        };
      }
    }
  }

  async function trySendEmailNotice(template, payload, label) {
    const finalPayload = { ...(payload || {}) };
    const emailContent = buildNoticeEmailContent(template, finalPayload);
    if (!String(finalPayload.subject || "").trim()) {
      finalPayload.subject = emailContent.subject;
    }
    if (!String(finalPayload.body || "").trim()) {
      finalPayload.body = emailContent.body;
    }
    if (!String(finalPayload.fromName || "").trim()) {
      finalPayload.fromName = "CoachFlow 請假系統";
    }
    const recipients = resolveNoticeRecipients(finalPayload);
    if (recipients.length) {
      finalPayload.to = recipients.join(",");
    }
    if (!String(finalPayload.to || "").trim()) {
      addLog(`[Email] ${label} 已略過：缺少收件人（請設定 defaultNotifyEmail 或學生/教練 email）。`);
      return false;
    }
    try {
      const result = await callAppsScriptApi("sendEmail", { template, ...finalPayload });
      addLog(`[Email] ${label} 已送出${result.mock ? "（模擬）" : ""}`);
      saveState();
      return true;
    } catch (error) {
      const message = String(error?.message || "未知錯誤");
      enqueueCompensationTask("sendEmail", { template, ...finalPayload }, message);
      addLog(`[Email] ${label} 發送失敗：${message}`);
      return false;
    }
  }

  function ensureLessonCalendarEventIds() {
    let changed = false;
    state.lessons.forEach((lesson) => {
      if (!lesson.calendarEventId) {
        lesson.calendarEventId = `GCAL_${lesson.id}`;
        changed = true;
      }
    });
    if (changed) {
      saveState();
    }
  }

  function toNonNegativeInt(value, fallback = 0) {
    const n = Number(value);
    if (!Number.isFinite(n)) {
      return fallback;
    }
    return Math.max(0, Math.floor(n));
  }

  function normalizePaymentStatus(value) {
    const text = String(value || "").trim();
    if (text === "paid" || text === "unpaid") {
      return text;
    }
    return "unknown";
  }

  function normalizeChargeReminderLogs(logs) {
    if (!Array.isArray(logs)) {
      return [];
    }
    return logs
      .map((item) => ({
        id: String(item?.id || newId("BILLMAIL")),
        milestone: toNonNegativeInt(item?.milestone, 0),
        status: item?.status === "success" ? "success" : "failed",
        sentAt: item?.sentAt ? String(item.sentAt) : "",
        to: String(item?.to || ""),
        note: String(item?.note || ""),
        triggerSource: String(item?.triggerSource || "")
      }))
      .filter((item) => item.milestone > 0)
      .sort((a, b) => new Date(b.sentAt || 0) - new Date(a.sentAt || 0))
      .slice(0, MAX_CHARGE_REMINDER_LOGS);
  }

  function ensureStudentBillingProfiles() {
    let changed = false;
    state.students = (state.students || []).map((student) => {
      if (!student || typeof student !== "object") {
        return student;
      }
      const existingStatus = normalizePaymentStatus(student.paymentStatus);
      const existingStartCount = toNonNegativeInt(student.chargeStartCount, 0);
      const existingChargedCount = (state.lessons || [])
        .filter((lesson) => (
          lesson.studentCode === student.code &&
          lesson.attendanceStatus !== "calendar-removed" &&
          lesson.charged
        ))
        .length;
      const fallbackPaidThrough = existingStatus === "paid"
        ? existingStartCount + existingChargedCount
        : 0;
      const normalized = {
        ...student,
        chargeStartCount: existingStartCount,
        paidThroughCount: student.paidThroughCount === undefined
          ? fallbackPaidThrough
          : toNonNegativeInt(student.paidThroughCount, 0),
        paymentStatus: existingStatus,
        paymentNote: String(student.paymentNote || ""),
        paymentConfirmedAt: String(student.paymentConfirmedAt || ""),
        paymentConfirmedBy: String(student.paymentConfirmedBy || ""),
        emailUpdatedAt: String(student.emailUpdatedAt || ""),
        emailUpdatedBy: String(student.emailUpdatedBy || ""),
        chargeReminderLogs: normalizeChargeReminderLogs(student.chargeReminderLogs)
      };
      const before = JSON.stringify(student);
      const after = JSON.stringify(normalized);
      if (before !== after) {
        changed = true;
      }
      return normalized;
    });
    if (changed) {
      saveState();
    }
  }

  function ensureParticipantEmails() {
    let changed = false;
    const fallback = String(window.APP_CONFIG?.defaultNotifyEmail || "").trim();
    state.students = (state.students || []).map((student) => {
      if (student && student.email === undefined) {
        changed = true;
        return { ...student, email: fallback };
      }
      return student;
    });
    state.coaches = (state.coaches || []).map((coach) => {
      if (coach && coach.email === undefined) {
        changed = true;
        return { ...coach, email: fallback };
      }
      return coach;
    });
    if (changed) {
      saveState();
    }
  }

  function getWeekdayFromDateKey(dateKey) {
    const day = makeTaipeiDateTime(dateKey, "12:00").getUTCDay();
    return ["sun", "mon", "tue", "wed", "thu", "fri", "sat"][day];
  }

  function getNextDateKeyByWeekday(startDateKey, targetWeekday, includeToday) {
    let key = includeToday ? startDateKey : addDays(startDateKey, 1);
    for (let i = 0; i < 14; i += 1) {
      if (getWeekdayFromDateKey(key) === targetWeekday) {
        return key;
      }
      key = addDays(key, 1);
    }
    return key;
  }

  function getCycleWindowByLesson(lesson) {
    let cycleStart = getDateKeyInTaipei(new Date(lesson.startAt));
    while (getWeekdayFromDateKey(cycleStart) !== "sat") {
      cycleStart = addDays(cycleStart, -1);
    }
    return {
      startKey: cycleStart,
      endKey: addDays(cycleStart, 2)
    };
  }

  function getCutoffForLesson(lesson) {
    const lessonDate = getDateKeyInTaipei(new Date(lesson.startAt));
    const cutoffDay = addDays(lessonDate, -1);
    return makeTaipeiDateTime(cutoffDay, "23:59");
  }

  function isLeaveOpen(lesson) {
    return new Date() <= getCutoffForLesson(lesson);
  }

  function getStudentByCode(studentCode) {
    return state.students.find((student) => student.code === studentCode);
  }

  function getPaymentStatusLabel(status) {
    return PAYMENT_STATUS_LABELS[normalizePaymentStatus(status)] || PAYMENT_STATUS_LABELS.unknown;
  }

  function getStudentChargeStats(studentCode) {
    const student = getStudentByCode(studentCode);
    const lessons = state.lessons
      .filter((lesson) => lesson.studentCode === studentCode && lesson.attendanceStatus !== "calendar-removed")
      .sort((a, b) => new Date(b.startAt) - new Date(a.startAt));
    const chargedLessons = lessons.filter((lesson) => lesson.charged);
    const startCount = toNonNegativeInt(student?.chargeStartCount, 0);
    return {
      student,
      lessons,
      chargedLessons,
      startCount,
      totalChargedCount: startCount + chargedLessons.length,
      noShowCount: lessons.filter((lesson) => lesson.attendanceStatus === "no-show").length,
      tempLeaveCount: lessons.filter((lesson) => lesson.attendanceStatus === "temporary-leave").length,
      majorCount: lessons.filter((lesson) => lesson.attendanceStatus === "major-case").length,
      normalLeaveCount: lessons.filter((lesson) => lesson.attendanceStatus === "leave-normal").length
    };
  }

  function getStudentBillingCycle(stats) {
    const student = stats?.student;
    const totalChargedCount = toNonNegativeInt(stats?.totalChargedCount, 0);
    const paidThroughCount = Math.min(
      toNonNegativeInt(student?.paidThroughCount, 0),
      totalChargedCount
    );
    const currentCycleChargedCount = Math.max(0, totalChargedCount - paidThroughCount);
    const nextPaymentDueCount = paidThroughCount + CHARGE_REMINDER_STEP;
    const remainingToNextPayment = Math.max(0, nextPaymentDueCount - totalChargedCount);
    const isPaymentDue = currentCycleChargedCount >= CHARGE_REMINDER_STEP;
    const storedStatus = normalizePaymentStatus(student?.paymentStatus);
    const effectivePaymentStatus = isPaymentDue ? "unpaid" : storedStatus;
    return {
      paidThroughCount,
      currentCycleChargedCount,
      nextPaymentDueCount,
      remainingToNextPayment,
      isPaymentDue,
      effectivePaymentStatus
    };
  }

  function getChargeReminderStatusPill(status) {
    if (status === "success") {
      return "<span class=\"status approved\">成功</span>";
    }
    return "<span class=\"status rejected\">失敗</span>";
  }

  async function maybeSendChargeReminder(studentCode, triggerSource) {
    const stats = getStudentChargeStats(studentCode);
    const student = stats.student;
    if (!student) {
      return false;
    }
    const billingCycle = getStudentBillingCycle(stats);
    const milestone = billingCycle.nextPaymentDueCount;
    if (!billingCycle.isPaymentDue || milestone <= 0) {
      return false;
    }
    const reminderLogs = Array.isArray(student.chargeReminderLogs) ? student.chargeReminderLogs : [];
    if (reminderLogs.some((item) => Number(item?.milestone) === milestone)) {
      return false;
    }
    const reminderKey = `${student.code}:${milestone}`;
    if (sendingChargeReminderKeys.has(reminderKey)) {
      return false;
    }
    sendingChargeReminderKeys.add(reminderKey);

    try {
      const payload = {
        studentCode: student.code,
        coachCode: student.coachCode,
        milestone,
        totalChargedCount: stats.totalChargedCount,
        baseChargedCount: stats.startCount,
        systemChargedCount: stats.chargedLessons.length,
        paidThroughCount: billingCycle.paidThroughCount,
        currentCycleChargedCount: billingCycle.currentCycleChargedCount,
        nextPaymentDueCount: billingCycle.nextPaymentDueCount,
        remainingToNextPayment: billingCycle.remainingToNextPayment,
        triggerSource: triggerSource || "system"
      };
      const recipientsText = resolveNoticeRecipients(payload).join(", ");
      const sent = await trySendEmailNotice(
        "billing_reminder",
        payload,
        `${student.code} 第 ${milestone} 堂扣堂提醒`
      );
      const logEntry = {
        id: newId("BILLMAIL"),
        milestone,
        status: sent ? "success" : "failed",
        sentAt: new Date().toISOString(),
        to: recipientsText,
        note: sent
          ? `第 ${milestone} 堂提醒已寄送`
          : (recipientsText ? "寄送失敗，已建立補償任務" : "缺少收件人，提醒未寄送"),
        triggerSource: triggerSource || "system"
      };
      student.chargeReminderLogs = [logEntry, ...reminderLogs].slice(0, MAX_CHARGE_REMINDER_LOGS);
      addLog(`[計費提醒] ${student.code} 第 ${milestone} 堂提醒${sent ? "已送出" : "未送出"}。`);
      saveState();
      if (selectedChargeStudentCode === student.code) {
        renderChargePanel();
      }
      return sent;
    } finally {
      sendingChargeReminderKeys.delete(reminderKey);
    }
  }

  async function sendSelectedStudentChargeEmail() {
    if (!requireCoachWriteAccess()) {
      return false;
    }
    const stats = getStudentChargeStats(selectedChargeStudentCode);
    const student = stats.student;
    if (!student) {
      alert("請先選擇學生。");
      return false;
    }
    const reminderLogs = Array.isArray(student.chargeReminderLogs) ? student.chargeReminderLogs : [];
    const billingCycle = getStudentBillingCycle(stats);
    const sendKey = `${student.code}:manual`;
    if (sendingChargeReminderKeys.has(sendKey)) {
      return false;
    }
    sendingChargeReminderKeys.add(sendKey);
    if (el.chargeEmailSendBtn) {
      el.chargeEmailSendBtn.disabled = true;
      el.chargeEmailSendBtn.textContent = "寄送中...";
    }

    try {
      const payload = {
        studentCode: student.code,
        coachCode: student.coachCode,
        milestone: stats.totalChargedCount,
        totalChargedCount: stats.totalChargedCount,
        baseChargedCount: stats.startCount,
        systemChargedCount: stats.chargedLessons.length,
        paidThroughCount: billingCycle.paidThroughCount,
        currentCycleChargedCount: billingCycle.currentCycleChargedCount,
        nextPaymentDueCount: billingCycle.nextPaymentDueCount,
        remainingToNextPayment: billingCycle.remainingToNextPayment,
        paymentStatusLabel: getPaymentStatusLabel(billingCycle.effectivePaymentStatus),
        paymentNote: String(student.paymentNote || ""),
        triggerSource: "manual_charge_panel"
      };
      const recipientsText = resolveNoticeRecipients(payload).join(", ");
      const sent = await trySendEmailNotice(
        "billing_summary",
        payload,
        `${student.code} 計費摘要`
      );
      const logEntry = {
        id: newId("BILLMAIL"),
        milestone: stats.totalChargedCount,
        label: `手動寄送（累計 ${stats.totalChargedCount} 堂）`,
        status: sent ? "success" : "failed",
        sentAt: new Date().toISOString(),
        to: recipientsText,
        note: sent
          ? "計費摘要已手動寄送"
          : (recipientsText ? "寄送失敗，已建立補償任務" : "缺少收件人，未寄送"),
        triggerSource: "manual_charge_panel"
      };
      student.chargeReminderLogs = [logEntry, ...reminderLogs].slice(0, MAX_CHARGE_REMINDER_LOGS);
      addLog(`[計費] ${student.code} 計費摘要${sent ? "已寄送" : "未寄送"}。`);
      saveState();
      renderChargePanel();
      return sent;
    } finally {
      sendingChargeReminderKeys.delete(sendKey);
      if (el.chargeEmailSendBtn) {
        el.chargeEmailSendBtn.disabled = false;
        el.chargeEmailSendBtn.textContent = "寄送計費 Email";
      }
    }
  }

  function saveSelectedStudentNotifyEmail() {
    if (!requireCoachWriteAccess()) {
      return;
    }
    const student = getStudentByCode(selectedChargeStudentCode);
    if (!student) {
      alert("請先選擇學生。");
      return;
    }
    const rawInput = String(el.chargeStudentEmailInput?.value || "").trim();
    const normalizedEmail = normalizeEmailList(rawInput).join(", ");
    student.email = normalizedEmail;
    student.emailUpdatedAt = new Date().toISOString();
    student.emailUpdatedBy = activeCoachCode || "SYSTEM";
    addLog(
      `[通知] ${student.code} 學生通知 Email 已更新：${normalizedEmail || "清空（將改用預設通知信箱）"}。`
    );
    saveState();
    renderChargePanel();
  }

  function saveChargeStartCount() {
    if (!requireCoachWriteAccess()) {
      return;
    }
    const student = getStudentByCode(selectedChargeStudentCode);
    if (!student) {
      alert("請先選擇學生。");
      return;
    }
    const nextCount = toNonNegativeInt(el.chargeBaseCountInput?.value, 0);
    student.chargeStartCount = nextCount;
    addLog(`[計費] ${student.code} 起始堂數調整為 ${nextCount}。`);
    saveState();
    renderChargePanel();
    maybeSendChargeReminder(student.code, "set_base_count").catch((error) => {
      console.error("billing reminder failed:", error);
    });
  }

  async function sendStudentPaymentConfirmation(student, stats, billingCycle) {
    const confirmationKey = `${student.code}:payment-confirmed`;
    if (sendingChargeReminderKeys.has(confirmationKey)) {
      return false;
    }
    sendingChargeReminderKeys.add(confirmationKey);
    try {
      const payload = {
        studentCode: student.code,
        coachCode: student.coachCode,
        studentEmail: getStudentNoticeEmail(student.code),
        coachEmail: getCoachNoticeEmail(student.coachCode),
        confirmedAt: student.paymentConfirmedAt,
        confirmedBy: student.paymentConfirmedBy,
        totalChargedCount: stats.totalChargedCount,
        baseChargedCount: stats.startCount,
        systemChargedCount: stats.chargedLessons.length,
        paidThroughCount: toNonNegativeInt(student.paidThroughCount, 0),
        currentCycleChargedCount: billingCycle.currentCycleChargedCount,
        nextPaymentDueCount: billingCycle.nextPaymentDueCount,
        remainingToNextPayment: billingCycle.remainingToNextPayment,
        paymentStatusLabel: getPaymentStatusLabel(student.paymentStatus),
        paymentNote: String(student.paymentNote || ""),
        triggerSource: "payment_confirmed_by_coach"
      };
      const recipientsText = resolveNoticeRecipients(payload).join(", ");
      const sent = await trySendEmailNotice(
        "billing_payment_confirmed",
        payload,
        `${student.code} 匯款確認`
      );
      const reminderLogs = Array.isArray(student.chargeReminderLogs) ? student.chargeReminderLogs : [];
      const logEntry = {
        id: newId("BILLMAIL"),
        milestone: stats.totalChargedCount,
        label: `匯款確認（已結算至第 ${toNonNegativeInt(student.paidThroughCount, 0)} 堂）`,
        status: sent ? "success" : "failed",
        sentAt: new Date().toISOString(),
        to: recipientsText,
        note: sent
          ? "匯款確認信已寄送給教練與學生"
          : (recipientsText ? "匯款確認信寄送失敗，已建立補償任務" : "缺少收件人，匯款確認信未寄送"),
        triggerSource: "payment_confirmed_by_coach"
      };
      student.chargeReminderLogs = [logEntry, ...reminderLogs].slice(0, MAX_CHARGE_REMINDER_LOGS);
      addLog(`[計費] ${student.code} 匯款確認信${sent ? "已寄送" : "未寄送"}。`);
      saveState();
      return sent;
    } finally {
      sendingChargeReminderKeys.delete(confirmationKey);
    }
  }

  async function saveStudentPaymentStatus() {
    if (!requireCoachWriteAccess()) {
      return;
    }
    const student = getStudentByCode(selectedChargeStudentCode);
    if (!student) {
      alert("請先選擇學生。");
      return;
    }
    const nextStatus = normalizePaymentStatus(el.chargePaymentStatusSelect?.value);
    const note = String(el.chargePaymentNoteInput?.value || "").trim();
    const stats = getStudentChargeStats(student.code);
    const settledThrough = stats.totalChargedCount;
    student.paymentStatus = nextStatus;
    student.paymentNote = note;
    student.paymentConfirmedAt = new Date().toISOString();
    student.paymentConfirmedBy = activeCoachCode || "SYSTEM";
    if (nextStatus === "paid") {
      student.paidThroughCount = settledThrough;
    }
    addLog(
      `[計費] ${student.code} 繳費狀態更新為 ${getPaymentStatusLabel(nextStatus)}（${student.paymentConfirmedBy}，已結算至第 ${toNonNegativeInt(student.paidThroughCount, 0)} 堂）。`
    );
    saveState();
    renderChargePanel();
    if (nextStatus === "paid") {
      if (el.chargePaymentSaveBtn) {
        el.chargePaymentSaveBtn.disabled = true;
        el.chargePaymentSaveBtn.textContent = "寄送確認信...";
      }
      try {
        const refreshedStats = getStudentChargeStats(student.code);
        const refreshedBillingCycle = getStudentBillingCycle(refreshedStats);
        await sendStudentPaymentConfirmation(student, refreshedStats, refreshedBillingCycle);
      } finally {
        if (el.chargePaymentSaveBtn) {
          el.chargePaymentSaveBtn.disabled = false;
          el.chargePaymentSaveBtn.textContent = "儲存繳費註記";
        }
        renderChargePanel();
      }
    }
  }

  function normalizeEmailList(value) {
    const list = String(value || "")
      .split(/[;,]/)
      .map((item) => item.trim())
      .filter((item) => item && item.includes("@"));
    return Array.from(new Set(list));
  }

  function resolveNoticeRecipients(payload = {}) {
    const recipients = new Set();
    normalizeEmailList(payload.to).forEach((item) => recipients.add(item));
    normalizeEmailList(payload.studentEmail).forEach((item) => recipients.add(item));
    normalizeEmailList(payload.coachEmail).forEach((item) => recipients.add(item));

    if (payload.studentCode) {
      const student = getStudentByCode(payload.studentCode);
      normalizeEmailList(student?.email).forEach((item) => recipients.add(item));
    }
    if (payload.coachCode) {
      const coach = getCoachByCode(payload.coachCode);
      normalizeEmailList(coach?.email).forEach((item) => recipients.add(item));
    }
    normalizeEmailList(window.APP_CONFIG?.defaultNotifyEmail).forEach((item) => recipients.add(item));
    return Array.from(recipients);
  }

  function getStudentNoticeEmail(studentCode) {
    const student = getStudentByCode(studentCode);
    return normalizeEmailList(student?.email).join(", ");
  }

  function getCoachNoticeEmail(coachCode) {
    const coach = getCoachByCode(coachCode);
    return normalizeEmailList(coach?.email).join(", ");
  }

  function getLessonById(lessonId) {
    return state.lessons.find((lesson) => lesson.id === lessonId);
  }

  function getLeaveByLesson(lessonId) {
    return state.leaveRequests.find((leave) => leave.lessonId === lessonId && !leave.revokedAt);
  }

  function getMakeupByLeave(leaveId) {
    return state.makeupRequests.find((request) => request.leaveId === leaveId && (request.status === "pending" || request.status === "approved"));
  }

  function getApprovedMakeupByLessonId(lessonId) {
    const leave = getLeaveByLesson(lessonId);
    if (!leave) {
      return null;
    }
    return state.makeupRequests.find((request) => request.leaveId === leave.id && request.status === "approved") || null;
  }

  function expirePendingRequests() {
    const now = new Date();
    let changed = false;
    state.makeupRequests.forEach((request) => {
      if (request.status !== "pending") {
        return;
      }
      const pendingAt = new Date(request.pendingAt);
      const hours = (now - pendingAt) / 36e5;
      if (hours >= PENDING_EXPIRE_HOURS) {
        request.status = "expired";
        request.resolvedAt = now.toISOString();
        addLog(`補課申請 ${request.id} 超過 ${PENDING_EXPIRE_HOURS} 小時未處理，已轉為逾時。`);
        changed = true;
      }
    });
    if (changed) {
      saveState();
    }
  }

  async function applyNormalLeave(lessonId) {
    const lesson = getLessonById(lessonId);
    if (!lesson || lesson.studentCode !== activeStudentCode) {
      alert("找不到課程，或你沒有權限操作。");
      return;
    }
    if (!isLeaveOpen(lesson)) {
      alert("已超過請假截止時間（前一天 23:59），不可請假。");
      return;
    }
    if (getLeaveByLesson(lessonId)) {
      alert("這堂課已經送出請假。");
      return;
    }

    try {
      await tryDeleteCalendarEventForLesson(lesson, "student_normal_leave", true);
    } catch (error) {
      alert(String(error?.message || "Google 日曆操作失敗，請稍後再試。"));
      return;
    }

    const leaveId = newId("LEAVE");
    const leaveRecord = {
      id: leaveId,
      lessonId,
      studentCode: lesson.studentCode,
      coachCode: lesson.coachCode,
      type: "normal",
      submittedAt: new Date().toISOString(),
      makeupEligible: true
    };
    state.leaveRequests.push(leaveRecord);
    lesson.calendarOccupied = false;
    lesson.attendanceStatus = "leave-normal";
    addLog(`學生 ${lesson.studentCode} 送出請假（課程 ${lesson.id}）。`);
    saveState();
    renderAll();
    try {
      await pushCloudLeaveRecord(leaveRecord, lesson);
      addLog(`[雲端請假] 已上傳請假紀錄 ${leaveRecord.id}。`);
      saveState();
    } catch (error) {
      const message = String(error?.message || "未知錯誤");
      addLog(`[雲端請假] 上傳請假紀錄失敗：${message}`);
      saveState();
    }
    const studentEmail = getStudentNoticeEmail(lesson.studentCode);
    const coachEmail = getCoachNoticeEmail(lesson.coachCode);
    if (!studentEmail) {
      addLog(`[Email] ${lesson.studentCode} 未設定學生 Email，請於計費面板補上學生通知 Email。`);
      saveState();
    }
    await trySendEmailNotice(
      "leave_submitted",
      {
        studentCode: lesson.studentCode,
        coachCode: lesson.coachCode,
        studentEmail,
        coachEmail,
        lessonId: lesson.id,
        lessonStartAt: lesson.startAt,
        leaveId
      },
      `學生 ${lesson.studentCode} 請假通知`
    );
  }

  async function cancelNormalLeaveByStudent(lessonId) {
    const lesson = getLessonById(lessonId);
    if (!lesson || lesson.studentCode !== activeStudentCode) {
      alert("找不到可取消請假的課程，或你沒有權限。");
      return;
    }
    if (lesson.attendanceStatus !== "leave-normal") {
      alert("這堂課目前不是「正常請假」狀態。");
      return;
    }
    if (!isLeaveOpen(lesson)) {
      alert("已超過取消截止時間（前一天 23:59），不可取消請假。");
      return;
    }

    const leave = getLeaveByLesson(lessonId);
    if (!leave) {
      alert("找不到對應的請假紀錄。");
      return;
    }

    const approvedMakeup = state.makeupRequests.find(
      (request) => request.leaveId === leave.id && request.status === "approved"
    );
    if (approvedMakeup) {
      alert("此請假已有核准補課，請聯繫教練處理。");
      return;
    }

    const nowIso = new Date().toISOString();
    state.makeupRequests.forEach((request) => {
      if (request.leaveId === leave.id && request.status === "pending") {
        request.status = "cancelled";
        request.resolvedAt = nowIso;
        request.rejectReason = "學生取消請假，待審補課已自動取消。";
      }
    });

    leave.revokedAt = nowIso;
    leave.revokedBy = activeStudentCode;
    leave.makeupEligible = false;

    lesson.attendanceStatus = "scheduled";
    lesson.calendarOccupied = true;
    lesson.charged = false;
    await tryCreateCalendarEventForLesson(lesson, "student_cancel_leave", false);

    addLog(`學生 ${activeStudentCode} 已取消課程 ${lesson.id} 的請假。`);
    saveState();
    renderAll();
    try {
      await pushCloudLeaveRecord(leave, lesson);
    } catch (error) {
      addLog(`[雲端請假] 上傳取消請假失敗：${String(error?.message || "未知錯誤")}`);
      saveState();
    }
    await trySendEmailNotice(
      "leave_cancelled_by_student",
      {
        studentCode: lesson.studentCode,
        coachCode: lesson.coachCode,
        lessonId: lesson.id,
        lessonStartAt: lesson.startAt
      },
      `學生 ${lesson.studentCode} 取消請假通知`
    );
  }

  function getCycleSlots(leave) {
    const lesson = getLessonById(leave.lessonId);
    if (!lesson) {
      return [];
    }
    const cycle = getCycleWindowByLesson(lesson);
    const slotKeys = [cycle.startKey, addDays(cycle.startKey, 1), addDays(cycle.startKey, 2)];
    const slots = [];

    slotKeys.forEach((dateKey) => {
      const day = getWeekdayFromDateKey(dateKey);
      const times = day === "sat" || day === "sun"
        ? ["09:00", "10:00", "11:00", "12:00", "13:00", "14:00"]
        : day === "mon"
          ? ["18:30", "19:30"]
          : [];
      times.forEach((timeText) => {
        const startAt = makeTaipeiDateTime(dateKey, timeText).toISOString();
        slots.push({
          slotId: `${dateKey}_${timeText}`,
          coachCode: leave.coachCode,
          startAt
        });
      });
    });

    const now = new Date();
    return slots.filter((slot) => {
      if (new Date(slot.startAt) <= now) {
        return false;
      }
      return !isSlotOccupied(slot.startAt, leave.coachCode);
    });
  }

  function isSlotOccupied(slotStartAt, coachCode, excludeRequestId, excludeLessonId) {
    const byLesson = state.lessons.some(
      (lesson) =>
        lesson.coachCode === coachCode &&
        lesson.calendarOccupied &&
        lesson.startAt === slotStartAt &&
        lesson.id !== excludeLessonId
    );
    if (byLesson) {
      return true;
    }
    const byCoachBlock = (state.coachBlocks || []).some(
      (block) =>
        block.coachCode === coachCode &&
        isTimeWithinRange(slotStartAt, block.startAt, block.endAt || block.startAt)
    );
    if (byCoachBlock) {
      return true;
    }
    return state.makeupRequests.some((request) => {
      if (excludeRequestId && request.id === excludeRequestId) {
        return false;
      }
      return request.coachCode === coachCode && request.startAt === slotStartAt && request.status === "pending";
    });
  }

  async function addCoachLeaveBlock() {
    if (!requireCoachWriteAccess()) {
      return;
    }
    if (!activeCoachCode) {
      alert("請先登入教練帳號。");
      return;
    }
    if (!el.coachLeaveDate || !el.coachLeaveTime || !el.coachLeaveEndTime) {
      return;
    }
    const dateKey = String(el.coachLeaveDate.value || "").trim();
    const startTimeText = String(el.coachLeaveTime.value || "").trim();
    const endTimeText = String(el.coachLeaveEndTime.value || "").trim();
    const reason = String(el.coachLeaveReason?.value || "").trim() || "教練請假";
    if (!dateKey || !startTimeText || !endTimeText) {
      alert("請輸入請假日期、開始時間與結束時間。");
      return;
    }
    const startAtDate = makeTaipeiDateTime(dateKey, startTimeText);
    const endAtDate = makeTaipeiDateTime(dateKey, endTimeText);
    if (endAtDate <= startAtDate) {
      alert("結束時間必須晚於開始時間。");
      return;
    }
    const startAt = startAtDate.toISOString();
    const endAt = endAtDate.toISOString();
    const duplicateBlock = (state.coachBlocks || []).some(
      (block) =>
        block.coachCode === activeCoachCode &&
        isRangeOverlap(startAt, endAt, block.startAt, block.endAt || block.startAt)
    );
    if (duplicateBlock) {
      alert("此時段與既有教練請假重疊。");
      return;
    }
    state.coachBlocks = state.coachBlocks || [];
    const block = {
      id: newId("BLOCK"),
      coachCode: activeCoachCode,
      startAt,
      endAt,
      reason
    };
    state.coachBlocks.push(block);

    const impactedLessons = state.lessons.filter(
      (lesson) =>
        lesson.coachCode === activeCoachCode &&
        isTimeWithinRange(lesson.startAt, startAt, endAt) &&
        lesson.attendanceStatus !== "coach-leave"
    );
    impactedLessons.forEach((lesson) => {
      if (lesson.cancelledByCoachLeaveBlockId !== block.id) {
        lesson.beforeCoachLeave = {
          attendanceStatus: lesson.attendanceStatus,
          calendarOccupied: lesson.calendarOccupied,
          charged: lesson.charged
        };
      }
      lesson.attendanceStatus = "coach-leave";
      lesson.calendarOccupied = false;
      lesson.charged = false;
      lesson.cancelledByCoachLeaveBlockId = block.id;
      lesson.coachLeaveReason = reason;
    });

    const impactedPendingRequests = state.makeupRequests.filter(
      (request) =>
        request.coachCode === activeCoachCode &&
        isTimeWithinRange(request.startAt, startAt, endAt) &&
        request.status === "pending"
    );
    impactedPendingRequests.forEach((request) => {
      request.status = "rejected";
      request.resolvedAt = new Date().toISOString();
      request.rejectReason = "教練請假時段重疊，無法安排此補課時段。";
    });

    for (const lesson of impactedLessons) {
      await tryDeleteCalendarEventForLesson(lesson, "coach_leave_block", false);
      await trySendEmailNotice(
        "coach_leave_blocked_lesson",
        {
          lessonId: lesson.id,
          studentCode: lesson.studentCode,
          coachCode: lesson.coachCode,
          lessonStartAt: lesson.startAt,
          reason
        },
        `教練請假異動通知 ${lesson.id}`
      );
    }

    for (const request of impactedPendingRequests) {
      await trySendEmailNotice(
        "makeup_auto_rejected_by_coach_leave",
        {
          requestId: request.id,
          requestCode: request.code,
          studentCode: request.studentCode,
          coachCode: request.coachCode,
          startAt: request.startAt,
          rejectReason: request.rejectReason
        },
        `補課自動退回通知 ${request.code || request.id}`
      );
    }

    addLog(`教練 ${activeCoachCode} 新增請假時段 ${formatDateTime(startAt)}-${getTimeText(endAt)}，影響課程 ${impactedLessons.length} 堂、退回待審 ${impactedPendingRequests.length} 筆。`);
    saveState();
    renderAll();
  }

  async function removeCoachLeaveBlock(blockId) {
    if (!requireCoachWriteAccess()) {
      return;
    }
    const block = (state.coachBlocks || []).find((item) => item.id === blockId);
    if (!block || block.coachCode !== activeCoachCode) {
      return;
    }
    state.coachBlocks = (state.coachBlocks || []).filter((item) => item.id !== blockId);

    const restorableLessons = state.lessons.filter(
      (lesson) => lesson.cancelledByCoachLeaveBlockId === blockId
    );
    restorableLessons.forEach((lesson) => {
      const previous = lesson.beforeCoachLeave || {
        attendanceStatus: "scheduled",
        calendarOccupied: true,
        charged: false
      };
      lesson.attendanceStatus = previous.attendanceStatus;
      lesson.calendarOccupied = previous.calendarOccupied;
      lesson.charged = previous.charged;
      delete lesson.cancelledByCoachLeaveBlockId;
      delete lesson.coachLeaveReason;
      delete lesson.beforeCoachLeave;
    });
    for (const lesson of restorableLessons) {
      await tryCreateCalendarEventForLesson(lesson, "coach_leave_block_removed", false);
      await trySendEmailNotice(
        "coach_leave_restored_lesson",
        {
          lessonId: lesson.id,
          studentCode: lesson.studentCode,
          coachCode: lesson.coachCode,
          lessonStartAt: lesson.startAt
        },
        `課程恢復通知 ${lesson.id}`
      );
    }
    addLog(`教練 ${activeCoachCode} 移除請假時段 ${formatDateTime(block.startAt)}-${getTimeText(block.endAt || block.startAt)}。`);
    saveState();
    renderAll();
  }

  async function submitMakeupRequest() {
    if (!el.makeupLeaveSelect || !el.makeupSlotSelect) {
      return;
    }
    const leaveId = el.makeupLeaveSelect.value;
    const startAt = el.makeupSlotSelect.value;
    if (!leaveId || !startAt) {
      alert("請先選擇請假單與補課時段。");
      return;
    }
    const leave = state.leaveRequests.find((item) => item.id === leaveId);
    if (!leave || !leave.makeupEligible) {
      alert("這張請假單目前不可補課。");
      return;
    }
    if (getMakeupByLeave(leave.id)) {
      alert("這張請假單已有待審或已核准的補課申請。");
      return;
    }
    if (isSlotOccupied(startAt, leave.coachCode)) {
      alert("該時段已被佔用，請改選其他時段。");
      return;
    }
    const pendingAt = new Date().toISOString();
    const request = {
      id: newId("MAKEUP"),
      code: buildMakeupCode(pendingAt),
      leaveId: leave.id,
      lessonId: leave.lessonId,
      studentCode: leave.studentCode,
      coachCode: leave.coachCode,
      status: "pending",
      pendingAt,
      resolvedAt: "",
      startAt,
      rejectReason: ""
    };
    state.makeupRequests.push(request);
    ensureMakeupCodes();
    addLog(`學生 ${leave.studentCode} 送出補課申請：${formatDateTime(startAt)}。`);
    saveState();
    renderAll();
    await trySendEmailNotice(
      "makeup_pending",
      {
        requestId: request.id,
        requestCode: request.code,
        studentCode: request.studentCode,
        coachCode: request.coachCode,
        startAt: request.startAt
      },
      `補課待審通知 ${request.code || request.id}`
    );
  }

  async function cancelPending(requestId) {
    const request = state.makeupRequests.find((item) => item.id === requestId);
    if (!request || request.studentCode !== activeStudentCode || request.status !== "pending") {
      alert("找不到待審申請，或你沒有權限。");
      return;
    }
    request.status = "cancelled";
    request.resolvedAt = new Date().toISOString();
    addLog(`學生 ${request.studentCode} 取消補課申請 ${request.id}。`);
    saveState();
    renderAll();
    await trySendEmailNotice(
      "makeup_cancelled",
      {
        requestId: request.id,
        requestCode: request.code,
        studentCode: request.studentCode,
        coachCode: request.coachCode,
        startAt: request.startAt
      },
      `補課申請取消通知 ${request.code || request.id}`
    );
  }

  async function revokeNormalLeave(lessonId) {
    if (!requireCoachWriteAccess()) {
      return;
    }
    const lesson = getLessonById(lessonId);
    if (!lesson || lesson.coachCode !== activeCoachCode) {
      alert("找不到可取消請假的課程。");
      return;
    }
    if (lesson.attendanceStatus !== "leave-normal") {
      alert("這堂課目前不是「正常請假」狀態。");
      return;
    }

    const leave = getLeaveByLesson(lessonId);
    if (!leave) {
      alert("找不到對應的請假紀錄。");
      return;
    }

    const approvedMakeup = state.makeupRequests.find(
      (request) => request.leaveId === leave.id && request.status === "approved"
    );
    if (approvedMakeup) {
      alert("此請假已有核准補課，不能直接取消請假。");
      return;
    }

    const nowIso = new Date().toISOString();
    state.makeupRequests.forEach((request) => {
      if (request.leaveId === leave.id && request.status === "pending") {
        request.status = "cancelled";
        request.resolvedAt = nowIso;
        request.rejectReason = "教練取消請假，待審補課已自動取消。";
      }
    });

    leave.revokedAt = nowIso;
    leave.revokedBy = activeCoachCode;
    leave.makeupEligible = false;

    lesson.attendanceStatus = "scheduled";
    lesson.calendarOccupied = true;
    lesson.charged = false;
    await tryCreateCalendarEventForLesson(lesson, "coach_revoke_leave", false);

    addLog(`教練 ${activeCoachCode} 已取消課程 ${lesson.id} 的請假。`);
    saveState();
    renderAll();
    try {
      await pushCloudLeaveRecord(leave, lesson);
    } catch (error) {
      addLog(`[雲端請假] 上傳教練取消請假失敗：${String(error?.message || "未知錯誤")}`);
      saveState();
    }
    await trySendEmailNotice(
      "leave_revoked_by_coach",
      {
        lessonId: lesson.id,
        studentCode: lesson.studentCode,
        coachCode: lesson.coachCode,
        lessonStartAt: lesson.startAt
      },
      `教練取消請假通知 ${lesson.id}`
    );
  }

  async function approveRequest(requestId) {
    if (!requireCoachWriteAccess()) {
      return;
    }
    const request = state.makeupRequests.find((item) => item.id === requestId);
    if (!request || request.coachCode !== activeCoachCode || request.status !== "pending") {
      alert("找不到待審申請，或你沒有權限。");
      return;
    }
    if (isSlotOccupied(request.startAt, request.coachCode, request.id)) {
      alert("申請時段已不可用。");
      return;
    }
    const newLesson = {
      id: newId("MAKEUP_LESSON"),
      studentCode: request.studentCode,
      coachCode: request.coachCode,
      startAt: request.startAt,
      sourceType: "MAKEUP",
      calendarEventId: "",
      calendarOccupied: true,
      attendanceStatus: "scheduled",
      charged: false
    };
    try {
      await tryCreateCalendarEventForLesson(newLesson, "makeup_approved", true);
    } catch (error) {
      alert(String(error?.message || "Google 日曆建立失敗，暫時無法核准。"));
      return;
    }
    request.status = "approved";
    request.resolvedAt = new Date().toISOString();
    state.lessons.push(newLesson);
    addLog(`教練 ${request.coachCode} 已核准補課申請 ${request.id}。`);
    saveState();
    renderAll();
    await trySendEmailNotice(
      "makeup_approved",
      {
        requestId: request.id,
        requestCode: request.code,
        studentCode: request.studentCode,
        coachCode: request.coachCode,
        startAt: request.startAt,
        lessonId: newLesson.id,
        eventId: newLesson.calendarEventId
      },
      `補課核准通知 ${request.code || request.id}`
    );
  }

  async function rejectRequest(requestId, reason) {
    if (!requireCoachWriteAccess()) {
      return;
    }
    const request = state.makeupRequests.find((item) => item.id === requestId);
    if (!request || request.coachCode !== activeCoachCode || request.status !== "pending") {
      alert("找不到待審申請，或你沒有權限。");
      return;
    }
    request.status = "rejected";
    request.resolvedAt = new Date().toISOString();
    request.rejectReason = reason || "教練退回";
    addLog(`教練 ${request.coachCode} 已退回補課申請 ${request.id}。`);
    saveState();
    renderAll();
    await trySendEmailNotice(
      "makeup_rejected",
      {
        requestId: request.id,
        requestCode: request.code,
        studentCode: request.studentCode,
        coachCode: request.coachCode,
        startAt: request.startAt,
        rejectReason: request.rejectReason
      },
      `補課退回通知 ${request.code || request.id}`
    );
  }

  function markLessonStatus(lessonId, statusType) {
    if (!requireCoachWriteAccess()) {
      return;
    }
    const lesson = getLessonById(lessonId);
    if (!lesson || lesson.coachCode !== activeCoachCode) {
      alert("找不到課程。");
      return;
    }
    const mapping = {
      temporary: { attendanceStatus: "temporary-leave", charged: true },
      noshow: { attendanceStatus: "no-show", charged: true },
      major: { attendanceStatus: "major-case", charged: false },
      reset: { attendanceStatus: "scheduled", charged: false }
    };
    const target = mapping[statusType];
    if (!target) {
      return;
    }
    const wasCharged = Boolean(lesson.charged);
    lesson.attendanceStatus = target.attendanceStatus;
    lesson.charged = target.charged;
    addLog(`教練 ${lesson.coachCode} 將課程 ${lesson.id} 標記為 ${target.attendanceStatus}。`);
    saveState();
    renderAll();
    if (!wasCharged && target.charged) {
      maybeSendChargeReminder(lesson.studentCode, "coach_mark_status").catch((error) => {
        console.error("billing reminder failed:", error);
      });
    }
  }

  function rescheduleLessonByCoach(lessonId) {
    if (!requireCoachWriteAccess()) {
      return;
    }
    const lesson = getLessonById(lessonId);
    if (!lesson || lesson.coachCode !== activeCoachCode) {
      alert("找不到課程。");
      return;
    }
    if (!lesson.calendarOccupied || lesson.attendanceStatus !== "scheduled") {
      alert("僅可調整「已排課」且仍佔用時段的課程。");
      return;
    }
    const now = new Date();
    if (new Date(lesson.startAt) <= now) {
      alert("已開始或已結束的課程不可調整時間。");
      return;
    }

    const currentDateKey = getDateKeyInTaipei(new Date(lesson.startAt));
    const currentTimeText = getTimeText(lesson.startAt);
    const inputDate = window.prompt("請輸入新上課日期（YYYY-MM-DD）", currentDateKey);
    if (!inputDate) {
      return;
    }
    const inputTime = window.prompt("請輸入新上課時間（HH:mm）", currentTimeText);
    if (!inputTime) {
      return;
    }
    const dateKey = normalizeDateInput(inputDate);
    const timeText = normalizeTimeInput(inputTime);
    if (!dateKey) {
      alert("日期格式錯誤，請用 YYYY-MM-DD 或 YYYY/MM/DD。");
      return;
    }
    if (!timeText) {
      alert("時間格式錯誤，請用 HH:mm。");
      return;
    }

    const nextStartAt = makeTaipeiDateTime(dateKey, timeText).toISOString();
    if (new Date(nextStartAt) <= now) {
      alert("新時段必須是未來時間。");
      return;
    }
    if (isSlotOccupied(nextStartAt, lesson.coachCode, "", lesson.id)) {
      alert("新時段已有課程或待審補課/教練請假，請改選其他時段。");
      return;
    }

    const oldStartAt = lesson.startAt;
    lesson.startAt = nextStartAt;
    addLog(`教練 ${lesson.coachCode} 將課程 ${lesson.id} 由 ${formatDateTime(oldStartAt)} 調整為 ${formatDateTime(nextStartAt)}。`);
    saveState();
    renderAll();
  }

  function isLessonPastForStudent(lesson) {
    if (!lesson || !lesson.startAt) {
      return false;
    }
    return new Date(lesson.startAt).getTime() <= Date.now();
  }

  function getStatusPill(status, lesson, viewMode) {
    let resolvedStatus = status;
    if (status === "scheduled" && viewMode === "student" && isLessonPastForStudent(lesson)) {
      resolvedStatus = "completed";
    }

    const map = {
      "leave-normal": "<span class=\"status normal\">正常請假</span>",
      "coach-leave": "<span class=\"status warn\">教練請假停課</span>",
      "calendar-removed": "<span class=\"status cancelled\">日曆已刪除</span>",
      pending: "<span class=\"status pending\">待審</span>",
      approved: "<span class=\"status approved\">已核准</span>",
      rejected: "<span class=\"status rejected\">已退回</span>",
      expired: "<span class=\"status expired\">已逾時</span>",
      cancelled: "<span class=\"status cancelled\">已取消</span>",
      "temporary-leave": "<span class=\"status warn\">臨時請假（扣堂）</span>",
      "no-show": "<span class=\"status warn\">未到課（扣堂）</span>",
      "major-case": "<span class=\"status normal\">重大急事（不扣堂）</span>",
      completed: "<span class=\"status normal\">已上課</span>",
      scheduled: "<span class=\"status\">已排課</span>"
    };
    return map[resolvedStatus] || `<span class="status">${resolvedStatus}</span>`;
  }

  function renderStudentLessons() {
    if (!el.studentLessonsTable) {
      return;
    }
    const student = getStudentByCode(activeStudentCode);
    if (!student) {
      el.studentLessonsTable.innerHTML = "<thead><tr><th>請先登入學生</th></tr></thead>";
      return;
    }
    const lessons = state.lessons
      .filter(
        (lesson) =>
          lesson.studentCode === student.code &&
          isGoogleSyncLesson(lesson) &&
          lesson.attendanceStatus !== "coach-leave" &&
          lesson.attendanceStatus !== "calendar-removed"
      )
      .sort((a, b) => new Date(a.startAt) - new Date(b.startAt));

    const rows = lessons.map((lesson) => {
      const leave = getLeaveByLesson(lesson.id);
      const approvedMakeup = getApprovedMakeupByLessonId(lesson.id);
      const canLeave = !leave && lesson.attendanceStatus === "scheduled" && isLeaveOpen(lesson);
      const canCancelLeave = !!leave && !approvedMakeup && lesson.attendanceStatus === "leave-normal" && isLeaveOpen(lesson);
      const cutoff = formatDateTime(getCutoffForLesson(lesson).toISOString());
      const statusInfo = approvedMakeup ? `<div class="hint">補課已核准：${formatDateTime(approvedMakeup.startAt)}</div>` : "";
      let actionHtml = "<span class=\"hint\">不可送出</span>";
      if (canLeave) {
        actionHtml = `<button class="student-action-btn" data-action="leave" data-id="${lesson.id}">送出請假申請</button>`;
      } else if (canCancelLeave) {
        actionHtml = `<button class="student-action-btn danger" data-action="cancel-leave" data-id="${lesson.id}">取消請假</button>`;
      } else if (approvedMakeup) {
        actionHtml = `<span class="hint">補課已核准：${formatDateTime(approvedMakeup.startAt)}</span>`;
      } else if (leave && lesson.attendanceStatus === "leave-normal" && !isLeaveOpen(lesson)) {
        actionHtml = "<span class=\"hint\">已過取消截止</span>";
      }
      return `<tr>
        <td>${formatDateTime(lesson.startAt)}</td>
        <td>${getStatusPill(lesson.attendanceStatus, lesson, "student")}${statusInfo}</td>
        <td>${lesson.charged ? "是" : "否"}</td>
        <td>${cutoff}</td>
        <td>${actionHtml}</td>
      </tr>`;
    }).join("");

    el.studentLessonsTable.innerHTML = `
      <thead><tr><th>時間</th><th>狀態</th><th>扣堂</th><th>請假截止</th><th>操作</th></tr></thead>
      <tbody>${rows || "<tr><td colspan=\"5\">目前沒有課程</td></tr>"}</tbody>
    `;
  }

  function renderStudentCalendar() {
    if (!el.studentCalendarGrid || !el.studentCalendarMonthLabel || !el.studentDayDetail) {
      return;
    }
    const student = getStudentByCode(activeStudentCode);
    if (!student) {
      el.studentCalendarMonthLabel.textContent = "--";
      el.studentCalendarGrid.innerHTML = "<div class='cal-weekday' style='grid-column:1 / span 7;'>請先登入學生</div>";
      el.studentDayDetail.innerHTML = "<h3>請先登入學生</h3>";
      closeStudentDayModal();
      return;
    }

    const meta = getMonthMeta(studentCalendarMonthStart);
    el.studentCalendarMonthLabel.textContent = `${meta.year}/${String(meta.month).padStart(2, "0")}`;

    const weekdays = ["日", "一", "二", "三", "四", "五", "六"];
    let html = weekdays.map((day) => `<div class="cal-weekday">${day}</div>`).join("");

    const totalCells = 42;
    for (let index = 0; index < totalCells; index += 1) {
      const dayNumber = index - meta.firstWeekday + 1;
      if (dayNumber < 1 || dayNumber > meta.daysInMonth) {
        html += `<div class="cal-cell dim"></div>`;
        continue;
      }
      const dateKey = `${String(meta.year).padStart(4, "0")}-${String(meta.month).padStart(2, "0")}-${String(dayNumber).padStart(2, "0")}`;
      const lessons = getStudentLessonsForDate(student.code, dateKey);
      const coachBlocks = getStudentCoachBlocksForDate(student.code, dateKey);
      const hasLessons = lessons.length > 0;
      const hasCoachBlocks = coachBlocks.length > 0;
      const todayKey = getDateKeyInTaipei(new Date());
      const classNames = [
        "cal-cell",
        hasLessons ? "has-lessons" : "no-lessons",
        hasCoachBlocks ? "has-coach-blocks" : "",
        dateKey === todayKey ? "today" : "",
        dateKey === selectedStudentDateKey ? "selected" : ""
      ].join(" ").trim();
      const todayBadge = dateKey === todayKey ? "<span class='today-badge'>今日</span>" : "";

      const lessonSnippets = lessons.slice(0, 3).map((lesson) => {
        const typeClass = lesson.sourceType === "MAKEUP"
          ? "makeup"
          : lesson.attendanceStatus === "leave-normal"
            ? "leave"
            : "";
        const prefix = lesson.sourceType === "MAKEUP" ? "補課" : "課";
        const selectedClass = lesson.id === selectedStudentLessonId ? "selected" : "";
        return `<button class="cal-item ${typeClass} ${selectedClass}" type="button" data-student-cal-date="${dateKey}" data-student-lesson-id="${lesson.id}">${prefix} ${getTimeText(lesson.startAt)} ${statusShortText(lesson)}</button>`;
      }).join("");
      const coachLeaveSnippet = coachBlocks.length
        ? `<button class="cal-item leave" type="button" data-student-cal-date="${dateKey}">教練請假 ${coachBlocks.length} 筆</button>`
        : "";
      const more = lessons.length > 3 ? `<button class="cal-item" type="button" data-student-cal-date="${dateKey}">+${lessons.length - 3}</button>` : "";
      const mainSnippet = lessonSnippets || (coachBlocks.length
        ? `<button class="cal-item leave" type="button" data-student-cal-date="${dateKey}">部分時段停課</button>`
        : `<button class="cal-item" type="button" data-student-cal-date="${dateKey}">無課程</button>`);

      html += `
        <div class="${classNames}" data-student-cal-date="${dateKey}">
          <div class="cal-date-row"><div class="cal-date">${dayNumber}</div>${todayBadge}</div>
          ${mainSnippet}
          ${coachLeaveSnippet}
          ${more}
        </div>
      `;
    }

    el.studentCalendarGrid.innerHTML = html;
    renderStudentDayDetail();
  }

  function statusShortText(lesson) {
    if (lesson.attendanceStatus === "leave-normal") {
      const approvedMakeup = getApprovedMakeupByLessonId(lesson.id);
      if (approvedMakeup) {
        return "已核補";
      }
      return "請假";
    }
    if (lesson.attendanceStatus === "temporary-leave") {
      return "臨假";
    }
    if (lesson.attendanceStatus === "no-show") {
      return "未到";
    }
    if (lesson.attendanceStatus === "major-case") {
      return "重大";
    }
    if (lesson.attendanceStatus === "calendar-removed") {
      return "已刪除";
    }
    if (lesson.attendanceStatus === "scheduled" && isLessonPastForStudent(lesson)) {
      return "已上課";
    }
    return "已排課";
  }

  function renderStudentDayDetail() {
    if (!el.studentDayDetail) {
      return;
    }
    const student = getStudentByCode(activeStudentCode);
    if (!student || !selectedStudentDateKey) {
      el.studentDayDetail.innerHTML = "<h3>請先選擇日期</h3>";
      return;
    }
    const lessons = getStudentLessonsForDate(student.code, selectedStudentDateKey);
    const coachBlocks = getStudentCoachBlocksForDate(student.code, selectedStudentDateKey);
    if (selectedStudentLessonId && !lessons.some((lesson) => lesson.id === selectedStudentLessonId)) {
      selectedStudentLessonId = "";
    }
    if (!lessons.length && !coachBlocks.length) {
      el.studentDayDetail.innerHTML = `<h3>${selectedStudentDateKey}</h3><p class="hint">這一天沒有課程。</p>`;
      return;
    }

    const rows = lessons.map((lesson) => {
      const leave = getLeaveByLesson(lesson.id);
      const approvedMakeup = getApprovedMakeupByLessonId(lesson.id);
      const canLeave = isGoogleSyncLesson(lesson) && lesson.attendanceStatus === "scheduled" && !leave && isLeaveOpen(lesson);
      const canCancelLeave = isGoogleSyncLesson(lesson) && !!leave && !approvedMakeup && lesson.attendanceStatus === "leave-normal" && isLeaveOpen(lesson);
      const cutoffText = isGoogleSyncLesson(lesson)
        ? formatDateTime(getCutoffForLesson(lesson).toISOString())
        : "-";
      const statusInfo = approvedMakeup ? `<div class='hint'>補課已核准：${formatDateTime(approvedMakeup.startAt)}</div>` : "";
      let actionHtml = "<span class='hint'>不可送出</span>";
      if (canLeave) {
        actionHtml = `<button class="student-action-btn" data-action="leave" data-id="${lesson.id}">送出請假申請</button>`;
      } else if (canCancelLeave) {
        actionHtml = `<button class="student-action-btn danger" data-action="cancel-leave" data-id="${lesson.id}">取消請假</button>`;
      } else if (approvedMakeup) {
        actionHtml = `<span class='hint'>補課已核准：${formatDateTime(approvedMakeup.startAt)}</span>`;
      } else if (leave && lesson.attendanceStatus === "leave-normal" && !isLeaveOpen(lesson)) {
        actionHtml = "<span class='hint'>已過取消截止</span>";
      }
      return `
        <tr ${lesson.id === selectedStudentLessonId ? "style='background:#eff6ff;'" : ""}>
          <td>${getTimeText(lesson.startAt)}</td>
          <td>${lesson.sourceType === "MAKEUP" ? "補課" : "原課"}</td>
          <td>${getStatusPill(lesson.attendanceStatus, lesson, "student")}${statusInfo}</td>
          <td>${lesson.charged ? "是" : "否"}</td>
          <td>${cutoffText}</td>
          <td>${actionHtml}</td>
        </tr>
      `;
    }).join("");

    const lessonCards = lessons.map((lesson) => {
      const leave = getLeaveByLesson(lesson.id);
      const approvedMakeup = getApprovedMakeupByLessonId(lesson.id);
      const canLeave = isGoogleSyncLesson(lesson) && lesson.attendanceStatus === "scheduled" && !leave && isLeaveOpen(lesson);
      const canCancelLeave = isGoogleSyncLesson(lesson) && !!leave && !approvedMakeup && lesson.attendanceStatus === "leave-normal" && isLeaveOpen(lesson);
      const cutoffText = isGoogleSyncLesson(lesson)
        ? formatDateTime(getCutoffForLesson(lesson).toISOString())
        : "-";
      const statusInfo = approvedMakeup ? `<div class="hint">補課已核准：${formatDateTime(approvedMakeup.startAt)}</div>` : "";
      let actionHtml = "<span class='hint'>目前不可送出請假</span>";
      if (canLeave) {
        actionHtml = `<button class="student-action-btn" data-action="leave" data-id="${lesson.id}">送出請假申請</button>`;
      } else if (canCancelLeave) {
        actionHtml = `<button class="student-action-btn danger" data-action="cancel-leave" data-id="${lesson.id}">取消請假</button>`;
      } else if (approvedMakeup) {
        actionHtml = `<span class="hint">補課已核准：${formatDateTime(approvedMakeup.startAt)}</span>`;
      } else if (leave && lesson.attendanceStatus === "leave-normal" && !isLeaveOpen(lesson)) {
        actionHtml = "<span class='hint'>已過取消截止</span>";
      }
      return `
        <article class="student-lesson-card ${lesson.id === selectedStudentLessonId ? "is-selected" : ""}">
          <div class="student-lesson-card-head">
            <div>
              <div class="student-lesson-time">${getTimeText(lesson.startAt)}</div>
              <div class="student-lesson-meta">${lesson.sourceType === "MAKEUP" ? "補課" : "原課"}</div>
            </div>
            <div>${getStatusPill(lesson.attendanceStatus, lesson, "student")}</div>
          </div>
          ${statusInfo}
          <div class="student-lesson-facts">
            <div>扣堂：${lesson.charged ? "是" : "否"}</div>
            <div>請假截止：${cutoffText}</div>
          </div>
          <div class="student-lesson-action">${actionHtml}</div>
        </article>
      `;
    }).join("");

    const blockRows = coachBlocks
      .map(
        (block) => `
      <tr>
        <td>${formatDateTime(block.startAt)} - ${getTimeText(block.endAt || block.startAt)}</td>
        <td>${block.reason || "-"}</td>
      </tr>`
      )
      .join("");

    const blockCards = coachBlocks
      .map((block) => `
        <div class="student-coach-block-card">
          <strong>教練請假時段</strong><br>
          ${formatDateTime(block.startAt)} - ${getTimeText(block.endAt || block.startAt)}<br>
          原因：${block.reason || "-"}
        </div>
      `)
      .join("");

    el.studentDayDetail.innerHTML = `
      <h3>${selectedStudentDateKey} 課程明細</h3>
      <div class="student-lesson-cards">${lessonCards || "<p class='hint'>無課程</p>"}</div>
      <div class="table-wrap student-day-table">
        <table>
          <thead><tr><th>時間</th><th>類型</th><th>狀態</th><th>扣堂</th><th>請假截止</th><th>操作</th></tr></thead>
          <tbody>${rows || "<tr><td colspan='6'>無課程</td></tr>"}</tbody>
        </table>
      </div>
      <div style="height:8px;"></div>
      <div class="student-coach-blocks">${blockCards || ""}</div>
      <div class="table-wrap student-coach-block-table">
        <table>
          <thead><tr><th>教練請假時段</th><th>原因</th></tr></thead>
          <tbody>${blockRows || "<tr><td colspan='2'>無教練請假</td></tr>"}</tbody>
        </table>
      </div>
    `;
  }

  function renderCoachCalendar() {
    if (!el.coachCalendarGrid || !el.coachCalendarMonthLabel || !el.coachDayDetail) {
      return;
    }
    if (!activeCoachCode) {
      el.coachCalendarMonthLabel.textContent = "--";
      el.coachCalendarGrid.innerHTML = "<div class='cal-weekday' style='grid-column:1 / span 7;'>請先登入教練</div>";
      el.coachDayDetail.innerHTML = "<h3>請先登入教練</h3>";
      if (el.coachCalendarSyncText) {
        el.coachCalendarSyncText.textContent = "請先登入教練後再同步。";
      }
      closeCoachDayModal();
      return;
    }
    if (el.coachCalendarSyncText && !el.coachCalendarSyncText.textContent.trim()) {
      el.coachCalendarSyncText.textContent = "可手動檢查是否有被外部刪除的事件。";
    }
    clampReadOnlyCoachMonth();

    const meta = getMonthMeta(coachCalendarMonthStart);
    el.coachCalendarMonthLabel.textContent = `${meta.year}/${String(meta.month).padStart(2, "0")}`;
    if (isCoachReadOnlyMode()) {
      const { minMonth, maxMonth } = getReadOnlyMonthBounds();
      const currentTime = getMonthStartTime(coachCalendarMonthStart);
      if (el.coachCalendarPrevBtn) {
        el.coachCalendarPrevBtn.disabled = currentTime <= minMonth.getTime();
      }
      if (el.coachCalendarNextBtn) {
        el.coachCalendarNextBtn.disabled = currentTime >= maxMonth.getTime();
      }
    } else {
      if (el.coachCalendarPrevBtn) {
        el.coachCalendarPrevBtn.disabled = false;
      }
      if (el.coachCalendarNextBtn) {
        el.coachCalendarNextBtn.disabled = false;
      }
    }

    const weekdays = ["日", "一", "二", "三", "四", "五", "六"];
    let html = weekdays.map((day) => `<div class="cal-weekday">${day}</div>`).join("");

    for (let index = 0; index < 42; index += 1) {
      const gridColumn = (index % 7) + 1;
      const gridRow = Math.floor(index / 7) + 2;
      const gridPlacement = `grid-column:${gridColumn};grid-row:${gridRow};`;
      const dayNumber = index - meta.firstWeekday + 1;
      if (dayNumber < 1 || dayNumber > meta.daysInMonth) {
        html += `<div class="cal-cell dim" style="${gridPlacement}"></div>`;
        continue;
      }
      const dateKey = `${String(meta.year).padStart(4, "0")}-${String(meta.month).padStart(2, "0")}-${String(dayNumber).padStart(2, "0")}`;
      const lessons = getCoachLessonsForDate(activeCoachCode, dateKey).filter(isReadOnlyVisibleCoachLesson);
      const pendings = getCoachPendingForDate(activeCoachCode, dateKey);
      const blocks = getCoachBlocksForDate(activeCoachCode, dateKey);
      const todayKey = getDateKeyInTaipei(new Date());
      const classNames = [
        "cal-cell",
        dateKey === todayKey ? "today" : "",
        dateKey === selectedCoachDateKey ? "selected" : "",
        lessons.length ? "" : "no-lessons"
      ].join(" ").trim();
      const todayBadge = dateKey === todayKey ? "<span class='today-badge'>今日</span>" : "";
      const readOnly = isCoachReadOnlyMode();
      const renderCalendarItem = (text, extraClass = "", attrs = "") => {
        const classText = ["cal-item", extraClass].filter(Boolean).join(" ");
        if (readOnly) {
          return `<span class="${classText}">${text}</span>`;
        }
        return `<button class="${classText}" type="button" data-coach-cal-date="${dateKey}" ${attrs}>${text}</button>`;
      };

      const readonlyLessonLimit = 2;
      const getLessonHour = (lesson) => Number(getDateTimePartsInTaipei(new Date(lesson.startAt)).hour || 0);
      const renderReadonlyPeriod = (label, periodLessons) => {
        if (!periodLessons.length) {
          return "";
        }
        const visibleLessons = periodLessons.slice(0, readonlyLessonLimit);
        const hiddenCount = periodLessons.length - visibleLessons.length;
        return `
          <div class="readonly-period">
            <span class="readonly-period-label">${label}</span>
            ${visibleLessons.map((lesson) => `<span class="cal-item readonly-name">${getStudentDisplayName(lesson.studentCode)}</span>`).join("")}
            ${hiddenCount > 0 ? `<span class="cal-item readonly-more">+${hiddenCount}</span>` : ""}
          </div>
        `;
      };
      const lessonSnippets = readOnly
        ? [
          renderReadonlyPeriod("上午", lessons.filter((lesson) => getLessonHour(lesson) < 12)),
          renderReadonlyPeriod("下午", lessons.filter((lesson) => getLessonHour(lesson) >= 12))
        ].join("")
        : lessons.slice(0, 3).map((lesson) => {
        const typeClass = lesson.sourceType === "MAKEUP" ? "makeup" : "";
        const prefix = lesson.sourceType === "MAKEUP" ? "補課" : "原課";
        const selectedClass = lesson.id === selectedCoachLessonId ? "selected" : "";
        return renderCalendarItem(
          `${prefix} ${getTimeText(lesson.startAt)} ${getStudentDisplayName(lesson.studentCode)}`,
          [typeClass, selectedClass].filter(Boolean).join(" "),
          `data-coach-lesson-id="${lesson.id}"`
        );
      }).join("");
      const lessonMoreSnippet = !readOnly && lessons.length > 3
        ? renderCalendarItem(`+${lessons.length - 3}`)
        : "";
      const pendingSnippet = pendings.length
        ? (readOnly ? "" : renderCalendarItem(`待審 ${pendings.length}`, "pending"))
        : "";
      const blockSnippet = blocks.length
        ? (readOnly ? "" : renderCalendarItem(`教練請假 ${blocks.length}`))
        : "";

      html += `
          <div class="${classNames}" data-coach-cal-date="${dateKey}" style="${gridPlacement}">
            <div class="cal-date-row"><div class="cal-date">${dayNumber}</div>${todayBadge}</div>
          ${lessonSnippets || (readOnly ? "" : renderCalendarItem("無課程"))}
          ${lessonMoreSnippet}
          ${pendingSnippet}
          ${blockSnippet}
        </div>
      `;
    }

    el.coachCalendarGrid.innerHTML = html;
    renderCoachDayDetail();
  }

  function renderCoachDayDetail() {
    if (!el.coachDayDetail) {
      return;
    }
    if (!activeCoachCode || !selectedCoachDateKey) {
      el.coachDayDetail.innerHTML = "<h3>請先選擇日期</h3>";
      return;
    }

    const lessons = getCoachLessonsForDate(activeCoachCode, selectedCoachDateKey).filter(isReadOnlyVisibleCoachLesson);
    const pendings = getCoachPendingForDate(activeCoachCode, selectedCoachDateKey);
    const blocks = getCoachBlocksForDate(activeCoachCode, selectedCoachDateKey);
    if (selectedCoachLessonId && !lessons.some((lesson) => lesson.id === selectedCoachLessonId)) {
      selectedCoachLessonId = "";
    }

    if (!lessons.length && !pendings.length && !blocks.length) {
      el.coachDayDetail.innerHTML = `<h3>${selectedCoachDateKey}</h3><p class="hint">無課程、無待審、無教練請假。</p>`;
      return;
    }

    const lessonRows = lessons
      .filter((lesson) => lesson.attendanceStatus !== "coach-leave")
      .map((lesson) => `
      <tr ${lesson.id === selectedCoachLessonId ? "style='background:#fff7ed;'" : ""}>
        <td>${getStudentDisplayName(lesson.studentCode)}</td>
        <td>${getTimeText(lesson.startAt)}</td>
        <td>${lesson.sourceType === "MAKEUP" ? "補課" : "原課"}</td>
        <td>${getStatusPill(lesson.attendanceStatus)}</td>
        <td>${lesson.charged ? "是" : "否"}</td>
        <td>${renderCoachLessonActionButtons(lesson)}</td>
      </tr>
    `).join("");

    const readOnly = isCoachReadOnlyMode();
    const pendingRows = pendings.map((request) => `
      <tr>
        <td>${request.code || request.id}</td>
        <td>${getStudentDisplayName(request.studentCode)}</td>
        <td>${getTimeText(request.startAt)}</td>
        <td>${formatDateTime(request.pendingAt)}</td>
        <td>${getStatusPill(request.status)}</td>
        <td>
          ${readOnly ? "<span class='hint'>唯讀</span>" : `<div class="btn-row">
            <button data-day-review-id="${request.id}" data-day-review-action="approve" class="primary">核准</button>
            <button data-day-review-id="${request.id}" data-day-review-action="reject" class="danger">退回</button>
          </div>`}
        </td>
      </tr>
    `).join("");

    const blockRows = blocks.map((block) => `
      <tr>
        <td>${formatDateTime(block.startAt)} - ${getTimeText(block.endAt || block.startAt)}</td>
        <td>${block.reason || "-"}</td>
      </tr>
    `).join("");

    el.coachDayDetail.innerHTML = `
      <h3>${selectedCoachDateKey} 教練日檢視</h3>
      <div class="table-wrap">
        <table>
          <thead><tr><th>學生</th><th>時間</th><th>類型</th><th>狀態</th><th>扣堂</th><th>操作</th></tr></thead>
          <tbody>${lessonRows || "<tr><td colspan='6'>無課程</td></tr>"}</tbody>
        </table>
      </div>
      <div style="height:8px;"></div>
      <div class="table-wrap">
        <table>
          <thead><tr><th>申請</th><th>學生</th><th>補課時段</th><th>提交時間</th><th>狀態</th><th>操作</th></tr></thead>
          <tbody>${pendingRows || "<tr><td colspan='6'>無待審申請</td></tr>"}</tbody>
        </table>
      </div>
      <div style="height:8px;"></div>
      <div class="table-wrap">
        <table>
          <thead><tr><th>教練請假時段</th><th>原因</th></tr></thead>
          <tbody>${blockRows || "<tr><td colspan='2'>無教練請假時段</td></tr>"}</tbody>
        </table>
      </div>
      <p class="hint">可直接在本日檢視中完成課程調整與補課審核。</p>
    `;
  }

  function renderCoachLessonActionButtons(lesson) {
    if (isCoachReadOnlyMode()) {
      return "<span class=\"hint\">唯讀</span>";
    }
    const isNormalLeave = lesson.attendanceStatus === "leave-normal";
    if (isNormalLeave) {
      return `
        <div class="btn-row">
          <button data-coach-lesson-id="${lesson.id}" data-coach-mark="cancel-leave" class="danger">取消請假</button>
        </div>
      `;
    }
    return `
      <div class="btn-row">
        <button data-coach-lesson-id="${lesson.id}" data-coach-mark="reschedule">調整時間</button>
        <button data-coach-lesson-id="${lesson.id}" data-coach-mark="temporary">臨時請假</button>
        <button data-coach-lesson-id="${lesson.id}" data-coach-mark="noshow">未到課</button>
        <button data-coach-lesson-id="${lesson.id}" data-coach-mark="major">重大急事</button>
        <button data-coach-lesson-id="${lesson.id}" data-coach-mark="reset">重置</button>
      </div>
    `;
  }

  function renderCoachLeaveTable() {
    if (!el.coachLeaveTable) {
      return;
    }
    if (!activeCoachCode) {
      el.coachLeaveTable.innerHTML = "<thead><tr><th>請先登入教練</th></tr></thead>";
      return;
    }
    const blocks = (state.coachBlocks || [])
      .filter((block) => block.coachCode === activeCoachCode)
      .sort((a, b) => new Date(a.startAt) - new Date(b.startAt));
    const rows = blocks.map((block) => `
      <tr>
        <td>${formatDateTime(block.startAt)} - ${getTimeText(block.endAt || block.startAt)}</td>
        <td>${block.reason || "-"}</td>
        <td><button class="danger" data-action="remove-coach-block" data-id="${block.id}">移除</button></td>
      </tr>
    `).join("");
    el.coachLeaveTable.innerHTML = `
      <thead><tr><th>教練請假時段</th><th>原因</th><th>操作</th></tr></thead>
      <tbody>${rows || "<tr><td colspan='3'>目前沒有教練請假時段</td></tr>"}</tbody>
    `;
  }

  function renderMakeupSection() {
    if (!el.makeupLeaveSelect || !el.studentMakeupTable) {
      return;
    }
    const student = getStudentByCode(activeStudentCode);
    const leaves = state.leaveRequests.filter(
      (leave) => leave.studentCode === activeStudentCode && leave.makeupEligible && !leave.revokedAt
    );
    el.makeupLeaveSelect.innerHTML = `<option value="">請選擇</option>${leaves.map((leave) => {
      const lesson = getLessonById(leave.lessonId);
      const existing = getMakeupByLeave(leave.id);
      const disabled = !!existing;
      const label = `${leave.id} / 原上課 ${lesson ? formatDateTime(lesson.startAt) : leave.lessonId}${disabled ? "（已有待審或核准）" : ""}`;
      return `<option value="${leave.id}" ${disabled ? "disabled" : ""}>${label}</option>`;
    }).join("")}`;

    if (!student) {
      el.studentMakeupTable.innerHTML = "<thead><tr><th>請先登入學生</th></tr></thead>";
      return;
    }

    const requests = state.makeupRequests
      .filter((request) => request.studentCode === student.code)
      .sort((a, b) => new Date(b.pendingAt) - new Date(a.pendingAt));

    const rows = requests.map((request) => `
      <tr>
        <td>${request.code || request.id}</td>
        <td>${formatDateTime(request.startAt)}</td>
        <td>${getStatusPill(request.status)}</td>
        <td>${formatDateTime(request.pendingAt)}</td>
        <td>${request.resolvedAt ? formatDateTime(request.resolvedAt) : "-"}</td>
        <td>${request.rejectReason || "-"}</td>
        <td>${request.status === "pending" ? `<button class="danger" data-action="cancel-pending" data-id="${request.id}">取消</button>` : "-"}</td>
      </tr>
    `).join("");

    el.studentMakeupTable.innerHTML = `
      <thead><tr><th>申請</th><th>補課時段</th><th>狀態</th><th>提交時間</th><th>處理時間</th><th>原因</th><th>操作</th></tr></thead>
      <tbody>${rows || "<tr><td colspan=\"7\">目前沒有補課申請</td></tr>"}</tbody>
    `;
  }

  function renderCoachPending() {
    if (!el.coachPendingTable) {
      return;
    }
    if (!activeCoachCode) {
      el.coachPendingTable.innerHTML = "<thead><tr><th>請先登入教練</th></tr></thead>";
      if (el.coachReviewSummary) {
        el.coachReviewSummary.textContent = "待審：0";
      }
      if (el.coachReviewUrgent) {
        el.coachReviewUrgent.textContent = "6 小時內逾時：0";
      }
      if (el.coachReviewStudentFilter) {
        el.coachReviewStudentFilter.innerHTML = "<option value='ALL'>全部學生</option>";
      }
      return;
    }
    const allPendings = state.makeupRequests
      .filter((request) => request.coachCode === activeCoachCode && request.status === "pending")
      .sort((a, b) => new Date(a.pendingAt) - new Date(b.pendingAt));

    const urgentCount = allPendings.filter((request) => getPendingHoursLeft(request) <= 6).length;
    if (el.coachReviewSummary) {
      el.coachReviewSummary.textContent = `待審：${allPendings.length}`;
    }
    if (el.coachReviewUrgent) {
      el.coachReviewUrgent.textContent = `6 小時內逾時：${urgentCount}`;
    }

    if (el.coachReviewStudentFilter) {
      const students = [...new Set(allPendings.map((request) => request.studentCode))].sort();
      if (coachReviewFilterStudent !== "ALL" && !students.includes(coachReviewFilterStudent)) {
        coachReviewFilterStudent = "ALL";
      }
      el.coachReviewStudentFilter.innerHTML = `<option value="ALL">全部學生</option>${students
        .map((studentCode) => `<option value="${studentCode}" ${coachReviewFilterStudent === studentCode ? "selected" : ""}>${getStudentDisplayName(studentCode)}</option>`)
        .join("")}`;
      if (el.coachReviewStudentFilter.value !== coachReviewFilterStudent) {
        el.coachReviewStudentFilter.value = coachReviewFilterStudent;
      }
    }

    const pendings =
      coachReviewFilterStudent === "ALL"
        ? allPendings
        : allPendings.filter((request) => request.studentCode === coachReviewFilterStudent);

    const rows = pendings.map((request) => {
      const leftHours = getPendingHoursLeft(request);
      const originalLesson = getLessonById(request.lessonId);
      const originalTimeText = originalLesson ? formatDateTime(originalLesson.startAt) : "-";
      return `
      <tr>
        <td>${request.code || request.id}</td>
        <td>${getStudentDisplayName(request.studentCode)}</td>
        <td>${originalTimeText}</td>
        <td>${formatDateTime(request.startAt)}</td>
        <td>${formatDateTime(request.pendingAt)}</td>
        <td>${leftHours.toFixed(1)} 小時</td>
        <td><span class="hint">請在月曆日檢視操作</span></td>
      </tr>
    `;
    }).join("");

    el.coachPendingTable.innerHTML = `
      <thead><tr><th>申請</th><th>學生</th><th>原上課時間</th><th>補課時段</th><th>提交時間</th><th>剩餘時數</th><th>操作</th></tr></thead>
      <tbody>${rows || "<tr><td colspan=\"7\">目前沒有待審補課申請</td></tr>"}</tbody>
    `;
  }

  function getCalendarRemovedLessonsForCoach(coachCode) {
    return state.lessons
      .filter(
        (lesson) =>
          lesson.coachCode === coachCode &&
          lesson.attendanceStatus === "calendar-removed"
      )
      .sort((a, b) => new Date(a.startAt) - new Date(b.startAt));
  }

  function applyRemovedSnapshot(lesson, snapshot) {
    lesson.attendanceStatus = snapshot.attendanceStatus;
    lesson.calendarOccupied = snapshot.calendarOccupied;
    lesson.charged = snapshot.charged;
    lesson.calendarEventId = snapshot.calendarEventId || "";
    lesson.calendarRemovedAt = snapshot.calendarRemovedAt;
    lesson.calendarRemovedReason = snapshot.calendarRemovedReason;
    lesson.beforeCalendarRemoved = snapshot.beforeCalendarRemoved;
  }

  async function restoreCalendarRemovedLesson(lessonId, options = {}) {
    if (!requireCoachWriteAccess()) {
      return false;
    }
    const { skipConfirm = false, skipRender = false, silent = false } = options;
    const lesson = getLessonById(lessonId);
    if (!lesson || lesson.coachCode !== activeCoachCode || lesson.attendanceStatus !== "calendar-removed") {
      if (!silent) {
        alert("找不到可還原的課程。");
      }
      return false;
    }
    if (!skipConfirm) {
      const ok = window.confirm(`確認還原這堂課？\n${formatDateTime(lesson.startAt)} / ${getStudentDisplayName(lesson.studentCode)}`);
      if (!ok) {
        return false;
      }
    }

    const removedSnapshot = {
      attendanceStatus: lesson.attendanceStatus,
      calendarOccupied: lesson.calendarOccupied,
      charged: lesson.charged,
      calendarEventId: lesson.calendarEventId || "",
      calendarRemovedAt: lesson.calendarRemovedAt,
      calendarRemovedReason: lesson.calendarRemovedReason,
      beforeCalendarRemoved: lesson.beforeCalendarRemoved
    };
    const previous = lesson.beforeCalendarRemoved || {
      attendanceStatus: "scheduled",
      calendarOccupied: true,
      charged: false,
      calendarEventId: ""
    };

    const snapshotEventId = String(previous.calendarEventId || "").trim();
    const restoreReason = String(lesson.calendarRemovedReason || "").trim().toLowerCase();
    const preferSnapshotRestore = restoreReason === "google_calendar_deleted" && !!snapshotEventId;

    if (preferSnapshotRestore) {
      lesson.attendanceStatus = previous.attendanceStatus || "scheduled";
      lesson.calendarOccupied = true;
      lesson.charged = Boolean(previous.charged);
      lesson.calendarEventId = snapshotEventId;
      delete lesson.calendarRemovedAt;
      delete lesson.calendarRemovedReason;
      delete lesson.beforeCalendarRemoved;
      addLog(`[日曆同步] 已依快照還原課程 ${lesson.id}（不重建 Google 事件）。`);
      saveState();
      if (!skipRender) {
        renderAll();
      }
      return true;
    }

    const now = new Date();
    if (new Date(lesson.startAt) <= now) {
      if (!silent) {
        alert("過去課程不可重建 Google 事件。");
      }
      return false;
    }

    lesson.attendanceStatus = previous.attendanceStatus || "scheduled";
    lesson.calendarOccupied = true;
    lesson.charged = Boolean(previous.charged);
    lesson.calendarEventId = "";
    delete lesson.calendarRemovedAt;
    delete lesson.calendarRemovedReason;

    try {
      await tryCreateCalendarEventForLesson(lesson, "coach_restore_calendar_removed", true);
      delete lesson.beforeCalendarRemoved;
      addLog(`[日曆同步] 已還原課程 ${lesson.id}（${formatDateTime(lesson.startAt)}）。`);
      saveState();
      if (!skipRender) {
        renderAll();
      }
      return true;
    } catch (error) {
      applyRemovedSnapshot(lesson, removedSnapshot);
      saveState();
      if (!skipRender) {
        renderAll();
      }
      if (!silent) {
        alert(`還原失敗：${String(error?.message || "未知錯誤")}`);
      }
      return false;
    }
  }

  async function restoreAllCalendarRemovedLessons() {
    if (!requireCoachWriteAccess()) {
      return;
    }
    if (!activeCoachCode) {
      alert("請先登入教練。");
      return;
    }
    const lessons = getCalendarRemovedLessonsForCoach(activeCoachCode).filter((lesson) => {
      const previous = lesson.beforeCalendarRemoved || {};
      const snapshotEventId = String(previous.calendarEventId || "").trim();
      const restoreReason = String(lesson.calendarRemovedReason || "").trim().toLowerCase();
      const canSnapshotRestore = restoreReason === "google_calendar_deleted" && !!snapshotEventId;
      return canSnapshotRestore || new Date(lesson.startAt) > new Date();
    });
    if (!lessons.length) {
      alert("目前沒有可還原的課程。");
      return;
    }
    const ok = window.confirm(`將還原 ${lessons.length} 堂課，並重建 Google 日曆事件。是否繼續？`);
    if (!ok) {
      return;
    }
    let success = 0;
    let failed = 0;
    for (const lesson of lessons) {
      const restored = await restoreCalendarRemovedLesson(lesson.id, { skipConfirm: true, skipRender: true, silent: true });
      if (restored) {
        success += 1;
      } else {
        failed += 1;
      }
    }
    const summary = `批次還原完成：成功 ${success} 堂，失敗 ${failed} 堂。`;
    if (el.coachCalendarSyncText) {
      el.coachCalendarSyncText.textContent = summary;
    }
    addLog(`[日曆同步] ${summary}`);
    saveState();
    renderAll();
  }

  function renderCalendarRemovedPanel() {
    if (!el.calendarRemovedTable) {
      return;
    }
    if (!activeCoachCode) {
      el.calendarRemovedTable.innerHTML = "<thead><tr><th>請先登入教練</th></tr></thead>";
      if (el.calendarRemovedSummary) {
        el.calendarRemovedSummary.textContent = "待還原：0";
      }
      if (el.calendarRemovedToggleBtn) {
        el.calendarRemovedToggleBtn.hidden = true;
      }
      isCalendarRemovedExpanded = false;
      return;
    }
    const lessons = getCalendarRemovedLessonsForCoach(activeCoachCode);
    if (el.calendarRemovedSummary) {
      el.calendarRemovedSummary.textContent = `待還原：${lessons.length}`;
    }
    const shouldCollapse = lessons.length > CALENDAR_REMOVED_PREVIEW_COUNT;
    if (shouldCollapse) {
      if (el.calendarRemovedToggleBtn) {
        el.calendarRemovedToggleBtn.hidden = false;
      }
    } else {
      isCalendarRemovedExpanded = false;
      if (el.calendarRemovedToggleBtn) {
        el.calendarRemovedToggleBtn.hidden = true;
      }
    }
    const visibleLessons =
      shouldCollapse && !isCalendarRemovedExpanded
        ? lessons.slice(0, CALENDAR_REMOVED_PREVIEW_COUNT)
        : lessons;
    const hiddenCount = lessons.length - visibleLessons.length;
    if (el.calendarRemovedToggleBtn && shouldCollapse) {
      el.calendarRemovedToggleBtn.textContent = isCalendarRemovedExpanded
        ? "收合清單"
        : `展開其餘 ${hiddenCount} 筆`;
    }
    const now = new Date();
    const rows = visibleLessons.map((lesson) => {
      const canRestore = new Date(lesson.startAt) > now;
      return `
      <tr>
        <td>${getStudentDisplayName(lesson.studentCode)}</td>
        <td>${formatDateTime(lesson.startAt)}</td>
        <td>${lesson.sourceType === "MAKEUP" ? "補課" : "原課"}</td>
        <td>${lesson.calendarRemovedAt ? formatDateTime(lesson.calendarRemovedAt) : "-"}</td>
        <td>${lesson.calendarRemovedReason || "google_calendar_deleted"}</td>
        <td>${canRestore ? `<button data-action="restore-removed" data-id="${lesson.id}">還原</button>` : "<span class='hint'>已過期</span>"}</td>
      </tr>
    `;
    }).join("");
    const collapsedHintRow =
      hiddenCount > 0
        ? `<tr><td colspan="6"><span class="hint">尚有 ${hiddenCount} 筆未顯示，請點「展開其餘 ${hiddenCount} 筆」。</span></td></tr>`
        : "";
    el.calendarRemovedTable.innerHTML = `
      <thead><tr><th>學生</th><th>課程時間</th><th>類型</th><th>移除時間</th><th>原因</th><th>操作</th></tr></thead>
      <tbody>${rows || "<tr><td colspan='6'>目前沒有同步移除課程</td></tr>"}${collapsedHintRow}</tbody>
    `;
  }

  function getCompensationStatusPill(status) {
    const mapped = status === "completed" ? "approved" : status === "failed" ? "rejected" : "pending";
    return getStatusPill(mapped);
  }

  function summarizeCompensationTask(task) {
    const payload = task?.payload || {};
    if (task?.type === "sendEmail") {
      return `模板 ${payload.template || "-"} / ${payload.studentCode || payload.coachCode || "-"}`;
    }
    if (task?.type === "deleteEvent" || task?.type === "deleteSingleEvent" || task?.type === "createEvent") {
      const lessonTypeText = payload.sourceType === "MAKEUP" ? "補課" : "原課";
      const studentText = payload.studentCode || "-";
      const timeText = payload.startAt ? formatDateTime(payload.startAt) : "-";
      return `${lessonTypeText} / ${studentText} / ${timeText}`;
    }
    return "-";
  }

  async function retryCompensationTask(taskId) {
    if (!requireCoachWriteAccess()) {
      return;
    }
    const tasks = Array.isArray(state.compensationTasks) ? state.compensationTasks : [];
    const task = tasks.find((item) => item.id === taskId);
    if (!task) {
      return;
    }
    task.attemptCount = Number(task.attemptCount || 0) + 1;
    task.lastTriedAt = new Date().toISOString();
    try {
      const finalPayload = { ...(task.payload || {}) };
      if (task.type === "sendEmail") {
        const recipients = resolveNoticeRecipients(finalPayload);
        if (recipients.length) {
          finalPayload.to = recipients.join(",");
        }
        task.payload = finalPayload;
      } else if (task.type === "createEvent" && isNormalLeaveCalendarCreateBlocked(finalPayload)) {
        task.status = "completed";
        task.completedAt = new Date().toISOString();
        task.lastError = "";
        addLog(`[補償重送] 已略過正常請假的建立事件任務（${task.id}）。`);
        saveState();
        renderAll();
        return;
      } else if (task.type === "deleteEvent" || task.type === "deleteSingleEvent") {
        Object.assign(finalPayload, addSingleEventDeleteGuards(finalPayload));
        task.payload = finalPayload;
      }
      const result = await callAppsScriptApi(task.type, finalPayload);
      task.status = "completed";
      task.completedAt = new Date().toISOString();
      task.lastError = "";
      if (task.type === "createEvent" && task.payload?.lessonId) {
        const lesson = getLessonById(task.payload.lessonId);
        if (lesson) {
          lesson.calendarEventId = String(result?.eventId || result?.calendarEventId || lesson.calendarEventId || "").trim();
        }
      } else if ((task.type === "deleteEvent" || task.type === "deleteSingleEvent") && task.payload?.lessonId) {
        const lesson = getLessonById(task.payload.lessonId);
        if (lesson) {
          lesson.calendarEventId = "";
        }
      }
      addLog(`[補償重送] ${task.type} 成功（${task.id}）`);
    } catch (error) {
      const message = String(error?.message || "未知錯誤");
      const deleteAlreadyHandled =
        (task.type === "deleteEvent" || task.type === "deleteSingleEvent") &&
        /event not found|already be deleted|already deleted/i.test(message);
      if (deleteAlreadyHandled) {
        task.status = "completed";
        task.completedAt = new Date().toISOString();
        task.lastError = "";
        addLog(`[補償重送] ${task.type} 視為成功（事件已不存在，${task.id}）。`);
      } else {
        task.status = "failed";
        task.lastError = message;
        addLog(`[補償重送] ${task.type} 失敗（${task.id}）：${task.lastError}`);
      }
    }
    saveState();
    renderAll();
  }

  async function retryAllCompensationTasks() {
    if (!requireCoachWriteAccess()) {
      return;
    }
    const tasks = (state.compensationTasks || []).filter((item) => item.status !== "completed");
    for (const task of tasks) {
      await retryCompensationTask(task.id);
    }
  }

  function deleteCompensationTask(taskId) {
    if (!requireCoachWriteAccess()) {
      return;
    }
    const tasks = Array.isArray(state.compensationTasks) ? state.compensationTasks : [];
    const task = tasks.find((item) => item.id === taskId);
    if (!task) {
      return;
    }
    const ok = window.confirm(`確定刪除這筆補償任務？\n${task.type || "-"} / ${summarizeCompensationTask(task)}`);
    if (!ok) {
      return;
    }
    state.compensationTasks = tasks.filter((item) => item.id !== taskId);
    addLog(`[補償任務] 已刪除 ${task.type || "-"}（${task.id}）。`);
    saveState();
    renderCompensationPanel();
    renderLogs();
  }

  function renderCompensationPanel() {
    if (!el.compensationTable) {
      return;
    }
    const tasks = (state.compensationTasks || []).slice().sort((a, b) => new Date(b.createdAt || 0) - new Date(a.createdAt || 0));
    const pendingCount = tasks.filter((item) => item.status !== "completed").length;
    if (el.compensationSummary) {
      el.compensationSummary.textContent = `待處理：${pendingCount}`;
    }
    const rows = tasks.map((task) => {
      const retryButton = task.status === "completed" ? "" : `<button data-action="retry-comp" data-id="${task.id}">重送</button>`;
      const deleteButton = `<button class="danger" data-action="delete-comp" data-id="${task.id}">刪除</button>`;
      const actionHtml = [retryButton, deleteButton].filter(Boolean).join(" ");
      return `
      <tr>
        <td>${formatDateTime(task.createdAt || new Date().toISOString())}</td>
        <td>${task.type || "-"}</td>
        <td>${summarizeCompensationTask(task)}</td>
        <td>${task.reason || task.lastError || "-"}</td>
        <td>${getCompensationStatusPill(task.status || "pending")}</td>
        <td>${actionHtml || "-"}</td>
      </tr>
    `;
    }).join("");
    el.compensationTable.innerHTML = `
      <thead><tr><th>建立時間</th><th>類型</th><th>摘要</th><th>原因</th><th>狀態</th><th>操作</th></tr></thead>
      <tbody>${rows || "<tr><td colspan='6'>目前沒有補償任務</td></tr>"}</tbody>
    `;
  }

  function renderLogs() {
    if (!el.eventLog) {
      return;
    }
    const isStudentPage = String(window.location.pathname || "").toLowerCase().includes("leave-student");
    if (!isStudentPage) {
      el.eventLog.innerHTML = state.eventLog.slice(0, 25).map((item) => `<li>${formatDateTime(item.at)} - ${item.message}</li>`).join("");
      return;
    }
    const studentCode = normalizeParticipantCode(activeStudentCode);
    if (!studentCode) {
      el.eventLog.innerHTML = "<li>載入學生後會顯示自己的事件紀錄。</li>";
      return;
    }
    const relatedIds = new Set([studentCode]);
    state.lessons
      .filter((lesson) => lesson.studentCode === studentCode)
      .forEach((lesson) => relatedIds.add(lesson.id));
    state.leaveRequests
      .filter((leave) => leave.studentCode === studentCode)
      .forEach((leave) => relatedIds.add(leave.id));
    state.makeupRequests
      .filter((request) => request.studentCode === studentCode)
      .forEach((request) => relatedIds.add(request.id));
    const otherStudentCodes = state.students
      .map((student) => normalizeParticipantCode(student.code))
      .filter((code) => code && code !== studentCode);
    const visibleLogs = state.eventLog
      .filter((item) => {
        const message = String(item.message || "");
        if (otherStudentCodes.some((code) => message.includes(code))) {
          return false;
        }
        return Array.from(relatedIds).some((id) => id && message.includes(id));
      })
      .slice(0, 25);
    el.eventLog.innerHTML = visibleLogs.length
      ? visibleLogs.map((item) => `<li>${formatDateTime(item.at)} - ${item.message}</li>`).join("")
      : "<li>目前沒有這位學生的事件紀錄。</li>";
  }

  function getHiddenStudentNameSet() {
    return new Set(
      (Array.isArray(window.APP_CONFIG?.hiddenStudentNames) ? window.APP_CONFIG.hiddenStudentNames : [])
        .map((name) => normalizeLooseText(name))
        .filter(Boolean)
    );
  }

  function getAliasTargetCodeByStudentName(studentName) {
    const aliases = window.APP_CONFIG?.calendarStudentAliases && typeof window.APP_CONFIG.calendarStudentAliases === "object"
      ? window.APP_CONFIG.calendarStudentAliases
      : {};
    const looseName = normalizeLooseText(studentName);
    if (!looseName) {
      return "";
    }
    const matched = Object.entries(aliases).find(([alias]) => normalizeLooseText(alias) === looseName);
    return normalizeParticipantCode(matched?.[1] || "");
  }

  function isHiddenChargeStudent(student) {
    const code = normalizeParticipantCode(student?.code);
    const looseName = normalizeLooseText(student?.name || "");
    if (!code) {
      return true;
    }
    if (["STU001", "STU002", "STU003"].includes(code)) {
      return true;
    }
    if (getHiddenStudentNameSet().has(looseName)) {
      return true;
    }
    const aliasTargetCode = getAliasTargetCodeByStudentName(student?.name || "");
    return Boolean(aliasTargetCode && aliasTargetCode !== code);
  }

  function getVisibleChargeStudents() {
    return (state.students || []).filter((student) => !isHiddenChargeStudent(student));
  }

  function getLastCompletedLesson(lessons) {
    const now = Date.now();
    const completedLessons = (lessons || [])
      .filter((lesson) => new Date(lesson.startAt).getTime() <= now)
      .sort((a, b) => new Date(b.startAt) - new Date(a.startAt));
    return completedLessons[0] || null;
  }

  function renderStudentOverviewPanel() {
    if (!el.studentOverviewTable) {
      return;
    }
    const students = getVisibleChargeStudents()
      .filter((student) => !activeCoachCode || student.coachCode === activeCoachCode)
      .sort((a, b) => String(a.name || a.code).localeCompare(String(b.name || b.code), "zh-Hant"));
    const rows = students.map((student) => {
      const stats = getStudentChargeStats(student.code);
      const billingCycle = getStudentBillingCycle(stats);
      const lastLesson = getLastCompletedLesson(stats.lessons);
      const paymentClass = billingCycle.effectivePaymentStatus === "unpaid"
        ? "rejected"
        : billingCycle.effectivePaymentStatus === "paid"
          ? "approved"
          : "pending";
      return `
        <tr>
          <td>${student.name || student.code}</td>
          <td>${lastLesson ? formatDateTime(lastLesson.startAt) : "尚未上課"}</td>
          <td>${stats.lessons.length}</td>
          <td>${stats.totalChargedCount}</td>
          <td>${billingCycle.paidThroughCount}</td>
          <td><span class="status ${paymentClass}">${getPaymentStatusLabel(billingCycle.effectivePaymentStatus)}</span></td>
        </tr>
      `;
    }).join("");
    el.studentOverviewTable.innerHTML = `
      <thead><tr><th>學生</th><th>最後上課日期</th><th>課程堂數</th><th>累計扣堂</th><th>已繳到</th><th>繳費狀態</th></tr></thead>
      <tbody>${rows || "<tr><td colspan=\"6\">目前沒有學生資料</td></tr>"}</tbody>
    `;
  }

  function renderChargePanel() {
    if (!el.chargeStudentSelect || !el.chargeMetricsBox || !el.chargeLedgerTable) {
      return;
    }
    const visibleStudents = getVisibleChargeStudents();
    if (!selectedChargeStudentCode || !visibleStudents.some((student) => student.code === selectedChargeStudentCode)) {
      selectedChargeStudentCode = visibleStudents[0]?.code || "";
    }

    el.chargeStudentSelect.innerHTML = visibleStudents
      .map((student) => `<option value="${student.code}" ${student.code === selectedChargeStudentCode ? "selected" : ""}>${student.name || student.code}</option>`)
      .join("");

    if (!selectedChargeStudentCode) {
      el.chargeMetricsBox.innerHTML = "";
      el.chargeLedgerTable.innerHTML = "<thead><tr><th>請先選擇學生</th></tr></thead>";
      if (el.chargeStudentEmailInput) {
        el.chargeStudentEmailInput.value = "";
      }
      if (el.chargeStudentEmailMeta) {
        el.chargeStudentEmailMeta.textContent = "尚未設定學生通知 Email，將改用預設通知信箱。";
      }
      if (el.chargeEmailSendBtn) {
        el.chargeEmailSendBtn.disabled = true;
      }
      if (el.chargeBaseCountInput) {
        el.chargeBaseCountInput.value = "0";
      }
      if (el.chargePaymentStatusSelect) {
        el.chargePaymentStatusSelect.value = "unknown";
      }
      if (el.chargePaymentNoteInput) {
        el.chargePaymentNoteInput.value = "";
      }
      if (el.chargePaymentMeta) {
        el.chargePaymentMeta.textContent = "尚未註記繳費狀態。";
      }
      if (el.chargeReminderSummary) {
        el.chargeReminderSummary.textContent = "第 4 堂提醒尚未觸發。";
      }
      if (el.chargeReminderTable) {
        el.chargeReminderTable.innerHTML = "<thead><tr><th>請先選擇學生</th></tr></thead>";
      }
      return;
    }

    const stats = getStudentChargeStats(selectedChargeStudentCode);
    const student = stats.student;
    if (!student) {
      return;
    }
    const reminderLogs = Array.isArray(student.chargeReminderLogs) ? student.chargeReminderLogs : [];
    const latestSuccessLog = reminderLogs.find((item) => item.status === "success");
    const billingCycle = getStudentBillingCycle(stats);

    if (el.chargeBaseCountInput) {
      el.chargeBaseCountInput.value = String(stats.startCount);
    }
    if (el.chargeStudentEmailInput) {
      el.chargeStudentEmailInput.value = String(student.email || "");
    }
    if (el.chargeEmailSendBtn) {
      el.chargeEmailSendBtn.disabled = false;
      el.chargeEmailSendBtn.textContent = "寄送計費 Email";
    }
    if (el.chargeStudentEmailMeta) {
      const fallback = String(window.APP_CONFIG?.defaultNotifyEmail || "").trim();
      const updatedText = student.emailUpdatedAt
        ? `最後更新：${formatDateTime(student.emailUpdatedAt)} / ${student.emailUpdatedBy || "-"}`
        : "尚未個別設定學生 Email。";
      el.chargeStudentEmailMeta.textContent = student.email
        ? `${updatedText}（目前寄送：${student.email}）`
        : `${updatedText}（目前改用預設：${fallback || "未設定"}）`;
    }
    if (el.chargePaymentStatusSelect) {
      el.chargePaymentStatusSelect.value = billingCycle.effectivePaymentStatus;
    }
    if (el.chargePaymentNoteInput) {
      el.chargePaymentNoteInput.value = String(student.paymentNote || "");
    }
    if (el.chargePaymentMeta) {
      const confirmedText = student.paymentConfirmedAt
        ? `最後確認：${formatDateTime(student.paymentConfirmedAt)} / ${student.paymentConfirmedBy || "-"}`
        : "尚未註記繳費狀態。";
      el.chargePaymentMeta.textContent = `${confirmedText}；已繳到第 ${billingCycle.paidThroughCount} 堂。`;
    }
    if (el.chargeReminderSummary) {
      const latestText = latestSuccessLog
        ? `最近寄送：第 ${latestSuccessLog.milestone} 堂（${formatDateTime(latestSuccessLog.sentAt)}）。`
        : "尚未成功寄送繳費提醒。";
      if (billingCycle.isPaymentDue) {
        el.chargeReminderSummary.textContent = `已達第 ${billingCycle.nextPaymentDueCount} 堂繳費門檻，需確認本期繳費。${latestText}`;
      } else {
        el.chargeReminderSummary.textContent = `已繳到第 ${billingCycle.paidThroughCount} 堂；本期已扣 ${billingCycle.currentCycleChargedCount}/${CHARGE_REMINDER_STEP} 堂，距下次繳費還差 ${billingCycle.remainingToNextPayment} 堂。${latestText}`;
      }
    }

    el.chargeMetricsBox.innerHTML = `
      <div class="metric"><div class="k">系統課程數</div><div class="v">${stats.lessons.length}</div></div>
      <div class="metric"><div class="k">起始堂數</div><div class="v">${stats.startCount}</div></div>
      <div class="metric"><div class="k">已扣堂數（系統）</div><div class="v">${stats.chargedLessons.length}</div></div>
      <div class="metric"><div class="k">累計已扣堂</div><div class="v">${stats.totalChargedCount}</div></div>
      <div class="metric"><div class="k">已繳到堂數</div><div class="v">${billingCycle.paidThroughCount}</div></div>
      <div class="metric"><div class="k">本期已扣</div><div class="v">${billingCycle.currentCycleChargedCount}/${CHARGE_REMINDER_STEP}</div></div>
      <div class="metric"><div class="k">下次應繳門檻</div><div class="v">${billingCycle.nextPaymentDueCount}</div></div>
      <div class="metric"><div class="k">繳費狀態</div><div class="v">${getPaymentStatusLabel(billingCycle.effectivePaymentStatus)}</div></div>
      <div class="metric"><div class="k">未到課次數</div><div class="v">${stats.noShowCount}</div></div>
      <div class="metric"><div class="k">臨時請假次數</div><div class="v">${stats.tempLeaveCount}</div></div>
      <div class="metric"><div class="k">重大急事次數</div><div class="v">${stats.majorCount}</div></div>
      <div class="metric"><div class="k">正常請假次數</div><div class="v">${stats.normalLeaveCount}</div></div>
    `;

    const reminderRows = reminderLogs.map((item) => `
      <tr>
        <td>${item.label || `第 ${item.milestone} 堂`}</td>
        <td>${getChargeReminderStatusPill(item.status)}</td>
        <td>${item.sentAt ? formatDateTime(item.sentAt) : "-"}</td>
        <td>${item.to || "-"}</td>
        <td>${item.note || "-"}</td>
      </tr>
    `).join("");
    if (el.chargeReminderTable) {
      el.chargeReminderTable.innerHTML = `
        <thead><tr><th>提醒門檻</th><th>寄送結果</th><th>時間</th><th>收件人</th><th>說明</th></tr></thead>
        <tbody>${reminderRows || "<tr><td colspan=\"5\">目前沒有提醒紀錄</td></tr>"}</tbody>
      `;
    }

    const ledgerRows = stats.chargedLessons.map((lesson) => `
      <tr>
        <td>${formatDateTime(lesson.startAt)}</td>
        <td>${getStatusPill(lesson.attendanceStatus)}</td>
        <td>${lesson.sourceType === "MAKEUP" ? "補課" : "原課"}</td>
        <td>是</td>
      </tr>
    `).join("");

    el.chargeLedgerTable.innerHTML = `
      <thead><tr><th>時間</th><th>狀態</th><th>類型</th><th>扣堂</th></tr></thead>
      <tbody>${ledgerRows || "<tr><td colspan=\"4\">目前沒有扣堂紀錄</td></tr>"}</tbody>
    `;
  }

  function renderSessionTexts() {
    if (el.studentSessionText) {
      const studentName = getStudentDisplayName(activeStudentCode);
      const coachName = getCoachDisplayName(activeCoachCode || "");
      el.studentSessionText.textContent = activeStudentCode
        ? `已載入學生：${studentName}（教練 ${coachName}）`
        : "尚未載入學生。";
    }
    if (el.coachSessionText) {
      const coachName = getCoachDisplayName(activeCoachCode || "");
      el.coachSessionText.textContent = activeCoachCode
        ? `已載入教練：${coachName}${isCoachReadOnlyMode() ? "（唯讀）" : ""}`
        : "尚未載入教練。";
    }
  }

  function renderAll() {
    expirePendingRequests();
    renderSessionTexts();
    renderStudentLessons();
    renderStudentCalendar();
    renderMakeupSection();
    renderCoachPending();
    renderCoachCalendar();
    renderCoachLeaveTable();
    renderCalendarRemovedPanel();
    renderCompensationPanel();
    renderLogs();
    renderStudentOverviewPanel();
    renderChargePanel();
    updateCoachReadOnlyUi();
  }

  function loadSlotsForSelectedLeave() {
    if (!el.makeupLeaveSelect || !el.makeupSlotSelect || !el.slotWindowHint) {
      return;
    }
    const leaveId = el.makeupLeaveSelect.value;
    if (!leaveId) {
      el.makeupSlotSelect.innerHTML = "<option value=''>先選請假單</option>";
      el.slotWindowHint.textContent = "請先選擇請假單，再載入可補課時段。";
      return;
    }
    const leave = state.leaveRequests.find((item) => item.id === leaveId);
    if (!leave) {
      return;
    }
    const slots = getCycleSlots(leave);
    el.makeupSlotSelect.innerHTML = `<option value="">請選擇時段</option>${slots.map((slot) => `<option value="${slot.startAt}">${formatDateTime(slot.startAt)}</option>`).join("")}`;

    const lesson = getLessonById(leave.lessonId);
    if (lesson) {
      const cycle = getCycleWindowByLesson(lesson);
      el.slotWindowHint.textContent = `補課週期：${cycle.startKey} ~ ${cycle.endKey}，可用時段 ${slots.length} 筆。`;
    }
  }

  function bindEvents() {
    if (el.studentLoginBtn) {
      el.studentLoginBtn.addEventListener("click", async () => {
        await syncCoachflowRosterFromCloud();
        syncCoachflowRosterFromLocalState();
        const loaded = activateStudentSession(el.studentCode?.value, el.studentCoachCode?.value, false);
        if (loaded) {
          try {
            const changed = await syncCloudLeaveRecords({ studentCode: activeStudentCode, coachCode: activeCoachCode });
            if (changed) {
              renderAll();
            }
          } catch (error) {
            console.warn("Cloud leave sync failed for student login:", error);
          }
        }
      });
    }

    if (el.coachLoginBtn) {
      el.coachLoginBtn.addEventListener("click", async () => {
        const coachCodeInput = el.coachCode?.value;
        const readOnlyLogin = isReadOnlyLoginCode(coachCodeInput);
        syncCoachflowRosterFromLocalState();
        const loaded = activateCoachSession(coachCodeInput, false);
        if (loaded && isCoachReadOnlyMode()) {
          setCoachLoginPromptVisibility(true);
        }
        const syncFromCloud = async () => {
          const cloudChanged = await syncCoachflowRosterFromCloud();
          const localChanged = syncCoachflowRosterFromLocalState();
          let leaveChanged = false;
          try {
            leaveChanged = await syncCloudLeaveRecords({ coachCode: activeCoachCode });
          } catch (error) {
            console.warn("Cloud leave sync failed for coach login:", error);
          }
          if (cloudChanged || localChanged || leaveChanged) {
            renderAll();
          }
        };
        if (readOnlyLogin) {
          syncReadOnlyCoachCalendarFromGoogle(activeCoachCode)
            .then((changed) => {
              if (changed) {
                renderAll();
              }
            })
            .catch((error) => {
              console.warn("Read-only Google calendar sync failed:", error);
              setUiStatus("唯讀月曆目前無法同步 Google，請稍後重新整理。");
            });
          syncFromCloud().catch((error) => {
            console.warn("CoachFlow roster sync failed via read-only login:", error);
          });
          return;
        }
        await syncFromCloud();
      });
    }

    if (el.studentResetBtn) {
      el.studentResetBtn.addEventListener("click", () => {
      if (!window.confirm("確定要重置沙盒資料嗎？")) {
        return;
      }
      state = seedState();
      activeStudentCode = "";
      activeCoachCode = "";
      selectedChargeStudentCode = "";
      studentCalendarMonthStart = getMonthStart(makeTaipeiDateTime(getDateKeyInTaipei(new Date()), "12:00"));
      selectedStudentDateKey = getDateKeyInTaipei(new Date());
      selectedStudentLessonId = "";
      coachCalendarMonthStart = getMonthStart(makeTaipeiDateTime(getDateKeyInTaipei(new Date()), "12:00"));
      selectedCoachDateKey = getDateKeyInTaipei(new Date());
      selectedCoachLessonId = "";
      closeStudentDayModal();
      closeCoachDayModal();
      closeCoachReviewModal();
      saveState();
      renderAll();
      });
    }

    if (el.studentLessonsTable) {
      el.studentLessonsTable.addEventListener("click", (event) => {
      const button = closestFromEvent(event, "button[data-action][data-id]");
      if (!button) {
        return;
      }
      if (button.dataset.action === "leave") {
        applyNormalLeave(button.dataset.id);
        return;
      }
      if (button.dataset.action === "cancel-leave") {
        cancelNormalLeaveByStudent(button.dataset.id);
      }
      });
    }

    if (el.studentCalendarPrevBtn) {
      el.studentCalendarPrevBtn.addEventListener("click", () => {
        studentCalendarMonthStart = shiftMonth(studentCalendarMonthStart, -1);
        renderStudentCalendar();
      });
    }
    if (el.studentCalendarNextBtn) {
      el.studentCalendarNextBtn.addEventListener("click", () => {
        studentCalendarMonthStart = shiftMonth(studentCalendarMonthStart, 1);
        renderStudentCalendar();
      });
    }
    if (el.studentCalendarGrid) {
      el.studentCalendarGrid.addEventListener("click", (event) => {
        const target = closestFromEvent(event, "[data-student-cal-date]");
        if (!target) {
          return;
        }
        event.preventDefault();
        event.stopPropagation();
        selectedStudentDateKey = target.dataset.studentCalDate;
        selectedStudentLessonId = target.dataset.studentLessonId || "";
        renderStudentCalendar();
        openStudentDayModal();
        setUiStatus(`學生月曆點擊：${selectedStudentDateKey}${selectedStudentLessonId ? ` / ${selectedStudentLessonId}` : ""}`);
      });
    }
    if (el.studentDayDetail) {
      el.studentDayDetail.addEventListener("click", (event) => {
        const button = closestFromEvent(event, "button[data-action][data-id]");
        if (!button) {
          return;
        }
        selectedStudentLessonId = button.dataset.id || "";
        if (button.dataset.action === "leave") {
          applyNormalLeave(button.dataset.id);
          return;
        }
        if (button.dataset.action === "cancel-leave") {
          cancelNormalLeaveByStudent(button.dataset.id);
        }
      });
    }
    if (el.studentDayModalCloseBtn) {
      el.studentDayModalCloseBtn.addEventListener("click", clearStudentCalendarSelection);
    }
    if (el.studentDayModal) {
      el.studentDayModal.addEventListener("click", (event) => {
        if (event.target === el.studentDayModal) {
          clearStudentCalendarSelection();
        }
      });
    }
    if (el.studentCalendarGrid) {
      document.addEventListener("click", (event) => {
        if (
          closestFromEvent(event, "#student-calendar-grid") ||
          closestFromEvent(event, "#student-calendar-prev-btn") ||
          closestFromEvent(event, "#student-calendar-next-btn") ||
          closestFromEvent(event, "#student-day-modal")
        ) {
          return;
        }
        clearStudentCalendarSelection();
      });
    }

    if (el.coachCalendarPrevBtn) {
      el.coachCalendarPrevBtn.addEventListener("click", () => {
        if (isCoachReadOnlyMode()) {
          const { minMonth } = getReadOnlyMonthBounds();
          if (getMonthStartTime(coachCalendarMonthStart) <= minMonth.getTime()) {
            return;
          }
        }
        coachCalendarMonthStart = shiftMonth(coachCalendarMonthStart, -1);
        clampReadOnlyCoachMonth();
        renderCoachCalendar();
      });
    }
    if (el.coachCalendarNextBtn) {
      el.coachCalendarNextBtn.addEventListener("click", () => {
        if (isCoachReadOnlyMode()) {
          const { maxMonth } = getReadOnlyMonthBounds();
          if (getMonthStartTime(coachCalendarMonthStart) >= maxMonth.getTime()) {
            return;
          }
        }
        coachCalendarMonthStart = shiftMonth(coachCalendarMonthStart, 1);
        clampReadOnlyCoachMonth();
        renderCoachCalendar();
      });
    }
    if (el.coachCalendarSyncBtn) {
      el.coachCalendarSyncBtn.addEventListener("click", () => {
        if (requireCoachWriteAccess()) {
          syncCoachCalendarEvents();
        }
      });
    }
    if (el.calendarRemovedRefreshBtn) {
      el.calendarRemovedRefreshBtn.addEventListener("click", () => {
        renderCalendarRemovedPanel();
      });
    }
    if (el.calendarRemovedRestoreAllBtn) {
      el.calendarRemovedRestoreAllBtn.addEventListener("click", () => {
        if (!requireCoachWriteAccess()) {
          return;
        }
        restoreAllCalendarRemovedLessons();
      });
    }
    if (el.calendarRemovedToggleBtn) {
      el.calendarRemovedToggleBtn.addEventListener("click", () => {
        isCalendarRemovedExpanded = !isCalendarRemovedExpanded;
        renderCalendarRemovedPanel();
      });
    }
    if (el.calendarRemovedTable) {
      el.calendarRemovedTable.addEventListener("click", (event) => {
        const button = closestFromEvent(event, "button[data-action='restore-removed'][data-id]");
        if (!button) {
          return;
        }
        if (!requireCoachWriteAccess()) {
          return;
        }
        restoreCalendarRemovedLesson(button.dataset.id);
      });
    }
    if (el.coachCalendarGrid) {
      el.coachCalendarGrid.addEventListener("click", (event) => {
        if (isCoachReadOnlyMode()) {
          return;
        }
        const target = closestFromEvent(event, "[data-coach-cal-date]");
        if (!target) {
          return;
        }
        event.preventDefault();
        event.stopPropagation();
        selectedCoachDateKey = target.dataset.coachCalDate;
        selectedCoachLessonId = target.dataset.coachLessonId || "";
        renderCoachCalendar();
        openCoachDayModal();
        setUiStatus(`教練月曆點擊：${selectedCoachDateKey}${selectedCoachLessonId ? ` / ${selectedCoachLessonId}` : ""}`);
      });
    }
    if (el.coachDayDetail) {
      el.coachDayDetail.addEventListener("click", (event) => {
        const lessonButton = closestFromEvent(event, "button[data-coach-lesson-id][data-coach-mark]");
        if (lessonButton) {
          if (!requireCoachWriteAccess()) {
            return;
          }
          const lessonId = lessonButton.dataset.coachLessonId;
          const mark = lessonButton.dataset.coachMark;
          selectedCoachLessonId = lessonId;
          if (mark === "cancel-leave") {
            revokeNormalLeave(lessonId);
            return;
          }
          if (mark === "reschedule") {
            rescheduleLessonByCoach(lessonId);
            return;
          }
          markLessonStatus(lessonId, mark);
          return;
        }

        const reviewButton = closestFromEvent(event, "button[data-day-review-id][data-day-review-action]");
        if (!reviewButton) {
          return;
        }
        if (!requireCoachWriteAccess()) {
          return;
        }
        openCoachReviewModal(reviewButton.dataset.dayReviewId, reviewButton.dataset.dayReviewAction);
      });
    }
    if (el.coachDayModalCloseBtn) {
      el.coachDayModalCloseBtn.addEventListener("click", clearCoachCalendarSelection);
    }
    if (el.coachDayModal) {
      el.coachDayModal.addEventListener("click", (event) => {
        if (event.target === el.coachDayModal) {
          clearCoachCalendarSelection();
        }
      });
    }
    if (el.coachCalendarGrid) {
      document.addEventListener("click", (event) => {
        if (
          closestFromEvent(event, "#coach-calendar-grid") ||
          closestFromEvent(event, "#coach-calendar-prev-btn") ||
          closestFromEvent(event, "#coach-calendar-next-btn") ||
          closestFromEvent(event, "#coach-day-modal") ||
          closestFromEvent(event, "#coach-review-modal")
        ) {
          return;
        }
        clearCoachCalendarSelection();
      });
    }

    if (el.loadSlotsBtn) {
      el.loadSlotsBtn.addEventListener("click", loadSlotsForSelectedLeave);
    }
    if (el.makeupLeaveSelect) {
      el.makeupLeaveSelect.addEventListener("change", loadSlotsForSelectedLeave);
    }
    if (el.submitMakeupBtn) {
      el.submitMakeupBtn.addEventListener("click", submitMakeupRequest);
    }

    if (el.studentMakeupTable) {
      el.studentMakeupTable.addEventListener("click", (event) => {
      const button = closestFromEvent(event, "button[data-action='cancel-pending']");
      if (!button) {
        return;
      }
      cancelPending(button.dataset.id);
      });
    }

    if (el.coachReviewModalConfirmBtn) {
      el.coachReviewModalConfirmBtn.addEventListener("click", () => {
        if (!requireCoachWriteAccess()) {
          return;
        }
        const requestId = coachReviewPendingRequestId;
        const action = coachReviewPendingAction;
        closeCoachReviewModal();
        if (!requestId || !action) {
          return;
        }
        if (action === "approve") {
          approveRequest(requestId);
          return;
        }
        const inputReason = String(el.coachReviewReasonInput?.value || "").trim();
        const rejectReason = inputReason || coachRejectReasonPreset || "教練退回";
        coachRejectReasonPreset = rejectReason;
        rejectRequest(requestId, rejectReason);
      });
    }
    if (el.coachReviewModalCancelBtn) {
      el.coachReviewModalCancelBtn.addEventListener("click", closeCoachReviewModal);
    }
    if (el.coachReviewModalCloseBtn) {
      el.coachReviewModalCloseBtn.addEventListener("click", closeCoachReviewModal);
    }
    if (el.coachReviewModal) {
      el.coachReviewModal.addEventListener("click", (event) => {
        if (event.target === el.coachReviewModal) {
          closeCoachReviewModal();
        }
      });
    }

    if (el.coachLeaveAddBtn) {
      el.coachLeaveAddBtn.addEventListener("click", () => {
        if (requireCoachWriteAccess()) {
          addCoachLeaveBlock();
        }
      });
    }
    if (el.coachLeaveTable) {
      el.coachLeaveTable.addEventListener("click", (event) => {
        const button = closestFromEvent(event, "button[data-action='remove-coach-block']");
        if (!button) {
          return;
        }
        if (!requireCoachWriteAccess()) {
          return;
        }
        removeCoachLeaveBlock(button.dataset.id);
      });
    }

    if (el.coachReviewStudentFilter) {
      el.coachReviewStudentFilter.addEventListener("change", () => {
        coachReviewFilterStudent = el.coachReviewStudentFilter.value || "ALL";
        renderCoachPending();
      });
    }
    if (el.coachReviewRefreshBtn) {
      el.coachReviewRefreshBtn.addEventListener("click", () => {
        renderAll();
      });
    }

    if (el.compensationRefreshBtn) {
      el.compensationRefreshBtn.addEventListener("click", () => {
        renderCompensationPanel();
      });
    }
    if (el.compensationRetryAllBtn) {
      el.compensationRetryAllBtn.addEventListener("click", () => {
        if (!requireCoachWriteAccess()) {
          return;
        }
        retryAllCompensationTasks();
      });
    }
    if (el.compensationTable) {
      el.compensationTable.addEventListener("click", (event) => {
        const button = closestFromEvent(event, "button[data-action='retry-comp'], button[data-action='delete-comp']");
        if (!button) {
          return;
        }
        if (!requireCoachWriteAccess()) {
          return;
        }
        if (button.dataset.action === "delete-comp") {
          deleteCompensationTask(button.dataset.id);
          return;
        }
        retryCompensationTask(button.dataset.id);
      });
    }

    if (el.chargeStudentSelect) {
      el.chargeStudentSelect.addEventListener("change", () => {
        selectedChargeStudentCode = el.chargeStudentSelect.value;
        renderChargePanel();
      });
    }
    if (el.chargeStudentEmailSaveBtn) {
      el.chargeStudentEmailSaveBtn.addEventListener("click", () => {
        if (requireCoachWriteAccess()) {
          saveSelectedStudentNotifyEmail();
        }
      });
    }
    if (el.chargeEmailSendBtn) {
      el.chargeEmailSendBtn.addEventListener("click", () => {
        if (!requireCoachWriteAccess()) {
          return;
        }
        sendSelectedStudentChargeEmail().catch((error) => {
          console.error("manual billing email failed:", error);
        });
      });
    }
    if (el.chargeBaseCountSaveBtn) {
      el.chargeBaseCountSaveBtn.addEventListener("click", () => {
        if (requireCoachWriteAccess()) {
          saveChargeStartCount();
        }
      });
    }
    if (el.chargePaymentSaveBtn) {
      el.chargePaymentSaveBtn.addEventListener("click", () => {
        if (requireCoachWriteAccess()) {
          saveStudentPaymentStatus().catch((error) => {
            console.error("payment confirmation failed:", error);
          });
        }
      });
    }
  }

  function readJsonFromStorage(storageKey) {
    try {
      const raw = localStorage.getItem(storageKey);
      return raw ? JSON.parse(raw) : null;
    } catch (error) {
      console.warn(`Failed to parse localStorage value for ${storageKey}:`, error);
      return null;
    }
  }

  function setCoachLoginPromptVisibility(hidden) {
    const shouldHide = Boolean(hidden);
    coachLoginPromptHidden = shouldHide;
    const coachLoginCard = el.coachCode?.closest("section");
    if (coachLoginCard) {
      coachLoginCard.hidden = shouldHide;
    }
    if (el.coachCode) {
      const coachCodeLabel = el.coachCode.closest("label");
      if (coachCodeLabel) {
        coachCodeLabel.hidden = shouldHide;
      }
      el.coachCode.readOnly = shouldHide;
    }
    if (el.coachLoginBtn) {
      el.coachLoginBtn.hidden = shouldHide;
      el.coachLoginBtn.disabled = shouldHide;
    }
  }

  function setStudentLoginPromptVisibility(hidden) {
    const shouldHide = Boolean(hidden);
    studentLoginPromptHidden = shouldHide;
    const studentLoginCard = el.studentCode?.closest("section");
    if (studentLoginCard) {
      studentLoginCard.hidden = shouldHide;
    }
    if (el.studentCode) {
      el.studentCode.readOnly = shouldHide;
    }
    if (el.studentCoachCode) {
      el.studentCoachCode.readOnly = shouldHide;
    }
    if (el.studentLoginBtn) {
      el.studentLoginBtn.hidden = shouldHide;
      el.studentLoginBtn.disabled = shouldHide;
    }
    if (el.studentResetBtn) {
      el.studentResetBtn.hidden = true;
      el.studentResetBtn.disabled = true;
    }
  }

  function emitLeavePrefillDebug(stage, payload = {}) {
    const params = new URLSearchParams(window.location.search);
    const snapshot = {
      stage: String(stage || ""),
      urlParams: {
        coachCode: String(params.get("coachCode") || ""),
        studentCode: String(params.get("studentCode") || ""),
        code: String(params.get("code") || "")
      },
      storedLeavePrefill: readJsonFromStorage(LEAVE_PREFILL_STORAGE_KEY),
      coachflowSession: readJsonFromStorage("coachflow-v2-session"),
      coachflowCurrentCoachId: String(readJsonFromStorage("coachflow-v2-state")?.currentCoachId || ""),
      mergedStudentCode: String(payload.mergedStudentCode || ""),
      mergedCoachCode: String(payload.mergedCoachCode || ""),
      stateCoaches: Array.isArray(state.coaches) ? state.coaches.map((coach) => ({ ...coach })) : [],
      activateStudentSessionResult: payload.activateStudentSessionResult,
      activateCoachSessionResult: payload.activateCoachSessionResult
    };
    window.__coachflowLeaveDebug = snapshot;
    console.log("[CoachFlowLeaveDebug]", snapshot);
  }

  function applySessionPrefillFromUrl() {
    const readCoachflowSessionPrefill = () => {
      try {
        const sessionRaw = localStorage.getItem("coachflow-v2-session");
        const stateRaw = localStorage.getItem("coachflow-v2-state");
        const session = sessionRaw ? JSON.parse(sessionRaw) : {};
        const coachflowState = stateRaw ? JSON.parse(stateRaw) : {};
        const coaches = Array.isArray(coachflowState?.coaches) ? coachflowState.coaches : [];
        const students = Array.isArray(coachflowState?.students) ? coachflowState.students : [];
        const currentCoachId = String(coachflowState?.currentCoachId || "").trim();
        const activeCoach = coaches.find((coach) => String(coach?.id || "").trim() === currentCoachId) || null;
        const coachCode = normalizeParticipantCode(
          session?.authenticatedCoachAccess ||
          activeCoach?.accessCode ||
          activeCoach?.access_code ||
          activeCoach?.code ||
          coaches[0]?.accessCode ||
          coaches[0]?.access_code ||
          coaches[0]?.code ||
          ""
        );
        const studentCode = normalizeParticipantCode(
          session?.currentStudentAccess ||
          students.find((student) => String(student?.id || "").trim() === String(session?.currentStudentId || "").trim())?.accessCode ||
          students.find((student) => String(student?.id || "").trim() === String(session?.currentStudentId || "").trim())?.access_code ||
          ""
        );
        return {
          coachCode,
          studentCode,
          from: "coachflow-session"
        };
      } catch (error) {
        console.warn("Failed to read CoachFlow session prefill:", error);
        return null;
      }
    };

    const readLeavePrefillSession = () => {
      try {
        const raw = localStorage.getItem(LEAVE_PREFILL_STORAGE_KEY);
        if (!raw) {
          return null;
        }
        const parsed = JSON.parse(raw);
        if (!parsed || typeof parsed !== "object") {
          return null;
        }
        const createdAt = Number(parsed.createdAt || 0);
        if (!Number.isFinite(createdAt) || Date.now() - createdAt > LEAVE_PREFILL_MAX_AGE_MS) {
          return null;
        }
        return {
          coachCode: normalizeParticipantCode(parsed.coachCode || ""),
          studentCode: normalizeParticipantCode(parsed.studentCode || ""),
          from: String(parsed.from || "").trim()
        };
      } catch (error) {
        console.warn("Failed to read leave prefill session:", error);
        return null;
      }
    };

    const params = new URLSearchParams(window.location.search);
    const pagePath = String(window.location.pathname || "").toLowerCase();
    const isStudentPage = pagePath.includes("leave-student");
    let studentCode = String(
      params.get("studentCode") ||
      params.get("student") ||
      params.get("stu") ||
      params.get("studentCodeInput") ||
      params.get("student_access_code") ||
      params.get("studentAccessCode") ||
      ""
    ).trim().toUpperCase();
    if (!studentCode && isStudentPage) {
      studentCode = String(params.get("code") || "").trim().toUpperCase();
    }
    let coachCode = String(
      params.get("coachCode") ||
      params.get("coach") ||
      params.get("coach_code") ||
      params.get("coach_access_code") ||
      params.get("coachAccessCode") ||
      ""
    ).trim().toUpperCase();
    if (!coachCode && !isStudentPage) {
      coachCode = String(params.get("code") || params.get("token") || "").trim().toUpperCase();
    }
    const from = String(params.get("from") || "").trim();
    const autoLoginRequested = ["1", "true", "yes", "y", "on"].includes(
      String(params.get("autoLogin") || params.get("autologin") || "").trim().toLowerCase()
    );
    const storedPrefill = readLeavePrefillSession();
    const coachflowPrefill = readCoachflowSessionPrefill();
    const hasUrlStudentCode = Boolean(studentCode);
    const hasUrlCoachCode = Boolean(coachCode);
    const hasStoredStudentPrefill = Boolean(storedPrefill?.studentCode);
    const hasStoredCoachPrefill = Boolean(storedPrefill?.coachCode);
    const coachflowSessionStudentCode = isStudentPage ? "" : coachflowPrefill?.studentCode;
    const mergedStudentCode = normalizeParticipantCode(
      studentCode ||
      storedPrefill?.studentCode ||
      coachflowSessionStudentCode ||
      ""
    );
    const mergedCoachCode = normalizeParticipantCode(
      coachCode ||
      storedPrefill?.coachCode ||
      coachflowPrefill?.coachCode ||
      ""
    );
    const mergedFrom = String(from || storedPrefill?.from || coachflowPrefill?.from || "").trim();
    const hasFallbackPrefill = Boolean(
      (storedPrefill && (storedPrefill.coachCode || storedPrefill.studentCode)) ||
      (coachflowPrefill && (coachflowPrefill.coachCode || coachflowPrefill.studentCode))
    );
    const hasActionablePrefill = isStudentPage
      ? Boolean(hasUrlStudentCode || hasStoredStudentPrefill || (autoLoginRequested && (hasUrlStudentCode || hasStoredStudentPrefill)))
      : Boolean(hasUrlCoachCode || hasStoredCoachPrefill || (autoLoginRequested && (hasUrlCoachCode || hasStoredCoachPrefill)));

    if (mergedStudentCode && el.studentCode) {
      el.studentCode.value = mergedStudentCode;
    }
    if (mergedCoachCode && el.studentCoachCode) {
      el.studentCoachCode.value = mergedCoachCode;
    }
    if (mergedCoachCode && el.coachCode) {
      el.coachCode.value = mergedCoachCode;
    }
    emitLeavePrefillDebug("applySessionPrefillFromUrl", {
      mergedStudentCode,
      mergedCoachCode
    });
    return {
      studentCode: mergedStudentCode,
      coachCode: mergedCoachCode,
      from: mergedFrom,
      autoLoginRequested: (autoLoginRequested || hasFallbackPrefill) && hasActionablePrefill,
      hasActionablePrefill
    };
  }

  async function bootstrap() {
    ensureMakeupCodes();
    ensureLessonCalendarEventIds();
    ensureParticipantEmails();
    ensureStudentBillingProfiles();
    syncCoachflowRosterFromLocalState();

    const cloudSyncPromise = (async () => {
      const syncedFromCloud = await syncCoachflowRosterFromCloud();
      const syncedFromLocal = syncCoachflowRosterFromLocalState();
      return Boolean(syncedFromCloud || syncedFromLocal);
    })();

    const sessionPrefill = applySessionPrefillFromUrl();
    const hasPrefillValue = Boolean(sessionPrefill.studentCode || sessionPrefill.coachCode);
    const shouldAutoLoginFromPrefill = Boolean(
      sessionPrefill.autoLoginRequested &&
      sessionPrefill.hasActionablePrefill
    );
    const shouldForceProfileFromUrl = hasPrefillValue &&
      shouldAutoLoginFromPrefill &&
      String(sessionPrefill.from || "").toLowerCase().startsWith("coachflow");
    if (shouldForceProfileFromUrl) {
      const fallbackCoachCode = normalizeParticipantCode(
        sessionPrefill.coachCode ||
        getStudentByCode(sessionPrefill.studentCode)?.coachCode ||
        state.coaches[0]?.code ||
        "CH001"
      );
      if (fallbackCoachCode) {
        ensureCoachProfile(fallbackCoachCode);
      }
      if (sessionPrefill.studentCode) {
        ensureStudentProfile(sessionPrefill.studentCode, fallbackCoachCode || "CH001", { silentMode: true });
      }
      if (sessionPrefill.coachCode) {
        ensureCoachProfile(sessionPrefill.coachCode);
      }
      ensureLessonCalendarEventIds();
      ensureParticipantEmails();
      ensureStudentBillingProfiles();
      saveState();
    }
    bindEvents();
    saveState();

    let autoStudentLoaded = false;
    let autoCoachLoaded = false;

    if (shouldAutoLoginFromPrefill) {
      if (sessionPrefill.studentCode) {
        autoStudentLoaded = activateStudentSession(sessionPrefill.studentCode, sessionPrefill.coachCode, true);
      }
      if (!autoStudentLoaded && sessionPrefill.coachCode) {
        autoCoachLoaded = activateCoachSession(sessionPrefill.coachCode, true);
      }
      if (!autoStudentLoaded && !autoCoachLoaded) {
        await new Promise((resolve) => setTimeout(resolve, 250));
        if (sessionPrefill.studentCode) {
          autoStudentLoaded = activateStudentSession(sessionPrefill.studentCode, sessionPrefill.coachCode, true);
        }
        if (!autoStudentLoaded && sessionPrefill.coachCode) {
          autoCoachLoaded = activateCoachSession(sessionPrefill.coachCode, true);
        }
      }
    } else if (!hasPrefillValue) {
      autoStudentLoaded = el.studentCode && el.studentCoachCode
        ? activateStudentSession(el.studentCode.value, el.studentCoachCode.value, true)
        : false;
      autoCoachLoaded = !autoStudentLoaded && el.coachCode
        ? activateCoachSession(el.coachCode.value, true)
        : false;
    }

    const shouldHideCoachLoginPrompt = Boolean(autoCoachLoaded);
    const shouldHideStudentLoginPrompt = Boolean(autoStudentLoaded);
    setStudentLoginPromptVisibility(shouldHideStudentLoginPrompt);
    setCoachLoginPromptVisibility(shouldHideCoachLoginPrompt);
    emitLeavePrefillDebug("bootstrap-auto-login", {
      mergedStudentCode: sessionPrefill.studentCode,
      mergedCoachCode: sessionPrefill.coachCode,
      activateStudentSessionResult: autoStudentLoaded,
      activateCoachSessionResult: autoCoachLoaded
    });

    if (!autoStudentLoaded && !autoCoachLoaded) {
      renderAll();
    }

    try {
      const cloudSyncChanged = await cloudSyncPromise;
      let studentCalendarSynced = false;
      if (autoStudentLoaded && activeStudentCode) {
        studentCalendarSynced = await syncStudentCalendarEventsFromGoogle();
      }
      let cloudLeaveSynced = false;
      if ((autoStudentLoaded || autoCoachLoaded) && (activeStudentCode || activeCoachCode)) {
        try {
          cloudLeaveSynced = await syncCloudLeaveRecords({
            studentCode: activeStudentCode,
            coachCode: activeCoachCode
          });
        } catch (error) {
          console.warn("Cloud leave sync failed during bootstrap:", error);
        }
      }
      let readOnlyCalendarSynced = false;
      if (autoCoachLoaded && activeCoachReadOnly && activeCoachCode) {
        readOnlyCalendarSynced = await syncReadOnlyCoachCalendarFromGoogle(activeCoachCode);
      }
      if (cloudSyncChanged) {
        if (studentLoginPromptHidden && el.studentCode && activeStudentCode && el.studentCode.value !== activeStudentCode) {
          el.studentCode.value = activeStudentCode;
        }
        if (coachLoginPromptHidden && el.coachCode && activeCoachCode && el.coachCode.value !== activeCoachCode) {
          el.coachCode.value = activeCoachCode;
        }
        renderAll();
      }
      if (studentCalendarSynced && !cloudSyncChanged) {
        renderAll();
      }
      if (cloudLeaveSynced && !cloudSyncChanged && !studentCalendarSynced) {
        renderAll();
      }
      if (readOnlyCalendarSynced && !cloudSyncChanged) {
        renderAll();
      }
    } catch (error) {
      console.warn("CoachFlow cloud sync completed with error:", error);
    }
  }

  bootstrap();
})();
