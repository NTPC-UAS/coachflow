(function () {
  "use strict";

  const STORAGE_KEY = "coachflow-leave-sandbox-v1";
  const SCHEMA_VERSION = 4;
  const TZ = "Asia/Taipei";
  const PENDING_EXPIRE_HOURS = 24;
  const DEFAULT_REJECT_REASON = "時段已被其他學生選走。";
  const CHARGE_REMINDER_STEP = 4;
  const MAX_CHARGE_REMINDER_LOGS = 24;
  const CLOUD_BILLING_RECORD_TYPE = "billing-profile";
  const CLOUD_BILLING_RECORD_PREFIX = "billing-profile:";
  const CLOUD_BILLING_REFRESH_INTERVAL_MS = 45 * 1000;
  const LOCAL_LEAVE_RECONCILE_GRACE_MS = 5 * 60 * 1000;
  const LEAVE_ACTION_DEDUPE_MS = 4000; // 防雙擊：4 秒內對同 lesson 只受理一次請假/取消
  const leaveActionInflight = new Map(); // lessonId -> timestamp
  const CALENDAR_REMOVED_PREVIEW_COUNT = 5;
  const RESET_IMPORTED_CHARGED_COUNT_MIGRATION = "reset-imported-charged-count-20260502-0007";
  const RESET_ALL_TRACKING_DATA_MIGRATION = "reset-all-tracking-data-20260503-0908";
  const PURGE_PRE_RESET_TRACKING_DATA_MIGRATION = "purge-pre-reset-tracking-data-20260503-0009";
  const RESTORE_AUTO_REMOVED_LOCAL_LESSONS_MIGRATION = "restore-auto-removed-local-lessons-20260503-0015";
  const TRACKING_DATA_RESET_AT = "2026-05-03T09:08:00+08:00";
  const BILLING_TRACKING_START_AT = "2026-05-03T00:00:00+08:00";
  const LEAVE_PREFILL_STORAGE_KEY = "coachflow-leave-prefill";
  const LEAVE_PREFILL_MAX_AGE_MS = 10 * 60 * 1000;
  const READ_ONLY_ACCOUNT_CODE = "READONLY";
  const READ_ONLY_DEFAULT_COACH_CODE = "MO001";
  const READ_ONLY_DEFAULT_COACH_NAME = "Monster Chang";
  const READ_ONLY_MONTH_LOOKAHEAD = 1;
  const DEFAULT_LEAVE_APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbxJV_y__nD9nD-6AepZFENxrkQbJv_h3cq59pWaCk2pudGaY5ew5v5WJ_N9zDaj_B7WVg/exec";
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
    coachStudentLeaveTable: document.getElementById("coach-student-leave-table"),
    coachStudentLeaveRefreshBtn: document.getElementById("coach-student-leave-refresh-btn"),
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
    studentOverviewTable: document.getElementById("student-overview-table"),
    coachCloudUploadBtn: document.getElementById("coach-cloud-upload-btn"),
    coachCloudDownloadBtn: document.getElementById("coach-cloud-download-btn"),
    coachCloudSyncText: document.getElementById("coach-cloud-sync-text"),
    pendingLeaveBanner: document.getElementById("pending-leave-banner"),
    pendingLeaveText: document.getElementById("pending-leave-text"),
    pendingLeaveRetryBtn: document.getElementById("pending-leave-retry-btn"),
    manualLessonStudent: document.getElementById("manual-lesson-student"),
    manualLessonDate: document.getElementById("manual-lesson-date"),
    manualLessonTime: document.getElementById("manual-lesson-time"),
    manualLessonAlreadyAttended: document.getElementById("manual-lesson-already-attended"),
    manualLessonAddBtn: document.getElementById("manual-lesson-add-btn")
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
  let lastCloudBillingRefreshAt = 0;
  let lastCloudSessionRefreshAt = 0;
  let cloudSnapshotUploadTimer = 0;
  let isApplyingCloudSnapshot = false;
  let isUploadingCloudSnapshot = false;
  let leaveStateSnapshotUnsupported = false;
  let isSubmittingMakeupRequest = false;
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
      return sanitizeLoadedState(parsed, true);
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
        chargeStartCountUpdatedAt: "",
        chargeStartCountUpdatedBy: "",
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
        chargeStartCountUpdatedAt: "",
        chargeStartCountUpdatedBy: "",
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
        chargeStartCountUpdatedAt: "",
        chargeStartCountUpdatedBy: "",
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

  function isKnownFakeStudentName(value) {
    const s = String(value === null || value === undefined ? "" : value).trim();
    if (!s) {
      return false;
    }
    // 讀取端也做最後一道防線：即使正式後端曾被測試流程污染，主系統名冊、
    // billing profile 或 snapshot 都不能再把 WF / WKND / TEST 學生建回本機。
    // 不使用「學生」單獨判斷，避免擋到缺姓名時的正常 `${code} 學生` fallback。
    return /grader|workflow|測試|週末|流程|學員|student|weekend|wf\d*|wflow|grd|grade|sync|eval|wknd|wkend|test|qa/i.test(s);
  }

  function isKnownFakeStudentRecord(student) {
    return isKnownFakeStudentName(student?.name || student?.studentName);
  }

  function sanitizeLoadedState(candidateState, persist = false) {
    const nextState = candidateState && typeof candidateState === "object" ? candidateState : seedState();
    const students = Array.isArray(nextState.students) ? nextState.students : [];
    const fakeStudents = students.filter(isKnownFakeStudentRecord);
    if (!fakeStudents.length) {
      return nextState;
    }

    const fakeCodes = new Set(fakeStudents.map((student) => normalizeParticipantCode(student.code || student.studentCode)).filter(Boolean));
    const realCodes = new Set(
      students
        .filter((student) => !isKnownFakeStudentRecord(student))
        .map((student) => normalizeParticipantCode(student.code || student.studentCode))
        .filter(Boolean)
    );
    const removableCodes = new Set([...fakeCodes].filter((code) => !realCodes.has(code)));

    nextState.students = students.filter((student) => !isKnownFakeStudentRecord(student));
    if (removableCodes.size) {
      nextState.lessons = (nextState.lessons || []).filter((lesson) => !removableCodes.has(normalizeParticipantCode(lesson.studentCode)));
      nextState.leaveRequests = (nextState.leaveRequests || []).filter((leave) => !removableCodes.has(normalizeParticipantCode(leave.studentCode)));
      nextState.makeupRequests = (nextState.makeupRequests || []).filter((request) => !removableCodes.has(normalizeParticipantCode(request.studentCode)));
      nextState.compensationTasks = (nextState.compensationTasks || []).filter((task) => {
        const payload = task?.payload || {};
        return !removableCodes.has(normalizeParticipantCode(payload.studentCode));
      });
    }

    if (persist) {
      try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(nextState));
      } catch (error) {
        console.warn("Failed to persist sanitized leave state:", error);
      }
    }
    return nextState;
  }

  function saveState() {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    queueCloudStateSnapshotUpload("local-save");
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

  function isAfterTrackingDataReset(dateValue) {
    const resetTime = new Date(TRACKING_DATA_RESET_AT).getTime();
    const valueTime = new Date(dateValue || "").getTime();
    return Number.isFinite(valueTime) && valueTime >= resetTime;
  }

  function hasValidDateValue(dateValue) {
    return Number.isFinite(new Date(dateValue || "").getTime());
  }

  function isAfterBillingTrackingStart(dateValue) {
    const startTime = new Date(BILLING_TRACKING_START_AT).getTime();
    const valueTime = new Date(dateValue || "").getTime();
    return Number.isFinite(valueTime) && valueTime >= startTime;
  }

  function isLessonAfterTrackingDataReset(lesson) {
    return hasValidDateValue(lesson?.startAt);
  }

  function isLessonAfterBillingTrackingStart(lesson) {
    return isAfterBillingTrackingStart(lesson?.startAt);
  }

  function isLessonActiveForTrackingStats(lesson) {
    if (!isLessonAfterBillingTrackingStart(lesson)) {
      return false;
    }
    const startTime = new Date(lesson?.startAt || "").getTime();
    return Number.isFinite(startTime) && (startTime <= Date.now() || lesson.attendanceStatus !== "scheduled");
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

  function getTodayKey() {
    return getDateKeyInTaipei(new Date());
  }

  function getTodayNoon() {
    return makeTaipeiDateTime(getTodayKey(), "12:00");
  }

  function focusStudentCalendarOnToday() {
    const todayKey = getTodayKey();
    studentCalendarMonthStart = getMonthStart(getTodayNoon());
    selectedStudentDateKey = todayKey;
    selectedStudentLessonId = "";
  }

  function focusCoachCalendarOnToday() {
    const todayKey = getTodayKey();
    coachCalendarMonthStart = getMonthStart(getTodayNoon());
    selectedCoachDateKey = todayKey;
    selectedCoachLessonId = "";
    clampReadOnlyCoachMonth();
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
          isLessonAfterTrackingDataReset(lesson) &&
          lesson.attendanceStatus !== "coach-leave" &&
          lesson.attendanceStatus !== "calendar-removed" &&
          getDateKeyInTaipei(new Date(lesson.startAt)) === dateKey
      )
      .sort((a, b) => new Date(a.startAt) - new Date(b.startAt));
  }

  // 教練月曆要看到該日的所有課程，包含學生「正常請假」(leave-normal) 與「教練
  // 停課」(coach-leave) 狀態的課。原本只列 calendarOccupied=true，會把這兩種
  // 整個藏掉 → 教練從月曆完全看不到該學生今天有請假/教練停課、也找不到入口
  // 去處理或核對。
  // calendar-removed / temporary-leave / no-show / major-case 仍藏起來（這些
  // 另有 calendarRemovedPanel / 計費面板可看）。
  function getCoachLessonsForDate(coachCode, dateKey) {
    return state.lessons
      .filter(
        (lesson) =>
          lesson.coachCode === coachCode &&
          isLessonAfterTrackingDataReset(lesson) &&
          (
            lesson.calendarOccupied ||
            lesson.attendanceStatus === "leave-normal" ||
            lesson.attendanceStatus === "coach-leave"
          ) &&
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

  function hasStudentMakeupRequestsForCloudSync(studentCode) {
    const normalizedStudentCode = normalizeParticipantCode(studentCode || "");
    return Boolean(
      normalizedStudentCode &&
      state.makeupRequests.some((request) => (
        normalizeParticipantCode(request.studentCode) === normalizedStudentCode &&
        request.status === "pending"
      ))
    );
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
      el.uiStatusText.hidden = !String(message || "").trim();
    }
  }

  function notifyUser(message, tone = "info") {
    if (!message) {
      return;
    }
    setUiStatus(message);
    if (el.uiStatusText) {
      el.uiStatusText.dataset.tone = tone;
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
      el.coachStudentLeaveTable,
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
    focusStudentCalendarOnToday();
    closeStudentDayModal();
    closeCoachDayModal();
    closeCoachReviewModal();
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
    activeStudentCode = "";
    activeCoachCode = coachCode;
    activeCoachReadOnly = isReadOnlyLogin || isReadonlyParamEnabled() || isReadonlyCoachRole(coach);
    focusCoachCalendarOnToday();
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
    return String(
      window.APP_CONFIG?.leaveAppsScriptUrl ||
      window.APP_CONFIG?.leaveBridgeAppsScriptUrl ||
      window.APP_CONFIG?.leaveSandbox?.appsScriptUrl ||
      DEFAULT_LEAVE_APPS_SCRIPT_URL ||
      ""
    ).trim();
  }

  function isLeaveBridgeEndpointError(message) {
    return /SpreadsheetApp\.getActiveSpreadsheet|spreadsheets\.currently|spreadsheets/i.test(String(message || ""));
  }

  function getCoachflowAppsScriptUrl() {
    return String(window.APP_CONFIG?.coachflowAppsScriptUrl || "").trim();
  }

  // 會搶後端 ScriptLock 的 Sheet 寫入 action。這些要走「單機排隊 + 重試」：
  // 頁面一打開常有三股寫入同時跑（日曆同步尾段 push 課表、billing 快照補種、
  // 手動上傳），全部並發打後端會互搶鎖 → LOCK_TIMEOUT。排隊後同一台裝置
  // 一次只送一個寫入，鎖競爭幾乎消失；剩餘的（跨裝置）由重試吸收。
  const CLOUD_WRITE_ACTIONS = ["saveLessons", "saveLeaveRecord", "saveBillingProfile"];
  let cloudWriteChain = Promise.resolve();

  function isRetryableCloudWriteError(error) {
    // LOCK_TIMEOUT：後端搶鎖逾時。逾時：client 12s abort，但後端可能仍在排隊
    // 完成寫入——三個寫入 action 都是 upsert 冪等，重送安全。
    return /LOCK_TIMEOUT|逾時/i.test(String(error?.message || error || ""));
  }

  async function callAppsScriptApi(action, payload = {}, method = "POST") {
    if (!CLOUD_WRITE_ACTIONS.includes(action)) {
      return callAppsScriptApiRaw(action, payload, method);
    }
    const run = async () => {
      const delays = [0, 2000, 5000, 10000];
      let lastError;
      for (const delay of delays) {
        if (delay) {
          await new Promise((resolve) => setTimeout(resolve, delay + Math.floor(Math.random() * 500)));
        }
        try {
          return await callAppsScriptApiRaw(action, payload, method);
        } catch (error) {
          lastError = error;
          if (!isRetryableCloudWriteError(error)) {
            throw error;
          }
        }
      }
      throw lastError;
    };
    // 排隊：前一個寫入無論成敗都接著跑下一個（錯誤由各自呼叫端處理）
    const result = cloudWriteChain.then(run, run);
    cloudWriteChain = result.catch(() => {});
    return result;
  }

  async function callAppsScriptApiRaw(action, payload = {}, method = "POST") {
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
      const requestPayload = {
        ...buildAppsScriptAuthPayload(),
        ...(payload || {})
      };
      let response;
      if (String(method).toUpperCase() === "GET") {
        endpoint.searchParams.set("action", action);
        endpoint.searchParams.set("_ts", String(Date.now()));
        Object.entries(requestPayload).forEach(([key, value]) => {
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
          body: JSON.stringify({ action, ...requestPayload, _ts: Date.now() }),
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

  function buildAppsScriptAuthPayload() {
    const auth = {};
    if (activeCoachCode) {
      auth.actorCoachAccess = activeCoachCode;
    }
    if (activeStudentCode) {
      auth.actorStudentAccess = activeStudentCode;
    }
    return auth;
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
      const requestPayload = {
        ...buildCoachflowAuthPayload(),
        ...(payload || {})
      };
      let response;
      if (String(method).toUpperCase() === "GET") {
        endpoint.searchParams.set("action", action);
        endpoint.searchParams.set("_ts", String(Date.now()));
        Object.entries(requestPayload).forEach(([key, value]) => {
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
          body: JSON.stringify({ action, ...requestPayload, _ts: Date.now() }),
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

  function buildCoachflowAuthPayload() {
    const auth = buildAppsScriptAuthPayload();
    const adminCode = String(window.APP_CONFIG?.adminAccessCode || "").trim();
    if (adminCode) {
      auth.adminAccess = adminCode;
    }
    return auth;
  }

  function enqueueCompensationTask(type, payload, reason) {
    state.compensationTasks = Array.isArray(state.compensationTasks) ? state.compensationTasks : [];
    const dedupeKey = getCompensationTaskDedupeKey(type, payload);
    if (dedupeKey) {
      const existing = state.compensationTasks.find((task) => (
        task.status !== "completed" &&
        (task.dedupeKey || getCompensationTaskDedupeKey(task.type, task.payload)) === dedupeKey
      ));
      if (existing) {
        existing.payload = payload;
        existing.reason = reason || existing.reason || "";
        existing.status = existing.status || "pending";
        existing.updatedAt = new Date().toISOString();
        addLog(`[補償任務] ${type} 已存在，略過重複建立：${reason}`);
        saveState();
        return existing;
      }
    }
    state.compensationTasks.unshift({
      id: newId("COMP"),
      type,
      payload,
      reason,
      status: "pending",
      dedupeKey,
      createdAt: new Date().toISOString()
    });
    state.compensationTasks = state.compensationTasks.slice(0, 80);
    addLog(`[補償任務] ${type} 已建立：${reason}`);
    saveState();
  }

  function getCompensationTaskDedupeKey(type, payload = {}) {
    payload = payload || {};
    const normalizedType = String(type || "").trim();
    const lessonId = String(payload.lessonId || "").trim();
    const studentCode = normalizeParticipantCode(payload.studentCode || "");
    const coachCode = normalizeParticipantCode(payload.coachCode || "");
    const eventId = normalizeCalendarEventId(payload.eventId || payload.calendarEventId || "");
    const startAt = String(payload.occurrenceStartAt || payload.lessonStartAt || payload.startAt || "").trim();
    if (normalizedType === "deleteEvent" || normalizedType === "deleteSingleEvent") {
      return [normalizedType, lessonId, eventId, startAt].join("|");
    }
    if (normalizedType === "createEvent") {
      return [normalizedType, lessonId, studentCode, coachCode, startAt].join("|");
    }
    if (normalizedType === "sendEmail") {
      return [
        normalizedType,
        String(payload.template || "").trim(),
        studentCode,
        coachCode,
        startAt || String(payload.when || "").trim(),
        String(payload.milestone || payload.totalChargedCount || "").trim()
      ].join("|");
    }
    return "";
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

  function collapseDuplicateGoogleLessons(scope = {}) {
    const syncScope = getCloudStateScope(scope);
    const groups = new Map();
    (state.lessons || []).forEach((lesson) => {
      if (!lesson || !isGoogleSyncLesson(lesson) || !isLessonInScope(lesson, syncScope) || !hasValidDateValue(lesson.startAt)) {
        return;
      }
      const key = getLessonCloudMatchKey(lesson);
      if (!key) return;
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(lesson);
    });

    const activeLeavesByLesson = new Map();
    (state.leaveRequests || []).forEach((leave) => {
      if (!leave || leave.revokedAt || String(leave.type || "normal") !== "normal") return;
      const lessonId = String(leave.lessonId || "").trim();
      if (!lessonId) return;
      if (!activeLeavesByLesson.has(lessonId)) activeLeavesByLesson.set(lessonId, []);
      activeLeavesByLesson.get(lessonId).push(leave);
    });

    const priority = (lesson) => {
      const leaves = activeLeavesByLesson.get(String(lesson.id || "")) || [];
      const hasRealLeaveEvent = leaves.some((leave) => normalizeCalendarEventId(leave.calendarEventId) && !isGeneratedLocalCalendarEventId(leave.calendarEventId));
      const hasRealLessonEvent = normalizeCalendarEventId(lesson.calendarEventId) && !isGeneratedLocalCalendarEventId(lesson.calendarEventId);
      return (hasRealLeaveEvent ? 1000 : 0) + (leaves.length ? 500 : 0) + (hasRealLessonEvent ? 200 : 0) +
        (lesson.attendanceStatus === "leave-normal" ? 100 : 0) + (lesson.calendarOccupied ? 20 : 0);
    };

    const duplicateIds = [];
    const nowIso = new Date().toISOString();
    groups.forEach((lessons) => {
      if (lessons.length < 2) return;
      const sorted = [...lessons].sort((left, right) => {
        const scoreDiff = priority(right) - priority(left);
        return scoreDiff || String(left.id || "").localeCompare(String(right.id || ""));
      });
      sorted.slice(1).forEach((lesson) => {
        lesson.sourceType = "DUPLICATE";
        lesson.calendarOccupied = false;
        lesson.attendanceStatus = "calendar-removed";
        lesson.charged = false;
        lesson.updatedAt = nowIso;
        duplicateIds.push(lesson.id);
      });
    });
    if (duplicateIds.length) {
      addLog(`[重複課程修復] 已停用 ${duplicateIds.length} 筆同學生、同時段的重複課程：${duplicateIds.join(", ")}。`);
    }
    return duplicateIds;
  }

  function getCloudLeaveMatchKeyFromValues(studentCode, coachCode, startAt) {
    const normalizedStudentCode = normalizeParticipantCode(studentCode);
    const normalizedCoachCode = normalizeParticipantCode(coachCode);
    const startTime = new Date(startAt || "").getTime();
    if (!normalizedStudentCode || !normalizedCoachCode || !Number.isFinite(startTime)) {
      return "";
    }
    return [
      normalizedStudentCode,
      normalizedCoachCode,
      new Date(startTime).toISOString()
    ].join("|");
  }

  function getCloudLeaveRecordMatchKey(record) {
    const lessonKey = String(record?.lessonKey || "").trim();
    // 欄位值是目前雲端紀錄實際顯示的學生／教練／時段；lessonKey 只是舊版
    // 冗餘快取，可能在課程曾被錯誤對齊後留下舊值。優先重算 tuple，只有
    // 舊紀錄缺欄位時才退回 lessonKey。
    return getCloudLeaveMatchKeyFromValues(record?.studentCode, record?.coachCode, record?.lessonStartAt) || lessonKey;
  }

  function isLeaveInCloudSyncScope(leave, lesson, scope) {
    const studentCode = normalizeParticipantCode(leave?.studentCode || lesson?.studentCode);
    const coachCode = normalizeParticipantCode(leave?.coachCode || lesson?.coachCode);
    if (scope.studentCode && studentCode !== scope.studentCode) {
      return false;
    }
    if (scope.coachCode && coachCode !== scope.coachCode) {
      return false;
    }
    return Boolean(studentCode || coachCode);
  }

  function reconcileLocalLeaveRecordsWithCloud(records, scope) {
    if (!scope.coachCode && !scope.studentCode) {
      return false;
    }

    const activeCloudIds = new Set();
    const activeCloudLessonKeys = new Set();
    records.forEach((record) => {
      if (!record || String(record.type || "normal") !== "normal" || String(record.revokedAt || "").trim()) {
        return;
      }
      const id = String(record.id || "").trim();
      const lessonKey = getCloudLeaveRecordMatchKey(record);
      if (id) {
        activeCloudIds.add(id);
      }
      if (lessonKey) {
        activeCloudLessonKeys.add(lessonKey);
      }
    });

    // 找出「本地有但雲端找不到」的請假紀錄，但**不要**自動取消。
    //
    // 之前的邏輯會把這類紀錄標 revokedAt + 把 lesson 改回 scheduled。問題是
    // 「雲端找不到」常常是「本地當初 push 失敗、雲端從沒收過」而不是
    // 「雲端確實取消了」。已上線案例：學生請假後 email + Google 日曆事件
    // 都已執行（不可逆），但因 push 失敗 5 分鐘後被誤判取消，造成「系統說
    // 排課，Google 日曆已刪」資料不一致，等課程當天還會被自動扣堂。
    //
    // 改成只記錄並補推；不自動修改本地狀態。雲端真的有取消會走
    // applyCloudLeaveRecord 的 revokedAt 分支，那條才是合法路徑。
    const pendingResync = [];
    state.leaveRequests.forEach((leave) => {
      if (!leave || String(leave.type || "normal") !== "normal" || leave.revokedAt) {
        return;
      }
      if (!isAfterTrackingDataReset(leave.submittedAt || leave.updatedAt)) {
        return;
      }
      const submittedTime = new Date(leave.submittedAt || "").getTime();
      if (Number.isFinite(submittedTime) && Date.now() - submittedTime < LOCAL_LEAVE_RECONCILE_GRACE_MS) {
        return;
      }

      const lesson = getLessonById(leave.lessonId);
      if (!isLeaveInCloudSyncScope(leave, lesson, scope)) {
        return;
      }

      const leaveId = String(leave.id || "").trim();
      const lessonKey = lesson ? getLessonCloudMatchKey(lesson) : "";
      if ((leaveId && activeCloudIds.has(leaveId)) || (lessonKey && activeCloudLessonKeys.has(lessonKey))) {
        return;
      }

      pendingResync.push({ leave, lesson });
    });

    if (!pendingResync.length) {
      return false;
    }

    // fire-and-forget 補推；錯誤靠 addLog 留痕，不擋住主流程
    pendingResync.forEach(({ leave, lesson }) => {
      addLog(`[雲端對帳] ⚠️ 本地請假 ${leave.id || "(無 ID)"} 雲端查無對應紀錄，嘗試補推到雲端（不取消本地）。`);
      Promise.resolve()
        .then(() => pushCloudLeaveRecord(leave, lesson))
        .then((ok) => {
          if (ok) {
            addLog(`[雲端對帳] ✓ 已補推請假 ${leave.id || "(無 ID)"} 至雲端。`);
            saveState();
          } else {
            addLog(`[雲端對帳] ✗ 補推請假 ${leave.id || "(無 ID)"} 雲端回應失敗，下次同步會再試。`);
          }
        })
        .catch((error) => {
          addLog(`[雲端對帳] ✗ 補推請假 ${leave.id || "(無 ID)"} 拋出例外：${String(error?.message || error)}`);
        });
    });

    // 沒有對 state.leaveRequests / state.lessons 做同步寫入，回 false 讓上層
    // 不要因此覆寫 state（補推由 promise 自己 saveState）
    return false;
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
      calendarEventId: leave.calendarEventId || targetLesson.calendarEventId || "",
      studentCode: normalizeParticipantCode(leave.studentCode || targetLesson.studentCode),
      coachCode: normalizeParticipantCode(leave.coachCode || targetLesson.coachCode),
      lessonStartAt: targetLesson.startAt,
      type: leave.type || "normal",
      submittedAt: leave.submittedAt || new Date().toISOString(),
      submittedBy: leave.submittedBy || "",
      submittedByRole: leave.submittedByRole || "",
      makeupEligible: leave.makeupEligible !== false,
      emailNoticeStatus: leave.emailNoticeStatus || "",
      emailNoticeAt: leave.emailNoticeSentAt || leave.emailNoticeQueuedAt || "",
      revokedAt: leave.revokedAt || "",
      revokedBy: leave.revokedBy || ""
    };
  }

  // 精簡成雲端 Lessons 分頁的 10 欄。calendarEventId / completedBy / beforeCalendarRemoved
  // 等本機專用欄位不上雲（計費用不到）。updatedAt fallback 鏈：
  // updatedAt → completedAt → startAt → now，確保每筆都有可比較的時間戳，
  // 讓後端 +2000ms stale 防護與主系統「比 updatedTime 取較新者」都能正確運作。
  function normalizeLessonForCloud(lesson) {
    if (!lesson || !lesson.id) {
      return null;
    }
    const startTime = new Date(lesson.startAt || "").getTime();
    if (!Number.isFinite(startTime)) {
      return null;
    }
    const startAt = new Date(startTime).toISOString();
    return {
      id: String(lesson.id),
      studentCode: normalizeParticipantCode(lesson.studentCode),
      coachCode: normalizeParticipantCode(lesson.coachCode),
      startAt,
      sourceType: String(lesson.sourceType || "REGULAR"),
      calendarOccupied: lesson.calendarOccupied !== false,
      attendanceStatus: String(lesson.attendanceStatus || "scheduled"),
      charged: Boolean(lesson.charged),
      completedAt: lesson.completedAt || "",
      updatedAt: lesson.updatedAt || lesson.completedAt || startAt || new Date().toISOString()
    };
  }

  function findLessonForCloudLeaveRecord(record) {
    if (!record) {
      return null;
    }
    const recordMatchKey = getCloudLeaveRecordMatchKey(record);
    const lessonId = String(record.lessonId || "").trim();
    if (lessonId) {
      const byId = getLessonById(lessonId);
      // lessonId 是最強識別，但仍要防止舊快照裡極少數重複／污染的 ID
      // 把請假套到不同學生或不同時段。舊資料若沒有完整 tuple，才保留
      // 單憑 lessonId 的相容行為。
      if (byId && (!recordMatchKey || getLessonCloudMatchKey(byId) === recordMatchKey)) {
        return byId;
      }
    }
    const recordEventId = normalizeCalendarEventId(record.calendarEventId);
    if (recordEventId && recordMatchKey) {
      // Google recurring event 的每個 occurrence 可能共用同一個 eventId。
      // 過去只憑 eventId 取第一堂，套用請假後又清掉該堂 eventId；下次同步
      // 就會依序命中下一週，造成單筆請假一路蔓延。eventId 現在只能當輔助
      // 識別，學生＋教練＋開始時間必須完全一致。
      const byEventId = state.lessons.find((lesson) => (
        normalizeCalendarEventId(lesson.calendarEventId) === recordEventId &&
        getLessonCloudMatchKey(lesson) === recordMatchKey
      ));
      if (byEventId) {
        return byEventId;
      }
    }
    // 使用 student+coach+startAt 嚴格比對。歷史版本曾以 ±2 分鐘模糊比對，
    // 在跨裝置 lessonId 不一致 + calendarEventId 缺失時，會把雲端某一天的
    // 請假紀錄誤套到本地另一筆相近時刻的 lesson。改成 ISO 完全一致才算數，
    // 避免再發生「請 5/10 結果套到 5/31」這種跨日誤配對。
    if (!recordMatchKey) {
      return null;
    }
    return state.lessons.find((lesson) => getLessonCloudMatchKey(lesson) === recordMatchKey) || null;
  }

  function restoreOrphanedNormalLeaveLessons(scope = {}) {
    const syncScope = getCloudStateScope(scope);
    const activeLeaveLessonIds = new Set(
      (state.leaveRequests || [])
        .filter((leave) => (
          leave &&
          String(leave.type || "normal") === "normal" &&
          !leave.revokedAt &&
          isLeaveInCloudSyncScope(leave, getLessonById(leave.lessonId), syncScope)
        ))
        .map((leave) => String(leave.lessonId || "").trim())
        .filter(Boolean)
    );
    const restoredLessonIds = [];
    const nowIso = new Date().toISOString();
    (state.lessons || []).forEach((lesson) => {
      if (
        !lesson ||
        lesson.attendanceStatus !== "leave-normal" ||
        !isLessonInScope(lesson, syncScope) ||
        activeLeaveLessonIds.has(String(lesson.id || "").trim())
      ) {
        return;
      }
      lesson.attendanceStatus = "scheduled";
      lesson.calendarOccupied = true;
      lesson.charged = false;
      lesson.updatedAt = nowIso;
      restoredLessonIds.push(lesson.id);
    });
    if (restoredLessonIds.length) {
      addLog(`[雲端請假修復] 已還原 ${restoredLessonIds.length} 堂沒有有效請假紀錄、但被錯標為請假的課程：${restoredLessonIds.join(", ")}。`);
    }
    return restoredLessonIds;
  }

  function ensureLessonForCloudLeaveRecord(record) {
    const studentCode = normalizeParticipantCode(record?.studentCode);
    const coachCode = normalizeParticipantCode(record?.coachCode);
    const startAt = String(record?.lessonStartAt || "").trim();
    if (!studentCode || !coachCode || !hasValidDateValue(startAt)) {
      return null;
    }
    // 只有在雲端紀錄帶有 lessonId 或 calendarEventId（強識別）時才允許
    // 自動補課。否則只憑 student+coach+startAt 就建出新 lesson，會在
    // 跨裝置同步時製造幽靈課程（例如把雲端某筆請假的時間對到另一個
    // student 的同時段，造成「系統自己跑出 5/31 請假」這種事故）。
    const hasStrongIdentity = Boolean(
      String(record?.lessonId || "").trim() ||
      normalizeCalendarEventId(record?.calendarEventId || "")
    );
    if (!hasStrongIdentity) {
      addLog(`[雲端請假] 跳過自動補課：雲端紀錄缺 lessonId 與 calendarEventId（${studentCode} / ${formatDateTime(startAt)}）。`);
      return null;
    }
    ensureCoachProfile(coachCode);
    ensureStudentProfile(studentCode, coachCode, { silentMode: true });
    const dateKey = getDateKeyInTaipei(new Date(startAt));
    const timeText = getTimeText(startAt);
    const lesson = makeLesson(newId("L"), studentCode, coachCode, dateKey, timeText);
    lesson.sourceType = "GOOGLE_CALENDAR";
    lesson.startAt = new Date(startAt).toISOString();
    lesson.calendarEventId = normalizeCalendarEventId(record.calendarEventId || "") || "";
    lesson.calendarOccupied = false;
    lesson.attendanceStatus = "leave-normal";
    lesson.charged = false;
    state.lessons.push(lesson);
    addLog(`[雲端請假] 已補回 ${studentCode} ${formatDateTime(lesson.startAt)} 的請假課程。`);
    return lesson;
  }

  function applyCloudLeaveRecord(record) {
    if (!record || String(record.type || "normal") !== "normal") {
      return false;
    }
    const lesson = findLessonForCloudLeaveRecord(record) || ensureLessonForCloudLeaveRecord(record);
    if (!lesson) {
      console.warn("[雲端請假] applyCloudLeaveRecord 找不到對應課程：", {
        recordId: record.id,
        studentCode: record.studentCode,
        coachCode: record.coachCode,
        lessonId: record.lessonId,
        calendarEventId: record.calendarEventId,
        lessonStartAt: record.lessonStartAt
      });
      addLog(`[雲端請假] 找不到對應課程，跳過 record ${record.id}（${record.studentCode} / ${formatDateTime(record.lessonStartAt)}）。lessonId=${record.lessonId || "(空)"} eventId=${record.calendarEventId || "(空)"}`);
      return false;
    }
    const existing = state.leaveRequests.find((leave) => leave.id === record.id) ||
      state.leaveRequests.find((leave) => leave.lessonId === lesson.id && leave.studentCode === record.studentCode && !leave.revokedAt);
    const recordCalendarEventId = String(record.calendarEventId || "").trim();
    let changed = false;
    const revokedAt = String(record.revokedAt || "").trim();
    if (revokedAt) {
      if (existing && !existing.revokedAt) {
        existing.revokedAt = revokedAt;
        existing.revokedBy = record.revokedBy || "cloud";
        changed = true;
      }
      // 同一個 lesson 上可能存在多筆 leave 紀錄（學生請假 → 取消 → 再請假）。
      // 不能只因為某一筆 revoked record 就把 lesson 改回 scheduled —— 必須先確認
      // 沒有「另一筆活著的 leave」也指向這個 lesson。否則會把已成立的請假誤砍。
      // 已上線案例：ST019 5/10 lesson 同時有 LEAVE_09cwcf（活）和 LEAVE_hox49h（revoked），
      // 處理到後者時把 lesson 蓋回 scheduled，造成 email 寄出但系統顯示已排課。
      const recordId = String(record.id || "").trim();
      const otherActiveLeave = state.leaveRequests.some((other) => (
        other &&
        other.lessonId === lesson.id &&
        String(other.id || "") !== recordId &&
        !other.revokedAt &&
        String(other.type || "normal") === "normal"
      ));
      if (lesson.attendanceStatus === "leave-normal" && !otherActiveLeave) {
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
        calendarEventId: recordCalendarEventId,
        type: "normal",
        submittedAt: record.submittedAt || new Date().toISOString(),
        submittedBy: record.submittedBy || "",
        submittedByRole: record.submittedByRole || "",
        makeupEligible: record.makeupEligible !== false
      });
      changed = true;
    } else {
      // 同一筆雲端 leave 曾因 recurring eventId 誤配到別堂時，除了修正 lesson
      // 狀態，也要把本機 leave 的指向搬回權威 tuple，否則孤立狀態無法清理。
      if (existing.lessonId !== lesson.id) {
        existing.lessonId = lesson.id;
        changed = true;
      }
      if (normalizeParticipantCode(existing.studentCode) !== normalizeParticipantCode(record.studentCode)) {
        existing.studentCode = record.studentCode;
        changed = true;
      }
      if (normalizeParticipantCode(existing.coachCode) !== normalizeParticipantCode(record.coachCode)) {
        existing.coachCode = record.coachCode;
        changed = true;
      }
      if (recordCalendarEventId && normalizeCalendarEventId(existing.calendarEventId) !== normalizeCalendarEventId(recordCalendarEventId)) {
        existing.calendarEventId = recordCalendarEventId;
        changed = true;
      }
    }
    if (lesson.attendanceStatus !== "leave-normal" || lesson.calendarOccupied !== false) {
      lesson.attendanceStatus = "leave-normal";
      lesson.calendarOccupied = false;
      changed = true;
    }
    if (lesson.calendarEventId) {
      lesson.calendarEventId = "";
      changed = true;
    }
    return changed;
  }

  // 把本機 state.leaveRequests 中還沒有 cloudSyncedAt 旗標的請假（或上次推送失敗的）
  // 重新推上雲端。用於救援：當時 saveLeaveRecord 因為 Apps Script 配額失敗、本機已有
  // 紀錄但雲端沒有的情況。saveLeaveRecord 端是 upsert，重複呼叫安全。
  //
  // 兩種 unsynced 來源：
  // 1. pendingCloudSync = true（新版 atomic flow）：雲端首發失敗，當時還沒動 GCal /
  //    寄信。雲端 push 成功後要把 lesson 狀態改成 leave-normal、刪 GCal、寄信、
  //    再 push 一次帶 email 狀態的雲端紀錄。
  // 2. cloudSyncedAt missing 但沒 pendingCloudSync 旗標（舊版資料）：當時本機已
  //    改 lesson、已刪 GCal、已寄信，只差雲端紀錄。補推一次就好。
  //
  // silent=true：背景補救（focus/visibility）用，不彈 notifyUser，依賴 banner 持續顯示。
  //
  // Scope 規則：跟 renderPendingSyncBanner 一致。
  // - 有 activeStudentCode → 只 retry 那位學生本機的請假
  // - 沒 student 但有 coach → retry 那位教練底下的（教練端可一鍵推完所有未同步的）
  // - 都沒有 → 跳過（沒登入時不該偷偷打 API）
  // 早期版本不 scope，會在「學生測試頁」這種共用 state 場景，把其他學生未同步的請假
  // 一併推上去、並在 notifyUser 統計裡把總數呈現給學生看，洩漏隔壁資料。
  async function retryUnsyncedLocalLeaves(options = {}) {
    const silent = Boolean(options && options.silent);
    if (!getAppsScriptUrl()) {
      if (!silent) {
        addLog("[雲端請假補救] 未設定 Apps Script URL，跳過。");
      }
      return { tried: 0, succeeded: 0, failed: 0, completedPending: 0 };
    }
    if (!activeStudentCode && !activeCoachCode) {
      if (!silent) {
        addLog("[雲端請假補救] 尚未登入學生或教練，跳過。");
      }
      return { tried: 0, succeeded: 0, failed: 0, completedPending: 0 };
    }
    const candidates = (state.leaveRequests || []).filter((leave) => {
      if (!leave || leave.revokedAt) return false;
      if (String(leave.type || "normal") !== "normal") return false;
      if (leave.cloudSyncedAt) return false;
      if (activeStudentCode) {
        return leave.studentCode === activeStudentCode;
      }
      return leave.coachCode === activeCoachCode;
    });
    if (!candidates.length) {
      if (!silent) {
        addLog("[雲端請假補救] 沒有需要補推的請假紀錄。");
      }
      return { tried: 0, succeeded: 0, failed: 0, completedPending: 0 };
    }
    let succeeded = 0;
    let failed = 0;
    let completedPending = 0;
    for (const leave of candidates) {
      const lesson = getLessonById(leave.lessonId);
      const wasPending = Boolean(leave.pendingCloudSync);
      try {
        const ok = await pushCloudLeaveRecord(leave, lesson);
        if (ok) {
          leave.cloudSyncedAt = new Date().toISOString();
          if (wasPending) {
            // 新版 pending leave 推上去後，把當時被擋掉的 lesson 狀態 / GCal / Email 補完
            delete leave.pendingCloudSync;
            if (lesson) {
              lesson.calendarOccupied = false;
              lesson.attendanceStatus = "leave-normal";
            }
            await completePendingLeaveSideEffects(leave, lesson);
            completedPending += 1;
          }
          succeeded += 1;
          addLog(`[雲端請假補救] 已補推 ${leave.id}（${leave.studentCode} / ${lesson ? formatDateTime(lesson.startAt) : leave.lessonId}）${wasPending ? "（pending → 補完 GCal+Email）" : ""}。`);
        } else {
          failed += 1;
          addLog(`[雲端請假補救] 補推回應 ok=false：${leave.id}`);
        }
      } catch (error) {
        failed += 1;
        addLog(`[雲端請假補救] 補推失敗：${leave.id} / ${String(error?.message || error)}`);
      }
    }
    if (succeeded > 0 || failed > 0) {
      saveState();
    }
    if (!silent) {
      if (succeeded > 0) {
        notifyUser(
          `已補推 ${succeeded} 筆請假紀錄到雲端${completedPending ? `（其中 ${completedPending} 筆同時補完 Google 日曆與通知信）` : ""}。${failed ? `（${failed} 筆失敗）` : ""}`,
          succeeded ? "success" : "warning"
        );
      } else if (failed > 0) {
        notifyUser(`本機請假紀錄補推失敗（${failed} 筆），請查 console。`, "warning");
      }
    }
    return { tried: candidates.length, succeeded, failed, completedPending };
  }

  // pendingCloudSync 的請假在雲端 push 成功後，把當時 atomic flow 沒做的 GCal 刪除
  // 與 email 通知補上。為了 retry 安全，所有動作都 try/catch，個別失敗只記 log，不影響
  // 已經寫到雲端的 leave record（畢竟雲端是 source of truth）。
  async function completePendingLeaveSideEffects(leave, lesson) {
    if (!lesson) {
      addLog(`[雲端請假補救] 找不到 lesson ${leave.lessonId}，跳過 GCal/Email 完成。`);
      return;
    }
    try {
      const ok = await tryDeleteCalendarEventForLesson(lesson, "retry_apply_normal_leave", false, {
        resolveMissingEventId: true
      });
      if (!ok) {
        addLog(`[Google日曆] 補推請假 ${leave.id} 後刪 GCal 事件未成功，留待下次同步處理。`);
      }
    } catch (error) {
      addLog(`[Google日曆] 補推請假 ${leave.id} 後刪 GCal 出錯：${String(error?.message || error)}`);
    }

    const isCoachLeave = String(leave.submittedByRole || "") === "coach";
    const studentEmail = getStudentNoticeEmail(leave.studentCode);
    if (!studentEmail) {
      addLog(`[Email] ${leave.studentCode} 未設定學生 Email，補寄請假通知信時將略過學生收件人。`);
    }
    const emailPayload = {
      studentCode: leave.studentCode,
      coachCode: leave.coachCode,
      studentEmail,
      coachEmail: getCoachNoticeEmail(leave.coachCode),
      lessonId: lesson.id,
      lessonStartAt: lesson.startAt,
      leaveId: leave.id,
      submittedBy: leave.submittedBy || "",
      submittedByRole: leave.submittedByRole || ""
    };
    leave.emailNoticeTo = resolveNoticeRecipients(emailPayload).join(", ");
    let sent = false;
    try {
      sent = await trySendEmailNotice(
        isCoachLeave ? "leave_submitted_by_coach" : "leave_submitted",
        emailPayload,
        isCoachLeave
          ? `教練代請假通知 ${leave.studentCode} / ${lesson.id}`
          : `學生 ${leave.studentCode} 請假通知`
      );
    } catch (error) {
      addLog(`[Email] 補推請假 ${leave.id} 寄信失敗：${String(error?.message || error)}`);
    }
    const emailAt = new Date().toISOString();
    leave.emailNoticeStatus = sent ? "sent" : "queued-or-skipped";
    if (sent) {
      leave.emailNoticeSentAt = emailAt;
    } else {
      leave.emailNoticeQueuedAt = emailAt;
    }
    try {
      await pushCloudLeaveRecord(leave, lesson);
    } catch (error) {
      addLog(`[雲端請假] 補推請假 ${leave.id} 之 Email 狀態同步失敗：${String(error?.message || error)}`);
    }
  }

  // 把一筆 lesson 從 leave-normal 強制還原成 scheduled，並把同 lesson 的 active
  // leave 一併 mark revoked、cloud push。
  //
  // 適用兩種 ghost 場景：
  // A. 有 active leave + lesson=leave-normal：跟 revokeNormalLeave 一樣，但跳過
  //    UI 的 isLeaveOpen 截止時間檢查，純救援。
  // B. 沒 active leave，但 lesson 卡在 leave-normal（更詭異的不一致）：
  //    revokeNormalLeave 會擋下「找不到對應請假紀錄」，這支照樣能還原 lesson。
  //
  // 若 calendarEventId 是 local placeholder（GCAL_L_...），清掉並嘗試重建 GCal
  // 事件；若是真實的 GCal 事件 ID（例如 recurring instance），保留並依賴下一次
  // 教練同步 Google 日曆對齊。
  async function restoreLessonToScheduledForDebug(lessonId) {
    if (!requireCoachWriteAccess()) {
      return;
    }
    const lesson = getLessonById(lessonId);
    if (!lesson) {
      console.warn("[restoreLesson] 找不到 lesson:", lessonId);
      return;
    }
    if (lesson.coachCode !== activeCoachCode) {
      console.warn("[restoreLesson] lesson 不屬於目前登入教練:", lesson.coachCode, "vs", activeCoachCode);
      return;
    }
    const before = {
      attendanceStatus: lesson.attendanceStatus,
      calendarOccupied: lesson.calendarOccupied,
      calendarEventId: lesson.calendarEventId
    };
    console.log("[restoreLesson] before:", JSON.stringify(before));

    const now = new Date().toISOString();
    const activeLeaves = (state.leaveRequests || []).filter((leave) => (
      leave && leave.lessonId === lessonId && !leave.revokedAt
    ));
    for (const leave of activeLeaves) {
      leave.revokedAt = now;
      leave.revokedBy = `${activeCoachCode}:debug-restore`;
      leave.makeupEligible = false;
      try {
        await pushCloudLeaveRecord(leave, lesson);
      } catch (error) {
        console.warn("[restoreLesson] cloud push for revoked leave failed:", leave.id, error);
      }
    }

    lesson.attendanceStatus = "scheduled";
    lesson.calendarOccupied = true;
    lesson.charged = false;
    if (lesson.calendarEventId && isGeneratedLocalCalendarEventId(lesson.calendarEventId)) {
      lesson.calendarEventId = "";
    }
    saveState();
    renderAll();

    if (!lesson.calendarEventId) {
      try {
        await tryCreateCalendarEventForLesson(lesson, "debug_restore_lesson", false);
      } catch (error) {
        console.warn("[restoreLesson] GCal recreate failed:", error);
      }
      saveState();
      renderAll();
    }

    addLog(`[debug] 還原 lesson ${lesson.id}（${formatDateTime(lesson.startAt)}），revoke ${activeLeaves.length} 筆 leave。`);
    console.log("[restoreLesson] after:", JSON.stringify({
      attendanceStatus: lesson.attendanceStatus,
      calendarOccupied: lesson.calendarOccupied,
      calendarEventId: lesson.calendarEventId,
      revokedLeaves: activeLeaves.map((l) => l.id)
    }));
  }

  // 暴露 debug helpers 到 window，console 直接呼叫用：
  // - retryUnsyncedLocalLeaves()：補推未同步請假
  // - revokeNormalLeave(lessonId)：教練取消學生的正常請假（含 GCal 重建 + 雲端同步）
  // - applyNormalLeaveByCoach(lessonId)：教練代學生請假
  // - restoreLessonToScheduled(lessonId)：強制把 leave-normal lesson 還原 scheduled，
  //   並 revoke 對應的 active leave。處理 ghost lesson 用。
  // - coachflowDebug.state：唯讀 access state（取代「state is not defined」）
  // - coachflowDebug.findLesson(studentCode, dateKey)：快速撈某學生某日的 lesson record
  // - coachflowDebug.findActiveLeaves(studentCode)：列出該學生未取消的請假
  if (typeof window !== "undefined") {
    window.retryUnsyncedLocalLeaves = retryUnsyncedLocalLeaves;
    window.revokeNormalLeave = revokeNormalLeave;
    window.applyNormalLeaveByCoach = applyNormalLeaveByCoach;
    window.restoreLessonToScheduled = restoreLessonToScheduledForDebug;
    window.coachflowDebug = {
      get state() { return state; },
      findLesson(studentCode, dateKey) {
        return (state.lessons || []).filter((l) => (
          l.studentCode === studentCode &&
          (!dateKey || getDateKeyInTaipei(new Date(l.startAt)) === dateKey)
        ));
      },
      findActiveLeaves(studentCode) {
        return (state.leaveRequests || []).filter((l) => (
          (!studentCode || l.studentCode === studentCode) && !l.revokedAt
        ));
      }
    };
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
    const payload = getCloudStateScope(scope);
    const result = await callAppsScriptApi("listLeaveRecords", payload, "GET");
    if (result?.ok === false || !Array.isArray(result?.records)) {
      return false;
    }
    const records = (Array.isArray(result?.records) ? result.records : [])
      .filter((record) => isAfterTrackingDataReset(record.submittedAt || record.updatedAt));
    let changed = false;
    records.forEach((record) => {
      if (applyCloudLeaveRecord(record)) {
        changed = true;
      }
    });
    // listLeaveRecords 成功後，雲端清單即為這個 scope 的請假真相來源。
    // 清掉舊版 recurring eventId bug 留下的「沒有 active leave、卻仍是
    // leave-normal」課程，並在後續既有流程推回 Lessons 雲端表。
    if (restoreOrphanedNormalLeaveLessons(payload).length) {
      changed = true;
    }
    if (reconcileLocalLeaveRecordsWithCloud(records, payload)) {
      changed = true;
    }
    if (collapseDuplicateGoogleLessons(payload).length) {
      changed = true;
    }
    if (changed) {
      addLog(`[雲端請假] 已同步 ${records.length} 筆請假紀錄。`);
      saveState();
    }
    return changed;
  }

  // 把本機完整課表（tracking-start 之後）推上雲端 Lessons 分頁。
  // 後端 saveLessons 是差量 upsert by id，重複 push 安全。分批 100 筆避免單次
  // payload 過大。後端未部署（Unsupported action）時 graceful 回 false，呼叫端
  // 退回現有 local/snapshot 兩來源，零行為退化。
  // 回傳：成功時為已 push 的課堂數；後端未支援/回 ok:false 時為 false。
  async function pushCloudLessons(scope = {}) {
    if (!getAppsScriptUrl()) {
      return false;
    }
    const syncScope = getCloudStateScope(scope);
    const lessons = (state.lessons || [])
      .filter((lesson) => (
        lesson &&
        isLessonAfterBillingTrackingStart(lesson) &&
        (!syncScope.coachCode || normalizeParticipantCode(lesson.coachCode) === syncScope.coachCode) &&
        (!syncScope.studentCode || normalizeParticipantCode(lesson.studentCode) === syncScope.studentCode)
      ))
      .map(normalizeLessonForCloud)
      .filter(Boolean);
    if (!lessons.length) {
      return 0;
    }
    const batchSize = 100;
    let pushed = 0;
    try {
      for (let i = 0; i < lessons.length; i += batchSize) {
        const batch = lessons.slice(i, i + batchSize);
        const result = await callAppsScriptApi("saveLessons", { lessons: batch });
        if (result?.ok === false) {
          return false;
        }
        pushed += batch.length;
      }
      return pushed;
    } catch (error) {
      if (isUnsupportedAppsScriptAction(error, "saveLessons")) {
        return false;
      }
      throw error;
    }
  }

  // Fire-and-forget 版：請假/改狀態/日曆同步等接入點呼叫，雲端課表持續保鮮，
  // 失敗只記 console 不影響主流程。
  function pushCloudLessonsQuietly(scope = {}) {
    pushCloudLessons(scope).catch((error) => {
      console.warn("Cloud lessons push failed:", error);
    });
  }

  function clonePlain(value) {
    return JSON.parse(JSON.stringify(value));
  }

  function getCloudStateScope(scope = {}) {
    const hasCoachScope = Object.prototype.hasOwnProperty.call(scope, "coachCode");
    const hasStudentScope = Object.prototype.hasOwnProperty.call(scope, "studentCode");
    const shouldUseActiveStudent = !hasCoachScope && !hasStudentScope;
    return {
      coachCode: normalizeParticipantCode(scope.coachCode || (!hasCoachScope ? activeCoachCode : "") || ""),
      studentCode: normalizeParticipantCode(scope.studentCode || (shouldUseActiveStudent ? activeStudentCode : "") || "")
    };
  }

  function isStudentInScope(student, scope) {
    const studentCode = normalizeParticipantCode(student?.code || student?.studentCode);
    const coachCode = normalizeParticipantCode(student?.coachCode);
    if (scope.studentCode && studentCode !== scope.studentCode) {
      return false;
    }
    if (scope.coachCode && coachCode !== scope.coachCode) {
      return false;
    }
    return Boolean(studentCode || coachCode);
  }

  function isCoachInScope(coach, scope) {
    const coachCode = normalizeParticipantCode(coach?.code || coach?.coachCode);
    return !scope.coachCode || coachCode === scope.coachCode;
  }

  function isLessonInScope(lesson, scope) {
    const studentCode = normalizeParticipantCode(lesson?.studentCode);
    const coachCode = normalizeParticipantCode(lesson?.coachCode);
    if (scope.studentCode && studentCode !== scope.studentCode) {
      return false;
    }
    if (scope.coachCode && coachCode !== scope.coachCode) {
      return false;
    }
    return Boolean(studentCode || coachCode);
  }

  function isMakeupRequestInScope(request, scope) {
    const studentCode = normalizeParticipantCode(request?.studentCode);
    const coachCode = normalizeParticipantCode(request?.coachCode);
    if (scope.studentCode && studentCode !== scope.studentCode) {
      return false;
    }
    if (scope.coachCode && coachCode !== scope.coachCode) {
      return false;
    }
    return Boolean(studentCode || coachCode);
  }

  function isCoachBlockInScope(block, scope) {
    const coachCode = normalizeParticipantCode(block?.coachCode);
    return Boolean(coachCode) && (!scope.coachCode || coachCode === scope.coachCode);
  }

  function isCompensationTaskInScope(task, scope) {
    const payload = task?.payload || {};
    const coachCode = normalizeParticipantCode(payload.coachCode || task?.coachCode);
    const studentCode = normalizeParticipantCode(payload.studentCode || task?.studentCode);
    if (scope.studentCode && studentCode !== scope.studentCode) {
      return false;
    }
    if (scope.coachCode && coachCode !== scope.coachCode) {
      return false;
    }
    return Boolean(studentCode || coachCode);
  }

  function buildLeaveStateSnapshot(scope = {}) {
    const syncScope = getCloudStateScope(scope);
    const snapshotState = {
      students: state.students.filter((student) => isStudentInScope(student, syncScope)).map(clonePlain),
      coaches: state.coaches.filter((coach) => isCoachInScope(coach, syncScope)).map(clonePlain),
      lessons: state.lessons.filter((lesson) => isLessonInScope(lesson, syncScope)).map(clonePlain),
      leaveRequests: state.leaveRequests
        .filter((leave) => isLeaveInCloudSyncScope(leave, getLessonById(leave.lessonId), syncScope))
        .map(clonePlain),
      makeupRequests: state.makeupRequests.filter((request) => isMakeupRequestInScope(request, syncScope)).map(clonePlain),
      coachBlocks: (state.coachBlocks || []).filter((block) => isCoachBlockInScope(block, syncScope)).map(clonePlain),
      compensationTasks: (state.compensationTasks || [])
        .filter((task) => isCompensationTaskInScope(task, syncScope))
        .map(clonePlain)
    };
    return {
      version: "1",
      schemaVersion: String(SCHEMA_VERSION),
      coachCode: syncScope.coachCode,
      studentCode: syncScope.studentCode,
      source: "coachflow-leave",
      sourceVersion: new URLSearchParams(window.location.search).get("v") || "",
      updatedAt: new Date().toISOString(),
      updatedBy: syncScope.coachCode || syncScope.studentCode || "SYSTEM",
      baseUpdatedAt: state.cloudSnapshotUpdatedAt || "",
      counts: {
        students: snapshotState.students.length,
        coaches: snapshotState.coaches.length,
        lessons: snapshotState.lessons.length,
        leaveRequests: snapshotState.leaveRequests.length,
        makeupRequests: snapshotState.makeupRequests.length,
        coachBlocks: snapshotState.coachBlocks.length,
        compensationTasks: snapshotState.compensationTasks.length
      },
      state: snapshotState
    };
  }

  function getItemKey(item, fallbackPrefix) {
    return String(item?.id || item?.code || item?.studentCode || item?.coachCode || `${fallbackPrefix}-${Math.random()}`).trim();
  }

  function mergeScopedStateItems(collectionName, incomingItems, isInScope, scope, keyFn) {
    const incoming = Array.isArray(incomingItems) ? incomingItems.map(clonePlain) : [];
    const incomingKeys = new Set();
    const incomingByKey = new Map();
    incoming.forEach((item) => {
      if (!isInScope(item, scope)) {
        return;
      }
      const key = keyFn(item);
      if (!key || incomingKeys.has(key)) {
        return;
      }
      incomingKeys.add(key);
      incomingByKey.set(key, item);
    });
    const nextItems = [];
    const seen = new Set();
    (state[collectionName] || []).forEach((item) => {
      const key = keyFn(item);
      if (key && incomingByKey.has(key)) {
        nextItems.push(incomingByKey.get(key));
        seen.add(key);
        return;
      }
      nextItems.push(item);
      if (key) {
        seen.add(key);
      }
    });
    incomingByKey.forEach((item, key) => {
      if (!seen.has(key)) {
        nextItems.push(item);
      }
    });
    state[collectionName] = nextItems;
  }

  function mergeScopedSnapshotItems(existingItems, incomingItems, isInScope, scope, keyFn) {
    const incomingByKey = new Map();
    (Array.isArray(incomingItems) ? incomingItems : []).forEach((item) => {
      if (!isInScope(item, scope)) {
        return;
      }
      const key = keyFn(item);
      if (key && !incomingByKey.has(key)) {
        incomingByKey.set(key, clonePlain(item));
      }
    });

    const nextItems = [];
    const seenKeys = new Set();
    (Array.isArray(existingItems) ? existingItems : []).forEach((item) => {
      const key = keyFn(item);
      if (key && incomingByKey.has(key)) {
        nextItems.push(incomingByKey.get(key));
        seenKeys.add(key);
        return;
      }
      nextItems.push(clonePlain(item));
      if (key) {
        seenKeys.add(key);
      }
    });
    incomingByKey.forEach((item, key) => {
      if (!seenKeys.has(key)) {
        nextItems.push(item);
      }
    });
    return nextItems;
  }

  function countSnapshotStateItems(snapshotState = {}) {
    return {
      students: (snapshotState.students || []).length,
      coaches: (snapshotState.coaches || []).length,
      lessons: (snapshotState.lessons || []).length,
      leaveRequests: (snapshotState.leaveRequests || []).length,
      makeupRequests: (snapshotState.makeupRequests || []).length,
      coachBlocks: (snapshotState.coachBlocks || []).length,
      compensationTasks: (snapshotState.compensationTasks || []).length
    };
  }

  function mergeLocalScopedStateIntoSnapshot(cloudSnapshot, scope = {}) {
    const syncScope = getCloudStateScope(scope);
    const localSnapshot = buildLeaveStateSnapshot(syncScope);
    const baseSnapshot = cloudSnapshot && typeof cloudSnapshot === "object" ? cloudSnapshot : {};
    const baseState = baseSnapshot.state && typeof baseSnapshot.state === "object" ? baseSnapshot.state : {};
    const localState = localSnapshot.state || {};
    const nextState = {
      ...baseState,
      students: mergeScopedSnapshotItems(baseState.students, localState.students, isStudentInScope, syncScope, (student) => normalizeParticipantCode(student.code || student.studentCode)),
      coaches: mergeScopedSnapshotItems(baseState.coaches, localState.coaches, isCoachInScope, syncScope, (coach) => normalizeParticipantCode(coach.code || coach.coachCode)),
      lessons: mergeScopedSnapshotItems(baseState.lessons, localState.lessons, isLessonInScope, syncScope, (lesson) => String(lesson.id || getLessonCloudMatchKey(lesson))),
      leaveRequests: mergeScopedSnapshotItems(baseState.leaveRequests, localState.leaveRequests, (leave, targetScope) => (
        isLeaveInCloudSyncScope(leave, getLessonById(leave.lessonId), targetScope)
      ), syncScope, (leave) => String(leave.id || leave.lessonId || "")),
      makeupRequests: mergeScopedSnapshotItems(baseState.makeupRequests, localState.makeupRequests, isMakeupRequestInScope, syncScope, (request) => String(request.id || "")),
      coachBlocks: mergeScopedSnapshotItems(baseState.coachBlocks, localState.coachBlocks, isCoachBlockInScope, syncScope, (block) => String(block.id || `${block.coachCode}|${block.startAt}`)),
      compensationTasks: mergeScopedSnapshotItems(baseState.compensationTasks, localState.compensationTasks, isCompensationTaskInScope, syncScope, (task) => String(task.id || ""))
    };
    return {
      ...baseSnapshot,
      version: "1",
      schemaVersion: String(SCHEMA_VERSION),
      coachCode: syncScope.coachCode || baseSnapshot.coachCode || "",
      studentCode: "",
      source: "coachflow-leave",
      sourceVersion: new URLSearchParams(window.location.search).get("v") || "",
      updatedAt: new Date().toISOString(),
      updatedBy: syncScope.studentCode || syncScope.coachCode || "SYSTEM",
      baseUpdatedAt: baseSnapshot.updatedAt || "",
      counts: countSnapshotStateItems(nextState),
      state: nextState
    };
  }

  function applyCloudLeaveStateSnapshot(snapshot, scope = {}) {
    if (!snapshot || !snapshot.state || typeof snapshot.state !== "object") {
      return false;
    }
    const syncScope = getCloudStateScope({
      coachCode: snapshot.coachCode || scope.coachCode,
      studentCode: snapshot.studentCode || scope.studentCode
    });
    if (!syncScope.coachCode && !syncScope.studentCode) {
      return false;
    }
    if (state.cloudSnapshotUpdatedAt && snapshot.updatedAt === state.cloudSnapshotUpdatedAt) {
      return false;
    }

    const incomingState = snapshot.state;
    isApplyingCloudSnapshot = true;
    try {
      mergeScopedStateItems("coaches", incomingState.coaches, isCoachInScope, syncScope, (coach) => normalizeParticipantCode(coach.code || coach.coachCode));
      mergeScopedStateItems("students", incomingState.students, isStudentInScope, syncScope, (student) => normalizeParticipantCode(student.code || student.studentCode));
      mergeScopedStateItems("lessons", incomingState.lessons, isLessonInScope, syncScope, (lesson) => String(lesson.id || getLessonCloudMatchKey(lesson)));
      mergeScopedStateItems("leaveRequests", incomingState.leaveRequests, (leave, targetScope) => (
        isLeaveInCloudSyncScope(leave, getLessonById(leave.lessonId), targetScope)
      ), syncScope, (leave) => String(leave.id || leave.lessonId || ""));
      mergeScopedStateItems("makeupRequests", incomingState.makeupRequests, isMakeupRequestInScope, syncScope, (request) => String(request.id || ""));
      mergeScopedStateItems("coachBlocks", incomingState.coachBlocks, isCoachBlockInScope, syncScope, (block) => String(block.id || `${block.coachCode}|${block.startAt}`));
      mergeScopedStateItems("compensationTasks", incomingState.compensationTasks || [], isCompensationTaskInScope, syncScope, (task) => String(task.id || ""));
      state.cloudSnapshotUpdatedAt = snapshot.updatedAt || new Date().toISOString();
      state.cloudSnapshotUpdatedBy = snapshot.updatedBy || "";
      saveState();
    } finally {
      isApplyingCloudSnapshot = false;
    }
    return true;
  }

  function isCorruptSnapshotError(error) {
    const message = String(error?.message || error || "");
    return /not readable|unreadable|snapshot.*parse|snapshot.*corrupt|cannot parse/i.test(message);
  }

  async function fetchLeaveStateSnapshotFromCloud(scope = {}) {
    if (!getAppsScriptUrl() || leaveStateSnapshotUnsupported) {
      return null;
    }
    const syncScope = getCloudStateScope(scope);
    if (!syncScope.coachCode && !syncScope.studentCode) {
      return null;
    }
    try {
      const result = await callAppsScriptApi("getLeaveStateSnapshot", syncScope, "GET");
      return result?.found ? result.snapshot : null;
    } catch (error) {
      if (isUnsupportedAppsScriptAction(error, "getLeaveStateSnapshot")) {
        leaveStateSnapshotUnsupported = true;
        return null;
      }
      // 雲端那筆 snapshot 損毀（JSON.parse 失敗）時，舊版後端會回 ok:false 讓前端 throw。
      // snapshot 只是狀態 cache，不是真資料 → 視為「沒 snapshot」讓 saveSnapshot 重新覆寫。
      if (isCorruptSnapshotError(error)) {
        addLog?.("[雲端 snapshot] 雲端 snapshot 損毀，視為沒 snapshot 處理（之後寫入會覆蓋掉壞的）。");
        return null;
      }
      throw error;
    }
  }

  async function pullLeaveStateSnapshotFromCloud(scope = {}) {
    const snapshot = await fetchLeaveStateSnapshotFromCloud(scope);
    if (!snapshot) {
      return false;
    }
    return applyCloudLeaveStateSnapshot(snapshot, scope);
  }

  async function saveLeaveStateSnapshotToCloud(scope = {}, options = {}) {
    if (!getAppsScriptUrl() || leaveStateSnapshotUnsupported) {
      return false;
    }
    const syncScope = getCloudStateScope(scope);
    if (!syncScope.coachCode && !syncScope.studentCode) {
      return false;
    }
    let snapshot = buildLeaveStateSnapshot(syncScope);
    if (syncScope.coachCode && !syncScope.studentCode) {
      let cloudSnapshot = null;
      let fetchFailedNonCorrupt = false;
      try {
        cloudSnapshot = await fetchLeaveStateSnapshotFromCloud({
          coachCode: syncScope.coachCode,
          studentCode: ""
        });
      } catch (error) {
        // fetchLeaveStateSnapshotFromCloud 內部已對「損毀」吞掉並回 null。
        // 走到這裡的 throw 是真正的網路錯誤等。為了不再讓壞掉的 snapshot 永遠寫不進新資料，
        // 改成「fetch 失敗」就跳過 merge、直接寫入本地版。
        console.warn("Cloud snapshot fetch before full save failed (will save without merge):", error);
        fetchFailedNonCorrupt = true;
      }
      if (cloudSnapshot?.state) {
        snapshot = mergeLocalScopedStateIntoSnapshot(cloudSnapshot, syncScope);
      } else if (fetchFailedNonCorrupt) {
        // 沒 merge 用的 cloud snapshot，但本地 snapshot 仍可寫，下次別人 fetch 就拿得到新版
      }
    }
    try {
      const result = await callAppsScriptApi("saveLeaveStateSnapshot", {
        snapshot,
        force: Boolean(options.force),
        baseUpdatedAt: snapshot.baseUpdatedAt || state.cloudSnapshotUpdatedAt || ""
      });
      if (result?.conflict && result.snapshot) {
        applyCloudLeaveStateSnapshot(result.snapshot, syncScope);
        return false;
      }
      if (result?.snapshot?.updatedAt) {
        state.cloudSnapshotUpdatedAt = result.snapshot.updatedAt;
        state.cloudSnapshotUpdatedBy = result.snapshot.updatedBy || "";
        localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
      }
      return result?.ok !== false;
    } catch (error) {
      if (isUnsupportedAppsScriptAction(error, "saveLeaveStateSnapshot")) {
        leaveStateSnapshotUnsupported = true;
        return false;
      }
      if (isStorageQuotaError(error)) {
        leaveStateSnapshotUnsupported = true;
        addLog("[雲端同步] Apps Script Properties 儲存配額爆滿，停用本次 session 的 snapshot 上傳。");
        return false;
      }
      throw error;
    }
  }

  async function saveMergedLeaveStateSnapshotToCloud(scope = {}, options = {}) {
    if (!getAppsScriptUrl() || leaveStateSnapshotUnsupported) {
      return false;
    }
    const syncScope = getCloudStateScope(scope);
    if (!syncScope.coachCode) {
      return saveLeaveStateSnapshotToCloud(scope, options);
    }
    if (!syncScope.studentCode) {
      return saveLeaveStateSnapshotToCloud({ coachCode: syncScope.coachCode, studentCode: "" }, options);
    }

    let cloudSnapshot = null;
    try {
      cloudSnapshot = await fetchLeaveStateSnapshotFromCloud({
        coachCode: syncScope.coachCode,
        studentCode: ""
      });
    } catch (error) {
      // 不再因 fetch 失敗就 return false（會讓壞 snapshot 永遠寫不掉）。
      // cloudSnapshot 維持 null，mergeLocalScopedStateIntoSnapshot 在 null 時應視為從零開始。
      console.warn("Cloud snapshot fetch before scoped save failed (proceed without merge):", error);
    }
    let mergedSnapshot = mergeLocalScopedStateIntoSnapshot(cloudSnapshot, syncScope);
    const savePayload = (snapshot) => ({
      snapshot,
      force: Boolean(options.force),
      baseUpdatedAt: snapshot.baseUpdatedAt || ""
    });

    try {
      let result = await callAppsScriptApi("saveLeaveStateSnapshot", savePayload(mergedSnapshot));
      if (result?.conflict && result.snapshot) {
        mergedSnapshot = mergeLocalScopedStateIntoSnapshot(result.snapshot, syncScope);
        result = await callAppsScriptApi("saveLeaveStateSnapshot", savePayload(mergedSnapshot));
      }
      if (result?.snapshot?.updatedAt) {
        state.cloudSnapshotUpdatedAt = result.snapshot.updatedAt;
        state.cloudSnapshotUpdatedBy = result.snapshot.updatedBy || "";
        localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
      }
      return result?.ok !== false && !result?.conflict;
    } catch (error) {
      if (isUnsupportedAppsScriptAction(error, "saveLeaveStateSnapshot")) {
        leaveStateSnapshotUnsupported = true;
        return false;
      }
      throw error;
    }
  }

  async function saveStudentScopedStateToCloud(studentCode, coachCode, reason = "") {
    const normalizedStudentCode = normalizeParticipantCode(studentCode || activeStudentCode || "");
    const normalizedCoachCode = normalizeParticipantCode(coachCode || activeCoachCode || "");
    if (!normalizedStudentCode || !normalizedCoachCode) {
      return false;
    }
    try {
      const saved = await saveMergedLeaveStateSnapshotToCloud({
        coachCode: normalizedCoachCode,
        studentCode: normalizedStudentCode
      }, { reason });
      if (saved) {
        addLog(`[雲端同步] 已同步 ${normalizedStudentCode} 的請假/補課資料。`);
        saveState();
      }
      return saved;
    } catch (error) {
      addLog(`[雲端同步] ${normalizedStudentCode} 請假/補課資料同步失敗：${String(error?.message || error)}`);
      saveState();
      return false;
    }
  }

  function queueCloudStateSnapshotUpload(reason = "") {
    if (
      isApplyingCloudSnapshot ||
      isUploadingCloudSnapshot ||
      leaveStateSnapshotUnsupported ||
      !state.cloudSnapshotUpdatedAt ||
      !activeCoachCode ||
      activeStudentCode ||
      !el.coachCloudUploadBtn ||
      isCoachReadOnlyMode() ||
      !getAppsScriptUrl()
    ) {
      return;
    }
    window.clearTimeout(cloudSnapshotUploadTimer);
    cloudSnapshotUploadTimer = window.setTimeout(() => {
      saveLeaveStateSnapshotToCloud({ coachCode: activeCoachCode }, { reason })
        .catch((error) => {
          // 把 Error 內容明確攤開，避免 console 只印出 Error {} 看不出原因。
          const detail = {
            name: error?.name || "(unknown)",
            message: error?.message || String(error),
            stack: error?.stack || "(no stack)"
          };
          console.warn("Cloud leave snapshot auto upload failed:", detail);
          addLog(`[雲端同步] 自動上傳 snapshot 失敗：${detail.name} / ${detail.message}`);
          // 一旦遇到配額爆滿，本 session 不再嘗試 snapshot 上傳（避免一直重撞）。
          // 單筆 saveLeaveRecord / saveBillingProfile 路徑不受影響。教練可在 Apps
          // Script 編輯器清理 LEAVE_STATE_SNAPSHOT_V1_* 後重新整理頁面恢復。
          if (isStorageQuotaError(error)) {
            leaveStateSnapshotUnsupported = true;
            addLog("[雲端同步] 偵測到 Apps Script 儲存配額爆滿，本次 session 不再嘗試 snapshot 上傳。請至 Apps Script 編輯器清理屬性後刷新頁面。");
            notifyUser("雲端 snapshot 儲存空間不足，自動同步已停用。請聯絡管理員清理 Apps Script Properties。", "warning");
          }
        });
    }, 1800);
  }

  async function pushScopedCloudRecords(scope = {}) {
    const syncScope = getCloudStateScope(scope);
    const leaves = state.leaveRequests.filter((leave) => (
      isLeaveInCloudSyncScope(leave, getLessonById(leave.lessonId), syncScope)
    ));
    let leaveCount = 0;
    for (const leave of leaves) {
      const lesson = getLessonById(leave.lessonId);
      if (lesson && await pushCloudLeaveRecord(leave, lesson)) {
        leaveCount += 1;
      }
    }

    const students = state.students.filter((student) => isStudentInScope(student, syncScope));
    let billingCount = 0;
    for (const student of students) {
      if (await pushCloudBillingProfile(student.code, "manual_snapshot_upload")) {
        billingCount += 1;
      }
    }

    // 完整課表雲端化的第一次灌入主入口：把這台「開過完整課表」裝置的完整課表
    // 一次推上雲。之後請假/改狀態/日曆同步會自動增量 push，雲端持續保鮮。
    // 後端未部署時 pushCloudLessons graceful 回 false → lessonsPushed 記 0。
    let lessonsPushed = 0;
    try {
      const pushed = await pushCloudLessons(syncScope);
      lessonsPushed = pushed === false ? 0 : pushed;
    } catch (error) {
      addLog(`[雲端課表] 上傳完整課表失敗：${String(error?.message || error)}`);
    }

    return { leaveCount, billingCount, lessonsPushed };
  }

  function renderCloudSnapshotControls() {
    const hasCoach = Boolean(activeCoachCode);
    [el.coachCloudUploadBtn, el.coachCloudDownloadBtn].filter(Boolean).forEach((button) => {
      button.disabled = !hasCoach || isCoachReadOnlyMode() || isUploadingCloudSnapshot;
    });
    if (el.coachCloudSyncText) {
      const updatedText = state.cloudSnapshotUpdatedAt
        ? `雲端資料最後同步：${formatDateTime(state.cloudSnapshotUpdatedAt)}`
        : "尚未把這台電腦的請假系統資料上傳到雲端。";
      el.coachCloudSyncText.textContent = hasCoach ? updatedText : "請先登入教練代碼。";
    }
  }

  // 列出本機所有 pendingCloudSync 的請假，組成持續可見的紅色 banner。
  // 學生請假時若雲端失敗會留下 pendingCloudSync，banner 提醒並提供「立即重試」按鈕；
  // 背景的 retryUnsyncedLocalLeaves 也會在 focus/visibility 時自動跑一次（silent）。
  //
  // Scope 規則：若有 activeStudentCode → 只看那位學生的；否則退回 activeCoachCode 的全部；
  // 都沒有就空。早期版本用 student OR coach 的 OR 條件，會在「學生測試頁」這種同時有
  // student/coach session 的情境，把同教練底下其他學生的請假紀錄也顯示給學生看。
  function renderPendingSyncBanner() {
    if (!el.pendingLeaveBanner) {
      return;
    }
    const all = (state.leaveRequests || []).filter((leave) => (
      leave && !leave.revokedAt && leave.pendingCloudSync && !leave.cloudSyncedAt
    ));
    let visible;
    if (activeStudentCode) {
      visible = all.filter((leave) => leave.studentCode === activeStudentCode);
    } else if (activeCoachCode) {
      visible = all.filter((leave) => leave.coachCode === activeCoachCode);
    } else {
      visible = [];
    }
    if (!visible.length) {
      el.pendingLeaveBanner.hidden = true;
      return;
    }
    const lessonsText = visible
      .map((leave) => {
        const lesson = getLessonById(leave.lessonId);
        const when = lesson ? formatDateTime(lesson.startAt) : "(找不到課程)";
        // 學生視角不需要再印自己的 student code，省版面；教練視角則保留以區分學生
        return activeStudentCode && !activeCoachCode
          ? when
          : `${leave.studentCode}@${when}`;
      })
      .join("、");
    if (el.pendingLeaveText) {
      el.pendingLeaveText.textContent = `有 ${visible.length} 筆請假尚未同步到雲端：${lessonsText}。系統會在頁面聚焦/重新整理時自動重試；要立刻試的話按右邊按鈕。`;
    }
    el.pendingLeaveBanner.hidden = false;
  }

  // ----- 手動新增學生上課（教練端） -----
  // 對應 leave-coach-sandbox.html 的「手動新增學生上課」card：
  // 不靠 Google 日曆主動加課，由教練從系統 push 一筆 lesson 並同步建出 GCal 事件。
  // sourceType 用 "REGULAR" 跟 GCal 來源等價，後續請假/補課/扣堂流程全自動沿用。
  function populateManualLessonStudentSelect() {
    if (!el.manualLessonStudent) {
      return;
    }
    const select = el.manualLessonStudent;
    if (!activeCoachCode) {
      select.innerHTML = "<option value=\"\">請先登入教練</option>";
      select.disabled = true;
      return;
    }
    const previousValue = select.value;
    const students = (state.students || [])
      .filter((student) => student && student.coachCode === activeCoachCode && !isHiddenChargeStudent(student))
      .sort((a, b) => String(a.name || a.code).localeCompare(String(b.name || b.code), "zh-Hant"));
    if (!students.length) {
      select.innerHTML = "<option value=\"\">這個教練底下還沒有學生</option>";
      select.disabled = true;
      return;
    }
    select.innerHTML = "<option value=\"\">請選擇學生</option>" + students.map((student) => (
      `<option value="${student.code}">${student.name || student.code}（${student.code}）</option>`
    )).join("");
    select.disabled = false;
    if (previousValue && students.some((student) => student.code === previousValue)) {
      select.value = previousValue;
    }
  }

  async function addManualStudentLesson() {
    if (!requireCoachWriteAccess()) {
      return;
    }
    if (!activeCoachCode) {
      alert("請先登入教練帳號。");
      return;
    }
    if (!el.manualLessonStudent || !el.manualLessonDate || !el.manualLessonTime) {
      return;
    }
    const studentCode = String(el.manualLessonStudent.value || "").trim();
    const dateKey = String(el.manualLessonDate.value || "").trim();
    const timeText = String(el.manualLessonTime.value || "").trim();
    if (!studentCode) {
      alert("請選擇學生。");
      return;
    }
    if (!dateKey || !timeText) {
      alert("請輸入日期與開始時間。");
      return;
    }
    const student = getStudentByCode(studentCode);
    if (!student || student.coachCode !== activeCoachCode) {
      alert("找不到該學生，或學生不屬於目前登入的教練。");
      return;
    }

    const startAtDate = makeTaipeiDateTime(dateKey, timeText);
    if (!Number.isFinite(startAtDate.getTime())) {
      alert("日期或時間格式不正確。");
      return;
    }
    const startAtIso = startAtDate.toISOString();

    // 同一學生在同一教練底下同一時刻已經有 lesson record → 擋下避免重複。
    // 注意：月曆只顯示 calendarOccupied === true 的課程，所以一個 leave-normal 的
    // lesson（學生已請假）月曆看不到，但 state 還在 → 教練容易誤以為「沒有課」而想
    // 重複新增。alert 必須點明衝突 lesson 的時間、狀態、id，讓教練知道從哪裡處理。
    const STATUS_LABEL = {
      "scheduled": "已排課",
      "leave-normal": "學生已請假",
      "coach-leave": "教練停課",
      "calendar-removed": "Google 日曆已移除",
      "temporary-leave": "臨時請假（已扣堂）",
      "no-show": "未到課（已扣堂）",
      "major-case": "重大急事（不扣堂）",
      "completed": "已上課"
    };
    const duplicate = state.lessons.find((lesson) => (
      lesson &&
      lesson.coachCode === activeCoachCode &&
      lesson.studentCode === studentCode &&
      new Date(lesson.startAt || "").toISOString() === startAtIso &&
      lesson.attendanceStatus !== "calendar-removed"
    ));
    if (duplicate) {
      const status = String(duplicate.attendanceStatus || "");
      const statusText = STATUS_LABEL[status] || status || "(未知狀態)";
      const whenText = formatDateTime(duplicate.startAt);
      // calendarOccupied === false 表示 slot 已被 leave / coach-leave 釋出，月曆會隱藏 →
      // 教練多半看不到，特別指引他到「學生請假紀錄」或「教練請假時段」處理
      const hiddenHint = duplicate.calendarOccupied === false
        ? "（此狀態的課程不會顯示在教練月曆，請到「學生請假紀錄」或「教練請假/停課時段」面板找到對應紀錄處理）"
        : "";
      alert(`這名學生在 ${whenText} 已有課程紀錄，不重複新增。\n狀態：${statusText}\n課程 ID：${duplicate.id}${hiddenHint ? "\n" + hiddenHint : ""}`);
      return;
    }

    const alreadyAttended = Boolean(el.manualLessonAlreadyAttended?.checked);
    const lesson = makeLesson(newId("L"), studentCode, activeCoachCode, dateKey, timeText);
    if (alreadyAttended) {
      // 仿 markAutoCompletedLessonsForBilling 把這堂直接記成已扣堂
      lesson.charged = true;
      lesson.completedAt = new Date().toISOString();
      lesson.completedBy = activeCoachCode;
    }
    lesson.updatedAt = new Date().toISOString();
    state.lessons.push(lesson);
    addLog(`教練 ${activeCoachCode} 手動新增學生 ${studentCode} 之上課：${formatDateTime(lesson.startAt)}（${alreadyAttended ? "已上完、立即扣堂" : "未來課程、不扣堂"}）。`);
    saveState();
    renderAll();
    // 完整課表雲端化：手動新增的課同步推上雲，主系統即時看得到。
    pushCloudLessonsQuietly({ coachCode: activeCoachCode, studentCode });
    notifyUser(`已建立 ${student.name || studentCode} 的上課（${formatDateTime(lesson.startAt)}），正在同步 Google 日曆...`, "info");

    // 同步建 GCal 事件。非 strict：失敗會走 compensation 任務排隊重試，本機 lesson 仍保留
    let calendarEventId = "";
    try {
      calendarEventId = await tryCreateCalendarEventForLesson(lesson, "coach_manual_add_lesson", false);
    } catch (error) {
      addLog(`[Google日曆] 手動新增 ${lesson.id} 建立事件失敗：${String(error?.message || error)}`);
    }
    saveState();
    renderAll();

    if (calendarEventId && !isGeneratedLocalCalendarEventId(calendarEventId)) {
      notifyUser(`已建立 ${student.name || studentCode} 的上課 + Google 日曆事件。${alreadyAttended ? "已計入扣堂。" : ""}`, "success");
    } else {
      notifyUser(`已建立 ${student.name || studentCode} 的上課，但 Google 日曆事件尚未確認；補償任務會自動重試。`, "warning");
    }

    // 清空 form（保留學生選擇方便連續加同學生不同時段）
    if (el.manualLessonAlreadyAttended) {
      el.manualLessonAlreadyAttended.checked = false;
    }
  }

  async function uploadLocalLeaveStateToCloud() {
    if (!activeCoachCode) {
      alert("請先載入教練資料。");
      return;
    }
    try {
      await syncCloudBillingProfiles({ coachCode: activeCoachCode, studentCode: "" });
      renderBillingPanels();
    } catch (error) {
      console.warn("Cloud billing pre-confirm sync failed:", error);
    }
    const snapshot = buildLeaveStateSnapshot({ coachCode: activeCoachCode, studentCode: "" });
    const confirmed = window.confirm(
      [
        "要把這台電腦目前的請假系統資料上傳到雲端嗎？",
        "",
        `課程：${snapshot.counts.lessons} 筆`,
        `請假：${snapshot.counts.leaveRequests} 筆`,
        `補課：${snapshot.counts.makeupRequests} 筆`,
        `學生：${snapshot.counts.students} 位`,
        "",
        "上傳後，手機和其他電腦會以這份雲端資料為準。"
      ].join("\n")
    );
    if (!confirmed) {
      return;
    }

    isUploadingCloudSnapshot = true;
    const uploadButtonText = el.coachCloudUploadBtn?.textContent || "";
    if (el.coachCloudUploadBtn) {
      el.coachCloudUploadBtn.disabled = true;
      el.coachCloudUploadBtn.textContent = "上傳中...";
    }
    try {
      notifyUser("正在把這台電腦的請假系統資料同步到雲端...", "info");
      try {
        await syncCloudBillingProfiles({ coachCode: activeCoachCode, studentCode: "" });
      } catch (error) {
        console.warn("Cloud billing pre-upload sync failed:", error);
      }
      const counts = await pushScopedCloudRecords({ coachCode: activeCoachCode, studentCode: "" });
      const saved = await saveLeaveStateSnapshotToCloud({ coachCode: activeCoachCode, studentCode: "" }, { force: true });
      if (!saved) {
        notifyUser("雲端快照端點尚未更新，請先重新部署 Apps Script 後再按一次。", "warning");
        return;
      }
      renderAll();
      notifyUser(`已上傳到雲端：請假 ${counts.leaveCount} 筆、扣課資料 ${counts.billingCount} 位、完整課表 ${counts.lessonsPushed} 堂，完整快照已更新。`, "success");
    } catch (error) {
      notifyUser(`雲端同步失敗：${String(error?.message || error)}`, "warning");
    } finally {
      isUploadingCloudSnapshot = false;
      if (el.coachCloudUploadBtn) {
        el.coachCloudUploadBtn.disabled = false;
        el.coachCloudUploadBtn.textContent = uploadButtonText || "把這台電腦資料上傳到雲端";
      }
      renderCloudSnapshotControls();
    }
  }

  async function downloadCloudLeaveStateToLocal() {
    if (!activeCoachCode) {
      alert("請先載入教練資料。");
      return;
    }
    const confirmed = window.confirm("要用雲端請假系統資料覆蓋這台裝置目前的本機資料嗎？");
    if (!confirmed) {
      return;
    }
    const downloadButtonText = el.coachCloudDownloadBtn?.textContent || "";
    if (el.coachCloudDownloadBtn) {
      el.coachCloudDownloadBtn.disabled = true;
      el.coachCloudDownloadBtn.textContent = "載入中...";
    }
    try {
      const changed = await pullLeaveStateSnapshotFromCloud({ coachCode: activeCoachCode });
      if (changed) {
        renderAll();
        notifyUser("已用雲端請假系統資料更新這台裝置。", "success");
      } else {
        notifyUser("目前沒有可套用的雲端請假系統資料。", "info");
      }
    } catch (error) {
      notifyUser(`雲端載入失敗：${String(error?.message || error)}`, "warning");
    } finally {
      if (el.coachCloudDownloadBtn) {
        el.coachCloudDownloadBtn.disabled = false;
        el.coachCloudDownloadBtn.textContent = downloadButtonText || "從雲端重新載入";
      }
      renderCloudSnapshotControls();
    }
  }

  function getStudentBillingUpdatedAt(student) {
    const logTimes = Array.isArray(student?.chargeReminderLogs)
      ? student.chargeReminderLogs.map((item) => item?.sentAt)
      : [];
    const times = [
      student?.billingUpdatedAt,
      student?.chargeStartCountUpdatedAt,
      student?.paymentConfirmedAt,
      student?.emailUpdatedAt,
      ...logTimes
    ]
      .map((value) => new Date(value || "").getTime())
      .filter((time) => Number.isFinite(time));
    return times.length ? new Date(Math.max(...times)).toISOString() : "";
  }

  function touchStudentBillingProfile(student, updatedBy) {
    if (!student) {
      return "";
    }
    const now = new Date().toISOString();
    student.billingUpdatedAt = now;
    student.billingUpdatedBy = normalizeParticipantCode(updatedBy || activeCoachCode || activeStudentCode || "SYSTEM");
    return now;
  }

  function isManualChargeStartFallbackSource(updatedBy) {
    const source = normalizeParticipantCode(updatedBy || "");
    if (!source || source === "SYSTEM") {
      return false;
    }
    if ((state.coaches || []).some((coach) => normalizeParticipantCode(coach?.code) === source)) {
      return true;
    }
    return /^(MO|CH)\d{3,}$/.test(source);
  }

  function getChargeStartBaselineAt(student) {
    const explicitBaseline = String(student?.chargeStartCountUpdatedAt || "").trim();
    if (hasValidDateValue(explicitBaseline)) {
      return explicitBaseline;
    }
    if (toNonNegativeInt(student?.chargeStartCount, 0) <= 0) {
      return "";
    }
    const fallbackBaseline = String(student?.billingUpdatedAt || "").trim();
    const fallbackBy = normalizeParticipantCode(student?.billingUpdatedBy || "");
    if (!isManualChargeStartFallbackSource(fallbackBy) || !hasValidDateValue(fallbackBaseline)) {
      return "";
    }
    const fallbackTime = new Date(fallbackBaseline).getTime();
    const paymentTime = new Date(student?.paymentConfirmedAt || "").getTime();
    if (Number.isFinite(paymentTime) && Math.abs(fallbackTime - paymentTime) <= 5000) {
      return "";
    }
    return fallbackBaseline;
  }

  function getChargeStartBaselineBy(student) {
    const explicitBy = normalizeParticipantCode(student?.chargeStartCountUpdatedBy || "");
    if (explicitBy) {
      return explicitBy;
    }
    return getChargeStartBaselineAt(student) ? normalizeParticipantCode(student?.billingUpdatedBy || "") : "";
  }

  function isLessonAfterChargeStartBaseline(lesson, student) {
    const baselineAt = getChargeStartBaselineAt(student);
    if (!baselineAt) {
      return true;
    }
    const lessonTime = new Date(lesson?.startAt || "").getTime();
    if (!Number.isFinite(lessonTime)) {
      return false;
    }
    const baselineTime = new Date(baselineAt).getTime();
    if (!Number.isFinite(baselineTime)) {
      return true;
    }
    // 以「日」為粒度比較：baseline 當天或之後的課都算系統內扣堂。
    // 過去用嚴格時刻比較，會造成「教練在當天 10:02 設 baseline，但學生 09:00 的課
    // 被當成在 baseline 之前」→ 系統內扣堂少算 1。教練設定 baseline 的意圖通常是
    // 「從今天起系統開始計算」，同一天的課應該算進來。
    //
    // 日界線用「台北時區」算（Codex review #38）：原本 setHours(0,0,0,0) 取的是
    // 瀏覽器本地午夜，教練/學生若在非台北裝置開系統，日界線會偏掉、把前一天傍晚
    // 的課誤算進當期。台灣無 DST、固定 +08:00，用 makeTaipeiDateTime 鎖死台北日。
    const baselineDayStart = makeTaipeiDateTime(getDateKeyInTaipei(new Date(baselineTime)), "00:00");
    return lessonTime >= baselineDayStart.getTime();
  }

  function isStorageQuotaError(error) {
    const message = String(error?.message || error || "");
    return /超出.*資源.*配額|配額上限|storage.*quota|exceeded.*quota|quota.*exceeded/i.test(message);
  }

  function isUnsupportedAppsScriptAction(error, action) {
    const text = String(error?.message || error || "");
    return /Unsupported action/i.test(text);
  }

  function buildCloudBillingProfile(student, triggerSource = "") {
    if (!student) {
      return null;
    }
    const studentCode = normalizeParticipantCode(student.code || student.studentCode);
    if (!studentCode) {
      return null;
    }
    const coachCode = normalizeParticipantCode(student.coachCode || activeCoachCode || "");
    const updatedAt = getStudentBillingUpdatedAt(student) || new Date().toISOString();
    const stats = getStudentChargeStats(studentCode);
    const billingCycle = getStudentBillingCycle(stats);
    return {
      studentCode,
      coachCode,
      studentName: String(student.name || ""),
      email: normalizeEmailList(student.email).join(", "),
      emailUpdatedAt: String(student.emailUpdatedAt || ""),
      emailUpdatedBy: normalizeParticipantCode(student.emailUpdatedBy || ""),
      chargeStartCount: toNonNegativeInt(student.chargeStartCount, 0),
      chargeStartCountUpdatedAt: getChargeStartBaselineAt(student),
      chargeStartCountUpdatedBy: getChargeStartBaselineBy(student),
      paidThroughCount: toNonNegativeInt(student.paidThroughCount, 0),
      paymentStatus: normalizePaymentStatus(student.paymentStatus),
      paymentNote: String(student.paymentNote || ""),
      paymentConfirmedAt: String(student.paymentConfirmedAt || ""),
      paymentConfirmedBy: normalizeParticipantCode(student.paymentConfirmedBy || ""),
      chargeReminderLogs: normalizeChargeReminderLogs(student.chargeReminderLogs),
      chargeReminderStep: CHARGE_REMINDER_STEP,
      systemChargedCount: billingCycle.systemChargedCount,
      totalChargedCount: billingCycle.totalChargedCount,
      currentCycleChargedCount: billingCycle.currentCycleChargedCount,
      remainingToNextPayment: billingCycle.remainingToNextPayment,
      nextPaymentDueCount: billingCycle.nextPaymentDueCount,
      effectivePaymentStatus: billingCycle.effectivePaymentStatus,
      updatedAt,
      updatedBy: normalizeParticipantCode(student.billingUpdatedBy || activeCoachCode || activeStudentCode || "SYSTEM"),
      triggerSource: String(triggerSource || "")
    };
  }

  function buildCloudBillingRecord(student, triggerSource = "") {
    const profile = buildCloudBillingProfile(student, triggerSource);
    if (!profile) {
      return null;
    }
    return {
      id: `${CLOUD_BILLING_RECORD_PREFIX}${profile.studentCode}`,
      lessonId: "",
      lessonKey: `${CLOUD_BILLING_RECORD_PREFIX}${profile.studentCode}`,
      calendarEventId: "",
      studentCode: profile.studentCode,
      coachCode: profile.coachCode,
      lessonStartAt: profile.updatedAt,
      type: CLOUD_BILLING_RECORD_TYPE,
      submittedAt: profile.updatedAt,
      submittedBy: profile.updatedBy,
      submittedByRole: "coach",
      makeupEligible: false,
      emailNoticeStatus: JSON.stringify(profile),
      emailNoticeAt: profile.updatedAt,
      revokedAt: "",
      revokedBy: ""
    };
  }

  function extractCloudBillingProfile(record) {
    if (!record || String(record.type || "") !== CLOUD_BILLING_RECORD_TYPE) {
      return null;
    }
    let parsed = {};
    try {
      parsed = JSON.parse(String(record.emailNoticeStatus || "{}"));
    } catch (error) {
      parsed = {};
    }
    return {
      ...parsed,
      studentCode: normalizeParticipantCode(parsed.studentCode || record.studentCode),
      coachCode: normalizeParticipantCode(parsed.coachCode || record.coachCode),
      updatedAt: String(parsed.updatedAt || record.updatedAt || record.submittedAt || record.emailNoticeAt || ""),
      updatedBy: normalizeParticipantCode(parsed.updatedBy || record.submittedBy || "")
    };
  }

  function getCloudBillingTimestamp(profile) {
    const candidates = [
      profile?.updatedAt,
      profile?.billingUpdatedAt,
      profile?.chargeStartCountUpdatedAt,
      profile?.paymentConfirmedAt,
      profile?.emailUpdatedAt
    ];
    const logTimes = Array.isArray(profile?.chargeReminderLogs)
      ? profile.chargeReminderLogs.map((item) => item?.sentAt)
      : [];
    const times = [...candidates, ...logTimes]
      .map((value) => new Date(value || "").getTime())
      .filter((time) => Number.isFinite(time));
    return times.length ? new Date(Math.max(...times)).toISOString() : "";
  }

  function getScopedBillingStudents(scope = {}) {
    const { coachCode, studentCode } = getCloudStateScope(scope);
    return (state.students || []).filter((student) => {
      if (isKnownFakeStudentRecord(student)) {
        return false;
      }
      const currentStudentCode = normalizeParticipantCode(student?.code);
      const currentCoachCode = normalizeParticipantCode(student?.coachCode);
      return (!studentCode || currentStudentCode === studentCode) &&
        (!coachCode || currentCoachCode === coachCode);
    });
  }

  // 自動測試/評分腳本會直接對正式後端灌入假學生的 billing profile，請假系統
  // 同步時就會自動把它們建成學生（WF1 Student / Grader Student / 測試學員…）。
  // 這個指紋用來在同步時略過、不落地。真學生（中文姓名或 Zoe/Elsie/Matt）不命中。
  function looksLikeTestParticipant(value) {
    const s = String(value === null || value === undefined ? "" : value).trim();
    if (!s) {
      return false;
    }
    return /grader|workflow|測試|週末|流程|學員|學生|student|\bwf\d|\bw2s?\b|\bw5\b|\bw6\b|wflow|grd|\bgrade\b|\bsync\b|\beval\b|wknd|wkend|weekend|\btest\b|\bqa\b/i.test(s);
  }

  function applyCloudBillingProfile(profile) {
    if (!profile) {
      return false;
    }
    // 略過自動測試假學生的 billing profile，避免請假系統自動建出測試學生。
    if (looksLikeTestParticipant(profile.studentName || profile.name)) {
      return false;
    }
    const studentCode = normalizeParticipantCode(profile.studentCode || profile.code);
    if (!studentCode) {
      return false;
    }
    const coachCode = normalizeParticipantCode(profile.coachCode || activeCoachCode || getStudentByCode(studentCode)?.coachCode || state.coaches[0]?.code || "CH001");
    const student = getStudentByCode(studentCode) || ensureStudentProfile(studentCode, coachCode, {
      studentName: profile.studentName || profile.name || "",
      studentEmail: profile.email || "",
      silentMode: true
    });
    if (!student) {
      return false;
    }

    const cloudUpdatedAt = getCloudBillingTimestamp(profile);
    const localUpdatedAt = getStudentBillingUpdatedAt(student);
    const cloudTime = new Date(cloudUpdatedAt || "").getTime();
    const localTime = new Date(localUpdatedAt || "").getTime();
    if (Number.isFinite(cloudTime) && Number.isFinite(localTime) && cloudTime + 2000 < localTime) {
      return false;
    }

    // 用 emailUpdatedAt timestamp 判斷誰新誰舊。雲端新就用雲端、本地新就保留本地。
    // 之前無條件「雲端非 undefined 就覆蓋」會把使用者剛輸入但 push 失敗的 email
    // 用雲端舊版（包括空字串、或更早的 default fallback）蓋掉，已上線案例：
    // 教練幫學生輸入 email 後消失、變回 hsnu115023@gmail.com 或空字串。
    const cloudEmailRaw = profile.email === undefined ? null : normalizeEmailList(profile.email).join(", ");
    const cloudEmailTime = new Date(profile.emailUpdatedAt || 0).getTime();
    const localEmailTime = new Date(student.emailUpdatedAt || 0).getTime();
    const localEmailNewer = Number.isFinite(localEmailTime) && localEmailTime > 0
      && (!Number.isFinite(cloudEmailTime) || localEmailTime > cloudEmailTime + 2000);
    let mergedEmail;
    let mergedEmailUpdatedAt;
    let mergedEmailUpdatedBy;
    if (cloudEmailRaw === null || localEmailNewer) {
      // 雲端沒帶或本地較新 → 保留本地
      mergedEmail = String(student.email || "");
      mergedEmailUpdatedAt = String(student.emailUpdatedAt || profile.emailUpdatedAt || "");
      mergedEmailUpdatedBy = normalizeParticipantCode(student.emailUpdatedBy || profile.emailUpdatedBy || "");
    } else {
      mergedEmail = cloudEmailRaw;
      mergedEmailUpdatedAt = String(profile.emailUpdatedAt || student.emailUpdatedAt || "");
      mergedEmailUpdatedBy = normalizeParticipantCode(profile.emailUpdatedBy || student.emailUpdatedBy || "");
    }

    const next = {
      ...student,
      coachCode: coachCode || student.coachCode,
      name: String(profile.studentName || profile.name || "").trim() || student.name,
      email: mergedEmail,
      emailUpdatedAt: mergedEmailUpdatedAt,
      emailUpdatedBy: mergedEmailUpdatedBy,
      chargeStartCount: profile.chargeStartCount === undefined ? toNonNegativeInt(student.chargeStartCount, 0) : toNonNegativeInt(profile.chargeStartCount, 0),
      chargeStartCountUpdatedAt: String(profile.chargeStartCountUpdatedAt || student.chargeStartCountUpdatedAt || ""),
      chargeStartCountUpdatedBy: normalizeParticipantCode(profile.chargeStartCountUpdatedBy || student.chargeStartCountUpdatedBy || ""),
      paidThroughCount: profile.paidThroughCount === undefined ? toNonNegativeInt(student.paidThroughCount, 0) : toNonNegativeInt(profile.paidThroughCount, 0),
      paymentStatus: profile.paymentStatus === undefined ? normalizePaymentStatus(student.paymentStatus) : normalizePaymentStatus(profile.paymentStatus),
      paymentNote: profile.paymentNote === undefined ? String(student.paymentNote || "") : String(profile.paymentNote || ""),
      paymentConfirmedAt: String(profile.paymentConfirmedAt || student.paymentConfirmedAt || ""),
      paymentConfirmedBy: normalizeParticipantCode(profile.paymentConfirmedBy || student.paymentConfirmedBy || ""),
      chargeReminderLogs: profile.chargeReminderLogs === undefined
        ? normalizeChargeReminderLogs(student.chargeReminderLogs)
        : normalizeChargeReminderLogs(profile.chargeReminderLogs),
      billingUpdatedAt: cloudUpdatedAt || localUpdatedAt || "",
      billingUpdatedBy: normalizeParticipantCode(profile.updatedBy || profile.billingUpdatedBy || student.billingUpdatedBy || "")
    };
    const before = JSON.stringify({
      coachCode: student.coachCode,
      name: student.name,
      email: student.email,
      emailUpdatedAt: student.emailUpdatedAt,
      emailUpdatedBy: student.emailUpdatedBy,
      chargeStartCount: student.chargeStartCount,
      chargeStartCountUpdatedAt: student.chargeStartCountUpdatedAt,
      chargeStartCountUpdatedBy: student.chargeStartCountUpdatedBy,
      paidThroughCount: student.paidThroughCount,
      paymentStatus: student.paymentStatus,
      paymentNote: student.paymentNote,
      paymentConfirmedAt: student.paymentConfirmedAt,
      paymentConfirmedBy: student.paymentConfirmedBy,
      chargeReminderLogs: normalizeChargeReminderLogs(student.chargeReminderLogs),
      billingUpdatedAt: student.billingUpdatedAt,
      billingUpdatedBy: student.billingUpdatedBy
    });
    const after = JSON.stringify({
      coachCode: next.coachCode,
      name: next.name,
      email: next.email,
      emailUpdatedAt: next.emailUpdatedAt,
      emailUpdatedBy: next.emailUpdatedBy,
      chargeStartCount: next.chargeStartCount,
      chargeStartCountUpdatedAt: next.chargeStartCountUpdatedAt,
      chargeStartCountUpdatedBy: next.chargeStartCountUpdatedBy,
      paidThroughCount: next.paidThroughCount,
      paymentStatus: next.paymentStatus,
      paymentNote: next.paymentNote,
      paymentConfirmedAt: next.paymentConfirmedAt,
      paymentConfirmedBy: next.paymentConfirmedBy,
      chargeReminderLogs: next.chargeReminderLogs,
      billingUpdatedAt: next.billingUpdatedAt,
      billingUpdatedBy: next.billingUpdatedBy
    });
    if (before === after) {
      return false;
    }
    Object.assign(student, next);
    return true;
  }

  async function listCloudBillingProfiles(scope = {}) {
    const payload = getCloudStateScope(scope);
    const listFallbackProfiles = async () => {
      const fallbackResult = await callAppsScriptApi("listLeaveRecords", payload, "GET");
      return (Array.isArray(fallbackResult?.records) ? fallbackResult.records : [])
        .map(extractCloudBillingProfile)
        .filter(Boolean);
    };
    try {
      const result = await callAppsScriptApi("listBillingProfiles", payload, "GET");
      return Array.isArray(result?.profiles) ? result.profiles : await listFallbackProfiles();
    } catch (error) {
      if (!isUnsupportedAppsScriptAction(error, "listBillingProfiles")) {
        throw error;
      }
      return listFallbackProfiles();
    }
  }

  async function pushCloudBillingProfile(studentCode, triggerSource = "") {
    if (!getAppsScriptUrl()) {
      return false;
    }
    const student = getStudentByCode(studentCode);
    if (!student) {
      return false;
    }
    const profile = buildCloudBillingProfile(student, triggerSource);
    if (!profile) {
      return false;
    }
    try {
      const result = await callAppsScriptApi("saveBillingProfile", { profile });
      if (result?.profile) {
        applyCloudBillingProfile(result.profile);
      }
      return result?.ok !== false;
    } catch (error) {
      if (!isUnsupportedAppsScriptAction(error, "saveBillingProfile")) {
        throw error;
      }
      const existingResult = await callAppsScriptApi("listLeaveRecords", {
        coachCode: profile.coachCode,
        studentCode: profile.studentCode
      }, "GET");
      const existingProfile = (Array.isArray(existingResult?.records) ? existingResult.records : [])
        .map(extractCloudBillingProfile)
        .filter(Boolean)
        .sort((a, b) => new Date(getCloudBillingTimestamp(b) || 0) - new Date(getCloudBillingTimestamp(a) || 0))[0];
      const existingTime = new Date(getCloudBillingTimestamp(existingProfile) || "").getTime();
      const incomingTime = new Date(getCloudBillingTimestamp(profile) || "").getTime();
      if (Number.isFinite(existingTime) && Number.isFinite(incomingTime) && existingTime > incomingTime + 2000) {
        applyCloudBillingProfile(existingProfile);
        return false;
      }
      const record = buildCloudBillingRecord(student, triggerSource);
      if (!record) {
        return false;
      }
      const fallbackResult = await callAppsScriptApi("saveLeaveRecord", { record });
      return fallbackResult?.ok !== false;
    }
  }

  function pushStudentBillingProfileQuietly(studentCode, triggerSource = "") {
    pushCloudBillingProfile(studentCode, triggerSource).catch((error) => {
      console.warn("Cloud billing profile sync failed:", error);
    });
  }

  async function seedMissingCloudBillingProfiles(cloudProfiles, scope = {}) {
    const cloudStudentCodes = new Set(
      (cloudProfiles || [])
        .map((profile) => normalizeParticipantCode(profile?.studentCode))
        .filter(Boolean)
    );
    const localStudents = getScopedBillingStudents(scope)
      .filter((student) => !cloudStudentCodes.has(normalizeParticipantCode(student?.code)));
    let seeded = false;
    for (const student of localStudents) {
      try {
        seeded = await pushCloudBillingProfile(student.code, "seed_missing_billing_profile") || seeded;
      } catch (error) {
        console.warn("Cloud billing profile seed failed:", error);
      }
    }
    return seeded;
  }

  async function syncCloudBillingProfiles(scope = {}) {
    if (!getAppsScriptUrl()) {
      return false;
    }
    const profiles = await listCloudBillingProfiles(scope);
    let changed = false;
    profiles.forEach((profile) => {
      if (applyCloudBillingProfile(profile)) {
        changed = true;
      }
    });
    if (changed) {
      addLog(`[同步] 已同步 ${profiles.length} 筆學生繳費資料。`);
      saveState();
    }
    const hasBillingScope = Boolean(
      normalizeParticipantCode(scope.coachCode || activeCoachCode || "") ||
      normalizeParticipantCode(scope.studentCode || activeStudentCode || "")
    );
    if (hasBillingScope) {
      await seedMissingCloudBillingProfiles(profiles, scope);
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
    const nowTime = Date.now();
    const defaultWeekday = pickDefaultLessonWeekdayForCoach(coachCode);
    const firstCycleDay = getNextDateKeyByWeekday(todayKey, defaultWeekday, true);
    const defaultTime = pickDefaultLessonTimeForCoach(coachCode);
    const dayOffsets = [-7, 0, 7, 14, 21, 28, 35, 42];

    dayOffsets.forEach((offset) => {
      const targetDateKey = addDays(firstCycleDay, offset);
      const targetStartAt = makeTaipeiDateTime(targetDateKey, defaultTime).toISOString();
      const targetTime = new Date(targetStartAt).getTime();
      if (!isAfterTrackingDataReset(targetStartAt) || !Number.isFinite(targetTime) || targetTime <= nowTime) {
        return;
      }
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
    if (isKnownFakeStudentName(studentName)) {
      return null;
    }
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

    const normalizedEmail = normalizeEmailList(studentEmail).join(", ");
    // 不再自動把 defaultNotifyEmail（教練自己的 email）填到新學生身上。
    // 過去會這樣做造成「8 位學生 email 都變成教練的，學生收不到通知」。
    // 沒填就空著，讓教練在計費面板看到再補。
    // defaultNotifyEmail 在 resolveNoticeRecipients 仍會作為「沒收件人時的最終備援」。
    const created = {
      code: normalizedStudentCode,
      name: String(studentName || "").trim() || `${normalizedStudentCode} 學生`,
      coachCode: normalizedCoachCode,
      email: normalizedEmail,
      chargeStartCount: 0,
      chargeStartCountUpdatedAt: "",
      chargeStartCountUpdatedBy: "",
      paidThroughCount: 0,
      paymentStatus: "unknown",
      paymentNote: "",
      paymentConfirmedAt: "",
      paymentConfirmedBy: "",
      chargeReminderLogs: [],
      billingUpdatedAt: "",
      billingUpdatedBy: ""
    };
    state.students.push(created);
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
      if (isKnownFakeStudentName(studentName)) {
        return;
      }
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
    // 移除 bootstrap POST 嘗試 —— 主後端 Code.gs 的 doPost 沒有 case "bootstrap"，
    // 只有 bootstrapAdmin / bootstrapStudent。每次同步都印兩條 "Unsupported action"
    // 警告純屬噪音。GET bootstrap 已涵蓋初始拉取的需求。
    //
    // 另外：syncCoachflowRosterFromPayload 即使成功處理 payload 但本地 state 已是
    // 最新狀態（changed=false），也會回 false 讓 loop 繼續往下。這在「資料早已同
    // 步」的常見情境會多走幾次無謂的 fallback。先嘗試「把 payload 處理成功就算
    // 成功」，再沒抓到資料才往下試。
    const attempts = [
      { action: "bootstrap", method: "GET" },
      { action: "bootstrapAdmin", method: "GET" },
      { action: "bootstrapAdmin", method: "POST" }
    ];
    for (const attempt of attempts) {
      try {
        const payload = await callCoachflowApi(attempt.action, {}, attempt.method);
        // payload 帶到實際的 students/coaches 就視為成功（即使本地沒變更）
        const hasRoster = (
          (Array.isArray(payload?.students) && payload.students.length > 0) ||
          (Array.isArray(payload?.coaches) && payload.coaches.length > 0) ||
          (payload?.data && (
            (Array.isArray(payload.data.students) && payload.data.students.length > 0) ||
            (Array.isArray(payload.data.coaches) && payload.data.coaches.length > 0)
          ))
        );
        const changed = syncCoachflowRosterFromPayload(payload, "雲端 CoachFlow");
        if (changed || hasRoster) {
          return changed;
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

    // 自動為補課 lesson 產生 description（含原請假課程時間、補課單號），讓教練在
    // Google 日曆事件詳情頁能直接看到「這堂是補哪一筆請假」。
    // 重排 lesson 的 description 由呼叫端透過 extra.description 帶入。
    let autoDescription = "";
    if (lesson.sourceType === "MAKEUP") {
      const makeupReq = (state.makeupRequests || []).find((req) => (
        req &&
        req.studentCode === lesson.studentCode &&
        req.coachCode === lesson.coachCode &&
        String(req.startAt) === String(startAt) &&
        (req.status === "approved" || req.lessonId === lesson.id)
      ));
      const studentLabel = `${student?.name || lesson.studentCode}（${lesson.studentCode}）`;
      const coachLabel = `${coach?.name || lesson.coachCode}（${lesson.coachCode}）`;
      let originalLessonTimeText = "(未知)";
      let leaveCode = "-";
      let makeupCode = "-";
      if (makeupReq) {
        const originalLeaveLesson = getLessonById(makeupReq.lessonId);
        if (originalLeaveLesson?.startAt) {
          originalLessonTimeText = formatDateTime(originalLeaveLesson.startAt);
        }
        leaveCode = String(makeupReq.leaveId || "-").trim() || "-";
        makeupCode = String(makeupReq.code || makeupReq.id || "-").trim() || "-";
      }
      autoDescription = [
        "【補課】",
        `學生：${studentLabel}`,
        `教練：${coachLabel}`,
        `原請假時段：${originalLessonTimeText}`,
        `對應請假單號：${leaveCode}`,
        `補課申請編號：${makeupCode}`,
        `補課課程編號：${lesson.id}`
      ].join("\n");
    }

    return {
      lessonId: lesson.id,
      eventId: lesson.calendarEventId || "",
      title: `${lesson.sourceType === "MAKEUP" ? "補課" : "課程"} ${student?.name || lesson.studentCode} / ${coach?.name || lesson.coachCode}`,
      description: autoDescription,
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

  function shouldAutoRemoveUnmatchedLocalCalendarLesson(lesson) {
    if (!lesson || lesson.attendanceStatus === "leave-normal") {
      return false;
    }
    return false;
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
    const hasStudentCodeInTitle = Boolean(studentCode && looseSearchable.includes(studentCode));
    const hasStudentNameInTitle = Boolean(studentName && looseSearchable.includes(studentName));
    const hasStudentNameLooseMatch = Boolean(student?.name && scoreLooseNameMatch(event?.title || "", student.name) > 0);
    const hasStudentText = hasStudentCodeInTitle || hasStudentNameInTitle || hasStudentNameLooseMatch;

    // Substring 防呆：若教練底下另一位學生名「包含」當前學生名（例：「朱朱媽媽」包「朱朱」），
    // 且事件標題中出現該更長名字，這個事件其實屬於那位學生，不應算作當前學生的事件。
    // 否則當前學生的 lesson 還沒綁 eventId 時，resolveMissingEventId 會誤把對方的事件
    // 綁過來，後續刪除就會把對方的 Google 事件刪掉（上線案例：朱朱請假，誤刪朱朱媽媽事件）。
    // 例外：studentCode 完整出現在標題時，視為強訊號，不套用此防呆。
    let ambiguousWithLongerName = false;
    if (hasStudentText && !hasStudentCodeInTitle && studentName && lesson?.coachCode) {
      const otherStudents = (state.students || []).filter((s) => (
        s && s.coachCode === lesson.coachCode && s.code !== lesson.studentCode
      ));
      for (let i = 0; i < otherStudents.length; i += 1) {
        const otherName = normalizeLooseText(otherStudents[i]?.name || "");
        if (
          otherName &&
          otherName.length > studentName.length &&
          otherName.includes(studentName) &&
          looseSearchable.includes(otherName)
        ) {
          ambiguousWithLongerName = true;
          break;
        }
      }
    }
    const reliableStudentText = hasStudentText && !ambiguousWithLongerName;

    let score = 0;
    if (sameEventId) {
      score += 10;
    }
    if (diffMs <= 2 * 60 * 1000) {
      score += 8;
    } else if (diffMs <= 15 * 60 * 1000) {
      score += 4;
    }
    if (reliableStudentText) {
      score += 4;
    }

    return {
      matched: (sameEventId && diffMs <= 15 * 60 * 1000) || (reliableStudentText && diffMs <= 2 * 60 * 1000),
      score
    };
  }

  async function listCalendarEventsForLessonDay(lesson) {
    if (!lesson?.startAt || !getAppsScriptUrl()) {
      return [];
    }
    const startDate = new Date(lesson.startAt);
    if (Number.isNaN(startDate.getTime())) {
      return [];
    }
    const calendarPayload = getCalendarPayloadForCoach(lesson.coachCode);
    if (!calendarPayload.calendarId && !calendarPayload.coachCalendarId) {
      return [];
    }
    const dateKey = getDateKeyInTaipei(startDate);
    const listResult = await callAppsScriptApi("listEvents", {
      ...getDaySyncRange(dateKey),
      ...calendarPayload,
      coachCode: lesson.coachCode,
      dateKey,
      occurrenceDate: dateKey
    });
    return (Array.isArray(listResult?.events) ? listResult.events : [])
      .map(normalizeGoogleCalendarEvent)
      .filter(Boolean);
  }

  function findSingleCalendarEventForLessonInEvents(events, lesson, requestedEventId) {
    const candidates = (Array.isArray(events) ? events : [])
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

  async function resolveSingleCalendarEventForLesson(lesson, requestedEventId) {
    if (!lesson?.startAt || !getAppsScriptUrl()) {
      return null;
    }
    const events = await listCalendarEventsForLessonDay(lesson);
    return findSingleCalendarEventForLessonInEvents(events, lesson, requestedEventId);
  }

  function pickClosestRegularLessonForStudent(studentCode, coachCode, usedLessonIds, targetStartAt, maxDistanceMs = Infinity) {
    const targetTime = new Date(targetStartAt).getTime();
    const candidates = state.lessons
      .filter(
        (lesson) =>
          lesson.studentCode === studentCode &&
          lesson.coachCode === coachCode &&
          isGoogleSyncLesson(lesson) &&
          isLessonAfterTrackingDataReset(lesson) &&
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
    const seenStudentOccurrences = new Set();
    const eventIdToLessons = new Map();
    // Build a separate index of MAKEUP lessons' Google event IDs. These
    // events are tracked by their own MAKEUP lesson record (which is NOT a
    // Google-sync lesson); align logic must NOT let those events fall back
    // to "closest regular lesson for student" — it stole leave-normal
    // lesson startAt and triggered downstream compensation tasks that
    // ended up deleting the makeup event itself.
    // 上線案例：ZO006 5/9 補課 event 把 leave 原 lesson 的 startAt 從
    // 14:00 拉到 11:00，後續 compensation task retry 用 11:00 找到補課
    // event 並刪掉它，coach 才看不到補課。
    const makeupEventIds = new Set();
    state.lessons.forEach((lesson) => {
      if (lesson.coachCode !== coachCode || !isLessonAfterTrackingDataReset(lesson)) {
        return;
      }
      if (lesson.sourceType === "MAKEUP") {
        const mk = normalizeCalendarEventId(lesson.calendarEventId);
        if (mk) makeupEventIds.add(mk);
      }
    });
    state.lessons.forEach((lesson) => {
      if (lesson.coachCode !== coachCode || !isGoogleSyncLesson(lesson) || !isLessonAfterTrackingDataReset(lesson)) {
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
    let skippedUnmatchedStudent = 0;
    const skippedUnmatchedTitles = [];
    let skippedNonLessonEvent = 0;
    let skippedMakeupClaimed = 0;
    let skippedDuplicateEvent = 0;

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
      if (!hasValidDateValue(startAt)) {
        skippedNonLessonEvent += 1;
        return;
      }
      const durationMs = new Date(String(event?.endAt || "")).getTime() - new Date(startAt).getTime();
      if (!Number.isFinite(durationMs) || durationMs <= 0 || durationMs > 4 * 60 * 60 * 1000) {
        skippedNonLessonEvent += 1;
        return;
      }
      // 若這個 event 已經被某個 MAKEUP lesson 認領，不要讓它再去 fallback
      // 拉「最接近的常規 lesson」。否則補課事件會把 leave/regular lesson 的
      // startAt 拉到補課時間，連帶讓後續 compensation task 用錯時間刪掉
      // 補課事件本身。
      const eventIdNorm = normalizeCalendarEventId(eventId);
      if (makeupEventIds.has(eventIdNorm)) {
        skippedMakeupClaimed += 1;
        return;
      }

      // 對不到學生的事件一律略過，不再自動建「Google 日曆來源學生」＋課程。
      // 之前自動建檔會把教練私人行程全部變成 G 開頭代碼的假學生與課表；
      // 需要納管的事件請改事件標題，或在 config 的 calendarStudentAliases 加別名。
      const studentCode = resolveStudentCodeFromCalendarEvent(event, coachCode);
      if (!studentCode) {
        skippedUnmatchedStudent += 1;
        const title = String(event?.title || "").trim();
        if (title && !skippedUnmatchedTitles.includes(title)) {
          skippedUnmatchedTitles.push(title);
        }
        return;
      }

      const occurrenceKey = getCloudLeaveMatchKeyFromValues(studentCode, coachCode, startAt);
      if (occurrenceKey && seenStudentOccurrences.has(occurrenceKey)) {
        skippedDuplicateEvent += 1;
        return;
      }
      if (occurrenceKey) {
        seenStudentOccurrences.add(occurrenceKey);
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

    const collapsedDuplicateLessonIds = collapseDuplicateGoogleLessons({ coachCode });

    return {
      totalEvents: sortedEvents.length,
      matchedEvents,
      updatedStart,
      relinkedEventId,
      createdLessons,
      skippedUnmatchedStudent,
      skippedUnmatchedTitles,
      skippedNonLessonEvent,
      skippedMakeupClaimed,
      skippedDuplicateEvent,
      collapsedDuplicateLessonIds,
      matchedLessonIds: Array.from(usedLessonIds)
    };
  }

  function buildLessonFromDeletePayload(payload = {}) {
    return {
      id: String(payload.lessonId || "").trim(),
      studentCode: normalizeParticipantCode(payload.studentCode || ""),
      coachCode: normalizeParticipantCode(payload.coachCode || activeCoachCode || ""),
      startAt: String(payload.occurrenceStartAt || payload.lessonStartAt || payload.startAt || "").trim(),
      calendarEventId: String(payload.eventId || payload.calendarEventId || "").trim(),
      sourceType: String(payload.sourceType || "REGULAR").trim() || "REGULAR"
    };
  }

  async function isDeleteOccurrenceAlreadyAbsent(payload = {}, lesson = null) {
    const targetLesson = lesson || getLessonById(payload.lessonId) || buildLessonFromDeletePayload(payload);
    if (!targetLesson?.startAt || !targetLesson?.coachCode) {
      return false;
    }
    const requestedEventId = String(payload.eventId || payload.calendarEventId || targetLesson.calendarEventId || "").trim();
    const events = await listCalendarEventsForLessonDay(targetLesson);
    return !findSingleCalendarEventForLessonInEvents(events, targetLesson, requestedEventId);
  }

  function markCalendarDeleteAlreadyHandled(lesson, message) {
    if (lesson) {
      lesson.calendarEventId = "";
      saveState();
    }
    addLog(`[Google日曆] ${message}`);
  }

  async function resolvePendingDeleteCompensationTasksFromGoogle(scope = {}) {
    const syncScope = getCloudStateScope(scope);
    const tasks = (state.compensationTasks || []).filter((task) => (
      task.status !== "completed" &&
      (task.type === "deleteEvent" || task.type === "deleteSingleEvent") &&
      isCompensationTaskInScope(task, syncScope)
    ));
    let changed = false;
    for (const task of tasks) {
      const finalPayload = addSingleEventDeleteGuards(task.payload || {});
      const lesson = getLessonById(finalPayload.lessonId);
      try {
        if (await isDeleteOccurrenceAlreadyAbsent(finalPayload, lesson)) {
          task.payload = finalPayload;
          task.status = "completed";
          task.completedAt = new Date().toISOString();
          task.lastError = "";
          task.resolvedBy = "google-day-check";
          task.resolvedReason = "Google day occurrence already absent";
          if (lesson) {
            lesson.calendarEventId = "";
          }
          addLog(`[補償任務] ${task.type} 視為完成（Google 當天已無此單堂事件，${task.id}）。`);
          changed = true;
        }
      } catch (error) {
        task.lastError = String(error?.message || error || task.lastError || "");
      }
    }
    if (changed) {
      saveState();
    }
    return changed;
  }

  async function tryDeleteCalendarEventForLesson(lesson, reason, strictMode, options = {}) {
    let eventId = String(lesson.calendarEventId || "").trim();
    const shouldResolveMissingEvent = Boolean(options.resolveMissingEventId);
    if (shouldResolveMissingEvent && (!eventId || isGeneratedLocalCalendarEventId(eventId))) {
      try {
        const resolvedEvent = await resolveSingleCalendarEventForLesson(lesson, "");
        if (resolvedEvent?.eventId) {
          eventId = String(resolvedEvent.eventId || "").trim();
          lesson.calendarEventId = eventId;
        } else {
          if (eventId && isGeneratedLocalCalendarEventId(eventId)) {
            lesson.calendarEventId = "";
            saveState();
          }
          return false;
        }
      } catch (error) {
        addLog(`[Google日曆] 請假課程 ${lesson.id} 無法確認待刪除事件：${String(error?.message || "未知錯誤")}`);
        return false;
      }
    }
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
        markCalendarDeleteAlreadyHandled(
          lesson,
          `已確認 ${lesson.id} 在該日期沒有可刪除的單堂事件，視為已同步完成。`
        );
        return true;
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
      try {
        if (await isDeleteOccurrenceAlreadyAbsent(deletePayload, lesson)) {
          markCalendarDeleteAlreadyHandled(
            lesson,
            `刪除 ${lesson.id} 時 Google 回覆失敗，但該日期已無此單堂事件，視為已同步完成。`
          );
          return true;
        }
      } catch (verifyError) {
        addLog(`[Google日曆] 刪除失敗後再次確認單堂事件失敗：${String(verifyError?.message || verifyError)}`);
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
      if (/unsupported|invalid action/i.test(message) || isLeaveBridgeEndpointError(message)) {
        return { exists: true, unsupported: true, message };
      }
      return { exists: true, error: true, message };
    }
  }

  async function syncActiveLeaveCalendarDeletes(scope = {}) {
    if (!getAppsScriptUrl()) {
      return false;
    }
    const coachCode = normalizeParticipantCode(scope.coachCode || activeCoachCode || "");
    const studentCode = normalizeParticipantCode(scope.studentCode || "");
    let syncedAny = false;
    const leaves = state.leaveRequests
      .filter((leave) => (
        String(leave.type || "normal") === "normal" &&
        !leave.revokedAt &&
        (!coachCode || leave.coachCode === coachCode) &&
        (!studentCode || leave.studentCode === studentCode)
      ))
      .sort((a, b) => new Date(a.submittedAt || 0) - new Date(b.submittedAt || 0));

    for (const leave of leaves) {
      const lesson = getLessonById(leave.lessonId);
      if (!lesson || lesson.attendanceStatus !== "leave-normal") {
        continue;
      }
      if (leave.calendarEventId && normalizeCalendarEventId(lesson.calendarEventId) !== normalizeCalendarEventId(leave.calendarEventId)) {
        lesson.calendarEventId = leave.calendarEventId;
      }
      const deleted = await tryDeleteCalendarEventForLesson(lesson, "synced_normal_leave", false, {
        resolveMissingEventId: true
      });
      if (deleted) {
        syncedAny = true;
      }
    }
    return syncedAny;
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
          isLessonAfterTrackingDataReset(lesson) &&
          lesson.calendarOccupied &&
          getDateKeyInTaipei(new Date(lesson.startAt)).startsWith(monthPrefix)
      )
      .sort((a, b) => new Date(a.startAt) - new Date(b.startAt));

    let lessons = collectMonthLessons();
    const syncScopeText = `${String(meta.year).padStart(4, "0")}/${String(meta.month).padStart(2, "0")} 當月`;
    const calendarPayload = getCalendarPayloadForCoach(activeCoachCode);

    const confirmed = window.confirm(
      `本次將以 Google 日曆同步 ${syncScopeText}。\n會先處理已請假的課程，再讀取 Google 當月事件並更新剩下課程；若本地課程在 Google 當月找不到對應事件，會列入同步移除確認。\n是否繼續？`
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
    await syncActiveLeaveCalendarDeletes({ coachCode: activeCoachCode });
    lessons = collectMonthLessons();

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
          const unmatchedTitles = Array.isArray(alignStats.skippedUnmatchedTitles) ? alignStats.skippedUnmatchedTitles : [];
          const unmatchedHint = unmatchedTitles.length
            ? `（略過事件：${unmatchedTitles.slice(0, 8).join("、")}${unmatchedTitles.length > 8 ? "…" : ""}）`
            : "";
          alignedSummary = `Google對齊：calendar ${syncedCalendarId || "-"}，抓到 ${alignStats.totalEvents} 筆，匹配 ${alignStats.matchedEvents} 筆，調整時間 ${alignStats.updatedStart} 筆，重綁事件ID ${alignStats.relinkedEventId} 筆，新增課程 ${alignStats.createdLessons} 筆，未匹配學生 ${alignStats.skippedUnmatchedStudent} 筆${unmatchedHint}，非課程事件 ${alignStats.skippedNonLessonEvent} 筆`;
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
      shiftMonth(studentCalendarMonthStart || new Date(), -1),
      getMonthStart(studentCalendarMonthStart || new Date()),
      shiftMonth(studentCalendarMonthStart || new Date(), 1),
      shiftMonth(studentCalendarMonthStart || new Date(), 2)
    ];
    await syncActiveLeaveCalendarDeletes({ coachCode: activeCoachCode, studentCode: activeStudentCode });
    let totalEvents = 0;
    let totalMatched = 0;
    let totalCreated = 0;
    let totalRemovedPlaceholders = 0;
    let syncedAny = false;
    let completionStats = { changed: false, count: 0, studentCodes: [] };
    let cloudLeaveChanged = false;
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
            isLessonAfterTrackingDataReset(lesson) &&
            lesson.calendarOccupied &&
            getDateKeyInTaipei(new Date(lesson.startAt)).startsWith(monthPrefix) &&
            !matchedLessonIds.has(lesson.id) &&
            shouldAutoRemoveUnmatchedLocalCalendarLesson(lesson)
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
      try {
        cloudLeaveChanged = await syncCloudLeaveRecords({
          coachCode: activeCoachCode,
          studentCode: activeStudentCode
        });
      } catch (error) {
        console.warn("Cloud leave sync failed after student calendar sync:", error);
      }
      completionStats = markAutoCompletedLessonsForBilling({
        coachCode: activeCoachCode,
        studentCode: activeStudentCode,
        sourceLabel: "學生端日曆同步"
      });
      saveState();
      const completedText = completionStats.count ? `，已上課扣堂 ${completionStats.count} 堂` : "";
      addLog(`[日曆同步] 學生端已自動同步 Google 課程：抓到 ${totalEvents} 筆，匹配 ${totalMatched} 筆，新增 ${totalCreated} 筆，移除暫存課 ${totalRemovedPlaceholders} 筆${completedText}。`);
      notifyUser(`已同步 Google 日曆，並更新今日/已過時間的上課與扣堂狀態${completedText}。`, "success");
      completionStats.studentCodes.forEach((code) => {
        maybeSendChargeReminder(code, "auto_completed_lesson").catch((error) => {
          console.error("billing reminder failed:", error);
        });
      });
      repushBillingForAutoCompletedStudents(completionStats, "student_calendar_auto_complete");
      // 完整課表雲端化：日曆同步後（含自動扣堂、移除暫存課）把完整課表推上雲保鮮。
      pushCloudLessonsQuietly({ coachCode: activeCoachCode, studentCode: activeStudentCode });
    }
    return syncedAny || completionStats.changed || cloudLeaveChanged;
  }

  async function syncCoachCalendarEventsFromGoogle(options = {}) {
    const coachCode = normalizeParticipantCode(options.coachCode || activeCoachCode);
    if (!coachCode || !getAppsScriptUrl()) {
      return false;
    }
    const calendarPayload = getCalendarPayloadForCoach(coachCode);
    if (!calendarPayload.calendarId && !calendarPayload.coachCalendarId) {
      return false;
    }

    const baseMonthStart = getMonthStart(options.monthStart || coachCalendarMonthStart || new Date());
    const parsedLookaheadMonths = Number(options.lookaheadMonths ?? 1);
    const lookaheadMonths = Math.max(
      0,
      Math.min(2, Number.isFinite(parsedLookaheadMonths) ? parsedLookaheadMonths : 1)
    );
    const monthStarts = [];
    const seenMonths = new Set();
    for (let offset = -1; offset <= lookaheadMonths; offset += 1) {
      const monthStart = shiftMonth(baseMonthStart, offset);
      const monthKey = getDateKeyInTaipei(monthStart).slice(0, 7);
      if (!seenMonths.has(monthKey)) {
        seenMonths.add(monthKey);
        monthStarts.push(monthStart);
      }
    }

    if (options.showStatus && el.coachCalendarSyncText) {
      el.coachCalendarSyncText.textContent = "正在自動同步 Google 日曆...";
    }

    await syncActiveLeaveCalendarDeletes({ coachCode });

    let totalEvents = 0;
    let totalMatched = 0;
    let totalCreated = 0;
    let totalRemovedPlaceholders = 0;
    let syncedAny = false;
    let cloudLeaveChanged = false;

    for (const monthStart of monthStarts) {
      const monthRange = getMonthSyncRange(monthStart);
      try {
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
            isLessonAfterTrackingDataReset(lesson) &&
            lesson.calendarOccupied &&
            getDateKeyInTaipei(new Date(lesson.startAt)).startsWith(monthPrefix) &&
            !matchedLessonIds.has(lesson.id) &&
            shouldAutoRemoveUnmatchedLocalCalendarLesson(lesson)
          ))
          .forEach((lesson) => {
            markLessonRemovedByCalendar(lesson, "coach_google_auto_sync");
            totalRemovedPlaceholders += 1;
          });
        totalEvents += stats.totalEvents;
        totalMatched += stats.matchedEvents;
        totalCreated += stats.createdLessons;
        syncedAny = true;
      } catch (error) {
        const message = String(error?.message || "未知錯誤");
        addLog(`[日曆同步] 教練端自動同步略過：${message}`);
      }
    }

    if (syncedAny) {
      ensureLessonCalendarEventIds();
      ensureParticipantEmails();
      ensureStudentBillingProfiles();
      try {
        cloudLeaveChanged = await syncCloudLeaveRecords({ coachCode });
      } catch (error) {
        console.warn("Cloud leave sync failed after coach calendar sync:", error);
      }
      const completionStats = markAutoCompletedLessonsForBilling({
        coachCode,
        sourceLabel: "教練端日曆同步"
      });
      saveState();
      const completedText = completionStats.count ? `，已上課扣堂 ${completionStats.count} 堂` : "";
      const summary = `已自動同步 Google 日曆：抓到 ${totalEvents} 筆，匹配 ${totalMatched} 筆，新增 ${totalCreated} 筆，移除暫存課 ${totalRemovedPlaceholders} 筆${completedText}。`;
      addLog(`[日曆同步] 教練端${summary}`);
      notifyUser(summary, "success");
      if (options.showStatus && el.coachCalendarSyncText) {
        el.coachCalendarSyncText.textContent = summary;
      }
      completionStats.studentCodes.forEach((code) => {
        maybeSendChargeReminder(code, "auto_completed_lesson").catch((error) => {
          console.error("billing reminder failed:", error);
        });
      });
      repushBillingForAutoCompletedStudents(completionStats, "coach_calendar_auto_complete");
      // 完整課表雲端化：教練端日曆同步後把整個 coach scope 的完整課表推上雲保鮮。
      pushCloudLessonsQuietly({ coachCode });
    }
    return syncedAny || cloudLeaveChanged;
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
          isLessonAfterTrackingDataReset(lesson) &&
          lesson.calendarOccupied &&
          getDateKeyInTaipei(new Date(lesson.startAt)).startsWith(monthPrefix) &&
          !matchedLessonIds.has(lesson.id) &&
          shouldAutoRemoveUnmatchedLocalCalendarLesson(lesson)
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
    // 完整課表雲端化：removed 沒有 completedAt 可當 updatedAt fallback，明確補上時間戳，
    // 否則雲端這筆會 fallback 到舊的 startAt、被其他裝置 stale 的 scheduled push 蓋回去。
    lesson.updatedAt = lesson.calendarRemovedAt;
    addLog(`[日曆同步] 已依 Google 日曆移除課程 ${lesson.id}（${formatDateTime(lesson.startAt)}）。`);
  }

  async function tryCreateCalendarEventForLesson(lesson, reason, strictMode, extras = {}) {
    if (isNormalLeaveCalendarCreateBlocked({
      reason,
      attendanceStatus: lesson?.attendanceStatus,
      type: lesson?.type
    })) {
      addLog(`[Google日曆] 已阻擋正常請假建立事件：${lesson?.id || "-"}`);
      return "";
    }
    const previousEventId = lesson.calendarEventId || "";
    try {
      const result = await callAppsScriptApi("createEvent", buildLessonEventPayload(lesson, { reason, ...extras }));
      const createdEventId = String(result?.eventId || result?.calendarEventId || lesson.calendarEventId || newId("GCAL")).trim();
      lesson.calendarEventId = createdEventId;
      if (!result?.mock) {
        const serverEventId = String(result?.eventId || result?.calendarEventId || "").trim();
        let verifiedEvent = null;
        let verificationMessage = "";
        try {
          verifiedEvent = await resolveSingleCalendarEventForLesson(lesson, createdEventId);
        } catch (error) {
          verificationMessage = String(error?.message || error || "驗證失敗");
        }
        if (verifiedEvent) {
          lesson.calendarEventId = String(verifiedEvent.eventId || createdEventId).trim();
        } else if (serverEventId) {
          // 後端 createEvent 只有在 calendar.createEvent 成功後才會回 eventId，
          // 所以「有拿到 serverEventId」= 事件確實建好了。建完馬上重讀驗證失敗，
          // 多半是 Google 日曆最終一致性延遲、或標題比對抓不到（學生沒名字用代碼），
          // 不是沒建成。此時務必信任 serverEventId、不要還原 + 排「再建一次」的補償
          // 任務 —— 那正是「系統一直亂生重複『課程…』事件」的根因。
          lesson.calendarEventId = serverEventId;
          addLog(`[Google日曆] 建立後即時重讀未確認，信任後端回傳事件 ${serverEventId}（不重建）：${lesson.id}${verificationMessage ? `（${verificationMessage}）` : ""}`);
        } else {
          // 後端沒回 eventId = 真的沒建成，照舊 strict throw / 排補償。
          lesson.calendarEventId = previousEventId;
          const message = verificationMessage || "後端未回傳事件 ID，視為建立失敗";
          addLog(`[Google日曆] 建立事件失敗：${lesson.id}（${message}）`);
          if (strictMode) {
            throw new Error(`Google 日曆建立失敗：${message}`);
          }
          const retryPayload = buildLessonEventPayload({ ...lesson, calendarEventId: "" }, { reason });
          enqueueCompensationTask("createEvent", retryPayload, `建立失敗：${message}`);
          saveState();
          return "";
        }
      }
      addLog(`[Google日曆] 已建立事件 ${createdEventId}${result.mock ? "（模擬）" : ""}`);
      saveState();
      return lesson.calendarEventId;
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
            `導入前已扣堂數：${payload.baseChargedCount ?? "-"}`,
            `本次系統扣堂：${payload.systemChargedCount ?? "-"}`,
            `已繳到堂數：${payload.paidThroughCount ?? "-"}`,
            `本期已扣：${payload.currentCycleChargedCount ?? "-"} / ${CHARGE_REMINDER_STEP}`,
            `下次應繳門檻：第 ${payload.nextPaymentDueCount ?? "-"} 堂`,
            `距下次繳費：${payload.remainingToNextPayment ?? "-"} 堂`,
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
            `導入前已扣堂數：${payload.baseChargedCount ?? "-"}`,
            `本次系統扣堂：${payload.systemChargedCount ?? "-"}`,
            `已繳到堂數：${payload.paidThroughCount ?? "-"}`,
            `本期已扣：${payload.currentCycleChargedCount ?? "-"} / ${CHARGE_REMINDER_STEP}`,
            `下次應繳門檻：第 ${payload.nextPaymentDueCount ?? "-"} 堂`,
            `距下次繳費：${payload.remainingToNextPayment ?? "-"} 堂`,
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
            `導入前已扣堂數：${payload.baseChargedCount ?? "-"}`,
            `本次系統扣堂：${payload.systemChargedCount ?? "-"}`,
            `已結算到第 ${payload.paidThroughCount ?? "-"} 堂`,
            `本期已扣：${payload.currentCycleChargedCount ?? "-"} / ${CHARGE_REMINDER_STEP}`,
            `距下次繳費：${payload.remainingToNextPayment ?? "-"} 堂`,
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
      case "leave_submitted_by_coach": {
        return {
          subject: `[CoachFlow] 教練已代學生請假：${studentText}`,
          body: [
            "您好，",
            "",
            "教練已在 CoachFlow 代學生送出請假，系統已同步處理原課程。",
            `學生：${studentText}`,
            `教練：${coachText}`,
            `原課時間：${lessonTimeText}`,
            `請假編號：${leaveCode}`,
            `代送者：${payload.submittedBy || payload.submittedByRole || "教練"}`,
            "",
            "如果資訊有誤，請直接聯絡教練確認。",
            "本信件由 CoachFlow 請假系統自動寄出。"
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
      case "lesson_rescheduled": {
        const oldText = safeFormatNoticeDateTime(payload.oldStartAt);
        const newText = safeFormatNoticeDateTime(payload.newStartAt || payload.lessonStartAt || payload.startAt);
        return {
          subject: `[CoachFlow] 課程時間調整通知｜${studentText}`,
          body: [
            `${studentText} 您好：`,
            "",
            "您原本的上課時段已由教練調整：",
            `學生：${studentText}`,
            `教練：${coachText}`,
            `原時段：${oldText}`,
            `新時段：${newText}`,
            `課程編號：${payload.lessonId || "-"}`,
            "",
            "如時間有疑問請直接聯繫教練。",
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
      notifyUser(`${label} 未寄出：缺少收件人 Email。`, "warning");
      return false;
    }
    try {
      const result = await callAppsScriptApi("sendEmail", { template, ...finalPayload });
      addLog(`[Email] ${label} 已送出${result.mock ? "（模擬）" : ""}`);
      notifyUser(`${label} 已寄出。`, "success");
      saveState();
      return true;
    } catch (error) {
      const message = String(error?.message || "未知錯誤");
      enqueueCompensationTask("sendEmail", { template, ...finalPayload }, message);
      addLog(`[Email] ${label} 發送失敗：${message}`);
      notifyUser(`${label} 發送失敗，已排入補償任務。`, "warning");
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
        label: String(item?.label || ""),
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
          isLessonActiveForTrackingStats(lesson) &&
          isLessonChargedForBilling(lesson) &&
          isLessonAfterChargeStartBaseline(lesson, student)
        ))
        .length;
      const existingSystemChargedCount = existingChargedCount;
      const fallbackPaidThrough = existingStatus === "paid"
        ? getPaidQuotaCeiling(existingSystemChargedCount)
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
        chargeStartCountUpdatedAt: String(student.chargeStartCountUpdatedAt || ""),
        chargeStartCountUpdatedBy: normalizeParticipantCode(student.chargeStartCountUpdatedBy || ""),
        chargeReminderLogs: normalizeChargeReminderLogs(student.chargeReminderLogs),
        billingUpdatedAt: String(student.billingUpdatedAt || getStudentBillingUpdatedAt(student) || ""),
        billingUpdatedBy: normalizeParticipantCode(student.billingUpdatedBy || "")
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

  function resetImportedChargedCountsOnce() {
    state.migrations = state.migrations && typeof state.migrations === "object" ? state.migrations : {};
    if (state.migrations[RESET_IMPORTED_CHARGED_COUNT_MIGRATION]) {
      return false;
    }

    let resetCount = 0;
    state.students = (state.students || []).map((student) => {
      if (!student || typeof student !== "object") {
        return student;
      }
      const currentCount = toNonNegativeInt(student.chargeStartCount, 0);
      if (currentCount > 0) {
        resetCount += 1;
      }
      return {
        ...student,
        chargeStartCount: 0,
        chargeStartCountUpdatedAt: "",
        chargeStartCountUpdatedBy: "",
        billingUpdatedAt: new Date().toISOString(),
        billingUpdatedBy: "SYSTEM"
      };
    });
    state.migrations[RESET_IMPORTED_CHARGED_COUNT_MIGRATION] = {
      at: new Date().toISOString(),
      resetCount
    };
    if (resetCount > 0) {
      addLog(`[計費] 已將 ${resetCount} 位學生的導入前已扣堂數歸零，請手動重新輸入。`);
    }
    saveState();
    return resetCount > 0;
  }

  function resetAllTrackingDataOnce() {
    state.migrations = state.migrations && typeof state.migrations === "object" ? state.migrations : {};
    if (state.migrations[RESET_ALL_TRACKING_DATA_MIGRATION]) {
      return false;
    }

    const before = {
      leaveRequests: (state.leaveRequests || []).length,
      makeupRequests: (state.makeupRequests || []).length,
      coachBlocks: (state.coachBlocks || []).length,
      compensationTasks: (state.compensationTasks || []).length,
      eventLog: (state.eventLog || []).length
    };

    state.students = (state.students || []).map((student) => ({
      ...student,
      chargeStartCount: 0,
      chargeStartCountUpdatedAt: "",
      chargeStartCountUpdatedBy: "",
      paidThroughCount: 0,
      paymentStatus: "unknown",
      paymentNote: "",
      paymentConfirmedAt: "",
      paymentConfirmedBy: "",
      chargeReminderLogs: [],
      billingUpdatedAt: new Date().toISOString(),
      billingUpdatedBy: "SYSTEM"
    }));

    state.lessons = (state.lessons || [])
      .filter((lesson) => lesson.sourceType !== "MAKEUP" && isLessonAfterTrackingDataReset(lesson))
      .map((lesson) => {
        const nextLesson = {
          ...lesson,
          attendanceStatus: "scheduled",
          charged: false
        };
        if (nextLesson.sourceType === "REGULAR" || nextLesson.sourceType === "GOOGLE_CALENDAR") {
          nextLesson.calendarOccupied = true;
        }
        delete nextLesson.calendarRemovedAt;
        delete nextLesson.calendarRemovedReason;
        delete nextLesson.beforeCalendarRemoved;
        delete nextLesson.completedAt;
        delete nextLesson.completedBy;
        return nextLesson;
      });

    state.leaveRequests = [];
    state.makeupRequests = [];
    state.coachBlocks = [];
    state.compensationTasks = [];
    state.eventLog = [{
      id: newId("LOG"),
      at: new Date().toISOString(),
      message: "已依需求將既有堂數、請假、補課、繳費與事件紀錄歸零，從此刻重新記錄。"
    }];
    state.migrations[RESET_ALL_TRACKING_DATA_MIGRATION] = {
      at: new Date().toISOString(),
      resetAt: TRACKING_DATA_RESET_AT,
      before
    };
    saveState();
    return true;
  }

  function purgePreResetTrackingDataOnce() {
    state.migrations = state.migrations && typeof state.migrations === "object" ? state.migrations : {};
    if (state.migrations[PURGE_PRE_RESET_TRACKING_DATA_MIGRATION]) {
      return false;
    }

    const before = {
      lessons: (state.lessons || []).length,
      leaveRequests: (state.leaveRequests || []).length,
      makeupRequests: (state.makeupRequests || []).length,
      coachBlocks: (state.coachBlocks || []).length,
      compensationTasks: (state.compensationTasks || []).length,
      eventLog: (state.eventLog || []).length
    };

    state.students = (state.students || []).map((student) => ({
      ...student,
      chargeStartCount: 0,
      chargeStartCountUpdatedAt: "",
      chargeStartCountUpdatedBy: "",
      paidThroughCount: 0,
      paymentStatus: "unknown",
      paymentNote: "",
      paymentConfirmedAt: "",
      paymentConfirmedBy: "",
      chargeReminderLogs: [],
      billingUpdatedAt: new Date().toISOString(),
      billingUpdatedBy: "SYSTEM"
    }));

    state.lessons = (state.lessons || [])
      .filter((lesson) => lesson.sourceType !== "MAKEUP" && isLessonAfterTrackingDataReset(lesson))
      .map((lesson) => {
        const nextLesson = {
          ...lesson,
          attendanceStatus: "scheduled",
          charged: false
        };
        if (nextLesson.sourceType === "REGULAR" || nextLesson.sourceType === "GOOGLE_CALENDAR") {
          nextLesson.calendarOccupied = true;
        }
        delete nextLesson.calendarRemovedAt;
        delete nextLesson.calendarRemovedReason;
        delete nextLesson.beforeCalendarRemoved;
        delete nextLesson.completedAt;
        delete nextLesson.completedBy;
        return nextLesson;
      });

    state.leaveRequests = [];
    state.makeupRequests = [];
    state.coachBlocks = [];
    state.compensationTasks = [];
    state.eventLog = [{
      id: newId("LOG"),
      at: new Date().toISOString(),
      message: "已重新清除歸零前課程與所有堂數、請假、補課、繳費紀錄，從歸零時間後重新計算。"
    }];
    state.migrations[PURGE_PRE_RESET_TRACKING_DATA_MIGRATION] = {
      at: new Date().toISOString(),
      resetAt: TRACKING_DATA_RESET_AT,
      before
    };
    saveState();
    return true;
  }

  function restoreAutoRemovedLocalLessonsOnce() {
    state.migrations = state.migrations && typeof state.migrations === "object" ? state.migrations : {};
    if (state.migrations[RESTORE_AUTO_REMOVED_LOCAL_LESSONS_MIGRATION]) {
      return false;
    }
    const autoSyncReasons = new Set([
      "student_google_auto_sync",
      "coach_google_auto_sync",
      "readonly_google_auto_sync"
    ]);
    let restoredCount = 0;
    state.lessons = (state.lessons || []).map((lesson) => {
      if (
        !lesson ||
        lesson.attendanceStatus !== "calendar-removed" ||
        !autoSyncReasons.has(String(lesson.calendarRemovedReason || "")) ||
        !isGeneratedLocalCalendarEventId(lesson.beforeCalendarRemoved?.calendarEventId || "")
      ) {
        return lesson;
      }
      const previous = lesson.beforeCalendarRemoved || {};
      restoredCount += 1;
      const nextLesson = {
        ...lesson,
        attendanceStatus: previous.attendanceStatus || "scheduled",
        calendarOccupied: previous.calendarOccupied !== false,
        calendarEventId: previous.calendarEventId || lesson.calendarEventId || "",
        charged: false
      };
      delete nextLesson.calendarRemovedAt;
      delete nextLesson.calendarRemovedReason;
      delete nextLesson.beforeCalendarRemoved;
      return nextLesson;
    });
    state.migrations[RESTORE_AUTO_REMOVED_LOCAL_LESSONS_MIGRATION] = {
      at: new Date().toISOString(),
      restoredCount
    };
    if (restoredCount > 0) {
      addLog(`[修復] 已還原 ${restoredCount} 堂被日曆同步誤標刪除的本機課程。`);
    }
    saveState();
    return restoredCount > 0;
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
      .filter(
        (lesson) =>
          lesson.studentCode === studentCode &&
          lesson.attendanceStatus !== "calendar-removed" &&
          isLessonActiveForTrackingStats(lesson)
      )
      .sort((a, b) => new Date(b.startAt) - new Date(a.startAt));
    const startCount = toNonNegativeInt(student?.chargeStartCount, 0);
    const rawChargedLessons = lessons
      .filter((lesson) => isLessonChargedForBilling(lesson) && isLessonAfterChargeStartBaseline(lesson, student))
      .sort((a, b) => new Date(a.startAt) - new Date(b.startAt));
    const chargedLessons = rawChargedLessons
      .sort((a, b) => new Date(b.startAt) - new Date(a.startAt));
    return {
      student,
      lessons,
      rawChargedLessons,
      chargedLessons,
      startCount,
      totalChargedCount: startCount + chargedLessons.length,
      noShowCount: lessons.filter((lesson) => lesson.attendanceStatus === "no-show").length,
      tempLeaveCount: lessons.filter((lesson) => lesson.attendanceStatus === "temporary-leave").length,
      majorCount: lessons.filter((lesson) => lesson.attendanceStatus === "major-case").length,
      normalLeaveCount: lessons.filter((lesson) => lesson.attendanceStatus === "leave-normal").length
    };
  }

  function getPaidQuotaCeiling(systemChargedCount) {
    const chargedCount = toNonNegativeInt(systemChargedCount, 0);
    if (chargedCount <= 0) {
      return CHARGE_REMINDER_STEP;
    }
    return Math.ceil(chargedCount / CHARGE_REMINDER_STEP) * CHARGE_REMINDER_STEP;
  }

  function normalizePaidQuotaCount(student, startCount) {
    const rawPaidThroughCount = toNonNegativeInt(student?.paidThroughCount, 0);
    const importedCount = toNonNegativeInt(startCount, 0);
    const legacyAdjustedCount = rawPaidThroughCount - importedCount;
    if (
      rawPaidThroughCount > 0 &&
      importedCount > 0 &&
      rawPaidThroughCount % CHARGE_REMINDER_STEP !== 0 &&
      legacyAdjustedCount > 0
    ) {
      return getPaidQuotaCeiling(legacyAdjustedCount);
    }
    return rawPaidThroughCount;
  }

  function getNextPaidQuotaCount(systemChargedCount) {
    const chargedCount = toNonNegativeInt(systemChargedCount, 0);
    return (Math.floor(chargedCount / CHARGE_REMINDER_STEP) + 1) * CHARGE_REMINDER_STEP;
  }

  function getStudentBillingCycle(stats) {
    const student = stats?.student;
    const systemChargedCount = toNonNegativeInt(stats?.chargedLessons?.length, 0);
    const importedCount = toNonNegativeInt(stats?.startCount, 0);
    const totalChargedCount = toNonNegativeInt(stats?.totalChargedCount, importedCount + systemChargedCount);
    const storedStatus = normalizePaymentStatus(student?.paymentStatus);
    const storedPaidQuotaCount = normalizePaidQuotaCount(student, stats?.startCount);
    const defaultQuotaCount = importedCount > 0 ? getPaidQuotaCeiling(importedCount) : CHARGE_REMINDER_STEP;
    const activeQuotaCount = storedPaidQuotaCount || defaultQuotaCount;
    const paidThroughCount = storedPaidQuotaCount || (storedStatus === "paid" || totalChargedCount > 0 ? activeQuotaCount : 0);
    const cycleRemainder = totalChargedCount % CHARGE_REMINDER_STEP;
    const currentCycleChargedCount = totalChargedCount <= 0
      ? 0
      : (cycleRemainder === 0 ? CHARGE_REMINDER_STEP : cycleRemainder);
    const remainingToNextPayment = currentCycleChargedCount === CHARGE_REMINDER_STEP
      ? 0
      : CHARGE_REMINDER_STEP - currentCycleChargedCount;
    const lastCompletedPaymentDueCount = Math.floor(totalChargedCount / CHARGE_REMINDER_STEP) * CHARGE_REMINDER_STEP;
    const coveredPaymentDueCount = storedPaidQuotaCount || (storedStatus === "paid" ? activeQuotaCount : 0);
    // 規則：cycle 結束點（currentCycle 達 step）翻未繳費。教練按確認匯款可清掉
    // 這個提醒——下個 cycle 結束時會再翻一次（per-cycle 提醒）。
    //
    // Cycle 是否「已被確認」用 paymentConfirmedAt 跟「最近一堂被計入扣堂的課」
    // 的 startAt 比：
    //   - confirmedAt 晚於最近扣堂課 startAt → 這 cycle 已 ack → 不翻未繳
    //   - confirmedAt 早於 → 這 cycle 還沒 ack → 翻未繳費
    // 學生上下一堂課後，「最近扣堂課」會更新到那堂、再次晚於 confirmedAt，
    // 下個 cycle 結束時自動翻未繳。
    //
    // 不再依賴 paidThroughCount 判斷：舊版按確認匯款會把 paidThroughCount 自動
    // 推到下個 cycle ceiling，副作用是下個 cycle 結束時系統認為「還在已繳區
    // 間」、不翻未繳，教練漏收一輪（朱朱、朱朱媽媽案例）。
    const chargedLessonsList = Array.isArray(stats?.chargedLessons) ? stats.chargedLessons : [];
    // stats.chargedLessons 已按 startAt 降序排序；[0] 就是最近被計入扣堂的課
    const lastChargedLesson = chargedLessonsList[0];
    const lastChargedLessonTime = lastChargedLesson ? new Date(lastChargedLesson.startAt).getTime() : 0;
    const confirmedTime = new Date(student?.paymentConfirmedAt || 0).getTime();
    const cycleAckedByCoach = Number.isFinite(confirmedTime)
      && confirmedTime > 0
      && (!Number.isFinite(lastChargedLessonTime) || lastChargedLessonTime <= 0 || confirmedTime > lastChargedLessonTime);
    const isPaymentDue = totalChargedCount > 0
      && currentCycleChargedCount === CHARGE_REMINDER_STEP
      && !cycleAckedByCoach;
    const overduePaymentDueCount = isPaymentDue ? lastCompletedPaymentDueCount : 0;
    const nextPaymentDueCount = overduePaymentDueCount || lastCompletedPaymentDueCount + CHARGE_REMINDER_STEP;
    const effectivePaymentStatus = isPaymentDue ? "unpaid" : storedStatus;
    return {
      paidThroughCount,
      systemChargedCount,
      totalChargedCount,
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

  function markStudentPaymentDue(student, billingCycle, triggerSource) {
    if (!student || !billingCycle?.isPaymentDue) {
      return false;
    }
    const milestone = billingCycle.nextPaymentDueCount;
    const currentStatus = normalizePaymentStatus(student.paymentStatus);
    const nextNote = `第 ${milestone} 堂已達繳費門檻，待確認繳費。`;
    let changed = false;
    if (currentStatus !== "unpaid") {
      student.paymentStatus = "unpaid";
      changed = true;
    }
    if (!String(student.paymentNote || "").includes(nextNote)) {
      student.paymentNote = nextNote;
      changed = true;
    }
    if (changed) {
      addLog(`[計費] ${student.code} 已達第 ${milestone} 堂，自動標記為未繳費。`);
    }
    return changed;
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
    const paymentStatusChanged = markStudentPaymentDue(student, billingCycle, triggerSource || "system");
    if (paymentStatusChanged) {
      touchStudentBillingProfile(student, triggerSource || "system");
    }
    const reminderLogs = Array.isArray(student.chargeReminderLogs) ? student.chargeReminderLogs : [];
    if (reminderLogs.some((item) => Number(item?.milestone) === milestone)) {
      if (paymentStatusChanged) {
        saveState();
        renderBillingPanels();
        pushStudentBillingProfileQuietly(student.code, triggerSource || "system");
      }
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
      touchStudentBillingProfile(student, triggerSource || "system");
      addLog(`[計費提醒] ${student.code} 第 ${milestone} 堂提醒${sent ? "已送出" : "未送出"}。`);
      saveState();
      renderBillingPanels();
      pushStudentBillingProfileQuietly(student.code, triggerSource || "system");
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
      touchStudentBillingProfile(student, "manual_charge_panel");
      addLog(`[計費] ${student.code} 計費摘要${sent ? "已寄送" : "未寄送"}。`);
      saveState();
      renderBillingPanels();
      pushStudentBillingProfileQuietly(student.code, "manual_charge_panel");
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
    touchStudentBillingProfile(student, activeCoachCode || "SYSTEM");
    addLog(
      `[通知] ${student.code} 學生通知 Email 已更新：${normalizedEmail || "清空（將改用預設通知信箱）"}。`
    );
    saveState();
    renderBillingPanels();
    pushStudentBillingProfileQuietly(student.code, "student_email");
    notifyUser(`學生 ${student.name || student.code} 的通知 Email 已儲存。`, "success");
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
    const updatedBy = activeCoachCode || "SYSTEM";
    touchStudentBillingProfile(student, updatedBy);
    // baseline 固定用「導入追蹤起點」（BILLING_TRACKING_START_AT = 5/3），
    // 而非「儲存當下時間」。
    //
    // 舊版每次存檔都把 baseline 重設成 now，造成：教練若晚於學生第一堂課才
    // 設定導入前堂數，夾在「學生首課」與「設定日」之間的課（已上、charged）
    // 既不算系統內扣堂（在 baseline 之前）、又不在手填的導入前數裡 → 整批
    // 憑空消失，教練怎麼調導入前堂數都對不上實際畫面。
    // 已上線案例：ST013 5/9、5/16 已上課，但 baseline 被鎖在 5/30，兩堂消失。
    //
    // 改用固定 tracking start 後，語意清楚：5/3（含）之後 charged 的課一律算
    // 系統內扣堂，5/3 之前的用手填 chargeStartCount 表示。
    student.chargeStartCountUpdatedAt = BILLING_TRACKING_START_AT;
    student.chargeStartCountUpdatedBy = normalizeParticipantCode(updatedBy);
    addLog(`[計費] ${student.code} 導入前已扣堂數調整為 ${nextCount}。`);
    saveState();
    renderBillingPanels();
    notifyUser(`已儲存 ${student.name || student.code} 的導入前已扣堂數：${nextCount}。`, "success");
    pushStudentBillingProfileQuietly(student.code, "set_base_count");
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
        paidThroughCount: billingCycle.paidThroughCount,
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
        label: `匯款確認（已繳到第 ${billingCycle.paidThroughCount} 堂）`,
        status: sent ? "success" : "failed",
        sentAt: new Date().toISOString(),
        to: recipientsText,
        note: sent
          ? "匯款確認信已寄送給教練與學生"
          : (recipientsText ? "匯款確認信寄送失敗，已建立補償任務" : "缺少收件人，匯款確認信未寄送"),
        triggerSource: "payment_confirmed_by_coach"
      };
      student.chargeReminderLogs = [logEntry, ...reminderLogs].slice(0, MAX_CHARGE_REMINDER_LOGS);
      touchStudentBillingProfile(student, "payment_confirmed_by_coach");
      addLog(`[計費] ${student.code} 匯款確認信${sent ? "已寄送" : "未寄送"}。`);
      saveState();
      pushStudentBillingProfileQuietly(student.code, "payment_confirmed_by_coach");
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
    const billingCycle = getStudentBillingCycle(stats);
    const nextPaidThrough = nextStatus === "paid"
      ? getNextPaidQuotaCount(billingCycle.totalChargedCount)
      : billingCycle.paidThroughCount;
    student.paymentStatus = nextStatus;
    student.paymentNote = note;
    student.paymentConfirmedAt = new Date().toISOString();
    student.paymentConfirmedBy = activeCoachCode || "SYSTEM";
    if (nextStatus === "paid") {
      student.paidThroughCount = nextPaidThrough;
    }
    touchStudentBillingProfile(student, activeCoachCode || "SYSTEM");
    addLog(
      `[計費] ${student.code} 繳費狀態更新為 ${getPaymentStatusLabel(nextStatus)}（${student.paymentConfirmedBy}，已繳到第 ${nextStatus === "paid" ? nextPaidThrough : billingCycle.paidThroughCount} 堂）。`
    );
    saveState();
    renderBillingPanels();
    notifyUser(`已更新 ${student.name || student.code} 的繳費狀態：${getPaymentStatusLabel(nextStatus)}。`, "success");
    pushStudentBillingProfileQuietly(student.code, "payment_status");
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
        renderBillingPanels();
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
    // pendingCloudSync 的請假還沒被雲端確認，視為「尚未成立」，所有
    // UI / 重複檢查 / 取消請假等流程都不應該看到它。retry 機制（讀 state
    // 直接掃 pendingCloudSync）會把它推上雲端再轉成正式請假。
    return state.leaveRequests.find((leave) => (
      leave.lessonId === lessonId &&
      !leave.revokedAt &&
      !leave.pendingCloudSync
    ));
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

  function buildMakeupPendingEmailPayload(request) {
    return {
      requestId: request.id,
      requestCode: request.code,
      leaveId: request.leaveId,
      lessonId: request.lessonId,
      studentCode: request.studentCode,
      coachCode: request.coachCode,
      studentEmail: getStudentNoticeEmail(request.studentCode),
      coachEmail: getCoachNoticeEmail(request.coachCode),
      pendingAt: request.pendingAt,
      startAt: request.startAt
    };
  }

  function isSameMakeupRequest(left, right) {
    const leftId = String(left?.id || "").trim();
    const rightId = String(right?.id || "").trim();
    if (leftId && rightId && leftId === rightId) {
      return true;
    }
    const leftCode = String(left?.code || "").trim();
    const rightCode = String(right?.code || "").trim();
    if (leftCode && rightCode && leftCode === rightCode) {
      return true;
    }
    return Boolean(
      String(left?.leaveId || "").trim() &&
      String(left?.leaveId || "").trim() === String(right?.leaveId || "").trim() &&
      normalizeParticipantCode(left?.studentCode) === normalizeParticipantCode(right?.studentCode) &&
      normalizeParticipantCode(left?.coachCode) === normalizeParticipantCode(right?.coachCode) &&
      String(left?.startAt || "").trim() === String(right?.startAt || "").trim()
    );
  }

  function findMakeupRequestInCollection(requests, targetRequest) {
    return (Array.isArray(requests) ? requests : []).find((request) => isSameMakeupRequest(request, targetRequest)) || null;
  }

  async function verifyMakeupRequestSavedToCloud(request) {
    const snapshot = await fetchLeaveStateSnapshotFromCloud({
      coachCode: request.coachCode,
      studentCode: ""
    });
    return Boolean(findMakeupRequestInCollection(snapshot?.state?.makeupRequests, request));
  }

  async function saveMakeupRequestToCloud(request, reason) {
    if (!request?.studentCode || !request?.coachCode) {
      return false;
    }
    for (let attempt = 0; attempt < 3; attempt += 1) {
      const saved = await saveStudentScopedStateToCloud(request.studentCode, request.coachCode, reason);
      if (saved) {
        const verified = await verifyMakeupRequestSavedToCloud(request).catch((error) => {
          console.warn("Makeup cloud verification failed:", error);
          return false;
        });
        if (verified) {
          request.cloudSyncStatus = "synced";
          request.cloudVerifiedAt = new Date().toISOString();
          request.cloudSyncErrorAt = "";
          saveState();
          return true;
        }
        addLog(`[雲端同步] ${request.code || request.id} 已送出保存但雲端尚未驗證，準備重試。`);
      }
      if (attempt < 2) {
        await pullLeaveStateSnapshotFromCloud({
          coachCode: request.coachCode,
          studentCode: request.studentCode
        }).catch((error) => {
          console.warn("Makeup cloud retry refresh failed:", error);
        });
      }
    }
    request.cloudSyncStatus = "pending";
    request.cloudSyncErrorAt = new Date().toISOString();
    saveState();
    return false;
  }

  function shouldSendMakeupPendingNotice(request) {
    return Boolean(
      request &&
      request.status === "pending" &&
      !request.emailNoticeSentAt &&
      (!request.emailNoticeStatus || request.emailNoticeStatus === "waiting-cloud")
    );
  }

  function isMakeupRequestWaitingForCloud(request) {
    return request?.cloudSyncStatus === "pending" || request?.emailNoticeStatus === "waiting-cloud";
  }

  function getMakeupRequestStatusPill(request) {
    if (isMakeupRequestWaitingForCloud(request)) {
      return "<span class=\"status warn\">同步雲端中</span>";
    }
    return getStatusPill(request.status);
  }

  async function sendMakeupPendingNotice(request) {
    if (!shouldSendMakeupPendingNotice(request)) {
      return false;
    }
    const emailPayload = buildMakeupPendingEmailPayload(request);
    request.emailNoticeTo = resolveNoticeRecipients(emailPayload).join(", ");
    const sent = await trySendEmailNotice(
      "makeup_pending",
      emailPayload,
      `補課待審通知 ${request.code || request.id}`
    );
    const emailAt = new Date().toISOString();
    request.emailNoticeStatus = sent ? "sent" : "queued-or-skipped";
    if (sent) {
      request.emailNoticeSentAt = emailAt;
      request.emailNoticeQueuedAt = "";
    } else {
      request.emailNoticeQueuedAt = emailAt;
    }
    saveState();
    await saveMakeupRequestToCloud(request, "makeup_request_email_updated");
    return true;
  }

  async function sendQueuedMakeupPendingNoticesForStudent(studentCode) {
    const normalizedStudentCode = normalizeParticipantCode(studentCode || "");
    if (!normalizedStudentCode) {
      return false;
    }
    let changed = false;
    const requests = state.makeupRequests.filter((request) => (
      normalizeParticipantCode(request.studentCode) === normalizedStudentCode &&
      shouldSendMakeupPendingNotice(request)
    ));
    for (const request of requests) {
      const cloudSaved = await saveMakeupRequestToCloud(request, "makeup_request_retry_before_email");
      if (!cloudSaved) {
        continue;
      }
      changed = await sendMakeupPendingNotice(request) || changed;
    }
    return changed;
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
    // 防雙擊：手指連點 / 按鈕送出後沒立刻 disable 都會造成同一筆 lesson 被請假兩次
    // （已上線案例：ST011 5/31 兩筆相隔 1.7 秒的 leave，明顯是 double-tap）
    const now = Date.now();
    const lastInflight = leaveActionInflight.get(lessonId) || 0;
    if (now - lastInflight < LEAVE_ACTION_DEDUPE_MS) {
      return; // 直接吞掉，不彈 alert（讓使用者覺得自己只點了一次）
    }
    leaveActionInflight.set(lessonId, now);
    try {
      await applyNormalLeaveImpl(lessonId);
    } finally {
      // 留在 map 直到 dedupe 視窗過完，避免快速重試
      setTimeout(() => leaveActionInflight.delete(lessonId), LEAVE_ACTION_DEDUPE_MS);
    }
  }

  async function applyNormalLeaveImpl(lessonId) {
    const lesson = getLessonById(lessonId);
    if (!lesson || lesson.studentCode !== activeStudentCode) {
      alert("找不到課程，或你沒有權限操作。");
      return;
    }
    // Telemetry：記錄學生點下去當下的 lesson 內容，下次再出現「請 5/10 結果系統處理 5/31」
    // 這種怪事，可以從 eventLog 對照學生看到的日期 vs 實際送出的 lesson 是否一致。
    try {
      const sameStudentLessons = state.lessons
        .filter((l) => l.studentCode === activeStudentCode && isGoogleSyncLesson(l) && isLessonAfterTrackingDataReset(l))
        .sort((a, b) => new Date(a.startAt) - new Date(b.startAt))
        .map((l) => `${l.id}@${formatDateTime(l.startAt)}[${l.attendanceStatus}]`);
      addLog(`[請假 telemetry] 學生 ${activeStudentCode} 點請假：lessonId=${lessonId} startAt=${lesson.startAt} eventId=${lesson.calendarEventId || "(空)"} 同學生其他課程=${sameStudentLessons.join(",")}`);
      saveState();
    } catch (telemetryError) {
      console.warn("Leave telemetry failed:", telemetryError);
    }
    if (!isLeaveOpen(lesson)) {
      alert("已超過請假截止時間（前一天 23:59），不可請假。");
      return;
    }
    if (getLeaveByLesson(lessonId)) {
      alert("這堂課已經送出請假。");
      return;
    }

    // 找有沒有先前因雲端 push 失敗暫存的 pending leave，有就重用做 retry。
    // 沒有就建一筆新的 pendingCloudSync leave。重用同一個 id 可避免在 Sheet 留下
    // 不同 id 的多筆重複紀錄。
    let leaveRecord = state.leaveRequests.find((leave) => (
      leave &&
      leave.lessonId === lessonId &&
      !leave.revokedAt &&
      leave.pendingCloudSync &&
      String(leave.type || "normal") === "normal"
    ));
    const isRetryAttempt = Boolean(leaveRecord);
    if (!leaveRecord) {
      leaveRecord = {
        id: newId("LEAVE"),
        lessonId,
        studentCode: lesson.studentCode,
        coachCode: lesson.coachCode,
        calendarEventId: lesson.calendarEventId || "",
        type: "normal",
        submittedAt: new Date().toISOString(),
        submittedBy: activeStudentCode,
        submittedByRole: "student",
        makeupEligible: true,
        pendingCloudSync: true
      };
      state.leaveRequests.push(leaveRecord);
    }
    addLog(
      isRetryAttempt
        ? `學生 ${lesson.studentCode} 重試 pending 請假（${leaveRecord.id}，課程 ${lesson.id}）。`
        : `學生 ${lesson.studentCode} 送出請假（課程 ${lesson.id}，等待雲端確認）。`
    );
    saveState();
    renderAll();
    notifyUser("正在送出請假給雲端，請稍候...", "info");

    // === 雲端 push 必須先成功，才會去動 Google 日曆與寄信 ===
    // 否則失敗時會出現「GCal 已刪 + 學生信箱有確認信 + 教練端永遠看不到請假」這
    // 種 5/16 上線事故的悲劇組合。Cloud 是 source of truth；cloud 沒收到就什麼
    // 都不算數，學生看到 warning 紅字後可以重試或聯絡教練。
    let cloudOk = false;
    let cloudErrorMessage = "";
    try {
      cloudOk = await pushCloudLeaveRecord(leaveRecord, lesson);
    } catch (error) {
      cloudErrorMessage = String(error?.message || error || "");
    }

    if (!cloudOk) {
      addLog(`[雲端請假] 上傳失敗，本機已暫存 pending 待重試：${leaveRecord.id}${cloudErrorMessage ? ` / ${cloudErrorMessage}` : "（Apps Script 回應 ok=false）"}`);
      saveState();
      renderAll();
      notifyUser(
        `請假尚未成立：雲端同步失敗${cloudErrorMessage ? `（${cloudErrorMessage}）` : ""}。系統會在你下次打開請假系統或重新整理時自動重試；如急著請假，請聯絡教練協助代為請假。`,
        "warning"
      );
      return;
    }

    // === 雲端確認收到，現在才正式套用到課程狀態 / 動 Google 日曆 / 寄信 ===
    leaveRecord.cloudSyncedAt = new Date().toISOString();
    delete leaveRecord.pendingCloudSync;
    lesson.calendarOccupied = false;
    lesson.attendanceStatus = "leave-normal";
    lesson.updatedAt = new Date().toISOString();
    addLog(`[雲端請假] 已上傳請假紀錄 ${leaveRecord.id}。`);
    saveState();
    renderAll();
    // 完整課表雲端化：請假後課程轉 leave-normal（不扣堂），同步推上雲讓主系統重算正確。
    pushCloudLessonsQuietly({ coachCode: lesson.coachCode, studentCode: lesson.studentCode });
    notifyUser("請假紀錄已建立，正在同步 Google 日曆與寄出確認信...", "info");
    const calendarSynced = await tryDeleteCalendarEventForLesson(lesson, "student_normal_leave", false, {
      resolveMissingEventId: true
    });
    if (!calendarSynced) {
      addLog(`[Google日曆] 請假 ${leaveRecord.id} 已建立，但 Google 日曆事件尚未確認刪除，將由教練端同步時再處理。`);
      saveState();
    }
    const studentEmail = getStudentNoticeEmail(lesson.studentCode);
    const coachEmail = getCoachNoticeEmail(lesson.coachCode);
    if (!studentEmail) {
      addLog(`[Email] ${lesson.studentCode} 未設定學生 Email，請於計費面板補上學生通知 Email。`);
      saveState();
    }
    const emailPayload = {
      studentCode: lesson.studentCode,
      coachCode: lesson.coachCode,
      studentEmail,
      coachEmail,
      lessonId: lesson.id,
      lessonStartAt: lesson.startAt,
      leaveId: leaveRecord.id
    };
    leaveRecord.emailNoticeTo = resolveNoticeRecipients(emailPayload).join(", ");
    const emailSent = await trySendEmailNotice(
      "leave_submitted",
      emailPayload,
      `學生 ${lesson.studentCode} 請假通知`
    );
    const emailAt = new Date().toISOString();
    leaveRecord.emailNoticeStatus = emailSent ? "sent" : "queued-or-skipped";
    if (emailSent) {
      leaveRecord.emailNoticeSentAt = emailAt;
    } else {
      leaveRecord.emailNoticeQueuedAt = emailAt;
    }
    saveState();
    try {
      await pushCloudLeaveRecord(leaveRecord, lesson);
    } catch (error) {
      addLog(`[雲端請假] 上傳 Email 狀態失敗：${String(error?.message || "未知錯誤")}`);
      saveState();
    }
    renderAll();
    notifyUser(
      emailSent
        ? `請假完成：${getTimeText(lesson.startAt)} 課程已標記請假，確認信已寄出。`
        : `請假完成：${getTimeText(lesson.startAt)} 課程已標記請假，但 Email 尚未成功寄出，請檢查收件人或補償任務。`,
      emailSent ? "success" : "warning"
    );
  }

  async function applyNormalLeaveByCoach(lessonId) {
    // 防雙擊（同上）
    const now = Date.now();
    const lastInflight = leaveActionInflight.get(lessonId) || 0;
    if (now - lastInflight < LEAVE_ACTION_DEDUPE_MS) {
      return;
    }
    leaveActionInflight.set(lessonId, now);
    try {
      await applyNormalLeaveByCoachImpl(lessonId);
    } finally {
      setTimeout(() => leaveActionInflight.delete(lessonId), LEAVE_ACTION_DEDUPE_MS);
    }
  }

  async function applyNormalLeaveByCoachImpl(lessonId) {
    if (!requireCoachWriteAccess()) {
      return;
    }
    const lesson = getLessonById(lessonId);
    if (!lesson || lesson.coachCode !== activeCoachCode) {
      alert("找不到課程。");
      return;
    }
    // Telemetry：教練代請假當下的 lesson 內容，便於追「教練幫忙請假後出現幽靈紀錄」之類的問題。
    try {
      addLog(`[請假 telemetry] 教練 ${activeCoachCode} 代學生 ${lesson.studentCode} 請假：lessonId=${lessonId} startAt=${lesson.startAt} eventId=${lesson.calendarEventId || "(空)"} status=${lesson.attendanceStatus} occupied=${lesson.calendarOccupied}`);
      saveState();
    } catch (telemetryError) {
      console.warn("Leave telemetry failed:", telemetryError);
    }
    if (!isGoogleSyncLesson(lesson) || !lesson.calendarOccupied || lesson.attendanceStatus !== "scheduled") {
      alert("僅可代學生請假仍排在 Google 日曆上的原課。");
      return;
    }
    if (getLeaveByLesson(lessonId)) {
      alert("這堂課已經有請假紀錄。");
      return;
    }
    const studentName = getStudentDisplayName(lesson.studentCode);
    const confirmed = window.confirm(
      `確定要幫 ${studentName} 申請 ${formatDateTime(lesson.startAt)} 的請假嗎？\n系統會先建立請假紀錄，再同步 Google 日曆並寄出確認信。`
    );
    if (!confirmed) {
      return;
    }

    // 同 applyNormalLeaveImpl：找有沒有先前因雲端 push 失敗暫存的 pending 請假，
    // 有就重用做 retry；沒有就建一筆新的 pendingCloudSync leave。
    let leaveRecord = state.leaveRequests.find((leave) => (
      leave &&
      leave.lessonId === lessonId &&
      !leave.revokedAt &&
      leave.pendingCloudSync &&
      String(leave.type || "normal") === "normal"
    ));
    const isRetryAttempt = Boolean(leaveRecord);
    if (!leaveRecord) {
      leaveRecord = {
        id: newId("LEAVE"),
        lessonId,
        studentCode: lesson.studentCode,
        coachCode: lesson.coachCode,
        calendarEventId: lesson.calendarEventId || "",
        type: "normal",
        submittedAt: new Date().toISOString(),
        submittedBy: activeCoachCode,
        submittedByRole: "coach",
        makeupEligible: true,
        pendingCloudSync: true
      };
      state.leaveRequests.push(leaveRecord);
    }
    addLog(
      isRetryAttempt
        ? `教練 ${activeCoachCode} 重試代學生 ${lesson.studentCode} 之 pending 請假（${leaveRecord.id}）。`
        : `教練 ${activeCoachCode} 代學生 ${lesson.studentCode} 送出請假（課程 ${lesson.id}，等待雲端確認）。`
    );
    saveState();
    renderAll();
    notifyUser(`正在把 ${studentName} 的請假送到雲端，請稍候...`, "info");

    // === 雲端 push 必須先成功才動 GCal / 寄信，避免「GCal 刪了 + 雲端沒紀錄」事故 ===
    let cloudOk = false;
    let cloudErrorMessage = "";
    try {
      cloudOk = await pushCloudLeaveRecord(leaveRecord, lesson);
    } catch (error) {
      cloudErrorMessage = String(error?.message || error || "");
    }

    if (!cloudOk) {
      addLog(`[雲端請假] 教練代請假上傳失敗，本機已暫存 pending 待重試：${leaveRecord.id}${cloudErrorMessage ? ` / ${cloudErrorMessage}` : "（Apps Script 回應 ok=false）"}`);
      saveState();
      renderAll();
      notifyUser(
        `代 ${studentName} 請假尚未成立：雲端同步失敗${cloudErrorMessage ? `（${cloudErrorMessage}）` : ""}。系統會在下次刷新或聚焦時自動重試，或可手動點頂部紅 banner 立即重試。`,
        "warning"
      );
      return;
    }

    // === 雲端確認後才正式生效 ===
    leaveRecord.cloudSyncedAt = new Date().toISOString();
    delete leaveRecord.pendingCloudSync;
    lesson.calendarOccupied = false;
    lesson.attendanceStatus = "leave-normal";
    lesson.updatedAt = new Date().toISOString();
    addLog(`[雲端請假] 已同步教練代請假 ${leaveRecord.id}。`);
    saveState();
    renderAll();
    // 完整課表雲端化：教練代請假後課程轉 leave-normal，同步推上雲讓主系統重算正確。
    pushCloudLessonsQuietly({ coachCode: lesson.coachCode, studentCode: lesson.studentCode });
    notifyUser(`已代 ${studentName} 建立請假紀錄，正在同步 Google 日曆...`, "info");

    const calendarSynced = await tryDeleteCalendarEventForLesson(lesson, "coach_student_normal_leave", false, {
      resolveMissingEventId: true
    });
    if (!calendarSynced) {
      addLog(`[Google日曆] 教練代請假 ${leaveRecord.id} 已建立，但 Google 日曆事件尚未確認刪除，將由下一次同步再處理。`);
      saveState();
    }

    const emailPayload = {
      studentCode: lesson.studentCode,
      coachCode: lesson.coachCode,
      studentEmail: getStudentNoticeEmail(lesson.studentCode),
      coachEmail: getCoachNoticeEmail(lesson.coachCode),
      lessonId: lesson.id,
      lessonStartAt: lesson.startAt,
      leaveId: leaveRecord.id,
      submittedBy: activeCoachCode,
      submittedByRole: "coach"
    };
    leaveRecord.emailNoticeTo = resolveNoticeRecipients(emailPayload).join(", ");
    const sent = await trySendEmailNotice(
      "leave_submitted_by_coach",
      emailPayload,
      `教練代請假通知 ${lesson.studentCode} / ${lesson.id}`
    );
    const emailAt = new Date().toISOString();
    leaveRecord.emailNoticeStatus = sent ? "sent" : "queued-or-skipped";
    if (sent) {
      leaveRecord.emailNoticeSentAt = emailAt;
    } else {
      leaveRecord.emailNoticeQueuedAt = emailAt;
    }
    saveState();
    try {
      await pushCloudLeaveRecord(leaveRecord, lesson);
    } catch (error) {
      addLog(`[雲端請假] 上傳代請假 Email 狀態失敗：${String(error?.message || "未知錯誤")}`);
      saveState();
    }
    renderAll();
    notifyUser(
      sent
        ? `代學生請假完成：${studentName} 的確認信已寄出。`
        : `代學生請假完成，但 Email 尚未成功寄出，請檢查補償任務。`,
      sent ? "success" : "warning"
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
    lesson.updatedAt = nowIso;
    notifyUser("正在取消請假並還原 Google 日曆課程...", "info");

    // 先確認雲端寫入，再做 calendar / email 等不可逆操作。
    // 否則雲端 push 失敗時，calendar event 已建、email 已寄出，但雲端仍認為 leave 還活著
    // → 下次 sync 把 lesson 改回 leave-normal，Google 日曆變成 ghost event。
    let cloudSyncedRevoke = false;
    try {
      cloudSyncedRevoke = await pushCloudLeaveRecord(leave, lesson);
    } catch (error) {
      addLog(`[雲端請假] 取消請假上傳失敗：${String(error?.message || "未知錯誤")}`);
    }
    if (!cloudSyncedRevoke) {
      // 雲端未確認 → 還原本地狀態，請學生重試
      leave.revokedAt = "";
      leave.revokedBy = "";
      leave.makeupEligible = true;
      lesson.attendanceStatus = "leave-normal";
      lesson.calendarOccupied = false;
      state.makeupRequests.forEach((request) => {
        if (request.leaveId === leave.id && request.resolvedAt === nowIso) {
          request.status = "pending";
          request.resolvedAt = "";
          request.rejectReason = "";
        }
      });
      saveState();
      renderAll();
      notifyUser("取消請假未完成：雲端同步失敗，請檢查網路後重試。", "warning");
      return;
    }

    await tryCreateCalendarEventForLesson(lesson, "student_cancel_leave", false);

    addLog(`學生 ${activeStudentCode} 已取消課程 ${lesson.id} 的請假。`);
    saveState();
    renderAll();
    // 完整課表雲端化：取消請假後課程還原 scheduled，同步推上雲讓主系統重算正確。
    pushCloudLessonsQuietly({ coachCode: lesson.coachCode, studentCode: lesson.studentCode });
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
    notifyUser("取消請假流程已完成，課程已還原。", "success");
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

  function isSlotOccupied(slotStartAt, coachCode, excludeRequestId, excludeLessonId, options = {}) {
    const byLesson = options.allowLessonOverlap
      ? false
      : state.lessons.some(
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
    notifyUser(`教練請假時段已新增，影響課程 ${impactedLessons.length} 堂。`, "success");
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
    notifyUser("教練請假時段已移除，原受影響課程已還原。", "success");
  }

  async function submitMakeupRequest() {
    if (isSubmittingMakeupRequest) {
      return;
    }
    if (!el.makeupLeaveSelect || !el.makeupSlotSelect) {
      return;
    }
    isSubmittingMakeupRequest = true;
    const submitButtonText = el.submitMakeupBtn?.textContent || "送出補課申請";
    if (el.submitMakeupBtn) {
      el.submitMakeupBtn.disabled = true;
      el.submitMakeupBtn.textContent = "送出中...";
    }
    try {
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
      // 防 IDOR：學生只能對自己的請假單送補課，避免拿到別人的 leaveId 也能送出
      if (leave.studentCode !== activeStudentCode) {
        alert("這張請假單不是你自己的，無法送出補課申請。");
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
      notifyUser("正在把補課申請同步到雲端...", "info");
      const initialCloudSaved = await saveMakeupRequestToCloud(request, "makeup_request_created");
      if (!initialCloudSaved) {
        request.emailNoticeStatus = "waiting-cloud";
        request.emailNoticeQueuedAt = new Date().toISOString();
        saveState();
        renderAll();
        notifyUser("補課申請尚未完成：雲端還沒有確認收到，所以待審 Email 不會寄出。請保持網路連線並重新開啟學生端，系統會自動補同步。", "warning");
        return;
      }
      const emailStateChanged = await sendMakeupPendingNotice(request);
      renderAll();
      notifyUser(
        request.emailNoticeStatus === "sent"
          ? `補課申請已送出，待審通知已寄出。`
          : emailStateChanged
          ? `補課申請已送出，但 Email 尚未成功寄出，請稍後重送補償任務。`
          : `補課申請已送出，待審資料已同步。`,
        request.emailNoticeStatus === "sent" ? "success" : "warning"
      );
    } finally {
      isSubmittingMakeupRequest = false;
      if (el.submitMakeupBtn) {
        el.submitMakeupBtn.disabled = false;
        el.submitMakeupBtn.textContent = submitButtonText;
      }
    }
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
    await saveStudentScopedStateToCloud(request.studentCode, request.coachCode, "makeup_request_cancelled");
    renderAll();
    notifyUser("補課申請已取消，正在寄送通知...", "info");
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
    notifyUser("補課申請取消流程已完成。", "success");
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
    lesson.updatedAt = nowIso;
    notifyUser("正在取消請假...", "info");

    // 雲端確認後才做不可逆操作（calendar create + email），否則網路抖留 ghost 資料
    let cloudSyncedRevoke = false;
    try {
      cloudSyncedRevoke = await pushCloudLeaveRecord(leave, lesson);
    } catch (error) {
      addLog(`[雲端請假] 教練取消請假上傳失敗：${String(error?.message || "未知錯誤")}`);
    }
    if (!cloudSyncedRevoke) {
      // 還原本地，請教練重試
      leave.revokedAt = "";
      leave.revokedBy = "";
      leave.makeupEligible = true;
      lesson.attendanceStatus = "leave-normal";
      lesson.calendarOccupied = false;
      state.makeupRequests.forEach((request) => {
        if (request.leaveId === leave.id && request.resolvedAt === nowIso) {
          request.status = "pending";
          request.resolvedAt = "";
          request.rejectReason = "";
        }
      });
      saveState();
      renderAll();
      notifyUser("取消請假未完成：雲端同步失敗，請檢查網路後重試。", "warning");
      return;
    }

    await tryCreateCalendarEventForLesson(lesson, "coach_revoke_leave", false);

    addLog(`教練 ${activeCoachCode} 已取消課程 ${lesson.id} 的請假。`);
    saveState();
    renderAll();
    // 完整課表雲端化：教練取消請假後課程還原 scheduled，同步推上雲讓主系統重算正確。
    pushCloudLessonsQuietly({ coachCode: lesson.coachCode, studentCode: lesson.studentCode });
    notifyUser("已取消請假，正在寄送通知...", "info");
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
    notifyUser("取消請假流程已完成。", "success");
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
    notifyUser("補課申請已核准，正在同步雲端...", "info");

    // 雲端確認後再寄信。否則雲端 push 失敗時，學生收到核准信、但其他裝置看不到
    // 補課事件，重複申請或衝堂的風險會出來。
    const cloudSyncedApproval = await saveMakeupRequestToCloud(request, "makeup_approved").catch((error) => {
      addLog(`[雲端同步] 核准補課上傳失敗：${String(error?.message || "未知錯誤")}`);
      return false;
    });
    if (!cloudSyncedApproval) {
      notifyUser("補課已核准（Google 日曆已建立），但雲端同步未完成。系統會自動重試；學生通知信暫緩寄送。", "warning");
      return;
    }

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
    notifyUser("補課核准流程已完成。", "success");
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
    notifyUser("補課申請已退回，正在同步雲端...", "info");

    // 雲端確認後再寄信，否則退回信寄出但雲端還是 pending → 學生會以為申請被退又看到 pending
    const cloudSyncedReject = await saveMakeupRequestToCloud(request, "makeup_rejected").catch((error) => {
      addLog(`[雲端同步] 退回補課上傳失敗：${String(error?.message || "未知錯誤")}`);
      return false;
    });
    if (!cloudSyncedReject) {
      notifyUser("補課已退回，但雲端同步未完成。系統會自動重試；學生通知信暫緩寄送。", "warning");
      return;
    }

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
    notifyUser("補課退回流程已完成。", "success");
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
    // 如果這堂課還掛著未撤銷的請假，按「重置」會把 lesson 改回 scheduled 但
    // leave 紀錄還活著，造成資料不一致（已上線案例：教練端代學生請假後
    // 不知道是誰按了重置，5/10 lesson 變 scheduled 但 LEAVE_09cwcf 還在，
    // 後續 sync 也修不回來）。改成擋住，要求先處理請假。
    if (statusType === "reset") {
      const activeLeave = getLeaveByLesson(lessonId);
      if (activeLeave) {
        alert("這堂課還有未取消的請假，無法重置。\n如要恢復排課，請先點「取消請假」（會同步還原 Google 日曆並通知學生）。");
        return;
      }
    }
    const wasCharged = Boolean(lesson.charged);
    lesson.attendanceStatus = target.attendanceStatus;
    lesson.charged = target.charged;
    lesson.updatedAt = new Date().toISOString();
    addLog(`教練 ${lesson.coachCode} 將課程 ${lesson.id} 標記為 ${target.attendanceStatus}。`);
    saveState();
    renderAll();
    // 完整課表雲端化最關鍵的接入點：臨時請假/未到課/重大急事/重置這四種狀態
    // 只改 lesson 內部欄位、不碰 Google 日曆，光看日曆永遠算不準扣堂。一定要把
    // 改完的 lesson 推上雲，主系統才能用權威算法即時重算正確的「本期已扣」。
    pushCloudLessonsQuietly({ coachCode: lesson.coachCode, studentCode: lesson.studentCode });
    notifyUser(`課程狀態已更新：${getStatusPill(target.attendanceStatus).replace(/<[^>]+>/g, "")}。`, "success");
    if (!wasCharged && target.charged) {
      maybeSendChargeReminder(lesson.studentCode, "coach_mark_status").catch((error) => {
        console.error("billing reminder failed:", error);
      });
    }
  }

  async function rescheduleLessonByCoach(lessonId) {
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
    if (isSlotOccupied(nextStartAt, lesson.coachCode, "", lesson.id, { allowLessonOverlap: true })) {
      alert("新時段與待審補課或教練請假重疊，請改選其他時段。");
      return;
    }

    // 已上線案例：原版只改本地 lesson.startAt 不動 Google 事件，下次 sync 時
    // alignCoachLessonsWithGoogleEvents 會把 Google 舊時間蓋回 local。
    // 必須做：刪舊 Google event → 改本地 + 建新 Google event → 寄通知信。

    const oldStartAt = lesson.startAt;
    const oldEventId = lesson.calendarEventId;
    const oldLessonSnapshot = JSON.parse(JSON.stringify(lesson));

    notifyUser(`正在調整課程時間為 ${formatDateTime(nextStartAt)} ...`, "info");

    // (1) 刪舊 Google event。失敗就 abort 不動本地，避免 Google 留 ghost 事件。
    let deletedOld = false;
    try {
      deletedOld = await tryDeleteCalendarEventForLesson(
        oldLessonSnapshot,
        "coach_reschedule_old",
        false,
        { resolveMissingEventId: true }
      );
    } catch (error) {
      addLog(`[Google日曆] 調整時間：刪除舊事件失敗 ${lesson.id}：${String(error?.message || error)}`);
    }
    if (!deletedOld) {
      notifyUser("調整時間未完成：無法刪除原 Google 日曆事件，請稍後重試。", "warning");
      return;
    }

    // (2) 改本地（雲端 snapshot 會自動把新時間推上去）
    lesson.startAt = nextStartAt;
    lesson.calendarEventId = "";
    addLog(`教練 ${lesson.coachCode} 將課程 ${lesson.id} 由 ${formatDateTime(oldStartAt)} 調整為 ${formatDateTime(nextStartAt)}。`);
    saveState();
    renderAll();

    // (3) 建新 Google event，附上「原時段 → 新時段」的 description 寫進 event 詳情
    const rescheduleDescription = [
      "【課程時間調整】",
      `學生：${getStudentDisplayName(lesson.studentCode)}（${lesson.studentCode}）`,
      `教練：${getCoachByCode(lesson.coachCode)?.name || lesson.coachCode}（${lesson.coachCode}）`,
      `原時段：${formatDateTime(oldStartAt)}`,
      `新時段：${formatDateTime(nextStartAt)}`,
      `課程編號：${lesson.id}`,
      `調整時間：${formatDateTime(new Date().toISOString())}`
    ].join("\n");
    let createdEventId = "";
    try {
      createdEventId = await tryCreateCalendarEventForLesson(lesson, "coach_reschedule_new", false, {
        description: rescheduleDescription
      });
    } catch (error) {
      addLog(`[Google日曆] 調整時間：建立新事件失敗 ${lesson.id}：${String(error?.message || error)}`);
    }
    if (!createdEventId) {
      notifyUser(
        `課程時間已改為 ${formatDateTime(nextStartAt)}，但 Google 日曆新事件尚未建立成功；下次同步會自動補上。`,
        "warning"
      );
    }

    // (4) 通知學生（不可逆，但已放在最後一步且雲端應已透過 snapshot 同步）
    try {
      await trySendEmailNotice(
        "lesson_rescheduled",
        {
          lessonId: lesson.id,
          studentCode: lesson.studentCode,
          coachCode: lesson.coachCode,
          oldStartAt,
          newStartAt: nextStartAt
        },
        `課程時間調整通知 ${lesson.id}`
      );
    } catch (error) {
      addLog(`[Email] 課程調整通知寄送失敗 ${lesson.id}：${String(error?.message || error)}`);
    }

    notifyUser(`課程時間已調整為 ${formatDateTime(nextStartAt)}。`, "success");
  }

  function isLessonPastForStudent(lesson) {
    if (!lesson || !lesson.startAt) {
      return false;
    }
    return new Date(lesson.startAt).getTime() <= Date.now();
  }

  function isLessonAutoCompleted(lesson) {
    return Boolean(
      lesson &&
      lesson.attendanceStatus === "scheduled" &&
      lesson.calendarOccupied &&
      (isGoogleSyncLesson(lesson) || lesson?.sourceType === "MAKEUP") &&
      isLessonPastForStudent(lesson)
    );
  }

  function isLessonChargedForBilling(lesson) {
    // 正常請假（leave-normal）、重大急事（major-case）、教練停課（coach-leave）
    // 依規則一律不扣堂——即使 lesson.charged 因歷史髒資料 / ghost 操作誤標 true。
    // 已上線案例：ST013 5/23 是 leave-normal 卻 charged=true，被誤算進扣堂，
    // 害「本期已扣」多算一堂。臨時請假 temporary-leave / 未到課 no-show 仍要扣，
    // 不在排除名單。
    const status = lesson?.attendanceStatus;
    if (status === "leave-normal" || status === "major-case" || status === "coach-leave") {
      return false;
    }
    return Boolean(isLessonAfterBillingTrackingStart(lesson) && (lesson?.charged || isLessonAutoCompleted(lesson)));
  }

  function markAutoCompletedLessonsForBilling(options = {}) {
    const coachCode = normalizeParticipantCode(options.coachCode || "");
    const studentCode = normalizeParticipantCode(options.studentCode || "");
    const sourceLabel = options.sourceLabel || "日曆同步";
    const affectedStudents = new Set();
    const affectedLessons = [];

    state.lessons.forEach((lesson) => {
      if (coachCode && lesson.coachCode !== coachCode) {
        return;
      }
      if (studentCode && lesson.studentCode !== studentCode) {
        return;
      }
      if (!isLessonAfterBillingTrackingStart(lesson)) {
        return;
      }
      if (!isLessonAutoCompleted(lesson) || lesson.charged) {
        return;
      }
      lesson.charged = true;
      lesson.completedAt = lesson.completedAt || new Date().toISOString();
      lesson.completedBy = lesson.completedBy || "google-calendar-auto";
      affectedStudents.add(lesson.studentCode);
      affectedLessons.push(lesson);
    });

    if (affectedLessons.length) {
      addLog(`[扣堂] ${sourceLabel} 已自動計算已上課 ${affectedLessons.length} 堂並納入扣堂。`);
    }

    return {
      changed: affectedLessons.length > 0,
      count: affectedLessons.length,
      studentCodes: Array.from(affectedStudents)
    };
  }

  // autoComplete 把課標扣堂後，主動把受影響學生的雲端 billing 快照重新 push。
  // 否則主系統 app.js 讀到的「本期已扣」會停在「上次教練手動操作計費面板」
  // 的凍結快照，學生上完新課後一直顯示舊堂數（本期已扣總是錯的根因之一）。
  // pushStudentBillingProfileQuietly 是容錯 fire-and-forget，雲端 upsert by
  // updatedAt，重複 push 安全。
  function repushBillingForAutoCompletedStudents(completionStats, triggerSource) {
    if (!completionStats || !completionStats.changed) {
      return;
    }
    (completionStats.studentCodes || []).forEach((code) => {
      pushStudentBillingProfileQuietly(code, triggerSource || "auto_completed_lesson");
    });
  }

  function getStatusPill(status, lesson, viewMode) {
    let resolvedStatus = status;
    if (status === "scheduled" && isLessonAutoCompleted(lesson)) {
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
          isLessonAfterTrackingDataReset(lesson) &&
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
        <td>${isLessonChargedForBilling(lesson) ? "是" : "否"}</td>
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
          <td>${isLessonChargedForBilling(lesson) ? "是" : "否"}</td>
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
            <div>扣堂：${isLessonChargedForBilling(lesson) ? "是" : "否"}</div>
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
      // 唯讀月曆只給其他教練/觀察者快速看排課，不需要顯示請假/停課狀態；
      // 過濾掉 leave-normal 與 coach-leave 兩種
      const lessonsForReadonly = lessons.filter((lesson) => (
        lesson.attendanceStatus !== "leave-normal" &&
        lesson.attendanceStatus !== "coach-leave"
      ));
      const lessonSnippets = readOnly
        ? [
          renderReadonlyPeriod("上午", lessonsForReadonly.filter((lesson) => getLessonHour(lesson) < 12)),
          renderReadonlyPeriod("下午", lessonsForReadonly.filter((lesson) => getLessonHour(lesson) >= 12))
        ].join("")
        : lessons.slice(0, 3).map((lesson) => {
        const isOnLeave = lesson.attendanceStatus === "leave-normal";
        const isCoachLeave = lesson.attendanceStatus === "coach-leave";
        // status 優先於 sourceType：補課 lesson 被教練請假覆蓋時，typeClass 也得是
        // coach-off / on-leave，否則 makeup 綠底會壓過警告色和刪除線。
        const typeClass = isCoachLeave
          ? "coach-off"
          : (isOnLeave ? "on-leave" : (lesson.sourceType === "MAKEUP" ? "makeup" : ""));
        const prefix = isCoachLeave
          ? "教練請假"
          : (isOnLeave ? "請假" : (lesson.sourceType === "MAKEUP" ? "補課" : "原課"));
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

      let mobileSummary = "";
      if (!readOnly) {
        const summaryParts = [];
        summaryParts.push(lessons.length > 0 ? `${lessons.length} 堂` : "無課");
        if (pendings.length) summaryParts.push(`待審 ${pendings.length}`);
        if (blocks.length) summaryParts.push("教練請假");
        mobileSummary = `<div class="cal-mobile-summary">${summaryParts.join(" · ")}</div>`;
      }

      html += `
          <div class="${classNames}" data-coach-cal-date="${dateKey}" style="${gridPlacement}">
            <div class="cal-date-row"><div class="cal-date">${dayNumber}</div>${todayBadge}</div>
          ${mobileSummary}
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
    const coachStudentLeaveButton = lesson.attendanceStatus === "scheduled" && lesson.calendarOccupied && isGoogleSyncLesson(lesson)
      ? `<button data-coach-lesson-id="${lesson.id}" data-coach-mark="student-leave" class="danger">代學生請假</button>`
      : "";
    return `
      <div class="btn-row">
        <button data-coach-lesson-id="${lesson.id}" data-coach-mark="reschedule">調整時間</button>
        ${coachStudentLeaveButton}
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

  function getLeaveRequestStatusPill(leave) {
    if (leave.revokedAt) {
      return "<span class=\"status cancelled\">已取消</span>";
    }
    return "<span class=\"status normal\">已請假</span>";
  }

  function getLeaveRequestSubmitterText(leave) {
    if (leave.submittedByRole === "coach") {
      return `教練代送 ${leave.submittedBy || ""}`.trim();
    }
    if (leave.submittedByRole === "student") {
      return "學生送出";
    }
    return leave.submittedBy || "-";
  }

  function getLeaveRequestMakeupText(leave) {
    const related = state.makeupRequests
      .filter((request) => request.leaveId === leave.id)
      .sort((a, b) => new Date(b.pendingAt || b.startAt || 0) - new Date(a.pendingAt || a.startAt || 0));
    if (!related.length) {
      return leave.makeupEligible ? "尚未申請補課" : "-";
    }
    const latest = related[0];
    const timeText = latest.startAt ? ` ${formatDateTime(latest.startAt)}` : "";
    return `${getMakeupRequestStatusPill(latest)}${timeText}`;
  }

  function renderCoachStudentLeaveTable() {
    if (!el.coachStudentLeaveTable) {
      return;
    }
    if (!activeCoachCode) {
      el.coachStudentLeaveTable.innerHTML = "<thead><tr><th>請先登入教練</th></tr></thead>";
      return;
    }

    // 顯示今天 00:00 起的請假紀錄，避免「今天但課時已過」被誤過濾。
    const startOfToday = new Date();
    startOfToday.setHours(0, 0, 0, 0);
    const cutoffTime = startOfToday.getTime();
    const normalizedActiveCoach = normalizeParticipantCode(activeCoachCode);
    const leaves = state.leaveRequests
      .filter((leave) => {
        if (normalizeParticipantCode(leave.coachCode) !== normalizedActiveCoach || leave.revokedAt) {
          return false;
        }
        const lesson = getLessonById(leave.lessonId);
        const lessonTime = new Date(lesson?.startAt || "").getTime();
        return !Number.isFinite(lessonTime) || lessonTime >= cutoffTime;
      })
      .sort((a, b) => new Date(b.submittedAt || 0) - new Date(a.submittedAt || 0));

    const rows = leaves.map((leave) => {
      const lesson = getLessonById(leave.lessonId);
      const lessonTime = lesson?.startAt ? formatDateTime(lesson.startAt) : (leave.lessonId || "-");
      return `
        <tr>
          <td>${getStudentDisplayName(leave.studentCode)}</td>
          <td>${lessonTime}</td>
          <td>${leave.submittedAt ? formatDateTime(leave.submittedAt) : "-"}</td>
          <td>${getLeaveRequestSubmitterText(leave)}</td>
          <td>${getLeaveRequestStatusPill(leave)}</td>
          <td>${getLeaveRequestMakeupText(leave)}</td>
        </tr>
      `;
    }).join("");

    el.coachStudentLeaveTable.innerHTML = `
      <thead><tr><th>學生</th><th>原課程時間</th><th>請假時間</th><th>送出方式</th><th>請假狀態</th><th>補課狀態</th></tr></thead>
      <tbody>${rows || "<tr><td colspan=\"6\">今天起目前沒有學生請假紀錄</td></tr>"}</tbody>
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
        <td>${getMakeupRequestStatusPill(request)}</td>
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
      .filter((request) => (
        request.coachCode === activeCoachCode &&
        request.status === "pending" &&
        !isMakeupRequestWaitingForCloud(request)
      ))
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
      } else if (task.type === "createEvent") {
        // 重送前先查同時段是否已有事件（多半上一次其實建成了、只是當下驗證
        // 沒讀到）。有就直接綁定、視為成功，不再重建——否則補償重送會一直
        // 疊出重複的「課程…」事件。
        const lesson = finalPayload?.lessonId ? getLessonById(finalPayload.lessonId) : null;
        if (lesson) {
          let existingEvent = null;
          try {
            existingEvent = await resolveSingleCalendarEventForLesson(
              lesson,
              finalPayload.eventId || lesson.calendarEventId || ""
            );
          } catch (error) {
            existingEvent = null;
          }
          if (existingEvent) {
            lesson.calendarEventId = String(existingEvent.eventId || lesson.calendarEventId || "").trim();
            task.status = "completed";
            task.completedAt = new Date().toISOString();
            task.lastError = "";
            addLog(`[補償重送] createEvent 視為成功（該時段已有事件，不重建，${task.id}）。`);
            saveState();
            renderAll();
            return;
          }
        }
      } else if (task.type === "deleteEvent" || task.type === "deleteSingleEvent") {
        Object.assign(finalPayload, addSingleEventDeleteGuards(finalPayload));
        task.payload = finalPayload;
        if (await isDeleteOccurrenceAlreadyAbsent(finalPayload, getLessonById(finalPayload.lessonId))) {
          task.status = "completed";
          task.completedAt = new Date().toISOString();
          task.lastError = "";
          const lesson = getLessonById(finalPayload.lessonId);
          if (lesson) {
            lesson.calendarEventId = "";
          }
          addLog(`[補償重送] ${task.type} 視為成功（該日期已無此單堂事件，${task.id}）。`);
          saveState();
          renderAll();
          return;
        }
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
    const tasks = (state.compensationTasks || [])
      .filter((item) => item.status !== "completed")
      .slice()
      .sort((a, b) => new Date(b.createdAt || 0) - new Date(a.createdAt || 0));
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
    if (isKnownFakeStudentRecord(student)) {
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
      .filter((lesson) => (
        new Date(lesson.startAt).getTime() <= now &&
        lesson.attendanceStatus !== "leave-normal" &&
        lesson.attendanceStatus !== "coach-leave" &&
        lesson.attendanceStatus !== "calendar-removed"
      ))
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
      const historyLessons = (state.lessons || [])
        .filter((lesson) => lesson.studentCode === student.code && isLessonAfterTrackingDataReset(lesson));
      const lastLesson = getLastCompletedLesson(historyLessons);
      const paymentClass = billingCycle.effectivePaymentStatus === "unpaid"
        ? "rejected"
        : billingCycle.effectivePaymentStatus === "paid"
          ? "approved"
          : "pending";
      return `
        <tr>
          <td>${student.name || student.code}</td>
          <td>${lastLesson ? formatDateTime(lastLesson.startAt) : "尚未上課"}</td>
          <td>${billingCycle.currentCycleChargedCount}</td>
          <td><span class="status ${paymentClass}">${getPaymentStatusLabel(billingCycle.effectivePaymentStatus)}</span></td>
        </tr>
      `;
    }).join("");
    el.studentOverviewTable.innerHTML = `
      <thead><tr><th>學生</th><th>最後上課日期</th><th>累計扣堂</th><th>繳費狀態</th></tr></thead>
      <tbody>${rows || "<tr><td colspan=\"4\">目前沒有學生資料</td></tr>"}</tbody>
    `;
  }

  function renderBillingPanels() {
    renderStudentOverviewPanel();
    renderChargePanel();
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
      <div class="metric"><div class="k">導入前已扣堂</div><div class="v">${stats.startCount}</div></div>
      <div class="metric"><div class="k">系統內已扣堂</div><div class="v">${stats.chargedLessons.length}</div></div>
      <div class="metric"><div class="k">累計已扣堂</div><div class="v">${stats.totalChargedCount}</div></div>
      <div class="metric"><div class="k">本期已扣</div><div class="v">${billingCycle.currentCycleChargedCount}/${CHARGE_REMINDER_STEP}</div></div>
      <div class="metric"><div class="k">距下次繳費</div><div class="v">${billingCycle.remainingToNextPayment}</div></div>
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
        <td>${getStatusPill(lesson.attendanceStatus, lesson)}</td>
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
    renderCoachStudentLeaveTable();
    renderCalendarRemovedPanel();
    renderCompensationPanel();
    renderLogs();
    renderStudentOverviewPanel();
    renderChargePanel();
    renderCloudSnapshotControls();
    updateCoachReadOnlyUi();
    renderPendingSyncBanner();
    populateManualLessonStudentSelect();
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

  async function refreshCloudBillingProfilesFromCurrentSession(force = false) {
    if (!activeCoachCode && !activeStudentCode) {
      return false;
    }
    const now = Date.now();
    if (!force && now - lastCloudBillingRefreshAt < CLOUD_BILLING_REFRESH_INTERVAL_MS) {
      return false;
    }
    lastCloudBillingRefreshAt = now;
    const changed = await syncCloudBillingProfiles({
      coachCode: activeCoachCode,
      studentCode: activeStudentCode
    });
    if (changed) {
      renderBillingPanels();
    }
    return changed;
  }

  async function refreshCloudSessionDataFromCurrentSession(force = false) {
    if (!activeCoachCode && !activeStudentCode) {
      return false;
    }
    const now = Date.now();
    if (!force && now - lastCloudSessionRefreshAt < CLOUD_BILLING_REFRESH_INTERVAL_MS) {
      return false;
    }
    lastCloudSessionRefreshAt = now;

    let changed = false;
    // 先補推任何 pending 請假，確保後續 cloud read 看到的是最新狀態，也讓 banner
    // 在使用者切回分頁時就能消掉。silent 模式不彈 notifyUser，避免每 45 秒打擾使用者；
    // banner 持續可見已足以提示。
    try {
      const retryResult = await retryUnsyncedLocalLeaves({ silent: true });
      if (retryResult && retryResult.succeeded > 0) {
        changed = true;
      }
    } catch (error) {
      console.warn("Pending leave retry on session refresh failed:", error);
    }
    try {
      changed = await pullLeaveStateSnapshotFromCloud({
        coachCode: activeCoachCode,
        studentCode: activeStudentCode
      }) || changed;
    } catch (error) {
      console.warn("Cloud leave snapshot refresh failed:", error);
    }
    try {
      changed = await syncCloudLeaveRecords({
        coachCode: activeCoachCode,
        studentCode: activeStudentCode
      }) || changed;
    } catch (error) {
      console.warn("Cloud leave session refresh failed:", error);
    }
    try {
      changed = await syncCloudBillingProfiles({
        coachCode: activeCoachCode,
        studentCode: activeStudentCode
      }) || changed;
    } catch (error) {
      console.warn("Cloud billing session refresh failed:", error);
    }

    try {
      if (activeStudentCode) {
        changed = await syncStudentCalendarEventsFromGoogle() || changed;
      } else if (activeCoachCode && activeCoachReadOnly) {
        changed = await syncReadOnlyCoachCalendarFromGoogle(activeCoachCode) || changed;
      } else if (activeCoachCode) {
        changed = await syncCoachCalendarEventsFromGoogle({ showStatus: false }) || changed;
      }
    } catch (error) {
      console.warn("Google calendar session refresh failed:", error);
    }

    try {
      changed = await resolvePendingDeleteCompensationTasksFromGoogle({
        coachCode: activeCoachCode,
        studentCode: activeStudentCode
      }) || changed;
    } catch (error) {
      console.warn("Delete compensation cleanup failed:", error);
    }

    if (activeStudentCode && activeCoachCode && hasStudentMakeupRequestsForCloudSync(activeStudentCode)) {
      const makeupCloudSaved = await saveStudentScopedStateToCloud(activeStudentCode, activeCoachCode, "student_session_makeup_sync");
      changed = makeupCloudSaved || changed;
      if (makeupCloudSaved) {
        changed = await sendQueuedMakeupPendingNoticesForStudent(activeStudentCode) || changed;
      }
    }

    if (changed) {
      renderAll();
    }
    return changed;
  }

  function bindEvents() {
    window.addEventListener("focus", () => {
      refreshCloudSessionDataFromCurrentSession().catch((error) => {
        console.warn("Cloud session focus refresh failed:", error);
      });
    });
    document.addEventListener("visibilitychange", () => {
      if (document.hidden) {
        return;
      }
      refreshCloudSessionDataFromCurrentSession().catch((error) => {
        console.warn("Cloud session visibility refresh failed:", error);
      });
    });
    if (el.pendingLeaveRetryBtn) {
      el.pendingLeaveRetryBtn.addEventListener("click", async () => {
        if (el.pendingLeaveRetryBtn.disabled) {
          return;
        }
        const originalText = el.pendingLeaveRetryBtn.textContent;
        el.pendingLeaveRetryBtn.disabled = true;
        el.pendingLeaveRetryBtn.textContent = "重試中...";
        try {
          await retryUnsyncedLocalLeaves();
        } catch (error) {
          notifyUser(`重試失敗：${String(error?.message || error)}`, "warning");
        } finally {
          el.pendingLeaveRetryBtn.disabled = false;
          el.pendingLeaveRetryBtn.textContent = originalText || "立即重試";
          renderAll();
        }
      });
    }
    if (el.studentLoginBtn) {
      el.studentLoginBtn.addEventListener("click", async () => {
        await syncCoachflowRosterFromCloud();
        syncCoachflowRosterFromLocalState();
        const loaded = activateStudentSession(el.studentCode?.value, el.studentCoachCode?.value, false);
        if (loaded) {
          notifyUser("學生資料已載入，正在同步 Google 日曆與請假紀錄...", "info");
          // 救援：把本機已有但雲端可能因配額失敗而沒收到的請假紀錄重新推一次。
          // saveLeaveRecord 端是 upsert by id，安全可重跑。
          try {
            await retryUnsyncedLocalLeaves();
          } catch (error) {
            console.warn("Retry unsynced leaves failed on student login:", error);
          }
          let leaveChanged = false;
          let cloudLeaveFailed = false;
          try {
            leaveChanged = await pullLeaveStateSnapshotFromCloud({ studentCode: activeStudentCode, coachCode: activeCoachCode }) || leaveChanged;
            leaveChanged = await syncCloudLeaveRecords({ studentCode: activeStudentCode, coachCode: activeCoachCode }) || leaveChanged;
          } catch (error) {
            console.warn("Cloud leave sync failed for student login:", error);
            cloudLeaveFailed = true;
          }
          let billingChanged = false;
          try {
            billingChanged = await syncCloudBillingProfiles({ studentCode: activeStudentCode, coachCode: activeCoachCode });
          } catch (error) {
            console.warn("Cloud billing sync failed for student login:", error);
          }
          let calendarChanged = false;
          try {
            calendarChanged = await syncStudentCalendarEventsFromGoogle();
          } catch (error) {
            console.warn("Student Google calendar sync failed for student login:", error);
          }
          let makeupCloudSaved = false;
          if (hasStudentMakeupRequestsForCloudSync(activeStudentCode)) {
            makeupCloudSaved = await saveStudentScopedStateToCloud(activeStudentCode, activeCoachCode, "student_login_makeup_sync");
          }
          let makeupNoticeChanged = false;
          if (makeupCloudSaved) {
            makeupNoticeChanged = await sendQueuedMakeupPendingNoticesForStudent(activeStudentCode);
          }
          if (leaveChanged || billingChanged || calendarChanged || makeupCloudSaved || makeupNoticeChanged) {
            renderAll();
          }
          if (cloudLeaveFailed) {
            notifyUser("學生資料已載入，但雲端請假紀錄同步失敗，請稍後再試。", "warning");
          } else {
            notifyUser("學生資料已更新完成，可以查看今日上課與請假狀態。", "success");
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
        if (!loaded) {
          return;
        }
        notifyUser("教練資料已載入，正在同步雲端資料與 Google 日曆...", "info");
        const syncFromCloud = async () => {
          const cloudChanged = await syncCoachflowRosterFromCloud();
          const localChanged = syncCoachflowRosterFromLocalState();
          let leaveChanged = false;
          try {
            leaveChanged = await pullLeaveStateSnapshotFromCloud({ coachCode: activeCoachCode }) || leaveChanged;
            leaveChanged = await syncCloudLeaveRecords({ coachCode: activeCoachCode }) || leaveChanged;
          } catch (error) {
            console.warn("Cloud leave sync failed for coach login:", error);
          }
          let billingChanged = false;
          try {
            billingChanged = await syncCloudBillingProfiles({ coachCode: activeCoachCode });
          } catch (error) {
            console.warn("Cloud billing sync failed for coach login:", error);
          }
          if (cloudChanged || localChanged || leaveChanged || billingChanged) {
            renderAll();
          }
          return Boolean(cloudChanged || localChanged || leaveChanged || billingChanged);
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
        try {
          const calendarChanged = await syncCoachCalendarEventsFromGoogle({ showStatus: true });
          const compensationCleaned = await resolvePendingDeleteCompensationTasksFromGoogle({ coachCode: activeCoachCode });
          if (calendarChanged || compensationCleaned) {
            renderAll();
          }
          notifyUser("教練資料與 Google 日曆同步完成。", "success");
        } catch (error) {
          console.warn("Coach Google calendar sync failed for coach login:", error);
          if (el.coachCalendarSyncText) {
            el.coachCalendarSyncText.textContent = "自動同步 Google 日曆失敗，請稍後再試或手動同步。";
          }
          notifyUser("教練資料已載入，但 Google 日曆同步失敗。", "warning");
        }
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
          if (mark === "student-leave") {
            applyNormalLeaveByCoach(lessonId);
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
    if (el.manualLessonAddBtn) {
      el.manualLessonAddBtn.addEventListener("click", async () => {
        if (el.manualLessonAddBtn.disabled) return;
        const originalText = el.manualLessonAddBtn.textContent;
        el.manualLessonAddBtn.disabled = true;
        el.manualLessonAddBtn.textContent = "新增中...";
        try {
          await addManualStudentLesson();
        } catch (error) {
          notifyUser(`新增失敗：${String(error?.message || error)}`, "warning");
        } finally {
          el.manualLessonAddBtn.disabled = false;
          el.manualLessonAddBtn.textContent = originalText || "新增學生上課";
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
      el.coachReviewRefreshBtn.addEventListener("click", async () => {
        if (activeCoachCode) {
          try {
            await pullLeaveStateSnapshotFromCloud({ coachCode: activeCoachCode });
            await syncCloudLeaveRecords({ coachCode: activeCoachCode });
          } catch (error) {
            console.warn("Coach pending makeup refresh failed:", error);
            notifyUser("待審清單雲端重新整理失敗，請稍後再試。", "warning");
          }
        }
        renderAll();
      });
    }

    if (el.coachStudentLeaveRefreshBtn) {
      el.coachStudentLeaveRefreshBtn.addEventListener("click", async () => {
        if (!activeCoachCode) {
          notifyUser("請先登入教練。", "warning");
          return;
        }
        el.coachStudentLeaveRefreshBtn.disabled = true;
        el.coachStudentLeaveRefreshBtn.textContent = "同步中…";
        try {
          await pullLeaveStateSnapshotFromCloud({ coachCode: activeCoachCode });
          await syncCloudLeaveRecords({ coachCode: activeCoachCode });
          renderAll();
          notifyUser("學生請假紀錄已從雲端重新整理。", "success");
        } catch (error) {
          console.warn("Coach student leave refresh failed:", error);
          notifyUser("雲端重新整理失敗，請稍後再試。", "warning");
        } finally {
          el.coachStudentLeaveRefreshBtn.disabled = false;
          el.coachStudentLeaveRefreshBtn.textContent = "重新整理";
        }
      });
    }

    if (el.compensationRefreshBtn) {
      el.compensationRefreshBtn.addEventListener("click", async () => {
        try {
          const changed = await resolvePendingDeleteCompensationTasksFromGoogle({
            coachCode: activeCoachCode,
            studentCode: activeStudentCode
          });
          if (changed) {
            renderAll();
            return;
          }
        } catch (error) {
          console.warn("Delete compensation refresh cleanup failed:", error);
        }
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
    if (el.coachCloudUploadBtn) {
      el.coachCloudUploadBtn.addEventListener("click", () => {
        if (!requireCoachWriteAccess()) {
          return;
        }
        uploadLocalLeaveStateToCloud();
      });
    }
    if (el.coachCloudDownloadBtn) {
      el.coachCloudDownloadBtn.addEventListener("click", () => {
        if (!requireCoachWriteAccess()) {
          return;
        }
        downloadCloudLeaveStateToLocal();
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
    collapseDuplicateGoogleLessons();
    ensureMakeupCodes();
    ensureLessonCalendarEventIds();
    ensureParticipantEmails();
    ensureStudentBillingProfiles();
    resetAllTrackingDataOnce();
    purgePreResetTrackingDataOnce();
    resetImportedChargedCountsOnce();
    restoreAutoRemovedLocalLessonsOnce();
    syncCoachflowRosterFromLocalState();

    const cloudSyncPromise = (async () => {
      const syncedFromCloud = await syncCoachflowRosterFromCloud();
      const syncedFromLocal = syncCoachflowRosterFromLocalState();
      let billingSynced = false;
      try {
        billingSynced = await syncCloudBillingProfiles();
      } catch (error) {
        console.warn("Cloud billing sync failed during bootstrap roster sync:", error);
      }
      return Boolean(syncedFromCloud || syncedFromLocal || billingSynced);
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
      let cloudLeaveSynced = false;
      let cloudBillingSynced = false;
      if ((autoStudentLoaded || autoCoachLoaded) && (activeStudentCode || activeCoachCode)) {
        try {
          cloudLeaveSynced = await pullLeaveStateSnapshotFromCloud({
            studentCode: activeStudentCode,
            coachCode: activeCoachCode
          }) || cloudLeaveSynced;
          cloudLeaveSynced = await syncCloudLeaveRecords({
            studentCode: activeStudentCode,
            coachCode: activeCoachCode
          }) || cloudLeaveSynced;
        } catch (error) {
          console.warn("Cloud leave sync failed during bootstrap:", error);
        }
        try {
          cloudBillingSynced = await syncCloudBillingProfiles({
            studentCode: activeStudentCode,
            coachCode: activeCoachCode
          });
        } catch (error) {
          console.warn("Cloud billing sync failed during bootstrap:", error);
        }
      }
      let studentCalendarSynced = false;
      if (autoStudentLoaded && activeStudentCode) {
        studentCalendarSynced = await syncStudentCalendarEventsFromGoogle();
      }
      let readOnlyCalendarSynced = false;
      if (autoCoachLoaded && activeCoachReadOnly && activeCoachCode) {
        readOnlyCalendarSynced = await syncReadOnlyCoachCalendarFromGoogle(activeCoachCode);
      }
      let coachCalendarSynced = false;
      if (autoCoachLoaded && !activeCoachReadOnly && activeCoachCode) {
        coachCalendarSynced = await syncCoachCalendarEventsFromGoogle({ showStatus: true });
      }
      let deleteCompensationCleaned = false;
      if (autoStudentLoaded || autoCoachLoaded) {
        deleteCompensationCleaned = await resolvePendingDeleteCompensationTasksFromGoogle({
          studentCode: activeStudentCode,
          coachCode: activeCoachCode
        });
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
      if (cloudBillingSynced && !cloudSyncChanged && !studentCalendarSynced && !cloudLeaveSynced) {
        renderAll();
      }
      if (readOnlyCalendarSynced && !cloudSyncChanged) {
        renderAll();
      }
      if (coachCalendarSynced && !cloudSyncChanged && !studentCalendarSynced && !cloudLeaveSynced && !cloudBillingSynced && !readOnlyCalendarSynced) {
        renderAll();
      }
      if (deleteCompensationCleaned && !cloudSyncChanged && !studentCalendarSynced && !cloudLeaveSynced && !cloudBillingSynced && !readOnlyCalendarSynced && !coachCalendarSynced) {
        renderAll();
      }
    } catch (error) {
      console.warn("CoachFlow cloud sync completed with error:", error);
    }
  }

  bootstrap();
})();
