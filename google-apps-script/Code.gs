const SHEETS = {
  coaches: "Coaches",
  students: "Students",
  programs: "Programs",
  programItems: "ProgramItems",
  workoutLogs: "WorkoutLogs",
  leaveRecords: "LeaveRecords",
  billingProfiles: "BillingProfiles",
  lessons: "Lessons",
  makeupRequests: "MakeupRequests",
  coachBlocks: "CoachBlocks"
};
const APP_TIME_ZONE = "Asia/Taipei";

const SCHEMA = {
  Coaches: [
    "coach_id",
    "coach_name",
    "status",
    "role",
    "access_code",
    "token",
    "last_used_at",
    "created_at",
    "updated_at"
  ],
  Students: [
    "student_id",
    "student_name",
    "status",
    "access_code",
    "token",
    "primary_coach_id",
    "primary_coach_name",
    "last_used_at",
    "created_at",
    "updated_at"
  ],
  Programs: [
    "program_id",
    "program_code",
    "program_date",
    "title",
    "coach_id",
    "coach_name",
    "notes",
    "published",
    "created_at",
    "updated_at",
    "target_student_ids"
  ],
  ProgramItems: [
    "item_id",
    "program_id",
    "sort_order",
    "category",
    "exercise",
    "target_sets",
    "target_type",
    "target_value",
    "item_note",
    "updated_at"
  ],
  WorkoutLogs: [
    "log_id",
    "program_id",
    "program_code",
    "program_date",
    "coach_id",
    "coach_name",
    "student_id",
    "student_name",
    "category",
    "exercise",
    "target_sets",
    "target_type",
    "target_value",
    "actual_weight",
    "actual_sets",
    "actual_reps",
    "student_note",
    "submitted_at",
    "updated_at"
  ],
  // 請假紀錄表：從 PropertiesService.LEAVE_RECORDS_V1 遷移過來。
  // 欄位對齊 normalizeLeaveRecord_，每筆 leave 一列。
  LeaveRecords: [
    "id",
    "lessonId",
    "lessonKey",
    "calendarEventId",
    "studentCode",
    "coachCode",
    "lessonStartAt",
    "type",
    "submittedAt",
    "submittedBy",
    "submittedByRole",
    "makeupEligible",
    "emailNoticeStatus",
    "emailNoticeAt",
    "revokedAt",
    "revokedBy",
    "createdAt",
    "updatedAt"
  ],
  // 學生扣費 profile：從 PropertiesService.BILLING_PROFILES_V1 遷移過來。
  // chargeReminderLogs 原本是物件陣列，cell 存不下 → 序列化成 JSON 存在
  // chargeReminderLogsJson，讀回時 parse。
  BillingProfiles: [
    "studentCode",
    "coachCode",
    "studentName",
    "email",
    "emailUpdatedAt",
    "emailUpdatedBy",
    "chargeStartCount",
    "chargeStartCountUpdatedAt",
    "chargeStartCountUpdatedBy",
    "paidThroughCount",
    "paymentStatus",
    "paymentNote",
    "paymentConfirmedAt",
    "paymentConfirmedBy",
    "chargeReminderLogsJson",
    "chargeReminderStep",
    "systemChargedCount",
    "totalChargedCount",
    "currentCycleChargedCount",
    "remainingToNextPayment",
    "nextPaymentDueCount",
    "effectivePaymentStatus",
    "updatedAt",
    "updatedBy",
    "createdAt"
  ],
  // 完整課表雲端化：每堂課一列，計費要算的「每堂 attendanceStatus + charged」權威來源。
  // 過去課表只散落在「開過請假系統那台裝置的 localStorage」+ Google Calendar，
  // 但臨時請假/未到課（扣堂）、重大急事（不扣堂）只改 lesson 內部欄位、不碰日曆，
  // 光看日曆算不準 → 主系統學生「本期已扣 X/4」會錯。把完整課表存上雲端，主系統
  // 直接讀完整課表用權威算法即時重算才是真根治。
  // calendarOccupied / charged 是 boolean，normalizeLessonRow_ 做 round-trip 防誤判。
  Lessons: [
    "id",
    "studentCode",
    "coachCode",
    "startAt",
    "sourceType",
    "calendarOccupied",
    "attendanceStatus",
    "charged",
    "completedAt",
    "updatedAt",
    "calendarEventId",
    "cancelledByCoachLeaveBlockId",
    "coachLeaveReason",
    "beforeCoachLeaveJson"
  ],
  MakeupRequests: [
    "id",
    "code",
    "leaveId",
    "lessonId",
    "studentCode",
    "coachCode",
    "status",
    "pendingAt",
    "resolvedAt",
    "startAt",
    "rejectReason",
    "emailNoticeStatus",
    "emailNoticeAt",
    "cloudSyncStatus",
    "cloudVerifiedAt",
    "createdAt",
    "updatedAt"
  ],
  CoachBlocks: [
    "id",
    "coachCode",
    "startAt",
    "endAt",
    "reason",
    "revokedAt",
    "createdAt",
    "updatedAt"
  ]
};

function doGet(e) {
  const action = getAction_(e);

  try {
    if (shouldEnsureSheetsForAction_(action)) {
      ensureSheets_();
    }

    switch (action) {
      case "resolveCoachAccess":
        return jsonResponse_(resolveCoachAccessResponse_(getParam_(e, "access")));

      case "resolveStudentAccess":
        return jsonResponse_(resolveStudentAccessResponse_(getParam_(e, "access")));

      case "bootstrapAdmin":
        requireAdminAccess_(e && e.parameter ? e.parameter : {});
        return jsonResponse_(buildFullBootstrap_());

      case "bootstrapCoach":
        requireCoachReadAccess_(e && e.parameter ? e.parameter : {}, getParam_(e, "coachId"));
        return jsonResponse_(buildCoachBootstrap_(getParam_(e, "coachId")));

      case "bootstrapStudent":
        requireStudentReadAccess_(e && e.parameter ? e.parameter : {}, getParam_(e, "studentId"));
        return jsonResponse_(buildStudentBootstrap_(getParam_(e, "studentId")));

      case "ping":
        return jsonResponse_(buildPingResponse_(action));

      case "checkEvent":
        return jsonResponse_(checkCalendarEvent_(e && e.parameter ? e.parameter : {}));

      case "listEvents":
        return jsonResponse_(listCalendarEvents_(e && e.parameter ? e.parameter : {}));

      case "listLeaveRecords":
        return jsonResponse_(listLeaveRecords_(e && e.parameter ? e.parameter : {}));

      case "listBillingProfiles":
        return jsonResponse_(listBillingProfiles_(e && e.parameter ? e.parameter : {}));

      case "listLessons":
        return jsonResponse_(listLessons_(e && e.parameter ? e.parameter : {}));

      case "listMakeupRequests":
        return jsonResponse_(listMakeupRequests_(e && e.parameter ? e.parameter : {}));

      case "listCoachBlocks":
        return jsonResponse_(listCoachBlocks_(e && e.parameter ? e.parameter : {}));

      case "getLeaveStateSnapshot":
        return jsonResponse_(getLeaveStateSnapshot_(e && e.parameter ? e.parameter : {}));

      case "bootstrap":
      default:
        requireAdminAccess_(e && e.parameter ? e.parameter : {});
        return jsonResponse_(buildFullBootstrap_());
    }
  } catch (error) {
    return jsonResponse_({
      ok: false,
      message: error.message || String(error)
    });
  }
}

function doPost(e) {
  const payload = parsePostBody_(e);
  const action = payload.action || getAction_(e) || "";

  try {
    if (shouldEnsureSheetsForAction_(action)) {
      ensureSheets_();
    }

    switch (action) {
      case "ping":
        return jsonResponse_(buildPingResponse_(action));

      case "resolveCoachAccess":
        return jsonResponse_(resolveCoachAccessResponse_(payload.access));

      case "resolveStudentAccess":
        return jsonResponse_(resolveStudentAccessResponse_(payload.access));

      case "bootstrapAdmin":
        requireAdminAccess_(payload);
        return jsonResponse_(buildFullBootstrap_());

      case "bootstrapCoach":
        requireCoachReadAccess_(payload, payload.coachId);
        return jsonResponse_(buildCoachBootstrap_(payload.coachId));

      case "bootstrapStudent":
        requireStudentReadAccess_(payload, payload.studentId);
        return jsonResponse_(buildStudentBootstrap_(payload.studentId));

      case "upsertCoach":
        if (isTestPollutionBlocked_() && coachLooksLikeTestPollution_(payload.coach || {})) {
          return jsonResponse_(testPollutionRejection_("upsertCoach"));
        }
        requireCoachUpsertAccess_(payload, payload.coach || {});
        upsertCoach_(payload.coach || {});
        return jsonResponse_(buildFullBootstrap_());

      case "setCoachStatus":
        requireAdminAccess_(payload);
        setCoachStatus_(payload.coachId, payload.status);
        return jsonResponse_(buildFullBootstrap_());

      case "deleteCoach":
        requireAdminAccess_(payload);
        deleteCoach_(payload.coachId);
        return jsonResponse_(buildFullBootstrap_());

      case "upsertStudent":
        if (isTestPollutionBlocked_() && studentLooksLikeTestPollution_(payload.student || {})) {
          return jsonResponse_(testPollutionRejection_("upsertStudent"));
        }
        requireStudentUpsertAccess_(payload, payload.student || {});
        upsertStudent_(payload.student || {});
        return jsonResponse_(buildFullBootstrap_());

      case "assignStudentCoach":
        requireStudentMutationAccess_(payload, payload.studentId);
        assignStudentCoach_(payload.studentId, payload.coachId);
        return jsonResponse_(buildFullBootstrap_());

      case "setStudentStatus":
        requireStudentMutationAccess_(payload, payload.studentId);
        setStudentStatus_(payload.studentId, payload.status);
        return jsonResponse_(buildFullBootstrap_());

      case "deleteStudent":
        requireStudentMutationAccess_(payload, payload.studentId);
        deleteStudent_(payload.studentId);
        return jsonResponse_(buildFullBootstrap_());

      case "saveProgram":
        if (isTestPollutionBlocked_() && programLooksLikeTestPollution_(payload.program || {})) {
          return jsonResponse_(testPollutionRejection_("saveProgram"));
        }
        requireProgramSaveAccess_(payload, payload.program || {});
        saveProgram_(payload.program || {}, payload.items || []);
        return jsonResponse_(buildFullBootstrap_());

      case "publishProgram":
        requireProgramMutationAccess_(payload, payload.programId, payload.coachId);
        publishProgram_(payload.programId, payload.coachId);
        return jsonResponse_(buildFullBootstrap_());

      case "deleteProgram":
        requireProgramMutationAccess_(payload, payload.programId || payload.program_id);
        deleteProgram_(payload.programId || payload.program_id);
        return jsonResponse_(buildFullBootstrap_());

      case "submitWorkoutLogs":
        requireWorkoutLogSubmitAccess_(payload, payload.logs || []);
        submitWorkoutLogs_(payload.logs || []);
        return jsonResponse_(buildFullBootstrap_());

      case "importWorkoutLogs":
        requireCoachWriteAccess_(payload);
        importWorkoutLogs_(payload.logs || []);
        return jsonResponse_(buildFullBootstrap_());

      case "touchCoach":
        requireCoachReadAccess_(payload, payload.coachId);
        touchCoach_(payload.coachId);
        return jsonResponse_(buildCoachBootstrap_(payload.coachId));

      case "touchStudent":
        requireStudentReadAccess_(payload, payload.studentId);
        touchStudent_(payload.studentId);
        return jsonResponse_(buildStudentBootstrap_(payload.studentId));

      case "createEvent":
        requireLeaveScopedWriteAccess_(payload, payload.coachCode, payload.studentCode);
        return jsonResponse_(createCalendarEvent_(payload));

      case "updateEvent":
        requireLeaveScopedWriteAccess_(payload, payload.coachCode, payload.studentCode);
        return jsonResponse_(updateCalendarEvent_(payload));

      case "checkEvent":
        return jsonResponse_(checkCalendarEvent_(payload));

      case "listEvents":
        return jsonResponse_(listCalendarEvents_(payload));

      case "deleteEvent":
        requireLeaveScopedWriteAccess_(payload, payload.coachCode, payload.studentCode);
        return jsonResponse_(deleteCalendarEvent_(payload));

      case "deleteSingleEvent":
        requireLeaveScopedWriteAccess_(payload, payload.coachCode, payload.studentCode);
        return jsonResponse_(deleteCalendarEvent_(Object.assign({}, payload || {}, {
          deleteScope: "single",
          singleEventOnly: true,
          strictSingleOccurrence: true
        })));

      case "sendEmail":
        requireLeaveScopedWriteAccess_(payload, payload.coachCode, payload.studentCode);
        return jsonResponse_(sendEmailNotice_(payload));

      case "listLeaveRecords":
        return jsonResponse_(listLeaveRecords_(payload));

      case "saveLeaveRecord":
        requireLeaveRecordWriteAccess_(payload);
        return jsonResponse_(saveLeaveRecord_(payload));

      case "listBillingProfiles":
        return jsonResponse_(listBillingProfiles_(payload));

      case "saveBillingProfile":
        if (isTestPollutionBlocked_() &&
          looksLikeTestParticipantName_((payload.profile || payload).studentName)) {
          return jsonResponse_(testPollutionRejection_("saveBillingProfile"));
        }
        requireBillingProfileWriteAccess_(payload);
        return jsonResponse_(saveBillingProfile_(payload));

      case "cleanupLeaveTestData":
        requireLeaveScopedWriteAccess_(payload, payload.coachCode, "");
        return jsonResponse_(cleanupLeaveTestData_(payload));

      case "listLessons":
        return jsonResponse_(listLessons_(payload));

      case "saveLessons":
        requireLessonsWriteAccess_(payload);
        return jsonResponse_(saveLessons_(payload));

      case "listMakeupRequests":
        return jsonResponse_(listMakeupRequests_(payload));

      case "saveMakeupRequest":
        requireLeaveScopedWriteAccess_(payload, (payload.request || payload).coachCode, (payload.request || payload).studentCode);
        return jsonResponse_(saveMakeupRequest_(payload));

      case "listCoachBlocks":
        return jsonResponse_(listCoachBlocks_(payload));

      case "saveCoachBlock":
        requireLeaveScopedWriteAccess_(payload, (payload.block || payload).coachCode, "");
        return jsonResponse_(saveCoachBlock_(payload));

      case "getLeaveStateSnapshot":
        return jsonResponse_(getLeaveStateSnapshot_(payload));

      case "saveLeaveStateSnapshot":
        requireLeaveScopedWriteAccess_(payload, payload.coachCode, payload.studentCode);
        return jsonResponse_(saveLeaveStateSnapshot_(payload));

      default:
        return jsonResponse_({
          ok: false,
          message: "Unsupported action. Use bootstrap/ping/checkEvent/listEvents/createEvent/updateEvent/deleteSingleEvent/listLeaveRecords/saveLeaveRecord/listBillingProfiles/saveBillingProfile/listLessons/saveLessons/listMakeupRequests/saveMakeupRequest/listCoachBlocks/saveCoachBlock/getLeaveStateSnapshot/saveLeaveStateSnapshot/sendEmail."
        });
    }
  } catch (error) {
    return jsonResponse_({
      ok: false,
      message: error.message || String(error)
    });
  }
}

function shouldEnsureSheetsForAction_(action) {
  const normalized = String(action || "bootstrap").trim() || "bootstrap";
  const sheetActions = [
    "resolveCoachAccess",
    "resolveStudentAccess",
    "bootstrapAdmin",
    "bootstrapCoach",
    "bootstrapStudent",
    "bootstrap",
    "upsertCoach",
    "setCoachStatus",
    "deleteCoach",
    "upsertStudent",
    "assignStudentCoach",
    "setStudentStatus",
    "deleteStudent",
    "saveProgram",
    "publishProgram",
    "deleteProgram",
    "submitWorkoutLogs",
    "importWorkoutLogs",
    "touchCoach",
    "touchStudent",
    // 遷移後請假/扣費也走 Sheet，確保分頁存在再操作
    "listLeaveRecords",
    "saveLeaveRecord",
    "listBillingProfiles",
    "saveBillingProfile",
    // 完整課表雲端化：listLessons/saveLessons 走 Lessons 分頁，確保分頁存在再操作
    "listLessons",
    "saveLessons",
    "listMakeupRequests",
    "saveMakeupRequest",
    "listCoachBlocks",
    "saveCoachBlock"
  ];
  return sheetActions.indexOf(normalized) !== -1;
}

function buildFullBootstrap_() {
  return {
    ok: true,
    coaches: readTable_(SHEETS.coaches),
    students: readTable_(SHEETS.students),
    programs: readTable_(SHEETS.programs),
    programItems: readTable_(SHEETS.programItems),
    workoutLogs: readTable_(SHEETS.workoutLogs)
  };
}

function buildCoachBootstrap_(coachId) {
  const coaches = readTable_(SHEETS.coaches);
  const students = readTable_(SHEETS.students);
  const programs = readTable_(SHEETS.programs);
  const programItems = readTable_(SHEETS.programItems);
  const workoutLogs = readTable_(SHEETS.workoutLogs);
  const studentCoachMap = {};
  const programCoachMap = {};

  students.forEach(function(row) {
    studentCoachMap[String(row.student_id || "")] = String(row.primary_coach_id || "");
  });
  programs.forEach(function(row) {
    programCoachMap[String(row.program_id || "")] = String(row.coach_id || "");
  });

  const activeCoach = coaches.filter(function(row) {
    return row.coach_id === coachId;
  })[0] || null;

  return {
    ok: true,
    mode: "coach",
    coach: activeCoach,
    coaches: coaches,
    students: students.filter(function(row) {
      return !coachId || row.primary_coach_id === coachId;
    }),
    programs: programs.filter(function(row) {
      return !coachId || row.coach_id === coachId;
    }),
    programItems: programItems,
    workoutLogs: workoutLogs.filter(function(row) {
      if (!coachId) {
        return true;
      }
      const explicitCoachId = String(row.coach_id || "");
      if (explicitCoachId) {
        return explicitCoachId === coachId;
      }
      const byStudent = studentCoachMap[String(row.student_id || "")] || "";
      if (byStudent) {
        return byStudent === coachId;
      }
      const byProgram = programCoachMap[String(row.program_id || "")] || "";
      return byProgram === coachId;
    })
  };
}

function buildStudentBootstrap_(studentId) {
  const coaches = readTable_(SHEETS.coaches);
  const students = readTable_(SHEETS.students);
  const programs = readTable_(SHEETS.programs);
  const programItems = readTable_(SHEETS.programItems);
  const workoutLogs = readTable_(SHEETS.workoutLogs);

  const student = students.filter(function(row) {
    return row.student_id === studentId;
  })[0] || null;

  const coachPrograms = student
    ? programs.filter(function(row) {
        return row.coach_id === student.primary_coach_id && isProgramAvailableForStudent_(row, studentId);
      })
    : [];
  const coachProgramIds = {};
  coachPrograms.forEach(function(row) {
    coachProgramIds[String(row.program_id || "")] = true;
  });

  const coach = student
    ? coaches.filter(function(row) {
        return row.coach_id === student.primary_coach_id;
      })[0] || null
    : null;

  return {
    ok: true,
    mode: "student",
    student: student,
    coach: coach,
    coaches: coach ? [coach] : [],
    programs: coachPrograms,
    programItems: programItems.filter(function(row) {
      return coachProgramIds[String(row.program_id || "")];
    }),
    workoutLogs: workoutLogs.filter(function(row) {
      return row.student_id === studentId;
    })
  };
}

function resolveCoachAccessResponse_(accessValue) {
  const coach = resolveCoachByAccess_(accessValue);
  if (!coach) {
    return {
      ok: false,
      mode: "coach",
      message: "找不到這組教練代碼。"
    };
  }

  if (String(coach.status || "").toLowerCase() === "inactive") {
    return {
      ok: false,
      mode: "coach",
      message: "這位教練帳號目前已停用。"
    };
  }

  return buildCoachBootstrap_(coach.coach_id);
}

function resolveStudentAccessResponse_(accessValue) {
  const student = resolveStudentByAccess_(accessValue);
  if (!student) {
    return {
      ok: false,
      mode: "student",
      message: "找不到這組學生代碼。"
    };
  }

  if (String(student.status || "").toLowerCase() === "inactive") {
    return {
      ok: false,
      mode: "student",
      message: "這位學生帳號目前已停用。"
    };
  }

  return buildStudentBootstrap_(student.student_id);
}

function ensureSheets_() {
  const spreadsheet = getCoachflowSpreadsheet_();
  Object.keys(SCHEMA).forEach(function(sheetName) {
    ensureSheet_(spreadsheet, sheetName, SCHEMA[sheetName]);
  });
}

function getCoachflowSpreadsheet_() {
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (activeSpreadsheet) {
    return activeSpreadsheet;
  }

  const spreadsheetId = String(
    PropertiesService.getScriptProperties().getProperty("COACHFLOW_SPREADSHEET_ID") || ""
  ).trim();
  if (spreadsheetId) {
    return SpreadsheetApp.openById(spreadsheetId);
  }

  throw new Error("找不到 CoachFlow 試算表。請將試算表 ID 設到 Script Properties 的 COACHFLOW_SPREADSHEET_ID。");
}

function ensureSheet_(spreadsheet, sheetName, headers) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }

  const width = headers.length;
  const current = sheet.getRange(1, 1, 1, width).getValues()[0];
  const mismatch = headers.some(function(header, index) {
    return current[index] !== header;
  });

  if (mismatch) {
    sheet.getRange(1, 1, 1, width).setValues([headers]);
  }

  if (sheet.getFrozenRows() < 1) {
    sheet.setFrozenRows(1);
  }
}

function readTable_(sheetName) {
  const sheet = getCoachflowSpreadsheet_().getSheetByName(sheetName);
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    return [];
  }

  const headers = values[0];
  return values.slice(1)
    .filter(function(row) {
      return row.some(function(cell) {
        return cell !== "";
      });
    })
    .map(function(row) {
      const obj = {};
      headers.forEach(function(header, index) {
        obj[header] = row[index];
      });
      return obj;
    });
}

function upsertCoach_(coach) {
  const id = coach.coach_id || createId_("coach");
  const now = nowString_();
  const rows = readTable_(SHEETS.coaches);
  const existing = rows.filter(function(row) {
    return row.coach_id === id;
  })[0];

  const record = {
    coach_id: id,
    coach_name: coach.coach_name || coach.name || "",
    status: coach.status || existing && existing.status || "active",
    role: coach.role || existing && existing.role || "coach",
    access_code: coach.access_code || existing && existing.access_code || "",
    token: coach.token || existing && existing.token || "",
    last_used_at: coach.last_used_at || existing && existing.last_used_at || "",
    created_at: existing && existing.created_at || now,
    updated_at: now
  };

  if (!record.access_code || isFieldValueUsedByOtherRow_(rows, "access_code", record.access_code, "coach_id", id)) {
    record.access_code = buildUniqueCoachAccessCode_(record.coach_name, rows, id);
  }
  if (!record.token || isFieldValueUsedByOtherRow_(rows, "token", record.token, "coach_id", id)) {
    record.token = buildUniqueToken_(record.coach_name, "coach", rows, "coach_id", id);
  }

  upsertRow_(SHEETS.coaches, "coach_id", record);
}

function setCoachStatus_(coachId, status) {
  updateRow_(SHEETS.coaches, "coach_id", coachId, {
    status: status,
    updated_at: nowString_()
  });
}

function deleteCoach_(coachId) {
  deleteRow_(SHEETS.coaches, "coach_id", coachId);

  const students = readTable_(SHEETS.students).filter(function(row) {
    return row.primary_coach_id === coachId;
  });

  students.forEach(function(student) {
    updateRow_(SHEETS.students, "student_id", student.student_id, {
      primary_coach_id: "",
      primary_coach_name: "",
      updated_at: nowString_()
    });
  });
}

function upsertStudent_(student) {
  const id = student.student_id || createId_("student");
  const now = nowString_();
  const rows = readTable_(SHEETS.students);
  const existing = rows.filter(function(row) {
    return row.student_id === id;
  })[0];

  const coach = resolveCoach_(student.primary_coach_id || existing && existing.primary_coach_id || "");

  const record = {
    student_id: id,
    student_name: student.student_name || student.name || "",
    status: student.status || existing && existing.status || "active",
    access_code: student.access_code || existing && existing.access_code || "",
    token: student.token || existing && existing.token || "",
    primary_coach_id: coach ? coach.coach_id : "",
    primary_coach_name: coach ? coach.coach_name : "",
    last_used_at: student.last_used_at || existing && existing.last_used_at || "",
    created_at: existing && existing.created_at || now,
    updated_at: now
  };

  if (!record.access_code || isFieldValueUsedByOtherRow_(rows, "access_code", record.access_code, "student_id", id)) {
    record.access_code = buildUniqueStudentAccessCode_(record.student_name, rows, id);
  }
  if (!record.token || isFieldValueUsedByOtherRow_(rows, "token", record.token, "student_id", id)) {
    record.token = buildUniqueToken_(record.student_name, "student", rows, "student_id", id);
  }

  upsertRow_(SHEETS.students, "student_id", record);
}

function assignStudentCoach_(studentId, coachId) {
  const coach = resolveCoach_(coachId);
  updateRow_(SHEETS.students, "student_id", studentId, {
    primary_coach_id: coach ? coach.coach_id : "",
    primary_coach_name: coach ? coach.coach_name : "",
    updated_at: nowString_()
  });
}

function setStudentStatus_(studentId, status) {
  updateRow_(SHEETS.students, "student_id", studentId, {
    status: status,
    updated_at: nowString_()
  });
}

function deleteStudent_(studentId) {
  deleteRow_(SHEETS.students, "student_id", studentId);
  const logs = readTable_(SHEETS.workoutLogs).filter(function(row) {
    return row.student_id === studentId;
  });
  logs.forEach(function(row) {
    deleteRow_(SHEETS.workoutLogs, "log_id", row.log_id);
  });
}

function saveProgram_(program, items) {
  const id = program.program_id || program.id || createId_("program");
  const now = nowString_();
  const existing = readTable_(SHEETS.programs).filter(function(row) {
    return row.program_id === id;
  })[0];

  const coach = resolveCoach_(program.coach_id || program.coachId || existing && existing.coach_id || "");
  const hasTargetStudentIds = Object.prototype.hasOwnProperty.call(program, "target_student_ids")
    || Object.prototype.hasOwnProperty.call(program, "targetStudentIds");
  const targetStudentIds = hasTargetStudentIds
    ? normalizeProgramStudentIds_(program.target_student_ids || program.targetStudentIds || "")
    : normalizeProgramStudentIds_(existing && existing.target_student_ids || "");

  const record = {
    program_id: id,
    program_code: program.program_code || program.code || "",
    program_date: program.program_date || program.date || "",
    title: program.title || "",
    coach_id: coach ? coach.coach_id : "",
    coach_name: coach ? coach.coach_name : (program.coach_name || program.coachName || ""),
    notes: program.notes || "",
    published: String(Boolean(program.published || existing && truthy_(existing.published))),
    created_at: existing && existing.created_at || now,
    updated_at: now,
    target_student_ids: targetStudentIds.join(",")
  };

  upsertRow_(SHEETS.programs, "program_id", record);
  replaceProgramItems_(id, items || []);
}

function publishProgram_(programId, coachId) {
  const coachPrograms = readTable_(SHEETS.programs).filter(function(row) {
    return row.coach_id === coachId;
  });
  const targetProgram = coachPrograms.filter(function(row) {
    return row.program_id === programId;
  })[0] || null;
  const targetStudentIds = normalizeProgramStudentIds_(targetProgram && targetProgram.target_student_ids || "");

  coachPrograms.forEach(function(row) {
    const rowStudentIds = normalizeProgramStudentIds_(row.target_student_ids || "");
    const shouldPublishTarget = row.program_id === programId;
    const shouldUnpublish = targetStudentIds.length
      ? rowStudentIds.some(function(studentId) {
          return targetStudentIds.indexOf(studentId) !== -1;
        })
      : rowStudentIds.length === 0;
    if (!shouldPublishTarget && !shouldUnpublish) {
      return;
    }
    updateRow_(SHEETS.programs, "program_id", row.program_id, {
      published: String(shouldPublishTarget),
      updated_at: nowString_()
    });
  });
}

function deleteProgram_(programId) {
  deleteRow_(SHEETS.programs, "program_id", programId);
  deleteRowsByValue_(SHEETS.programItems, "program_id", programId);
}

function normalizeProgramStudentIds_(value) {
  const rawIds = Array.isArray(value)
    ? value
    : String(value || "").split(/[,，\n]/);
  const seen = {};
  return rawIds
    .map(function(id) {
      return String(id || "").trim();
    })
    .filter(function(id) {
      if (!id || seen[id]) {
        return false;
      }
      seen[id] = true;
      return true;
    });
}

function isProgramAvailableForStudent_(program, studentId) {
  const targetIds = normalizeProgramStudentIds_(program && program.target_student_ids || "");
  return !targetIds.length || targetIds.indexOf(studentId) !== -1;
}

function replaceProgramItems_(programId, items) {
  deleteRowsByValue_(SHEETS.programItems, "program_id", programId);

  if (!items.length) {
    return;
  }

  const rows = items.map(function(item, index) {
    return {
      item_id: item.item_id || item.id || createId_("item"),
      program_id: programId,
      sort_order: Number(item.sort_order || item.sortOrder || index + 1),
      category: item.category || "",
      exercise: item.exercise || "",
      target_sets: Number(item.target_sets || item.targetSets || 0),
      target_type: item.target_type || item.targetType || "",
      target_value: Number(item.target_value || item.targetValue || 0),
      item_note: item.item_note || item.itemNote || "",
      updated_at: nowString_()
    };
  });

  appendRows_(SHEETS.programItems, rows);
}

function submitWorkoutLogs_(logs) {
  // 用 ScriptLock 避免並行請求 race：
  // client 端 callCloudApiCritical 失敗會 retry，三次請求可能同時到達 Apps Script。
  // 沒鎖的話三個 execution 各跑 readTable_、都看不到對方 append 的 row、各自 append
  // → 同一筆 workout log 在 Sheet 變成 N 份重複。
  //
  // 同時：把整個 batch 改用「一次讀 + in-memory match map + 批次寫」邏輯，比
  // 原本 logs.forEach(replaceMatchingWorkoutLog_) 每次都 readTable_ 快很多，
  // 也避開「同一 batch 內前面 row 寫進去後 readTable_ 可能因 Apps Script 未
  // flush 看不到」這個 read-after-write 一致性風險。
  return withScriptLock_(function () {
    return submitWorkoutLogsLocked_(logs);
  });
}

function submitWorkoutLogsLocked_(logs) {
  const now = nowString_();
  const students = readTable_(SHEETS.students);
  const programs = readTable_(SHEETS.programs);
  const coaches = readTable_(SHEETS.coaches);
  const existingRows = readTable_(SHEETS.workoutLogs);

  const studentById = {};
  const programById = {};
  const coachById = {};
  students.forEach(function(row) { studentById[String(row.student_id || "")] = row; });
  programs.forEach(function(row) { programById[String(row.program_id || "")] = row; });
  coaches.forEach(function(row) { coachById[String(row.coach_id || "")] = row; });

  // Match key 改用 log_id（client 端 generate 的 UUID）：
  // - retry 重送同筆 → 同一個 log_id → upsert，不重複
  // - 不同時段送同動作 → 不同 log_id → 各自一筆，不誤合
  // - client 沒有 edit-resubmit 流程，每次送出都是新 log_id（user 確認過）
  //
  // 舊版 match key 用 student+program+activityDate+exercise，會把「同一天
  // 同一份 program 的同動作」整批硬合成一筆——dedupeWorkoutLogs 用同樣
  // key 跑下去殺掉一堆其實合法的「同動作不同時段」紀錄，這次一併修。
  const existingByLogId = {};
  existingRows.forEach(function(row) {
    const id = String(row.log_id || "");
    if (id) {
      existingByLogId[id] = row;
    }
  });

  const toAppend = [];
  const pendingUpdates = [];
  const batchSeenLogIds = {};
  logs.forEach(function(log) {
    const studentId = String(log.student_id || log.studentId || "");
    const programId = String(log.program_id || log.programId || "");
    const sourceStudent = studentById[studentId] || null;
    const sourceProgram = programById[programId] || null;
    const effectiveCoachId = String(
      log.coach_id
      || log.coachId
      || sourceProgram && sourceProgram.coach_id
      || sourceStudent && sourceStudent.primary_coach_id
      || ""
    );
    const sourceCoach = coachById[effectiveCoachId] || null;

    const record = {
      log_id: log.log_id || log.id || createId_("log"),
      program_id: programId,
      program_code: log.program_code || log.programCode || "",
      program_date: log.program_date || log.programDate || "",
      coach_id: effectiveCoachId,
      coach_name: log.coach_name || log.coachName || sourceCoach && sourceCoach.coach_name || "",
      student_id: studentId,
      student_name: log.student_name || log.studentName || "",
      category: log.category || "",
      exercise: log.exercise || "",
      target_sets: Number(log.target_sets || log.targetSets || 0),
      target_type: log.target_type || log.targetType || "",
      target_value: Number(log.target_value || log.targetValue || 0),
      actual_weight: blankIfNull_(log.actual_weight || log.actualWeight),
      actual_sets: blankIfNull_(log.actual_sets || log.actualSets),
      actual_reps: blankIfNull_(log.actual_reps || log.actualReps),
      student_note: log.student_note || log.studentNote || "",
      submitted_at: log.submitted_at || log.submittedAt || now,
      updated_at: now
    };

    if (existingByLogId[record.log_id] || batchSeenLogIds[record.log_id]) {
      // 同 log_id 已存在於 Sheet 或同批次中 → upsert（retry 自動收斂）
      pendingUpdates.push(record);
    } else {
      toAppend.push(record);
    }
    batchSeenLogIds[record.log_id] = true;
  });

  pendingUpdates.forEach(function(record) {
    updateRow_(SHEETS.workoutLogs, "log_id", record.log_id, record);
  });
  if (toAppend.length) {
    appendRows_(SHEETS.workoutLogs, toAppend);
  }
  SpreadsheetApp.flush();
}

function importWorkoutLogs_(logs) {
  const now = nowString_();
  const rows = logs.map(function(log) {
    return {
      log_id: log.log_id || log.id || createId_("log"),
      program_id: log.program_id || log.programId || "",
      program_code: log.program_code || log.programCode || "",
      program_date: log.program_date || log.programDate || log.date || "",
      coach_id: log.coach_id || log.coachId || "",
      coach_name: log.coach_name || log.coachName || "",
      student_id: log.student_id || log.studentId || "",
      student_name: log.student_name || log.studentName || "",
      category: log.category || "",
      exercise: log.exercise || "",
      target_sets: blankIfNull_(log.target_sets || log.targetSets),
      target_type: log.target_type || log.targetType || "",
      target_value: blankIfNull_(log.target_value || log.targetValue),
      actual_weight: blankIfNull_(log.actual_weight || log.actualWeight),
      actual_sets: blankIfNull_(log.actual_sets || log.actualSets),
      actual_reps: blankIfNull_(log.actual_reps || log.actualReps),
      student_note: log.student_note || log.studentNote || "",
      submitted_at: log.submitted_at || log.submittedAt || now,
      updated_at: now
    };
  });

  rows.forEach(replaceMatchingWorkoutLog_);
}

function replaceMatchingWorkoutLog_(record) {
  const existingRows = readTable_(SHEETS.workoutLogs);
  const matched = existingRows.filter(function(row) {
    return row.student_id === record.student_id
      && row.program_id === record.program_id
      && getWorkoutLogActivityDate_(row) === getWorkoutLogActivityDate_(record)
      && row.exercise === record.exercise;
  })[0];

  if (matched) {
    record.log_id = matched.log_id;
    updateRow_(SHEETS.workoutLogs, "log_id", matched.log_id, record);
    return;
  }

  appendRows_(SHEETS.workoutLogs, [record]);
}

function touchCoach_(coachId) {
  updateRow_(SHEETS.coaches, "coach_id", coachId, {
    last_used_at: nowString_(),
    updated_at: nowString_()
  });
}

function touchStudent_(studentId) {
  updateRow_(SHEETS.students, "student_id", studentId, {
    last_used_at: nowString_(),
    updated_at: nowString_()
  });
}

function resolveCoach_(coachId) {
  return readTable_(SHEETS.coaches).filter(function(row) {
    return row.coach_id === coachId;
  })[0] || null;
}

function resolveCoachByAccess_(accessValue) {
  const value = String(accessValue || "").trim().toLowerCase();
  if (!value) {
    return null;
  }

  return readTable_(SHEETS.coaches).filter(function(row) {
    const code = String(row.access_code || "").trim().toLowerCase();
    const token = String(row.token || "").trim().toLowerCase();
    return value === code || value === token || value === String(row.coach_id || "").trim().toLowerCase();
  })[0] || null;
}

function resolveStudentByAccess_(accessValue) {
  const value = String(accessValue || "").trim().toLowerCase();
  if (!value) {
    return null;
  }

  return readTable_(SHEETS.students).filter(function(row) {
    const code = String(row.access_code || "").trim().toLowerCase();
    const token = String(row.token || "").trim().toLowerCase();
    return value === code || value === token || value === String(row.student_id || "").trim().toLowerCase();
  })[0] || null;
}

function getAuthValue_(payload, keys) {
  payload = payload || {};
  const auth = payload.auth && typeof payload.auth === "object" ? payload.auth : {};
  for (var i = 0; i < keys.length; i += 1) {
    const key = keys[i];
    const value = payload[key] || auth[key];
    if (String(value || "").trim()) {
      return String(value || "").trim();
    }
  }
  return "";
}

function isActiveRecord_(record) {
  return !!record && String(record.status || "").toLowerCase() !== "inactive";
}

function getActiveCoachFromPayload_(payload) {
  const access = getAuthValue_(payload, ["coachAccess", "actorCoachAccess", "coachCode", "access"]);
  const coach = resolveCoachByAccess_(access);
  return isActiveRecord_(coach) ? coach : null;
}

function getActiveStudentFromPayload_(payload) {
  const access = getAuthValue_(payload, ["studentAccess", "actorStudentAccess", "studentCode", "access"]);
  const student = resolveStudentByAccess_(access);
  return isActiveRecord_(student) ? student : null;
}

function getAdminAccessValue_(payload) {
  return getAuthValue_(payload, ["adminAccess", "adminKey", "writeAdminAccess"]);
}

function isAdminAccessAllowed_(payload) {
  const expected = safeString_(
    PropertiesService.getScriptProperties().getProperty("COACHFLOW_ADMIN_ACCESS_CODE"),
    ""
  );
  if (!expected) {
    return false;
  }
  return String(getAdminAccessValue_(payload) || "").trim() === expected;
}

function requireAdminAccess_(payload) {
  const expected = safeString_(
    PropertiesService.getScriptProperties().getProperty("COACHFLOW_ADMIN_ACCESS_CODE"),
    ""
  );
  if (!expected) {
    throw new Error("UNAUTHORIZED: admin access is disabled. Set Script Property COACHFLOW_ADMIN_ACCESS_CODE before using admin bootstrap or admin-only writes.");
  }
  if (String(getAdminAccessValue_(payload) || "").trim() !== expected) {
    throw new Error("UNAUTHORIZED: missing or invalid admin access code.");
  }
}

function requireCoachWriteAccess_(payload, expectedCoachId) {
  if (isAdminAccessAllowed_(payload)) {
    return { admin: true };
  }
  const coach = getActiveCoachFromPayload_(payload);
  if (!coach) {
    throw new Error("UNAUTHORIZED: a valid coach access code is required for this write.");
  }
  if (expectedCoachId && String(coach.coach_id || "") !== String(expectedCoachId || "")) {
    throw new Error("UNAUTHORIZED: coach access does not match this record.");
  }
  return coach;
}

function requireCoachReadAccess_(payload, coachId) {
  return requireCoachWriteAccess_(payload, coachId);
}

function requireStudentReadAccess_(payload, studentId) {
  if (isAdminAccessAllowed_(payload)) {
    return { admin: true };
  }
  const student = getActiveStudentFromPayload_(payload);
  if (student && String(student.student_id || "") === String(studentId || "")) {
    return student;
  }

  const existing = readTable_(SHEETS.students).filter(function(row) {
    return String(row.student_id || "") === String(studentId || "");
  })[0];
  const coach = getActiveCoachFromPayload_(payload);
  if (coach && existing && String(existing.primary_coach_id || "") === String(coach.coach_id || "")) {
    return coach;
  }

  throw new Error("UNAUTHORIZED: a valid student or coach access code is required for this record.");
}

function requireCoachUpsertAccess_(payload, coach) {
  coach = coach || {};
  const id = coach.coach_id || coach.id || "";
  const rows = readTable_(SHEETS.coaches);
  const existing = rows.filter(function(row) {
    return String(row.coach_id || "") === String(id || "");
  })[0];

  if (!rows.length) {
    return true;
  }
  if (isAdminAccessAllowed_(payload)) {
    return true;
  }
  const actor = requireCoachWriteAccess_(payload);
  if (existing && String(existing.coach_id || "") !== String(actor.coach_id || "")) {
    throw new Error("UNAUTHORIZED: coach access can only update its own coach record.");
  }
  return true;
}

function requireStudentUpsertAccess_(payload, student) {
  student = student || {};
  if (isAdminAccessAllowed_(payload)) {
    return true;
  }

  const id = student.student_id || student.id || "";
  const existing = readTable_(SHEETS.students).filter(function(row) {
    return String(row.student_id || "") === String(id || "");
  })[0];
  const targetCoachId = student.primary_coach_id || existing && existing.primary_coach_id || "";
  const actor = requireCoachWriteAccess_(payload, targetCoachId || "");
  if (!targetCoachId) {
    student.primary_coach_id = actor.coach_id;
  }
  return true;
}

function requireStudentMutationAccess_(payload, studentId) {
  if (isAdminAccessAllowed_(payload)) {
    return true;
  }
  const existing = readTable_(SHEETS.students).filter(function(row) {
    return String(row.student_id || "") === String(studentId || "");
  })[0];
  const targetCoachId = existing && existing.primary_coach_id || "";
  const actor = requireCoachWriteAccess_(payload, targetCoachId || "");
  if (!targetCoachId) {
    return actor;
  }
  return actor;
}

function requireProgramSaveAccess_(payload, program) {
  program = program || {};
  if (isAdminAccessAllowed_(payload)) {
    return true;
  }

  const id = program.program_id || program.id || "";
  const existing = readTable_(SHEETS.programs).filter(function(row) {
    return String(row.program_id || "") === String(id || "");
  })[0];
  const targetCoachId = program.coach_id || program.coachId || existing && existing.coach_id || "";
  const actor = requireCoachWriteAccess_(payload, targetCoachId || "");
  if (!targetCoachId) {
    program.coach_id = actor.coach_id;
  }
  return true;
}

function requireProgramMutationAccess_(payload, programId, coachId) {
  if (isAdminAccessAllowed_(payload)) {
    return true;
  }
  const existing = readTable_(SHEETS.programs).filter(function(row) {
    return String(row.program_id || "") === String(programId || "");
  })[0];
  const targetCoachId = coachId || existing && existing.coach_id || "";
  return requireCoachWriteAccess_(payload, targetCoachId || "");
}

function requireWorkoutLogSubmitAccess_(payload, logs) {
  logs = logs || [];
  if (!logs.length || isAdminAccessAllowed_(payload)) {
    return true;
  }

  const student = getActiveStudentFromPayload_(payload);
  if (student) {
    const allForStudent = logs.every(function(log) {
      return String(log.student_id || log.studentId || "") === String(student.student_id || "");
    });
    if (allForStudent) {
      return true;
    }
  }

  const coach = getActiveCoachFromPayload_(payload);
  if (coach) {
    const students = readTable_(SHEETS.students);
    const programs = readTable_(SHEETS.programs);
    const allForCoach = logs.every(function(log) {
      const logCoachId = String(log.coach_id || log.coachId || "");
      if (logCoachId) {
        return logCoachId === String(coach.coach_id || "");
      }
      const studentRow = students.filter(function(row) {
        return String(row.student_id || "") === String(log.student_id || log.studentId || "");
      })[0];
      if (studentRow && String(studentRow.primary_coach_id || "") === String(coach.coach_id || "")) {
        return true;
      }
      const programRow = programs.filter(function(row) {
        return String(row.program_id || "") === String(log.program_id || log.programId || "");
      })[0];
      return programRow && String(programRow.coach_id || "") === String(coach.coach_id || "");
    });
    if (allForCoach) {
      return true;
    }
  }

  throw new Error("UNAUTHORIZED: a valid student or coach access code is required to submit logs.");
}

function codeMatchesRecord_(code, record, idField, accessField) {
  const normalized = normalizeCode_(code);
  if (!normalized || !record) {
    return false;
  }
  const candidates = [
    record[idField],
    record[accessField],
    record.token,
    record.student_id,
    record.coach_id
  ].map(normalizeCode_);
  return candidates.indexOf(normalized) !== -1;
}

function getActorCoachAccessCode_(payload) {
  return normalizeCode_(getAuthValue_(payload, ["coachAccess", "actorCoachAccess", "coachCode", "access"]));
}

function getActorStudentAccessCode_(payload) {
  return normalizeCode_(getAuthValue_(payload, ["studentAccess", "actorStudentAccess", "studentCode", "access"]));
}

function activeCoachCanWriteLeaveScope_(payload, coachCode) {
  const actorCode = getActorCoachAccessCode_(payload);
  const coach = getActiveCoachFromPayload_(payload);
  if (!coach) {
    return false;
  }
  return !coachCode ||
    normalizeCode_(coachCode) === actorCode ||
    codeMatchesRecord_(coachCode, coach, "coach_id", "access_code");
}

function activeStudentCanWriteLeaveScope_(payload, studentCode) {
  const actorCode = getActorStudentAccessCode_(payload);
  const student = getActiveStudentFromPayload_(payload);
  if (!student) {
    return false;
  }
  return !studentCode ||
    normalizeCode_(studentCode) === actorCode ||
    codeMatchesRecord_(studentCode, student, "student_id", "access_code");
}

function requireLeaveScopedWriteAccess_(payload, coachCode, studentCode) {
  if (isAdminAccessAllowed_(payload)) {
    return true;
  }
  if (coachCode && activeCoachCanWriteLeaveScope_(payload, coachCode)) {
    return true;
  }
  if (studentCode && activeStudentCanWriteLeaveScope_(payload, studentCode)) {
    return true;
  }
  if (!coachCode && !studentCode && getActiveCoachFromPayload_(payload)) {
    return true;
  }
  throw new Error("UNAUTHORIZED: a valid coach or student code is required for this leave-system write.");
}

function requireLeaveRecordWriteAccess_(payload) {
  const record = payload && (payload.record || payload) || {};
  return requireLeaveScopedWriteAccess_(payload, record.coachCode, record.studentCode);
}

function requireBillingProfileWriteAccess_(payload) {
  const profile = payload && (payload.profile || payload) || {};
  return requireLeaveScopedWriteAccess_(payload, profile.coachCode, profile.studentCode);
}

function requireLessonsWriteAccess_(payload) {
  const lessons = payload && payload.lessons || [];
  if (!lessons.length || isAdminAccessAllowed_(payload)) {
    return true;
  }

  const coach = getActiveCoachFromPayload_(payload);
  if (coach && lessons.every(function(lesson) {
    return activeCoachCanWriteLeaveScope_(payload, lesson.coachCode);
  })) {
    return true;
  }

  const student = getActiveStudentFromPayload_(payload);
  if (student && lessons.every(function(lesson) {
    return activeStudentCanWriteLeaveScope_(payload, lesson.studentCode);
  })) {
    return true;
  }

  throw new Error("UNAUTHORIZED: lesson writes must match the active coach or student code.");
}

function upsertRow_(sheetName, keyField, record) {
  const sheet = getCoachflowSpreadsheet_().getSheetByName(sheetName);
  const headers = SCHEMA[sheetName];
  const values = sheet.getDataRange().getValues();
  const keyIndex = headers.indexOf(keyField);

  for (var rowIndex = 1; rowIndex < values.length; rowIndex += 1) {
    if (values[rowIndex][keyIndex] === record[keyField]) {
      sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([headers.map(function(header) {
        return toCellValue_(record[header]);
      })]);
      return;
    }
  }

  appendRows_(sheetName, [record]);
}

function updateRow_(sheetName, keyField, keyValue, patch) {
  const rows = readTable_(sheetName);
  const existing = rows.filter(function(row) {
    return row[keyField] === keyValue;
  })[0];

  if (!existing) {
    return;
  }

  const next = {};
  SCHEMA[sheetName].forEach(function(header) {
    next[header] = patch.hasOwnProperty(header) ? patch[header] : existing[header];
  });

  upsertRow_(sheetName, keyField, next);
}

function deleteRow_(sheetName, keyField, keyValue) {
  const sheet = getCoachflowSpreadsheet_().getSheetByName(sheetName);
  const headers = SCHEMA[sheetName];
  const values = sheet.getDataRange().getValues();
  const keyIndex = headers.indexOf(keyField);

  for (var rowIndex = values.length - 1; rowIndex >= 1; rowIndex -= 1) {
    if (values[rowIndex][keyIndex] === keyValue) {
      sheet.deleteRow(rowIndex + 1);
    }
  }
}

function deleteRowsByValue_(sheetName, keyField, keyValue) {
  deleteRow_(sheetName, keyField, keyValue);
}

function appendRows_(sheetName, rows) {
  if (!rows.length) {
    return;
  }

  const sheet = getCoachflowSpreadsheet_().getSheetByName(sheetName);
  const headers = SCHEMA[sheetName];
  const values = rows.map(function(record) {
    return headers.map(function(header) {
      return toCellValue_(record[header]);
    });
  });

  sheet.getRange(sheet.getLastRow() + 1, 1, values.length, headers.length).setValues(values);
}

function toCellValue_(value) {
  if (value === null || typeof value === "undefined") {
    return "";
  }
  return value;
}

function blankIfNull_(value) {
  return (value === null || typeof value === "undefined") ? "" : value;
}

function buildCoachAccessCode_(name, index) {
  const letters = String(name || "CL")
    .replace(/[^A-Za-z]/g, "")
    .toUpperCase()
    .slice(0, 2) || "CL";
  return letters + Utilities.formatString("%03d", index);
}

function buildStudentAccessCode_(name, index) {
  const letters = String(name || "ST")
    .replace(/[^A-Za-z]/g, "")
    .toUpperCase()
    .slice(0, 2) || "ST";
  return letters + Utilities.formatString("%03d", index);
}

function buildUniqueCoachAccessCode_(name, rows, coachId) {
  return buildUniqueAccessCode_(name, rows, "coach_id", coachId, "CL");
}

function buildUniqueStudentAccessCode_(name, rows, studentId) {
  return buildUniqueAccessCode_(name, rows, "student_id", studentId, "ST");
}

function buildUniqueAccessCode_(name, rows, idField, recordId, fallbackPrefix) {
  const prefix = String(name || fallbackPrefix)
    .replace(/[^A-Za-z]/g, "")
    .toUpperCase()
    .slice(0, 2) || fallbackPrefix;
  const used = {};
  let maxIndex = 0;

  rows.forEach(function(row) {
    if (String(row[idField] || "") === String(recordId || "")) {
      return;
    }
    const code = String(row.access_code || "").trim().toUpperCase();
    if (!code) {
      return;
    }
    used[code] = true;
    const match = code.match(/^([A-Z]+)(\d+)$/);
    if (match && match[1] === prefix) {
      maxIndex = Math.max(maxIndex, Number(match[2]) || 0);
    }
  });

  let nextIndex = maxIndex + 1;
  let candidate = prefix + Utilities.formatString("%03d", nextIndex);
  while (used[candidate]) {
    nextIndex += 1;
    candidate = prefix + Utilities.formatString("%03d", nextIndex);
  }
  return candidate;
}

function buildToken_(name, fallbackPrefix) {
  const token = String(name || fallbackPrefix)
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[^a-z0-9]/g, "");
  return token || (fallbackPrefix + "-" + new Date().getTime());
}

function buildUniqueToken_(name, fallbackPrefix, rows, idField, recordId) {
  const used = {};
  rows.forEach(function(row) {
    if (String(row[idField] || "") === String(recordId || "")) {
      return;
    }
    const token = String(row.token || "").trim().toLowerCase();
    if (token) {
      used[token] = true;
    }
  });

  const baseToken = buildToken_(name, fallbackPrefix);
  if (baseToken && !used[String(baseToken).toLowerCase()]) {
    return baseToken;
  }

  let nextIndex = 1;
  let candidate = fallbackPrefix + Utilities.formatString("%03d", nextIndex);
  while (used[String(candidate).toLowerCase()]) {
    nextIndex += 1;
    candidate = fallbackPrefix + Utilities.formatString("%03d", nextIndex);
  }
  return candidate;
}

function isFieldValueUsedByOtherRow_(rows, fieldName, value, idField, recordId) {
  const normalizedValue = String(value || "").trim().toLowerCase();
  if (!normalizedValue) {
    return false;
  }

  return rows.some(function(row) {
    return String(row[idField] || "") !== String(recordId || "")
      && String(row[fieldName] || "").trim().toLowerCase() === normalizedValue;
  });
}

function createId_(prefix) {
  return prefix + "-" + Utilities.getUuid().slice(0, 8);
}

function nowString_() {
  return Utilities.formatDate(new Date(), APP_TIME_ZONE, "yyyy-MM-dd HH:mm:ss");
}

function getWorkoutLogActivityDate_(row) {
  const submittedAt = String(row.submitted_at || row.submittedAt || "").trim();
  const submittedMatch = submittedAt.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (submittedMatch) {
    return [
      submittedMatch[1],
      submittedMatch[2].padStart(2, "0"),
      submittedMatch[3].padStart(2, "0")
    ].join("-");
  }

  const programDate = String(row.program_date || row.programDate || "").trim();
  const programMatch = programDate.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (programMatch) {
    return [
      programMatch[1],
      programMatch[2].padStart(2, "0"),
      programMatch[3].padStart(2, "0")
    ].join("-");
  }

  return programDate;
}

function truthy_(value) {
  return String(value).toLowerCase() === "true";
}

function getAction_(e) {
  return (e && e.parameter && e.parameter.action) || "bootstrap";
}

function getParam_(e, key) {
  return (e && e.parameter && e.parameter[key]) || "";
}

function buildPingResponse_(action) {
  return {
    ok: true,
    action: action || "ping",
    message: "ok",
    timeZone: APP_TIME_ZONE,
    now: nowIso_()
  };
}

function checkCalendarEvent_(payload) {
  const calendar = resolveCalendar_(payload || {});
  if (!calendar) {
    return {
      ok: false,
      action: "checkEvent",
      message: "找不到可用日曆，請檢查 calendarId 或授權。"
    };
  }

  const eventId = safeString_(payload.eventId || payload.calendarEventId, "");
  if (!eventId) {
    return {
      ok: true,
      action: "checkEvent",
      exists: false,
      eventId: "",
      calendarId: calendar.getId(),
      message: "缺少 eventId。"
    };
  }

  const event = findCalendarEventById_(calendar, eventId);
  return {
    ok: true,
    action: "checkEvent",
    exists: Boolean(event),
    eventId: eventId,
    calendarId: calendar.getId(),
    message: event ? "事件存在。" : "找不到指定事件。"
  };
}

function listCalendarEvents_(payload) {
  const calendar = resolveCalendar_(payload || {});
  if (!calendar) {
    return {
      ok: false,
      action: "listEvents",
      message: "找不到可用日曆，請檢查 calendarId 或授權。"
    };
  }

  const startAt = new Date(payload.startAt || "");
  const endAt = new Date(payload.endAt || "");
  if (String(startAt) === "Invalid Date" || String(endAt) === "Invalid Date") {
    return { ok: false, action: "listEvents", message: "startAt / endAt 格式錯誤。" };
  }
  if (endAt <= startAt) {
    return { ok: false, action: "listEvents", message: "endAt 必須晚於 startAt。" };
  }

  const events = calendar.getEvents(startAt, endAt).map(function(event) {
    return {
      eventId: String(event.getId() || ""),
      title: String(event.getTitle() || ""),
      startAt: event.getStartTime().toISOString(),
      endAt: event.getEndTime().toISOString(),
      description: String(event.getDescription() || ""),
      location: String(event.getLocation() || "")
    };
  });

  return {
    ok: true,
    action: "listEvents",
    calendarId: calendar.getId(),
    startAt: startAt.toISOString(),
    endAt: endAt.toISOString(),
    count: events.length,
    events: events
  };
}

function createCalendarEvent_(payload) {
  if (isNormalLeaveCalendarCreateBlocked_(payload || {})) {
    return {
      ok: false,
      action: "createEvent",
      blocked: true,
      message: "正常請假只能刪除原課程事件，不可新增請假事件。"
    };
  }

  const calendar = resolveCalendar_(payload || {});
  if (!calendar) {
    return { ok: false, message: "找不到可用日曆，請檢查 calendarId 或授權。" };
  }

  const startAt = new Date(payload.startAt || "");
  const endAt = new Date(payload.endAt || "");
  if (String(startAt) === "Invalid Date" || String(endAt) === "Invalid Date" || endAt <= startAt) {
    return { ok: false, message: "startAt / endAt 格式錯誤。" };
  }

  const title = String(payload.title || "CoachFlow 課程");
  const lines = [
    "來源：CoachFlow 請假系統",
    "學生：" + String(payload.studentCode || payload.studentName || "-"),
    "教練：" + String(payload.coachCode || payload.coachName || "-"),
    "類型：" + String(payload.sourceType || "-"),
    "原因：" + String(payload.reason || "-")
  ];
  const options = {
    description: lines.join("\n")
  };
  if (payload.location) {
    options.location = String(payload.location);
  }

  const event = calendar.createEvent(title, startAt, endAt, options);
  return {
    ok: true,
    eventId: event.getId(),
    calendarId: calendar.getId(),
    message: "createEvent 成功"
  };
}

function updateCalendarEvent_(payload) {
  const calendar = resolveCalendar_(payload || {});
  if (!calendar) {
    return { ok: false, action: "updateEvent", message: "找不到可用日曆，請檢查 calendarId 或授權。" };
  }

  const eventId = String(payload.eventId || payload.calendarEventId || "").trim();
  if (!eventId) {
    return { ok: false, action: "updateEvent", message: "缺少 eventId。" };
  }

  const event = findCalendarEventById_(calendar, eventId);
  if (!event) {
    return { ok: false, action: "updateEvent", message: "找不到指定事件。", eventId: eventId };
  }

  const beforeTitle = String(event.getTitle() || "");
  const title = String(payload.title || payload.summary || payload.name || beforeTitle).trim();
  if (title && title !== beforeTitle) {
    event.setTitle(title);
  }
  if (payload.description !== undefined) {
    event.setDescription(String(payload.description || ""));
  }
  if (payload.location !== undefined) {
    event.setLocation(String(payload.location || ""));
  }

  return {
    ok: true,
    action: "updateEvent",
    eventId: eventId,
    calendarId: calendar.getId(),
    beforeTitle: beforeTitle,
    title: String(event.getTitle() || ""),
    message: "updateEvent 成功"
  };
}

function deleteCalendarEvent_(payload) {
  const calendar = resolveCalendar_(payload || {});
  if (!calendar) {
    return { ok: false, message: "找不到可用日曆，請檢查 calendarId 或授權。" };
  }

  const eventId = String(payload.eventId || payload.calendarEventId || "").trim();
  if (!eventId) {
    return { ok: false, message: "缺少 eventId。" };
  }

  // 安全預設：除非呼叫端明確指定 deleteScope:"series"，否則一律走 single-occurrence 路徑。
  // 過去 deleteEvent 的非 single 分支會直接呼 event.deleteEvent()，若 eventId 是
  // recurring master，會把整個系列（包括未來所有實例）一起刪掉，無法復原。
  // 已知前端只有「刪一堂」需求；series 刪除應該走額外明確的 deleteScope。
  const explicitSeriesScope = String((payload || {}).deleteScope || "").toLowerCase() === "series";
  const singleDelete = explicitSeriesScope ? false : true;

  const event = singleDelete
    ? findSingleOccurrenceForDelete_(calendar, payload || {}, eventId)
    : findCalendarEventById_(calendar, eventId);
  if (!event) {
    if (singleDelete) {
      return {
        ok: false,
        action: "deleteEvent",
        message: "無法確認單一課程事件，已停止刪除以避免刪除整個週期課程。",
        eventId: eventId,
        blocked: true,
        singleEventOnly: true
      };
    }
    return { ok: false, message: "找不到指定事件，可能已刪除。", eventId: eventId };
  }

  // 保險：只在「非 single delete」路徑（會把整個 series 刪光）才檢查 recurring。
  // event.isRecurringEvent() 對單一實例也回 true，但對單一實例呼叫 deleteEvent()
  // 只刪那一筆 instance、不會動到整個系列，所以 single delete 不需要這道 guard。
  if (!singleDelete && event.isRecurringEvent && event.isRecurringEvent() && !explicitSeriesScope) {
    return {
      ok: false,
      action: "deleteEvent",
      message: "目標是重複事件系列；單堂刪除請用 deleteSingleEvent action，整個系列刪除請帶 deleteScope:\"series\"。",
      eventId: eventId,
      blocked: true,
      isRecurring: true
    };
  }

  event.deleteEvent();
  return {
    ok: true,
    eventId: String(event.getId() || eventId),
    requestedEventId: eventId,
    calendarId: calendar.getId(),
    singleEventOnly: singleDelete,
    seriesScope: explicitSeriesScope,
    message: "deleteEvent 成功"
  };
}

function sendEmailNotice_(payload) {
  const to = String(payload.to || "").trim();
  if (!to) {
    return { ok: false, message: "缺少收件者 to。" };
  }

  const template = String(payload.template || "coachflow_notice");
  const subject = String(payload.subject || ("[CoachFlow] " + template));
  const body = String(payload.body || buildEmailBody_(payload));
  MailApp.sendEmail({
    to: to,
    subject: subject,
    body: body,
    name: "CoachFlow"
  });
  return {
    ok: true,
    message: "sendEmail 成功",
    to: to
  };
}

function buildEmailBody_(payload) {
  const lines = [
    "這是 CoachFlow 系統通知。",
    "",
    "模板：" + String(payload.template || "coachflow_notice"),
    "學生：" + String(payload.studentCode || payload.studentName || "-"),
    "教練：" + String(payload.coachCode || payload.coachName || "-"),
    "時間：" + String(payload.when || payload.startAt || "-")
  ];
  return lines.join("\n");
}

function resolveCalendar_(payload) {
  const scriptProps = PropertiesService.getScriptProperties();
  const fallbackCalendarId = String(scriptProps.getProperty("DEFAULT_CALENDAR_ID") || "").trim();
  const calendarId = String(payload.calendarId || payload.coachCalendarId || fallbackCalendarId).trim();
  if (calendarId) {
    const byId = CalendarApp.getCalendarById(calendarId);
    if (byId) {
      return byId;
    }
  }
  return CalendarApp.getDefaultCalendar();
}

function findCalendarEventById_(calendar, eventId) {
  const raw = String(eventId || "").trim();
  if (!raw) {
    return null;
  }

  const candidates = [raw];
  if (raw.indexOf("@google.com") === -1) {
    candidates.push(raw + "@google.com");
  } else {
    candidates.push(raw.replace("@google.com", ""));
  }

  for (var index = 0; index < candidates.length; index += 1) {
    const found = calendar.getEventById(candidates[index]);
    if (found) {
      return found;
    }
  }
  return null;
}

function isSingleEventDeleteRequested_(payload) {
  const scope = String(payload.deleteScope || payload.deleteMode || "").trim().toLowerCase();
  return truthy_(payload.singleEventOnly) || truthy_(payload.strictSingleOccurrence) || scope === "single" || scope === "singleevent";
}

function isNormalLeaveCalendarCreateBlocked_(payload) {
  const reason = String(payload.reason || "").trim().toLowerCase();
  const attendanceStatus = String(payload.attendanceStatus || "").trim().toLowerCase();
  const leaveType = String(payload.leaveType || payload.type || "").trim().toLowerCase();
  return reason === "student_normal_leave" ||
    attendanceStatus === "leave-normal" ||
    leaveType === "normal_leave" ||
    leaveType === "student_normal_leave";
}

function normalizeCalendarEventId_(eventId) {
  const suffix = "@google.com";
  const raw = String(eventId || "").trim();
  return raw.slice(-suffix.length) === suffix
    ? raw.slice(0, raw.length - suffix.length)
    : raw;
}

function findSingleOccurrenceForDelete_(calendar, payload, rawEventId) {
  const occurrenceStart = new Date(payload.occurrenceStartAt || payload.lessonStartAt || payload.startAt || "");
  if (String(occurrenceStart) === "Invalid Date") {
    return null;
  }

  const dateKey = String(payload.dateKey || payload.occurrenceDate || Utilities.formatDate(occurrenceStart, APP_TIME_ZONE, "yyyy-MM-dd")).trim();
  const dayStart = new Date(dateKey + "T00:00:00+08:00");
  const dayEnd = new Date(dayStart.getTime() + 24 * 60 * 60 * 1000);
  const requestedId = normalizeCalendarEventId_(rawEventId);
  const studentCode = String(payload.studentCode || "").trim().toLowerCase();
  const studentName = String(payload.studentName || "").trim().toLowerCase();
  const events = calendar.getEvents(dayStart, dayEnd);
  let bestEvent = null;
  let bestScore = -1;
  let tied = false;

  // 若呼叫端提供了 requestedId（leave 紀錄裡的 calendarEventId），就嚴格只接受
  // ID 相符的 event。學生姓名 + 時間 fallback 在這個情境下會誤殺別的事件 ——
  // 已上線案例：ZO006 leave 的舊 lesson eventId 已不存在於 Google，但同一天
  // 11:00 有補課事件，sync 嘗試刪 leave 的 event 時，學生姓名匹配成功、時間
  // 也對得上 (因為 lesson.startAt 被 align 拉到補課時間)，就把補課事件刪掉了。
  const hasRequestedId = Boolean(requestedId);
  for (var i = 0; i < events.length; i += 1) {
    const event = events[i];
    const eventId = normalizeCalendarEventId_(event.getId());
    const startDiff = Math.abs(event.getStartTime().getTime() - occurrenceStart.getTime());
    const searchable = (String(event.getTitle() || "") + "\n" + String(event.getDescription() || "") + "\n" + String(event.getLocation() || "")).toLowerCase();
    const idMatched = requestedId && eventId && requestedId === eventId;
    const studentMatched = Boolean(
      (studentCode && searchable.indexOf(studentCode) !== -1) ||
      (studentName && searchable.indexOf(studentName) !== -1)
    );
    const timeMatched = startDiff <= 2 * 60 * 1000;
    const looseTimeMatched = startDiff <= 15 * 60 * 1000;
    const accepted = hasRequestedId
      ? (idMatched && looseTimeMatched)
      : (studentMatched && timeMatched);
    if (!accepted) {
      continue;
    }

    let score = 0;
    if (idMatched) {
      score += 10;
    }
    if (timeMatched) {
      score += 8;
    } else if (looseTimeMatched) {
      score += 4;
    }
    if (studentMatched) {
      score += 4;
    }
    if (score > bestScore) {
      bestEvent = event;
      bestScore = score;
      tied = false;
    } else if (score === bestScore) {
      tied = true;
    }
  }

  return tied ? null : bestEvent;
}

function listLeaveRecords_(payload) {
  const coachCode = normalizeCode_(payload.coachCode);
  const studentCode = normalizeCode_(payload.studentCode);
  const records = getLeaveRecords_()
    .filter(function(record) {
      const recordCoachCode = normalizeCode_(record.coachCode);
      const recordStudentCode = normalizeCode_(record.studentCode);
      return (!coachCode || recordCoachCode === coachCode) &&
        (!studentCode || recordStudentCode === studentCode);
    })
    .sort(function(a, b) {
      return new Date(b.submittedAt || 0) - new Date(a.submittedAt || 0);
    });
  return {
    ok: true,
    action: "listLeaveRecords",
    count: records.length,
    records: records
  };
}

function saveLeaveRecord_(payload) {
  const record = normalizeLeaveRecord_(payload.record || payload);
  if (!record.id && !record.lessonKey) {
    return {
      ok: false,
      action: "saveLeaveRecord",
      message: "缺少請假紀錄 id 或 lessonKey。"
    };
  }

  // 包進 ScriptLock：多裝置同時 saveLeaveRecord 時 read-modify-write 會 lost update。
  return withScriptLock_(function () {
    const records = getLeaveRecords_();
    const recordKey = getLeaveRecordKey_(record);
    const recordLessonKey = safeString_(record.lessonKey, "") ||
      lessonCloudKey_(record.studentCode, record.coachCode, record.lessonStartAt);
    if (!safeString_(record.revokedAt, "") && String(record.type || "normal") === "normal" && recordLessonKey) {
      const duplicateActive = records.filter(function(item) {
        const itemLessonKey = safeString_(item.lessonKey, "") ||
          lessonCloudKey_(item.studentCode, item.coachCode, item.lessonStartAt);
        return getLeaveRecordKey_(item) !== recordKey &&
          !safeString_(item.revokedAt, "") &&
          String(item.type || "normal") === "normal" &&
          itemLessonKey === recordLessonKey;
      })[0];
      if (duplicateActive) {
        return {
          ok: true,
          action: "saveLeaveRecord",
          record: duplicateActive,
          replaced: false,
          duplicate: true
        };
      }
    }
    let replaced = false;
    const nextRecords = records.map(function(item) {
      if (getLeaveRecordKey_(item) === recordKey) {
        replaced = true;
        const merged = Object.assign({}, item, record, {
          updatedAt: nowIso_()
        });
        // 防復活 ratchet：同一筆 leave id 一旦 revoked 就不可被「未帶 revokedAt
        // 的舊狀態」蓋回活著。沒有任何合法流程會清空同一筆 leave 的 revokedAt
        // （重試請假一律建新 id），會清空的只有「stale 裝置手動上傳整包舊資料」
        // —— 已上線案例：教練在 A 裝置取消請假，B 裝置（資料較舊）按「上傳到
        // 雲端」把整批 leave push 回雲，revokedAt 被空值覆蓋 → 請假復活 →
        // 所有裝置同步後課程又變回「請假」（幽靈請假）。
        if (safeString_(item.revokedAt, "") && !safeString_(record.revokedAt, "")) {
          merged.revokedAt = item.revokedAt;
          merged.revokedBy = item.revokedBy;
          merged.makeupEligible = item.makeupEligible;
        }
        return merged;
      }
      return item;
    });

    if (!replaced) {
      nextRecords.unshift(Object.assign({}, record, {
        createdAt: record.createdAt || nowIso_(),
        updatedAt: nowIso_()
      }));
    }

    const truncated = nextRecords.length > 500;
    setLeaveRecords_(nextRecords.slice(0, 500));
    return {
      ok: true,
      action: "saveLeaveRecord",
      record: record,
      replaced: replaced,
      truncated: truncated,
      droppedCount: truncated ? nextRecords.length - 500 : 0
    };
  });
}

function normalizeLeaveRecord_(record) {
  return {
    id: safeString_(record.id, ""),
    lessonId: safeString_(record.lessonId, ""),
    lessonKey: safeString_(record.lessonKey, ""),
    calendarEventId: safeString_(record.calendarEventId, ""),
    studentCode: normalizeCode_(record.studentCode),
    coachCode: normalizeCode_(record.coachCode),
    lessonStartAt: safeString_(record.lessonStartAt, ""),
    type: safeString_(record.type, "normal"),
    submittedAt: safeString_(record.submittedAt, nowIso_()),
    submittedBy: normalizeCode_(record.submittedBy),
    submittedByRole: safeString_(record.submittedByRole, ""),
    makeupEligible: record.makeupEligible === false || String(record.makeupEligible).toLowerCase() === "false" ? false : true,
    emailNoticeStatus: safeString_(record.emailNoticeStatus, ""),
    emailNoticeAt: safeString_(record.emailNoticeAt || record.emailNoticeSentAt || record.emailNoticeQueuedAt, ""),
    revokedAt: safeString_(record.revokedAt, ""),
    revokedBy: normalizeCode_(record.revokedBy),
    createdAt: safeString_(record.createdAt, ""),
    updatedAt: safeString_(record.updatedAt, "")
  };
}

function getLeaveRecordKey_(record) {
  return safeString_(record.id, "") || safeString_(record.lessonKey, "");
}

// 已從 PropertiesService.LEAVE_RECORDS_V1 遷移到 Sheet（一筆一列）。
// PropertiesService 整個 script 上限 500KB，請假紀錄會持續累積撞牆；Sheet 容量
// 實質沒有上限，也不會跟其他 properties 互相排擠。
function getLeaveRecords_() {
  const rows = readTable_(SHEETS.leaveRecords);
  return rows.map(normalizeLeaveRecord_);
}

// 整批覆寫：清掉資料區（保留 header）再 bulk write。
// saveLeaveRecord_ 仍是 read-modify-write 整包；寫入量低（請假頻率不高），Sheet
// 寫 500 列約一秒可接受，且語意最直接、不會留半成品。
function setLeaveRecords_(records) {
  const sheet = getCoachflowSpreadsheet_().getSheetByName(SHEETS.leaveRecords);
  const headers = SCHEMA.LeaveRecords;
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
  }
  if (records && records.length) {
    appendRows_(SHEETS.leaveRecords, records.map(normalizeLeaveRecord_));
  }
}

function listBillingProfiles_(payload) {
  const coachCode = normalizeCode_(payload.coachCode);
  const studentCode = normalizeCode_(payload.studentCode);
  const profiles = getBillingProfiles_()
    .filter(function(profile) {
      return (!coachCode || normalizeCode_(profile.coachCode) === coachCode) &&
        (!studentCode || normalizeCode_(profile.studentCode) === studentCode);
    })
    .sort(function(a, b) {
      return new Date(b.updatedAt || 0) - new Date(a.updatedAt || 0);
    });
  return {
    ok: true,
    action: "listBillingProfiles",
    count: profiles.length,
    profiles: profiles
  };
}

function saveBillingProfile_(payload) {
  const profile = normalizeBillingProfile_(payload.profile || payload);
  if (!profile.studentCode) {
    return {
      ok: false,
      action: "saveBillingProfile",
      message: "Missing studentCode."
    };
  }

  return withScriptLock_(function () {
    const now = nowIso_();
    const profiles = getBillingProfiles_();
    const profileKey = getBillingProfileKey_(profile);
    let replaced = false;
    let ignored = false;
    const existingIndex = profiles.findIndex(function(item) {
      return getBillingProfileKey_(item) === profileKey;
    });
    if (existingIndex >= 0) {
      const existing = profiles[existingIndex];
      const incomingTime = new Date(profile.updatedAt || 0).getTime();
      const existingTime = new Date(existing.updatedAt || 0).getTime();
      if (Number.isFinite(incomingTime) && Number.isFinite(existingTime) && incomingTime + 2000 < existingTime) {
        ignored = true;
        profile.createdAt = existing.createdAt || now;
        profile.updatedAt = existing.updatedAt || now;
        setBillingProfiles_(profiles);
        return {
          ok: true,
          action: "saveBillingProfile",
          profile: existing,
          replaced: false,
          ignored: ignored
        };
      }
      profile.createdAt = existing.createdAt || profile.createdAt || now;
      profile.updatedAt = profile.updatedAt || now;
      profiles[existingIndex] = profile;
      replaced = true;
    } else {
      profile.createdAt = profile.createdAt || now;
      profile.updatedAt = profile.updatedAt || now;
      profiles.push(profile);
    }
    setBillingProfiles_(profiles);
    return {
      ok: true,
      action: "saveBillingProfile",
      profile: profile,
      replaced: replaced,
      ignored: ignored
    };
  });
}

function normalizeBillingProfile_(profile) {
  profile = profile || {};
  return {
    studentCode: normalizeCode_(profile.studentCode || profile.code),
    coachCode: normalizeCode_(profile.coachCode),
    studentName: safeString_(profile.studentName || profile.name, ""),
    email: safeString_(profile.email, ""),
    emailUpdatedAt: safeString_(profile.emailUpdatedAt, ""),
    emailUpdatedBy: normalizeCode_(profile.emailUpdatedBy),
    chargeStartCount: toNonNegativeInt_(profile.chargeStartCount, 0),
    chargeStartCountUpdatedAt: safeString_(profile.chargeStartCountUpdatedAt, ""),
    chargeStartCountUpdatedBy: normalizeCode_(profile.chargeStartCountUpdatedBy),
    paidThroughCount: toNonNegativeInt_(profile.paidThroughCount, 0),
    paymentStatus: normalizePaymentStatus_(profile.paymentStatus),
    paymentNote: safeString_(profile.paymentNote, ""),
    paymentConfirmedAt: safeString_(profile.paymentConfirmedAt, ""),
    paymentConfirmedBy: normalizeCode_(profile.paymentConfirmedBy),
    chargeReminderLogs: normalizeBillingReminderLogs_(profile.chargeReminderLogs),
    chargeReminderStep: Math.max(1, toNonNegativeInt_(profile.chargeReminderStep, 4)),
    systemChargedCount: hasNumericValue_(profile.systemChargedCount) ? toNonNegativeInt_(profile.systemChargedCount, 0) : "",
    totalChargedCount: hasNumericValue_(profile.totalChargedCount) ? toNonNegativeInt_(profile.totalChargedCount, 0) : "",
    currentCycleChargedCount: hasNumericValue_(profile.currentCycleChargedCount) ? toNonNegativeInt_(profile.currentCycleChargedCount, 0) : "",
    remainingToNextPayment: hasNumericValue_(profile.remainingToNextPayment) ? toNonNegativeInt_(profile.remainingToNextPayment, 0) : "",
    nextPaymentDueCount: hasNumericValue_(profile.nextPaymentDueCount) ? toNonNegativeInt_(profile.nextPaymentDueCount, 0) : "",
    effectivePaymentStatus: normalizePaymentStatus_(profile.effectivePaymentStatus || profile.paymentStatus),
    updatedAt: safeString_(profile.updatedAt || profile.billingUpdatedAt, nowIso_()),
    updatedBy: normalizeCode_(profile.updatedBy || profile.billingUpdatedBy || "SYSTEM"),
    createdAt: safeString_(profile.createdAt, "")
  };
}

function getBillingProfileKey_(profile) {
  return normalizeCode_(profile.studentCode);
}

// 已從 PropertiesService.BILLING_PROFILES_V1 遷移到 Sheet。
// chargeReminderLogs 是物件陣列，cell 存不下 → 序列化進 chargeReminderLogsJson 欄位。
function getBillingProfiles_() {
  const rows = readTable_(SHEETS.billingProfiles);
  return rows.map(function(row) {
    var logs;
    try {
      const raw = String(row.chargeReminderLogsJson || "").trim();
      logs = raw ? JSON.parse(raw) : [];
      if (!Array.isArray(logs)) {
        logs = [];
      }
    } catch (error) {
      logs = [];
    }
    row.chargeReminderLogs = logs;
    return normalizeBillingProfile_(row);
  });
}

function setBillingProfiles_(profiles) {
  const sheet = getCoachflowSpreadsheet_().getSheetByName(SHEETS.billingProfiles);
  const headers = SCHEMA.BillingProfiles;
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
  }
  if (profiles && profiles.length) {
    const rowsToWrite = profiles.map(function(profile) {
      const normalized = normalizeBillingProfile_(profile);
      const row = Object.assign({}, normalized);
      row.chargeReminderLogsJson = JSON.stringify(normalized.chargeReminderLogs || []);
      delete row.chargeReminderLogs;
      return row;
    });
    appendRows_(SHEETS.billingProfiles, rowsToWrite);
  }
}


// ====================================================================
// 補課申請與教練停課：雲端唯一來源
// ====================================================================

function normalizeMakeupRequest_(request) {
  request = request || {};
  const allowed = {
    pending: true,
    approved: true,
    rejected: true,
    cancelled: true,
    expired: true
  };
  const status = safeString_(request.status, "pending").toLowerCase();
  return {
    id: safeString_(request.id, ""),
    code: safeString_(request.code, ""),
    leaveId: safeString_(request.leaveId, ""),
    lessonId: safeString_(request.lessonId, ""),
    studentCode: normalizeCode_(request.studentCode),
    coachCode: normalizeCode_(request.coachCode),
    status: allowed[status] ? status : "pending",
    pendingAt: safeString_(request.pendingAt, ""),
    resolvedAt: safeString_(request.resolvedAt, ""),
    startAt: safeString_(request.startAt, ""),
    rejectReason: safeString_(request.rejectReason, ""),
    emailNoticeStatus: safeString_(request.emailNoticeStatus, ""),
    emailNoticeAt: safeString_(request.emailNoticeAt || request.emailNoticeSentAt || request.emailNoticeQueuedAt, ""),
    cloudSyncStatus: safeString_(request.cloudSyncStatus, ""),
    cloudVerifiedAt: safeString_(request.cloudVerifiedAt, ""),
    createdAt: safeString_(request.createdAt, ""),
    updatedAt: safeString_(request.updatedAt, "")
  };
}

function getMakeupRequests_() {
  return readTable_(SHEETS.makeupRequests).map(normalizeMakeupRequest_);
}

function listMakeupRequests_(payload) {
  const coachCode = normalizeCode_(payload.coachCode);
  const studentCode = normalizeCode_(payload.studentCode);
  const requests = getMakeupRequests_().filter(function(request) {
    return (!coachCode || request.coachCode === coachCode) &&
      (!studentCode || request.studentCode === studentCode);
  }).sort(function(a, b) {
    return new Date(b.pendingAt || b.createdAt || 0) - new Date(a.pendingAt || a.createdAt || 0);
  });
  return {
    ok: true,
    action: "listMakeupRequests",
    count: requests.length,
    requests: requests
  };
}

function saveMakeupRequest_(payload) {
  const request = normalizeMakeupRequest_(payload.request || payload);
  if (!request.id || !request.studentCode || !request.coachCode) {
    return {
      ok: false,
      action: "saveMakeupRequest",
      message: "Missing makeup request id/studentCode/coachCode."
    };
  }
  return withScriptLock_(function() {
    const requests = getMakeupRequests_();
    const existing = requests.filter(function(item) { return item.id === request.id; })[0] || null;
    const finalStatuses = { approved: true, rejected: true, cancelled: true, expired: true };
    if (existing) {
      const incomingTime = new Date(request.updatedAt || request.resolvedAt || request.pendingAt || 0).getTime();
      const existingTime = new Date(existing.updatedAt || existing.resolvedAt || existing.pendingAt || 0).getTime();
      if ((finalStatuses[existing.status] && request.status === "pending") ||
        (Number.isFinite(incomingTime) && Number.isFinite(existingTime) && incomingTime + 2000 < existingTime)) {
        return {
          ok: true,
          action: "saveMakeupRequest",
          request: existing,
          ignored: true
        };
      }
    }
    const now = nowIso_();
    request.createdAt = existing && existing.createdAt || request.createdAt || now;
    request.updatedAt = now;
    upsertRow_(SHEETS.makeupRequests, "id", request);
    return {
      ok: true,
      action: "saveMakeupRequest",
      request: request,
      replaced: Boolean(existing),
      ignored: false
    };
  });
}

function normalizeCoachBlock_(block) {
  block = block || {};
  return {
    id: safeString_(block.id, ""),
    coachCode: normalizeCode_(block.coachCode),
    startAt: safeString_(block.startAt, ""),
    endAt: safeString_(block.endAt, ""),
    reason: safeString_(block.reason, ""),
    revokedAt: safeString_(block.revokedAt, ""),
    createdAt: safeString_(block.createdAt, ""),
    updatedAt: safeString_(block.updatedAt, "")
  };
}

function getCoachBlocks_() {
  return readTable_(SHEETS.coachBlocks).map(normalizeCoachBlock_);
}

function listCoachBlocks_(payload) {
  const coachCode = normalizeCode_(payload.coachCode);
  const blocks = getCoachBlocks_().filter(function(block) {
    return !coachCode || block.coachCode === coachCode;
  }).sort(function(a, b) {
    return new Date(b.startAt || b.createdAt || 0) - new Date(a.startAt || a.createdAt || 0);
  });
  return {
    ok: true,
    action: "listCoachBlocks",
    count: blocks.length,
    blocks: blocks
  };
}

function saveCoachBlock_(payload) {
  const block = normalizeCoachBlock_(payload.block || payload);
  if (!block.id || !block.coachCode) {
    return {
      ok: false,
      action: "saveCoachBlock",
      message: "Missing coach block id/coachCode."
    };
  }
  return withScriptLock_(function() {
    const blocks = getCoachBlocks_();
    const existing = blocks.filter(function(item) { return item.id === block.id; })[0] || null;
    if (existing && existing.revokedAt && !block.revokedAt) {
      return {
        ok: true,
        action: "saveCoachBlock",
        block: existing,
        ignored: true
      };
    }
    const now = nowIso_();
    block.createdAt = existing && existing.createdAt || block.createdAt || now;
    block.updatedAt = now;
    upsertRow_(SHEETS.coachBlocks, "id", block);
    return {
      ok: true,
      action: "saveCoachBlock",
      block: block,
      replaced: Boolean(existing),
      ignored: false
    };
  });
}
function cleanupLeaveTestData_(payload) {
  const coachCode = normalizeCode_(payload && payload.coachCode);
  if (!coachCode) {
    return {
      ok: false,
      action: "cleanupLeaveTestData",
      message: "Missing coachCode."
    };
  }

  return withScriptLock_(function () {
    const profiles = getBillingProfiles_();
    const targetCodes = {};
    profiles.forEach(function(profile) {
      if (normalizeCode_(profile.coachCode) === coachCode &&
        looksLikeTestParticipantName_(profile.studentName || profile.studentCode)) {
        targetCodes[normalizeCode_(profile.studentCode)] = true;
      }
    });

    const codeList = Object.keys(targetCodes).filter(Boolean);
    if (!codeList.length) {
      return {
        ok: true,
        action: "cleanupLeaveTestData",
        removedBillingProfiles: 0,
        removedLessons: 0,
        removedLeaveRecords: 0,
        targetCodes: []
      };
    }

    const nextProfiles = profiles.filter(function(profile) {
      return !(normalizeCode_(profile.coachCode) === coachCode && targetCodes[normalizeCode_(profile.studentCode)]);
    });
    const lessons = getLessons_();
    const nextLessons = lessons.filter(function(lesson) {
      return !(normalizeCode_(lesson.coachCode) === coachCode && targetCodes[normalizeCode_(lesson.studentCode)]);
    });
    const leaveRecords = getLeaveRecords_();
    const nextLeaveRecords = leaveRecords.filter(function(record) {
      return !(normalizeCode_(record.coachCode) === coachCode && targetCodes[normalizeCode_(record.studentCode)]);
    });

    setBillingProfiles_(nextProfiles);
    setLessons_(nextLessons);
    setLeaveRecords_(nextLeaveRecords);

    return {
      ok: true,
      action: "cleanupLeaveTestData",
      removedBillingProfiles: profiles.length - nextProfiles.length,
      removedLessons: lessons.length - nextLessons.length,
      removedLeaveRecords: leaveRecords.length - nextLeaveRecords.length,
      targetCodes: codeList
    };
  });
}

// ====================================================================
// 完整課表雲端化（Lessons 分頁）
//
// 計費要算「每堂 attendanceStatus + charged」，但課表過去沒完整存雲端：雲端只有
// LeaveRecords（請假紀錄）+ BillingProfiles（凍結快照），完整課表只在開過請假系統
// 那台裝置的 localStorage + Google Calendar。但 Google Calendar 不是完整真相
// （臨時請假/未到課要扣堂、重大急事不扣堂，這三種只改 lesson 內部欄位、不碰日曆），
// 光看日曆算不準。把完整課表存上雲端，主系統 listLessons 讀回用權威算法即時重算，
// 才是「學生本期已扣 X/4 顯示錯誤」的真根治。
// ====================================================================

// calendarOccupied / charged 是 boolean，但 Sheet cell round-trip 後可能回 boolean
// 或字串（手動編輯、匯入、不同 client）。兩種都吃，避免把扣堂與否判錯。
function normalizeLessonBoolean_(value) {
  if (value === true) {
    return true;
  }
  if (value === false) {
    return false;
  }
  const text = String(value === null || value === undefined ? "" : value).trim().toLowerCase();
  if (text === "true") {
    return true;
  }
  if (text === "false") {
    return false;
  }
  return false;
}

function normalizeLessonRow_(row) {
  row = row || {};
  return {
    id: safeString_(row.id, ""),
    studentCode: normalizeCode_(row.studentCode),
    coachCode: normalizeCode_(row.coachCode),
    startAt: safeString_(row.startAt, ""),
    sourceType: safeString_(row.sourceType, "REGULAR"),
    calendarOccupied: normalizeLessonBoolean_(row.calendarOccupied),
    attendanceStatus: safeString_(row.attendanceStatus, "scheduled"),
    charged: normalizeLessonBoolean_(row.charged),
    completedAt: safeString_(row.completedAt, ""),
    updatedAt: safeString_(row.updatedAt, ""),
    calendarEventId: safeString_(row.calendarEventId, ""),
    cancelledByCoachLeaveBlockId: safeString_(row.cancelledByCoachLeaveBlockId, ""),
    coachLeaveReason: safeString_(row.coachLeaveReason, ""),
    beforeCoachLeaveJson: safeString_(row.beforeCoachLeaveJson, "")
  };
}

function getLessons_() {
  return readTable_(SHEETS.lessons).map(normalizeLessonRow_);
}

// 整批寫回：清資料區（保留 header）再 bulk write。
// 重要：絕不可 slice 截斷。歷史扣堂課（已上完、charged=true）一旦被截掉，
// 主系統重算的「本期已扣」會少算。Lessons 會持續累積但 Sheet 容量實質無上限。
function setLessons_(lessons) {
  const sheet = getCoachflowSpreadsheet_().getSheetByName(SHEETS.lessons);
  const headers = SCHEMA.Lessons;
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
  }
  if (lessons && lessons.length) {
    appendRows_(SHEETS.lessons, lessons.map(normalizeLessonRow_));
  }
}

function listLessons_(payload) {
  const coachCode = normalizeCode_(payload.coachCode);
  const studentCode = normalizeCode_(payload.studentCode);
  const lessons = getLessons_()
    .filter(function(lesson) {
      return (!coachCode || normalizeCode_(lesson.coachCode) === coachCode) &&
        (!studentCode || normalizeCode_(lesson.studentCode) === studentCode);
    })
    .sort(function(a, b) {
      return new Date(b.startAt || 0) - new Date(a.startAt || 0);
    });
  return {
    ok: true,
    action: "listLessons",
    count: lessons.length,
    lessons: lessons
  };
}

// 差量 upsert by id（不是整 scope 覆寫）：
// 學生端只會 push 自己的課；若做整 scope 覆寫，學生一推就會把「同教練底下其他
// 學生的課」整批清掉。所以只 upsert 進來的那幾筆，其餘原封不動。
//
// +2000ms stale 防護：incoming.updatedAt 明顯比現有舊（落後 2 秒以上）就忽略，
// 避免落後裝置用過期狀態覆寫較新的扣堂/請假結果（對齊 saveBillingProfile_）。
function saveLessons_(payload) {
  const incoming = Array.isArray(payload.lessons) ? payload.lessons : [];
  const authoritativeClient = payload.authoritativeClient === true ||
    String(payload.authoritativeClient || "").toLowerCase() === "true";
  const normalizedIncoming = incoming
    .map(normalizeLessonRow_)
    .filter(function(lesson) {
      return lesson.id;
    });
  if (!normalizedIncoming.length) {
    return {
      ok: true,
      action: "saveLessons",
      saved: 0,
      message: "No lessons to save."
    };
  }

  return withScriptLock_(function () {
    const existing = getLessons_();
    const indexById = {};
    existing.forEach(function(lesson, index) {
      if (lesson.id) {
        indexById[lesson.id] = index;
      }
    });

    let inserted = 0;
    let updated = 0;
    let ignored = 0;
    normalizedIncoming.forEach(function(lesson) {
      const existingIndex = indexById[lesson.id];
      if (existingIndex === undefined) {
        existing.push(lesson);
        indexById[lesson.id] = existing.length - 1;
        inserted += 1;
        return;
      }
      const current = existing[existingIndex];
      const incomingTime = new Date(lesson.updatedAt || 0).getTime();
      const currentTime = new Date(current.updatedAt || 0).getTime();
      if (Number.isFinite(incomingTime) && Number.isFinite(currentTime) && incomingTime + 2000 < currentTime) {
        ignored += 1;
        return;
      }
      existing[existingIndex] = authoritativeClient
        ? lesson
        : Object.assign({}, lesson, {
          calendarEventId: current.calendarEventId || "",
          cancelledByCoachLeaveBlockId: current.cancelledByCoachLeaveBlockId || "",
          coachLeaveReason: current.coachLeaveReason || "",
          beforeCoachLeaveJson: current.beforeCoachLeaveJson || ""
        });
      updated += 1;
    });

    setLessons_(existing);
    return {
      ok: true,
      action: "saveLessons",
      saved: inserted + updated,
      inserted: inserted,
      updated: updated,
      ignored: ignored,
      total: existing.length
    };
  });
}

// ====================================================================
// 重複課堂盤點／一次性清理（只供 Apps Script 編輯器手動執行）
// ====================================================================

function getLessonDuplicateGroups_(lessons) {
  const grouped = {};
  const singles = [];
  (lessons || []).forEach(function(lesson) {
    const key = lessonCloudKey_(lesson.studentCode, lesson.coachCode, lesson.startAt);
    if (!key) {
      singles.push([lesson]);
      return;
    }
    grouped[key] = grouped[key] || [];
    grouped[key].push(lesson);
  });
  return Object.keys(grouped).map(function(key) {
    return { key: key, lessons: grouped[key] };
  }).concat(singles.map(function(items) {
    return { key: "", lessons: items };
  }));
}

function getLessonDedupScore_(lesson, activeLeaveLessonIds) {
  let score = 0;
  if (activeLeaveLessonIds[String(lesson.id || "")]) score += 10000;
  if (String(lesson.sourceType || "") !== "DUPLICATE") score += 1000;
  if (String(lesson.attendanceStatus || "") === "leave-normal") score += 500;
  if (normalizeLessonBoolean_(lesson.charged)) score += 200;
  if (normalizeLessonBoolean_(lesson.calendarOccupied)) score += 100;
  if (safeString_(lesson.calendarEventId, "")) score += 50;
  return score;
}

function auditDuplicateLessons() {
  ensureSheets_();
  const lessons = getLessons_();
  const groups = getLessonDuplicateGroups_(lessons).filter(function(group) {
    return group.key && group.lessons.length > 1;
  });
  const report = {
    ok: true,
    totalLessons: lessons.length,
    duplicateGroups: groups.length,
    duplicateRows: groups.reduce(function(total, group) { return total + group.lessons.length; }, 0),
    removableRows: groups.reduce(function(total, group) { return total + group.lessons.length - 1; }, 0)
  };
  Logger.log("[lesson-audit] " + JSON.stringify(report));
  return report;
}

function cleanupDuplicateLessons() {
  ensureSheets_();
  return withScriptLock_(function() {
    const lessons = getLessons_();
    const leaveRecords = getLeaveRecords_();
    const activeLeaveLessonIds = {};
    const activeLeaveKeys = {};
    leaveRecords.forEach(function(record) {
      if (safeString_(record.revokedAt, "") || String(record.type || "normal") !== "normal") {
        return;
      }
      if (record.lessonId) activeLeaveLessonIds[String(record.lessonId)] = true;
      const key = safeString_(record.lessonKey, "") ||
        lessonCloudKey_(record.studentCode, record.coachCode, record.lessonStartAt);
      if (key) activeLeaveKeys[key] = true;
    });

    const replacementIds = {};
    const cleaned = [];
    let duplicateGroups = 0;
    let removedLessons = 0;
    getLessonDuplicateGroups_(lessons).forEach(function(group) {
      if (!group.key || group.lessons.length < 2) {
        cleaned.push.apply(cleaned, group.lessons);
        return;
      }
      duplicateGroups += 1;
      const sorted = group.lessons.slice().sort(function(left, right) {
        const scoreDifference = getLessonDedupScore_(right, activeLeaveLessonIds) -
          getLessonDedupScore_(left, activeLeaveLessonIds);
        if (scoreDifference) return scoreDifference;
        const timeDifference = new Date(right.updatedAt || 0).getTime() -
          new Date(left.updatedAt || 0).getTime();
        if (Number.isFinite(timeDifference) && timeDifference) return timeDifference;
        return String(left.id || "").localeCompare(String(right.id || ""));
      });
      const winner = sorted[0];
      sorted.slice(1).forEach(function(loser) {
        replacementIds[String(loser.id || "")] = String(winner.id || "");
        if (!winner.calendarEventId && loser.calendarEventId) winner.calendarEventId = loser.calendarEventId;
        removedLessons += 1;
      });
      if (activeLeaveKeys[group.key]) {
        winner.attendanceStatus = "leave-normal";
        winner.calendarOccupied = false;
        winner.charged = false;
        winner.updatedAt = nowIso_();
      }
      cleaned.push(winner);
    });

    let repointedLeaveRecords = 0;
    leaveRecords.forEach(function(record) {
      const replacementId = replacementIds[String(record.lessonId || "")];
      if (replacementId) {
        record.lessonId = replacementId;
        record.lessonKey = lessonCloudKey_(record.studentCode, record.coachCode, record.lessonStartAt);
        record.updatedAt = nowIso_();
        repointedLeaveRecords += 1;
      }
    });

    if (removedLessons) {
      setLessons_(cleaned);
      setLeaveRecords_(leaveRecords);
    }
    const report = {
      ok: true,
      before: lessons.length,
      after: cleaned.length,
      duplicateGroups: duplicateGroups,
      removedLessons: removedLessons,
      repointedLeaveRecords: repointedLeaveRecords
    };
    Logger.log("[lesson-cleanup] " + JSON.stringify(report));
    return report;
  });
}
// ====================================================================
// 雲端自動整理排程（時間觸發器）
//
// 過去課表狀態的「整理」（請假套用到課、課程時間過後自動扣堂）全部跑在瀏覽器：
// 教練電腦沒開，雲端資料就停止保鮮——這正是「系統更新都有問題」的根因。
// 這裡把整理邏輯搬進 Apps Script 時間觸發器，每小時自動跑，完全不依賴任何
// 裝置開著：
//
//   1. 請假對齊：active 的 normal 請假紀錄 → 對應 Lessons 列改 leave-normal
//      （不扣堂）。防護「學生請假成功但 lesson push 失敗 → 課表列還停在
//      scheduled → 被誤扣堂」。
//   2. 取消還原：請假已 revoked 但課表列還卡 leave-normal → 還原 scheduled。
//   3. 自動扣堂：scheduled + 佔用日曆 + 時間已過 → charged=true。
//      主系統 app.js 讀雲端課表時本來就會即時算，這裡是把結果落地回 Sheet，
//      讓 Sheet 本身永遠是對的（教練直接看 Sheet 也對）。
//
// 部署：貼上這版 Code.gs 後，在 Apps Script Editor 手動執行一次
// setupCoachflowTriggers（會要求授權），之後每小時自動跑。idempotent，
// 多跑不會壞資料。
// ====================================================================

// 與前端 leave-sandbox.js / app.js 的 BILLING_TRACKING_START_AT 對齊：
// 5/3（含）之後的課才納入追蹤與自動整理。
const LESSONS_TRACKING_START_AT = "2026-05-03T00:00:00+08:00";

function lessonCloudKey_(studentCode, coachCode, startAt) {
  const time = new Date(startAt || "").getTime();
  if (!studentCode || !coachCode || !Number.isFinite(time)) {
    return "";
  }
  return normalizeCode_(studentCode) + "|" + normalizeCode_(coachCode) + "|" + new Date(time).toISOString();
}

function reconcileAndAutoCompleteLessons() {
  return withScriptLock_(function () {
    const nowMs = Date.now();
    const trackingStartMs = new Date(LESSONS_TRACKING_START_AT).getTime();
    const lessons = getLessons_();
    if (!lessons.length) {
      Logger.log("[reconcile] Lessons 分頁沒有資料，跳過。");
      return;
    }

    // 只看 normal 請假（temporary-leave / no-show / major-case 是 lesson 內部
    // 狀態、教練停課另有流程，都不在這裡動）。
    const leaves = getLeaveRecords_().filter(function(record) {
      return String(record.type || "normal") === "normal";
    });
    const activeByLessonId = {};
    const activeByKey = {};
    const revokedByLessonId = {};
    const revokedByKey = {};
    leaves.forEach(function(record) {
      const idKey = safeString_(record.lessonId, "");
      const lessonKey = safeString_(record.lessonKey, "") ||
        lessonCloudKey_(record.studentCode, record.coachCode, record.lessonStartAt);
      const isActive = !safeString_(record.revokedAt, "");
      const idMap = isActive ? activeByLessonId : revokedByLessonId;
      const keyMap = isActive ? activeByKey : revokedByKey;
      if (idKey) {
        idMap[idKey] = true;
      }
      if (lessonKey) {
        keyMap[lessonKey] = true;
      }
    });

    let leaveApplied = 0;
    let leaveRestored = 0;
    let autoCompleted = 0;
    const now = nowIso_();

    lessons.forEach(function(lesson) {
      const startMs = new Date(lesson.startAt || "").getTime();
      if (!Number.isFinite(startMs) || startMs < trackingStartMs) {
        return;
      }
      const key = lessonCloudKey_(lesson.studentCode, lesson.coachCode, lesson.startAt);
      const hasActiveLeave = Boolean(activeByLessonId[lesson.id] || (key && activeByKey[key]));
      const hasRevokedLeave = Boolean(revokedByLessonId[lesson.id] || (key && revokedByKey[key]));

      // (1) 請假成立但課表列還停在 scheduled → 套用 leave-normal、不扣堂
      if (hasActiveLeave && lesson.attendanceStatus === "scheduled") {
        lesson.attendanceStatus = "leave-normal";
        lesson.calendarOccupied = false;
        lesson.charged = false;
        lesson.updatedAt = now;
        leaveApplied += 1;
        return;
      }

      // (2) 請假已取消但課表列卡在 leave-normal → 還原 scheduled
      //（接著仍會走 (3)：若時間已過，照常自動扣堂）
      if (!hasActiveLeave && hasRevokedLeave && lesson.attendanceStatus === "leave-normal") {
        lesson.attendanceStatus = "scheduled";
        lesson.calendarOccupied = true;
        lesson.charged = false;
        lesson.updatedAt = now;
        leaveRestored += 1;
      }

      // (3) 自動扣堂：與前端 isLessonAutoCompleted 同語意
      if (
        lesson.attendanceStatus === "scheduled" &&
        lesson.calendarOccupied &&
        !lesson.charged &&
        startMs <= nowMs
      ) {
        lesson.charged = true;
        lesson.completedAt = lesson.completedAt || now;
        lesson.updatedAt = now;
        autoCompleted += 1;
      }
    });

    if (leaveApplied || leaveRestored || autoCompleted) {
      setLessons_(lessons);
    }
    Logger.log(
      "[reconcile] 完成：套用請假 " + leaveApplied +
      " 堂、還原已取消請假 " + leaveRestored +
      " 堂、自動扣堂 " + autoCompleted + " 堂（共掃 " + lessons.length + " 堂）。"
    );
  });
}

// 一次性安裝：刪掉同名舊觸發器再建新的，避免重複。在 Editor 手動執行一次即可。
function setupCoachflowTriggers() {
  const handler = "reconcileAndAutoCompleteLessons";
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === handler) {
      ScriptApp.deleteTrigger(trigger);
      removed += 1;
    }
  });
  ScriptApp.newTrigger(handler).timeBased().everyHours(1).create();
  Logger.log("[trigger] 已建立每小時排程 " + handler + "（移除舊觸發器 " + removed + " 個）。");
}

// === Snapshot 端點停用 ===
// 過去 snapshot 把整份 coach state（lessons / leaves / makeups / billing）序列化
// 進 PropertiesService，是 500KB 配額爆炸的主要元兇。原始資料現在都在 Sheet
// 與 Google Calendar，snapshot 只是個 cache 層，可以直接拿掉：
//
// - getLeaveStateSnapshot_ 一律回 found:false，前端會 fallback 走
//   listLeaveRecords + listBillingProfiles + listEvents 重建狀態。
// - saveLeaveStateSnapshot_ 直接 no-op，回 ok:true 避免前端 retry。
//
// 老 client 仍會 call 這兩支，所以保留 action handler，只是 body 改空殼。
// 待所有 PWA 都更新到新版後可考慮整個 case 區段刪除。
function getLeaveStateSnapshot_(payload) {
  return {
    ok: true,
    action: "getLeaveStateSnapshot",
    found: false,
    snapshot: null,
    migrated: true,
    message: "Snapshot storage discontinued; client should fall back to listLeaveRecords + listBillingProfiles + listEvents."
  };
}

function saveLeaveStateSnapshot_(payload) {
  return {
    ok: true,
    action: "saveLeaveStateSnapshot",
    saved: false,
    migrated: true,
    message: "Snapshot storage discontinued; saveLeaveRecord / saveBillingProfile is now the canonical persistence."
  };
}

function normalizeLeaveStateSnapshot_(snapshot) {
  snapshot = snapshot || {};
  const now = nowIso_();
  return {
    version: safeString_(snapshot.version, "1"),
    schemaVersion: safeString_(snapshot.schemaVersion, ""),
    coachCode: normalizeCode_(snapshot.coachCode),
    studentCode: normalizeCode_(snapshot.studentCode),
    source: safeString_(snapshot.source, ""),
    sourceVersion: safeString_(snapshot.sourceVersion, ""),
    updatedAt: safeString_(snapshot.updatedAt, now),
    updatedBy: normalizeCode_(snapshot.updatedBy || snapshot.coachCode || "SYSTEM"),
    createdAt: safeString_(snapshot.createdAt, ""),
    counts: snapshot.counts && typeof snapshot.counts === "object" ? snapshot.counts : {},
    state: snapshot.state && typeof snapshot.state === "object" ? snapshot.state : {}
  };
}

function getLeaveStateSnapshotKey_(payload) {
  const coachCode = normalizeCode_(payload && payload.coachCode);
  return "LEAVE_STATE_SNAPSHOT_V1_" + (coachCode || "GLOBAL");
}

function normalizeBillingReminderLogs_(logs) {
  let source = logs;
  if (typeof source === "string") {
    try {
      source = JSON.parse(source);
    } catch (error) {
      source = [];
    }
  }
  if (!Array.isArray(source)) {
    return [];
  }
  return source
    .map(function(item) {
      return {
        id: safeString_(item && item.id, ""),
        milestone: toNonNegativeInt_(item && item.milestone, 0),
        label: safeString_(item && item.label, ""),
        status: safeString_(item && item.status, "") === "success" ? "success" : "failed",
        sentAt: safeString_(item && item.sentAt, ""),
        to: safeString_(item && item.to, ""),
        note: safeString_(item && item.note, ""),
        triggerSource: safeString_(item && item.triggerSource, "")
      };
    })
    .filter(function(item) {
      return item.milestone > 0;
    })
    .sort(function(a, b) {
      return new Date(b.sentAt || 0) - new Date(a.sentAt || 0);
    })
    .slice(0, 24);
}

function normalizePaymentStatus_(value) {
  const text = safeString_(value, "unknown");
  return text === "paid" || text === "unpaid" ? text : "unknown";
}

function toNonNegativeInt_(value, fallback) {
  const n = Number(value);
  if (!Number.isFinite(n)) {
    return fallback || 0;
  }
  return Math.max(0, Math.floor(n));
}

function hasNumericValue_(value) {
  return value !== undefined && value !== null && value !== "" && Number.isFinite(Number(value));
}

function readChunkedProperty_(baseKey) {
  const props = PropertiesService.getScriptProperties();
  // 新格式：用 _ACTIVE 指向 A 或 B chunk set，原子切換
  const activeSuffix = props.getProperty(baseKey + "_ACTIVE") || "";
  if (activeSuffix) {
    const chunkCount = Number(props.getProperty(baseKey + "_" + activeSuffix + "_CHUNKS") || 0);
    if (!chunkCount) {
      return "";
    }
    const chunks = [];
    for (var i = 0; i < chunkCount; i += 1) {
      chunks.push(props.getProperty(baseKey + "_" + activeSuffix + "_" + i) || "");
    }
    return chunks.join("");
  }
  // 舊格式：保留向後相容（讀舊資料用），新版寫入會把它清掉
  const chunkCount = Number(props.getProperty(baseKey + "_CHUNKS") || 0);
  if (!chunkCount) {
    return props.getProperty(baseKey) || "";
  }
  const chunks = [];
  for (var index = 0; index < chunkCount; index += 1) {
    chunks.push(props.getProperty(baseKey + "_" + index) || "");
  }
  return chunks.join("");
}

function writeChunkedProperty_(baseKey, value) {
  const props = PropertiesService.getScriptProperties();
  const text = String(value || "");
  const chunkSize = 8000;
  const chunkCount = Math.max(1, Math.ceil(text.length / chunkSize));

  // 500KB 總配額護欄：每筆 value 自己若超過 480KB 就拒絕，
  // 避免把整份 properties 撐爆而中斷其他寫入。
  if (text.length > 480000) {
    throw new Error("writeChunkedProperty_: payload too large (" + text.length + " bytes), refusing to write.");
  }

  // Atomic A/B swap：先寫到 standby suffix、最後一步才切 ACTIVE。
  // 即使 setProperty 中斷，舊 active 資料完整保留，下一次寫入會重試。
  const activeSuffix = props.getProperty(baseKey + "_ACTIVE") || "A";
  const standbySuffix = activeSuffix === "A" ? "B" : "A";

  for (var i = 0; i < chunkCount; i += 1) {
    props.setProperty(baseKey + "_" + standbySuffix + "_" + i, text.slice(i * chunkSize, (i + 1) * chunkSize));
  }
  props.setProperty(baseKey + "_" + standbySuffix + "_CHUNKS", String(chunkCount));

  // 原子切換指針
  props.setProperty(baseKey + "_ACTIVE", standbySuffix);

  // 清舊 active 資料（這之後出問題不影響讀取）
  const oldActiveCount = Number(props.getProperty(baseKey + "_" + activeSuffix + "_CHUNKS") || 0);
  for (var j = 0; j < oldActiveCount; j += 1) {
    try {
      props.deleteProperty(baseKey + "_" + activeSuffix + "_" + j);
    } catch (e) {
      // 失敗忽略，下次 swap 會再清
    }
  }
  try { props.deleteProperty(baseKey + "_" + activeSuffix + "_CHUNKS"); } catch (e) {}

  // 一次性清舊格式（pre-A/B 時代留下的）
  const legacyCount = Number(props.getProperty(baseKey + "_CHUNKS") || 0);
  if (legacyCount) {
    for (var k = 0; k < legacyCount; k += 1) {
      try { props.deleteProperty(baseKey + "_" + k); } catch (e) {}
    }
    try { props.deleteProperty(baseKey + "_CHUNKS"); } catch (e) {}
  }
  try { props.deleteProperty(baseKey); } catch (e) {}
}

// ====================================================================
// 一次性遷移：把 PropertiesService 的 LEAVE_RECORDS_V1 / BILLING_PROFILES_V1
// 倒進 LeaveRecords / BillingProfiles 兩張 Sheet，然後清掉舊 properties 與
// snapshot blob，釋放配額空間。
//
// 部署流程：
//   1. 部署這版 Code.gs 到 Apps Script Editor
//   2. 在 Editor 手動執行 migrateLeaveAndBillingToSheets
//      （Run 選單 → 選函式 → Run；看 Log 確認筆數正確）
//   3. 開 Sheet 肉眼確認 LeaveRecords / BillingProfiles 內容無誤
//   4. 再執行 cleanupLegacyLeaveProperties 清掉 Properties
//   5. 在 Editor「部署 → 管理部署 → 新版本」更新 URL 部署版
//
// 兩個函式都是 idempotent，多執行幾次不會壞資料。
// ====================================================================

function migrateLeaveAndBillingToSheets() {
  ensureSheets_();

  // ---- LEAVE_RECORDS_V1 → LeaveRecords ----
  const leaveRaw = readChunkedProperty_("LEAVE_RECORDS_V1");
  if (leaveRaw) {
    try {
      const parsed = JSON.parse(leaveRaw);
      if (Array.isArray(parsed) && parsed.length) {
        const existing = readTable_(SHEETS.leaveRecords);
        if (existing.length === 0) {
          setLeaveRecords_(parsed);
          Logger.log("[migrate] LeaveRecords 從 properties 寫入 " + parsed.length + " 筆。");
        } else {
          Logger.log("[migrate] LeaveRecords Sheet 已有 " + existing.length + " 筆，跳過寫入避免覆蓋。");
        }
      } else {
        Logger.log("[migrate] LEAVE_RECORDS_V1 內容為空或非陣列，跳過。");
      }
    } catch (error) {
      Logger.log("[migrate] LEAVE_RECORDS_V1 parse 失敗：" + error);
    }
  } else {
    Logger.log("[migrate] PropertiesService 沒有 LEAVE_RECORDS_V1，跳過。");
  }

  // ---- BILLING_PROFILES_V1 → BillingProfiles ----
  const billingRaw = readChunkedProperty_("BILLING_PROFILES_V1");
  if (billingRaw) {
    try {
      const parsed = JSON.parse(billingRaw);
      if (Array.isArray(parsed) && parsed.length) {
        const existing = readTable_(SHEETS.billingProfiles);
        if (existing.length === 0) {
          setBillingProfiles_(parsed);
          Logger.log("[migrate] BillingProfiles 從 properties 寫入 " + parsed.length + " 筆。");
        } else {
          Logger.log("[migrate] BillingProfiles Sheet 已有 " + existing.length + " 筆，跳過寫入避免覆蓋。");
        }
      } else {
        Logger.log("[migrate] BILLING_PROFILES_V1 內容為空或非陣列，跳過。");
      }
    } catch (error) {
      Logger.log("[migrate] BILLING_PROFILES_V1 parse 失敗：" + error);
    }
  } else {
    Logger.log("[migrate] PropertiesService 沒有 BILLING_PROFILES_V1，跳過。");
  }

  Logger.log("[migrate] 完成。請肉眼檢查 Sheet 兩張分頁的資料筆數與內容，再執行 cleanupLegacyLeaveProperties() 清掉 Properties。");
}

function cleanupLegacyLeaveProperties() {
  const props = PropertiesService.getScriptProperties();
  const allKeys = props.getKeys();
  const LEGACY_BASE_KEYS = ["LEAVE_RECORDS_V1", "BILLING_PROFILES_V1"];
  const SNAPSHOT_PREFIX = "LEAVE_STATE_SNAPSHOT_V1_";

  let deletedCount = 0;
  let snapshotDeleted = 0;
  let legacyDeleted = 0;

  allKeys.forEach(function(key) {
    let shouldDelete = false;
    let isSnapshot = false;

    // base 與其 chunked 子鍵（_ACTIVE, _A_CHUNKS, _A_0, _B_CHUNKS, _B_0...; 舊格式 _CHUNKS, _0...）
    for (var i = 0; i < LEGACY_BASE_KEYS.length; i += 1) {
      var base = LEGACY_BASE_KEYS[i];
      if (key === base || key.indexOf(base + "_") === 0) {
        shouldDelete = true;
        break;
      }
    }
    // snapshot 系列
    if (!shouldDelete && key.indexOf(SNAPSHOT_PREFIX) === 0) {
      shouldDelete = true;
      isSnapshot = true;
    }

    if (shouldDelete) {
      try {
        props.deleteProperty(key);
        deletedCount += 1;
        if (isSnapshot) {
          snapshotDeleted += 1;
        } else {
          legacyDeleted += 1;
        }
      } catch (error) {
        Logger.log("[cleanup] 刪 " + key + " 失敗：" + error);
      }
    }
  });

  Logger.log("[cleanup] 完成。共刪 " + deletedCount + " 個 properties（legacy " + legacyDeleted + " + snapshot " + snapshotDeleted + "）。");
  Logger.log("[cleanup] 剩下的 properties keys：" + JSON.stringify(props.getKeys()));
}

/**
 * 一次性：清掉 WorkoutLogs 分頁裡有相同 log_id 的真重複列。
 * 重複定義 = 同 log_id 出現多次（並行 submit race / 同一筆 retry 直接 append 出來
 * 的副本）。多筆並存時保留 updated_at 最新那一筆、其餘刪除。
 *
 * 注意：先前版本誤用 (student + program + activity_date + exercise) 當重複 key，
 * 會把「同一天送了多次同動作」這種合法多筆紀錄硬合掉。已修正成只比 log_id。
 * 部署新版 Code.gs 後在 Editor 跑一次即可。idempotent。
 */
function dedupeWorkoutLogs() {
  return withScriptLock_(function () {
    const sheet = getCoachflowSpreadsheet_().getSheetByName(SHEETS.workoutLogs);
    if (!sheet) {
      Logger.log("[dedupe] 找不到 WorkoutLogs 分頁。");
      return;
    }
    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) {
      Logger.log("[dedupe] WorkoutLogs 沒有資料列。");
      return;
    }
    const headers = values[0];
    const headerIndex = {};
    headers.forEach(function(h, i) { headerIndex[h] = i; });
    if (!("log_id" in headerIndex) || !("updated_at" in headerIndex)) {
      Logger.log("[dedupe] 缺 log_id 或 updated_at 欄位，中止。");
      return;
    }

    const records = [];
    for (var r = 1; r < values.length; r += 1) {
      const row = values[r];
      const obj = {};
      headers.forEach(function(h, i) { obj[h] = row[i]; });
      records.push({ rowIndex: r + 1, data: obj });
    }

    // 群組：只依 log_id
    const groups = {};
    records.forEach(function(rec) {
      const key = String(rec.data.log_id || "").trim();
      if (!key) {
        return; // 沒 log_id 的列就不動，避免把 import 進來的舊資料誤殺
      }
      if (!groups[key]) {
        groups[key] = [];
      }
      groups[key].push(rec);
    });

    const toDelete = [];
    Object.keys(groups).forEach(function(key) {
      const list = groups[key];
      if (list.length < 2) {
        return;
      }
      list.sort(function(a, b) {
        return new Date(b.data.updated_at || 0) - new Date(a.data.updated_at || 0);
      });
      list.slice(1).forEach(function(rec) { toDelete.push(rec.rowIndex); });
      Logger.log("[dedupe] log_id=" + key + " 重複 " + list.length + " 筆，保留 updated_at " + list[0].data.updated_at + "，刪除其餘 " + (list.length - 1) + " 筆。");
    });

    toDelete.sort(function(a, b) { return b - a; });
    toDelete.forEach(function(rowIndex) {
      sheet.deleteRow(rowIndex);
    });

    SpreadsheetApp.flush();
    Logger.log("[dedupe] 完成。共刪 " + toDelete.length + " 列同 log_id 的重複 WorkoutLogs。");
  });
}

/**
 * 用法：withScriptLock_(function () { ... 寫入動作 ... })
 * 多裝置同時寫同一份 chunked property 會 lost update。
 * 包進 ScriptLock 後同一時間只允許一個寫入交易。
 * 取不到鎖在 30 秒內 throw，呼叫端可決定是否回傳「忙碌請重試」。
 *
 * 等待時間從 5s 拉長到 30s：完整課表第一次灌入（saveLessons 分批）+
 * billing 快照補種 + 請假紀錄會同時到達，5s 太短會大量 LOCK_TIMEOUT。
 * Apps Script 單次執行上限 6 分鐘，30s 等鎖完全安全；client 端 12s 逾時
 * 中斷也沒關係——server 端會繼續排隊完成寫入，寫入全部是 upsert 冪等，
 * client 重試不會產生重複資料。
 */
function withScriptLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error("LOCK_TIMEOUT: could not acquire script lock within 30s, please retry.");
  }
  try {
    return fn();
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function normalizeCode_(value) {
  return safeString_(value, "").toUpperCase();
}

function safeString_(value, fallback) {
  const text = value === null || value === undefined ? "" : String(value).trim();
  return text || (fallback || "");
}

function nowIso_() {
  return new Date().toISOString();
}

// ====================================================================
// 阻擋自動測試／評分腳本把假資料灌進正式資料庫
//
// 已上線事故：外部自動化測試直接對正式後端 POST saveProgram / upsertCoach /
// upsertStudent，每天在教練的正式資料庫塞滿 WF* / GRADE* / TEST* 課表與
// Grader Coach / WF2 Coach / 測試教練 等假帳號。一直手動刪是打地鼠。
//
// 這裡在「寫入當下」就擋掉這些假資料。設計目標：真實資料 0 誤殺。
//   - 真教練：Monster Chang（不含測試字樣）
//   - 真學生：中文姓名（不含測試字樣）
//   - 真課表代碼：純數字（日期，如 503 / 6060612），絕不含英文字母
// 測試指紋：課表代碼含英文字母、或教練/學生名含 WF/GRADE/TEST/Grader/
// Workflow/測試/Weekend 等字樣。
//
// 逃生門：若有正當需求要暫時放行（例如自己要建英文代碼課表），到
// Script Properties 設 COACHFLOW_BLOCK_TEST_DATA = "false" 即可關閉本守門。
// ====================================================================
function isTestPollutionBlocked_() {
  const v = String(PropertiesService.getScriptProperties().getProperty("COACHFLOW_BLOCK_TEST_DATA") || "")
    .trim().toLowerCase();
  return v !== "false" && v !== "0" && v !== "off" && v !== "no";
}

function looksLikeTestParticipantName_(value) {
  const s = String(value === null || value === undefined ? "" : value).trim();
  if (!s) {
    return false;
  }
  const explicitPrefixes = [
    "\u6e2c\u8a66\u5b78\u751f",
    "\u9031\u672b\u6e2c\u8a66\u5b78\u751f",
    "\u5468\u672b\u6e2c\u8a66\u5b78\u751f",
    "WFStuA"
  ];
  if (explicitPrefixes.some(function(prefix) {
    return s.indexOf(prefix) === 0;
  })) {
    return true;
  }
  // 真資料：Monster Chang、中文姓名，都不會命中。
  return /grader|workflow|測試|週末|流程|學員|學生|student|weekend|\bwf\d*\b|\bw2s?\b|\bw5\b|\bw6\b|\bwflow\b|\bgrd\w*|\bgrade\b|\bsync\b|\beval\b|\bwknd\b|\bwkend\b|\btest\b|\bqa\b/i.test(s);
}

function looksLikeTestProgramCode_(code) {
  const s = String(code === null || code === undefined ? "" : code).trim();
  if (!s) {
    return false; // 空白代碼不擋（草稿）
  }
  // 真課表代碼是純數字（日期）。含任何英文字母一律視為自動測試命名。
  return /[A-Za-z]/.test(s);
}

function coachLooksLikeTestPollution_(coach) {
  coach = coach || {};
  return looksLikeTestParticipantName_(coach.coach_name || coach.name);
}

function studentLooksLikeTestPollution_(student) {
  student = student || {};
  return looksLikeTestParticipantName_(student.student_name || student.name) ||
    looksLikeTestParticipantName_(student.primary_coach_name);
}

function programLooksLikeTestPollution_(program) {
  program = program || {};
  return looksLikeTestProgramCode_(program.program_code || program.code) ||
    looksLikeTestParticipantName_(program.coach_name || program.coachName);
}

function testPollutionRejection_(action) {
  return {
    ok: false,
    blocked: true,
    action: action,
    message: "Rejected: payload matches automated-test fingerprint and was blocked to protect production data. Set Script Property COACHFLOW_BLOCK_TEST_DATA=false to disable."
  };
}

function parsePostBody_(e) {
  const params = (e && e.parameter) || {};

  if (params && params.payload) {
    try {
      const parsedPayload = JSON.parse(params.payload);
      return Object.assign({}, parsedPayload, {
        action: params.action || parsedPayload.action || ""
      });
    } catch (error) {
      return Object.assign({}, params, {
        action: params.action || ""
      });
    }
  }

  try {
    const parsedBody = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    return Object.assign({}, parsedBody, {
      action: parsedBody.action || params.action || ""
    });
  } catch (error) {
    return Object.assign({}, params, {
      action: params.action || ""
    });
  }
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
