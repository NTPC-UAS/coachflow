const SHEETS = {
  coaches: "Coaches",
  students: "Students",
  programs: "Programs",
  programItems: "ProgramItems",
  workoutLogs: "WorkoutLogs"
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
    "updated_at"
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
  ]
};

function doGet(e) {
  const action = getAction_(e);

  try {
    ensureSheets_();

    switch (action) {
      case "resolveCoachAccess":
        return jsonResponse_(resolveCoachAccessResponse_(getParam_(e, "access")));

      case "resolveStudentAccess":
        return jsonResponse_(resolveStudentAccessResponse_(getParam_(e, "access")));

      case "bootstrapAdmin":
        return jsonResponse_(buildFullBootstrap_());

      case "bootstrapCoach":
        return jsonResponse_(buildCoachBootstrap_(getParam_(e, "coachId")));

      case "bootstrapStudent":
        return jsonResponse_(buildStudentBootstrap_(getParam_(e, "studentId")));

      case "bootstrap":
      default:
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
    ensureSheets_();

    switch (action) {
      case "resolveCoachAccess":
        return jsonResponse_(resolveCoachAccessResponse_(payload.access));

      case "resolveStudentAccess":
        return jsonResponse_(resolveStudentAccessResponse_(payload.access));

      case "bootstrapAdmin":
        return jsonResponse_(buildFullBootstrap_());

      case "bootstrapCoach":
        return jsonResponse_(buildCoachBootstrap_(payload.coachId));

      case "bootstrapStudent":
        return jsonResponse_(buildStudentBootstrap_(payload.studentId));

      case "upsertCoach":
        upsertCoach_(payload.coach || {});
        return jsonResponse_(buildFullBootstrap_());

      case "setCoachStatus":
        setCoachStatus_(payload.coachId, payload.status);
        return jsonResponse_(buildFullBootstrap_());

      case "deleteCoach":
        deleteCoach_(payload.coachId);
        return jsonResponse_(buildFullBootstrap_());

      case "upsertStudent":
        upsertStudent_(payload.student || {});
        return jsonResponse_(buildFullBootstrap_());

      case "assignStudentCoach":
        assignStudentCoach_(payload.studentId, payload.coachId);
        return jsonResponse_(buildFullBootstrap_());

      case "setStudentStatus":
        setStudentStatus_(payload.studentId, payload.status);
        return jsonResponse_(buildFullBootstrap_());

      case "deleteStudent":
        deleteStudent_(payload.studentId);
        return jsonResponse_(buildFullBootstrap_());

      case "saveProgram":
        saveProgram_(payload.program || {}, payload.items || []);
        return jsonResponse_(buildFullBootstrap_());

      case "publishProgram":
        publishProgram_(payload.programId, payload.coachId);
        return jsonResponse_(buildFullBootstrap_());

      case "submitWorkoutLogs":
        submitWorkoutLogs_(payload.logs || []);
        return jsonResponse_(buildFullBootstrap_());

      case "importWorkoutLogs":
        importWorkoutLogs_(payload.logs || []);
        return jsonResponse_(buildFullBootstrap_());

      case "touchCoach":
        touchCoach_(payload.coachId);
        return jsonResponse_(buildCoachBootstrap_(payload.coachId));

      case "touchStudent":
        touchStudent_(payload.studentId);
        return jsonResponse_(buildStudentBootstrap_(payload.studentId));

      case "createEvent":
        return jsonResponse_(createCalendarEvent_(payload));

      case "updateEvent":
        return jsonResponse_(updateCalendarEvent_(payload));

      case "deleteEvent":
        return jsonResponse_(deleteCalendarEvent_(payload));

      case "deleteSingleEvent":
        return jsonResponse_(deleteCalendarEvent_(Object.assign({}, payload || {}, {
          deleteScope: "single",
          singleEventOnly: true,
          strictSingleOccurrence: true
        })));

      case "sendEmail":
        return jsonResponse_(sendEmailNotice_(payload));

      default:
        return jsonResponse_({
          ok: false,
          message: "Unsupported action."
        });
    }
  } catch (error) {
    return jsonResponse_({
      ok: false,
      message: error.message || String(error)
    });
  }
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
        return row.coach_id === student.primary_coach_id;
      })
    : [];

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
    programItems: programItems,
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
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SCHEMA).forEach(function(sheetName) {
    ensureSheet_(spreadsheet, sheetName, SCHEMA[sheetName]);
  });
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
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
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

  if (!record.access_code) {
    record.access_code = buildCoachAccessCode_(record.coach_name, rows.length + 1);
  }
  if (!record.token) {
    record.token = buildToken_(record.coach_name, "coach");
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

  if (!record.access_code) {
    record.access_code = buildStudentAccessCode_(record.student_name, rows.length + 1);
  }
  if (!record.token) {
    record.token = buildToken_(record.student_name, "student");
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
    updated_at: now
  };

  upsertRow_(SHEETS.programs, "program_id", record);
  replaceProgramItems_(id, items || []);
}

function publishProgram_(programId, coachId) {
  const coachPrograms = readTable_(SHEETS.programs).filter(function(row) {
    return row.coach_id === coachId;
  });

  coachPrograms.forEach(function(row) {
    updateRow_(SHEETS.programs, "program_id", row.program_id, {
      published: String(row.program_id === programId),
      updated_at: nowString_()
    });
  });
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
  const now = nowString_();
  const students = readTable_(SHEETS.students);
  const programs = readTable_(SHEETS.programs);
  const coaches = readTable_(SHEETS.coaches);
  const studentById = {};
  const programById = {};
  const coachById = {};

  students.forEach(function(row) {
    studentById[String(row.student_id || "")] = row;
  });
  programs.forEach(function(row) {
    programById[String(row.program_id || "")] = row;
  });
  coaches.forEach(function(row) {
    coachById[String(row.coach_id || "")] = row;
  });

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

    replaceMatchingWorkoutLog_(record);
  });
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

function upsertRow_(sheetName, keyField, record) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
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
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
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

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
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

function buildToken_(name, fallbackPrefix) {
  const token = String(name || fallbackPrefix)
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[^a-z0-9]/g, "");
  return token || (fallbackPrefix + "-" + new Date().getTime());
}

function createId_(prefix) {
  return prefix + "-" + Utilities.getUuid().slice(0, 8);
}

function nowString_() {
  return Utilities.formatDate(new Date(), APP_TIME_ZONE, "yyyy-MM-dd HH:mm:ss");
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

  const event = calendar.getEventById(eventId) || (
    eventId.indexOf("@google.com") === -1
      ? calendar.getEventById(eventId + "@google.com")
      : calendar.getEventById(eventId.replace("@google.com", ""))
  );
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

  const singleDelete = isSingleEventDeleteRequested_(payload || {});
  const event = singleDelete
    ? findSingleOccurrenceForDelete_(calendar, payload || {}, eventId)
    : calendar.getEventById(eventId);
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
  event.deleteEvent();
  return {
    ok: true,
    eventId: String(event.getId() || eventId),
    requestedEventId: eventId,
    calendarId: calendar.getId(),
    singleEventOnly: singleDelete,
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
    const accepted = (idMatched && looseTimeMatched) || (studentMatched && timeMatched);
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
