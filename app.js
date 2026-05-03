const STORAGE_KEY = "coachflow-v2-state";

const SESSION_KEY = "coachflow-v2-session";

const defaultState = {
  coaches: [],
  currentCoachId: "",
  students: [],
  programs: [],
  programItems: [],
  workoutLogs: []
};

const state = structuredClone(defaultState);

let editingProgramId = null;
let coachProgramEditorMode = "create";
let loadedStudentEntries = [];
let pendingSubmission = null;
let currentStudentId = "";
let currentStudentProgramId = "";
let studentViewMode = window.innerWidth <= 820 ? "card" : "table";
let studentHistoryOpened = false;
let lastSubmittedLogs = [];
let studentSubmissionCompleted = false;
let expandedProgramLibraryId = null;
let selectedTodayStudentId = "";
let programLibraryVisibleCount = 6;
let studentHistoryVisibleCount = 1;
let coachTodayVisibleCount = 6;
let coachHistoryVisibleCount = 6;
let todayStudentDetailVisibleCount = 6;
let coachViewMode = "desktop";
let coachHistoryOpened = false;
let pendingHistoryImportRows = [];
let assigningStudentId = "";
let authenticatedCoachId = "";
let authenticatedCoachAccess = "";
let currentStudentAccess = "";
let coachEditorDirty = false;
let studentEntriesDirty = false;
let loadedStudentProgramId = "";
let lastCloudSyncAt = 0;
let lastSubmittedSnapshot = null;
let studentAutoLoginPending = false;
let studentEditNoticeTimer = null;
let sessionStudentDraft = null;
let sessionStudentProgramEditDraft = null;
let coachLoginBusy = false;
let studentLoginBusy = false;
let studentLoadBusy = false;
let studentSubmitBusy = false;

function getAppMode() {
  const params = new URLSearchParams(window.location.search);
  const modeParam = params.get("mode");
  if (modeParam === "coach" || modeParam === "student" || modeParam === "admin") {
    return modeParam;
  }

  const fileName = String(window.location.pathname || "").split(/[\\/]/).pop().toLowerCase();
  if (fileName === "coach.html") {
    return "coach";
  }
  if (fileName === "admin.html") {
    return "admin";
  }
  if (fileName === "student.html") {
    return "student";
  }

  return "dual";
}

const APP_MODE = getAppMode();

const APP_CONFIG = window.APP_CONFIG || {
  mode: "local",
  appsScriptUrl: "",
  requestTimeoutMs: 12000
};
const LEAVE_SANDBOX_CONFIG = APP_CONFIG.leaveSandbox || {};
const IS_LEAVE_SANDBOX_ENABLED = LEAVE_SANDBOX_CONFIG.enabled !== false;
const LEAVE_SANDBOX_COACH_PAGE = String(LEAVE_SANDBOX_CONFIG.coachPage || "leave-coach-sandbox.html").trim();
const LEAVE_SANDBOX_STUDENT_PAGE = String(LEAVE_SANDBOX_CONFIG.studentPage || "leave-student-sandbox.html").trim();

const PUBLIC_APP_VERSION = "20260503-0016";
const APP_TIME_ZONE = "Asia/Taipei";
const LEAVE_PREFILL_STORAGE_KEY = "coachflow-leave-prefill";

const IS_CLOUD_MODE =
  String(APP_CONFIG.mode || "local").toLowerCase() === "cloud" &&
  !!String(APP_CONFIG.appsScriptUrl || "").trim();

const RETRYABLE_CLOUD_ACTIONS = new Set([
  "resolveCoachAccess",
  "resolveStudentAccess",
  "bootstrapAdmin",
  "bootstrapCoach",
  "bootstrapStudent",
  "bootstrap",
  "submitWorkoutLogs",
  "touchCoach",
  "touchStudent"
]);

function padDateNumber(value, length = 2) {
  return String(Math.max(0, Number(value) || 0)).padStart(length, "0");
}

function parseDateTimeInAppZone(rawValue) {
  if (rawValue instanceof Date && !Number.isNaN(rawValue.getTime())) {
    const parts = getDateTimePartsInAppZone(rawValue);
    return {
      ...parts,
      second: "00",
      hasTime: true
    };
  }

  const text = String(rawValue ?? "").trim();
  if (!text) {
    return null;
  }

  const literalMatch = text.match(
    /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:\s*[T ]\s*(\d{1,2})(?::(\d{1,2}))?(?::(\d{1,2}))?)?$/
  );
  if (literalMatch) {
    const [, year, month, day, hour, minute, second] = literalMatch;
    const yearNumber = Number(year);
    const monthNumber = Number(month);
    const dayNumber = Number(day);
    const hourNumber = hour === undefined ? 0 : Number(hour);
    const minuteNumber = minute === undefined ? 0 : Number(minute);
    const secondNumber = second === undefined ? 0 : Number(second);
    if (
      !Number.isFinite(yearNumber) ||
      monthNumber < 1 || monthNumber > 12 ||
      dayNumber < 1 || dayNumber > 31 ||
      hourNumber < 0 || hourNumber > 23 ||
      minuteNumber < 0 || minuteNumber > 59 ||
      secondNumber < 0 || secondNumber > 59
    ) {
      return null;
    }
    return {
      year: padDateNumber(yearNumber, 4),
      month: padDateNumber(monthNumber),
      day: padDateNumber(dayNumber),
      hour: padDateNumber(hourNumber),
      minute: padDateNumber(minuteNumber),
      second: padDateNumber(secondNumber),
      hasTime: hour !== undefined
    };
  }

  if (/^\d{10,13}$/.test(text)) {
    const epoch = Number(text);
    const millis = text.length === 13 ? epoch : epoch * 1000;
    const parsedEpoch = new Date(millis);
    if (!Number.isNaN(parsedEpoch.getTime())) {
      const parts = getDateTimePartsInAppZone(parsedEpoch);
      return {
        ...parts,
        second: "00",
        hasTime: true
      };
    }
  }

  const parsed = new Date(text);
  if (Number.isNaN(parsed.getTime())) {
    return null;
  }

  const parts = getDateTimePartsInAppZone(parsed);
  return {
    ...parts,
    second: "00",
    hasTime: true
  };
}

function normalizeDateKey(rawValue) {
  const parsed = parseDateTimeInAppZone(rawValue);
  if (!parsed) {
    return "";
  }
  return `${parsed.year}-${parsed.month}-${parsed.day}`;
}

function normalizeDateTimeValue(rawValue) {
  const parsed = parseDateTimeInAppZone(rawValue);
  if (!parsed) {
    return "";
  }
  return `${parsed.year}-${parsed.month}-${parsed.day} ${parsed.hour}:${parsed.minute}`;
}

function formatDateDisplay(rawValue, fallback = "-") {
  const parsed = parseDateTimeInAppZone(rawValue);
  if (!parsed) {
    const text = String(rawValue ?? "").trim();
    return text || fallback;
  }
  return `${parsed.year}-${parsed.month}-${parsed.day}`;
}

function formatDateTimeDisplay(rawValue, fallback = "-") {
  const parsed = parseDateTimeInAppZone(rawValue);
  if (!parsed) {
    const text = String(rawValue ?? "").trim();
    return text || fallback;
  }
  return `${parsed.year}-${parsed.month}-${parsed.day} ${parsed.hour}:${parsed.minute}`;
}

function formatRocDateDisplay(rawValue, fallback = "尚無紀錄") {
  const parsed = parseDateTimeInAppZone(rawValue);
  if (!parsed) {
    const text = String(rawValue ?? "").trim();
    return text || fallback;
  }
  return `${Number(parsed.year) - 1911}.${Number(parsed.month)}.${Number(parsed.day)}`;
}

function formatRocDateTimeDisplay(rawValue, fallback = "尚無紀錄") {
  const parsed = parseDateTimeInAppZone(rawValue);
  if (!parsed) {
    const text = String(rawValue ?? "").trim();
    return text || fallback;
  }
  return `${Number(parsed.year) - 1911}.${Number(parsed.month)}.${Number(parsed.day)} ${parsed.hour}:${parsed.minute}`;
}

function getComparableDateKey(rawValue) {
  const parsed = parseDateTimeInAppZone(rawValue);
  if (!parsed) {
    return "";
  }
  return `${parsed.year}-${parsed.month}-${parsed.day}`;
}

function getComparableDateTimeKey(rawValue) {
  const parsed = parseDateTimeInAppZone(rawValue);
  if (!parsed) {
    return "";
  }
  return `${parsed.year}-${parsed.month}-${parsed.day} ${parsed.hour}:${parsed.minute}:${parsed.second}`;
}

function getLogDateTimeSortKey(log) {
  const submitted = getComparableDateTimeKey(log?.submittedAt);
  if (submitted) {
    return submitted;
  }
  const updated = getComparableDateTimeKey(log?.updatedAt);
  if (updated) {
    return updated;
  }
  const date = getComparableDateKey(log?.programDate);
  return date ? `${date} 00:00:00` : "";
}

function getLogActivityDateKey(log) {
  return getComparableDateKey(log?.submittedAt) || getComparableDateKey(log?.programDate);
}

function compareLogsByDateTimeDesc(a, b) {
  return getLogDateTimeSortKey(b).localeCompare(getLogDateTimeSortKey(a));
}

function compareLogsByDateTimeAsc(a, b) {
  return getLogDateTimeSortKey(a).localeCompare(getLogDateTimeSortKey(b));
}

function normalizeCloudCoach(row, index = 0) {
  const name = row.coach_name || row.name || `Coach ${index + 1}`;
  return {
    id: row.coach_id || row.id || `coach-${index + 1}`,
    name,
    status: row.status || "active",
    role: row.role || "coach",
    accessCode: row.access_code || row.accessCode || buildCoachAccessCode(name, index + 1),
    token: row.token || "",
    lastUsedAt: normalizeDateTimeValue(row.last_used_at || row.lastUsedAt || "")
  };
}

function isUsableCoach(record) {
  return !!String(record?.id || "").trim() && !!String(record?.name || "").trim();
}

function normalizeCloudStudent(row, index = 0) {
  const name = row.student_name || row.name || `學生 ${index + 1}`;
  return {
    id: row.student_id || row.id || `stu-${index + 1}`,
    name,
    className: row.class_name || row.className || "",
    status: row.status || "active",
    accessCode: row.access_code || row.accessCode || buildAccessCode(name, index + 1),
    token: row.token || "",
    primaryCoachId: row.primary_coach_id || row.primaryCoachId || "",
    primaryCoachName: row.primary_coach_name || row.primaryCoachName || "",
    lastUsedAt: normalizeDateTimeValue(row.last_used_at || row.lastUsedAt || "")
  };
}

function isUsableStudent(record) {
  return !!String(record?.id || "").trim() && !!String(record?.name || "").trim();
}

function normalizeCloudProgram(row, index = 0) {
  return {
    id: row.program_id || row.id || `program-${index + 1}`,
    code: row.program_code || row.code || "",
    date: normalizeDateKey(row.program_date || row.date || ""),
    title: row.title || "",
    coachId: row.coach_id || row.coachId || "",
    coachName: row.coach_name || row.coachName || "",
    notes: row.notes || "",
    published: String(row.published || "").toLowerCase() === "true" || row.published === true,
    targetStudentIds: normalizeProgramStudentIds(row.target_student_ids || row.targetStudentIds || row.student_ids || row.studentIds || ""),
    createdAt: normalizeDateTimeValue(row.created_at || row.createdAt || ""),
    updatedAt: normalizeDateTimeValue(row.updated_at || row.updatedAt || "")
  };
}

function isUsableProgram(record) {
  return !!String(record?.id || "").trim() && !!String(record?.code || "").trim();
}

function normalizeProgramStudentIds(value = []) {
  const rawIds = Array.isArray(value)
    ? value
    : String(value || "")
        .split(/[,，\n]/);
  return [...new Set(
    rawIds
      .map((id) => String(id || "").trim())
      .filter(Boolean)
  )];
}

function getProgramStudentIds(program = {}) {
  return normalizeProgramStudentIds(program.targetStudentIds || program.studentIds || "");
}

function getProgramTargetStudentNames(program = {}) {
  const targetIds = getProgramStudentIds(program);
  if (!targetIds.length) {
    return [];
  }
  return targetIds.map((studentId) => {
    const student = state.students.find((item) => item.id === studentId);
    return student?.name || studentId;
  });
}

function isProgramTargeted(program = {}) {
  return getProgramStudentIds(program).length > 0;
}

function isProgramAvailableForStudent(program = {}, studentId = "") {
  const targetIds = getProgramStudentIds(program);
  return !targetIds.length || targetIds.includes(studentId);
}

function normalizeCloudProgramItem(row, index = 0) {
  return {
    id: row.item_id || row.id || `item-${index + 1}`,
    programId: row.program_id || row.programId || "",
    sortOrder: Number(row.sort_order || row.sortOrder || index + 1),
    category: row.category || "",
    exercise: row.exercise || "",
    targetSets: Number(row.target_sets || row.targetSets || 0),
    targetType: row.target_type || row.targetType || "reps",
    targetValue: Number(row.target_value || row.targetValue || 0),
    itemNote: row.item_note || row.itemNote || "",
    updatedAt: normalizeDateTimeValue(row.updated_at || row.updatedAt || "")
  };
}

function isUsableProgramItem(record) {
  return !!String(record?.id || "").trim()
    && !!String(record?.programId || "").trim()
    && !!String(record?.exercise || "").trim();
}

function normalizeCloudLog(row, index = 0) {
  return {
    id: row.log_id || row.id || `log-${index + 1}`,
    programId: row.program_id || row.programId || "",
    programCode: row.program_code || row.programCode || "",
    programDate: normalizeDateKey(row.program_date || row.programDate || ""),
    coachId: row.coach_id || row.coachId || "",
    coachName: row.coach_name || row.coachName || "",
    studentId: row.student_id || row.studentId || "",
    studentName: row.student_name || row.studentName || "",
    category: row.category || "",
    exercise: row.exercise || "",
    targetSets: Number(row.target_sets || row.targetSets || 0),
    targetType: row.target_type || row.targetType || "reps",
    targetValue: Number(row.target_value || row.targetValue || 0),
    actualWeight: Number(row.actual_weight || row.actualWeight || 0),
    actualSets: Number(row.actual_sets || row.actualSets || 0),
    actualReps: Number(row.actual_reps || row.actualReps || 0),
    studentNote: row.student_note || row.studentNote || "",
    submittedAt: normalizeDateTimeValue(row.submitted_at || row.submittedAt || ""),
    updatedAt: normalizeDateTimeValue(row.updated_at || row.updatedAt || "")
  };
}

function isUsableWorkoutLog(record) {
  return !!String(record?.id || "").trim()
    && !!String(record?.studentId || "").trim()
    && !!String(record?.exercise || "").trim();
}

function applyCloudPayloadToState(payload) {
  if (!payload || payload.ok === false) {
    throw new Error(payload?.message || "雲端資料回傳失敗。");
  }

  if (Array.isArray(payload.coaches)) {
    state.coaches = payload.coaches
      .map(normalizeCloudCoach)
      .filter(isUsableCoach);
  }
  if (Array.isArray(payload.students)) {
    state.students = payload.students
      .map(normalizeCloudStudent)
      .filter(isUsableStudent);
  }
  if (Array.isArray(payload.programs)) {
    state.programs = payload.programs
      .map(normalizeCloudProgram)
      .filter(isUsableProgram);
  }
  if (Array.isArray(payload.programItems)) {
    state.programItems = payload.programItems
      .map(normalizeCloudProgramItem)
      .filter(isUsableProgramItem);
  }
  if (Array.isArray(payload.workoutLogs)) {
    state.workoutLogs = payload.workoutLogs
      .map(normalizeCloudLog)
      .filter(isUsableWorkoutLog);
  }

  if (payload.coach) {
    const normalizedCoach = normalizeCloudCoach(payload.coach);
    const existingCoach = state.coaches.find((coach) => coach.id === normalizedCoach.id);
    if (existingCoach) {
      Object.assign(existingCoach, normalizedCoach);
    } else {
      state.coaches.push(normalizedCoach);
    }
  }

  if (payload.student) {
    const normalizedStudent = normalizeCloudStudent(payload.student);
    const existingStudent = state.students.find((student) => student.id === normalizedStudent.id);
    if (existingStudent) {
      Object.assign(existingStudent, normalizedStudent);
    } else {
      state.students.push(normalizedStudent);
    }
  }

  persistState();
}

async function callCloudApi(action, payload = {}, method = "POST") {
  if (!IS_CLOUD_MODE) {
    throw new Error("目前不是雲端模式。");
  }

  const controller = new AbortController();
  const timeout = window.setTimeout(
    () => controller.abort(),
    Number(APP_CONFIG.requestTimeoutMs || 12000)
  );

  try {
    const url = new URL(APP_CONFIG.appsScriptUrl);

    let response;
    if (String(method).toUpperCase() === "GET") {
      url.searchParams.set("action", action);
      url.searchParams.set("_ts", String(Date.now()));
      Object.entries(payload || {}).forEach(([key, value]) => {
        if (value !== undefined && value !== null) {
          url.searchParams.set(key, String(value));
        }
      });
      response = await fetch(url.toString(), {
        method: "GET",
        cache: "no-store",
        signal: controller.signal
      });
    } else {
      response = await fetch(url.toString(), {
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
    if (!response.ok || json.ok === false) {
      throw new Error(json?.message || "雲端 API 執行失敗。");
    }
    return json;
  } catch (error) {
    if (error?.name === "AbortError") {
      throw new Error("雲端請求逾時，請稍後再試。");
    }
    throw error;
  } finally {
    window.clearTimeout(timeout);
  }
}

function shouldRetryCloudError(error) {
  const message = String(error?.message || "");
  if (!message) {
    return false;
  }

  if (/找不到|停用|Unsupported action|目前不是雲端模式/i.test(message)) {
    return false;
  }

  return /逾時|timeout|network|failed to fetch|load failed|unexpected end|cloud|request/i.test(message);
}

async function callCloudApiCritical(action, payload = {}, method = "POST", retryCount = 2) {
  const canRetry = RETRYABLE_CLOUD_ACTIONS.has(String(action || ""));
  const attempts = canRetry ? Math.max(0, Number(retryCount || 0)) + 1 : 1;
  let lastError = null;

  for (let index = 0; index < attempts; index += 1) {
    try {
      return await callCloudApi(action, payload, method);
    } catch (error) {
      lastError = error;
      const shouldRetry = canRetry && index < attempts - 1 && shouldRetryCloudError(error);
      if (!shouldRetry) {
        throw error;
      }
      await delay(400 * (index + 1));
    }
  }

  throw lastError || new Error("雲端連線失敗。");
}

async function bootstrapCloudMode() {
  if (!IS_CLOUD_MODE) {
    return;
  }

  if (APP_MODE === "admin") {
    try {
      const payload = await callCloudApiCritical("bootstrapAdmin", {}, "GET", 3);
      applyCloudPayloadToState(payload);
    } catch (error) {
      console.warn("Admin cloud bootstrap failed, using local state:", error);
    }
    return;
  }

  if (APP_MODE === "coach") {
    const urlAccess = getModeAccessFromUrl("coach");
    if (urlAccess) {
      const previousCoachId = authenticatedCoachId;
      const previousCoachAccess = authenticatedCoachAccess;
      authenticatedCoachAccess = normalizeAccessInput(urlAccess);
      if (els.coachAccessCode) {
        els.coachAccessCode.value = authenticatedCoachAccess;
      }
      try {
        const matchedCoach = await resolveCoachAccessFromCloud(authenticatedCoachAccess);
        if (matchedCoach?.id) {
          authenticatedCoachId = matchedCoach.id;
          authenticatedCoachAccess = matchedCoach.accessCode || authenticatedCoachAccess;
          state.currentCoachId = matchedCoach.id;
          if (els.coachAccessCode) {
            els.coachAccessCode.value = matchedCoach.accessCode || authenticatedCoachAccess;
          }
          const payload = await callCloudApiCritical("bootstrapCoach", { coachId: matchedCoach.id }, "POST", 3);
          applyCloudPayloadToState(payload);
        } else if (!previousCoachId) {
          authenticatedCoachId = "";
        } else {
          authenticatedCoachId = previousCoachId;
          authenticatedCoachAccess = previousCoachAccess || authenticatedCoachAccess;
        }
      } catch (error) {
        console.warn("Coach cloud bootstrap by URL access failed:", error);
        authenticatedCoachId = previousCoachId;
        authenticatedCoachAccess = previousCoachAccess || authenticatedCoachAccess;
      }
      persistSession();
      return;
    }
    if (authenticatedCoachId) {
      try {
        const payload = await callCloudApiCritical("bootstrapCoach", { coachId: authenticatedCoachId }, "POST", 2);
        applyCloudPayloadToState(payload);
        return;
      } catch (error) {
        console.warn("Coach cloud bootstrap by session failed, falling back to URL:", error);
      }
    }
    if (authenticatedCoachAccess) {
      try {
        const matchedCoach = await resolveCoachAccessFromCloud(authenticatedCoachAccess);
        if (matchedCoach?.id) {
          authenticatedCoachId = matchedCoach.id;
          state.currentCoachId = matchedCoach.id;
          if (els.coachAccessCode) {
            els.coachAccessCode.value = matchedCoach.accessCode || authenticatedCoachAccess;
          }
          const payload = await callCloudApiCritical("bootstrapCoach", { coachId: matchedCoach.id }, "POST", 2);
          applyCloudPayloadToState(payload);
          persistSession();
          return;
        }
      } catch (error) {
        console.warn("Coach cloud bootstrap by access failed, falling back to URL:", error);
      }
    }
    await hydrateCoachAccessFromUrl();
    return;
  }

  if (APP_MODE === "student") {
    const urlAccess = getModeAccessFromUrl("student");
    if (urlAccess) {
      const previousStudentId = currentStudentId;
      const previousStudentAccess = currentStudentAccess;
      currentStudentAccess = normalizeAccessInput(urlAccess);
      if (els.studentAccessCode) {
        els.studentAccessCode.value = currentStudentAccess;
      }
      try {
        const matchedStudent = await resolveStudentAccessFromCloud(currentStudentAccess);
        if (matchedStudent?.id) {
          currentStudentId = matchedStudent.id;
          currentStudentAccess = matchedStudent.accessCode || currentStudentAccess;
          if (els.studentAccessCode) {
            els.studentAccessCode.value = matchedStudent.accessCode || currentStudentAccess;
          }
          const payload = await callCloudApiCritical("bootstrapStudent", { studentId: matchedStudent.id }, "POST", 3);
          applyCloudPayloadToState(payload);
        } else if (!previousStudentId) {
          currentStudentId = "";
        } else {
          currentStudentId = previousStudentId;
          currentStudentAccess = previousStudentAccess || currentStudentAccess;
        }
      } catch (error) {
        console.warn("Student cloud bootstrap by URL access failed:", error);
        currentStudentId = previousStudentId;
        currentStudentAccess = previousStudentAccess || currentStudentAccess;
      }
      persistSession();
      return;
    }
    if (currentStudentId) {
      try {
        const payload = await callCloudApiCritical("bootstrapStudent", { studentId: currentStudentId }, "POST", 2);
        applyCloudPayloadToState(payload);
        return;
      } catch (error) {
        console.warn("Student cloud bootstrap by session failed, falling back to URL:", error);
      }
    }
    if (currentStudentAccess) {
      try {
        const matchedStudent = await resolveStudentAccessFromCloud(currentStudentAccess);
        if (matchedStudent?.id) {
          currentStudentId = matchedStudent.id;
          if (els.studentAccessCode) {
            els.studentAccessCode.value = matchedStudent.accessCode || currentStudentAccess;
          }
          const payload = await callCloudApiCritical("bootstrapStudent", { studentId: matchedStudent.id }, "POST", 2);
          applyCloudPayloadToState(payload);
          persistSession();
          return;
        }
      } catch (error) {
        console.warn("Student cloud bootstrap by access failed, falling back to URL:", error);
      }
    }
    await hydrateStudentAccessFromUrl();
  }
}

function getModeAccessFromUrl(mode) {
  const params = new URLSearchParams(window.location.search);
  if (mode === "coach") {
    return normalizeAccessInput(params.get("coach") || params.get("token") || params.get("code") || "");
  }
  if (mode === "student") {
    return normalizeAccessInput(params.get("student") || params.get("token") || params.get("code") || "");
  }
  return "";
}

async function resolveStudentAccessFromCloud(accessValue) {
  const candidates = buildAccessCandidates(accessValue);
  let lastError = null;

  for (const candidate of candidates) {
    try {
      const payload = await callCloudApiCritical("resolveStudentAccess", { access: candidate }, "POST", 3);
      applyCloudPayloadToState(payload);
      const matched = payload.student ? normalizeCloudStudent(payload.student) : resolveStudentByAccess(candidate);
      if (matched) {
        return matched;
      }
    } catch (error) {
      lastError = error;
    }
  }

  if (lastError) {
    throw lastError;
  }
  return null;
}

async function resolveCoachAccessFromCloud(accessValue) {
  const candidates = buildAccessCandidates(accessValue);
  let lastError = null;

  for (const candidate of candidates) {
    try {
      const payload = await callCloudApiCritical("resolveCoachAccess", { access: candidate }, "POST", 3);
      applyCloudPayloadToState(payload);
      const matched = payload.coach ? normalizeCloudCoach(payload.coach) : resolveCoachByAccess(candidate);
      if (matched) {
        return matched;
      }
    } catch (error) {
      lastError = error;
    }
  }

  if (lastError) {
    throw lastError;
  }
  return null;
}

async function submitStudentLogsToCloud(logs) {
  const payload = await callCloudApiCritical("submitWorkoutLogs", {
    logs: logs.map((log) => ({
      log_id: log.id,
      program_id: log.programId,
      program_code: log.programCode,
      program_date: log.programDate,
      coach_id: log.coachId || "",
      coach_name: log.coachName || "",
      student_id: log.studentId,
      student_name: log.studentName,
      category: log.category,
      exercise: log.exercise,
      target_sets: log.targetSets,
      target_type: log.targetType,
      target_value: log.targetValue,
      actual_weight: log.actualWeight,
      actual_sets: log.actualSets,
      actual_reps: log.actualReps,
      student_note: log.studentNote,
      submitted_at: log.submittedAt
    }))
  }, "POST", 2);
  applyCloudPayloadToState(payload);
}

async function saveProgramToCloud(program, items) {
  const payload = await callCloudApi("saveProgram", {
    program: {
      program_id: program.id,
      program_code: program.code,
      program_date: program.date,
      title: program.title || "",
      coach_id: program.coachId || "",
      coach_name: program.coachName || "",
      notes: program.notes || "",
      published: String(!!program.published),
      target_student_ids: getProgramStudentIds(program).join(",")
    },
    items: items.map((item) => ({
      item_id: item.id,
      program_id: program.id,
      sort_order: item.sortOrder,
      category: item.category,
      exercise: item.exercise,
      target_sets: item.targetSets,
      target_type: item.targetType,
      target_value: item.targetValue,
      item_note: item.itemNote || ""
    }))
  });
  applyCloudPayloadToState(payload);
}

async function publishProgramToCloud(program, items) {
  await saveProgramToCloud({ ...program, published: true }, items);
  const payload = await callCloudApi("publishProgram", {
    programId: program.id,
    coachId: program.coachId
  });
  applyCloudPayloadToState(payload);
}

async function deleteProgramFromCloud(programId) {
  const payload = await callCloudApi("deleteProgram", { programId, program_id: programId });
  if (Array.isArray(payload.programs) && payload.programs.some((program) => (program.program_id || program.id) === programId)) {
    throw new Error("雲端尚未刪除這份課表，請確認 Apps Script 已重新部署。");
  }
  applyCloudPayloadToState(payload);
}

async function importHistoryRowsToCloud(rows) {
  const payload = await callCloudApi("importWorkoutLogs", {
    logs: rows.map((row) => {
      const student = state.students.find((item) => item.id === row.studentId);
      const coach = state.coaches.find((item) => item.id === (student?.primaryCoachId || ""));
      return {
        log_id: row.logId || "",
        program_id: "",
        program_code: row.programCode || "",
        program_date: row.programDate,
        coach_id: coach?.id || "",
        coach_name: coach?.name || "",
        student_id: row.studentId,
        student_name: row.studentName,
        category: row.category,
        exercise: row.exercise,
        target_sets: row.targetSets,
        target_type: row.targetType,
        target_value: row.targetValue,
        actual_weight: row.actualWeight,
        actual_sets: row.actualSets,
        actual_reps: row.actualReps,
        student_note: row.studentNote || ""
      };
    })
  });
  applyCloudPayloadToState(payload);
}

function setNodeText(node, text) {
  if (node) {
    node.textContent = text;
  }
}

function setLabelText(label, text) {
  if (!label) {
    return;
  }
  const textNode = [...label.childNodes].find((node) => node.nodeType === Node.TEXT_NODE);
  if (textNode) {
    textNode.textContent = `${text}\n`;
    return;
  }
  label.prepend(document.createTextNode(`${text}\n`));
}

function setTableHeaders(tableSelector, labels) {
  const headers = document.querySelectorAll(`${tableSelector} thead th`);
  labels.forEach((label, index) => {
    if (headers[index]) {
      headers[index].textContent = label;
    }
  });
}

function applyStaticCopy() {
  document.title = "CoachFlow 重訓課表與紀錄系統";

  setNodeText(document.querySelector(".hero-copy .eyebrow"), "CoachFlow 重訓課表與紀錄系統");
  setNodeText(document.querySelector(".hero-copy h1"), "教練建立課表，學生完成紀錄");
  setNodeText(document.querySelector(".hero-note .hero-chip"), "雲端連線");
  setNodeText(document.querySelector(".hero-note h3"), "Google Sheets + Apps Script");
  setNodeText(document.querySelector(".hero-note p"), "首次登入可能短暫等待。");

  if (els.mainTabs[0]) setNodeText(els.mainTabs[0], "教練管理");
  if (els.mainTabs[1]) setNodeText(els.mainTabs[1], "學生填寫");

  if (els.coachGlobalViewDesktopButton) setNodeText(els.coachGlobalViewDesktopButton, "網頁版介面");
  if (els.coachGlobalViewMobileButton) setNodeText(els.coachGlobalViewMobileButton, "手機版介面");
  if (els.studentGlobalViewDesktopButton) setNodeText(els.studentGlobalViewDesktopButton, "網頁版介面");
  if (els.studentGlobalViewMobileButton) setNodeText(els.studentGlobalViewMobileButton, "手機版介面");

  const coachTabLabels = {
    "coach-editor": "建立課表",
    "coach-coaches": "教練管理",
    "coach-students": "學生管理",
    "coach-today": "今日紀錄",
    "coach-history": "歷史查詢"
  };
  els.coachTabs.forEach((button) => {
    const label = coachTabLabels[button.dataset.coachTab];
    if (label) {
      button.textContent = label;
    }
  });
  if (els.openCoachLeaveSystemTab) setNodeText(els.openCoachLeaveSystemTab, "請假系統");

  setNodeText(document.querySelector("#coach-auth-card .section-kicker"), "教練登入");
  setNodeText(document.querySelector("#coach-auth-card h3"), "教練專屬入口");
  if (els.coachAccessCode) {
    els.coachAccessCode.placeholder = "例如：CL001";
    setLabelText(els.coachAccessCode.closest("label"), "教練代碼");
  }
  setNodeText(document.querySelector("#coach-auth-card .muted-copy"), "也可以直接使用教練專屬連結進入，系統會自動帶入你的身份。");
  if (els.confirmCoachAccess) setNodeText(els.confirmCoachAccess, "確認身份並進入教練工作台");

  setNodeText(document.querySelector("#student-entry-card .section-kicker"), "身份確認");
  setNodeText(document.querySelector("#student-entry-card h3"), "學生專屬入口");
  if (els.studentAccessCode) {
    els.studentAccessCode.placeholder = "例如：WM001";
    setLabelText(els.studentAccessCode.closest("label"), "專屬代碼");
  }
  const studentEntryNote = document.querySelector("#student-entry-card .muted-copy");
  if (studentEntryNote) {
    studentEntryNote.textContent = "也可以直接使用專屬連結進入，系統會自動帶入你的身份。";
  }
  if (els.confirmStudentAccess) setNodeText(els.confirmStudentAccess, "確認身份並載入本輪課表");

  setNodeText(document.querySelector("#coach-coaches .section-kicker"), "教練設定");
  setNodeText(document.querySelector("#coach-coaches h3"), "教練管理");
  setNodeText(document.querySelector("#coach-roster-card > .muted-copy"), "可新增教練、修改名稱、停用教練，並查看每位教練的最後使用狀態。");
  setNodeText(document.querySelector("#coach-roster-card .student-manager-block:first-child .student-manager-label"), "新增教練");
  if (els.newCoachName) {
    els.newCoachName.placeholder = "例如：王教練";
    setLabelText(els.newCoachName.closest("label"), "教練姓名");
  }
  if (els.addCoachButton) {
    els.addCoachButton.textContent = "新增教練";
  }
  setNodeText(document.querySelector("#coach-roster-card .student-manager-block:last-child .student-manager-label"), "目前登入");
  const coachSessionBlockCopy = document.querySelector("#coach-roster-card .student-manager-block:last-child .muted-copy");
  if (coachSessionBlockCopy) {
    coachSessionBlockCopy.textContent = "正式版將由教練專屬入口登入，管理者頁不再切換目前教練。";
  }
  if (els.coachSessionName) {
    els.coachSessionName.placeholder = "管理者模式不使用";
  }
  setNodeText(document.querySelector("#coach-students .section-kicker"), "學生管理");
  setNodeText(document.querySelector("#coach-students h3"), "學生管理");
  setNodeText(document.querySelector("#coach-student-links-card > .muted-copy"), "可依學生姓名快速查看專屬網址與共用入口，方便傳給學生登入。");
  if (els.coachStudentLinkName) {
    els.coachStudentLinkName.placeholder = "請輸入學生姓名";
    setLabelText(els.coachStudentLinkName.closest("label"), "學生姓名");
  }
  if (els.coachStudentLinkSelect) {
    setLabelText(els.coachStudentLinkSelect.closest("label"), "或直接選擇學生");
    if (els.coachStudentLinkSelect.options.length && !els.coachStudentLinkSelect.options[0].value) {
      els.coachStudentLinkSelect.options[0].textContent = "請選擇學生";
    }
  }
  setNodeText(document.querySelector("#coach-student-links-card .section-kicker"), "學生專屬網址");
  setNodeText(document.querySelector("#coach-student-links-card h3"), "學生專屬網址");
  setNodeText(document.querySelector("#history-import-card .section-kicker"), "歷史資料匯入");
  setNodeText(document.querySelector("#history-import-card h3"), "批次匯入訓練紀錄");
  setNodeText(document.querySelector("#history-import-card > .muted-copy"), "可下載匯入範本 CSV，整理完成後再批次匯入舊訓練紀錄。");
  setNodeText(document.querySelector("#history-import-card .student-manager-block:first-child .student-manager-label"), "步驟 1");
  if (els.downloadHistoryTemplateButton) {
    els.downloadHistoryTemplateButton.textContent = "下載匯入範本";
  }
  setNodeText(document.querySelector("#history-import-card .student-manager-block:last-child .student-manager-label"), "步驟 2");
  if (els.historyImportFile) {
    setLabelText(els.historyImportFile.closest("label"), "上傳 CSV 檔案");
  }
  if (els.confirmHistoryImportButton) {
    els.confirmHistoryImportButton.textContent = "確認匯入";
  }
  setNodeText(document.querySelector("#coach-editor .section-kicker"), "課表編輯");
  setNodeText(document.querySelector("#coach-editor h3"), "建立本輪課表");
  if (els.coachActiveCopy) {
    els.coachActiveCopy.textContent = "";
  }
  if (els.changeCoachAccess) {
    els.changeCoachAccess.textContent = "切換教練";
  }
  if (els.programCode) setLabelText(els.programCode.closest("label"), "課表代碼");
  if (els.programDate) setLabelText(els.programDate.closest("label"), "課表日期");
  if (els.coachName) {
    setLabelText(els.coachName.closest("label"), "教練名稱");
    els.coachName.placeholder = "例如：Coach Lin";
  }
  if (els.programNotes) {
    setLabelText(els.programNotes.closest("label"), "課表備註");
    els.programNotes.placeholder = "可填寫本輪提醒或訓練重點";
  }
  const addProgramItemButton = document.querySelector("#add-program-item");
  if (addProgramItemButton) addProgramItemButton.textContent = "新增一列";
  const saveProgramButton = document.querySelector("#save-program");
  if (saveProgramButton) saveProgramButton.textContent = "儲存課表";
  const publishTargetedProgramButton = document.querySelector("#publish-targeted-program");
  if (publishTargetedProgramButton) publishTargetedProgramButton.textContent = "發布給指定學生";
  const publishProgramButton = document.querySelector("#publish-program");
  if (publishProgramButton) publishProgramButton.textContent = "發布全班共用課表";
  const coachTodayTab = document.querySelector('.sub-tab[data-coach-tab="coach-today"]');
  if (coachTodayTab) coachTodayTab.textContent = "本週紀錄";
  setNodeText(document.querySelector("#coach-today .section-kicker"), "提交狀況");
  setNodeText(document.querySelector("#coach-today h3"), "本輪全班紀錄");
  if (els.coachTodayDate) setLabelText(els.coachTodayDate.closest("label"), "查看本輪日期");
  if (els.todayLogsSearch) {
    setLabelText(els.todayLogsSearch.closest("label"), "搜尋本輪紀錄");
    els.todayLogsSearch.placeholder = "可搜尋學生姓名或動作";
  }
  setNodeText(document.querySelector("#coach-history .section-kicker"), "歷史查詢");
  setNodeText(document.querySelector("#coach-history h3"), "學生／動作歷史紀錄");
  if (els.coachHistoryStudent) setLabelText(els.coachHistoryStudent.closest("label"), "學生");
  if (els.coachHistoryExercise) {
    setLabelText(els.coachHistoryExercise.closest("label"), "動作");
    els.coachHistoryExercise.placeholder = "請輸入或選擇動作";
  }
  if (els.coachHistoryDateFrom) setLabelText(els.coachHistoryDateFrom.closest("label"), "開始日期");
  if (els.coachHistoryDateTo) setLabelText(els.coachHistoryDateTo.closest("label"), "結束日期");
  if (els.coachHistoryLast30) els.coachHistoryLast30.textContent = "最近 30 天";
  if (els.runCoachHistoryButton) els.runCoachHistoryButton.textContent = "查看歷史紀錄";
  if (els.resetCoachHistoryButton) els.resetCoachHistoryButton.textContent = "清除條件";
  if (els.exportCoachHistoryButton) els.exportCoachHistoryButton.textContent = "匯出 CSV";

  setTableHeaders(".editor-table", ["分類", "動作名稱", "組數", "類型", "次數", "備註", "排序", ""]);
  setTableHeaders("#coach-today table", ["學生", "分類", "動作", "目標", "實際結果", "備註", "送出時間"]);
  setTableHeaders("#today-student-detail-card table", ["分類", "動作", "目標", "實際結果", "備註", "送出時間"]);
  setTableHeaders("#coach-history .wide-card table", ["日期", "學生", "分類", "動作", "目標", "實際結果", "備註"]);
  setTableHeaders("#student-table-wrap table", ["分類", "動作", "目標", "填寫內容"]);
  setTableHeaders("#student-history-card table", ["日期", "分類", "動作", "目標", "實際結果", "備註"]);
  setTableHeaders("#success-modal table", ["分類", "動作", "目標", "實際結果", "備註"]);
  setTableHeaders("#history-import-card table", ["學生姓名", "日期", "動作", "分類", "目標格式", "匯入狀態"]);

  if (els.loadStudentProgramInline) els.loadStudentProgramInline.textContent = "載入本輪課表";
  if (els.loadStudentProgramMobile) els.loadStudentProgramMobile.textContent = "載入本輪課表";
  if (els.editStudentProgramMobileTop) els.editStudentProgramMobileTop.textContent = "編輯課表項目";
  if (els.editStudentProgramMobile) els.editStudentProgramMobile.textContent = "編輯課表項目";
  if (els.openStudentHistoryMobile) els.openStudentHistoryMobile.textContent = "查看我的歷史紀錄";
  if (els.submitStudentLog) els.submitStudentLog.textContent = "確認本次填寫內容";
  if (els.submitStudentLogMobile) els.submitStudentLogMobile.textContent = "確認本次填寫內容";
  if (els.openStudentHistory) els.openStudentHistory.textContent = "查看我的歷史紀錄";
  if (els.downloadSuccessImage) els.downloadSuccessImage.textContent = "下載圖片";
  if (els.successOpenHistory) els.successOpenHistory.textContent = "返回歷史紀錄";
  if (els.closeSuccessModal) els.closeSuccessModal.textContent = "完成";
  setNodeText(document.querySelector("#success-modal .section-kicker"), "完成送出");
  setNodeText(document.querySelector("#success-title"), "課表已完成");
  setNodeText(document.querySelector("#student-history-card .section-kicker"), "歷史紀錄");
  setNodeText(document.querySelector("#student-history-card h3"), "我的歷史紀錄");
  if (els.studentHistoryExercise) {
    setLabelText(els.studentHistoryExercise.closest("label"), "項目");
    els.studentHistoryExercise.placeholder = "請選擇或輸入項目";
  }

  const studentRosterHint = document.querySelector("#coach-students .student-roster-heading > .muted-copy");
  if (studentRosterHint) {
    studentRosterHint.textContent = APP_MODE === "admin"
      ? "可指派教練、修改姓名、複製代碼、停用或刪除學生。"
      : "可直接查看網址、修改姓名、複製代碼、查看全部紀錄、停用或刪除學生。";
  }
}

function formatUsageTimestamp(value) {
  return formatRocDateTimeDisplay(value, "尚無紀錄");
}

function markCoachUsed(coachId) {
  if (!coachId) {
    return;
  }
  const coach = state.coaches.find((item) => item.id === coachId);
  if (!coach) {
    return;
  }
  coach.lastUsedAt = timestampNow();
  persistState();
}

function markStudentUsed(studentId) {
  if (!studentId) {
    return;
  }
  const student = state.students.find((item) => item.id === studentId);
  if (!student) {
    return;
  }
  student.lastUsedAt = timestampNow();
  persistState();
}

function buildCoachAccessCode(name, index) {
  const latin = String(name || "")
    .replace(/[^A-Za-z]/g, "")
    .toUpperCase()
    .slice(0, 2);
  const prefix = latin || "CO";
  return `${prefix}${String(index).padStart(3, "0")}`;
}

function normalizeAccessInput(rawValue) {
  return String(rawValue || "")
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .trim();
}

function buildAccessCandidates(rawValue) {
  const normalized = normalizeAccessInput(rawValue);
  if (!normalized) {
    return [];
  }
  return [...new Set([normalized, normalized.toUpperCase(), normalized.toLowerCase()])];
}

function resolveCoachByAccess(rawValue) {
  const value = normalizeAccessInput(rawValue).toLowerCase();
  if (!value) {
    return null;
  }
  return state.coaches.find((coach) => {
    const code = String(coach.accessCode || "").trim().toLowerCase();
    const token = String(coach.token || "").trim().toLowerCase();
    return coach.status !== "inactive" && (value === code || value === token || value === String(coach.id).toLowerCase());
  }) || null;
}

function getCurrentCoach() {
  if (APP_MODE === "coach") {
    return state.coaches?.find((coach) => coach.id === authenticatedCoachId)
      || null;
  }
  return state.coaches?.find((coach) => coach.id === state.currentCoachId)
    || state.coaches?.[0]
    || null;
}

function getCoachScopedStudents() {
  const coach = getCurrentCoach();
  if (!coach) {
    return [];
  }
  if (APP_MODE === "admin") {
    return state.students;
  }
  return state.students.filter((student) => (student.primaryCoachId || coach.id) === coach.id);
}

function getCoachScopedPrograms() {
  const coach = getCurrentCoach();
  if (!coach) {
    return [];
  }
  if (APP_MODE === "admin") {
    return state.programs;
  }
  return state.programs.filter((program) => (program.coachId || coach.id) === coach.id);
}

function getCoachScopedLogs() {
  const coach = getCurrentCoach();
  if (!coach) {
    return [];
  }
  if (APP_MODE === "admin") {
    return state.workoutLogs;
  }
  return state.workoutLogs.filter((log) => {
    const program = state.programs.find((item) => item.id === log.programId);
    const student = state.students.find((item) => item.id === log.studentId);
    const coachId = log.coachId || program?.coachId || student?.primaryCoachId || coach.id;
    return coachId === coach.id;
  });
}

function getDefaultEditingProgramId() {
  const coach = getCurrentCoach();
  if (!coach) {
    return null;
  }

  const coachPrograms = state.programs
    .filter((program) => (program.coachId || coach.id) === coach.id)
    .sort((a, b) => {
      const dateCompare = String(b.date || "").localeCompare(String(a.date || ""));
      if (dateCompare !== 0) {
        return dateCompare;
      }
      return String(b.createdAt || "").localeCompare(String(a.createdAt || ""));
    });

  return getPublishedProgram(coach.id)?.id || coachPrograms[0]?.id || createBlankProgram().id;
}

function cleanupRedundantBlankPrograms() {
  const removableIds = new Set();

  state.coaches.forEach((coach) => {
    const blankPrograms = state.programs
      .filter((program) => (program.coachId || "") === coach.id)
      .filter((program) => !program.published)
      .filter((program) => !String(program.code || "").trim())
      .filter((program) => !String(program.title || "").trim())
      .filter((program) => !String(program.notes || "").trim())
      .filter((program) => getProgramItems(program.id).length === 0)
      .sort((a, b) => String(b.createdAt || "").localeCompare(String(a.createdAt || "")));

    blankPrograms.slice(1).forEach((program) => removableIds.add(program.id));
  });

  if (!removableIds.size) {
    return;
  }

  state.programs = state.programs.filter((program) => !removableIds.has(program.id));
  state.programItems = state.programItems.filter((item) => !removableIds.has(item.programId));
  persistState();
}


const COMMON_EXERCISES = [
  "\u7206\u767c\u9ad8\u62c9",
  "\u5bec\u63e1\u9ad8\u62c9",
  "\u6df1\u8e72",
  "SSB\u6df1\u8e72",
  "\u80a9\u63a8",
  "\u81e5\u63a8",
  "\u5730\u9707\u81e5\u63a8",
  "\u786c\u8209",
  "\u50b3\u7d71\u786c\u8209",
  "RDL",
  "\u5c48\u9ad4\u5212\u8239",
  "\u55ae\u624b\u5212\u8239",
  "\u5f8c\u8173\u62ac\u9ad8\u8e72",
  "\u524d\u8173\u62ac\u9ad8\u5206\u817f\u8e72",
  "\u55ae\u8dea\u59ff\u55ae\u80a9\u63a8",
  "\u55ae\u908aRDL",
  "\u58fa\u9234\u64fa\u76ea",
  "\u6536\u62cc\u5f0f",
  "\u5f15\u9ad4\u5411\u4e0a",
  "\u624b\u81c2\u88dc\u5f37",
  "\u4e8c\u982d\u808c\u5f4e\u8209"
];

const els = {
  mainTabsNav: document.querySelector(".main-tabs"),
  mainTabs: [...document.querySelectorAll(".main-tab")],
  panels: [...document.querySelectorAll(".panel")],
  coachTabs: [...document.querySelectorAll(".sub-tab[data-coach-tab]")],
  coachPanels: [...document.querySelectorAll(".sub-panel")],
  currentCoachSelect: document.querySelector("#current-coach-select"),
  coachSessionName: document.querySelector("#coach-session-name"),
  coachActiveBar: document.querySelector("#coach-active-bar"),
  coachActiveCopy: document.querySelector("#coach-active-copy"),
  changeCoachAccess: document.querySelector("#change-coach-access"),
  newCoachName: document.querySelector("#new-coach-name"),
  addCoachButton: document.querySelector("#add-coach"),
  coachRosterCount: document.querySelector("#coach-roster-count"),
  coachRoster: document.querySelector("#coach-roster"),
  programCode: document.querySelector("#program-code"),
  programDate: document.querySelector("#program-date"),
  coachName: document.querySelector("#coach-name"),
  programNotes: document.querySelector("#program-notes"),
  programTargetStudents: document.querySelector("#program-target-students"),
  programItemsBody: document.querySelector("#program-items-body"),
  programItemTemplate: document.querySelector("#program-item-template"),
  coachExerciseOptions: document.querySelector("#coach-exercise-options"),
  programLibraryCount: document.querySelector("#program-library-count"),
  programLibrarySort: document.querySelector("#program-library-sort"),
  programLibrarySearch: document.querySelector("#program-library-search"),
  programLibraryList: document.querySelector("#program-library-list"),
  programLibraryMore: document.querySelector("#program-library-more"),
  coachStats: document.querySelector("#coach-stats"),
  coachTodayDate: document.querySelector("#coach-today-date"),
  submittedList: document.querySelector("#submitted-list"),
  pendingList: document.querySelector("#pending-list"),
  todayLogsCount: document.querySelector("#today-logs-count"),
  todayLogsSearch: document.querySelector("#today-logs-search"),
  todayLogsBody: document.querySelector("#today-logs-body"),
  todayLogsMore: document.querySelector("#today-logs-more"),
  todayStudentDetailStatus: document.querySelector("#today-student-detail-status"),
  todayStudentDetailBody: document.querySelector("#today-student-detail-body"),
  todayStudentDetailMore: document.querySelector("#today-student-detail-more"),
  coachHistoryStudent: document.querySelector("#coach-history-student"),
  coachHistoryExercise: document.querySelector("#coach-history-exercise"),
  coachHistoryDateFrom: document.querySelector("#coach-history-date-from"),
  coachHistoryDateTo: document.querySelector("#coach-history-date-to"),
  coachHistoryLast30: document.querySelector("#coach-history-last-30"),
  coachHistoryCount: document.querySelector("#coach-history-count"),
  coachHistoryBody: document.querySelector("#coach-history-body"),
  coachHistoryMore: document.querySelector("#coach-history-more"),
  coachHistoryResultsWrap: document.querySelector("#coach-history-results-wrap"),
  toggleCoachHistoryResults: document.querySelector("#toggle-coach-history-results"),
  coachAuthShell: document.querySelector("#coach-auth-shell"),
  coachAccessCode: document.querySelector("#coach-access-code"),
  confirmCoachAccess: document.querySelector("#confirm-coach-access"),
  coachSummary: document.querySelector("#coach-summary"),
  runCoachHistoryButton: document.querySelector("#run-coach-history"),
  resetCoachHistoryButton: document.querySelector("#reset-coach-history"),
  exportCoachHistoryButton: document.querySelector("#export-coach-history"),
  coachStudentLinkName: document.querySelector("#coach-student-link-name"),
  coachStudentLinkSelect: document.querySelector("#coach-student-link-select"),
  coachStudentLinkOptions: document.querySelector("#coach-student-link-options"),
  coachStudentLinks: document.querySelector("#coach-student-links"),
  coachStudentLinksCard: document.querySelector("#coach-student-links-card"),
  historyImportCard: document.querySelector("#history-import-card"),
  historyImportCount: document.querySelector("#history-import-count"),
  downloadHistoryTemplateButton: document.querySelector("#download-history-template"),
  historyImportFile: document.querySelector("#history-import-file"),
  historyImportPreviewBody: document.querySelector("#history-import-preview-body"),
  confirmHistoryImportButton: document.querySelector("#confirm-history-import"),
  newStudentName: document.querySelector("#new-student-name"),
  addStudentButton: document.querySelector("#add-student"),
  coachStudentCount: document.querySelector("#coach-student-count"),
  coachStudentRosterSearch: document.querySelector("#coach-student-roster-search"),
  clearStudentRosterSearchButton: document.querySelector("#clear-student-roster-search"),
  coachStudentRosterSummary: document.querySelector("#coach-student-roster-summary"),
  coachStudentRoster: document.querySelector("#coach-student-roster"),
  coachGlobalViewDesktopButton: document.querySelector("#coach-global-view-desktop"),
  coachGlobalViewMobileButton: document.querySelector("#coach-global-view-mobile"),
  studentPanelViewToggle: document.querySelector("#student-panel .panel-header .view-toggle"),
  studentAccessCode: document.querySelector("#student-access-code"),
  studentProgramSelect: document.querySelector("#student-program-select"),
  studentSummary: document.querySelector("#student-summary"),
  studentAutoLoginBanner: document.querySelector("#student-autologin-banner"),
  studentAuthShell: document.querySelector("#student-auth-shell"),
  studentAuthCard: document.querySelector("#student-auth-card"),
  studentProgramShell: document.querySelector(".student-program-shell"),
  studentActiveBar: document.querySelector("#student-active-bar"),
  studentActiveCopy: document.querySelector("#student-active-copy"),
  studentRecordedBanner: document.querySelector("#student-recorded-banner"),
  studentProgramStatus: document.querySelector("#student-program-status"),
  studentProgramCard: document.querySelector("#student-program-card"),
  studentLeaveEntryRow: document.querySelector("#student-leave-entry-row"),
  openStudentLeaveSystemMain: document.querySelector("#open-student-leave-system-main"),
  openCoachLeaveSystemTab: document.querySelector("#open-coach-leave-system-tab"),
  studentPhoneFrame: document.querySelector("#student-phone-frame"),
  studentMobileTools: document.querySelector("#student-mobile-tools"),
  studentMobileSecondaryActions: document.querySelector("#student-mobile-secondary-actions"),
  studentCardList: document.querySelector("#student-card-list"),
  studentTableWrap: document.querySelector("#student-table-wrap"),
  studentProgramBody: document.querySelector("#student-program-body"),
  studentEntryTemplate: document.querySelector("#student-entry-template"),
  studentHistoryExercise: document.querySelector("#student-history-exercise"),
  studentHistoryExerciseOptions: document.querySelector("#student-history-exercise-options"),
  studentHistoryCount: document.querySelector("#student-history-count"),
  studentHistoryBody: document.querySelector("#student-history-body"),
  studentHistoryCard: document.querySelector("#student-history-card"),
  studentHistoryCardList: document.querySelector("#student-history-card-list"),
  studentHistoryMore: document.querySelector("#student-history-more"),
  studentGlobalViewCardButton: document.querySelector("#student-global-view-card"),
  studentGlobalViewTableButton: document.querySelector("#student-global-view-table"),
  studentViewCardButton: document.querySelector("#student-view-card"),
  studentViewTableButton: document.querySelector("#student-view-table"),
  loadStudentProgramInline: document.querySelector("#load-student-program-inline"),
  loadStudentProgramMobile: document.querySelector("#load-student-program-mobile"),
  editStudentProgramMobileTop: document.querySelector("#edit-student-program-mobile-top"),
  editStudentProgram: document.querySelector("#edit-student-program"),
  editStudentProgramMobile: document.querySelector("#edit-student-program-mobile"),
  openStudentLeaveSystemInline: document.querySelector("#open-student-leave-system-inline"),
  openStudentLeaveSystemMobilePrimary: document.querySelector("#open-student-leave-system-mobile-primary"),
  openStudentLeaveSystemMobileSecondary: document.querySelector("#open-student-leave-system-mobile-secondary"),
  openStudentHistoryMobile: document.querySelector("#open-student-history-mobile"),
  changeStudentAccess: document.querySelector("#change-student-access"),
  submitStudentLog: document.querySelector("#submit-student-log"),
  submitStudentLogMobile: document.querySelector("#submit-student-log-mobile"),
  studentMobileSubmitBar: document.querySelector("#student-mobile-submit-bar"),
  confirmModal: document.querySelector("#confirm-modal"),
  confirmModalCard: document.querySelector("#confirm-modal-card"),
  confirmSummary: document.querySelector("#confirm-summary"),
  confirmCardList: document.querySelector("#confirm-card-list"),
  confirmBody: document.querySelector("#confirm-body"),
  closeConfirmModal: document.querySelector("#close-confirm-modal"),
  cancelSubmit: document.querySelector("#cancel-submit"),
  confirmSubmit: document.querySelector("#confirm-submit"),
  studentProgramEditModal: document.querySelector("#student-program-edit-modal"),
  studentProgramEditModalCard: document.querySelector("#student-program-edit-modal-card"),
  studentProgramEditBody: document.querySelector("#student-program-edit-body"),
  studentEditMobileList: document.querySelector("#student-edit-mobile-list"),
  studentProgramEditNotice: document.querySelector("#student-program-edit-notice"),
  studentProgramEditRowTemplate: document.querySelector("#student-program-edit-row-template"),
  programPreviewModal: document.querySelector("#program-preview-modal"),
  modalPreviewCode: document.querySelector("#modal-preview-code"),
  modalPreviewBody: document.querySelector("#modal-program-preview-body"),
  successModal: document.querySelector("#success-modal"),
  successModalCard: document.querySelector("#success-modal-card"),
  successSummary: document.querySelector("#success-summary"),
  successCardList: document.querySelector("#success-card-list"),
  successBody: document.querySelector("#success-body"),
  downloadSuccessImage: document.querySelector("#download-success-image"),
  successOpenHistory: document.querySelector("#success-open-history"),
  assignCoachModal: document.querySelector("#assign-coach-modal"),
  assignCoachSummary: document.querySelector("#assign-coach-summary"),
  assignCoachSelect: document.querySelector("#assign-coach-select"),
  closeAssignCoachModal: document.querySelector("#close-assign-coach-modal"),
  cancelAssignCoach: document.querySelector("#cancel-assign-coach"),
  saveAssignCoach: document.querySelector("#save-assign-coach")
};

init().catch((error) => {
  console.error("CoachFlow init failed:", error);
  window.alert(error?.message || "系統初始化失敗，請重新整理後再試。");
});

function refreshAdminWorkspace() {
  renderCoachRoster();
  renderCoachStudentRoster();
}

function refreshCoachWorkspace(options = {}) {
  const preserveEditor = options.preserveEditor ?? (APP_MODE === "coach" && coachEditorDirty);
  const preservedTargetStudentIds = preserveEditor ? getSelectedProgramTargetStudentIds() : [];
  const nextEditingProgramId =
    coachProgramEditorMode === "edit"
      ? (
          editingProgramId && state.programs.some((program) => program.id === editingProgramId)
            ? editingProgramId
            : getDefaultEditingProgramId()
        )
      : null;
  const canPreserveEditor = Boolean(
    preserveEditor &&
    (
      coachProgramEditorMode === "create" ||
      (
        editingProgramId &&
        nextEditingProgramId &&
        editingProgramId === nextEditingProgramId &&
        state.programs.some((program) => program.id === editingProgramId)
      )
    )
  );

  editingProgramId = nextEditingProgramId;
  renderCurrentCoachOptions();
  renderCoachRoster();
  renderCoachExerciseOptions();
  getCoachTodayDate();
  if (!canPreserveEditor) {
    if (coachProgramEditorMode === "edit" && editingProgramId) {
      seedEditorFromProgram(editingProgramId);
    } else {
      seedEditorFromProgram(null);
    }
  }
  renderCoachHistoryFilters();
  renderProgramLibrary();
  renderCoachStudentLinks();
  renderHistoryImportPreview();
  renderCoachStudentRoster();
  if (canPreserveEditor) {
    renderProgramTargetStudents(preservedTargetStudentIds);
  }
  syncEditorPreviewState();
  renderCoachToday();
  renderCoachHistory();
  renderCoachSummary();
  syncCoachAccessUI();
}

function refreshStudentWorkspace() {
  renderStudentProgramOptions();
  renderStudentHistoryFilters();
  renderStudentHistory();
  renderStudentSummary();
  syncStudentAccessUI();
}

function delay(ms) {
  return new Promise((resolve) => window.setTimeout(resolve, ms));
}

function setControlDisabled(control, disabled) {
  if (!control) {
    return;
  }

  if (disabled) {
    if (!control.dataset.busyDisabledState) {
      control.dataset.busyDisabledState = control.disabled ? "disabled" : "enabled";
    }
    control.disabled = true;
    return;
  }

  const previousState = control.dataset.busyDisabledState;
  if (previousState) {
    control.disabled = previousState === "disabled";
    delete control.dataset.busyDisabledState;
  } else {
    control.disabled = false;
  }
}

function setButtonBusy(button, busy, busyLabel = "") {
  if (!button) {
    return;
  }

  if (busy) {
    if (!button.dataset.defaultLabel) {
      button.dataset.defaultLabel = button.textContent || "";
    }
    if (busyLabel) {
      button.textContent = busyLabel;
    }
    button.classList.add("is-busy");
    button.setAttribute("aria-busy", "true");
    setControlDisabled(button, true);
    return;
  }

  if (button.dataset.defaultLabel) {
    button.textContent = button.dataset.defaultLabel;
    delete button.dataset.defaultLabel;
  }
  button.classList.remove("is-busy");
  button.removeAttribute("aria-busy");
  setControlDisabled(button, false);
}

function setButtonsBusy(buttons, busy, busyLabel = "") {
  buttons.filter(Boolean).forEach((button) => setButtonBusy(button, busy, busyLabel));
}

function setStudentAutoLoginPending(pending, message = "正在自動登入並載入課表，請稍候...") {
  studentAutoLoginPending = Boolean(pending);
  if (els.studentAutoLoginBanner) {
    els.studentAutoLoginBanner.textContent = message;
    els.studentAutoLoginBanner.classList.toggle("is-visible", studentAutoLoginPending);
    els.studentAutoLoginBanner.classList.toggle("is-hidden", !studentAutoLoginPending);
  }
  if (els.studentAuthShell) {
    els.studentAuthShell.classList.toggle("is-autologin", studentAutoLoginPending);
  }
  if (studentAutoLoginPending && els.studentProgramStatus) {
    els.studentProgramStatus.textContent = "正在自動登入...";
  }
}

async function ensureCoachAccessOnInit() {
  if (APP_MODE !== "coach") {
    return;
  }

  const urlAccess = getModeAccessFromUrl("coach");
  const candidateAccess = normalizeAccessInput(urlAccess || authenticatedCoachAccess || els.coachAccessCode?.value || "");
  if (!candidateAccess) {
    return;
  }

  if (els.coachAccessCode) {
    els.coachAccessCode.value = candidateAccess;
  }

  if (urlAccess || !getCurrentCoach()) {
    let confirmed = false;
    const retryDelays = [0, 700, 1500];
    for (let i = 0; i < retryDelays.length; i += 1) {
      if (retryDelays[i] > 0) {
        await delay(retryDelays[i]);
      }
      confirmed = await confirmCoachAccess({ silentError: true, skipTouch: true });
      if (confirmed) {
        break;
      }
    }
    if (!confirmed) {
      authenticatedCoachId = "";
      authenticatedCoachAccess = candidateAccess;
      persistSession();
      syncCoachAccessUI();
    }
  }
}

async function ensureStudentAccessOnInit() {
  if (APP_MODE !== "student") {
    return;
  }

  const urlAccess = getModeAccessFromUrl("student");
  const candidateAccess = normalizeAccessInput(urlAccess || currentStudentAccess || els.studentAccessCode?.value || "");
  if (!candidateAccess) {
    return;
  }

  if (els.studentAccessCode) {
    els.studentAccessCode.value = candidateAccess;
  }

  let shouldAutoLogin = false;

  if (urlAccess || !getSelectedStudent()) {
    shouldAutoLogin = true;
    setStudentAutoLoginPending(true);
    let confirmed = false;
    const retryDelays = [0, 700, 1500];
    try {
      for (let i = 0; i < retryDelays.length; i += 1) {
        if (retryDelays[i] > 0) {
          await delay(retryDelays[i]);
        }
        confirmed = await confirmStudentAccess({ silentError: true, skipTouch: true });
        if (confirmed) {
          break;
        }
      }
      if (!confirmed) {
        currentStudentId = "";
        currentStudentAccess = candidateAccess;
        persistSession();
        syncStudentAccessUI();
        return;
      }
    } finally {
      setStudentAutoLoginPending(false);
    }
  }

  if (!shouldAutoLogin) {
    setStudentAutoLoginPending(false);
  }

  if (currentStudentId && !loadedStudentEntries.length) {
    loadStudentProgram({ silent: true, preserveIfDirty: true });
  }
}

async function init() {
  applyStaticCopy();
  applyAppMode();
  if (IS_CLOUD_MODE) {
    resetStateForCloudBootstrap();
  } else {
    hydrateFromStorage();
  }
  hydrateSession();
  await bootstrapCloudMode();
  if (APP_MODE === "coach" && !authenticatedCoachId) {
    await hydrateCoachAccessFromUrl();
  } else if (APP_MODE === "student" && !currentStudentId) {
    await hydrateStudentAccessFromUrl();
  } else if (APP_MODE === "dual") {
    await hydrateCoachAccessFromUrl();
    await hydrateStudentAccessFromUrl();
  }
  await ensureCoachAccessOnInit();
  await ensureStudentAccessOnInit();
  if (!IS_CLOUD_MODE) {
    cleanupRedundantBlankPrograms();
  }
  bindEvents();

  if (APP_MODE === "admin") {
    refreshAdminWorkspace();
    applyCoachViewMode();
    applyAppMode();
    switchCoachPanel("coach-coaches");
    return;
  }

  if (APP_MODE === "coach") {
    refreshCoachWorkspace();
  } else if (APP_MODE === "student") {
    refreshStudentWorkspace();
    if (!getSelectedStudent() && getModeAccessFromUrl("student")) {
      await ensureStudentAccessOnInit();
      refreshStudentWorkspace();
    }
  } else {
    refreshCoachWorkspace();
    refreshStudentWorkspace();
  }
  if (APP_MODE === "coach" && !authenticatedCoachId) {
    const coachAccessFromUrl = new URLSearchParams(window.location.search).get("coach")
      || new URLSearchParams(window.location.search).get("token")
      || new URLSearchParams(window.location.search).get("code")
      || "";
    if (coachAccessFromUrl && els.coachAccessCode) {
      els.coachAccessCode.value = coachAccessFromUrl;
      await confirmCoachAccess();
    }
  }
  if (APP_MODE === "coach") {
    markCoachUsed(getCurrentCoach()?.id || "");
  }
  if (currentStudentId) {
    loadStudentProgram({ silent: true, preserveIfDirty: true });
  }
  if (APP_MODE === "coach") {
    syncCoachAccessUI();
  }
  if (APP_MODE === "student") {
    syncStudentAccessUI();
  }
  applyCoachViewMode();
  applyStudentViewMode();
  applyAppMode();
}

function renderCoachExerciseOptions() {
  if (!els.coachExerciseOptions) {
    return;
  }

  const known = new Set(COMMON_EXERCISES);
  state.programItems.forEach((item) => {
    if (item.exercise) {
      known.add(item.exercise);
    }
  });

  els.coachExerciseOptions.innerHTML = [...known]
    .sort((a, b) => a.localeCompare(b, "zh-Hant"))
    .map((exercise) => `<option value="${exercise}"></option>`)
    .join("");
}

function renderCurrentCoachOptions() {
  if (!els.currentCoachSelect) {
    return;
  }
  const currentCoach = getCurrentCoach();
  els.currentCoachSelect.innerHTML = (state.coaches || [])
    .filter((coach) => coach.status !== "inactive")
    .map((coach) => `<option value="${coach.id}">${coach.name}</option>`)
    .join("");
  if (currentCoach) {
    els.currentCoachSelect.value = currentCoach.id;
  }
  if (els.coachName) {
    els.coachName.value = currentCoach?.name || "";
  }
  if (els.coachSessionName) {
    els.coachSessionName.value = currentCoach?.name || "";
  }
}

function renderCoachRoster() {
  if (!els.coachRoster) {
    return;
  }

  const coaches = state.coaches || [];
  if (els.coachRosterCount) {
    els.coachRosterCount.textContent = `共 ${coaches.length} 位`;
  }

  els.coachRoster.innerHTML = coaches.length
    ? coaches
        .map((coach) => {
          const studentCount = state.students.filter((student) => (student.primaryCoachId || "") === coach.id).length;
          const programCount = state.programs.filter((program) => (program.coachId || "") === coach.id).length;
          const logCount = state.workoutLogs.filter((log) => (log.coachId || "") === coach.id).length;
          const isCurrent = state.currentCoachId === coach.id;
          const isInactive = coach.status === "inactive";
          const lastUsedText = formatUsageTimestamp(coach.lastUsedAt);

          return `
            <article class="coach-link-card" data-coach-card="${coach.id}">
              <div class="coach-link-top">
                <div>
                  <h4>${coach.name}</h4>
                  <p class="coach-student-meta">教練代碼：${coach.accessCode || "-"}</p>
                  <p class="coach-student-meta">學生 ${studentCount} 人｜課表 ${programCount} 份｜紀錄 ${logCount} 筆</p>
                  <p class="coach-student-meta">最後使用：${lastUsedText}</p>
                </div>
                <div class="roster-pill-stack">
                  ${APP_MODE !== "admin" ? `<span class="status-pill ${isCurrent ? "is-success" : "is-muted"}">${isCurrent ? "目前教練" : "可切換"}</span>` : ""}
                  <span class="status-pill ${isInactive ? "is-muted" : "is-success"}">${isInactive ? "已停用" : "使用中"}</span>
                </div>
              </div>
              <div class="coach-link-actions">
                ${APP_MODE !== "admin" ? `<button class="ghost-button" type="button" data-set-current-coach="${coach.id}" ${isCurrent ? "disabled" : ""}>設為目前教練</button>` : ""}
                <button class="ghost-button" type="button" data-edit-coach="${coach.id}">修改姓名</button>
                <button class="ghost-button" type="button" data-toggle-coach-status="${coach.id}">
                  ${isInactive ? "啟用教練" : "停用教練"}
                </button>
                ${APP_MODE === "admin" ? `<button class="ghost-button" type="button" data-delete-coach="${coach.id}">刪除教練</button>` : ""}
              </div>
            </article>
          `;
        })
        .join("")
    : `<div class="empty-card">目前尚未建立教練。</div>`;
}

function getActiveStudents() {
  return state.students.filter((student) => student.status !== "inactive");
}

function getCoachActiveStudents() {
  return getCoachScopedStudents().filter((student) => student.status !== "inactive");
}

function getCoachTodayDate() {
  const fallback = getTodayDateInAppZone();
  if (els.coachTodayDate && !els.coachTodayDate.value) {
    els.coachTodayDate.value = fallback;
  }
  return normalizeDateKey(els.coachTodayDate?.value || fallback) || fallback;
}

function getCoachRoundProgram(viewDate = getCoachTodayDate()) {
  const viewDateKey = getComparableDateKey(viewDate) || getTodayDateInAppZone();
  const scopedPrograms = getCoachScopedPrograms();
  const datedPrograms = scopedPrograms
    .map((program) => ({
      program,
      dateKey: getComparableDateKey(program.date),
      createdKey: getComparableDateTimeKey(program.createdAt)
    }))
    .filter((item) => item.dateKey);

  const currentOrPast = datedPrograms
    .filter((item) => item.dateKey <= viewDateKey)
    .sort((a, b) => {
      const dateCompare = b.dateKey.localeCompare(a.dateKey);
      if (dateCompare !== 0) {
        return dateCompare;
      }
      if (a.program.published !== b.program.published) {
        return a.program.published ? -1 : 1;
      }
      return b.createdKey.localeCompare(a.createdKey);
    });

  if (currentOrPast.length) {
    return currentOrPast[0].program;
  }

  return getPublishedProgram(getCurrentCoach()?.id)
    || datedPrograms.sort((a, b) => a.dateKey.localeCompare(b.dateKey))[0]?.program
    || null;
}

function getCoachNextRoundDate(program, studentId = "") {
  const programDateKey = getComparableDateKey(program?.date);
  if (!programDateKey) {
    return "";
  }

  return getCoachScopedPrograms()
    .filter((item) => !studentId || isProgramAvailableForStudent(item, studentId))
    .map((item) => ({
      id: item.id,
      dateKey: getComparableDateKey(item.date)
    }))
    .filter((item) => item.id !== program.id && item.dateKey && item.dateKey > programDateKey)
    .sort((a, b) => a.dateKey.localeCompare(b.dateKey))[0]?.dateKey || "";
}

function isLogInProgramRound(log, program) {
  if (!program || log?.programId !== program.id) {
    return false;
  }

  const roundStart = getComparableDateKey(program.date);
  const logDate = getLogActivityDateKey(log);
  if (!roundStart || !logDate || logDate < roundStart) {
    return false;
  }

  const nextRoundDate = getCoachNextRoundDate(program, log.studentId || "");
  return !nextRoundDate || logDate < nextRoundDate;
}

function getStudentRoundProgram(student, viewDate = getCoachTodayDate()) {
  const viewDateKey = getComparableDateKey(viewDate) || getTodayDateInAppZone();
  const studentId = student?.id || "";
  const coachId = student?.primaryCoachId || getCurrentCoach()?.id || "";
  const datedPrograms = getCoachScopedPrograms()
    .filter((program) => !coachId || (program.coachId || "") === coachId)
    .filter((program) => isProgramAvailableForStudent(program, studentId))
    .map((program) => ({
      program,
      dateKey: getComparableDateKey(program.date),
      createdKey: getComparableDateTimeKey(program.createdAt)
    }))
    .filter((item) => item.dateKey);

  const currentOrPast = datedPrograms
    .filter((item) => item.dateKey <= viewDateKey)
    .sort((a, b) => {
      const dateCompare = b.dateKey.localeCompare(a.dateKey);
      if (dateCompare !== 0) {
        return dateCompare;
      }
      if (isProgramTargeted(a.program) !== isProgramTargeted(b.program)) {
        return isProgramTargeted(a.program) ? -1 : 1;
      }
      if (a.program.published !== b.program.published) {
        return a.program.published ? -1 : 1;
      }
      return b.createdKey.localeCompare(a.createdKey);
    });

  if (currentOrPast.length) {
    return currentOrPast[0].program;
  }

  return getPublishedProgram(coachId, studentId)
    || datedPrograms.sort((a, b) => a.dateKey.localeCompare(b.dateKey))[0]?.program
    || null;
}

function getCoachRoundContext(viewDate = getCoachTodayDate()) {
  const normalizedViewDate = normalizeDateKey(viewDate) || getTodayDateInAppZone();
  const viewDateKey = getComparableDateKey(normalizedViewDate);
  const program = getCoachRoundProgram(normalizedViewDate);
  const coachLogs = getCoachScopedLogs();
  const studentProgramMap = new Map();
  getCoachActiveStudents().forEach((student) => {
    const studentProgram = getStudentRoundProgram(student, normalizedViewDate);
    if (studentProgram) {
      studentProgramMap.set(student.id, studentProgram);
    }
  });
  const logs = studentProgramMap.size
    ? coachLogs.filter((log) => {
        const studentProgram = studentProgramMap.get(log.studentId);
        return studentProgram ? isLogInProgramRound(log, studentProgram) : false;
      })
    : program
      ? coachLogs.filter((log) => isLogInProgramRound(log, program))
      : coachLogs.filter((log) => getLogActivityDateKey(log) === viewDateKey);

  return {
    viewDate: normalizedViewDate,
    viewDateKey,
    program,
    studentProgramMap,
    logs: [...logs].sort(compareLogsByDateTimeDesc)
  };
}

function bindEvents() {
  els.mainTabs.forEach((button) => {
    button.addEventListener("click", () => switchMainPanel(button.dataset.panel));
  });

  els.coachTabs.forEach((button) => {
    button.addEventListener("click", () => {
      if (button.dataset.coachTab === "coach-editor") {
        coachProgramEditorMode = "create";
        editingProgramId = null;
        if (!coachEditorDirty) {
          seedEditorFromProgram(null);
        }
      }
      switchCoachPanel(button.dataset.coachTab);
    });
  });
  els.currentCoachSelect?.addEventListener("change", handleCurrentCoachChange);
  els.addCoachButton?.addEventListener("click", createCoachFromForm);
  els.confirmCoachAccess?.addEventListener("click", confirmCoachAccess);
  els.changeCoachAccess?.addEventListener("click", resetCoachAccess);

  document.querySelector("#add-program-item").addEventListener("click", () => {
    addProgramItemRow();
    syncSortNumbers();
    syncEditorPreviewState();
    setCoachEditorDirty(true);
  });
  document.querySelector("#save-program").addEventListener("click", saveProgram);
  document.querySelector("#publish-program").addEventListener("click", publishProgram);
  document.querySelector("#publish-targeted-program")?.addEventListener("click", publishTargetedProgram);
  els.programTargetStudents?.addEventListener("change", (event) => {
    const input = event.target.closest("input[type='checkbox']");
    input?.closest(".program-target-option")?.classList.toggle("is-selected", input.checked);
    setCoachEditorDirty(true);
  });
  els.programLibraryList?.addEventListener("click", handleProgramLibraryAction);
  els.programLibrarySearch?.addEventListener("input", () => {
    programLibraryVisibleCount = 6;
    renderProgramLibrary();
  });
  els.programLibrarySort?.addEventListener("change", () => {
    programLibraryVisibleCount = 6;
    renderProgramLibrary();
  });
  els.programLibraryMore?.addEventListener("click", toggleProgramLibraryVisibleCount);
  document.querySelector("#close-program-preview-modal").addEventListener("click", closeProgramPreviewModal);
  document.querySelector("#modal-close-preview").addEventListener("click", closeProgramPreviewModal);
  document.querySelector("#modal-export-program-image").addEventListener("click", exportProgramImage);
  els.toggleCoachHistoryResults?.addEventListener("click", toggleCoachHistoryResults);
  els.confirmStudentAccess?.addEventListener("click", confirmStudentAccessAndLoadProgram);
  els.loadStudentProgramInline.addEventListener("click", () => loadStudentProgramWithFeedback({ preserveIfDirty: true }));
  els.loadStudentProgramMobile.addEventListener("click", () => loadStudentProgramWithFeedback({ preserveIfDirty: true }));
  els.editStudentProgramMobileTop?.addEventListener("click", handleStudentDetailAction);
  els.editStudentProgram.addEventListener("click", handleStudentDetailAction);
  els.editStudentProgramMobile.addEventListener("click", handleStudentDetailAction);
  els.submitStudentLog.addEventListener("click", openSubmissionConfirm);
  els.submitStudentLogMobile.addEventListener("click", openSubmissionConfirm);
  els.openStudentHistoryMobile.addEventListener("click", () => {
    studentHistoryOpened = true;
    applyStudentViewMode();
    renderStudentHistory();
    els.studentHistoryCard.scrollIntoView({ behavior: "smooth", block: "start" });
  });
  els.changeStudentAccess.addEventListener("click", showStudentAccessPanel);
  [
    els.openStudentLeaveSystemMain,
    els.openStudentLeaveSystemInline,
    els.openStudentLeaveSystemMobilePrimary,
    els.openStudentLeaveSystemMobileSecondary
  ].forEach((button) => button?.addEventListener("click", openStudentLeaveSystemEntry));
  els.openCoachLeaveSystemTab?.addEventListener("click", openCoachLeaveSystemEntry);
  document.querySelector("#open-student-history").addEventListener("click", () => {
    renderStudentHistory();
    document.querySelector("#student-history-exercise").focus();
  });
  els.closeConfirmModal.addEventListener("click", closeConfirmModal);
  els.cancelSubmit.addEventListener("click", closeConfirmModal);
  els.confirmSubmit.addEventListener("click", finalizeSubmission);
  document.querySelector("#close-success-modal").addEventListener("click", closeSuccessModal);
  els.downloadSuccessImage?.addEventListener("click", downloadSuccessImage);
  els.successOpenHistory?.addEventListener("click", openSuccessHistory);
  document.querySelector("#close-student-program-edit-modal").addEventListener("click", closeStudentProgramEditModal);
  document.querySelector("#cancel-student-program-edit").addEventListener("click", closeStudentProgramEditModal);
  document.querySelector("#save-student-program-edit").addEventListener("click", saveStudentProgramEdits);
  document.querySelector("#add-student-edit-row").addEventListener("click", () => {
    addStudentProgramEditRow();
    showStudentProgramEditNotice("已新增 1 個課表項目，記得按「套用到本次課表」。");
    captureStudentProgramEditDraftFromDom();
  });
  els.studentProgramEditBody?.addEventListener("input", captureStudentProgramEditDraftFromDom);
  els.studentProgramEditBody?.addEventListener("change", captureStudentProgramEditDraftFromDom);
  els.studentEditMobileList?.addEventListener("input", captureStudentProgramEditDraftFromDom);
  els.studentEditMobileList?.addEventListener("change", captureStudentProgramEditDraftFromDom);
  els.closeAssignCoachModal?.addEventListener("click", closeAssignCoachModal);
  els.cancelAssignCoach?.addEventListener("click", closeAssignCoachModal);
  els.saveAssignCoach?.addEventListener("click", saveAssignedCoach);
  els.coachGlobalViewDesktopButton?.addEventListener("click", () => setCoachViewMode("desktop"));
  els.coachGlobalViewMobileButton?.addEventListener("click", () => setCoachViewMode("mobile"));
  els.studentGlobalViewCardButton.addEventListener("click", () => setStudentViewMode("card"));
  els.studentGlobalViewTableButton.addEventListener("click", () => setStudentViewMode("table"));
  els.studentViewCardButton.addEventListener("click", () => setStudentViewMode("card"));
  els.studentViewTableButton.addEventListener("click", () => setStudentViewMode("table"));

  const markStudentEntriesDirty = () => {
    if (!loadedStudentEntries.length) {
      return;
    }
    studentEntriesDirty = true;
    captureStudentEntryDraftsFromDom();
  };

  [els.programCode, els.programDate, els.coachName, els.programNotes].forEach((input) => {
    input.addEventListener("input", () => {
      setCoachEditorDirty(true);
      syncEditorPreviewState();
    });
  });

  els.studentAccessCode.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      confirmStudentAccessAndLoadProgram();
    }
  });
  els.coachAccessCode?.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      confirmCoachAccess();
    }
  });

  els.studentProgramSelect.addEventListener("change", renderStudentSummary);

  els.coachHistoryStudent.addEventListener("change", () => {
    coachHistoryVisibleCount = 6;
    renderCoachHistory();
  });
  els.coachHistoryExercise.addEventListener("change", () => {
    coachHistoryVisibleCount = 6;
    renderCoachHistory();
  });
  els.coachHistoryDateFrom?.addEventListener("change", () => {
    coachHistoryVisibleCount = 6;
    renderCoachHistory();
  });
  els.coachHistoryDateTo?.addEventListener("change", () => {
    coachHistoryVisibleCount = 6;
    renderCoachHistory();
  });
  els.todayLogsSearch?.addEventListener("input", () => {
    coachTodayVisibleCount = 6;
    renderCoachToday();
  });
  els.coachTodayDate?.addEventListener("change", () => {
    coachTodayVisibleCount = 6;
    renderCoachToday();
  });
  els.todayLogsBody?.addEventListener("click", handleTodayLogAction);
  els.todayLogsMore?.addEventListener("click", toggleTodayLogsVisibleCount);
  els.todayStudentDetailMore?.addEventListener("click", toggleTodayStudentDetailVisibleCount);
  els.submittedList?.addEventListener("click", handleTodayStudentListAction);
  els.pendingList?.addEventListener("click", handleTodayStudentListAction);
  els.coachStudentLinks?.addEventListener("click", handleCoachStudentLinkAction);
  els.studentProgramBody?.addEventListener("input", markStudentEntriesDirty);
  els.studentProgramBody?.addEventListener("change", markStudentEntriesDirty);
  els.studentCardList?.addEventListener("input", markStudentEntriesDirty);
  els.studentCardList?.addEventListener("change", markStudentEntriesDirty);
  els.coachStudentRoster?.addEventListener("click", handleCoachStudentRosterAction);
  els.coachRoster?.addEventListener("click", handleCoachRosterAction);
  els.coachStudentLinkName?.addEventListener("input", renderCoachStudentLinks);
  els.coachStudentLinkName?.addEventListener("change", renderCoachStudentLinks);
  els.downloadHistoryTemplateButton?.addEventListener("click", downloadHistoryImportTemplate);
  els.historyImportFile?.addEventListener("change", handleHistoryImportFile);
  els.confirmHistoryImportButton?.addEventListener("click", confirmHistoryImport);
  els.coachStudentRosterSearch?.addEventListener("input", renderCoachStudentRoster);
  els.clearStudentRosterSearchButton?.addEventListener("click", () => {
    if (els.coachStudentRosterSearch) {
      els.coachStudentRosterSearch.value = "";
    }
    renderCoachStudentRoster();
    els.coachStudentRosterSearch?.focus();
  });
  els.coachStudentLinkSelect?.addEventListener("change", () => {
    els.coachStudentLinkName.value = els.coachStudentLinkSelect.value;
    renderCoachStudentLinks();
  });
  els.addStudentButton?.addEventListener("click", createStudentFromCoachForm);
  els.studentHistoryExercise.addEventListener("input", () => {
    studentHistoryVisibleCount = 1;
    renderStudentHistory();
  });
  els.studentHistoryExercise.addEventListener("change", () => {
    studentHistoryVisibleCount = 1;
    renderStudentHistory();
  });
  els.studentHistoryMore?.addEventListener("click", toggleStudentHistoryVisibleCount);
  els.coachHistoryLast30?.addEventListener("click", applyCoachHistoryLast30Days);
  els.runCoachHistoryButton?.addEventListener("click", () => {
    coachHistoryVisibleCount = 6;
    renderCoachHistory();
  });
  els.resetCoachHistoryButton?.addEventListener("click", resetCoachHistoryFilters);
  els.exportCoachHistoryButton?.addEventListener("click", exportCoachHistoryCsv);
  els.coachHistoryMore?.addEventListener("click", toggleCoachHistoryVisibleCount);
  window.addEventListener("resize", handleViewportChange);

  els.confirmModal.addEventListener("click", (event) => {
    if (event.target === els.confirmModal) {
      closeConfirmModal();
    }
  });
  els.programPreviewModal.addEventListener("click", (event) => {
    if (event.target === els.programPreviewModal) {
      closeProgramPreviewModal();
    }
  });
  els.successModal.addEventListener("click", (event) => {
    if (event.target === els.successModal) {
      closeSuccessModal();
    }
  });
  els.studentProgramEditModal.addEventListener("click", (event) => {
    if (event.target === els.studentProgramEditModal) {
      closeStudentProgramEditModal();
    }
  });
  els.assignCoachModal?.addEventListener("click", (event) => {
    if (event.target === els.assignCoachModal) {
      closeAssignCoachModal();
    }
  });
  window.addEventListener("storage", handleStorageSync);
  window.addEventListener("focus", handleVisibilityRefresh);
  document.addEventListener("visibilitychange", handleVisibilityRefresh);
  window.addEventListener("beforeunload", handleBeforeUnload);
  window.addEventListener("pagehide", handlePageHide);
}

function hydrateFromStorage() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      persistState();
      return;
    }
    const parsed = JSON.parse(raw);
    state.coaches = (parsed.coaches || structuredClone(defaultState.coaches)).map((coach, index) => {
      const fallback = defaultState.coaches[index] || {};
      const generatedAccessCode = buildCoachAccessCode(coach.name || fallback.name || `Coach ${index + 1}`, index + 1);
      const generatedToken = String(coach.name || fallback.name || `coach${index + 1}`)
        .replace(/\s+/g, "")
        .toLowerCase()
        .replace(/[^a-z0-9]/g, "");
      return {
        ...fallback,
        ...coach,
        accessCode: coach.accessCode || fallback.accessCode || generatedAccessCode,
        token: coach.token || fallback.token || generatedToken || `coach${index + 1}`,
        lastUsedAt: normalizeDateTimeValue(coach.lastUsedAt || fallback.lastUsedAt || ""),
        status: coach.status || fallback.status || "active",
        role: coach.role || fallback.role || "coach"
      };
    });
    state.currentCoachId = parsed.currentCoachId || state.coaches[0]?.id || defaultState.currentCoachId;
    state.students = (parsed.students || structuredClone(defaultState.students)).map((student, index) => {
      const fallback = defaultState.students[index] || {};
      return {
        ...fallback,
        ...student,
        lastUsedAt: normalizeDateTimeValue(student.lastUsedAt || fallback.lastUsedAt || ""),
        accessCode: student.accessCode || fallback.accessCode || `CODE${index + 1}`,
        token: student.token || fallback.token || `token${index + 1}`,
        primaryCoachId: student.primaryCoachId || fallback.primaryCoachId || state.currentCoachId
      };
    });
    state.programs = (parsed.programs || structuredClone(defaultState.programs)).map((program) => ({
      ...program,
      date: normalizeDateKey(program.date || ""),
      createdAt: normalizeDateTimeValue(program.createdAt || ""),
      updatedAt: normalizeDateTimeValue(program.updatedAt || ""),
      coachId: program.coachId || state.currentCoachId,
      coachName: program.coachName || getCurrentCoach()?.name || "",
      targetStudentIds: normalizeProgramStudentIds(program.targetStudentIds || program.studentIds || "")
    }));
    state.programItems = parsed.programItems || structuredClone(defaultState.programItems);
    state.workoutLogs = (parsed.workoutLogs || structuredClone(defaultState.workoutLogs)).map((log) => {
      const program = (parsed.programs || structuredClone(defaultState.programs)).find((item) => item.id === log.programId);
      const student = (parsed.students || structuredClone(defaultState.students)).find((item) => item.id === log.studentId);
      const coachId = log.coachId || program?.coachId || student?.primaryCoachId || state.currentCoachId;
      const coachName = log.coachName || state.coaches.find((coach) => coach.id === coachId)?.name || "";
      return {
        ...log,
        programDate: normalizeDateKey(log.programDate || ""),
        submittedAt: normalizeDateTimeValue(log.submittedAt || ""),
        updatedAt: normalizeDateTimeValue(log.updatedAt || ""),
        coachId,
        coachName
      };
    }).filter(isUsableWorkoutLog);
    persistState();
  } catch {
    persistState();
  }
}

function resetStateForCloudBootstrap() {
  state.coaches = [];
  state.currentCoachId = "";
  state.students = [];
  state.programs = [];
  state.programItems = [];
  state.workoutLogs = [];
}

function persistState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  persistSession();
}

function setCoachEditorDirty(value = true) {
  coachEditorDirty = Boolean(value);
}

function hydrateSession() {
  try {
    const raw = localStorage.getItem(SESSION_KEY);
    if (!raw) {
      persistSession();
      return;
    }

    const parsed = JSON.parse(raw);
      authenticatedCoachId = String(parsed.authenticatedCoachId || "").trim();
      authenticatedCoachAccess = String(parsed.authenticatedCoachAccess || "").trim();
      currentStudentId = String(parsed.currentStudentId || "").trim();
      currentStudentAccess = String(parsed.currentStudentAccess || "").trim();
      currentStudentProgramId = String(parsed.currentStudentProgramId || "").trim();
      sessionStudentDraft = parsed.studentDraft && typeof parsed.studentDraft === "object"
        ? parsed.studentDraft
        : null;
      sessionStudentProgramEditDraft =
        parsed.studentProgramEditDraft && typeof parsed.studentProgramEditDraft === "object"
          ? parsed.studentProgramEditDraft
          : null;
      coachViewMode = parsed.coachViewMode || coachViewMode;
      studentViewMode = parsed.studentViewMode || studentViewMode;
    } catch {
    persistSession();
  }
}

function persistSession() {
  let existing = {};
  try {
    existing = JSON.parse(localStorage.getItem(SESSION_KEY) || "{}");
  } catch {
    existing = {};
  }

  const shouldPersistCoachSession = APP_MODE === "coach" || APP_MODE === "dual";
  const shouldPersistStudentSession = APP_MODE === "student" || APP_MODE === "dual";

  const nextSession = {
    authenticatedCoachId: shouldPersistCoachSession
      ? String(authenticatedCoachId || "")
      : String(existing.authenticatedCoachId || ""),
    authenticatedCoachAccess: shouldPersistCoachSession
      ? String(authenticatedCoachAccess || "")
      : String(existing.authenticatedCoachAccess || ""),
      currentStudentId: shouldPersistStudentSession
        ? String(currentStudentId || "")
        : String(existing.currentStudentId || ""),
      currentStudentAccess: shouldPersistStudentSession
        ? String(currentStudentAccess || "")
        : String(existing.currentStudentAccess || ""),
      currentStudentProgramId: shouldPersistStudentSession
        ? String(currentStudentProgramId || "")
        : String(existing.currentStudentProgramId || ""),
      coachViewMode: shouldPersistCoachSession
        ? coachViewMode
        : String(existing.coachViewMode || coachViewMode),
    studentViewMode: shouldPersistStudentSession
      ? studentViewMode
      : String(existing.studentViewMode || studentViewMode),
    studentDraft: shouldPersistStudentSession
      ? buildCurrentStudentDraftSession()
      : (existing.studentDraft && typeof existing.studentDraft === "object" ? existing.studentDraft : null),
    studentProgramEditDraft: shouldPersistStudentSession
      ? buildCurrentStudentProgramEditDraftSession(
          existing.studentProgramEditDraft && typeof existing.studentProgramEditDraft === "object"
            ? existing.studentProgramEditDraft
            : null
        )
      : (existing.studentProgramEditDraft && typeof existing.studentProgramEditDraft === "object"
          ? existing.studentProgramEditDraft
          : null)
  };

  sessionStudentDraft = nextSession.studentDraft || null;
  sessionStudentProgramEditDraft = nextSession.studentProgramEditDraft || null;
  localStorage.setItem(SESSION_KEY, JSON.stringify(nextSession));
}

function sanitizeStudentDraftEntries(entries = []) {
  return entries
    .map((entry) => ({
      itemId: String(entry.itemId || `draft-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`),
      category: String(entry.category || "主項目"),
      exercise: String(entry.exercise || "").trim(),
      targetSets: Number(entry.targetSets || 0),
      targetType: entry.targetType === "time" || entry.targetType === "rm" ? entry.targetType : "reps",
      targetValue: Number(entry.targetType === "rm" ? 0 : (entry.targetValue || 0)),
      itemNote: String(entry.itemNote || ""),
      actualWeight: String(entry.actualWeight || ""),
      actualSets: String(entry.actualSets || ""),
      actualReps: String(entry.actualReps || ""),
      studentNote: String(entry.studentNote || "")
    }))
    .filter((entry) => entry.exercise && entry.targetSets > 0 && (entry.targetType === "rm" || entry.targetValue > 0));
}

function buildCurrentStudentDraftSession() {
  const studentId = String(currentStudentId || "").trim();
  const programId = String(currentStudentProgramId || loadedStudentProgramId || "").trim();
  if (!studentId || !programId || !loadedStudentEntries.length) {
    return null;
  }

  const entries = sanitizeStudentDraftEntries(loadedStudentEntries);
  if (!entries.length) {
    return null;
  }

  return {
    studentId,
    programId,
    dirty: Boolean(studentEntriesDirty),
    savedAt: timestampNow(),
    entries
  };
}

function getSessionStudentDraft(studentId, programId) {
  if (!sessionStudentDraft || typeof sessionStudentDraft !== "object") {
    return null;
  }
  if (String(sessionStudentDraft.studentId || "") !== String(studentId || "")) {
    return null;
  }
  if (String(sessionStudentDraft.programId || "") !== String(programId || "")) {
    return null;
  }
  const entries = Array.isArray(sessionStudentDraft.entries) ? sessionStudentDraft.entries : [];
  if (!entries.length) {
    return null;
  }
  return sessionStudentDraft;
}

function sanitizeStudentProgramEditDraftRows(rows = []) {
  return rows.map((row) => ({
    sourceIndex:
      String(row.sourceIndex ?? "").trim() !== "" && Number.isInteger(Number(row.sourceIndex))
        ? Number(row.sourceIndex)
        : -1,
    category: String(row.category || "主項目"),
    exercise: String(row.exercise || "").trim(),
    targetSets: String(row.targetSets ?? ""),
    targetType: row.targetType === "time" || row.targetType === "rm" ? row.targetType : "reps",
    targetValue: String(row.targetValue ?? "")
  }));
}

function collectStudentProgramEditDraftRowsFromDom() {
  const editRows = studentViewMode === "card"
    ? [...(els.studentEditMobileList?.querySelectorAll(".student-edit-mobile-card") || [])]
    : [...(els.studentProgramEditBody?.querySelectorAll("tr") || [])];

  return sanitizeStudentProgramEditDraftRows(
    editRows.map((row) => ({
      sourceIndex: Number(row.dataset.sourceIndex),
      category: row.querySelector(".student-edit-category")?.value || "主項目",
      exercise: row.querySelector(".student-edit-exercise")?.value || "",
      targetSets: row.querySelector(".student-edit-sets")?.value || "",
      targetType: row.querySelector(".student-edit-type")?.value || "reps",
      targetValue: row.querySelector(".student-edit-value")?.value || ""
    }))
  );
}

function buildCurrentStudentProgramEditDraftSession(existingDraft = null) {
  const studentId = String(currentStudentId || "").trim();
  const programId = String(currentStudentProgramId || loadedStudentProgramId || "").trim();
  if (!studentId || !programId) {
    return null;
  }

  if (els.studentProgramEditModal && !els.studentProgramEditModal.classList.contains("is-hidden")) {
    const rows = collectStudentProgramEditDraftRowsFromDom();
    if (!rows.length) {
      return null;
    }
    return {
      studentId,
      programId,
      savedAt: timestampNow(),
      rows
    };
  }

  if (
    sessionStudentProgramEditDraft
    && String(sessionStudentProgramEditDraft.studentId || "") === studentId
    && String(sessionStudentProgramEditDraft.programId || "") === programId
    && Array.isArray(sessionStudentProgramEditDraft.rows)
    && sessionStudentProgramEditDraft.rows.length
  ) {
    return sessionStudentProgramEditDraft;
  }

  if (
    existingDraft
    && String(existingDraft.studentId || "") === studentId
    && String(existingDraft.programId || "") === programId
    && Array.isArray(existingDraft.rows)
    && existingDraft.rows.length
  ) {
    return existingDraft;
  }

  return null;
}

function getSessionStudentProgramEditDraft(studentId, programId) {
  if (!sessionStudentProgramEditDraft || typeof sessionStudentProgramEditDraft !== "object") {
    return null;
  }
  if (String(sessionStudentProgramEditDraft.studentId || "") !== String(studentId || "")) {
    return null;
  }
  if (String(sessionStudentProgramEditDraft.programId || "") !== String(programId || "")) {
    return null;
  }
  const rows = Array.isArray(sessionStudentProgramEditDraft.rows) ? sessionStudentProgramEditDraft.rows : [];
  if (!rows.length) {
    return null;
  }
  return sessionStudentProgramEditDraft;
}

function clearStudentProgramEditDraft(shouldPersist = true) {
  sessionStudentProgramEditDraft = null;
  if (shouldPersist) {
    persistSession();
  }
}

function captureStudentProgramEditDraftFromDom() {
  if (!els.studentProgramEditModal || els.studentProgramEditModal.classList.contains("is-hidden")) {
    return;
  }
  const studentId = String(currentStudentId || "").trim();
  const programId = String(currentStudentProgramId || loadedStudentProgramId || "").trim();
  if (!studentId || !programId) {
    return;
  }
  const rows = collectStudentProgramEditDraftRowsFromDom();
  sessionStudentProgramEditDraft = rows.length
    ? {
        studentId,
        programId,
        savedAt: timestampNow(),
        rows
      }
    : null;
  persistSession();
}

function captureStudentEntryDraftsFromDom() {
  if (!loadedStudentEntries.length) {
    return;
  }

  loadedStudentEntries = loadedStudentEntries.map((entry, index) => {
    const nextEntry = { ...entry };
    const actualWeightInput = getActiveEntryField(index, "actualWeight");
    const actualSetsInput = getActiveEntryField(index, "actualSets");
    const actualRepsInput = getActiveEntryField(index, "actualReps");
    const noteInput = getActiveEntryField(index, "studentNote");

    if (actualWeightInput) nextEntry.actualWeight = actualWeightInput.value;
    if (actualSetsInput) nextEntry.actualSets = actualSetsInput.value;
    if (actualRepsInput) nextEntry.actualReps = actualRepsInput.value;
    if (noteInput) nextEntry.studentNote = noteInput.value;

    return nextEntry;
  });
  persistSession();
}

function buildSubmittedSnapshot(logs, program = getSelectedStudentProgram(), student = getSelectedStudent()) {
  if (!logs?.length) {
    return null;
  }

  const firstLog = logs[0] || {};
  return {
    studentId: student?.id || firstLog.studentId || "",
    studentName: student?.name || firstLog.studentName || "",
    programId: program?.id || firstLog.programId || "",
    programCode: program?.code || firstLog.programCode || "",
    programDate: normalizeDateKey(getLogActivityDateKey(firstLog) || program?.date || ""),
    submittedAt: normalizeDateTimeValue(firstLog.submittedAt || ""),
    logs: logs.map((log) => ({ ...log }))
  };
}

function getActiveSubmittedSnapshot() {
  if (lastSubmittedSnapshot?.logs?.length) {
    return lastSubmittedSnapshot;
  }

  const logs = getLatestSubmittedLogsForCurrentStudentProgram();
  if (!logs.length) {
    return null;
  }

  return buildSubmittedSnapshot(logs);
}

function updateModeAccessInUrl(mode, accessValue) {
  const value = String(accessValue || "").trim();
  const url = new URL(window.location.href);
  ["code", "token", "coach", "student"].forEach((key) => url.searchParams.delete(key));
  if (mode === "coach" && value) {
    url.searchParams.set("code", value);
  }
  if (mode === "student" && value) {
    url.searchParams.set("code", value);
  }
  window.history.replaceState({}, "", url.toString());
}

function clearModeAccessInUrl() {
  const url = new URL(window.location.href);
  ["code", "token", "coach", "student"].forEach((key) => url.searchParams.delete(key));
  window.history.replaceState({}, "", url.toString());
}

function switchMainPanel(panelId) {
  els.mainTabs.forEach((button) => button.classList.toggle("is-active", button.dataset.panel === panelId));
  els.panels.forEach((panel) => panel.classList.toggle("is-active", panel.id === panelId));
}

function applyCoachTabVisibility() {
  const tabConfig = {
    "coach-editor": APP_MODE !== "admin",
    "coach-coaches": APP_MODE === "admin" || APP_MODE === "dual",
    "coach-students": true,
    "coach-today": APP_MODE !== "admin",
    "coach-history": APP_MODE !== "admin"
  };

  els.coachTabs.forEach((button) => {
    const visible = tabConfig[button.dataset.coachTab] !== false;
    button.hidden = !visible;
    button.style.display = visible ? "" : "none";
  });

  els.coachPanels.forEach((panel) => {
    const visible = tabConfig[panel.id] !== false;
    panel.dataset.available = visible ? "true" : "false";
    if (!visible) {
      panel.classList.remove("is-active");
      panel.hidden = true;
      panel.style.display = "none";
    } else {
      panel.hidden = false;
      panel.style.display = "";
    }
  });

  const activeCoachPanel = els.coachPanels.find((panel) => panel.classList.contains("is-active"));
  if (!activeCoachPanel || activeCoachPanel.dataset.available === "false") {
    if (APP_MODE === "admin") {
      switchCoachPanel("coach-coaches");
    } else {
      switchCoachPanel("coach-editor");
    }
  }
}

function applyAppMode() {
  const coachPanel = document.querySelector("#coach-panel");
  const studentPanel = document.querySelector("#student-panel");
  const coachHeaderBlock = document.querySelector("#coach-panel .panel-header > div:first-child");
  const coachHeaderKicker = coachHeaderBlock?.querySelector(".section-kicker");
  const coachHeaderTitle = coachHeaderBlock?.querySelector("h2");
  const coachHeaderCopy = coachHeaderBlock?.querySelector(".muted-copy");

  document.body.dataset.appMode = APP_MODE;
  if (els.mainTabsNav) {
    els.mainTabsNav.classList.toggle("is-hidden", APP_MODE !== "dual");
    els.mainTabsNav.hidden = APP_MODE !== "dual";
    els.mainTabsNav.style.display = APP_MODE === "dual" ? "" : "none";
  }

  if (coachPanel) {
    coachPanel.hidden = APP_MODE === "student";
    coachPanel.style.display = APP_MODE === "student" ? "none" : "";
  }
  if (studentPanel) {
    studentPanel.hidden = APP_MODE === "coach" || APP_MODE === "admin";
    studentPanel.style.display = APP_MODE === "coach" || APP_MODE === "admin" ? "none" : "";
  }

  const coachSessionInline = document.querySelector(".coach-session-inline");
  if (coachSessionInline) {
    const shouldShowCoachSwitcher = APP_MODE === "dual";
    coachSessionInline.hidden = !shouldShowCoachSwitcher;
    coachSessionInline.style.display = shouldShowCoachSwitcher ? "" : "none";
  }
  const coachSessionBlock = els.coachSessionName?.closest(".student-manager-block");
  if (coachSessionBlock) {
    const showCoachSessionBlock = APP_MODE !== "admin";
    coachSessionBlock.hidden = !showCoachSessionBlock;
    coachSessionBlock.style.display = showCoachSessionBlock ? "" : "none";
  }
  if (els.coachAuthShell) {
    const showCoachAuthShell = APP_MODE === "coach";
    els.coachAuthShell.hidden = !showCoachAuthShell;
    els.coachAuthShell.style.display = showCoachAuthShell ? "" : "none";
  }
  if (els.coachStudentLinksCard) {
    const showStudentLinks = APP_MODE !== "student";
    els.coachStudentLinksCard.hidden = !showStudentLinks;
    els.coachStudentLinksCard.style.display = showStudentLinks ? "" : "none";
  }
  if (els.historyImportCard) {
    const showImport = APP_MODE !== "admin";
    els.historyImportCard.hidden = !showImport;
    els.historyImportCard.style.display = showImport ? "" : "none";
  }
  if (els.openCoachLeaveSystemTab) {
    const showCoachLeaveTab = APP_MODE !== "admin";
    els.openCoachLeaveSystemTab.hidden = !showCoachLeaveTab;
    els.openCoachLeaveSystemTab.style.display = showCoachLeaveTab ? "" : "none";
  }

  applyCoachTabVisibility();

  if (APP_MODE === "coach" || APP_MODE === "admin") {
    switchMainPanel("coach-panel");
  } else if (APP_MODE === "student") {
    switchMainPanel("student-panel");
  }
  if (APP_MODE === "admin") {
    if (coachHeaderKicker) coachHeaderKicker.textContent = "管理者控制台";
    if (coachHeaderTitle) coachHeaderTitle.textContent = "人員管理者控制台";
    if (coachHeaderCopy) coachHeaderCopy.textContent = "管理教練、管理學生，並指派學生的主要教練。";
    switchCoachPanel("coach-coaches");
  } else {
    if (coachHeaderKicker) coachHeaderKicker.textContent = "教練工作台";
    if (coachHeaderTitle) coachHeaderTitle.textContent = "教練課表與紀錄管理";
    if (coachHeaderCopy) coachHeaderCopy.textContent = "建立課表、管理學生、查看今日紀錄與歷史資料。";
  }

  if (APP_MODE === "coach") {
    syncCoachAccessUI();
  }
  if (APP_MODE === "student") {
    syncStudentAccessUI();
  }
}

function handleStorageSync(event) {
  if (event.key !== STORAGE_KEY && event.key !== SESSION_KEY) {
    return;
  }
  if (hasUnsavedDraft()) {
    return;
  }

  const activeCoachPanel = els.coachPanels.find((panel) => panel.classList.contains("is-active"))?.id || (APP_MODE === "admin" ? "coach-coaches" : "coach-editor");
  const activeMainPanel = els.panels.find((panel) => panel.classList.contains("is-active"))?.id || (APP_MODE === "student" ? "student-panel" : "coach-panel");

  hydrateFromStorage();
  hydrateSession();
  cleanupRedundantBlankPrograms();

  if (APP_MODE === "admin") {
    refreshAdminWorkspace();
  } else if (APP_MODE === "coach") {
    refreshCoachWorkspace({ preserveEditor: coachEditorDirty });
  } else if (APP_MODE === "student") {
    refreshStudentWorkspace();
    if (currentStudentId) {
      loadStudentProgram({ silent: true, preserveIfDirty: true });
    }
  } else {
    refreshCoachWorkspace({ preserveEditor: coachEditorDirty });
    refreshStudentWorkspace();
  }

  switchMainPanel(activeMainPanel);
  switchCoachPanel(activeCoachPanel);
  if (APP_MODE === "coach") {
    syncCoachAccessUI();
  }
  if (APP_MODE === "student") {
    syncStudentAccessUI();
  }
  applyCoachViewMode();
  applyStudentViewMode();
  applyAppMode();
}

function switchCoachPanel(panelId) {
  const targetPanel = els.coachPanels.find((panel) => panel.id === panelId);
  if (targetPanel?.dataset.available === "false") {
    return;
  }
  els.coachTabs.forEach((button) => {
    button.classList.toggle("is-active", button.dataset.coachTab === panelId && button.style.display !== "none");
  });
  els.coachPanels.forEach((panel) => {
    const shouldActivate = panel.id === panelId && panel.dataset.available !== "false";
    panel.classList.toggle("is-active", shouldActivate);
  });
  if (panelId !== "coach-history") {
    coachHistoryOpened = false;
  }
  if (panelId === "coach-coaches") {
    renderCoachRoster();
  }
  if (panelId === "coach-students") {
    renderCoachStudentRoster();
  }
  applyCoachHistoryVisibility();
  if (APP_MODE === "coach") {
    syncCoachAccessUI();
  }
  if (APP_MODE === "student" && currentStudentId) {
    loadStudentProgram({ silent: true });
  }
}

async function syncCloudWorkspaceOnActivate() {
  if (!IS_CLOUD_MODE) {
    return;
  }

  const now = Date.now();
  if (now - lastCloudSyncAt < 1200) {
    return;
  }
  lastCloudSyncAt = now;

  try {
    if (APP_MODE === "admin") {
      const payload = await callCloudApiCritical("bootstrapAdmin", {}, "GET", 2);
      applyCloudPayloadToState(payload);
      refreshAdminWorkspace();
      return;
    }

    if (APP_MODE === "coach" && authenticatedCoachId) {
      const payload = await callCloudApiCritical("bootstrapCoach", { coachId: authenticatedCoachId }, "POST", 2);
      applyCloudPayloadToState(payload);
      refreshCoachWorkspace({ preserveEditor: coachEditorDirty });
      return;
    }

    if (APP_MODE === "student" && currentStudentId) {
      const payload = await callCloudApiCritical("bootstrapStudent", { studentId: currentStudentId }, "POST", 2);
      applyCloudPayloadToState(payload);
      refreshStudentWorkspace();
      loadStudentProgram({ silent: true, preserveIfDirty: true });
    }
  } catch (error) {
    console.warn("Cloud workspace sync failed:", error);
  }
}

function handleVisibilityRefresh() {
  if (document.visibilityState && document.visibilityState !== "visible") {
    tryPersistStudentProgress();
    return;
  }
  if (hasUnsavedDraft()) {
    return;
  }
  syncCloudWorkspaceOnActivate();
}

function hasUnsavedDraft() {
  if (APP_MODE === "coach") {
    return coachEditorDirty;
  }
  if (APP_MODE === "student") {
    return studentEntriesDirty;
  }
  if (APP_MODE === "dual") {
    const coachActive = Boolean(els.coachPanel?.classList.contains("is-active"));
    const studentActive = Boolean(els.studentPanel?.classList.contains("is-active"));
    if (coachActive && !studentActive) {
      return coachEditorDirty;
    }
    if (studentActive && !coachActive) {
      return studentEntriesDirty;
    }
    return coachEditorDirty || studentEntriesDirty;
  }
  return false;
}

function handleBeforeUnload(event) {
  if (!hasUnsavedDraft()) {
    return;
  }
  tryPersistStudentProgress();
  event.preventDefault();
  event.returnValue = "";
}

function tryPersistStudentProgress() {
  if (loadedStudentEntries.length) {
    captureStudentEntryDraftsFromDom();
  }
  if (els.studentProgramEditModal && !els.studentProgramEditModal.classList.contains("is-hidden")) {
    captureStudentProgramEditDraftFromDom();
  } else if (APP_MODE === "student" || APP_MODE === "dual") {
    persistSession();
  }
}

function handlePageHide() {
  tryPersistStudentProgress();
}

function handleCurrentCoachChange() {
  if (!els.currentCoachSelect?.value) {
    return;
  }
  state.currentCoachId = els.currentCoachSelect.value;
  persistState();
  renderCurrentCoachOptions();
  renderCoachRoster();
  coachProgramEditorMode = "create";
  editingProgramId = null;
  seedEditorFromProgram(null);
  renderProgramLibrary();
  renderCoachStudentLinks();
  renderCoachStudentRoster();
  renderCoachRoster();
  renderCoachToday();
  renderCoachHistoryFilters();
  renderCoachHistory();
  renderCoachRoster();
  renderStudentProgramOptions();
  renderStudentHistoryFilters();
  renderStudentHistory();
  renderStudentSummary();
}

function createProgramId() {
  return `program-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function createBlankProgram() {
  const today = getTodayDateInAppZone();
  const currentCoach = getCurrentCoach();
  const program = {
    id: createProgramId(),
    code: "",
    date: today,
    title: "",
    coachId: currentCoach?.id || "",
    coachName: currentCoach?.name || "",
    notes: "",
    published: false,
    targetStudentIds: [],
    createdAt: timestampNow()
  };
  state.programs.unshift(program);
  persistState();
  return program;
}

function renderProgramTargetStudents(selectedIds = []) {
  if (!els.programTargetStudents) {
    return;
  }

  const selectedSet = new Set(normalizeProgramStudentIds(selectedIds));
  const students = getCoachActiveStudents()
    .sort((a, b) => String(a.name || "").localeCompare(String(b.name || ""), "zh-Hant", { numeric: true, sensitivity: "base" }));

  els.programTargetStudents.innerHTML = students
    .map((student) => `
      <label class="program-target-option ${selectedSet.has(student.id) ? "is-selected" : ""}">
        <input type="checkbox" value="${escapeHtml(student.id)}" ${selectedSet.has(student.id) ? "checked" : ""}>
        <span class="program-target-check" aria-hidden="true"></span>
        <span>${escapeHtml(student.name)}</span>
      </label>
    `)
    .join("");
}

function getSelectedProgramTargetStudentIds() {
  if (!els.programTargetStudents) {
    return [];
  }
  return [...els.programTargetStudents.querySelectorAll("input[type='checkbox']:checked")]
    .map((input) => input.value)
    .filter(Boolean);
}

function seedEditorFromProgram(programId) {
  if (!els.programItemsBody || !els.programCode || !els.programDate || !els.coachName || !els.programNotes) {
    return;
  }

  if (!programId) {
    els.programCode.value = "";
    els.programDate.value = getTodayDateInAppZone();
    els.coachName.value = getCurrentCoach()?.name || "";
    els.programNotes.value = "";
    renderProgramTargetStudents([]);
    els.programItemsBody.innerHTML = "";
    return;
  }

  const program = state.programs.find((item) => item.id === programId) || createBlankProgram();
  const items = getProgramItems(program.id);

  els.programCode.value = program.code || "";
  els.programDate.value = normalizeDateKey(program.date || getTodayDateInAppZone()) || getTodayDateInAppZone();
  els.coachName.value = program.coachName || "";
  els.programNotes.value = program.notes || "";
  renderProgramTargetStudents(getProgramStudentIds(program));

  els.programItemsBody.innerHTML = "";
  if (items.length) {
    items.forEach((item) => addProgramItemRow(item));
  } else {
    addProgramItemRow();
    addProgramItemRow({ category: "\u4e3b\u9805\u76ee", targetSets: 5, targetType: "reps", targetValue: 5 });
    addProgramItemRow({ category: "\u8f14\u52a9\u9805", targetSets: 4, targetType: "time", targetValue: 60 });
    syncSortNumbers();
  }

  setCoachEditorDirty(false);
}

function seedEditorFromProgramCopy(programId) {
  if (!els.programItemsBody || !els.programCode || !els.programDate || !els.coachName || !els.programNotes) {
    return false;
  }

  const program = state.programs.find((item) => item.id === programId);
  if (!program) {
    return false;
  }

  const items = getProgramItems(program.id);
  coachProgramEditorMode = "create";
  editingProgramId = null;
  els.programCode.value = program.code || "";
  els.programDate.value = getTodayDateInAppZone();
  els.coachName.value = getCurrentCoach()?.name || program.coachName || "";
  els.programNotes.value = program.notes || "";
  renderProgramTargetStudents(getProgramStudentIds(program));
  els.programItemsBody.innerHTML = "";

  if (items.length) {
    items.forEach((item) => addProgramItemRow({ ...item, id: "", programId: "" }));
  } else {
    addProgramItemRow();
  }

  setCoachEditorDirty(true);
  syncEditorPreviewState();
  return true;
}

function addProgramItemRow(item = {}) {
  const fragment = els.programItemTemplate.content.cloneNode(true);
  const row = fragment.querySelector("tr");
  const category = fragment.querySelector(".category-input");
  const exercise = fragment.querySelector(".exercise-input");
  const sets = fragment.querySelector(".sets-input");
  const type = fragment.querySelector(".type-input");
  const value = fragment.querySelector(".value-input");
  const note = fragment.querySelector(".item-note-input");
  const sort = fragment.querySelector(".sort-input");

  category.value = item.category || "\u7206\u767c\u529b";
  exercise.value = item.exercise || "";
  sets.value = item.targetSets || 4;
  type.value = item.targetType || "reps";
  value.value = item.targetType === "rm" ? "" : (item.targetValue || 8);
  note.value = item.itemNote || "";
  sort.value = item.sortOrder || els.programItemsBody.children.length + 1;

  [category, exercise, sets, type, value, note, sort].forEach((input) => {
    input.addEventListener("input", () => {
      setCoachEditorDirty(true);
      syncSortNumbers();
      syncEditorPreviewState();
    });
    input.addEventListener("change", () => {
      setCoachEditorDirty(true);
      syncSortNumbers();
      syncEditorPreviewState();
    });
  });

  syncProgramItemRowState(row);
  type.addEventListener("change", () => {
    syncProgramItemRowState(row);
  });

  fragment.querySelector(".remove-item").addEventListener("click", () => {
    row.remove();
    setCoachEditorDirty(true);
    syncSortNumbers();
    syncEditorPreviewState();
  });

  els.programItemsBody.appendChild(fragment);
}

function syncProgramItemRowState(row) {
  const typeInput = row.querySelector(".type-input");
  const valueInput = row.querySelector(".value-input");

  if (typeInput.value === "rm") {
    valueInput.value = "";
    valueInput.disabled = true;
    valueInput.placeholder = "RM";
    return;
  }

  valueInput.disabled = false;
  valueInput.placeholder = typeInput.value === "time" ? "\u79d2\u6578" : "\u6b21\u6578";
  if (!valueInput.value) {
    valueInput.value = typeInput.value === "time" ? 60 : 8;
  }
}

function addStudentProgramEditRow(item = {}, sourceIndex = -1) {
  const useMobile = studentViewMode === "card";

  if (useMobile) {
    const card = document.createElement("article");
    card.className = "student-entry-card student-edit-mobile-card";
    if (Number.isInteger(sourceIndex) && sourceIndex >= 0) {
      card.dataset.sourceIndex = String(sourceIndex);
    }
    card.innerHTML = `
      <div class="student-entry-fields">
        <label class="mini-field">
          <span>\u5206\u985e</span>
          <select class="student-edit-category">
            <option value="\u7206\u767c\u529b">\u7206\u767c\u529b</option>
            <option value="\u4e3b\u9805\u76ee">\u4e3b\u9805\u76ee</option>
            <option value="\u8f14\u52a9\u9805">\u8f14\u52a9\u9805</option>
          </select>
        </label>
        <label class="mini-field">
          <span>\u52d5\u4f5c\u540d\u7a31</span>
          <input type="text" class="student-edit-exercise" placeholder="\u8acb\u8f38\u5165\u52d5\u4f5c\u540d\u7a31">
        </label>
        <div class="card-inline-fields">
          <label class="mini-field">
            <span>\u7d44\u6578</span>
            <input type="number" class="student-edit-sets" min="1" value="4">
          </label>
          <label class="mini-field">
            <span>\u985e\u578b</span>
            <select class="student-edit-type">
              <option value="reps">reps</option>
              <option value="time">time</option>
              <option value="rm">rm</option>
            </select>
          </label>
        </div>
        <label class="mini-field">
          <span>\u6b21\u6578</span>
          <input type="number" class="student-edit-value" min="1" value="8">
        </label>
        <button type="button" class="icon-button remove-student-edit-row">\u522a\u9664</button>
      </div>
    `;

    const category = card.querySelector(".student-edit-category");
    const exercise = card.querySelector(".student-edit-exercise");
    const sets = card.querySelector(".student-edit-sets");
    const type = card.querySelector(".student-edit-type");
    const value = card.querySelector(".student-edit-value");

    category.value = item.category || "\u4e3b\u9805\u76ee";
    exercise.value = item.exercise || "";
    sets.value = item.targetSets || 4;
    type.value = item.targetType || "reps";
    value.value = item.targetType === "rm" ? "" : (item.targetValue || 8);

    const sync = () => {
      if (type.value === "rm") {
        value.value = "";
        value.disabled = true;
        value.placeholder = "RM";
        return;
      }
      value.disabled = false;
      value.placeholder = type.value === "time" ? "\u79d2\u6578" : "\u6b21\u6578";
      if (!value.value) {
        value.value = type.value === "time" ? 60 : 8;
      }
    };

    type.addEventListener("change", sync);
    card.querySelector(".remove-student-edit-row").addEventListener("click", () => {
      card.remove();
      const count = getStudentProgramEditRowCount();
      showStudentProgramEditNotice(
        count > 0
          ? "已刪除 1 個課表項目。"
          : "已刪除最後一個項目，至少保留 1 筆才能套用。",
        count > 0 ? "info" : "warning"
      );
      captureStudentProgramEditDraftFromDom();
    });
    sync();
    els.studentEditMobileList.appendChild(card);
    return;
  }

  const fragment = els.studentProgramEditRowTemplate.content.cloneNode(true);
  const row = fragment.querySelector("tr");
  if (Number.isInteger(sourceIndex) && sourceIndex >= 0) {
    row.dataset.sourceIndex = String(sourceIndex);
  }
  const category = fragment.querySelector(".student-edit-category");
  const exercise = fragment.querySelector(".student-edit-exercise");
  const sets = fragment.querySelector(".student-edit-sets");
  const type = fragment.querySelector(".student-edit-type");
  const value = fragment.querySelector(".student-edit-value");

  category.value = item.category || "\u4e3b\u9805\u76ee";
  exercise.value = item.exercise || "";
  sets.value = item.targetSets || 4;
  type.value = item.targetType || "reps";
  value.value = item.targetType === "rm" ? "" : (item.targetValue || 8);

  const sync = () => {
    if (type.value === "rm") {
      value.value = "";
      value.disabled = true;
      value.placeholder = "RM";
      return;
    }

    value.disabled = false;
    value.placeholder = type.value === "time" ? "\u79d2\u6578" : "\u6b21\u6578";
    if (!value.value) {
      value.value = type.value === "time" ? 60 : 8;
    }
  };

  type.addEventListener("change", sync);
  fragment.querySelector(".remove-student-edit-row").addEventListener("click", () => {
    row.remove();
    const count = getStudentProgramEditRowCount();
    showStudentProgramEditNotice(
      count > 0
        ? "已刪除 1 個課表項目。"
        : "已刪除最後一個項目，至少保留 1 筆才能套用。",
      count > 0 ? "info" : "warning"
    );
    captureStudentProgramEditDraftFromDom();
  });
  sync();
  els.studentProgramEditBody.appendChild(fragment);
}

function getStudentProgramEditRowCount() {
  if (studentViewMode === "card") {
    return els.studentEditMobileList?.querySelectorAll(".student-edit-mobile-card").length || 0;
  }
  return els.studentProgramEditBody?.querySelectorAll("tr").length || 0;
}

function syncSortNumbers() {
  [...els.programItemsBody.querySelectorAll("tr")].forEach((row, index) => {
    const sortInput = row.querySelector(".sort-input");
    if (!Number(sortInput.value)) {
      sortInput.value = index + 1;
    }
  });
}

function collectProgramPayload(options = {}) {
  const existingProgram =
    coachProgramEditorMode === "edit" && editingProgramId
      ? state.programs.find((item) => item.id === editingProgramId) || null
      : null;
  const nextDate = normalizeDateKey(els.programDate.value);
  const existingDate = normalizeDateKey(existingProgram?.date || "");
  const dateChangedFromExisting = Boolean(existingProgram && nextDate && existingDate && nextDate !== existingDate);
  const shouldCreateNewProgram = !existingProgram || dateChangedFromExisting;
  const targetStudentIds = Object.prototype.hasOwnProperty.call(options, "targetStudentIds")
    ? normalizeProgramStudentIds(options.targetStudentIds)
    : getSelectedProgramTargetStudentIds();
  const program = {
    id: shouldCreateNewProgram ? createProgramId() : existingProgram.id,
    code: els.programCode.value.trim(),
    date: nextDate,
    coachId: getCurrentCoach()?.id || "",
    coachName: getCurrentCoach()?.name || els.coachName.value.trim(),
    notes: els.programNotes.value.trim(),
    published: shouldCreateNewProgram ? false : Boolean(existingProgram?.published),
    targetStudentIds,
    createdAt: shouldCreateNewProgram ? timestampNow() : existingProgram.createdAt || timestampNow()
  };

  const items = [...els.programItemsBody.querySelectorAll("tr")]
    .map((row, index) => ({
      id: row.dataset.itemId || `item-${program.id}-${index + 1}`,
      programId: program.id,
      sortOrder: Number(row.querySelector(".sort-input").value || index + 1),
      category: row.querySelector(".category-input").value,
      exercise: row.querySelector(".exercise-input").value.trim(),
      targetSets: Number(row.querySelector(".sets-input").value || 0),
      targetType: row.querySelector(".type-input").value,
      targetValue: row.querySelector(".type-input").value === "rm"
        ? 0
        : Number(row.querySelector(".value-input").value || 0),
      itemNote: row.querySelector(".item-note-input").value.trim()
    }))
    .filter((item) => item.exercise && item.targetSets > 0 && (item.targetType === "rm" || item.targetValue > 0))
    .sort((a, b) => a.sortOrder - b.sortOrder);

  return { program, items, existingProgram, dateChangedFromExisting };
}

async function saveProgram() {
  const { program, items, dateChangedFromExisting } = collectProgramPayload();

  if (!program.code || !program.date || !items.length) {
    window.alert("\u8acb\u5148\u5b8c\u6210\u8ab2\u8868\u4ee3\u78bc\u3001\u65e5\u671f\u8207\u81f3\u5c11\u4e00\u7b46\u8ab2\u8868\u9805\u76ee\u3002");
    return;
  }

  if (IS_CLOUD_MODE) {
    try {
      await saveProgramToCloud(program, items);
    } catch (error) {
      console.warn("Cloud saveProgram failed, falling back to local state:", error);
      upsertProgram(program);
      state.programItems = state.programItems.filter((item) => item.programId !== program.id).concat(items);
      persistState();
    }
  } else {
    upsertProgram(program);
    state.programItems = state.programItems.filter((item) => item.programId !== program.id).concat(items);
    persistState();
  }

  coachProgramEditorMode = "edit";
  editingProgramId = program.id;
  setCoachEditorDirty(false);
  renderCoachExerciseOptions();
  syncEditorPreviewState();
  renderProgramLibrary();
  renderCoachRoster();
  renderCoachHistoryFilters();
  renderCoachStudentLinks();
  renderStudentProgramOptions();
  renderStudentHistoryFilters();
  openProgramPreviewModal();
  if (dateChangedFromExisting) {
    window.alert("已另存為新的課表，原本的歷史課表已保留。");
  }
}

function applyProgramPublishState(program) {
  const publishedTargetIds = getProgramStudentIds(program);
  state.programs = state.programs.map((item) => {
    if ((item.coachId || "") !== (program.coachId || "")) {
      return item;
    }
    if (item.id === program.id) {
      return { ...item, published: true, targetStudentIds: publishedTargetIds };
    }

    const itemTargetIds = getProgramStudentIds(item);
    const shouldUnpublish = publishedTargetIds.length
      ? itemTargetIds.some((id) => publishedTargetIds.includes(id))
      : itemTargetIds.length === 0;
    return shouldUnpublish ? { ...item, published: false } : item;
  });
}

async function publishProgram() {
  await publishProgramWithTargets([]);
}

async function publishTargetedProgram() {
  const targetStudentIds = getSelectedProgramTargetStudentIds();
  if (!targetStudentIds.length) {
    window.alert("請先在「指定學生」勾選至少 1 位學生。");
    return;
  }
  await publishProgramWithTargets(targetStudentIds);
}

async function publishProgramWithTargets(targetStudentIds = []) {
  const { program, items, dateChangedFromExisting } = collectProgramPayload({ targetStudentIds });
  const isTargeted = getProgramStudentIds(program).length > 0;

  if (!program.code || !program.date || !items.length) {
    window.alert("\u8acb\u5148\u5b8c\u6210\u8ab2\u8868\u4ee3\u78bc\u3001\u65e5\u671f\u8207\u81f3\u5c11\u4e00\u7b46\u8ab2\u8868\u9805\u76ee\u3002");
    return;
  }

  if (IS_CLOUD_MODE) {
    try {
      await publishProgramToCloud(program, items);
    } catch (error) {
      console.warn("Cloud publishProgram failed, falling back to local state:", error);
      applyProgramPublishState(program);
      upsertProgram({ ...program, published: true });
      state.programItems = state.programItems.filter((item) => item.programId !== program.id).concat(items);
      persistState();
    }
  } else {
    applyProgramPublishState(program);
    upsertProgram({ ...program, published: true });
    state.programItems = state.programItems.filter((item) => item.programId !== program.id).concat(items);
    persistState();
  }

  coachProgramEditorMode = "edit";
  editingProgramId = program.id;
  setCoachEditorDirty(false);
  if (els.coachTodayDate && !els.coachTodayDate.value) {
    els.coachTodayDate.value = getTodayDateInAppZone();
  }
  renderCoachExerciseOptions();
  renderProgramTargetStudents(getProgramStudentIds(program));
  syncEditorPreviewState();
  renderProgramLibrary();
  renderCoachRoster();
  renderCoachToday();
  renderCoachHistoryFilters();
  renderCoachStudentLinks();
  renderStudentProgramOptions();
  renderStudentHistoryFilters();
  renderStudentSummary();
  window.alert(dateChangedFromExisting
    ? "已另存並發布為新的目前課表，原本的歷史課表已保留。"
    : isTargeted
      ? "課表已發布給指定學生。學生端會優先載入自己的專屬課表。"
      : "課表已發布為全班共用課表。沒有專屬課表的學生會載入這份課表。");
}

function upsertProgram(program) {
  const index = state.programs.findIndex((item) => item.id === program.id);
  if (index >= 0) {
    state.programs[index] = { ...state.programs[index], ...program };
  } else {
    state.programs.unshift(program);
  }
}

function renderProgramLibrary() {
  if (!els.programLibraryList) {
    return;
  }

  const sortBy = els.programLibrarySort?.value || "date-desc";
  const programs = [...getCoachScopedPrograms()].sort((a, b) => comparePrograms(a, b, sortBy));
  const keyword = (els.programLibrarySearch?.value || "").trim().toLowerCase();
  const filteredPrograms = programs.filter((program) => {
    if (!keyword) {
      return true;
    }
    return (
      String(program.code || "").toLowerCase().includes(keyword) ||
      String(program.date || "").toLowerCase().includes(keyword)
    );
  });

  if (els.programLibraryCount) {
    els.programLibraryCount.textContent = `共 ${filteredPrograms.length} / ${programs.length} 份`;
  }

  const visiblePrograms = filteredPrograms.slice(0, programLibraryVisibleCount);
  const hasMorePrograms = filteredPrograms.length > programLibraryVisibleCount;
  const canCollapsePrograms = filteredPrograms.length > 6 && programLibraryVisibleCount >= filteredPrograms.length;

  els.programLibraryList.innerHTML = visiblePrograms.length
    ? visiblePrograms
        .map((program) => {
          const items = getProgramItems(program.id);
          const itemCount = items.length;
          const targetNames = getProgramTargetStudentNames(program);
          const scopeText = targetNames.length ? `指定：${targetNames.join("、")}` : "全班共用";
          const itemPreview = items.length
            ? items
                .map((item) => `${item.category}｜${item.exercise}｜${formatTarget(item)}`)
                .join("<br>")
            : "目前沒有課表項目";
          const isExpanded = expandedProgramLibraryId === program.id;
          return `
            <article class="coach-link-card">
              <div class="coach-link-top">
                <div>
                  <h4>${program.code || "未命名課表"}</h4>
                  <p class="coach-student-meta">${formatDateDisplay(program.date)}｜${itemCount} 個項目</p>
                  <p class="coach-student-meta">${escapeHtml(scopeText)}</p>
                </div>
                <div class="roster-pill-stack">
                  <span class="status-pill ${program.published ? "is-success" : "is-muted"}">${program.published ? "目前課表" : "未發布"}</span>
                  <span class="status-pill ${isProgramTargeted(program) ? "is-success" : "is-muted"}">${isProgramTargeted(program) ? "指定學生" : "全班共用"}</span>
                </div>
              </div>
              <div class="coach-program-preview ${isExpanded ? "is-expanded" : ""}">${itemPreview}</div>
              <div class="coach-link-actions">
                <button class="ghost-button" type="button" data-toggle-program-preview="${program.id}">${isExpanded ? "收起項目" : "查看項目"}</button>
                <button class="ghost-button" type="button" data-copy-program="${program.id}">沿用成新課表</button>
                <button class="ghost-button" type="button" data-edit-existing-program="${program.id}">編輯原課表</button>
                <button class="ghost-button" type="button" data-publish-existing-program="${program.id}">設為目前課表</button>
                <button class="ghost-button danger-button" type="button" data-delete-existing-program="${program.id}">刪除課表</button>
              </div>
            </article>
          `;
        })
        .join("")
    : `<div class="empty-card">目前沒有符合條件的歷史課表。</div>`;

  if (els.programLibraryMore) {
    const shouldShowButton = filteredPrograms.length > 6;
    els.programLibraryMore.classList.toggle("is-hidden", !shouldShowButton);
    els.programLibraryMore.textContent = hasMorePrograms ? "顯示更多" : "收合";
  }
}

function toggleProgramLibraryVisibleCount() {
  const sortBy = els.programLibrarySort?.value || "date-desc";
  const keyword = (els.programLibrarySearch?.value || "").trim().toLowerCase();
  const filteredCount = [...state.programs]
    .sort((a, b) => comparePrograms(a, b, sortBy))
    .filter((program) => {
      if (!keyword) {
        return true;
      }
      return (
        String(program.code || "").toLowerCase().includes(keyword) ||
        String(program.date || "").toLowerCase().includes(keyword)
      );
    }).length;

  if (programLibraryVisibleCount >= filteredCount) {
    programLibraryVisibleCount = 6;
  } else {
    programLibraryVisibleCount += 6;
  }

  renderProgramLibrary();
}

function comparePrograms(a, b, sortBy) {
  const dateA = getComparableDateKey(a.date) || String(a.date || "");
  const dateB = getComparableDateKey(b.date) || String(b.date || "");
  const codeA = String(a.code || "");
  const codeB = String(b.code || "");
  const createdA = getComparableDateTimeKey(a.createdAt) || String(a.createdAt || "");
  const createdB = getComparableDateTimeKey(b.createdAt) || String(b.createdAt || "");

  if (sortBy === "date-asc") {
    const dateCompare = dateA.localeCompare(dateB);
    if (dateCompare !== 0) {
      return dateCompare;
    }
    return createdA.localeCompare(createdB);
  }

  if (sortBy === "code-asc") {
    const codeCompare = codeA.localeCompare(codeB, "zh-Hant", { numeric: true, sensitivity: "base" });
    if (codeCompare !== 0) {
      return codeCompare;
    }
    return dateB.localeCompare(dateA);
  }

  if (sortBy === "code-desc") {
    const codeCompare = codeB.localeCompare(codeA, "zh-Hant", { numeric: true, sensitivity: "base" });
    if (codeCompare !== 0) {
      return codeCompare;
    }
    return dateB.localeCompare(dateA);
  }

  const dateCompare = dateB.localeCompare(dateA);
  if (dateCompare !== 0) {
    return dateCompare;
  }
  return createdB.localeCompare(createdA);
}

async function handleProgramLibraryAction(event) {
  const button = event.target.closest("[data-toggle-program-preview], [data-copy-program], [data-load-program], [data-edit-existing-program], [data-publish-existing-program], [data-delete-existing-program]");
  if (!button) {
    return;
  }

  if (button.dataset.toggleProgramPreview) {
    const programId = button.dataset.toggleProgramPreview;
    expandedProgramLibraryId = expandedProgramLibraryId === programId ? null : programId;
    renderProgramLibrary();
    return;
  }

  if (button.dataset.copyProgram || button.dataset.loadProgram) {
    const programId = button.dataset.copyProgram || button.dataset.loadProgram;
    const copied = seedEditorFromProgramCopy(programId);
    if (!copied) {
      return;
    }
    switchCoachPanel("coach-editor");
    document.querySelector("#coach-editor")?.scrollIntoView({ behavior: "smooth", block: "start" });
    window.alert("已沿用為新的課表草稿，日期已帶入今天；儲存或發布後會新增一份課表，不會覆蓋原本歷史課表。");
    return;
  }

  if (button.dataset.editExistingProgram) {
    const programId = button.dataset.editExistingProgram;
    const ok = window.confirm("這會直接修改原本的歷史課表。若要做本週新課表，請使用「沿用成新課表」。確定要編輯原課表嗎？");
    if (!ok) {
      return;
    }
    coachProgramEditorMode = "edit";
    editingProgramId = programId;
    seedEditorFromProgram(programId);
    setCoachEditorDirty(false);
    syncEditorPreviewState();
    switchCoachPanel("coach-editor");
    document.querySelector("#coach-editor")?.scrollIntoView({ behavior: "smooth", block: "start" });
    return;
  }

  if (button.dataset.deleteExistingProgram) {
    const programId = button.dataset.deleteExistingProgram;
    const program = state.programs.find((item) => item.id === programId);
    if (!program) {
      return;
    }

    const confirmed = window.confirm(`確定要刪除「${program.code || "未命名課表"}」嗎？這會移除課表與課表項目，但不會刪除學生已送出的訓練紀錄。`);
    if (!confirmed) {
      return;
    }

    if (IS_CLOUD_MODE) {
      try {
        await deleteProgramFromCloud(programId);
      } catch (error) {
        console.warn("Cloud deleteProgram failed:", error);
        window.alert(error?.message || "雲端刪除課表失敗，請確認 Apps Script 已重新部署後再試。");
        return;
      }
    } else {
      state.programs = state.programs.filter((item) => item.id !== programId);
      state.programItems = state.programItems.filter((item) => item.programId !== programId);
      persistState();
    }

    if (editingProgramId === programId) {
      coachProgramEditorMode = "create";
      editingProgramId = null;
      seedEditorFromProgram(null);
      setCoachEditorDirty(false);
    }
    if (expandedProgramLibraryId === programId) {
      expandedProgramLibraryId = null;
    }

    renderCoachExerciseOptions();
    renderProgramLibrary();
    renderCoachToday();
    renderCoachStudentLinks();
    renderCoachStudentRoster();
    renderStudentProgramOptions();
    renderStudentHistoryFilters();
    renderStudentSummary();
    window.alert("課表已刪除，學生訓練紀錄已保留。");
    return;
  }

  const publishProgramId = button.dataset.publishExistingProgram;
  const program = state.programs.find((item) => item.id === publishProgramId);
  if (!program) {
    return;
  }

  applyProgramPublishState(program);
  if (IS_CLOUD_MODE) {
    try {
      await publishProgramToCloud(program, getProgramItems(program.id));
    } catch (error) {
      console.warn("Cloud publish existing program failed, falling back to local state:", error);
      persistState();
    }
  } else {
    persistState();
  }
  coachProgramEditorMode = "edit";
  editingProgramId = publishProgramId;
  if (els.coachTodayDate && !els.coachTodayDate.value) {
    els.coachTodayDate.value = getTodayDateInAppZone();
  }
  seedEditorFromProgram(publishProgramId);
  setCoachEditorDirty(false);
  syncEditorPreviewState();
  renderProgramLibrary();
  renderCoachToday();
  renderCoachStudentLinks();
  renderCoachStudentRoster();
  renderStudentProgramOptions();
  renderStudentSummary();
  switchCoachPanel("coach-editor");
  window.alert(isProgramTargeted(program)
    ? "已設為指定學生目前課表。"
    : "已設為全班共用目前課表。");
}

function renderPreview() {
  const { program, items } = collectProgramPayload();
  els.modalPreviewCode.textContent = program.code || "\u8acb\u5148\u8f38\u5165\u8ab2\u8868\u4ee3\u78bc";

  if (!items.length) {
    els.modalPreviewBody.innerHTML = `<div class="program-preview-empty">\u8acb\u5148\u8f38\u5165\u8ab2\u8868\u5167\u5bb9\u3002</div>`;
    return;
  }

  const previewMarkup = items
    .map(
      (item, index) => `
        <div class="preview-row ${index % 2 === 0 ? "is-accent" : ""}">
          <div class="preview-cell category">${item.category}</div>
          <div class="preview-cell exercise">${item.exercise}</div>
          <div class="preview-cell target">${formatTarget(item)}</div>
        </div>
      `
    )
    .join("");

  els.modalPreviewBody.innerHTML = previewMarkup;
}

function syncEditorPreviewState() {
  renderPreview();
}

function renderStudentProgramOptions() {
  if (!els.studentProgramSelect) {
    return;
  }
  const student = getSelectedStudent();
  if (student?.status === "inactive") {
    els.studentProgramSelect.innerHTML = `<option value="">\u8a72\u5b78\u751f\u5df2\u505c\u7528</option>`;
    currentStudentProgramId = "";
    return;
  }

  const studentCoachId = student?.primaryCoachId || "";
  const studentId = student?.id || "";
  const assignedCoach = state.coaches.find((coach) => coach.id === studentCoachId);
  const programs = [...state.programs]
    .filter((program) => !studentCoachId || (program.coachId || "") === studentCoachId)
    .filter((program) => !studentId || isProgramAvailableForStudent(program, studentId))
    .sort((a, b) => (getComparableDateKey(b.date) || "").localeCompare(getComparableDateKey(a.date) || ""));
  const currentValue = currentStudentProgramId || els.studentProgramSelect.value;

  els.studentProgramSelect.innerHTML = programs.length
    ? programs
        .map((program) => {
          const scopeLabel = isProgramTargeted(program) ? " | 專屬課表" : "";
          const label = `${formatDateDisplay(program.date)} | ${program.code || "\u76ee\u524d\u8ab2\u8868"}${assignedCoach?.name ? ` | ${assignedCoach.name}` : ""}${scopeLabel}${program.published ? " | \u76ee\u524d\u8ab2\u8868" : ""}`;
          return `<option value="${program.id}">${label}</option>`;
        })
        .join("")
    : `<option value="">${assignedCoach?.name ? `${assignedCoach.name}\u76ee\u524d\u6c92\u6709\u8ab2\u8868` : "\u76ee\u524d\u6c92\u6709\u8ab2\u8868"}</option>`;

  if (programs.length) {
    const published = getPublishedProgram(studentCoachId, studentId);
    const selectedProgram = programs.find((program) => program.id === currentValue);
    const fallbackProgram = published || selectedProgram || programs[0];
    currentStudentProgramId =
      published?.id
        || selectedProgram?.id
        || fallbackProgram.id;
    els.studentProgramSelect.value = currentStudentProgramId;
  } else {
    currentStudentProgramId = "";
  }
}

function renderStudentSummary() {
  const student = getSelectedStudent();
  const program = getSelectedStudentProgram();
  const assignedCoach = state.coaches.find((coach) => coach.id === (student?.primaryCoachId || ""));
  const assignedCoachName = assignedCoach?.name || "\u5c1a\u672a\u6307\u6d3e";

  if (!student) {
    els.studentSummary.innerHTML = `<p class="muted-copy">\u8acb\u5148\u8f38\u5165\u5c08\u5c6c\u4ee3\u78bc\uff0c\u6216\u76f4\u63a5\u4f7f\u7528\u4f60\u7684\u5c08\u5c6c\u9023\u7d50\u3002</p>`;
    els.studentActiveCopy.textContent = "";
    els.studentMobileSubmitBar.classList.remove("is-visible");
    els.studentMobileTools.classList.remove("is-visible");
    syncStudentLeaveSystemEntry();
    syncStudentAccessUI();
    return;
  }

  if (!program) {
    els.studentSummary.innerHTML = `
      <p><strong>${student.name}</strong></p>
      <p>\u6559\u7df4\uff1a${assignedCoachName}</p>
      <p>\u5c08\u5c6c\u4ee3\u78bc\uff1a${student.accessCode}</p>
      <p class="muted-copy">\u76ee\u524d\u9084\u6c92\u6709 ${assignedCoachName} \u7684\u53ef\u8f09\u5165\u8ab2\u8868\u3002</p>
    `;
    els.studentActiveCopy.textContent = `${student.name}\uff5c${assignedCoachName}`;
    els.studentMobileSubmitBar.classList.remove("is-visible");
    els.studentMobileTools.classList.remove("is-visible");
    syncStudentLeaveSystemEntry();
    syncStudentAccessUI();
    return;
  }

  els.studentSummary.innerHTML = `
    <p><strong>${student.name}</strong></p>
    <p>\u6559\u7df4\uff1a${assignedCoachName}</p>
    <p>\u5c08\u5c6c\u4ee3\u78bc\uff1a${student.accessCode}</p>
    <p>\u8ab2\u8868\uff1a${program.code || "\u76ee\u524d\u8ab2\u8868"}</p>
    <p>類型：${isProgramTargeted(program) ? "專屬課表" : "全班共用"}</p>
    <p>${formatDateDisplay(program.date)}${program.published ? "\uff5c\u5df2\u767c\u5e03" : "\uff5c\u672a\u767c\u5e03"}</p>
  `;
  els.studentActiveCopy.textContent = `${student.name}\uff5c${assignedCoachName}\uff5c${program.code || "\u76ee\u524d\u8ab2\u8868"}`;
  syncStudentLeaveSystemEntry();
  syncStudentAccessUI();
}

function buildStudentLeaveSystemUrl() {
  if (!IS_LEAVE_SANDBOX_ENABLED || !LEAVE_SANDBOX_STUDENT_PAGE) {
    return "";
  }
  const student = getSelectedStudent();
  if (!student) {
    return "";
  }

  const configuredPublicBase = String(APP_CONFIG.publicBaseUrl || "").trim();
  const isHosted = /^https?:/i.test(window.location.protocol);
  const baseHref = configuredPublicBase || (isHosted ? window.location.href : "");
  if (!baseHref) {
    return "";
  }

  try {
    const leaveUrl = new URL(LEAVE_SANDBOX_STUDENT_PAGE, baseHref);
    leaveUrl.search = "";
    const studentCode = String(student.accessCode || student.id || "").trim();
    if (studentCode) {
      leaveUrl.searchParams.set("studentCode", studentCode);
    }
    const assignedCoach = state.coaches.find((coach) => coach.id === (student.primaryCoachId || ""));
    const coachCode = String(assignedCoach?.accessCode || "").trim();
    if (coachCode) {
      leaveUrl.searchParams.set("coachCode", coachCode);
    }
    leaveUrl.searchParams.set("from", "coachflow-student");
    leaveUrl.searchParams.set("autoLogin", "1");
    leaveUrl.searchParams.set("v", PUBLIC_APP_VERSION);
    return leaveUrl.toString();
  } catch (error) {
    console.warn("Failed to build student leave url:", error);
    return "";
  }
}

function buildCoachLeaveSystemUrl(options = {}) {
  if (!IS_LEAVE_SANDBOX_ENABLED || !LEAVE_SANDBOX_COACH_PAGE) {
    return "";
  }

  const configuredPublicBase = String(APP_CONFIG.publicBaseUrl || "").trim();
  const isHosted = /^https?:/i.test(window.location.protocol);
  const candidateBaseHref = String(options.baseHref || "").trim();
  const baseHref = candidateBaseHref || configuredPublicBase || (isHosted ? window.location.href : "");
  if (!baseHref) {
    return "";
  }

  try {
    const leaveUrl = new URL(LEAVE_SANDBOX_COACH_PAGE, baseHref);
    leaveUrl.search = "";
    const fallbackCoachCode = String(
      options.forceCoachCode ||
      getCurrentCoach()?.accessCode ||
      authenticatedCoachAccess ||
      els.coachAccessCode?.value ||
      ""
    ).trim();
    if (fallbackCoachCode) {
      leaveUrl.searchParams.set("coachCode", fallbackCoachCode);
    }
    if (options.readOnly) {
      leaveUrl.searchParams.set("readonly", "1");
      leaveUrl.searchParams.set("mode", "readonly");
    }
    leaveUrl.searchParams.set("from", String(options.from || "coachflow-coach"));
    leaveUrl.searchParams.set("autoLogin", "1");
    leaveUrl.searchParams.set("v", PUBLIC_APP_VERSION);
    return leaveUrl.toString();
  } catch (error) {
    console.warn("Failed to build coach leave url:", error);
    return "";
  }
}

function saveLeavePrefillSession(payload = {}) {
  try {
    const data = {
      coachCode: String(payload.coachCode || "").trim(),
      studentCode: String(payload.studentCode || "").trim(),
      from: String(payload.from || "").trim(),
      createdAt: Date.now()
    };
    if (!data.coachCode && !data.studentCode) {
      return;
    }
    localStorage.setItem(LEAVE_PREFILL_STORAGE_KEY, JSON.stringify(data));
  } catch (error) {
    console.warn("Failed to persist leave prefill session:", error);
  }
}

function getActiveCoachAccessCode() {
  return String(
    getCurrentCoach()?.accessCode ||
    authenticatedCoachAccess ||
    els.coachAccessCode?.value ||
    ""
  ).trim();
}

function buildLeavePrefillPayload(options = {}) {
  const selectedStudent = getSelectedStudent();
  const studentCode = String(
    options.studentCode ||
    selectedStudent?.accessCode ||
    selectedStudent?.id ||
    ""
  ).trim();
  const coachCode = String(
    options.coachCode ||
    getActiveCoachAccessCode()
  ).trim();
  const from = String(options.from || "coachflow").trim();
  if (!coachCode && !studentCode) {
    return null;
  }
  return {
    coachCode,
    studentCode,
    from,
    createdAt: Date.now()
  };
}

function appendLeavePrefillToUrl(rawUrl, payload = {}) {
  if (!rawUrl) {
    return "";
  }
  const url = new URL(String(rawUrl), window.location.href);
  const coachCode = String(payload.coachCode || "").trim();
  const studentCode = String(payload.studentCode || "").trim();
  const from = String(payload.from || "coachflow").trim();
  if (coachCode) {
    url.searchParams.set("coachCode", coachCode);
    url.searchParams.set("code", coachCode);
  }
  if (studentCode) {
    url.searchParams.set("studentCode", studentCode);
  }
  url.searchParams.set("from", from);
  url.searchParams.set("autoLogin", "1");
  url.searchParams.set("v", PUBLIC_APP_VERSION);
  return url.toString();
}

function resolveCoachLeaveTargetUrl(rawUrl = "") {
  const fallbackUrl = buildCoachLeaveSystemUrl() || "";
  const sourceUrl = String(rawUrl || fallbackUrl || "").trim();
  if (!sourceUrl) {
    return "";
  }
  const payload = buildLeavePrefillPayload({ from: "coachflow-coach" });
  if (payload) {
    saveLeavePrefillSession(payload);
  }
  return appendLeavePrefillToUrl(sourceUrl, payload || {});
}

function syncStudentLeaveSystemEntry() {
  const url = buildStudentLeaveSystemUrl();
  const hasStudent = Boolean(getSelectedStudent());
  if (els.studentLeaveEntryRow) {
    if (hasStudent) {
      els.studentLeaveEntryRow.hidden = false;
      els.studentLeaveEntryRow.removeAttribute("hidden");
      els.studentLeaveEntryRow.classList.remove("is-hidden");
    } else {
      els.studentLeaveEntryRow.hidden = true;
      els.studentLeaveEntryRow.setAttribute("hidden", "hidden");
      els.studentLeaveEntryRow.classList.add("is-hidden");
    }
  }
  const buttons = [
    els.openStudentLeaveSystemMain,
    els.openStudentLeaveSystemInline
  ].filter(Boolean);
  buttons.forEach((button) => {
    const visible = hasStudent;
    if (visible) {
      button.hidden = false;
      button.removeAttribute("hidden");
      button.classList.remove("is-hidden");
    } else {
      button.hidden = true;
      button.setAttribute("hidden", "hidden");
      button.classList.add("is-hidden");
    }
    button.dataset.leaveUrl = url || "";
    if (visible) {
      button.removeAttribute("disabled");
    } else {
      button.setAttribute("disabled", "disabled");
    }
  });
  if (els.openStudentLeaveSystemMobileSecondary) {
    els.openStudentLeaveSystemMobileSecondary.hidden = true;
    els.openStudentLeaveSystemMobileSecondary.setAttribute("hidden", "hidden");
    els.openStudentLeaveSystemMobileSecondary.classList.add("is-hidden");
    els.openStudentLeaveSystemMobileSecondary.setAttribute("disabled", "disabled");
    els.openStudentLeaveSystemMobileSecondary.dataset.leaveUrl = "";
  }
  if (els.openStudentLeaveSystemMobilePrimary) {
    els.openStudentLeaveSystemMobilePrimary.hidden = true;
    els.openStudentLeaveSystemMobilePrimary.setAttribute("hidden", "hidden");
    els.openStudentLeaveSystemMobilePrimary.classList.add("is-hidden");
    els.openStudentLeaveSystemMobilePrimary.setAttribute("disabled", "disabled");
    els.openStudentLeaveSystemMobilePrimary.dataset.leaveUrl = "";
  }
}

function openStudentLeaveSystemEntry(event) {
  event?.preventDefault();
  const url = String(event?.currentTarget?.dataset?.leaveUrl || buildStudentLeaveSystemUrl() || "").trim();
  if (!url) {
    window.alert("請假系統入口尚未啟用，請先聯絡教練。");
    return;
  }
  const payload = buildLeavePrefillPayload({ from: "coachflow-student" });
  if (payload) {
    saveLeavePrefillSession(payload);
  }
  const targetUrl = appendLeavePrefillToUrl(url, payload || {});
  window.location.assign(targetUrl);
}

function openCoachLeaveSystemEntry(event) {
  event?.preventDefault();
  const targetUrl = resolveCoachLeaveTargetUrl(
    String(event?.currentTarget?.dataset?.leaveUrl || "").trim()
  );
  if (!targetUrl) {
    window.alert("請假系統入口尚未啟用，請先確認教練身份。");
    return;
  }
  window.open(targetUrl, "_blank", "noopener");
}

function renderCoachSummary() {
  if (!els.coachSummary) {
    return;
  }
  const coach = getCurrentCoach();
  if (!coach || APP_MODE !== "coach") {
    els.coachSummary.innerHTML = `<p class="muted-copy">請先輸入教練代碼，或直接使用你的專屬連結。</p>`;
    return;
  }

  els.coachSummary.innerHTML = `
    <p><strong>${coach.name}</strong></p>
    <p>教練代碼：${coach.accessCode || "-"}</p>
    <p>身份已確認，已載入你的教練工作台。</p>
  `;
}

function resetCoachAccess() {
  authenticatedCoachId = "";
  authenticatedCoachAccess = "";
  setCoachEditorDirty(false);
  persistSession();
  clearModeAccessInUrl();
  if (els.coachAccessCode) {
    els.coachAccessCode.value = "";
  }
  renderCoachSummary();
  refreshCoachWorkspace();
  if (els.coachAccessCode) {
    els.coachAccessCode.focus();
  }
}

function syncCoachAccessUI() {
  if (APP_MODE !== "coach") {
    return;
  }

  const hasCoach = Boolean(getCurrentCoach());
  const coachAuthShell = els.coachAuthShell;
  const coachHeader = document.querySelector("#coach-panel .panel-header");
  const coachTabs = document.querySelector("#coach-panel .sub-tabs");
  const coachPanelsWrap = Array.from(document.querySelectorAll("#coach-panel .sub-panel"));
  const coachActiveBar = els.coachActiveBar;

  if (coachAuthShell) {
    coachAuthShell.hidden = hasCoach;
    coachAuthShell.style.display = hasCoach ? "none" : "";
    coachAuthShell.classList.toggle("is-collapsed", hasCoach);
  }

  if (coachHeader) {
    coachHeader.style.display = hasCoach ? "" : "none";
  }

  if (coachActiveBar) {
    coachActiveBar.hidden = !hasCoach;
    coachActiveBar.classList.toggle("is-hidden", !hasCoach);
    coachActiveBar.style.display = hasCoach ? "" : "none";
  }

  if (els.coachActiveCopy) {
    const coach = getCurrentCoach();
    els.coachActiveCopy.textContent = hasCoach && coach
      ? `${coach.name}｜${coach.accessCode || "-"}`
      : "";
  }

  if (coachTabs) {
    coachTabs.style.display = hasCoach ? "" : "none";
  }
  if (els.openCoachLeaveSystemTab) {
    const leaveUrl = hasCoach ? buildCoachLeaveSystemUrl({ from: "coachflow-coach" }) : "";
    els.openCoachLeaveSystemTab.dataset.leaveUrl = leaveUrl;
    els.openCoachLeaveSystemTab.hidden = !hasCoach;
    els.openCoachLeaveSystemTab.style.display = hasCoach ? "" : "none";
    if (hasCoach && leaveUrl) {
      els.openCoachLeaveSystemTab.removeAttribute("disabled");
    } else {
      els.openCoachLeaveSystemTab.setAttribute("disabled", "disabled");
    }
  }

  coachPanelsWrap.forEach((panel) => {
    panel.style.display = hasCoach && panel.classList.contains("is-active") ? "" : "none";
    panel.hidden = !hasCoach || !panel.classList.contains("is-active");
  });
}

async function confirmCoachAccess(options = {}) {
  const { silentError = false } = options;
  const showBusy = !silentError;
  if (showBusy) {
    if (coachLoginBusy) {
      return false;
    }
    coachLoginBusy = true;
    setButtonBusy(els.confirmCoachAccess, true, "登入中...");
    setControlDisabled(els.coachAccessCode, true);
    if (els.coachSummary) {
      els.coachSummary.innerHTML = `<p class="muted-copy">正在確認教練身份，請稍候...</p>`;
    }
  }

  try {
    return await confirmCoachAccessCore(options);
  } finally {
    if (showBusy) {
      coachLoginBusy = false;
      setButtonBusy(els.confirmCoachAccess, false);
      setControlDisabled(els.coachAccessCode, false);
    }
  }
}

async function confirmCoachAccessCore(options = {}) {
  const { silentError = false, skipTouch = false } = options;
  const accessValue = normalizeAccessInput(els.coachAccessCode?.value || "");
  if (els.coachAccessCode) {
    els.coachAccessCode.value = accessValue;
  }
  let matchedCoach = null;
  if (IS_CLOUD_MODE) {
    try {
      matchedCoach = await resolveCoachAccessFromCloud(accessValue);
    } catch (error) {
      console.warn("Coach cloud login failed, falling back to local state:", error);
      matchedCoach = resolveCoachByAccess(accessValue);
    }
  } else {
    matchedCoach = resolveCoachByAccess(accessValue);
  }
  authenticatedCoachId = matchedCoach?.id || "";

  if (!matchedCoach) {
    authenticatedCoachAccess = "";
    persistSession();
    clearModeAccessInUrl();
    renderCoachSummary();
    if (!silentError) {
      window.alert("找不到這組教練代碼，請重新確認後再試一次。");
    }
    return false;
  }

  if (matchedCoach.status === "inactive") {
    authenticatedCoachId = "";
    authenticatedCoachAccess = "";
    persistSession();
    renderCoachSummary();
    if (!silentError) {
      window.alert("這位教練帳號目前已停用。");
    }
    return false;
  }

  if (els.coachAccessCode) {
    els.coachAccessCode.value = matchedCoach.accessCode || els.coachAccessCode.value;
  }
  authenticatedCoachAccess = matchedCoach.accessCode || String(accessValue || "").trim();
  state.currentCoachId = matchedCoach.id;
  if (IS_CLOUD_MODE && !skipTouch) {
    callCloudApiCritical("touchCoach", { coachId: matchedCoach.id }, "POST", 2)
      .then((payload) => applyCloudPayloadToState(payload))
      .catch(() => markCoachUsed(matchedCoach.id));
  } else if (!skipTouch) {
    markCoachUsed(matchedCoach.id);
  }
  setCoachEditorDirty(false);
  refreshCoachWorkspace();
  persistSession();
  updateModeAccessInUrl("coach", authenticatedCoachAccess);
  return true;
}

async function confirmStudentAccess(options = {}) {
  const { silentError = false } = options;
  const showBusy = !silentError;
  if (showBusy) {
    if (studentLoginBusy) {
      return false;
    }
    studentLoginBusy = true;
    setButtonBusy(els.confirmStudentAccess, true, "載入中...");
    setControlDisabled(els.studentAccessCode, true);
    if (els.studentProgramStatus) {
      els.studentProgramStatus.textContent = "正在確認身份並載入資料...";
    }
    if (els.studentSummary) {
      els.studentSummary.innerHTML = `<p class="muted-copy">正在確認身份，請稍候...</p>`;
    }
  }

  try {
    return await confirmStudentAccessCore(options);
  } finally {
    if (showBusy) {
      studentLoginBusy = false;
      setButtonBusy(els.confirmStudentAccess, false);
      setControlDisabled(els.studentAccessCode, false);
    }
  }
}

async function confirmStudentAccessCore(options = {}) {
  const { silentError = false, skipTouch = false } = options;
  const accessValue = normalizeAccessInput(els.studentAccessCode.value);
  const normalizedAccessValue = accessValue;
  els.studentAccessCode.value = normalizedAccessValue;
  let matchedStudent = null;
  if (IS_CLOUD_MODE) {
    try {
      matchedStudent = await resolveStudentAccessFromCloud(normalizedAccessValue);
    } catch (error) {
      console.warn("Student cloud login failed, falling back to local state:", error);
      matchedStudent = resolveStudentByAccess(normalizedAccessValue);
    }
  } else {
    matchedStudent = resolveStudentByAccess(normalizedAccessValue);
  }
  currentStudentId = matchedStudent?.id || "";

  if (!matchedStudent) {
    currentStudentAccess = "";
    persistSession();
    clearModeAccessInUrl();
    els.studentProgramStatus.textContent = "\u8eab\u4efd\u672a\u78ba\u8a8d";
    els.studentProgramBody.innerHTML = `<tr><td colspan="4" class="empty-state">\u8acb\u5148\u8f38\u5165\u6b63\u78ba\u7684\u5c08\u5c6c\u4ee3\u78bc\uff0c\u518d\u8f09\u5165\u8ab2\u8868\u3002</td></tr>`;
    els.studentCardList.innerHTML = `<div class="empty-card">\u8acb\u5148\u8f38\u5165\u6b63\u78ba\u7684\u5c08\u5c6c\u4ee3\u78bc\uff0c\u518d\u8f09\u5165\u8ab2\u8868\u3002</div>`;
    renderStudentSummary();
    renderStudentHistoryFilters();
    renderStudentHistory();
    if (!silentError) {
      window.alert("\u627e\u4e0d\u5230\u9019\u7d44\u5c08\u5c6c\u4ee3\u78bc\uff0c\u8acb\u91cd\u65b0\u78ba\u8a8d\u5f8c\u518d\u8a66\u4e00\u6b21\u3002");
    }
    return false;
  }

  if (matchedStudent.status === "inactive") {
    currentStudentId = "";
    currentStudentAccess = "";
    persistSession();
    clearModeAccessInUrl();
    if (!silentError) {
      window.alert("\u9019\u4f4d\u5b78\u751f\u5e33\u865f\u76ee\u524d\u5df2\u505c\u7528\uff0c\u8acb\u806f\u7d61\u6559\u7df4\u3002");
    }
    return false;
  }

  const matchedCode = String(matchedStudent.accessCode || "").trim();
  const matchedToken = String(matchedStudent.token || "").trim();
  els.studentAccessCode.value = matchedCode || matchedToken || normalizedAccessValue;
  currentStudentAccess = matchedCode || normalizedAccessValue || matchedToken;
  if (IS_CLOUD_MODE && !skipTouch) {
    callCloudApiCritical("touchStudent", { studentId: matchedStudent.id }, "POST", 2)
      .then((payload) => applyCloudPayloadToState(payload))
      .catch(() => markStudentUsed(matchedStudent.id));
  } else if (!skipTouch) {
    markStudentUsed(matchedStudent.id);
  }
  els.studentProgramStatus.textContent = `${matchedStudent.name} \u8eab\u4efd\u5df2\u78ba\u8a8d`;
  if (els.studentPanelViewToggle) {
    els.studentPanelViewToggle.hidden = false;
    els.studentPanelViewToggle.classList.remove("is-hidden");
    els.studentPanelViewToggle.style.display = "";
  }
  renderStudentProgramOptions();
  renderStudentSummary();
  renderStudentHistoryFilters();
  renderStudentHistory();
  syncStudentAccessUI();
  persistSession();
  updateModeAccessInUrl("student", currentStudentAccess);
  return true;
}

async function confirmStudentAccessAndLoadProgram() {
  const isConfirmed = await confirmStudentAccess();
  if (!isConfirmed) {
    return;
  }
  await loadStudentProgramWithFeedback();
}

async function loadStudentProgramWithFeedback(options = {}) {
  if (studentLoadBusy) {
    return;
  }

  studentLoadBusy = true;
  setButtonsBusy([els.loadStudentProgramInline, els.loadStudentProgramMobile], true, "載入中...");
  if (els.studentProgramStatus) {
    els.studentProgramStatus.textContent = "正在載入本輪課表...";
  }

  try {
    await delay(0);
    loadStudentProgram(options);
  } finally {
    studentLoadBusy = false;
    setButtonsBusy([els.loadStudentProgramInline, els.loadStudentProgramMobile], false);
  }
}

function loadStudentProgram(options = {}) {
  const { silent = false, preserveIfDirty = false } = options;
  const program = getSelectedStudentProgram();
  const student = getSelectedStudent();

  if (!student) {
    if (!silent) {
      window.alert("\u8acb\u5148\u78ba\u8a8d\u5b78\u751f\u8eab\u4efd\u3002");
    }
    return;
  }

  if (!program) {
    els.studentProgramStatus.textContent = "\u6c92\u6709\u53ef\u8f09\u5165\u8ab2\u8868";
    els.studentProgramBody.innerHTML = `<tr><td colspan="4" class="empty-state">\u76ee\u524d\u6c92\u6709\u53ef\u4f7f\u7528\u7684\u8ab2\u8868\u3002</td></tr>`;
    els.studentCardList.innerHTML = `<div class="empty-card">\u76ee\u524d\u6c92\u6709\u53ef\u4f7f\u7528\u7684\u8ab2\u8868\u3002</div>`;
    els.studentMobileSubmitBar.classList.remove("is-visible");
    els.studentMobileTools.classList.remove("is-visible");
    return;
  }

  const nextProgramId = String(program.id || "");
  if (
    preserveIfDirty &&
    studentEntriesDirty &&
    loadedStudentProgramId &&
    loadedStudentProgramId === nextProgramId &&
    loadedStudentEntries.length
  ) {
    renderStudentProgramEntries();
    els.studentProgramStatus.textContent = `${student.name}｜已載入 ${program.code || "目前課表"}`;
    els.studentMobileSubmitBar.classList.toggle("is-visible", studentViewMode === "card" && loadedStudentEntries.length > 0);
    els.studentMobileTools.classList.toggle("is-visible", studentViewMode === "card");
    syncStudentAccessUI();
    return;
  }

  const baseEntries = getProgramItems(program.id).map((item) => ({
    itemId: item.id,
    category: item.category,
    exercise: item.exercise,
    targetSets: item.targetSets,
    targetType: item.targetType,
    targetValue: item.targetValue,
    itemNote: item.itemNote,
    referenceLog: findLatestReferenceLog(student.id, item),
    actualWeight: "",
    actualSets: item.targetType === "time" || item.targetType === "reps" ? String(item.targetSets ?? "") : "",
    actualReps: item.targetType === "reps" ? String(item.targetValue ?? "") : "",
    studentNote: ""
  }));

  const draftSession = getSessionStudentDraft(student.id, nextProgramId);
  if (draftSession?.entries?.length) {
    loadedStudentEntries = sanitizeStudentDraftEntries(draftSession.entries).map((entry) => ({
      ...entry,
      referenceLog: findLatestReferenceLog(student.id, entry)
    }));
    studentEntriesDirty = Boolean(loadedStudentEntries.length && draftSession.dirty !== false);
  } else {
    loadedStudentEntries = baseEntries;
    studentEntriesDirty = false;
  }

  studentHistoryOpened = false;
  studentSubmissionCompleted = false;
  lastSubmittedLogs = [];
  loadedStudentProgramId = nextProgramId;
  currentStudentProgramId = nextProgramId;
  persistSession();
  renderStudentProgramEntries();
  els.studentProgramStatus.textContent = draftSession?.entries?.length
    ? `${student.name}｜已恢復未送出填寫內容`
    : `${student.name}｜已載入 ${program.code || "目前課表"}`;
  els.studentMobileSubmitBar.classList.toggle("is-visible", studentViewMode === "card" && loadedStudentEntries.length > 0);
  els.studentMobileTools.classList.toggle("is-visible", studentViewMode === "card");
  syncStudentAccessUI();
}

function buildEntryFields(entry, index) {
  const recommendation = getRecommendedWeightText(entry);
  const actualWeight = escapeHtml(String(entry.actualWeight ?? ""));
  const actualSets = escapeHtml(String(entry.actualSets ?? entry.targetSets ?? ""));
  const actualReps = escapeHtml(String(entry.actualReps ?? entry.targetValue ?? ""));
  const noteValue = escapeHtml(String(entry.studentNote ?? ""));
  const actualSetsPlaceholder = escapeHtml(String(entry.referenceLog?.actualSets || ""));
  const actualWeightPlaceholder = escapeHtml(String(entry.referenceLog?.actualWeight || "kg"));
  if (entry.targetType === "time") {
    return `
      <div class="field-stack">
        <label class="mini-field">
          <span>\u5be6\u969b\u7d44\u6578</span>
          <input type="number" min="0" value="${actualSets}" placeholder="${actualSetsPlaceholder}" data-entry-index="${index}" data-field="actualSets">
        </label>
        <label class="mini-field">
          <span>\u5099\u8a3b</span>
          <input type="text" value="${noteValue}" placeholder="\u72c0\u6cc1\u3001\u611f\u53d7" data-entry-index="${index}" data-field="studentNote">
        </label>
      </div>
    `;
  }

  if (entry.targetType === "rm") {
    return `
      <div class="field-stack">
        <label class="mini-field">
          <span>\u5be6\u969b\u91cd\u91cf</span>
          <input type="number" min="0" step="0.5" value="${actualWeight}" placeholder="${actualWeightPlaceholder}" data-entry-index="${index}" data-field="actualWeight">
          ${recommendation ? `<small class="field-hint recommendation-hint">${recommendation}</small>` : ""}
        </label>
        <label class="mini-field">
          <span>\u5099\u8a3b</span>
          <input type="text" value="${noteValue}" placeholder="\u72c0\u6cc1\u3001\u611f\u53d7" data-entry-index="${index}" data-field="studentNote">
        </label>
      </div>
    `;
  }

  return `
    <div class="field-grid">
      <label class="mini-field">
        <span>\u91cd\u91cf</span>
        <input type="number" min="0" step="0.5" value="${actualWeight}" placeholder="${actualWeightPlaceholder}" data-entry-index="${index}" data-field="actualWeight">
        ${recommendation ? `<small class="field-hint recommendation-hint">${recommendation}</small>` : ""}
      </label>
      <label class="mini-field">
        <span>\u7d44\u6578</span>
        <input type="number" min="0" value="${actualSets}" data-entry-index="${index}" data-field="actualSets">
      </label>
      <label class="mini-field">
        <span>\u6b21\u6578</span>
        <input type="number" min="0" value="${actualReps}" data-entry-index="${index}" data-field="actualReps">
      </label>
      <label class="mini-field span-full">
        <span>\u5099\u8a3b</span>
        <input type="text" value="${noteValue}" placeholder="\u72c0\u6cc1\u3001\u611f\u53d7" data-entry-index="${index}" data-field="studentNote">
      </label>
    </div>
  `;
}

function buildCardEntryFields(entry, index) {
  const recommendation = getRecommendedWeightText(entry);
  const actualWeight = escapeHtml(String(entry.actualWeight ?? ""));
  const actualSets = escapeHtml(String(entry.actualSets ?? entry.targetSets ?? ""));
  const actualReps = escapeHtml(String(entry.actualReps ?? entry.targetValue ?? ""));
  const noteValue = escapeHtml(String(entry.studentNote ?? ""));
  const actualSetsPlaceholder = escapeHtml(String(entry.referenceLog?.actualSets || ""));
  const actualWeightPlaceholder = escapeHtml(String(entry.referenceLog?.actualWeight || "kg"));
  if (entry.targetType === "time") {
    return `
      <label class="mini-field">
        <span>\u5be6\u969b\u7d44\u6578</span>
        <input type="number" min="0" value="${actualSets}" placeholder="${actualSetsPlaceholder}" data-entry-index="${index}" data-field="actualSets">
      </label>
      <label class="mini-field">
        <span>\u5099\u8a3b</span>
        <input type="text" value="${noteValue}" placeholder="\u72c0\u6cc1\u3001\u611f\u53d7" data-entry-index="${index}" data-field="studentNote">
      </label>
    `;
  }

  if (entry.targetType === "rm") {
    return `
      <label class="mini-field">
        <span>\u5be6\u969b\u91cd\u91cf</span>
        <input type="number" min="0" step="0.5" value="${actualWeight}" placeholder="${actualWeightPlaceholder}" data-entry-index="${index}" data-field="actualWeight">
        ${recommendation ? `<small class="field-hint recommendation-hint">${recommendation}</small>` : ""}
      </label>
      <label class="mini-field">
        <span>\u5099\u8a3b</span>
        <input type="text" value="${noteValue}" placeholder="\u72c0\u6cc1\u3001\u611f\u53d7" data-entry-index="${index}" data-field="studentNote">
      </label>
    `;
  }

  return `
    <label class="mini-field">
      <span>\u91cd\u91cf</span>
      <input type="number" min="0" step="0.5" value="${actualWeight}" placeholder="${actualWeightPlaceholder}" data-entry-index="${index}" data-field="actualWeight">
      ${recommendation ? `<small class="field-hint recommendation-hint">${recommendation}</small>` : ""}
    </label>
    <div class="card-inline-fields">
      <label class="mini-field">
        <span>\u7d44\u6578</span>
        <input type="number" min="0" value="${actualSets}" data-entry-index="${index}" data-field="actualSets">
      </label>
      <label class="mini-field">
        <span>\u6b21\u6578</span>
        <input type="number" min="0" value="${actualReps}" data-entry-index="${index}" data-field="actualReps">
      </label>
    </div>
    <label class="mini-field">
      <span>\u5099\u8a3b</span>
      <input type="text" value="${noteValue}" placeholder="\u72c0\u6cc1\u3001\u611f\u53d7" data-entry-index="${index}" data-field="studentNote">
    </label>
  `;
}

function formatLogForSummary(log) {
  return `
    <article class="confirm-entry-card">
      <div class="confirm-entry-top">
        <span class="confirm-entry-category">${log.category}</span>
        <span class="confirm-entry-target">${formatTarget(log)}</span>
      </div>
      <h4>${log.exercise}</h4>
      <p class="confirm-entry-result">${formatActual(log)}</p>
      <p class="confirm-entry-note">${log.studentNote || "-"}</p>
    </article>
  `;
}

function renderStudentProgramEntries() {
  els.studentProgramBody.innerHTML = "";
  els.studentCardList.innerHTML = "";
  els.studentRecordedBanner.classList.toggle("is-hidden", !studentSubmissionCompleted);
  els.editStudentProgram.textContent = studentSubmissionCompleted ? "\u67e5\u770b\u672c\u6b21\u8a13\u7df4\u7d00\u9304" : "\u7de8\u8f2f\u8ab2\u8868\u9805\u76ee";
  if (els.editStudentProgramMobileTop) {
    els.editStudentProgramMobileTop.textContent = studentSubmissionCompleted ? "\u67e5\u770b\u672c\u6b21\u8a13\u7df4\u7d00\u9304" : "\u7de8\u8f2f\u8ab2\u8868\u9805\u76ee";
  }
  els.editStudentProgramMobile.textContent = studentSubmissionCompleted ? "\u67e5\u770b\u672c\u6b21\u8a13\u7df4\u7d00\u9304" : "\u7de8\u8f2f\u8ab2\u8868\u9805\u76ee";

  if (!loadedStudentEntries.length) {
    const completedLogs = lastSubmittedLogs.length ? lastSubmittedLogs : getLatestSubmittedLogsForCurrentStudentProgram();
    if (studentSubmissionCompleted && completedLogs.length) {
      els.studentProgramBody.innerHTML = completedLogs
        .map(
          (log) => `
            <tr>
              <td>${log.category}</td>
              <td>${log.exercise}</td>
              <td>${formatTarget(log)}</td>
              <td>${formatActual(log)}</td>
            </tr>
          `
        )
        .join("");
      els.studentCardList.innerHTML = `
        <div class="empty-card success-state-card">\u8ab2\u8868\u5df2\u5b8c\u6210</div>
        ${completedLogs.map((log) => formatLogForSummary(log)).join("")}
      `;
    } else {
      els.studentProgramBody.innerHTML = `<tr><td colspan="4" class="empty-state">\u8acb\u5148\u8f09\u5165\u4e00\u4efd\u8ab2\u8868\u3002</td></tr>`;
      els.studentCardList.innerHTML = `<div class="empty-card">\u8acb\u5148\u8f09\u5165\u4e00\u4efd\u8ab2\u8868\u3002</div>`;
    }
    return;
  }

  loadedStudentEntries.forEach((entry, index) => {
    const fragment = els.studentEntryTemplate.content.cloneNode(true);
    const row = fragment.querySelector("tr");
    row.dataset.index = index;
    fragment.querySelector(".entry-category").textContent = entry.category;
    fragment.querySelector(".entry-exercise").textContent = entry.exercise;
    fragment.querySelector(".entry-target").textContent = formatTarget(entry);
    fragment.querySelector(".entry-fields").innerHTML = buildEntryFields(entry, index);
    els.studentProgramBody.appendChild(fragment);

    const card = document.createElement("article");
    card.className = "student-entry-card";
    card.innerHTML = `
      <div class="student-entry-top">
        <span class="student-entry-category">${entry.category}</span>
        <span class="student-entry-target">${formatTarget(entry)}</span>
      </div>
      <h4>${entry.exercise}</h4>
      <div class="student-entry-fields">
        ${buildCardEntryFields(entry, index)}
      </div>
    `;
    els.studentCardList.appendChild(card);
  });
}

function openSubmissionConfirm() {
  const student = getSelectedStudent();
  const program = getSelectedStudentProgram();

  if (!student || !program) {
    window.alert("\u8acb\u5148\u78ba\u8a8d\u5b78\u751f\u8eab\u4efd\u8207\u8ab2\u8868\u3002");
    return;
  }

  if (!loadedStudentEntries.length) {
    window.alert("\u8acb\u5148\u8f09\u5165\u4e00\u4efd\u8ab2\u8868\uff0c\u518d\u78ba\u8a8d\u672c\u6b21\u586b\u5beb\u5167\u5bb9\u3002");
    return;
  }

  if (!validateSubmissionInputs()) {
    els.studentProgramStatus.textContent = "\u8acb\u5148\u5b8c\u6210\u5fc5\u586b\u6b04\u4f4d";
    return;
  }

  const submissionDate = getTodayDateInAppZone();
  const alreadySubmitted = state.workoutLogs.some(
    (log) => log.studentId === student.id
      && log.programId === program.id
      && getLogActivityDateKey(log) === submissionDate
  );
  if (alreadySubmitted) {
    els.studentProgramStatus.textContent = "\u672c\u6b21\u9001\u51fa\u5c07\u8986\u84cb\u820a\u7d00\u9304";
  }

  pendingSubmission = loadedStudentEntries.map((entry, index) => buildLogPayload(student, program, entry, index));
  els.confirmSummary.textContent = alreadySubmitted
    ? `${student.name}\uff5c${formatDateDisplay(submissionDate)}\uff5c${program.code || "\u76ee\u524d\u8ab2\u8868"}\uff0c\u78ba\u8a8d\u5f8c\u6703\u8986\u84cb\u4f60\u4eca\u5929\u7684\u540c\u4efd\u8ab2\u8868\u7d00\u9304\u3002`
    : `${student.name}\uff5c${formatDateDisplay(submissionDate)}\uff5c${program.code || "\u76ee\u524d\u8ab2\u8868"}\uff0c\u8acb\u78ba\u8a8d\u9019\u6b21\u586b\u5beb\u5167\u5bb9\u3002`;
  els.confirmCardList.innerHTML = pendingSubmission.map((log) => formatLogForSummary(log)).join("");
  els.confirmBody.innerHTML = pendingSubmission
    .map(
      (log) => `
        <tr>
          <td>${log.category}</td>
          <td>${log.exercise}</td>
          <td>${formatTarget(log)}</td>
          <td>${formatActual(log)}</td>
          <td>${log.studentNote || "-"}</td>
        </tr>
      `
    )
    .join("");

  els.confirmModal.classList.toggle("is-mobile-preview", studentViewMode === "card");
  els.confirmModalCard.classList.toggle("is-mobile-preview", studentViewMode === "card");
  els.confirmModal.classList.remove("is-hidden");
  els.confirmModal.setAttribute("aria-hidden", "false");
}

function getActiveEntryField(index, field) {
  const activeRoot = studentViewMode === "card" ? els.studentCardList : els.studentProgramBody;
  return activeRoot?.querySelector(`[data-entry-index="${index}"][data-field="${field}"]`) || null;
}

function buildLogPayload(student, program, entry, index) {
  const read = (field) => getActiveEntryField(index, field);
  const assignedCoach = state.coaches.find((coach) => coach.id === (student?.primaryCoachId || ""));
  const effectiveCoachId = program?.coachId || assignedCoach?.id || "";
  const effectiveCoachName = program?.coachName || assignedCoach?.name || "";
  const submittedAt = timestampNow();
  const submittedDate = getComparableDateKey(submittedAt) || getTodayDateInAppZone();

  return {
    id: `log-${Date.now()}-${index}`,
    programId: program.id,
    programCode: program.code,
    programDate: submittedDate,
    coachId: effectiveCoachId,
    coachName: effectiveCoachName,
    studentId: student.id,
    studentName: student.name,
    category: entry.category,
    exercise: entry.exercise,
    targetSets: entry.targetSets,
    targetType: entry.targetType,
    targetValue: entry.targetValue,
    actualWeight: Number(read("actualWeight")?.value || 0),
    actualSets: Number(read("actualSets")?.value || 0),
    actualReps: Number(read("actualReps")?.value || 0),
    studentNote: read("studentNote")?.value.trim() || "",
    submittedAt
  };
}

function setInputError(input, message) {
  if (!input) {
    return;
  }
  input.classList.add("input-error");
  const wrapper = input.closest(".entry-input, .card-field, .field-group, label") || input.parentElement;
  wrapper?.classList.add("has-error");
  let hint = wrapper.querySelector(".field-error");
  if (!hint) {
    hint = document.createElement("span");
    hint.className = "field-error";
    wrapper.appendChild(hint);
  }
  hint.textContent = message;
}

function clearInputErrors() {
  document.querySelectorAll(".input-error").forEach((input) => input.classList.remove("input-error"));
  document.querySelectorAll(".has-error").forEach((node) => node.classList.remove("has-error"));
  document.querySelectorAll(".field-error").forEach((hint) => hint.remove());
}

function validateSubmissionInputs() {
  clearInputErrors();
  let isValid = true;

  loadedStudentEntries.forEach((entry, index) => {
    const read = (field) => getActiveEntryField(index, field);
    const weightInput = read("actualWeight");
    const setsInput = read("actualSets");
    const repsInput = read("actualReps");

    if (entry.targetType === "reps") {
      if (!weightInput?.value) {
        setInputError(weightInput, "\u8acb\u586b\u5beb\u91cd\u91cf");
        isValid = false;
      }
      if (!setsInput?.value) {
        setInputError(setsInput, "\u8acb\u586b\u5beb\u7d44\u6578");
        isValid = false;
      }
      if (!repsInput?.value) {
        setInputError(repsInput, "\u8acb\u586b\u5beb\u6b21\u6578");
        isValid = false;
      }
    }

    if (entry.targetType === "time" && !setsInput?.value) {
      setInputError(setsInput, "\u8acb\u586b\u5beb\u7d44\u6578");
      isValid = false;
    }

    if (entry.targetType === "rm" && !weightInput?.value) {
      setInputError(weightInput, "\u8acb\u586b\u5beb\u91cd\u91cf");
      isValid = false;
    }
  });

  if (!isValid) {
    const firstError = document.querySelector(".input-error");
    firstError?.scrollIntoView({ behavior: "smooth", block: "center" });
    firstError?.focus?.();
    window.alert("\u9084\u6709\u5fc5\u586b\u6b04\u4f4d\u672a\u5b8c\u6210\uff0c\u8acb\u5148\u88dc\u9f4a\u3002");
  }

  return isValid;
}

function closeConfirmModal() {
  pendingSubmission = null;
  els.confirmModal.classList.add("is-hidden");
  els.confirmModal.classList.remove("is-mobile-preview");
  els.confirmModalCard.classList.remove("is-mobile-preview");
  els.confirmModal.setAttribute("aria-hidden", "true");
  els.confirmCardList.innerHTML = "";
}

function openSuccessModal() {
  const snapshot = getActiveSubmittedSnapshot();
  const logs = snapshot?.logs || [];
  const summaryDate = formatDateDisplay(snapshot?.programDate || "", "");
  const summaryCode = snapshot?.programCode || "\u76ee\u524d\u8ab2\u8868";
  els.successSummary.textContent = `${summaryDate}\uff5c${summaryCode}\uff0c\u672c\u6b21\u8a13\u7df4\u5167\u5bb9\u5df2\u8a18\u9304\u3002`;
  els.successCardList.innerHTML = logs.map((log) => formatLogForSummary(log)).join("");
  els.successBody.innerHTML = logs
    .map(
      (log) => `
        <tr>
          <td>${log.category}</td>
          <td>${log.exercise}</td>
          <td>${formatTarget(log)}</td>
          <td>${formatActual(log)}</td>
          <td>${log.studentNote || "-"}</td>
        </tr>
      `
    )
    .join("");
  els.successModal.classList.toggle("is-mobile-preview", studentViewMode === "card");
  els.successModalCard.classList.toggle("is-mobile-preview", studentViewMode === "card");
  els.successModal.classList.remove("is-hidden");
  els.successModal.setAttribute("aria-hidden", "false");
}

function closeSuccessModal() {
  els.successModal.classList.add("is-hidden");
  els.successModal.classList.remove("is-mobile-preview");
  els.successModalCard.classList.remove("is-mobile-preview");
  els.successModal.setAttribute("aria-hidden", "true");
}

function buildSuccessImageFileName(student, program) {
  return `${student?.name || "學生"}_${program?.code || "課表"}_完成紀錄`;
}

async function downloadSuccessImage() {
  const snapshot = getActiveSubmittedSnapshot();
  const logs = snapshot?.logs || [];
  if (!logs.length) {
    window.alert("目前沒有可下載的完成紀錄。");
    return;
  }

  const student = snapshot
    ? { id: snapshot.studentId, name: snapshot.studentName }
    : getSelectedStudent();
  const program = snapshot
    ? { id: snapshot.programId, code: snapshot.programCode, date: snapshot.programDate }
    : getSelectedStudentProgram();
  const title = buildSuccessImageFileName(student, program);

  try {
    const imageData = await renderSuccessSheetAsJpeg(student, program, logs);
    const link = document.createElement("a");
    link.href = imageData.data;
    link.download = `${title}.png`;
    document.body.appendChild(link);
    link.click();
    link.remove();
  } catch (error) {
    console.error(error);
    window.alert("圖片下載失敗，請稍後再試一次。");
  }
}

function openSuccessHistory() {
  closeSuccessModal();
  studentHistoryOpened = true;
  renderStudentHistoryFilters();
  renderStudentHistory();
  applyStudentViewMode();
  els.studentHistoryCard?.scrollIntoView({ behavior: "smooth", block: "start" });
}

async function renderSuccessSheetAsJpeg(student, program, logs) {
  const width = 1120;
  const rowHeight = 78;
  const height = 250 + logs.length * rowHeight + 40;
  const rowsSvg = logs
    .map((log, index) => {
      const y = 210 + index * rowHeight;
      const fill = index % 2 === 0 ? "#fffdf9" : "#f7f1e8";
      const note = truncateForSheet(log.studentNote || "-", 24);
      return `
        <rect x="40" y="${y}" width="1040" height="${rowHeight}" rx="18" fill="${fill}" />
        <text x="72" y="${y + 30}" font-size="20" font-weight="700" fill="#8a552f">${escapeXml(log.category)}</text>
        <text x="240" y="${y + 30}" font-size="26" font-weight="700" fill="#2f241c">${escapeXml(log.exercise)}</text>
        <text x="240" y="${y + 58}" font-size="18" fill="#8e8171">${escapeXml(formatTarget(log))}</text>
        <text x="700" y="${y + 30}" font-size="22" font-weight="700" fill="#2f241c">${escapeXml(formatActual(log))}</text>
        <text x="700" y="${y + 58}" font-size="18" fill="#8e8171">${escapeXml(note)}</text>
      `;
    })
    .join("");

  const svg = `
    <svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
      <rect width="100%" height="100%" fill="#f7f1e8" />
      <rect x="20" y="20" width="${width - 40}" height="${height - 40}" rx="28" fill="#ffffff" stroke="#e6d8c7" />
      <text x="60" y="78" font-size="18" font-weight="800" fill="#8a552f">CoachFlow 完成紀錄</text>
      <text x="60" y="128" font-size="38" font-weight="800" fill="#2f241c">${escapeXml(student?.name || "學生")}｜${escapeXml(program?.code || "課表")}</text>
      <text x="60" y="168" font-size="22" fill="#8e8171">日期：${escapeXml(formatDateDisplay(program?.date || "-"))}　送出時間：${escapeXml(formatDateTimeDisplay(logs[0]?.submittedAt || "-"))}</text>
      ${rowsSvg}
    </svg>
  `;

  const blob = new Blob([svg], { type: "image/svg+xml;charset=utf-8" });
  const url = URL.createObjectURL(blob);

  try {
    const image = await loadImage(url);
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const context = canvas.getContext("2d");
    context.fillStyle = "#f7f1e8";
    context.fillRect(0, 0, width, height);
    context.drawImage(image, 0, 0, width, height);
    return {
      data: canvas.toDataURL("image/png"),
      width,
      height
    };
  } finally {
    URL.revokeObjectURL(url);
  }
}

function loadImage(url) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = reject;
    image.src = url;
  });
}

function buildPdfBlobFromJpeg(dataUrl, imageWidth, imageHeight) {
  const base64 = String(dataUrl).split(",")[1] || "";
  const imageBytes = Uint8Array.from(atob(base64), (char) => char.charCodeAt(0));
  const pdfWidth = 595.28;
  const pdfHeight = (imageHeight / imageWidth) * pdfWidth;
  const pageHeight = Math.max(pdfHeight + 40, 842);

  const parts = [];
  const offsets = [];
  let position = 0;

  const pushString = (value) => {
    const bytes = new TextEncoder().encode(value);
    parts.push(bytes);
    position += bytes.length;
  };

  const pushBytes = (bytes) => {
    parts.push(bytes);
    position += bytes.length;
  };

  pushString("%PDF-1.4\n");

  offsets[1] = position;
  pushString("1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");

  offsets[2] = position;
  pushString("2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n");

  offsets[3] = position;
  pushString(`3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${pdfWidth.toFixed(2)} ${pageHeight.toFixed(2)}] /Resources << /XObject << /Im0 4 0 R >> >> /Contents 5 0 R >>\nendobj\n`);

  offsets[4] = position;
  pushString(`4 0 obj\n<< /Type /XObject /Subtype /Image /Width ${imageWidth} /Height ${imageHeight} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${imageBytes.length} >>\nstream\n`);
  pushBytes(imageBytes);
  pushString("\nendstream\nendobj\n");

  const imageDrawHeight = (imageHeight / imageWidth) * (pdfWidth - 40);
  const contentStream = `q\n${(pdfWidth - 40).toFixed(2)} 0 0 ${imageDrawHeight.toFixed(2)} 20 ${(pageHeight - imageDrawHeight - 20).toFixed(2)} cm\n/Im0 Do\nQ\n`;
  const contentBytes = new TextEncoder().encode(contentStream);

  offsets[5] = position;
  pushString(`5 0 obj\n<< /Length ${contentBytes.length} >>\nstream\n`);
  pushBytes(contentBytes);
  pushString("endstream\nendobj\n");

  const xrefStart = position;
  pushString("xref\n0 6\n");
  pushString("0000000000 65535 f \n");
  for (let i = 1; i <= 5; i += 1) {
    pushString(`${String(offsets[i]).padStart(10, "0")} 00000 n \n`);
  }
  pushString(`trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n${xrefStart}\n%%EOF`);

  return new Blob(parts, { type: "application/pdf" });
}

function openProgramPreviewModal() {
  renderPreview();
  els.programPreviewModal.classList.remove("is-hidden");
  els.programPreviewModal.setAttribute("aria-hidden", "false");
}

function closeProgramPreviewModal() {
  els.programPreviewModal.classList.add("is-hidden");
  els.programPreviewModal.setAttribute("aria-hidden", "true");
}

function handleStudentDetailAction() {
  if (studentSubmissionCompleted && (lastSubmittedLogs.length || getLatestSubmittedLogsForCurrentStudentProgram().length)) {
    openSuccessModal();
    return;
  }
  openStudentProgramEditModal();
}

function openStudentProgramEditModal() {
  if (!loadedStudentEntries.length) {
    window.alert("\u8acb\u5148\u8f09\u5165\u4e00\u4efd\u8ab2\u8868\uff0c\u518d\u8abf\u6574\u9805\u76ee\u3002");
    return;
  }

  els.studentProgramEditBody.innerHTML = "";
  els.studentEditMobileList.innerHTML = "";
  clearStudentProgramEditNotice();
  const editDraft = getSessionStudentProgramEditDraft(
    currentStudentId || getSelectedStudent()?.id || "",
    currentStudentProgramId || loadedStudentProgramId || ""
  );
  if (editDraft?.rows?.length) {
    editDraft.rows.forEach((entry) => addStudentProgramEditRow(entry, entry.sourceIndex));
  } else {
    loadedStudentEntries.forEach((entry, index) => addStudentProgramEditRow(entry, index));
  }
  els.studentProgramEditModal.classList.toggle("is-mobile-preview", studentViewMode === "card");
  els.studentProgramEditModalCard.classList.toggle("is-mobile-preview", studentViewMode === "card");
  els.studentProgramEditModal.classList.remove("is-hidden");
  els.studentProgramEditModal.setAttribute("aria-hidden", "false");
  captureStudentProgramEditDraftFromDom();
}

function closeStudentProgramEditModal() {
  clearStudentProgramEditNotice();
  els.studentProgramEditModal.classList.add("is-hidden");
  els.studentProgramEditModal.classList.remove("is-mobile-preview");
  els.studentProgramEditModalCard.classList.remove("is-mobile-preview");
  els.studentProgramEditModal.setAttribute("aria-hidden", "true");
}

function showStudentProgramEditNotice(message, tone = "info") {
  const text = String(message || "").trim();
  if (!text) {
    return;
  }

  if (!els.studentProgramEditNotice) {
    window.alert(text);
    return;
  }
  els.studentProgramEditNotice.textContent = text;
  els.studentProgramEditNotice.classList.add("is-visible");
  els.studentProgramEditNotice.classList.toggle("is-warning", tone === "warning");
  if (els.studentProgramStatus) {
    els.studentProgramStatus.textContent = text;
  }
  window.alert(text);
  window.clearTimeout(studentEditNoticeTimer);
  studentEditNoticeTimer = window.setTimeout(() => {
    clearStudentProgramEditNotice();
  }, 3200);
}

function clearStudentProgramEditNotice() {
  window.clearTimeout(studentEditNoticeTimer);
  studentEditNoticeTimer = null;
  if (!els.studentProgramEditNotice) {
    return;
  }
  els.studentProgramEditNotice.textContent = "";
  els.studentProgramEditNotice.classList.remove("is-visible", "is-warning");
}

function openAssignCoachModal(studentId) {
  const student = state.students.find((item) => item.id === studentId);
  if (!student || !els.assignCoachModal || !els.assignCoachSelect) {
    return;
  }
  assigningStudentId = studentId;
  const activeCoaches = state.coaches.filter((coach) => coach.status !== "inactive");
  els.assignCoachSelect.innerHTML = activeCoaches
    .map((coach) => `<option value="${coach.id}">${coach.name}</option>`)
    .join("");
  els.assignCoachSelect.value = student.primaryCoachId || activeCoaches[0]?.id || "";
  if (els.assignCoachSummary) {
    els.assignCoachSummary.textContent = `請選擇 ${student.name} 的主要教練。`;
  }
  els.assignCoachModal.classList.remove("is-hidden");
  els.assignCoachModal.setAttribute("aria-hidden", "false");
}

function closeAssignCoachModal() {
  assigningStudentId = "";
  if (!els.assignCoachModal) {
    return;
  }
  els.assignCoachModal.classList.add("is-hidden");
  els.assignCoachModal.setAttribute("aria-hidden", "true");
}

async function saveAssignedCoach() {
  const student = state.students.find((item) => item.id === assigningStudentId);
  const nextCoachId = els.assignCoachSelect?.value || "";
  const nextCoach = state.coaches.find((coach) => coach.id === nextCoachId && coach.status !== "inactive");
  if (!student || !nextCoach) {
    window.alert("請先選擇有效的教練。");
    return;
  }
  if (IS_CLOUD_MODE) {
    try {
      const payload = await callCloudApi("assignStudentCoach", {
        studentId: student.id,
        coachId: nextCoach.id
      });
      applyCloudPayloadToState(payload);
    } catch (error) {
      console.warn("Cloud assign coach failed, falling back to local state:", error);
      student.primaryCoachId = nextCoach.id;
      student.primaryCoachName = nextCoach.name;
      persistState();
    }
  } else {
    student.primaryCoachId = nextCoach.id;
    persistState();
  }
  if (APP_MODE === "admin") {
    refreshAdminWorkspace();
  } else {
    renderCoachStudentLinks();
    renderCoachStudentRoster();
    renderCoachRoster();
    renderCoachToday();
    renderCoachHistoryFilters();
    renderCoachHistory();
    renderStudentProgramOptions();
    renderStudentHistoryFilters();
    renderStudentHistory();
    renderStudentSummary();
  }
  closeAssignCoachModal();
  window.alert(`已指派給 ${nextCoach.name} 教練。`);
}

function saveStudentProgramEdits() {
  const editRows = studentViewMode === "card"
    ? [...els.studentEditMobileList.querySelectorAll(".student-edit-mobile-card")]
    : [...els.studentProgramEditBody.querySelectorAll("tr")];

  const editedEntries = editRows
    .map((row) => {
      const type = row.querySelector(".student-edit-type").value;
      const sourceIndex = Number(row.dataset.sourceIndex);
      const sourceEntry = Number.isInteger(sourceIndex) && sourceIndex >= 0
        ? loadedStudentEntries[sourceIndex]
        : null;
      return {
        itemId: sourceEntry?.itemId || `custom-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
        category: row.querySelector(".student-edit-category").value,
        exercise: row.querySelector(".student-edit-exercise").value.trim(),
        targetSets: Number(row.querySelector(".student-edit-sets").value || 0),
        targetType: type,
        targetValue: type === "rm" ? 0 : Number(row.querySelector(".student-edit-value").value || 0),
        itemNote: sourceEntry?.itemNote || "",
        actualWeight: sourceEntry?.actualWeight || "",
        actualSets: sourceEntry?.actualSets || "",
        actualReps: sourceEntry?.actualReps || "",
        studentNote: sourceEntry?.studentNote || "",
        referenceLog: null
      };
    })
    .filter((entry) => entry.exercise && entry.targetSets > 0 && (entry.targetType === "rm" || entry.targetValue > 0))
    .map((entry) => ({
      ...entry,
      referenceLog: findLatestReferenceLog(getSelectedStudent()?.id || "", entry)
    }));

  if (!editedEntries.length) {
    window.alert("\u8acb\u81f3\u5c11\u4fdd\u7559\u4e00\u7b46\u8ab2\u8868\u9805\u76ee\u3002");
    return;
  }

  loadedStudentEntries = editedEntries;
  studentEntriesDirty = true;
  clearStudentProgramEditDraft(false);
  persistSession();
  renderStudentProgramEntries();
  closeStudentProgramEditModal();
}

function setStudentViewMode(mode) {
  studentViewMode = mode;
  persistSession();
  applyStudentViewMode();
}

function setCoachViewMode(mode) {
  coachViewMode = mode;
  persistSession();
  applyCoachViewMode();
}

function applyCoachHistoryVisibility() {
  const isMobile = coachViewMode === "mobile";
  const isOpen = !isMobile || coachHistoryOpened;
  els.coachHistoryResultsWrap?.classList.toggle("is-hidden", !isOpen);
  els.toggleCoachHistoryResults?.classList.toggle("is-hidden", !isMobile);
  if (els.coachHistoryMore) {
    els.coachHistoryMore.classList.toggle("is-hidden", !isOpen || els.coachHistoryMore.dataset.empty === "true");
  }
  if (els.toggleCoachHistoryResults) {
    els.toggleCoachHistoryResults.textContent = isOpen ? "收起歷史紀錄" : "查看歷史紀錄";
  }
}

function applyCoachViewMode() {
  const isMobile = coachViewMode === "mobile";
  const coachPanel = document.querySelector("#coach-panel");
  coachPanel?.classList.toggle("coach-mobile-preview", isMobile);
  els.coachGlobalViewDesktopButton?.classList.toggle("is-active", !isMobile);
  els.coachGlobalViewMobileButton?.classList.toggle("is-active", isMobile);
  applyCoachHistoryVisibility();
}

function toggleCoachHistoryResults() {
  coachHistoryOpened = !coachHistoryOpened;
  applyCoachHistoryVisibility();
}

function toggleTodayLogsVisibleCount() {
  const { logs: allTodayLogs } = getCoachRoundContext();
  const keyword = (els.todayLogsSearch?.value || "").trim().toLowerCase();
  const filteredLogs = allTodayLogs.filter((log) => {
    if (!keyword) {
      return true;
    }
    return [log.studentName, log.exercise]
      .some((value) => String(value || "").toLowerCase().includes(keyword));
  });

  coachTodayVisibleCount = coachTodayVisibleCount < filteredLogs.length ? coachTodayVisibleCount + 6 : 6;
  renderCoachToday();
}

function toggleCoachHistoryVisibleCount() {
  const logs = getCoachHistoryLogs();
  coachHistoryVisibleCount = coachHistoryVisibleCount < logs.length ? coachHistoryVisibleCount + 6 : 6;
  renderCoachHistory();
}

function toggleTodayStudentDetailVisibleCount() {
  const { logs: roundLogs } = getCoachRoundContext();
  const studentLogs = roundLogs
    .filter((log) => log.studentId === selectedTodayStudentId)
    .sort(compareLogsByDateTimeAsc);

  todayStudentDetailVisibleCount = todayStudentDetailVisibleCount < studentLogs.length ? todayStudentDetailVisibleCount + 6 : 6;
  renderTodayStudentDetail(roundLogs);
}

function applyStudentViewMode() {
  const isCard = studentViewMode === "card";
  const hasStudent = Boolean(getSelectedStudent());
  const hasLoadedProgram = loadedStudentEntries.length > 0;
  els.studentGlobalViewCardButton.classList.toggle("is-active", isCard);
  els.studentGlobalViewTableButton.classList.toggle("is-active", !isCard);
  els.studentViewCardButton.classList.toggle("is-active", isCard);
  els.studentViewTableButton.classList.toggle("is-active", !isCard);
  els.studentCardList.classList.toggle("is-active", isCard);
  els.studentPhoneFrame.classList.toggle("is-active", isCard);
  els.studentTableWrap.classList.toggle("is-hidden", isCard);
  els.studentMobileSubmitBar.classList.toggle("is-visible", isCard && hasLoadedProgram);
  els.studentMobileTools.classList.toggle("is-visible", isCard && hasStudent);
  els.studentMobileSecondaryActions.classList.toggle("is-visible", isCard && hasLoadedProgram);
  els.editStudentProgramMobileTop?.classList.toggle("is-hidden", !isCard || !hasLoadedProgram);
  els.studentProgramCard.classList.toggle("is-mobile-preview", isCard);
  els.studentHistoryCard.classList.toggle("is-mobile-preview", isCard);
  els.studentHistoryCard.classList.toggle("is-mobile-open", !isCard || studentHistoryOpened);
  els.studentAuthShell.classList.toggle("is-mobile-preview", isCard && !hasStudent);
}

function handleViewportChange() {
  if (window.innerWidth <= 820 && studentViewMode !== "card") {
    setStudentViewMode("card");
  }
}

function syncStudentAccessUI() {
  const hasStudent = Boolean(getSelectedStudent());
  const isCard = studentViewMode === "card";
  const hasLoadedProgram = loadedStudentEntries.length > 0;
  syncStudentLeaveSystemEntry();
  document.body.dataset.studentAuth = hasStudent ? "authenticated" : "anonymous";
  els.studentAuthShell.classList.toggle("is-autologin", studentAutoLoginPending);
  if (els.studentAutoLoginBanner) {
    els.studentAutoLoginBanner.classList.toggle("is-visible", studentAutoLoginPending);
    els.studentAutoLoginBanner.classList.toggle("is-hidden", !studentAutoLoginPending);
  }
  els.studentAuthShell.classList.toggle("is-collapsed", hasStudent);
  els.studentAuthShell.classList.toggle("is-mobile-preview", isCard && !hasStudent);
  if (els.studentProgramSelect) {
    els.studentProgramSelect.hidden = true;
    els.studentProgramSelect.style.display = "none";
  }
  if (els.studentProgramShell) {
    els.studentProgramShell.hidden = !hasStudent;
    els.studentProgramShell.style.display = hasStudent ? "" : "none";
  }
  if (els.studentHistoryCard) {
    const shouldShowHistory = hasStudent && (!isCard || studentHistoryOpened || loadedStudentEntries.length > 0);
    els.studentHistoryCard.hidden = !shouldShowHistory;
    els.studentHistoryCard.style.display = shouldShowHistory ? "" : "none";
  }
  els.studentActiveBar.classList.toggle("is-hidden", !hasStudent || isCard);
  els.studentProgramCard.classList.toggle("is-mobile-preview", isCard);
  els.studentProgramCard.classList.toggle("is-hidden-mobile-entry", isCard && !hasStudent);
  els.studentHistoryCard.classList.toggle("is-hidden-mobile-entry", isCard && !hasStudent);
  els.loadStudentProgramInline?.classList.toggle("is-hidden", hasLoadedProgram);
  els.loadStudentProgramMobile?.classList.toggle("is-hidden", hasLoadedProgram);
  els.editStudentProgramMobileTop?.classList.toggle("is-hidden", !hasLoadedProgram);
  if (els.loadStudentProgramInline) {
    els.loadStudentProgramInline.hidden = hasLoadedProgram;
  }
  if (els.loadStudentProgramMobile) {
    els.loadStudentProgramMobile.hidden = hasLoadedProgram;
  }
  if (els.editStudentProgramMobileTop) {
    els.editStudentProgramMobileTop.hidden = !hasLoadedProgram;
  }
  els.studentMobileTools.classList.toggle("is-visible", hasStudent && isCard);
  els.studentMobileSecondaryActions.classList.toggle("is-visible", hasStudent && isCard && hasLoadedProgram);
  if (els.studentPanelViewToggle) {
    els.studentPanelViewToggle.hidden = false;
    els.studentPanelViewToggle.classList.remove("is-hidden");
    els.studentPanelViewToggle.style.display = "";
  }
  if (!hasStudent) {
    els.studentMobileSubmitBar.classList.remove("is-visible");
    els.studentMobileTools.classList.remove("is-visible");
    els.studentMobileSecondaryActions.classList.remove("is-visible");
    els.editStudentProgramMobileTop?.classList.add("is-hidden");
  }
}

function showStudentAccessPanel() {
  setStudentAutoLoginPending(false);
  studentHistoryOpened = false;
  loadedStudentEntries = [];
  loadedStudentProgramId = "";
  currentStudentProgramId = "";
  studentEntriesDirty = false;
  studentSubmissionCompleted = false;
  lastSubmittedLogs = [];
  lastSubmittedSnapshot = null;
  clearStudentProgramEditDraft(false);
  clearModeAccessInUrl();
  applyStudentViewMode();
  els.studentAuthShell.classList.remove("is-collapsed");
  els.studentAuthShell.scrollIntoView({ behavior: "smooth", block: "start" });
  els.studentAccessCode.focus();
}

async function finalizeSubmission() {
  if (studentSubmitBusy) {
    return;
  }

  studentSubmitBusy = true;
  setButtonBusy(els.confirmSubmit, true, "送出中...");
  setControlDisabled(els.cancelSubmit, true);
  setControlDisabled(els.closeConfirmModal, true);
  if (els.studentProgramStatus) {
    els.studentProgramStatus.textContent = "正在送出訓練紀錄，請稍候...";
  }

  try {
    await finalizeSubmissionCore();
  } finally {
    studentSubmitBusy = false;
    setButtonBusy(els.confirmSubmit, false);
    setControlDisabled(els.cancelSubmit, false);
    setControlDisabled(els.closeConfirmModal, false);
  }
}

async function finalizeSubmissionCore() {
  if (!pendingSubmission?.length) {
    closeConfirmModal();
    return;
  }

  const submittedLogs = [...pendingSubmission];
  lastSubmittedLogs = [...submittedLogs];
  lastSubmittedSnapshot = buildSubmittedSnapshot(submittedLogs);
  studentSubmissionCompleted = false;
  if (IS_CLOUD_MODE) {
    try {
      await submitStudentLogsToCloud(pendingSubmission);
      const cloudLogs = getLatestSubmittedLogsForCurrentStudentProgram();
      if (cloudLogs.length) {
        lastSubmittedLogs = [...cloudLogs];
        lastSubmittedSnapshot = buildSubmittedSnapshot(cloudLogs);
      }
      studentSubmissionCompleted = true;
    } catch (error) {
      console.error("Cloud submit logs failed:", error);
      closeConfirmModal();
      window.alert("送出失敗，雲端尚未儲存。請稍候再試一次。");
      return;
    }
  } else {
    const submissionKeys = new Set(
      pendingSubmission.map((log) => `${log.programId}::${log.studentId}::${log.programDate}::${log.exercise}`)
    );
    state.workoutLogs = state.workoutLogs.filter(
      (log) => !submissionKeys.has(`${log.programId}::${log.studentId}::${getLogActivityDateKey(log)}::${log.exercise}`)
    );
    state.workoutLogs = submittedLogs.concat(state.workoutLogs);
    persistState();
    lastSubmittedSnapshot = buildSubmittedSnapshot(submittedLogs);
    studentSubmissionCompleted = true;
  }
  closeConfirmModal();
  renderCoachToday();
  renderCoachHistoryFilters();
  renderCoachStudentLinks();
  renderCoachStudentRoster();
  renderCoachRoster();
  renderCoachHistory();
  renderStudentHistoryFilters();
  renderStudentHistory();
  loadedStudentEntries = [];
  loadedStudentProgramId = currentStudentProgramId || loadedStudentProgramId;
  studentEntriesDirty = false;
  clearStudentProgramEditDraft(false);
  persistSession();
  renderStudentProgramEntries();
  els.studentProgramStatus.textContent = `\u672c\u6b21\u8a13\u7df4\u5df2\u8a18\u9304\uff5c\u6700\u5f8c\u66f4\u65b0 ${formatDateTimeDisplay(lastSubmittedLogs[0]?.submittedAt || "", "")}`;
  els.studentMobileSubmitBar.classList.remove("is-visible");
  els.studentMobileTools.classList.remove("is-visible");
  els.studentMobileSecondaryActions.classList.add("is-visible");
  openSuccessModal();
}

function renderCoachToday() {
  const { viewDate, program: roundProgram, studentProgramMap, logs: allRoundLogs } = getCoachRoundContext();
  const coachStudents = getCoachActiveStudents();
  const keyword = (els.todayLogsSearch?.value || "").trim().toLowerCase();
  const roundLogs = allRoundLogs.filter((log) => {
    if (!keyword) {
      return true;
    }
    return (
      String(log.studentName || "").toLowerCase().includes(keyword) ||
      String(log.exercise || "").toLowerCase().includes(keyword)
    );
  });
  const visibleRoundLogs = roundLogs.slice(0, coachTodayVisibleCount);
  const submittedStudentIds = new Set(allRoundLogs.map((log) => log.studentId).filter(Boolean));
  const submittedStudents = coachStudents.filter((student) => submittedStudentIds.has(student.id));
  const pendingStudents = coachStudents.filter((student) => !submittedStudentIds.has(student.id));
  const roundPrograms = [...new Set([...studentProgramMap.values()].map((program) => program.id))]
    .map((programId) => state.programs.find((program) => program.id === programId))
    .filter(Boolean);
  const roundLabel = roundPrograms.length > 1
    ? `${roundPrograms.length} 份課表｜含指定學生`
    : roundProgram
    ? `${roundProgram.code || "\u672a\u547d\u540d"}｜${formatDateDisplay(roundProgram.date, "-")}`
    : "\u5c1a\u672a\u5efa\u7acb";

  const statItems = [
    { label: "\u67e5\u770b\u672c\u8f2a\u65e5\u671f", value: formatDateDisplay(viewDate || "-", "-") },
    { label: "\u672c\u8f2a\u8ab2\u8868", value: roundLabel },
    { label: "\u5df2\u9001\u51fa", value: `${submittedStudents.length} \u4eba` },
    { label: "\u672a\u9001\u51fa", value: `${pendingStudents.length} \u4eba` },
    { label: "\u672c\u8f2a\u7d00\u9304", value: `${allRoundLogs.length} \u7b46` }
  ];

  els.coachStats.innerHTML = statItems
    .map(
      (item) => `
        <div class="stat-card">
          <div class="stat-label">${item.label}</div>
          <div class="stat-value">${item.value}</div>
        </div>
      `
    )
    .join("");

  els.submittedList.innerHTML = submittedStudents.length
    ? submittedStudents
        .map((student) => `<li><button class="ghost-button today-student-list-button" type="button" data-open-submitted-student="${student.id || ""}">${student.name}</button></li>`)
        .join("")
    : `<li class="muted-copy">\u672c\u8f2a\u76ee\u524d\u9084\u6c92\u6709\u4eba\u9001\u51fa\u3002</li>`;

  els.pendingList.innerHTML = pendingStudents.length
    ? pendingStudents
        .map((student) => `<li><button class="ghost-button today-student-list-button" type="button" data-open-pending-student="${student.id || ""}">${student.name}</button></li>`)
        .join("")
    : `<li class="muted-copy">${coachStudents.length ? "\u672c\u8f2a\u6240\u6709\u5b78\u751f\u90fd\u5df2\u9001\u51fa\u3002" : "\u76ee\u524d\u6c92\u6709\u53ef\u8ffd\u8e64\u7684\u5b78\u751f\u3002"}</li>`;

  if (els.todayLogsCount) {
    els.todayLogsCount.textContent = `\u5171 ${visibleRoundLogs.length} / ${roundLogs.length} \u7b46`;
  }

  els.todayLogsBody.innerHTML = visibleRoundLogs.length
    ? visibleRoundLogs
        .map(
          (log) => `
            <tr class="${selectedTodayStudentId === log.studentId ? "is-selected-log" : ""}">
              <td><button class="ghost-button ${selectedTodayStudentId === log.studentId ? "is-selected-log-button" : ""}" type="button" data-open-today-student="${log.studentId}">${log.studentName}</button></td>
              <td>${log.category}</td>
              <td>${log.exercise}</td>
              <td>${formatTarget(log)}</td>
              <td>${formatActual(log)}</td>
              <td>${log.studentNote || "-"}</td>
              <td>${formatDateTimeDisplay(log.submittedAt)}</td>
            </tr>
          `
        )
        .join("")
    : `<tr><td colspan="7" class="empty-state">\u76ee\u524d\u9084\u6c92\u6709\u672c\u8f2a\u7d00\u9304\u3002</td></tr>`;

  if (els.todayLogsMore) {
    const remaining = roundLogs.length - visibleRoundLogs.length;
    const shouldHide = roundLogs.length <= 6;
    els.todayLogsMore.classList.toggle("is-hidden", shouldHide);
    if (!shouldHide) {
      els.todayLogsMore.textContent = remaining > 0 ? `查看更多（剩餘 ${remaining} 筆）` : "收合";
    }
  }

  if (!selectedTodayStudentId || !allRoundLogs.some((log) => log.studentId === selectedTodayStudentId)) {
    selectedTodayStudentId = allRoundLogs[0]?.studentId || "";
  }
  renderTodayStudentDetail(allRoundLogs);
}

function handleTodayLogAction(event) {
  const button = event.target.closest("[data-open-today-student]");
  if (!button) {
    return;
  }

  selectedTodayStudentId = button.dataset.openTodayStudent || "";
  todayStudentDetailVisibleCount = 6;
  const { logs: roundLogs } = getCoachRoundContext();
  renderTodayStudentDetail(roundLogs);
  document.querySelector("#today-student-detail-card")?.scrollIntoView({ behavior: "smooth", block: "start" });
}

function handleTodayStudentListAction(event) {
  const submittedButton = event.target.closest("[data-open-submitted-student]");
  if (submittedButton) {
    selectedTodayStudentId = submittedButton.dataset.openSubmittedStudent || "";
    todayStudentDetailVisibleCount = 6;
    const { logs: roundLogs } = getCoachRoundContext();
    renderCoachToday();
    renderTodayStudentDetail(roundLogs);
    document.querySelector("#today-student-detail-card")?.scrollIntoView({ behavior: "smooth", block: "start" });
    return;
  }

  const pendingButton = event.target.closest("[data-open-pending-student]");
  if (!pendingButton) {
    return;
  }

  const student = state.students.find((item) => item.id === pendingButton.dataset.openPendingStudent);
  if (!student) {
    return;
  }

  switchCoachPanel("coach-students");
  if (els.coachStudentLinkName) {
    els.coachStudentLinkName.value = student.name;
  }
  if (els.coachStudentLinkSelect) {
    els.coachStudentLinkSelect.value = student.name;
  }
  renderCoachStudentLinks();
  els.coachStudentLinkName?.scrollIntoView({ behavior: "smooth", block: "center" });
}

function renderTodayStudentDetail(todayLogs = []) {
  if (!els.todayStudentDetailBody || !els.todayStudentDetailStatus) {
    return;
  }

  if (!todayLogs.length) {
    els.todayStudentDetailStatus.textContent = "\u672c\u8f2a\u5c1a\u7121\u7d00\u9304";
    els.todayStudentDetailBody.innerHTML = `<tr><td colspan="6" class="empty-state">\u8acb\u5148\u7b49\u5b78\u751f\u9001\u51fa\u672c\u8f2a\u7d00\u9304\u3002</td></tr>`;
    els.todayStudentDetailMore?.classList.add("is-hidden");
    return;
  }

  const studentLogs = todayLogs
    .filter((log) => log.studentId === selectedTodayStudentId)
    .sort(compareLogsByDateTimeAsc);
  const studentName = studentLogs[0]?.studentName || "";

  if (!studentLogs.length) {
    els.todayStudentDetailStatus.textContent = "\u8acb\u5148\u9078\u64c7\u5b78\u751f";
    els.todayStudentDetailBody.innerHTML = `<tr><td colspan="6" class="empty-state">\u8acb\u5f9e\u4e0a\u65b9\u672c\u8f2a\u7d00\u9304\u8868\u683c\u9ede\u9078\u4e00\u4f4d\u5b78\u751f\u3002</td></tr>`;
    els.todayStudentDetailMore?.classList.add("is-hidden");
    return;
  }

  const visibleLogs = studentLogs.slice(0, todayStudentDetailVisibleCount);
  els.todayStudentDetailStatus.textContent = `${studentName}\u672c\u8f2a\u5171 ${visibleLogs.length} / ${studentLogs.length} \u7b46`;
  els.todayStudentDetailBody.innerHTML = visibleLogs
    .map(
      (log) => `
        <tr>
          <td>${log.category}</td>
          <td>${log.exercise}</td>
          <td>${formatTarget(log)}</td>
          <td>${formatActual(log)}</td>
          <td>${log.studentNote || "-"}</td>
          <td>${formatDateTimeDisplay(log.submittedAt)}</td>
        </tr>
      `
    )
    .join("");

  if (els.todayStudentDetailMore) {
    const remaining = studentLogs.length - visibleLogs.length;
    const shouldHide = studentLogs.length <= 6;
    els.todayStudentDetailMore.classList.toggle("is-hidden", shouldHide);
    if (!shouldHide) {
      els.todayStudentDetailMore.textContent = remaining > 0 ? `查看更多（剩餘 ${remaining} 筆）` : "收合";
    }
  }
}

async function hydrateCoachAccessFromUrl() {
  if (APP_MODE !== "coach") {
    return;
  }
  const params = new URLSearchParams(window.location.search);
  const accessValue = normalizeAccessInput(params.get("coach") || params.get("token") || params.get("code") || "");
  if (!accessValue) {
    return;
  }
  let matchedCoach = null;
  if (IS_CLOUD_MODE) {
    try {
      matchedCoach = await resolveCoachAccessFromCloud(accessValue);
    } catch (error) {
      console.warn("Coach cloud hydrate failed, falling back to local state:", error);
      matchedCoach = resolveCoachByAccess(accessValue);
    }
  } else {
    matchedCoach = resolveCoachByAccess(accessValue);
  }
  if (!matchedCoach) {
    return;
  }
  authenticatedCoachId = matchedCoach.id;
  authenticatedCoachAccess = matchedCoach.accessCode || String(accessValue || "").trim();
  state.currentCoachId = matchedCoach.id;
  if (els.coachAccessCode) {
    els.coachAccessCode.value = matchedCoach.accessCode || accessValue;
  }
  persistSession();
}

function renderCoachHistoryFilters() {
  const studentOptions = [
    `<option value="">\u5168\u90e8\u5b78\u751f</option>`,
    ...getCoachScopedStudents().map((student) => `<option value="${student.id}">${student.name}</option>`)
  ];

  const exercises = [...new Set(getCoachScopedLogs().map((log) => log.exercise))].sort();
  const exerciseOptions = [
    `<option value="">\u5168\u90e8\u52d5\u4f5c</option>`,
    ...exercises.map((exercise) => `<option value="${exercise}">${exercise}</option>`)
  ];

  els.coachHistoryStudent.innerHTML = studentOptions.join("");
  els.coachHistoryExercise.innerHTML = exerciseOptions.join("");
}

function renderCoachStudentLinks() {
  if (!els.coachStudentLinks) {
    return;
  }

  const configuredPublicBase = String(APP_CONFIG.publicBaseUrl || "").trim();
  const isHosted = /^https?:/i.test(window.location.protocol);
  const baseHref = configuredPublicBase || (isHosted ? window.location.href : "");
  const studentBaseUrl = baseHref ? new URL("student.html", baseHref) : null;
  const canGenerateScannableQr = Boolean(studentBaseUrl);

  if (!studentBaseUrl) {
    els.coachStudentLinks.innerHTML = `
      <div class="empty-card">
        目前是本機檔案模式，QR code 無法提供可掃描的公開網址。
        <br>
        請先將系統部署到網站，或在 <code>config.js</code> 設定 <code>publicBaseUrl</code>。
      </div>
    `;
    return;
  }

  studentBaseUrl.search = "";
  studentBaseUrl.hash = "";
  studentBaseUrl.searchParams.set("mode", "student");
  studentBaseUrl.searchParams.set("v", PUBLIC_APP_VERSION);
  const sharedBaseUrl = new URL(studentBaseUrl.toString());
  sharedBaseUrl.search = "";
  sharedBaseUrl.searchParams.set("mode", "student");
  sharedBaseUrl.searchParams.set("v", PUBLIC_APP_VERSION);
  const activeCoach = getCurrentCoach();
  const coachAccessCode = String(activeCoach?.accessCode || "").trim();
  const leaveCoachQuickUrl = buildCoachLeaveSystemUrl({
    baseHref,
    from: "coachflow",
    forceCoachCode: coachAccessCode
  });
  const leaveStudentBaseUrl =
    IS_LEAVE_SANDBOX_ENABLED && LEAVE_SANDBOX_STUDENT_PAGE
      ? new URL(LEAVE_SANDBOX_STUDENT_PAGE, baseHref)
      : null;

  const keyword = (els.coachStudentLinkName?.value || "").trim().toLowerCase();
  if (els.coachStudentLinkSelect && els.coachStudentLinkName?.value) {
    els.coachStudentLinkSelect.value = els.coachStudentLinkName.value;
  }

  const scopedStudents = getCoachScopedStudents();
  const selectedNameKeyword = String(els.coachStudentLinkSelect?.value || "").trim().toLowerCase();
  const activeKeyword = keyword || selectedNameKeyword;
  const matchedStudents = scopedStudents.filter((student) => {
    if (!activeKeyword) {
      return false;
    }

      return (
        String(student.name || "").trim().toLowerCase().includes(activeKeyword) ||
        String(student.accessCode || "").trim().toLowerCase().includes(activeKeyword)
      );
    });
  const displayStudents = matchedStudents.length
    ? matchedStudents
    : (!activeKeyword && scopedStudents.length ? [scopedStudents[0]] : []);

  if (els.coachStudentLinkOptions) {
    els.coachStudentLinkOptions.innerHTML = scopedStudents
      .map((student) => `<option value="${student.name}"></option>`)
      .join("");
  }

  if (els.coachStudentLinkSelect) {
    const currentValue = els.coachStudentLinkSelect.value;
    els.coachStudentLinkSelect.innerHTML = [
      `<option value="">\u8acb\u9078\u64c7\u5b78\u751f</option>`,
      ...scopedStudents.map((student) => `<option value="${student.name}">${student.name}</option>`)
    ].join("");
    els.coachStudentLinkSelect.value = scopedStudents.some((student) => student.name === currentValue)
      ? currentValue
      : (displayStudents[0]?.name || "");
  }

  if (!displayStudents.length) {
    els.coachStudentLinks.innerHTML = `
      <div class="empty-card">\u627e\u4e0d\u5230\u7b26\u5408\u7684\u5b78\u751f\uff0c\u8acb\u91cd\u65b0\u78ba\u8a8d\u59d3\u540d\u6216\u4ee3\u78bc\u3002</div>
      ${leaveCoachQuickUrl
        ? `<div class="coach-link-actions"><a class="ghost-button" href="${leaveCoachQuickUrl}" target="_blank" rel="noopener" data-open-coach-leave="1">開啟請假教練頁（測試）</a></div>`
        : ""}
    `;
    return;
  }

  els.coachStudentLinks.innerHTML = displayStudents
    .slice(0, 1)
    .map((student) => {
      const accessUrl = new URL(studentBaseUrl.toString());
      const directAccessValue = student.accessCode || student.id;
      accessUrl.searchParams.set("code", directAccessValue);
      const sharedUrl = sharedBaseUrl.toString();
      const qrImageUrl = canGenerateScannableQr
        ? `https://api.qrserver.com/v1/create-qr-code/?size=180x180&data=${encodeURIComponent(accessUrl.toString())}`
        : "";
      const leaveStudentUrl = leaveStudentBaseUrl ? new URL(leaveStudentBaseUrl.toString()) : null;
      if (leaveStudentUrl) {
        leaveStudentUrl.searchParams.set("studentCode", directAccessValue);
        if (coachAccessCode) {
          leaveStudentUrl.searchParams.set("coachCode", coachAccessCode);
        }
        leaveStudentUrl.searchParams.set("from", "coachflow");
        leaveStudentUrl.searchParams.set("autoLogin", "1");
        leaveStudentUrl.searchParams.set("v", PUBLIC_APP_VERSION);
      }
      const leaveCoachUrl = buildCoachLeaveSystemUrl({
        baseHref,
        from: "coachflow",
        forceCoachCode: coachAccessCode
      });
      return `
        <article class="coach-link-card">
            <div class="coach-link-top">
              <div>
                <h4>${student.name}</h4>
              </div>
            <div class="roster-pill-stack">
              <span class="status-pill">\u4ee3\u78bc\uff1a${student.accessCode || "-"}</span>
              <span class="status-pill ${student.status === "inactive" ? "is-muted" : "is-success"}">${student.status === "inactive" ? "\u5df2\u505c\u7528" : "\u4f7f\u7528\u4e2d"}</span>
            </div>
          </div>
          <label class="coach-link-field">
            <span>\u5c08\u5c6c\u7db2\u5740</span>
            <input type="text" readonly value="${accessUrl.toString()}" data-link-value="${student.id}">
          </label>
          <label class="coach-link-field">
            <span>\u5171\u7528\u5165\u53e3\u7db2\u5740</span>
            <input type="text" readonly value="${sharedUrl}" data-shared-link-value="${student.id}">
          </label>
          ${leaveStudentUrl
            ? `
              <label class="coach-link-field">
                <span>請假測試版（學生頁）</span>
                <input type="text" readonly value="${leaveStudentUrl.toString()}" data-leave-student-link-value="${student.id}">
              </label>
            `
            : ""}
          <div class="coach-link-inline-meta">
            <span class="status-pill">\u8acb\u5b78\u751f\u8f38\u5165\u4ee3\u78bc\uff1a${student.accessCode || "-"}</span>
            ${leaveStudentUrl ? `<span class="status-pill">請假測試整合已啟用</span>` : ""}
          </div>
          <div class="coach-link-qr-block">
            <div class="coach-link-qr-card">
              <span>學生專屬 QR</span>
              ${canGenerateScannableQr
                ? `<img src="${qrImageUrl}" alt="${student.name} 的專屬 QR">`
                : `<p class="muted-copy">需設定公開網址後才能產生 QR。</p>`}
            </div>
          </div>
          <div class="coach-link-actions">
            <button class="ghost-button" type="button" data-copy-student-link="${student.id}">\u8907\u88fd\u7db2\u5740</button>
            <button class="ghost-button" type="button" data-copy-shared-link="${student.id}">\u8907\u88fd\u5171\u7528\u5165\u53e3</button>
            ${leaveStudentUrl
              ? `<button class="ghost-button" type="button" data-copy-leave-student-link="${student.id}">複製請假學生頁</button>`
              : ""}
            ${leaveStudentUrl
              ? `<a class="ghost-button" href="${leaveStudentUrl.toString()}" target="_blank" rel="noopener">開啟請假學生頁</a>`
              : ""}
            ${leaveCoachUrl
              ? `<a class="ghost-button" href="${leaveCoachUrl}" target="_blank" rel="noopener" data-open-coach-leave="1">開啟請假教練頁</a>`
              : ""}
            <button class="ghost-button" type="button" data-copy-student-qr="${student.id}" data-qr-link="${accessUrl.toString()}">複製 QR 連結內容</button>
            <button class="primary-button" type="button" data-copy-student-message="${student.id}" data-student-name="${student.name}" data-student-code="${student.accessCode || ""}" data-student-link="${accessUrl.toString()}" data-shared-url="${sharedUrl}">\u8907\u88fd\u7d66\u5b78\u751f\u7684\u8a0a\u606f</button>
          </div>
        </article>
      `;
    })
    .join("");
}

function downloadHistoryImportTemplate() {
  const rows = [
    ["學生姓名", "日期", "動作", "分類", "類型", "目標組數", "目標次數", "實際重量", "備註"],
    ["王小明", "2026-03-01", "深蹲", "主項目", "reps", "5", "5", "100", "狀況穩定"],
    ["王小明", "2026-03-08", "深蹲", "主項目", "rm", "", "3", "110", "3RM 測試"],
    ["陳泳蓁", "2026-03-08", "後腳抬高蹲", "輔助項", "time", "4", "60", "", ""]
  ];
  const csv = rows.map((row) => row.map(escapeCsvValue).join(",")).join("\r\n");
  downloadBlob(new Blob([`\uFEFF${csv}`], { type: "text/csv;charset=utf-8;" }), "coachflow-history-import-template.csv");
}

function escapeCsvValue(value) {
  const text = String(value ?? "");
  if (/[",\r\n]/.test(text)) {
    return `"${text.replace(/"/g, "\"\"")}"`;
  }
  return text;
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 0);
}

function parseCsv(text) {
  const rows = [];
  let current = "";
  let row = [];
  let inQuotes = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];

    if (char === "\"") {
      if (inQuotes && next === "\"") {
        current += "\"";
        index += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (char === "," && !inQuotes) {
      row.push(current);
      current = "";
      continue;
    }

    if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && next === "\n") {
        index += 1;
      }
      row.push(current);
      if (row.some((cell) => String(cell).trim() !== "")) {
        rows.push(row);
      }
      row = [];
      current = "";
      continue;
    }

    current += char;
  }

  row.push(current);
  if (row.some((cell) => String(cell).trim() !== "")) {
    rows.push(row);
  }

  return rows;
}

function normalizeImportHeader(header) {
  const normalized = String(header || "").trim().toLowerCase();
  const aliases = {
    "學生姓名": "student_name",
    "student_name": "student_name",
    "姓名": "student_name",
    "日期": "date",
    "date": "date",
    "動作": "exercise",
    "exercise": "exercise",
    "分類": "category",
    "category": "category",
    "類型": "target_type",
    "target_type": "target_type",
    "目標組數": "target_sets",
    "組數": "target_sets",
    "target_sets": "target_sets",
    "目標次數": "target_value",
    "次數": "target_value",
    "目標值": "target_value",
    "target_value": "target_value",
    "實際重量": "actual_weight",
    "actual_weight": "actual_weight",
    "實際組數": "actual_sets",
    "actual_sets": "actual_sets",
    "實際次數": "actual_reps",
    "actual_reps": "actual_reps",
    "備註": "student_note",
    "student_note": "student_note",
    "課表代碼": "program_code",
    "program_code": "program_code"
  };
  return aliases[header] || aliases[normalized] || normalized;
}

function handleHistoryImportFile(event) {
  const file = event.target?.files?.[0];
  if (!file) {
    return;
  }

  const reader = new FileReader();
  reader.onload = () => {
    try {
      pendingHistoryImportRows = buildHistoryImportPreview(String(reader.result || ""));
      renderHistoryImportPreview();
    } catch (error) {
      pendingHistoryImportRows = [];
      renderHistoryImportPreview();
      window.alert("CSV 解析失敗，請確認檔案格式。");
    }
  };
  reader.readAsText(file, "utf-8");
}

function buildHistoryImportPreview(text) {
  const rows = parseCsv(text.replace(/^\uFEFF/, ""));
  if (!rows.length) {
    return [];
  }

  const headers = rows[0].map((value) => normalizeImportHeader(String(value || "").trim()));
  const dataRows = rows.slice(1);

  return dataRows.map((row, index) => {
    const raw = Object.fromEntries(headers.map((header, cellIndex) => [header, String(row[cellIndex] || "").trim()]));
    const student = state.students.find((item) => item.name === raw.student_name);
    // Allow CSV imports to use RM / rm / Reps / TIME without failing on case alone.
    const targetType = String(raw.target_type || "").trim().toLowerCase();
    const targetSets = Number(raw.target_sets || 0);
    const targetValue = Number(raw.target_value || 0);
    const actualWeight = raw.actual_weight ? Number(raw.actual_weight) : "";
    const actualSets = raw.actual_sets ? Number(raw.actual_sets) : "";
    const actualReps = raw.actual_reps ? Number(raw.actual_reps) : "";
    const normalizedDate = normalizeDateKey(raw.date || "");
    const hasValidDate = Boolean(parseDateTimeInAppZone(raw.date || ""));

    const errors = [];
    if (!student) errors.push("找不到學生");
    if (!raw.date) errors.push("缺少日期");
    if (raw.date && !hasValidDate) errors.push("日期格式錯誤");
    if (!raw.exercise) errors.push("缺少動作");
    if (!["reps", "time", "rm"].includes(targetType)) errors.push("類型錯誤");

    const valid = errors.length === 0;
    return {
      rowNumber: index + 2,
      valid,
      statusText: valid ? "可匯入" : errors.join("、"),
      studentId: student?.id || "",
      studentName: raw.student_name,
      programDate: normalizedDate,
      programCode: raw.program_code || "",
      exercise: raw.exercise,
      category: raw.category || "主項目",
      targetType,
      targetSets,
      targetValue,
      actualWeight,
      actualSets,
      actualReps,
      studentNote: raw.student_note || ""
    };
  });
}

function renderHistoryImportPreview() {
  if (!els.historyImportPreviewBody) {
    return;
  }

  const validCount = pendingHistoryImportRows.filter((row) => row.valid).length;
  const previewRows = [...pendingHistoryImportRows].sort((a, b) => {
    if (a.valid === b.valid) {
      return a.rowNumber - b.rowNumber;
    }
    return a.valid ? 1 : -1;
  });
  if (els.historyImportCount) {
    els.historyImportCount.textContent = pendingHistoryImportRows.length
      ? `可匯入 ${validCount} / ${pendingHistoryImportRows.length} 筆`
      : "尚未匯入";
  }

  if (els.confirmHistoryImportButton) {
    els.confirmHistoryImportButton.disabled = validCount === 0;
  }

  els.historyImportPreviewBody.innerHTML = previewRows.length
    ? previewRows
        .map((row) => `
          <tr>
            <td>${row.studentName || "-"}</td>
            <td>${formatDateDisplay(row.programDate || "-")}</td>
            <td>${row.exercise || "-"}</td>
            <td>${formatTarget(row)}</td>
            <td>${formatActual(row)}</td>
            <td>${row.valid ? "可匯入" : `第 ${row.rowNumber} 列：${row.statusText}`}</td>
          </tr>
        `)
        .join("")
    : `<tr><td colspan="6" class="empty-state">請先下載範本並上傳 CSV。</td></tr>`;
}

async function confirmHistoryImport() {
  const validRows = pendingHistoryImportRows.filter((row) => row.valid);
  if (!validRows.length) {
    window.alert("目前沒有可匯入的資料。");
    return;
  }
  let importedCount = validRows.length;

  if (IS_CLOUD_MODE) {
    try {
      await importHistoryRowsToCloud(validRows);
    } catch (error) {
      console.warn("Cloud import history failed, falling back to local state:", error);
      const importKeys = new Set(validRows.map((row) => `${row.studentId}::${row.programDate}::${row.exercise}`));
      state.workoutLogs = state.workoutLogs.filter(
        (log) => !importKeys.has(`${log.studentId}::${log.programDate}::${log.exercise}`)
      );

      const importedLogs = validRows.map((row, index) => ({
        id: `import-log-${Date.now()}-${index}`,
        programId: "",
        programCode: row.programCode || "",
        programDate: row.programDate,
        studentId: row.studentId,
        studentName: row.studentName,
        category: row.category,
        exercise: row.exercise,
        targetSets: row.targetSets || "",
        targetType: row.targetType,
        targetValue: row.targetValue || "",
        actualWeight: row.actualWeight || "",
        actualSets: row.actualSets || "",
        actualReps: row.actualReps || "",
        studentNote: row.studentNote || "",
        submittedAt: normalizeDateTimeValue(`${row.programDate} 00:00`)
      }));

      state.workoutLogs = importedLogs.concat(state.workoutLogs);
      importedCount = importedLogs.length;
      persistState();
    }
  } else {
    const importKeys = new Set(validRows.map((row) => `${row.studentId}::${row.programDate}::${row.exercise}`));
    state.workoutLogs = state.workoutLogs.filter(
      (log) => !importKeys.has(`${log.studentId}::${log.programDate}::${log.exercise}`)
    );

    const importedLogs = validRows.map((row, index) => ({
      id: `import-log-${Date.now()}-${index}`,
      programId: "",
      programCode: row.programCode || "",
      programDate: row.programDate,
      studentId: row.studentId,
      studentName: row.studentName,
      category: row.category,
      exercise: row.exercise,
      targetSets: row.targetSets || "",
      targetType: row.targetType,
      targetValue: row.targetValue || "",
      actualWeight: row.actualWeight || "",
      actualSets: row.actualSets || "",
      actualReps: row.actualReps || "",
      studentNote: row.studentNote || "",
      submittedAt: normalizeDateTimeValue(`${row.programDate} 00:00`)
    }));

    state.workoutLogs = importedLogs.concat(state.workoutLogs);
    importedCount = importedLogs.length;
    persistState();
  }
  pendingHistoryImportRows = [];
  if (els.historyImportFile) {
    els.historyImportFile.value = "";
  }
  renderHistoryImportPreview();
  renderCoachHistoryFilters();
  renderCoachHistory();
  renderCoachToday();
  renderCoachStudentRoster();
  renderCoachRoster();
  renderStudentHistoryFilters();
  renderStudentHistory();
  renderStudentSummary();
  window.alert(`已匯入 ${importedCount} 筆歷史紀錄。`);
}

function renderCoachStudentRoster() {
  if (!els.coachStudentRoster) {
    return;
  }

  const studentsPool = getCoachScopedStudents();
  const coachLogs = getCoachScopedLogs();
  const roundContext = getCoachRoundContext(getCoachTodayDate());
  const program = roundContext.program || getPublishedProgram(getCurrentCoach()?.id);
  const rosterRoundLogs = roundContext.studentProgramMap?.size
    ? roundContext.logs
    : (program ? coachLogs.filter((log) => isLogInProgramRound(log, program)) : []);
  const submittedIds = new Set(
    rosterRoundLogs.map((log) => log.studentId)
  );
  const displayDate = program?.date
    ? formatRocDateDisplay(program.date, "\u4eca\u65e5")
    : "\u4eca\u65e5";
  const keyword = (els.coachStudentRosterSearch?.value || "").trim().toLowerCase();
  const students = studentsPool.filter((student) => {
    if (!keyword) {
      return true;
    }

    return (
      String(student.name || "").trim().toLowerCase().includes(keyword) ||
      String(student.accessCode || "").trim().toLowerCase().includes(keyword)
    );
  });

  if (els.coachStudentCount) {
    els.coachStudentCount.textContent = `\u5171 ${studentsPool.length} \u4eba`;
  }
  if (els.coachStudentRosterSummary) {
    els.coachStudentRosterSummary.textContent = `\u76ee\u524d\u986f\u793a ${students.length} / ${studentsPool.length} \u4eba`;
  }

  els.coachStudentRoster.innerHTML = students.length
    ? students
        .map((student) => {
          const studentLogs = coachLogs
            .filter((log) => log.studentId === student.id);
          const lastLog = studentLogs
            .sort(compareLogsByDateTimeDesc)[0];
          const lastLogDate = lastLog?.programDate
            ? formatRocDateDisplay(lastLog.programDate, "\u5c1a\u7121\u7d00\u9304")
            : "\u5c1a\u7121\u7d00\u9304";
          const assignedCoach = state.coaches.find((coach) => coach.id === student.primaryCoachId);
          const lastUsedText = formatUsageTimestamp(student.lastUsedAt);

          return `
            <article class="coach-link-card">
              <div class="coach-link-top">
                <div>
                  <h4>${student.name}</h4>
                  <p class="coach-student-meta">\u6700\u5f8c\u586b\u5beb\uff1a${lastLogDate}</p>
                  <p class="coach-student-meta">最後使用：${lastUsedText}</p>
                  <p class="coach-student-meta">\u6b77\u53f2\u7b46\u6578\uff1a${studentLogs.length}</p>
                  ${APP_MODE === "admin" ? `<p class="coach-student-meta">主要教練：${assignedCoach?.name || "暫無教練"}</p>` : ""}
                </div>
                <div class="roster-pill-stack">
                  <span class="status-pill">\u4ee3\u78bc\uff1a${student.accessCode || "-"}</span>
                  <span class="status-pill ${student.status === "inactive" ? "is-muted" : "is-success"}">${student.status === "inactive" ? "\u5df2\u505c\u7528" : "\u4f7f\u7528\u4e2d"}</span>
                  <span class="status-pill ${submittedIds.has(student.id) ? "is-success" : "is-muted"}">${displayDate} ${submittedIds.has(student.id) ? "\u5df2\u9001\u51fa" : "\u672a\u9001\u51fa"}</span>
                </div>
              </div>
              <div class="coach-link-actions">
                ${APP_MODE !== "admin" ? `<button class="ghost-button" type="button" data-open-student-link="${student.id}" data-student-name="${student.name}">\u67e5\u770b\u7db2\u5740</button>` : ""}
                ${APP_MODE !== "admin" ? `<button class="ghost-button" type="button" data-open-student-history="${student.id}">\u67e5\u770b\u5168\u90e8\u7d00\u9304</button>` : ""}
                ${APP_MODE === "admin" ? `<button class="ghost-button" type="button" data-assign-student-coach="${student.id}">指派教練</button>` : ""}
                <button class="ghost-button" type="button" data-edit-student="${student.id}">\u4fee\u6539\u59d3\u540d</button>
                <button class="ghost-button" type="button" data-copy-student-code="${student.id}" data-student-code="${student.accessCode || ""}">\u8907\u88fd\u4ee3\u78bc</button>
                <button class="ghost-button" type="button" data-toggle-student-status="${student.id}">${student.status === "inactive" ? "\u555f\u7528\u5b78\u751f" : "\u505c\u7528\u5b78\u751f"}</button>
                <button class="ghost-button" type="button" data-delete-student="${student.id}">\u522a\u9664\u5b78\u751f</button>
              </div>
            </article>
          `;
        })
        .join("")
    : `<div class="empty-card">\u76ee\u524d\u6c92\u6709\u7b26\u5408\u689d\u4ef6\u7684\u5b78\u751f\u3002</div>`;
}

async function createCoachFromForm() {
  const name = (els.newCoachName?.value || "").trim();
  if (!name) {
    window.alert("請先輸入教練姓名。");
    els.newCoachName?.focus();
    return;
  }

  const exists = (state.coaches || []).some((coach) => String(coach.name || "").trim() === name);
  if (exists) {
    window.alert("這位教練已經存在。");
    els.newCoachName?.focus();
    return;
  }

  const nextIndex = state.coaches.length + 1;
  const accessCode = buildCoachAccessCode(name, nextIndex);
  const tokenBase = String(name || "")
    .replace(/\s+/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");

  const coach = {
    id: `coach-${Date.now()}`,
    name,
    accessCode,
    token: tokenBase || `coach${String(nextIndex).padStart(3, "0")}`,
    status: "active",
    role: "coach",
    lastUsedAt: ""
  };

  if (IS_CLOUD_MODE) {
    try {
      const payload = await callCloudApi("upsertCoach", {
        coach: {
          coach_id: coach.id,
          coach_name: coach.name,
          access_code: coach.accessCode,
          token: coach.token,
          status: coach.status,
          role: coach.role,
          last_used_at: coach.lastUsedAt
        }
      });
      applyCloudPayloadToState(payload);
      if (APP_MODE === "admin") {
        const freshAdminPayload = await callCloudApi("bootstrapAdmin", {}, "GET");
        applyCloudPayloadToState(freshAdminPayload);
      }
    } catch (error) {
      console.warn("Cloud createCoach failed, falling back to local state:", error);
      state.coaches.push(coach);
      if (APP_MODE !== "admin") {
        state.currentCoachId = coach.id;
      }
      persistState();
    }
  } else {
    state.coaches.push(coach);
    if (APP_MODE !== "admin") {
      state.currentCoachId = coach.id;
    }
    persistState();
  }
  if (APP_MODE === "admin") {
    refreshAdminWorkspace();
    els.newCoachName.value = "";
    switchCoachPanel("coach-coaches");
    window.alert(`教練已新增，代碼為 ${coach.accessCode}。`);
    return;
  }
  refreshCoachWorkspace();
  els.newCoachName.value = "";
  switchCoachPanel("coach-coaches");
  window.alert(`教練已新增，代碼為 ${coach.accessCode}。`);
}

async function createStudentFromCoachForm() {
  const name = (els.newStudentName?.value || "").trim();

  if (!name) {
    window.alert("\u8acb\u5148\u8f38\u5165\u5b78\u751f\u59d3\u540d\u3002");
    els.newStudentName?.focus();
    return;
  }

  const exists = state.students.some((student) => String(student.name || "").trim() === name);
  if (exists) {
    window.alert("\u9019\u4f4d\u5b78\u751f\u5df2\u7d93\u5b58\u5728\u3002");
    els.newStudentName?.focus();
    return;
  }

  const tokenBase = getNextAvailableStudentToken();
  const accessCode = getNextAvailableStudentAccessCode(name);
  const nextStudentId = `stu-${Date.now()}`;
  const defaultCoachId = APP_MODE === "admin" ? "" : (getCurrentCoach()?.id || "");

  const studentRecord = {
    id: nextStudentId,
    name,
    primaryCoachId: defaultCoachId,
    status: "active",
    accessCode,
    token: tokenBase
  };

  if (IS_CLOUD_MODE) {
    try {
      const payload = await callCloudApi("upsertStudent", {
        student: {
          student_id: studentRecord.id,
          student_name: studentRecord.name,
          primary_coach_id: studentRecord.primaryCoachId,
          status: studentRecord.status,
          access_code: studentRecord.accessCode,
          token: studentRecord.token
        }
      });
      applyCloudPayloadToState(payload);
    } catch (error) {
      console.warn("Cloud createStudent failed, falling back to local state:", error);
      state.students.push(studentRecord);
      persistState();
    }
  } else {
    state.students.push(studentRecord);
    persistState();
  }

  switchCoachPanel("coach-students");
  if (APP_MODE !== "admin" && els.coachStudentLinkName) {
    els.coachStudentLinkName.value = name;
  }
  if (APP_MODE !== "admin" && els.coachStudentLinkSelect) {
    els.coachStudentLinkSelect.value = name;
  }
  if (APP_MODE !== "admin") {
    renderCoachStudentLinks();
  }
  if (APP_MODE === "admin") {
    refreshAdminWorkspace();
  } else {
    renderCoachStudentRoster();
    renderCoachRoster();
    renderCoachToday();
    renderCoachHistoryFilters();
    renderStudentProgramOptions();
    renderStudentHistoryFilters();
  }

  els.newStudentName.value = "";
  window.alert("\u5b78\u751f\u5df2\u65b0\u589e\u3002");
  const nextStudentCard = document.querySelector(`[data-student-card="${nextStudentId}"]`);
  nextStudentCard?.scrollIntoView({ behavior: "smooth", block: "center" });
  const activeTabButton = document.querySelector('.sub-tab[data-coach-tab="coach-students"]');
  activeTabButton?.scrollIntoView({ behavior: "smooth", block: "nearest", inline: "center" });
}

function buildAccessCode(name, index) {
  const prefix = getAccessCodePrefix(name, "ST");
  return `${prefix}${String(index).padStart(3, "0")}`;
}

function getAccessCodePrefix(name, fallback = "ST") {
  const latin = String(name || "")
    .replace(/[^A-Za-z]/g, "")
    .toUpperCase()
    .slice(0, 2);

  return latin || fallback;
}

function getNextAvailableStudentAccessCode(name) {
  const prefix = getAccessCodePrefix(name, "ST");
  const usedCodes = new Set(
    state.students
      .map((student) => String(student.accessCode || "").trim().toUpperCase())
      .filter(Boolean)
  );
  const maxExistingNumber = state.students.reduce((max, student) => {
    const match = String(student.accessCode || "").trim().toUpperCase().match(/^([A-Z]+)(\d+)$/);
    if (!match || match[1] !== prefix) {
      return max;
    }
    return Math.max(max, Number(match[2]) || 0);
  }, 0);

  let nextIndex = maxExistingNumber + 1;
  let candidate = `${prefix}${String(nextIndex).padStart(3, "0")}`;
  while (usedCodes.has(candidate.toUpperCase())) {
    nextIndex += 1;
    candidate = `${prefix}${String(nextIndex).padStart(3, "0")}`;
  }
  return candidate;
}

function getNextAvailableStudentToken() {
  const usedTokens = new Set(
    state.students
      .map((student) => String(student.token || "").trim().toLowerCase())
      .filter(Boolean)
  );
  const maxExistingNumber = state.students.reduce((max, student) => {
    const match = String(student.token || "").trim().toLowerCase().match(/^student(\d+)$/);
    return match ? Math.max(max, Number(match[1]) || 0) : max;
  }, 0);

  let nextIndex = maxExistingNumber + 1;
  let candidate = `student${String(nextIndex).padStart(3, "0")}`;
  while (usedTokens.has(candidate.toLowerCase())) {
    nextIndex += 1;
    candidate = `student${String(nextIndex).padStart(3, "0")}`;
  }
  return candidate;
}

async function handleCoachRosterAction(event) {
  const button = event.target.closest("[data-set-current-coach], [data-edit-coach], [data-toggle-coach-status], [data-delete-coach]");
  if (!button) {
    return;
  }

  if (button.dataset.setCurrentCoach) {
    const coachId = button.dataset.setCurrentCoach;
    const nextCoach = state.coaches.find((coach) => coach.id === coachId && coach.status !== "inactive");
    if (!nextCoach) {
      window.alert("這位教練目前已停用，請先啟用後再切換。");
      return;
    }
    state.currentCoachId = coachId;
    persistState();
    renderCurrentCoachOptions();
    renderCoachRoster();
    coachProgramEditorMode = "create";
    editingProgramId = null;
    seedEditorFromProgram(null);
    renderProgramLibrary();
    renderCoachStudentLinks();
    renderCoachStudentRoster();
    renderCoachToday();
    renderCoachHistoryFilters();
    renderCoachHistory();
    renderStudentProgramOptions();
    renderStudentHistoryFilters();
    renderStudentHistory();
    renderStudentSummary();
    return;
  }

  if (button.dataset.editCoach) {
    const coach = state.coaches.find((item) => item.id === button.dataset.editCoach);
    if (!coach) {
      return;
    }
    const nextName = window.prompt("請輸入新的教練姓名。", coach.name || "");
    if (nextName === null) {
      return;
    }
    const trimmed = nextName.trim();
    if (!trimmed) {
      window.alert("教練姓名不可空白。");
      return;
    }
    const duplicated = state.coaches.some((item) => item.id !== coach.id && item.name === trimmed);
    if (duplicated) {
      window.alert("已有同名教練，請換一個名稱。");
      return;
    }
    if (IS_CLOUD_MODE) {
      try {
        const payload = await callCloudApi("upsertCoach", {
          coach: {
            coach_id: coach.id,
            coach_name: trimmed,
            access_code: coach.accessCode,
            token: coach.token,
            status: coach.status,
            role: coach.role,
            last_used_at: coach.lastUsedAt
          }
        });
        applyCloudPayloadToState(payload);
      } catch (error) {
        console.warn("Cloud editCoach failed, falling back to local state:", error);
        coach.name = trimmed;
        state.programs = state.programs.map((program) =>
          program.coachId === coach.id ? { ...program, coachName: trimmed } : program
        );
        state.workoutLogs = state.workoutLogs.map((log) =>
          log.coachId === coach.id ? { ...log, coachName: trimmed } : log
        );
        persistState();
      }
    } else {
      coach.name = trimmed;
      state.programs = state.programs.map((program) =>
        program.coachId === coach.id ? { ...program, coachName: trimmed } : program
      );
      state.workoutLogs = state.workoutLogs.map((log) =>
        log.coachId === coach.id ? { ...log, coachName: trimmed } : log
      );
      persistState();
    }
    if (APP_MODE === "admin") {
      refreshAdminWorkspace();
    } else {
      renderCurrentCoachOptions();
      renderCoachRoster();
      renderProgramLibrary();
      renderCoachStudentLinks();
      renderCoachStudentRoster();
      renderCoachHistoryFilters();
      renderCoachToday();
      renderCoachHistory();
      renderStudentProgramOptions();
      renderStudentHistoryFilters();
      renderStudentHistory();
      renderStudentSummary();
      syncStudentAccessUI();
    }
    window.alert(`教練名稱已更新為 ${trimmed}。`);
    return;
  }

  if (button.dataset.toggleCoachStatus) {
    const coach = state.coaches.find((item) => item.id === button.dataset.toggleCoachStatus);
    if (!coach) {
      return;
    }
    const activeCoachCount = state.coaches.filter((item) => item.status !== "inactive").length;
    const nextStatus = coach.status === "inactive" ? "active" : "inactive";
    if (nextStatus === "inactive" && activeCoachCount <= 1) {
      window.alert("至少要保留一位啟用中的教練。");
      return;
    }
    if (IS_CLOUD_MODE) {
      try {
        const payload = await callCloudApi("setCoachStatus", {
          coachId: coach.id,
          status: nextStatus
        });
        applyCloudPayloadToState(payload);
      } catch (error) {
        console.warn("Cloud toggleCoachStatus failed, falling back to local state:", error);
        coach.status = nextStatus;
        if (state.currentCoachId === coach.id && nextStatus === "inactive") {
          const fallbackCoach = state.coaches.find((item) => item.status !== "inactive" && item.id !== coach.id);
          state.currentCoachId = fallbackCoach?.id || state.currentCoachId;
        }
        persistState();
      }
    } else {
      coach.status = nextStatus;
      if (state.currentCoachId === coach.id && nextStatus === "inactive") {
        const fallbackCoach = state.coaches.find((item) => item.status !== "inactive" && item.id !== coach.id);
        state.currentCoachId = fallbackCoach?.id || state.currentCoachId;
      }
      persistState();
    }
    if (APP_MODE === "admin") {
      refreshAdminWorkspace();
    } else {
      renderCurrentCoachOptions();
      renderCoachRoster();
      coachProgramEditorMode = "create";
      editingProgramId = null;
      seedEditorFromProgram(null);
      renderProgramLibrary();
      renderCoachStudentLinks();
      renderCoachStudentRoster();
      renderCoachToday();
      renderCoachHistoryFilters();
      renderCoachHistory();
      renderStudentProgramOptions();
      renderStudentHistoryFilters();
      renderStudentHistory();
      renderStudentSummary();
    }
    return;
  }

  if (button.dataset.deleteCoach) {
    const coach = state.coaches.find((item) => item.id === button.dataset.deleteCoach);
    if (!coach) {
      return;
    }
    if ((state.coaches || []).length <= 1) {
      window.alert("至少要保留一位教練。");
      return;
    }

    const linkedStudentCount = state.students.filter((student) => (student.primaryCoachId || "") === coach.id).length;
    const linkedProgramCount = state.programs.filter((program) => (program.coachId || "") === coach.id).length;
    const linkedLogCount = state.workoutLogs.filter((log) => (log.coachId || "") === coach.id).length;
    const confirmed = window.confirm(
      `確定要刪除 ${coach.name} 嗎？\n\n名下學生：${linkedStudentCount} 人\n課表：${linkedProgramCount} 份\n紀錄：${linkedLogCount} 筆\n\n刪除後，原本學生會顯示為「暫無教練」。`
    );
    if (!confirmed) {
      return;
    }

    if (IS_CLOUD_MODE) {
      try {
        const payload = await callCloudApi("deleteCoach", { coachId: coach.id });
        applyCloudPayloadToState(payload);
      } catch (error) {
        console.warn("Cloud deleteCoach failed, falling back to local state:", error);
        state.students = state.students.map((student) =>
          student.primaryCoachId === coach.id ? { ...student, primaryCoachId: "" } : student
        );
        state.programs = state.programs.map((program) =>
          program.coachId === coach.id ? { ...program, coachId: "" } : program
        );
        state.workoutLogs = state.workoutLogs.map((log) =>
          log.coachId === coach.id ? { ...log, coachId: "" } : log
        );
        state.coaches = state.coaches.filter((item) => item.id !== coach.id);

        if (state.currentCoachId === coach.id) {
          state.currentCoachId = state.coaches[0]?.id || "";
        }
        if (authenticatedCoachId === coach.id) {
          authenticatedCoachId = "";
        }

        persistState();
        persistSession();
      }
    } else {
      state.students = state.students.map((student) =>
        student.primaryCoachId === coach.id ? { ...student, primaryCoachId: "" } : student
      );
      state.programs = state.programs.map((program) =>
        program.coachId === coach.id ? { ...program, coachId: "" } : program
      );
      state.workoutLogs = state.workoutLogs.map((log) =>
        log.coachId === coach.id ? { ...log, coachId: "" } : log
      );
      state.coaches = state.coaches.filter((item) => item.id !== coach.id);

      if (state.currentCoachId === coach.id) {
        state.currentCoachId = state.coaches[0]?.id || "";
      }
      if (authenticatedCoachId === coach.id) {
        authenticatedCoachId = "";
      }

      persistState();
      persistSession();
    }
    if (APP_MODE === "admin") {
      refreshAdminWorkspace();
    } else {
      refreshCoachWorkspace();
      refreshStudentWorkspace();
    }
    window.alert(`已刪除 ${coach.name}，其名下學生目前為暫無教練。`);
  }
}

async function handleCoachStudentRosterAction(event) {
  const button = event.target.closest("[data-delete-student], [data-copy-student-code], [data-edit-student], [data-open-student-link], [data-open-student-history], [data-toggle-student-status], [data-assign-student-coach]");
  if (!button) {
    return;
  }

  if (button.dataset.openStudentHistory) {
    const studentId = button.dataset.openStudentHistory;
    switchCoachPanel("coach-history");
    els.coachHistoryStudent.value = studentId;
    els.coachHistoryExercise.value = "";
    renderCoachHistory();
    document.querySelector("#coach-history").scrollIntoView({ behavior: "smooth", block: "start" });
    return;
  }

  if (button.dataset.openStudentLink) {
    const studentName = button.dataset.studentName || "";
    if (!studentName) {
      return;
    }

    switchCoachPanel("coach-students");
    if (els.coachStudentLinkName) {
      els.coachStudentLinkName.value = studentName;
    }
    if (els.coachStudentLinkSelect) {
      els.coachStudentLinkSelect.value = studentName;
    }
    renderCoachStudentLinks();
    els.coachStudentLinkName?.scrollIntoView({ behavior: "smooth", block: "center" });
    els.coachStudentLinkName?.focus();
    return;
  }

  if (button.dataset.editStudent) {
    const studentId = button.dataset.editStudent;
    const student = state.students.find((item) => item.id === studentId);
    if (!student) {
      return;
    }

    const nextName = window.prompt("\u8acb\u8f38\u5165\u65b0\u7684\u5b78\u751f\u59d3\u540d\uff1a", student.name)?.trim();
    if (!nextName || nextName === student.name) {
      return;
    }

    const exists = state.students.some((item) => item.id !== studentId && String(item.name || "").trim() === nextName);
    if (exists) {
      window.alert("\u9019\u4f4d\u5b78\u751f\u540d\u7a31\u5df2\u7d93\u5b58\u5728\u3002");
      return;
    }
    if (IS_CLOUD_MODE) {
      try {
        const payload = await callCloudApi("upsertStudent", {
          student: {
            student_id: student.id,
            student_name: nextName,
            access_code: student.accessCode,
            token: student.token,
            status: student.status,
            primary_coach_id: student.primaryCoachId,
            last_used_at: student.lastUsedAt
          }
        });
        applyCloudPayloadToState(payload);
      } catch (error) {
        console.warn("Cloud editStudent failed, falling back to local state:", error);
        state.students = state.students.map((item) =>
          item.id === studentId
            ? { ...item, name: nextName }
            : item
        );
        state.workoutLogs = state.workoutLogs.map((log) =>
          log.studentId === studentId
            ? { ...log, studentName: nextName }
            : log
        );
        persistState();
      }
    } else {
      state.students = state.students.map((item) =>
        item.id === studentId
          ? { ...item, name: nextName }
          : item
      );
      state.workoutLogs = state.workoutLogs.map((log) =>
        log.studentId === studentId
          ? { ...log, studentName: nextName }
          : log
      );
      persistState();
    }

    if (currentStudentId === studentId) {
      const currentStudent = state.students.find((item) => item.id === studentId);
      if (currentStudent) {
        els.studentAccessCode.value = currentStudent.accessCode || els.studentAccessCode.value;
      }
    }

    if (APP_MODE === "admin") {
      refreshAdminWorkspace();
    } else {
      renderCoachStudentLinks();
      renderCoachStudentRoster();
      renderCoachToday();
      renderCoachHistoryFilters();
      renderCoachHistory();
      renderStudentProgramOptions();
      renderStudentHistoryFilters();
      renderStudentHistory();
      renderStudentSummary();
    }
    return;
  }

  if (button.dataset.assignStudentCoach) {
    openAssignCoachModal(button.dataset.assignStudentCoach);
    return;
  }

  if (button.dataset.toggleStudentStatus) {
    const studentId = button.dataset.toggleStudentStatus;
    const student = state.students.find((item) => item.id === studentId);
    if (!student) {
      return;
    }

    const nextStatus = student.status === "inactive" ? "active" : "inactive";
    if (IS_CLOUD_MODE) {
      try {
        const payload = await callCloudApi("setStudentStatus", {
          studentId,
          status: nextStatus
        });
        applyCloudPayloadToState(payload);
      } catch (error) {
        console.warn("Cloud toggleStudentStatus failed, falling back to local state:", error);
        state.students = state.students.map((item) =>
          item.id === studentId ? { ...item, status: nextStatus } : item
        );

        if (currentStudentId === studentId && nextStatus === "inactive") {
          currentStudentId = "";
          loadedStudentEntries = [];
          lastSubmittedLogs = [];
          studentSubmissionCompleted = false;
        }

        persistState();
        persistSession();
      }
    } else {
      state.students = state.students.map((item) =>
        item.id === studentId ? { ...item, status: nextStatus } : item
      );

      if (currentStudentId === studentId && nextStatus === "inactive") {
        currentStudentId = "";
        loadedStudentEntries = [];
        lastSubmittedLogs = [];
        studentSubmissionCompleted = false;
      }

      persistState();
      persistSession();
    }
    if (APP_MODE === "admin") {
      refreshAdminWorkspace();
    } else {
      renderCoachStudentLinks();
      renderCoachStudentRoster();
      renderCoachToday();
      renderCoachHistoryFilters();
      renderCoachHistory();
      renderStudentProgramOptions();
      renderStudentHistoryFilters();
      renderStudentHistory();
      renderStudentSummary();
      syncStudentAccessUI();
    }
    return;
  }

  if (button.dataset.copyStudentCode) {
    const code = button.dataset.studentCode || "";
    if (!code) {
      return;
    }

    const resetText = () => {
      window.setTimeout(() => {
        button.textContent = "\u8907\u88fd\u4ee3\u78bc";
      }, 1200);
    };

    if (navigator.clipboard?.writeText) {
      navigator.clipboard.writeText(code).then(() => {
        button.textContent = "\u5df2\u8907\u88fd";
        resetText();
      }).catch(() => {
        window.alert(`\u8acb\u624b\u52d5\u8907\u88fd\u4ee3\u78bc\uff1a${code}`);
      });
      return;
    }

    window.alert(`\u8acb\u624b\u52d5\u8907\u88fd\u4ee3\u78bc\uff1a${code}`);
    return;
  }

  const studentId = button.dataset.deleteStudent;
  const student = state.students.find((item) => item.id === studentId);
  if (!student) {
    return;
  }

  const hasLogs = state.workoutLogs.some((log) => log.studentId === studentId);
  const confirmed = window.confirm(
    hasLogs
      ? `${student.name} \u5df2\u7d93\u6709\u6b77\u53f2\u7d00\u9304\uff0c\u522a\u9664\u5f8c\u6703\u4e00\u8d77\u79fb\u9664\u8a72\u5b78\u751f\u7684\u6b77\u53f2\u8cc7\u6599\u3002\u78ba\u5b9a\u8981\u522a\u9664\u55ce\uff1f`
      : `\u78ba\u5b9a\u8981\u522a\u9664 ${student.name} \u55ce\uff1f`
  );

  if (!confirmed) {
    return;
  }

    if (IS_CLOUD_MODE) {
      try {
        const payload = await callCloudApi("deleteStudent", { studentId });
        applyCloudPayloadToState(payload);
      } catch (error) {
        console.warn("Cloud deleteStudent failed, falling back to local state:", error);
        state.students = state.students.filter((item) => item.id !== studentId);
        state.workoutLogs = state.workoutLogs.filter((log) => log.studentId !== studentId);

        if (currentStudentId === studentId) {
          currentStudentId = "";
          loadedStudentEntries = [];
          lastSubmittedLogs = [];
          studentSubmissionCompleted = false;
        }

        persistState();
        persistSession();
      }
    } else {
      state.students = state.students.filter((item) => item.id !== studentId);
      state.workoutLogs = state.workoutLogs.filter((log) => log.studentId !== studentId);

      if (currentStudentId === studentId) {
        currentStudentId = "";
        loadedStudentEntries = [];
        lastSubmittedLogs = [];
        studentSubmissionCompleted = false;
      }

      persistState();
      persistSession();
    }
  if (APP_MODE === "admin") {
    refreshAdminWorkspace();
  } else {
    renderCoachStudentLinks();
    renderCoachStudentRoster();
    renderCoachToday();
    renderCoachHistoryFilters();
    renderCoachHistory();
    renderStudentProgramOptions();
    renderStudentHistoryFilters();
    renderStudentHistory();
    renderStudentSummary();
    syncStudentAccessUI();
  }
}

function handleCoachStudentLinkAction(event) {
  const coachLeaveAnchor = event.target.closest("a[data-open-coach-leave='1']");
  if (coachLeaveAnchor) {
    event.preventDefault();
    const targetUrl = resolveCoachLeaveTargetUrl(coachLeaveAnchor.href);
    if (!targetUrl) {
      window.alert("請假系統入口尚未啟用，請先確認教練身份。");
      return;
    }
    window.open(targetUrl, "_blank", "noopener");
    return;
  }

  const button = event.target.closest("[data-copy-student-link], [data-copy-shared-link], [data-copy-leave-student-link], [data-copy-student-message], [data-copy-student-qr]");
  if (!button) {
    return;
  }

  if (button.dataset.copyStudentMessage) {
    const message = `${button.dataset.studentName}\u4f60\u597d\uff0c
\u4f60\u7684\u5c08\u5c6c\u7db2\u5740\uff1a${button.dataset.studentLink}
\u5982\u679c\u4f7f\u7528\u5171\u7528\u5165\u53e3\uff0c\u8acb\u524d\u5f80\uff1a${button.dataset.sharedUrl}
\u4e26\u8f38\u5165\u4ee3\u78bc\uff1a${button.dataset.studentCode}`;

    if (navigator.clipboard?.writeText) {
      navigator.clipboard.writeText(message).then(() => {
        button.textContent = "\u5df2\u8907\u88fd\u8a0a\u606f";
        window.setTimeout(() => {
          button.textContent = "\u8907\u88fd\u7d66\u5b78\u751f\u7684\u8a0a\u606f";
        }, 1200);
      }).catch(() => {
        window.alert("\u8acb\u624b\u52d5\u8907\u88fd\u7d66\u5b78\u751f\u7684\u8a0a\u606f\u3002");
      });
      return;
    }

    window.alert(message);
    return;
  }

  if (button.dataset.copyStudentQr) {
    const qrValue = button.dataset.qrLink || "";
    if (!qrValue) {
      return;
    }

    if (navigator.clipboard?.writeText) {
      navigator.clipboard.writeText(qrValue).then(() => {
        button.textContent = "已複製 QR 連結";
        window.setTimeout(() => {
          button.textContent = "複製 QR 連結內容";
        }, 1200);
      }).catch(() => {
        window.alert("請手動複製 QR 連結內容。");
      });
      return;
    }

    window.alert(qrValue);
    return;
  }

  const studentId = button.dataset.copyStudentLink || button.dataset.copySharedLink || button.dataset.copyLeaveStudentLink;
  const input = button.dataset.copyStudentLink
    ? els.coachStudentLinks.querySelector(`[data-link-value="${studentId}"]`)
    : button.dataset.copySharedLink
      ? els.coachStudentLinks.querySelector(`[data-shared-link-value="${studentId}"]`)
      : els.coachStudentLinks.querySelector(`[data-leave-student-link-value="${studentId}"]`);
  if (!input) {
    return;
  }

  const defaultLabel = button.dataset.copyStudentLink
    ? "\u8907\u88fd\u7db2\u5740"
    : button.dataset.copySharedLink
      ? "\u8907\u88fd\u5171\u7528\u5165\u53e3"
      : "複製請假學生頁";

  const resetText = () => {
    window.setTimeout(() => {
      button.textContent = defaultLabel;
    }, 1200);
  };

  if (navigator.clipboard?.writeText) {
    navigator.clipboard.writeText(input.value).then(() => {
      button.textContent = "\u5df2\u8907\u88fd";
      resetText();
    }).catch(() => {
      input.focus();
      input.select();
      window.alert("\u5df2\u9078\u53d6\u7db2\u5740\uff0c\u8acb\u624b\u52d5\u8907\u88fd\u3002");
    });
    return;
  }

  input.focus();
  input.select();
  window.alert("\u5df2\u9078\u53d6\u7db2\u5740\uff0c\u8acb\u624b\u52d5\u8907\u88fd\u3002");
}

function getCoachHistoryLogs() {
  const studentId = els.coachHistoryStudent?.value || "";
  const exercise = els.coachHistoryExercise?.value || "";
  const dateFrom = els.coachHistoryDateFrom?.value || "";
  const dateTo = els.coachHistoryDateTo?.value || "";

  return getCoachScopedLogs()
    .filter((log) => !studentId || log.studentId === studentId)
    .filter((log) => !exercise || log.exercise === exercise)
    .filter((log) => !dateFrom || getLogActivityDateKey(log) >= dateFrom)
    .filter((log) => !dateTo || getLogActivityDateKey(log) <= dateTo)
    .sort(compareLogsByDateTimeDesc);
}

function renderCoachHistory() {
  const logs = getCoachHistoryLogs();
  const visibleLogs = logs.slice(0, coachHistoryVisibleCount);

  if (els.coachHistoryCount) {
    els.coachHistoryCount.textContent = `共 ${visibleLogs.length} / ${logs.length} 筆`;
  }

  els.coachHistoryBody.innerHTML = visibleLogs.length
    ? visibleLogs
        .map(
          (log) => `
            <tr>
              <td>${formatDateDisplay(getLogActivityDateKey(log))}</td>
              <td>${log.studentName}</td>
              <td>${log.category}</td>
              <td>${log.exercise}</td>
              <td>${formatTarget(log)}</td>
              <td>${formatActual(log)}</td>
              <td>${log.studentNote || "-"}</td>
            </tr>
          `
        )
        .join("")
    : `<tr><td colspan="7" class="empty-state">目前沒有符合條件的紀錄。</td></tr>`;

  if (els.coachHistoryMore) {
    const remaining = logs.length - visibleLogs.length;
    const shouldHide = logs.length <= 6;
    els.coachHistoryMore.dataset.empty = shouldHide ? "true" : "false";
    els.coachHistoryMore.classList.toggle("is-hidden", shouldHide || (coachViewMode === "mobile" && !coachHistoryOpened));
    if (!shouldHide) {
      els.coachHistoryMore.textContent = remaining > 0 ? `查看更多（剩餘 ${remaining} 筆）` : "收合";
    }
  }
}

function resetCoachHistoryFilters() {
  if (els.coachHistoryStudent) {
    els.coachHistoryStudent.value = "";
  }
  if (els.coachHistoryExercise) {
    els.coachHistoryExercise.value = "";
  }
  if (els.coachHistoryDateFrom) {
    els.coachHistoryDateFrom.value = "";
  }
  if (els.coachHistoryDateTo) {
    els.coachHistoryDateTo.value = "";
  }
  coachHistoryVisibleCount = 6;
  renderCoachHistory();
}

function formatDateInputValue(date) {
  const { year, month, day } = getDateTimePartsInAppZone(date);
  return `${year}-${month}-${day}`;
}

function applyCoachHistoryLast30Days() {
  const end = new Date();
  const start = new Date(end.getTime() - 29 * 24 * 60 * 60 * 1000);

  if (els.coachHistoryDateFrom) {
    els.coachHistoryDateFrom.value = formatDateInputValue(start);
  }
  if (els.coachHistoryDateTo) {
    els.coachHistoryDateTo.value = formatDateInputValue(end);
  }

  coachHistoryVisibleCount = 6;
  renderCoachHistory();
}

function renderStudentHistoryFilters() {
  const student = getSelectedStudent();
  const exercises = [...new Set(
    state.workoutLogs
      .filter((log) => !student || log.studentId === student.id)
      .map((log) => log.exercise)
  )].sort();
  const currentValue = els.studentHistoryExercise.value;
  const nextValue = exercises.includes(currentValue) ? currentValue : "";
  els.studentHistoryExerciseOptions.innerHTML = exercises
    .map((exercise) => `<option value="${exercise}"></option>`)
    .join("");
  els.studentHistoryExercise.value = nextValue;
  if (nextValue !== currentValue) {
    studentHistoryVisibleCount = 1;
  }
}

function renderStudentHistory() {
  const student = getSelectedStudent();
  const exercise = els.studentHistoryExercise.value;

  if (!student) {
    if (els.studentHistoryCount) {
      els.studentHistoryCount.textContent = "共 0 筆";
    }
    els.studentHistoryBody.innerHTML = `<tr><td colspan="6" class="empty-state">\u8acb\u5148\u78ba\u8a8d\u5b78\u751f\u8eab\u4efd\u3002</td></tr>`;
    els.studentHistoryCardList.innerHTML = `<div class="empty-card">\u8acb\u5148\u78ba\u8a8d\u5b78\u751f\u8eab\u4efd\u3002</div>`;
    els.studentHistoryMore?.classList.add("is-hidden");
    return;
  }

  const logs = state.workoutLogs
    .filter((log) => log.studentId === student.id)
    .filter((log) => !exercise || log.exercise === exercise)
    .sort(compareLogsByDateTimeDesc);

  const visibleLogs = logs.slice(0, studentHistoryVisibleCount);

  if (els.studentHistoryCount) {
    els.studentHistoryCount.textContent = `共 ${visibleLogs.length} / ${logs.length} 筆`;
  }

  els.studentHistoryBody.innerHTML = visibleLogs.length
    ? visibleLogs
        .map(
          (log) => `
            <tr>
              <td>${formatDateDisplay(getLogActivityDateKey(log))}</td>
              <td>${log.category}</td>
              <td>${log.exercise}</td>
              <td>${formatTarget(log)}</td>
              <td>${formatActual(log)}</td>
              <td>${log.studentNote || "-"}</td>
            </tr>
          `
        )
        .join("")
    : `<tr><td colspan="6" class="empty-state">\u76ee\u524d\u6c92\u6709\u7b26\u5408\u689d\u4ef6\u7684\u6b77\u53f2\u7d00\u9304\u3002</td></tr>`;

  els.studentHistoryCardList.innerHTML = visibleLogs.length
    ? visibleLogs
        .map(
          (log) => `
            <article class="confirm-entry-card">
              <div class="confirm-entry-top">
                <span class="confirm-entry-category">${log.category}</span>
                <span class="confirm-entry-target">${formatDateDisplay(getLogActivityDateKey(log))}</span>
              </div>
              <h4>${log.exercise}</h4>
              <p class="confirm-entry-result">${formatActual(log)}</p>
              <p class="confirm-entry-note">${log.studentNote || "-"}</p>
            </article>
          `
        )
        .join("")
    : `<div class="empty-card">\u76ee\u524d\u6c92\u6709\u7b26\u5408\u689d\u4ef6\u7684\u6b77\u53f2\u7d00\u9304\u3002</div>`;

  if (els.studentHistoryMore) {
    const shouldShowButton = logs.length > 1;
    const hasMore = logs.length > studentHistoryVisibleCount;
    els.studentHistoryMore.classList.toggle("is-hidden", !shouldShowButton);
    const remainCount = Math.max(logs.length - visibleLogs.length, 0);
    els.studentHistoryMore.textContent = hasMore ? `查看更多（剩餘 ${remainCount} 筆）` : "收合";
  }
}

function toggleStudentHistoryVisibleCount() {
  const student = getSelectedStudent();
  const exercise = els.studentHistoryExercise?.value || "";
  const total = state.workoutLogs
    .filter((log) => !student || log.studentId === student.id)
    .filter((log) => !exercise || log.exercise === exercise).length;

  if (studentHistoryVisibleCount >= total) {
    studentHistoryVisibleCount = 1;
  } else {
    studentHistoryVisibleCount += 5;
  }

  renderStudentHistory();
}
function exportProgramImage() {
  const { program, items } = collectProgramPayload();

  if (!program.code || !items.length) {
    window.alert("\u8acb\u5148\u5b8c\u6210\u8ab2\u8868\u4ee3\u78bc\u8207\u8ab2\u8868\u9805\u76ee\uff0c\u518d\u532f\u51fa PNG\u3002");
    return;
  }

  const width = 1200;
  const headerHeight = 120;
  const rowHeight = 94;
  const height = headerHeight + items.length * rowHeight + 2;
  const canvas = document.createElement("canvas");
  const ctx = canvas.getContext("2d");

  canvas.width = width;
  canvas.height = height;
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, width, height);

  ctx.strokeStyle = "#29231c";
  ctx.lineWidth = 2;
  ctx.strokeRect(0, 0, width, height);

  ctx.font = "700 66px 'Noto Sans TC'";
  ctx.fillStyle = "#111111";
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillText(program.code, width / 2, headerHeight / 2);

  const categoryWidth = 260;
  const targetWidth = 220;
  const exerciseWidth = width - categoryWidth - targetWidth;

  items.forEach((item, index) => {
    const y = headerHeight + index * rowHeight;
    ctx.fillStyle = index % 2 === 0 ? "#d7d7d7" : "#ffffff";
    ctx.fillRect(0, y, width, rowHeight);

    ctx.fillStyle = "#171717";
    ctx.textBaseline = "middle";
    ctx.textAlign = "center";
    ctx.font = "500 38px 'Noto Sans TC'";
    ctx.fillText(item.category, categoryWidth / 2, y + rowHeight / 2);
    ctx.fillText(item.exercise, categoryWidth + exerciseWidth / 2, y + rowHeight / 2);
    ctx.fillText(formatTarget(item), width - targetWidth / 2, y + rowHeight / 2);
  });

  const link = document.createElement("a");
  link.href = canvas.toDataURL("image/png");
  link.download = `${program.code || "program"}.png`;
  link.click();
}

function getPublishedProgram(coachId = "", studentId = "") {
  const scopedPrograms = coachId
    ? state.programs.filter((program) => (program.coachId || "") === coachId)
    : state.programs;
  const availablePrograms = studentId
    ? scopedPrograms.filter((program) => isProgramAvailableForStudent(program, studentId))
    : scopedPrograms;
  const publishedPrograms = availablePrograms
    .filter((program) => program.published)
    .sort((a, b) => {
      if (studentId && isProgramTargeted(a) !== isProgramTargeted(b)) {
        return isProgramTargeted(a) ? -1 : 1;
      }
      const dateCompare = (getComparableDateKey(b.date) || "").localeCompare(getComparableDateKey(a.date) || "");
      if (dateCompare !== 0) {
        return dateCompare;
      }
      return (getComparableDateTimeKey(b.createdAt) || "").localeCompare(getComparableDateTimeKey(a.createdAt) || "");
    });
  return publishedPrograms[0] || null;
}

function getSelectedStudentProgram() {
  const student = getSelectedStudent();
  const studentCoachId = student?.primaryCoachId || "";
  const studentId = student?.id || "";
  const scopedPrograms = [...state.programs]
    .filter((program) => !studentCoachId || (program.coachId || "") === studentCoachId)
    .filter((program) => !studentId || isProgramAvailableForStudent(program, studentId))
    .sort((a, b) => (getComparableDateKey(b.date) || "").localeCompare(getComparableDateKey(a.date) || ""));
  const selectedProgramId = currentStudentProgramId || els.studentProgramSelect?.value || "";
  return getPublishedProgram(studentCoachId, studentId)
    || scopedPrograms.find((program) => program.id === selectedProgramId)
    || scopedPrograms[0]
    || null;
}

function getLatestSubmittedLogsForCurrentStudentProgram() {
  const student = getSelectedStudent();
  const program = getSelectedStudentProgram();
  if (!student || !program) {
    return [];
  }

  const matches = state.workoutLogs
    .filter((log) => log.studentId === student.id)
    .filter((log) => log.programId === program.id)
    .sort(compareLogsByDateTimeDesc);

  if (!matches.length) {
    return [];
  }

  const latestSubmittedAt = normalizeDateTimeValue(matches[0].submittedAt || "");
  if (!latestSubmittedAt) {
    return matches;
  }

  return matches.filter((log) => normalizeDateTimeValue(log.submittedAt || "") === latestSubmittedAt);
}

function findLatestReferenceLog(studentId, item) {
  return state.workoutLogs
    .filter((log) => log.studentId === studentId)
    .filter((log) => log.exercise === item.exercise)
    .filter((log) => log.targetType === item.targetType)
    .filter((log) => Number(log.targetSets) === Number(item.targetSets))
    .filter((log) => Number(log.targetValue) === Number(item.targetValue))
    .sort(compareLogsByDateTimeDesc)[0] || null;
}

function getRecommendedWeightText(entry) {
  const oneRm = getEntryBaseOneRm(entry);
  if (!oneRm) {
    return "";
  }

  const targetWeight = getRecommendedWeightFromOneRm(entry, oneRm);
  if (!targetWeight) {
    return "";
  }

  return `建議重量：${formatWeightToDisplay(targetWeight)} kg`;
}

function getEntryBaseOneRm(entry) {
  const student = getSelectedStudent();
  if (!student || entry.category !== "主項目") {
    return null;
  }

  const candidateLogs = state.workoutLogs
    .filter((log) => log.studentId === student.id)
    .filter((log) => log.exercise === entry.exercise)
    .filter((log) => Number(log.actualWeight) > 0)
    .sort(compareLogsByDateTimeDesc);

  for (const log of candidateLogs) {
    const inferred = inferOneRmFromLog(log);
    if (inferred) {
      return inferred;
    }
  }

  return null;
}

function inferOneRmFromLog(log) {
  const weight = Number(log.actualWeight || 0);
  if (!weight) {
    return null;
  }

  if (log.targetType === "rm" && Number(log.targetValue) === 1) {
    return weight;
  }

  if (log.targetType === "rm" && Number(log.targetValue) === 3) {
    return weight / 0.9;
  }

  if (log.targetType === "reps" && Number(log.targetSets) === 3 && Number(log.targetValue) === 5) {
    return weight / 0.85;
  }

  if (log.targetType === "reps" && Number(log.targetSets) === 5 && Number(log.targetValue) === 5) {
    return weight / 0.8;
  }

  return null;
}

function getRecommendedWeightFromOneRm(entry, oneRm) {
  if (entry.targetType === "rm" && Number(entry.targetValue) === 1) {
    return roundToNearestPlate(oneRm);
  }

  if (entry.targetType === "rm" && Number(entry.targetValue) === 3) {
    return roundToNearestPlate(oneRm * 0.9);
  }

  if (entry.targetType === "reps" && Number(entry.targetSets) === 3 && Number(entry.targetValue) === 5) {
    return roundToNearestPlate(oneRm * 0.85);
  }

  if (entry.targetType === "reps" && Number(entry.targetSets) === 5 && Number(entry.targetValue) === 5) {
    return roundToNearestPlate(oneRm * 0.8);
  }

  return null;
}

function roundToNearestPlate(value) {
  const numeric = Number(value || 0);
  if (!numeric) {
    return null;
  }
  return Math.round(numeric / 2.5) * 2.5;
}

function formatWeightToDisplay(value) {
  const numeric = Number(value || 0);
  if (!numeric) {
    return "";
  }
  return Number.isInteger(numeric) ? String(numeric) : numeric.toFixed(1).replace(/\.0$/, "");
}

function getProgramItems(programId) {
  return state.programItems
    .filter((item) => item.programId === programId)
    .sort((a, b) => a.sortOrder - b.sortOrder);
}

function getSelectedStudent() {
  return state.students.find((student) => student.id === currentStudentId) || null;
}

function resolveStudentByAccess(rawValue) {
  const value = normalizeAccessInput(rawValue).toLowerCase();
  if (!value) {
    return null;
  }

  return state.students.find((student) => {
    const code = String(student.accessCode || "").trim().toLowerCase();
    const token = String(student.token || "").trim().toLowerCase();
    return student.status !== "inactive" && (value === code || value === token || value === String(student.id).toLowerCase());
  }) || null;
}

function exportCoachHistoryCsv() {
  const logs = getCoachHistoryLogs();

  if (!logs.length) {
    window.alert("\u76ee\u524d\u6c92\u6709\u53ef\u532f\u51fa\u7684\u6b77\u53f2\u7d00\u9304\u3002");
    return;
  }

  const rows = [
    ["日期", "學生", "分類", "動作", "目標", "實際結果", "備註"],
    ...logs.map((log) => [
      formatDateDisplay(getLogActivityDateKey(log)),
      log.studentName,
      log.category,
      log.exercise,
      formatTarget(log),
      formatActual(log),
      log.studentNote || ""
    ])
  ];
  const csv = rows
    .map((row) => row.map((cell) => `"${String(cell).replace(/"/g, '""')}"`).join(","))
    .join("\n");

  const blob = new Blob(["\ufeff" + csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "coach-history-export.csv";
  link.click();
  URL.revokeObjectURL(url);
}

async function hydrateStudentAccessFromUrl() {
  const params = new URLSearchParams(window.location.search);
  const accessValue = normalizeAccessInput(params.get("student") || params.get("token") || params.get("code") || "");
  if (!accessValue) {
    return;
  }
  let matchedStudent = null;
  if (IS_CLOUD_MODE) {
    try {
      matchedStudent = await resolveStudentAccessFromCloud(accessValue);
    } catch (error) {
      console.warn("Student cloud hydrate failed, falling back to local state:", error);
      matchedStudent = resolveStudentByAccess(accessValue);
    }
  } else {
    matchedStudent = resolveStudentByAccess(accessValue);
  }

  if (matchedStudent) {
    currentStudentId = matchedStudent.id;
    const normalizedAccessValue = accessValue;
    const matchedCode = String(matchedStudent.accessCode || "").trim();
    const matchedToken = String(matchedStudent.token || "").trim();
    currentStudentAccess = matchedCode || normalizedAccessValue || matchedToken;
    els.studentAccessCode.value = matchedCode || matchedToken || normalizedAccessValue;
    persistSession();
    return;
  }
  if (!currentStudentId) {
    currentStudentAccess = "";
    persistSession();
  }
}

function formatTarget(item) {
  if (item.targetType === "time") {
    return `${item.targetSets}x${item.targetValue}sec`;
  }
  if (item.targetType === "rm") {
    return item.targetValue ? `${item.targetValue}RM` : "RM";
  }
  return `${item.targetSets}x${item.targetValue}`;
}

function formatActual(log) {
  if (log.targetType === "time") {
    return `${log.actualSets || 0} \u7d44`;
  }
  if (log.targetType === "rm") {
    return `${log.actualWeight || 0} kg`;
  }
  return `${log.actualWeight || 0} kg / ${log.actualSets || 0} \u7d44 / ${log.actualReps || 0} \u6b21`;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function escapeXml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function truncateForSheet(value, maxLength) {
  const text = String(value ?? "");
  return text.length > maxLength ? `${text.slice(0, maxLength - 1)}…` : text;
}

function getDateTimePartsInAppZone(date = new Date()) {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: APP_TIME_ZONE,
    hourCycle: "h23",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit"
  }).formatToParts(date);

  const map = {};
  parts.forEach((part) => {
    map[part.type] = part.value;
  });

  return {
    year: map.year || "0000",
    month: map.month || "01",
    day: map.day || "01",
    hour: map.hour || "00",
    minute: map.minute || "00",
    second: map.second || "00"
  };
}

function getTodayDateInAppZone() {
  const { year, month, day } = getDateTimePartsInAppZone(new Date());
  return `${year}-${month}-${day}`;
}

function timestampNow() {
  const { year, month, day, hour, minute } = getDateTimePartsInAppZone(new Date());
  return `${year}-${month}-${day} ${hour}:${minute}`;
}
