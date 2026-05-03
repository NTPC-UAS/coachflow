/**
 * CoachFlow Leave Bridge (No Google Sheet dependency)
 *
 * Supported actions:
 * - bootstrap / ping
 * - checkEvent
 * - listEvents
 * - createEvent
 * - updateEvent
 * - deleteEvent
 * - deleteSingleEvent
 * - listLeaveRecords
 * - saveLeaveRecord
 * - listBillingProfiles
 * - saveBillingProfile
 * - sendEmail
 *
 * Optional Script Property:
 * - DEFAULT_CALENDAR_ID: fallback calendar id used when payload.calendarId is empty.
 */

const APP_TIME_ZONE = "Asia/Taipei";

function doGet(e) {
  try {
    const action = getAction_(e, "bootstrap");
    return jsonResponse_(handleAction_(action, e && e.parameter ? e.parameter : {}));
  } catch (error) {
    return jsonResponse_(failResponse_(error));
  }
}

function doPost(e) {
  try {
    const payload = parsePostBody_(e);
    const action = String(payload.action || "").trim() || "bootstrap";
    return jsonResponse_(handleAction_(action, payload));
  } catch (error) {
    return jsonResponse_(failResponse_(error));
  }
}

function handleAction_(action, payload) {
  switch (String(action || "").trim()) {
    case "bootstrap":
    case "ping":
      return {
        ok: true,
        action: action,
        message: "bridge ok",
        timeZone: APP_TIME_ZONE,
        now: nowIso_()
      };

    case "createEvent":
      return createCalendarEvent_(payload || {});

    case "updateEvent":
      return updateCalendarEvent_(payload || {});

    case "checkEvent":
      return checkCalendarEvent_(payload || {});

    case "listEvents":
      return listCalendarEvents_(payload || {});

    case "deleteEvent":
      return deleteCalendarEvent_(payload || {});

    case "deleteSingleEvent":
      return deleteCalendarEvent_(Object.assign({}, payload || {}, {
        deleteScope: "single",
        singleEventOnly: true,
        strictSingleOccurrence: true
      }));

    case "listLeaveRecords":
      return listLeaveRecords_(payload || {});

    case "saveLeaveRecord":
      return saveLeaveRecord_(payload || {});

    case "listBillingProfiles":
      return listBillingProfiles_(payload || {});

    case "saveBillingProfile":
      return saveBillingProfile_(payload || {});

    case "sendEmail":
      return sendEmailNotice_(payload || {});

    default:
      return {
        ok: false,
        action: action,
        message: "Unsupported action. Use bootstrap/ping/checkEvent/listEvents/createEvent/updateEvent/deleteEvent/deleteSingleEvent/listLeaveRecords/saveLeaveRecord/listBillingProfiles/saveBillingProfile/sendEmail."
      };
  }
}

function checkCalendarEvent_(payload) {
  const calendar = resolveCalendar_(payload);
  if (!calendar) {
    return {
      ok: false,
      action: "checkEvent",
      message: "No available calendar. Check calendar access or DEFAULT_CALENDAR_ID."
    };
  }

  const rawEventId = safeString_(payload.eventId || payload.calendarEventId, "");
  if (!rawEventId) {
    return {
      ok: true,
      action: "checkEvent",
      exists: false,
      eventId: "",
      calendarId: calendar.getId(),
      message: "Missing eventId."
    };
  }

  const event = findEventById_(calendar, rawEventId);
  return {
    ok: true,
    action: "checkEvent",
    exists: Boolean(event),
    eventId: rawEventId,
    calendarId: calendar.getId(),
    message: event ? "Event exists." : "Event not found."
  };
}

function listCalendarEvents_(payload) {
  const calendar = resolveCalendar_(payload);
  if (!calendar) {
    return {
      ok: false,
      action: "listEvents",
      message: "No available calendar. Check calendar access or DEFAULT_CALENDAR_ID."
    };
  }

  const startAt = new Date(payload.startAt || "");
  const endAt = new Date(payload.endAt || "");
  if (String(startAt) === "Invalid Date" || String(endAt) === "Invalid Date") {
    return { ok: false, action: "listEvents", message: "Invalid startAt/endAt." };
  }
  if (endAt <= startAt) {
    return { ok: false, action: "listEvents", message: "endAt must be later than startAt." };
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
  if (isNormalLeaveCalendarCreateBlocked_(payload)) {
    return {
      ok: false,
      action: "createEvent",
      blocked: true,
      message: "Normal student leave must delete the original lesson event; creating a new leave event is blocked."
    };
  }

  const calendar = resolveCalendar_(payload);
  if (!calendar) {
    return {
      ok: false,
      message: "No available calendar. Check calendar access or DEFAULT_CALENDAR_ID."
    };
  }

  const startAt = new Date(payload.startAt || "");
  const endAt = new Date(payload.endAt || "");
  if (String(startAt) === "Invalid Date" || String(endAt) === "Invalid Date") {
    return { ok: false, message: "Invalid startAt/endAt." };
  }
  if (endAt <= startAt) {
    return { ok: false, message: "endAt must be later than startAt." };
  }

  const title = safeString_(payload.title, "CoachFlow Lesson");
  const options = {
    description: buildEventDescription_(payload)
  };
  if (payload.location) {
    options.location = String(payload.location);
  }
  if (payload.guests) {
    options.guests = String(payload.guests);
  }
  if (payload.sendInvites !== undefined) {
    options.sendInvites = toBoolean_(payload.sendInvites);
  }

  const event = calendar.createEvent(title, startAt, endAt, options);
  return {
    ok: true,
    action: "createEvent",
    eventId: event.getId(),
    calendarId: calendar.getId(),
    startAt: startAt.toISOString(),
    endAt: endAt.toISOString(),
    message: "createEvent success"
  };
}

function updateCalendarEvent_(payload) {
  const calendar = resolveCalendar_(payload);
  if (!calendar) {
    return {
      ok: false,
      action: "updateEvent",
      message: "No available calendar. Check calendar access or DEFAULT_CALENDAR_ID."
    };
  }

  const rawEventId = safeString_(payload.eventId || payload.calendarEventId, "");
  if (!rawEventId) {
    return { ok: false, action: "updateEvent", message: "Missing eventId." };
  }

  const event = findEventById_(calendar, rawEventId);
  if (!event) {
    return {
      ok: false,
      action: "updateEvent",
      eventId: rawEventId,
      calendarId: calendar.getId(),
      message: "Event not found."
    };
  }

  const beforeTitle = String(event.getTitle() || "");
  const title = safeString_(payload.title || payload.summary || payload.name, beforeTitle);
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
    eventId: rawEventId,
    calendarId: calendar.getId(),
    beforeTitle: beforeTitle,
    title: String(event.getTitle() || ""),
    message: "updateEvent success"
  };
}

function deleteCalendarEvent_(payload) {
  const calendar = resolveCalendar_(payload);
  if (!calendar) {
    return {
      ok: false,
      message: "No available calendar. Check calendar access or DEFAULT_CALENDAR_ID."
    };
  }

  const rawEventId = safeString_(payload.eventId || payload.calendarEventId, "");
  if (!rawEventId) {
    return { ok: false, message: "Missing eventId." };
  }

  const singleDelete = isSingleEventDeleteRequested_(payload);
  const event = singleDelete
    ? findSingleOccurrenceForDelete_(calendar, payload, rawEventId)
    : findEventById_(calendar, rawEventId);
  if (!event) {
    if (singleDelete) {
      return {
        ok: false,
        action: "deleteEvent",
        eventId: rawEventId,
        calendarId: calendar.getId(),
        blocked: true,
        singleEventOnly: true,
        message: "Single occurrence was not confirmed; deleteEvent blocked to avoid deleting a recurring series."
      };
    }
    return {
      ok: true,
      action: "deleteEvent",
      eventId: rawEventId,
      calendarId: calendar.getId(),
      skipped: true,
      message: "Event not found, treated as already deleted."
    };
  }

  event.deleteEvent();
  return {
    ok: true,
    action: "deleteEvent",
    eventId: String(event.getId() || rawEventId),
    requestedEventId: rawEventId,
    calendarId: calendar.getId(),
    singleEventOnly: singleDelete,
    message: "deleteEvent success"
  };
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
      message: "Missing leave record id or lessonKey."
    };
  }
  const records = getLeaveRecords_();
  const recordKey = getLeaveRecordKey_(record);
  let replaced = false;
  const nextRecords = records.map(function(item) {
    if (getLeaveRecordKey_(item) === recordKey) {
      replaced = true;
      return Object.assign({}, item, record, {
        updatedAt: nowIso_()
      });
    }
    return item;
  });
  if (!replaced) {
    nextRecords.unshift(Object.assign({}, record, {
      createdAt: record.createdAt || nowIso_(),
      updatedAt: nowIso_()
    }));
  }
  setLeaveRecords_(nextRecords.slice(0, 500));
  return {
    ok: true,
    action: "saveLeaveRecord",
    record: record,
    replaced: replaced
  };
}

function sendEmailNotice_(payload) {
  const to = safeString_(payload.to, "");
  if (!to) {
    return { ok: false, message: "Missing recipient email: to." };
  }

  const subject = safeString_(payload.subject, "[CoachFlow] Notification");
  const textBody = safeString_(payload.body, buildEmailBody_(payload));
  const options = {
    name: safeString_(payload.fromName, "CoachFlow")
  };
  if (payload.htmlBody) {
    options.htmlBody = String(payload.htmlBody);
  }
  if (payload.cc) {
    options.cc = String(payload.cc);
  }
  if (payload.bcc) {
    options.bcc = String(payload.bcc);
  }
  if (payload.replyTo) {
    options.replyTo = String(payload.replyTo);
  }

  MailApp.sendEmail(to, subject, textBody, options);
  return {
    ok: true,
    action: "sendEmail",
    to: to,
    subject: subject,
    message: "sendEmail success"
  };
}

function resolveCalendar_(payload) {
  const props = PropertiesService.getScriptProperties();
  const fallbackCalendarId = safeString_(props.getProperty("DEFAULT_CALENDAR_ID"), "");
  const requestedCalendarId = safeString_(payload.calendarId || payload.coachCalendarId, "");
  const calendarId = requestedCalendarId || fallbackCalendarId;

  if (calendarId) {
    const byId = CalendarApp.getCalendarById(calendarId);
    if (byId) {
      return byId;
    }
  }
  return CalendarApp.getDefaultCalendar();
}

function findEventById_(calendar, eventId) {
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

  for (var i = 0; i < candidates.length; i += 1) {
    const found = calendar.getEventById(candidates[i]);
    if (found) {
      return found;
    }
  }
  return null;
}

function isSingleEventDeleteRequested_(payload) {
  const scope = safeString_(payload.deleteScope || payload.deleteMode, "").toLowerCase();
  return toBoolean_(payload.singleEventOnly) || toBoolean_(payload.strictSingleOccurrence) || scope === "single" || scope === "singleevent";
}

function isNormalLeaveCalendarCreateBlocked_(payload) {
  const reason = safeString_(payload.reason, "").toLowerCase();
  const attendanceStatus = safeString_(payload.attendanceStatus, "").toLowerCase();
  const leaveType = safeString_(payload.leaveType || payload.type, "").toLowerCase();
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

  const dateKey = safeString_(
    payload.dateKey || payload.occurrenceDate,
    Utilities.formatDate(occurrenceStart, APP_TIME_ZONE, "yyyy-MM-dd")
  );
  const dayStart = new Date(dateKey + "T00:00:00+08:00");
  const dayEnd = new Date(dayStart.getTime() + 24 * 60 * 60 * 1000);
  const requestedId = normalizeCalendarEventId_(rawEventId);
  const studentCode = safeString_(payload.studentCode, "").toLowerCase();
  const studentName = safeString_(payload.studentName, "").toLowerCase();
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

function normalizeCode_(value) {
  return safeString_(value, "").toUpperCase();
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

function getLeaveRecords_() {
  const raw = readChunkedProperty_("LEAVE_RECORDS_V1");
  if (!raw) {
    return [];
  }
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed.map(normalizeLeaveRecord_) : [];
  } catch (error) {
    return [];
  }
}

function setLeaveRecords_(records) {
  writeChunkedProperty_("LEAVE_RECORDS_V1", JSON.stringify(records || []));
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
    updatedAt: safeString_(profile.updatedAt || profile.billingUpdatedAt, nowIso_()),
    updatedBy: normalizeCode_(profile.updatedBy || profile.billingUpdatedBy || "SYSTEM"),
    createdAt: safeString_(profile.createdAt, "")
  };
}

function getBillingProfileKey_(profile) {
  return normalizeCode_(profile.studentCode);
}

function getBillingProfiles_() {
  const raw = readChunkedProperty_("BILLING_PROFILES_V1");
  if (!raw) {
    return [];
  }
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed.map(normalizeBillingProfile_) : [];
  } catch (error) {
    return [];
  }
}

function setBillingProfiles_(profiles) {
  writeChunkedProperty_("BILLING_PROFILES_V1", JSON.stringify(profiles || []));
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

function readChunkedProperty_(baseKey) {
  const props = PropertiesService.getScriptProperties();
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
  const previousCount = Number(props.getProperty(baseKey + "_CHUNKS") || 0);
  for (var oldIndex = 0; oldIndex < previousCount; oldIndex += 1) {
    props.deleteProperty(baseKey + "_" + oldIndex);
  }
  props.deleteProperty(baseKey);

  const chunkSize = 8000;
  const chunkCount = Math.max(1, Math.ceil(text.length / chunkSize));
  for (var index = 0; index < chunkCount; index += 1) {
    props.setProperty(baseKey + "_" + index, text.slice(index * chunkSize, (index + 1) * chunkSize));
  }
  props.setProperty(baseKey + "_CHUNKS", String(chunkCount));
}

function buildEventDescription_(payload) {
  const rows = [
    "Source: CoachFlow Leave Bridge",
    "Student: " + safeString_(payload.studentCode || payload.studentName, "-"),
    "Coach: " + safeString_(payload.coachCode || payload.coachName, "-"),
    "Type: " + safeString_(payload.sourceType, "-"),
    "Reason: " + safeString_(payload.reason, "-")
  ];
  return rows.join("\n");
}

function buildEmailBody_(payload) {
  const rows = [
    "CoachFlow notification",
    "",
    "Template: " + safeString_(payload.template, "-"),
    "Student: " + safeString_(payload.studentCode || payload.studentName, "-"),
    "Coach: " + safeString_(payload.coachCode || payload.coachName, "-"),
    "When: " + safeString_(payload.when || payload.startAt, "-")
  ];
  return rows.join("\n");
}

function parsePostBody_(e) {
  const params = (e && e.parameter) || {};
  const body = (e && e.postData && e.postData.contents) || "";

  if (!body) {
    return params;
  }

  try {
    return JSON.parse(body);
  } catch (error) {
    // Some clients post payload JSON in query param "payload".
    if (params.payload) {
      try {
        return JSON.parse(params.payload);
      } catch (nestedError) {
        return params;
      }
    }
    return params;
  }
}

function getAction_(e, fallback) {
  return (e && e.parameter && e.parameter.action) || fallback || "";
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function failResponse_(error) {
  return {
    ok: false,
    message: error && error.message ? String(error.message) : String(error)
  };
}

function safeString_(value, fallback) {
  const text = value === null || value === undefined ? "" : String(value).trim();
  return text || (fallback || "");
}

function toBoolean_(value) {
  if (typeof value === "boolean") {
    return value;
  }
  const text = String(value || "").trim().toLowerCase();
  return text === "true" || text === "1" || text === "yes" || text === "y";
}

function nowIso_() {
  return new Date().toISOString();
}
