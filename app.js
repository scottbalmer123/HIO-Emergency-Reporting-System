const INCIDENT_STORAGE_KEY = "mine-emergency-incidents-v1";
const PERSONNEL_STORAGE_KEY = "mine-emergency-personnel-v1";
const SETTINGS_STORAGE_KEY = "mine-emergency-settings-v1";
const API_BASE = "/api";
const VALID_VIEWS = new Set(["create", "incidents", "reporting", "personnel", "settings"]);
const DEFAULT_RADIO_STATUS = "Entries appear below in time order and are saved with the incident.";
const DEFAULT_RADIO_UNITS = [
  "ERT Buggy 1",
  "Rescue 1",
  "Ambulance 2",
  "Surface Supervisor",
  "Traffic Control 1",
  "Fire Response 1",
  "Incident Control Vehicle",
];
const RADIO_CHANNELS = [
  "ESO Mine 1",
  "ESO Mine 2",
  "ESO Port 1",
  "ESO Port 2",
  "Emergency Mine",
  "Emergency Port",
  "Incident Management",
  "BA OPS 1 Mine",
  "BA OPS 2 Mine",
  "BA OPS 1 Port",
  "BA OPS 2 Port",
];
const CREW_OPTIONS = [
  "SV A",
  "SV B",
  "ESO A",
  "ESO B",
  "ESO C",
  "ESO D",
  "ICP A",
  "ICP B",
  "ICP C",
  "ICP D",
  "RN A",
  "RN B",
  "RN C",
  "RN D",
  "Security A",
  "Security B",
  "Security C",
  "Security D",
  "Shut",
  "Casual",
];
const REPORTING_PRESETS = new Set(["last7", "last30", "last90", "ytd", "all", "custom"]);

const severityRank = {
  Low: 0,
  Moderate: 1,
  High: 2,
  Critical: 3,
};

const state = {
  incidents: [],
  personnel: [],
  currentView: "create",
  storageMode: "local",
  selectedIncidentId: null,
  editingIncidentId: null,
  editingWorkerId: null,
  selectedPersonnelIds: [],
  currentRadioLogs: [],
  editingRadioLogId: null,
  currentIncidentPhotos: [],
  settings: {
    radioUnits: [],
  },
  reportingFilters: null,
  reportingSnapshot: null,
};

let previousDocumentTitle = document.title;

const dom = {
  navButtons: [...document.querySelectorAll("[data-view-target]")],
  views: [...document.querySelectorAll(".view")],
  metricOpen: document.querySelector("#metric-open"),
  metricHigh: document.querySelector("#metric-high"),
  metricMedical: document.querySelector("#metric-medical"),
  metricPending: document.querySelector("#metric-pending"),
  metricPersonnel: document.querySelector("#metric-personnel"),
  storageModeIndicator: document.querySelector("#storage-mode-indicator"),
  workspaceTitle: document.querySelector("#workspace-title"),
  incidentForm: document.querySelector("#incident-form"),
  incidentFormStatus: document.querySelector("#incident-form-status"),
  startNewIncidentButton: document.querySelector("#start-new-incident-btn"),
  openRegisterButton: document.querySelector("#open-register-btn"),
  createFromRegisterButton: document.querySelector("#create-from-register-btn"),
  copyReportingSummaryButton: document.querySelector("#copy-reporting-summary-btn"),
  exportReportingCsvButton: document.querySelector("#export-reporting-csv-btn"),
  openCreateFromPersonnelButton: document.querySelector("#open-create-from-personnel-btn"),
  managePersonnelButton: document.querySelector("#manage-personnel-btn"),
  resetFormButton: document.querySelector("#reset-form-btn"),
  incidentSearch: document.querySelector("#incident-search"),
  incidentList: document.querySelector("#incident-list"),
  detailTitle: document.querySelector("#detail-title"),
  editSelectedIncidentButton: document.querySelector("#edit-selected-incident-btn"),
  exportJsonButton: document.querySelector("#export-json-btn"),
  copySummaryButton: document.querySelector("#copy-summary-btn"),
  printReportButton: document.querySelector("#print-report-btn"),
  reportEmpty: document.querySelector("#report-empty"),
  reportContent: document.querySelector("#report-content"),
  reportTitle: document.querySelector("#report-title"),
  reportStatusBadge: document.querySelector("#report-status-badge"),
  reportGrid: document.querySelector("#report-grid"),
  reportSummary: document.querySelector("#report-summary"),
  reportAssignedPersonnel: document.querySelector("#report-assigned-personnel"),
  reportTimeline: document.querySelector("#report-timeline"),
  reportRadioTimeline: document.querySelector("#report-radio-timeline"),
  reportNotes: document.querySelector("#report-notes"),
  reportPhotos: document.querySelector("#report-photos"),
  reportGeneratedAt: document.querySelector("#report-generated-at"),
  hioExportPages: document.querySelector("#hio-export-pages"),
  hioExportPage1: document.querySelector("#hio-export-page-1"),
  hioExportPage2: document.querySelector("#hio-export-page-2"),
  draftTitle: document.querySelector("#draft-title"),
  draftSummaryText: document.querySelector("#draft-summary-text"),
  draftBadges: document.querySelector("#draft-badges"),
  draftMiniMetrics: document.querySelector("#draft-mini-metrics"),
  draftAssignedList: document.querySelector("#draft-assigned-list"),
  draftRadioPreview: document.querySelector("#draft-radio-preview"),
  personnelPickerToggle: document.querySelector("#personnel-picker-toggle"),
  personnelPickerPanel: document.querySelector("#personnel-picker-panel"),
  personnelPickerOptions: document.querySelector("#personnel-picker-options"),
  assignedPersonnelChips: document.querySelector("#assigned-personnel-chips"),
  radioLogTime: document.querySelector("#radio-log-time"),
  radioLogUnit: document.querySelector("#radio-log-unit"),
  radioLogMovement: document.querySelector("#radio-log-movement"),
  radioLogChannel: document.querySelector("#radio-log-channel"),
  radioLogLocation: document.querySelector("#radio-log-location"),
  radioLogNotes: document.querySelector("#radio-log-notes"),
  radioLogList: document.querySelector("#radio-log-list"),
  radioLogStatus: document.querySelector("#radio-log-status"),
  incidentPhotoTriggerButton: document.querySelector("#incident-photo-trigger-btn"),
  incidentPhotoUpload: document.querySelector("#incident-photo-upload"),
  incidentPhotoStatus: document.querySelector("#incident-photo-status"),
  incidentPhotoList: document.querySelector("#incident-photo-list"),
  saveRadioLogButton: document.querySelector("#save-radio-log-btn"),
  cancelRadioEditButton: document.querySelector("#cancel-radio-edit-btn"),
  personnelForm: document.querySelector("#personnel-form"),
  personnelWorkspaceTitle: document.querySelector("#personnel-workspace-title"),
  personnelFormStatus: document.querySelector("#personnel-form-status"),
  resetPersonnelButton: document.querySelector("#reset-personnel-btn"),
  personnelSearch: document.querySelector("#personnel-search"),
  personnelCountLabel: document.querySelector("#personnel-count-label"),
  personnelList: document.querySelector("#personnel-list"),
  settingsForm: document.querySelector("#settings-form"),
  settingsFormStatus: document.querySelector("#settings-form-status"),
  settingsRadioUnits: document.querySelector("#settings-radio-units"),
  resetSettingsButton: document.querySelector("#reset-settings-btn"),
  settingsChannelList: document.querySelector("#settings-channel-list"),
  reportingForm: document.querySelector("#reporting-form"),
  reportingFormStatus: document.querySelector("#reporting-form-status"),
  reportingPreset: document.querySelector("#reporting-preset"),
  reportingLocationFilter: document.querySelector("#reporting-location-filter"),
  reportingDateFrom: document.querySelector("#reporting-date-from"),
  reportingDateTo: document.querySelector("#reporting-date-to"),
  reportingStatusFilter: document.querySelector("#reporting-status-filter"),
  reportingSeverityFilter: document.querySelector("#reporting-severity-filter"),
  reportingIncidentTypeFilter: document.querySelector("#reporting-incident-type-filter"),
  resetReportingButton: document.querySelector("#reset-reporting-btn"),
  reportingTitle: document.querySelector("#reporting-title"),
  reportingSummary: document.querySelector("#reporting-summary"),
  reportingCountBadge: document.querySelector("#reporting-count-badge"),
  reportingGeneratedAt: document.querySelector("#reporting-generated-at"),
  reportingMetrics: document.querySelector("#reporting-metrics"),
  reportingBreakdowns: document.querySelector("#reporting-breakdowns"),
  reportingPerformance: document.querySelector("#reporting-performance"),
  reportingPersonnel: document.querySelector("#reporting-personnel"),
  reportingIncidents: document.querySelector("#reporting-incidents"),
  incidentTemplate: document.querySelector("#incident-item-template"),
  personnelTemplate: document.querySelector("#personnel-item-template"),
};

function toLocalDateTimeValue(date = new Date()) {
  return new Date(date.getTime() - date.getTimezoneOffset() * 60000).toISOString().slice(0, 16);
}

function toLocalDateValue(date = new Date()) {
  return new Date(date.getTime() - date.getTimezoneOffset() * 60000).toISOString().slice(0, 10);
}

function createIncidentId(date = new Date()) {
  return `INC-${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(
    2,
    "0"
  )}-${String(date.getHours()).padStart(2, "0")}${String(date.getMinutes()).padStart(2, "0")}`;
}

function createRadioLogId() {
  return `radio-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function createWorkerId() {
  return `worker-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function createIncidentPhotoId() {
  return `photo-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function normalizeNumber(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : 0;
}

function normalizeBoolean(value, fallback = false) {
  if (typeof value === "boolean") {
    return value;
  }

  if (typeof value === "string") {
    if (value === "true") {
      return true;
    }

    if (value === "false") {
      return false;
    }
  }

  if (value == null) {
    return fallback;
  }

  return Boolean(value);
}

function singleLine(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function formatDateTime(value) {
  if (!value) {
    return "Not recorded";
  }

  const date = new Date(value);
  if (Number.isNaN(date.valueOf())) {
    return value;
  }

  return new Intl.DateTimeFormat(undefined, {
    dateStyle: "medium",
    timeStyle: "short",
  }).format(date);
}

function formatDateOnly(value) {
  if (!value) {
    return "Not recorded";
  }

  const normalizedValue = /^\d{4}-\d{2}-\d{2}$/.test(value) ? `${value}T00:00` : value;
  const date = new Date(normalizedValue);
  if (Number.isNaN(date.valueOf())) {
    return value;
  }

  return new Intl.DateTimeFormat("en-AU", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
  }).format(date);
}

function formatTimeOnly(value) {
  if (!value) {
    return "Not recorded";
  }

  if (/^\d{2}:\d{2}/.test(value)) {
    return value.slice(0, 5);
  }

  const date = new Date(value);
  if (Number.isNaN(date.valueOf())) {
    return value;
  }

  return new Intl.DateTimeFormat("en-AU", {
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  }).format(date);
}

function parseDateValue(value) {
  if (!value) {
    return null;
  }

  if (typeof value === "string" && /^\d{1,2}:\d{2}$/.test(value.trim())) {
    return null;
  }

  const normalizedValue = typeof value === "string" && /^\d{4}-\d{2}-\d{2}$/.test(value) ? `${value}T00:00` : value;
  const date = new Date(normalizedValue);
  return Number.isNaN(date.valueOf()) ? null : date;
}

function formatDateLabel(value) {
  const date = value instanceof Date ? value : parseDateValue(value);
  if (!date) {
    return "Not recorded";
  }

  return new Intl.DateTimeFormat(undefined, {
    dateStyle: "medium",
  }).format(date);
}

function incidentReportedAtDate(incident) {
  return parseDateValue(incident.reportedAt) || parseDateValue(incident.reportCompiledDate) || parseDateValue(incident.updatedAt);
}

function combineIncidentDateAndTime(incident, value) {
  if (!value) {
    return null;
  }

  const directDate = parseDateValue(value);
  if (directDate) {
    return directDate;
  }

  const baseDate = incidentReportedAtDate(incident);
  const timeMatch = String(value).match(/^(\d{1,2}):(\d{2})/);
  if (!baseDate || !timeMatch) {
    return null;
  }

  const combined = new Date(baseDate);
  combined.setHours(Number(timeMatch[1]), Number(timeMatch[2]), 0, 0);

  if (combined.valueOf() < baseDate.valueOf()) {
    combined.setDate(combined.getDate() + 1);
  }

  return combined;
}

function minutesBetween(start, end) {
  if (!start || !end) {
    return null;
  }

  const diff = end.valueOf() - start.valueOf();
  return diff >= 0 ? Math.round(diff / 60000) : null;
}

function averageNumbers(values) {
  const numbers = values.filter((value) => Number.isFinite(value));
  if (numbers.length === 0) {
    return null;
  }

  return numbers.reduce((total, value) => total + value, 0) / numbers.length;
}

function formatDurationMinutes(value) {
  if (!Number.isFinite(value)) {
    return "Not recorded";
  }

  const rounded = Math.round(value);
  if (rounded < 60) {
    return `${rounded} min`;
  }

  const hours = Math.floor(rounded / 60);
  const minutes = rounded % 60;
  return minutes > 0 ? `${hours}h ${minutes}m` : `${hours}h`;
}

function formatHoursValue(value) {
  if (!Number.isFinite(value)) {
    return "Not recorded";
  }

  const precision = value >= 10 || Number.isInteger(value) ? 0 : 1;
  return `${value.toFixed(precision).replace(/\.0$/, "")} h`;
}

function csvCell(value) {
  const text = String(value ?? "");
  return /[",\n]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
}

function tallyBy(items, getLabel) {
  const counts = new Map();

  items.forEach((item) => {
    const label = singleLine(getLabel(item)) || "Not recorded";
    counts.set(label, (counts.get(label) || 0) + 1);
  });

  return [...counts.entries()]
    .map(([label, count]) => ({ label, count }))
    .sort((a, b) => b.count - a.count || a.label.localeCompare(b.label));
}

function radioTimestampValue(value) {
  const parsed = new Date(value).valueOf();
  return Number.isNaN(parsed) ? 0 : parsed;
}

function normalizeRadioLogs(logs) {
  if (!Array.isArray(logs)) {
    return [];
  }

  return logs
    .filter((entry) => entry && typeof entry === "object")
    .map((entry) => ({
      id: entry.id || createRadioLogId(),
      timestamp: entry.timestamp || "",
      unit: entry.unit || "",
      movement: entry.movement || "Other",
      channel: entry.channel || "",
      location: entry.location || "",
      notes: entry.notes || "",
    }))
    .sort((a, b) => radioTimestampValue(a.timestamp) - radioTimestampValue(b.timestamp));
}

function normalizeIncidentPhotos(photos) {
  if (!Array.isArray(photos)) {
    return [];
  }

  return photos
    .filter(
      (photo) =>
        photo &&
        typeof photo === "object" &&
        (typeof photo.dataUrl === "string" || typeof photo.url === "string" || typeof photo.storagePath === "string")
    )
    .map((photo) => ({
      id: photo.id || createIncidentPhotoId(),
      name: singleLine(photo.name) || "Incident photo",
      type: photo.type || "image/jpeg",
      dataUrl: typeof photo.dataUrl === "string" ? photo.dataUrl : "",
      url: typeof photo.url === "string" ? photo.url : "",
      storagePath: typeof photo.storagePath === "string" ? photo.storagePath : "",
    }));
}

function incidentPhotoSource(photo) {
  return photo.dataUrl || photo.url || "";
}

function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(new Error(`Unable to read ${file.name}.`));
    reader.readAsDataURL(file);
  });
}

function loadImageElement(source) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("Image could not be loaded."));
    image.src = source;
  });
}

async function prepareIncidentPhoto(file) {
  const originalDataUrl = await readFileAsDataUrl(file);

  try {
    const image = await loadImageElement(originalDataUrl);
    const maxDimension = 1600;
    const scale = Math.min(1, maxDimension / Math.max(image.naturalWidth, image.naturalHeight));
    const width = Math.max(1, Math.round(image.naturalWidth * scale));
    const height = Math.max(1, Math.round(image.naturalHeight * scale));
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const context = canvas.getContext("2d");

    if (!context) {
      return {
        id: createIncidentPhotoId(),
        name: file.name,
        type: file.type || "image/jpeg",
        dataUrl: originalDataUrl,
      };
    }

    context.drawImage(image, 0, 0, width, height);
    const outputType = file.type === "image/png" ? "image/png" : "image/jpeg";
    const dataUrl = outputType === "image/png" ? canvas.toDataURL(outputType) : canvas.toDataURL(outputType, 0.82);

    return {
      id: createIncidentPhotoId(),
      name: file.name,
      type: outputType,
      dataUrl,
    };
  } catch {
    return {
      id: createIncidentPhotoId(),
      name: file.name,
      type: file.type || "image/jpeg",
      dataUrl: originalDataUrl,
    };
  }
}

function inferTemplateIncidentCategory(incident) {
  const type = singleLine(incident.incidentType).toLowerCase();

  if (incident.medicalResponseRequired || type.includes("medical")) {
    return "Medical";
  }

  if (type.includes("security")) {
    return "Security";
  }

  return "Emergency Response";
}

function defaultIncident() {
  return {
    incidentId: createIncidentId(),
    status: "Open",
    incidentType: "Ground fall",
    severity: "Moderate",
    templateLocationType: "Mine",
    siteRegisterNo: "",
    esoTeam: "",
    esoSupervisor: "",
    mineSite: "",
    shaftLevel: "",
    locationDetail: "",
    shift: "Day",
    reportedAt: toLocalDateTimeValue(),
    reportCompiledBy: "",
    reportCompiledDate: toLocalDateValue(),
    templateIncidentCategory: "Emergency Response",
    controlledAt: "",
    incidentController: "",
    leadResponder: "",
    esoPrimaryRole: "",
    personnelCount: 0,
    casualtyCount: 0,
    evacuatedCount: 0,
    downtimeHours: 0,
    activationMethodRadio: false,
    activationMethodSmartGraphics: false,
    activationMethodInPerson: false,
    activationMethodPhone: false,
    timeEsoSuperintendentNotified: "",
    timeSseAltNotified: "",
    timeWhispirOtherNotified: "",
    arrivalAtSiteTime: "",
    departureTime: "",
    medicalResponseRequired: false,
    regulatorNotified: false,
    ventilationAffected: false,
    reentryRestricted: false,
    hazards: "",
    environmentalConditions: "",
    casualtySummary: "",
    equipmentDeployed: "",
    initialAlert: "",
    initialActions: "",
    responseActions: "",
    communicationsLog: "",
    investigationNotes: "",
    recoveryActions: "",
    involvedPerson1Name: "",
    involvedPerson1Company: "",
    involvedPerson2Name: "",
    involvedPerson2Company: "",
    involvedPerson3Name: "",
    involvedPerson3Company: "",
    involvedPerson4Name: "",
    involvedPerson4Company: "",
    externalAgenciesInvolved: "",
    handoverAgency: "",
    logisticsIssues: "",
    communicationIssues: "",
    policyProcedureIssues: "",
    mutualAidIssues: "",
    otherIssuesComments: "",
    incidentPhotoReferences: "",
    incidentPhotos: [],
    assignedPersonnelIds: [],
    radioLogs: [],
    updatedAt: new Date().toISOString(),
  };
}

function normalizeIncident(raw = {}) {
  const incident = {
    ...defaultIncident(),
    ...raw,
  };
  const hadAssignedPersonnel = Array.isArray(raw.assignedPersonnelIds) && raw.assignedPersonnelIds.length > 0;

  incident.personnelCount = normalizeNumber(incident.personnelCount);
  incident.casualtyCount = normalizeNumber(incident.casualtyCount);
  incident.evacuatedCount = normalizeNumber(incident.evacuatedCount);
  incident.downtimeHours = normalizeNumber(incident.downtimeHours);
  incident.assignedPersonnelIds = Array.isArray(raw.assignedPersonnelIds)
    ? [...new Set(raw.assignedPersonnelIds.map(String))]
    : [];
  incident.radioLogs = normalizeRadioLogs(raw.radioLogs);
  incident.incidentPhotos = normalizeIncidentPhotos(raw.incidentPhotos);
  incident.medicalResponseRequired = normalizeBoolean(incident.medicalResponseRequired);
  incident.regulatorNotified = normalizeBoolean(incident.regulatorNotified);
  incident.ventilationAffected = normalizeBoolean(incident.ventilationAffected);
  incident.reentryRestricted = normalizeBoolean(incident.reentryRestricted);
  incident.activationMethodRadio = normalizeBoolean(raw.activationMethodRadio, incident.radioLogs.length > 0);
  incident.activationMethodSmartGraphics = normalizeBoolean(raw.activationMethodSmartGraphics);
  incident.activationMethodInPerson = normalizeBoolean(raw.activationMethodInPerson);
  incident.activationMethodPhone = normalizeBoolean(raw.activationMethodPhone);
  incident.templateLocationType = incident.templateLocationType || "Mine";
  incident.templateIncidentCategory = raw.templateIncidentCategory || inferTemplateIncidentCategory(incident);
  incident.esoTeam = buildEsoTeamValue(incident.assignedPersonnelIds, hadAssignedPersonnel ? "" : raw.esoTeam ?? incident.esoTeam);
  incident.reportCompiledDate =
    raw.reportCompiledDate || (incident.reportedAt ? toLocalDateValue(new Date(incident.reportedAt)) : toLocalDateValue());
  incident.updatedAt = raw.updatedAt || new Date().toISOString();

  return incident;
}

function defaultWorker() {
  return {
    employeeId: "",
    fullName: "",
    role: "",
    crew: "",
    phone: "",
    certifications: "",
    status: "Available",
  };
}

function normalizeWorker(raw = {}) {
  return {
    ...defaultWorker(),
    ...raw,
    id: raw.id || createWorkerId(),
  };
}

function defaultSettings() {
  return {
    radioUnits: [...DEFAULT_RADIO_UNITS],
  };
}

function normalizeSettings(raw = {}) {
  const sourceUnits = Array.isArray(raw.radioUnits) ? raw.radioUnits : defaultSettings().radioUnits;

  return {
    radioUnits: [...new Set(sourceUnits.map(singleLine).filter(Boolean))],
  };
}

function buildEsoTeamValue(workerIds, fallback = "") {
  const selectedIds = Array.isArray(workerIds) ? workerIds : [];
  const workerNames = [...new Set(
    selectedIds
      .map((workerId) => state.personnel.find((worker) => worker.id === workerId))
      .filter(Boolean)
      .map((worker) => singleLine(worker.fullName))
      .filter(Boolean)
  )];

  return workerNames.length > 0 ? workerNames.join(", ") : singleLine(fallback);
}

function syncEsoTeamField(fallback = "") {
  const field = dom.incidentForm.elements.namedItem("esoTeam");
  if (!(field instanceof HTMLInputElement)) {
    return;
  }

  field.value = buildEsoTeamValue(state.selectedPersonnelIds, fallback);
}

function startOfDay(date) {
  const copy = new Date(date);
  copy.setHours(0, 0, 0, 0);
  return copy;
}

function endOfDay(date) {
  const copy = new Date(date);
  copy.setHours(23, 59, 59, 999);
  return copy;
}

function shiftDays(date, amount) {
  const copy = new Date(date);
  copy.setDate(copy.getDate() + amount);
  return copy;
}

function resolveReportingPresetRange(preset, referenceDate = new Date()) {
  const today = startOfDay(referenceDate);

  switch (preset) {
    case "last7":
      return {
        dateFrom: toLocalDateValue(shiftDays(today, -6)),
        dateTo: toLocalDateValue(today),
      };
    case "last90":
      return {
        dateFrom: toLocalDateValue(shiftDays(today, -89)),
        dateTo: toLocalDateValue(today),
      };
    case "ytd":
      return {
        dateFrom: toLocalDateValue(new Date(today.getFullYear(), 0, 1)),
        dateTo: toLocalDateValue(today),
      };
    case "all":
      return {
        dateFrom: "",
        dateTo: "",
      };
    case "custom":
      return {
        dateFrom: "",
        dateTo: "",
      };
    case "last30":
    default:
      return {
        dateFrom: toLocalDateValue(shiftDays(today, -29)),
        dateTo: toLocalDateValue(today),
      };
  }
}

function defaultReportingFilters() {
  return {
    preset: "last30",
    ...resolveReportingPresetRange("last30"),
    location: "",
    status: "",
    severity: "",
    incidentType: "",
  };
}

function normalizeReportingFilters(raw = {}) {
  const filters = {
    ...defaultReportingFilters(),
    ...raw,
  };

  filters.preset = REPORTING_PRESETS.has(filters.preset) ? filters.preset : "last30";
  filters.location = singleLine(filters.location);
  filters.status = singleLine(filters.status);
  filters.severity = singleLine(filters.severity);
  filters.incidentType = singleLine(filters.incidentType);
  filters.dateFrom = singleLine(filters.dateFrom);
  filters.dateTo = singleLine(filters.dateTo);

  if (filters.preset !== "custom") {
    const presetRange = resolveReportingPresetRange(filters.preset);
    filters.dateFrom = presetRange.dateFrom;
    filters.dateTo = presetRange.dateTo;
  }

  if (filters.preset === "custom" && filters.dateFrom && filters.dateTo && filters.dateFrom > filters.dateTo) {
    [filters.dateFrom, filters.dateTo] = [filters.dateTo, filters.dateFrom];
  }

  return filters;
}

function safeParseStorage(key) {
  try {
    const stored = localStorage.getItem(key);
    return stored ? JSON.parse(stored) : [];
  } catch {
    return [];
  }
}

function safeParseObject(key, fallback) {
  try {
    const stored = localStorage.getItem(key);
    if (!stored) {
      return fallback;
    }

    const parsed = JSON.parse(stored);
    return parsed && typeof parsed === "object" ? parsed : fallback;
  } catch {
    return fallback;
  }
}

function canUseCloudStorage() {
  return window.location.protocol !== "file:";
}

function setStorageMode(mode, detail = "") {
  state.storageMode = mode;

  if (!dom.storageModeIndicator) {
    return;
  }

  switch (mode) {
    case "cloud":
      dom.storageModeIndicator.textContent = detail || "Google Cloud storage connected";
      break;
    case "local-fallback":
      dom.storageModeIndicator.textContent = detail || "Local fallback active: cloud unavailable";
      break;
    default:
      dom.storageModeIndicator.textContent = detail || "Local browser storage active";
      break;
  }
}

function cacheIncidentsLocally(notifyOnFailure = true) {
  try {
    localStorage.setItem(INCIDENT_STORAGE_KEY, JSON.stringify(state.incidents));
    return true;
  } catch {
    if (notifyOnFailure) {
      window.alert("Unable to save the incident in browser storage. Attached images may be too large for local storage.");
    }
    return false;
  }
}

function cachePersonnelLocally() {
  try {
    localStorage.setItem(PERSONNEL_STORAGE_KEY, JSON.stringify(state.personnel));
    return true;
  } catch {
    return false;
  }
}

function cacheSettingsLocally() {
  try {
    localStorage.setItem(SETTINGS_STORAGE_KEY, JSON.stringify(state.settings));
    return true;
  } catch {
    return false;
  }
}

function cacheAllDataLocally() {
  cacheSettingsLocally();
  cachePersonnelLocally();
  return cacheIncidentsLocally(false);
}

function loadPersonnelFromLocal() {
  state.personnel = safeParseStorage(PERSONNEL_STORAGE_KEY).map(normalizeWorker);
}

function loadSettingsFromLocal() {
  state.settings = normalizeSettings(safeParseObject(SETTINGS_STORAGE_KEY, defaultSettings()));
}

function loadIncidentsFromLocal() {
  state.incidents = safeParseStorage(INCIDENT_STORAGE_KEY).map(normalizeIncident);
}

async function apiRequest(path, options = {}) {
  const response = await fetch(`${API_BASE}${path}`, {
    headers: {
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
    ...options,
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(errorText || `${response.status} ${response.statusText}`);
  }

  if (response.status === 204) {
    return null;
  }

  return response.json();
}

async function apiBootstrap() {
  return apiRequest("/bootstrap");
}

async function apiSaveSettings(settings) {
  const response = await apiRequest("/settings", {
    method: "PUT",
    body: JSON.stringify({ settings }),
  });
  return normalizeSettings(response.settings);
}

async function apiSavePersonnel(personnel) {
  const response = await apiRequest("/personnel", {
    method: "PUT",
    body: JSON.stringify({ personnel }),
  });
  return Array.isArray(response.personnel) ? response.personnel.map(normalizeWorker) : [];
}

async function apiSaveIncident(incident) {
  const response = await apiRequest(`/incidents/${encodeURIComponent(incident.incidentId)}`, {
    method: "PUT",
    body: JSON.stringify({ incident }),
  });
  return normalizeIncident(response.incident);
}

async function apiSaveIncidentsBulk(incidents) {
  const response = await apiRequest("/incidents/bulk", {
    method: "POST",
    body: JSON.stringify({ incidents }),
  });
  return Array.isArray(response.incidents) ? response.incidents.map(normalizeIncident) : [];
}

async function apiDeleteIncident(incidentId) {
  await apiRequest(`/incidents/${encodeURIComponent(incidentId)}`, {
    method: "DELETE",
  });
}

function applyBootstrapData(data = {}) {
  state.settings = normalizeSettings(data.settings || defaultSettings());
  state.personnel = Array.isArray(data.personnel) ? data.personnel.map(normalizeWorker) : [];
  state.incidents = Array.isArray(data.incidents) ? data.incidents.map(normalizeIncident) : [];
}

async function loadAppData() {
  if (canUseCloudStorage()) {
    try {
      const bootstrap = await apiBootstrap();
      applyBootstrapData(bootstrap);
      cacheAllDataLocally();
      setStorageMode("cloud", "Google Cloud storage connected");
      return;
    } catch (error) {
      console.error("Cloud bootstrap failed, using local cache instead.", error);
      loadSettingsFromLocal();
      loadPersonnelFromLocal();
      loadIncidentsFromLocal();
      setStorageMode("local-fallback", "Local fallback active: cloud unavailable");
      return;
    }
  }

  loadSettingsFromLocal();
  loadPersonnelFromLocal();
  loadIncidentsFromLocal();
  setStorageMode("local", "Local browser storage active");
}

async function persistSettings() {
  const localCached = cacheSettingsLocally();

  if (!canUseCloudStorage() || state.storageMode === "local") {
    return localCached;
  }

  try {
    state.settings = await apiSaveSettings(state.settings);
    cacheSettingsLocally();
    setStorageMode("cloud", "Google Cloud storage connected");
    return true;
  } catch (error) {
    console.error("Unable to sync settings to cloud storage.", error);
    setStorageMode("local-fallback", "Local fallback active: cloud sync failed");
    return localCached;
  }
}

async function persistPersonnel() {
  const localCached = cachePersonnelLocally();

  if (!canUseCloudStorage() || state.storageMode === "local") {
    return localCached;
  }

  try {
    state.personnel = await apiSavePersonnel(state.personnel);
    cachePersonnelLocally();
    setStorageMode("cloud", "Google Cloud storage connected");
    return true;
  } catch (error) {
    console.error("Unable to sync personnel to cloud storage.", error);
    setStorageMode("local-fallback", "Local fallback active: cloud sync failed");
    return localCached;
  }
}

async function persistIncidents(options = {}) {
  const localCached = cacheIncidentsLocally(state.storageMode === "local");

  if (!canUseCloudStorage() || state.storageMode === "local") {
    return localCached;
  }

  try {
    if (options.mode === "single" && options.incident) {
      const savedIncident = await apiSaveIncident(options.incident);
      const index = state.incidents.findIndex((incident) => incident.incidentId === savedIncident.incidentId);

      if (index >= 0) {
        state.incidents[index] = savedIncident;
      } else {
        state.incidents.push(savedIncident);
      }
    } else {
      state.incidents = await apiSaveIncidentsBulk(state.incidents);
    }

    cacheIncidentsLocally(false);
    setStorageMode("cloud", "Google Cloud storage connected");
    return true;
  } catch (error) {
    console.error("Unable to sync incidents to cloud storage.", error);
    setStorageMode("local-fallback", "Local fallback active: cloud sync failed");
    return localCached;
  }
}

async function seedSettingsIfEmpty() {
  if (state.settings.radioUnits.length > 0) {
    return;
  }

  state.settings = defaultSettings();
  await persistSettings();
}

async function seedPersonnelIfEmpty() {
  if (state.personnel.length > 0) {
    return;
  }

  state.personnel = [
    normalizeWorker({
      id: "worker-sample-1",
      employeeId: "RR-014",
      fullName: "Kara Morgan",
      role: "Emergency Response Team Lead",
      crew: "ERT Alpha",
      phone: "ERT-1",
      certifications: "Underground rescue, medic, breathing apparatus",
      status: "On shift",
    }),
    normalizeWorker({
      id: "worker-sample-2",
      employeeId: "RR-022",
      fullName: "Ari Fletcher",
      role: "Shift Superintendent",
      crew: "Supervision",
      phone: "SUP-2",
      certifications: "Incident control, ventilation management",
      status: "Available",
    }),
    normalizeWorker({
      id: "worker-sample-3",
      employeeId: "RR-105",
      fullName: "Jules Bennett",
      role: "Underground Medic",
      crew: "Medical",
      phone: "MED-4",
      certifications: "Trauma care, extraction support",
      status: "On shift",
    }),
    normalizeWorker({
      id: "worker-sample-4",
      employeeId: "RR-211",
      fullName: "Tama Reid",
      role: "Traffic Controller",
      crew: "Operations",
      phone: "OPS-7",
      certifications: "Traffic isolation, vehicle control",
      status: "Underground",
    }),
  ];

  await persistPersonnel();
}

async function seedIncidentsIfEmpty() {
  if (state.incidents.length > 0) {
    return;
  }

  const sample = normalizeIncident({
    incidentId: "INC-2026-03-10-0715",
    status: "Contained",
    incidentType: "Vehicle collision",
    severity: "High",
    templateLocationType: "Mine",
    siteRegisterNo: "RRM-ESO-240310-01",
    esoTeam: "Red Ridge ERT Alpha",
    esoSupervisor: "Ari Fletcher",
    mineSite: "Red Ridge Mine",
    shaftLevel: "Ramp 640",
    locationDetail: "North decline intersection",
    shift: "Day",
    reportedAt: "2026-03-10T07:15",
    reportCompiledBy: "Ari Fletcher",
    reportCompiledDate: "2026-03-10",
    templateIncidentCategory: "Emergency Response",
    controlledAt: "2026-03-10T08:10",
    incidentController: "Ari Fletcher",
    leadResponder: "Kara Morgan",
    esoPrimaryRole: "Traffic isolation, casualty extraction, and scene coordination.",
    personnelCount: 5,
    casualtyCount: 1,
    evacuatedCount: 18,
    downtimeHours: 4.5,
    activationMethodRadio: true,
    activationMethodPhone: true,
    timeEsoSuperintendentNotified: "07:18",
    timeSseAltNotified: "07:23",
    timeWhispirOtherNotified: "07:27",
    arrivalAtSiteTime: "07:29",
    departureTime: "08:18",
    medicalResponseRequired: true,
    regulatorNotified: false,
    ventilationAffected: false,
    reentryRestricted: true,
    hazards: "Heavy vehicle congestion, low visibility from dust, damaged barricade.",
    environmentalConditions: "Ventilation stable. Visibility reduced at incident point until water cart arrived.",
    casualtySummary: "One operator with suspected shoulder fracture; treated underground and transferred topside.",
    equipmentDeployed: "Rescue vehicle, stretcher, trauma kit, gas monitor, barricade replacements.",
    initialAlert: "Control room notified by loader operator after light vehicle was struck at intersection.",
    initialActions: "Stopped traffic on decline, isolated zone, dispatched emergency response team and medic.",
    responseActions: "Assessed casualty, stabilized scene, extracted patient, reviewed visibility controls and traffic management.",
    communicationsLog: "Updates sent to shift superintendent, control room, and site medic every 15 minutes.",
    investigationNotes: "Preliminary review indicates failed right-of-way compliance and ineffective dust suppression.",
    recoveryActions: "Intersection remains barricaded pending supervisor review and signage replacement.",
    involvedPerson1Name: "M. Ellis",
    involvedPerson1Company: "Hancock Iron Ore",
    involvedPerson2Name: "P. Singh",
    involvedPerson2Company: "Hancock Iron Ore",
    externalAgenciesInvolved: "Site medic and surface ambulance crew notified for topside transfer.",
    handoverAgency: "Casualty handed to ambulance crew at portal at 08:04.",
    logisticsIssues: "Traffic congestion delayed secondary response vehicle until decline was cleared.",
    communicationIssues: "Heavy equipment traffic required repeated channel clearance calls.",
    policyProcedureIssues: "Dust suppression controls and intersection right-of-way compliance under review.",
    mutualAidIssues: "No mutual aid activation required.",
    otherIssuesComments: "Intersection remains under supervisor control pending signage replacement.",
    incidentPhotoReferences: "No incident photos attached. CCTV and supervisor scene photos requested.",
    assignedPersonnelIds: state.personnel.slice(0, 3).map((worker) => worker.id),
    radioLogs: [
      {
        id: "radio-sample-1",
        timestamp: "2026-03-10T07:18",
        unit: "ERT Buggy 1",
        movement: "Dispatched",
        channel: "Emergency Mine",
        location: "Surface emergency bay to portal",
        notes: "First response crew deployed with medic and trauma kit.",
      },
      {
        id: "radio-sample-2",
        timestamp: "2026-03-10T07:29",
        unit: "Ambulance 2",
        movement: "Entered mine",
        channel: "Emergency Mine",
        location: "Portal descending to ramp 640",
        notes: "Escorted by traffic controller following decline shutdown.",
      },
      {
        id: "radio-sample-3",
        timestamp: "2026-03-10T07:42",
        unit: "Rescue 1",
        movement: "Arrived on scene",
        channel: "Emergency Mine",
        location: "North decline intersection",
        notes: "Scene secure and casualty handover confirmed.",
      },
    ],
  });

  state.incidents.push(sample);
  await persistIncidents();
}

async function syncIncidentPersonnelAssignments() {
  const validIds = new Set(state.personnel.map((worker) => worker.id));
  state.incidents = state.incidents.map((incident) =>
    normalizeIncident({
      ...incident,
      assignedPersonnelIds: incident.assignedPersonnelIds.filter((workerId) => validIds.has(workerId)),
    })
  );
  state.selectedPersonnelIds = state.selectedPersonnelIds.filter((workerId) => validIds.has(workerId));
  return persistIncidents();
}

function getWorkerById(workerId) {
  return state.personnel.find((worker) => worker.id === workerId) || null;
}

function getAssignedWorkers(workerIds) {
  return workerIds.map(getWorkerById).filter(Boolean);
}

function sortIncidents(incidents) {
  return [...incidents].sort((a, b) => new Date(b.updatedAt).valueOf() - new Date(a.updatedAt).valueOf());
}

function getSelectedIncident() {
  return state.incidents.find((incident) => incident.incidentId === state.selectedIncidentId) || null;
}

function getEditingIncident() {
  return state.incidents.find((incident) => incident.incidentId === state.editingIncidentId) || null;
}

function statusTone(status) {
  switch (status) {
    case "Closed":
      return "var(--good)";
    case "Contained":
      return "var(--accent)";
    case "Under Investigation":
      return "var(--alert)";
    default:
      return "rgba(255, 255, 255, 0.1)";
  }
}

function statusTextColor(status) {
  switch (status) {
    case "Contained":
      return "#241503";
    case "Closed":
      return "#06120d";
    default:
      return "";
  }
}

function workerStatusTone(status) {
  switch (status) {
    case "Available":
      return { background: "rgba(56, 194, 134, 0.18)", color: "#c9ffe6" };
    case "On shift":
      return { background: "rgba(255, 159, 28, 0.18)", color: "#ffe0ab" };
    case "Underground":
      return { background: "rgba(95, 156, 255, 0.18)", color: "#d5e4ff" };
    default:
      return { background: "rgba(255, 95, 69, 0.18)", color: "#ffd5ce" };
  }
}

function currentReportedAtValue() {
  const field = dom.incidentForm.elements.namedItem("reportedAt");
  return field && field.value ? field.value : toLocalDateTimeValue();
}

function getActivationMethods(incident) {
  const methods = [];

  if (incident.activationMethodRadio) {
    methods.push("Radio");
  }

  if (incident.activationMethodSmartGraphics) {
    methods.push("Smart Graphics");
  }

  if (incident.activationMethodInPerson) {
    methods.push("In Person");
  }

  if (incident.activationMethodPhone) {
    methods.push("Phone");
  }

  return methods;
}

function getInvolvedPeople(incident) {
  const people = [];

  for (let index = 1; index <= 4; index += 1) {
    const name = singleLine(incident[`involvedPerson${index}Name`]);
    const company = singleLine(incident[`involvedPerson${index}Company`]);

    if (!name && !company) {
      continue;
    }

    people.push({
      name: name || "Not recorded",
      company: company || "Not recorded",
    });
  }

  return people.slice(0, 4);
}

function truncateBlockText(value, maxLength) {
  const text = String(value || "").trim();
  if (!text || text.length <= maxLength) {
    return text;
  }

  return `${text.slice(0, Math.max(0, maxLength - 3)).trimEnd()}...`;
}

const HIO_DETAIL_MEASURE_WIDTH_MM = 138;
const HIO_DETAIL_FIRST_PAGE_HEIGHT_MM = 102;
const HIO_DETAIL_CONTINUATION_HEIGHT_MM = 220;

function createHioTextMeasure(heightMm) {
  const measure = document.createElement("div");
  measure.className = "hio-large-box";
  Object.assign(measure.style, {
    position: "fixed",
    left: "-10000px",
    top: "0",
    width: `${HIO_DETAIL_MEASURE_WIDTH_MM}mm`,
    minHeight: `${heightMm}mm`,
    height: `${heightMm}mm`,
    overflow: "hidden",
    visibility: "hidden",
    pointerEvents: "none",
    boxSizing: "border-box",
    zIndex: "-1",
  });
  document.body.appendChild(measure);
  return measure;
}

function takeHioTextChunk(lines, startIndex, heightMm) {
  const measure = createHioTextMeasure(heightMm);
  let endIndex = startIndex;
  let bestText = "";

  while (endIndex < lines.length) {
    const candidate = lines.slice(startIndex, endIndex + 1).join("\n");
    measure.textContent = candidate || "\u00a0";

    if (measure.scrollHeight > measure.clientHeight) {
      if (endIndex === startIndex) {
        bestText = candidate;
        endIndex += 1;
      }
      break;
    }

    bestText = candidate;
    endIndex += 1;
  }

  measure.remove();

  return {
    text: bestText.trimEnd(),
    nextIndex: endIndex,
  };
}

function paginateHioDetailBlock(detailText) {
  const normalizedText = String(detailText || "").trim() || "No incident / activation detail recorded.";
  const lines = normalizedText.split("\n");
  const pages = [];
  let cursor = 0;

  while (cursor < lines.length) {
    const heightMm = pages.length === 0 ? HIO_DETAIL_FIRST_PAGE_HEIGHT_MM : HIO_DETAIL_CONTINUATION_HEIGHT_MM;
    const chunk = takeHioTextChunk(lines, cursor, heightMm);

    if (!chunk.text && chunk.nextIndex <= cursor) {
      break;
    }

    pages.push(chunk.text || " ");
    cursor = chunk.nextIndex;
  }

  return pages.length > 0 ? pages : [normalizedText];
}

function buildHioChronologyEntries(incident) {
  const reportedAt = incidentReportedAtDate(incident);
  const entries = [];

  const addEntry = (dateValue, label, detail = "") => {
    const date = dateValue instanceof Date ? dateValue : parseDateValue(dateValue);
    if (!date) {
      return;
    }

    const parts = [singleLine(label)];
    const normalizedDetail = singleLine(detail);

    if (normalizedDetail) {
      parts.push(normalizedDetail);
    }

    entries.push({
      sortValue: date.valueOf(),
      text: `${formatTimeOnly(date)} ${parts.join(": ")}`,
    });
  };

  addEntry(reportedAt, "Initial alert", incident.initialAlert || "Control room notification received");
  addEntry(reportedAt, "Immediate actions", incident.initialActions);
  addEntry(combineIncidentDateAndTime(incident, incident.timeEsoSuperintendentNotified), "ESO superintendent notified");
  addEntry(combineIncidentDateAndTime(incident, incident.timeSseAltNotified), "SSE or ALT SSE notified");
  addEntry(combineIncidentDateAndTime(incident, incident.timeWhispirOtherNotified), "Whispir / other notified");
  addEntry(combineIncidentDateAndTime(incident, incident.arrivalAtSiteTime), "Arrived at site");
  addEntry(parseDateValue(incident.controlledAt), "Incident controlled");
  addEntry(combineIncidentDateAndTime(incident, incident.departureTime), "Departed site");

  incident.radioLogs.forEach((entry) => {
    const radioDetail = [
      singleLine(entry.unit) || "Unit",
      singleLine(entry.movement),
      singleLine(entry.location),
      singleLine(entry.notes),
    ]
      .filter(Boolean)
      .join(" | ");

    addEntry(entry.timestamp, "Radio", radioDetail);
  });

  return entries.sort((a, b) => a.sortValue - b.sortValue).map((entry) => entry.text);
}

function buildHioDetailBlock(incident) {
  const chronology = buildHioChronologyEntries(incident);
  const sections = [];

  if (singleLine(incident.equipmentDeployed)) {
    sections.push(`Assets deployed: ${singleLine(incident.equipmentDeployed)}`);
  }

  if (chronology.length > 0) {
    sections.push(["Chronology:", ...chronology].join("\n"));
  }

  [
    ["Ongoing response", incident.responseActions],
    ["Casualty summary", incident.casualtySummary],
    ["Primary hazards", incident.hazards],
    ["Environmental conditions", incident.environmentalConditions],
    ["Communications", incident.communicationsLog],
    ["Recovery", incident.recoveryActions],
  ]
    .filter(([, value]) => singleLine(value))
    .forEach(([label, value]) => {
      sections.push(`${label}: ${singleLine(value)}`);
    });

  return sections.join("\n\n") || "No incident / activation detail recorded.";
}

function buildHioPhotoBlock(incident) {
  const photoSection = singleLine(incident.incidentPhotoReferences)
    ? `Photo references: ${singleLine(incident.incidentPhotoReferences)}`
    : "Photo references: None recorded.";
  const uploadedPhotosSection =
    incident.incidentPhotos.length > 0
      ? `Uploaded photos: ${incident.incidentPhotos.map((photo) => photo.name).join(", ")}`
      : "Uploaded photos: None.";

  const assignedWorkers = getAssignedWorkers(incident.assignedPersonnelIds);
  const personnelSection =
    assignedWorkers.length > 0
      ? `Assigned workers: ${assignedWorkers.map((worker) => worker.fullName).join(", ")}`
      : "Assigned workers: None recorded.";

  return truncateBlockText([photoSection, uploadedPhotosSection, "", personnelSection].join("\n"), 900);
}

function getCompiledByValue(incident) {
  return singleLine(incident.reportCompiledBy || incident.esoSupervisor || incident.incidentController);
}

function getCompiledDateValue(incident) {
  return incident.reportCompiledDate || (incident.reportedAt ? incident.reportedAt.slice(0, 10) : toLocalDateValue());
}

function renderHioCheck(label, selected) {
  return `
    <span class="hio-check-item">
      <span class="hio-check-box">${selected ? "X" : ""}</span>
      <span>${escapeHtml(label)}</span>
    </span>
  `;
}

function renderHioLineValue(value) {
  const text = singleLine(value);
  return text ? escapeHtml(text) : "&nbsp;";
}

function renderHioBlockValue(value) {
  const text = String(value || "").trim();
  return text ? escapeHtml(text) : "&nbsp;";
}

function renderHioPageMeta(pageNumber, totalPages) {
  return `
    <footer class="hio-page-meta">
      <span><strong>Rev</strong>3</span>
      <span><strong>Document #</strong>QP-TEM-00270</span>
      <span><strong>Author</strong>D. Newspoint</span>
      <span><strong>Approver</strong>Drew Duddy</span>
      <span><strong>Issue Date</strong>13/11/2023</span>
      <span><strong>Page</strong>${pageNumber} of ${totalPages}</span>
    </footer>
  `;
}

function renderHioPhotoGallery(photos) {
  const galleryPhotos = photos.slice(0, 4);

  if (galleryPhotos.length === 0) {
    return '<div class="hio-photo-empty">No uploaded images attached.</div>';
  }

  return `
    <div class="hio-photo-grid">
      ${galleryPhotos
        .map(
          (photo) => `
            <figure class="hio-photo-tile">
              <img src="${escapeHtml(incidentPhotoSource(photo))}" alt="${escapeHtml(photo.name)}" />
              <figcaption>${escapeHtml(photo.name)}</figcaption>
            </figure>
          `
        )
        .join("")}
    </div>
  `;
}

function renderHioFirstPage(incident, involvedPeople, detailBlock, pageNumber, totalPages) {
  const category = incident.templateIncidentCategory || inferTemplateIncidentCategory(incident);

  return `
    <section class="hio-page">
      <article class="hio-export-sheet">
        <header class="hio-export-header">
          <h2>Emergency &amp; Security Activation Report</h2>
          <img class="hio-export-logo" src="logo-hancock-iron-ore.svg" alt="Hancock Iron Ore" />
        </header>
        <div class="hio-export-rule"></div>

        <table class="hio-export-table">
          <tr>
            <th colspan="5" class="hio-section-bar">Initial Notification Details</th>
          </tr>
          <tr>
            <th class="hio-label-cell">Location:</th>
            <td colspan="4">
              <div class="hio-checkbox-row">
                ${renderHioCheck("Port", incident.templateLocationType === "Port")}
                ${renderHioCheck("Rail", incident.templateLocationType === "Rail")}
                ${renderHioCheck("Mine", incident.templateLocationType === "Mine")}
                ${renderHioCheck("Other", incident.templateLocationType === "Other")}
              </div>
            </td>
          </tr>
          <tr>
            <th class="hio-label-cell">Exact Location:</th>
            <td colspan="4">${renderHioLineValue(incident.locationDetail)}</td>
          </tr>
          <tr>
            <th class="hio-label-cell">ESO Team:</th>
            <td colspan="4">${renderHioLineValue(incident.esoTeam)}</td>
          </tr>
          <tr>
            <th class="hio-label-cell">Date:</th>
            <td colspan="4">${renderHioLineValue(formatDateOnly(incident.reportedAt))}</td>
          </tr>
          <tr>
            <th class="hio-label-cell">ESO Supervisor:</th>
            <td colspan="4">${renderHioLineValue(incident.esoSupervisor)}</td>
          </tr>
          <tr>
            <th class="hio-label-cell">Total no. of Personnel:</th>
            <td colspan="4">${renderHioLineValue(incident.personnelCount)}</td>
          </tr>
          <tr>
            <th class="hio-label-cell">Incident Type:</th>
            <td colspan="4">
              <div class="hio-checkbox-row">
                ${renderHioCheck("Security", category === "Security")}
                ${renderHioCheck("Medical", category === "Medical")}
                ${renderHioCheck("Emergency Response", category === "Emergency Response")}
              </div>
            </td>
          </tr>
          <tr>
            <th class="hio-label-cell">Activation Method:</th>
            <td colspan="4">
              <div class="hio-checkbox-row">
                ${renderHioCheck("Radio", incident.activationMethodRadio)}
                ${renderHioCheck("Smart Graphics", incident.activationMethodSmartGraphics)}
                ${renderHioCheck("In Person", incident.activationMethodInPerson)}
                ${renderHioCheck("Phone", incident.activationMethodPhone)}
              </div>
            </td>
          </tr>
          <tr>
            <th class="hio-label-cell">Time notified:</th>
            <td colspan="4">${renderHioLineValue(formatTimeOnly(incident.reportedAt))}</td>
          </tr>
          <tr>
            <th class="hio-label-cell">Time Superintendent and Managers notified:</th>
            <td>
              <div class="hio-mini-label">ESO Supt.</div>
              <div class="hio-mini-value">${renderHioLineValue(incident.timeEsoSuperintendentNotified)}</div>
            </td>
            <td>
              <div class="hio-mini-label">SSE or ALT SSE</div>
              <div class="hio-mini-value">${renderHioLineValue(incident.timeSseAltNotified)}</div>
            </td>
            <td colspan="2">
              <div class="hio-mini-label">Whispir/Other</div>
              <div class="hio-mini-value">${renderHioLineValue(incident.timeWhispirOtherNotified)}</div>
            </td>
          </tr>
          <tr>
            <th class="hio-label-cell">Time of arrival at Site:</th>
            <td colspan="2">${renderHioLineValue(incident.arrivalAtSiteTime)}</td>
            <th class="hio-label-cell">Time of departure:</th>
            <td>${renderHioLineValue(incident.departureTime)}</td>
          </tr>
          <tr>
            <th colspan="5" class="hio-section-bar">Incident / Activation Details</th>
          </tr>
          <tr>
            <th class="hio-label-cell hio-top">Assets Deployed:</th>
            <td colspan="4">
              <div class="hio-large-box hio-large-box--page-one">${renderHioBlockValue(detailBlock)}</div>
            </td>
          </tr>
          <tr>
            <th class="hio-label-cell">ESO Primary Role:</th>
            <td colspan="4">${renderHioLineValue(incident.esoPrimaryRole)}</td>
          </tr>
          ${involvedPeople
            .map(
              (person) => `
                <tr>
                  <th class="hio-label-cell">Person involved name:</th>
                  <td colspan="2">${renderHioLineValue(person.name)}</td>
                  <th class="hio-label-cell">Company name:</th>
                  <td>${renderHioLineValue(person.company)}</td>
                </tr>
              `
            )
            .join("")}
          <tr>
            <th class="hio-label-cell">External Agencies Involved:</th>
            <td colspan="4">${renderHioLineValue(incident.externalAgenciesInvolved)}</td>
          </tr>
          <tr>
            <th class="hio-label-cell">Handover (to another agency):</th>
            <td colspan="4">${renderHioLineValue(incident.handoverAgency)}</td>
          </tr>
          <tr>
            <th class="hio-label-cell">Logistics Issues:</th>
            <td colspan="4">${renderHioLineValue(incident.logisticsIssues)}</td>
          </tr>
          <tr>
            <th class="hio-label-cell">Communication Issues:</th>
            <td colspan="4">${renderHioLineValue(incident.communicationIssues)}</td>
          </tr>
          <tr>
            <th class="hio-label-cell">Policy/Procedure:</th>
            <td colspan="4">${renderHioLineValue(incident.policyProcedureIssues)}</td>
          </tr>
        </table>

        ${renderHioPageMeta(pageNumber, totalPages)}
      </article>
    </section>
  `;
}

function renderHioContinuationPage(detailBlock, pageNumber, totalPages) {
  return `
    <section class="hio-page">
      <article class="hio-export-sheet">
        <header class="hio-export-header">
          <h2>Emergency &amp; Security Activation Report</h2>
          <img class="hio-export-logo" src="logo-hancock-iron-ore.svg" alt="Hancock Iron Ore" />
        </header>
        <div class="hio-export-rule"></div>

        <table class="hio-export-table">
          <tr>
            <th colspan="5" class="hio-section-bar">Incident / Activation Details (continued)</th>
          </tr>
          <tr>
            <th class="hio-label-cell hio-top">Assets Deployed:</th>
            <td colspan="4">
              <div class="hio-large-box hio-large-box--continuation">${renderHioBlockValue(detailBlock)}</div>
            </td>
          </tr>
        </table>

        ${renderHioPageMeta(pageNumber, totalPages)}
      </article>
    </section>
  `;
}

function renderHioPhotoPage(incident, pageNumber, totalPages) {
  return `
    <section class="hio-page">
      <article class="hio-export-sheet">
        <header class="hio-export-header">
          <h2>Emergency &amp; Security Activation Report</h2>
          <img class="hio-export-logo" src="logo-hancock-iron-ore.svg" alt="Hancock Iron Ore" />
        </header>
        <div class="hio-export-rule"></div>

        <table class="hio-export-table hio-export-table-page-two">
          <tr>
            <th class="hio-label-cell">Mutual Aid Issues:</th>
            <td>${renderHioLineValue(incident.mutualAidIssues)}</td>
          </tr>
          <tr>
            <th class="hio-label-cell">Other issues/comments:</th>
            <td>${renderHioLineValue(incident.otherIssuesComments)}</td>
          </tr>
          <tr>
            <th colspan="2" class="hio-section-bar">Incident Photos if Available</th>
          </tr>
          <tr>
            <td colspan="2">
              <div class="hio-photo-box">
                <div class="hio-photo-layout">
                  ${renderHioPhotoGallery(incident.incidentPhotos)}
                  <div class="hio-photo-summary">${renderHioBlockValue(buildHioPhotoBlock(incident))}</div>
                </div>
              </div>
            </td>
          </tr>
        </table>

        <table class="hio-export-table hio-export-signoff">
          <tr>
            <th class="hio-label-cell">Report Compiled By:</th>
            <td>${renderHioLineValue(getCompiledByValue(incident))}</td>
            <th class="hio-label-cell">Date:</th>
            <td>${renderHioLineValue(formatDateOnly(getCompiledDateValue(incident)))}</td>
          </tr>
        </table>

        ${renderHioPageMeta(pageNumber, totalPages)}
      </article>
    </section>
  `;
}

function incidentLocationText(incident) {
  return [incident.templateLocationType, incident.locationDetail].filter(Boolean).join(" · ") || "location pending";
}

function makeSummary(incident) {
  const casualtyText =
    incident.casualtyCount > 0
      ? `${incident.casualtyCount} ${incident.casualtyCount === 1 ? "casualty" : "casualties"} reported`
      : "no casualties reported";
  const evacuationText =
    incident.evacuatedCount > 0 ? `${incident.evacuatedCount} personnel evacuated` : "no evacuation recorded";
  const radioText =
    incident.radioLogs.length > 0
      ? `${incident.radioLogs.length} radio movement entries logged.`
      : "Radio movement log not yet populated.";
  const assignedText =
    incident.assignedPersonnelIds.length > 0
      ? `${incident.assignedPersonnelIds.length} workers assigned.`
      : "No workers assigned yet.";
  const photoText =
    incident.incidentPhotos.length > 0
      ? `${photoCountLabel(incident.incidentPhotos.length)} attached.`
      : "No incident photos attached.";
  const regulatorText = incident.regulatorNotified ? "Regulatory notification completed." : "Regulatory notification still pending.";

  return `${incident.incidentType} at ${incidentLocationText(incident)} was reported ${formatDateTime(
    incident.reportedAt
  )}. Severity is ${incident.severity.toLowerCase()} and status is ${incident.status.toLowerCase()}; ${casualtyText}, ${evacuationText}. ${regulatorText} ${radioText} ${assignedText} ${photoText}`;
}

function resolveViewFromHash() {
  const view = window.location.hash.replace("#", "");
  return VALID_VIEWS.has(view) ? view : "create";
}

function applyView(view) {
  state.currentView = view;

  dom.views.forEach((section) => {
    section.classList.toggle("is-active", section.dataset.view === view);
  });

  dom.navButtons.forEach((button) => {
    button.classList.toggle("is-active", button.dataset.viewTarget === view);
  });

  if (view === "incidents" && !getSelectedIncident() && state.incidents.length > 0) {
    state.selectedIncidentId = sortIncidents(state.incidents)[0].incidentId;
  }

  renderIncidentList();
  renderIncidentDetail();
  renderReporting();
}

function navigateToView(view) {
  if (!VALID_VIEWS.has(view)) {
    return;
  }

  if (window.location.hash !== `#${view}`) {
    window.location.hash = view;
    return;
  }

  applyView(view);
}

function renderMetrics() {
  const openIncidents = state.incidents.filter((incident) => incident.status !== "Closed").length;
  const highSeverity = state.incidents.filter((incident) => severityRank[incident.severity] >= severityRank.High).length;
  const medicalCases = state.incidents.filter((incident) => incident.medicalResponseRequired).length;
  const pendingNotifications = state.incidents.filter((incident) => !incident.regulatorNotified && incident.status !== "Closed").length;

  dom.metricOpen.textContent = String(openIncidents);
  dom.metricHigh.textContent = String(highSeverity);
  dom.metricMedical.textContent = String(medicalCases);
  dom.metricPending.textContent = String(pendingNotifications);
  dom.metricPersonnel.textContent = String(state.personnel.length);
}

function renderSettingsChannelList() {
  dom.settingsChannelList.innerHTML = RADIO_CHANNELS.map((channel) => `<span class="chip">${escapeHtml(channel)}</span>`).join("");
}

function renderReportingIncidentTypeOptions(selectedValue = "") {
  const normalizedValue = singleLine(selectedValue);
  const formField = dom.incidentForm.elements.namedItem("incidentType");
  const typeSet = new Set();

  if (formField instanceof HTMLSelectElement) {
    [...formField.options]
      .map((option) => singleLine(option.value || option.textContent))
      .filter(Boolean)
      .forEach((type) => typeSet.add(type));
  }

  state.incidents
    .map((incident) => singleLine(incident.incidentType))
    .filter(Boolean)
    .forEach((type) => typeSet.add(type));

  const types = [...typeSet].sort((a, b) => a.localeCompare(b));
  if (normalizedValue && !types.includes(normalizedValue)) {
    types.unshift(normalizedValue);
  }

  dom.reportingIncidentTypeFilter.innerHTML = [
    '<option value="">All incident types</option>',
    ...types.map((type) => `<option value="${escapeHtml(type)}">${escapeHtml(type)}</option>`),
  ].join("");
  dom.reportingIncidentTypeFilter.value = normalizedValue;
}

function setReportingFormValues(filters, statusMessage = "") {
  renderReportingIncidentTypeOptions(filters.incidentType);
  dom.reportingPreset.value = filters.preset;
  dom.reportingLocationFilter.value = filters.location;
  dom.reportingDateFrom.value = filters.dateFrom;
  dom.reportingDateTo.value = filters.dateTo;
  dom.reportingStatusFilter.value = filters.status;
  dom.reportingSeverityFilter.value = filters.severity;
  dom.reportingIncidentTypeFilter.value = filters.incidentType;
  dom.reportingDateFrom.disabled = filters.preset === "all";
  dom.reportingDateTo.disabled = filters.preset === "all";
  dom.reportingFormStatus.textContent =
    statusMessage || "Filter the incident register and generate a current reporting view.";
}

function getReportingFormValues() {
  return normalizeReportingFilters({
    preset: dom.reportingPreset.value,
    location: dom.reportingLocationFilter.value,
    dateFrom: dom.reportingDateFrom.value,
    dateTo: dom.reportingDateTo.value,
    status: dom.reportingStatusFilter.value,
    severity: dom.reportingSeverityFilter.value,
    incidentType: dom.reportingIncidentTypeFilter.value,
  });
}

function reportPresetLabel(preset) {
  switch (preset) {
    case "last7":
      return "Last 7 days";
    case "last90":
      return "Last 90 days";
    case "ytd":
      return "Year to date";
    case "all":
      return "All incidents";
    case "custom":
      return "Custom range";
    case "last30":
    default:
      return "Last 30 days";
  }
}

function describeReportingPeriod(filters, incidents) {
  if (filters.dateFrom || filters.dateTo) {
    const fromText = filters.dateFrom ? formatDateLabel(filters.dateFrom) : "Start";
    const toText = filters.dateTo ? formatDateLabel(filters.dateTo) : "Now";
    return `${fromText} to ${toText}`;
  }

  if (filters.preset === "all" && incidents.length > 0) {
    const datedIncidents = incidents
      .map((incident) => incidentReportedAtDate(incident))
      .filter((date) => date instanceof Date)
      .sort((a, b) => a.valueOf() - b.valueOf());

    if (datedIncidents.length > 0) {
      return `${formatDateLabel(datedIncidents[0])} to ${formatDateLabel(datedIncidents[datedIncidents.length - 1])}`;
    }
  }

  return reportPresetLabel(filters.preset);
}

function describeReportingScope(filters) {
  const scope = [];

  if (filters.location) {
    scope.push(filters.location);
  }

  if (filters.status) {
    scope.push(filters.status);
  }

  if (filters.severity) {
    scope.push(filters.severity);
  }

  if (filters.incidentType) {
    scope.push(filters.incidentType);
  }

  return scope.length > 0 ? scope.join(" / ") : "All incidents";
}

function incidentMatchesReportingFilters(incident, filters) {
  if (filters.location && incident.templateLocationType !== filters.location) {
    return false;
  }

  if (filters.status && incident.status !== filters.status) {
    return false;
  }

  if (filters.severity && incident.severity !== filters.severity) {
    return false;
  }

  if (filters.incidentType && incident.incidentType !== filters.incidentType) {
    return false;
  }

  const reportedAt = incidentReportedAtDate(incident);
  const fromDate = filters.dateFrom ? startOfDay(`${filters.dateFrom}T00:00`) : null;
  const toDate = filters.dateTo ? endOfDay(`${filters.dateTo}T00:00`) : null;

  if (fromDate && (!reportedAt || reportedAt.valueOf() < fromDate.valueOf())) {
    return false;
  }

  if (toDate && (!reportedAt || reportedAt.valueOf() > toDate.valueOf())) {
    return false;
  }

  return true;
}

function buildReportingSummary(snapshot) {
  if (snapshot.incidents.length === 0) {
    return `No incidents matched ${snapshot.scopeLabel.toLowerCase()} for ${snapshot.periodLabel}. Adjust the filters and generate the report again.`;
  }

  const incidentCountText = `${snapshot.incidents.length} incident${snapshot.incidents.length === 1 ? "" : "s"}`;
  const openText =
    snapshot.openIncidents > 0
      ? `${snapshot.openIncidents} still open or under active management`
      : "no incidents currently open";
  const highText =
    snapshot.highSeverity > 0
      ? `${snapshot.highSeverity} were rated high or critical`
      : "no high or critical incidents were recorded";
  const medicalText =
    snapshot.medicalCases > 0
      ? `${snapshot.medicalCases} required medical response`
      : "no medical response was recorded";
  const typeText = snapshot.typeBreakdown[0]?.label ? `${snapshot.typeBreakdown[0].label} was the most common type` : "";
  const channelText =
    snapshot.channelBreakdown[0]?.label && snapshot.channelBreakdown[0].label !== "Not recorded"
      ? `${snapshot.channelBreakdown[0].label} carried the highest radio traffic`
      : "radio channel activity was limited or not recorded";
  const controlText = Number.isFinite(snapshot.avgControlMinutes)
    ? `average time to control was ${formatDurationMinutes(snapshot.avgControlMinutes)}`
    : "time-to-control data was not consistently recorded";

  return `${incidentCountText} matched ${snapshot.scopeLabel.toLowerCase()} for ${snapshot.periodLabel}. ${openText}, ${highText}, and ${medicalText}. ${typeText ? `${typeText}, ` : ""}${controlText}. ${channelText}.`;
}

function buildReportingSnapshot(filters) {
  const incidents = sortIncidents(state.incidents).filter((incident) => incidentMatchesReportingFilters(incident, filters));
  const typeBreakdown = tallyBy(incidents, (incident) => incident.incidentType);
  const severityBreakdown = tallyBy(incidents, (incident) => incident.severity);
  const statusBreakdown = tallyBy(incidents, (incident) => incident.status);
  const locationBreakdown = tallyBy(incidents, (incident) => incident.templateLocationType);
  const channelBreakdown = tallyBy(
    incidents.flatMap((incident) => incident.radioLogs),
    (entry) => entry.channel || "Not recorded"
  );
  const arrivalMinutes = incidents.map((incident) =>
    minutesBetween(incidentReportedAtDate(incident), combineIncidentDateAndTime(incident, incident.arrivalAtSiteTime))
  );
  const controlMinutes = incidents.map((incident) =>
    minutesBetween(incidentReportedAtDate(incident), parseDateValue(incident.controlledAt))
  );
  const downtimeValues = incidents
    .map((incident) => Number(incident.downtimeHours))
    .filter((value) => Number.isFinite(value) && value > 0);
  const personnelMap = new Map();
  const uniquePersonnelIds = new Set();

  incidents.forEach((incident) => {
    getAssignedWorkers(incident.assignedPersonnelIds).forEach((worker) => {
      uniquePersonnelIds.add(worker.id);
      const entry = personnelMap.get(worker.id) || {
        id: worker.id,
        name: worker.fullName,
        role: worker.role,
        crew: worker.crew,
        count: 0,
      };
      entry.count += 1;
      personnelMap.set(worker.id, entry);
    });
  });

  const topPersonnel = [...personnelMap.values()]
    .sort((a, b) => b.count - a.count || a.name.localeCompare(b.name))
    .slice(0, 6);

  const snapshot = {
    filters,
    generatedAt: new Date(),
    incidents,
    periodLabel: describeReportingPeriod(filters, incidents),
    scopeLabel: describeReportingScope(filters),
    openIncidents: incidents.filter((incident) => incident.status !== "Closed").length,
    highSeverity: incidents.filter((incident) => severityRank[incident.severity] >= severityRank.High).length,
    medicalCases: incidents.filter((incident) => incident.medicalResponseRequired).length,
    regulatorPending: incidents.filter((incident) => !incident.regulatorNotified && incident.status !== "Closed").length,
    incidentsWithPhotos: incidents.filter((incident) => incident.incidentPhotos.length > 0).length,
    totalCasualties: incidents.reduce((total, incident) => total + normalizeNumber(incident.casualtyCount), 0),
    totalEvacuated: incidents.reduce((total, incident) => total + normalizeNumber(incident.evacuatedCount), 0),
    totalRadioEntries: incidents.reduce((total, incident) => total + incident.radioLogs.length, 0),
    totalAssignedSlots: incidents.reduce((total, incident) => total + incident.assignedPersonnelIds.length, 0),
    uniquePersonnelCount: uniquePersonnelIds.size,
    avgArrivalMinutes: averageNumbers(arrivalMinutes),
    avgControlMinutes: averageNumbers(controlMinutes),
    avgDowntimeHours: averageNumbers(downtimeValues),
    typeBreakdown,
    severityBreakdown,
    statusBreakdown,
    locationBreakdown,
    channelBreakdown,
    topPersonnel,
  };

  snapshot.summaryText = buildReportingSummary(snapshot);
  return snapshot;
}

function renderReportingMetricCards(snapshot) {
  const metrics = [
    ["Matched incidents", String(snapshot.incidents.length)],
    ["Open / active", String(snapshot.openIncidents)],
    ["High / critical", String(snapshot.highSeverity)],
    ["Medical cases", String(snapshot.medicalCases)],
    ["Regulatory pending", String(snapshot.regulatorPending)],
    ["Workers engaged", String(snapshot.uniquePersonnelCount)],
  ];

  return metrics
    .map(
      ([label, value]) => `
        <article class="reporting-metric-card">
          <span class="metric-label">${escapeHtml(label)}</span>
          <strong>${escapeHtml(value)}</strong>
        </article>
      `
    )
    .join("");
}

function renderBreakdownCard(title, items, emptyMessage) {
  if (items.length === 0) {
    return `
      <article class="breakdown-card">
        <div class="breakdown-card-head">
          <h4>${escapeHtml(title)}</h4>
        </div>
        <div class="empty-note">${escapeHtml(emptyMessage)}</div>
      </article>
    `;
  }

  const visibleItems = items.slice(0, 6);
  const maxCount = Math.max(...visibleItems.map((item) => item.count), 1);

  return `
    <article class="breakdown-card">
      <div class="breakdown-card-head">
        <h4>${escapeHtml(title)}</h4>
        <span>${visibleItems.length} item${visibleItems.length === 1 ? "" : "s"}</span>
      </div>
      <div class="breakdown-list">
        ${visibleItems
          .map(
            (item) => `
              <div class="breakdown-row">
                <div class="breakdown-row-head">
                  <strong>${escapeHtml(item.label)}</strong>
                  <span>${item.count}</span>
                </div>
                <div class="breakdown-track">
                  <span class="breakdown-fill" style="width: ${(item.count / maxCount) * 100}%"></span>
                </div>
              </div>
            `
          )
          .join("")}
      </div>
    </article>
  `;
}

function renderReportingPerformance(snapshot) {
  const stats = [
    ["Average arrival to site", formatDurationMinutes(snapshot.avgArrivalMinutes)],
    ["Average time to control", formatDurationMinutes(snapshot.avgControlMinutes)],
    ["Average downtime", formatHoursValue(snapshot.avgDowntimeHours)],
    ["Radio log entries", String(snapshot.totalRadioEntries)],
    ["Personnel assignment slots", String(snapshot.totalAssignedSlots)],
    ["Incidents with photos", String(snapshot.incidentsWithPhotos)],
    ["Total casualties", String(snapshot.totalCasualties)],
    ["Total evacuated", String(snapshot.totalEvacuated)],
    ["Most active channel", snapshot.channelBreakdown[0]?.label || "Not recorded"],
  ];

  return stats
    .map(
      ([label, value]) => `
        <article class="reporting-stat-card">
          <strong>${escapeHtml(label)}</strong>
          <span>${escapeHtml(value)}</span>
        </article>
      `
    )
    .join("");
}

function renderReportingPersonnel(snapshot) {
  if (snapshot.topPersonnel.length === 0) {
    return '<div class="empty-note">No assigned workers were found in the filtered incident set.</div>';
  }

  return snapshot.topPersonnel
    .map(
      (worker) => `
        <article class="people-card">
          <strong>${escapeHtml(worker.name)}</strong>
          <span>${escapeHtml(worker.role || "Role not recorded")} · ${escapeHtml(worker.crew || "Crew not recorded")}</span>
          <span>${worker.count} incident${worker.count === 1 ? "" : "s"} assigned</span>
        </article>
      `
    )
    .join("");
}

function renderReportingIncidents(snapshot) {
  if (snapshot.incidents.length === 0) {
    return '<tr><td class="reporting-empty-row" colspan="9">No incidents match the current report filters.</td></tr>';
  }

  return snapshot.incidents
    .map((incident) => {
      const assignedWorkers = getAssignedWorkers(incident.assignedPersonnelIds);
      return `
        <tr>
          <td><strong>${escapeHtml(incident.incidentId)}</strong></td>
          <td>${escapeHtml(formatDateTime(incident.reportedAt))}</td>
          <td>${escapeHtml(incident.status)}</td>
          <td>${escapeHtml(incident.severity)}</td>
          <td>${escapeHtml(incident.incidentType)}</td>
          <td>${escapeHtml(incidentLocationText(incident))}</td>
          <td>${assignedWorkers.length}</td>
          <td>${normalizeNumber(incident.casualtyCount)}</td>
          <td>
            <button class="report-link-button" type="button" data-report-open="${escapeHtml(incident.incidentId)}">
              Open
            </button>
          </td>
        </tr>
      `;
    })
    .join("");
}

function buildReportingClipboardText(snapshot) {
  const topTypes = snapshot.typeBreakdown
    .slice(0, 3)
    .map((item) => `${item.label} (${item.count})`)
    .join(", ");
  const topPersonnel = snapshot.topPersonnel
    .slice(0, 5)
    .map((worker) => `${worker.name} (${worker.count})`)
    .join(", ");

  return [
    `Incident reporting summary`,
    `Scope: ${snapshot.scopeLabel}`,
    `Period: ${snapshot.periodLabel}`,
    `Generated: ${formatDateTime(snapshot.generatedAt)}`,
    "",
    snapshot.summaryText,
    "",
    `Matched incidents: ${snapshot.incidents.length}`,
    `Open / active: ${snapshot.openIncidents}`,
    `High / critical: ${snapshot.highSeverity}`,
    `Medical cases: ${snapshot.medicalCases}`,
    `Regulatory pending: ${snapshot.regulatorPending}`,
    `Average arrival to site: ${formatDurationMinutes(snapshot.avgArrivalMinutes)}`,
    `Average time to control: ${formatDurationMinutes(snapshot.avgControlMinutes)}`,
    `Average downtime: ${formatHoursValue(snapshot.avgDowntimeHours)}`,
    `Total radio entries: ${snapshot.totalRadioEntries}`,
    `Workers engaged: ${snapshot.uniquePersonnelCount}`,
    `Top incident types: ${topTypes || "None"}`,
    `Top personnel: ${topPersonnel || "None"}`,
  ].join("\n");
}

function buildReportingCsv(snapshot) {
  const headers = [
    "Incident ID",
    "Reported At",
    "Status",
    "Severity",
    "Incident Type",
    "Location",
    "ESO Supervisor",
    "Lead Responder",
    "Personnel Count",
    "Assigned Workers",
    "Casualties",
    "Evacuated",
    "Medical Response Required",
    "Regulator Notified",
    "Radio Entries",
    "Photos Attached",
    "Controlled At",
  ];

  const rows = snapshot.incidents.map((incident) => {
    const assignedWorkers = getAssignedWorkers(incident.assignedPersonnelIds)
      .map((worker) => worker.fullName)
      .join(", ");

    return [
      incident.incidentId,
      formatDateTime(incident.reportedAt),
      incident.status,
      incident.severity,
      incident.incidentType,
      incidentLocationText(incident),
      incident.esoSupervisor,
      incident.leadResponder,
      incident.personnelCount,
      assignedWorkers,
      incident.casualtyCount,
      incident.evacuatedCount,
      incident.medicalResponseRequired ? "Yes" : "No",
      incident.regulatorNotified ? "Yes" : "No",
      incident.radioLogs.length,
      incident.incidentPhotos.length,
      formatDateTime(incident.controlledAt),
    ];
  });

  return [headers, ...rows].map((row) => row.map(csvCell).join(",")).join("\n");
}

function setReportingActionState(enabled) {
  [dom.copyReportingSummaryButton, dom.exportReportingCsvButton].forEach((button) => {
    button.disabled = !enabled;
  });
}

function renderReporting(statusMessage = "") {
  const filters = state.reportingFilters || defaultReportingFilters();
  const snapshot = buildReportingSnapshot(filters);

  state.reportingFilters = filters;
  state.reportingSnapshot = snapshot;

  setReportingFormValues(filters, statusMessage || `Showing ${snapshot.incidents.length} incident${snapshot.incidents.length === 1 ? "" : "s"} for ${snapshot.scopeLabel.toLowerCase()}.`);
  dom.reportingTitle.textContent = `Operational report · ${snapshot.periodLabel}`;
  dom.reportingSummary.textContent = snapshot.summaryText;
  dom.reportingCountBadge.textContent = `${snapshot.incidents.length} incident${snapshot.incidents.length === 1 ? "" : "s"}`;
  dom.reportingGeneratedAt.textContent = `Generated ${new Intl.DateTimeFormat(undefined, {
    dateStyle: "full",
    timeStyle: "short",
  }).format(snapshot.generatedAt)}`;
  dom.reportingMetrics.innerHTML = renderReportingMetricCards(snapshot);
  dom.reportingBreakdowns.innerHTML = [
    renderBreakdownCard("Incident type", snapshot.typeBreakdown, "No incident types are available for this filter set."),
    renderBreakdownCard("Severity", snapshot.severityBreakdown, "No severity data is available for this filter set."),
    renderBreakdownCard("Status", snapshot.statusBreakdown, "No status data is available for this filter set."),
    renderBreakdownCard("Location", snapshot.locationBreakdown, "No location data is available for this filter set."),
  ].join("");
  dom.reportingPerformance.innerHTML = renderReportingPerformance(snapshot);
  dom.reportingPersonnel.innerHTML = renderReportingPersonnel(snapshot);
  dom.reportingIncidents.innerHTML = renderReportingIncidents(snapshot);
  setReportingActionState(snapshot.incidents.length > 0);
}

async function copyReportingSummary() {
  const snapshot = state.reportingSnapshot;
  if (!snapshot || snapshot.incidents.length === 0) {
    return;
  }

  try {
    await navigator.clipboard.writeText(buildReportingClipboardText(snapshot));
    const previous = dom.copyReportingSummaryButton.textContent;
    dom.copyReportingSummaryButton.textContent = "Copied";
    window.setTimeout(() => {
      dom.copyReportingSummaryButton.textContent = previous;
    }, 1200);
  } catch {
    window.alert("Clipboard access is unavailable in this browser.");
  }
}

function exportReportingCsv() {
  const snapshot = state.reportingSnapshot;
  if (!snapshot || snapshot.incidents.length === 0) {
    return;
  }

  const blob = new Blob([buildReportingCsv(snapshot)], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = `incident-report-${toLocalDateValue(snapshot.generatedAt)}.csv`;
  anchor.click();
  URL.revokeObjectURL(url);
}

function setSettingsFormValues(settings, statusMessage = "") {
  dom.settingsRadioUnits.value = settings.radioUnits.join("\n");
  dom.settingsFormStatus.textContent = statusMessage || "Maintain the dropdown list used in the radio operator log.";
}

function getSettingsFormValues() {
  return normalizeSettings({
    radioUnits: dom.settingsRadioUnits.value.split("\n"),
  });
}

function renderRadioUnitOptions(selectedValue = "") {
  const normalizedValue = singleLine(selectedValue);
  const units = [...state.settings.radioUnits];

  if (normalizedValue && !units.includes(normalizedValue)) {
    units.unshift(normalizedValue);
  }

  dom.radioLogUnit.innerHTML = [
    '<option value="">Select unit</option>',
    ...units.map((unit) => `<option value="${escapeHtml(unit)}">${escapeHtml(unit)}</option>`),
  ].join("");
  dom.radioLogUnit.value = normalizedValue;
}

function renderRadioChannelOptions(selectedValue = "") {
  const normalizedValue = singleLine(selectedValue);
  const channels = [...RADIO_CHANNELS];

  if (normalizedValue && !channels.includes(normalizedValue)) {
    channels.unshift(normalizedValue);
  }

  dom.radioLogChannel.innerHTML = [
    '<option value="">Select channel</option>',
    ...channels.map((channel) => `<option value="${escapeHtml(channel)}">${escapeHtml(channel)}</option>`),
  ].join("");
  dom.radioLogChannel.value = normalizedValue;
}

async function saveSettings(settings) {
  const wasCloudBacked = state.storageMode === "cloud";
  state.settings = settings;
  const synced = await persistSettings();
  setSettingsFormValues(
    state.settings,
    synced
      ? "Radio log settings saved to storage."
      : wasCloudBacked
        ? "Radio log settings saved locally. Cloud sync failed."
        : "Radio log settings could not be saved in browser storage."
  );
  renderRadioUnitOptions(dom.radioLogUnit.value);
  renderRadioChannelOptions(dom.radioLogChannel.value);
}

function renderPersonnelSelect(fieldName, placeholder, selectedValue = "") {
  const field = dom.incidentForm.elements.namedItem(fieldName);
  if (!(field instanceof HTMLSelectElement)) {
    return;
  }

  const normalizedValue = singleLine(selectedValue);
  const names = [...new Set(state.personnel.map((worker) => singleLine(worker.fullName)).filter(Boolean))].sort((a, b) =>
    a.localeCompare(b)
  );

  if (normalizedValue && !names.includes(normalizedValue)) {
    names.unshift(normalizedValue);
  }

  field.innerHTML = [
    `<option value="">${escapeHtml(placeholder)}</option>`,
    ...names.map((name) => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`),
  ].join("");
  field.value = normalizedValue;
}

function renderIncidentPersonnelRoleOptions(values = {}) {
  renderPersonnelSelect(
    "esoSupervisor",
    "Select supervisor",
    values.esoSupervisor ?? dom.incidentForm.elements.namedItem("esoSupervisor")?.value ?? ""
  );
  renderPersonnelSelect(
    "reportCompiledBy",
    "Select report compiler",
    values.reportCompiledBy ?? dom.incidentForm.elements.namedItem("reportCompiledBy")?.value ?? ""
  );
  renderPersonnelSelect(
    "leadResponder",
    "Select lead responder",
    values.leadResponder ?? dom.incidentForm.elements.namedItem("leadResponder")?.value ?? ""
  );
}

function photoCountLabel(count) {
  return `${count} photo${count === 1 ? "" : "s"}`;
}

function renderIncidentPhotoList() {
  dom.incidentPhotoList.innerHTML = "";

  if (state.currentIncidentPhotos.length === 0) {
    dom.incidentPhotoList.innerHTML = '<div class="empty-note">No images uploaded for this incident yet.</div>';
    return;
  }

  dom.incidentPhotoList.innerHTML = state.currentIncidentPhotos
    .map(
      (photo) => `
        <article class="photo-card">
          <div class="photo-thumb">
            <img src="${escapeHtml(incidentPhotoSource(photo))}" alt="${escapeHtml(photo.name)}" loading="lazy" />
          </div>
          <div class="photo-card-meta">
            <strong>${escapeHtml(photo.name)}</strong>
            <button class="btn btn-ghost btn-small" type="button" data-photo-remove="${escapeHtml(photo.id)}">Remove</button>
          </div>
        </article>
      `
    )
    .join("");
}

function renderPhotoGrid(photos, emptyMessage) {
  if (photos.length === 0) {
    return `<div class="empty-note">${emptyMessage}</div>`;
  }

  return photos
    .map(
      (photo) => `
        <article class="photo-card">
          <div class="photo-thumb">
            <img src="${escapeHtml(incidentPhotoSource(photo))}" alt="${escapeHtml(photo.name)}" loading="lazy" />
          </div>
          <div class="photo-card-meta">
            <strong>${escapeHtml(photo.name)}</strong>
          </div>
        </article>
      `
    )
    .join("");
}

async function handleIncidentPhotoUpload(fileList) {
  const files = [...fileList].filter((file) => file.type.startsWith("image/"));

  if (files.length === 0) {
    dom.incidentPhotoStatus.textContent = "Select one or more image files to attach them to the incident.";
    dom.incidentPhotoUpload.value = "";
    return;
  }

  dom.incidentPhotoStatus.textContent = `Processing ${photoCountLabel(files.length)}...`;

  try {
    const preparedPhotos = await Promise.all(files.map((file) => prepareIncidentPhoto(file)));
    state.currentIncidentPhotos = normalizeIncidentPhotos([...state.currentIncidentPhotos, ...preparedPhotos]);
    dom.incidentPhotoUpload.value = "";
    renderIncidentPhotoList();
    renderDraftSidebar();
    dom.incidentPhotoStatus.textContent = `Added ${photoCountLabel(preparedPhotos.length)} to this incident draft.`;
  } catch {
    dom.incidentPhotoStatus.textContent = "Unable to process the selected image files.";
  }
}

function renderPersonnelPicker(esoTeamFallback = "") {
  renderIncidentPersonnelRoleOptions();
  state.selectedPersonnelIds = state.selectedPersonnelIds.filter((workerId) => Boolean(getWorkerById(workerId)));
  const assignedWorkers = getAssignedWorkers(state.selectedPersonnelIds);
  syncEsoTeamField(esoTeamFallback);

  dom.personnelPickerToggle.textContent =
    assignedWorkers.length > 0
      ? `${assignedWorkers.length} worker${assignedWorkers.length === 1 ? "" : "s"} selected`
      : "Select Workers";

  dom.assignedPersonnelChips.innerHTML =
    assignedWorkers.length > 0
      ? assignedWorkers
          .map(
            (worker) =>
              `<span class="chip">${escapeHtml(worker.fullName)} · ${escapeHtml(worker.role)}</span>`
          )
          .join("")
      : `<div class="empty-note">No workers assigned yet.</div>`;

  if (state.personnel.length === 0) {
    dom.personnelPickerOptions.innerHTML =
      '<div class="empty-state">No workers are in the roster yet. Open the Personnel view and add them first.</div>';
    return;
  }

  dom.personnelPickerOptions.innerHTML = [...state.personnel]
    .sort((a, b) => a.fullName.localeCompare(b.fullName))
    .map(
      (worker) => `
        <label class="multiselect-option">
          <input
            type="checkbox"
            data-picker-checkbox
            value="${escapeHtml(worker.id)}"
            ${state.selectedPersonnelIds.includes(worker.id) ? "checked" : ""}
          />
          <span>
            <strong>${escapeHtml(worker.fullName)}</strong>
            <span>${escapeHtml(worker.role)} · ${escapeHtml(worker.crew || "Crew not recorded")}</span>
          </span>
        </label>
      `
    )
    .join("");
}

function togglePersonnelPicker(forceOpen) {
  const shouldOpen =
    typeof forceOpen === "boolean" ? forceOpen : dom.personnelPickerPanel.classList.contains("hidden");

  dom.personnelPickerPanel.classList.toggle("hidden", !shouldOpen);
  dom.personnelPickerToggle.setAttribute("aria-expanded", String(shouldOpen));
}

function clearRadioLogDraft(timestamp = "", statusMessage = "") {
  dom.radioLogTime.value = timestamp || currentReportedAtValue();
  renderRadioUnitOptions("");
  dom.radioLogMovement.value = "Dispatched";
  renderRadioChannelOptions("");
  dom.radioLogLocation.value = "";
  dom.radioLogNotes.value = "";
  state.editingRadioLogId = null;
  dom.saveRadioLogButton.textContent = "Add Radio Log Entry";
  dom.cancelRadioEditButton.classList.add("hidden");
  dom.radioLogStatus.textContent = statusMessage || DEFAULT_RADIO_STATUS;
}

function setRadioLogDraft(entry) {
  dom.radioLogTime.value = entry.timestamp || "";
  renderRadioUnitOptions(entry.unit || "");
  dom.radioLogMovement.value = entry.movement || "Other";
  renderRadioChannelOptions(entry.channel || "");
  dom.radioLogLocation.value = entry.location || "";
  dom.radioLogNotes.value = entry.notes || "";
  state.editingRadioLogId = entry.id;
  dom.saveRadioLogButton.textContent = "Update Radio Log Entry";
  dom.cancelRadioEditButton.classList.remove("hidden");
  dom.radioLogStatus.textContent = `Editing ${entry.unit || "radio entry"} from ${formatDateTime(entry.timestamp)}.`;
}

function renderRadioLogList() {
  dom.radioLogList.innerHTML = "";

  if (state.currentRadioLogs.length === 0) {
    dom.radioLogList.innerHTML = '<div class="empty-note">No radio movement entries added yet.</div>';
    return;
  }

  state.currentRadioLogs.forEach((entry) => {
    const item = document.createElement("article");
    item.className = "radio-log-item";
    item.innerHTML = `
      <div class="radio-log-item-top">
        <span class="radio-log-stamp">${escapeHtml(formatDateTime(entry.timestamp))}</span>
        <div class="radio-log-item-actions">
          <button class="btn btn-ghost btn-small" type="button" data-radio-action="edit">Edit</button>
          <button class="btn btn-ghost btn-small" type="button" data-radio-action="remove">Remove</button>
        </div>
      </div>
      <div>
        <h4>${escapeHtml(entry.unit)} · ${escapeHtml(entry.movement)}</h4>
        <p>${escapeHtml(entry.notes || "No operator notes recorded.")}</p>
      </div>
      <div class="radio-log-item-meta">
        <span class="radio-log-chip">Channel: ${escapeHtml(entry.channel || "Not recorded")}</span>
        <span class="radio-log-chip">Location: ${escapeHtml(entry.location || "Not recorded")}</span>
      </div>
    `;

    item.querySelector('[data-radio-action="edit"]').addEventListener("click", () => {
      setRadioLogDraft(entry);
    });

    item.querySelector('[data-radio-action="remove"]').addEventListener("click", () => {
      state.currentRadioLogs = state.currentRadioLogs.filter((log) => log.id !== entry.id);
      renderRadioLogList();
      renderDraftSidebar();

      if (state.editingRadioLogId === entry.id) {
        clearRadioLogDraft(currentReportedAtValue(), `Removed radio entry for ${entry.unit || "unit"}.`);
        return;
      }

      dom.radioLogStatus.textContent = `Removed radio entry for ${entry.unit || "unit"}.`;
    });

    dom.radioLogList.append(item);
  });
}

function getRadioLogDraft() {
  return {
    id: state.editingRadioLogId || createRadioLogId(),
    timestamp: dom.radioLogTime.value,
    unit: dom.radioLogUnit.value.trim(),
    movement: dom.radioLogMovement.value,
    channel: dom.radioLogChannel.value.trim(),
    location: dom.radioLogLocation.value.trim(),
    notes: dom.radioLogNotes.value.trim(),
  };
}

function setIncidentFormValues(incident, statusMessage = "") {
  Array.from(dom.incidentForm.elements).forEach((field) => {
    if (!field.name) {
      return;
    }

    if (field.type === "checkbox") {
      field.checked = Boolean(incident[field.name]);
      return;
    }

    field.value = incident[field.name] ?? "";
  });

  renderIncidentPersonnelRoleOptions(incident);
  state.selectedPersonnelIds = incident.assignedPersonnelIds.filter((workerId) => Boolean(getWorkerById(workerId)));
  state.currentRadioLogs = normalizeRadioLogs(incident.radioLogs);
  state.currentIncidentPhotos = normalizeIncidentPhotos(incident.incidentPhotos);
  dom.workspaceTitle.textContent = state.editingIncidentId ? `Editing ${incident.incidentId}` : "Create Incident";
  dom.incidentFormStatus.textContent =
    statusMessage ||
    (state.editingIncidentId
      ? `Editing saved incident ${incident.incidentId}.`
      : "Build a new incident record or load one from the register to edit it.");

  renderPersonnelPicker(incident.esoTeam);
  renderIncidentPhotoList();
  dom.incidentPhotoUpload.value = "";
  dom.incidentPhotoStatus.textContent =
    state.currentIncidentPhotos.length > 0
      ? `${photoCountLabel(state.currentIncidentPhotos.length)} attached to this incident.`
      : "Uploaded images are stored with the incident in this browser.";
  clearRadioLogDraft(incident.reportedAt || toLocalDateTimeValue());
  renderRadioLogList();
  renderDraftSidebar();
}

function getIncidentFormValues() {
  const data = {};

  Array.from(dom.incidentForm.elements).forEach((field) => {
    if (!field.name) {
      return;
    }

    if (field.type === "checkbox") {
      data[field.name] = field.checked;
      return;
    }

    if (field.type === "number") {
      data[field.name] = field.value === "" ? 0 : Number(field.value);
      return;
    }

    data[field.name] = field.value.trim();
  });

  data.assignedPersonnelIds = [...new Set(state.selectedPersonnelIds)].filter((workerId) => Boolean(getWorkerById(workerId)));
  data.esoTeam = buildEsoTeamValue(data.assignedPersonnelIds, data.esoTeam);
  data.radioLogs = normalizeRadioLogs(state.currentRadioLogs);
  data.incidentPhotos = normalizeIncidentPhotos(state.currentIncidentPhotos);
  data.updatedAt = new Date().toISOString();

  return normalizeIncident(data);
}

function startNewIncident(statusMessage = "") {
  state.editingIncidentId = null;
  setIncidentFormValues(defaultIncident(), statusMessage);
}

function loadIncidentIntoForm(incidentId, navigate = true) {
  const incident = state.incidents.find((item) => item.incidentId === incidentId);
  if (!incident) {
    return;
  }

  state.editingIncidentId = incident.incidentId;
  setIncidentFormValues(incident);

  if (navigate) {
    navigateToView("create");
  }
}

async function saveIncident(incidentData) {
  const wasCloudBacked = state.storageMode === "cloud";
  const previousIncidentId = state.editingIncidentId;
  const editingIndex = state.editingIncidentId
    ? state.incidents.findIndex((incident) => incident.incidentId === state.editingIncidentId)
    : -1;
  const exactIndex = state.incidents.findIndex((incident) => incident.incidentId === incidentData.incidentId);

  if (editingIndex >= 0 && state.editingIncidentId !== incidentData.incidentId) {
    if (exactIndex >= 0 && exactIndex !== editingIndex) {
      window.alert(`Incident ID ${incidentData.incidentId} already exists. Use a unique ID.`);
      return;
    }
    state.incidents[editingIndex] = incidentData;
  } else if (exactIndex >= 0) {
    state.incidents[exactIndex] = incidentData;
  } else {
    state.incidents.push(incidentData);
  }

  state.selectedIncidentId = incidentData.incidentId;
  state.editingIncidentId = incidentData.incidentId;
  const synced = await persistIncidents({ mode: "single", incident: incidentData });

  if (synced && previousIncidentId && previousIncidentId !== incidentData.incidentId && state.storageMode === "cloud") {
    try {
      await apiDeleteIncident(previousIncidentId);
    } catch (error) {
      console.error(`Unable to remove renamed cloud incident ${previousIncidentId}.`, error);
    }
  }

  renderMetrics();
  renderIncidentList();
  renderIncidentDetail();
  renderReporting();
  const savedIncident = state.incidents.find((incident) => incident.incidentId === incidentData.incidentId) || incidentData;
  setIncidentFormValues(
    savedIncident,
    synced
      ? `Saved incident ${incidentData.incidentId}.`
      : wasCloudBacked
        ? `Saved incident ${incidentData.incidentId} locally. Cloud sync failed.`
        : `Incident ${incidentData.incidentId} could not be saved in browser storage.`
  );
}

function renderDraftSidebar() {
  const incident = getIncidentFormValues();
  const assignedWorkers = getAssignedWorkers(incident.assignedPersonnelIds);

  dom.draftTitle.textContent = state.editingIncidentId ? `Editing ${incident.incidentId}` : `Draft ${incident.incidentId}`;
  dom.draftSummaryText.textContent = makeSummary(incident);
  dom.draftBadges.innerHTML = `
    <span class="status-pill">${escapeHtml(incident.status)}</span>
    <span class="status-pill">${escapeHtml(incident.severity)}</span>
    <span class="status-pill">${escapeHtml(incident.incidentType)}</span>
  `;
  dom.draftMiniMetrics.innerHTML = `
    <article class="mini-metric">
      <span class="metric-label">Assigned workers</span>
      <strong>${assignedWorkers.length}</strong>
    </article>
    <article class="mini-metric">
      <span class="metric-label">Radio log entries</span>
      <strong>${incident.radioLogs.length}</strong>
    </article>
    <article class="mini-metric">
      <span class="metric-label">Casualties</span>
      <strong>${incident.casualtyCount}</strong>
    </article>
    <article class="mini-metric">
      <span class="metric-label">Location</span>
      <strong>${escapeHtml(incident.locationDetail || "Pending")}</strong>
    </article>
  `;

  dom.draftAssignedList.innerHTML =
    assignedWorkers.length > 0
      ? assignedWorkers
          .map(
            (worker) => `
              <div class="stack-item">
                <strong>${escapeHtml(worker.fullName)}</strong>
                <span>${escapeHtml(worker.role)} · ${escapeHtml(worker.crew || "Crew not recorded")}</span>
              </div>
            `
          )
          .join("")
      : '<div class="empty-note">No workers assigned to this draft yet.</div>';

  dom.draftRadioPreview.innerHTML =
    incident.radioLogs.length > 0
      ? incident.radioLogs
          .map(
            (entry) => `
              <article class="timeline-item">
                <strong>${escapeHtml(formatDateTime(entry.timestamp))} · ${escapeHtml(entry.unit)}</strong>
                <p>${escapeHtml(entry.movement)}${entry.location ? ` · ${escapeHtml(entry.location)}` : ""}</p>
              </article>
            `
          )
          .join("")
      : '<div class="empty-note">No radio timeline entries added yet.</div>';
}

function renderIncidentList() {
  const query = dom.incidentSearch.value.trim().toLowerCase();
  const incidents = sortIncidents(state.incidents).filter((incident) => {
    if (!query) {
      return true;
    }

    const haystack = [incident.incidentId, incident.incidentType, incident.templateLocationType, incident.locationDetail]
      .join(" ")
      .toLowerCase();
    return haystack.includes(query);
  });

  dom.incidentList.innerHTML = "";

  if (incidents.length === 0) {
    dom.incidentList.innerHTML = `<div class="empty-state">${query ? "No incidents match the search." : "No incidents saved yet."}</div>`;
    return;
  }

  incidents.forEach((incident) => {
    const fragment = dom.incidentTemplate.content.cloneNode(true);
    const button = fragment.querySelector(".incident-card");
    const statusElement = fragment.querySelector(".incident-card-status");

    button.classList.toggle("is-active", incident.incidentId === state.selectedIncidentId);
    fragment.querySelector(".incident-card-id").textContent = incident.incidentId;
    statusElement.textContent = incident.status;
    statusElement.style.background = statusTone(incident.status);
    statusElement.style.color = statusTextColor(incident.status);
    fragment.querySelector(".incident-card-type").textContent = `${incident.severity} · ${incident.incidentType}`;
    fragment.querySelector(".incident-card-location").textContent = incidentLocationText(incident);
    fragment.querySelector(".incident-card-updated").textContent = `Updated ${formatDateTime(incident.updatedAt)}`;
    fragment.querySelector(".incident-card-assigned").textContent = `${incident.assignedPersonnelIds.length} worker${
      incident.assignedPersonnelIds.length === 1 ? "" : "s"
    } assigned`;

    button.addEventListener("click", () => {
      state.selectedIncidentId = incident.incidentId;
      renderIncidentList();
      renderIncidentDetail();
    });

    dom.incidentList.append(fragment);
  });
}

function setDetailActionState() {
  const enabled = Boolean(getSelectedIncident());
  [dom.editSelectedIncidentButton, dom.exportJsonButton, dom.copySummaryButton, dom.printReportButton].forEach((button) => {
    button.disabled = !enabled;
  });
}

function renderIncidentDetail() {
  const incident = getSelectedIncident();
  setDetailActionState();

  if (!incident) {
    dom.detailTitle.textContent = "Incident Detail";
    dom.reportEmpty.classList.remove("hidden");
    dom.reportContent.classList.add("hidden");
    dom.reportPhotos.innerHTML = "";
    return;
  }

  const assignedWorkers = getAssignedWorkers(incident.assignedPersonnelIds);

  dom.detailTitle.textContent = `${incident.incidentId} · ${incident.incidentType}`;
  dom.reportEmpty.classList.add("hidden");
  dom.reportContent.classList.remove("hidden");
  dom.reportTitle.textContent = `${incident.incidentId} · ${incident.incidentType}`;
  dom.reportStatusBadge.textContent = incident.status;
  dom.reportStatusBadge.style.background = statusTone(incident.status);
  dom.reportStatusBadge.style.color = statusTextColor(incident.status);

  const activationMethods = getActivationMethods(incident);
  const gridItems = [
    ["Template location", incident.templateLocationType || "Not recorded"],
    ["Exact Location", incident.locationDetail || "Not recorded"],
    ["ESO Team", incident.esoTeam || "Not recorded"],
    ["ESO Supervisor", incident.esoSupervisor || "Not recorded"],
    ["Report Compiled By", incident.reportCompiledBy || "Not recorded"],
    ["Incident category", incident.templateIncidentCategory || "Not recorded"],
    ["Shift", incident.shift || "Not recorded"],
    ["Incident Controller", incident.incidentController || "Not recorded"],
    ["Lead Responder", incident.leadResponder || "Not recorded"],
    ["ESO Primary Role", incident.esoPrimaryRole || "Not recorded"],
    ["Reported At", formatDateTime(incident.reportedAt)],
    ["Controlled At", formatDateTime(incident.controlledAt)],
    ["Arrival At Site", formatTimeOnly(incident.arrivalAtSiteTime)],
    ["Departure", formatTimeOnly(incident.departureTime)],
    ["Activation Method", activationMethods.join(", ") || "Not recorded"],
    ["Personnel Involved", String(incident.personnelCount)],
    ["Assigned Workers", String(assignedWorkers.length)],
    ["Casualties", String(incident.casualtyCount)],
    ["Re-entry Restricted", incident.reentryRestricted ? "Yes" : "No"],
  ];

  dom.reportGrid.innerHTML = gridItems
    .map(
      ([label, value]) => `<div class="report-grid-item"><strong>${label}</strong><span>${escapeHtml(value)}</span></div>`
    )
    .join("");

  dom.reportSummary.textContent = makeSummary(incident);

  dom.reportAssignedPersonnel.innerHTML =
    assignedWorkers.length > 0
      ? assignedWorkers
          .map(
            (worker) => `
              <div class="people-card">
                <strong>${escapeHtml(worker.fullName)}</strong>
                <span>${escapeHtml(worker.role)} · ${escapeHtml(worker.crew || "Crew not recorded")}</span>
              </div>
            `
          )
          .join("")
      : '<div class="empty-note">No workers assigned to this incident.</div>';

  const timelineItems = [
    ["Initial alert", incident.initialAlert],
    ["Immediate actions", incident.initialActions],
    ["Ongoing response", incident.responseActions],
    ["Communications", incident.communicationsLog],
    ["Investigation", incident.investigationNotes],
    ["Recovery", incident.recoveryActions],
  ].filter(([, value]) => value);

  dom.reportTimeline.innerHTML =
    timelineItems.length > 0
      ? timelineItems
          .map(
            ([label, value]) =>
              `<article class="timeline-item"><strong>${escapeHtml(label)}</strong><p>${escapeHtml(value)}</p></article>`
          )
          .join("")
      : '<div class="empty-note">No action timeline has been recorded for this incident.</div>';

  dom.reportRadioTimeline.innerHTML =
    incident.radioLogs.length > 0
      ? incident.radioLogs
          .map(
            (entry) => `
              <article class="timeline-item">
                <strong>${escapeHtml(formatDateTime(entry.timestamp))} · ${escapeHtml(entry.unit)} · ${escapeHtml(
                  entry.movement
                )}</strong>
                <p>${escapeHtml(
                  [entry.location, entry.channel ? `Channel ${entry.channel}` : "", entry.notes].filter(Boolean).join(" · ")
                )}</p>
              </article>
            `
          )
          .join("")
      : '<div class="empty-note">No radio operator log has been recorded for this incident.</div>';

  const noteItems = [
    ["Assets Deployed", incident.equipmentDeployed || "Not recorded"],
    ["Primary Hazards", incident.hazards || "Not recorded"],
    ["Environmental Conditions", incident.environmentalConditions || "Not recorded"],
    ["Casualty Summary", incident.casualtySummary || "Not recorded"],
    ["External Agencies", incident.externalAgenciesInvolved || "Not recorded"],
    ["Handover", incident.handoverAgency || "Not recorded"],
    ["Logistics Issues", incident.logisticsIssues || "Not recorded"],
    ["Communication Issues", incident.communicationIssues || "Not recorded"],
    ["Policy / Procedure Issues", incident.policyProcedureIssues || "Not recorded"],
    ["Mutual Aid Issues", incident.mutualAidIssues || "Not recorded"],
    ["Other Comments", incident.otherIssuesComments || "Not recorded"],
    ["Incident Photo References", incident.incidentPhotoReferences || "Not recorded"],
  ];

  dom.reportNotes.innerHTML = noteItems
    .map(
      ([label, value]) => `<div class="note-card"><strong>${escapeHtml(label)}</strong><span>${escapeHtml(value)}</span></div>`
    )
    .join("");

  dom.reportPhotos.innerHTML = renderPhotoGrid(incident.incidentPhotos, "No images are attached to this incident report.");

  dom.reportGeneratedAt.textContent = `Generated ${new Intl.DateTimeFormat(undefined, {
    dateStyle: "full",
    timeStyle: "short",
  }).format(new Date())}`;
}

function exportSelectedIncident() {
  const incident = getSelectedIncident();
  if (!incident) {
    return;
  }

  const blob = new Blob([JSON.stringify(incident, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = `${incident.incidentId}.json`;
  anchor.click();
  URL.revokeObjectURL(url);
}

function renderHioExport(incident) {
  const involvedPeople = getInvolvedPeople(incident);
  const detailPages = paginateHioDetailBlock(buildHioDetailBlock(incident));
  const totalPages = detailPages.length + 1;
  const continuationPages = detailPages
    .slice(1)
    .map((detailBlock, index) => renderHioContinuationPage(detailBlock, index + 2, totalPages))
    .join("");

  dom.hioExportPages.innerHTML = [
    renderHioFirstPage(incident, involvedPeople, detailPages[0], 1, totalPages),
    continuationPages,
    renderHioPhotoPage(incident, totalPages, totalPages),
  ].join("");

  dom.hioExportPages.setAttribute("aria-hidden", "false");
}

function clearHioPrintMode() {
  document.body.classList.remove("print-hio-pdf");
  dom.hioExportPages.setAttribute("aria-hidden", "true");
  document.title = previousDocumentTitle;
}

function exportSelectedIncidentAsPdf() {
  const incident = getSelectedIncident();
  if (!incident) {
    return;
  }

  renderHioExport(incident);
  previousDocumentTitle = document.title;
  document.title = `${incident.incidentId} HIO Incident Report`;
  document.body.classList.add("print-hio-pdf");
  window.setTimeout(() => {
    window.print();
  }, 50);
}

async function copySummary() {
  const incident = getSelectedIncident();
  if (!incident) {
    return;
  }

  const assignedWorkers = getAssignedWorkers(incident.assignedPersonnelIds);
  const summary = [
    `Incident ${incident.incidentId}`,
    makeSummary(incident),
    "",
    `Assigned workers: ${assignedWorkers.length > 0 ? assignedWorkers.map((worker) => worker.fullName).join(", ") : "None"}`,
    `Incident photos: ${incident.incidentPhotos.length > 0 ? incident.incidentPhotos.map((photo) => photo.name).join(", ") : "None"}`,
    `Hazards: ${incident.hazards || "Not recorded"}`,
    `Immediate actions: ${incident.initialActions || "Not recorded"}`,
    `Response actions: ${incident.responseActions || "Not recorded"}`,
    `Recovery actions: ${incident.recoveryActions || "Not recorded"}`,
    "",
    "Radio operator log:",
    ...(incident.radioLogs.length > 0
      ? incident.radioLogs.map(
          (entry) =>
            `${formatDateTime(entry.timestamp)} | ${entry.unit} | ${entry.movement} | ${entry.location || "Location not recorded"}${
              entry.notes ? ` | ${entry.notes}` : ""
            }`
        )
      : ["No radio movement entries recorded."]),
  ].join("\n");

  try {
    await navigator.clipboard.writeText(summary);
    const previous = dom.copySummaryButton.textContent;
    dom.copySummaryButton.textContent = "Copied";
    window.setTimeout(() => {
      dom.copySummaryButton.textContent = previous;
    }, 1200);
  } catch {
    window.alert("Clipboard access is unavailable in this browser.");
  }
}

function renderPersonnelCrewOptions(selectedValue = "") {
  const field = dom.personnelForm.elements.namedItem("crew");
  if (!(field instanceof HTMLSelectElement)) {
    return;
  }

  const normalizedValue = singleLine(selectedValue);
  const options = [...CREW_OPTIONS];

  if (normalizedValue && !options.includes(normalizedValue)) {
    options.unshift(normalizedValue);
  }

  field.innerHTML = [
    '<option value="">Select crew / team</option>',
    ...options.map((option) => `<option value="${escapeHtml(option)}">${escapeHtml(option)}</option>`),
  ].join("");
  field.value = normalizedValue;
}

function setWorkerFormValues(worker, statusMessage = "") {
  renderPersonnelCrewOptions(worker.crew);

  Array.from(dom.personnelForm.elements).forEach((field) => {
    if (!field.name) {
      return;
    }

    field.value = worker[field.name] ?? "";
  });

  dom.personnelWorkspaceTitle.textContent = state.editingWorkerId ? `Editing ${worker.fullName || "Worker"}` : "Add Worker";
  dom.personnelFormStatus.textContent =
    statusMessage ||
    (state.editingWorkerId
      ? `Editing ${worker.fullName || "worker"} in the roster.`
      : "Add workers to the roster so they can be assigned to incidents.");
}

function getWorkerFormValues() {
  const worker = {};

  Array.from(dom.personnelForm.elements).forEach((field) => {
    if (!field.name) {
      return;
    }

    worker[field.name] = field.value.trim();
  });

  return normalizeWorker({
    id: state.editingWorkerId || createWorkerId(),
    ...worker,
  });
}

function resetPersonnelForm(statusMessage = "") {
  state.editingWorkerId = null;
  setWorkerFormValues(defaultWorker(), statusMessage);
}

async function saveWorker(worker) {
  const wasCloudBacked = state.storageMode === "cloud";
  const editingIndex = state.editingWorkerId ? state.personnel.findIndex((item) => item.id === state.editingWorkerId) : -1;

  if (editingIndex >= 0) {
    state.personnel[editingIndex] = worker;
  } else {
    state.personnel.push(worker);
  }

  state.personnel = [...state.personnel].sort((a, b) => a.fullName.localeCompare(b.fullName));
  state.editingWorkerId = worker.id;
  const personnelSynced = await persistPersonnel();
  const assignmentsSynced = await syncIncidentPersonnelAssignments();
  renderMetrics();
  renderPersonnelList();
  renderPersonnelPicker();
  renderDraftSidebar();
  renderIncidentList();
  renderIncidentDetail();
  renderReporting();
  setWorkerFormValues(
    worker,
    personnelSynced && assignmentsSynced
      ? `Saved ${worker.fullName}.`
      : wasCloudBacked
        ? `Saved ${worker.fullName} locally. Cloud sync failed.`
        : `${worker.fullName} could not be saved in browser storage.`
  );
}

function editWorker(workerId) {
  const worker = getWorkerById(workerId);
  if (!worker) {
    return;
  }

  state.editingWorkerId = worker.id;
  setWorkerFormValues(worker);
  navigateToView("personnel");
}

async function removeWorker(workerId) {
  const worker = getWorkerById(workerId);
  if (!worker) {
    return;
  }
  const wasCloudBacked = state.storageMode === "cloud";

  if (!window.confirm(`Remove ${worker.fullName} from the personnel roster?`)) {
    return;
  }

  state.personnel = state.personnel.filter((item) => item.id !== workerId);
  const personnelSynced = await persistPersonnel();
  const assignmentsSynced = await syncIncidentPersonnelAssignments();

  if (state.editingWorkerId === workerId) {
    resetPersonnelForm(
      personnelSynced && assignmentsSynced
        ? "Worker removed from the roster."
        : wasCloudBacked
          ? "Worker removed locally. Cloud sync failed."
          : "Worker could not be removed from browser storage."
    );
  }

  renderMetrics();
  renderPersonnelList();
  renderPersonnelPicker();
  renderDraftSidebar();
  renderIncidentList();
  renderIncidentDetail();
  renderReporting();
}

function renderPersonnelList() {
  const query = dom.personnelSearch.value.trim().toLowerCase();
  const workers = [...state.personnel]
    .sort((a, b) => a.fullName.localeCompare(b.fullName))
    .filter((worker) => {
      if (!query) {
        return true;
      }

      const haystack = [worker.fullName, worker.role, worker.crew, worker.employeeId].join(" ").toLowerCase();
      return haystack.includes(query);
    });

  dom.personnelCountLabel.textContent = `${workers.length} worker${workers.length === 1 ? "" : "s"} ${
    query ? "match the search" : "in roster"
  }`;
  dom.personnelList.innerHTML = "";

  if (workers.length === 0) {
    dom.personnelList.innerHTML = `<div class="empty-state">${query ? "No workers match the search." : "No workers in the roster yet."}</div>`;
    return;
  }

  workers.forEach((worker) => {
    const fragment = dom.personnelTemplate.content.cloneNode(true);
    const status = fragment.querySelector(".person-status");
    const usageCount = state.incidents.filter((incident) => incident.assignedPersonnelIds.includes(worker.id)).length;
    const tone = workerStatusTone(worker.status);

    fragment.querySelector(".person-name").textContent = worker.fullName;
    fragment.querySelector(".person-role").textContent = worker.role;
    status.textContent = worker.status;
    status.style.background = tone.background;
    status.style.color = tone.color;
    fragment.querySelector(".person-employee").textContent = worker.employeeId || "No employee ID";
    fragment.querySelector(".person-crew").textContent = worker.crew || "No crew";
    fragment.querySelector(".person-phone").textContent = worker.phone || "No contact";
    fragment.querySelector(".person-notes").textContent =
      worker.certifications || `${usageCount} incident${usageCount === 1 ? "" : "s"} currently reference this worker.`;

    fragment.querySelector('[data-person-action="edit"]').addEventListener("click", () => {
      editWorker(worker.id);
    });

    fragment.querySelector('[data-person-action="remove"]').addEventListener("click", () => {
      removeWorker(worker.id);
    });

    dom.personnelList.append(fragment);
  });
}

function wireEvents() {
  dom.navButtons.forEach((button) => {
    button.addEventListener("click", () => {
      navigateToView(button.dataset.viewTarget);
    });
  });

  dom.startNewIncidentButton.addEventListener("click", () => {
    startNewIncident("Blank incident form ready.");
    navigateToView("create");
  });

  dom.createFromRegisterButton.addEventListener("click", () => {
    startNewIncident("Blank incident form ready.");
    navigateToView("create");
  });

  dom.openRegisterButton.addEventListener("click", () => {
    navigateToView("incidents");
  });

  dom.copyReportingSummaryButton.addEventListener("click", copyReportingSummary);
  dom.exportReportingCsvButton.addEventListener("click", exportReportingCsv);

  dom.openCreateFromPersonnelButton.addEventListener("click", () => {
    navigateToView("create");
  });

  dom.managePersonnelButton.addEventListener("click", () => {
    togglePersonnelPicker(false);
    navigateToView("personnel");
  });

  dom.resetFormButton.addEventListener("click", () => {
    const incident = getEditingIncident();
    if (incident) {
      setIncidentFormValues(incident, `Reset form to saved incident ${incident.incidentId}.`);
      return;
    }

    startNewIncident("Blank incident form restored.");
  });

  dom.incidentForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    await saveIncident(getIncidentFormValues());
  });

  dom.incidentForm.addEventListener("input", () => {
    renderDraftSidebar();
  });

  dom.incidentForm.addEventListener("change", () => {
    renderDraftSidebar();
  });

  dom.incidentSearch.addEventListener("input", () => {
    renderIncidentList();
  });

  dom.editSelectedIncidentButton.addEventListener("click", () => {
    const incident = getSelectedIncident();
    if (!incident) {
      return;
    }

    loadIncidentIntoForm(incident.incidentId, true);
  });

  dom.exportJsonButton.addEventListener("click", exportSelectedIncident);
  dom.copySummaryButton.addEventListener("click", copySummary);
  dom.printReportButton.addEventListener("click", exportSelectedIncidentAsPdf);

  dom.personnelPickerToggle.addEventListener("click", () => {
    togglePersonnelPicker();
  });

  dom.personnelPickerOptions.addEventListener("change", (event) => {
    if (!(event.target instanceof HTMLInputElement) || !event.target.matches("[data-picker-checkbox]")) {
      return;
    }

    state.selectedPersonnelIds = [...dom.personnelPickerOptions.querySelectorAll("[data-picker-checkbox]:checked")].map(
      (input) => input.value
    );
    renderPersonnelPicker();
    renderDraftSidebar();
  });

  document.addEventListener("click", (event) => {
    if (dom.personnelPickerPanel.classList.contains("hidden")) {
      return;
    }

    if (event.target instanceof Element && event.target.closest(".picker-field")) {
      return;
    }

    togglePersonnelPicker(false);
  });

  dom.incidentPhotoUpload.addEventListener("change", async () => {
    if (!dom.incidentPhotoUpload.files || dom.incidentPhotoUpload.files.length === 0) {
      return;
    }

    await handleIncidentPhotoUpload(dom.incidentPhotoUpload.files);
  });

  dom.incidentPhotoTriggerButton.addEventListener("click", () => {
    dom.incidentPhotoUpload.click();
  });

  dom.incidentPhotoList.addEventListener("click", (event) => {
    if (!(event.target instanceof HTMLElement)) {
      return;
    }

    const removeButton = event.target.closest("[data-photo-remove]");
    if (!(removeButton instanceof HTMLElement)) {
      return;
    }

    const photoId = removeButton.dataset.photoRemove;
    const removedPhoto = state.currentIncidentPhotos.find((photo) => photo.id === photoId);
    state.currentIncidentPhotos = state.currentIncidentPhotos.filter((photo) => photo.id !== photoId);
    renderIncidentPhotoList();
    renderDraftSidebar();
    dom.incidentPhotoStatus.textContent = removedPhoto
      ? `Removed ${removedPhoto.name} from this incident draft.`
      : "Incident photo removed.";
  });

  dom.saveRadioLogButton.addEventListener("click", () => {
    const entry = getRadioLogDraft();

    if (!entry.timestamp || !entry.unit) {
      window.alert("Add at least a timestamp and vehicle/unit for the radio log entry.");
      return;
    }

    const index = state.currentRadioLogs.findIndex((log) => log.id === entry.id);
    let statusMessage = "";

    if (index >= 0) {
      state.currentRadioLogs[index] = entry;
      statusMessage = `Updated radio entry for ${entry.unit}.`;
    } else {
      state.currentRadioLogs.push(entry);
      statusMessage = `Added radio entry for ${entry.unit}.`;
    }

    state.currentRadioLogs = normalizeRadioLogs(state.currentRadioLogs);
    clearRadioLogDraft(entry.timestamp, statusMessage);
    renderRadioLogList();
    renderDraftSidebar();
  });

  dom.cancelRadioEditButton.addEventListener("click", () => {
    clearRadioLogDraft(currentReportedAtValue());
  });

  dom.personnelForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    await saveWorker(getWorkerFormValues());
  });

  dom.resetPersonnelButton.addEventListener("click", () => {
    resetPersonnelForm("Worker form reset.");
  });

  dom.personnelSearch.addEventListener("input", () => {
    renderPersonnelList();
  });

  dom.settingsForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    await saveSettings(getSettingsFormValues());
  });

  dom.reportingForm.addEventListener("submit", (event) => {
    event.preventDefault();
    state.reportingFilters = getReportingFormValues();
    renderReporting("Incident report generated using the selected filters.");
  });

  dom.reportingForm.addEventListener("change", (event) => {
    if (event.target === dom.reportingPreset) {
      return;
    }

    state.reportingSnapshot = null;
    setReportingActionState(false);
    dom.reportingFormStatus.textContent = "Filters updated. Generate the report to refresh the dashboard.";
  });

  dom.reportingPreset.addEventListener("change", () => {
    state.reportingSnapshot = null;
    setReportingActionState(false);
    const nextFilters = getReportingFormValues();

    if (nextFilters.preset === "custom") {
      setReportingFormValues(
        {
          ...nextFilters,
          dateFrom: dom.reportingDateFrom.value,
          dateTo: dom.reportingDateTo.value,
        },
        "Custom range selected. Set the date window, then generate the report."
      );
      return;
    }

    setReportingFormValues(nextFilters, `${reportPresetLabel(nextFilters.preset)} loaded. Generate the report to refresh the dashboard.`);
  });

  [dom.reportingDateFrom, dom.reportingDateTo].forEach((input) => {
    input.addEventListener("input", () => {
      state.reportingSnapshot = null;
      setReportingActionState(false);
      if (dom.reportingPreset.value !== "custom") {
        dom.reportingPreset.value = "custom";
      }

      dom.reportingDateFrom.disabled = false;
      dom.reportingDateTo.disabled = false;
      dom.reportingFormStatus.textContent = "Custom date range selected. Generate the report to apply it.";
    });
  });

  dom.resetReportingButton.addEventListener("click", () => {
    state.reportingFilters = defaultReportingFilters();
    renderReporting("Reporting filters reset to the default last 30 day view.");
  });

  dom.reportingIncidents.addEventListener("click", (event) => {
    if (!(event.target instanceof HTMLElement)) {
      return;
    }

    const openButton = event.target.closest("[data-report-open]");
    if (!(openButton instanceof HTMLElement) || !openButton.dataset.reportOpen) {
      return;
    }

    state.selectedIncidentId = openButton.dataset.reportOpen;
    navigateToView("incidents");
  });

  dom.resetSettingsButton.addEventListener("click", () => {
    setSettingsFormValues(defaultSettings(), "Default unit list loaded. Save settings to apply it.");
  });

  window.addEventListener("hashchange", () => {
    applyView(resolveViewFromHash());
  });

  window.addEventListener("afterprint", clearHioPrintMode);
  window.addEventListener("focus", () => {
    if (!document.body.classList.contains("print-hio-pdf")) {
      return;
    }

    window.setTimeout(clearHioPrintMode, 200);
  });
}

async function init() {
  await loadAppData();
  await seedSettingsIfEmpty();
  await seedPersonnelIfEmpty();
  await seedIncidentsIfEmpty();
  await syncIncidentPersonnelAssignments();

  state.selectedIncidentId = sortIncidents(state.incidents)[0]?.incidentId || null;
  resetPersonnelForm();
  startNewIncident();
  renderMetrics();
  renderPersonnelList();
  renderSettingsChannelList();
  setSettingsFormValues(state.settings);
  state.reportingFilters = defaultReportingFilters();
  renderReportingIncidentTypeOptions();
  setReportingActionState(false);
  setReportingFormValues(state.reportingFilters);
  renderRadioUnitOptions();
  renderRadioChannelOptions();
  renderIncidentList();
  renderIncidentDetail();
  renderReporting();
  wireEvents();
  applyView(resolveViewFromHash());
}

init().catch((error) => {
  console.error("Application initialization failed.", error);
  setStorageMode("local-fallback", "Local fallback active: startup error");
});
