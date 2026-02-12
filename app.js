const DATA_WORKBOOK_URL = "data/Mobile%20Credentials%20Departments.xlsx";

const state = {
  departments: [],
  timeline: [],
  phases: [],
  health: [],
  summary: null,
  tableSort: {
    key: "acronym",
    dir: "asc"
  },
  meta: {
    source: DATA_WORKBOOK_URL,
    sheetsScanned: 0,
    deptRows: 0,
    timelineRows: 0,
    lastLoadedAt: null
  }
};

const kpiGrid = document.getElementById("kpiGrid");
const departmentBars = document.getElementById("departmentBars");
const deptCoverageNote = document.getElementById("deptCoverageNote");
const phaseBars = document.getElementById("phaseBars");
const timelineList = document.getElementById("timelineList");
const timelineWindow = document.getElementById("timelineWindow");
const healthGrid = document.getElementById("healthGrid");
const forecastGrid = document.getElementById("forecastGrid");
const departmentRows = document.getElementById("departmentRows");
const globalStatus = document.getElementById("globalStatus");
const readinessTableHeaders = document.querySelectorAll("#readinessTable thead th[data-sort]");

function setup() {
  setupTableSorting();
  loadHostedData();
}

async function loadHostedData() {
  try {
    globalStatus.textContent = "Loading latest roadmap data...";

    if (typeof XLSX === "undefined") {
      throw new Error("Workbook parser failed to load");
    }

    const response = await fetch(DATA_WORKBOOK_URL, { cache: "no-store" });
    if (!response.ok) {
      throw new Error(`Unable to load ${DATA_WORKBOOK_URL} (${response.status})`);
    }

    const buffer = await response.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const parsed = parseRoadmapWorkbook(workbook);
    applyParsedData(parsed);
    render();
  } catch (error) {
    state.departments = [];
    state.timeline = [];
    state.phases = [];
    state.health = [];
    state.summary = null;
    render();
    globalStatus.textContent = `Data load error: ${error.message}`;
  }
}

function parseRoadmapWorkbook(workbook) {
  const matrices = workbook.SheetNames.map((sheetName) => ({
    sheetName,
    matrix: XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "" })
  }));

  const correctedMap = new Map();
  const quarterThemeMap = new Map();
  const deptRows = [];
  const timelineRows = [];

  for (const { sheetName, matrix } of matrices) {
    const normalizedSheet = normalizeHeader(sheetName);

    if (normalizedSheet.includes("correcteddeptnames")) {
      parseCorrectedNames(matrix).forEach((row) => correctedMap.set(row.acronym, row.name));
    }

    if (normalizedSheet.includes("quarterlyrollout")) {
      parseQuarterThemes(matrix).forEach((row) => quarterThemeMap.set(row.quarter, row.theme));
    }

    if (normalizedSheet.includes("deptcnt")) {
      deptRows.push(...parseDeptCnt(matrix));
    }
  }

  const timelineV2 = matrices.find((s) => normalizeHeader(s.sheetName).includes("communicationtimelinev2"));
  const timelineV1 = matrices.find((s) => normalizeHeader(s.sheetName).includes("communicationtimelinev1"));
  const anyTimeline = matrices.find((s) => normalizeHeader(s.sheetName).includes("communicationtimeline"));
  const timelineSource = timelineV2 || timelineV1 || anyTimeline;
  if (timelineSource) {
    timelineRows.push(...parseCommunicationTimeline(timelineSource.matrix));
  }

  if (!deptRows.length && !timelineRows.length) {
    throw new Error("No usable roadmap data found in workbook");
  }

  return {
    deptRows,
    timelineRows,
    correctedMap,
    quarterThemeMap,
    sheetsScanned: workbook.SheetNames.length
  };
}

function parseCorrectedNames(matrix) {
  const header = findHeaderRow(matrix, ["abbreviation", "fulldepartmentname"]);
  if (!header) return [];

  const rows = matrix.slice(header.rowIndex + 1);
  const out = [];

  rows.forEach((row) => {
    const acronym = normalizeAcronym(readCellText(row, header.index.abbreviation));
    const name = readCellText(row, header.index.fulldepartmentname);
    if (!acronym || !name) return;
    out.push({ acronym, name });
  });

  return out;
}

function parseQuarterThemes(matrix) {
  const out = [];
  matrix.forEach((row) => {
    const quarter = normalizeQuarter(readCellText(row, 0));
    const theme = readCellText(row, 1);
    if (!quarter || !theme) return;
    out.push({ quarter, theme });
  });
  return out;
}

function parseDeptCnt(matrix) {
  const header = findHeaderRow(matrix, ["abbreviation", "fulldepartmentname", "headcount"]);
  if (!header) return [];

  const rows = matrix.slice(header.rowIndex + 1);
  const out = [];

  rows.forEach((row) => {
    const acronym = normalizeAcronym(readCellText(row, header.index.abbreviation));
    const name = readCellText(row, header.index.fulldepartmentname);
    const headcount = Math.max(0, Math.round(readCellNumber(row, header.index.headcount)));
    const quarter = normalizeQuarter(readCellText(row, header.index.quarter));
    const conversionRate = normalizePercent(readCellNumber(row, header.index.conversionrate));

    if (!acronym || isSummaryLabel(name)) return;
    out.push({ acronym, name, headcount, quarter, conversionRate });
  });

  return out;
}

function parseCommunicationTimeline(matrix) {
  const header = findHeaderRow(matrix, ["dept", "rolloutdate"]);
  if (!header) return [];

  const rows = matrix.slice(header.rowIndex + 1);
  const out = [];

  rows.forEach((row) => {
    const acronym = normalizeAcronym(readCellText(row, header.index.dept));
    if (!acronym) return;

    out.push({
      acronym,
      name: readCellText(row, header.index.fulldepartmentname),
      quarter: normalizeQuarter(readCellText(row, header.index.qtr)),
      rolloutDate: readCellDate(row, header.index.rolloutdate),
      owner: readCellText(row, header.index.commsteward),
      note: readCellText(row, header.index.note),
      count: Math.max(0, Math.round(readCellNumber(row, header.index.count))),
      conversionRate: normalizePercent(readCellNumber(row, header.index.conversionrate))
    });
  });

  return out;
}

function findHeaderRow(matrix, requiredKeys) {
  const alias = {
    abbreviation: ["abbreviation", "abbr", "dept"],
    fulldepartmentname: ["fulldepartmentname", "departmentname", "fulldepartmentname"],
    headcount: ["headcount", "count"],
    quarter: ["quarter", "qtr"],
    dept: ["dept", "department", "abbreviation"],
    rolloutdate: ["rolloutdate", "rolloutdate", "date"],
    qtr: ["qtr", "quarter"],
    commsteward: ["commsteward", "commsteward", "steward", "owner"],
    note: ["note", "notes"],
    count: ["count", "headcount"],
    conversionrate: ["conversionrate", "conversion", "conversion", "conversionpercent", "digitalbadgeconversion"]
  };

  const limit = Math.min(matrix.length, 30);

  for (let rowIndex = 0; rowIndex < limit; rowIndex += 1) {
    const normalized = (matrix[rowIndex] || []).map((cell) => normalizeHeader(cell));
    const index = {};

    for (const [key, aliases] of Object.entries(alias)) {
      index[key] = findColumnIndex(normalized, aliases);
    }

    if (requiredKeys.every((key) => index[key] >= 0)) {
      return { rowIndex, index };
    }
  }

  return null;
}

function findColumnIndex(headers, aliases) {
  for (let i = 0; i < headers.length; i += 1) {
    const value = headers[i];
    if (!value) continue;
    if (aliases.some((a) => value === normalizeHeader(a) || value.startsWith(normalizeHeader(a)))) {
      return i;
    }
  }
  return -1;
}

function applyParsedData(parsed) {
  const deptMap = new Map();

  parsed.deptRows.forEach((row) => {
    if (!deptMap.has(row.acronym)) deptMap.set(row.acronym, createDeptRecord(row.acronym));
    const dept = deptMap.get(row.acronym);

    dept.name = row.name || dept.name;
    dept.headcount = Math.max(dept.headcount, row.headcount);
    dept.quarter = row.quarter || dept.quarter;
    if (row.conversionRate != null) dept.conversionRate = row.conversionRate;
  });

  parsed.timelineRows.forEach((row) => {
    if (!deptMap.has(row.acronym)) deptMap.set(row.acronym, createDeptRecord(row.acronym));
    const dept = deptMap.get(row.acronym);

    dept.name = row.name || dept.name;
    if (row.count > 0) dept.headcount = Math.max(dept.headcount, row.count);
    dept.quarter = row.quarter || dept.quarter;
    dept.owner = row.owner || dept.owner;
    dept.note = row.note || dept.note;
    if (row.conversionRate != null) dept.conversionRate = row.conversionRate;

    if (row.rolloutDate instanceof Date && (!dept.rolloutDate || row.rolloutDate < dept.rolloutDate)) {
      dept.rolloutDate = row.rolloutDate;
    }
  });

  for (const dept of deptMap.values()) {
    if (!dept.name && parsed.correctedMap.has(dept.acronym)) dept.name = parsed.correctedMap.get(dept.acronym);
    if (!dept.quarter && dept.rolloutDate instanceof Date) dept.quarter = quarterFromDate(dept.rolloutDate);

    dept.milestoneTheme = parsed.quarterThemeMap.get(dept.quarter) || "-";
    dept.status = deriveStatus(dept.rolloutDate, dept.note);
    dept.badgeUsers = dept.conversionRate != null && dept.headcount > 0
      ? Math.round((dept.conversionRate / 100) * dept.headcount)
      : 0;
  }

  const departments = [...deptMap.values()].sort((a, b) => b.headcount - a.headcount || a.acronym.localeCompare(b.acronym));
  const withDates = departments
    .filter((d) => d.rolloutDate instanceof Date)
    .sort((a, b) => a.rolloutDate - b.rolloutDate);
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const upcoming = withDates.filter((d) => d.rolloutDate >= today);
  const source = upcoming.length ? upcoming : withDates;

  const timeline = source
    .slice(0, 12)
    .map((d) => ({
      department: d.acronym,
      date: d.rolloutDate,
      milestone: d.milestoneTheme !== "-" ? d.milestoneTheme : `${d.quarter || "Unspecified"} rollout`,
      status: d.status
    }));

  state.departments = departments;
  state.timeline = timeline;
  state.phases = buildPhaseStats(departments);
  state.health = buildHealthStats(departments);
  state.summary = buildSummary(departments);
  state.meta.sheetsScanned = parsed.sheetsScanned;
  state.meta.deptRows = parsed.deptRows.length;
  state.meta.timelineRows = parsed.timelineRows.length;
  state.meta.lastLoadedAt = new Date();
}

function createDeptRecord(acronym) {
  return {
    acronym,
    name: "",
    headcount: 0,
    badgeUsers: 0,
    conversionRate: null,
    quarter: "",
    rolloutDate: null,
    owner: "",
    note: "",
    milestoneTheme: "-",
    status: "No Data"
  };
}

function buildPhaseStats(departments) {
  const map = new Map();
  departments.forEach((dept) => {
    const quarter = dept.quarter || "Unspecified";
    if (!map.has(quarter)) map.set(quarter, { phase: quarter, total: 0, headcount: 0 });
    const value = map.get(quarter);
    value.total += 1;
    value.headcount += dept.headcount;
  });
  return [...map.values()].sort((a, b) => quarterSortValue(a.phase) - quarterSortValue(b.phase));
}

function buildHealthStats(departments) {
  const counters = { "At Risk": 0, Watch: 0, "On Track": 0, Complete: 0, "No Data": 0 };
  departments.forEach((dept) => {
    counters[dept.status] = (counters[dept.status] || 0) + 1;
  });
  return Object.entries(counters)
    .filter(([, count]) => count > 0)
    .map(([label, count]) => ({ label, count }));
}

function buildSummary(departments) {
  const totalHeadcount = departments.reduce((sum, d) => sum + d.headcount, 0);
  const withDates = departments.filter((d) => d.rolloutDate instanceof Date).length;
  const withConversion = departments.filter((d) => d.conversionRate != null && d.headcount > 0);
  const totalBadgeUsers = departments.reduce((sum, d) => sum + (d.badgeUsers || 0), 0);
  const conversionRate = totalHeadcount > 0 ? (totalBadgeUsers / totalHeadcount) * 100 : 0;

  return {
    departments: departments.length,
    totalHeadcount,
    withDates,
    conversionDeptCount: withConversion.length,
    totalBadgeUsers,
    conversionRate
  };
}

function render() {
  renderKpis();
  renderDepartmentBars();
  renderPhaseBars();
  renderTimeline();
  renderHealth();
  renderForecast();
  renderTable();
  renderStatus();
}

function renderKpis() {
  if (!state.summary) {
    kpiGrid.innerHTML = "";
    return;
  }

  const cards = [
    { label: "Departments", value: formatNumber(state.summary.departments), trend: "Using acronym-keyed department records" },
    { label: "Total Headcount", value: formatNumber(state.summary.totalHeadcount), trend: "Headcount from roadmap sheets" },
    { label: "With Rollout Dates", value: formatNumber(state.summary.withDates), trend: "Departments with a scheduled rollout date" },
    { label: "Digital Badge", value: formatNumber(state.summary.totalBadgeUsers), trend: "Derived from Excel conversion rates" },
    {
      label: "Conversion Rate",
      value: state.summary.totalHeadcount > 0 ? `${state.summary.conversionRate.toFixed(1)}%` : "-",
      trend: `${formatNumber(state.summary.totalBadgeUsers)} / ${formatNumber(state.summary.totalHeadcount)} across all depts`
    }
  ];

  kpiGrid.innerHTML = cards
    .map(
      (card) => `
      <article class="kpi-card">
        <p class="kpi-label">${card.label}</p>
        <p class="kpi-value">${card.value}</p>
        <p class="kpi-trend">${card.trend}</p>
      </article>
    `
    )
    .join("");
}

function renderDepartmentBars() {
  if (!state.departments.length) {
    departmentBars.innerHTML = '<p class="empty-state">Roadmap data is not available yet.</p>';
    deptCoverageNote.textContent = "";
    return;
  }

  const allDepartments = state.departments.slice().sort((a, b) => a.acronym.localeCompare(b.acronym));

  departmentBars.innerHTML = allDepartments
    .map(
      (d) => `
      <div class="coverage-item">
        <div class="coverage-item-head">
          <span class="metric-row-name" title="${escapeHtml(d.acronym)}">${escapeHtml(d.acronym)}</span>
          <span class="coverage-item-value">${(d.conversionRate || 0).toFixed(1)}%</span>
        </div>
        <div class="track">
          <div class="fill" style="width:100%"></div>
          <div class="fill-secondary" style="width:${Math.max(0, Math.min(100, d.conversionRate || 0)).toFixed(1)}%"></div>
        </div>
      </div>
    `
    )
    .join("");

  deptCoverageNote.textContent = state.summary.totalHeadcount > 0
    ? `${state.summary.conversionRate.toFixed(1)}% enterprise conversion`
    : "No conversion-rate data found in workbook";
}

function renderPhaseBars() {
  if (!state.phases.length) {
    phaseBars.innerHTML = '<p class="empty-state">No quarter/phase data found.</p>';
    return;
  }

  const max = Math.max(...state.phases.map((p) => p.total), 1);
  phaseBars.innerHTML = state.phases
    .map(
      (p) => `
      <div class="metric-row">
        <span class="metric-row-name">${escapeHtml(p.phase)}</span>
        <div class="track"><div class="fill" style="width:${((p.total / max) * 100).toFixed(1)}%"></div></div>
        <span>${p.total} depts</span>
      </div>
    `
    )
    .join("");
}

function renderTimeline() {
  if (!state.timeline.length) {
    timelineList.innerHTML = '<p class="empty-state">No rollout dates found.</p>';
    timelineWindow.textContent = "";
    return;
  }

  const minDate = state.timeline[0].date;
  const maxDate = state.timeline[state.timeline.length - 1].date;
  timelineWindow.textContent = `${formatDate(minDate)} to ${formatDate(maxDate)}`;

  timelineList.innerHTML = state.timeline
    .map(
      (item) => `
      <li>
        <strong>${formatDate(item.date)}</strong> - ${escapeHtml(item.department)}
        <span class="badge ${statusClass(item.status)}">${escapeHtml(item.status)}</span>
      </li>
    `
    )
    .join("");
}

function renderHealth() {
  if (!state.health.length) {
    healthGrid.innerHTML = '<p class="empty-state">No health distribution available.</p>';
    return;
  }

  healthGrid.innerHTML = state.health
    .map(
      (h) => `
      <div class="health-row">
        <span>${escapeHtml(h.label)}</span>
        <strong>${h.count}</strong>
      </div>
    `
    )
    .join("");
}

function renderForecast() {
  if (!state.departments.length) {
    forecastGrid.innerHTML = '<p class="empty-state">No forecast available.</p>';
    return;
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const next30 = new Date(today.getTime() + (30 * 86400000));
  const next90 = new Date(today.getTime() + (90 * 86400000));

  const scheduled = state.departments.filter((d) => d.rolloutDate instanceof Date);
  const byDate = (endDate) =>
    scheduled.filter((d) => d.rolloutDate >= today && d.rolloutDate <= endDate);

  const next30Rows = byDate(next30);
  const next90Rows = byDate(next90);
  const unscheduled = state.departments.length - scheduled.length;
  const next30Headcount = next30Rows.reduce((sum, d) => sum + d.headcount, 0);

  const rows = [
    { label: "Next 30 Days", value: `${next30Rows.length} depts` },
    { label: "30-Day Headcount", value: formatNumber(next30Headcount) },
    { label: "Next 90 Days", value: `${next90Rows.length} depts` },
    { label: "Unscheduled", value: `${unscheduled} depts` }
  ];

  forecastGrid.innerHTML = rows
    .map(
      (row) => `
      <div class="health-row">
        <span>${escapeHtml(row.label)}</span>
        <strong>${escapeHtml(row.value)}</strong>
      </div>
    `
    )
    .join("");
}

function renderTable() {
  updateSortHeaderIndicators();

  if (!state.departments.length) {
    departmentRows.innerHTML = '<tr><td colspan="7" class="empty-state">No department data available.</td></tr>';
    return;
  }

  const sorted = sortDepartments(state.departments, state.tableSort);
  departmentRows.innerHTML = sorted
    .map(
      (d) => `
      <tr>
        <td>${escapeHtml(d.acronym)}</td>
        <td>${formatNumber(d.headcount)}</td>
        <td>${d.conversionRate != null ? formatNumber(d.badgeUsers) : "-"}</td>
        <td>${d.conversionRate != null ? `${d.conversionRate.toFixed(1)}%` : "-"}</td>
        <td>${escapeHtml(d.quarter || "-")}</td>
        <td>${d.rolloutDate ? formatDate(d.rolloutDate) : "-"}</td>
        <td><span class="status-text ${statusClass(d.status)}">${escapeHtml(d.status)}</span></td>
      </tr>
    `
    )
    .join("");
}

function setupTableSorting() {
  readinessTableHeaders.forEach((header) => {
    header.style.cursor = "pointer";
    header.addEventListener("click", () => {
      const key = header.dataset.sort;
      if (!key) return;

      if (state.tableSort.key === key) {
        state.tableSort.dir = state.tableSort.dir === "asc" ? "desc" : "asc";
      } else {
        state.tableSort.key = key;
        state.tableSort.dir = "asc";
      }

      renderTable();
    });
  });
}

function updateSortHeaderIndicators() {
  readinessTableHeaders.forEach((header) => {
    const label = header.dataset.label || header.textContent || "";
    if (header.dataset.sort === state.tableSort.key) {
      const arrow = state.tableSort.dir === "asc" ? " ▲" : " ▼";
      header.textContent = `${label}${arrow}`;
    } else {
      header.textContent = label;
    }
  });
}

function sortDepartments(departments, sort) {
  const rows = departments.slice();
  const dir = sort.dir === "asc" ? 1 : -1;

  rows.sort((a, b) => {
    const av = sortableValue(a, sort.key);
    const bv = sortableValue(b, sort.key);

    if (av == null && bv == null) return 0;
    if (av == null) return 1 * dir;
    if (bv == null) return -1 * dir;

    if (typeof av === "number" && typeof bv === "number") {
      return (av - bv) * dir;
    }

    return String(av).localeCompare(String(bv), undefined, { numeric: true, sensitivity: "base" }) * dir;
  });

  return rows;
}

function sortableValue(dept, key) {
  if (key === "rolloutDate") {
    return dept.rolloutDate instanceof Date ? dept.rolloutDate.getTime() : null;
  }
  if (key === "conversionRate") {
    return Number.isFinite(dept.conversionRate) ? dept.conversionRate : null;
  }
  return dept[key];
}

function renderStatus() {
  if (!state.summary) {
    globalStatus.textContent = "Loading latest roadmap data...";
    return;
  }

  const loadedTime = state.meta.lastLoadedAt
    ? new Date(state.meta.lastLoadedAt).toLocaleString("en-US", { month: "short", day: "numeric", hour: "numeric", minute: "2-digit" })
    : "-";

  globalStatus.textContent = `Live data loaded (${state.meta.sheetsScanned} sheets) • Updated ${loadedTime}`;
}

function deriveStatus(rolloutDate, note) {
  const noteText = String(note || "").toLowerCase();
  if (noteText.includes("missed") || noteText.includes("delay") || noteText.includes("overdue") || noteText.includes("risk")) {
    return "At Risk";
  }

  if (!(rolloutDate instanceof Date) || Number.isNaN(rolloutDate.getTime())) {
    return "Watch";
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  if (rolloutDate < today) return "Complete";

  const days = (rolloutDate.getTime() - today.getTime()) / 86400000;
  if (days <= 30) return "Watch";
  return "On Track";
}

function quarterFromDate(date) {
  const month = date.getMonth() + 1;
  if (month <= 3) return "Q1";
  if (month <= 6) return "Q2";
  if (month <= 9) return "Q3";
  return "Q4";
}

function normalizeAcronym(value) {
  const text = String(value || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
  if (!text || text.length > 8) return "";
  return text;
}

function normalizeQuarter(value) {
  const text = String(value || "").trim().toLowerCase();
  if (!text) return "";
  const match = text.match(/q\s*([1-4])/i) || text.match(/quarter\s*([1-4])/i) || text.match(/^([1-4])$/);
  return match ? `Q${match[1]}` : "";
}

function normalizePercent(value) {
  if (!Number.isFinite(value)) return null;
  if (value <= 0) return null;
  const percent = value <= 1 ? value * 100 : value;
  return Math.max(0, Math.min(100, percent));
}

function isSummaryLabel(name) {
  const text = String(name || "").toLowerCase();
  return text.includes("sum of") || text.includes("cummulative") || text.includes("each quarter") || text.includes("planned") || text.includes("remaining") || text.includes("need");
}

function quarterSortValue(value) {
  const match = String(value || "").toUpperCase().match(/^Q([1-4])$/);
  return match ? Number(match[1]) : 99;
}

function readCellText(row, index) {
  if (index == null || index < 0) return "";
  return String(row[index] ?? "").trim();
}

function readCellNumber(row, index) {
  if (index == null || index < 0) return NaN;
  const value = row[index];
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const parsed = Number(String(value ?? "").replace(/[,%\s]/g, ""));
  return Number.isFinite(parsed) ? parsed : NaN;
}

function readCellDate(row, index) {
  if (index == null || index < 0) return null;
  const value = row[index];
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;

  if (typeof value === "number" && Number.isFinite(value) && value > 20000 && value < 70000) {
    const base = new Date(1899, 11, 30, 12, 0, 0, 0);
    const date = new Date(base.getTime() + value * 86400000);
    if (!Number.isNaN(date.getTime())) return date;
  }

  const text = String(value ?? "").trim();
  if (!text) return null;

  const iso = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (iso) {
    const y = Number(iso[1]);
    const m = Number(iso[2]) - 1;
    const d = Number(iso[3]);
    const date = new Date(y, m, d, 12, 0, 0, 0);
    return Number.isNaN(date.getTime()) ? null : date;
  }

  const date = new Date(text);
  return Number.isNaN(date.getTime()) ? null : date;
}

function normalizeHeader(value) {
  return String(value || "")
    .toLowerCase()
    .trim()
    .replace(/[^a-z0-9]/g, "");
}

function formatNumber(value) {
  return new Intl.NumberFormat("en-US").format(Math.round(value || 0));
}

function formatDate(value) {
  return new Date(value).toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
}

function statusClass(status) {
  if (status === "At Risk") return "at-risk";
  if (status === "Watch") return "watch";
  if (status === "Complete") return "complete";
  return "on-track";
}

function escapeHtml(text) {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

setup();
