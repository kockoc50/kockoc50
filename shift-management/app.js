"use strict";

/* ============ קבועים ============ */

const SHIFTS = [
  { id: "morning", name: "משמרת בוקר", start: "06:00", end: "14:00", hours: 8 },
  { id: "evening", name: "משמרת צהריים", start: "14:00", end: "22:00", hours: 8 },
  { id: "night", name: "משמרת לילה", start: "22:00", end: "06:00", hours: 8 },
];

const DAY_NAMES = ["ראשון", "שני", "שלישי", "רביעי", "חמישי", "שישי", "שבת"];

const STORAGE_DRIVERS = "stm_drivers_v1";
const STORAGE_ASSIGNMENTS = "stm_assignments_v1";
const STORAGE_SEEDED = "stm_seeded_v1";

const MIN_REST_HOURS = 11; // מנוחה מינימלית מומלצת בין משמרות

/* ============ מצב ============ */

let drivers = [];
let assignments = [];
let currentWeekStart = getWeekStart(new Date());
let assignContext = { date: null, shiftId: null };
let editingDriverId = null;

/* ============ אחסון ============ */

function loadDrivers() {
  try {
    drivers = JSON.parse(localStorage.getItem(STORAGE_DRIVERS)) || [];
  } catch (e) {
    drivers = [];
  }
}

function saveDrivers() {
  localStorage.setItem(STORAGE_DRIVERS, JSON.stringify(drivers));
}

function loadAssignments() {
  try {
    assignments = JSON.parse(localStorage.getItem(STORAGE_ASSIGNMENTS)) || [];
  } catch (e) {
    assignments = [];
  }
}

function saveAssignments() {
  localStorage.setItem(STORAGE_ASSIGNMENTS, JSON.stringify(assignments));
}

function seedIfNeeded() {
  if (localStorage.getItem(STORAGE_SEEDED)) return;
  if (drivers.length === 0) {
    drivers = [
      { id: uid(), name: "יוסי כהן", phone: "050-1234567", license: "12345678", truck: "12-345-67", active: true },
      { id: uid(), name: "דוד לוי", phone: "052-2345678", license: "23456789", truck: "34-567-89", active: true },
      { id: uid(), name: "משה אברהם", phone: "054-3456789", license: "34567890", truck: "56-789-01", active: true },
    ];
    saveDrivers();
  }
  localStorage.setItem(STORAGE_SEEDED, "1");
}

/* ============ עזרים ============ */

function uid() {
  return "id-" + Date.now().toString(36) + "-" + Math.random().toString(36).slice(2, 8);
}

function getWeekStart(date) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  const day = d.getDay(); // 0 = ראשון
  d.setDate(d.getDate() - day);
  return d;
}

function addDays(date, n) {
  const d = new Date(date);
  d.setDate(d.getDate() + n);
  return d;
}

function toISODate(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function formatDateHeb(date) {
  const d = String(date.getDate()).padStart(2, "0");
  const m = String(date.getMonth() + 1).padStart(2, "0");
  return `${d}/${m}`;
}

function getDriver(id) {
  return drivers.find((d) => d.id === id);
}

function getShift(id) {
  return SHIFTS.find((s) => s.id === id);
}

function toast(msg) {
  const el = document.getElementById("toast");
  el.textContent = msg;
  el.classList.remove("hidden");
  clearTimeout(toast._t);
  toast._t = setTimeout(() => el.classList.add("hidden"), 2500);
}

/* ============ טאבים ============ */

function initTabs() {
  document.querySelectorAll(".tab-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".tab-btn").forEach((b) => b.classList.remove("active"));
      document.querySelectorAll(".tab-panel").forEach((p) => p.classList.remove("active"));
      btn.classList.add("active");
      document.getElementById("tab-" + btn.dataset.tab).classList.add("active");
      if (btn.dataset.tab === "summary") renderSummary();
    });
  });
}

/* ============ לוח משמרות ============ */

function renderWeekLabel() {
  const end = addDays(currentWeekStart, 6);
  document.getElementById("weekLabel").textContent =
    `${formatDateHeb(currentWeekStart)} – ${formatDateHeb(end)} (${currentWeekStart.getFullYear()})`;
}

function renderScheduleTable() {
  renderWeekLabel();

  const headRow = document.getElementById("scheduleHeadRow");
  headRow.innerHTML = "<th>משמרת</th>";
  for (let i = 0; i < 7; i++) {
    const date = addDays(currentWeekStart, i);
    const th = document.createElement("th");
    th.innerHTML = `${DAY_NAMES[i]}<br><small>${formatDateHeb(date)}</small>`;
    headRow.appendChild(th);
  }

  const body = document.getElementById("scheduleBody");
  body.innerHTML = "";

  SHIFTS.forEach((shift) => {
    const tr = document.createElement("tr");
    tr.dataset.shift = shift.id;

    const labelTd = document.createElement("td");
    labelTd.className = "shift-row-label";
    labelTd.innerHTML = `${shift.name}<small>${shift.start}–${shift.end}</small>`;
    tr.appendChild(labelTd);

    for (let i = 0; i < 7; i++) {
      const date = addDays(currentWeekStart, i);
      const iso = toISODate(date);
      const td = document.createElement("td");
      td.className = "shift-cell";
      td.dataset.date = iso;
      td.dataset.shift = shift.id;

      const dayAssignments = assignments.filter((a) => a.date === iso && a.shiftId === shift.id);

      if (dayAssignments.length === 0) {
        td.innerHTML = `<span class="empty-cell-hint">+</span>`;
      } else {
        dayAssignments.forEach((a) => {
          const driver = getDriver(a.driverId);
          const chip = document.createElement("div");
          chip.className = "assignment-chip";
          const name = driver ? driver.name : "נהג לא ידוע";
          const truck = a.truck || (driver ? driver.truck : "");
          chip.innerHTML = `${name}${truck ? `<span class="truck"> · ${truck}</span>` : ""}`;
          td.appendChild(chip);
        });
      }

      td.addEventListener("click", () => openAssignModal(iso, shift.id));
      tr.appendChild(td);
    }

    body.appendChild(tr);
  });
}

function changeWeek(deltaWeeks) {
  currentWeekStart = addDays(currentWeekStart, deltaWeeks * 7);
  renderScheduleTable();
}

/* ============ בדיקת התנגשויות ומנוחה ============ */

function checkConflicts(driverId, date, shiftId) {
  const warnings = [];
  const sameDayOther = assignments.filter(
    (a) => a.driverId === driverId && a.date === date && a.shiftId !== shiftId
  );
  if (sameDayOther.length > 0) {
    const names = sameDayOther.map((a) => getShift(a.shiftId).name).join(", ");
    warnings.push(`הנהג כבר משובץ באותו יום למשמרת: ${names}. ייתכן חוסר מנוחה/חפיפה.`);
  }

  const alreadyThisShift = assignments.some(
    (a) => a.driverId === driverId && a.date === date && a.shiftId === shiftId
  );
  if (alreadyThisShift) {
    warnings.push("הנהג כבר משובץ למשמרת הזו.");
  }

  const prevDateISO = toISODate(addDays(new Date(date), -1));
  const nextDateISO = toISODate(addDays(new Date(date), 1));

  if (shiftId === "morning") {
    const hadNightBefore = assignments.some(
      (a) => a.driverId === driverId && a.date === prevDateISO && a.shiftId === "night"
    );
    if (hadNightBefore) {
      warnings.push(`הנהג סיים משמרת לילה בבוקר זה — פחות מ-${MIN_REST_HOURS} שעות מנוחה.`);
    }
  }

  if (shiftId === "night") {
    const hasMorningAfter = assignments.some(
      (a) => a.driverId === driverId && a.date === nextDateISO && a.shiftId === "morning"
    );
    if (hasMorningAfter) {
      warnings.push(`לנהג יש משמרת בוקר למחרת — פחות מ-${MIN_REST_HOURS} שעות מנוחה.`);
    }
  }

  return warnings;
}

/* ============ מודאל שיבוץ ============ */

function openAssignModal(date, shiftId) {
  assignContext = { date, shiftId };
  const shift = getShift(shiftId);
  const d = new Date(date);
  const dayName = DAY_NAMES[d.getDay()];

  document.getElementById("assignModalTitle").textContent =
    `שיבוץ – ${shift.name} (${shift.start}–${shift.end}) · יום ${dayName} ${formatDateHeb(d)}`;

  const select = document.getElementById("assignDriverSelect");
  select.innerHTML = "";
  const activeDrivers = drivers.filter((dr) => dr.active);
  if (activeDrivers.length === 0) {
    select.innerHTML = `<option value="">אין נהגים פעילים — הוסיפו נהג בלשונית "נהגים"</option>`;
  } else {
    activeDrivers.forEach((dr) => {
      const opt = document.createElement("option");
      opt.value = dr.id;
      opt.textContent = dr.name;
      select.appendChild(opt);
    });
  }

  document.getElementById("assignTruckInput").value = "";
  document.getElementById("assignNoteInput").value = "";
  document.getElementById("assignWarning").classList.add("hidden");

  renderExistingAssignments();
  document.getElementById("assignModal").classList.remove("hidden");
}

function renderExistingAssignments() {
  const container = document.getElementById("existingAssignments");
  const list = assignments.filter(
    (a) => a.date === assignContext.date && a.shiftId === assignContext.shiftId
  );
  if (list.length === 0) {
    container.innerHTML = `<p class="hint">אין שיבוצים עדיין למשמרת זו.</p>`;
    return;
  }
  container.innerHTML = "<strong>שיבוצים קיימים:</strong>";
  list.forEach((a) => {
    const driver = getDriver(a.driverId);
    const row = document.createElement("div");
    row.className = "existing-assignment-row";
    row.innerHTML = `<span>${driver ? driver.name : "נהג לא ידוע"}${a.truck ? " · " + a.truck : ""}</span>`;
    const removeBtn = document.createElement("button");
    removeBtn.className = "danger";
    removeBtn.textContent = "הסר";
    removeBtn.addEventListener("click", () => {
      assignments = assignments.filter((x) => x.id !== a.id);
      saveAssignments();
      renderScheduleTable();
      renderExistingAssignments();
      toast("השיבוץ הוסר");
    });
    row.appendChild(removeBtn);
    container.appendChild(row);
  });
}

function closeAssignModal() {
  document.getElementById("assignModal").classList.add("hidden");
}

function saveAssignment() {
  const driverId = document.getElementById("assignDriverSelect").value;
  if (!driverId) {
    toast("יש לבחור נהג");
    return;
  }
  const truck = document.getElementById("assignTruckInput").value.trim();
  const note = document.getElementById("assignNoteInput").value.trim();
  const { date, shiftId } = assignContext;

  const warnings = checkConflicts(driverId, date, shiftId);
  const warnBox = document.getElementById("assignWarning");
  if (warnings.length > 0) {
    warnBox.innerHTML = "⚠ " + warnings.join("<br>⚠ ");
    warnBox.classList.remove("hidden");
    if (warnings.some((w) => w.includes("כבר משובץ למשמרת הזו"))) {
      return; // מניעת כפילות מוחלטת
    }
  } else {
    warnBox.classList.add("hidden");
  }

  assignments.push({
    id: uid(),
    date,
    shiftId,
    driverId,
    truck,
    note,
  });
  saveAssignments();
  renderScheduleTable();
  renderExistingAssignments();
  toast("השיבוץ נשמר");
}

/* ============ נהגים ============ */

function renderDriversTable() {
  const body = document.getElementById("driversBody");
  body.innerHTML = "";
  if (drivers.length === 0) {
    body.innerHTML = `<tr><td colspan="6" class="hint">אין נהגים עדיין. לחצו על "נהג חדש" כדי להוסיף.</td></tr>`;
    return;
  }
  drivers.forEach((dr) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${dr.name}</td>
      <td>${dr.phone || ""}</td>
      <td>${dr.license || ""}</td>
      <td>${dr.truck || ""}</td>
      <td><span class="badge ${dr.active ? "active" : "inactive"}">${dr.active ? "פעיל" : "לא פעיל"}</span></td>
      <td><button class="secondary edit-driver-btn">עריכה</button></td>
    `;
    tr.querySelector(".edit-driver-btn").addEventListener("click", () => openDriverModal(dr.id));
    body.appendChild(tr);
  });
}

function openDriverModal(driverId) {
  editingDriverId = driverId || null;
  const isEdit = !!driverId;
  document.getElementById("driverModalTitle").textContent = isEdit ? "עריכת נהג" : "נהג חדש";
  document.getElementById("driverDeleteBtn").classList.toggle("hidden", !isEdit);

  if (isEdit) {
    const dr = getDriver(driverId);
    document.getElementById("driverName").value = dr.name || "";
    document.getElementById("driverPhone").value = dr.phone || "";
    document.getElementById("driverLicense").value = dr.license || "";
    document.getElementById("driverTruck").value = dr.truck || "";
    document.getElementById("driverActive").checked = !!dr.active;
  } else {
    document.getElementById("driverName").value = "";
    document.getElementById("driverPhone").value = "";
    document.getElementById("driverLicense").value = "";
    document.getElementById("driverTruck").value = "";
    document.getElementById("driverActive").checked = true;
  }

  document.getElementById("driverModal").classList.remove("hidden");
}

function closeDriverModal() {
  document.getElementById("driverModal").classList.add("hidden");
}

function saveDriver() {
  const name = document.getElementById("driverName").value.trim();
  if (!name) {
    toast("יש להזין שם נהג");
    return;
  }
  const data = {
    name,
    phone: document.getElementById("driverPhone").value.trim(),
    license: document.getElementById("driverLicense").value.trim(),
    truck: document.getElementById("driverTruck").value.trim(),
    active: document.getElementById("driverActive").checked,
  };

  if (editingDriverId) {
    const dr = getDriver(editingDriverId);
    Object.assign(dr, data);
  } else {
    drivers.push({ id: uid(), ...data });
  }
  saveDrivers();
  renderDriversTable();
  renderScheduleTable();
  closeDriverModal();
  toast("הנהג נשמר");
}

function deleteDriver() {
  if (!editingDriverId) return;
  if (!confirm("למחוק את הנהג? שיבוצים קיימים שלו בלוח יישארו מקושרים ל'נהג לא ידוע'.")) return;
  drivers = drivers.filter((d) => d.id !== editingDriverId);
  saveDrivers();
  renderDriversTable();
  renderScheduleTable();
  closeDriverModal();
  toast("הנהג נמחק");
}

/* ============ סיכום שעות ============ */

function renderSummary() {
  const weekDates = [];
  for (let i = 0; i < 7; i++) weekDates.push(toISODate(addDays(currentWeekStart, i)));

  const weekAssignments = assignments.filter((a) => weekDates.includes(a.date));

  const perDriver = {};
  weekAssignments.forEach((a) => {
    if (!perDriver[a.driverId]) perDriver[a.driverId] = [];
    perDriver[a.driverId].push(a);
  });

  const body = document.getElementById("summaryBody");
  body.innerHTML = "";

  const driverIds = Object.keys(perDriver);
  if (driverIds.length === 0) {
    body.innerHTML = `<tr><td colspan="4" class="hint">אין שיבוצים בשבוע המוצג.</td></tr>`;
    return;
  }

  driverIds.forEach((driverId) => {
    const driver = getDriver(driverId);
    const list = perDriver[driverId];
    const totalHours = list.reduce((sum, a) => sum + getShift(a.shiftId).hours, 0);
    const detail = list
      .map((a) => {
        const d = new Date(a.date);
        return `${DAY_NAMES[d.getDay()]} ${formatDateHeb(d)} – ${getShift(a.shiftId).name}`;
      })
      .join("; ");

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${driver ? driver.name : "נהג לא ידוע"}</td>
      <td>${list.length}</td>
      <td>${totalHours}</td>
      <td style="text-align:right;">${detail}</td>
    `;
    body.appendChild(tr);
  });
}

/* ============ העתקת שבוע ============ */

function copyWeekToNext() {
  const weekDates = [];
  for (let i = 0; i < 7; i++) weekDates.push(toISODate(addDays(currentWeekStart, i)));
  const weekAssignments = assignments.filter((a) => weekDates.includes(a.date));

  if (weekAssignments.length === 0) {
    toast("אין שיבוצים בשבוע הנוכחי להעתקה");
    return;
  }

  let added = 0;
  weekAssignments.forEach((a) => {
    const newDate = toISODate(addDays(new Date(a.date), 7));
    const exists = assignments.some(
      (x) => x.date === newDate && x.shiftId === a.shiftId && x.driverId === a.driverId
    );
    if (!exists) {
      assignments.push({ id: uid(), date: newDate, shiftId: a.shiftId, driverId: a.driverId, truck: a.truck, note: a.note });
      added++;
    }
  });
  saveAssignments();
  toast(`הועתקו ${added} שיבוצים לשבוע הבא`);
}

/* ============ ייצוא CSV ============ */

function exportCSV() {
  const rows = [["תאריך", "יום", "משמרת", "שעות", "נהג", "משאית", "הערה"]];
  for (let i = 0; i < 7; i++) {
    const date = addDays(currentWeekStart, i);
    const iso = toISODate(date);
    SHIFTS.forEach((shift) => {
      const dayAssignments = assignments.filter((a) => a.date === iso && a.shiftId === shift.id);
      if (dayAssignments.length === 0) {
        rows.push([iso, DAY_NAMES[i], shift.name, `${shift.start}-${shift.end}`, "", "", ""]);
      } else {
        dayAssignments.forEach((a) => {
          const driver = getDriver(a.driverId);
          rows.push([
            iso,
            DAY_NAMES[i],
            shift.name,
            `${shift.start}-${shift.end}`,
            driver ? driver.name : "",
            a.truck || (driver ? driver.truck : ""),
            a.note || "",
          ]);
        });
      }
    });
  }

  const csvContent = "﻿" + rows.map((r) => r.map(csvEscape).join(",")).join("\r\n");
  const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `לוח-משמרות-${toISODate(currentWeekStart)}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function csvEscape(value) {
  const str = String(value ?? "");
  if (/[",\n]/.test(str)) {
    return `"${str.replace(/"/g, '""')}"`;
  }
  return str;
}

/* ============ אתחול ============ */

function init() {
  loadDrivers();
  loadAssignments();
  seedIfNeeded();

  initTabs();
  renderScheduleTable();
  renderDriversTable();

  document.getElementById("prevWeek").addEventListener("click", () => changeWeek(-1));
  document.getElementById("nextWeek").addEventListener("click", () => changeWeek(1));
  document.getElementById("todayBtn").addEventListener("click", () => {
    currentWeekStart = getWeekStart(new Date());
    renderScheduleTable();
  });
  document.getElementById("copyWeekBtn").addEventListener("click", copyWeekToNext);
  document.getElementById("exportCsvBtn").addEventListener("click", exportCSV);
  document.getElementById("printBtn").addEventListener("click", () => window.print());

  document.getElementById("assignSaveBtn").addEventListener("click", saveAssignment);
  document.getElementById("assignCloseBtn").addEventListener("click", closeAssignModal);

  document.getElementById("addDriverBtn").addEventListener("click", () => openDriverModal(null));
  document.getElementById("driverSaveBtn").addEventListener("click", saveDriver);
  document.getElementById("driverDeleteBtn").addEventListener("click", deleteDriver);
  document.getElementById("driverCloseBtn").addEventListener("click", closeDriverModal);

  [document.getElementById("assignModal"), document.getElementById("driverModal")].forEach((overlay) => {
    overlay.addEventListener("click", (e) => {
      if (e.target === overlay) overlay.classList.add("hidden");
    });
  });
}

document.addEventListener("DOMContentLoaded", init);
