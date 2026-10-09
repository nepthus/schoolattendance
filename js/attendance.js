/**
 * =====================================================
 * ⚠️ CRITICAL ATTENDANCE LOGIC
 * Do NOT rewrite or refactor automatically.
 * Do NOT touch HTML or Cyrillic text.
 * Edit JS logic only.
 * =====================================================
 */

/* ================= DOM READY ================= */
document.addEventListener("DOMContentLoaded", async () => {
  updateCurrentDate();
  await loadClassRosters();
  initTotals();
  initValidation();
  initNameMultiSelects();
  initLateStudentDropdowns();
  initExportToolbar();
  restoreDraft();
  calculateTotals();

  document.addEventListener("input", () => {
    calculateTotals();
    saveDraft();
  });

  document.addEventListener("change", () => {
    calculateTotals();
    saveDraft();
  });

  window.addEventListener("pagehide", saveDraft);
});

function updateCurrentDate() {
  const dateLine = document.querySelector(".sub-line");
  if (!dateLine) return;

  const today = new Date();
  dateLine.textContent = `${today.getFullYear()} оны ${today.getMonth() + 1} сарын ${today.getDate()}-ны өдөр`;
}

/* ================= REFRESH-SAFE DRAFT ================= */
const DRAFT_STORAGE_KEY = `attendance-draft-v${document.body.dataset.shift || "1"}`;

function formControlValues(row) {
  return [...row.querySelectorAll("input, select, textarea")]
    .filter(control => control.type !== "checkbox")
    .map(control => control.value);
}

function restoreControlValues(row, values = []) {
  [...row.querySelectorAll("input, select, textarea")]
    .filter(control => control.type !== "checkbox")
    .forEach((control, index) => {
      if (values[index] !== undefined) control.value = values[index];
    });
}

function saveDraft() {
  try {
    const draft = {
      attendance: [...document.querySelectorAll(".attendance-table tbody tr")].map(row => ({
        values: [3, 4, 5, 6, 7].map(index => row.cells[index]?.querySelector("input")?.value || ""),
        absentNames: selectedNames(row.cells[8]),
        leaveNames: selectedNames(row.cells[9])
      })),
      elective: [...document.querySelectorAll(".elective-table tbody tr")].map(formControlValues),
      dutyTeacherNotes: document.getElementById("duty-teacher-notes")?.value || "",
      late: [...document.querySelectorAll(".late-students-table tbody tr")].map(row => ({
        values: formControlValues(row),
        studentNames: selectedNames(row.cells[2])
      }))
    };
    localStorage.setItem(DRAFT_STORAGE_KEY, JSON.stringify(draft));
  } catch (error) {
    console.warn("Attendance draft could not be saved.", error);
  }
}

function restoreDraft() {
  try {
    const rawDraft = localStorage.getItem(DRAFT_STORAGE_KEY);
    if (!rawDraft) return;
    const draft = JSON.parse(rawDraft);

    [...document.querySelectorAll(".attendance-table tbody tr")].forEach((row, rowIndex) => {
      const savedRow = draft.attendance?.[rowIndex];
      if (!savedRow) return;
      [3, 4, 5, 6, 7].forEach((cellIndex, valueIndex) => {
        const input = row.cells[cellIndex]?.querySelector("input");
        if (input && savedRow.values?.[valueIndex] !== undefined) input.value = savedRow.values[valueIndex];
      });
      [[8, savedRow.absentNames], [9, savedRow.leaveNames]].forEach(([cellIndex, names]) => {
        const selected = new Set(String(names || "").split(", ").filter(Boolean));
        row.cells[cellIndex]?.querySelectorAll(".name-multiselect input").forEach(input => {
          input.checked = selected.has(input.value);
        });
        const control = row.cells[cellIndex]?.querySelector(".name-multiselect");
        if (control) updateNameMultiSelect(control);
      });
    });

    [...document.querySelectorAll(".elective-table tbody tr")]
      .forEach((row, index) => restoreControlValues(row, draft.elective?.[index]));

    const dutyTeacherNotes = document.getElementById("duty-teacher-notes");
    if (dutyTeacherNotes) dutyTeacherNotes.value = draft.dutyTeacherNotes || "";

    [...document.querySelectorAll(".late-students-table tbody tr")].forEach((row, index) => {
      const savedRow = draft.late?.[index];
      if (!savedRow) return;
      const values = Array.isArray(savedRow) ? savedRow : savedRow.values;
      const classSelect = row.querySelector(".late-class-select");
      if (classSelect && values[0] !== undefined) {
        classSelect.value = values[0];
        classSelect.dispatchEvent(new Event("change"));
      }
      restoreControlValues(row, values);
      const selected = new Set(String(savedRow.studentNames || "").split(", ").filter(Boolean));
      row.cells[2]?.querySelectorAll(".name-multiselect input").forEach(input => {
        input.checked = selected.has(input.value);
      });
      const control = row.cells[2]?.querySelector(".name-multiselect");
      if (control) updateNameMultiSelect(control);
    });
  } catch (error) {
    console.warn("Attendance draft could not be restored.", error);
  }
}

/* ================= STUDENT ROSTER DATA ================= */
let CLASS_ROSTERS = {};

async function loadClassRosters() {
  const errorBox = document.querySelector(".floating-error");

  try {
    if (typeof XLSX === "undefined") {
      throw new Error("Excel library could not be loaded");
    }

    const response = await fetch("students_by_class.xlsx", { cache: "no-store" });
    if (!response.ok) throw new Error(`Workbook request failed (${response.status})`);

    const workbook = XLSX.read(await response.arrayBuffer(), { type: "array" });
    CLASS_ROSTERS = rostersFromWorkbook(workbook);
    if (!Object.keys(CLASS_ROSTERS).length) {
      throw new Error("Workbook contains no readable student list");
    }

    if (errorBox) errorBox.textContent = "";
  } catch (error) {
    console.warn("Student roster workbook could not be loaded.", error);
    CLASS_ROSTERS = await readFallbackRosters();

    if (!Object.keys(CLASS_ROSTERS).length) {
      if (errorBox) {
        errorBox.textContent = `Сурагчийн мэдээлэл уншигдаагүй: ${error.message}`;
      }
      return;
    }

    if (errorBox) {
      errorBox.textContent = "Excel sheet унших боломжгүй тул хадгалагдсан сурагчийн жагсаалтаар ажиллаж байна.";
    }
  }
}

function normaliseGroupName(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/^0+/, "")
    .replace(/^(\d+)([а-яөүе])$/i, "$1-$2");
}

function normaliseRosters(rosters) {
  return Object.fromEntries(Object.entries(rosters)
    .map(([group, students]) => [group.trim(), [...new Set((students || [])
      .map(name => String(name).trim()).filter(Boolean))]])
    .filter(([group, students]) => group && students.length));
}

async function readFallbackRosters() {
  try {
    const response = await fetch("data/students.json", { cache: "no-store" });
    if (!response.ok) return {};
    const data = await response.json();
    return normaliseRosters(data);
  } catch (error) {
    console.warn("Fallback student roster could not be loaded.", error);
    return {};
  }
}

function rostersFromWorkbook(workbook) {
  const rosters = {};

  workbook.SheetNames.forEach(sheetName => {
    if (["Заавар", "Guide", "Instructions"].includes(sheetName)) return;

    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "" });
    const students = rows.slice(1).map(row => {
      const ovog = String(row[2] || "").trim();
      const givenName = String(row[1] || "").trim();
      return [ovog, givenName].filter(Boolean).join(" ");
    });

    const cleanStudents = [...new Set(students.filter(Boolean))];
    if (cleanStudents.length) {
      const normalizedGroup = normaliseGroupName(sheetName);
      rosters[normalizedGroup] = cleanStudents;
    }
  });

  return normaliseRosters(rosters);
}

/* ================= LATE STUDENT DROPDOWNS ================= */
function initLateStudentDropdowns() {
  const classGroups = Object.keys(CLASS_ROSTERS);

  document.querySelectorAll(".late-students-table tbody tr").forEach(row => {
    const classCell = row.cells[1];
    const studentCell = row.cells[2];
    if (!classCell || !studentCell) return;

    const classSelect = document.createElement("select");
    classSelect.className = "late-dropdown late-class-select";
    classSelect.setAttribute("aria-label", "Анги бүлэг сонгох");
    classSelect.innerHTML = '<option value="">Анги сонгох</option>';
    classGroups.forEach(group => classSelect.add(new Option(group, group)));

    const populateStudents = () => {
      const selected = selectedNames(studentCell).split(", ").filter(Boolean);
      const emptyLabel = classSelect.value
        ? "Сурагчийн жагсаалт алга"
        : "Эхлээд анги сонгоно уу";
      createNameMultiSelect(studentCell, selected, classSelect.value, emptyLabel);
    };

    classSelect.addEventListener("change", populateStudents);
    classCell.replaceChildren(classSelect);
    populateStudents();
  });
}

/* ================= STUDENT NAME MULTI-SELECTS ================= */

function initNameMultiSelects() {
  document.querySelectorAll(".attendance-table tbody tr").forEach(row => {
    [8, 9].forEach(index => {
      const cell = row.cells[index];
      if (!cell) return;
      const classGroup = row.cells[2]?.textContent.trim();

      const selected = [...cell.querySelectorAll("option:checked, .name-multiselect input:checked")]
        .map(option => option.value || option.textContent.trim())
        .filter(Boolean);
      createNameMultiSelect(cell, selected, classGroup);
    });
  });

  document.addEventListener("click", event => {
    if (!event.target.closest(".name-multiselect")) {
      document.querySelectorAll(".name-multiselect.is-open")
        .forEach(control => control.classList.remove("is-open"));
    }
  });
}

function createNameMultiSelect(cell, selected = [], classGroup = "", emptyLabel = "Жагсаалт оруулаагүй") {
  cell.replaceChildren();
  const control = document.createElement("div");
  control.className = "name-multiselect";
  control.dataset.emptyLabel = emptyLabel;
  control.innerHTML = `
    <button type="button" class="name-multiselect__toggle" aria-expanded="false">
      <span class="name-multiselect__value"></span><span aria-hidden="true">⌄</span>
    </button>
    <div class="name-multiselect__menu" role="group" aria-label="Сурагч сонгох"></div>`;

  const menu = control.querySelector(".name-multiselect__menu");
  const students = CLASS_ROSTERS[classGroup] || [];
  students.forEach(name => {
    const label = document.createElement("label");
    label.className = "name-multiselect__option";
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.value = name;
    checkbox.checked = selected.includes(name);
    label.append(checkbox, document.createTextNode(displayStudentName(name)));
    menu.appendChild(label);
  });

  if (!students.length) {
    const notice = document.createElement("span");
    notice.className = "name-multiselect__empty";
    notice.textContent = "Сурагчийн жагсаалт алга";
    menu.appendChild(notice);
  }

  const toggle = control.querySelector(".name-multiselect__toggle");
  toggle.addEventListener("click", event => {
    event.stopPropagation();
    const opening = !control.classList.contains("is-open");
    document.querySelectorAll(".name-multiselect.is-open")
      .forEach(item => item.classList.remove("is-open"));
    control.classList.toggle("is-open", opening);
    toggle.setAttribute("aria-expanded", String(opening));
  });

  menu.addEventListener("change", () => {
    updateNameMultiSelect(control);
    control.dispatchEvent(new Event("change", { bubbles: true }));
  });

  cell.appendChild(control);
  updateNameMultiSelect(control, students.length > 0);
}

function updateNameMultiSelect(control, hasRoster = !control.querySelector(".name-multiselect__empty")) {
  const names = [...control.querySelectorAll("input:checked")].map(input => input.value);
  control.querySelector(".name-multiselect__value").textContent = names.length
    ? names.map(displayStudentName).join("\n")
    : hasRoster ? "Сурагч сонгох" : control.dataset.emptyLabel || "Жагсаалт оруулаагүй";
  control.title = names.map(displayStudentName).join(", ");
}

function displayStudentName(name) {
  const [ovog, ...givenNames] = String(name).trim().split(/\s+/);
  return givenNames.length ? `${[...ovog][0]}. ${givenNames.join(" ")}` : name;
}

function selectedNames(cell) {
  return [...cell.querySelectorAll(".name-multiselect input:checked")]
    .map(input => input.value)
    .join(", ");
}

function displaySelectedNames(cell) {
  return selectedNames(cell)
    .split(", ")
    .filter(Boolean)
    .map(displayStudentName)
    .join(", ");
}

/* ================= TOTAL CALCULATION ================= */
function initTotals() {
  calculateTotals();
}

function calculateTotals() {
  let incomingSum = 0;
  let presentSum = 0;
  const absenceSums = [0, 0, 0];
  const studentNameSums = [[], []];

  document.querySelectorAll(".attendance-table tbody tr").forEach(row => {
    const tds = row.querySelectorAll("td");
    const incoming = parseInt(tds[3]?.querySelector("input")?.value || 0, 10);
    const present  = parseInt(tds[4]?.querySelector("input")?.value || 0, 10);

    incomingSum += isNaN(incoming) ? 0 : incoming;
    presentSum  += isNaN(present)  ? 0 : present;

    [5, 6, 7].forEach((cellIndex, sumIndex) => {
      const value = parseInt(tds[cellIndex]?.querySelector("input")?.value || 0, 10);
      absenceSums[sumIndex] += isNaN(value) ? 0 : value;
    });

    [8, 9].forEach((cellIndex, sumIndex) => {
      const names = displaySelectedNames(tds[cellIndex]);
      if (names) studentNameSums[sumIndex].push(`${tds[2].innerText.trim()}: ${names}`);
    });
  });

  const incomingTotal = document.getElementById("incoming-total");
  const presentTotal  = document.getElementById("present-total");

  if (incomingTotal) incomingTotal.value = incomingSum;
  if (presentTotal)  presentTotal.value  = presentSum;

  const totalCells = document.querySelectorAll(".attendance-table tfoot tr td");
  [3, 4, 5].forEach((cellIndex, sumIndex) => {
    const totalInput = totalCells[cellIndex]?.querySelector("input");
    if (totalInput) {
      totalInput.value = absenceSums[sumIndex];
      totalInput.readOnly = true;
    }
  });

  [6, 7].forEach((cellIndex, sumIndex) => {
    const totalCell = totalCells[cellIndex];
    if (totalCell) totalCell.textContent = studentNameSums[sumIndex].join("; ");
  });
}

/* ================= VALIDATION ================= */
function initValidation() {
  document.querySelectorAll(".attendance-table tbody tr").forEach(row => {
    const tds = row.querySelectorAll("td");

    const incomingInput = tds[3]?.querySelector("input");
    const presentInput  = tds[4]?.querySelector("input");
    const spinners = [
      tds[5]?.querySelector("input"),
      tds[6]?.querySelector("input"),
      tds[7]?.querySelector("input")
    ];

    if (!incomingInput || !presentInput) return;

    const errorBox = document.createElement("div");
    errorBox.className = "cell-error";
    errorBox.style.color = "red";
    errorBox.style.fontSize = "11px";
    tds[7].appendChild(errorBox);

    const validate = () => {
      const incoming = parseInt(incomingInput.value || 0, 10);
      const present  = parseInt(presentInput.value || 0, 10);
      const spinnerSum = spinners.reduce((s, i) => s + (parseInt(i?.value || 0, 10)), 0);

      errorBox.textContent = "";

      if (present > incoming) {
        errorBox.textContent = "Ирсэн > Ирэх";
      } else if (spinnerSum > incoming - present) {
        errorBox.textContent = "Ө/Ч/Т нийлбэр буруу";
      }
    };

    [incomingInput, presentInput, ...spinners].forEach(i => {
      if (i) i.addEventListener("input", validate);
    });
  });
}

/* ================= EXPORT TOOLBAR ================= */
function initExportToolbar() {
  const toolbar = document.querySelector(".export-toolbar");
  if (!toolbar) return;

  toolbar.addEventListener("click", e => {
    const btn = e.target.closest("button");
    if (!btn) return;

    const action = btn.dataset.action;

    if (action === "excel")  exportExcel();
    if (action === "pdf")    exportPDF();
    if (action === "exit")   resetInputs(); // 🧹 цэвэрлэх
  });
}

/* ================= RESET ================= */
function resetInputs() {
  if (!confirm("Бүх өгөгдлийг цэвэрлэх үү?")) return;

  localStorage.removeItem(DRAFT_STORAGE_KEY);

  document.querySelectorAll("input").forEach(input => {
    if (input.type === "number") input.value = "0";
    else if (input.type === "checkbox") input.checked = false;
    else input.value = "";
  });
  document.querySelectorAll("textarea").forEach(textarea => {
    textarea.value = "";
  });

  document.querySelectorAll(".late-students-table select").forEach(select => {
    select.selectedIndex = 0;
    if (select.classList.contains("late-class-select")) {
      select.dispatchEvent(new Event("change"));
    }
  });

  document.querySelectorAll(".name-multiselect").forEach(updateNameMultiSelect);

  calculateTotals();
}

/* ================= EXCEL EXPORT ================= */
function exportExcel() {
  if (typeof XLSX === "undefined") {
    alert("Excel library ачаалагдаагүй байна");
    return;
  }

  const title = "2026-2027 оны хичээлийн жилийн сурагчдын ирцийн тэмдэглэл";
  const headers = ["№","Анги №","Анги бүлэг","Ирэх","Ирсэн","Ө","Ч","Т","Тасалсан","Чөлөө"];
  const currentShift = document.body.dataset.shift || "1";
  const otherShift = currentShift === "1" ? "2" : "1";
  const currentDraft = captureAttendanceDraft();
  let otherDraft = null;
  try {
    otherDraft = JSON.parse(localStorage.getItem(`attendance-draft-v${otherShift}`) || "null");
  } catch (error) {
    console.warn(`Shift ${otherShift} attendance draft could not be read.`, error);
  }

  const data = [[title]];
  ["1", "2"].forEach(shift => {
    const draft = shift === currentShift ? currentDraft : otherDraft;
    data.push([`${shift}-р ээлж`], [], headers);
    const rows = document.querySelectorAll(".attendance-table tbody tr");
    let incomingTotal = 0;
    let presentTotal = 0;
    rows.forEach((row, index) => {
      const cells = row.querySelectorAll("td");
      const saved = draft?.attendance?.[index];
      const values = saved?.values || [];
      const incoming = shift === currentShift ? cells[3].querySelector("input").value : (values[0] || "");
      const present = shift === currentShift ? cells[4].querySelector("input").value : (values[1] || "");
      incomingTotal += Number(incoming) || 0;
      presentTotal += Number(present) || 0;
      data.push([
        cells[0].innerText.trim(), cells[1].innerText.trim(), cells[2].innerText.trim(),
        incoming, present,
        ...(shift === currentShift
          ? [cells[5].querySelector("input").value, cells[6].querySelector("input").value, cells[7].querySelector("input").value]
          : [values[2] || "", values[3] || "", values[4] || ""]),
        shift === currentShift ? displaySelectedNames(cells[8]) : (saved?.absentNames || ""),
        shift === currentShift ? displaySelectedNames(cells[9]) : (saved?.leaveNames || "")
      ]);
    });
    data.push(["", "", "НИЙТ", incomingTotal, presentTotal], []);
  });

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(data);
  formatExportSheet(ws, data.length, 10, "A4:J4", [6, 12, 12, 9, 9, 7, 7, 7, 32, 32]);
  ws["!merges"] = [XLSX.utils.decode_range("A1:J1")];
  XLSX.utils.book_append_sheet(wb, ws, "Ирц");

  const electiveData = [["Сонгон судлах хичээлийн ирц"], [],
    ["Хичээл / багш", "Ирэх", "Ирсэн", "Ө", "Ч", "Т", "Тасалсан сурагч"]];
  document.querySelectorAll(".elective-table tbody tr").forEach(row => {
    const cells = row.querySelectorAll("td");
    electiveData.push([...cells].map(cell => exportCellValue(cell)));
  });
  const electiveWs = XLSX.utils.aoa_to_sheet(electiveData);
  formatExportSheet(electiveWs, electiveData.length, 7, "A3:G3", [28, 9, 9, 7, 7, 7, 34]);
  electiveWs["!merges"] = [XLSX.utils.decode_range("A1:G1")];
  XLSX.utils.book_append_sheet(wb, electiveWs, "Сонгон хичээл");

  const lateData = [["Хоцорсон сурагчдын бүртгэл"], [],
    ["№", "Анги бүлэг", "Сурагчийн нэр", "Асран хамгаалагчийн утас", "Гэрийн хаяг", "Хоцорсон давталт"]];
  document.querySelectorAll(".late-students-table tbody tr").forEach(row => {
    lateData.push([...row.querySelectorAll("td")].map(cell => exportCellValue(cell)));
  });
  const lateWs = XLSX.utils.aoa_to_sheet(lateData);
  formatExportSheet(lateWs, lateData.length, 6, "A3:F3", [6, 12, 25, 23, 30, 18]);
  lateWs["!merges"] = [XLSX.utils.decode_range("A1:F1")];
  XLSX.utils.book_append_sheet(wb, lateWs, "Хоцорсон");

  const dutyTeacherNotes = document.getElementById("duty-teacher-notes");
  if (dutyTeacherNotes) {
    const notesData = [["Жижүүр багшийн тэмдэглэл"], [dutyTeacherNotes.value]];
    const notesWs = XLSX.utils.aoa_to_sheet(notesData);
    notesWs["!cols"] = [{ wch: 100 }];
    notesWs["!rows"] = [{ hpt: 22 }, { hpt: 100 }];
    notesWs["!merges"] = [
      XLSX.utils.decode_range("A1:A1"),
      XLSX.utils.decode_range("A2:A2")
    ];
    XLSX.utils.book_append_sheet(wb, notesWs, "Багшийн тэмдэглэл");
  }

  const exportDate = new Date();
  const dateStamp = exportDate.toISOString().slice(0, 10);
  XLSX.writeFile(wb, `irts_both_shifts_${dateStamp}.xlsx`);
}

function captureAttendanceDraft() {
  return {
    attendance: [...document.querySelectorAll(".attendance-table tbody tr")].map(row => ({
      values: [3, 4, 5, 6, 7].map(index => row.cells[index]?.querySelector("input")?.value || ""),
      absentNames: selectedNames(row.cells[8]),
      leaveNames: selectedNames(row.cells[9])
    }))
  };
}

function exportCellValue(cell) {
  const nameSelector = cell.querySelector(".name-multiselect");
  if (nameSelector) return displaySelectedNames(cell);
  return cell.querySelector("input, select")?.value?.trim() || cell.textContent.trim();
}

function formatExportSheet(ws, rowCount, columnCount, filterRange, widths) {
  ws["!cols"] = widths.map(wch => ({ wch }));
  ws["!autofilter"] = { ref: filterRange };
  ws["!margins"] = { left: 0.25, right: 0.25, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 };
  ws["!pageSetup"] = { orientation: "landscape", paperSize: "9", fitToWidth: 1, fitToHeight: 0 };
  ws["!rows"] = Array.from({ length: rowCount }, (_, index) => ({ hpt: index < 2 ? 22 : 18 }));
  const endCell = XLSX.utils.encode_cell({ r: rowCount - 1, c: columnCount - 1 });
  ws["!ref"] = `A1:${endCell}`;
}

/* ================= PDF EXPORT ================= */
function exportPDF() {
  window.print();
}
