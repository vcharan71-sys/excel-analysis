const fileInput = document.getElementById("fileInput");
const sheetSelect = document.getElementById("sheetSelect");
const categorySelect = document.getElementById("categorySelect");
const valueSelect = document.getElementById("valueSelect");
const generateBtn = document.getElementById("generateBtn");
const statusText = document.getElementById("statusText");
const tableBody = document.getElementById("tableBody");
const chartMeta = document.getElementById("chartMeta");

let workbookData = [];
let pieChart = null;

fileInput.addEventListener("change", handleFileUpload);
sheetSelect.addEventListener("change", populateColumnSelectors);
generateBtn.addEventListener("click", generateAnalysis);

async function handleFileUpload(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  statusText.textContent = "Reading file...";
  resetResults();

  try {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });

    workbookData = workbook.SheetNames.map((name) => {
      const sheet = workbook.Sheets[name];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
      return { name, rows };
    }).filter((sheet) => sheet.rows.length > 0);

    if (!workbookData.length) {
      statusText.textContent = "No readable data found in this file.";
      return;
    }

    sheetSelect.innerHTML = workbookData
      .map((sheet) => `<option value="${escapeHtml(sheet.name)}">${escapeHtml(sheet.name)}</option>`)
      .join("");

    sheetSelect.disabled = false;
    populateColumnSelectors();
    generateBtn.disabled = false;
    statusText.textContent = "File loaded. Choose columns and generate analysis.";
  } catch (error) {
    statusText.textContent = "Could not read this file.";
    console.error(error);
  }
}

function populateColumnSelectors() {
  const currentSheet = getCurrentSheet();
  if (!currentSheet || !currentSheet.rows.length) return;

  const headers = [...new Set(currentSheet.rows.flatMap((row) => Object.keys(row)))];
  const numericHeaders = headers.filter((header) =>
    currentSheet.rows.some((row) => Number.isFinite(parseNumber(row[header])))
  );

  const suggestedCategory = headers.find((header) =>
    currentSheet.rows.some((row) => {
      const value = String(row[header] || "").trim();
      return value && !Number.isFinite(parseNumber(value));
    })
  ) || headers[0] || "";

  categorySelect.innerHTML = headers
    .map((header) => `<option value="${escapeHtml(header)}">${escapeHtml(header)}</option>`)
    .join("");

  valueSelect.innerHTML =
    `<option value="">Use row count</option>` +
    numericHeaders
      .map((header) => `<option value="${escapeHtml(header)}">${escapeHtml(header)}</option>`)
      .join("");

  categorySelect.disabled = false;
  valueSelect.disabled = false;

  categorySelect.value = suggestedCategory;
}

function generateAnalysis() {
  const sheet = getCurrentSheet();
  if (!sheet) return;

  const categoryColumn = categorySelect.value;
  const valueColumn = valueSelect.value;

  if (!categoryColumn) {
    statusText.textContent = "Please choose a category column.";
    return;
  }

  const grouped = new Map();

  for (const row of sheet.rows) {
    const label = String(row[categoryColumn] || "").trim() || "Blank";
    const value = valueColumn ? parseNumber(row[valueColumn]) : 1;

    if (!Number.isFinite(value)) continue;
    grouped.set(label, (grouped.get(label) || 0) + value);
  }

  const summary = [...grouped.entries()]
    .map(([label, value]) => ({ label, value }))
    .filter((item) => item.value > 0)
    .sort((a, b) => b.value - a.value);

  if (!summary.length) {
    statusText.textContent = "No chartable data found for the selected columns.";
    tableBody.innerHTML = `<tr><td colspan="3">No chartable data found.</td></tr>`;
    chartMeta.textContent = "Try another sheet or column combination.";
    if (pieChart) pieChart.destroy();
    return;
  }

  const total = summary.reduce((sum, item) => sum + item.value, 0);
  const colors = summary.map((_, index) => palette[index % palette.length]);

  renderTable(summary, total);
  renderChart(summary, colors);
  chartMeta.textContent = `${sheet.name} • ${categoryColumn} • ${valueColumn || "Row count"}`;
  statusText.textContent = "Analysis generated successfully.";
}

function renderTable(summary, total) {
  tableBody.innerHTML = summary
    .map((item) => {
      const percent = ((item.value / total) * 100).toFixed(1);
      return `
        <tr>
          <td>${escapeHtml(item.label)}</td>
          <td>${formatValue(item.value)}</td>
          <td>${percent}%</td>
        </tr>
      `;
    })
    .join("");
}

function renderChart(summary, colors) {
  const ctx = document.getElementById("pieChart").getContext("2d");

  if (pieChart) pieChart.destroy();

  pieChart = new Chart(ctx, {
    type: "pie",
    data: {
      labels: summary.map((item) => item.label),
      datasets: [
        {
          data: summary.map((item) => item.value),
          backgroundColor: colors,
          borderColor: "#ffffff",
          borderWidth: 2
        }
      ]
    },
    options: {
      responsive: true,
      plugins: {
        legend: {
          position: "bottom"
        }
      }
    }
  });
}

function getCurrentSheet() {
  return workbookData.find((sheet) => sheet.name === sheetSelect.value) || workbookData[0] || null;
}

function parseNumber(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : NaN;

  const normalized = String(value || "")
    .replaceAll(",", "")
    .replace(/[^\d.-]/g, "")
    .trim();

  if (!normalized) return NaN;

  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : NaN;
}

function formatValue(value) {
  const hasDecimal = Math.abs(value % 1) > 0.001;
  return new Intl.NumberFormat(undefined, {
    maximumFractionDigits: hasDecimal ? 2 : 0
  }).format(value);
}

function resetResults() {
  tableBody.innerHTML = `<tr><td colspan="3">No analysis yet.</td></tr>`;
  chartMeta.textContent = "Your chart will appear here.";
  if (pieChart) {
    pieChart.destroy();
    pieChart = null;
  }
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

const palette = [
  "#0a84d6",
  "#18b3b9",
  "#34c759",
  "#f59e0b",
  "#ef4444",
  "#8b5cf6",
  "#14b8a6",
  "#f97316"
];
