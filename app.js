const SHEET_SOURCE = {
  sheetId: "1pFzKzVdvlfXfXQyZybL6aYHBR2bTCvbYPgK2uQYBFQg",
  gid: "0",
  title: "LRMS Inventory System",
};

const SNAPSHOT_CACHE_KEY = "20260408-7";

const state = {
  rows: [],
  filteredRows: [],
  charts: {},
  activeSource: "loading",
  embedMode: detectEmbedMode(),
};

const ui = {
  dataSourceLabel: document.getElementById("dataSourceLabel"),
  dataSourceSubtext: document.getElementById("dataSourceSubtext"),
  statusBanner: document.getElementById("statusBanner"),
  refreshButton: document.getElementById("refreshButton"),
  exportButton: document.getElementById("exportButton"),
  resetFiltersButton: document.getElementById("resetFiltersButton"),
  schoolFilter: document.getElementById("schoolFilter"),
  categoryFilter: document.getElementById("categoryFilter"),
  typeFilter: document.getElementById("typeFilter"),
  statusFilter: document.getElementById("statusFilter"),
  gradeFilter: document.getElementById("gradeFilter"),
  areaFilter: document.getElementById("areaFilter"),
  searchFilter: document.getElementById("searchFilter"),
  metricRecords: document.getElementById("metricRecords"),
  metricQuantity: document.getElementById("metricQuantity"),
  metricSchools: document.getElementById("metricSchools"),
  metricDistributed: document.getElementById("metricDistributed"),
  metricInStock: document.getElementById("metricInStock"),
  metricLatestUpdate: document.getElementById("metricLatestUpdate"),
  stockTableBody: document.getElementById("stockTableBody"),
  updatesTableBody: document.getElementById("updatesTableBody"),
  detailTableBody: document.getElementById("detailTableBody"),
  detailSummary: document.getElementById("detailSummary"),
};

const statusClasses = {
  Distributed: "pill pill-distributed",
  "In-Stock": "pill pill-in-stock",
  "Library Use": "pill pill-library-use",
  Unspecified: "pill pill-unspecified",
};

const valueLabelsPlugin = {
  id: "valueLabels",
  afterDatasetsDraw(chart, _args, pluginOptions) {
    if (!pluginOptions?.enabled) return;

    const dataset = chart.data.datasets?.[0];
    if (!dataset) return;

    const ctx = chart.ctx;
    const meta = chart.getDatasetMeta(0);
    const chartType = pluginOptions.mode || chart.config.type;
    const numbers = (dataset.data || []).map((value) => Number(value) || 0);

    ctx.save();
    ctx.font = "600 11px 'Space Grotesk', sans-serif";

    if (chartType === "bar") {
      const indexAxis = chart.options?.indexAxis || "x";
      const rightEdge = chart.chartArea?.right || 0;
      const topEdge = chart.chartArea?.top || 0;

      numbers.forEach((value, index) => {
        if (!value) return;
        const element = meta.data[index];
        if (!element) return;

        const props = element.getProps(["x", "y"], true);
        const text = value.toLocaleString();

        if (indexAxis === "y") {
          let x = props.x + 12;
          let textAlign = "left";
          if (x > rightEdge - 6) {
            x = props.x - 8;
            textAlign = "right";
          }

          ctx.textAlign = textAlign;
          ctx.textBaseline = "middle";
          ctx.fillStyle = "#213144";
          ctx.fillText(text, x, props.y);
        } else {
          const y = Math.max(props.y - 8, topEdge + 12);
          ctx.textAlign = "center";
          ctx.textBaseline = "middle";
          ctx.fillStyle = "#213144";
          ctx.fillText(text, props.x, y);
        }
      });
    }

    if (chartType === "doughnut") {
      const total = numbers.reduce((sum, value) => sum + value, 0);
      if (!total) {
        ctx.restore();
        return;
      }

      numbers.forEach((value, index) => {
        if (!value) return;
        const ratio = value / total;
        if (ratio < 0.03) return;

        const arc = meta.data[index];
        if (!arc) return;

        const {
          x: centerX,
          y: centerY,
          startAngle,
          endAngle,
          innerRadius,
          outerRadius,
        } = arc.getProps(
          ["x", "y", "startAngle", "endAngle", "innerRadius", "outerRadius"],
          true
        );

        const angle = (startAngle + endAngle) / 2;
        const radius = (innerRadius + outerRadius) / 2;
        const x = centerX + Math.cos(angle) * radius;
        const y = centerY + Math.sin(angle) * radius;
        const text = value.toLocaleString();

        ctx.textAlign = "center";
        ctx.textBaseline = "middle";
        ctx.strokeStyle = "rgba(0, 0, 0, 0.45)";
        ctx.lineWidth = 3;
        ctx.strokeText(text, x, y);
        ctx.fillStyle = "#ffffff";
        ctx.fillText(text, x, y);
      });
    }

    ctx.restore();
  },
};

Chart.register(valueLabelsPlugin);

initialize();

async function initialize() {
  applyRuntimeMode();
  bindEvents();
  await loadDashboardData({ preferLive: true });
}

function bindEvents() {
  [
    ui.schoolFilter,
    ui.categoryFilter,
    ui.typeFilter,
    ui.statusFilter,
    ui.gradeFilter,
    ui.areaFilter,
    ui.searchFilter,
  ].forEach((element) => {
    element.addEventListener("input", applyFiltersAndRender);
  });

  ui.resetFiltersButton.addEventListener("click", () => {
    [
      ui.schoolFilter,
      ui.categoryFilter,
      ui.typeFilter,
      ui.statusFilter,
      ui.gradeFilter,
      ui.areaFilter,
    ].forEach((element) => {
      element.value = "";
    });
    ui.searchFilter.value = "";
    applyFiltersAndRender();
  });

  ui.refreshButton.addEventListener("click", async () => {
    ui.refreshButton.disabled = true;
    await loadDashboardData({ preferLive: true, forceRefreshMessage: true });
    ui.refreshButton.disabled = false;
  });

  ui.exportButton.addEventListener("click", exportFilteredCsv);
}

async function loadDashboardData({ preferLive, forceRefreshMessage = false }) {
  setStatus("Loading dashboard data...", "loading");

  if (preferLive) {
    try {
      const liveTable = await loadGoogleSheetTable();
      state.rows = normalizeTable(liveTable);
      state.activeSource = "live";
      hydrateDashboard({
        banner: forceRefreshMessage
          ? "Live data refreshed from Google Sheet."
          : "Connected to the live Google Sheet.",
      });
      return;
    } catch (error) {
      console.warn("Live sheet load failed:", error);
    }
  }

  const snapshotLoaded = await ensureSnapshotLoaded();
  if (snapshotLoaded && window.DASHBOARD_SNAPSHOT?.table) {
    state.rows = normalizeTable(window.DASHBOARD_SNAPSHOT.table);
    state.activeSource = "snapshot";
    const generatedAt = window.DASHBOARD_SNAPSHOT.generatedAt
      ? formatDateTime(new Date(window.DASHBOARD_SNAPSHOT.generatedAt))
      : "an earlier saved copy";
    hydrateDashboard({
      banner: `Live sheet request failed, so the dashboard is using the local snapshot saved on ${generatedAt}.`,
    });
    return;
  }

  state.rows = [];
  state.filteredRows = [];
  setSourceMeta("No data loaded", "Live Google Sheet and local snapshot are both unavailable.");
  setStatus("Unable to load dashboard data.", "error");
  applyFiltersAndRender();
}

function applyRuntimeMode() {
  document.body.classList.toggle("embed-mode", state.embedMode);
}

function detectEmbedMode() {
  const params = new URLSearchParams(window.location.search);
  if (params.get("embed") === "1" || params.get("mode") === "embed") {
    return true;
  }

  try {
    return window.self !== window.top;
  } catch (_error) {
    return false;
  }
}

function hydrateDashboard({ banner }) {
  populateFilters(state.rows);
  applyFiltersAndRender();

  if (state.activeSource === "live") {
    setSourceMeta("Live Google Sheet", `${SHEET_SOURCE.title} (${state.rows.length.toLocaleString()} rows loaded)`);
    setStatus(banner, "live");
  } else {
    setSourceMeta(
      "Local snapshot",
      `${SHEET_SOURCE.title} (${state.rows.length.toLocaleString()} rows available offline)`
    );
    setStatus(banner, "fallback");
  }
}

function setSourceMeta(label, subtext) {
  ui.dataSourceLabel.textContent = label;
  ui.dataSourceSubtext.textContent = subtext;
}

function setStatus(message, kind) {
  ui.statusBanner.textContent = message;
  ui.statusBanner.className = "status-banner";
  if (kind === "live") ui.statusBanner.classList.add("is-live");
  if (kind === "fallback") ui.statusBanner.classList.add("is-fallback");
  if (kind === "error") ui.statusBanner.classList.add("is-error");
  if (kind === "loading") ui.statusBanner.classList.add("is-loading");
}

function populateFilters(rows) {
  populateSelect(ui.schoolFilter, rows.map((row) => row.school), "All schools");
  populateSelect(ui.categoryFilter, rows.map((row) => row.category), "All categories");
  populateSelect(ui.typeFilter, rows.map((row) => row.type), "All types");
  populateSelect(ui.statusFilter, rows.map((row) => row.status), "All statuses");
  populateSelect(ui.gradeFilter, rows.map((row) => row.gradeLevel), "All grade levels");
  populateSelect(ui.areaFilter, rows.map((row) => row.learningArea), "All learning areas");
}

function populateSelect(select, values, placeholder) {
  const currentValue = select.value;
  const uniqueValues = [...new Set(values.filter(Boolean))].sort((left, right) =>
    left.localeCompare(right)
  );

  select.innerHTML = "";
  const blankOption = document.createElement("option");
  blankOption.value = "";
  blankOption.textContent = placeholder;
  select.append(blankOption);

  uniqueValues.forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    select.append(option);
  });

  select.value = uniqueValues.includes(currentValue) ? currentValue : "";
}

function applyFiltersAndRender() {
  const searchTerm = ui.searchFilter.value.trim().toLowerCase();

  state.filteredRows = state.rows.filter((row) => {
    if (ui.schoolFilter.value && row.school !== ui.schoolFilter.value) return false;
    if (ui.categoryFilter.value && row.category !== ui.categoryFilter.value) return false;
    if (ui.typeFilter.value && row.type !== ui.typeFilter.value) return false;
    if (ui.statusFilter.value && row.status !== ui.statusFilter.value) return false;
    if (ui.gradeFilter.value && row.gradeLevel !== ui.gradeFilter.value) return false;
    if (ui.areaFilter.value && row.learningArea !== ui.areaFilter.value) return false;
    if (searchTerm && !row.searchableText.includes(searchTerm)) return false;
    return true;
  });

  renderMetrics();
  renderCharts();
  renderTables();
}

function renderMetrics() {
  const rows = state.filteredRows;
  const totalQuantity = sumBy(rows, (row) => row.quantity);
  const distributedQuantity = sumBy(rows.filter((row) => row.status === "Distributed"), (row) => row.quantity);
  const inStockQuantity = sumBy(rows.filter((row) => row.status === "In-Stock"), (row) => row.quantity);
  const latestUpdatedRow = [...rows]
    .filter((row) => row.lastUpdated || row.receivedDate)
    .sort((left, right) => (right.lastUpdatedOrReceivedMs || 0) - (left.lastUpdatedOrReceivedMs || 0))[0];

  ui.metricRecords.textContent = rows.length.toLocaleString();
  ui.metricQuantity.textContent = totalQuantity.toLocaleString();
  ui.metricSchools.textContent = new Set(rows.map((row) => row.school)).size.toLocaleString();
  ui.metricDistributed.textContent = totalQuantity
    ? `${Math.round((distributedQuantity / totalQuantity) * 100)}%`
    : "0%";
  ui.metricInStock.textContent = inStockQuantity.toLocaleString();
  ui.metricLatestUpdate.textContent = latestUpdatedRow
    ? latestUpdatedRow.lastUpdated
      ? formatDateTime(latestUpdatedRow.lastUpdated)
      : formatDate(latestUpdatedRow.receivedDate)
    : "--";
}

function renderCharts() {
  const rows = state.filteredRows;

  renderBarChart("schoolChart", aggregateTop(rows, "school", 8), {
    label: "Quantity",
    color: "#1f6f68",
    showValues: true,
  });

  renderDonutChart("statusChart", aggregate(rows, "status"), {
    colors: ["#1f6f68", "#d96c3a", "#6674c7", "#a0a8b0"],
    showValues: true,
  });

  renderBarChart("typeChart", aggregateTop(rows, "type", 8), {
    label: "Quantity",
    color: "#d96c3a",
  });

  renderLineChart("timelineChart", aggregateTimeline(rows), {
    label: "Received quantity",
    color: "#135e7b",
  });
}

function renderBarChart(canvasId, dataset, options) {
  const canvas = document.getElementById(canvasId);
  const labels = dataset.map((entry) => entry.label);
  const values = dataset.map((entry) => entry.value);
  const chartOptions = baseChartOptions({
    indexAxis: canvasId === "schoolChart" ? "y" : "x",
  });

  if (options.showValues) {
    chartOptions.plugins.valueLabels = {
      enabled: true,
      mode: "bar",
    };
  }

  const config = {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          label: options.label,
          data: values,
          backgroundColor: options.color,
          borderRadius: 12,
          maxBarThickness: 34,
        },
      ],
    },
    options: chartOptions,
  };

  upsertChart(canvasId, canvas, config);
}

function renderDonutChart(canvasId, dataset, options) {
  const canvas = document.getElementById(canvasId);
  const labels = dataset.map((entry) => entry.label);
  const values = dataset.map((entry) => entry.value);

  const config = {
    type: "doughnut",
    data: {
      labels,
      datasets: [
        {
          data: values,
          backgroundColor: options.colors,
          borderWidth: 0,
          hoverOffset: 6,
        },
      ],
    },
    options: {
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: "bottom",
          labels: {
            usePointStyle: true,
            color: "#5f6c7c",
            padding: 18,
            font: {
              family: "Space Grotesk",
            },
          },
        },
        tooltip: tooltipOptions(),
        valueLabels: {
          enabled: Boolean(options.showValues),
          mode: "doughnut",
        },
      },
      cutout: "62%",
    },
  };

  upsertChart(canvasId, canvas, config);
}

function renderLineChart(canvasId, dataset, options) {
  const canvas = document.getElementById(canvasId);
  const labels = dataset.map((entry) => entry.label);
  const values = dataset.map((entry) => entry.value);

  const config = {
    type: "line",
    data: {
      labels,
      datasets: [
        {
          label: options.label,
          data: values,
          borderColor: options.color,
          backgroundColor: "rgba(19, 94, 123, 0.12)",
          fill: true,
          tension: 0.32,
          pointRadius: 3,
          pointHoverRadius: 5,
        },
      ],
    },
    options: baseChartOptions(),
  };

  upsertChart(canvasId, canvas, config);
}

function baseChartOptions(extraOptions = {}) {
  return {
    maintainAspectRatio: false,
    responsive: true,
    plugins: {
      legend: {
        display: false,
      },
      tooltip: tooltipOptions(),
    },
    scales: {
      x: {
        grid: {
          color: "rgba(33, 49, 68, 0.07)",
        },
        ticks: {
          color: "#5f6c7c",
          font: {
            family: "Space Grotesk",
          },
        },
      },
      y: {
        grid: {
          color: "rgba(33, 49, 68, 0.07)",
        },
        ticks: {
          color: "#5f6c7c",
          font: {
            family: "Space Grotesk",
          },
        },
      },
    },
    ...extraOptions,
  };
}

function tooltipOptions() {
  return {
    backgroundColor: "#213144",
    titleColor: "#fff",
    bodyColor: "#fff",
    padding: 12,
    displayColors: false,
    callbacks: {
      label(context) {
        return `${context.raw.toLocaleString()} pcs`;
      },
    },
  };
}

function upsertChart(key, canvas, config) {
  if (state.charts[key]) {
    state.charts[key].destroy();
  }
  state.charts[key] = new Chart(canvas, config);
}

function renderTables() {
  const rows = state.filteredRows;
  const inStockRows = [...rows]
    .filter((row) => row.status === "In-Stock")
    .sort((left, right) => right.quantity - left.quantity)
    .slice(0, 15);

  const latestRows = [...rows]
    .filter((row) => row.lastUpdated || row.receivedDate)
    .sort((left, right) => (right.lastUpdatedOrReceivedMs || 0) - (left.lastUpdatedOrReceivedMs || 0))
    .slice(0, 15);

  const detailRows = [...rows]
    .sort((left, right) => (right.lastUpdatedOrReceivedMs || 0) - (left.lastUpdatedOrReceivedMs || 0))
    .slice(0, 20);

  renderRows(ui.stockTableBody, inStockRows, (row) => `
    <tr>
      <td>${row.receivedDate ? formatDate(row.receivedDate) : "--"}</td>
      <td><div class="cell-clamp-2">${escapeHtml(row.title)}</div></td>
      <td><div class="cell-clamp-1">${escapeHtml(row.school)}</div></td>
      <td><div class="cell-clamp-1">${escapeHtml(row.category)}</div></td>
      <td><div class="cell-clamp-1">${escapeHtml(row.type)}</div></td>
      <td><div class="cell-clamp-1">${escapeHtml(row.gradeLevel)}</div></td>
      <td><div class="cell-clamp-2">${escapeHtml(row.learningArea)}</div></td>
      <td><span class="${statusClass(row.status)}">${escapeHtml(row.status)}</span></td>
      <td class="num">${row.quantity.toLocaleString()}</td>
    </tr>
  `, 9);

  renderRows(ui.updatesTableBody, latestRows, (row) => `
    <tr>
      <td>${row.lastUpdated ? formatDateTime(row.lastUpdated) : formatDate(row.receivedDate)}</td>
      <td>${row.receivedDate ? formatDate(row.receivedDate) : "--"}</td>
      <td><div class="cell-clamp-2">${escapeHtml(row.title)}</div></td>
      <td><div class="cell-clamp-1">${escapeHtml(row.school)}</div></td>
      <td><div class="cell-clamp-1">${escapeHtml(row.category)}</div></td>
      <td><div class="cell-clamp-1">${escapeHtml(row.type)}</div></td>
      <td><div class="cell-clamp-1">${escapeHtml(row.gradeLevel)}</div></td>
      <td><span class="${statusClass(row.status)}">${escapeHtml(row.status)}</span></td>
      <td class="num">${row.quantity.toLocaleString()}</td>
    </tr>
  `, 9);

  renderRows(ui.detailTableBody, detailRows, (row) => `
    <tr>
      <td>${row.receivedDate ? formatDate(row.receivedDate) : "--"}</td>
      <td><div class="cell-clamp-1">${escapeHtml(row.school)}</div></td>
      <td><div class="cell-clamp-1">${escapeHtml(row.category)}</div></td>
      <td><div class="cell-clamp-1">${escapeHtml(row.type)}</div></td>
      <td><div class="cell-clamp-1">${escapeHtml(row.gradeLevel)}</div></td>
      <td><div class="cell-clamp-2">${escapeHtml(row.learningArea)}</div></td>
      <td><div class="cell-clamp-3">${escapeHtml(row.title)}</div></td>
      <td><span class="${statusClass(row.status)}">${escapeHtml(row.status)}</span></td>
      <td class="num">${row.quantity.toLocaleString()}</td>
    </tr>
  `, 9);

  const total = rows.length.toLocaleString();
  const showing = Math.min(rows.length, 20).toLocaleString();
  ui.detailSummary.textContent = `Showing ${showing} of ${total} filtered rows`;
}

function renderRows(target, rows, template, colspan) {
  if (!rows.length) {
    target.innerHTML = `<tr><td colspan="${colspan}" class="empty-state">No matching rows found.</td></tr>`;
    return;
  }

  target.innerHTML = rows.map(template).join("");
}

function exportFilteredCsv() {
  const headers = [
    "Received Date",
    "School",
    "Category",
    "Type",
    "Grade Level",
    "Learning Area",
    "Title",
    "Quantity",
    "Unit",
    "Status",
    "Last Updated",
  ];

  const lines = state.filteredRows.map((row) => [
    row.receivedDate ? formatDate(row.receivedDate) : "",
    row.school,
    row.category,
    row.type,
    row.gradeLevel,
    row.learningArea,
    row.title,
    row.quantity,
    row.unit,
    row.status,
    row.lastUpdated ? formatDateTime(row.lastUpdated) : "",
  ]);

  const csv = [headers, ...lines]
    .map((line) => line.map(csvEscape).join(","))
    .join("\n");

  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = "lrms-dashboard-filtered.csv";
  link.click();
  URL.revokeObjectURL(link.href);
}

async function loadGoogleSheetTable() {
  return new Promise((resolve, reject) => {
    const callbackName = `lrmsSheetCallback_${Date.now()}`;
    const timeout = window.setTimeout(() => {
      cleanup();
      reject(new Error("Timed out while loading the Google Sheet."));
    }, 12000);

    window[callbackName] = (payload) => {
      cleanup();
      if (payload?.status !== "ok" || !payload?.table) {
        reject(new Error("Google Sheet returned an invalid response."));
        return;
      }
      resolve(payload.table);
    };

    const script = document.createElement("script");
    const query = encodeURIComponent("select *");
    script.src = `https://docs.google.com/spreadsheets/d/${SHEET_SOURCE.sheetId}/gviz/tq?tqx=out:json;responseHandler:${callbackName}&gid=${SHEET_SOURCE.gid}&headers=1&tq=${query}`;
    script.async = true;
    script.onerror = () => {
      cleanup();
      reject(new Error("The browser could not load the Google Sheet script."));
    };

    function cleanup() {
      window.clearTimeout(timeout);
      delete window[callbackName];
      script.remove();
    }

    document.head.appendChild(script);
  });
}

async function ensureSnapshotLoaded() {
  if (window.DASHBOARD_SNAPSHOT?.table) {
    return true;
  }

  return new Promise((resolve) => {
    const script = document.createElement("script");
    const timeout = window.setTimeout(() => {
      cleanup();
      resolve(false);
    }, 12000);

    script.src = `./data/snapshot.js?v=${SNAPSHOT_CACHE_KEY}`;
    script.async = true;
    script.onload = () => {
      cleanup();
      resolve(Boolean(window.DASHBOARD_SNAPSHOT?.table));
    };
    script.onerror = () => {
      cleanup();
      resolve(false);
    };

    function cleanup() {
      window.clearTimeout(timeout);
      script.remove();
    }

    document.head.appendChild(script);
  });
}

function normalizeTable(table) {
  const columns = table.cols.map((column, index) => {
    const label = (column.label || "").trim();
    return label || column.id || `Column ${index + 1}`;
  });

  return table.rows.map((row, index) => {
    const raw = {};
    columns.forEach((columnName, columnIndex) => {
      raw[columnName] = parseCell(row.c?.[columnIndex]);
    });
    return normalizeRow(raw, index);
  });
}

function parseCell(cell) {
  if (!cell) return "";
  if (cell.f && typeof cell.v !== "string") return cell.f;
  return cell.v ?? "";
}

function normalizeRow(raw, index) {
  const receivedDate = parseSheetDate(raw["Received Date"]);
  const lastUpdated = parseSheetDate(raw["Last Updated"]);
  const quantity = parseNumber(raw.Quantity);
  const status = cleanText(raw.Status) || "Unspecified";

  const school = cleanText(raw.School) || "Unspecified";
  const category = cleanText(raw.Category) || "Unspecified";
  const type = normalizeSpacing(raw.Type) || "Unspecified";
  const gradeLevel = cleanText(raw["Grade Level"]) || "Unspecified";
  const learningArea = cleanText(raw["Learning Area"]) || "Unspecified";
  const title = cleanText(raw.Title) || "Untitled resource";
  const searchableText = [
    raw["ITR No. / DR No."],
    raw["Publisher / Supplier"],
    title,
    school,
    category,
    type,
    gradeLevel,
    learningArea,
    raw.Remarks,
    status,
  ]
    .map(cleanText)
    .join(" ")
    .toLowerCase();

  return {
    id: cleanText(raw["No."]) || String(index + 1),
    school,
    category,
    type,
    quarter: cleanText(raw.Quarter),
    weekNo: cleanText(raw["Week No."]),
    gradeLevel,
    learningArea,
    title,
    quantity,
    unit: cleanText(raw.Unit) || "Pcs",
    status,
    remarks: cleanText(raw.Remarks),
    procuringEntity: cleanText(raw["Procuring Entity"]),
    itrNo: cleanText(raw["ITR No. / DR No."]),
    supplier: cleanText(raw["Publisher / Supplier"]),
    receivedDate,
    lastUpdated,
    lastUpdatedOrReceivedMs:
      lastUpdated?.getTime() ?? receivedDate?.getTime() ?? 0,
    searchableText,
  };
}

function parseSheetDate(value) {
  if (!value) return null;

  const text = String(value).trim();
  const gvizMatch = text.match(/^Date\((.+)\)$/);
  if (gvizMatch) {
    const parts = gvizMatch[1].split(",").map((part) => Number(part.trim()));
    const [year, month, day = 1, hours = 0, minutes = 0, seconds = 0] = parts;
    const date = new Date(year, month, day, hours, minutes, seconds);
    return Number.isNaN(date.getTime()) ? null : date;
  }

  const slashMatch = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (slashMatch) {
    const [, monthText, dayText, yearText] = slashMatch;
    const date = new Date(Number(yearText), Number(monthText) - 1, Number(dayText));
    return Number.isNaN(date.getTime()) ? null : date;
  }

  const isoDate = new Date(text);
  return Number.isNaN(isoDate.getTime()) ? null : isoDate;
}

function parseNumber(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const text = String(value || "").replace(/,/g, "").trim();
  const parsed = Number(text);
  return Number.isFinite(parsed) ? parsed : 0;
}

function normalizeSpacing(value) {
  return cleanText(value).replace(/\s+/g, " ");
}

function cleanText(value) {
  return String(value ?? "").trim();
}

function aggregate(rows, key) {
  const map = new Map();
  rows.forEach((row) => {
    const label = row[key] || "Unspecified";
    map.set(label, (map.get(label) || 0) + row.quantity);
  });

  return [...map.entries()]
    .map(([label, value]) => ({ label, value }))
    .sort((left, right) => right.value - left.value);
}

function aggregateTop(rows, key, count) {
  return aggregate(rows, key).slice(0, count);
}

function aggregateTimeline(rows) {
  const map = new Map();

  rows.forEach((row) => {
    if (!row.receivedDate) return;
    const year = row.receivedDate.getFullYear();
    if (year < 2024 || year > 2035) return;
    const label = `${year}-${String(row.receivedDate.getMonth() + 1).padStart(2, "0")}`;
    map.set(label, (map.get(label) || 0) + row.quantity);
  });

  return [...map.entries()]
    .sort((left, right) => left[0].localeCompare(right[0]))
    .map(([label, value]) => ({ label, value }));
}

function sumBy(rows, getter) {
  return rows.reduce((total, row) => total + getter(row), 0);
}

function formatDate(date) {
  return new Intl.DateTimeFormat("en-PH", {
    year: "numeric",
    month: "short",
    day: "numeric",
  }).format(date);
}

function formatDateTime(date) {
  return new Intl.DateTimeFormat("en-PH", {
    year: "numeric",
    month: "short",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit",
  }).format(date);
}

function statusClass(status) {
  return statusClasses[status] || statusClasses.Unspecified;
}

function csvEscape(value) {
  const text = String(value ?? "");
  return `"${text.replace(/"/g, '""')}"`;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}
