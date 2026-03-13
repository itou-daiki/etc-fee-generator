// ============================================================
// Constants
// ============================================================

const HIGHWAY_SECTIONS = [
    "安心院", "朝地", "院内", "宇佐", "大分", "大分光吉",
    "大分宮河内", "大分松岡", "大分米良", "蒲江", "北浦",
    "玖珠", "佐伯", "佐伯堅田", "竹田", "津久見", "津久見南",
    "日田", "別府", "湯布院"
].sort();

const MONTHLY_ALLOWANCE = 112560;

// ============================================================
// State (immutable pattern - replaced, never mutated)
// ============================================================

let appState = Object.freeze({
    csvData: null,
    csvHeaders: null,
    year: null,
    month: null,
    generatedBlob: null,
});

function updateState(partial) {
    appState = Object.freeze({ ...appState, ...partial });
}

// ============================================================
// Initialization
// ============================================================

document.addEventListener("DOMContentLoaded", () => {
    initializeSelects();
    setupEventListeners();
    updateSettingsDisplay();
});

function initializeSelects() {
    const fromSelect = document.getElementById("highway-from");
    const toSelect = document.getElementById("highway-to");

    HIGHWAY_SECTIONS.forEach((section) => {
        fromSelect.add(new Option(section, section));
        toSelect.add(new Option(section, section));
    });

    fromSelect.value = "大分米良";
    toSelect.value = "日田";
}

function setupEventListeners() {
    const fileInput = document.getElementById("csv-file");
    const dropzone = document.getElementById("dropzone");

    // Dropzone interactions
    dropzone.addEventListener("click", () => fileInput.click());
    dropzone.addEventListener("dragover", (e) => {
        e.preventDefault();
        dropzone.classList.add("dragover");
    });
    dropzone.addEventListener("dragleave", () => {
        dropzone.classList.remove("dragover");
    });
    dropzone.addEventListener("drop", (e) => {
        e.preventDefault();
        dropzone.classList.remove("dragover");
        if (e.dataTransfer.files.length > 0) {
            handleFile(e.dataTransfer.files[0]);
        }
    });
    fileInput.addEventListener("change", (e) => {
        if (e.target.files.length > 0) {
            handleFile(e.target.files[0]);
        }
    });

    // Generate button
    document.getElementById("generate-btn").addEventListener("click", generateExcel);

    // Settings change listeners
    const settingIds = [
        "organization", "position", "user-name",
        "highway-from", "highway-to", "one-way-fee",
    ];
    settingIds.forEach((id) => {
        document.getElementById(id).addEventListener("input", updateSettingsDisplay);
    });

    // Mobile sidebar toggle
    const toggle = document.getElementById("sidebar-toggle");
    toggle.addEventListener("click", () => {
        document.getElementById("sidebar-content").classList.toggle("open");
    });
}

// ============================================================
// Settings Display
// ============================================================

function updateSettingsDisplay() {
    const org = document.getElementById("organization").value;
    const pos = document.getElementById("position").value;
    const name = document.getElementById("user-name").value;
    const from = document.getElementById("highway-from").value;
    const to = document.getElementById("highway-to").value;
    const fee = parseInt(document.getElementById("one-way-fee").value, 10) || 0;

    document.getElementById("settings-display").innerHTML = [
        org ? `<p><strong>所属:</strong> ${escapeHtml(org)}</p>` : "",
        pos ? `<p><strong>職:</strong> ${escapeHtml(pos)}</p>` : "",
        name ? `<p><strong>氏名:</strong> ${escapeHtml(name)}</p>` : "",
        `<p><strong>利用区間:</strong> ${escapeHtml(from)} ⇔ ${escapeHtml(to)}</p>`,
        `<p><strong>片道料金:</strong> ¥${fee.toLocaleString()}</p>`,
    ].join("");
}

// ============================================================
// File Handling
// ============================================================

async function handleFile(file) {
    if (!file.name.toLowerCase().endsWith(".csv")) {
        showMessage("error", "CSVファイルを選択してください。");
        return;
    }

    clearMessages();

    try {
        const buffer = await file.arrayBuffer();
        const uint8Array = new Uint8Array(buffer);

        // Detect encoding
        const detectedEncoding = Encoding.detect(uint8Array);
        showMessage("info", `検出されたエンコーディング: ${detectedEncoding}`);

        // Convert to Unicode string
        const unicodeArray = Encoding.convert(uint8Array, {
            to: "UNICODE",
            from: detectedEncoding,
        });
        const text = Encoding.codeToString(unicodeArray);

        // Parse CSV
        const result = Papa.parse(text, {
            header: true,
            skipEmptyLines: true,
            dynamicTyping: false,
        });

        if (result.data.length === 0) {
            showMessage("error", "CSVファイルにデータがありません。");
            return;
        }

        const headers = result.meta.fields;
        const data = result.data;

        // Extract year/month
        const yearMonth = extractYearMonth(data);
        if (!yearMonth) {
            showMessage("error", "データから年月を抽出できませんでした。");
            return;
        }

        updateState({
            csvData: data,
            csvHeaders: headers,
            year: yearMonth.year,
            month: yearMonth.month,
            generatedBlob: null,
        });

        showMessage("success", `データ期間: ${yearMonth.year}年${yearMonth.month}月`);

        // Update UI
        showPreview(data, headers);
        showStats(data);
        document.getElementById("data-section").classList.remove("hidden");
        document.getElementById("download-section").classList.add("hidden");
        document.getElementById("warnings").classList.add("hidden");

        // Update dropzone text
        document.getElementById("dropzone-text").innerHTML =
            `📄 ${escapeHtml(file.name)}<br><span class="dropzone-hint">読み込み済み - 再選択するにはクリック</span>`;
    } catch (e) {
        showMessage("error", `ファイルの読み込みに失敗しました: ${e.message}`);
        console.error(e);
    }
}

// ============================================================
// Data Processing
// ============================================================

function extractYearMonth(data) {
    const yearMonths = [];

    for (const row of data) {
        const dateStr = row["利用年月日（自）"];
        if (!dateStr || !String(dateStr).includes("/")) continue;

        const parts = String(dateStr).split("/");
        if (parts.length < 2) continue;

        const rawYear = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10);
        if (isNaN(rawYear) || isNaN(month)) continue;

        const year = normalizeYear(rawYear);
        yearMonths.push({ year, month });
    }

    if (yearMonths.length === 0) return null;

    // Return most recent year/month
    yearMonths.sort((a, b) => b.year - a.year || b.month - a.month);
    return yearMonths[0];
}

function normalizeYear(year) {
    if (year < 50) return year + 2000;
    if (year < 100) return year + 1900;
    return year;
}

function calculateDailyUsage(data, targetYear, targetMonth, targetDay) {
    let morningAmount = 0;
    let afternoonAmount = 0;
    let morningConfirmed = null;
    let afternoonConfirmed = null;

    for (const row of data) {
        const dateStr = String(row["利用年月日（自）"] || "").trim();
        const timeStr = String(row["時分（自）"] || "").trim();
        const amount = parseFloat(row["後納料金"]);

        if (isNaN(amount)) continue;
        if (!dateStr.includes("/")) continue;

        const parts = dateStr.split("/");
        if (parts.length < 3) continue;

        const csvYear = normalizeYear(parseInt(parts[0], 10));
        const csvMonth = parseInt(parts[1], 10);
        const csvDay = parseInt(parts[2], 10);

        if (isNaN(csvYear) || isNaN(csvMonth) || isNaN(csvDay)) continue;
        if (csvYear !== targetYear || csvMonth !== targetMonth || csvDay !== targetDay) {
            continue;
        }

        // Determine morning/afternoon by time
        let isMorning = true;
        if (timeStr.includes(":")) {
            const hour = parseInt(timeStr.split(":")[0], 10);
            if (!isNaN(hour)) {
                isMorning = hour < 12;
            }
        }

        if (isMorning) {
            morningAmount += amount;
            morningConfirmed = "○";
        } else {
            afternoonAmount += amount;
            afternoonConfirmed = "○";
        }
    }

    return { morningAmount, afternoonAmount, morningConfirmed, afternoonConfirmed };
}

function calculateAllUsageAmounts(data, year, month) {
    const lastDay = new Date(year, month, 0).getDate();
    const usageAmounts = {};

    for (let day = 1; day <= lastDay; day++) {
        usageAmounts[day] = calculateDailyUsage(data, year, month, day);
    }

    return usageAmounts;
}

// ============================================================
// Excel Generation (JSZip + direct XML manipulation)
// ============================================================

const XLSX_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const SHEET_PATH = "xl/worksheets/sheet1.xml";

async function generateExcel() {
    const generateBtn = document.getElementById("generate-btn");
    const loadingEl = document.getElementById("loading");

    generateBtn.disabled = true;
    loadingEl.classList.add("active");

    try {
        const settings = readSettings();

        // Load template as zip
        const response = await fetch("templates/template.xlsx");
        if (!response.ok) throw new Error("テンプレートファイルの読み込みに失敗しました");
        const zip = await JSZip.loadAsync(await response.arrayBuffer());

        // Parse sheet XML
        const sheetXml = await zip.file(SHEET_PATH).async("string");
        const parser = new DOMParser();
        const doc = parser.parseFromString(sheetXml, "application/xml");

        // Build cell modification map
        const cellMap = buildCellMap(settings, appState.year, appState.month, appState.csvData);

        // Apply modifications
        applyCellModifications(doc, cellMap);

        // Serialize back to XML
        const serializer = new XMLSerializer();
        const newXml = serializer.serializeToString(doc);
        zip.file(SHEET_PATH, newXml);

        // Remove calcChain.xml to force Excel to recalculate formulas
        zip.remove("xl/calcChain.xml");

        // Generate blob
        const blob = await zip.generateAsync({
            type: "blob",
            mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });
        updateState({ generatedBlob: blob });

        setupDownloadButton(settings.userName);
        showMessage("success", "利用実績簿が正常に生成されました！");
        document.getElementById("warnings").classList.remove("hidden");
    } catch (e) {
        showMessage("error", `実績簿の生成に失敗しました: ${e.message}`);
        console.error(e);
    } finally {
        generateBtn.disabled = false;
        loadingEl.classList.remove("active");
    }
}

function readSettings() {
    return Object.freeze({
        organization: document.getElementById("organization").value,
        position: document.getElementById("position").value,
        userName: document.getElementById("user-name").value,
        highwayFrom: document.getElementById("highway-from").value,
        highwayTo: document.getElementById("highway-to").value,
        oneWayFee: parseInt(document.getElementById("one-way-fee").value, 10) || 2680,
    });
}

function dateToExcelSerial(year, month, day) {
    const target = Date.UTC(year, month - 1, day);
    const base = Date.UTC(1900, 0, 1);
    const days = Math.round((target - base) / 86400000);
    return days + 2; // +1 for serial starting at 1, +1 for Lotus 1-2-3 bug
}

function buildCellMap(settings, year, month, csvData) {
    const cells = {};

    // Header info (inline strings)
    if (settings.organization) cells["C3"] = { type: "s", value: settings.organization };
    if (settings.position) cells["K3"] = { type: "s", value: settings.position };
    if (settings.userName) cells["N3"] = { type: "s", value: settings.userName };

    // Date info (numbers)
    cells["B5"] = { type: "n", value: year - 2018 }; // Reiwa year
    cells["D5"] = { type: "n", value: month };

    // Highway info
    cells["M5"] = { type: "s", value: settings.highwayFrom };
    cells["P5"] = { type: "s", value: settings.highwayTo };
    cells["M6"] = { type: "n", value: settings.oneWayFee };

    // Date values (overwrite formulas with serial numbers)
    const lastDayNum = new Date(year, month, 0).getDate();
    cells["E56"] = { type: "n", value: dateToExcelSerial(year, month, 1), removeFormula: true };
    cells["E57"] = { type: "n", value: dateToExcelSerial(year, month, lastDayNum), removeFormula: true };

    // Usage data
    const usageAmounts = calculateAllUsageAmounts(csvData, year, month);
    for (let day = 1; day <= lastDayNum; day++) {
        const d = usageAmounts[day];
        if (!d) continue;
        if (!d.morningAmount && !d.afternoonAmount && !d.morningConfirmed && !d.afternoonConfirmed) continue;

        let row, mC, mA, aC, aA;
        if (day <= 15) {
            row = day + 12;
            [mC, mA, aC, aA] = ["D", "E", "G", "H"];
        } else {
            row = day <= 30 ? day - 15 + 12 : 28;
            [mC, mA, aC, aA] = ["L", "M", "O", "P"];
        }

        if (d.morningConfirmed) {
            cells[`${mC}${row}`] = { type: "s", value: "○" };
            cells[`${mA}${row}`] = { type: "n", value: d.morningAmount };
        }
        if (d.afternoonConfirmed) {
            cells[`${aC}${row}`] = { type: "s", value: "○" };
            cells[`${aA}${row}`] = { type: "n", value: d.afternoonAmount };
        }
    }

    return cells;
}

function applyCellModifications(doc, cellMap) {
    const rows = doc.getElementsByTagNameNS(XLSX_NS, "row");

    for (const row of rows) {
        const cells = row.getElementsByTagNameNS(XLSX_NS, "c");
        for (const cell of cells) {
            const ref = cell.getAttribute("r");
            if (!cellMap[ref]) continue;

            const mod = cellMap[ref];
            const styleAttr = cell.getAttribute("s"); // preserve style

            // Clear existing children (v, f, is)
            while (cell.firstChild) cell.removeChild(cell.firstChild);

            // Restore style
            if (styleAttr) cell.setAttribute("s", styleAttr);

            if (mod.type === "s") {
                // Inline string
                cell.setAttribute("t", "inlineStr");
                const is = doc.createElementNS(XLSX_NS, "is");
                const t = doc.createElementNS(XLSX_NS, "t");
                t.textContent = mod.value;
                is.appendChild(t);
                cell.appendChild(is);
            } else {
                // Number
                cell.removeAttribute("t");
                const v = doc.createElementNS(XLSX_NS, "v");
                v.textContent = String(mod.value);
                cell.appendChild(v);
            }

            delete cellMap[ref]; // mark as processed
        }
    }
}

function setupDownloadButton(userName) {
    const downloadSection = document.getElementById("download-section");
    downloadSection.classList.remove("hidden");

    const downloadBtn = document.getElementById("download-btn");
    downloadBtn.onclick = () => {
        if (!appState.generatedBlob) return;

        const monthStr = String(appState.month).padStart(2, "0");
        const filename = `${appState.year}_${monthStr}_高速道路等利用実績簿（${userName}）.xlsx`;
        saveAs(appState.generatedBlob, filename);
    };
}

// ============================================================
// UI Helpers
// ============================================================

function showPreview(data, headers) {
    const container = document.getElementById("preview-table");

    const headerCells = headers.map((h) => `<th>${escapeHtml(h)}</th>`).join("");
    const rows = data
        .map((row) => {
            const cells = headers
                .map((h) => `<td>${escapeHtml(String(row[h] || ""))}</td>`)
                .join("");
            return `<tr>${cells}</tr>`;
        })
        .join("");

    container.innerHTML = `
        <div class="preview-table-wrapper">
            <table>
                <thead><tr>${headerCells}</tr></thead>
                <tbody>${rows}</tbody>
            </table>
        </div>`;
}

function showStats(data) {
    const oneWayFee = parseInt(document.getElementById("one-way-fee").value, 10) || 2680;

    let totalFee = 0;
    data.forEach((row) => {
        const fee = parseFloat(row["後納料金"]);
        if (!isNaN(fee)) totalFee += fee;
    });

    const expectedTrips = Math.floor(MONTHLY_ALLOWANCE / oneWayFee);

    document.getElementById("total-records").textContent = `${data.length}回`;
    document.getElementById("total-fee").textContent = `¥${totalFee.toLocaleString()}`;
    document.getElementById("expected-trips").textContent = `${expectedTrips}回`;
}

function showMessage(type, text) {
    const container = document.getElementById("messages");
    const msg = document.createElement("div");
    msg.className = `message ${type}`;
    msg.textContent = text;
    container.appendChild(msg);

    if (type !== "error") {
        setTimeout(() => msg.remove(), 10000);
    }
}

function clearMessages() {
    document.getElementById("messages").innerHTML = "";
}

function escapeHtml(text) {
    const div = document.createElement("div");
    div.textContent = text;
    return div.innerHTML;
}
