const MM_TO_PX = 96 / 25.4;
const PT_TO_PX = 96 / 72;
const SVG_NS = "http://www.w3.org/2000/svg";

const fonts = [
    { label: "Arial", value: "Arial, Helvetica, sans-serif" },
    { label: "Helvetica", value: "Helvetica, Arial, sans-serif" },
    { label: "Times New Roman", value: "\"Times New Roman\", Times, serif" },
    { label: "Courier New", value: "\"Courier New\", Courier, monospace" },
    { label: "Verdana", value: "Verdana, Geneva, sans-serif" },
    { label: "Futura", value: "\"Futura\", Arial, sans-serif" },
    { label: "Roboto", value: "Roboto, Arial, sans-serif" },
    { label: "Open Sans", value: "\"Open Sans\", Arial, sans-serif" }
];

const mmToPx = mm => mm * MM_TO_PX;
const ptToPx = pt => pt * PT_TO_PX;
const pxToMm = px => px / MM_TO_PX;

const widthInput = document.getElementById("labelWidth");
const heightInput = document.getElementById("labelHeight");
const cornerRadiusInput = document.getElementById("cornerRadius");
const borderThicknessInput = document.getElementById("borderThickness");
const backgroundColorInput = document.getElementById("backgroundColor");
const textColorInput = document.getElementById("textColor");
const lineSpacingInput = document.getElementById("lineSpacing");
const letterSpacingInput = document.getElementById("letterSpacing");
const autoBorderToggle = document.getElementById("autoBorder");
const autoBorderPaddingInput = document.getElementById("autoPadding");
const canvasMarginInput = document.getElementById("canvasMargin");
const linesContainer = document.getElementById("linesContainer");
const addLineBtn = document.getElementById("addLineBtn");
const exportSvgBtn = document.getElementById("exportSvgBtn");
const previewContainer = document.getElementById("previewContainer");
const labelForm = document.getElementById("labelForm");

const projectNameInput = document.getElementById("projectName");
const labelNameInput = document.getElementById("labelName");
const projectLabelList = document.getElementById("labelList");
const addProjectLabelBtn = document.getElementById("addLabelBtn");
const duplicateLabelBtn = document.getElementById("duplicateLabelBtn");
const removeLabelBtn = document.getElementById("removeLabelBtn");
const saveProjectBtn = document.getElementById("saveProjectBtn");
const openProjectBtn = document.getElementById("openProjectBtn");
const projectFileInput = document.getElementById("projectFileInput");

const PROJECT_FILE_VERSION = 1;
const ALLOWED_FONT_WEIGHTS = ["400", "500", "600", "700"];

const projectState = {
    name: "Untitled Project",
    labels: []
};

let activeLabelId = null;
let isHydratingForm = false;

function formatNumber(value, digits = 4) {
    if (!Number.isFinite(value)) {
        return "0";
    }
    const fixed = value.toFixed(digits);
    return fixed.replace(/\.?0+$/, "");
}

function sanitizeNumber(value, fallback, { min, max } = {}) {
    const num = Number(value);
    if (!Number.isFinite(num)) {
        return fallback;
    }
    let result = num;
    if (typeof min === "number" && result < min) {
        result = min;
    }
    if (typeof max === "number" && result > max) {
        result = max;
    }
    return result;
}

function sanitizeBoolean(value, fallback = false) {
    if (typeof value === "boolean") {
        return value;
    }
    if (typeof value === "string") {
        const normalized = value.trim().toLowerCase();
        if (["true", "1", "yes", "on"].includes(normalized)) {
            return true;
        }
        if (["false", "0", "no", "off"].includes(normalized)) {
            return false;
        }
    }
    return fallback;
}

function sanitizeColor(value, fallback) {
    if (typeof value === "string") {
        const trimmed = value.trim();
        if (/^#[0-9a-fA-F]{6}$/.test(trimmed)) {
            return trimmed.toLowerCase();
        }
    }
    return fallback;
}

function sanitizeLine(rawLine = {}) {
    const fallbackFont = fonts[0].value;
    const sanitized = {
        text: typeof rawLine.text === "string" ? rawLine.text : "",
        fontFamily: fonts.some(font => font.value === rawLine.fontFamily) ? rawLine.fontFamily : fallbackFont,
        fontSizePt: sanitizeNumber(rawLine.fontSizePt, 18, { min: 4 }),
        fontWeight: ALLOWED_FONT_WEIGHTS.includes(String(rawLine.fontWeight))
            ? String(rawLine.fontWeight)
            : "600"
    };
    return sanitized;
}

function sanitizeSettings(rawSettings = {}) {
    const defaults = getDefaultLabelSettings();
    const sanitized = {
        widthMm: sanitizeNumber(rawSettings.widthMm, defaults.widthMm, { min: 1 }),
        heightMm: sanitizeNumber(rawSettings.heightMm, defaults.heightMm, { min: 1 }),
        cornerRadiusMm: sanitizeNumber(rawSettings.cornerRadiusMm, defaults.cornerRadiusMm, { min: 0 }),
        borderThicknessMm: sanitizeNumber(rawSettings.borderThicknessMm, defaults.borderThicknessMm, { min: 0 }),
        canvasMarginMm: sanitizeNumber(rawSettings.canvasMarginMm, defaults.canvasMarginMm, { min: 0 }),
        autoBorder: sanitizeBoolean(rawSettings.autoBorder, defaults.autoBorder),
        autoBorderPaddingMm: sanitizeNumber(rawSettings.autoBorderPaddingMm, defaults.autoBorderPaddingMm, { min: 0 }),
        backgroundColor: sanitizeColor(rawSettings.backgroundColor, defaults.backgroundColor),
        textColor: sanitizeColor(rawSettings.textColor, defaults.textColor),
        lineSpacingMm: sanitizeNumber(rawSettings.lineSpacingMm, defaults.lineSpacingMm, { min: 0 }),
        letterSpacingMm: sanitizeNumber(rawSettings.letterSpacingMm, defaults.letterSpacingMm, { min: 0 }),
        lines: Array.isArray(rawSettings.lines) && rawSettings.lines.length
            ? rawSettings.lines.map(line => sanitizeLine(line))
            : defaults.lines.map(line => ({ ...line }))
    };
    if (!sanitized.lines.length) {
        sanitized.lines = defaults.lines.map(line => ({ ...line }));
    }
    return sanitized;
}

function sanitizeLabel(rawLabel = {}, fallbackIndex = 0, seenIds = new Set()) {
    let id = typeof rawLabel.id === "string" ? rawLabel.id.trim() : "";
    if (!id || seenIds.has(id)) {
        do {
            id = generateId();
        } while (seenIds.has(id));
    }
    seenIds.add(id);

    const name = typeof rawLabel.name === "string" ? rawLabel.name.trim() : "";
    const fallbackName = getFallbackLabelName(fallbackIndex);
    return {
        id,
        name: name || fallbackName,
        settings: sanitizeSettings(rawLabel.settings)
    };
}

let measurementCanvas = null;
let measurementContext = null;

function getMeasurementContext() {
    if (!measurementCanvas) {
        measurementCanvas = document.createElement("canvas");
        measurementCanvas.width = 1024;
        measurementCanvas.height = 1024;
    }
    if (!measurementContext) {
        measurementContext = measurementCanvas.getContext("2d");
    }
    return measurementContext;
}

function computeAutoBorderSize({ lines, lineSpacingMm, letterSpacingMm, paddingMm }) {
    const ctx = getMeasurementContext();
    const lineSpacingPx = mmToPx(lineSpacingMm);
    const letterSpacingPx = mmToPx(letterSpacingMm);
    const paddingPx = mmToPx(paddingMm);

    let maxLineWidthPx = 0;
    let totalHeightPx = 0;

    lines.forEach((line, index) => {
        const fontSizePx = ptToPx(line.fontSizePt);
        totalHeightPx += fontSizePx;
        if (index > 0) {
            totalHeightPx += lineSpacingPx;
        }

        const text = (line.text || "").trim();
        if (text.length > 0) {
            const font = `${line.fontWeight} ${fontSizePx}px ${line.fontFamily}`;
            ctx.font = font;
            const metrics = ctx.measureText(text);
            let lineWidth = metrics.width;
            if (letterSpacingPx > 0) {
                lineWidth += letterSpacingPx * Math.max(0, text.length - 1);
            }
            maxLineWidthPx = Math.max(maxLineWidthPx, lineWidth);
        } else {
            maxLineWidthPx = Math.max(maxLineWidthPx, fontSizePx * 0.6);
        }
    });

    if (lines.length === 0) {
        const defaultFontSizePx = ptToPx(18);
        totalHeightPx = defaultFontSizePx;
        maxLineWidthPx = defaultFontSizePx * 4;
    }

    if (totalHeightPx <= 0) {
        totalHeightPx = ptToPx(10);
    }

    if (maxLineWidthPx <= 0) {
        maxLineWidthPx = ptToPx(10) * 4;
    }

    const widthPx = maxLineWidthPx + paddingPx * 2;
    const heightPx = totalHeightPx + paddingPx * 2;

    return {
        widthMm: pxToMm(Math.max(widthPx, paddingPx * 2)),
        heightMm: pxToMm(Math.max(heightPx, paddingPx * 2))
    };
}

function generateId() {
    if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
        return crypto.randomUUID();
    }
    return `label-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}

function getDefaultLabelSettings() {
    return {
        widthMm: 100,
        heightMm: 50,
        cornerRadiusMm: 4,
        borderThicknessMm: 0.5,
        canvasMarginMm: 0.2,
        autoBorder: false,
        autoBorderPaddingMm: 2,
        backgroundColor: "#f2f2f2",
        textColor: "#000000",
        lineSpacingMm: 3,
        letterSpacingMm: 0,
        lines: [
            { text: "SAMPLE", fontFamily: fonts[0].value, fontSizePt: 26, fontWeight: "700" },
            { text: "IDENTIFICATION", fontFamily: fonts[3].value, fontSizePt: 18, fontWeight: "600" }
        ]
    };
}

function cloneSettings(settings) {
    const source = settings || getDefaultLabelSettings();
    return {
        ...source,
        lines: (source.lines || []).map(line => ({ ...line }))
    };
}

function createLabel({ name, settings } = {}) {
    return {
        id: generateId(),
        name: name || "",
        settings: cloneSettings(settings)
    };
}

function getFallbackLabelName(index) {
    return `Label ${index + 1}`;
}

function getActiveLabelIndex() {
    return projectState.labels.findIndex(label => label.id === activeLabelId);
}

function getActiveLabel() {
    const index = getActiveLabelIndex();
    return index >= 0 ? projectState.labels[index] : null;
}

function updateAutoBorderUIState() {
    const enabled = autoBorderToggle.checked;
    widthInput.disabled = enabled;
    heightInput.disabled = enabled;
    autoBorderPaddingInput.disabled = !enabled;
}

function createFontSelect(defaultValue) {
    const select = document.createElement("select");
    fonts.forEach(font => {
        const option = document.createElement("option");
        option.value = font.value;
        option.textContent = font.label;
        select.appendChild(option);
    });
    if (defaultValue && fonts.some(font => font.value === defaultValue)) {
        select.value = defaultValue;
    }
    select.dataset.role = "font";
    return select;
}

function createLineEditor(initial = {}) {
    const wrapper = document.createElement("div");
    wrapper.className = "line-editor";

    const header = document.createElement("header");
    const title = document.createElement("span");
    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.textContent = "Remove";
    header.append(title, removeBtn);
    wrapper.appendChild(header);

    const textLabel = document.createElement("label");
    textLabel.textContent = "Text";
    const textInput = document.createElement("input");
    textInput.type = "text";
    textInput.placeholder = "Enter line text";
    textInput.value = initial.text || "";
    textInput.dataset.role = "text";
    textLabel.appendChild(textInput);

    const fontLabel = document.createElement("label");
    fontLabel.textContent = "Font";
    const fontSelect = createFontSelect(initial.fontFamily);
    fontLabel.appendChild(fontSelect);

    const sizeLabel = document.createElement("label");
    sizeLabel.textContent = "Font Size (pt)";
    const sizeInput = document.createElement("input");
    sizeInput.type = "number";
    sizeInput.min = "4";
    sizeInput.step = "0.5";
    sizeInput.value = initial.fontSizePt != null ? initial.fontSizePt : 22;
    sizeInput.dataset.role = "size";
    sizeLabel.appendChild(sizeInput);

    const weightLabel = document.createElement("label");
    weightLabel.textContent = "Weight";
    const weightSelect = document.createElement("select");
    weightSelect.dataset.role = "weight";
    [
        { label: "Regular", value: "400" },
        { label: "Medium", value: "500" },
        { label: "Semi-bold", value: "600" },
        { label: "Bold", value: "700" }
    ].forEach(optionData => {
        const opt = document.createElement("option");
        opt.value = optionData.value;
        opt.textContent = optionData.label;
        weightSelect.appendChild(opt);
    });
    weightSelect.value = initial.fontWeight || "600";
    weightLabel.appendChild(weightSelect);

    wrapper.append(textLabel, fontLabel, sizeLabel, weightLabel);

    [textInput, fontSelect, sizeInput, weightSelect].forEach(control => {
        const eventName = control.tagName === "SELECT" ? "change" : "input";
        control.addEventListener(eventName, () => {
            if (!isHydratingForm) {
                updatePreview();
            }
        });
    });

    removeBtn.addEventListener("click", () => {
        if (linesContainer.children.length > 1) {
            linesContainer.removeChild(wrapper);
        } else {
            textInput.value = "";
        }
        refreshLineHeaders();
        if (!isHydratingForm) {
            updatePreview();
        }
    });

    return wrapper;
}

function refreshLineHeaders() {
    Array.from(linesContainer.children).forEach((line, index) => {
        const headerTitle = line.querySelector("header span");
        if (headerTitle) {
            headerTitle.textContent = `Line ${index + 1}`;
        }
    });
}

function addLine(initial = {}, options = {}) {
    const editor = createLineEditor(initial);
    linesContainer.appendChild(editor);
    refreshLineHeaders();
    if (!isHydratingForm && !options.suppressPreview) {
        updatePreview();
    }
    return editor;
}

function applySettingsToForm(settings) {
    widthInput.value = settings.widthMm != null ? settings.widthMm : "";
    heightInput.value = settings.heightMm != null ? settings.heightMm : "";
    cornerRadiusInput.value = settings.cornerRadiusMm != null ? settings.cornerRadiusMm : 0;
    borderThicknessInput.value = settings.borderThicknessMm != null ? settings.borderThicknessMm : 0;
    canvasMarginInput.value = settings.canvasMarginMm != null ? settings.canvasMarginMm : 0;
    autoBorderToggle.checked = !!settings.autoBorder;
    autoBorderPaddingInput.value = settings.autoBorderPaddingMm != null ? settings.autoBorderPaddingMm : 0;
    updateAutoBorderUIState();
    backgroundColorInput.value = settings.backgroundColor || "#ffffff";
    textColorInput.value = settings.textColor || "#000000";
    lineSpacingInput.value = settings.lineSpacingMm != null ? settings.lineSpacingMm : 3;
    letterSpacingInput.value = settings.letterSpacingMm != null ? settings.letterSpacingMm : 0;

    linesContainer.innerHTML = "";
    const lines = settings.lines && settings.lines.length ? settings.lines : [{ text: "", fontFamily: fonts[0].value, fontSizePt: 22, fontWeight: "600" }];
    lines.forEach(line => addLine(line, { suppressPreview: true }));
}

function renderLabelList() {
    projectLabelList.innerHTML = "";
    projectState.labels.forEach((label, index) => {
        const li = document.createElement("li");
        li.className = "label-list-item";
        if (label.id === activeLabelId) {
            li.classList.add("active");
        }

        const button = document.createElement("button");
        button.type = "button";
        button.dataset.labelId = label.id;
        const displayName = label.name && label.name.trim().length > 0 ? label.name.trim() : getFallbackLabelName(index);
        const nameSpan = document.createElement("span");
        nameSpan.textContent = displayName;
        const metaSpan = document.createElement("span");
        metaSpan.className = "meta";
        if (label.settings) {
            const width = Math.round(label.settings.widthMm || 0);
            const height = Math.round(label.settings.heightMm || 0);
            metaSpan.textContent = `${width} × ${height} mm`;
        } else {
            metaSpan.textContent = "";
        }
        button.append(nameSpan, metaSpan);
        li.appendChild(button);
        projectLabelList.appendChild(li);
    });
    updateLabelActionState();
}

function updateLabelActionState() {
    const hasActive = !!getActiveLabel();
    duplicateLabelBtn.disabled = !hasActive;
    removeLabelBtn.disabled = projectState.labels.length <= 1 || !hasActive;
}

function selectLabel(labelId) {
    const label = projectState.labels.find(item => item.id === labelId);
    if (!label) {
        return;
    }
    activeLabelId = labelId;

    isHydratingForm = true;
    projectNameInput.value = projectState.name;
    const index = projectState.labels.findIndex(item => item.id === labelId);
    const displayName = label.name && label.name.trim().length > 0 ? label.name.trim() : getFallbackLabelName(index);
    labelNameInput.value = displayName;
    label.name = displayName;
    applySettingsToForm(label.settings);
    isHydratingForm = false;

    refreshLineHeaders();
    renderLabelList();
    updatePreview();
}

function getLabelSettings() {
    let widthMm = parseFloat(widthInput.value);
    let heightMm = parseFloat(heightInput.value);
    const cornerRadiusMm = Math.max(0, parseFloat(cornerRadiusInput.value) || 0);
    const borderThicknessMm = Math.max(0, parseFloat(borderThicknessInput.value) || 0);
    const canvasMarginMm = Math.max(0, parseFloat(canvasMarginInput.value) || 0);
    const backgroundColor = backgroundColorInput.value || "#ffffff";
    const textColor = textColorInput.value || "#000000";
    const lineSpacingMm = Math.max(0, parseFloat(lineSpacingInput.value) || 0);
    const letterSpacingMm = Math.max(0, parseFloat(letterSpacingInput.value) || 0);
    const autoBorder = !!autoBorderToggle.checked;
    const autoBorderPaddingMm = Math.max(0, parseFloat(autoBorderPaddingInput.value) || 0);

    const lines = Array.from(linesContainer.children).map(line => {
        const textInput = line.querySelector("input[data-role='text']");
        const fontSelect = line.querySelector("select[data-role='font']");
        const sizeInput = line.querySelector("input[data-role='size']");
        const weightSelect = line.querySelector("select[data-role='weight']");

        const rawFontSize = parseFloat(sizeInput ? sizeInput.value : 18);
        const fontSizePt = Number.isFinite(rawFontSize) ? Math.max(4, rawFontSize) : 18;

        return {
            text: textInput ? textInput.value || "" : "",
            fontFamily: fontSelect ? fontSelect.value : fonts[0].value,
            fontSizePt,
            fontWeight: weightSelect ? weightSelect.value : "400"
        };
    });

    if (autoBorder) {
        const autoSize = computeAutoBorderSize({
            lines,
            lineSpacingMm,
            letterSpacingMm,
            paddingMm: autoBorderPaddingMm
        });
        if (Number.isFinite(autoSize.widthMm) && autoSize.widthMm > 0) {
            widthMm = autoSize.widthMm;
        }
        if (Number.isFinite(autoSize.heightMm) && autoSize.heightMm > 0) {
            heightMm = autoSize.heightMm;
        }
        widthInput.value = formatNumber(widthMm, 3);
        heightInput.value = formatNumber(heightMm, 3);
    }

    if (!Number.isFinite(widthMm) || widthMm <= 0 || !Number.isFinite(heightMm) || heightMm <= 0) {
        return null;
    }

    return {
        widthMm,
        heightMm,
        cornerRadiusMm,
        borderThicknessMm,
        canvasMarginMm,
        autoBorder,
        autoBorderPaddingMm,
        backgroundColor,
        textColor,
        lineSpacingMm,
        letterSpacingMm,
        lines
    };
}

function persistActiveLabel(settings) {
    const activeLabel = getActiveLabel();
    if (!activeLabel) {
        return;
    }
    activeLabel.settings = cloneSettings(settings);
    activeLabel.name = labelNameInput.value.trim();
    renderLabelList();
}

function buildProjectPayload() {
    const settings = getLabelSettings();
    if (settings) {
        persistActiveLabel(settings);
    }

    const projectName = projectState.name && projectState.name.trim().length
        ? projectState.name.trim()
        : "Untitled Project";

    const labels = projectState.labels.map(label => ({
        id: label.id,
        name: label.name,
        settings: cloneSettings(label.settings)
    }));

    return {
        version: PROJECT_FILE_VERSION,
        project: {
            name: projectName,
            labels
        }
    };
}

function downloadProjectFile() {
    const payload = buildProjectPayload();
    const json = JSON.stringify(payload, null, 2);
    const blob = new Blob([json], { type: "application/json;charset=utf-8" });
    const fileName = `${slugify(payload.project.name)}_project.json`;
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setTimeout(() => URL.revokeObjectURL(link.href), 0);
}

function loadProjectFromData(rawData) {
    if (!rawData || typeof rawData !== "object") {
        throw new Error("Invalid project file.");
    }

    if (typeof rawData.version === "number" && rawData.version > PROJECT_FILE_VERSION) {
        console.warn("Project file version is newer than the current application version.");
    }

    const payload = rawData.project && typeof rawData.project === "object" ? rawData.project : rawData;
    const rawLabels = Array.isArray(payload.labels) ? payload.labels : [];
    const seenIds = new Set();
    const sanitizedLabels = rawLabels.map((label, index) => sanitizeLabel(label, index, seenIds));

    if (!sanitizedLabels.length) {
        sanitizedLabels.push(createLabel({ name: "Label 1", settings: getDefaultLabelSettings() }));
    }

    projectState.name = payload.name && typeof payload.name === "string" && payload.name.trim().length
        ? payload.name.trim()
        : "Untitled Project";

    projectState.labels = sanitizedLabels.map(label => ({
        id: label.id,
        name: label.name,
        settings: cloneSettings(label.settings)
    }));

    projectNameInput.value = projectState.name;
    selectLabel(projectState.labels[0].id);
}

function handleProjectFileSelected(event) {
    const file = event.target.files && event.target.files[0];
    if (!file) {
        return;
    }

    const reader = new FileReader();
    reader.onload = () => {
        try {
            const content = typeof reader.result === "string" ? reader.result : "";
            const data = JSON.parse(content);
            loadProjectFromData(data);
        } catch (error) {
            console.error(error);
            alert("Unable to open project file. Please select a valid project export.");
        } finally {
            projectFileInput.value = "";
        }
    };
    reader.onerror = () => {
        alert("Unable to read the selected project file.");
        projectFileInput.value = "";
    };
    reader.readAsText(file);
}

function buildSvg(settings) {
    const widthPx = mmToPx(settings.widthMm);
    const heightPx = mmToPx(settings.heightMm);
    const cornerRadiusPx = mmToPx(settings.cornerRadiusMm);
    const borderThicknessPx = mmToPx(settings.borderThicknessMm);
    const lineSpacingPx = mmToPx(settings.lineSpacingMm);
    const letterSpacingPx = mmToPx(settings.letterSpacingMm);
    const canvasMarginMm = Math.max(0, settings.canvasMarginMm || 0);
    const canvasMarginPx = mmToPx(canvasMarginMm);
    const canvasWidthMm = settings.widthMm + canvasMarginMm * 2;
    const canvasHeightMm = settings.heightMm + canvasMarginMm * 2;
    const canvasWidthPx = widthPx + canvasMarginPx * 2;
    const canvasHeightPx = heightPx + canvasMarginPx * 2;
    const originX = canvasMarginPx;
    const originY = canvasMarginPx;

    const svg = document.createElementNS(SVG_NS, "svg");
    svg.setAttribute("xmlns", SVG_NS);
    svg.setAttribute("width", `${canvasWidthMm}mm`);
    svg.setAttribute("height", `${canvasHeightMm}mm`);
    svg.setAttribute("viewBox", `0 0 ${formatNumber(canvasWidthPx)} ${formatNumber(canvasHeightPx)}`);
    svg.style.overflow = "visible";
    svg.setAttribute("role", "img");

    const rect = document.createElementNS(SVG_NS, "rect");
    const strokeInsetPx = borderThicknessPx > 0 ? borderThicknessPx / 2 : 0;
    const rectWidthPx = Math.max(0, widthPx - borderThicknessPx);
    const rectHeightPx = Math.max(0, heightPx - borderThicknessPx);
    const rectRadiusPx = Math.max(0, cornerRadiusPx - strokeInsetPx);

    rect.setAttribute("x", formatNumber(originX + strokeInsetPx, 4));
    rect.setAttribute("y", formatNumber(originY + strokeInsetPx, 4));
    rect.setAttribute("width", formatNumber(rectWidthPx, 4));
    rect.setAttribute("height", formatNumber(rectHeightPx, 4));
    rect.setAttribute("rx", formatNumber(rectRadiusPx, 4));
    rect.setAttribute("ry", formatNumber(rectRadiusPx, 4));
    rect.setAttribute("fill", settings.backgroundColor);

    if (borderThicknessPx > 0) {
        rect.setAttribute("stroke", settings.textColor);
        rect.setAttribute("stroke-width", formatNumber(borderThicknessPx, 4));
    } else {
        rect.setAttribute("stroke", "none");
    }

    svg.appendChild(rect);

    if (settings.lines.length > 0) {
        const lineCount = settings.lines.length;
        const fontSizesPx = settings.lines.map(line => ptToPx(line.fontSizePt));
        const totalTextHeightPx = fontSizesPx.reduce((acc, size, index) => {
            return acc + size + (index > 0 ? lineSpacingPx : 0);
        }, 0);

        let baselineOffset = originY + (heightPx - totalTextHeightPx) / 2;

        if (lineCount % 2 === 0) {
            const gapIndex = lineCount / 2;
            const sumFontBeforeGap = fontSizesPx.reduce((acc, size, index) => {
                return index < gapIndex ? acc + size : acc;
            }, 0);
            const spacingBeforeGap = lineSpacingPx * Math.max(0, gapIndex - 1);
            const baselineAtGap = baselineOffset + sumFontBeforeGap + spacingBeforeGap;
            const centerY = originY + heightPx / 2;
            const gapCenter = baselineAtGap + lineSpacingPx / 2;
            const delta = centerY - gapCenter;
            baselineOffset += delta;
        } else {
            const middleIndex = Math.floor(lineCount / 2);
            const sumFontThroughMiddle = fontSizesPx.reduce((acc, size, index) => {
                return index <= middleIndex ? acc + size : acc;
            }, 0);
            const spacingBeforeMiddle = lineSpacingPx * middleIndex;
            const baselineAtMiddle = baselineOffset + sumFontThroughMiddle + spacingBeforeMiddle;
            const centerY = originY + heightPx / 2;
            const middleLineCenter = baselineAtMiddle - fontSizesPx[middleIndex] / 2;
            const delta = centerY - middleLineCenter;
            baselineOffset += delta;
        }

        let currentBaseline = baselineOffset;

        settings.lines.forEach((line, index) => {
            const fontSizePx = fontSizesPx[index];
            currentBaseline += fontSizePx;
            const textContent = (line.text || "").trim();

            if (textContent.length > 0) {
                const textElement = document.createElementNS(SVG_NS, "text");
                textElement.setAttribute("x", formatNumber(originX + widthPx / 2, 4));
                textElement.setAttribute("y", formatNumber(currentBaseline, 4));
                textElement.setAttribute("fill", settings.textColor);
                textElement.setAttribute("font-size", formatNumber(fontSizePx, 4));
                textElement.setAttribute("font-family", line.fontFamily);
                textElement.setAttribute("font-weight", line.fontWeight);
                textElement.setAttribute("text-anchor", "middle");
                textElement.setAttribute("dominant-baseline", "alphabetic");
                if (letterSpacingPx > 0) {
                    textElement.setAttribute("letter-spacing", formatNumber(letterSpacingPx, 4));
                } else {
                    textElement.removeAttribute("letter-spacing");
                }
                textElement.textContent = textContent;
                svg.appendChild(textElement);
            }

            currentBaseline += lineSpacingPx;
        });
    }

    return svg;
}

function updatePreview() {
    if (isHydratingForm) {
        return;
    }

    const settings = getLabelSettings();
    if (!settings) {
        previewContainer.textContent = "Enter valid width and height to preview the label.";
        return;
    }

    const svg = buildSvg(settings);
    previewContainer.innerHTML = "";
    previewContainer.appendChild(svg);

    persistActiveLabel(settings);
}

function slugify(value) {
    return value
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, "-")
        .replace(/^-+|-+$/g, "") || "label";
}

function exportSvg() {
    const settings = getLabelSettings();
    if (!settings) {
        alert("Please provide valid label dimensions before exporting.");
        return;
    }

    const activeLabel = getActiveLabel();
    const labelIndex = getActiveLabelIndex();
    const labelName = activeLabel && activeLabel.name && activeLabel.name.trim().length > 0
        ? activeLabel.name.trim()
        : getFallbackLabelName(labelIndex >= 0 ? labelIndex : 0);
    const fileSlug = slugify(labelName);

    const svg = buildSvg(settings);
    const serializer = new XMLSerializer();
    let source = serializer.serializeToString(svg);

    if (!source.startsWith("<?xml")) {
        source = `<?xml version="1.0" encoding="UTF-8"?>\n${source}`;
    }

    const blob = new Blob([source], { type: "image/svg+xml;charset=utf-8" });
    const fileName = `${fileSlug}_${Math.round(settings.widthMm)}x${Math.round(settings.heightMm)}mm.svg`;
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setTimeout(() => URL.revokeObjectURL(link.href), 0);
}

function handleAddLabel() {
    const activeLabel = getActiveLabel();
    const baseSettings = activeLabel ? cloneSettings(activeLabel.settings) : getDefaultLabelSettings();
    baseSettings.lines = (baseSettings.lines || []).map(line => ({ ...line, text: "" }));
    if (baseSettings.lines.length === 0) {
        baseSettings.lines.push({ text: "", fontFamily: fonts[0].value, fontSizePt: 22, fontWeight: "600" });
    }

    const newLabel = createLabel({
        name: "",
        settings: baseSettings
    });

    projectState.labels.push(newLabel);
    selectLabel(newLabel.id);
    labelNameInput.focus();
    labelNameInput.select();
}

function handleDuplicateLabel() {
    const activeLabel = getActiveLabel();
    if (!activeLabel) {
        return;
    }
    const index = getActiveLabelIndex();
    const duplicate = createLabel({
        name: `${activeLabel.name || getFallbackLabelName(index)} Copy`,
        settings: activeLabel.settings
    });
    projectState.labels.splice(index + 1, 0, duplicate);
    selectLabel(duplicate.id);
}

function handleRemoveLabel() {
    if (projectState.labels.length <= 1) {
        return;
    }
    const index = getActiveLabelIndex();
    if (index < 0) {
        return;
    }
    projectState.labels.splice(index, 1);
    const nextIndex = index >= projectState.labels.length ? projectState.labels.length - 1 : index;
    const nextLabel = projectState.labels[nextIndex];
    if (nextLabel) {
        selectLabel(nextLabel.id);
    } else {
        updateLabelActionState();
        previewContainer.textContent = "Add a label to begin.";
    }
}

function attachGlobalListeners() {
    [
        widthInput,
        heightInput,
        cornerRadiusInput,
        borderThicknessInput,
        canvasMarginInput,
        backgroundColorInput,
        textColorInput,
        lineSpacingInput,
        letterSpacingInput,
        autoBorderPaddingInput
    ].forEach(input => {
        input.addEventListener("input", () => {
            if (!isHydratingForm) {
                updatePreview();
            }
        });
    });

    addLineBtn.addEventListener("click", () => addLine({ text: "" }));
    exportSvgBtn.addEventListener("click", exportSvg);

    projectLabelList.addEventListener("click", event => {
        const button = event.target.closest("button[data-label-id]");
        if (!button) {
            return;
        }
        const labelId = button.dataset.labelId;
        if (labelId && labelId !== activeLabelId) {
            selectLabel(labelId);
        }
    });

    addProjectLabelBtn.addEventListener("click", handleAddLabel);
    duplicateLabelBtn.addEventListener("click", handleDuplicateLabel);
    removeLabelBtn.addEventListener("click", handleRemoveLabel);
    saveProjectBtn.addEventListener("click", downloadProjectFile);
    openProjectBtn.addEventListener("click", () => projectFileInput.click());
    projectFileInput.addEventListener("change", handleProjectFileSelected);
    autoBorderToggle.addEventListener("change", () => {
        updateAutoBorderUIState();
        if (!isHydratingForm) {
            updatePreview();
        }
    });

    projectNameInput.addEventListener("input", () => {
        projectState.name = projectNameInput.value;
    });

    labelNameInput.addEventListener("input", () => {
        if (isHydratingForm) {
            return;
        }
        const activeLabel = getActiveLabel();
        if (!activeLabel) {
            return;
        }
        activeLabel.name = labelNameInput.value.trim();
        renderLabelList();
    });

    labelNameInput.addEventListener("blur", () => {
        const activeLabel = getActiveLabel();
        const index = getActiveLabelIndex();
        if (!activeLabel || index < 0) {
            return;
        }
        const trimmed = labelNameInput.value.trim();
        if (trimmed.length === 0) {
            const fallback = getFallbackLabelName(index);
            activeLabel.name = fallback;
            labelNameInput.value = fallback;
            renderLabelList();
        }
    });

    labelForm.addEventListener("reset", event => {
        event.preventDefault();
        const activeLabel = getActiveLabel();
        if (!activeLabel) {
            return;
        }
        isHydratingForm = true;
        applySettingsToForm(activeLabel.settings);
        labelNameInput.value = activeLabel.name && activeLabel.name.trim().length > 0
            ? activeLabel.name.trim()
            : getFallbackLabelName(getActiveLabelIndex());
        isHydratingForm = false;
        updatePreview();
    });
}

function init() {
    attachGlobalListeners();
    const initialLabel = createLabel({ name: "Label 1", settings: getDefaultLabelSettings() });
    projectState.labels.push(initialLabel);
    selectLabel(initialLabel.id);
}

if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
} else {
    init();
}
