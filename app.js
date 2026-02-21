import * as pdfjsLib from "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.min.mjs";

pdfjsLib.GlobalWorkerOptions.workerSrc =
  "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.worker.min.mjs";

const THEME = {
  bg: "1E3A5F",
  bgSoft: "18324F",
  text: "FFFFFF",
  accent: "00B4D8",
  line: "4F789F",
  muted: "C5DDF4"
};

const i18n = {
  ko: {
    uploadTitle: "파일 업로드",
    uploadDesc: "PDF 또는 Word(.docx) 파일을 드래그하거나 선택해 주세요.",
    dropTitle: "파일을 여기에 놓으세요",
    dropSub: "또는 클릭해서 업로드",
    noFile: "선택된 파일이 없습니다.",
    selectedFilePrefix: "선택 파일",
    previewTitle: "PPT 미리보기",
    previewDesc: "BCG·Deloitte 스타일의 컨설팅 슬라이드로 자동 변환된 결과입니다.",
    previewEmpty: "업로드 후 미리보기가 표시됩니다.",
    downloadDesc: "파일 이름을 확인한 뒤 원하는 위치에 저장하세요.",
    fileNameLabel: "파일 이름",
    downloadBtn: "PPT 다운로드",
    processing: "파일을 분석하는 중입니다...",
    extracting: "텍스트 추출 중...",
    converting: "컨설팅 템플릿으로 변환 중...",
    done: "변환 완료: {count}장 슬라이드가 생성되었습니다.",
    unsupported: "지원하지 않는 형식입니다. PDF 또는 .docx 파일을 업로드해 주세요.",
    noText: "문서에서 텍스트를 찾지 못했습니다.",
    saving: "PPT를 생성하는 중입니다...",
    savedPicker: "선택한 위치에 저장되었습니다.",
    savedFallback: "브라우저 다운로드를 시작했습니다.",
    saveCanceled: "저장이 취소되었습니다.",
    saveError: "저장 중 오류가 발생했습니다.",
    slideLabel: "슬라이드",
    coverSuffix: "컨설팅 리포트",
    moduleTitle: "전략 모듈",
    matrixFallback: "핵심 데이터 추출 중",
    matrixQ1: "시장 매력도",
    matrixQ2: "실행 난이도",
    matrixQ3: "수익 기여도",
    matrixQ4: "전략 적합성",
    chartTitle: "Performance Indicators",
    chartL1: "Growth",
    chartL2: "Feasibility",
    chartL3: "Impact",
    chartL4: "Priority"
  },
  en: {
    uploadTitle: "Upload File",
    uploadDesc: "Drag a PDF or Word (.docx) file, or choose one.",
    dropTitle: "Drop your file here",
    dropSub: "or click to upload",
    noFile: "No file selected.",
    selectedFilePrefix: "Selected",
    previewTitle: "PPT Preview",
    previewDesc: "Auto-converted consulting slides in a BCG/Deloitte report style.",
    previewEmpty: "Preview appears after upload.",
    downloadDesc: "Confirm the file name and save to your preferred location.",
    fileNameLabel: "File Name",
    downloadBtn: "Download PPT",
    processing: "Analyzing file...",
    extracting: "Extracting text...",
    converting: "Converting to consulting template...",
    done: "Done: {count} slides created.",
    unsupported: "Unsupported file type. Please upload a PDF or .docx file.",
    noText: "No readable text was found in the document.",
    saving: "Generating PPT...",
    savedPicker: "Saved to your selected location.",
    savedFallback: "Browser download started.",
    saveCanceled: "Save canceled.",
    saveError: "An error occurred while saving.",
    slideLabel: "Slide",
    coverSuffix: "Consulting Report",
    moduleTitle: "Strategic Module",
    matrixFallback: "Extracting core signal",
    matrixQ1: "Market Attractiveness",
    matrixQ2: "Execution Complexity",
    matrixQ3: "Financial Impact",
    matrixQ4: "Strategic Fit",
    chartTitle: "Performance Indicators",
    chartL1: "Growth",
    chartL2: "Feasibility",
    chartL3: "Impact",
    chartL4: "Priority"
  }
};

const state = {
  lang: "ko",
  selectedFileName: "",
  sourceText: "",
  slides: []
};

const el = {
  langKo: document.getElementById("langKo"),
  langEn: document.getElementById("langEn"),
  fileInput: document.getElementById("fileInput"),
  dropZone: document.getElementById("dropZone"),
  selectedFile: document.getElementById("selectedFile"),
  statusText: document.getElementById("statusText"),
  previewContainer: document.getElementById("previewContainer"),
  fileNameInput: document.getElementById("fileNameInput"),
  downloadBtn: document.getElementById("downloadBtn")
};

function t(key, vars = {}) {
  let text = i18n[state.lang][key] ?? key;
  Object.entries(vars).forEach(([k, v]) => {
    text = text.replace(`{${k}}`, String(v));
  });
  return text;
}

function setStatus(message, type = "") {
  el.statusText.textContent = message;
  el.statusText.className = `status ${type}`.trim();
}

function getBaseName(fileName) {
  return fileName.replace(/\.[^.]+$/, "");
}

function ensurePptxExtension(name) {
  if (!name.trim()) return "woosik-slide-deck.pptx";
  return name.toLowerCase().endsWith(".pptx") ? name : `${name}.pptx`;
}

function updateSelectedFileLabel() {
  if (state.selectedFileName) {
    el.selectedFile.textContent = `${t("selectedFilePrefix")}: ${state.selectedFileName}`;
    return;
  }
  el.selectedFile.textContent = t("noFile");
}

function setLanguage(lang) {
  state.lang = lang;
  document.documentElement.lang = lang;

  el.langKo.classList.toggle("active", lang === "ko");
  el.langEn.classList.toggle("active", lang === "en");

  document.querySelectorAll("[data-i18n]").forEach((node) => {
    const key = node.getAttribute("data-i18n");
    node.textContent = t(key);
  });

  updateSelectedFileLabel();

  if (state.sourceText && state.selectedFileName) {
    state.slides = buildSlidesFromText(state.sourceText, state.selectedFileName);
  }
  renderPreview();
}

function splitIntoBulletCandidates(text) {
  const chunks = text
    .replace(/\r/g, "\n")
    .split(/\n+/)
    .map((line) => line.trim())
    .filter(Boolean)
    .flatMap((line) => line.split(/(?<=[.!?])\s+|[;:]/g))
    .map((line) => line.replace(/\s+/g, " ").trim())
    .filter((line) => line.length > 9);

  const unique = [];
  const seen = new Set();
  for (const line of chunks) {
    const key = line.toLowerCase();
    if (!seen.has(key)) {
      seen.add(key);
      unique.push(line);
    }
  }
  return unique;
}

function shortenLine(line, maxLen) {
  return line.length > maxLen ? `${line.slice(0, maxLen - 3)}...` : line;
}

function getQuadrantLabels() {
  return [t("matrixQ1"), t("matrixQ2"), t("matrixQ3"), t("matrixQ4")];
}

function getChartLabels() {
  return [t("chartL1"), t("chartL2"), t("chartL3"), t("chartL4")];
}

function buildMatrixFromChunk(chunk) {
  const quadrantLabels = getQuadrantLabels();
  const quadrants = quadrantLabels.map((label) => ({
    title: label,
    items: []
  }));

  chunk.forEach((line, idx) => {
    const target = quadrants[idx % 4];
    if (target.items.length < 2) {
      target.items.push(shortenLine(line, 78));
    }
  });

  quadrants.forEach((quadrant) => {
    if (quadrant.items.length === 0) {
      quadrant.items.push(t("matrixFallback"));
    }
  });

  return quadrants;
}

function buildChartValues(quadrants, slideIndex) {
  return quadrants.map((quadrant, idx) => {
    const totalLength = quadrant.items.join(" ").length;
    return 35 + ((totalLength * 3 + slideIndex * 17 + idx * 19) % 56);
  });
}

function buildSlidesFromText(text, fileName) {
  const bullets = splitIntoBulletCandidates(text);
  if (!bullets.length) return [];

  const baseTitle = getBaseName(fileName);
  const allPoints = bullets.map((line) => shortenLine(line, 95));
  const chunkSize = 8;
  const chunked = [];

  for (let i = 0; i < allPoints.length; i += chunkSize) {
    chunked.push(allPoints.slice(i, i + chunkSize));
  }

  if (!chunked.length) {
    chunked.push(allPoints.slice(0, chunkSize));
  }

  const slides = [];
  chunked.forEach((chunk, idx) => {
    const quadrants = buildMatrixFromChunk(chunk);
    slides.push({
      title:
        idx === 0
          ? `${baseTitle} | ${t("coverSuffix")}`
          : `${t("moduleTitle")} ${idx}`,
      quadrants,
      chart: {
        title: t("chartTitle"),
        labels: getChartLabels(),
        values: buildChartValues(quadrants, idx + 1)
      }
    });
  });

  if (slides.length === 1) {
    const fallbackChunk = allPoints.slice(0, chunkSize).reverse();
    const quadrants = buildMatrixFromChunk(fallbackChunk);
    slides.push({
      title: `${t("moduleTitle")} 1`,
      quadrants,
      chart: {
        title: t("chartTitle"),
        labels: getChartLabels(),
        values: buildChartValues(quadrants, 2)
      }
    });
  }

  return slides;
}

async function extractPdfText(file) {
  const data = new Uint8Array(await file.arrayBuffer());
  const loadingTask = pdfjsLib.getDocument({ data });
  const pdf = await loadingTask.promise;
  const pages = [];

  for (let i = 1; i <= pdf.numPages; i += 1) {
    const page = await pdf.getPage(i);
    const textContent = await page.getTextContent();
    const pageText = textContent.items
      .map((item) => ("str" in item ? item.str : ""))
      .join(" ")
      .replace(/\s+/g, " ")
      .trim();

    if (pageText) pages.push(pageText);
  }

  return pages.join("\n");
}

async function extractDocxText(file) {
  if (!window.mammoth) throw new Error("Mammoth not loaded");
  const arrayBuffer = await file.arrayBuffer();
  const result = await window.mammoth.extractRawText({ arrayBuffer });
  return result.value.trim();
}

function createMatrixCell(quadrant) {
  const cell = document.createElement("section");
  cell.className = "matrix-cell";

  const title = document.createElement("h5");
  title.textContent = quadrant.title;

  const list = document.createElement("ul");
  quadrant.items.forEach((item) => {
    const li = document.createElement("li");
    li.textContent = item;
    list.appendChild(li);
  });

  cell.appendChild(title);
  cell.appendChild(list);
  return cell;
}

function renderPreview() {
  el.previewContainer.innerHTML = "";

  if (!state.slides.length) {
    const placeholder = document.createElement("div");
    placeholder.className = "empty-preview";
    placeholder.textContent = t("previewEmpty");
    el.previewContainer.appendChild(placeholder);
    return;
  }

  state.slides.forEach((slide, index) => {
    const card = document.createElement("article");
    card.className = "slide-card";

    const shell = document.createElement("div");
    shell.className = "slide-shell";

    const slideIndex = document.createElement("p");
    slideIndex.className = "slide-index";
    slideIndex.textContent = `${t("slideLabel")} ${index + 1}`;

    const head = document.createElement("header");
    head.className = "slide-head";

    const logo = document.createElement("div");
    logo.className = "slide-logo";
    logo.textContent = "WSG";

    const title = document.createElement("h4");
    title.className = "slide-title";
    title.textContent = slide.title;

    head.appendChild(logo);
    head.appendChild(title);

    const matrix = document.createElement("div");
    matrix.className = "matrix-grid";
    slide.quadrants.forEach((quadrant) => {
      matrix.appendChild(createMatrixCell(quadrant));
    });

    const chartWrap = document.createElement("section");
    chartWrap.className = "chart-wrap";

    const chartTitle = document.createElement("p");
    chartTitle.className = "chart-title";
    chartTitle.textContent = slide.chart.title;

    const bars = document.createElement("div");
    bars.className = "chart-bars";
    slide.chart.values.forEach((value) => {
      const bar = document.createElement("div");
      bar.className = "chart-bar";
      bar.style.height = `${value}%`;
      bars.appendChild(bar);
    });

    const labels = document.createElement("div");
    labels.className = "chart-labels";
    slide.chart.labels.forEach((label) => {
      const span = document.createElement("span");
      span.textContent = label;
      labels.appendChild(span);
    });

    chartWrap.appendChild(chartTitle);
    chartWrap.appendChild(bars);
    chartWrap.appendChild(labels);

    shell.appendChild(slideIndex);
    shell.appendChild(head);
    shell.appendChild(matrix);
    shell.appendChild(chartWrap);
    card.appendChild(shell);

    el.previewContainer.appendChild(card);
  });
}

async function parseFileToSlides(file) {
  const lower = file.name.toLowerCase();
  let text = "";

  setStatus(t("extracting"), "warn");

  if (lower.endsWith(".pdf")) {
    text = await extractPdfText(file);
  } else if (lower.endsWith(".docx")) {
    text = await extractDocxText(file);
  } else {
    throw new Error(t("unsupported"));
  }

  if (!text.trim()) {
    throw new Error(t("noText"));
  }

  setStatus(t("converting"), "warn");
  const slides = buildSlidesFromText(text, file.name);

  if (!slides.length) {
    throw new Error(t("noText"));
  }

  return { text, slides };
}

async function handleFile(file) {
  if (!file) return;

  state.selectedFileName = file.name;
  updateSelectedFileLabel();
  setStatus(t("processing"), "warn");
  el.downloadBtn.disabled = true;

  try {
    const parsed = await parseFileToSlides(file);
    state.sourceText = parsed.text;
    state.slides = parsed.slides;

    renderPreview();
    el.fileNameInput.value = `${getBaseName(file.name)}.pptx`;
    el.downloadBtn.disabled = false;
    setStatus(t("done", { count: parsed.slides.length }), "ok");
  } catch (error) {
    state.slides = [];
    state.sourceText = "";
    renderPreview();
    const message = error instanceof Error ? error.message : t("saveError");
    setStatus(message, "warn");
  }
}

function fallbackDownload(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function drawSlideHeader(slide, data, index, rectShape) {
  slide.addShape(rectShape, {
    x: 0,
    y: 0,
    w: 13.333,
    h: 7.5,
    line: { color: THEME.bg, transparency: 100 },
    fill: { color: THEME.bg }
  });

  slide.addShape(rectShape, {
    x: 0.5,
    y: 0.38,
    w: 1.45,
    h: 0.66,
    line: { color: THEME.accent, pt: 1.5 },
    fill: { color: THEME.bgSoft }
  });

  slide.addText("WSG", {
    x: 0.5,
    y: 0.5,
    w: 1.45,
    h: 0.38,
    fontFace: "Roboto",
    color: THEME.accent,
    bold: true,
    fontSize: 18,
    align: "center"
  });

  slide.addText(data.title, {
    x: 2.2,
    y: 0.44,
    w: 9.9,
    h: 0.52,
    fontFace: "Roboto",
    color: THEME.text,
    bold: true,
    fontSize: 24,
    fit: "shrink"
  });

  slide.addText(`${t("slideLabel")} ${index + 1}`, {
    x: 11.6,
    y: 0.48,
    w: 1.2,
    h: 0.3,
    fontFace: "Roboto",
    color: THEME.muted,
    fontSize: 10,
    align: "right"
  });
}

function drawMatrix(slide, data, rectShape) {
  const matrixX = 0.5;
  const matrixY = 1.2;
  const cellW = 6.05;
  const cellH = 1.86;
  const gap = 0.23;

  data.quadrants.forEach((quadrant, idx) => {
    const col = idx % 2;
    const row = Math.floor(idx / 2);
    const x = matrixX + col * (cellW + gap);
    const y = matrixY + row * (cellH + gap);

    slide.addShape(rectShape, {
      x,
      y,
      w: cellW,
      h: cellH,
      line: { color: THEME.line, pt: 1 },
      fill: { color: THEME.bgSoft, transparency: 8 }
    });

    slide.addText(quadrant.title, {
      x: x + 0.2,
      y: y + 0.12,
      w: cellW - 0.4,
      h: 0.22,
      fontFace: "Roboto",
      color: THEME.accent,
      bold: true,
      fontSize: 11,
      fit: "shrink"
    });

    const bulletText = quadrant.items.map((item) => `- ${shortenLine(item, 70)}`).join("\n");
    slide.addText(bulletText, {
      x: x + 0.2,
      y: y + 0.42,
      w: cellW - 0.35,
      h: 1.22,
      fontFace: "Roboto",
      color: THEME.text,
      fontSize: 10,
      breakLine: true,
      valign: "top"
    });
  });
}

function drawChart(slide, chart, rectShape) {
  const chartX = 0.5;
  const chartY = 5.38;
  const chartW = 12.35;
  const chartH = 1.6;

  slide.addShape(rectShape, {
    x: chartX,
    y: chartY,
    w: chartW,
    h: chartH,
    line: { color: THEME.line, pt: 1 },
    fill: { color: THEME.bgSoft, transparency: 14 }
  });

  slide.addText(chart.title, {
    x: chartX + 0.2,
    y: chartY + 0.08,
    w: 3.4,
    h: 0.24,
    fontFace: "Roboto",
    color: THEME.accent,
    bold: true,
    fontSize: 11
  });

  const baseX = chartX + 1.45;
  const baseY = chartY + 1.28;
  const gap = 1.05;
  const barW = 0.72;

  chart.values.forEach((value, idx) => {
    const barH = 0.35 + (value / 100) * 0.72;
    const x = baseX + idx * (barW + gap);
    const y = baseY - barH;

    slide.addShape(rectShape, {
      x,
      y,
      w: barW,
      h: barH,
      line: { color: THEME.accent, transparency: 100 },
      fill: { color: THEME.accent }
    });

    slide.addText(String(value), {
      x: x - 0.06,
      y: y - 0.16,
      w: 0.85,
      h: 0.15,
      fontFace: "Roboto",
      color: THEME.text,
      fontSize: 8,
      align: "center"
    });

    slide.addText(chart.labels[idx], {
      x: x - 0.16,
      y: baseY + 0.08,
      w: 1,
      h: 0.2,
      fontFace: "Roboto",
      color: THEME.muted,
      fontSize: 8,
      align: "center"
    });
  });
}

async function exportPpt() {
  if (!state.slides.length) return;

  setStatus(t("saving"), "warn");

  const pptx = new window.PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "Woosik's Slide Generator";
  pptx.company = "Woosik's Slide Generator";
  pptx.subject = "Consulting report style deck";
  pptx.title = "Woosik Slide Deck";

  state.slides.forEach((data, idx) => {
    const slide = pptx.addSlide();
    const rectShape = pptx.ShapeType.rect;
    drawSlideHeader(slide, data, idx, rectShape);
    drawMatrix(slide, data, rectShape);
    drawChart(slide, data.chart, rectShape);
  });

  const fileName = ensurePptxExtension(el.fileNameInput.value.trim());

  try {
    const blob = await pptx.write({ outputType: "blob" });

    if (typeof window.showSaveFilePicker === "function") {
      try {
        const fileHandle = await window.showSaveFilePicker({
          suggestedName: fileName,
          types: [
            {
              description: "PowerPoint Presentation",
              accept: {
                "application/vnd.openxmlformats-officedocument.presentationml.presentation": [
                  ".pptx"
                ]
              }
            }
          ]
        });

        const writable = await fileHandle.createWritable();
        await writable.write(blob);
        await writable.close();
        setStatus(t("savedPicker"), "ok");
        return;
      } catch (pickerError) {
        if (pickerError && pickerError.name === "AbortError") {
          setStatus(t("saveCanceled"), "warn");
          return;
        }
      }
    }

    fallbackDownload(blob, fileName);
    setStatus(t("savedFallback"), "ok");
  } catch (_error) {
    setStatus(t("saveError"), "warn");
  }
}

el.dropZone.addEventListener("dragover", (event) => {
  event.preventDefault();
  el.dropZone.classList.add("dragover");
});

el.dropZone.addEventListener("dragleave", () => {
  el.dropZone.classList.remove("dragover");
});

el.dropZone.addEventListener("drop", (event) => {
  event.preventDefault();
  el.dropZone.classList.remove("dragover");
  const [file] = event.dataTransfer?.files || [];
  void handleFile(file);
});

el.fileInput.addEventListener("change", (event) => {
  const target = event.target;
  if (!(target instanceof HTMLInputElement)) return;
  const [file] = target.files || [];
  void handleFile(file);
});

el.downloadBtn.addEventListener("click", () => {
  void exportPpt();
});

el.langKo.addEventListener("click", () => setLanguage("ko"));
el.langEn.addEventListener("click", () => setLanguage("en"));

setLanguage("ko");
setStatus("");
