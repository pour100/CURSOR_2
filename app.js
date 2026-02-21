import * as pdfjsLib from "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.min.mjs";

pdfjsLib.GlobalWorkerOptions.workerSrc =
  "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.worker.min.mjs";

const i18n = {
  ko: {
    uploadTitle: "파일 업로드",
    uploadDesc: "PDF 또는 Word(.docx) 파일을 드래그하거나 선택해 주세요.",
    dropTitle: "파일을 여기에 놓으세요",
    dropSub: "또는 클릭해서 업로드",
    noFile: "선택된 파일이 없습니다.",
    selectedFilePrefix: "선택 파일",
    previewTitle: "PPT 미리보기",
    previewDesc: "맥킨지 스타일의 슬라이드 구조로 자동 변환된 결과를 확인하세요.",
    previewEmpty: "업로드 후 미리보기가 표시됩니다.",
    downloadDesc: "파일 이름을 확인한 뒤 원하는 위치에 저장하세요.",
    fileNameLabel: "파일 이름",
    downloadBtn: "PPT 다운로드",
    processing: "파일을 분석하는 중입니다...",
    extracting: "텍스트 추출 중...",
    converting: "슬라이드로 변환 중...",
    done: "변환 완료: {count}장 슬라이드가 생성되었습니다.",
    unsupported: "지원하지 않는 형식입니다. PDF 또는 .docx 파일을 업로드해 주세요.",
    noText: "문서에서 텍스트를 찾지 못했습니다.",
    downloadReady: "다운로드 준비 완료",
    saving: "PPT를 생성하는 중입니다...",
    savedPicker: "선택한 위치에 저장되었습니다.",
    savedFallback: "브라우저 다운로드를 시작했습니다.",
    saveCanceled: "저장이 취소되었습니다.",
    saveError: "저장 중 오류가 발생했습니다.",
    slideLabel: "슬라이드",
    coverSuffix: "핵심 요약",
    insightTitle: "핵심 인사이트",
    previewOnly: "미리보기 전용 텍스트입니다."
  },
  en: {
    uploadTitle: "Upload File",
    uploadDesc: "Drag a PDF or Word (.docx) file, or choose one.",
    dropTitle: "Drop your file here",
    dropSub: "or click to upload",
    noFile: "No file selected.",
    selectedFilePrefix: "Selected",
    previewTitle: "PPT Preview",
    previewDesc: "Review the auto-converted deck in a McKinsey-style structure.",
    previewEmpty: "Preview appears after upload.",
    downloadDesc: "Confirm the file name and save to your preferred location.",
    fileNameLabel: "File Name",
    downloadBtn: "Download PPT",
    processing: "Analyzing file...",
    extracting: "Extracting text...",
    converting: "Converting to slides...",
    done: "Done: {count} slides created.",
    unsupported: "Unsupported file type. Please upload a PDF or .docx file.",
    noText: "No readable text was found in the document.",
    downloadReady: "Ready to download",
    saving: "Generating PPT...",
    savedPicker: "Saved to your selected location.",
    savedFallback: "Browser download started.",
    saveCanceled: "Save canceled.",
    saveError: "An error occurred while saving.",
    slideLabel: "Slide",
    coverSuffix: "Executive Summary",
    insightTitle: "Key Insights",
    previewOnly: "Preview-only text."
  }
};

const state = {
  lang: "ko",
  selectedFileName: "",
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
  if (!name.trim()) {
    return "woosik-slide-deck.pptx";
  }
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
  document.documentElement.lang = lang === "ko" ? "ko" : "en";

  el.langKo.classList.toggle("active", lang === "ko");
  el.langEn.classList.toggle("active", lang === "en");

  document.querySelectorAll("[data-i18n]").forEach((node) => {
    const key = node.getAttribute("data-i18n");
    node.textContent = t(key);
  });

  updateSelectedFileLabel();
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
    .filter((line) => line.length > 8);

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

function buildSlidesFromText(text, fileName) {
  const bullets = splitIntoBulletCandidates(text);
  if (bullets.length === 0) {
    return [];
  }

  const baseTitle = getBaseName(fileName);
  const coverBullets = bullets.slice(0, 4).map((line) => shortenLine(line, 105));
  const contentBullets = bullets.slice(4).map((line) => shortenLine(line, 120));

  const slides = [
    {
      title: `${baseTitle} - ${t("coverSuffix")}`,
      bullets: coverBullets.length ? coverBullets : [t("previewOnly")]
    }
  ];

  const chunkSize = 5;
  for (let i = 0; i < contentBullets.length; i += chunkSize) {
    const group = contentBullets.slice(i, i + chunkSize);
    slides.push({
      title: `${t("insightTitle")} ${Math.floor(i / chunkSize) + 1}`,
      bullets: group
    });
  }

  if (slides.length === 1) {
    slides.push({
      title: `${t("insightTitle")} 1`,
      bullets: bullets.slice(0, 5).map((line) => shortenLine(line, 120))
    });
  }

  return slides;
}

function shortenLine(line, max) {
  return line.length > max ? `${line.slice(0, max - 1)}...` : line;
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

    if (pageText) {
      pages.push(pageText);
    }
  }
  return pages.join("\n");
}

async function extractDocxText(file) {
  if (!window.mammoth) {
    throw new Error("Mammoth not loaded");
  }
  const arrayBuffer = await file.arrayBuffer();
  const result = await window.mammoth.extractRawText({ arrayBuffer });
  return result.value.trim();
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

    const top = document.createElement("div");
    top.className = "slide-top";

    const inner = document.createElement("div");
    inner.className = "slide-inner";

    const idx = document.createElement("p");
    idx.className = "slide-index";
    idx.textContent = `${t("slideLabel")} ${index + 1}`;

    const title = document.createElement("h4");
    title.className = "slide-title";
    title.textContent = slide.title;

    const list = document.createElement("ul");
    list.className = "slide-bullets";
    slide.bullets.forEach((bullet) => {
      const li = document.createElement("li");
      li.textContent = bullet;
      list.appendChild(li);
    });

    inner.appendChild(idx);
    inner.appendChild(title);
    inner.appendChild(list);
    card.appendChild(top);
    card.appendChild(inner);
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

  return slides;
}

async function handleFile(file) {
  if (!file) return;

  state.selectedFileName = file.name;
  updateSelectedFileLabel();
  setStatus(t("processing"), "warn");
  el.downloadBtn.disabled = true;

  try {
    const slides = await parseFileToSlides(file);
    state.slides = slides;
    renderPreview();
    el.fileNameInput.value = `${getBaseName(file.name)}.pptx`;
    el.downloadBtn.disabled = false;
    setStatus(t("done", { count: slides.length }), "ok");
  } catch (error) {
    state.slides = [];
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

async function exportPpt() {
  if (!state.slides.length) return;

  setStatus(t("saving"), "warn");

  const pptx = new window.PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "Woosik's Slide Generator";
  pptx.company = "Woosik's Slide Generator";
  pptx.subject = "Auto-generated presentation";
  pptx.title = "Woosik Slide Deck";

  state.slides.forEach((slideData, idx) => {
    const slide = pptx.addSlide();
    slide.background = { color: "FFFFFF" };

    slide.addShape(pptx.ShapeType.rect, {
      x: 0,
      y: 0,
      w: 13.333,
      h: 0.18,
      line: { color: "0F4C81", transparency: 100 },
      fill: { color: "0F4C81" }
    });

    slide.addText(slideData.title, {
      x: 0.7,
      y: 0.45,
      w: 11.9,
      h: 0.7,
      fontFace: "Calibri",
      color: "1A1F2B",
      bold: true,
      fontSize: 28
    });

    const bullets = slideData.bullets.map((line) => `• ${line}`).join("\n");
    slide.addText(bullets, {
      x: 0.9,
      y: 1.45,
      w: 11.7,
      h: 4.9,
      fontFace: "Calibri",
      color: "2F3745",
      fontSize: 18,
      valign: "top",
      breakLine: true
    });

    slide.addText(String(idx + 1), {
      x: 12.2,
      y: 7.0,
      w: 0.8,
      h: 0.3,
      color: "5A6475",
      fontFace: "Calibri",
      fontSize: 10,
      align: "right"
    });
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
