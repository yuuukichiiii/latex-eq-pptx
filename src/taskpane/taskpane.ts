import katex from 'katex';
import { toPng } from 'html-to-image';
import './taskpane.css';
import 'katex/dist/katex.min.css';

// PowerPoint.ShapeCollection.addImage was added in PowerPointApi 1.8.
// The type definitions lag behind, so we extend the type locally.
type ShapeCollectionWithImage = PowerPoint.ShapeCollection & {
  addImage(base64: string, options?: Record<string, unknown>): PowerPoint.Shape;
};

// ---- Constants ----
const HISTORY_KEY = 'latex-eq-history';
const MAX_HISTORY = 10;
const DEFAULT_SCALE = 1.5;
// Render at 4× for crisp output when scaled in PowerPoint
const CAPTURE_PIXEL_RATIO = 4;

// ---- State ----
let isDisplayMode = true;

// ---- DOM references ----
const latexInput = document.getElementById('latex-input') as HTMLTextAreaElement;
const scaleSlider = document.getElementById('scale-slider') as HTMLInputElement;
const scaleDisplay = document.getElementById('scale-display') as HTMLSpanElement;
const colorPicker = document.getElementById('color-picker') as HTMLInputElement;
const colorLabel = document.getElementById('color-label') as HTMLSpanElement;
const previewRender = document.getElementById('preview-render') as HTMLDivElement;
const previewPlaceholder = document.getElementById('preview-placeholder') as HTMLParagraphElement;
const errorMsg = document.getElementById('error-msg') as HTMLParagraphElement;
const insertBtn = document.getElementById('insert-btn') as HTMLButtonElement;
const statusMsg = document.getElementById('status-msg') as HTMLParagraphElement;
const displayModeBtn = document.getElementById('display-mode-btn') as HTMLButtonElement;
const inlineModeBtn = document.getElementById('inline-mode-btn') as HTMLButtonElement;
const historySection = document.getElementById('history-section') as HTMLElement;
const historyList = document.getElementById('history-list') as HTMLDivElement;
const clearHistoryBtn = document.getElementById('clear-history-btn') as HTMLButtonElement;

// ---- Preview rendering ----
function renderPreview(): void {
  const latex = latexInput.value.trim();
  const scale = parseFloat(scaleSlider.value);
  const color = colorPicker.value;

  errorMsg.textContent = '';
  statusMsg.textContent = '';

  if (!latex) {
    previewRender.innerHTML = '';
    previewPlaceholder.style.display = 'block';
    insertBtn.disabled = true;
    return;
  }

  previewPlaceholder.style.display = 'none';

  try {
    katex.render(latex, previewRender, {
      displayMode: isDisplayMode,
      throwOnError: true,
      output: 'html',
      trust: false,
    });
    previewRender.style.fontSize = `${scale * 16}px`;
    previewRender.style.color = color;
    insertBtn.disabled = false;
  } catch (e: unknown) {
    previewRender.innerHTML = '';
    errorMsg.textContent = e instanceof Error ? e.message : String(e);
    insertBtn.disabled = true;
  }
}

// ---- Capture rendered equation as PNG (base64) ----
async function captureAsPng(): Promise<string> {
  // Temporarily pin size for stable capture
  const savedOverflow = previewRender.style.overflow;
  previewRender.style.overflow = 'visible';

  try {
    const dataUrl = await toPng(previewRender, {
      pixelRatio: CAPTURE_PIXEL_RATIO,
      // omitting backgroundColor → transparent (html-to-image default)
      style: {
        display: 'inline-block',
        padding: '8px',
      },
    });
    return dataUrl;
  } finally {
    previewRender.style.overflow = savedOverflow;
  }
}

// ---- Insert into PowerPoint ----
async function insertEquation(): Promise<void> {
  const latex = latexInput.value.trim();
  if (!latex) return;

  insertBtn.disabled = true;
  insertBtn.textContent = '挿入中...';
  errorMsg.textContent = '';
  statusMsg.textContent = '';

  try {
    const dataUrl = await captureAsPng();
    // Office JS expects base64 without the data URL prefix
    const base64 = dataUrl.replace(/^data:image\/\w+;base64,/, '');

    await PowerPoint.run(async (context) => {
      // getSelectedSlides().items[0] gives the active slide (API 1.5).
      // Falls back to slides.getItemAt(0) when nothing is selected.
      const selected = context.presentation.getSelectedSlides();
      selected.load('items');
      await context.sync();

      const slide = selected.items[0] ?? context.presentation.slides.getItemAt(0);
      (slide.shapes as ShapeCollectionWithImage).addImage(base64);
      await context.sync();
    });

    saveToHistory(latex);
    statusMsg.textContent = '✓ 挿入しました';
    setTimeout(() => { statusMsg.textContent = ''; }, 2500);
  } catch (e: unknown) {
    errorMsg.textContent = `挿入エラー: ${e instanceof Error ? e.message : String(e)}`;
  } finally {
    insertBtn.textContent = 'スライドに挿入';
    insertBtn.disabled = false;
  }
}

// ---- History ----
function getHistory(): string[] {
  try {
    return JSON.parse(localStorage.getItem(HISTORY_KEY) ?? '[]') as string[];
  } catch {
    return [];
  }
}

function saveToHistory(latex: string): void {
  const history = getHistory().filter(h => h !== latex);
  history.unshift(latex);
  history.splice(MAX_HISTORY);
  localStorage.setItem(HISTORY_KEY, JSON.stringify(history));
  renderHistory();
}

function renderHistory(): void {
  const history = getHistory();

  if (history.length === 0) {
    historySection.style.display = 'none';
    return;
  }

  historySection.style.display = '';
  historyList.innerHTML = '';

  history.forEach((latex) => {
    const btn = document.createElement('button');
    btn.className = 'history-item';
    btn.title = latex;
    btn.textContent = latex.length > 50 ? `${latex.slice(0, 50)}…` : latex;
    btn.addEventListener('click', () => {
      latexInput.value = latex;
      renderPreview();
      latexInput.focus();
    });
    historyList.appendChild(btn);
  });
}

function clearHistory(): void {
  localStorage.removeItem(HISTORY_KEY);
  renderHistory();
}

// ---- Event listeners ----
latexInput.addEventListener('input', renderPreview);

scaleSlider.addEventListener('input', () => {
  scaleDisplay.textContent = `${scaleSlider.value}×`;
  renderPreview();
});

colorPicker.addEventListener('input', () => {
  colorLabel.textContent = colorPicker.value;
  renderPreview();
});

displayModeBtn.addEventListener('click', () => {
  isDisplayMode = true;
  displayModeBtn.classList.add('active');
  inlineModeBtn.classList.remove('active');
  renderPreview();
});

inlineModeBtn.addEventListener('click', () => {
  isDisplayMode = false;
  inlineModeBtn.classList.add('active');
  displayModeBtn.classList.remove('active');
  renderPreview();
});

insertBtn.addEventListener('click', () => void insertEquation());

clearHistoryBtn.addEventListener('click', clearHistory);

// Snippet buttons: click → load into textarea
document.querySelectorAll<HTMLButtonElement>('.snippet-btn').forEach((btn) => {
  btn.addEventListener('click', () => {
    const snippet = btn.dataset['latex'];
    if (snippet) {
      latexInput.value = snippet;
      renderPreview();
      latexInput.focus();
    }
  });
});

// ---- Keyboard shortcut: Ctrl+Enter to insert ----
latexInput.addEventListener('keydown', (e: KeyboardEvent) => {
  if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) {
    e.preventDefault();
    if (!insertBtn.disabled) void insertEquation();
  }
});

// ---- Office initialization ----
Office.onReady((info) => {
  if (info.host !== Office.HostType.PowerPoint) {
    // Running in a browser for development preview
    insertBtn.textContent = '⚠ PowerPoint 外では挿入不可';
    insertBtn.title = '開発プレビュー: PowerPoint 内で実行してください';
  }

  // Initial scale label
  scaleDisplay.textContent = `${DEFAULT_SCALE}×`;

  renderHistory();
  renderPreview();
});
