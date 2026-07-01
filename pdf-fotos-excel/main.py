import base64
import io
import json
import os
import re
import tempfile
import unicodedata
from copy import copy
from typing import Any, Dict, List, Optional

from fastapi import FastAPI, UploadFile, File, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter, column_index_from_string
from PIL import Image as PILImage, ImageOps
from starlette.background import BackgroundTask


app = FastAPI(title="Exportador de Fotos para Excel")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ============================================================
# Utilitários
# ============================================================

def normalize_text(value: Any) -> str:
    """
    Normaliza texto para comparação:
    - remove acentos
    - converte para minúsculo
    - remove espaços excedentes
    """
    if value is None:
        return ""

    text = str(value).strip()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.lower()
    text = re.sub(r"\s+", " ", text)

    return text


def is_foto_header(value: Any) -> bool:
    """
    Identifica cabeçalhos do tipo:
    Foto
    Foto_1
    Foto 1
    Foto-1
    FOTO_2
    """
    text = normalize_text(value)
    if not text:
        return False

    return bool(re.match(r"^foto([\s_-]*\d+)?$", text))


def safe_int(value: Any, default: int) -> int:
    try:
        return int(value)
    except Exception:
        return default


def str_to_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value

    if value is None:
        return False

    return str(value).strip().lower() in ["true", "1", "yes", "sim", "s"]


def decode_data_url(data_url: str) -> bytes:
    """
    Aceita:
    - data:image/jpeg;base64,/9j/...
    - /9j/...
    """
    if not data_url:
        raise ValueError("Imagem vazia.")

    if "," in data_url:
        data_url = data_url.split(",", 1)[1]

    return base64.b64decode(data_url)


def prepare_image(
    image_bytes: bytes,
    max_width: int,
    max_height: int,
    quality: int = 88,
):
    """
    Redimensiona a imagem mantendo proporção e converte para JPEG.
    Retorna:
    - buffer BytesIO
    - largura final
    - altura final
    """
    input_buffer = io.BytesIO(image_bytes)

    with PILImage.open(input_buffer) as img:
        img = ImageOps.exif_transpose(img)

        if img.mode not in ["RGB"]:
            img = img.convert("RGB")

        img.thumbnail((max_width, max_height), PILImage.Resampling.LANCZOS)

        output_buffer = io.BytesIO()
        img.save(
            output_buffer,
            format="JPEG",
            quality=quality,
            optimize=True
        )
        output_buffer.seek(0)

        return output_buffer, img.width, img.height


def copy_cell_style(source_cell, target_cell):
    """
    Copia estilo básico de uma célula para outra.
    Útil para dar aparência parecida aos cabeçalhos Foto_x.
    """
    if source_cell is None or target_cell is None:
        return

    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.fill = copy(source_cell.fill)
        target_cell.border = copy(source_cell.border)
        target_cell.alignment = copy(source_cell.alignment)
        target_cell.number_format = source_cell.number_format
        target_cell.protection = copy(source_cell.protection)


def detect_photo_columns(ws, header_row: int) -> Dict[str, Any]:
    """
    Detecta colunas Foto_x.

    Regra:
    1. Se existirem colunas com cabeçalho Foto_1, Foto_2 etc., usa essas colunas.
    2. Se não existirem, usa as colunas vazias após a última coluna nomeada.
       No template enviado, isso equivale às colunas O até AF.
    3. Se o Excel não preservar as colunas vazias, cria virtualmente 18 colunas após a última coluna nomeada.
    """

    max_col = ws.max_column

    header_cells = []
    last_named_col = 0

    for col_idx in range(1, max_col + 1):
        cell = ws.cell(row=header_row, column=col_idx)
        raw_value = cell.value
        name = str(raw_value).strip() if raw_value is not None else ""

        if name:
            last_named_col = col_idx

        header_cells.append({
            "col": col_idx,
            "letter": get_column_letter(col_idx),
            "name": name,
            "is_blank": not bool(name)
        })

    explicit_photo_columns = []

    for item in header_cells:
        if is_foto_header(item["name"]):
            explicit_photo_columns.append({
                "name": item["name"],
                "col": item["col"],
                "letter": item["letter"],
                "source": "explicit"
            })

    if explicit_photo_columns:
        explicit_photo_columns.sort(key=lambda x: x["col"])

        return {
            "mode": "explicit",
            "last_named_col": last_named_col,
            "columns": explicit_photo_columns,
            "message": "Colunas Foto_x encontradas no cabeçalho do Excel."
        }

    

# Caso não encontre Foto_x, usa as colunas vazias depois da última coluna nomeada.
    

# No seu template, última coluna nomeada é N, depois vêm O:AF vazias.
    blank_photo_columns = []

    if max_col > last_named_col:
        count_blank = max_col - last_named_col
    else:
        

# Fallback caso o Excel não preserve colunas vazias.
        

# Cria Foto_1 até Foto_18 após a última coluna nomeada.
        count_blank = 18

    for i in range(count_blank):
        col_idx = last_named_col + 1 + i
        blank_photo_columns.append({
            "name": f"Foto_{i + 1}",
            "col": col_idx,
            "letter": get_column_letter(col_idx),
            "source": "blank_after_last_named"
        })

    return {
        "mode": "blank_after_last_named",
        "last_named_col": last_named_col,
        "columns": blank_photo_columns,
        "message": (
            "Nenhuma coluna Foto_x foi encontrada. "
            "As colunas vazias após a última coluna nomeada serão tratadas como Foto_1, Foto_2 etc."
        )
    }


# ============================================================
# Frontend embutido no backend
# ============================================================

INDEX_HTML = """
<!DOCTYPE html>
<html lang="pt-BR"

> <head>
  <meta charset="UTF-8" />
  <title>Exportador de Fotos para Excel</title>
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />

  <style>
    :root {
      --bg: #f4f6f8;
      --card: #ffffff;
      --primary: #0057ff;
      --primary-dark: #003bb5;
      --text: #1f2937;
      --muted: #6b7280;
      --border: #d1d5db;
      --success: #16a34a;
      --danger: #dc2626;
      --warn: #f59e0b;
    }

    * {
      box-sizing: border-box;
    }

    body {
      margin: 0;
      font-family: Arial, Helvetica, sans-serif;
      background: var(--bg);
      color: var(--text);
    }

    header {
      background: #111827;
      color: white;
      padding: 22px 28px;
    }

    header h1 {
      margin: 0;
      font-size: 22px;
    }

    header p {
      margin: 6px 0 0;
      color: #d1d5db;
      font-size: 14px;
    }

    main {
      max-width: 1280px;
      margin: 24px auto;
      padding: 0 18px 40px;
    }

    .grid {
      display: grid;
      grid-template-columns: 360px 1fr;
      gap: 18px;
    }

    @media (max-width: 900px) {
      .grid {
        grid-template-columns: 1fr;
      }
    }

    .card {
      background: var(--card);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 18px;
      box-shadow: 0 4px 12px rgba(0, 0, 0, 0.04);
    }

    .card h2 {
      margin: 0 0 14px;
      font-size: 18px;
    }

    .field {
      margin-bottom: 14px;
    }

    label {
      display: block;
      font-size: 13px;
      font-weight: 700;
      margin-bottom: 6px;
    }

    input[type="file"],
    input[type="number"],
    select {
      width: 100%;
      padding: 9px;
      border: 1px solid var(--border);
      border-radius: 8px;
      background: white;
      font-size: 14px;
    }

    input[type="checkbox"] {
      transform: scale(1.05);
      margin-right: 6px;
    }

    button {
      border: 0;
      background: var(--primary);
      color: white;
      padding: 11px 14px;
      border-radius: 8px;
      cursor: pointer;
      font-size: 14px;
      font-weight: 700;
    }

    button:hover {
      background: var(--primary-dark);
    }

    button:disabled {
      background: #9ca3af;
      cursor: not-allowed;
    }

    .btn-secondary {
      background: #374151;
    }

    .btn-secondary:hover {
      background: #1f2937;
    }

    .btn-danger {
      background: var(--danger);
    }

    .btn-danger:hover {
      background: #991b1b;
    }

    .row {
      display: flex;
      gap: 10px;
      align-items: center;
    }

    .row > * {
      flex: 1;
    }

    .status {
      padding: 10px 12px;
      border-radius: 8px;
      font-size: 13px;
      margin-top: 10px;
      display: none;
    }

    .status.visible {
      display: block;
    }

    .status.info {
      background: #e0ecff;
      color: #1e3a8a;
    }

    .status.success {
      background: #dcfce7;
      color: #166534;
    }

    .status.warn {
      background: #fef3c7;
      color: #92400e;
    }

    .status.error {
      background: #fee2e2;
      color: #991b1b;
    }

    .photo-columns {
      margin-top: 10px;
      padding: 10px;
      border: 1px dashed var(--border);
      border-radius: 8px;
      background: #f9fafb;
      font-size: 13px;
      max-height: 180px;
      overflow: auto;
    }

    .image-list {
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(230px, 1fr));
      gap: 14px;
    }

    .image-card {
      border: 1px solid var(--border);
      border-radius: 12px;
      background: white;
      padding: 10px;
    }

    .image-card img {
      width: 100%;
      height: 145px;
      object-fit: contain;
      background: #f3f4f6;
      border-radius: 8px;
      border: 1px solid #e5e7eb;
    }

    .image-card .name {
      font-size: 13px;
      font-weight: 700;
      margin: 8px 0;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }

    .image-card .meta {
      color: var(--muted);
      font-size: 12px;
      margin-bottom: 8px;
    }

    .actions {
      display: flex;
      gap: 10px;
      margin-top: 18px;
    }

    .actions button {
      flex: 1;
    }

    .small {
      font-size: 12px;
      color: var(--muted);
    }

    .badge {
      display: inline-block;
      padding: 3px 7px;
      background: #e5e7eb;
      color: #374151;
      border-radius: 999px;
      font-size: 12px;
      margin: 2px;
    }

    .overlay {
      position: fixed;
      inset: 0;
      display: none;
      align-items: center;
      justify-content: center;
      background: rgba(17, 24, 39, 0.55);
      z-index: 9999;
    }

    .overlay.visible {
      display: flex;
    }

    .modal {
      background: white;
      padding: 22px;
      border-radius: 12px;
      width: min(420px, calc(100% - 32px));
      box-shadow: 0 20px 60px rgba(0, 0, 0, 0.25);
      text-align: center;
    }

    .progress {
      height: 10px;
      background: #e5e7eb;
      border-radius: 999px;
      overflow: hidden;
      margin-top: 16px;
    }

    .progress > div {
      height: 100%;
      width: 0;
      background: var(--primary);
      transition: width 0.25s;
    }
  </style>
</head>

<body>
  <header>
    <h1>Exportador de Fotos para Excel</h1>
    <p>Frontend e backend juntos em uma única aplicação FastAPI.</p>
  </header>

  <main>
    <div class="grid"

>       <section class="card"

>         <h2>1. Template Excel</h2>

        <div class="field"

>           <label>Arquivo Excel .xlsx</label>
          <input id="excelFile" type="file" accept=".xlsx" />
        </div>

        <div class="field"

>           <label>Aba</label>
          <select id="sheetSelect" disabled></select>
        </div>

        <div class="field"

>           <label>Linha do cabeçalho</label>
          <input id="headerRow" type="number" min="1" value="1" />
          <div class="small">No template enviado, aparentemente é a linha 1.</div>
        </div>

        <button id="btnAnalyze" type="button" disabled>Analisar Template</button>

        <div id="templateStatus" class="status"></div>

        <div id="photoColumnsBox" class="photo-columns" style="display:none;"></div>

        <hr style="margin: 18px 0; border:0; border-top:1px solid #e5e7eb;" />

        <h2>2. Imagens</h2>

        <div class="field"

>           <label>Selecionar imagens</label>
          <input id="imageFiles" type="file" accept="image/*" multiple />
        </div>

        <div class="row"

>           <div class="field"

>             <label>Largura máxima</label>
            <input id="imgWidth" type="number" value="130" min="30" />
          </div>

          <div class="field"

>             <label>Altura máxima</label>
            <input id="imgHeight" type="number" value="95" min="30" />
          </div>
        </div>

        <div class="field"

>           <label>
            <input id="writeHeaders" type="checkbox" checked />
            Escrever Foto_1, Foto_2 etc. no cabeçalho quando a coluna estiver vazia
          </label>
        </div>

        <div class="actions"

>           <button id="btnAuto" type="button" class="btn-secondary" disabled>Distribuir Automaticamente</button>
          <button id="btnClear" type="button" class="btn-danger" disabled>Limpar Imagens</button>
        </div>

        <div class="actions"

>           <button id="btnExport" type="button" disabled>Exportar Excel</button>
        </div>
      </section>

      <section class="card"

>         <h2>3. Mapeamento das Imagens</h2>
        <p class="small"

>           Escolha a coluna Foto_x e a linha em que cada imagem será inserida.
        </p>

        <div id="imageList" class="image-list"></div>
      </section>
    </div>
  </main>

  <div id="overlay" class="overlay"

>     <div class="modal"

>       <h2 id="modalTitle">Processando...</h2>
      <p id="modalText" class="small">Aguarde.</p>
      <div class="progress"

>         <div id="progressBar"></div>
      </div>
    </div>
  </div>

  <script>
    const $ = (id) => document.getElementById(id);

    let excelFile = null;
    let sheetNames = [];
    let photoColumns = [];
    let images = [];

    function setStatus(type, message) {
      const box = $("templateStatus");
      box.className = "status visible " + type;
      box.textContent = message;
    }

    function showModal(title, text, percent) {
      $("overlay").classList.add("visible");
      $("modalTitle").textContent = title;
      $("modalText").textContent = text;
      $("progressBar").style.width = percent + "%";
    }

    function hideModal() {
      $("overlay").classList.remove("visible");
    }

    function startRow() {
      const headerRow = parseInt($("headerRow").value || "1", 10);
      return headerRow + 1;
    }

    function escapeHtml(text) {
      return String(text || "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
    }

    function getColumnOptions(selectedCol) {
      return photoColumns.map(col => {
        const selected = String(col.col) === String(selectedCol) ? "selected" : "";
        return `<option value="${col.col}" ${selected}>${escapeHtml(col.name)} - coluna ${col.letter}</option>`;
      }).join("");
    }

    function updateButtons() {
      $("btnExport").disabled = !(excelFile && images.length > 0 && photoColumns.length > 0);
      $("btnAuto").disabled = !(images.length > 0 && photoColumns.length > 0);
      $("btnClear").disabled = images.length === 0;
    }

    function renderPhotoColumns() {
      const box = $("photoColumnsBox");

      if (!photoColumns.length) {
        box.style.display = "none";
        box.innerHTML = "";
        return;
      }

      box.style.display = "block";
      box.innerHTML = `
        <strong>Colunas de fotos detectadas:</strong><br/>
        ${photoColumns.map(col => `
          <span class="badge">${escapeHtml(col.name)} = ${col.letter}</span>
        `).join("")}
      `;
    }

    function renderImages() {
      const list = $("imageList");

      if (!images.length) {
        list.innerHTML = `
          <div class="small"

>             Nenhuma imagem selecionada ainda.
          </div>
        `;
        updateButtons();
        return;
      }

      list.innerHTML = images.map((img, idx) => `
        <div class="image-card"

>           <img src="${img.dataUrl}" alt="${escapeHtml(img.name)}" />

          <div class="name">${escapeHtml(img.name)}</div>
          <div class="meta">${img.width} x ${img.height}px</div>

          <div class="field"

>             <label>Coluna de destino</label>
            <select data-idx="${idx}" class="colSelect"

>               ${getColumnOptions(img.col)}
            </select>
          </div>

          <div class="field"

>             <label>Linha de destino</label>
            <input data-idx="${idx}" class="rowInput" type="number" min="1" value="${img.row}" />
          </div>

          <button type="button" class="btn-danger removeBtn" data-idx="${idx}"

>             Remover
          </button>
        </div>
      `).join("");

      document.querySelectorAll(".colSelect").forEach(el => {
        el.addEventListener("change", () => {
          const idx = parseInt(el.dataset.idx, 10);
          images[idx].col = parseInt(el.value, 10);
        });
      });

      document.querySelectorAll(".rowInput").forEach(el => {
        el.addEventListener("change", () => {
          const idx = parseInt(el.dataset.idx, 10);
          images[idx].row = parseInt(el.value || startRow(), 10);
        });
      });

      document.querySelectorAll(".removeBtn").forEach(el => {
        el.addEventListener("click", () => {
          const idx = parseInt(el.dataset.idx, 10);
          images.splice(idx, 1);
          renderImages();
        });
      });

      updateButtons();
    }

    async function analyzeTemplate() {
      if (!excelFile) {
        setStatus("warn", "Selecione um arquivo Excel primeiro.");
        return;
      }

      showModal("Analisando template", "Lendo abas e colunas do Excel...", 30);

      try {
        const fd = new FormData();
        fd.append("template", excelFile);
        fd.append("header_row", $("headerRow").value || "1");

        if ($("sheetSelect").value) {
          fd.append("sheet_name", $("sheetSelect").value);
        }

        const res = await fetch("/api/template-info", {
          method: "POST",
          body: fd
        });

        const data = await res.json();

        if (!res.ok) {
          throw new Error(data.error || "Erro ao analisar template.");
        }

        sheetNames = data.sheets || [];

        const currentValue = $("sheetSelect").value;

        $("sheetSelect").innerHTML = sheetNames.map(name => {
          const selected = name === data.selected_sheet ? "selected" : "";
          return `<option value="${escapeHtml(name)}" ${selected}>${escapeHtml(name)}</option>`;
        }).join("");

        $("sheetSelect").disabled = sheetNames.length === 0;

        if (currentValue && sheetNames.includes(currentValue)) {
          $("sheetSelect").value = currentValue;
        } else {
          $("sheetSelect").value = data.selected_sheet;
        }

        photoColumns = data.photo_columns || [];

        renderPhotoColumns();

        if (data.mode === "explicit") {
          setStatus("success", data.message);
        } else {
          setStatus("warn", data.message);
        }

        autoDistribute(false);
        renderImages();

      } catch (err) {
        setStatus("error", err.message);
      } finally {
        hideModal();
      }
    }

    function autoDistribute(showMessage = true) {
      if (!photoColumns.length || !images.length) {
        return;
      }

      const firstRow = startRow();
      const colCount = photoColumns.length;

      images.forEach((img, idx) => {
        const colIndex = idx % colCount;
        const rowOffset = Math.floor(idx / colCount);

        img.col = photoColumns[colIndex].col;
        img.row = firstRow + rowOffset;
      });

      renderImages();

      if (showMessage) {
        setStatus("info", "Imagens distribuídas automaticamente nas colunas Foto_x.");
      }
    }

    async function resizeImageFile(file, maxSize = 1600, quality = 0.88) {
      return new Promise((resolve, reject) => {
        const reader = new FileReader();

        reader.onload = () => {
          const img = new Image();

          img.onload = () => {
            let width = img.width;
            let height = img.height;

            if (width > maxSize || height > maxSize) {
              const scale = Math.min(maxSize / width, maxSize / height);
              width = Math.round(width * scale);
              height = Math.round(height * scale);
            }

            const canvas = document.createElement("canvas");
            canvas.width = width;
            canvas.height = height;

            const ctx = canvas.getContext("2d");
            ctx.drawImage(img, 0, 0, width, height);

            const dataUrl = canvas.toDataURL("image/jpeg", quality);

            resolve({
              name: file.name,
              dataUrl,
              width,
              height
            });
          };

          img.onerror = reject;
          img.src = reader.result;
        };

        reader.onerror = reject;
        reader.readAsDataURL(file);
      });
    }

    async function handleImagesSelected(files) {
      if (!files || !files.length) {
        return;
      }

      showModal("Carregando imagens", "Preparando imagens para exportação...", 20);

      try {
        const fileArray = Array.from(files);

        for (let i = 0; i < fileArray.length; i++) {
          const prepared = await resizeImageFile(fileArray[i]);
          images.push({
            ...prepared,
            col: photoColumns[0] ? photoColumns[0].col : null,
            row: startRow()
          });

          const percent = 20 + Math.round(((i + 1) / fileArray.length) * 70);
          showModal("Carregando imagens", `Imagem ${i + 1} de ${fileArray.length}`, percent);
        }

        autoDistribute(false);
        renderImages();

      } catch (err) {
        setStatus("error", "Erro ao carregar imagem: " + err.message);
      } finally {
        hideModal();
      }
    }

    async function exportExcel() {
      if (!excelFile) {
        setStatus("warn", "Selecione o template Excel.");
        return;
      }

      if (!images.length) {
        setStatus("warn", "Selecione pelo menos uma imagem.");
        return;
      }

      if (!photoColumns.length) {
        setStatus("warn", "Nenhuma coluna de foto foi detectada.");
        return;
      }

      showModal("Exportando Excel", "Enviando dados para o backend...", 20);

      try {
        const assignments = images.map(img => ({
          name: img.name,
          image: img.dataUrl,
          col: img.col,
          row: img.row
        }));

        const fd = new FormData();
        fd.append("template", excelFile);
        fd.append("sheet_name", $("sheetSelect").value);
        fd.append("header_row", $("headerRow").value || "1");
        fd.append("photo_columns", JSON.stringify(photoColumns));
        fd.append("assignments", JSON.stringify(assignments));
        fd.append("write_headers", $("writeHeaders").checked ? "true" : "false");
        fd.append("img_width", $("imgWidth").value || "130");
        fd.append("img_height", $("imgHeight").value || "95");

        showModal("Exportando Excel", "Inserindo imagens no arquivo...", 55);

        const res = await fetch("/api/exportar", {
          method: "POST",
          body: fd
        });

        if (!res.ok) {
          let msg = "Erro ao exportar Excel.";

          try {
            const err = await res.json();
            msg = err.error || msg;
          } catch (_) {}

          throw new Error(msg);
        }

        showModal("Exportando Excel", "Baixando arquivo final...", 85);

        const blob = await res.blob();
        const url = URL.createObjectURL(blob);

        const a = document.createElement("a");
        a.href = url;
        a.download = "resultado_fotos.xlsx";
        document.body.appendChild(a);
        a.click();
        a.remove();

        URL.revokeObjectURL(url);

        showModal("Concluído", "Arquivo exportado com sucesso.", 100);

        setTimeout(() => {
          hideModal();
          setStatus("success", "Excel exportado com sucesso.");
        }, 800);

      } catch (err) {
        hideModal();
        setStatus("error", err.message);
      }
    }

    $("excelFile").addEventListener("change", async (event) => {
      excelFile = event.target.files[0] || null;
      $("btnAnalyze").disabled = !excelFile;

      if (excelFile) {
        await analyzeTemplate();
      }
    });

    $("btnAnalyze").addEventListener("click", analyzeTemplate);

    $("sheetSelect").addEventListener("change", analyzeTemplate);

    $("headerRow").addEventListener("change", analyzeTemplate);

    $("imageFiles").addEventListener("change", async (event) => {
      await handleImagesSelected(event.target.files);
      event.target.value = "";
    });

    $("btnAuto").addEventListener("click", () => autoDistribute(true));

    $("btnClear").addEventListener("click", () => {
      images = [];
      renderImages();
      setStatus("info", "Imagens removidas.");
    });

    $("btnExport").addEventListener("click", exportExcel);

    renderImages();
  </script>
</body>
</html>
"""


# ============================================================
# Rotas
# ============================================================

@app.get("/", response_class=HTMLResponse)
async def index():
    return HTMLResponse(INDEX_HTML)


@app.get("/api/health")
async def health():
    return {
        "status": "ok",
        "app": "Exportador de Fotos para Excel"
    }


@app.post("/api/template-info")
async def template_info(
    template: UploadFile = File(...),
    header_row: int = Form(1),
    sheet_name: Optional[str] = Form(None)
):
    """
    Lê o template Excel e retorna:
    - abas disponíveis
    - aba selecionada
    - colunas Foto_x detectadas
    """

    try:
        filename = template.filename or ""

        if not filename.lower().endswith(".xlsx"):
            return JSONResponse(
                {"error": "Envie um arquivo .xlsx."},
                status_code=400
            )

        content = await template.read()
        wb = load_workbook(io.BytesIO(content))

        sheets = wb.sheetnames

        if not sheets:
            return JSONResponse(
                {"error": "O arquivo Excel não possui abas."},
                status_code=400
            )

        if sheet_name and sheet_name in sheets:
            selected_sheet = sheet_name
        else:
            selected_sheet = wb.active.title

        ws = wb[selected_sheet]

        header_row = safe_int(header_row, 1)

        if header_row < 1:
            header_row = 1

        detection = detect_photo_columns(ws, header_row)

        return {
            "sheets": sheets,
            "selected_sheet": selected_sheet,
            "header_row": header_row,
            "mode": detection["mode"],
            "message": detection["message"],
            "last_named_col": detection["last_named_col"],
            "photo_columns": detection["columns"]
        }

    except Exception as e:
        return JSONResponse(
            {"error": f"Erro ao ler template: {str(e)}"},
            status_code=500
        )


@app.post("/api/exportar")
async def exportar(
    template: UploadFile = File(...),
    sheet_name: str = Form(...),
    header_row: int = Form(1),
    photo_columns: str = Form("[]"),
    assignments: str = Form("[]"),
    write_headers: str = Form("true"),
    img_width: int = Form(130),
    img_height: int = Form(95)
):
    """
    Exporta o Excel inserindo as fotos nas células escolhidas.

    assignments esperado:
    [
      {
        "name": "foto1.jpg",
        "image": "data:image/jpeg;base64,...",
        "col": 15,
        "row": 2
      }
    ]
    """

    output_path = None

    try:
        filename = template.filename or ""

        if not filename.lower().endswith(".xlsx"):
            return JSONResponse(
                {"error": "Envie um arquivo .xlsx."},
                status_code=400
            )

        header_row = safe_int(header_row, 1)
        img_width = safe_int(img_width, 130)
        img_height = safe_int(img_height, 95)

        if img_width < 20:
            img_width = 130

        if img_height < 20:
            img_height = 95

        try:
            photo_columns_data = json.loads(photo_columns)
        except Exception:
            return JSONResponse(
                {"error": "JSON inválido em photo_columns."},
                status_code=400
            )

        try:
            assignments_data = json.loads(assignments)
        except Exception:
            return JSONResponse(
                {"error": "JSON inválido em assignments."},
                status_code=400
            )

        if not isinstance(assignments_data, list) or len(assignments_data) == 0:
            return JSONResponse(
                {"error": "Nenhuma imagem foi enviada para exportação."},
                status_code=400
            )

        content = await template.read()
        wb = load_workbook(io.BytesIO(content))

        if sheet_name not in wb.sheetnames:
            return JSONResponse(
                {"error": f"A aba '{sheet_name}' não foi encontrada no Excel."},
                status_code=400
            )

        ws = wb[sheet_name]

        

# Se solicitado, escreve Foto_1, Foto_2 etc. nas colunas vazias
        if str_to_bool(write_headers) and isinstance(photo_columns_data, list):
            previous_header_cell = ws.cell(row=header_row, column=max(1, ws.max_column))

            for col_info in photo_columns_data:
                col_idx = safe_int(col_info.get("col"), 0)
                name = str(col_info.get("name") or "").strip()

                if col_idx <= 0 or not name:
                    continue

                target_cell = ws.cell(row=header_row, column=col_idx)
                current_value = target_cell.value

                if current_value is None or str(current_value).strip() == "" or is_foto_header(current_value):
                    target_cell.value = name

                    

# Copia estilo da célula anterior quando possível
                    if col_idx > 1:
                        source_cell = ws.cell(row=header_row, column=col_idx - 1)
                        copy_cell_style(source_cell, target_cell)

        image_buffers = []
        used_cells = set()
        warnings = []

        for idx, item in enumerate(assignments_data, start=1):
            image_data = item.get("image")
            target_row = safe_int(item.get("row"), header_row + 1)
            target_col = safe_int(item.get("col"), 0)

            if target_row <= header_row:
                target_row = header_row + 1

            if target_col <= 0:
                warnings.append(f"Imagem {idx} ignorada: coluna inválida.")
                continue

            cell_ref = f"{get_column_letter(target_col)}{target_row}"

            if cell_ref in used_cells:
                warnings.append(
                    f"A célula {cell_ref} recebeu mais de uma imagem. "
                    f"As imagens podem ficar sobrepostas."
                )

            used_cells.add(cell_ref)

            try:
                image_bytes = decode_data_url(image_data)
                img_buffer, final_w, final_h = prepare_image(
                    image_bytes=image_bytes,
                    max_width=img_width,
                    max_height=img_height,
                    quality=88
                )
            except Exception as e:
                warnings.append(f"Imagem {idx} ignorada por erro de leitura: {str(e)}")
                continue

            xl_image = XLImage(img_buffer)
            xl_image.width = final_w
            xl_image.height = final_h
            xl_image.anchor = cell_ref

            ws.add_image(xl_image)

            image_buffers.append(img_buffer)

            col_letter = get_column_letter(target_col)

            

# Ajuste de largura da coluna.
            

# openpyxl mede largura em unidade aproximada de caracteres.
            desired_col_width = max(12, img_width / 7.0)

            current_width = ws.column_dimensions[col_letter].width
            if current_width is None or current_width < desired_col_width:
                ws.column_dimensions[col_letter].width = desired_col_width

            

# Ajuste de altura da linha.
            

# Excel mede altura em pontos. Aproximação: 1 px = 0.75 pt.
            desired_row_height = max(35, img_height * 0.75)

            current_height = ws.row_dimensions[target_row].height
            if current_height is None or current_height < desired_row_height:
                ws.row_dimensions[target_row].height = desired_row_height

        if len(used_cells) == 0:
            return JSONResponse(
                {"error": "Nenhuma imagem válida foi inserida no Excel."},
                status_code=400
            )

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        output_path = tmp.name
        tmp.close()

        wb.save(output_path)

        def cleanup_file(path: str):
            try:
                if path and os.path.exists(path):
                    os.remove(path)
            except Exception:
                pass

        return FileResponse(
            path=output_path,
            filename="resultado_fotos.xlsx",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            background=BackgroundTask(cleanup_file, output_path)
        )

    except Exception as e:
        if output_path and os.path.exists(output_path):
            try:
                os.remove(output_path)
            except Exception:
                pass

        return JSONResponse(
            {"error": f"Erro ao exportar Excel: {str(e)}"},
            status_code=500
        )
