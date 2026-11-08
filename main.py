import os
import sqlite3
from datetime import datetime
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
import shutil
import html
import subprocess
from shutil import which
from pypdf import PdfReader
from docx import Document
import openpyxl
from pptx import Presentation
from PIL import Image
from pydantic import BaseModel

# Import ReportLab untuk pembuatan PDF
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

app = FastAPI(title="Document & Image Converter API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

UPLOAD_DIR = "uploads"
OUTPUT_DIR = "outputs"
DB_FILE = "converter.db"

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)


# ============================================================
# HELPER: Deteksi LibreOffice (untuk konversi kualitas tinggi)
# ============================================================
def _find_soffice():
    """Cari executable LibreOffice (soffice) di sistem."""
    candidates = [
        "soffice", "libreoffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        r"C:\Program Files\LibreOffice 7\program\soffice.exe",
        r"C:\Program Files\LibreOffice 24\program\soffice.exe",
        "/usr/bin/soffice", "/usr/bin/libreoffice",
        "/snap/bin/libreoffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    for c in candidates:
        try:
            if os.path.sep in c or (len(c) > 1 and c[1] == ":"):
                if os.path.exists(c):
                    print(f"[LibreOffice] Ditemukan: {c}")
                    return c
            else:
                found = which(c)
                if found:
                    print(f"[LibreOffice] Ditemukan via PATH: {found}")
                    return found
        except Exception:
            continue
    print("[LibreOffice] TIDAK DITEMUKAN — konversi akan pakai fallback ReportLab (kualitas terbatas)")
    return None


SOFFICE = _find_soffice()


def convert_with_libreoffice(input_path: str, output_path: str, target_ext: str) -> bool:
    """Konversi menggunakan LibreOffice headless. Return True jika berhasil."""
    if not SOFFICE:
        return False
    try:
        out_dir = os.path.dirname(os.path.abspath(output_path))
        cmd = [
            SOFFICE, "--headless", "--norestore", "--nolockcheck",
            "--convert-to", target_ext,
            "--outdir", out_dir,
            input_path,
        ]
        result = subprocess.run(cmd, capture_output=True, timeout=120)
        if result.returncode != 0:
            print(f"[LibreOffice] stderr: {result.stderr.decode(errors='ignore')}")
            return False
        produced = os.path.join(
            out_dir,
            os.path.splitext(os.path.basename(input_path))[0] + "." + target_ext
        )
        if os.path.exists(produced):
            if os.path.abspath(produced) != os.path.abspath(output_path):
                shutil.move(produced, output_path)
            return True
        return False
    except Exception as e:
        print(f"[LibreOffice] gagal: {e}")
        return False


# ============================================================
# HELPER: PDF -> DOCX (pdf2docx)
# ============================================================
def pdf_to_docx(input_path: str, output_path: str) -> bool:
    try:
        from pdf2docx import Converter
        cv = Converter(input_path)
        cv.convert(output_path)
        cv.close()
        return os.path.exists(output_path)
    except Exception as e:
        print(f"[pdf2docx] gagal: {e}")
        return False


# ============================================================
# HELPER: PDF -> PPTX (python-pptx, teks per halaman)
# ============================================================
def pdf_to_pptx(input_path: str, output_path: str) -> bool:
    try:
        from pptx import Presentation as PptxPresentation
        from pptx.util import Inches, Pt
        reader = PdfReader(input_path)
        prs = PptxPresentation()
        blank_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[0]
        # Hapus semua slide default
        while len(prs.slides) > 0:
            prs.slides._sldIdLst.remove(prs.slides._sldIdLst[0])

        for page_idx, page in enumerate(reader.pages, 1):
            text = page.extract_text() or ""
            slide = prs.slides.add_slide(blank_layout)
            # Judul halaman
            txBox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.8))
            tf = txBox.text_frame
            tf.text = f"Halaman {page_idx}"
            for p in tf.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(24)
                    r.font.bold = True
            # Isi
            if text.strip():
                body = slide.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(9), Inches(6))
                btf = body.text_frame
                btf.word_wrap = True
                btf.text = text
                for p in btf.paragraphs:
                    for r in p.runs:
                        r.font.size = Pt(14)
        prs.save(output_path)
        return os.path.exists(output_path)
    except Exception as e:
        print(f"[pdf2pptx] gagal: {e}")
        return False


# ============================================================
# HELPER: DOCX / PPTX -> PDF via ReportLab (fallback)
# ============================================================
def office_to_pdf_fallback(input_path: str, output_path: str, ext: str):
    """Fallback ReportLab. Kualitas terbatas: hanya teks + tabel + gambar statis."""
    from reportlab.platypus import Image as RLImage
    import io as _io

    pdf_doc = SimpleDocTemplate(
        output_path,
        pagesize=landscape(letter),
        rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40
    )
    styles = getSampleStyleSheet()
    story = []

    if ext == "docx":
        doc = Document(input_path)
        story.append(Paragraph("<b>Dokumen Word (.docx)</b>", styles['Heading1']))
        story.append(Spacer(1, 10))
        for p in doc.paragraphs:
            if p.text.strip():
                story.append(Paragraph(html.escape(p.text), styles['Normal']))
                story.append(Spacer(1, 4))
        for table in doc.tables:
            data = []
            for row in table.rows:
                data.append([html.escape(c.text) for c in row.cells])
            if data:
                t = Table(data)
                t.setStyle(TableStyle([
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                    ('FONTSIZE', (0, 0), (-1, -1), 8)
                ]))
                story.append(Spacer(1, 6))
                story.append(t)

    elif ext in ["ppt", "pptx"]:
        prs = Presentation(input_path)
        total = len(prs.slides)
        story.append(Paragraph("<b>Hasil Konversi PowerPoint</b>", styles['Title']))
        story.append(Paragraph(
            "<i>Catatan: Install LibreOffice untuk hasil konversi yang mempertahankan "
            "layout, gambar, dan desain asli.</i>", styles['Italic']))
        story.append(Spacer(1, 14))

        for idx, slide in enumerate(prs.slides, 1):
            # Kumpulkan teks & gambar dari slide
            texts = []
            images = []
            for shape in slide.shapes:
                # Teks
                if shape.has_text_frame:
                    for p in shape.text_frame.paragraphs:
                        if p.text.strip():
                            texts.append(p.text.strip())
                # Gambar (PIL picture)
                try:
                    if shape.shape_type == 13:  # MSO_SHAPE_TYPE.PICTURE
                        img_bytes = shape.image.blob
                        images.append(img_bytes)
                except Exception:
                    pass
                # Tabel
                if getattr(shape, "has_table", False) and shape.has_table:
                    tbl_data = []
                    for row in shape.table.rows:
                        tbl_data.append([cell.text for cell in row.cells])
                    if tbl_data:
                        t = Table(tbl_data)
                        t.setStyle(TableStyle([
                            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                            ('FONTSIZE', (0, 0), (-1, -1), 8),
                        ]))
                        story.append(t)
                        story.append(Spacer(1, 6))

            # Header slide
            story.append(Paragraph(f"<b>SLIDE {idx} dari {total}</b>", styles['Heading2']))
            story.append(Spacer(1, 6))

            # Gambar (maks 2 per slide supaya tidak kepanjangan)
            for ib in images[:2]:
                try:
                    img = RLImage(_io.BytesIO(ib))
                    # skala agar muat
                    max_w, max_h = 600, 350
                    ratio = min(max_w / img.drawWidth, max_h / img.drawHeight, 1.0)
                    img.drawWidth *= ratio
                    img.drawHeight *= ratio
                    story.append(img)
                    story.append(Spacer(1, 6))
                except Exception:
                    pass

            # Teks
            if texts:
                # Gabung tapi hindari duplikat beruntun
                seen = None
                for t in texts:
                    if t != seen:
                        story.append(Paragraph(html.escape(t), styles['Normal']))
                        story.append(Spacer(1, 3))
                        seen = t
            elif not images:
                story.append(Paragraph("<i>(Slide kosong / hanya berisi shape grafis)</i>",
                                       styles['Italic']))

            story.append(PageBreak())

    if story:
        pdf_doc.build(story)
    else:
        raise Exception("Dokumen kosong, tidak ada yang bisa dikonversi.")


# ============================================================
# HELPER: DOCX -> PPTX (tiap chunk paragraf jadi slide)
# ============================================================
def docx_to_pptx(input_path: str, output_path: str):
    from pptx import Presentation as PptxPresentation
    from pptx.util import Inches, Pt
    doc = Document(input_path)
    prs = PptxPresentation()
    blank_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[0]
    while len(prs.slides) > 0:
        prs.slides._sldIdLst.remove(prs.slides._sldIdLst[0])

    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    if not paragraphs:
        paragraphs = ["(Dokumen kosong)"]

    chunk_size = 8
    for i in range(0, len(paragraphs), chunk_size):
        chunk = paragraphs[i:i + chunk_size]
        slide = prs.slides.add_slide(blank_layout)
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.9))
        tf = title_box.text_frame
        tf.text = chunk[0][:80]
        for p in tf.paragraphs:
            for r in p.runs:
                r.font.size = Pt(28)
                r.font.bold = True
        if len(chunk) > 1:
            body = slide.shapes.add_textbox(Inches(0.5), Inches(1.3), Inches(9), Inches(5.5))
            btf = body.text_frame
            btf.word_wrap = True
            btf.text = "\n".join(chunk[1:])
            for p in btf.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(16)
    prs.save(output_path)
    return os.path.exists(output_path)


# ============================================================
# HELPER: PPTX -> DOCX
# ============================================================
def pptx_to_docx(input_path: str, output_path: str):
    prs = Presentation(input_path)
    doc = Document()
    doc.add_heading("Hasil Konversi dari PowerPoint", level=0)
    for idx, slide in enumerate(prs.slides, 1):
        doc.add_heading(f"Slide {idx}", level=1)
        for shape in slide.shapes:
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    if p.text.strip():
                        doc.add_paragraph(p.text)
        doc.add_paragraph("")
    doc.save(output_path)
    return os.path.exists(output_path)


# ============================================================
# DATABASE
# ============================================================
def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS conversions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            original_filename TEXT NOT NULL,
            source_format TEXT NOT NULL,
            target_format TEXT NOT NULL,
            file_size INTEGER NOT NULL,
            status TEXT DEFAULT 'completed',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL
        )
    """)
    conn.commit()
    conn.close()

init_db()


class UserAuth(BaseModel):
    username: str
    password: str


@app.post("/api/register")
async def register_user(user: UserAuth):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO users (username, password) VALUES (?, ?)", (user.username, user.password))
        conn.commit()
        return {"status": "success", "message": "Registrasi berhasil! Silakan login."}
    except sqlite3.IntegrityError:
        raise HTTPException(status_code=400, detail="Username sudah digunakan.")
    finally:
        conn.close()


@app.post("/api/login")
async def login_user(user: UserAuth):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM users WHERE username = ? AND password = ?", (user.username, user.password))
    db_user = cursor.fetchone()
    conn.close()

    if db_user:
        return {"status": "success", "message": "Login berhasil", "username": user.username}
    raise HTTPException(status_code=401, detail="Username atau password salah.")


# ============================================================
# ENDPOINT KONVERSI (diperbaiki)
# ============================================================
@app.post("/api/convert")
async def convert_document(
    file: UploadFile = File(...),
    target_format: str = Form(...)
):
    filename = file.filename
    ext = filename.split(".")[-1].lower()
    target_format = target_format.lower()

    input_path = os.path.join(UPLOAD_DIR, filename)
    with open(input_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    file_size = os.path.getsize(input_path)
    base_name = os.path.splitext(filename)[0]
    output_filename = f"converted_{base_name}.{target_format}"
    output_path = os.path.join(OUTPUT_DIR, output_filename)

    supported_image_exts = ["png", "jpg", "jpeg", "webp"]
    office_exts = ["doc", "docx", "xls", "xlsx", "ppt", "pptx"]

    try:
        # 1. Format sama -> copy
        if ext == target_format:
            shutil.copy(input_path, output_path)

        # 2. Gambar <-> Gambar
        elif ext in supported_image_exts and target_format in supported_image_exts:
            with Image.open(input_path) as img:
                if target_format in ["jpg", "jpeg"] and img.mode in ("RGBA", "P"):
                    img = img.convert("RGB")
                img.save(output_path)

        # 3. Gambar -> PDF
        elif ext in supported_image_exts and target_format == "pdf":
            with Image.open(input_path) as img:
                img.convert("RGB").save(output_path, "PDF", resolution=100.0)

        # 4. Semua -> PDF
        elif target_format == "pdf":
            if ext in office_exts and SOFFICE:
                # Prioritas LibreOffice (mempertahankan layout, gambar, tabel)
                if not convert_with_libreoffice(input_path, output_path, "pdf"):
                    office_to_pdf_fallback(input_path, output_path, ext)
            elif ext in ["doc", "docx", "ppt", "pptx"]:
                office_to_pdf_fallback(input_path, output_path, ext)
            elif ext == "xlsx":
                wb = openpyxl.load_workbook(input_path, data_only=True)
                sheet = wb.active
                pdf_doc = SimpleDocTemplate(output_path, pagesize=landscape(letter),
                                            rightMargin=40, leftMargin=40,
                                            topMargin=40, bottomMargin=40)
                styles = getSampleStyleSheet()
                story = [Paragraph("<b>Dokumen Excel (.xlsx)</b>", styles['Heading1']), Spacer(1, 10)]
                data = []
                for row in sheet.iter_rows(values_only=True):
                    row_data = [html.escape(str(c)) if c is not None else "" for c in row]
                    if any(row_data):
                        data.append(row_data[:8])
                if data:
                    t = Table(data[:60])
                    t.setStyle(TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                        ('FONTSIZE', (0, 0), (-1, -1), 8)
                    ]))
                    story.append(t)
                pdf_doc.build(story)
            elif ext == "pdf":
                shutil.copy(input_path, output_path)
            else:
                raise Exception(f"Format {ext} tidak didukung untuk konversi ke PDF.")

        # 5. PDF -> DOCX
        elif ext == "pdf" and target_format == "docx":
            if not pdf_to_docx(input_path, output_path):
                # Fallback: ekstrak teks saja
                reader = PdfReader(input_path)
                doc = Document()
                for i, page in enumerate(reader.pages, 1):
                    doc.add_heading(f"Halaman {i}", level=2)
                    txt = page.extract_text()
                    if txt:
                        for line in txt.split("\n"):
                            doc.add_paragraph(line)
                doc.save(output_path)

        # 6. PDF -> PPTX
        elif ext == "pdf" and target_format == "pptx":
            if not pdf_to_pptx(input_path, output_path):
                raise Exception("Gagal mengonversi PDF ke PPTX (pastikan python-pptx terinstall).")

        # 7. DOCX -> PPTX
        elif ext == "docx" and target_format == "pptx":
            if not (SOFFICE and convert_with_libreoffice(input_path, output_path, "pptx")):
                docx_to_pptx(input_path, output_path)

        # 8. PPT/PPTX -> DOCX
        elif ext in ["ppt", "pptx"] and target_format == "docx":
            if not (SOFFICE and convert_with_libreoffice(input_path, output_path, "docx")):
                pptx_to_docx(input_path, output_path)

        # 9. -> TXT
        elif target_format == "txt":
            text_content = ""
            if ext == "docx":
                doc = Document(input_path)
                text_content = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
            elif ext == "xlsx":
                wb = openpyxl.load_workbook(input_path, data_only=True)
                sheet = wb.active
                for row in sheet.iter_rows(values_only=True):
                    text_content += "\t".join([str(c) if c is not None else "" for c in row]) + "\n"
            elif ext in ["ppt", "pptx"]:
                prs = Presentation(input_path)
                for idx, slide in enumerate(prs.slides, 1):
                    text_content += f"\n=== SLIDE {idx} ===\n"
                    for shape in slide.shapes:
                        if shape.has_text_frame:
                            for p in shape.text_frame.paragraphs:
                                if p.text.strip():
                                    text_content += p.text + "\n"
            elif ext == "pdf":
                reader = PdfReader(input_path)
                for page in reader.pages:
                    extracted = page.extract_text()
                    if extracted:
                        text_content += extracted + "\n"
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(text_content)

        # 10. Fallback umum via LibreOffice
        else:
            if SOFFICE and convert_with_libreoffice(input_path, output_path, target_format):
                pass
            else:
                raise Exception(
                    f"Konversi {ext} -> {target_format} tidak didukung. "
                    f"Install LibreOffice untuk konversi format office yang lebih luas."
                )

        # Catat ke DB
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO conversions (original_filename, source_format, target_format, file_size, status, created_at)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (filename, ext, target_format, file_size, "completed", datetime.now()))
        conn.commit()
        conn.close()

        return {
            "status": "success",
            "message": f"File berhasil dikonversi ke .{target_format}",
            "download_url": f"/api/download/{output_filename}"
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/api/history")
async def get_conversion_history():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM conversions ORDER BY created_at DESC LIMIT 15")
    rows = cursor.fetchall()
    conn.close()
    return {"history": [dict(row) for row in rows]}


@app.delete("/api/history/{history_id}")
async def delete_history_item(history_id: int):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM conversions WHERE id = ?", (history_id,))
    conn.commit()
    conn.close()
    return {"status": "success", "message": "Riwayat berhasil dihapus"}


@app.delete("/api/history")
async def clear_all_history():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM conversions")
    conn.commit()
    conn.close()
    return {"status": "success", "message": "Semua riwayat berhasil dibersihkan"}


@app.get("/api/download/{filename}")
async def download_file(filename: str):
    file_path = os.path.join(OUTPUT_DIR, filename)
    if os.path.exists(file_path):
        return FileResponse(file_path, media_type="application/octet-stream", filename=filename)
    raise HTTPException(status_code=404, detail="File tidak ditemukan.")


@app.get("/", response_class=HTMLResponse)
async def serve_home():
    with open("index.html", "r", encoding="utf-8") as f:
        return f.read()


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True)
