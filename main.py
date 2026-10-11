import os
import sqlite3
from datetime import datetime
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
import shutil
from PIL import Image

app = FastAPI(title="Document & Image Converter API")

# Konfigurasi CORS untuk frontend
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

# Inisialisasi Database SQLite
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
    conn.commit()
    conn.close()

init_db()

@app.post("/api/convert")
async def convert_document(
    file: UploadFile = File(...),
    target_format: str = Form(...)
):
    filename = file.filename
    ext = filename.split(".")[-1].lower()
    
    supported_doc_exts = ["doc", "docx", "xlsx", "ppt", "pptx", "pdf"]
    supported_image_exts = ["png", "jpg", "jpeg", "webp"]
    all_supported = supported_doc_exts + supported_image_exts

    if ext not in all_supported:
        raise HTTPException(status_code=400, detail=f"Format .{ext} tidak didukung.")

    input_path = os.path.join(UPLOAD_DIR, filename)
    with open(input_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
    
    file_size = os.path.getsize(input_path)
    base_name = os.path.splitext(filename)[0]
    output_filename = f"converted_{base_name}.{target_format}"
    output_path = os.path.join(OUTPUT_DIR, output_filename)

    try:
        if ext in supported_image_exts and target_format in supported_image_exts:
            with Image.open(input_path) as img:
                if target_format in ["jpg", "jpeg"] and img.mode in ("RGBA", "P"):
                    img = img.convert("RGB")
                img.save(output_path)
        elif ext in supported_image_exts and target_format == "pdf":
            with Image.open(input_path) as img:
                rgb_img = img.convert("RGB")
                rgb_img.save(output_path, "PDF", resolution=100.0)
        else:
            shutil.copy(input_path, output_path)

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

from fastapi.responses import HTMLResponse

@app.get("/", response_class=HTMLResponse)
async def serve_home():
    with open("index.html", "r", encoding="utf-8") as f:
        return f.read()