from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, StreamingResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from uuid import uuid4
import os, io, json, base64, zipfile, tempfile

# ================== APP SETUP ==================
app = FastAPI(title="ToolForge Backend API 🚀")

UPLOAD_DIR = "temp"
os.makedirs(UPLOAD_DIR, exist_ok=True)

# ================== CORS ==================
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://www.tooconvert.in",
        "https://tooconvert.in",
        "https://api.tooconvert.in"
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ================== UTILS ==================
async def save_upload(file: UploadFile) -> str:
    ext = os.path.splitext(file.filename)[1]
    path = os.path.join(UPLOAD_DIR, f"{uuid4().hex}{ext}")
    with open(path, "wb") as f:
        f.write(await file.read())
    return path

# ================== ROOT ==================
@app.get("/")
def root():
    return {"message": "ToolForge Backend 🚀", "docs": "/docs"}

@app.get("/health")
def health():
    return {"status": "Server Alive ✅"}

# =================================================
# ================== PDF TOOLS ====================
# =================================================

@app.post("/pdf-to-doc")
async def pdf_to_docx(file: UploadFile = File(...)):
    from pdf2docx import Converter

    input_path = await save_upload(file)
    output_path = input_path.replace(".pdf", ".docx")

    cv = Converter(input_path)
    cv.convert(output_path)
    cv.close()

    return FileResponse(output_path, filename="converted.docx")

@app.post("/merge-pdf")
async def merge_pdf(files: list[UploadFile] = File(...)):
    from PyPDF2 import PdfReader, PdfWriter

    writer = PdfWriter()
    for file in files:
        reader = PdfReader(io.BytesIO(await file.read()))
        for page in reader.pages:
            writer.add_page(page)

    output_path = os.path.join(UPLOAD_DIR, "merged.pdf")
    with open(output_path, "wb") as f:
        writer.write(f)

    return FileResponse(output_path, filename="merged.pdf")

@app.post("/split-pdf")
async def split_pdf(file: UploadFile = File(...), pages_per_split: int = Form(...)):
    from PyPDF2 import PdfReader, PdfWriter

    reader = PdfReader(io.BytesIO(await file.read()))
    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, "w") as zipf:
        for i in range(0, len(reader.pages), pages_per_split):
            writer = PdfWriter()
            for j in range(i, min(i + pages_per_split, len(reader.pages))):
                writer.add_page(reader.pages[j])
            buf = io.BytesIO()
            writer.write(buf)
            zipf.writestr(f"split_{i+1}.pdf", buf.getvalue())

    zip_buffer.seek(0)
    return StreamingResponse(zip_buffer, media_type="application/zip",
        headers={"Content-Disposition": "attachment; filename=splits.zip"}
    )

@app.post("/extract-text")
async def extract_text(file: UploadFile = File(...)):
    from PyPDF2 import PdfReader

    reader = PdfReader(io.BytesIO(await file.read()))
    text = "".join(page.extract_text() or "" for page in reader.pages)
    return {"text": text}

@app.post("/compress-pdf")
async def compress_pdf(file: UploadFile = File(...)):
    import fitz  # PyMuPDF

    pdf = fitz.open(stream=await file.read(), filetype="pdf")
    buf = io.BytesIO()
    pdf.save(buf, garbage=4, deflate=True)
    buf.seek(0)

    return StreamingResponse(buf, media_type="application/pdf",
        headers={"Content-Disposition": "attachment; filename=compressed.pdf"}
    )

# =================================================
# ================= IMAGE TOOLS ===================
# =================================================

@app.post("/resize-image")
async def resize_image(
    file: UploadFile = File(...),
    width: int = Form(...),
    height: int = Form(...)
):
    from PIL import Image

    img = Image.open(io.BytesIO(await file.read()))
    img = img.resize((width, height))

    buf = io.BytesIO()
    img.save(buf, "JPEG")
    buf.seek(0)

    return StreamingResponse(buf, media_type="image/jpeg")

@app.post("/compress-image")
async def compress_image(file: UploadFile = File(...), target_kb: int = Form(...)):
    from PIL import Image

    img = Image.open(io.BytesIO(await file.read())).convert("RGB")
    quality = 95

    while True:
        buf = io.BytesIO()
        img.save(buf, "JPEG", quality=quality)
        if buf.tell() / 1024 <= target_kb or quality <= 10:
            break
        quality -= 5

    buf.seek(0)
    return StreamingResponse(buf, media_type="image/jpeg")

@app.post("/watermark")
async def watermark_image(
    file: UploadFile = File(...),
    text: str = Form(...)
):
    from PIL import Image, ImageDraw, ImageFont

    img = Image.open(io.BytesIO(await file.read())).convert("RGBA")
    layer = Image.new("RGBA", img.size)
    draw = ImageDraw.Draw(layer)

    font = ImageFont.load_default()
    draw.text((20, 20), text, fill=(255, 255, 255, 120), font=font)

    out = Image.alpha_composite(img, layer)
    buf = io.BytesIO()
    out.convert("RGB").save(buf, "JPEG")
    buf.seek(0)

    return StreamingResponse(buf, media_type="image/jpeg")

# =================================================
# ============== FORMAT CONVERSIONS ===============
# =================================================

@app.post("/image-to-pdf")
async def image_to_pdf(file: UploadFile = File(...)):
    from PIL import Image

    img = Image.open(io.BytesIO(await file.read())).convert("RGB")
    buf = io.BytesIO()
    img.save(buf, "PDF")
    buf.seek(0)

    return StreamingResponse(buf, media_type="application/pdf")

@app.post("/pdf-to-image")
async def pdf_to_image(file: UploadFile = File(...)):
    from pdf2image import convert_from_bytes

    images = convert_from_bytes(await file.read())
    zip_buf = io.BytesIO()

    with zipfile.ZipFile(zip_buf, "w") as zipf:
        for i, img in enumerate(images):
            buf = io.BytesIO()
            img.save(buf, "JPEG")
            zipf.writestr(f"page_{i+1}.jpg", buf.getvalue())

    zip_buf.seek(0)
    return StreamingResponse(zip_buf, media_type="application/zip")

@app.post("/ppt-to-pdf")
async def ppt_to_pdf(file: UploadFile = File(...)):
    import aspose.slides as slides

    path = await save_upload(file)
    output = path.replace(".pptx", ".pdf")

    pres = slides.Presentation(path)
    pres.save(output, slides.export.SaveFormat.PDF)

    return FileResponse(output)

# =================================================
# =================== UTILITIES ===================
# =================================================

@app.post("/generate-qr")
async def generate_qr(text: str = Form(...)):
    import qrcode

    img = qrcode.make(text)
    buf = io.BytesIO()
    img.save(buf, "PNG")
    buf.seek(0)

    return StreamingResponse(buf, media_type="image/png")

@app.post("/base64-encode")
async def base64_encode(file: UploadFile = File(None), text: str = Form(None)):
    if not file and not text:
        raise HTTPException(400, "No input provided")

    data = await file.read() if file else text.encode()
    return {"base64": base64.b64encode(data).decode()}

@app.post("/base64-decode")
async def base64_decode(encoded: str = Form(...)):
    data = base64.b64decode(encoded)
    return StreamingResponse(io.BytesIO(data), media_type="application/octet-stream")

class JsonInput(BaseModel):
    json_text: str

@app.post("/format-json")
async def format_json(data: JsonInput):
    try:
        parsed = json.loads(data.json_text)
        return {"formatted": json.dumps(parsed, indent=4)}
    except json.JSONDecodeError:
        raise HTTPException(400, "Invalid JSON")
