from fastapi import FastAPI, File, UploadFile, Request, HTTPException
from fastapi.responses import JSONResponse, FileResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
import os
import uuid
import shutil

app = FastAPI()

# ✅ Proper CORS setup
origins = [
    "http://localhost:5173",
    "https://fastapi-cs6g.onrender.com",
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ✅ Explicit preflight handler (important for Render)
@app.options("/{rest_of_path:path}")
async def preflight_handler(rest_of_path: str):
    return Response(status_code=200)

UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

app.mount("/uploads", StaticFiles(directory=UPLOAD_DIR), name="uploads")


# ✅ Export Excel
@app.post("/export")
def export_excel(data: dict):
    import openpyxl

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for group, rows in data.items():
        ws = wb.create_sheet(title=group)
        ws.append(["Sort Order", "Entity", "Amount"])

        for row in rows:
            ws.append([
                row.get("Sort Order"),
                row.get("IDN_Legal_Entity_Formal_Name"),
                row.get("Amount")
            ])

    file_path = os.path.join(UPLOAD_DIR, "report.xlsx")
    wb.save(file_path)

    return FileResponse(
        file_path,
        filename="report.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ✅ Upload API
@app.post("/upload")
async def upload_file(request: Request, file: UploadFile = File(...)):

    allowed_ext = [".xlsx", ".xls"]
    ext = os.path.splitext(file.filename)[1].lower()

    if ext not in allowed_ext:
        raise HTTPException(status_code=400, detail="Only Excel files allowed")

    unique_name = f"{uuid.uuid4()}{ext}"
    file_path = os.path.join(UPLOAD_DIR, unique_name)

    with open(file_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    public_url = f"https://fastapi-cs6g.onrender.com/uploads/{unique_name}"

    return {
        "fileName": file.filename,
        "publicUrl": public_url
    }
