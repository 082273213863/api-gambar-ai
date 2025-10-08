from fastapi import FastAPI, UploadFile, File
from fastapi.responses import JSONResponse, FileResponse
from PIL import Image
import pytesseract
import io
from openpyxl import Workbook

app = FastAPI()

@app.post("/ocr")
async def ocr_image(file: UploadFile = File(...)):
    image_data = await file.read()
    image = Image.open(io.BytesIO(image_data))
    text = pytesseract.image_to_string(image)
    return JSONResponse(content={"text": text})

@app.post("/ocr-to-excel")
async def ocr_to_excel(file: UploadFile = File(...)):
    image_data = await file.read()
    image = Image.open(io.BytesIO(image_data))
    text = pytesseract.image_to_string(image)

    wb = Workbook()
    ws = wb.active
    ws.title = "OCR Result"

    for i, line in enumerate(text.splitlines(), start=1):
        ws.cell(row=i, column=1, value=line)

    excel_filename = "ocr_result.xlsx"
    wb.save(excel_filename)

    return FileResponse(
        excel_filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=excel_filename
    )