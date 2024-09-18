import os
import sys
import io
from io import BytesIO

ROOT_PATH = os.path.dirname(os.path.abspath(__file__))
# print(f"ROOT_PATH:{ROOT_PATH}")
sys.path.append(ROOT_PATH)

import fastapi
from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
import uvicorn
import openpyxl

from configs import TITLE, VERSION
from function.pdf2excel import func_pdf2excel

app = FastAPI(title=TITLE, version=VERSION)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.post("/uploadpdf/")
async def upload_pdf(file: UploadFile = File(...)):
    contents = await file.read()
    file_name = file.filename
    new_file_name = file_name.replace(" ", "_")
    bytes_io = io.BytesIO(contents)
    excel_content = func_pdf2excel(pdf_content=bytes_io)

    # 将 Excel 文件保存到 BytesIO 对象中
    output = BytesIO()
    excel_content.save(output)
    output.seek(0)  # 重置文件指针到开始位置

    # 创建一个 StreamingResponse 对象
    headers = {
        "Content-Disposition": f'attachment; filename="sheet_{new_file_name}.xlsx"'
    }
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


if __name__ == "__main__":
    uvicorn.run(
        app,
        # host=server_ip,
        # port=server_port
    )
