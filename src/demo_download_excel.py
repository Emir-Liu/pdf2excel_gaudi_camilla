from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse
import openpyxl
from io import BytesIO

app = FastAPI()


@app.post("/create_excel/")
async def create_excel():
    # 创建一个 Excel 文件
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    # 向 Excel 文件添加一些数据
    sheet["A1"] = "Hello"
    sheet["B1"] = "World"

    # 将 Excel 文件保存到 BytesIO 对象中
    output = BytesIO()
    workbook.save(output)
    output.seek(0)  # 重置文件指针到开始位置

    # 创建一个 StreamingResponse 对象
    headers = {"Content-Disposition": 'attachment; filename="example.xlsx"'}
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


# 用于启动服务器
if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="127.0.0.1", port=8001)
