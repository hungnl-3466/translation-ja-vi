from fastapi import FastAPI, File, UploadFile
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import aiofiles
from pptx import Presentation
import uvicorn
import json
import io
import sys

sys.path.insert(0, "app")
from domain.pipline import Pineline


app = FastAPI()
origins = ["*"]
app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/upload_pptx/")
async def upload_pptx(file: UploadFile = File(...)):
    # Lưu file pptx vào hệ thống
    # json_data = await json_file.read()
    content = await file.read()        
    
    pptx_stream = io.BytesIO(content)
    pipeline_pptx = Pineline()
    slides_data = pipeline_pptx.extract_text(pptx_stream)
    dict_output = pipeline_pptx.extract_text(slides_data)
    
    # dict_batch = pipeline_pptx.split_batch(dict_output, 6)
    
    # print(dict_batch)
    # print("---------------------------------------------")
    with open('dict_slide_text_test.json', 'w', encoding='utf-8') as f:
        json.dump(dict_output, f,ensure_ascii=False, indent=4)
    # with open('data.json', 'w', encoding='utf-8') as f:
    # json.dump(data, f, ensure_ascii=False, indent=4)
    # Trả về nội dung đã đọc từ file pptx
    return {"filename": file.filename, "content": str(dict_output)}

if __name__ == "__main__":
    import os
    # os.makedirs('files', exist_ok=True)
    uvicorn.run(app, host="0.0.0.0", port=8000)
    