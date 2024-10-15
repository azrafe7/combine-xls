# main.py
import os
import tempfile
import pandas as pd
from fastapi import FastAPI, File, UploadFile, Form
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
import uvicorn
from combine_xls import combine_excel_files, get_column

app = FastAPI()

# Serve static files
app.mount("/static", StaticFiles(directory="static"), name="static")

@app.get("/", response_class=HTMLResponse)
async def read_root():
    with open("static/index.html", "r") as f:
        content = f.read()
    return content

@app.post("/get_columns")
async def get_columns(file: UploadFile = File(...)):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_file:
        temp_file.write(await file.read())
        temp_file.close()
        df = pd.read_excel(temp_file.name)
        columns = df.columns.tolist()
        os.unlink(temp_file.name)
    return JSONResponse(content={"columns": columns})

@app.post("/combine")
async def combine_files(
    file_a: UploadFile = File(...),
    file_b: UploadFile = File(...),
    column_a: str = Form(...),
    column_b: str = Form(...),
    case_sensitive: bool = Form(True),
    like_comparison: bool = Form(False),
    debug: bool = Form(True)
):
    # Create temporary files
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_a, \
         tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_b, \
         tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_output:
        
        # Save uploaded files
        temp_a.write(await file_a.read())
        temp_b.write(await file_b.read())
        
        # Close the files to ensure all data is written
        temp_a.close()
        temp_b.close()

        # Combine Excel files using the function from combine_xls.py
        combine_excel_files(
            temp_a.name, temp_b.name, column_a, column_b, 
            temp_output.name, case_sensitive, like_comparison, debug
        )

        # Return the merged file
        return FileResponse(
            temp_output.name,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="merged.xlsx"
        )

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)