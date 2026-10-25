import os
import shutil
from pathlib import Path
from fastapi import FastAPI, UploadFile, File, Form, Request
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
from dotenv import load_dotenv

# Carica le variabili di ambiente (inclusa GEMINI_API_KEY)
load_dotenv()

from backend.pipeline_adapter import process_audio_with_pipeline, process_multiple_audio_with_pipeline, process_audio_repeated_with_pipeline, process_multiple_audio_repeated_with_pipeline, fuse_reports_with_pipeline, match_incident_with_pipeline, chat_with_incident_pipeline, process_audio_quality_with_pipeline

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/api/health")
async def health():
    return {"ok": True}

# Monta le api PRIMA di montare i file statici sulla root
@app.post("/api/analyze-audio")
async def analyze_audio(file: UploadFile = File(...)):
    # Crea una cartella temporanea per l'audio
    temp_dir = Path("temp_audio")
    temp_dir.mkdir(exist_ok=True)
    temp_file_path = temp_dir / file.filename
    
    with open(temp_file_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
        
    try:
        # Usa il pipeline adapter (che salva anche in output_report)
        result = process_audio_with_pipeline(temp_file_path)
        return {"success": True, "data": result}
    except Exception as e:
        return {"success": False, "error": str(e)}
    finally:
        # Cleanup
        if temp_file_path.exists():
            os.remove(temp_file_path)

@app.post("/api/analyze-audio-batch")
async def analyze_audio_batch(files: list[UploadFile] = File(...)):
    temp_dir = Path("temp_audio")
    temp_dir.mkdir(exist_ok=True)
    saved_paths = []

    try:
        for file in files:
            temp_file_path = temp_dir / file.filename
            with open(temp_file_path, "wb") as buffer:
                shutil.copyfileobj(file.file, buffer)
            saved_paths.append(temp_file_path)

        result = process_multiple_audio_with_pipeline(saved_paths)
        return {"success": True, "data": result}
    except Exception as e:
        return {"success": False, "error": str(e)}
    finally:
        for path in saved_paths:
            if path.exists():
                os.remove(path)

@app.post("/api/analyze-audio-repeated")
async def analyze_audio_repeated(files: list[UploadFile] = File(...), runs_per_file: int = Form(3)):
    temp_dir = Path("temp_audio")
    temp_dir.mkdir(exist_ok=True)
    saved_paths = []

    try:
        for file in files:
            temp_file_path = temp_dir / file.filename
            with open(temp_file_path, "wb") as buffer:
                shutil.copyfileobj(file.file, buffer)
            saved_paths.append(temp_file_path)

        result = process_audio_repeated_with_pipeline(saved_paths, runs_per_file)
        return {"success": True, "data": result}
    except Exception as e:
        return {"success": False, "error": str(e)}
    finally:
        for path in saved_paths:
            if path.exists():
                os.remove(path)

@app.post("/api/analyze-audio-batch-repeated")
async def analyze_audio_batch_repeated(files: list[UploadFile] = File(...), runs_per_batch: int = Form(3)):
    temp_dir = Path("temp_audio")
    temp_dir.mkdir(exist_ok=True)
    saved_paths = []

    try:
        for file in files:
            temp_file_path = temp_dir / file.filename
            with open(temp_file_path, "wb") as buffer:
                shutil.copyfileobj(file.file, buffer)
            saved_paths.append(temp_file_path)

        result = process_multiple_audio_repeated_with_pipeline(saved_paths, runs_per_batch)
        return {"success": True, "data": result}
    except Exception as e:
        return {"success": False, "error": str(e)}
    finally:
        for path in saved_paths:
            if path.exists():
                os.remove(path)

@app.post("/api/analyze-quality")
async def analyze_quality(files: list[UploadFile] = File(...), runs_per_file: int = Form(3)):
    temp_dir = Path("temp_audio")
    temp_dir.mkdir(exist_ok=True)
    saved_paths = []

    try:
        for file in files:
            temp_file_path = temp_dir / file.filename
            with open(temp_file_path, "wb") as buffer:
                shutil.copyfileobj(file.file, buffer)
            saved_paths.append(temp_file_path)

        result = process_audio_quality_with_pipeline(saved_paths, runs_per_file)
        return {"success": True, "data": result}
    except Exception as e:
        return {"success": False, "error": str(e)}
    finally:
        for path in saved_paths:
            if path.exists():
                os.remove(path)


@app.post("/api/fuse-reports")
async def fuse_reports(request: Request):
    try:
        body = await request.json()
        singoli = body.get("singoli", [])
        result = fuse_reports_with_pipeline(singoli)
        return {"success": True, "data": result}
    except Exception as e:
        return {"success": False, "error": str(e)}

@app.post("/api/match-incident")
async def match_incident(request: Request):
    try:
        body = await request.json()
        new_report = body.get("new_report", {})
        trascrizione = body.get("trascrizione", "")
        existing_incidents = body.get("existing_incidents", [])
        
        match_id = match_incident_with_pipeline(new_report, existing_incidents, trascrizione)
        return {"success": True, "match_id": match_id}
    except Exception as e:
        return {"success": False, "error": str(e)}

@app.post("/api/chat")
async def chat_endpoint(request: Request):
    try:
        body = await request.json()
        messages = body.get("messages", [])
        incident_data = body.get("incident_data", {})
        
        if not messages:
            return {"success": False, "error": "No messages provided"}
            
        response_text = chat_with_incident_pipeline(messages, incident_data)
        return {"success": True, "response": response_text}
    except Exception as e:
        return {"success": False, "error": str(e)}

@app.post("/api/set-api-key")
async def set_api_key(request: Request):
    try:
        body = await request.json()
        key = body.get("key", "").strip()
        if key:
            os.environ["GEMINI_API_KEY"] = key
        else:
            from dotenv import load_dotenv
            load_dotenv(override=True)
        return {"success": True}
    except Exception as e:
        return {"success": False, "error": str(e)}

# Serve la UI
@app.get("/")
async def get_index():
    with open("index.html", "r", encoding="utf-8") as f:
        return HTMLResponse(f.read())

# Serve i file statici (se ci sono altri file nella cartella root)
app.mount("/", StaticFiles(directory=".", html=True), name="static")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("server:app", host="127.0.0.1", port=8000, reload=False)
