from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse, JSONResponse
import io
from processes.proceso1 import procesar_pago1

app = FastAPI(title="Pagos App")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
def root():
    return {"status": "ok", "procesos": 7}

@app.options("/proceso1/procesar")
async def options_proceso1():
    return JSONResponse(content={}, headers={
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "POST, OPTIONS",
        "Access-Control-Allow-Headers": "*",
    })

@app.post("/proceso1/procesar")
async def proceso1(
    pagot1: UploadFile = File(...),
    referencia: UploadFile = File(...)
):
    try:
        txt_bytes = await pagot1.read()
        xls_bytes = await referencia.read()
        resultado_txt = procesar_pago1(txt_bytes, xls_bytes)
        return StreamingResponse(
            io.BytesIO(resultado_txt.encode("latin-1")),
            media_type="text/plain",
            headers={
                "Content-Disposition": "attachment; filename=PAGOT1_Procesado.txt",
                "Access-Control-Allow-Origin": "*",
            }
        )
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))
