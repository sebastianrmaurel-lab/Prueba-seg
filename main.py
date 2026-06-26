from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
import io
from processes.proceso1 import procesar_pago1

app = FastAPI(title="Pagos App")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
def root():
    return {"status": "ok", "procesos": 7}

@app.post("/proceso1/procesar")
async def proceso1(
    pagot1: UploadFile = File(..., description="Archivo PAGOT1.txt"),
    referencia: UploadFile = File(..., description="Excel de referencia (.xlsx)")
):
    try:
        txt_bytes = await pagot1.read()
        xls_bytes = await referencia.read()

        resultado_txt = procesar_pago1(txt_bytes, xls_bytes)

        return StreamingResponse(
            io.BytesIO(resultado_txt.encode("latin-1")),
            media_type="text/plain",
            headers={"Content-Disposition": "attachment; filename=PAGOT1_Procesado.txt"}
        )
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))
