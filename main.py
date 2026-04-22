from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from routers import inventory, warehouse, outgoing, incoming, shipment, purchase, ledger, evaluation, incentive

app = FastAPI(title="재고수불현황 관리 시스템")

app.include_router(inventory.router,   prefix="/api/inventory",   tags=["재고수불"])
app.include_router(warehouse.router,   prefix="/api/warehouse",   tags=["창고재고실사"])
app.include_router(outgoing.router,    prefix="/api/outgoing",    tags=["양품출고"])
app.include_router(incoming.router,    prefix="/api/incoming",    tags=["입고현황"])
app.include_router(shipment.router,    prefix="/api/shipment",    tags=["출하현황"])
app.include_router(purchase.router,    prefix="/api/purchase",    tags=["매입영수증"])
app.include_router(ledger.router,      prefix="/api/ledger",      tags=["재고원장"])
app.include_router(evaluation.router,  prefix="/api/evaluation",  tags=["재고평가"])
app.include_router(incentive.router,   prefix="/api/incentive",   tags=["인센티브"])

app.mount("/static", StaticFiles(directory="static"), name="static")

@app.get("/")
def index():
    return FileResponse("static/index.html")
