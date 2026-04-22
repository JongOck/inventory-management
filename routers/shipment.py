from fastapi import APIRouter, Query
from database import query
from typing import Optional

router = APIRouter()

@router.get("/")
async def get_shipment(work_month: Optional[str] = Query(None)):
    sql = "SELECT * FROM shipment_status WHERE TO_CHAR(shipment_date, 'YYYY-MM') = $1 ORDER BY shipment_date"
    return await query(sql, (work_month,))
