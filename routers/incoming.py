from fastapi import APIRouter, Query
from database import query
from typing import Optional

router = APIRouter()

@router.get("/")
async def get_incoming(work_month: Optional[str] = Query(None)):
    sql = "SELECT * FROM purchase_receipt_status WHERE TO_CHAR(receipt_date, 'YYYY-MM') = $1 ORDER BY receipt_date"
    return await query(sql, (work_month,))
