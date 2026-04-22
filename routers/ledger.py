from fastapi import APIRouter, Query
from database import query
from typing import Optional

router = APIRouter()

@router.get("/")
async def get_ledger(work_month: Optional[str] = Query(None)):
    sql = "SELECT * FROM monthly_inventory_status WHERE reference_month = $1 ORDER BY item_code"
    return await query(sql, (work_month,))
