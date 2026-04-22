from fastapi import APIRouter, Query
from database import query
from typing import Optional

router = APIRouter()

@router.get("/")
def get_ledger(work_month: Optional[str] = Query(None)):
    """재고 원장 (treeview6)"""
    sql = """
        SELECT *
        FROM monthly_inventory_status
        WHERE reference_month = %s
        ORDER BY item_code
    """
    return query(sql, (work_month,))
