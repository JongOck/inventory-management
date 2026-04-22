from fastapi import APIRouter, Query
from database import query
from typing import Optional

router = APIRouter()

@router.get("/")
def get_outgoing(work_month: Optional[str] = Query(None)):
    """양품 출고 현황 (treeview2)"""
    sql = """
        SELECT *
        FROM shipment_status
        WHERE TO_CHAR(shipment_date, 'YYYY-MM') = %s
        ORDER BY shipment_date, item_code
    """
    return query(sql, (work_month,))
