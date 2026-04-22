from fastapi import APIRouter, Query
from database import query
from typing import Optional

router = APIRouter()

@router.get("/")
def get_incentive(work_month: Optional[str] = Query(None)):
    """인센티브 배분 (treeview8)"""
    sql = """
        SELECT *
        FROM account_substitution
        WHERE TO_CHAR(reference_date, 'YYYY-MM') = %s
        ORDER BY item_code
    """
    return query(sql, (work_month,))
