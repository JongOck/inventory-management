from fastapi import APIRouter, Query
from database import query
from typing import Optional

router = APIRouter()

@router.get("/")
def get_incentive(work_month: Optional[str] = Query(None)):
    month_fmt = f"{work_month[:4]}/{work_month[4:6]}" if work_month and len(work_month) >= 6 else ""
    sql = "SELECT * FROM mds_incentive_result WHERE reference_month = %s ORDER BY item_code"
    return query(sql, (month_fmt,))
