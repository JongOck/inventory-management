from fastapi import APIRouter, Query
from database import query
from typing import Optional

router = APIRouter()

@router.get("/")
def get_inventory(work_month: Optional[str] = Query(None)):
    month = work_month or ""
    year = month[:4] if len(month) >= 4 else ""
    # reference_year 형식: '2025', '2025/06', '2025-06' 모두 처리
    ym_slash = f"{month[:4]}/{month[4:6]}" if len(month) >= 6 else ""
    ym_dash  = f"{month[:4]}-{month[4:6]}" if len(month) >= 6 else ""
    sql = """
        SELECT item_code, item_name, specification, unit, category,
            beginning_unit_price, beginning_quantity, beginning_amount,
            reference_year
        FROM mds_basic_data
        WHERE reference_year IN (%s, %s, %s)
        ORDER BY item_code
    """
    return query(sql, (year, ym_slash, ym_dash))
