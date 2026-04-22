from fastapi import APIRouter, Query
from database import query
from typing import Optional

router = APIRouter()

@router.get("/")
async def get_inventory(work_month: Optional[str] = Query(None)):
    month = work_month or ""
    year = month[:4] if len(month) >= 4 else ""
    sql = """
        SELECT item_code, item_name, specification, unit, category,
            beginning_unit_price, beginning_quantity, beginning_amount,
            incoming_unit_price, incoming_quantity, incoming_amount,
            misc_profit_amount, incentive_amount,
            outgoing_unit_price, outgoing_quantity, outgoing_amount,
            current_unit_price, current_quantity, current_amount
        FROM mds_basic_data
        WHERE reference_year = $1
        ORDER BY item_code
    """
    return await query(sql, (year,))
