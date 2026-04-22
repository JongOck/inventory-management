from fastapi import APIRouter, Query
from database import query
from typing import Optional

router = APIRouter()

@router.get("/")
def get_evaluation(work_year: Optional[str] = Query(None)):
    sql = "SELECT * FROM inventory_evaluation WHERE reference_year = %s ORDER BY item_code"
    return query(sql, (work_year,))
