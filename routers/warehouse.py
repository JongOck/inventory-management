from fastapi import APIRouter, Query
from database import query
from typing import Optional

router = APIRouter()

@router.get("/")
def get_warehouse(work_month: Optional[str] = Query(None)):
    """창고별 재고실사 현황 (treeview1)"""
    sql = """
        SELECT
            item_code, item_name, specification, unit,
            차산점, 차산점A, 수입창고, 청량리점, 이천점,
            케이터링, 하남점, 이커머스, 선매입창고,
            차산점반품, 청량리반품, 이천점반품, 하남점반품,
            차산점폐기, 이천점폐기,
            last_updated
        FROM master
        ORDER BY item_code
    """
    return query(sql)

@router.get("/last-updated")
def get_last_updated():
    """마지막 업데이트 시간"""
    sql = "SELECT MAX(last_updated) as last_updated FROM master"
    return query(sql)
