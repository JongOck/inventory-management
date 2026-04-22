from fastapi import APIRouter
from database import query

router = APIRouter()

@router.get("/")
async def get_warehouse():
    sql = "SELECT * FROM master ORDER BY item_code"
    return await query(sql)
