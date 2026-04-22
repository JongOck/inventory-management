from fastapi import APIRouter
from database import query

router = APIRouter()

@router.get("/")
def get_warehouse():
    return query("SELECT * FROM master ORDER BY item_code")
