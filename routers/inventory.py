from fastapi import APIRouter, Query
from database import query
from typing import Optional

router = APIRouter()

@router.get("/")
def get_inventory(work_month: Optional[str] = Query(None)):
    """재고수불 마스터 (treeview0)"""
    sql = """
        SELECT
            item_code, item_name, specification, unit, category,
            beginning_unit_price, beginning_quantity, beginning_amount,
            incoming_unit_price, incoming_quantity, incoming_amount,
            misc_profit_amount, incentive_amount,
            transfer_in_free_quantity, transfer_in_code_change_quantity, transfer_in_code_change_amount,
            outgoing_unit_price, outgoing_quantity, outgoing_amount,
            transfer_out_donation_quantity, transfer_out_donation_amount,
            transfer_out_free_quantity, transfer_out_free_amount,
            transfer_out_internal_use_quantity, transfer_out_internal_use_amount,
            transfer_out_sample_quantity, transfer_out_sample_amount,
            transfer_out_loss_quantity, transfer_out_loss_amount,
            transfer_out_expired_quantity, transfer_out_expired_amount,
            transfer_out_inventory_adjustment_quantity, transfer_out_inventory_adjustment_amount,
            current_unit_price, current_quantity, current_amount,
            verification_quantity, verification_amount
        FROM mds_basic_data
        WHERE reference_year = %s
        ORDER BY item_code
    """
    month = work_month or ""
    year = month[:4] if len(month) >= 4 else ""
    return query(sql, (year,))

@router.get("/summary")
def get_inventory_summary(work_month: Optional[str] = Query(None)):
    """재고수불 합계"""
    sql = """
        SELECT
            COUNT(*) as total_items,
            SUM(beginning_amount) as total_beginning_amount,
            SUM(incoming_amount) as total_incoming_amount,
            SUM(current_amount) as total_current_amount
        FROM mds_basic_data
        WHERE reference_year = %s
    """
    month = work_month or ""
    year = month[:4] if len(month) >= 4 else ""
    return query(sql, (year,))
