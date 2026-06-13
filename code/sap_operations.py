# Re-exports de compatibilidad — este módulo fue dividido en:
#   sap_utils.py   -> utilidades y constantes compartidas
#   sap_vl06f.py   -> lectura de VL06F y eliminación de BOL
#   sap_batches.py -> operaciones VF11, VI05, VT02N, VL09
#   sap_orders.py  -> ZCMR, VA02, ME22N

from sap_utils import (
    _navigate_to,
    _go_back,
    _wait_ready,
    _enter_multi_values,
    _POPUP_TABLE,
    _wnd_exists,
    _find_popup_wnd,
    _popup_table,
)
from sap_vl06f import read_vl06f_data, delete_bol
from sap_batches import (
    delete_billing_document,
    delete_billing_documents_bulk,
    delete_shipment_cost,
    delete_shipment_costs_all,
    delete_shipment_number,
    delete_shipment_numbers_all,
    reverse_pgi,
    reverse_pgi_bulk,
)
from sap_orders import delete_orders_from_zcmr

__all__ = [
    "_navigate_to", "_go_back", "_wait_ready", "_enter_multi_values", "_POPUP_TABLE",
    "_wnd_exists", "_find_popup_wnd", "_popup_table",
    "read_vl06f_data", "delete_bol",
    "delete_billing_document", "delete_billing_documents_bulk",
    "delete_shipment_cost", "delete_shipment_costs_all",
    "delete_shipment_number", "delete_shipment_numbers_all",
    "reverse_pgi", "reverse_pgi_bulk",
    "delete_orders_from_zcmr",
]
