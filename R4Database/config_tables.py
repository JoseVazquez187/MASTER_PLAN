# config_tables.py - Configuración para múltiples tablas
# VERSIÓN COMPLETA SIN ERRORES - Con soporte para todas las tablas

import os
from pathlib import Path

# Configuración centralizada para todas las tablas
TABLES_CONFIG = {
# ==================== TABLA Credit_Memos ====================


# ==================== TABLA SALES_ORDER_TABLE ====================
"sales_order_table": {
    "source_file": r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\WHS PLAN\FILES\SalesOrder\sales_order.csv",
    "table_name": "sales_order_table",
    "file_type": "csv",
    "excel_params": {
        "skiprows": 3,
        "dtype": "str"
    },
    "expected_columns": [
        'Entity', 'Proj', 'SO-No', 'Ln', 'Ord-Cd', 'Spr-CD', 'Order-Dt',
        'Cancel-Dt', 'CustReqDt', 'Req-Dt', 'Pln-Ship', 'Prd-Dt', 'Inv-Dt',
        'Recover-Dt', 'TimeToConfirm(hr)', 'Std LT', 'Cust#', 'Cust-Name',
        'Type-Code', 'Cust-PO', 'Buyer Name', 'A/C', 'Item-Number',
        'Description', 'Item-Rev', 'UM', 'ML', 'Orig-Q', 'Opn-Q', 'OH Netable',
        'OH Non Netable', 'TR', 'Issue-Q', 'Plnr-Intl', 'Plnr-Code', 'PlanType',
        'Family', 'Misc-Code', 'WO-No', 'WO Due-dt', 'Unit Price', 'OpenValue',
        'Curr', 'T-Desc', 'So-Ln Memo', 'WO-Memo', 'ST', 'AC-Priority',
        'T-Card', 'Std-Cost', 'Cust-NCR', 'Reason-Cd', 'User-ID', 'Invoice-No',
        'Shipped-Date', 'Ship-no', 'Ship-via', 'Tracking-No', 'Vendor-Code',
        'Vendor Name', 'DG', 'Planner Notes', 'Responsibility', 'Last Notes',
        'Manuf Charge', 'Memo 1', 'Memo 2'
    ],
    "columns_mapping": {
        'SO-No': 'SO_No',
        'Ord-Cd': 'Ord_Cd',
        'Spr-CD': 'Spr_CD',
        'Order-Dt': 'Order_Dt',
        'Cancel-Dt': 'Cancel_Dt',
        'Req-Dt': 'Req_Dt',
        'Pln-Ship': 'Pln_Ship',
        'Prd-Dt': 'Prd_Dt',
        'Inv-Dt': 'Inv_Dt',
        'Recover-Dt': 'Recover_Dt',
        'TimeToConfirm(hr)': 'TimeToConfirm_hr',
        'Std LT': 'Std_LT',
        'Cust#': 'Cust',
        'Cust-Name': 'Cust_Name',
        'Type-Code': 'Type_Code',
        'Cust-PO': 'Cust_PO',
        'Buyer Name': 'Buyer_Name',
        'A/C': 'AC',
        'Item-Number': 'Item_Number',
        'Item-Rev': 'Item_Rev',
        'Orig-Q': 'Orig_Q',
        'Opn-Q': 'Opn_Q',
        'OH Netable': 'OH_Netable',
        'OH Non Netable': 'OH_Non_Netable',
        'Issue-Q': 'Issue_Q',
        'Plnr-Intl': 'Plnr_Intl',
        'Plnr-Code': 'Plnr_Code',
        'Misc-Code': 'Misc_Code',
        'WO-No': 'WO_No',
        'WO Due-dt': 'WO_Due_dt',
        'Unit Price': 'Unit_Price',
        'T-Desc': 'T_Desc',
        'So-Ln Memo': 'So_Ln_Memo',
        'WO-Memo': 'WO_Memo',
        'AC-Priority': 'AC_Priority',
        'T-Card': 'T_Card',
        'Std-Cost': 'Std_Cost',
        'Cust-NCR': 'Cust_NCR',
        'Reason-Cd': 'Reason_Cd',
        'User-ID': 'User_ID',
        'Invoice-No': 'Invoice_No',
        'Shipped-Date': 'Shipped_Date',
        'Ship-no': 'Ship_no',
        'Ship-via': 'Ship_via',
        'Tracking-No': 'Tracking_No',
        'Vendor-Code': 'Vendor_Code',
        'Vendor Name': 'Vendor_Name',
        'Planner Notes': 'Planner_Notes',
        'Last Notes': 'Last_Notes',
        'Manuf Charge': 'Manuf_Charge',
        'Memo 1': 'Memo_1',
        'Memo 2': 'Memo_2'
    },
    "columns_order_original": [
        'Entity', 'Proj', 'SO-No', 'Ln', 'Ord-Cd', 'Spr-CD', 'Order-Dt',
        'Cancel-Dt', 'CustReqDt', 'Req-Dt', 'Pln-Ship', 'Prd-Dt', 'Inv-Dt',
        'Recover-Dt', 'TimeToConfirm(hr)', 'Std LT', 'Cust#', 'Cust-Name',
        'Type-Code', 'Cust-PO', 'Buyer Name', 'A/C', 'Item-Number',
        'Description', 'Item-Rev', 'UM', 'ML', 'Orig-Q', 'Opn-Q', 'OH Netable',
        'OH Non Netable', 'TR', 'Issue-Q', 'Plnr-Intl', 'Plnr-Code', 'PlanType',
        'Family', 'Misc-Code', 'WO-No', 'WO Due-dt', 'Unit Price', 'OpenValue',
        'Curr', 'T-Desc', 'So-Ln Memo', 'WO-Memo', 'ST', 'AC-Priority',
        'T-Card', 'Std-Cost', 'Cust-NCR', 'Reason-Cd', 'User-ID', 'Invoice-No',
        'Shipped-Date', 'Ship-no', 'Ship-via', 'Tracking-No', 'Vendor-Code',
        'Vendor Name', 'DG', 'Planner Notes', 'Responsibility', 'Last Notes',
        'Manuf Charge', 'Memo 1', 'Memo 2'
    ],
    "columns_order_renamed": [
        'Entity', 'Proj', 'SO_No', 'Ln', 'Ord_Cd', 'Spr_CD', 'Order_Dt',
        'Cancel_Dt', 'CustReqDt', 'Req_Dt', 'Pln_Ship', 'Prd_Dt', 'Inv_Dt',
        'Recover_Dt', 'TimeToConfirm_hr', 'Std_LT', 'Cust', 'Cust_Name',
        'Type_Code', 'Cust_PO', 'Buyer_Name', 'AC', 'Item_Number',
        'Description', 'Item_Rev', 'UM', 'ML', 'Orig_Q', 'Opn_Q', 'OH_Netable',
        'OH_Non_Netable', 'TR', 'Issue_Q', 'Plnr_Intl', 'Plnr_Code', 'PlanType',
        'Family', 'Misc_Code', 'WO_No', 'WO_Due_dt', 'Unit_Price', 'OpenValue',
        'Curr', 'T_Desc', 'So_Ln_Memo', 'WO_Memo', 'ST', 'AC_Priority',
        'T_Card', 'Std_Cost', 'Cust_NCR', 'Reason_Cd', 'User_ID', 'Invoice_No',
        'Shipped_Date', 'Ship_no', 'Ship_via', 'Tracking_No', 'Vendor_Code',
        'Vendor_Name', 'DG', 'Planner_Notes', 'Responsibility', 'Last_Notes',
        'Manuf_Charge', 'Memo_1', 'Memo_2'
    ],
    "special_processing": {
        "clear_before_insert": True,
        "validate_columns": True,
        "custom_cleaning": True,
        "column_name_cleaning": True,
        "numeric_conversion": ["TimeToConfirm_hr", "Std_LT", "Orig_Q", "Opn_Q", "OH_Netable", "OH_Non_Netable", "TR", "Issue_Q", "Unit_Price", "OpenValue", "Std_Cost"],
        "auto_detect_header": True
    },
    "create_table_sql": """CREATE TABLE IF NOT EXISTS sales_order_table(
        Entity TEXT,
        Proj TEXT,
        SO_No TEXT,
        Ln TEXT,
        Ord_Cd TEXT,
        Spr_CD TEXT,
        Order_Dt TEXT,
        Cancel_Dt TEXT,
        CustReqDt TEXT,
        Req_Dt TEXT,
        Pln_Ship TEXT,
        Prd_Dt TEXT,
        Inv_Dt TEXT,
        Recover_Dt TEXT,
        TimeToConfirm_hr REAL,
        Std_LT REAL,
        Cust TEXT,
        Cust_Name TEXT,
        Type_Code TEXT,
        Cust_PO TEXT,
        Buyer_Name TEXT,
        AC TEXT,
        Item_Number TEXT,
        Description TEXT,
        Item_Rev TEXT,
        UM TEXT,
        ML TEXT,
        Orig_Q REAL,
        Opn_Q REAL,
        OH_Netable REAL,
        OH_Non_Netable REAL,
        TR REAL,
        Issue_Q REAL,
        Plnr_Intl TEXT,
        Plnr_Code TEXT,
        PlanType TEXT,
        Family TEXT,
        Misc_Code TEXT,
        WO_No TEXT,
        WO_Due_dt TEXT,
        Unit_Price REAL,
        OpenValue REAL,
        Curr TEXT,
        T_Desc TEXT,
        So_Ln_Memo TEXT,
        WO_Memo TEXT,
        ST TEXT,
        AC_Priority TEXT,
        T_Card TEXT,
        Std_Cost REAL,
        Cust_NCR TEXT,
        Reason_Cd TEXT,
        User_ID TEXT,
        Invoice_No TEXT,
        Shipped_Date TEXT,
        Ship_no TEXT,
        Ship_via TEXT,
        Tracking_No TEXT,
        Vendor_Code TEXT,
        Vendor_Name TEXT,
        DG TEXT,
        Planner_Notes TEXT,
        Responsibility TEXT,
        Last_Notes TEXT,
        Manuf_Charge TEXT,
        Memo_1 TEXT,
        Memo_2 TEXT
    )"""
},
# ==================== TABLA EXPEDITE ====================
"expedite": {
    "source_file": r"J:\Departments\Material Control\Purchasing\Tools\ComprasDB\Expedite.csv",
    "table_name": "expedite",
    "columns_mapping": {
        'Entity Group': 'EntityGroup',
        'A/C': 'AC',
        'Item-No': 'ItemNo',
        'Req-Qty': 'ReqQty',
        'Demand-Source': 'DemandSource',
        'Req-Date': 'ReqDate',
        'Ship-Date': 'ShipDate',
        'MLIK Code': 'MLIKCode',
        'STD-Cost': 'STDCost',
        'Lot-Size': 'LotSize',
        'Fill-Doc': 'FillDoc',
        'Demand - Type': 'DemandType'
    },
    "columns_order_original": [
        'Entity Group', 'Project', 'A/C', 'Item-No', 'Description',
        'PlanTp', 'Ref', 'Sub', 'Sort', 'Fill-Doc', 'Demand - Type', 'Req-Qty',
        'Demand-Source', 'Unit', 'Vendor', 'Req-Date', 'Ship-Date', 'OH', 
        'MLIK Code', 'LT', 'STD-Cost', 'Lot-Size', 'UOM'
    ],
    "columns_order_renamed": [
        'EntityGroup', 'Project', 'AC', 'ItemNo', 'Description',
        'PlanTp', 'Ref', 'Sub', 'Sort', 'FillDoc', 'DemandType', 'ReqQty',
        'DemandSource', 'Unit', 'Vendor', 'ReqDate', 'ShipDate', 'OH', 
        'MLIKCode', 'LT', 'STDCost', 'LotSize', 'UOM'
    ],
    "create_table_sql": """CREATE TABLE IF NOT EXISTS expedite(
        id INTEGER PRIMARY KEY,
        EntityGroup TEXT,
        Project TEXT,
        AC TEXT,
        ItemNo TEXT,
        Description TEXT,
        PlanTp TEXT,
        Ref TEXT,
        Sub TEXT,
        FillDoc TEXT,
        DemandType TEXT,
        Sort TEXT,
        ReqQty TEXT,
        DemandSource TEXT,
        Unit TEXT,
        Vendor TEXT,
        ReqDate TEXT,
        ShipDate TEXT,
        OH TEXT,
        MLIKCode TEXT,
        LT TEXT,
        STDCost TEXT,
        LotSize TEXT,
        UOM TEXT
    )"""
},
# ==================== TABLA IN92 ====================
"in92": {
    "source_file": r"J:\Departments\Material Control\Purchasing\Tools\ComprasDB\in92.txt",
    "table_name": "in92",
    "file_type": "txt_delimited",
    "delimiter": "@",
    "columns_mapping": {
        'ITEM NUMBER': 'ItemNo',
        'Vendor CD': 'Vendor',
        'L/T': 'LT',
        'STD COST 1': 'STDCost',
        'MLI': 'MLIKCode',
        'LOT SIZE': 'LotSize'
    },
    "columns_order_original": [
        'ITEM NUMBER', 'Vendor CD', 'L/T', 'STD COST 1', 'PLANTYPE', 'MLI', 'ACTIVE', 'LOT SIZE'
    ],
    "columns_order_renamed": [
        'ItemNo', 'Vendor', 'LT', 'STDCost', 'PLANTYPE', 'MLIKCode', 'ACTIVE', 'LotSize'
    ],
    "uppercase_columns": [
        'ItemNo', 'Vendor', 'PLANTYPE', 'MLIKCode', 'ACTIVE'
    ],
    "create_table_sql": """CREATE TABLE IF NOT EXISTS in92(
        id INTEGER PRIMARY KEY,
        ItemNo TEXT,
        Vendor TEXT,
        LT TEXT,
        STDCost TEXT,
        PLANTYPE TEXT,
        MLIKCode TEXT,
        ACTIVE TEXT,
        LotSize TEXT
    )"""
},
# ==================== TABLA IN521 ====================
"in521": {
    "source_file": r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\FEFO\FEFO_in521\in521.txt",
    "table_name": "in521",
    "file_type": "fixed_width",
    "fixed_width_params": {
        "widths": [2, 32, 21, 3, 5, 11, 10, 10, 8, 20, 9, 5, 5, 10, 9, 5, 10, 15],
        "header": 3,
        "skip_rows": [0]
    },
    "data_filters": {
        "exclude_rows": {"Wh": "En"}
    },
    "columns_mapping": {
        'Item-No': 'ItemNo',
        'Exp-Date': 'ExpDate', 
        'Rec.#': 'Rec',
        'Qty-OH': 'QtyOH',
        'Order By': 'OrderBy',
        'SL %': 'Shelf_Life'
    },
    "columns_order_original": [
        'Wh', 'Item-No', 'Description', 'UM', 'Prod', 'Exp-Date', 'Rec.#', 'Qty-OH',
        'Bin', 'Lot', 'Order By', 'LRev', 'IRev', 'PlanType', 'EntCode', 'Proj',
        'ManufDate', 'SL %'
    ],
    "columns_order_renamed": [
        'Wh', 'ItemNo', 'Description', 'UM', 'Prod', 'ExpDate', 'Rec', 'QtyOH',
        'Bin', 'Lot', 'OrderBy', 'LRev', 'IRev', 'PlanType', 'EntCode', 'Proj',
        'ManufDate', 'Shelf_Life'
    ],
    "create_table_sql": """CREATE TABLE IF NOT EXISTS in521(
        id INTEGER PRIMARY KEY,
        Wh TEXT,
        ItemNo TEXT,
        Description TEXT,
        UM TEXT,
        Prod TEXT,
        ExpDate TEXT,
        Rec TEXT,
        QtyOH TEXT,
        Bin TEXT,
        Lot TEXT,
        OrderBy TEXT,
        LRev TEXT,
        IRev TEXT,
        PlanType TEXT,
        EntCode TEXT,
        Proj TEXT,
        ManufDate TEXT,
        Shelf_Life TEXT
    )"""
},
# ==================== TABLA PO351 ====================
"po351": {
    "source_file": r"J:\Departments\Material Control\Purchasing\Tools\ComprasDB\po351.txt",
    "table_name": "po351",
    "file_type": "txt_delimited",
    "delimiter": "@",
    "columns_mapping": {
        'Item-No': 'ItemNo',
        'endor-code': 'Vendor',
        'Fob-vendor[1]': 'FobVendor'
    },
    "columns_order_original": [
        'Item-No', 'endor-code', 'Fob-vendor[1]'
    ],
    "columns_order_renamed": [
        'ItemNo', 'Vendor', 'FobVendor'
    ],
    "create_table_sql": """CREATE TABLE IF NOT EXISTS po351(
        id INTEGER PRIMARY KEY,
        ItemNo TEXT,
        Vendor TEXT,
        FobVendor TEXT
    )"""
},
# ==================== TABLA OH_GEMBA_SHORTAGES ====================
"oh_gemba_shortages": {
    "source_file": r"J:\Departments\Material Control\Purchasing\Tools\ComprasDB\ohNeteable.xlsx",
    "table_name": "oh_gemba_shortages",
    "file_type": "excel",
    "excel_params": {
        "skiprows": 4
    },
    "data_filters": {
        "exclude_rows": {"Bin": "QA"}
    },
    "columns_mapping": {
        'Item-Number': 'ItemNo'
    },
    "uppercase_columns": [
        'Bin'
    ],
    "columns_order_original": [
        'Item-Number', 'Description', 'Lot', 'OH', 'Bin'
    ],
    "columns_order_renamed": [
        'ItemNo', 'Description', 'Lot', 'OH', 'Bin'
    ],
    "create_table_sql": """CREATE TABLE IF NOT EXISTS oh_gemba_shortages(
        id INTEGER PRIMARY KEY,
        ItemNo TEXT,
        Description TEXT,
        Lot TEXT,
        OH TEXT,
        Bin TEXT
    )"""
},
# ==================== TABLA OH ====================
"oh": {
    "source_file": r"J:\Departments\Material Control\Purchasing\Tools\ComprasDB\ohNeteable.xlsx",
    "table_name": "oh",
    "file_type": "excel",
    "excel_params": {
        "skiprows": 4
    },
    "columns_mapping": {
        'Item-Number': 'ItemNo'
    },
    "columns_order_original": [
        'Item-Number', 'Description', 'Lot', 'OH'
    ],
    "columns_order_renamed": [
        'ItemNo', 'Description', 'Lot', 'OH'
    ],
    "create_table_sql": """CREATE TABLE IF NOT EXISTS oh(
        id INTEGER PRIMARY KEY,
        ItemNo TEXT,
        Description TEXT,
        Lot TEXT,
        OH TEXT
    )"""
},
# ==================== TABLA VENDOR ====================
"vendor": {
    "source_file": r"J:\Departments\Material Control\Purchasing\Tools\ComprasDB\vendor.xlsx",
    "table_name": "Vendor",
    "file_type": "excel",
    "columns_mapping": {
        'Vendor Code': 'Vendor',
        'Sort-name': 'Shortname'
    },
    "expected_columns": [
        'Vendor Code', 'Name', 'Shortname', 'Tactical', 'Operational', 'Process', 'Strategic'
    ],
    "columns_order_original": [
        'Vendor Code', 'Name', 'Shortname', 'Tactical', 'Operational', 'Process', 'Strategic'
    ],
    "columns_order_renamed": [
        'Vendor', 'Name', 'Shortname', 'Tactical', 'Operational', 'Process', 'Strategic'
    ],
    "create_table_sql": """CREATE TABLE IF NOT EXISTS Vendor(
        id INTEGER PRIMARY KEY,
        Vendor TEXT,
        Name TEXT,
        Shortname TEXT,
        Tactical TEXT,
        Operational TEXT,
        Process TEXT,
        Strategic TEXT
    )"""
},
# ==================== TABLA HOLIDAY ====================
"holiday": {
    "source_file": r"J:\Departments\Material Control\Purchasing\Tools\ComprasDB\holiday.xlsx",
    "table_name": "Holiday",
    "file_type": "excel",
    "columns_order_original": [
        'Holiday', 'Day'
    ],
    "columns_order_renamed": [
        'Holiday', 'Day'
    ],
    "create_table_sql": """CREATE TABLE IF NOT EXISTS Holiday(
        id INTEGER PRIMARY KEY,
        Holiday TEXT,
        Day DATE
    )"""
},
# ==================== TABLA OPENORDER ====================
"openorder": {
    "source_file": r"J:\Departments\Material Control\Purchasing\Tools\ComprasDB\openOrder.xlsx",
    "table_name": "openOrder",
    "file_type": "excel",
    "excel_params": {
        "skiprows": 3
    },
    "uppercase_columns": [
        'ItemNo'
    ],
    "columns_mapping": {
        'Entity-Code': 'EntityCode',
        'PO-No': 'PONo',
        'WO-No': 'WONo',
        'Vendor #': 'Vendor',
        'Vendor Name': 'VendorName',
        'Item-Number': 'ItemNo',
        'Item-Description': 'ItemDescription',
        'U/M': 'UM',
        'Ord-Q': 'OrdQ',
        'Rcvd-Q': 'RcvdQ',
        'Opn-Q': 'OpnQ',
        'Std-cost': 'Stdcost',
        'Unit-$': 'Unit$',
        'Opn-Val': 'OpnVal',
        'Net-Opn-Val': 'NetOpnVal',
        'PO-Dt': 'PODt',
        'Req-Dt': 'ReqDt',
        'Prom-Dt': 'PromDt',
        'Rev-Pr-Dt': 'RevPrDt',
        'Buyer Name': 'BuyerName',
        'Line-Memo': 'LineMemo',
        'Last Note/Status': 'NoteStatus'
    },
    "columns_order_original": [
        'Entity-Code', 'Proj', 'PO-No', 'Ln', 'Type', 'WO-No', 'Vendor #',
        'Vendor Name', 'Item-Number', 'Item-Description', 'Rev', 'U/M', 'Ord-Q', 'Rcvd-Q', 'Opn-Q',
        'Std-cost', 'Unit-$', 'Opn-Val', 'Net-Opn-Val', 'PO-Dt', 'Req-Dt', 'Prom-Dt', 'Rev-Pr-Dt', 'Buyer Name',
        'LT', 'Line-Memo', 'Last Note/Status'
    ],
    "columns_order_renamed": [
        'EntityCode', 'Proj', 'PONo', 'Ln', 'Type', 'WONo', 'Vendor',
        'VendorName', 'ItemNo', 'ItemDescription', 'Rev', 'UM', 'OrdQ', 'RcvdQ', 'OpnQ',
        'Stdcost', 'Unit$', 'OpnVal', 'NetOpnVal', 'PODt', 'ReqDt', 'PromDt', 'RevPrDt', 'BuyerName',
        'LT', 'LineMemo', 'NoteStatus'
    ],
    "create_table_sql": """CREATE TABLE IF NOT EXISTS openOrder(
        id INTEGER PRIMARY KEY,
        EntityCode TEXT,
        Proj TEXT,
        PONo TEXT,
        Ln TEXT,
        Type TEXT,
        WONo TEXT,
        Vendor TEXT,
        VendorName TEXT,
        ItemNo TEXT,
        ItemDescription TEXT,
        Rev TEXT,
        UM TEXT,
        OrdQ TEXT,
        RcvdQ TEXT,
        OpnQ TEXT,
        Stdcost TEXT,
        Unit$ TEXT,
        OpnVal TEXT,
        NetOpnVal TEXT,
        PODt TEXT,
        ReqDt TEXT,
        PromDt TEXT,
        RevPrDt TEXT,
        BuyerName TEXT,
        LT TEXT,
        LineMemo TEXT,
        NoteStatus TEXT
    )"""
},
# ==================== TABLA REWORKLOC_ALL ====================
"reworkloc_all": {
    "source_file": r"J:\Departments\Material Control\Purchasing\Tools\ComprasDB\reworkLoc.xlsx",
    "table_name": "ReworkLoc_all",
    "file_type": "excel",
    "excel_params": {
        "skiprows": 2
    },
    "columns_mapping": {
        'Item-Number': 'ItemNo',
        'Expire-Date': 'Expire_Date'
    },
    "special_processing": {
        "fill_expire_date": True,
        "expire_years": 2
    },
    "columns_order_original": [
        'Item-Number', 'Description', 'Lot', 'Expire-Date', 'OH', 'Bin'
    ],
    "columns_order_renamed": [
        'ItemNo', 'Description', 'Lot', 'Expire_Date', 'OH', 'Bin'
    ],
    "create_table_sql": """CREATE TABLE IF NOT EXISTS ReworkLoc_all(
        id INTEGER PRIMARY KEY,
        ItemNo TEXT,
        Description TEXT,
        Lot TEXT,
        Expire_Date TEXT,
        OH TEXT,
        Bin TEXT
    )"""
},
# ==================== TABLA REWORKLOC ====================
"reworkloc": {
    "source_file": r"J:\Departments\Material Control\Purchasing\Tools\ComprasDB\reworkLoc.xlsx",
    "table_name": "ReworkLoc",
    "file_type": "excel",
    "excel_params": {
        "skiprows": 2
    },
    "columns_mapping": {
        'Item-Number': 'ItemNo'
    },
    "columns_order_original": [
        'Item-Number', 'Description', 'Lot', 'OH', 'Bin'
    ],
    "columns_order_renamed": [
        'ItemNo', 'Description', 'Lot', 'OH', 'Bin'
    ],
    "create_table_sql": """CREATE TABLE IF NOT EXISTS ReworkLoc(
        id INTEGER PRIMARY KEY,
        ItemNo TEXT,
        Description TEXT,
        Lot TEXT,
        OH TEXT,
        Bin TEXT
    )"""
},
# ==================== TABLA ACTIONMESSAGES ====================
"actionmessages": {
    "source_file": r"J:\Departments\Material Control\Purchasing\Tools\ComprasDB\actionMessages.xlsx",
    "table_name": "ActionMessages",
    "file_type": "excel",
    "excel_params": {
        "skiprows": 3
    },
    "columns_mapping": {
        'ITEM-NO': 'ItemNo',
        'ITEM-DESC': 'ItemDesc',
        'Plnr Code': 'PlnrCode',
        'Std-Cost': 'StdCost',
        'REF': 'PO',
        'SUB': 'Line',
        'A-CD': 'ACD',
        'Act-Qty': 'ActQty',
        'Fm-Qty': 'FmQty',
        'To-Qty': 'ToQty',
        'A-DT': 'ADT',
        'REQ-DT': 'REQDT',
        'PO Line Note': 'POLineNote',
        'Planning Notes': 'PlanningNotes'
    },
    "columns_order_original": [
        'ITEM-NO', 'ITEM-DESC', 'PlanTp', 'Plnr Code', 'Std-Cost', 'REF', 'SUB', 'A-CD',
        'Act-Qty', 'Fm-Qty', 'To-Qty', 'A-DT', 'REQ-DT', 'Vendor', 'Type', 'OH', 'MLI', 'PO Line Note', 'Planning Notes'
    ],
    "columns_order_renamed": [
        'ItemNo', 'ItemDesc', 'PlanTp', 'PlnrCode', 'StdCost', 'PO', 'Line', 'ACD',
        'ActQty', 'FmQty', 'ToQty', 'ADT', 'REQDT', 'Vendor', 'Type', 'OH', 'MLI', 'POLineNote', 'PlanningNotes'
    ],
    "create_table_sql": """CREATE TABLE IF NOT EXISTS ActionMessages(
        id INTEGER PRIMARY KEY,
        ItemNo TEXT,
        ItemDesc TEXT,
        PlanTp TEXT,
        PlnrCode TEXT,
        StdCost TEXT,
        PO TEXT,
        Line TEXT,
        ACD TEXT,
        ActQty TEXT,
        FmQty TEXT,
        ToQty TEXT,
        ADT TEXT,
        REQDT TEXT,
        Vendor TEXT,
        Type TEXT,
        OH TEXT,
        MLI TEXT,
        POLineNote TEXT,
        PlanningNotes TEXT
    )"""
},
# ==================== TABLA ENTITY ====================
"entity": {
    "source_file": r"J:\Departments\Material Control\Purchasing\Tools\ComprasDB\entity.xlsx",
    "table_name": "entity",
    "file_type": "excel",
    "columns_mapping": {
        'Entity': 'EntityGroup'
    },
    "columns_order_original": [
        'Entity', 'Name', 'EntityName', 'Commodity'
    ],
    "columns_order_renamed": [
        'EntityGroup', 'Name', 'EntityName', 'Commodity'
    ],
    "create_table_sql": """CREATE TABLE IF NOT EXISTS entity(
        id INTEGER PRIMARY KEY,
        EntityGroup TEXT,
        Name TEXT,
        EntityName TEXT,
        Commodity TEXT
    )"""
},
# ==================== TABLA TRANSIT ====================
"transit": {
    "source_file": r"J:\Departments\Material Control\Purchasing\Tools\ComprasDB\transit_table.xlsx",
    "table_name": "transit",
    "file_type": "excel",
    "columns_order_original": [
        'Vendor', 'VendorName', 'TransitTime'
    ],
    "columns_order_renamed": [
        'Vendor', 'VendorName', 'TransitTime'
    ],
    "create_table_sql": """CREATE TABLE IF NOT EXISTS transit(
        id INTEGER PRIMARY KEY,
        Vendor TEXT,
        VendorName TEXT,
        TransitTime TEXT
    )"""
},
# ==================== TABLA ENTITY_PROJECT ====================
"entity_project": {
    "source_file": r"J:\Departments\Material Control\Purchasing\Tools\ComprasDB\project_entity.xlsx",
    "table_name": "entity_project",
    "file_type": "excel",
    "columns_order_original": [
        'EntityGroup', 'Project'
    ],
    "columns_order_renamed": [
        'EntityGroup', 'Project'
    ],
    "create_table_sql": """CREATE TABLE IF NOT EXISTS entity_project(
        id INTEGER PRIMARY KEY,
        EntityGroup TEXT,
        Project TEXT
    )"""
},
# ==================== TABLA FCST_SUPPLIER ====================
"fcst_supplier": {
    "source_file": r"J:\Departments\Material Control\Purchasing\Tools\ComprasDB\supplier_mail.xlsx",
    "table_name": "fcst_supplier",
    "file_type": "excel",
    "columns_order_original": [
        'Vendor', 'Shortname', 'Owner1', 'Owner2', 'Owner3', 'email', 'email2', 'email3'
    ],
    "columns_order_renamed": [
        'Vendor', 'Shortname', 'Owner1', 'Owner2', 'Owner3', 'email', 'email2', 'email3'
    ],
    "create_table_sql": """CREATE TABLE IF NOT EXISTS fcst_supplier(
        id INTEGER PRIMARY KEY,
        Vendor TEXT,
        Shortname TEXT,
        Owner1 TEXT,
        Owner2 TEXT,
        Owner3 TEXT,
        email TEXT,
        email2 TEXT,
        email3 TEXT
    )"""
},
# ==================== TABLA KPT ====================
"kpt": {
    "source_file": r"J:\Departments\Material Control\Purchasing\Tools\ComprasDB\KPT.txt",
    "table_name": "kpt",
    "file_type": "fixed_width",
    "fixed_width_params": {
        "widths": [9, 13, 21, 21, 10, 7, 9, 9, 10, 12],
        "header": 3,
        "skip_rows": [0]
    },
    "data_filters": {
        "exclude_rows": {"In-entity": "End-of-Re"}
    },
    "columns_mapping": {
        'Entity-code': 'EntityGroup',
        'Sort-code': 'Sort',
        'Mfg-time': 'Mfgtime',
        'Route-Sort': 'RouteSort'
    },
    "columns_order_original": [
        'Entity-code', 'Unit', 'WorkCenter', 'Sort-code', 'Active', 'leadtime', 'Mfg-time', 'Route-Sort', 'Sequence'
    ],
    "columns_order_renamed": [
        'EntityGroup', 'Unit', 'WorkCenter', 'Sort', 'Active', 'leadtime', 'Mfgtime', 'RouteSort', 'Sequence'
    ],
    "create_table_sql": """CREATE TABLE IF NOT EXISTS kpt(
        id INTEGER PRIMARY KEY,
        EntityGroup TEXT,
        Unit TEXT,
        WorkCenter TEXT,
        Sort TEXT,
        Active TEXT,
        leadtime TEXT,
        Mfgtime TEXT,
        RouteSort TEXT,
        Sequence TEXT
    )"""
},
# ==================== TABLA QA_OH ====================
"qa_oh": {
    "source_file": r"J:\Departments\Material Control\Purchasing\Tools\ComprasDB\qaLoc.xlsx",
    "table_name": "qa_oh",
    "file_type": "excel",
    "excel_params": {
        "skiprows": 3
    },
    "columns_mapping": {
        'Item-Number': 'ItemNo'
    },
    "columns_order_original": [
        'Item-Number', 'Description', 'OH', 'Bin', 'Lot'
    ],
    "columns_order_renamed": [
        'ItemNo', 'Description', 'OH', 'Bin', 'Lot'
    ],
    "create_table_sql": """CREATE TABLE IF NOT EXISTS qa_oh(
        id INTEGER PRIMARY KEY,
        ItemNo TEXT,
        Description TEXT,
        OH TEXT,
        Bin TEXT,
        Lot TEXT
    )"""
},
# ==================== TABLA TABLE_YEARWEEK ====================
"table_yearweek": {
    "source_file": r"J:\Departments\Material Control\Purchasing\Tools\ComprasDB\year_week_table_dates.xlsx",
    "table_name": "table_yearweek",
    "file_type": "excel",
    "columns_order_original": [
        'year_week', 'Date'
    ],
    "columns_order_renamed": [
        'year_week', 'Date'
    ],
    "create_table_sql": """CREATE TABLE IF NOT EXISTS table_yearweek(
        id INTEGER PRIMARY KEY,
        year_week TEXT,
        Date TEXT
    )"""
},
# ==================== TABLA WOINQUIRY ====================
"woinquiry": {
    "source_file": r"J:\Departments\Material Control\Purchasing\Tools\ComprasDB\WO_Inquiry.xlsx",
    "table_name": "WOInquiry",
    "file_type": "excel",
    "excel_params": {
        "skiprows": 3
    },
    "columns_mapping": {
        'Project No': 'ProjectNo',
        'WO No': 'WONo',
        'SO/FCST': 'SO_FCST',
        'Parent WO': 'ParentWO',
        'Item-Number': 'ItemNumber',
        'A/C': 'AC',
        'Due-Dt': 'DueDt',
        'Create-Dt': 'CreateDt',
        'Wo-Type': 'WoType',
        'Plan Type': 'PlanType',
        'Opn-Q': 'OpnQ',
        'QA Aprvl': 'QAAprvl',
        'Prt-No': 'PrtNo',
        'Prt-User': 'PrtUser',
        'Prt-date': 'Prtdate'
    },
    "columns_order_original": [
        'Entity', 'Project No', 'WO No', 'SO/FCST', 'Sub', 'Parent WO', 'Item-Number', 'Rev', 'Description',
        'A/C', 'Due-Dt', 'Create-Dt', 'Wo-Type', 'Srt', 'Plan Type', 'Opn-Q', 'St', 'QA Aprvl', 'Prt-No',
        'Prt-User', 'Prt-date'
    ],
    "columns_order_renamed": [
        'Entity', 'ProjectNo', 'WONo', 'SO_FCST', 'Sub', 'ParentWO', 'ItemNumber', 'Rev', 'Description',
        'AC', 'DueDt', 'CreateDt', 'WoType', 'Srt', 'PlanType', 'OpnQ', 'St', 'QAAprvl', 'PrtNo',
        'PrtUser', 'Prtdate'
    ],
    "create_table_sql": """CREATE TABLE IF NOT EXISTS WOInquiry(
        id INTEGER PRIMARY KEY,
        Entity TEXT,
        ProjectNo TEXT,
        WONo TEXT,
        SO_FCST TEXT,
        Sub TEXT,
        ParentWO TEXT,
        ItemNumber TEXT,
        Rev TEXT,
        Description TEXT,
        AC TEXT,
        DueDt TEXT,
        CreateDt TEXT,
        WoType TEXT,
        Srt TEXT,
        PlanType TEXT,
        OpnQ TEXT,
        St TEXT,
        QAAprvl TEXT,
        PrtNo TEXT,
        PrtUser TEXT,
        Prtdate TEXT
    )"""
},
# ==================== TABLA WHERE_USE_TABLE_PLANTYPE ====================
"where_use_table_plantype": {
    "source_file": r"J:\Departments\Material Control\Purchasing\Tools\ComprasDB\where_use_plantype_table.xlsx",
    "table_name": "where_use_table_plantype",
    "file_type": "excel",
    "columns_order_original": [
        'PlanTp', 'Area'
    ],
    "columns_order_renamed": [
        'PlanTp', 'Area'
    ],
    "create_table_sql": """CREATE TABLE IF NOT EXISTS where_use_table_plantype(
        PlanTp TEXT,
        Area TEXT
    )"""
},
# ==================== TABLA SAFETY_STOCK ====================
"safety_stock": {
    "source_file": r"J:\Departments\Material Control\Purchasing\Tools\ComprasDB\safety_stock.txt",
    "table_name": "safety_stock",
    "file_type": "fixed_width",
    "fixed_width_params": {
        "widths": [9, 32, 30, 4, 13, 13, 13, 6, 9, 100],
        "header": 3
    },
    "data_filters": {
        "remove_index_0_if_dash_in": "Item-No"
    },
    "columns_mapping": {
        "Item-No": "ItemNo",
        "Safety-Stock": "Qty"
    },
    "columns_order_original": [
        "Item-No", "Safety-Stock"
    ],
    "columns_order_renamed": [
        "ItemNo", "Qty"
    ],
    "create_table_sql": """
        CREATE TABLE IF NOT EXISTS safety_stock (
            ItemNo TEXT,
            Qty TEXT
        )
    """
},
# ==================== TABLA SAFETY_STOCK_SHOULD_BE ====================
"safety_stock_should_be": {
    "source_file":  r"J:\Departments\Material Control\Purchasing\Tools\ComprasDB\safety_stock_should_be.xlsx",
    "table_name": "safety_stock_should_be",
    "file_type": "excel",
    "columns_order_original": [
        'ItemNo', 'New_Qty'
    ],
    "columns_order_renamed": [
        'ItemNo', 'New_Qty'
    ],
    "create_table_sql": """CREATE TABLE IF NOT EXISTS safety_stock_should_be(
        ItemNo TEXT,
        New_Qty TEXT
    )"""
},
# ==================== TABLA OH_PR_TABLE ====================
"oh_pr_table": {
    "source_file": r"J:\Departments\Material Control\Purchasing\Tools\ComprasDB\ohNeteable.xlsx",
    "table_name": "oh_pr_table",
    "file_type": "excel",
    "excel_params": {
        "skiprows": 4
    },
    "columns_mapping": {
        'Item-Number': 'ItemNo',
        'Manuf-Date': 'Manuf_Date',
        'Expire-Date': 'Expire_Date'
    },
    "uppercase_columns": [
        'Bin'
    ],
    "special_processing": {
        "fix_expire_date_year": True,
        "year_threshold": 1950,
        "year_adjustment": 100
    },
    "columns_order_original": [
        'Item-Number', 'Description', 'UM', 'Lot', 'Manuf-Date', 'Expire-Date', 'OH', 'Bin'
    ],
    "columns_order_renamed": [
        'ItemNo', 'Description', 'UM', 'Lot', 'Manuf_Date', 'Expire_Date', 'OH', 'Bin'
    ],
    "create_table_sql": """CREATE TABLE IF NOT EXISTS oh_pr_table(
        id INTEGER PRIMARY KEY,
        ItemNo TEXT,
        Description TEXT,
        UM TEXT,
        Lot TEXT,
        Manuf_Date TEXT,
        Expire_Date TEXT,
        OH TEXT,
        Bin TEXT
    )"""
},
# ==================== TABLA FCST ====================
"fcst": {
    "source_file": r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\WHS PLAN\FILES\MP 5 9\MP59.txt",
    "table_name": "fcst",
    "file_type": "fixed_width",
    "fixed_width_params": {
        "widths": [9, 5, 9, 30, 7, 24, 17, 4, 3, 8, 9, 10, 10, 10, 3, 2, 30],
        "header": 3,
        "skip_rows": [0],
        "drop_last_row": True
    },
    "columns_mapping": {
        'A/C-No': 'AC',
        'Config-ID            CRev   Fc': 'ConfigID',
        'st  Ite': 'FcstNo',
        'm-No                 Des': 'ItemNo',
        'UM': 'Rev',
        'Pla': 'UM',
        'cription      Rev': 'Description',
        'nned-By': 'PlannedBy',
        'Rqmt-Date': 'ReqDate',
        'Qty-Fcst': 'QtyFcst',
        'Opn-Qty': 'OpenQty',
        'WO          SO-AC': 'WO'
    },
    "columns_order_original": [
        'Entity', 'Proj', 'A/C-No', 'Config-ID            CRev   Fc', 'st  Ite', 'm-No                 Des',
        'UM', 'Pla', 'cription      Rev', 'nned-By', 'Rqmt-Date', 'Qty-Fcst', 'Opn-Qty', 'WO          SO-AC'
    ],
    "columns_order_renamed": [
        'Entity', 'Proj', 'AC', 'ConfigID', 'FcstNo', 'Description', 'ItemNo',
        'Rev', 'UM', 'PlannedBy', 'ReqDate', 'QtyFcst', 'OpenQty', 'WO'
    ],
    "special_processing": {
        "clear_before_insert": True,
        "has_control_table": True,
        "custom_cleaning": True,
        "numeric_columns": ["QtyFcst", "OpenQty"],
        "final_columns_only": True
    },
    "create_table_sql": """CREATE TABLE IF NOT EXISTS fcst(
        Entity TEXT,
        Proj TEXT,
        AC TEXT,
        ConfigID TEXT,
        FcstNo TEXT,
        Description TEXT,
        ItemNo TEXT,
        Rev TEXT,
        UM TEXT,
        PlannedBy TEXT,
        ReqDate TEXT,
        QtyFcst INTEGER,
        OpenQty INTEGER,
        WO TEXT
    )""",
    "control_table": {
        "table_name": "fcst_loaded",
        "create_table_sql": """CREATE TABLE IF NOT EXISTS fcst_loaded(
            last_loaded_date TEXT,
            file_name TEXT
        )"""
    }
},
# ==================== TABLA PR561 ====================
"pr561": {
    "source_file": r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\WHS PLAN\FILES\PR 5 61\pr561.txt",
    "table_name": "pr561",
    "file_type": "fixed_width",
    "fixed_width_params": {
        "widths": [2, 10, 7, 32, 14, 30, 9, 3, 16, 13, 12, 14, 12, 12, 14, 9],
        "header": 3,
        "skip_rows": [0]
    },
    "data_filters": {
        "exclude_rows_multiple": {
            "column": "In",
            "values": ["**", "", "TO", "En"]
        }
    },
    "columns_mapping": {
        'Fuse-No': 'FuseNo',
        'Component Description': 'Description',
        'Qty-Oh': 'QtyOh',
        'Qty-Issue': 'QtyIssue',
        'Qty-no-Iss': 'QtyPending',
        'Qty-Required': 'ReqQty',
        'Wo-No': 'WoNo'
    },
    "columns_order_original": [
        'In', 'Entity', 'Project', 'Component', 'Fuse-No', 'Component Description',
        'PlnType', 'Srt', 'Qty-Oh', 'Qty-Issue', 'Qty-no-Iss', 'Qty-Required', 'Wo-No'
    ],
    "columns_order_renamed": [
        'Entity', 'Project', 'Component', 'FuseNo', 'Description',
        'PlnType', 'Srt', 'QtyOh', 'QtyIssue', 'QtyPending', 'ReqQty', 'WoNo'
    ],
    "special_processing": {
        "clear_before_insert": True,
        "custom_cleaning": True,
        "final_columns_only": True
    },
    "create_table_sql": """CREATE TABLE IF NOT EXISTS pr561(
        id INTEGER PRIMARY KEY,
        Entity TEXT,
        Project TEXT,
        Component TEXT,
        FuseNo TEXT,
        Description TEXT,
        PlnType TEXT,
        Srt TEXT,
        QtyOh TEXT,
        QtyIssue TEXT,
        QtyPending TEXT,
        ReqQty TEXT,
        WoNo TEXT
    )"""
},
# ==================== TABLA Credit Memos ====================
"Credit_Memos": {
    "source_file": r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\WHS PLAN\FILES\Credit Memos\Credit_Memos.xlsx",
    "table_name": "Credit_Memos",
    "file_type": "excel",
    "excel_params": {
        "skiprows": 3
    },
    "columns_mapping": {
    'Entity-Code': 'Entity_Code',
    'Invoice-No': 'Invoice_No',
    'Line': 'Line',
    'SO-No': 'SO_No',
    'Customer': 'Customer',
    'Type': 'Type',
    'Cust Name': 'Cust_Name',
    'PO-No': 'PO_No',
    'A/C': 'AC',
    'Spr-C': 'Spr_C',
    'O-Cd': 'O_Cd',
    'Proj': 'Proj',
    'TC': 'TC',
    'Item Number': 'Item_Number',
    'Description': 'Description',
    'UM': 'UM',
    'S-Qty': 'S_Qty',
    'Price': 'Price',
    'Total': 'Total',
    'Curr': 'Curr',
    'CustReqDt': 'CustReqDt',
    'Req-Date': 'Req_Date',
    'Pln-Ship': 'Pln_Ship',
    'Inv-Dt': 'Inv_Dt',
    'Inv-Time': 'Inv_Time',
    'Ord-Dt': 'Ord_Dt',
    'SODueDt': 'SODueDt',
    'Due-Dt': 'Due_Dt',
    'TimeToConfirm(hr)': 'TimeToConfirm_hr',
    'Plan-FillRate': 'Plan_FillRate',
    'TAT to fill an order': 'TAT_to_fill_an_order',
    'Cust-FillRate': 'Cust_FillRate',
    'SH': 'SH',
    'DG': 'DG',
    'Std Cost': 'Std_Cost',
    'Vendor-C': 'Vendor_C',
    'Stk': 'Stk',
    'Std LT': 'Std_LT',
    'Shipped-Dt': 'Shipped_Dt',
    'ShipTo': 'ShipTo',
    'ViaDesc': 'ViaDesc',
    'Tracking No': 'Tracking_No',
    'AddntlTracking': 'AddntlTracking',
    'CreditedInv': 'CreditedInv',
    'Invoice Line Memo': 'Invoice_Line_Memo',
    'Lot No (Qty)': 'Lot_No_Qty',
    'Manuf Charge': 'Manuf_Charge',
    'Inv-Credit': 'Inv_Credit',
    'CM-Reason': 'CM_Reason',
    'User-Id': 'User_Id',
    'Cust-Item': 'Cust_Item',
    'Buyer-Name': 'Buyer_Name',
    'Issue-Date': 'Issue_Date',
    'Memo[1]': 'Memo_1',
    'Memo[2]': 'Memo_2'
},
    "columns_order_original": [
        'Entity-Code', 'Invoice-No', 'Line', 'SO-No', 'Customer', 'Type',
        'Cust Name', 'PO-No', 'A/C', 'Spr-C', 'O-Cd', 'Proj', 'TC',
        'Item Number', 'Description', 'UM', 'S-Qty', 'Price', 'Total', 'Curr',
        'CustReqDt', 'Req-Date', 'Pln-Ship', 'Inv-Dt', 'Inv-Time', 'Ord-Dt',
        'SODueDt', 'Due-Dt', 'TimeToConfirm(hr)', 'Plan-FillRate',
        'TAT to fill an order', 'Cust-FillRate', 'SH', 'DG', 'Std Cost',
        'Vendor-C', 'Stk', 'Std LT', 'Shipped-Dt', 'ShipTo', 'ViaDesc',
        'Tracking No', 'AddntlTracking', 'CreditedInv', 'Invoice Line Memo',
        'Lot No (Qty)', 'Manuf Charge', 'Inv-Credit', 'CM-Reason', 'User-Id',
        'Cust-Item', 'Buyer-Name', 'Issue-Date', 'Memo[1]', 'Memo[2]'
    ],
    "columns_order_renamed": [
        'Entity_Code', 'Invoice_No', 'Line', 'SO_No', 'Customer'
        , 'Type', 'Cust_Name', 'PO_No', 'A_C', 'Spr_C', 'O_Cd'
        , 'Proj', 'TC', 'Item_Number', 'Description', 'UM'
        , 'S_Qty', 'Price', 'Total', 'Curr', 'CustReqDt'
        , 'Req_Date', 'Pln_Ship', 'Inv_Dt', 'Inv_Time'
        , 'Ord_Dt', 'SODueDt', 'Due_Dt'
        , 'TimeToConfirmhr', 'Plan_FillRate', 'TAT_to_fill_an_order', 'Cust_FillRate', 'SH'
        , 'DG', 'Std_Cost', 'Vendor_C', 'Stk', 'Std_LT', 'Shipped_Dt', 'ShipTo', 'ViaDesc'
        , 'Tracking_No', 'AddntlTracking', 'CreditedInv', 'Invoice_Line_Memo', 'Lot_No_Qty'
        , 'Manuf_Charge', 'Inv_Credit', 'CM_Reason', 'User_Id', 'Cust_Item', 'Buyer_Name'
        , 'Issue_Date', 'Memo1', 'Memo2'
    ],
    "create_table_sql": """CREATE TABLE IF NOT EXISTS Credit_Memos(
        Entity_Code TEXT,
        Invoice_No TEXT,
        Line TEXT,
        SO_No TEXT,
        Customer TEXT,
        Type TEXT,
        Cust_Name TEXT,
        PO_No TEXT,
        AC TEXT,
        Spr_C TEXT,
        O_Cd TEXT,
        Proj TEXT,
        TC TEXT,
        Item_Number TEXT,
        Description TEXT,
        UM TEXT,
        S_Qty REAL,
        Price REAL,
        Total REAL,
        Curr TEXT,
        CustReqDt TEXT,
        Req_Date TEXT,
        Pln_Ship TEXT,
        Inv_Dt TEXT,
        Inv_Time TEXT,
        Ord_Dt TEXT,
        SODueDt TEXT,
        Due_Dt TEXT,
        TimeToConfirm_hr REAL,
        Plan_FillRate REAL,
        TAT_to_fill_an_order REAL,
        Cust_FillRate REAL,
        SH TEXT,
        DG TEXT,
        Std_Cost REAL,
        Vendor_C TEXT,
        Stk TEXT,
        Std_LT REAL,
        Shipped_Dt TEXT,
        ShipTo TEXT,
        ViaDesc TEXT,
        Tracking_No TEXT,
        AddntlTracking TEXT,
        CreditedInv TEXT,
        Invoice_Line_Memo TEXT,
        Lot_No_Qty TEXT,
        Manuf_Charge TEXT,
        Inv_Credit TEXT,
        CM_Reason TEXT,
        User_Id TEXT,
        Cust_Item TEXT,
        Buyer_Name TEXT,
        Issue_Date TEXT,
        Memo_1 TEXT,
        Memo_2 TEXT
    )"""
    ,
    "control_table": {
    "table_name": "Credit_Memos",
    "create_table_sql": """CREATE TABLE IF NOT EXISTS Credit_Memos(
        last_loaded_date TEXT,
        file_name TEXT
    )"""
}
},
# ==================== TABLA OpenWO_with_Closed_SO ====================
"OpenWO_with_Closed_SO": {
    "source_file": r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\WHS PLAN\FILES\PR 5 58\pr558.txt",
    "table_name": "OpenWO_with_Closed_SO",
    "file_type": "fixed_width",
    "fixed_width_params": {
        "widths": [10,2,10,32,31,9,5,10,20],
        "header": 3,
        "skip_rows": [0]
    },
    "data_filters": {
        "exclude_rows_multiple": {
            "column": "WO-No   Su",
            "values": ["**", "End-of-Rep"]
        }
    },
    "columns_mapping": {
        'WO-No   Su': 'WONo',
        'b': 'Sub',
        'Req-Date': 'DueDate_WO',
        'Order-No': 'OrderNo',
        'Item-No': 'ItemNo',
        'Req-Date.1': 'ReqDate1'  # Mantener ReqDate1
    },
    "columns_order_original": [
        'WO-No   Su', 'b', 'Req-Date', 'Item-No', 'Description', 'Order-No', 'Ln', 'Req-Date.1', 'Entity'
    ],
    "columns_order_renamed": [
        'WONo', 'Sub', 'DueDate_WO', 'ItemNo', 'Description', 'OrderNo', 'Ln', 'ReqDate1', 'Entity'  # Incluir ReqDate1
    ],
    "special_processing": {
        "clear_before_insert": True,
        "custom_cleaning": True,
        "final_columns_only": True
    },
    "create_table_sql": """CREATE TABLE IF NOT EXISTS OpenWO_with_Closed_SO(
        WONo TEXT,
        Sub TEXT,
        DueDate_WO TEXT,
        ItemNo TEXT,
        Description TEXT,
        OrderNo TEXT,
        Ln TEXT,
        ReqDate1 TEXT,  -- Cambiar a ReqDate1
        Entity TEXT
    )"""
},
# ==================== TABLA IN 5 6====================
"prepetual_in56": {
    "source_file": r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\WHS PLAN\FILES\IN 5 6\HISTORICO.txt",
    "table_name": "prepetual_in56",
    "file_type": "fixed_width",
    "fixed_width_params": {
        "widths": [9,7,5,32,8,7,10,16,16,9,9,18,8,13,28,17,7,15,12,15,10,9,7,18,10],
        "header": 0,
        "skip_rows": [0],
        "drop_last_row": True
    },
    "columns_mapping": {
        'Date': 'Date',
        'Reason': 'Reason',
        'Acct': 'Action',
        'Item Number': 'ItemNo',
        'Product': 'PRD',
        'Group': 'Group',
        'A/C': 'AC',
        'Lot No': 'LotNo',
        'Fam-Code': 'FamCode',
        'PlanTp': 'PlanType',
        'Ent': 'Entity',
        'Qty': 'Qty',
        'UM': 'UM',
        'Cur Std': 'CurStd',
        'DR': 'DR',
        'CR': 'CR',
        'NET': 'NET',
        'Reference': 'Reference',
        'To-Entity': 'ToEntity',
        'Seq No': 'SeqNo',
        'Userid': 'Userid',
        'BinLoc': 'BinLoc',
        'Zone': 'Zone',
        'Reference2': 'Reference2',
        'DateTime': 'DateTime'
    },


    "columns_order_original": [
        'Date', 'Reason', 'Acct', 'Item Number', 'Product', 'Group', 'A/C',
        'Lot No', 'Fam-Code', 'PlanTp', 'Ent', 'Qty', 'UM', 'Cur Std', 'DR',
        'CR', 'NET', 'Reference', 'To-Entity', 'Seq No', 'Userid', 'BinLoc',
        'Zone', 'Reference2', 'DateTime'
    ],
    "columns_order_renamed": [
        'Date', 'Reason', 'Action', 'ItemNo', 'PRD', 'Group', 'AC',
        'LotNo', 'FamCode', 'PlanType', 'Entity', 'Qty', 'UM', 'CurStd', 'DR',
        'CR', 'NET', 'Reference', 'ToEntity', 'SeqNo', 'Userid', 'BinLoc',
        'Zone', 'Reference2', 'DateTime'
    ],
    "special_processing": {
        "clear_before_insert": True,
        "has_control_table": True,
        "custom_cleaning": True,
        "numeric_columns": ["Qty"],
        "final_columns_only": False
    },
    "create_table_sql": """
    CREATE TABLE IF NOT EXISTS prepetual_in56 (
        Date TEXT,
        Reason TEXT,
        Action TEXT,
        ItemNo TEXT,
        PRD TEXT,
        [Group] TEXT,
        AC TEXT,
        LotNo TEXT,
        FamCode TEXT,
        PlanType TEXT,
        Entity TEXT,
        Qty INTEGER,
        UM TEXT,
        CurStd REAL,
        DR REAL,
        CR REAL,
        NET REAL,
        Reference TEXT,
        ToEntity TEXT,
        SeqNo TEXT,
        Userid TEXT,
        BinLoc TEXT,
        Zone TEXT,
        Reference2 TEXT,
        DateTime TEXT
    )""",
    "control_table": {
        "table_name": "prepetual_in56",
        "create_table_sql": """CREATE TABLE IF NOT EXISTS prepetual_in56(
            last_loaded_date TEXT,
            file_name TEXT
        )"""
    }

}

}


# Rutas globales usando raw strings para evitar problemas con backslashes
BASE_PATHS = {
    "db_folder": r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\R4Database",
    "source_folder": r"J:\Departments\Material Control\Purchasing\Tools\ComprasDB",
    "tracking_file": r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\R4Database\file_tracking.json"
}

def get_table_config(table_name):
    """Obtiene la configuración de una tabla específica"""
    return TABLES_CONFIG.get(table_name)

def get_all_tables():
    """Obtiene lista de todas las tablas configuradas"""
    return list(TABLES_CONFIG.keys())

def add_new_table_config(table_name, source_file, columns_mapping, columns_order_original, columns_order_renamed, create_table_sql):
    """Función helper para agregar nueva tabla dinámicamente"""
    TABLES_CONFIG[table_name] = {
        "source_file": source_file,
        "table_name": table_name,
        "columns_mapping": columns_mapping,
        "columns_order_original": columns_order_original,
        "columns_order_renamed": columns_order_renamed,
        "create_table_sql": create_table_sql
    }
    print(f"Configuración agregada para tabla: {table_name}")

def read_file_data(file_path, table_config):
    """Lee datos del archivo según su tipo (CSV, TXT delimitado, ancho fijo, o Excel)"""
    import pandas as pd
    
    file_type = table_config.get("file_type", "csv")
    table_name = table_config.get("table_name", "")
    
    # Procesamiento especial para FCST
    if table_name == "fcst" and table_config.get("special_processing", {}).get("custom_cleaning", False):
        print(f"📖 Procesamiento especial para tabla FCST...")
        
        # Leer el archivo con anchos fijos
        fixed_params = table_config.get("fixed_width_params", {})
        widths = fixed_params.get("widths", [])
        header = fixed_params.get("header", 0)
        
        df = pd.read_fwf(file_path, widths=widths, header=header)
        
        # Limpiar datos
        df = df.fillna('')
        df.drop([0], axis=0, inplace=True)
        df = df.iloc[:-1]  # Eliminar última fila
        
        # Renombrar columnas según el mapeo específico
        columns_mapping = table_config.get("columns_mapping", {})
        df = df.rename(columns=columns_mapping)
        
        # Seleccionar solo las columnas necesarias
        final_columns = table_config.get("columns_order_renamed", [])
        available_columns = [col for col in final_columns if col in df.columns]
        df = df[available_columns]
        
        # Convertir columnas numéricas
        numeric_columns = table_config.get("special_processing", {}).get("numeric_columns", [])
        for col in numeric_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
                print(f"   🔢 Convertido a numérico: {col}")
        
        print(f"   ✅ Datos FCST procesados: {len(df)} filas, {len(df.columns)} columnas")
        return df
    
    
    # Procesamiento especial para PR561
    if table_name == "pr561" and table_config.get("special_processing", {}).get("custom_cleaning", False):
        print(f"📖 Procesamiento especial para tabla PR561...")
        
        # Leer el archivo con anchos fijos
        fixed_params = table_config.get("fixed_width_params", {})
        widths = fixed_params.get("widths", [])
        header = fixed_params.get("header", 0)
        
        df = pd.read_fwf(file_path, widths=widths, header=header)
        
        # Eliminar primera fila de datos
        df.drop([0], axis=0, inplace=True)
        print(f"   🗑️ Eliminada primera fila de datos")
        
        # Aplicar filtros individuales para la columna 'In' (como en tu código original)
        initial_count = len(df)
        df = df.loc[(df['In'] != '**')]
        df = df.loc[(df['In'] != '')]
        df = df.loc[(df['In'] != 'TO')]
        df = df.loc[(df['In'] != 'En')]
        filtered_count = initial_count - len(df)
        print(f"   🔽 Filtradas {filtered_count} filas por valores excluidos en 'In'")
        
        # Rellenar valores nulos
        df = df.fillna('')
        print(f"   🔧 Valores nulos rellenados")
        
        # Renombrar columnas
        columns_mapping = table_config.get("columns_mapping", {})
        df = df.rename(columns=columns_mapping)
        print(f"   🔄 Columnas renombradas")
        
        # Seleccionar solo las columnas finales (sin 'In')
        final_columns = table_config.get("columns_order_renamed", [])
        available_columns = [col for col in final_columns if col in df.columns]
        df = df[available_columns]
        print(f"   📋 Seleccionadas {len(available_columns)} columnas finales")
        
        # Aplicar strip a todas las columnas de texto
        for col in df.columns:
            if df[col].dtype == 'object':
                df[col] = df[col].str.strip()
        print(f"   🧹 Strip aplicado a columnas de texto")
        
        print(f"   ✅ Datos PR561 procesados: {len(df)} filas, {len(df.columns)} columnas")
        return df
    
    # Procesamiento normal para otras tablas
    if file_type == "txt_delimited":
        # Leer archivo .txt con delimitador especial
        delimiter = table_config.get("delimiter", "@")
        
        print(f"📖 Leyendo archivo TXT con delimitador '{delimiter}'...")
        
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()
        except UnicodeDecodeError:
            # Intentar con encoding diferente
            with open(file_path, 'r', encoding='latin-1') as f:
                lines = f.readlines()
        
        if not lines:
            raise Exception("Archivo vacío")
        
        # Primera línea son los headers
        headers = lines[0].strip().split(delimiter)
        print(f"   📋 Headers encontrados: {headers}")
        
        # Resto son datos
        data = []
        for i, line in enumerate(lines[1:], 2):
            if line.strip():  # Ignorar líneas vacías
                row_data = line.strip().split(delimiter)
                if len(row_data) == len(headers):
                    data.append(row_data)
                else:
                    print(f"   ⚠️ Línea {i} tiene {len(row_data)} columnas, esperaba {len(headers)} - IGNORADA")
        
        df = pd.DataFrame(data, columns=headers)
        print(f"   ✅ Datos cargados: {len(df)} filas, {len(df.columns)} columnas")
        
        return df
    
    elif file_type == "fixed_width":
        # Leer archivo de ancho fijo
        fixed_params = table_config.get("fixed_width_params", {})
        widths = fixed_params.get("widths", [])
        header = fixed_params.get("header", 0)
        skip_rows = fixed_params.get("skip_rows", [])
        
        print(f"📖 Leyendo archivo de ancho fijo...")
        print(f"   📏 Anchos: {widths}")
        print(f"   📋 Header en línea: {header}")
        
        # Leer con pandas read_fwf
        df = pd.read_fwf(file_path, widths=widths, header=header)
        
        # Eliminar filas específicas si se especifica
        if skip_rows:
            print(f"   🗑️ Eliminando filas: {skip_rows}")
            df.drop(skip_rows, axis=0, inplace=True, errors='ignore')
        
        print(f"   ✅ Datos cargados: {len(df)} filas, {len(df.columns)} columnas")
        
        return df
    
    elif file_type == "excel":
        # Leer archivo Excel
        excel_params = table_config.get("excel_params", {})
        skiprows = excel_params.get("skiprows", 0)
        
        print(f"📖 Leyendo archivo Excel...")
        print(f"   ⏭️ Saltando filas: {skiprows}")
        
        try:
            # Intentar leer con diferentes métodos para manejar problemas de codificación
            df = pd.read_excel(file_path, skiprows=skiprows, dtype=str, engine='openpyxl')
        except Exception as e1:
            try:
                # Intentar con motor xlrd
                df = pd.read_excel(file_path, skiprows=skiprows, dtype=str, engine='xlrd')
            except Exception as e2:
                try:
                    # Intentar sin dtype específico
                    df = pd.read_excel(file_path, skiprows=skiprows, engine='openpyxl')
                    # Convertir todo a string después
                    df = df.astype(str)
                except Exception as e3:
                    raise Exception(f"Error leyendo Excel: {e1}, {e2}, {e3}")
        
        print(f"   ✅ Datos cargados: {len(df)} filas, {len(df.columns)} columnas")
        
        return df
    
    else:
        # Archivo CSV estándar
        print(f"📖 Leyendo archivo CSV...")
        try:
            return pd.read_csv(file_path, encoding='utf-8')
        except UnicodeDecodeError:
            return pd.read_csv(file_path, encoding='latin-1')

def apply_data_filters(df, table_config):
    """Aplica filtros de datos según configuración"""
    data_filters = table_config.get("data_filters", {})
    
    # Filtro para excluir filas
    exclude_rows = data_filters.get("exclude_rows", {})
    if exclude_rows:
        for column, value in exclude_rows.items():
            if column in df.columns:
                initial_count = len(df)
                df = df.loc[df[column] != value]
                filtered_count = initial_count - len(df)
                print(f"   🔽 Filtro aplicado: Excluidas {filtered_count} filas donde {column} = '{value}'")
    
    return df

def apply_special_processing(df, table_config):
    """Aplica procesamiento especial según configuración"""
    special_processing = table_config.get("special_processing", {})
    
    # Procesamiento de fechas de expiración para reworkloc_all
    if special_processing.get("fill_expire_date", False):
        expire_col = None
        for col in df.columns:
            if 'expire' in col.lower() or 'expir' in col.lower():
                expire_col = col
                break
        
        if expire_col:
            print(f"   📅 Procesando fechas de expiración en columna: {expire_col}")
            import pandas as pd
            from datetime import datetime
            
            current_date = datetime.today().date()
            years_to_add = special_processing.get("expire_years", 2)
            future_date = pd.Timestamp(current_date) + pd.DateOffset(years=years_to_add)
            
            df[expire_col].fillna(future_date.normalize(), inplace=True)
            print(f"   ✅ Fechas nulas rellenadas con: {future_date.date()}")
    
    return df

def clean_data_remove_quotes(df):
    """Remueve comillas simples de todas las columnas"""
    print("   🧹 Removiendo comillas simples...")
    for i, col in enumerate(df.columns):
        if df[col].dtype == 'object':
            df.iloc[:, i] = df.iloc[:, i].astype(str).str.replace("'", '', regex=False)
    return df

def apply_data_transformations(df, table_config):
    """Aplica transformaciones especiales de datos (uppercase, strip, etc.)"""
    # Aplicar strip a todas las columnas de texto
    print("   🧹 Aplicando strip a columnas de texto...")
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].astype(str).str.strip()
    
    # Remover comillas simples
    df = clean_data_remove_quotes(df)
    
    # Aplicar uppercase a columnas específicas
    uppercase_columns = table_config.get("uppercase_columns", [])
    for col in uppercase_columns:
        if col in df.columns:
            df[col] = df[col].astype(str).str.upper()
            print(f"   🔤 Aplicado uppercase a columna: {col}")
    
    # Aplicar procesamiento especial
    df = apply_special_processing(df, table_config)
    
    return df

def validate_vendor_columns(df, table_config):
    """Validación especial para tabla vendor"""
    expected_columns = table_config.get("expected_columns", [])
    if expected_columns:
        actual_columns = df.columns.tolist()
        print(f"   📋 Validando columnas para tabla vendor...")
        print(f"   📋 Esperadas: {expected_columns}")
        print(f"   📋 Encontradas: {actual_columns}")
        
        if actual_columns != expected_columns:
            print(f"   ⚠️ Columnas no coinciden - usando DataFrame tal como está")
            return False
        else:
            print(f"   ✅ Columnas válidas para vendor")
            return True
    return True

def detect_column_format(file_path, table_config):
    """Detecta automáticamente si las columnas están en formato original o ya renombradas"""
    try:
        # Leer archivo según su tipo
        df_sample = read_file_data(file_path, table_config)
        if df_sample.empty:
            return 'unknown'
        
        # Validación especial para vendor
        table_name = table_config.get("table_name", "")
        if table_name.lower() == "vendor":
            if not validate_vendor_columns(df_sample, table_config):
                return 'unknown'
        
        # Aplicar filtros si existen
        df_sample = apply_data_filters(df_sample, table_config)
        
        # Solo tomar los headers para análisis
        csv_columns = set(df_sample.columns.tolist())
        
        # Verificar si tiene columnas originales
        original_columns = set(table_config.get("columns_order_original", []))
        renamed_columns = set(table_config.get("columns_order_renamed", []))
        
        # Si no hay configuración de columnas, asumir que está bien
        if not original_columns and not renamed_columns:
            print(f"   ℹ️ Sin configuración de columnas específica - usando tal como está")
            return 'renamed'
        
        # Calcular coincidencias
        original_matches = len(csv_columns.intersection(original_columns))
        renamed_matches = len(csv_columns.intersection(renamed_columns))
        
        print(f"🔍 Análisis de columnas en {os.path.basename(file_path)}:")
        print(f"   📊 Columnas encontradas: {len(csv_columns)}")
        print(f"   🔤 Coincidencias formato original: {original_matches}/{len(original_columns)}")
        print(f"   🔄 Coincidencias formato renombrado: {renamed_matches}/{len(renamed_columns)}")
        
        # Determinar formato
        if original_matches > renamed_matches and original_matches >= len(original_columns) * 0.8:
            print("   ✅ Detectado: Formato ORIGINAL (necesita renombrado)")
            return 'original'
        elif renamed_matches > original_matches and renamed_matches >= len(renamed_columns) * 0.8:
            print("   ✅ Detectado: Formato RENOMBRADO (listo para usar)")
            return 'renamed'
        else:
            print("   ⚠️ Detectado: Formato DESCONOCIDO")
            print(f"   📋 Columnas en archivo: {sorted(list(csv_columns))}")
            return 'unknown'
            
    except Exception as e:
        print(f"❌ Error detectando formato de columnas: {e}")
        return 'unknown'

def process_dataframe_columns(df, table_config, detected_format):
    """Procesa el DataFrame según el formato detectado"""
    
    # Aplicar filtros de datos primero
    df = apply_data_filters(df, table_config)
    
    if detected_format == 'original':
        print("🔄 Aplicando mapeo de columnas...")
        
        # Seleccionar columnas en orden original
        columns_order = table_config.get("columns_order_original", [])
        available_columns = [col for col in columns_order if col in df.columns]
        
        if len(available_columns) != len(columns_order):
            missing = set(columns_order) - set(available_columns)
            print(f"⚠️ Columnas faltantes: {missing}")
        
        # Seleccionar y renombrar
        df_processed = df[available_columns].copy()
        columns_mapping = table_config.get("columns_mapping", {})
        df_processed = df_processed.rename(columns=columns_mapping)
        
        print(f"✅ Columnas procesadas: {len(df_processed.columns)}")
        
    elif detected_format == 'renamed':
        print("✅ Columnas ya están renombradas, usando directamente...")
        
        # Seleccionar columnas en orden renombrado
        columns_order = table_config.get("columns_order_renamed", [])
        available_columns = [col for col in columns_order if col in df.columns]
        
        if len(available_columns) != len(columns_order):
            missing = set(columns_order) - set(available_columns)
            print(f"⚠️ Columnas faltantes: {missing}")
        
        df_processed = df[available_columns].copy()
        print(f"✅ Columnas utilizadas: {len(df_processed.columns)}")
        
    else:
        print("❌ Formato de columnas desconocido, usando DataFrame original...")
        df_processed = df
    
    # Aplicar transformaciones de datos (uppercase, strip, etc.)
    df_processed = apply_data_transformations(df_processed, table_config)
    
    return df_processed

def verify_paths():
    """Verifica que las rutas configuradas existan"""
    print("🔍 Verificando rutas configuradas...")
    
    # Verificar rutas base
    for path_name, path_value in BASE_PATHS.items():
        if os.path.exists(path_value):
            print(f"✅ {path_name}: {path_value}")
        else:
            print(f"❌ {path_name}: {path_value} - NO EXISTE")
    
    # Verificar archivos fuente para todas las tablas
    print(f"🔍 Verificando archivos fuente para {len(TABLES_CONFIG)} tablas...")
    for table_name, config in TABLES_CONFIG.items():
        source_file = config.get("source_file", "")
        if source_file != "dynamic_file_selection" and os.path.exists(source_file):
            size = os.path.getsize(source_file)
            file_type = config.get("file_type", "csv").upper()
            print(f"✅ {table_name}: {os.path.basename(source_file)} ({file_type}, {size:,} bytes)")
        elif source_file == "dynamic_file_selection":
            print(f"⚙️ {table_name}: Archivo de selección dinámica")
        else:
            print(f"❌ {table_name}: {os.path.basename(source_file)} - NO EXISTE")

if __name__ == "__main__":
    print("=== Verificación de Configuración Multi-Tabla ===")
    print(f"📊 Tablas configuradas: {len(TABLES_CONFIG)}")
    for table_name in TABLES_CONFIG.keys():
        print(f"   - {table_name}")
    print()
    verify_paths()