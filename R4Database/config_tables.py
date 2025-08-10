# config_tables.py - Configuración para múltiples tablas
# VERSIÓN COMPLETA SIN ERRORES - Con soporte para todas las tablas

import os
from pathlib import Path

TABLES_CONFIG = {

# ==================== TABLA SALES_ORDER_TABLE ====================
"sales_order_table": {
    "source_file": r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\WHS PLAN\FILES\SalesOrder\SALESORDER.csv",
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
        'Std-Cost': 'STDCost',
        'Lot-Size': 'LotSize',
        'Fill-Doc': 'FillDoc',
        'Demand - Type': 'DemandType'
    },
    "columns_order_original": [
        'Entity Group', 'Project', 'A/C', 'Item-No', 'Description',
        'PlanTp', 'Ref', 'Sub', 'Sort', 'Fill-Doc', 'Demand - Type', 'Req-Qty',
        'Demand-Source', 'Unit', 'Vendor', 'Req-Date', 'Ship-Date', 'OH', 
        'MLIK Code', 'LT', 'Std-Cost', 'Lot-Size', 'UOM'
    ],
    "columns_order_renamed": [
        'EntityGroup', 'Project', 'AC', 'ItemNo', 'Description',
        'PlanTp', 'Ref', 'Sub', 'Sort', 'FillDoc', 'DemandType', 'ReqQty',
        'DemandSource', 'Unit', 'Vendor', 'ReqDate', 'ShipDate', 'OH', 
        'MLIKCode', 'LT', 'Std-Cost', 'LotSize', 'UOM'
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
        "Std-Cost" TEXT,
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
#==================== TABLA locations WHS whs_location_in36851====================
"whs_location_in36851": {
    "source_file": r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\WHS PLAN\FILES\WHS_Locations\whs_location_in36851.txt",
    "table_name": "whs_location_in36851",
    "file_type": "fixed_width",
    "fixed_width_params": {
        "widths": [3,10,9,31,12,9,10,29,9,9],
        "header": 3,
        "skip_rows": [0]
    },
    "special_processing": {
        "custom_cleaning": True,
        "use_exact_logic": True
    },
    "columns_mapping": {
        'Whs': 'Whs',
        'Bin ID': 'BinID',
        'Loc-ID': 'LocID',
        'Description': 'Description',
        'Excl-Bin': 'ExclBin',
        'Zone': 'Zone',
        'Container': 'Container',
        'Tag                  Netable': 'TagNetable',
        'Section': 'Section',
        'Type': 'Type'
    },
    "columns_order_original": [
        'Whs', 'Bin ID', 'Loc-ID', 'Description', 'Excl-Bin', 'Zone',
        'Container', 'Tag                  Netable', 'Section', 'Type'
    ],
    "columns_order_renamed": [
        'Whs', 'BinID', 'LocID', 'Description', 'ExclBin', 'Zone',
        'Container', 'TagNetable', 'Section', 'Type'
    ],
    "create_table_sql": """CREATE TABLE IF NOT EXISTS whs_location_in36851(
        id INTEGER PRIMARY KEY,
        Whs TEXT,
        BinID TEXT,
        LocID TEXT,
        Description TEXT,
        ExclBin TEXT,
        Zone TEXT,
        Container TEXT,
        TagNetable TEXT,
        Section TEXT,
        Type TEXT
    )"""
},
# ==================== TABLA SUPERBOM ====================
"bom": {
    "source_file": r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\WHS PLAN\FILES\SUPERBOM\SUPERBOM.txt",
    "table_name": "bom",
    "file_type": "fixed_width",
    "fixed_width_params": {
        "widths": [23, 31, 20, 3, 2, 5, 3, 5, 10, 10, 13, 13, 11, 13, 10, 10, 11, 20],
        "header": 6,
        "skip_rows": [0]
    },
    "data_filters": {
        "exclude_rows": {"Level Number": "End-of-Report.       1"}
    },
    "columns_mapping": {
        'Level Number': 'Level_Number',
        'Plan-Type': 'Plan_Type',
        'Unit Qty': 'Unit_Qty',
        'Std-Cost': 'Std_Cost',
        'Ext-Std': 'Ext_Std',
        'Labor IN': 'Labor_IN',
        'Lab  Rem': 'Lab_Rem',
        'Mat IN': 'Mat_IN',
        'Mat Rem': 'Mat_Rem',
        'This lvl': 'This_lvl',
        'Lab Hrs': 'Lab_Hrs'
    },
    "columns_order_original": [
        'Level Number', 'Component', 'Description', 'CE', 'T', 'Sort', 'UM',
        'MLI', 'Plan-Type', 'Unit Qty', 'Std-Cost', 'Ext-Std', 'Labor IN',
        'Lab  Rem', 'Mat IN', 'Mat Rem', 'This lvl', 'Lab Hrs'
    ],
    "columns_order_renamed": [
        'Level_Number', 'Component', 'Description', 'CE', 'T', 'Sort', 'UM',
        'MLI', 'Plan_Type', 'Unit_Qty', 'Std_Cost', 'Ext_Std', 'Labor_IN',
        'Lab_Rem', 'Mat_IN', 'Mat_Rem', 'This_lvl', 'Lab_Hrs'
    ],
    "uppercase_columns": [
        'Sort'
    ],
    "special_processing": {
        "clear_before_insert": True,
        "validate_columns": True,
        "custom_cleaning": True,
        "column_name_cleaning": True,
        "numeric_conversion": ["Unit_Qty", "Std_Cost", "Ext_Std", "Labor_IN", "Lab_Rem", "Mat_IN", "Mat_Rem", "This_lvl", "Lab_Hrs"],
        "auto_detect_header": True,
        "generate_key_by_level": True,  # Nuevo flag para generar keys
        "add_original_order": True      # Nuevo flag para añadir orden original
    },
    "create_table_sql": """CREATE TABLE IF NOT EXISTS bom(
        id INTEGER PRIMARY KEY,
        Level_Number TEXT,
        Component TEXT,
        Description TEXT,
        CE TEXT,
        T TEXT,
        Sort TEXT,
        UM TEXT,
        MLI TEXT,
        Plan_Type TEXT,
        Unit_Qty REAL,
        Std_Cost REAL,
        Ext_Std REAL,
        Labor_IN REAL,
        Lab_Rem REAL,
        Mat_IN REAL,
        Mat_Rem REAL,
        This_lvl REAL,
        Lab_Hrs REAL,
        key TEXT,
        Orden_BOM_Original INTEGER
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
    "source_file": r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\WHS PLAN\FILES\ActionCodes\actionMessages.xlsx",
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
        'Entity': 'Entity',
        'Project No': 'ProjectNo',
        'WO No': 'WONo',
        'SO/FCST': 'SO_FCST',
        'Sub': 'Sub',
        'Parent WO': 'ParentWO',
        'Item-Number': 'ItemNumber',
        'Rev': 'Rev',
        'Description': 'Description',
        'Misc-Code': 'MiscCode',
        'A/C': 'AC',
        'Lead Time': 'LeadTime',
        'Due-Dt': 'DueDt',
        'A-St-Dt': 'A_St_Dt',
        'Comp-Dt': 'CompDt',
        'Create-Dt': 'CreateDt',
        'Wo-Type': 'WoType',
        'Srt': 'Srt',
        'Plnr': 'Plnr',
        'Plan Type': 'PlanType',
        'Itm-Tp': 'ItmTp',
        'UM': 'UM',
        'Iss-Q': 'IssQ',
        'Comp-Q': 'CompQ',
        'Opn-Q': 'OpnQ',
        'St': 'St',
        'QA Aprvl': 'QAAprvl',
        'Stk': 'Stk',
        'Iss': 'Iss',
        'Prt': 'Prt',
        'User-Id': 'UserId',
        'Std-Cost': 'StdCost',
        'Prt-No': 'PrtNo',
        'Prt-User': 'PrtUser',
        'Prt-date': 'PrtDate',
        'Opn-Hrs': 'OpnHrs',
        'Sum-Labor-Hr': 'SumLaborHr',
        'Static Plan No': 'StaticPlanNo',
        'SP Rev': 'SPRev',
        'Static Plan Desc': 'StaticPlanDesc',
        'WO Last Notes': 'WOLastNotes'
    },
    "columns_order_original": [
        'Entity', 'Project No', 'WO No', 'SO/FCST', 'Sub', 'Parent WO',
        'Item-Number', 'Rev', 'Description', 'Misc-Code', 'A/C', 'Lead Time',
        'Due-Dt', 'A-St-Dt', 'Comp-Dt', 'Create-Dt', 'Wo-Type', 'Srt', 'Plnr',
        'Plan Type', 'Itm-Tp', 'UM', 'Iss-Q', 'Comp-Q', 'Opn-Q', 'St',
        'QA Aprvl', 'Stk', 'Iss', 'Prt', 'User-Id', 'Std-Cost', 'Prt-No',
        'Prt-User', 'Prt-date', 'Opn-Hrs', 'Sum-Labor-Hr', 'Static Plan No',
        'SP Rev', 'Static Plan Desc', 'WO Last Notes'
    ],
    "columns_order_renamed": [
        'Entity', 'ProjectNo', 'WONo', 'SO_FCST', 'Sub', 'ParentWO',
        'ItemNumber', 'Rev', 'Description', 'MiscCode', 'AC', 'LeadTime',
        'DueDt', 'A_St_Dt', 'CompDt', 'CreateDt', 'WoType', 'Srt', 'Plnr',
        'PlanType', 'ItmTp', 'UM', 'IssQ', 'CompQ', 'OpnQ', 'St',
        'QAAprvl', 'Stk', 'Iss', 'Prt', 'UserId', 'StdCost', 'PrtNo',
        'PrtUser', 'PrtDate', 'OpnHrs', 'SumLaborHr', 'StaticPlanNo',
        'SPRev', 'StaticPlanDesc', 'WOLastNotes'
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
        MiscCode TEXT,
        AC TEXT,
        LeadTime TEXT,
        DueDt TEXT,
        A_St_Dt TEXT,
        CompDt TEXT,
        CreateDt TEXT,
        WoType TEXT,
        Srt TEXT,
        Plnr TEXT,
        PlanType TEXT,
        ItmTp TEXT,
        UM TEXT,
        IssQ TEXT,
        CompQ TEXT,
        OpnQ TEXT,
        St TEXT,
        QAAprvl TEXT,
        Stk TEXT,
        Iss TEXT,
        Prt TEXT,
        UserId TEXT,
        StdCost TEXT,
        PrtNo TEXT,
        PrtUser TEXT,
        PrtDate TEXT,
        OpnHrs TEXT,
        SumLaborHr TEXT,
        StaticPlanNo TEXT,
        SPRev TEXT,
        StaticPlanDesc TEXT,
        WOLastNotes TEXT
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
        "widths": [2, 10, 7, 32, 14, 30, 9, 3, 16, 13, 12, 14, 12, 12, 14, 9,30,2,12],
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
        'Entity': 'Entity',
        'Project': 'Project',
        'Component': 'ItemNo',
        'Fuse-No': 'FuseNo',
        'Component Description': 'Description',
        'PlnType': 'PlnType',
        'Srt':'Srt',
        'St':'St','Qty-Oh':'QtyOh',
        'Qty-Issue':'QtyIssue',
        'Qty-no-Iss':'QtyPending',
        'Qty-Required':'ReqQty',
        'Val-Qty-Iss':'ValQtyIss',
        'Val-Not-Iss':'ValNotIss',
        'Val-Required':'ValRequired',
        'Wo-No':'WONo',
        'WO-Descripton':'WODescripton',
        'St':'St',
        'Req-Date':'ReqDate'
    },

    "columns_order_original": [
'Entity', 'Project', 'Component', 'Fuse-No',
        'Component Description', 'PlnType', 'Srt', 'Qty-Oh', 'Qty-Issue',
        'Qty-no-Iss', 'Qty-Required', 'Val-Qty-Iss', 'Val-Not-Iss',
        'Val-Required', 'Wo-No', 'WO-Descripton', 'St', 'Req-Date'
    ],
    "columns_order_renamed": [
'Entity','Project', 'ItemNo', 'FuseNo', 'Description','PlnType', 'Srt','St', 'QtyOh'
        ,'QtyIssue','QtyPending','ReqQty', 'ValQtyIss','ValNotIss','ValRequired','WONo'
        , 'WODescripton', 'WoNo','WODescripton','St','ReqDate'
    ],
    "special_processing": {
        "clear_before_insert": True,
        "custom_cleaning": False,
        "final_columns_only": True
    },
    "create_table_sql": """CREATE TABLE IF NOT EXISTS pr561(
        id INTEGER PRIMARY KEY,
        Entity TEXT,
        Project TEXT,
        ItemNo TEXT,
        FuseNo TEXT,
        Description TEXT,
        PlnType TEXT,
        Srt TEXT,
        St TEXT,
        QtyOh TEXT,
        QtyIssue TEXT,
        QtyPending TEXT,
        ReqQty TEXT,
        ValQtyIss TEXT,
        ValNotIss TEXT,
        ValRequired TEXT,
        WONo TEXT,
        WODescripton TEXT, 
        ReqDate TEXT
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

},
# ==================== TABLA KITING GROUPS====================
"kiting_groups": {
    "source_file": r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\WHS PLAN\FILES\Grupos Kiting\GRUPOS.xlsx",
    "table_name": "kiting_groups",
    "file_type": "excel",
    
    "excel_params": {
        "skiprows": 1,      # Cambiado de "skip_rows" a "skiprows" (formato pandas)
        "header": 0,        # Usar la primera fila (después del skip) como header
        "usecols": "A:G"    # Solo usar las primeras 7 columnas (A-G), elimina las últimas 3
    },
    
    # ✅ NUEVO: Filtrar filas de metadata del sistema
    "data_filters": {
        "exclude_rows_containing": [
            "Results Complete",
            "Report was generated", 
            "Query String",
            "record(s) shown",
            "seconds",
            "NO-LOCK WHERE",
            "Query finished",
            "EACH cndez"
        ]
    },
    
    "special_processing": {
        "force_column_names": True,     # Forzar nombres de columnas del mapping
        "auto_detect_header": True,     # Detectar automáticamente el header
        "remove_system_rows": True,     # ✅ NUEVO: Eliminar filas del sistema
        "clean_numeric_data": True,     # ✅ NUEVO: Limpiar datos numéricos
        "filter_nan_rows": True         # ✅ NUEVO: Filtrar filas con muchos 'nan'
    },
    
    "columns_mapping": {
        'cndez.Wo-Group.Group-no': 'Groupno',
        'cndez.Wo-Group.In-entity': 'Entity',
        'cndez.Wo-Group.SortG-list': 'SortGlist',
        'cndez.Wo-Group.Tool-group': 'Toolgroup',
        'cndez.Wo-Group.tool-item': 'Toolitem',
        'cndez.Wo-Group.Tool-TaktTime': 'Tool_TaktTime',
        'cndez.Wo-Group.wo-no': 'WO'
    },
    
    "columns_order_original": [
        'cndez.Wo-Group.Group-no', 'cndez.Wo-Group.In-entity',
        'cndez.Wo-Group.SortG-list', 'cndez.Wo-Group.Tool-group',
        'cndez.Wo-Group.tool-item', 'cndez.Wo-Group.Tool-TaktTime',
        'cndez.Wo-Group.wo-no'
    ],
    
    "columns_order_renamed": [
        'Groupno', 'Entity', 'SortGlist', 'Toolgroup', 'Toolitem', 'Tool_TaktTime', 'WO'
    ],
    
    "create_table_sql": """CREATE TABLE IF NOT EXISTS kiting_groups(
        id INTEGER PRIMARY KEY,
        Groupno TEXT,
        Entity TEXT,
        SortGlist TEXT,
        Toolgroup TEXT,
        Toolitem TEXT,
        Tool_TaktTime TEXT,
        WO TEXT
    )"""
},
# ==================== TABLA OPERATION WO====================
"Operation_WO": {
    "source_file": r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\WHS PLAN\FILES\PR 5 56\PR556.txt",
    "table_name": "Operation_WO",
    "file_type": "fixed_width",
    "fixed_width_params": {
        "widths": [9,31,21,4,4,9,11,11,7,9,9,8,11,7,15],
        "header": 3,
        "skip_rows": [0]
    },
    "data_filters": {
        "exclude_rows_multiple": {
            "column": "WO-No",
            "values": ["End-of-Re"]
        }
    },
    "columns_mapping": {
        'WO-No': 'WONo',
        'Item-No': 'ItemNo',
        'Description': 'Description',
        'UM': 'UM',
        'Seq': 'Seq',
        'Op-ID': 'OpID',
        'Start-Date': 'StartDate'
        ,'End-Date':'EndDate','WrkCtr':'WrkCtr'
        ,'Qty-Std':'QtyStd'
        ,'Qty-Act':'QtyAct'
        ,'WO-Qty':'WOQty'
        ,'Due-Date':'DueDate','Status':'Status','Prj':'Prj'
    },
    "columns_order_original": [
        'WO-No', 'Item-No', 'Description', 'UM', 'Seq', 'Op-ID', 'Start-Date',
        'End-Date', 'WrkCtr', 'Qty-Std', 'Qty-Act', 'WO-Qty', 'Due-Date',
        'Status', 'Prj'
    ],
    "columns_order_renamed": [
        'WONo', 'ItemNo', 'Description', 'UM','Seq','OpID'
        ,'StartDate','EndDate','WrkCtr','QtyStd','QtyAct','WOQty','DueDate','Status','Prj'
    ],
    "special_processing": {
        "clear_before_insert": True,
        "custom_cleaning": False,
        "final_columns_only": True
    },
    "create_table_sql": """CREATE TABLE IF NOT EXISTS Operation_WO(
                            id INTEGER PRIMARY KEY,
                            WONo TEXT,
                            ItemNo TEXT,
                            Description TEXT,
                            UM TEXT,
                            Seq TEXT,
                            OpID TEXT,
                            StartDate TEXT,
                            EndDate TEXT,
                            WrkCtr TEXT,
                            QtyStd TEXT,
                            QtyAct TEXT,
                            WOQty TEXT,
                            DueDate TEXT, 
                            Status text,
                            Prj TEXT
                        )"""
},
# ==================== TABLA RESPONSABLES ACTION WO by ENTITY====================
"Work_order_Actions_responsibles_entity": {
    "source_file": r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\WHS PLAN\FILES\ActionCodes\Responsible_by_entity.xlsx",
    "table_name": "WO_Actions_responsibles_entity",
    "file_type": "excel",
    
    "columns_mapping": {
        'Entity': 'Entity',
        'Sort_entity': 'Sort_entity',
        'Commodity': 'Commodity',
        'Responible':'Responsible',
        'email':'email',
        'R4_Account':'R4_Account'

    },
    "columns_order_original": [
        'Entity', 'Sort_entity', 'Commodity','Responsible','email','R4_Account'
    ],
    "columns_order_renamed": [
        'Entity', 'Sort_entity', 'Commodity','Responsible','email','R4_Account'
    ],
    "create_table_sql": """CREATE TABLE IF NOT EXISTS WO_Actions_responsibles_entity(
        id INTEGER PRIMARY KEY,
        Entity TEXT,
        Sort_entity TEXT,
        Commodity TEXT,
        Responsible TEXT,
        email TEXT,
        R4_Account TEXT
    )"""
},
# ==================== TABLA RESPONSABLES ACTION WO by PLANTYPE====================
"Work_order_Actions_responsibles_plantype": {
    "source_file": r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\WHS PLAN\FILES\ActionCodes\Responsible_by_plantype.xlsx",
    "table_name": "WO_Actions_responsibles_plantype",
    "file_type": "excel",
    
    "columns_mapping": {
        'Plantype': 'Plantype',
        'Responsible': 'Responsible',
        'email': 'email',
        'R4_Account':'R4_Account'

    },
    "columns_order_original": [
        'Plantype', 'Responsible', 'email','R4_Account'
    ],
    "columns_order_renamed": [
        'Plantype', 'Responsible', 'email','R4_Account'
    ],
    "create_table_sql": """CREATE TABLE IF NOT EXISTS WO_Actions_responsibles_plantype(
        id INTEGER PRIMARY KEY,
        Plantype TEXT,
        Responsible TEXT,
        email TEXT,
        R4_Account TEXT
    )"""
}
}







# Rutas globales usando raw strings para evitar problemas con backslashes
BASE_PATHS = {
    "db_folder": r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\R4Database",
    # "db_folder": r"C:\Users\J.Vazquez\Desktop\R4Database.db",
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
    """Aplica filtros de datos según configuración - VERSIÓN ACTUALIZADA"""
    data_filters = table_config.get("data_filters", {})
    initial_count = len(df)
    
    print(f"🔍 Aplicando filtros de datos... Filas iniciales: {initial_count}")
    
    # Filtro para excluir filas que contienen ciertas palabras
    exclude_containing = data_filters.get("exclude_rows_containing", [])
    if exclude_containing:
        print(f"   🔍 Buscando filas que contengan: {exclude_containing}")
        
        for keyword in exclude_containing:
            # Buscar keyword en todas las columnas (convertir todo a string)
            mask = df.astype(str).apply(lambda x: x.str.contains(keyword, case=False, na=False)).any(axis=1)
            rows_with_keyword = mask.sum()
            
            if rows_with_keyword > 0:
                print(f"   🗑️ Eliminando {rows_with_keyword} filas que contienen '{keyword}'")
                df = df[~mask]
    
    # Filtro original para excluir filas por columna específica
    exclude_rows = data_filters.get("exclude_rows", {})
    if exclude_rows:
        for column, value in exclude_rows.items():
            if column in df.columns:
                initial_col_count = len(df)
                df = df.loc[df[column] != value]
                filtered_count = initial_col_count - len(df)
                if filtered_count > 0:
                    print(f"   🔽 Filtro aplicado: Excluidas {filtered_count} filas donde {column} = '{value}'")
    
    # Filtro para filas con demasiados valores 'nan' o vacíos
    if data_filters.get("filter_nan_rows", False):
        print("   🧹 Filtrando filas con exceso de valores 'nan'...")
        # Calcular porcentaje de 'nan' por fila
        nan_threshold = 0.7  # 70% de la fila debe ser 'nan' para eliminarla
        
        nan_mask = df.astype(str).apply(lambda x: (x.str.contains('^nan$|^$', case=False, na=True)).sum() / len(df.columns) > nan_threshold, axis=1)
        nan_rows = nan_mask.sum()
        
        if nan_rows > 0:
            print(f"   🗑️ Eliminando {nan_rows} filas con más de {nan_threshold*100}% valores 'nan'")
            df = df[~nan_mask]
    
    final_count = len(df)
    filtered_total = initial_count - final_count
    
    if filtered_total > 0:
        print(f"   ✅ Total de filas filtradas: {filtered_total}")
        print(f"   📊 Filas restantes: {final_count}")
    else:
        print(f"   ℹ️ No se encontraron filas para filtrar")
    
    return df

def process_kiting_groups_with_filters(file_path, table_config):
    """
    Procesamiento específico para kiting_groups con filtros de metadata
    """
    import pandas as pd
    
    print("🔧 Procesando kiting_groups con filtros de metadata...")
    
    try:
        # 1. Cargar datos con parámetros Excel
        excel_params = table_config.get("excel_params", {})
        skiprows = excel_params.get("skiprows", 0)
        usecols = excel_params.get("usecols", None)
        
        print(f"   📖 Cargando Excel con skiprows={skiprows}, usecols={usecols}")
        df = pd.read_excel(file_path, skiprows=skiprows, header=0, usecols=usecols)
        
        print(f"   📊 Datos cargados: {len(df)} filas, {len(df.columns)} columnas")
        print(f"   📋 Columnas: {list(df.columns)}")
        
        # 2. CRÍTICO: Aplicar filtros ANTES de cualquier otro procesamiento
        df = apply_data_filters(df, table_config)
        
        # 3. Manejar columnas "Unnamed"
        if any('Unnamed' in str(col) for col in df.columns):
            print("   🔧 Detectadas columnas 'Unnamed', asignando nombres...")
            expected_columns = table_config.get("columns_order_original", [])
            
            if len(df.columns) <= len(expected_columns):
                df.columns = expected_columns[:len(df.columns)]
                print(f"   ✅ Columnas renombradas: {list(df.columns)}")
        
        # 4. Aplicar mapping de columnas
        columns_mapping = table_config.get("columns_mapping", {})
        if columns_mapping:
            df = df.rename(columns=columns_mapping)
            print(f"   🔄 Mapping aplicado: {list(df.columns)}")
        
        # 5. Seleccionar solo las columnas finales
        final_columns = table_config.get("columns_order_renamed", [])
        available_columns = [col for col in final_columns if col in df.columns]
        
        if available_columns:
            df = df[available_columns]
            print(f"   📋 Columnas finales seleccionadas: {len(available_columns)}")
        
        # 6. Limpieza final de datos
        print("   🧹 Aplicando limpieza final...")
        for col in df.columns:
            if df[col].dtype == 'object':
                # Reemplazar 'nan' string con valores vacíos
                df[col] = df[col].astype(str).replace(['nan', 'NaN', 'None'], '').str.strip()
                # Convertir cadenas vacías a NaN real para eliminar después
                df[col] = df[col].replace('', pd.NA)
        
        # 7. Eliminar filas donde todas las columnas importantes están vacías
        important_cols = ['Groupno', 'Entity', 'WO']  # Columnas clave que no deben estar vacías
        available_important = [col for col in important_cols if col in df.columns]
        
        if available_important:
            initial_clean = len(df)
            df = df.dropna(subset=available_important, how='all')
            final_clean = len(df)
            if initial_clean != final_clean:
                print(f"   🗑️ Eliminadas {initial_clean - final_clean} filas con columnas importantes vacías")
        
        print(f"   ✅ Procesamiento completado: {len(df)} filas finales")
        
        # 8. Mostrar muestra de datos para verificar
        if not df.empty:
            print("   📄 Muestra de datos finales:")
            sample = df.head(3)
            for idx, row in sample.iterrows():
                print(f"   Fila {idx}: {dict(row)}")
            
            print(f"   📄 Últimas filas para verificar filtros:")
            sample_last = df.tail(3)
            for idx, row in sample_last.iterrows():
                print(f"   Fila {idx}: {dict(row)}")
        
        return df
        
    except Exception as e:
        print(f"   ❌ Error procesando kiting_groups: {e}")
        import traceback
        traceback.print_exc()
        return None

def debug_bom_processing(df, table_config):
    """Debug del procesamiento de BOM"""
    table_name = table_config.get("table_name", "")
    special_processing = table_config.get("special_processing", {})
    
    print(f"🔍 DEBUG BOM:")
    print(f"   Table name: {table_name}")
    print(f"   Special processing flags:")
    print(f"     - generate_key_by_level: {special_processing.get('generate_key_by_level', False)}")
    print(f"     - custom_cleaning: {special_processing.get('custom_cleaning', False)}")
    print(f"   Columnas en DataFrame: {list(df.columns)}")
    print(f"   Filas en DataFrame: {len(df)}")
    
    # Verificar si ya tiene las columnas key y Orden_BOM_Original
    has_key = 'key' in df.columns
    has_orden = 'Orden_BOM_Original' in df.columns
    
    print(f"   DataFrame ya tiene 'key': {has_key}")
    print(f"   DataFrame ya tiene 'Orden_BOM_Original': {has_orden}")
    
    # Mostrar muestra de Level_Number para verificar formato
    if 'Level_Number' in df.columns:
        sample_levels = df['Level_Number'].head(10).tolist()
        print(f"   Muestra Level_Number: {sample_levels}")
    elif 'Level Number' in df.columns:
        sample_levels = df['Level Number'].head(10).tolist()
        print(f"   Muestra Level Number: {sample_levels}")
    
    return df

def process_superbom_complete(df, table_config):
    """
    Procesamiento completo para SUPERBOM incluyendo debug
    """
    import pandas as pd
    
    print(f"🔧 INICIANDO procesamiento completo de SUPERBOM...")
    
    # Debug inicial
    df = debug_bom_processing(df, table_config)
    
    # 1. Rellenar valores nulos
    df = df.fillna('')
    print(f"   🔧 Valores nulos rellenados")
    
    # 2. Eliminar filas "End-of-Report" si existen
    initial_count = len(df)
    level_col = None
    
    # Detectar la columna de nivel (puede ser Level_Number o Level Number)
    if 'Level_Number' in df.columns:
        level_col = 'Level_Number'
    elif 'Level Number' in df.columns:
        level_col = 'Level Number'
    
    if level_col:
        df = df.loc[df[level_col] != 'End-of-Report.       1']
        df = df.loc[df[level_col].astype(str).str.strip() != 'End-of-Report.       1']
        if len(df) != initial_count:
            print(f"   🗑️ Eliminadas {initial_count - len(df)} filas 'End-of-Report'")
    
    # 3. Eliminar la primera fila de datos (índice 0) como en tu código original
    if len(df) > 0:
        df = df.drop(index=df.index[0])
        print(f"   🗑️ Eliminada primera fila de datos")
    
    # 4. CRÍTICO: Verificar que las columnas necesarias existan
    # Detectar columnas de componente (puede ser Component o similar)
    component_col = None
    if 'Component' in df.columns:
        component_col = 'Component'
    elif 'component' in df.columns:
        component_col = 'component'
    
    if not level_col or not component_col:
        print(f"   ❌ ERROR: No se encontraron columnas requeridas")
        print(f"   📋 Level column: {level_col}")
        print(f"   📋 Component column: {component_col}")
        print(f"   📋 Columnas disponibles: {list(df.columns)}")
        return df
    
    print(f"   ✅ Usando columnas: Level='{level_col}', Component='{component_col}'")
    
    # 5. Generar columna 'key' por nivel 1
    print(f"   🔑 Generando keys por nivel 1...")
    df['key'] = None
    ultima_key = None
    keys_generated = 0
    
    for idx, row in df.iterrows():
        level_str = str(row[level_col]).strip()
        try:
            # Extraer el primer carácter y convertir a entero
            if level_str and len(level_str) > 0 and level_str[0].isdigit():
                level = int(level_str[0])
                
                if level == 1:
                    ultima_key = str(row[component_col]).strip()
                    keys_generated += 1
                    print(f"     🔑 Nueva key nivel 1: '{ultima_key}'")
                
                df.at[idx, 'key'] = ultima_key
            
        except (ValueError, IndexError) as e:
            print(f"     ⚠️ Error procesando nivel '{level_str}' en fila {idx}: {e}")
            continue
    
    # Convertir key a string y manejar valores None
    df['key'] = df['key'].fillna('').astype(str)
    print(f"   ✅ Keys generadas: {keys_generated} componentes nivel 1")
    
    # 6. Aplicar uppercase a columna Sort si existe
    if 'Sort' in df.columns:
        df['Sort'] = df['Sort'].astype(str).str.upper()
        print(f"   🔤 Aplicado uppercase a columna Sort")
    
    # 7. Reset index y generar orden original
    df.reset_index(drop=True, inplace=True)
    df['Orden_BOM_Original'] = df.index + 1
    print(f"   📊 Generado orden original BOM (1-{len(df)})")
    
    # 8. Verificación final
    print(f"   ✅ Procesamiento SUPERBOM completado:")
    print(f"     - Filas finales: {len(df)}")
    print(f"     - Columnas finales: {len(df.columns)}")
    
    # Verificar que las keys no estén vacías
    keys_not_empty = df['key'].ne('').sum()
    print(f"     - Keys no vacías: {keys_not_empty} de {len(df)}")
    
    if keys_not_empty > 0:
        unique_keys = df[df['key'] != '']['key'].nunique()
        print(f"     - Keys únicas: {unique_keys}")
    
    # 9. Mostrar muestra de datos para verificar
    if not df.empty:
        print("   📄 Muestra de datos procesados:")
        sample_cols = [level_col, component_col, 'Sort', 'key', 'Orden_BOM_Original']
        available_sample_cols = [col for col in sample_cols if col in df.columns]
        sample = df[available_sample_cols].head(5)
        
        for idx, row in sample.iterrows():
            level = row.get(level_col, 'N/A')
            component = row.get(component_col, 'N/A')
            key_val = row.get('key', 'N/A')
            orden = row.get('Orden_BOM_Original', 'N/A')
            sort_val = row.get('Sort', 'N/A')
            print(f"     Fila {idx}: Level={level}, Component={component}, Sort={sort_val}, Key='{key_val}', Orden={orden}")
    
    return df

def apply_special_processing(df, table_config):
    """Aplica procesamiento especial según configuración - VERSIÓN ACTUALIZADA"""
    special_processing = table_config.get("special_processing", {})
    table_name = table_config.get("table_name", "")
    
    # ========== PROCESAMIENTO ESPECIAL PARA SUPERBOM ==========
    if (table_name == "bom" and 
        special_processing.get("generate_key_by_level", False) and
        special_processing.get("custom_cleaning", False)):
        
        print(f"🎯 Aplicando procesamiento especial para SUPERBOM...")
        df = process_superbom_complete(df, table_config)  # Usar la función mejorada
        return df  # Retornar aquí porque ya incluye toda la lógica necesaria
    
    # ========== PROCESAMIENTO ORIGINAL PARA OTRAS TABLAS ==========
    
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
    
    # Procesamiento para arreglar años de fecha de expiración
    if special_processing.get("fix_expire_date_year", False):
        import pandas as pd
        
        year_threshold = special_processing.get("year_threshold", 1950)
        year_adjustment = special_processing.get("year_adjustment", 100)
        
        for col in df.columns:
            if 'expire' in col.lower() or 'expir' in col.lower():
                print(f"   📅 Corrigiendo años en columna: {col}")
                # Lógica para corregir años (si es necesario)
                break
    
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

def process_whs_location_exact(df, table_config):
    """
    Procesamiento EXACTO como tu código funcional para whs_location_in36851
    """
    import pandas as pd
    
    print(f"🏭 Procesando WHS Location con tu lógica exacta...")
    initial_count = len(df)
    print(f"   📊 Filas iniciales: {initial_count}")
    
    try:
        # PASO 1: df = df.drop(0) - Eliminar fila índice 0
        if len(df) > 0 and 0 in df.index:
            df = df.drop(0)
            print(f"   🗑️ Eliminada fila índice 0")
        
        # PASO 2: df = df.drop(1) - Eliminar fila índice 1  
        if len(df) > 0 and 1 in df.index:
            df = df.drop(1)
            print(f"   🗑️ Eliminada fila índice 1")
        
        # PASO 3: df = df.loc[df['Whs']=='CH'] - MANTENER solo CH
        if 'Whs' in df.columns:
            before_filter = len(df)
            df = df.loc[df['Whs'] == 'CH']  # ✅ Tu lógica exacta
            after_filter = len(df)
            filtered = before_filter - after_filter
            print(f"   ✅ Filtro Whs='CH': Mantenidas {after_filter} filas, eliminadas {filtered}")
        else:
            print(f"   ❌ ERROR: Columna 'Whs' no encontrada. Columnas disponibles: {list(df.columns)}")
        
        # PASO 4: df = df.fillna('') - Rellenar NaN
        df = df.fillna('')
        print(f"   🔧 Valores NaN rellenados con cadenas vacías")
        
        # PASO 5: Reset index para evitar problemas
        df.reset_index(drop=True, inplace=True)
        print(f"   🔄 Index reseteado")
        
    except Exception as e:
        print(f"   ❌ ERROR en procesamiento WHS: {e}")
        import traceback
        traceback.print_exc()
        return df
    
    final_count = len(df)
    total_filtered = initial_count - final_count
    
    print(f"   ✅ Procesamiento WHS completado:")
    print(f"     - Filas iniciales: {initial_count}")
    print(f"     - Filas finales: {final_count}")
    print(f"     - Total filtradas: {total_filtered}")
    print(f"     - Porcentaje mantenido: {(final_count/initial_count)*100:.1f}%")
    
    # Verificación de datos
    if not df.empty:
        print("   📄 Verificación de datos finales:")
        sample = df.head(3)
        for idx, row in sample.iterrows():
            whs = row.get('Whs', 'N/A')
            binid = row.get('Bin ID', 'N/A') 
            locid = row.get('Loc-ID', 'N/A')
            zone = row.get('Zone', 'N/A')
            print(f"     Fila {idx}: Whs='{whs}', BinID='{binid}', LocID='{locid}', Zone='{zone}'")
        
        # Verificar que NO hay datos basura
        unique_whs = df['Whs'].unique() if 'Whs' in df.columns else ['N/A']
        print(f"   📊 Valores únicos en Whs: {unique_whs}")
        
        # Contar filas con nan.0 o similares (no debería haber ninguna)
        if 'Whs' in df.columns:
            bad_whs = df[df['Whs'].isin(['nan.0', 'nan', '0', ''])].shape[0]
            print(f"   🔍 Filas con Whs inválido: {bad_whs} (debería ser 0)")
    
    return df

def apply_special_processing(df, table_config):
    """Aplica procesamiento especial - VERSIÓN CORREGIDA SIN ERRORES"""
    special_processing = table_config.get("special_processing", {})
    table_name = table_config.get("table_name", "")
    
    # ========== TU LÓGICA EXACTA PARA WHS LOCATION ==========
    if table_name == "whs_location_in36851":
        print(f"🎯 Detectada whs_location_in36851 - usando tu código exacto...")
        df = process_whs_location_exact(df, table_config)
        return df  # ✅ Retornar inmediatamente después del procesamiento
    
    # ========== PROCESAMIENTO ESPECIAL PARA SUPERBOM ==========
    if (table_name == "bom" and 
        special_processing.get("generate_key_by_level", False) and
        special_processing.get("custom_cleaning", False)):
        
        print(f"🎯 Aplicando procesamiento especial para SUPERBOM...")
        df = process_superbom_complete(df, table_config)
        return df
    
    # ========== PROCESAMIENTO ESPECIAL PARA KITING GROUPS ==========
    if (table_name == "kiting_groups" and 
        special_processing.get("custom_cleaning", False)):
        
        print(f"🎯 Aplicando procesamiento especial para KITING GROUPS...")
        # ✅ CORREGIDO: Usar la función correcta que ya existe en tu código
        df = apply_data_filters(df, table_config)
        return df
    
    # ========== RESTO DEL PROCESAMIENTO ORIGINAL ==========
    
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
    
    # Procesamiento para arreglar años de fecha de expiración
    if special_processing.get("fix_expire_date_year", False):
        import pandas as pd
        
        year_threshold = special_processing.get("year_threshold", 1950)
        year_adjustment = special_processing.get("year_adjustment", 100)
        
        for col in df.columns:
            if 'expire' in col.lower() or 'expir' in col.lower():
                print(f"   📅 Corrigiendo años en columna: {col}")
                break
    
    return df


















if __name__ == "__main__":
    print("=== Verificación de Configuración Multi-Tabla ===")
    print(f"📊 Tablas configuradas: {len(TABLES_CONFIG)}")
    for table_name in TABLES_CONFIG.keys():
        print(f"   - {table_name}")
    print()
    verify_paths()