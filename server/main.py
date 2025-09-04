from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
import pandas as pd
from pywinauto import Application
import os
import re
import logging

logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

app = FastAPI()

class QueryParams(BaseModel):
    delivery_date: str

def validate_date(date_str: str) -> bool:
    pattern = r"^(>|>=|<|<=|=)?\d{1,2}/\d{1,2}/\d{4}$"
    return bool(re.match(pattern, date_str))

@app.post("/export-csv")
async def export_csv(params: QueryParams):
    try:
        logger.debug(f"Received delivery_date: {params.delivery_date}")
        if not validate_date(params.delivery_date):
            raise HTTPException(status_code=400, detail="Invalid delivery_date format. Use e.g., '>9/2/2025'")
        logger.debug("Connecting to FileMaker")
        app = Application(backend="uia").connect(title="Pacific_Solutionsv10 (CF-VM-App)", timeout=10)
        window = app.window(title="Pacific_Solutionsv10 (CF-VM-App)")
        window.wait("ready", timeout=10)
        logger.debug("Window ready, entering Find mode")
        window.type_keys("^f")
        window.wait("ready", timeout=2)
        logger.debug(f"Typing query: {params.delivery_date}")
        window.type_keys(params.delivery_date + "{ENTER}")
        window.wait("ready", timeout=2)
        logger.debug("Opening Export dialog")
        csv_path = r"C:\Temp\export.csv"
        window.type_keys("^+e")
        export_dialog = app.window(title_re=".*Export.*")
        export_dialog.wait("ready", timeout=5)
        logger.debug("Selecting CSV format")
        export_dialog.ComboBox.select("Comma-Separated Text (*.csv)")
        export_dialog.Edit.set_text(csv_path)
        export_dialog.Save.click()
        window.wait("ready", timeout=5)
        logger.debug(f"Checking for CSV at {csv_path}")
        if not os.path.exists(csv_path):
            raise HTTPException(status_code=500, detail="CSV export failed")
        logger.debug("Reading CSV")
        df = pd.read_csv(csv_path, encoding="utf-8", quotechar='"')
        expected_headers = [
            "PO Number", "Job Number", "Customer", "Address", "Vendor",
            "Order Date", "Entry Date", "Ship Date", "Delivery Date",
            "Status", "Amount", "Vendor Order Number", "Ship Via",
            "Salesperson", "Notes"
        ]
        if list(df.columns) != expected_headers:
            raise HTTPException(status_code=400, detail="Unexpected CSV headers")
        df = df.fillna({"Ship Date": "", "Delivery Date": "", "Notes": ""})
        df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce")
        for col in ["Order Date", "Entry Date", "Ship Date", "Delivery Date"]:
            df[col] = pd.to_datetime(df[col], format="%m/%d/%Y", errors="coerce").dt.strftime("%Y-%m-%d")
        df["Days to Delivery"] = df["Delivery Date"].apply(
            lambda x: (pd.to_datetime(x) - pd.to_datetime("2025-09-02")).days if pd.notnull(x) else None
        )
        logger.debug("Returning data")
        return {"data": df.to_dict(orient="records")}
    except Exception as e:
        logger.error(f"Error: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error: {str(e)}")