from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import pandas as pd
from pywinauto import Application
from pywinauto.keyboard import send_keys
import os
import re
import logging
import time

logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "http://127.0.0.1:5173"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class QueryParams(BaseModel):
    delivery_date: str

def validate_date(date_str: str) -> bool:
    pattern = r"^(>|>=|<|<=|=)?\d{1,2}/\d{1,2}/\d{4}$"
    return bool(re.match(pattern, date_str))

@app.get("/test")
async def test():
    return {"message": "Server is up!"}

@app.post("/export-csv")
async def export_csv(params: QueryParams):
    try:
        logger.debug(f"Received delivery_date: {params.delivery_date}")
        if not validate_date(params.delivery_date):
            raise HTTPException(status_code=400, detail="Invalid delivery_date format. Use e.g., '>9/2/2025'")

        logger.debug("Connecting to FileMaker")
        app = Application(backend="uia").connect(title_re=".*Pacific Solutionsv10.*", timeout=10)
        window = app.window(title_re=".*Pacific Solutionsv10.*")
        logger.debug("Waiting for window")
        window.wait("ready", timeout=10)

        logger.debug("Clicking Purchase Orders button")
        window.set_focus()
        window.click_input(coords=(1528, 726))  # From Accessibility Insights
        time.sleep(3)

        logger.debug("Switching to POv 10 window")
        window = app.window(title_re=".*POv 10.*")
        window.wait("ready", timeout=10)

        logger.debug("Clicking first Find button")
        window.click_input(coords=(1280, 724))  # From Accessibility Insights
        time.sleep(2)

        logger.debug("Clicking Order Date field")
        window.click_input(coords=(1259, 571))  # From Accessibility Insights
        time.sleep(1)
        logger.debug(f"Typing query in Order Date field: {params.delivery_date}")
        send_keys(params.delivery_date + "{ENTER}")
        time.sleep(2)

        logger.debug("Clicking second Find button to execute query")
        window.click_input(coords=(1275, 706))  # From Accessibility Insights
        time.sleep(2)

        logger.debug("Clicking Excel button")
        window.click_input(coords=(785, 55))  # From Accessibility Insights
        time.sleep(2)

        logger.debug("Selecting PO Items in export options")
        export_options_dialog = app.window(title_re=".*Export.*|.*Options.*")
        export_options_dialog.wait("ready", timeout=10)
        export_options_dialog.child_window(title="PO Items", control_type="Button").click()
        time.sleep(2)

        logger.debug("Opening Save As dialog")
        csv_path = r"C:\Temp\export.csv"
        save_dialog = app.window(title_re="Save As")
        save_dialog.wait("ready", timeout=10)
        save_dialog.ComboBox.select("Comma-Separated Text (*.csv)")
        save_dialog.Edit.set_text(csv_path)
        save_dialog.child_window(title="Save", control_type="Button").click()
        time.sleep(2)

        logger.debug("Handling Specify Field Order dialog")
        field_order_dialog = app.window(title_re="Specify Field Order.*")
        field_order_dialog.wait("ready", timeout=10)
        field_order_dialog.child_window(title="Export", control_type="Button").click()
        time.sleep(2)

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
        logger.error(f"Unexpected error: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Unexpected error: {str(e)}")