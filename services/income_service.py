import pandas as pd
from typing import Optional
from io import BytesIO
from services.order_service import OrderService


class IncomeService:
    def __init__(self, source, order_service: OrderService):
        """
        source can be:
        - file path
        - file-like object (BytesIO)
        """
        self.source = source
        self.order_service = order_service
        self.income_df = None
        self.order_id_column = None

    # --------------------------------------------------
    # LOAD & PREPARE INCOME DATA
    # --------------------------------------------------
    def load_income_data(self):
        excel = pd.ExcelFile(self.source)
        sheet_names = excel.sheet_names

        # Prefer explicit Income sheet
        sheet_name = "Income" if "Income" in sheet_names else sheet_names[0]

        # Preview rows to detect header
        preview = pd.read_excel(
            excel,
            sheet_name=sheet_name,
            header=None,
            nrows=30
        )

        header_row = self._detect_header_row(preview)

        if header_row is None:
            raise ValueError("❌ Could not detect header row in Income sheet")

        # Load actual data
        self.income_df = pd.read_excel(
            excel,
            sheet_name=sheet_name,
            header=header_row
        )

        # Normalize column names
        self.income_df.columns = (
            self.income_df.columns
            .astype(str)
            .str.strip()
            .str.lower()
        )

        self.order_id_column = self._detect_order_id_column()

        if self.order_id_column is None:
            raise ValueError(
                f"❌ Order ID column not found. Columns: {list(self.income_df.columns)}"
            )

    def _detect_header_row(self, preview_df: pd.DataFrame) -> Optional[int]:
        keywords = ["order id", "order no", "order number", "ordersn"]

        for idx, row in preview_df.iterrows():
            values = row.astype(str).str.lower()
            if any(any(k in v for k in keywords) for v in values):
                return idx

        return None

    def _detect_order_id_column(self) -> Optional[str]:
        candidates = ["order id", "order no", "order number", "ordersn"]

        for col in self.income_df.columns:
            for key in candidates:
                if key in col:
                    return col

        return None

    # --------------------------------------------------
    # CORE LOGIC
    # --------------------------------------------------
    def get_income_order_ids(self) -> set:
        if self.income_df is None:
            self.load_income_data()

        return set(
            self.income_df[self.order_id_column]
            .astype(str)
            .str.strip()
            .unique()
        )

    def find_missing_income_orders(self) -> pd.DataFrame:
        completed_orders = self.order_service.get_completed_orders()
        income_ids = self.get_income_order_ids()

        return completed_orders[
            ~completed_orders["Order ID"]
            .astype(str)
            .str.strip()
            .isin(income_ids)
        ]

    def get_reconciliation_summary(self) -> dict:
        completed_ids = set(
            self.order_service
            .get_completed_orders()["Order ID"]
            .astype(str)
            .str.strip()
        )

        income_ids = self.get_income_order_ids()

        return {
            "completed_orders": len(completed_ids),
            "orders_with_income": len(income_ids),
            "missing_income_orders": len(completed_ids - income_ids),
        }

    # --------------------------------------------------
    # REPORTING
    # --------------------------------------------------
    def get_missing_income_report(self) -> pd.DataFrame:
        """
        Detailed report for completed orders with no income record.
        """
        missing = self.find_missing_income_orders()

        # Normalize columns
        missing.columns = (
            missing.columns
            .astype(str)
            .str.strip()
            .str.lower()
        )

        column_map = {
            "order id": "Order ID",
            "tracking number*": "Tracking Number",
            "estimated ship out date": "Estimated Ship Out Date",
            "order creation date": "Order Creation Date",
            "product name": "Product Name",
            "variation name": "Variation Name",
            "original price": "Original Price",
            "deal price": "Deal Price",
            "product subtotal": "Product Subtotal",
            "username (buyer)": "Username (Buyer)",
        }

        selected_cols = []
        rename_map = {}

        for col, display in column_map.items():
            if col in missing.columns:
                selected_cols.append(col)
                rename_map[col] = display

        report = missing[selected_cols].rename(columns=rename_map)

        # ✅ FINAL FIX:
        # Replace NaN / None with empty string (clean UI & Excel)
        report = report.fillna("")

        return report

    def export_missing_orders_to_excel(self) -> BytesIO:
        """
        Export missing income report to an in-memory Excel file.
        """
        report = self.get_missing_income_report()

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            report.to_excel(
                writer,
                index=False,
                sheet_name="Missing Income Orders"
            )

        output.seek(0)
        return output
