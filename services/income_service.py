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
        if self.income_df is not None:
            return

        excel = pd.ExcelFile(self.source)
        sheet_names = excel.sheet_names

        # Prefer Income sheet if present
        sheet_name = "Income" if "Income" in sheet_names else sheet_names[0]

        preview = pd.read_excel(
            excel,
            sheet_name=sheet_name,
            header=None,
            nrows=30
        )

        header_row = self._detect_header_row(preview)
        if header_row is None:
            raise ValueError("❌ Could not detect header row in Income sheet")

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

    def _safe_sum(self, df: pd.DataFrame, column_name: str) -> float:
        col = column_name.lower()
        if col not in df.columns:
            return 0.0
        return float(
            pd.to_numeric(df[col], errors="coerce")
            .fillna(0)
            .sum()
        )

    # --------------------------------------------------
    # CORE MATCHING
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

    # --------------------------------------------------
    # 📦 RECONCILIATION SUMMARY (RESTORED)
    # --------------------------------------------------
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
    # 💵 ACTUAL RECEIVED INCOME SUMMARY
    # --------------------------------------------------
    def get_actual_received_income_summary(self) -> dict:
        if self.income_df is None:
            self.load_income_data()

        completed_ids = set(
            self.order_service
            .get_completed_orders()["Order ID"]
            .astype(str)
            .str.strip()
        )

        matched = self.income_df[
            self.income_df[self.order_id_column]
            .astype(str)
            .str.strip()
            .isin(completed_ids)
        ]

        projected_income = self.order_service.get_projected_income_total()

        return {
            "Projected Income": projected_income,
            "Actual Received Income": self._safe_sum(
                matched, "total released amount (₱)"
            ),
            "AMS Commission Fee": self._safe_sum(matched, "ams commission fee"),
            "Commission Fee": self._safe_sum(matched, "commission fee"),
            "Service Fee": self._safe_sum(matched, "service fee"),
            "Support Program Fee": self._safe_sum(matched, "support program fee"),
            "Transaction Fee": self._safe_sum(matched, "transaction fee"),
            "Withholding Tax": self._safe_sum(matched, "withholding tax"),
            "Total Released Amount (₱)": self._safe_sum(
                matched, "total released amount (₱)"
            ),
        }

    # --------------------------------------------------
    # 🔄 RETURN / REFUND SUMMARY
    # --------------------------------------------------
    def get_return_refund_summary(self) -> dict:
        if self.income_df is None:
            self.load_income_data()

        completed_ids = set(
            self.order_service
            .get_completed_orders()["Order ID"]
            .astype(str)
            .str.strip()
        )

        refunded = self.income_df[
            (self.income_df[self.order_id_column]
             .astype(str)
             .str.strip()
             .isin(completed_ids)) &
            (self.income_df.get("refund id").notna()) &
            (self.income_df["refund id"].astype(str).str.strip() != "")
        ]

        refund_amount = self._safe_sum(refunded, "refund amount")
        reverse_shipping_fee = self._safe_sum(
            refunded, "reverse shipping fee"
        )

        return {
            "Return/Refund Count": len(refunded),
            "Total Revenue Loss": abs(refund_amount),
            "Reverse Shipping Fee Charged to Seller": abs(reverse_shipping_fee),
        }

    # --------------------------------------------------
    # 🔄 RETURN / REFUND DETAILS (CLEAN VIEW)
    # --------------------------------------------------
    def get_return_refund_details(self) -> pd.DataFrame:
        if self.income_df is None:
            self.load_income_data()

        completed_ids = set(
            self.order_service
            .get_completed_orders()["Order ID"]
            .astype(str)
            .str.strip()
        )

        refunded = self.income_df[
            (self.income_df[self.order_id_column]
             .astype(str)
             .str.strip()
             .isin(completed_ids)) &
            (self.income_df.get("refund id").notna()) &
            (self.income_df["refund id"].astype(str).str.strip() != "")
        ].copy()

        column_map = {
            self.order_id_column: "Order ID",
            "refund id": "Refund ID",
            "username (buyer)": "Username (Buyer)",
            "order creation date": "Order Creation Date",
            "refund amount": "Refund Amount",
            "shipping fee rebate from shopee":
                "Shipping Fee Rebate From Shopee",
            "3rd party logistics - defined shipping fee":
                "3rd Party Logistics - Defined Shipping Fee",
            "reverse shipping fee": "Reverse Shipping Fee",
            "total released amount (₱)": "Total Released Amount (₱)",
            "cash refund to buyer amount": "Cash Refund to Buyer Amount",
        }

        existing = {
            c: l for c, l in column_map.items() if c in refunded.columns
        }

        cleaned = refunded[list(existing.keys())].rename(columns=existing)

        numeric_cols = [
            "Refund Amount",
            "Shipping Fee Rebate From Shopee",
            "3rd Party Logistics - Defined Shipping Fee",
            "Reverse Shipping Fee",
            "Total Released Amount (₱)",
            "Cash Refund to Buyer Amount",
        ]

        for col in numeric_cols:
            if col in cleaned.columns:
                cleaned[col] = (
                    pd.to_numeric(cleaned[col], errors="coerce")
                    .fillna(0)
                    .abs()
                )

        return cleaned

    # --------------------------------------------------
    # 🚚 OVERCHARGE SHIPPING FEE SUMMARY
    # --------------------------------------------------
    def get_overcharge_shipping_fee_summary(self, limit: int = 50) -> pd.DataFrame:
        if self.income_df is None:
            self.load_income_data()

        completed_ids = set(
            self.order_service
            .get_completed_orders()["Order ID"]
            .astype(str)
            .str.strip()
        )

        df = self.income_df[
            self.income_df[self.order_id_column]
            .astype(str)
            .str.strip()
            .isin(completed_ids)
        ].copy()

        column_map = {
            self.order_id_column: "Order ID",
            "username (buyer)": "Username (Buyer)",
            "buyer paid shipping fee": "Buyer Paid Shipping Fee",
            "shipping fee rebate from shopee":
                "Shipping Fee Rebate From Shopee",
            "3rd party logistics - defined shipping fee":
                "3rd Party Logistics - Defined Shipping Fee",
        }

        existing = {
            c: l for c, l in column_map.items() if c in df.columns
        }

        if "3rd party logistics - defined shipping fee" not in df.columns:
            return pd.DataFrame(columns=existing.values())

        cleaned = df[list(existing.keys())].rename(columns=existing)

        for col in [
            "Buyer Paid Shipping Fee",
            "Shipping Fee Rebate From Shopee",
            "3rd Party Logistics - Defined Shipping Fee",
        ]:
            if col in cleaned.columns:
                cleaned[col] = (
                    pd.to_numeric(cleaned[col], errors="coerce")
                    .fillna(0)
                    .abs()
                )

        cleaned = cleaned.sort_values(
            by="3rd Party Logistics - Defined Shipping Fee",
            ascending=False
        ).head(limit)

        return cleaned.reset_index(drop=True)

    # --------------------------------------------------
    # 📤 EXPORT
    # --------------------------------------------------
    def export_missing_orders_to_excel(self) -> BytesIO:
        report = self.get_missing_income_report()
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            report.to_excel(writer, index=False)
        output.seek(0)
        return output

    def get_missing_income_report(self) -> pd.DataFrame:
        missing = self.find_missing_income_orders()

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

        cols, rename = [], {}
        for c, d in column_map.items():
            if c in missing.columns:
                cols.append(c)
                rename[c] = d

        report = missing[cols].rename(columns=rename)
        return report.fillna("")
