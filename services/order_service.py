import pandas as pd


class OrderService:
    def __init__(self, source):
        """
        source can be:
        - file path
        - file-like object (e.g. BytesIO)
        """
        self.source = source
        self.df = None

    # --------------------------------------------------
    # LOAD & PREPARE ORDER DATA
    # --------------------------------------------------
    def load_data(self):
        self.df = pd.read_excel(self.source)

        # Normalize column names (trim spaces)
        self.df.columns = (
            self.df.columns
            .astype(str)
            .str.strip()
        )

    # --------------------------------------------------
    # CORE ORDER QUERIES
    # --------------------------------------------------
    def get_completed_orders(self) -> pd.DataFrame:
        if self.df is None:
            self.load_data()

        return self.df[
            self.df["Order Status"]
            .astype(str)
            .str.lower()
            == "completed"
        ]

    def get_completed_count(self) -> int:
        return len(self.get_completed_orders())

    def get_summary(self) -> dict:
        if self.df is None:
            self.load_data()

        return {
            "total_orders": len(self.df),
            "completed_orders": self.get_completed_count(),
        }

    # --------------------------------------------------
    # 💰 PROJECTED INCOME (NEW FEATURE)
    # --------------------------------------------------
    def get_projected_income_total(self) -> float:
        """
        Sum of 'Product Subtotal' for all completed orders.
        Represents projected gross income.
        """
        completed = self.get_completed_orders()

        if "Product Subtotal" not in completed.columns:
            raise ValueError(
                "❌ 'Product Subtotal' column not found in orders file"
            )

        # Convert values safely to numeric
        subtotals = pd.to_numeric(
            completed["Product Subtotal"],
            errors="coerce"
        ).fillna(0)

        return float(subtotals.sum())
