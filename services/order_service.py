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
    # 💰 PROJECTED INCOME
    # --------------------------------------------------
    def get_projected_income_total(self) -> float:
        completed = self.get_completed_orders()

        if "Product Subtotal" not in completed.columns:
            raise ValueError(
                "❌ 'Product Subtotal' column not found in orders file"
            )

        subtotals = pd.to_numeric(
            completed["Product Subtotal"],
            errors="coerce"
        ).fillna(0)

        return float(subtotals.sum())

    # --------------------------------------------------
    # 🏆 TOP 20 PRODUCTS (SORTED BY TOTAL REVENUE)
    # --------------------------------------------------
    def get_top_20_products_completed(self) -> list[dict]:
        """
        Returns Top 20 products from COMPLETED orders,
        sorted by TOTAL REVENUE (highest first).
        """
        completed = self.get_completed_orders()

        required_cols = {"Product Name", "Quantity", "Product Subtotal"}
        if not required_cols.issubset(completed.columns):
            missing = required_cols - set(completed.columns)
            raise ValueError(
                f"❌ Missing required columns: {', '.join(missing)}"
            )

        completed = completed.copy()

        completed["Quantity"] = pd.to_numeric(
            completed["Quantity"],
            errors="coerce"
        ).fillna(0)

        completed["Product Subtotal"] = pd.to_numeric(
            completed["Product Subtotal"],
            errors="coerce"
        ).fillna(0)

        grouped = (
            completed
            .groupby("Product Name", dropna=False)
            .agg(
                total_quantity=("Quantity", "sum"),
                total_revenue=("Product Subtotal", "sum")
            )
            .reset_index()
            # 🔥 SORT BY REVENUE INSTEAD OF QUANTITY
            .sort_values("total_revenue", ascending=False)
            .head(20)
        )

        return [
            {
                "Product Name": row["Product Name"],
                "Total Quantity Sold": int(row["total_quantity"]),
                "Total Revenue": float(row["total_revenue"]),
            }
            for _, row in grouped.iterrows()
        ]
