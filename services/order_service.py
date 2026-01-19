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
        if self.df is not None:
            return

        self.df = pd.read_excel(self.source)

        # Normalize column names (trim spaces only, keep case)
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
    # 📦 TOP 20 HIGH SALES PRODUCTS (COMPLETED)
    # --------------------------------------------------
    def get_top_20_products_completed(self) -> pd.DataFrame:
        completed = self.get_completed_orders()

        required_columns = [
            "Product Name",
            "Quantity",
            "Product Subtotal",
        ]

        for col in required_columns:
            if col not in completed.columns:
                raise ValueError(f"❌ Missing column: {col}")

        completed["Quantity"] = pd.to_numeric(
            completed["Quantity"], errors="coerce"
        ).fillna(0)

        completed["Product Subtotal"] = pd.to_numeric(
            completed["Product Subtotal"], errors="coerce"
        ).fillna(0)

        grouped = (
            completed
            .groupby("Product Name", as_index=False)
            .agg({
                "Quantity": "sum",
                "Product Subtotal": "sum"
            })
        )

        grouped.rename(columns={
            "Quantity": "Total Quantity Sold",
            "Product Subtotal": "Total Revenue"
        }, inplace=True)

        return (
            grouped
            .sort_values("Total Revenue", ascending=False)
            .head(20)
            .reset_index(drop=True)
        )

    # --------------------------------------------------
    # 📉 TOP 20 LEAST SALES PRODUCTS (COMPLETED)
    # --------------------------------------------------
    def get_top_20_least_products_completed(self) -> pd.DataFrame:
        completed = self.get_completed_orders()

        required_columns = [
            "Product Name",
            "Quantity",
            "Product Subtotal",
        ]

        for col in required_columns:
            if col not in completed.columns:
                raise ValueError(f"❌ Missing column: {col}")

        completed["Quantity"] = pd.to_numeric(
            completed["Quantity"], errors="coerce"
        ).fillna(0)

        completed["Product Subtotal"] = pd.to_numeric(
            completed["Product Subtotal"], errors="coerce"
        ).fillna(0)

        grouped = (
            completed
            .groupby("Product Name", as_index=False)
            .agg({
                "Quantity": "sum",
                "Product Subtotal": "sum"
            })
        )

        grouped.rename(columns={
            "Quantity": "Total Quantity Sold",
            "Product Subtotal": "Total Revenue"
        }, inplace=True)

        return (
            grouped
            .sort_values("Total Revenue", ascending=True)
            .head(20)
            .reset_index(drop=True)
        )
