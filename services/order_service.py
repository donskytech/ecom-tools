import pandas as pd


class OrderService:
    def __init__(self, source):
        """
        source can be:
        - file path
        - file-like object (BytesIO)
        """
        self.source = source
        self.df = None

    def load_data(self):
        self.df = pd.read_excel(self.source)

        self.df.columns = (
            self.df.columns
            .astype(str)
            .str.strip()
        )

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
