from flask import Blueprint, render_template, request, send_file
from io import BytesIO
import os

from services.order_service import OrderService
from services.income_service import IncomeService

main_bp = Blueprint("main", __name__)

# -----------------------------
# CONFIG
# -----------------------------
MAX_FILE_SIZE_MB = 10
ALLOWED_EXTENSIONS = {".xlsx", ".xls"}

# Simple in-memory cache for Excel download
_cached_excel = None


# -----------------------------
# HELPER FUNCTIONS
# -----------------------------
def _is_allowed_excel(file) -> bool:
    if not file or not file.filename:
        return False
    ext = os.path.splitext(file.filename)[1].lower()
    return ext in ALLOWED_EXTENSIONS


def _is_file_size_valid(file) -> bool:
    file.seek(0, os.SEEK_END)
    size_mb = file.tell() / (1024 * 1024)
    file.seek(0)
    return size_mb <= MAX_FILE_SIZE_MB


def _filename_starts_with(file, prefix: str) -> bool:
    if not file or not file.filename:
        return False
    return file.filename.lower().strip().startswith(prefix.lower())


# -----------------------------
# ROUTES
# -----------------------------
@main_bp.route("/", methods=["GET", "POST"])
def index():
    global _cached_excel

    if request.method == "POST":
        orders_file = request.files.get("orders_file")
        income_file = request.files.get("income_file")

        # --------------------------------------------------
        # BASIC PRESENCE CHECK
        # --------------------------------------------------
        if not orders_file or not income_file:
            return render_template(
                "index.html",
                error="❌ Please upload both Orders and Income Excel files."
            )

        # --------------------------------------------------
        # FILE TYPE VALIDATION
        # --------------------------------------------------
        if not _is_allowed_excel(orders_file):
            return render_template(
                "index.html",
                error="❌ Orders file must be an Excel file (.xlsx or .xls)."
            )

        if not _is_allowed_excel(income_file):
            return render_template(
                "index.html",
                error="❌ Income file must be an Excel file (.xlsx or .xls)."
            )

        # --------------------------------------------------
        # FILE SIZE VALIDATION
        # --------------------------------------------------
        if not _is_file_size_valid(orders_file):
            return render_template(
                "index.html",
                error="❌ Orders file exceeds the 10MB size limit."
            )

        if not _is_file_size_valid(income_file):
            return render_template(
                "index.html",
                error="❌ Income file exceeds the 10MB size limit."
            )

        # --------------------------------------------------
        # FILE NAME PREFIX VALIDATION
        # --------------------------------------------------
        if not _filename_starts_with(orders_file, "order"):
            return render_template(
                "index.html",
                error=(
                    "❌ Orders file name must start with the word "
                    "'order' (e.g. Order.all.20251201.xlsx)."
                )
            )

        if not _filename_starts_with(income_file, "income"):
            return render_template(
                "index.html",
                error=(
                    "❌ Income file name must start with the word "
                    "'income' (e.g. Income.released.ph.xlsx)."
                )
            )

        # --------------------------------------------------
        # LOAD FILES INTO MEMORY
        # --------------------------------------------------
        orders_bytes = BytesIO(orders_file.read())
        income_bytes = BytesIO(income_file.read())

        # --------------------------------------------------
        # INITIALIZE SERVICES
        # --------------------------------------------------
        order_service = OrderService(orders_bytes)
        income_service = IncomeService(income_bytes, order_service)

        # --------------------------------------------------
        # CORE SUMMARIES
        # --------------------------------------------------
        order_summary = order_service.get_summary()
        recon_summary = income_service.get_reconciliation_summary()

        # --------------------------------------------------
        # REPORTS & ANALYTICS
        # --------------------------------------------------
        missing_report_df = income_service.get_missing_income_report()
        income_summary = income_service.get_actual_received_income_summary()

        refund_summary = income_service.get_return_refund_summary()
        refund_details_df = income_service.get_return_refund_details()

        # Product analytics
        top_products = order_service.get_top_20_products_completed()
        least_products = order_service.get_top_20_least_products_completed()

        # Shipping overcharge analytics
        overcharge_shipping_df = (
            income_service.get_overcharge_shipping_fee_summary()
        )

        # --------------------------------------------------
        # CACHE EXCEL FOR DOWNLOAD
        # --------------------------------------------------
        _cached_excel = income_service.export_missing_orders_to_excel()
        shipping_overcharge = income_service.get_shipping_overcharge_status()
        region_summary = order_service.get_region_analysis_summary()

        return render_template(
            "results.html",
            order_summary=order_summary,
            recon_summary=recon_summary,
            missing_report=missing_report_df.to_dict(orient="records"),
            income_summary=income_summary,
            refund_summary=refund_summary,
            refund_details=refund_details_df.to_dict(orient="records"),
            top_products=top_products.to_dict(orient="records"),
            least_products=least_products.to_dict(orient="records"),
            overcharge_shipping=overcharge_shipping_df.to_dict(orient="records"),
            can_download=len(missing_report_df) > 0,
            shipping_overcharge=shipping_overcharge.to_dict(orient="records"),
            region_summary=region_summary,
        )

    # --------------------------------------------------
    # GET REQUEST
    # --------------------------------------------------
    return render_template("index.html")


@main_bp.route("/download")
def download_report():
    global _cached_excel

    if _cached_excel is None:
        return "No report available to download.", 400

    return send_file(
        _cached_excel,
        as_attachment=True,
        download_name="missing_income_orders.xlsx",
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )


@main_bp.route("/how-to")
def how_to():
    return render_template("how_to.html")


@main_bp.route("/about")
def about():
    return render_template("about.html")
