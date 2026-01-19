from flask import Blueprint, render_template, request, send_file
from io import BytesIO

from services.order_service import OrderService
from services.income_service import IncomeService

main_bp = Blueprint("main", __name__)

# Simple in-memory cache for Excel download
_cached_excel = None


@main_bp.route("/", methods=["GET", "POST"])
def index():
    global _cached_excel

    if request.method == "POST":
        orders_file = request.files.get("orders_file")
        income_file = request.files.get("income_file")

        if not orders_file or not income_file:
            return render_template(
                "index.html",
                error="Please upload both Orders and Income files."
            )

        # --------------------------------------------------
        # Load uploaded files into memory
        # --------------------------------------------------
        orders_bytes = BytesIO(orders_file.read())
        income_bytes = BytesIO(income_file.read())

        # --------------------------------------------------
        # Initialize services
        # --------------------------------------------------
        order_service = OrderService(orders_bytes)
        income_service = IncomeService(income_bytes, order_service)

        # --------------------------------------------------
        # CORE SUMMARIES
        # --------------------------------------------------
        order_summary = order_service.get_summary()
        recon_summary = income_service.get_reconciliation_summary()

        # --------------------------------------------------
        # REPORTS
        # --------------------------------------------------
        missing_report_df = income_service.get_missing_income_report()
        income_summary = income_service.get_actual_received_income_summary()

        refund_summary = income_service.get_return_refund_summary()
        refund_details_df = income_service.get_return_refund_details()

        # Product analytics
        top_products = order_service.get_top_20_products_completed()
        least_products = order_service.get_top_20_least_products_completed()

        # 🚚 Overcharge Shipping Fee Summary
        overcharge_shipping_df = (
            income_service.get_overcharge_shipping_fee_summary()
        )

        # --------------------------------------------------
        # Cache Excel download
        # --------------------------------------------------
        _cached_excel = income_service.export_missing_orders_to_excel()

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
            can_download=len(missing_report_df) > 0
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
