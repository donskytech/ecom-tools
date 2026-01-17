from flask import Blueprint, render_template, request, send_file
from io import BytesIO
from services.order_service import OrderService
from services.income_service import IncomeService

main_bp = Blueprint("main", __name__)

# In-memory cache (per request cycle)
_cached_report = None


@main_bp.route("/", methods=["GET", "POST"])
def index():
    global _cached_report

    if request.method == "POST":
        orders_file = request.files.get("orders_file")
        income_file = request.files.get("income_file")

        if not orders_file or not income_file:
            return render_template(
                "index.html",
                error="Please upload both files."
            )

        orders_bytes = BytesIO(orders_file.read())
        income_bytes = BytesIO(income_file.read())

        order_service = OrderService(orders_bytes)
        income_service = IncomeService(income_bytes, order_service)

        order_summary = order_service.get_summary()
        recon_summary = income_service.get_reconciliation_summary()
        missing_report = income_service.get_missing_income_report()

        # Cache report for download
        _cached_report = income_service.export_missing_orders_to_excel()

        return render_template(
            "results.html",
            order_summary=order_summary,
            recon_summary=recon_summary,
            missing_report=missing_report.to_dict(orient="records"),
            can_download=len(missing_report) > 0
        )

    return render_template("index.html")


@main_bp.route("/download")
def download_report():
    global _cached_report

    if _cached_report is None:
        return "No report available", 400

    return send_file(
        _cached_report,
        as_attachment=True,
        download_name="missing_income_orders.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
