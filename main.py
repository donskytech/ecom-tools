import argparse
from services.order_service import OrderService
from services.income_service import IncomeService


def main():
    parser = argparse.ArgumentParser(
        description="ecom-tools | Order vs Income Reconciliation"
    )

    parser.add_argument("--orders", required=True, help="Orders Excel file")
    parser.add_argument("--income", required=True, help="Income Excel file")
    parser.add_argument("--export-missing", help="Export missing income report")

    args = parser.parse_args()

    print("\n📦 Loading orders...\n")
    order_service = OrderService(args.orders)

    order_summary = order_service.get_summary()
    print("📊 Order Summary")
    print(f"Total Orders     : {order_summary['total_orders']}")
    print(f"Completed Orders : {order_summary['completed_orders']}")

    income_service = IncomeService(args.income, order_service)
    recon = income_service.get_reconciliation_summary()

    print("\n💰 Income Reconciliation")
    print(f"Completed Orders        : {recon['completed_orders']}")
    print(f"Orders with Income      : {recon['orders_with_income']}")
    print(f"❌ Missing Income Orders: {recon['missing_income_orders']}")

    if recon["missing_income_orders"] > 0:
        print("\n🔍 Missing Income Order Details\n")
        report = income_service.get_missing_income_report()
        print(report.to_string(index=False))

        if args.export_missing:
            report.to_excel(args.export_missing, index=False)
            print(f"\n📤 Exported to {args.export_missing}")


if __name__ == "__main__":
    main()
