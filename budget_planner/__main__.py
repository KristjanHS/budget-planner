"""CLI entry point: python -m budget_planner [--update path]"""
import argparse
from budget_planner.workbook import generate_workbook, snapshot_workbook, restore_workbook


def main():
    parser = argparse.ArgumentParser(description="Generate programme budget workbook")
    parser.add_argument("--update", metavar="PATH",
                        help="Existing workbook to preserve data from")
    parser.add_argument("--skeleton", action="store_true",
                        help="Production skeleton (empty Capacity)")
    parser.add_argument("--snapshot", metavar="PATH",
                        help="Save snapshot for scenario comparison")
    parser.add_argument("--restore", metavar="PATH",
                        help="Restore from snapshot")
    parser.add_argument("-o", "--output", default="output/budget_plan.xlsx")
    args = parser.parse_args()

    if args.snapshot:
        snapshot_workbook(args.snapshot)
        return

    if args.restore:
        restore_workbook(args.restore)
        return

    mode = "skeleton" if args.skeleton else "demo"
    generate_workbook(output_path=args.output, existing_path=args.update, mode=mode)


if __name__ == "__main__":
    main()
