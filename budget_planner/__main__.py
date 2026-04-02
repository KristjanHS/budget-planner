"""CLI entry point: python -m budget_planner [--skeleton]"""
import argparse
from budget_planner.workbook import generate_workbook, snapshot_workbook, restore_workbook


def main():
    parser = argparse.ArgumentParser(description="Generate programme budget workbook")
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
    output_path = args.output
    if args.skeleton and output_path == "output/budget_plan.xlsx":
        output_path = "output/budget_plan_skeleton.xlsx"
    generate_workbook(output_path=output_path, mode=mode)


if __name__ == "__main__":
    main()
