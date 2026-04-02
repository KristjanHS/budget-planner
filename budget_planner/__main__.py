"""CLI entry point: python -m budget_planner [init|rebuild|--snapshot|--restore]"""
import argparse
import sys

from budget_planner.workbook import (
    init_workbook, rebuild_workbook, snapshot_workbook, restore_workbook,
)


def main():
    if len(sys.argv) > 1 and sys.argv[1] in ("init", "rebuild"):
        _subcommand_mode()
    else:
        _legacy_mode()


def _subcommand_mode():
    command = sys.argv[1]
    parser = argparse.ArgumentParser(
        prog=f"python -m budget_planner {command}",
    )

    if command == "init":
        parser.add_argument("--skeleton", action="store_true",
                            help="Empty Capacity for manual entry")
        parser.add_argument("-o", "--output",
                            default="output/budget_plan.xlsx")
        args = parser.parse_args(sys.argv[2:])
        mode = "skeleton" if args.skeleton else "demo"
        output = args.output
        if args.skeleton and output == "output/budget_plan.xlsx":
            output = "output/budget_plan_skeleton.xlsx"
        init_workbook(output_path=output, mode=mode)

    elif command == "rebuild":
        parser.add_argument("path", nargs="?",
                            default="output/budget_plan.xlsx",
                            help="Path to existing workbook")
        parser.add_argument("--skeleton", action="store_true",
                            help="Clear Capacity FTE values, preserve Actuals")
        args = parser.parse_args(sys.argv[2:])
        mode = "skeleton" if args.skeleton else "demo"
        rebuild_workbook(input_path=args.path, mode=mode)


def _legacy_mode():
    """Backward-compatible CLI: python -m budget_planner [--skeleton] [-o PATH]"""
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
    init_workbook(output_path=output_path, mode=mode)


if __name__ == "__main__":
    main()
