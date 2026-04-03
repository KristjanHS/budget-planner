"""CLI entry point: python -m budget_planner [init|rebuild|--snapshot|--restore]"""
import argparse
import sys

from budget_planner.workbook import (
    init_workbook, rebuild_workbook, snapshot_workbook, restore_workbook,
    seed_demo_data,
)


def main():
    if len(sys.argv) > 1 and sys.argv[1] in ("init", "rebuild"):
        _subcommand_mode()
    else:
        _snapshot_restore_mode()


def _subcommand_mode():
    command = sys.argv[1]
    parser = argparse.ArgumentParser(
        prog=f"python -m budget_planner {command}",
    )

    if command == "init":
        parser.add_argument("--demo", action="store_true",
                            help="Seed Capacity + Actuals with demo data")
        parser.add_argument("-o", "--output",
                            default="output/budget_plan.xlsx")
        args = parser.parse_args(sys.argv[2:])
        output = args.output
        init_workbook(output_path=output)
        if args.demo:
            seed_demo_data(output)

    elif command == "rebuild":
        parser.add_argument("path", nargs="?",
                            default="output/budget_plan.xlsx",
                            help="Path to existing workbook")
        args = parser.parse_args(sys.argv[2:])
        rebuild_workbook(input_path=args.path)


def _snapshot_restore_mode():
    """CLI for snapshot/restore: python -m budget_planner [--snapshot|--restore] PATH"""
    parser = argparse.ArgumentParser(description="programme budget workbook utilities")
    parser.add_argument("--snapshot", metavar="PATH",
                        help="Save snapshot for scenario comparison")
    parser.add_argument("--restore", metavar="PATH",
                        help="Restore from snapshot")
    args = parser.parse_args()

    if args.snapshot:
        snapshot_workbook(args.snapshot)
        return

    if args.restore:
        restore_workbook(args.restore)
        return

    parser.print_help()


if __name__ == "__main__":
    main()
