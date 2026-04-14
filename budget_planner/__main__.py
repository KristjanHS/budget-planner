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
        # Default is None so we can distinguish "user overrode" from
        # "auto-pick based on --demo" — demo runs land at a _demo-suffixed
        # path by default so they don't clobber the real workbook.
        parser.add_argument("-o", "--output", default=None)
        args = parser.parse_args(sys.argv[2:])
        if args.output is not None:
            output = args.output
        elif args.demo:
            output = "output/budget_plan_demo.xlsx"
        else:
            output = "output/budget_plan.xlsx"
        if args.demo:
            from budget_planner.config import DEMO_CAPACITY_ASSIGNMENTS
            init_workbook(output_path=output,
                          assignments=DEMO_CAPACITY_ASSIGNMENTS)
            seed_demo_data(output)
        else:
            init_workbook(output_path=output)

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
