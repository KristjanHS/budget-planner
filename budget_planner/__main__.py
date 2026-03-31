"""CLI entry point: python -m budget_planner [--update path]"""
import argparse
from budget_planner.workbook import generate_workbook

def main():
    parser = argparse.ArgumentParser(description="Generate programme budget workbook")
    parser.add_argument("--update", metavar="PATH", help="Existing workbook to preserve data from")
    parser.add_argument("-o", "--output", default="output/budget_plan.xlsx")
    args = parser.parse_args()
    generate_workbook(output_path=args.output, existing_path=args.update)

if __name__ == "__main__":
    main()
