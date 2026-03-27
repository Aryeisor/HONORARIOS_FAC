import argparse
from honorarios_app.core.processor import process_excel


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input")
    ap.add_argument("output")
    ap.add_argument("--especialidad", default="ortopedia")
    ap.add_argument("--original-sheet", default=None)
    ap.add_argument("--hosvi-sheet", default="HOSVIREPORT")
    ap.add_argument("--coosalud-sheet", default="COOSALUD")
    args = ap.parse_args()

    res = process_excel(
        input_path=args.input,
        output_path=args.output,
        especialidad=args.especialidad,
        original_sheet=args.original_sheet,
        hosvi_sheet=args.hosvi_sheet,
        coosalud_sheet=args.coosalud_sheet,
    )

    print("OK ->", res.output_path)
    print("PRE:", res.pre_rows)
    print("PEDIR:", res.pedir_rows)
    print("AMARILLO:", res.amarillo_rows)
    print("ANULADOS:", res.anulados_rows)


if __name__ == "__main__":
    main()