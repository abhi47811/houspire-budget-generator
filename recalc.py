import sys, json, openpyxl

def verify_boq(path, max_rows=500):
    wb = openpyxl.load_workbook(path)   # NOT data_only=True — that strips formulas
    ws = wb.active
    errors = []
    for r in range(2, min(ws.max_row + 1, max_rows + 2)):
        qty = ws.cell(r, 4).value
        if qty is None:
            continue
        expected = f"=D{r}*E{r}"
        found = ws.cell(r, 6).value
        if found != expected:
            errors.append({"row": r, "expected": expected, "found": str(found)})
    result = {"total_errors": len(errors), "errors": errors}
    print(json.dumps(result, indent=2))
    return result

if __name__ == "__main__":
    r = verify_boq(sys.argv[1], int(sys.argv[2]) if len(sys.argv) > 2 else 500)
    sys.exit(0 if r["total_errors"] == 0 else 1)
