import numpy as np
import pandas as pd
import re, os, argparse



def well_to_index(well):
    return ord(well[0]) - ord('A'), int(well[1:]) - 1


def convert(fp):
    # --- load file (xls + xlsx) ---
    if fp.endswith(".xlsx"):
        df = pd.read_excel(fp, header=None, engine="openpyxl")
    elif fp.endswith(".xls"):
        df = pd.read_excel(fp, header=None, engine="xlrd")
    else:
        raise ValueError("Unsupported file format")

    # --- output filename ---
    base = os.path.splitext(os.path.basename(fp))[0]
    out_fp = f"{base}_plates.xlsx"

    plates = {}
    label_num = None
    col_mappings = {}
    in_label_block = False

    for _, row in df.iterrows():

        first = str(row[0]).strip()

        # enter label block
        if "Lumin. Label" in first:
            match = re.search(r'Label\s*(\d+)', first)
            label_num = int(match.group(1)) if match else None
            col_mappings = {}
            in_label_block = True
            continue

        # exit label block
        if in_label_block and ("LB1 / LB2" in first or "LB2 / LB1" in first):
            in_label_block = False
            label_num = None
            col_mappings = {}
            continue

        if not in_label_block:
            continue

        # header
        if first == "Time [sec]":
            col_mappings = {}
            for j in range(len(row)):
                val = str(row[j]).strip()
                if re.match(r'^[A-H](0[1-9]|1[0-2])$', val):
                    col_mappings[j] = val
            continue

        # data row
        first_val = row[0]
        if not isinstance(first_val, (int, float)) or pd.isna(first_val):
            continue

        time_val = int(first_val)

        if label_num is None or not col_mappings:
            continue

        if time_val not in plates:
            plates[time_val] = {}

        if label_num not in plates[time_val]:
            plates[time_val][label_num] = np.zeros((8, 12))

        plate = plates[time_val][label_num]

        for j in range(len(row)):
            if j not in col_mappings:
                continue

            val = row[j]
            if isinstance(val, (int, float)) and not pd.isna(val):
                i, k = well_to_index(col_mappings[j])
                plate[i, k] = int(val)

    # --- write output Excel ---
    with pd.ExcelWriter(out_fp, engine="openpyxl") as writer:

        for time, labels in plates.items():
            sheet_name = f"time_{time}"
            start_row = 0

            for label in sorted(labels.keys()):
                grid = labels[label]

                df_out = pd.DataFrame(
                    grid,
                    index=[chr(ord('A') + i) for i in range(8)],
                    columns=[f"{j:02d}" for j in range(1, 13)]
                )

                # label title
                pd.DataFrame([[f"Label {label}"]]).to_excel(
                    writer,
                    sheet_name=sheet_name,
                    startrow=start_row,
                    index=False,
                    header=False
                )

                # grid
                df_out.to_excel(
                    writer,
                    sheet_name=sheet_name,
                    startrow=start_row + 1
                )

                start_row += len(df_out) + 3

    return out_fp


# ---------------- CLI ---------------- #

def main():

    parser = argparse.ArgumentParser(
        description="Convert plate reader Excel (.xls/.xlsx) into formatted plates.xlsx"
    )
    parser.add_argument("input", help="Path to input .xls or .xlsx file")
    parser.add_argument(
        "-o", "--output",
        help="Optional output file path (default: <input_basename>_plates.xlsx)"
    )
    args = parser.parse_args()
    out_fp = convert(args.input, args.output)
    print(f"Saved: {out_fp}")
if __name__ == "__main__":
    main()