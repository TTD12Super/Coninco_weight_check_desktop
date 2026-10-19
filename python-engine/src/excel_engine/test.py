from pathlib import Path

from excel_engine.excel_reader import ExcelReader
from excel_engine.excel_converter import ExcelConverter
from excel_engine.product_formatter import ProductFormat, calculate_section

def main():
    file_path = Path(__file__).with_name("L01GF0087_value_only.xls")

    reader = ExcelReader(file_path)

    df = reader.read_sheet(
        sheet_name="02",
        skiprows=16,
        header=None
    )

    print("Original DataFrame:")
    print(df.head(10))

    df = ExcelConverter(df)

    df.remove_columns([0,3])


    df.rename_columns({
        1: "No.Stt",
        2: "description",
        4: "Unit",
        5: "W",
        6: "H",
        7: "W_prime",
        8: "H_prime",
        9: "W_double_prime",
        10: "H_double_prime",
        11: "L_M",
        12: "R_D",
        13: "a",
        14: "t",
        15: "Q",
        16: "K",
        17: "plan",
        18: "up_down",
        19: "vertical_axis",
        20: "S or Scn",
        21: "Sta",
        22: "insulation_quantity_detail",
        23: "acoustic_fireproof_detail",
        24: "Note"
    })

    df.add_k_by_function()
    df.set_column("section", df.get_data().apply(calculate_section, axis=1))

    df.remove_empty_rows([
        "W",
        "H",
        "W_prime",
        "H_prime",
        "W_double_prime",
        "H_double_prime",
        "L_M",
        "R_D",
    ])

    result = df.get_data()

    print("\nConverted DataFrame:")
    print(result.head(20))

    print("\nFormula results:")
    for _, row in result.iterrows():
        try:
            product = ProductFormat.from_row(row)
            calculation = product.calculate(row)
            print(
                product.section,
                product.product_type,
                calculation.to_dict(),
            )
        except (TypeError, ValueError) as error:
            print("Skipped:", error)
    # df.export_excel("output.xlsx")

if __name__ == "__main__":
    main()
