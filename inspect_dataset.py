import pandas as pd

FILE = "organizer_dataset.xlsx"

excel = pd.ExcelFile(FILE)

print("Листы в файле:")
print(excel.sheet_names)

for sheet in excel.sheet_names:
    print("\n" + "=" * 50)
    print("Лист:", sheet)

    df = pd.read_excel(FILE, sheet_name=sheet, nrows=5)

    print("Размер первых строк:", df.shape)
    print("Колонки:")
    for i, col in enumerate(df.columns):
        print(f"{i}: {col}")

    print("\nПервые строки:")
    print(df.head())
