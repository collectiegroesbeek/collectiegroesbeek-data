import os

from xlsx2csv import Xlsx2csv


def main():
    for filename in sorted(os.listdir(".")):
        if filename.endswith(".xlsx"):
            print(f"Converting {filename} to csv")
            filename_csv = filename.replace(".xlsx", ".csv")
            Xlsx2csv(
                filename,
                outputencoding="utf-8",
                skip_empty_lines=True,
                lineterminator=os.linesep,
            ).convert(filename_csv)

    print("Merge 16. jaartallen into 3. jaartallen")
    with open('Coll Gr 3 Jaartallen.csv', 'a') as f3, open('Coll Gr 16 Jaartallen.csv') as f16:
        next(f16)  # Skip header
        f3.writelines(f16)

    os.remove("Coll Gr 16 Jaartallen.csv")


if __name__ == "__main__":
    main()
