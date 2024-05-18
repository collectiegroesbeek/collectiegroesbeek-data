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


if __name__ == "__main__":
    main()
