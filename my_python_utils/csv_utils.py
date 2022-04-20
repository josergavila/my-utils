import csv
from datetime import date

__all__ = "save_to_csv"

DATE = date.today().strftime("%d-%m-%y")


def save_to_csv(data, filename: str, columns: list):
    filename = f"{filename}_{DATE}.csv"
    with open(filename, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(columns)
        writer.writerows(data)
