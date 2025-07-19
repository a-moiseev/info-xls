import os
import re
import shutil
from pathlib import Path

import numpy as np
import pandas as pd
from Levenshtein import ratio
from openpyxl.reader.excel import load_workbook

from config import Config

SPECIALIZATION_MAP = {
    "МАНИКЮР": "МАНИКЮР-ПЕДИКЮР",
    "ПЕДИКЮР": "МАНИКЮР-ПЕДИКЮР",
    "ВИЗАЖ": "РЕСНИЦЫ, ВИЗАЖ",
    "РЕСНИЦЫ": "РЕСНИЦЫ, ВИЗАЖ",
    "МАССАЖ": "МАССАЖ лица",
}


class CalcZP:

    def __init__(self, config: Config):
        self.config = config

        self.zp_df = None
        self.files = None

        self.config.to_files_path.mkdir(parents=True, exist_ok=True)

    def get_zp_df(self) -> pd.DataFrame:
        def split_empls(empls_value: str) -> list[str]:
            empls_value = [empls_value]
            for key in ["\n", ",", "  "]:
                empls_value = [x for emp in empls_value for x in emp.split(key)]
            return empls_value

        df = pd.read_excel(self.config.info_path, header=1, sheet_name="расчет ЗП")
        df[["Правило", "Сотрудник"]] = df[["Правило", "Сотрудник"]].ffill()
        # new_df = pd.DataFrame(columns=df.columns)
        rows = []

        for _, row in df.iterrows():
            empls = split_empls(row["Сотрудник"])

            for emp in empls:
                if emp:
                    emp = emp.strip()
                    try:
                        spec = row["Специализация"].split("\n")
                    except AttributeError:
                        spec = [row["Специализация"]]
                    for s in spec:
                        if not s:
                            continue
                        s = str(s).strip()
                        new_row = row.copy()
                        new_row["Специализация"] = s
                        new_row["Сотрудник"] = emp
                        rows.append(new_row)
        return pd.DataFrame(rows, columns=df.columns)

    def get_files_df(self) -> list[Path]:
        return [
            file
            for file in self.config.from_files_path.iterdir()
            if file.suffix in [".xlsx", ".xls"]
        ]

    def calc_zp(self, fl: Path, zp_df: pd.DataFrame) -> None:
        def get_proc_to_zp_dict(zp_df: pd.DataFrame) -> dict:
            proc_to_zp_dict = dict(zip(zp_df["Специализация"], zp_df["Процент в ЗП"]))
            return proc_to_zp_dict

        def keep_only_digits(input_string: str) -> [int, None]:
            try:
                return int(re.sub(r"\D", "", str(input_string)))
            except ValueError:
                return

        def get_closest_match(procedure: str, proc_to_zp: dict) -> [str, None]:
            max_ratio = self.config.similarity_ratio
            max_proc = None
            for proc in proc_to_zp:
                r = ratio(procedure, proc)
                if r > max_ratio:
                    max_ratio = r
                    max_proc = proc
            return max_proc

        def get_match(first: str, second: str, similarity_ratio: float = 0.8) -> bool:
            if not first or not second:
                return False
            return ratio(first, second) > similarity_ratio

        def convert_xls_to_xlsx(file_path: Path) -> Path:
            if file_path.suffix.lower() != ".xls":
                return file_path

            import xlwings as xw

            xlsx_path = file_path.with_suffix(".xlsx")

            app = xw.App(visible=False)
            try:
                wb = app.books.open(str(file_path))
                wb.save(str(xlsx_path))
                wb.close()
                os.remove(file_path)
            finally:
                app.quit()

            return xlsx_path

        export_fl = self.config.to_files_path / fl.name
        suffix = export_fl.suffix.lower()
        if suffix == ".xls":
            shutil.copy(fl, export_fl)
            export_fl = convert_xls_to_xlsx(export_fl)
        elif suffix == ".xlsx":
            shutil.copy(fl, export_fl)
        else:
            print(f"File {fl} has unsupported format")
            return

        df = pd.read_excel(export_fl)
        df = df.replace(np.nan, None)
        employee = (export_fl.stem.split(" ")[0]).upper()

        zp_df = zp_df[
            zp_df["Сотрудник"].apply(
                lambda x: get_match(employee, x.split(" ")[0]),
                self.config.similarity_ratio,
            )
        ]
        if zp_df.empty:
            print(f"Employee {employee} not found in the ZP file")
            return

        proc_to_zp = get_proc_to_zp_dict(zp_df)

        zp_row = []
        for _, row in df.iterrows():
            procedure = row.iloc[0]
            quantity = row.iloc[1]
            all_price = row.iloc[6]
            if not procedure or not quantity:
                zp_row.append("")
                continue

            zp = ""

            match procedure:
                case "УСЛУГИ СОТРУДНИКАМ":
                    zp = all_price
                case "ТОВАРЫ НА ПРОДАЖУ":
                    zp = all_price * 0.05
                case "СТРИЖКИ" | "УКЛАДКИ":
                    zp = all_price * 0.5
                case "ОКРАШИВАНИЕ ВОЛОС" | "УХОДЫ ДЛЯ ВОЛОС":
                    zp = (all_price - all_price * 0.1) * 0.5
                case "РЕСНИЦЫ" | "ВИЗАЖ":
                    zp = all_price * 0.5
                case _:
                    mp = proc_to_zp.get(procedure)
                    if not mp:
                        try_procedure = SPECIALIZATION_MAP.get(procedure)
                        if try_procedure:
                            mp = proc_to_zp.get(try_procedure)
                    if not mp:
                        procedure = get_closest_match(procedure, proc_to_zp)
                        mp = proc_to_zp.get(procedure)
                    if isinstance(mp, float):
                        if mp > 100:
                            zp = mp * quantity
                        else:
                            zp = mp * all_price
                    elif isinstance(mp, str):
                        if mp := keep_only_digits(mp):
                            zp = mp * quantity

            zp_row.append(zp)

        # Load the existing workbook and sheet
        book = load_workbook(export_fl)
        sheet = book.active
        new_column_index = sheet.max_column + 1

        # Add the new column to the existing sheet
        for idx, value in enumerate(
            zp_row, start=2
        ):  # Assuming header is in the first row
            sheet.cell(row=idx, column=new_column_index, value=value)

        # Удаляем столбцы E, G и H
        cols_to_delete = ["H", "G", "E"]
        for col in cols_to_delete:
            sheet.delete_cols(sheet[col + "1"].column)
        new_column_index = new_column_index - 3

        sum_formula = f"=SUM({sheet.cell(row=2, column=new_column_index).coordinate}:{sheet.cell(row=len(zp_row), column=new_column_index).coordinate})"
        sheet.cell(row=len(zp_row) + 1, column=new_column_index, value=sum_formula)

        book.save(export_fl)

    def calculate(self):
        self.zp_df = self.get_zp_df()
        self.files = self.get_files_df()

        if not self.files:
            return

        for fl in self.files:
            print(f"Processing file: {fl.name}")
            self.calc_zp(fl, self.zp_df)
