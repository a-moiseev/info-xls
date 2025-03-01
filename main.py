import io
import os
import re
import shutil
import yaml

import numpy as np
import pandas as pd
import msoffcrypto
from pathlib import Path
from Levenshtein import ratio

from openpyxl.reader.excel import load_workbook


def get_config() -> dict:
    with open("config.yaml", "r") as file:
        config = yaml.safe_load(file)
    return config


config = get_config()

CURRENT_PATH = Path(os.getcwd())
INFO_PATH = CURRENT_PATH / "zp_file" / "ЗАРАБОТНАЯ ПЛАТА  2025.xlsx"
FROM_FILES_PATH = CURRENT_PATH / "files"
TO_FILES_PATH = CURRENT_PATH / "files_new"
PASSWD = str(config.get("password"))
SIMILARITY_RATIO = 0.8

SPECIALIZATION_MAP = {
    "МАНИКЮР": "МАНИКЮР-ПЕДИКЮР",
    "ПЕДИКЮР": "МАНИКЮР-ПЕДИКЮР",
    "ВИЗАЖ": "РЕСНИЦЫ, ВИЗАЖ",
    "РЕСНИЦЫ": "РЕСНИЦЫ, ВИЗАЖ",
    "МАССАЖ": "МАССАЖ лица",
}

decrypted_workbook = io.BytesIO()
with open(INFO_PATH, "rb") as file:
    office_file = msoffcrypto.OfficeFile(file)
    office_file.load_key(password=PASSWD)
    office_file.decrypt(decrypted_workbook)


def get_zp_df():
    df = pd.read_excel(decrypted_workbook, header=1, sheet_name="расчет ЗП")
    df[["Правило", "Сотрудник"]] = df[["Правило", "Сотрудник"]].ffill()
    new_df = pd.DataFrame(columns=df.columns)

    for _, row in df.iterrows():
        empls = row["Сотрудник"].split("\n")
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
                    row["Специализация"] = s
                    row["Сотрудник"] = emp
                    new_df.loc[len(new_df)] = row
    # new_df["Сотрудник"] = new_df["Сотрудник"].apply(
    #     lambda x: str(x).split(" ")[0] if x else ""
    # )
    return new_df


def get_files_df():
    return [
        file for file in FROM_FILES_PATH.iterdir() if file.suffix in [".xlsx", ".xls"]
    ]


def calc_zp(fl, zp_df):
    def get_proc_to_zp_dict(zp_df):
        proc_to_zp_dict = dict(zip(zp_df["Специализация"], zp_df["Процент в ЗП"]))
        return proc_to_zp_dict

    def keep_only_digits(input_string):
        try:
            return int(re.sub(r"\D", "", str(input_string)))
        except ValueError:
            return

    def get_closest_match(procedure, proc_to_zp):
        max_ratio = SIMILARITY_RATIO
        max_proc = None
        for proc in proc_to_zp:
            r = ratio(procedure, proc)
            if r > max_ratio:
                max_ratio = r
                max_proc = proc
                print(
                    f"Max ratio: {max_ratio}, proc: {procedure}, max proc: {max_proc}"
                )
        return max_proc

    def get_match():
        pass

    export_fl = TO_FILES_PATH / fl.name
    shutil.copy(fl, export_fl)

    df = pd.read_excel(export_fl)
    df = df.replace(np.nan, None)
    employee = (export_fl.stem.split(" ")[-1]).upper()

    # zp_df = zp_df[zp_df["Сотрудник"].apply(lambda x: employee in x.upper())]
    zp_df = zp_df[zp_df["Сотрудник"].str.upper().str.contains(employee.upper(), regex=False)]

    # if zp_df.empty:
    #     zp_df = zp_df[
    #         zp_df["Сотрудник"].apply(lambda x: employee == str(x).split(" ")[0])
    #     ]
    # zp_df = zp_df[zp_df["Сотрудник"] == employee]

    if zp_df.empty:
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
        mp = proc_to_zp.get(procedure)
        if not mp:
            try_procedure = SPECIALIZATION_MAP.get(procedure)
            if try_procedure:
                mp = proc_to_zp.get(try_procedure)
        if not mp:
            procedure = get_closest_match(procedure, proc_to_zp)
            mp = proc_to_zp.get(procedure)

        if procedure == "УСЛУГИ СОТРУДНИКАМ":
            zp = all_price
        elif isinstance(mp, float):
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
    for idx, value in enumerate(zp_row, start=2):  # Assuming header is in the first row
        sheet.cell(row=idx, column=new_column_index, value=value)

    # Save the workbook
    book.save(export_fl)


if __name__ == "__main__":
    zp_df = get_zp_df()

    TO_FILES_PATH.mkdir(parents=True, exist_ok=True)

    files = get_files_df()
    for fl in files:
        calc_zp(fl, zp_df)
