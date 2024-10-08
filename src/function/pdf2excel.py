"""demo read pdf
"""

import os
import sys
from typing import List, Optional
from copy import deepcopy

from typing import Tuple

ROOT_PATH = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
print(f"ROOT_PATH:{ROOT_PATH}")
sys.path.append(ROOT_PATH)

import json

import numpy as np
import pandas as pd
from pandas import DataFrame
import pymupdf
from pymupdf import Page, Rect, Document
from pymupdf.table import Table
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

from configs import ORG_PDF_PATH, BOOL_MERGE_HEAD_COL, BOOL_MERGE_HEAD_ROW


def get_table(doc: Document) -> Tuple[List[Table], List[DataFrame]]:
    """get target table in doc

    Args:
        doc (Document): doc object

    Returns:
        List[Table]: target table list
    """
    table_list = []
    table_df_list = []
    for page in doc:
        next_annot = page.first_annot
        while next_annot:
            new_next_annot = page.delete_annot(next_annot)
            next_annot = new_next_annot

        page_width = page.rect[2]
        table_width_percent_threshold = 0.8
        tables = page.find_tables()
        print("get tables")

        for table in tables.tables:
            table.width = abs(table.bbox[0] - table.bbox[2])

            if table.width > table_width_percent_threshold * page_width:
                print("find target table")
                table_df = table.to_pandas()
                table_list.append(table)
                table_df_list.append(table_df)

            for cell in table.cells:
                table.page.draw_rect(cell, color=(1, 0, 0))

    for table in table_list:
        table.page.draw_rect(table.bbox, color=(0, 1, 0))

    return table_list, table_df_list


def change_df2target(table_df: DataFrame) -> DataFrame:

    table_json = json.loads(table_df.to_json(orient="records"))
    print("get doc")

    new_table_json = []
    for table_json_idx, tmp_table_json in enumerate(table_json):
        if table_json_idx % 3 == 0:
            color_json = table_json[table_json_idx]

            if table_json_idx + 2 < len(table_json):
                size_json = table_json[table_json_idx + 1]
                qua_json = table_json[table_json_idx + 2]
                size_qua_pair_json = {}
                for key in size_json:
                    if size_json[key] and qua_json[key] != "Quantity":
                        tmp_size = size_json[key]
                        tmp_qua = qua_json[key]
                        print(f'color_json["Qty"]:{color_json["Qty"]}')
                        if int(color_json["Qty"]) == 0:
                            size_qua_pair_json[tmp_size] = int(tmp_qua)

                        else:
                            size_qua_pair_json[tmp_size] = int(tmp_qua) * int(
                                color_json["Qty"]
                            )
                        print(
                            f"size_qua_pair_json[tmp_size]:{size_qua_pair_json[tmp_size]}"
                        )

                tmp_table_json.update(size_qua_pair_json)
                new_table_json.append(tmp_table_json)
        else:
            pass

    df = pd.DataFrame(new_table_json)

    df = df.drop("Pack code", axis=1)
    df = df.drop("Col2", axis=1)
    df = df.drop("Barcode/SKU", axis=1)
    df = df.drop("Col4", axis=1)
    df = df.drop("Pack size", axis=1)
    df = df.drop("Unit price", axis=1)
    df = df.drop("Total (USD)", axis=1)
    df = df.drop("Qty", axis=1)
    # df.fillna(0, inplace=True)
    # 将值为0的元素替换为NaN
    df.replace(0, np.nan, inplace=True)
    df = df.rename(
        columns={"Colour": "颜色色号", "Pack/Loose": "包装方式", "Total": "合计"}
    )

    df = df.sort_index(axis=1)
    new_order = ["颜色色号"] + [col for col in df.columns if col != "颜色色号"]
    # new_order = ['颜色色号color'] + [col for col in df.columns if col != '颜色色号color']
    df = df[new_order]
    df["合计"] = df["合计"].str.replace(",", "").astype(int)

    df["包装方式"] = df["包装方式"].replace("Pack", "配比包装")
    df["包装方式"] = df["包装方式"].replace("Loose", "散装走货")

    return df


def func_pdf2excel(pdf_content):

    # # ORG_PDF_PATH = "D:/projects/pdf2excel_supplier_purchase/others/org_sample_file/2411082002 SDNZ-P0014217.pdf"
    # ORG_PDF_PATH = "D:/projects/pdf2excel_supplier_purchase/others/org_sample_file/2411082002 SDAU-P0014217.pdf"
    # # read pdf file
    # doc = pymupdf.open(ORG_PDF_PATH)
    doc = pymupdf.open(stream=pdf_content)
    table_list, table_df_list = get_table(doc=doc)

    out_path = "a.pdf"
    doc.save(out_path)
    # get table info

    if len(table_df_list) == 1:
        table_df = table_df_list[0]
    else:
        col = []
        table_df = None
        for tmp_table_id, tmp_table_df in enumerate(table_df_list):
            if tmp_table_id == 0:
                col = tmp_table_df.columns
                table_df = tmp_table_df
            else:
                fake_head = tmp_table_df.columns.to_list()
                new_fake_head = []
                for tmp_fake_head in fake_head:
                    if "-" in tmp_fake_head:
                        tmp_fake_head = tmp_fake_head.split("-")[-1]
                    elif "Col" in tmp_fake_head:
                        tmp_fake_head = pd.NA
                    new_fake_head.append(tmp_fake_head)
                fake_head_df = pd.DataFrame(new_fake_head, index=col)
                table_df = pd.concat([table_df, fake_head_df.T], ignore_index=True)
                tmp_table_df_reset = tmp_table_df.reset_index(drop=True)
                tmp_table_df_reset.columns = col
                table_df = pd.concat([table_df, tmp_table_df_reset], ignore_index=True)

    df = change_df2target(table_df)

    df.to_excel("c.xlsx", index=False)
    # 创建一个Workbook对象
    wb = Workbook()

    # 获取当前活跃的工作表
    ws = wb.active

    # 将DataFrame的数据写入工作表
    for r in dataframe_to_rows(df, index=False, header=True):
        print(f"r:{r}")
        ws.append(r)

    wb.save("back.xlsx")
    return wb


if __name__ == "__main__":

    # ORG_PDF_PATH = "D:/projects/pdf2excel_supplier_purchase/others/org_sample_file/2411082002 SDNZ-P0014217.pdf"
    # ORG_PDF_PATH = "D:/projects/pdf2excel_supplier_purchase/others/org_sample_file/2411082002 SDAU-P0014217.pdf"
    ORG_PDF_PATH = "D:/projects/pdf2excel_supplier_purchase/others/org_sample_file/2411082002 SDAU-P0014217_annot.pdf"
    # read pdf file
    doc = pymupdf.open(ORG_PDF_PATH)
    table_list, table_df_list = get_table(doc=doc)

    out_path = "a.pdf"
    doc.save(out_path)
    # get table info

    if len(table_df_list) == 1:
        table_df = table_df_list[0]
    else:
        col = []
        table_df = None
        for tmp_table_id, tmp_table_df in enumerate(table_df_list):
            if tmp_table_id == 0:
                col = tmp_table_df.columns
                table_df = tmp_table_df
            else:
                fake_head = tmp_table_df.columns.to_list()
                new_fake_head = []
                for tmp_fake_head in fake_head:
                    if "-" in tmp_fake_head:
                        tmp_fake_head = tmp_fake_head.split("-")[-1]
                    elif "Col" in tmp_fake_head:
                        tmp_fake_head = pd.NA
                    new_fake_head.append(tmp_fake_head)
                fake_head_df = pd.DataFrame(new_fake_head, index=col)
                table_df = pd.concat([table_df, fake_head_df.T], ignore_index=True)
                tmp_table_df_reset = tmp_table_df.reset_index(drop=True)
                tmp_table_df_reset.columns = col
                table_df = pd.concat([table_df, tmp_table_df_reset], ignore_index=True)

    df = change_df2target(table_df)

    df.to_excel("c.xlsx", index=False)
