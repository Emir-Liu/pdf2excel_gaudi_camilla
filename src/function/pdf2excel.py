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

            # for cell in table.cells:
            #     table.page.draw_rect(cell, color=(1, 0, 0))

    # for table in table_list:
    #     table.page.draw_rect(table.bbox, color=(0, 1, 0))

    return table_list, table_df_list


def change_df2target(
    table_df: DataFrame, order_number="", style_number="", order_type=""
) -> DataFrame:
    """Convert the original tabular data in PDF file to target format of tabular data

    Args:
        table_df (DataFrame): original tabular data

    Returns:
        DataFrame: target format of tabular data
    """

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
                proportion_pair_json = {}
                for key in size_json:
                    if size_json[key] and qua_json[key] != "Quantity":
                        tmp_size = size_json[key]
                        tmp_qua = qua_json[key]
                        print(f'color_json["Qty"]:{color_json["Qty"]}')
                        proportion_pair_json[tmp_size] = int(tmp_qua)
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
                if tmp_table_json["Pack/Loose"] == "Pack":
                    proportion_str = ":".join(
                        str(val) for val in proportion_pair_json.values()
                    )
                    # proportion_list = []
                    # for val in proportion_pair_json.values():
                    #     proportion_list
                    num_per_pack = sum(proportion_pair_json.values())

                    if num_per_pack != int(tmp_table_json["Pack size"]):
                        tmp_table_json["备注"] = (
                            f"配比{proportion_str}/ {tmp_table_json['Pack size']}件(注意：数量不一致)"
                        )
                    else:
                        tmp_table_json["备注"] = (
                            f"配比{proportion_str}/ {num_per_pack}件"
                        )
                else:
                    tmp_table_json["备注"] = "可溢2%"
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
    # df = df.drop("Qty", axis=1)
    # df.fillna(0, inplace=True)
    # 将值为0的元素替换为NaN
    df["Qty"] = df["Qty"].astype(int)
    df.replace(0, np.nan, inplace=True)
    df = df.rename(
        columns={
            "Colour": "颜色色号",
            "Pack/Loose": "包装方式",
            "Total": "合计",
            "Qty": "包数",
        }
    )

    df = df.sort_index(axis=1)
    # 对列进行排序 款号 颜色色号 PO号 订单形式 4 ... 18 合计 包装方式 备注 包数
    front_list = ["款号", "颜色色号", "PO号", "订单形式"]
    end_list = ["合计", "包装方式", "备注", "包数"]
    new_order = (
        front_list
        + [col for col in df.columns if col not in front_list + end_list]
        + end_list
    )

    df["款号"] = style_number
    df["PO号"] = order_number
    df["订单形式"] = order_type

    # new_order = ['颜色色号color'] + [col for col in df.columns if col != '颜色色号color']
    print(f"new_order:{new_order}")
    exist_index = df.columns
    print(f"exist_index:{exist_index}")
    for ordered_key in new_order:
        if ordered_key in exist_index:
            pass
        else:
            df[ordered_key] = pd.NA
    df = df[new_order]

    print(f"df:{df}")
    df["合计"] = df["合计"].str.replace(",", "").astype(int)

    df["包装方式"] = df["包装方式"].replace("Pack", "配比包装")
    df["包装方式"] = df["包装方式"].replace("Loose", "散装走货")

    return df


def func_pdf2excel(pdf_content):

    # # ORG_PDF_PATH = "D:/projects/pdf2excel_supplier_purchase/others/org_sample_file/2411082002 SDNZ-P0014217.pdf"
    # ORG_PDF_PATH = "D:/projects/pdf2excel_supplier_purchase/others/org_sample_file/2411082002 SDAU-P0014217.pdf"
    # # read pdf file
    # doc = pymupdf.open(ORG_PDF_PATH)

    # convert reading local file into reading data stream,
    # avoiding the need to save the file locally
    doc = pymupdf.open(stream=pdf_content)

    # remove annotation information from pdf files
    # to avoid the impact of annotation information on form extraction
    for page in doc:

        for annot in page.annots():
            page.delete_annot(annot=annot)

    # identify text in PDF file

    # identify text in PDF file of page 1
    page = doc[0]

    page_content = page.get_text(option="dict")

    print("get page content")
    content_list = []
    for block in page_content["blocks"]:
        tmp_block_content_list = []
        # tmp_block_content = ""
        # page.draw_rect(pymupdf.Rect(block["bbox"]))
        for line in block["lines"]:
            # page.draw_rect(pymupdf.Rect(line["bbox"]))
            for span in line["spans"]:
                # page.draw_rect(pymupdf.Rect(span["bbox"]))
                # tmp_block_content += span["text"]
                tmp_block_content_list.append(span["text"])
                # pass
        tmp_block_content = " ".join(tmp_block_content_list)
        print(f"tmp_block_content:{tmp_block_content}")
        # while "  " in tmp_block_content:
        #     tmp_block_content.replace("  ", " ")
        while True:
            if "  " in tmp_block_content:
                tmp_block_content = tmp_block_content.replace("  ", " ")
            else:
                break
        content_list.append(tmp_block_content)

    order_number = ""
    style_number = ""
    order_type = ""
    for tmp_content in content_list:
        if "Order number:" in tmp_content:
            order_number = tmp_content.split(" ")[-1]
        elif "Style number:" in tmp_content:
            style_number = tmp_content.split(" ")[-1]
        elif "SUPPLIER PURCHASE ORDER" in tmp_content:
            order_type = tmp_content.split("SUPPLIER PURCHASE ORDER")[0]

    print(
        f"order_number:{order_number}\nstyle_number:{style_number}\norder_type:{order_type}"
    )
    # doc.save("a.pdf")

    # get table in pdf file
    table_list, table_df_list = get_table(doc=doc)

    # out_path = "a.pdf"
    # doc.save(out_path)
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

    df = change_df2target(table_df, order_number, style_number, order_type)

    # df.to_excel("c.xlsx", index=False)
    # 创建一个Workbook对象
    wb = Workbook()

    # 获取当前活跃的工作表
    ws = wb.active

    # 将DataFrame的数据写入工作表
    for r in dataframe_to_rows(df, index=False, header=True):
        print(f"r:{r}")
        ws.append(r)

    # wb.save("back.xlsx")
    return wb


if __name__ == "__main__":

    ORG_PDF_PATH = "D:/projects/pdf2excel_supplier_purchase/others/org_sample_file/2411082002 SDNZ-P0014217.pdf"
    # ORG_PDF_PATH = "D:/projects/pdf2excel_supplier_purchase/others/org_sample_file/2411082002 SDAU-P0014217.pdf"
    # ORG_PDF_PATH = "D:/projects/pdf2excel_supplier_purchase/others/org_sample_file/2411082002 SDAU-P0014217_annot.pdf"
    # read pdf file
    doc = pymupdf.open(ORG_PDF_PATH)

    # remove annotation information from pdf files
    # to avoid the impact of annotation information on form extraction
    for page in doc:
        for annot in page.annots():
            page.delete_annot(annot=annot)

    # identify text in PDF file of page 1
    page = doc[0]

    page_content = page.get_text(option="dict")

    print("get page content")
    content_list = []
    for block in page_content["blocks"]:
        tmp_block_content_list = []
        # tmp_block_content = ""
        # page.draw_rect(pymupdf.Rect(block["bbox"]))
        for line in block["lines"]:
            # page.draw_rect(pymupdf.Rect(line["bbox"]))
            for span in line["spans"]:
                # page.draw_rect(pymupdf.Rect(span["bbox"]))
                # tmp_block_content += span["text"]
                tmp_block_content_list.append(span["text"])
                # pass
        tmp_block_content = " ".join(tmp_block_content_list)
        print(f"tmp_block_content:{tmp_block_content}")
        # while "  " in tmp_block_content:
        #     tmp_block_content.replace("  ", " ")
        while True:
            if "  " in tmp_block_content:
                tmp_block_content = tmp_block_content.replace("  ", " ")
            else:
                break
        content_list.append(tmp_block_content)

    order_number = ""
    style_number = ""
    order_type = ""
    for tmp_content in content_list:
        if "Order number:" in tmp_content:
            order_number = tmp_content.split(" ")[-1]
        elif "Style number:" in tmp_content:
            style_number = tmp_content.split(" ")[-1]
        elif "SUPPLIER PURCHASE ORDER" in tmp_content:
            order_type = tmp_content.split("SUPPLIER PURCHASE ORDER")[0]

    print(
        f"order_number:{order_number}\nstyle_number:{style_number}\norder_type:{order_type}"
    )

    for page in doc:

        page_content = page.get_text(option="dict")

        print("get page content")

        for block in page_content["blocks"]:
            # page.draw_rect(pymupdf.Rect(block["bbox"]))
            for line in block["lines"]:
                # page.draw_rect(pymupdf.Rect(line["bbox"]))
                for span in line["spans"]:
                    # page.draw_rect(pymupdf.Rect(span["bbox"]))
                    pass
    out_path = "block.pdf"
    doc.save(out_path)

    table_list, table_df_list = get_table(doc=doc)

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

    df = change_df2target(table_df, order_number, style_number, order_type)

    df.to_excel("c.xlsx", index=False)
