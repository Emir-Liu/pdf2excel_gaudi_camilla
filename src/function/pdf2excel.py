"""demo read pdf
"""

import os
import sys
from typing import List, Optional
from copy import deepcopy

ROOT_PATH = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
print(f"ROOT_PATH:{ROOT_PATH}")
sys.path.append(ROOT_PATH)


import pymupdf
from pymupdf import Page, Rect, Document
from openpyxl import Workbook

from configs import ORG_PDF_PATH, BOOL_MERGE_HEAD_COL, BOOL_MERGE_HEAD_ROW


def get_annot_info(path: str = None, doc: Optional[Document] = None) -> List[dict]:
    """get annot in pdf file

    Args:
        path (str): pdf file path

    Returns:
        List[dict]: annot info
        [
            {
                'page':Page,
                'rect':Rect
            },
            ...
        ]
    """
    if doc:
        pass
    else:
        doc = pymupdf.open(ORG_PDF_PATH)

    annot_info = []
    for page in doc:
        print(f"page num:{page.number}")
        # determine if get annots
        annots = page.annots()

        for tmp_annot_idx, tmp_annot in enumerate(annots):
            print(f"find annot{tmp_annot_idx} in page {page.number}:{tmp_annot.rect}")
            tmp_annot_info = {"page": page, "rect": tmp_annot.rect}
            annot_info.append(tmp_annot_info)

    return annot_info


def str_is_only_digits_and_spaces(content: str) -> bool:
    """determine if the string only contains digits and space

    Args:
        content (str): string

    Returns:
        bool: if only contains digits and space return true
    """
    no_space = content.replace(" ", "")
    return no_space.isnumeric()


def get_table_info_in_annot_rect(page: Page, rect: Rect) -> List[dict]:
    """get table content info in annot rect

    Args:
        page (Page): page object
        rect (Rect): annot position

    Returns:
        List[dict]:
        [
            {
                'content':'',
                'rect':(x,x,x,x),
                'center': (x,x),
            },
            ...
        ]
    """
    table_info_in_annot = []
    text_info = page.get_text(
        option="dict",
        clip=rect,
    )
    for tmp_text_block in text_info["blocks"]:
        for tmp_text_line in tmp_text_block["lines"]:
            for tmp_text_span in tmp_text_line["spans"]:
                # print(f"tmp text block content:{tmp_text_span['text']}")
                # page.draw_rect(tmp_text_span["bbox"], color=(1, 0, 0))

                # if all numbers and space
                if (
                    str_is_only_digits_and_spaces(tmp_text_span["text"])
                    and " " in tmp_text_span["text"].strip()
                ):
                    content_list = tmp_text_span["text"].split(" ")
                    num_split = len(content_list)
                    width = tmp_text_span["bbox"][2] - tmp_text_span["bbox"][0]
                    part_wid = width / num_split
                    for tmp_content_id, tmp_content in enumerate(content_list):
                        tmp_table_info_in_annot = {
                            "content": tmp_content,
                            "rect": (
                                tmp_text_span["bbox"][0] + tmp_content_id * part_wid,
                                tmp_text_span["bbox"][1],
                                tmp_text_span["bbox"][0]
                                + (tmp_content_id + 1) * part_wid,
                                tmp_text_span["bbox"][3],
                            ),
                            "center": (
                                (
                                    tmp_text_span["bbox"][0]
                                    + tmp_content_id * part_wid
                                    + tmp_text_span["bbox"][0]
                                    + (tmp_content_id + 1) * part_wid
                                )
                                / 2,
                                (tmp_text_span["bbox"][1] + tmp_text_span["bbox"][3])
                                / 2,
                            ),
                        }
                        table_info_in_annot.append(tmp_table_info_in_annot)
                else:
                    tmp_table_info_in_annot = {
                        "content": tmp_text_span["text"],
                        "rect": tmp_text_span["bbox"],
                        "center": (
                            (tmp_text_span["bbox"][0] + tmp_text_span["bbox"][2]) / 2,
                            (tmp_text_span["bbox"][1] + tmp_text_span["bbox"][3]) / 2,
                        ),
                    }
                    table_info_in_annot.append(tmp_table_info_in_annot)
    return table_info_in_annot


def build_row_info(table_info_in_annot: List[dict]) -> List[dict]:
    """get row info by table info

    Args:
        table_info_in_annot(List[dict]):
        [
            {
                'content':'',
                'rect':(x,x,x,x),
                'center': (x,x),
            },
            ...
        ]

    Returns:
        List[dict]: row info
        [
            {
                'max_y':xx,
                'min_y':xx,
                'center_y':xx,
                'content_id':[
                    xx,
                    xx,
                    ..
                ]
            },
            ...
        ]
    """
    row_info = []

    for tmp_table_info_in_annot_idx, tmp_table_info_in_annot in enumerate(
        table_info_in_annot
    ):
        max_y = max(
            tmp_table_info_in_annot["rect"][1], tmp_table_info_in_annot["rect"][3]
        )
        min_y = min(
            tmp_table_info_in_annot["rect"][1], tmp_table_info_in_annot["rect"][3]
        )
        center_y = (max_y + min_y) / 2

        bool_find_row = False
        for tmp_row_info in row_info:
            if center_y < tmp_row_info["max_y"] and center_y > tmp_row_info["min_y"]:
                tmp_row_info["content_id"].append(tmp_table_info_in_annot_idx)
                if max_y > tmp_row_info["max_y"]:
                    tmp_row_info["max_y"] = max_y
                if min_y < tmp_row_info["min_y"]:
                    tmp_row_info["min_y"] = min_y
                tmp_row_info["center_y"] = (
                    tmp_row_info["min_y"] + tmp_row_info["max_y"]
                ) / 2
                bool_find_row = True
                break
            else:
                pass

        if bool_find_row is True:
            pass
        else:
            row_info.append(
                {
                    "max_y": max_y,
                    "min_y": min_y,
                    "center_y": (max_y + min_y) / 2,
                    "content_id": [tmp_table_info_in_annot_idx],
                }
            )

    for tmp_row_info in row_info:
        tmp_row_info["content_id"].sort(
            key=lambda x: table_info_in_annot[x]["center"][0],
            reverse=False,
        )

    row_info.sort(key=lambda x: x["center_y"], reverse=False)

    # merge head row
    # attention
    if BOOL_MERGE_HEAD_ROW is True:
        head_couple_row = [
            (
                "Direct",
                "Disc.",
            ),
            (
                "Unit Cost",
                "%",
            ),
        ]
        for row_id in [0, 1]:
            new_content_id = []
            for tmp_row_info_content_id_id in range(
                len(row_info[row_id]["content_id"]) - 1
            ):
                # for tmp_row_content_id_new in row_info[0]["content_id"]:
                #     if tmp_row_content_id_new == tmp_row_content_id:
                #         continue
                if tmp_row_info_content_id_id == 0:
                    new_content_id.append(
                        row_info[row_id]["content_id"][tmp_row_info_content_id_id]
                    )
                tmp_table_info = table_info_in_annot[
                    row_info[row_id]["content_id"][tmp_row_info_content_id_id]
                ]
                tmp_table_info_new = table_info_in_annot[
                    row_info[row_id]["content_id"][tmp_row_info_content_id_id + 1]
                ]

                bool_find_row = False
                for head_couple_row_idx in range(len(head_couple_row)):
                    if (
                        tmp_table_info["content"].strip()
                        == head_couple_row[head_couple_row_idx][0]
                        and tmp_table_info_new["content"].strip()
                        == head_couple_row[head_couple_row_idx][1]
                    ):

                        tmp_row_info_content_id_id += 1
                        tmp_table_info["content"] = (
                            tmp_table_info["content"]
                            + " "
                            + tmp_table_info_new["content"]
                        )
                        tmp_table_info["rect"] = (
                            min(tmp_table_info["rect"][0], tmp_table_info["rect"][0]),
                            tmp_table_info["rect"][1],
                            max(tmp_table_info["rect"][2], tmp_table_info["rect"][2]),
                            tmp_table_info["rect"][3],
                        )
                        bool_find_row = True
                        break
                if bool_find_row is False:
                    new_content_id.append(
                        row_info[row_id]["content_id"][tmp_row_info_content_id_id + 1]
                    )
            row_info[row_id]["content_id"] = new_content_id

    # print("finish merge head row")

    if BOOL_MERGE_HEAD_COL is True:
        # merge head col
        head_couple = [
            ("Location", "Code"),
            ("Unit of", "Measure"),
            ("Direct Disc.", "Unit Cost %"),
        ]
        for tmp_row_content_id_row0 in row_info[0]["content_id"]:
            bool_find_couple = False
            tmp_table_info = table_info_in_annot[tmp_row_content_id_row0]
            tmp_table_info_content = tmp_table_info["content"]
            # if tmp_table_info_content
            for tmp_row_content_id_row1 in row_info[1]["content_id"]:
                tmp_table_info_row1 = table_info_in_annot[tmp_row_content_id_row1]
                tmp_table_info_content_row1 = tmp_table_info_row1["content"]

                for tmp_couple in head_couple:
                    if (
                        tmp_table_info_content.strip() == tmp_couple[0]
                        and tmp_table_info_content_row1.strip() == tmp_couple[1]
                    ):
                        tmp_table_info_row1["content"] = (
                            tmp_table_info_content + "\n" + tmp_table_info_content_row1
                        )
                        tmp_table_info_row1["rect"] = (
                            min(
                                tmp_table_info["rect"][0],
                                tmp_table_info_row1["rect"][0],
                            ),
                            tmp_table_info_row1["rect"][1],
                            max(
                                tmp_table_info["rect"][2],
                                tmp_table_info_row1["rect"][2],
                            ),
                            tmp_table_info_row1["rect"][3],
                        )
                        bool_find_couple = True
                        break
                if bool_find_couple == True:
                    break

        row_info.remove(row_info[0])
    # print("finish merge head col")
    # for tmp_sorted_row_info in row_info:
    #     print(f"row:{tmp_sorted_row_info}\n")

    #     for content_id in tmp_sorted_row_info["content_id"]:
    #         print(f"content:{table_info_in_annot[content_id]}\n")

    return row_info


def build_col_info(row_info: List[dict], table_info_in_annot: List[dict]) -> List[dict]:
    """get row info by table info

    Args:
        row_info(List[dict]): row info list
        [
            {
                'max_y':xx,
                'min_y':xx,
                'center_y':xx,
                'content_id':[
                    xx,
                    xx,
                    ..
                ]
            },
            ...
        ]

        table_info_in_annot(List[dict]): table info in annot
        [
            {
                'content':'',
                'rect':(x,x,x,x),
                'center': (x,x),
            },
            ...
        ]

    Returns:
        List[dict]: col info
        [
            {
                'max_x':xx,
                'min_x':xx,
                'center_x':xx,
                'content_id':[
                    xx,
                    xx,
                    ..
                ]
            },
            ...
        ]
    """
    col_info = []

    # head info
    for tmp_row_content_id_id in range(len(row_info[0]["content_id"]) - 1):
        tmp_table_info_in_annot_id = row_info[0]["content_id"][tmp_row_content_id_id]
        tmp_table_info_in_annot = table_info_in_annot[tmp_table_info_in_annot_id]

        if tmp_table_info_in_annot["content"] == "Location\nCode":
            left_col = tmp_table_info_in_annot["rect"][0]
            right_col = tmp_table_info_in_annot["rect"][2]
        else:
            next_table_info_in_annot_id = row_info[0]["content_id"][
                tmp_row_content_id_id + 1
            ]
            next_table_info_in_annot = table_info_in_annot[next_table_info_in_annot_id]
            left_col = tmp_table_info_in_annot["rect"][0]
            right_col = next_table_info_in_annot["rect"][0]
        tmp_col_info = {
            "min_x": left_col,
            "max_x": right_col,
            "center_x": (left_col + right_col) / 2,
            "content_id": [tmp_table_info_in_annot_id],
        }
        col_info.append(tmp_col_info)

    # final col in head
    tmp_table_info_in_annot_id = row_info[0]["content_id"][-1]
    tmp_table_info_in_annot = table_info_in_annot[tmp_table_info_in_annot_id]
    left_col = tmp_table_info_in_annot["rect"][0]
    right_col = 100000
    tmp_col_info = {
        "min_x": left_col,
        "max_x": right_col,
        "center_x": right_col,
        "content_id": [tmp_table_info_in_annot_id],
    }
    col_info.append(tmp_col_info)

    # body info
    for tmp_table_info_in_annot_idx, tmp_table_info_in_annot in enumerate(
        table_info_in_annot
    ):
        # print(f"当前数据:{tmp_table_info_in_annot}")
        min_x = min(
            tmp_table_info_in_annot["rect"][0], tmp_table_info_in_annot["rect"][2]
        )
        max_x = max(
            tmp_table_info_in_annot["rect"][0], tmp_table_info_in_annot["rect"][2]
        )
        center_x = (min_x + max_x) / 2

        bool_find_col = False
        # print(f"待对比数据:")
        # for tmp_col_info in col_info:
        #     print(f"{tmp_col_info}")
        for tmp_col_info_id, tmp_col_info in enumerate(col_info):
            # print(f"当前对比列数据:{tmp_col_info}")
            if center_x < tmp_col_info["max_x"] and center_x > tmp_col_info["min_x"]:
                # print(f"找到旧列")
                if tmp_table_info_in_annot_idx in tmp_col_info["content_id"]:
                    # print(f"已存在")
                    pass
                else:
                    # print(f"未存在")
                    if min_x < tmp_col_info["min_x"]:
                        tmp_col_info["min_x"] = min_x

                    if max_x > tmp_col_info["max_x"]:
                        tmp_col_info["max_x"] = max_x

                    tmp_col_info["center_x"] = (
                        tmp_col_info["min_x"] + tmp_col_info["max_x"]
                    ) / 2

                    tmp_col_info["content_id"].append(tmp_table_info_in_annot_idx)
                bool_find_col = True
                break

        if bool_find_col is False:
            tmp_col_info = {
                "min_x": min_x,
                "max_x": max_x,
                "center_x": (min_x + max_x) / 2,
                "content_id": [tmp_table_info_in_annot_idx],
            }
            # print(f"创建新列:{tmp_col_info}")
            col_info.append(tmp_col_info)
    col_info.sort(key=lambda x: x["center_x"], reverse=False)
    # for tmp_sorted_col_info in col_info:
    #     print(f"col:{tmp_sorted_col_info}\n")

    #     for content_id in tmp_sorted_col_info["content_id"]:
    #         print(f"content:{table_info_in_annot[content_id]}\n")
    return col_info


def build_excel_info(
    row_info: List[dict], col_info: List[dict], table_info_in_annot: List[dict]
) -> dict:
    """get excel info

    Args:
        row_info (List[dict]): row info
        [
            {
                'max_y':xx,
                'min_y':xx,
                'center_y':xx,
                'content_id':[
                    xx,
                    xx,
                    ..
                ]
            },
            ...
        ]
        col_info (List[dict]): col info
        [
            {
                'max_y':xx,
                'min_y':xx,
                'center_y':xx,
                'content_id':[
                    xx,
                    xx,
                    ..
                ]
            },
            ...
        ]
        table_info_in_annot(List[dict]): table info in annot
        [
            {
                'content':'',
                'rect':(x,x,x,x),
                'center': (x,x),
            },
            ...
        ]
    Returns:
        {
            0:{
                'row':xx,
                'col':xx,
                'content':'xx',
                'rect':(x,x,x,x),
                'center':(x,x),
            }
            ...
        }
    """
    excel_table_info = {}

    for row_idx, tmp_row_info in enumerate(row_info):

        for tmp_table_info_in_annot_id in tmp_row_info["content_id"]:
            tmp_excel_table_info = {
                "row": row_idx,
                "col": -1,
            }
            tmp_table_info_in_annot = table_info_in_annot[tmp_table_info_in_annot_id]
            tmp_excel_table_info.update(table_info_in_annot[tmp_table_info_in_annot_id])

            # excel_table_info.append(tmp_excel_table_info)
            excel_table_info[tmp_table_info_in_annot_id] = tmp_excel_table_info

    for col_idx, tmp_col_info in enumerate(col_info):
        for tmp_table_in_annot_id in tmp_col_info["content_id"]:
            try:
                excel_table_info[tmp_table_in_annot_id]["col"] = col_idx
            except:
                pass

    # print(f"excel结构数据")
    # for tmp_excel_table_id in excel_table_info:
    #     print(f"{excel_table_info[tmp_excel_table_id]}")

    return excel_table_info


def func_pdf2excel(pdf_content=ORG_PDF_PATH):
    # read pdf file
    doc = pymupdf.open(stream=pdf_content)
    # doc = pymupdf.open(pdf_content)

    # read annot info
    annot_info = get_annot_info(doc=doc)
    # print("get annot info")

    excel_table_info_total = []
    for tmp_annot_info_idx, tmp_annot_info in enumerate(annot_info):
        # choose one annot info for example
        # tmp_annot_info = annot_info[0]
        page = tmp_annot_info["page"]
        rect = tmp_annot_info["rect"]
        table_info_in_annot = get_table_info_in_annot_rect(page=page, rect=rect)

        # draw info
        for tmp_table_info_in_annot in table_info_in_annot:
            page.draw_rect(tmp_table_info_in_annot["rect"])

        # build row info
        row_info = build_row_info(table_info_in_annot=table_info_in_annot)

        # build col info
        col_info = build_col_info(row_info, table_info_in_annot)

        # build excel info
        excel_table_info = build_excel_info(row_info, col_info, table_info_in_annot)

        excel_table_info_total.append(deepcopy(excel_table_info))

    # write excel file
    wb = Workbook()

    for tmp_excel_table_info_id, tmp_excel_table_info in enumerate(
        excel_table_info_total
    ):
        if tmp_excel_table_info_id == 0:
            ws = wb.active
            ws.title = f"sheet_{tmp_excel_table_info_id}"
        else:
            ws = wb.create_sheet(title=f"sheet_{tmp_excel_table_info_id}")

        for tmp_key, tmp_item in tmp_excel_table_info.items():
            # row and col must be at least 1
            row = tmp_item["row"] + 1
            col = tmp_item["col"] + 1
            content = tmp_item["content"]
            ws.cell(row=row, column=col, value=content)
    # wb.save("a.xlsx")
    return wb


if __name__ == "__main__":
    # read pdf file
    doc = pymupdf.open(ORG_PDF_PATH)

    # read annot info
    annot_info = get_annot_info(doc=doc)
    # print("get annot info")

    excel_table_info_total = []
    for tmp_annot_info_idx, tmp_annot_info in enumerate(annot_info):
        # choose one annot info for example
        # tmp_annot_info = annot_info[0]
        page = tmp_annot_info["page"]
        rect = tmp_annot_info["rect"]
        table_info_in_annot = get_table_info_in_annot_rect(page=page, rect=rect)

        # draw info
        for tmp_table_info_in_annot in table_info_in_annot:
            page.draw_rect(tmp_table_info_in_annot["rect"])

        # build row info
        row_info = build_row_info(table_info_in_annot=table_info_in_annot)

        # build col info
        col_info = build_col_info(row_info, table_info_in_annot)

        # build excel info
        excel_table_info = build_excel_info(row_info, col_info, table_info_in_annot)

        excel_table_info_total.append(deepcopy(excel_table_info))

    # write excel file
    wb = Workbook()

    for tmp_excel_table_info_id, tmp_excel_table_info in enumerate(
        excel_table_info_total
    ):
        if tmp_excel_table_info_id == 0:
            ws = wb.active
            ws.title = f"sheet_{tmp_excel_table_info_id}"
        else:
            ws = wb.create_sheet(title=f"sheet_{tmp_excel_table_info_id}")

        for tmp_key, tmp_item in tmp_excel_table_info.items():
            # row and col must be at least 1
            row = tmp_item["row"] + 1
            col = tmp_item["col"] + 1
            content = tmp_item["content"]
            ws.cell(row=row, column=col, value=content)

    wb.save("a.xlsx")
