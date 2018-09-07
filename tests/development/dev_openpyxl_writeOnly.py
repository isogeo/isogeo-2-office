# -*- coding: UTF-8 -*-
#! python3

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.worksheet.write_only import WriteOnlyCell

wb = Workbook(write_only=True)
ws = wb.create_sheet("i2o_thumbnails")

# columns dimensions
ws.column_dimensions["A"].width = 35
ws.column_dimensions["B"].width = 75
ws.column_dimensions["C"].width = 75

# headers
head_col1 = WriteOnlyCell(ws, value="isogeo_uuid")
head_col2 = WriteOnlyCell(ws, value="isogeo_title_slugged")
head_col3 = WriteOnlyCell(ws, value="img_abs_path")
# comments
comment = Comment(text="Do not modify worksheet structure",
                  author="Isogeo")

head_col1.comment = head_col2.comment = head_col3.comment = comment

# styling
head_col1.style = head_col2.style = head_col3.style = "Headline 2"
# insert
ws.append((head_col1,
           head_col2,
           head_col3)
)

# realist fake
ws.append(("c4b7ad9732454beca1ab3ec1958ffa50",
           "title-slugged",
           "resources/table.svg")
)

# random values
for letter in "Isogeo, easy access to geodata!":
    ws.append((letter, "hop", 50))

wb.save("test_xl_writeOnly.xlsx")
