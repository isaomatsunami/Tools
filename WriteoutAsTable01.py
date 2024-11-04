# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import uno
import xml.etree.ElementTree as ET

# LibreOfiiceの設定->->セキュリティ->マクロのセキュリティで
# HighをMediumに変更する（これでPythonが実行できる）
# Pythonは以下の位置にある（システムのPythonではない）
# /Applications/LibreOffice.app/Contents/Resources/python
# Macの場合、~/Library/Application Support/LibreOffice/4/user/Scripts/python/
# にこのpyファイルを置くと、Tools/Macro/OrganizeMacro/pythonに表示される
# 初めて使う場合、Scripts/python/のフォルダを作る必要がある

def WriteoutAsTable01():
    """
        選択範囲をHTML TABLE形式で書き出す
        1行目をTHEAD/TH、それ以降をTBODY/TDで書き出す
    """
    # 出力位置はパーミッションがあるフルパス
    #output_filename = "/Users/**your username**/Desktop/out.json"
    output_filename = "/Users/isao/Desktop/writeout.html"

    desktop = XSCRIPTCONTEXT.getDesktop()
    component = desktop.getCurrentComponent()
    try:
        sheets = component.getSheets()
    except AttributeError:
        raise Exception("This script is for LibreOffice Calc only")

    sheet = component.CurrentController.getActiveSheet()
    objSelection = component.getCurrentSelection()
    objArea = objSelection.getRangeAddress()
    first_row, last_row = objArea.StartRow, objArea.EndRow
    first_col, last_col = objArea.StartColumn, objArea.EndColumn
    table = ET.Element('table')
    thead = ET.SubElement(table, 'thead')
    tbody = ET.SubElement(table, 'tbody')
    thead_tr = ET.SubElement(thead, 'tr')
    for col in range(first_col, last_col + 1):
        th = ET.SubElement(thead_tr, 'th')
        th.text = sheet.getCellByPosition(col, first_row).String
    for row in range(first_row + 1, last_row + 1):
        tbody_tr = ET.SubElement(tbody, 'tr')
        for col in range(first_col, last_col + 1):
            td = ET.SubElement(tbody_tr, 'td')
            td.text = sheet.getCellByPosition(col, row).String
            # 文字列か値かで出力は違う
            # Valueは数値、日付や％も内部数値が表示される、文字型は0.0
            # Stringは「表示通り」
    root = ET.ElementTree(table)
    root.write(output_filename, encoding='UTF-8', method='xml')
