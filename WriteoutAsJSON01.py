# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import uno, json

# LibreOfiiceの設定->->セキュリティ->マクロのセキュリティで
# HighをMediumに変更する（これでPythonが実行できる）
# Pythonは以下の位置にある（システムのPythonではない）
# /Applications/LibreOffice.app/Contents/Resources/python
# Macの場合、~/Library/Application Support/LibreOffice/4/user/Scripts/python/
# にこのpyファイルを置くと、Tools/Macro/OrganizeMacro/pythonに表示される
# 初めて使う場合、Scripts/python/のフォルダを作る必要がある

def WriteoutAsJSON01():
    """
        選択範囲をJSON形式で書き出す
        1行目をキー、2行目以降を値とする
        辞書オブジェクトを配列として書き出す
    """
    # 出力位置はパーミッションがあるフルパス
    #output_filename = "/Users/**your username**/Desktop/out.json"
    output_filename = "/Users/isao/Desktop/writeout.json"

    fout = open( output_filename, 'w', newline='', encoding='utf-8' )
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
    obj_result = []
    for row in range(first_row + 1, last_row + 1):
        obj_row = dict()
        for col in range(first_col, last_col + 1):
            k = sheet.getCellByPosition(col, first_row).String  # 文字列か値かで出力は違う
            # v = sheet.getCellByPosition(col, row).Value        # Valueは数値、日付や％も内部数値が表示される、文字型は0.0
            v = sheet.getCellByPosition(col, row).String         # Stringは「表示通り」
            obj_row[k] = v
        obj_result.append(obj_row)
    json.dump(obj_result, fout, sort_keys=True, indent=4, ensure_ascii=False)

    fout.close()