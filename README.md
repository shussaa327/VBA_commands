# VBA_commands

## 他のブックを開く
 Workbooks.Open Filename:="フォルダパス¥ファイル名", ReadOnly:=True

同じフォルダ内のファイルを指定する場合
Workbooks.Open Filename:=ThisWorkbook.Path&"¥ファイル名", ReadOnly:=True
