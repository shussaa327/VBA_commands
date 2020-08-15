# VBA_commands

## 他のブックを開く
```
Workbooks.Open Filename:="フォルダパス¥ファイル名", ReadOnly:=True
```  
  
同じフォルダ内のファイルを指定する場合
```  
Workbooks.Open Filename:=ThisWorkbook.Path&"¥ファイル名", ReadOnly:=True
```  
他ファイルの抽出
```  
Dim wb As Workbook
Set wb = Workbooks.Open(Filename:=ThisWorkbook.Path & "\機器20200815", ReadOnly:=True)
ThisWorkbook.Sheets = wb.Sheets
    
wb.Close SaveChanges:=False
```
