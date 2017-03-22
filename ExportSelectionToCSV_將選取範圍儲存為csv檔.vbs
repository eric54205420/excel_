Option Explicit

' =============== 將選取範圍 儲存為 csv 檔 ===============



' 將選取範圍 儲存為 csv 檔
Public Sub ExcelRowsToCSV()
 
  Dim iPtr As Integer
  Dim sFileName As String
  Dim intFH As Integer
  Dim aRange As Range
  Dim iLastColumn As Integer
  Dim oCell As Range
  Dim iRec As Long
 
 ' 由互動窗口取得選取資料
  Set aRange = Application.InputBox("Select a range:-", , Selection.Address, , , , , Type:=8)
  iLastColumn = aRange.Column + aRange.Columns.Count - 1
  
  ' 處理儲存檔名問題
  ' 由互動窗口決定要儲存位置
  iPtr = InStrRev(ActiveWorkbook.FullName, ".")
  sFileName = Left(ActiveWorkbook.FullName, iPtr - 1) & ".csv"
  sFileName = Application.GetSaveAsFilename(InitialFileName:=sFileName, _
                                                        FileFilter:="CSV (Comma delimited) (*.csv), *.csv")
  If sFileName = "False" Then Exit Sub
    
  Close
  
  ' 將選取資料寫入
  intFH = FreeFile()
  Open sFileName For Output As intFH
  
  iRec = 0
  For Each oCell In aRange
    If oCell.Column = iLastColumn Then
      Print #intFH, oCell.Value
      iRec = iRec + 1
    Else
      Print #intFH, oCell.Value; ",";
    End If
  Next oCell
   
  Close intFH
  
  ' 完成操作提示
  MsgBox "Finished: " & CStr(iRec) & " records written to " _
     & sFileName & Space(10), vbOKOnly + vbInformation
 
End Sub
