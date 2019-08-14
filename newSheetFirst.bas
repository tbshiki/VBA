Sub newSheetFirst(ByVal strSheetName As String)
'新しいシートを一番左に挿入して任意の名前(引数)を付ける
'同じ名前のシートがあった場合はシートの見出しの色を"色なし"にして(*)で連番を付ける

Worksheets.Add before:=Worksheets(1)

'同じ名前のシートがある場合、見出しの色を"色なし"に
Dim WS As Worksheet
For Each WS In Worksheets
    If WS.Name = strSheetName Then Sheets(strSheetName).Tab.ColorIndex = xlColorIndexNone
Next WS

On Error Resume Next
ActiveSheet.Name = strSheetName

Dim n As Long
n = 1
Do Until Err.Number = 0
    Err.Clear
    n = n + 1
    Sheets(strSheetName).Name = strSheetName & "(" & n & ")"
Loop

ActiveSheet.Name = strSheetName

On Error GoTo 0

End Sub
