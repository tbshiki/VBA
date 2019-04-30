Attribute VB_Name = "Module1"
'Sleep関数宣言
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Sub ファイル移動()

Dim strDestinationFileDir As String '貼り付け先のディレクトリ
Dim strFileName As String 'ファイル名(元も先も同じ)
Dim strSourceFile As String 'ソースディレクトリ+ファイル名
Dim pos As Long 'ファイル名をわける為の\の位置
Dim MaxRow As Long

strDestinationFileDir = ActiveWorkbook.Path

MaxRow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To MaxRow
    
    For j = 1 To 10

        strSourceFile = Cells(i, 0 + j).Value
        
        If strSourceFile = "" Then
        
            Range(Cells(i, 0 + j), Cells(i, 10)).Interior.ColorIndex = 36
        
            Exit For
        
        End If
        
        pos = InStrRev(strSourceFile, "\")
        strFileName = Mid(strSourceFile, pos + 1)
        
        FileCopy strSourceFile, strDestinationFileDir & "\" & strFileName
        
        Sleep 300
        
    Next j
    
    pos = 0
    strSourceFile = ""
    strFileName = ""

Next i

MsgBox "完了"

End Sub
