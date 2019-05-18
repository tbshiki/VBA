Attribute VB_Name = "Module2"
'Sleep関数宣言
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub 画像移動()

Dim strDestinationFileDir As String '貼り付け先のディレクトリ
strDestinationFileDir = ActiveWorkbook.Path

Dim lastRow As Long
lastRow = Cells(Rows.Count, 8).End(xlUp).Row

Dim i As Long '列用
Dim j As Long '行用
For i = 2 To lastRow
    
    For j = 1 To 10
    
        Dim strSourceFile As String 'ソースディレクトリ+ファイル名
        strSourceFile = Cells(i, 7 + j).Value
        
        If strSourceFile = "" Then
        
            Range(Cells(i, 7 + j), Cells(i, 17)).Interior.ColorIndex = 36
        
            Exit For
        
        End If
        
        Dim pos As Long 'ファイル名をわける為の\の位置
        Dim strFileName As String 'ファイル名
        pos = InStrRev(strSourceFile, "\")
        strFileName = Mid(strSourceFile, pos + 1)
        strFileName = Replace(strFileName, "_", "")
        strFileName = LCase(strFileName)
        
        Dim squareFlag As Boolean
        If j = 1 And squareFlag = False Then '1枚目用square条件分岐
            
            Dim dot As Long 'ファイル名をわける為の.の位置
            Dim strNonExtension As String
            Dim strExtension As String
            dot = InStrRev(strFileName, ".")
            strNonExtension = Left(strFileName, dot - 1)
            strExtension = Right(strFileName, Len(strFileName) - dot + 1)
        
            FileCopy strSourceFile, strDestinationFileDir & "\" & strNonExtension & "square" & strExtension  '移動とリネームを同時に
            
            Sleep 300
            
            squareFlag = True
        
        Else
            
            FileCopy strSourceFile, strDestinationFileDir & "\" & strFileName
            
            Sleep 300
            
        End If
        
    Next j
    
    squareFlag = False
    pos = 0
    dot = 0
    strSourceFile = ""
    strFileName = ""
    strNonExtension = ""
    strExtension = ""

Next i

MsgBox "完了"

End Sub
