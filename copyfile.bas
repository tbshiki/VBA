Attribute VB_Name = "Module1"
'Sleep�֐��錾
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Sub �t�@�C���ړ�()

Dim strDestinationFileDir As String '�\��t����̃f�B���N�g��
Dim strFileName As String '�t�@�C����(�����������)
Dim strSourceFile As String '�\�[�X�f�B���N�g��+�t�@�C����
Dim pos As Long '�t�@�C�������킯��ׂ�\�̈ʒu
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

MsgBox "����"

End Sub
