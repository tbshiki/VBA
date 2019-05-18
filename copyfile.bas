Attribute VB_Name = "Module2"
'Sleep�֐��錾
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub �摜�ړ�()

Dim strDestinationFileDir As String '�\��t����̃f�B���N�g��
strDestinationFileDir = ActiveWorkbook.Path

Dim lastRow As Long
lastRow = Cells(Rows.Count, 8).End(xlUp).Row

Dim i As Long '��p
Dim j As Long '�s�p
For i = 2 To lastRow
    
    For j = 1 To 10
    
        Dim strSourceFile As String '�\�[�X�f�B���N�g��+�t�@�C����
        strSourceFile = Cells(i, 7 + j).Value
        
        If strSourceFile = "" Then
        
            Range(Cells(i, 7 + j), Cells(i, 17)).Interior.ColorIndex = 36
        
            Exit For
        
        End If
        
        Dim pos As Long '�t�@�C�������킯��ׂ�\�̈ʒu
        Dim strFileName As String '�t�@�C����
        pos = InStrRev(strSourceFile, "\")
        strFileName = Mid(strSourceFile, pos + 1)
        strFileName = Replace(strFileName, "_", "")
        strFileName = LCase(strFileName)
        
        Dim squareFlag As Boolean
        If j = 1 And squareFlag = False Then '1���ڗpsquare��������
            
            Dim dot As Long '�t�@�C�������킯��ׂ�.�̈ʒu
            Dim strNonExtension As String
            Dim strExtension As String
            dot = InStrRev(strFileName, ".")
            strNonExtension = Left(strFileName, dot - 1)
            strExtension = Right(strFileName, Len(strFileName) - dot + 1)
        
            FileCopy strSourceFile, strDestinationFileDir & "\" & strNonExtension & "square" & strExtension  '�ړ��ƃ��l�[���𓯎���
            
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

MsgBox "����"

End Sub
