Attribute VB_Name = "Module2"
'Sleep�֐��錾
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub �摜�ړ�()

Dim strDestinationFileDir As String '�\��t����̃f�B���N�g��
strDestinationFileDir = ActiveWorkbook.Path

Dim lastRow As Long
lastRow = Cells(Rows.Count, 8).End(xlUp).Row

Dim firstCol As Long
Dim lastCol As Long
firstCol = 8 'firstCol��ڂ���
lastCol = 17 'lastCol��ڂ܂�

Dim i As Long '�s�p
Dim j As Long '��p
For i = 2 To lastRow
    
    For j = 0 To lastCol - firstCol
CONTINUE:
        Dim strSourceFile As String '�\�[�X�f�B���N�g��+�t�@�C����
        strSourceFile = Cells(i, firstCol + j).Value
        
        If strSourceFile = "" Then
        
            Range(Cells(i, firstCol + j), Cells(i, lastCol)).Interior.ColorIndex = 36
        
            Exit For
        
        End If
        
        Dim pos As Long '�t�@�C�������킯��ׂ�\�̈ʒu
        Dim strFileName As String '�t�@�C����
        pos = InStrRev(strSourceFile, "\")
        strFileName = Mid(strSourceFile, pos + 1)
        strFileName = Replace(strFileName, "_", "")
        strFileName = LCase(strFileName)
        
        Dim squareFlag As Boolean
        If j = 0 And squareFlag = False Then '1���ڗpsquare��������
            
            Dim dot As Long '�t�@�C�������킯��ׂ�.�̈ʒu
            Dim strNonExtension As String
            Dim strExtension As String
            dot = InStrRev(strFileName, ".")
            strNonExtension = Left(strFileName, dot - 1)
            strExtension = Right(strFileName, Len(strFileName) - dot + 1)
        
            FileCopy strSourceFile, strDestinationFileDir & "\" & strNonExtension & "square" & strExtension  '�ړ��ƃ��l�[���𓯎���
            Cells(i, firstCol - 1).Value = "xxxxxxxxxx" & strNonExtension & "square" & strExtension
            
            Sleep 300
            
            squareFlag = True
            pos = 0
            dot = 0
            strSourceFile = ""
            strFileName = ""
            strNonExtension = ""
            strExtension = ""
            
            GoTo CONTINUE
        
        Else
            
            FileCopy strSourceFile, strDestinationFileDir & "\" & strFileName
            Cells(i, firstCol + j).Value = "xxxxxxxxxx" & strFileName
            
            Sleep 300
            
        End If
        
        pos = 0
        dot = 0
        strSourceFile = ""
        strFileName = ""
        strNonExtension = ""
        strExtension = ""
        
    Next j
    
    squareFlag = False

Next i

MsgBox "����"

End Sub
