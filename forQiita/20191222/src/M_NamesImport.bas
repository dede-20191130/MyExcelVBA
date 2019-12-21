Attribute VB_Name = "M_NamesImport"
'@Folder("VBAProject")
Option Explicit

'******************************************************************************************
'*�֐���    �FNamesImport
'*�@�\      �Fcsv�t�@�C���̖��O��`��ǂݎ��A�w�肵���u�b�N�̖��O��`�ݒ�ɐV�����擾�������O��`���C���|�[�g����B
'*����(1)   �Fcsv�t�@�C���̃p�X
'*����(2)   �F�C���|�[�g�Ώۃu�b�N�I�u�W�F�N�g
'******************************************************************************************
Public Sub NamesImport(ByVal prmCsvFilePath As String, _
                       ByRef prmWB As Workbook)
    
    '�萔
    Const FUNC_NAME As String = "NamesInput"
    
    '�ϐ�
    Dim namesArr As Variant
    Dim cnt As Long
    
    On Error GoTo ErrorHandler
    '---�ȉ��ɏ������L�q---
    
    'csv�ǂݍ���
    namesArr = NamesImport_Sub_CsvToArray(prmCsvFilePath)
    
    '�w�肵���u�b�N�̖��O��`�ݒ�ɐV�����擾�������O��`��}������
    For cnt = LBound(namesArr, 1) To UBound(namesArr, 1)
        '�Q�ƌ`�����Ƃɏ����𕪊�
        Select Case Application.ReferenceStyle
        Case xlA1
            prmWB.Names.Add namesArr(cnt, 0), namesArr(cnt, 1), True
        Case Else
            prmWB.Names.Add Name:=namesArr(cnt, 0), RefersToR1C1:=namesArr(cnt, 2), Visible:=True
        End Select
        
        '�R�����g��ݒ�
        prmWB.Names(namesArr(cnt, 0)).Comment = namesArr(cnt, 3)
    Next
    
    MsgBox "���O��`�̃C���|�[�g���������܂����B"

ExitHandler:
    
    
    Exit Sub
    
ErrorHandler:

    MsgBox "�G���[���������܂����̂ŏI�����܂�" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ�" & Err.Number & Chr(13) & Err.Description, vbCritical
        
    GoTo ExitHandler
        
End Sub

'******************************************************************************************
'*�֐���    �FNamesImport_Sub_CsvToArray
'*�@�\      �Fcsv�t�@�C���̓��e��񎟌��z��Ɋi�[
'*����(1)   �Fcsv�t�@�C���p�X
'*�߂�l    �F�񎟌��z��
'******************************************************************************************
Private Function NamesImport_Sub_CsvToArray(ByVal prmCsvFilePath As String) As Variant
    
    '�萔
    Const FUNC_NAME As String = "NamesInput_Sub_CsvToArray"
    
    '�ϐ�
    Dim strArr() As String
    Dim lineArr As Variant
    Dim iFile As Long
    Dim buf As String
    Dim cnt_1 As Long
    Dim cnt_2 As Long
    Dim tempNum As Long
    
    On Error GoTo ErrorHandler
    
    '---�ȉ��ɏ������L�q---
    
    '������
    cnt_1 = 0
    cnt_2 = 0
    
    '�t�@�C���ԍ��擾
    iFile = FreeFile
    
    'csv�t�@�C���̍s���擾
    With CreateObject("Scripting.FileSystemObject").OpenTextFile(prmCsvFilePath, 8)
        tempNum = .Line
    End With
    
    'Redim
    ReDim strArr(tempNum - 2, 3)
    
    'csv���J���A�f�[�^���s�����z��Ɋi�[
    Open prmCsvFilePath For Input As #iFile
    Do Until EOF(iFile)
        Line Input #iFile, buf
        '�w�b�_�[�s�͖�������
        If cnt_1 <> 0 Then
            lineArr = Split(buf, DELIMITER)
            For cnt_2 = 0 To UBound(lineArr)
                strArr(cnt_1 - 1, cnt_2) = lineArr(cnt_2)
            Next cnt_2
        End If
        
        cnt_1 = cnt_1 + 1
            
    Loop
    Close #iFile
    
    '������������������������������������������������������������������������������������������������������
    '���L�����̓G���[�ɂȂ�FVBA�ł͓񎟌��z��ɑ΂��Ă�Redim�͍Ō�̎����̗v�f���̂݋��e
    '    '�g�p���Ă��Ȃ��z������T�C�Y����
    '    tempNum = cnt_1 - 1
    '    ReDim Preserve strArr(tempNum, 3)
    '������������������������������������������������������������������������������������������������������

    '�߂�l�ݒ�
    NamesImport_Sub_CsvToArray = strArr
    
ExitHandler:

    Exit Function
    
ErrorHandler:

    MsgBox "�G���[���������܂����̂ŏI�����܂�" & _
           vbLf & _
           "�֐����F" & FUNC_NAME & _
           vbLf & _
           "�G���[�ԍ�" & Err.Number & Chr(13) & Err.Description, vbCritical
        
    GoTo ExitHandler
        
End Function


