Attribute VB_Name = "M_NamesExport"
Option Explicit

'******************************************************************************************
'*�֐���    �FNamesExport
'*�@�\      �F�u�b�N�̖��O��`�����G�N�X�|�[�g����csv���o�͂���B
'*            �����̃t�H���_�ɁA�����̃u�b�N�Ɋ֘A�t�������O��`���G�N�X�|�[�g����csv�t�@�C�����쐬����
'*����(1)   �F�o�͐�t�H���_�p�X
'*����(2)   �F���O��`�o�͑Ώۂ̃u�b�N�I�u�W�F�N�g
'******************************************************************************************
Public Sub NamesExport(ByVal prmFolderPath As String, _
                       ByVal prmWB As Workbook)
    
    '�萔
    Const FUNC_NAME As String = "NamesOutput"
    Const ForWriting As Long = 2                 '�V�K��������
    
    '�ϐ�
    Dim outPutCsvPath As String
    
    On Error GoTo ErrorHandler
    '---�ȉ��ɏ������L�q---
    
    With CreateObject("Scripting.FileSystemObject")
    
        '�o��csv�̃p�X��ݒ�
        outPutCsvPath = (prmFolderPath & "\" & .GetBaseName(prmWB.Name) & "_Names.csv")
    
        'csv�t�@�C�����쐬����
        '���łɓ����t�@�C�������݂���ꍇ�͏㏑���̉ۂ�₤
        If .FileExists(outPutCsvPath) Then
            If MsgBox("�㏑�����܂��B" & vbLf & "��낵���ł����B", vbYesNo) = vbNo Then
                GoTo ExitHandler
            End If
        End If
        
        .CreateTextFile (outPutCsvPath)
        
        'csv�t�@�C�����J��
        With .OpenTextFile(outPutCsvPath, ForWriting, True)
            '���O��`�̏�������
            .Write NamesExport_Sub_PullNamesData(prmWB)
            .Close
        End With
        
    End With
    
    MsgBox "�G�N�X�|�[�g���������܂����B"
    
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
'*�֐���    �FNamesExport_Sub_PullNamesData
'*�@�\      �F�����̃u�b�N�̖��O��`�f�[�^��csv�`��������Ƃ��Ď擾
'*����(1)   �F�Ώۂ̃u�b�N�I�u�W�F�N�g
'*�߂�l    �Fcsv�`��������
'******************************************************************************************
Private Function NamesExport_Sub_PullNamesData(ByVal prmWB As Workbook) As String
    
    '�萔
    Const FUNC_NAME As String = "NamesOutput_Sub_PullNamesData"
    
    '�ϐ�
    Dim cnt As Long
    Dim tempStr As String
    
    On Error GoTo ErrorHandler
    '�߂�l�����l
    NamesExport_Sub_PullNamesData = ""
    
    '---�ȉ��ɏ������L�q---
    
    'csv�w�b�_�[��ݒ�
    tempStr = """���O""" & DELIMITER & _
              """�Q�Ɣ͈́iA1�j""" & DELIMITER & _
              """�Q�Ɣ͈́iR1C1�j""" & DELIMITER & _
              """�R�����g"""
    tempStr = tempStr & vbCrLf
    
    '��`���ꂽ���O�̐��������[�v
    For cnt = 1 To prmWB.Names.Count
        
        '�u�b�N�͈̖͂��O��`�̂ݒ��o
        If TypeName(prmWB.Names(cnt).Parent) = "Workbook" Then
        
            '���擾
            tempStr = tempStr & prmWB.Names(cnt).Name & DELIMITER & _
                                                      prmWB.Names(cnt).RefersTo & DELIMITER & _
                                                      prmWB.Names(cnt).RefersToR1C1 & DELIMITER & _
                                                      prmWB.Names(cnt).Comment
                                                      
            '���s�ݒ�
            tempStr = tempStr & vbCrLf
        End If
    Next
    
    '�Ō�̉��s�R�[�h������
    tempStr = Mid(tempStr, 1, Len(tempStr) - 2)
    
    '�߂�l�ݒ�
    NamesExport_Sub_PullNamesData = tempStr
    
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


