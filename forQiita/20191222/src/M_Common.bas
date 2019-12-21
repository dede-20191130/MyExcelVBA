Attribute VB_Name = "M_Common"
'@Folder("VBAProject")
Option Explicit



'******************************************************************************************
'*�֐���    �FReturnCurrentBookNames
'*�@�\      �F���݊J����Ă���G�N�Z���u�b�N�̖��O�̌���������i�J���}��؂�j��Ԃ�
'*����(1)   �F����
'*�߂�l    �F����������
'******************************************************************************************
Public Function ReturnCurrentBookNames() As String
    
    '�萔
    Const FUNC_NAME As String = "ReturnCurrentBookNames"
    
    '�ϐ�
    Dim tempStr As String
    Dim cnt As Long
    
    On Error GoTo ErrorHandler
    
    '---�ȉ��ɏ������L�q---
    
    '�u�b�N�̖��O�擾
    For cnt = 1 To Workbooks.Count
        tempStr = tempStr & Workbooks(cnt).Name & ","
    Next
    
    '�Ō�̃J���}����
    tempStr = Mid(tempStr, 1, Len(tempStr) - 1)
    
    '�߂�l�ݒ�
    ReturnCurrentBookNames = tempStr
    
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



