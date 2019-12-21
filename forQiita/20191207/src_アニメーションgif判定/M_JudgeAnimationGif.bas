Attribute VB_Name = "M_JudgeAnimationGif"
Option Explicit



'******************************************************************************************
'*�֐���    �FisAnimationGifImage
'*�@�\      �F�w���GIF�摜�t�@�C�����A�j���[�V�����ł��邩�ǂ������ʂ���
'*����(1)   �F�w�肷��摜�t�@�C���̃t���p�X
'*����(2)   �F�A�j���[�V�����ł��邩�ǂ����̔�����o�͂���Z���I�u�W�F�N�g
'******************************************************************************************
Public Sub isAnimationGifImage(ByVal filePath As String, _
                                ByRef outputCell As Range)
    
    '�萔
    Const FUNC_NAME As String = "isAnimationGifImage"
    
    '�ϐ�
    Dim fileExpression As String            '�t�@�C���g���q
    Dim picObject As Object                 '�摜�t�@�C���i�[�I�u�W�F�N�g
    Dim byteData() As Byte                  '�o�C�g��
    Dim fileNum As Long                     '�t�@�C���ԍ�
    Dim cnt As Long                         '���[�v�^�E���^
        
    On Error GoTo ErrorHandler
    '---�ȉ��ɏ������L�q---
    
    '///�g���q�擾
    fileExpression = Mid(filePath, InStrRev(filePath, ".") + 1)
    '�g���q��gif�ł���t�@�C���̂ݏ����𑱍s
    If fileExpression <> "gif" Then
        outputCell.Value = "�w�肳�ꂽ�t�@�C����GIF�t�@�C���ł͂���܂���B"
        GoTo ExitHandler
    End If
    
    '///�摜�t�@�C���̂ݏ����𑱍s
    On Error GoTo 0
    On Error Resume Next
    Set picObject = LoadPicture(filePath)
    On Error GoTo 0
    On Error GoTo ErrorHandler
    '������load�ł��Ă��Ȃ��Ȃ�ExitHandler��
    If picObject Is Nothing Then
        outputCell.Value = "�w�肳�ꂽ�t�@�C���͉摜�t�@�C���ł͂���܂���B"
        GoTo ExitHandler
    End If
    
    '///�o�C�g��̊i�[����
    '�]���Ă���t�@�C���ԍ��擾
    fileNum = FreeFile

    '�w�肳�ꂽ�t�@�C�����o�C�i�����[�h�ŊJ��
    Open filePath For Binary As #fileNum
        '�t�@�C���̒����Ŕz���������
        ReDim byteData(0 To LOF(fileNum))
        '�t�@�C�����o�C�i���œǂݍ����Byte�z��Ɋi�[
        Get fileNum, , byteData
    Close #fileNum
    
    '///gif�A�j�����ʏ���
    '�o�C�g������[�v����
    For cnt = 0 To UBound(byteData) - 7
        '�A�j���[�V����GIF�̎��ʕ�����uNETSCAPE�v�̃o�C�i���R�[�h��������܂Ń��[�v����
        If byteData(cnt) = 78 Then
            If byteData(cnt + 1) = 69 Then
                If byteData(cnt + 2) = 84 Then
                    If byteData(cnt + 3) = 83 Then
                        If byteData(cnt + 4) = 67 Then
                            If byteData(cnt + 5) = 65 Then
                                If byteData(cnt + 6) = 80 Then
                                    If byteData(cnt + 7) = 69 Then
                                        outputCell.Value = "TRUE:�w�肵���t�@�C���̓A�j���[�V����GIF�摜�t�@�C���ł��B"
                                        GoTo ExitHandler
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next
    
    '������Ȃ���Δ�A�j���[�V����GIF
    outputCell.Value = "FALSE:�w�肵���t�@�C���̓A�j���[�V����GIF�摜�t�@�C���ł͂���܂���B"
                                        

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
'*�֐���    �FconvertByteArrayFromString
'*�@�\      �F�w�蕶������o�C�g��ɕϊ����A�w��Z���ɃJ���}��؂�ŏo��
'*����(1)   �F�ϊ��Ώە�����i�[�Z���I�u�W�F�N�g
'*����(2)   �F�o�͐�̃Z���I�u�W�F�N�g
'******************************************************************************************
Public Sub convertByteArrayFromString(ByVal inputStrCell As Range, _
                                        ByRef outputCell As Range)
    
    '�萔
    Const FUNC_NAME As String = "convertByteArrayFromString"
    
    '�ϐ�
    Dim inputStr As String               '�Ώە�����
    Dim outputByte() As Byte             '�i�[��o�C�g��
    Dim cnt As Long                      '���[�v�J�E���^
    
    
    On Error GoTo ErrorHandler
    '---�ȉ��ɏ������L�q---
    
    '///�o�͐��������
    outputCell.Value = ""
    
    '///��������擾
    inputStr = StrConv(inputStrCell.Value, vbFromUnicode)
    
    '///�o�C�g��Ɋi�[
    outputByte = inputStr
    
    '///�o�͏���
    With outputCell
        '�o��
        For cnt = LBound(outputByte) To UBound(outputByte)
            .Value = .Value & outputByte(cnt)
            '�J���}��t��
            If cnt <> UBound(outputByte) Then
                .Value = .Value & ","
            End If
        Next
    End With
    

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
