Attribute VB_Name = "M_JudgeAnimationGif"
Option Explicit



'******************************************************************************************
'*関数名    ：isAnimationGifImage
'*機能      ：指定のGIF画像ファイルがアニメーションであるかどうか判別する
'*引数(1)   ：指定する画像ファイルのフルパス
'*引数(2)   ：アニメーションであるかどうかの判定を出力するセルオブジェクト
'******************************************************************************************
Public Sub isAnimationGifImage(ByVal filePath As String, _
                                ByRef outputCell As Range)
    
    '定数
    Const FUNC_NAME As String = "isAnimationGifImage"
    
    '変数
    Dim fileExpression As String            'ファイル拡張子
    Dim picObject As Object                 '画像ファイル格納オブジェクト
    Dim byteData() As Byte                  'バイト列
    Dim fileNum As Long                     'ファイル番号
    Dim cnt As Long                         'ループタウンタ
        
    On Error GoTo ErrorHandler
    '---以下に処理を記述---
    
    '///拡張子取得
    fileExpression = Mid(filePath, InStrRev(filePath, ".") + 1)
    '拡張子がgifであるファイルのみ処理を続行
    If fileExpression <> "gif" Then
        outputCell.Value = "指定されたファイルはGIFファイルではありません。"
        GoTo ExitHandler
    End If
    
    '///画像ファイルのみ処理を続行
    On Error GoTo 0
    On Error Resume Next
    Set picObject = LoadPicture(filePath)
    On Error GoTo 0
    On Error GoTo ErrorHandler
    '正しくloadできていないならExitHandlerへ
    If picObject Is Nothing Then
        outputCell.Value = "指定されたファイルは画像ファイルではありません。"
        GoTo ExitHandler
    End If
    
    '///バイト列の格納処理
    '余っているファイル番号取得
    fileNum = FreeFile

    '指定されたファイルをバイナリモードで開く
    Open filePath For Binary As #fileNum
        'ファイルの長さで配列を初期化
        ReDim byteData(0 To LOF(fileNum))
        'ファイルをバイナリで読み込んでByte配列に格納
        Get fileNum, , byteData
    Close #fileNum
    
    '///gifアニメ判別処理
    'バイト列をループする
    For cnt = 0 To UBound(byteData) - 7
        'アニメーションGIFの識別文字列「NETSCAPE」のバイナリコードが見つかるまでループする
        If byteData(cnt) = 78 Then
            If byteData(cnt + 1) = 69 Then
                If byteData(cnt + 2) = 84 Then
                    If byteData(cnt + 3) = 83 Then
                        If byteData(cnt + 4) = 67 Then
                            If byteData(cnt + 5) = 65 Then
                                If byteData(cnt + 6) = 80 Then
                                    If byteData(cnt + 7) = 69 Then
                                        outputCell.Value = "TRUE:指定したファイルはアニメーションGIF画像ファイルです。"
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
    
    '見つからなければ非アニメーションGIF
    outputCell.Value = "FALSE:指定したファイルはアニメーションGIF画像ファイルではありません。"
                                        

ExitHandler:

    Exit Sub
    
ErrorHandler:

        MsgBox "エラーが発生しましたので終了します" & _
                vbLf & _
                "関数名：" & FUNC_NAME & _
                vbLf & _
                "エラー番号" & Err.Number & Chr(13) & Err.Description, vbCritical
        
        GoTo ExitHandler
        
End Sub







'******************************************************************************************
'*関数名    ：convertByteArrayFromString
'*機能      ：指定文字列をバイト列に変換し、指定セルにカンマ区切りで出力
'*引数(1)   ：変換対象文字列格納セルオブジェクト
'*引数(2)   ：出力先のセルオブジェクト
'******************************************************************************************
Public Sub convertByteArrayFromString(ByVal inputStrCell As Range, _
                                        ByRef outputCell As Range)
    
    '定数
    Const FUNC_NAME As String = "convertByteArrayFromString"
    
    '変数
    Dim inputStr As String               '対象文字列
    Dim outputByte() As Byte             '格納先バイト列
    Dim cnt As Long                      'ループカウンタ
    
    
    On Error GoTo ErrorHandler
    '---以下に処理を記述---
    
    '///出力先を初期化
    outputCell.Value = ""
    
    '///文字列を取得
    inputStr = StrConv(inputStrCell.Value, vbFromUnicode)
    
    '///バイト列に格納
    outputByte = inputStr
    
    '///出力処理
    With outputCell
        '出力
        For cnt = LBound(outputByte) To UBound(outputByte)
            .Value = .Value & outputByte(cnt)
            'カンマを付加
            If cnt <> UBound(outputByte) Then
                .Value = .Value & ","
            End If
        Next
    End With
    

ExitHandler:

    Exit Sub
    
ErrorHandler:

        MsgBox "エラーが発生しましたので終了します" & _
                vbLf & _
                "関数名：" & FUNC_NAME & _
                vbLf & _
                "エラー番号" & Err.Number & Chr(13) & Err.Description, vbCritical
        
        GoTo ExitHandler
        
End Sub
