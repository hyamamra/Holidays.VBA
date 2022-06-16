Attribute VB_Name = "Sample"
Option Explicit


Public Sub Sample_Is祝日()
    '指定した日付が祝日か確認します。


    Dim 日付 As Date
    日付 = 日付を入力()

    On Error GoTo Catch

    Debug.Print Is祝日(日付)

    Exit Sub
    
Catch:
    Call MsgBox("実行時エラー: " & Err & vbNewLine _
        & Err.Description, vbCritical, Err.Source)
    
End Sub


Public Sub Sample_祝日名を取得()
    '指定された日付が祝日であれば祝日の名前が返されます。
    '祝日ではない場合、空文字列が返されます。
    'そのため、空文字列であれば祝日ではないと判断できます。


    Dim 日付 As Date
    日付 = 日付を入力()

    On Error GoTo Catch

    '指定した日付の祝日名を取得します。
    Dim 祝日 As String
    祝日 = 祝日名を取得(日付)
    
    If 祝日 = "" Then
        Debug.Print "祝日ではありません。"
    Else
        Debug.Print 祝日
    End If

    Exit Sub
    
Catch:
    Call MsgBox("実行時エラー: " & Err & vbNewLine _
        & Err.Description, vbCritical, Err.Source)
    
End Sub


Private Function 日付を入力() As Date

    Dim 日付 As Date
    日付 = InputBox("日付を入力してください。")
    日付 = DateValue(日付)
    日付を入力 = 日付

End Function


Public Sub Sample_祝日csvを作成()
    '指定した階層に祝日のcsvファイルを作成します。

    Dim 年 As Long
    年 = 年を入力()

    Dim 保存先 As String
    保存先 = 保存先を入力()

    On Error GoTo Catch

    Call 祝日csvを作成(年, 保存先)

    Exit Sub

Catch:
    Call MsgBox("実行時エラー: " & Err & vbNewLine _
        & Err.Description, vbCritical, Err.Source)

End Sub


Public Sub Sample_祝日一覧を取得()
    '祝日（Dictionary型）を取得して For Each で取り出します。

    Dim 年 As Long
    年 = 年を入力()
    
    On Error GoTo Catch

    Dim 祝日一覧 As Dictionary
    Set 祝日一覧 = 祝日一覧を取得(年)
    
    Dim 日付
    For Each 日付 In 祝日一覧
        Debug.Print 日付, 祝日一覧.item(日付)
    Next

    Exit Sub
    
Catch:
    Call MsgBox("実行時エラー: " & Err & vbNewLine _
        & Err.Description, vbCritical, Err.Source)
    
End Sub


Private Function 年を入力() As Long

    Dim 年 As Long
    年 = InputBox("年を入力してください。")
    年を入力 = 年

End Function


Private Function 保存先を入力() As String

    Dim 保存先 As String
    保存先 = InputBox("保存先を絶対パスで指定してください。")
    保存先を入力 = 保存先

End Function
