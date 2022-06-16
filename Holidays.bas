Attribute VB_Name = "Holidays"

'Holidays JP API から json を文字列として取得し、
'Dictionaryオブジェクトに加工してから処理します。
'Holidays JP API に存在しない年が指定された場合、
'ユーザー定義の実行時エラー: 513 を発生させます。
'使用する際は例外処理の導入を検討してください。

'参照設定
    'Microsoft XML, v6.0
    'Microsoft Scripting Runtime

'Holidays JP API 公式サイト
    'https://holidays-jp.github.io


Option Explicit


Public Function Is祝日(ByVal 日付 As Date) As Boolean
    '渡された日付が祝日かどうかを真偽値で返します。


    Dim 年 As Long
    年 = Year(日付)

    Dim 祝日一覧 As Dictionary
    Set 祝日一覧 = 祝日一覧を取得(年)

    Is祝日 = 祝日一覧.exists(日付)

End Function


Public Function 祝日名を取得(ByVal 日付 As Date) As String
    '渡された日付が祝日なら祝日の名前を返します。
    '祝日でなければ空文字列を返します。


    Dim 年 As Long
    年 = Year(日付)

    Dim 祝日一覧 As Dictionary
    Set 祝日一覧 = 祝日一覧を取得(年)

    祝日名を取得 = 祝日一覧.item(日付)

End Function


Public Function 祝日csvを作成(Optional ByVal 年 As Long, Optional ByVal 保存先 As String)
    '指定された年の祝日一覧をcsv化して保存します。
    '年が指定されなければ今年を指定します。
    '保存先が指定されなければ同階層に保存します。
    'すでに同名のファイルがあれば実行時エラー: 58 が発生します。
    '保存先フォルダが存在しなければ実行時エラー: 76 が発生します。


    If 年 = 0 Then
        年 = Year(Now)
    End If

    If 保存先 = "" Then
        保存先 = ThisWorkbook.Path & "\" & 年 & "年の祝日.csv"
    Else
        保存先 = 末尾一文字を削除(保存先, "/", "\")
        保存先 = 保存先 & "\" & 年 & "年の祝日.csv"
    End If

    Dim 祝日一覧 As Dictionary
    Set 祝日一覧 = 祝日一覧を取得(年)

    Dim FSO As New FileSystemObject

    On Error GoTo Finally

    Dim csv As TextStream
    Set csv = FSO.CreateTextFile(保存先, False)

    Dim 日付
    For Each 日付 In 祝日一覧
        Call csv.WriteLine(日付 & "," & 祝日一覧.item(日付))
    Next

Finally:
    If Not csv Is Nothing Then
        Call csv.Close
    End If
    Set FSO = Nothing

    If Err Then
        Call Err.Raise(Err, "祝日csvを作成", Err.Description)
    End If

End Function


Public Function 祝日一覧を取得(Optional ByVal 年 As Long) As Dictionary
    '指定された年の祝日を取得します。
    '年が指定されなければ今年を指定します。


    Dim 祝日json As String
    祝日json = 祝日jsonを取得(年)

    Set 祝日一覧を取得 = JsonをDictionayに変換(祝日json)

End Function


Private Function 祝日jsonを取得(ByVal 年 As Long) As String
    '指定した年の祝日のjsonファイルを文字列として取得します。
    '年が指定されていなければ今年を指定します。
    '指定された年のjsonファイルがなければエラーになります。


    Dim URL As String
    URL = Holidays_JP_APIのURLを取得(年)
    
    祝日jsonを取得 = Webサイトの全文字列を取得(URL)
    
    '取得した文字列がjson形式でなければエラーを発生させます。
    If 祝日jsonを取得 Like "*html*" Then
        Call Err.Raise(513, "祝日jsonを取得", _
            "指定された年のWebAPIが存在しません。（ " & 年 & " 年 )")
    End If
    
End Function


Private Function Holidays_JP_APIのURLを取得(ByVal 年 As Long) As String
    '指定された年のHolidays JP APIのURLを返します。


    '年が指定されていなければ今年を指定します。
    If 年 = 0 Then
        年 = Year(Now)
    End If

    Holidays_JP_APIのURLを取得 = _
        "https://holidays-jp.github.io/api/v1/" & 年 & "/date.json"

End Function


Private Function Webサイトの全文字列を取得(ByVal URL As String) As String
    '指定されたURLのページから全文字列を取得します。


    With New MSXML2.XMLHTTP60
        Call .Open("get", URL)
        Call .send
        Webサイトの全文字列を取得 = .responseText()
    End With

End Function


Private Function JsonをDictionayに変換(ByVal json As String) As Object
    'jsonをDictionaryオブジェクトに変換します。
    
    '引数
        'json : string

    '戻り値
        'Object : Dictionary


    json = Jsonから余分な文字列を除去(json)

    Dim 項目一覧() As String
    項目一覧() = Split(json, ",")

    Dim jsonDictionary As New Dictionary

    Dim 項目
    For Each 項目 In 項目一覧()
        Dim key As Date
        Dim item As String

        ' "項目名:値" をコロンで分けて代入します。
        key = Split(項目, ":")(0)
        item = Split(項目, ":")(1)

        Call jsonDictionary.Add(key, item)
    Next

    Set JsonをDictionayに変換 = jsonDictionary

End Function


Private Function Jsonから余分な文字列を除去(ByVal json As String) As String
    '値とコロンとカンマ以外を削除します。（末尾のカンマを除く）


    Dim 除去する文字一覧() As Variant
    除去する文字一覧() = Array("{", "}", """", " ", vbNewLine)

    Dim 除去する文字
    For Each 除去する文字 In 除去する文字一覧()
        json = Replace(json, 除去する文字, "")
    Next

    json = 末尾一文字を削除(json, ",")

    Jsonから余分な文字列を除去 = json

End Function


Private Function 末尾一文字を削除(ByVal 文字列 As String, ParamArray 削除する文字()) As String
    '削除する文字に合致していれば末尾一文字を削除します。


    Dim 文字数 As Long
    文字数 = Len(文字列)
    
    Dim 文字
    For Each 文字 In 削除する文字()
        If (Right(文字列, 1) = 文字) Then
            末尾一文字を削除 = Left(文字列, 文字数 - 1)
            Exit Function
        End If
    Next

    末尾一文字を削除 = 文字列

End Function
