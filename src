Sub メールで参加者名簿を送信()

    Dim wb As Workbook
    Dim tempWb As Workbook
    Dim sheetName As String
    Dim tempPath As String
    Dim tempFileName As String
    Dim outlookApp As Object
    Dim outlookMail As Object

    ' メール情報取得（N6～N9）
    Dim mailTo As String, mailCC As String
    Dim mailSubject As String, mailBody As String

    ' 「参加者名簿」シートからメール情報を取得
    sheetName = "参加者名簿"
    With ThisWorkbook.Sheets(sheetName)
        mailTo = .Range("N6").Value
        mailCC = .Range("N7").Value
        mailSubject = .Range("N8").Value
        mailBody = .Range("N9").Value
    End With

    tempPath = Environ("TEMP") & "\"

    ' ファイル名に使用できない文字を除去し、A3セルの値を使う
    Dim fileNamePart As String
    fileNamePart = CleanFileName(ThisWorkbook.Sheets(sheetName).Range("A3").Value)
    tempFileName = "部員名簿" & fileNamePart & ".xlsx"

    ' I3セルの数字から行番号を取得
    Dim rowNum As Long
    rowNum = GetRowFromJ3()

    ' 元のブックとシート参照
    Set wb = ThisWorkbook
    Set srcSheet = wb.Sheets(sheetName)

    ' 一時的な新しいワークブックを作成
    Set tempWb = Workbooks.Add
    Set destSheet = tempWb.Sheets(1)

    ' I3の数字から5行分を選択してコピー
    srcSheet.Range(srcSheet.Cells(1, "A"), srcSheet.Cells(rowNum + 5, "K")).Copy
    With destSheet.Range("A1")
        .PasteSpecial Paste:=xlPasteValues
        .PasteSpecial Paste:=xlPasteFormats
    End With

    ' 列幅をコピー
    Dim col As Integer
    For col = 1 To 11
        destSheet.Columns(col).ColumnWidth = srcSheet.Columns(col).ColumnWidth
    Next col

    ' 一時ファイルとして保存
    tempWb.SaveAs fileName:=tempPath & tempFileName, FileFormat:=xlOpenXMLWorkbook
    tempWb.Close SaveChanges:=False

    ' Outlookアプリケーションを起動
    Set outlookApp = CreateObject("Outlook.Application")
    Set outlookMail = outlookApp.CreateItem(0)

    ' メールの内容を設定（HTML形式）
    With outlookMail
        .To = mailTo
        .CC = mailCC
        .Subject = mailSubject
        .HTMLBody = "<p>" & Replace(mailBody, vbCrLf, "<br>") & "</p>"
        .Attachments.Add tempPath & tempFileName
        .Display
    End With

    ' 後始末
    Set outlookMail = Nothing
    Set outlookApp = Nothing

End Sub

' I3セルの値に基づいて行番号を抽出
Function GetRowFromJ3() As Long
    Dim cellValue As String
    Dim startPos As Long
    Dim endPos As Long
    Dim number As Long
    
    cellValue = ThisWorkbook.Sheets("参加者名簿").Range("I3").Value

    ' "参加者："の位置を見つける
    startPos = InStr(cellValue, "参加者：") + Len("参加者：")
    
    ' "名"の位置を見つける
    endPos = InStr(cellValue, "名")
    
    ' "参加者："から"名"の間の文字列を抽出して数字に変換
    If startPos > 0 And endPos > startPos Then
        number = CLng(Mid(cellValue, startPos, endPos - startPos))
    Else
        MsgBox "無効なデータ形式です。"
        Exit Function
    End If
    
    ' 数字が見つからない場合
    If number = 0 Then
        MsgBox "数字が見つかりませんでした。"
        Exit Function
    End If
    
    GetRowFromJ3 = number
End Function

' ファイル名として使えない文字を削除・置換する関数
Function CleanFileName(str As String) As String
    Dim illegalChars As Variant
    illegalChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    Dim i As Integer
    For i = LBound(illegalChars) To UBound(illegalChars)
        str = Replace(str, illegalChars(i), "_")
    Next i

    CleanFileName = str
End Function


