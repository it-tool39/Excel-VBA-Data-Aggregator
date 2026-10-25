Attribute VB_Name = "payment_aggregation"
Option Explicit
'メイン処理
Sub Main()

    Dim inputFolder As String
    Dim doneFolder As String
    Dim masterWS As Worksheet
    Dim fileName As String
    Dim firstNewRow As Long
    Dim lastNewRow As Long
    Dim importedRows As Long
    Dim processingFile As Boolean
    Dim importCompleted As Boolean
    Dim errDescription As String

    On Error GoTo ErrHandler

    '大量データ処理時の画面更新停止による高速化
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    '集計シート設定
    Set masterWS = ThisWorkbook.Sheets("入金金額")

    'フォルダ設定
    inputFolder = ThisWorkbook.Path & "\data\"
    doneFolder = ThisWorkbook.Path & "\処理済み\"

    '処理済みフォルダ確認
    Call CheckFolder(doneFolder)


    'Excelファイル取得
    fileName = Dir(inputFolder & "*.xlsx")

    Do While fileName <> ""

        'マクロブック自身は除外
        If fileName <> ThisWorkbook.Name Then
            firstNewRow = 0
            importedRows = 0

            processingFile = True
            importCompleted = False

            Call ImportData( _
                inputFolder & fileName, _
                masterWS, fileName, firstNewRow, importedRows)

            importCompleted = True

            Call MoveFile(inputFolder, doneFolder, fileName)

            processingFile = False
        End If

        fileName = Dir()

    Loop

    '合計表示
    Call DisplayTotal(masterWS)

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    
   '商品コードから商品名取得
    Call SetProductName(masterWS)
     MsgBox "データの集計が完了しました！", vbInformation

    Exit Sub


ErrHandler:
    'エラー情報を保存
    errDescription = Err.Description

    'ファイル処理中にエラーが発生した場合
    If processingFile And Not importCompleted _
    And Not masterWS Is Nothing Then

    If firstNewRow > 0 And importedRows > 0 Then

        lastNewRow = firstNewRow + importedRows - 1

        masterWS.Range( _
            "A" & firstNewRow & ":H" & lastNewRow _
        ).ClearContents

        End If
    End If

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "エラーが発生しました。" & vbCrLf & _
           errDescription, vbExclamation


End Sub

' フォルダ確認
Sub CheckFolder(doneFolder As String)


    If Dir(doneFolder, vbDirectory) = "" Then
        MkDir doneFolder

    End If

End Sub

' データ取込
Sub ImportData(filePath As String, _
               masterWS As Worksheet, fileName As String, _
               ByRef firstImportedRow As Long, ByRef importedRowCount As Long)


    Dim targetWB As Workbook
    Dim targetWS As Worksheet

    Dim lastRow As Long
    Dim copyRows As Long
    Dim nextRow As Long
    Dim errNumber As Long
    Dim errDescription As String
    
    On Error GoTo ErrHandler


    Set targetWB = Workbooks.Open(filePath)
    Set targetWS = targetWB.Sheets(1)

    '最終行取得
    lastRow = targetWS.Cells(targetWS.Rows.Count, 1).End(xlUp).Row
    
    If lastRow >= 2 Then
        copyRows = lastRow - 1
        
        '貼付開始行
        nextRow = GetLastDataRow(masterWS) + 1
        firstImportedRow = nextRow
        importedRowCount = copyRows

        'A～E列コピー
        masterWS.Range("A" & nextRow).Resize(copyRows, 5).Value = _
        targetWS.Range("A2:E" & lastRow).Value

        '元ファイル名記録
        masterWS.Range("H" & nextRow).Resize(copyRows).Value = fileName

        'データ整形
        Call FormatData(masterWS, nextRow, copyRows)

    End If
    
    targetWB.Close SaveChanges:=False
    
    Set targetWB = Nothing
    Exit Sub
    
ErrHandler:
    '元のエラー情報を保存
    errNumber = Err.Number
    errDescription = Err.Description

     '開いているブックを閉じる
    If Not targetWB Is Nothing Then
        On Error Resume Next
        targetWB.Close SaveChanges:=False
        On Error GoTo 0
    End If

    '元のエラーを呼び出し元に伝える
    Err.Raise errNumber, , errDescription

End Sub
' データ整形
Sub FormatData(masterWS As Worksheet, _
               startRow As Long, _
               rowCount As Long)

    Dim r As Long
    Dim c As Long

    For r = startRow To startRow + rowCount - 1

        '日付統一
        If IsDate(masterWS.Cells(r, 1).Value) Then

            masterWS.Cells(r, 1).Value = _
            CDate(masterWS.Cells(r, 1).Value)

            masterWS.Cells(r, 1).NumberFormat = "m/d"

        End If

        'B～E列スペース削除
        For c = 2 To 5

            masterWS.Cells(r, c).Value = _
            Trim(masterWS.Cells(r, c).Value)

            masterWS.Cells(r, c).Value = _
            Replace(masterWS.Cells(r, c).Value, "　", "")

        Next c

        '処理日
        masterWS.Cells(r, 7).Value = Date

        masterWS.Cells(r, 7).NumberFormat = "m/d"

    Next r

End Sub


' ファイル移動
Sub MoveFile(inputFolder As String, _
             doneFolder As String, _
             fileName As String)

    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")

    fso.MoveFile _
    inputFolder & fileName, _
    doneFolder & fileName

    Set fso = Nothing

End Sub

' 集計表示
Sub DisplayTotal(masterWS As Worksheet)

    masterWS.Range("J2").Value = _
    WorksheetFunction.Sum(masterWS.Range("D:D"))

    masterWS.Range("J2").NumberFormat = "#,##0"

End Sub

'商品コードから商品名を表示
Sub SetProductName(masterWS As Worksheet)

    Dim wsCode As Worksheet
    Dim targetRow As Long
    Dim lastRow As Long
    Dim result As Variant

    Set wsCode = ThisWorkbook.Sheets("商品コード表")

    lastRow = masterWS.Cells(masterWS.Rows.Count, "E").End(xlUp).Row

    For targetRow = 2 To lastRow

        result = Application.VLookup( _
                masterWS.Cells(targetRow, "E").Value, _
                wsCode.Range("A:B"), 2, False)

    If IsError(result) Then
        result = "不明"
    End If

    masterWS.Cells(targetRow, "F").Value = result

Next targetRow


End Sub

' A～H列全体から最終データ行を取得
Function GetLastDataRow(ws As Worksheet) As Long

    Dim lastCell As Range

    Set lastCell = ws.Range("A:H").Find( _
        What:="*", _
        LookIn:=xlFormulas, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious)

    If lastCell Is Nothing Then
        GetLastDataRow = 0
    Else
        GetLastDataRow = lastCell.Row
    End If

End Function

