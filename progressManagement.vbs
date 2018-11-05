Sub makeAggregateTable()
    '各試験仕様書のにバリエーションを管理するための管理シートを作成する。
    Dim macroWbName As String                           'マクロのブック名
    Dim macroWsName As String                           'マクロのシート名
    Dim path As String                                            '試験仕様書の配置ディレクトリ
    Dim wbNum As Integer                                     '試験仕様書数
    Dim wbName As String                                     '試験仕様書のブック名
    Dim wsName As String                                     '試験仕様書のシート名
    Dim executingDate As String                            '実行日
    Dim variationKW As String                                'バリエーションエリアを探すためのキーワード
    Dim ws As Worksheets                                      '？
    Dim inputedexecutingDateCell As Range         '実行日が定義されているセル
    Dim toCellsInVariationRng As Range                'バリエーションエリアの左上のセル
    Dim variationRng As Range                              'バリエーションエリア
    Dim variationMaxNum As Integer                     'バリエーションの最大数
    Dim testCaseNum As Integer                           'テストケース数
    Dim toCellsInVariationRngRow As Integer       'バリエーションエリアの左上のセルの行番号
    Dim toCellsInVariationRngColumn As Integer  'バリエーションエリアの左上のセルの列番号
    Dim endCellsInVariationRng As Range             'バリエーションエリアの右下のセル
    Dim columnId As String                                    'セルの列番号(アルファベット)
    Dim aggregateTableName As String                '集計表の名前
    
    Application.ScreenUpdating = False
    'マクロブックのブック名、シート名、日付名、キーワード、試験仕様書数をprogressManageent.xlsmから取得
    macroWbName = Range("C3").Value
    macroWsName = Range("C4").Value
    If Range("B11").Value = "" Then
        wbNum = 1
    Else
        wbNum = Range(Workbooks(macroWbName).Worksheets(macroWsName).Range("B10"), Workbooks(macroWbName).Worksheets(macroWsName).Range("B10").End(xlDown)).Rows.count
    End If
    variationKW = Workbooks(macroWbName).Worksheets(macroWsName).Range("C6").Value
    executingDate = Workbooks(macroWbName).Worksheets(macroWsName).Range("C7").Value
    Debug.Print ("マクロのブック名：" & macroWbName)
    Debug.Print ("マクロのシート名：" & macroWsName)
    Debug.Print ("試験仕様書数：" & wbNum)
    Debug.Print ("日付名：" & executingDate)
    Debug.Print ("バリエーションを検索するためのキーワード：" & variationKW)
    
    Dim l As Integer
    Dim count As Long
    count = 0
    For l = 0 To wbNum - 1
L1:
        'ターゲットの試験仕様書名、シート名、進捗表のあるシート名を取得
        wbName = Workbooks(macroWbName).Worksheets(macroWsName).Cells(10 + l, 2).Value
        wsName = Workbooks(macroWbName).Worksheets(macroWsName).Cells(10 + l, 3).Value
        aggregateTableName = Workbooks(macroWbName).Worksheets(macroWsName).Cells(10 + l, 4).Value
        Debug.Print ("試験仕様書名：" & wbName)
        Debug.Print ("シート名：" & wsName)
        Debug.Print ("進捗表のあるシート名：" & aggregateTableName)
        
        'もし先ほどまでと試験仕様書が異なるならば新しく試験仕様書を開く
        If isSameTestingSpecification(macroWbName, macroWsName, l - 1) = False Then
            Call openTestingSpecification(getPath(macroWbName, macroWsName, "C5"), wbName)
        End If
        
        'シート内でバリエーションエリアを特定する
        Workbooks(wbName).Worksheets(wsName).Activate
        Set toCellsInVariationRng = findCells(variationKW, usingRng(wbName, wsName))
        
        'バリエーションエリアを特定できなかったときのエラーハンドリング
        If toCellsInVariationRng Is Nothing Then
            MsgBox "試験仕様書：" & wbName & "　シート名：" & wsName & "のバリエーションエリアの特定に失敗。skipします。"
            l = l + 1
            GoTo L1
        End If
        
        '特定できた場合は処理を続行
        '#30対応でコメントアウト
        'Set toCellsInVariationRng = Cells(findCells(variationKW, usingRng(wbName, wsName)).Row + 1, findCells(variationKW, usingRng(wbName, wsName)).Column + 1)
        'Debug.Print ("バリエーションの範囲の左上：" & toCellsInVariationRng.Address(RowAbsolute:=False, ColumnAbsolute:=False))
        Set variationRng = findArea(toCellsInVariationRng)
        Debug.Print ("バリエーションの範囲：" & variationRng.Address(RowAbsolute:=False, ColumnAbsolute:=False))
        
        '#30対応でコメントアウト
        '特定したバリエーションエリアの範囲からバリエーション最大数を取得
        'variationMaxNum = (variationRng.Rows.count - 1) / testCaseNum
        'Debug.Print ("バリエーションの最大数：" & variationMaxNum)
        
        '特定したバリエーションエリアの範囲からテストケース数を取得
        'testCaseNum = toCellsInVariationRng.End(xlDown).Value
        'Debug.Print ("テストケース数：" & testCaseNum)
        
        '実行日が入力されているセルの位置を取得
        'Set inputedexecutingDateCell = findCells(executingDate, usingRng(wbName, wsName))
        'Debug.Print ("実行日が入力されているセルの位置：" & inputedexecutingDateCell.Address(RowAbsolute:=False, ColumnAbsolute:=False))
        
        '実行日が入力されているセルの位置の特定に失敗した時のエラーハンドリング
        'If inputedexecutingDateCell Is Nothing Then
        '    MsgBox "試験仕様書：" & wbName & "　シート名：" & wsName & "の実行日が入力されているセルの位置の特定に失敗しました。skipします"
        '    l = l + 1
        '    GoTo L1
        'End If
        
        '集計表に書き込みを開始
        Dim i As Integer
            For i = 0 To (variationRng.Rows.count - 1) - 1
                Call writingInAggregateTable(i, count, aggregateTableName, wsName, toCellsInVariationRng)
            Next i
        count = count + variationRng.Rows.count - 1
        Debug.Print ("合計数：" & count)
        
        'もし試験仕様書が異なるのであれば試験仕様書を保存して閉じる
        If isSameTestingSpecification(macroWbName, macroWsName, l) = False Then
            Call closeTestingSpecification(wbName)
            count = 0
        End If
    Next l
    Application.ScreenUpdating = True
    MsgBox "進捗管理表の作成が完了しました。"
End Sub

Sub addAggregateTable()
    '各試験仕様書に進捗管理表のシートを追加する。
    Dim wbName As String
    Dim wsName As String
    Dim path As String
    Dim addWsName As String
    Dim testingSpecificationName As String
    
    Application.ScreenUpdating = False
    
    wbName = ActiveSheet.Range("C3").Value
    wsName = ActiveSheet.Range("C4").Value
    path = getPath(wbName, wsName, "C5")
    addWsName = Range("C6").Value
    Debug.Print ("マクロのブック名：" & wbName)
    Debug.Print ("マクロのシート名：" & wsName)
    Debug.Print ("シート追加するブックが配置されているパス：" & path)
    Debug.Print ("シート追加するブック数：" & addWsCount)
    Debug.Print ("シート追加するブック名：" & testingSpecificationName)
    
    'シート追加対象の試験仕様書を取得する
    testingSpecificationName = Dir(path & "*.xls*")
    Debug.Print ("試験仕様書名：" & testingSpecificationName)
    
    'ファイルがなかったときのエラーハンドリング
    If testingSpecificationName = "" Then
        MsgBox "試験仕様書が存在しません。"
        Exit Sub
    End If
    
    'ファイル名を順次開く
    Do While testingSpecificationName <> ""
        Debug.Print ("開く試験仕様書名：" & testingSpecificationName)
        Call openTestingSpecification(path, testingSpecificationName)
        
        '//追加するシートと同名のシートが存在したときは削除する
        If isSheetDuplicationCheck(addWsName) = True Then
            '//ブックが共有か排他的かチェック。共有であれば排他的にする。
            If Workbooks(testingSpecificationName).MultiUserEditing = True Then
               '//共有を外す
                Workbooks(testingSpecificationName).UnprotectSharing
                Workbooks(testingSpecificationName).ExclusiveAccess
            End If
            '//シート削除
            Application.DisplayAlerts = False
            Workbooks(testingSpecificationName).Worksheets(addWsName).Delete
            Application.DisplayAlerts = True
            '//共有にする
            '//Workbooks(testingSpecificationName).ProtectSharing
        End If
            
        'シートを追加
        Call addNewWorksheets(testingSpecificationName, addWsName)
        Call makeInstrumentationArea(testingSpecificationName, addWsName)
        '追加済みの試験仕様書を閉じる。
        Call closeTestingSpecification(testingSpecificationName)
        testingSpecificationName = Dir()
    Loop
    
    Application.ScreenUpdating = True
    
    MsgBox "シートの追加が完了しました。"
End Sub

Sub transcription()
    '各試験試験仕様書の進捗管理表シートからマスターへ転記する
    Dim macroWb As String
    Dim macroWs As String
    Dim pathOfTestingSpecification As String
    Dim pathOfVariationMngWb As String
    Dim variationMngWb As String
    Dim variationMngWs As String
    Dim testingSpecification As String
    Dim aggregateTableName As String
    Dim timeStampCells As String
    Dim searchWord As String
    Dim i As Long
    Dim wsName As String
    Dim caseNum As String
    Dim executingDate As String
    Dim result As String
    Dim faultNum As String
    Dim tester As String
    Dim executingKubun As String
    Dim achievement As String
    Dim remaining As String
    Dim sum As String
    Dim num As Long
    Dim overWritingFlag As Boolean
    Dim rslt As VbMsgBoxResult
    
    Application.ScreenUpdating = False
    
    macroWb = ActiveSheet.Range("C3")
    macroWs = ActiveSheet.Range("C4")
    pathOfTestingSpecification = getPath(macroWb, macroWs, "C5")
    pathOfVariationMngWb = getPath(macroWb, macroWs, "C6")
    variationMngWb = ActiveSheet.Range("C7")
    variationMngWs = ActiveSheet.Range("C8")
    aggregateTableName = ActiveSheet.Range("C9")
    timeStampCells = ActiveSheet.Range("C10")
    
    '前日分を上書きするか確認する
    rslt = MsgBox("前日分を上書きしますか？", Buttons:=vbYesNo)
    If rslt = vbYes Then
        overWritingFlag = True
    Else
        overWritingFlag = False
    End If
    
    '進捗管理表_バリエーション.xlsxを開く
    Call openTestingSpecification(pathOfVariationMngWb, variationMngWb)
    Call checkFilterModeStatus(Worksheets(variationMngWs))
    
    'overWritingFlag=trueのとき、データを前日項目に移動
    If overWritingFlag = True Then
        Workbooks(variationMngWb).Worksheets(variationMngWs).Range("E9:L10000").Copy Range("M9")
    End If
    
    'ここから転記を開始する
    '集計対象の試験仕様書名を取得
    testingSpecification = Dir(pathOfTestingSpecification & "*.xls*")
    Debug.Print ("試験仕様書名：" & testingSpecification)
    
    'ファイルがなかったときのエラーハンドリング
    If testingSpecification = "" Then
        MsgBox "Excelファイルがありません。"
        Exit Sub
    End If
    
    'ファイル名を順次開く
    Do While testingSpecification <> ""
        Debug.Print ("開く試験仕様書名：" & testingSpecification)
        
        Call openTestingSpecification(pathOfTestingSpecification, testingSpecification)
        
        '転記するテストケースの数を計算する
        Workbooks(testingSpecification).Worksheets(aggregateTableName).Activate
        num = Range(Workbooks(testingSpecification).Worksheets(aggregateTableName).Range("B4"), Workbooks(testingSpecification).Worksheets(aggregateTableName).Range("B4").End(xlDown)).Rows.count
        Debug.Print ("ケース数：" & num)
        
        For i = 0 To num - 1
L3:
            '転記に必要な情報を取得
            wsName = Cells(4 + i, 2).Value
            caseNum = Cells(4 + i, 3).Value
            executingDate = Cells(4 + i, 4).Value
            result = Cells(4 + i, 5).Value
            faultNum = Cells(4 + i, 6).Value
            tester = Cells(4 + i, 7).Value
            executingKubun = Cells(4 + i, 8).Value
            achievement = Cells(4 + i, 9).Value
            remaining = Cells(4 + i, 10).Value
            sum = Cells(4 + i, 11).Value
            
            '転記先のセル位置を取得
            searchWord = testingSpecification & wsName & caseNum
            Debug.Print ("検索ワード" & searchWord)
            'Workbooks(variationMngWb).Worksheets(variationMngWs).Activate
            Set copyTarget = findCells(searchWord, Workbooks(variationMngWb).Worksheets(variationMngWs).Range("AB:AB"))
            
            '転記先のセル位置を取得できなかったときのエラーハンドリング
            If copyTarget Is Nothing Then
                MsgBox "試験仕様書名：" & testingSpecification & "シート名：" & wsName & "ケース番号：" & caseNum & "の転記先のセル位置の取得失敗。処理をskipします。"
                i = i + 1
                GoTo L3
            End If
            
            '転記
            Workbooks(variationMngWb).Worksheets(variationMngWs).Cells(copyTarget.Row, 5) = executingDate
            Workbooks(variationMngWb).Worksheets(variationMngWs).Cells(copyTarget.Row, 6) = result
            Workbooks(variationMngWb).Worksheets(variationMngWs).Cells(copyTarget.Row, 7) = faultNum
            Workbooks(variationMngWb).Worksheets(variationMngWs).Cells(copyTarget.Row, 8) = tester
            Workbooks(variationMngWb).Worksheets(variationMngWs).Cells(copyTarget.Row, 9) = executingKubun
            Workbooks(variationMngWb).Worksheets(variationMngWs).Cells(copyTarget.Row, 10) = achievement
            Workbooks(variationMngWb).Worksheets(variationMngWs).Cells(copyTarget.Row, 11) = remaining
            Workbooks(variationMngWb).Worksheets(variationMngWs).Cells(copyTarget.Row, 12) = sum
        Next i
        
        '転記が完了した試験仕様書を閉じる
        closeTestingSpecification (testingSpecification)
        
        testingSpecification = Dir()
    Loop
    
    'タイムスタンプを記入
    Dim timeStamp As String
    timeStamp = Format(Now, "yyyy/mm/dd/　hh:mm:ss")
    Workbooks(variationMngWb).Worksheets(variationMngWs).Range(timeStampCells).Value = "更新日時：" & timeStamp
    
    Application.ScreenUpdating = False
    
    MsgBox "転記完了"
End Sub

Function getPath(ByVal wbName As String, ByVal wsName As String, ByVal rngAddress As String) As String
    getPath = Workbooks(wbName).Worksheets(wsName).Range(rngAddress).Value
End Function

Function openTestingSpecification(ByVal path As String, ByVal wbName As String)
    Workbooks.Open (path & wbName)
    Workbooks(wbName).Activate
End Function

Function closeTestingSpecification(ByVal wbName As String)
    Application.DisplayAlerts = False
    Debug.Print ("閉じる試験仕様書名：" & wbName)
    Workbooks(wbName).Save
    Workbooks(wbName).Close
    Application.DisplayAlerts = True
End Function

Function columnNumberToAlphabet(ByVal i As Long) As String
    Dim alpha As String
    alpha = Cells(1, i).Address(True, False)
    columnNumberToAlphabet = Left(alpha, InStr(alpha, "$") - 1)
End Function

Function findCells(ByVal keyword As String, ByVal usedRng As Range) As Range
    Set findCells = usedRng.Find(What:=keyword, LookIn:=xlValues, LookAt:=xlWhole)
End Function

Function usingRng(ByVal wbName As String, ByVal wsName As String) As Range
    Set usingRng = Workbooks(wbName).Worksheets(wsName).usedRange
End Function

Function findArea(ByVal rng As Range) As Range
     Dim columnNum As Integer
     Dim rowNum As Integer
     columnNum = Range(rng, rng.End(xlToRight)).Columns.count
     rowNum = Range(rng, rng.End(xlDown)).Rows.count
     Set findArea = Range(rng, Cells(rng.Row + rowNum - 1, rng.Column + columnNum - 1))
End Function

Function writingInAggregateTable(ByVal i As Integer, _
                                                        ByVal count As Long, _
                                                        ByVal aggregateTableName As String, _
                                                        ByVal wsName As String, _
                                                        ByVal toCellsInVariationRng As Range)
        '//シート名を該当セルに入力
        Worksheets(aggregateTableName).Cells(4 + i + count, 2) = wsName
        '//セルの列番号を数字から英語に変換
        columnId = columnNumberToAlphabet(toCellsInVariationRng.Column)
        '//ケース番号を該当セルに入力
        Worksheets(aggregateTableName).Cells(4 + i + count, 3) = "='" & wsName & "'!" & columnNumberToAlphabet(toCellsInVariationRng.Column) & toCellsInVariationRng.Row + 1 + i
        '//観点を該当セルに入力
        Worksheets(aggregateTableName).Cells(4 + i + count, 4) = "='" & wsName & "'!" & columnNumberToAlphabet(toCellsInVariationRng.Column + 1) & toCellsInVariationRng.Row + 1 + i
        '//実行日を表示する数式を該当セルに入力
        Worksheets(aggregateTableName).Cells(4 + i + count, 5) = "='" & wsName & "'!" & columnNumberToAlphabet(toCellsInVariationRng.Column + 2) & toCellsInVariationRng.Row + 1 + i
        '//実行者を表示する数式を該当セルに入力
        Worksheets(aggregateTableName).Cells(4 + i + count, 6) = "='" & wsName & "'!" & columnNumberToAlphabet(toCellsInVariationRng.Column + 3) & toCellsInVariationRng.Row + 1 + i
        '//実行結果を表示する数式を該当セルに入力
        Worksheets(aggregateTableName).Cells(4 + i + count, 7) = "='" & wsName & "'!" & columnNumberToAlphabet(toCellsInVariationRng.Column + 4) & toCellsInVariationRng.Row + 1 + i
        '//障害番号を表示する数式を該当セルに入力
        Worksheets(aggregateTableName).Cells(4 + i + count, 8) = "='" & wsName & "'!" & columnNumberToAlphabet(toCellsInVariationRng.Column + 5) & toCellsInVariationRng.Row + 1 + i
        '//実行区分/備考を表示する数式を該当セルに入力
        Worksheets(aggregateTableName).Cells(4 + i + count, 9) = "='" & wsName & "'!" & columnNumberToAlphabet(toCellsInVariationRng.Column + 6) & toCellsInVariationRng.Row + 1 + i
End Function

Function isSameTestingSpecification(ByVal wb As String, ByVal ws As String, ByVal i As Integer) As Boolean
    Dim result As Boolean
    
    If Workbooks(wb).Sheets(ws).Cells(10 + i, 2).Value = Workbooks(wb).Sheets(ws).Cells(10 + i + 1, 2).Value Then
        result = True
    Else
        result = False
    End If
    
    isSameTestingSpecification = result
End Function

Function addNewWorksheets(ByVal wbName As String, ByVal wsName As String)
    Dim newWorkSheet As Worksheet
    '//試験仕様書の左から二番目にシートを追加する
    Set newWorkSheet = Worksheets.Add(Before:=Worksheets(2))
    newWorkSheet.Name = wsName
    Debug.Print ("追加")
    Workbooks(wbName).Worksheets(wsName).Range("B3") = "シート名"
    Workbooks(wbName).Worksheets(wsName).Range("C3") = "ケース番号"
    Workbooks(wbName).Worksheets(wsName).Range("D3") = "観点"
    Workbooks(wbName).Worksheets(wsName).Range("E3") = "実行日"
    Workbooks(wbName).Worksheets(wsName).Range("F3") = "実行者"
    Workbooks(wbName).Worksheets(wsName).Range("G3") = "実行結果"
    Workbooks(wbName).Worksheets(wsName).Range("H3") = "障害番号"
    Workbooks(wbName).Worksheets(wsName).Range("I3") = "実行区分/備考"
End Function

Function makeInstrumentationArea(ByVal wbName As String, ByVal wsName As String)
        Workbooks(wbName).Worksheets(wsName).Range("L2").Interior.ColorIndex = 24
        Workbooks(wbName).Worksheets(wsName).Range("L3") = "打鍵数"
        Workbooks(wbName).Worksheets(wsName).Range("L4") = "=COUNTIF(E:E,L2)"
        Workbooks(wbName).Worksheets(wsName).Range("M2") = "←検索した打鍵日を記入"
        Workbooks(wbName).Worksheets(wsName).Range("M3") = "打鍵OK数"
        Workbooks(wbName).Worksheets(wsName).Range("M4") = "=COUNTIFS(E:E,L2,G:G," & """" & "OK" & """" & ")"
        Workbooks(wbName).Worksheets(wsName).Range("N3") = "成功率"
        Workbooks(wbName).Worksheets(wsName).Range("N4") = "=IFERROR(TEXT(M4/L4," & """" & "0.0%" & """" & "),0)"
End Function

Function isSheetDuplicationCheck(ByVal wsName As String) As Boolean
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.Name = wsName Then isSheetDuplicationCheck = True
    Next ws
End Function

Function checkFilterModeStatus(ByVal ws As Worksheet)
    'オートフィルタ未設定時は処理を抜ける
    If (ws.AutoFilterMode = False) Then
        Exit Function
    End If
    
    '絞り込みされている場合
    If (ws.AutoFilter.FilterMode = True) Then
        ws.AutoFilterMode = False
    End If
End Function
