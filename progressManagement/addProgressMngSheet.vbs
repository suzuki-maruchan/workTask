Option Explicit

Sub addProgressMngSheet()
    Dim mcrWb As MacroWorkBook
    Set mcrWb = New MacroWorkBook
    Dim tstSpcfctn As TestSpecification
    Set tstSpcfctn = New TestSpecification
    '//進捗管理表シート名
    Dim addWsName As String
    
    Application.ScreenUpdating = False
    
    tstSpcfctn.setPath = Workbooks(mcrWb.getMacroWbName).Worksheets(mcrWb.getMacroWsName).Range("C2").Value
    addWsName = Workbooks(mcrWb.getMacroWbName).Worksheets(mcrWb.getMacroWsName).Range("C3").Value
    '//シート追加対象補試験仕様書を取得する
    tstSpcfctn.setTestSpecificationName = Dir(tstSpcfctn.getPath() & "*.xls*")
    '//取得した試験仕様書の件数が0件だったときのエラーハンドリング
    If "" = tstSpcfctn.getTestSpecificationName Then
        MsgBox "試験仕様書が" & tstSpcfctn.getPath() & "に存在しません"
        Exit Sub
    End If
    ''//試験仕様書のブックを順番に開く
    'ファイル名を順次開く
    Do While tstSpcfctn.getTestSpecificationName <> ""
        Call tstSpcfctn.openTestSpecification
        '//追加するシートと同名のシートが存在したときは削除する
        If tstSpcfctn.isSheetDuplicationCheck(addWsName) = True Then
            '//ブックが共有か排他的かチェック。共有であれば排他的にする。
            If Workbooks(tstSpcfctn.getTestSpecificationName()).MultiUserEditing = True Then
               '//共有を外す
                Workbooks(tstSpcfctn.getTestSpecificationName()).UnprotectSharing
                Workbooks(tstSpcfctn.getTestSpecificationName()).ExclusiveAccess
            End If
            '//シート削除
            Application.DisplayAlerts = False
            Workbooks(tstSpcfctn.getTestSpecificationName()).Worksheets(addWsName).Delete
            Application.DisplayAlerts = True
            '//共有にする
            '//Workbooks(tstSpcfctn.getTestSpecificationName()).ProtectSharing
        End If
        Call addNewWorksheets(tstSpcfctn.getTestSpecificationName(), addWsName)
        Call tstSpcfctn.closeTestSpecification
        tstSpcfctn.setTestSpecificationName = Dir()
    Loop
    
    Application.ScreenUpdating = True
    MsgBox "シートの追加が完了しました。"
End Sub

Public Function addNewWorksheets(ByVal wbName As String, ByVal wsName As String)
    Dim newWorkSheet As Worksheet
    '//試験仕様書の左から二番目にシートを追加する
    Set newWorkSheet = Worksheets.Add(Before:=Worksheets(2))
    newWorkSheet.Name = wsName
    Workbooks(wbName).Worksheets(wsName).Range("B3") = "シート名"
    Workbooks(wbName).Worksheets(wsName).Range("C3") = "ケース番号"
    Workbooks(wbName).Worksheets(wsName).Range("D3") = "実行日"
    Workbooks(wbName).Worksheets(wsName).Range("E3") = "実行結果"
    Workbooks(wbName).Worksheets(wsName).Range("F3") = "障害番号"
    Workbooks(wbName).Worksheets(wsName).Range("G3") = "実行者"
    Workbooks(wbName).Worksheets(wsName).Range("H3") = "実行区分"
    Workbooks(wbName).Worksheets(wsName).Range("I3") = "■の数"
    Workbooks(wbName).Worksheets(wsName).Range("J3") = "□の数"
    Workbooks(wbName).Worksheets(wsName).Range("K3") = "総数"
End Function
