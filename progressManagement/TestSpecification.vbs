Option Explicit

'//試験仕様書名
Private testSpecificationName As String
'//進捗管理表名
Private progressChart As String
'//試験仕様書が配置されているパス
Private path As String
'// バリエーションエリアを特定するためのキーワード
Private variationAreaKeyWord As String
'//バリエーションエリア
Private variationArea As Range
'//ケース番号
Private caseNum As String
'//実行結果
Private result As String
'//実行日
Private executingDate As String
'//障害番号
Private faultNum As String
'//打鍵者
Private testerName As String
'//実行区分（手動 or 自動）
Private executingKubun As String
'//実績
Private Achievement As String
'//残件
Private remaining As String
'//進捗管理表名
Private progressMngName As String

Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()
    testSpecificationName = ""
    progressChart = ""
    path = ""
End Sub

Public Property Get getTestSpecificationName() As String
    getTestSpecificationName = testSpecificationName
End Property

Public Property Get getProgressChart() As String
    getProgressChart = progressChart
End Property

Public Property Get getPath() As String
    getPath = path
End Property

Public Property Get getVariationAreaKeyWord() As String
    variationAreaKeyWord = variationAreaKeyWord
End Property

Public Property Get getVariationArea() As Range
    getVariationArea = variationArea
End Property

Public Property Get getCaseNum() As String
    getCaseNum = caseNum
End Property

Public Property Get getResult() As String
    getResult = result
End Property

Public Property Get getExecutingDate() As String
    getExecutingDate = executingDate
End Property

Public Property Get getFaultNum() As String
    getFaultNum = faultNum
End Property

Public Property Get getTesterName() As String
    getTesterName = testerName
End Property

Public Property Get getExecutingKubun() As String
    getExecutingKubun = executingKubun
End Property

Public Property Get Achievement() As Stirng
    getAchievement = Achievement
End Property

Public Property Get getRemaining() As Stirng
    getRemaining = remaining
End Property

Public Property Get getProgressMngName() As String
    getProgressMngName = progressMngName
End Property
Public Property Let setTestSpecificationName(ByVal newTestSpecificationName As String)
    testSpecificationName = newTestSpecificationName
End Property

Public Property Let setProgressChart(ByVal newProgressChart As String)
    progressChart = newProgressChart
End Property

Public Property Let setPath(ByVal newPath As String)
    path = newPath
End Property

Public Property Let setVariationAreaKeyWord(ByVal newVariationAreaKeyWord As String)
    variationAreaKeyWord = newVariationAreaKeyWord
End Property

Public Property Set setVariationArea(ByVal newVariationArea As Range)
    variationArea = newVariationArea
End Property

Public Property Let setCaseNum(ByVal newCaseNum As String)
    caseNum = newCaseNum
End Property

Public Property Let setResult(ByVal newResult As String)
    result = newResult
End Property

Public Property Let setExecutingDate(ByVal newExecutingDate As String)
    executingDate = newExecutingDate
End Property

Public Property Let setFaultNum(ByVal newFaultNum As String)
    faultNum = newFaultNum
End Property

Public Property Let setTesterName(ByVal newTesterName As String)
    testerName = newTesterName
End Property

Public Property Let setExecutingKubun(ByVal newExecutingKubun As String)
    executingKubun = newExecutingKubun
End Property

Public Property Let setAchievement(ByVal newAchievement As String)
    achiecement = newAchievement
End Property

Public Property Let setRemaining(ByVal newRemainig As String)
    remaining = newRemaining
End Property

Public Property Let setProgressMngName(ByVal newProgressMngName As String)
    progressMngName = newProgressMngName
End Property

Public Function openTestSpecification()
    Workbook.Open (path & testSpecificationName)
    Workbook(testSpecificationName).Activate
End Function

Public Function closeTestSpecification()
    Application.DisplayAlerts = False
    Workbooks(testSpecificationName).Save
    Workbooks(testSpecificationName).Close
    Application.DisplayAlerts = True
End Function

Public Function isSheetDuplicationCheck(ByVal wsName As String) As Boolean
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.Name = wsName Then isSheetDuplicationCheck = True
    Next ws
End Function
