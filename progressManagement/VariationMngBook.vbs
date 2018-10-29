Option Explicit

'//バリエーション管理ブック名
Private variationMngBookName As String
'//バリエーション管理ブックのシート名
Private variationMngSheetName As String
'//バリエーション管理ブックが配置されているパス
Private path As String
'//タイムスタンプを記入するセル
Private timeStampCells As Range

Public Property Get getVariationMngBookName() As String
    getVariationMngBookName = variationMngBookName
End Property

Public Property Get getVariationMngSheetName() As String
    getVariationMngSheetName = variationMngSheetName
End Property

Public Property Get getPath() As String
    getPath = path
End Property

Public Property Get getTimeStampCells() As Range
    getTimeStampCells = timeStampCells
End Property

Public Property Let setVariationMngBookName(ByVal newVariationMngBookName As String)
    variationMngBookName = newVariationMngBookName
End Property

Public Property Let setVariationMngSheetName(ByVal newVariationMngSheetName As String)
    variationMngSheetName = newVariationMngSheetName
End Property

Public Property Let setPath(ByVal newPath As String)
    path = newPath
End Property

Public Property Set setTimeStampCells(ByVal newTimeStampCells As Range)
    timeStampCells = newTimeStampCells
End Property
    
