Option Explicit

'//マクロのブック名
Public macroWbName As String
'//マクロのシート名
Public macroWsName As String
'//マクロが配置されているパス
Public path As String
    
Private Sub Class_Initialize()
    macroWbName = ActiveWorkbook.Name
    macroWsName = ActiveSheet.Name
End Sub

Private Sub Class_Terminate()
    macroWbName = ""
    macroWsName = ""
    path = ""
End Sub

Public Property Get getMacroWbName() As String
    getMacroWbName = macroWbName
End Property

Public Property Let setMacroWbName(ByVal newMacroWbName As String)
    macroWbName = newMacroWbName
End Property

Public Property Get getMacroWsName() As String
    getMacroWsName = macroWsName
End Property

Public Property Let setMacroWsName(ByVal newMacroWsName As String)
    macroWsName = newMacroWsName
End Property

Public Property Get getPath()
    getPath = path
End Property

Public Property Let setPath(ByVal newPath As String)
    path = newPath
End Property
