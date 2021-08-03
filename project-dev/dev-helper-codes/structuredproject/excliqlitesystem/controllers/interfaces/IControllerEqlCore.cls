VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICoreController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Private EHGLOBAL As AppErrorHandler




'PROPERTIES
Public Property Set ErrorHandler(ByRef ObjEH As AppErrorHandler)
'  Set EHGLOBAL = Nothing
'  Set EHGLOBAL = ObjEH
End Property




'CONSTRUCTOR/DESTRUCTOR
Private Sub Class_Terminate()
  'Set EHGLOBAL = Nothing
End Sub




' ABSTRACT METHODS PUBLIC -----------------------------------------------------------'
' INTERFACE METHODS LIST (PUBLIC) ---------------------------------------------------'
Public Function GetSolution(ByVal CStrRequester As String, ByVal CVarRequest As Variant, ByRef CVarResponse As Variant) As Boolean
End Function

Public Function GetSolutionUDF(ByVal CStrRequester As String, ByVal CVarRequest As Variant) As Variant
End Function
' ABSTRACT METHODS PUBLIC -----------------------------------------------------------'

