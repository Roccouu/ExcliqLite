VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppExcliq_current"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'============================= CLASS MODULES STRUCTURE =============================='
' OBJECT VARIABLES (GLOOBJ_) --------------------------------------------------------'
Private EHGLOBAL As AppErrorHandler
'Ribbon Controller
Private CTR_RIBBON As RibbonExcliq_current

' VARIABLES (GLOStr_) ---------------------------------------------------------------'
'Private VERSION As String

' CONSTANTS ERRORS AND ERROR MESSAGES (GLOSTR_ERR_) ---------------------------------'
Private Const CUSTOM_ERROR As Long = VBA.vbObjectError + 513








' PROPERTIES ------------------------------------------------------------------------'
'Public Property Get APPVERSION() As String
'  APPVERSION = VERSION
'End Property







' CONSTRUCTOR/DESTRUCTOR ------------------------------------------------------------'
Private Sub Class_Initialize()
  Set CTR_RIBBON = New RibbonExcliq_current
  Set EHGLOBAL = New AppErrorHandler
End Sub

Private Sub Class_Terminate()
  Set CTR_RIBBON = Nothing
  Set EHGLOBAL = Nothing
End Sub






' CORE METHODS LIST PRIVATE ---------------------------------------------------------'
Public Function AppRibbon_CommonsStatusGetter(ByVal ControlId As String) As Variant
  
  Dim CTR As ControllerApp_current
  
  On Error GoTo EH
  If CTR_RIBBON Is Nothing Then GoTo EH
  #If Debugging Then
    Debug.Print "Segundo", ControlId
  #End If
  
  'Ribbon controls main enabling status setter
  Call Debug_assert(Not EHGLOBAL Is Nothing)
  Set CTR = New ControllerApp_current
  Set CTR.ErrorHandler = EHGLOBAL
  Let AppRibbon_CommonsStatusGetter = CTR.CtrlExecute(ControlId, eqlCtrGet)
  Set CTR = Nothing

EH:
  Set CTR = Nothing
  Call EHGLOBAL.ErrorHandlerDisplay("APPEQL::CommonsStatusGetter")

End Function

'RIBBON Action listeners
Public Sub AppRibbon_CommonsActionsExecutorControls(ByVal ControlId As String, Optional id As String, Optional Index As Integer)
  
  Dim CTR As ControllerApp_current
  
  Dim Action As String
  
  Const sFN As String = "RIBBON::ExecutorControls"
  
  'So, if EH is not previously registered, then create a new EH object
  'If RegisteredError Then Set eROR = New AppErrorHandler
  On Error GoTo EH
  If CTR_RIBBON Is Nothing Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "La aplicación debe ser cerrada debido a un error interno.")
  
  #If Debugging Then
    Call VBA.MsgBox(ControlId)
  #End If

  Set CTR = New ControllerApp_current
  Set CTR.ErrorHandler = EHGLOBAL
  Call CTR.CtrlExecute( _
    ControlId, _
    IIf(ControlId = "configtabletype" Or ControlId = "configtablestyle" Or ControlId = "configformulas" Or ControlId = "configtablerowshadows" Or ControlId = "configcellforselections", eqlCtrSet, eqlCtrShw), _
    IIf(ControlId = "configtabletype" Or ControlId = "configtablestyle", Index, id))
  Set CTR = Nothing

EH:
  Set CTR = Nothing
  Call EHGLOBAL.ErrorHandlerDisplay(sFN)

End Sub







' ASSETS ----------------------------------------------------------------------------'






' ERRORS MANNAGEMENT AND ASSERTSIONS ------------------------------------------------'
Private Sub Debug_assert(ByVal Cond As Boolean, Optional Origin As String = VBA.vbNullString)
  #If Debugging Then
    Debug.Assert Cond
    If Not Cond Then Debug.Print "Debugged from: " & Origin
  #End If
End Sub





' METHODS LIST PUBLIC ---------------------------------------------------------------'
'Ribbon
Public Sub AppInit(ByVal AppRibbon As IRibbonUI)
  If Not CTR_RIBBON Is Nothing Then Call CTR_RIBBON.RibbonCommonsInit(AppRibbon)
End Sub

Public Sub AppRibbonExecutorControls(ByVal ControlId As String, Optional id As String, Optional Index As Integer)
  Call AppRibbon_CommonsActionsExecutorControls(VBA.LCase(VBA.Mid(ControlId, 14)), id, Index)
End Sub

Public Function AppRibbonStatusSetter(ByVal ControlId As String, Optional VisibleEnable As Boolean = False) As Variant
  If Not CTR_RIBBON Is Nothing Then Let AppRibbonStatusSetter = CTR_RIBBON.RibbonCommonsStatusSetter(VBA.LCase(VBA.Mid(ControlId, 14)), VisibleEnable)
End Function

Public Function AppRibbonStatusGetter(ByVal ControlId As String) As Variant
  Let AppRibbonStatusGetter = AppRibbon_CommonsStatusGetter(VBA.LCase(VBA.Mid(ControlId, 14)))
End Function






'UDF'S METHODS
'Category Registration UDF's mehotd
'This method has a rare way. Excel says: "Do not can change a macro in a hidde book" (insane, when values are passed from an array direction),
'so, custom category does not created, but all UDF's are stored in a new category called "Commands".
'This is the reason of "On Error Resume Next" and the MacroNameS & MacroDescriptionS data strings.
'Oh MS!, ...just VSCode is a marvel in your world!.
Public Sub CallUDFsRegistration()
'  '"EQL_REGRESION_LINEAL","EQL_REGRESION_LINEAL_a","EQL_REGRESION_LINEAL_b","EQL_REGRESION_LINEAL_r","EQL_REGRESION_LINEAL_r2","EQL_REGRESION_LINEAL_Se", _
'  "EQL_REGRESION_LINEAL_n","EQL_BALMET", "EQL_BALMET_Rec","EQL_BALMET_Rc","EQL_BALMET_COMPLEJOS","EQL_BALMET_COMPLEJOS_Rec", _
'  "EQL_BALMET_ECONOMICO_Mineral","EQL_BALMET_ECONOMICO_LeyCabeza","EQL_BALMET_ECONOMICO_Concentrado","EQL_BALMET_ECONOMICO_LeyConcentrado", _
'  "EQL_NUMERO_A_TEXTO"
'  Dim MacroName() As Variant
'  Dim MacroDescription() As Variant
'  Dim MacroParameters() As Variant
'  Dim ECategory As String
'  Dim MacroNameS As String
'  Dim MacroDescriptionS As String
'  Dim i As Single
'
'  Set CV = New AppResVersioning
'  If CV.VersionVerify() Then
'
'    On Error Resume Next
'    Let MacroName = Array( _
'      "EQL_REGRESION_LINEAL", _
'      "EQL_REGRESION_LINEAL_a", _
'      "EQL_REGRESION_LINEAL_b", _
'      "EQL_REGRESION_LINEAL_r", _
'      "EQL_REGRESION_LINEAL_r2", _
'      "EQL_REGRESION_LINEAL_Se", _
'      "EQL_REGRESION_LINEAL_n", _
'      "EQL_BALMET", _
'      "EQL_BALMET_Rec", _
'      "EQL_BALMET_Rc", _
'      "EQL_BALMET_COMPLEJOS", _
'      "EQL_BALMET_COMPLEJOS_Rec", _
'      "EQL_BALMET_ECONOMICO_Mineral", _
'      "EQL_BALMET_ECONOMICO_LeyCabeza", _
'      "EQL_BALMET_ECONOMICO_Concentrado", _
'      "EQL_BALMET_ECONOMICO_LeyConcentrado", _
'      "EQL_NUMERO_A_TEXTO")
'
'    Let MacroDescription = Array( _
'      "Obtiene la regresión lineal de varios datos.", "Obtiene el valor del Coeficiente a de una regresión.", "Obtiene el valor del Coeficiente b de una regresión.", "Obtiene el valor del Coeficiente de correlación.", "Obtiene el valor del Coeficiente de determinación.", "Obtiene el valor del Error estándar.", _
'      "Obtiene el total de datos que intervienen en la regresión.", "Obtiene el peso del concentrado.", "Obtiene el Porcentaje de Recuperación.", "Obtiene el Radio de Concentración.", "Obtiene el Peso del concentrado a partir de un balance de tratamiento de minerales complejos.", "Obtiene el Porcentaje de Recuperación de un balance de tratamiento de minerales complejos.", _
'      "Obtiene el Peso de Alimentación.", "Obtiene la ley de Cabeza.", "Obtiene el Peso del Concentrado.", "Obtiene la Ley del Concentrado.", _
'      "Obtiene una expresión literal del número dado.")
'
'    Let ECategory = "ExcliqLite"
'
'    For i = 0 To UBound(MacroName)
'      If i = 0 Then
'        Let MacroParameters = Array("Requerido. Un rango de celdas con más de tres números.", "Opcional. Un rango de celdas con más de tres números.", "Opcional. Un número o una lista de números.")
'      ElseIf i = 1 Then Let MacroParameters = Array("Requerido. Un rango de celdas con más de tres números.", "Opcional. Un rango de celdas con más de tres números.")
'      ElseIf i = 2 Then Let MacroParameters = Array("Requerido. Un rango de celdas con más de tres números.", "Opcional. Un rango de celdas con más de tres números.")
'      ElseIf i = 3 Then Let MacroParameters = Array("Requerido. Un rango de celdas con más de tres números.", "Opcional. Un rango de celdas con más de tres números.")
'      ElseIf i = 4 Then Let MacroParameters = Array("Requerido. Un rango de celdas con más de tres números.", "Opcional. Un rango de celdas con más de tres números.")
'      ElseIf i = 5 Then Let MacroParameters = Array("Requerido. Un rango de celdas con más de tres números.", "Opcional. Un rango de celdas con más de tres números.")
'      ElseIf i = 6 Then Let MacroParameters = Array("Requerido. Un rango de celdas con más de tres números.", "Opcional. Un rango de celdas con más de tres números.")
'      ElseIf i = 7 Then Let MacroParameters = Array("Requerido. Un número, el peso de alimentación.", _
'                            "Requerido. Un número, Ley del primer mineral.", _
'                            "Opcional. Un número, Ley del segundo mineral.", _
'                            "Opcional. Un número, Ley del tercer mineral.", _
'                            "Opcional. Un número, Ley del cuarto mineral.", _
'                            "Opcional. Un número, corresponde a un concentrado (producto).")
'      ElseIf i = 8 Then Let MacroParameters = Array("Requerido. Un número, el peso de alimentación.", _
'                            "Requerido. Un número, Ley del primer mineral.", _
'                            "Opcional. Un número, Ley del segundo mineral.", _
'                            "Opcional. Un número, Ley del tercer mineral.", _
'                            "Opcional. Un número, Ley del cuarto mineral.", _
'                            "Opcional. Un número, corresponde a un concentrado (producto).")
'      ElseIf i = 9 Then Let MacroParameters = Array("Requerido. Un número, el peso de alimentación.", _
'                            "Requerido. Un número, Ley del primer mineral.", _
'                            "Opcional. Un número, Ley del segundo mineral.", _
'                            "Opcional. Un número, Ley del tercer mineral.", _
'                            "Opcional. Un número, Ley del cuarto mineral.", _
'                            "Opcional. Un número, corresponde a un concentrado (producto).")
'      ElseIf i = 10 Then Let MacroParameters = Array("Requerido. Un número, Ley de concentrado", "Requerido. Un número. Peso de concentrado.", "Opcional. 0 para DM, 1 para g/T")
'      ElseIf i = 11 Then Let MacroParameters = Array("Requerido. Un número, Ley de alimentación.", "Requerido. Un número. Peso de alimentación.", "Requerido. Un número. Ley del concentrado", "Requerido. Un número. Peso del concentrado")
'      ElseIf i = 12 Then Let MacroParameters = Array("Requerido. Un número, Peso del concentrado.", "Requerido. Un número. Ley del concentrado.", "Requerido. Un número. Ley de cabeza", "Opcional. Un número. Porcentaje de Recuperación.")
'      ElseIf i = 13 Then Let MacroParameters = Array("Requerido. Un número, Peso del concentrado.", "Requerido. Un número. Ley del concentrado.", "Requerido. Un número. Peso de alimentación.", "Opcional. Un número. Porcentaje de Recuperación.")
'      ElseIf i = 14 Then Let MacroParameters = Array("Requerido. Un número, Peso de alimentación.", "Requerido. Un número. Ley de cabeza.", "Requerido. Un número. Ley del concentrado.", "Opcional. Un número. Porcentaje de Recuperación.")
'      ElseIf i = 15 Then Let MacroParameters = Array("Requerido. Un número, Peso de alimentación.", "Requerido. Un número. Ley de cabeza.", "Requerido. Un número. Peso del concentrado.", "Opcional. Un número. Porcentaje de Recuperación.")
'      Else
'        Let MacroParameters = Array("Requerido. Un número.", "Opcional. Un texto.", "Opcional. Un texto.", "Opcional. Un número. 0 ó 1.")
'      End If
'
'      Let MacroNameS = MacroName(i)
'      Let MacroDescriptionS = MacroDescription(i)
'      'Application.MacroOptions MacroName(i), MacroDescription(i), ECategory:=ECategory, argumentdescriptions:=MacroParameters
'      Call Application.MacroOptions( _
'        macro:=MacroNameS, _
'        Description:=MacroDescriptionS, _
'        category:=ECategory, _
'        ArgumentDescriptions:=MacroParameters)
'
'      Let MacroNameS = VBA.vbNullString
'      Let MacroDescriptionS = VBA.vbNullString
'      Erase MacroParameters
'    Next i
'
'  Else
'    'Low Version, close thisWorkbook
'    Call Application.ThisWorkbook.Close(SaveChanges:=False)
'  End If
'
'  Erase MacroParameters
'  Erase MacroName
'  Erase MacroDescription
'  Set CV = Nothing
End Sub

'Linear Regression UDF'S. Works with Object instances from respectively App classes
Public Function LinearRegression(ByVal AStrRequester As String, ByVal AVarY As Variant, ByVal AVarX As Variant, ByVal AVarExtrapolacion As Variant) As Variant
  
  Dim CTR As ControllerApp_current
  Set CTR = New ControllerApp_current
  Set CTR.ErrorHandler = EHGLOBAL
  Let LinearRegression = CTR.CtrlExecuteUDF("regressionudf", Array(AStrRequester, AVarY, AVarX, AVarExtrapolacion))
  Set CTR = Nothing

End Function

'Balmet
Public Function Balmet( _
  ByVal AStrRequester As String, _
  ByVal AVarFeed As Variant, _
  ByVal AVarGrad As Variant, _
  ByVal AVarGradCx As Variant, _
  ByVal AVarName As Variant, _
  ByVal AVarUnit As Variant, _
  ByVal AMethod As EQLBAL_ENU_METHOD, _
  ByVal AResult As EQLBAL_ENU_RESULT, _
  ByVal ADirection As EQLBAL_ENU_RESULTDIRECTION, _
  ByVal ABooAllResult As Boolean) As Variant
  
  Dim CTR As ControllerApp_current
  Set CTR = New ControllerApp_current
  Set CTR.ErrorHandler = EHGLOBAL
  Let Balmet = CTR.CtrlExecuteUDF(AStrRequester, Array(AVarFeed, AVarGrad, AVarGradCx, AVarName, AVarUnit, AMethod, AResult, ADirection, ABooAllResult))
  Set CTR = Nothing

End Function

Public Function BalmetEco( _
  ByVal AStrRequester As String, _
  ByVal AVarFeed As Variant, _
  ByVal AVarWeig As Variant, _
  ByVal AVarGrad As Variant, _
  ByVal AVarRecs As Variant, _
  ByVal AVarGradCx As Variant, _
  ByVal AVarRecsCx As Variant, _
  ByVal AVarUnit As Variant, _
  ByVal AEnuTypeBalEco As EQLBAL_ENU_RESULTECO, _
  ByVal AEnuTypeResult As EQLBAL_ENU_TYPERESULT, _
  ByVal AEnuDirectionBalEco As EQLBAL_ENU_RESULTDIRECTION) As Variant

  Dim CTR As ControllerApp_current
  Set CTR = New ControllerApp_current
  Set CTR.ErrorHandler = EHGLOBAL
  Let BalmetEco = CTR.CtrlExecuteUDF(AStrRequester, Array(AVarFeed, AVarWeig, AVarGrad, AVarRecs, AVarGradCx, AVarRecsCx, AVarUnit, AEnuTypeBalEco, AEnuTypeResult, AEnuDirectionBalEco))
  Set CTR = Nothing

End Function

'Projections
Public Function Projections(ByVal AVarCashflow As Variant) As Variant
  
  Dim CTR As ControllerApp_current
  Set CTR = New ControllerApp_current
  Set CTR.ErrorHandler = EHGLOBAL
  Let Projections = CTR.CtrlExecuteUDF(VBA.vbNullString, AVarCashflow) 'Net flow cash
  Set CTR = Nothing

End Function


'Miscellaneous
Public Function NumberToStringGet(ByVal VarNum As Variant, ByVal StrCursingle As String, ByVal StrCurmultiple As String, ByVal BooLiteralcents As Boolean) As Variant
  
  Dim CTR As ControllerApp_current
  Set CTR = New ControllerApp_current
  Let NumberToStringGet = CTR.CtrlExecuteUDF(VBA.vbNullString, Array(VarNum, StrCursingle, StrCurmultiple, BooLiteralcents))
  Set CTR = Nothing

End Function
' METHODS LIST PUBLIC ---------------------------------------------------------------'
'============================= CLASS MODULES STRUCTURE =============================='



