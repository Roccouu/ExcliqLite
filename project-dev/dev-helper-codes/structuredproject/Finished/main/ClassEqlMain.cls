VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassEqlMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ========================= CLASS EXCLIQ LITE APP STRUCTURE ======================== '
' ============================== CREDITS AND LICENSE =============================== '
' LICENSE & ACKNOWLEDGMENTS
'
' MIT License
'
' Copyright (c) 2019 - 2021
' Roberto Carlos Romay Medina
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'
'
' Acknowledgments to
'  - StarUML v2.8.0
'  - Pencil v3.0.3
'  - VSCode v1.41.0
'  - InkScape v0.92.4
'  - Gimp v2.10.2
'  - Just Color Picker v4.6
'  - Office RibbonX Editor | Fernando Andreu v1.5.1.418
'  - MS Excel v1911
'  - VBA7.1 V1091
'  - Paul Kelly | https://excelmacromastery.com
'  - Andrew Gould | https://www.wiseowl.co.uk/
'  - David Asurmendi | www.davidasurmendi.blogspot.com
'  - Ismael Romero | www.excelforo.blogspot.com
'  - Sergio Alejandro Campos | www.exceleinfo.com
'  - Otto Javier González | www.excelymas.com"
'
'
' Roccou: I think RefEdits works very well with a good treatment.
' GitHub: https://github.com/roccouu/
' ExcliqLite's home page: https://roccouu.github.io/ExcliqLite/
' ExcliqLite's repo: https://github.com/Roccouu/ExcliqLite
' 2019, POTOSÍ - BOLÍVIA
' ============================== CREDITS AND LICENSE =============================== '


' ========================= CLASS EXCLIQ LITE APP STRUCTURE ======================== '
' VARIABLES (GLOStr_) -------------------------------------------------------------- '
' CONSTRUCTOR/DESTRUCTOR ----------------------------------------------------------- '
' CORE METHODS LIST (PRIVATE) ------------------------------------------------------ '
' AppRibbon_CommonsStatusGetter
' AppRibbon_CommonsActionsExecutorControls
' ERRORS MANNAGEMENT AND ASSERTSIONS ----------------------------------------------- '
' Debug_assert
' CORE METHODS LIST (PUBLIC) ------------------------------------------------------- '
' AppBegin
' AppInit
' AppRibbonExecutorControls
' AppRibbonStatusSetter
' AppRibbonStatusGetter
' EqlVersion
' LinearRegression
' Balmet
' BalmetEco
' Projections
' NumberToStringGet
' ======================== CLASS EXCLIQ LITE APP STRUCTURE ========================= '



' ========================= CLASS EXCLIQ LITE APP STRUCTURE ======================== '
' VARIABLES (GLOStr_) -------------------------------------------------------------- '
Private EHGLOBAL As ClassEqlErrorHandler
' Ribbon Controller
Private CTR_RIBBON As ClassEqlRibbon



' CONSTRUCTOR/DESTRUCTOR ----------------------------------------------------------- '
Private Sub Class_Initialize()

  Set EHGLOBAL = New ClassEqlErrorHandler
  Set CTR_RIBBON = New ClassEqlRibbon
  Set CTR_RIBBON.ErrorHandler = EHGLOBAL
  Let CTR_RIBBON.AppTitle = "Excliq Lite" ' FIXME: fill with name from db

End Sub

Private Sub Class_Terminate()

  Set CTR_RIBBON = Nothing
  Set EHGLOBAL = Nothing

End Sub




' CORE METHODS LIST (PRIVATE) ------------------------------------------------------ '
Private Function AppRibbon_CommonsStatusGetter(ByVal ControlId As String) As Variant

  Dim CTR As ClassEqlController

  On Error GoTo EH
  If CTR_RIBBON Is Nothing Then GoTo EH
  #If Debugging Then
    Debug.Print "Segundo", ControlId
  #End If

  ' Ribbon controls main enabling status setter
  #If Debugging Then
    Call Debug_assert(Not EHGLOBAL Is Nothing, "CGSS")
  #End If
  Set CTR = New ClassEqlController
  Set CTR.ErrorHandler = EHGLOBAL
  Let AppRibbon_CommonsStatusGetter = CTR.CtrlExecute(ControlId, eqlCtrGet)
  Set CTR = Nothing

EH:
  Set CTR = Nothing
  Call EHGLOBAL.ErrorHandlerDisplay("MAIN::CommonsStatusGetter")

End Function

' RIBBON Action listeners
Private Sub AppRibbon_CommonsActionsExecutorControls(ByVal ControlId As String, Optional id As String, Optional Index As Integer)

  Dim CTR As ClassEqlController

  Dim Action As String, StrAux0 As String

  Const sFN As String = "MAIN::ExecutorControls"

  ' So, if EH is not previously registered, then create a new EH object
  ' If RegisteredError Then Set eROR = New ClassEqlErrorHandler
  On Error GoTo EH
  If CTR_RIBBON Is Nothing Then Call VBA.Err.Raise(CUSTOM_ERROR_APP, sFN, "La aplicación debe ser cerrada debido a un error interno.")
  If CTR_RIBBON.RibbonCommonsNoWorkbooksOpenend() Then

    #If Debugging Then
      Call VBA.MsgBox(ControlId)
      Call Debug_assert(1 = 1)
    #End If

    Set CTR = New ClassEqlController
    Set CTR.ErrorHandler = EHGLOBAL
    Let StrAux0 = VBA.Left(ControlId, 6)
    Call CTR.CtrlExecute( _
      IIf(StrAux0 = "close-", ControlId, VBA.LCase(ControlId)), _
      IIf(StrAux0 = "config", eqlCtrSet, IIf(StrAux0 = "close-", eqlCtrHlp, eqlCtrShw)), _
      IIf(ControlId = "configtabletype" Or ControlId = "configtablestyle", Index, id))
    Set CTR = Nothing

  End If

EH:
  Set CTR = Nothing
  Call EHGLOBAL.ErrorHandlerDisplay(sFN)

End Sub




' ERRORS MANNAGEMENT AND ASSERTSIONS ----------------------------------------------- '
Private Sub Debug_assert(ByVal Cond As Boolean, Optional Origin As String = VBA.vbNullString)

  #If Debugging Then
    Debug.Assert Cond
    If Not Cond Then Debug.Print "Debugged from: " & Origin
  #End If

End Sub





' CORE METHODS LIST (PUBLIC) ------------------------------------------------------- '
Public Sub AppBegin()

  Dim CTR As ClassEqlController
  Dim BooAux0 As Boolean

  #If Debugging Then
    'Call Debug_assert(1 = 0)
  #End If

  With Application.ThisWorkbook

    ' Verify version
    Set CTR = New ClassEqlController
    Set CTR.ErrorHandler = EHGLOBAL
    Let BooAux0 = VBA.CBool(CTR.CtrlExecute("version", eqlCtrVer, Empty))
    Set CTR = Nothing

    ' Initialize or close system
    If Not BooAux0 Then Call .Close(SaveChanges:=False)

  End With

EH:
  Set CTR = Nothing
  Call EHGLOBAL.ErrorHandlerDisplay("MAIN")

End Sub

' Ribbon
Public Sub AppInit(ByVal AppRibbon As IRibbonUI)

  If Not CTR_RIBBON Is Nothing Then Call CTR_RIBBON.RibbonCommonsInit(AppRibbon) ' Set up ribbon

End Sub

Public Sub AppRibbonExecutorControls(ByVal ControlId As String, Optional id As String, Optional Index As Integer)

  Call AppRibbon_CommonsActionsExecutorControls(VBA.Mid(ControlId, 14), id, Index)

End Sub

Public Function AppRibbonStatusSetter(ByVal ControlId As String, Optional VisibleEnable As Boolean = False) As Variant

  If Not CTR_RIBBON Is Nothing Then Let AppRibbonStatusSetter = CTR_RIBBON.RibbonCommonsStatusSetter(VBA.LCase(VBA.Mid(ControlId, 14)), VisibleEnable)

End Function

Public Function AppRibbonStatusGetter(ByVal ControlId As String) As Variant

  Let AppRibbonStatusGetter = AppRibbon_CommonsStatusGetter(VBA.LCase(VBA.Mid(ControlId, 14)))

End Function


' App version
Public Function EqlVersion() As Variant

  Dim CTR As ClassEqlController
  Set CTR = New ClassEqlController
  Set CTR.ErrorHandler = EHGLOBAL
  Let EqlVersion = CTR.CtrlExecuteUDF("version", Empty)
  Set CTR = Nothing

End Function

' Linear Regression UDF'S. Works with Object instances from respectively App classes
Public Function LinearRegression(ByVal AStrRequester As String, ByVal AVarY As Variant, ByVal AVarX As Variant, ByVal AVarExtrapolacion As Variant) As Variant

  Dim CTR As ClassEqlController
  Set CTR = New ClassEqlController
  Set CTR.ErrorHandler = EHGLOBAL
  Let LinearRegression = CTR.CtrlExecuteUDF("regressionudf", Array(AStrRequester, AVarY, AVarX, AVarExtrapolacion))
  Set CTR = Nothing

End Function

' Balmet
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

  Dim CTR As ClassEqlController
  Set CTR = New ClassEqlController
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

  Dim CTR As ClassEqlController
  Set CTR = New ClassEqlController
  Set CTR.ErrorHandler = EHGLOBAL
  Let BalmetEco = CTR.CtrlExecuteUDF(AStrRequester, Array(AVarFeed, AVarWeig, AVarGrad, AVarRecs, AVarGradCx, AVarRecsCx, AVarUnit, AEnuTypeBalEco, AEnuTypeResult, AEnuDirectionBalEco))
  Set CTR = Nothing

End Function

' Projections
Public Function Projections(ByVal AVarCashflow As Variant) As Variant

  Dim CTR As ClassEqlController
  Set CTR = New ClassEqlController
  Set CTR.ErrorHandler = EHGLOBAL
  Let Projections = CTR.CtrlExecuteUDF(VBA.vbNullString, AVarCashflow) 'Net flow cash
  Set CTR = Nothing

End Function


' Miscellaneous
Public Function NumberToStringGet(ByVal VarNum As Variant, ByVal StrCursingle As String, ByVal StrCurmultiple As String, ByVal BooLiteralcents As Boolean) As Variant

  Dim CTR As ClassEqlController
  Set CTR = New ClassEqlController
  Set CTR.ErrorHandler = EHGLOBAL
  Let NumberToStringGet = CTR.CtrlExecuteUDF("numtostrudf", Array(VarNum, StrCursingle, StrCurmultiple, BooLiteralcents))
  Set CTR = Nothing

End Function
' CORE METHODS LIST (PUBLIC) ------------------------------------------------------- '
' ======================== CLASS EXCLIQ LITE APP STRUCTURE ========================= '



