VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassEqlController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ===================== CLASS EXCLIQ LITE CONTROLLER STRUCTURE ===================== '
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



' ===================== CLASS EXCLIQ LITE CONTROLLER STRUCTURE ===================== '
' OBJECT VARIABLES (GLOOBJ_) ------------------------------------------------------- '
' PROPERTIES ----------------------------------------------------------------------- '
' CONSTRUCTOR/DESTRUCTOR ----------------------------------------------------------- '
' METHODS PRIVATE ------------------------------------------------------------------ '
' EVENT METHODS LIST (PRIVATE) ----------------------------------------------------- '
' VIE_RequestEvent
' CORE METHODS LIST (PRIVATE) ------------------------------------------------------ '
' Ctrl_get_object
' Ctrl_process_data
' Ctrl_get_data
' Ctrl_set_data
' Ctrl_Excecutor_processor
' Ctrl_execute_udf
' Ctrl_del_help_sheet
' ASSETS --------------------------------------------------------------------------- '
' Ctrl_asset_set_res_and_model
' ERROR ASSERTSIONS ---------------------------------------------------------------- '
' Debug_assert
' Debug_print
' CORE METHODS LIST (PUBLIC) ------------------------------------------------------- '
' CtrlExecute
' CtrlExecuteUDF
' ===================== CLASS EXCLIQ LITE CONTROLLER STRUCTURE ===================== '



' ===================== CLASS EXCLIQ LITE CONTROLLER STRUCTURE ===================== '
' DESCRIPTION: _
  1.  Two public methods: Udf(do not uses view), Exe(uses view) _
  2.  A private method creates the current object using the reference parameter _
  3.  Request de ans at current object and send the data to viewprint if requester is not an UDF _
  4.  That's all folks

' OBJECT VARIABLES (GLOOBJ_) ------------------------------------------------------- '
Private EHGLOBAL As ClassEqlErrorHandler
Private WithEvents VIE As ClassEqlView
Attribute VIE.VB_VarHelpID = -1



' PROPERTIES ----------------------------------------------------------------------- '
Public Property Set ErrorHandler(ByRef ObjEH As ClassEqlErrorHandler)

  Set EHGLOBAL = ObjEH

End Property



' CONSTRUCTOR/DESTRUCTOR ----------------------------------------------------------- '
Private Sub Class_Terminate()

  Set EHGLOBAL = Nothing
  Set VIE = Nothing

End Sub



' METHODS PRIVATE ------------------------------------------------------------------ '
' EVENT METHODS LIST (PRIVATE) ----------------------------------------------------- '
Private Sub VIE_RequestEvent(ByVal EStrRequester As String, ByRef EVarRequest As Variant, ByRef EVarResponse As Variant, ByRef EBooCancel As Boolean)

  On Error GoTo EH
  Let EBooCancel = Ctrl_process_data(EStrRequester, EVarRequest, EVarResponse)

EH:
  Call EHGLOBAL.ErrorHandlerDisplay("CONTROLLER::-RequestEvent")

End Sub

' CORE METHODS LIST (PRIVATE) ------------------------------------------------------ '
Private Function Ctrl_get_object(ByVal StrId As String) As ClassEqlControllerInterface

  Dim CtrCore As ClassEqlControllerInterface

  On Error GoTo EH
  Select Case StrId
    Case "version"
      Set CtrCore = New ClassEqlControllerVersioning
    Case "regression", "regressionselection", "regressionudf"
      Set CtrCore = New ClassEqlControllerRegression
    Case "balmet", "balmetselection", "balmeteco", "balmetudf", "balmetudfeco"
      Set CtrCore = New ClassEqlControllerBalmet
    Case "projection"
      Set CtrCore = New ClassEqlControllerProjections
    Case "oreinvoice"
      Set CtrCore = New ClassEqlControllerPurchaseMin
    Case "concentrateinvoice"
      Set CtrCore = New ClassEqlControllerPurchaseCnc
    Case "numtostrudf"
      Set CtrCore = New ClassEqlControllerNumToStr
    Case Else: GoTo EH
  End Select

  ' Return
  Set Ctrl_get_object = CtrCore

EH:
  Set CtrCore = Nothing
  Call EHGLOBAL.ErrorHandlerRaise("CONTROLLER::-get_object")

End Function

Private Function Ctrl_process_data(ByVal StrId As String, ByRef CVarRequest As Variant, ByRef CVarResponse As Variant) As Boolean 'TODO: TO EVENT HNDL

  Dim RES As ClassEqlResources
  Dim MDL As ClassEqlModel
  Dim ICC As ClassEqlControllerInterface '<--

  On Error GoTo EH
  ' Get the specific controller
  Set ICC = Ctrl_get_object(StrId)

  ' Initialize Resources and Models
  Call Ctrl_asset_set_res_and_model(RES, MDL)

  ' Set controller object dependencies
  Set ICC.ResourcesApp = RES

  ' Process the request
  #If Debugging Then
    Call VBA.MsgBox("Greetings from Controller!: " & StrId, vbExclamation, "EQL")
  #End If
  Let Ctrl_process_data = ICC.GetSolution(StrId, CVarRequest, CVarResponse)
  
EH:
  Set ICC = Nothing '<--
  Set MDL = Nothing
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise("CONTROLLER::-process_data")

End Function

Private Function Ctrl_get_data(ByVal CStrId As String, ByVal CMDL As ClassEqlModel) As Variant

  On Error GoTo EH

  ' Calls ModelApp and request some data
  #If Debugging Then
    Call Debug_assert(Not EHGLOBAL Is Nothing, "CONTROLLER::-getdata")
  #End If

  Select Case CStrId
    Case "configtabletype": Let Ctrl_get_data = CMDL.Tables
    Case "configtablestyle": Let Ctrl_get_data = CMDL.TableStyles
    Case "configformulas": Let Ctrl_get_data = CMDL.Formulas
    Case "configtablerowshadows": Let Ctrl_get_data = CMDL.TableRowStripes
    Case "configcellforselections": Let Ctrl_get_data = CMDL.TableSelectionsCell
  End Select

EH:
  Call EHGLOBAL.ErrorHandlerRaise("CONTROLLER::-getdata")

End Function

Private Sub Ctrl_set_data(ByVal CStrId As String, ByVal CVarData As Variant, ByVal CMDL As ClassEqlModel)
  
  ' Calls ModelApp and sets some data
  On Error GoTo EH
  
  #If Debugging Then
    Call Debug_assert(Not EHGLOBAL Is Nothing, "CONTROLLER::-getdata")
  #End If

  Select Case CStrId
    Case "configtabletype": Let CMDL.Tables = VBA.CByte(CVarData)
    Case "configtablestyle": Let CMDL.TableStyles = VBA.CByte(CVarData)
    Case "configformulas": Let CMDL.Formulas = VBA.CBool(CVarData)
    Case "configtablerowshadows": Let CMDL.TableRowStripes = VBA.CBool(CVarData)
    Case "configcellforselections": Let CMDL.TableSelectionsCell = VBA.CBool(CVarData)
  End Select
  
EH:
  Call EHGLOBAL.ErrorHandlerRaise("CONTROLLER::-setdata")

End Sub

Private Function Ctrl_Excecutor_processor(ByVal CStrId As String, CEnuAction As EQLCTR_ENU_TABLE_AS, ByVal CVarRequest As Variant) As Variant
  
  Dim RES As ClassEqlResources
  Dim MDL As ClassEqlModel
  Dim ICC As ClassEqlControllerInterface '<--

  On Error GoTo EH
  ' Initialize Resources and Models
  Call Ctrl_asset_set_res_and_model(RES, MDL, CJustRES:=(CEnuAction = eqlCtrShw))

  ' Execute tasks
  Select Case CEnuAction

    Case eqlCtrGet
      Let Ctrl_Excecutor_processor = Ctrl_get_data(CStrId, MDL)

    Case eqlCtrSet
      Call Ctrl_set_data(CStrId, CVarRequest, MDL)

    Case eqlCtrShw
      Set VIE = New ClassEqlView
      Set VIE.ResourcesApp = RES
      Call VIE.ViewUI(CStrId)
      Set VIE = Nothing

    Case eqlCtrHlp
      #If Debugging Then
        Call VBA.MsgBox("Help!: " & CStrId)
      #End If
      Call Ctrl_del_help_sheet(RES, VBA.Replace(CStrId, "close-", VBA.vbNullString), VBA.CStr(CVarRequest))

    Case eqlCtrVer
      #If Debugging Then
        Call VBA.MsgBox("Version!: " & CStrId)
      #End If
      ' Set the selected controller
      Set ICC = Ctrl_get_object(CStrId)
      Set ICC.ResourcesApp = RES
      ' Send Request to acquire a Response
      Let Ctrl_Excecutor_processor = ICC.GetSolution(CStrId, CVarRequest, Empty)

    Case Else: GoTo EH
  End Select

EH:
  Set ICC = Nothing
  Set VIE = Nothing
  Set MDL = Nothing
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise("CONTROLLER::Excecutor_processor")

End Function

Public Function Ctrl_execute_udf(ByVal CStrId As String, ByVal CVarRequest As Variant) As Variant

  Dim RES As ClassEqlResources
  Dim MDL As ClassEqlModel
  Dim ICC As ClassEqlControllerInterface '<--

  On Error GoTo EH
  ' Initialize Resources and Models
  Call Ctrl_asset_set_res_and_model(RES, MDL)

  ' Set the selected controller
  Set ICC = Ctrl_get_object(CStrId)
  Set ICC.ResourcesApp = RES

  ' Send Request to acquire a Response
  Let Ctrl_execute_udf = ICC.GetSolutionUDF(CStrId, CVarRequest)

EH:
  Set ICC = Nothing
  Set MDL = Nothing
  Set RES = Nothing
  If Not VBA.Err.Number = 0 Then Call Debug_print(VBA.Err.Description, VBA.Err.Number)

End Function

Private Sub Ctrl_del_help_sheet(ByVal RES As ClassEqlResources, ByVal CStrWbk As String, ByVal CStrSht As String)

  On Error GoTo EH

  ' Delete worksheet
  If Not CStrWbk = Application.ThisWorkbook.Name Then
    Call RES.ExcelDelSheet(RStrNameSheet:=CStrSht, RStrWbk:=CStrWbk)
  End If

EH:
  Call EHGLOBAL.ErrorHandlerRaise("CONTROLLER::-del_help_sheet")

End Sub



' ASSETS --------------------------------------------------------------------------- '
Private Sub Ctrl_asset_set_res_and_model( _
  ByRef CRES As ClassEqlResources, _
  ByRef CMDL As ClassEqlModel, _
  Optional CJustRES As Boolean = False)

  On Error GoTo EH
  ' Initialize Resources and Models
  Set CRES = New ClassEqlResources
  Set CRES.ErrorHandler = EHGLOBAL

  Set CMDL = New ClassEqlModel
  Set CMDL.ResourcesApp = CRES

  Let CRES.Formulas = CMDL.Formulas
  Let CRES.AppStrTitle = VBA.StrConv(CMDL.AppTitle, vbProperCase)

EH:
  If CJustRES Then Set CMDL = Nothing
  Call EHGLOBAL.ErrorHandlerRaise("CONTROLLER::-asset_set_res_and_model")

End Sub



' ERRORS MANNAGEMENT AND ASSERTSIONS ----------------------------------------------- '
Private Sub Debug_assert(ByVal Cond As Boolean, Optional Origin As String = VBA.vbNullString)

  #If Debugging Then
    Debug.Assert Cond
    If Not Cond Then Debug.Print "Debugged from: " & Origin
  #End If

End Sub

Private Sub Debug_print(ParamArray CPrintableData() As Variant)

  #If Debugging Then
    Debug.Print VBA.Join(CPrintableData, vbTab)
  #End If

End Sub
' METHODS PRIVATE ------------------------------------------------------------------ '



' CORE METHODS LIST (PUBLIC) ------------------------------------------------------- '
Public Function CtrlExecute(ByVal CStrId As String, CEnuAction As EQLCTR_ENU_TABLE_AS, Optional CVarRequest As Variant) As Variant

  On Error GoTo EH

  #If Debugging Then
    Call Debug_assert(Not CStrId = VBA.vbNullString, "Exe")
  #End If

  ' Send Request to acquire a Response
  Let CtrlExecute = Ctrl_Excecutor_processor(CStrId, CEnuAction, CVarRequest)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("CONTROLLER::Execute")

End Function

Public Function CtrlExecuteUDF(ByVal CStrId As String, ByVal CVarRequest As Variant) As Variant

  On Error GoTo EH
  ' Send Request to acquire a Response
  Let CtrlExecuteUDF = Ctrl_execute_udf(CStrId, CVarRequest)

EH:
  If Not VBA.Err.Number = 0 Then Call Debug_print(VBA.Err.Description, VBA.Err.Number)

End Function
' CORE METHODS LIST (PUBLIC) ------------------------------------------------------- '
' ===================== CLASS EXCLIQ LITE CONTROLLER STRUCTURE ===================== '




