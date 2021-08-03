VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControllerApp_current"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' OBJECT VARIABLES (GLOOBJ_) --------------------------------------------------------'
Private EHGLOBAL As AppErrorHandler
Private ICC As ICoreController
Private WithEvents VIE As ViewApp_current
Attribute VIE.VB_VarHelpID = -1

' VARIABLES (GLOStr_) ---------------------------------------------------------------'

' CUSTOM VARIABLES (GLOTyp|Enu_) ----------------------------------------------------'
Public Enum EQLCTR_ENU_TABLE_AS
  eqlCtrGet
  eqlCtrSet
  eqlCtrShw
End Enum

' CONSTANTS ERROR MESSAGES (GLOSTR_ERR_) --------------------------------------------'
Private Const CUSTOM_ERROR = VBA.vbObjectError + 518



' DESCRIPTION: _
  1.  Two public methods: Udf(do not uses view), Exe(uses view) _
  2.  A private method creates the current object using the reference parameter _
  3.  Request de ans at current object and send the data to viewprint if requester is not an UDF _
  4.  That's all folks








'PROPERTIES
Public Property Set ErrorHandler(ByRef ObjEH As AppErrorHandler)
  
  Set EHGLOBAL = Nothing
  Set EHGLOBAL = ObjEH

End Property







'CONSTRUCTOR/DESTRUCTOR
Private Sub Class_Terminate()
  
  Set EHGLOBAL = Nothing
  Set ICC = Nothing
  Set VIE = Nothing

End Sub







' METHODS PRIVATE -------------------------------------------------------------------'
' EVENT METHODS LIST (PRIVATE) ------------------------------------------------------'
Private Sub VIE_RequestEvent(ByVal EStrRequester As String, ByVal EVarRequest As Variant, EVarResponse As Variant, EBooCancel As Boolean)
  
  On Error GoTo EH
  Let EBooCancel = Ctrl_process_data(EStrRequester, EVarRequest, EVarResponse)
  
EH:
  Call EHGLOBAL.ErrorHandlerDisplay("CONTROLLER::-RequestEvent")

End Sub


' CORE METHODS LIST (PRIVATE) -------------------------------------------------------'
Private Function Ctrl_get_object(ByVal StrId As String) As ICoreController
  
  Dim CtrCore As ICoreController
  
  On Error GoTo EH
  Select Case StrId
    Case "regression", "regressionselection", "regressionudf"
      Set CtrCore = New AppExcliqRegression_current
    Case "balmet", "balmetselection", "balmeteco", "balmetudf", "balmetudfeco"
      Set CtrCore = New AppExcliqBalance_current
    Case "projection"
      Set CtrCore = New AppExcliqProjections_current
    Case "oreinvoice"
      Set CtrCore = New AppExcliqPurchaseMin_current
    Case "concentrateinvoice"
      Set CtrCore = New AppExcliqPurchaseCnc_current
    Case Else: GoTo EH
  End Select
  Set CtrCore.ErrorHandler = EHGLOBAL
  
  'Return
  Set Ctrl_get_object = CtrCore

EH:
  Set CtrCore = Nothing
  Call EHGLOBAL.ErrorHandlerRaise("CONTROLLER::-get_object")

End Function

Private Function Ctrl_process_data(ByVal StrId As String, ByVal CVarRequest As Variant, ByRef CVarResponse As Variant) As Boolean 'TODO: TO EVENT HNDL
  
  On Error GoTo EH
  ' Get the specific controller
  Set ICC = Ctrl_get_object(StrId)

  ' Process the request
  Call VBA.MsgBox("Greetings from Controller!: " & StrId, vbExclamation, "EQL")
  Let Ctrl_process_data = ICC.GetSolution(StrId, CVarRequest, CVarResponse)
  
EH:
  Call EHGLOBAL.ErrorHandlerRaise("CONTROLLER::-process_data")

End Function

Private Function Ctrl_get_data(ByVal CStrId As String) As Variant
  
  Dim MDL As ModelExcliqliteDatasheet
  
  On Error GoTo EH
  'Calls ModelApp and request some data
  Call Debug_assert(Not EHGLOBAL Is Nothing, "CONTROLLER::-getdata")
  Set MDL = New ModelExcliqliteDatasheet
  Set MDL.ErrorHandler = EHGLOBAL
  Select Case CStrId
    Case "configtabletype": Let Ctrl_get_data = MDL.Tables
    Case "configtablestyle": Let Ctrl_get_data = MDL.TableStyles
    Case "configformulas": Let Ctrl_get_data = MDL.Formulas
    Case "configtablerowshadows": Let Ctrl_get_data = MDL.TableRowStripes
    Case "configcellforselections": Let Ctrl_get_data = MDL.TableSelectionsCell
  End Select
  
EH:
  Set MDL = Nothing
  Call EHGLOBAL.ErrorHandlerRaise("CONTROLLER::-getdata")

End Function

Private Sub Ctrl_set_data(ByVal CStrId As String, ByVal CVarData As Variant)
  
  Dim MDL As ModelExcliqliteDatasheet
  
  On Error GoTo EH
  'Calls ModelApp and sets some data
  Call Debug_assert(Not EHGLOBAL Is Nothing, "CONTROLLER::-getdata")
  Set MDL = New ModelExcliqliteDatasheet
  Set MDL.ErrorHandler = EHGLOBAL
  Select Case CStrId
    Case "configtabletype": Let MDL.Tables = VBA.CByte(CVarData)
    Case "configtablestyle": Let MDL.TableStyles = VBA.CByte(CVarData)
    Case "configformulas": Let MDL.Formulas = VBA.CBool(CVarData)
    Case "configtablerowshadows": Let MDL.TableRowStripes = VBA.CBool(CVarData)
    Case "configcellforselections": Let MDL.TableSelectionsCell = VBA.CBool(CVarData)
  End Select
  
EH:
  Set MDL = Nothing
  Call EHGLOBAL.ErrorHandlerRaise("CONTROLLER::-setdata")

End Sub

Private Function Ctrl_Excecutor_processor(ByVal CStrId As String, CEnuAction As EQLCTR_ENU_TABLE_AS, ByVal CVarRequest As Variant) As Variant
  
  On Error GoTo EH
  Select Case CEnuAction
    Case eqlCtrGet
      Let Ctrl_Excecutor_processor = Ctrl_get_data(CStrId)
    
    Case eqlCtrSet
      Call Ctrl_set_data(CStrId, CVarRequest)
    
    Case eqlCtrShw
      Set VIE = New ViewApp_current
      Set VIE.ErrorHandler = EHGLOBAL
      Call VIE.ViewUI(CStrId)
    
    Case Else: GoTo EH
  End Select
  
EH:
  Set VIE = Nothing
  Call EHGLOBAL.ErrorHandlerRaise("CONTROLLER::Excecutor_processor")

End Function






' ERRORS MANNAGEMENT AND ASSERTSIONS ------------------------------------------------'
Private Sub Debug_assert(ByVal Cond As Boolean, Optional Origin As String = VBA.vbNullString)
  #If Debugging Then
    Debug.Assert Cond
    If Not Cond Then Debug.Print "Debugged from: " & Origin
  #End If
End Sub
' METHODS PRIVATE -------------------------------------------------------------------'







' METHODS PUBLIC --------------------------------------------------------------------'
Public Function CtrlExecute(ByVal CStrId As String, CEnuAction As EQLCTR_ENU_TABLE_AS, Optional CVarRequest As Variant) As Variant
  
  On Error GoTo EH
  'Send Request to acquire a Response
  Let CtrlExecute = Ctrl_Excecutor_processor(CStrId, CEnuAction, CVarRequest)
  
EH:
  Call EHGLOBAL.ErrorHandlerRaise("CONTROLLER::Execute")

End Function

Public Function CtrlExecuteUDF(ByVal CStrId As String, ByVal CVarRequest As Variant) As Variant
  
  On Error GoTo EH
  Set ICC = Ctrl_get_object(CStrId)
  
  'Send Request to acquire a Response
  Let CtrlExecuteUDF = ICC.GetSolutionUDF(CStrId, CVarRequest)
  
EH:
  Debug.Print VBA.Err.Description, VBA.Err.Number
  Set ICC = Nothing

End Function
' METHODS PUBLIC --------------------------------------------------------------------'


