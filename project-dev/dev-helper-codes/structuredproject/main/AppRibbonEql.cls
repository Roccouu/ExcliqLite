VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RibbonExcliq_current"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' OBJECT VARIABLES (GLOOBJ_) --------------------------------------------------------'
'App
Private WithEvents APP As Application
Attribute APP.VB_VarHelpID = -1
'Errors
Private EHGLOBAL As AppErrorHandler
'System Ribbon
Private EQRIBBON As IRibbonUI
'Ribbon
Private GLOOBJRIBBON As RibbonExcliqItems_current

' VARIABLES (GLOStr_) ---------------------------------------------------------------'
'System name
Private GLOSTR_APPTITLE As String
'Ribbon status controls
Private RIBBONINITIALIZED As Boolean
Private RIBBONSTATUS As Boolean

' CONSTANTS ERRORS AND ERROR MESSAGES (GLOSTR_ERR_) ---------------------------------'
Private Const CUSTOM_ERROR As Long = VBA.vbObjectError + 514
Private Const EQPREFIX As String = "btnexcliqlite"



'============================= CLASS RIBBON STRUCTURE ==============================='
' OBJECT VARIABLES (GLOOBJ_) --------------------------------------------------------'
' VARIABLES (GLOStr_) ---------------------------------------------------------------'
' CONSTANTS ERRORS AND ERROR MESSAGES (GLOSTR_ERR_) ---------------------------------'
' CONSTANTS (GLOSTR_) ---------------------------------------------------------------'
' PROPERTIES ------------------------------------------------------------------------'
' CONSTRUCTOR/DESTRUCTOR ------------------------------------------------------------'
' CORE METHODS LIST PRIVATE ---------------------------------------------------------'
' Ribbon_initial_values
' Ribbon_SetStandBy
' Ribbon_Commons_Reset
' Ribbon_Enabling
' APP_NewOpenWorkbook
' APP_DisableEnableRibbon
' ASSETS ----------------------------------------------------------------------------'
' ERRORS MANNAGEMENT AND ASSERTSIONS ------------------------------------------------'
' Debug_assert
' INTERFACE METHODS LIST ------------------------------------------------------------'
' RibbonCommonsInit
' RibbonCommonsStatusSetter
' RibbonCommonsStatusGetter
' RibbonCommonsActionsExecutorControls
'============================= CLASS RIBBON STRUCTURE ==============================='



' PROPERTIES ------------------------------------------------------------------------'
Public Property Let AppTitle(ByVal PStrTitle As String)
  Let GLOSTR_APPTITLE = PStrTitle
End Property

Public Property Get APPRIBBONINITIALIZED() As Boolean
  Let APPRIBBONINITIALIZED = RIBBONINITIALIZED
End Property



' CONSTRUCTOR/DESTRUCTOR ------------------------------------------------------------'
Private Sub Class_Initialize()
  'Dim RES As AppResources

  If APP Is Nothing Then Set APP = Application
  Set EHGLOBAL = New AppErrorHandler
  Set GLOOBJRIBBON = New RibbonExcliqItems_current
  Set GLOOBJRIBBON.PROPObjEH = EHGLOBAL

  'Set RES = New AppResources
  'Let GLOSTR_APPTITLE = RES.pRESTITLE '"Excliq"
  'Set RES = Nothing
End Sub

Private Sub Class_Terminate()
  Set APP = Nothing
  Set GLOOBJRIBBON = Nothing
  Set EHGLOBAL = Nothing
End Sub








' CORE METHODS LIST PRIVATE ---------------------------------------------------------'
' Initialize ribbon values
Private Function Ribbon_initial_values() As Boolean
  On Error GoTo EH
  Let Ribbon_initial_values = GLOOBJRIBBON.RIBBONITEMINITIAL()

EH:
  Call EHGLOBAL.ErrorHandlerRaise("RIBBON:-initial_values")
End Function

'RIBBON StandBy
Private Sub Ribbon_SetStandBy()
  Dim VecRi As Variant
  Dim i As Long
  
  On Error GoTo EH
  Let VecRi = GLOOBJRIBBON.PROPVecItems
  
  For i = 0 To UBound(VecRi)
    #If Debugging Then
      Call Debug_assert(Not VecRi(i, 0) = VBA.vbNullString, "standby")
    #End If
    Call Ribbon_Commons_Reset(EQPREFIX & VBA.StrConv(VBA.CStr(VecRi(i, 0)), vbProperCase), VecRi(i, 1))
  Next i

EH:
  Let VecRi = Empty
  Call EHGLOBAL.ErrorHandlerRaise("RIBBON::-setstandby")
End Sub

' Commons
Private Sub Ribbon_Commons_Reset(ByVal ControlId As String, Optional Status As Variant)
  On Error GoTo EH
  #If Debugging Then
    Debug.Print "RESET EQRIBBON STATUS: ", ControlId, Status '(EQRIBBON Is Nothing)
  #End If
  
  If RIBBONSTATUS Then
    #If Debugging Then
      Debug.Print "Reseteo el control: ", ControlId & " A", Status
    #End If
    Call EQRIBBON.InvalidateControl(ControlId)
  End If

EH:
  Call EHGLOBAL.ErrorHandlerRaise("RIBBON::-commons_reset")
End Sub

'Ribbon enabling method
Private Sub Ribbon_Enabling(Optional BooEnabling As Boolean = False)
  Dim VecRi As Variant
  Dim i As Long
  
  On Error GoTo EH
  Let VecRi = GLOOBJRIBBON.PROPVecItems
  
  For i = 0 To UBound(VecRi)
    Call GLOOBJRIBBON.RIBBONITEMENABLINGALL(BooEnabling)
  Next i

EH:
  Let VecRi = Empty
  Call EHGLOBAL.ErrorHandlerRaise("RIBBON::-enabling")
End Sub

'APPLICATION MAIN EVENTS TO SETTING BASIC AND MOST IMPORTANT SYSTEM RESOURCES
'Methods Private
Private Sub APP_NewOpenWorkbook(ByVal Wb As Workbook, Optional NewOpen As Boolean = False)
  Const sFN As String = "RIBBON::-new_open_workbook"
  
  #If Debugging Then
    Call VBA.MsgBox(IIf(NewOpen, "Open", "New") & " Workbook is: " & Wb.Name)
  #End If

  On Error GoTo EH
  If Not EQRIBBON Is Nothing Then
    If Not RIBBONSTATUS And Application.Workbooks.Count > 0 Then
      Call Ribbon_Enabling(True)
      Call APP_DisableEnableRibbon(True)
    End If
  Else
    Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "No existe Objeto Ribbon para iniciar Sistema.")
  End If

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)
End Sub

Private Sub APP_DisableEnableRibbon(Optional DisableEnable As Boolean = False)
  Dim vRibbon As Variant
  Dim i As Long
  
  On Error GoTo EH
  Let vRibbon = Array("SysGrpExcliq", "SysBtnCloseSession", "SysBtnNewUser", "SysBtnNewTerminal", "GesMnuGestions", "GesMnuPersons", "GesMnuValues", "GesMnuProduction", "GesMnuEconomics", "GesMnuExport", "SimSplExcliqOne", "SimSplExcliqTwo", "SimBtnNewProjection", "SimBtnNewOreInvoice", "SimBtnNewConcentrateInvoice", "SimBtnSimulations", "ConMnuUsers", "ConMnuGestion", "ConMnuSystem", "HlpBtnHelp", "HlpSplExcliq")
  Let vRibbon = GLOOBJRIBBON.PROPVecItems
  For i = 0 To UBound(vRibbon)
    Call Ribbon_Commons_Reset(vRibbon(i, 0), vRibbon(i, 1))
  Next i
  Let RIBBONSTATUS = DisableEnable
  
EH:
  Let vRibbon = Empty
  Call EHGLOBAL.ErrorHandlerRaise("RIBBON::-disable_enable_ribbon")
End Sub


'Events
'Private Sub APP_NewWorkbook(ByVal Wb As Workbook)
'  'Check for Ribbon
'  Call APP_NewOpenWorkbook(Wb)
'  'Shows a subscribe message when current workbook is opened
'  'Call RES.resAppRegisterMsg(APPREGISTERED) 'TODO: MEJORAR
'End Sub
'
'Private Sub APP_WorkbookOpen(ByVal Wb As Workbook)
'  'Check for Ribbon
'  Call APP_NewOpenWorkbook(Wb, True)
'  'Shows a subscribe message when current workbook is opened
'  'Call RES.resAppRegisterMsg(APPREGISTERED) 'TODO: MEJORAR
'End Sub
'
'Private Sub APP_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
'  Dim x As Long
'  Dim y As Boolean
'
'  If Not EQRIBBON Is Nothing And Not Wb Is Nothing Then
'    If Wb.IsAddin Then
'      If Wb.Name = ThisWorkbook.Name Then
'        'Shows a subscribe message when current workbook is closed
'        'Call RES.resAppRegisterMsg(APPREGISTERED) 'TODO: MEJORAR
'        'Call cSYS.SystemAppQuit 'TODO: MEJORAR
'        'TODO: AGREGAR FUNCIONALIDAD QUE GENERE UN TOKEN Y LO CUARDE EN LA DB Y ESHEET
'        #If Debugging Then
'          Call VBA.MsgBox("Se pueden ejecutar tareas de cierre de addin")
'        #End If
'      End If
'    Else
'      Let x = Application.Workbooks.Count - 1
'      #If Debugging Then
'        MsgBox "Existen " & x & " libros por cerrar"
'      #End If
'      If x < 1 Then
'        If RIBBONSTATUS Then
'          Call Ribbon_Enabling
'          Call Wb.Close
'          Let y = (Application.Workbooks.Count <= 0)
'          Select Case y
'            Case True
'              Call APP_DisableEnableRibbon
'            Case False
'              Call Ribbon_Enabling(True)
'              Let RIBBONSTATUS = True
'              Let Cancel = True
'          End Select
'        End If
'      End If
'    End If
'  Else
'    'Error
'  End If
'End Sub

'Events vendors
'Private Sub App_WorkbookBeforePrint(ByVal Wb As Workbook, Cancel As Boolean)
'  '1. Revisa si el sistema no está registrado.
'  '2. Revisa si el libro actual tiene hojas generadas con unExcliq
'  '3. Antes de imprimir sella la hoja en el encabezado.
'  'MsgBox "Libro actual: " & ActiveWorkbook.Name & ", la hoja actual: " & ActiveSheet.Name & ", y el Wb es: " & wb.Name & ", y ActiveSheet en el Wb es: " & wb.ActiveSheet.Name
'  Call RES.resExcelPageSetupExtras(True, Wb, APPREGISTERED)
'  Cancel = True
'End Sub
'
'Private Sub App_WorkbookAfterSave(ByVal Wb As Workbook, ByVal Success As Boolean)
'  'Shows a subscribe message when current workbook is saved for fiveth time
'  APPSAVES = APPSAVES + 1
'  If APPSAVES > 5 Then APPSAVES = 0: Call RES.resAppRegisterMsg(APPREGISTERED)
'End Sub
'
'Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
'  'Deletes CenterHeader's unExcliq rangeName when user points any cell
'  'MsgBox "Libro actual: " & Sh.Parent.Name & ", la hoja actual: " & Sh.Name
'  Call RES.resExcelPageSetupExtras(False, Sh.Parent, APPREGISTERED)
'End Sub








' ERRORS MANNAGEMENT AND ASSERTSIONS ------------------------------------------------'
' Debug_assert
Private Sub Debug_assert(ByVal Cond As Boolean, Optional Origin As String = VBA.vbNullString)
  #If Debugging Then
    Debug.Assert Cond
    If Not Cond Then Debug.Print "Debugged from: " & Origin
  #End If
End Sub








' INTERFACE METHODS LIST ------------------------------------------------------------'
'Commons
Public Sub RibbonCommonsInit(ByVal Ribbon As IRibbonUI)
  
  Const sFN As String = "RIBBON::Init"
  
  On Error GoTo EH
  Let RIBBONSTATUS = (Not Ribbon Is Nothing)
  If RIBBONSTATUS Then
    Set EQRIBBON = Ribbon
    
    #If Debugging Then
      Call VBA.MsgBox("Aplicación" & IIf(RIBBONINITIALIZED, " iniciada.", " iniciando."), vbInformation, Application.Name)
    #End If

    'Initialize System included Ribbon just if System is started for first time. This is to work with only one instance.
    If Not RIBBONINITIALIZED Then
      
      #If Debugging Then
        Call VBA.MsgBox(IIf(RIBBONINITIALIZED, "Continúo la ", "Inicio la ") & "Aplicación", vbInformation, Application.Name)
      #End If

      Let RIBBONINITIALIZED = Ribbon_initial_values()

      #If Debugging Then
        Call VBA.MsgBox("Estableciendo Ribbon en modo de espera.", vbInformation, GLOSTR_APPTITLE)
      #End If

      Call Ribbon_SetStandBy 'Setting Ribbon values to Stand by.
    End If
  Else
    Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "No existe Objeto Ribbon para iniciar Sistema.")
  End If

  #If Debugging Then
    Debug.Print "Primero"
  #End If

EH:
  Call EHGLOBAL.ErrorHandlerDisplay(sFN)
End Sub

Public Function RibbonCommonsStatusSetter(ByVal ControlId As String, Optional GetVisibleEnable As Boolean = False) As Variant
  On Error GoTo EH
  #If Debugging Then
    Debug.Print "Segundo", ControlId
  #End If
  
  'Ribbon controls enabling status setter
  If RIBBONINITIALIZED Then
    If GetVisibleEnable Then 'Enable/Disable
      Let RibbonCommonsStatusSetter = GLOOBJRIBBON.RIBBONITEMGET(ControlId)
    Else 'Visibility
      'Let RibbonCommonsStatusSetter = uRIBBONCOL.Item(ControlId)
    End If
  End If
EH:
  Call EHGLOBAL.ErrorHandlerDisplay("RIBBON::CommonsStatusSetter")
End Function

'Public Function RibbonCommonsStatusGetter(ByVal ControlId As String) As Variant
'  Dim RES As AppResources
'
'  On Error GoTo EH
'  #If Debugging Then
'    Debug.Print "Segundo", ControlId
'  #End If
'
'  'Ribbon controls main enabling status setter
'  If RIBBONINITIALIZED Then
'    Set RES = New AppResources
'    Set RES.pPROPOBJ_EH = EHGLOBAL
'    Select Case ControlId
'      Case "configtabletype": Let RibbonCommonsStatusGetter = RES.resAPPEQGETTABLES()
'      Case "configtablestyle": Let RibbonCommonsStatusGetter = RES.resAPPEQGETTABLESTYLES()
'      Case "configformulas": Let RibbonCommonsStatusGetter = RES.resAPPEQGETFORMULAS()
'      Case "configtablerowshadows": Let RibbonCommonsStatusGetter = RES.resAPPEQGETTABLEROWSTRIPES()
'      Case "configcellforselections": Let RibbonCommonsStatusGetter = RES.resAPPEQGETTABLESELECTIONSMODS()
'    End Select
'    Set RES = Nothing
'  End If
'
'EH:
'  Set RES = Nothing
'  Call EHGLOBAL.ErrorHandlerDisplay("RIBBON::CommonsStatusGetter")
'End Function
'
''RIBBON Action listeners
'Public Sub RibbonCommonsActionsExecutorControls(ByVal ControlId As String, Optional id As String, Optional index As Integer)
'  Dim RES As AppResources
'  Dim VIE As View
'  'Dim vAll As ViewAlltasks
'  Dim Action As String
'  Const sFN As String = "RIBBON::ExecutorControls"
'  'So, if EH is not previously registered, then create a new EH object
'  'If RegisteredError Then Set eROR = New AppErrorHandler
'  On Error GoTo REH
'  If Not RIBBONINITIALIZED Then
'    Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "La aplicación debe ser cerrada debido a un error interno.")
'  End If
'
'  #If Debugging Then
'    Call VBA.MsgBox(ControlId)
'  #End If
'  Set VIE = New View
'  Set VIE.ErrorHandler = EHGLOBAL
'  Select Case ControlId
'    Case "regression", "regressionselection"
'      Call VIE.ViewUI(ControlId)
'    Case "balmet", "balmeteco", "balmetselection"
'      Call VIE.ViewUI(ControlId)
'    Case "projection"
'      Call VIE.ViewUI(ControlId)
'    Case "oreinvoice"
'      Call VIE.ViewUI(ControlId)
'    Case "concentrateinvoice"
'      Call VIE.ViewUI(ControlId)
'    Case "configtabletype", "configtablestyle"
'      Set VIE = Nothing
'      Set RES = New AppResources
'      If ControlId = "configtabletype" Then Call RES.resAPPEQSETTABLES(VBA.CSng(index))
'      If ControlId = "configtablestyle" Then Call RES.resAPPEQSETTABLESTYLES(VBA.CSng(index))
'      Set RES = Nothing
'    Case "configformulas", "configtablerowshadows", "configcellforselections"
'      Set VIE = Nothing
'      Set RES = New AppResources
'      If ControlId = "configformulas" Then Call RES.resAPPEQSETFORMULAS(VBA.CBool(id))
'      If ControlId = "configtablerowshadows" Then Call RES.resAPPEQSETTABLEROWSTRIPES(VBA.CBool(id))
'      If ControlId = "configcellforselections" Then Call RES.resAPPEQSETTABLESELECTIONSMODS(VBA.CBool(id))
'      Set RES = Nothing
'    Case "configlauncher"
'      Call VIE.ViewUI(ControlId)
'    Case "help"
'    Case "about"
'    Case "license"
'    Case Else
'  End Select
'  Set VIE = Nothing
'  Let RIBBONINITIALIZED = Not EQRIBBON Is Nothing
'
'REH:
'  Set VIE = Nothing
'  Set RES = Nothing
'  Call EHGLOBAL.ErrorHandlerDisplay(sFN)
'End Sub
'============================= CLASS RIBBON STRUCTURE ==============================='


