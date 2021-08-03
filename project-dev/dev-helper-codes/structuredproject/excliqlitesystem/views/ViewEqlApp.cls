VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ViewApp_current"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' CUSTOM EVENTS ---------------------------------------------------------------------'
Public Event RequestEvent(ByVal EStrRequester As String, ByVal EVarRequest As Variant, ByRef EVarResponse As Variant, ByRef EBooCancel As Boolean)

' OBJECT VARIABLES (GLOOBJ_) --------------------------------------------------------'
Private EHGLOBAL As AppErrorHandler
Private DPR As ClassDatePicker
Private CMB As ClassCustomMsgbox

Private WithEvents BtnAccept As MSForms.CommandButton
Attribute BtnAccept.VB_VarHelpID = -1
Private WithEvents BtnCancel As MSForms.CommandButton
Attribute BtnCancel.VB_VarHelpID = -1
Private WithEvents Btn0 As MSForms.CommandButton 'Buttons
Attribute Btn0.VB_VarHelpID = -1
Private WithEvents Btn1 As MSForms.CommandButton
Attribute Btn1.VB_VarHelpID = -1
Private WithEvents Btn2 As MSForms.CommandButton
Attribute Btn2.VB_VarHelpID = -1
Private WithEvents Btn3 As MSForms.CommandButton
Attribute Btn3.VB_VarHelpID = -1
Private WithEvents Btn4 As MSForms.CommandButton
Attribute Btn4.VB_VarHelpID = -1
Private WithEvents Btn5 As MSForms.CommandButton
Attribute Btn5.VB_VarHelpID = -1
Private WithEvents Btn6 As MSForms.CommandButton
Attribute Btn6.VB_VarHelpID = -1
Private WithEvents Btn7 As MSForms.CommandButton
Attribute Btn7.VB_VarHelpID = -1
Private WithEvents Btn8 As MSForms.CommandButton
Attribute Btn8.VB_VarHelpID = -1
Private WithEvents Btn9 As MSForms.CommandButton
Attribute Btn9.VB_VarHelpID = -1
Private WithEvents Btn10 As MSForms.CommandButton
Attribute Btn10.VB_VarHelpID = -1
Private WithEvents Btn11 As MSForms.CommandButton
Attribute Btn11.VB_VarHelpID = -1
Private WithEvents Btn12 As MSForms.CommandButton
Attribute Btn12.VB_VarHelpID = -1
Private WithEvents Btn13 As MSForms.CommandButton
Attribute Btn13.VB_VarHelpID = -1
Private WithEvents Btn14 As MSForms.CommandButton
Attribute Btn14.VB_VarHelpID = -1
Private WithEvents Btn15 As MSForms.CommandButton
Attribute Btn15.VB_VarHelpID = -1

Private WithEvents Chk0 As MSForms.CheckBox
Attribute Chk0.VB_VarHelpID = -1
Private WithEvents Chk1 As MSForms.CheckBox
Attribute Chk1.VB_VarHelpID = -1
Private WithEvents Chk2 As MSForms.CheckBox
Attribute Chk2.VB_VarHelpID = -1
Private WithEvents Chk3 As MSForms.CheckBox
Attribute Chk3.VB_VarHelpID = -1
Private WithEvents Chk4 As MSForms.CheckBox
Attribute Chk4.VB_VarHelpID = -1
Private WithEvents Chk5 As MSForms.CheckBox
Attribute Chk5.VB_VarHelpID = -1

Private WithEvents Cbo0 As MSForms.ComboBox
Attribute Cbo0.VB_VarHelpID = -1
Private WithEvents Cbo1 As MSForms.ComboBox
Attribute Cbo1.VB_VarHelpID = -1
Private WithEvents Cbo2 As MSForms.ComboBox
Attribute Cbo2.VB_VarHelpID = -1
Private WithEvents Cbo3 As MSForms.ComboBox
Attribute Cbo3.VB_VarHelpID = -1

Private WithEvents Lst0 As MSForms.ListBox
Attribute Lst0.VB_VarHelpID = -1
Private WithEvents Lst1 As MSForms.ListBox
Attribute Lst1.VB_VarHelpID = -1

Private WithEvents Txt0 As MSForms.TextBox
Attribute Txt0.VB_VarHelpID = -1
Private WithEvents Txt1 As MSForms.TextBox
Attribute Txt1.VB_VarHelpID = -1
Private WithEvents Txt2 As MSForms.TextBox
Attribute Txt2.VB_VarHelpID = -1
Private WithEvents Txt3 As MSForms.TextBox
Attribute Txt3.VB_VarHelpID = -1

Private WithEvents Spn0 As MSForms.SpinButton
Attribute Spn0.VB_VarHelpID = -1
Private WithEvents Spn1 As MSForms.SpinButton
Attribute Spn1.VB_VarHelpID = -1
Private WithEvents Spn2 As MSForms.SpinButton
Attribute Spn2.VB_VarHelpID = -1
Private WithEvents Spn3 As MSForms.SpinButton
Attribute Spn3.VB_VarHelpID = -1
Private WithEvents Spn4 As MSForms.SpinButton
Attribute Spn4.VB_VarHelpID = -1
Private WithEvents Spn5 As MSForms.SpinButton
Attribute Spn5.VB_VarHelpID = -1

' VARIABLES (GLOStr_) ---------------------------------------------------------------'
Private GloStrRngA As String
Private GloStrRngB As String
Private GloStrRngC As String
Private GloStrRngD As String
Private GloStrRngE As String
Private GLORNG_CELL As Range

'For data management
Private STRAPPTITLE As String
Private GLOBOO_FORMULAS As Boolean
Private GLOBOO_SELECCTIONCELL As Boolean
Private GLOBOO_COMPACT As Boolean
Private GLOBOO_HEADS As Boolean

' CONSTANTS ERROR MESSAGES (GLOSTR_ERR_) --------------------------------------------'
Private Const CUSTOM_ERROR As Long = VBA.vbObjectError + 517

' CONSTANTS (GLOSTR_) ---------------------------------------------------------------'
Private Const STR_TITLE8 As String = "Seleccione una celda."
Private Const BAL_LIMIT As Single = 5
Private Const GLOSNG_PUR_LIMIT As Single = 10
Private Const CDBL_DM_TOP As Double = 10000
Private Const CDBL_PC_TOP As Double = 100
Private Const CDBL_GT_TOP As Double = 1000000
Private Const CDBL_MAX As Double = 10000000000000#













' ============================== CLASS VIEWS STRUCTURE =============================='
' OBJECT VARIABLES (GLOOBJ_) --------------------------------------------------------'
' VARIABLES (GLOStr_) ---------------------------------------------------------------'
' CONSTANTS ERROR MESSAGES (GLOSTR_ERR_) --------------------------------------------'
' CONSTANTS (GLOSTR_) ---------------------------------------------------------------'
' PROPERTIES ------------------------------------------------------------------------'
' CONSTRUCTOR/DESTRUCTOR ------------------------------------------------------------'
' CONTROLS EVENTS -------------------------------------------------------------------'
' METHODS PRIVATE -------------------------------------------------------------------'
' CORE CONTROLS METHODS LIST (PRIVATE) ----------------------------------------------'
' View_Controls_control_event...:                                        View methods'
' CORE VIEWS METHODS LIST -----------------------------------------------------------'
' View_Core_activity...:                                                 View methods'
' CORE VIEWS FORMS METHODS LIST -----------------------------------------------------'
' View_form_module...:                                                   View methods'
' CORE VIEW ASSETS ------------------------------------------------------------------'
' View_assets_Module_...                                     All Assets in module'
' ERROR ASSERTSIONS -----------------------------------------------------------------'
' Debug_asset                                                    All Assets in module'
' METHODS PUBLIC --------------------------------------------------------------------'
' INTERFACE METHODS LIST (PUBLIC) ---------------------------------------------------'
' View_Module_...:                                                     All interfaces'
' ============================== CLASS VIEWS STRUCTURE =============================='



' PROPERTIES ------------------------------------------------------------------------'
Public Property Set ErrorHandler(ByRef ObjEH As AppErrorHandler)
  Set EHGLOBAL = ObjEH
End Property









' CONSTRUCTOR/DESTRUCTOR ------------------------------------------------------------'
Private Sub Class_Initialize()
  
  Dim MDL As ModelExcliqliteDatasheet
  On Error GoTo EH
  Set EHGLOBAL = New AppErrorHandler
  Set MDL = New ModelExcliqliteDatasheet
  Set MDL.ErrorHandler = EHGLOBAL
  Let STRAPPTITLE = MDL.AppTitle
  Let GLOBOO_FORMULAS = MDL.Formulas
  Let GLOBOO_SELECCTIONCELL = MDL.TableSelectionsCell
  Let GLOBOO_COMPACT = False 'MDL.Tables
  Let GLOBOO_HEADS = True 'MDL.resAPPEQGETTABLESELECTIONSMODSHEADS()

EH:
  Set MDL = Nothing
  Call EHGLOBAL.ErrorHandlerDisplay("VIEW::-initialize")

End Sub

Private Sub Class_Terminate()
  
  'Objects
  Set GLORNG_CELL = Nothing
  If Not DPR Is Nothing Then Set DPR = Nothing
  If Not CMB Is Nothing Then Set CMB = Nothing
  
  'Form Controls
  Call View_assets_controls_free
  Set EHGLOBAL = Nothing

End Sub






' CONTROLS EVENTS -------------------------------------------------------------------'
' Buttons
Private Sub BtnCancel_Click()
  Call View_assets_closeform(BtnCancel.Parent)
End Sub

Private Sub BtnAccept_Click()
  If View_Controls_buttons_click(BtnAccept, View_assets_controls_getformparent(BtnAccept)) Then Call View_assets_closeform(BtnCancel.Parent)
End Sub

Private Sub Btn0_Click()
  If View_Controls_buttons_click(Btn0, View_assets_controls_getformparent(Btn0)) Then Call View_assets_closeform(BtnCancel.Parent)
End Sub

Private Sub Btn1_Click()
  If View_Controls_buttons_click(Btn1, View_assets_controls_getformparent(Btn1)) Then Call View_assets_closeform(BtnCancel.Parent)
End Sub

Private Sub Btn2_Click()
  If View_Controls_buttons_click(Btn2, View_assets_controls_getformparent(Btn2)) Then Call View_assets_closeform(BtnCancel.Parent)
End Sub

Private Sub Btn3_Click()
  If View_Controls_buttons_click(Btn3, View_assets_controls_getformparent(Btn3)) Then Call View_assets_closeform(BtnCancel.Parent)
End Sub

Private Sub Btn4_Click()
  If View_Controls_buttons_click(Btn4, View_assets_controls_getformparent(Btn4)) Then Call View_assets_closeform(BtnCancel.Parent)
End Sub

Private Sub Btn5_Click()
  If View_Controls_buttons_click(Btn5, View_assets_controls_getformparent(Btn5)) Then Call View_assets_closeform(BtnCancel.Parent)
End Sub

Private Sub Btn6_Click()
  If View_Controls_buttons_click(Btn6, View_assets_controls_getformparent(Btn6)) Then Call View_assets_closeform(BtnCancel.Parent)
End Sub

Private Sub Btn7_Click()
  If View_Controls_buttons_click(Btn7, View_assets_controls_getformparent(Btn7)) Then Call View_assets_closeform(BtnCancel.Parent)
End Sub

Private Sub Btn8_Click()
  If View_Controls_buttons_click(Btn8, View_assets_controls_getformparent(Btn8)) Then Call View_assets_closeform(BtnCancel.Parent)
End Sub

Private Sub Btn9_Click()
  If View_Controls_buttons_click(Btn9, View_assets_controls_getformparent(Btn9)) Then Call View_assets_closeform(BtnCancel.Parent)
End Sub

Private Sub Btn10_Click()
  If View_Controls_buttons_click(Btn10, View_assets_controls_getformparent(Btn10)) Then Call View_assets_closeform(BtnCancel.Parent)
End Sub

Private Sub Btn11_Click()
  If View_Controls_buttons_click(Btn11, View_assets_controls_getformparent(Btn11)) Then Call View_assets_closeform(BtnCancel.Parent)
End Sub

Private Sub Btn12_Click()
  If View_Controls_buttons_click(Btn12, View_assets_controls_getformparent(Btn12)) Then Call View_assets_closeform(BtnCancel.Parent)
End Sub

Private Sub Btn13_Click()
  If View_Controls_buttons_click(Btn13, View_assets_controls_getformparent(Btn13)) Then Call View_assets_closeform(BtnCancel.Parent)
End Sub

Private Sub Btn14_Click()
  If View_Controls_buttons_click(Btn14, View_assets_controls_getformparent(Btn14)) Then Call View_assets_closeform(BtnCancel.Parent)
End Sub

Private Sub Btn15_Click()
  If View_Controls_buttons_click(Btn15, View_assets_controls_getformparent(Btn15)) Then Call View_assets_closeform(BtnCancel.Parent)
End Sub

' Checks
Private Sub Chk0_Change()
  Call View_Controls_checkboxes_change(Chk0, View_assets_controls_getformparent(Chk0))
End Sub

Private Sub Chk1_Change()
  Call View_Controls_checkboxes_change(Chk1, View_assets_controls_getformparent(Chk1))
End Sub

Private Sub Chk2_Change()
  Call View_Controls_checkboxes_change(Chk2, View_assets_controls_getformparent(Chk2))
End Sub

Private Sub Chk3_Change()
  Call View_Controls_checkboxes_change(Chk3, View_assets_controls_getformparent(Chk3))
End Sub

Private Sub Chk4_Change()
  Call View_Controls_checkboxes_change(Chk4, View_assets_controls_getformparent(Chk4))
End Sub

Private Sub Chk5_Change()
  Call View_Controls_checkboxes_change(Chk5, View_assets_controls_getformparent(Chk5))
End Sub

Private Sub Chk1_Click()
  Call View_Controls_checkboxes_click(Chk1, View_assets_controls_getformparent(Chk1))
End Sub

' Combos
Private Sub Cbo0_Change()
  Call View_Controls_combos_change(Cbo0, View_assets_controls_getformparent(Cbo0))
End Sub

Private Sub Cbo1_Change()
  Call View_Controls_combos_change(Cbo1, View_assets_controls_getformparent(Cbo1))
End Sub

Private Sub Cbo2_Change()
  Call View_Controls_combos_change(Cbo2, View_assets_controls_getformparent(Cbo2))
End Sub

Private Sub Cbo3_Change()
  Call View_Controls_combos_change(Cbo3, View_assets_controls_getformparent(Cbo3))
End Sub

' Lists
Private Sub Lst0_Change()
  Call View_Controls_listboxes_change(Lst0, View_assets_controls_getformparent(Lst0))
End Sub

Private Sub Lst0_Click()
  Call View_Controls_listboxes_click(Lst0, View_assets_controls_getformparent(Lst0))
End Sub

Private Sub Lst0_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  Call View_Controls_listboxes_dblclick(Lst0, View_assets_controls_getformparent(Lst0))
End Sub

Private Sub Lst1_Change()
  Call View_Controls_listboxes_change(Lst1, View_assets_controls_getformparent(Lst1))
End Sub

Private Sub Lst1_Click()
  Call View_Controls_listboxes_click(Lst1, View_assets_controls_getformparent(Lst1))
End Sub

Private Sub Lst1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  Call View_Controls_listboxes_dblclick(Lst1, View_assets_controls_getformparent(Lst1))
End Sub

' Spins
Private Sub Spn0_Change()
  Call View_Controls_spinbuttons_change(Spn0, View_assets_controls_getformparent(Spn0))
End Sub

Private Sub Spn1_Change()
  Call View_Controls_spinbuttons_change(Spn1, View_assets_controls_getformparent(Spn1))
End Sub

Private Sub Spn2_Change()
  Call View_Controls_spinbuttons_change(Spn2, View_assets_controls_getformparent(Spn2))
End Sub

Private Sub Spn3_Change()
  Call View_Controls_spinbuttons_change(Spn3, View_assets_controls_getformparent(Spn3))
End Sub

Private Sub Spn4_Change()
  Call View_Controls_spinbuttons_change(Spn4, View_assets_controls_getformparent(Spn4))
End Sub

Private Sub Spn5_Change()
  Call View_Controls_spinbuttons_change(Spn5, View_assets_controls_getformparent(Spn5))
End Sub

' Texts
Private Sub Txt0_DropButtonClick()
  Call View_Controls_texts_dropbuttonclick(Txt0, View_assets_controls_getformparent(Txt0))
End Sub

Private Sub Txt1_DropButtonClick()
  Call View_Controls_texts_dropbuttonclick(Txt1, View_assets_controls_getformparent(Txt1))
End Sub

Private Sub Txt2_DropButtonClick()
  Call View_Controls_texts_dropbuttonclick(Txt2, View_assets_controls_getformparent(Txt2))
End Sub

Private Sub Txt3_DropButtonClick()
  Call View_Controls_texts_dropbuttonclick(Txt3, View_assets_controls_getformparent(Txt3))
End Sub


Private Sub Txt0_Change()
  Call View_Controls_texts_change(Txt0, View_assets_controls_getformparent(Txt0))
End Sub

Private Sub Txt1_Change()
  Call View_Controls_texts_change(Txt1, View_assets_controls_getformparent(Txt1))
End Sub









' METHODS PRIVATE -------------------------------------------------------------------'
' CORE CONTROLS METHODS LIST (PRIVATE) ----------------------------------------------'
'Events handlers - Buttons
Private Function View_Controls_buttons_click( _
  ByVal VCtrBtn As MSForms.CommandButton, _
  ByVal FRM As FRMBasics, _
  Optional VBooSelections As Boolean = False, _
  Optional ControlId As String) As Boolean
  
  Dim RES As AppResources_current
  Dim REGEX As AppResRegEx
  Dim VWP As ViewsParent_current
  
  Dim RngBox As Range
  Dim VecAux As Variant, VecAux1 As Variant
  Dim sAux0 As String, sAux1 As String, sAux2 As String
  Dim BooAux As Boolean
  Dim X As Long
  Dim i As Single
  
  Const sFN As String = "VIEW::-buttons_click"
  
  On Error GoTo EH
  
  'Validate first
  If Not View_Core_validations(VCtrBtn, FRM, VBooSelections, ControlId) Then GoTo EH
  
  'Execute task
  If Not VBooSelections Then 'Form tasks
    
    #If Debugging Then
      Call Debug_assert((Not VCtrBtn Is Nothing Or FRM Is Nothing), sFN)
    #End If
    If VCtrBtn Is Nothing Or FRM Is Nothing Then GoTo EH
    
    Select Case FRM.Tag
      Case "regression"
        If VCtrBtn.Name = "Btnmain1" Then Let View_Controls_buttons_click = View_Core_execution(FRM.Tag, FRM)
      
'      Case "balmet"
'        Select Case VCtrBtn.Name
'          Case "Btnbal0", "Btnbal1", "Btnbal4", "Btnbal6" 'capture ranges datas
'            With FRM
'              Set VWP = New ViewsParent_current
'              Set VWP.ErrorHandler = EHGLOBAL
'              If VCtrBtn.Name = "Btnbal0" Then Set RngBox = VWP.ViewsCaptureRange(FRM, STR_TITLE8, STRAPPTITLE)
'              If VCtrBtn.Name = "Btnbal1" Then Set RngBox = VWP.ViewsCaptureRange(FRM, STR_TITLE8, STRAPPTITLE)
'              If VCtrBtn.Name = "Btnbal4" Then Set RngBox = VWP.ViewsCaptureRange(FRM, STR_TITLE8, STRAPPTITLE)
'              If VCtrBtn.Name = "Btnbal6" Then Set RngBox = VWP.ViewsCaptureRange(FRM, STR_TITLE8, STRAPPTITLE)
'              Set VWP = Nothing
'              If Not RngBox Is Nothing Then
'                Set RES = New AppResources
'                If VCtrBtn.Name = "Btnbal0" Then Let GloStrRngA = RES.resArrayGetDataColonSeparatedStringFromVariants(RngBox, X, False, True, RBooRngRowsCols:=True)
'                If VCtrBtn.Name = "Btnbal1" Then Let GloStrRngB = RES.resArrayGetDataColonSeparatedStringFromVariants(RngBox, X, False, True, RBooRngRowsCols:=True)
'                If VCtrBtn.Name = "Btnbal4" Then Let GloStrRngC = RES.resArrayGetDataColonSeparatedStringFromVariants(RngBox, X, False, True, RBooRngRowsCols:=True)
'                If VCtrBtn.Name = "Btnbal6" Then Set GLORNG_CELL = RngBox 'RES.resArrayGetDataAsStringFromRanges(RngBox, x, False, True)
'                Set RES = Nothing
'                Let sAux1 = IIf(VCtrBtn.Name = "Btnbal0", 0, IIf(VCtrBtn.Name = "Btnbal1", 1, IIf(VCtrBtn.Name = "Btnbal4", 8, 9)))
'                Let .Controls("Txtbal" & sAux1).Tag = "rng"
'                Let .Controls("Txtbal" & sAux1).Value = RngBox.address
'                Let .Controls("Txtbal" & sAux1).Enabled = False
'                Set RngBox = Nothing
'              End If
'            End With
'
'          Case "Btnbal2", "Btnbal3", "Btnbal5"
'            With FRM
'              If VCtrBtn.Name = "Btnbal2" Then Let GloStrRngA = VBA.vbNullString
'              If VCtrBtn.Name = "Btnbal3" Then Let GloStrRngB = VBA.vbNullString
'              If VCtrBtn.Name = "Btnbal5" Then Let GloStrRngC = VBA.vbNullString
'              Let sAux1 = IIf(VCtrBtn.Name = "Btnbal2", 0, IIf(VCtrBtn.Name = "Btnbal3", 1, 8))
'              Let .Controls("Txtbal" & sAux1).Tag = VBA.vbNullString
'              Let .Controls("Txtbal" & sAux1).Value = VBA.vbNullString
'              Let .Controls("Txtbal" & sAux1).Enabled = True
'              Call .Controls("Txtbal" & sAux1).SetFocus
'            End With
'
'          Case "Btnmain1" 'execute balmet task
'            Let View_Controls_buttons_click = View_Core_execution(FRM.Tag, FRM)
'
'          Case Else: GoTo EH
'        End Select
'
'      Case "balmeteco"
'        With FRM
'          Select Case VCtrBtn.Name
'            Case "Btnbal8"
'              If Not .Controls("Chkbal2").Value Then GoTo EH
'              Let sAux0 = .Controls("Txtbal6").Value
'              Let .Controls("Txtbal6").Value = sAux0 & IIf(sAux0 = VBA.vbNullString, VBA.vbNullString, ";") & .Controls("Cbobal2").Value
'
'            Case "Btnbal9"
'              If Not .Controls("Chkbal2").Value Then GoTo EH
'              Let .Controls("Txtbal6").Value = VBA.vbNullString
'
'            Case "Btnbal0", "Btnbal1", "Btnbal2", "Btnbal3", "Btnbal10", "Btnbal12"
'              With FRM
'                Set VWP = New ViewsParent_current
'                Set VWP.ErrorHandler = EHGLOBAL
'                If VCtrBtn.Name = "Btnbal0" Then Set RngBox = VWP.ViewsCaptureRange(FRM, STR_TITLE8 & " (O varias. Deben contener números):", STRAPPTITLE)
'                If VCtrBtn.Name = "Btnbal1" Then Set RngBox = VWP.ViewsCaptureRange(FRM, STR_TITLE8 & "(O varias. Deben contener números):", STRAPPTITLE)
'                If VCtrBtn.Name = "Btnbal2" Then Set RngBox = VWP.ViewsCaptureRange(FRM, STR_TITLE8 & "(O varias. Deben contener números):", STRAPPTITLE)
'                If VCtrBtn.Name = "Btnbal3" Then Set RngBox = VWP.ViewsCaptureRange(FRM, STR_TITLE8 & "(O varias. Deben contener números):", STRAPPTITLE)
'                If VCtrBtn.Name = "Btnbal10" And .Controls("Chkbal4").Value Then Set RngBox = VWP.ViewsCaptureRange(FRM, STR_TITLE8 & "(O varias. Deben contener números):", STRAPPTITLE)
'                If VCtrBtn.Name = "Btnbal12" And .Controls("Chkbal4").Value Then Set RngBox = VWP.ViewsCaptureRange(FRM, STR_TITLE8 & "(O varias. Deben contener números):", STRAPPTITLE)
'                Set VWP = Nothing
'                If Not RngBox Is Nothing Then
'                  Set RES = New AppResources
'                  If Not VCtrBtn.Name = "Btnbal10" And Not VCtrBtn.Name = "Btnbal12" Then
'                    Let sAux0 = VBA.Right(VCtrBtn.Name, 1)
'                    Let .Controls("Txtbal" & sAux0).Value = RES.resArrayGetDataColonSeparatedStringFromVariants(RngBox, X, False, True, RBooRngRowsCols:=True)
'                  Else
'                    If VCtrBtn.Name = "Btnbal10" Then Let .Controls("Txtbal8").Value = RES.resArrayGetDataColonSeparatedStringFromVariants(RngBox, X, False, True, RBooRngRowsCols:=True)
'                    If VCtrBtn.Name = "Btnbal12" Then Let .Controls("Txtbal9").Value = RES.resArrayGetDataColonSeparatedStringFromVariants(RngBox, X, False, True, RBooRngRowsCols:=True)
'                  End If
'                  Set RES = Nothing
'                  Set RngBox = Nothing
'                End If
'              End With
'
'            Case "Btnbal4", "Btnbal5", "Btnbal6", "Btnbal7", "Btnbal11", "Btnbal13"
'              With FRM
'                If Not VCtrBtn.Name = "Btnbal11" And Not VCtrBtn.Name = "Btnbal13" Then
'                  Let sAux0 = VBA.CSng(VBA.Right(VCtrBtn.Name, 1)) - 4
'                  Let .Controls("Txtbal" & sAux0).Value = VBA.vbNullString
'                Else
'                  If Not .Controls("Chkbal4").Value Then GoTo EH
'                  If VCtrBtn.Name = "Btnbal11" And .Controls("Chkbal4").Value Then Let .Controls("Txtbal8").Value = VBA.vbNullString: Let sAux0 = 8
'                  If VCtrBtn.Name = "Btnbal13" And .Controls("Chkbal4").Value Then Let .Controls("Txtbal9").Value = VBA.vbNullString: Let sAux0 = 9
'                End If
'                Call .Controls("Txtbal" & sAux0).SetFocus
'              End With
'
'            Case "Btnbal14" 'capture ranges cell
'              With FRM
'                If Not .Controls("Optbal1").Value Then GoTo EH
'                Set VWP = New ViewsParent_current
'                Set VWP.ErrorHandler = EHGLOBAL
'                If VCtrBtn.Name = "Btnbal14" Then Set RngBox = VWP.ViewsCaptureRange(FRM, STR_TITLE8 & " Será la referencia donde imprimir el resultado:", STRAPPTITLE)
'                Set VWP = Nothing
'                If Not RngBox Is Nothing Then
'                  Set RES = New AppResources
'                  Set GLORNG_CELL = RngBox
'                  Set RES = Nothing
'                  Let .Controls("Txtbal10").Value = RngBox.address
'                  Let .Controls("Txtbal10").Enabled = False
'                  Set RngBox = Nothing
'                End If
'              End With
'
'            Case "Btnmain1"
'              Let View_Controls_buttons_click = View_Core_execution(FRM.Tag, FRM)
'
'            Case Else: GoTo EH
'          End Select
'        End With
'
'      Case "projection"
'        With FRM
'          Select Case VCtrBtn.Name
'            Case "Btnpro0", "Btnpro1", "Btnpro3"
'              Call View_assets_capture_numbers_from_ranges(FRM, VCtrBtn.Name, "Txtpro", STR_TITLE8 & " (O varias. Deben contener números):", VBooSum:=(VCtrBtn.Name = "Btnpro0"))
'
'            Case "Btnpro2" '1;2;3/3;2;1
'              Call View_assets_capture_numbers_from_ranges(FRM, VCtrBtn.Name, "Txtpro", STR_TITLE8 & " (O varias. Deben contener números):", VBooSlashGrouped:=True)
'
'            Case "Btnpro4", "Btnpro6" 'Clear txts
'              With FRM
'                Let sAux1 = IIf(VCtrBtn.Name = "Btnpro4", "Txtpro4", "Txtpro5")
'                Let sAux0 = VBA.CStr(.Controls(sAux1).Value)
'                Let sAux0 = VBA.StrReverse(sAux0)
'                Let .Controls(sAux1).Value = VBA.StrReverse(VBA.Mid(sAux0, VBA.InStr(1, sAux0, ";", vbTextCompare) + 1))
'              End With
'
'            Case "Btnpro5", "Btnpro7" 'Clear txts
'              Let sAux0 = IIf(VCtrBtn.Name = "Btnpro5", "Txtpro4", "Txtpro5")
'              Let .Controls(sAux0).Value = VBA.vbNullString
'
'            Case "Btnpro8" 'capture ranges cell
'              If Not .Controls("Optpro3").Value Then GoTo EH
'              Call View_assets_capture_range_cell(FRM, VCtrBtn.Name, "Txtpro9")
'
'            Case "Btnmain1"
'              Let View_Controls_buttons_click = View_Core_execution(FRM.Tag, FRM)
'
'            Case Else: GoTo EH
'          End Select
'        End With
'
'      Case "oreinvoice"
'        With FRM
'          Select Case VCtrBtn.Name
'            Case "Btnpur0", "Btnpur9" 'Dates
'              Let sAux0 = "Txtpur" & IIf(VCtrBtn.Name = "Btnpur0", 1, 6)
'              Call DPR.DatePicker(FRM, DControl:=.Controls(sAux0))
'
'            Case "Btnpur1", "Btnpur2" 'Names & Type
'              Let sAux0 = "Txtpur" & IIf(VCtrBtn.Name = "Btnpur1", 3, 4)
'              Let sAux1 = "Seleccione un rango con datos " & IIf(VCtrBtn.Name = "Btnpur1", "Nombres de Provedores:", "Tipos Provedores:")
'              Call View_assets_capture_strings_from_ranges_customtextbox(FRM, VCtrBtn.Name, sAux0, sAux1)
'              Let VCtrBtn.Tag = "validate"
'              Let BooAux = View_Core_validations(VCtrBtn, FRM, VBooSelections, ControlId)
'              Let VCtrBtn.Tag = "selection"
'              If BooAux = False Then
'                Let .Controls(sAux0).Value = VBA.vbNullString
'                GoTo EH
'              End If
'
'            Case "Btnpur3", "Btnpur4" 'Add & Edit
'              Set RES = New AppResources
'              If VCtrBtn.Name = "Btnpur3" Then
'                Let .Controls("Lstpur0").RowSource = VBA.vbNullString
'                Select Case .Controls("Cbopur1").ListIndex
'                  Case 0 'Base Data
'                    Let sAux0 = "excliqlitepurchasebasedata"
'                    ReDim VecAux(1 To 1, 0 To 4) 'BASE 1 ROWS BECAUSE VBA IS A #$"!... VARIANTS FROM RANGES BEGINS AT 1,N
'                    For i = 0 To UBound(VecAux, 2)
'                      Let VecAux(1, i) = .Controls("Txtpur" & i + 5).Value
'                    Next i
'                    Call RES.resExcelxlTABLEADDROWFROMARRAY(excliqlitedatasheet.ListObjects(sAux0), VecAux)
'                    Let VecAux = Empty
'
'                  Case 1 'Main Data
'                    Let sAux0 = "excliqlitepurchasemaindata"
'                    If .Controls("Cbopur2").ListIndex = 1 Then 'Complex
'                      ReDim VecAux(1 To 1, 0 To 4)
'                      Let VecAux(1, 0) = .Controls("Cbopur2").Value
'                      Let VecAux(1, 1) = .Controls("Txtpur6").Value
'                      Let VecAux(1, 2) = .Controls("Txtpur7").Value
'                      Let VecAux(1, 3) = .Controls("Txtpur8").Value
'                      Let VecAux(1, 4) = .Controls("Txtpur9").Value
'                      Call RES.resExcelxlTABLEADDROWFROMARRAY(excliqlitedatasheet.ListObjects(sAux0), VecAux)
'                      Let VecAux = Empty
'                    Else
'                      ReDim VecAux(1 To 1, 0 To 4)
'                      Let VecAux(1, 0) = .Controls("Cbopur2").Value
'                      Let VecAux(1, 1) = .Controls("Cbopur3").Value
'                      Let VecAux(1, 2) = .Controls("Txtpur7").Value
'                      Let VecAux(1, 3) = .Controls("Cbopur4").Value
'                      Let VecAux(1, 4) = .Controls("Txtpur9").Value
'                      Call RES.resExcelxlTABLEADDROWFROMARRAY(excliqlitedatasheet.ListObjects(sAux0), VecAux)
'                      Let VecAux = Empty
'                    End If
'
'                  Case 2 'Bonus
'                    Let sAux0 = "excliqlitepurchasebonus"
'                    ReDim VecAux(1 To 1, 0 To 1)
'                    Let VecAux(1, 0) = .Controls("Txtpur5").Value
'                    Let VecAux(1, 1) = .Controls("Txtpur6").Value
'                    Call RES.resExcelxlTABLEADDROWFROMARRAY(excliqlitedatasheet.ListObjects(sAux0), VecAux)
'                    Let VecAux = Empty
'
'                  Case 6 'Discounts
'                    Let sAux0 = "excliqlitepurchasediscounts"
'                    ReDim VecAux(1 To 1, 0 To 1)
'                    Let VecAux(1, 0) = .Controls("Txtpur5").Value
'                    Let VecAux(1, 1) = .Controls("Txtpur6").Value
'                    Call RES.resExcelxlTABLEADDROWFROMARRAY(excliqlitedatasheet.ListObjects(sAux0), VecAux)
'                    Let VecAux = Empty
'
'                  Case Else: GoTo EH
'                End Select
'                Let .Controls("Btnpur5").Enabled = True
'                Let .Controls("Btnpur6").Enabled = True
'
'              Else 'Edit
'                Select Case .Controls("Cbopur1").ListIndex
'                  Case 0: Let sAux0 = "excliqlitepurchasebasedata" 'Base Data
'                  Case 1: Let sAux0 = "excliqlitepurchasemaindata" 'Main Data
'                  Case 2: Let sAux0 = "excliqlitepurchasebonus" 'Bonus
'                  Case 6: Let sAux0 = "excliqlitepurchasediscounts" 'Discounts
'                  Case Else: GoTo EH
'                End Select
'                Let X = .Controls("Lstpur0").ListIndex + 1
'                Let .Controls("Lstpur0").RowSource = VBA.vbNullString
'                Select Case .Controls("Cbopur1").ListIndex
'                  Case 0 'Base Data
'                    ReDim VecAux(0 To 0, 0 To 4)
'                    For i = 0 To UBound(VecAux, 2)
'                      Let VecAux(0, i) = .Controls("Txtpur" & i + 5).Value
'                    Next i
'
'                  Case 1 'Main Data
'                    If .Controls("Cbopur2").ListIndex = 1 Then 'Complex
'                      ReDim VecAux(0 To 0, 0 To 4)
'                      Let VecAux(0, 0) = .Controls("Cbopur2").Value
'                      Let VecAux(0, 1) = .Controls("Txtpur6").Value
'                      Let VecAux(0, 2) = .Controls("Txtpur7").Value
'                      Let VecAux(0, 3) = .Controls("Txtpur8").Value
'                      Let VecAux(0, 4) = .Controls("Txtpur9").Value
'                    Else
'                      ReDim VecAux(0 To 0, 0 To 4)
'                      Let VecAux(0, 0) = .Controls("Cbopur2").Value
'                      Let VecAux(0, 1) = .Controls("Cbopur3").Value
'                      Let VecAux(0, 2) = .Controls("Txtpur7").Value
'                      Let VecAux(0, 3) = .Controls("Cbopur4").Value
'                      Let VecAux(0, 4) = .Controls("Txtpur9").Value
'                    End If
'
'                  Case 2 'Bonus
'                    ReDim VecAux(0 To 0, 0 To 1)
'                    Let VecAux(0, 0) = .Controls("Txtpur5").Value
'                    Let VecAux(0, 1) = .Controls("Txtpur6").Value
'
'                  Case 5 'Discounts
'                    ReDim VecAux(0 To 0, 0 To 1)
'                    Let VecAux(0, 0) = .Controls("Txtpur5").Value
'                    Let VecAux(0, 1) = .Controls("Txtpur6").Value
'
'                  Case Else: GoTo EH
'                End Select
'                Let excliqlitedatasheet.ListObjects(sAux0).ListRows(X).Range.Value = VecAux
'                Let VecAux = Empty
'                For i = 3 To 8
'                  Let .Controls("Btnpur" & i).Enabled = IIf(i > 3 And i < 6, False, True)
'                Next i
'                Let .Controls("Cbopur1").Enabled = True
'                Let .Controls("Lstpur0").Enabled = True
'                Let .Controls("Btnmain1").Enabled = True
'              End If
'              Set RES = Nothing
'              Let .Controls("Lstpur0").RowSource = excliqlitedatasheet.ListObjects(sAux0)
'
'            Case "Btnpur5" 'Cancel Edition
'              For i = 3 To 8
'                Let .Controls("Btnpur" & i).Enabled = IIf(i > 3 And i < 6, False, True)
'              Next i
'              Let .Controls("Cbopur1").Enabled = True
'              Let .Controls("Lstpur0").Enabled = True
'              Let .Controls("Btnmain1").Enabled = True
'
'            Case "Btnpur6", "Btnpur7" 'Del once & Del all
'              Select Case .Controls("Cbopur1").ListIndex
'                Case 0: Let sAux0 = "excliqlitepurchasebasedata" 'Base Data
'                Case 1: Let sAux0 = "excliqlitepurchasemaindata" 'Main Data
'                Case 2: Let sAux0 = "excliqlitepurchasebonus" 'Bonus
'                Case 5: Let sAux0 = "excliqlitepurchasediscounts" 'Discounts
'                Case Else: GoTo EH
'              End Select
'              If VCtrBtn.Name = "Btnpur6" Then
'                Let i = .Controls("Lstpur0").ListIndex + 1
'                If Not (i > 0 And .Controls("Lstpur0").ListCount > 0) Then GoTo EH
'                Let .Controls("Lstpur0").RowSource = VBA.vbNullString
'                Call excliqlitedatasheet.ListObjects(sAux0).ListRows(i).Delete
'                Let .Controls("Lstpur0").RowSource = excliqlitedatasheet.ListObjects(sAux0)
'              Else
'                If Not .Controls("Lstpur0").ListCount > 0 Then GoTo EH
'                Let .Controls("Lstpur0").RowSource = VBA.vbNullString
'                Call excliqlitedatasheet.ListObjects(sAux0).DataBodyRange.Delete
'                Let .Controls("Lstpur0").RowSource = excliqlitedatasheet.ListObjects(sAux0)
'              End If
'              Let .Controls("Btnpur6").Enabled = (excliqlitedatasheet.ListObjects(sAux0).ListRows.Count > 0)
'              Let .Controls("Btnpur7").Enabled = (excliqlitedatasheet.ListObjects(sAux0).ListRows.Count > 0)
'
'            Case "Btnpur10" 'Rangecell
'              If Not .Controls("Optpur1").Value Then GoTo EH
'              Call View_assets_capture_range_cell(FRM, VCtrBtn.Name, "Txtpur10")
'
'            Case "Btnpur8" 'Rangedata
'              Call View_assets_capture_range_cell(FRM, VCtrBtn.Name, VBA.vbNullString, VStrTitle:="Seleccionar datos de " & .Controls("Cbopur1").Value, VBooJustCaptureRange:=True)
'              Let VCtrBtn.Tag = "validate"
'              Let BooAux = View_Core_validations(VCtrBtn, FRM, VBooSelections, ControlId)
'              Let VCtrBtn.Tag = "selection"
'              If BooAux = False Then
'                Set GLORNG_CELL = Nothing
'                GoTo EH
'              Else 'Correct information
'                Select Case .Controls("Cbopur1").ListIndex
'                  Case 0: Let sAux0 = "excliqlitepurchasebasedata" 'Base Data
'                  Case 1: Let sAux0 = "excliqlitepurchasemaindata" 'Main Data
'                  Case 2: Let sAux0 = "excliqlitepurchasebonus" 'Bonus
'                  Case 6: Let sAux0 = "excliqlitepurchasediscounts" 'Discounts
'                  Case Else: GoTo EH
'                End Select
'                Let VecAux = GLORNG_CELL.Value 'BASE 1 ROWS BECAUSE VBA IS A #$"!... VARIANTS FROM RANGES BEGINS AT 1,N
'                Set RES = New AppResources
'                Let .Controls("Lstpur0").RowSource = VBA.vbNullString
'                Debug.Print "Filas: ", excliqlitedatasheet.ListObjects(sAux0).ListRows.Count
'                Call RES.resExcelxlTABLEADDROWFROMARRAY(excliqlitedatasheet.ListObjects(sAux0), VecAux)
'                Set RES = Nothing
'                Let .Controls("Lstpur0").RowSource = excliqlitedatasheet.ListObjects(sAux0)
'                Let .Controls("Btnpur6").Enabled = True
'                Let .Controls("Btnpur7").Enabled = True
'              End If
'
'            Case "Btnmain1" 'Exe
'              Let View_Controls_buttons_click = View_Core_execution(FRM.Tag, FRM)
'
'            Case Else: GoTo EH
'          End Select
'        End With
'
'      Case "concentrateinvoice"
'        With FRM
'          Select Case VCtrBtn.Name
'            Case "Btncnc0", "Btncnc1" 'Dates
'              Let sAux0 = "Txtcnc" & IIf(VCtrBtn.Name = "Btncnc0", 1, 5)
'              Call DPR.DatePicker(FRM, DControl:=.Controls(sAux0))
'
'            Case "Btncnc2", "Btncnc3" 'Add & Edit UNIDAD MINERAL
'              Set RES = New AppResources
'              Let .Controls("Lstcnc0").Tag = "noevent"
'              'Let .Controls("Lstcnc0").RowSource = VBA.vbNullString
'              Select Case .Controls("Cbocnc1").ListIndex
'                Case 0 'Contents
'                  Let i = .Controls("Cbocnc2").ListIndex + 1
'                  Let sAux0 = .Controls("Txtcnc12").Value & "-" & excliqlitedatasheet.ListObjects("excliqliteconfigpurrm").ListColumns("UNIDAD MINERAL").DataBodyRange.Cells(i, 1).Value
'                  ReDim VecAux(1 To 1, 1 To 5)
'                  Let VecAux(1, 1) = .Controls("Cbocnc2").Value
'                  Let VecAux(1, 2) = sAux0
'                  Let VecAux(1, 3) = .Controls("Txtcnc13").Value
'                  Let VecAux(1, 4) = .Controls("Cbocnc3").Value
'                  Let VecAux(1, 5) = .Controls("Txtcnc14").Value
'                  Let sAux0 = "excliqliteconpurcontents"
'
'                Case 1 'Refinement
'                  ReDim VecAux(1 To 1, 1 To 3)
'                  Let VecAux(1, 1) = .Controls("Cbocnc2").Value
'                  Let VecAux(1, 2) = .Controls("Txtcnc12").Value
'                  Let VecAux(1, 3) = .Controls("Txtcnc13").Value
'                  'Let VecAux = Array(.Controls("Cbocnc2").Value, .Controls("Txtcnc12").Value, .Controls("Txtcnc13").Value)
'                  Let sAux0 = "excliqliteconpurrefination"
'
'                Case 2: GoTo EH 'Penalties
'
'                Case 3 'Deductions
'                  Select Case .Controls("Cbocnc2").ListIndex
'                    Case 0 'Exportations
'                      ReDim VecAux(1 To 1, 1 To 5)
'                      Let VecAux(1, 1) = .Controls("Txtcnc12").Value
'                      Let VecAux(1, 2) = .Controls("Txtcnc13").Value
'                      Let VecAux(1, 3) = .Controls("Txtcnc14").Value
'                      Let VecAux(1, 4) = IIf(.Controls("Chkcnc1").Value, 1, 0)
'                      Let VecAux(1, 5) = 1
'                      'Let VecAux = Array(.Controls("Txtcnc12").Value, .Controls("Txtcnc13").Value, .Controls("Txtcnc14").Value)
'                      Let sAux0 = "excliqliteconpurdedexp"
'
'                    Case 1: GoTo EH 'Retentions
'                    Case 2: GoTo EH 'Retentions others
'
''                    Case 3 'Rollback
''                      Let VecAux = Array(.Controls("Txtcnc12").Value, .Controls("Txtcnc13").Value, .Controls("Txtcnc14").Value)
''                      Let sAux0 = "excliqliteconpurdedrollback"
'
'                    Case 3 'Discounts others
'                      ReDim VecAux(1 To 1, 1 To 2)
'                      Let VecAux(1, 1) = .Controls("Txtcnc12").Value
'                      Let VecAux(1, 2) = .Controls("Txtcnc13").Value
'                      'Let VecAux = Array(.Controls("Txtcnc12").Value, .Controls("Txtcnc13").Value, .Controls("Txtcnc14").Value)
'                      Let sAux0 = "excliqliteconpurdeddiscounts"
'
'                    Case Else: GoTo EH
'                  End Select
'
'                Case 4 'Bonus
'                  ReDim VecAux(1 To 1, 1 To 2)
'                  Let VecAux(1, 1) = .Controls("Txtcnc12").Value
'                  Let VecAux(1, 2) = .Controls("Txtcnc13").Value
'                  'Let VecAux = Array(.Controls("Txtcnc12").Value, .Controls("Txtcnc13").Value, .Controls("Txtcnc14").Value)
'                  Let sAux0 = "excliqliteconpurbonus"
'
'                Case Else: GoTo EH
'              End Select
'              If VCtrBtn.Name = "Btncnc2" Then
'                Let .Controls("Lstcnc0").RowSource = VBA.vbNullString
'                Call RES.resExcelxlTABLEADDROWFROMARRAY(excliqlitedatasheet.ListObjects(sAux0), VecAux)
'                Let .Controls("Lstcnc0").RowSource = excliqlitedatasheet.ListObjects(sAux0)
'              Else
'                Let X = .Controls("Lstcnc0").ListIndex + 1
'                Let .Controls("Lstcnc0").RowSource = VBA.vbNullString
'                Let excliqlitedatasheet.ListObjects(sAux0).ListRows(X).Range.Value = VecAux
'                For i = 2 To 8
'                  Let .Controls("Btncnc" & i).Enabled = Not (i = 3 Or i = 4)
'                Next i
'                Let .Controls("Cbocnc1").Enabled = True
'                Let .Controls("Lstcnc0").Enabled = True
'                Let .Controls("Btnmain1").Enabled = True
'                Let .Controls("Lstcnc0").RowSource = excliqlitedatasheet.ListObjects(sAux0)
'              End If
'
'              'Let .Controls("Lstcnc0").Tag = VBA.vbNullString
'              'Let .Controls("Lstcnc0").Tag = "noevent"
'              Let i = .Controls("Cbocnc1").ListIndex
'              Let sAux0 = IIf(i = 3, "Cbocnc2", "Cbocnc1")
'              Let i = .Controls(sAux0).ListIndex
'              Let .Controls(sAux0).ListIndex = -1
'              Let .Controls(sAux0).ListIndex = i
'              Let .Controls("Lstcnc0").Tag = VBA.vbNullString
'
'
'              Let VecAux = Empty
'              Set RES = Nothing
'
'            Case "Btncnc4" 'Cancel Edition
'              For i = 2 To 8
'                Let .Controls("Btncnc" & i).Enabled = Not (i = 3 Or i = 4)
'              Next i
'              Let .Controls("Cbocnc1").Enabled = True
'              Let .Controls("Lstcnc0").Enabled = True
'              Let .Controls("Btnmain1").Enabled = True
'
'              Let .Controls("Lstcnc0").Tag = "noevent"
'              Let i = .Controls("Cbocnc1").ListIndex
'              Let sAux0 = IIf(i = 3, "Cbocnc2", "Cbocnc1")
'              Let i = .Controls(sAux0).ListIndex
'              Let .Controls(sAux0).ListIndex = -1
'              Let .Controls(sAux0).ListIndex = i
'              Let .Controls("Lstcnc0").Tag = VBA.vbNullString
'
'            Case "Btncnc5", "Btncnc6" 'Del once & Del all
'              Select Case .Controls("Cbocnc1").ListIndex
'                Case 0: Let sAux0 = "excliqliteconpurcontents" 'Contents
'                Case 1: Let sAux0 = "excliqliteconpurrefination" 'Refinement
'                Case 3 'Deductions
'                  If Not .Controls("Cbocnc2").ListIndex = 3 Then GoTo EH
'                  Let sAux0 = "excliqliteconpurdeddiscounts" 'Retentions others
'                  'Select Case .Controls("Cbocnc2").ListIndex
'                  '  'Case 0: Let sAux0 = "excliqliteconpurdedexp" 'Exportations
'                  '  'Case 3: Let sAux0 = "excliqliteconpurdedrollback" 'Rollback
'                  '  Case 3: Let sAux0 = "excliqliteconpurdeddiscounts" 'Retentions others
'                  '  Case Else: GoTo EH
'                  'End Select
'                Case 4: Let sAux0 = "excliqliteconpurbonus" 'Bonus
'                Case Else: GoTo EH
'              End Select
'              If VCtrBtn.Name = "Btncnc5" Then
'                Let i = .Controls("Lstcnc0").ListIndex + 1
'                If Not (i > 0 And .Controls("Lstcnc0").ListCount > 0) Then GoTo EH
'                Let .Controls("Lstcnc0").RowSource = VBA.vbNullString
'                Call excliqlitedatasheet.ListObjects(sAux0).ListRows(i).Delete
'                Let .Controls("Lstcnc0").RowSource = excliqlitedatasheet.ListObjects(sAux0)
'              Else
'                If Not .Controls("Lstcnc0").ListCount > 0 Then GoTo EH
'                Let .Controls("Lstcnc0").RowSource = VBA.vbNullString
'                Call excliqlitedatasheet.ListObjects(sAux0).DataBodyRange.Delete
'                Let .Controls("Lstcnc0").RowSource = excliqlitedatasheet.ListObjects(sAux0)
'              End If
'              If sAux0 = "excliqliteconpurcontents" Then 'Contents
'                Let sAux0 = "excliqliteconpurrefination" 'Refinement
'                If Not excliqlitedatasheet.ListObjects(sAux0).DataBodyRange Is Nothing Then
'                  Call excliqlitedatasheet.ListObjects(sAux0).DataBodyRange.Delete
'                End If
'              End If
'              Let .Controls("Btncnc5").Enabled = (excliqlitedatasheet.ListObjects(sAux0).ListRows.Count > 0)
'              Let .Controls("Btncnc6").Enabled = (excliqlitedatasheet.ListObjects(sAux0).ListRows.Count > 0)
'
'            Case "Btncnc7" 'Rangedata
'              Call View_assets_capture_range_cell(FRM, VCtrBtn.Name, VBA.vbNullString, VStrTitle:="Seleccionar datos de " & .Controls("Cbocnc1").Value, VBooJustCaptureRange:=True)
'              Let VCtrBtn.Tag = "validate"
'              Let BooAux = View_Core_validations(VCtrBtn, FRM, VBooSelections, ControlId)
'              'Let BooAux = False '* TESTING
'              Let VCtrBtn.Tag = "selection"
'              If BooAux = False Then
'                Set GLORNG_CELL = Nothing
'                GoTo EH
'
'              Else 'Correct information
'                Select Case .Controls("Cbocnc1").ListIndex
'                  Case 0
'                    Let VecAux = GLORNG_CELL.Value 'BASE 1 ROWS BECAUSE VBA IS A #$"!... VARIANTS FROM RANGES BEGINS AT 1,N
'                    Let sAux0 = "-" & excliqlitedatasheet.ListObjects("excliqliteconfigpurrm").ListColumns("UNIDAD MINERAL").DataBodyRange.Cells(i, 1).Value
'                    Set RES = New AppResources
'                    For i = 1 To UBound(VecAux)
'                      Let sAux1 = VBA.CStr(VecAux(i, 1))
'                      Let sAux0 = RES.resExcelVLockup(sAux1, excliqlitedatasheet.ListObjects("excliqliteconfigpurrm").ListColumns("MINERAL").DataBodyRange, excliqlitedatasheet.ListObjects("excliqliteconfigpurrm").ListColumns("UNIDAD MINERAL").DataBodyRange)
'                      Let sAux1 = VBA.CStr(VecAux(i, 2))
'                      Let sAux1 = sAux1 & "-" & sAux0
'                      Let VecAux(i, 2) = sAux1
'                    Next i
'                    Set RES = Nothing
'                    Let sAux0 = "excliqliteconpurcontents" 'Contents
'
'                  Case 1
'                    Let sAux0 = "excliqliteconpurcontents"
'                    ReDim VecAux(1 To GLORNG_CELL.Rows.Count, 1 To 3)
'                    For i = 1 To UBound(VecAux)
'                      Let VecAux(i, 1) = "Refinación " & excliqlitedatasheet.ListObjects(sAux0).ListColumns("CONTENIDOS").DataBodyRange.Cells(i, 1).Value
'                      Let VecAux(i, 2) = GLORNG_CELL.Cells(i, 1).Value 'BASE 1 ROWS BECAUSE VBA IS A #$"!... VARIANTS FROM RANGES BEGINS AT 1,N
'                      Let VecAux(i, 3) = GLORNG_CELL.Cells(i, 2).Value 'BASE 1 ROWS BECAUSE VBA IS A #$"!... VARIANTS FROM RANGES BEGINS AT 1,N
'                    Next i
'                    Let sAux0 = "excliqliteconpurrefination" 'Reinement
'                    'Let VecAux = GLORNG_CELL.Value 'BASE 1 ROWS BECAUSE VBA IS A #$"!... VARIANTS FROM RANGES BEGINS AT 1,N
'                  Case 3 'Deductions
'                    If Not .Controls("Cbocnc2").ListIndex = 3 Then GoTo EH
'                    Let sAux0 = "excliqliteconpurdeddiscounts" 'Retentions others
'                    Let VecAux = GLORNG_CELL.Value 'BASE 1 ROWS BECAUSE VBA IS A #$"!... VARIANTS FROM RANGES BEGINS AT 1,N
'                    'Select Case .Controls("Cbocnc2").ListIndex
'                    '  Case 0: Let sAux0 = "excliqliteconpurdedexp" 'Exportations
'                    '  Case 3: Let sAux0 = "excliqliteconpurdedrollback" 'Rollback
'                    '  Case 4: Let sAux0 = "excliqliteconpurdeddiscounts" 'Retentions others
'                    '  Case Else: GoTo EH
'                    'End Select
'                  Case 4
'                    Let sAux0 = "excliqliteconpurbonus" 'Bonus
'                    Let VecAux = GLORNG_CELL.Value 'BASE 1 ROWS BECAUSE VBA IS A #$"!... VARIANTS FROM RANGES BEGINS AT 1,N
'                  Case Else: GoTo EH
'                End Select
'                'GoTo EH:
'
'                Let .Controls("Lstcnc0").Tag = "noevent"
'                Set RES = New AppResources
'                Let .Controls("Lstcnc0").RowSource = VBA.vbNullString
'
'                #If Debugging Then
'                  Debug.Print "Filas: ", excliqlitedatasheet.ListObjects(sAux0).ListRows.Count
'                #End If
'
'                Call RES.resExcelxlTABLEADDROWFROMARRAY(excliqlitedatasheet.ListObjects(sAux0), VecAux)
'                Set RES = Nothing
'                Let .Controls("Lstcnc0").RowSource = excliqlitedatasheet.ListObjects(sAux0)
'                Let .Controls("Lstcnc0").Tag = VBA.vbNullString
'                Let .Controls("Btncnc5").Enabled = True
'                Let .Controls("Btncnc6").Enabled = True
'              End If
'
'            Case "Btncnc8"
'              Call View_assets_capture_range_cell(FRM, VCtrBtn.Name, VBA.vbNullString, VStrTitle:="Seleccionar datos base", VBooJustCaptureRange:=True)
'              Let VCtrBtn.Tag = "validate"
'              Let BooAux = View_Core_validations(VCtrBtn, FRM, VBooSelections, ControlId)
'              'Let BooAux = False '* TESTING
'              Let VCtrBtn.Tag = "selection"
'
'              If BooAux = False Then
'                Set GLORNG_CELL = Nothing
'                GoTo EH
'              Else 'Correct information
'                Let i = 0
'                Let VecAux = Array("Txtcnc1", "Txtcnc2", "Txtcnc3", "Chkcnc0", "Txtcnc4", "Cbocnc0", "Txtcnc5", "Txtcnc6", "Txtcnc7", "Txtcnc8", "Txtcnc9", "Txtcnc10", "Txtcnc11")
'                For Each RngBox In GLORNG_CELL
'                  Let sAux0 = VBA.CStr(VecAux(i))
'                  If i = 3 Then 'Chk
'                    Let .Controls(sAux0).Value = VBA.CBool(RngBox.Value)
'                  ElseIf i = 5 Then 'Cbo
'                    'Let .Controls(sAux0).ListIndex = IIf(VBA.LCase(RngBox.Value) = "particular", 1, IIf(VBA.LCase(RngBox.Value) = "cooperativista", 0, IIf(VBA.CBool(RngBox.Value), 1, 0)))
'                    Let .Controls(sAux0).ListIndex = IIf(VBA.LCase(RngBox.Value) = "particular", 1, 0)
'                  Else
'                    Let .Controls(sAux0).Value = RngBox.Value
'                  End If
'                  Let i = i + 1
'                Next RngBox
'                Set RngBox = Nothing
'              End If
'
'            Case "Btnmain3" 'Rangecell
'              If Not .Controls("Optcnc1").Value Then GoTo EH
'              Call View_assets_capture_range_cell(FRM, VCtrBtn.Name, "Txtcnc0")
'
'            Case "Btnmain1" 'Exe
'              'TODO: COLOCAR LOS VALORES DE LOS CONTROLES BASE A LA TABLA MAIN
'              Let VecAux = Array("SIMULACIÓN", "SIMULACIÓN", VBA.CDate(.Controls("Txtcnc1").Value), .Controls("Txtcnc2").Value, VBA.CDbl(.Controls("Txtcnc3").Value), .Controls("Chkcnc0").Value, .Controls("Txtcnc4").Value, "N/D", .Controls("Cbocnc0").Value, "N/D", "N/D", "N/D", "N/D", "N/D", "N/D", VBA.CDate(.Controls("Txtcnc5").Value), VBA.CDbl(.Controls("Txtcnc6").Value), VBA.CDbl(.Controls("Txtcnc7").Value), VBA.CDbl(.Controls("Txtcnc8").Value), VBA.CDbl(.Controls("Txtcnc9").Value), VBA.CDbl(.Controls("Txtcnc10").Value), VBA.CDbl(.Controls("Txtcnc11").Value))
'              If Not excliqlitedatasheet.ListObjects("excliqliteconpurbase").DataBodyRange Is Nothing Then
'                Call excliqlitedatasheet.ListObjects("excliqliteconpurbase").DataBodyRange.Delete
'              End If
'              Set RES = New AppResources
'              Call RES.resExcelxlTABLEADDROWFROMARRAY(excliqlitedatasheet.ListObjects("excliqliteconpurbase"), VecAux, True)
'              Set RES = Nothing
'              Let View_Controls_buttons_click = View_Core_execution(FRM.Tag, FRM)
'
'            Case Else: GoTo EH
'          End Select
'        End With
'
'      Case "configlauncher"
'        Select Case VCtrBtn.Name
'          'DIVISA 'Add, Upd, Ucn, Del, Local, Foreign
'          Case "Btncon0", "Btncon1", "Btncon2", "Btncon3", "Btncon4", "Btncon5", _
'               "Btncon6", "Btncon7", "Btncon8", "Btncon9", "Btncon10", "Btncon11", _
'               "Btncon12", "Btncon13", "Btncon14": Let View_Controls_buttons_click = View_Core_execution(FRM.Tag, FRM, VCtrBtn.Name)
'
'          'Case "Btncon1" 'Upd (Divisa)
'          'Case "Btncon2" 'Ucn (Divisa)
'          'Case "Btncon3" 'Del (Divisa)
'          'Case "Btncon4" 'Local (Divisa)
'          'Case "Btncon5" 'Foreign (Divisa)
'        End Select
      
      Case Else: GoTo EH
    End Select
      
  
  
  
  
  Else ' Selections tasks
    
    Select Case ControlId
      
      Case "regressionselection" ' Go to get a Cellrange
        ' 1.  Evaluates first validation with direct selection.
        ' 2.  In validations, avoid evaluation of RngCell in first instance 'cause it isn't captured yet and its value is nothing
        ' 3.  Capture the RngCell and make a new call to validations, this time, send RngCell, if evaluates, call exe
        ' Select RngCell
        If Not GLOBOO_SELECCTIONCELL Then Let View_Controls_buttons_click = View_Core_execution(ControlId, Nothing): GoTo EH ' SELECTIONCELLS IN RIBBON IS INACTIVE, SO, NEED TO CREATE A NEW SHEET
        Set VWP = New ViewsParent_current ' TODO: ASK MDL IF SELECTIONCELLS IS ACTIVE, ELSE, CREATE NEW SHEET IF FORMULAS IS ACTIVE
        Set VWP.ErrorHandler = EHGLOBAL
        Set RngBox = VWP.ViewsCaptureRange(FRM, "Seleccione una celda donde generar la Regresión Lineal:", STRAPPTITLE)
        Set VWP = Nothing
        ' Validate
        If View_Core_validations(Nothing, Nothing, True, ControlId & "-check_cell", RngBox) Then Let View_Controls_buttons_click = View_Core_execution(ControlId, Nothing, RngBox)
        
        
        
        
'      Case "regressionselection-getcell" ' Go to get a Cellrange
'        Set RngBox = Application.Selection
'        If RngBox Is Nothing Then GoTo EH
'        Let sAux0 = "Seleccione datos iniciales para la Regresión. " & VBA.vbNewLine & VBA.vbNewLine & _
'                     "Dato 1: Variables 'Y', puede ser un rango o un área." & VBA.vbNewLine & _
'                     "Dato 2 (Opcional): Variables 'X', puede ser un rango o un área." & VBA.vbNewLine & _
'                     "Dato 3 (Opcional): Valores a extrapolar, puede ser un rango o un área" & VBA.vbNewLine & VBA.vbNewLine & _
'                     "Recomendado, datos en columnas por cada valor."
'        If Application.WorksheetFunction.Count(RngBox) < 3 Then Call VBA.MsgBox(sAux0, vbInformation, STRAPPTITLE): GoTo EH
'        Set RngBox = Nothing
'
'        If Not GLOBOO_SELECCTIONCELL Then Let View_Controls_buttons_click = True: GoTo EH
'        Set VWP = New ViewsParent_current
'        Set VWP.ErrorHandler = EHGLOBAL
'        Set GLORNG_CELL = VWP.ViewsCaptureRange(FRM, STR_TITLE8 & " para generar una Regresión Lineal:", STRAPPTITLE)
'        Set VWP = Nothing
'        Let View_Controls_buttons_click = (Not GLORNG_CELL Is Nothing)

'      Case "regressionselection"
'        Let View_Controls_buttons_click = View_Core_execution(ControlId, Nothing)
'
'      Case "balmetselection-getcell" 'Go to get a Cellrange
'        Set RngBox = Selection
'        If RngBox Is Nothing Then GoTo EH
'        Let sAux0 = "Seleccione datos iniciales para el Balance Metalúrgico. " & VBA.vbNewLine & VBA.vbNewLine & _
'                     "Dato 1: Peso Alimentación" & VBA.vbNewLine & _
'                     "Dato 2: Leyes de Productos y Alimentación (Puede incluir leyes de Minerales complejos)." & VBA.vbNewLine & _
'                     "Dato 3 (Opcional): Las Unidades de las leyes de cada Producto. Ej. %" & VBA.vbNewLine & _
'                     "Dato 4 (Opcional): Los Nombres de los Productos. Ej. Cu, Zn, Pb." & VBA.vbNewLine & VBA.vbNewLine & _
'                     "Si omite los Nombres y Unidades, se asumirán nombres genéricos y como unidades porcentajes (%) para Balances generales y decimarcos (DM) para Minerales Complejos." & VBA.vbNewLine & _
'                     "Recomiendado, datos en columnas por cada Producto."
'        If Application.WorksheetFunction.Count(RngBox) < 4 Then Call VBA.MsgBox(sAux0, vbInformation, STRAPPTITLE): GoTo EH
'        Set RngBox = Nothing
'
'        If Not GLOBOO_SELECCTIONCELL Then Let View_Controls_buttons_click = True: GoTo EH
'        Set VWP = New ViewsParent_current
'        Set VWP.ErrorHandler = EHGLOBAL
'        Set GLORNG_CELL = VWP.ViewsCaptureRange(FRM, STR_TITLE8 & " para imprimir el Balance Metalúrgico:", STRAPPTITLE)
'        Set VWP = Nothing
'        Let View_Controls_buttons_click = (Not GLORNG_CELL Is Nothing)
'
'      Case "balmetselection"
'        Let View_Controls_buttons_click = View_Core_execution(ControlId, Nothing)

      Case Else: GoTo EH
    
    End Select
  
  End If

EH:
  Let VecAux = Empty
  Let VecAux1 = Empty
  Set RngBox = Nothing
  Set VWP = Nothing
  Set REGEX = Nothing
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerDisplay(sFN)

End Function

'Events handlers - Checkboxes
Private Sub View_Controls_checkboxes_change( _
  ByVal VCtrChk As MSForms.CheckBox, _
  ByVal FRM As FRMBasics)
  
  Dim RES As AppResources_current
  Dim REGEX As AppResRegEx
  Dim RCA As AppResourcesController_current
  
  Dim VecAux0 As Variant, VecAux1 As Variant
  Dim StrAux0 As String, StrAux1 As String, StrAux2 As String
  Dim LngX As Long
  Dim i As Single
  
  Const sFN As String = "VIEW::-checkboxes_change"
  
  On Error GoTo EH
  
  With FRM
    
    Select Case .Tag
      
      Case "balmet"
        Select Case VCtrChk.Name
          Case "Chkbal0"
            Let .Controls("Txtbal1").Locked = Not VCtrChk.Value
            Let .Controls("Txtbal1").ForeColor = IIf(VCtrChk.Value, .ForeColor, VBA.vbGrayText)
            Let .Controls("Txtbal1").Text = IIf(VCtrChk.Value, VBA.vbNullString, "A")
            If VCtrChk.Value Then Call .Controls("Txtbal1").SetFocus
        
          Case "Chkbal1"
            Let .Controls("Cbobal1").Enabled = VCtrChk.Value
            Let .Controls("Btnbal0").Enabled = VCtrChk.Value
            Let .Controls("Btnbal1").Enabled = VCtrChk.Value
            With .Controls("Txtbal2")
              Let StrAux0 = VBA.vbNullString
              If Not VCtrChk.Value Then
                Set REGEX = New AppResRegEx
                If REGEX.isNumberEntireAndDecimalVectorColonSeparated(.Text) Then
                  Set RCA = New AppResourcesController_current
                  Set RCA.ErrorHandler = EHGLOBAL
                  Let StrAux0 = RCA.ResCtrGetGenericNames(.Text)
                  Set RCA = Nothing
                End If
                Set REGEX = Nothing
              End If
            End With
            If VCtrChk.Value Then Call .Controls("Cbobal1").SetFocus
            Let .Controls("Txtbal3").ForeColor = IIf(VCtrChk.Value, .ForeColor, VBA.vbGrayText)
            Let .Controls("Txtbal3").Text = StrAux0
        
          Case "Chkbal2"
            Let .Controls("Cbobal2").Enabled = VCtrChk.Value
            Let .Controls("Btnbal2").Enabled = VCtrChk.Value
            Let .Controls("Btnbal3").Enabled = VCtrChk.Value
            With .Controls("Txtbal2")
              Let StrAux0 = VBA.vbNullString
              If Not VCtrChk.Value Then
                Set REGEX = New AppResRegEx
                If REGEX.isNumberEntireAndDecimalVectorColonSeparated(.Text) Then
                  Set RCA = New AppResourcesController_current
                  Set RCA.ErrorHandler = EHGLOBAL
                  Let StrAux0 = RCA.ResCtrGetGenericNames(.Text, BBooUnits:=True)
                  Set RCA = Nothing
                End If
                Set REGEX = Nothing
              End If
            End With
            If VCtrChk.Value Then Call .Controls("Cbobal2").SetFocus
            Let .Controls("Txtbal4").ForeColor = IIf(VCtrChk.Value, .ForeColor, VBA.vbGrayText)
            Let .Controls("Txtbal4").Text = StrAux0
        
        End Select
      
      Case "oreinvoice"
        Select Case VCtrChk.Name
          Case "Chkpur1"
'            Let .Controls("Txtpur4").Visible = VCtrChk.Value
'            Let .Controls("Btnpur1").Visible = VCtrChk.Value
'            Let .Controls("Btnpur2").Visible = VCtrChk.Value
'            Let .Controls("Cbopur0").Visible = Not VCtrChk.Value
'            Let .Controls("Cbopur1").ListIndex = 0
'            Let .Controls("Chkpur2").Value = False
'            Let .Controls("Chkpur3").Enabled = Not VCtrChk.Value
'            Let .Controls("Chkpur3").Value = False
'
'            If .Controls("Lstpur0").ListCount > 0 Then
'              Set RES = New AppResources
'              If Not excliqlitedatasheet.ListObjects("excliqlitepurchasevariousretentions").DataBodyRange Is Nothing Then
'                Call excliqlitedatasheet.ListObjects("excliqlitepurchasevariousretentions").DataBodyRange.Delete
'                Call excliqlitedatasheet.ListObjects("excliqlitepurchasevariousretentionsother").DataBodyRange.Delete
'              End If
'              If VCtrChk.Value Then
'                Call RES.resExcel_TABLE_ADD_ROWS_EMPTIES(excliqlitedatasheet.ListObjects("excliqlitepurchasevariousretentions"), .Controls("Lstpur0").ListCount)
'                Call RES.resExcel_TABLE_ADD_ROWS_EMPTIES(excliqlitedatasheet.ListObjects("excliqlitepurchasevariousretentionsother"), .Controls("Lstpur0").ListCount)
'              End If
'              Set RES = Nothing
'            End If

          Case Else: GoTo EH
        End Select
    
      Case "configlauncher"
'        Select Case VCtrChk.Name
'          Case "Chkcon0", "Chkcon1", "Chkcon2", "Chkcon3": Call View_Core_execution(.Tag, FRM)
'        End Select
      
      Case Else: GoTo EH
    End Select
  End With

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Set RCA = Nothing
  Set REGEX = Nothing
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerDisplay(sFN)

End Sub

'Events handlers - Checkboxes
Private Sub View_Controls_checkboxes_click( _
  ByVal VCtrChk As MSForms.CheckBox, _
  ByVal FRM As FRMBasics)
  
  Dim RES As AppResources
  
  Dim VecAux As Variant, VecAux1 As Variant
  Dim StrAux0 As String, StrAux1 As String, StrAux2 As String
  Dim LngX As Long
  Dim i As Single
  
  Const sFN As String = "VIEW::-controls_checkboxes_click"
  
  On Error GoTo EH
  
  'Execute task
  With FRM
    Select Case .Tag
      Case "oreinvoice"
        Select Case VCtrChk.Name
          Case "Chkpur2": Call View_assets_controls_listbox_mannagement(FRM, VCtrChk.Name)
          
          Case Else: GoTo EH
        End Select
    
      Case Else: GoTo EH
    End Select
  End With

EH:
  Let VecAux = Empty
  Let VecAux1 = Empty
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerDisplay(sFN)
End Sub


'Events handlers - Combos
Private Sub View_Controls_combos_change(ByVal xCombo As MSForms.ComboBox, ByVal FRM As FRMBasics)
  Dim RES As AppResources
  Dim REGEX As AppResRegEx
  
  Dim xControl As MSForms.Control
  Dim LO1 As ListObject, LO2 As ListObject
  Dim Row0 As ListRow, Row1 As ListRow, Row2 As ListRow
  
  Dim VecAux As Variant
  Dim Controlis() As Variant, aControlis() As Variant, aControlisI() As Variant
  Dim xMsg As String, xIndex As String, yIndex As String, zIndex As String, StrAux As String
  Dim ixIndex As Long, iyIndex As Long, izIndex As Long
  Dim i As Long, j As Long, k As Long, l As Long, m As Long, n As Long, o As Long, CurIndex As Currency
  Dim bIndex As Boolean, cIndex As Boolean, BooAux As Boolean, BooAux1 As Boolean
  
  Const sFN As String = "VIEW::-ComboEvents"

  On Error GoTo EH
  #If Debugging Then
    Call Debug_assert(FRM Is Nothing, sFN)
  #End If
  If FRM Is Nothing Then GoTo EH
    
  Select Case FRM.Tag
    Case "balmet"
      With FRM
        Select Case xCombo.Name
          Case "Cbobal1"
            Let .Controls("Cbobal0").ListIndex = IIf(xCombo.ListIndex = 1, 0, .Controls("Cbobal0").ListIndex)
            Let .Controls("Cbobal0").Enabled = Not (xCombo.ListIndex = 1)
            For Each xControl In .Controls("Framebal1").Controls
              Let xControl.Enabled = Not (xCombo.ListIndex = 1)
              If xControl.Name = "Chkbal6" Then Let xControl.Value = False
            Next xControl
            Set xControl = Nothing
          
          Case Else: GoTo EH
        End Select
      End With
    
    Case "balmeteco"
      Select Case xCombo.Name
        Case "Cbobal0"
          With FRM
            Let .Controls("Lblbal4").Caption = IIf(xCombo.ListIndex = 0, "Ley Producto/s (Concentrado):", "Ley Cabeza/s (Alimentación):")
          End With

        Case "Cbobal2"
          With FRM
            For Each xControl In .Controls("Framebal1").Controls
              Let xControl.Enabled = Not (xCombo.ListIndex = 1)
              If xControl.Name = "Chkbal5" Then Let xControl.Value = False
            Next xControl
            Set xControl = Nothing
            If (xCombo.ListIndex = 1) Then Let .Controls("Chkbal2").Value = False
            'Let .Controls("Txtbal6").Value = IIf((xCombo.ListIndex = 1), xCombo.Value, .Controls("Txtbal6").Tag)
            If (xCombo.ListIndex = 1) Then Let .Controls("Txtbal6").Value = xCombo.Value
            Let .Controls("Chkbal2").Enabled = Not (xCombo.ListIndex = 1)
            Let .Controls("Btnbal8").Enabled = Not (xCombo.ListIndex = 1)
            Let .Controls("Btnbal9").Enabled = Not (xCombo.ListIndex = 1)
          End With

        Case "Cbobal3"
          With FRM
            If Not .Controls("Chkbal4").Value Then GoTo EH
            If xCombo.ListIndex = 0 Then Let .Controls("Cbobal4").List = Array("DM", "g/T", "g/L")
            If xCombo.ListIndex > 0 Then Let .Controls("Cbobal4").List = Array("g/T", "g/L")
            Let .Controls("Cbobal4").ListIndex = 0
          End With
        
        Case Else: GoTo EH
      End Select

    Case "projection"
      Select Case xCombo.Name
        Case "Cbopro1", "Cbopro2"
          With FRM
            Let xIndex = IIf(xCombo.Name = "Cbopro1", "Txtpro4", "Txtpro5")
            Set REGEX = New AppResRegEx
            If Not REGEX.isEmptyStringReg(.Controls(xIndex).Value) Then
              Let VecAux = VBA.Split(.Controls(xIndex), ";")
              ReDim Preserve VecAux(UBound(VecAux) + 1)
              Let VecAux(UBound(VecAux)) = xCombo.Value
              Let .Controls(xIndex).Value = VBA.Join(VecAux, ";")
              Let VecAux = Empty
            Else
              Let .Controls(xIndex).Value = xCombo.Value
            End If
            Set REGEX = Nothing
          End With
          
        Case Else: GoTo EH
      End Select
      
    Case "oreinvoice"
      With FRM
        Select Case xCombo.Name
          Case "Cbopur1"
            For i = 5 To 9
              Let .Controls("Txtpur" & i).Value = VBA.vbNullString
              Let .Controls("Txtpur" & i).Visible = Not (xCombo.ListIndex > 2 And xCombo.ListIndex < 6)
            Next i
            Let .Controls("Btnpur9").Visible = (xCombo.ListIndex = 0)
            Let .Controls("Txtpur6").Locked = False
            Let .Controls("Txtpur8").Locked = False
            Let .Controls("Chkpur2").Visible = False
            Select Case xCombo.ListIndex
              Case 0 'Base Data
                For i = 5 To 9
                  Let .Controls("Txtpur" & i).Width = (.Controls("Frame2").Width * 0.95) / 5
                  Let .Controls("Txtpur" & i).Visible = True
                Next i
                Let .Controls("Txtpur6").Left = .Controls("Txtpur5").Left + .Controls("Txtpur5").Width
                Let .Controls("Txtpur6").Width = .Controls("Txtpur5").Width - .Controls("Txtpur6").Height
                Let .Controls("Txtpur6").Value = VBA.Format$(VBA.Date(), "dd/mm/yy")
                
                Let .Controls("Txtpur5").ControlTipText = "El nombre de la muestra del mineral. Ejemplo: M-1"
                Let .Controls("Txtpur6").ControlTipText = "Ejemplo: dd/mm/aa o dd/mm/aaaa"
                Let .Controls("Txtpur7").ControlTipText = "El Porcentaje de humedad del mineral. Ejemplo: 5"
                Let .Controls("Txtpur8").ControlTipText = "El Porcentaje de Merma del mineral. Ejemplo: 1"
                Let .Controls("Txtpur9").ControlTipText = "El Peso del mineral en Toneladas. Ejemplo: 1000"
              
              Case 1 'Main Data
                For i = 5 To 9
                  Let .Controls("Txtpur" & i).Width = (.Controls("Frame2").Width * 0.95) / 5
                  Let .Controls("Txtpur" & i).Visible = True
                Next i
                Let .Controls("Txtpur5").Visible = False
                Let .Controls("Cbopur2").Visible = True
                Let .Controls("Cbopur3").Visible = True
                Let .Controls("Cbopur4").Visible = True
                Let .Controls("Txtpur6").Left = .Controls("Txtpur5").Left + .Controls("Txtpur5").Width
                Let .Controls("Cbopur2").RowSource = VBA.vbNullString
                Let .Controls("Cbopur2").List = Array("Mineral", "Mineral Complejo", "Brosa")
                Let .Controls("Cbopur2").ListIndex = -1
                Let .Controls("Cbopur2").ListIndex = 0
              
                Let .Controls("Cbopur2").ControlTipText = "Seleccione el tipo de mineral"
                Let .Controls("Cbopur3").ControlTipText = "Seleccione el mineral o los minerales"
                Let .Controls("Cbopur4").ControlTipText = "Seleccione la Unidad de ley del mineral"
                Let .Controls("Txtpur6").ControlTipText = "Seleccione un mineral de la lista"
                Let .Controls("Txtpur7").ControlTipText = "Introduzca la ley o las leyes del mineral (separados por punto y coma si son complejos). Ejemplo: 5"
                Let .Controls("Txtpur8").ControlTipText = "Seleccione una unidad de ley de la lista"
                Let .Controls("Txtpur9").ControlTipText = "El Precio del mineral. Ejemplo: 3.5"
              
              'Case 2, 3, 4, 5 'BONUS, RETENTIONS, OTHER, DISCOUNTS
              Case 3, 4 'RETENTIONS, OTHER
                For i = 5 To 9
                  Let .Controls("Txtpur" & i).Visible = False
                Next i
                Let .Controls("Cbopur2").Visible = Not (xCombo.ListIndex > 2 And xCombo.ListIndex < 5) 'False
                Let .Controls("Cbopur3").Visible = Not (xCombo.ListIndex > 2 And xCombo.ListIndex < 5) 'False
                Let .Controls("Cbopur4").Visible = Not (xCombo.ListIndex > 2 And xCombo.ListIndex < 5) 'False
                Let .Controls("Txtpur5").Visible = Not (xCombo.ListIndex > 2 And xCombo.ListIndex < 5) 'True
                Let .Controls("Txtpur6").Visible = Not (xCombo.ListIndex > 2 And xCombo.ListIndex < 5) 'True
                Let .Controls("Txtpur5").Width = (.Controls("Frame2").Width * 0.95) / 2
                Let .Controls("Txtpur6").Width = .Controls("Txtpur5").Width
                Let .Controls("Txtpur6").Left = .Controls("Txtpur5").Left + .Controls("Txtpur5").Width
                Let .Controls("Chkpur2").Visible = False
'                If xCombo.ListIndex = 2 Then
'                  Let .Controls("Txtpur5").ControlTipText = "El nombre del Bono. Ejemplo: Bono transporte"
'                  Let .Controls("Txtpur6").ControlTipText = "El valor del Bono. Ejemplo: 500"
'                End If
'                If xCombo.ListIndex = 5 Then
'                  Let .Controls("Txtpur5").ControlTipText = "El nombre del Descuento. Ejemplo: Ensayes químicos"
'                  Let .Controls("Txtpur6").ControlTipText = "El valor del Descuento. Ejemplo: 500"
'                End If
                If xCombo.ListIndex = 3 Or xCombo.ListIndex = 4 Then 'Retentions or Other retentions
                  If .Controls("Chkpur1").Value Then 'Show controls if Multiple is selected
                    Let .Controls("Chkpur2").Left = .Controls("Cbopur2").Left + .Controls("Cbopur2").Width + 5
                    Let .Controls("Chkpur2").Visible = True
                    Let .Controls("Chkpur2").Caption = "Todas las " & .Controls("Cbopur1").Value
                    Let .Controls("Cbopur2").Visible = True
                    
                    Set LO1 = excliqlitedatasheet.ListObjects("excliqlitepurchasebasedata")
                    If LO1 Is Nothing Then GoTo EH
                    If LO1.ListColumns("MUESTRA").DataBodyRange Is Nothing Then GoTo EH
                    Let .Controls("Cbopur2").RowSource = LO1 & "[MUESTRA]"
                    Set LO1 = Nothing
                    Let .Controls("Cbopur2").ListIndex = 0
                    Let .Controls("Cbopur2").ControlTipText = "Seleccione el código de muestra de cada liquidación y asigne datos de " & .Controls("Cbopur1").Value
                  Else
                    Let .Controls("Chkpur2").Left = .Controls("Cbopur2").Left
                    Let .Controls("Chkpur2").Visible = True
                    Let .Controls("Chkpur2").Caption = "Todas las " & .Controls("Cbopur1").Value
                  End If
                End If
              
              Case 2, 5 'BONUS, DISCOUNTS
                For i = 5 To 9
                  Let .Controls("Txtpur" & i).Visible = False
                Next i
                Let .Controls("Cbopur2").Visible = Not (xCombo.ListIndex > 2 And xCombo.ListIndex < 6) 'False
                Let .Controls("Cbopur3").Visible = Not (xCombo.ListIndex > 2 And xCombo.ListIndex < 6) 'False
                Let .Controls("Cbopur4").Visible = Not (xCombo.ListIndex > 2 And xCombo.ListIndex < 6) 'False
                Let .Controls("Txtpur5").Visible = Not (xCombo.ListIndex > 2 And xCombo.ListIndex < 6) 'True
                Let .Controls("Txtpur6").Visible = Not (xCombo.ListIndex > 2 And xCombo.ListIndex < 6) 'True
                Let .Controls("Txtpur5").Width = (.Controls("Frame2").Width * 0.95) / 2
                Let .Controls("Txtpur6").Width = .Controls("Txtpur5").Width
                Let .Controls("Txtpur6").Left = .Controls("Txtpur5").Left + .Controls("Txtpur5").Width
                If xCombo.ListIndex = 2 Then
                  Let .Controls("Txtpur5").ControlTipText = "El nombre del Bono. Ejemplo: Bono transporte"
                  Let .Controls("Txtpur6").ControlTipText = "El valor del Bono. Ejemplo: 500"
                End If
                If xCombo.ListIndex = 5 Then
                  Let .Controls("Txtpur5").ControlTipText = "El nombre del Descuento. Ejemplo: Ensayos químicos"
                  Let .Controls("Txtpur6").ControlTipText = "El valor del Descuento. Ejemplo: 500"
                End If
                Let .Controls("Lstpur0").Enabled = True
              
              Case Else: GoTo EH
            End Select
            
            Let .Controls("Lstpur0").Visible = False
            Let .Controls("Lstpur0").RowSource = VBA.vbNullString
            Select Case xCombo.ListIndex
              Case 0 'Data A
                Let j = 5: Let StrAux = "excliqlitepurchasebasedata"
                Let BooAux = (excliqlitedatasheet.ListObjects(StrAux).ListRows.Count = GLOSNG_PUR_LIMIT)
                Let .Controls("Lstpur0").RowSource = excliqlitedatasheet.ListObjects(StrAux)
              Case 1 'Data B
                Let j = 5: Let StrAux = "excliqlitepurchasemaindata"
                Let BooAux = (excliqlitedatasheet.ListObjects(StrAux).ListRows.Count = GLOSNG_PUR_LIMIT)
                Let .Controls("Lstpur0").RowSource = excliqlitedatasheet.ListObjects(StrAux)
              Case 2 'Bonus
                Let j = 2: Let StrAux = "excliqlitepurchasebonus"
                Let .Controls("Lstpur0").RowSource = excliqlitedatasheet.ListObjects(StrAux)
              Case 3 'Retentions
                Let j = 2
                Let .Controls("Lstpur0").RowSource = excliqlitedatasheet.ListObjects("excliqliteconfigpurretentions")
              Case 4 'Other retentions
                Let j = 2
                Let .Controls("Lstpur0").RowSource = excliqlitedatasheet.ListObjects("excliqliteconfigpurretentionsother")
              Case 5 'Discounts
                Let j = 2: Let StrAux = "excliqlitepurchasediscounts"
                Let BooAux = (excliqlitedatasheet.ListObjects(StrAux).ListRows.Count = GLOSNG_PUR_LIMIT)
                Let .Controls("Lstpur0").RowSource = excliqlitedatasheet.ListObjects(StrAux)
              Case Else: GoTo EH
            End Select
            Let .Controls("Lstpur0").ColumnCount = j
            Let .Controls("Lstpur0").ColumnWidths = .Controls("Lstpur0").Width / j
            Let .Controls("Lstpur0").ColumnHeads = True
            Let j = xCombo.ListIndex
            Set REGEX = New AppResRegEx
            Let BooAux1 = (REGEX.isEmptyStringReg(VBA.CStr(.Controls("Lstpur0").List(0, 0))))
            Set REGEX = Nothing
            Let .Controls("Lstpur0").MultiSelect = IIf(j > 2 And j < 5, fmMultiSelectMulti, fmMultiSelectSingle)
            Let .Controls("Lstpur0").ListStyle = IIf(j > 2 And j < 5, fmListStyleOption, fmListStylePlain)
            Let .Controls("Lstpur0").Height = .Controls("Frame2").Height * 0.6
            Let .Controls("Lstpur0").Width = .Controls("Frame2").Width * 0.95
            Let .Controls("Btnpur3").Enabled = IIf(j > 2 And j < 6 Or BooAux, False, True)
            Let .Controls("Btnpur4").Enabled = False
            Let .Controls("Btnpur5").Enabled = False
            Let .Controls("Btnpur6").Enabled = IIf(j > 2 And j < 5, False, .Controls("Lstpur0").ListCount > 1)
            Let .Controls("Btnpur7").Enabled = IIf(j > 2 And j < 5, False, .Controls("Lstpur0").ListCount >= 1 And Not BooAux1)
            Let .Controls("Btnpur8").Enabled = IIf(j > 2 And j < 5 Or BooAux, False, True)
            Let .Controls("Lstpur0").Visible = True
            Let .Controls("Lstpur0").ControlTipText = IIf(j > 2 And j < 5, "Seleccione los ítems de " & xCombo.Value & " que desea habilitar", "Doble click en el ítem que desee modificar")
            'Mark all previously selected items in the list
            If (j > 2 And j < 5) Then
              Let BooAux = .Controls("Chkpur1").Value 'If multiple, all items will be unselected
              With .Controls("Lstpur0")
                Let .Tag = "change"
                If .ListCount > 0 Then
                  For i = 0 To .ListCount - 1
                    Let .Selected(i) = IIf(BooAux, False, VBA.CBool(.List(i, 2)))
                  Next i
                End If
                Let .Tag = VBA.vbNullString
              End With
              Let .Controls("Lstpur0").Tag = "change"
              Let .Controls("Chkpur2").Value = False
              Let .Controls("Lstpur0").Tag = VBA.vbNullString
            End If
            
          Case "Cbopur2"
            If .Controls("Cbopur1").ListIndex = 3 Or .Controls("Cbopur1").ListIndex = 4 Then
              Let .Controls("Lstpur0").Tag = "change"
              Let .Controls("Chkpur2").Value = False
              Let .Controls("Lstpur0").Tag = VBA.vbNullString
              Call View_assets_controls_listbox_mannagement(FRM, xCombo.Name)
            Else
              If xCombo.ListIndex > -1 Then
                Let bIndex = (xCombo.ListIndex = 1)
                Let j = xCombo.Width * IIf(bIndex, 0.5, 1)
                Let .Controls("Txtpur6").Width = j
                Let .Controls("Txtpur8").Width = j
                Let .Controls("Cbopur3").Width = j
                Let .Controls("Cbopur3").Left = .Controls("Txtpur6").Left + IIf(bIndex, j, 0)
                Let .Controls("Cbopur4").Width = j
                Let .Controls("Cbopur4").Left = .Controls("Txtpur8").Left + IIf(bIndex, j, 0)
                Let .Controls("Txtpur6").Visible = bIndex
                Let .Controls("Txtpur8").Visible = bIndex
                If xCombo.ListIndex = 1 Then
                  Set LO1 = excliqlitedatasheet.ListObjects("excliqliteconfigpurrm")
                  If LO1 Is Nothing Then GoTo EH
                  If LO1.ListColumns("MINERAL").DataBodyRange Is Nothing Then GoTo EH
                  If LO1.ListColumns("ELEMENTO").DataBodyRange Is Nothing Then GoTo EH
                  Let i = LO1.ListRows.Count + 1
                  ReDim VecAux(0 To i, 0 To 1)
                  For i = 0 To UBound(VecAux)
                    If i = 0 Then
                      Let VecAux(i, 0) = "Quitar último"
                      Let VecAux(i, 1) = VBA.vbNullString
                    ElseIf i = 1 Then
                      Let VecAux(i, 0) = "Limpiar lista"
                      Let VecAux(i, 1) = VBA.vbNullString
                    Else
                      Let VecAux(i, 0) = LO1.ListColumns("ELEMENTO").DataBodyRange.Cells(i - 1, 1).Value
                      Let VecAux(i, 1) = LO1.ListColumns("MINERAL").DataBodyRange.Cells(i - 1, 1).Value
                    End If
                  Next i
                  Set LO1 = Nothing
                  Let .Controls("Cbopur3").ColumnCount = 2
                  Let .Controls("Cbopur3").List = VecAux
                  Let VecAux = Empty
                  Let .Controls("Cbopur4").List = Array("Quitar último", "Limpiar lista", "DM", "%", "OT", "g/T", "g/L")
                  Let .Controls("Cbopur3").ListIndex = 2
                Else
                  Set LO1 = excliqlitedatasheet.ListObjects("excliqliteconfigpurrm")
                  If LO1 Is Nothing Then GoTo EH
                  If LO1.ListColumns("MINERAL").DataBodyRange Is Nothing Then GoTo EH
                  If LO1.ListColumns("ELEMENTO").DataBodyRange Is Nothing Then GoTo EH
                  Let i = LO1.ListRows.Count - 1
                  ReDim VecAux(0 To i, 0 To 1)
                  For i = 0 To UBound(VecAux)
                    Let VecAux(i, 0) = LO1.ListColumns("ELEMENTO").DataBodyRange.Cells(i + 1, 1).Value
                    Let VecAux(i, 1) = LO1.ListColumns("MINERAL").DataBodyRange.Cells(i + 1, 1).Value
                  Next i
                  Set LO1 = Nothing
                  Let .Controls("Cbopur3").ColumnCount = 2
                  Let .Controls("Cbopur3").List = VecAux
                  Let .Controls("Cbopur4").List = Array("DM", "%", "OT", "g/T", "g/L")
                  Let .Controls("Cbopur3").ListIndex = 0
                End If
                Let VecAux = Empty
              End If
              Let .Controls("Txtpur6").Locked = True
              Let .Controls("Txtpur8").Locked = True
              Let .Controls("Txtpur6").Value = VBA.vbNullString
              Let .Controls("Txtpur8").Value = VBA.vbNullString
            End If
          Case "Cbopur3", "Cbopur4"
            If .Controls("Cbopur2").ListIndex = 1 Then
              Let xIndex = IIf(xCombo.Name = "Cbopur3", "Txtpur6", "Txtpur8")
              Select Case xCombo.ListIndex
                Case 0 'Clear once
                  Let yIndex = VBA.CStr(.Controls(xIndex).Value)
                  Let yIndex = VBA.StrReverse(yIndex)
                  Let .Controls(xIndex).Value = VBA.StrReverse(VBA.Mid(yIndex, VBA.InStr(1, yIndex, ";", vbTextCompare) + 1))
                Case 1 'Clear all
                  Let .Controls(xIndex).Value = VBA.vbNullString
                Case Else
                  Let yIndex = VBA.CStr(.Controls(xIndex).Value)
                  Let yIndex = yIndex & IIf(yIndex = VBA.vbNullString, xCombo.Value, ";" & xCombo.Value)
                  Let .Controls(xIndex).Value = yIndex
              End Select
            End If
          
          Case Else: GoTo EH
        End Select
      End With
    
    Case "concentrateinvoice"
'      Call VBA.MsgBox(Frm.Tag)
      With FRM
        Let .Controls("Lstcnc0").MultiSelect = fmMultiSelectSingle
        Let .Controls("Lstcnc0").ListStyle = fmListStylePlain
        Let .Controls("Lstcnc0").Height = .Controls("Frame4").Height * 0.45 '70.25
        Let .Controls("Lstcnc0").Width = (.Controls("Frame4").Width * 0.95)
        Let .Controls("Lstcnc0").Visible = False
        Let .Controls("Chkcnc1").Visible = False
        Let xCombo.ControlTipText = VBA.vbNullString
        Select Case xCombo.Name
          Case "Cbocnc1"
            For i = 12 To 15
              Let .Controls("Txtcnc" & i).Visible = False
              Let .Controls("Txtcnc" & i).Value = VBA.vbNullString
            Next i
            For i = 2 To 7
              Let .Controls("Btncnc" & i).Enabled = False
            Next i
            Let .Controls("Cbocnc2").Visible = False
            Let .Controls("Cbocnc3").Visible = False
            Let .Controls("Lstcnc0").RowSource = VBA.vbNullString
            Select Case xCombo.ListIndex
              Case 0 'Contents
                Set LO1 = excliqlitedatasheet.ListObjects("excliqliteconfigpurrm")
                If LO1.DataBodyRange Is Nothing Then GoTo EH
                Let .Controls("Cbocnc2").RowSource = LO1 & "[MINERAL]"
                If .Controls("Cbocnc2").ListCount > 0 Then Let .Controls("Cbocnc2").ListIndex = 0
                Set LO1 = Nothing
                
                Let .Controls("Cbocnc2").Width = ((.Controls("Frame4").Width * 0.95) / 5)
                Let .Controls("Cbocnc2").Left = .Controls("Cbocnc1").Left
                Let .Controls("Cbocnc2").Visible = True
                Let VecAux = Array("Txtcnc12", "Txtcnc13", "Cbocnc3", "Txtcnc14")
                For i = 0 To 3
                  Let .Controls(VecAux(i)).Visible = True
                  Let .Controls(VecAux(i)).Width = .Controls("Cbocnc2").Width
                  Let .Controls(VecAux(i)).Left = (.Controls("Cbocnc2").Left + (.Controls("Cbocnc2").Width * (i + 1)))
                Next i
                Let VecAux = Empty
                Let StrAux = "excliqliteconpurcontents"
                Let BooAux = (excliqlitedatasheet.ListObjects(StrAux).ListRows.Count <= 10)
                Let BooAux1 = (excliqlitedatasheet.ListObjects(StrAux).ListRows.Count >= 1)
              
              Case 1 'Refination
                Set LO1 = excliqlitedatasheet.ListObjects("excliqliteconpurcontents")
                If LO1.DataBodyRange Is Nothing Then GoTo EH
                Let .Controls("Cbocnc2").RowSource = VBA.vbNullString
                For i = 0 To LO1.ListRows.Count - 1
                  Call .Controls("Cbocnc2").AddItem("Refinación " & LO1.ListColumns("CONTENIDOS").DataBodyRange.Cells(i + 1, 1).Value, i)
                Next i
                Let .Controls("Cbocnc2").ListIndex = 0
                Set LO1 = Nothing
                
                Let .Controls("Cbocnc2").Width = ((.Controls("Frame4").Width * 0.95) / 3)
                Let .Controls("Cbocnc2").Left = .Controls("Cbocnc1").Left
                Let .Controls("Cbocnc2").Visible = True
                'For i = 13 To 14
                For i = 12 To 13
                  Let .Controls("Txtcnc" & i).Visible = True
                  Let .Controls("Txtcnc" & i).Width = .Controls("Cbocnc2").Width
                  'Let .Controls("Txtcnc" & i).Left = (.Controls("Cbocnc2").Left + (.Controls("Cbocnc2").Width * (i - 12)))
                  Let .Controls("Txtcnc" & i).Left = (.Controls("Cbocnc2").Left + (.Controls("Cbocnc2").Width * (i - 11)))
                Next i
                Let StrAux = "excliqliteconpurrefination"
                Let BooAux = (excliqlitedatasheet.ListObjects(StrAux).ListRows.Count <= 10)
                Let BooAux1 = (excliqlitedatasheet.ListObjects(StrAux).ListRows.Count >= 1)
              
              Case 2 'Penalties (Items are added from config)
                Set LO1 = excliqlitedatasheet.ListObjects("excliqliteconpurpenalties")
                If LO1.DataBodyRange Is Nothing Then GoTo EH
                Let .Controls("Cbocnc2").RowSource = LO1 & "[ELEMENTO]"
                If .Controls("Cbocnc2").ListCount > 0 Then Let .Controls("Cbocnc2").ListIndex = 0
                Set LO1 = Nothing
                
                Let .Controls("Cbocnc2").Width = ((.Controls("Frame4").Width * 0.95) / 5)
                Let .Controls("Cbocnc2").Left = .Controls("Cbocnc1").Left
                'Let .Controls("Cbocnc2").Visible = True 'FIXME: SHOW WHEN DBLCLICK IN LIST0 *
                For i = 12 To 15
                  'Let .Controls("Txtcnc" & i).Visible = True '*
                  Let .Controls("Txtcnc" & i).Width = .Controls("Cbocnc2").Width
                  Let .Controls("Txtcnc" & i).Left = (.Controls("Cbocnc2").Left + (.Controls("Cbocnc2").Width * (i - 11)))
                Next i
                Let .Controls("Lstcnc0").Tag = "noevent"
                Let .Controls("Lstcnc0").MultiSelect = fmMultiSelectMulti
                Let .Controls("Lstcnc0").ListStyle = fmListStyleOption
                Let .Controls("Lstcnc0").Tag = VBA.vbNullString
                Let StrAux = "excliqliteconpurpenalties"
                Let BooAux = True ' (excliqlitedatasheet.ListObjects(StrAux).ListRows.Count <= 15)
                Let BooAux1 = False '(excliqlitedatasheet.ListObjects(StrAux).ListRows.Count >= 1)
                Let bIndex = True
                
              Case 3 'Deductions
                Let .Controls("Cbocnc2").RowSource = VBA.vbNullString
                'Let .Controls("Cbocnc2").List = Array("Descuentos por Exportación", "Retenciones de Ley", "Retenciones Institucionales", "Rollback", "Descuentos Otros")
                Let .Controls("Cbocnc2").List = Array("Descuentos por Exportación", "Retenciones de Ley", "Retenciones Institucionales", "Descuentos Otros")
                Let .Controls("Cbocnc2").ListIndex = 0
                GoTo EH
              
              Case 4 'Bonus
                'Let .Controls("Txtcnc12").Width = ((.Controls("Frame4").Width * 0.95) / 3)
                Let .Controls("Txtcnc12").Width = ((.Controls("Frame4").Width * 0.95) / 2)
                Let .Controls("Txtcnc12").Left = .Controls("Cbocnc1").Left
                Let .Controls("Txtcnc12").Visible = True
                'For i = 13 To 14
                For i = 13 To 13
                  Let .Controls("Txtcnc" & i).Visible = True
                  Let .Controls("Txtcnc" & i).Width = .Controls("Txtcnc12").Width
                  'Let .Controls("Txtcnc" & i).Left = (.Controls("Txtcnc12").Left + (.Controls("Txtcnc12").Width * (i - 12)))
                  Let .Controls("Txtcnc" & i).Left = (.Controls("Txtcnc12").Left + (.Controls("Txtcnc12").Width * (i - 12)))
                Next i
                Let StrAux = "excliqliteconpurbonus"
                Let BooAux = (excliqlitedatasheet.ListObjects(StrAux).ListRows.Count <= 10)
                Let BooAux1 = (excliqlitedatasheet.ListObjects(StrAux).ListRows.Count >= 1)
              
              Case Else: GoTo EH
            End Select
            
            Let j = excliqlitedatasheet.ListObjects(StrAux).ListColumns.Count + IIf(xCombo.ListIndex = 2, -1, 0)
            Let .Controls("Lstcnc0").Tag = "noevent"
            Let .Controls("Lstcnc0").RowSource = excliqlitedatasheet.ListObjects(StrAux)
            Let .Controls("Lstcnc0").ColumnCount = j
            Let .Controls("Lstcnc0").ColumnWidths = .Controls("Lstcnc0").Width / j
            
            If xCombo.ListIndex = 2 Then 'Penalties
              With .Controls("Lstcnc0")
                For i = 0 To .ListCount - 1
                  Let .Selected(i) = VBA.CBool(.List(i, 5))
                Next i
                Let .Tag = VBA.vbNullString
              End With
            End If
            
            Let .Controls("Btncnc2").Enabled = BooAux And bIndex = False
            Let .Controls("Btncnc5").Enabled = (BooAux1 Or BooAux = False)
            Let .Controls("Btncnc6").Enabled = (BooAux1 Or BooAux = False)
            Let .Controls("Btncnc7").Enabled = BooAux And bIndex = False
            
            Let bIndex = False

          Case "Cbocnc2"
            If .Controls("Cbocnc1").ListIndex = 3 Then
              For i = 12 To 14
                Let .Controls("Txtcnc" & i).Visible = True
                Let .Controls("Txtcnc" & i).Value = VBA.vbNullString
              Next i
              Let .Controls("Chkcnc1").Visible = False
              'Let i = xCombo.ListIndex
              Let k = xCombo.ListIndex
              Select Case xCombo.ListIndex
                Case 0 'Exports
                  Let .Controls("Cbocnc2").Width = ((.Controls("Frame4").Width * 0.95) / 5)
                  Let .Controls("Cbocnc2").Left = .Controls("Cbocnc1").Left
                  Let .Controls("Cbocnc2").Visible = True
                  For i = 12 To 14
                    Let .Controls("Txtcnc" & i).Visible = False
                    Let .Controls("Txtcnc" & i).Width = .Controls("Cbocnc2").Width
                    Let .Controls("Txtcnc" & i).Left = (.Controls("Cbocnc2").Left + (.Controls("Cbocnc2").Width * (i - 11)))
                  Next i
                  'Let .Controls("Chkcnc1").Visible = True
                  Let .Controls("Chkcnc1").Width = .Controls("Cbocnc2").Width
                  Let .Controls("Chkcnc1").Left = (.Controls("Txtcnc14").Left + .Controls("Txtcnc14").Width) + 2
                  Let StrAux = "excliqliteconpurdedexp" 'EXP
                  'Let BooAux = (excliqlitedatasheet.ListObjects(StrAux).ListRows.Count <= 10)
                  Let BooAux1 = False '(excliqlitedatasheet.ListObjects(StrAux).ListRows.Count >= 1)
                Case 1 'LawRets
                  For j = 12 To 15
                    Let .Controls("Txtcnc" & j).Visible = False
                  Next j
                  Let .Controls("Cbocnc2").Visible = True
                  Let StrAux = "excliqliteconfigpurretentions" 'RETLAW
                  Let BooAux1 = False
                Case 2 'OtherRets
                  For j = 12 To 15
                    Let .Controls("Txtcnc" & j).Visible = False
                  Next j
                  Let .Controls("Cbocnc2").Visible = True
                  Let StrAux = "excliqliteconfigpurretentionsother" 'RETINS
                  Let BooAux1 = False
'                Case 3 'XXXXXXXXXXXXXXXXXXXXXX
'                  Let .Controls("Cbocnc2").Width = ((.Controls("Frame4").Width * 0.95) / 4)
'                  Let .Controls("Cbocnc2").Left = .Controls("Cbocnc1").Left
'                  Let .Controls("Cbocnc2").Visible = True
'                  For i = 12 To 14
'                    Let .Controls("Txtcnc" & i).Visible = True
'                    Let .Controls("Txtcnc" & i).Width = .Controls("Cbocnc2").Width
'                    Let .Controls("Txtcnc" & i).Left = (.Controls("Cbocnc2").Left + (.Controls("Cbocnc2").Width * (i - 11)))
'                  Next i
'                  Let StrAux = "excliqliteconpurdedrollback" 'ROLL
'                  Let BooAux = (excliqlitedatasheet.ListObjects(StrAux).ListRows.Count <= 5)
'                  Let BooAux1 = (excliqlitedatasheet.ListObjects(StrAux).ListRows.Count >= 1)
                Case 3 'Discounts others
                  For j = 12 To 15
                    Let .Controls("Txtcnc" & j).Visible = False
                  Next j
                  Let .Controls("Cbocnc2").Width = ((.Controls("Frame4").Width * 0.95) / 3)
                  Let .Controls("Cbocnc2").Left = .Controls("Cbocnc1").Left
                  Let .Controls("Cbocnc2").Visible = True
                  For i = 12 To 13
                    Let .Controls("Txtcnc" & i).Visible = True
                    Let .Controls("Txtcnc" & i).Width = .Controls("Cbocnc2").Width
                    Let .Controls("Txtcnc" & i).Left = (.Controls("Cbocnc2").Left + (.Controls("Cbocnc2").Width * (i - 11)))
                  Next i
                  Let StrAux = "excliqliteconpurdeddiscounts" 'OTHERS
                  Let BooAux = (excliqlitedatasheet.ListObjects(StrAux).ListRows.Count <= 10)
                  Let BooAux1 = (excliqlitedatasheet.ListObjects(StrAux).ListRows.Count >= 1)
                Case Else: GoTo EH
              End Select
              
              Let BooAux1 = (BooAux1 Or BooAux = False)
              Let .Controls("Lstcnc0").Tag = "noevent"
              Let .Controls("Lstcnc0").RowSource = VBA.vbNullString
              Let j = excliqlitedatasheet.ListObjects(StrAux).ListColumns.Count
              'Let j = IIf(i = 1 Or i = 2, 2, j) + IIf(i = 0, -1, 0)
              Let j = IIf(k = 1 Or k = 2, 2, j) + IIf(k = 0, -1, 0)
              'If (i = 1 Or i = 2) Then
              If (k <= 2) Then ' Or k = 1 Or k = 2) Then
                Let .Controls("Lstcnc0").MultiSelect = fmMultiSelectMulti
                Let .Controls("Lstcnc0").ListStyle = fmListStyleOption
                Let BooAux1 = False
                Let BooAux = False
              End If
              Let .Controls("Lstcnc0").RowSource = excliqlitedatasheet.ListObjects(StrAux)
              Let .Controls("Lstcnc0").ColumnCount = j
              Let .Controls("Lstcnc0").ColumnWidths = .Controls("Lstcnc0").Width / j
              
              If xCombo.ListIndex <= 2 Then ' Or xCombo.ListIndex = 2 Then 'LawRets or OtherRets
                Let j = IIf(xCombo.ListIndex = 0, 4, 2)
                With .Controls("Lstcnc0")
                  For i = 0 To .ListCount - 1
                    Let .Selected(i) = VBA.CBool(.List(i, j))
                  Next i
                  Let .Tag = VBA.vbNullString
                End With
              End If
              
              Let .Controls("Btncnc2").Enabled = BooAux
              Let .Controls("Btncnc5").Enabled = BooAux1
              Let .Controls("Btncnc6").Enabled = BooAux1
              Let .Controls("Btncnc7").Enabled = BooAux
              Let xCombo.ControlTipText = xCombo.Value
            End If

          Case Else: GoTo EH
        End Select
        Let .Controls("Lstcnc0").Visible = True
        Call .Repaint
      End With




    Case "configlauncher"
      Select Case xCombo.Name
        Case "Cbocon7"
          If xCombo.Tag = "noevent" Or xCombo.ListIndex < 0 Or xCombo.ListCount <= 0 Then GoTo EH
          Let xCombo.Tag = xCombo.ListIndex
        
        Case "Cbocon8"
          With FRM
            For i = 8 To 12
              With .Controls("Txtcon" & i)
                Let .Enabled = True
                Let .Visible = False
                Let .Value = VBA.vbNullString
              End With
            Next i
            For i = 11 To 14
              Let .Controls("Btncon" & i).Enabled = False
            Next i
            Let .Controls("Cbocon9").Visible = False
            Let .Controls("Cbocon10").Visible = False
            Let .Controls("Cbocon9").Enabled = True
            Let .Controls("Cbocon10").Enabled = True
            Let .Controls("Chkcon4").Visible = False
            Let l = .Controls("Cbocon8").Left
            Let i = .Controls("Cbocon8").Width
            Select Case xCombo.ListIndex
              Case 0 'RM
                'Show input controls
                Let i = i / 7
                Let VecAux = Array("Txtcon8", "Txtcon9", "Cbocon9", "Txtcon10", "Txtcon11", "Txtcon12", "Cbocon10")
                For j = 0 To UBound(VecAux)
                  With .Controls(VecAux(j))
                    Let .Width = i
                    Let .Left = l + (i * j) - 1.1
                    Let .Visible = True
                  End With
                Next j
                Let VecAux = Empty
                'Load listbox control
                Let j = 7
                With .Controls("Lstcon1")
                  Let .Height = FRM.Controls("Multi").Height * 0.5
                  Let .Tag = "noevent"
                  Let .RowSource = VBA.vbNullString
                  Let .RowSource = excliqlitedatasheet.ListObjects("excliqliteconfigpurrm")
                  Let .ColumnCount = j
                  Let .ColumnWidths = .Width / j
                  Let .ColumnHeads = True
                  Let .Tag = VBA.vbNullString
                End With
                'Set up CRUD controls
                Let .Controls("Btncon11").Enabled = True
                Let .Controls("Btncon14").Enabled = (.Controls("Lstcon1").ListCount > 11) '11 initial items in db
                'Focus on first Txt
                Call .Controls("Txtcon8").SetFocus
                
              Case 1, 2 'Retentions, Retentions others
                'Show input controls
                Let VecAux = Array("Txtcon8", "Txtcon9")
                Let i = i / 2
                For j = 0 To UBound(VecAux)
                  With .Controls(VecAux(j))
                    Let .Width = i
                    Let .Left = l + (i * j) - 1.1
                    Let .Visible = True
                  End With
                Next j
                Let VecAux = Empty
                'Load listbox control
                Let StrAux = IIf(xCombo.ListIndex = 1, "excliqliteconfigpurretentions", "excliqliteconfigpurretentionsother")
                Let j = 2
                Let k = IIf(xCombo.ListIndex = 1, 2, 4)
                With .Controls("Lstcon1")
                  Let .Height = FRM.Controls("Multi").Height * 0.5
                  Let .Tag = "noevent"
                  Let .RowSource = VBA.vbNullString
                  Let .RowSource = excliqlitedatasheet.ListObjects(StrAux)
                  Let .ColumnCount = j
                  Let .ColumnWidths = .Width / j
                  Let .ColumnHeads = True
                  Let .Tag = VBA.vbNullString
                End With
                'Set up CRUD controls
                Let .Controls("Btncon11").Enabled = True
                Let .Controls("Btncon14").Enabled = (.Controls("Lstcon1").ListCount > k) '2 initial items in db
                'Focus on first Txt
                Call .Controls("Txtcon8").SetFocus

              Case 3 'Penalties
                'Show input controls
                Let VecAux = Array("Txtcon8", "Txtcon9", "Txtcon10", "Txtcon11", "Txtcon12")
                Let i = i / 5
                For j = 0 To UBound(VecAux)
                  With .Controls(VecAux(j))
                    Let .Width = i
                    Let .Left = l + (i * j) - 1.1
                    Let .Visible = True
                  End With
                Next j
                Let VecAux = Empty
                'Load listbox control
                Let StrAux = "excliqliteconpurpenalties"
                Let j = 5
                Let k = 5
                With .Controls("Lstcon1")
                  Let .Height = FRM.Controls("Multi").Height * 0.5
                  Let .Tag = "noevent"
                  Let .RowSource = VBA.vbNullString
                  Let .RowSource = excliqlitedatasheet.ListObjects(StrAux)
                  Let .ColumnCount = j
                  Let .ColumnWidths = .Width / j
                  Let .ColumnHeads = True
                  Let .Tag = VBA.vbNullString
                End With
                'Set up CRUD controls
                Let .Controls("Btncon11").Enabled = True
                Let .Controls("Btncon14").Enabled = (.Controls("Lstcon1").ListCount > k) '5 initial items in db
                'Focus on first Txt
                Call .Controls("Txtcon8").SetFocus

              Case 4 'Deductions in exports
                'Show input controls
                Let VecAux = Array("Txtcon8", "Txtcon9", "Txtcon10", "Chkcon4")
                Let i = i / 4
                For j = 0 To UBound(VecAux)
                  With .Controls(VecAux(j))
                    Let .Width = i
                    Let .Left = l + (i * j) - 1.1
                    Let .Visible = True
                  End With
                Next j
                Let .Controls("Chkcon4").Left = .Controls("Chkcon4").Left * 1.02
                Let .Controls("Chkcon4").Top = .Controls("Txtcon10").Top
                Let .Controls("Chkcon4").Value = False
                Let VecAux = Empty
                'Load listbox control
                Let StrAux = "excliqliteconpurdedexp"
                Let j = 4
                Let k = 1
                With .Controls("Lstcon1")
                  Let .Height = FRM.Controls("Multi").Height * 0.5
                  Let .Tag = "noevent"
                  Let .RowSource = VBA.vbNullString
                  Let .RowSource = excliqlitedatasheet.ListObjects(StrAux)
                  Let .ColumnCount = j
                  Let .ColumnWidths = .Width / j
                  Let .ColumnHeads = True
                  Let .Tag = VBA.vbNullString
                End With
                'Set up CRUD controls
                Let .Controls("Btncon11").Enabled = True
                Let .Controls("Btncon14").Enabled = (.Controls("Lstcon1").ListCount > k) '1 initial item in db
                'Focus on first Txt
                Call .Controls("Txtcon8").SetFocus

            End Select
          End With
      End Select
      
    Case Else: GoTo EH
  End Select

EH:
  Erase Controlis
  Erase aControlis
  Set Row0 = Nothing
  Set Row1 = Nothing
  Set Row2 = Nothing
  Set LO1 = Nothing
  Set LO2 = Nothing
  Set xControl = Nothing
  Set REGEX = Nothing
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerDisplay(sFN)
End Sub

'Events handlers - Listboxes
Private Sub View_Controls_listboxes_change(ByVal xList As MSForms.ListBox, ByVal FRM As FRMBasics)
  Dim RES As AppResources
  Dim REGEX As AppResRegEx
  
  Dim LO As ListObject, LC As ListColumn
  
  Dim VecAux As Variant
  Dim StrMsg As String, StrAux0 As String, StrAux1 As String, StrAux2 As String
  Dim i As Long, j As Long, k As Long, l As Long, m As Long, n As Long, o As Long
  Dim BooAux0 As Boolean, BooAux1 As Boolean
  
  Const sFN As String = "VIEW::-Controls_listboxes_change"

  On Error GoTo EH
  #If Debugging Then
    Call Debug_assert(FRM Is Nothing, sFN)
  #End If
  If FRM Is Nothing Then GoTo EH
    
  With FRM
    Select Case .Tag
      Case "oreinvoice"
        Select Case xList.Name
          Case "Lstpur0": Call View_assets_controls_listbox_mannagement(FRM, xList.Name)
          
          Case Else: GoTo EH
        End Select

      Case "concentrateinvoice"
        Select Case xList.Name
          Case "Lstcnc0"
            If .Controls("Lstcnc0").Tag = "noevent" Then GoTo EH
            Call View_assets_controls_listbox_mannagement(FRM, xList.Name)
          
          Case Else: GoTo EH
        End Select
      
      Case Else: GoTo EH
    End Select
  End With

EH:
  Set LC = Nothing
  Set LO = Nothing
  Set REGEX = Nothing
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerDisplay(sFN)
End Sub

Private Sub View_Controls_listboxes_click(ByVal xList As MSForms.ListBox, ByVal FRM As FRMBasics)
  Dim RES As AppResources 'FIXME: MÉTODO DESECHABLE, PODRÍA SER ELIMINADO
  Dim REGEX As AppResRegEx
  
  Dim LO As ListObject, LC As ListColumn
  
  Dim VecAux As Variant
  Dim StrMsg As String, StrAux0 As String, StrAux1 As String, StrAux2 As String
  Dim i As Long, j As Long, k As Long, l As Long, m As Long, n As Long, o As Long
  Dim BooAux0 As Boolean, BooAux1 As Boolean
  
  Const sFN As String = "VIEW::-controls_listboxes_click"

  On Error GoTo EH
  #If Debugging Then
    Call Debug_assert(FRM Is Nothing, sFN)
  #End If
  If FRM Is Nothing Then GoTo EH
    
  With FRM
    Select Case .Tag
      Case "oreinvoice"
  
      Case "configlauncher"
        Select Case xList.Name
          'DIVISA
          Case "Lstcon0"
            If Not (xList.ListCount > 0) Then GoTo EH
            Let .Controls("Btncon3").Enabled = (xList.ListIndex > 1)
            
          'LIQUIDACIONES
          Case "Lstcon1"
            If Not (xList.ListCount > 0) Then GoTo EH
            Let i = .Controls("Lstcon1").ListIndex + 1
            Select Case .Controls("Cbocon8").ListIndex
              Case 0: Let i = 11 'RM
              Case 1: Let i = 2 'Ret
              Case 2: Let i = 4 'RetOth
              Case 3: Let i = 5 'Penalties
              Case 4: Let i = 0 'Deductions in Exports
              Case Else: GoTo EH
            End Select
            Let .Controls("Btncon14").Enabled = (xList.ListIndex >= i)
          
          Case Else: GoTo EH
        End Select
  
      Case Else: GoTo EH
    End Select
  End With

EH:
  Set LC = Nothing
  Set LO = Nothing
  Set REGEX = Nothing
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerDisplay(sFN)
End Sub

Private Sub View_Controls_listboxes_dblclick(ByVal xList As MSForms.ListBox, ByVal FRM As FRMBasics)
  Dim REGEX As AppResRegEx
  
  Dim VecAux As Variant
  Dim StrMsg As String, StrAux0 As String, StrAux1 As String, StrAux2 As String
  Dim i As Long, j As Long, k As Long, l As Long, m As Long, n As Long, o As Long
  Dim BooAux0 As Boolean, BooAux1 As Boolean
  
  Const sFN As String = "VIEW::-Controls_listboxes_dblclick"

  On Error GoTo EH
  #If Debugging Then
    Call Debug_assert(FRM Is Nothing, sFN)
  #End If
  If FRM Is Nothing Then GoTo EH
    
  With FRM
    Set REGEX = New AppResRegEx
    Select Case .Tag
      Case "oreinvoice"
        If Not (xList.ListCount > 0 And xList.ListIndex > -1) Then GoTo EH
        If REGEX.isEmptyStringReg(VBA.CStr(xList.List(xList.ListIndex, 0))) Then GoTo EH
        Let l = .Controls("Cbopur1").ListIndex
        Select Case l
          Case 0 'Base Data
            For i = 5 To 9
              Let .Controls("Txtpur" & i).Value = xList.List(xList.ListIndex, i - 5)
            Next i
            Let .Controls("Txtpur6").Value = VBA.Format$(VBA.CDate(.Controls("Txtpur6").Value), "dd/mm/yy")
          
          Case 1 'Main Data
            Let StrAux0 = VBA.CStr(xList.List(xList.ListIndex, 0))
            Let BooAux0 = REGEX.isMineralComplex(StrAux0)
            If BooAux0 Then 'Mineral type
              Let .Controls("Cbopur2").ListIndex = 1
            ElseIf REGEX.isMineralBrosa(StrAux0) Then
              Let .Controls("Cbopur2").ListIndex = 2
            Else
              Let .Controls("Cbopur2").ListIndex = 0
            End If
            If BooAux0 Then 'Case complex
              For i = 6 To 9
                Let .Controls("Txtpur" & i).Value = xList.List(xList.ListIndex, i - 5)
              Next i
            Else 'Minerals or brosas
              For i = 0 To .Controls("Cbopur3").ListCount - 1
                If VBA.LCase(VBA.CStr(xList.List(xList.ListIndex, 1))) = VBA.LCase(VBA.CStr(.Controls("Cbopur3").List(i, 0))) Then Exit For
              Next i
              Let .Controls("Cbopur3").ListIndex = i
              Let .Controls("Txtpur7").Value = VBA.CStr(xList.List(xList.ListIndex, 2))
              For i = 0 To .Controls("Cbopur4").ListCount - 1
                If VBA.LCase(VBA.CStr(xList.List(xList.ListIndex, 3))) = VBA.LCase(VBA.CStr(.Controls("Cbopur4").List(i, 0))) Then Exit For
              Next i
              Let .Controls("Cbopur4").ListIndex = i
              Let .Controls("Txtpur9").Value = VBA.CStr(xList.List(xList.ListIndex, 4))
            End If
            
          Case 2, 5 'Bonus, Discounts
            Let .Controls("Txtpur5").Value = VBA.CStr(xList.List(xList.ListIndex, 0))
            Let .Controls("Txtpur6").Value = VBA.CStr(xList.List(xList.ListIndex, 1))
          
          Case Else: GoTo EH
        End Select
        For i = 3 To 8
          Let .Controls("Btnpur" & i).Enabled = IIf(i > 3 And i < 6, True, False)
        Next i
        Let .Controls("Cbopur1").Enabled = False
        Let .Controls("Lstpur0").Enabled = False
        Let .Controls("Btnmain1").Enabled = False
  
      
      
      
      Case "concentrateinvoice"
        If Not (xList.ListCount > 0 And xList.ListIndex > -1) Then GoTo EH
        If REGEX.isEmptyStringReg(VBA.CStr(xList.List(xList.ListIndex, 0))) Then GoTo EH
        Let l = .Controls("Cbocnc1").ListIndex
        Select Case l
          Case 0 'Contents
            Let VecAux = .Controls("Cbocnc2").List
            For i = 0 To UBound(VecAux)
              If VBA.LCase(xList.List(xList.ListIndex, 0)) = VBA.LCase(VecAux(i, 0)) Then Let .Controls("Cbocnc2").ListIndex = i: Exit For
            Next i
            Let VecAux = Empty
            Let StrAux0 = xList.List(xList.ListIndex, 1)
            Let StrAux0 = VBA.Left(StrAux0, VBA.InStr(1, StrAux0, "-", vbTextCompare) - 1)
            Let .Controls("Txtcnc12").Value = StrAux0
            Let .Controls("Txtcnc13").Value = xList.List(xList.ListIndex, 2)
            Let .Controls("Txtcnc14").Value = xList.List(xList.ListIndex, 4)
            Let .Controls("Cbocnc3").ListIndex = IIf(VBA.LCase(xList.List(xList.ListIndex, 4)) = "y", 0, 1)
          
          Case 1 'Refination
            Let VecAux = .Controls("Cbocnc2").List
            For i = 0 To UBound(VecAux)
              If VBA.LCase(xList.List(xList.ListIndex, 0)) = VBA.LCase(VecAux(i, 0)) Then Let .Controls("Cbocnc2").ListIndex = i: Exit For
            Next i
            Let VecAux = Empty
            Let .Controls("Txtcnc12").Value = xList.List(xList.ListIndex, 1)
            Let .Controls("Txtcnc13").Value = xList.List(xList.ListIndex, 2)
          
          Case 2 'Penalties
            Let .Controls("Txtcnc12").Value = xList.List(xList.ListIndex, 1)
            Let .Controls("Txtcnc13").Value = xList.List(xList.ListIndex, 2)
            Let .Controls("Txtcnc14").Value = xList.List(xList.ListIndex, 3)
            Let .Controls("Txtcnc15").Value = xList.List(xList.ListIndex, 4)
            For i = 12 To 15
              Let .Controls("Txtcnc" & i).Visible = True
            Next i
          
          Case 3 'Deductions
            Select Case .Controls("Cbocnc2").ListIndex
              Case 0 'Exports
                Let .Controls("Txtcnc12").Visible = True
                Let .Controls("Txtcnc13").Visible = True
                Let .Controls("Txtcnc14").Visible = True
                Let .Controls("Chkcnc1").Visible = True
                Let .Controls("Txtcnc12").Value = xList.List(xList.ListIndex, 0)
                Let .Controls("Txtcnc13").Value = xList.List(xList.ListIndex, 1)
                Let .Controls("Txtcnc14").Value = xList.List(xList.ListIndex, 2)
                Let .Controls("Chkcnc1").Value = VBA.CBool(xList.List(xList.ListIndex, 3))
              Case 3 'Discounts
                Let .Controls("Txtcnc12").Value = xList.List(xList.ListIndex, 0)
                Let .Controls("Txtcnc13").Value = xList.List(xList.ListIndex, 1)
              Case Else: GoTo EH
            End Select
          Case 4 'Bonus
            Let .Controls("Txtcnc12").Value = xList.List(xList.ListIndex, 0)
            Let .Controls("Txtcnc13").Value = xList.List(xList.ListIndex, 1)
          
          Case Else: GoTo EH
        End Select
        For i = 2 To 7
          Let .Controls("Btncnc" & i).Enabled = (i = 3 Or i = 4)
        Next i
        Let .Controls("Cbocnc1").Enabled = False
        Let .Controls("Lstcnc0").Enabled = False
        Let .Controls("Btnmain1").Enabled = False
        
        
        
      Case "configlauncher"
        Select Case xList.Name
          'DIVISA
          Case "Lstcon0"
            If xList.ListCount > 0 And xList.ListIndex > 1 Then
              Let .Controls("Btncon0").Enabled = False
              Let .Controls("Btncon1").Enabled = True
              Let .Controls("Btncon2").Enabled = True
              Let .Controls("Btncon3").Enabled = False
              Let xList.Enabled = False
              
              Let .Controls("Multi").Pages(1).Enabled = False
              Let .Controls("Multi").Pages(2).Enabled = False
              Let .Controls("Multi").Pages(3).Enabled = False
            
              Let .Controls("Txtcon0").Value = xList.List(xList.ListIndex, 0)
              Let .Controls("Txtcon1").Value = xList.List(xList.ListIndex, 1)
              Let .Controls("Txtcon2").Value = xList.List(xList.ListIndex, 2)
            End If
          
          'LIQUIDACIONES
          Case "Lstcon1"
            If Not .Controls("Lstcon1").ListCount > 0 Then GoTo EH
            If .Controls("Lstcon1").Tag = "noevent" Then GoTo EH
            Let i = .Controls("Lstcon1").ListIndex
            Let .Controls("Txtcon8").Enabled = True
            Let .Controls("Txtcon9").Enabled = True
            Let .Controls("Txtcon10").Enabled = True
            Let .Controls("Txtcon11").Enabled = True
            Let .Controls("Txtcon12").Enabled = True
            Let .Controls("Cbocon9").Enabled = True
            Let .Controls("Cbocon10").Enabled = True
            Select Case .Controls("Cbocon8").ListIndex
              Case 0 'RM
                Let BooAux0 = (.Controls("Lstcon1").ListIndex <= 10)
                If BooAux0 Then 'Catalog values
                  Let .Controls("Txtcon8").Enabled = False
                  Let .Controls("Txtcon9").Enabled = False
                  Let .Controls("Cbocon9").Enabled = False
                  Let .Controls("Cbocon10").Enabled = False
                End If
                Let .Controls("Txtcon8").Value = .Controls("Lstcon1").List(i, 0)
                Let .Controls("Txtcon9").Value = .Controls("Lstcon1").List(i, 1)
                For j = 0 To .Controls("Cbocon9").ListCount - 1
                  If VBA.LCase(.Controls("Cbocon9").List(j, 0)) = VBA.LCase(.Controls("Lstcon1").List(i, 2)) Then Exit For
                Next j
                Let .Controls("Cbocon9").ListIndex = j
                Let .Controls("Txtcon10").Value = .Controls("Lstcon1").List(i, 3)
                Let .Controls("Txtcon11").Value = .Controls("Lstcon1").List(i, 4)
                Let .Controls("Txtcon12").Value = .Controls("Lstcon1").List(i, 5)
                For j = 0 To .Controls("Cbocon10").ListCount - 1
                  If VBA.LCase(.Controls("Cbocon10").List(j, 0)) = VBA.LCase(.Controls("Lstcon1").List(i, 6)) Then Exit For
                Next j
                Let .Controls("Cbocon10").ListIndex = j
                With .Controls("Txtcon10")
                  Let .SelStart = 0
                  Let .SelLength = VBA.Len(.Value)
                  Call .SetFocus
                End With
                
              Case 1 'Ret
                Let BooAux0 = (.Controls("Lstcon1").ListIndex <= 2)
                If BooAux0 Then 'Catalog values
                  Let .Controls("Txtcon8").Enabled = False
                End If
                Let .Controls("Txtcon8").Value = .Controls("Lstcon1").List(i, 0)
                Let .Controls("Txtcon9").Value = .Controls("Lstcon1").List(i, 1)
                With .Controls("Txtcon9")
                  Let .SelStart = 0
                  Let .SelLength = VBA.Len(.Value)
                  Call .SetFocus
                End With
              
              Case 2 'RetOth
                Let BooAux0 = (.Controls("Lstcon1").ListIndex <= 4)
                If BooAux0 Then 'Catalog values
                  Let .Controls("Txtcon8").Enabled = False
                End If
                Let .Controls("Txtcon8").Value = .Controls("Lstcon1").List(i, 0)
                Let .Controls("Txtcon9").Value = .Controls("Lstcon1").List(i, 1)
                With .Controls("Txtcon9")
                  Let .SelStart = 0
                  Let .SelLength = VBA.Len(.Value)
                  Call .SetFocus
                End With
              
              Case 3 'Penalties
                Let BooAux0 = (.Controls("Lstcon1").ListIndex <= 5)
                If BooAux0 Then 'Catalog values
                  Let .Controls("Txtcon8").Enabled = False
                End If
                Let .Controls("Txtcon8").Value = .Controls("Lstcon1").List(i, 0)
                Let .Controls("Txtcon9").Value = .Controls("Lstcon1").List(i, 1)
                Let .Controls("Txtcon10").Value = .Controls("Lstcon1").List(i, 2)
                Let .Controls("Txtcon11").Value = .Controls("Lstcon1").List(i, 3)
                Let .Controls("Txtcon12").Value = .Controls("Lstcon1").List(i, 4)
                With .Controls("Txtcon9")
                  Let .SelStart = 0
                  Let .SelLength = VBA.Len(.Value)
                  Call .SetFocus
                End With

              Case 4 'Deductions in Exports
                Let .Controls("Txtcon8").Value = .Controls("Lstcon1").List(i, 0)
                Let .Controls("Txtcon9").Value = .Controls("Lstcon1").List(i, 1)
                Let .Controls("Txtcon10").Value = .Controls("Lstcon1").List(i, 2)
                Let .Controls("Chkcon4").Value = VBA.CBool(.Controls("Lstcon1").List(i, 3))
                With .Controls("Txtcon8")
                  Let .SelStart = 0
                  Let .SelLength = VBA.Len(.Value)
                  Call .SetFocus
                End With
              
              Case Else: GoTo EH
            End Select
            Let .Controls("Multi").Pages(0).Enabled = False
            Let .Controls("Multi").Pages(1).Enabled = False
            Let .Controls("Multi").Pages(2).Enabled = False
            Let .Controls("Btncon11").Enabled = False
            Let .Controls("Btncon12").Enabled = True
            Let .Controls("Btncon13").Enabled = True
            Let .Controls("Btncon14").Enabled = False
            Let .Controls("Cbocon8").Enabled = False
          
          Case Else: GoTo EH
        End Select
      
      Case Else: GoTo EH
    End Select
    Set REGEX = Nothing
  End With

EH:
  Set REGEX = Nothing
  Call EHGLOBAL.ErrorHandlerDisplay(sFN)
End Sub

'Events handlers - Texts
Private Sub View_Controls_texts_dropbuttonclick(ByVal VTxt As MSForms.TextBox, ByVal FRM As FRMBasics)

  Const sFN As String = "VIEW::-TextEvents_dropbuttonclick"
  
  On Error GoTo EH
  'Call VBA.MsgBox(VTxt.Name)
  Select Case FRM.Tag
    
    Case "regression"
      Select Case VTxt.Name
        Case "Txtreg0": Call View_assets_capture_data_from_ranges(FRM, VTxt, "Seleccione datos numéricos para la variable X:", VBooZeroes:=True)
        Case "Txtreg1": Call View_assets_capture_data_from_ranges(FRM, VTxt, "Seleccione datos numéricos para la variable Y:", VBooZeroes:=True)
        Case "Txtreg2": Call View_assets_capture_data_from_ranges(FRM, VTxt, "Seleccione datos numéricos a Extrapolar:", VBooZeroes:=True)
        Case "Txtreg3": If FRM.Controls("Optreg3").Value Then Call View_assets_capture_data_from_ranges(FRM, VTxt, "Seleccione una celda donde imprimir el resultado:", eqlResRanges, VBooRngAddressToTxt:=True)
      End Select
      
    Case "balmet"
      Select Case VTxt.Name
        Case "Txtbal0": Call View_assets_capture_data_from_ranges(FRM, VTxt, "Seleccione datos numéricos de Alimentación del proceso:", eqlResNumbers, True)
        Case "Txtbal2": Call View_assets_capture_data_from_ranges(FRM, VTxt, "Seleccione datos numéricos de Leyes de minerales del proceso:", eqlResNumbers, True)
        Case "Txtbal5": If FRM.Controls("Chkbal3").Value Then Call View_assets_capture_data_from_ranges(FRM, VTxt, "Seleccione datos numéricos de Leyes de minerales complejos del proceso:", eqlResNumbers, True)
        Case "Txtbal8": If FRM.Controls("Optbal1").Value Then Call View_assets_capture_data_from_ranges(FRM, VTxt, "Seleccione una celda donde imprimir el resultado:", eqlResRanges, VBooRngAddressToTxt:=True)
      End Select
    
    Case "balmeteco"
    Case "projections"
    Case "orepurchase"
    Case "cncpurchase"
    Case Else: GoTo EH
  End Select

EH:
  Call EHGLOBAL.ErrorHandlerDisplay(sFN)
  
End Sub

Private Sub View_Controls_texts_change(ByVal xTxt As MSForms.TextBox, ByVal FRM As FRMBasics)
  Dim RES As AppResources
  Dim REGEX As AppResRegEx
  
  Dim vSplits() As Variant
  Dim StrAux As String
  Dim i As Long, j As Long, k As Long, l As Long
  Dim bRange As Boolean, bIndex As Boolean

  Const xProducts As Long = 5
  Const sFN As String = "VIEW::-TextEvents"

  On Error GoTo EH
  #If Debugging Then
    Call Debug_assert(xTxt Is Nothing, sFN)
  #End If
  If xTxt Is Nothing Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "Se requiere un objeto Control.")
  
  Let bIndex = False
  Select Case FRM.Tag
    Case "balmet"
      Select Case xTxt.Name
        Case "Txtbal1"
          Set RES = New AppResources
          Let vSplits = RES.resArrayGetDataColonSeparatedFromString(GloStrRngB, j, False, True)
          Set RES = Nothing
          Erase vSplits
          
          For i = 1 To xProducts '11 (Get the balmet dimension)
            Let k = (i + 2) * i 'normal, i: Cols
            If j = k Then
              Let l = i
              Let bIndex = True
              Exit For
            End If
          Next i
          
          With FRM 'Show/hide controls
            For i = 0 To 4
              Let .Controls("Chkbal" & i + 1).Visible = True
              Let .Controls("Txtbal" & i + 3).Visible = True
              Let .Controls("Cbobal" & i + 1).Visible = True
            Next i
            If bIndex Then
              Let j = l '+ 1 '(dynamic products)
              For i = j To 4
                Let .Controls("Chkbal" & i + 1).Visible = False
                Let .Controls("Txtbal" & i + 3).Visible = False
                Let .Controls("Cbobal" & i + 1).Visible = False
              Next i
            End If
          End With
        
        Case Else: GoTo EH
      End Select
    
    Case "balmeteco"
      Select Case xTxt.Name
        Case "Txtbal1" 'Weights
          Set RES = New AppResources
          Let vSplits = RES.resArrayGetDataColonSeparatedFromString(xTxt.Value, j, False, True)
          Set RES = Nothing
          Erase vSplits
          Let StrAux = VBA.vbNullString
          If j = 0 Then GoTo EH
          For i = 0 To j - 1
            Let StrAux = StrAux & IIf(StrAux = VBA.vbNullString, VBA.vbNullString, ";") & "B" & i + 1
          Next i
          Let FRM.Controls("Txtbal5").Value = StrAux
    
        Case "Txtbal2" 'Units
          Set RES = New AppResources
          Let vSplits = RES.resArrayGetDataColonSeparatedFromString(xTxt.Value, j, False, True)
          Set RES = Nothing
          Erase vSplits
          Let StrAux = VBA.vbNullString
          If j = 0 Then GoTo EH
          For i = 0 To j - 1
            Let StrAux = StrAux & IIf(StrAux = VBA.vbNullString, VBA.vbNullString, ";") & "%"
          Next i
          Let FRM.Controls("Txtbal6").Value = StrAux
      
      Case Else: GoTo EH
      End Select
    
    Case Else: GoTo EH
  End Select

EH:
  Erase vSplits
  Set REGEX = Nothing
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerDisplay(sFN)
End Sub

'Events handlers - Spinbuttons
Private Sub View_Controls_spinbuttons_change( _
  ByVal xSpin As MSForms.SpinButton, _
  ByVal FRM As FRMBasics)
  
  Dim REGEX As AppResRegEx
  
  Dim StrCtrlName As String, StrAux As String
  Dim DblValue As Double, DblValue1 As Double
  
  Const sFN As String = "VIEW::-spinbuttons_change"
  
  On Error GoTo EH
  
  With FRM
    Select Case .Tag
      Case "projection"
        Select Case xSpin.Name
          Case "Spnpro0", "Spnpro1": Let StrCtrlName = "Txtpro6"
          Case "Spnpro2": Let StrCtrlName = "Txtpro7"
          Case "Spnpro3", "Spnpro4": Let StrCtrlName = "Txtpro8"
          Case Else: GoTo EH
        End Select
      Case Else: GoTo EH
    End Select
    Set REGEX = New AppResRegEx
    Let .Controls(StrCtrlName).Value = IIf(Not REGEX.isPercentReg(.Controls(StrCtrlName).Value), xSpin.Value, .Controls(StrCtrlName).Value)
    Let StrAux = VBA.Replace(.Controls(StrCtrlName).Value, "%", VBA.vbNullString)
    If Not REGEX.isNumberReg(StrAux) Then GoTo EH
    Let DblValue = VBA.CDbl(StrAux)
    Set REGEX = Nothing
    
    'Work with spin values
    With xSpin
      Select Case .Tag
        Case "A"
          Let DblValue1 = IIf(.Value >= .Max, .Max, IIf(.Value <= .MIN, .MIN, .Value))
          Let DblValue = (DblValue - VBA.Int(DblValue)) + DblValue1
        
        Case "B"
          Let DblValue1 = IIf(.Value >= .Max, .Max, IIf(.Value <= .MIN, .MIN, .Value))
          Let DblValue = VBA.Int(DblValue) + (DblValue1 / 100)
        
        Case Else: GoTo EH
      End Select
    End With
  
    Select Case .Tag
      Case "projection"
        Select Case xSpin.Name
          Case "Spnpro0", "Spnpro1", "Spnpro3", "Spnpro4": Let .Controls(StrCtrlName).Value = VBA.Format(DblValue, "0.00""%""")
          Case "Spnpro2": Let .Controls(StrCtrlName).Value = VBA.Format(DblValue, "0")
          Case Else: GoTo EH
        End Select
      Case Else: GoTo EH
    End Select
  End With
  
EH:
  Set REGEX = Nothing
  Call EHGLOBAL.ErrorHandlerDisplay(sFN)
End Sub






' CORE VIEWS METHODS LIST -----------------------------------------------------------'
'Validations
Private Function View_Core_validations( _
  ByVal VCmdBtn As MSForms.CommandButton, _
  ByVal FRM As FRMBasics, _
  Optional VFBooSelections As Boolean = False, _
  Optional ControlId As String, _
  Optional VRngBox As Range = Nothing) As Boolean
  
  Dim RES As AppResources_current
  Dim REGEX As AppResRegEx
  
  Const sFN As String = "VIEW::-Validations"

  On Error GoTo EH
  Let View_Core_validations = False
  
  Set RES = New AppResources_current
  Set RES.ErrorHandler = EHGLOBAL
  Set REGEX = New AppResRegEx
  
  If VFBooSelections Then  ' Validations for Selections contents
    
    Select Case ControlId
      Case "regressionselection", "regressionselection-check_cell": Let View_Core_validations = View_Core_validations_regression(Nothing, Nothing, RES, REGEX, VFBooSelections, ControlId, VRngBox)
'      Case "regressionselection-getcell": Let View_Core_validations = True ' Go to get a Cellrange
'      Case "regressionselection": Let View_Core_validations = View_Core_validations_regression(VCmdBtn, FRM, VFBooSelections, ControlId)
'
'      Case "balmetselection-getcell": Let View_Core_validations = True 'Go to get a Cellrange
'      Case "balmetselection": Let View_Core_validations = View_Core_validations_balmet(VCmdBtn, FRM, VFBooSelections, ControlId)
      
      Case Else: GoTo EH
    End Select
  
  Else ' Validation to Forms contents
    
    If CMB Is Nothing Then Set CMB = New ClassCustomMsgbox
    If Not VCmdBtn.Tag = "validate" Then GoTo EH 'Just pass enabled for valitations buttons
    
    #If Debugging Then
      Call Debug_assert(((FRM Is Nothing) Or (VCmdBtn Is Nothing)), sFN)
    #End If
        
    Select Case FRM.Tag
      Case "regression": Let View_Core_validations = View_Core_validations_regression(VCmdBtn, FRM, RES, REGEX, VFBooSelections, ControlId)
'      Case "balmet", "balmeteco": Let View_Core_validations = View_Core_validations_balmet(VCmdBtn, FRM, VFBooSelections, ControlId)
'      Case "projection": Let View_Core_validations = View_Core_validations_projections(VCmdBtn, FRM)
'      Case "oreinvoice": Let View_Core_validations = View_Core_validations_oreinvoice(VCmdBtn, FRM)
'      Case "concentrateinvoice": Let View_Core_validations = View_Core_validations_concentrateinvoice(VCmdBtn, FRM)
'      Case "configlauncher": Let View_Core_validations = View_Core_validations_config(VCmdBtn, FRM)
      Case Else: GoTo EH
    End Select
  
  End If
  
EH:
  'Set CMB = Nothing
  Set REGEX = Nothing
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function View_Core_validations_regression( _
  ByVal VCmdBtn As MSForms.CommandButton, _
  ByVal FRM As FRMBasics, _
  ByVal RES As AppResources_current, _
  ByVal REGEX As AppResRegEx, _
  Optional VFBooSelections As Boolean = False, _
  Optional ControlId As String, _
  Optional VRngBox As Range = Nothing) As Boolean
  
  Dim RngBox As Range
  Dim vSplits As Variant
  Dim xMsg As String, StrAux0 As String
  Dim iIndex As Long, i As Long, j As Long
  
  Const sFN As String = "VIEW::-Validations_regression"

  On Error GoTo EH
  Let View_Core_validations_regression = False
  
  If VFBooSelections Then 'Validations for Selections contents
    
    If ControlId = "regressionselection" Then
      
      Set RngBox = Application.Selection
      Let StrAux0 = "Seleccione datos iniciales para la Regresión. " & VBA.vbNewLine & VBA.vbNewLine & _
                   "Dato 1: Variables Y, puede ser un rango o un área." & VBA.vbNewLine & _
                   "Dato 2 (Opcional): Variables X, puede ser un rango o un área." & VBA.vbNewLine & _
                   "Dato 3 (Opcional): Valores a extrapolar, puede ser un rango o un área" & VBA.vbNewLine & VBA.vbNewLine & _
                   "Recomendado, datos en columnas por cada valor."
      If RngBox Is Nothing Then Call VBA.MsgBox(StrAux0, vbInformation, STRAPPTITLE): GoTo EH
      If Application.WorksheetFunction.Count(RngBox) < 3 Then Call VBA.MsgBox(StrAux0, vbInformation, STRAPPTITLE): GoTo EH
  
      If RngBox.Areas.Count = 1 Then
        Let i = Application.WorksheetFunction.Count(RngBox.Columns("A:A"))
        If i < 3 Then Call VBA.MsgBox(StrAux0, vbInformation, STRAPPTITLE): GoTo EH
        Let i = Application.WorksheetFunction.Sum(RngBox.Columns("A:A"))
        If i = 0 Then Call VBA.MsgBox("Los elementos deben ser distintos de cero.", vbInformation, STRAPPTITLE): GoTo EH
        If RngBox.Columns.Count > 1 Then
          Let i = Application.WorksheetFunction.Count(RngBox.Columns("A:A"))
          Let j = Application.WorksheetFunction.Count(RngBox.Columns("B:B"))
          If Not j = i Then Call VBA.MsgBox("La cantidad de elementos de la primera columna debe ser igual a la cantidad de elementos de la segunda columna.", vbInformation, STRAPPTITLE): GoTo EH
          Let j = Application.WorksheetFunction.Sum(RngBox.Columns("B:B"))
          If j = 0 Then Call VBA.MsgBox("Los elementos deben ser distintos de cero.", vbInformation, STRAPPTITLE): GoTo EH
        End If
      ElseIf RngBox.Areas.Count > 1 Then
        Let i = Application.WorksheetFunction.Count(RngBox.Areas(1))
        If i < 3 Then Call VBA.MsgBox(StrAux0, vbInformation, STRAPPTITLE): GoTo EH
        Let i = Application.WorksheetFunction.Sum(RngBox.Areas(1))
        If i = 0 Then Call VBA.MsgBox("Los elementos deben ser distintos de cero.", vbInformation, STRAPPTITLE): GoTo EH
        Let i = Application.WorksheetFunction.Count(RngBox.Areas(1))
        Let j = Application.WorksheetFunction.Count(RngBox.Areas(2))
        If Not j = i Then Call VBA.MsgBox("La cantidad de elementos de la primera columna debe ser igual a la cantidad de elementos de la segunda columna.", vbInformation, STRAPPTITLE): GoTo EH
        Let j = Application.WorksheetFunction.Sum(RngBox.Areas(2))
        If j = 0 Then Call VBA.MsgBox("Los elementos deben ser distintos de cero.", vbInformation, STRAPPTITLE): GoTo EH
      End If
      Set RngBox = Nothing
      Let View_Core_validations_regression = True
    
    Else ' "regressionselection-check_cell"
    
      'Check for spreadsheet edges and data inside the region selected
      If VRngBox Is Nothing Then GoTo EH
      Let iIndex = ((Application.Selection.Rows.Count + 3) * 3) ' A1 to TotalsRows aprox
      If RES.ExcelEdges(VRngBox, xRow:=iIndex, xCol:=20) Then Call VBA.MsgBox("La celda que eligió está muy cerca de los bordes de la hoja de cálculo, este ejercicio requiere más espacio. Seleccione otra celda mpas alejada de los bordes.", STRAPPTITLE): GoTo EH
      If (Application.WorksheetFunction.CountA(VRngBox.Range("A1").Resize(iIndex, 20)) > 0) Then
        If VBA.MsgBox(STRAPPTITLE & " sobreescribirá cualquier dato que se encuentre al rededor de la celda seleccionada. Existen datos que podrían ser eliminados." & VBA.vbNewLine & "¿continuar de todos modos?", vbYesNo + vbQuestion, STRAPPTITLE) = vbNo Then GoTo EH
      End If
      Let View_Core_validations_regression = (VBA.MsgBox(STRAPPTITLE & " calculará la regresión lineal con los datos que seleccionó." & VBA.vbNewLine & "¿Proceder?", vbQuestion + vbYesNo, STRAPPTITLE) = vbYes)
    
    End If

    
    
'    Select Case ControlId
'      Case "regressionselection-getcell": Let View_Core_validations_regression = True 'Go to get a Cellrange
'      Case "regressionselection"
'        Set RngBox = Selection
'        If RngBox Is Nothing Then GoTo EH
'        Let StrAux0 = "Seleccione datos iniciales para la Regresión. " & VBA.vbNewLine & VBA.vbNewLine & _
'                     "Dato 1: Variables Y, puede ser un rango o un área." & VBA.vbNewLine & _
'                     "Dato 2 (Opcional): Variables X, puede ser un rango o un área." & VBA.vbNewLine & _
'                     "Dato 3 (Opcional): Valores a extrapolar, puede ser un rango o un área" & VBA.vbNewLine & VBA.vbNewLine & _
'                     "Recomendado, datos en columnas por cada valor."
'        If Application.WorksheetFunction.Count(RngBox) < 3 Then Call VBA.MsgBox(StrAux0, vbInformation, STRAPPTITLE): GoTo EH
'
'        If RngBox.Areas.Count = 1 Then
'          Let i = Application.WorksheetFunction.Count(RngBox.Columns("A:A"))
'          If i < 3 Then Call VBA.MsgBox(StrAux0, vbInformation, STRAPPTITLE): GoTo EH
'          Let i = Application.WorksheetFunction.Sum(RngBox.Columns("A:A"))
'          If i = 0 Then Call VBA.MsgBox("Los elementos deben ser distintos de cero.", vbInformation, STRAPPTITLE): GoTo EH
'          If RngBox.Columns.Count > 1 Then
'            Let i = Application.WorksheetFunction.Count(RngBox.Columns("A:A"))
'            Let j = Application.WorksheetFunction.Count(RngBox.Columns("B:B"))
'            If Not j = i Then Call VBA.MsgBox("La cantidad de elementos de la primera columna debe ser igual a la cantidad de elementos de la segunda columna.", vbInformation, STRAPPTITLE): GoTo EH
'            Let j = Application.WorksheetFunction.Sum(RngBox.Columns("B:B"))
'            If j = 0 Then Call VBA.MsgBox("Los elementos deben ser distintos de cero.", vbInformation, STRAPPTITLE): GoTo EH
'          End If
'        ElseIf RngBox.Areas.Count > 1 Then
'          Let i = Application.WorksheetFunction.Count(RngBox.Areas(1))
'          If i < 3 Then Call VBA.MsgBox(StrAux0, vbInformation, STRAPPTITLE): GoTo EH
'          Let i = Application.WorksheetFunction.Sum(RngBox.Areas(1))
'          If i = 0 Then Call VBA.MsgBox("Los elementos deben ser distintos de cero.", vbInformation, STRAPPTITLE): GoTo EH
'          Let i = Application.WorksheetFunction.Count(RngBox.Areas(1))
'          Let j = Application.WorksheetFunction.Count(RngBox.Areas(2))
'          If Not j = i Then Call VBA.MsgBox("La cantidad de elementos de la primera columna debe ser igual a la cantidad de elementos de la segunda columna.", vbInformation, STRAPPTITLE): GoTo EH
'          Let j = Application.WorksheetFunction.Sum(RngBox.Areas(2))
'          If j = 0 Then Call VBA.MsgBox("Los elementos deben ser distintos de cero.", vbInformation, STRAPPTITLE): GoTo EH
'        End If
'        'Check for spreadsheet edges and data inside the region selected
'        If GLOBOO_SELECCTIONCELL Then
'          Let iIndex = (i * 5) '5: A1 to totalsrow aprox
'          If GLORNG_CELL Is Nothing Then Call VBA.MsgBox("¡Debe seleccionar una celda para imprimir los resultados!", vbInformation, STRAPPTITLE): GoTo EH
'          If RES.resExcelEdges(GLORNG_CELL, iIndex, 20) Then
'            Call VBA.MsgBox("La celda que eligió está muy cerca de los bordes de la hoja de cálculo, este ejercicio requiere más espacio. Seleccione otra celda mpas alejada de los bordes.", STRAPPTITLE): GoTo EH
'          End If
'          If (Application.WorksheetFunction.CountA(GLORNG_CELL.Range("A1").Resize(iIndex, 20)) > 0) Then
'            If VBA.MsgBox(STRAPPTITLE & " sobreescribirá cualquier dato que se encuentre al rededor de la celda seleccionada. Existen datos que podrían ser eliminados." & VBA.vbNewLine & "¿continuar de todos modos?", vbYesNo + vbQuestion, STRAPPTITLE) = vbNo Then GoTo EH
'          End If
'        End If
'
'        Let xMsg = "Se calculará la regresión lineal con los datos que seleccionó." & VBA.vbNewLine & "¿Proceder?"
'        Let View_Core_validations_regression = (VBA.MsgBox(xMsg, vbQuestion + vbYesNo, STRAPPTITLE) = vbYes)
'        Set RngBox = Nothing
'
'      Case Else: GoTo EH
'    End Select

  Else 'Validation to Forms contents
    
    With FRM
      Select Case VCmdBtn.Name
        Case "Btnmain1"
          ' Y
          Let vSplits = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(.Controls("Txtreg1").Value, i, RBooIncludeZeroes:=True)
          If i <= 0 Then Call CMB.CMsgBox("¡Se esperaban valores base para la regresión en Y!", STRAPPTITLE, "Error", cMsgType:=cmbError, cUFrm:=FRM, cCtrlTargetName:="Txtreg1", cCtrlSourceName:=VCmdBtn.Name, cCtrlRequired:=cmbRequired, cPositionX:=cmbRight, cPositionY:=cmbBottom, cAspect:=cmbTagLine): GoTo EH
          If Application.WorksheetFunction.Sum(vSplits) = 0 Then Call CMB.CMsgBox("¡Se esperaben valores no nulos para la regresión en Y!", STRAPPTITLE, "Error", cMsgType:=cmbError, cUFrm:=FRM, cCtrlTargetName:="Txtreg1", cCtrlSourceName:=VCmdBtn.Name, cCtrlRequired:=cmbRequired, cPositionX:=cmbRight, cPositionY:=cmbBottom, cAspect:=cmbTagLine): GoTo EH
          Let vSplits = Empty
          ' X
          Let vSplits = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(.Controls("Txtreg0").Value, j, RBooIncludeZeroes:=True)
          If j > 0 Then
            If Not j = i Then Call CMB.CMsgBox("¡Se requieren mismas cantidades de valores en X e Y!", STRAPPTITLE, "Error", cMsgType:=cmbError, cUFrm:=FRM, cCtrlTargetName:="Txtreg0", cCtrlSourceName:=VCmdBtn.Name, cCtrlRequired:=cmbRequired, cPositionX:=cmbRight, cPositionY:=cmbBottom, cAspect:=cmbTagLine): GoTo EH
            If Application.WorksheetFunction.Sum(vSplits) = 0 Then Call CMB.CMsgBox("¡Se esperaben valores en X no nulos!", STRAPPTITLE, "Error", cMsgType:=cmbError, cUFrm:=FRM, cCtrlTargetName:="Txtreg0", cCtrlSourceName:=VCmdBtn.Name, cCtrlRequired:=cmbRequired, cPositionX:=cmbRight, cPositionY:=cmbBottom, cAspect:=cmbTagLine): GoTo EH
          End If
          Let vSplits = Empty

          ' Check for spreadsheet edges and data inside the region selected
          If .Controls("Optreg3").Value Then
            Let iIndex = (i * 5) '5: A1 to totalsrow aprox
            If GLORNG_CELL Is Nothing Then Call CMB.CMsgBox("¡Debe seleccionar una celda para imprimir los resultados!", STRAPPTITLE, "Error", cMsgType:=cmbError, cUFrm:=FRM, cCtrlTargetName:="Txtreg3", cCtrlSourceName:=VCmdBtn.Name, cCtrlRequired:=cmbRequired, cAspect:=cmbPopupAutoclose): GoTo EH
            If RES.ExcelEdges(GLORNG_CELL, iIndex, 20) Then
              Call VBA.MsgBox("La celda que eligió está muy cerca de los bordes de la hoja de cálculo, este ejercicio requiere más espacio. Seleccione otra celda mpas alejada de los bordes.", STRAPPTITLE): GoTo EH
            End If
            If (Application.WorksheetFunction.CountA(GLORNG_CELL.Range("A1").Resize(iIndex, 20)) > 0) Then
              If VBA.MsgBox(STRAPPTITLE & " sobreescribirá cualquier dato que se encuentre al rededor de la celda seleccionada. Existen datos que podrían ser eliminados." & VBA.vbNewLine & "¿continuar de todos modos?", vbYesNo + vbQuestion, STRAPPTITLE) = vbNo Then GoTo EH
            End If
          Else
            Set GLORNG_CELL = Nothing
          End If

          Let xMsg = "Se calculará la regresión lineal con los datos que ingresó." & VBA.vbNewLine & "¿Proceder?"
          Let View_Core_validations_regression = (VBA.MsgBox(xMsg, vbQuestion + vbYesNo, STRAPPTITLE) = vbYes)

        Case Else: GoTo EH
      End Select
    End With
    
  End If
  
EH:
  Let vSplits = Empty
  Set RngBox = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function View_Core_validations_balmet( _
  ByVal VCmdBtn As MSForms.CommandButton, _
  ByVal FRM As FRMBasics, _
  Optional VFBooSelections As Boolean = False, _
  Optional ControlId As String) As Boolean
  
  Dim RES As New AppResources
  Dim REGEX As New AppResRegEx
  Dim CMsgBox As New ClassCustomMsgbox
  
  Dim RngBox As Range
  Dim vSplits As Variant
  Dim xMsg As String, StrAux0 As String, StrAux1 As String, StrAux2 As String
  Dim iIndex As Long, i As Long, j As Long, l As Long, m As Long, n As Long, CurIndex As Currency
  Dim BooAux As Boolean
  
  Const sFN As String = "VIEW::-Validations_balmet"

  On Error GoTo EH
'  Let View_Core_validations_balmet = False
'
'  If Not VFBooSelections Then 'Validations to Forms contents
'    With FRM
'      Select Case FRM.Tag
'        Case "balmet"
'          Select Case VCmdBtn.Name 'NEW
'            Case "Btnbal0", "Btnbal1", "Btnbal2", "Btnbal3", "Btnbal5": Let View_Core_validations_balmet = True
'            Case "Btnbal4": Let View_Core_validations_balmet = .Controls("Chkbal6").Value
'            Case "Btnbal6": Let View_Core_validations_balmet = .Controls("Optbal1").Value
'            Case "Btnmain1"
'              '1. [x]: Validate Feed and grades
'              With .Controls("Txtbal0") 'Feed
'                Let StrAux0 = IIf(.Tag = "rng", GloStrRngA, .Value)
'                If REGEX.isEmptyStringReg(StrAux0) Then Call CMSGBOX.CMsgBoxFluid("¡Falta valor de Alimentación!", "Error", MMsgType:="Error", MFrm:=FRM, MTargetControl:=.Name, MRequired:=True, MSourceControl:=VCmdBtn.Name): GoTo EH
'                If Not REGEX.isNumberEntireAndDecimalVectorColonSeparated(StrAux0) Then
'                  Let StrAux1 = "Para la Alimentación debe introducir un número o varios números separados por punto y coma (Ejemplo: 1,3;4;10) o un rango de una o varias celdas mediante el botón correspondiente." & VBA.vbNewLine & "Se tomará en cuenta la sumatoria y el valor absoluto de los datos que ingrese, además la Alimentación debe ser diferente de cero."
'                  Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtbal0"), StrAux1, VBA.vbInformation, True): GoTo EH
'                End If
'                Let vSplits = RES.resArrayGetDataColonSeparatedFromString(StrAux0, i, False, True)
'                If Application.WorksheetFunction.Sum(vSplits) = 0 Or i = 0 Then Call CMSGBOX.CMsgBoxFluid("¡Alimentación debe ser un número mayor a cero!", "Error", MMsgType:="Error", MFrm:=FRM, MTargetControl:=.Name, MRequired:=True, MSourceControl:=VCmdBtn.Name): GoTo EH
'                Let vSplits = Empty
'              End With
'
'              With .Controls("Txtbal1")
'                Let StrAux0 = "Conjunto de leyes inválido." & VBA.vbNewLine & _
'                    "La referencia debe ser un rango de celdas o un grupo de números separados por punto y coma (Ej.: 1;2;3) con las siguientes dimensiones: " & VBA.vbNewLine & VBA.vbNewLine & _
'                    "Para Balance de 2 Productos: 3 filas x 1 columna." & VBA.vbNewLine & _
'                    "Para Balance de 3 Productos: 4 filas x 2 columnas." & VBA.vbNewLine & _
'                    "Para Balance de 4 Productos: 5 filas x 3 columnas." & VBA.vbNewLine & _
'                    "Para Balance de 5 Productos: 6 filas x 4 columnas, etc."
'                Let StrAux1 = IIf(.Tag = "rng", GloStrRngB, .Value)
'                If REGEX.isEmptyStringReg(StrAux1) Then Call CMSGBOX.CMsgBoxFluid("¡Falta valores de leyes de minerales (cabeza y productos)!", "Error", MMsgType:="Error", MFrm:=FRM, MTargetControl:=.Name, MRequired:=True, MSourceControl:=VCmdBtn.Name): GoTo EH
'                If Not REGEX.isNumberEntireAndDecimalVectorColonSeparated(StrAux1) Then Call RES.resFrmCustomMsgboxRequiredControl(FRM.Controls("Txtbal1"), StrAux0, VBA.vbInformation, True): GoTo EH
'                Let vSplits = RES.resArrayGetDataColonSeparatedFromString(StrAux1, j, False, True)
'                Let vSplits = Empty
'                '1. [x]: Just numbers
'                'P: RxC-CX: ('2: 3x1+3=6  '3: 4x2+4=12  '4: 5x3+5=20  '5: 6x4+6=30  '6: 7x5+7=42)
'                'Accept balances with five products
'                If j > 35 Then Call RES.resFrmCustomMsgboxRequiredControl(.Parent.Controls("Txtbal1"), "Puede obtener Balances Metalúrgicos de más de cinco productos (hasta de diez) mediante la opción Balmet Selección de la cinta de opciones " & STRAPPTITLE & ".", VBA.vbInformation): GoTo EH
'                Let BooAux = False
'                For i = 1 To 5 '4 '11
'                  Let n = (i + 2) * i 'normal
'                  Let l = i 'Cols
'                  If FRM.Controls("Chkbal6").Value Then Let FRM.Controls("Chkbal6").Tag = (i + 2) 'complex
'                  If j = n Then
'                    Let FRM.Controls("Lblbal2").Tag = i
'                    Let BooAux = True
'                    Exit For
'                  End If
'                Next i
'                If Not BooAux Then Call RES.resFrmCustomMsgboxRequiredControl(FRM.Controls("Txtbal1"), StrAux0): GoTo EH
'                If FRM.Controls("Cbobal0").ListIndex = 0 And j > 3 Then Call RES.resFrmCustomMsgboxRequiredControl(FRM.Controls("Cbobal0"), "El método Convencional sólo admite Balances Metalúrgicos de dos productos.", VBA.vbInformation): GoTo EH
'              End With
'
'              '2. [x]: Complex grades
'              Let m = 0 '
'              If .Controls("Chkbal6").Value Then
'                If (.Controls("Cbobal6").ListIndex > 0 And .Controls("Cbobal7").ListIndex = 0) Then Call CMSGBOX.CMsgBoxFluid("¿" & .Controls("Cbobal7").Value & " para minerales complejos de " & .Controls("Cbobal6").Value & "?", "Error", MMsgType:="Error", MFrm:=FRM, MTargetControl:="Cbobal7", MRequired:=True, MSourceControl:=VCmdBtn.Name): GoTo EH
'                With .Controls("Txtbal8")
'                  Let i = FRM.Controls("Chkbal6").Tag 'complex length
'                  Let StrAux0 = IIf(.Tag = "rng", GloStrRngC, .Value)
'                  If REGEX.isEmptyStringReg(StrAux0) Then Call CMSGBOX.CMsgBoxFluid("¡Faltan los valores de leyes de minerales (cabeza y productos) para complejos!", "Error", MMsgType:="Error", MFrm:=FRM, MTargetControl:=.Name, MRequired:=True, MSourceControl:=VCmdBtn.Name): GoTo EH
'                  If Not REGEX.isNumberEntireAndDecimalVectorColonSeparated(StrAux0) Then
'                    Let StrAux1 = "Para las leyes de minerales complejos debe introducir varios números separados por punto y coma (Ejemplo: 1,3;4;10) o un rango de una o varias celdas mediante el botón correspondiente." & VBA.vbNewLine & "La cantidad de datos que ingrese, debe ser igual a la cantidad de leyes de los minerales principales del tratamiento."
'                    Call RES.resFrmCustomMsgboxRequiredControl(FRM.Controls("Txtbal8"), StrAux1, VBA.vbInformation, True): GoTo EH
'                  End If
'                  Let vSplits = RES.resArrayGetDataColonSeparatedFromString(StrAux0, j, False, True)
'                  Let vSplits = Empty
'                  'Accept balances at five products
'                  If Not i = j Then
'                    Let m = 1
'                    Call RES.resFrmCustomMsgboxRequiredControl(FRM.Controls("Txtbal8"), "Parámetro de valores para minerales complejos inválida." & VBA.vbNewLine & _
'                      "La referencia debe ser un conjunto de celdas con las siguientes dimensiones: " & VBA.vbNewLine & VBA.vbNewLine & _
'                      "Para Balance de 2 Productos: 3 filas x 1 columna." & VBA.vbNewLine & _
'                      "Para Balance de 3 Productos: 4 filas x 1 columnas." & VBA.vbNewLine & _
'                      "Para Balance de 4 Productos: 5 filas x 1 columnas." & VBA.vbNewLine & _
'                      "Para Balance de 5 Productos: 6 filas x 1 columnas, etc.")
'                    GoTo EH
'                  End If
'                End With
'              End If
'
'              '3. [x]: Validate range place if Opt is active
'              'Check for spreadsheet edges and data inside the region selected
'              If .Controls("Optbal1").Value Then
'                If .Controls("Chkbal7").Value Then 'All results
'                  Let iIndex = (3 + (l * 4) + (m * 5) + 2) 'l: grades cols, m: complex cols, 3: initial fixed cols, 2: %p & Rc
'                Else
'                  Let iIndex = ( _
'                      3 + _
'                      (l * (2 + IIf(.Controls("Chkbal9").Value, 1, 0) + IIf(.Controls("Chkbal10").Value, 1, 0)) + _
'                      (m * (2 + IIf(.Controls("Chkbal9").Value, 1, 0) + IIf(.Controls("Chkbal10").Value, 1, 0) + IIf(.Controls("Chkbal11").Value, 1, 0) + IIf(.Controls("Chkbal12").Value, 1, 0))) + _
'                      (IIf(.Controls("Chkbal8").Value, 1, 0) + IIf(.Controls("Chkbal13").Value, 1, 0)))) + _
'                      5 + (IIf(.Controls("Chkbal14").Value, 1, 0) * 2) 'l: grades cols, m: complex cols, 3: initial fixed cols, 2: %p & Rc
'                End If
'                If GLORNG_CELL Is Nothing Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtbal9"), "¡Debe seleccionar una celda para imprimir los resultados!", VBA.vbInformation): GoTo EH
'                If RES.resExcelEdges(GLORNG_CELL, iIndex, 10) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtbal9"), "La celda que eligió está muy cerca de los bordes de la hoja de cálculo, este ejercicio requiere más espacio. Seleccione otra celda mpas alejada de los bordes.", VBA.vbInformation): GoTo EH
'                If (Application.WorksheetFunction.CountA(GLORNG_CELL.Range("A1").Resize(iIndex, 20)) > 0) Then
'                  If VBA.MsgBox(STRAPPTITLE & " sobreescribirá cualquier dato que se encuentre alrededor de la celda seleccionada. Existen datos que podrían ser eliminados." & VBA.vbNewLine & "¿continuar de todos modos?", vbYesNo + vbQuestion, STRAPPTITLE) = vbNo Then GoTo EH
'                End If
'              End If
'
'              Let xMsg = "Se calculará el Balance Metalúrgico con los datos que ingresó." & VBA.vbNewLine & "¿Proceder?"
'              Let View_Core_validations_balmet = (VBA.MsgBox(xMsg, vbQuestion + vbYesNo, STRAPPTITLE) = vbYes)
'
'            Case Else: GoTo EH
'          End Select
'
'        Case "balmeteco"
'          Select Case VCmdBtn.Name
'            Case "Btnbal0", "Btnbal1", "Btnbal2", "Btnbal3", "Btnbal4", "Btnbal5", "Btnbal6", "Btnbal7", "Btnbal8", "Btnbal9", "Btnbal10", "Btnbal11", "Btnbal12", "Btnbal13", "Btnbal14": Let View_Core_validations_balmet = True
'
'            Case "Btnmain1"
'              Dim StrFeed As String, StrWeights As String, StrGrades As String, StrRecs As String, StrUnits As String, StrNames As String, StrGradesCx As String, StrUnitsCx As String, StrRecsCx As String
'              Dim VecFeed As Variant, VecWeights As Variant, VecGrades As Variant, VecRecs As Variant, VecUnits As Variant, VecNames As Variant, VecGradesCx As Variant, VecUnitsCx As Variant, VecRecsCx As Variant
'
'              Dim DblSum As Double, DblAux As Double
'              Dim x1 As Long, x2 As Long, x3 As Long, x4 As Long, x5 As Long
'              Dim BooBalCx As Boolean
'              Dim StrAux As String
'
'              Let StrFeed = .Controls("Txtbal0").Value
'              Let StrWeights = .Controls("Txtbal1").Value
'              Let StrGrades = .Controls("Txtbal2").Value
'              Let StrRecs = .Controls("Txtbal3").Value
'              Let StrUnits = .Controls("Txtbal6").Value
'              Let StrGradesCx = .Controls("Txtbal8").Value
'              Let StrRecsCx = .Controls("Txtbal9").Value
'              Let StrUnitsCx = .Controls("Cbobal4").Value
'              Let BooBalCx = .Controls("Chkbal4").Value
'
'              'Initialize validations
'              Set REGEX = New AppResRegEx
'              Set RES = New AppResources
'
'                'Feed
'              With .Controls("Txtbal0")
'                If REGEX.isEmptyStringReg(StrFeed) Then Call CMSGBOX.CMsgBoxFluid("¡Falta valor de Alimentación!", "Error", MMsgType:="Error", MFrm:=FRM, MTargetControl:=.Name, MRequired:=True, MSourceControl:=VCmdBtn.Name): GoTo EH   'Feed required
'                Let VecFeed = RES.resArrayGetDataColonSeparatedFromString(StrFeed, i, False, True)
'                If i = 0 Then Call CMSGBOX.CMsgBoxFluid("¡Valor de Alimentación debe ser un número!", "Error", MMsgType:="Error", MFrm:=FRM, MTargetControl:=.Name, MRequired:=True, MSourceControl:=VCmdBtn.Name): GoTo EH 'Feed required
'                Let DblSum = Application.WorksheetFunction.Sum(VecFeed)
'                Let VecFeed = Empty
'                If DblSum = 0 Then Call CMSGBOX.CMsgBoxFluid("¿Cero Toneladas de Alimentación?", "Error", MMsgType:="Error", MFrm:=FRM, MTargetControl:=.Name, MRequired:=True, MSourceControl:=VCmdBtn.Name): GoTo EH 'Feed required
'                If DblSum > CDBL_MAX Then Call CMSGBOX.CMsgBoxFluid("¡Valor de Alimentación muy grande!", "Error", MMsgType:="Error", MFrm:=FRM, MTargetControl:=.Name, MRequired:=True, MSourceControl:=VCmdBtn.Name): GoTo EH 'Feed too big
'              End With
'
'                'Weight(s)
'              With .Controls("Txtbal1")
'                If REGEX.isEmptyStringReg(StrWeights) Then Call CMSGBOX.CMsgBoxFluid("¡Falta valor de Pesos de Productos!", "Error", MMsgType:="Error", MFrm:=FRM, MTargetControl:=.Name, MRequired:=True, MSourceControl:=VCmdBtn.Name): GoTo EH 'Weights required
'                Let VecWeights = RES.resArrayGetDataColonSeparatedFromString(StrWeights, x1, False, True)
'                If x1 = 0 Or x1 > BAL_LIMIT Then Call CMSGBOX.CMsgBoxFluid("Puede calcular Balances Metalúrgicos de hasta cinco Productos.", "Error", MMsgType:="Error", MFrm:=FRM, MTargetControl:=.Name, MRequired:=True, MSourceControl:=VCmdBtn.Name): GoTo EH 'Just 5 products as limit of balmets and One product at least
'                Let DblAux = Application.WorksheetFunction.Sum(VecWeights)
'                If DblAux > DblSum Then Call CMSGBOX.CMsgBoxFluid("¡Suma de Productos excede el valor de Alimentación!", "Error", MMsgType:="Error", MFrm:=FRM, MTargetControl:=.Name, MRequired:=True, MSourceControl:=VCmdBtn.Name): GoTo EH 'Weights can't be greater than Feed
'                Let VecWeights = Empty
'              End With
'
'                'Product Names
'              With .Controls("Txtbal5")
'                If Not REGEX.isEmptyStringReg(StrNames) Then
'                  Let VecNames = RES.resArrayGetDataColonSeparatedFromString(StrNames, j, True, StrCustomRegex:="%")
'                  If Not (j = 0) Then 'Just 5 products as limit of balmets and One product at least
'                    If Not (j = x1) Then Call CMSGBOX.CMsgBoxFluid("Cantidad de Nombres de Productos deben ser iguales a cantidad de Productos.", "Error", MMsgType:="Error", MFrm:=FRM, MTargetControl:=.Name, MRequired:=True, MSourceControl:=VCmdBtn.Name): GoTo EH 'just one DM on balmet And Weights must be equal or less than Grade units
'                  End If
'                End If
'              End With
'
'                'Grades Units
'              With .Controls("Txtbal6")
'                If Not REGEX.isEmptyStringReg(StrUnits) Then
'                  Let VecUnits = RES.resArrayGetDataColonSeparatedFromString(StrUnits, x2, True, StrCustomRegex:="%")
'                  If Not (x2 = 0) Then 'Just 5 products as limit of balmets and One product at least
'                    Let i = RES.resArrayHasValueN(VecUnits, "dm", ALCase:=True)
'                    If (x2 > 1 And i >= 1) Or (x1 > 1 And i >= 1) Or (x2 > x1 Or x2 = 0) Or (x2 < x1 And x2 > 1) Then Call CMSGBOX.CMsgBoxFluid("¿Minerales complejos en el Balance Principal?", "Error", MMsgType:="Error", MFrm:=FRM, MTargetControl:=.Name, MRequired:=True, MSourceControl:=VCmdBtn.Name): GoTo EH 'just one DM on balmet And Weights must be equal or less than Grade units
'                    'Refill array with same unit if more than 1 product is required for the balance
'                    If x2 = 1 And x1 > 1 Then 'User send just one value, the optional or a custom grade unit an more than 1 Grades
'                      For i = 1 To (x1 - 1)
'                        ReDim Preserve VecUnits(i)
'                        Let VecUnits(i) = VecUnits(0)
'                      Next i
'                    End If
'                  End If
'                End If
'              End With
'
'                'Grade(s)
'              With .Controls("Txtbal2")
'                If REGEX.isEmptyStringReg(StrGrades) Then Call CMSGBOX.CMsgBoxFluid("¡Leyes son requeridas!", "Error", MMsgType:="Error", MFrm:=FRM, MTargetControl:=.Name, MRequired:=True, MSourceControl:=VCmdBtn.Name): GoTo EH 'Grades are required
'                Let VecGrades = RES.resArrayGetDataColonSeparatedFromString(StrGrades, x3, StrAbs:=True)
'                If x3 = 0 Or Not (x3 = x1) Then Call CMSGBOX.CMsgBoxFluid("¡Cantidad de Leyes debe ser igual a la de Productos!", "Error", MMsgType:="Error", MFrm:=FRM, MTargetControl:=.Name, MRequired:=True, MSourceControl:=VCmdBtn.Name): Call CMSGBOX.CMsgBoxFluid("¡Falta valor de Alimentación!", "Error", MMsgType:="Error", MFrm:=FRM, MTargetControl:=.Name, MRequired:=True, MSourceControl:=VCmdBtn.Name): GoTo EH 'Same Grades, same weights
'                For i = 0 To UBound(VecGrades) 'Validate values for each Grade
'                  Let DblAux = VBA.Abs(VBA.CDbl(VecGrades(i)))
'                  Select Case VBA.LCase(VecUnits(i))
'                    Case "%": If DblAux > CDBL_PC_TOP Then Call CMSGBOX.CMsgBoxFluid("¿Ley mayor al 100%?", "Error", MMsgType:="Error", MFrm:=FRM, MTargetControl:=.Name, MRequired:=True, MSourceControl:=VCmdBtn.Name): GoTo EH 'Same Grades, same weights'No one Grade can exceed 100%
'                    Case "dm": If DblAux > CDBL_DM_TOP Then Call CMSGBOX.CMsgBoxFluid("¡Ley DM mayor al máximo!", "Error", MMsgType:="Error", MFrm:=FRM, MTargetControl:=.Name, MRequired:=True, MSourceControl:=VCmdBtn.Name): GoTo EH 'Same Grades, same weights 'No one DM Grade can exceed 10000DM
'                    Case "g/t", "g/l": If DblAux > CDBL_GT_TOP Then Call CMSGBOX.CMsgBoxFluid("¡Ley mayor al máximo!", "Error", MMsgType:="Error", MFrm:=FRM, MTargetControl:=.Name, MRequired:=True, MSourceControl:=VCmdBtn.Name): GoTo EH 'Same Grades, same weights 'No one DM Grade can exceed 1000000[g/T|g/L]
'                    Case Else: GoTo EH
'                  End Select
'                Next i
'              End With
'
'                'Recuperation(s)
'              With .Controls("Txtbal3")
'                If Not REGEX.isEmptyStringReg(StrRecs) Then 'Recs are required
'                  Let VecRecs = RES.resArrayGetDataColonSeparatedFromString(StrRecs, x4, StrAbs:=True)
'                  If x4 = 0 Or Not (x4 = x1) Then Call CMSGBOX.CMsgBoxFluid("¡Cantidad de Recuperaciones debe ser igual a cantidad de Productos!", "Error", MMsgType:="Error", MFrm:=FRM, MTargetControl:=.Name, MRequired:=True, MSourceControl:=VCmdBtn.Name): GoTo EH 'Same Recs, same weights
'                  For i = 0 To UBound(VecRecs)
'                    If VecRecs(i) > CDBL_PC_TOP Then Call CMSGBOX.CMsgBoxFluid("¿Porcentaje de Recuperación mayor al 100%?", "Error", MMsgType:="Error", MFrm:=FRM, MTargetControl:=.Name, MRequired:=True, MSourceControl:=VCmdBtn.Name): GoTo EH 'Recuperations are percentages, so its limit is 100%
'                  Next i
'                End If
'                Let VecFeed = Empty
'                Let VecWeights = Empty
'                Let VecGrades = Empty
'                Let VecRecs = Empty
'                Let VecUnits = Empty
'              End With
'
'              'GradesCx Units and GradesCX
'              If BooBalCx Then
'                Let VecGradesCx = RES.resArrayGetDataColonSeparatedFromString(StrGradesCx, x5, StrAbs:=True)
'                With .Controls("Txtbal8")
'                  If x5 > 1 Or x5 = 0 Then Call CMSGBOX.CMsgBoxFluid("¡Sólo  se requiere un valor de Ley para el Complejo!", "Error", MMsgType:="Error", MFrm:=FRM, MTargetControl:=.Name, MRequired:=True, MSourceControl:=VCmdBtn.Name): GoTo EH 'Just one Grade cx
'                End With
'                'GradesCx Units
'                If Not REGEX.isEmptyStringReg(StrUnitsCx) Then 'Units cx are required
'                  Let VecUnitsCx = RES.resArrayGetDataColonSeparatedFromString(StrUnitsCx, x2, True, StrCustomRegex:="%")
'                  With .Controls("Cbobal3")
'                    If x2 = 0 Or x2 > 1 Then Call CMSGBOX.CMsgBoxFluid("¡Sólo una Unidad de Ley para el material Complejo!", "Error", MMsgType:="Error", MFrm:=FRM, MTargetControl:=.Name, MRequired:=True, MSourceControl:=VCmdBtn.Name): GoTo EH 'Just one Unit cx
'                    Let StrAux = VBA.LCase(VBA.CStr(VecUnitsCx(0)))
'                    If Not (StrAux = "dm" Or StrAux = "g/t" Or StrAux = "g/l") Then Call CMSGBOX.CMsgBoxFluid("¡Unidad de Ley incorrecta! (introduzca DM, g/T o g/L)", "Error", MMsgType:="Error", MFrm:=FRM, MTargetControl:=.Name, MRequired:=True, MSourceControl:=VCmdBtn.Name): GoTo EH
'                  End With
'
'                  'Complex Grade
'                  With .Controls("Txtbal8")
'                    For i = 0 To UBound(VecGradesCx) 'Validate values for each Grade
'                      Let DblAux = VBA.Abs(VBA.CDbl(VecGradesCx(i)))
'                      Select Case VBA.LCase(VecUnitsCx(i))
'                        Case "dm": If DblAux > CDBL_DM_TOP Then Call CMSGBOX.CMsgBoxFluid("¡DM superior al máximo!", "Error", MMsgType:="Error", MFrm:=FRM, MTargetControl:=.Name, MRequired:=True, MSourceControl:=VCmdBtn.Name): GoTo EH 'No one DM Grade can exceed 10000DM
'                        Case "g/t", "g/l": If DblAux > CDBL_GT_TOP Then Call CMSGBOX.CMsgBoxFluid("¡Ley superior al máximo!", "Error", MMsgType:="Error", MFrm:=FRM, MTargetControl:=.Name, MRequired:=True, MSourceControl:=VCmdBtn.Name): GoTo EH 'No one DM Grade can exceed 1000000[g/T|g/L]
'                      End Select
'                    Next i
'                  End With
'                End If
'                Let VecGradesCx = Empty
'                Let VecUnitsCx = Empty
'
'                'Complex Recs
'                With .Controls("Txtbal0")
'                  If Not REGEX.isEmptyStringReg(StrRecsCx) Then 'Recs are required, but by default it is 50%
'                    Let VecRecsCx = RES.resArrayGetDataColonSeparatedFromString(StrRecsCx, x3, StrAbs:=True)
'                    If x3 = 0 Or x3 > 1 Then Call CMSGBOX.CMsgBoxFluid("¡Sólo se requiere una Recuperación en el Complejo!", "Error", MMsgType:="Error", MFrm:=FRM, MTargetControl:=.Name, MRequired:=True, MSourceControl:=VCmdBtn.Name): GoTo EH  'Just one Recueration cx
'                    If Application.WorksheetFunction.Sum(VecRecsCx) > 100 Then Call CMSGBOX.CMsgBoxFluid("¡Porcentaje de Recuperación en Complejos mayor a 100%!", "Error", MMsgType:="Error", MFrm:=FRM, MTargetControl:=.Name, MRequired:=True, MSourceControl:=VCmdBtn.Name): GoTo EH 'Recuperations are percentages, so its limit is 100%
'                    Let VecRecsCx = Empty
'                  End If
'                End With
'              End If
'              Set RES = Nothing
'              Set REGEX = Nothing
'
'              'Method validation
'              If FRM.Controls("Cbobal1").ListIndex = 0 And x1 > 1 Then Call RES.resFrmCustomMsgboxRequiredControl(FRM.Controls("Cbobal1"), "El método Convencional sólo admite Balances Metalúrgicos de dos productos.", VBA.vbInformation): GoTo EH
'
'              '3. [x]: Validate range place if Opt is active
'              'Check for spreadsheet edges and data inside the region selected
'              If .Controls("Optbal1").Value Then
'                If .Controls("Chkbal7").Value Then 'All results
'                  Let iIndex = (3 + (l * 4) + (m * 5) + 2) 'l: grades cols, m: complex cols, 3: initial fixed cols, 2: %p & Rc
'                Else
'                  Let iIndex = ( _
'                      3 + _
'                      (l * (2 + IIf(.Controls("Chkbal7").Value, 1, 0) + IIf(.Controls("Chkbal8").Value, 1, 0)) + _
'                      (m * (2 + IIf(.Controls("Chkbal7").Value, 1, 0) + IIf(.Controls("Chkbal8").Value, 1, 0) + IIf(.Controls("Chkbal9").Value, 1, 0) + IIf(.Controls("Chkbal10").Value, 1, 0))) + _
'                      (IIf(.Controls("Chkbal6").Value, 1, 0) + IIf(.Controls("Chkbal11").Value, 1, 0)))) + _
'                      5 + (IIf(.Controls("Chkbal12").Value, 1, 0) * 2) 'l: grades cols, m: complex cols, 3: initial fixed cols, 2: %p & Rc
'                End If
'                If GLORNG_CELL Is Nothing Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtbal10"), "¡Debe seleccionar una celda para imprimir los resultados!", VBA.vbInformation): GoTo EH
'                If RES.resExcelEdges(GLORNG_CELL, iIndex, 10) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtbal10"), "La celda que eligió está muy cerca de los bordes de la hoja de cálculo, este ejercicio requiere más espacio. Seleccione otra celda mpas alejada de los bordes.", VBA.vbInformation): GoTo EH
'                If (Application.WorksheetFunction.CountA(GLORNG_CELL.Range("A1").Resize(iIndex, 20)) > 0) Then
'                  If VBA.MsgBox(STRAPPTITLE & " sobreescribirá cualquier dato que se encuentre alrededor de la celda seleccionada. Existen datos que podrían ser eliminados." & VBA.vbNewLine & "¿continuar de todos modos?", vbYesNo + vbQuestion, STRAPPTITLE) = vbNo Then GoTo EH
'                End If
'              End If
'
'              Let xMsg = "Se calculará el Balance Metalúrgico con los datos que ingresó." & VBA.vbNewLine & "¿Proceder?"
'              Let View_Core_validations_balmet = (VBA.MsgBox(xMsg, vbQuestion + vbYesNo, STRAPPTITLE) = vbYes)
'
'            Case Else: GoTo EH
'          End Select
'
'        Case Else: GoTo EH
'      End Select
'    End With
'
'  Else 'Validation for Selections contents
'    Select Case ControlId
'      Case "balmetselection-getcell": Let View_Core_validations_balmet = True 'Go to get a Cellrange
'      Case "balmetselection"
'        Set RngBox = Selection
'        If RngBox Is Nothing Then GoTo EH
'
'        'Get number of Balmet elements
'        Set RES = New AppResources
'        Let StrAux0 = RES.resArrayGetDataColonSeparatedStringFromVariants(RngBox, i, False, True, RBooRngRowsCols:=True) 'Numbers
'        '(i-1: just grades, no feed)
'        If Not RES.ResBalmetGetIsBalmet(i - 1) Then Call VBA.MsgBox(StrAux1, vbInformation, STRAPPTITLE): GoTo EH 'Invalid data for Balmet
'        Let vSplits = VBA.Split(StrAux0, ";")
'        If VBA.CDbl(vSplits(0)) = 0 Then Call VBA.MsgBox("¡La Alimentación nunca debe ser cero!", vbInformation, STRAPPTITLE): GoTo EH 'Invalid data for Balmet (i-1: just grades, no feed)
'        If VBA.Len(VBA.CStr(vSplits(0))) > 10 Then Call VBA.MsgBox("¡La Alimentación es un número muy grande o con muchos decimales!", vbInformation, STRAPPTITLE): GoTo EH 'Invalid data for Balmet (i-1: just grades, no feed)
'
'        'Debug.Print "Alim + Leyes: ", StrAux0, i, i - 1
'        'Get Strings: Names and Units
'        Let StrAux2 = RES.resArrayGetDataColonSeparatedStringFromVariants(RngBox, j, True, RStrCustomRegex:="%", RBooRngRowsCols:=True)
'        Set RngBox = Nothing
'        Debug.Print "Nombres + Unidades", StrAux2, j
'        If j > 0 Then 'just if selection have some unity or name
'          Let n = RES.ResBalmetGetProductsBounds(i - 1, "cols") 'Get number of Balmet columns
'          Let m = RES.ResBalmetGetProductsBounds(i - 1, "rows") 'Get number of Balmet rows
'
'          Debug.Print "Cols:", n, "Rows:", m
'
'          Let StrAux0 = RES.ResBalmetGetType(i - 1) 'Get the type of Balmet
'          Let StrAux1 = "Datos incorrectos." & VBA.vbNewLine & VBA.vbNewLine & _
'                       "los Nombres y Leyes de los Productos deben ser iguales al número de Productos menos uno." & VBA.vbNewLine & _
'                       "Ej. Para el Balance de dos productos se requiere un Nombre y una Ley." & VBA.vbNewLine & _
'                       "Los Productos Complejos deben tener sólo un Nombre y una Ley." & VBA.vbNewLine & VBA.vbNewLine & _
'                       "Corrija el error y vuelva a intentarlo."
'
'          Debug.Print "Tipo Balmet:", StrAux0
'
'          If StrAux0 = VBA.vbNullString Then Call VBA.MsgBox("Datos incorrectos, corríjalos y vuelva a intentarlo.", vbInformation, STRAPPTITLE): GoTo EH
'
'          If StrAux0 = "complex" Then 'Names and Units would be zero, equal to cols + 1 or equal to 2cols + 1
'            If RES.ResBalmetGetComplexDatas(StrAux2, i) = VBA.vbNullString Then 'If exists more than one Name Complex Mineral: wrong data
'              Call VBA.MsgBox("Los Minerales Complejos deben ser tratados independientemente o en Procesos Diferenciales.", vbInformation, STRAPPTITLE): GoTo EH
'            End If
'            If RES.ResBalmetGetComplexDatas(StrAux2, i, True) = VBA.vbNullString Then 'If exists more than one Unit Complex Mineral: wrong data
'              Call VBA.MsgBox("Los Minerales Complejos deben ser tratados independientemente o en Procesos Diferenciales.", vbInformation, STRAPPTITLE): GoTo EH
'            End If
'          End If
'
'          'Names and Units would be zero, equal to cols or equal to 2cols
'          If RES.ResBalmetGetUnits(StrAux2, i) = VBA.vbNullString Then 'Exist Complex in normal Balmet
'            Call VBA.MsgBox("Los Minerales Complejos deben ser tratados independientemente o en Procesos Diferenciales.", vbInformation, STRAPPTITLE): GoTo EH
'          End If
'          If Not RES.ResBalmetGetComplexDatas(StrAux2, i, False, False) = VBA.vbNullString Then Call VBA.MsgBox("Los Minerales Complejos deben ser tratados independientemente o en Procesos Diferenciales.", vbInformation, STRAPPTITLE): GoTo EH
'          Let StrAux2 = RES.resArrayGetDataColonSeparatedStringFromVariants(StrAux2, i, True) 'Get Names
'          If Not (i = n) And (i > 0) Then Call VBA.MsgBox("Error, Cantidad de Nombres de Productos diferentes de cantidad de Productos valiosos.", vbInformation, STRAPPTITLE): GoTo EH
'        End If
'        Set RES = Nothing
'
'        '3. [x]: Validate range place if Opt is active
'        'Check for spreadsheet edges and data inside the region selected
'        If GLOBOO_SELECCTIONCELL Then
'          Let iIndex = (3 + (n * 4) + (n * 5) + 2) 'l: grades cols, m: complex cols, 3: initial fixed cols, 2: %p & Rc
'          If GLORNG_CELL Is Nothing Then Call VBA.MsgBox("¡Debe seleccionar una celda para imprimir los resultados!", VBA.vbInformation, STRAPPTITLE): GoTo EH
'          If RES.resExcelEdges(GLORNG_CELL, iIndex, 10) Then Call VBA.MsgBox("La celda que eligió está muy cerca de los bordes finales de la hoja de cálculo, este ejercicio requiere más espacio. Seleccione otra celda más alejada de los bordes finales.", VBA.vbInformation, STRAPPTITLE): GoTo EH
'          If (Application.WorksheetFunction.CountA(GLORNG_CELL.Range("A1").Resize(iIndex, 20)) > 0) Then
'            If VBA.MsgBox(STRAPPTITLE & " sobreescribirá cualquier dato que se encuentre alrededor de la celda seleccionada. Existen datos que podrían ser eliminados." & VBA.vbNewLine & "¿Continuar de todos modos?", vbYesNo + vbQuestion, STRAPPTITLE) = vbNo Then GoTo EH
'          End If
'        End If
'
'        Let View_Core_validations_balmet = False ' True
'
'      Case Else: GoTo EH
'    End Select
'  End If
  
EH:
  Let vSplits = Empty
  Set RngBox = Nothing
  Set CMsgBox = Nothing
  Set REGEX = Nothing
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)
End Function

Private Function View_Core_validations_projections( _
  ByVal VCmdBtn As MSForms.CommandButton, _
  ByVal FRM As FRMBasics) As Boolean
  
  Dim RES As New AppResources
  Dim REGEX As New AppResRegEx
  Dim CMsgBox As New ClassCustomMsgbox
  
  Dim vSplits As Variant
  Dim xMsg As String, StrAux0 As String
  Dim iIndex As Long, i As Long, j As Long, m As Long
  Dim DblAux0 As Double
  
  Const sFN As String = "VIEW::-Validations_projections"

  On Error GoTo EH
  Let View_Core_validations_projections = False
  
'  With FRM
'    Select Case VCmdBtn.Name
'      Case "Btnpro0", "Btnpro1", "Btnpro2", "Btnpro3", "Btnpro4", "Btnpro5", "Btnpro6", "Btnpro7", "Btnpro8": Let View_Core_validations_projections = True
'
'      Case "Btnmain1"
'        'Inversion
'        With .Controls("Txtpro0")
'          Let vSplits = RES.resArrayGetDataArrayFromVariants(.Value, i, False, False, True)
'          If i = 0 Or Application.WorksheetFunction.Sum(vSplits) = 0 Then Call CMSGBOX.CMsgBoxFluid("¡Se requiere el valor de Inversión!", "Error", MMsgType:="Error", MFrm:=FRM, MTargetControl:=.Name, MRequired:=True, MSourceControl:=VCmdBtn.Name): GoTo EH
'        End With
'
'        'Fine Weight
'        With .Controls("Txtpro1")
'          Let vSplits = RES.resArrayGetDataArrayFromVariants(.Value, i, False, True)
'          If i = 0 Then Call CMSGBOX.CMsgBoxFluid("¡Se requiere como mínimo un valor de Peso de Mineral Fino!", "Error", MMsgType:="Error", MFrm:=FRM, MTargetControl:=.Name, MRequired:=True, MSourceControl:=VCmdBtn.Name): GoTo EH
'          If i > BAL_LIMIT Then Call RES.resFrmCustomMsgboxRequiredControl(FRM.Controls("Txtpro1"), "¡Sólo hasta " & BAL_LIMIT & " minerales a la vez!", VBA.vbInformation): GoTo EH
'        End With
'
'        'Ctz
'        With .Controls("Txtpro2")
'          Let vSplits = RES.resArrayGetDataSemicolonSeparatedSlashGroupedStringFromVariants(.Value, j, False)
'          If j = 0 Then Call CMSGBOX.CMsgBoxFluid("¡Se requiere como mínimo un rango de tres valores de Cotización del Mineral Fino!", "Error", MMsgType:="Error", MFrm:=FRM, MTargetControl:=.Name, MRequired:=True, MSourceControl:=VCmdBtn.Name): GoTo EH
'          If Not j = i Then Call RES.resFrmCustomMsgboxRequiredControl(FRM.Controls("Txtpro2"), "¡Tiene " & IIf(i = 1, "un mineral", i & " minerales") & ". Es necesario un rango de Cotizaciones por cada Peso Fino de Mineral!", VBA.vbInformation): GoTo EH
'          Let vSplits = VBA.Split(.Value, "/")
'          For m = 0 To UBound(vSplits)
'            Let StrAux0 = RES.resArrayGetDataColonSeparatedStringFromVariants(VBA.CStr(vSplits(m)), j, False, False, True)
'            If j = 0 Or j < 3 Then Call RES.resFrmCustomMsgboxRequiredControl(FRM.Controls("Txtpro2"), "¡Cada rango de Cotizaciones debe contener como mínimo tres valores!", VBA.vbInformation): GoTo EH
'          Next m
'          Let vSplits = Empty
'        End With
'
'        'Mineral Names
'        With .Controls("Txtpro4")
'          Let vSplits = RES.resArrayGetDataArrayFromVariants(.Value, j, True)
'          If j > 0 And Not j = i Then Call CMSGBOX.CMsgBoxFluid("¡Falta el Nombre de algún Mineral!", "Error", MMsgType:="Error", MFrm:=FRM, MTargetControl:=.Name, MRequired:=True, MSourceControl:=VCmdBtn.Name): GoTo EH
'        End With
'
'        'Mineral Ctz Units
'        With .Controls("Txtpro5")
'          Let vSplits = RES.resArrayGetDataArrayFromVariants(.Value, j, True)
'          If j > 0 And Not j = i Then Call CMSGBOX.CMsgBoxFluid("¡Falta la Unidad de Cotización de algún Mineral!", "Error", MMsgType:="Error", MFrm:=FRM, MTargetControl:=.Name, MRequired:=True, MSourceControl:=VCmdBtn.Name): GoTo EH
'        End With
'        Let vSplits = Empty
'
'        '%Outs
'        With .Controls("Txtpro6")
'          If Not REGEX.isPercentReg(.Value) Then Call RES.resFrmCustomMsgboxRequiredControl(FRM.Controls("Txtpro6"), "¡Se esperaba un porcentaje!", VBA.vbInformation): GoTo EH
'          Let DblAux0 = VBA.CDbl(VBA.Replace(.Value, "%", VBA.vbNullString))
'          If DblAux0 > 100 Or DblAux0 < 0 Then Call RES.resFrmCustomMsgboxRequiredControl(FRM.Controls("Txtpro6"), "¡Ningún porcentaje puede ser mayor a 100% ni menor que cero!", VBA.vbInformation): GoTo EH
'        End With
'
'        '%Time
'        With .Controls("Txtpro7")
'          If Not REGEX.isNumberReg(.Value) Then Call RES.resFrmCustomMsgboxRequiredControl(FRM.Controls("Txtpro7"), "¡Se esperaba un número!", VBA.vbInformation): GoTo EH
'        End With
'
'        '%Rate
'        With .Controls("Txtpro8")
'          If Not REGEX.isPercentReg(.Value) Then Call RES.resFrmCustomMsgboxRequiredControl(FRM.Controls("Txtpro8"), "¡Se esperaba un porcentaje!", VBA.vbInformation): GoTo EH
'          Let DblAux0 = VBA.CDbl(VBA.Replace(.Value, "%", VBA.vbNullString))
'          If DblAux0 > 100 Or DblAux0 < 0 Then Call RES.resFrmCustomMsgboxRequiredControl(FRM.Controls("Txtpro8"), "¡Ningún porcentaje puede ser mayor a 100% ni menor que cero!", VBA.vbInformation): GoTo EH
'        End With
'
'        'Check for spreadsheet edges and data inside the region selected
'        If .Controls("Optpro3").Value Then
'          Let iIndex = ((i + 6) * 5) '5: A1 to totalsrow aprox
'          If GLORNG_CELL Is Nothing Then Call CMSGBOX.CMsgBoxFluid("¡Debe seleccionar una celda para imprimir los resultados!", "Error", MMsgType:="Error", MFrm:=FRM, MTargetControl:="Txtpro9", MRequired:=True, MSourceControl:=VCmdBtn.Name): GoTo EH
'          If RES.resExcelEdges(GLORNG_CELL, iIndex, 50) Then
'            Call VBA.MsgBox("La celda que eligió está muy cerca de los bordes de la hoja de cálculo, este ejercicio requiere más espacio. Seleccione otra celda mpas alejada de los bordes.", STRAPPTITLE): GoTo EH
'          End If
'          If (Application.WorksheetFunction.CountA(GLORNG_CELL.Range("A1").Resize(iIndex, 50)) > 0) Then
'            If VBA.MsgBox(STRAPPTITLE & " sobreescribirá cualquier dato que se encuentre al rededor de la celda seleccionada. Existen datos que podrían ser eliminados." & VBA.vbNewLine & "¿continuar de todos modos?", vbYesNo + vbQuestion, STRAPPTITLE) = vbNo Then GoTo EH
'          End If
'        End If
'
'        Let xMsg = "Se calculará la Proyección del Proyecto con los datos que ingresó." & VBA.vbNewLine & "¿Proceder?"
'        Let View_Core_validations_projections = (VBA.MsgBox(xMsg, vbQuestion + vbYesNo, STRAPPTITLE) = vbYes)
'
'      Case Else: GoTo EH
'    End Select
'  End With
  
EH:
  Let vSplits = Empty
  Set CMsgBox = Nothing
  Set REGEX = Nothing
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)
End Function

Private Function View_Core_validations_oreinvoice( _
  ByVal VCmdBtn As MSForms.CommandButton, _
  ByVal FRM As FRMBasics) As Boolean
  
  Dim RES As New AppResources
  Dim REGEX As New AppResRegEx
  Dim CMsgBox As New ClassCustomMsgbox
  
  Dim RngBox As Range
  Dim vSplits As Variant
  Dim xMsg As String, StrAux0 As String, StrAux1 As String
  Dim iIndex As Long, jIndex As Long, i As Long, j As Long, k As Long, l As Long, m As Long
  Dim BooAux0 As Boolean
  
  Const sFN As String = "VIEW::-Validations_oreinvoice"

  On Error GoTo EH
  Let View_Core_validations_oreinvoice = False
  
'  With FRM
'    Select Case VCmdBtn.Name
'      Case "Btnpur0", "Btnpur5", "Btnpur9", "Btnpur10": Let View_Core_validations_oreinvoice = True 'Date btns & RangeCell
'
'      Case "Btnpur1", "Btnpur2" 'Provider names & Provider type
'        If VCmdBtn.Tag = "selection" Then 'To select range
'          Let View_Core_validations_oreinvoice = True
'        Else 'To validate range
'          Let StrAux0 = "Txtpur" & IIf(VCmdBtn.Name = "Btnpur1", 3, 4)
'          Let vSplits = RES.resArrayGetDataArrayFromVariants(.Controls(StrAux0).Value, i, True)
'          If i <= 1 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls(StrAux0), "¡Se esperaba al menos dos valores!", VBA.vbInformation): GoTo EH
'          If i > GLOSNG_PUR_LIMIT Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls(StrAux0), "¡Demasiados datos, la lista excede del límite de " & GLOSNG_PUR_LIMIT & " registros!", VBA.vbInformation): GoTo EH
'          If VCmdBtn.Name = "Btnpur2" Then
'            For i = 1 To UBound(vSplits)
'              If Not REGEX.isProviderType(VBA.CStr(vSplits(i))) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls(StrAux0), "Dato incorrecto. ¡No se admite el Tipo de provedor " & VBA.CStr(vSplits(i)) & "!, sólo se admiten los tipos: Particular ó  Cooperativista", VBA.vbInformation): GoTo EH
'            Next i
'          End If
'          Let vSplits = Empty
'        End If
'        Let View_Core_validations_oreinvoice = True
'
'      Case "Btnpur3", "Btnpur4" 'Add & Edit
'        Select Case .Controls("Cbopur1").ListIndex
'          Case 0 'Base data: 'Texts 4 - 8
'            'MUESTRA
'            If excliqlitedatasheet.ListObjects("excliqlitepurchasebasedata").ListRows.Count > GLOSNG_PUR_LIMIT Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¡Demasiados datos, la lista excede del límite de " & GLOSNG_PUR_LIMIT & " registros!", VBA.vbInformation): GoTo EH
'            If REGEX.isStringReg(.Controls("Txtpur5").Value) Then
'              Set RngBox = excliqlitedatasheet.ListObjects("excliqlitepurchasebasedata").ListColumns("MUESTRA").DataBodyRange
'              If Not RngBox Is Nothing Then
'                If Application.WorksheetFunction.CountIf(RngBox, .Controls("Txtpur5").Value) > 0 Then
'                  Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtpur5"), "¡La Muestra: " & .Controls("Txtpur5").Value & " ya existe!", VBA.vbInformation): GoTo EH
'                End If
'              End If
'              Set RngBox = Nothing
'            End If
'            'FECHA
'            If Not REGEX.isDateReg(.Controls("Txtpur6").Value) Then
'              Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtpur6"), "¡Se requiere una fecha válida!", VBA.vbInformation): GoTo EH
'            End If
'            '%HUMEDAD
'            If Not REGEX.isPercentReg(.Controls("Txtpur7").Value) Then
'              Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtpur7"), "¡El valor de Porcentaje de Humedad es inválido!", VBA.vbInformation): GoTo EH
'            End If
'            '%MERMA
'            If Not REGEX.isPercentReg(.Controls("Txtpur8").Value) Then
'              Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtpur8"), "¡El valor de Porcentaje de Merma es inválido!", VBA.vbInformation): GoTo EH
'            End If
'            'PBH
'            If Not REGEX.isNumberReg(.Controls("Txtpur9").Value) Then
'              Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtpur9"), "¡Se requiere un valor de Peso Bruto Húmedo válido!", VBA.vbInformation): GoTo EH
'            End If
'
'          Case 1 'Main data: 'Texts 6,7,8,9
'            If excliqlitedatasheet.ListObjects("excliqlitepurchasemaindata").ListRows.Count > GLOSNG_PUR_LIMIT Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¡Demasiados datos, la lista excede del límite de " & GLOSNG_PUR_LIMIT & " registros!", VBA.vbInformation): GoTo EH
'            If .Controls("Cbopur2").ListIndex = 1 Then 'TIPO
'              'CONTENIDOS
'              Let vSplits = RES.resArrayGetDataArrayFromVariants(.Controls("Txtpur6").Value, i, True)
'              If i = 0 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtpur6"), "¡Se esperaba al menos dos elementos!", VBA.vbInformation): GoTo EH
'              If RES.resArrayHasRepeatedElements(vSplits) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtpur6"), "¡Un mineral complejo debe contener distintos elementos!", VBA.vbInformation): GoTo EH
'              'LEYES
'              Let vSplits = RES.resArrayGetDataArrayFromVariants(.Controls("Txtpur7").Value, j)
'              If j = 0 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtpur7"), "¿Las leyes de los elementos?", VBA.vbInformation): GoTo EH
'              If Not j = i Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtpur7"), "¡Debe existir la misma cantidad de leyes y elementos!", VBA.vbInformation): GoTo EH
'              If RES.resArrayHasValueN(vSplits, "dm", False, True) > 1 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtpur7"), "¡Un mineral complejo sólo puede contener una unidad DM!", VBA.vbInformation): GoTo EH
'              'UNIDADES
'              Let vSplits = RES.resArrayGetDataArrayFromVariants(.Controls("Txtpur8").Value, j, True, RStrCustomRegex:="%")
'              If j = 0 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtpur8"), "¿Las unidades de los elementos?", VBA.vbInformation): GoTo EH
'              If Not j = i Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtpur8"), "¡Debe existir la misma cantidad de unidades, leyes y elementos!", VBA.vbInformation): GoTo EH
'              Let j = RES.resArrayHasValueN(vSplits, "dm", False, True)
'              If j > 1 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtpur8"), "¡Un mineral complejo no puede tener más de un elemento de Plata!", VBA.vbInformation): GoTo EH
'              'PRECIO/PTO
'              Let vSplits = RES.resArrayGetDataArrayFromVariants(.Controls("Txtpur9").Value, j)
'              If j = 0 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtpur9"), "¿Los precios de los elementos?", VBA.vbInformation): GoTo EH
'              If Not j = i Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtpur9"), "¡Debe existir la misma cantidad de precios, unidades, leyes y elementos!", VBA.vbInformation): GoTo EH
'              Let vSplits = Empty
'            Else
'              'LEYES
'              If Not REGEX.isNumberReg(.Controls("Txtpur7").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtpur7"), "¡Ley inválida!", VBA.vbInformation): GoTo EH
'              'PRECIO/PTO
'              If Not REGEX.isNumberReg(.Controls("Txtpur9").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtpur9"), "¡Precio inválido!", VBA.vbInformation): GoTo EH
'            End If
'
'          Case 2 'Bonus: 'Texts 5 - 6
'            If excliqlitedatasheet.ListObjects("excliqlitepurchasebonus").ListRows.Count > GLOSNG_PUR_LIMIT Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¡Demasiados datos, la lista excede del límite de " & GLOSNG_PUR_LIMIT & " registros!", VBA.vbInformation): GoTo EH
'            'BONO
'            If REGEX.isEmptyStringReg(.Controls("Txtpur5").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtpur5"), "¡El bono debe tener un nombre!", VBA.vbInformation): GoTo EH
'            'VALOR
'            If Not REGEX.isNumberReg(.Controls("Txtpur6").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtpur6"), "¡Bono inválido!", VBA.vbInformation): GoTo EH
'
'          Case 5 'Discounts: 'Texts 5 - 6
'            If excliqlitedatasheet.ListObjects("excliqlitepurchasediscounts").ListRows.Count > GLOSNG_PUR_LIMIT Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¡Demasiados datos, la lista excede del límite de " & GLOSNG_PUR_LIMIT & " registros!", VBA.vbInformation): GoTo EH
'            'DESCUENTO
'            If REGEX.isEmptyStringReg(.Controls("Txtpur5").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtpur5"), "¡El descuento debe tener un nombre!", VBA.vbInformation): GoTo EH
'            'VALOR
'            If Not REGEX.isNumberReg(.Controls("Txtpur6").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtpur6"), "¡Descuento inválido!", VBA.vbInformation): GoTo EH
'
'          Case Else: GoTo EH
'        End Select
'        Let View_Core_validations_oreinvoice = True
'
'      Case "Btnpur6" 'Del once
'        If Not (.Controls("Lstpur0").ListCount >= 1 And .Controls("Lstpur0").ListIndex > -1) Then GoTo EH
'        Let View_Core_validations_oreinvoice = (VBA.MsgBox("¿Realmente desea eliminar el ítem seleccionado?", vbYesNo, STRAPPTITLE) = vbYes)
'
'      Case "Btnpur7" 'Del all
'        If Not (.Controls("Lstpur0").ListCount > 0) Then GoTo EH
'        Let View_Core_validations_oreinvoice = (VBA.MsgBox("¿Realmente desea eliminar todos los datos?", vbYesNo, STRAPPTITLE) = vbYes)
'
'      Case "Btnpur8"
'        If VCmdBtn.Tag = "selection" Then 'To select range
'          Let View_Core_validations_oreinvoice = True
'        Else 'To validate range
'          If GLORNG_CELL Is Nothing Then GoTo EH
'          Select Case .Controls("Cbopur1").ListIndex
'            Case 0 'Base data: 'Texts 4 - 8
'              Let i = excliqlitedatasheet.ListObjects("excliqlitepurchasebasedata").ListRows.Count
'              Let j = GLORNG_CELL.Rows.Count
'              If i >= GLOSNG_PUR_LIMIT Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¡Demasiados datos, la lista de registros ya llegó al límite de " & GLOSNG_PUR_LIMIT & "!", VBA.vbInformation): GoTo EH
'              If (i + j) > GLOSNG_PUR_LIMIT Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¡Demasiados datos, la lista excede del límite de " & GLOSNG_PUR_LIMIT & " registros!", VBA.vbInformation): GoTo EH
'              If GLORNG_CELL.Columns.Count > 5 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¡Demasiadas columnas, sólo se requieren cinco en " & .Controls("Cbopur1").Value & "!", VBA.vbInformation): GoTo EH
'              If GLORNG_CELL.Columns.Count < 5 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¡Muy pocas columnas, se requieren cinco en " & .Controls("Cbopur1").Value & "!", VBA.vbInformation): GoTo EH
'              If Application.WorksheetFunction.CountA(GLORNG_CELL) = 0 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¿Y los Datos literales?", VBA.vbInformation): GoTo EH
'              For j = 1 To GLORNG_CELL.Columns.Count
'                For i = 1 To GLORNG_CELL.Rows.Count
'                  Let StrAux0 = GLORNG_CELL.Cells(i, j).Value
'                  Let StrAux1 = GLORNG_CELL.Cells(i, j).address(False, False)
'                  Select Case j
'                    Case 1 'MUESTRA
'                      If REGEX.isStringReg(StrAux0) Then
'                        Set RngBox = excliqlitedatasheet.ListObjects("excliqlitepurchasebasedata").ListColumns("MUESTRA").DataBodyRange
'                        If Not RngBox Is Nothing Then
'                          If Application.WorksheetFunction.CountIf(RngBox, StrAux0) > 1 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¡La Muestra: " & StrAux0 & " ya existe!", VBA.vbInformation): GoTo EH
'                        End If
'                        Set RngBox = Nothing
'                        If Application.WorksheetFunction.CountIf(GLORNG_CELL.Columns("A:A"), StrAux0) > 1 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¡La Muestra: " & StrAux0 & " ya existe en la lista de datos que introdujo, no se admiten Muestras repetidas!", VBA.vbInformation): GoTo EH
'                      Else
'                        Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¡Se requiere el nombre de la Muestra en la celda: " & StrAux1, VBA.vbInformation): GoTo EH
'                      End If
'
'                    Case 2 'FECHA
'                      If Not REGEX.isDateReg(StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¡Se requiere una fecha válida en la celda: " & StrAux1, VBA.vbInformation): GoTo EH
'                    Case 3 '%HUMEDAD
'                      If Not REGEX.isPercentReg(StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¡Se requiere un Porcentaje de humedad válido en la celda: " & StrAux1, VBA.vbInformation): GoTo EH
'                    Case 4 '%MERMA
'                      If Not REGEX.isPercentReg(StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¡Se requiere un Porcentaje de Merma válido en la celda: " & StrAux1, VBA.vbInformation): GoTo EH
'                    Case 5 'PBH
'                      If Not REGEX.isNumberReg(StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¡Se requiere un valor de Peso Bruto válido en la celda: " & StrAux1, VBA.vbInformation): GoTo EH
'                  End Select
'                Next i
'              Next j
'
'            Case 1 'Main data: 'Texts 6,7,8,9
'              Let i = excliqlitedatasheet.ListObjects("excliqlitepurchasemaindata").ListRows.Count
'              Let j = GLORNG_CELL.Rows.Count
'              If i >= GLOSNG_PUR_LIMIT Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¡Demasiados datos, la lista de registros ya llegó al límite de " & GLOSNG_PUR_LIMIT & "!", VBA.vbInformation): GoTo EH
'              If (i + j) > GLOSNG_PUR_LIMIT Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¡Demasiados datos, la lista excede del límite de " & GLOSNG_PUR_LIMIT & " registros!", VBA.vbInformation): GoTo EH
'              If GLORNG_CELL.Columns.Count > 5 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¡Demasiadas columnas, sólo se requieren cinco en " & .Controls("Cbopur1").Value & "!", VBA.vbInformation): GoTo EH
'              If GLORNG_CELL.Columns.Count < 5 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¡Muy pocas columnas, se requieren cinco en " & .Controls("Cbopur1").Value & "!", VBA.vbInformation): GoTo EH
'              If Application.WorksheetFunction.CountA(GLORNG_CELL) = 0 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¿Y los datos numéricos?", VBA.vbInformation): GoTo EH
'              For j = 1 To GLORNG_CELL.Columns.Count
'                For i = 1 To GLORNG_CELL.Rows.Count
'                  Let StrAux0 = GLORNG_CELL.Cells(i, j).Value
'                  Let StrAux1 = GLORNG_CELL.Cells(i, j).address(False, False)
'                  Select Case j
'                    Case 1 'TIPO MINERAL
'                      If Not REGEX.isMineralType(StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¡El tipo de mineral: " & StrAux0 & " de la celda " & StrAux1 & " no es admitido!", VBA.vbInformation): GoTo EH
'
'                    Case 2 'CONTENIDOS
'                      Let vSplits = RES.resArrayGetDataArrayFromVariants(StrAux0, k, True)
'                      If REGEX.isMineralComplex(GLORNG_CELL.Range("A" & i).Value) Then 'Mineral Complejo
'                        If k <= 1 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "Dato incorrecto en celda " & StrAux1 & ". ¡Se esperaba al menos dos elementos!", VBA.vbInformation): GoTo EH
'                        If RES.resArrayHasRepeatedElements(vSplits) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "Dato incorrecto en celda " & StrAux1 & " ¡Un mineral complejo debe contener distintos elementos!", VBA.vbInformation): GoTo EH
'                        For m = 0 To UBound(vSplits)
'                          If Not REGEX.isMineralExcliqlite(VBA.CStr(vSplits(m))) Then
'                            Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "Dato incorrecto en celda " & StrAux1 & " ¡El mineral: " & VBA.CStr(vSplits(m)) & " no es admitido!", VBA.vbInformation): GoTo EH
'                          End If
'                        Next m
'                      Else
'                        If k = 0 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "Dato incorrecto en celda " & StrAux1 & ". ¡Se esperaba al menos un elemento!", VBA.vbInformation): GoTo EH
'                        If RES.resArrayHasRepeatedElements(vSplits) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "Dato incorrecto en celda " & StrAux1 & " ¡Un mineral complejo debe contener distintos elementos!", VBA.vbInformation): GoTo EH
'                        If Not REGEX.isMineralExcliqlite(VBA.CStr(vSplits(0))) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "Dato incorrecto en celda " & StrAux1 & " ¡El mineral: " & VBA.CStr(vSplits(0)) & " no es admitido!", VBA.vbInformation): GoTo EH
'                      End If
'
'                    Case 3 'LEYES
'                      Let vSplits = RES.resArrayGetDataArrayFromVariants(StrAux0, l, False, True, True)
'                      If REGEX.isMineralComplex(GLORNG_CELL.Range("A" & i).Value) Then 'Mineral Complejo
'                        If l <= 1 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "Dato incorrecto en celda " & StrAux1 & ". ¡Se esperaba al menos dos Leyes de elementos!", VBA.vbInformation): GoTo EH
'                        If Not l = k Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "Dato incorrecto en celda " & StrAux1 & ". ¡Se esperaba igual cantidad de datos en Leyes de elementos y Contenidos!", VBA.vbInformation): GoTo EH
'                        If RES.resArrayHasRepeatedElements(vSplits) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "Dato incorrecto en celda " & StrAux1 & " ¡Un mineral complejo debe contener distintos elementos!", VBA.vbInformation): GoTo EH
'                      Else
'                        If l = 0 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "Dato incorrecto en celda " & StrAux1 & ". ¡Se esperaba al menos una ley de mineral!", VBA.vbInformation): GoTo EH
'                      End If
'
'                    Case 4 'UNIDADES
'                      Let vSplits = RES.resArrayGetDataArrayFromVariants(StrAux0, l, True, RStrCustomRegex:="%")
'                      If REGEX.isMineralComplex(GLORNG_CELL.Range("A" & i).Value) Then 'Mineral Complejo
'                        If l <= 1 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "Dato incorrecto en celda " & StrAux1 & ". ¡Se esperaba al menos dos unidades de elementos!", VBA.vbInformation): GoTo EH
'                        If RES.resArrayHasValueN(vSplits, "dm", False, True) > 1 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "Dato incorrecto en celda " & StrAux1 & " ¡Un mineral complejo sólo puede contener una unidad DM!", VBA.vbInformation): GoTo EH
'                        If Not l = k Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "Dato incorrecto en celda " & StrAux1 & ". ¡Se esperaba igual cantidad de datos en Unidad de Ley de elementos y Contenidos!", VBA.vbInformation): GoTo EH
'                        For m = 0 To UBound(vSplits)
'                          If Not REGEX.isMineralUnitExcliqlite(VBA.CStr(vSplits(m))) Then
'                            Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "Dato incorrecto en celda " & StrAux1 & " ¡La Unidad de mineral: " & VBA.CStr(vSplits(m)) & " no es admitida!", VBA.vbInformation): GoTo EH
'                          End If
'                        Next m
'                      Else
'                        If l = 0 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "Dato incorrecto en celda " & StrAux1 & ". ¡Se esperaba al menos una Unidad de ley de elemento!", VBA.vbInformation): GoTo EH
'                        If Not REGEX.isMineralUnitExcliqlite(VBA.CStr(vSplits(0))) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "Dato incorrecto en celda " & StrAux1 & " ¡La Unidad de ley mineral: " & VBA.CStr(vSplits(0)) & " no es admitida!", VBA.vbInformation): GoTo EH
'                      End If
'
'                    Case 5 'PRECIO/PTO
'                      Let vSplits = RES.resArrayGetDataArrayFromVariants(StrAux0, l, False, True, True)
'                      If REGEX.isMineralComplex(GLORNG_CELL.Range("A" & i).Value) Then 'Mineral Complejo
'                        If l <= 1 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "Dato incorrecto en celda " & StrAux1 & ". ¡Se esperaba al menos dos Precios de elementos!", VBA.vbInformation): GoTo EH
'                        If Not l = k Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "Dato incorrecto en celda " & StrAux1 & ". ¡Se esperaba igual cantidad de datos en Precios de elementos y Contenidos!", VBA.vbInformation): GoTo EH
'                      Else
'                        If l = 0 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "Dato incorrecto en celda " & StrAux1 & ". ¡Se esperaba al menos un Precio de mineral!", VBA.vbInformation): GoTo EH
'                      End If
'                      Let vSplits = Empty
'
'                  End Select
'                Next i
'              Next j
'
'            Case 2 'Bonus: 'Texts 5 - 6
'              Let i = excliqlitedatasheet.ListObjects("excliqlitepurchasebonus").ListRows.Count
'              Let j = GLORNG_CELL.Rows.Count
'              If i >= GLOSNG_PUR_LIMIT Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¡Demasiados datos, la lista de registros ya llegó al límite de " & GLOSNG_PUR_LIMIT & "!", VBA.vbInformation): GoTo EH
'              If (i + j) > GLOSNG_PUR_LIMIT Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¡Demasiados datos, la lista excede del límite de " & GLOSNG_PUR_LIMIT & " registros!", VBA.vbInformation): GoTo EH
'              If GLORNG_CELL.Columns.Count > 2 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¡Demasiadas columnas, sólo se requieren dos en " & .Controls("Cbopur1").Value & "!", VBA.vbInformation): GoTo EH
'              If GLORNG_CELL.Columns.Count < 2 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¡Muy pocas columnas, se requieren dos en " & .Controls("Cbopur1").Value & "!", VBA.vbInformation): GoTo EH
'              If Application.WorksheetFunction.CountA(GLORNG_CELL) = 0 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¿Y los Bonos? (Rellene con ceros si no existen)", VBA.vbInformation): GoTo EH
'              For j = 1 To GLORNG_CELL.Columns.Count
'                For i = 1 To GLORNG_CELL.Rows.Count
'                  Let StrAux0 = GLORNG_CELL.Cells(i, j).Value
'                  Let StrAux1 = GLORNG_CELL.Cells(i, j).address(False, False)
'                  Select Case j
'                    Case 1 'BONO
'                      If REGEX.isEmptyStringReg(StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "Dato incorrecto en celda " & StrAux1 & ". ¡Cada bono debe tener un nombre!", VBA.vbInformation): GoTo EH
'
'                    Case 2 'VALOR
'                      If Not REGEX.isNumberReg(StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "Dato incorrecto en celda " & StrAux1 & ". ¡Bono inválido!", VBA.vbInformation): GoTo EH
'                  End Select
'                Next i
'              Next j
'
'            Case 5 'Discounts: 'Texts 5 - 6
'              Let i = excliqlitedatasheet.ListObjects("excliqlitepurchasebonus").ListRows.Count
'              Let j = GLORNG_CELL.Rows.Count
'              If i >= GLOSNG_PUR_LIMIT Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¡Demasiados datos, la lista de registros ya llegó al límite de " & GLOSNG_PUR_LIMIT & "!", VBA.vbInformation): GoTo EH
'              If (i + j) > GLOSNG_PUR_LIMIT Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¡Demasiados datos, la lista excede del límite de " & GLOSNG_PUR_LIMIT & " registros!", VBA.vbInformation): GoTo EH
'              If GLORNG_CELL.Columns.Count > 2 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¡Demasiadas columnas, sólo se requieren dos en " & .Controls("Cbopur1").Value & "!", VBA.vbInformation): GoTo EH
'              If GLORNG_CELL.Columns.Count < 2 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¡Muy pocas columnas, se requieren dos en " & .Controls("Cbopur1").Value & "!", VBA.vbInformation): GoTo EH
'              If Application.WorksheetFunction.CountA(GLORNG_CELL) = 0 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¿Y los Descuentos? (Rellene con ceros si no existen)", VBA.vbInformation): GoTo EH
'              For j = 1 To GLORNG_CELL.Columns.Count
'                For i = 1 To GLORNG_CELL.Rows.Count
'                  Let StrAux0 = GLORNG_CELL.Cells(i, j).Value
'                  Let StrAux1 = GLORNG_CELL.Cells(i, j).address(False, False)
'                  Select Case j
'                    Case 1 'DESCUENTO
'                      If REGEX.isEmptyStringReg(StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "Dato incorrecto en celda " & StrAux1 & ". ¡Cada descuento debe tener un nombre!", VBA.vbInformation): GoTo EH
'
'                    Case 2 'VALOR
'                      If Not REGEX.isNumberReg(StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "Dato incorrecto en celda " & StrAux1 & ". ¡Descuento inválido!", VBA.vbInformation): GoTo EH
'                  End Select
'                Next i
'              Next j
'
'            Case Else: GoTo EH
'          End Select
'        End If
'        Let View_Core_validations_oreinvoice = True
'
'      Case "Btnmain1"
'        'Base Data
'        Set RngBox = excliqlitedatasheet.ListObjects("excliqlitepurchasebasedata").DataBodyRange
'        If RngBox Is Nothing Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¡No existen los Datos literales del mineral!", VBA.vbInformation): GoTo EH
'        If Application.WorksheetFunction.CountA(RngBox) = 0 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¡No existen los Datos literales del mineral!", VBA.vbInformation): GoTo EH
'        Let i = RngBox.Rows.Count: Set RngBox = Nothing
'        'Main Data
'        Set RngBox = excliqlitedatasheet.ListObjects("excliqlitepurchasemaindata").DataBodyRange
'        If RngBox Is Nothing Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¡No existen los Datos numéricos del mineral!", VBA.vbInformation): GoTo EH
'        If Application.WorksheetFunction.CountA(RngBox) = 0 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¡No existen los Datos numéricos del mineral!", VBA.vbInformation): GoTo EH
'        Let j = RngBox.Rows.Count: Set RngBox = Nothing
'        If Not j = i Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¡Se requiere la misma cantidad de filas en Datos numéricos del mineral y Datos literales del mineral!", VBA.vbInformation): GoTo EH
'        'Multiple
'        If .Controls("Chkpur1").Value Then
'          Let vSplits = RES.resArrayGetDataArrayFromVariants(.Controls("Txtpur4").Value, j, True)
'          If j = 0 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtpur4"), "¡El tipo de provedor debe ser un texto separado por puntos y comas especificando que sea de tipo Particular o Cooperativista!", VBA.vbInformation): GoTo EH
'          If Not j = i Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtpur4"), "¡Para generar varias liquidaciones, es necesario que existan tantos provedores como datos de liquidación!", VBA.vbInformation): GoTo EH
'          Let i = RES.resArrayHasValueN(vSplits, "particular", False, True)
'          Let i = i + RES.resArrayHasValueN(vSplits, "cooperativista", False, True)
'          If Not i = UBound(vSplits) + 1 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtpur4"), "¡Especifique el tipo de provedor utilizando los términos: Particular o Cooperativista!", VBA.vbInformation): GoTo EH
'          For i = 1 To j
'            'Retentions
'            Let BooAux0 = (VBA.CBool(excliqlitedatasheet.ListObjects("excliqlitepurchasevariousretentions").ListColumns("all").DataBodyRange.Cells(i, 1).Value) = False And _
'               excliqlitedatasheet.ListObjects("excliqlitepurchasevariousretentions").ListColumns("values").DataBodyRange.Cells(i, 1).Value = VBA.vbNullString)
'            If BooAux0 Then
'              If VBA.MsgBox("¡Ninguna Retención fue seleccionada para el Provedor N° " & i & ".!" & VBA.vbNewLine & "¿Asignar Retenciones?", vbYesNo + vbQuestion, STRAPPTITLE) = vbYes Then GoTo EH
'            End If
'            'Retentions other
'            Let BooAux0 = (VBA.CBool(excliqlitedatasheet.ListObjects("excliqlitepurchasevariousretentionsother").ListColumns("all").DataBodyRange.Cells(i, 1).Value) = False And _
'               excliqlitedatasheet.ListObjects("excliqlitepurchasevariousretentionsother").ListColumns("values").DataBodyRange.Cells(i, 1).Value = VBA.vbNullString)
'            If BooAux0 Then
'              If VBA.LCase(VBA.CStr(vSplits(i - 1))) = "cooperativista" Then
'                If VBA.MsgBox("¡Ninguna Retención Otra fue seleccionada para el Provedor N° " & i & ".!" & VBA.vbNewLine & "¿Asignar Retenciones Otras?", vbYesNo + vbQuestion, STRAPPTITLE) = vbYes Then GoTo EH
'              End If
'            Else
'              If VBA.LCase(VBA.CStr(vSplits(i - 1))) = "particular" Then
'                If VBA.MsgBox("¡El provedor N° " & i & " es de tipo " & VBA.StrConv(VBA.CStr(vSplits(i - 1)), vbProperCase) & ". No corresponde la aplicación de Retenciones otras, debe eliminar las retenciones!" & VBA.vbNewLine & "¿Corregir?", vbYesNo + vbQuestion, STRAPPTITLE) = vbYes Then GoTo EH
'              End If
'            End If
'            'Discounts
'            Let StrAux0 = Application.WorksheetFunction.TextJoin(";", True, excliqlitedatasheet.ListObjects("excliqlitepurchasediscounts").ListColumns("DESCUENTO").DataBodyRange)
'            Let StrAux1 = Application.WorksheetFunction.TextJoin(";", True, excliqlitedatasheet.ListObjects("excliqlitepurchasediscounts").ListColumns("VALOR").DataBodyRange)
'            Call RES.resArrayGetDataArrayFromVariants(StrAux0, k, True, RBooIncludeEmptyStrings:=True)
'            Call RES.resArrayGetDataArrayFromVariants(StrAux1, l, False, False, True)
'            If Not k = l Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¡La cantidad de Descuentos debe coincidir con la cantidad de sus valores!", VBA.vbInformation): GoTo EH
'            'Bonus
'            Let StrAux0 = Application.WorksheetFunction.TextJoin(";", True, excliqlitedatasheet.ListObjects("excliqlitepurchasebonus").ListColumns("BONO").DataBodyRange)
'            Let StrAux1 = Application.WorksheetFunction.TextJoin(";", True, excliqlitedatasheet.ListObjects("excliqlitepurchasebonus").ListColumns("VALOR").DataBodyRange)
'            Call RES.resArrayGetDataArrayFromVariants(StrAux0, k, True, RBooIncludeEmptyStrings:=True)
'            Call RES.resArrayGetDataArrayFromVariants(StrAux1, l, False, False, True)
'            If Not k = l Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¡La cantidad de Bonos debe coincidir con la cantidad de sus valores!", VBA.vbInformation): GoTo EH
'          Next i
'          Let vSplits = Empty
'        Else 'Simple
'          'Retentions
'          Let i = Application.WorksheetFunction.CountIf(excliqlitedatasheet.ListObjects("excliqliteconfigpurretentions").ListColumns("SELECTED").DataBodyRange, "=1")
'          If i = 0 Then
'            If VBA.MsgBox("¡Ninguna Retención fue seleccionada!" & VBA.vbNewLine & "¿Corregir?", vbYesNo + vbQuestion, STRAPPTITLE) = vbYes Then GoTo EH
'          End If
'          'Retentions other
'          Let i = Application.WorksheetFunction.CountIf(excliqlitedatasheet.ListObjects("excliqliteconfigpurretentionsother").ListColumns("SELECTED").DataBodyRange, "=1")
'          If i = 0 Then
'            If .Controls("Cbopur0").ListIndex = 0 Then
'              If VBA.MsgBox("¡Ninguna Retención Otra fue seleccionada!" & VBA.vbNewLine & "¿Corregir?", vbYesNo + vbQuestion, STRAPPTITLE) = vbYes Then GoTo EH
'            End If
'          Else
'            If .Controls("Cbopur1").ListIndex = 1 Then
'              If VBA.MsgBox("¡Este provedor es de tipo " & .Controls("Cbopur0").Value & ". No corresponde la aplicación de Retenciones otras!" & VBA.vbNewLine & "¿Corregir?", vbYesNo + vbQuestion, STRAPPTITLE) = vbYes Then GoTo EH
'            End If
'          End If
'          'Discounts
'          Let StrAux0 = Application.WorksheetFunction.TextJoin(";", True, excliqlitedatasheet.ListObjects("excliqlitepurchasediscounts").ListColumns("DESCUENTO").DataBodyRange)
'          Let StrAux1 = Application.WorksheetFunction.TextJoin(";", True, excliqlitedatasheet.ListObjects("excliqlitepurchasediscounts").ListColumns("VALOR").DataBodyRange)
'          Let vSplits = RES.resArrayGetDataArrayFromVariants(StrAux0, i, True, RBooIncludeEmptyStrings:=True)
'          Let vSplits = RES.resArrayGetDataArrayFromVariants(StrAux1, j, False, False, True)
'          If Not i = j Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¡La cantidad de Descuentos debe coincidir con la cantidad de sus valores!", VBA.vbInformation): GoTo EH
'          'Bonus
'          Let StrAux0 = Application.WorksheetFunction.TextJoin(";", True, excliqlitedatasheet.ListObjects("excliqlitepurchasebonus").ListColumns("BONO").DataBodyRange)
'          Let StrAux1 = Application.WorksheetFunction.TextJoin(";", True, excliqlitedatasheet.ListObjects("excliqlitepurchasebonus").ListColumns("VALOR").DataBodyRange)
'          Let vSplits = RES.resArrayGetDataArrayFromVariants(StrAux0, i, True, RBooIncludeEmptyStrings:=True)
'          Let vSplits = RES.resArrayGetDataArrayFromVariants(StrAux1, j, False, False, True)
'          If Not i = j Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstpur0"), "¡La cantidad de Bonos debe coincidir con la cantidad de sus valores!", VBA.vbInformation): GoTo EH
'        End If
'
'        'Check for spreadsheet edges and data inside the region selected
'        If .Controls("Optpur1").Value Then
'          Let iIndex = 150 '50: A1 to totalsrow aprox
'          Let jIndex = 9 * IIf(.Controls("Chkpur5").Value, 2, 1) * IIf(.Controls("Chkpur1").Value, .Controls("Lstpur0").ListCount, 1) + 100
'          If GLORNG_CELL Is Nothing Then Call CMSGBOX.CMsgBoxFluid("¡Debe seleccionar una celda para imprimir los resultados!", "Error", MMsgType:="Error", MFrm:=FRM, MTargetControl:="Txtpur10", MRequired:=True, MSourceControl:=VCmdBtn.Name): GoTo EH
'          If RES.resExcelEdges(GLORNG_CELL, iIndex, jIndex) Then
'            Call VBA.MsgBox("La celda que eligió está muy cerca de los bordes de la hoja de cálculo, este ejercicio requiere más espacio. Seleccione otra celda más alejada de los bordes.", STRAPPTITLE): GoTo EH
'          End If
'          If (Application.WorksheetFunction.CountA(GLORNG_CELL.Range("A1").Resize(iIndex, jIndex)) > 0) Then
'            If VBA.MsgBox(STRAPPTITLE & " sobreescribirá cualquier dato que se encuentre al rededor de la celda seleccionada. Existen datos que podrían ser eliminados." & VBA.vbNewLine & "¿Continuar de todos modos?", vbYesNo + vbQuestion, STRAPPTITLE) = vbNo Then GoTo EH
'          End If
'        End If
'
'        Let xMsg = "Se generarán las Liquidaciones de minerales con los datos que ingresó." & VBA.vbNewLine & "¿Proceder?"
'        Let View_Core_validations_oreinvoice = (VBA.MsgBox(xMsg, vbQuestion + vbYesNo, STRAPPTITLE) = vbYes)
'
'      Case Else: GoTo EH
'    End Select
'  End With
  
EH:
  Let vSplits = Empty
  Set RngBox = Nothing
  Set CMsgBox = Nothing
  Set REGEX = Nothing
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)
End Function

Private Function View_Core_validations_concentrateinvoice( _
  ByVal VCmdBtn As MSForms.CommandButton, _
  ByVal FRM As FRMBasics) As Boolean
  
  Dim RES As New AppResources
  Dim REGEX As New AppResRegEx
  Dim CMsgBox As New ClassCustomMsgbox

  Dim RngBox As Range
  Dim vSplits As Variant
  Dim xMsg As String, StrAux0 As String, StrAux1 As String, StrAux2 As String, StrAux3 As String
  Dim iIndex As Long, jIndex As Long, i As Long, j As Long, k As Long, l As Long, m As Long
  Dim BooAux0 As Boolean
  Dim DblAux0 As Double, DblAux1 As Double

  Const sFN As String = "VIEW::-Validations_concentrateinvoice"

  On Error GoTo EH
  Let View_Core_validations_concentrateinvoice = False

'  With FRM
'    Select Case VCmdBtn.Name
'      Case "Btncnc0", "Btncnc1", "Btncnc4", "Btnmain3": Let View_Core_validations_concentrateinvoice = True 'Date btns & RangeCell
'
'      Case "Btncnc2", "Btncnc3" 'Add & Edit
'        Select Case .Controls("Cbocnc1").ListIndex
'          Case 0 'Contents
'            Let StrAux0 = "excliqliteconpurcontents"
'            'Contents
'            Let i = IIf(VCmdBtn.Name = "Btncnc2", 0, 1)
'            If excliqlitedatasheet.ListObjects(StrAux0).ListRows.Count > GLOSNG_PUR_LIMIT Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Demasiados datos, la lista excede del límite de " & GLOSNG_PUR_LIMIT & " registros!", VBA.vbInformation): GoTo EH
'            Set RngBox = excliqlitedatasheet.ListObjects(StrAux0).ListColumns("CONTENIDOS").DataBodyRange
'            If Not RngBox Is Nothing Then
'              If Application.WorksheetFunction.CountIf(RngBox, .Controls("Cbocnc2").Value) > i Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Cbocnc2"), "¡El Mineral: " & .Controls("Cbocnc2").Value & " ya existe en los registros!", VBA.vbInformation): GoTo EH
'            End If
'            Set RngBox = Nothing
'
'            'Grade
'            If Not REGEX.isNumberReg(.Controls("Txtcnc12").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcnc12"), "¡Falta la Ley del Mineral!", VBA.vbInformation): GoTo EH
'            If VBA.Abs(VBA.CDbl(.Controls("Txtcnc12").Value)) = 0 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcnc12"), "¡La Ley del Mineral debe ser un número positivo mayor a cero!", VBA.vbInformation): GoTo EH
'            'Min
'            If Not REGEX.isNumberReg(.Controls("Txtcnc13").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcnc13"), "¡Deducción mínima inválida!", VBA.vbInformation): GoTo EH
'            '%Payable
'            If Not REGEX.isPercentReg(.Controls("Txtcnc14").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcnc14"), "¡El Porcentaje Pagable es inválido!", VBA.vbInformation): GoTo EH
'
'          Case 1 'Refinement
'            Let StrAux0 = "excliqliteconpurrefination"
'            'Concept
'            Let i = IIf(VCmdBtn.Name = "Btncnc2", 0, 1)
'            If excliqlitedatasheet.ListObjects(StrAux0).ListRows.Count > GLOSNG_PUR_LIMIT Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Demasiados datos, la lista excede del límite de " & GLOSNG_PUR_LIMIT & " registros!", VBA.vbInformation): GoTo EH
'            Set RngBox = excliqlitedatasheet.ListObjects(StrAux0).ListColumns("CONCEPTO").DataBodyRange
'            If Not RngBox Is Nothing Then
'              If Application.WorksheetFunction.CountIf(RngBox, .Controls("Cbocnc2").Value) > i Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Cbocnc2"), "¡El Concepto: " & .Controls("Cbocnc2").Value & ", ya existe en los registros!", VBA.vbInformation): GoTo EH
'            End If
'            Set RngBox = Nothing
'            'Condition '2.
'            Let BooAux0 = REGEX.isNumberReg(.Controls("Txtcnc12").Value)
'            If Not BooAux0 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcnc12"), "¡La condición es inválida!", VBA.vbInformation): GoTo EH
'            '%Base|Cur/T '1.
'            If BooAux0 Then
'              If Not REGEX.isPercentReg(.Controls("Txtcnc13").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcnc13"), "¡Porcentaje inválido!", VBA.vbInformation): GoTo EH
'            Else
'              If Not REGEX.isNumberReg(.Controls("Txtcnc13").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcnc13"), "¡Base inválida!", VBA.vbInformation): GoTo EH
'            End If
'
'          Case 2 'Penalties: just in edition
'            If excliqlitedatasheet.ListObjects("excliqliteconpurpenalties").ListRows.Count > 15 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Demasiados datos, la lista excede del límite de 15 registros!", VBA.vbInformation): GoTo EH
'            'Current
'            If Not REGEX.isNumberReg(.Controls("Txtcnc12").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcnc12"), "¡Ley actual del elemento inválida!", VBA.vbInformation): GoTo EH
'            'Base
'            If Not REGEX.isNumberReg(.Controls("Txtcnc13").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcnc13"), "¡Ley Base del elemento inválida!", VBA.vbInformation): GoTo EH
'            'Deduction
'            If Not REGEX.isNumberReg(.Controls("Txtcnc14").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcnc14"), "¡Cantidad deducible inválida!", VBA.vbInformation): GoTo EH
'            'Limit
'            If Not REGEX.isPercentReg(.Controls("Txtcnc15").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcnc15"), "¡Porcentaje de deducción inválido!", VBA.vbInformation): GoTo EH
'
'          Case 3 'Deductions
'            Select Case .Controls("Cbocnc2").ListIndex
'              Case 0 'Deductions: by exportations just edition
'                Let StrAux0 = "excliqliteconpurdedexp"
'                'Concept
'                Let i = IIf(VCmdBtn.Name = "Btncnc2", 0, 1)
'                If REGEX.isEmptyStringReg(.Controls("Txtcnc12").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcnc12"), "¡Se esperaba un concepto!", VBA.vbInformation): GoTo EH
'                If excliqlitedatasheet.ListObjects(StrAux0).ListRows.Count > GLOSNG_PUR_LIMIT Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Demasiados datos, la lista excede del límite de " & GLOSNG_PUR_LIMIT & " registros!", VBA.vbInformation): GoTo EH
'                Set RngBox = excliqlitedatasheet.ListObjects(StrAux0).ListColumns("CONCEPTO").DataBodyRange
'                If Not RngBox Is Nothing Then
'                  If Application.WorksheetFunction.CountIf(RngBox, .Controls("Txtcnc12").Value) > i Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcnc12"), "¡El concepto: " & .Controls("Txtcnc12").Value & " ya existe en los registros!", VBA.vbInformation): GoTo EH
'                End If
'                Set RngBox = Nothing
'                'Percent
'                If Not REGEX.isPercentReg(.Controls("Txtcnc13").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcnc13"), "¡El Porcentaje es inválido!", VBA.vbInformation): GoTo EH
'                Let DblAux0 = VBA.Abs(VBA.CDbl(.Controls("Txtcnc13").Value))
'                'Value
'                If Not REGEX.isNumberReg(.Controls("Txtcnc14").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcnc14"), "¡Valor inválido!", VBA.vbInformation): GoTo EH
'                Let DblAux1 = VBA.Abs(VBA.CDbl(.Controls("Txtcnc14").Value))
'
'                If (DblAux0 = 0 And DblAux1 = 0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Se esperaba un porcentaje o un valor!", VBA.vbInformation): GoTo EH
'                If (DblAux0 > 0 And DblAux1 > 0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Sólo puede habilitar un porcentaje o un valor!", VBA.vbInformation): GoTo EH
'
''              Case 0 ', 3 ', 4 'Deductions: by exportations and discounts', rollback and discounts
''                Let StrAux0 = IIf(i = 0, "excliqliteconpurdedexp", IIf(i = 3, "excliqliteconpurdedrollback", "excliqliteconpurdeddiscounts"))
''                'Concept
''                If REGEX.isEmptyStringReg(.Controls("Txtcnc12").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcnc12"), "¡Se esperaba un concepto!", VBA.vbInformation): GoTo EH
''                If excliqlitedatasheet.ListObjects(StrAux0).ListRows.Count > GLOSNG_PUR_LIMIT Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Demasiados datos, la lista excede del límite de " & GLOSNG_PUR_LIMIT & " registros!", VBA.vbInformation): GoTo EH
''                Set RngBox = excliqlitedatasheet.ListObjects(StrAux0).ListColumns("CONCEPTO").DataBodyRange
''                If Not RngBox Is Nothing Then
''                  If Application.WorksheetFunction.CountIf(RngBox, .Controls("Txtcnc12").Value) > 0 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcnc12"), "¡El concepto: " & .Controls("Txtcnc12").Value & " ya existe en los registros!", VBA.vbInformation): GoTo EH
''                End If
''                Set RngBox = Nothing
''                'Percent
''                If Not REGEX.isPercentReg(.Controls("Txtcnc13").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcnc13"), "¡El Porcentaje es inválido!", VBA.vbInformation): GoTo EH
''                Let DblAux0 = VBA.Abs(VBA.CDbl(.Controls("Txtcnc13").Value))
''                'Value
''                If Not REGEX.isNumberReg(.Controls("Txtcnc14").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcnc14"), "¡Valor inválido!", VBA.vbInformation): GoTo EH
''                Let DblAux0 = VBA.Abs(VBA.CDbl(.Controls("Txtcnc14").Value))
''
''                If (DblAux0 = 0 And DblAux1 = 0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Se esperaba un porcentaje o un valor!", VBA.vbInformation): GoTo EH
''                If (DblAux0 > 0 And DblAux1 > 0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Sólo puede habilitar un porcentaje o un valor!", VBA.vbInformation): GoTo EH
'
'              Case 3 'Deductions: by discounts
'                Let StrAux0 = "excliqliteconpurdeddiscounts"
'                'Concept
'                Let i = IIf(VCmdBtn.Name = "Btncnc2", 0, 1)
'                If REGEX.isEmptyStringReg(.Controls("Txtcnc12").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcnc12"), "¡Se esperaba un concepto!", VBA.vbInformation): GoTo EH
'                If excliqlitedatasheet.ListObjects(StrAux0).ListRows.Count > GLOSNG_PUR_LIMIT Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Demasiados datos, la lista excede del límite de " & GLOSNG_PUR_LIMIT & " registros!", VBA.vbInformation): GoTo EH
'                Set RngBox = excliqlitedatasheet.ListObjects(StrAux0).ListColumns("CONCEPTO").DataBodyRange
'                If Not RngBox Is Nothing Then
'                  If Application.WorksheetFunction.CountIf(RngBox, .Controls("Txtcnc12").Value) > i Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcnc12"), "¡El concepto: " & .Controls("Txtcnc12").Value & " ya existe en los registros!", VBA.vbInformation): GoTo EH
'                End If
'                Set RngBox = Nothing
'                'Value
'                If Not REGEX.isNumberReg(.Controls("Txtcnc13").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcnc13"), "¡Valor inválido!", VBA.vbInformation): GoTo EH
'                Let DblAux0 = VBA.Abs(VBA.CDbl(.Controls("Txtcnc13").Value))
'                If (DblAux0 = 0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcnc13"), "¡Se requiere un valor diferente de cero!", VBA.vbInformation): GoTo EH
'
''              Case 1, 2 'Deductions: by retentions per law and retentions others 'OJO, SÓLO DEBE SER REVISADO EN BTN GENERAR
''                If .Controls("Cbocnc0").Value = "Particular" Then
''                  Set RngBox = excliqlitedatasheet.ListObjects("excliqliteconfigpurretentions").ListColumns("SELECTED").DataBodyRange
''                  If Application.WorksheetFunction.CountIf(RngBox, "=1") = 0 Then
''                    If VBA.MsgBox("El Provedor es de tipo Particular, debería realizar las retenciones de ley. ¿Pasar por alto?", vbQuestion + vbYesNo) = vbNo Then GoTo EH
''                  End If
''                  Set RngBox = excliqlitedatasheet.ListObjects("excliqliteconfigpurretentionsother").ListColumns("SELECTED").DataBodyRange
''                  If Application.WorksheetFunction.CountIf(RngBox, "=1") > 0 Then
''                    If VBA.MsgBox("El Provedor es de tipo Particular, no debería aplicarle las retenciones otras. ¿Pasar por alto?", vbQuestion + vbYesNo) = vbNo Then GoTo EH
''                  End If
''                Else
''                  Set RngBox = excliqlitedatasheet.ListObjects("excliqliteconfigpurretentions").ListColumns("SELECTED").DataBodyRange
''                  If Application.WorksheetFunction.CountIf(RngBox, "=1") = 0 Then
''                    If VBA.MsgBox("El Provedor es de tipo Cooperativista, debería realizar las retenciones de ley. ¿Pasar por alto?", vbQuestion + vbYesNo) = vbNo Then GoTo EH
''                  End If
''                  Set RngBox = excliqlitedatasheet.ListObjects("excliqliteconfigpurretentionsother").ListColumns("SELECTED").DataBodyRange
''                  If Application.WorksheetFunction.CountIf(RngBox, "=1") = 0 Then
''                    If VBA.MsgBox("El Provedor es de tipo Cooperativista, debería aplicarle las retenciones otras. ¿Pasar por alto?", vbQuestion + vbYesNo) = vbNo Then GoTo EH
''                  End If
''                End If
''                Set RngBox = Nothing
'
'              Case Else: GoTo EH
'            End Select
'
'          Case 4 'Bonus
'            Let StrAux0 = "excliqliteconpurbonus"
'            'Concept
'            Let i = IIf(VCmdBtn.Name = "Btncnc2", 0, 1)
'            If REGEX.isEmptyStringReg(.Controls("Txtcnc12").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcnc12"), "¡Se esperaba el concepto del bono!", VBA.vbInformation): GoTo EH
'            If excliqlitedatasheet.ListObjects(StrAux0).ListRows.Count > GLOSNG_PUR_LIMIT Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Demasiados datos, la lista excede del límite de " & GLOSNG_PUR_LIMIT & " registros!", VBA.vbInformation): GoTo EH
'            Set RngBox = excliqlitedatasheet.ListObjects(StrAux0).ListColumns("CONCEPTO").DataBodyRange
'            If Not RngBox Is Nothing Then
'              If Application.WorksheetFunction.CountIf(RngBox, .Controls("Txtcnc12").Value) > i Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcnc12"), "¡El Mineral: " & .Controls("Txtcnc12").Value & " ya existe en los registros!", VBA.vbInformation): GoTo EH
'            End If
'            Set RngBox = Nothing
'            'Value
'            If Not REGEX.isNumberReg(.Controls("Txtcnc13").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcnc14"), "¡Valor inválido!", VBA.vbInformation): GoTo EH
'            Let DblAux0 = VBA.Abs(VBA.CDbl(.Controls("Txtcnc13").Value))
'
'            If (DblAux0 = 0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Se esperaba un valor numérico!", VBA.vbInformation): GoTo EH
'
'          Case Else: GoTo EH
'        End Select
'        Let View_Core_validations_concentrateinvoice = True
'
'      Case "Btncnc5" 'Del once
'        If Not (.Controls("Lstcnc0").ListCount >= 1 And .Controls("Lstcnc0").ListIndex > -1) Then GoTo EH
'        Let View_Core_validations_concentrateinvoice = (VBA.MsgBox("¿Realmente desea eliminar el ítem seleccionado?", vbYesNo, STRAPPTITLE) = vbYes)
'
'      Case "Btncnc6" 'Del all
'        If Not (.Controls("Lstcnc0").ListCount > 0) Then GoTo EH
'        Let View_Core_validations_concentrateinvoice = (VBA.MsgBox("¿Realmente desea eliminar todos los datos?", vbYesNo, STRAPPTITLE) = vbYes)
'
'      Case "Btncnc7"
'        If VCmdBtn.Tag = "validate" Then 'F: To select range, T: To validate range
'          If GLORNG_CELL Is Nothing Then GoTo EH
'          If GLORNG_CELL.Areas.Count > 1 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Btncnc7"), "¡Realice solo una selección!", VBA.vbInformation): GoTo EH 'Just one area
'          Select Case .Controls("Cbocnc1").ListIndex
'            Case 0 'Contents
'              Let i = excliqlitedatasheet.ListObjects("excliqliteconpurcontents").ListRows.Count
'              Let j = GLORNG_CELL.Rows.Count
'              If i >= GLOSNG_PUR_LIMIT Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Demasiados datos, la lista de registros ya llegó al límite de " & GLOSNG_PUR_LIMIT & "!", VBA.vbInformation): GoTo EH
'              If (i + j) > GLOSNG_PUR_LIMIT Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Demasiados datos, la lista excede del límite de " & GLOSNG_PUR_LIMIT & " registros!", VBA.vbInformation): GoTo EH
'              If Not GLORNG_CELL.Columns.Count = 5 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Sólo se requiere una selección con cinco columnas, puede seleccionar hasta diez filas para los " & .Controls("Cbocnc1").Value & "!", VBA.vbInformation): GoTo EH
'              If Application.WorksheetFunction.CountA(GLORNG_CELL) = 0 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¿Y los Datos?", VBA.vbInformation): GoTo EH
'
'              For i = 1 To GLORNG_CELL.Rows.Count
'                'Contents
'                Let StrAux0 = GLORNG_CELL.Cells(i, 1).Value
'                Let StrAux1 = GLORNG_CELL.Cells(i, 1).address(False, False)
'                If REGEX.isStringReg(StrAux0) Then
'                  Set RngBox = excliqlitedatasheet.ListObjects("excliqliteconfigpurrm").ListColumns("MINERAL").DataBodyRange
'                  If Not RngBox Is Nothing Then
'                    If Application.WorksheetFunction.CountIf(RngBox, StrAux0) = 0 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡El contenido: " & StrAux0 & " no es admitido!", VBA.vbInformation): GoTo EH
'                  End If
'                  Set RngBox = Nothing
'                  If Application.WorksheetFunction.CountIf(GLORNG_CELL.Columns("A:A"), StrAux0) > 1 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡El contenido: " & StrAux0 & " ya existe en la lista de datos que introdujo, no se admiten Muestras repetidas!", VBA.vbInformation): GoTo EH
'                  If Not excliqlitedatasheet.ListObjects("excliqliteconpurcontents").DataBodyRange Is Nothing Then
'                    Set RngBox = excliqlitedatasheet.ListObjects("excliqliteconpurcontents").ListColumns("CONTENIDOS").DataBodyRange
'                    If Not RngBox Is Nothing Then
'                      If Application.WorksheetFunction.CountIf(RngBox, StrAux0) >= 1 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡La Muestra: " & StrAux0 & ", ya existe en la base de datos! No se admiten duplicados.", VBA.vbInformation): GoTo EH
'                    End If
'                    Set RngBox = Nothing
'                  End If
'                Else
'                  Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Se requiere el nombre del contenido en la celda: " & StrAux1 & ". Introduzca un nombre, ejemplo: Zinc", VBA.vbInformation): GoTo EH
'                End If
'                'Grade
'                Let StrAux0 = GLORNG_CELL.Cells(i, 2).Value
'                Let StrAux1 = GLORNG_CELL.Cells(i, 2).address(False, False)
'                If Not REGEX.isNumberReg(StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Se requiere un dato de ley válido en la celda: " & StrAux1, VBA.vbInformation): GoTo EH
'                If VBA.CDbl(StrAux0) = 0 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Ley inválida en celda: " & StrAux1 & ". No se admiten ceros.", VBA.vbInformation): GoTo EH
'                'Min
'                Let StrAux0 = GLORNG_CELL.Cells(i, 3).Value
'                Let StrAux1 = GLORNG_CELL.Cells(i, 3).address(False, False)
'                If Not REGEX.isNumberReg(StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Deducción mínima inválida en la celda: " & StrAux1, VBA.vbInformation): GoTo EH
'                Let DblAux0 = VBA.CDbl(StrAux0)
'                'Y/O
'                '%Payable
'                Let StrAux0 = GLORNG_CELL.Cells(i, 5).Value
'                Let StrAux1 = GLORNG_CELL.Cells(i, 5).address(False, False)
'                If Not REGEX.isPercentReg(StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Porcentaje pagable inválido en la celda: " & StrAux1, VBA.vbInformation): GoTo EH
'                Let DblAux1 = VBA.CDbl(StrAux0)
'                If DblAux1 > 100 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Porcentaje pagable inválido en la celda: " & StrAux1, VBA.vbInformation): GoTo EH
'                If DblAux0 = 0 And DblAux1 = 0 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Deducción mínima y Porcentaje pagable inválidos. Como mínimo, uno debe ser mayor a cero.", VBA.vbInformation): GoTo EH
'              Next i
'
'            Case 1 'Refinement
'              Let i = excliqlitedatasheet.ListObjects("excliqliteconpurrefination").ListRows.Count
'              Let j = GLORNG_CELL.Rows.Count
'              If i >= GLOSNG_PUR_LIMIT Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Demasiados datos, la lista de registros ya llegó al límite de " & GLOSNG_PUR_LIMIT & "!", VBA.vbInformation): GoTo EH
'              If (i + j) > GLOSNG_PUR_LIMIT Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Demasiados datos, la lista excede del límite de " & GLOSNG_PUR_LIMIT & " registros!", VBA.vbInformation): GoTo EH
'              If Not GLORNG_CELL.Columns.Count = 2 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Demasiadas columnas, sólo se requieren dos en " & .Controls("Cbocnc1").Value & "!", VBA.vbInformation): GoTo EH
'              Let i = excliqlitedatasheet.ListObjects("excliqliteconpurcontents").ListRows.Count
'              If Not i > 0 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Cbocnc1"), "¡Primero ingrese registros en " & .Controls("Cbocnc1").List(0) & "!", VBA.vbInformation): GoTo EH
'              If Not j = i Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡La cantidad de filas debe coincidir con la cantidad de registros de " & .Controls("Cbocnc1").List(0) & "!", VBA.vbInformation): GoTo EH
'              If Application.WorksheetFunction.CountA(GLORNG_CELL) = 0 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¿Y los datos?", VBA.vbInformation): GoTo EH
'              For i = 1 To GLORNG_CELL.Rows.Count
'                Let StrAux0 = GLORNG_CELL.Cells(i, 1).Value
'                Let StrAux1 = GLORNG_CELL.Cells(i, 1).address(False, False)
'                Let StrAux2 = GLORNG_CELL.Cells(i, 2).Value
'                Let StrAux3 = GLORNG_CELL.Cells(i, 2).address(False, False)
'                'Condition '2.
'                Let BooAux0 = REGEX.isNumberReg(StrAux0)
'                If Not BooAux0 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡La condición: " & StrAux0 & " de la celda " & StrAux1 & " es inválida!", VBA.vbInformation): GoTo EH
'                '%Base|Cur/T '1.
'                If BooAux0 Then
'                  If Not REGEX.isPercentReg(StrAux2) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡La base: " & StrAux2 & " de la celda " & StrAux3 & " es inválida!", VBA.vbInformation): GoTo EH
'                Else
'                  If Not REGEX.isNumberReg(StrAux2) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡La base: " & StrAux2 & " de la celda " & StrAux3 & " es inválida!", VBA.vbInformation): GoTo EH
'                End If
'              Next i
'
''            Case 2 'Penalties NOT RANGE SELECTION FOR PENALTIES, THIS CODE CAN BE UTIL IN CONFIG
''              Let i = excliqlitedatasheet.ListObjects("excliqliteconpurpenalties").ListRows.Count
''              Let j = GLORNG_CELL.Rows.Count
''              If j > i Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Los registros exceden el límite de " & i & "!", VBA.vbInformation): GoTo EH
''              If Not GLORNG_CELL.Columns.Count = 5 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Demasiadas columnas, sólo se requieren cinco en " & .Controls("Cbocnc1").Value & "!", VBA.vbInformation): GoTo EH
''              If Application.WorksheetFunction.CountA(GLORNG_CELL) = 0 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¿Y los datos?", VBA.vbInformation): GoTo EH
''              Set RngBox = excliqlitedatasheet.ListObjects("excliqliteconpurpenalties").ListColumns("ELEMENTO").DataBodyRange
''              If RngBox Is Nothing Then GoTo EH
''              For j = 1 To GLORNG_CELL.Columns.Count
''                For i = 1 To GLORNG_CELL.Rows.Count
''                  Let StrAux0 = GLORNG_CELL.Cells(i, j).Value
''                  Let StrAux1 = GLORNG_CELL.Cells(i, j).address(False, False)
''                  Select Case j
''                    Case 0 'Element
''                      If Application.WorksheetFunction.CountIf(RngBox, StrAux0) = 0 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡El elemento: " & StrAux0 & " de la celda: " & StrAux1 & " no existe en la base de datos!", VBA.vbInformation): GoTo EH
''                    Case 1 'Current
''                      If Not REGEX.isNumberReg(StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Ley actual: " & StrAux0 & " inválida en celda: " & StrAux1 & "!", VBA.vbInformation): GoTo EH
''                    Case 2 'Base
''                      If Not REGEX.isNumberReg(StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Ley base: " & StrAux0 & " inválida en celda: " & StrAux1 & "!", VBA.vbInformation): GoTo EH
''                    Case 3 'Deduction
''                      If Not REGEX.isNumberReg(StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Cantidad deducible: " & StrAux0 & " inválida en celda: " & StrAux1 & "!", VBA.vbInformation): GoTo EH
''                    Case 4 'Limit
''                      If Not REGEX.isPercentReg(StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Porcentaje de deducción: " & StrAux0 & " inválido en celda: " & StrAux1 & "!", VBA.vbInformation): GoTo EH
''                  End Select
''                Next i
''              Next j
'
'            Case 3
'              Select Case .Controls("Cbocnc2").ListIndex
'                Case 0 'Exports
'                  Let StrAux0 = "excliqliteconpurdedexp"
'                  Let i = excliqlitedatasheet.ListObjects(StrAux0).ListRows.Count
'                  Let j = GLORNG_CELL.Rows.Count
'                  If (j + i) > 10 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Los registros exceden el límite de 10!", VBA.vbInformation): GoTo EH
'                  If Not GLORNG_CELL.Columns.Count = 4 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Demasiadas columnas, sólo se requieren cuatro en " & .Controls("Cbocnc1").Value & "!", VBA.vbInformation): GoTo EH
'                  If Application.WorksheetFunction.CountA(GLORNG_CELL) = 0 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¿Y los datos?", VBA.vbInformation): GoTo EH
'
'                  Set RngBox = excliqlitedatasheet.ListObjects(StrAux0).ListColumns("CONCEPTO").DataBodyRange
'
'                  For i = 1 To GLORNG_CELL.Rows.Count
'                    Let StrAux0 = GLORNG_CELL.Cells(i, 1).Value
'                    Let StrAux1 = GLORNG_CELL.Cells(i, 1).address(False, False)
'                    'Concept
'                    If REGEX.isEmptyStringReg(StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Se esperaba un concepto en la celda " & StrAux1 & "!", VBA.vbInformation): GoTo EH
'                    If Application.WorksheetFunction.CountIf(GLORNG_CELL, StrAux0) > 1 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡El concepto: " & StrAux0 & " de la celda: " & StrAux1 & " tiene duplicados, solo se admiten valores únicos!", VBA.vbInformation): GoTo EH
'                    If Not RngBox Is Nothing Then
'                      If Application.WorksheetFunction.CountIf(RngBox, StrAux0) >= 1 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡El concepto: " & StrAux0 & " de la celda: " & StrAux1 & " ya existe en los registros, utilice otro concepto porque no se admiten duplicados!", VBA.vbInformation): GoTo EH
'                    End If
'
'                    'Percent
'                    Let StrAux0 = GLORNG_CELL.Cells(i, 2).Value
'                    Let StrAux1 = GLORNG_CELL.Cells(i, 2).address(False, False)
'                    If Not REGEX.isPercentReg(StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡El porcentaje: " & StrAux0 & " de la celda: " & StrAux1 & " es inválido!", VBA.vbInformation): GoTo EH
'                    Let DblAux0 = VBA.Abs(VBA.CDbl(StrAux0))
'
'                    'Value
'                    Let StrAux0 = GLORNG_CELL.Cells(i, 3).Value
'                    Let StrAux1 = GLORNG_CELL.Cells(i, 3).address(False, False)
'                    If Not REGEX.isNumberReg(StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡El valor: " & StrAux0 & " en la celda: " & StrAux1 & " es inválido!", VBA.vbInformation): GoTo EH
'                    Let DblAux1 = VBA.Abs(VBA.CDbl(StrAux0))
'
'                    If (DblAux0 = 0 And DblAux1 = 0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Se esperaba un porcentaje o un valor en las celdas: " & StrAux1 & " y su adyacente izquierdo!", VBA.vbInformation): GoTo EH
'                    If (DblAux0 > 0 And DblAux1 > 0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Sólo puede habilitar un porcentaje o un valor en las celdas: " & StrAux1 & " y su adyacente izquierdo!", VBA.vbInformation): GoTo EH
'                  Next i
'                  Set RngBox = Nothing
'
'                Case 3 'Discounts
'                  Let StrAux0 = "excliqliteconpurdeddiscounts"
'                  Let i = excliqlitedatasheet.ListObjects(StrAux0).ListRows.Count
'                  Let j = GLORNG_CELL.Rows.Count
'                  If (j + i) > 10 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Los registros exceden el límite de 10!", VBA.vbInformation): GoTo EH
'                  If Not GLORNG_CELL.Columns.Count = 2 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Demasiadas columnas, sólo se requieren dos en " & .Controls("Cbocnc1").Value & "!", VBA.vbInformation): GoTo EH
'                  If Application.WorksheetFunction.CountA(GLORNG_CELL) = 0 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¿Y los datos?", VBA.vbInformation): GoTo EH
'                  Set RngBox = excliqlitedatasheet.ListObjects(StrAux0).ListColumns("CONCEPTO").DataBodyRange
'
'                  For i = 1 To GLORNG_CELL.Rows.Count
'                    Let StrAux0 = GLORNG_CELL.Cells(i, 1).Value
'                    Let StrAux1 = GLORNG_CELL.Cells(i, 1).address(False, False)
'                    'Concept
'                    If REGEX.isEmptyStringReg(StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Se esperaba un concepto en la celda " & StrAux1 & "!", VBA.vbInformation): GoTo EH
'                    If Application.WorksheetFunction.CountIf(GLORNG_CELL, StrAux0) > 1 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡El concepto: " & StrAux0 & " de la celda: " & StrAux1 & " tiene duplicados, solo se admiten valores únicos!", VBA.vbInformation): GoTo EH
'                    If Not RngBox Is Nothing Then
'                      If Application.WorksheetFunction.CountIf(RngBox, StrAux0) >= 1 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡El concepto: " & StrAux0 & " de la celda: " & StrAux1 & " ya existe en los registros, utilice otro concepto porque no se admiten duplicados!", VBA.vbInformation): GoTo EH
'                    End If
'
'                    'Value
'                    Let StrAux0 = GLORNG_CELL.Cells(i, 2).Value
'                    Let StrAux1 = GLORNG_CELL.Cells(i, 2).address(False, False)
'                    If Not REGEX.isNumberReg(StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡El valor: " & StrAux0 & " en la celda: " & StrAux1 & " es inválido!", VBA.vbInformation): GoTo EH
'                    Let DblAux1 = VBA.Abs(VBA.CDbl(StrAux0))
'
'                    If (DblAux1 = 0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Se esperaba un valor numérico en la celda: " & StrAux1 & "!", VBA.vbInformation): GoTo EH
'                  Next i
'                  Set RngBox = Nothing
'
''                Case 1, 2 'Retentions per law and Retentions others  NOT RANGE SELECTION FOR RETENTIONS, THIS CODE CAN BE UTIL IN CONFIG
''                  If .Controls("Cbocnc0").Value = "Particular" Then
''                    Set RngBox = excliqlitedatasheet.ListObjects("excliqliteconfigpurretentions").ListColumns("SELECTED").DataBodyRange
''                    If Application.WorksheetFunction.CountIf(RngBox, "=1") = 0 Then
''                      If VBA.MsgBox("El Provedor es de tipo Particular, debería realizar las retenciones de ley. ¿Pasar por alto?", vbQuestion + vbYesNo) = vbNo Then GoTo EH
''                    End If
''                    Set RngBox = excliqlitedatasheet.ListObjects("excliqliteconfigpurretentionsother").ListColumns("SELECTED").DataBodyRange
''                    If Application.WorksheetFunction.CountIf(RngBox, "=1") > 0 Then
''                      If VBA.MsgBox("El Provedor es de tipo Particular, no debería aplicarle las retenciones otras. ¿Pasar por alto?", vbQuestion + vbYesNo) = vbNo Then GoTo EH
''                    End If
''                  Else
''                    Set RngBox = excliqlitedatasheet.ListObjects("excliqliteconfigpurretentions").ListColumns("SELECTED").DataBodyRange
''                    If Application.WorksheetFunction.CountIf(RngBox, "=1") = 0 Then
''                      If VBA.MsgBox("El Provedor es de tipo Cooperativista, debería realizar las retenciones de ley. ¿Pasar por alto?", vbQuestion + vbYesNo) = vbNo Then GoTo EH
''                    End If
''                    Set RngBox = excliqlitedatasheet.ListObjects("excliqliteconfigpurretentionsother").ListColumns("SELECTED").DataBodyRange
''                    If Application.WorksheetFunction.CountIf(RngBox, "=1") = 0 Then
''                      If VBA.MsgBox("El Provedor es de tipo Cooperativista, debería aplicarle las retenciones otras. ¿Pasar por alto?", vbQuestion + vbYesNo) = vbNo Then GoTo EH
''                    End If
''                  End If
''                  Set RngBox = Nothing
'
'                Case Else: GoTo EH
'              End Select
'
'            Case 4 'Bonus
'              Let StrAux0 = "excliqliteconpurbonus"
'              Let i = excliqlitedatasheet.ListObjects(StrAux0).ListRows.Count
'              Let j = GLORNG_CELL.Rows.Count
'              If (j + i) > 10 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Los registros exceden el límite de 10!", VBA.vbInformation): GoTo EH
'              If Not GLORNG_CELL.Columns.Count = 2 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Demasiadas columnas, sólo se requieren dos en " & .Controls("Cbocnc1").Value & "!", VBA.vbInformation): GoTo EH
'              If Application.WorksheetFunction.CountA(GLORNG_CELL) = 0 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¿Y los datos?", VBA.vbInformation): GoTo EH
'              Set RngBox = excliqlitedatasheet.ListObjects(StrAux0).ListColumns("CONCEPTO").DataBodyRange
'
'              For i = 1 To GLORNG_CELL.Rows.Count
'                'Concept
'                Let StrAux0 = GLORNG_CELL.Cells(i, 1).Value
'                Let StrAux1 = GLORNG_CELL.Cells(i, 1).address(False, False)
'                If REGEX.isEmptyStringReg(StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Se esperaba un concepto en la celda " & StrAux1 & "!", VBA.vbInformation): GoTo EH
'                If Application.WorksheetFunction.CountIf(GLORNG_CELL, StrAux0) > 1 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡El concepto: " & StrAux0 & " de la celda: " & StrAux1 & " tiene duplicados, solo se admiten valores únicos!", VBA.vbInformation): GoTo EH
'                If Not RngBox Is Nothing Then
'                  If Application.WorksheetFunction.CountIf(RngBox, StrAux0) >= 1 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡El concepto: " & StrAux0 & " de la celda: " & StrAux1 & " ya existe en los registros, utilice otro concepto porque no se admiten duplicados!", VBA.vbInformation): GoTo EH
'                End If
'
'                'Value
'                Let StrAux0 = GLORNG_CELL.Cells(i, 2).Value
'                Let StrAux1 = GLORNG_CELL.Cells(i, 2).address(False, False)
'                If Not REGEX.isNumberReg(StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡El valor: " & StrAux0 & " en la celda: " & StrAux1 & " es inválido!", VBA.vbInformation): GoTo EH
'                Let DblAux1 = VBA.Abs(VBA.CDbl(StrAux0))
'
'                If (DblAux1 = 0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Se esperaba un valor numérico en la celda: " & StrAux1 & "!", VBA.vbInformation): GoTo EH
'              Next i
'              Set RngBox = Nothing
'
'            Case Else: GoTo EH
'          End Select
'        End If
'        Let View_Core_validations_concentrateinvoice = True
'
'      Case "Btncnc8" 'Purchase, provider, concentrate and treatment costs datas
'        If VCmdBtn.Tag = "validate" Then 'T To select range, F To validate range
'          If GLORNG_CELL Is Nothing Then GoTo EH
'          If GLORNG_CELL.Areas.Count > 1 Then Call CMSGBOX.CMsgBoxFluid("¡Realice solo una selección!", "Error", MMsgType:="Error", MFrm:=FRM, MSourceControl:="Btncnc8"): GoTo EH 'Just one area
'          If Not GLORNG_CELL.Cells.Count = 13 Then Call CMSGBOX.CMsgBoxFluid("¡Se esperaba una selección de trece celdas!", "Error", MMsgType:="Error", MFrm:=FRM, MSourceControl:="Btncnc8"): GoTo EH 'Just one area
'          If Not Application.WorksheetFunction.CountA(GLORNG_CELL) = 13 Then Call CMSGBOX.CMsgBoxFluid("¡Se esperaba que cada celda de la selección tuviese datos!", "Error", MMsgType:="Error", MFrm:=FRM, MSourceControl:="Btncnc8"): GoTo EH 'Just one area
'          Let i = 0
'          ReDim vSplits(0)
'          For Each RngBox In GLORNG_CELL
'            ReDim Preserve vSplits(i)
'            Let vSplits(i) = RngBox.Value
'            Let i = i + 1
'          Next RngBox
'          Set RngBox = Nothing
'          If Not View_Core_validations_concentrateinvoice_basics(vSplits, FRM, RES, REGEX) Then GoTo EH
'        End If
'        Let View_Core_validations_concentrateinvoice = True
'
'      Case "Btnmain1"
'        'Base
'        Let vSplits = Array(.Controls("Txtcnc1").Value, .Controls("Txtcnc2").Value, .Controls("Txtcnc3").Value, "", "", "", .Controls("Txtcnc5").Value, .Controls("Txtcnc6").Value, .Controls("Txtcnc7").Value, .Controls("Txtcnc8").Value, .Controls("Txtcnc9").Value, .Controls("Txtcnc10").Value, .Controls("Txtcnc11").Value)
'        If Not View_Core_validations_concentrateinvoice_basics(vSplits, FRM, RES, REGEX) Then GoTo EH
'        Let vSplits = Empty
'        'Main
'        Let StrAux0 = "excliqliteconpurcontents"
'        Let StrAux1 = .Controls("Cbocnc1").List(0)
'        Let i = excliqlitedatasheet.ListObjects(StrAux0).ListRows.Count
'        If i > GLOSNG_PUR_LIMIT Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Demasiados datos, la lista excede del límite de " & GLOSNG_PUR_LIMIT & " registros!", VBA.vbInformation): GoTo EH
'        If i <= 0 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Faltan datos de " & StrAux1 & "!", VBA.vbInformation): GoTo EH
'
'        Let StrAux0 = "excliqliteconpurrefination"
'        Let j = excliqlitedatasheet.ListObjects(StrAux0).ListRows.Count
'        Let StrAux0 = .Controls("Cbocnc1").List(0)
'        Let StrAux1 = .Controls("Cbocnc1").List(1)
'        If j > 0 Then
'          If j > GLOSNG_PUR_LIMIT Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Demasiados datos, la lista excede del límite de " & GLOSNG_PUR_LIMIT & " registros!", VBA.vbInformation): GoTo EH
'          If Not j = i Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡La cantidad de registros de " & StrAux1 & "debe ser igual a la cantidad de registros de " & StrAux1 & "!", VBA.vbInformation): GoTo EH
'        End If
'
'        Let StrAux0 = "excliqliteconpurpenalties"
'        Let StrAux1 = .Controls("Cbocnc1").List(2)
'        Set RngBox = excliqlitedatasheet.ListObjects(StrAux0).ListColumns("SELECTED").DataBodyRange
'        If RngBox Is Nothing Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡No existen registros en " & StrAux1 & "!", VBA.vbInformation): GoTo EH
'        Let i = Application.WorksheetFunction.CountIf(RngBox, "=1")
'        If i = 0 Then
'          If VBA.MsgBox("No registró ninguna penalidad, ¿Pasar por alto?", vbQuestion + vbYesNo) = vbNo Then GoTo EH
'        End If
'        Set RngBox = Nothing
'
'        If .Controls("Cbocnc0").Value = "Particular" Then
'          Set RngBox = excliqlitedatasheet.ListObjects("excliqliteconfigpurretentions").ListColumns("SELECTED").DataBodyRange
'          If Application.WorksheetFunction.CountIf(RngBox, "=1") = 0 Then
'            If VBA.MsgBox("El Provedor es de tipo Particular, debería realizar las retenciones de ley. ¿Pasar por alto?", vbQuestion + vbYesNo) = vbNo Then GoTo EH
'          End If
'          Set RngBox = excliqlitedatasheet.ListObjects("excliqliteconfigpurretentionsother").ListColumns("SELECTED").DataBodyRange
'          If Application.WorksheetFunction.CountIf(RngBox, "=1") > 0 Then
'            If VBA.MsgBox("El Provedor es de tipo Particular, no debería aplicarle las retenciones otras. ¿Pasar por alto?", vbQuestion + vbYesNo) = vbNo Then GoTo EH
'          End If
'        Else
'          Set RngBox = excliqlitedatasheet.ListObjects("excliqliteconfigpurretentions").ListColumns("SELECTED").DataBodyRange
'          If Application.WorksheetFunction.CountIf(RngBox, "=1") = 0 Then
'            If VBA.MsgBox("El Provedor es de tipo Cooperativista, debería realizar las retenciones de ley. ¿Pasar por alto?", vbQuestion + vbYesNo) = vbNo Then GoTo EH
'          End If
'          Set RngBox = excliqlitedatasheet.ListObjects("excliqliteconfigpurretentionsother").ListColumns("SELECTED").DataBodyRange
'          If Application.WorksheetFunction.CountIf(RngBox, "=1") = 0 Then
'            If VBA.MsgBox("El Provedor es de tipo Cooperativista, debería aplicarle las retenciones otras. ¿Pasar por alto?", vbQuestion + vbYesNo) = vbNo Then GoTo EH
'          End If
'        End If
'        Set RngBox = Nothing
'        Let View_Core_validations_concentrateinvoice = True
'
'      Case Else: GoTo EH
'    End Select
'  End With

EH:
  Let vSplits = Empty
  Set RngBox = Nothing
  Set CMsgBox = Nothing
  Set REGEX = Nothing
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)
End Function

Private Function View_Core_validations_concentrateinvoice_basics( _
  ByVal VVecBasics As Variant, _
  ByVal FRM As FRMBasics, _
  ByVal RES As AppResources, _
  ByVal REGEX As AppResRegEx) As Boolean
  
  Dim StrAux0 As String
  Dim i As Byte
  
  Const sFN As String = "VIEW::-Validations_concentrateinvoice_basics"

  On Error GoTo EH
  Let View_Core_validations_concentrateinvoice_basics = False
  If Not VBA.IsArray(VVecBasics) Then GoTo EH
  
  With FRM
    For i = 0 To UBound(VVecBasics)
      Let StrAux0 = VBA.CStr(VVecBasics(i))
      Select Case i
        Case 0 'Basic Data: Date
          If Not REGEX.isDateReg(StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcnc0"), "¡Se esperaba una fecha!", VBA.vbInformation): GoTo EH
        Case 1 'Basic Data: STRPlace
          If REGEX.isEmptyStringReg(StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcnc1"), "¡Se esperaba el nombre de un Lugar!", VBA.vbInformation): GoTo EH
        Case 2 'Basic Data: DBLTC
          If Not REGEX.isNumberReg(StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcnc2"), "¡Se esperaba un número!", VBA.vbInformation): GoTo EH
        Case 3 'Basic Data: BOOME
  
        Case 4 'Provider Data: STRProvider
        Case 5 'Provider Data: BOOType
  
        Case 6 'Concentrate Data: Date
          If Not REGEX.isDateReg(StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcnc5"), "¡Se esperaba una fecha!", VBA.vbInformation): GoTo EH
        Case 7 'Concentrate Data: DBLWeight
          If Not REGEX.isNumberReg(StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcnc6"), "¡Se esperaba un número!", VBA.vbInformation): GoTo EH
        Case 8 'Concentrate Data: DBLH2O
          If Not REGEX.isNumberReg(StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcnc7"), "¡Se esperaba un número!", VBA.vbInformation): GoTo EH
        Case 9 'Concentrate Data: DBLDecrease
          If Not REGEX.isNumberReg(StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcnc8"), "¡Se esperaba un número!", VBA.vbInformation): GoTo EH
  
        Case 10 'Treatment cost Data: DBLMaquila
          If Not REGEX.isNumberReg(StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcnc9"), "¡Se esperaba un número!", VBA.vbInformation): GoTo EH
        Case 11 'Treatment cost Data: DBLBase
          If Not REGEX.isNumberReg(StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcnc10"), "¡Se esperaba un número!", VBA.vbInformation): GoTo EH
        Case 12 'Treatment cost Data: DBLScaler
          If Not REGEX.isNumberReg(StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcnc11"), "¡Se esperaba un número!", VBA.vbInformation): GoTo EH
      End Select
    Next i
  End With
  
  Let View_Core_validations_concentrateinvoice_basics = True

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)
End Function

Private Function View_Core_validations_config( _
  ByVal VCmdBtn As MSForms.CommandButton, _
  ByVal FRM As FRMBasics) As Boolean
  
  Dim RES As New AppResources
  Dim REGEX As New AppResRegEx
  Dim CMsgBox As New ClassCustomMsgbox

  Dim RngBox As Range
  Dim vSplits As Variant
  Dim xMsg As String, StrAux0 As String, StrAux1 As String, StrAux2 As String, StrAux3 As String
  Dim iIndex As Long, jIndex As Long, i As Long, j As Long, k As Long, l As Long, m As Long
  Dim BooAux0 As Boolean
  Dim DblAux0 As Double, DblAux1 As Double

  Const sFN As String = "VIEW::-Validations_config"

  On Error GoTo EH
  Let View_Core_validations_config = False

  With FRM
    Select Case VCmdBtn.Name
      'DIVISA
      Case "Btncon2": Let View_Core_validations_config = True: GoTo EH 'Ucn
      Case "Btncon4" 'Local
        If VBA.MsgBox("¿Realmente desea cambiar la moneda local a: " & .Controls("Cbocon0").Value & "?", vbQuestion + vbYesNo) = vbNo Then GoTo EH
      Case "Btncon5" 'Foreign
        If VBA.MsgBox("¿Realmente desea cambiar la moneda extranjera a: " & .Controls("Cbocon1").Value & "?", vbQuestion + vbYesNo) = vbNo Then GoTo EH
      Case "Btncon3" 'Del
        If VBA.MsgBox("¿Realmente desea eliminar la divisa seleccionada?", vbQuestion + vbYesNo) = vbNo Then GoTo EH
      Case "Btncon0", "Btncon1" 'Add, Upd (Divisa)
        If REGEX.isEmptyStringReg(.Controls("Txtcon0").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcon0"), "¡Se esperaba el nombre de una divisa!", VBA.vbInformation): GoTo EH
        If REGEX.isEmptyStringReg(.Controls("Txtcon1").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcon1"), "¡Se esperaba el nombre de una divisa!", VBA.vbInformation): GoTo EH
        If REGEX.isEmptyStringReg(.Controls("Txtcon2").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcon2"), "¡Se esperaba el símbolo de una divisa!", VBA.vbInformation): GoTo EH
        
        Let StrAux0 = "excliqlitecurrenciescon"
        Let j = IIf(VCmdBtn.Name = "Btncon2", 0, 1)
        If Not excliqlitedatasheet.ListObjects(StrAux0) Is Nothing Then
          If excliqlitedatasheet.ListObjects(StrAux0).ListRows.Count >= GLOSNG_PUR_LIMIT Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Lstcnc0"), "¡Demasiados datos, la lista excede del límite de " & GLOSNG_PUR_LIMIT & " registros!", VBA.vbInformation): GoTo EH
          Set RngBox = excliqlitedatasheet.ListObjects(StrAux0).ListColumns("currency").DataBodyRange
          If Not RngBox Is Nothing Then
            If Application.WorksheetFunction.CountIf(RngBox, .Controls("Txtcon0").Value) > j Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcon0"), "¡La divisa: " & .Controls("Txtcnc0").Value & " ya existe en los registros!", VBA.vbInformation): GoTo EH
          End If
          Set RngBox = Nothing
          Set RngBox = excliqlitedatasheet.ListObjects(StrAux0).ListColumns("currencies").DataBodyRange
          If Not RngBox Is Nothing Then
            If Application.WorksheetFunction.CountIf(RngBox, .Controls("Txtcon1").Value) > j Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcon0"), "¡La divisa: " & .Controls("Txtcnc1").Value & " ya existe en los registros!", VBA.vbInformation): GoTo EH
          End If
          Set RngBox = Nothing
        End If
        
      'PÁGINA
      Case "Btncon7" 'Save all
        If REGEX.isEmptyStringReg(.Controls("Txtcon4").Value) Then
          If VBA.MsgBox("¿Pasar por alto el nombre de su empresa?", vbQuestion + vbYesNo, STRAPPTITLE) = vbNo Then GoTo EH
        End If
        If REGEX.isEmptyStringReg(.Controls("Txtcon5").Value) Then
          If VBA.MsgBox("¿Pasar por alto la dirección de su empresa?", vbQuestion + vbYesNo, STRAPPTITLE) = vbNo Then GoTo EH
        End If
        If REGEX.isEmptyStringReg(.Controls("Txtcon6").Value) Then
          If VBA.MsgBox("¿Pasar por alto el teléfono de su empresa?", vbQuestion + vbYesNo, STRAPPTITLE) = vbNo Then GoTo EH
        Else
          If Not REGEX.isPhonenumber(.Controls("Txtcon6").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcon6"), "¡Ingrese un número de teléfono válido!", VBA.vbInformation): GoTo EH
        End If
        If REGEX.isEmptyStringReg(.Controls("Txtcon7").Value) Then
          If VBA.MsgBox("¿Pasar por alto el correo electrónico de su empresa?", vbQuestion + vbYesNo, STRAPPTITLE) = vbNo Then GoTo EH
        Else
          If Not REGEX.isEmail(.Controls("Txtcon7").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcon7"), "¡Ingrese un correo electrónico válido!", VBA.vbInformation): GoTo EH
        End If
    
      Case "Btncon10"
        If Not .Controls("Cbocon7").ListCount > 0 Then GoTo EH
        If Not .Controls("Cbocon7").Tag >= 0 Then GoTo EH
        If .Controls("Cbocon7").Tag = 0 Then GoTo EH
        If VBA.MsgBox("¿Realmente desea eliminar el ítem " & VBA.CStr(.Controls("Cbocon7").List(.Controls("Cbocon7").Tag, 0)) & " de los registros?", vbQuestion + vbYesNo, STRAPPTITLE) = vbNo Then GoTo EH
        
      Case "Btncon8", "Btncon9"
        If Not .Controls("Cbocon7").ListCount > 0 Then GoTo EH
        If Not .Controls("Cbocon7").Tag >= 0 Then GoTo EH
        Let StrAux0 = .Controls("Cbocon7").Value
        If REGEX.isEmptyStringReg(StrAux0) Then
          Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Cbocon7"), "¡Dato inválido!", VBA.vbInformation): GoTo EH
        Else
          If Application.WorksheetFunction.CountIf(excliqlitedatasheet.ListObjects("excliqliteconfigpurdestiny").ListColumns("purch_destiny").DataBodyRange, StrAux0) > 0 Then
            If VCmdBtn.Name = "Btncon8" Then GoTo EH
            If VCmdBtn.Name = "Btncon9" Then Call VBA.MsgBox("Dato modificado correctamente.", VBA.vbInformation, STRAPPTITLE): GoTo EH
          Else
            If VCmdBtn.Name = "Btncon8" Then
              If .Controls("Cbocon7").ListCount > 5 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Cbocon7"), "¡Límite de items alcanzado!", VBA.vbInformation): GoTo EH
            Else
              If .Controls("Cbocon7").Tag = 0 Then GoTo EH
            End If
          End If
        End If
        
      'LIQUIDACIONES
      Case "Btncon11", "Btncon12" 'Add, Upd
        Select Case .Controls("Cbocon8").ListIndex
          Case 0 'RM
            Let StrAux0 = VBA.CStr(.Controls("Txtcon8").Value)
            Let StrAux0 = VBA.CStr(.Controls("Txtcon9").Value)
            If REGEX.isEmptyStringReg(StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcon8"), "Se requiere el nombre de un mineral."): GoTo EH
            If Application.WorksheetFunction.CountIf(excliqlitedatasheet.ListObjects("excliqliteconfigpurrm").ListColumns("MINERAL").DataBodyRange, StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcon8"), "¡El ítem " & StrAux0 & " ya existe en la base de datos!"): GoTo EH
            If REGEX.isEmptyStringReg(.Controls("Txtcon9").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcon9"), "Se requiere el símbolo del elemento principal del mineral."): GoTo EH
            If Application.WorksheetFunction.CountIf(excliqlitedatasheet.ListObjects("excliqliteconfigpurrm").ListColumns("ELEMENTO").DataBodyRange, StrAux1) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcon9"), "¡El ítem " & StrAux1 & " ya existe en la base de datos!"): GoTo EH
            If Not REGEX.isNumberReg(.Controls("Txtcon10").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcon10"), "Se requiere un número."): GoTo EH
            If Not REGEX.isNumberReg(.Controls("Txtcon11").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcon11"), "Se requiere un número."): GoTo EH
            If Not REGEX.isNumberReg(.Controls("Txtcon12").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcon12"), "Se requiere un número."): GoTo EH
          Case 1 'Ret
            Let StrAux0 = VBA.CStr(.Controls("Txtcon8").Value)
            If REGEX.isEmptyStringReg(.Controls("Txtcon8").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcon8"), "El concepto es requerido."): GoTo EH
            If Application.WorksheetFunction.CountIf(excliqlitedatasheet.ListObjects("excliqliteconfigpurretentions").ListColumns("CONCEPTO").DataBodyRange, StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcon8"), "¡El ítem " & StrAux0 & " ya existe en la base de datos!"): GoTo EH
            If Not REGEX.isPercentReg(.Controls("Txtcon9").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcon9"), "Se requiere un porcentaje válido."): GoTo EH
          Case 2 'RetOth
            Let StrAux0 = VBA.CStr(.Controls("Txtcon8").Value)
            If REGEX.isEmptyStringReg(.Controls("Txtcon8").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcon8"), "El concepto es requerido."): GoTo EH
            If Application.WorksheetFunction.CountIf(excliqlitedatasheet.ListObjects("excliqliteconfigpurretentionsother").ListColumns("CONCEPTO").DataBodyRange, StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcon8"), "¡El concepto " & StrAux0 & " ya existe en la base de datos!"): GoTo EH
            If Not REGEX.isPercentReg(.Controls("Txtcon9").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcon9"), "Se requiere un porcentaje válido."): GoTo EH
          Case 3 'Penalties
            Let StrAux0 = VBA.CStr(.Controls("Txtcon8").Value)
            If REGEX.isEmptyStringReg(.Controls("Txtcon8").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcon8"), "Se requiere el símbolo del elemento principal de un mineral."): GoTo EH
            If Application.WorksheetFunction.CountIf(excliqlitedatasheet.ListObjects("excliqliteconpurpenalties").ListColumns("ELEMENTO").DataBodyRange, StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcon8"), "¡El elemento " & StrAux0 & " ya existe en la base de datos!"): GoTo EH
            If Not REGEX.isNumberReg(.Controls("Txtcon9").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcon9"), "Se requiere un número."): GoTo EH
            If Not REGEX.isNumberReg(.Controls("Txtcon10").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcon10"), "Se requiere un número."): GoTo EH
            If Not REGEX.isNumberReg(.Controls("Txtcon11").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcon11"), "Se requiere un número."): GoTo EH
            If Not REGEX.isNumberReg(.Controls("Txtcon12").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcon12"), "Se requiere un número."): GoTo EH
          Case 4 'Deductions in Exports
            Let StrAux0 = VBA.CStr(.Controls("Txtcon8").Value)
            If REGEX.isEmptyStringReg(.Controls("Txtcon8").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcon8"), "El concepto es requerido."): GoTo EH
            If Application.WorksheetFunction.CountIf(excliqlitedatasheet.ListObjects("excliqliteconpurdedexp").ListColumns("CONCEPTO").DataBodyRange, StrAux0) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcon8"), "¡El concepto " & StrAux0 & " ya existe en la base de datos!"): GoTo EH
            If Not REGEX.isPercentReg(.Controls("Txtcon9").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcon9"), "Se requiere un porcentaje válido."): GoTo EH
            If Not REGEX.isNumberReg(.Controls("Txtcon10").Value) Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcon10"), "Se requiere un número."): GoTo EH
            Let DblAux0 = VBA.CDbl(.Controls("Txtcon9").Value)
            Let DblAux1 = VBA.CDbl(.Controls("Txtcon10").Value)
            If DblAux0 = 0 And DblAux1 = 0 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcon9"), "Sólo un valor puede ser cero.", SelTxt:=True): GoTo EH
            If Not DblAux0 = 0 And Not DblAux1 = 0 Then Call RES.resFrmCustomMsgboxRequiredControl(.Controls("Txtcon9"), "Sólo un valor puede ser mayor que cero.", SelTxt:=True): GoTo EH
            
          Case Else: GoTo EH
        End Select
      
      Case "Btncon13" 'Ucn
      
      Case "Btncon14" 'Del
        If Not .Controls("Lstcon1").ListCount > 0 Then GoTo EH
        Select Case .Controls("Cbocon8").ListIndex
          Case 0 'RM
            If Not .Controls("Lstcon1").ListIndex > 10 Then GoTo EH
          Case 1 'Ret
            If Not .Controls("Lstcon1").ListIndex > 2 Then GoTo EH
          Case 2 'RetOth
            If Not .Controls("Lstcon1").ListIndex > 4 Then GoTo EH
          Case 3 'Penalties
            If Not .Controls("Lstcon1").ListIndex > 5 Then GoTo EH
          Case 4 'Deductions in Exports
          
          Case Else: GoTo EH
        End Select
        If VBA.MsgBox("¿Realmente desea eliminar el ítem " & VBA.CStr(.Controls("Lstcon1").List(.Controls("Lstcon1").ListIndex, 0)) & " de los registros?", vbQuestion + vbYesNo, STRAPPTITLE) = vbNo Then GoTo EH
    
    End Select
    Let View_Core_validations_config = True
  End With

EH:
  Let vSplits = Empty
  Set RngBox = Nothing
  Set CMsgBox = Nothing
  Set REGEX = Nothing
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)
End Function





'Values Preparation
Private Function View_Core_preparations( _
  ByVal RES As AppResources_current, _
  ByVal FRM As FRMBasics, _
  ByVal ControlId As String, _
  ByRef VVarRequest As Variant, _
  ByRef VBooNewSheet As Boolean, _
  ByRef VLngVecRngIndex As Long, _
  ByVal VRngBox As Range) As Boolean
  
  Const sFN As String = "VIEW::-ValuesPreparation"

  On Error GoTo EH
  Select Case ControlId
    Case "regression", "regressionselection": Let View_Core_preparations = View_Core_preparations_regression(RES, FRM, ControlId, VVarRequest, VBooNewSheet, VLngVecRngIndex, VRngBox)
'    Case "balmet", "balmetselection", "balmeteco": Let View_Core_preparations = View_Core_preparations_balmet(FRM, ControlId, VVarRequestBools, VVarRequest)
'
'    Case "projection": Let View_Core_preparations = View_Core_preparations_projection(FRM, ControlId, VVarRequestBools, VVarRequest)
'
'    Case "oreinvoice": Let View_Core_preparations = View_Core_preparations_oreinvoice(FRM, ControlId, VVarRequestBools, VVarRequest)
'
'    Case "concentrateinvoice": Let View_Core_preparations = View_Core_preparations_concentrateinvoice(FRM, ControlId, VVarRequestBools, VVarRequest)
'
'    Case "configlauncher": Let View_Core_preparations = True
    
    Case Else: GoTo EH
  End Select

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function View_Core_preparations_regression( _
  ByVal RES As AppResources_current, _
  ByVal FRM As FRMBasics, _
  ByVal ControlId As String, _
  ByRef VVarRequest As Variant, _
  ByRef VBooNewSheet As Boolean, _
  ByRef VLngVecRngIndex As Long, _
  ByVal VRngBox As Range) As Boolean
  
  Dim RngBox As Range
  Dim BooAux0 As Boolean, BooShowAll As Boolean
  
  Const sFN As String = "VIEW::-ValuesPreparation_regression"

  On Error GoTo EH
  Let View_Core_preparations_regression = False
  
  Select Case ControlId
    
    Case "regression"
      With FRM
        Let BooShowAll = .Controls("Chkreg0").Value 'ShowAll
        Let BooAux0 = .Controls("Chkreg1").Value 'JustRegression
        Let VBooNewSheet = .Controls("Optreg2").Value
        ReDim VVarRequest(16)
        Let VVarRequest(0) = .Controls("Txtreg0").Value                                        '0-X
        Let VVarRequest(1) = .Controls("Txtreg1").Value                                        '1-Y
        Let VVarRequest(2) = .Controls("Txtreg2").Value                                        '2-E
        Let VVarRequest(3) = BooShowAll                                                        '3-ShowAll
        Let VVarRequest(4) = BooAux0                                                           '4-JustRegression
        Let VVarRequest(5) = IIf(BooAux0, False, (.Controls("Chkreg2").Value Or BooShowAll))   '5-X2
        Let VVarRequest(6) = IIf(BooAux0, False, (.Controls("Chkreg3").Value Or BooShowAll))   '6-Y2
        Let VVarRequest(7) = IIf(BooAux0, False, (.Controls("Chkreg4").Value Or BooShowAll))   '7-XY
        Let VVarRequest(8) = IIf(BooAux0, False, (.Controls("Chkreg5").Value Or BooShowAll))   '8-a
        Let VVarRequest(9) = IIf(BooAux0, False, (.Controls("Chkreg6").Value Or BooShowAll))   '9-b
        Let VVarRequest(10) = IIf(BooAux0, False, (.Controls("Chkreg7").Value Or BooShowAll))  '10-r
        Let VVarRequest(11) = IIf(BooAux0, False, (.Controls("Chkreg8").Value Or BooShowAll))  '11-r2
        Let VVarRequest(12) = IIf(BooAux0, False, (.Controls("Chkreg9").Value Or BooShowAll))  '12-Se
        Let VVarRequest(13) = .Controls("Chkreg10").Value                                      '13-Graph
        Let VVarRequest(14) = .Controls("Optreg0").Value                                       '14-SheetGraph
        Let VVarRequest(15) = VBooNewSheet                                                     '15-NewSheet
        Set VVarRequest(16) = GLORNG_CELL                                                      '13-RngBox
        Set GLORNG_CELL = Nothing
        Let VLngVecRngIndex = 16
      End With
      Let View_Core_preparations_regression = True
    
    Case "regressionselection"
      Dim StrY As String, StrX As String, StrExt As String
      Dim i As Long
      Set RngBox = Application.Selection
      If RngBox Is Nothing Then GoTo EH
      If RngBox.Areas.Count = 1 Then
        Let StrY = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(RngBox.Columns("A:A"), i, RValueType:=eqlResNumbers, RBooIncludeZeroes:=True)
        If RngBox.Columns.Count > 1 Then Let StrX = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(RngBox.Columns("B:B"), i, RValueType:=eqlResNumbers, RBooIncludeZeroes:=True)
        If RngBox.Columns.Count > 2 Then Let StrExt = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(RngBox.Columns("C:Z"), i, RValueType:=eqlResNumbers, RBooIncludeZeroes:=True)
      ElseIf RngBox.Areas.Count > 1 Then
        Let StrY = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(RngBox.Areas(1), i, RValueType:=eqlResNumbers, RBooIncludeZeroes:=True)
        If RngBox.Areas.Count > 1 Then Let StrX = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(RngBox.Areas(2), i, RValueType:=eqlResNumbers, RBooIncludeZeroes:=True)
        If RngBox.Areas.Count > 2 Then Let StrExt = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(RngBox.Areas(3), i, RValueType:=eqlResNumbers, RBooIncludeZeroes:=True)
      End If
      Set RngBox = Nothing
      
      Let BooShowAll = True 'ShowAll
      Let BooAux0 = BooShowAll 'JustRegression
      Let VBooNewSheet = Not GLOBOO_SELECCTIONCELL
      ReDim VVarRequest(16)
      Let VVarRequest(0) = StrX                         '0-X
      Let VVarRequest(1) = StrY                        '1-Y
      Let VVarRequest(2) = StrExt                       '2-E
      Let VVarRequest(3) = BooShowAll                       '3-ShowAll
      Let VVarRequest(4) = BooAux0                          '4-JustRegression
      Let VVarRequest(5) = BooShowAll                       '5-X2
      Let VVarRequest(6) = BooShowAll                       '6-Y2
      Let VVarRequest(7) = BooShowAll                       '7-XY
      Let VVarRequest(8) = BooShowAll                       '8-a
      Let VVarRequest(9) = BooShowAll                       '9-b
      Let VVarRequest(10) = BooShowAll                      '10-r
      Let VVarRequest(11) = BooShowAll                      '11-r2
      Let VVarRequest(12) = BooShowAll                      '12-Se
      Let VVarRequest(13) = False                           '13-Graph
      Let VVarRequest(14) = False                           '14-SheetGraph
      Let VVarRequest(15) = VBooNewSheet ' GLORNG_CELL Is Nothing           '15-NewSheet
      Set VVarRequest(16) = VRngBox 'GLORNG_CELL                      '13-RngBox
      Let VLngVecRngIndex = 16
      
      Let View_Core_preparations_regression = True
      
    Case Else: GoTo EH
  End Select

EH:
  Set RngBox = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function View_Core_preparations_balmet( _
  ByVal FRM As FRMBasics, _
  ByVal ControlId As String, _
  ByRef VVarRequestBools As Variant, _
  ByRef VVarRequest As Variant) As Boolean
'  ByVal RES As AppResources_current, _
 ByRef VBooNewSheet As Boolean, _
  ByRef VLngVecRngIndex As Long, _
  ByVal VRngBox As Range) As Boolean
  
  Dim RES As AppResources
  
  Dim RngBox As Range
  Dim StrFeeds As String, StrGrades As String, StrWeights As String, StrRecs As String, StrNames As String, StrUnits As String, StrGradeCx As String, StrRecCx As String, StrNameCx As String, StrUnitCx As String
  Dim VecNames() As Variant, VecUnits() As Variant, VLocalValues() As Variant
  Dim xlSplits As Variant
  Dim BooAux1 As Boolean, BooShowAll As Boolean
  Dim i As Long, j As Long, k As Long, l As Long, m As Long
  Dim StrAux0 As String, StrAux1 As String
  
  Const sFN As String = "VIEW::-ValuesPreparation_balmet"

  On Error GoTo EH
  Let View_Core_preparations_balmet = False
  
  ' Data required in BVarRequest: (0.DblFeed, 1.DblFeedVol, 2.VecGrad, 3.VecGradCx, 4.VecName, 5.VecUnit, 6.BytMethod, 7.Booleans, 8.EnuTypeFeed, 9.RngBox)
  Select Case ControlId
    Case "balmet"
      With FRM
        'Feed
        Let StrFeeds = IIf(.Controls("Txtbal0").Tag = "rng", GloStrRngA, .Controls("Txtbal0").Value)
        'Grades
        Let StrGrades = IIf(.Controls("Txtbal1").Tag = "rng", GloStrRngB, .Controls("Txtbal1").Value)
        'GradesComplex
        If .Controls("Chkbal6").Value Then Let StrGradeCx = IIf(.Controls("Txtbal8").Tag = "rng", GloStrRngC, .Controls("Txtbal8").Value)
        'Names & units
        Let j = IIf(.Controls("Lblbal2").Tag = VBA.vbNullString, 3, VBA.CLng(.Controls("Lblbal2").Tag) + 1)
        Let k = j - 2
        ReDim VecUnits(0)
        For i = 0 To k 'Units
          ReDim Preserve VecUnits(i)
          Let VecUnits(i) = .Controls("Cbobal" & i + 1).Value
        Next i
        ReDim VecNames(0)
        For i = 0 To j
          ReDim Preserve VecNames(i) 'Products Names
          If (i = j - 1) Then Let VecNames(i) = "C"
          If .Controls("Chkbal" & i + 1).Value And Not .Controls("Txtbal" & i + 3).Value = VBA.vbNullString Then
            If (i < j - 1) Then Let VecNames(i) = .Controls("Txtbal" & i + 3).Value
          Else
            If (i < j - 1) Then Let VecNames(i) = .Controls("Txtbal" & i + 3).Tag
          End If
        Next i
        ReDim Preserve VecNames(i - 1)
        If .Controls("Chkbal0").Value And Not .Controls("Txtbal2").Value = VBA.vbNullString Then 'Feed Name
          Let VecNames(i - 1) = .Controls("Txtbal2").Value
        Else
          Let VecNames(i - 1) = .Controls("Txtbal2").Tag
        End If

        'Set array of values
        'BOOLEANS: 0-SngMethod | 1-BooComplx | 2-BooShwAll | 3-BooPrcnts | 4-BooUnits | 5-BooFines | 6-BooRatio | 7-BooWOTCx | 8-BooWTNCx | 9-BooHeads | 10-BooHeadDwn | 11-BooNewSheet
        Let BooShowAll = .Controls("Chkbal7").Value
        ReDim VVarRequestBools(11)
        Let VVarRequestBools(0) = .Controls("Cbobal0").ListIndex                  '0-SngMethod  X
        Let VVarRequestBools(1) = .Controls("Chkbal6").Value                      '1-BooComplx  X
        Let VVarRequestBools(2) = BooShowAll                                      '2-BooShwAll  X
        Let VVarRequestBools(3) = (.Controls("Chkbal8").Value Or BooShowAll)      '3-BooPrcnts
        Let VVarRequestBools(4) = (.Controls("Chkbal9").Value Or BooShowAll)      '4-BooUnities
        Let VVarRequestBools(5) = (.Controls("Chkbal10").Value Or BooShowAll)     '5-BooFines
        Let VVarRequestBools(6) = (.Controls("Chkbal13").Value Or BooShowAll)     '6-BooRatio
        Let VVarRequestBools(7) = ((.Controls("Chkbal11").Value And .Controls("Chkbal6").Value) Or (BooShowAll And .Controls("Chkbal6").Value)) '7-BooWOTCx
        Let VVarRequestBools(8) = ((.Controls("Chkbal12").Value And .Controls("Chkbal6").Value) Or (BooShowAll And .Controls("Chkbal6").Value)) '8-BooWTNCx
        Let VVarRequestBools(9) = (BooShowAll Or .Controls("Chkbal14").Value)     '9-BooHeads
        Let VVarRequestBools(10) = .Controls("Tglbal0").Value                     '10-BooHeadDwn
        Let VVarRequestBools(11) = .Controls("Optbal0").Value                     '11-BooNewSheet X
  ' Data required in BVarRequest: (0.DblFeed, 1.DblFeedVol, 2.VecGrad, 3.VecGradCx, 4.VecName, 5.VecUnit, 6.BytMethod, 7.Booleans, 8.EnuTypeFeed, 9.RngBox)

'  Let DblFeed = VBA.CDbl(BVarRequest(0))
'  Let DblFeedVol = VBA.CDbl(BVarRequest(1))
'  Let EnuMethod = IIf(VBA.CByte(BVarRequest(6)) = 0, eqlBalConventional, IIf(VBA.CByte(BVarRequest(6)) = 1, eqlBalCramer, eqlBalInverseMatrix))
'  Let VecAux0 = BVarRequest(7)
'  Let BooAll = VBA.CBool(VecAux0(0))
'  Let BooPercents = VBA.CBool(VecAux0(1))
'  Let BooUnities = VBA.CBool(VecAux0(2))
'  Let BooFines = VBA.CBool(VecAux0(3))
'  Let BooGrams = VBA.CBool(VecAux0(4))
'  Let BooOT = VBA.CBool(VecAux0(5))
'  Let BooRatio = VBA.CBool(VecAux0(6))
'  Let BooHeads = VBA.CBool(VecAux0(7))
'  Let EnuTypeFeed = IIf(VBA.CByte(BVarRequest(8)) = 0, eqlBalSolids, IIf(VBA.CByte(BVarRequest(8)) = 1, eqlBalVolumes, eqlBalBoth))

        'VALUES: 0-StrFeeds | 1-StrGrades | 2-LngCols | 3-StrNames | 4-StrUnits | 5-StrNameCx | 6-StrCxGrades | 7-StrUnitsCx | 8-RngPlace | x 9-StrxlUnits
        ReDim VVarRequest(8)
        Let VVarRequest(0) = StrFeeds                    ' 0-StrFeeds
        Let VVarRequest(1) = StrGrades                   ' 1-StrGrades
        Let VVarRequest(2) = .Controls("Lblbal2").Tag    ' 2-LngCols
        Let VVarRequest(3) = VBA.Join(VecNames, ";")     ' 3-StrNames
        Let VVarRequest(4) = VBA.Join(VecUnits, ";")     ' 4-StrUnits
        Let VVarRequest(5) = .Controls("Cbobal6").Value  ' 5-StrNameCx
        Let VVarRequest(6) = StrGradeCx                 ' 6-StrCxGrades
        Let VVarRequest(7) = .Controls("Cbobal7").Value  ' 7-StrUnitsCx
        Set VVarRequest(8) = GLORNG_CELL                  ' 8-RngPlace
        'Let VVarRequest(9) = VBA.Join(VecUnits, ";")     ' 9-VecUnits
          
        Erase VecNames
        Erase VecUnits
        Let View_Core_preparations_balmet = True
      End With
    
    Case "balmetselection"
      Set RngBox = Selection
      If RngBox Is Nothing Then GoTo EH
      Set RES = New AppResources
      Let StrAux0 = RES.resArrayGetDataColonSeparatedStringFromVariants(RngBox, i, False, True, RBooRngRowsCols:=True) 'Numbers
      Let StrAux1 = RES.ResBalmetGetType(i - 1)  'Get the type of Balmet
      If StrAux1 = VBA.vbNullString Then GoTo EH
      Let BooAux1 = StrAux1 = "normal"
      Let l = RES.ResBalmetGetProductsBounds(i - 1, "cols") 'Get number of Balmet cols
    
      'Feed
      Let xlSplits = VBA.Split(StrAux0, ";")
      Let StrFeeds = VBA.CStr(xlSplits(0))
      
      'Grades
      If BooAux1 Then 'Normal Balmet
        Let StrGrades = VBA.Mid(StrAux0, VBA.Len(StrFeeds & ";"))
      Else 'Complex Balmet
        Let StrGrades = VBA.Mid(StrAux0, VBA.Len(StrFeeds & ";") + 1)
        Let j = RES.ResBalmetGetProductsBounds(i - 1, "full") 'Get number of Balmet grades without cx
        ReDim VLocalValues(0)
        For k = 1 To j
          ReDim Preserve VLocalValues(k - 1)
          Let VLocalValues(k - 1) = xlSplits(k)
        Next k
        Let StrGrades = VBA.Join(VLocalValues, ";") 'grades
        Erase VLocalValues
        ReDim VLocalValues(0)
        Let j = 0
        For k = k To UBound(xlSplits)
          ReDim Preserve VLocalValues(j)
          Let VLocalValues(j) = xlSplits(k)
          Let j = j + 1
        Next k
        Let StrGradeCx = VBA.Join(VLocalValues, ";") 'grades cx
      End If
      Erase VLocalValues
      Let xlSplits = Empty
      
      'Names & units
      Let StrAux0 = RES.resArrayGetDataColonSeparatedStringFromVariants(RngBox, j, True, RStrCustomRegex:="%", RBooRngRowsCols:=True) 'Get Strings: Names and Units
      Set RngBox = Nothing
      'Complex
      If Not BooAux1 Then
        Let StrNameCx = RES.ResBalmetGetComplexDatas(StrAux0, k) 'get Name cx
        Let StrNameCx = IIf(k = 0, "Ag", StrNameCx)
        Let StrUnitCx = RES.ResBalmetGetComplexDatas(StrAux0, k, True) 'get Unit cx
        Let StrUnitCx = IIf(k = 0, "DM", StrUnitCx)
      End If
      'Normal
      Let StrUnits = RES.ResBalmetGetUnits(StrAux0, k) 'Get Units
      Let j = l - 1
      'GoTo EH
      If k = 0 Then
        ReDim xlSplits(0)
        For i = 0 To j 'Cols
          ReDim Preserve xlSplits(i)
          Let xlSplits(i) = "%"
        Next i
        Let StrUnits = VBA.Join(xlSplits, ";")
      End If
      Let StrNames = RES.resArrayGetDataColonSeparatedStringFromVariants(StrAux0, k, True) 'Get Names
      If k = 0 Then
        ReDim xlSplits(0)
        For i = 0 To j 'Cols
          ReDim Preserve xlSplits(i)
          Let xlSplits(i) = "B" & i + 1
        Next i
        Let StrNames = VBA.Join(xlSplits, ";")
      End If
      Let StrNames = StrNames & ";C;A"
      
      'BOOLEANS: 0-SngMethod | 1-BooComplx | 2-BooShwAll | 3-BooPrcnts | 4-BooUnits | 5-BooFines | 6-BooRatio | 7-BooWOTCx | 8-BooWTNCx | 9-BooHeads | 10-BooHeadDwn | 11-BooNewSheet
      Let BooShowAll = Not GLOBOO_COMPACT
      ReDim VVarRequestBools(11)
      Let VVarRequestBools(0) = 2                                 '0-SngMethod  X
      Let VVarRequestBools(1) = Not BooAux1                       '1-BooComplx  X
      Let VVarRequestBools(2) = BooShowAll                        '2-BooShwAll  X
      Let VVarRequestBools(3) = (GLOBOO_COMPACT Or BooShowAll)     '3-BooPrcnts
      Let VVarRequestBools(4) = (False Or BooShowAll)             '4-BooUnities
      Let VVarRequestBools(5) = (GLOBOO_COMPACT Or BooShowAll)     '5-BooFines
      Let VVarRequestBools(6) = (False Or BooShowAll)             '6-BooRatio
      Let VVarRequestBools(7) = (False Or BooShowAll)             '7-BooWOTCx
      Let VVarRequestBools(8) = (False Or BooShowAll)             '8-BooWTNCx
      Let VVarRequestBools(9) = (GLOBOO_HEADS Or BooShowAll)       '9-BooHeads
      Let VVarRequestBools(10) = True                             '10-BooHeadDwn
      Let VVarRequestBools(11) = GLORNG_CELL Is Nothing            '11-BooNewSheet X

      'VALUES: 0-StrFeeds | 1-StrGrades | 2-LngCols | 3-StrNames | 4-StrUnits | 5-StrNameCx | 6-StrCxGrades | 7-StrUnitsCx | 8-RngPlace | x 9-StrxlUnits
      ReDim VVarRequest(8)
      Let VVarRequest(0) = StrFeeds     ' 0-StrFeeds
      Let VVarRequest(1) = StrGrades    ' 1-StrGrades
      Let VVarRequest(2) = l            ' 2-LngCols
      Let VVarRequest(3) = StrNames     ' 3-StrNames
      Let VVarRequest(4) = StrUnits     ' 4-StrUnits
      Let VVarRequest(5) = StrNameCx    ' 5-StrNameCx
      Let VVarRequest(6) = StrGradeCx  ' 6-StrCxGrades
      Let VVarRequest(7) = StrUnitCx    ' 7-StrUnitsCx
      Set VVarRequest(8) = GLORNG_CELL   ' 8-RngPlace
      
      Let View_Core_preparations_balmet = True
  
    Case "balmeteco"
      With FRM
        Set RES = New AppResources
        '[x] Feed
        Let xlSplits = RES.resArrayGetDataColonSeparatedFromString(.Controls("Txtbal0").Value, i, StrAbs:=True)
        Let StrFeeds = VBA.CStr(VBA.Abs(Application.WorksheetFunction.Sum(xlSplits)))
        '[ ] Weight(s)
        Let xlSplits = RES.resArrayGetDataColonSeparatedFromString(.Controls("Txtbal1").Value, m, StrAbs:=True)
        Let StrWeights = VBA.Join(xlSplits, ";")
        '[x] Grade(s)
        Let StrGrades = RES.resArrayGetDataColonSeparatedStringFromVariants(.Controls("Txtbal2").Value, i, RBooAbs:=True)
        '[ ] Recuperation(s)
        Let StrRecs = RES.resArrayGetDataColonSeparatedStringFromVariants(.Controls("Txtbal3").Value, i, RBooAbs:=True)
        If i = 0 Or Not i = m Then
          Let StrRecs = VBA.vbNullString
          For j = 0 To UBound(xlSplits)
            Let StrRecs = StrRecs & IIf(StrRecs = VBA.vbNullString, VBA.vbNullString, ";") & 50
          Next j
        End If
        '[x] Products Names
        Let StrNames = RES.resArrayGetDataColonSeparatedStringFromVariants(.Controls("Txtbal5").Value, i, True)
        If i = 0 Or Not i = m Then
          Let StrNames = VBA.vbNullString
          For j = 0 To UBound(xlSplits)
            Let StrNames = StrNames & IIf(StrUnits = VBA.vbNullString, VBA.vbNullString, ";") & "B" & j
          Next j
        End If
        '[x] Grades Units
        Let StrUnits = RES.resArrayGetDataColonSeparatedStringFromVariants(.Controls("Txtbal6").Value, i, True, RStrCustomRegex:="%")
        If i = 0 Or Not i = m Then
          Let StrUnits = VBA.vbNullString
          For j = 0 To UBound(xlSplits)
            Let StrUnits = StrUnits & IIf(StrUnits = VBA.vbNullString, VBA.vbNullString, ";") & "%"
          Next j
        End If
        Let xlSplits = Empty
        '[x] Alim label
        Let StrAux0 = RES.resArrayGetDataColonSeparatedStringFromVariants(.Controls("Txtbal4").Value, i, True)
        Let StrAux0 = ";C;" & IIf(StrAux0 = VBA.vbNullString, "A", StrAux0) '& "A"
        
        Let BooAux1 = .Controls("Chkbal4").Value
        If BooAux1 Then
          '[x] Complex Grade
          Let StrGradeCx = .Controls("Txtbal8").Value
          '[ ] Complex Recs
          Let StrRecCx = RES.resArrayGetDataColonSeparatedStringFromVariants(.Controls("Txtbal9").Value, i, RBooAbs:=True)
          Let StrRecCx = IIf(StrRecCx = VBA.vbNullString, "50", StrRecCx)
          '[x] Complex Name
          Let StrNameCx = .Controls("Cbobal3").Value
          '[x] Complex Unit
          Let StrUnitCx = .Controls("Cbobal4").Value
        End If
        Set RES = Nothing
        '1. Forming the request: (Feed, Weight, Grade, Rec, Unit, [Gradecx], [Unitcx], [Reccx])
        'BOOLEANS: 0-SngMethod | 1-BooComplx | 2-BooShwAll | 3-BooPrcnts | 4-BooUnits | 5-BooFines | 6-BooRatio | 7-BooWOTCx | 8-BooWTNCx | 9-BooHeads | 10-BooHeadDwn | 11-BooNewSheet
        Let BooShowAll = .Controls("Chkbal5").Value
        ReDim VVarRequestBools(11)
        Let VVarRequestBools(0) = .Controls("Cbobal1").ListIndex                  '0-SngMethod  X
        Let VVarRequestBools(1) = .Controls("Chkbal4").Value                      '1-BooComplx  X
        Let VVarRequestBools(2) = BooShowAll                                      '2-BooShwAll  X
        Let VVarRequestBools(3) = (.Controls("Chkbal6").Value Or BooShowAll)      '3-BooPrcnts
        Let VVarRequestBools(4) = (.Controls("Chkbal7").Value Or BooShowAll)      '4-BooUnities
        Let VVarRequestBools(5) = (.Controls("Chkbal8").Value Or BooShowAll)      '5-BooFines
        Let VVarRequestBools(6) = (.Controls("Chkbal11").Value Or BooShowAll)     '6-BooRatio
        Let VVarRequestBools(7) = ((.Controls("Chkbal10").Value And .Controls("Chkbal4").Value) Or (BooShowAll And .Controls("Chkbal4").Value)) '7-BooWOTCx
        Let VVarRequestBools(8) = ((.Controls("Chkbal11").Value And .Controls("Chkbal4").Value) Or (BooShowAll And .Controls("Chkbal4").Value)) '8-BooWTNCx
        Let VVarRequestBools(9) = (BooShowAll Or .Controls("Chkbal12").Value)     '9-BooHeads
        Let VVarRequestBools(10) = True                                           '10-BooHeadDwn
        Let VVarRequestBools(11) = .Controls("Optbal0").Value                     '11-BooNewSheet X

        'VALUES: 0-StrFeeds | 1-StrGrades | 2-LngCols | 3-StrNames | 4-StrUnits | 5-StrNameCx | 6-StrCxGrades | 7-StrUnitsCx | 8-RngPlace | x 9-StrxlUnits
        ReDim VVarRequest(11)
        Let VVarRequest(0) = StrFeeds                    ' 0-StrFeeds
        Let VVarRequest(1) = StrGrades                   ' 1-StrGrades    Eco[Grades|HeadsGrades]
        Let VVarRequest(2) = m                           ' 2-LngCols
        Let VVarRequest(3) = StrNames & StrAux0            ' 3-StrNames
        Let VVarRequest(4) = StrUnits                    ' 4-StrUnits
        Let VVarRequest(5) = StrNameCx                   ' 5-StrNameCx    Eco|Eng
        Let VVarRequest(6) = StrGradeCx                  ' 6-StrCxGrades  Eco|Eng
        Let VVarRequest(7) = StrUnitCx                   ' 7-StrUnitsCx   Eco|Eng
        Set VVarRequest(8) = GLORNG_CELL                  ' 8-RngPlace
        Let VVarRequest(9) = StrWeights                  ' 9-StrWeights   Eco
        Let VVarRequest(10) = StrRecs                    ' 10-StrRecs     Eco
        Let VVarRequest(11) = StrRecCx                   ' 11-StrRecCx    Eco
      End With
      Let View_Core_preparations_balmet = True
    
    Case Else: GoTo EH
  End Select

EH:
  Erase VecNames
  Erase VecUnits
  Let xlSplits = Empty
  Erase VLocalValues
  Set RngBox = Nothing
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)
End Function

Private Function View_Core_preparations_balmet_OOLD( _
  ByVal RES As AppResources_current, _
  ByVal FRM As FRMBasics, _
  ByVal ControlId As String, _
  ByRef VVarRequestBools As Variant, _
  ByRef VVarRequest As Variant) As Boolean
'' ByRef VBooNewSheet As Boolean, _
'  ByRef VLngVecRngIndex As Long, _
'  ByVal VRngBox As Range) As Boolean
'
'  Dim RES As AppResources
'
'  Dim RngBox As Range
'  Dim StrFeeds As String, StrGrades As String, StrWeights As String, StrRecs As String, StrNames As String, StrUnits As String, StrGradeCx As String, StrRecCx As String, StrNameCx As String, StrUnitCx As String
'  Dim VecNames() As Variant, VecUnits() As Variant, VLocalValues() As Variant
'  Dim xlSplits As Variant
'  Dim BooAux1 As Boolean, BooShowAll As Boolean
'  Dim i As Long, j As Long, k As Long, l As Long, m As Long
'  Dim StrAux0 As String, StrAux1 As String
'
'  Const sFN As String = "VIEW::-ValuesPreparation_balmet"
'
'  On Error GoTo EH
'  Let View_Core_preparations_balmet = False
'
'  ' Data required in BVarRequest: (0.DblFeed, 1.DblFeedVol, 2.VecGrad, 3.VecGradCx, 4.VecName, 5.VecUnit, 6.BytMethod, 7.Booleans, 8.EnuTypeFeed, 9.RngBox)
'  Select Case ControlId
'    Case "balmet"
'      With FRM
'        'Feed
'        Let StrFeeds = IIf(.Controls("Txtbal0").Tag = "rng", GloStrRngA, .Controls("Txtbal0").Value)
'        'Grades
'        Let StrGrades = IIf(.Controls("Txtbal1").Tag = "rng", GloStrRngB, .Controls("Txtbal1").Value)
'        'GradesComplex
'        If .Controls("Chkbal6").Value Then Let StrGradeCx = IIf(.Controls("Txtbal8").Tag = "rng", GloStrRngC, .Controls("Txtbal8").Value)
'        'Names & units
'        Let j = IIf(.Controls("Lblbal2").Tag = VBA.vbNullString, 3, VBA.CLng(.Controls("Lblbal2").Tag) + 1)
'        Let k = j - 2
'        ReDim VecUnits(0)
'        For i = 0 To k 'Units
'          ReDim Preserve VecUnits(i)
'          Let VecUnits(i) = .Controls("Cbobal" & i + 1).Value
'        Next i
'        ReDim VecNames(0)
'        For i = 0 To j
'          ReDim Preserve VecNames(i) 'Products Names
'          If (i = j - 1) Then Let VecNames(i) = "C"
'          If .Controls("Chkbal" & i + 1).Value And Not .Controls("Txtbal" & i + 3).Value = VBA.vbNullString Then
'            If (i < j - 1) Then Let VecNames(i) = .Controls("Txtbal" & i + 3).Value
'          Else
'            If (i < j - 1) Then Let VecNames(i) = .Controls("Txtbal" & i + 3).Tag
'          End If
'        Next i
'        ReDim Preserve VecNames(i - 1)
'        If .Controls("Chkbal0").Value And Not .Controls("Txtbal2").Value = VBA.vbNullString Then 'Feed Name
'          Let VecNames(i - 1) = .Controls("Txtbal2").Value
'        Else
'          Let VecNames(i - 1) = .Controls("Txtbal2").Tag
'        End If
'
'        'Set array of values
'        'BOOLEANS: 0-SngMethod | 1-BooComplx | 2-BooShwAll | 3-BooPrcnts | 4-BooUnits | 5-BooFines | 6-BooRatio | 7-BooWOTCx | 8-BooWTNCx | 9-BooHeads | 10-BooHeadDwn | 11-BooNewSheet
'        Let BooShowAll = .Controls("Chkbal7").Value
'        ReDim VVarRequestBools(11)
'        Let VVarRequestBools(0) = .Controls("Cbobal0").ListIndex                  '0-SngMethod  X
'        Let VVarRequestBools(1) = .Controls("Chkbal6").Value                      '1-BooComplx  X
'        Let VVarRequestBools(2) = BooShowAll                                      '2-BooShwAll  X
'        Let VVarRequestBools(3) = (.Controls("Chkbal8").Value Or BooShowAll)      '3-BooPrcnts
'        Let VVarRequestBools(4) = (.Controls("Chkbal9").Value Or BooShowAll)      '4-BooUnities
'        Let VVarRequestBools(5) = (.Controls("Chkbal10").Value Or BooShowAll)     '5-BooFines
'        Let VVarRequestBools(6) = (.Controls("Chkbal13").Value Or BooShowAll)     '6-BooRatio
'        Let VVarRequestBools(7) = ((.Controls("Chkbal11").Value And .Controls("Chkbal6").Value) Or (BooShowAll And .Controls("Chkbal6").Value)) '7-BooWOTCx
'        Let VVarRequestBools(8) = ((.Controls("Chkbal12").Value And .Controls("Chkbal6").Value) Or (BooShowAll And .Controls("Chkbal6").Value)) '8-BooWTNCx
'        Let VVarRequestBools(9) = (BooShowAll Or .Controls("Chkbal14").Value)     '9-BooHeads
'        Let VVarRequestBools(10) = .Controls("Tglbal0").Value                     '10-BooHeadDwn
'        Let VVarRequestBools(11) = .Controls("Optbal0").Value                     '11-BooNewSheet X
'
''  Let DblFeed = VBA.CDbl(BVarRequest(0))
''  Let DblFeedVol = VBA.CDbl(BVarRequest(1))
''  Let EnuMethod = IIf(VBA.CByte(BVarRequest(6)) = 0, eqlBalConventional, IIf(VBA.CByte(BVarRequest(6)) = 1, eqlBalCramer, eqlBalInverseMatrix))
''  Let VecAux0 = BVarRequest(7)
''  Let BooAll = VBA.CBool(VecAux0(0))
''  Let BooPercents = VBA.CBool(VecAux0(1))
''  Let BooUnities = VBA.CBool(VecAux0(2))
''  Let BooFines = VBA.CBool(VecAux0(3))
''  Let BooGrams = VBA.CBool(VecAux0(4))
''  Let BooOT = VBA.CBool(VecAux0(5))
''  Let BooRatio = VBA.CBool(VecAux0(6))
''  Let BooHeads = VBA.CBool(VecAux0(7))
''  Let EnuTypeFeed = IIf(VBA.CByte(BVarRequest(8)) = 0, eqlBalSolids, IIf(VBA.CByte(BVarRequest(8)) = 1, eqlBalVolumes, eqlBalBoth))
'
'        'VALUES: 0-StrFeeds | 1-StrGrades | 2-LngCols | 3-StrNames | 4-StrUnits | 5-StrNameCx | 6-StrCxGrades | 7-StrUnitsCx | 8-RngPlace | x 9-StrxlUnits
'        ReDim VVarRequest(8)
'        Let VVarRequest(0) = StrFeeds                    ' 0-StrFeeds
'        Let VVarRequest(1) = StrGrades                   ' 1-StrGrades
'        Let VVarRequest(2) = .Controls("Lblbal2").Tag    ' 2-LngCols
'        Let VVarRequest(3) = VBA.Join(VecNames, ";")     ' 3-StrNames
'        Let VVarRequest(4) = VBA.Join(VecUnits, ";")     ' 4-StrUnits
'        Let VVarRequest(5) = .Controls("Cbobal6").Value  ' 5-StrNameCx
'        Let VVarRequest(6) = StrGradeCx                 ' 6-StrCxGrades
'        Let VVarRequest(7) = .Controls("Cbobal7").Value  ' 7-StrUnitsCx
'        Set VVarRequest(8) = GLORNG_CELL                  ' 8-RngPlace
'        'Let VVarRequest(9) = VBA.Join(VecUnits, ";")     ' 9-VecUnits
'
'        Erase VecNames
'        Erase VecUnits
'        Let View_Core_preparations_balmet = True
'      End With
'
'    Case "balmetselection"
'      Set RngBox = Selection
'      If RngBox Is Nothing Then GoTo EH
'      Set RES = New AppResources
'      Let StrAux0 = RES.resArrayGetDataColonSeparatedStringFromVariants(RngBox, i, False, True, RBooRngRowsCols:=True) 'Numbers
'      Let StrAux1 = RES.ResBalmetGetType(i - 1)  'Get the type of Balmet
'      If StrAux1 = VBA.vbNullString Then GoTo EH
'      Let BooAux1 = StrAux1 = "normal"
'      Let l = RES.ResBalmetGetProductsBounds(i - 1, "cols") 'Get number of Balmet cols
'
'      'Feed
'      Let xlSplits = VBA.Split(StrAux0, ";")
'      Let StrFeeds = VBA.CStr(xlSplits(0))
'
'      'Grades
'      If BooAux1 Then 'Normal Balmet
'        Let StrGrades = VBA.Mid(StrAux0, VBA.Len(StrFeeds & ";"))
'      Else 'Complex Balmet
'        Let StrGrades = VBA.Mid(StrAux0, VBA.Len(StrFeeds & ";") + 1)
'        Let j = RES.ResBalmetGetProductsBounds(i - 1, "full") 'Get number of Balmet grades without cx
'        ReDim VLocalValues(0)
'        For k = 1 To j
'          ReDim Preserve VLocalValues(k - 1)
'          Let VLocalValues(k - 1) = xlSplits(k)
'        Next k
'        Let StrGrades = VBA.Join(VLocalValues, ";") 'grades
'        Erase VLocalValues
'        ReDim VLocalValues(0)
'        Let j = 0
'        For k = k To UBound(xlSplits)
'          ReDim Preserve VLocalValues(j)
'          Let VLocalValues(j) = xlSplits(k)
'          Let j = j + 1
'        Next k
'        Let StrGradeCx = VBA.Join(VLocalValues, ";") 'grades cx
'      End If
'      Erase VLocalValues
'      Let xlSplits = Empty
'
'      'Names & units
'      Let StrAux0 = RES.resArrayGetDataColonSeparatedStringFromVariants(RngBox, j, True, RStrCustomRegex:="%", RBooRngRowsCols:=True) 'Get Strings: Names and Units
'      Set RngBox = Nothing
'      'Complex
'      If Not BooAux1 Then
'        Let StrNameCx = RES.ResBalmetGetComplexDatas(StrAux0, k) 'get Name cx
'        Let StrNameCx = IIf(k = 0, "Ag", StrNameCx)
'        Let StrUnitCx = RES.ResBalmetGetComplexDatas(StrAux0, k, True) 'get Unit cx
'        Let StrUnitCx = IIf(k = 0, "DM", StrUnitCx)
'      End If
'      'Normal
'      Let StrUnits = RES.ResBalmetGetUnits(StrAux0, k) 'Get Units
'      Let j = l - 1
'      'GoTo EH
'      If k = 0 Then
'        ReDim xlSplits(0)
'        For i = 0 To j 'Cols
'          ReDim Preserve xlSplits(i)
'          Let xlSplits(i) = "%"
'        Next i
'        Let StrUnits = VBA.Join(xlSplits, ";")
'      End If
'      Let StrNames = RES.resArrayGetDataColonSeparatedStringFromVariants(StrAux0, k, True) 'Get Names
'      If k = 0 Then
'        ReDim xlSplits(0)
'        For i = 0 To j 'Cols
'          ReDim Preserve xlSplits(i)
'          Let xlSplits(i) = "B" & i + 1
'        Next i
'        Let StrNames = VBA.Join(xlSplits, ";")
'      End If
'      Let StrNames = StrNames & ";C;A"
'
'      'BOOLEANS: 0-SngMethod | 1-BooComplx | 2-BooShwAll | 3-BooPrcnts | 4-BooUnits | 5-BooFines | 6-BooRatio | 7-BooWOTCx | 8-BooWTNCx | 9-BooHeads | 10-BooHeadDwn | 11-BooNewSheet
'      Let BooShowAll = Not GLOBOO_COMPACT
'      ReDim VVarRequestBools(11)
'      Let VVarRequestBools(0) = 2                                 '0-SngMethod  X
'      Let VVarRequestBools(1) = Not BooAux1                       '1-BooComplx  X
'      Let VVarRequestBools(2) = BooShowAll                        '2-BooShwAll  X
'      Let VVarRequestBools(3) = (GLOBOO_COMPACT Or BooShowAll)     '3-BooPrcnts
'      Let VVarRequestBools(4) = (False Or BooShowAll)             '4-BooUnities
'      Let VVarRequestBools(5) = (GLOBOO_COMPACT Or BooShowAll)     '5-BooFines
'      Let VVarRequestBools(6) = (False Or BooShowAll)             '6-BooRatio
'      Let VVarRequestBools(7) = (False Or BooShowAll)             '7-BooWOTCx
'      Let VVarRequestBools(8) = (False Or BooShowAll)             '8-BooWTNCx
'      Let VVarRequestBools(9) = (GLOBOO_HEADS Or BooShowAll)       '9-BooHeads
'      Let VVarRequestBools(10) = True                             '10-BooHeadDwn
'      Let VVarRequestBools(11) = GLORNG_CELL Is Nothing            '11-BooNewSheet X
'
'      'VALUES: 0-StrFeeds | 1-StrGrades | 2-LngCols | 3-StrNames | 4-StrUnits | 5-StrNameCx | 6-StrCxGrades | 7-StrUnitsCx | 8-RngPlace | x 9-StrxlUnits
'      ReDim VVarRequest(8)
'      Let VVarRequest(0) = StrFeeds     ' 0-StrFeeds
'      Let VVarRequest(1) = StrGrades    ' 1-StrGrades
'      Let VVarRequest(2) = l            ' 2-LngCols
'      Let VVarRequest(3) = StrNames     ' 3-StrNames
'      Let VVarRequest(4) = StrUnits     ' 4-StrUnits
'      Let VVarRequest(5) = StrNameCx    ' 5-StrNameCx
'      Let VVarRequest(6) = StrGradeCx  ' 6-StrCxGrades
'      Let VVarRequest(7) = StrUnitCx    ' 7-StrUnitsCx
'      Set VVarRequest(8) = GLORNG_CELL   ' 8-RngPlace
'
'      Let View_Core_preparations_balmet = True
'
'    Case "balmeteco"
'      With FRM
'        Set RES = New AppResources
'        '[x] Feed
'        Let xlSplits = RES.resArrayGetDataColonSeparatedFromString(.Controls("Txtbal0").Value, i, StrAbs:=True)
'        Let StrFeeds = VBA.CStr(VBA.Abs(Application.WorksheetFunction.Sum(xlSplits)))
'        '[ ] Weight(s)
'        Let xlSplits = RES.resArrayGetDataColonSeparatedFromString(.Controls("Txtbal1").Value, m, StrAbs:=True)
'        Let StrWeights = VBA.Join(xlSplits, ";")
'        '[x] Grade(s)
'        Let StrGrades = RES.resArrayGetDataColonSeparatedStringFromVariants(.Controls("Txtbal2").Value, i, RBooAbs:=True)
'        '[ ] Recuperation(s)
'        Let StrRecs = RES.resArrayGetDataColonSeparatedStringFromVariants(.Controls("Txtbal3").Value, i, RBooAbs:=True)
'        If i = 0 Or Not i = m Then
'          Let StrRecs = VBA.vbNullString
'          For j = 0 To UBound(xlSplits)
'            Let StrRecs = StrRecs & IIf(StrRecs = VBA.vbNullString, VBA.vbNullString, ";") & 50
'          Next j
'        End If
'        '[x] Products Names
'        Let StrNames = RES.resArrayGetDataColonSeparatedStringFromVariants(.Controls("Txtbal5").Value, i, True)
'        If i = 0 Or Not i = m Then
'          Let StrNames = VBA.vbNullString
'          For j = 0 To UBound(xlSplits)
'            Let StrNames = StrNames & IIf(StrUnits = VBA.vbNullString, VBA.vbNullString, ";") & "B" & j
'          Next j
'        End If
'        '[x] Grades Units
'        Let StrUnits = RES.resArrayGetDataColonSeparatedStringFromVariants(.Controls("Txtbal6").Value, i, True, RStrCustomRegex:="%")
'        If i = 0 Or Not i = m Then
'          Let StrUnits = VBA.vbNullString
'          For j = 0 To UBound(xlSplits)
'            Let StrUnits = StrUnits & IIf(StrUnits = VBA.vbNullString, VBA.vbNullString, ";") & "%"
'          Next j
'        End If
'        Let xlSplits = Empty
'        '[x] Alim label
'        Let StrAux0 = RES.resArrayGetDataColonSeparatedStringFromVariants(.Controls("Txtbal4").Value, i, True)
'        Let StrAux0 = ";C;" & IIf(StrAux0 = VBA.vbNullString, "A", StrAux0) '& "A"
'
'        Let BooAux1 = .Controls("Chkbal4").Value
'        If BooAux1 Then
'          '[x] Complex Grade
'          Let StrGradeCx = .Controls("Txtbal8").Value
'          '[ ] Complex Recs
'          Let StrRecCx = RES.resArrayGetDataColonSeparatedStringFromVariants(.Controls("Txtbal9").Value, i, RBooAbs:=True)
'          Let StrRecCx = IIf(StrRecCx = VBA.vbNullString, "50", StrRecCx)
'          '[x] Complex Name
'          Let StrNameCx = .Controls("Cbobal3").Value
'          '[x] Complex Unit
'          Let StrUnitCx = .Controls("Cbobal4").Value
'        End If
'        Set RES = Nothing
'        '1. Forming the request: (Feed, Weight, Grade, Rec, Unit, [Gradecx], [Unitcx], [Reccx])
'        'BOOLEANS: 0-SngMethod | 1-BooComplx | 2-BooShwAll | 3-BooPrcnts | 4-BooUnits | 5-BooFines | 6-BooRatio | 7-BooWOTCx | 8-BooWTNCx | 9-BooHeads | 10-BooHeadDwn | 11-BooNewSheet
'        Let BooShowAll = .Controls("Chkbal5").Value
'        ReDim VVarRequestBools(11)
'        Let VVarRequestBools(0) = .Controls("Cbobal1").ListIndex                  '0-SngMethod  X
'        Let VVarRequestBools(1) = .Controls("Chkbal4").Value                      '1-BooComplx  X
'        Let VVarRequestBools(2) = BooShowAll                                      '2-BooShwAll  X
'        Let VVarRequestBools(3) = (.Controls("Chkbal6").Value Or BooShowAll)      '3-BooPrcnts
'        Let VVarRequestBools(4) = (.Controls("Chkbal7").Value Or BooShowAll)      '4-BooUnities
'        Let VVarRequestBools(5) = (.Controls("Chkbal8").Value Or BooShowAll)      '5-BooFines
'        Let VVarRequestBools(6) = (.Controls("Chkbal11").Value Or BooShowAll)     '6-BooRatio
'        Let VVarRequestBools(7) = ((.Controls("Chkbal10").Value And .Controls("Chkbal4").Value) Or (BooShowAll And .Controls("Chkbal4").Value)) '7-BooWOTCx
'        Let VVarRequestBools(8) = ((.Controls("Chkbal11").Value And .Controls("Chkbal4").Value) Or (BooShowAll And .Controls("Chkbal4").Value)) '8-BooWTNCx
'        Let VVarRequestBools(9) = (BooShowAll Or .Controls("Chkbal12").Value)     '9-BooHeads
'        Let VVarRequestBools(10) = True                                           '10-BooHeadDwn
'        Let VVarRequestBools(11) = .Controls("Optbal0").Value                     '11-BooNewSheet X
'
'        'VALUES: 0-StrFeeds | 1-StrGrades | 2-LngCols | 3-StrNames | 4-StrUnits | 5-StrNameCx | 6-StrCxGrades | 7-StrUnitsCx | 8-RngPlace | x 9-StrxlUnits
'        ReDim VVarRequest(11)
'        Let VVarRequest(0) = StrFeeds                    ' 0-StrFeeds
'        Let VVarRequest(1) = StrGrades                   ' 1-StrGrades    Eco[Grades|HeadsGrades]
'        Let VVarRequest(2) = m                           ' 2-LngCols
'        Let VVarRequest(3) = StrNames & StrAux0            ' 3-StrNames
'        Let VVarRequest(4) = StrUnits                    ' 4-StrUnits
'        Let VVarRequest(5) = StrNameCx                   ' 5-StrNameCx    Eco|Eng
'        Let VVarRequest(6) = StrGradeCx                  ' 6-StrCxGrades  Eco|Eng
'        Let VVarRequest(7) = StrUnitCx                   ' 7-StrUnitsCx   Eco|Eng
'        Set VVarRequest(8) = GLORNG_CELL                  ' 8-RngPlace
'        Let VVarRequest(9) = StrWeights                  ' 9-StrWeights   Eco
'        Let VVarRequest(10) = StrRecs                    ' 10-StrRecs     Eco
'        Let VVarRequest(11) = StrRecCx                   ' 11-StrRecCx    Eco
'      End With
'      Let View_Core_preparations_balmet = True
'
'    Case Else: GoTo EH
'  End Select
'
'EH:
'  Erase VecNames
'  Erase VecUnits
'  Let xlSplits = Empty
'  Erase VLocalValues
'  Set RngBox = Nothing
'  Set RES = Nothing
'  Call EHGLOBAL.ErrorHandlerRaise(sFN)
End Function

Private Function View_Core_preparations_projection( _
  ByVal FRM As FRMBasics, _
  ByVal ControlId As String, _
  ByRef VVarRequestBools As Variant, _
  ByRef VVarRequest As Variant) As Boolean
  
  Dim REGEX As AppResRegEx
  
  Const sFN As String = "VIEW::-ValuesPreparation_projection"

  On Error GoTo EH
  Let View_Core_preparations_projection = False
  
  ' Request
  ' 0-DblIo      5-VecW[]           9-StrDivise         11-BooVAN           16-Rng
  ' 1-DblVR      6-VecWName[]       10-StrTimeUnit      12-BooTIR
  ' 2-DblEgr     7-VecWUnitCtz[]                        13-BooGraph
  ' 3-DblT       8-VecCtz[]                             14-BooGraphSheet
  ' 4-Dbl%                                              15-BooNewSheet
  'Inversion
  With FRM
    Set REGEX = New AppResRegEx
    ReDim VVarRequest(16)
    Let VVarRequest(0) = VBA.CDbl(.Controls("Txtpro0").Value) '0-DblIo
    Let VVarRequest(1) = VBA.CDbl(IIf(Not REGEX.isNumberReg(.Controls("Txtpro3").Value), 0, .Controls("Txtpro3").Value)) '1-DblVR
    Let VVarRequest(2) = VBA.CDbl(VBA.Replace(.Controls("Txtpro6").Value, "%", VBA.vbNullString)) '2-DblEgr
    Let VVarRequest(3) = VBA.CDbl(IIf(Not REGEX.isNumberReg(.Controls("Txtpro7").Value), 1, .Controls("Txtpro7").Value)) '3-DblT
    Let VVarRequest(4) = VBA.CDbl(VBA.Replace(.Controls("Txtpro8").Value, "%", VBA.vbNullString)) '4-Dbl%
    
    Let VVarRequest(5) = .Controls("Txtpro1").Value '5-VecW[]
    Let VVarRequest(6) = .Controls("Txtpro4").Value '6-VecWName[]
    Let VVarRequest(7) = .Controls("Txtpro5").Value '7-VecWUnitCtz[]
    Let VVarRequest(8) = .Controls("Txtpro2").Value '8-VecCtz[]
    
    Let VVarRequest(9) = .Controls("Cbopro0").Value '9-StrDivise
    Let VVarRequest(10) = .Controls("Cbopro3").Value '10-StrTimeUnit
    
    Let VVarRequest(11) = .Controls("Chkpro0").Value '11-BooVAN
    Let VVarRequest(12) = .Controls("Chkpro1").Value '12-BooTIR
    Let VVarRequest(13) = .Controls("Chkpro2").Value '13-BooGraph
    Let VVarRequest(14) = .Controls("Optpro0").Value '14-BooSelfSheet
    Let VVarRequest(15) = .Controls("Optpro2").Value '15-BooNewSheet
    Set VVarRequest(16) = GLORNG_CELL '13-RngBox
  End With
  Set REGEX = Nothing
  Let View_Core_preparations_projection = True

EH:
  Set REGEX = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)
End Function

Private Function View_Core_preparations_oreinvoice( _
  ByVal FRM As FRMBasics, _
  ByVal ControlId As String, _
  ByRef VVarRequestBools As Variant, _
  ByRef VVarRequest As Variant) As Boolean
  
  Dim RES As AppResources
  Dim REGEX As AppResRegEx
  
  Dim xlSplits As Variant
  Dim aControl() As Variant
  Dim i As Long, j As Long
  Dim StrAux0 As String
  
  Const sFN As String = "VIEW::-ValuesPreparation_oreinvoice"

  On Error GoTo EH
  Let View_Core_preparations_oreinvoice = False
  
  ' Request
  ' 0-StrPlace      5-StrProNames[]      10-BooNewSheet
  ' 1-StrDate       6-StrProType[]       11-RngBox
  ' 2-StrTC         7-BooAccurate
  ' 3-BooME         8-BooPage
  ' 4-BooVarious    9-BooPrintTwo
  Set REGEX = New AppResRegEx
  With FRM
    ReDim VVarRequest(11)
    Let VVarRequest(0) = IIf(REGEX.isEmptyStringReg(.Controls("Txtpur0").Value), "Unknown Place", .Controls("Txtpur0").Value) '0-StrPlace
    Let VVarRequest(1) = VBA.CStr(IIf(REGEX.isDateReg(.Controls("Txtpur1").Value), .Controls("Txtpur1").Value, VBA.Format$(VBA.Date(), "dd/mm/yy"))) '1-StrDate
    Let VVarRequest(2) = VBA.CStr(IIf(REGEX.isNumberReg(.Controls("Txtpur2").Value), .Controls("Txtpur2").Value, 1))  '2-StrTC
    Let VVarRequest(3) = .Controls("Chkpur0").Value '3-BooME
    Let VVarRequest(4) = .Controls("Chkpur1").Value '4-BooVarious
    
    If .Controls("Chkpur1").Value Then 'If various providers
      Set RES = New AppResources
      Let xlSplits = RES.resArrayGetDataArrayFromVariants(.Controls("Txtpur3").Value, i, True) 'names
      Let xlSplits = RES.resArrayGetDataArrayFromVariants(.Controls("Txtpur4").Value, j, True)
      Set RES = Nothing
      If i = j Then
        Let StrAux0 = .Controls("Txtpur3").Value
      Else
        For i = 0 To UBound(xlSplits)
          Let xlSplits(i) = "Unknown Provider " & i
        Next i
        Let StrAux0 = VBA.Join(xlSplits, ";")
      End If
      Let xlSplits = Empty
      Let VVarRequest(5) = StrAux0 '5-StrProNames[]
      Let VVarRequest(6) = .Controls("Txtpur4").Value '6-StrProType[]
    Else
      Let VVarRequest(5) = IIf(REGEX.isEmptyStringReg(.Controls("Txtpur3").Value), "Unknown Provider", .Controls("Txtpur3").Value) '5-StrProNames[]
      Let VVarRequest(6) = .Controls("Cbopur0").Value '6-StrProType[]
    End If
    Let VVarRequest(7) = .Controls("Chkpur3").Value '7-BooAccurate
    Let VVarRequest(8) = .Controls("Chkpur4").Value '8-BooPage
    Let VVarRequest(9) = .Controls("Chkpur5").Value '9-BooPrintTwo
    Let VVarRequest(10) = .Controls("Optpur0").Value '10-BooNewSheet
    Set VVarRequest(11) = GLORNG_CELL '11-RngBox
  End With
  Set REGEX = Nothing
  
  Let View_Core_preparations_oreinvoice = True

EH:
  Let xlSplits = Empty
  Set REGEX = Nothing
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)
End Function

Private Function View_Core_preparations_concentrateinvoice( _
  ByVal FRM As FRMBasics, _
  ByVal ControlId As String, _
  ByRef VVarRequestBools As Variant, _
  ByRef VVarRequest As Variant) As Boolean
    
  Const sFN As String = "VIEW::-ValuesPreparation_concentrateinvoice"

  On Error GoTo EH
  Let View_Core_preparations_concentrateinvoice = False
  
  ' Request
  ' 0-BooPage
  ' 1-BooPrintTwo
  ' 2-BooNewSheet
  ' 3-RngBox
  With FRM
    ReDim VVarRequest(3)
    Let VVarRequest(0) = .Controls("Chkcnc2").Value '8-BooPage
    Let VVarRequest(1) = .Controls("Chkcnc3").Value '9-BooPrintTwo
    Let VVarRequest(2) = .Controls("Optcnc0").Value '10-BooNewSheet
    Set VVarRequest(3) = GLORNG_CELL '11-RngBox
  End With
  
  Let View_Core_preparations_concentrateinvoice = True

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)
End Function


'Executor method
Private Function View_Core_execution(ByVal ControlId As String, ByVal FRM As FRMBasics, Optional VRngBox As Range = Nothing) As Boolean
  
  Dim RES As AppResources_current
  
  Dim Wbk As Workbook
  Dim VecRes As Variant, VecReq As Variant
  Dim StrSheetName As String
  Dim BooCancel As Boolean, BooNewSheet As Boolean, BooDeleteNewSheet As Boolean
  Dim LngVecRngIndex As Long

  Const sFN As String = "VIEW::-EXE"
  
  On Error GoTo EH
  Let Application.ScreenUpdating = False
  'Call VBA.MsgBox("Ready to exe!", vbInformation, STRAPPTITLE): GoTo EH
  Set RES = New AppResources_current
  Set RES.ErrorHandler = EHGLOBAL
  If Not View_Core_preparations(RES, FRM, ControlId, VecReq, BooNewSheet, LngVecRngIndex, VRngBox) Then GoTo EH
  Set RES = Nothing
    
  ' Event implementation -----------------------------------------------------------------
  
  ' Get range from selected cell or by creating a new sheet
  Call View_assets_get_selected_or_new_range(ControlId, VecReq)
  
  ' Raise the custom event
  RaiseEvent RequestEvent(ControlId, VecReq, VecRes, View_Core_execution) 'BooCancel)
  
  ' Process the received data
  If View_Core_execution Then
    
    ' Print results on spreadsheet
    Call VBA.MsgBox("Success!!!", vbInformation, STRAPPTITLE)
    Dim VWP As ViewPrint_current
    Set VWP = New ViewPrint_current
    Set VWP.ErrorHandler = EHGLOBAL
    Let View_Core_execution = VWP.ViewPrint(ControlId, VecReq, VecRes)
    Set VWP = Nothing
    If View_Core_execution Then Let BooNewSheet = False
    
  End If
  
  If Not View_Core_execution Then
    
    ' Delete new sheet created
    If BooNewSheet Then  'LngVecRngIndex
      With Application
        Let .DisplayAlerts = False
        Set Wbk = VecReq(LngVecRngIndex).Parent.Parent
        Let StrSheetName = VecReq(LngVecRngIndex).Parent.Name
        Call Wbk.Worksheets(StrSheetName).Delete
        Set Wbk = Nothing
        Let BooNewSheet = False
        Let .DisplayAlerts = True
      End With
    End If
    
    ' Other tasks
    Let View_Core_execution = View_Core_error_handler(ControlId, FRM)
  
  End If
  
  ' Event implementation -----------------------------------------------------------------
  
  
  
'  Set OBJCTR = New Controller 'Creating a new controller
'  With OBJCTR
'    Set .ErrorHandler = EHGLOBAL 'Setting the Error local Handler Object to Local Error Handler of new controller
'
'      Case "balmet": Let View_Core_execution = OBJCTR.Balmet(ControlId, VecBOOLEANS, VecREQUEST)
'
'      Case "balmetselection": Let View_Core_execution = OBJCTR.Balmet("balmet", VecBOOLEANS, VecREQUEST)
'
'      Case "balmeteco"
'        Let ControlId = IIf(FRM.Controls("Cbobal0").ListIndex = 0, "balmetgetecohead", "balmetgetecogrades")
'        Let View_Core_execution = OBJCTR.Balmet(ControlId, VecBOOLEANS, VecREQUEST)
'
'      Case "projection": Let View_Core_execution = OBJCTR.Projection(VecREQUEST)
'
'      Case "oreinvoice": Let View_Core_execution = OBJCTR.Oreinvoice(VecREQUEST)
'
'      Case "concentrateinvoice": Let View_Core_execution = OBJCTR.Concentrateinvoice(VecREQUEST)
'
'      Case "configlauncher": Let View_Core_execution = View_Core_execution_config(StrControl, FRM)
'
'      Case Else: GoTo EH
'    End Select
'  End With
'
'  If View_Core_execution = False Then Call View_Core_error_handler(ControlId, FRM)
'
'  With Application 'TODO: ¿REALMENTE DEBE IR AQUÍ?, ...PUEDE SER MEJOR EN BEFORECLOSE DE THISWORKBOOK
'    'Let .DisplayAlerts = False
'    'Call .ThisWorkbook.Save
'    'Let .DisplayAlerts = True
'  End With
  
EH:
  If Not View_Core_execution And BooNewSheet Then
    With Application
      Let .DisplayAlerts = False
      Set Wbk = VecReq(LngVecRngIndex).Parent.Parent
      Let StrSheetName = VecReq(LngVecRngIndex).Parent.Name
      Call Wbk.Worksheets(StrSheetName).Delete ' Delete new sheet created
      Let .DisplayAlerts = True
    End With
  End If
  Let VecReq = Empty
  Let VecRes = Empty
  Set Wbk = Nothing
  Set RES = Nothing
  Let Application.ScreenUpdating = True
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function View_Core_execution_config(ByVal StrControl As String, ByVal FRM As FRMBasics) As Boolean
  Dim RES As AppResources
  
  Dim FD As FileDialog
  Dim RngBox As Range
  Dim VecAux0 As Variant
  Dim StrAux0 As String
  Dim BooAux0 As Boolean
  Dim i As Long, j As Long, k As Long
  
  Const sFN As String = "VIEW::-EXE_config"

  On Error GoTo EH
  Let View_Core_execution_config = False
  
  With FRM
    Select Case StrControl
      'DIVISA
      Case "Btncon0" 'Add
        Let .Controls("Lstcon0").Tag = "noevent"
        Let .Controls("Lstcon0").RowSource = VBA.vbNullString
        Set RES = New AppResources
        Let VecAux0 = Array(.Controls("Txtcon0").Value, .Controls("Txtcon1").Value, .Controls("Txtcon2").Value, 0, 0)
        Call RES.resExcelxlTABLEADDROWFROMARRAY(excliqlitedatasheet.ListObjects("excliqlitecurrenciescon"), VecAux0, True)
        Let VecAux0 = Empty
        Set RES = Nothing
        Call View_Core_execution_config("Btncon2", FRM)
      Case "Btncon1" 'Upd
        Let i = .Controls("Lstcon0").ListIndex
        If .Controls("Lstcon0").ListCount > 0 And i > 1 Then
          Let VecAux0 = Array(.Controls("Txtcon0").Value, .Controls("Txtcon1").Value, .Controls("Txtcon2").Value, .Controls("Lstcon0").List(i, 3), .Controls("Lstcon0").List(i, 4))
          Let .Controls("Lstcon0").Tag = "noevent"
          Let .Controls("Lstcon0").RowSource = VBA.vbNullString
          Let excliqlitedatasheet.ListObjects("excliqlitecurrenciescon").ListRows(i + 1).Range.Value = VecAux0
          Let VecAux0 = Empty
          Call View_Core_execution_config("Btncon2", FRM)
        End If
      Case "Btncon2" 'Dcn
        Let .Controls("Lstcon0").Tag = "noevent"
        Let .Controls("Lstcon0").RowSource = VBA.vbNullString
        Let .Controls("Lstcon0").RowSource = excliqlitedatasheet.ListObjects("excliqlitecurrenciescon")
        Let .Controls("Lstcon0").Tag = VBA.vbNullString
        Let BooAux0 = (.Controls("Lstcon0").ListCount > 0)
        Let .Controls("Btncon0").Enabled = BooAux0
        Let .Controls("Btncon1").Enabled = False
        Let .Controls("Btncon2").Enabled = False
        Let .Controls("Btncon3").Enabled = (.Controls("Lstcon0").ListCount > 2)
        Let .Controls("Lstcon0").Enabled = BooAux0
        Let .Controls("Multi").Pages(1).Enabled = True
        Let .Controls("Multi").Pages(2).Enabled = True
        Let .Controls("Multi").Pages(3).Enabled = True
        Let j = 0: Let k = 0
        If .Controls("Lstcon0").ListCount > 0 Then
          Call .Controls("Cbocon0").Clear
          Call .Controls("Cbocon1").Clear
          For i = 0 To .Controls("Lstcon0").ListCount - 1
            If VBA.CBool(.Controls("Lstcon0").List(i, 3)) Then Let j = i
            If VBA.CBool(.Controls("Lstcon0").List(i, 4)) Then Let k = i
            Let StrAux0 = .Controls("Lstcon0").List(i, 0) & " [" & .Controls("Lstcon0").List(i, 2) & "]"
            Call .Controls("Cbocon0").AddItem(StrAux0)
            Call .Controls("Cbocon1").AddItem(StrAux0)
          Next i
          Let .Controls("Cbocon0").List(j, 0) = .Controls("Cbocon0").List(j, 0) & " *"
          Let .Controls("Cbocon1").List(k, 0) = .Controls("Cbocon1").List(k, 0) & " *"
          Let .Controls("Cbocon0").ListIndex = j
          Let .Controls("Cbocon1").ListIndex = k
          Let .Controls("Btncon4").Enabled = (.Controls("Cbocon0").ListCount > 0)
          Let .Controls("Btncon5").Enabled = (.Controls("Cbocon1").ListCount > 0)
        End If
        GoTo EH
      Case "Btncon3" 'Del
        Let i = .Controls("Lstcon0").ListIndex
        If .Controls("Lstcon0").ListCount > 0 And i > 1 Then
          Let .Controls("Lstcon0").Tag = "noevent"
          Let .Controls("Lstcon0").RowSource = VBA.vbNullString
          Call excliqlitedatasheet.ListObjects("excliqlitecurrenciescon").ListRows(i + 1).Delete
          Call View_Core_execution_config("Btncon2", FRM)
        End If
      Case "Btncon4" 'Local
        Let i = .Controls("Cbocon0").ListIndex + 1
        If .Controls("Cbocon0").ListCount > 0 And i >= 1 Then
          Let .Controls("Lstcon0").Tag = "noevent"
          Let .Controls("Lstcon0").RowSource = VBA.vbNullString
          Let VecAux0 = excliqlitedatasheet.ListObjects("excliqlitecurrenciescon").DataBodyRange
          For j = LBound(VecAux0) To UBound(VecAux0) 'Local
            If VBA.CByte(VecAux0(j, 4)) = 1 Then Exit For
          Next j
          For k = LBound(VecAux0) To UBound(VecAux0) 'Foreign
            If VBA.CByte(VecAux0(k, 5)) = 1 Then Exit For
          Next k
          Let excliqlitedatasheet.ListObjects("excliqlitecurrenciescon").ListColumns("currency_main").DataBodyRange.Value = 0
          Let excliqlitedatasheet.ListObjects("excliqlitecurrenciescon").ListColumns("currency_main").DataBodyRange.Cells(i, 1).Value = 1
          If i = k Then
            Let excliqlitedatasheet.ListObjects("excliqlitecurrenciescon").ListColumns("currency_foreign").DataBodyRange.Value = 0
            Let excliqlitedatasheet.ListObjects("excliqlitecurrenciescon").ListColumns("currency_foreign").DataBodyRange.Cells(j, 1).Value = 1
          End If
          Call View_Core_execution_config("Btncon2", FRM)
        End If
      Case "Btncon5" 'Foreign
        Let i = .Controls("Cbocon1").ListIndex + 1
        If .Controls("Cbocon1").ListCount > 0 And i >= 1 Then
          Let .Controls("Lstcon0").Tag = "noevent"
          Let .Controls("Lstcon0").RowSource = VBA.vbNullString
          Let VecAux0 = excliqlitedatasheet.ListObjects("excliqlitecurrenciescon").DataBodyRange
          For j = LBound(VecAux0) To UBound(VecAux0) 'Local
            If VBA.CByte(VecAux0(j, 4)) = 1 Then Exit For
          Next j
          For k = LBound(VecAux0) To UBound(VecAux0) 'Foreign
            If VBA.CByte(VecAux0(k, 5)) = 1 Then Exit For
          Next k
          Let excliqlitedatasheet.ListObjects("excliqlitecurrenciescon").ListColumns("currency_foreign").DataBodyRange.Value = 0
          Let excliqlitedatasheet.ListObjects("excliqlitecurrenciescon").ListColumns("currency_foreign").DataBodyRange.Cells(i, 1).Value = 1
          If i = j Then
            Let excliqlitedatasheet.ListObjects("excliqlitecurrenciescon").ListColumns("currency_main").DataBodyRange.Value = 0
            Let excliqlitedatasheet.ListObjects("excliqlitecurrenciescon").ListColumns("currency_main").DataBodyRange.Cells(k, 1).Value = 1
          End If
          Call View_Core_execution_config("Btncon2", FRM)
        End If
      
      'SELECCIONES
      Case "Chkcon0", "Chkcon1"
        Set RES = New AppResources
        If StrControl = "Chkcon0" Then
          Call RES.resAPPEQSETTABLESELECTIONSMODSCOMPACT(.Controls(StrControl).Value)
        Else
          Call RES.resAPPEQSETTABLESELECTIONSMODSHEADS(.Controls(StrControl).Value)
        End If
        Set RES = Nothing
      
      'PÁGINA
      Case "Chkcon2" 'Always setup page
        Set RES = New AppResources
        Call RES.resAPPEQSETTABLE_PAGEOPTIONS_EXISTS(.Controls(StrControl).Value)
        Set RES = Nothing
      
      Case "Btncon6" 'Get Image
        Set FD = Application.FileDialog(msoFileDialogFilePicker)
        With FD
          Let .AllowMultiSelect = False
          Let .ButtonName = "Cargar imagen"
          Let .Title = "Cargar imagen para Encabezado de página - " & STRAPPTITLE
          Call .Filters.Add("Archivos de imagen", "*.jpg; *.png; *.gif; *.bmp; *.tiff", 1)
          Let i = .Show()
          If i = -1 Then
            Let FRM.Controls("Txtcon3").Value = VBA.CStr(.SelectedItems.Item(1))
          End If
        End With
      
      Case "Btncon7" 'Save all
        Set RES = New AppResources
        Call RES.resAPPEQSETTABLE_PAGEOPTIONS_LOGO(.Controls("Txtcon3").Value)
        Call RES.resAPPEQSETTABLE_PAGEOPTIONS_COMPANY(.Controls("Txtcon4").Value)
        Call RES.resAPPEQSETTABLE_PAGEOPTIONS_ADDRESS(.Controls("Txtcon5").Value)
        Call RES.resAPPEQSETTABLE_PAGEOPTIONS_PHONE(.Controls("Txtcon6").Value)
        Call RES.resAPPEQSETTABLE_PAGEOPTIONS_EMAIL(.Controls("Txtcon7").Value)
        Call RES.resAPPEQSETTABLE_PAGEOPTIONS_LOGO_EXISTS(.Controls("Cbocon2").ListIndex)
        Call RES.resAPPEQSETTABLE_PAGEOPTIONS_COMPANY_EXISTS(.Controls("Cbocon3").ListIndex)
        Call RES.resAPPEQSETTABLE_PAGEOPTIONS_ADDRESS_EXISTS(.Controls("Cbocon4").ListIndex)
        Call RES.resAPPEQSETTABLE_PAGEOPTIONS_PHONE_EXISTS(.Controls("Cbocon5").ListIndex)
        Call RES.resAPPEQSETTABLE_PAGEOPTIONS_EMAIL_EXISTS(.Controls("Cbocon6").ListIndex)
        Set RES = Nothing
        Call VBA.MsgBox("Datos guardados correctamente", vbInformation, STRAPPTITLE)
      
      Case "Btncon8" 'Add Destiny
        Let .Controls("Cbocon7").Tag = "noevent"
        Let .Controls("Cbocon7").RowSource = VBA.vbNullString
        Set RES = New AppResources
        Call RES.resExcelxlTABLEADDROWFROMARRAY(excliqlitedatasheet.ListObjects("excliqliteconfigpurdestiny"), Array(.Controls("Cbocon7").Value), True)
        Set RES = Nothing
        Let .Controls("Cbocon7").RowSource = excliqlitedatasheet.ListObjects("excliqliteconfigpurdestiny")
        Let .Controls("Cbocon7").Tag = VBA.vbNullString
        Call VBA.MsgBox("Dato insertado correctamente", vbInformation, STRAPPTITLE)
        
      Case "Btncon9" 'Upd Destiny
        Let i = .Controls("Cbocon7").Tag + 1
        Let .Controls("Cbocon7").Tag = "noevent"
        Let .Controls("Cbocon7").RowSource = VBA.vbNullString
        Let excliqlitedatasheet.ListObjects("excliqliteconfigpurdestiny").ListColumns("purch_destiny").DataBodyRange.Cells(i, 1).Value = VBA.CStr(.Controls("Cbocon7").Value)
        Let .Controls("Cbocon7").RowSource = excliqlitedatasheet.ListObjects("excliqliteconfigpurdestiny")
        Let .Controls("Cbocon7").Tag = VBA.vbNullString
        Call VBA.MsgBox("Dato modificado correctamente", vbInformation, STRAPPTITLE)

      Case "Btncon10" 'Del Destiny
        Let i = .Controls("Cbocon7").Tag + 1
        Let .Controls("Cbocon7").Tag = "noevent"
        Let .Controls("Cbocon7").RowSource = VBA.vbNullString
        Call excliqlitedatasheet.ListObjects("excliqliteconfigpurdestiny").ListRows(i).Delete
        Let .Controls("Cbocon7").RowSource = excliqlitedatasheet.ListObjects("excliqliteconfigpurdestiny")
        Let .Controls("Cbocon7").ListIndex = 0
        Let .Controls("Cbocon7").Tag = VBA.vbNullString
        Call VBA.MsgBox("Registo eliminado correctamente", vbInformation, STRAPPTITLE)

      'LIQUIDACIONES
      Case "Btncon11", "Btncon12" 'Add, Upd
        Let i = .Controls("Lstcon1").ListIndex + 1
        Select Case .Controls("Cbocon8").ListIndex
          Case 0 'RM
            Let StrAux0 = VBA.Replace(.Controls("Cbocon9").Value, ".", VBA.vbNullString)
            Let StrAux0 = .Controls("Cbocon10").Value & " a " & StrAux0
            Let VecAux0 = Array( _
              .Controls("Txtcon8").Value, _
              .Controls("Txtcon9").Value, _
              .Controls("Cbocon9").Value, _
              VBA.Replace(.Controls("Txtcon10").Value, ",", "."), _
              VBA.Replace(.Controls("Txtcon11").Value, ",", "."), _
              VBA.Replace(.Controls("Txtcon12").Value, ",", "."), _
              .Controls("Cbocon10").Value, StrAux0)
            Let StrAux0 = "excliqliteconfigpurrm"
          Case 1 'Ret
            Let VecAux0 = Array(.Controls("Txtcon8").Value, VBA.Replace(.Controls("Txtcon9").Value, ",", "."), 0)
            Let StrAux0 = "excliqliteconfigpurretentions"
          Case 2 'RetOth
            Let VecAux0 = Array(.Controls("Txtcon8").Value, VBA.Replace(.Controls("Txtcon9").Value, ",", "."), 0)
            Let StrAux0 = "excliqliteconfigpurretentionsother"
          Case 3 'Penalties
            Let VecAux0 = Array( _
                .Controls("Txtcon8").Value, _
                VBA.Replace(.Controls("Txtcon9").Value, ",", "."), _
                VBA.Replace(.Controls("Txtcon10").Value, ",", "."), _
                VBA.Replace(.Controls("Txtcon11").Value, ",", "."), _
                VBA.Replace(.Controls("Txtcon12").Value, ",", "."), 0)
            Let StrAux0 = "excliqliteconpurpenalties"
          Case 4 'Deductions in Exports
            Let VecAux0 = Array( _
                .Controls("Txtcon8").Value, _
                VBA.Replace(.Controls("Txtcon9").Value, ",", "."), _
                VBA.Replace(.Controls("Txtcon10").Value, ",", "."), _
                IIf(.Controls("Chkcon4").Value, 1, 0), 0)
            Let StrAux0 = "excliqliteconpurdedexp"

          Case Else: GoTo EH
        End Select
        
        Let .Controls("Lstcon1").Tag = "noevent"
        Let .Controls("Lstcon1").RowSource = VBA.vbNullString
        If StrControl = "Btncon11" Then
          Set RES = New AppResources
          Call RES.resExcelxlTABLEADDROWFROMARRAY(excliqlitedatasheet.ListObjects(StrAux0), VecAux0, True)
          Set RES = Nothing
          Let StrAux0 = "Registro insertado correctamente"
        Else
          Let excliqlitedatasheet.ListObjects(StrAux0).ListRows(i).Range.Value = VecAux0
          Let StrAux0 = "Registro modificado correctamente"
        End If
        Let VecAux0 = Empty
        Call View_Core_execution_config("Btncon13", FRM)
        Call VBA.MsgBox(StrAux0, vbInformation, STRAPPTITLE)

      Case "Btncon13" 'Ucn
        Let i = .Controls("Cbocon8").ListIndex
        Let .Controls("Cbocon8").ListIndex = -1
        Let .Controls("Cbocon8").ListIndex = i
        Let .Controls("Cbocon8").Enabled = True
        Let .Controls("Multi").Pages(0).Enabled = True
        Let .Controls("Multi").Pages(1).Enabled = True
        Let .Controls("Multi").Pages(2).Enabled = True

      Case "Btncon14" 'Del
        Let i = .Controls("Lstcon1").ListIndex + 1
        Select Case .Controls("Cbocon8").ListIndex
          Case 0 'RM
            Let StrAux0 = "excliqliteconfigpurrm"
          Case 1 'Ret
            Let StrAux0 = "excliqliteconfigpurretentions"
          Case 2 'RetOth
            Let StrAux0 = "excliqliteconfigpurretentionsother"
          Case 3 'Penalties
            Let StrAux0 = "excliqliteconpurpenalties"
          Case 4 'Deductions in Exports
            Let StrAux0 = "excliqliteconpurdedexp"

          Case Else: GoTo EH
        End Select
        
        Let .Controls("Lstcon1").Tag = "noevent"
        Let .Controls("Lstcon1").RowSource = VBA.vbNullString
        Call excliqlitedatasheet.ListObjects(StrAux0).ListRows(i).Delete
        Let StrAux0 = "Registro eliminado correctamente"
        Call View_Core_execution_config("Btncon13", FRM)
        Call VBA.MsgBox(StrAux0, vbInformation, STRAPPTITLE)
      
      Case "Chkcon3"
        Set RES = New AppResources
        Call RES.resAPPEQSETTABLE_PUROPTIONS_RMDETAIL(.Controls(StrControl).Value)
        Set RES = Nothing
      
    End Select
  End With
  
  'Let View_Core_execution_config = True

EH:
  Let VecAux0 = Empty
  Set RngBox = Nothing
  Set FD = Nothing
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)
End Function








' CORE VIEWS FORMS METHODS LIST -----------------------------------------------------'
' Regression <
Private Sub View_form_regression(ByVal ControlId As String, ByVal FRM As FRMBasics)
  
  'Main objects
  Dim VIEWS As ViewsParent_current
  'ListObjects
  'LO data names and other data
  'Controls
  Dim aLabels() As Variant
  Dim aMaxLengths() As Variant
  Dim aControls() As Variant
  Dim aControlis() As Variant
  'Standard vars
  Dim StrTitle As String, sTable As String, TitleForm As String
  Dim i As Long, j As Long, k As Long, l As Long
  Dim xLeft As Long
  Dim Aux0 As String, Aux1 As String
  
  Const sFN As String = "VIEWREG::-FormRegression"


  #If Debugging Then
    Call Debug_assert(Not ControlId = VBA.vbNullString, sFN)
    Call Debug_assert(Not FRM Is Nothing, sFN)
  #End If
  
  If ControlId = VBA.vbNullString Then GoTo EH
  If FRM Is Nothing Then GoTo EH
  On Error GoTo EH

  '1. [x] Get data for Form Controls

  '2. [x] Desing the Form
  Let TitleForm = "Regresión Lineal"
  Let StrTitle = TitleForm & " - " & STRAPPTITLE
  
  Set VIEWS = New ViewsParent_current
  Set VIEWS.ErrorHandler = EHGLOBAL
  Call VIEWS.ViewFormCleanner(FRM)
  With FRM
    'Set the form dimensions
    Let .Height = 415
    Let .Width = 450
    Let .Tag = ControlId
    Let .Caption = StrTitle
    Let xLeft = .InsideWidth * 0.05

    'Build and design the form
    Call VIEWS.ViewFormAddControl(FRM, "CommandButton", "Btnmain", 2) 'All with events (Cancel, Exe and Help)
    Call VIEWS.ViewFormAddControl(FRM, "Frame", "Framereg", 3)
    
    Call VIEWS.ViewFormAddControl(.Controls("Framereg0"), "Label", "Lblreg", 2)
    Call VIEWS.ViewFormAddControl(.Controls("Framereg0"), "TextBox", "Txtreg", 2) ' EVENT all
    With .Controls("Framereg0")
      
      Let .Caption = "Ingresar Datos"
      Let .Width = .Parent.InsideWidth * 0.9
      Let .Height = .Parent.InsideHeight * 0.3
      Let .Left = xLeft
      Let .Top = .Parent.InsideHeight * 0.04

      Let aLabels = Array("Datos de Variable X (Opcional)", "Datos de Variable Y", "Datos a extrapolar (Opcional)")
      Let aControls = Array("Lblreg0", "Lblreg1", "Lblreg2")
      Let j = xLeft
      For i = 0 To UBound(aLabels)
        Let .Controls(aControls(i)).Caption = aLabels(i)
        Let .Controls(aControls(i)).Left = xLeft
        Let .Controls(aControls(i)).AutoSize = True
        Let .Controls(aControls(i)).WordWrap = False
        Let .Controls(aControls(i)).Top = xLeft + ((.Controls(aControls(i)).Height * 2.5) * i)
      Next i
      Erase aLabels: Erase aControlis
      Let .Controls("Txtreg0").Top = .Controls("Lblreg0").Top
      Let .Controls("Txtreg1").Top = .Controls("Lblreg1").Top
      Let .Controls("Txtreg2").Top = .Controls("Lblreg2").Top
      Let k = .Controls("Lblreg0").Left + .Controls("Lblreg0").Width + 5
      Let j = .Width - .Controls("Lblreg0").Left - .Controls("Lblreg0").Width - (.Controls("Txtreg2").Height * 1.5) ' (.Controls("Btnreg0").Height * 3)

      For i = 0 To 2
        With .Controls("Txtreg" & i)
          Let .Left = k
          Let .Width = j
          Let .SpecialEffect = fmSpecialEffectRaised
          Let .DropButtonStyle = fmDropButtonStyleReduce
          Let .ShowDropButtonWhen = fmShowDropButtonWhenAlways
        End With
      Next i
    
    End With

    Call VIEWS.ViewFormAddControl(.Controls("Framereg1"), "CheckBox", "Chkreg", 9)
    With .Controls("Framereg1")
      
      Let .Caption = "Incluir en resultado"
      Let .Width = .Parent.InsideWidth * 0.9
      Let .Height = .Parent.InsideHeight * 0.3
      Let .Left = xLeft
      Let .Top = .Parent.Controls("Framereg0").Top + .Parent.Controls("Framereg0").Height + 5

      Let aControls = Array("Chkreg0", "Chkreg1", "Chkreg2", "Chkreg3", "Chkreg4", "Chkreg5", "Chkreg6", "Chkreg7", "Chkreg8", "Chkreg9")
      Let aLabels = Array("Todo", "Sólo Regresión", "Valores de X^2", "Valores de Y^2", "Valores de XY", "Valor de a", "Valor de b", "Coeficiente de correlación (r)", "Coeficiente de determinación (r^2)", "Error estándar de determinación (Se)")

      For i = 0 To UBound(aControls)
        Let .Controls(aControls(i)).Caption = aLabels(i)
        Let .Controls(aControls(i)).AutoSize = True
        Let .Controls(aControls(i)).WordWrap = False
        Let .Controls(aControls(i)).Left = xLeft + IIf(i < 5, 0, (.Controls(aControls(0)).Left + .Controls(aControls(0)).Width) * 2)
        Let .Controls(aControls(i)).Top = (xLeft / 1.4) * (IIf(i < 5, i, (i - 5)) + 1)
        If i = 0 Then Let .Controls(aControls(i)).Value = True
      Next i
    
    End With

    Call VIEWS.ViewFormAddControl(.Controls("Framereg2"), "CheckBox", "Chkreg", 10, CtlFromX:=10)
    Call VIEWS.ViewFormAddControl(.Controls("Framereg2"), "OptionButton", "Optreg", 1)
    With .Controls("Framereg2")
      
      Let .Caption = "Gráfico del ejercicio"
      Let .Width = ((.Parent.InsideWidth * 0.9) / 2) - 2.5
      Let .Height = .Parent.InsideHeight * 0.15
      Let .Left = xLeft
      Let .Top = .Parent.Controls("Framereg1").Top + .Parent.Controls("Framereg1").Height + 5

      Let .Controls("Chkreg10").Caption = "Generar gráfico en:"
      Let .Controls("Optreg0").Caption = "Hoja actual"
      Let .Controls("Optreg1").Caption = "Nueva Hoja de gráfico"

      Let .Controls("Chkreg10").AutoSize = True
      Let .Controls("Optreg0").AutoSize = True
      Let .Controls("Optreg1").AutoSize = True
      Let .Controls("Chkreg10").WordWrap = False
      Let .Controls("Optreg0").WordWrap = False
      Let .Controls("Optreg1").WordWrap = False

      Let .Controls("Optreg0").Value = True

      Let i = xLeft / 2
      Let .Controls("Chkreg10").Top = i
      Let .Controls("Optreg0").Top = i + .Controls("Chkreg10").Height
      Let .Controls("Optreg1").Top = i + .Controls("Chkreg10").Height

      Let .Controls("Chkreg10").Left = xLeft
      Let .Controls("Optreg0").Left = xLeft
      Let .Controls("Optreg1").Left = (xLeft / 2) + .Controls("Optreg0").Left + .Controls("Optreg0").Width
    
    End With

    Call VIEWS.ViewFormAddControl(.Controls("Framereg3"), "OptionButton", "Optreg", 3, CtlFromX:=2)
    Call VIEWS.ViewFormAddControl(.Controls("Framereg3"), "TextBox", "Txtreg", 3, CtlFromX:=3)
    With .Controls("Framereg3")
      
      Let .Caption = "Dónde mostrar resultados"
      Let .Width = ((.Parent.InsideWidth * 0.9) / 2) - 2.5
      Let .Height = .Parent.InsideHeight * 0.15
      Let .Left = .Parent.Controls("Framereg2").Left + .Parent.Controls("Framereg2").Width + 5
      Let .Top = .Parent.Controls("Framereg1").Top + .Parent.Controls("Framereg1").Height + 5

      Let .Controls("Optreg2").Caption = "Hoja nueva"
      Let .Controls("Optreg3").Caption = "Hoja actual"

      Let .Controls("Optreg2").AutoSize = True
      Let .Controls("Optreg3").AutoSize = True
      Let .Controls("Optreg2").WordWrap = False
      Let .Controls("Optreg3").WordWrap = False

      Let .Controls("Optreg2").Left = xLeft
      Let .Controls("Optreg3").Left = xLeft
      Let .Controls("Optreg2").Top = xLeft / 2.3
      Let .Controls("Optreg3").Top = (xLeft / 2.3) + .Controls("Optreg2").Height

      Let .Controls("Optreg2").Value = True

      With .Controls("Txtreg3")
        Let .Left = FRM.Controls("Optreg3").Left + FRM.Controls("Optreg3").Width
        Let .Top = FRM.Controls("Optreg3").Top
        Let .Width = .Width * 1.3
        Let .SpecialEffect = fmSpecialEffectRaised
        Let .DropButtonStyle = fmDropButtonStyleReduce
        Let .ShowDropButtonWhen = fmShowDropButtonWhenAlways
      End With
    
    End With

    'Push data to controls

    'General purpose CommandButtons
    '3. Fill Form Controls with data and Set general purpose CommandButtons
    Let k = (.InsideWidth * 0.95)
    With .Controls("Btnmain0") '("BtnCancel")
      Let .Accelerator = "c"
      Let .Cancel = True
      Let .Top = (.Parent.InsideHeight * 0.95) - (.Height)
      Let .Left = k - .Width
      Let .Caption = "Cerrar"
    End With
    With .Controls("Btnmain2") '("BtnHelp")
      Let .Caption = "Ayuda"
      Let .Accelerator = "a"
      Let .ControlTipText = "Ver ayuda del Ejercicio"
      Let .Top = .Parent.Controls("Btnmain0").Top
      Let .Left = xLeft
      Let .Tag = ControlId '& cGet
    End With
    With .Controls("Btnmain1") '("BtnEXE")
      Let .Caption = "Resolver"
      Let .Accelerator = "r"
      Let .ControlTipText = "Permite procesar los datos y mostrar los resultados en una hoja Excel."
      Let .Top = .Parent.Controls("Btnmain0").Top
      Let .Left = .Parent.Controls("Btnmain0").Left - .Width - 2
      Let .Tag = ControlId & "-reg"
      Let .Default = True
    End With

    '4. Assing handler for each button
    Set BtnCancel = .Controls("Btnmain0") 'Main
    Set BtnAccept = .Controls("Btnmain1") 'Main

    Set Btn0 = .Controls("Btnmain2") 'Main
    Set Txt0 = .Controls("Txtreg0") 'Frame0
    Set Txt1 = .Controls("Txtreg1") 'Frame0
    Set Txt2 = .Controls("Txtreg2") 'Frame0
    Set Txt3 = .Controls("Txtreg3") 'Frame3
    
    '5. Give styles to controls
    Call VIEWS.ViewFormStyles(FRM): Set VIEWS = Nothing

    '6. Show the finished Form
    Call Load(FRM)
    Call .Show

    '7. Close form
  End With

EH:
  Erase aLabels
  Erase aMaxLengths
  Erase aControlis
  Erase aControls
  Set VIEWS = Nothing
  Debug.Print VBA.Err.Description
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Sub

'Selection Regression UI
Private Sub View_form_regressionselection(ByVal ControlId As String)
  
  Const sFN As String = "VIEW::-regressionselection"
  
  On Error GoTo EH
  Call View_Controls_buttons_click(Nothing, Nothing, True, ControlId)
'  If View_Controls_buttons_click(Nothing, Nothing, True, ControlId & "-getcell") Then
'    Call View_Controls_buttons_click(Nothing, Nothing, True, ControlId)
'  End If

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Sub


'Balmet <
Private Sub View_form_balmet(ByVal ControlId As String, ByVal FRM As FRMBasics)
  
  'Main objects
  Dim VWP As ViewsParent_current
  Dim RES As AppResources_current
  Dim MDL As ModelExcliqliteDatasheet
  'Controls
  Dim FmeCurrent As MSForms.Frame
  Dim VecLabels As Variant
  Dim VecMaxLgt As Variant
  Dim VecCntrls0 As Variant
  Dim VecCntrls1 As Variant
  'Standard vars
  Dim StrTitle As String, StrTitleForm As String, StrAux0 As String
  Dim i As Long, j As Long, k As Long, l As Long, xLeft As Long
  
  Const sFN As String = "VIEW::-form_balmet"

  #If Debugging Then
    Call Debug_assert(Not ControlId = VBA.vbNullString, sFN)
    Call Debug_assert(Not FRM Is Nothing, sFN)
  #End If
  
  On Error GoTo EH
  '1. [x] Get data for Form Controls

  '2. [x] Desing the Form
  Let StrTitleForm = "Balance Metalúrgico"
  Let StrTitle = StrTitleForm & " - " & STRAPPTITLE
  
  Set VWP = New ViewsParent_current
  Set VWP.ErrorHandler = EHGLOBAL
  Call VWP.ViewFormCleanner(FRM)
  
  With FRM
    
    'Set the form dimensions
    Let .Height = 415
    Let .Width = 450
    Let .Tag = ControlId
    Let .Caption = StrTitle
    Let xLeft = .InsideWidth * 0.05 ' Left for all main controls

    'Build and design the form
    Call VWP.ViewFormAddControl(FRM, "CommandButton", "Btnmain", 2) 'All with events (Cancel, Exe and Help)
    Call VWP.ViewFormAddControl(FRM, "Frame", "Framebal", 3)
    
    Set FmeCurrent = .Controls("Framebal0")
    With FmeCurrent
      
      ' Set Frame properties
      Let .Caption = "Ingresar Datos"
      Let .Width = .Parent.InsideWidth * 0.9
      Let .Height = .Parent.InsideHeight * 0.325
      Let .Left = xLeft
      Let .Top = .Parent.InsideHeight * 0.04
      
      ' Create Frame controls
      Call VWP.ViewFormAddControl(FmeCurrent, "Label", "Lblbal", 2)
      Call VWP.ViewFormAddControl(FmeCurrent, "ComboBox", "Cbobal", 2) ' EVENT 1,2
      Call VWP.ViewFormAddControl(FmeCurrent, "TextBox", "Txtbal", 4) ' EVENT 0,2
      Call VWP.ViewFormAddControl(FmeCurrent, "CheckBox", "Chkbal", 2) ' EVENT all
      Call VWP.ViewFormAddControl(FmeCurrent, "CommandButton", "Btnbal", 3) 'EVENT all
      
      ' Set Frame controls layout
      Let VecLabels = Array("Método de resolución:", "Peso Alimentación [T]", "Leyes de minerales y productos:")
      Let VecCntrls0 = Array("Lblbal0", "Lblbal1", "Lblbal2")
      Let j = xLeft * 0.7
      For i = 0 To UBound(VecLabels)
        Let .Controls(VecCntrls0(i)).Caption = VecLabels(i)
        Let .Controls(VecCntrls0(i)).Left = xLeft
        Let .Controls(VecCntrls0(i)).AutoSize = True
        Let .Controls(VecCntrls0(i)).WordWrap = False
        Let .Controls(VecCntrls0(i)).Top = j + ((.Controls(VecCntrls0(i)).Height * 2) * i)
      Next i
      Let VecLabels = Empty
      
      Let j = .Controls("Lblbal2").Left + .Controls("Lblbal2").Width + 3
      Let .Controls("Cbobal0").Left = j
      Let .Controls("Txtbal0").Left = j
      Let .Controls("Chkbal1").Left = j
      
      Let .Controls("Cbobal0").Top = .Controls("Lblbal0").Top
      Let .Controls("Txtbal0").Top = .Controls("Lblbal1").Top
      Let .Controls("Chkbal1").Top = .Controls("Lblbal2").Top * 0.98
      
      Let .Controls("Cbobal0").Width = .Width - (.Controls("Cbobal0").Left + xLeft)
      Let .Controls("Txtbal0").Width = .Controls("Cbobal0").Width * 0.5
      Let .Controls("Chkbal0").Left = .Controls("Txtbal0").Left + .Controls("Txtbal0").Width + 2
      Let .Controls("Chkbal0").Top = .Controls("Txtbal0").Top
      Let .Controls("Chkbal0").AutoSize = True
      Let .Controls("Chkbal0").WordWrap = False
      Let .Controls("Txtbal1").Left = .Controls("Chkbal0").Left + .Controls("Chkbal0").Width
      Let .Controls("Txtbal1").Top = .Controls("Txtbal0").Top
      Let .Controls("Txtbal1").Width = .Controls("Txtbal0").Width - .Controls("Chkbal0").Width - 2
      
      Let j = ((.Width - (xLeft * 2)) / 3) - 2
      Let .Controls("Txtbal2").Left = .Controls("Lblbal2").Left
      Let .Controls("Txtbal2").Top = .Controls("Chkbal1").Top + .Controls("Chkbal1").Height
      Let .Controls("Txtbal2").Width = j
      
      Let .Controls("Cbobal1").Left = .Controls("Chkbal1").Left + 2
      Let .Controls("Cbobal1").Top = .Controls("Txtbal2").Top
      Let .Controls("Cbobal1").Width = j - (.Controls("Cbobal1").Height * 2)
      Let .Controls("Btnbal0").Left = .Controls("Cbobal1").Left + .Controls("Cbobal1").Width
      Let .Controls("Btnbal0").Top = .Controls("Cbobal1").Top
      Let .Controls("Btnbal0").Height = .Controls("Cbobal1").Height
      Let .Controls("Btnbal0").Width = .Controls("Btnbal0").Height
      Let .Controls("Btnbal1").Left = .Controls("Btnbal0").Left + .Controls("Btnbal0").Width
      Let .Controls("Btnbal1").Top = .Controls("Btnbal0").Top
      Let .Controls("Btnbal1").Height = .Controls("Btnbal0").Height
      Let .Controls("Btnbal1").Width = .Controls("Btnbal0").Height
      Let .Controls("Txtbal3").Left = .Controls("Cbobal1").Left
      Let .Controls("Txtbal3").Top = .Controls("Cbobal1").Top + .Controls("Cbobal1").Height
      Let .Controls("Txtbal3").Width = j

      Let .Controls("Chkbal2").Left = .Controls("Chkbal0").Left
      Let .Controls("Chkbal2").Top = .Controls("Chkbal1").Top
      Let .Controls("Chkbal2").AutoSize = True
      Let .Controls("Chkbal2").WordWrap = False
      Let .Controls("Cbobal2").Left = .Controls("Chkbal2").Left - 2
      Let .Controls("Cbobal2").Top = .Controls("Btnbal0").Top
      Let .Controls("Cbobal2").Width = .Controls("Cbobal1").Width
      Let .Controls("Btnbal2").Left = .Controls("Cbobal2").Left + .Controls("Cbobal2").Width
      Let .Controls("Btnbal2").Top = .Controls("Cbobal2").Top
      Let .Controls("Btnbal2").Height = .Controls("Cbobal2").Height
      Let .Controls("Btnbal2").Width = .Controls("Btnbal2").Height
      Let .Controls("Btnbal3").Left = .Controls("Btnbal2").Left + .Controls("Btnbal2").Width
      Let .Controls("Btnbal3").Top = .Controls("Btnbal2").Top
      Let .Controls("Btnbal3").Height = .Controls("Btnbal2").Height
      Let .Controls("Btnbal3").Width = .Controls("Btnbal2").Height
      Let .Controls("Txtbal4").Left = .Controls("Cbobal2").Left
      Let .Controls("Txtbal4").Top = .Controls("Cbobal2").Top + .Controls("Cbobal2").Height
      Let .Controls("Txtbal4").Width = j

      ' Set Frame controls ToolTipTexts
      Let VecLabels = Array( _
          "Seleccione un método para resolver el Balance Metalúrgico", _
          "Peso[;Volumen]: Cantidad(es) separada(s) por punto y coma (Ej. 1;2;3) o seleccione una celda o rango de celdas con datos numéricos", _
          "Active esta casilla y cambie el nombre de la Alimentación (A) a una personalizada", _
          "Escriba un nombre personalizado para la Alimentación (A)", _
          "Cantidad(es) separada(s) separadas por punto y coma (Ej. 1;2;3) o seleccione una celda o rango de celdas con datos numéricos", _
          "Active esta casilla si desea seleccionar nombres de minerales específicos", _
          "Seleccione un mineral o elemento", _
          "Elimina del cuadro de texto el último mineral o elemento seleccionado", _
          "Elimina del cuadro de texto todos los minerals o elementos seleccionados", _
          "Active esta casilla si desea seleccionar unidades de ley específicas", _
          "Seleccione una unidad de ley de mineral", _
          "Elimina del cuadro de texto la última unidad seleccionada", _
          "Elimina del cuadro de texto todas las unidades seleccionadas")
      Let VecCntrls0 = Array("Cbobal0", "Txtbal0", "Chkbal0", "Txtbal1", "Txtbal2", "Chkbal1", "Cbobal1", "Btnbal0", "Btnbal1", "Chkbal2", "Cbobal2", "Btnbal2", "Btnbal3")
      For i = 0 To UBound(VecLabels)
        Let .Controls(VBA.CStr(VecCntrls0(i))).ControlTipText = VBA.CStr(VecLabels(i))
      Next i
      Let VecCntrls0 = Empty
      Let VecLabels = Empty

    End With
    Set FmeCurrent = Nothing

    Set FmeCurrent = .Controls("Framebal1")
    With FmeCurrent
      
      ' Set Frame properties
      Let .Caption = "Complejos de metales preciosos"
      Let .Width = .Parent.InsideWidth * 0.9
      Let .Height = .Parent.InsideHeight * 0.25
      Let .Left = xLeft
      Let .Top = .Parent.Controls("Framebal0").Top + .Parent.Controls("Framebal0").Height + 5

      ' Create Frame controls
      Call VWP.ViewFormAddControl(FmeCurrent, "Label", "Lblbal", 3, CtlFromX:=3)
      Call VWP.ViewFormAddControl(FmeCurrent, "CheckBox", "Chkbal", 5, CtlFromX:=3)
      Call VWP.ViewFormAddControl(FmeCurrent, "TextBox", "Txtbal", 7, CtlFromX:=5)
      Call VWP.ViewFormAddControl(FmeCurrent, "CommandButton", "Btnbal", 7, CtlFromX:=4) 'EVENT all
      Call VWP.ViewFormAddControl(FmeCurrent, "ComboBox", "Cbobal", 4, CtlFromX:=3)

      ' Set Frame controls properties
      Let .Controls("Chkbal3").Left = xLeft
      Let .Controls("Chkbal3").Top = xLeft / 2
      Let .Controls("Chkbal3").AutoSize = True
      Let .Controls("Chkbal3").WordWrap = False

      Let .Controls("Lblbal3").Left = xLeft
      Let .Controls("Lblbal3").Top = .Controls("Chkbal3").Top + .Controls("Chkbal3").Height
      Let .Controls("Lblbal3").AutoSize = True
      Let .Controls("Lblbal3").WordWrap = False
      Let .Controls("Txtbal5").Left = xLeft
      Let .Controls("Txtbal5").Top = .Controls("Lblbal3").Top + .Controls("Lblbal3").Height + 2
      Let .Controls("Txtbal5").Width = j
      
      Let .Controls("Cbobal3").Width = .Parent.Controls("Cbobal1").Width
      Let .Controls("Cbobal3").Left = .Controls("Txtbal5").Left + .Controls("Txtbal5").Width + 2
      Let .Controls("Cbobal3").Top = .Controls("Txtbal5").Top
      Let .Controls("Btnbal4").Top = .Controls("Cbobal3").Top
      Let .Controls("Btnbal4").Left = .Controls("Cbobal3").Left + .Controls("Cbobal3").Width
      Let .Controls("Btnbal4").Height = .Controls("Cbobal3").Height
      Let .Controls("Btnbal4").Width = .Controls("Btnbal4").Height
      Let .Controls("Btnbal5").Top = .Controls("Cbobal3").Top
      Let .Controls("Btnbal5").Left = .Controls("Btnbal4").Left + .Controls("Btnbal4").Width
      Let .Controls("Btnbal5").Height = .Controls("Cbobal3").Height
      Let .Controls("Btnbal5").Width = .Controls("Btnbal5").Height
      Let .Controls("Txtbal6").Left = .Controls("Cbobal3").Left
      Let .Controls("Txtbal6").Top = .Controls("Cbobal3").Top + .Controls("Cbobal3").Height
      Let .Controls("Txtbal6").Width = j
      Let .Controls("Chkbal4").Left = .Controls("Cbobal3").Left
      Let .Controls("Chkbal4").Top = .Controls("Lblbal3").Top - 2
      Let .Controls("Chkbal4").AutoSize = True
      Let .Controls("Chkbal4").WordWrap = False
      
      Let .Controls("Cbobal4").Width = .Controls("Cbobal3").Width
      Let .Controls("Cbobal4").Left = .Controls("Btnbal5").Left + .Controls("Btnbal5").Width + 2
      Let .Controls("Cbobal4").Top = .Controls("Cbobal3").Top
      Let .Controls("Btnbal6").Top = .Controls("Cbobal4").Top
      Let .Controls("Btnbal6").Left = .Controls("Cbobal4").Left + .Controls("Cbobal4").Width
      Let .Controls("Btnbal6").Height = .Controls("Cbobal4").Height
      Let .Controls("Btnbal6").Width = .Controls("Btnbal6").Height
      Let .Controls("Btnbal7").Top = .Controls("Cbobal4").Top
      Let .Controls("Btnbal7").Left = .Controls("Btnbal6").Left + .Controls("Btnbal6").Width
      Let .Controls("Btnbal7").Height = .Controls("Cbobal4").Height
      Let .Controls("Btnbal7").Width = .Controls("Btnbal7").Height
      Let .Controls("Txtbal7").Left = .Controls("Cbobal4").Left
      Let .Controls("Txtbal7").Top = .Controls("Cbobal4").Top + .Controls("Cbobal4").Height
      Let .Controls("Txtbal7").Width = j
      Let .Controls("Chkbal5").Left = .Controls("Cbobal4").Left
      Let .Controls("Chkbal5").Top = .Controls("Chkbal4").Top
      Let .Controls("Chkbal5").AutoSize = True
      Let .Controls("Chkbal5").WordWrap = False
      
      ' Set Frame controls ToolTipTexts
      Let VecLabels = Array( _
          "Active esta casilla si los contenidos del mineral principal son de tipo complejo", _
          "Cantidad(s) separada(s) por punto y coma (Ej. 1;2;3) o seleccione una celda o rango de celdas con datos numéricos", _
          "Desactive esta casilla si desea seleccionar nombres de minerales específicos", _
          "Seleccione un mineral o elemento", _
          "Elimina del cuadro de texto el último mineral o elemento seleccionado", _
          "Elimina del cuadro de texto todos los minerals o elementos seleccionados", _
          "Desactive esta casilla si desea seleccionar unidades de ley específicas", _
          "Seleccione una unidad de ley de mineral", _
          "Elimina del cuadro de texto la última unidad seleccionada", _
          "Elimina del cuadro de texto todas las unidades seleccionadas")
      Let VecCntrls0 = Array("Chkbal3", "Txtbal5", "Chkbal4", "Cbobal3", "Btnbal4", "Btnbal5", "Chkbal5", "Cbobal4", "Btnbal6", "Btnbal7")
      For i = 0 To UBound(VecLabels)
        Let .Controls(VBA.CStr(VecCntrls0(i))).ControlTipText = VBA.CStr(VecLabels(i))
      Next i
      Let VecCntrls0 = Empty
      Let VecLabels = Empty
      

      Let .Controls("Chkbal3").Caption = "Es complejo de Plata, u otro metal precioso"
      Let .Controls("Chkbal3").ControlTipText = "Active esta casilla si su mineral es de tipo complejo e ingrese los datos correspondientes en los controles inferiores"
      Let .Controls("Lblbal3").Caption = "Leyes de minerales y productos:"





'      Let .Controls("Chkbal6").Left = xLeft
'      Let .Controls("Chkbal6").Top = xLeft / 2
'      Let .Controls("Chkbal6").Caption = "Es complejo de Plata, u otro metal precioso"
'      Let .Controls("Chkbal6").AutoSize = True
'      Let .Controls("Chkbal6").WordWrap = False
'      Let .Controls("Chkbal6").ControlTipText = "Active esta casilla si su mineral es de tipo complejo e ingrese los datos correspondientes en los controles inferiores"
'
'      Let VecCntrls0 = Array("Lblbal3", "Lblbal4", "Lblbal5")
'      Let VecCntrls1 = Array("Cbobal6", "Txtbal8", "Cbobal7")
'      Let VecLabels = Array("Mineral:", "Leyes:", "Unidad")
'      For i = 0 To UBound(VecCntrls0)
'        Let .Controls(VecCntrls0(i)).Caption = VecLabels(i)
'        Let .Controls(VecCntrls0(i)).AutoSize = True
'        Let .Controls(VecCntrls0(i)).WordWrap = False
'        Let .Controls(VecCntrls0(i)).Left = xLeft + (((.Width - (xLeft * 2)) / 3) * i) + IIf(i = 2, .Controls(VecCntrls1(0)).Height * 2, 0)
'        Let .Controls(VecCntrls0(i)).Top = (.Controls("Chkbal6").Top + (xLeft / 3)) + .Controls(VecCntrls0(i)).Height
'        Let .Controls(VecCntrls1(i)).Left = .Controls(VecCntrls0(i)).Left
'        Let .Controls(VecCntrls1(i)).Top = (.Controls(VecCntrls0(i)).Top + .Controls(VecCntrls0(i)).Height + 2)
'      Next i
'      Erase VecCntrls0: Erase VecCntrls1: Erase VecLabels
'
'      Let .Controls("Btnbal4").Left = .Controls("Txtbal8").Left + .Controls("Txtbal8").Width + 1
'      Let .Controls("Btnbal4").Top = .Controls("Txtbal8").Top
'      Let .Controls("Btnbal4").Height = .Controls("Txtbal8").Height
'      Let .Controls("Btnbal4").Width = .Controls("Btnbal4").Height
'      Let .Controls("Btnbal4").Caption = "_"
'      Let .Controls("Btnbal4").ControlTipText = "Pulse para seleccionar una celda"
'      Let .Controls("Btnbal4").Font.Bold = True
'      Let .Controls("Btnbal4").Tag = ControlId & "-rng"
'
'      Let .Controls("Btnbal5").Left = .Controls("Btnbal4").Left + .Controls("Btnbal4").Width + 1
'      Let .Controls("Btnbal5").Top = .Controls("Btnbal4").Top
'      Let .Controls("Btnbal5").Height = .Controls("Btnbal4").Height
'      Let .Controls("Btnbal5").Width = .Controls("Btnbal4").Height
'      Let .Controls("Btnbal5").Caption = VBA.ChrW(&H2716)
'      Let .Controls("Btnbal5").ControlTipText = "Descartar"
'      Let .Controls("Btnbal5").Font.Bold = True
'      Let .Controls("Btnbal5").Tag = ControlId & "-clr"
'
'      Let .Controls("Cbobal6").List = Array("Ag", "Au", "Pt")
'      Let .Controls("Cbobal7").List = Array("DM", "g/T", "g/L")
'      Let .Controls("Txtbal8").ControlTipText = "Ingrese las leyes del elemento en cada producto (Alim., Conc. y Relave) separados por punto y comas (Ej. 1;2;3) o un rango mediante el botón"
'      Let .Controls("Cbobal6").ControlTipText = "Seleccione un elemento metal precioso"
'      Let .Controls("Cbobal7").ControlTipText = "Seleccione la unidad de ley del elemento"
    End With
    Set FmeCurrent = Nothing

    Set FmeCurrent = .Controls("Framebal2")
    With FmeCurrent
      
      ' Set Frame properties
      Let .Caption = "Incluir en resultado"
      Let .Width = ((.Parent.InsideWidth * 0.9) / 2) - 2.5
      Let .Height = .Parent.InsideHeight * 0.2
      Let .Left = xLeft
      Let .Top = .Parent.Controls("Framebal1").Top + .Parent.Controls("Framebal1").Height + 5

      ' Create Frame controls
      Call VWP.ViewFormAddControl(FmeCurrent, "CheckBox", "Chkbal", 13, CtlFromX:=6)
      
      ' Set Frame controls properties (including toolTipTexts and values)
      Let VecLabels = Array("Todo", "Peso [%]", "Unidades", "Finos", "Peso Complejo [g]", "Peso Complejo [OT]", "Radio de concentración", "Leyes de Cabeza Ens. y Calc.")
      Let VecCntrls1 = Array("Imprimir todos los valores", "Porcentaje en Peso", "Unidades", "Finos en Peso (Contenido Metálico)", "Peso del mineral complejo [Toneladas]", "Peso del mineral complejo [Onzas Troy]", "Radio de Concentración", "Mostrar las Leyes de Cabeza Ensayada y Calculada")
      For i = 0 To 7
        With .Controls("Chkbal" & i + 6)
          Let .Caption = VecLabels(i)
          Let .ControlTipText = VecCntrls1(i)
          Let .AutoSize = True
          Let .WordWrap = False
          Let .Left = (xLeft * 0.5) + IIf(i > 3, .Parent.Controls("Chkbal8").Left + .Parent.Controls("Chkbal8").Width, 0)
          Let j = IIf(i > 3, i - 4, i)
          Let .Top = (xLeft * 0.35) + ((.Parent.Controls("Chkbal" & j + 6).Height * 0.9) * j)
        End With
      Next i
          
    End With
    Set FmeCurrent = Nothing

    Set FmeCurrent = .Controls("Framebal3")
    With FmeCurrent
      
      ' Set Frame properties
      Let .Caption = "Dónde mostrar resultados"
      Let .Width = ((.Parent.InsideWidth * 0.9) / 2) - 2.5
      Let .Height = .Parent.InsideHeight * 0.2
      Let .Left = .Parent.Controls("Framebal2").Left + .Parent.Controls("Framebal2").Width + 5
      Let .Top = .Parent.Controls("Framebal1").Top + .Parent.Controls("Framebal1").Height + 5

      ' Create Frame controls
      Call VWP.ViewFormAddControl(FmeCurrent, "OptionButton", "Optbal", 1)
      Call VWP.ViewFormAddControl(FmeCurrent, "TextBox", "Txtbal", 8, CtlFromX:=8) ' EVENT
      
      ' Set Frame controls properties (including values)
      Let .Controls("Optbal0").Caption = "Hoja nueva"
      Let .Controls("Optbal1").Caption = "Hoja actual"

      Let .Controls("Optbal1").AutoSize = True
      Let .Controls("Optbal0").AutoSize = True
      Let .Controls("Optbal1").WordWrap = False
      Let .Controls("Optbal0").WordWrap = False

      Let .Controls("Optbal1").Left = xLeft
      Let .Controls("Optbal0").Left = xLeft
      Let .Controls("Optbal0").Top = xLeft / 1.3
      Let .Controls("Optbal1").Top = (xLeft / 1.3) + .Controls("Optbal1").Height

      Let .Controls("Optbal0").Value = True
      
      Let j = .Width - (xLeft * 2) - .Controls("Optbal1").Width
      Let .Controls("Txtbal8").Left = .Controls("Optbal1").Left + .Controls("Optbal1").Width
      Let .Controls("Txtbal8").Top = .Controls("Optbal1").Top
      Let .Controls("Txtbal8").Width = j

      ' Set Frame controls ToolTipTexts
      Let .Controls("Txtbal8").ControlTipText = "Seleccionar una celda para mostrar el resultado"

    End With
    Set FmeCurrent = Nothing
    Let VecLabels = Empty
    Let VecMaxLgt = Empty
    Let VecCntrls1 = Empty
    Let VecCntrls0 = Empty
    
    ' Set Form textbox controls dropbuttons
    Let VecLabels = Array(0, 2, 5, 8)
    For i = 0 To UBound(VecLabels)
      With .Controls("Txtbal" & VecLabels(i))
        Let .SpecialEffect = fmSpecialEffectRaised
        Let .DropButtonStyle = fmDropButtonStyleReduce
        Let .ShowDropButtonWhen = fmShowDropButtonWhenAlways
      End With
    Next i
    Let VecLabels = Empty

    ' Set Form controls values
    Let .Controls("Cbobal0").List = Array("Convencional (sólo para balances de dos productos)", "Por determinantes (Cramer)", "Por Matriz inversa")
    Let .Controls("Txtbal1").Text = "A"
    Set RES = New AppResources_current
    Set RES.ErrorHandler = EHGLOBAL
    Set MDL = New ModelExcliqliteDatasheet
    Set MDL.ErrorHandler = EHGLOBAL
    Let VecLabels = RES.ArrayToBaseZero(MDL.MGet(eqlMdlTblChemicalelements, eqlMdlArray, "Símbolo"))
    Let .Controls("Cbobal1").List = VecLabels
    Let VecLabels = Empty
    Let VecLabels = RES.ArrayToBaseZero(MDL.MGet(eqlMdlTblChemicalunits, eqlMdlArray))
    Let .Controls("Cbobal2").List = VecLabels 'Array("%", "DM", "g/T", "g/L")
    Let VecLabels = RES.ArrayToBaseZero(MDL.MGet(eqlMdlTblChemicalelementscx, eqlMdlArray, "Símbolo"))
    Let .Controls("Cbobal3").List = VecLabels
    Let VecLabels = Empty
    Let VecLabels = RES.ArrayToBaseZero(MDL.MGet(eqlMdlTblChemicalunitscx, eqlMdlArray))
    Let .Controls("Cbobal4").List = VecLabels 'Array("DM", "g/T", "g/L")
    Set MDL = Nothing
    Set RES = Nothing
    For i = 0 To 4
      Let .Controls("Cbobal" & i).ListIndex = 0
      'Let .Controls("Chkbal" & i + 1).Value = Not (i = 2)
      If i = 0 Or i = 3 Then Let .Controls("Chkbal" & i + 1).Caption = "Seleccionar minerales"
      If i = 1 Or i = 4 Then Let .Controls("Chkbal" & i + 1).Caption = "Seleccionar unidades"
    Next i
    Let VecLabels = Array(1, 3, 4, 6, 7, 8)
    For i = 0 To UBound(VecLabels)
      Let .Controls("Txtbal" & VecLabels(i)).ForeColor = VBA.vbGrayText
      Let .Controls("Txtbal" & VecLabels(i)).Locked = Not (i = UBound(VecLabels))
    Next i
    Let VecLabels = Empty
    For i = 1 To 4
      Let .Controls("Cbobal" & i).Enabled = False
    Next i
    For i = 0 To 7
      Let .Controls("Btnbal" & i).Caption = IIf(i Mod 2 = 0, VBA.ChrW(&H2A31), VBA.ChrW(&H2716))
      Let .Controls("Btnbal" & i).Enabled = False
    Next i
    Let .Controls("Chkbal6").Value = True

    ' Set Tabindex
    Let VecLabels = Array("Lblbal0", "Cbobal0", "Lblbal1", "Txtbal0", "Chkbal0", "Txtbal1", "Lblbal2", "Txtbal2", "Chkbal1", "Cbobal1", "Btnbal0", "Btnbal1", "Txtbal3", "Chkbal2", "Cbobal2", "Btnbal2", "Btnbal3", "Txtbal4", "Chkbal3", "Lblbal3", "Txtbal5", "Chkbal4", "Cbobal3", "Btnbal4", "Btnbal5", "Txtbal6", "Chkbal5", "Cbobal4", "Btnbal6", "Btnbal7", "Txtbal7", "Chkbal6", "Chkbal7", "Chkbal8", "Chkbal9", "Chkbal10", "Chkbal11", "Chkbal12", "Chkbal13", "Optbal0", "Optbal1", "Txtbal8", "Btnmain0", "Btnmain1", "Btnmain2")
    For i = 0 To UBound(VecLabels)
      Let .Controls(VBA.CStr(VecLabels(i))).TabIndex = i
    Next i
    
    ' Set Hotkeys
    Let .Controls("Lblbal0").Accelerator = "m"
    Let .Controls("Lblbal1").Accelerator = "p"
    Let .Controls("Lblbal2").Accelerator = "l"
    Let .Controls("Lblbal3").Accelerator = "y"
    Let .Controls("Chkbal1").Accelerator = "n"
    Let .Controls("Chkbal2").Accelerator = "u"
    Let .Controls("Chkbal3").Accelerator = "s"
    Let .Controls("Chkbal4").Accelerator = "i"
    Let .Controls("Chkbal5").Accelerator = "d"
    Let .Controls("Chkbal6").Accelerator = "t"
    Let .Controls("Optbal0").Accelerator = "v"
    Let .Controls("Optbal1").Accelerator = "j"
    
    '3. [x] Fill Form Controls with data and Set general purpose CommandButtons
    Let k = (.InsideWidth * 0.95)
    With .Controls("Btnmain0") '("BtnCancel")
      Let .Accelerator = "c"
      Let .Cancel = True
      Let .Top = (.Parent.InsideHeight * 0.95) - (.Height)
      Let .Left = k - .Width
      Let .Caption = "Cerrar"
    End With
    With .Controls("Btnmain2") '("BtnHelp")
      Let .Caption = "Ayuda"
      Let .Accelerator = "a"
      Let .ControlTipText = "Ver ayuda del Ejercicio"
      Let .Top = .Parent.Controls("Btnmain0").Top
      Let .Left = xLeft
      Let .Tag = ControlId '& cGet
    End With
    With .Controls("Btnmain1") '("BtnEXE")
      Let .Caption = "Resolver"
      Let .Accelerator = "r"
      Let .ControlTipText = "Permite procesar los datos y mostrar los resultados en una hoja Excel."
      Let .Top = .Parent.Controls("Btnmain0").Top
      Let .Left = .Parent.Controls("Btnmain0").Left - .Width - 2
      Let .Tag = "validation" 'ControlId & "-bal"
      Let .Default = True
    End With
    
    '4. Events: Assing handler for each control
    Set BtnCancel = .Controls("Btnmain0") 'Main
    Set BtnAccept = .Controls("Btnmain1") 'Main
    Set Btn0 = .Controls("Btnmain2") 'Main
    
    Set Btn1 = .Controls("Btnbal0") ' Frame0
    Set Btn2 = .Controls("Btnbal1") ' Frame0
    Set Btn3 = .Controls("Btnbal2") ' Frame0
    Set Btn4 = .Controls("Btnbal3") ' Frame0
    Set Btn5 = .Controls("Btnbal4") ' Frame1
    Set Btn6 = .Controls("Btnbal5") ' Frame1
    Set Btn7 = .Controls("Btnbal6") ' Frame1
    Set Btn8 = .Controls("Btnbal7") ' Frame1
    
    Set Txt0 = .Controls("Txtbal0") ' Frame0
    Set Txt1 = .Controls("Txtbal2") ' Frame0
    Set Txt2 = .Controls("Txtbal5") ' Frame1
    Set Txt3 = .Controls("Txtbal8") ' Frame3
    
    Set Chk0 = .Controls("Chkbal0") ' Frame0
    Set Chk1 = .Controls("Chkbal1") ' Frame0
    Set Chk2 = .Controls("Chkbal2") ' Frame0
    Set Chk3 = .Controls("Chkbal4") ' Frame1
    Set Chk4 = .Controls("Chkbal5") ' Frame1
    Set Chk5 = .Controls("Chkbal6") ' Frame2
    
    Set Cbo0 = .Controls("Cbobal1") ' Frame0
    Set Cbo1 = .Controls("Cbobal2") ' Frame0
    Set Cbo2 = .Controls("Cbobal3") ' Frame1
    Set Cbo2 = .Controls("Cbobal4") ' Frame1
    
    '5. Give styles to controls
    Call VWP.ViewFormStyles(FRM): Set VWP = Nothing

    '6. [x] Show the finished Form
    Call Load(FRM)
    Call .Show

    '7. [x] Close form
  End With

EH:
  Let VecLabels = Empty
  Let VecMaxLgt = Empty
  Let VecCntrls1 = Empty
  Let VecCntrls0 = Empty
  Set FmeCurrent = Nothing
  Set VWP = Nothing
  Set MDL = Nothing
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Sub

'Balmet <
Private Sub View_form_balmet_OOLD(ByVal ControlId As String, ByVal FRM As FRMBasics)
'  'Main objects
'  Dim VIEWS As ViewsApp
'  Dim RES As AppResources
'  'ListObjects
'  'LO data names and other data
'  'Controls
'  Dim aLabels() As Variant
'  Dim aMaxLengths() As Variant
'  Dim aControls() As Variant
'  Dim aControlis() As Variant
'  'Standard vars
'  Dim sTitle As String, TitleForm As String
'  Dim i As Long, j As Long, k As Long, l As Long
'  Dim xLeft As Long
'  Const sFN As String = "VIEW::-FormBalmet"
'
'
'  Call Debug_assert(Not ControlId = VBA.vbNullString, sFN)
'  Call Debug_assert(Not FRM Is Nothing, sFN)
'  If ControlId = VBA.vbNullString Then GoTo EH
'  If FRM Is Nothing Then GoTo EH
'  On Error GoTo EH
'
'  '1. [x] Get data for Form Controls
'
'
'  '2. [x] Desing the Form
'  Let TitleForm = "Balance Metalúrgico"
'  Let sTitle = TitleForm & " - " & STRAPPTITLE
'
'  Set VIEWS = New ViewsApp
'  Set VIEWS.ErrorHandler = EHGLOBAL
'  Call VIEWS.ViewFormCleanner(FRM)
'  With FRM
'    'Set the form dimensions
'    Let .Height = 415
'    Let .Width = 450
'    Let .Tag = ControlId
'    Let .Caption = sTitle
'    Let xLeft = .InsideWidth * 0.05
'
'    'Build and design the form
'    Call VIEWS.ViewFormAddControl(FRM, "CommandButton", "Btnmain", 2) 'All with events (Cancel, Exe and Help)
'    Call VIEWS.ViewFormAddControl(FRM, "Frame", "Framebal", 3)
'
'    Call VIEWS.ViewFormAddControl(.Controls("Framebal0"), "Label", "Lblbal", 2)
'    Call VIEWS.ViewFormAddControl(.Controls("Framebal0"), "ToggleButton", "Tglbal", 0)
'    Call VIEWS.ViewFormAddControl(.Controls("Framebal0"), "ComboBox", "Cbobal", 5) 'EVENT 0,1
'    Call VIEWS.ViewFormAddControl(.Controls("Framebal0"), "TextBox", "Txtbal", 7) 'EVENT 1
'    Call VIEWS.ViewFormAddControl(.Controls("Framebal0"), "CheckBox", "Chkbal", 5)
'    Call VIEWS.ViewFormAddControl(.Controls("Framebal0"), "CommandButton", "Btnbal", 3) 'EVENT all
'    With .Controls("Framebal0")
'      Let .Caption = "Ingresar Datos"
'      Let .Width = .Parent.InsideWidth * 0.9
'      Let .Height = .Parent.InsideHeight * 0.4
'      Let .Left = xLeft
'      Let .Top = .Parent.InsideHeight * 0.04
'
'      Let aLabels = Array("Método de resolución:", "Peso Alimentación", "Leyes de minerales y productos:")
'      Let aControls = Array("Lblbal0", "Lblbal1", "Lblbal2")
'      Let j = xLeft / 2
'      For i = 0 To UBound(aLabels)
'        Let .Controls(aControls(i)).Caption = aLabels(i)
'        Let .Controls(aControls(i)).Left = xLeft
'        Let .Controls(aControls(i)).AutoSize = True
'        Let .Controls(aControls(i)).WordWrap = False
'        Let .Controls(aControls(i)).Top = j + ((.Controls(aControls(i)).Height * 2) * i)
'      Next i
'      Erase aLabels
'
'      Erase aLabels: Let aLabels = Array("Cbobal0", "Txtbal0", "Txtbal1")
'      Erase aControlis: Let aControlis = Array("Btnbal0", "Btnbal1", "Btnbal2", "Btnbal3")
'      Let j = .Controls("Lblbal2").Left + .Controls("Lblbal2").Width '+ xLeft
'      For i = 0 To UBound(aLabels)
'        Let .Controls(aLabels(i)).Left = j
'        Let .Controls(aLabels(i)).Top = .Controls(aControls(i)).Top
'        If i > 0 Then
'          Let .Controls(aControlis(i - 1)).Left = .Controls(aLabels(i)).Left + .Controls(aLabels(i)).Width + 1
'          Let .Controls(aControlis(i - 1)).Top = .Controls(aLabels(i)).Top
'          Let .Controls(aControlis(i - 1)).Height = .Controls(aLabels(1)).Height
'          Let .Controls(aControlis(i - 1)).Width = .Controls(aControlis(i - 1)).Height * 0.8
'          Let .Controls(aControlis(i - 1)).Caption = "_"
'          Let .Controls(aControlis(i - 1)).Font.Bold = True
'          Let .Controls(aControlis(i - 1)).ControlTipText = "Pulse para seleccionar una celda o rango de celdas"
'          Let .Controls(aControlis(i - 1)).Tag = ControlId & "-rng"
'
'          Let .Controls(aControlis(i + 1)).Left = .Controls(aControlis(i - 1)).Left + .Controls(aControlis(i - 1)).Width + 1
'          Let .Controls(aControlis(i + 1)).Top = .Controls(aControlis(i - 1)).Top
'          Let .Controls(aControlis(i + 1)).Height = .Controls(aControlis(i - 1)).Height
'          Let .Controls(aControlis(i + 1)).Width = .Controls(aControlis(i + 1)).Height * 0.9
'          Let .Controls(aControlis(i + 1)).Caption = VBA.ChrW(&H2716)
'          Let .Controls(aControlis(i + 1)).Font.Bold = True
'          Let .Controls(aControlis(i + 1)).ControlTipText = "Descartar"
'          Let .Controls(aControlis(i + 1)).Tag = ControlId & "-clr"
'        End If
'      Next i
'      Erase aLabels: Erase aControls
'      Let .Controls("Cbobal0").Width = .Width - (.Controls("Lblbal0").Width + (xLeft * 3.7)) + 1
'      Let .Controls("Cbobal0").ControlTipText = "Seleccione un método para resolver el Balance Metalúrgico"
'      Let .Controls("Cbobal1").ControlTipText = "Seleccione la unidad de ley que tendrá el mineral principal"
'
'      Erase aLabels: Let aLabels = Array("Chkbal0", "Chkbal1", "Chkbal2", "Chkbal3", "Chkbal4", "Chkbal5")
'      Erase aControls: Let aControls = Array("Txtbal2", "Txtbal3", "Txtbal4", "Txtbal5", "Txtbal6", "Txtbal7")
'      Let j = .Controls("Btnbal3").Left + .Controls("Btnbal3").Width + (xLeft / 3)
'      For i = 0 To UBound(aLabels)
'        Let .Controls(aLabels(i)).Left = j
'        Let .Controls(aLabels(i)).Top = .Controls(aControlis(0)).Top + (.Controls(aControlis(0)).Height * i)
'        Let .Controls(aLabels(i)).AutoSize = True
'        Let .Controls(aLabels(i)).ControlTipText = "Active esta casilla y cambie el nombre de la variable del cuadro de texto a una personalizada"
'        Let .Controls(aControls(i)).Left = j + (.Controls(aLabels(i)).Width * 0.7) ' + 1
'        Let .Controls(aControls(i)).Top = .Controls(aLabels(i)).Top
'        Let .Controls(aControls(i)).Width = .Controls("Cbobal0").Width / 4
'        Let .Controls(aControls(i)).MaxLength = 12
'        Let .Controls("Txtbal" & i).ControlTipText = IIf(i < 2, "Ingrese un número o varios separados por punto y coma (Ej. 1;2;3) o un rango mediante el botón", _
'                                                             "Active la casilla y cambie el nombre de la variable de este cuadro te texto a una personalizada")
'      Next i
'      Erase aLabels: Erase aControls: Erase aControlis
'      Let aLabels = Array("A", "B1", "B2", "B3", "B4", "B5")
'      For i = 0 To UBound(aLabels)
'        Let .Controls("Txtbal" & i + 2).Tag = aLabels(i)
'        Let .Controls("Txtbal" & i + 2).Value = aLabels(i)
'      Next i
'      Erase aLabels
'      Let j = (.Controls("Cbobal0").Left + .Controls("Cbobal0").Width) - (.Controls("Txtbal2").Left + .Controls("Txtbal2").Width) - 1
'      For i = 0 To 4
'        Let .Controls("Cbobal" & i + 1).Left = .Controls("Txtbal" & i + 3).Left + .Controls("Txtbal" & i + 3).Width + 1
'        Let .Controls("Cbobal" & i + 1).Top = .Controls("Txtbal" & i + 3).Top + 1
'        Let .Controls("Cbobal" & i + 1).Width = j
'        If i > 0 Then Let .Controls("Cbobal" & i + 1).List = Array("%", "g/T", "g/L")
'      Next i
'
'      Let .Controls("Tglbal0").Caption = "INICIO|fin"
'      Let .Controls("Tglbal0").Left = .Controls("Cbobal1").Left
'      Let .Controls("Tglbal0").Width = .Controls("Cbobal1").Width
'      Let .Controls("Tglbal0").Top = .Controls("Txtbal2").Top
'      Let .Controls("Tglbal0").Height = .Controls("Txtbal2").Height
'      Let .Controls("Tglbal0").ControlTipText = "Los Datos de Alimentación están al INICIO/final (DESACTIVADO|activado)"
'
'      Let .Controls("Cbobal0").List = Array("Convencional (sólo balances de dos productos)", "Por determinantes (Cramer)", "Por Matriz inversa")
'      Let .Controls("Cbobal1").List = Array("%", "DM", "g/T", "g/L")
'      Let .Controls("Cbobal0").Tag = ControlId
'      Let .Controls("Cbobal1").Tag = ControlId
'    End With
'
'    Call VIEWS.ViewFormAddControl(.Controls("Framebal1"), "Label", "Lblbal", 5, CtlFromX:=3)
'    Call VIEWS.ViewFormAddControl(.Controls("Framebal1"), "CheckBox", "Chkbal", 6, CtlFromX:=6)
'    Call VIEWS.ViewFormAddControl(.Controls("Framebal1"), "TextBox", "Txtbal", 8, CtlFromX:=8)
'    Call VIEWS.ViewFormAddControl(.Controls("Framebal1"), "CommandButton", "Btnbal", 5, CtlFromX:=4) 'EVENT all
'    Call VIEWS.ViewFormAddControl(.Controls("Framebal1"), "ComboBox", "Cbobal", 7, CtlFromX:=6)
'    With .Controls("Framebal1")
'      Let .Caption = "Complejos de metales preciosos"
'      Let .Width = .Parent.InsideWidth * 0.9
'      Let .Height = .Parent.InsideHeight * 0.2
'      Let .Left = xLeft
'      Let .Top = .Parent.Controls("Framebal0").Top + .Parent.Controls("Framebal0").Height + 5
'
'      Let .Controls("Chkbal6").Left = xLeft
'      Let .Controls("Chkbal6").Top = xLeft / 2
'      Let .Controls("Chkbal6").Caption = "Es complejo de Plata, u otro metal precioso"
'      Let .Controls("Chkbal6").AutoSize = True
'      Let .Controls("Chkbal6").WordWrap = False
'      Let .Controls("Chkbal6").ControlTipText = "Active esta casilla si su mineral es de tipo complejo e ingrese los datos correspondientes en los controles inferiores"
'
'      Let aControls = Array("Lblbal3", "Lblbal4", "Lblbal5")
'      Let aControlis = Array("Cbobal6", "Txtbal8", "Cbobal7")
'      Let aLabels = Array("Mineral:", "Leyes:", "Unidad")
'      For i = 0 To UBound(aControls)
'        Let .Controls(aControls(i)).Caption = aLabels(i)
'        Let .Controls(aControls(i)).AutoSize = True
'        Let .Controls(aControls(i)).WordWrap = False
'        Let .Controls(aControls(i)).Left = xLeft + (((.Width - (xLeft * 2)) / 3) * i) + IIf(i = 2, .Controls(aControlis(0)).Height * 2, 0)
'        Let .Controls(aControls(i)).Top = (.Controls("Chkbal6").Top + (xLeft / 3)) + .Controls(aControls(i)).Height
'        Let .Controls(aControlis(i)).Left = .Controls(aControls(i)).Left
'        Let .Controls(aControlis(i)).Top = (.Controls(aControls(i)).Top + .Controls(aControls(i)).Height + 2)
'      Next i
'      Erase aControls: Erase aControlis: Erase aLabels
'
'      Let .Controls("Btnbal4").Left = .Controls("Txtbal8").Left + .Controls("Txtbal8").Width + 1
'      Let .Controls("Btnbal4").Top = .Controls("Txtbal8").Top
'      Let .Controls("Btnbal4").Height = .Controls("Txtbal8").Height
'      Let .Controls("Btnbal4").Width = .Controls("Btnbal4").Height
'      Let .Controls("Btnbal4").Caption = "_"
'      Let .Controls("Btnbal4").ControlTipText = "Pulse para seleccionar una celda"
'      Let .Controls("Btnbal4").Font.Bold = True
'      Let .Controls("Btnbal4").Tag = ControlId & "-rng"
'
'      Let .Controls("Btnbal5").Left = .Controls("Btnbal4").Left + .Controls("Btnbal4").Width + 1
'      Let .Controls("Btnbal5").Top = .Controls("Btnbal4").Top
'      Let .Controls("Btnbal5").Height = .Controls("Btnbal4").Height
'      Let .Controls("Btnbal5").Width = .Controls("Btnbal4").Height
'      Let .Controls("Btnbal5").Caption = VBA.ChrW(&H2716)
'      Let .Controls("Btnbal5").ControlTipText = "Descartar"
'      Let .Controls("Btnbal5").Font.Bold = True
'      Let .Controls("Btnbal5").Tag = ControlId & "-clr"
'
'      Let .Controls("Cbobal6").List = Array("Ag", "Au", "Pt")
'      Let .Controls("Cbobal7").List = Array("DM", "g/T", "g/L")
'      Let .Controls("Txtbal8").ControlTipText = "Ingrese las leyes del elemento en cada producto (Alim., Conc. y Relave) separados por punto y comas (Ej. 1;2;3) o un rango mediante el botón"
'      Let .Controls("Cbobal6").ControlTipText = "Seleccione un elemento metal precioso"
'      Let .Controls("Cbobal7").ControlTipText = "Seleccione la unidad de ley del elemento"
'    End With
'
'    Call VIEWS.ViewFormAddControl(.Controls("Framebal2"), "CheckBox", "Chkbal", 14, CtlFromX:=7)
'    With .Controls("Framebal2")
'      Let .Caption = "Incluir en resultado"
'      Let .Width = ((.Parent.InsideWidth * 0.9) / 2) - 2.5
'      Let .Height = .Parent.InsideHeight * 0.2
'      Let .Left = xLeft
'      Let .Top = .Parent.Controls("Framebal1").Top + .Parent.Controls("Framebal1").Height + 5
'
'      Let aControls = Array("Chkbal7", "Chkbal8", "Chkbal9", "Chkbal10", "Chkbal11", "Chkbal12", "Chkbal13", "Chkbal14")
'      Let aLabels = Array("Todo", "Peso [%]", "Unidades", "Finos", "Peso Complejo [T]", "Peso Complejo [OT]", "Ratio", "Leyes de Cabeza")
'      Let aControlis = Array("Imprimir todos los valores", "Porcentaje en Peso", "Unidades", "Finos en Peso (Contenido Metálico)", "Peso del mineral complejo [Toneladas]", "Peso del mineral complejo [Onzas Troy]", "Radio de Concentración", "Mostrar las Leyes de Cabeza Ensayada y Calculada (Sólo en Complejos)")
'      For i = 0 To UBound(aControls)
'        Let .Controls(aControls(i)).Caption = aLabels(i)
'        Let .Controls(aControls(i)).ControlTipText = aControlis(i)
'        Let .Controls(aControls(i)).AutoSize = True
'        Let .Controls(aControls(i)).WordWrap = False
'        Let .Controls(aControls(i)).Left = xLeft + IIf(i > 3, .Controls(aControls(2)).Left + .Controls(aControls(2)).Width, 0)
'        Let j = IIf(i > 3, i - 4, i)
'        Let .Controls(aControls(i)).Top = (xLeft / 2) + ((.Controls(aControls(j)).Height * 0.7) * j)
'      Next i
'      Erase aControls: Erase aControlis: Erase aLabels
'      Let .Controls("Chkbal7").Value = True
'    End With
'
'    Call VIEWS.ViewFormAddControl(.Controls("Framebal3"), "OptionButton", "Optbal", 1)
'    Call VIEWS.ViewFormAddControl(.Controls("Framebal3"), "TextBox", "Txtbal", 9, CtlFromX:=9)
'    Call VIEWS.ViewFormAddControl(.Controls("Framebal3"), "CommandButton", "Btnbal", 6, CtlFromX:=6) 'EVENT
'    With .Controls("Framebal3")
'      Let .Caption = "Dónde mostrar resultados"
'      Let .Width = ((.Parent.InsideWidth * 0.9) / 2) - 2.5
'      Let .Height = .Parent.InsideHeight * 0.2
'      Let .Left = .Parent.Controls("Framebal2").Left + .Parent.Controls("Framebal2").Width + 5
'      Let .Top = .Parent.Controls("Framebal1").Top + .Parent.Controls("Framebal1").Height + 5
'
'      Let .Controls("Optbal0").Caption = "Hoja nueva"
'      Let .Controls("Optbal1").Caption = "Hoja actual"
'
'      Let .Controls("Optbal1").AutoSize = True
'      Let .Controls("Optbal0").AutoSize = True
'      Let .Controls("Optbal1").WordWrap = False
'      Let .Controls("Optbal0").WordWrap = False
'
'      Let .Controls("Optbal1").Left = xLeft
'      Let .Controls("Optbal0").Left = xLeft
'      Let .Controls("Optbal0").Top = xLeft / 1.3
'      Let .Controls("Optbal1").Top = (xLeft / 1.3) + .Controls("Optbal1").Height
'
'      Let .Controls("Optbal0").Value = True
'
'      Let .Controls("Txtbal9").Left = .Controls("Optbal1").Left + .Controls("Optbal1").Width
'      Let .Controls("Txtbal9").Top = .Controls("Optbal1").Top
'      Let .Controls("Btnbal6").Left = .Controls("Txtbal9").Left + .Controls("Txtbal9").Width + 1
'      Let .Controls("Btnbal6").Top = .Controls("Txtbal9").Top
'      Let .Controls("Btnbal6").Height = .Controls("Txtbal9").Height
'      Let .Controls("Btnbal6").Width = .Controls("Txtbal9").Height
'      Let .Controls("Btnbal6").Caption = "_"
'      Let .Controls("Btnbal6").Font.Bold = True
'      Let .Controls("Btnbal6").ControlTipText = "Click para seleccionar una celda"
'      Let .Controls("Btnbal6").Tag = ControlId & "-rng"
'    End With
'    Erase aLabels
'    Erase aMaxLengths
'    Erase aControlis
'    Erase aControls
'
'    For i = 0 To 7
'      Let .Controls("Cbobal" & i).ListIndex = 0
'    Next i
'
'    '3. [x] Fill Form Controls with data and Set general purpose CommandButtons
'    Let k = (.InsideWidth * 0.95)
'    With .Controls("Btnmain0") '("BtnCancel")
'      Let .Accelerator = "c"
'      Let .Cancel = True
'      Let .Top = (.Parent.InsideHeight * 0.95) - (.Height)
'      Let .Left = k - .Width
'      Let .Caption = "Cerrar"
'    End With
'    With .Controls("Btnmain2") '("BtnHelp")
'      Let .Caption = "Ayuda"
'      Let .Accelerator = "a"
'      Let .ControlTipText = "Ver ayuda del Ejercicio"
'      Let .Top = .Parent.Controls("Btnmain0").Top
'      Let .Left = xLeft
'      Let .Tag = ControlId '& cGet
'    End With
'    With .Controls("Btnmain1") '("BtnEXE")
'      Let .Caption = "Resolver"
'      Let .Accelerator = "r"
'      Let .ControlTipText = "Permite procesar los datos y mostrar los resultados en una hoja Excel."
'      Let .Top = .Parent.Controls("Btnmain0").Top
'      Let .Left = .Parent.Controls("Btnmain0").Left - .Width - 2
'      Let .Tag = ControlId & "-bal"
'      Let .Default = True
'    End With
'
'    '4. Assing handler for each button
'    Set BtnCancel = .Controls("Btnmain0") 'Main
'    Set BtnAccept = .Controls("Btnmain1") 'Main
'    Set Btn0 = .Controls("Btnmain2") 'Main
'    Set Btn1 = .Controls("Btnbal0") 'Frame0
'    Set Btn2 = .Controls("Btnbal1") 'Frame0
'    Set Btn3 = .Controls("Btnbal2") 'Frame0
'    Set Btn4 = .Controls("Btnbal3") 'Frame0
'    Set Btn5 = .Controls("Btnbal4") 'Frame1
'    Set Btn6 = .Controls("Btnbal5") 'Frame1
'    Set Btn7 = .Controls("Btnbal6") 'Frame2
'
'    Set Cbo0 = .Controls("Cbobal0") 'Frame0
'    Set Cbo1 = .Controls("Cbobal1") 'Frame0
'
'    Set Txt0 = .Controls("Txtbal1") 'Frame0
'
'    '5. Give styles to controls
'    Call VIEWS.ViewFormStyles(FRM): Set VIEWS = Nothing
'
'    '6. [x] Show the finished Form
'    Call Load(FRM)
'    Call .Show
'
'    '7. [x] Close form
'  End With
'
'EH:
'  Erase aLabels
'  Erase aMaxLengths
'  Erase aControlis
'  Erase aControls
'  Set VIEWS = Nothing
'  Set RES = Nothing
'  Call EHGLOBAL.ErrorHandlerRaise(sFN)
End Sub

Private Sub View_form_balmeteco(ByVal ControlId As String, ByVal FRM As FRMBasics)
  'Main objects
  Dim VIEWS As ViewsApp
  Dim RES As AppResources
  'ListObjects
  'LO data names and other data
  'Controls
  Dim aLabels() As Variant
  Dim aMaxLengths() As Variant
  Dim aControls() As Variant
  Dim aControlis() As Variant
  'Standard vars
  Dim sTitle As String, TitleForm As String
  Dim i As Long, j As Long, k As Long, l As Long
  Dim xLeft As Long
  Const sFN As String = "VIEW::-FormBalmeteco"

  #If Debugging Then
    Call Debug_assert(Not ControlId = VBA.vbNullString, sFN)
    Call Debug_assert(Not FRM Is Nothing, sFN)
  #End If
  
  If ControlId = VBA.vbNullString Then GoTo EH
  If FRM Is Nothing Then GoTo EH
  On Error GoTo EH
  
  '1. [x] Get data for Form Controls


  '2. [x] Desing the Form
  Let TitleForm = "Balance Metalúrgico"
  Let sTitle = TitleForm & " - " & STRAPPTITLE
  
  Set VIEWS = New ViewsApp
  Set VIEWS.ErrorHandler = EHGLOBAL
  Call VIEWS.ViewFormCleanner(FRM)
  With FRM
    'Set the form dimensions
    Let .Height = 415
    Let .Width = 450
    Let .Tag = ControlId
    Let .Caption = sTitle
    Let xLeft = .InsideWidth * 0.05

    'Build and design the form
    Call VIEWS.ViewFormAddControl(FRM, "CommandButton", "Btnmain", 2) 'All with events (Cancel, Exe and Help)
    Call VIEWS.ViewFormAddControl(FRM, "Frame", "Framebal", 3)

    Call VIEWS.ViewFormAddControl(.Controls("Framebal0"), "Label", "Lblbal", 5)
    Call VIEWS.ViewFormAddControl(.Controls("Framebal0"), "ComboBox", "Cbobal", 2) 'EVENT 0,2
    Call VIEWS.ViewFormAddControl(.Controls("Framebal0"), "TextBox", "Txtbal", 7)
    Call VIEWS.ViewFormAddControl(.Controls("Framebal0"), "CheckBox", "Chkbal", 3)
    Call VIEWS.ViewFormAddControl(.Controls("Framebal0"), "CommandButton", "Btnbal", 9) 'EVENT all
    With .Controls("Framebal0")
      Let .Caption = "Ingresar Datos"
      Let .Width = .Parent.InsideWidth * 0.9
      Let .Height = .Parent.InsideHeight * 0.4
      Let .Left = xLeft
      Let .Top = .Parent.InsideHeight * 0.04

      Let aLabels = Array("Objetivo:", "Método de resolución:", "Alimentación [T]:", "Peso Producto/s (Concentrado)[T]:", "Ley Producto/s (Concentrado):", "Recuperación[%] (opcional):")
      Let aControls = Array("Lblbal0", "Lblbal1", "Lblbal2", "Lblbal3", "Lblbal4", "Lblbal5")
      Let j = xLeft / 1.3
      For i = 0 To UBound(aLabels)
        Let .Controls(aControls(i)).Caption = aLabels(i)
        Let .Controls(aControls(i)).Left = xLeft
        Let .Controls(aControls(i)).AutoSize = True
        Let .Controls(aControls(i)).WordWrap = False
        Let .Controls(aControls(i)).Top = j + ((.Controls(aControls(i)).Height * 2) * i)
      Next i
      Erase aLabels
      
      Let .Controls("Cbobal0").Width = .Width - (.Controls("Lblbal3").Left + .Controls("Lblbal3").Width + xLeft)
      Let .Controls("Cbobal1").Width = .Controls("Cbobal0").Width
      Let .Controls("Cbobal0").ControlTipText = "Seleccione un método para resolver el Balance Metalúrgico"
      Let .Controls("Cbobal1").ControlTipText = "Seleccione la unidad de ley que tendrá el mineral principal"
      
      Erase aLabels: Let aLabels = Array("Cbobal0", "Cbobal1", "Txtbal0", "Txtbal1", "Txtbal2", "Txtbal3")
      Erase aControlis: Let aControlis = Array("Btnbal0", "Btnbal1", "Btnbal2", "Btnbal3", "Btnbal4", "Btnbal5", "Btnbal6", "Btnbal7")
      Let j = .Controls("Lblbal3").Left + .Controls("Lblbal3").Width '+ xLeft
      For i = 0 To UBound(aLabels)
        Let .Controls(aLabels(i)).Left = j
        Let .Controls(aLabels(i)).Top = .Controls(aControls(i)).Top
        If i > 1 Then
          Let .Controls(aLabels(i)).Width = .Controls("Cbobal0").Width / 4

          Let k = i - 2
          Let .Controls(aControlis(k)).Left = .Controls(aLabels(i)).Left + .Controls(aLabels(i)).Width + 1
          Let .Controls(aControlis(k)).Top = .Controls(aLabels(i)).Top
          Let .Controls(aControlis(k)).Height = .Controls(aLabels(1)).Height
          Let .Controls(aControlis(k)).Width = .Controls(aControlis(k)).Height * 0.8
          Let .Controls(aControlis(k)).Caption = VBA.ChrW(&H2583) '"_"
          Let .Controls(aControlis(k)).Font.Bold = True
          Let .Controls(aControlis(k)).ControlTipText = "Pulse para seleccionar una celda o rango de celdas"
          Let .Controls(aControlis(k)).Tag = ControlId & "-rng"
          Let l = i + 2
          Let .Controls(aControlis(l)).Left = .Controls(aControlis(k)).Left + .Controls(aControlis(k)).Width + 1
          Let .Controls(aControlis(l)).Top = .Controls(aControlis(k)).Top
          Let .Controls(aControlis(l)).Height = .Controls(aControlis(k)).Height
          Let .Controls(aControlis(l)).Width = .Controls(aControlis(l)).Height * 0.9
          Let .Controls(aControlis(l)).Caption = VBA.ChrW(&H2716)
          Let .Controls(aControlis(l)).Font.Bold = True
          Let .Controls(aControlis(l)).ControlTipText = "Descartar"
          Let .Controls(aControlis(l)).Tag = ControlId & "-clr"
        End If
      Next i
      Erase aLabels: Erase aControls

      Erase aLabels: Let aLabels = Array("Chkbal0", "Chkbal1", "Chkbal2", "Chkbal3")
      Erase aControls: Let aControls = Array("Txtbal4", "Txtbal5", "Txtbal6", "Txtbal7")
      Let j = .Controls("Btnbal4").Left + .Controls("Btnbal4").Width + (xLeft / 3)
      For i = 0 To UBound(aLabels)
        Let .Controls(aLabels(i)).Left = j
        Let .Controls(aLabels(i)).Top = .Controls(aControlis(i)).Top
        Let .Controls(aLabels(i)).AutoSize = True
        Let .Controls(aLabels(i)).ControlTipText = "Active esta casilla y cambie el valor del cuadro de texto a uno personalizado"
        Let .Controls(aControls(i)).Left = j + (.Controls(aLabels(i)).Width * 0.7) ' + 1
        Let .Controls(aControls(i)).Top = .Controls(aLabels(i)).Top
        Let .Controls(aControls(i)).Width = .Controls("Cbobal0").Width / 5.5
        Let .Controls(aControls(i)).MaxLength = 12
        Let .Controls(aControls(i)).ControlTipText = "Active la casilla y cambie el valor de este cuadro de texto a uno personalizado"
        Let .Controls("Txtbal" & i).ControlTipText = "Ingrese un número o varios separados por punto y coma (Ej. 1;2;3) o un rango mediante el botón"
      Next i
      Erase aLabels: Erase aControls: Erase aControlis

      Let aLabels = Array("A", "B", "%", "%R")
      For i = 0 To UBound(aLabels)
        Let .Controls("Txtbal" & i + 4).Value = aLabels(i)
        Let .Controls("Txtbal" & i + 4).Tag = aLabels(i)
      Next i
      Erase aLabels
      Let .Controls("Txtbal7").Enabled = False
      Let .Controls("Chkbal3").Visible = False
      
      Let j = (.Controls("Cbobal0").Left + .Controls("Cbobal0").Width) - (.Controls("Txtbal4").Left + .Controls("Txtbal4").Width) - 1
      Let .Controls("Cbobal2").Left = .Controls("Txtbal6").Left + .Controls("Txtbal6").Width + 1
      Let .Controls("Cbobal2").Top = .Controls("Txtbal6").Top ' + 1
      Let .Controls("Cbobal2").Width = j * 0.5
      
      Let .Controls("Btnbal8").Caption = VBA.ChrW(&H271A)
      Let .Controls("Btnbal8").Font.Bold = True
      Let .Controls("Btnbal8").ControlTipText = "Añadir Unidad"
      Let .Controls("Btnbal8").Left = .Controls("Cbobal2").Left + .Controls("Cbobal2").Width + 1
      Let .Controls("Btnbal8").Top = .Controls("Cbobal2").Top
      Let .Controls("Btnbal8").Height = .Controls("Cbobal2").Height
      Let .Controls("Btnbal8").Width = .Controls("Btnbal8").Height

      Let .Controls("Btnbal9").Caption = VBA.ChrW(&H2716)
      Let .Controls("Btnbal9").Font.Bold = True
      Let .Controls("Btnbal9").ControlTipText = "Limpiar contenido del cuadro de texto"
      Let .Controls("Btnbal9").Left = .Controls("Btnbal8").Left + .Controls("Btnbal8").Width + 1
      Let .Controls("Btnbal9").Top = .Controls("Cbobal2").Top
      Let .Controls("Btnbal9").Height = .Controls("Cbobal2").Height
      Let .Controls("Btnbal9").Width = .Controls("Btnbal9").Height
      

      Let .Controls("Cbobal0").List = Array("Obtener Ley de Alimentación (Cabeza)", "Obtener Ley de productos")
      Let .Controls("Cbobal1").List = Array("Convencional (sólo balances de dos productos)", "Por determinantes (Cramer)", "Por Matriz inversa")
      Let .Controls("Cbobal2").List = Array("%", "DM", "g/T", "g/L")
      Let .Controls("Cbobal2").ControlTipText = "Seleccione una Unidad y pulse el botón + para añadir al cuadro de texto"
      Let .Controls("Cbobal0").Tag = ControlId
      Let .Controls("Cbobal1").Tag = ControlId
    End With

    Call VIEWS.ViewFormAddControl(.Controls("Framebal1"), "Label", "Lblbal", 9, CtlFromX:=6)
    Call VIEWS.ViewFormAddControl(.Controls("Framebal1"), "ComboBox", "Cbobal", 4, CtlFromX:=3) 'EVENT 3
    Call VIEWS.ViewFormAddControl(.Controls("Framebal1"), "TextBox", "Txtbal", 9, CtlFromX:=8)
    Call VIEWS.ViewFormAddControl(.Controls("Framebal1"), "CheckBox", "Chkbal", 4, CtlFromX:=4)
    Call VIEWS.ViewFormAddControl(.Controls("Framebal1"), "CommandButton", "Btnbal", 13, CtlFromX:=10) 'EVENT all
    With .Controls("Framebal1")
      Let .Caption = "Complejos de metales preciosos"
      Let .Width = .Parent.InsideWidth * 0.9
      Let .Height = .Parent.InsideHeight * 0.2
      Let .Left = xLeft
      Let .Top = .Parent.Controls("Framebal0").Top + .Parent.Controls("Framebal0").Height + 5

      Let .Controls("Chkbal4").Left = xLeft
      Let .Controls("Chkbal4").Top = xLeft / 3
      Let .Controls("Chkbal4").Caption = "Es complejo de Plata, u otro metal precioso"
      Let .Controls("Chkbal4").AutoSize = True
      Let .Controls("Chkbal4").WordWrap = False
      Let .Controls("Chkbal4").ControlTipText = "Active esta casilla si el mineral es de tipo complejo e ingrese los datos correspondientes en los controles inferiores"

      Let aControls = Array("Lblbal6", "Lblbal7", "Lblbal8", "Lblbal9")
      Let aControlis = Array("Cbobal3", "Txtbal8", "Txtbal9", "Cbobal4")
      Let aLabels = Array("Mineral:", "Leyes:", "Recuperaciones:", "Unidad")
      Let k = 5
      For i = 0 To UBound(aControls)
        Let .Controls(aControls(i)).Caption = aLabels(i)
        Let .Controls(aControls(i)).AutoSize = True
        Let .Controls(aControls(i)).WordWrap = False
      Next i
      Erase aControls: Erase aControlis: Erase aLabels

      Let .Controls("Lblbal6").Left = .Controls("Chkbal4").Left
      Let .Controls("Lblbal6").Top = (.Controls("Chkbal4").Top + .Controls("Chkbal4").Height) * 1.1
      Let .Controls("Lblbal9").Left = .Controls("Chkbal4").Left
      Let .Controls("Lblbal9").Top = (.Controls("Lblbal6").Top + .Controls("Lblbal6").Height) * 1.25

      Let .Controls("Cbobal3").Left = (.Controls("Lblbal6").Left + .Controls("Lblbal6").Width) * 1.2
      Let .Controls("Cbobal4").Left = .Controls("Cbobal3").Left
      Let .Controls("Cbobal3").Top = .Controls("Lblbal6").Top
      Let .Controls("Cbobal4").Top = .Controls("Lblbal9").Top

      Let .Controls("Lblbal7").Left = .Parent.Controls("Framebal0").Controls("Cbobal0").Left
      Let .Controls("Lblbal7").Top = .Controls("Lblbal6").Top
      Let .Controls("Lblbal8").Left = .Controls("Lblbal7").Left
      Let .Controls("Lblbal8").Top = .Controls("Lblbal9").Top

      Let .Controls("Txtbal8").Left = (.Controls("Lblbal8").Left + .Controls("Lblbal8").Width)
      Let .Controls("Txtbal9").Left = .Controls("Txtbal8").Left
      Let .Controls("Txtbal8").Top = .Controls("Lblbal7").Top
      Let .Controls("Txtbal9").Top = .Controls("Lblbal9").Top
      Let .Controls("Txtbal8").Width = .Controls("Txtbal8").Width + .Controls("Txtbal8").Height + 4
      Let .Controls("Txtbal9").Width = .Controls("Txtbal8").Width

      Let .Controls("Btnbal10").Left = .Controls("Txtbal8").Left + .Controls("Txtbal8").Width + 1
      Let .Controls("Btnbal10").Top = .Controls("Txtbal8").Top
      Let .Controls("Btnbal10").Height = .Controls("Txtbal8").Height
      Let .Controls("Btnbal10").Width = .Controls("Btnbal10").Height
      Let .Controls("Btnbal10").Caption = "_"
      Let .Controls("Btnbal10").ControlTipText = "Pulse para seleccionar una celda"
      Let .Controls("Btnbal10").Font.Bold = True
      Let .Controls("Btnbal10").Tag = ControlId & "-rng"

      Let .Controls("Btnbal11").Left = .Controls("Btnbal10").Left + .Controls("Btnbal10").Width + 1
      Let .Controls("Btnbal11").Top = .Controls("Btnbal10").Top
      Let .Controls("Btnbal11").Height = .Controls("Btnbal10").Height
      Let .Controls("Btnbal11").Width = .Controls("Btnbal10").Height
      Let .Controls("Btnbal11").Caption = VBA.ChrW(&H2716)
      Let .Controls("Btnbal11").ControlTipText = "Descartar"
      Let .Controls("Btnbal11").Font.Bold = True
      Let .Controls("Btnbal11").Tag = ControlId & "-clr"

      Let .Controls("Btnbal12").Left = .Controls("Txtbal9").Left + .Controls("Txtbal9").Width + 1
      Let .Controls("Btnbal12").Top = .Controls("Txtbal9").Top
      Let .Controls("Btnbal12").Height = .Controls("Txtbal9").Height
      Let .Controls("Btnbal12").Width = .Controls("Btnbal12").Height
      Let .Controls("Btnbal12").Caption = "_"
      Let .Controls("Btnbal12").ControlTipText = "Pulse para seleccionar una celda"
      Let .Controls("Btnbal12").Font.Bold = True
      Let .Controls("Btnbal12").Tag = ControlId & "-rng"

      Let .Controls("Btnbal13").Left = .Controls("Btnbal12").Left + .Controls("Btnbal12").Width + 1
      Let .Controls("Btnbal13").Top = .Controls("Btnbal12").Top
      Let .Controls("Btnbal13").Height = .Controls("Btnbal12").Height
      Let .Controls("Btnbal13").Width = .Controls("Btnbal12").Height
      Let .Controls("Btnbal13").Caption = VBA.ChrW(&H2716)
      Let .Controls("Btnbal13").ControlTipText = "Descartar"
      Let .Controls("Btnbal13").Font.Bold = True
      Let .Controls("Btnbal13").Tag = ControlId & "-clr"

      Let .Controls("Cbobal3").List = Array("Ag", "Au", "Pt")
      Let .Controls("Cbobal4").List = Array("DM", "g/T", "g/L")
      Let .Controls("Txtbal8").ControlTipText = "Ley del producto. Debe ser un número o un rango mediante el botón"
      Let .Controls("Txtbal9").ControlTipText = "Recuperación aproximada del producto. Debe ser un número o un rango mediante el botón"
      Let .Controls("Cbobal3").ControlTipText = "Seleccione un elemento metal precioso"
      Let .Controls("Cbobal4").ControlTipText = "Seleccione la unidad de ley del elemento"
    End With

    Call VIEWS.ViewFormAddControl(.Controls("Framebal2"), "CheckBox", "Chkbal", 12, CtlFromX:=5)
    With .Controls("Framebal2")
      Let .Caption = "Incluir en resultado"
      Let .Width = ((.Parent.InsideWidth * 0.9) / 2) - 2.5
      Let .Height = .Parent.InsideHeight * 0.2
      Let .Left = xLeft
      Let .Top = .Parent.Controls("Framebal1").Top + .Parent.Controls("Framebal1").Height + 5

      Let aControls = Array("Chkbal5", "Chkbal6", "Chkbal7", "Chkbal8", "Chkbal9", "Chkbal10", "Chkbal11", "Chkbal12")
      Let aLabels = Array("Todo", "%P", "Unidades", "Finos", "Peso Complejo", "OT Complejo", "Rc", "Leyes de Cabeza")
      Let aControlis = Array("Imprimir todos los valores", "Porcentaje en Peso", "Unidades", "Finos en Peso (Contenido Metálico)", "Peso del mineral complejo [Toneladas]", "Peso del mineral complejo [Onzas Troy]", "Radio de Concentración", "Mostrar las Leyes de Cabeza Ensayada y Calculada (Sólo en Complejos)")
      For i = 0 To UBound(aControls)
        Let .Controls(aControls(i)).Caption = aLabels(i)
        Let .Controls(aControls(i)).ControlTipText = aControlis(i)
        Let .Controls(aControls(i)).AutoSize = True
        Let .Controls(aControls(i)).WordWrap = False
        Let .Controls(aControls(i)).Left = xLeft + IIf(i > 3, .Controls(aControls(2)).Left + .Controls(aControls(2)).Width, 0)
        Let j = IIf(i > 3, i - 4, i)
        Let .Controls(aControls(i)).Top = (xLeft / 2) + ((.Controls(aControls(j)).Height * 0.7) * j)
      Next i
      Erase aControls: Erase aControlis: Erase aLabels
      Let .Controls("Chkbal5").Value = True
    End With

    Call VIEWS.ViewFormAddControl(.Controls("Framebal3"), "OptionButton", "Optbal", 1)
    Call VIEWS.ViewFormAddControl(.Controls("Framebal3"), "TextBox", "Txtbal", 10, CtlFromX:=10)
    Call VIEWS.ViewFormAddControl(.Controls("Framebal3"), "CommandButton", "Btnbal", 14, CtlFromX:=14) 'EVENT all
    With .Controls("Framebal3")
      Let .Caption = "Dónde mostrar resultados"
      Let .Width = ((.Parent.InsideWidth * 0.9) / 2) - 2.5
      Let .Height = .Parent.InsideHeight * 0.2
      Let .Left = .Parent.Controls("Framebal2").Left + .Parent.Controls("Framebal2").Width + 5
      Let .Top = .Parent.Controls("Framebal1").Top + .Parent.Controls("Framebal1").Height + 5

      Let .Controls("Optbal0").Caption = "Hoja nueva"
      Let .Controls("Optbal1").Caption = "Hoja actual"

      Let .Controls("Optbal1").AutoSize = True
      Let .Controls("Optbal0").AutoSize = True
      Let .Controls("Optbal1").WordWrap = False
      Let .Controls("Optbal0").WordWrap = False

      Let .Controls("Optbal1").Left = xLeft
      Let .Controls("Optbal0").Left = xLeft
      Let .Controls("Optbal0").Top = xLeft / 1.3
      Let .Controls("Optbal1").Top = (xLeft / 1.3) + .Controls("Optbal1").Height

      Let .Controls("Optbal0").Value = True

      Let .Controls("Txtbal10").Left = .Controls("Optbal1").Left + .Controls("Optbal1").Width
      Let .Controls("Txtbal10").Top = .Controls("Optbal1").Top
      Let .Controls("Btnbal14").Left = .Controls("Txtbal10").Left + .Controls("Txtbal10").Width + 1
      Let .Controls("Btnbal14").Top = .Controls("Txtbal10").Top
      Let .Controls("Btnbal14").Height = .Controls("Txtbal10").Height
      Let .Controls("Btnbal14").Width = .Controls("Txtbal10").Height
      Let .Controls("Btnbal14").Caption = "_"
      Let .Controls("Btnbal14").Font.Bold = True
      Let .Controls("Btnbal14").ControlTipText = "Click para seleccionar una celda"
      Let .Controls("Btnbal14").Tag = ControlId & "-rng"
    End With

    For i = 0 To 4
      Let .Controls("Cbobal" & i).ListIndex = 0
    Next i

    '3. [x] Fill Form Controls with data and Set general purpose CommandButtons
    Let k = (.InsideWidth * 0.95)
    With .Controls("Btnmain0") '("BtnCancel")
      Let .Accelerator = "c"
      Let .Cancel = True
      Let .Top = (.Parent.InsideHeight * 0.95) - (.Height)
      Let .Left = k - .Width
      Let .Caption = "Cerrar"
    End With
    With .Controls("Btnmain2") '("BtnHelp")
      Let .Caption = "Ayuda"
      Let .Accelerator = "a"
      Let .ControlTipText = "Ver ayuda del Ejercicio"
      Let .Top = .Parent.Controls("Btnmain0").Top
      Let .Left = xLeft
      Let .Tag = ControlId '& cGet
    End With
    With .Controls("Btnmain1") '("BtnEXE")
      Let .Caption = "Resolver"
      Let .Accelerator = "r"
      Let .ControlTipText = "Permite procesar los datos y mostrar los resultados en una hoja Excel."
      Let .Top = .Parent.Controls("Btnmain0").Top
      Let .Left = .Parent.Controls("Btnmain0").Left - .Width - 2
      Let .Tag = ControlId & "-bal"
      Let .Default = True
    End With
    
    '4. Assing handler for each button
    Set BtnCancel = .Controls("Btnmain0") 'Main
    Set BtnAccept = .Controls("Btnmain1") 'Main
    Set Btn0 = .Controls("Btnmain2")      'Main
    Set Btn1 = .Controls("Btnbal0") 'Frame0
    Set Btn2 = .Controls("Btnbal1") 'Frame0
    Set Btn3 = .Controls("Btnbal2") 'Frame0
    Set Btn4 = .Controls("Btnbal3") 'Frame0
    Set Btn5 = .Controls("Btnbal4") 'Frame1
    Set Btn6 = .Controls("Btnbal5") 'Frame1
    Set Btn7 = .Controls("Btnbal6") 'Frame2
    Set Btn8 = .Controls("Btnbal7") 'Frame2
    Set Btn9 = .Controls("Btnbal8") 'Frame2
    Set Btn10 = .Controls("Btnbal9") 'Frame2
    Set Btn11 = .Controls("Btnbal10") 'Frame2
    Set Btn12 = .Controls("Btnbal11") 'Frame2
    Set Btn13 = .Controls("Btnbal12") 'Frame2
    Set Btn14 = .Controls("Btnbal13") 'Frame2
    Set Btn15 = .Controls("Btnbal14") 'Frame2

    Set Cbo0 = .Controls("Cbobal0") 'Frame0
    Set Cbo1 = .Controls("Cbobal2") 'Frame0
    Set Cbo2 = .Controls("Cbobal3") 'Frame1

    Set Txt0 = .Controls("Txtbal1") 'Frame0
    Set Txt1 = .Controls("Txtbal2") 'Frame0

    '5. Give styles to controls
    Call VIEWS.ViewFormStyles(FRM): Set VIEWS = Nothing

    '6. [x] Show the finished Form
    Call Load(FRM)
    Call .Show

    '7. [x] Close form
    'Erase vControls 'Global vector of controls
  End With

EH:
  Erase aLabels
  Erase aMaxLengths
  Erase aControlis
  Erase aControls
  Set VIEWS = Nothing
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)
End Sub


'Selections UI
Private Sub View_form_balmetselection(ByVal ControlId As String)
  Const sFN As String = "VIEW::-SelectionUIBalmet"
  
  On Error GoTo EH
  If View_Controls_buttons_click(Nothing, Nothing, True, ControlId & "-getcell") Then
    Call View_Controls_buttons_click(Nothing, Nothing, True, ControlId)
  End If

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)
End Sub


'Projections UI
Private Sub View_form_projection(ByVal ControlId As String, ByVal FRM As FRMBasics)
  'Main objects
  Dim RES As AppResources
  Dim VIEWS As ViewsApp
  'ListObjects
  'LO data names and other data
  'Controls
  Dim aLabels() As Variant
  Dim aMaxLengths() As Variant
  Dim aControls() As Variant
  Dim aControlis() As Variant
  'Standard vars
  Dim bNoData As Boolean, bTables As Boolean
  Dim sTitle As String, sTable As String, TitleForm As String
  Dim i As Long, j As Long, k As Long, l As Long
  Dim xLeft As Long
  Dim Aux0 As String, Aux1 As String, Aux2 As String
  
  Const cLeft As Single = 10
  Const sFN As String = "VIEW::-form_projection"

  #If Debugging Then
    Call Debug_assert(Not ControlId = VBA.vbNullString, sFN)
    Call Debug_assert(Not FRM Is Nothing, sFN)
  #End If
  
  If ControlId = VBA.vbNullString Then GoTo EH
  If FRM Is Nothing Then GoTo EH
  On Error GoTo EH
  
  'Get data for controls

  'Set Form basic datas
  Let TitleForm = "Proyección Proyectos"
  Let sTitle = TitleForm & " - " & STRAPPTITLE
  
  Set VIEWS = New ViewsApp
  Set VIEWS.ErrorHandler = EHGLOBAL
  Call VIEWS.ViewFormCleanner(FRM)
  With FRM
    'Set the form dimensions
    Let .Height = 415
    Let .Width = 450
    Let .Tag = ControlId
    Let .Caption = sTitle
    Let xLeft = .InsideWidth * 0.05

    'Build and design the form
    Call VIEWS.ViewFormAddControl(FRM, "CommandButton", "Btnmain", 2) 'All with events (Cancel, Exe and Help)
    Call VIEWS.ViewFormAddControl(FRM, "Frame", "Frame", 3)

    Call VIEWS.ViewFormAddControl(.Controls("Frame0"), "Label", "Lblpro", 4)
    Call VIEWS.ViewFormAddControl(.Controls("Frame0"), "ComboBox", "Cbopro", 2) 'EVENT 0
    Call VIEWS.ViewFormAddControl(.Controls("Frame0"), "TextBox", "Txtpro", 6)
    Call VIEWS.ViewFormAddControl(.Controls("Frame0"), "SpinButton", "Spnpro", 1) 'EVENT all
    Call VIEWS.ViewFormAddControl(.Controls("Frame0"), "CommandButton", "Btnpro", 7) 'EVENT all
    With .Controls("Frame0")
      Let .Caption = "Ingresar Datos"
      Let .Width = .Parent.InsideWidth * 0.9
      Let .Height = .Parent.InsideHeight * 0.4
      Let .Left = xLeft
      Let .Top = .Parent.InsideHeight * 0.04
      
      Let aLabels = Array("Inversión Inicial:", "Peso Fino Mineral:", "Cotización Mineral:", "Valor Residual:", "Porcentaje estimado (Impuestos, Egresos, etc.) sobre ingreso bruto:")
      For i = 0 To UBound(aLabels) - 1
        Let Aux0 = "Lblpro" & i
        Let .Controls(Aux0).Caption = aLabels(i)
        Let .Controls(Aux0).Left = xLeft
        Let .Controls(Aux0).AutoSize = True
        Let .Controls(Aux0).WordWrap = False
        Let .Controls(Aux0).Top = (xLeft / 1.2) + ((.Controls(Aux0).Height * 2) * i)
      Next i
      For i = 2 To UBound(aLabels)
        Let Aux0 = "Lblpro" & i
        Let .Controls(Aux0).Caption = aLabels(i)
        Let .Controls(Aux0).Left = xLeft
        Let .Controls(Aux0).AutoSize = True
        Let .Controls(Aux0).WordWrap = False
        Let .Controls(Aux0).Top = (.Controls("Lblpro1").Top) + ((.Controls(Aux0).Height * 2) * i) * IIf(i = UBound(aLabels), 1.1, 1)
      Next i
      Erase aLabels

      Let j = .Width - (xLeft * 2.5) - .Controls("Lblpro2").Width
      For i = 0 To 3
        Let Aux0 = "Txtpro" & i
        Let Aux1 = "Btnpro" & i
        Let .Controls(Aux0).Left = xLeft + .Controls("Lblpro2").Width + (xLeft / 2)
        Let .Controls(Aux0).Top = .Controls("Lblpro" & i).Top
        Let .Controls(Aux0).Width = j / 2.5
        Let .Controls(Aux0).ControlTipText = "Ingrese un número o un conjunto de número separados por punto y coma (;)"
        Let .Controls(Aux1).Caption = VBA.ChrW(&H2583)
        Let .Controls(Aux1).ControlTipText = "Pulse para seleccionar un rango o una celda"
        Let .Controls(Aux1).Left = .Controls(Aux0).Left + .Controls(Aux0).Width + 1
        Let .Controls(Aux1).Top = .Controls(Aux0).Top
        Let .Controls(Aux1).Height = .Controls(Aux0).Height
        Let .Controls(Aux1).Width = .Controls(Aux1).Height
      Next i
      Let .Controls("Txtpro2").ControlTipText = "Ingrese cotizaciones separadas por punto y coma (;), para varios grupos, sepárelos con diagonal (/)"
      
      Let .Controls("Cbopro0").ControlTipText = "Seleccione una unidad de moneda"
      Let .Controls("Cbopro0").Left = .Controls("Btnpro0").Left + .Controls("Btnpro0").Width + 3
      Let .Controls("Cbopro0").Top = .Controls("Btnpro0").Top
      Let .Controls("Cbopro0").Width = .Width / 2.8 '(j - .Controls("Btnpro0").Width) / 4
            
      Let .Controls("Txtpro4").Locked = True
      Let .Controls("Txtpro4").Left = .Controls("Btnpro1").Left + .Controls("Btnpro1").Width + 3
      Let .Controls("Txtpro4").Top = .Controls("Btnpro1").Top
      Let .Controls("Txtpro4").Width = .Controls("Cbopro0").Width * 0.4
            
      Let .Controls("Cbopro1").ControlTipText = "Seleccione el mineral deseado para insertarlo en la lista"
      Let .Controls("Cbopro1").Left = .Controls("Txtpro4").Left + .Controls("Txtpro4").Width + 1
      Let .Controls("Cbopro1").Top = .Controls("Txtpro4").Top
      Let .Controls("Cbopro1").Width = .Controls("Cbopro0").Width * 0.323 ' - (xLeft * 0.5)

      Let .Controls("Btnpro4").Caption = VBA.ChrW(&H2716)
      Let .Controls("Btnpro4").ControlTipText = "Eliminar último"
      Let .Controls("Btnpro4").Left = .Controls("Cbopro1").Left + .Controls("Cbopro1").Width + 1
      Let .Controls("Btnpro4").Top = .Controls("Cbopro1").Top
      Let .Controls("Btnpro4").Height = .Controls("Cbopro1").Height ' / 2 ' - (xLeft * 0.5)
      Let .Controls("Btnpro4").Width = .Controls("Btnpro4").Height ' / 2 ' - (xLeft * 0.5)

      Let .Controls("Btnpro5").Caption = VBA.ChrW(&H2716)
      Let .Controls("Btnpro5").ControlTipText = "Eliminar todos los minerales insertados"
      Let .Controls("Btnpro5").Font.Bold = True
      Let .Controls("Btnpro5").Left = .Controls("Btnpro4").Left + .Controls("Btnpro4").Width + 1
      Let .Controls("Btnpro5").Top = .Controls("Btnpro4").Top
      Let .Controls("Btnpro5").Height = .Controls("Btnpro4").Height ' / 2 ' - (xLeft * 0.5)
      Let .Controls("Btnpro5").Width = .Controls("Btnpro5").Height ' / 2 ' - (xLeft * 0.5)

      Let .Controls("Txtpro5").Locked = True
      Let .Controls("Txtpro5").Left = (.Controls("Btnpro2").Left + .Controls("Btnpro2").Width + 3)
      Let .Controls("Txtpro5").Top = .Controls("Btnpro2").Top - .Controls("Btnpro2").Height - 2
      Let .Controls("Txtpro5").Width = .Controls("Txtpro4").Width
            
      Let .Controls("Cbopro2").ControlTipText = "Seleccione la unidad del mineral deseado para insertarlo en la lista"
      Let .Controls("Cbopro2").Left = .Controls("Txtpro5").Left + .Controls("Txtpro5").Width + 1
      Let .Controls("Cbopro2").Top = .Controls("Txtpro5").Top
      Let .Controls("Cbopro2").Width = .Controls("Cbopro1").Width

      Let .Controls("Btnpro6").Caption = VBA.ChrW(&H2716)
      Let .Controls("Btnpro6").ControlTipText = "Eliminar último"
      Let .Controls("Btnpro6").Left = .Controls("Cbopro2").Left + .Controls("Cbopro2").Width + 1
      Let .Controls("Btnpro6").Top = .Controls("Cbopro2").Top
      Let .Controls("Btnpro6").Height = .Controls("Cbopro2").Height ' / 2 ' - (xLeft * 0.5)
      Let .Controls("Btnpro6").Width = .Controls("Btnpro6").Height ' / 2 ' - (xLeft * 0.5)

      Let .Controls("Btnpro7").Caption = VBA.ChrW(&H2716)
      Let .Controls("Btnpro7").ControlTipText = "Eliminar todas las unidades insertadas"
      Let .Controls("Btnpro7").Font.Bold = True
      Let .Controls("Btnpro7").Left = .Controls("Btnpro6").Left + .Controls("Btnpro6").Width + 1
      Let .Controls("Btnpro7").Top = .Controls("Btnpro6").Top
      Let .Controls("Btnpro7").Height = .Controls("Btnpro6").Height ' / 2 ' - (xLeft * 0.5)
      Let .Controls("Btnpro7").Width = .Controls("Btnpro7").Height ' / 2 ' - (xLeft * 0.5)

      Let i = .Controls("Txtpro6").Height
      Let j = .Width - (xLeft * 2.2) - .Controls("Lblpro4").Width
      Let .Controls("Txtpro6").Left = .Controls("Lblpro4").Left + .Controls("Lblpro4").Width '+ 1
      Let .Controls("Txtpro6").Top = .Controls("Lblpro4").Top * 0.97
      Let .Controls("Txtpro6").Width = j - i
      Let j = .Controls("Txtpro6").Left + .Controls("Txtpro6").Width
      Let .Controls("Txtpro6").TextAlign = fmTextAlignRight
      Let .Controls("Spnpro0").Left = j + 1
      Let .Controls("Spnpro0").Height = i
      Let .Controls("Spnpro0").Top = .Controls("Txtpro6").Top
      Let .Controls("Spnpro0").Width = i / 2
      Let .Controls("Spnpro1").Left = .Controls("Spnpro0").Left + .Controls("Spnpro0").Width
      Let .Controls("Spnpro1").Height = i
      Let .Controls("Spnpro1").Top = .Controls("Spnpro0").Top
      Let .Controls("Spnpro1").Width = i / 2
      Let .Controls("Spnpro0").Max = 100
      Let .Controls("Spnpro0").MIN = 0
      Let .Controls("Spnpro0").Tag = 0
      Let .Controls("Spnpro1").Max = 10
      Let .Controls("Spnpro1").MIN = 0
      Let .Controls("Spnpro1").Tag = 0
      Let .Controls("Txtpro6").Value = VBA.Format(0, "0.00%")
      Let .Controls("Txtpro6").Tag = 0
    End With
    
    Call VIEWS.ViewFormAddControl(.Controls("Frame1"), "Label", "Lblpro", 6, True, 5)
    Call VIEWS.ViewFormAddControl(.Controls("Frame1"), "TextBox", "Txtpro", 8, True, 7)
    Call VIEWS.ViewFormAddControl(.Controls("Frame1"), "SpinButton", "Spnpro", 4, True, 2) 'EVENT all
    Call VIEWS.ViewFormAddControl(.Controls("Frame1"), "ComboBox", "Cbopro", 3, True, 3)
    With .Controls("Frame1")
      Let .Caption = "Cálculo Proyección"
      Let .Width = .Parent.InsideWidth * 0.9
      Let .Height = .Parent.InsideHeight * 0.2
      Let .Left = xLeft
      Let .Top = .Parent.Controls("Frame0").Top + .Parent.Controls("Frame0").Height + 5

      Let .Controls("Lblpro5").Caption = "Tiempo de proyección del Proyecto:"
      Let .Controls("Lblpro5").AutoSize = True
      Let .Controls("Lblpro5").WordWrap = False
      Let .Controls("Lblpro5").Left = xLeft
      Let .Controls("Lblpro5").Top = xLeft / 1.5
      Let .Controls("Lblpro6").Caption = "Tasa de ganancia requerida:"
      Let .Controls("Lblpro6").AutoSize = True
      Let .Controls("Lblpro6").WordWrap = False
      Let .Controls("Lblpro6").Left = xLeft
      Let .Controls("Lblpro6").Top = xLeft + .Controls("Lblpro6").Height * 1.5
      
      Let .Controls("Txtpro7").Left = .Controls("Lblpro5").Left + .Controls("Lblpro5").Width + (xLeft / 2)
      Let .Controls("Txtpro7").Top = .Controls("Lblpro5").Top
      Let .Controls("Txtpro7").TextAlign = fmTextAlignRight
      Let .Controls("Txtpro8").Left = .Controls("Txtpro7").Left
      Let .Controls("Txtpro8").Top = .Controls("Lblpro6").Top
      Let .Controls("Txtpro8").TextAlign = fmTextAlignRight

      Let .Controls("Spnpro2").Left = .Controls("Txtpro7").Left + .Controls("Txtpro7").Width + 1
      Let .Controls("Spnpro2").Top = .Controls("Txtpro7").Top
      Let .Controls("Spnpro2").Height = .Controls("Txtpro7").Height
      Let .Controls("Spnpro2").Width = .Controls("Txtpro7").Height / 2
      Let .Controls("Spnpro3").Left = .Controls("Spnpro2").Left
      Let .Controls("Spnpro3").Top = .Controls("Txtpro8").Top
      Let .Controls("Spnpro3").Height = .Controls("Spnpro2").Height
      Let .Controls("Spnpro3").Width = .Controls("Spnpro2").Width

      Let .Controls("Spnpro4").Left = .Controls("Spnpro3").Left + .Controls("Spnpro3").Width
      Let .Controls("Spnpro4").Top = .Controls("Spnpro3").Top
      Let .Controls("Spnpro4").Height = .Controls("Spnpro3").Height
      Let .Controls("Spnpro4").Width = .Controls("Spnpro3").Width

      Let .Controls("Spnpro2").Max = 100
      Let .Controls("Spnpro2").MIN = 1
      Let .Controls("Txtpro7").Value = VBA.Format(1, "0")
      Let .Controls("Txtpro7").Tag = 1

      Let .Controls("Spnpro3").Max = 100
      Let .Controls("Spnpro3").MIN = 0
      Let .Controls("Spnpro3").Tag = 0
      Let .Controls("Spnpro4").Max = 100
      Let .Controls("Spnpro4").MIN = 0
      Let .Controls("Spnpro4").Tag = 0
      Let .Controls("Txtpro8").Value = VBA.Format(0, "50.00%")
      Let .Controls("Txtpro8").Tag = 0

      Let .Controls("Cbopro3").Left = .Controls("Spnpro2").Left + .Controls("Spnpro2").Width
      Let .Controls("Cbopro3").Top = .Controls("Spnpro2").Top
    End With

    Call VIEWS.ViewFormAddControl(.Controls("Frame2"), "CheckBox", "Chkpro", 2)
    Call VIEWS.ViewFormAddControl(.Controls("Frame2"), "OptionButton", "Optpro", 1)
    With .Controls("Frame2")
      Let .Caption = "Gráfico y resultados del ejercicio"
      Let .Width = ((.Parent.InsideWidth * 0.9) / 2) - 2.5
      Let .Height = .Parent.InsideHeight * 0.2
      Let .Left = xLeft
      Let .Top = .Parent.Controls("Frame1").Top + .Parent.Controls("Frame1").Height + 5

      Let .Controls("Chkpro0").Caption = "Mostrar VAN"
      Let .Controls("Chkpro0").AutoSize = True
      Let .Controls("Chkpro0").WordWrap = False
      Let .Controls("Chkpro1").Caption = "Mostrar TIR"
      Let .Controls("Chkpro1").AutoSize = True
      Let .Controls("Chkpro1").WordWrap = False

      Let .Controls("Chkpro2").Caption = "Generar gráfico en:"
      Let .Controls("Optpro0").Caption = "Hoja actual"
      Let .Controls("Optpro1").Caption = "Nueva Hoja de gráfico"

      Let .Controls("Chkpro2").AutoSize = True
      Let .Controls("Optpro0").AutoSize = True
      Let .Controls("Optpro1").AutoSize = True
      Let .Controls("Chkpro2").WordWrap = False
      Let .Controls("Optpro0").WordWrap = False
      Let .Controls("Optpro1").WordWrap = False

      Let .Controls("Optpro0").Value = True

      Let i = xLeft / 2.5
      Let .Controls("Chkpro0").Top = i
      Let .Controls("Chkpro1").Top = i + (.Controls("Chkpro0").Height * 0.8)

      Let .Controls("Chkpro2").Top = i + (.Controls("Chkpro1").Height * 0.8) * 2
      Let .Controls("Optpro0").Top = i + (.Controls("Chkpro2").Height * 0.8) * 3
      Let .Controls("Optpro1").Top = i + (.Controls("Chkpro2").Height * 0.8) * 3

      Let .Controls("Chkpro0").Left = xLeft
      Let .Controls("Chkpro1").Left = xLeft
      Let .Controls("Chkpro2").Left = xLeft
      Let .Controls("Optpro0").Left = xLeft
      Let .Controls("Optpro1").Left = (xLeft / 2) + .Controls("Optpro0").Left + .Controls("Optpro0").Width

      Let .Controls("Chkpro0").Value = True
      Let .Controls("Chkpro1").Value = True
    End With
    
    Call VIEWS.ViewFormAddControl(.Controls("Frame3"), "OptionButton", "Optpro", 3, True, 2)
    Call VIEWS.ViewFormAddControl(.Controls("Frame3"), "TextBox", "Txtpro", 9, True, 9)
    Call VIEWS.ViewFormAddControl(.Controls("Frame3"), "CommandButton", "Btnpro", 8, True, 8) 'EVENT all
    With .Controls("Frame3")
      Let .Caption = "Dónde mostrar resultados"
      Let .Width = ((.Parent.InsideWidth * 0.9) / 2) - 2.5
      Let .Height = .Parent.InsideHeight * 0.2
      Let .Left = .Parent.Controls("Frame2").Left + .Parent.Controls("Frame2").Width + 5
      Let .Top = .Parent.Controls("Frame1").Top + .Parent.Controls("Frame1").Height + 5

      Let .Controls("Optpro2").Caption = "Hoja nueva"
      Let .Controls("Optpro3").Caption = "Hoja actual"

      Let .Controls("Optpro2").AutoSize = True
      Let .Controls("Optpro3").AutoSize = True
      Let .Controls("Optpro2").WordWrap = False
      Let .Controls("Optpro3").WordWrap = False

      Let .Controls("Optpro2").Left = xLeft
      Let .Controls("Optpro3").Left = xLeft
      Let .Controls("Optpro2").Top = xLeft / 1.3
      Let .Controls("Optpro3").Top = (xLeft / 1.3) + .Controls("Optpro3").Height

      Let .Controls("Optpro2").Value = True

      Let .Controls("Txtpro9").Left = .Controls("Optpro3").Left + .Controls("Optpro3").Width
      Let .Controls("Txtpro9").Top = .Controls("Optpro3").Top
      Let .Controls("Btnpro8").Left = .Controls("Txtpro9").Left + .Controls("Txtpro9").Width + 1
      Let .Controls("Btnpro8").Top = .Controls("Txtpro9").Top
      Let .Controls("Btnpro8").Height = .Controls("Txtpro9").Height
      Let .Controls("Btnpro8").Width = .Controls("Txtpro9").Height
      Let .Controls("Btnpro8").Caption = VBA.ChrW(&H2583)
      Let .Controls("Btnpro8").Font.Bold = True
      Let .Controls("Btnpro8").ControlTipText = "Click para seleccionar una celda"
    End With

    '2. [ ] Fill controls with values
    Let .Controls("Cbopro0").RowSource = excliqlitedatasheet.ListObjects("excliqlitecurrenciescon") & "[symbol]"
    Let .Controls("Cbopro1").List = Array("Pb", "Ag", "Zn", "Cu", "Au", "Sn", "Sb", "Co", "Bi", "Al", "Ta", "Fe", "Ni", "Pt", "Pa")
    Let .Controls("Cbopro2").List = Array("TF", "LF", "OT", "KF", "TC", "TL")
    Let .Controls("Cbopro3").List = Array("Días", "Semanas", "Quincenas", "Meses", "Años")
    Let .Controls("Cbopro0").ListIndex = 0
    Let .Controls("Cbopro3").ListIndex = 0

    Let .Controls("Spnpro0").Tag = "A" 'Frame0
    Let .Controls("Spnpro1").Tag = "B" 'Frame0
    Let .Controls("Spnpro2").Tag = "A" 'Frame1
    Let .Controls("Spnpro3").Tag = "A" 'Frame1
    Let .Controls("Spnpro4").Tag = "B" 'Frame1

    '3. [x] Fill Form Controls with data and Set general purpose CommandButtons
    Let k = (.InsideWidth * 0.95)
    With .Controls("Btnmain0") '("BtnCancel")
      Let .Accelerator = "c"
      Let .Cancel = True
      Let .Top = (.Parent.InsideHeight * 0.95) - (.Height)
      Let .Left = k - .Width
      Let .Caption = "Cerrar"
    End With
    With .Controls("Btnmain2") '("BtnHelp")
      Let .Caption = "Ayuda"
      Let .Accelerator = "a"
      Let .ControlTipText = "Ver ayuda del Ejercicio"
      Let .Top = .Parent.Controls("Btnmain0").Top
      Let .Left = xLeft
      Let .Tag = ControlId '& cGet
    End With
    With .Controls("Btnmain1") '("BtnEXE")
      Let .Caption = "Resolver"
      Let .Accelerator = "r"
      Let .ControlTipText = "Permite procesar los datos y mostrar los resultados en una hoja Excel."
      Let .Top = .Parent.Controls("Btnmain0").Top
      Let .Left = .Parent.Controls("Btnmain0").Left - .Width - 2
      Let .Tag = ControlId & "-bal"
      Let .Default = True
    End With

    '4. Assing handler for each button
    Set BtnCancel = .Controls("Btnmain0") 'Main
    Set BtnAccept = .Controls("Btnmain1") 'Main
    Set Btn0 = .Controls("Btnmain2")      'Main
    Set Btn1 = .Controls("Btnpro0") 'Frame0
    Set Btn2 = .Controls("Btnpro1") 'Frame0
    Set Btn3 = .Controls("Btnpro2") 'Frame0
    Set Btn4 = .Controls("Btnpro3") 'Frame0
    Set Btn5 = .Controls("Btnpro4") 'Frame0
    Set Btn6 = .Controls("Btnpro5") 'Frame0
    Set Btn7 = .Controls("Btnpro6") 'Frame0
    Set Btn8 = .Controls("Btnpro7") 'Frame0
    Set Btn9 = .Controls("Btnpro8") 'Frame3

    Set Cbo1 = .Controls("Cbopro1") 'Frame0
    Set Cbo2 = .Controls("Cbopro2") 'Frame0

    Set Spn0 = .Controls("Spnpro0") 'Frame0
    Set Spn1 = .Controls("Spnpro1") 'Frame0
    Set Spn2 = .Controls("Spnpro2") 'Frame1
    Set Spn3 = .Controls("Spnpro3") 'Frame1
    Set Spn4 = .Controls("Spnpro4") 'Frame1

    '5. Give styles to controls
    Let .Controls("Cbopro1").ListIndex = 0
    Let .Controls("Cbopro2").ListIndex = 0
    Call VIEWS.ViewFormStyles(FRM): Set VIEWS = Nothing

    '6. [x] Show the finished Form
    Call Load(FRM)
    Call .Show

    '7. [x] Close form
    Let .Controls("Cbopro0").RowSource = VBA.vbNullString
  End With

EH:
  Erase aLabels
  Erase aMaxLengths
  Erase aControlis
  Erase aControls
  
  Set VIEWS = Nothing
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)
End Sub

'Purchase Mineral UI
Private Sub View_form_purchasemineral(ByVal ControlId As String, ByVal FRM As FRMBasics)
  'Main objects
  Dim RES As AppResources
  Dim VIEWS As ViewsApp
  'ListObjects
  'LO data names and other data
  'Controls
  Dim aLabels() As Variant
  Dim aMaxLengths() As Variant
  Dim aControls() As Variant
  Dim aControlis() As Variant
  'Standard vars
  Dim bNoData As Boolean, bTables As Boolean
  Dim sTitle As String, sTable As String, TitleForm As String, StrAux As String
  Dim i As Long, j As Long, k As Long, l As Long
  Dim xLeft As Long
  Dim Aux0 As String, Aux1 As String, Aux2 As String
  Dim StrDivlocal As String, StrDivforeign As String

  
  Const cLeft As Single = 10
  Const sFN As String = "VIEW::-form_purchasemineral"

  #If Debugging Then
    Call Debug_assert(Not ControlId = VBA.vbNullString, sFN)
    Call Debug_assert(Not FRM Is Nothing, sFN)
  #End If
  
  If ControlId = VBA.vbNullString Then GoTo EH
  If FRM Is Nothing Then GoTo EH
  On Error GoTo EH
  
  'Get data for controls

  'Set Form basic datas
  Let TitleForm = "Liquidación de Minerales"
  Let sTitle = TitleForm & " - " & STRAPPTITLE
  
  Set VIEWS = New ViewsApp
  Set VIEWS.ErrorHandler = EHGLOBAL
  Call VIEWS.ViewFormCleanner(FRM)
  With FRM
    'Set the form dimensions
    Let .Height = 415
    Let .Width = 450
    Let .Tag = ControlId
    Let .Caption = sTitle
    Let xLeft = .InsideWidth * 0.05

    'Build and design the form
    Call VIEWS.ViewFormAddControl(FRM, "CommandButton", "Btnmain", 2) 'All with events (Cancel, Exe and Help)
    Call VIEWS.ViewFormAddControl(FRM, "Frame", "Frame", 3)

    Let StrAux = "Frame0"
    Call VIEWS.ViewFormAddControl(.Controls(StrAux), "Label", "Lblpur", 2)
    Call VIEWS.ViewFormAddControl(.Controls(StrAux), "TextBox", "Txtpur", 2)
    Call VIEWS.ViewFormAddControl(.Controls(StrAux), "CheckBox", "Chkpur", 0)
    Call VIEWS.ViewFormAddControl(.Controls(StrAux), "CommandButton", "Btnpur", 0) 'EVENT all
    With .Controls(StrAux)
      Let .Caption = "Datos Base"
      Let .Width = (.Parent.InsideWidth * 0.95) '* 0.6
      Let .Height = .Parent.InsideHeight * 0.15
      Let .Left = xLeft * 0.45
      Let .Top = .Parent.InsideHeight * 0.02
      
      Let .Controls("Lblpur0").Caption = "Lugar:"
      Let .Controls("Lblpur1").Caption = "Fecha:"
      Let .Controls("Lblpur0").Left = xLeft * 0.4
      Let .Controls("Lblpur1").Left = .Controls("Lblpur0").Left
      Let .Controls("Lblpur0").AutoSize = True
      Let .Controls("Lblpur1").WordWrap = False
      Let .Controls("Lblpur0").AutoSize = True
      Let .Controls("Lblpur1").WordWrap = False
      Let .Controls("Lblpur0").Top = (xLeft * 0.32) '+ .Controls("Lblpur0").Height ' * 2
      Let .Controls("Lblpur1").Top = (.Controls("Lblpur0").Top + .Controls("Lblpur0").Top) * 2 ' + .Controls("Lblpur1").Top

      Let .Controls("Txtpur0").Left = .Controls("Lblpur0").Left + .Controls("Lblpur0").Width + 5
      Let .Controls("Txtpur1").Left = .Controls("Txtpur0").Left
      Let .Controls("Txtpur0").Top = .Controls("Lblpur0").Top
      Let .Controls("Txtpur1").Top = .Controls("Txtpur0").Top + .Controls("Txtpur0").Height + 1
      Let .Controls("Txtpur0").Width = .Controls("Txtpur0").Width * 2
      Let .Controls("Txtpur1").Width = .Controls("Txtpur0").Width - .Controls("Txtpur1").Height
      Let .Controls("Txtpur0").ControlTipText = "Ejemplo: Potosí"
      Let .Controls("Txtpur1").ControlTipText = "Ejemplo: 01/01/21 o seleccione una fecha mediante el botón"
      Let .Controls("Txtpur1").Value = VBA.Format$(VBA.Date(), "dd/mm/yy")
      Let .Controls("Txtpur1").TextAlign = fmTextAlignRight

      Let .Controls("Btnpur0").Left = .Controls("Txtpur1").Left + .Controls("Txtpur1").Width
      Let .Controls("Btnpur0").Top = .Controls("Txtpur1").Top
      Let .Controls("Btnpur0").Height = .Controls("Txtpur1").Height
      Let .Controls("Btnpur0").Width = .Controls("Btnpur0").Height
      Let .Controls("Btnpur0").Caption = VBA.ChrW(&H2705)
      Let .Controls("Btnpur0").Font.Bold = True
      Let .Controls("Btnpur0").ControlTipText = "Pulse para seleccionar una fecha"
      
      Let StrDivlocal = "BOB": Let StrDivforeign = "USD" 'TODO: OBTENER RE RECURSOS
      Let .Controls("Lblpur2").Caption = "TC [" & StrDivlocal & "/" & StrDivforeign & "]:"
      Let .Controls("Chkpur0").Caption = "Imprimir Moneda Extranjera"
      Let .Controls("Lblpur2").Left = .Controls("Txtpur0").Left + .Controls("Txtpur0").Width + xLeft
      Let .Controls("Chkpur0").Left = .Controls("Lblpur2").Left
      Let .Controls("Lblpur2").AutoSize = True
      Let .Controls("Chkpur0").AutoSize = True
      Let .Controls("Lblpur2").WordWrap = False
      Let .Controls("Chkpur0").WordWrap = False
      Let .Controls("Lblpur2").Top = .Controls("Txtpur0").Top
      Let .Controls("Chkpur0").Top = .Controls("Btnpur0").Top
      Let .Controls("Chkpur0").ControlTipText = "Generará una columna con los valores en moneda extranjera en la Liquidación Final"

      Let .Controls("Txtpur2").Left = .Controls("Lblpur2").Left + .Controls("Lblpur2").Width + 2
      Let .Controls("Txtpur2").Top = .Controls("Lblpur2").Top
      Let .Controls("Txtpur2").Width = .Controls("Txtpur2").Width * 2
      Let .Controls("Txtpur2").Value = "1"
      Let .Controls("Txtpur2").ControlTipText = "Ingrese el equivalente al tipo de cambio actual"
    End With
    
    Let StrAux = "Frame1"
    Call VIEWS.ViewFormAddControl(.Controls(StrAux), "Label", "Lblpur", 4, CtlFromX:=3)
    Call VIEWS.ViewFormAddControl(.Controls(StrAux), "TextBox", "Txtpur", 4, CtlFromX:=3)
    Call VIEWS.ViewFormAddControl(.Controls(StrAux), "ComboBox", "Cbopur", 0)
    Call VIEWS.ViewFormAddControl(.Controls(StrAux), "CheckBox", "Chkpur", 1, CtlFromX:=1) 'EVENT all
    Call VIEWS.ViewFormAddControl(.Controls(StrAux), "CommandButton", "Btnpur", 2, CtlFromX:=1) 'EVENT all
    With .Controls(StrAux)
      Let .Caption = "Datos Provedor"
      Let .Width = (.Parent.InsideWidth * 0.95) '* 0.4
      Let .Height = .Parent.InsideHeight * 0.12
      Let .Left = .Parent.Controls("Frame0").Left '+ .Parent.Controls("Frame0").Width + 2
      Let .Top = .Parent.Controls("Frame0").Top + .Parent.Controls("Frame0").Height + 1

      Let .Controls("Chkpur1").Caption = "Varias liquidaciones a la vez"
      Let .Controls("Chkpur1").ControlTipText = "Permite generar hasta diez diferentes liquidaciones una por página"
      Let .Controls("Chkpur1").Left = xLeft * 0.4
      Let .Controls("Chkpur1").AutoSize = True
      Let .Controls("Chkpur1").WordWrap = False
      Let .Controls("Chkpur1").Top = (xLeft * 0.15)
      
      Let .Controls("Lblpur3").Caption = "Nombres y Apellidos:"
      Let .Controls("Lblpur3").Left = .Controls("Chkpur1").Left + .Controls("Chkpur1").Width + xLeft 'xLeft * 0.4
      Let .Controls("Lblpur3").AutoSize = True
      Let .Controls("Lblpur3").WordWrap = False
      Let .Controls("Lblpur3").Top = .Controls("Chkpur1").Top * 1.6 '(xLeft * 0.3)

      Let .Controls("Txtpur3").Left = .Controls("Lblpur3").Left ' + .Controls("Lblpur3").Width + 2
      Let .Controls("Txtpur3").Top = .Controls("Txtpur3").Top + .Controls("Lblpur3").Top + .Controls("Lblpur3").Height ' .Controls("Lblpur3").Top
      Let .Controls("Txtpur3").Width = .Controls("Txtpur3").Width * 1.5
      
      Let .Controls("Btnpur1").Left = .Controls("Txtpur3").Left + .Controls("Txtpur3").Width + 1
      Let .Controls("Btnpur1").Top = .Controls("Txtpur3").Top
      Let .Controls("Btnpur1").Height = .Controls("Txtpur3").Height
      Let .Controls("Btnpur1").Width = .Controls("Btnpur1").Height
      Let .Controls("Btnpur1").Caption = VBA.ChrW(&H2583)
      Let .Controls("Btnpur1").ControlTipText = "Pulse para seleccionar un rango de valores"
      Let .Controls("Btnpur1").Visible = False
      Let .Controls("Btnpur1").Tag = "selection"
      
      Let .Controls("Lblpur4").Caption = "Tipo:"
      Let .Controls("Lblpur4").Left = .Controls("Btnpur1").Left + .Controls("Btnpur1").Width + xLeft
      Let .Controls("Lblpur4").AutoSize = True
      Let .Controls("Lblpur4").WordWrap = False
      Let .Controls("Lblpur4").Top = .Controls("Lblpur3").Top
      
      Let .Controls("Cbopur0").Left = .Controls("Lblpur4").Left
      Let .Controls("Cbopur0").Top = .Controls("Btnpur1").Top
      Let .Controls("Cbopur0").Width = (.Controls("Cbopur0").Width * 1.5) '- .Controls("Cbopur0").Height
    
      Let .Controls("Txtpur4").Left = .Controls("Cbopur0").Left ' + .Controls("Lblpur3").Width + 2
      Let .Controls("Txtpur4").Top = .Controls("Cbopur0").Top
      Let .Controls("Txtpur4").Width = (.Controls("Txtpur4").Width * 1.5) - .Controls("Txtpur4").Height
      Let .Controls("Txtpur4").Visible = False
    
      Let .Controls("Btnpur2").Left = .Controls("Txtpur4").Left + .Controls("Txtpur4").Width + 1
      Let .Controls("Btnpur2").Top = .Controls("Txtpur4").Top
      Let .Controls("Btnpur2").Height = .Controls("Txtpur4").Height
      Let .Controls("Btnpur2").Width = .Controls("Btnpur2").Height
      Let .Controls("Btnpur2").Caption = VBA.ChrW(&H2583)
      Let .Controls("Btnpur2").ControlTipText = "Pulse para seleccionar un rango de valores"
      Let .Controls("Btnpur2").Visible = False
      Let .Controls("Btnpur2").Tag = "selection"
    End With
    
    Let StrAux = "Frame2"
    Call VIEWS.ViewFormAddControl(.Controls(StrAux), "ComboBox", "Cbopur", 4, CtlFromX:=1, CtlVisible:=False) 'EVENT all
    Call VIEWS.ViewFormAddControl(.Controls(StrAux), "CommandButton", "Btnpur", 9, CtlFromX:=3) 'EVENT all
    Call VIEWS.ViewFormAddControl(.Controls(StrAux), "TextBox", "Txtpur", 9, CtlFromX:=5)
    Call VIEWS.ViewFormAddControl(.Controls(StrAux), "ListBox", "Lstpur", 0) 'EVENT
    Call VIEWS.ViewFormAddControl(.Controls(StrAux), "CheckBox", "Chkpur", 5, CtlFromX:=2)
    With .Controls(StrAux)
      Let .Caption = "Datos Liquidación"
      Let .Width = (.Parent.InsideWidth * 0.95) + 2
      Let .Height = .Parent.InsideHeight * 0.5
      Let .Left = .Parent.Controls("Frame0").Left
      Let .Top = .Parent.Controls("Frame1").Top + .Parent.Controls("Frame1").Height + 1

      Let .Controls("Cbopur1").Left = xLeft * 0.4
      Let .Controls("Cbopur1").Top = xLeft * 0.4
      Let .Controls("Cbopur1").Width = (.Width * 0.95) - ((.Controls("Cbopur1").Height * 6) + 1) '6 Buttons
      Let .Controls("Cbopur1").Visible = True
      Let .Controls("Cbopur1").ControlTipText = "Seleccione una opción e ingrese los datos respectivos en la lista"

      For i = 3 To 8
        Let .Controls("Btnpur" & i).Top = .Controls("Cbopur1").Top
        Let .Controls("Btnpur" & i).Height = .Controls("Cbopur1").Height
        Let .Controls("Btnpur" & i).Width = .Controls("Btnpur" & i).Height
        Let .Controls("Btnpur" & i).Left = (.Controls("Cbopur1").Left + .Controls("Cbopur1").Width) + (.Controls("Btnpur" & i).Width * (i - 3))
      Next i
      Let .Controls("Btnpur3").Caption = VBA.ChrW(&H271A)
      Let .Controls("Btnpur3").ControlTipText = "Pulse para añadir los valores de los controles a la lista"
      Let .Controls("Btnpur4").Caption = VBA.ChrW(&H2710)
      Let .Controls("Btnpur4").ControlTipText = "Pulse para confirmar la edición de los valores que modifique"
      Let .Controls("Btnpur5").Caption = VBA.ChrW(&H270F)
      Let .Controls("Btnpur5").ControlTipText = "Cancelar edición"
      Let .Controls("Btnpur6").Caption = VBA.ChrW(&H2A31)
      Let .Controls("Btnpur6").ControlTipText = "Pulse para eliminar la fila de datos seleccionada en la lista"
      Let .Controls("Btnpur5").Font.Bold = True
      Let .Controls("Btnpur4").Font.Size = 10
      Let .Controls("Btnpur6").Font.Bold = True
      
      Let .Controls("Btnpur7").Caption = VBA.ChrW(&H2716)
      Let .Controls("Btnpur7").ControlTipText = "Pulse para eliminar todos los datos que ingresó"
      
      Let .Controls("Btnpur8").Caption = VBA.ChrW(&H2583)
      Let .Controls("Btnpur8").ControlTipText = "Seleccionar datos desde un rango de celdas de " & Application.Name
      Let .Controls("Btnpur8").Tag = "selection"
      
      Let .Controls("Btnpur4").Enabled = False
      Let .Controls("Btnpur5").Enabled = False
      Let .Controls("Btnpur6").Enabled = False
      Let .Controls("Btnpur7").Enabled = False

      For i = 5 To 9
        Let .Controls("Txtpur" & i).Top = (.Controls("Cbopur1").Top + .Controls("Cbopur1").Height) + 5
        Let .Controls("Txtpur" & i).Width = (.Width * 0.95) / 5
        Let .Controls("Txtpur" & i).Left = (.Controls("Cbopur1").Left) + (.Controls("Txtpur" & i).Width * (i - 5))
      Next i
      Let .Controls("Txtpur6").Width = .Controls("Txtpur6").Width - .Controls("Txtpur6").Height
      Let .Controls("Txtpur6").Value = VBA.Format$(VBA.Date(), "dd/mm/yy")
      Let .Controls("Txtpur6").TextAlign = fmTextAlignRight

      Let .Controls("Btnpur9").Caption = VBA.ChrW(&H2705)
      Let .Controls("Btnpur9").ControlTipText = "Pulse para seleccionar una fecha"
      Let .Controls("Btnpur9").Font.Bold = True
      Let .Controls("Btnpur9").Height = .Controls("Txtpur6").Height
      Let .Controls("Btnpur9").Width = .Controls("Txtpur6").Height
      Let .Controls("Btnpur9").Top = .Controls("Txtpur6").Top
      Let .Controls("Btnpur9").Left = .Controls("Txtpur6").Left + .Controls("Txtpur6").Width

      Let .Controls("Cbopur2").Width = .Controls("Txtpur5").Width
      Let .Controls("Cbopur2").Left = .Controls("Txtpur5").Left
      Let .Controls("Cbopur2").Top = .Controls("Txtpur5").Top

      Let .Controls("Cbopur3").Width = .Controls("Txtpur6").Width / 2
      Let .Controls("Cbopur3").Left = (.Controls("Txtpur6").Left + .Controls("Txtpur6").Width / 2)
      Let .Controls("Cbopur3").Top = .Controls("Txtpur6").Top

      Let .Controls("Cbopur4").Width = .Controls("Txtpur8").Width / 2
      Let .Controls("Cbopur4").Left = (.Controls("Txtpur8").Left + .Controls("Txtpur8").Width / 2)
      Let .Controls("Cbopur4").Top = .Controls("Txtpur8").Top

      Let .Controls("Lstpur0").Width = .Width * 0.95
      Let .Controls("Lstpur0").Left = .Controls("Txtpur5").Left
      Let .Controls("Lstpur0").Top = .Controls("Txtpur5").Top + .Controls("Txtpur5").Height + 2
      Let .Controls("Lstpur0").Height = .Height * 0.6 ' .Controls("Lstpur0").Height * 1.9
      Let .Controls("Lstpur0").ControlTipText = "Doble click en el ítem que desee modificar"

      Let .Controls("Chkpur2").Left = .Controls("Txtpur6").Left + (xLeft * 0.5)
      Let .Controls("Chkpur2").Top = .Controls("Txtpur6").Top
      Let .Controls("Chkpur2").AutoSize = True
      Let .Controls("Chkpur2").WordWrap = False
      Let .Controls("Chkpur2").Visible = False
      Let .Controls("Chkpur2").Caption = "Todas las "
      
      Let .Controls("Chkpur3").Left = .Controls("Txtpur5").Left '.Controls("Txtpur5").Left + .Controls("Chkpur3").Width + (xLeft * 0.7)
      Let .Controls("Chkpur3").Top = (.Controls("Lstpur0").Top + .Controls("Lstpur0").Height)
      Let .Controls("Chkpur3").AutoSize = True
      Let .Controls("Chkpur3").WordWrap = False
      Let .Controls("Chkpur3").Caption = "Promediar"
      Let .Controls("Chkpur3").ControlTipText = "Obtendrá los promedios de leyes, %Humedad y %Merma en la Liquidación Final"
      
      Let .Controls("Chkpur4").Left = .Controls("Chkpur3").Left + .Controls("Chkpur3").Width + (xLeft * 0.7)
      Let .Controls("Chkpur4").Top = .Controls("Chkpur3").Top
      Let .Controls("Chkpur4").AutoSize = True
      Let .Controls("Chkpur4").WordWrap = False
      Let .Controls("Chkpur4").Caption = "Autoconfigurar página"
      Let .Controls("Chkpur4").ControlTipText = "El sistema configurará las páginas donde imprimirá la Liquidación Final"
      
      Let .Controls("Chkpur5").Left = .Controls("Chkpur4").Left + .Controls("Chkpur4").Width + (xLeft * 0.7)
      Let .Controls("Chkpur5").Top = .Controls("Chkpur4").Top
      Let .Controls("Chkpur5").AutoSize = True
      Let .Controls("Chkpur5").WordWrap = False
      Let .Controls("Chkpur5").Caption = "Dos copias por página"
      Let .Controls("Chkpur5").ControlTipText = "Obtendrá cada Liquidación Final con una copia, ambas en una página independiente en orientación horizontal"
    End With

    Let StrAux = "Frame3"
    Call VIEWS.ViewFormAddControl(.Controls(StrAux), "OptionButton", "Optpur", 1)
    Call VIEWS.ViewFormAddControl(.Controls(StrAux), "TextBox", "Txtpur", 10, True, 10)
    Call VIEWS.ViewFormAddControl(.Controls(StrAux), "CommandButton", "Btnpur", 10, True, 10) 'EVENT all
    With .Controls(StrAux)
      Let .Caption = "Dónde mostrar resultados"
      Let .Width = .Parent.Controls("Frame2").Width
      Let .Height = .Parent.InsideHeight * 0.1
      Let .Left = .Parent.Controls("Frame2").Left
      Let .Top = .Parent.Controls("Frame2").Top + .Parent.Controls("Frame2").Height + 1

      Let .Controls("Optpur0").Caption = "Hoja nueva"
      Let .Controls("Optpur1").Caption = "Hoja actual"

      Let .Controls("Optpur0").AutoSize = True
      Let .Controls("Optpur1").AutoSize = True
      Let .Controls("Optpur0").WordWrap = False
      Let .Controls("Optpur1").WordWrap = False

      Let .Controls("Optpur0").Left = xLeft
      Let .Controls("Optpur1").Left = xLeft + .Width * 0.5
      Let .Controls("Optpur0").Top = xLeft * 0.3
      Let .Controls("Optpur1").Top = .Controls("Optpur0").Top

      Let .Controls("Optpur0").Value = True

      Let .Controls("Txtpur10").Left = .Controls("Optpur1").Left + .Controls("Optpur1").Width
      Let .Controls("Txtpur10").Top = .Controls("Optpur1").Top
      Let .Controls("Btnpur10").Left = .Controls("Txtpur10").Left + .Controls("Txtpur10").Width + 1
      Let .Controls("Btnpur10").Top = .Controls("Txtpur10").Top
      Let .Controls("Btnpur10").Height = .Controls("Txtpur10").Height
      Let .Controls("Btnpur10").Width = .Controls("Txtpur10").Height
      Let .Controls("Btnpur10").Caption = VBA.ChrW(&H2583)
      Let .Controls("Btnpur10").ControlTipText = "Click para seleccionar una celda"
    End With

    '2. [ ] Fill controls with values
    Let .Controls("Cbopur0").List = Array("Cooperativista", "Particular")
    Let .Controls("Cbopur1").List = Array("Datos base del mineral", "Datos de calidad y valor del mineral", "Bonos", "Retenciones de ley", "Otras retenciones", "Descuentos")
    Let .Controls("Cbopur2").List = Array("Mineral", "Mineral Complejo", "Brosa")
    'Let .Controls("Cbopur3").List = Array("Pb", "Ag", "Zn", "Cu", "Au", "Sn", "Sb", "Co", "Bi", "Al", "Ta", "Fe", "Ni", "Pt", "Pa")
    Let .Controls("Cbopur4").List = Array("%", "DM", "OT", "g/T", "g/L")

    Let .Controls("Cbopur0").ListIndex = 0
    'Let .Controls("Cbopur3").ListIndex = 0
    Let .Controls("Cbopur4").ListIndex = 0
    Let .Controls("Lstpur0").ColumnHeads = True

    '3. [x] Fill Form Controls with data and Set general purpose CommandButtons
    Let k = (.InsideWidth * 0.95)
    With .Controls("Btnmain0") '("BtnCancel")
      Let .Accelerator = "c"
      Let .Cancel = True
      Let .Top = (.Parent.InsideHeight * 0.97) - (.Height)
      Let .Left = k - .Width + xLeft * 0.5
      Let .Caption = "Cerrar"
    End With
    With .Controls("Btnmain2") '("BtnHelp")
      Let .Caption = "Ayuda"
      Let .Accelerator = "a"
      Let .ControlTipText = "Ver ayuda del Ejercicio"
      Let .Top = .Parent.Controls("Btnmain0").Top
      Let .Left = xLeft * 0.45
      Let .Tag = ControlId '& cGet
    End With
    With .Controls("Btnmain1") '("BtnEXE")
      Let .Caption = "Generar"
      Let .Accelerator = "g"
      Let .ControlTipText = "Permite purcesar los datos y mostrar los resultados en una hoja Excel."
      Let .Top = .Parent.Controls("Btnmain0").Top
      Let .Left = .Parent.Controls("Btnmain0").Left - .Width - 2
      Let .Tag = ControlId & "-bal"
      Let .Default = True
    End With

    '4. Assing handler for each button
    Set BtnCancel = .Controls("Btnmain0") 'Main
    Set BtnAccept = .Controls("Btnmain1") 'Main
    Set Btn0 = .Controls("Btnmain2")      'Main
    Set Btn1 = .Controls("Btnpur0") 'Frame0
    Set Btn2 = .Controls("Btnpur1") 'Frame1
    Set Btn3 = .Controls("Btnpur2") 'Frame1
    Set Btn4 = .Controls("Btnpur3") 'Frame2
    Set Btn5 = .Controls("Btnpur4") 'Frame2
    Set Btn6 = .Controls("Btnpur5") 'Frame2
    Set Btn7 = .Controls("Btnpur6") 'Frame2
    Set Btn8 = .Controls("Btnpur7") 'Frame2
    Set Btn9 = .Controls("Btnpur8") 'Frame2
    Set Btn10 = .Controls("Btnpur9") 'Frame2
    Set Btn11 = .Controls("Btnpur10") 'Frame3

    Set Chk0 = .Controls("Chkpur1") 'Frame2
    Set Chk1 = .Controls("Chkpur2") 'Frame2
    'Set Chk2 = .Controls("Chkpur4") 'Frame2
    
    Set Cbo0 = .Controls("Cbopur1") 'Frame2
    Set Cbo1 = .Controls("Cbopur2") 'Frame2
    Set Cbo2 = .Controls("Cbopur3") 'Frame2
    Set Cbo3 = .Controls("Cbopur4") 'Frame2

    Set Lst0 = .Controls("Lstpur0") 'Frame2

    Set DPR = New ClassDatePicker

    '5. Give styles to controls
    Let .Controls("Cbopur1").ListIndex = 0
    Call VIEWS.ViewFormStyles(FRM): Set VIEWS = Nothing
    
    '6. [x] Show the finished Form
    Call Load(FRM)
    Call .Show

    '7. [x] Close form
    Let .Controls("Lstpur0").RowSource = VBA.vbNullString
    Let .Controls("Cbopur0").List = Empty
    Let .Controls("Cbopur1").List = Empty
    Let .Controls("Cbopur2").List = Empty
    Let .Controls("Cbopur3").List = Empty
    Let .Controls("Cbopur4").List = Empty
  End With

EH:
  Erase aLabels
  Erase aMaxLengths
  Erase aControlis
  Erase aControls
  Set DPR = Nothing
  Set VIEWS = Nothing
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)
End Sub

Private Sub View_form_purchaseconcentrate(ByVal ControlId As String, ByVal FRM As FRMBasics)
  'Main objects
  Dim RES As AppResources
  Dim VIEWS As ViewsApp
  'ListObjects
  'LO data names and other data
  'Controls
  Dim aLabels() As Variant
  Dim aMaxLengths() As Variant
  Dim aControls() As Variant
  Dim aControlis() As Variant
  'Standard vars
  Dim sTitle As String, TitleForm As String, StrAux As String
  Dim i As Long, j As Long, k As Long, l As Long
  Dim xLeft As Long
  Dim StrDivlocal As String, StrDivforeign As String

  
  Const cLeft As Single = 10
  Const sFN As String = "VIEW::-form_purchaseconcentrate"

  #If Debugging Then
    Call Debug_assert(Not ControlId = VBA.vbNullString, sFN)
    Call Debug_assert(Not FRM Is Nothing, sFN)
  #End If
  
  If ControlId = VBA.vbNullString Then GoTo EH
  If FRM Is Nothing Then GoTo EH
  On Error GoTo EH
  
  'Get data for controls

  'Set Form basic datas
  Let TitleForm = "Liquidación de Concentrados Minerales"
  Let sTitle = TitleForm & " - " & STRAPPTITLE
  
  Set VIEWS = New ViewsApp
  Set VIEWS.ErrorHandler = EHGLOBAL
  Call VIEWS.ViewFormCleanner(FRM)
  With FRM
    'Set the form dimensions
    Let .Height = 415
    Let .Width = 450
    Let .Tag = ControlId
    Let .Caption = sTitle
    Let xLeft = .InsideWidth * 0.05

    'Build and design the form
    Call VIEWS.ViewFormAddControl(FRM, "CommandButton", "Btnmain", 3) 'All with events (Cancel, Exe and Help, and All purcnc data)
    Call VIEWS.ViewFormAddControl(FRM, "Frame", "Frame", 4)
    Call VIEWS.ViewFormAddControl(FRM, "TextBox", "Txtcnc", 0)
    Call VIEWS.ViewFormAddControl(FRM, "OptionButton", "Optcnc", 1)

    Let StrAux = "Frame0"
    Call VIEWS.ViewFormAddControl(.Controls(StrAux), "Label", "Lblcnc", 2)
    Call VIEWS.ViewFormAddControl(.Controls(StrAux), "TextBox", "Txtcnc", 3, True, 1)
    Call VIEWS.ViewFormAddControl(.Controls(StrAux), "CheckBox", "Chkcnc", 0)
    Call VIEWS.ViewFormAddControl(.Controls(StrAux), "CommandButton", "Btncnc", 0) 'EVENT all
    With .Controls(StrAux)
      Let .Caption = "Datos Base"
      Let .Width = (.Parent.InsideWidth * 0.47) '* 0.6
      Let .Height = .Parent.InsideHeight * 0.2
      Let .Left = xLeft * 0.45
      Let .Top = .Parent.InsideHeight * 0.02

      Let StrDivlocal = "BOB": Let StrDivforeign = "USD" 'TODO: OBTENER RE RECURSOS
      Let .Controls("Lblcnc0").Caption = "Fecha:"
      Let .Controls("Lblcnc0").Accelerator = "f"
      Let .Controls("Lblcnc0").Left = xLeft * 0.4
      Let .Controls("Lblcnc0").AutoSize = True
      Let .Controls("Lblcnc0").WordWrap = False
      Let .Controls("Lblcnc0").Top = (xLeft * 0.3)
      
      Let .Controls("Txtcnc1").Left = .Controls("Lblcnc0").Left
      Let .Controls("Txtcnc1").Top = .Controls("Lblcnc0").Top + .Controls("Lblcnc0").Height
      Let .Controls("Txtcnc1").ControlTipText = "Ejemplo: 01/01/21 o seleccione una fecha mediante el botón"
      Let .Controls("Txtcnc1").Value = VBA.Format$(VBA.Date(), "dd/mm/yy")
      
      Let .Controls("Btncnc0").Left = .Controls("Txtcnc1").Left + .Controls("Txtcnc1").Width
      Let .Controls("Btncnc0").Top = .Controls("Txtcnc1").Top
      Let .Controls("Btncnc0").Height = .Controls("Txtcnc1").Height
      Let .Controls("Btncnc0").Width = .Controls("Btncnc0").Height
      Let .Controls("Btncnc0").Caption = VBA.ChrW(&H2705)
      Let .Controls("Btncnc0").Font.Bold = True
      Let .Controls("Btncnc0").ControlTipText = "Pulse para seleccionar una fecha"
      
      
      Let .Controls("Lblcnc1").Caption = "Lugar:"
      Let .Controls("Lblcnc0").Accelerator = "l"
      Let .Controls("Lblcnc1").Left = .Controls("Lblcnc0").Left
      Let .Controls("Lblcnc1").AutoSize = True
      Let .Controls("Lblcnc1").WordWrap = False
      Let .Controls("Lblcnc1").Top = (.Controls("Txtcnc1").Top + .Controls("Txtcnc1").Height) + 3
      
      Let .Controls("Txtcnc2").Left = .Controls("Lblcnc1").Left
      Let .Controls("Txtcnc2").Top = .Controls("Lblcnc1").Top + .Controls("Lblcnc1").Height
      Let .Controls("Txtcnc2").ControlTipText = "Ejemplo: Potosí"
      Let .Controls("Txtcnc2").TextAlign = fmTextAlignRight


      Let .Controls("Lblcnc2").Caption = "TC [" & StrDivlocal & "/" & StrDivforeign & "]:"
      Let .Controls("Lblcnc2").Left = (.Controls("Txtcnc2").Left + .Controls("Txtcnc2").Width) * 1.3
      Let .Controls("Lblcnc2").AutoSize = True
      Let .Controls("Lblcnc2").WordWrap = False
      Let .Controls("Lblcnc2").Top = .Controls("Lblcnc0").Top

      Let .Controls("Txtcnc3").Left = .Controls("Lblcnc2").Left
      Let .Controls("Txtcnc3").Top = (.Controls("Lblcnc2").Top + .Controls("Lblcnc2").Height)
      Let .Controls("Txtcnc3").ControlTipText = "Ingrese el equivalente al tipo de cambio actual"
      
      Let .Controls("Chkcnc0").Caption = "Imprimir Moneda Local"
      Let .Controls("Chkcnc0").Left = .Controls("Lblcnc2").Left
      Let .Controls("Chkcnc0").AutoSize = True
      Let .Controls("Chkcnc0").WordWrap = False
      Let .Controls("Chkcnc0").Top = .Controls("Txtcnc2").Top
      Let .Controls("Chkcnc0").ControlTipText = "Generará una columna con los valores en moneda local en la Liquidación Final"
    End With

    Let StrAux = "Frame1"
    Call VIEWS.ViewFormAddControl(.Controls(StrAux), "Label", "Lblcnc", 4, CtlFromX:=3)
    Call VIEWS.ViewFormAddControl(.Controls(StrAux), "TextBox", "Txtcnc", 4, CtlFromX:=4)
    Call VIEWS.ViewFormAddControl(.Controls(StrAux), "ComboBox", "Cbocnc", 0)
    With .Controls(StrAux)
      Let .Caption = "Datos Provedor"
      Let .Width = (.Parent.InsideWidth * 0.47) '* 0.4
      Let .Height = .Parent.InsideHeight * 0.2
      Let .Left = .Parent.Controls("Frame0").Left + .Parent.Controls("Frame0").Width + 5
      Let .Top = .Parent.Controls("Frame0").Top '+ .Parent.Controls("Frame0").Height + 1

      Let .Controls("Lblcnc3").Caption = "Provedor:"
      Let .Controls("Lblcnc3").Left = xLeft * 0.4
      Let .Controls("Lblcnc3").AutoSize = True
      Let .Controls("Lblcnc3").WordWrap = False
      Let .Controls("Lblcnc3").Top = (xLeft * 0.6)

      Let .Controls("Lblcnc4").Caption = "Tipo:"
      Let .Controls("Lblcnc4").Left = .Controls("Lblcnc3").Left
      Let .Controls("Lblcnc4").AutoSize = True
      Let .Controls("Lblcnc4").WordWrap = False
      Let .Controls("Lblcnc4").Top = .Controls("Lblcnc3").Top + .Controls("Lblcnc3").Height + (xLeft * 0.5)

      Let .Controls("Txtcnc4").Left = .Controls("Lblcnc3").Left + .Controls("Lblcnc3").Width + 2
      Let .Controls("Txtcnc4").Top = .Controls("Lblcnc3").Top '+ .Controls("Lblcnc3").Top + .Controls("Lblcnc3").Height ' .Controls("Lblcnc3").Top
      Let .Controls("Txtcnc4").Width = .Controls("Txtcnc4").Width * 2

      Let .Controls("Cbocnc0").Left = .Controls("Txtcnc4").Left
      Let .Controls("Cbocnc0").Top = .Controls("Lblcnc4").Top
      Let .Controls("Cbocnc0").Width = .Controls("Txtcnc4").Width

      Let .Controls("Txtcnc4").ControlTipText = "Nombres y apellidos del Provedor"
      Let .Controls("Cbocnc0").ControlTipText = "Seleccione el tipo de Provedor"
    End With

    Let StrAux = "Frame2"
    Call VIEWS.ViewFormAddControl(.Controls(StrAux), "Label", "Lblcnc", 8, CtlFromX:=5)
    Call VIEWS.ViewFormAddControl(.Controls(StrAux), "CommandButton", "Btncnc", 1, CtlFromX:=1) 'EVENT all
    Call VIEWS.ViewFormAddControl(.Controls(StrAux), "TextBox", "Txtcnc", 8, CtlFromX:=5)
    With .Controls(StrAux)
      Let .Caption = "Datos Concentrado"
      Let .Width = (.Parent.InsideWidth * 0.47)
      Let .Height = .Parent.InsideHeight * 0.2
      Let .Left = .Parent.Controls("Frame0").Left
      Let .Top = .Parent.Controls("Frame0").Top + .Parent.Controls("Frame0").Height + 1

      Let .Controls("Lblcnc5").Caption = "Fecha Ingreso:"
      Let .Controls("Lblcnc5").Left = xLeft * 0.4
      Let .Controls("Lblcnc5").AutoSize = True
      Let .Controls("Lblcnc5").WordWrap = False
      Let .Controls("Lblcnc5").Top = (xLeft * 0.3)
      
      Let .Controls("Txtcnc5").Left = .Controls("Lblcnc5").Left
      Let .Controls("Txtcnc5").Top = .Controls("Lblcnc5").Top + .Controls("Lblcnc5").Height
      Let .Controls("Txtcnc5").ControlTipText = "Ejemplo: 01/01/21 o seleccione una fecha mediante el botón"
      Let .Controls("Txtcnc5").Value = VBA.Format$(VBA.Date(), "dd/mm/yy")
      
      Let .Controls("Btncnc1").Left = .Controls("Txtcnc5").Left + .Controls("Txtcnc5").Width
      Let .Controls("Btncnc1").Top = .Controls("Txtcnc5").Top
      Let .Controls("Btncnc1").Height = .Controls("Txtcnc5").Height
      Let .Controls("Btncnc1").Width = .Controls("Btncnc1").Height
      Let .Controls("Btncnc1").Caption = VBA.ChrW(&H2705)
      Let .Controls("Btncnc1").Font.Bold = True
      Let .Controls("Btncnc1").ControlTipText = "Pulse para seleccionar una fecha"
      
      
      Let .Controls("Lblcnc6").Caption = "Peso Húmedo:"
      Let .Controls("Lblcnc6").Left = .Controls("Lblcnc5").Left
      Let .Controls("Lblcnc6").AutoSize = True
      Let .Controls("Lblcnc6").WordWrap = False
      Let .Controls("Lblcnc6").Top = (.Controls("Txtcnc5").Top + .Controls("Txtcnc5").Height) + 3
      
      Let .Controls("Txtcnc6").Left = .Controls("Lblcnc6").Left
      Let .Controls("Txtcnc6").Top = .Controls("Lblcnc6").Top + .Controls("Lblcnc6").Height
      Let .Controls("Txtcnc6").ControlTipText = "Ejemplo: 100"
      Let .Controls("Txtcnc6").TextAlign = fmTextAlignRight


      Let .Controls("Lblcnc7").Caption = "Porcentaje de humedad:"
      Let .Controls("Lblcnc7").Left = (.Controls("Txtcnc6").Left + .Controls("Txtcnc6").Width) * 1.3
      Let .Controls("Lblcnc7").AutoSize = True
      Let .Controls("Lblcnc7").WordWrap = False
      Let .Controls("Lblcnc7").Top = .Controls("Lblcnc5").Top

      Let .Controls("Txtcnc7").Left = .Controls("Lblcnc7").Left
      Let .Controls("Txtcnc7").Top = (.Controls("Lblcnc7").Top + .Controls("Lblcnc7").Height)
      Let .Controls("Txtcnc7").ControlTipText = "Ingrese el equivalente al tipo de cambio actual"
      
      Let .Controls("Lblcnc8").Caption = "Porcentaje de Merma:"
      Let .Controls("Lblcnc8").Left = .Controls("Lblcnc7").Left
      Let .Controls("Lblcnc8").AutoSize = True
      Let .Controls("Lblcnc8").WordWrap = False
      Let .Controls("Lblcnc8").Top = .Controls("Lblcnc6").Top
    
      Let .Controls("Txtcnc8").Left = .Controls("Lblcnc8").Left
      Let .Controls("Txtcnc8").Top = (.Controls("Lblcnc8").Top + .Controls("Lblcnc8").Height)
      Let .Controls("Txtcnc8").ControlTipText = "Ingrese el equivalente al tipo de cambio actual"
    End With

    Let StrAux = "Frame3"
    Call VIEWS.ViewFormAddControl(.Controls(StrAux), "Label", "Lblcnc", 11, CtlFromX:=9)
    Call VIEWS.ViewFormAddControl(.Controls(StrAux), "TextBox", "Txtcnc", 11, CtlFromX:=9)
    With .Controls(StrAux)
      Let .Caption = "Costo Tratamiento"
      Let .Width = (.Parent.InsideWidth * 0.47)
      Let .Height = .Parent.InsideHeight * 0.2
      Let .Left = .Parent.Controls("Frame1").Left
      Let .Top = .Parent.Controls("Frame0").Top + .Parent.Controls("Frame0").Height + 1

      Let .Controls("Lblcnc9").Caption = "Maquila:"
      Let .Controls("Lblcnc9").Left = xLeft * 0.4
      Let .Controls("Lblcnc9").AutoSize = True
      Let .Controls("Lblcnc9").WordWrap = False
      Let .Controls("Lblcnc9").Top = (xLeft * 0.4)

      Let .Controls("Lblcnc10").Caption = "Base:"
      Let .Controls("Lblcnc10").Left = .Controls("Lblcnc9").Left
      Let .Controls("Lblcnc10").AutoSize = True
      Let .Controls("Lblcnc10").WordWrap = False
      Let .Controls("Lblcnc10").Top = .Controls("Lblcnc9").Top + .Controls("Lblcnc9").Height + (xLeft * 0.5)

      Let .Controls("Lblcnc11").Caption = "Escalador:"
      Let .Controls("Lblcnc11").Left = .Controls("Lblcnc10").Left
      Let .Controls("Lblcnc11").AutoSize = True
      Let .Controls("Lblcnc11").WordWrap = False
      Let .Controls("Lblcnc11").Top = .Controls("Lblcnc10").Top + .Controls("Lblcnc10").Height + (xLeft * 0.5)


      Let .Controls("Txtcnc9").Left = .Controls("Lblcnc11").Left + .Controls("Lblcnc11").Width + 2
      Let .Controls("Txtcnc9").Top = .Controls("Lblcnc9").Top
      Let .Controls("Txtcnc9").Width = .Controls("Txtcnc9").Width * 2

      Let .Controls("Txtcnc10").Left = .Controls("Txtcnc9").Left
      Let .Controls("Txtcnc10").Top = .Controls("Txtcnc9").Top + .Controls("Txtcnc9").Height
      Let .Controls("Txtcnc10").Width = .Controls("Txtcnc9").Width

      Let .Controls("Txtcnc11").Left = .Controls("Txtcnc10").Left
      Let .Controls("Txtcnc11").Top = .Controls("Txtcnc10").Top + .Controls("Txtcnc10").Height
      Let .Controls("Txtcnc11").Width = .Controls("Txtcnc10").Width

      Let .Controls("Txtcnc9").ControlTipText = "El valor de Maquila. Ejemplo: 150"
      Let .Controls("Txtcnc10").ControlTipText = "La cotización base. Ejemplo: 2000"
      Let .Controls("Txtcnc11").ControlTipText = "La penalidad aplicable. Ejemplo: 0.15"
    End With
        
    Let StrAux = "Frame4"
    Call VIEWS.ViewFormAddControl(.Controls(StrAux), "ComboBox", "Cbocnc", 3, CtlFromX:=1) 'EVENT all
    Call VIEWS.ViewFormAddControl(.Controls(StrAux), "CommandButton", "Btncnc", 8, CtlFromX:=2) 'EVENT all
    Call VIEWS.ViewFormAddControl(.Controls(StrAux), "TextBox", "Txtcnc", 15, CtlFromX:=12)
    Call VIEWS.ViewFormAddControl(.Controls(StrAux), "ListBox", "Lstcnc", 0) 'EVENT
    Call VIEWS.ViewFormAddControl(.Controls(StrAux), "CheckBox", "Chkcnc", 3, CtlFromX:=1)
    With .Controls(StrAux)
      Let .Caption = "Datos Liquidación" 'Dónde mostrar resultados"
      Let .Width = (.Parent.InsideWidth * 0.9525)
      Let .Height = .Parent.InsideHeight * 0.4
      Let .Left = .Parent.Controls("Frame0").Left
      Let .Top = .Parent.Controls("Frame3").Top + .Parent.Controls("Frame3").Height + 1
      
      Let .Controls("Cbocnc1").Left = xLeft * 0.4
      Let .Controls("Cbocnc1").Top = xLeft * 0.4
      Let .Controls("Cbocnc1").Width = (.Width * 0.95) - ((.Controls("Cbocnc1").Height * 7) + 1) '6 Buttons
      Let .Controls("Cbocnc1").Visible = True
      Let .Controls("Cbocnc1").ControlTipText = "Seleccione una opción e ingrese los datos respectivos en la lista"

      For i = 2 To 8
        Let .Controls("Btncnc" & i).Top = .Controls("Cbocnc1").Top
        Let .Controls("Btncnc" & i).Height = .Controls("Cbocnc1").Height
        Let .Controls("Btncnc" & i).Width = .Controls("Btncnc" & i).Height
        Let .Controls("Btncnc" & i).Left = (.Controls("Cbocnc1").Left + .Controls("Cbocnc1").Width) + (.Controls("Btncnc" & i).Width * (i - 2))
      Next i
      Let .Controls("Btncnc2").Caption = VBA.ChrW(&H271A)
      Let .Controls("Btncnc2").ControlTipText = "Pulse para añadir los valores de los controles a la lista"
      Let .Controls("Btncnc3").Caption = VBA.ChrW(&H2710)
      Let .Controls("Btncnc3").ControlTipText = "Pulse para confirmar la edición de los valores que modifique"
      Let .Controls("Btncnc4").Caption = VBA.ChrW(&H270F)
      Let .Controls("Btncnc4").ControlTipText = "Cancelar edición"
      Let .Controls("Btncnc5").Caption = VBA.ChrW(&H2A31)
      Let .Controls("Btncnc5").ControlTipText = "Pulse para eliminar la fila de datos seleccionada en la lista"
      Let .Controls("Btncnc4").Font.Bold = True
      Let .Controls("Btncnc3").Font.Size = 10
      Let .Controls("Btncnc5").Font.Bold = True

      Let .Controls("Btncnc6").Caption = VBA.ChrW(&H2716)
      Let .Controls("Btncnc6").ControlTipText = "Pulse para eliminar todos los datos que ingresó"

      Let .Controls("Btncnc7").Caption = VBA.ChrW(&H2583)
      Let .Controls("Btncnc7").ControlTipText = "Seleccionar datos desde un rango de celdas de " & Application.Name
      Let .Controls("Btncnc7").Tag = "selection"

      Let .Controls("Btncnc8").Caption = VBA.ChrW(&H2583)
      Let .Controls("Btncnc8").ControlTipText = "Seleccionar Datos Base, Provedor, Concentrado y Tratamiento desde un rango de celdas de " & Application.Name
      Let .Controls("Btncnc8").Tag = "selection"

      Let .Controls("Btncnc3").Enabled = False
      Let .Controls("Btncnc4").Enabled = False
      Let .Controls("Btncnc5").Enabled = False
      Let .Controls("Btncnc6").Enabled = False
      
      Let .Controls("Cbocnc2").Width = ((.Width * 0.95) / 4)
      Let .Controls("Txtcnc12").Width = .Controls("Cbocnc2").Width
      Let .Controls("Txtcnc13").Width = .Controls("Cbocnc2").Width
      Let .Controls("Txtcnc14").Width = .Controls("Cbocnc2").Width
      Let .Controls("Txtcnc15").Width = .Controls("Cbocnc2").Width
      
      Let .Controls("Cbocnc2").Left = .Controls("Cbocnc1").Left
      Let .Controls("Txtcnc12").Left = (.Controls("Cbocnc2").Left + .Controls("Cbocnc2").Width)
      Let .Controls("Txtcnc13").Left = (.Controls("Txtcnc12").Left + .Controls("Txtcnc12").Width)
      Let .Controls("Txtcnc14").Left = (.Controls("Txtcnc13").Left + .Controls("Txtcnc13").Width)
      Let .Controls("Txtcnc15").Left = (.Controls("Txtcnc14").Left + .Controls("Txtcnc14").Width)

      Let .Controls("Cbocnc2").Top = (.Controls("Cbocnc1").Top + .Controls("Cbocnc1").Height)
      Let .Controls("Txtcnc12").Top = .Controls("Cbocnc2").Top
      Let .Controls("Txtcnc13").Top = .Controls("Cbocnc2").Top
      Let .Controls("Txtcnc14").Top = .Controls("Cbocnc2").Top
      Let .Controls("Txtcnc15").Top = .Controls("Cbocnc2").Top
      Let .Controls("Cbocnc3").Top = .Controls("Cbocnc2").Top

      Let .Controls("Txtcnc15").Visible = False

      Let .Controls("Lstcnc0").Left = .Controls("Cbocnc1").Left
      Let .Controls("Lstcnc0").Top = .Controls("Cbocnc2").Top + .Controls("Cbocnc2").Height
      Let .Controls("Lstcnc0").Width = (.Width * 0.95)
    
      Let .Controls("Chkcnc1").Visible = False
      Let .Controls("Chkcnc1").AutoSize = True
      Let .Controls("Chkcnc1").WordWrap = False
      Let .Controls("Chkcnc1").Caption = "Valor directo"
      Let .Controls("Chkcnc1").Accelerator = "v"
      Let .Controls("Chkcnc1").ControlTipText = "La deducción será directa por el VALOR que especifique"
      Let .Controls("Chkcnc1").Top = .Controls("Cbocnc3").Top
      Let .Controls("Chkcnc1").Left = .Controls("Txtcnc15").Left
      
      Let .Controls("Chkcnc2").Left = .Controls("Cbocnc1").Left
      Let .Controls("Chkcnc2").Top = (.Controls("Lstcnc0").Top + .Controls("Lstcnc0").Height) * 1.05
      Let .Controls("Chkcnc2").AutoSize = True
      Let .Controls("Chkcnc2").WordWrap = False
      Let .Controls("Chkcnc2").Accelerator = "p"
      Let .Controls("Chkcnc2").Caption = "Autoconfigurar página"
      Let .Controls("Chkcnc2").ControlTipText = "Configurará la página con los valores predeterminados por el Sistema"
    
      Let .Controls("Chkcnc3").Left = .Controls("Chkcnc2").Left + .Controls("Chkcnc2").Width
      Let .Controls("Chkcnc3").Top = .Controls("Chkcnc2").Top
      Let .Controls("Chkcnc3").AutoSize = True
      Let .Controls("Chkcnc3").WordWrap = False
      Let .Controls("Chkcnc3").Accelerator = "d"
      Let .Controls("Chkcnc3").Caption = "Dos copias por página"
      Let .Controls("Chkcnc3").ControlTipText = "Generará dos copias en una página"
    End With
    
    With .Controls("Optcnc0") 'New Sheet
      Let .WordWrap = False
      Let .AutoSize = True
      Let .Top = (.Parent.Controls("Frame4").Top + .Parent.Controls("Frame4").Height) * 1.01
      Let .Left = (.Parent.Width * (0.125 / 2))
      Let .Caption = "Mostrar resultado en hoja nueva"
      Let .Value = True
    End With

    With .Controls("Optcnc1") 'New Sheet
      Let .WordWrap = False
      Let .AutoSize = True
      Let .Top = .Parent.Controls("Optcnc0").Top
      Let .Left = (.Parent.Width * 0.4)
      Let .Caption = "Mostrar resultado en hoja actual"
    End With

    With .Controls("Txtcnc0") 'New Sheet
      Let .Top = .Parent.Controls("Optcnc1").Top
      Let .Left = (.Parent.Controls("Optcnc1").Left + .Parent.Controls("Optcnc1").Width)
    End With

    With .Controls("Btnmain3") 'All pur cnc data
      Let .Height = .Parent.Controls("Txtcnc0").Height
      Let .Width = .Height
      Let .Top = .Parent.Controls("Txtcnc0").Top
      Let .Left = (.Parent.Controls("Txtcnc0").Left + .Parent.Controls("Txtcnc0").Width) '- .Width
      Let .Caption = VBA.ChrW(&H2583)
      Let .ControlTipText = "Pulse para seleccionar una celda donde generar el resultado"
    End With


    '2. [ ] Fill controls with values
    Let .Controls("Cbocnc0").List = Array("Cooperativista", "Particular")
    Let .Controls("Cbocnc1").List = Array("Contenidos y condiciones", "Refinación", "Penalidades", "Deducciones", "Bonos")
    Let .Controls("Cbocnc3").List = Array("Y", "O")

    Let .Controls("Lstcnc0").ColumnHeads = True
    Let .Controls("Cbocnc0").ListIndex = 0
    Let .Controls("Cbocnc3").ListIndex = 0

    '3. [x] Fill Form Controls with data and Set general purpose CommandButtons
    Let k = (.InsideWidth * 0.95)
    With .Controls("Btnmain0") '("BtnCancel")
      Let .Cancel = True
      Let .Accelerator = "c"
      Let .Top = (.Parent.InsideHeight * 0.97) - (.Height)
      Let .Left = k - .Width + xLeft * 0.5
      Let .Caption = "Cerrar"
    End With
    With .Controls("Btnmain2") '("BtnHelp")
      Let .Caption = "Ayuda"
      Let .Accelerator = "a"
      Let .ControlTipText = "Ver ayuda del Ejercicio"
      Let .Top = .Parent.Controls("Btnmain0").Top
      Let .Left = xLeft * 0.45
      Let .Tag = ControlId '& cGet
    End With
    With .Controls("Btnmain1") '("BtnEXE")
      Let .Caption = "Generar"
      Let .Accelerator = "g"
      Let .ControlTipText = "Permite procesar los datos y mostrar los resultados en una hoja Excel."
      Let .Top = .Parent.Controls("Btnmain0").Top
      Let .Left = .Parent.Controls("Btnmain0").Left - .Width - 2
      Let .Tag = ControlId & "-bal"
      Let .Default = True
    End With

    '4. Assing handler for each button
    Set BtnCancel = .Controls("Btnmain0") 'Main
    Set BtnAccept = .Controls("Btnmain1") 'Main
    Set Btn0 = .Controls("Btnmain2")      'Main
    Set Btn1 = .Controls("Btnmain3")      'Main (Rng)
    Set Btn2 = .Controls("Btncnc0") 'Frame0
    Set Btn3 = .Controls("Btncnc1") 'Frame2
    Set Btn4 = .Controls("Btncnc2") 'Frame4
    Set Btn5 = .Controls("Btncnc3") 'Frame4
    Set Btn6 = .Controls("Btncnc4") 'Frame4
    Set Btn7 = .Controls("Btncnc5") 'Frame4
    Set Btn8 = .Controls("Btncnc6") 'Frame4
    Set Btn9 = .Controls("Btncnc7") 'Frame4
    Set Btn10 = .Controls("Btncnc8") 'Frame4

    Set Cbo0 = .Controls("Cbocnc1") 'Frame4
    Set Cbo1 = .Controls("Cbocnc2") 'Frame4
'    Set Cbo2 = .Controls("Cbopur3") 'Frame2
'    Set Cbo3 = .Controls("Cbopur4") 'Frame2

    Set Lst0 = .Controls("Lstcnc0") 'Frame4

    Set DPR = New ClassDatePicker

    '5. Give styles to controls
    Let .Controls("Cbocnc1").ListIndex = 0
    Call VIEWS.ViewFormStyles(FRM): Set VIEWS = Nothing

    '6. [x] Show the finished Form
    Call Load(FRM)
    Call .Show

'    '7. [x] Close form
'    Let .Controls("Lstpur0").RowSource = VBA.vbNullString
'    Let .Controls("Cbopur0").List = Empty
'    Let .Controls("Cbopur1").List = Empty
'    Let .Controls("Cbopur2").List = Empty
'    Let .Controls("Cbopur3").List = Empty
'    Let .Controls("Cbopur4").List = Empty
  End With

EH:
  Erase aLabels
  Erase aMaxLengths
  Erase aControlis
  Erase aControls
  Set DPR = Nothing
  Set VIEWS = Nothing
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)
End Sub

Private Sub View_form_config(ByVal ControlId As String, ByVal FRM As FRMBasics)
  'Main objects
  Dim RES As AppResources
  Dim VIEWS As ViewsApp
  'ListObjects
  'LO data names and other data
  'Controls
  Dim aLabels() As Variant
  Dim aMaxLengths() As Variant
  Dim aControls() As Variant
  Dim aControlis() As Variant
  'Standard vars
  Dim sTitle As String, TitleForm As String, StrAux As String, StrAux0 As String
  Dim i As Long, j As Long, k As Long, l As Long, t As Long
  Dim xLeft As Long
  Dim StrDivlocal As String, StrDivforeign As String

  
  Const cLeft As Single = 10
  Const sFN As String = "VIEW::-form_config"

  #If Debugging Then
    Call Debug_assert(Not ControlId = VBA.vbNullString, sFN)
    Call Debug_assert(Not FRM Is Nothing, sFN)
  #End If
  
  If ControlId = VBA.vbNullString Then GoTo EH
  If FRM Is Nothing Then GoTo EH
  On Error GoTo EH
  
  'Get data for controls

  'Set Form basic datas
  Let TitleForm = "Configuración"
  Let sTitle = TitleForm & " - " & STRAPPTITLE
  
  Set VIEWS = New ViewsApp
  Set VIEWS.ErrorHandler = EHGLOBAL
  Call VIEWS.ViewFormCleanner(FRM)
  With FRM
    'Set the form dimensions
    Let .Height = 415
    Let .Width = 450
    Let .Tag = ControlId
    Let .Caption = sTitle
    Let xLeft = .InsideWidth * 0.05

    'Build and design the form
    Call VIEWS.ViewFormAddControl(FRM, "CommandButton", "Btnmain", 0) 'All with events (Cancel and Help)
    Call VIEWS.ViewFormAddControl(FRM, "MultiPage", "Multi", 0, False)
    
    'Dim multi As MSForms.MultiPage
    With .Controls("Multi") 'multi
      Call .Pages.Add
      Call .Pages.Add
      Let aLabels = Array("Divisas", "Selecciones", "Página", "Liquidaciones")
      For i = 0 To .Pages.Count - 1
        Let .Pages(i).Caption = aLabels(i)
        Let .Pages(i).Accelerator = VBA.Left(VBA.LCase(VBA.CStr(aLabels(i))), 1)
      Next i
      Let .Width = .Parent.InsideWidth
      Let .Height = (.Parent.InsideHeight * 0.85)
    End With

    
    
    'Add controls to Page 0 - Divisa
    Let j = 0
    Call VIEWS.ViewFormAddControl(.Controls("Multi").Pages(j), "Label", "Lblcon", 4)
    Call VIEWS.ViewFormAddControl(.Controls("Multi").Pages(j), "TextBox", "Txtcon", 2)
    Call VIEWS.ViewFormAddControl(.Controls("Multi").Pages(j), "ListBox", "Lstcon", 0)
    Call VIEWS.ViewFormAddControl(.Controls("Multi").Pages(j), "CommandButton", "Btncon", 5) 'EVENT all
    Call VIEWS.ViewFormAddControl(.Controls("Multi").Pages(j), "ComboBox", "Cbocon", 1)
    
    Let i = ((.Controls("Multi").Width - (xLeft * 2.3)) / 3)
      
    For j = 0 To 2
      Let t = j * 2
      Let .Controls("Txtcon" & j).Width = i
      Let .Controls("Txtcon" & j).Left = xLeft + ((.Controls("Txtcon" & j).Left + .Controls("Txtcon" & j).Width) * j)
      Let .Controls("Txtcon" & j).Top = xLeft + (.Controls("Txtcon" & j).Height * 2)
      Let .Controls("Txtcon" & j).TabIndex = (t + 1)
      Let .Controls("Txtcon" & j).MaxLength = 30
      Let k = .Controls("Txtcon" & j).Left
      Let l = .Controls("Txtcon" & j).Top
      With .Controls("Lblcon" & j)
        Let .Left = k
        Let .Width = i
        Let .WordWrap = False
        Let .Top = l - (xLeft * 0.75)
        Let .TabIndex = t
      End With
    Next j
    Let .Controls("Txtcon2").MaxLength = 3
    
    Let k = (.Controls("Txtcon2").Left + .Controls("Txtcon2").Width) - (.Controls("Txtcon2").Height)
    For j = 0 To 3
      Let .Controls("Btncon" & j).Width = .Controls("Txtcon0").Height
      Let .Controls("Btncon" & j).Height = .Controls("Txtcon0").Height
      Let .Controls("Btncon" & j).Left = (k - (.Controls("Btncon" & j).Width * j))
      Let .Controls("Btncon" & j).Top = (.Controls("Txtcon2").Top - .Controls("Btncon" & j).Height)
      Let .Controls("Btncon" & j).TabIndex = 12 + j
    Next j
    Let .Controls("Btncon0").Caption = VBA.ChrW(&H271A)
    Let .Controls("Btncon0").ControlTipText = "Añadir los datos al registro"
    Let .Controls("Btncon1").Caption = VBA.ChrW(&H2710)
    Let .Controls("Btncon1").ControlTipText = "Editar el registro seleccionado"
    Let .Controls("Btncon2").Caption = VBA.ChrW(&H270F)
    Let .Controls("Btncon2").ControlTipText = "Cancelar edición"
    Let .Controls("Btncon3").Caption = VBA.ChrW(&H2A31)
    Let .Controls("Btncon3").ControlTipText = "Eliminar registro"
    Let .Controls("Btncon1").Enabled = False
    Let .Controls("Btncon2").Enabled = False
    Let .Controls("Btncon3").Enabled = False
    
    With .Controls("Lblcon0")
      Let .Caption = "Divisa en singular:"
      Let .Accelerator = "i"
    End With
    With .Controls("Lblcon1")
      Let .Caption = "Divisa en plural:"
      Let .Accelerator = "u"
    End With
    With .Controls("Lblcon2")
      Let .Caption = "Símbolo:"
      Let .Accelerator = "m"
      Let .AutoSize = True
    End With
    
    With .Controls("Lstcon0")
      Let .Top = (FRM.Controls("Txtcon0").Top + FRM.Controls("Txtcon0").Height)
      Let .Left = FRM.Controls("Txtcon0").Left
      Let .Height = FRM.Controls("Multi").Height * 0.5
      Let .Width = i * 3
      Let .TabIndex = 11
    End With
    
    With .Controls("Lblcon3")
      Let .Caption = "Moneda local:"
      Let .Accelerator = "o"
      Let .WordWrap = False
      Let .AutoSize = True
      Let .Left = FRM.Controls("Lstcon0").Left
      Let .Top = FRM.Controls("Lstcon0").Top + FRM.Controls("Lstcon0").Height + (xLeft * 0.5)
      Let .TabIndex = 7
    End With
    With .Controls("Lblcon4")
      Let .Caption = "Moneda Extranjera:"
      Let .Accelerator = "j"
      Let .WordWrap = False
      Let .AutoSize = True
      Let .Left = FRM.Controls("Lblcon3").Left
      Let .Top = FRM.Controls("Lblcon3").Top + FRM.Controls("Lblcon3").Height + (xLeft * 0.5)
      Let .TabIndex = 9
    End With
    
    With .Controls("Cbocon0")
      Let .Left = FRM.Controls("Lblcon4").Left + FRM.Controls("Lblcon4").Width + 5
      Let .Top = FRM.Controls("Lblcon3").Top
      Let .Width = .Width * 2
      Let .TabIndex = 8
    End With
    With .Controls("Cbocon1")
      Let .Left = FRM.Controls("Cbocon0").Left
      Let .Top = FRM.Controls("Lblcon4").Top
      Let .Width = .Width * 2
      Let .TabIndex = 10
    End With
    
    With .Controls("Btncon4")
      Let .Caption = "Establecer"
      Let .Accelerator = "t"
      Let .Left = FRM.Controls("Cbocon0").Left + FRM.Controls("Cbocon0").Width + 5
      Let .Top = FRM.Controls("Cbocon0").Top
      Let .Height = FRM.Controls("Cbocon0").Height
    End With
    With .Controls("Btncon5")
      Let .Caption = "Establecer"
      Let .Accelerator = "r"
      Let .Left = FRM.Controls("Btncon4").Left
      Let .Top = FRM.Controls("Cbocon1").Top
      Let .Height = FRM.Controls("Cbocon0").Height
    End With




    'Add controls to Page 1 - Selecciones
    Let j = 1
    Call VIEWS.ViewFormAddControl(.Controls("Multi").Pages(j), "CheckBox", "Chkcon", 1)
    
    With .Controls("Chkcon0")
      Let .Caption = "Resultados compactos en Balances metalúrgicos y Regresión lineal Selección"
      Let .Accelerator = "r"
      Let .WordWrap = False
      Let .AutoSize = True
      Let .Left = xLeft
      Let .Top = FRM.Controls("Multi").Height / 6
    End With
    With .Controls("Chkcon1")
      Let .Caption = "Mostrar Resultados de Cabeza calculada en Balances metalúrgicos"
      Let .Accelerator = "z"
      Let .WordWrap = False
      Let .AutoSize = True
      Let .Left = xLeft
      Let .Top = FRM.Controls("Multi").Height / 4
    End With
    


    'Add controls to Page 2 - Página
    Let j = 2
    Call VIEWS.ViewFormAddControl(.Controls("Multi").Pages(j), "Label", "Lblcon", 10, CtlFromX:=5)
    Call VIEWS.ViewFormAddControl(.Controls("Multi").Pages(j), "TextBox", "Txtcon", 7, CtlFromX:=3)
    Call VIEWS.ViewFormAddControl(.Controls("Multi").Pages(j), "CheckBox", "Chkcon", 2, CtlFromX:=2)
    Call VIEWS.ViewFormAddControl(.Controls("Multi").Pages(j), "CommandButton", "Btncon", 10, CtlFromX:=6) 'EVENT all
    Call VIEWS.ViewFormAddControl(.Controls("Multi").Pages(j), "ComboBox", "Cbocon", 7, CtlFromX:=2)

    For j = 5 To 9
      Let t = (j - 5) * 2
      With .Controls("Lblcon" & j)
        Let .TabIndex = t
        Let .AutoSize = True
        Let .WordWrap = False
        Let .Left = xLeft
        Let .Top = xLeft + ((FRM.Controls("Txtcon3").Height * 1.05) * (j - 5))
        Let l = .Top
      End With
      With .Controls("Txtcon" & (j - 2))
        Let .Width = .Width * 2
        Let .Top = l
        Let .TabIndex = (t + 1)
      End With
    Next j
    Let .Controls("Lblcon5").Caption = "Logo:"
    Let .Controls("Lblcon6").Caption = "Empresa:"
    Let .Controls("Lblcon7").Caption = "Dirección:"
    Let .Controls("Lblcon8").Caption = "Teléfono:"
    Let .Controls("Lblcon9").Caption = "Correo electrónico:"
    Let .Controls("Lblcon5").Accelerator = "o"
    Let .Controls("Lblcon6").Accelerator = "E:"
    Let .Controls("Lblcon7").Accelerator = "n:"
    Let .Controls("Lblcon8").Accelerator = "t:"
    Let .Controls("Lblcon9").Accelerator = "r"

    Let k = .Controls("Lblcon9").Left + .Controls("Lblcon9").Width + 5
    For j = 3 To 7
      With .Controls("Txtcon" & j)
        Let .Left = k
      End With
    Next j

    Let .Controls("Txtcon3").Locked = True
    
    With .Controls("Btncon6")
      Let .Caption = "..."
      Let .ControlTipText = "Elegir imagen desde archivo"
      Let .Font.Bold = True
      Let .Height = FRM.Controls("Txtcon3").Height
      Let .Width = .Height
      Let .Top = FRM.Controls("Txtcon3").Top
      Let .Left = FRM.Controls("Txtcon3").Left + FRM.Controls("Txtcon3").Width + 2
    End With

    Let k = .Controls("Btncon6").Left + .Controls("Btncon6").Width + 2
    Let aLabels = Array("Sólo en encabezado de página", "Sólo en pie de página", "En ambos")
    For j = 2 To 6
      With .Controls("Cbocon" & j)
        Let .List = aLabels
        Let .ControlTipText = "El contenido se mostrará en..."
        Let .Left = k
        Let .Top = FRM.Controls("Txtcon" & (j + 1)).Top
        Let .Width = .Width * 2
        Let .ListIndex = 0
      End With
    Next j
    Erase aLabels
    
    With .Controls("Btncon7")
      Let .Caption = "Guardar cambios"
      Let .ControlTipText = "Guarda los cambios realizados en los datos de Página"
      Let .Accelerator = "g"
      Let .Width = .Width * 2
      Let .Top = (FRM.Controls("Cbocon6").Top + FRM.Controls("Cbocon6").Height) * 1.1
      Let .Left = FRM.Controls("Cbocon6").Left
    End With

    With .Controls("Lblcon10")
      Let .Caption = "Destinos de documento:"
      Let .Accelerator = "i"
      Let .TabIndex = 10
      Let .AutoSize = True
      Let .WordWrap = False
      Let .Left = xLeft
      Let .Top = (FRM.Controls("Btncon7").Top + FRM.Controls("Btncon7").Height) * 1.3
    End With

    With .Controls("Cbocon7")
      Let .TabIndex = 11
      Let .Left = (FRM.Controls("Lblcon10").Left + FRM.Controls("Lblcon10").Width)
      Let .Top = FRM.Controls("Lblcon10").Top
      Let .Width = .Width * 2
    End With

    Let k = .Controls("Cbocon7").Left + .Controls("Cbocon7").Width
    Let l = .Controls("Cbocon7").Height
    Let t = .Controls("Cbocon7").Top
    For j = 8 To 10
      With .Controls("Btncon" & j)
        Let .Height = l
        Let .Width = l
        Let .Left = k + (.Width * (j - 8))
        Let .Top = t
      End With
    Next j
    Let .Controls("Btncon8").Caption = VBA.ChrW(&H271A)
    Let .Controls("Btncon8").ControlTipText = "Añadir los datos al registro"
    Let .Controls("Btncon9").Caption = VBA.ChrW(&H2710)
    Let .Controls("Btncon9").ControlTipText = "Editar el registro seleccionado"
    Let .Controls("Btncon10").Caption = VBA.ChrW(&H2A31)
    Let .Controls("Btncon10").ControlTipText = "Eliminar registro"

    Let t = .Controls("Cbocon7").Top + (.Controls("Cbocon7").Height * 1.5)
    With .Controls("Chkcon2")
      Let .Caption = "Siempre configurar página automáticamente"
      Let .ControlTipText = "Establece la configuración personalizada de página en las liquidaciones"
      Let .Accelerator = "f"
      Let .WordWrap = False
      Let .AutoSize = True
      Let .Left = xLeft
      Let .Top = t
    End With



    'Add controls to Page 3 - Liquidaciones
    Let j = 3
    Call VIEWS.ViewFormAddControl(.Controls("Multi").Pages(j), "Label", "Lblcon", 11, CtlFromX:=11)
    Call VIEWS.ViewFormAddControl(.Controls("Multi").Pages(j), "TextBox", "Txtcon", 12, CtlFromX:=8)
    Call VIEWS.ViewFormAddControl(.Controls("Multi").Pages(j), "CheckBox", "Chkcon", 4, CtlFromX:=3)
    Call VIEWS.ViewFormAddControl(.Controls("Multi").Pages(j), "ListBox", "Lstcon", 1, CtlFromX:=1)
    Call VIEWS.ViewFormAddControl(.Controls("Multi").Pages(j), "CommandButton", "Btncon", 14, CtlFromX:=11) 'EVENT all
    Call VIEWS.ViewFormAddControl(.Controls("Multi").Pages(j), "ComboBox", "Cbocon", 10, CtlFromX:=8)

    With .Controls("Lblcon11")
      Let .Caption = "Seleccione una opción de la lista:"
      Let .Accelerator = "t"
      Let .Left = xLeft
      Let .Top = xLeft
      Let .WordWrap = False
      Let .AutoSize = True
      Let .TabIndex = 0
    End With
    
    Let aLabels = Array("Configurar regalías mineras", "Configurar retenciones de ley", "Configurar retenciones otras (institucionales)", "Configurar penalidades para compra-venta de concentrados minerales", "Configurar deducciones en exportaciones")
    With .Controls("Cbocon8")
      Let .List = aLabels
      Let .ControlTipText = "Seleccione una opción y cree, modifique o elimine registros mediante los botones"
      Let .Left = xLeft
      Let .Top = (FRM.Controls("Lblcon6").Top + 5) 'xLeft
      Let .Width = i * 3
      Let .ListIndex = 0
      Let .TabIndex = 1
    End With
    Erase aLabels

    Let k = .Controls("Cbocon8").Left + .Controls("Cbocon8").Width - .Controls("Cbocon8").Height
    Let l = .Controls("Cbocon8").Height
    Let t = .Controls("Cbocon8").Top + .Controls("Cbocon8").Height
    For j = 11 To 14
      With .Controls("Btncon" & j)
        Let .Height = l
        Let .Width = l
        Let .Left = k - (.Width * (j - 11))
        Let .Top = t
      End With
    Next j
    Let .Controls("Btncon11").Caption = VBA.ChrW(&H271A)
    Let .Controls("Btncon11").ControlTipText = "Añadir los datos al registro"
    Let .Controls("Btncon12").Caption = VBA.ChrW(&H2710)
    Let .Controls("Btncon12").ControlTipText = "Editar el registro seleccionado"
    Let .Controls("Btncon13").Caption = VBA.ChrW(&H270F)
    Let .Controls("Btncon13").ControlTipText = "Cancelar edición"
    Let .Controls("Btncon14").Caption = VBA.ChrW(&H2A31)
    Let .Controls("Btncon14").ControlTipText = "Eliminar registro"

    'JUST TO REMEMBER THAT MS IS A $#%@!, DOESN'T SUPPORTS AfterUpdate EVENT(S) IN DYNAMIC CONTROLS
    'Let t = .Controls("Btncon11").Top + (.Controls("Btncon11").Height * 0.25)
    'With .Controls("Lblcon12")
    '  Let .Caption = "X" 'VBA.vbNullString
    '  Let .Left = xLeft
    '  Let .Top = t
    '  Let .WordWrap = False
    '  Let .AutoSize = True
    'End With

    Let aLabels = Array("Txtcon8", "Txtcon9", "Cbocon9", "Txtcon10", "Txtcon11", "Txtcon12", "Cbocon10")
    Let i = .Controls("Cbocon8").Width
    Let k = (i / 7)
    Let t = .Controls("Btncon11").Top + .Controls("Btncon11").Height
    Let l = .Controls("Cbocon8").Left
    For j = 0 To UBound(aLabels)
      With .Controls(aLabels(j))
        Let .Width = k
        Let .Left = l + (k * j) - 1.1
        Let .Top = t
      End With
    Next j
    Erase aLabels
    
    Let t = .Controls("Txtcon8").Top + .Controls("Txtcon8").Height
    With .Controls("Lstcon1")
      Let .Height = FRM.Controls("Multi").Height * 0.5
      Let .Width = i
      Let .Left = l
      Let .Top = t
    End With
    
    Let t = .Controls("Lstcon1").Top + (.Controls("Lstcon1").Height * 1.1)
    With .Controls("Chkcon3")
      Let .Caption = "Resumir regalías en las liquidaciones"
      Let .ControlTipText = "Suma los porcentajes de las regalías y genera un solo resultado en las liquidaciones"
      Let .Accelerator = "m"
      Let .WordWrap = False
      Let .AutoSize = True
      Let .Left = l
      Let .Top = t
    End With

    Let l = .Controls("Cbocon10").Left
    With .Controls("Chkcon4")
      Let .Caption = "Valor directo"
      Let .ControlTipText = "La cantidad de la columna VALOR será descontada directamente en las liquidaciones"
      Let .Accelerator = "v"
      Let .WordWrap = False
      Let .AutoSize = True
      Let .Visible = False
      Let .Left = l
      Let .Top = t
    End With



    '2. [ ] Fill controls with values
    With .Controls("Lstcon0")
      Let .RowSource = excliqlitedatasheet.ListObjects("excliqlitecurrenciescon")
      Let .ColumnCount = 3
      Let FRM.Controls("Btncon3").Enabled = (.ListCount > 2)
    End With
    
    Let j = 0: Let k = 0
    If .Controls("Lstcon0").ListCount > 0 Then
      For i = 0 To .Controls("Lstcon0").ListCount - 1
        Let StrAux0 = .Controls("Lstcon0").List(i, 0) & " [" & .Controls("Lstcon0").List(i, 2) & "]"
        Call .Controls("Cbocon0").AddItem(StrAux0)
        Call .Controls("Cbocon1").AddItem(StrAux0)
        If VBA.CBool(.Controls("Lstcon0").List(i, 3)) Then Let j = i
        If VBA.CBool(.Controls("Lstcon0").List(i, 4)) Then Let k = i
      Next i
      Let .Controls("Cbocon0").ListIndex = j
      Let .Controls("Cbocon1").ListIndex = k
      Let .Controls("Btncon4").Enabled = (.Controls("Cbocon0").ListCount > 0)
      Let .Controls("Btncon5").Enabled = (.Controls("Cbocon1").ListCount > 0)
    End If
    
    Set RES = New AppResources
    Let .Controls("Chkcon0").Value = RES.resAPPEQGETTABLESELECTIONSMODSCOMPACT()
    Let .Controls("Chkcon1").Value = RES.resAPPEQGETTABLESELECTIONSMODSHEADS()
    Let .Controls("Chkcon2").Value = RES.resAPPEQGETTABLE_PAGEOPTIONS_EXISTS()
    Let .Controls("Chkcon3").Value = RES.resAPPEQGETTABLE_PUROPTIONS_RMDETAIL()
    
    Let .Controls("Txtcon3").Value = RES.resAPPEQGETTABLE_PAGEOPTIONS_LOGO()
    Let .Controls("Txtcon4").Value = RES.resAPPEQGETTABLE_PAGEOPTIONS_COMPANY()
    Let .Controls("Txtcon5").Value = RES.resAPPEQGETTABLE_PAGEOPTIONS_ADDRESS()
    Let .Controls("Txtcon6").Value = RES.resAPPEQGETTABLE_PAGEOPTIONS_PHONE()
    Let .Controls("Txtcon7").Value = RES.resAPPEQGETTABLE_PAGEOPTIONS_EMAIL()
    Let .Controls("Cbocon2").ListIndex = RES.resAPPEQGETTABLE_PAGEOPTIONS_LOGO_EXISTS()
    Let .Controls("Cbocon3").ListIndex = RES.resAPPEQGETTABLE_PAGEOPTIONS_COMPANY_EXISTS()
    Let .Controls("Cbocon4").ListIndex = RES.resAPPEQGETTABLE_PAGEOPTIONS_ADDRESS_EXISTS()
    Let .Controls("Cbocon5").ListIndex = RES.resAPPEQGETTABLE_PAGEOPTIONS_PHONE_EXISTS()
    Let .Controls("Cbocon6").ListIndex = RES.resAPPEQGETTABLE_PAGEOPTIONS_EMAIL_EXISTS()
    Set RES = Nothing
    
    Let .Controls("Cbocon7").RowSource = excliqlitedatasheet.ListObjects("excliqliteconfigpurdestiny")
    Let .Controls("Cbocon7").ListIndex = 0

    Let .Controls("Cbocon9").List = Array("L.F.", "O.T.", "T.M.F.")
    Let .Controls("Cbocon9").ListIndex = 0
    Let .Controls("Cbocon10").List = Array("%", "g/T", "DM")
    Let .Controls("Cbocon10").ListIndex = 0

    
    '3. [x] Fill Form Controls with data and Set general purpose CommandButtons
    Let k = (.InsideWidth * 0.95)
    With .Controls("Btnmain0") '("BtnCancel")
      Let .Cancel = True
      Let .Accelerator = "c"
      Let .Top = (.Parent.InsideHeight * 0.97) - (.Height)
      Let .Left = k - .Width + xLeft * 0.5
      Let .Caption = "Cerrar"
    End With
'    With .Controls("Btnmain1") '("BtnHelp")
'      Let .Caption = "Ayuda"
'      Let .Accelerator = "a"
'      Let .ControlTipText = "Ver ayuda del Ejercicio"
'      Let .Top = .Parent.Controls("Btnmain0").Top
'      Let .Left = xLeft * 0.45
'    End With
        
    '4. Assing handler for each button
    Set BtnCancel = .Controls("Btnmain0") 'Main
    'Set Btn0 = .Controls("Btnmain1")      'Main
    Set Btn1 = .Controls("Btncon0") 'Page 0
    Set Btn2 = .Controls("Btncon1") 'Page 0
    Set Btn3 = .Controls("Btncon2") 'Page 0
    Set Btn4 = .Controls("Btncon3") 'Page 0
    Set Btn5 = .Controls("Btncon4") 'Page 0
    Set Btn6 = .Controls("Btncon5") 'Page 0
    Set Btn7 = .Controls("Btncon6") 'Page 2
    Set Btn8 = .Controls("Btncon7") 'Page 2
    Set Btn9 = .Controls("Btncon8") 'Page 2
    Set Btn10 = .Controls("Btncon9") 'Page 2
    Set Btn11 = .Controls("Btncon10") 'Page 2
    Set Btn12 = .Controls("Btncon11") 'Page 3
    Set Btn13 = .Controls("Btncon12") 'Page 3
    Set Btn14 = .Controls("Btncon13") 'Page 3
    Set Btn15 = .Controls("Btncon14") 'Page 3

    Set Chk0 = .Controls("Chkcon0") 'Page 1
    Set Chk1 = .Controls("Chkcon1") 'Page 1
    Set Chk2 = .Controls("Chkcon2") 'Page 2
    Set Chk3 = .Controls("Chkcon3") 'Page 3

    Set Cbo0 = .Controls("Cbocon7") 'Page 2
    Set Cbo1 = .Controls("Cbocon8") 'Page 3
    
    Set Lst0 = .Controls("Lstcon0") 'Page 0
    Set Lst1 = .Controls("Lstcon1") 'Page 3
    
    '5. Give styles to controls
    Call VIEWS.ViewFormStyles(FRM): Set VIEWS = Nothing
    
    'Set initial values for cbo7
    Debug.Print .Controls("Cbocon7").ListCount
    With .Controls("Cbocon7")
      If (.ListCount < 5) Then Let .Style = fmStyleDropDownCombo
      Let FRM.Controls("Btncon8").Enabled = (.ListCount < 5)
      Let FRM.Controls("Btncon9").Enabled = (.ListCount < 5)
      Let FRM.Controls("Btncon10").Enabled = (.ListCount < 5)
    End With
    Let .Controls("Multi").Value = 3
    With .Controls("Cbocon8")
      Let .ListStyle = fmListStyleOption
      Let .ListIndex = -1
      Let .ListIndex = 0
    End With
    Let .Controls("Multi").Value = 0
    

    '6. [x] Show the finished Form
    Call Load(FRM)
    Call .Show

    '7. [x] Close form
    Let .Controls("Lstpur0").RowSource = VBA.vbNullString
    Let .Controls("Lstpur1").RowSource = VBA.vbNullString
    Let .Controls("Cbopur0").RowSource = VBA.vbNullString
    Let .Controls("Cbopur1").RowSource = VBA.vbNullString
    Let .Controls("Cbopur2").RowSource = VBA.vbNullString
  End With

EH:
  Erase aLabels
  Erase aMaxLengths
  Erase aControlis
  Erase aControls
  Set DPR = Nothing
  Set VIEWS = Nothing
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)
End Sub









' CORE VIEW ASSETS ------------------------------------------------------------------'
Private Sub View_assets_controls_free()
  On Error GoTo EH

  Set BtnAccept = Nothing
  Set BtnCancel = Nothing
  Set Btn0 = Nothing
  Set Btn1 = Nothing
  Set Btn2 = Nothing
  Set Btn3 = Nothing
  Set Btn4 = Nothing
  Set Btn5 = Nothing
  Set Btn6 = Nothing
  Set Btn7 = Nothing
  Set Btn8 = Nothing
  Set Btn9 = Nothing
  Set Btn10 = Nothing
  Set Btn11 = Nothing
  Set Btn12 = Nothing
  Set Btn13 = Nothing
  Set Btn14 = Nothing
  Set Btn15 = Nothing
  
  Set Chk0 = Nothing
  Set Chk1 = Nothing
  Set Chk2 = Nothing
  Set Chk3 = Nothing
  Set Chk4 = Nothing
  Set Chk5 = Nothing
  
  Set Cbo0 = Nothing
  Set Cbo1 = Nothing
  Set Cbo2 = Nothing
  Set Cbo3 = Nothing
  
  Set Lst0 = Nothing
  Set Lst1 = Nothing
  
  Set Txt0 = Nothing
  Set Txt1 = Nothing

  Set Spn0 = Nothing
  Set Spn1 = Nothing
  Set Spn2 = Nothing
  Set Spn3 = Nothing
  Set Spn4 = Nothing
  Set Spn5 = Nothing

EH:
  'Call EHGLOBAL.ErrorHandlerRaise("VIEW::-ControlsFree")
  'Call EHGLOBAL.ErrorHandlerRaise("VIEW::-ControlsFree")
End Sub

Private Function View_assets_controls_getformparent(ByVal UControl As MSForms.Control) As FRMBasics
  'Search across each container control as Frames, MultiPages, etc., to find its main UserForm Parent
  Dim CtlParent As Object
  
  On Error GoTo EH
  
  Set CtlParent = UControl.Parent 'Find control's parent UserForm
  Do While Not (VBA.TypeName(CtlParent) = "FRMBasics")
    #If Debugging Then
      Debug.Print VBA.TypeName(CtlParent)
    #End If
    Set CtlParent = CtlParent.Parent
    'Exit Do
  Loop
  'GoTo EH
  
  Set View_assets_controls_getformparent = CtlParent
  
EH:
  Set CtlParent = Nothing
  Call EHGLOBAL.ErrorHandlerDisplay("VIEW::-Assets_GetFormParent")
End Function

'Controls Helper methods
Private Sub View_assets_closeform(ByVal FRM As FRMBasics)
  On Error GoTo EH
  Call VBA.Unload(FRM)
EH:
  Call EHGLOBAL.ErrorHandlerDisplay("VIEW::-assets_closeform")
End Sub

Private Sub View_assets_capture_data_from_ranges( _
  ByVal VFRM As FRMBasics, _
  ByVal VTxt As MSForms.TextBox, _
  ByVal VStrTitle As String, _
  Optional VEnuType As EQLRES_ENU_VALUETYPE = eqlResNumbers, _
  Optional VBooAbs As Boolean = False, _
  Optional VBooZeroes As Boolean = False, _
  Optional VBooRngCols As Boolean = True, _
  Optional VBooSlashGrouped As Boolean, _
  Optional VBooSum As Boolean = False, _
  Optional VBooRngAddressToTxt As Boolean = False)
  
  Dim RES As AppResources_current
  Dim VIEWS As ViewsParent_current
  
  Dim RngBox As Range
  Dim VecAux0 As Variant
  Dim i As Long
  
  On Error GoTo EH
    
  With VFRM
    
    Set VIEWS = New ViewsParent_current
    Set VIEWS.ErrorHandler = EHGLOBAL
    Set RngBox = VIEWS.ViewsCaptureRange(VFRM, VStrTitle, STRAPPTITLE)
    Set VIEWS = Nothing
    
    If Not RngBox Is Nothing Then
      
      If VEnuType = eqlResRanges Then ' Ranges
        Set GLORNG_CELL = Nothing
        Set GLORNG_CELL = RngBox
        If VBooRngAddressToTxt Then Let VTxt.Value = RngBox.Range("A1").Address
      Else
        Set RES = New AppResources_current
        If VBooSlashGrouped Then ' Grouped by /
          Let VTxt.Value = RES.ArrayGetDataSemicolonSeparatedSlashGroupedStringFromVariants(RngBox, i, False, VBooAbs, VBooZeroes, RBooRngRowsCols:=VBooRngCols)
        Else ' ; separation
          If VBooSum Then ' Sum elements
            Let VTxt.Value = VBA.vbNullString
            Let VecAux0 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(RngBox, i, RValueType:=VEnuType, RBooAbs:=VBooAbs, RBooIncludeZeroes:=VBooZeroes)
            If i > 0 Then Let VTxt.Value = Application.WorksheetFunction.Sum(VecAux0)
            Let VecAux0 = Empty
          Else
            Let VTxt.Value = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(RngBox, i, RValueType:=VEnuType, RBooAbs:=VBooAbs, RBooIncludeZeroes:=VBooZeroes)
          End If
        End If
      End If
      Set RngBox = Nothing
      Set RES = Nothing
    
    End If
  
  End With

EH:
  Set RngBox = Nothing
  Set VIEWS = Nothing
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise("VIEW::-assets_capture_data_from_ranges")

End Sub

Private Sub View_assets_capture_numbers_from_ranges( _
  ByVal VFRM As FRMBasics, _
  ByVal VStrBtnName As String, _
  ByVal VStrTxtName As String, _
  ByVal VStrTitle As String, _
  Optional VBooAbs As Boolean = False, _
  Optional VBooZeroes As Boolean = False, _
  Optional VBooRngCols As Boolean = True, _
  Optional VBooSlashGrouped As Boolean, _
  Optional VBooSum As Boolean = False)
    
  Dim RES As AppResources
  Dim VIEWS As ViewsApp
  
  Dim RngBox As Range
  Dim VecAux As Variant
  Dim StrAux1 As String
  Dim i As Long
  
  On Error GoTo EH
    
  With VFRM
    Set VIEWS = New ViewsApp
    Set VIEWS.ErrorHandler = EHGLOBAL
    Set RngBox = VIEWS.ViewsCaptureRange(VFRM, VStrTitle, STRAPPTITLE)
    Set VIEWS = Nothing
    If Not RngBox Is Nothing Then
      Set RES = New AppResources
      Let StrAux1 = VBA.Right(VStrBtnName, 1)
      If VBooSlashGrouped = False Then '; separation
        If VBooSum Then
          Let VecAux = RES.resArrayGetDataArrayFromVariants(RngBox, i, False, VBooAbs, VBooZeroes, RBooRngRowsCols:=VBooRngCols)
          If i = 0 Then
            Let .Controls(VStrTxtName & StrAux1).Value = VBA.vbNullString
          Else
            Let .Controls(VStrTxtName & StrAux1).Value = Application.WorksheetFunction.Sum(VecAux)
          End If
          Let VecAux = Empty
        Else
          Let .Controls(VStrTxtName & StrAux1).Value = RES.resArrayGetDataColonSeparatedStringFromVariants(RngBox, i, False, VBooAbs, VBooZeroes, RBooRngRowsCols:=VBooRngCols)
        End If
      Else 'Grouped by /
        Let .Controls(VStrTxtName & StrAux1).Value = RES.resArrayGetDataSemicolonSeparatedSlashGroupedStringFromVariants(RngBox, i, False, VBooAbs, VBooZeroes, RBooRngRowsCols:=VBooRngCols)
      End If
      Set RES = Nothing
      Set RngBox = Nothing
    End If
  End With

EH:
  Set RngBox = Nothing
  Set VIEWS = Nothing
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise("VIEW::-assets_capture_numbers_from_ranges")
End Sub

Private Sub View_assets_capture_strings_from_ranges( _
  ByVal VFRM As FRMBasics, _
  ByVal VStrBtnName As String, _
  ByVal VStrTxtName As String, _
  ByVal VStrTitle As String, _
  Optional VBooEmptyStrings As Boolean = False, _
  Optional VStrRgx As String = VBA.vbNullString, _
  Optional VBooRngCols As Boolean = True)
  
  Dim RES As AppResources
  Dim VIEWS As ViewsApp
  
  Dim RngBox As Range
  Dim StrAux1 As String
  Dim i As Long
  
  On Error GoTo EH
    
  With VFRM
    Set VIEWS = New ViewsApp
    Set VIEWS.ErrorHandler = EHGLOBAL
    Set RngBox = VIEWS.ViewsCaptureRange(VFRM, VStrTitle, STRAPPTITLE)
    Set VIEWS = Nothing
    If Not RngBox Is Nothing Then
      Set RES = New AppResources
      Let StrAux1 = VBA.Right(VStrBtnName, 1)
      Let .Controls(VStrTxtName & StrAux1).Value = RES.resArrayGetDataColonSeparatedStringFromVariants(RngBox, i, True, RBooIncludeEmptyStrings:=VBooEmptyStrings, RStrCustomRegex:=VStrRgx, RBooRngRowsCols:=VBooRngCols)
      Set RES = Nothing
      Set RngBox = Nothing
    End If
  End With

EH:
  Set RngBox = Nothing
  Set VIEWS = Nothing
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise("VIEW::-assets_capture_strings_from_ranges")
End Sub

Private Sub View_assets_capture_strings_from_ranges_customtextbox( _
  ByVal VFRM As FRMBasics, _
  ByVal VStrBtnName As String, _
  ByVal VStrTxtName As String, _
  ByVal VStrTitle As String, _
  Optional VBooEmptyStrings As Boolean = False, _
  Optional VStrRgx As String = VBA.vbNullString, _
  Optional VBooRngCols As Boolean = True)
  
  Dim RES As AppResources
  Dim VIEWS As ViewsApp
  
  Dim RngBox As Range
  Dim StrAux1 As String
  Dim i As Long
  
  On Error GoTo EH
    
  With VFRM
    Set VIEWS = New ViewsApp
    Set VIEWS.ErrorHandler = EHGLOBAL
    Set RngBox = VIEWS.ViewsCaptureRange(VFRM, VStrTitle, STRAPPTITLE)
    Set VIEWS = Nothing
    If Not RngBox Is Nothing Then
      Set RES = New AppResources
      Let .Controls(VStrTxtName).Value = RES.resArrayGetDataColonSeparatedStringFromVariants(RngBox, i, True, RBooIncludeEmptyStrings:=VBooEmptyStrings, RStrCustomRegex:=VStrRgx, RBooRngRowsCols:=VBooRngCols)
      Set RES = Nothing
      Set RngBox = Nothing
    End If
  End With

EH:
  Set RngBox = Nothing
  Set VIEWS = Nothing
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise("VIEW::-assets_capture_strings_from_ranges_customtextbox")
End Sub

Private Sub View_assets_capture_range_cell( _
  ByVal VFRM As FRMBasics, _
  ByVal VStrBtnName As String, _
  ByVal VStrTxtName As String, _
  Optional VStrTitle As String = STR_TITLE8 & " Será la referencia donde imprimir el resultado:", _
  Optional VBooJustCaptureRange As Boolean = False)
  
  Dim RES As AppResources
  Dim VIEWS As ViewsApp
  
  Dim RngBox As Range
  Dim StrAux1 As String
  Dim i As Long
  
  On Error GoTo EH
    
  With VFRM
    Set VIEWS = New ViewsApp
    Set VIEWS.ErrorHandler = EHGLOBAL
    Set RngBox = VIEWS.ViewsCaptureRange(VFRM, VStrTitle, STRAPPTITLE)
    Set VIEWS = Nothing
    If Not RngBox Is Nothing Then
      Set RES = New AppResources
      Set GLORNG_CELL = RngBox
      Set RES = Nothing
      If Not VBooJustCaptureRange Then
        Let .Controls(VStrTxtName).Value = RngBox.Range("A1").Address
        Let .Controls(VStrTxtName).Enabled = False
      End If
      Set RngBox = Nothing
    End If
  End With

EH:
  Set RngBox = Nothing
  Set VIEWS = Nothing
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise("VIEW::-assets_capture_range_cell")
End Sub

Private Sub View_assets_controls_listbox_mannagement( _
  ByVal VFRM As FRMBasics, _
  ByVal VStrCaller As String)
  
  Dim RES As AppResources
  
  Dim RngBox As Range
  Dim LO As ListObject
  Dim VecAux As Variant
  Dim StrAux0 As String, StrAux1 As String
  Dim BooAux0 As Boolean, BooAux1 As Boolean
  Dim i As Long, j As Long, k As Long, l As Long
  
  On Error GoTo EH
    
  With VFRM
    Select Case VFRM.Tag
      Case "oreinvoice"
        Select Case .Controls("Cbopur1").ListIndex
          Case 3: Let StrAux0 = "excliqliteconfigpurretentions"
          Case 4: Let StrAux0 = "excliqliteconfigpurretentionsother"
          Case Else: GoTo EH
        End Select
        If .Controls("Lstpur0").Tag = "change" Then GoTo EH
        
        Select Case VStrCaller
          Case "Chkpur2"
            Let .Controls("Lstpur0").Enabled = Not .Controls(VStrCaller).Value
            If .Controls("Chkpur1").Value Then 'Multiple
              Select Case .Controls("Cbopur1").ListIndex
                Case 3: Let StrAux0 = "excliqlitepurchasevariousretentions"
                Case 4: Let StrAux0 = "excliqlitepurchasevariousretentionsother"
                Case Else: GoTo EH
              End Select
              Let i = .Controls("Cbopur2").ListIndex + 1
              If i <= 0 Then GoTo EH
              Let excliqlitedatasheet.ListObjects(StrAux0).ListColumns("all").DataBodyRange.Cells(i, 1).Value = IIf(.Controls(VStrCaller).Value, 1, 0)
              If .Controls(VStrCaller).Value = False Then Let excliqlitedatasheet.ListObjects(StrAux0).ListColumns("values").DataBodyRange.Cells(i, 1).Value = VBA.vbNullString
              Let .Controls("Lstpur0").Tag = "change"
              Let j = .Controls("Lstpur0").ListCount - 1
              If j < 0 Then GoTo EH
              For i = 0 To j
                Let .Controls("Lstpur0").Selected(i) = .Controls(VStrCaller).Value
              Next i
              Let .Controls("Lstpur0").Tag = VBA.vbNullString
            Else
              Select Case .Controls("Cbopur1").ListIndex
                Case 3: Let StrAux0 = "excliqliteconfigpurretentions"
                Case 4: Let StrAux0 = "excliqliteconfigpurretentionsother"
                Case Else: GoTo EH
              End Select
              Let .Controls("Lstpur0").Tag = "change"
              Let j = .Controls("Lstpur0").ListCount
              Let .Controls("Lstpur0").RowSource = VBA.vbNullString
              If j < 0 Then GoTo EH
              For i = 1 To j
                Let excliqlitedatasheet.ListObjects(StrAux0).ListColumns("SELECTED").DataBodyRange.Cells(i, 1).Value = IIf(.Controls(VStrCaller).Value, 1, 0)
              Next i
              Let .Controls("Lstpur0").RowSource = excliqlitedatasheet.ListObjects(StrAux0)
              For i = 0 To j - 1
                Let .Controls("Lstpur0").Selected(i) = .Controls(VStrCaller).Value
              Next i
              Let .Controls("Lstpur0").Tag = VBA.vbNullString
            End If
            
          Case "Cbopur2"
            If .Controls("Lstpur0").ListIndex >= 0 Then
              If .Controls("Chkpur1").Value Then 'Multiple
                If Not excliqlitedatasheet.ListObjects("excliqlitepurchasevariousretentions").DataBodyRange Is Nothing Then
                  Select Case .Controls("Cbopur1").ListIndex
                    Case 3: Let StrAux0 = "excliqlitepurchasevariousretentions"
                    Case 4: Let StrAux0 = "excliqlitepurchasevariousretentionsother"
                    Case Else: GoTo EH
                  End Select
                  Let StrAux1 = excliqlitedatasheet.ListObjects(StrAux0).DataBodyRange.Cells(.Controls(VStrCaller).ListIndex + 1, 1)
                  For i = 0 To .Controls("Lstpur0").ListCount - 1
                    Let .Controls("Lstpur0").Selected(i) = False
                  Next i
                  Let .Controls("Chkpur2").Value = excliqlitedatasheet.ListObjects(StrAux0).DataBodyRange.Cells(.Controls(VStrCaller).ListIndex + 1, 3) '2) 'False
                  Let j = .Controls("Lstpur0").ListCount - 1
                  Let k = 0
                  Let VecAux = VBA.Split(StrAux1, ";")
                  If Not StrAux1 = VBA.vbNullString Then
                    For i = 0 To j
                      For k = 0 To UBound(VecAux)
                        If .Controls("Lstpur0").List(i, 0) = VBA.CStr(VecAux(k)) Then
                          Let .Controls("Lstpur0").Selected(i) = True
                        End If
                      Next k
                    Next i
                  End If
                  Let VecAux = Empty
                  Let .Controls("Lstpur0").Enabled = True
                End If
              End If
            End If
          
          Case "Lstpur0"
            If .Controls("Lstpur0").ListIndex >= 0 Then
              If .Controls("Chkpur1").Value Then 'Multiple
                If Not excliqlitedatasheet.ListObjects("excliqlitepurchasevariousretentions").DataBodyRange Is Nothing Then
                  Select Case .Controls("Cbopur1").ListIndex
                    Case 3: Let StrAux0 = "excliqlitepurchasevariousretentions"
                    Case 4: Let StrAux0 = "excliqlitepurchasevariousretentionsother"
                    Case Else: GoTo EH
                  End Select
                  Let j = .Controls("Lstpur0").ListCount - 1
                  Let k = 0
                  Let l = 0
                  ReDim VecAux(0)
                  Let VecAux(0) = VBA.vbNullString
                  For i = 0 To j
                    If .Controls("Lstpur0").Selected(i) Then
                      ReDim Preserve VecAux(l)
                      Let VecAux(l) = .Controls("Lstpur0").List(i, 0)
                      Let l = l + 1
                    End If
                    Let k = k + IIf(.Controls("Lstpur0").Selected(i), 1, 0)
                  Next i
                  If k = (j + 1) Then .Controls("Chkpur2").Value = True
                  If k = 0 Or k <= j Then .Controls("Chkpur2").Value = False
                  Let StrAux1 = VBA.Join(VecAux, ";")
                  Let VecAux = Empty
                  Let i = .Controls("Cbopur2").ListIndex + 1
                  Let excliqlitedatasheet.ListObjects(StrAux0).DataBodyRange.Cells(i, 1).Value = StrAux1
                End If
              Else 'Single
                Let j = .Controls("Lstpur0").ListCount - 1
                Let k = 0
                For i = 0 To j
                  Let excliqlitedatasheet.ListObjects(StrAux0).ListColumns("SELECTED").DataBodyRange.Cells(i + 1, 1).Value = IIf(.Controls("Lstpur0").Selected(i), 1, 0)
                  Let k = k + IIf(.Controls("Lstpur0").Selected(i), 1, 0)
                Next i
                If k = (j + 1) Then .Controls("Chkpur2").Value = True: .Controls("Lstpur0").Enabled = False
                If k = 0 Or k <= j Then .Controls("Chkpur2").Value = False
              End If
            End If
          
          Case Else: GoTo EH
        End Select
      
      Case "concentrateinvoice"
        Select Case VStrCaller
          Case "Lstcnc0"
            If .Controls(VStrCaller).Tag = "noevent" Then GoTo EH
            Select Case .Controls("Cbocnc1").ListIndex
              Case 2 'Penalties
                Let StrAux0 = "excliqliteconpurpenalties"
                Let i = .Controls(VStrCaller).ListIndex
                Let BooAux0 = .Controls(VStrCaller).Selected(i)
                Let .Controls(VStrCaller).RowSource = VBA.vbNullString
                Let excliqlitedatasheet.ListObjects(StrAux0).ListColumns("SELECTED").DataBodyRange.Cells(i + 1, 1).Value = IIf(BooAux0, 1, 0)
                Let .Controls(VStrCaller).RowSource = excliqlitedatasheet.ListObjects(StrAux0)
              
              Case 3 'Deductions
                '0: Exports, 1: Retentions, 2: Retentions others
                Select Case .Controls("Cbocnc2").ListIndex
                  Case 0, 1, 2
                    Let i = .Controls("Cbocnc2").ListIndex
                    Let StrAux0 = IIf(i = 0, "excliqliteconpurdedexp", IIf(i = 1, "excliqliteconfigpurretentions", "excliqliteconfigpurretentionsother")) 'IIf(i = 1, "excliqliteconfigpurretentions", "excliqliteconfigpurretentionsother")
                    Let i = .Controls(VStrCaller).ListIndex
                    Let BooAux0 = .Controls(VStrCaller).Selected(i)
                    Let .Controls(VStrCaller).RowSource = VBA.vbNullString
                    Let excliqlitedatasheet.ListObjects(StrAux0).ListColumns("SELECTED").DataBodyRange.Cells(i + 1, 1).Value = IIf(BooAux0, 1, 0)
                    Let .Controls(VStrCaller).RowSource = excliqlitedatasheet.ListObjects(StrAux0)
                End Select
                  
            End Select
        End Select
      Case Else: GoTo EH
    End Select
  End With

EH:
  Set RngBox = Nothing
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise("VIEW::-assets_controls_listbox_mannagement")
End Sub

Private Sub View_assets_get_selected_or_new_range(ByVal ControlId As String, ByRef VVarReq As Variant)
  
  ' [Create new sheet and assign its first range] or assign a user selected range
  Dim RES As AppResources_current
  
  Dim RngBox As Range
  Dim StrName As String
  Dim BooNewSheet As Boolean
  Dim LngVecReqIndex As Long
  
  Const sFN As String = "VIEW::-assets_get_selected_or_new_range"
  
  On Error GoTo EH
  Select Case ControlId
    
    Case "regression", "regressionselection": Let BooNewSheet = VBA.CBool(VVarReq(15)): Let LngVecReqIndex = 16: Let StrName = "Regresión"
        
'    Case "regressionselection"
'    Case "balmet"
'    Case "balmetselection"
'    Case "balmeteco"
'    Case "projection"
'    Case "oreinvoice"
    
    Case Else: GoTo EH
  End Select
  
  Set RES = New AppResources_current
  Set RES.ErrorHandler = EHGLOBAL
  If Not BooNewSheet Then Set RngBox = VVarReq(LngVecReqIndex)
  Set VVarReq(LngVecReqIndex) = View_assets_getcell(BooNewSheet, RES, RngBox, StrName)

EH:
  Set RngBox = Nothing
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Sub

Private Function View_assets_getcell( _
  ByVal VBooNewSheet As Boolean, _
  ByVal RES As AppResources_current, _
  ByVal VRngBox As Range, _
  ByVal VStrSheetName As String) As Range
  
  Dim ShtNewSheet As Worksheet
  
  Const sFN As String = "VIEW:-assets_getcell"
  
  On Error GoTo EH
  
  If VBooNewSheet Then
    
    Set ShtNewSheet = RES.ExcelCurrentOrNewSheet(VStrSheetName)
    
    #If Debugging Then
      Call Debug_assert(Not ShtNewSheet Is Nothing, sFN)
    #End If
    
    If ShtNewSheet Is Nothing Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "No se pudo crear la nueva hoja: '" & VStrSheetName & "' para imrpimir los resultados.")
    
    Set View_assets_getcell = ShtNewSheet.Range("A1")
    Set ShtNewSheet = Nothing
    
  Else
    
    Set View_assets_getcell = VRngBox.Range("A1")
    
  End If

EH:
  If Not VBA.Err.Number = 0 And Not ShtNewSheet Is Nothing Then Call ShtNewSheet.Parent.Worksheets(ShtNewSheet.Name).Delete
  Set ShtNewSheet = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function





' ERROR ASSERTSIONS -----------------------------------------------------------------'
Private Sub Debug_assert(ByVal Cond As Boolean, Optional Origin As String = VBA.vbNullString)
  
  #If Debbuging Then
    Debug.Assert Cond
    If Not Cond Then Debug.Print "Assert on: " & Origin
  #End If

End Sub

Private Function View_Core_error_handler(ByVal ControlId As String, ByVal FRM As FRMBasics) As Boolean
  
  Const sFN As String = "VIEW::-core_error_handler"
  
  On Error GoTo EH
  Let View_Core_error_handler = False
  If Not CMB Is Nothing Then Set CMB = New ClassCustomMsgbox
  
  Select Case ControlId
    
    Case "regression"
      Select Case EHGLOBAL.ErrorNum
        Case 0, 1: Call CMB.CMsgBox(EHGLOBAL.ErrorMsg, STRAPPTITLE, "Error", cmbError, cmbNoButtons, FRM, "Txtreg1", "BtnMain1", cmbRequired, cmbCenter, cmbMiddle, cmbGreenDefault, cmb3Seconds, cmbPopupAutoclose)
        Case 2: Call CMB.CMsgBox(EHGLOBAL.ErrorMsg, STRAPPTITLE, "Error", cmbError, cmbNoButtons, FRM, "Txtreg0", "BtnMain1", cmbRequired, cAspect:=cmbPopupAutoclose)
        Case Else: Call CMB.CMsgBox(EHGLOBAL.ErrorMsg, STRAPPTITLE, "Error", cmbError, cmbClose, FRM): Let View_Core_error_handler = True
      End Select
      
    Case "balmet"
    
    Case "balmeteco"
    
    Case "projection"
    
    Case "oreinvoice"
    
    Case Else
      Call VBA.MsgBox(EHGLOBAL.ErrorMsg, vbInformation, STRAPPTITLE)
    
  End Select
  
EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)
  
End Function








' METHODS PUBLIC --------------------------------------------------------------------'
' INTERFACE METHODS LIST (PUBLIC) ---------------------------------------------------'
Public Sub ViewUI(ByVal ControlId As String)
  Dim oFrm As New FRMBasics
  
  On Error GoTo EH
  Select Case ControlId
    Case "regression": Call View_form_regression(ControlId, oFrm)
    Case "regressionselection": Call View_form_regressionselection(ControlId)
    
    Case "balmet": Call View_form_balmet(ControlId, oFrm)
    Case "balmeteco": Call View_form_balmeteco(ControlId, oFrm)
    Case "balmetselection": Call View_form_balmetselection(ControlId)
    
    Case "projection": Call View_form_projection(ControlId, oFrm) 'Call ProjectionsTest
    
    Case "oreinvoice": Call View_form_purchasemineral(ControlId, oFrm)
    
    Case "concentrateinvoice": Call View_form_purchaseconcentrate(ControlId, oFrm)
    
    Case "configlauncher": Call View_form_config(ControlId, oFrm)
    
    'Help, About, License
    
    Case Else: GoTo EH
  End Select
EH:
  Set oFrm = Nothing
  Call View_assets_controls_free
  Call EHGLOBAL.ErrorHandlerDisplay("VIEW::UI")
End Sub
' METHODS PUBLIC --------------------------------------------------------------------'


