VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassEqlView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================== CLASS EXCLIQ LITE VIEW STRUCTURE ======================== '
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


' ======================== CLASS EXCLIQ LITE VIEW STRUCTURE ======================== '
' CUSTOM EVENTS -------------------------------------------------------------------- '
' OBJECT VARIABLES (GLOOBJ_) ------------------------------------------------------- '
' VARIABLES (GLOStr_) -------------------------------------------------------------- '
' PROPERTIES ----------------------------------------------------------------------- '
' CONSTRUCTOR/DESTRUCTOR ----------------------------------------------------------- '
' CONTROLS EVENTS ------------------------------------------------------------------ '
' CORE METHODS LIST (PRIVATE) ------------------------------------------------------ '
' CONTROLS ADMINITRATORS ----------------------------------------------------------- '
' View_Controls_buttons_click
' View_Controls_checkboxes_change
' View_Controls_combos_change
' View_Controls_listboxes_change
' View_Controls_listboxes_click
' View_Controls_listboxes_dblclick
' View_Controls_texts_dropbuttonclick
' View_Controls_spinbuttons_change
' View_Core_validations
' CORE VIEWS METHODS LIST ---------------------------------------------------------- '
' View_Core_validations
' View_Core_validations_regression
' View_Core_validations_balmet
' View_Core_validations_projections
' View_Core_validations_oreinvoice
' View_Core_validations_concentrateinvoice
' View_core_validations_config
' View_Core_preparations
' View_Core_preparations_regression
' View_Core_preparations_balmet
' View_Core_preparations_projection
' View_Core_preparations_oreinvoice
' View_Core_preparations_concentrateinvoice
' View_core_preparations_help
' View_Core_execution
' CORE VIEWS FORMS METHODS LIST ---------------------------------------------------- '
' View_form_regression
' View_form_regressionselection
' View_form_balmet
' View_form_balmeteco
' View_form_balmetselection
' View_form_projection
' View_form_purchasemineral
' View_form_purchaseconcentrate
' View_form_config
' View_form_help
' View_form_about
' CORE VIEW ASSETS ----------------------------------------------------------------- '
' View_assets_initialize_model
' View_assets_controls_free
' View_assets_controls_getformparent
' View_assets_closeform
' View_assets_capture_data_from_ranges
' View_assets_get_selected_or_new_range
' View_assets_getcell
' View_assets_read_data_for_balmet
' View_assets_clear_last_colonseparatedstring_item
' View_assets_check_for_spreadsheet_edges
' View_assets_oreinvoice_basedata
' View_assets_oreinvoice_maindata
' View_assets_oreinvoice_bonusdiscounts
' View_assets_oreinvoice_retentions
' View_assets_cncinvoice_contents
' View_assets_cncinvoice_refination
' View_assets_cncinvoice_penalties
' View_assets_cncinvoice_deductions
' View_assets_cncinvoice_bonus
' View_assets_cncinvoice_dedexports
' View_assets_cncinvoice_dedtax
' View_assets_cncinvoice_dedothertax
' View_assets_cncinvoice_deddiscounts
' View_assets_config_rm
' View_assets_config_retentions
' View_assets_config_retentionsother
' View_assets_config_penalties
' View_assets_config_exports
' View_Core_error_handler
' ERRORS MANNAGEMENT AND ASSERTSIONS ----------------------------------------------- '
' Debug_assert
' CORE METHODS LIST (PUBLIC) ------------------------------------------------------- '
' ViewUI
' ======================== CLASS EXCLIQ LITE VIEW STRUCTURE ======================== '



' ======================== CLASS EXCLIQ LITE VIEW STRUCTURE ======================== '
' CUSTOM EVENTS -------------------------------------------------------------------- '
Public Event RequestEvent(ByVal EStrRequester As String, ByRef EVarRequest As Variant, ByRef EVarResponse As Variant, ByRef EBooCancel As Boolean)

' OBJECT VARIABLES (GLOOBJ_) ------------------------------------------------------- '
Private RES As ClassEqlResources
Private MDL As ClassEqlModel
Private DPR As ClassEqlVendorsDatePicker
Private CMB As ClassEqlVendorsCustomMsgBox

' Controls
Private WithEvents BtnAccept As MSForms.CommandButton
Attribute BtnAccept.VB_VarHelpID = -1
Private WithEvents BtnCancel As MSForms.CommandButton
Attribute BtnCancel.VB_VarHelpID = -1
Private WithEvents Btn0 As MSForms.CommandButton ' Buttons
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
Private WithEvents Btn16 As MSForms.CommandButton
Attribute Btn16.VB_VarHelpID = -1
Private WithEvents Btn17 As MSForms.CommandButton
Attribute Btn17.VB_VarHelpID = -1

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
Private WithEvents Chk6 As MSForms.CheckBox
Attribute Chk6.VB_VarHelpID = -1
Private WithEvents Chk7 As MSForms.CheckBox
Attribute Chk7.VB_VarHelpID = -1
Private WithEvents Chk8 As MSForms.CheckBox
Attribute Chk8.VB_VarHelpID = -1
Private WithEvents Chk9 As MSForms.CheckBox
Attribute Chk9.VB_VarHelpID = -1
Private WithEvents Chk10 As MSForms.CheckBox
Attribute Chk10.VB_VarHelpID = -1
Private WithEvents Chk11 As MSForms.CheckBox
Attribute Chk11.VB_VarHelpID = -1
Private WithEvents Chk12 As MSForms.CheckBox
Attribute Chk12.VB_VarHelpID = -1
Private WithEvents Chk13 As MSForms.CheckBox
Attribute Chk13.VB_VarHelpID = -1

Private WithEvents Cbo0 As MSForms.ComboBox
Attribute Cbo0.VB_VarHelpID = -1
Private WithEvents Cbo1 As MSForms.ComboBox
Attribute Cbo1.VB_VarHelpID = -1
Private WithEvents Cbo2 As MSForms.ComboBox
Attribute Cbo2.VB_VarHelpID = -1
Private WithEvents Cbo3 As MSForms.ComboBox
Attribute Cbo3.VB_VarHelpID = -1
Private WithEvents Cbo4 As MSForms.ComboBox
Attribute Cbo4.VB_VarHelpID = -1

Private WithEvents Lst0 As MSForms.ListBox
Attribute Lst0.VB_VarHelpID = -1
Private WithEvents Lst1 As MSForms.ListBox
Attribute Lst1.VB_VarHelpID = -1
Private WithEvents Lst2 As MSForms.ListBox
Attribute Lst2.VB_VarHelpID = -1

Private WithEvents Txt0 As MSForms.TextBox
Attribute Txt0.VB_VarHelpID = -1
Private WithEvents Txt1 As MSForms.TextBox
Attribute Txt1.VB_VarHelpID = -1
Private WithEvents Txt2 As MSForms.TextBox
Attribute Txt2.VB_VarHelpID = -1
Private WithEvents Txt3 As MSForms.TextBox
Attribute Txt3.VB_VarHelpID = -1
Private WithEvents Txt4 As MSForms.TextBox
Attribute Txt4.VB_VarHelpID = -1
Private WithEvents Txt5 As MSForms.TextBox
Attribute Txt5.VB_VarHelpID = -1
Private WithEvents Txt6 As MSForms.TextBox
Attribute Txt6.VB_VarHelpID = -1
Private WithEvents Txt7 As MSForms.TextBox
Attribute Txt7.VB_VarHelpID = -1
Private WithEvents Txt8 As MSForms.TextBox
Attribute Txt8.VB_VarHelpID = -1
Private WithEvents Txt9 As MSForms.TextBox
Attribute Txt9.VB_VarHelpID = -1
Private WithEvents Txt10 As MSForms.TextBox
Attribute Txt10.VB_VarHelpID = -1
Private WithEvents Txt11 As MSForms.TextBox
Attribute Txt11.VB_VarHelpID = -1
Private WithEvents Txt12 As MSForms.TextBox
Attribute Txt12.VB_VarHelpID = -1
Private WithEvents Txt13 As MSForms.TextBox
Attribute Txt13.VB_VarHelpID = -1
Private WithEvents Txt14 As MSForms.TextBox
Attribute Txt14.VB_VarHelpID = -1
Private WithEvents Txt15 As MSForms.TextBox
Attribute Txt15.VB_VarHelpID = -1

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

' VARIABLES (GLOStr_) -------------------------------------------------------------- '
Private GLORNG_CELL As Range

' For data management
Private STRAPPTITLE As String
Private GLOBOO_FORMULAS As Boolean
Private GLOBOO_SELECTIONCELL As Boolean
Private GLOBOO_COMPACT As Boolean
Private GLOBOO_HEADS As Boolean
Private GLOBYT_TABLES As Byte



' PROPERTIES ----------------------------------------------------------------------- '
Public Property Set ResourcesApp(ByRef ObjRES As ClassEqlResources)

  Set RES = ObjRES
  Let GLOBOO_FORMULAS = RES.Formulas
  Let STRAPPTITLE = RES.AppStrTitle

End Property



' CONSTRUCTOR/DESTRUCTOR ----------------------------------------------------------- '
Private Sub Class_Initialize()

  If CMB Is Nothing Then Set CMB = New ClassEqlVendorsCustomMsgBox

End Sub

Private Sub Class_Terminate()

  ' Objects
  Set GLORNG_CELL = Nothing
  If Not DPR Is Nothing Then Set DPR = Nothing
  If Not CMB Is Nothing Then Set CMB = Nothing

  ' Form Controls
  Call View_assets_controls_free

  Set RES = Nothing

End Sub



' CONTROLS EVENTS ------------------------------------------------------------------ '
' Buttons
Private Sub BtnCancel_Click()
  Call View_assets_closeform(BtnCancel.Parent)
End Sub

Private Sub BtnAccept_Click()
  If View_controls_buttons_click(BtnAccept, View_assets_controls_getformparent(BtnAccept)) Then Call View_assets_closeform(BtnCancel.Parent)
End Sub

Private Sub Btn0_Click()
  If View_controls_buttons_click(Btn0, View_assets_controls_getformparent(Btn0)) Then Call View_assets_closeform(BtnCancel.Parent)
End Sub

Private Sub Btn1_Click()
  If View_controls_buttons_click(Btn1, View_assets_controls_getformparent(Btn1)) Then Call View_assets_closeform(BtnCancel.Parent)
End Sub

Private Sub Btn2_Click()
  If View_controls_buttons_click(Btn2, View_assets_controls_getformparent(Btn2)) Then Call View_assets_closeform(BtnCancel.Parent)
End Sub

Private Sub Btn3_Click()
  If View_controls_buttons_click(Btn3, View_assets_controls_getformparent(Btn3)) Then Call View_assets_closeform(BtnCancel.Parent)
End Sub

Private Sub Btn4_Click()
  If View_controls_buttons_click(Btn4, View_assets_controls_getformparent(Btn4)) Then Call View_assets_closeform(BtnCancel.Parent)
End Sub

Private Sub Btn5_Click()
  If View_controls_buttons_click(Btn5, View_assets_controls_getformparent(Btn5)) Then Call View_assets_closeform(BtnCancel.Parent)
End Sub

Private Sub Btn6_Click()
  If View_controls_buttons_click(Btn6, View_assets_controls_getformparent(Btn6)) Then Call View_assets_closeform(BtnCancel.Parent)
End Sub

Private Sub Btn7_Click()
  If View_controls_buttons_click(Btn7, View_assets_controls_getformparent(Btn7)) Then Call View_assets_closeform(BtnCancel.Parent)
End Sub

Private Sub Btn8_Click()
  If View_controls_buttons_click(Btn8, View_assets_controls_getformparent(Btn8)) Then Call View_assets_closeform(BtnCancel.Parent)
End Sub

Private Sub Btn9_Click()
  If View_controls_buttons_click(Btn9, View_assets_controls_getformparent(Btn9)) Then Call View_assets_closeform(BtnCancel.Parent)
End Sub

Private Sub Btn10_Click()
  If View_controls_buttons_click(Btn10, View_assets_controls_getformparent(Btn10)) Then Call View_assets_closeform(BtnCancel.Parent)
End Sub

Private Sub Btn11_Click()
  If View_controls_buttons_click(Btn11, View_assets_controls_getformparent(Btn11)) Then Call View_assets_closeform(BtnCancel.Parent)
End Sub

Private Sub Btn12_Click()
  If View_controls_buttons_click(Btn12, View_assets_controls_getformparent(Btn12)) Then Call View_assets_closeform(BtnCancel.Parent)
End Sub

Private Sub Btn13_Click()
  If View_controls_buttons_click(Btn13, View_assets_controls_getformparent(Btn13)) Then Call View_assets_closeform(BtnCancel.Parent)
End Sub

Private Sub Btn14_Click()
  If View_controls_buttons_click(Btn14, View_assets_controls_getformparent(Btn14)) Then Call View_assets_closeform(BtnCancel.Parent)
End Sub

Private Sub Btn15_Click()
  If View_controls_buttons_click(Btn15, View_assets_controls_getformparent(Btn15)) Then Call View_assets_closeform(BtnCancel.Parent)
End Sub

Private Sub Btn16_Click()
  If View_controls_buttons_click(Btn16, View_assets_controls_getformparent(Btn16)) Then Call View_assets_closeform(BtnCancel.Parent)
End Sub

' Checks
Private Sub Chk0_Change()
  Call View_controls_checkboxes_change(Chk0, View_assets_controls_getformparent(Chk0))
End Sub

Private Sub Chk1_Change()
  Call View_controls_checkboxes_change(Chk1, View_assets_controls_getformparent(Chk1))
End Sub

Private Sub Chk2_Change()
  Call View_controls_checkboxes_change(Chk2, View_assets_controls_getformparent(Chk2))
End Sub

Private Sub Chk3_Change()
  Call View_controls_checkboxes_change(Chk3, View_assets_controls_getformparent(Chk3))
End Sub

Private Sub Chk4_Change()
  Call View_controls_checkboxes_change(Chk4, View_assets_controls_getformparent(Chk4))
End Sub

Private Sub Chk5_Change()
  Call View_controls_checkboxes_change(Chk5, View_assets_controls_getformparent(Chk5))
End Sub

Private Sub Chk6_Change()
  Call View_controls_checkboxes_change(Chk6, View_assets_controls_getformparent(Chk6))
End Sub

Private Sub Chk7_Change()
  Call View_controls_checkboxes_change(Chk7, View_assets_controls_getformparent(Chk7))
End Sub

Private Sub Chk8_Change()
  Call View_controls_checkboxes_change(Chk8, View_assets_controls_getformparent(Chk8))
End Sub

Private Sub Chk9_Change()
  Call View_controls_checkboxes_change(Chk9, View_assets_controls_getformparent(Chk9))
End Sub

Private Sub Chk10_Change()
  Call View_controls_checkboxes_change(Chk10, View_assets_controls_getformparent(Chk10))
End Sub

Private Sub Chk11_Change()
  Call View_controls_checkboxes_change(Chk11, View_assets_controls_getformparent(Chk11))
End Sub

Private Sub Chk12_Change()
  Call View_controls_checkboxes_change(Chk12, View_assets_controls_getformparent(Chk12))
End Sub

Private Sub Chk13_Change()
  Call View_controls_checkboxes_change(Chk13, View_assets_controls_getformparent(Chk13))
End Sub

' Combos
Private Sub Cbo0_Change()
  Call View_controls_combos_change(Cbo0, View_assets_controls_getformparent(Cbo0))
End Sub

Private Sub Cbo1_Change()
  Call View_controls_combos_change(Cbo1, View_assets_controls_getformparent(Cbo1))
End Sub

Private Sub Cbo2_Change()
  Call View_controls_combos_change(Cbo2, View_assets_controls_getformparent(Cbo2))
End Sub

Private Sub Cbo3_Change()
  Call View_controls_combos_change(Cbo3, View_assets_controls_getformparent(Cbo3))
End Sub

Private Sub Cbo4_Change()
  Call View_controls_combos_change(Cbo4, View_assets_controls_getformparent(Cbo4))
End Sub

' Lists
Private Sub Lst0_Change()
  Call View_Controls_listboxes_change(Lst0, View_assets_controls_getformparent(Lst0))
End Sub

Private Sub Lst0_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  Call View_controls_listboxes_dblclick(Lst0, View_assets_controls_getformparent(Lst0))
End Sub

Private Sub Lst0_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  If KeyAscii = vbKeySpace Then Call View_controls_listboxes_dblclick(Lst0, View_assets_controls_getformparent(Lst0))
End Sub

Private Sub Lst1_Change()
  Call View_Controls_listboxes_change(Lst1, View_assets_controls_getformparent(Lst1))
End Sub

Private Sub Lst1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  Call View_controls_listboxes_dblclick(Lst1, View_assets_controls_getformparent(Lst1))
End Sub

Private Sub Lst2_Change()
  Call View_Controls_listboxes_change(Lst2, View_assets_controls_getformparent(Lst2))
End Sub

Private Sub Lst2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  Call View_controls_listboxes_dblclick(Lst2, View_assets_controls_getformparent(Lst2))
End Sub

' Spins
Private Sub Spn0_SpinUp()
  Call View_controls_spinbuttons_change(Spn0, View_assets_controls_getformparent(Spn0), True)
End Sub

Private Sub Spn0_SpinDown()
  Call View_controls_spinbuttons_change(Spn0, View_assets_controls_getformparent(Spn0))
End Sub

Private Sub Spn1_SpinUp()
  Call View_controls_spinbuttons_change(Spn1, View_assets_controls_getformparent(Spn1), True)
End Sub

Private Sub Spn1_SpinDown()
  Call View_controls_spinbuttons_change(Spn1, View_assets_controls_getformparent(Spn1))
End Sub

Private Sub Spn2_SpinUp()
  Call View_controls_spinbuttons_change(Spn2, View_assets_controls_getformparent(Spn2), True)
End Sub

Private Sub Spn2_SpinDown()
  Call View_controls_spinbuttons_change(Spn2, View_assets_controls_getformparent(Spn2))
End Sub

Private Sub Spn3_SpinUp()
  Call View_controls_spinbuttons_change(Spn3, View_assets_controls_getformparent(Spn3), True)
End Sub

Private Sub Spn3_SpinDown()
  Call View_controls_spinbuttons_change(Spn3, View_assets_controls_getformparent(Spn3))
End Sub

Private Sub Spn4_SpinUp()
  Call View_controls_spinbuttons_change(Spn4, View_assets_controls_getformparent(Spn4), True)
End Sub

Private Sub Spn4_SpinDown()
  Call View_controls_spinbuttons_change(Spn4, View_assets_controls_getformparent(Spn4))
End Sub

Private Sub Spn5_SpinUp()
  Call View_controls_spinbuttons_change(Spn5, View_assets_controls_getformparent(Spn5), True)
End Sub

Private Sub Spn5_SpinDown()
  Call View_controls_spinbuttons_change(Spn5, View_assets_controls_getformparent(Spn5))
End Sub

' Texts
Private Sub Txt0_DropButtonClick()
  Call View_controls_texts_dropbuttonclick(Txt0, View_assets_controls_getformparent(Txt0))
End Sub

Private Sub Txt1_DropButtonClick()
  Call View_controls_texts_dropbuttonclick(Txt1, View_assets_controls_getformparent(Txt1))
End Sub

Private Sub Txt2_DropButtonClick()
  Call View_controls_texts_dropbuttonclick(Txt2, View_assets_controls_getformparent(Txt2))
End Sub

Private Sub Txt3_DropButtonClick()
  Call View_controls_texts_dropbuttonclick(Txt3, View_assets_controls_getformparent(Txt3))
End Sub

Private Sub Txt4_DropButtonClick()
  Call View_controls_texts_dropbuttonclick(Txt4, View_assets_controls_getformparent(Txt4))
End Sub

Private Sub Txt5_DropButtonClick()
  Call View_controls_texts_dropbuttonclick(Txt5, View_assets_controls_getformparent(Txt5))
End Sub

Private Sub Txt6_DropButtonClick()
  Call View_controls_texts_dropbuttonclick(Txt6, View_assets_controls_getformparent(Txt6))
End Sub

Private Sub Txt7_DropButtonClick()
  Call View_controls_texts_dropbuttonclick(Txt7, View_assets_controls_getformparent(Txt7))
End Sub

Private Sub Txt8_DropButtonClick()
  Call View_controls_texts_dropbuttonclick(Txt8, View_assets_controls_getformparent(Txt8))
End Sub

Private Sub Txt9_DropButtonClick()
  Call View_controls_texts_dropbuttonclick(Txt9, View_assets_controls_getformparent(Txt9))
End Sub

Private Sub Txt10_DropButtonClick()
  Call View_controls_texts_dropbuttonclick(Txt10, View_assets_controls_getformparent(Txt10))
End Sub

Private Sub Txt11_DropButtonClick()
  Call View_controls_texts_dropbuttonclick(Txt11, View_assets_controls_getformparent(Txt11))
End Sub

Private Sub Txt12_DropButtonClick()
  Call View_controls_texts_dropbuttonclick(Txt12, View_assets_controls_getformparent(Txt12))
End Sub

Private Sub Txt13_DropButtonClick()
  Call View_controls_texts_dropbuttonclick(Txt13, View_assets_controls_getformparent(Txt13))
End Sub

Private Sub Txt14_DropButtonClick()
  Call View_controls_texts_dropbuttonclick(Txt14, View_assets_controls_getformparent(Txt14))
End Sub

Private Sub Txt15_DropButtonClick()
  Call View_controls_texts_dropbuttonclick(Txt15, View_assets_controls_getformparent(Txt15))
End Sub



' CORE METHODS LIST (PRIVATE) ------------------------------------------------------ '
' CONTROLS ADMINITRATORS ----------------------------------------------------------- '
' Events handlers - Buttons
Private Function View_controls_buttons_click( _
  ByVal VBtnCurrent As MSForms.CommandButton, _
  ByVal FRM As FRMBasics, _
  Optional VBooSelections As Boolean = False, _
  Optional ControlId As String) As Boolean

  Dim VWP As ClassEqlViewParent
  Dim RngBox As Range
  Dim VecAux0 As Variant, VecAux1 As Variant
  Dim StrAux0 As String, StrAux1 As String
  Dim BooAux0 As Boolean, BooAux1 As Boolean
  Dim i As Long, j As Long, k As Long, u As Long, v As Long, w As Long
  Dim DblAux0 As Double, DblAux1 As Double, DblAux2 As Double, DblAux3 As Double, DblAux4 As Double, DblAux5 As Double, DblAux6 As Double

  Const sFN As String = "VIEW::-buttons_click"

  On Error GoTo EH

  ' Validate first
  If Not CMB Is Nothing Then
    If CMB.IsActive Then GoTo EH
  End If
  If Not View_core_validations(VBtnCurrent, FRM, VBooSelections, ControlId) Then GoTo EH

  ' Execute task
  If VBooSelections Then ' =========================================================== Validation for Selections contents'

    Set VWP = New ClassEqlViewParent
    Select Case ControlId

      Case "regressionselection" ' Go to get a Cellrange
        ' 1.  Evaluates first validation with direct selection.
        ' 2.  In validations, avoid evaluation of RngCell in first instance 'cause it isn't captured yet and its value is nothing
        ' 3.  Capture the RngCell and make a new call to validations, this time, send RngCell, if evaluates, call exe
        ' Select RngCell
        If Not GLOBOO_SELECTIONCELL Then Let View_controls_buttons_click = View_Core_execution(ControlId, Nothing, VStatusMsg:="RL Selección"): GoTo EH ' SELECTIONCELLS IN RIBBON IS INACTIVE, SO, NEED TO CREATE A NEW SHEET
        Set RngBox = VWP.ViewsCaptureRange(RES, FRM, "Seleccione una celda donde generar la Regresión Lineal:", STRAPPTITLE)
        ' Validate
        If View_core_validations(Nothing, Nothing, True, ControlId & "-check_cell", RngBox) Then Let View_controls_buttons_click = View_Core_execution(ControlId, Nothing, RngBox, VStatusMsg:="RL Selección")

      Case "balmetselection" ' Go to get a Cellrange
        ' Select RngCell
        If Not GLOBOO_SELECTIONCELL Then Let View_controls_buttons_click = View_Core_execution(ControlId, Nothing, VStatusMsg:="Balmet Selección"): GoTo EH ' SELECTIONCELLS IN RIBBON IS INACTIVE, SO, NEED TO CREATE A NEW SHEET
        Set RngBox = VWP.ViewsCaptureRange(RES, FRM, "Seleccione una celda donde generar el Balance Metalúrgico:", STRAPPTITLE)
        ' Validate
        If View_core_validations(Nothing, Nothing, True, ControlId & "-check_cell", RngBox) Then Let View_controls_buttons_click = View_Core_execution(ControlId, Nothing, RngBox, VStatusMsg:="Balmet Selección")

      Case Else: GoTo EH

    End Select
    Set VWP = Nothing

  Else ' ============================================================================= Validations to Forms contents

    #If Debugging Then
      Call Debug_assert((Not VBtnCurrent Is Nothing Or FRM Is Nothing), sFN)
    #End If
    If VBtnCurrent Is Nothing Or FRM Is Nothing Then GoTo EH

    Select Case FRM.Tag

      Case "regression" ' ----------------------------------------------- regression'
        If VBtnCurrent.Name = "Btnmain1" Then Let View_controls_buttons_click = View_Core_execution(FRM.Tag, FRM, VStatusMsg:=FRM.Caption)
        If VBtnCurrent.Name = "Btnmain2" Then Let View_controls_buttons_click = View_Core_execution("hlp" & FRM.Tag, Nothing, VStatusMsg:="Ayuda: " & FRM.Caption, VBooHelp:=True)

      Case "balmet" ' --------------------------------------------------- balmet'
        With FRM
          Select Case VBtnCurrent.Name
            Case "Btnmain1": Let View_controls_buttons_click = View_Core_execution(FRM.Tag, FRM, VStatusMsg:=FRM.Caption) ' Execute balmet task
            Case "Btnmain2": Let View_controls_buttons_click = View_Core_execution("hlp" & FRM.Tag, Nothing, VStatusMsg:="Ayuda: " & FRM.Caption, VBooHelp:=True)
            Case "Btnbal0": Let .Controls("Txtbal3").Text = View_assets_clear_last_colonseparatedstring_item(.Controls("Txtbal3").Text): Let .Controls("Cbobal1").ListIndex = 0
            Case "Btnbal1": Let .Controls("Txtbal3").Text = VBA.vbNullString: Let .Controls("Cbobal1").ListIndex = 0
            Case "Btnbal2": Let .Controls("Txtbal4").Text = View_assets_clear_last_colonseparatedstring_item(.Controls("Txtbal4").Text): Let .Controls("Cbobal2").ListIndex = 0
            Case "Btnbal3": Let .Controls("Txtbal4").Text = VBA.vbNullString: Let .Controls("Cbobal2").ListIndex = 0
            Case "Btnbal4": Let .Controls("Txtbal6").Text = View_assets_clear_last_colonseparatedstring_item(.Controls("Txtbal6").Text): Let .Controls("Cbobal3").ListIndex = 0
            Case "Btnbal5": Let .Controls("Txtbal6").Text = VBA.vbNullString: Let .Controls("Cbobal3").ListIndex = 0
            Case "Btnbal6": Let .Controls("Txtbal7").Text = View_assets_clear_last_colonseparatedstring_item(.Controls("Txtbal7").Text): Let .Controls("Cbobal4").ListIndex = 0
            Case "Btnbal7": Let .Controls("Txtbal7").Text = VBA.vbNullString: Let .Controls("Cbobal4").ListIndex = 0
            Case Else: GoTo EH
          End Select
        End With

      Case "balmeteco" ' ------------------------------------------------ balmeteco'
        With FRM
          Select Case VBtnCurrent.Name
            Case "Btnmain1": Let View_controls_buttons_click = View_Core_execution(FRM.Tag, FRM, VStatusMsg:=FRM.Caption) ' Execute balmeteco task
            Case "Btnmain2": Let View_controls_buttons_click = View_Core_execution("hlp" & FRM.Tag, Nothing, VStatusMsg:="Ayuda: " & FRM.Caption, VBooHelp:=True)
            Case "Btnbal0": Let .Controls("Txtbal5").Text = View_assets_clear_last_colonseparatedstring_item(.Controls("Txtbal5").Text): Let .Controls("Cbobal2").ListIndex = 0
            Case "Btnbal2": Let .Controls("Txtbal5").Text = VBA.vbNullString: Let .Controls("Cbobal2").ListIndex = 0
            Case "Btnbal1": Let .Controls("Txtbal6").Text = View_assets_clear_last_colonseparatedstring_item(.Controls("Txtbal6").Text): Let .Controls("Cbobal3").ListIndex = 0
            Case "Btnbal3": Let .Controls("Txtbal6").Text = VBA.vbNullString: Let .Controls("Cbobal3").ListIndex = 0
            Case "Btnbal4": Let .Controls("Txtbal8").Text = View_assets_clear_last_colonseparatedstring_item(.Controls("Txtbal8").Text): .Controls("Cbobal4").ListIndex = 0
            Case "Btnbal5": Let .Controls("Txtbal8").Text = VBA.vbNullString: .Controls("Cbobal4").ListIndex = 0
            Case "Btnbal6": Let .Controls("Txtbal9").Text = View_assets_clear_last_colonseparatedstring_item(.Controls("Txtbal9").Text): .Controls("Cbobal5").ListIndex = 0
            Case "Btnbal7": Let .Controls("Txtbal9").Text = VBA.vbNullString: .Controls("Cbobal5").ListIndex = 0
            Case Else: GoTo EH
          End Select
        End With

      Case "projection" ' ----------------------------------------------- projection'
        With FRM
          Select Case VBtnCurrent.Name
            Case "Btnpro0": Let .Controls("Txtpro5").Text = View_assets_clear_last_colonseparatedstring_item(.Controls("Txtpro5").Text): Let .Controls("Cbopro1").ListIndex = 0 ' 1;2;3
            Case "Btnpro1": Let .Controls("Txtpro5").Text = VBA.vbNullString: Let .Controls("Cbopro1").ListIndex = 0
            Case "Btnpro2": Let .Controls("Txtpro6").Text = View_assets_clear_last_colonseparatedstring_item(.Controls("Txtpro6").Text): Let .Controls("Cbopro2").ListIndex = 0 ' 1;2;3/3;2;1
            Case "Btnpro3": Let .Controls("Txtpro6").Text = VBA.vbNullString: Let .Controls("Cbopro2").ListIndex = 0
            Case "Btnmain1": Let View_controls_buttons_click = View_Core_execution(FRM.Tag, FRM)
            Case "Btnmain2": Let View_controls_buttons_click = View_Core_execution("hlp" & FRM.Tag, Nothing, VStatusMsg:="Ayuda: " & FRM.Caption, VBooHelp:=True)
          End Select
        End With

      Case "oreinvoice" ' ----------------------------------------------- oreinvoice'
        With FRM
          Select Case VBtnCurrent.Name

            Case "Btnpur0", "Btnpur1" ' <+++ Set & Upd
              Let j = .Controls("Lstpur0").ListIndex + 1
              Let .Controls("Lstpur0").RowSource = Empty
              Let .Controls("Lstpur0").Tag = "noevent"
              Let .Controls("Chkpur2").Tag = "noevent"
              Let .Controls("Cbopur2").Tag = "FillList"
              If VBtnCurrent.Name = "Btnpur0" Then ' Add

                Select Case .Controls("Cbopur1").ListIndex

                  Case 0 ' Base Data (Create new row)
                    Let BooAux0 = RES.REGEX.isEmptyStringReg(.Controls("Txtpur5").Value)
                    Let VecAux0 = Array(IIf(BooAux0, "M-" & Format(Now(), "yymmddssms"), .Controls("Txtpur5").Value), VBA.CDate(.Controls("Txtpur6").Value), VBA.CDbl(.Controls("Txtpur7").Text), VBA.CDbl(.Controls("Txtpur8").Text), VBA.CDbl(.Controls("Txtpur9").Text), _
                      Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, 0, _
                      Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty)
                    Call MDL.MSet(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur, VecAux0)

                  Case 1 ' Main Data (update second part of the table)
                    Let VecAux0 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur, eqlMdlArray, "PARTE2"), eqlRes1D)
                    Let i = RES.ArrayGetValuePosition(VecAux0, "0") + 1
                    Let VecAux0 = Empty
                    Let VecAux0 = Array(.Controls("Cbopur2").Value, .Controls(IIf(.Controls("Cbopur2").ListIndex = 1, "Txtpur6", "Cbopur3")).Value, .Controls("Txtpur7").Value, .Controls(IIf(.Controls("Cbopur2").ListIndex = 1, "Txtpur8", "Cbopur4")).Value, .Controls("Txtpur9").Value, 1)
                    Call MDL.MUpd(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur, VecAux0, "TIPO, CONTENIDOS, LEYES, UNIDADES, PRECIO, PARTE2", i)

                  Case 2, 5 ' Bonus, Discounts
                    Let BooAux0 = .Controls("Chkpur1").Value
                    Let BooAux1 = (.Controls("Cbopur1").ListIndex = 2)
                    ' Determine value of Sample column
                    Let StrAux0 = IIf(BooAux0, .Controls("Cbopur2").Value, "N/D") ' Sample
                    Let StrAux1 = IIf(RES.REGEX.isEmptyStringReg(.Controls("Txtpur5").Value), IIf(BooAux1, "Bono-", "Descuento-") & Format(Now(), "yymmddssms"), .Controls("Txtpur5").Value) ' Bonus or Discount Name
                    ' Set data to table
                    Call MDL.MSet(eqlMdlSheetTmp, IIf(BooAux1, eqlMdlTblPurbonus_tmp, eqlMdlTblPurdiscounts_tmp), Array(StrAux0, StrAux1, .Controls("Txtpur6").Text))
                    ' Prepare data to Main table
                    Let StrAux0 = VBA.Join(RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetTmp, IIf(BooAux1, eqlMdlTblPurbonus_tmp, eqlMdlTblPurdiscounts_tmp), eqlMdlArray, IIf(BooAux1, "BONO", "DESCUENTO")), eqlRes1D), VBA.vbTab) ' Concept
                    Let StrAux0 = VBA.Replace(VBA.Replace(VBA.Replace(StrAux0, ";", "[eqlsemicolon]"), "|", "[eqlpipe]"), VBA.vbTab, ";")
                    Let StrAux1 = VBA.Join(RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetTmp, IIf(BooAux1, eqlMdlTblPurbonus_tmp, eqlMdlTblPurdiscounts_tmp), eqlMdlArray, "VALOR"), eqlRes1D), ";") ' Value
                    ' Update data in Main table
                    Call MDL.MUpd(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur, Array(StrAux0 & "|" & StrAux1), IIf(BooAux1, "BONOS", "DESCUENTOS"), IIf(BooAux0, .Controls("Cbopur2").ListIndex + 1, 1))

                End Select
                Let VecAux0 = Empty

              Else ' Update

                Select Case .Controls("Cbopur1").ListIndex
                  Case 0 ' Base Data
                    Let BooAux0 = RES.REGEX.isEmptyStringReg(.Controls("Txtpur5").Value)
                    Let VecAux0 = Array(IIf(BooAux0, "M-" & Format(Now(), "yymmddssms"), .Controls("Txtpur5").Text), VBA.CDate(.Controls("Txtpur6").Text), .Controls("Txtpur7").Text, .Controls("Txtpur8").Text, .Controls("Txtpur9").Text)
                    Call MDL.MUpd(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur, VecAux0, "MUESTRA, FECHA INGRESO, HUMEDAD, MERMA, PESO HÚMEDO BRUTO", j)
                  Case 1 ' Main Data (Cbopur2.ListIndex = 1: Complex)
                    Let VecAux0 = Array(.Controls("Cbopur2").Value, .Controls(IIf(.Controls("Cbopur2").ListIndex = 1, "Txtpur6", "Cbopur3")).Value, .Controls("Txtpur7").Text, .Controls(IIf(.Controls("Cbopur2").ListIndex = 1, "Txtpur8", "Cbopur4")).Value, .Controls("Txtpur9").Text)
                    Call MDL.MUpd(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur, VecAux0, "TIPO, CONTENIDOS, LEYES, UNIDADES, PRECIO", j)
                  Case 2, 5 ' Bonus, Discounts
                    Let BooAux1 = (.Controls("Cbopur1").ListIndex = 2)
                    ' Determine value of Sample column
                    Let StrAux1 = IIf(RES.REGEX.isEmptyStringReg(.Controls("Txtpur5").Value), IIf(BooAux1, "Bono-", "Descuento-") & Format(Now(), "yymmddssms"), .Controls("Txtpur5").Value) ' Bonus or Discount Name
                    ' Set data to table
                    Let VecAux0 = Array(StrAux1, .Controls("Txtpur6").Text)
                    Call MDL.MUpd(eqlMdlSheetTmp, IIf(BooAux1, eqlMdlTblPurbonus_tmp, eqlMdlTblPurdiscounts_tmp), VecAux0, IIf(BooAux1, "BONO", "DESCUENTO") & ", VALOR", j)
                    ' Prepare data to Main table
                    Let StrAux0 = VBA.Join(RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetTmp, IIf(BooAux1, eqlMdlTblPurbonus_tmp, eqlMdlTblPurdiscounts_tmp), eqlMdlArray, IIf(BooAux1, "BONO", "DESCUENTO")), eqlRes1D), VBA.vbTab) ' Concept
                    Let StrAux0 = VBA.Replace(VBA.Replace(VBA.Replace(StrAux0, ";", "[eqlsemicolon]"), "|", "[eqlpipe]"), VBA.vbTab, ";")
                    Let StrAux1 = VBA.Join(RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetTmp, IIf(BooAux1, eqlMdlTblPurbonus_tmp, eqlMdlTblPurdiscounts_tmp), eqlMdlArray, "VALOR"), eqlRes1D), ";") ' Value
                    ' Update data in Main table
                    Call MDL.MUpd(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur, Array(StrAux0 & "|" & StrAux1), IIf(BooAux1, "BONOS", "DESCUENTOS"), IIf(BooAux0, .Controls("Cbopur2").ListIndex + 1, 1))
                End Select
                Let VecAux0 = Empty

              End If
              Call View_controls_combos_change(.Controls("Cbopur1"), FRM)
              Let .Controls("Cbopur2").Tag = VBA.vbNullString
              Let .Controls("Lstpur0").Tag = VBA.vbNullString
              Let .Controls("Chkpur1").Enabled = True
              Let .Controls("Chkpur2").Tag = VBA.vbNullString
              Let .Controls("Cbopur1").Enabled = True
              Let .Controls("Lstpur0").Enabled = True
              Let .Controls("Btnmain1").Enabled = True

            Case "Btnpur2" ' <+++ Cancel Edition
              Let .Controls("Cbopur2").Tag = "FillList"
              Let .Controls("Lstpur0").Tag = "noevent"
              Let .Controls("Chkpur2").Tag = "noevent"
              Call View_controls_combos_change(.Controls("Cbopur1"), FRM)
              Let .Controls("Cbopur2").Tag = VBA.vbNullString
              Let .Controls("Lstpur0").Tag = VBA.vbNullString
              Let .Controls("Chkpur2").Tag = VBA.vbNullString
              Let .Controls("Chkpur1").Enabled = True
              Let .Controls("Cbopur1").Enabled = True
              Let .Controls("Lstpur0").Enabled = True
              Let .Controls("Btnmain1").Enabled = True

            Case "Btnpur3", "Btnpur4" ' <+++ Del once & Del all
              Let j = .Controls("Lstpur0").ListIndex + 1
              Let .Controls("Lstpur0").RowSource = Empty
              Let .Controls("Cbopur2").Tag = "FillList"
              Let .Controls("Lstpur0").Tag = "noevent"
              Let .Controls("Chkpur2").Tag = "noevent"
              Select Case .Controls("Cbopur1").ListIndex
                Case 0, 1 ' Base Data & Main Data
                  Call MDL.MDel(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur, IIf(VBtnCurrent.Name = "Btnpur3", j, 0))
                  If VBtnCurrent.Name = "Btnpur4" Then ' Del all
                    Call MDL.MDel(eqlMdlSheetTmp, eqlMdlTblPurbonus_tmp)
                    Call MDL.MDel(eqlMdlSheetTmp, eqlMdlTblPurdiscounts_tmp)
                    Call MDL.MUpd(eqlMdlSheetTmp, eqlMdlTblConfigpurretentions_tmp, Array(0), "SELETED")
                    Call MDL.MUpd(eqlMdlSheetTmp, eqlMdlTblConfigpurretentionsother_tmp, Array(0), "SELETED")
                  End If
                Case 2, 5 ' Bonus, Discounts
                  Let BooAux0 = (.Controls("Cbopur1").ListIndex = 2)
                  Call MDL.MDel(eqlMdlSheetTmp, IIf(BooAux0, eqlMdlTblPurbonus_tmp, eqlMdlTblPurdiscounts_tmp), IIf(VBtnCurrent.Name = "Btnpur3", j, 0))
                  Let i = IIf(.Controls("Chkpur1").Value, .Controls("Cbopur2").ListIndex + 1, 1)
                  Let k = MDL.LengthRows(eqlMdlSheetTmp, IIf(BooAux0, eqlMdlTblPurbonus_tmp, eqlMdlTblPurdiscounts_tmp))
                  ' Prepare data to Main table
                  Let StrAux0 = VBA.Join(RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetTmp, IIf(BooAux0, eqlMdlTblPurbonus_tmp, eqlMdlTblPurdiscounts_tmp), eqlMdlArray, IIf(BooAux0, "BONO", "DESCUENTO")), eqlRes1D), VBA.vbTab) ' Concept
                  Let StrAux0 = VBA.Replace(VBA.Replace(VBA.Replace(StrAux0, ";", "[eqlsemicolon]"), "|", "[eqlpipe]"), VBA.vbTab, ";")
                  Let StrAux1 = VBA.Join(RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetTmp, IIf(BooAux0, eqlMdlTblPurbonus_tmp, eqlMdlTblPurdiscounts_tmp), eqlMdlArray, "VALOR"), eqlRes1D), ";") ' Value
                  ' Update data in Main table
                  Call MDL.MUpd(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur, Array(IIf(k > 0, StrAux0 & "|" & StrAux1, Empty)), IIf(BooAux0, "BONOS", "DESCUENTOS"), IIf(VBtnCurrent.Name = "Btnpur3", j, 1))
              End Select
              Call View_controls_combos_change(.Controls("Cbopur1"), FRM)
              Let .Controls("Cbopur2").Tag = VBA.vbNullString
              Let .Controls("Lstpur0").Tag = VBA.vbNullString
              Let .Controls("Chkpur2").Tag = VBA.vbNullString

            Case "Btnpur5" ' <+++ Rangedata
              Call View_assets_capture_data_from_ranges(FRM, Nothing, "Seleccionar " & .Controls("Cbopur1").Value, eqlResJustRanges)
              Let VBtnCurrent.Tag = "validate"
              If View_core_validations(VBtnCurrent, FRM, VBooSelections, ControlId) Then ' Correct information

                ' Stop controls events
                Let .Controls("Cbopur2").Tag = "FillList"
                Let .Controls("Lstpur0").Tag = "noevent"
                Let .Controls("Chkpur2").Tag = "noevent"
                Let .Controls("Lstpur0").RowSource = Empty

                ' #####
                Let k = .Controls("Cbopur1").ListIndex
                Let i = MDL.LengthRows(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur)
                Let j = GLORNG_CELL.Rows.Count

                Select Case k

                  Case 0, 1 ' Base Data And/Or Main data: Texts 5,6,7,8,9 (5 or 10 cols)
                    ' 0: Get 5 cols in array, if there are 10 cols, slice it in two arrays
                    ' 1: Get 5 cols in array, if there are 10 cols, slice it in two arrays and check if Main data have at least 1 row
                    Let VecAux0 = RES.ArrayToBaseZero(GLORNG_CELL.Value)
                    If (i + j) > GLOSNG_PUR_LIMIT And k = 0 Then Let VecAux0 = RES.ArraySlice(VecAux0, (GLOSNG_PUR_LIMIT - i) - 1, RBooGetLeftSide:=True)
                    If Not VBA.IsArray(VecAux0) And k = 0 Then Call RES.FrmCMsgbox("La base de datos ya alcanzó el límite de " & GLOSNG_PUR_LIMIT & " registros.", STRAPPTITLE, vbInformation, True, .Controls("Btnpur5"), vbCritical): GoTo EH
                    If GLORNG_CELL.Columns.Count >= 10 Then
                      Let VecAux0 = RES.ArraySlice(VecAux0, 9, RByCols:=True, RBooGetLeftSide:=True)
                      For u = LBound(VecAux0) To UBound(VecAux0)
                        If RES.REGEX.isEmptyStringReg(VBA.CStr(VecAux0(u, 0))) Then Let VecAux0(u, 0) = "M-" & Format(Now(), "yymmddssms")
                      Next u
                    ElseIf GLORNG_CELL.Columns.Count >= 5 Then
                      If k = 0 Then
                        Let VecAux0 = RES.ArraySlice(VecAux0, 4, RByCols:=True, RBooGetLeftSide:=True)
                        For u = LBound(VecAux0) To UBound(VecAux0)
                          If RES.REGEX.isEmptyStringReg(VBA.CStr(VecAux0(u, 0))) Then Let VecAux0(u, 0) = "M-" & Format(Now(), "yymmddssms")
                        Next u
                        ReDim VecAux1(0 To UBound(VecAux0), 0 To 4)
                        Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1, RByCols:=True)
                        Let VecAux1 = Empty
                      End If
                      If k = 1 Then Let VecAux0 = RES.ArraySlice(VecAux0, 4, RByCols:=True, RBooGetLeftSide:=True)
                    End If
                    ' Validate each array into its corresponding position (0 or 1)
                    If k = 0 Then ' Base data (0)

                      ReDim VecAux1(0 To UBound(VecAux0), 0 To 19)
                      For u = 0 To UBound(VecAux0)
                        Let VecAux1(u, 4) = IIf(GLORNG_CELL.Columns.Count >= 10, 1, 0) ' PARTE2
                      Next u
                      Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1, RByCols:=True)
                      Let VecAux1 = Empty
                      Call MDL.MSet(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur, VecAux0)

                    Else ' Main data (1)

                      ' Check if table has 0's in PARTE2 column, if it has, update second part, else cancel process
                      Let VecAux1 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur, eqlMdlArray, "PARTE2"), eqlRes1D)
                      Let k = RES.ArrayGetValuePosition(VecAux1, 0) + 1
                      Let VecAux1 = Empty
                      ' Slice all excedent rows updated
                      Let k = (GLOSNG_PUR_LIMIT - k) - 1
                      If k < UBound(VecAux0) Then Let VecAux0 = RES.ArraySlice(VecAux0, k, RBooGetLeftSide:=True) ' Slice array 'cause have more data than expected
                      ' Update
                      Call MDL.MUpd(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur, Array(VecAux0), "TIPO", k)

                    End If

                  Case 2, 5 ' Bonus & Discounts
                    Let BooAux0 = .Controls("Chkpur1").Value
                    Let i = MDL.LengthRows(eqlMdlSheetTmp, IIf(k = 2, eqlMdlTblPurbonus_tmp, eqlMdlTblPurdiscounts_tmp))
                    Let j = GLORNG_CELL.Rows.Count
                    Let u = CSNG_BONUS_DISCOUNT_LIMIT - i
                    Let VecAux0 = GLORNG_CELL.Value
                    If GLORNG_CELL.Columns.Count > 2 Then Let VecAux0 = RES.ArraySlice(VecAux0, 2, RByCols:=True, RBooGetLeftSide:=True)
                    If j > u Then Let VecAux0 = RES.ArraySlice(VecAux0, (u - 1), RBooGetLeftSide:=True)

                    ' Upd data in table
                    Call MDL.MSet(IIf(k = 2, eqlMdlTblPurbonus_tmp, eqlMdlTblPurdiscounts_tmp), IIf(k = 2, eqlMdlTblPurbonus_tmp, eqlMdlTblConfigpurretentions_tmp), Array(VecAux0))
                    ' Prepare data to Main table
                    Let StrAux0 = VBA.Join(RES.ArrayToBaseZero(MDL.MGet(IIf(k = 2, eqlMdlTblPurbonus_tmp, eqlMdlTblPurdiscounts_tmp), eqlMdlArray, IIf(k = 2, "BONO", "DESCUENTO")), eqlRes1D), VBA.vbTab) ' Concept
                    Let StrAux0 = VBA.Replace(VBA.Replace(VBA.Replace(StrAux0, ";", "[eqlsemicolon]"), "|", "[eqlpipe]"), VBA.vbTab, ";")
                    Let StrAux1 = VBA.Join(RES.ArrayToBaseZero(MDL.MGet(IIf(k = 2, eqlMdlTblPurbonus_tmp, eqlMdlTblPurdiscounts_tmp), eqlMdlArray, "VALOR"), eqlRes1D), ";") ' Value
                    ' Update data in Main table
                    Call MDL.MUpd(eqlMdlSheetPur, Array(StrAux0 & "|" & StrAux1), IIf(k = 2, "BONOS", "DESCUENTOS"), IIf(BooAux0, .Controls("Cbopur2").ListIndex + 1, 1))

                  Case Else: GoTo EH

                End Select
                ' #####

                ' Restore controls
                Call View_controls_combos_change(.Controls("Cbopur1"), FRM)
                Let .Controls("Cbopur2").Tag = VBA.vbNullString
                Let .Controls("Lstpur0").Tag = VBA.vbNullString
                Let .Controls("Chkpur2").Tag = VBA.vbNullString

              End If
              Let VBtnCurrent.Tag = Empty
              Set GLORNG_CELL = Nothing

            Case "Btnmain1" ' Exe
              Call VBA.MsgBox("Let's prepare " & .Tag & "!", Title:=STRAPPTITLE) ': GoTo EH ' <====!
              Let View_controls_buttons_click = View_Core_execution(FRM.Tag, FRM)

            Case "Btnmain2": Let View_controls_buttons_click = View_Core_execution("hlp" & FRM.Tag, Nothing, VStatusMsg:="Ayuda: " & FRM.Caption, VBooHelp:=True)

            Case Else: GoTo EH
          End Select
        End With

      Case "concentrateinvoice" ' --------------------------------------- concentrateinvoice'
        With FRM
          Select Case VBtnCurrent.Name

            Case "Btncnc0", "Btncnc1" ' Set
              Let .Controls("Lstcnc0").RowSource = Empty
              Select Case .Controls("Cbocnc1").ListIndex

                Case 0 ' Contents
                  Let VecAux0 = Array(.Controls("Cbocnc2").Value, VBA.Abs(VBA.CDbl(.Controls("Txtcnc11").Text)), VBA.CDbl(.Controls("Cbocnc2").Text), .Controls("Cbocnc3").Value, VBA.Abs(VBA.CDbl(.Controls("Txtcnc14").Text)), .Controls("Cbocnc2").List(.Controls("Cbocnc2").ListIndex, 2), "Refinación " & .Controls("Cbocnc2").Value, 0, 0)
                  Call MDL.MSet(eqlMdlSheetPur, eqlMdlTblpurchasecnccontents_pur, VecAux0)

                Case 3, 4 ' Deductions, bonus
                  If (.Controls("Cbocnc1").ListIndex = 3 And .Controls("Cbocnc2").ListIndex = 3) Or .Controls("Cbocnc1").ListIndex = 4 Then ' Discounts or bonus
                    Let VecAux0 = Array(Empty, .Controls("Txtcnc11").Value, VBA.Abs(VBA.CDbl(.Controls("Txtcnc12").Value)))
                    Call MDL.MSet(eqlMdlSheetTmp, IIf(.Controls("Cbocnc1").ListIndex = 3, eqlMdlTblPurdiscounts_tmp, eqlMdlTblPurbonus_tmp), VecAux0)
                  End If

                Case Else: GoTo EH
              End Select
              Let VecAux0 = Empty
              Call View_controls_combos_change(.Controls("Cbocnc1"), FRM)
              If .Controls("Cbocnc1").ListIndex = 3 Then .Controls("Cbocnc2").ListIndex = 3 'Call View_Controls_combos_change(.Controls("Cbocnc2"), FRM)

            Case "Btncnc1" ' Upd
              Let i = .Controls("Lstcnc0").ListIndex
              Let .Controls("Lstcnc0").RowSource = Empty
              Select Case .Controls("Cbocnc1").ListIndex

                Case 0 ' Contents
                  Let VecAux0 = Array(.Controls("Cbocnc2").Value, VBA.Abs(VBA.CDbl(.Controls("Txtcnc11").Text)), VBA.CDbl(.Controls("Cbocnc2").Text), .Controls("Cbocnc3").Value, VBA.Abs(VBA.CDbl(.Controls("Txtcnc14").Text)), .Controls("Cbocnc2").List(.Controls("Cbocnc2").ListIndex, 2), "Refinación " & .Controls("Cbocnc2").Value, .Controls("Lstcnc0").List(i, 7), .Controls("Lstcnc0").List(i, 8))
                  Call MDL.MUpd(eqlMdlSheetPur, eqlMdlTblpurchasecnccontents_pur, VecAux0, MLngIndex:=i + 1)

                Case 1 'Refinement
                  Let VecAux0 = RES.ArrayToBaseZero(RES.ArrayDelIndex(.Controls("Lstcnc0").List, i, True), eqlRes1D)
                  Let VecAux0(7) = VBA.Abs(VBA.CDbl(.Controls("Txtcnc11").Text))
                  Let VecAux0(8) = VBA.Abs(VBA.CDbl(.Controls("Txtcnc12").Text))
                  Call MDL.MUpd(eqlMdlSheetPur, eqlMdlTblpurchasecnccontents_pur, VecAux0, MLngIndex:=i + 1)

                Case 2 ' Penalties
                  Let VecAux0 = Array(.Controls("Chkcnc1").Caption, VBA.Abs(VBA.CDbl(.Controls("Txtcnc11").Text)), VBA.Abs(VBA.CDbl(.Controls("Txtcnc12").Text)), VBA.Abs(VBA.CDbl(.Controls("Txtcnc13").Text)), VBA.Abs(VBA.CDbl(.Controls("Txtcnc14").Text)), .Controls("Lstcnc0").List(i, 5))
                  Call MDL.MUpd(eqlMdlSheetPur, eqlMdlTblpurchasecncpenalties_pur, VecAux0, MLngIndex:=i + 1)

                Case 3, 4 ' Deductions, bonus
                  If (.Controls("Cbocnc1").ListIndex = 3 And .Controls("Cbocnc2").ListIndex = 0) Then ' Exports
                    Let VecAux0 = RES.ArrayToBaseZero(RES.ArrayDelIndex(.Controls("Lstcnc0").List, i, True), eqlRes1D)
                    Let VecAux0(1) = VBA.Abs(VBA.CDbl(.Controls("Txtcnc12").Text))
                    Let VecAux0(2) = VBA.Abs(VBA.CDbl(.Controls("Txtcnc13").Text))
                    Let VecAux0(3) = IIf(.Controls("Chkcnc1").Value, 1, 0)
                    Call MDL.MUpd(eqlMdlSheetPur, eqlMdlTblpurchasecncdedexp_pur, VecAux0, MLngIndex:=i + 1)

                  ElseIf (.Controls("Cbocnc1").ListIndex = 3 And .Controls("Cbocnc2").ListIndex = 3) Or .Controls("Cbocnc1").ListIndex = 4 Then ' Discounts or bonus
                    Let VecAux0 = Array(Empty, .Controls("Txtcnc11").Value, VBA.Abs(VBA.CDbl(.Controls("Txtcnc12").Value)))
                    Call MDL.MUpd(eqlMdlSheetTmp, IIf(.Controls("Cbocnc1").ListIndex = 3, eqlMdlTblPurdiscounts_tmp, eqlMdlTblPurbonus_tmp), VecAux0, MLngIndex:=i + 1)

                  End If
                  Let VecAux0 = Empty
                  Call View_controls_combos_change(.Controls("Cbocnc1"), FRM)
                  If .Controls("Cbocnc1").ListIndex = 3 Then Call View_controls_combos_change(.Controls("Cbocnc2"), FRM)

              End Select

            Case "Btncnc2" ' Cancel Edition
              Call View_controls_combos_change(.Controls("Cbocnc1"), FRM)
              If .Controls("Cbocnc1").ListIndex = 3 Then Call View_controls_combos_change(.Controls("Cbocnc2"), FRM)
              Let .Controls("Cbocnc1").Enabled = True
              Let .Controls("Btnmain1").Enabled = True

            Case "Btncnc3", "Btncnc4" ' Del once & all
              Let i = .Controls("Lstcnc0").ListIndex + 1
              Let .Controls("Lstcnc0").RowSource = Empty
              Select Case .Controls("Cbocnc1").ListIndex
                Case 0 ' Contents
                  Call MDL.MDel(eqlMdlSheetPur, eqlMdlTblpurchasecnccontents_pur, i)
                Case 1 ' Refinement DelA
                  Call MDL.MUpd(eqlMdlSheetPur, eqlMdlTblpurchasecnccontents_pur, Array(0, 0), "BASE, CONDICIÓN [%]", i)
                Case 2 ' Penalties Del>x
                  If VBtnCurrent.Name = "Btncnc3" Then
                    If i > 6 Then Call MDL.MDel(eqlMdlSheetPur, eqlMdlTblpurchasecncpenalties_pur, i)
                  Else
                    For j = 7 To .Controls("Lstcnc0").ListCount
                      Call MDL.MDel(eqlMdlSheetPur, eqlMdlTblpurchasecncpenalties_pur, j)
                    Next j
                  End If
                Case 3 ' Deductions - Discounts
                  If .Controls("Cbocnc2").ListIndex = 3 Then Call MDL.MDel(eqlMdlSheetTmp, eqlMdlTblPurdiscounts_tmp, i)
                Case 4 ' Bonus
                  Call MDL.MDel(eqlMdlSheetTmp, eqlMdlTblPurbonus_tmp, i)
              End Select
              Call View_controls_combos_change(.Controls("Cbocnc1"), FRM)
              If .Controls("Cbocnc1").ListIndex = 3 Then Call View_controls_combos_change(.Controls("Cbocnc2"), FRM)

            Case "Btncnc5" ' Rangedata
              Call View_assets_capture_data_from_ranges(FRM, Nothing, "Seleccionar " & .Controls("Cbocnc1").Value, eqlResJustRanges)
              Let VBtnCurrent.Tag = "validate"
              If View_core_validations(VBtnCurrent, FRM, VBooSelections, ControlId) Then ' Correct information


                Let .Controls("Lstcnc0").RowSource = Empty
                Select Case .Controls("Cbocnc1").ListIndex

                  Case 0 ' Contents
                    Let VecAux0 = RES.ArrayToBaseZero(GLORNG_CELL.Value)
                    ReDim VecAux1(0 To UBound(VecAux0), 0 To 3)
                    Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1, RByCols:=True)
                    Let VecAux1 = Empty
                    ' Name contents - eqlMdlSheetSys, eqlMdlTblConfigpurrm_sys, eqlMdlArray, "MINERAL, ELEMENTO, UNIDAD MINERAL"
                    Let VecAux1 = RES.ArrayToBaseZero(RES.ArrayDelIndex(.Controls("Cbocnc1").List, 1, True, True), eqlRes1D)
                    For i = LBound(VecAux0) To UBound(VecAux0)
                      Let j = RES.ArrayHasValueN(VecAux1, VBA.LCase(VBA.CStr(VecAux0(i, 0))), ALCase:=True)
                      If j >= 0 Then Let VecAux0(i, 0) = .Controls("Cbocnc1").List(0, j)
                    Next i
                    Let VecAux1 = Empty
                    Let VecAux1 = RES.ArrayToBaseZero(RES.ArrayShift(.Controls("Cbocnc1").List, True, True), eqlRes1D)
                    For i = LBound(VecAux0) To UBound(VecAux0)
                      Let j = RES.ArrayGetValuePosition(VecAux1, VBA.LCase(VBA.CStr(VecAux0(i, 0))))
                      If j >= 0 Then
                        Let VecAux0(i, 1) = VBA.CStr(VBA.Abs(VBA.CDbl(VecAux0(i, 1)))) & "[" & VBA.CStr(.Controls("Cbocnc1").List(j, 2)) & "]"
                        Let VecAux0(i, 4) = VBA.Abs(VBA.CDbl(VecAux0(i, 1)))
                        Let VecAux0(i, 5) = VBA.CStr(.Controls("Cbocnc1").List(j, 2))
                        Let VecAux0(i, 6) = "Refinación" & VBA.CStr(.Controls("Cbocnc1").List(j, 0))
                        Let VecAux0(i, 7) = 0
                        Let VecAux0(i, 8) = 0
                      End If
                    Next i
                    Call MDL.MSet(eqlMdlSheetPur, eqlMdlTblpurchasecnccontents_pur, VecAux0)

                  Case 1 ' Refinement
                    Let VecAux0 = RES.ArrayToBaseZero(GLORNG_CELL.Value)
                    For i = LBound(VecAux0) To UBound(VecAux0)
                      Let VecAux0(i, 0) = VBA.Abs(VBA.CDbl(VecAux0(i, 0)))
                      Let VecAux0(i, 1) = VBA.Abs(VBA.CDbl(VecAux0(i, 1)))
                    Next i
                    Call MDL.MUpd(eqlMdlSheetPur, eqlMdlTblpurchasecnccontents_pur, VecAux0, "BASE")

                  Case 3, 4 ' Discounts, bonus
                    If (.Controls("Cbocnc1").ListIndex = 3 And .Controls("Cbocnc2").ListIndex = 3) Or .Controls("Cbocnc1").ListIndex = 4 Then ' Discounts or bonus
                      Let VecAux0 = RES.ArrayToBaseZero(GLORNG_CELL.Value)
                      For i = LBound(VecAux0) To UBound(VecAux0)
                        Let VecAux0(i, 1) = VBA.Abs(VBA.CDbl(VecAux0(i, 1)))
                      Next i
                      ReDim VecAux1(0 To UBound(VecAux0), 0 To 0)
                      Let VecAux0 = RES.ArrayConcat(VecAux1, VecAux0, RByCols:=True)
                      Let VecAux1 = Empty
                      Call MDL.MSet(eqlMdlSheetTmp, IIf(.Controls("Cbocnc1").ListIndex = 3, eqlMdlTblPurdiscounts_tmp, eqlMdlTblPurbonus_tmp), VecAux0)
                    End If

                  Case Else: GoTo EH
                End Select

              End If
              Let VecAux0 = Empty
              Let VecAux1 = Empty
              Call View_controls_combos_change(.Controls("Cbocnc1"), FRM)
              If .Controls("Cbocnc1").ListIndex = 3 Then Call View_controls_combos_change(.Controls("Cbocnc2"), FRM)
              Set GLORNG_CELL = Nothing
              Let VBtnCurrent.Tag = Empty

            Case "Btnmain1" ' Exe
              Let DblAux0 = IIf(RES.REGEX.isNumberReg(.Controls("Txtcnc2").Text), VBA.Abs(VBA.CDbl(.Controls("Txtcnc2").Text)), 1): Let DblAux0 = IIf(DblAux0 = 0, 1, DblAux0)
              Let DblAux1 = IIf(RES.REGEX.isNumberReg(.Controls("Txtcnc5").Text), VBA.Abs(VBA.CDbl(.Controls("Txtcnc5").Text)), 1): Let DblAux1 = IIf(DblAux1 = 0, 1, DblAux1)
              Let DblAux2 = IIf(RES.REGEX.isNumberReg(.Controls("Txtcnc7").Text), VBA.Abs(VBA.CDbl(.Controls("Txtcnc7").Text)), 1): Let DblAux2 = IIf(DblAux2 = 0, 1, DblAux2)
              Let DblAux3 = IIf(RES.REGEX.isNumberReg(.Controls("Txtcnc9").Text), VBA.Abs(VBA.CDbl(.Controls("Txtcnc9").Text)), 1): Let DblAux3 = IIf(DblAux3 = 0, 1, DblAux3)
              Let DblAux4 = IIf(RES.REGEX.isNumberReg(.Controls("Txtcnc6").Text), VBA.Abs(VBA.CDbl(.Controls("Txtcnc6").Text)), 1): Let DblAux4 = IIf(DblAux4 = 0, 1, DblAux4)
              Let DblAux5 = IIf(RES.REGEX.isNumberReg(.Controls("Txtcnc8").Text), VBA.Abs(VBA.CDbl(.Controls("Txtcnc8").Text)), 1): Let DblAux5 = IIf(DblAux5 = 0, 1, DblAux5)
              Let DblAux6 = IIf(RES.REGEX.isNumberReg(.Controls("Txtcnc10").Text), VBA.Abs(VBA.CDbl(.Controls("Txtcnc10").Text)), 1): Let DblAux6 = IIf(DblAux6 = 0, 1, DblAux6)
              Let VecAux0 = Array("SIMULACIÓN", "SIMULACIÓN", VBA.CDate(.Controls("Txtcnc1").Text), .Controls("Txtcnc0").Text, DblAux0, .Controls("Chkcnc0").Value, .Controls("Txtcnc3").Value, "N/D", .Controls("Cbocnc0").Value, "N/D", "N/D", "N/D", "N/D", "N/D", "N/D", VBA.CDate(.Controls("Txtcnc4").Value), DblAux1, DblAux2, DblAux3, DblAux4, DblAux5, DblAux6)
              If MDL.LengthRows(eqlMdlSheetPur, eqlMdlTblpurchasecnc_pur) > 0 Then ' Upd
                Call MDL.MUpd(eqlMdlSheetPur, eqlMdlTblpurchasecnc_pur, VecAux0, MLngIndex:=1)
              Else ' Set
                Call MDL.MSet(eqlMdlSheetPur, eqlMdlTblpurchasecnc_pur, VecAux0)
              End If
              Let VecAux0 = Empty
              Let View_controls_buttons_click = View_Core_execution(FRM.Tag, FRM)

            Case "Btnmain2": Let View_controls_buttons_click = View_Core_execution("hlp" & FRM.Tag, Nothing, VStatusMsg:="Ayuda: " & FRM.Caption, VBooHelp:=True)

            Case Else: GoTo EH
          End Select
        End With


      Case "configlauncher" ' ------------------------------------------- configlauncher'
        With FRM
          Select Case VBtnCurrent.Name

            ' DIVISA: set, upd, ucn, del, Local, Foreign
            ' Page 0 ------------------------------------------------------------------------------------------------------ x
            Case "Btncon0" ' set
              Let .Controls("Lstcon0").RowSource = Empty
              Let VecAux0 = Array(VBA.StrConv(.Controls("Txtcon0").Text, vbProperCase), VBA.StrConv(.Controls("Txtcon1").Text, vbProperCase), VBA.UCase(.Controls("Txtcon2").Text), 0, 0)
              'Call MDL.MSet(eqlMdlSheetSys, eqlMdlTblConfigcurrencies_sys, VecAux0)
              Call VBA.MsgBox("Registro añadido satisfactoriamente.", vbInformation, STRAPPTITLE)
              Let .Controls("Lstcon0").RowSource = MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigcurrencies_sys, eqlMdlStrTableName)

            Case "Btncon1" ' upd
              If .Controls("Lstcon0").ListCount > 0 And .Controls("Lstcon0").ListIndex > 1 Then
                Let j = .Controls("Lstcon0").ListIndex
                Let VecAux0 = RES.ArrayToBaseZero(RES.ArrayDelIndex(.Controls("Lstcon0").List, j, True), eqlRes1D)
                Let .Controls("Lstcon0").RowSource = Empty
                For i = 0 To 2
                  If Not VBA.LCase(.Controls("Txtcon" & i).Text) = VBA.LCase(VBA.CStr(VecAux0(i))) Then
                    Let StrAux0 = IIf(i < 2, VBA.StrConv(.Controls("Txtcon" & i).Text, vbProperCase), VBA.UCase(.Controls("Txtcon2").Text))
                    Let StrAux1 = IIf(i = 0, "currency", IIf(i = 1, "currencies", "symbol"))
                    'Call MDL.MUpd(eqlMdlSheetSys, eqlMdlTblConfigcurrencies_sys, Array(StrAux0), StrAux1, j + 1)
                    Call VBA.MsgBox("Registro modificado satisfactoriamente.", vbInformation, STRAPPTITLE)
                  End If
                Next i
                Let .Controls("Lstcon0").RowSource = MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigcurrencies_sys, eqlMdlStrTableName)
                Call View_controls_buttons_click(.Controls("Btncon2"), FRM)
              End If

            Case "Btncon2" ' upd cancel
              If .Controls("Lstcon0").ListCount > 0 Then
                Let VecAux0 = Array("Btncon0", "Btncon0", "Btncon0", "Btncon0")
                For i = 0 To 3
                  Let .Controls("Btncon" & i).Enabled = IIf(i = 0, True, IIf(i < 3, False, (.Controls("Lstcon0").ListCount > 2)))
                  Let FRM.Controls("Multicon").Pages(i + 1).Enabled = True
                  If i < 3 Then Let FRM.Controls("Txtcon" & i).Text = Empty
                Next i
                Let .Controls("Lstcon0").Enabled = True
                Let VecAux0 = Empty
              End If

            Case "Btncon3" ' del
              If .Controls("Lstcon0").ListCount > 0 And .Controls("Lstcon0").ListIndex > 1 Then
                Let j = .Controls("Lstcon0").ListIndex + 1
                Let .Controls("Lstcon0").RowSource = Empty
                'Call MDL.MDel(eqlMdlSheetSys, eqlMdlTblConfigcurrencies_sys, j)
                Call VBA.MsgBox("Registro modificado satisfactoriamente.", vbInformation, STRAPPTITLE)
                Let .Controls("Lstcon0").RowSource = MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigcurrencies_sys, eqlMdlStrTableName)
              End If

            Case "Btncon4", "Btncon5" ' Local, Foreign
              Let v = IIf(VBtnCurrent.Name = "Btncon4", 4, 5)
              Let w = IIf(VBtnCurrent.Name = "Btncon4", 0, 1)
              Let i = .Controls("Cbocon" & w).ListIndex
              If .Controls("Cbocon" & w).ListCount > 0 And i > 0 Then
                Let VecAux0 = .Controls("Lstcon0").List
                Let .Controls("Lstcon0").RowSource = Empty
                Let j = RES.ArrayGetValuePosition(VecAux0, 1, 3) ' Local cur
                Let k = RES.ArrayGetValuePosition(VecAux0, 1, 4) ' Foreign cur
                For u = LBound(VecAux0) To UBound(VecAux0) ' Local
                  Let VecAux0(u, 3) = 0
                  Let VecAux0(u, 4) = 0
                Next u
                Let u = IIf(i = j, k, IIf(w = 0, i, j)): Let VecAux0(u, 3) = 1 ' Local
                Let i = IIf(i = k, j, IIf(w = 1, i, k)): Let VecAux0(i, 4) = 1 ' Foreign
                'Call MDL.MUpd(eqlMdlSheetSys, eqlMdlTblConfigcurrencies_sys, Array(RES.ArrayDelIndex(VecAux0, IIf(v = 4, 3, 4), True, True)), IIf(v = 4, "currency_main", "currency_foreign"))
                Let VecAux0 = Empty
                Let .Controls("Lstcon0").RowSource = MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigcurrencies_sys, eqlMdlStrTableName)
                Let .Controls("Cbocon0").ListIndex = u
                Let .Controls("Cbocon1").ListIndex = i
                Call VBA.MsgBox("Divisa " & IIf(v = 4, "local", "extranjera") & " establecida satisfactoriamente a: '" & .Controls("Cbocon" & w).Value & "'.", vbInformation, STRAPPTITLE)
              End If

            ' PÁGINA: upd-1, set, upd, del
            ' Page 2 ------------------------------------------------------------------------------------------------------ x
            Case "Btncon6"
              Let BooAux0 = False
              If Not (VBA.CStr(.Controls("Cbocon2").ListIndex) = VBA.CStr(.Controls("Cbocon2").Tag)) Then ' Header
                'Let MDL.CompanyHeader = .Controls("Cbocon2").ListIndex
                Let BooAux0 = True
              End If
              Let i = IIf(VBA.IsNull(.Controls("Chkcon2").Value), 2, IIf(.Controls("Chkcon2").Value, 1, 0)) ' Img
              Let StrAux0 = .Controls("Txtcon3").Text & "[" & i & "]"
              If Not .Controls("Txtcon3").Tag = StrAux0 Then
                'Let MDL.CompanyLogo = StrAux0
                Let BooAux0 = True
              End If
              Let i = IIf(VBA.IsNull(.Controls("Chkcon3").Value), 2, IIf(.Controls("Chkcon3").Value, 1, 0)) ' Name
              Let StrAux0 = .Controls("Txtcon4").Text & "[" & i & "]"
              If Not .Controls("Txtcon4").Tag = StrAux0 Then
                'Let MDL.CompanyName = StrAux0
                Let BooAux0 = True
              End If
              Let i = IIf(VBA.IsNull(.Controls("Chkcon4").Value), 2, IIf(.Controls("Chkcon4").Value, 1, 0)) ' Address
              Let StrAux0 = .Controls("Txtcon5").Text & "[" & i & "]"
              If Not .Controls("Txtcon5").Tag = StrAux0 Then
                'Let MDL.CompanyName = StrAux0
                Let BooAux0 = True
              End If
              Let i = IIf(VBA.IsNull(.Controls("Chkcon5").Value), 2, IIf(.Controls("Chkcon5").Value, 1, 0)) ' Phone
              Let StrAux0 = .Controls("Txtcon6").Text & "[" & i & "]"
              If Not .Controls("Txtcon6").Tag = StrAux0 Then
                'Let MDL.CompanyName = StrAux0
                Let BooAux0 = True
              End If
              Let i = IIf(VBA.IsNull(.Controls("Chkcon6").Value), 2, IIf(.Controls("Chkcon6").Value, 1, 0)) ' Email
              Let StrAux0 = .Controls("Txtcon7").Text & "[" & i & "]"
              If Not .Controls("Txtcon7").Tag = StrAux0 Then
                'Let MDL.CompanyName = StrAux0
                Let BooAux0 = True
              End If
              If BooAux0 Then Call VBA.MsgBox("Datos guardados satisfactoriamente.", vbInformation, STRAPPTITLE)

            Case "Btncon7", "Btncon8" ' Destiny: set & upd
              Let BooAux0 = (VBtnCurrent.Name = "Btncon7")
              With .Controls("Cbocon3")
                Let VecAux0 = RES.ArrayToBaseZero(.List, eqlRes1D)
                If BooAux0 Then ' Set
                  Let VecAux0 = RES.ArrayAddAtLast(VecAux0, .Value)
                Else ' Upd
                  Let i = VBA.CLng(.Tag)
                  Let VecAux0(i) = .Value
                End If
                Let .Tag = "noevent"
                Call .Clear
                Let .List = VecAux0
                Let StrAux0 = VBA.Join(VecAux0, ";")
                Let VecAux0 = Empty
                'Let MDL.PurDestination = StrAux0
                Let .Value = Empty
                Let .Tag = Empty
                Let BooAux0 = (.ListCount > 4)
              End With
              Let .Controls("Btncon8").Enabled = BooAux0
              Let .Controls("Btncon9").Enabled = BooAux0
              Call VBA.MsgBox("Dato " & IIf(BooAux0, "insertado", "modificado") & " correctamente.", vbInformation, STRAPPTITLE)

            Case "Btncon9" ' Destiny: Del
              With .Controls("Cbocon3")
                Let VecAux0 = RES.ArrayToBaseZero(.List, eqlRes1D)
                Let StrAux1 = VBA.CStr(VecAux0(.ListIndex))
                Call RES.ArrayDelIndex(VecAux0, .ListIndex)
                Let .Tag = "noevent"
                Call .Clear
                Let .List = VecAux0
                Let StrAux0 = VBA.Join(VecAux0, ";")
                Let VecAux0 = Empty
                'Let MDL.PurDestination = StrAux0
                Let .Value = Empty
                Let .Tag = Empty
                Let BooAux0 = (.ListCount > 4)
              End With
              Let .Controls("Btncon8").Enabled = BooAux0
              Let .Controls("Btncon9").Enabled = BooAux0
              Call VBA.MsgBox("El dato '" & StrAux1 & "' fue eliminado.", vbInformation, STRAPPTITLE)

            ' LIQUIDACIONES: set, upd, upd-cancel, del
            ' Page 3 ------------------------------------------------------------------------------------------------------ x
            Case "Btncon8", "Btncon9" ' Liquidaciones - set, upd
              Let BooAux0 = (VBtnCurrent.Name = "Btncon8")
              If Not (.Controls("Lstcon1").ListCount > 0) And Not BooAux0 Then GoTo EH
              Let i = .Controls("Lstcon1").ListIndex + 1
              Let j = .Controls("Cbocon4").ListIndex
              Let .Controls("Lstcon1").RowSource = Empty
              Select Case j
                Case 0 ' RM
                  Let StrAux0 = VBA.Replace(.Controls("Cbocon6").Value, ".", VBA.vbNullString)
                  Let StrAux1 = .Controls("Cbocon7").Value
                  Let VecAux0 = Array(.Controls("Cbocon5").List(.Controls("Cbocon5").ListIndex, 0), .Controls("Cbocon5").List(.Controls("Cbocon5").ListIndex, 1), .Controls("Cbocon6").Value, VBA.Abs(VBA.CDbl(.Controls("Txtcon8").Value)), VBA.Abs(VBA.CDbl(.Controls("Txtcon9").Value)), VBA.Abs(VBA.CDbl(.Controls("Txtcon10").Value)), StrAux1, StrAux0, IIf(StrAux0 = "LF", 2204.6223, IIf(StrAux0 = "OT", 31.1035, 1)), IIf(StrAux1 = "%", 1, IIf(StrAux1 = "g/T", 0.032150723, 3.315072259)))
                  If BooAux0 Then
                    Call MDL.MSet(eqlMdlSheetSys, eqlMdlTblConfigpurrm_sys, VecAux0)
                  Else
                    Call MDL.MUpd(eqlMdlSheetSys, eqlMdlTblConfigpurrm_sys, VecAux0, MLngIndex:=i)
                  End If

                Case 1, 2 ' Retentions
                  Let VecAux0 = Array(.Controls("Txtcon8").Text, VBA.Abs(VBA.CDbl(.Controls("Txtcon9").Text)), 0)
                  If BooAux0 Then
                    Call MDL.MSet(eqlMdlSheetTmp, IIf(j = 1, eqlMdlTblConfigpurretentions_tmp, eqlMdlTblConfigpurretentionsother_tmp), VecAux0)
                  Else
                    Call MDL.MUpd(eqlMdlSheetTmp, IIf(j = 1, eqlMdlTblConfigpurretentions_tmp, eqlMdlTblConfigpurretentionsother_tmp), VecAux0, MLngIndex:=i)
                  End If

                Case 3 ' Penalties
                  Let VecAux0 = Array(.Controls("Cbocon5").Value, VBA.Abs(VBA.CDbl(.Controls("Txtcon8").Value)), VBA.Abs(VBA.CDbl(.Controls("Txtcon9").Value)), VBA.Abs(VBA.CDbl(.Controls("Txtcon10").Value)), VBA.Abs(VBA.CDbl(.Controls("Txtcon11").Value)), 0)
                  If BooAux0 Then
                    Call MDL.MSet(eqlMdlSheetPur, eqlMdlTblpurchasecncpenalties_pur, VecAux0)
                  Else
                    Call MDL.MUpd(eqlMdlSheetPur, eqlMdlTblpurchasecncpenalties_pur, VecAux0, MLngIndex:=i)
                  End If

                Case 4 ' Exports
                  Let VecAux0 = Array(.Controls("Txtcon8").Value, VBA.Abs(VBA.CDbl(.Controls("Txtcon8").Value)), VBA.Abs(VBA.CDbl(.Controls("Txtcon9").Value)), VBA.Abs(VBA.CDbl(.Controls("Txtcon10").Value)), IIf(.Controls("Chkcon7").Value, 1, 0), 0)
                  If BooAux0 Then
                    Call MDL.MSet(eqlMdlSheetPur, eqlMdlTblpurchasecncdedexp_pur, VecAux0)
                  Else
                    Call MDL.MUpd(eqlMdlSheetPur, eqlMdlTblpurchasecncdedexp_pur, VecAux0, MLngIndex:=i)
                  End If

              End Select
              Let VecAux0 = Empty
              Call View_controls_combos_change(.Controls("Cbocon4"), FRM)
              Call VBA.MsgBox("El registro fue " & IIf(BooAux0, "creado", "modificado") & " satisfactoriamente.", vbInformation, STRAPPTITLE)

            Case "Btncon12" ' Liquidaciones - upd cancel
              If Not (.Controls("Lstcon1").ListCount > 0) Then GoTo EH
              Call View_controls_combos_change(.Controls("Cbocon4"), FRM)
              For i = 0 To 4
                Let .Controls("Multicon").Pages(i).Enabled = True
              Next i

            Case "Btncon13" ' Liquidaciones - del
              If Not (.Controls("Lstcon1").ListCount > 0) Then GoTo EH
              Let i = .Controls("Lstcon1").ListIndex + 1 ' List row selected - model register
              Let j = .Controls("Cbocon4").ListIndex ' Menu item
              Let .Controls("Lstcon1").RowSource = Empty
              Select Case j

                Case 0: Call MDL.MDel(eqlMdlSheetSys, eqlMdlTblConfigpurrm_sys, i) ' RM
                Case 1, 2: Call MDL.MDel(eqlMdlSheetTmp, IIf(j = 1, eqlMdlTblConfigpurretentions_tmp, eqlMdlTblConfigpurretentionsother_tmp), i) ' Retentions
                Case 3: Call MDL.MDel(eqlMdlSheetPur, eqlMdlTblpurchasecncpenalties_pur, i) ' Penalties
                Case 4: Call MDL.MDel(eqlMdlSheetPur, eqlMdlTblpurchasecncdedexp_pur, i) ' Exports

              End Select
              Call View_controls_combos_change(.Controls("Cbocon4"), FRM)
              Call VBA.MsgBox("El registro seleccionado fue eliminado.", vbInformation, STRAPPTITLE)

            ' ELEMENTOS: set, upd, upd-cancel, del
            ' Page 4 ------------------------------------------------------------------------------------------------------ x
            Case "Btncon14", "Btncon15" ' Elementos - set, upd
              Let BooAux0 = (VBtnCurrent.Name = "Btncon14")
              If Not (.Controls("Lstcon2").ListCount > 0) And Not BooAux0 Then GoTo EH
              Let i = .Controls("Lstcon2").ListIndex + 1
              Let .Controls("Lstcon2").RowSource = Empty

              Let VecAux0 = Array(.Controls("Txtcon8").Text, .Controls("Txtcon9").Text, IIf(.Controls("Chkcon15").Value, "SI", "NO"))
              If BooAux0 Then ' set
                Call MDL.MSet(eqlMdlSheetSys, eqlMdlTblConfigchemicalelements_sys, VecAux0)
                ' Set new element in complex elements table
                If .Controls("Chkcon15").Value Then Call MDL.MSet(eqlMdlSheetSys, eqlMdlTblConfigchemicalelementscx_sys, Array(.Controls("Txtcon8").Text, .Controls("Txtcon9").Text))
              Else ' upd
                Call MDL.MUpd(eqlMdlSheetSys, eqlMdlTblConfigchemicalelements_sys, VecAux0, MLngIndex:=i)
                ' Upd element in complex elements table
                Let VecAux1 = MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalelementscx_sys, eqlMdlArray)
                If .Controls("Chkcon15").Value Then
                  If Not RES.ArrayHasValue2D(VecAux1, .Controls("Txtcon8").Text, 0) Then
                    Call MDL.MSet(eqlMdlSheetSys, eqlMdlTblConfigchemicalelementscx_sys, Array(.Controls("Txtcon8").Text, .Controls("Txtcon9").Text))
                  End If
                Else
                  Let j = RES.ArrayHasValueN2D(VecAux1, .Controls("Txtcon8").Text, 0) + 1
                  If j > 0 Then Call MDL.MDel(eqlMdlSheetSys, eqlMdlTblConfigchemicalelementscx_sys, j)
                End If
              End If
              Let VecAux0 = Empty
              Let VecAux1 = Empty

              Let .Controls("Lstcon2").RowSource = MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalelements_sys, eqlMdlStrTableName)
              Call VBA.MsgBox("El registro fue " & IIf(BooAux0, "creado", "modificado") & " satisfactoriamente.", vbInformation, STRAPPTITLE)

            Case "Btncon16" ' Elementos - upd cancel
              If .Controls("Lstcon2").ListCount > 0 Then
                For i = 0 To 4
                  If i < 4 Then Let .Controls("Btncon" & i + 14).Enabled = IIf(i = 0, True, IIf(i < 3, False, (.Controls("Lstcon2").ListCount > 60)))
                  Let FRM.Controls("Multicon").Pages(i).Enabled = True
                  If i < 2 Then Let .Controls("Txtcon" & i + 13).Text = Empty: Let .Controls("Txtcon" & i + 13).Enabled = True
                Next i
                Let .Controls("Lstcon2").Enabled = True
              End If

            Case "Btncon17" ' Elementos - del
              If Not (.Controls("Lstcon2").ListCount > 0) Then GoTo EH
              Let i = .Controls("Lstcon2").ListIndex + 1 ' List row selected - model register
              Let StrAux0 = VBA.CStr(.Controls("Lstcon2").List(i, 0))
              Let .Controls("Lstcon2").RowSource = Empty
              Call MDL.MDel(eqlMdlSheetSys, eqlMdlTblConfigchemicalelements_sys, i) ' del
              ' Upd element in complex elements table
              Let VecAux0 = MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalelementscx_sys, eqlMdlArray)
              Let j = RES.ArrayHasValueN2D(VecAux1, StrAux0, 0) + 1
              If j > 0 Then Call MDL.MDel(eqlMdlSheetSys, eqlMdlTblConfigchemicalelementscx_sys, j)
              Let VecAux0 = Empty
              Let .Controls("Lstcon2").RowSource = MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalelements_sys, eqlMdlStrTableName)
              Call VBA.MsgBox("El registro seleccionado fue eliminado.", vbInformation, STRAPPTITLE)

          End Select
        End With

      Case Else: GoTo EH
    End Select

  End If

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Set RngBox = Nothing
  Set VWP = Nothing
  Call RES.ErrorHandler.ErrorHandlerDisplay(sFN)

End Function

' Events handlers - Checkboxes
Private Sub View_controls_checkboxes_change( _
  ByVal VChkCurrent As MSForms.CheckBox, _
  ByVal FRM As FRMBasics)

  Dim VecAux0 As Variant, VecAux1 As Variant
  Dim StrAux0 As String, StrAux1 As String, StrAux2 As String
  Dim LngX As Long, i As Long, j As Long, k As Long
  Dim EnuSht As EQLMDL_ENU_SHEETS, EnuTbl As EQLMDL_ENU_TABLES

  Const sFN As String = "VIEW::-checkboxes_change"

  On Error GoTo EH

  With FRM
    If Not CMB Is Nothing Then
      If CMB.IsActive Then GoTo EH
    End If

    Select Case .Tag

      Case "balmet" ' ------------------------------------------------ balmet'
        Select Case VChkCurrent.Name

          Case "Chkbal0"
            Let .Controls("Txtbal1").Locked = Not VChkCurrent.Value
            Let .Controls("Txtbal1").ForeColor = IIf(VChkCurrent.Value, .ForeColor, VBA.vbGrayText)
            Let .Controls("Txtbal1").Text = IIf(VChkCurrent.Value, VBA.vbNullString, "A")
            If VChkCurrent.Value Then Call .Controls("Txtbal1").SetFocus

          Case "Chkbal1"
            With .Controls("Txtbal2")
              If RES.REGEX.isNumberEntireAndDecimalVectorColonSeparated(.Text) Then
                Let StrAux0 = VBA.vbNullString
                If Not VChkCurrent.Value Then Let StrAux0 = RES.RESCORE.GetGenericNamesAndUnits(.Text, .Tag, StrAux1)
                Let .Parent.Controls("Txtbal3").ForeColor = IIf(VChkCurrent.Value, .ForeColor, VBA.vbGrayText)
                Let .Parent.Controls("Txtbal3").Text = StrAux0
                Let .Parent.Controls("Cbobal1").Enabled = VChkCurrent.Value
                Let .Parent.Controls("Btnbal0").Enabled = VChkCurrent.Value
                Let .Parent.Controls("Btnbal1").Enabled = VChkCurrent.Value
                If VChkCurrent.Value Then Call .Parent.Controls("Cbobal1").SetFocus
              End If
            End With

          Case "Chkbal2"
            With .Controls("Txtbal2")
              If RES.REGEX.isNumberEntireAndDecimalVectorColonSeparated(.Text) Then
                Let StrAux0 = VBA.vbNullString
                If Not VChkCurrent.Value Then Let StrAux0 = RES.RESCORE.GetGenericNamesAndUnits(.Text, .Tag, StrAux1, BBooUnits:=True)
                Let .Parent.Controls("Txtbal4").ForeColor = IIf(VChkCurrent.Value, .ForeColor, VBA.vbGrayText)
                Let .Parent.Controls("Txtbal4").Text = StrAux0
                Let .Parent.Controls("Cbobal2").Enabled = VChkCurrent.Value
                Let .Parent.Controls("Btnbal2").Enabled = VChkCurrent.Value
                Let .Parent.Controls("Btnbal3").Enabled = VChkCurrent.Value
                If VChkCurrent.Value Then Call .Parent.Controls("Cbobal2").SetFocus
              End If
            End With

          Case "Chkbal4"
            With .Controls("Txtbal5")
              If RES.REGEX.isNumberEntireAndDecimalVectorColonSeparated(.Text) Then
                Let StrAux0 = VBA.vbNullString
                If Not VChkCurrent.Value Then Let StrAux0 = RES.RESCORE.GetGenericNamesAndUnits(FRM.Controls("Txtbal2").Text, StrAux1, .Tag, .Text, BBooCx:=True)
                Let .Parent.Controls("Txtbal6").ForeColor = IIf(VChkCurrent.Value, .ForeColor, VBA.vbGrayText)
                Let .Parent.Controls("Txtbal6").Text = StrAux0
                Let .Parent.Controls("Cbobal3").Enabled = VChkCurrent.Value
                Let .Parent.Controls("Btnbal4").Enabled = VChkCurrent.Value
                Let .Parent.Controls("Btnbal5").Enabled = VChkCurrent.Value
                If VChkCurrent.Value Then Call .Parent.Controls("Cbobal3").SetFocus
              End If
            End With

          Case "Chkbal5"
            With .Controls("Txtbal5")
              If RES.REGEX.isNumberEntireAndDecimalVectorColonSeparated(.Text) Then
                Let StrAux0 = VBA.vbNullString
                If Not VChkCurrent.Value Then Let StrAux0 = RES.RESCORE.GetGenericNamesAndUnits(FRM.Controls("Txtbal2").Text, StrAux1, .Tag, .Text, BBooCx:=True, BBooUnits:=True)
                Let .Parent.Controls("Txtbal7").ForeColor = IIf(VChkCurrent.Value, .ForeColor, VBA.vbGrayText)
                Let .Parent.Controls("Txtbal7").Text = StrAux0
                Let .Parent.Controls("Cbobal4").Enabled = VChkCurrent.Value
                Let .Parent.Controls("Btnbal6").Enabled = VChkCurrent.Value
                Let .Parent.Controls("Btnbal7").Enabled = VChkCurrent.Value
                If VChkCurrent.Value Then Call .Parent.Controls("Cbobal4").SetFocus
              End If
            End With

          Case "Chkbal6"
            For i = 7 To 13
              With .Controls("Chkbal" & i)
                Let .Value = VChkCurrent.Value
                Let .Enabled = Not VChkCurrent.Value
              End With
            Next i

        End Select

      Case "balmeteco" ' ------------------------------------------------ balmeteco'
        Let i = -1
        Select Case VChkCurrent.Name

          Case "Chkbal0"
            Let .Controls("Txtbal3").Locked = Not VChkCurrent.Value
            Let .Controls("Txtbal3").ForeColor = IIf(VChkCurrent.Value, .ForeColor, VBA.vbGrayText)
            If Not .Controls("Txtbal3").Tag = VBA.vbNullString Then Let i = VBA.CLng(.Controls("Txtbal3").Tag)
            Let StrAux0 = RES.RESCORE.BalGetGenericRecoveries(i)
            Let .Controls("Txtbal3").Text = IIf(VChkCurrent.Value, VBA.vbNullString, StrAux0)
            If VChkCurrent.Value Then Call .Controls("Txtbal3").SetFocus

          Case "Chkbal1"
            Let .Controls("Txtbal4").Enabled = VChkCurrent.Value
            Let .Controls("Txtbal4").Text = IIf(VChkCurrent.Value, VBA.vbNullString, "A")
            If VChkCurrent.Value Then Call .Controls("Txtbal4").SetFocus

          Case "Chkbal2"
            Let .Controls("Txtbal5").ForeColor = IIf(VChkCurrent.Value, .ForeColor, VBA.vbGrayText)
            If Not VChkCurrent.Value Then Let .Controls("Cbobal2").ListIndex = 0
            If Not .Controls("Txtbal1").Tag = VBA.vbNullString Then Let i = VBA.CLng(.Controls("Txtbal1").Tag)
            Let StrAux0 = RES.RESCORE.BalGetGenericNames(i)
            Let .Controls("Txtbal5").Text = IIf(VChkCurrent.Value, VBA.vbNullString, StrAux0)
            Let .Controls("Cbobal2").Enabled = VChkCurrent.Value
            Let .Controls("Btnbal0").Enabled = VChkCurrent.Value
            Let .Controls("Btnbal2").Enabled = VChkCurrent.Value
            If VChkCurrent.Value Then Call .Controls("Cbobal2").SetFocus

          Case "Chkbal3"
            Let .Controls("Txtbal6").ForeColor = IIf(VChkCurrent.Value, .ForeColor, VBA.vbGrayText)
            If Not VChkCurrent.Value Then Let .Controls("Cbobal3").ListIndex = 0
            If Not .Controls("Txtbal2").Tag = VBA.vbNullString Then Let i = VBA.CLng(.Controls("Txtbal2").Tag)
            Let StrAux0 = RES.RESCORE.BalGetGenericUnits(i)
            Let .Controls("Txtbal6").Text = IIf(VChkCurrent.Value, VBA.vbNullString, StrAux0)
            Let .Controls("Cbobal3").Enabled = VChkCurrent.Value
            Let .Controls("Btnbal1").Enabled = VChkCurrent.Value
            Let .Controls("Btnbal3").Enabled = VChkCurrent.Value
            If VChkCurrent.Value Then Call .Controls("Cbobal3").SetFocus

          Case "Chkbal5"
            If Not .Controls("Chkbal4").Value Then GoTo EH
            Let .Controls("Txtbal8").ForeColor = IIf(VChkCurrent.Value, .ForeColor, VBA.vbGrayText)
            If Not VChkCurrent.Value Then Let .Controls("Cbobal4").ListIndex = 0
            If Not .Controls("Txtbal7").Tag = VBA.vbNullString Then Let i = VBA.CLng(.Controls("Txtbal7").Tag)
            Let StrAux0 = RES.RESCORE.BalGetGenericNames(i, BBooCx:=True)
            Let .Controls("Txtbal8").Text = IIf(VChkCurrent.Value, VBA.vbNullString, StrAux0)
            Let .Controls("Cbobal4").Enabled = VChkCurrent.Value
            Let .Controls("Btnbal4").Enabled = VChkCurrent.Value
            Let .Controls("Btnbal5").Enabled = VChkCurrent.Value
            If VChkCurrent.Value Then Call .Controls("Cbobal4").SetFocus

          Case "Chkbal6"
            If Not .Controls("Chkbal4").Value Then GoTo EH
            Let .Controls("Txtbal9").ForeColor = IIf(VChkCurrent.Value, .ForeColor, VBA.vbGrayText)
            If Not VChkCurrent.Value Then Let .Controls("Cbobal5").ListIndex = 0
            If Not .Controls("Txtbal7").Tag = VBA.vbNullString Then Let i = VBA.CLng(.Controls("Txtbal7").Tag)
            Let StrAux0 = RES.RESCORE.BalGetGenericUnits(i, BBooCx:=True)
            Let .Controls("Txtbal9").Text = IIf(VChkCurrent.Value, VBA.vbNullString, StrAux0)
            Let .Controls("Cbobal5").Enabled = VChkCurrent.Value
            Let .Controls("Btnbal6").Enabled = VChkCurrent.Value
            Let .Controls("Btnbal7").Enabled = VChkCurrent.Value
            If VChkCurrent.Value Then Call .Controls("Cbobal5").SetFocus

          Case "Chkbal7"
            If Not .Controls("Chkbal4").Value Then GoTo EH
            Let .Controls("Txtbal10").Locked = Not VChkCurrent.Value
            Let .Controls("Txtbal10").ForeColor = IIf(VChkCurrent.Value, .ForeColor, VBA.vbGrayText)
            If Not .Controls("Txtbal7").Tag = VBA.vbNullString Then Let i = VBA.CLng(.Controls("Txtbal7").Tag)
            Let StrAux0 = RES.RESCORE.BalGetGenericRecoveries(i)
            Let .Controls("Txtbal10").Text = IIf(VChkCurrent.Value, VBA.vbNullString, StrAux0)
            If VChkCurrent.Value Then Call .Controls("Txtbal10").SetFocus

          Case "Chkbal8"
            For i = 9 To 15
              With .Controls("Chkbal" & i)
                Let .Value = VChkCurrent.Value
                Let .Enabled = Not VChkCurrent.Value
              End With
            Next i

        End Select

      Case "oreinvoice" ' ------------------------------------------------ oreinvoice'
        Select Case VChkCurrent.Name

          Case "Chkpur1"
            ' Set Tabindex
            If VChkCurrent.Value Then
              Let VecAux0 = Array("Lblpur0", "Txtpur0", "Lblpur1", "Txtpur1", "Lblpur2", "Txtpur2", "Chkpur0", "Chkpur1", "Lblpur3", "Txtpur3", "Lblpur4", "Txtpur4", "Cbopur0", "Cbopur1", "Btnpur0", "Btnpur1", "Btnpur2", "Btnpur3", "Btnpur4", "Btnpur5", "Txtpur5", "Txtpur6", "Cbopur2", "Chkpur2", "Txtpur7", "Txtpur8", "Txtpur9", "Lstpur0", "Chkpur3", "Chkpur4", "Chkpur5", "Optpur0", "Optpur1", "Txtpur10", "Btnmain0", "Btnmain1", "Btnmain2")
            Else
              Let VecAux0 = Array("Lblpur0", "Txtpur0", "Lblpur1", "Txtpur1", "Lblpur2", "Txtpur2", "Chkpur0", "Chkpur1", "Lblpur3", "Txtpur3", "Lblpur4", "Cbopur0", "Txtpur4", "Cbopur1", "Btnpur0", "Btnpur1", "Btnpur2", "Btnpur3", "Btnpur4", "Btnpur5", "Txtpur5", "Txtpur6", "Cbopur2", "Chkpur2", "Txtpur7", "Txtpur8", "Txtpur9", "Lstpur0", "Chkpur3", "Chkpur4", "Chkpur5", "Optpur0", "Optpur1", "Txtpur10", "Btnmain0", "Btnmain1", "Btnmain2")
            End If
            For i = 0 To UBound(VecAux0)
              Let .Controls(VBA.CStr(VecAux0(i))).TabIndex = i
            Next i
            Let VecAux0 = Empty
            Let .Controls("Lblpur3").Caption = IIf(VChkCurrent.Value, "Provedores:", "Provedor:")
            Let .Controls("Lblpur4").Caption = IIf(VChkCurrent.Value, "Tipo Provedores:", "Tipo Provedor:")
            Let .Controls("Txtpur4").Visible = VChkCurrent.Value
            Let .Controls("Cbopur0").Visible = Not VChkCurrent.Value
            Let .Controls("Cbopur1").ListIndex = 0
            Let .Controls("Chkpur2").Value = False
            Let .Controls("Chkpur3").Enabled = Not VChkCurrent.Value
            Let .Controls("Chkpur3").Value = False

          Case "Chkpur2"
            ' Save current MenuCbo index
            Debug.Print .Controls("Chkpur2").Tag
            Let i = .Controls("Cbopur1").ListIndex
            If .Controls("Chkpur2").Tag = "noevent" Then GoTo EH ' Accept events
            If .Controls("Lstpur0").ListCount <= 0 Then GoTo EH ' Empty list

            If i = 3 Or i = 4 Then ' Change list selections
              With .Controls("Lstpur0")
                Let k = .ListCount
                Let .Tag = "noevent"
                Let .RowSource = VBA.vbNullString ' Unlink list
                ' Set 0 or 1 LO Selected field and add new values to pur table values
                ReDim VecAux0(k - 1)
                For j = 1 To k
                  Call MDL.MUpd(eqlMdlSheetTmp, IIf(i = 3, eqlMdlTblConfigpurretentions_tmp, eqlMdlTblConfigpurretentionsother_tmp), Array(IIf(VChkCurrent.Value, 1, 0)), "SELECTED", j)
                  Let VecAux0(j - 1) = IIf(VChkCurrent.Value, 1, 0)
                Next j
                Call MDL.MUpd(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur, Array(VBA.Join(VecAux0, ";")), IIf(i = 3, "RETENCIONES", "RETENCIONES OTRAS"), IIf(FRM.Controls("Chkpur1").Value, FRM.Controls("Cbopur2").ListIndex + 1, 1)) ' Multiple providers: Lsti+1, 1
                Let VecAux0 = Empty
                ' Link list
                Let .RowSource = MDL.MGet(eqlMdlSheetTmp, IIf(i = 3, eqlMdlTblConfigpurretentions_tmp, eqlMdlTblConfigpurretentionsother_tmp), eqlMdlListObject)
                ' Highlight selected items (Set tag=empty)
                For j = 0 To k - 1
                  Let .Selected(j) = VChkCurrent.Value
                Next j
                Let .Tag = VBA.vbNullString
              End With
            End If

          Case Else: GoTo EH

        End Select

      Case "concentrateinvoice" ' ------------------------------------------------ concentrateinvoice'
        Select Case VChkCurrent.Name

          Case "Chkcnc1"
            ' Save current MenuCbo index
            Let i = .Controls("Cbocnc1").ListIndex
            'If .Controls("Chkcnc2").Tag = "noevent" Then GoTo EH ' Accept events
            If .Controls("Lstcnc0").ListCount <= 0 Then GoTo EH ' Empty list

            If i = 2 Or i = 3 Then ' Change list selections
              Let j = .Controls("Cbocnc2").ListIndex
              Let EnuSht = IIf(i = 2, eqlMdlSheetPur, IIf(j = 0, eqlMdlSheetPur, eqlMdlSheetTmp))
              Let EnuTbl = IIf(i = 2, eqlMdlTblpurchasecncpenalties_pur, IIf(j = 0, eqlMdlTblpurchasecncdedexp_pur, IIf(j = 1, eqlMdlTblConfigpurretentions_tmp, eqlMdlTblConfigpurretentionsother_tmp)))
              Set Lst0 = Nothing
              With .Controls("Lstcnc0")
                Let k = .ListCount - 1
                Let .RowSource = Empty ' Unlink list
                ' Set 0 or 1 LO Selected field and add new values to cnc table values
                Call MDL.MUpd(EnuSht, EnuTbl, Array(IIf(VChkCurrent.Value, 1, 0)), "SELECTED")
                ' Link list
                Let .RowSource = MDL.MGet(EnuSht, EnuTbl, eqlMdlStrTableName)
                ' Highlight selected items (Set tag=empty)
                If VChkCurrent.Value Then
                  For j = 0 To k
                    Let .Selected(j) = VChkCurrent.Value
                  Next j
                End If
              End With
              Set Lst0 = .Controls("Lstcnc0")
            End If

          Case Else: GoTo EH

        End Select

      Case "configlauncher" ' ---------------------------------------------------- configlauncher'
        Select Case VChkCurrent.Name
          Case "Chkcon0": Let MDL.TableSelectionsCompact = VChkCurrent.Value
          Case "Chkcon1": Let MDL.TableSelectionsHeads = VChkCurrent.Value

          Case "Chkcon2", "Chkcon3", "Chkcon4", "Chkcon5", "Chkcon6"
            With VChkCurrent
              Let .Caption = IIf(.Value = False, "En encabezado de página", IIf(.Value = True, "En pie de página", "En ambos"))
            End With

          Case "Chkcon8": Let MDL.PurAverage = VChkCurrent.Value
          Case "Chkcon9": Let MDL.PurRMDetailed = VChkCurrent.Value
          Case "Chkcon10": Let MDL.PurRetentions = VChkCurrent.Value
          Case "Chkcon11": Let MDL.PurRetentionsOther = VChkCurrent.Value
          Case "Chkcon12": Let MDL.PurDiscounts = VChkCurrent.Value
          Case "Chkcon13": Let MDL.PurBonus = VChkCurrent.Value
          Case "Chkcon14": Let MDL.PurExports = VChkCurrent.Value

        End Select

      Case Else: GoTo EH
    End Select
  End With

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Call RES.ErrorHandler.ErrorHandlerDisplay(sFN)

End Sub

' Events handlers - Combos
Private Sub View_controls_combos_change(ByVal VCboCurrent As MSForms.ComboBox, ByVal FRM As FRMBasics)

  Dim VecAux0 As Variant, VecAux1 As Variant
  Dim StrAux0 As String, StrAux1 As String
  Dim xLeft As Long, i As Long, j As Long, k As Long, LngPurRows As Long

  Const sFN As String = "VIEW::-comboEvents"

  On Error GoTo EH
  If Not CMB Is Nothing Then
    If CMB.IsActive Then GoTo EH
  End If

  Select Case FRM.Tag

    Case "balmet" ' -------------------------------------------------- balmet'
      With FRM
        Select Case VCboCurrent.Name

          Case "Cbobal1" ' Names
            If VCboCurrent.ListIndex <= 0 Then GoTo EH
            Let VecAux0 = VBA.Split(.Controls("Txtbal6").Text, ";")
            If RES.ArrayHasValue(VecAux0, VCboCurrent.Value) Then GoTo EH
            Let VecAux0 = VBA.Split(.Controls("Txtbal3").Text, ";")
            If RES.ArrayHasValue(VecAux0, VCboCurrent.Value) Then GoTo EH
            Let i = UBound(VBA.Split(.Controls("Txtbal3").Text, ";")) + 1
            Let j = VBA.CLng(.Controls("Txtbal2").Tag)
            If i <= j Then Let .Controls("Txtbal3").Text = VBA.Join(RES.ArrayAddAtLast(VecAux0, VCboCurrent.Value), ";")

          Case "Cbobal2" ' Units
            If VCboCurrent.ListIndex <= 0 Then GoTo EH
            Let VecAux0 = VBA.Split(.Controls("Txtbal4").Text, ";")
            If RES.ArrayHasValue(VecAux0, "dm", ALCase:=True) And VBA.LCase(VCboCurrent.Value) = "dm" Then GoTo EH
            Let i = UBound(VBA.Split(.Controls("Txtbal4").Text, ";")) + 1
            Let j = VBA.CLng(.Controls("Txtbal2").Tag)
            If i <= j Then Let .Controls("Txtbal4").Text = VBA.Join(RES.ArrayAddAtLast(VecAux0, VCboCurrent.Value), ";")

          Case "Cbobal3" ' Namescx
            If VCboCurrent.ListIndex <= 0 Then GoTo EH
            Let VecAux0 = VBA.Split(.Controls("Txtbal3").Text, ";")
            If RES.ArrayHasValue(VecAux0, VCboCurrent.Value) Then GoTo EH
            Let VecAux0 = VBA.Split(.Controls("Txtbal6").Text, ";")
            If RES.ArrayHasValue(VecAux0, VCboCurrent.Value) Then GoTo EH
            Let i = UBound(VBA.Split(.Controls("Txtbal6").Text, ";")) + 1
            Let j = VBA.CLng(.Controls("Txtbal5").Tag)
            If i <= j Then Let .Controls("Txtbal6").Text = VBA.Join(RES.ArrayAddAtLast(VecAux0, VCboCurrent.Value), ";")

          Case "Cbobal4" ' Unitscx
            If VCboCurrent.ListIndex <= 0 Then GoTo EH
            Let VecAux0 = VBA.Split(.Controls("Txtbal7").Text, ";")
            If RES.ArrayHasValue(VecAux0, "dm", ALCase:=True) And VBA.LCase(VCboCurrent.Value) = "dm" Then GoTo EH
            Let i = UBound(VBA.Split(.Controls("Txtbal7").Text, ";")) + 1
            Let j = VBA.CLng(.Controls("Txtbal5").Tag)
            If i <= j Then Let .Controls("Txtbal7").Text = VBA.Join(RES.ArrayAddAtLast(VecAux0, VCboCurrent.Value), ";")

          Case Else: GoTo EH
        End Select
      End With

    Case "balmeteco" ' ----------------------------------------------- balmeteco'
      With FRM
        Select Case VCboCurrent.Name
          Case "Cbobal0"
            Let .Controls("Lblbal4").Caption = IIf(VCboCurrent.ListIndex = 0, "Ley Producto/s (Conc.):", "Ley Cabeza/s (Alim.):")
            Let .Controls("Lblbal5").WordWrap = False
            Let .Controls("Lblbal5").Caption = .Controls("Lblbal4").Caption
            Let .Controls("Lblbal5").AutoSize = True

          Case "Cbobal2" ' Names
            If VCboCurrent.ListIndex <= 0 Then GoTo EH
            Let VecAux0 = VBA.Split(.Controls("Txtbal5").Text, ";")
            If RES.ArrayHasValue(VecAux0, VCboCurrent.Value) Then GoTo EH
            Let i = UBound(VecAux0) + 1: Let j = -1
            If Not .Controls("Txtbal1").Tag = VBA.vbNullString Then Let j = VBA.CLng(.Controls("Txtbal1").Tag)
            If i <= j Then Let .Controls("Txtbal5").Text = VBA.Join(RES.ArrayAddAtLast(VecAux0, VCboCurrent.Value), ";")

          Case "Cbobal3" ' Units
            If VCboCurrent.ListIndex <= 0 Then GoTo EH
            Let VecAux0 = VBA.Split(.Controls("Txtbal6").Text, ";")
            If RES.ArrayHasValue(VecAux0, "dm", ALCase:=True) And VBA.LCase(VCboCurrent.Value) = "dm" Then GoTo EH
            Let i = UBound(VecAux0) + 1: Let j = -1
            If Not .Controls("Txtbal2").Tag = VBA.vbNullString Then Let j = VBA.CLng(.Controls("Txtbal2").Tag)
            If i <= j Then Let .Controls("Txtbal6").Text = VBA.Join(RES.ArrayAddAtLast(VecAux0, VCboCurrent.Value), ";")

          Case "Cbobal4" ' Namescx
            If VCboCurrent.ListIndex <= 0 Then GoTo EH
            Let VecAux0 = VBA.Split(.Controls("Txtbal8").Text, ";")
            If RES.ArrayHasValue(VecAux0, VCboCurrent.Value) Then GoTo EH
            Let i = UBound(VecAux0) + 1: Let j = -1
            If Not .Controls("Txtbal7").Tag = VBA.vbNullString Then Let j = VBA.CLng(.Controls("Txtbal7").Tag)
            If i <= j Then Let .Controls("Txtbal8").Text = VBA.Join(RES.ArrayAddAtLast(VecAux0, VCboCurrent.Value), ";")

          Case "Cbobal5" ' Unitscx
            If VCboCurrent.ListIndex <= 0 Then GoTo EH
            Let VecAux0 = VBA.Split(.Controls("Txtbal9").Text, ";")
            If RES.ArrayHasValue(VecAux0, "dm", ALCase:=True) And VBA.LCase(VCboCurrent.Value) = "dm" Then GoTo EH
            Let i = UBound(VecAux0) + 1: Let j = -1
            If Not .Controls("Txtbal7").Tag = VBA.vbNullString Then Let j = VBA.CLng(.Controls("Txtbal7").Tag)
            If i <= j Then Let .Controls("Txtbal9").Text = VBA.Join(RES.ArrayAddAtLast(VecAux0, VCboCurrent.Value), ";")

          Case Else: GoTo EH
        End Select
      End With

    Case "projection" ' ----------------------------------------------- projection'
      With VCboCurrent
        If .ListIndex <= 0 Then GoTo EH
        Let StrAux0 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(FRM.Controls("Txtpro1").Text, i, RValueType:=eqlResNumbers, RBooAbs:=True)
        If i <= 0 Then GoTo EH
        Select Case .Name
          Case "Cbopro1"
            Let VecAux0 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(FRM.Controls("Txtpro5").Text, j, RValueType:=eqlResNotNumbers, RValueAs:=eqlResNotNumbers)
            If j > 0 Then
              If RES.ArrayHasValue(VecAux0, .Value) Then GoTo EH
            End If
            If j >= i Then GoTo EH
            If j <= i Then Let FRM.Controls("Txtpro5").Text = VBA.Join(RES.ArrayAddAtLast(VecAux0, .Value), ";")
          Case "Cbopro2"
            Let VecAux0 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(FRM.Controls("Txtpro6").Text, j, RValueType:=eqlResNotNumbers, RValueAs:=eqlResNotNumbers)
            If j > 0 Then
              If RES.ArrayHasValue(VecAux0, .Value) Then GoTo EH
            End If
            If j >= i Then GoTo EH
            If j <= i Then Let FRM.Controls("Txtpro6").Text = VBA.Join(RES.ArrayAddAtLast(VecAux0, .Value), ";")
        End Select
      End With

    Case "oreinvoice" ' ----------------------------------------------- oreinvoice'
      With FRM
        If VCboCurrent.ListIndex < 0 Then GoTo EH
        Let xLeft = .InsideWidth * 0.05 ' Left for all main controls
        Let LngPurRows = MDL.LengthRows(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur)
        Select Case VCboCurrent.Name

          Case "Cbopur1" '_____________________________________________ Cbopur1_
            ' Deactivate controls
            Set Lst0 = Nothing ' Lstpur0 - Frame2
            Set Chk1 = Nothing ' Chkpur2 - Frame2
            ' Reset controls
            For i = 5 To 9
              With .Controls("Txtpur" & i)
                Let .Value = VBA.vbNullString
                Let .Width = (FRM.Controls("Framepur2").InsideWidth - (xLeft * 2) + 4) / 5
                Let .Left = FRM.Controls("Cbopur1").Left + (FRM.Controls("Txtpur" & i).Width * (i - 5))
                Let .ShowDropButtonWhen = fmShowDropButtonWhenAlways
                Let .Visible = Not (VCboCurrent.ListIndex = 3 Or VCboCurrent.ListIndex = 4)
                Let .MaxLength = 0
                If i < 8 Then Let FRM.Controls("Cbopur" & i - 3).Visible = (VCboCurrent.ListIndex = 1): Let FRM.Controls("Cbopur" & i - 3).Top = FRM.Controls("Txtpur5").Top
              End With
            Next i
            Let .Controls("Chkpur2").Visible = (VCboCurrent.ListIndex = 3 Or VCboCurrent.ListIndex = 4)
            Let .Controls("Lstpur0").Visible = False
            Let .Controls("Lstpur0").Tag = "noevent"
            Let .Controls("Lstpur0").RowSource = VBA.vbNullString

            Select Case VCboCurrent.ListIndex

              Case 0 ' ------------------------------------------- Base Data
                Call View_assets_oreinvoice_basedata(VCboCurrent, FRM, LngPurRows)

              Case 1 ' ------------------------------------------- Main Data
                Call View_assets_oreinvoice_maindata(VCboCurrent, FRM, LngPurRows)

              Case 2, 5 ' ---------------------------------------- BONUS, DISCOUNTS
                Call View_assets_oreinvoice_bonusdiscounts(VCboCurrent, FRM, LngPurRows)

              Case 3, 4 ' ---------------------------------------- RETENTIONS, OTHER
                Call View_assets_oreinvoice_retentions(VCboCurrent, FRM, LngPurRows)

              Case Else: GoTo EH

            End Select

            ' Activate controls
            Set Lst0 = .Controls("Lstpur0") ' Frame2
            Set Chk1 = .Controls("Chkpur2") ' Frame2

          Case "Cbopur2" '_____________________________________________ Cbopur2_

            Let k = .Controls("Cbopur1").ListIndex
            If VCboCurrent.Tag = "noevent" Then GoTo EH
            Select Case .Controls("Cbopur1").ListIndex
              Case 2, 5 ' Bonus or Discounts
                If LngPurRows > 0 Then Call MDL.MDel(eqlMdlSheetTmp, IIf(k = 2, eqlMdlTblPurbonus_tmp, eqlMdlTblPurdiscounts_tmp))
                If .Controls("Chkpur1").Value And VCboCurrent.ListIndex >= 0 Then ' Just for Multiple providers
                  ' Get SELECTED values in array
                  If k = 2 Or k = 5 Then
                    ' Unlink list
                    Let .Controls("Lstpur0").Tag = "noevent"
                    Let .Controls("Lstpur0").RowSource = Empty
                    If LngPurRows > 0 Then
                      Let VecAux0 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur, eqlMdlArray, IIf(k = 2, "BONOS", "DESCUENTOS"), VCboCurrent.ListIndex + 1), eqlRes1D)
                      If VBA.IsArray(VecAux0) Then
                        If Not VBA.CStr(VecAux0(0)) = VBA.vbNullString Then
                          Let VecAux1 = VBA.Split(VBA.CStr(VecAux0(0)), "|") ' Separe Concept | Values
                          Let VecAux0 = RES.ArrayToBaseZero(VBA.Split(VBA.CStr(VecAux1(0)), ";"), eqlRes2D, False)  ' Separe Concept1 ; Concept2...
                          For i = 0 To UBound(VecAux0)
                            Let VecAux0(i, 0) = VBA.Replace(VBA.Replace(VBA.CStr(VecAux0(i, 0)), "[eqlsemicolon]", ";"), "[eqlpipe]", "|")
                          Next i
                          Let VecAux0 = RES.ArrayConcat(VecAux0, VBA.Split(VBA.CStr(VecAux1(1)), ";"), RByCols:=True) ' Join Concepts - Values
                          Let VecAux0 = RES.ArrayAddAtFirst(VecAux0, .Controls("Cbopur2").Value, True) ' Join #Sample - Concepts - Values
                          Call MDL.MSet(eqlMdlSheetTmp, IIf(k = 2, eqlMdlTblPurbonus_tmp, eqlMdlTblPurdiscounts_tmp), VecAux0)
                        End If
                      End If
                    End If
                    ' Link list
                    Let .Controls("Lstpur0").RowSource = MDL.MGet(eqlMdlSheetTmp, IIf(k = 2, eqlMdlTblPurbonus_tmp, eqlMdlTblPurdiscounts_tmp), eqlMdlStrTableName) ' Link list
                    Let .Controls("Lstpur0").Tag = VBA.vbNullString
                  End If
                End If

              Case 3, 4 ' Retentions or Retentions others
                Let .Controls("Lstpur0").RowSource = Empty ' Unlink list
                Let k = .Controls("Cbopur1").ListIndex
                Let j = MDL.LengthRows(eqlMdlSheetTmp, IIf(k = 3, eqlMdlTblConfigpurretentions_tmp, eqlMdlTblConfigpurretentionsother_tmp))
                For i = 1 To j
                  Call MDL.MUpd(eqlMdlSheetTmp, IIf(k = 3, eqlMdlTblConfigpurretentions_tmp, eqlMdlTblConfigpurretentionsother_tmp), Array(0), "SELECTED", i)
                Next i
                If .Controls("Chkpur1").Value And VCboCurrent.ListIndex >= 0 Then ' Just for Multiple providers
                  ' Get SELECTED values in array
                  If LngPurRows > 0 Then
                    Let VecAux0 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur, eqlMdlArray, IIf(k = 3, "RETENCIONES", "RETENCIONES OTRAS")), eqlRes1D)
                    If VBA.IsArray(VecAux0) Then
                      If Not VBA.CStr(VecAux0(VCboCurrent.ListIndex)) = VBA.vbNullString Then
                        Let VecAux0 = VBA.Split(VBA.CStr(VecAux0(VCboCurrent.ListIndex)), ";")
                        For i = 1 To j
                          Call MDL.MUpd(eqlMdlSheetTmp, IIf(k = 3, eqlMdlTblConfigpurretentions_tmp, eqlMdlTblConfigpurretentionsother_tmp), Array(VBA.CByte(VecAux0(i - 1))), "SELECTED", i)
                        Next i
                      End If
                    End If
                  End If
                End If
                Let VCboCurrent.Tag = "noevent"
                Let .Controls("Cbopur1").Tag = "FillList"
                With .Controls("Lstpur0")
                  Let .Tag = "noevent"
                  Let .RowSource = MDL.MGet(eqlMdlSheetTmp, IIf(k = 3, eqlMdlTblConfigpurretentions_tmp, eqlMdlTblConfigpurretentionsother_tmp), eqlMdlStrTableName) ' Link list
                  If .ListCount > 0 Then ' Mark all previously selected items in the list
                    Let j = 0
                    For i = 0 To .ListCount - 1
                      Let .Selected(i) = VBA.CBool(.List(i, 2))
                      Let j = j + IIf(.Selected(i), 1, 0)
                    Next i
                  End If
                  Let .Tag = VBA.vbNullString
                End With
                Let .Controls("Chkpur2").Tag = "noevent"
                Let .Controls("Chkpur2").Value = (j = .Controls("Lstpur0").ListCount)
                Let .Controls("Chkpur2").Tag = VBA.vbNullString
                Let .Controls("Cbopur1").Tag = VBA.vbNullString
                Let VCboCurrent.Tag = VBA.vbNullString

              Case Else
                Select Case VCboCurrent.ListIndex

                  Case 0, 2
                    With .Controls("Cbopur3")
                      Let .Width = .Parent.Controls("Txtpur5").Width
                      Let .Left = .Parent.Controls("Txtpur6").Left
                      Call .Clear
                      Let .List = MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigpurrm_sys, eqlMdlArray, "ELEMENTO, MINERAL")
                      Let .ListIndex = 0
                    End With
                    With .Controls("Cbopur4")
                      Let .Width = .Parent.Controls("Txtpur5").Width
                      Let .Left = .Parent.Controls("Txtpur8").Left
                      Call .Clear
                      Let .List = MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalunits_sys, eqlMdlArray)
                      Let .ListIndex = 0
                    End With

                  Case 1
                    With .Controls("Txtpur6")
                      Let .Width = .Parent.Controls("Txtpur5").Width * 0.5
                      Let .ShowDropButtonWhen = fmShowDropButtonWhenNever
                    End With
                    With .Controls("Txtpur8")
                      Let .Width = .Parent.Controls("Txtpur5").Width * 0.5
                      Let .ShowDropButtonWhen = fmShowDropButtonWhenNever
                    End With
                    With .Controls("Cbopur3")
                      Let .Width = .Parent.Controls("Txtpur5").Width * 0.5
                      Let .Left = .Parent.Controls("Txtpur6").Left + .Parent.Controls("Txtpur6").Width
                      Call .Clear
                      Let .List = MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigpurrm_sys, eqlMdlArray, "ELEMENTO, MINERAL")
                      Call .AddItem("---", 0)
                      Call .AddItem("Quitar último", 1)
                      Call .AddItem("Limpiar lista", 2)
                      Let .ListIndex = 0
                    End With
                    With .Controls("Cbopur4")
                      Let .Width = .Parent.Controls("Txtpur5").Width * 0.5
                      Let .Left = .Parent.Controls("Txtpur8").Left + .Parent.Controls("Txtpur8").Width
                      Call .Clear
                      Let .List = MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalunits_sys, eqlMdlArray)
                      Call .AddItem("---", 0)
                      Call .AddItem("Quitar último", 1)
                      Call .AddItem("Limpiar lista", 2)
                      Let .ListIndex = 0
                    End With

                  Case Else: GoTo EH
                End Select

                With .Controls("Txtpur6")
                  Let .Visible = (VCboCurrent.ListIndex = 1)
                  Let .Locked = (VCboCurrent.ListIndex = 1)
                  Let .Value = VBA.vbNullString
                End With
                With .Controls("Txtpur8")
                  Let .Visible = (VCboCurrent.ListIndex = 1)
                  Let .Locked = (VCboCurrent.ListIndex = 1)
                  Let .Value = VBA.vbNullString
                End With
                Let .Controls("Txtpur7").Text = VBA.vbNullString
                Let .Controls("Txtpur8").Text = VBA.vbNullString
                Let .Controls("Txtpur9").Text = VBA.vbNullString

            End Select

          Case "Cbopur3", "Cbopur4" '__________________________________ Cbopur3, Cbopur4_
            If .Controls("Cbopur2").ListIndex = 1 Then
              Let StrAux0 = IIf(VCboCurrent.Name = "Cbopur3", "Txtpur6", "Txtpur8")
              Select Case VCboCurrent.ListIndex
                Case 0 ' Skip
                Case 1 ' Clear once
                  Let StrAux1 = VBA.CStr(.Controls(StrAux0).Value)
                  Let StrAux1 = VBA.StrReverse(StrAux1)
                  Let .Controls(StrAux0).Value = VBA.StrReverse(VBA.Mid(StrAux1, VBA.InStr(1, StrAux1, ";", vbTextCompare) + 1))
                Case 2 ' Clear all
                  Let .Controls(StrAux0).Value = VBA.vbNullString
                Case Else
                  Let StrAux1 = VBA.CStr(.Controls(StrAux0).Value)
                  If UBound(VBA.Split(StrAux1, ";")) + 1 = CLNG_BAL_LIMITCX Then GoTo EH
                  Let StrAux1 = StrAux1 & IIf(StrAux1 = VBA.vbNullString, VCboCurrent.Value, IIf(VBA.InStr(StrAux1, VCboCurrent.Value) = 0, ";" & VCboCurrent.Value, IIf(VCboCurrent.Name = "Cbopur3", VBA.vbNullString, ";" & VCboCurrent.Value)))
                  Let .Controls(StrAux0).Value = StrAux1
              End Select
            End If

          Case Else: GoTo EH
        End Select
      End With

    Case "concentrateinvoice" ' ----------------------------------------------- concentrateinvoice'
      'Call VBA.MsgBox(Frm.Tag)
      With FRM

        Set Lst0 = Nothing
        Set Chk0 = Nothing
        If VCboCurrent.Name = "Cbocnc1" Then Let .Controls("Cbocnc2").Tag = Empty
        If .Controls("Cbocnc2").Tag = Empty Then
          Let VecAux0 = Array("Cbocnc2", "Cbocnc3", "Txtcnc11", "Txtcnc12", "Txtcnc13", "Txtcnc14", "Chkcnc1", "Btncnc0", "Btncnc1", "Btncnc2", "Btncnc3", "Btncnc4", "Btncnc5")
          For i = LBound(VecAux0) To UBound(VecAux0)
            With .Controls(VBA.CStr(VecAux0(i)))
              Let .Visible = (i >= 7)
              If VBA.InStr(.Name, "Txt") > 0 Then
                Let .Value = VBA.vbNullString
                Let .ControlTipText = Empty
              End If
              Let .Enabled = (i < 7)
            End With
          Next i
          Let VecAux0 = Empty
          With .Controls("Lstcnc0")
            Let .RowSource = Empty
            Let .MultiSelect = fmMultiSelectSingle
            Let .ListStyle = fmListStylePlain
            Let .Visible = False
          End With
        End If

        Select Case VCboCurrent.Name

          Case "Cbocnc1"
            Set Cbo1 = Nothing
            Select Case VCboCurrent.ListIndex

              Case 0: Call View_assets_cncinvoice_contents(VCboCurrent, FRM) ' Contents
              Case 1: Call View_assets_cncinvoice_refination(VCboCurrent, FRM) ' Refination
              Case 2: Call View_assets_cncinvoice_penalties(VCboCurrent, FRM) ' Penalties (Items are added from config)
              Case 3: Call View_assets_cncinvoice_deductions(VCboCurrent, FRM) ' Deductions
              Case 4: Call View_assets_cncinvoice_bonus(VCboCurrent, FRM) ' Bonus
              Case Else: GoTo EH

            End Select

          Case "Cbocnc2"
            Let j = .Controls("Cbocnc1").ListIndex
            Select Case j

              Case 0
                Let .Controls("Txtcnc11").ControlTipText = "Ley de " & VCboCurrent.List(VCboCurrent.ListIndex, 1) & " en: " & VCboCurrent.List(VCboCurrent.ListIndex, 2)

              Case 1
                If VCboCurrent.ListCount > 0 Then
                  Let .Controls("Txtcnc11").Text = .Controls("Lstcnc0").List(VCboCurrent.ListIndex, 7)
                  Let .Controls("Txtcnc12").Text = .Controls("Lstcnc0").List(VCboCurrent.ListIndex, 8)
                End If

              Case 3
                Let VCboCurrent.Visible = True

                Select Case VCboCurrent.ListIndex

                  Case 0: Call View_assets_cncinvoice_dedexports(VCboCurrent, FRM) ' Exports
                  Case 1: Call View_assets_cncinvoice_dedtax(VCboCurrent, FRM) ' Law Taxes
                  Case 2: Call View_assets_cncinvoice_dedothertax(VCboCurrent, FRM) ' Other Taxes
                  Case 3: Call View_assets_cncinvoice_deddiscounts(VCboCurrent, FRM) ' Discounts others

                End Select

            End Select

          Case Else: GoTo EH
        End Select

        Let VCboCurrent.ControlTipText = VCboCurrent.Value
        Let .Controls("Lstcnc0").Visible = True
        Set Lst0 = .Controls("Lstcnc0")
        Set Chk0 = .Controls("Chkcnc1")

      End With

    Case "configlauncher" ' --------------------------------------------------- configlauncher'
      With FRM
        Select Case VCboCurrent.Name
          Case "Cbocon3"
            If VCboCurrent.Tag = "noevent" Or VCboCurrent.ListIndex < 0 Or VCboCurrent.ListCount <= 0 Then GoTo EH
            Let VCboCurrent.Tag = VCboCurrent.ListIndex

          Case "Cbocon4"
            With FRM
              Let VecAux0 = Array("Cbocon5", "Cbocon6", "Cbocon7", "Txtcon8", "Txtcon9", "Txtcon10", "Txtcon11", "Txtcon12", "Chkcon7")
              For i = LBound(VecAux0) To UBound(VecAux0)
                With .Controls(VBA.CStr(VecAux0(i)))
                  Let .Enabled = True
                  Let .Visible = False
                  If i > 2 And i < 8 Then Let .Value = Empty
                  If i = 8 Then Let .WordWrap = False
                End With
              Next i
              Let VecAux0 = Empty

              Set Lst1 = Nothing
              Let .Controls("Lstcon1").RowSource = Empty
              Select Case VCboCurrent.ListIndex

                Case 0: Call View_assets_config_rm(VCboCurrent, FRM) ' RM
                Case 1: Call View_assets_config_retentions(VCboCurrent, FRM) ' Retentions
                Case 2: Call View_assets_config_retentionsother(VCboCurrent, FRM) ' Retentions others
                Case 3: Call View_assets_config_penalties(VCboCurrent, FRM) ' Penalties
                Case 4: Call View_assets_config_exports(VCboCurrent, FRM) ' Exports

              End Select
              Let .Controls("Lstcon1").Enabled = True
              Let VCboCurrent.Enabled = True
              Set Lst1 = .Controls("Lstcon1")
            End With

        End Select
      End With

    Case Else: GoTo EH
  End Select

EH:
  Let VecAux0 = Empty
  Call RES.ErrorHandler.ErrorHandlerDisplay(sFN)

End Sub

' Events handlers - Listboxes
Private Sub View_Controls_listboxes_change(ByVal VLstCurrent As MSForms.ListBox, ByVal FRM As FRMBasics)

  Dim VecAux0 As Variant
  Dim BooAux0 As Boolean
  Dim i As Long, j As Long, k As Long, l As Long
  Dim EnuSht As EQLMDL_ENU_SHEETS, EnuTbl As EQLMDL_ENU_TABLES

  Const sFN As String = "VIEW::-Controls_listboxes_change"

  On Error GoTo EH
  #If Debugging Then
    Call Debug_assert(FRM Is Nothing, sFN)
  #End If
  If FRM Is Nothing Then GoTo EH
  If Not CMB Is Nothing Then
    If CMB.IsActive Then GoTo EH
  End If

  With FRM
    Select Case .Tag
      Case "oreinvoice"
        Select Case VLstCurrent.Name
          Case "Lstpur0"
            If VLstCurrent.Tag = "noevent" Or VLstCurrent.ListCount <= 0 Then GoTo EH
            Let i = MDL.LengthRows(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur)
            If i <= 0 Then GoTo EH
            If .Controls("Cbopur1").ListIndex = 3 Or .Controls("Cbopur1").ListIndex = 4 Then
              If .Controls("Chkpur1").Value Then ' Multiple providers

                Let j = .Controls("Cbopur1").ListIndex
                Let k = VLstCurrent.ListCount
                ReDim VecAux0(k - 1)
                For i = 0 To k - 1
                  Let VecAux0(i) = IIf(VLstCurrent.Selected(i), 1, 0)
                Next i
                Let VLstCurrent.RowSource = VBA.vbNullString
                Call MDL.MUpd(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur, Array(VBA.Join(VecAux0, ";")), IIf(j = 3, "RETENCIONES", "RETENCIONES OTRAS"), .Controls("Cbopur2").ListIndex + 1)
                For i = 1 To k
                  Call MDL.MUpd(eqlMdlSheetTmp, IIf(j = 3, eqlMdlTblConfigpurretentions_tmp, eqlMdlTblConfigpurretentionsother_tmp), Array(VecAux0(i - 1)), "SELECTED", i)
                Next i
                With VLstCurrent
                  Let .Tag = "noevent"
                  Let .RowSource = VBA.vbNullString ' Unlink list
                  Let .RowSource = MDL.MGet(eqlMdlSheetTmp, IIf(j = 3, eqlMdlTblConfigpurretentions_tmp, eqlMdlTblConfigpurretentionsother_tmp), eqlMdlListObject) ' Link list
                  If .ListCount > 0 Then ' Mark all previously selected items in the list
                    Let j = 0
                    For i = 0 To .ListCount - 1
                      Let .Selected(i) = VBA.CBool(.List(i, 2))
                      Let j = j + IIf(.Selected(i), 1, 0)
                    Next i
                  End If
                  Let .Tag = VBA.vbNullString
                End With
                Let .Controls("Chkpur2").Tag = "noevent"
                Let .Controls("Chkpur2").Value = (j = VLstCurrent.ListCount)
                Let .Controls("Chkpur2").Tag = VBA.vbNullString

              Else ' Just one provider

                Let j = .Controls("Cbopur1").ListIndex
                Let k = VLstCurrent.ListCount
                ReDim VecAux0(k - 1)
                For i = 0 To k - 1
                  Let VecAux0(i) = IIf(VLstCurrent.Selected(i), 1, 0)
                Next i
                Let VLstCurrent.RowSource = VBA.vbNullString
                For i = 1 To k
                  Call MDL.MUpd(eqlMdlSheetTmp, IIf(j = 3, eqlMdlTblConfigpurretentions_tmp, eqlMdlTblConfigpurretentionsother_tmp), Array(VecAux0(i - 1)), "SELECTED", i)
                Next i
                Call MDL.MUpd(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur, Array(VBA.Join(VecAux0, ";")), IIf(j = 3, "RETENCIONES", "RETENCIONES OTRAS"), .Controls("Cbopur2").ListIndex + 1)
                With VLstCurrent
                  Let .Tag = "noevent"
                  Let .RowSource = VBA.vbNullString ' Unlink list
                  Let .RowSource = MDL.MGet(eqlMdlSheetTmp, IIf(j = 3, eqlMdlTblConfigpurretentions_tmp, eqlMdlTblConfigpurretentionsother_tmp), eqlMdlListObject) ' Link list
                  If .ListCount > 0 Then ' Mark all previously selected items in the list
                    Let j = 0
                    For i = 0 To .ListCount - 1
                      Let .Selected(i) = VBA.CBool(.List(i, 2))
                      Let j = j + IIf(.Selected(i), 1, 0)
                    Next i
                  End If
                  Let .Tag = VBA.vbNullString
                End With
                Let .Controls("Chkpur2").Tag = "noevent"
                Let .Controls("Chkpur2").Value = (j = VLstCurrent.ListCount)
                Let .Controls("Chkpur2").Tag = VBA.vbNullString

              End If
            End If

          Case Else: GoTo EH
        End Select

      Case "concentrateinvoice"
        Select Case VLstCurrent.Name

          Case "Lstcnc0"
            ' Save current MenuCbo index
            Let i = .Controls("Cbocnc1").ListIndex
            If VLstCurrent.ListCount <= 0 Or VLstCurrent.Tag = "noevent" Then GoTo EH ' Empty list

            If i = 2 Or i = 3 Then ' Change list selections
              Let j = .Controls("Cbocnc2").ListIndex
              Let EnuSht = IIf(i = 2, eqlMdlSheetPur, IIf(j = 0, eqlMdlSheetPur, eqlMdlSheetTmp))
              Let EnuTbl = IIf(i = 2, eqlMdlTblpurchasecncpenalties_pur, IIf(j = 0, eqlMdlTblpurchasecncdedexp_pur, IIf(j = 1, eqlMdlTblConfigpurretentions_tmp, eqlMdlTblConfigpurretentionsother_tmp)))
              Let l = IIf(i = 2, 5, IIf(j = 0, 4, 2))
              Set Lst0 = Nothing
              With VLstCurrent
                Let k = .ListCount - 1
                Let j = .ListIndex + 1
                Let .RowSource = Empty ' Unlink list
                Call MDL.MUpd(EnuSht, EnuTbl, Array(IIf(.Selected(.ListIndex), 1, 0)), "SELECTED", j)
                ' Link list
                Let .RowSource = MDL.MGet(EnuSht, EnuTbl, eqlMdlStrTableName)
                ' Highlight selected items (Set tag=empty)
                Let i = 0
                For j = 0 To k
                  Let .Selected(j) = VBA.CBool(.List(j, l))
                  Let i = i + IIf(.Selected(j), 1, 0)
                Next j
                Set Chk0 = Nothing
                Let .Parent.Controls("Chkcnc1").Value = (i = .ListCount)
                Set Chk0 = .Parent.Controls("Chkcnc1")
              End With
              Set Lst0 = .Controls("Lstcnc0")
            End If

        End Select

      Case "configlauncher"
        If Not (VLstCurrent.ListCount > 0) Then GoTo EH
        Select Case VLstCurrent.Name
          ' DIVISA
          Case "Lstcon0"
            Let .Controls("Btncon3").Enabled = (VLstCurrent.ListIndex > 1)

          ' LIQUIDACIONES
          Case "Lstcon1"
            Let k = VLstCurrent.ListIndex
            Let j = .Controls("Cbocon4").ListIndex
            Let i = IIf(j = 0, 12, IIf(j = 1, 3, IIf(j = 2, 4, IIf(j = 3, 6, 0))))
            Let BooAux0 = IIf(j = 0, k <= 50 And k >= i, (k <= GLOSNG_PUR_LIMIT And k >= i))
            Let .Controls("Btncon13").Enabled = BooAux0

          ' ELEMENTOS
          Case "Lstcon2"
            Let .Controls("Btncon15").Enabled = (VLstCurrent.ListIndex >= 60)
            Let .Controls("Btncon17").Enabled = (VLstCurrent.ListIndex >= 60)

          Case Else: GoTo EH
        End Select

      Case Else: GoTo EH
    End Select
  End With

EH:
  Call RES.ErrorHandler.ErrorHandlerDisplay(sFN)

End Sub

Private Sub View_controls_listboxes_dblclick(ByVal VLstCurrent As MSForms.ListBox, ByVal FRM As FRMBasics)

  Dim VecAux0 As Variant
  Dim StrAux0 As String
  Dim BooAux0 As Boolean
  Dim i As Long, j As Long, k As Long, l As Long

  Const sFN As String = "VIEW::-Controls_listboxes_dblclick"

  On Error GoTo EH
  ' Cancel all actions while CMB is active on view
  If Not CMB Is Nothing Then
    If CMB.IsActive Then GoTo EH
  End If

  ' Process actions
  With FRM
    Select Case .Tag

      Case "oreinvoice" '-------------------------------------------------------------------- oreinvoice'
        If Not (VLstCurrent.ListCount > 0 And VLstCurrent.ListIndex > -1) Then GoTo EH
        If RES.REGEX.isEmptyStringReg(VBA.CStr(VLstCurrent.List(VLstCurrent.ListIndex, 0))) Then GoTo EH
        Let l = .Controls("Cbopur1").ListIndex
        Select Case .Controls("Cbopur1").ListIndex

          Case 0 ' Base Data
            For i = 5 To 9
              If i = 6 Then Let .Controls("Txtpur" & i).Text = VBA.Format(VBA.CDate(VLstCurrent.List(VLstCurrent.ListIndex, i - 5)), "dd/mm/yy")
              If Not i = 6 Then Let .Controls("Txtpur" & i).Text = VLstCurrent.List(VLstCurrent.ListIndex, i - 5)
            Next i

          Case 1 ' Main Data
            If Not VBA.CBool(VLstCurrent.List(VLstCurrent.ListIndex, 14)) Then GoTo EH
            Let StrAux0 = VBA.CStr(VLstCurrent.List(VLstCurrent.ListIndex, 5))
            Let BooAux0 = RES.REGEX.isMineralComplex(StrAux0)
            If BooAux0 Then ' Mineral type
              Let .Controls("Cbopur2").ListIndex = 1
            ElseIf RES.REGEX.isMineralBrosa(StrAux0) Then
              Let .Controls("Cbopur2").ListIndex = 2
            Else
              Let .Controls("Cbopur2").ListIndex = 0
            End If
            If BooAux0 Then ' Case complex
              For i = 6 To 9
                Let .Controls("Txtpur" & i).Value = VLstCurrent.List(VLstCurrent.ListIndex, i)
              Next i
            Else ' Minerals or brosas
              For i = 0 To .Controls("Cbopur3").ListCount - 1
                If VBA.LCase(VBA.CStr(VLstCurrent.List(VLstCurrent.ListIndex, 6))) = VBA.LCase(VBA.CStr(.Controls("Cbopur3").List(i, 0))) Then Exit For
              Next i
              Let .Controls("Cbopur3").ListIndex = i
              Let .Controls("Txtpur7").Value = VBA.CStr(VLstCurrent.List(VLstCurrent.ListIndex, 7))
              For i = 0 To .Controls("Cbopur4").ListCount - 1
                If VBA.LCase(VBA.CStr(VLstCurrent.List(VLstCurrent.ListIndex, 8))) = VBA.LCase(VBA.CStr(.Controls("Cbopur4").List(i, 0))) Then Exit For
              Next i
              Let .Controls("Cbopur4").ListIndex = i
              Let .Controls("Txtpur9").Value = VBA.CStr(VLstCurrent.List(VLstCurrent.ListIndex, 9))
            End If

          Case 2, 5 ' Bonus, Discounts
            Let .Controls("Txtpur5").Value = VBA.CStr(VLstCurrent.List(VLstCurrent.ListIndex, 1))
            Let .Controls("Txtpur6").Value = VBA.CStr(VLstCurrent.List(VLstCurrent.ListIndex, 2))

          Case Else: GoTo EH
        End Select
        For i = 0 To 5
          Let .Controls("Btnpur" & i).Enabled = (i = 1 Or i = 2)
        Next i
        Let .Controls("Chkpur1").Enabled = False
        Let .Controls("Cbopur1").Enabled = False
        Let .Controls("Lstpur0").Enabled = False
        Let .Controls("Btnmain1").Enabled = False




      Case "concentrateinvoice" '-------------------------------------------------------------------- concentrateinvoice'
        If VLstCurrent.ListCount <= 0 And VLstCurrent.ListIndex < 0 Then GoTo EH
        Let l = .Controls("Cbocnc1").ListIndex
        Select Case l

          Case 0 ' Contents
            For i = 0 To .Controls("Cbocnc2").ListCount - 1
              If VBA.LCase(VLstCurrent.List(VLstCurrent.ListIndex, 0)) = VBA.LCase(.Controls("Cbocnc2").List(i, 0)) Then Let .Controls("Cbocnc2").ListIndex = i: Exit For
            Next i
            Let StrAux0 = VLstCurrent.List(VLstCurrent.ListIndex, 1)
            Let StrAux0 = VBA.Replace(StrAux0, "[" & VBA.CStr(VLstCurrent.List(VLstCurrent.ListIndex, 5)) & "]", VBA.vbNullString)
            Let .Controls("Txtcnc11").Value = StrAux0
            Let .Controls("Txtcnc12").Value = VLstCurrent.List(VLstCurrent.ListIndex, 2)
            Let .Controls("Txtcnc14").Value = VLstCurrent.List(VLstCurrent.ListIndex, 4)
            Let .Controls("Cbocnc3").ListIndex = IIf(VBA.LCase(VLstCurrent.List(VLstCurrent.ListIndex, 4)) = "y", 0, 1)

          Case 1 ' Refination
            For i = 0 To VLstCurrent.ListCount - 1
              If VBA.LCase(VLstCurrent.List(VLstCurrent.ListIndex, 6)) = VBA.LCase(VLstCurrent.List(i, 6)) Then Let .Controls("Cbocnc2").ListIndex = i: Exit For
            Next i
            Let .Controls("Txtcnc11").Value = VLstCurrent.List(VLstCurrent.ListIndex, 7)
            Let .Controls("Txtcnc12").Value = VLstCurrent.List(VLstCurrent.ListIndex, 8)

          Case 2 ' Penalties
            Let j = (.Controls("Framecnc2").InsideWidth - (.InsideWidth * 0.05 * 2)) / 5 ' Left for all main controls
            With .Controls("Chkcnc1")
              Let .Caption = VLstCurrent.List(VLstCurrent.ListIndex, 0)
              Let .Width = j
              Let .Left = .Parent.Controls("Cbocnc1").Left
              Let .Enabled = False
            End With
            For i = 1 To 4
              With .Controls("Txtcnc" & i + 10)
                Let .Value = VLstCurrent.List(VLstCurrent.ListIndex, i)
                Let .Width = j
                Let .Left = .Parent.Controls("Chkcnc1").Left + (.Width * i)
                Let .Visible = True
              End With
            Next i

          Case 3 ' Deductions
            Select Case .Controls("Cbocnc2").ListIndex
              Case 0 ' Exports
                Set Chk0 = Nothing
                ' Set controls layout
                Let j = (.Controls("Framecnc2").InsideWidth - (.InsideWidth * 0.05 * 2)) / 4 ' Left for all main controls
                Let VecAux0 = Array("Txtcnc11", "Txtcnc12", "Txtcnc13", "Chkcnc1")
                For i = 0 To 3
                  Let .Controls(VBA.CStr(VecAux0(i))).Visible = True
                  Let .Controls(VBA.CStr(VecAux0(i))).Value = VLstCurrent.List(VLstCurrent.ListIndex, i)
                  Let .Controls(VBA.CStr(VecAux0(i))).Width = j
                  Let .Controls(VBA.CStr(VecAux0(i))).Left = .Controls("Cbocnc1").Left + (.Controls(VBA.CStr(VecAux0(i))).Width * i)
                Next i
                Let .Controls("Chkcnc1").Caption = "Valor directo"
                Let .Controls("Cbocnc2").Visible = False
              Case 3 ' Discounts
                Let .Controls("Txtcnc11").Value = VLstCurrent.List(VLstCurrent.ListIndex, 1)
                Let .Controls("Txtcnc12").Value = VLstCurrent.List(VLstCurrent.ListIndex, 2)
              Case Else: GoTo EH
            End Select

          Case 4 ' Bonus
            Let .Controls("Txtcnc11").Value = VLstCurrent.List(VLstCurrent.ListIndex, 1)
            Let .Controls("Txtcnc12").Value = VLstCurrent.List(VLstCurrent.ListIndex, 2)

          Case Else: GoTo EH
        End Select
        For i = 0 To 5
          Let .Controls("Btncnc" & i).Enabled = (i = 1 Or i = 2)
        Next i
        Let .Controls("Cbocnc1").Enabled = False
        Let .Controls("Lstcnc0").Enabled = False
        Let .Controls("Btnmain1").Enabled = False



      Case "configlauncher"
        If Not (VLstCurrent.ListCount > 0) Then GoTo EH
        Select Case VLstCurrent.Name
          ' DIVISA
          Case "Lstcon0"
            If VLstCurrent.ListIndex > 1 Then
              For i = 0 To 3
                Let .Controls("Btncon" & i).Enabled = (i = 1 Or i = 2)
                Let FRM.Controls("Multicon").Pages(i + 1).Enabled = False
                If i < 3 Then Let FRM.Controls("Txtcon" & i).Text = VLstCurrent.List(VLstCurrent.ListIndex, i)
              Next i
              Let VLstCurrent.Enabled = False
            End If

          ' LIQUIDACIONES
          Case "Lstcon1"
            Let k = VLstCurrent.ListIndex
            Let j = .Controls("Cbocon4").ListIndex
            Let i = IIf(j = 0, 11, IIf(j = 1, 3, IIf(j = 2, 4, IIf(j = 3, 6, 0))))
            Select Case j

              Case 0 ' RM
                Let .Controls("Cbocon5").Value = VLstCurrent.List(VLstCurrent.ListIndex, 0)
                Let .Controls("Cbocon6").Value = VLstCurrent.List(VLstCurrent.ListIndex, 2)
                Let .Controls("Cbocon7").Value = VLstCurrent.List(VLstCurrent.ListIndex, 6)
                Let .Controls("Cbocon5").Enabled = (k >= i)
                Let .Controls("Cbocon6").Enabled = (k >= i)
                Let .Controls("Cbocon7").Enabled = (k >= i)
                Let .Controls("Txtcon8").Value = VLstCurrent.List(VLstCurrent.ListIndex, 3)
                Let .Controls("Txtcon9").Value = VLstCurrent.List(VLstCurrent.ListIndex, 4)
                Let .Controls("Txtcon10").Value = VLstCurrent.List(VLstCurrent.ListIndex, 5)

              Case 1, 2 ' Retentions
                Let .Controls("Txtcon8").Value = VLstCurrent.List(VLstCurrent.ListIndex, 0)
                Let .Controls("Txtcon9").Value = VLstCurrent.List(VLstCurrent.ListIndex, 1)
                Let .Controls("Txtcon8").Enabled = (k >= i)

              Case 3 ' Penalties
                Let .Controls("Cbocon5").Value = VLstCurrent.List(VLstCurrent.ListIndex, 0)
                Let .Controls("Cbocon5").Enabled = (k >= i)
                Let .Controls("Txtcon8").Value = VLstCurrent.List(VLstCurrent.ListIndex, 1)
                Let .Controls("Txtcon9").Value = VLstCurrent.List(VLstCurrent.ListIndex, 2)
                Let .Controls("Txtcon10").Value = VLstCurrent.List(VLstCurrent.ListIndex, 3)
                Let .Controls("Txtcon11").Value = VLstCurrent.List(VLstCurrent.ListIndex, 4)

              Case 4
                Let .Controls("Txtcon8").Value = VLstCurrent.List(VLstCurrent.ListIndex, 0)
                Let .Controls("Txtcon9").Value = VLstCurrent.List(VLstCurrent.ListIndex, 1)
                Let .Controls("Txtcon10").Value = VLstCurrent.List(VLstCurrent.ListIndex, 2)

              Case Else: GoTo EH
            End Select
            Call RES.FrmSelTextControl(.Controls(IIf((k < i) And (j = 1 Or j = 2), "Txtcon9", "Txtcon8")))
            For i = 0 To 4
              If i < 4 Then Let .Controls("Btncon" & i + 10).Enabled = (i = 1 Or i = 2)
              Let FRM.Controls("Multicon").Pages(i).Enabled = (i = 3)
            Next i
            Let VLstCurrent.Enabled = False
            Let .Controls("Cbocon4").Enabled = False

          ' ELEMENTOS
          Case "Lstcon2"
            For i = 0 To 4
              If i < 3 Then Let .Controls("Btncon" & i + 14).Enabled = (i = 1 Or i = 2)
              Let .Controls("Multicon").Pages(i).Enabled = (i = 4)
              If i < 2 Then
                Let .Controls("Txtcon" & i + 13).Text = VLstCurrent.List(VLstCurrent.ListIndex, i)
                Let .Controls("Txtcon" & i + 13).Enabled = (VLstCurrent.ListIndex > 60)
              End If
            Next i
            Let .Controls("Chkcon15").Value = (VLstCurrent.List(VLstCurrent.ListIndex, 2) = "SI")
            If (VLstCurrent.ListIndex > 60) Then
              Call RES.FrmSelTextControl(.Controls("Txtcon13"))
            Else
              Call .Controls("Chkcon15").SetFocus
            End If
            Let VLstCurrent.Enabled = False

          Case Else: GoTo EH
        End Select

      Case Else: GoTo EH
    End Select
  End With

EH:
  Let VecAux0 = Empty
  Call RES.ErrorHandler.ErrorHandlerDisplay(sFN)

End Sub

' Events handlers - Texts
Private Sub View_controls_texts_dropbuttonclick(ByVal VTxtCurrent As MSForms.TextBox, ByVal FRM As FRMBasics)

  Dim FD As FileDialog
  Dim StrAux0 As String
  Dim i As Long

  Const sFN As String = "VIEW::-TextEvents_dropbuttonclick"

  On Error GoTo EH
  If Not CMB Is Nothing Then
    If CMB.IsActive Then GoTo EH
  End If

  Select Case FRM.Tag
    Case "regression" ' ================================================================== Regression'
      Select Case VTxtCurrent.Name
        Case "Txtreg0": Call View_assets_capture_data_from_ranges(FRM, VTxtCurrent, "Seleccione datos numéricos para la variable X:", VBooZeroes:=True, RDblTop:=CDBL_MAX)
        Case "Txtreg1": Call View_assets_capture_data_from_ranges(FRM, VTxtCurrent, "Seleccione datos numéricos para la variable Y:", VBooZeroes:=True, RDblTop:=CDBL_MAX)
        Case "Txtreg2": Call View_assets_capture_data_from_ranges(FRM, VTxtCurrent, "Seleccione datos numéricos a Extrapolar:", VBooZeroes:=True, RDblTop:=CDBL_MAX)
        Case "Txtreg3": If FRM.Controls("Optreg3").Value Then Call View_assets_capture_data_from_ranges(FRM, VTxtCurrent, "Seleccione una celda donde imprimir el resultado:", eqlResRanges, VBooRngAddressToTxt:=True)
      End Select

    Case "balmet", "balmeteco" ' ======================================================== Balmet'
      Call View_assets_read_data_for_balmet(VTxtCurrent, FRM)

    Case "projection" ' ================================================================== Projection'
      Select Case VTxtCurrent.Name
        Case "Txtpro0": Call View_assets_capture_data_from_ranges(FRM, VTxtCurrent, FRM.Controls("Lblpro0").Caption & " seleccione uno o varios datos numéricos:", eqlResNumbers, True, VBooSum:=True, RDblTop:=CDBL_MAX)
        Case "Txtpro1"
          With VTxtCurrent
            Call View_assets_capture_data_from_ranges(FRM, VTxtCurrent, FRM.Controls("Lblpro1").Caption & " seleccione uno o varios datos numéricos:", eqlResNumbers, True, RDblTop:=CDBL_MAX)
            If Not RES.REGEX.isNumberEntireAndDecimalVectorColonSeparated(.Text) Then GoTo EH
            Let i = UBound(VBA.Split(.Text, ";"))
            If (i - 1) > CLNG_BAL_LIMIT Then Call VBA.MsgBox("Puede proyectar proyectos con hasta " & CLNG_BAL_LIMIT & " productos minerales.", vbInformation, STRAPPTITLE): Let .Text = VBA.vbNullString: GoTo EH
            Let FRM.Controls("Txtpro5").Text = RES.RESCORE.BalGetGenericNames(i)
            Let FRM.Controls("Txtpro6").Text = RES.RESCORE.BalGetGenericUnits(i, BBooTFs:=True)
          End With

        Case "Txtpro2": Call View_assets_capture_data_from_ranges(FRM, VTxtCurrent, FRM.Controls("Lblpro2").Caption & " seleccione uno o distintos rangos con datos numéricos:", eqlResNumbers, True, VBooSlashGrouped:=True, RDblTop:=CDBL_MAX)
        Case "Txtpro3": Call View_assets_capture_data_from_ranges(FRM, VTxtCurrent, FRM.Controls("Lblpro3").Caption & " seleccione uno o varios datos numéricos:", eqlResNumbers, True, VBooSum:=True, RDblTop:=CDBL_MAX)
        Case "Txtpro9": If FRM.Controls("Optpro3").Value Then Call View_assets_capture_data_from_ranges(FRM, VTxtCurrent, "Seleccione una celda donde imprimir el resultado:", eqlResRanges, VBooRngAddressToTxt:=True)
      End Select

    Case "oreinvoice" ' ================================================================ Oreinvoice'
      With FRM
        Select Case VTxtCurrent.Name
          Case "Txtpur0": Call View_assets_capture_data_from_ranges(FRM, VTxtCurrent, .Controls("Lblpur0").Caption, eqlResNotNumbers)
          Case "Txtpur1": Call DPR.DatePicker(FRM, DControl:=.Controls("Txtpur1"), DAlign:=dpcRight)
          Case "Txtpur2": Call View_assets_capture_data_from_ranges(FRM, VTxtCurrent, .Controls("Lblpur2").Caption & " seleccione uno o varios datos numéricos:", eqlResNumbers, True, VBooSum:=True, RDblTop:=CDBL_MAX)
          Case "Txtpur3", "Txtpur4" 'Names & Type
            Call View_assets_capture_data_from_ranges(FRM, VTxtCurrent, .Controls(IIf(VTxtCurrent.Name = "Txtpur3", "Lblpur3", "Lblpur4")).Caption, eqlResNotNumbers)
          Case "Txtpur5"
            Call View_assets_capture_data_from_ranges(FRM, VTxtCurrent, .Controls("Cbopur1").Value & ": seleccione un dato textual:", eqlResNotNumbers)
            If Not VTxtCurrent.Value = VBA.vbNullString Then
              If VBA.InStr(VTxtCurrent.Value, ";") > 0 Then Let VTxtCurrent.Value = VBA.Mid(VTxtCurrent.Value, 1, VBA.InStr(VTxtCurrent.Value, ";") - 1)
            End If
          Case "Txtpur6"
            If .Controls("Cbopur1").ListIndex = 2 Or .Controls("Cbopur1").ListIndex = 5 Then
              Call View_assets_capture_data_from_ranges(FRM, VTxtCurrent, .Controls("Cbopur1").Value & ": seleccione uno o varios datos numéricos:", eqlResNumbers, True, VBooSum:=True, RDblTop:=CDBL_MAX)
            Else
              Call DPR.DatePicker(FRM, DControl:=.Controls("Txtpur6"), DAlign:=dpcRight)
            End If
          Case "Txtpur7", "Txtpur8", "Txtpur9": Call View_assets_capture_data_from_ranges(FRM, VTxtCurrent, .Controls("Cbopur1").Value & ": seleccione uno o varios datos numéricos:", eqlResNumbers, True, VBooSum:=(Not .Controls("Cbopur2").ListIndex = 1), RDblTop:=CDBL_MAX)
          Case "Txtpur10": If .Controls("Optpur1").Value Then Call View_assets_capture_data_from_ranges(FRM, VTxtCurrent, "Seleccione una celda donde imprimir el resultado:", eqlResRanges, VBooRngAddressToTxt:=True)
        End Select
      End With

    Case "concentrateinvoice" ' ======================================================== concentrateinvoice'
      With FRM
        Select Case VTxtCurrent.Name

          Case "Txtcnc0", "Txtcnc3": Call View_assets_capture_data_from_ranges(FRM, VTxtCurrent, .Controls(IIf(VTxtCurrent.Name = "Txtcnc0", "Lblcnc0", "Lblcnc3")).Caption, eqlResNotNumbers, VBooJustOneWord:=True)
          Case "Txtcnc1", "Txtcnc4": Call DPR.DatePicker(FRM, DControl:=VTxtCurrent, DAlign:=dpcRight)

          Case "Txtcnc2", "Txtcnc5", "Txtcnc6", "Txtcnc8", "Txtcnc10"
            Let StrAux0 = "Lblcnc" & IIf(VTxtCurrent.Name = "Txtcnc2", 2, IIf(VTxtCurrent.Name = "Txtcnc5", 6, IIf(VTxtCurrent.Name = "Txtcnc6", 7, IIf(VTxtCurrent.Name = "Txtcnc8", 9, "11"))))
            Call View_assets_capture_data_from_ranges(FRM, VTxtCurrent, .Controls(StrAux0).Caption, eqlResNumbers, True, VBooSum:=True, RDblTop:=CDBL_MAX)

          Case "Txtcnc7", "Txtcnc9": Call View_assets_capture_data_from_ranges(FRM, VTxtCurrent, .Controls("Lblcnc" & IIf(VTxtCurrent.Name = "Txtcnc7", 8, 10)).Caption, eqlResNumbers, True, True, VBooSum:=True, RDblTop:=100)
          Case "Txtcnc12", "Txtcnc13", "Txtcnc14": Call View_assets_capture_data_from_ranges(FRM, VTxtCurrent, .Controls("Cbocnc1").Value & ": Valor numérico", eqlResNumbers, True, VBooSum:=True, RDblTop:=CDBL_MAX)

          Case "Txtcnc11"
            If .Controls("Cbocnc1").ListIndex = 3 And .Controls("Cbocnc2").ListIndex = 3 Then
              Call View_assets_capture_data_from_ranges(FRM, VTxtCurrent, .Controls("Cbocnc2").Value & ": Nombre", eqlResNotNumbers, VBooJustOneWord:=True)
            ElseIf .Controls("Cbocnc1").ListIndex = 4 Then
              Call View_assets_capture_data_from_ranges(FRM, VTxtCurrent, .Controls("Cbocnc1").Value & ": Nombre", eqlResNotNumbers, VBooJustOneWord:=True)
            Else
              Call View_assets_capture_data_from_ranges(FRM, VTxtCurrent, .Controls("Cbocnc1").Value & ": Valor numérico", eqlResNumbers, True, VBooSum:=True, RDblTop:=CDBL_MAX)
            End If

          Case "Txtcnc15": If .Controls("Optcnc1").Value Then Call View_assets_capture_data_from_ranges(FRM, VTxtCurrent, "Seleccione una celda donde imprimir el resultado:", eqlResRanges, VBooRngAddressToTxt:=True)
        End Select
      End With

    Case "configlauncher" ' ============================================================ configlauncher'
      Select Case VTxtCurrent.Name

        Case "Txtcon3"
          Set FD = Application.FileDialog(msoFileDialogFilePicker)
          With FD
            Let .AllowMultiSelect = False
            Let .ButtonName = "Cargar imagen"
            Let .Title = "Cargar imagen para Encabezado de página - " & STRAPPTITLE
            Call .Filters.Add("Archivos de imagen", "*.jpg; *.png; *.gif; *.bmp; *.tiff", 1)
            Let i = .Show()
            If i = -1 Then Let VTxtCurrent.Value = VBA.CStr(.SelectedItems.Item(1))
          End With

        Case Else: GoTo EH
      End Select

    Case Else: GoTo EH
  End Select

EH:
  Set FD = Nothing
  Call RES.ErrorHandler.ErrorHandlerDisplay(sFN)

End Sub

' Events handlers - Spinbuttons
Private Sub View_controls_spinbuttons_change( _
  ByVal VSpnCurrent As MSForms.SpinButton, _
  ByVal FRM As FRMBasics, _
  Optional BooUp As Boolean = False)

  Dim StrCtrlName As String, StrAux0 As String
  Dim BooDecs As Boolean, BooPercent As Boolean
  Dim DblValue0 As Double, DblValue1 As Double

  Const sFN As String = "VIEW::-spinbuttons_change"

  On Error GoTo EH
  If Not CMB Is Nothing Then
    If CMB.IsActive Then GoTo EH
  End If

  Let BooPercent = True
  With FRM

    Select Case .Tag

      Case "projection"
        Select Case VSpnCurrent.Name
          Case "Spnpro0", "Spnpro1"
            Let StrCtrlName = "Txtpro4"
            Let BooDecs = (VSpnCurrent.Name = "Spnpro1")
          Case "Spnpro2"
            Let StrCtrlName = "Txtpro7"
            Let BooPercent = False
          Case "Spnpro3", "Spnpro4"
            Let StrCtrlName = "Txtpro8"
            Let BooDecs = (VSpnCurrent.Name = "Spnpro4")
          Case Else: GoTo EH
        End Select

      Case "concentrateinvoice"
        Select Case VSpnCurrent.Name
          Case "Spncnc0", "Spncnc1"
            Let StrCtrlName = "Txtcnc7"
            Let BooDecs = (VSpnCurrent.Name = "Spncnc1")
          Case "Spncnc2", "Spncnc3"
            Let StrCtrlName = "Txtcnc9"
            Let BooDecs = (VSpnCurrent.Name = "Spncnc2")
          Case Else: GoTo EH
        End Select

      Case Else: GoTo EH

    End Select

    Let StrAux0 = .Controls(StrCtrlName).Text
    If RES.REGEX.isPercentReg(StrAux0) Then
      Let DblValue0 = VBA.CDbl(VBA.Replace(VBA.Replace(.Controls(StrCtrlName).Text, "%", VBA.vbNullString), ".", ","))
      Let DblValue1 = IIf(BooDecs, 0.01, 1)
      Let DblValue0 = DblValue0 + IIf(BooUp, DblValue1, -DblValue1)
      If DblValue0 > 100 Then Let DblValue0 = 100
      If DblValue0 < 0 Then Let DblValue0 = 0
    End If
    Let .Controls(StrCtrlName).Value = VBA.Format(DblValue0, IIf(BooPercent, "0.00""%""", "0"))

  End With

EH:
  Call RES.ErrorHandler.ErrorHandlerDisplay(sFN)

End Sub



' CORE VIEWS METHODS LIST ---------------------------------------------------------- '
' Validations
Private Function View_core_validations( _
  ByVal VCmdBtn As MSForms.CommandButton, _
  ByVal FRM As FRMBasics, _
  Optional VFBooSelections As Boolean = False, _
  Optional ControlId As String, _
  Optional VRngBox As Range = Nothing) As Boolean

  Const sFN As String = "VIEW::-Validations"

  On Error GoTo EH
  Let View_core_validations = False

  If VFBooSelections Then  ' ================================================= Validations for Selections contents

    Select Case ControlId
      Case "regressionselection", "regressionselection-check_cell": Let View_core_validations = View_core_validations_regression(Nothing, Nothing, VFBooSelections, ControlId, VRngBox)
      Case "balmetselection", "balmetselection-check_cell": Let View_core_validations = View_core_validations_balmet(Nothing, Nothing, VFBooSelections, ControlId, VRngBox)
    End Select

  Else ' ===================================================================== Validation to Forms contents

    If Not VCmdBtn.Tag = "validate" Then Let View_core_validations = True: GoTo EH ' Just pass enabled for valitations buttons

    Select Case FRM.Tag
      Case "regression": Let View_core_validations = View_core_validations_regression(VCmdBtn, FRM, VFBooSelections, ControlId)
      Case "balmet", "balmeteco": Let View_core_validations = View_core_validations_balmet(VCmdBtn, FRM, VFBooSelections, ControlId)
      Case "projection": Let View_core_validations = View_Core_validations_projections(VCmdBtn, FRM)
      Case "oreinvoice": Let View_core_validations = View_Core_validations_oreinvoice(VCmdBtn, FRM)
      Case "concentrateinvoice": Let View_core_validations = View_Core_validations_concentrateinvoice(VCmdBtn, FRM)
      Case "configlauncher": Let View_core_validations = View_core_validations_config(VCmdBtn, FRM)
      Case Else: GoTo EH
    End Select

  End If

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function View_core_validations_regression( _
  ByVal VCmdBtn As MSForms.CommandButton, _
  ByVal FRM As FRMBasics, _
  Optional VFBooSelections As Boolean = False, _
  Optional ControlId As String, _
  Optional VRngBox As Range = Nothing) As Boolean

  Dim RngBox As Range
  Dim VecAux0 As Variant
  Dim StrMsg0 As String, StrAux0 As String
  Dim i As Long, j As Long

  Const sFN As String = "VIEW::-validations_regression"

  On Error GoTo EH
  Let View_core_validations_regression = False

  If VFBooSelections Then ' ---------------------------------------- Validations for Selections contents

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
      Let View_core_validations_regression = True

    Else ' "regressionselection-check_cell"

      'Check for spreadsheet edges and data inside the region selected
      If VRngBox Is Nothing Then GoTo EH
      ' Check for protected data
      With VRngBox
        If GLOBOO_SELECTIONCELL Then
          If .Parent.Parent.ProtectStructure Then Call VBA.MsgBox("No se puede crear una nueva hoja en el libro actual porque está protegido. Quite la protección o seleccione otro libro para continuar.", vbInformation, STRAPPTITLE): GoTo EH
        Else
          If RES.ExcelIsProtectedSheet(VRngBox.Parent) Then Call VBA.MsgBox("La celda que seleccionó para imprimir los resultados está en una hoja protegida. Para realizar un cambio, debe quitar la protección de la hoja. Es posible que se le pida que escriba una contraseña.", vbInformation, STRAPPTITLE): GoTo EH
        End If
      End With
      If Not View_assets_check_for_spreadsheet_edges(VRngBox, Application.Selection.Rows.Count + 3, 20, StrMsg0, VBooSelections:=True) Then Call VBA.MsgBox(StrMsg0, vbInformation, STRAPPTITLE): GoTo EH
      Let View_core_validations_regression = (VBA.MsgBox(STRAPPTITLE & " calculará la regresión lineal con los datos que seleccionó." & VBA.vbNewLine & "¿Proceder?", vbQuestion + vbYesNo, STRAPPTITLE) = vbYes)

    End If

  Else ' ----------------------------------------------------------- Validation to Forms contents

    With FRM
      Select Case VCmdBtn.Name
        Case "Btnmain1"
          ' Y
          Let VecAux0 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(.Controls("Txtreg1").Value, i, RBooIncludeZeroes:=True, RDblTop:=CDBL_MAX)
          If i <= 0 Then Call CMB.CMsgBox("¡Se esperaban valores base para la regresión en Y!", STRAPPTITLE, "Error", cMsgType:=cmbError, cMsgButtons:=cmbClose, cUFrm:=FRM, cCtrlTargetName:="Txtreg1", cCtrlSourceName:=VCmdBtn.Name, cCtrlRequired:=cmbRequired, cPositionX:=cmbRight, cPositionY:=cmbBottom, cAspect:=cmbTagLine): GoTo EH
          If Application.WorksheetFunction.Sum(VecAux0) = 0 Then Call CMB.CMsgBox("¡Se esperaben valores no nulos para la regresión en Y!", STRAPPTITLE, "Error", cMsgType:=cmbError, cMsgButtons:=cmbClose, cUFrm:=FRM, cCtrlTargetName:="Txtreg1", cCtrlSourceName:=VCmdBtn.Name, cCtrlRequired:=cmbRequired, cPositionX:=cmbRight, cPositionY:=cmbBottom, cAspect:=cmbTagLine): GoTo EH
          Let VecAux0 = Empty
          ' X
          Let VecAux0 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(.Controls("Txtreg0").Value, j, RBooIncludeZeroes:=True, RDblTop:=CDBL_MAX)
          If j > 0 Then
            If Not j = i Then Call CMB.CMsgBox("¡Se requieren mismas cantidades de valores en X e Y!", STRAPPTITLE, "Error", cMsgType:=cmbError, cMsgButtons:=cmbClose, cUFrm:=FRM, cCtrlTargetName:="Txtreg0", cCtrlSourceName:=VCmdBtn.Name, cCtrlRequired:=cmbRequired, cPositionX:=cmbRight, cPositionY:=cmbBottom, cAspect:=cmbTagLine): GoTo EH
            If Application.WorksheetFunction.Sum(VecAux0) = 0 Then Call CMB.CMsgBox("¡Se esperaben valores en X no nulos!", STRAPPTITLE, "Error", cMsgType:=cmbError, cMsgButtons:=cmbClose, cUFrm:=FRM, cCtrlTargetName:="Txtreg0", cCtrlSourceName:=VCmdBtn.Name, cCtrlRequired:=cmbRequired, cPositionX:=cmbRight, cPositionY:=cmbBottom, cAspect:=cmbTagLine): GoTo EH
          End If
          Let VecAux0 = Empty

          ' Check for protected data
          ' Check for spreadsheet edges and data inside the region selected
          If .Controls("Optreg3").Value Then
            If RES.ExcelIsProtectedSheet(GLORNG_CELL.Parent) Then Call RES.FrmCMsgbox("La celda que seleccionó para imprimir los resultados está en una hoja protegida. Para realizar un cambio, debe quitar la protección de la hoja. Es posible que se le pida que escriba una contraseña.", STRAPPTITLE, vbInformation, True, .Controls("Txtreg3")): GoTo EH
            If Not View_assets_check_for_spreadsheet_edges(GLORNG_CELL, i * 5, 20, StrMsg0) Then Call RES.FrmCMsgbox(StrMsg0, STRAPPTITLE, vbInformation, True, .Controls("Txtreg3")): GoTo EH
          Else
            If Application.ActiveWorkbook.ProtectStructure Then Call RES.FrmCMsgbox("No se puede crear una nueva hoja en el libro actual porque está protegido. Quite la protección o seleccione otro libro para continuar.", STRAPPTITLE, vbInformation, True, .Controls("Optbal3")): GoTo EH
            Set GLORNG_CELL = Nothing
          End If

          Let StrMsg0 = "Se calculará la regresión lineal con los datos que ingresó." & VBA.vbNewLine & "¿Proceder?"
          Let View_core_validations_regression = (VBA.MsgBox(StrMsg0, vbQuestion + vbYesNo, STRAPPTITLE) = vbYes)

        Case Else: GoTo EH
      End Select
    End With

  End If

EH:
  Let VecAux0 = Empty
  Set RngBox = Nothing
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function View_core_validations_balmet( _
  ByVal VCmdBtn As MSForms.CommandButton, _
  ByVal FRM As FRMBasics, _
  Optional VFBooSelections As Boolean = False, _
  Optional ControlId As String, _
  Optional VRngBox As Range = Nothing) As Boolean

  Dim RngBox As Range
  Dim VecFeed As Variant, VecGrad As Variant, VecName As Variant, VecUnit As Variant, VecGradCx As Variant, VecNameCx As Variant, VecUnitCx As Variant
  Dim VecAux0 As Variant, VecAux1 As Variant, VecAux2 As Variant, VecAux3 As Variant, VecAux4 As Variant, VecAux5 As Variant
  Dim StrMsg0 As String, StrAux0 As String, StrAux1 As String
  Dim i As Long, j As Long, k As Long, l As Long, m As Long, n As Long
  Dim EnuTypeFeed As EQLBAL_ENU_TYPECONTENT, EnuBalType As EQLBAL_ENU_TYPE

  Const StrBALPRODUCTS As String = "Conjunto de leyes inválido." & VBA.vbNewLine & _
      "La referencia debe ser un rango de celdas o un grupo de números separados por punto y coma (Ej.: 1;2;3) con las siguientes dimensiones: " & VBA.vbNewLine & VBA.vbNewLine & _
      "Para Balance de 2 Productos: 3 filas x 1 columna." & VBA.vbNewLine & _
      "Para Balance de 3 Productos: 4 filas x 2 columnas." & VBA.vbNewLine & _
      "Para Balance de 4 Productos: 5 filas x 3 columnas..., etc." & VBA.vbNewLine & VBA.vbNewLine & _
      "Puede generar balances de hasta " & CLNG_BAL_LIMIT & " productos"
  Const StrBALPRODUCTSeco As String = "Conjunto de datos inválido." & VBA.vbNewLine & _
      "La referencia debe ser un rango de celdas o un grupo de números separados por punto y coma (Ej.: 1;2;3)." & VBA.vbNewLine & _
      "Puede generar balances de hasta " & CLNG_BAL_LIMIT & " productos"
  Const sFN As String = "VIEW::-Validations_balmet"

  On Error GoTo EH
  Let View_core_validations_balmet = False

  If VFBooSelections Then ' ========================================================== Validation for Selections contents'

    If ControlId = "balmetselection" Then ' ------------------------------------------ balmetselection'

      ' Data required in BVarRequest: (0.DblFeed, 1.DblFeedVol, 2.VecGrad, 3.VecGradCx, 4.VecName, 5.VecUnit, 6.EnuMethod, 7.Booleans, 8.EnuTypeFeed, 9.RngBox (, 10.StrFeedName, 11.StrNameA, 12.StrUnitA, 13. BooNewSheet))
      Set RngBox = Application.Selection
      If RngBox Is Nothing Then Call VBA.MsgBox("Balmet Selección: " & StrBALPRODUCTS & ".", vbInformation, STRAPPTITLE): GoTo EH
      With RngBox

        ' Requirements: Area 1: Feeds, Area 2: Grades, [Area 3: Grades Cx], [Names & Units]
        If .Areas.Count <= 1 Then Call VBA.MsgBox("Balmet Selección: Como mínimo debe seleccionar dos rangos, uno para la alimentación y otro para las leyes.", vbInformation, STRAPPTITLE): GoTo EH

        ' Feed
        Let VecFeed = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(.Areas(1), i, RValueType:=eqlResNumbers, RBooAbs:=True, RValueAs:=eqlResNumbers, RDblTop:=CDBL_MAX)
        If i <= 0 Then GoTo EH ' Feed is required

        ' Extract all balmet and cx grades
        Let VecGrad = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(.Areas(2), j, RValueType:=eqlResNumbers, RBooAbs:=True, RValueAs:=eqlResNumbers, RDblTop:=CDBL_MAX)
        If .Areas.Count > 2 Then Let VecGradCx = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(.Areas(3), k, RValueType:=eqlResNumbers, RBooAbs:=True, RValueAs:=eqlResNumbers, RDblTop:=CDBL_MAX)
        If j <= 0 Then Call VBA.MsgBox("Balmet Selección: ¡Se requieren valores numéricos para las leyes iniciales del balance metalúrgico!", vbInformation, STRAPPTITLE): GoTo EH
        If Not RES.RESCORE.BalIsBalmet(j, k) Then Call VBA.MsgBox(StrBALPRODUCTS & ".", vbInformation, STRAPPTITLE): GoTo EH ' Not a balmet
        ' Get the type of balmet to evaluate
        Let EnuBalType = RES.RESCORE.BalGetType(j, k) ' Balmet type
        If EnuBalType = eqlBalNothing Then Call VBA.MsgBox(StrBALPRODUCTS & ".", vbInformation, STRAPPTITLE): GoTo EH ' Not a balmet
        ' Get cols of products
        Let l = RES.RESCORE.BalGetProductsBounds(j, BWhat:=eqlBalColumns) ' Main Balmet cols
        Let m = RES.RESCORE.BalGetProductsBounds(j, k, BWhat:=eqlBalColumnsCx) ' Cx Balmet cols

        ' Filter names & units
        Let StrAux1 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(RngBox, n, RValueType:=eqlResNotNumbers, RStrCustomRegex:="%")

        ' Get info from db
        Let VecAux0 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalelements_sys, eqlMdlArray, "Símbolo"), eqlRes1D)
        Let VecAux1 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalelementscx_sys, eqlMdlArray, "Símbolo"), eqlRes1D)
        Let VecAux2 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalunits_sys, eqlMdlArray, "Unidad"), eqlRes1D)
        Let VecAux3 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalunitscx_sys, eqlMdlArray, "Unidad"), eqlRes1D)

        ' Names
        Call RES.RESCORE.BalFilterNames(RES, VecAux0, VBA.Split(StrAux1, ";"), VecName, True)
        Let StrAux0 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(VBA.CStr(VecName), n, RValueType:=eqlResNotNumbers, RBooUniques:=True)
        If Not RES.RESCORE.BalRefillNames2(RES, VecAux0, StrAux0, BLngGradCol:=(l - 1), BLngGradCxCol:=(m - 1), BVarName:=VecName, BVarNameCx:=VecNameCx, BStrErrMsg:=StrAux0, BBooCx:=(EnuBalType = eqlBalWithComplex), BVarDBNamesCx:=VecAux1) Then Call VBA.MsgBox("Balmet Selección: Los nombres de elementos son incorrectos.", vbInformation, STRAPPTITLE): GoTo EH

        ' Units
        Call RES.RESCORE.BalFilterUnits(RES, VecAux2, VBA.Split(StrAux1, ";"), VecUnit, True)
        Let StrAux0 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(VBA.CStr(VecUnit), n, RValueType:=eqlResNotNumbers, RStrCustomRegex:="%")
        If Not RES.RESCORE.BalRefillUnits2(RES, VecAux2, StrAux0, BLngGradCol:=(l - 1), BLngGradCxCol:=(m - 1), BVarUnit:=VecUnit, BVarUnitCx:=VecUnitCx, BStrErrMsg:=StrAux0, BBooCx:=(EnuBalType = eqlBalWithComplex), BVarDBUnitsCx:=VecAux3) Then Call VBA.MsgBox("Balmet Selección: Las unidades de elementos son incorrectas.", vbInformation, STRAPPTITLE): GoTo EH

        ' Validations normal Balmet
        If Not RES.RESCORE.BalCheckCustomNamesDoesNotHasWrongUnits(RES, StrAux0, BVarName:=VecName, BVarUnit:=VecUnit, BEnuCheckWhat:=eqlBalAgUnit) Then Call VBA.MsgBox("Balmet Selección: " & StrAux0, vbInformation, STRAPPTITLE): GoTo EH
        If Not RES.RESCORE.BalCheckCustomNamesDoesNotHasWrongUnits(RES, StrAux0, VecAux1, VecName, VecUnit, eqlBalCxElemsUnitOnNormalBalmet) Then Call VBA.MsgBox("Balmet Selección: " & StrAux0, vbInformation, STRAPPTITLE): GoTo EH
        If Not RES.RESCORE.BalCheckCustomNamesDoesNotHasWrongUnits(RES, StrAux0, BVarName:=VecName, BVarUnit:=VecUnit, BEnuCheckWhat:=eqlBalDMOnOtherElements) Then Call VBA.MsgBox("Balmet Selección: " & StrAux0, vbInformation, STRAPPTITLE): GoTo EH
        If Not RES.RESCORE.BalCheckForMaximousValuesInGrades(StrAux0, VecGrad, VecUnit) Then Call VBA.MsgBox("Balmet Selección: " & StrAux0, vbInformation, STRAPPTITLE): GoTo EH

        ' Validation Cx Balmet
        If EnuBalType = eqlBalWithComplex Then
          If Not RES.RESCORE.BalCheckCustomNamesDoesNotHasWrongUnits(RES, StrAux0, BVarSrcElems:=VecAux1, BVarName:=VecNameCx, BEnuCheckWhat:=eqlBalNonCx) Then Call VBA.MsgBox("Balmet Selección: " & StrAux0, vbInformation, STRAPPTITLE): GoTo EH
          If Not RES.RESCORE.BalCheckCustomNamesDoesNotHasWrongUnits(RES, StrAux0, BVarSrcElems:=Empty, BVarName:=VecNameCx, BVarUnit:=VecUnitCx, BEnuCheckWhat:=eqlBalDMOnOtherElements) Then Call VBA.MsgBox("Balmet Selección: " & StrAux0, vbInformation, STRAPPTITLE): GoTo EH
          If Not RES.RESCORE.BalCheckForMaximousValuesInGrades(StrAux0, VecGrad, VecUnit, VecGradCx, VecUnitCx, True) Then Call VBA.MsgBox("Balmet Selección: " & StrAux0, vbInformation, STRAPPTITLE): GoTo EH
        End If

        ' Feeds
        Let StrAux0 = VBA.Join(VecUnit, ";")
        If EnuBalType = eqlBalWithComplex Then Let StrAux0 = StrAux0 & ";" & VBA.Join(VecUnitCx, ";")
        Let EnuTypeFeed = RES.RESCORE.BalTypeFeed(RES, StrAux0)
        If Not RES.RESCORE.BalCheckForFeedsValues(VecFeed, EnuTypeFeed, CDBL_MAX, StrAux1) Then Call VBA.MsgBox("Balmet Selección: " & StrAux1, vbInformation, STRAPPTITLE): GoTo EH

      End With

      Set RngBox = Nothing
      Let View_core_validations_balmet = True

    Else ' --------------------------------------------------------------------------- "balmetselection-check_cell"

      ' Check for spreadsheet edges and data inside the region selected
      If VRngBox Is Nothing Then GoTo EH
      ' Check for protected data
      With VRngBox
        If GLOBOO_SELECTIONCELL Then
          If .Parent.Parent.ProtectStructure Then Call VBA.MsgBox("No se puede crear una nueva hoja en el libro actual porque está protegido. Quite la protección o seleccione otro libro para continuar.", vbInformation, STRAPPTITLE): GoTo EH
        Else
          If RES.ExcelIsProtectedSheet(VRngBox.Parent) Then Call VBA.MsgBox("La celda que seleccionó para imprimir los resultados está en una hoja protegida. Para realizar un cambio, debe quitar la protección de la hoja. Es posible que se le pida que escriba una contraseña.", vbInformation, STRAPPTITLE): GoTo EH
        End If
      End With
      If Not View_assets_check_for_spreadsheet_edges(VRngBox, (Application.Selection.Rows.Count * 3) + 10, 250, StrMsg0, VBooSelections:=True) Then Call VBA.MsgBox(StrMsg0, vbInformation, STRAPPTITLE): GoTo EH
      Let View_core_validations_balmet = (VBA.MsgBox("Balmet Selección: " & STRAPPTITLE & " calculará el balance metalúrgico con los datos que seleccionó." & VBA.vbNewLine & "¿Proceder?", vbQuestion + vbYesNo, STRAPPTITLE) = vbYes)

    End If

  Else ' ============================================================================= Validations to Forms contents

    With FRM
      Select Case FRM.Tag

        Case "balmet" ' -------------------------------------------------------------- balmet
          Select Case VCmdBtn.Name

            Case "Btnmain1"
              'Feed
              With .Controls("Txtbal0") ' Feed
                Let VecAux0 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(.Text, i, RBooAbs:=True, RValueAs:=eqlResNumbers, RDblTop:=CDBL_MAX)
                If i <= 0 Then Call CMB.CMsgBox("¡Alimentación debe ser un número mayor a cero!", "Error", cMsgType:=cmbError, cMsgButtons:=cmbClose, cUFrm:=FRM, cCtrlTargetName:=.Name, cCtrlSourceName:=VCmdBtn.Name, cCtrlRequired:=cmbRequired, cPositionX:=cmbRight, cPositionY:=cmbBottom, cAspect:=cmbTagLine): GoTo EH
              End With
              ' Feed name
              If .Controls("Chkbal0").Value And RES.REGEX.isEmptyStringReg(.Controls("Txtbal1").Text) Then Call CMB.CMsgBox("¿Y el dato?", "Error", cMsgType:=cmbQuestion, cMsgButtons:=cmbClose, cUFrm:=FRM, cCtrlTargetName:="Txtbal1", cCtrlSourceName:=VCmdBtn.Name, cPositionY:=cmbBottom, cCtrlRequired:=cmbRequired, cAspect:=cmbTagLine): GoTo EH

              ' Extract all balmet and cx grades
              With .Controls("Txtbal2")
                If Not RES.REGEX.isNumberEntireAndDecimalVectorColonSeparated(.Text) Then Call CMB.CMsgBox("¡Datos incorrectos!", "Error", cMsgType:=cmbError, cMsgButtons:=cmbClose, cUFrm:=FRM, cCtrlTargetName:=.Name, cCtrlSourceName:=VCmdBtn.Name, cPositionY:=cmbBottom, cCtrlRequired:=cmbRequired, cAspect:=cmbTagLine): GoTo EH
                Let VecGrad = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(.Value, j, RValueType:=eqlResNumbers, RBooAbs:=True, RValueAs:=eqlResNumbers, RDblTop:=CDBL_MAX)
                If j <= 0 Then Call CMB.CMsgBox("¡Datos incorrectos!", "Error", cMsgType:=cmbError, cMsgButtons:=cmbClose, cUFrm:=FRM, cCtrlTargetName:=.Name, cCtrlSourceName:=VCmdBtn.Name, cPositionY:=cmbBottom, cCtrlRequired:=cmbRequired, cAspect:=cmbTagLine): GoTo EH
                If Not RES.RESCORE.BalIsBalmet(j) Then Call RES.FrmCMsgbox(StrBALPRODUCTS & ".", STRAPPTITLE, vbInformation, True, FRM.Controls("Txtbal2"), RCtrlFocus:=True, RSelTxt:=True): GoTo EH
                ' Get cols of products
                Let l = RES.RESCORE.BalGetProductsBounds(j, BWhat:=eqlBalColumns) ' Main Balmet cols
                Let .Tag = l - 1
              End With

              ' Complex grades
              If .Controls("Chkbal3").Value Then
                With .Controls("Txtbal5")
                  If Not RES.REGEX.isNumberEntireAndDecimalVectorColonSeparated(.Text) Then Call CMB.CMsgBox("¡Datos incorrectos!", "Error", cMsgType:=cmbError, cMsgButtons:=cmbClose, cUFrm:=FRM, cCtrlTargetName:=.Name, cCtrlSourceName:=VCmdBtn.Name, cPositionY:=cmbBottom, cCtrlRequired:=cmbRequired, cAspect:=cmbTagLine): GoTo EH
                  Let VecGradCx = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(.Value, k, RValueType:=eqlResNumbers, RBooAbs:=True, RValueAs:=eqlResNumbers, RDblTop:=CDBL_MAX)
                  If k <= 0 Then Call CMB.CMsgBox("¡Datos incorrectos!", "Error", cMsgType:=cmbError, cMsgButtons:=cmbClose, cUFrm:=FRM, cCtrlTargetName:=.Name, cCtrlSourceName:=VCmdBtn.Name, cPositionY:=cmbBottom, cCtrlRequired:=cmbRequired, cAspect:=cmbTagLine): GoTo EH
                  If Not RES.RESCORE.BalIsBalmet(j, k) Then Call RES.FrmCMsgbox(StrBALPRODUCTS & " y hasta " & CLNG_BAL_LIMITCX & " tipos de minerales complejos, estos además, deben tener igual cantidad de filas que las leyes del balance principal.", STRAPPTITLE, vbInformation, True, FRM.Controls("Txtbal5"), RCtrlFocus:=True, RSelTxt:=True): GoTo EH
                  ' Get cols of products
                  Let m = RES.RESCORE.BalGetProductsBounds(j, k, BWhat:=eqlBalColumnsCx) ' BalmetCx cols
                  Let .Tag = m - 1
                End With
              End If

              ' Get the type of balmet to evaluate
              Let EnuBalType = RES.RESCORE.BalGetType(j, k) ' Balmet type
              If EnuBalType = eqlBalNothing Then Call VBA.MsgBox("Datos incorrectos", vbInformation, STRAPPTITLE): GoTo EH ' Not a balmet

              ' Get info from db
              Let VecAux0 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalelements_sys, eqlMdlArray, "Símbolo"), eqlRes1D)
              Let VecAux1 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalelementscx_sys, eqlMdlArray, "Símbolo"), eqlRes1D)
              Let VecAux2 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalunits_sys, eqlMdlArray, "Unidad"), eqlRes1D)
              Let VecAux3 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalunitscx_sys, eqlMdlArray, "Unidad"), eqlRes1D)

              ' Names
              With .Controls("Txtbal3")
                If RES.REGEX.isEmptyStringReg(.Text) Then Let .Parent.Controls("Chkbal1").Value = False
                Let StrAux0 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(.Value, n, RValueType:=eqlResNotNumbers, RBooUniques:=True)
                If Not RES.RESCORE.BalRefillNames2(RES, VecAux0, StrAux0, BLngGradCol:=(l - 1), BLngGradCxCol:=0, BVarName:=VecName, BVarNameCx:=Empty, BStrErrMsg:=StrAux0) Then Call RES.FrmCMsgbox(StrAux0, STRAPPTITLE, vbInformation, True, .Controls("Txtbal3"), vbCritical): GoTo EH
                Let .Value = VBA.Join(VecName, ";")
              End With

              ' Units
              With .Controls("Txtbal4")
                If RES.REGEX.isEmptyStringReg(.Text) Then Let .Parent.Controls("Chkbal2").Value = False
                Let StrAux0 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(.Value, n, RValueType:=eqlResNotNumbers, RStrCustomRegex:="%")
                If Not RES.RESCORE.BalRefillUnits2(RES, VecAux2, StrAux0, BLngGradCol:=(l - 1), BLngGradCxCol:=0, BVarUnit:=VecUnit, BVarUnitCx:=Empty, BStrErrMsg:=StrAux0) Then Call RES.FrmCMsgbox(StrAux0, STRAPPTITLE, vbInformation, True, .Controls("Txtbal4"), vbCritical): GoTo EH
                Let .Value = VBA.Join(VecUnit, ";")
              End With

              ' Validations normal Balmet
              If Not RES.RESCORE.BalCheckCustomNamesDoesNotHasWrongUnits(RES, StrAux0, BVarName:=VecName, BVarUnit:=VecUnit, BEnuCheckWhat:=eqlBalAgUnit) Then Call RES.FrmCMsgbox(StrAux0, STRAPPTITLE, vbInformation, True, .Controls("Txtbal3"), vbCritical): GoTo EH
              If Not RES.RESCORE.BalCheckCustomNamesDoesNotHasWrongUnits(RES, StrAux0, VecAux1, VecName, VecUnit, eqlBalCxElemsUnitOnNormalBalmet) Then Call RES.FrmCMsgbox(StrAux0, STRAPPTITLE, vbInformation, True, .Controls("Txtbal3"), vbCritical): GoTo EH
              If Not RES.RESCORE.BalCheckCustomNamesDoesNotHasWrongUnits(RES, StrAux0, BVarName:=VecName, BVarUnit:=VecUnit, BEnuCheckWhat:=eqlBalDMOnOtherElements) Then Call RES.FrmCMsgbox(StrAux0, STRAPPTITLE, vbInformation, True, .Controls("Txtbal4"), vbCritical): GoTo EH
              If Not RES.RESCORE.BalCheckForMaximousValuesInGrades(StrAux0, VecGrad, VecUnit) Then Call RES.FrmCMsgbox(StrAux0, STRAPPTITLE, vbInformation, True, .Controls("Txtbal2"), vbCritical): GoTo EH

              ' Validations Cx Balmet
              If .Controls("Chkbal3").Value Then
                If EnuBalType = eqlBalWithComplex Then

                  ' NamesCx
                  With .Controls("Txtbal6")
                    If RES.REGEX.isEmptyStringReg(.Text) Then Let .Parent.Controls("Chkbal4").Value = False
                    Let StrAux0 = FRM.Controls("Txtbal3").Value & ";" & .Value ' Name & Namecx
                    If Not RES.RESCORE.BalRefillNames2(RES, VecAux0, StrAux0, BLngGradCol:=(l - 1), BLngGradCxCol:=(m - 1), BVarName:=VecName, BVarNameCx:=VecNameCx, BStrErrMsg:=StrAux0, BBooCx:=True, BVarDBNamesCx:=VecAux1) Then Call RES.FrmCMsgbox(StrAux0, STRAPPTITLE, vbInformation, True, FRM.Controls("Txtbal6"), vbCritical): GoTo EH
                    Let .Value = VBA.Join(VecNameCx, ";")
                  End With

                  ' UnitsCx
                  With .Controls("Txtbal7")
                    If RES.REGEX.isEmptyStringReg(.Text) Then Let .Parent.Controls("Chkbal5").Value = False
                    Let StrAux0 = FRM.Controls("Txtbal4").Value & ";" & .Value ' Unit & Unitcx
                    If Not RES.RESCORE.BalRefillUnits2(RES, VecAux2, StrAux0, BLngGradCol:=(l - 1), BLngGradCxCol:=(m - 1), BVarUnit:=VecUnit, BVarUnitCx:=VecUnitCx, BStrErrMsg:=StrAux0, BBooCx:=True, BVarDBUnitsCx:=VecAux3) Then Call RES.FrmCMsgbox(StrAux0, STRAPPTITLE, vbInformation, True, FRM.Controls("Txtbal7"), vbCritical): GoTo EH
                    Let .Value = VBA.Join(VecUnitCx, ";")
                  End With

                  If Not RES.RESCORE.BalCheckCustomNamesDoesNotHasWrongUnits(RES, StrAux0, BVarSrcElems:=VecAux1, BVarName:=VecNameCx, BEnuCheckWhat:=eqlBalNonCx) Then Call RES.FrmCMsgbox(StrAux0, STRAPPTITLE, vbInformation, True, .Controls("Txtbal6"), vbCritical): GoTo EH
                  If Not RES.RESCORE.BalCheckCustomNamesDoesNotHasWrongUnits(RES, StrAux0, BVarSrcElems:=Empty, BVarName:=VecNameCx, BVarUnit:=VecUnitCx, BEnuCheckWhat:=eqlBalDMOnOtherElements) Then Call RES.FrmCMsgbox(StrAux0, STRAPPTITLE, vbInformation, True, .Controls("Txtbal6"), vbCritical): GoTo EH
                  If Not RES.RESCORE.BalCheckForMaximousValuesInGrades(StrAux0, VecGrad, VecUnit, VecGradCx, VecUnitCx, True) Then Call RES.FrmCMsgbox(StrAux0, STRAPPTITLE, vbInformation, True, .Controls("Txtbal7"), vbCritical): GoTo EH
                  If Not RES.RESCORE.BalCheckCustomNamesDoesNotHasWrongUnits(RES, StrAux0, BVarName:=VecNameCx, BVarUnit:=VecUnitCx, BEnuCheckWhat:=eqlBalDMOnOtherElements) Then Call RES.FrmCMsgbox(StrAux0, STRAPPTITLE, vbInformation, True, .Controls("Txtbal6"), vbCritical): GoTo EH
                End If
              End If

              ' Feed values
              Let VecAux0 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(.Controls("Txtbal0").Text, i, RBooAbs:=True, RValueAs:=eqlResNumbers, RDblTop:=CDBL_MAX)
              Let StrAux0 = .Controls("Txtbal4").Value & IIf(.Controls("Txtbal7").Value = VBA.vbNullString, VBA.vbNullString, ";" & .Controls("Txtbal7").Value)
              Let EnuTypeFeed = RES.RESCORE.BalTypeFeed(RES, StrAux0)
              If Not RES.RESCORE.BalCheckForFeedsValues(VecAux0, EnuTypeFeed, CDBL_MAX, StrAux0) Then Call VBA.MsgBox(StrAux0, vbInformation, STRAPPTITLE): GoTo EH
              Let VecAux0 = Empty

              ' Check for possible data in range area
              Let l = RES.RESCORE.BalGetProductsBounds(j, BWhat:=eqlBalProducts)

              Let n = (3 + 3 + (i * 4) + (j * 5) + (4 + (i * 2))) ' Cols: i: grades cols, j: complex cols, 3: initial fixed cols, 3: Vol, 4: %p & Rc
              Let m = (i * 3) + 10 ' Rows
              If .Controls("Optbal1").Value Then
                If RES.ExcelIsProtectedSheet(GLORNG_CELL.Parent) Then Call RES.FrmCMsgbox("La celda que seleccionó para imprimir los resultados está en una hoja protegida. Para realizar un cambio, debe quitar la protección de la hoja. Es posible que se le pida que escriba una contraseña.", STRAPPTITLE, vbInformation, True, .Controls("Txtbal8")): GoTo EH
                If Not View_assets_check_for_spreadsheet_edges(GLORNG_CELL, m, n, StrMsg0) Then
                  If Not StrMsg0 = VBA.vbNullString Then Call RES.FrmCMsgbox(StrMsg0, STRAPPTITLE, vbInformation, True, .Controls("Txtbal8"))
                  GoTo EH
                End If
              Else
                If Application.ActiveWorkbook.ProtectStructure Then Call RES.FrmCMsgbox("No se puede crear una nueva hoja en el libro actual porque está protegido. Quite la protección o seleccione otro libro para continuar.", STRAPPTITLE, vbInformation, True, .Controls("Optbal0")): GoTo EH
                Set GLORNG_CELL = Nothing: Let .Controls("Txtbal8").Text = VBA.vbNullString
              End If
              Call VBA.MsgBox("Lets operate!", vbInformation, STRAPPTITLE)
              'GoTo EH ' <=====!
              Let View_core_validations_balmet = (VBA.MsgBox("Se calculará el Balance Metalúrgico con los datos que ingresó." & VBA.vbNewLine & "¿Proceder?", vbQuestion + vbYesNo, STRAPPTITLE) = vbYes)

            Case Else: GoTo EH
          End Select

        Case "balmeteco" ' ----------------------------------------------------------- balmeteco
          If VCmdBtn.Name = "Btnmain1" Then

            ' Before: refilter each textbox data
            Let .Controls("Txtbal0").Value = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(.Controls("Txtbal0").Value, i, RBooAbs:=True, RDblTop:=CDBL_MAX)
            Let .Controls("Txtbal1").Value = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(.Controls("Txtbal1").Value, i, RBooAbs:=True, RDblTop:=CDBL_MAX)
            Let .Controls("Txtbal2").Value = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(.Controls("Txtbal2").Value, i, RBooAbs:=True, RDblTop:=CDBL_MAX)
            If .Controls("Chkbal0").Value Then Let .Controls("Txtbal3").Value = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(.Controls("Txtbal3").Value, i, RBooAbs:=True, RDblTop:=CDBL_MAX)

            ' 1. [x]  Type feed
            If Not RES.REGEX.isNumberEntireAndDecimalVectorColonSeparated(.Controls("Txtbal0").Value) Then Call CMB.CMsgBox("Dato requerido.", cMsgType:=cmbError, cMsgButtons:=cmbClose, cUFrm:=FRM, cCtrlTargetName:="Txtbal0", cCtrlRequired:=cmbRequired, cPositionX:=cmbRight, cPositionY:=cmbBottom, cAspect:=cmbTagLine): GoTo EH

            ' 2. [x]  Weigths, Grades, Recs, Names & Units lengths
            Let VecAux0 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(.Controls("Txtbal1").Value, i, RBooAbs:=True, RValueAs:=eqlResNumbers)
            Let VecAux1 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(.Controls("Txtbal2").Value, j, RBooAbs:=True, RValueAs:=eqlResNumbers)
            Let VecAux2 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(.Controls("Txtbal3").Value, k, RBooAbs:=True, RValueAs:=eqlResNumbers)
            Let StrAux0 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(.Controls("Txtbal5").Value, l, RValueType:=eqlResNotNumbers)
            Let StrAux0 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(.Controls("Txtbal6").Value, m, RValueType:=eqlResNotNumbers)
            ' Equal length for all vectors
            If Not RES.REGEX.isNumberEntireAndDecimalVectorColonSeparated(.Controls("Txtbal1").Value) Then Call CMB.CMsgBox("Dato incorrecto", cMsgType:=cmbError, cMsgButtons:=cmbClose, cUFrm:=FRM, cCtrlTargetName:="Txtbal1", cCtrlRequired:=cmbRequired, cPositionX:=cmbRight, cPositionY:=cmbBottom, cAspect:=cmbTagLine): GoTo EH
            If Not RES.REGEX.isNumberEntireAndDecimalVectorColonSeparated(.Controls("Txtbal2").Value) Then Call CMB.CMsgBox("Dato incorrecto", cMsgType:=cmbError, cMsgButtons:=cmbClose, cUFrm:=FRM, cCtrlTargetName:="Txtbal2", cCtrlRequired:=cmbRequired, cPositionX:=cmbRight, cPositionY:=cmbBottom, cAspect:=cmbTagLine): GoTo EH
            If i > CLNG_BAL_LIMIT Then Call RES.FrmCMsgbox(StrBALPRODUCTSeco & ".", STRAPPTITLE, vbInformation, True, FRM.Controls("Txtbal1"), RCtrlFocus:=True, RSelTxt:=True): GoTo EH
            If Not j = i Then Call CMB.CMsgBox("Dato incorrecto", cMsgType:=cmbError, cUFrm:=FRM, cCtrlTargetName:="Txtbal2", cCtrlRequired:=cmbRequired, cMsgButtons:=cmbClose, cPositionX:=cmbRight, cPositionY:=cmbBottom, cAspect:=cmbTagLine): GoTo EH
            If Not k = i And .Controls("Chkbal0").Value Then Call CMB.CMsgBox("Dato incorrecto", cMsgType:=cmbError, cUFrm:=FRM, cCtrlTargetName:="Txtbal3", cCtrlRequired:=cmbRequired, cMsgButtons:=cmbClose, cPositionX:=cmbRight, cPositionY:=cmbBottom, cAspect:=cmbTagLine): GoTo EH

            Let StrAux0 = VBA.vbNullString
            If l > 0 And .Controls("Chkbal2").Value Then Let StrAux0 = .Controls("Txtbal5").Value ' Get Names
            If Not RES.RESCORE.BalRefillNames(RES, StrAux0, BLngGradCol:=(i - 1), BLngGradCxCol:=0, BVarName:=VecAux3, BVarNameCx:=Empty) Then Call RES.FrmCMsgbox("Los nombres de elementos son incorrectas.", STRAPPTITLE, vbInformation, True, .Controls("Txtbal5"), vbCritical): GoTo EH
            Let .Controls("Txtbal5") = VBA.Join(VecAux3, ";")

            Let StrAux0 = VBA.vbNullString
            If m > 0 And .Controls("Chkbal3").Value Then Let StrAux0 = .Controls("Txtbal6").Value ' Get Units
            If Not RES.RESCORE.BalRefillUnits(RES, StrAux0, BLngGradCol:=(i - 1), BLngGradCxCol:=0, BVarUnit:=VecAux4, BVarUnitCx:=Empty) Then Call RES.FrmCMsgbox("Las unidades que introdujo son incorrectas.", STRAPPTITLE, vbInformation, True, .Controls("Txtbal6"), vbCritical): GoTo EH
            Let .Controls("Txtbal6") = VBA.Join(VecAux4, ";")

            ' 3. [x]  Feed Name
            If .Controls("Chkbal1").Value Then
              If RES.REGEX.isEmptyStringReg(.Controls("Txtbal4").Value) Then Call CMB.CMsgBox("¿Y el nombre de la alimentación?", cMsgType:=cmbQuestion, cMsgButtons:=cmbClose, cUFrm:=FRM, cCtrlTargetName:="Txtbal4", cCtrlRequired:=cmbRequired, cPositionX:=cmbLeft, cPositionY:=cmbBottom, cAspect:=cmbTagLine): GoTo EH
            End If

            ' 4. [x]  Names vs Units
            If .Controls("Chkbal2").Value Then
              Let VecAux5 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalelementscx_sys, eqlMdlArray, "Símbolo"), eqlRes1D)
              If Not RES.RESCORE.BalCheckCustomNamesDoesNotHasWrongUnits(RES, StrAux1, BVarSrcElems:=VecAux5, BVarName:=VecAux3, BVarUnit:=VecAux4, BEnuCheckWhat:=eqlBalCxElemsUnitOnNormalBalmet) Then Call RES.FrmCMsgbox(StrAux1, STRAPPTITLE, vbInformation, True, .Controls("Txtbal5"), vbCritical): GoTo EH
              If Not RES.RESCORE.BalCheckCustomNamesDoesNotHasWrongUnits(RES, StrAux1, BVarName:=VecAux3, BVarUnit:=VecAux4, BEnuCheckWhat:=eqlBalDMOnOtherElements) Then Call RES.FrmCMsgbox(StrAux1, STRAPPTITLE, vbInformation, True, .Controls("Txtbal6"), vbCritical): GoTo EH
              Let VecAux5 = Empty
            End If
            ' Check for maximous values in grades
            If Not RES.RESCORE.BalCheckForMaximousValuesInGrades(StrAux1, RES.ArrayToBaseZero(VecAux1, eqlRes2D), VecAux4, BBooOneToOne:=True) Then Call RES.FrmCMsgbox(StrAux1, STRAPPTITLE, vbInformation, True, .Controls("Txtbal2"), vbCritical): GoTo EH
            ' Check for maximous values in recs
            If Application.WorksheetFunction.Max(VecAux2) >= 100 Then Call RES.FrmCMsgbox("Ningún porcentaje de recuperación puede tener un valor mayor o igual a 100%.", STRAPPTITLE, vbInformation, True, .Controls("Txtbal3"), vbCritical): GoTo EH
            ' Check for volumes in feeds
            Let EnuTypeFeed = RES.RESCORE.BalTypeFeed(RES, VBA.Join(VecAux4, ";")) ' Feed values
            Let VecAux5 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(.Controls("Txtbal0").Value, n, RBooAbs:=True, RValueAs:=eqlResNumbers)
            If n > CLNG_BAL_LIMIT Or Not n = i Then Call RES.FrmCMsgbox(StrBALPRODUCTSeco & ".", STRAPPTITLE, vbInformation, True, FRM.Controls("Txtbal0"), RCtrlFocus:=True, RSelTxt:=True): GoTo EH
            If Not RES.RESCORE.BalCheckForVolumesAndSolidsFeeds(RES, VecAux5, VecAux0, VecAux4, EnuTypeFeed, CDBL_MAX, StrAux1) Then Call RES.FrmCMsgbox(StrAux1, STRAPPTITLE, vbInformation, True, .Controls("Txtbal1"), vbCritical): GoTo EH
            Let VecAux0 = Empty
            Let VecAux1 = Empty
            Let VecAux2 = Empty
            Let VecAux3 = Empty
            Let VecAux4 = Empty
            Let VecAux5 = Empty
            'GoTo EH ' <<<======!

            ' 5. [x]  Cx:
            If .Controls("Chkbal4").Value Then
              ' Before: refilter each textbox data
              Let .Controls("Txtbal7").Value = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(.Controls("Txtbal7").Value, i, RBooAbs:=True, RDblTop:=CDBL_MAX)
              If .Controls("Chkbal7").Value Then Let .Controls("Txtbal10").Value = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(.Controls("Txtbal10").Value, i, RBooAbs:=True, RDblTop:=CDBL_MAX)
              ' 6. [x]  Grades, Names, Units & Recs lengths - end cx
              Let VecAux0 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(.Controls("Txtbal7").Value, j, RBooAbs:=True, RValueAs:=eqlResNumbers)
              Let VecAux1 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(.Controls("Txtbal8").Value, k, RValueType:=eqlResNotNumbers, RValueAs:=eqlResStrings)
              Let VecAux2 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(.Controls("Txtbal9").Value, l, RValueType:=eqlResNotNumbers, RValueAs:=eqlResStrings)
              Let VecAux3 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(.Controls("Txtbal10").Value, m, RBooAbs:=True, RValueAs:=eqlResNumbers)

              ' Equal length for all vectors
              If Not RES.REGEX.isNumberEntireAndDecimalVectorColonSeparated(.Controls("Txtbal7").Value) Then Call CMB.CMsgBox("Dato incorrecto", cMsgType:=cmbError, cMsgButtons:=cmbClose, cUFrm:=FRM, cCtrlTargetName:="Txtbal7", cCtrlRequired:=cmbRequired, cPositionX:=cmbRight, cPositionY:=cmbBottom, cAspect:=cmbTagLine): GoTo EH
              If Not RES.REGEX.isNumberEntireAndDecimalVectorColonSeparated(.Controls("Txtbal10").Value) Then Call CMB.CMsgBox("Dato incorrecto", cMsgType:=cmbError, cMsgButtons:=cmbClose, cUFrm:=FRM, cCtrlTargetName:="Txtbal10", cCtrlRequired:=cmbRequired, cPositionX:=cmbRight, cPositionY:=cmbBottom, cAspect:=cmbTagLine): GoTo EH
              If Not m = j And .Controls("Chkbal7").Value Then Call CMB.CMsgBox("Dato incorrecto", cMsgType:=cmbError, cUFrm:=FRM, cCtrlTargetName:="Txtbal10", cCtrlRequired:=cmbRequired, cMsgButtons:=cmbClose, cPositionX:=cmbRight, cPositionY:=cmbBottom, cAspect:=cmbTagLine): GoTo EH
              If j > CLNG_BAL_LIMITCX Then Call RES.FrmCMsgbox(StrBALPRODUCTSeco & " y hasta " & CLNG_BAL_LIMITCX & " tipos de minerales complejos.", STRAPPTITLE, vbInformation, True, FRM.Controls("Txtbal7"), RCtrlFocus:=True, RSelTxt:=True): GoTo EH

              Let StrAux0 = VBA.vbNullString
              If j > 0 And .Controls("Chkbal5").Value Then Let StrAux0 = .Controls("Txtbal5").Value & IIf(RES.REGEX.isEmptyStringReg(.Controls("Txtbal8").Value), VBA.vbNullString, ";" & .Controls("Txtbal8").Value) ' Get Namescx
              If Not RES.RESCORE.BalRefillNames(RES, StrAux0, (i - 1), (j - 1), BVarName:=VecAux4, BVarNameCx:=VecAux1, BBooCx:=True) Then Call RES.FrmCMsgbox("Los nombres de especies complejas son incorrectos.", STRAPPTITLE, vbInformation, True, .Controls("Txtbal8"), vbCritical): GoTo EH
              Let .Controls("Txtbal8") = VBA.Join(VecAux1, ";")
              Let VecAux5 = Empty

              Let StrAux0 = VBA.vbNullString
              If m > 0 And .Controls("Chkbal6").Value Then Let StrAux0 = .Controls("Txtbal6").Value & IIf(RES.REGEX.isEmptyStringReg(.Controls("Txtbal9").Value), VBA.vbNullString, ";" & .Controls("Txtbal9").Value) ' Get Units
              If Not RES.RESCORE.BalRefillUnits(RES, StrAux0, (i - 1), (j - 1), BVarUnit:=VecAux5, BVarUnitCx:=VecAux2, BBooCx:=True) Then Call RES.FrmCMsgbox("Las unidades de especies complejas que introdujo son incorrectas.", STRAPPTITLE, vbInformation, True, .Controls("Txtbal9"), vbCritical): GoTo EH
              Let .Controls("Txtbal9") = VBA.Join(VecAux2, ";")
              Let VecAux5 = Empty

              ' Namescx vs Unitscx
              If .Controls("Chkbal5").Value Then
                Let VecAux5 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalelementscx_sys, eqlMdlArray, "Símbolo"), eqlRes1D)
                If Not RES.RESCORE.BalCheckCustomNamesDoesNotHasWrongUnits(RES, StrAux1, BVarSrcElems:=Empty, BVarName:=VecAux1, BVarUnit:=VecAux2, BEnuCheckWhat:=eqlBalDMOnOtherElements) Then Call RES.FrmCMsgbox(StrAux1, STRAPPTITLE, vbInformation, True, .Controls("Txtbal8"), vbCritical): GoTo EH
                If Not RES.RESCORE.BalCheckCustomNamesDoesNotHasWrongUnits(RES, StrAux1, BVarSrcElems:=VecAux5, BVarName:=VecAux1, BVarUnit:=VecAux2, BEnuCheckWhat:=eqlBalNonCx) Then Call RES.FrmCMsgbox(StrAux1, STRAPPTITLE, vbInformation, True, .Controls("Txtbal8"), vbCritical): GoTo EH
                Let VecAux5 = Empty
              End If

              ' Check for maximous values in gradescx
              If Not RES.RESCORE.BalCheckForMaximousValuesInGrades(StrAux1, RES.ArrayToBaseZero(VecAux0, eqlRes2D), VecAux2, BBooOneToOne:=True) Then Call RES.FrmCMsgbox(StrAux1, STRAPPTITLE, vbInformation, True, .Controls("Txtbal7"), vbCritical): GoTo EH
              ' Check for maximous values in recs
              If Application.WorksheetFunction.Max(VecAux3) >= 100 Then Call RES.FrmCMsgbox("Ningún porcentaje de recuperación puede tener un valor mayor o igual a 100%.", STRAPPTITLE, vbInformation, True, .Controls("Txtbal10"), vbCritical): GoTo EH
            End If
            Let VecAux0 = Empty
            Let VecAux1 = Empty
            Let VecAux2 = Empty
            Let VecAux3 = Empty
            Let VecAux4 = Empty
            Let VecAux5 = Empty

            ' 7. [x]  RngCell
            Let n = (3 + 3 + (i * 4) + (j * 5) + (4 + (i * 2))) ' Cols: i: grades cols, j: complex cols, 3: initial fixed cols, 3: Vol, 4: %p & Rc
            Let m = (i * 3) + 10 ' Rows
            If .Controls("Optbal1").Value Then
              If RES.ExcelIsProtectedSheet(GLORNG_CELL.Parent) Then Call RES.FrmCMsgbox("La celda que seleccionó para imprimir los resultados está en una hoja protegida. Para realizar un cambio, debe quitar la protección de la hoja. Es posible que se le pida que escriba una contraseña.", STRAPPTITLE, vbInformation, True, .Controls("Txtbal11")): GoTo EH
              If Not View_assets_check_for_spreadsheet_edges(GLORNG_CELL, m, n, StrMsg0) Then
                If Not StrMsg0 = VBA.vbNullString Then Call RES.FrmCMsgbox(StrMsg0, STRAPPTITLE, vbInformation, True, .Controls("Txtbal11"))
                GoTo EH
              End If
            Else
              If Application.ActiveWorkbook.ProtectStructure Then Call RES.FrmCMsgbox("No se puede crear una nueva hoja en el libro actual porque está protegido. Quite la protección o seleccione otro libro para continuar.", STRAPPTITLE, vbInformation, True, .Controls("Optbal0")): GoTo EH
              Set GLORNG_CELL = Nothing: Let .Controls("Txtbal11").Text = VBA.vbNullString
            End If
            Let View_core_validations_balmet = (VBA.MsgBox("Se calculará el Balance Metalúrgico con los datos que ingresó." & VBA.vbNewLine & "¿Proceder?", vbQuestion + vbYesNo, STRAPPTITLE) = vbYes)

          End If

        Case Else: GoTo EH
      End Select
    End With

  End If

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Let VecAux2 = Empty
  Let VecAux3 = Empty
  Let VecAux4 = Empty
  Let VecAux5 = Empty
  Let VecFeed = Empty
  Let VecGrad = Empty
  Let VecName = Empty
  Let VecUnit = Empty
  Let VecGradCx = Empty
  Let VecNameCx = Empty
  Let VecUnitCx = Empty
  Set RngBox = Nothing
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function View_Core_validations_projections( _
  ByVal VCmdBtn As MSForms.CommandButton, _
  ByVal FRM As FRMBasics) As Boolean

  Dim VecAux0 As Variant
  Dim StrMsg0 As String, StrAux0 As String
  Dim i As Long, j As Long, k As Long, m As Long
  Dim DblAux0 As Double

  Const sFN As String = "VIEW::-Validations_projections"

  On Error GoTo EH
  Let View_Core_validations_projections = False

  With FRM
    Select Case VCmdBtn.Name

      Case "Btnmain1"
        ' Inversion
        With .Controls("Txtpro0")
          Let VecAux0 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(.Value, i, RValueType:=eqlResNumbers, RBooAbs:=True, RValueAs:=eqlResNumbers, RDblTop:=CDBL_MAX)
          If i = 0 Or Application.WorksheetFunction.Sum(VecAux0) = 0 Then Call CMB.CMsgBox("¡Se requiere el valor de Inversión!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, .Name, VCmdBtn.Name, cmbRequired, cmbRight, cmbBottom, cAspect:=cmbTagLine): GoTo EH
        End With

        ' Fine Weight
        With .Controls("Txtpro1")
          If Not RES.REGEX.isNumberEntireAndDecimalVectorColonSeparated(.Value) Then Call CMB.CMsgBox("¡Dato incorrecto!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, .Name, VCmdBtn.Name, cmbRequired, cmbRight, cmbBottom, cAspect:=cmbTagLine): GoTo EH
          Let VecAux0 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(.Value, i, RValueType:=eqlResNumbers, RBooAbs:=True, RValueAs:=eqlResNumbers, RDblTop:=CDBL_MAX)
          If i = 0 Then Call CMB.CMsgBox("¡Como mínimo un número!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, .Name, VCmdBtn.Name, cmbRequired, cmbRight, cmbBottom, cAspect:=cmbTagLine): GoTo EH
          If i > CLNG_BAL_LIMIT Then Call RES.FrmCMsgbox("¡Sólo hasta " & CLNG_BAL_LIMIT & " minerales a la vez!", STRAPPTITLE, vbInformation, True, FRM.Controls("Txtpro1"), vbCritical, True, True): GoTo EH
        End With

        ' Ctz
        With .Controls("Txtpro2")
          If Not RES.REGEX.isNumberWithDecimalsSemicolonAndSeparatedSlashGrouped(.Value) Then Call CMB.CMsgBox("¡Dato incorrecto!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, .Name, VCmdBtn.Name, cmbRequired, cmbRight, cmbBottom, cAspect:=cmbTagLine): GoTo EH
          Let VecAux0 = RES.ArrayGetDataSemicolonSeparatedSlashGroupedStringFromVariants(.Value, j, False)
          If j = 0 Then Call RES.FrmCMsgbox("¡Se requiere como mínimo un rango de tres valores de Cotización del Mineral Fino!", STRAPPTITLE, vbInformation, True, FRM.Controls("Txtpro2"), vbCritical, True, True): GoTo EH
          If Not j = i Then Call RES.FrmCMsgbox("¡Tiene " & IIf(i = 1, "un mineral", i & " minerales") & ". Es necesario un rango de Cotizaciones por cada Peso Fino de Mineral!", STRAPPTITLE, vbInformation, True, FRM.Controls("Txtpro2"), vbCritical, True, True): GoTo EH
          Let VecAux0 = VBA.Split(.Value, "/")
          For m = 0 To UBound(VecAux0)
            Let StrAux0 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(VBA.CStr(VecAux0(m)), j, RValueType:=eqlResNumbers, RBooAbs:=False, RBooIncludeZeroes:=True, RDblTop:=CDBL_MAX)
            If j = 0 Or j < 3 Then Call RES.FrmCMsgbox("¡Cada rango de Cotizaciones debe contener como mínimo tres valores!", STRAPPTITLE, vbInformation, True, FRM.Controls("Txtpro2"), vbCritical, True, True): GoTo EH
          Next m
          Let VecAux0 = Empty
        End With

        ' RV
        With .Controls("Txtpro3")
          Let .Value = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(.Value, i, RValueType:=eqlResNumbers, RBooAbs:=True, RDblTop:=CDBL_MAX)
        End With

        ' %Outs
        With .Controls("Txtpro4")
          If Not RES.REGEX.isPercentReg(.Value) Then Call CMB.CMsgBox("¡Dato incorrecto!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, .Name, VCmdBtn.Name, cmbRequired, cmbRight, cmbBottom, cAspect:=cmbTagLine): GoTo EH
          Let DblAux0 = VBA.CDbl(VBA.Replace(VBA.Replace(.Value, "%", VBA.vbNullString), ".", ","))
          If DblAux0 > 100 Or DblAux0 < 0 Then Call CMB.CMsgBox("¡Porcentaje incorrecto!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, .Name, VCmdBtn.Name, cmbRequired, cmbRight, cmbBottom, cAspect:=cmbTagLine): GoTo EH
        End With

        ' %Time
        With .Controls("Txtpro7")
          If Not RES.REGEX.isNumberReg(.Value) Then Call CMB.CMsgBox("¡Dato incorrecto!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, .Name, VCmdBtn.Name, cmbRequired, cmbRight, cmbBottom, cAspect:=cmbTagLine): GoTo EH
          If VBA.CLng(.Value) <= 0 Then Call CMB.CMsgBox("¡Dato incorrecto!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, .Name, VCmdBtn.Name, cmbRequired, cmbRight, cmbBottom, cAspect:=cmbTagLine): GoTo EH
        End With

        ' %Rate
        With .Controls("Txtpro8")
          If Not RES.REGEX.isPercentReg(.Value) Then Call CMB.CMsgBox("¡Dato incorrecto!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, .Name, VCmdBtn.Name, cmbRequired, cmbRight, cmbBottom, cAspect:=cmbTagLine): GoTo EH
          Let DblAux0 = VBA.CDbl(VBA.Replace(.Value, "%", VBA.vbNullString))
          If DblAux0 > 100 Or DblAux0 < 0 Then Call CMB.CMsgBox("¡Porcentaje incorrecto!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, .Name, VCmdBtn.Name, cmbRequired, cmbRight, cmbBottom, cAspect:=cmbTagLine): GoTo EH
        End With

        ' Check for spreadsheet edges and data inside the region selected
        If .Controls("Optpro3").Value Then
          Let k = ((i + 6) * 5) '5: A1 to totalsrow aprox
          If GLORNG_CELL Is Nothing Then Call RES.FrmCMsgbox("¡Debe seleccionar una celda para imprimir los resultados!", STRAPPTITLE, vbInformation, True, FRM.Controls("Txtpro9"), vbCritical): GoTo EH
          If RES.ExcelIsProtectedSheet(GLORNG_CELL.Parent) Then Call RES.FrmCMsgbox("La celda que seleccionó para imprimir los resultados está en una hoja protegida. Para realizar un cambio, debe quitar la protección de la hoja. Es posible que se le pida que escriba una contraseña.", STRAPPTITLE, vbInformation, True, .Controls("Txtpro9")): GoTo EH
          If Not View_assets_check_for_spreadsheet_edges(GLORNG_CELL, k, 50, StrMsg0) Then Call RES.FrmCMsgbox(StrMsg0, STRAPPTITLE, vbInformation, True, .Controls("Txtpro9")): GoTo EH
        Else
          If Application.ActiveWorkbook.ProtectStructure Then Call RES.FrmCMsgbox("No se puede crear una nueva hoja en el libro actual porque está protegido. Quite la protección o seleccione otro libro para continuar.", STRAPPTITLE, vbInformation, True, .Controls("Optpro2")): GoTo EH
          Set GLORNG_CELL = Nothing
        End If

        Let StrMsg0 = "Se calculará la Proyección del Proyecto con los datos que ingresó." & VBA.vbNewLine & "¿Proceder?"
        Let View_Core_validations_projections = (VBA.MsgBox(StrMsg0, vbQuestion + vbYesNo, STRAPPTITLE) = vbYes)

      Case Else: GoTo EH
    End Select
  End With

EH:
  Let VecAux0 = Empty
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function View_Core_validations_oreinvoice( _
  ByVal VBtnCurrent As MSForms.CommandButton, _
  ByVal FRM As FRMBasics) As Boolean

  Dim VecAux0 As Variant, VecAux1 As Variant, VecAux2 As Variant
  Dim StrMsg0 As String, StrAux0 As String, StrAux1 As String, StrAux2 As String
  Dim i As Long, j As Long, k As Long, l As Long, m As Long, n As Long
  Dim BooAux0 As Boolean, BooAux1 As Boolean, BooAux2 As Boolean

  Const sFN As String = "VIEW::-Validations_oreinvoice"

  On Error GoTo EH
  Let View_Core_validations_oreinvoice = False

  With FRM
    Select Case VBtnCurrent.Name

      Case "Btnpur0", "Btnpur1" ' Insert, Update
        Let j = .Controls("Cbopur1").ListIndex
        Let i = MDL.LengthRows(IIf(j < 2, eqlMdlSheetPur, eqlMdlSheetTmp), IIf(j = 0 Or j = 1, eqlMdlTblPurchasemin_pur, IIf(j = 2, eqlMdlTblPurbonus_tmp, eqlMdlTblPurdiscounts_tmp)))
        Let k = IIf(j = 0 Or j = 1, GLOSNG_PUR_LIMIT, CSNG_BONUS_DISCOUNT_LIMIT)
        If VBtnCurrent.Name = "Btnpur1" And i <= 0 Then GoTo EH

        Select Case j

          Case 0 ' Base Data: Texts 5,6,7,8,9
            If i = k And VBtnCurrent.Name = "Btnpur0" Then Call RES.FrmCMsgbox("Sólo se admiten hasta " & k & " registros de " & .Controls("Cbopur1").Value, STRAPPTITLE, vbInformation, True, .Controls("Lstpur0"), vbCritical): GoTo EH ' All inside the limit
            ' MUESTRA
            If i > 0 Then
              Let VecAux0 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur, eqlMdlArray, "MUESTRA"), eqlRes1D)
              If RES.REGEX.isStringReg(.Controls("Txtpur5").Value) Then
                If VBtnCurrent.Name = "Btnpur1" Then ' Update
                  If Not VBA.LCase(.Controls("Txtpur5").Value) = VBA.LCase(.Controls("Lstpur0").List(.Controls("Lstpur0").ListIndex, 0)) Then
                    If RES.ArrayHasValue(VecAux0, VBA.LCase(.Controls("Txtpur5").Value), ALCase:=True) Then Call CMB.CMsgBox("¡La Muestra: " & .Controls("Txtpur5").Value & " ya existe!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtpur5", "Btnmain1;Chkpur1;Chkpur2;Cbopur1;Cbopur2;Cbopur3;Cbopur4", cmbRequired, cmbLeft, cmbBottom, cAspect:=cmbTagLine): GoTo EH
                  End If
                Else ' Insert
                  If RES.ArrayHasValue(VecAux0, VBA.LCase(.Controls("Txtpur5").Value), ALCase:=True) Then Call CMB.CMsgBox("¡La Muestra: " & .Controls("Txtpur5").Value & " ya existe!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtpur5", "Btnmain1;Chkpur1;Chkpur2;Cbopur1;Cbopur2;Cbopur3;Cbopur4", cmbRequired, cmbLeft, cmbBottom, cAspect:=cmbTagLine): GoTo EH
                End If
              End If
            End If
            Let VecAux0 = Empty
            ' FECHA
            If Not RES.REGEX.isDateReg(.Controls("Txtpur6").Value) Then Call CMB.CMsgBox("¡Fecha inválida!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtpur6", "Btnmain1;Chkpur1;Chkpur2;Cbopur1;Cbopur2;Cbopur3;Cbopur4", cmbRequired, cmbLeft, cmbBottom, cAspect:=cmbTagLine): GoTo EH
            ' %HUMEDAD
            If Not RES.REGEX.isPercentReg(.Controls("Txtpur7").Value) Then Call CMB.CMsgBox("¡Porcentaje de Humedad inválido!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtpur7", "Btnmain1;Chkpur1;Chkpur2;Cbopur1;Cbopur2;Cbopur3;Cbopur4", cmbRequired, cmbLeft, cmbBottom, cAspect:=cmbTagLine): GoTo EH
            ' %MERMA
            If Not RES.REGEX.isPercentReg(.Controls("Txtpur8").Value) Then Call CMB.CMsgBox("¡Porcentaje de Merma inválido!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtpur8", "Btnmain1;Chkpur1;Chkpur2;Cbopur1;Cbopur2;Cbopur3;Cbopur4", cmbRequired, cmbLeft, cmbBottom, cAspect:=cmbTagLine): GoTo EH
            ' PBH
            If Not RES.REGEX.isNumberReg(.Controls("Txtpur9").Value) Then Call CMB.CMsgBox("¡Peso Bruto Húmedo inválido!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtpur9", "Btnmain1;Chkpur1;Chkpur2;Cbopur1;Cbopur2;Cbopur3;Cbopur4", cmbRequired, cmbRight, cmbBottom, cAspect:=cmbTagLine): GoTo EH

          Case 1 ' Main data: Texts 6,7,8,9
            If i <= 0 Then Call RES.FrmCMsgbox("Primero ingrese datos en la opción: " & VBA.CStr(.Controls("Cbopur1").List(0, 0)) & ".", STRAPPTITLE, vbInformation, True, .Controls("Cbopur1"), vbCritical): GoTo EH
            ' Check if table has 0's in PARTE2 column, if it has, update second part, else cancel process
            Let VecAux0 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur, eqlMdlArray, "PARTE2"), eqlRes1D)
            ' Table has all rows updated
            If RES.ArrayHasValue(VecAux0, "0") Then Call RES.FrmCMsgbox("Primero ingrese datos en la opción: " & VBA.CStr(.Controls("Cbopur1").List(0, 0)) & ".", STRAPPTITLE, vbInformation, True, .Controls("Cbopur1"), vbCritical): GoTo EH
            Let VecAux0 = Empty
            Let k = .Controls("Cbopur2").ListIndex
            ' Table has rows for update
            If k = 1 Then ' TIPO: Complejos

              ' CONTENIDOS
              If RES.REGEX.isEmptyStringReg(.Controls("Txtpur6").Value) Then Call CMB.CMsgBox("¡Se esperaba al menos dos elementos!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtpur6", "Btnmain1;Chkpur1;Chkpur2;Cbopur1;Cbopur2;Cbopur3;Cbopur4", cmbRequired, cmbLeft, cmbBottom, cAspect:=cmbTagLine): GoTo EH
              Let VecAux0 = VBA.Split(.Controls("Txtpur6").Value, ";"): Let i = UBound(VecAux0) - LBound(VecAux0)
              If i < 1 Then Call CMB.CMsgBox("¡Se esperaba al menos dos elementos!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtpur6", "Btnmain1;Chkpur1;Chkpur2;Cbopur1;Cbopur2;Cbopur3;Cbopur4", cmbRequired, cmbLeft, cmbBottom, cAspect:=cmbTagLine): GoTo EH
              If RES.ArrayHasRepeatedElements(VecAux0) Then Call RES.FrmCMsgbox("¡Un mineral complejo debe contener distintos elementos!", STRAPPTITLE, vbInformation, True, .Controls("Txtpur6"), vbCritical): GoTo EH
              Let VecAux0 = Empty
              ' LEYES
              If Not RES.REGEX.isNumberEntireAndDecimalVectorColonSeparated(.Controls("Txtpur7").Value) Then Call CMB.CMsgBox("¡Se esperaba al menos dos leyes!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtpur7", "Btnmain1;Chkpur1;Chkpur2;Cbopur1;Cbopur2;Cbopur3;Cbopur4", cmbRequired, cmbLeft, cmbBottom, cAspect:=cmbTagLine): GoTo EH
              Let VecAux0 = VBA.Split(.Controls("Txtpur7").Value, ";"): Let j = UBound(VecAux0) - LBound(VecAux0)
              If Not j = i Then Call RES.FrmCMsgbox("¡Debe existir la misma cantidad de leyes y elementos!", STRAPPTITLE, vbInformation, True, .Controls("Txtpur7"), vbCritical): GoTo EH
              Let VecAux0 = Empty
              ' UNIDADES
              If RES.REGEX.isEmptyStringReg(.Controls("Txtpur8").Value) Then Call CMB.CMsgBox("¿Las unidades de los elementos?", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtpur8", "Btnmain1;Chkpur1;Chkpur2;Cbopur1;Cbopur2;Cbopur3;Cbopur4", cmbRequired, cmbLeft, cmbBottom, cAspect:=cmbTagLine): GoTo EH
              Let VecAux0 = VBA.Split(.Controls("Txtpur8").Value, ";"): Let j = UBound(VecAux0) - LBound(VecAux0)
              If Not j = i Then Call RES.FrmCMsgbox("¡Debe existir la misma cantidad de unidades y elementos!", STRAPPTITLE, vbInformation, True, .Controls("Txtpur8"), vbCritical): GoTo EH
              Let VecAux0 = Empty
              ' PRECIO/PTO
              If Not RES.REGEX.isNumberEntireAndDecimalVectorColonSeparated(.Controls("Txtpur9").Value) Then Call CMB.CMsgBox("¿Los precios?", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtpur9", "Btnmain1;Chkpur1;Chkpur2;Cbopur1;Cbopur2;Cbopur3;Cbopur4", cmbRequired, cmbLeft, cmbBottom, cAspect:=cmbTagLine): GoTo EH
              Let VecAux0 = VBA.Split(.Controls("Txtpur9").Value, ";"): Let j = UBound(VecAux0) - LBound(VecAux0)
              If Not j = i Then Call RES.FrmCMsgbox("¡Debe existir la misma cantidad de precios y elementos!", STRAPPTITLE, vbInformation, True, .Controls("Txtpur9"), vbCritical): GoTo EH
              Let VecAux0 = Empty

            Else

              ' LEYES & PRECIO/PTO
              If Not RES.REGEX.isNumberReg(.Controls("Txtpur7").Value) Then Call CMB.CMsgBox("¡Se esperaba al menos una ley!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtpur7", "Btnmain1;Chkpur1;Chkpur2;Cbopur1;Cbopur2;Cbopur3;Cbopur4", cmbRequired, cmbLeft, cmbBottom, cAspect:=cmbTagLine): GoTo EH
              If Not RES.REGEX.isNumberEntireAndDecimalVectorColonSeparated(.Controls("Txtpur9").Value) Then Call CMB.CMsgBox("¿Y el precio?", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtpur9", "Btnmain1;Chkpur1;Chkpur2;Cbopur1;Cbopur2;Cbopur3;Cbopur4", cmbRequired, cmbRight, cmbBottom, cAspect:=cmbTagLine): GoTo EH

            End If
            ' Check for DM's & Ag's in elements & units lists
            If Not RES.RESCORE.BalCheckCustomNamesDoesNotHasWrongUnits(RES, StrAux0, Empty, VBA.Split(.Controls(IIf(k = 1, "Txtpur6", "Cbopur3")).Value, ";"), VBA.Split(.Controls(IIf(k = 1, "Txtpur8", "Cbopur4")).Value, ";"), eqlBalAgUnit) Then Call RES.FrmCMsgbox(StrAux0, STRAPPTITLE, vbInformation, True, .Controls(IIf(k = 1, "Txtpur8", "Cbopur4")), vbCritical): GoTo EH
            If Not RES.RESCORE.BalCheckCustomNamesDoesNotHasWrongUnits(RES, StrAux0, Empty, VBA.Split(.Controls(IIf(k = 1, "Txtpur6", "Cbopur3")).Value, ";"), VBA.Split(.Controls(IIf(k = 1, "Txtpur8", "Cbopur4")).Value, ";"), eqlBalDMOnOtherElements) Then Call RES.FrmCMsgbox(StrAux0, STRAPPTITLE, vbInformation, True, .Controls(IIf(k = 1, "Txtpur8", "Cbopur4")), vbCritical): GoTo EH
            If Not RES.RESCORE.BalCheckForMaximousValuesInGrades(StrAux0, RES.ArrayToBaseZero(VBA.Split(.Controls("Txtpur7").Value, ";"), eqlRes2D), VBA.Split(.Controls(IIf(k = 1, "Txtpur8", "Cbopur4")).Value, ";"), BBooOneToOne:=True) Then Call RES.FrmCMsgbox(StrAux0, STRAPPTITLE, vbInformation, True, .Controls("Txtpur7"), vbCritical): GoTo EH

          Case 2, 5 ' Bonus & Discounts: Texts 5 - 6
            If i = k Then Call RES.FrmCMsgbox("Sólo se admiten hasta " & k & " registros de " & .Controls("Cbopur1").Value, STRAPPTITLE, vbInformation, True, .Controls("Lstpur0"), vbCritical): GoTo EH ' All inside the limit
            ' BONO|DESCUENTO - VALOR
            If RES.REGEX.isEmptyStringReg(.Controls("Txtpur5").Value) And VBtnCurrent.Name = "Btnpur1" Then Call CMB.CMsgBox("¡El " & IIf(j = 2, "bono", "descuento") & " debe tener un nombre!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtpur5", "Btnmain1", cmbRequired, cmbLeft, cmbBottom, cAspect:=cmbTagLine): GoTo EH
            If i > 0 Then
              Let VecAux0 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetTmp, IIf(j = 2, eqlMdlTblPurbonus_tmp, eqlMdlTblPurdiscounts_tmp), eqlMdlArray, IIf(j = 2, "BONO", "DESCUENTO")), eqlRes1D)
              If RES.REGEX.isStringReg(.Controls("Txtpur5").Value) Then
                If VBtnCurrent.Name = "Btnpur1" Then ' Update
                  If Not VBA.LCase(.Controls("Txtpur5").Value) = VBA.LCase(.Controls("Lstpur0").List(.Controls("Lstpur0").ListIndex, 0)) Then
                    If RES.ArrayHasValue(VecAux0, VBA.LCase(.Controls("Txtpur5").Value), ALCase:=True) Then Call CMB.CMsgBox("¡EL " & IIf(j = 2, "BONO", "DESCUENTO") & ": " & .Controls("Txtpur5").Value & " ya existe!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtpur5", "Btnmain1;Chkpur1;Chkpur2;Cbopur1;Cbopur2;Cbopur3;Cbopur4", cmbRequired, cmbLeft, cmbBottom, cAspect:=cmbTagLine): GoTo EH
                  End If
                Else ' Insert
                  If RES.ArrayHasValue(VecAux0, VBA.LCase(.Controls("Txtpur5").Value), ALCase:=True) Then Call CMB.CMsgBox("¡El " & IIf(j = 2, "BONO", "DESCUENTO") & ": " & .Controls("Txtpur5").Value & " ya existe!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtpur5", "Btnmain1;Chkpur1;Chkpur2;Cbopur1;Cbopur2;Cbopur3;Cbopur4", cmbRequired, cmbLeft, cmbBottom, cAspect:=cmbTagLine): GoTo EH
                End If
              End If
              Let VecAux0 = Empty
            End If
            If Not RES.REGEX.isNumberReg(.Controls("Txtpur6").Text) Then Call CMB.CMsgBox("¡" & IIf(j = 2, "Bono", "Descuento") & " inválido!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtpur6", "Btnmain1;Chkpur1;Chkpur2;Cbopur1;Cbopur2;Cbopur3;Cbopur4", cmbRequired, cmbLeft, cmbBottom, cAspect:=cmbTagLine): GoTo EH
            If VBA.Abs(VBA.CBool(.Controls("Txtpur6").Text)) = 0 Then Call CMB.CMsgBox("¡" & IIf(j = 2, "Bono", "Descuento") & " inválido!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtpur6", "Btnmain1;Chkpur1;Chkpur2;Cbopur1;Cbopur2;Cbopur3;Cbopur4", cmbRequired, cmbLeft, cmbBottom, cAspect:=cmbTagLine): GoTo EH

          Case Else: GoTo EH

        End Select
        Let View_Core_validations_oreinvoice = True

      Case "Btnpur3", "Btnpur4" ' Del once, del all
        Let i = .Controls("Cbopur1").ListIndex
        Let j = .Controls("Lstpur0").ListIndex + 1
        If j < 1 And VBtnCurrent.Name = "Btnpur3" Or MDL.LengthRows(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur) <= 0 Then GoTo EH
        Let View_Core_validations_oreinvoice = (VBA.MsgBox("Eliminará " & IIf(VBtnCurrent.Name = "Btnpur3", "el registro seleccionado ", "todos los registros ") & "de la tabla: " & VBA.CStr(.Controls("Cbopur1").List(i, 0)) & IIf(i < 2, ", también se eliminarán los datos de: " & VBA.CStr(.Controls("Cbopur1").List(IIf(i = 0, 1, 0), 0)), ".") & VBA.vbNewLine & VBA.vbNewLine & "¿Proceder?", vbYesNo + vbQuestion, STRAPPTITLE) = vbYes)

      Case "Btnpur5" ' Validate range
        If GLORNG_CELL Is Nothing Then GoTo EH
        If GLORNG_CELL.Areas.Count > 1 Then Call RES.FrmCMsgbox("Sólo se requiere la selección de un rango de datos.", STRAPPTITLE, vbInformation, True, .Controls("Btnpur5"), vbCritical): GoTo EH
        Let k = .Controls("Cbopur1").ListIndex
        Let i = MDL.LengthRows(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur)
        Let j = GLORNG_CELL.Rows.Count
        If i = GLOSNG_PUR_LIMIT And k = 0 Then Call RES.FrmCMsgbox("La base de datos ya alcanzó el límite de " & GLOSNG_PUR_LIMIT & " registros.", STRAPPTITLE, vbInformation, True, .Controls("Btnpur5"), vbCritical): GoTo EH
        If Application.WorksheetFunction.CountA(GLORNG_CELL) = 0 Then Call RES.FrmCMsgbox("Existen celdas sin datos en la selección, todas las celdas del rango seleccionado deben tener datos.", STRAPPTITLE, vbInformation, True, .Controls("Btnpur5"), vbCritical): GoTo EH

        Select Case k

          Case 0, 1 ' Base Data And/Or Main data: Texts 5,6,7,8,9 (5 or 10 cols)
            ' 0: Get 5 cols in array, if there are 10 cols, slice it in two arrays
            ' 1: Get 5 cols in array, if there are 10 cols, slice it in two arrays and check if Main data have at least 1 row
            Let VecAux0 = RES.ArrayToBaseZero(GLORNG_CELL.Value)
            If (i + j) > GLOSNG_PUR_LIMIT And k = 0 Then Let VecAux0 = RES.ArraySlice(VecAux0, (GLOSNG_PUR_LIMIT - i) - 1, RBooGetLeftSide:=True)
            If Not VBA.IsArray(VecAux0) And k = 0 Then Call RES.FrmCMsgbox("La base de datos ya alcanzó el límite de " & GLOSNG_PUR_LIMIT & " registros.", STRAPPTITLE, vbInformation, True, .Controls("Btnpur5"), vbCritical): GoTo EH
            If GLORNG_CELL.Columns.Count >= 10 And k = 0 Then
              Let VecAux0 = RES.ArraySlice(VecAux0, 9, RByCols:=True, RBooGetLeftSide:=True)
              Let VecAux1 = RES.ArraySlice(VecAux0, 4, RByCols:=True) ', RBooGetLeftSide:=True)
              Let BooAux0 = True
              Let BooAux1 = True
              Let BooAux2 = True
            ElseIf GLORNG_CELL.Columns.Count >= 5 Then
              If k = 0 Then Let VecAux0 = RES.ArraySlice(VecAux0, 4, RByCols:=True, RBooGetLeftSide:=True)
              If k = 1 Then Let VecAux1 = RES.ArraySlice(VecAux0, 4, RByCols:=True, RBooGetLeftSide:=True)
              Let BooAux0 = (k = 0)
              Let BooAux1 = (k = 1)
              Let BooAux2 = False
            Else
              Call RES.FrmCMsgbox("Como mínimo cinco columnas" & IIf(k = 0, ", como máximo diez.", "."), STRAPPTITLE, vbInformation, True, .Controls("Btnpur5"), vbCritical): GoTo EH
            End If
            ' Validate each array into its corresponding position (0 or 1)
            If BooAux0 Then ' Base data (0)

              If i > 0 Then Let VecAux2 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur, eqlMdlArray, "MUESTRA"), eqlRes1D)
              For l = LBound(VecAux0) To UBound(VecAux0)
                ' MUESTRA
                If i > 0 Then
                  Let StrAux0 = VBA.CStr(VecAux0(l, 0))
                  Let StrAux1 = GLORNG_CELL.Cells(l + 1, 1).Address(False, False)
                  If RES.REGEX.isStringReg(StrAux0) Then
                    If RES.ArrayHasValue(VecAux2, VBA.LCase(StrAux0), ALCase:=True) Then _
                      Call RES.FrmCMsgbox("Celda: " & StrAux1 & VBA.vbNewLine & "¡La Muestra: " & StrAux0 & " ya existe en la base de datos!", STRAPPTITLE, vbInformation, True, .Controls("Btnpur5"), vbCritical): GoTo EH
                  End If
                End If
                ' FECHA
                Let StrAux0 = VBA.CStr(VecAux0(l, 1))
                Let StrAux1 = GLORNG_CELL.Cells(l + 1, 2).Address(False, False)
                If Not RES.REGEX.isDateReg(StrAux0) And Not VBA.IsDate(StrAux0) Then Call RES.FrmCMsgbox("Celda: " & StrAux1 & VBA.vbNewLine & "¡Fecha inválida! (" & StrAux0 & ")", STRAPPTITLE, vbInformation, True, .Controls("Btnpur5"), vbCritical): GoTo EH
                ' %HUMEDAD
                Let StrAux0 = VBA.CStr(VecAux0(l, 2))
                Let StrAux1 = GLORNG_CELL.Cells(l + 1, 3).Address(False, False)
                If Not RES.REGEX.isPercentReg(StrAux0) Then Call RES.FrmCMsgbox("Celda: " & StrAux1 & VBA.vbNewLine & "¡Porcentaje de Humedad inválido! (" & StrAux0 & ")", STRAPPTITLE, vbInformation, True, .Controls("Btnpur5"), vbCritical): GoTo EH
                ' %MERMA
                Let StrAux0 = VBA.CStr(VecAux0(l, 3))
                Let StrAux1 = GLORNG_CELL.Cells(l + 1, 4).Address(False, False)
                If Not RES.REGEX.isPercentReg(StrAux0) Then Call RES.FrmCMsgbox("Celda: " & StrAux1 & VBA.vbNewLine & "¡Porcentaje de Merma inválido! (" & StrAux0 & ")", STRAPPTITLE, vbInformation, True, .Controls("Btnpur5"), vbCritical): GoTo EH
                ' PBH
                Let StrAux0 = VBA.CStr(VecAux0(l, 4))
                Let StrAux1 = GLORNG_CELL.Cells(l + 1, 5).Address(False, False)
                If Not RES.REGEX.isNumberReg(StrAux0) Then Call RES.FrmCMsgbox("Celda: " & StrAux1 & VBA.vbNewLine & "¡Peso Bruto Húmedo inválido! (" & StrAux0 & ")", STRAPPTITLE, vbInformation, True, .Controls("Btnpur5"), vbCritical): GoTo EH
              Next l
              Let VecAux2 = Empty

            End If

            If BooAux1 Then ' Main data (1)

              If i <= 0 And Not BooAux0 Then Call RES.FrmCMsgbox("Primero ingrese datos en la opción: " & VBA.CStr(.Controls("Cbopur1").List(0, 0)) & ".", STRAPPTITLE, vbInformation, True, .Controls("Cbopur1"), vbCritical): GoTo EH
              ' Check if table has 0's in PARTE2 column, if it has, update second part, else cancel process
              Let k = 1
              If i > 0 Then
                Let VecAux0 = Empty
                Let VecAux0 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur, eqlMdlArray, "PARTE2"), eqlRes1D)
                Let k = RES.ArrayGetValuePosition(VecAux0, 0) + 1
                Let VecAux0 = Empty
              End If
              ' Table has all rows updated
              If k >= GLOSNG_PUR_LIMIT Then Call RES.FrmCMsgbox("La base de datos ya alcanzó el límite de " & GLOSNG_PUR_LIMIT & " registros.", STRAPPTITLE, vbInformation, True, .Controls("Btnpur5"), vbCritical): GoTo EH
              If ((GLOSNG_PUR_LIMIT - k) - 1) < UBound(VecAux1) Then Let VecAux1 = RES.ArraySlice(VecAux1, ((GLOSNG_PUR_LIMIT - k) - 1), RBooGetLeftSide:=True) ' Slice array 'cause have more data than expected
              If Not VBA.IsArray(VecAux1) Then Call RES.FrmCMsgbox("La base de datos ya alcanzó el límite de " & GLOSNG_PUR_LIMIT & " registros.", STRAPPTITLE, vbInformation, True, .Controls("Btnpur5"), vbCritical): GoTo EH
              ' Table has rows for update
              For i = LBound(VecAux1) To UBound(VecAux1)
                ' TIPO
                Let StrAux0 = VBA.CStr(VecAux1(i, 0))
                Let StrAux1 = GLORNG_CELL.Cells(i + 1, IIf(BooAux2, 6, 1)).Address(False, False)
                If Not RES.REGEX.isMineralType(StrAux0) Then Call RES.FrmCMsgbox("Celda: " & StrAux1 & VBA.vbNewLine & "¡Tipo mineral inválido! (" & StrAux0 & ")" & VBA.vbNewLine & "Sólo se admiten los siguientes tipos: Minera, Complejo, Brosa.", STRAPPTITLE, vbInformation, True, .Controls("Btnpur5"), vbCritical): GoTo EH
                Let BooAux0 = RES.REGEX.isMineralComplex(StrAux0) ' Complejos
                ' CONTENIDOS
                Let StrAux0 = VBA.CStr(VecAux1(i, 1))
                Let StrAux1 = GLORNG_CELL.Cells(i + 1, IIf(BooAux2, 7, 2)).Address(False, False)
                Let VecAux0 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(StrAux0, l, RValueType:=eqlResNotNumbers, RValueAs:=eqlResNotNumbers)
                If l <= 0 Then Call RES.FrmCMsgbox("Celda: " & StrAux1 & VBA.vbNewLine & "¡Se requieren los nombres de elementos o minerales! (" & StrAux0 & ")", STRAPPTITLE, vbInformation, True, .Controls("Btnpur5"), vbCritical): GoTo EH
                If BooAux0 And l <= 1 Then Call RES.FrmCMsgbox("Celda: " & StrAux1 & VBA.vbNewLine & "¡Se requiere como mínimo dos elementos o minerales para especies complejas! (" & StrAux0 & ")", STRAPPTITLE, vbInformation, True, .Controls("Btnpur5"), vbCritical): GoTo EH
                If RES.ArrayHasRepeatedElements(VecAux0) Then Call RES.FrmCMsgbox("Celda: " & StrAux1 & VBA.vbNewLine & "¡No pueden existir elementos repetidos para una liquidación! (" & StrAux0 & ")", STRAPPTITLE, vbInformation, True, .Controls("Btnpur5"), vbCritical): GoTo EH
                Let VecAux2 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalelements_sys, eqlMdlArray, "SÍMBOLO"), eqlRes1D)
                For j = LBound(VecAux0) To UBound(VecAux0)
                  If Not RES.ArrayHasValue(VecAux2, VBA.LCase(VBA.CStr(VecAux0(j))), ALCase:=True) Then _
                    Call RES.FrmCMsgbox("Celda: " & StrAux1 & VBA.vbNewLine & "¡No se reconoce al elemento o mineral: '" & VBA.CStr(VecAux0(j)) & "'. No existe en la base de datos! (" & StrAux0 & ")", STRAPPTITLE, vbInformation, True, .Controls("Btnpur5"), vbCritical): GoTo EH
                Next j
                Let VecAux0 = Empty
                Let VecAux2 = Empty
                ' LEYES
                Let StrAux0 = VBA.CStr(VecAux1(i, 2))
                Let StrAux1 = GLORNG_CELL.Cells(i + 1, IIf(BooAux2, 8, 3)).Address(False, False)
                Let VecAux0 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(StrAux0, m, RValueType:=eqlResNumbers, RBooAbs:=True, RValueAs:=eqlResNumbers, RDblTop:=CDBL_MAX)
                If m <= 0 Then Call RES.FrmCMsgbox("Celda: " & StrAux1 & VBA.vbNewLine & "¡Se requieren los nombres de elementos o minerales!", STRAPPTITLE, vbInformation, True, .Controls("Btnpur5"), vbCritical): GoTo EH
                If Not m = l Then Call RES.FrmCMsgbox("Celda: " & StrAux1 & VBA.vbNewLine & "¡La cantidad de leyes debe ser la misma que de elementos! (" & StrAux0 & ")", STRAPPTITLE, vbInformation, True, .Controls("Btnpur5"), vbCritical): GoTo EH
                Let VecAux0 = Empty
                ' UNIDADES
                Let StrAux0 = VBA.CStr(VecAux1(i, 3))
                Let StrAux1 = GLORNG_CELL.Cells(i + 1, IIf(BooAux2, 9, 4)).Address(False, False)
                Let VecAux0 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(StrAux0, n, RValueType:=eqlResNotNumbers, RValueAs:=eqlResNotNumbers)
                If n <= 0 Then Call RES.FrmCMsgbox("Celda: " & StrAux1 & VBA.vbNewLine & "¡Se requieren unidades de elementos o minerales!", STRAPPTITLE, vbInformation, True, .Controls("Btnpur5"), vbCritical): GoTo EH
                If Not n = l Then Call RES.FrmCMsgbox("Celda: " & StrAux1 & VBA.vbNewLine & "¡La cantidad de unidades debe ser la misma que de elementos! (" & StrAux0 & ")", STRAPPTITLE, vbInformation, True, .Controls("Btnpur5"), vbCritical): GoTo EH
                Let VecAux2 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalunits_sys, eqlMdlArray, "Unidad"), eqlRes1D)
                For j = LBound(VecAux0) To UBound(VecAux0)
                  If Not RES.ArrayHasValue(VecAux2, VBA.LCase(VBA.CStr(VecAux0(j))), ALCase:=True) Then _
                    Call RES.FrmCMsgbox("Celda: " & StrAux1 & VBA.vbNewLine & "¡No se reconoce la unidad: '" & VBA.CStr(VecAux0(j)) & "'. No existe en la base de datos! (" & StrAux0 & ")", STRAPPTITLE, vbInformation, True, .Controls("Btnpur5"), vbCritical): GoTo EH
                Next j
                Let VecAux2 = Empty
                ' Check for DM's & Ag's in elements & units lists
                Let VecAux2 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(VBA.CStr(VecAux1(i, 1)), n, RValueType:=eqlResNotNumbers, RValueAs:=eqlResNotNumbers)
                If Not RES.RESCORE.BalCheckCustomNamesDoesNotHasWrongUnits(RES, StrAux2, Empty, VecAux2, VecAux0, eqlBalAgUnit) Then Call RES.FrmCMsgbox("Celda: " & StrAux1 & VBA.vbNewLine & StrAux2, STRAPPTITLE, vbInformation, True, .Controls("Btnpur5"), vbCritical): GoTo EH
                If Not RES.RESCORE.BalCheckCustomNamesDoesNotHasWrongUnits(RES, StrAux2, Empty, VecAux2, VecAux0, eqlBalDMOnOtherElements) Then Call RES.FrmCMsgbox("Celda: " & StrAux1 & VBA.vbNewLine & StrAux2, STRAPPTITLE, vbInformation, True, .Controls("Btnpur5"), vbCritical): GoTo EH
                Let VecAux2 = Empty
                Let VecAux2 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(VBA.CStr(VecAux1(i, 2)), m, RValueType:=eqlResNumbers, RBooAbs:=True, RValueAs:=eqlResNumbers, RDblTop:=CDBL_MAX)
                Let VecAux2 = RES.ArrayToBaseZero(VecAux2, eqlRes2D)
                If Not RES.RESCORE.BalCheckForMaximousValuesInGrades(StrAux2, VecAux2, VecAux0, BBooOneToOne:=True) Then Call RES.FrmCMsgbox("Celda: " & StrAux1 & VBA.vbNewLine & StrAux2, STRAPPTITLE, vbInformation, True, .Controls("Btnpur5"), vbCritical): GoTo EH
                Let VecAux0 = Empty
                Let VecAux2 = Empty
                ' PRECIO/PTO
                Let StrAux0 = VBA.CStr(VecAux1(i, 4))
                Let StrAux1 = GLORNG_CELL.Cells(i + 1, IIf(BooAux2, 10, 5)).Address(False, False)
                Let VecAux0 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(StrAux0, m, RValueType:=eqlResNumbers, RBooAbs:=True, RValueAs:=eqlResNumbers, RDblTop:=CDBL_MAX)
                If m <= 0 Then Call RES.FrmCMsgbox("Celda: " & StrAux1 & VBA.vbNewLine & "¡Se requieren los precios de minerales!", STRAPPTITLE, vbInformation, True, .Controls("Btnpur5"), vbCritical): GoTo EH
                If Not m = l Then Call RES.FrmCMsgbox("Celda: " & StrAux1 & VBA.vbNewLine & "¡La cantidad de precios de minerales debe ser la misma que de elementos! (" & StrAux0 & ")", STRAPPTITLE, vbInformation, True, .Controls("Btnpur5"), vbCritical): GoTo EH
                Let VecAux0 = Empty
              Next i

            End If

          Case 2, 5 ' Bonus & Discounts
            If GLORNG_CELL.Columns.Count < 2 Then Call RES.FrmCMsgbox("Seleccione dos columnas.", STRAPPTITLE, vbInformation, True, .Controls("Btnpur5"), vbCritical): GoTo EH
            Let i = MDL.LengthRows(eqlMdlSheetTmp, IIf(k = 2, eqlMdlTblPurbonus_tmp, eqlMdlTblPurdiscounts_tmp))
            Let j = GLORNG_CELL.Rows.Count
            If i = CSNG_BONUS_DISCOUNT_LIMIT Then Call RES.FrmCMsgbox("La tabla '" & .Controls("Cbopur1").Value & "' ya alcanzó el límite de " & CSNG_BONUS_DISCOUNT_LIMIT & " registros.", STRAPPTITLE, vbInformation, True, .Controls("Btnpur5"), vbCritical): GoTo EH
            Let l = CSNG_BONUS_DISCOUNT_LIMIT - i
            Let VecAux0 = GLORNG_CELL.Value
            If j > l Then Let VecAux0 = RES.ArraySlice(VecAux0, l - 1, RBooGetLeftSide:=True)
            If GLORNG_CELL.Columns.Count > 2 Then Let VecAux0 = RES.ArraySlice(VecAux0, 1, RByCols:=True, RBooGetLeftSide:=True)
            ' Validate
            For i = LBound(VecAux0) To UBound(VecAux0)
              ' BONO|DESCUENTO
              ' VALOR
              Let StrAux0 = VBA.CStr(VecAux0(i, 1))
              Let StrAux1 = GLORNG_CELL.Cells(i, 2).Address(False, False)
              If Not RES.REGEX.isNumberReg(StrAux0) Then Call RES.FrmCMsgbox("Celda: " & StrAux1 & VBA.vbNewLine & "¡" & IIf(j = 2, "Bono", "Descuento") & " inválido!", STRAPPTITLE, vbInformation, True, .Controls("Btnpur5"), vbCritical): GoTo EH
            Next i

          Case Else: GoTo EH

        End Select
        Let View_Core_validations_oreinvoice = True

      Case "Btnmain1"
        Let i = MDL.LengthRows(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur)
        ' Provider type
        If .Controls("Chkpur1").Value And i > 0 Then
          If Not RES.REGEX.isProvidersTypes(.Controls("Txtpur4").Text) Then Call RES.FrmCMsgbox("¡Dato incorrecto!" & VBA.vbNewLine & "Debe ingresar un tipo de provedor para cada liquidación, cada tipo debe ir separado por punto y coma, Ej. Particular;Cooperativista;Particular", STRAPPTITLE, vbInformation, True, .Controls("Txtpur4"), vbCritical): GoTo EH
          Let StrAux0 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(.Controls("Txtpur4").Text, j, RValueType:=eqlResNotNumbers)
          If Not i = j Then Call RES.FrmCMsgbox("¡Dato incorrecto!" & VBA.vbNewLine & "Se requiere un tipo de provedor por cada liquidación de minerales", STRAPPTITLE, vbInformation, True, .Controls("Txtpur4"), vbCritical): GoTo EH
        End If
        ' Base Data
        If i <= 0 Then Call RES.FrmCMsgbox("Debe ingresar todos los datos de liquidación de minerales", STRAPPTITLE, vbInformation, True, .Controls("Lstpur0"), vbCritical): GoTo EH
        ' Main Data
        Let VecAux0 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur, eqlMdlArray, "PARTE2"), eqlRes1D)
        If RES.ArrayHasValue(VecAux0, "0") Then Call RES.FrmCMsgbox("Debe ingresar todos los datos de liquidación de minerales", STRAPPTITLE, vbInformation, True, .Controls("Lstpur0"), vbCritical): GoTo EH
        Let VecAux0 = Empty

        ' Retentions & Retentions others
        Let VecAux0 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur, eqlMdlArray, "RETENCIONES"), eqlRes1D)
        If .Controls("Chkpur1").Value Then
          If RES.ArrayHasValue(VecAux0, VBA.vbNullString) Then Call RES.FrmCMsgbox("Las retenciones de ley son obligatorias, existen liquidaciones sin retenciones establecidas.", STRAPPTITLE, vbInformation, True, .Controls("Lstpur0"), vbCritical): GoTo EH
        Else
          If VBA.CStr(VecAux0(0)) = VBA.vbNullString Then Call RES.FrmCMsgbox("Las retenciones de ley son obligatorias, existen liquidaciones sin retenciones establecidas.", STRAPPTITLE, vbInformation, True, .Controls("Lstpur0"), vbCritical): GoTo EH
          Let StrAux0 = VBA.CStr(VecAux0(0))
          If Not VBA.InStr(StrAux0, "1") > 0 Then Call RES.FrmCMsgbox("Las retenciones de ley son obligatorias, existen liquidaciones sin retenciones establecidas.", STRAPPTITLE, vbInformation, True, .Controls("Lstpur0"), vbCritical): GoTo EH
        End If
        Let VecAux0 = Empty
        Let VecAux0 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur, eqlMdlArray, "RETENCIONES OTRAS"), eqlRes1D)
        If .Controls("Chkpur1").Value Then
          Let VecAux1 = VBA.Split(.Controls("Txtpur4").Value, ";")
        Else
          Let VecAux1 = Array(.Controls("Cbopur0").Value)
        End If
        For n = 0 To UBound(VecAux1)
          If VBA.LCase(VBA.CStr(VecAux1(n))) = "cooperativista" Then
            If .Controls("Chkpur1").Value Then ' Lot providers
              If VBA.CStr(VecAux0(n)) = VBA.vbNullString Then
                If VBA.MsgBox("Las retenciones institucionales son requeridas en el caso de cooperativistas mineros, existen liquidaciones sin retenciones establecidas." & VBA.vbNewLine & "¿Realizar las retenciones?", vbInformation + vbYesNo, STRAPPTITLE) = vbYes Then
                  GoTo EH
                Else
                  Exit For
                End If
              Else
                Let StrAux0 = VBA.CStr(VecAux0(n))
                If Not VBA.InStr(StrAux0, "1") > 0 Then
                  If VBA.MsgBox("Las retenciones institucionales son requeridas en el caso de cooperativistas mineros, existen liquidaciones sin retenciones establecidas." & VBA.vbNewLine & "¿Realizar las retenciones?", vbInformation + vbYesNo, STRAPPTITLE) = vbYes Then
                    GoTo EH
                  Else
                    Exit For
                  End If
                End If
              End If
            Else ' Single provider
              If VBA.CStr(VecAux0(0)) = VBA.vbNullString Then
                If VBA.MsgBox("Las retenciones institucionales son requeridas en el caso de cooperativistas mineros, esta liquidación está sin retenciones establecidas." & VBA.vbNewLine & "¿Realizar las retenciones?", vbInformation + vbYesNo, STRAPPTITLE) = vbYes Then GoTo EH
              Else
                Let StrAux0 = VBA.CStr(VecAux0(0))
                If Not VBA.InStr(StrAux0, "1") > 0 Then
                  If VBA.MsgBox("Las retenciones institucionales son requeridas en el caso de cooperativistas mineros, esta liquidación está sin retenciones establecidas." & VBA.vbNewLine & "¿Realizar las retenciones?", vbInformation + vbYesNo, STRAPPTITLE) = vbYes Then GoTo EH
                End If
              End If
            End If
          End If
        Next n
        Let VecAux0 = Empty
        Let VecAux1 = Empty

        ' Check for spreadsheet edges and data inside the region selected
        If .Controls("Optpur1").Value Then
          Let j = 80 * IIf(i <= 5, 1, 2) + 50 ' 80: A1 to totalsrow aprox
          Let k = 9 * IIf(.Controls("Chkpur5").Value, 2, 1) * IIf(.Controls("Chkpur1").Value, i, 1) + 50
          If GLORNG_CELL Is Nothing Then Call CMB.CMsgBox("¡Seleccione una celda!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtpur10", "Btnmain1", cmbRequired, cmbRight, cmbBottom, cAspect:=cmbTagLine): GoTo EH
          If RES.ExcelIsProtectedSheet(GLORNG_CELL.Parent) Then Call RES.FrmCMsgbox("La celda que seleccionó para imprimir los resultados está en una hoja protegida. Para realizar un cambio, debe quitar la protección de la hoja. Es posible que se le pida que escriba una contraseña.", STRAPPTITLE, vbInformation, True, .Controls("Txtpur10")): GoTo EH
          If Not View_assets_check_for_spreadsheet_edges(GLORNG_CELL, j, k, StrMsg0) Then Call RES.FrmCMsgbox(StrMsg0, STRAPPTITLE, vbInformation, True, .Controls("Txtpur10")): GoTo EH
        Else
          If Application.ActiveWorkbook.ProtectStructure Then Call RES.FrmCMsgbox("No se puede crear una nueva hoja en el libro actual porque está protegido. Quite la protección o seleccione otro libro para continuar.", STRAPPTITLE, vbInformation, True, .Controls("Optpro0")): GoTo EH
          Set GLORNG_CELL = Nothing
        End If

        If .Controls("Chkpur1").Value Then Let StrMsg0 = "Se generarán las Liquidaciones de minerales con los datos que ingresó." & VBA.vbNewLine & "¿Proceder?"
        If Not .Controls("Chkpur1").Value Then Let StrMsg0 = "Se generará la Liquidación de minerales con los datos que ingresó." & VBA.vbNewLine & "¿Proceder?"
        Let View_Core_validations_oreinvoice = (VBA.MsgBox(StrMsg0, vbQuestion + vbYesNo, STRAPPTITLE) = vbYes)

      Case Else: GoTo EH

    End Select
  End With

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Let VecAux2 = Empty
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function View_Core_validations_concentrateinvoice( _
  ByVal VBtnCurrent As MSForms.CommandButton, _
  ByVal FRM As FRMBasics) As Boolean

  Dim VecAux0 As Variant, VecAux1 As Variant, VecAux2 As Variant, VecAux3 As Variant
  Dim StrAux0 As String, StrAux1 As String, StrAux2 As String
  Dim i As Long, j As Long, k As Long
  Dim BooAux0 As Boolean
  Dim DblAux0 As Double, DblAux1 As Double
  Dim EnuTbl As EQLMDL_ENU_TABLES

  Const sFN As String = "VIEW::-Validations_concentrateinvoice"

  On Error GoTo EH
  Let View_Core_validations_concentrateinvoice = False

  With FRM
    Select Case VBtnCurrent.Name

      Case "Btncnc0", "Btncnc1" ' Set & Upd
        Select Case .Controls("Cbocnc1").ListIndex

          Case 0
            ' Contents
            Let j = MDL.LengthRows(eqlMdlSheetPur, eqlMdlTblpurchasecnccontents_pur)
            If j <= 0 And VBtnCurrent.Name = "Btncnc1" Then GoTo EH
            If j >= GLOSNG_PUR_LIMIT Then Call RES.FrmCMsgbox("¡Ya se alcanzó el límite de " & GLOSNG_PUR_LIMIT & " registros!", STRAPPTITLE, VBA.vbInformation, True, .Controls("Lstcnc0")): GoTo EH
            Let i = IIf(VBtnCurrent.Name = "Btncnc0", 0, 1)
            If j > 0 Then
              Let VecAux0 = RES.ArrayToBaseZero(RES.ArrayShift(.Controls("Lstcnc0").List, True, True), eqlRes1D)
              Let j = RES.ArrayHasValueN(VecAux0, VBA.LCase(.Controls("Cbocnc2").Value), ALCase:=True)
              If j > i Then Call CMB.CMsgBox("¡El Mineral: " & .Controls("Cbocnc2").Value & " ya existe en los registros!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Cbocnc2", "Cbocnc1;Cbocnc2;Btnmain1", cmbRequired, cmbLeft, cmbBottom, cAspect:=cmbTagLine): GoTo EH
            End If
            ' Grade
            If Not RES.REGEX.isNumberReg(.Controls("Txtcnc11").Text) Then Call CMB.CMsgBox("¡Falta la Ley del Mineral!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtcnc11", "Cbocnc1;Cbocnc2;Btnmain1", cmbRequired, cmbLeft, cmbBottom, cAspect:=cmbTagLine): GoTo EH
            If VBA.Abs(VBA.CDbl(.Controls("Txtcnc11").Value)) = 0 Then Call CMB.CMsgBox("¡Dato incorrecto!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtcnc11", "Cbocnc1,Cbocnc2,Btnmain1", cmbRequired, cmbLeft, cmbBottom, cAspect:=cmbTagLine): GoTo EH
            ' Min
            If Not RES.REGEX.isNumberReg(.Controls("Txtcnc12").Value) Then Call CMB.CMsgBox("¡Deducción mínima inválida!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtcnc12", "Cbocnc1;Cbocnc2;Btnmain1", cmbRequired, cmbLeft, cmbBottom, cAspect:=cmbTagLine): GoTo EH
            ' %Payable
            If Not RES.REGEX.isPercentReg(.Controls("Txtcnc14").Value) Then Call CMB.CMsgBox("¡Porcentaje Pagable inválido!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtcnc14", "Cbocnc1;Cbocnc2;Btnmain1", cmbRequired, cmbRight, cmbBottom, cAspect:=cmbTagLine): GoTo EH
            If VBA.Abs(VBA.CDbl(.Controls("Txtcnc14").Value)) = 0 Or VBA.Abs(VBA.CDbl(.Controls("Txtcnc14").Value)) > 100 Then Call CMB.CMsgBox("¡Dato incorrecto!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtcnc14", "Cbocnc1;Cbocnc2;Btnmain1", cmbRequired, cmbLeft, cmbBottom, cAspect:=cmbTagLine): GoTo EH

          Case 1 ' Refinement - just edition
            If VBtnCurrent.Name = "Btncnc1" Then
              ' Condition '2.
              Let BooAux0 = RES.REGEX.isNumberReg(.Controls("Txtcnc11").Value)
              If Not BooAux0 Then Call CMB.CMsgBox("¡La condición es inválida!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtcnc11", "Cbocnc1;Cbocnc2;Btnmain1", cmbRequired, cmbLeft, cmbBottom, cAspect:=cmbTagLine): GoTo EH
              ' %Base|Cur/T '1.
              If BooAux0 Then
                If Not RES.REGEX.isPercentReg(.Controls("Txtcnc12").Value) Then Call CMB.CMsgBox("¡Porcentaje inválido!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtcnc12", "Cbocnc1;Cbocnc2;Btnmain1", cmbRequired, cmbRight, cmbBottom, cAspect:=cmbTagLine): GoTo EH
                If VBA.Abs(VBA.CDbl(.Controls("Txtcnc12").Value)) = 0 Or VBA.Abs(VBA.CDbl(.Controls("Txtcnc12").Value)) > 100 Then Call CMB.CMsgBox("¡Porcentaje inválido!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtcnc12", "Cbocnc1;Cbocnc2;Btnmain1", cmbRequired, cmbRight, cmbBottom, cAspect:=cmbTagLine): GoTo EH
              Else
                If Not RES.REGEX.isNumberReg(.Controls("Txtcnc12").Value) Then Call CMB.CMsgBox("¡Base inválida!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtcnc12", "Cbocnc1;Cbocnc2;Btnmain1", cmbRequired, cmbRight, cmbBottom, cAspect:=cmbTagLine): GoTo EH
              End If
            End If

          Case 2 ' Penalties: just in edition
              If VBtnCurrent.Name = "Btncnc1" Then
              If MDL.LengthRows(eqlMdlSheetPur, eqlMdlTblpurchasecncpenalties_pur) > GLOSNG_PUR_PEN_LIMIT Then Call RES.FrmCMsgbox("¡Ya fue alcanzado el límite de " & GLOSNG_PUR_PEN_LIMIT & " registros!", STRAPPTITLE, VBA.vbInformation, True, .Controls("Lstcnc0")): GoTo EH
              ' Current
              If Not RES.REGEX.isNumberReg(.Controls("Txtcnc11").Value) Then Call CMB.CMsgBox("¡Ley actual del elemento inválida!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtcnc11", "Cbocnc1;Cbocnc2;Btnmain1", cmbRequired, cmbLeft, cmbBottom, cAspect:=cmbTagLine): GoTo EH
              ' Base
              If Not RES.REGEX.isNumberReg(.Controls("Txtcnc12").Value) Then Call CMB.CMsgBox("¡Ley Base del elemento inválida!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtcnc12", "Cbocnc1;Cbocnc2;Btnmain1", cmbRequired, cmbLeft, cmbBottom, cAspect:=cmbTagLine): GoTo EH
              ' Deduction
              If Not RES.REGEX.isNumberReg(.Controls("Txtcnc13").Value) Then Call CMB.CMsgBox("¡Cantidad deducible inválida!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtcnc13", "Cbocnc1;Cbocnc2;Btnmain1", cmbRequired, cmbRight, cmbBottom, cAspect:=cmbTagLine): GoTo EH
              ' Limit
              If Not RES.REGEX.isPercentReg(.Controls("Txtcnc14").Value) Then Call CMB.CMsgBox("¡Porcentaje de deducción inválido!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtcnc14", "Cbocnc1;Cbocnc2;Btnmain1", cmbRequired, cmbRight, cmbBottom, cAspect:=cmbTagLine): GoTo EH
            End If

          Case 3 ' Deductions
            Select Case .Controls("Cbocnc2").ListIndex
              Case 0 ' Deductions: by exportations just edition
                If VBtnCurrent.Name = "Btncnc1" Then
                  ' Percent
                  If Not RES.REGEX.isPercentReg(.Controls("Txtcnc12").Text) Then Call CMB.CMsgBox("¡Porcentaje inválido!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtcnc12", "Cbocnc1;Cbocnc2;Btnmain1", cmbRequired, cmbLeft, cmbBottom, cAspect:=cmbTagLine): GoTo EH
                  Let DblAux0 = VBA.Abs(VBA.CDbl(.Controls("Txtcnc12").Value))
                  ' Value
                  If Not RES.REGEX.isNumberReg(.Controls("Txtcnc13").Text) Then Call CMB.CMsgBox("¡Valor incorrecto!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtcnc13", "Cbocnc1;Cbocnc2;Btnmain1", cmbRequired, cmbLeft, cmbBottom, cAspect:=cmbTagLine): GoTo EH
                  Let DblAux1 = VBA.Abs(VBA.CDbl(.Controls("Txtcnc13").Text))

                  If (DblAux0 = 0 And DblAux1 = 0) Then Call RES.FrmCMsgbox("¡Se esperaba un porcentaje o un valor!", STRAPPTITLE, VBA.vbInformation, True, .Controls("Btncnc1")): GoTo EH
                  If (DblAux0 > 0 And DblAux1 > 0) Then Call RES.FrmCMsgbox("¡Sólo puede habilitar un porcentaje o un valor!", STRAPPTITLE, VBA.vbInformation, True, .Controls("Btncnc1")): GoTo EH
                End If

              Case 3 ' Deductions: by discounts
                ' Concept
                Let i = MDL.LengthRows(eqlMdlSheetTmp, eqlMdlTblPurdiscounts_tmp)
                If i <= 0 And VBtnCurrent.Name = "Btncnc1" Then GoTo EH
                If i > GLOSNG_PUR_LIMIT Then Call RES.FrmCMsgbox("¡La lista de descuentos ya llegó al límite de " & GLOSNG_PUR_LIMIT & " registros!", STRAPPTITLE, VBA.vbInformation, True, .Controls("Lstcnc0")): GoTo EH
                Let j = IIf(VBtnCurrent.Name = "Btncnc2", 0, 1)
                If RES.REGEX.isEmptyStringReg(.Controls("Txtcnc11").Text) Then Call CMB.CMsgBox("¡Concepto de descuento inválido!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtcnc11", "Cbocnc1;Cbocnc2;Btnmain1", cmbRequired, cmbLeft, cmbBottom, cAspect:=cmbTagLine): GoTo EH
                If i > 0 Then
                  Let VecAux0 = RES.ArrayToBaseZero(RES.ArrayDelIndex(.Controls("Lstcnc0").List, 1, True, True), eqlRes1D)
                  Let i = RES.ArrayHasValueN(VecAux0, VBA.LCase(.Controls("Txtcnc11").Text), ALCase:=True)
                  If i > j Then Call CMB.CMsgBox("¡El descuento: '" & .Controls("Txtcnc11").Text & "' ya existe en los registros!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtcnc11", "Cbocnc1;Cbocnc2;Btnmain1", cmbRequired, cmbLeft, cmbBottom, cAspect:=cmbTagLine): GoTo EH
                  Let VecAux0 = Empty
                End If
                ' Value
                If Not RES.REGEX.isNumberReg(.Controls("Txtcnc12").Value) Then Call CMB.CMsgBox("¡Valor de descuento inválido!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtcnc12", "Cbocnc1;Cbocnc2;Btnmain1", cmbRequired, cmbRight, cmbBottom, cAspect:=cmbTagLine): GoTo EH
                Let DblAux0 = VBA.Abs(VBA.CDbl(.Controls("Txtcnc12").Value))
                If (DblAux0 = 0) Then Call CMB.CMsgBox("¡Valor de descuento inválido!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtcnc12", "Cbocnc1;Cbocnc2;Btnmain1", cmbRequired, cmbRight, cmbBottom, cAspect:=cmbTagLine): GoTo EH

              Case Else: GoTo EH
            End Select

          Case 4 ' Bonus
            ' Concept
            Let i = MDL.LengthRows(eqlMdlSheetTmp, eqlMdlTblPurbonus_tmp)
            If i <= 0 And VBtnCurrent.Name = "Btncnc1" Then GoTo EH
            If i > GLOSNG_PUR_LIMIT Then Call RES.FrmCMsgbox("¡La lista de bonos ya llegó al límite de " & GLOSNG_PUR_LIMIT & " registros!", STRAPPTITLE, VBA.vbInformation, True, .Controls("Lstcnc0")): GoTo EH
            Let j = IIf(VBtnCurrent.Name = "Btncnc2", 0, 1)
            If RES.REGEX.isEmptyStringReg(.Controls("Txtcnc11").Text) Then Call CMB.CMsgBox("¡Concepto de bono inválido!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtcnc11", "Cbocnc1;Cbocnc2;Btnmain1", cmbRequired, cmbLeft, cmbBottom, cAspect:=cmbTagLine): GoTo EH
            If i > 0 Then
              Let VecAux0 = RES.ArrayToBaseZero(RES.ArrayDelIndex(.Controls("Lstcnc0").List, 1, True, True), eqlRes1D)
              Let i = RES.ArrayHasValueN(VecAux0, VBA.LCase(.Controls("Txtcnc11").Text), ALCase:=True)
              If i > j Then Call CMB.CMsgBox("¡El bono: '" & .Controls("Txtcnc11").Text & "' ya existe en los registros!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtcnc11", "Cbocnc1;Cbocnc2;Btnmain1", cmbRequired, cmbLeft, cmbBottom, cAspect:=cmbTagLine): GoTo EH
              Let VecAux0 = Empty
            End If
            ' Value
            If Not RES.REGEX.isNumberReg(.Controls("Txtcnc12").Value) Then Call CMB.CMsgBox("¡Valor de bono inválido!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtcnc12", "Cbocnc1;Cbocnc2;Btnmain1", cmbRequired, cmbLeft, cmbBottom, cAspect:=cmbTagLine): GoTo EH
            Let DblAux0 = VBA.Abs(VBA.CDbl(.Controls("Txtcnc12").Value))
            If (DblAux0 = 0) Then Call CMB.CMsgBox("¡Valor de bono inválido!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtcnc12", "Cbocnc1;Cbocnc2;Btnmain1", cmbRequired, cmbLeft, cmbBottom, cAspect:=cmbTagLine): GoTo EH

          Case Else: GoTo EH
        End Select
        Let View_Core_validations_concentrateinvoice = True

      Case "Btncnc3", "Btncnc4" ' Del once, del all
        If VBtnCurrent.Name = "Btncnc3" And .Controls("Lstcnc0").ListIndex < 0 Then GoTo EH
        Let i = .Controls("Cbocnc1").ListIndex
        Let j = .Controls("Cbocnc2").ListIndex
        Let StrAux0 = "Eliminará " & IIf(VBtnCurrent.Name = "Btncnc3", "el registro seleccionado ", "todos los datos ") & "de la tabla '" & _
           .Controls(IIf(i = 3, "Cbocnc2", "Cbocnc1")).Value & "'." & IIf(i = 0, VBA.vbNewLine & "(También será eliminado el contenido relacionado en la tabla '" & .Controls("Cbocnc1").List(1, 0) & "')", VBA.vbNullString)
        Let View_Core_validations_concentrateinvoice = (VBA.MsgBox(StrAux0 & VBA.vbNewLine & VBA.vbNewLine & "¿Realmente desea proceder con la eliminación de datos?", vbYesNo + vbQuestion, STRAPPTITLE) = vbYes)

      Case "Btncnc5"
        If GLORNG_CELL Is Nothing Then GoTo EH
        If GLORNG_CELL.Areas.Count > 1 Then Call RES.FrmCMsgbox("¡Realice solo una selección!", STRAPPTITLE, VBA.vbInformation, True, .Controls("Btncnc5")): GoTo EH 'Just one area
        Select Case .Controls("Cbocnc1").ListIndex

          Case 0 ' Contents
            Let i = MDL.LengthRows(eqlMdlSheetPur, eqlMdlTblpurchasecnccontents_pur)
            Let j = GLORNG_CELL.Rows.Count
            If i >= GLOSNG_PUR_LIMIT Then Call RES.FrmCMsgbox("¡La lista de registros ya llegó al límite de " & GLOSNG_PUR_LIMIT & "!", STRAPPTITLE, VBA.vbInformation, True, .Controls("Lstcnc0")): GoTo EH
            If (i + j) > GLOSNG_PUR_LIMIT Then Call RES.FrmCMsgbox("¡La lista de registros ya llegó al límite de " & GLOSNG_PUR_LIMIT & "!", STRAPPTITLE, VBA.vbInformation, True, .Controls("Lstcnc0")): GoTo EH
            If Not GLORNG_CELL.Columns.Count = 5 Then Call RES.FrmCMsgbox("La selección debe poseer sólo cinco columnas.", STRAPPTITLE, VBA.vbInformation, True, VBtnCurrent): GoTo EH
            If Application.WorksheetFunction.CountA(GLORNG_CELL) = 0 Then Call RES.FrmCMsgbox("¿Y los Datos?", STRAPPTITLE, VBA.vbInformation, True, VBtnCurrent): GoTo EH
            ' Search for duplicates
            Let VecAux0 = RES.ArrayToBaseZero(GLORNG_CELL.Columns("A:A").Value, eqlRes1D)
            If RES.ArrayHasRepeatedElements(VecAux0) Then Call RES.FrmCMsgbox("La selección contiene nombres de minerales duplicados. Es necesario que todos los nombres sean únicos.", STRAPPTITLE, VBA.vbInformation, True, VBtnCurrent): GoTo EH
            Let VecAux0 = Empty
            ' Validate data
            If i > 0 Then Let VecAux0 = RES.ArrayToBaseZero(RES.ArrayShift(.Controls("Cbocnc2").List, True, True), eqlRes1D)
            If i > 0 Then Let VecAux1 = RES.ArrayToBaseZero(RES.ArrayDelIndex(.Controls("Cbocnc2").List, 1, True, True), eqlRes1D)
            If i > 0 Then Let VecAux2 = RES.ArrayToBaseZero(RES.ArrayShift(.Controls("Lstcnc0").List, True, True), eqlRes1D)
            If i > 0 Then Let VecAux3 = RES.ArrayToBaseZero(RES.ArrayDelIndex(.Controls("Lstcnc0").List, 1, True, True), eqlRes1D)
            For j = 1 To GLORNG_CELL.Rows.Count
              ' Contents
              Let StrAux0 = GLORNG_CELL.Cells(j, 1).Value
              Let StrAux1 = "Celda '" & GLORNG_CELL.Cells(j, 1).Address(False, False) & "': "
              Let j = RES.ArrayHasValueN(VecAux0, VBA.LCase(StrAux0), ALCase:=True)
              Let k = RES.ArrayHasValueN(VecAux1, VBA.LCase(StrAux0), ALCase:=True)
              If j <= 0 And k <= 0 Then Call RES.FrmCMsgbox(StrAux1 & "¡El Mineral: '" & StrAux0 & "' no existe en la base de datos!", STRAPPTITLE, vbInformation, True, VBtnCurrent): GoTo EH
              If i > 0 Then
                Let j = RES.ArrayHasValueN(VecAux2, VBA.LCase(StrAux0), ALCase:=True)
                Let k = RES.ArrayHasValueN(VecAux3, VBA.LCase(StrAux0), ALCase:=True)
                If j > 0 Or k > 0 Then Call RES.FrmCMsgbox(StrAux1 & "¡El Mineral: '" & StrAux0 & "' ya existe en los registros!", STRAPPTITLE, vbInformation, True, VBtnCurrent): GoTo EH
              End If
              ' Grade
              Let StrAux0 = GLORNG_CELL.Cells(j, 2).Value
              Let StrAux1 = "Celda '" & GLORNG_CELL.Cells(j, 2).Address(False, False) & "': "
              If Not RES.REGEX.isNumberReg(StrAux0) Then Call RES.FrmCMsgbox(StrAux1 & "¡Ley de Mineral '" & StrAux0 & "' inválida!", STRAPPTITLE, vbInformation, True, VBtnCurrent): GoTo EH
              If VBA.Abs(VBA.CDbl(StrAux0)) = 0 Then Call RES.FrmCMsgbox(StrAux1 & "¡Ley de Mineral '" & StrAux0 & "' inválida!", STRAPPTITLE, vbInformation, True, VBtnCurrent): GoTo EH
              ' Min
              Let StrAux0 = GLORNG_CELL.Cells(j, 3).Value
              Let StrAux1 = "Celda '" & GLORNG_CELL.Cells(j, 3).Address(False, False) & "': "
              If Not RES.REGEX.isNumberReg(StrAux0) Then Call RES.FrmCMsgbox(StrAux1 & "¡Deducción mínima '" & StrAux0 & "' inválida!", STRAPPTITLE, vbInformation, True, VBtnCurrent): GoTo EH
              Let DblAux0 = VBA.CDbl(StrAux0)
              ' %Payable
              Let StrAux0 = GLORNG_CELL.Cells(j, 5).Value
              Let StrAux1 = "Celda '" & GLORNG_CELL.Cells(j, 5).Address(False, False) & "': "
              If Not RES.REGEX.isPercentReg(StrAux0) Then Call RES.FrmCMsgbox(StrAux1 & "¡Porcentaje pagable '" & StrAux0 & "' inválido!", STRAPPTITLE, vbInformation, True, VBtnCurrent): GoTo EH
              If VBA.Abs(VBA.CDbl(StrAux0)) = 0 Or VBA.Abs(VBA.CDbl(StrAux0)) > 100 Then Call RES.FrmCMsgbox(StrAux1 & "¡Porcentaje pagable '" & StrAux0 & "' inválido!", STRAPPTITLE, vbInformation, True, VBtnCurrent): GoTo EH
              Let DblAux1 = VBA.CDbl(StrAux0)

              If DblAux0 = 0 And DblAux1 = 0 Then Call RES.FrmCMsgbox(StrAux1 & "¡Deducción mínima y Porcentaje pagable inválidos. Como mínimo, uno debe ser mayor a cero.", STRAPPTITLE, vbInformation, True, VBtnCurrent): GoTo EH
            Next j
            Let VecAux0 = Empty
            Let VecAux1 = Empty
            Let VecAux2 = Empty
            Let VecAux3 = Empty

          Case 1 ' Refinement
            Let i = MDL.LengthRows(eqlMdlSheetPur, eqlMdlTblpurchasecnccontents_pur)
            Let j = GLORNG_CELL.Rows.Count
            If i >= GLOSNG_PUR_LIMIT Then Call RES.FrmCMsgbox("¡La lista de registros ya llegó al límite de " & GLOSNG_PUR_LIMIT & "!", STRAPPTITLE, VBA.vbInformation, True, .Controls("Lstcnc0")): GoTo EH
            If j > i Then Call RES.FrmCMsgbox("La selección debe contener como máximo " & i & " registros.", STRAPPTITLE, VBA.vbInformation, True, VBtnCurrent): GoTo EH
            If Not GLORNG_CELL.Columns.Count = 2 Then Call RES.FrmCMsgbox("La selección debe poseer sólo dos columnas.", STRAPPTITLE, VBA.vbInformation, True, VBtnCurrent): GoTo EH
            If Application.WorksheetFunction.CountA(GLORNG_CELL) = 0 Then Call RES.FrmCMsgbox("¿Y los Datos?", STRAPPTITLE, VBA.vbInformation, True, VBtnCurrent): GoTo EH

            For j = 1 To GLORNG_CELL.Rows.Count
              ' Condition '2.
              Let StrAux0 = GLORNG_CELL.Cells(j, 1).Value
              Let StrAux1 = "Celda '" & GLORNG_CELL.Cells(j, 1).Address(False, False) & "': "
              Let BooAux0 = RES.REGEX.isNumberReg(StrAux0)
              If Not BooAux0 Then Call RES.FrmCMsgbox(StrAux1 & "¡Condición '" & StrAux0 & "' inválida!", STRAPPTITLE, vbInformation, True, VBtnCurrent): GoTo EH
              ' %Base|Cur/T '1.
              If BooAux0 Then
                If Not RES.REGEX.isPercentReg(StrAux0) Then Call RES.FrmCMsgbox(StrAux1 & "¡Porcentaje '" & StrAux0 & "' inválido!", STRAPPTITLE, vbInformation, True, VBtnCurrent): GoTo EH
                If VBA.Abs(VBA.CDbl(StrAux0)) = 0 Or VBA.Abs(VBA.CDbl(StrAux0)) > 100 Then Call RES.FrmCMsgbox(StrAux1 & "¡Porcentaje '" & StrAux0 & "' inválido!", STRAPPTITLE, vbInformation, True, VBtnCurrent): GoTo EH
              Else
                If Not RES.REGEX.isNumberReg(.Controls("Txtcnc12").Value) Then Call RES.FrmCMsgbox(StrAux1 & "¡Base '" & StrAux0 & "' inválida!", STRAPPTITLE, vbInformation, True, VBtnCurrent): GoTo EH
              End If
            Next j

          Case 3, 4

            If (.Controls("Cbocnc1").ListIndex = 3 And .Controls("Cbocnc2").ListIndex = 3) Or .Controls("Cbocnc1").ListIndex = 4 Then ' Discounts or bonus
              Let EnuTbl = IIf(.Controls("Cbocnc1").ListIndex = 3, eqlMdlTblPurdiscounts_tmp, eqlMdlTblPurbonus_tmp)

              Let i = MDL.LengthRows(eqlMdlSheetTmp, EnuTbl)
              Let j = GLORNG_CELL.Rows.Count
              If i >= GLOSNG_PUR_LIMIT Then Call RES.FrmCMsgbox("¡Ya se alcanzó al límite de " & GLOSNG_PUR_LIMIT & " registros!", STRAPPTITLE, VBA.vbInformation, True, .Controls("Lstcnc0")): GoTo EH
              If (j + i) > GLOSNG_PUR_LIMIT Then Call RES.FrmCMsgbox("¡Ya se alcanzó al límite de " & GLOSNG_PUR_LIMIT & " registros!", STRAPPTITLE, VBA.vbInformation, True, .Controls("Lstcnc0")): GoTo EH
              If Not GLORNG_CELL.Columns.Count = 2 Then Call RES.FrmCMsgbox("Dos columnas serán suficientes en la selección", STRAPPTITLE, VBA.vbInformation, True, VBtnCurrent): GoTo EH
              If Application.WorksheetFunction.CountA(GLORNG_CELL) = 0 Then Call RES.FrmCMsgbox("¿Y los datos?", STRAPPTITLE, VBA.vbInformation, True, VBtnCurrent): GoTo EH

              ' Search for duplicates
              Let StrAux2 = IIf(.Controls("Cbocnc1").ListIndex = 3, "descuento", "bono")
              Let VecAux0 = RES.ArrayToBaseZero(GLORNG_CELL.Columns("A:A").Value, eqlRes1D)
              If RES.ArrayHasRepeatedElements(VecAux0) Then Call RES.FrmCMsgbox("La selección contiene nombres de " & StrAux2 & "s duplicados. Es necesario que todos los nombres sean únicos.", STRAPPTITLE, VBA.vbInformation, True, VBtnCurrent): GoTo EH
              Let VecAux0 = Empty
              ' Validate data
              If i > 0 Then Let VecAux0 = RES.ArrayToBaseZero(RES.ArrayDelIndex(.Controls("Lstcnc0").List, 1, True, True), eqlRes1D)

              For j = 1 To GLORNG_CELL.Rows.Count
                ' Concept
                Let StrAux0 = GLORNG_CELL.Cells(j, 1).Value
                Let StrAux1 = "Celda '" & GLORNG_CELL.Cells(j, 1).Address(False, False) & "': "
                If RES.REGEX.isEmptyStringReg(StrAux0) Then Call RES.FrmCMsgbox(StrAux1 & "El concepto del " & StrAux2 & " es requerido.", STRAPPTITLE, VBA.vbInformation, True, VBtnCurrent): GoTo EH
                If i > 0 Then
                  Let i = RES.ArrayHasValueN(VecAux0, VBA.LCase(StrAux0), ALCase:=True)
                  If i > 0 Then Call RES.FrmCMsgbox(StrAux1 & "El " & StrAux2 & ": '" & StrAux0 & "' ya existe en los registros. Ingrese otro valor.", STRAPPTITLE, vbInformation, True, VBtnCurrent): GoTo EH
                End If
                ' Value
                Let StrAux0 = GLORNG_CELL.Cells(j, 2).Value
                Let StrAux1 = "Celda '" & GLORNG_CELL.Cells(j, 2).Address(False, False) & "': "
                If Not RES.REGEX.isNumberReg(StrAux0) Then Call RES.FrmCMsgbox(StrAux1 & "El valor del " & StrAux2 & " '" & StrAux0 & "' ya existe en los registros. Ingrese otro dato.", STRAPPTITLE, vbInformation, True, VBtnCurrent): GoTo EH
                Let DblAux1 = VBA.Abs(VBA.CDbl(StrAux0))
                If (DblAux1 = 0) Then Call RES.FrmCMsgbox(StrAux1 & "¡Se esperaba un valor numérico!", STRAPPTITLE, vbInformation, True, VBtnCurrent): GoTo EH
              Next j
            Else
              GoTo EH
            End If

          Case Else: GoTo EH
        End Select
        Let View_Core_validations_concentrateinvoice = True

      Case "Btnmain1"
        ' Basic Data: Date
        If Not RES.REGEX.isDateReg(.Controls("Txtcnc1").Text) Then Call RES.FrmCMsgbox("¡Se esperaba una fecha válida!", STRAPPTITLE, VBA.vbInformation, True, .Controls("Txtcnc1")): GoTo EH
        ' Concentrate Data: Date
        If Not RES.REGEX.isDateReg(.Controls("Txtcnc4").Text) Then Call RES.FrmCMsgbox("¡Se esperaba una fecha válida!", STRAPPTITLE, VBA.vbInformation, True, .Controls("Txtcnc4")): GoTo EH
        ' Concentrate Data: DBLWeight
        If Not RES.REGEX.isNumberReg(.Controls("Txtcnc5").Text) Then Call RES.FrmCMsgbox("¡Peso inválido!", STRAPPTITLE, VBA.vbInformation, True, .Controls("Txtcnc5")): GoTo EH
        Let DblAux0 = VBA.Abs(VBA.CDbl(.Controls("Txtcnc5").Text))
        If DblAux0 = 0 Then Call RES.FrmCMsgbox("¡Peso inválido!", STRAPPTITLE, VBA.vbInformation, True, .Controls("Txtcnc5")): GoTo EH
        ' Concentrate Data: DBLH2O
        If Not RES.REGEX.isPercentReg(.Controls("Txtcnc7").Text) Then Call RES.FrmCMsgbox("¡Porcentaje de humedad inválido!", STRAPPTITLE, VBA.vbInformation, True, .Controls("Txtcnc7")): GoTo EH
        Let DblAux0 = VBA.Abs(VBA.CDbl(.Controls("Txtcnc7").Text))
        If DblAux0 = 0 Or DblAux0 > 100 Then Call RES.FrmCMsgbox("¡Porcentaje de humedad inválido!", STRAPPTITLE, VBA.vbInformation, True, .Controls("Txtcnc7")): GoTo EH
        ' Concentrate Data: DBLDecrease
        If Not RES.REGEX.isPercentReg(.Controls("Txtcnc9").Text) Then Call RES.FrmCMsgbox("¡Porcentaje de merma inválido!", STRAPPTITLE, VBA.vbInformation, True, .Controls("Txtcnc9")): GoTo EH
        Let DblAux0 = VBA.Abs(VBA.CDbl(.Controls("Txtcnc9").Text))
        If DblAux0 = 0 Or DblAux0 > 100 Then Call RES.FrmCMsgbox("¡Porcentaje de merma inválido!", STRAPPTITLE, VBA.vbInformation, True, .Controls("Txtcnc9")): GoTo EH
        ' Treatment cost Data: DBLMaquila
        If Not RES.REGEX.isNumberReg(.Controls("Txtcnc6").Text) Then Call RES.FrmCMsgbox("¡Maquila inválida!", STRAPPTITLE, VBA.vbInformation, True, .Controls("Txtcnc6")): GoTo EH
        Let DblAux0 = VBA.Abs(VBA.CDbl(.Controls("Txtcnc6").Text))
        If DblAux0 = 0 Then Call RES.FrmCMsgbox("¡Maquila inválida!", STRAPPTITLE, VBA.vbInformation, True, .Controls("Txtcnc6")): GoTo EH
        ' Treatment cost Data: DBLBase
        If Not RES.REGEX.isNumberReg(.Controls("Txtcnc8").Text) Then Call RES.FrmCMsgbox("¡Precio base inválido!", STRAPPTITLE, VBA.vbInformation, True, .Controls("Txtcnc8")): GoTo EH
        Let DblAux0 = VBA.Abs(VBA.CDbl(.Controls("Txtcnc8").Text))
        If DblAux0 = 0 Then Call RES.FrmCMsgbox("¡Precio base inválido!", STRAPPTITLE, VBA.vbInformation, True, .Controls("Txtcnc8")): GoTo EH
        ' Treatment cost Data: DBLScaler
        If Not RES.REGEX.isNumberReg(.Controls("Txtcnc10").Text) Then Call RES.FrmCMsgbox("¡Escalador inválido!", STRAPPTITLE, VBA.vbInformation, True, .Controls("Txtcnc10")): GoTo EH
        Let DblAux0 = VBA.Abs(VBA.CDbl(.Controls("Txtcnc10").Text))
        If DblAux0 = 0 Then Call RES.FrmCMsgbox("¡Escalador inválido!", STRAPPTITLE, VBA.vbInformation, True, .Controls("Txtcnc10")): GoTo EH

        ' Main
        If MDL.LengthRows(eqlMdlSheetPur, eqlMdlTblpurchasecnccontents_pur) <= 0 Then Call RES.FrmCMsgbox("¡Faltan datos de " & .Controls("Cbocnc1").List(0, 0) & "!", STRAPPTITLE, VBA.vbInformation, True, .Controls("Lstcnc0")): GoTo EH
        ' Penalties
        Let StrAux0 = VBA.Join(RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetPur, eqlMdlTblpurchasecncpenalties_pur, eqlMdlArray, "SELECTED"), eqlRes1D), ";")
        If VBA.InStr(StrAux0, "1") <= 0 Then
          If VBA.MsgBox("No registró ninguna penalidad, ¿Pasar por alto?", vbQuestion + vbYesNo) = vbNo Then GoTo EH
        End If
        ' Discounts
        Let StrAux0 = VBA.Join(RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetTmp, eqlMdlTblConfigpurretentions_tmp, eqlMdlArray, "SELECTED"), eqlRes1D), ";")
        Let StrAux1 = VBA.Join(RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetTmp, eqlMdlTblConfigpurretentionsother_tmp, eqlMdlArray, "SELECTED"), eqlRes1D), ";")
        If VBA.LCase(.Controls("Cbocnc0").Value) = "particular" Then
          If VBA.InStr(StrAux0, "1") <= 0 Then Call VBA.MsgBox("El Provedor es de tipo Particular, debe realizar las retenciones de ley", vbInformation): GoTo EH
          If VBA.InStr(StrAux1, "1") > 0 Then
            If VBA.MsgBox("El Provedor es de tipo Particular, no debería aplicar retenciones institucionales otras. ¿Pasar por alto?", vbQuestion + vbYesNo) = vbNo Then GoTo EH
          End If
        Else
          If VBA.InStr(StrAux0, "1") <= 0 Then Call VBA.MsgBox("El Provedor es de tipo Cooperativista, debería realizar las retenciones de ley", vbInformation): GoTo EH
          If VBA.InStr(StrAux0, "1") <= 0 Then
            If VBA.MsgBox("El Provedor es de tipo Cooperativista, debería aplicarle las retenciones institucionales otras. ¿Pasar por alto?", vbQuestion + vbYesNo) = vbNo Then GoTo EH
          End If
        End If

        ' Check for spreadsheet edges and data inside the region selected
        If .Controls("Optcnc1").Value Then
          Let j = 150 ' 80: A1 to totalsrow aprox
          Let k = 50
          If GLORNG_CELL Is Nothing Then Call CMB.CMsgBox("¡Seleccione una celda!", STRAPPTITLE, "Error", cmbError, cmbClose, FRM, "Txtcnc15", "Cbocnc1;Cbocnc2;Btnmain1", cmbRequired, cmbRight, cmbBottom, cAspect:=cmbTagLine): GoTo EH
          If RES.ExcelIsProtectedSheet(GLORNG_CELL.Parent) Then Call RES.FrmCMsgbox("La celda que seleccionó para imprimir los resultados está en una hoja protegida. Para realizar un cambio, debe quitar la protección de la hoja. Es posible que se le pida que escriba una contraseña.", STRAPPTITLE, vbInformation, True, .Controls("Txtcnc15")): GoTo EH
          If Not View_assets_check_for_spreadsheet_edges(GLORNG_CELL, j, k, StrAux0) Then Call RES.FrmCMsgbox(StrAux0, STRAPPTITLE, vbInformation, True, .Controls("Txtcnc15")): GoTo EH
        Else
          If Application.ActiveWorkbook.ProtectStructure Then Call RES.FrmCMsgbox("No se puede crear una nueva hoja en el libro actual porque está protegido. Quite la protección o seleccione otro libro para continuar.", STRAPPTITLE, vbInformation, True, .Controls("Optcnc0")): GoTo EH
          Set GLORNG_CELL = Nothing
        End If
        Let StrAux0 = "Se generará la Liquidación de concentrados con los datos que ingresó." & VBA.vbNewLine & "¿Proceder?"
        Let View_Core_validations_concentrateinvoice = (VBA.MsgBox(StrAux0, vbQuestion + vbYesNo, STRAPPTITLE) = vbYes)

      Case Else: GoTo EH
    End Select
  End With

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Let VecAux2 = Empty
  Let VecAux3 = Empty
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function View_core_validations_config( _
  ByVal VCmdBtn As MSForms.CommandButton, _
  ByVal FRM As FRMBasics) As Boolean

  Dim StrAux0 As String
  Dim i As Long, j As Long, k As Long, l As Long
  Dim DblAux0 As Double, DblAux1 As Double

  Const sFN As String = "VIEW::-validations_config"

  On Error GoTo EH
  Let View_core_validations_config = False

  With FRM
    Select Case VCmdBtn.Name

      ' DIVISA
      Case "Btncon0", "Btncon1" ' Add, Upd (Divisa)
        Let j = IIf(VCmdBtn.Name = "Btncon0", 0, 1)
        Let k = MDL.LengthRows(eqlMdlSheetSys, eqlMdlTblConfigcurrencies_sys)
        For i = 0 To 2
          If Not RES.REGEX.isStringNoSemicolon(.Controls("Txtcon" & i).Text) Then Call RES.FrmCMsgbox("¡Se esperaba el " & IIf(i > 1, "símbolo", "nombre") & " de una divisa!", STRAPPTITLE, vbInformation, True, .Controls("Txtcon" & i), vbCritical, True, True): GoTo EH
          If k > 0 Then
            If k >= GLOSNG_PUR_LIMIT Then Call RES.FrmCMsgbox("¡La cantidad divisas ya llegó al límite de " & GLOSNG_PUR_LIMIT & " registros!", STRAPPTITLE, vbInformation, True, VCmdBtn, vbCritical, True, True): GoTo EH
            Let l = RES.ArrayHasValueN2D(.Controls("Lstcon0").List, .Controls("Txtcon" & i).Text, i)
            If l > j Then Call RES.FrmCMsgbox("¡La divisa: '" & .Controls("Txtcon" & i).Text & "' ya existe en los registros, introduzca otro valor!", STRAPPTITLE, vbInformation, True, .Controls("Txtcon" & i), vbCritical, True, True): GoTo EH
          End If
        Next i

      Case "Btncon3" ' Del
        If Not (.Controls("Lstcon0").ListCount > 0 And .Controls("Lstcon0").ListIndex > 1) Then GoTo EH
        If VBA.MsgBox("¿Realmente desea eliminar la divisa seleccionada?", vbQuestion + vbYesNo) = vbNo Then GoTo EH

      Case "Btncon4" ' Local
        If .Controls("Cbocon0").Tag = .Controls("Cbocon0").Value Then GoTo EH
        If VBA.MsgBox("¿Realmente desea cambiar la moneda local de '" & .Controls("Cbocon0").Tag & "' a: '" & .Controls("Cbocon0").Value & "'?", vbQuestion + vbYesNo) = vbNo Then GoTo EH

      Case "Btncon5" ' Foreign
        If .Controls("Cbocon1").Tag = .Controls("Cbocon1").Value Then GoTo EH
        If VBA.MsgBox("¿Realmente desea cambiar la moneda extranjera de '" & .Controls("Cbocon1").Tag & "' a: '" & .Controls("Cbocon1").Value & "'?", vbQuestion + vbYesNo) = vbNo Then GoTo EH

      ' PÁGINA
      Case "Btncon6" ' Save all button
        If RES.REGEX.isStringReg(.Controls("Txtcon3").Value) Then ' Img file
          Let StrAux0 = .Controls("Txtcon3").Value & "[" & IIf(.Controls("Chkcon3").Value = False, 0, IIf(.Controls("Chkcon3").Value, 1, 2)) & "]"
          If Not (VBA.LCase(StrAux0) = VBA.LCase(.Controls("Txtcon3").Tag)) Then
            If Not RES.REGEX.isValidImageFile(.Controls("Txtcon3").Text) Then Call RES.FrmCMsgbox("¡Se requiere un archivo de imagen válido!", STRAPPTITLE, vbInformation, True, .Controls("Txtcon3"), vbCritical, True, True): GoTo EH
            If VBA.Dir(.Controls("Txtcon3").Value) = VBA.vbNullString Then Call RES.FrmCMsgbox("¡Se requiere un archivo de imagen válido!", STRAPPTITLE, vbInformation, True, .Controls("Txtcon3"), vbCritical, True, True): GoTo EH
          End If
        End If
        For i = 0 To 2 ' Company data
          If Not RES.REGEX.isStringNoSemicolon(.Controls("Txtcon" & i + 4).Text) Then
            If VBA.MsgBox("¿Pasar por alto " & IIf(i = 1, "la ", "el ") & IIf(i = 0, "nombre", VBA.StrReverse(VBA.Mid(VBA.StrReverse(.Controls("Lblcon" & i + 7).Caption), 2))) & " de su empresa?", vbQuestion + vbYesNo, STRAPPTITLE) = vbNo Then GoTo EH
          End If
        Next i
        If RES.REGEX.isEmptyStringReg(.Controls("Txtcon7").Value) Then
          If VBA.MsgBox("¿Pasar por alto el correo electrónico de su empresa?", vbQuestion + vbYesNo, STRAPPTITLE) = vbNo Then GoTo EH
        Else
          If Not RES.REGEX.isEmail(.Controls("Txtcon7").Value) Then Call RES.FrmCMsgbox("¡Ingrese un correo electrónico válido!", STRAPPTITLE, vbInformation, True, .Controls("Txtcon7"), vbCritical, True, True): GoTo EH
        End If

      Case "Btncon7", "Btncon8" ' Destiny: set & upd
        Let i = IIf(VCmdBtn.Name = "Btncon7", 0, 1)
        If .Controls("Cbocon3").ListCount <= 0 Then GoTo EH
        If .Controls("Cbocon3").ListIndex < 4 And i = 1 Then GoTo EH
        If .Controls("Cbocon3").Tag = "noevent" Then GoTo EH
        If Not .Controls("Cbocon3").Tag >= 0 Then GoTo EH
        Let StrAux0 = .Controls("Cbocon3").Value
        If RES.REGEX.isEmptyStringReg(StrAux0) Then
          Call RES.FrmCMsgbox("¡Dato inválido!", STRAPPTITLE, vbInformation, True, .Controls("Cbocon3"), vbCritical, True, True): GoTo EH
        Else
          If Not RES.REGEX.isStringNoSemicolon(StrAux0) Then Call RES.FrmCMsgbox("¡Dato inválido!", STRAPPTITLE, vbInformation, True, .Controls("Cbocon3"), vbCritical, True, True): GoTo EH
          Let j = RES.ArrayHasValueN2D(.Controls("Cbocon3").List, StrAux0, 0)
          If j >= i Then GoTo EH
          If VCmdBtn.Name = "Btncon7" Then
            If .Controls("Cbocon3").ListCount > GLOSNG_PUR_LIMIT Then Call RES.FrmCMsgbox("¡Límite de items alcanzado!", STRAPPTITLE, vbInformation, True, .Controls("Cbocon3"), vbCritical, True, True): GoTo EH
          Else
            If .Controls("Cbocon3").Tag <= 3 Then GoTo EH
          End If
        End If

      Case "Btncon9"
        If .Controls("Cbocon3").ListCount <= 0 Then GoTo EH
        If .Controls("Cbocon3").ListIndex < 4 Then GoTo EH
        If Not .Controls("Cbocon3").Tag >= 0 Then GoTo EH
        If .Controls("Cbocon3").Tag = 0 Then GoTo EH
        If VBA.MsgBox("¿Realmente desea eliminar el ítem " & VBA.CStr(.Controls("Cbocon3").List(.Controls("Cbocon3").Tag, 0)) & " de los registros?", vbQuestion + vbYesNo, STRAPPTITLE) = vbNo Then GoTo EH

      ' LIQUIDACIONES
      Case "Btncon10", "Btncon11" ' set, upd
        Let j = IIf(VCmdBtn.Name = "Btncon10", 0, 1)
        Select Case .Controls("Cbocon4").ListIndex

          Case 0 ' RM
            Let StrAux0 = .Controls("Cbocon5").Value
            Let i = RES.ArrayHasValueN2D(.Controls("Lstcon1").List, StrAux0, 0)
            If i >= j Then Call RES.FrmCMsgbox("¡El mineral '" & StrAux0 & "' ya existe en la base de datos, seleccione otro valor!", STRAPPTITLE, vbInformation, True, .Controls("Cbocon5"), RSelTxt:=True): GoTo EH

            If Not RES.REGEX.isNumberReg(.Controls("Txtcon8").Text) Then Call RES.FrmCMsgbox("Valor incorrecto. Se requiere un número.", STRAPPTITLE, vbInformation, True, .Controls("Txtcon8"), RSelTxt:=True): GoTo EH
            If Not RES.REGEX.isNumberReg(.Controls("Txtcon9").Text) Then Call RES.FrmCMsgbox("Valor incorrecto. Se requiere un número.", STRAPPTITLE, vbInformation, True, .Controls("Txtcon9"), RSelTxt:=True): GoTo EH
            If Not RES.REGEX.isNumberReg(.Controls("Txtcon10").Text) Then Call RES.FrmCMsgbox("Valor incorrecto. Se requiere un número.", STRAPPTITLE, vbInformation, True, .Controls("Txtcon10"), RSelTxt:=True): GoTo EH

          Case 1, 2 ' Ret, RetOth
            Let StrAux0 = VBA.CStr(.Controls("Txtcon8").Text)
            If Not RES.REGEX.isStringNoSemicolon(StrAux0) Then Call RES.FrmCMsgbox("¡Dato incorrecto!. Se requiere un concepto de retención.", STRAPPTITLE, vbInformation, True, .Controls("Txtcon8"), RSelTxt:=True): GoTo EH
            Let i = RES.ArrayHasValueN2D(.Controls("Lstcon1").List, StrAux0, 0)
            If i >= j Then Call RES.FrmCMsgbox("¡El concepto '" & StrAux0 & "' ya existe en la base de datos, introduzca otro valor!", STRAPPTITLE, vbInformation, True, .Controls("Txtcon8"), RSelTxt:=True): GoTo EH

            If Not RES.REGEX.isPercentReg(.Controls("Txtcon9").Text) Then Call RES.FrmCMsgbox("Porcentaje incorrecto.", STRAPPTITLE, vbInformation, True, .Controls("Txtcon9"), RSelTxt:=True): GoTo EH
            If VBA.Abs(VBA.CDbl(.Controls("Txtcon9").Text)) > 100 Or VBA.Abs(VBA.CDbl(.Controls("Txtcon9").Text)) = 0 Then Call RES.FrmCMsgbox("Porcentaje incorrecto.", STRAPPTITLE, vbInformation, True, .Controls("Txtcon9"), RSelTxt:=True): GoTo EH

          Case 3 ' Penalties
            Let StrAux0 = .Controls("Cbocon5").Value
            Let i = RES.ArrayHasValueN2D(.Controls("Lstcon1").List, StrAux0, 1)
            If i >= j Then Call RES.FrmCMsgbox("¡El elemento o especie '" & StrAux0 & "' ya está sujeto a penalización, introduzca otro valor!", STRAPPTITLE, vbInformation, True, .Controls("Cbocon5"), RSelTxt:=True): GoTo EH

            If Not RES.REGEX.isPercentReg(.Controls("Txtcon8").Text) Then Call RES.FrmCMsgbox("Valor incorrecto.", STRAPPTITLE, vbInformation, True, .Controls("Txtcon8"), RSelTxt:=True): GoTo EH
            If VBA.Abs(VBA.CDbl(.Controls("Txtcon8").Text)) > 100 Or VBA.Abs(VBA.CDbl(.Controls("Txtcon8").Text)) = 0 Then Call RES.FrmCMsgbox("Valor incorrecto, se requiere un número entre 0 y 100%.", STRAPPTITLE, vbInformation, True, .Controls("Txtcon8"), RSelTxt:=True): GoTo EH
            If Not RES.REGEX.isPercentReg(.Controls("Txtcon9").Text) Then Call RES.FrmCMsgbox("Valor incorrecto.", STRAPPTITLE, vbInformation, True, .Controls("Txtcon9"), RSelTxt:=True): GoTo EH
            If VBA.Abs(VBA.CDbl(.Controls("Txtcon9").Text)) > 100 Or VBA.Abs(VBA.CDbl(.Controls("Txtcon9").Text)) = 0 Then Call RES.FrmCMsgbox("Valor incorrecto, se requiere un número entre 0 y 100%.", STRAPPTITLE, vbInformation, True, .Controls("Txtcon9"), RSelTxt:=True): GoTo EH
            If Not RES.REGEX.isNumberReg(.Controls("Txtcon10").Text) Then Call RES.FrmCMsgbox("Valor incorrecto.", STRAPPTITLE, vbInformation, True, .Controls("Txtcon10"), RSelTxt:=True): GoTo EH
            If VBA.Abs(VBA.CDbl(.Controls("Txtcon10").Text)) > CDBL_MAX Or VBA.Abs(VBA.CDbl(.Controls("Txtcon10").Text)) = 0 Then Call RES.FrmCMsgbox("Valor incorrecto.", STRAPPTITLE, vbInformation, True, .Controls("Txtcon10"), RSelTxt:=True): GoTo EH
            If Not RES.REGEX.isPercentReg(.Controls("Txtcon11").Text) Then Call RES.FrmCMsgbox("Valor incorrecto.", STRAPPTITLE, vbInformation, True, .Controls("Txtcon11"), RSelTxt:=True): GoTo EH
            If VBA.Abs(VBA.CDbl(.Controls("Txtcon11").Text)) > 100 Or VBA.Abs(VBA.CDbl(.Controls("Txtcon11").Text)) = 0 Then Call RES.FrmCMsgbox("Valor incorrecto, re requiere un número entre 0 y 100%.", STRAPPTITLE, vbInformation, True, .Controls("Txtcon11"), RSelTxt:=True): GoTo EH

          Case 4 'Deductions in Exports
            Let StrAux0 = VBA.CStr(.Controls("Txtcon8").Text)
            If Not RES.REGEX.isStringNoSemicolon(StrAux0) Then Call RES.FrmCMsgbox("¡Dato incorrecto!. Se requiere un concepto.", STRAPPTITLE, vbInformation, True, .Controls("Txtcon8"), RSelTxt:=True): GoTo EH
            Let i = RES.ArrayHasValueN2D(.Controls("Lstcon1").List, StrAux0, 0)
            If i >= j Then Call RES.FrmCMsgbox("¡El concepto '" & StrAux0 & "' ya existe en la base de datos, introduzca otro valor!", STRAPPTITLE, vbInformation, True, .Controls("Txtcon8"), RSelTxt:=True): GoTo EH

            Let DblAux0 = VBA.Abs(VBA.CDbl(.Controls("Txtcon9").Value))
            Let DblAux1 = VBA.Abs(VBA.CDbl(.Controls("Txtcon10").Value))
            If Not RES.REGEX.isPercentReg(.Controls("Txtcon9").Text) Then Call RES.FrmCMsgbox("Valor incorrecto.", STRAPPTITLE, vbInformation, True, .Controls("Txtcon9"), RSelTxt:=True): GoTo EH
            If DblAux0 > 100 Then Call RES.FrmCMsgbox("Valor incorrecto, se requiere un número entre 0 y 100%.", STRAPPTITLE, vbInformation, True, .Controls("Txtcon9"), RSelTxt:=True): GoTo EH
            If Not RES.REGEX.isNumberReg(.Controls("Txtcon10").Text) Then Call RES.FrmCMsgbox("Valor incorrecto.", STRAPPTITLE, vbInformation, True, .Controls("Txtcon10"), RSelTxt:=True): GoTo EH
            If DblAux1 > CDBL_MAX Then Call RES.FrmCMsgbox("Valor incorrecto.", STRAPPTITLE, vbInformation, True, .Controls("Txtcon10"), RSelTxt:=True): GoTo EH
            If (DblAux0 = 0 And DblAux1 = 0) Or (DblAux0 > 0 And DblAux1 > 0) Then Call RES.FrmCMsgbox("Sólo un valor es requerido: o el porcentaje de descuento, o el valor de descuento.", STRAPPTITLE, vbInformation, True, .Controls("Txtcon9"), RSelTxt:=True): GoTo EH
            If (DblAux1 = 0 And .Controls("Chkcon7").Value) Then Call RES.FrmCMsgbox("El valor de descuento debe ser mayor a cero para aplicar esta opción.", STRAPPTITLE, vbInformation, True, .Controls("Chkcon7")): GoTo EH

          Case Else: GoTo EH
        End Select

      Case "Btncon13"
        If .Controls("Cbocon4").ListCount <= 0 Then GoTo EH
        If VBA.MsgBox("¿Realmente desea eliminar de la base de datos el registro seleccionado?", vbQuestion + vbYesNo, STRAPPTITLE) = vbNo Then GoTo EH

      '   ELEMENTOS QUÍMICOS
      Case "Btncon14", "Btncon15" ' set, upd
        Let StrAux0 = .Controls("Txtcon8").Text
        If Not RES.REGEX.isStringNoSemicolon(StrAux0) Then Call RES.FrmCMsgbox("¡Dato incorrecto!. El nombre del elemento o especie mineral es requerido.", STRAPPTITLE, vbInformation, True, .Controls("Txtcon13"), RSelTxt:=True): GoTo EH
        Let i = RES.ArrayHasValueN2D(.Controls("Lstcon2").List, StrAux0, 0)
        If i >= j Then Call RES.FrmCMsgbox("¡El elemento o mineral '" & StrAux0 & "' ya existe en la base de datos, introduzca otro valor!", STRAPPTITLE, vbInformation, True, .Controls("Txtcon13"), RSelTxt:=True): GoTo EH
        Let StrAux0 = .Controls("Txtcon9").Text
        If Not RES.REGEX.isStringNoSemicolon(StrAux0) Then Call RES.FrmCMsgbox("¡Dato incorrecto!. El símbolo del elemento o especie mineral es requerido.", STRAPPTITLE, vbInformation, True, .Controls("Txtcon14"), RSelTxt:=True): GoTo EH
        Let i = RES.ArrayHasValueN2D(.Controls("Lstcon2").List, StrAux0, 1)
        If i >= j Then Call RES.FrmCMsgbox("¡El símbolo '" & StrAux0 & "' ya existe en la base de datos, introduzca otro valor!", STRAPPTITLE, vbInformation, True, .Controls("Txtcon14"), RSelTxt:=True): GoTo EH

      Case "Btncon17" ' del
        If .Controls("Lstcon2").ListCount <= 0 Then GoTo EH
        If VBA.MsgBox("¿Realmente desea eliminar de la base de datos el registro seleccionado?", vbQuestion + vbYesNo, STRAPPTITLE) = vbNo Then GoTo EH

      Case Else: GoTo EH
    End Select
    Let View_core_validations_config = True
  End With

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function


' Values Preparation
Private Function View_Core_preparations( _
  ByVal FRM As FRMBasics, _
  ByVal ControlId As String, _
  ByRef VVarRequest As Variant, _
  ByRef VBooNewSheet As Boolean, _
  ByRef VLngVecRngIndex As Long, _
  ByVal VRngBox As Range) As Boolean

  Const sFN As String = "VIEW::-ValuesPreparation"

  On Error GoTo EH
  Select Case ControlId
    Case "regression", "regressionselection": Let View_Core_preparations = View_Core_preparations_regression(FRM, ControlId, VVarRequest, VBooNewSheet, VLngVecRngIndex, VRngBox)
    Case "balmet", "balmetselection", "balmeteco": Let View_Core_preparations = View_Core_preparations_balmet(FRM, ControlId, VVarRequest, VBooNewSheet, VLngVecRngIndex, VRngBox)
    Case "projection": Let View_Core_preparations = View_Core_preparations_projection(FRM, ControlId, VVarRequest, VBooNewSheet, VLngVecRngIndex)
    Case "oreinvoice": Let View_Core_preparations = View_Core_preparations_oreinvoice(FRM, ControlId, VVarRequest, VBooNewSheet, VLngVecRngIndex)
    Case "concentrateinvoice": Let View_Core_preparations = View_Core_preparations_concentrateinvoice(FRM, ControlId, VVarRequest, VBooNewSheet, VLngVecRngIndex)

    Case "hlpregression", "hlprlselection", "hlpbalmet", "hlpbalmetselection", "hlpbalmeteco", "hlpprojection", "hlporeinvoice", "hlpconcentrateinvoice", "hlpudf"
      Let View_Core_preparations = View_core_preparations_help(VVarRequest, VBooNewSheet, VLngVecRngIndex)

    Case Else: GoTo EH
  End Select

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function View_Core_preparations_regression( _
  ByVal FRM As FRMBasics, _
  ByVal ControlId As String, _
  ByRef VVarRequest As Variant, _
  ByRef VBooNewSheet As Boolean, _
  ByRef VLngVecRngIndex As Long, _
  ByVal VRngBox As Range) As Boolean

  Dim RngBox As Range
  Dim StrY As String, StrX As String, StrExt As String
  Dim BooJustReg As Boolean, BooShowAll As Boolean, BooGraph As Boolean, BooSheetGraph As Boolean
  Dim i As Long

  Const sFN As String = "VIEW::-ValuesPreparation_regression"

  On Error GoTo EH
  Let View_Core_preparations_regression = False

  Select Case ControlId

    Case "regression"
      ' Request content: (0.StrX, 1.StrY, 2.StrE, 3.BooShowAll, 4.BooJustRegression, 5.BooX2, 6.BooY2, 7.BooXY, 8.Booa, 9.Boob, 10.Boor, 11.Boor2, 12.BooSe, 13.BooGraph, 14.BooSheetGraph, 15.BooNewSheet, 16.RngBox)
      With FRM
        ' Get main data
        Let StrX = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(.Controls("Txtreg0").Value, i, RValueType:=eqlResNumbers, RBooIncludeZeroes:=True, RDblTop:=CDBL_MAX) '0-X
        Let StrY = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(.Controls("Txtreg1").Value, i, RValueType:=eqlResNumbers, RBooIncludeZeroes:=True, RDblTop:=CDBL_MAX) '0-X'1-Y
        Let StrExt = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(.Controls("Txtreg2").Value, i, RValueType:=eqlResNumbers, RBooIncludeZeroes:=True, RDblTop:=CDBL_MAX) '0-X'2-E
        Let BooShowAll = .Controls("Chkreg0").Value ' ShowAll
        Let BooJustReg = .Controls("Chkreg1").Value ' JustRegression
        Let BooGraph = .Controls("Chkreg10").Value ' 13.Graph
        Let BooSheetGraph = .Controls("Optreg0").Value ' 14.SheetGraph
        Let VBooNewSheet = .Controls("Optreg2").Value ' 15.NewSheet
        ' Set request vector
        Let VVarRequest = Array(StrX, StrY, StrExt, BooShowAll, BooJustReg, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, BooGraph, BooSheetGraph, VBooNewSheet, Nothing)
        For i = 5 To 12 ' Boolean options
          Let VVarRequest(i) = IIf(BooJustReg, False, (.Controls("Chkreg" & (i - 3)).Value Or BooShowAll))
        Next i
        Set VVarRequest(16) = GLORNG_CELL ' 13.RngBox
        Set GLORNG_CELL = Nothing
        Let VLngVecRngIndex = 16
      End With
      Let View_Core_preparations_regression = True

    Case "regressionselection"
      ' Request content: (0.StrX, 1.StrY, 2.StrE, 3.BooShowAll, 4.BooJustRegression, 5.BooX2, 6.BooY2, 7.BooXY, 8.Booa, 9.Boob, 10.Boor, 11.Boor2, 12.BooSe, 13.BooGraph, 14.BooSheetGraph, 15.BooNewSheet, 16.RngBox)
      ' Get main data
      Set RngBox = Application.Selection
      If RngBox Is Nothing Then GoTo EH
      With RngBox
        If .Areas.Count = 1 Then
          Let StrY = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(.Columns("A:A"), i, RValueType:=eqlResNumbers, RBooIncludeZeroes:=True, RDblTop:=CDBL_MAX) ' 0.X
          If .Columns.Count > 1 Then Let StrX = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(RngBox.Columns("B:B"), i, RValueType:=eqlResNumbers, RBooIncludeZeroes:=True, RDblTop:=CDBL_MAX) ' 1.Y
          If .Columns.Count > 2 Then Let StrExt = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(RngBox.Columns("C:Z"), i, RValueType:=eqlResNumbers, RBooIncludeZeroes:=True, RDblTop:=CDBL_MAX) ' 2.E
        ElseIf .Areas.Count > 1 Then
          Let StrY = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(.Areas(1), i, RValueType:=eqlResNumbers, RBooIncludeZeroes:=True, RDblTop:=CDBL_MAX) ' 1.Y
          If .Areas.Count > 1 Then Let StrX = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(.Areas(2), i, RValueType:=eqlResNumbers, RBooIncludeZeroes:=True, RDblTop:=CDBL_MAX) ' 0.X
          If .Areas.Count > 2 Then Let StrExt = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(.Areas(3), i, RValueType:=eqlResNumbers, RBooIncludeZeroes:=True, RDblTop:=CDBL_MAX) ' 2.E
        End If
      End With
      Set RngBox = Nothing
      Let BooShowAll = Not GLOBOO_COMPACT ' ShowAll
      Let BooJustReg = GLOBOO_COMPACT ' JustRegression
      Let BooGraph = False ' 13.Graph
      Let BooSheetGraph = False ' 14.SheetGraph
      Let VBooNewSheet = Not GLOBOO_SELECTIONCELL ' 15.NewSheet
      ' Set request vector
      Let VVarRequest = Array(StrX, StrY, StrExt, BooShowAll, BooJustReg, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, BooGraph, BooSheetGraph, VBooNewSheet, Nothing)
      For i = 5 To 12 ' Boolean options
        Let VVarRequest(i) = IIf(BooJustReg, False, BooShowAll)
      Next i
      Set VVarRequest(16) = VRngBox 'GLORNG_CELL ' 13.RngBox
      Let VLngVecRngIndex = 16
      Let View_Core_preparations_regression = True

    Case Else: GoTo EH
  End Select

EH:
  Set RngBox = Nothing
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function View_Core_preparations_balmet( _
  ByVal FRM As FRMBasics, _
  ByVal ControlId As String, _
  ByRef VVarRequest As Variant, _
  ByRef VBooNewSheet As Boolean, _
  ByRef VLngVecRngIndex As Long, _
  ByVal VRngBox As Range) As Boolean

  Dim RngBox As Range
  Dim VecGrade As Variant, VecWeig As Variant, VecRecs As Variant, VecGradeCx As Variant, VecRecCx As Variant
  Dim VecName As Variant, VecUnit As Variant, VecNameCx As Variant, VecUnitCx As Variant
  Dim VecAux0 As Variant, VecAux1 As Variant, VecAux2 As Variant, VecAux3 As Variant, VecEcos As Variant
  Dim StrGrade As String, StrWeig As String, StrRecs As String, StrName As String, StrUnit As String, StrUnitCx As String, StrGradeCx As String, StrRecCx As String, _
      StrNameA As String, StrUnitA As String, StrFeedName As String, StrAux0 As String, StrAux1 As String
  Dim BooAux0 As Boolean, BooShowAll As Boolean
  Dim DblFeed As Double, DblFeedVol As Double
  Dim i As Long, j As Long, k As Long, l As Long, m As Long, n As Long
  Dim EnuBalType As EQLBAL_ENU_TYPE, EnuMethod As EQLBAL_ENU_METHOD, EnuTypeFeed As EQLBAL_ENU_TYPECONTENT, EnuEcoType As EQLBAL_ENU_RESULTECO

  Const sFN As String = "VIEW::-ValuesPreparation_balmet"

  On Error GoTo EH
  Let View_Core_preparations_balmet = False

  Select Case ControlId

    Case "balmet" ' -------------------------------------------------------------------- Balmet
      With FRM ' Get data from controls

        ' FeedName
        Let StrFeedName = .Controls("Txtbal1").Value
        With .Controls("Cbobal0") ' Method
          Let StrAux0 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(FRM.Controls("Txtbal2").Value, i, RValueType:=eqlResNumbers, RBooAbs:=True) ' Grades
          Let i = RES.RESCORE.BalGetProductsBounds(i, BWhat:=eqlBalProducts)
          Let EnuMethod = IIf(.ListIndex = 0, eqlBalConventional, IIf(.ListIndex = 1, eqlBalCramer, eqlBalInverseMatrix))
          Let EnuMethod = IIf(i > 1 And EnuMethod = eqlBalConventional, eqlBalCramer, EnuMethod)
        End With
        With .Controls("Txtbal0") ' Feed
          Let StrAux0 = VBA.LCase(FRM.Controls("Txtbal4").Text & IIf(FRM.Controls("Chkbal3").Value, IIf(FRM.Controls("Txtbal7").Text = VBA.vbNullString, VBA.vbNullString, ";" & FRM.Controls("Txtbal7").Text), VBA.vbNullString))
          Let EnuTypeFeed = RES.RESCORE.BalTypeFeed(RES, StrAux0)
          Let VecAux0 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(.Value, k, RBooAbs:=True, RValueAs:=eqlResNumbers)
          If Not RES.RESCORE.BalGetFeed(EnuTypeFeed, VecAux0, DblFeed, DblFeedVol) Then Call VBA.MsgBox("Falló la preparación del valor de alimentación", vbInformation, STRAPPTITLE): GoTo EH
          Let VecAux0 = Empty
        End With
        ' Grades
        Let StrGrade = .Controls("Txtbal2").Value
        ' Names
        Let StrName = .Controls("Txtbal3").Value: Let StrNameA = StrName ' (A: when formulas is active)
        ' Units
        Let StrUnit = .Controls("Txtbal4").Value: Let StrUnitA = StrUnit ' (A: when formulas is active)

        ' GradesCx
        Let StrGradeCx = IIf(.Controls("Chkbal3").Value, .Controls("Txtbal5").Value, VBA.vbNullString)
        ' &NamesCx
        With .Controls("Txtbal6")
          Let StrName = StrName & IIf(FRM.Controls("Chkbal3").Value, IIf(.Value = VBA.vbNullString, VBA.vbNullString, ";" & .Value), VBA.vbNullString)
        End With
        ' &UnitCx
        With .Controls("Txtbal7")
          Let StrUnit = StrUnit & IIf(FRM.Controls("Chkbal3").Value, IIf(.Value = VBA.vbNullString, VBA.vbNullString, ";" & .Value), VBA.vbNullString)
        End With

        ' BooNewSheet
        Let VBooNewSheet = .Controls("Optbal0").Value

        ' Booleans values: 0.BooAll, 1.BooPercents, 2.BooUnities, 3.BooFines, 4.BooGrams, 5.BooOT, 6.BooRatio, 7.BooHeads (, 8.Formulas, 9.Tables)
        Let BooShowAll = .Controls("Chkbal6").Value
        ReDim VecAux0(9)
        For i = 0 To 7
          With .Controls("Chkbal" & i + 6)
            Let VecAux0(i) = IIf(BooShowAll, BooShowAll, .Value)
          End With
        Next i
        Let VecAux0(8) = GLOBOO_FORMULAS
        Let VecAux0(9) = GLOBYT_TABLES

        ' Data required in BVarRequest: (0.DblFeed, 1.DblFeedVol, 2.VecGrad, 3.VecGradCx, 4.VecName, 5.VecUnit, 6.EnuMethod, 7.Booleans, 8.EnuTypeFeed, 9.RngBox (, 10.StrFeedName, 11.StrNameA, 12.StrUnitA, 13. BooNewSheet))
        Let VVarRequest = Array(DblFeed, DblFeedVol, StrGrade, StrGradeCx, StrName, StrUnit, EnuMethod, VecAux0, EnuTypeFeed, Nothing, StrFeedName, StrNameA, StrUnitA, VBooNewSheet)
        Set VVarRequest(9) = GLORNG_CELL
        Let VLngVecRngIndex = 9

      End With
      Let View_Core_preparations_balmet = True

    Case "balmetselection" ' ------------------------------------------------------------ Balmetselection

      Set RngBox = Application.Selection
      With RngBox

        ' Method
        Let EnuMethod = eqlBalCramer

        ' Extract all balmet and cx grades
        Let StrGrade = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(.Areas(2), j, RValueType:=eqlResNumbers, RBooAbs:=True)
        If .Areas.Count > 2 Then Let StrGradeCx = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(.Areas(3), k, RValueType:=eqlResNumbers, RBooAbs:=True)
        Let EnuBalType = RES.RESCORE.BalGetType(j, k) ' Balmet type
        ' Get cols of products
        Let l = RES.RESCORE.BalGetProductsBounds(j, BWhat:=eqlBalColumns) - 1 ' Main Balmet cols
        Let m = RES.RESCORE.BalGetProductsBounds(j, k, BWhat:=eqlBalColumnsCx) - 1 ' Cx Balmet cols

        ' Filter names & units
        Let StrAux1 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(RngBox, n, RValueType:=eqlResNotNumbers, RStrCustomRegex:="%")

        ' Get info from db
        Let VecAux0 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalelements_sys, eqlMdlArray, "Símbolo"), eqlRes1D)
        Let VecAux1 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalelementscx_sys, eqlMdlArray, "Símbolo"), eqlRes1D)
        Let VecAux2 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalunits_sys, eqlMdlArray, "Unidad"), eqlRes1D)
        Let VecAux3 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalunitscx_sys, eqlMdlArray, "Unidad"), eqlRes1D)

        ' Names
        Call RES.RESCORE.BalFilterNames(RES, VecAux0, VBA.Split(StrAux1, ";"), VecName, True)
        Let StrAux0 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(VBA.CStr(VecName), n, RValueType:=eqlResNotNumbers, RBooUniques:=True)
        If Not RES.RESCORE.BalRefillNames2(RES, VecAux0, StrAux0, BLngGradCol:=l, BLngGradCxCol:=m, BVarName:=VecName, BVarNameCx:=VecNameCx, BStrErrMsg:=StrAux0, BBooCx:=(EnuBalType = eqlBalWithComplex), BVarDBNamesCx:=VecAux1, BBooAsString:=True) Then Call VBA.MsgBox("Balmet Selección: Los nombres de elementos son incorrectos.", vbInformation, STRAPPTITLE): GoTo EH
        Let StrNameA = VBA.CStr(VecName)
        Let StrName = VBA.CStr(VecName) & IIf(EnuBalType = eqlBalWithComplex, ";" & VBA.CStr(VecNameCx), VBA.vbNullString)

        ' Units
        Call RES.RESCORE.BalFilterUnits(RES, VecAux2, VBA.Split(StrAux1, ";"), VecUnit, True)
        Let StrAux0 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(VBA.CStr(VecUnit), n, RValueType:=eqlResNotNumbers, RStrCustomRegex:="%")
        If Not RES.RESCORE.BalRefillUnits2(RES, VecAux2, StrAux0, BLngGradCol:=l, BLngGradCxCol:=m, BVarUnit:=VecUnit, BVarUnitCx:=VecUnitCx, BStrErrMsg:=StrAux0, BBooCx:=(EnuBalType = eqlBalWithComplex), BVarDBUnitsCx:=VecAux3, BBooAsString:=True) Then Call VBA.MsgBox("Balmet Selección: Las unidades de elementos son incorrectas.", vbInformation, STRAPPTITLE): GoTo EH
        Let StrUnitA = VBA.CStr(VecUnit)
        Let StrUnit = VBA.CStr(VecUnit) & IIf(EnuBalType = eqlBalWithComplex, ";" & VBA.CStr(VecUnitCx), VBA.vbNullString)
        Let VecAux0 = Empty

        ' Feeds
        Let EnuTypeFeed = RES.RESCORE.BalTypeFeed(RES, StrUnit)
        Let VecAux0 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(.Areas(1), i, RBooAbs:=True, RValueAs:=eqlResNumbers)
        If Not RES.RESCORE.BalGetFeed(EnuTypeFeed, VecAux0, DblFeed, DblFeedVol) Then Call VBA.MsgBox("Falló la preparación del valor de alimentación", vbInformation, STRAPPTITLE): GoTo EH
        Let VecAux0 = Empty

        ' BooNewSheet
        Let VBooNewSheet = Not GLOBOO_SELECTIONCELL

        ' Booleans values: 0.BooAll, 1.BooPercents, 2.BooUnities, 3.BooFines, 4.BooGrams, 5.BooOT, 6.BooRatio, 7.BooHeads (, 8.Formulas, 9.Tables)
        ReDim VecAux0(9)
        For i = 0 To 7
          Let VecAux0(i) = GLOBOO_COMPACT
        Next i
        Let VecAux0(7) = GLOBOO_HEADS
        Let VecAux0(8) = GLOBOO_FORMULAS
        Let VecAux0(9) = GLOBYT_TABLES

        ' Data required in BVarRequest: (0.DblFeed, 1.DblFeedVol, 2.VecGrad, 3.VecGradCx, 4.VecName, 5.VecUnit, 6.EnuMethod, 7.Booleans, 8.EnuTypeFeed, 9.RngBox (, 10.StrFeedName, 11.StrNameA, 12.StrUnitA, 13. BooNewSheet))
        Let VVarRequest = Array(DblFeed, DblFeedVol, StrGrade, StrGradeCx, StrName, StrUnit, EnuMethod, VecAux0, EnuTypeFeed, Nothing, StrFeedName, StrNameA, StrUnitA, VBooNewSheet)
        Set VVarRequest(9) = VRngBox
        Let VLngVecRngIndex = 9

      End With

      Let View_Core_preparations_balmet = True

    Case "balmeteco" ' ------------------------------------------------------------------ Balmet
      ' Data required in BVarRequest: (0.DblFeed, 1.DblFeedVol, 2.VecGrad-, 3.VecGradCx-, 4.VecName, 5.VecUnit, 6.EnuMethod, 7.Booleans, 8.EnuTypeFeed, 9.RngBox (, 10.StrFeedName, 11.StrNameA, 12.StrUnitA, 13. BooNewSheet, 14.BEnuTypeBalEco, 15.VecEcos))
      ' Booleans values: 0.BooAll, 1.BooPercents, 2.BooUnities, 3.BooFines, 4.BooGrams, 5.BooOT, 6.BooRatio, 7.BooHeads (, 8.Formulas, 9.Tables)
      ' VecEcos: (0.Array(DblFeedSolid, DblFeedVolume), 1.VecWeights, 2.VecGrades, 3.VecRecs, 4.VecGradesCx, 5.VecRecsCx, 6.VecUnits, 7.VecUnitsCx, 8.BEnuTypeBalEco)

      With FRM

        ' FeedName
        Let StrFeedName = .Controls("Txtbal4").Value
        ' Method
        With .Controls("Cbobal1")
          Let EnuMethod = IIf(.ListIndex = 0, eqlBalConventional, IIf(.ListIndex = 1, eqlBalCramer, eqlBalInverseMatrix))
        End With
        ' Feed
        With .Controls("Txtbal0")
          Let StrAux0 = VBA.LCase(FRM.Controls("Txtbal6").Text & IIf(FRM.Controls("Chkbal4").Value, IIf(FRM.Controls("Txtbal9").Text = VBA.vbNullString, VBA.vbNullString, ";" & FRM.Controls("Txtbal9").Text), VBA.vbNullString))
          Let EnuTypeFeed = RES.RESCORE.BalTypeFeed(RES, StrAux0)
          Let VecAux0 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(.Value, k, RValueType:=eqlResNumbers, RBooAbs:=True, RValueAs:=eqlResNumbers)
          If Not RES.RESCORE.BalGetFeed(EnuTypeFeed, VecAux0, DblFeed, DblFeedVol) Then Call VBA.MsgBox("Falló la preparación del valor de alimentación", vbInformation, STRAPPTITLE): GoTo EH
          Let VecAux0 = Empty
        End With
        ' Prods
        Let VecWeig = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(.Controls("Txtbal1").Value, i, RValueType:=eqlResNumbers, RBooAbs:=True, RValueAs:=eqlResNumbers)
        ' Grades
        Let VecGrade = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(.Controls("Txtbal2").Value, i, RValueType:=eqlResNumbers, RBooAbs:=True, RValueAs:=eqlResNumbers)
        ' Recs
        Let VecRecs = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(.Controls("Txtbal3").Value, i, RValueType:=eqlResNumbers, RBooAbs:=True, RValueAs:=eqlResNumbers)
        ' Names
        Let StrName = .Controls("Txtbal5").Value: Let StrNameA = StrName ' (A: when formulas is active)
        ' Units
        Let StrUnit = .Controls("Txtbal6").Value: Let StrUnitA = StrUnit ' (A: when formulas is active)

        ' Cx
        Let BooAux0 = .Controls("Chkbal4").Value
        If BooAux0 Then
          ' GradesCx
          Let VecGradeCx = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(.Controls("Txtbal7").Value, i, RValueType:=eqlResNumbers, RBooAbs:=True, RValueAs:=eqlResNumbers)
          ' &NamesCx
          Let StrName = StrName & IIf(.Controls("Txtbal8").Value = VBA.vbNullString, VBA.vbNullString, ";" & .Controls("Txtbal8").Value)
          ' &UnitCx
          Let StrUnit = StrUnit & IIf(.Controls("Txtbal9").Value = VBA.vbNullString, VBA.vbNullString, ";" & .Controls("Txtbal9").Value)
          Let StrUnitCx = .Controls("Txtbal9").Value
          ' RecsCx
          Let VecRecCx = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(.Controls("Txtbal10").Value, i, RValueType:=eqlResNumbers, RBooAbs:=True, RValueAs:=eqlResNumbers)
        End If

        ' Adjust Method
        Let StrAux0 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(VBA.Join(VecGrade, ";"), i, RValueType:=eqlResNumbers, RBooAbs:=True)
        If EnuMethod = eqlBalConventional And i > 1 Then Let EnuMethod = eqlBalCramer
        ' Type Bal Eco
        Let EnuEcoType = IIf(.Controls("Cbobal0").ListIndex = 0, eqlBalHeadsGrades, eqlBalProdsGrades)
        ' BooNewSheet
        Let VBooNewSheet = .Controls("Optbal0").Value

        ' VecEcos: (0.Array(DblFeedSolid, DblFeedVolume), 1.VecWeights, 2.VecGrades, 3.VecRecs, 4.VecGradesCx, 5.VecRecsCx, 6.VecUnits, 7.VecUnitsCx, 8.BEnuTypeBalEco)
        Let VecEcos = Array(Array(DblFeed, DblFeedVol), VecWeig, VecGrade, VecRecs, VecGradeCx, VecRecCx, VBA.Split(StrUnitA, ";"), IIf(BooAux0, VBA.Split(StrUnitCx, ";"), Empty), EnuEcoType)

        ' Booleans values: 0.BooAll, 1.BooPercents, 2.BooUnities, 3.BooFines, 4.BooGrams, 5.BooOT, 6.BooRatio, 7.BooHeads (, 8.Formulas, 9.Tables)
        Let BooShowAll = .Controls("Chkbal8").Value
        ReDim VecAux0(9)
        For i = 0 To 7
          With .Controls("Chkbal" & i + 8)
            Let VecAux0(i) = IIf(BooShowAll, BooShowAll, .Value)
          End With
        Next i
        Let VecAux0(8) = GLOBOO_FORMULAS
        Let VecAux0(9) = GLOBYT_TABLES

        ' Data required in BVarRequest: (0.DblFeed, 1.DblFeedVol, 2.VecGrad-, 3.VecGradCx-, 4.VecName, 5.VecUnit, 6.EnuMethod, 7.Booleans, 8.EnuTypeFeed, 9.RngBox (, 10.StrFeedName, 11.StrNameA, 12.StrUnitA, 13. BooNewSheet, 14.BEnuTypeBalEco, 15.VecEcos))
        Let VVarRequest = Array(DblFeed, DblFeedVol, Empty, Empty, StrName, StrUnit, EnuMethod, VecAux0, EnuTypeFeed, Nothing, StrFeedName, StrNameA, StrUnitA, VBooNewSheet, EnuEcoType, VecEcos)
        Set VVarRequest(9) = GLORNG_CELL
        Let VLngVecRngIndex = 9
        Let VecEcos = Empty
        Let VecAux0 = Empty

      End With
      Let View_Core_preparations_balmet = True

    Case Else: GoTo EH
  End Select

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Let VecAux2 = Empty
  Let VecAux3 = Empty
  Let VecName = Empty
  Let VecNameCx = Empty
  Let VecUnit = Empty
  Let VecUnitCx = Empty
  Let VecGrade = Empty
  Let VecWeig = Empty
  Let VecRecs = Empty
  Let VecGradeCx = Empty
  Let VecRecCx = Empty
  Let VecEcos = Empty
  Set RngBox = Nothing
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function View_Core_preparations_projection( _
  ByVal FRM As FRMBasics, _
  ByVal ControlId As String, _
  ByRef VVarRequest As Variant, _
  ByRef VBooNewSheet As Boolean, _
  ByRef VLngVecRngIndex As Long) As Boolean

  Dim VecAux0 As Variant
  Dim StrWeigh As String, StrWName As String, StrWUnit As String, StrCtz As String, StrDivise As String, StrTimeUnit As String
  Dim DblIo As Double, DblVR As Double, DblEg As Double, DblTx As Double
  Dim LngTm As Long, i As Long

  Const sFN As String = "VIEW::-preparations_projection"

  On Error GoTo EH
  Let View_Core_preparations_projection = False

  With FRM

    ' Prepare source data
    Let DblIo = VBA.CDbl(VBA.Replace(.Controls("Txtpro0").Value, ".", ",")) ' 0.DblIo
    Let DblVR = VBA.CDbl(VBA.Replace(.Controls("Txtpro3").Value, ".", ",")) ' 1.DblVR
    Let DblEg = VBA.CDbl(VBA.Replace(VBA.Replace(.Controls("Txtpro4").Value, ".", ","), "%", VBA.vbNullString)) ' 2.DblEgr
    Let LngTm = VBA.CLng(.Controls("Txtpro7").Value) ' 3.LngTm
    Let DblTx = VBA.CDbl(VBA.Replace(VBA.Replace(.Controls("Txtpro8").Value, ".", ","), "%", VBA.vbNullString)) ' 4.Dbl%
    Let StrWeigh = .Controls("Txtpro1").Value ' 5.VecW[]
    Let StrWName = .Controls("Txtpro5").Value ' 6.VecWName[]
    Let StrWUnit = .Controls("Txtpro6").Value ' 7.VecWUnitCtz[]
    Let StrCtz = .Controls("Txtpro2").Value ' 8.VecCtz[]
    Let StrDivise = .Controls("Cbopro0").Value ' 9.StrDivise
    Let StrTimeUnit = .Controls("Cbopro3").Value ' 10.StrTimeUnit

    ' Fill Request
    ' Data required in BVarRequest: (0.DblIo, 1.DblVR, 2.DblEgr, 3.DblTm, 4.Dbl%, 5.VecW, 6.VecWName, 7.VecWUnitCtz, 8.VecCtz, 9.StrDivise, 10.StrTimeUnit, 11.BooVAN, 12.BooTIR, 13.BooGraph, 14.BooGraphSheet, 15.BooNewSheet, 16.RngBox)
    Let VVarRequest = Array(DblIo, DblVR, DblEg, LngTm, DblTx, StrWeigh, StrWName, StrWUnit, StrCtz, StrDivise, StrTimeUnit, Empty, Empty, Empty, Empty, Empty, Nothing)
    Let VecAux0 = Array("Chkpro0", "Chkpro1", "Chkpro2", "Optpro1", "Optpro2")
    For i = 11 To 15
      Let VVarRequest(i) = .Controls(VBA.CStr(VecAux0(i - 11))).Value
    Next i
    Set VVarRequest(16) = GLORNG_CELL ' 16.RngBox
    Let VLngVecRngIndex = 16 ' New sheet variables to get_cell
    Let VBooNewSheet = VBA.CBool(VVarRequest(15))

  End With
  Let View_Core_preparations_projection = True

EH:
  Let VecAux0 = Empty
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function View_Core_preparations_oreinvoice( _
  ByVal FRM As FRMBasics, _
  ByVal ControlId As String, _
  ByRef VVarRequest As Variant, _
  ByRef VBooNewSheet As Boolean, _
  ByRef VLngVecRngIndex As Long) As Boolean

  Dim VecAux0 As Variant
  Dim StrAux0 As String
  Dim DblAux0 As Double
  Dim i As Long

  Const sFN As String = "VIEW::-ValuesPreparation_oreinvoice"

  On Error GoTo EH
  Let View_Core_preparations_oreinvoice = False

  ' Request: 0.StrPlace, 1.BooME, 2.BooVarious, 3.BooAverage, 4.BooPage, 5.BooPrintTwo, 6.BooNewSheet, 7.RngBox
  With FRM

    ' Fill "FECHA LIQUIDACIÓN" in db
    Let StrAux0 = IIf(RES.REGEX.isEmptyStringReg(.Controls("Txtpur1").Text), VBA.Format(VBA.Date, "dd/mm/yy"), IIf(RES.REGEX.isDateReg(.Controls("Txtpur1").Text), .Controls("Txtpur1").Text, VBA.Format(VBA.Date, "dd/mm/yy")))
    Call MDL.MUpd(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur, Array(VBA.CDate(StrAux0)), "FECHA LIQUIDACIÓN")

    ' Fill "TC" in db
    Let DblAux0 = IIf(RES.REGEX.isEmptyStringReg(.Controls("Txtpur2").Text), 1, IIf(RES.REGEX.isNumberReg(.Controls("Txtpur2").Text), VBA.CDbl(.Controls("Txtpur2").Text), 1))
    Call MDL.MUpd(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur, Array(DblAux0), "TC")

    ' Get Provider(s) name(s) and Fill Provider column in db
    Let VecAux0 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(.Controls("Txtpur3").Text, i, RValueType:=eqlResNotNumbers, RValueAs:=eqlResNotNumbers)
    If i = 1 Then Let StrAux0 = VBA.CStr(VecAux0(0))
    Let StrAux0 = IIf(i <= 0, "Desconocido", IIf(i = 1, StrAux0, VBA.vbNullString))
    If i > 1 Then
      Call MDL.MUpd(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur, Array(VecAux0), "PROVEDOR")
    Else
      Call MDL.MUpd(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur, Array(StrAux0), "PROVEDOR")
    End If
    Let VecAux0 = Empty

    ' Get Provider(s) type(s) and Fill Provider type column in db
    If .Controls("Chkpur1").Value Then
      Let VecAux0 = VBA.Split(.Controls("Txtpur4").Text, ";")
      Call MDL.MUpd(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur, Array(VecAux0), "TIPO PROVEDOR")
      Let VecAux0 = Empty
    Else
      Call MDL.MUpd(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur, Array(.Controls("Cbopur0").Value), "TIPO PROVEDOR")
    End If

    ' Get Place
    Let StrAux0 = IIf(RES.REGEX.isEmptyStringReg(.Controls("Txtpur0").Text), "N/D", .Controls("Txtpur0").Text)

    ' Prepare data
    ' 0.StrPlace, 1.BooME, 2.BooVarious, 3.BooAverage, 4.BooPage, 5.BooPrintTwo, 6.BooNewSheet, 7.RngBox
    Let VVarRequest = Array(StrAux0, .Controls("Chkpur0").Value, .Controls("Chkpur1").Value, .Controls("Chkpur3").Value, .Controls("Chkpur4").Value, .Controls("Chkpur5").Value, .Controls("Optpur0").Value, Nothing)
    Set VVarRequest(7) = GLORNG_CELL ' 7.RngBox
    Let VLngVecRngIndex = 7 ' New sheet variables to get_cell
    Let VBooNewSheet = VBA.CBool(VVarRequest(6))

  End With

  Let View_Core_preparations_oreinvoice = True

EH:
  Let VecAux0 = Empty
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function View_Core_preparations_concentrateinvoice( _
  ByVal FRM As FRMBasics, _
  ByVal ControlId As String, _
  ByRef VVarRequest As Variant, _
  ByRef VBooNewSheet As Boolean, _
  ByRef VLngVecRngIndex As Long) As Boolean

  Const sFN As String = "VIEW::-ValuesPreparation_concentrateinvoice"

  On Error GoTo EH
  Let View_Core_preparations_concentrateinvoice = False

  ' Request: 0.BooPage, 1.BooNewSheet, 2.RngBox, 3.SIMULATION
  With FRM
    Let VVarRequest = Array(.Controls("Chkcnc2").Value, .Controls("Optcnc0").Value, GLORNG_CELL, True)
    Let VLngVecRngIndex = 2 ' New sheet variables to get_cell
    Let VBooNewSheet = VBA.CBool(VVarRequest(1))
  End With

  Let View_Core_preparations_concentrateinvoice = True

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function View_core_preparations_help( _
  ByRef VVarRequest As Variant, _
  ByRef VBooNewSheet As Boolean, _
  ByRef VLngVecRngIndex As Long) As Boolean

  Const sFN As String = "VIEW::-core_preparations_help"

  On Error GoTo EH
  Let View_core_preparations_help = False

  ' Request: 0.RngBox, 1.BooNewSheet
  Let VVarRequest = Array(Nothing, True)
  Let VLngVecRngIndex = 0 ' New sheet variables to get_cell
  Let VBooNewSheet = True

  Let View_core_preparations_help = True

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

' Executor method
Private Function View_Core_execution( _
  ByVal ControlId As String, _
  ByVal FRM As FRMBasics, _
  Optional VRngBox As Range = Nothing, _
  Optional VStatusMsg As String = VBA.vbNullString, _
  Optional VBooHelp As Boolean = False) As Boolean

  Dim Wbk As Workbook
  Dim VecRes As Variant, VecReq As Variant, VarOldStatusBar As Variant
  Dim StrSheetName As String
  Dim BooCancel As Boolean, BooNewSheet As Boolean, BooDeleteNewSheet As Boolean
  Dim LngVecRngIndex As Long

  Const sFN As String = "VIEW::-EXE"

  On Error GoTo EH
  ' ------------------------------------------------ Control application objects
  Let Application.ScreenUpdating = False
  If Not VStatusMsg = VBA.vbNullString Then
    Let VarOldStatusBar = Application.DisplayStatusBar
    Let Application.DisplayStatusBar = True
    Let Application.StatusBar = "Trabajando en...: " & VStatusMsg
    #If Debugging Then
      Call VBA.MsgBox("Trabajando en...: " & VStatusMsg, vbInformation, STRAPPTITLE)
    #End If
  End If
  ' ------------------------------------------------ Control application objects

  #If Debugging Then
    Call VBA.MsgBox("Ready to exe!", vbInformation, STRAPPTITLE) ': GoTo EH
  #End If
  If Not View_Core_preparations(FRM, ControlId, VecReq, BooNewSheet, LngVecRngIndex, VRngBox) Then GoTo EH

'  GoTo EH ' <<<====

  ' Event implementation -----------------------------------------------------------------
  ' Get range from selected cell or by creating a new sheet
  If Not View_assets_get_selected_or_new_range(ControlId, VecReq) Then GoTo EH

  ' Raise the custom event
  If VBooHelp Then
    Let View_Core_execution = True
  Else
    RaiseEvent RequestEvent(ControlId, VecReq, VecRes, View_Core_execution) 'BooCancel)
  End If

'  GoTo EH ' <<<====

  ' Process the received data
  If View_Core_execution Then

    ' Print results on spreadsheet
    #If Debugging Then
      Call VBA.MsgBox("Success!!!", vbInformation, STRAPPTITLE)
    #End If
    Dim VPR As ClassEqlViewPrint
    Set VPR = New ClassEqlViewPrint
    Let View_Core_execution = VPR.ViewPrint(RES, MDL, ControlId, VecReq, VecRes)
    Set VPR = Nothing
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
    Let View_Core_execution = View_core_error_handler(ControlId, FRM)

  End If
  ' Event implementation -----------------------------------------------------------------

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

  ' ------------------------------------------------ Control application objects
  If Not VStatusMsg = VBA.vbNullString Then
    Let Application.StatusBar = False
    Let Application.DisplayStatusBar = VarOldStatusBar
  End If
  Let Application.ScreenUpdating = True
  ' ------------------------------------------------ Control application objects

  Let VarOldStatusBar = Empty
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function


' CORE VIEWS FORMS METHODS LIST ---------------------------------------------------- '
' Regression <
Private Sub View_form_regression(ByVal ControlId As String, ByVal FRM As FRMBasics)

  ' Main objects
  Dim VWP As ClassEqlViewParent
  ' ListObjects
  ' LO data names and other data
  ' Controls
  Dim aLabels() As Variant
  Dim aMaxLengths() As Variant
  Dim aControls() As Variant
  Dim aControlis() As Variant
  ' Standard vars
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

  ' 1. [x] Get data for Form Controls

  ' 2. [x] Desing the Form
  Let TitleForm = "Regresión Lineal"
  Let StrTitle = TitleForm & " - " & STRAPPTITLE

  Set VWP = New ClassEqlViewParent
  Call VWP.ViewFormCleanner(RES, FRM)
  With FRM
    ' Set the form dimensions
    Let .Height = 415
    Let .Width = 450
    Let .Tag = ControlId
    Let .Caption = StrTitle
    Let xLeft = .InsideWidth * 0.05

    ' Build and design the form
    Call VWP.ViewFormAddControl(RES, FRM, "CommandButton", "Btnmain", 2) 'All with events (Cancel, Exe and Help)
    Call VWP.ViewFormAddControl(RES, FRM, "Frame", "Framereg", 3)

    Call VWP.ViewFormAddControl(RES, .Controls("Framereg0"), "Label", "Lblreg", 2)
    Call VWP.ViewFormAddControl(RES, .Controls("Framereg0"), "TextBox", "Txtreg", 2) ' EVENT all
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

    Call VWP.ViewFormAddControl(RES, .Controls("Framereg1"), "CheckBox", "Chkreg", 9)
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

    Call VWP.ViewFormAddControl(RES, .Controls("Framereg2"), "CheckBox", "Chkreg", 10, CtlFromX:=10)
    Call VWP.ViewFormAddControl(RES, .Controls("Framereg2"), "OptionButton", "Optreg", 1)
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

    Call VWP.ViewFormAddControl(RES, .Controls("Framereg3"), "OptionButton", "Optreg", 3, CtlFromX:=2)
    Call VWP.ViewFormAddControl(RES, .Controls("Framereg3"), "TextBox", "Txtreg", 3, CtlFromX:=3)
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

    ' Push data to controls

    ' General purpose CommandButtons
    ' 3. Fill Form Controls with data and Set general purpose CommandButtons
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
      Let .Tag = "validate"
      Let .Default = True
    End With

    ' 4. Assing handler for each button
    Set BtnCancel = .Controls("Btnmain0") 'Main
    Set BtnAccept = .Controls("Btnmain1") 'Main

    Set Btn0 = .Controls("Btnmain2") 'Main
    Set Txt0 = .Controls("Txtreg0") 'Frame0
    Set Txt1 = .Controls("Txtreg1") 'Frame0
    Set Txt2 = .Controls("Txtreg2") 'Frame0
    Set Txt3 = .Controls("Txtreg3") 'Frame3

    ' 5. Give styles to controls
    Call VWP.ViewFormStyles(RES, FRM)
    Set VWP = Nothing

    ' 6. Show the finished Form
    Call Load(FRM)
    Call .Show

    ' 7. Close form
  End With

EH:
  Erase aLabels
  Erase aMaxLengths
  Erase aControlis
  Erase aControls
  Set VWP = Nothing
  Debug.Print VBA.Err.Description
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Sub

' Selection Regression UI <
Private Sub View_form_regressionselection(ByVal ControlId As String)

  Const sFN As String = "VIEW::-regressionselection"

  On Error GoTo EH
  Call View_controls_buttons_click(Nothing, Nothing, True, ControlId)

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Sub


' Balmet <
Private Sub View_form_balmet(ByVal ControlId As String, ByVal FRM As FRMBasics)

  ' Main objects
  Dim VWP As ClassEqlViewParent
  ' Controls
  Dim FmeCurrent As MSForms.Frame
  Dim VecLabels As Variant
  Dim VecMaxLgt As Variant
  Dim VecCntrls0 As Variant
  Dim VecCntrls1 As Variant
  ' Standard vars
  Dim StrTitle As String, StrTitleForm As String, StrAux0 As String
  Dim i As Long, j As Long, k As Long, l As Long, xLeft As Long

  Const sFN As String = "VIEW::-form_balmet"

  #If Debugging Then
    Call Debug_assert(Not ControlId = VBA.vbNullString, sFN)
    Call Debug_assert(Not FRM Is Nothing, sFN)
  #End If

  On Error GoTo EH
  ' 1. [x] Get data for Form Controls

  ' 2. [x] Desing the Form
  Let StrTitleForm = "Balance Metalúrgico"
  Let StrTitle = StrTitleForm & " - " & STRAPPTITLE

  Set VWP = New ClassEqlViewParent
  Call VWP.ViewFormCleanner(RES, FRM)
  With FRM

    ' Set the form dimensions
    Let .Height = 415
    Let .Width = 450
    Let .Tag = ControlId
    Let .Caption = StrTitle
    Let xLeft = .InsideWidth * 0.05 ' Left for all main controls

    ' Build and design the form
    Call VWP.ViewFormAddControl(RES, FRM, "CommandButton", "Btnmain", 2) 'All with events (Cancel, Exe and Help)
    Call VWP.ViewFormAddControl(RES, FRM, "Frame", "Framebal", 3)

    Set FmeCurrent = .Controls("Framebal0")
    With FmeCurrent

      ' Set Frame properties
      Let .Caption = "Ingresar Datos"
      Let .Width = .Parent.InsideWidth * 0.9
      Let .Height = .Parent.InsideHeight * 0.325
      Let .Left = xLeft
      Let .Top = .Parent.InsideHeight * 0.04

      ' Create Frame controls
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "Label", "Lblbal", 2)
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "ComboBox", "Cbobal", 2) ' EVENT 1,2
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "TextBox", "Txtbal", 4) ' EVENT 0,2
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "CheckBox", "Chkbal", 2) ' EVENT all
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "CommandButton", "Btnbal", 3) 'EVENT all

      ' Set Frame controls layout
      Let VecLabels = Array("Método de resolución:", "Peso Alimentación [T]:", "Leyes de minerales y productos:")
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
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "Label", "Lblbal", 3, CtlFromX:=3)
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "CheckBox", "Chkbal", 5, CtlFromX:=3)
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "TextBox", "Txtbal", 7, CtlFromX:=5)
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "CommandButton", "Btnbal", 7, CtlFromX:=4) 'EVENT all
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "ComboBox", "Cbobal", 4, CtlFromX:=3)

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
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "CheckBox", "Chkbal", 13, CtlFromX:=6)

      ' Set Frame controls properties (including toolTipTexts and values)
      Let VecLabels = Array("Todo", "Peso [%]", "Unidades", "Finos", "Finos [T]", "Peso Complejo [OT]", "Radio de concentración", "Leyes de Cabeza Ens. y Calc.")
      Let VecCntrls1 = Array("Imprimir todos los valores", "Porcentaje en Peso", "Unidades", "Finos en Peso (Contenido Metálico)", "Finos en Peso, predeterminado: [Toneladas]. Si desactiva, el cálculo se obtendrá en [gramos]", "Peso del mineral complejo [Onzas Troy]", "Radio de Concentración", "Mostrar las Leyes de Cabeza Ensayada y Calculada")
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
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "OptionButton", "Optbal", 1)
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "TextBox", "Txtbal", 8, CtlFromX:=8) ' EVENT

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
    Let .Controls("Txtbal1").MaxLength = 12
    Let .Controls("Txtbal2").Tag = 0
    Let .Controls("Txtbal5").Tag = 0
    Let .Controls("Cbobal0").List = Array("Convencional (sólo para balances de dos productos)", "Por determinantes (Cramer)", "Por Matriz inversa")
    Let .Controls("Cbobal0").ListStyle = fmListStyleOption
    Let .Controls("Txtbal1").Text = "A"
    Let VecLabels = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalelements_sys, eqlMdlArray, "Símbolo"))
    Let .Controls("Cbobal1").List = VecLabels
    Call .Controls("Cbobal1").AddItem("---", 0)
    Let VecLabels = Empty
    Let VecLabels = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalunits_sys, eqlMdlArray))
    Let .Controls("Cbobal2").List = VecLabels 'Array("%", "DM", "g/T", "g/L")
    Call .Controls("Cbobal2").AddItem("---", 0)
    Let VecLabels = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalelementscx_sys, eqlMdlArray, "Símbolo"))
    Let .Controls("Cbobal3").List = VecLabels
    Call .Controls("Cbobal3").AddItem("---", 0)
    Let VecLabels = Empty
    Let VecLabels = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalunitscx_sys, eqlMdlArray))
    Let .Controls("Cbobal4").List = VecLabels 'Array("DM", "g/T", "g/L")
    Call .Controls("Cbobal4").AddItem("---", 0)
    For i = 0 To 4
      Let .Controls("Cbobal" & i).ListIndex = 0
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
    For i = 7 To 13
      With .Controls("Chkbal" & i)
        Let .Value = True
        Let .Enabled = False
      End With
    Next i

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

    ' 3. [x] Fill Form Controls with data and Set general purpose CommandButtons
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
      Let .Tag = "validate" 'ControlId & "-bal"
      Let .Default = True
    End With

    ' 4. Events: Assing handler for each control
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
    Set Cbo3 = .Controls("Cbobal4") ' Frame1

    ' 5. Give styles to controls
    Call VWP.ViewFormStyles(RES, FRM)
    Set VWP = Nothing

    ' 6. [x] Show the finished Form
    Call Load(FRM)
    Call .Show

    ' 7. [x] Close form
  End With

EH:
  Let VecLabels = Empty
  Let VecMaxLgt = Empty
  Let VecCntrls1 = Empty
  Let VecCntrls0 = Empty
  Set FmeCurrent = Nothing
  Set VWP = Nothing
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Sub

Private Sub View_form_balmeteco(ByVal ControlId As String, ByVal FRM As FRMBasics)

  ' Main objects
  Dim VWP As ClassEqlViewParent
  ' Controls
  Dim FmeCurrent As MSForms.Frame
  Dim VecLabels As Variant
  Dim VecCntrls0 As Variant
  ' Standard vars
  Dim StrTitle As String, StrTitleForm As String
  Dim i As Long, j As Long, k As Long, l As Long, xLeft As Long

  Const sFN As String = "VIEW::-form_balmeteco"

  #If Debugging Then
    Call Debug_assert(Not ControlId = VBA.vbNullString, sFN)
    Call Debug_assert(Not FRM Is Nothing, sFN)
  #End If

  On Error GoTo EH
  ' 1. [x] Get data for Form Controls

  ' 2. [x] Desing the Form
  Let StrTitleForm = "Balance Metalúrgico (Análisis económico)"
  Let StrTitle = StrTitleForm & " - " & STRAPPTITLE

  Set VWP = New ClassEqlViewParent
  Call VWP.ViewFormCleanner(RES, FRM)
  With FRM

    ' Set the form dimensions
    Let .Height = 415
    Let .Width = 450
    Let .Tag = ControlId
    Let .Caption = StrTitle
    Let xLeft = .InsideWidth * 0.05 ' Left for all main controls

    ' Build and design the form
    Call VWP.ViewFormAddControl(RES, FRM, "CommandButton", "Btnmain", 2) 'All with events (Cancel, Exe and Help)
    Call VWP.ViewFormAddControl(RES, FRM, "Frame", "Framebal", 3)

    Set FmeCurrent = .Controls("Framebal0") ' ---------------------------- Framebal0'
    With FmeCurrent

      ' Set Frame properties
      Let .Caption = "Ingresar Datos"
      Let .Width = .Parent.InsideWidth * 0.9
      Let .Height = .Parent.InsideHeight * 0.35
      Let .Left = xLeft
      Let .Top = .Parent.InsideHeight * 0.04

      ' Create Frame controls
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "Label", "Lblbal", 4)
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "ComboBox", "Cbobal", 3) ' EVENT 2,3
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "TextBox", "Txtbal", 6) ' EVENT 0-3
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "CheckBox", "Chkbal", 3) ' EVENT all
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "CommandButton", "Btnbal", 3) 'EVENT all

      ' Set Frame controls layout (Vertical)
      Let VecLabels = Array("Objetivo:", "Método de resolución:", "Peso Alimentación [T]:", "Peso producto/s (Conc) [T]:", "Ley producto/s (Conc):", "Recuperación [%]:")
      Let VecCntrls0 = Array("Lblbal0", "Lblbal1", "Lblbal2", "Lblbal3", "Lblbal4", "Chkbal0")
      ' Labels
      Let j = xLeft * 0.45
      For i = 0 To UBound(VecLabels)
        With .Controls(VBA.CStr(VecCntrls0(i)))
          Let .Caption = VBA.CStr(VecLabels(i))
          Let .Left = xLeft
          Let .AutoSize = True
          Let .WordWrap = False
          Let .Top = j + (.Parent.Controls("Cbobal0").Height * i)
        End With
      Next i
      Let VecCntrls0 = Empty
      Let VecLabels = Empty
      ' Combos & texts (First controls of labels)
      Let VecLabels = Array("Cbobal0", "Cbobal1", "Txtbal0", "Txtbal1", "Txtbal2", "Txtbal3")
      Let VecCntrls0 = Array("Lblbal0", "Lblbal1", "Lblbal2", "Lblbal3", "Lblbal4", "Chkbal0")
      Let j = .Controls("Lblbal3").Left + .Controls("Lblbal3").Width + 3
      For i = 0 To UBound(VecLabels)
        With .Controls(VBA.CStr(VecLabels(i)))
          Let .Left = j
          Let .Top = .Parent.Controls(VBA.CStr(VecCntrls0(i))).Top
        End With
      Next i
      Let VecCntrls0 = Empty
      Let VecLabels = Empty
      ' Check controls
      Let j = .Controls("Txtbal0").Left + .Controls("Txtbal0").Width + 5
      For i = 0 To 2
        With .Controls("Chkbal" & i + 1)
          Let .AutoSize = True
          Let .WordWrap = False
          Let .Left = j
          Let .Top = .Parent.Controls("Txtbal" & i).Top
        End With
      Next i
      Let VecLabels = Empty
      ' Other controls
      Let VecLabels = Array("Txtbal4", "Txtbal5", "Txtbal6")
      Let j = .Controls("Chkbal1").Left + .Controls("Chkbal1").Width
      For i = 0 To UBound(VecLabels)
        With .Controls(VBA.CStr(VecLabels(i)))
          Let .Left = j
          Let .Top = .Parent.Controls("Txtbal" & i).Top
        End With
      Next i
      Let VecLabels = Empty
      Let j = .Controls("Txtbal5").Left + .Controls("Txtbal5").Width
      For i = 0 To 1
        With .Controls("Cbobal" & i + 2)
          Let .Left = j
          Let .Top = .Parent.Controls("Txtbal" & i + 5).Top
        End With
      Next i
      For i = 0 To 1
        With .Controls("Btnbal" & i)
          Let .Width = .Parent.Controls("Cbobal" & i + 2).Height
          Let .Height = .Width
          Let .Top = .Parent.Controls("Cbobal" & i + 2).Top
        End With
      Next i
      For i = 0 To 1
        With .Controls("Btnbal" & i + 2)
          Let .Width = .Parent.Controls("Btnbal" & i).Height
          Let .Height = .Width
          Let .Top = .Parent.Controls("Btnbal" & i).Top
        End With
      Next i

      ' Set Frame controls layout (Horizontal)
      Let .Controls("Cbobal0").Width = .Width - (.Controls("Cbobal0").Left + xLeft)
      Let .Controls("Cbobal1").Width = .Controls("Cbobal0").Width
      Let .Controls("Txtbal4").Width = .Controls("Cbobal0").Width - (.Controls("Txtbal4").Left - .Controls("Cbobal0").Left)
      Let .Controls("Cbobal2").Width = .Controls("Txtbal4").Width - (.Controls("Txtbal5").Width + (.Controls("Btnbal0").Width * 2))
      Let .Controls("Cbobal3").Width = .Controls("Cbobal2").Width
      Let .Controls("Btnbal0").Left = .Controls("Cbobal2").Left + .Controls("Cbobal2").Width
      Let .Controls("Btnbal1").Left = .Controls("Btnbal0").Left
      Let .Controls("Btnbal2").Left = .Controls("Btnbal0").Left + .Controls("Btnbal0").Width
      Let .Controls("Btnbal3").Left = .Controls("Btnbal2").Left

      ' Set Frame controls ToolTipTexts
      Let VecLabels = Array( _
          "Seleccione el objetivo que desea encontrar tras resolver el Balance Metalúrgico", _
          "Seleccione un método para resolver el Balance Metalúrgico", _
          "Peso[;Volumen]: Cantidad(es) separada(s) por punto y coma (Ej. 1;2;3) o seleccione una celda o rango de celdas con datos numéricos", _
          "Peso productos concentrados: Cantidad(es) separada(s) por punto y coma (Ej. 1;2;3) o seleccione una celda o rango de celdas con datos numéricos", _
          "Leyes productos concentrados: Cantidad(es) separada(s) por punto y coma (Ej. 1;2;3) o seleccione una celda o rango de celdas con datos numéricos", _
          "Recuperación (valores entre 0 y 100%): Cantidad(es) separada(s) por punto y coma (Ej. 1;2;3) o seleccione una celda o rango de celdas con datos numéricos", _
          "Active esta casilla e introduzca los valores de Recuperación deseados", _
          "Active esta casilla y cambie el nombre de la Alimentación (A) a una personalizada", _
          "Escriba un nombre personalizado para la Alimentación (A)", _
          "Active esta casilla si desea seleccionar nombres de minerales específicos", _
          "Seleccione un mineral o elemento", _
          "Elimina del cuadro de texto el último mineral o elemento seleccionado", _
          "Elimina del cuadro de texto todos los minerals o elementos seleccionados", _
          "Active esta casilla si desea seleccionar unidades de ley específicas", _
          "Seleccione una unidad de ley de mineral", _
          "Elimina del cuadro de texto la última unidad seleccionada", _
          "Elimina del cuadro de texto todas las unidades seleccionadas")
      Let VecCntrls0 = Array("Cbobal0", "Cbobal1", "Txtbal0", "Txtbal1", "Txtbal2", "Txtbal3", "Chkbal0", "Chkbal1", "Txtbal4", "Chkbal2", "Cbobal2", "Btnbal0", "Btnbal2", "Chkbal3", "Cbobal3", "Btnbal1", "Btnbal3")

      For i = 0 To UBound(VecLabels)
        Let .Controls(VBA.CStr(VecCntrls0(i))).ControlTipText = VBA.CStr(VecLabels(i))
      Next i
      Let VecCntrls0 = Empty
      Let VecLabels = Empty

    End With
    Set FmeCurrent = Nothing

    Set FmeCurrent = .Controls("Framebal1") ' ---------------------------- Framebal1'
    With FmeCurrent

      ' Set Frame properties
      Let .Caption = "Complejos de metales preciosos"
      Let .Width = .Parent.InsideWidth * 0.9
      Let .Height = .Parent.InsideHeight * 0.25
      Let .Left = xLeft
      Let .Top = .Parent.Controls("Framebal0").Top + .Parent.Controls("Framebal0").Height + 5

      ' Create Frame controls
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "Label", "Lblbal", 5, CtlFromX:=5)
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "CheckBox", "Chkbal", 7, CtlFromX:=4) ' EVENT all - 4
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "TextBox", "Txtbal", 10, CtlFromX:=7) ' EVENT 9,10
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "CommandButton", "Btnbal", 7, CtlFromX:=4) ' EVENT all
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "ComboBox", "Cbobal", 5, CtlFromX:=4) ' EVENT all

      ' Set Frame controls layout (Vertical)
      Let j = xLeft * 0.45
      Let VecLabels = Array("Chkbal4", "Chkbal5", "Cbobal4", "Txtbal8") ' First col
      For i = 0 To UBound(VecLabels)
        With .Controls(VBA.CStr(VecLabels(i)))
          Let .Left = xLeft
          Let .Top = j + (.Parent.Controls("Cbobal4").Height * i)
        End With
      Next i
      Let VecLabels = Empty

      ' Set Frame controls layout (Horizontal)
      ' Define width for all txts
      Let k = (.InsideWidth - (xLeft * 2)) / 4
      For i = 7 To 10
        Let .Controls("Txtbal" & i).Width = k
      Next i
      ' First col
      Let .Controls("Lblbal5").Left = xLeft
      Let .Controls("Lblbal5").Top = .Parent.Controls("Chkbal5").Top '* 1.13
      Let .Controls("Txtbal7").Left = xLeft
      Let .Controls("Txtbal7").Top = .Parent.Controls("Cbobal4").Top
      ' Second col
      Let j = .Controls("Txtbal7").Left + .Controls("Txtbal7").Width
      Let .Controls("Chkbal5").Left = j
      Let .Controls("Cbobal4").Left = j
      Let .Controls("Txtbal8").Left = j
      ' Third col
      Let j = .Controls("Txtbal8").Left + .Controls("Txtbal8").Width
      Let .Controls("Chkbal6").Left = j
      Let .Controls("Chkbal6").Top = .Parent.Controls("Chkbal5").Top
      Let .Controls("Cbobal5").Left = j
      Let .Controls("Cbobal5").Top = .Parent.Controls("Cbobal4").Top
      Let .Controls("Txtbal9").Left = j
      Let .Controls("Txtbal9").Top = .Parent.Controls("Txtbal8").Top
      ' Fourth col
      Let j = .Controls("Txtbal9").Left + .Controls("Txtbal9").Width
      Let .Controls("Chkbal7").Left = j
      Let .Controls("Chkbal7").Top = .Parent.Controls("Chkbal6").Top
      Let .Controls("Txtbal10").Left = j
      Let .Controls("Txtbal10").Top = .Parent.Controls("Cbobal5").Top

      ' Buttons and combos layout
      Let .Controls("Cbobal4").Width = k - (.Controls("Cbobal4").Height * 2)
      Let .Controls("Cbobal5").Width = .Controls("Cbobal4").Width
      For i = 4 To 7
        With .Controls("Btnbal" & i)
          Let .Width = .Parent.Controls("Cbobal4").Height
          Let .Height = .Width
          Let .Top = .Parent.Controls("Cbobal4").Top
        End With
      Next i
      Let .Controls("Btnbal4").Left = .Controls("Cbobal4").Left + .Controls("Cbobal4").Width
      Let .Controls("Btnbal5").Left = .Controls("Btnbal4").Left + .Controls("Btnbal4").Width
      Let .Controls("Btnbal6").Left = .Controls("Cbobal5").Left + .Controls("Cbobal5").Width
      Let .Controls("Btnbal7").Left = .Controls("Btnbal6").Left + .Controls("Btnbal6").Width

      ' Set Frame controls ToolTipTexts
      Let VecLabels = Array( _
        "Active esta casilla si los contenidos del mineral principal son de tipo complejo", _
        "Active esta casilla si desea seleccionar nombres de minerales específicos", _
        "Seleccione un mineral o elemento", _
        "Elimina del cuadro de texto el último mineral o elemento seleccionado", _
        "Elimina del cuadro de texto todos los minerals o elementos seleccionados", _
        "Active esta casilla si desea seleccionar unidades de ley específicas", _
        "Seleccione una unidad de ley de mineral", _
        "Elimina del cuadro de texto la última unidad seleccionada", _
        "Elimina del cuadro de texto todas las unidades seleccionadas", _
        "Leyes mienrales complejos: Cantidad(s) separada(s) por punto y coma (Ej. 1;2;3) o seleccione una celda o rango de celdas con datos numéricos", _
        "Active esta casilla para introducir recuperaciones personalizadas", _
        "Recuperación (valores entre 0 y 100%): Cantidad(s) separada(s) por punto y coma (Ej. 1;2;3) o seleccione una celda o rango de celdas con datos numéricos")
      Let VecCntrls0 = Array("Chkbal4", "Chkbal5", "Cbobal4", "Btnbal4", "Btnbal5", "Chkbal6", "Cbobal5", "Btnbal6", "Btnbal7", "Txtbal7", "Chkbal7", "Txtbal10")
      For i = 0 To UBound(VecLabels)
        Let .Controls(VBA.CStr(VecCntrls0(i))).ControlTipText = VBA.CStr(VecLabels(i))
      Next i
      Let VecCntrls0 = Empty
      Let VecLabels = Empty

      ' Set values
      Let .Controls("Chkbal4").Caption = "Es complejo de Plata, u otro metal precioso"
      Let .Controls("Chkbal5").Caption = "Seleccionar minerales"
      Let .Controls("Chkbal6").Caption = "Seleccionar unidades"
      Let .Controls("Lblbal5").Caption = "Leyes:"
      Let .Controls("Chkbal7").Caption = "Recuperación [%]:"

      Let .Controls("Chkbal4").AutoSize = True
      Let .Controls("Chkbal4").WordWrap = False

    End With
    Set FmeCurrent = Nothing

    Set FmeCurrent = .Controls("Framebal2") ' ---------------------------- Framebal2'
    With FmeCurrent

      ' Set Frame properties
      Let .Caption = "Incluir en resultado"
      Let .Width = ((.Parent.InsideWidth * 0.9) / 2) - 2.5
      Let .Height = .Parent.InsideHeight * 0.2
      Let .Left = xLeft
      Let .Top = .Parent.Controls("Framebal1").Top + .Parent.Controls("Framebal1").Height + 5

      ' Create Frame controls
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "CheckBox", "Chkbal", 15, CtlFromX:=8) ' EVENT 8

      ' Set Frame controls properties (including toolTipTexts and values)
      Let VecLabels = Array("Todo", "Peso [%]", "Unidades", "Finos", "Finos [T]", "Peso Complejo [OT]", "Radio de concentración", "Leyes de Cabeza Ens. y Calc.")
      Let VecCntrls0 = Array("Imprimir todos los valores", "Porcentaje en Peso", "Unidades", "Finos en Peso (Contenido Metálico)", "Finos en Peso, predeterminado: [Toneladas]. Si desactiva, el cálculo se obtendrá en [gramos]", "Peso del mineral complejo [Onzas Troy]", "Radio de Concentración", "Mostrar las Leyes de Cabeza Ensayada y Calculada")

      For i = 0 To 7
        With .Controls("Chkbal" & i + 8)
          Let .Caption = VecLabels(i)
          Let .ControlTipText = VecCntrls0(i)
          Let .AutoSize = True
          Let .WordWrap = False
          Let .Left = (xLeft * 0.5) + IIf(i > 3, .Parent.Controls("Chkbal10").Left + .Parent.Controls("Chkbal10").Width, 0)
          Let j = IIf(i > 3, i - 4, i)
          Let .Top = (xLeft * 0.35) + ((.Parent.Controls("Chkbal" & j + 8).Height * 0.9) * j)
        End With
      Next i

    End With
    Set FmeCurrent = Nothing

    Set FmeCurrent = .Controls("Framebal3") ' ---------------------------- Framebal3'
    With FmeCurrent

      ' Set Frame properties
      Let .Caption = "Dónde mostrar resultados"
      Let .Width = ((.Parent.InsideWidth * 0.9) / 2) - 2.5
      Let .Height = .Parent.InsideHeight * 0.2
      Let .Left = .Parent.Controls("Framebal2").Left + .Parent.Controls("Framebal2").Width + 5
      Let .Top = .Parent.Controls("Framebal1").Top + .Parent.Controls("Framebal1").Height + 5

      ' Create Frame controls
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "OptionButton", "Optbal", 1)
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "TextBox", "Txtbal", 11, CtlFromX:=11) ' EVENT

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
      Let .Controls("Txtbal11").Left = .Controls("Optbal1").Left + .Controls("Optbal1").Width
      Let .Controls("Txtbal11").Top = .Controls("Optbal1").Top
      Let .Controls("Txtbal11").Width = j

      ' Set Frame controls ToolTipTexts
      Let .Controls("Txtbal11").ControlTipText = "Seleccionar una celda para mostrar el resultado"

    End With
    Set FmeCurrent = Nothing
    Let VecLabels = Empty
    Let VecCntrls0 = Empty

    ' Set Form textbox controls dropbuttons
    Let VecLabels = Array(0, 1, 2, 3, 7, 10, 11)
    For i = 0 To UBound(VecLabels)
      With .Controls("Txtbal" & VecLabels(i))
        Let .SpecialEffect = fmSpecialEffectRaised
        Let .DropButtonStyle = fmDropButtonStyleReduce
        Let .ShowDropButtonWhen = fmShowDropButtonWhenAlways
      End With
    Next i
    Let VecLabels = Empty

'    ' Set Form controls values
    Let .Controls("Txtbal4").Text = "A"
    Let .Controls("Txtbal4").MaxLength = 12
    Let .Controls("Cbobal0").List = Array("Obtener Ley de Alimentación (Cabeza)", "Obtener Ley de productos")
    Let .Controls("Cbobal0").ListStyle = fmListStyleOption
    Let .Controls("Cbobal1").List = Array("Convencional (sólo para balances de dos productos)", "Por determinantes (Cramer)", "Por Matriz inversa")
    Let .Controls("Cbobal1").ListStyle = fmListStyleOption
    Let VecLabels = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalelements_sys, eqlMdlArray, "Símbolo"))
    Let .Controls("Cbobal2").List = VecLabels
    Call .Controls("Cbobal2").AddItem("---", 0)
    Let VecLabels = Empty
    Let VecLabels = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalunits_sys, eqlMdlArray))
    Let .Controls("Cbobal3").List = VecLabels 'Array("%", "DM", "g/T", "g/L")
    Call .Controls("Cbobal3").AddItem("---", 0)
    Let VecLabels = Empty
    Let VecLabels = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalelementscx_sys, eqlMdlArray, "Símbolo"))
    Let .Controls("Cbobal4").List = VecLabels
    Call .Controls("Cbobal4").AddItem("---", 0)
    Let VecLabels = Empty
    Let VecLabels = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalunitscx_sys, eqlMdlArray))
    Let .Controls("Cbobal5").List = VecLabels 'Array("DM", "g/T", "g/L")
    Call .Controls("Cbobal5").AddItem("---", 0)
    Let VecLabels = Empty
    For i = 0 To 7
      Let .Controls("Btnbal" & i).Caption = IIf(i Mod 2 = 0, VBA.ChrW(&H2A31), VBA.ChrW(&H2716))
    Next i
    For i = 0 To 5
      Let .Controls("Cbobal" & i).ListIndex = 0
    Next i

    Let .Controls("Btnbal1").Caption = VBA.ChrW(&H2A31) ' Frame 0
    Let .Controls("Btnbal2").Caption = VBA.ChrW(&H2716)
    Let VecCntrls0 = Array("Txtbal4", "Cbobal2", "Cbobal3", "Btnbal0", "Btnbal1", "Btnbal2", "Btnbal3")
    For i = 0 To UBound(VecCntrls0)
      Let .Controls(VBA.CStr(VecCntrls0(i))).Enabled = False
    Next i
    Let VecCntrls0 = Empty
    Let VecCntrls0 = Array("Txtbal3", "Txtbal5", "Txtbal6", "Txtbal8", "Txtbal9", "Txtbal10")
    For i = 0 To UBound(VecCntrls0)
      Let .Controls(VBA.CStr(VecCntrls0(i))).Locked = True
      Let .Controls(VBA.CStr(VecCntrls0(i))).ForeColor = VBA.vbGrayText
    Next i
    Let VecCntrls0 = Empty
    Let VecCntrls0 = Array("Cbobal4", "Btnbal4", "Btnbal5", "Cbobal5", "Btnbal6", "Btnbal7") ' Frame 1
    For i = 0 To UBound(VecCntrls0)
      Let .Controls(VBA.CStr(VecCntrls0(i))).Enabled = False
    Next i
    Let VecCntrls0 = Empty
    Let .Controls("Chkbal8").Value = True
    For i = 9 To 15
      With .Controls("Chkbal" & i)
        Let .Value = True
        Let .Enabled = False
      End With
    Next i

    ' Set Tabindex
    Let VecLabels = Array("Lblbal0", "Cbobal0", "Lblbal1", "Cbobal1", "Lblbal2", "Txtbal0", "Lblbal3", "Txtbal1", "Lblbal4", "Txtbal2", "Chkbal0", "Txtbal3", "Chkbal1", "Chkbal2", "Chkbal3", "Txtbal4", "Txtbal5", "Cbobal2", "Btnbal0", "Btnbal2", "Txtbal6", "Cbobal3", "Btnbal1", "Btnbal3", "Chkbal4", "Lblbal5", "Txtbal7", "Chkbal5", "Cbobal4", "Btnbal4", "Btnbal5", "Txtbal8", "Chkbal6", "Cbobal5", "Btnbal6", "Btnbal7", "Txtbal9", "Chkbal7", "Txtbal10", "Chkbal8", "Chkbal9", "Chkbal10", "Chkbal11", "Chkbal12", "Chkbal13", "Chkbal14", "Chkbal15", "Optbal0", "Optbal1", "Txtbal11", "Btnmain0", "Btnmain1", "Btnmain2") '    For i = 0 To UBound(VecLabels)
    For i = 0 To UBound(VecLabels)
      Let .Controls(VBA.CStr(VecLabels(i))).TabIndex = i
    Next i

    ' Set Hotkeys
    Let .Controls("Lblbal0").Accelerator = "b"
    Let .Controls("Lblbal1").Accelerator = "m"
    Let .Controls("Lblbal2").Accelerator = "e"
    Let .Controls("Lblbal3").Accelerator = "o"
    Let .Controls("Lblbal4").Accelerator = "l"
    Let .Controls("Chkbal0").Accelerator = "n"

    Let .Controls("Chkbal4").Accelerator = "u"
    Let .Controls("Chkbal5").Accelerator = "i"
    Let .Controls("Chkbal6").Accelerator = "g"
    Let .Controls("Lblbal5").Accelerator = "y"
    Let .Controls("Chkbal7").Accelerator = "p"
    Let .Controls("Chkbal8").Accelerator = "t"

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
      Let .Tag = "validate" 'ControlId & "-bal"
      Let .Default = True
    End With

    '4. Events: Assing handler for each control
    Set BtnCancel = .Controls("Btnmain0") 'Main
    Set BtnAccept = .Controls("Btnmain1") 'Main
    Set Btn0 = .Controls("Btnmain2") 'Main

    Set Btn1 = .Controls("Btnbal0") ' Frame0
    Set Btn2 = .Controls("Btnbal2") ' Frame0
    Set Btn3 = .Controls("Btnbal1") ' Frame0
    Set Btn4 = .Controls("Btnbal3") ' Frame0
    Set Btn5 = .Controls("Btnbal4") ' Frame1
    Set Btn6 = .Controls("Btnbal5") ' Frame1
    Set Btn7 = .Controls("Btnbal6") ' Frame1
    Set Btn8 = .Controls("Btnbal7") ' Frame1

    Set Txt0 = .Controls("Txtbal0") ' Frame0
    Set Txt1 = .Controls("Txtbal1") ' Frame0
    Set Txt2 = .Controls("Txtbal2") ' Frame0
    Set Txt3 = .Controls("Txtbal3") ' Frame0
    Set Txt4 = .Controls("Txtbal7") ' Frame1
    Set Txt5 = .Controls("Txtbal10") ' Frame1
    Set Txt6 = .Controls("Txtbal11") ' Frame3

    Set Chk0 = .Controls("Chkbal0") ' Frame0
    Set Chk1 = .Controls("Chkbal1") ' Frame0
    Set Chk2 = .Controls("Chkbal2") ' Frame0
    Set Chk3 = .Controls("Chkbal3") ' Frame0
    Set Chk4 = .Controls("Chkbal5") ' Frame1
    Set Chk5 = .Controls("Chkbal6") ' Frame1
    Set Chk6 = .Controls("Chkbal7") ' Frame1
    Set Chk7 = .Controls("Chkbal8") ' Frame2

    Set Cbo0 = .Controls("Cbobal0") ' Frame0
    Set Cbo1 = .Controls("Cbobal2") ' Frame0
    Set Cbo2 = .Controls("Cbobal3") ' Frame0
    Set Cbo3 = .Controls("Cbobal4") ' Frame1
    Set Cbo4 = .Controls("Cbobal5") ' Frame1

    ' 5. Give styles to controls
    Call VWP.ViewFormStyles(RES, FRM)
    Set VWP = Nothing

    ' 6. [x] Show the finished Form
    Call Load(FRM)
    Call .Show

    ' 7. [x] Close form
  End With

EH:
  Let VecLabels = Empty
  Let VecCntrls0 = Empty
  Set FmeCurrent = Nothing
  Set VWP = Nothing
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Sub

Private Sub View_form_balmetselection(ByVal ControlId As String)

  Const sFN As String = "VIEW::-SelectionUIBalmet"

  On Error GoTo EH
  Call View_controls_buttons_click(Nothing, Nothing, True, ControlId)

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Sub


' Projections UI <
Private Sub View_form_projection(ByVal ControlId As String, ByVal FRM As FRMBasics)

  ' Main objects
  Dim VWP As ClassEqlViewParent
  ' Controls
  Dim FmeCurrent As MSForms.Frame
  Dim VecLabels As Variant
  Dim VecCntrls0 As Variant
  ' Standard vars
  Dim StrTitle As String, StrTitleForm As String
  Dim i As Long, j As Long, k As Long, l As Long, xLeft As Long

  Const sFN As String = "VIEW::-form_projection"

  #If Debugging Then
    Call Debug_assert(Not ControlId = VBA.vbNullString, sFN)
    Call Debug_assert(Not FRM Is Nothing, sFN)
  #End If

  On Error GoTo EH
  ' 1. [x] Get data for Form Controls

  ' 2. [x] Desing the Form
  Let StrTitleForm = "Proyección Proyectos"
  Let StrTitle = StrTitleForm & " - " & STRAPPTITLE

  Set VWP = New ClassEqlViewParent
  Call VWP.ViewFormCleanner(RES, FRM)
  With FRM

    ' Set the form dimensions
    Let .Height = 415
    Let .Width = 450
    Let .Tag = ControlId
    Let .Caption = StrTitle
    Let xLeft = .InsideWidth * 0.05 ' Left for all main controls

    ' Build and design the form
    Call VWP.ViewFormAddControl(RES, FRM, "CommandButton", "Btnmain", 2) ' All with events (Cancel, Exe and Help)
    Call VWP.ViewFormAddControl(RES, FRM, "Frame", "Framepro", 3)

    Set FmeCurrent = .Controls("Framepro0") ' ---------------------------- Framepro0'
    With FmeCurrent

      ' Set Frame properties
      Let .Caption = "Ingresar Datos"
      Let .Width = .Parent.InsideWidth * 0.9
      Let .Height = .Parent.InsideHeight * 0.35
      Let .Left = xLeft
      Let .Top = .Parent.InsideHeight * 0.04

      ' Create Frame controls
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "Label", "Lblpro", 4)
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "ComboBox", "Cbopro", 2) ' EVENT 1,2
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "TextBox", "Txtpro", 6) ' EVENT 0,3,4,5
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "CommandButton", "Btnpro", 3) 'EVENT all
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "SpinButton", "Spnpro", 1) ' EVENT all

      ' Set Frame controls layout (Vertical)
      Let VecLabels = Array("Inversión Inicial:", "Peso Fino Mineral:", "Cotización Mineral:", "Valor Residual:", "Porcentaje estimado (Impuestos, Egresos, etc.) sobre ingreso bruto:")
      ' Labels - col 1
      Let j = xLeft * 0.45
      For i = 0 To 4
        With .Controls("Lblpro" & i)
          Let .Caption = VBA.CStr(VecLabels(i))
          Let .Left = xLeft
          Let .WordWrap = False
          Let .AutoSize = True
          Let .Top = j + (.Parent.Controls("Txtpro0").Height * i) + IIf(i > 1, .Parent.Controls("Txtpro0").Height, 0)
        End With
      Next i
      Let VecLabels = Empty
      ' Texts - col 2
      Let j = .Controls("Lblpro2").Left + .Controls("Lblpro2").Width + 5
      Let k = .Controls("Lblpro4").Left + .Controls("Lblpro4").Width + 5
      Let l = (.InsideWidth - (j + xLeft)) / 3
      For i = 0 To 4
        With .Controls("Txtpro" & i)
          Let .Left = IIf(i = 4, k, j)
          Let .Top = .Parent.Controls("Lblpro" & i).Top
          Let .Width = IIf(i < 4, l, ((.Parent.InsideWidth - ((xLeft * 2) + .Parent.Controls("Lblpro4").Width + 5)) - .Height))
        End With
      Next i
      ' Texts & Cbos - col 3
      Let VecLabels = Array("Cbopro0", "Txtpro5", "Txtpro6")
      Let j = .Controls("Txtpro0").Left + .Controls("Txtpro0").Width
      For i = 0 To 2
        With .Controls(VBA.CStr(VecLabels(i)))
          Let .Left = j
          Let .Top = .Parent.Controls("Txtpro" & i).Top + IIf(i = 2, -.Height, 0)
          Let .Width = IIf(i = 0, l * 2, l)
        End With
      Next i
      Let VecLabels = Empty
      ' Cbos - col 4
      Let j = .Controls("Txtpro5").Left + .Controls("Txtpro5").Width
      Let .Controls("Cbopro1").Left = j
      Let .Controls("Cbopro1").Top = .Controls("Txtpro5").Top
      Let .Controls("Cbopro1").Width = l - (.Controls("Cbopro1").Height * 2)
      Let .Controls("Cbopro2").Left = j
      Let .Controls("Cbopro2").Top = .Controls("Txtpro6").Top
      Let .Controls("Cbopro2").Width = l - (.Controls("Cbopro2").Height * 2)
      ' Btns & Spns - col 5 & 6
      Let j = .Controls("Cbopro1").Left + .Controls("Cbopro1").Width
      Let .Controls("Btnpro0").Width = .Controls("Cbopro1").Height
      Let .Controls("Btnpro0").Height = .Controls("Cbopro1").Height
      Let .Controls("Btnpro0").Left = j
      Let .Controls("Btnpro0").Top = .Controls("Cbopro1").Top
      Let .Controls("Btnpro2").Width = .Controls("Cbopro2").Height
      Let .Controls("Btnpro2").Height = .Controls("Cbopro2").Height
      Let .Controls("Btnpro2").Left = j
      Let .Controls("Btnpro2").Top = .Controls("Cbopro2").Top
      Let j = .Controls("Btnpro0").Left + .Controls("Btnpro0").Width
      Let .Controls("Btnpro1").Width = .Controls("Btnpro0").Height
      Let .Controls("Btnpro1").Height = .Controls("Btnpro0").Height
      Let .Controls("Btnpro1").Left = j
      Let .Controls("Btnpro1").Top = .Controls("Btnpro0").Top
      Let .Controls("Btnpro3").Width = .Controls("Btnpro0").Height
      Let .Controls("Btnpro3").Height = .Controls("Btnpro0").Height
      Let .Controls("Btnpro3").Left = j
      Let .Controls("Btnpro3").Top = .Controls("Btnpro2").Top
      Let j = .Controls("Txtpro4").Left + .Controls("Txtpro4").Width
      Let .Controls("Spnpro0").Height = .Controls("Txtpro4").Height
      Let .Controls("Spnpro0").Width = .Controls("Txtpro4").Height * 0.5
      Let .Controls("Spnpro0").Left = j
      Let .Controls("Spnpro0").Top = .Controls("Txtpro4").Top
      Let .Controls("Spnpro1").Height = .Controls("Spnpro0").Height
      Let .Controls("Spnpro1").Width = .Controls("Txtpro4").Height * 0.5
      Let .Controls("Spnpro1").Left = j + .Controls("Spnpro0").Width
      Let .Controls("Spnpro1").Top = .Controls("Spnpro0").Top
      ' Set Frame controls ToolTipTexts
      Let VecLabels = Array( _
          "Valor numérico de la inversión inicial del proyecto", _
          "Un número. Para varios minerales, valores numéricos separados por punto y coma Ej.: 1;2;3", _
          "Valores numéricos separados por punto y coma Ej.: 1;2;3. Para varios minerales, haga grupos Ej.: 1;2;3/4;5;6", _
          "Seleccione el valor en una celda o introduzca un valor numérico", _
          "Seleccione la divisa con que trabaja el proyecto", _
          "Seleccione un mineral o elemento", _
          "Elimina del cuadro de texto el último mineral o elemento seleccionado", _
          "Elimina del cuadro de texto todos los minerals o elementos seleccionados", _
          "Seleccione una unidad de peso para el mineral", _
          "Elimina del cuadro de texto la última unidad seleccionada", _
          "Elimina del cuadro de texto todas las unidades seleccionadas", _
          "Un número entre 0 y 100")
      Let VecCntrls0 = Array("Txtpro0", "Txtpro1", "Txtpro2", "Txtpro3", "Cbopro0", "Cbopro1", "Btnpro0", "Btnpro1", "Cbopro2", "Btnpro2", "Btnpro3", "Txtpro4")
      For i = 0 To UBound(VecLabels)
        Let .Controls(VBA.CStr(VecCntrls0(i))).ControlTipText = VBA.CStr(VecLabels(i))
      Next i
      Let VecCntrls0 = Empty
      Let VecLabels = Empty

    End With
    Set FmeCurrent = Nothing

    Set FmeCurrent = .Controls("Framepro1") ' ---------------------------- Framepro1'
    With FmeCurrent

      ' Set Frame properties
      Let .Caption = "Parámetros Proyección"
      Let .Width = .Parent.InsideWidth * 0.9
      Let .Height = .Parent.InsideHeight * 0.2
      Let .Left = xLeft
      Let .Top = .Parent.Controls("Framepro0").Top + .Parent.Controls("Framepro0").Height + 5

      ' Create Frame controls
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "Label", "Lblpro", 6, CtlFromX:=5)
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "TextBox", "Txtpro", 8, CtlFromX:=7) ' EVENT 9,10
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "SpinButton", "Spnpro", 4, CtlFromX:=2) ' EVENT all
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "ComboBox", "Cbopro", 3, CtlFromX:=3) ' EVENT all

      ' Set Frame controls layout
      Let VecLabels = Array("Tiempo de proyección:", "Tasa de ganancia requerida:")
      ' Labels - col 1
      Let j = xLeft * 0.85
      For i = 5 To 6
        With .Controls("Lblpro" & i)
          Let .WordWrap = False
          Let .Caption = VBA.CStr(VecLabels(i - 5))
          Let .AutoSize = True
          Let .Left = xLeft
          Let .Top = j + (.Parent.Controls("Txtpro7").Height * (i - 5))
        End With
      Next i
      Let VecLabels = Empty
      ' Texts - col 2
      Let j = (.InsideWidth - (.Controls("Lblpro6").Left + .Controls("Lblpro6").Width + xLeft + 5)) / 3
      For i = 7 To 8
        With .Controls("Txtpro" & i)
          Let .Left = .Parent.Controls("Lblpro6").Left + .Parent.Controls("Lblpro6").Width + 5
          Let .Top = .Parent.Controls("Lblpro" & i - 2).Top
          Let .Width = j * 2
        End With
      Next i
      ' Spns - col 3
      For i = 2 To 3
        With .Controls("Spnpro" & i)
          Let .Height = .Parent.Controls("Txtpro7").Height
          Let .Width = .Height * 0.5
          Let .Left = .Parent.Controls("Txtpro7").Left + .Parent.Controls("Txtpro7").Width
          Let .Top = .Parent.Controls("Txtpro" & i + 5).Top
        End With
      Next i
      ' Spns & Cbo - col 4
      Let VecLabels = Array("Cbopro3", "Spnpro4")
      For i = 0 To 1
        With .Controls(VBA.CStr(VecLabels(i)))
          Let .Height = .Parent.Controls("Cbopro3").Height
          Let .Width = IIf(i = 0, j - .Parent.Controls("Spnpro2").Width, .Height * 0.5)
          Let .Left = .Parent.Controls("Spnpro2").Left + .Parent.Controls("Spnpro2").Width
          Let .Top = .Parent.Controls("Spnpro" & i + 2).Top
        End With
      Next i
      Let VecLabels = Empty

      ' Set Frame controls ToolTipTexts
      Let VecLabels = Array("Txtpro7", "Txtpro8", "Cbopro3", "Valor numérico", "Valor numérico entre 0 y 100%", "Seleccione una unidad de tiempo")
      For i = 0 To 2
        Let .Controls(VBA.CStr(VecLabels(i))).ControlTipText = VBA.CStr(VecLabels(i + 3))
      Next i
      Let VecLabels = Empty

    End With
    Set FmeCurrent = Nothing

    Set FmeCurrent = .Controls("Framepro2") ' ---------------------------- Framepro2'
    With FmeCurrent

      ' Set Frame properties
      Let .Caption = "Incluir en resultado"
      Let .Width = ((.Parent.InsideWidth * 0.9) / 2) - 2.5
      Let .Height = .Parent.InsideHeight * 0.2
      Let .Left = xLeft
      Let .Top = .Parent.Controls("Framepro1").Top + .Parent.Controls("Framepro1").Height + 5

      ' Create Frame controls
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "CheckBox", "Chkpro", 2, CtlFromX:=0)
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "OptionButton", "Optpro", 1, CtlFromX:=0)

      ' Set Frame controls properties (including toolTipTexts and values)
      Let VecLabels = Array("Permite imprimir en el resultado el Valor Anual Neto", "Permite imprimir la Tasa Interna de Retorno", "Active la casilla para generar el gráfico del resultado", "Seleccione para Imprimir el gráfico en la hoja actual")
      Let VecCntrls0 = Array("Chkpro0", "Chkpro1", "Chkpro2", "Optpro0", "Mostrar VAN", "Mostrar TIR", "Generar gráfico en:", "Hoja actual")
      For i = 0 To 3
        With .Controls(VBA.CStr(VecCntrls0(i)))
          Let .WordWrap = False
          Let .Caption = VBA.CStr(VecCntrls0(i + 4))
          Let .AutoSize = True
          Let .ControlTipText = VecLabels(i)
          Let .Left = xLeft * 0.5
          Let .Top = (xLeft * 0.5) + ((.Parent.Controls("Chkpro0").Height * 0.8) * i)
        End With
      Next i
      With .Controls("Optpro1")
        Let .WordWrap = False
        Let .Caption = "Nueva hoja de gráfico"
        Let .AutoSize = True
        Let .ControlTipText = "Seleccione para obtener el gráfico del ejercicio en una hoja de gráfico independiente"
        Let .Left = .Parent.Controls("Optpro0").Left + .Parent.Controls("Optpro0").Width + 5
        Let .Top = .Parent.Controls("Optpro0").Top
      End With

    End With
    Set FmeCurrent = Nothing

    Set FmeCurrent = .Controls("Framepro3") ' ---------------------------- Framepro3'
    With FmeCurrent

      ' Set Frame properties
      Let .Caption = "Dónde mostrar resultados"
      Let .Width = ((.Parent.InsideWidth * 0.9) / 2) - 2.5
      Let .Height = .Parent.InsideHeight * 0.2
      Let .Left = .Parent.Controls("Framepro2").Left + .Parent.Controls("Framepro2").Width + 5
      Let .Top = .Parent.Controls("Framepro1").Top + .Parent.Controls("Framepro1").Height + 5

      ' Create Frame controls
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "OptionButton", "Optpro", 3, CtlFromX:=2)
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "TextBox", "Txtpro", 9, CtlFromX:=9) ' EVENT

      ' Set Frame controls properties (including values)
      Let .Controls("Optpro2").Caption = "Hoja nueva"
      Let .Controls("Optpro3").Caption = "Hoja actual"

      Let .Controls("Optpro3").AutoSize = True
      Let .Controls("Optpro2").AutoSize = True
      Let .Controls("Optpro3").WordWrap = False
      Let .Controls("Optpro2").WordWrap = False

      Let .Controls("Optpro3").Left = xLeft
      Let .Controls("Optpro2").Left = xLeft
      Let .Controls("Optpro2").Top = xLeft / 1.3
      Let .Controls("Optpro3").Top = (xLeft / 1.3) + .Controls("Optpro3").Height

      Let j = .Width - (xLeft * 2) - .Controls("Optpro3").Width
      Let .Controls("Txtpro9").Left = .Controls("Optpro3").Left + .Controls("Optpro3").Width
      Let .Controls("Txtpro9").Top = .Controls("Optpro3").Top
      Let .Controls("Txtpro9").Width = j

      ' Set Frame controls ToolTipTexts
      Let .Controls("Txtpro9").ControlTipText = "Seleccionar una celda para mostrar el resultado"

    End With
    Set FmeCurrent = Nothing
    Let VecLabels = Empty
    Let VecCntrls0 = Empty

    ' Set Form textbox controls dropbuttons
    Let VecLabels = Array(0, 1, 2, 3, 9)
    For i = 0 To UBound(VecLabels)
      With .Controls("Txtpro" & VecLabels(i))
        Let .SpecialEffect = fmSpecialEffectRaised
        Let .DropButtonStyle = fmDropButtonStyleReduce
        Let .ShowDropButtonWhen = fmShowDropButtonWhenAlways
      End With
    Next i
    Let VecLabels = Empty

    ' Set Form controls values
    ' Combos
    Let .Controls("Cbopro0").List = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigcurrencies_sys, eqlMdlArray, "symbol"))
    Let .Controls("Cbopro1").List = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalelements_sys, eqlMdlArray, "Símbolo"))
    Call .Controls("Cbopro1").AddItem("---", 0)
    Let .Controls("Cbopro2").List = Array("TF", "LF", "OT", "KF", "TC", "TL")
    Call .Controls("Cbopro2").AddItem("---", 0)
    Let .Controls("Cbopro3").List = Array("Días", "Semanas", "Quincenas", "Meses", "Años")
    Let .Controls("Cbopro0").ListIndex = 0
    Let .Controls("Cbopro1").ListIndex = 0
    Let .Controls("Cbopro2").ListIndex = 0
    Let .Controls("Cbopro3").ListIndex = 0
    ' Texts
    Let .Controls("Txtpro5").Locked = True
    Let .Controls("Txtpro6").Locked = True
    Let .Controls("Txtpro5").ForeColor = VBA.vbGrayText
    Let .Controls("Txtpro6").ForeColor = VBA.vbGrayText
    With .Controls("Txtpro4")
      Let .Value = VBA.Format(2, "0.00""%""")
      Let .MaxLength = 6
      Let .TextAlign = fmTextAlignRight
    End With
    With .Controls("Txtpro7")
      Let .Value = VBA.Format(1, "0")
      Let .MaxLength = 2
      Let .TextAlign = fmTextAlignRight
    End With
    With .Controls("Txtpro8")
      Let .Value = VBA.Format(50, "0.00""%""")
      Let .MaxLength = 6
      Let .TextAlign = fmTextAlignRight
    End With
    ' Checks & Options
    Let .Controls("Chkpro0").Value = True
    Let .Controls("Chkpro1").Value = True
    Let .Controls("Optpro0").Value = True
    Let .Controls("Optpro2").Value = True
    ' Buttons
    For i = 0 To 3
      Let .Controls("Btnpro" & i).Caption = IIf(i Mod 2 = 0, VBA.ChrW(&H2A31), VBA.ChrW(&H2716)) ' Frame 0
      Let .Controls("Btnpro" & i).Font.Bold = True
    Next i

    ' Set Tabindex
    Let VecLabels = Array("Lblpro0", "Txtpro0", "Cbopro0", "Lblpro1", "Txtpro1", "Txtpro5", "Cbopro1", "Btnpro0", "Btnpro1", "Txtpro6", "Cbopro2", "Btnpro2", "Btnpro3", "Lblpro2", "Txtpro2", "Lblpro3", "Txtpro3", "Lblpro4", "Txtpro4", "Spnpro0", "Spnpro1", "Lblpro5", "Txtpro7", "Spnpro2", "Cbopro3", "Lblpro6", "Txtpro8", "Spnpro3", "Spnpro4", "Chkpro0", "Chkpro1", "Chkpro2", "Optpro0", "Optpro1", "Optpro2", "Optpro3", "Txtpro9", "Btnmain0", "Btnmain1", "Btnmain2")
    For i = 0 To UBound(VecLabels)
      Let .Controls(VBA.CStr(VecLabels(i))).TabIndex = i
    Next i

    ' Set Hotkeys
    Let .Controls("Lblpro0").Accelerator = "i"
    Let .Controls("Lblpro1").Accelerator = "p"
    Let .Controls("Lblpro2").Accelerator = "l"
    Let .Controls("Lblpro3").Accelerator = "d"
    Let .Controls("Lblpro4").Accelerator = "r"
    Let .Controls("Lblpro5").Accelerator = "e"
    Let .Controls("Lblpro6").Accelerator = "s"

    Let .Controls("Chkpro0").Accelerator = "v"
    Let .Controls("Chkpro1").Accelerator = "t"
    Let .Controls("Chkpro2").Accelerator = "g"
    Let .Controls("Optpro0").Accelerator = "h"
    Let .Controls("Optpro1").Accelerator = "n"

    Let .Controls("Optpro2").Accelerator = "o"
    Let .Controls("Optpro3").Accelerator = "u"

    ' 3. [x] Fill Form Controls with data and Set general purpose CommandButtons
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
      Let .Tag = "validate" 'ControlId & "-pro"
      Let .Default = True
    End With

    ' 4. Events: Assing handler for each control
    Set BtnCancel = .Controls("Btnmain0") 'Main
    Set BtnAccept = .Controls("Btnmain1") 'Main
    Set Btn0 = .Controls("Btnmain2") 'Main

    Set Btn1 = .Controls("Btnpro0") ' Frame0
    Set Btn2 = .Controls("Btnpro1") ' Frame0
    Set Btn3 = .Controls("Btnpro2") ' Frame0
    Set Btn4 = .Controls("Btnpro3") ' Frame0

    Set Txt0 = .Controls("Txtpro0") ' Frame0
    Set Txt1 = .Controls("Txtpro1") ' Frame0
    Set Txt2 = .Controls("Txtpro2") ' Frame0
    Set Txt3 = .Controls("Txtpro3") ' Frame0
    Set Txt4 = .Controls("Txtpro9") ' Frame3

    Set Cbo0 = .Controls("Cbopro1") ' Frame0
    Set Cbo1 = .Controls("Cbopro2") ' Frame0

    Set Spn0 = .Controls("Spnpro0") 'Frame0
    Set Spn1 = .Controls("Spnpro1") 'Frame0
    Set Spn2 = .Controls("Spnpro2") 'Frame1
    Set Spn3 = .Controls("Spnpro3") 'Frame1
    Set Spn4 = .Controls("Spnpro4") 'Frame1

    ' 5. Give styles to controls
    Call VWP.ViewFormStyles(RES, FRM)
    Set VWP = Nothing

    ' 6. [x] Show the finished Form
    Call Load(FRM)
    Call .Show

    ' 7. [x] Close form
  End With

EH:
  Let VecLabels = Empty
  Let VecCntrls0 = Empty
  Set FmeCurrent = Nothing
  Set VWP = Nothing
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Sub

' Purchase Mineral UI
Private Sub View_form_purchasemineral(ByVal ControlId As String, ByVal FRM As FRMBasics)

  Dim VWP As ClassEqlViewParent
  ' Controls
  Dim FmeCurrent As MSForms.Frame
  Dim VecLabels As Variant, VecCntrls0 As Variant
  ' Standard vars
  Dim StrTitle As String, StrTitleForm As String, StrDivlocal As String, StrDivforeign As String
  Dim i As Long, j As Long, k As Long, l As Long, xLeft As Long

  Const sFN As String = "VIEW::-form_purchasemineral"

  #If Debugging Then
    Call Debug_assert(Not ControlId = VBA.vbNullString, sFN)
    Call Debug_assert(Not FRM Is Nothing, sFN)
  #End If

  On Error GoTo EH
  ' 0. [x] Get data for Form Controls

  ' 1. [x] Desing the Form
  Let StrTitleForm = "Liquidación de Minerales"
  Let StrTitle = StrTitleForm & " - " & STRAPPTITLE

  Set VWP = New ClassEqlViewParent
  Call VWP.ViewFormCleanner(RES, FRM)
  With FRM

    ' Set the form dimensions
    Let .Height = 415
    Let .Width = 450
    Let .Tag = ControlId
    Let .Caption = StrTitle
    Let xLeft = .InsideWidth * 0.05 ' Left for all main controls

    ' Build and design the form
    Call VWP.ViewFormAddControl(RES, FRM, "CommandButton", "Btnmain", 2) ' All with events (Cancel, Exe and Help)
    Call VWP.ViewFormAddControl(RES, FRM, "Frame", "Framepur", 3)

    Set FmeCurrent = .Controls("Framepur0") ' ---------------------------- Framepur0'
    With FmeCurrent

      ' Set Frame properties
      Let .Caption = "Datos base"
      Let .Width = (.Parent.InsideWidth * 0.9)
      Let .Height = .Parent.InsideHeight * 0.15
      Let .Left = xLeft
      Let .Top = .Parent.InsideHeight * 0.04

      ' Create Frame controls
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "Label", "Lblpur", 2)
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "TextBox", "Txtpur", 2) ' EVENT 0,2
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "CheckBox", "Chkpur", 0)

      'Set Frame controls layout
      ' Labels - col 1
      Let j = (.InsideWidth - (xLeft * 2) - 10) / 4
      For i = 0 To 1
        With .Controls("Lblpur" & i)
          Let .WordWrap = False
          Let .Caption = IIf(i = 0, "Lugar:", "Fecha:")
          Let .AutoSize = True
          Let .Left = xLeft
          Let .Top = (xLeft * 0.35) + IIf(i = 0, 0, .Parent.Controls("Txtpur0").Height)
        End With
      Next i
      ' Txts - col 2
      For i = 0 To 1
        Let .Controls("Txtpur" & i).Left = xLeft + .Controls("Lblpur0").Width + 5
        Let .Controls("Txtpur" & i).Top = .Controls("Lblpur" & i).Top
        Let .Controls("Txtpur" & i).Width = j + (j - .Controls("Lblpur0").Width)
      Next i
      ' Lbls & Chks - col 3
      Let StrDivlocal = MDL.CurrencyMainSymbolLocal
      Let StrDivforeign = MDL.CurrencyMainSymbolForeign
      For i = 0 To 1
        Let .Controls(IIf(i = 0, "Lblpur2", "Chkpur0")).WordWrap = False
        Let .Controls(IIf(i = 0, "Lblpur2", "Chkpur0")).Caption = IIf(i = 0, "TC [" & StrDivlocal & "/" & StrDivforeign & "]:", "Imprimir Moneda Extranjera")
        Let .Controls(IIf(i = 0, "Lblpur2", "Chkpur0")).AutoSize = True
        Let .Controls(IIf(i = 0, "Lblpur2", "Chkpur0")).Left = (.Controls("Txtpur" & i).Left + .Controls("Txtpur" & i).Width) + 5 'xLeft
        Let .Controls(IIf(i = 0, "Lblpur2", "Chkpur0")).Top = .Controls("Lblpur" & i).Top
      Next i
      Let .Controls("Txtpur2").Left = .Controls("Lblpur2").Left + .Controls("Lblpur2").Width + 5
      Let .Controls("Txtpur2").Top = .Controls("Txtpur0").Top
      Let .Controls("Txtpur2").Width = j + (j - .Controls("Lblpur2").Width)

      ' Set Frame controls ToolTipTexts
      Let VecLabels = Array( _
          "Orígen/localidad/ciudad donde se realiza este documento. Ej.: Potosí", _
          "Seleccione o introduzca una fecha, formato: dd/mm/aa", _
          "Tipo de cambio: una cifra monetaria equivalente al tipo de cambio actual", _
          "Si activa la casilla, se creará en el documento una columna con valores en moneda extranjera al tipo de cambio dado")
      Let VecCntrls0 = Array("Txtpur0", "Txtpur1", "Txtpur2", "Chkpur0")
      For i = 0 To UBound(VecLabels)
        Let .Controls(VBA.CStr(VecCntrls0(i))).ControlTipText = VBA.CStr(VecLabels(i))
      Next i
      Let VecCntrls0 = Empty
      Let VecLabels = Empty

    End With
    Set FmeCurrent = Nothing

    Set FmeCurrent = .Controls("Framepur1") ' ---------------------------- Framepur1'
    With FmeCurrent

      ' Set Frame properties
      Let .Caption = "Datos Provedor"
      Let .Width = (.Parent.InsideWidth * 0.9)
      Let .Height = .Parent.InsideHeight * 0.15
      Let .Left = .Parent.Controls("Framepur0").Left
      Let .Top = .Parent.Controls("Framepur0").Top + .Parent.Controls("Framepur0").Height + 5

      ' Create Frame controls
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "Label", "Lblpur", 4, CtlFromX:=3)
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "TextBox", "Txtpur", 4, CtlFromX:=3) ' EVENT all
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "CheckBox", "Chkpur", 1, CtlFromX:=1) ' EVENT 1
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "ComboBox", "Cbopur", 0) ' EVENT all

      'Set Frame controls layout
      ' Chk - col 1
      Let .Controls("Chkpur1").WordWrap = False
      Let .Controls("Chkpur1").Caption = "Varios provedores" '"Varias liquidaciones"
      Let .Controls("Chkpur1").AutoSize = True
      Let .Controls("Chkpur1").Left = xLeft
      Let .Controls("Chkpur1").Top = (xLeft * 0.35) + .Controls("Txtpur3").Height
      ' Lbl & Txt - col 2
      Let j = (.InsideWidth - (.Controls("Chkpur1").Width + (xLeft * 2) + 5)) * 0.5
      For i = 0 To 1
        Let .Controls(IIf(i = 0, "Lblpur3", "Txtpur3")).Left = .Controls("Chkpur1").Left + .Controls("Chkpur1").Width + 5
        Let .Controls(IIf(i = 0, "Lblpur3", "Txtpur3")).Top = .Controls("Chkpur1").Top - IIf(i = 1, 0, .Controls("Txtpur3").Height)
        Let .Controls(IIf(i = 0, "Lblpur3", "Txtpur3")).Width = j
      Next i
      ' Lbl & Txt - col 3
      For i = 0 To 1
        Let .Controls(IIf(i = 0, "Lblpur4", "Txtpur4")).Left = .Controls("Txtpur3").Left + .Controls("Txtpur3").Width + 5
        Let .Controls(IIf(i = 0, "Lblpur4", "Txtpur4")).Top = .Controls("Chkpur1").Top - IIf(i = 1, 0, .Controls("Txtpur3").Height)
        Let .Controls(IIf(i = 0, "Lblpur4", "Txtpur4")).Width = j
      Next i
      Let .Controls("Cbopur0").Left = .Controls("Txtpur4").Left
      Let .Controls("Cbopur0").Top = .Controls("Txtpur4").Top
      Let .Controls("Cbopur0").Width = .Controls("Txtpur4").Width

      ' Set Frame controls ToolTipTexts
      Let VecLabels = Array( _
          "Generar varias liquidaciones para uno o varios provedores", _
          "Nombres y apellidos del provedor (o provedores)", _
          "Seleccione un rango de celdas o escriba los valores correspondientes separados por punto y coma. Ej.: 'Particular;Particular;Cooperativista'", _
          "Seleccione una opción de la lista")
      Let VecCntrls0 = Array("Chkpur1", "Txtpur3", "Txtpur4", "Cbopur0")
      For i = 0 To UBound(VecLabels)
        Let .Controls(VBA.CStr(VecCntrls0(i))).ControlTipText = VBA.CStr(VecLabels(i))
      Next i
      Let VecCntrls0 = Empty
      Let VecLabels = Empty

    End With
    Set FmeCurrent = Nothing

    Set FmeCurrent = .Controls("Framepur2") ' ---------------------------- Framepur2'
    With FmeCurrent

      ' Set Frame properties
      Let .Caption = "Datos Liquidación"
      Let .Width = (.Parent.InsideWidth * 0.9)
      Let .Height = .Parent.InsideHeight * 0.4
      Let .Left = .Parent.Controls("Framepur1").Left
      Let .Top = .Parent.Controls("Framepur1").Top + .Parent.Controls("Framepur1").Height + 5

      ' Create Frame controls
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "TextBox", "Txtpur", 9, CtlFromX:=5) ' EVENT all
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "CheckBox", "Chkpur", 5, CtlFromX:=2) ' EVENT 2
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "ComboBox", "Cbopur", 4, CtlFromX:=1, CtlVisible:=False) ' EVENT all
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "CommandButton", "Btnpur", 5, CtlFromX:=0) ' EVENT all
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "ListBox", "Lstpur", 0) ' EVENT all

      'Set Frame controls layout
      ' Cbo & Btns - col 1 (unique col) - Row 1
      Let .Controls("Cbopur1").Left = xLeft
      Let .Controls("Cbopur1").Top = xLeft * 0.5
      Let .Controls("Cbopur1").Width = (.InsideWidth - ((.Controls("Cbopur1").Height * 6) + (xLeft * 2))) + 5
      Let .Controls("Cbopur1").Visible = True
      For i = 0 To 5
        Let .Controls("Btnpur" & i).Height = .Controls("Cbopur1").Height
        Let .Controls("Btnpur" & i).Width = .Controls("Cbopur1").Height
        Let .Controls("Btnpur" & i).Top = .Controls("Cbopur1").Top
        Let .Controls("Btnpur" & i).Left = (.Controls("Cbopur1").Left + .Controls("Cbopur1").Width) + (.Controls("Btnpur" & i).Width * i)
      Next i
      Let .Controls("Btnpur2").Enabled = False
      Let .Controls("Btnpur3").Enabled = False
      Let .Controls("Btnpur4").Enabled = False

      ' Txts - col 1 (unique col) - Row 2
      For i = 5 To 9
        Let .Controls("Txtpur" & i).Top = (.Controls("Cbopur1").Top + .Controls("Cbopur1").Height) + 2.5
        Let .Controls("Txtpur" & i).Width = (.InsideWidth - (xLeft * 2) + 5) / 5
        Let .Controls("Txtpur" & i).Left = (.Controls("Cbopur1").Left) + (.Controls("Txtpur" & i).Width * (i - 5))
      Next i
      Let .Controls("Chkpur2").Left = .Controls("Txtpur6").Left + (xLeft * 0.5)
      Let .Controls("Chkpur2").Top = .Controls("Txtpur6").Top
      Let .Controls("Chkpur2").WordWrap = False
      Let .Controls("Chkpur2").Caption = "Todas las "
      Let .Controls("Chkpur2").AutoSize = True
      Let .Controls("Chkpur2").Visible = False

      ' Lst - col 1 (unique col) - Row 3
      Let .Controls("Lstpur0").Width = (.InsideWidth - (xLeft * 2) + 5)
      Let .Controls("Lstpur0").Left = xLeft
      Let .Controls("Lstpur0").Top = .Controls("Txtpur5").Top + .Controls("Txtpur5").Height
      Let .Controls("Lstpur0").Height = (.InsideHeight - ((.Controls("Cbopur1").Height * 3) + (xLeft * 0.5) + 5))

      ' Chks - col 1 (unique col) - Row 4
      For i = 3 To 5
        With .Controls("Chkpur" & i)
          Let .WordWrap = False
          Let .Width = (.Parent.InsideWidth - (xLeft * 2) + 5) / 3
          Let .Top = (.Parent.InsideHeight - xLeft)
          Let .Left = (xLeft + (.Width * (i - 3)))
        End With
      Next i

      ' Set Frame controls ToolTipTexts
      Let VecLabels = Array( _
          "Seleccione una opción de la lista", _
          "Añadir los valores de los controles a la lista", _
          "Confirmar edición", _
          "Cancelar edición", _
          "Eliminar fila de datos seleccionada", _
          "Eliminar todos los datos ingresados", _
          "Seleccionar datos desde un rango de celdas de " & Application.Name, _
          "Permite generar leyes, %Humedad y %Merma promediadas en la Liquidación Final", _
          "Doble click en el ítem que desee modificar", _
          "El sistema configurará las páginas donde imprimirá la Liquidación Final", _
          "Permite generar cada Liquidación Final con su copia, ambas en una página independiente en orientación horizontal")
      Let VecCntrls0 = Array("Cbopur1", "Btnpur0", "Btnpur1", "Btnpur2", "Btnpur3", "Btnpur4", "Btnpur5", "Lstpur0", "Chkpur3", "Chkpur4", "Chkpur5")
      For i = 0 To UBound(VecLabels)
        Let .Controls(VBA.CStr(VecCntrls0(i))).ControlTipText = VBA.CStr(VecLabels(i))
      Next i
      Let VecCntrls0 = Empty
      Let VecLabels = Empty

    End With
    Set FmeCurrent = Nothing

    Set FmeCurrent = .Controls("Framepur3") ' ---------------------------- Framepur3'
    With FmeCurrent

      ' Set Frame properties
      Let .Caption = "Dónde mostrar resultados"
      Let .Width = .Parent.Controls("Framepur2").Width
      Let .Height = .Parent.InsideHeight * 0.1
      Let .Left = .Parent.Controls("Framepur2").Left
      Let .Top = .Parent.Controls("Framepur2").Top + .Parent.Controls("Framepur2").Height + 5

      ' Create Frame controls
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "TextBox", "Txtpur", 10, True, 10) ' EVENT all
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "OptionButton", "Optpur", 1)

      'Set Frame controls layout
      ' Opts & Txt - Row 1
      Let j = (.InsideWidth - (xLeft * 2) + 5) / 3
      Let .Controls("Optpur0").WordWrap = False
      Let .Controls("Optpur0").Width = j
      Let .Controls("Optpur0").Left = xLeft
      Let .Controls("Optpur0").Top = xLeft * 0.35
      Let .Controls("Optpur1").WordWrap = False
      Let .Controls("Optpur1").Width = j
      Let .Controls("Optpur1").Left = xLeft + .Controls("Optpur1").Width
      Let .Controls("Optpur1").Top = .Controls("Optpur0").Top
      Let .Controls("Txtpur10").Left = .Controls("Optpur1").Left + .Controls("Optpur1").Width
      Let .Controls("Txtpur10").Top = .Controls("Optpur0").Top

      ' Set Frame controls ToolTipTexts
      Let .Controls("Optpur0").ControlTipText = STRAPPTITLE & " creará una hoja de cálculo nueva en el libro actual"
      Let .Controls("Optpur1").ControlTipText = "Active esta opción para generar el resultado en la hoja actual"
      Let .Controls("Txtpur10").ControlTipText = "Seleccionar una celda"

    End With

    ' 2. [x] Fill controls with values
    ' Frame0
    Let .Controls("Txtpur0").Text = "N/D"
    Let .Controls("Txtpur1").Text = VBA.Format(Date, "dd/mm/yy")
    Let .Controls("Txtpur1").TextAlign = fmTextAlignRight
    Let .Controls("Txtpur2").Text = "1"
    ' Frame1
    Let .Controls("Cbopur0").List = Array("Cooperativista", "Particular")
    Let .Controls("Cbopur0").ListIndex = 0
    ' Frame2
    Let .Controls("Lblpur3").Caption = "Provedor:"
    Let .Controls("Lblpur4").Caption = "Tipo Provedor:"
    Let .Controls("Cbopur1").List = Array("Datos base del mineral", "Datos de calidad y valor del mineral", "Bonos", "Retenciones de ley", "Otras retenciones", "Descuentos")
    Let .Controls("Cbopur1").ListIndex = -1
    Let .Controls("Cbopur1").ListStyle = fmListStyleOption
    Let VecLabels = Array(&H271A, &H2710, &H270F, &H2A31, &H2716, &H2583)
    For i = 0 To 5
      Let .Controls("Btnpur" & i).Caption = VBA.ChrW(VecLabels(i))
      Let .Controls("Btnpur" & i).Font.Bold = True
      If Not i = 2 And Not i = 5 Then Let .Controls("Btnpur" & i).Tag = "validate"
    Next i
    Let .Controls("Chkpur3").Caption = "Promediar valores"
    Let .Controls("Chkpur4").Caption = "Autoconfigurar página"
    Let .Controls("Chkpur5").Caption = "Incluir copia"
    Let .Controls("Cbopur2").List = Array("Mineral", "Mineral Complejo", "Brosa")
    Let .Controls("Cbopur2").ListIndex = 0
    Let .Controls("Cbopur4").List = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalunits_sys, eqlMdlArray)) 'Array("%", "DM", "OT", "g/T", "g/L")
    Call .Controls("Cbopur4").AddItem("---", 0)
    Let .Controls("Cbopur4").ListIndex = 0
    Let .Controls("Txtpur6").Text = VBA.Format(Date, "dd/mm/yy")
    Let .Controls("Lstpur0").ColumnHeads = True
    ' Frame3
    Let .Controls("Optpur0").Value = True
    Let .Controls("Optpur0").Caption = "Hoja nueva"
    Let .Controls("Optpur1").Caption = "Hoja actual"
    Let .Controls("Optpur1").AutoSize = True
    Let .Controls("Txtpur10").Width = j + (j - .Controls("Optpur1").Width)
    Let .Controls("Txtpur10").Left = .Controls("Optpur1").Left + .Controls("Optpur1").Width

    ' Set Tabindex
    Let VecLabels = Array("Lblpur0", "Txtpur0", "Lblpur1", "Txtpur1", "Lblpur2", "Txtpur2", "Chkpur0", "Chkpur1", "Lblpur3", "Txtpur3", "Lblpur4", "Cbopur0", "Txtpur4", "Cbopur1", "Btnpur0", "Btnpur1", "Btnpur2", "Btnpur3", "Btnpur4", "Btnpur5", "Txtpur5", "Txtpur6", "Cbopur2", "Chkpur2", "Txtpur7", "Txtpur8", "Txtpur9", "Lstpur0", "Chkpur3", "Chkpur4", "Chkpur5", "Optpur0", "Optpur1", "Txtpur10", "Btnmain0", "Btnmain1", "Btnmain2")
    For i = 0 To UBound(VecLabels)
      Let .Controls(VBA.CStr(VecLabels(i))).TabIndex = i
    Next i

    ' Set Hotkeys
    Let .Controls("Lblpur0").Accelerator = "l"
    Let .Controls("Lblpur1").Accelerator = "f"
    Let .Controls("Lblpur2").Accelerator = "t"
    Let .Controls("Lblpur3").Accelerator = "v"
    Let .Controls("Lblpur4").Accelerator = "o"

    Let .Controls("Chkpur0").Accelerator = "i"
    Let .Controls("Chkpur1").Accelerator = "s"
    Let .Controls("Chkpur3").Accelerator = "p"
    Let .Controls("Chkpur4").Accelerator = "u"
    Let .Controls("Chkpur5").Accelerator = "r"

    Let .Controls("Optpur0").Accelerator = "h"
    Let .Controls("Optpur1").Accelerator = "j"

    ' Set Form textbox controls dropbuttons
    For i = 0 To 10
      With .Controls("Txtpur" & i)
        Let .SpecialEffect = fmSpecialEffectRaised
        Let .DropButtonStyle = IIf(i = 1 Or i = 6, fmDropButtonStylePlain, fmDropButtonStyleReduce)
        Let .ShowDropButtonWhen = fmShowDropButtonWhenAlways
      End With
    Next i

    ' 3. [x] Fill Form Controls with data and Set general purpose CommandButtons
    Let k = (.InsideWidth - (xLeft * 2))
    With .Controls("Btnmain0") '("BtnCancel")
      Let .Accelerator = "c"
      Let .Cancel = True
      Let .Top = (.Parent.InsideHeight * 0.97) - (.Height)
      Let .Left = (xLeft + k) - .Width
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
      Let .Caption = "Generar"
      Let .Accelerator = "g"
      Let .ControlTipText = "Permite purcesar los datos y mostrar los resultados en una hoja Excel."
      Let .Top = .Parent.Controls("Btnmain0").Top
      Let .Left = .Parent.Controls("Btnmain0").Left - .Width - 2
      Let .Tag = "validate" 'ControlId & "-bal"
      Let .Default = True
    End With

    ' 4. [x] Assing handler for each button
    Set BtnCancel = .Controls("Btnmain0") 'Main
    Set BtnAccept = .Controls("Btnmain1") 'Main
    Set Btn0 = .Controls("Btnmain2")      'Main

    Set Txt0 = .Controls("Txtpur0") ' Frame0
    Set Txt1 = .Controls("Txtpur1") ' Frame0
    Set Txt2 = .Controls("Txtpur2") ' Frame0
    Set Txt3 = .Controls("Txtpur3") ' Frame1
    Set Txt4 = .Controls("Txtpur4") ' Frame1
    Set Txt5 = .Controls("Txtpur5") ' Frame2
    Set Txt6 = .Controls("Txtpur6") ' Frame2
    Set Txt7 = .Controls("Txtpur7") ' Frame2
    Set Txt8 = .Controls("Txtpur8") ' Frame2
    Set Txt9 = .Controls("Txtpur9") ' Frame3
    Set Txt10 = .Controls("Txtpur10") ' Frame3

    Set Chk0 = .Controls("Chkpur1") ' Frame1
    Set Chk1 = .Controls("Chkpur2") ' Frame2

    Set Cbo0 = .Controls("Cbopur1") ' Frame2
    Set Cbo1 = .Controls("Cbopur2") ' Frame2
    Set Cbo2 = .Controls("Cbopur3") ' Frame2
    Set Cbo3 = .Controls("Cbopur4") ' Frame2


    Set Btn1 = .Controls("Btnpur0") ' Frame2
    Set Btn2 = .Controls("Btnpur1") ' Frame2
    Set Btn3 = .Controls("Btnpur2") ' Frame2
    Set Btn4 = .Controls("Btnpur3") ' Frame2
    Set Btn5 = .Controls("Btnpur4") ' Frame2
    Set Btn6 = .Controls("Btnpur5") ' Frame2

    Set Lst0 = .Controls("Lstpur0") ' Frame2

    Set DPR = New ClassEqlVendorsDatePicker

    ' 5. [x] Give styles to controls
    Let .Controls("Cbopur1").ListIndex = 0
    Call VWP.ViewFormStyles(RES, FRM)
    Set VWP = Nothing

    ' 6. [x] Show the finished Form
    Call Load(FRM)
    Call .Show

    ' 7. [x] Close form
  End With

EH:
  Let VecLabels = Empty
  Let VecCntrls0 = Empty
  Set DPR = Nothing
  Set VWP = Nothing
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Sub

' Purchase concentrate UI
Private Sub View_form_purchaseconcentrate(ByVal ControlId As String, ByVal FRM As FRMBasics)

  Dim VWP As ClassEqlViewParent
  ' Controls
  Dim FmeCurrent As MSForms.Frame
  Dim VecLabels As Variant, VecCntrls0 As Variant
  ' Standard vars
  Dim StrAux0 As String, StrTitle As String, StrTitleForm As String, StrDivlocal As String, StrDivforeign As String
  Dim i As Long, j As Long, k As Long, l As Long, xLeft As Long

  Const sFN As String = "VIEW::-form_purchaseconcentrate"

  #If Debugging Then
    Call Debug_assert(Not ControlId = VBA.vbNullString, sFN)
    Call Debug_assert(Not FRM Is Nothing, sFN)
  #End If

  On Error GoTo EH
  ' 0. [x] Get data for Form Controls

  ' 1. [x] Desing the Form
  Let StrTitleForm = "Liquidación de Concentrados Minerales"
  Let StrTitle = StrTitleForm & " - " & STRAPPTITLE

  Set VWP = New ClassEqlViewParent
  Call VWP.ViewFormCleanner(RES, FRM)
  With FRM

    ' Set the form dimensions
    Let .Height = 415
    Let .Width = 450
    Let .Tag = ControlId
    Let .Caption = StrTitle
    Let xLeft = .InsideWidth * 0.05 ' Left for all main controls

    ' Build and design the form
    Call VWP.ViewFormAddControl(RES, FRM, "CommandButton", "Btnmain", 2) ' All with events (Cancel, Exe and Help)
    Call VWP.ViewFormAddControl(RES, FRM, "Frame", "Framecnc", 3)

    Set FmeCurrent = .Controls("Framecnc0") ' ---------------------------- Framecnc0'
    With FmeCurrent

      ' Set Frame properties
      Let .Caption = "Datos base y de provedor"
      Let .Width = (.Parent.InsideWidth * 0.9)
      Let .Height = .Parent.InsideHeight * 0.15
      Let .Left = xLeft
      Let .Top = .Parent.InsideHeight * 0.04

      ' Create Frame controls
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "Label", "Lblcnc", 4)
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "TextBox", "Txtcnc", 3) ' EVENT all
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "ComboBox", "Cbocnc", 0)
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "CheckBox", "Chkcnc", 0)

      ' Set Frame controls layout
      ' Labels - 6 cols (Lbl|Txt...)
      Let StrDivlocal = MDL.CurrencyMainSymbolLocal
      Let StrDivforeign = MDL.CurrencyMainSymbolForeign
      Let j = (.InsideWidth - (xLeft * 2) - 10) / 6
      For i = 0 To 10 ' Asign widths to each control
        Let StrAux0 = IIf(i <= 4, "Lbl", IIf(i <= 5, "Chk", IIf(i <= 9, "Txt", "Cbo"))) & "cnc" & IIf(i <= 4, i, IIf(i <= 5, 0, IIf(i <= 9, i - 6, 0)))
        With .Controls(StrAux0)
          If i < 6 Then Let .WordWrap = False: Let .Caption = IIf(i = 0, "Lugar:", IIf(i = 1, "Fecha:", IIf(i = 2, "TC [" & StrDivlocal & "/" & StrDivforeign & "]:", IIf(i = 3, "Provedor:", IIf(i = 4, "Tipo:", "Imprimir moneda local")))))
          Let .Width = j
        End With
      Next i
      Let VecLabels = Array("Lblcnc0", "Txtcnc0", "Lblcnc1", "Txtcnc1", "Lblcnc2", "Txtcnc2", "Lblcnc3", "Txtcnc3", "Lblcnc4", "Cbocnc0", "Chkcnc0")
      For i = 0 To 10 ' Asign positions
        With .Controls(VBA.CStr(VecLabels(i)))
          Let .Left = xLeft + ((.Width + 2.5) * IIf(i > 5, i - 6, i))
          Let .Top = (xLeft * 0.35) + IIf(i > 5, .Parent.Controls("Txtcnc0").Height, 0)
          If i Mod 2 = 0 Then Let .AutoSize = True
        End With
      Next i
      For i = 0 To 10 ' Adjust widths
        With .Controls(VBA.CStr(VecLabels(i)))
          If VBA.InStr(VBA.CStr(VecLabels(i)), "Txt") > 0 Or VBA.InStr(VBA.CStr(VecLabels(i)), "Cbo") > 0 Then
            Let StrAux0 = IIf(.Name = "Txtcnc0" Or .Name = "Txtcnc3", "Lblcnc3", IIf(.Name = "Txtcnc1" Or .Name = "Cbocnc0", "Lblcnc1", "Lblcnc2"))
            Let .Width = (.Width * 2) - .Parent.Controls(StrAux0).Width: Let .Left = .Parent.Controls(StrAux0).Left + .Parent.Controls(StrAux0).Width + 2.5
          End If
        End With
      Next i
      Let VecLabels = Empty

      ' Set Frame controls ToolTipTexts
      Let VecLabels = Array("Txtcnc0", "Txtcnc1", "Txtcnc2", "Chkcnc0", "Txtcnc3", "Cbocnc0", _
          "Orígen/localidad/ciudad donde se realiza este documento. Ej.: Potosí", _
          "Seleccione o introduzca una fecha, formato: dd/mm/aa", _
          "Tipo de cambio: una cifra monetaria equivalente al tipo de cambio actual", _
          "Si activa la casilla, se creará en el documento una columna con valores en moneda extranjera al tipo de cambio dado", _
          "Nombres y apellidos del provedor", _
          "Seleccione una opción de la lista")
      For i = 0 To 5
        Let .Controls(VBA.CStr(VecLabels(i))).ControlTipText = VBA.CStr(VecLabels(i + 6))
      Next i
      Let VecLabels = Empty

    End With
    Set FmeCurrent = Nothing

    Set FmeCurrent = .Controls("Framecnc1") ' ---------------------------- Framecnc1'
    With FmeCurrent

      ' Set Frame properties
      Let .Caption = "Datos concentrado y costo tratamiento"
      Let .Width = (.Parent.InsideWidth * 0.9)
      Let .Height = .Parent.InsideHeight * 0.2
      Let .Left = .Parent.Controls("Framecnc0").Left
      Let .Top = .Parent.Controls("Framecnc0").Top + .Parent.Controls("Framecnc0").Height + 5

      ' Create Frame controls
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "Label", "Lblcnc", 11, CtlFromX:=5)
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "TextBox", "Txtcnc", 10, CtlFromX:=4) ' EVENT all
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "SpinButton", "Spncnc", 3) ' EVENT all

      ' Set Frame controls layout: Labels - 6 cols (Lbl|Txt...)
      Let VecLabels = Array("Lblcnc5", "Txtcnc4", "Lblcnc6", "Txtcnc5", "Lblcnc7", "Txtcnc6", "Lblcnc8", "Txtcnc7", "Lblcnc9", "Txtcnc8", "Lblcnc10", "Txtcnc9", "Lblcnc11", "Txtcnc10")
      Let j = (.InsideWidth - (xLeft * 2) - 10) / 6
      For i = LBound(VecLabels) To UBound(VecLabels) ' Asign widths to each control
        With .Controls(VBA.CStr(VecLabels(i)))
          If i Mod 2 = 0 Then Let .WordWrap = False: Let .Caption = IIf(i = 0, "Fecha ingreso:", IIf(i = 2, "Peso [T]:", IIf(i = 4, "Maquila:", IIf(i = 6, "Humedad [%]:", IIf(i = 8, "Base:", IIf(i = 10, "Merma [%]:", "Escalador:"))))))
          Let .Width = j
        End With
      Next i
      For i = LBound(VecLabels) To UBound(VecLabels) ' Asign positions
        With .Controls(VBA.CStr(VecLabels(i)))
          Let .Left = xLeft + ((.Width + 2.5) * (i - IIf(i > 9, 8, IIf(i > 5, 4, 0))))
          Let .Top = (xLeft * 0.35) + (.Parent.Controls("Txtcnc4").Height * IIf(i > 9, 2, IIf(i > 5, 1, 0)))
          If i Mod 2 = 0 Then Let .AutoSize = True
        End With
      Next i
      For i = LBound(VecLabels) To UBound(VecLabels) ' Adjust widths
        If VBA.InStr(VBA.CStr(VecLabels(i)), "Txt") > 0 Then
          With .Controls(VBA.CStr(VecLabels(i)))
            Let StrAux0 = IIf(.Name = "Txtcnc4", "Lblcnc5", IIf(.Name = "Txtcnc5" Or .Name = "Txtcnc7" Or .Name = "Txtcnc9", "Lblcnc8", "Lblcnc11"))
            Let .Width = (.Width * 2) - (.Parent.Controls(StrAux0).Width + IIf(i = 7 Or i = 11, .Height, 0)): Let .Left = .Parent.Controls(StrAux0).Left + .Parent.Controls(StrAux0).Width + 2.5
          End With
        End If
      Next i
      Let VecLabels = Empty
      For i = 0 To 3 ' Spins adjustment
        With .Controls("Spncnc" & i)
          Let .Height = .Parent.Controls("Txtcnc7").Height
          Let .Width = .Parent.Controls("Txtcnc7").Height / 2
          Let k = IIf(i = 0 Or i = 1, 7, 9)
          Let .Left = .Parent.Controls("Txtcnc" & k).Left + .Parent.Controls("Txtcnc" & k).Width + IIf(i = 1 Or i = 2, .Width, 0)
          Let .Top = .Parent.Controls("Txtcnc" & k).Top
        End With
      Next i

      ' Set Frame controls ToolTipTexts
      Let VecLabels = Array("Txtcnc4", "Txtcnc5", "Txtcnc6", "Txtcnc7", "Txtcnc8", "Txtcnc9", "Txtcnc10", _
          "Seleccione o introduzca una fecha (de recepción del concentrado mineral), formato: dd/mm/aa", _
          "Seleccione o introduzca el peso bruto húmedo del concentrado mineral", _
          "Maquila: una cifra equivalente al costo de tratamiento del concentrado mineral", _
          "Humedad: una cifra en porcentaje (entre 0% - 100%)", _
          "Base: una cifra equivalente a la base de cotización internacional del concentrado mineral", _
          "Merma: una cifra en porcentaje (entre 0% - 100%)", _
          "Escalador: una cifra determinada en el contrato de compra-venta del concentrado mineral")
      For i = 0 To 6
        Let .Controls(VBA.CStr(VecLabels(i))).ControlTipText = VBA.CStr(VecLabels(i + 7))
      Next i
      Let VecLabels = Empty

    End With
    Set FmeCurrent = Nothing

    Set FmeCurrent = .Controls("Framecnc2") ' ---------------------------- Framecnc2'
    With FmeCurrent

      ' Set Frame properties
      Let .Caption = "Datos Liquidación"
      Let .Width = (.Parent.InsideWidth * 0.9)
      Let .Height = .Parent.InsideHeight * 0.35
      Let .Left = .Parent.Controls("Framecnc1").Left
      Let .Top = .Parent.Controls("Framecnc1").Top + .Parent.Controls("Framecnc1").Height + 5

      ' Create Frame controls
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "TextBox", "Txtcnc", 14, CtlFromX:=11) ' EVENT all
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "CheckBox", "Chkcnc", 2, CtlFromX:=1) ' EVENT 0
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "ComboBox", "Cbocnc", 3, CtlFromX:=1, CtlVisible:=False) ' EVENT all
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "CommandButton", "Btncnc", 5, CtlFromX:=0) ' EVENT all
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "ListBox", "Lstcnc", 0) ' EVENT all

      ' Set Frame controls layout
      ' Cbo & Btns - col 1 (unique col) - Row 1
      With .Controls("Cbocnc1")
        Let .Left = xLeft
        Let .Top = xLeft * 0.5
        Let .Width = (.Parent.InsideWidth - ((xLeft * 2) + (.Height * 6)))
        Let .Visible = True
      End With
      For i = 0 To 5
        With .Controls("Btncnc" & i)
          Let .Height = .Parent.Controls("Cbocnc1").Height
          Let .Width = .Parent.Controls("Cbocnc1").Height
          Let .Top = .Parent.Controls("Cbocnc1").Top
          Let .Left = (.Parent.Controls("Cbocnc1").Left + .Parent.Controls("Cbocnc1").Width) + (.Width * i)
          Let .Enabled = Not (i = 2 Or i = 3 Or i = 4)
        End With
      Next i

      ' Txts - col 1 (unique col) - Row 2
      Let j = (.InsideWidth - (xLeft * 2)) / 5
      Let VecLabels = Array("Cbocnc2", "Txtcnc11", "Txtcnc12", "Txtcnc13", "Txtcnc14", "Cbocnc3", "Chkcnc1")
      For i = LBound(VecLabels) To UBound(VecLabels)
        With .Controls(VBA.CStr(VecLabels(i)))
          Let .Width = j
          Let .Top = (.Parent.Controls("Cbocnc1").Top + .Parent.Controls("Cbocnc1").Height) + 1.5
          Let .Left = xLeft + (.Width * i)
          Let .Visible = (i <= 4)
        End With
      Next i
      Let VecLabels = Empty
      Let .Controls("Chkcnc1").WordWrap = False
      Let .Controls("Chkcnc1").AutoSize = True

      ' Lst - col 1 (unique col) - Row 3
      With .Controls("Lstcnc0")
        Let .Width = (.Parent.InsideWidth - (xLeft * 2))
        Let .Left = xLeft
        Let .Top = .Parent.Controls("Txtcnc11").Top + .Parent.Controls("Txtcnc11").Height
        Let .Height = (.Parent.InsideHeight - ((.Parent.Controls("Cbocnc1").Height * 3.5)))
      End With

      ' Chks - col 1 (unique col) - Row 4
      With .Controls("Chkcnc2")
        Let .WordWrap = False
        Let .Caption = "Autoconfigurar página"
        Let .AutoSize = True
        Let .Top = (.Parent.InsideHeight - xLeft * 0.8)
        Let .Left = xLeft
      End With

      ' Set Frame controls ToolTipTexts
      Let VecLabels = Array("Cbocnc1", "Btncnc0", "Btncnc1", "Btncnc2", "Btncnc3", "Btncnc4", "Btncnc5", "Lstcnc0", "Chkcnc2", _
          "Seleccione una opción de la lista", _
          "Añadir los valores de los controles a la lista", _
          "Confirmar edición", _
          "Cancelar edición", _
          "Eliminar fila de datos seleccionada", _
          "Eliminar todos los datos ingresados", _
          "Seleccionar datos desde un rango de celdas de " & Application.Name, _
          "Doble click en el ítem que desee modificar", _
          "El sistema autoconfigurará las páginas donde imprimirá la Liquidación Final")
      For i = 0 To 8
        Let .Controls(VBA.CStr(VecLabels(i))).ControlTipText = VBA.CStr(VecLabels(i + 9))
      Next i
      Let VecLabels = Empty

    End With
    Set FmeCurrent = Nothing

    Set FmeCurrent = .Controls("Framecnc3") ' ---------------------------- Framecnc3'
    With FmeCurrent

      ' Set Frame properties
      Let .Caption = "Dónde mostrar resultados"
      Let .Width = .Parent.Controls("Framecnc2").Width
      Let .Height = .Parent.InsideHeight * 0.1
      Let .Left = .Parent.Controls("Framecnc2").Left
      Let .Top = .Parent.Controls("Framecnc2").Top + .Parent.Controls("Framecnc2").Height + 5

      ' Create Frame controls
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "TextBox", "Txtcnc", 15, True, 15) ' EVENT all
      Call VWP.ViewFormAddControl(RES, FmeCurrent, "OptionButton", "Optcnc", 1)

      ' Set Frame controls layout
      ' Opts & Txt - Row 1
      Let j = (.InsideWidth - (xLeft * 2)) / 3
      Let .Controls("Optcnc0").WordWrap = False
      Let .Controls("Optcnc0").Width = j
      Let .Controls("Optcnc0").Left = xLeft
      Let .Controls("Optcnc0").Top = xLeft * 0.35
      Let .Controls("Optcnc1").WordWrap = False
      Let .Controls("Optcnc1").Width = j
      Let .Controls("Optcnc1").Left = xLeft + .Controls("Optcnc1").Width
      Let .Controls("Optcnc1").Top = .Controls("Optcnc0").Top
      Let .Controls("Txtcnc15").Left = .Controls("Optcnc1").Left + .Controls("Optcnc1").Width
      Let .Controls("Txtcnc15").Top = .Controls("Optcnc0").Top

      ' Set Frame controls ToolTipTexts
      Let .Controls("Optcnc0").ControlTipText = STRAPPTITLE & " creará una hoja de cálculo nueva en el libro actual"
      Let .Controls("Optcnc1").ControlTipText = "Active esta opción para generar el resultado en la hoja actual"
      Let .Controls("Txtcnc15").ControlTipText = "Seleccionar una celda"

    End With

    ' 2. [x] Fill controls with values
    ' Frame0
    Let .Controls("Txtcnc0").Text = "N/D"
    Let .Controls("Txtcnc1").Text = VBA.Format(Date, "dd/mm/yy")
    Let .Controls("Txtcnc1").TextAlign = fmTextAlignRight
    Let .Controls("Txtcnc4").Text = VBA.Format(Date, "dd/mm/yy")
    Let .Controls("Txtcnc4").TextAlign = fmTextAlignRight
    Let .Controls("Txtcnc2").Text = "1"
    ' Frame1
    Let .Controls("Cbocnc0").List = Array("Cooperativista", "Particular")
    Let .Controls("Cbocnc0").ListIndex = 0
    ' Frame2
    Let .Controls("Cbocnc1").List = Array("Contenidos y condiciones", "Refinación", "Penalidades", "Deducciones", "Bonos")
    Let .Controls("Cbocnc1").ListStyle = fmListStyleOption
    Let .Controls("Cbocnc3").List = Array("Y", "O")
    Let .Controls("Cbocnc3").ListIndex = 0
    Let VecLabels = Array(&H271A, &H2710, &H270F, &H2A31, &H2716, &H2583)
    For i = 0 To 5
      Let .Controls("Btncnc" & i).Caption = VBA.ChrW(VecLabels(i))
      If Not (i = 2 Or i = 5) Then Let .Controls("Btncnc" & i).Tag = "validate"
    Next i
    Let .Controls("Lstcnc0").ColumnHeads = True
    Let .Controls("Lstcnc0").IntegralHeight = False
    ' Frame3
    Let .Controls("Optcnc0").Value = True
    Let .Controls("Optcnc0").Caption = "Hoja nueva"
    Let .Controls("Optcnc1").Caption = "Hoja actual"
    Let .Controls("Optcnc1").AutoSize = True
    Let .Controls("Txtcnc15").Width = j + (j - .Controls("Optcnc1").Width)
    Let .Controls("Txtcnc15").Left = .Controls("Optcnc1").Left + .Controls("Optcnc1").Width

    ' Set Tabindex
    Let VecLabels = Array("Lblcnc0", "Txtcnc0", "Lblcnc1", "Txtcnc1", "Lblcnc2", "Txtcnc2", "Lblcnc3", "Txtcnc3", "Lblcnc4", "Cbocnc0", "Chkcnc0", "Lblcnc5", "Txtcnc4", "Lblcnc6", "Txtcnc5", "Lblcnc7", "Txtcnc6", "Lblcnc8", "Txtcnc7", "Spncnc0", "Spncnc1", "Lblcnc9", "Txtcnc8", "Lblcnc10", "Txtcnc9", "Spncnc3", "Spncnc2", "Lblcnc11", "Txtcnc10", "Cbocnc1", "Btncnc0", "Btncnc1", "Btncnc2", "Btncnc3", "Btncnc4", "Btncnc5", "Cbocnc2", "Txtcnc11", "Txtcnc12", "Cbocnc3", "Txtcnc13", "Txtcnc14", "Lstcnc0", "Chkcnc1", "Chkcnc2", "Optcnc0", "Optcnc1", "Txtcnc15", "Btnmain0", "Btnmain1", "Btnmain2")
    For i = LBound(VecLabels) To UBound(VecLabels)
      Let .Controls(VBA.CStr(VecLabels(i))).TabIndex = i
    Next i
    Let VecLabels = Empty

    ' Set Hotkeys
    Let VecLabels = Array("Lblcnc0", "Lblcnc1", "Lblcnc2", "Lblcnc3", "Lblcnc4", "Chkcnc0", "Lblcnc5", "Lblcnc6", "Lblcnc7", "Lblcnc8", "Lblcnc9", "Lblcnc10", "Lblcnc11", "Chkcnc2", "Optcnc0", "Optcnc1", "l", "f", "t", "r", "o", "n", "s", "p", "q", "h", "b", "m", "e", "i", "v", "j")
    For i = 0 To 15
      Let .Controls(VBA.CStr(VecLabels(i))).Accelerator = VBA.CStr(VecLabels(i + 16))
    Next i
    Let VecLabels = Empty

    ' Set Form textbox controls dropbuttons
    For i = 0 To 15
      With .Controls("Txtcnc" & i)
        Let .DropButtonStyle = IIf(i = 1 Or i = 4, fmDropButtonStylePlain, fmDropButtonStyleReduce)
        Let .ShowDropButtonWhen = fmShowDropButtonWhenAlways
      End With
    Next i

    ' 3. [x] Fill Form Controls with data and Set general cncpose CommandButtons
    Let k = (.InsideWidth - (xLeft * 2))
    With .Controls("Btnmain0") '("BtnCancel")
      Let .Accelerator = "c"
      Let .Cancel = True
      Let .Top = (.Parent.InsideHeight * 0.97) - (.Height)
      Let .Left = (xLeft + k) - .Width
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
      Let .Caption = "Generar"
      Let .Accelerator = "g"
      Let .ControlTipText = "Permite cnccesar los datos y mostrar los resultados en una hoja Excel."
      Let .Top = .Parent.Controls("Btnmain0").Top
      Let .Left = .Parent.Controls("Btnmain0").Left - .Width - 2
      Let .Tag = "validate" 'ControlId & "-bal"
      Let .Default = True
    End With

    ' 4. [x] Assing handler for each button
    Set BtnCancel = .Controls("Btnmain0") 'Main
    Set BtnAccept = .Controls("Btnmain1") 'Main
    Set Btn0 = .Controls("Btnmain2")      'Main

    Set Txt0 = .Controls("Txtcnc0") ' Frame0
    Set Txt1 = .Controls("Txtcnc1") ' Frame0
    Set Txt2 = .Controls("Txtcnc2") ' Frame0
    Set Txt3 = .Controls("Txtcnc3") ' Frame0
    Set Txt4 = .Controls("Txtcnc4") ' Frame1
    Set Txt5 = .Controls("Txtcnc5") ' Frame1
    Set Txt6 = .Controls("Txtcnc6") ' Frame1
    Set Txt7 = .Controls("Txtcnc7") ' Frame1
    Set Txt8 = .Controls("Txtcnc8") ' Frame1
    Set Txt9 = .Controls("Txtcnc9") ' Frame1
    Set Txt10 = .Controls("Txtcnc10") ' Frame1
    Set Txt11 = .Controls("Txtcnc11") ' Frame2
    Set Txt12 = .Controls("Txtcnc12") ' Frame2
    Set Txt13 = .Controls("Txtcnc13") ' Frame2
    Set Txt14 = .Controls("Txtcnc14") ' Frame2
    Set Txt15 = .Controls("Txtcnc15") ' Frame3

    Set Spn0 = .Controls("Spncnc0") ' Frame1
    Set Spn1 = .Controls("Spncnc1") ' Frame1
    Set Spn2 = .Controls("Spncnc2") ' Frame1
    Set Spn3 = .Controls("Spncnc3") ' Frame1

    Set Chk0 = .Controls("Chkcnc1") ' Frame2

    Set Cbo0 = .Controls("Cbocnc1") ' Frame2

    Set Btn1 = .Controls("Btncnc0") ' Frame2
    Set Btn2 = .Controls("Btncnc1") ' Frame2
    Set Btn3 = .Controls("Btncnc2") ' Frame2
    Set Btn4 = .Controls("Btncnc3") ' Frame2
    Set Btn5 = .Controls("Btncnc4") ' Frame2
    Set Btn6 = .Controls("Btncnc5") ' Frame2

    Set Lst0 = .Controls("Lstcnc0") ' Frame2

    Set DPR = New ClassEqlVendorsDatePicker

    ' 5. [x] Give styles to controls
    Let .Controls("Cbocnc1").ListIndex = 0
    Call VWP.ViewFormStyles(RES, FRM)
    Set VWP = Nothing

    ' 6. [x] Show the finished Form
    Call Load(FRM)
    Call .Show

    ' 7. [x] Close form
  End With

EH:
  Let VecLabels = Empty
  Let VecCntrls0 = Empty
  Set DPR = Nothing
  Set VWP = Nothing
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Sub

' Config UI
Private Sub View_form_config(ByVal ControlId As String, ByVal FRM As FRMBasics)

  ' Main objects
  Dim VIEWS As ClassEqlViewParent
  Dim MpgMulti As MSForms.MultiPage
  ' Controls
  Dim VecAux0 As Variant
  ' Standard vars
  Dim sTitle As String, TitleForm As String, StrAux0 As String
  Dim i As Long, j As Long, k As Long
  Dim xLeft As Long

  Const sFN As String = "VIEW::-form_config"

  #If Debugging Then
    Call Debug_assert(Not ControlId = VBA.vbNullString, sFN)
    Call Debug_assert(Not FRM Is Nothing, sFN)
  #End If

  On Error GoTo EH

  ' Set Form basic datas
  Let TitleForm = "Configuración"
  Let sTitle = TitleForm & " - " & STRAPPTITLE

  Set VIEWS = New ClassEqlViewParent
  Call VIEWS.ViewFormCleanner(RES, FRM)
  With FRM

    ' Set the form dimensions
    Let .Height = 415
    Let .Width = 450
    Let .Tag = ControlId
    Let .Caption = sTitle
    Let xLeft = .InsideWidth * 0.05

    ' Build and design the form
    Call VIEWS.ViewFormAddControl(RES, FRM, "CommandButton", "Btnmain", 0) ' All with events (Cancel)
    Call VIEWS.ViewFormAddControl(RES, FRM, "MultiPage", "Multicon", 0, False)

    ' Set multipage config
    Set MpgMulti = .Controls("Multicon")
    With MpgMulti
      Call .Pages.Add
      Call .Pages.Add
      Call .Pages.Add
      Let VecAux0 = Array("Divisas", "Selecciones", "Página", "Liquidaciones", "Elementos/Especies minerales")
      For i = 0 To .Pages.Count - 1
        Let .Pages(i).Caption = VecAux0(i)
        Let .Pages(i).Accelerator = VBA.Left(VBA.LCase(VBA.CStr(VecAux0(i))), 1)
        'Debug.Print .Pages(i).Name
      Next i
      Let VecAux0 = Empty
      Let .Width = .Parent.InsideWidth
      Let .Height = (.Parent.InsideHeight * 0.85)
    End With

    ' ----------------------------------------------------------- Page 0 - Divisa
    ' Add controls
    With MpgMulti
      Call VIEWS.ViewFormAddControl(RES, .Pages(0), "Label", "Lblcon", 4)
      Call VIEWS.ViewFormAddControl(RES, .Pages(0), "TextBox", "Txtcon", 2)
      Call VIEWS.ViewFormAddControl(RES, .Pages(0), "ListBox", "Lstcon", 0)
      Call VIEWS.ViewFormAddControl(RES, .Pages(0), "CommandButton", "Btncon", 5) ' EVENT all
      Call VIEWS.ViewFormAddControl(RES, .Pages(0), "ComboBox", "Cbocon", 1)
    End With

    Let i = ((MpgMulti.Width - (xLeft * 2)) / 3) - 2
    ' Rows 1, 2 (labels, textbox and buttons)
    Let VecAux0 = Array("Lblcon0", "Lblcon1", "Lblcon2", "Txtcon0", "Txtcon1", "Txtcon2", "Divisa en singular:", "Divisa en plural:", "Símbolo:")
    For j = 0 To 5
      With .Controls(VBA.CStr(VecAux0(j)))
        If j < 3 Then
          Let .WordWrap = False
          Let .Caption = VBA.CStr(VecAux0(j + 6))
        Else
          Let .MaxLength = IIf(j = 5, 3, 30)
        End If
        Let .Width = i
        Let .Left = xLeft + (.Width * (j + IIf(j > 2, -3, 0)))
        Let .Top = (xLeft * 0.5) + (FRM.Controls("Txtcon0").Height * IIf(j < 3, 1, 2))
      End With
    Next j
    Let VecAux0 = Empty

    Let VecAux0 = Array(&H271A, &H2710, &H270F, &H2A31)
    Let k = (i * 3) - (.Controls("Txtcon0").Height * 3) + 4
    For j = 0 To 3
      With .Controls("Btncon" & j)
        Let .Width = .Parent.Controls("Txtcon0").Height
        Let .Height = .Width
        Let .Left = k + (.Width * j)
        Let .Top = .Parent.Controls("Lblcon0").Top
        Let .Caption = VBA.ChrW(VecAux0(j))
        Let .Enabled = (j = 0)
      End With
    Next j
    Let VecAux0 = Empty

    ' Row 3 (List)
    With .Controls("Lstcon0")
      Let .Top = (FRM.Controls("Txtcon0").Top + FRM.Controls("Txtcon0").Height)
      Let .Left = FRM.Controls("Txtcon0").Left
      Let .Height = FRM.Controls("Multicon").Height * 0.5
      Let .Width = i * 3
      Let .IntegralHeight = False
      Let .ColumnHeads = True
      Let .ColumnCount = 3
    End With

    ' Rows 4, 5 (Labels, combos and buttons)
    Let VecAux0 = Array("Lblcon3", "Lblcon4", "Cbocon0", "Cbocon1", "Btncon4", "Btncon5")
    For j = 0 To 5
      With .Controls(VBA.CStr(VecAux0(j)))
        If j < 2 Or j > 3 Then
          If j < 2 Then Let .WordWrap = False: Let .Caption = IIf(j = 0, "Moneda local:", "Moneda Extranjera:")
          If j > 3 Then Let .Caption = "Establecer"
        End If
        Let .Width = i
        Let .Height = IIf(j > 3, .Parent.Controls("Txtcon0").Height, .Height)
        Let .Left = xLeft + (.Width * (j + IIf(j Mod 2 = 0, IIf(j = 0, 0, (-j + IIf(j < 4, 1, 2))), IIf(j = 1, -j, (-j + IIf(j < 4, 1, 2))))))
        Let .Top = (FRM.Controls("Lstcon0").Top + FRM.Controls("Lstcon0").Height) + (.Parent.Controls("Txtcon0").Height * IIf(j Mod 2 = 1, 2, 1))
        If j < 2 Then Let .AutoSize = True ': Let .BackColor = VBA.vbRed
        If j = 2 Or j = 3 Then Let .Width = i + (i - (.Parent.Controls("Lblcon4").Width + 2)): Let .Left = xLeft + .Parent.Controls("Lblcon4").Width + 2
      End With
    Next j
    Let VecAux0 = Empty

    ' Set tool tip texts
    Let VecAux0 = Array("Txtcon0", "Txtcon1", "Txtcon2", "Btncon0", "Btncon1", "Btncon2", "Btncon3", "Lstcon0", "Cbocon0", "Cbocon1", "Btncon4", "Btncon5", "Nombre de la divisa en forma singular. Ej. Boliviano", "Nombre de la divisa en forma plural. Ej. Bolivianos", "Símbolo de la divisa. Ej. BOB", "Crear moneda", "Editar moneda", "Cancelar edición", "Eliminar moneda seleccionada", "Doble click o pulse barra espaciadora para editar el ítem seleccionado", "Seleccione una opción de la lista", "Seleccione una opción de la lista", "Establece la moneda seleccionada como moneda local", "Establece la moneda seleccionada como moneda extranjera")
    For j = 0 To 11
      Let .Controls(VBA.CStr(VecAux0(j))).ControlTipText = VBA.CStr(VecAux0(j + 12))
    Next j
    Let VecAux0 = Empty

    ' Set tabindex
    Let VecAux0 = Array("Lblcon0", "Txtcon0", "Lblcon1", "Txtcon1", "Lblcon2", "Txtcon2", "Btncon0", "Btncon1", "Btncon2", "Btncon3", "Lstcon0", "Lblcon3", "Cbocon0", "Btncon4", "Lblcon4", "Cbocon1", "Btncon5")
    For j = 0 To UBound(VecAux0)
      Let .Controls(VBA.CStr(VecAux0(j))).TabIndex = j
    Next j
    Let VecAux0 = Empty

    ' Hot keys
    Let VecAux0 = Array("Lblcon0", "Lblcon1", "Lblcon2", "Lblcon3", "Btncon4", "Lblcon4", "Btncon5", "i", "u", "o", "a", "t", "j", "r")
    For j = 0 To 6
      Let .Controls(VBA.CStr(VecAux0(j))).Accelerator = VBA.CStr(VecAux0(j + 7))
    Next j
    Let VecAux0 = Empty


    ' ----------------------------------------------------------- Page 1 - Selecciones
    ' Add controls to Page 1 - Selecciones
    Call VIEWS.ViewFormAddControl(RES, MpgMulti.Pages(1), "CheckBox", "Chkcon", 1)

    Let VecAux0 = Array("Resultados compactos en Balances metalúrgicos y Regresión lineal Selección", "Mostrar Resultados de Cabeza calculada en Balances metalúrgicos", "Permite generar un resultado básico en los Balances Metalúrgicos", "Genera resultados de leyes de cabeza ensayada y calculada en Balances Metalúrgicos", "r", "z")
    For j = 0 To 1
      With .Controls("Chkcon" & j)
        Let .Caption = VBA.CStr(VecAux0(j))
        Let .ControlTipText = VBA.CStr(VecAux0(j + 2))
        Let .Accelerator = VBA.CStr(VecAux0(j + 4))
        Let .TabIndex = j
        Let .WordWrap = False
        Let .AutoSize = True
        Let .Left = xLeft
        Let .Top = xLeft + (FRM.Controls("Txtcon0").Height * j)
      End With
    Next j
    Let VecAux0 = Empty


    ' ----------------------------------------------------------- Page 2 - Página
    ' Add controls
    With MpgMulti
      Call VIEWS.ViewFormAddControl(RES, .Pages(2), "Label", "Lblcon", 11, CtlFromX:=5)
      Call VIEWS.ViewFormAddControl(RES, .Pages(2), "TextBox", "Txtcon", 7, CtlFromX:=3) ' EVENT 3
      Call VIEWS.ViewFormAddControl(RES, .Pages(2), "CheckBox", "Chkcon", 6, CtlFromX:=2) ' EVENT all
      Call VIEWS.ViewFormAddControl(RES, .Pages(2), "CommandButton", "Btncon", 9, CtlFromX:=6) ' EVENT all
      Call VIEWS.ViewFormAddControl(RES, .Pages(2), "ComboBox", "Cbocon", 3, CtlFromX:=2)
    End With

    ' Row 1
    For j = 0 To 1
      With .Controls(IIf(j = 0, "Lblcon5", "Cbocon2"))
        If j = 0 Then Let .Caption = "Encabezado y pie de página:"
        Let .Width = i
        Let .Top = xLeft
        Let .Left = xLeft + (.Width * j)
        If j = 1 Then Let .List = Array("No establecer", "Encabezado de página", "Pie de página", "Ambos, encabezado y pie")
      End With
    Next j

    ' Row 2 - Lbl, Txt, Chk
    Let VecAux0 = Array("Logo:", "Empresa:", "Dirección:", "Teléfono:", "Correo electrónico:")
    For j = 0 To 14
      Let k = IIf(j = 0 Or j = 5 Or j = 10, 0, k)
      Let StrAux0 = IIf(j < 5, "Lbl", IIf(j < 10, "Txt", "Chk")) & "con" & k + IIf(j < 5, 6, IIf(j < 10, 3, 2))
      With .Controls(StrAux0)
        Let .Width = i
        If j < 5 Then Let .WordWrap = False: Let .Caption = VBA.CStr(VecAux0(j)): Let .BackColor = VBA.vbRed
        If j > 9 Then Let .WordWrap = False: Let .TripleState = True
        Let .Left = xLeft + (.Width * IIf(j < 5, 0, IIf(j < 10, 1, 2)))
        Let .Top = (xLeft * 2) + (FRM.Controls("Txtcon3").Height * k)
      End With
      Let k = k + 1
    Next j
    Let VecAux0 = Empty
    With .Controls("Btncon6")
      Let .Caption = "Guardar cambios"
      Let .Width = i
      Let .Top = (FRM.Controls("Txtcon6").Top + (FRM.Controls("Txtcon6").Height * 2.5))
      Let .Left = FRM.Controls("Chkcon6").Left
    End With

    ' Row 3
    For j = 0 To 1
      With .Controls(IIf(j = 0, "Lblcon11", "Cbocon3"))
        Let .Width = i
        If j = 0 Then Let .Caption = "Destinos de documento:": Let .WordWrap = False
        Let .Left = xLeft + IIf(j = 0, 0, .Width)
        Let .Top = (FRM.Controls("Btncon6").Top + (FRM.Controls("Btncon6").Height * 2))
      End With
    Next j
    Let VecAux0 = Array(&H271A, &H2710, &H2A31)
    For j = 7 To 9
      With .Controls("Btncon" & j)
        Let .Caption = VBA.ChrW(VecAux0(j - 7))
        Let .Height = FRM.Controls("Cbocon3").Height
        Let .Width = .Height
        Let .Left = (FRM.Controls("Cbocon3").Left + FRM.Controls("Cbocon3").Width) + (.Width * (j - 7))
        Let .Top = FRM.Controls("Cbocon3").Top
      End With
    Next j
    Let VecAux0 = Empty

    ' Set tool tip texts
    Let VecAux0 = Array("Cbocon2", "Txtcon3", "Chkcon2", "Txtcon4", "Chkcon3", "Txtcon5", "Chkcon4", "Txtcon6", "Chkcon5", "Txtcon7", "Chkcon6", "Btncon6", "Cbocon3", "Btncon7", "Btncon8", "Btncon9", _
     "Seleccione una opción de la lista para el encabezado y pie de pagina", "Seleccione el logotipo de su empresa", "Mostrará el logotipo en la sección de la hoja que establezca", "Escriba el nombre de su empresa", _
     "Mostrará el nombre de su empresa en la sección de la hoja que establezca", "Escriba la dirección de su empresa", "Mostrará la dirección de su empresa en la sección de la hoja que establezca", _
     "Escriba el número telefónico de su empresa", "Mostrará número telefónico de su empresa en la sección de la hoja que establezca", "Escriba el correo electrónico de su empresa", _
     "Mostrará correo electrónico de su empresa en la sección de la hoja que establezca", "Guarda los cambios realizados en la sección superior en la base de datos", "Escriba un nuevo dato o seleccione un ítem de la lista para editar o eliminar", "Crear nuevo destino de documento", "Editar destino de documento seleccionado", "Eliminar destino de documento seleccionado")
    For j = 0 To 15
      Let .Controls(VBA.CStr(VecAux0(j))).ControlTipText = VBA.CStr(VecAux0(j + 16))
    Next j
    Let VecAux0 = Empty

    ' Set tabindex
    Let VecAux0 = Array("Lblcon5", "Cbocon2", "Lblcon6", "Txtcon3", "Chkcon2", "Lblcon7", "Txtcon4", "Chkcon3", "Lblcon8", "Txtcon5", "Chkcon4", "Lblcon9", "Txtcon6", "Chkcon5", "Lblcon10", "Txtcon7", "Chkcon6", "Btncon6", "Lblcon11", "Cbocon3", "Btncon7", "Btncon8", "Btncon9")
    For j = 0 To UBound(VecAux0)
      Let .Controls(VBA.CStr(VecAux0(j))).TabIndex = j
    Next j
    Let VecAux0 = Empty

    ' Hot keys
    Let VecAux0 = Array("Lblcon5", "Lblcon6", "Lblcon7", "Lblcon8", "Lblcon9", "Lblcon10", "Btncon6", "Lblcon11", "n", "o", "m", "r", "t", "i", "g", "u")
    For j = 0 To 7
      Let .Controls(VBA.CStr(VecAux0(j))).Accelerator = VBA.CStr(VecAux0(j + 8))
    Next j
    Let VecAux0 = Empty


    ' ----------------------------------------------------------- Page 3 - Liquidaciones
    ' Add controls
    With MpgMulti
      Call VIEWS.ViewFormAddControl(RES, .Pages(3), "Label", "Lblcon", 12, CtlFromX:=12)
      Call VIEWS.ViewFormAddControl(RES, .Pages(3), "TextBox", "Txtcon", 12, CtlFromX:=8)
      Call VIEWS.ViewFormAddControl(RES, .Pages(3), "CheckBox", "Chkcon", 14, CtlFromX:=7)
      Call VIEWS.ViewFormAddControl(RES, .Pages(3), "ListBox", "Lstcon", 1, CtlFromX:=1)
      Call VIEWS.ViewFormAddControl(RES, .Pages(3), "CommandButton", "Btncon", 13, CtlFromX:=10) 'EVENT all
      Call VIEWS.ViewFormAddControl(RES, .Pages(3), "ComboBox", "Cbocon", 7, CtlFromX:=4)
    End With

    ' Rows 1, 2, 3
    Let VecAux0 = Array("Lblcon12", "Cbocon4", "Btncon10", "Btncon11", "Btncon12", "Btncon13", "Cbocon5", "Cbocon6", "Txtcon8", "Txtcon9", "Txtcon10", "Cbocon7", "Txtcon11", "Txtcon12", "Chkcon7", &H271A, &H2710, &H270F, &H2A31)
    For j = 0 To 14
      With .Controls(VBA.CStr(VecAux0(j)))
        Let .Width = IIf(j = 0, .Width, IIf(j = 1, ((i * 3) - (.Height * 4)), IIf(j > 5, (((i * 3) / 7) * IIf(j = 6, 2, 1)), 0)))
        If j > 1 And j < 6 Then Let .Width = .Parent.Controls("Cbocon4").Height: Let .Height = .Width: Let .Caption = VBA.ChrW(VecAux0(j + 13))
        If j = 0 Then Let .Caption = "Seleccione una opción de la lista:": Let .WordWrap = False: Let .AutoSize = True
        If j = 1 Then Let .ListStyle = fmListStyleOption: Let .List = Array("Configurar regalías mineras", "Configurar retenciones de ley", "Configurar retenciones otras (institucionales)", "Configurar penalidades para compra-venta de concentrados minerales", "Configurar deducciones en exportaciones")
        Let .Left = xLeft + (.Parent.Controls("Cbocon4").Width * IIf(j > 1 And j < 6, 1, 0)) + IIf(j > 1 And j < 6, (.Width * (j - 2)), 0)
        Let .Top = (xLeft * 0.7) + (FRM.Controls("Cbocon4").Height * IIf(j > 5, 2, IIf(j = 0, 0, 1)))
        If j = 14 Then Let .Visible = False: Let .Caption = "Valor directo"
      End With
    Next j
    Let VecAux0 = Empty

    ' Row 4
    With .Controls("Lstcon1")
      Let .Height = FRM.Controls("Multicon").Height * 0.35
      Let .Width = i * 3
      Let .Left = xLeft
      Let .Top = FRM.Controls("Txtcon8").Top + FRM.Controls("Txtcon8").Height
      Let .IntegralHeight = False
      Let .ColumnHeads = True
    End With

    ' Next several rows (Chks)
    Let VecAux0 = Array("Forzar promedio de leyes minerales en liquidación de minerales", "Resumir regalías en liquidaciones", "Resumir retenciones de ley en liquidaciones", "Resumir retenciones institucionales en liquidaciones", "Resumir descuentos en liquidaciones", "Resumir bonos en liquidaciones", "Resumir deducción por exportación en liquidación de concentrados")
    For j = 0 To 6
      With .Controls("Chkcon" & j + 8)
        Let .WordWrap = False
        Let .Caption = VBA.CStr(VecAux0(j))
        Let .Width = i * 3
        Let .Left = xLeft
        Let .Top = (FRM.Controls("Lstcon1").Top + FRM.Controls("Lstcon1").Height + (xLeft * 0.5)) + ((.Height * 0.8) * j)
      End With
    Next j
    Let VecAux0 = Empty

    ' Set tool tip texts
    Let VecAux0 = Array("Cbocon4", "Btncon10", "Btncon11", "Btncon12", "Btncon13", "Lstcon1", "Chkcon8", "Chkcon9", "Chkcon10", "Chkcon11", "Chkcon12", "Chkcon13", "Chkcon14", _
     "Seleccione una opción de la lista", "Crear nuevo registro", "Editar datos de registro seleccionado", "Cancelar edición", "Eliminar registro seleccionado", _
     "Doble click o barra espaciadora en el ítem que desee modificar", "Promedia las leyes (por defecto) o los precios en múltiples liquidaciones de minerales para un provedor", "Obtiene una sola fila con el resultado de la retención regalía minera en las liquidaciones", "Obtiene una sola fila con el resultado de las retenciones de ley en las liquidaciones", _
     "Obtiene una sola fila con el resultado de las retenciones institucionales en las liquidaciones", "Obtiene una sola fila con el resultado de los descuentos en las liquidaciones", _
     "Obtiene una sola fila con el resultado de los bonos en las liquidaciones", "Obtiene una sola fila con el resultado de los bonos en las liquidaciones")
    For j = 0 To 12
      Let .Controls(VBA.CStr(VecAux0(j))).ControlTipText = VBA.CStr(VecAux0(j + 13))
    Next j
    Let VecAux0 = Empty

    ' Set tabindex
    Let VecAux0 = Array("Lblcon12", "Cbocon4", "Btncon10", "Btncon11", "Btncon12", "Btncon13", "Btncon13", "Cbocon5", "Cbocon6", "Txtcon8", "Txtcon9", "Txtcon10", "Txtcon11", "Txtcon12", "Cbocon7", "Chkcon7", "Lstcon1", "Chkcon8", "Chkcon9", "Chkcon10", "Chkcon11", "Chkcon12", "Chkcon13", "Chkcon14")
    For j = 0 To UBound(VecAux0)
      Let .Controls(VBA.CStr(VecAux0(j))).TabIndex = j
    Next j
    Let VecAux0 = Empty

    ' Hot keys
    Let VecAux0 = Array("Lblcon12", "Chkcon7", "Chkcon8", "Chkcon9", "Chkcon10", "Chkcon11", "Chkcon12", "Chkcon13", "Chkcon14", "o", "v", "f", "r", "y", "q", "n", "x", "i")
    For j = 0 To 8
      Let .Controls(VBA.CStr(VecAux0(j))).Accelerator = VBA.CStr(VecAux0(j + 9))
    Next j
    Let VecAux0 = Empty


    ' ----------------------------------------------------------- Page 4 - Elementos
    ' Add controls
    With MpgMulti
      Call VIEWS.ViewFormAddControl(RES, .Pages(4), "Label", "Lblcon", 15, CtlFromX:=13)
      Call VIEWS.ViewFormAddControl(RES, .Pages(4), "TextBox", "Txtcon", 14, CtlFromX:=13)
      Call VIEWS.ViewFormAddControl(RES, .Pages(4), "ListBox", "Lstcon", 2, CtlFromX:=2) ' EVENT all
      Call VIEWS.ViewFormAddControl(RES, .Pages(4), "CommandButton", "Btncon", 17, CtlFromX:=14) ' EVENT all
      Call VIEWS.ViewFormAddControl(RES, .Pages(4), "CheckBox", "Chkcon", 15, CtlFromX:=15)
    End With

    ' Rows 1, 2 (labels, textbox and buttons)
    Let VecAux0 = Array("Lblcon13", "Lblcon14", "Lblcon15", "Txtcon13", "Txtcon14", "Chkcon15", "Elemento:", "Símbolo:", VBA.vbNullString)
    For j = 0 To 5
      With .Controls(VBA.CStr(VecAux0(j)))
        If j < 3 Then
          Let .WordWrap = False
          Let .Caption = VBA.CStr(VecAux0(j + 6))
        Else
          If Not j = 5 Then Let .MaxLength = IIf(j = 4, 3, 15)
          If j = 5 Then Let .Caption = "Mineral tipo complejo"
        End If
        Let .Width = i
        Let .Left = xLeft + (.Width * (j + IIf(j > 2, -3, 0)))
        Let .Top = xLeft + (FRM.Controls("Txtcon13").Height * IIf(j < 3, 1, 2))
      End With
    Next j
    Let VecAux0 = Empty

    Let VecAux0 = Array(&H271A, &H2710, &H270F, &H2A31)
    Let k = (i * 3) - (.Controls("Txtcon13").Height * 3) + 4
    For j = 0 To 3
      With .Controls("Btncon" & j + 14)
        Let .Width = .Parent.Controls("Txtcon13").Height
        Let .Height = .Width
        Let .Left = k + (.Width * j)
        Let .Top = .Parent.Controls("Lblcon13").Top
        Let .Caption = VBA.ChrW(VecAux0(j))
        Let .Enabled = (j = 0)
      End With
    Next j
    Let VecAux0 = Empty

    ' Row 3 (List)
    With .Controls("Lstcon2")
      Let .Top = (FRM.Controls("Txtcon13").Top + FRM.Controls("Txtcon13").Height)
      Let .Left = xLeft
      Let .Height = FRM.Controls("Multicon").Height * 0.5
      Let .Width = i * 3
      Let .IntegralHeight = False
      Let .ColumnHeads = True
      Let .ColumnCount = 3
    End With

    ' Set tool tip texts
    Let VecAux0 = Array("Btncon14", "Btncon15", "Btncon16", "Btncon17", "Txtcon13", "Txtcon14", "Chkcon15", "Lstcon2", "Crear nuevo registro", "Editar datos de registro seleccionado", "Cancelar edición", "Eliminar registro seleccionado", "Nombre del elemento o especie mineral", "Símbolo del elemento o especie mineral", "Active la casilla si el elemento o especie mineral es de tipo complejo", "Doble click o barra espaciadora en el ítem que desee modificar")
    For j = 0 To 7
      Let .Controls(VBA.CStr(VecAux0(j))).ControlTipText = VBA.CStr(VecAux0(j + 8))
    Next j
    Let VecAux0 = Empty

    ' Set tabindex
    Let VecAux0 = Array("Lblcon13", "Txtcon13", "Lblcon14", "Txtcon14", "Chkcon15", "Btncon14", "Btncon15", "Btncon16", "Btncon17", "Lstcon2")
    For j = 0 To UBound(VecAux0)
      Let .Controls(VBA.CStr(VecAux0(j))).TabIndex = j
    Next j
    Let VecAux0 = Empty

    ' Hot keys
    Let VecAux0 = Array("Lblcon13", "Lblcon14", "Chkcon15", "m", "b", "t")
    For j = 0 To 2
      Let .Controls(VBA.CStr(VecAux0(j))).Accelerator = VBA.CStr(VecAux0(j + 3))
    Next j
    Let VecAux0 = Empty



    ' 2. [ ] Fill controls with values
    ' Page 0
    With .Controls("Lstcon0")
      Let .RowSource = MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigcurrencies_sys, eqlMdlStrTableName)
      Let .ColumnCount = 3
      Let FRM.Controls("Btncon3").Enabled = (.ListCount > 2)
    End With
    Let j = -1: Let k = -1
    If .Controls("Lstcon0").ListCount > 0 Then
      For i = 0 To .Controls("Lstcon0").ListCount - 1
        Let StrAux0 = .Controls("Lstcon0").List(i, 0) & " [" & .Controls("Lstcon0").List(i, 2) & "]"
        Call .Controls("Cbocon0").AddItem(StrAux0)
        Call .Controls("Cbocon1").AddItem(StrAux0)
        If VBA.CBool(.Controls("Lstcon0").List(i, 3)) Then Let j = i
        If VBA.CBool(.Controls("Lstcon0").List(i, 4)) Then Let k = i
      Next i
      Let .Controls("Cbocon0").ListIndex = j: Let .Controls("Cbocon0").Tag = .Controls("Cbocon0").Value
      Let .Controls("Cbocon1").ListIndex = k: Let .Controls("Cbocon1").Tag = .Controls("Cbocon1").Value
      Let .Controls("Btncon4").Enabled = (.Controls("Cbocon0").ListCount > 0)
      Let .Controls("Btncon5").Enabled = (.Controls("Cbocon1").ListCount > 0)
    End If

    ' Page 1
    Let .Controls("Chkcon0").Value = MDL.TableSelectionsCompact
    Let .Controls("Chkcon1").Value = MDL.TableSelectionsHeads

    ' Page 2
    Let VecAux0 = Array("Cbocon2", "Txtcon3", "Txtcon4", "Txtcon5", "Txtcon6", "Txtcon7", MDL.CompanyHeader, MDL.CompanyLogo, MDL.CompanyName, MDL.CompanyAddress, MDL.CompanyPhone, MDL.CompanyEmail)
    For j = 0 To 5
      If j = 0 Then
        Let .Controls("Cbocon2").ListIndex = VBA.CLng(VecAux0(j + 6))
        Let .Controls("Cbocon2").Tag = VBA.CLng(VecAux0(j + 6))
      Else
        Let StrAux0 = VBA.CStr(VecAux0(j + 6))
        Let .Controls(VBA.CStr(VecAux0(j))).Value = VBA.StrReverse(VBA.Mid(VBA.StrReverse(StrAux0), 4))
        Let .Controls(VBA.CStr(VecAux0(j))).Tag = StrAux0
        Let StrAux0 = VBA.Right(StrAux0, 3)
        Let StrAux0 = VBA.Replace(VBA.Replace(StrAux0, "[", VBA.vbNullString), "]", VBA.vbNullString)
        Let k = VBA.CLng(StrAux0)
        With .Controls("Chkcon" & j + 1)
          Let .Value = IIf(k = 0, False, IIf(k = 1, True, Null)): Let .Tag = k: Let .Caption = IIf(k = 0, "En encabezado de página", IIf(k = 1, "En pie de página", "En ambos"))
        End With
      End If
    Next j
    Let VecAux0 = Empty
    Let .Controls("Cbocon3").List = VBA.Split(MDL.PurDestination, ";")
    Let .Controls("Btncon8").Enabled = (.Controls("Cbocon3").ListCount > 4)
    Let .Controls("Btncon9").Enabled = (.Controls("Cbocon3").ListCount > 4)
    ' Set Form textbox controls dropbuttons
    With .Controls("Txtcon3")
      Let .DropButtonStyle = fmDropButtonStyleEllipsis
      Let .ShowDropButtonWhen = fmShowDropButtonWhenAlways
    End With

    ' Page 3
    Let .Controls("Cbocon4").ListIndex = 0
    Let .Controls("Cbocon5").List = MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalelements_sys, eqlMdlArray, "Elemento, Símbolo")
    Let .Controls("Cbocon5").ListIndex = 0
    Let .Controls("Cbocon6").List = Array("L.F.", "O.T.", "T.M.F.")
    Let .Controls("Cbocon6").ListIndex = 0
    Let .Controls("Cbocon7").List = Array("%", "g/T", "DM")
    Let .Controls("Cbocon7").ListIndex = 0
    Call View_controls_combos_change(.Controls("Cbocon4"), FRM)
    Let VecAux0 = Array(MDL.PurAverage, MDL.PurRMDetailed, MDL.PurRetentions, MDL.PurRetentionsOther, MDL.PurDiscounts, MDL.PurBonus, MDL.PurExports)
    For j = 0 To 6
      Let .Controls("Chkcon" & j + 8).Value = VBA.CBool(VecAux0(j))
    Next j
    Let VecAux0 = Empty

    ' Page 4
    Let .Controls("Lstcon2").RowSource = MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalelements_sys, eqlMdlStrTableName)
    Let .Controls("Btncon15").Enabled = (.Controls("Lstcon2").ListCount > 57)
    Let .Controls("Btncon17").Enabled = (.Controls("Lstcon2").ListCount > 57)
    Let VecAux0 = Array("Btncon14", "Btncon15", "Btncon16", "Btncon17")
    For j = LBound(VecAux0) To UBound(VecAux0)
      Let .Controls(VBA.CStr(VecAux0(j))).Enabled = (j = LBound(VecAux0))
    Next j
    Let VecAux0 = Empty

    ' Enable validations to controls
    Let VecAux0 = Array("Btncon0", "Btncon1", "Btncon3", "Btncon4", "Btncon5", "Btncon6", "Btncon7", "Btncon8", "Btncon9", "Btncon10", "Btncon11", "Btncon13", "Btncon14", "Btncon15", "Btncon17")
    For j = LBound(VecAux0) To UBound(VecAux0)
      Let .Controls(VBA.CStr(VecAux0(j))).Tag = "validate"
    Next j
    Let VecAux0 = Empty

    ' 3. [x] Fill Form Controls with data and Set general purpose CommandButtons
    Let k = (.InsideWidth - (xLeft * 2))
    With .Controls("Btnmain0") '("BtnCancel")
      Let .Cancel = True
      Let .Accelerator = "c"
      Let .Top = (.Parent.InsideHeight * 0.95) - (.Height)
      Let .Left = (xLeft + k) - .Width
      Let .Caption = "Cerrar"
    End With

    ' 4. Assing handler for each button
    Set BtnCancel = .Controls("Btnmain0") 'Main
    Set Btn0 = .Controls("Btncon0") ' Page 0
    Set Btn1 = .Controls("Btncon1") ' Page 0
    Set Btn2 = .Controls("Btncon2") ' Page 0
    Set Btn3 = .Controls("Btncon3") ' Page 0
    Set Btn4 = .Controls("Btncon4") ' Page 0
    Set Btn5 = .Controls("Btncon5") ' Page 0
    Set Btn6 = .Controls("Btncon6") ' Page 2
    Set Btn7 = .Controls("Btncon7") ' Page 2
    Set Btn8 = .Controls("Btncon8") ' Page 2
    Set Btn9 = .Controls("Btncon9") ' Page 2
    Set Btn10 = .Controls("Btncon10") ' Page 3
    Set Btn11 = .Controls("Btncon11") ' Page 3
    Set Btn12 = .Controls("Btncon12") ' Page 3
    Set Btn13 = .Controls("Btncon13") ' Page 3
    Set Btn14 = .Controls("Btncon14") ' Page 4
    Set Btn15 = .Controls("Btncon15") ' Page 4
    Set Btn16 = .Controls("Btncon16") ' Page 4
    Set Btn17 = .Controls("Btncon17") ' Page 4

    Set Chk0 = .Controls("Chkcon0") ' Page 1
    Set Chk1 = .Controls("Chkcon1") ' Page 1
    Set Chk2 = .Controls("Chkcon2") ' Page 2
    Set Chk3 = .Controls("Chkcon3") ' Page 2
    Set Chk4 = .Controls("Chkcon4") ' Page 2
    Set Chk5 = .Controls("Chkcon5") ' Page 2
    Set Chk6 = .Controls("Chkcon6") ' Page 2
    Set Chk7 = .Controls("Chkcon8") ' Page 3
    Set Chk8 = .Controls("Chkcon9") ' Page 3
    Set Chk9 = .Controls("Chkcon10") ' Page 3
    Set Chk10 = .Controls("Chkcon11") ' Page 3
    Set Chk11 = .Controls("Chkcon12") ' Page 3
    Set Chk12 = .Controls("Chkcon13") ' Page 3
    Set Chk13 = .Controls("Chkcon14") ' Page 3

    Set Txt0 = .Controls("Txtcon3") ' Page 2

    Set Cbo0 = .Controls("Cbocon3") ' Page 2
    Set Cbo1 = .Controls("Cbocon4") ' Page 3
'
    Set Lst0 = .Controls("Lstcon0") ' Page 0
'    Set Lst1 = .Controls("Lstcon1") ' Page 3
    Set Lst2 = .Controls("Lstcon2") ' Page 4

    ' 5. Give styles to controls
    Call VIEWS.ViewFormStyles(RES, FRM): Set VIEWS = Nothing

    ' Set post appearance properties to controls
    Let .Controls("Cbocon3").Style = fmStyleDropDownCombo
    Let .Controls("Multicon").Value = 0

    ' 6. [x] Show the finished Form
    Call Load(FRM)
    Call .Show

    ' 7. [x] Close form
    Let .Controls("Lstcon0").RowSource = Empty
    Let .Controls("Lstcon1").RowSource = Empty
    Let .Controls("Lstcon2").RowSource = Empty

  End With

EH:
  Let VecAux0 = Empty
  Set MpgMulti = Nothing
  Set VIEWS = Nothing
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Sub

' Help UI
Private Sub View_form_help()

  Dim StrEqURL As String

  Const sFN As String = "VIEW::-form_help"

  On Error GoTo EH

  Let StrEqURL = "https://roccouu.github.io/ExcliqLite"
  Call Application.ActiveWorkbook.FollowHyperlink(Address:=StrEqURL, NewWindow:=True)

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Sub

' About UI
Private Sub View_form_about(ByVal ControlId As String, ByVal FRM As FRMBasics, Optional VBooWhat As Boolean = False)

  ' Main objects
  Dim VIEWS As ClassEqlViewParent
  ' Standard vars
  Dim sTitle As String, TitleForm As String
  Dim i As Long
  Dim xLeft As Long

  Const sFN As String = "VIEW::-form_about"

  #If Debugging Then
    Call Debug_assert(Not ControlId = VBA.vbNullString, sFN)
    Call Debug_assert(Not FRM Is Nothing, sFN)
  #End If

  On Error GoTo EH

  ' Set Form basic datas
  Let TitleForm = IIf(VBooWhat, "Licencia", "Acerca de")
  Let sTitle = TitleForm & " - " & STRAPPTITLE

  Set VIEWS = New ClassEqlViewParent
  Call VIEWS.ViewFormCleanner(RES, FRM)
  With FRM

    ' Set the form dimensions
    Let .Height = 300
    Let .Width = 250
    Let .Tag = ControlId
    Let .Caption = sTitle
    Let .BackColor = GLOLNG_MAIN_COLOR
    Let xLeft = .InsideWidth * 0.05

    ' Build and design the form
    Call VIEWS.ViewFormAddControl(RES, FRM, "CommandButton", "Btnmain", 0) ' All with events (Cancel)
    Call VIEWS.ViewFormAddControl(RES, FRM, "Label", "Lblabt", 2, CtlVisible:=False)
    Call VIEWS.ViewFormAddControl(RES, FRM, "TextBox", "Txtabt", 0, CtlVisible:=False)

    Let i = .InsideWidth - (xLeft * 2)
    ' Rows 1
    With .Controls("Lblabt0")
      Let .Caption = "Excliq"
      Let .Font.Size = 85
      Let .ForeColor = VBA.vbWhite
      Let .AutoSize = True
      Let .WordWrap = False
      Let .Left = xLeft
      Let .Top = 0
      Let .Visible = True
    End With

    ' Rows 2
    With .Controls("Lblabt1")
      Let .Caption = "Lite"
      Let .Font.Name = "Calibri"
      Let .Font.Size = 50
      Let .Font.Italic = True
      Let .ForeColor = rgbMediumSeaGreen
      Let .AutoSize = True
      Let .WordWrap = False
      Let .Left = (.Parent.InsideWidth - (xLeft * 1.2) - .Width)
      Let .Top = .Parent.Controls("Lblabt0").Top + (.Parent.Controls("Lblabt0").Height * 0.85)
      Let .Visible = True
    End With

    ' Rows 3
    If VBooWhat Then

      With .Controls("Txtabt0")
        Let .Text = "LICENSE & ACKNOWLEDGMENTS" & VBA.vbNewLine & VBA.vbNewLine & _
          STRAPPTITLE & " v" & MDL.AppVersion & " License." & VBA.vbNewLine & VBA.vbNewLine & _
          "Copyright (c) " & MDL.AppYear & " - " & VBA.Format(VBA.Date, "YYYY") & VBA.vbNewLine & MDL.AppAuthor & VBA.vbNewLine & VBA.vbNewLine & _
          "Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the 'Software'), to deal in the Software without restriction, including without limitation the rights" & _
          "to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:" & VBA.vbNewLine & VBA.vbNewLine & _
          "The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software." & VBA.vbNewLine & VBA.vbNewLine & _
          "THE SOFTWARE IS PROVIDED 'AS IS', WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, " & _
          "FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER" & _
          "LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE." & VBA.vbNewLine & VBA.vbNewLine & _
          "Acknowledgments to:" & VBA.vbNewLine & _
          " - StarUML v2.8.0" & VBA.vbNewLine & _
          " - Pencil v3.0.3" & VBA.vbNewLine & _
          " - VSCode v1.41.0" & VBA.vbNewLine & _
          " - InkScape v0.92.4" & VBA.vbNewLine & _
          " - Gimp v2.10.2" & VBA.vbNewLine & _
          " - Just Color Picker v4.6" & VBA.vbNewLine & _
          " - Office RibbonX Editor | Fernando Andreu v1.5.1.418" & VBA.vbNewLine & _
          " - MS Excel v1911" & VBA.vbNewLine & _
          " - VBA7.1 V1091" & VBA.vbNewLine & _
          " - Paul Kelly | https://excelmacromastery.com" & VBA.vbNewLine & _
          " - Andrew Gould | https://www.wiseowl.co.uk/" & VBA.vbNewLine & _
          " - David Asurmendi | www.davidasurmendi.blogspot.com" & VBA.vbNewLine & _
          " - Ismael Romero | www.excelforo.blogspot.com" & VBA.vbNewLine & _
          " - Sergio Alejandro Campos | www.exceleinfo.com" & VBA.vbNewLine & _
          " - Otto Javier González | www.excelymas.com"
        Let .Font.Size = 9
        Let .Height = .Parent.Height * 0.35
        Let .Width = i * 0.95
        Let .Left = xLeft * 1.5
        Let .Top = .Parent.Controls("Lblabt1").Top + (.Parent.Controls("Lblabt1").Height * 0.95)
        Let .MultiLine = True
        Let .ScrollBars = fmScrollBarsVertical
        Let .Locked = True
        Let .Visible = True
      End With

    Else

      With .Controls("Lblabt2")
        Let .Caption = MDL.AppSlogan & VBA.vbNewLine & VBA.vbNewLine & _
          "Versión: v" & MDL.AppVersion & VBA.vbNewLine & _
          "Autor: " & MDL.AppAuthor & VBA.vbNewLine & _
          "E-Mail: <" & MDL.AppEmail & ">" & VBA.vbNewLine & _
          "Respositorio Git: " & MDL.AppGit & VBA.vbNewLine & _
          "Licencia: " & MDL.AppLicense & VBA.vbNewLine & _
          MDL.AppYear & " - " & VBA.Format(VBA.Date, "YYYY") & VBA.vbNewLine & _
          MDL.AppOrigin
        Let .Font.Size = 9
        Let .ForeColor = rgbIvory
        Let .AutoSize = True
        Let .Width = i
        Let .Left = xLeft * 1.5
        Let .Top = .Parent.Controls("Lblabt1").Top + (.Parent.Controls("Lblabt1").Height * 0.95)
        Let .Visible = True
      End With

    End If



    ' 3. [x] Fill Form Controls with data and Set general purpose CommandButtons
    With .Controls("Btnmain0") '("BtnCancel")
      Let .Cancel = True
      Let .Accelerator = "c"
      Let .Top = -.Height
      Let .Left = -.Width
      Let .Caption = "Cerrar"
    End With

    ' 4. Assing handler for each button
    Set BtnCancel = .Controls("Btnmain0") 'Main

    ' 5. Give styles to controls
    Call VIEWS.ViewFormStyles(RES, FRM): Set VIEWS = Nothing

    ' 6. [x] Show the finished Form
    Call Load(FRM)
    Call .Show

    ' 7. [x] Close form

  End With

EH:
  Set VIEWS = Nothing
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Sub



' CORE VIEW ASSETS ----------------------------------------------------------------- '
Private Sub View_assets_initialize_model()

  On Error GoTo EH
  ' Define model
  Set MDL = New ClassEqlModel
  Set MDL.ResourcesApp = RES

  Let GLOBOO_SELECTIONCELL = MDL.TableSelectionsCell
  Let GLOBOO_COMPACT = MDL.TableSelectionsCompact
  Let GLOBOO_HEADS = MDL.TableSelectionsHeads
  Let GLOBYT_TABLES = MDL.Tables

EH:
  Call RES.ErrorHandler.ErrorHandlerDisplay("VIEW::-initialize_model")

End Sub

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
  Set Btn16 = Nothing
  Set Btn17 = Nothing

  Set Chk0 = Nothing
  Set Chk1 = Nothing
  Set Chk2 = Nothing
  Set Chk3 = Nothing
  Set Chk4 = Nothing
  Set Chk5 = Nothing
  Set Chk6 = Nothing
  Set Chk7 = Nothing
  Set Chk8 = Nothing
  Set Chk9 = Nothing
  Set Chk10 = Nothing
  Set Chk11 = Nothing
  Set Chk12 = Nothing
  Set Chk13 = Nothing

  Set Cbo0 = Nothing
  Set Cbo1 = Nothing
  Set Cbo2 = Nothing
  Set Cbo3 = Nothing
  Set Cbo4 = Nothing

  Set Lst0 = Nothing
  Set Lst1 = Nothing
  Set Lst2 = Nothing

  Set Txt0 = Nothing
  Set Txt1 = Nothing
  Set Txt2 = Nothing
  Set Txt3 = Nothing
  Set Txt4 = Nothing
  Set Txt5 = Nothing
  Set Txt6 = Nothing
  Set Txt7 = Nothing
  Set Txt8 = Nothing
  Set Txt9 = Nothing
  Set Txt10 = Nothing
  Set Txt11 = Nothing
  Set Txt12 = Nothing
  Set Txt13 = Nothing
  Set Txt14 = Nothing
  Set Txt15 = Nothing

  Set Spn0 = Nothing
  Set Spn1 = Nothing
  Set Spn2 = Nothing
  Set Spn3 = Nothing
  Set Spn4 = Nothing
  Set Spn5 = Nothing

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise("VIEW::-ControlsFree")

End Sub

Private Function View_assets_controls_getformparent(ByVal UControl As MSForms.Control) As FRMBasics

  ' Search across each container control as Frames, MultiPages, etc., to find its main UserForm Parent
  Dim CtlParent As Object

  On Error GoTo EH

  Set CtlParent = UControl.Parent ' Find control's parent UserForm
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
  Call RES.ErrorHandler.ErrorHandlerDisplay("VIEW::-Assets_GetFormParent")

End Function

' Controls Helper methods
Private Sub View_assets_closeform(ByVal FRM As FRMBasics)

  On Error GoTo EH
  Call VBA.Unload(FRM)

EH:
  Call RES.ErrorHandler.ErrorHandlerDisplay("VIEW::-assets_closeform")

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
  Optional VBooRngAddressToTxt As Boolean = False, _
  Optional RDblTop As Double = -1, _
  Optional VBooJustOneWord As Boolean = False)

  Dim VWP As ClassEqlViewParent
  Dim RngBox As Range
  Dim VecAux0 As Variant
  Dim i As Long

  On Error GoTo EH

  With VFRM

    Set VWP = New ClassEqlViewParent
    Set RngBox = VWP.ViewsCaptureRange(RES, VFRM, VStrTitle, STRAPPTITLE)

    If Not RngBox Is Nothing Then

      Select Case VEnuType

        Case eqlResJustRanges
          Set GLORNG_CELL = Nothing
          Set GLORNG_CELL = RngBox

        Case eqlResRanges ' Ranges for addres on txt
          Set GLORNG_CELL = Nothing
          Set GLORNG_CELL = RngBox
          If VBooRngAddressToTxt Then Let VTxt.Value = RngBox.Range("A1").Address

        Case eqlResNotNumbers
          If VBooJustOneWord Then
            Let VTxt.Value = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(RngBox, i, RValueType:=VEnuType, RBooAbs:=VBooAbs, RBooIncludeZeroes:=VBooZeroes)
            If RES.REGEX.isStringReg(VTxt.Text) Then
              Let VecAux0 = VBA.Split(VTxt.Text, ";")
              Let VTxt.Text = VBA.CStr(VecAux0(0))
              Let VecAux0 = Empty
            End If
          Else
            Let VTxt.Value = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(RngBox, i, RValueType:=VEnuType, RBooAbs:=VBooAbs, RBooIncludeZeroes:=VBooZeroes)
          End If

        Case eqlResDates
          Let VTxt.Value = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(RngBox, i, RValueType:=VEnuType, RBooAbs:=VBooAbs, RBooIncludeZeroes:=VBooZeroes)

        Case eqlResStrings
          Let VTxt.Value = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(RngBox, i, RValueType:=VEnuType, RBooAbs:=VBooAbs, RBooIncludeZeroes:=VBooZeroes)

        Case eqlResNumbers
          If VBooSlashGrouped Then ' Grouped by /
            Let VTxt.Value = RES.ArrayGetDataSemicolonSeparatedSlashGroupedStringFromVariants(RngBox, i, False, VBooAbs, VBooZeroes, RBooRngRowsCols:=VBooRngCols)
          Else ' ; separation
            If VBooSum Then ' Sum elements
              Let VTxt.Value = VBA.vbNullString
              Let VecAux0 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(RngBox, i, RValueType:=VEnuType, RBooAbs:=VBooAbs, RBooIncludeZeroes:=VBooZeroes, RDblTop:=RDblTop)
              If i > 0 Then
                Let VTxt.Value = Application.WorksheetFunction.Sum(VecAux0)
                If VBA.CDbl(VTxt.Value) >= RDblTop Then Call VBA.MsgBox("La suma de los datos seleccionados es un número muy grande. Seleccione otros datos", vbInformation, STRAPPTITLE): Let VTxt.Value = VBA.vbNullString
              End If
              Let VecAux0 = Empty
            Else
              Let VTxt.Value = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(RngBox, i, RValueType:=VEnuType, RBooAbs:=VBooAbs, RBooIncludeZeroes:=VBooZeroes)
            End If
          End If

      End Select
      Set RngBox = Nothing

    End If

  End With

EH:
  Set RngBox = Nothing
  Set VWP = Nothing
  Call RES.ErrorHandler.ErrorHandlerRaise("VIEW::-assets_capture_data_from_ranges")

End Sub

Private Function View_assets_get_selected_or_new_range(ByVal ControlId As String, ByRef VVarReq As Variant) As Boolean

  ' [Create new sheet and assign its first range] or assign a user selected range
  Dim RngBox As Range
  Dim StrName As String
  Dim BooNewSheet As Boolean
  Dim LngVecReqIndex As Long

  Const sFN As String = "VIEW::-assets_get_selected_or_new_range"

  On Error GoTo EH
  Select Case ControlId

    Case "regression", "regressionselection": Let BooNewSheet = VBA.CBool(VVarReq(15)): Let LngVecReqIndex = 16: Let StrName = "Regresión"
    Case "balmet", "balmetselection", "balmeteco": Let BooNewSheet = VBA.CBool(VVarReq(13)): Let LngVecReqIndex = 9: Let StrName = "Balmet"
    Case "projection": Let BooNewSheet = VBA.CBool(VVarReq(15)): Let LngVecReqIndex = 16: Let StrName = "Proyecto"

    Case "oreinvoice": Let BooNewSheet = VBA.CBool(VVarReq(6)): Let LngVecReqIndex = 7: Let StrName = "Liqmin"
    Case "concentrateinvoice": Let BooNewSheet = VBA.CBool(VVarReq(1)): Let LngVecReqIndex = 2: Let StrName = "Liqcnc"

    Case "hlpregression", "hlprlselection", "hlpbalmet", "hlpbalmetselection", "hlpbalmeteco", "hlpprojection", "hlporeinvoice", "hlpconcentrateinvoice", "hlpudf"
      Let BooNewSheet = VBA.CBool(VVarReq(1)): Let LngVecReqIndex = 0: Let StrName = "AyudaExcliqLite"

    Case Else: GoTo EH
  End Select

  If Not BooNewSheet Then Set RngBox = VVarReq(LngVecReqIndex)
  Set VVarReq(LngVecReqIndex) = View_assets_getcell(BooNewSheet, RngBox, StrName)

  Let View_assets_get_selected_or_new_range = (Not VVarReq(LngVecReqIndex) Is Nothing)

EH:
  Set RngBox = Nothing
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function View_assets_getcell( _
  ByVal VBooNewSheet As Boolean, _
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

    If ShtNewSheet Is Nothing Then Call VBA.Err.Raise(CUSTOM_ERROR_VWS, sFN, "No se pudo crear la nueva hoja: '" & VStrSheetName & "' para imrpimir los resultados.")

    Set View_assets_getcell = ShtNewSheet.Range("A1")
    Set ShtNewSheet = Nothing

  Else

    With VRngBox
      ' Get range
      Set View_assets_getcell = .Range("A1")
    End With

  End If

EH:
  If Not VBA.Err.Number = 0 And Not ShtNewSheet Is Nothing Then Call ShtNewSheet.Parent.Worksheets(ShtNewSheet.Name).Delete
  Set ShtNewSheet = Nothing
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Sub View_assets_read_data_for_balmet(ByVal VTxtCurrent As MSForms.TextBox, ByVal FRM As FRMBasics)

  Dim StrAux0 As String, StrAux1 As String, StrAux2 As String
  Dim i As Long, j As Long

  Const StrBALPRODUCTS As String = "Conjunto de leyes inválido." & VBA.vbNewLine & _
      "La referencia debe ser un rango de celdas o un grupo de números separados por punto y coma (Ej.: 1;2;3) con las siguientes dimensiones: " & VBA.vbNewLine & VBA.vbNewLine & _
      "Para Balance de 2 Productos: 3 filas x 1 columna." & VBA.vbNewLine & _
      "Para Balance de 3 Productos: 4 filas x 2 columnas." & VBA.vbNewLine & _
      "Para Balance de 4 Productos: 5 filas x 3 columnas..., etc." & VBA.vbNewLine & VBA.vbNewLine & _
      "Puede generar balances de hasta " & CLNG_BAL_LIMIT & " productos"
  Const sFN As String = "VIEW:-assets_read_data_for_balmet"

  On Error GoTo EH
  Select Case FRM.Tag

    Case "balmet" ' ---------------------------------------------------------- balmet'
      With VTxtCurrent
        Select Case .Name
          Case "Txtbal0": Call View_assets_capture_data_from_ranges(FRM, VTxtCurrent, FRM.Controls("Lblbal1").Caption & " seleccione uno o varios datos numéricos:", eqlResNumbers, True)

          Case "Txtbal2"
            Call View_assets_capture_data_from_ranges(FRM, VTxtCurrent, FRM.Controls("Lblbal2").Caption & " seleccione un rango de datos numéricos:", eqlResNumbers, True)
            If Not RES.REGEX.isNumberEntireAndDecimalVectorColonSeparated(.Text) Then GoTo EH
            Let i = UBound(VBA.Split(.Text, ";")) + 1
            If Not RES.RESCORE.BalIsBalmet(i) Then Call VBA.MsgBox(StrBALPRODUCTS & ".", vbInformation, STRAPPTITLE): Let .Text = VBA.vbNullString: GoTo EH
            Let FRM.Controls("Txtbal3").Text = RES.RESCORE.GetGenericNamesAndUnits(.Text, StrAux0, StrAux1)
            Let .Tag = StrAux0
            Let FRM.Controls("Txtbal4").Text = RES.RESCORE.GetGenericNamesAndUnits(.Text, StrAux0, StrAux1, BBooUnits:=True)

          Case "Txtbal5"
            If Not FRM.Controls("Txtbal2").Text = VBA.vbNullString Then
              If FRM.Controls("Chkbal3").Value Then Call View_assets_capture_data_from_ranges(FRM, VTxtCurrent, FRM.Controls("Lblbal3").Caption & " seleccione un rango de datos numéricos:", eqlResNumbers, True)
              If Not RES.REGEX.isNumberEntireAndDecimalVectorColonSeparated(.Text) Then GoTo EH
              Let i = UBound(VBA.Split(FRM.Controls("Txtbal2").Text, ";")) + 1
              Let j = UBound(VBA.Split(.Text, ";")) + 1
              If Not RES.RESCORE.BalIsBalmet(i, j) Then Call VBA.MsgBox(StrBALPRODUCTS & " y hasta " & CLNG_BAL_LIMITCX & " tipos de minerales complejos, estos además, deben tener igual cantidad de filas que las leyes del balance principal.", vbInformation, STRAPPTITLE): Let .Text = VBA.vbNullString: GoTo EH
              Let FRM.Controls("Txtbal6").Text = RES.RESCORE.GetGenericNamesAndUnits(FRM.Controls("Txtbal2").Text, StrAux0, StrAux1, .Text, BBooCx:=True)
              Let .Tag = StrAux1
              Let FRM.Controls("Txtbal7").Text = RES.RESCORE.GetGenericNamesAndUnits(FRM.Controls("Txtbal2").Text, StrAux0, StrAux1, .Text, BBooCx:=True, BBooUnits:=True)
            End If

          Case "Txtbal8": If FRM.Controls("Optbal1").Value Then Call View_assets_capture_data_from_ranges(FRM, VTxtCurrent, "Seleccione una celda donde imprimir el resultado:", eqlResRanges, VBooRngAddressToTxt:=True)
        End Select
      End With

    Case "balmeteco" ' ---------------------------------------------------------- balmeteco'
      With VTxtCurrent
        Select Case .Name

          Case "Txtbal0": Call View_assets_capture_data_from_ranges(FRM, VTxtCurrent, FRM.Controls("Lblbal2").Caption & " Seleccione uno o varios datos numéricos:", eqlResNumbers, True)

          Case "Txtbal1"
            Call View_assets_capture_data_from_ranges(FRM, VTxtCurrent, FRM.Controls("Lblbal3").Caption & " Seleccione uno o varios datos numéricos:", eqlResNumbers, True)
            If Not RES.REGEX.isNumberEntireAndDecimalVectorColonSeparated(.Text) Then GoTo EH
            Let i = UBound(VBA.Split(.Text, ";"))
            If (i - 1) > CLNG_BAL_LIMIT Then Call VBA.MsgBox("Puede procesar balances metalúrgicos de hasta " & CLNG_BAL_LIMIT & " productos.", vbInformation, STRAPPTITLE): Let .Text = VBA.vbNullString: GoTo EH
            Let FRM.Controls("Txtbal5").Text = IIf(FRM.Controls("Chkbal2").Value, VBA.vbNullString, RES.RESCORE.BalGetGenericNames(i))
            If Not FRM.Controls("Chkbal0").Value Then Let FRM.Controls("Txtbal3").Text = RES.RESCORE.BalGetGenericRecoveries(i)
            Let FRM.Controls("Txtbal3").Tag = i
            Let .Tag = i

          Case "Txtbal2"
            Call View_assets_capture_data_from_ranges(FRM, VTxtCurrent, FRM.Controls("Lblbal4").Caption & " Seleccione un rango con datos numéricos:", eqlResNumbers, True)
            If Not RES.REGEX.isNumberEntireAndDecimalVectorColonSeparated(.Text) Then GoTo EH
            Let i = UBound(VBA.Split(.Text, ";"))
            If (i - 1) > CLNG_BAL_LIMIT Then Call VBA.MsgBox("Puede procesar balances metalúrgicos de hasta " & CLNG_BAL_LIMIT & " productos.", vbInformation, STRAPPTITLE): Let .Text = VBA.vbNullString: GoTo EH
            Let FRM.Controls("Txtbal6").Text = IIf(FRM.Controls("Chkbal3").Value, VBA.vbNullString, RES.RESCORE.BalGetGenericUnits(i))
            If Not FRM.Controls("Chkbal0").Value Then Let FRM.Controls("Txtbal3").Text = RES.RESCORE.BalGetGenericRecoveries(i)
            Let .Tag = i

          Case "Txtbal3"
            If Not FRM.Controls("Chkbal0").Value Then GoTo EH
            Call View_assets_capture_data_from_ranges(FRM, VTxtCurrent, FRM.Controls("Chkbal0").Caption & " Seleccione un rango con datos numéricos:", eqlResNumbers, True)
            If Not RES.REGEX.isNumberEntireAndDecimalVectorColonSeparated(.Text) Then GoTo EH
            Let i = UBound(VBA.Split(.Text, ";"))
            If (i - 1) > CLNG_BAL_LIMIT Then Call VBA.MsgBox("Puede procesar balances metalúrgicos de hasta " & CLNG_BAL_LIMIT & " productos.", vbInformation, STRAPPTITLE): Let .Text = VBA.vbNullString: GoTo EH

          Case "Txtbal7"
            If Not FRM.Controls("Chkbal4").Value Then GoTo EH
            Call View_assets_capture_data_from_ranges(FRM, VTxtCurrent, FRM.Controls("Lblbal5").Caption & " Seleccione un rango con datos numéricos:", eqlResNumbers, True)
            If Not RES.REGEX.isNumberEntireAndDecimalVectorColonSeparated(.Text) Then GoTo EH
            Let i = UBound(VBA.Split(.Text, ";"))
            If (i - 1) > CLNG_BAL_LIMITCX Then Call VBA.MsgBox("Puede procesar balances metalúrgicos de hasta " & CLNG_BAL_LIMIT & " productos y hasta " & CLNG_BAL_LIMITCX & " productos de minerales complejos.", vbInformation, STRAPPTITLE): Let .Text = VBA.vbNullString: GoTo EH
            Let FRM.Controls("Txtbal8").Text = IIf(FRM.Controls("Chkbal5").Value, VBA.vbNullString, RES.RESCORE.BalGetGenericNames(i, BBooCx:=True))
            Let FRM.Controls("Txtbal9").Text = IIf(FRM.Controls("Chkbal6").Value, VBA.vbNullString, RES.RESCORE.BalGetGenericUnits(i, BBooCx:=True))
            Let FRM.Controls("Txtbal10").Text = IIf(FRM.Controls("Chkbal7").Value, VBA.vbNullString, RES.RESCORE.BalGetGenericRecoveries(i))
            Let FRM.Controls("Txtbal10").Tag = i
            Let .Tag = i

          Case "Txtbal10"
            If Not FRM.Controls("Chkbal4").Value And Not FRM.Controls("Chkbal7").Value Then GoTo EH
            Call View_assets_capture_data_from_ranges(FRM, VTxtCurrent, FRM.Controls("Chkbal7").Caption & " Seleccione un rango con datos numéricos:", eqlResNumbers, True)
            If Not RES.REGEX.isNumberEntireAndDecimalVectorColonSeparated(.Text) Then GoTo EH
            Let i = UBound(VBA.Split(.Text, ";"))
            If (i - 1) > CLNG_BAL_LIMIT Then Call VBA.MsgBox("Puede procesar balances metalúrgicos de hasta " & CLNG_BAL_LIMIT & " productos.", vbInformation, STRAPPTITLE): Let .Text = VBA.vbNullString: GoTo EH

          Case "Txtbal11": If FRM.Controls("Optbal1").Value Then Call View_assets_capture_data_from_ranges(FRM, VTxtCurrent, "Seleccione una celda donde imprimir el resultado:", eqlResRanges, VBooRngAddressToTxt:=True)

        End Select
      End With

  End Select

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Sub

Private Function View_assets_clear_last_colonseparatedstring_item(ByVal StrTxtColonSeparated As String) As String

  Dim StrAux0 As String

  Const sFN As String = "VIEW:-assets_clear_last_colonseparatedstring_item"

  On Error GoTo EH
  Let StrAux0 = StrTxtColonSeparated
  If Not StrAux0 = VBA.vbNullString Then
    Let StrAux0 = VBA.StrReverse(StrAux0)
    Let StrAux0 = VBA.StrReverse(VBA.Mid(StrAux0, VBA.InStr(1, StrAux0, ";", vbTextCompare) + 1))
  End If

  Let View_assets_clear_last_colonseparatedstring_item = StrAux0

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function View_assets_check_for_spreadsheet_edges( _
  ByVal VRngBox As Range, _
  ByVal VLngRows As Long, _
  ByVal VLngCols As Long, _
  ByRef VStrMsg As String, _
  Optional VBooSelections As Boolean = False) As Boolean

  ' Check for spreadsheet edges and data inside the region selected
  Dim LngRowsAprox As Long, LngColsAprox As Long

  Const sFN As String = "VIEW:-assets_check_for_spreadsheet_edges"

  On Error GoTo EH
  Let View_assets_check_for_spreadsheet_edges = False

  Let LngRowsAprox = (VLngRows * 5) ' 5: A1 to totalsrow aprox
  Let LngColsAprox = (VLngCols * 2) ' 2: A1 to totalscol aprox

  If VBooSelections Then

    If RES.ExcelEdges(VRngBox, xRow:=LngRowsAprox, xCol:=LngColsAprox) Then Let VStrMsg = "La celda que eligió está muy cerca de los bordes de la hoja de cálculo, este ejercicio requiere más espacio. Seleccione otra celda mpas alejada de los bordes.": GoTo EH
    If (Application.WorksheetFunction.CountA(VRngBox.Range("A1").Resize(LngRowsAprox, LngColsAprox)) > 0) Then
      If VBA.MsgBox(STRAPPTITLE & " sobreescribirá cualquier dato que se encuentre al rededor de la celda seleccionada. Existen datos que podrían ser eliminados." & VBA.vbNewLine & VBA.vbNewLine & "¿Continuar de todos modos?", vbYesNo + vbQuestion, STRAPPTITLE) = vbNo Then GoTo EH
    End If

  Else

    If VRngBox Is Nothing Then Let VStrMsg = "¡Debe seleccionar una celda para imprimir los resultados!": GoTo EH
    If RES.ExcelEdges(VRngBox, LngRowsAprox, LngColsAprox) Then Let VStrMsg = "La celda que eligió está muy cerca de los bordes de la hoja de cálculo, este ejercicio requiere más espacio. Seleccione otra celda mpas alejada de los bordes.": GoTo EH
    If (Application.WorksheetFunction.CountA(VRngBox.Range("A1").Resize(LngRowsAprox, LngColsAprox)) > 0) Then
      If VBA.MsgBox(STRAPPTITLE & " sobreescribirá cualquier dato que se encuentre alrededor de la celda seleccionada. Existen datos que podrían ser eliminados." & VBA.vbNewLine & VBA.vbNewLine & "¿Continuar de todos modos?", vbYesNo + vbQuestion, STRAPPTITLE) = vbNo Then GoTo EH
    End If

  End If

  Let View_assets_check_for_spreadsheet_edges = True

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Sub View_assets_oreinvoice_basedata( _
  ByVal xCombo As MSForms.ComboBox, _
  ByVal FRM As FRMBasics, _
  ByVal VLngPurRows As Long)

  Dim VecAux0 As Variant
  Dim BooAux0 As Boolean
  Dim i As Long

  Const sFN As String = "VIEW::-assets_oreinvoice_basedata"

  On Error GoTo EH

  If xCombo.ListIndex < 0 Then GoTo EH
  With FRM

    ' Layout & Status (& some values)
    With .Controls("Txtpur6")
      Let .Value = VBA.Format$(VBA.Date(), "dd/mm/yy")
      Let .DropButtonStyle = fmDropButtonStylePlain
      Let .ShowDropButtonWhen = fmShowDropButtonWhenAlways
    End With
    Let VecAux0 = Array("El nombre de la muestra del mineral. Ejemplo: M-1", "Ejemplo: dd/mm/aa o dd/mm/aaaa", "El Porcentaje de humedad del mineral. Ejemplo: 5", "El Porcentaje de Merma del mineral. Ejemplo: 1", "El Peso del mineral en Toneladas. Ejemplo: 1000")
    For i = 5 To 9
      Let .Controls("Txtpur" & i).ControlTipText = VBA.CStr(VecAux0(i - 5))
    Next i
    Let VecAux0 = Empty
    ' Values Data A
    Let .Controls("Lstpur0").RowSource = VBA.CStr(MDL.MGet(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur, eqlMdlStrTableName))

    ' Set List properties
    With .Controls("Lstpur0")
      Let .ColumnHeads = False
      Let .MultiSelect = fmMultiSelectSingle
      Let .ListStyle = fmListStylePlain
      Let .ColumnCount = 5
      Let .ColumnWidths = (.Width / 5) * 0.9
      Let .ColumnHeads = True
      Let .IntegralHeight = False
      Let .ControlTipText = "Doble click (o pulse la Barra Espaciadora) en el ítem que desee modificar"
      Let .Visible = True
      Let .Enabled = (VLngPurRows > 0)

      Let BooAux0 = (.ListCount = GLOSNG_PUR_LIMIT)
    End With
    ' Set buttons status
    Let .Controls("Btnpur0").Enabled = IIf(BooAux0, False, VLngPurRows >= 0)
    Let .Controls("Btnpur1").Enabled = False
    Let .Controls("Btnpur2").Enabled = False
    Let .Controls("Btnpur3").Enabled = (VLngPurRows > 1)
    Let .Controls("Btnpur4").Enabled = (VLngPurRows > 0)
    Let .Controls("Btnpur5").Enabled = IIf(BooAux0, False, VLngPurRows >= 0)

  End With

EH:
  Let VecAux0 = Empty
  Call RES.ErrorHandler.ErrorHandlerDisplay(sFN)

End Sub

Private Sub View_assets_oreinvoice_maindata( _
  ByVal xCombo As MSForms.ComboBox, _
  ByVal FRM As FRMBasics, _
  ByVal VLngPurRows As Long)

  Dim VecAux0 As Variant, VecAux1 As Variant
  Dim BooAux0 As Boolean
  Dim i As Long, j As Long

  Const sFN As String = "VIEW::-assets_oreinvoice_maindata"

  On Error GoTo EH

  With FRM

    ' Layout & Status (& some values)
    For i = 5 To 9
      Let .Controls("Txtpur" & i).Visible = (i = 7 Or i = 9)
      If i < 8 Then Let .Controls("Cbopur" & i - 3).Visible = True
    Next i
    With .Controls("Cbopur2")
      Let .Width = .Parent.Controls("Txtpur5").Width
      Let .Left = .Parent.Controls("Txtpur5").Left
      Let .RowSource = VBA.vbNullString: Call .Clear
      Let .List = Array("Mineral", "Mineral Complejo", "Brosa")
      Let .ListIndex = 0
    End With
    Let .Controls("Cbopur3").Left = .Controls("Txtpur6").Left
    Let .Controls("Cbopur4").Left = .Controls("Txtpur8").Left
    Let VecAux0 = Array("Seleccione el tipo de mineral", "Seleccione el mineral o los minerales", "Seleccione la Unidad de ley del mineral", "Seleccione un mineral de la lista", "Introduzca la ley o las leyes del mineral (separados por punto y coma si son complejos). Ejemplo: 5", "Seleccione una unidad de ley de la lista", "El Precio del mineral. Ejemplo: 3.5", "Cbopur2", "Cbopur3", "Cbopur4", "Txtpur6", "Txtpur7", "Txtpur8", "Txtpur9")
    For i = 0 To 6
      Let .Controls(VBA.CStr(VecAux0(i + 7))).ControlTipText = VBA.CStr(VecAux0(i))
    Next i
    Let VecAux0 = Empty

    ' Values Data B
    Let .Controls("Lstpur0").RowSource = VBA.CStr(MDL.MGet(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur, eqlMdlStrTableName))

    ' Set List properties
    With .Controls("Lstpur0")
      Let .ColumnHeads = False
      Let .MultiSelect = fmMultiSelectSingle
      Let .ListStyle = fmListStylePlain
      Let .ColumnCount = 14
      Let j = (.Width / 5) * 0.9
      Let .ColumnWidths = "0;0;0;0;0;" & j & ";" & j & ";" & j & ";" & j & ";" & j & ";0;0;0;0"
      Let .ColumnHeads = True
      Let .IntegralHeight = False
      Let .ControlTipText = "Doble click (o pulse la Barra Espaciadora) en el ítem que desee modificar"
      Let .Enabled = (VLngPurRows > 0)
      Let .Visible = True

      Let BooAux0 = (.ListCount = GLOSNG_PUR_LIMIT)
    End With

    ' Set buttons status
    Let .Controls("Btnpur0").Enabled = IIf(BooAux0, False, VLngPurRows > 0)
    Let .Controls("Btnpur1").Enabled = False
    Let .Controls("Btnpur2").Enabled = False
    Let .Controls("Btnpur3").Enabled = (VLngPurRows > 1)
    Let .Controls("Btnpur4").Enabled = (VLngPurRows > 0)
    Let .Controls("Btnpur5").Enabled = IIf(BooAux0, False, VLngPurRows > 0)

  End With

EH:
  Let VecAux0 = Empty
  Call RES.ErrorHandler.ErrorHandlerDisplay(sFN)

End Sub

Private Sub View_assets_oreinvoice_bonusdiscounts( _
  ByVal xCombo As MSForms.ComboBox, _
  ByVal FRM As FRMBasics, _
  ByVal VLngPurRows As Long)

  Dim VecAux0 As Variant, VecAux1 As Variant
  Dim BooAux0 As Boolean, BooAux1 As Boolean
  Dim xLeft As Long
  Dim i As Long, j As Long, k As Long

  Const sFN As String = "VIEW::-assets_oreinvoice_bonusdiscounts"

  On Error GoTo EH

  With FRM

    ' Left for all main controls
    Let xLeft = .InsideWidth * 0.05

    ' Clear tables before work with data
    Let k = xCombo.ListIndex
    If VLngPurRows > 0 Then Call MDL.MDel(eqlMdlSheetTmp, IIf(k = 2, eqlMdlTblPurbonus_tmp, eqlMdlTblPurdiscounts_tmp))
    ' Layout & Status (& some values)
    For i = 5 To 9
      Let .Controls("Txtpur" & i).Visible = (i < 7)
    Next i
    Let j = (.Controls("Framepur2").InsideWidth - (xLeft * 2) + 5) / IIf(.Controls("Chkpur1").Value, 3, 2)
    If .Controls("Chkpur1").Value Then
      With .Controls("Cbopur2")
        Let .Width = j
        Let .Left = .Parent.Controls("Cbopur1").Left
        Let .ControlTipText = "Seleccione el código de muestra de cada liquidación y asigne los " & .Parent.Controls("Cbopur1").Value
      End With
    End If
    With .Controls("Txtpur5")
      Let .Width = j
      Let .Left = xCombo.Left + IIf(FRM.Controls("Chkpur1").Value, .Parent.Controls("Cbopur2").Width, 0) '(.Parent.InsideWidth - (xLeft * 2) + 3.5) / 2
      Let .MaxLength = 30
      Let .ControlTipText = IIf(xCombo.ListIndex = 2, "El nombre del Bono. Ejemplo: Bono transporte", "El nombre del Descuento. Ejemplo: Ensayos químicos")
    End With
    With .Controls("Txtpur6")
      Let .Width = j
      Let .Left = .Parent.Controls("Txtpur5").Left + .Parent.Controls("Txtpur5").Width
      Let .ControlTipText = IIf(xCombo.ListIndex = 2, "El valor del Bono. Ejemplo: 500", "El valor del Descuento. Ejemplo: 500")
      Let .Value = VBA.vbNullString
      Let .DropButtonStyle = fmDropButtonStyleReduce
      Let .ShowDropButtonWhen = fmShowDropButtonWhenAlways
    End With
    ' Fill list if multiple providers is not selected
    If .Controls("Chkpur1").Value Then
      With .Controls("Cbopur2")
        Let .Tag = "noevent"
        Let .RowSource = VBA.vbNullString: Call .Clear
        If VLngPurRows > 0 Then Let .List = MDL.MGet(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur, eqlMdlArray, "MUESTRA")
        Let .Tag = VBA.vbNullString
        If VLngPurRows > 0 Then Let .ListIndex = 0
        Let .Visible = True
      End With
    Else
      If VLngPurRows > 0 Then
        Let VecAux0 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur, eqlMdlArray, IIf(k = 2, "BONOS", "DESCUENTOS"), 1), eqlRes1D)
        If VBA.IsArray(VecAux0) Then
          If Not VBA.CStr(VecAux0(0)) = VBA.vbNullString Then
            Let VecAux1 = VBA.Split(VBA.CStr(VecAux0(0)), "|") ' Separe Concept | Values
            Let VecAux0 = RES.ArrayToBaseZero(VBA.Split(VBA.CStr(VecAux1(0)), ";"), eqlRes2D, False)  ' Separe Concept1 ; Concept2...
            For i = 0 To UBound(VecAux0)
              Let VecAux0(i, 0) = VBA.Replace(VBA.Replace(VBA.CStr(VecAux0(i, 0)), "[eqlsemicolon]", ";"), "[eqlpipe]", "|")
            Next i
            Let VecAux0 = RES.ArrayConcat(VecAux0, VBA.Split(VBA.CStr(VecAux1(1)), ";"), RByCols:=True) ' Join Concepts - Values
            Let VecAux0 = RES.ArrayAddAtFirst(VecAux0, "M-" & Format(Now(), "yymmddssms"), True) ' Join #Sample - Concepts - Values
            Call MDL.MSet(eqlMdlSheetTmp, IIf(k = 2, eqlMdlTblPurbonus_tmp, eqlMdlTblPurdiscounts_tmp), VecAux0)
          End If
        End If
        Let VecAux0 = Empty
        Let VecAux1 = Empty
      End If
    End If
    Let .Controls("Lstpur0").RowSource = VBA.CStr(MDL.MGet(eqlMdlSheetTmp, IIf(k = 2, eqlMdlTblPurbonus_tmp, eqlMdlTblPurdiscounts_tmp), eqlMdlStrTableName))

    ' Set List properties
    Let BooAux1 = False
    If VLngPurRows > 0 Then Let BooAux1 = (MDL.LengthRows(eqlMdlSheetTmp, IIf(xCombo.ListIndex = 2, eqlMdlTblPurbonus_tmp, eqlMdlTblPurdiscounts_tmp)) > 0)
    With .Controls("Lstpur0")
      Let .ColumnHeads = False
      Let .MultiSelect = fmMultiSelectSingle
      Let .ListStyle = fmListStylePlain
      Let .ColumnCount = 3
      Let .ColumnWidths = IIf(FRM.Controls("Chkpur1").Value, VBA.vbNullString, "0;;")
      Let .ColumnHeads = True
      Let .IntegralHeight = False
      Let .ControlTipText = "Doble click (o pulse la Barra Espaciadora) en el ítem que desee modificar"
      Let .Enabled = (VLngPurRows > 0 And BooAux1)
      Let .Visible = True

      Let BooAux0 = (.ListCount = GLOSNG_PUR_LIMIT Or VLngPurRows <= 0)
    End With

    ' Set buttons status
    Let .Controls("Btnpur0").Enabled = Not BooAux0
    Let .Controls("Btnpur1").Enabled = False
    Let .Controls("Btnpur2").Enabled = False
    Let .Controls("Btnpur3").Enabled = BooAux1
    Let .Controls("Btnpur4").Enabled = BooAux1
    Let .Controls("Btnpur5").Enabled = Not BooAux0

  End With

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Call RES.ErrorHandler.ErrorHandlerDisplay(sFN)

End Sub

Private Sub View_assets_oreinvoice_retentions( _
  ByVal xCombo As MSForms.ComboBox, _
  ByVal FRM As FRMBasics, _
  ByVal VLngPurRows As Long)

  Dim VecAux0 As Variant
  Dim i As Long, j As Long, k As Long

  Const sFN As String = "VIEW::-assets_oreinvoice_retentions"

  On Error GoTo EH

  With FRM

    ' Clear tables before work with data & Update LO SELECTIONS TO 0
    Let k = xCombo.ListIndex
    Let j = MDL.LengthRows(eqlMdlSheetTmp, IIf(k = 3, eqlMdlTblConfigpurretentions_tmp, eqlMdlTblConfigpurretentionsother_tmp))
    If VLngPurRows > 0 And j > 0 Then Call MDL.MUpd(eqlMdlSheetTmp, IIf(k = 3, eqlMdlTblConfigpurretentions_tmp, eqlMdlTblConfigpurretentionsother_tmp), Array(0), "SELECTED")

    ' Layout & Status (& some values)
    With .Controls("Cbopur2")
      Let .ControlTipText = "Seleccione el código de muestra de cada liquidación y asigne datos de " & .Parent.Controls("Cbopur1").Value
      Let .Left = FRM.Controls("Cbopur1").Left
      Let .Enabled = (VLngPurRows > 0)
      Let .Visible = FRM.Controls("Chkpur1").Value
    End With
    With .Controls("Chkpur2")
      Let .Left = FRM.Controls("Cbopur1").Left + IIf(FRM.Controls("Chkpur1").Value, FRM.Controls("Cbopur2").Width + 5, 0)
      Let .Caption = "Todas las " & FRM.Controls("Cbopur1").Value
      Let .Enabled = (VLngPurRows > 0)
      Let .Visible = True
    End With

    ' Values Retentions, Other retentions
    If .Controls("Chkpur1").Value And k >= 0 Then ' Just for Multiple providers
      With .Controls("Cbopur2")
        Let .Tag = "noevent"
        Let .RowSource = Empty: Call .Clear
        If VLngPurRows > 0 Then Let .List = MDL.MGet(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur, eqlMdlArray, "MUESTRA")
        Let .Tag = VBA.vbNullString
        If VLngPurRows > 0 Then Let .ListIndex = IIf(FRM.Controls("Cbopur1").Tag = "FillList", .ListIndex, 0)
      End With
    Else ' Get SELECTED values in array
      If VLngPurRows > 0 Then
        Let VecAux0 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur, eqlMdlArray, IIf(k = 3, "RETENCIONES", "RETENCIONES OTRAS"), 1), eqlRes1D)
        If VBA.IsArray(VecAux0) Then
          If Not VBA.CStr(VecAux0(0)) = VBA.vbNullString Then
            Let VecAux0 = VBA.Split(VBA.CStr(VecAux0(0)), ";")
            Call MDL.MUpd(eqlMdlSheetTmp, IIf(k = 3, eqlMdlTblConfigpurretentions_tmp, eqlMdlTblConfigpurretentionsother_tmp), Array(VecAux0), "SELECTED")
            'For i = 1 To j
            '  Call MDL.MUpd(eqlMdlSheetTmp, IIf(k = 3, eqlMdlTblConfigpurretentions_tmp, eqlMdlTblConfigpurretentionsother_tmp), Array(VBA.CByte(VecAux0(i - 1))), "SELECTED", i)
            'Next i
          End If
        End If
      End If
    End If
    Let .Controls("Lstpur0").RowSource = VBA.CStr(MDL.MGet(eqlMdlSheetTmp, IIf(xCombo.ListIndex = 3, eqlMdlTblConfigpurretentions_tmp, eqlMdlTblConfigpurretentionsother_tmp), eqlMdlStrTableName))

    ' Set List properties
    With .Controls("Lstpur0")
      Let .ColumnHeads = False
      Let .MultiSelect = fmMultiSelectMulti
      Let .ListStyle = fmListStyleOption
      Let .ColumnCount = 2
      Let .ColumnWidths = VBA.vbNullString
      Let .ColumnHeads = True
      Let .IntegralHeight = False
      Let .ControlTipText = "Seleccione los ítems de " & xCombo.Value & " que desea habilitar"
      Let .Enabled = (VLngPurRows > 0)
      Let .Visible = True

      ' Mark all previously selected items in the list
      If .ListCount > 0 Then
        Let j = 0
        For i = 0 To .ListCount - 1
          Let .Selected(i) = VBA.CBool(.List(i, 2))
          Let j = j + IIf(.Selected(i), 1, 0)
        Next i
      End If
      Let FRM.Controls("Chkpur2").Tag = "noevent"
      Let FRM.Controls("Chkpur2").Value = (j = .ListCount)
      Let FRM.Controls("Chkpur2").Tag = VBA.vbNullString
      Let .Tag = VBA.vbNullString
    End With

    ' Set buttons status
    Let .Controls("Btnpur0").Enabled = False
    Let .Controls("Btnpur1").Enabled = False
    Let .Controls("Btnpur2").Enabled = False
    Let .Controls("Btnpur3").Enabled = False
    Let .Controls("Btnpur4").Enabled = False
    Let .Controls("Btnpur5").Enabled = False

  End With

EH:
  Let VecAux0 = Empty
  Call RES.ErrorHandler.ErrorHandlerDisplay(sFN)

End Sub

Private Sub View_assets_cncinvoice_contents( _
  ByVal xCombo As MSForms.ComboBox, _
  ByVal FRM As FRMBasics)

  Dim VecAux0 As Variant
  Dim i As Long, j As Long, k As Long, xLeft As Long

  Const sFN As String = "VIEW::-assets_cncinvoice_contents"

  On Error GoTo EH

  With FRM

    ' Set controls layout
    Let xLeft = .InsideWidth * 0.05 ' Left for all main controls
    Let j = (.Controls("Framecnc2").InsideWidth - (xLeft * 2)) / 5
    Let VecAux0 = Array("Cbocnc2", "Txtcnc11", "Txtcnc12", "Cbocnc3", "Txtcnc14")
    For i = LBound(VecAux0) To UBound(VecAux0)
      With .Controls(VBA.CStr(VecAux0(i)))
        Let .Width = j
        Let .Left = (.Parent.Controls("Cbocnc1").Left + (.Width * i))
        If i > 1 Then Let .ControlTipText = IIf(i = 2, "Deducción mínima", IIf(i = 3, "Seleccione una opción", "Pagable al... [%]"))
        Let .Visible = True
      End With
    Next i
    Let VecAux0 = Empty

    ' Set Combo2 and List values
    With .Controls("Cbocnc2")
      Let .Tag = "noevent"
      Let .List = MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigpurrm_sys, eqlMdlArray, "MINERAL, ELEMENTO, UNIDAD MINERAL")
      Let .ListIndex = 0
    End With
    Let .Controls("Txtcnc11").ControlTipText = "Ley de " & .Controls("Cbocnc2").List(0, 1) & " en: " & .Controls("Cbocnc2").List(0, 2)
    Set Cbo1 = .Controls("Cbocnc2")

    Let k = MDL.LengthRows(eqlMdlSheetPur, eqlMdlTblpurchasecnccontents_pur)
    With .Controls("Lstcnc0")
      Let .ColumnHeads = False
      Let .RowSource = MDL.MGet(eqlMdlSheetPur, eqlMdlTblpurchasecnccontents_pur, eqlMdlStrTableName)
      Let .ColumnCount = 5
      Let j = (.Width / 5) * 0.95
      Let .ColumnWidths = j & ";" & j & ";" & j & ";" & j & ";" & j & ";0;0;0;0"
      Let .ColumnHeads = True
      Let .ControlTipText = "Doble click (o pulse la Barra Espaciadora) en el ítem que desee modificar"
      Let .Enabled = (k > 0)
    End With

    ' Set buttons status
    Let .Controls("Btncnc0").Enabled = (k <= GLOSNG_PUR_LIMIT)
    Let .Controls("Btncnc1").Enabled = False
    Let .Controls("Btncnc2").Enabled = False
    Let .Controls("Btncnc3").Enabled = (k <= GLOSNG_PUR_LIMIT And k > 1)
    Let .Controls("Btncnc4").Enabled = (k <= GLOSNG_PUR_LIMIT And k > 0)
    Let .Controls("Btncnc5").Enabled = (k <= GLOSNG_PUR_LIMIT)

  End With

EH:
  Let VecAux0 = Empty
  Call RES.ErrorHandler.ErrorHandlerDisplay(sFN)

End Sub

Private Sub View_assets_cncinvoice_refination( _
  ByVal xCombo As MSForms.ComboBox, _
  ByVal FRM As FRMBasics)

  Dim VecAux0 As Variant
  Dim i As Long, j As Long, k As Long, xLeft As Long

  Const sFN As String = "VIEW::-assets_cncinvoice_refination"

  On Error GoTo EH

  With FRM

    ' Set controls layout
    Let xLeft = .InsideWidth * 0.05 ' Left for all main controls
    Let j = (.Controls("Framecnc2").InsideWidth - (xLeft * 2)) / 3
    For i = 0 To 2
      With .Controls(IIf(i = 0, "Cbocnc2", IIf(i = 1, "Txtcnc11", "Txtcnc12")))
        Let .ControlTipText = IIf(i = 0, "Seleccione un ítem para modificar los datos", IIf(i = 1, "Costo Base", "Pagable al... [%]"))
        Let .Width = j
        Let .Left = (xLeft + (.Width * i))
        Let .Visible = True
      End With
    Next i

    ' Set Combo2 and List values
    Let k = MDL.LengthRows(eqlMdlSheetPur, eqlMdlTblpurchasecnccontents_pur)
    With .Controls("Lstcnc0")
      Let .ColumnHeads = False
      Let .RowSource = MDL.MGet(eqlMdlSheetPur, eqlMdlTblpurchasecnccontents_pur, eqlMdlStrTableName)
      If k > 0 Then Let VecAux0 = RES.ArrayShift(.List, True, True)
      Let .ColumnCount = 9
      Let .ColumnWidths = "0;0;0;0;0;0;;;"
      Let .ColumnHeads = True
      Let .ControlTipText = "Doble click (o pulse la Barra Espaciadora) en el ítem que desee modificar"
      Let .Enabled = (k > 0)
    End With
    If k > 0 Then
      For i = LBound(VecAux0) To UBound(VecAux0)
        Let VecAux0(i, 0) = "Refinación de " & VBA.CStr(VecAux0(i, 0))
      Next i
    End If

    With .Controls("Cbocnc2")
      Let .Tag = "noevent"
      If k > 0 Then Let .List = VecAux0
      If k > 0 Then Let .ListIndex = 0
      Let .Enabled = (k > 0)
    End With
    Let VecAux0 = Empty
    If k > 0 Then
      Let .Controls("Txtcnc11").Text = .Controls("Lstcnc0").List(0, 7)
      Let .Controls("Txtcnc12").Text = .Controls("Lstcnc0").List(0, 8)
      Set Cbo1 = .Controls("Cbocnc2")
    End If

    ' Set buttons status
    Let .Controls("Btncnc0").Enabled = False
    Let .Controls("Btncnc1").Enabled = (k <= GLOSNG_PUR_LIMIT And k > 0)
    Let .Controls("Btncnc2").Enabled = False
    Let .Controls("Btncnc3").Enabled = (k <= GLOSNG_PUR_LIMIT And k > 1)
    Let .Controls("Btncnc4").Enabled = (k <= GLOSNG_PUR_LIMIT And k > 0)
    Let .Controls("Btncnc5").Enabled = (k <= GLOSNG_PUR_LIMIT And k > 0)

  End With

EH:
  Let VecAux0 = Empty
  Call RES.ErrorHandler.ErrorHandlerDisplay(sFN)

End Sub

Private Sub View_assets_cncinvoice_penalties( _
  ByVal xCombo As MSForms.ComboBox, _
  ByVal FRM As FRMBasics)

  Dim i As Long, j As Long, k As Long, xLeft As Long

  Const sFN As String = "VIEW::-assets_cncinvoice_penalties"

  On Error GoTo EH

  With FRM

    ' Set controls layout
    Let xLeft = .InsideWidth * 0.05 ' Left for all main controls
    With .Controls("Chkcnc1")
      Let .Caption = "Seleccionar todas las " & .Parent.Controls("Cbocnc1").Value
      Let .Left = xLeft
      Let .Visible = True
    End With

    ' Set Combo2 and List values
    Let k = MDL.LengthRows(eqlMdlSheetPur, eqlMdlTblpurchasecncpenalties_pur)
    'Set Lst0 = Nothing
    With .Controls("Lstcnc0")
      Let .ColumnHeads = False
      Let .RowSource = MDL.MGet(eqlMdlSheetPur, eqlMdlTblpurchasecncpenalties_pur, eqlMdlStrTableName)
      Let j = (.Width / 5) * 0.95
      Let .ColumnCount = 6
      Let .ColumnWidths = j & ";" & j & ";" & j & ";" & j & ";" & j & ";0"
      Let .ColumnHeads = True
      Let .MultiSelect = fmMultiSelectMulti
      Let .ListStyle = fmListStyleOption
      Let .ControlTipText = "Doble click (o pulse la Barra Espaciadora) en el ítem que desee modificar"
      Let .Enabled = (k > 0)
      ' Highlight selected items (Set tag=empty)
      If k > 0 Then
        Let i = 0
        For j = 0 To k - 1
          Let .Selected(j) = VBA.CBool(.List(j, 5))
          Let i = i + IIf(.Selected(j), 1, 0)
        Next j
      End If
      Let .Parent.Controls("Chkcnc1").Value = (i = .ListCount)
    End With
    'Set Lst0 = .Controls("Lstcnc0")

    ' Set buttons status
    Let .Controls("Btncnc0").Enabled = False
    Let .Controls("Btncnc1").Enabled = False
    Let .Controls("Btncnc2").Enabled = False
    Let .Controls("Btncnc3").Enabled = False '(k <= GLOSNG_PUR_LIMIT And k > 7)
    Let .Controls("Btncnc4").Enabled = False '(k <= GLOSNG_PUR_LIMIT And k > 6)
    Let .Controls("Btncnc5").Enabled = False

  End With

EH:
  Call RES.ErrorHandler.ErrorHandlerDisplay(sFN)

End Sub

Private Sub View_assets_cncinvoice_deductions( _
  ByVal xCombo As MSForms.ComboBox, _
  ByVal FRM As FRMBasics)

  Const sFN As String = "VIEW::-assets_cncinvoice_deductions"

  On Error GoTo EH

  With FRM

    ' Set controls layout
    With .Controls("Cbocnc2")
      Let .Tag = Empty
      Call .Clear
      Let .List = Array("Descuentos por Exportación", "Retenciones de Ley", "Retenciones Institucionales", "Descuentos Otros")
      Let .Width = .Parent.Controls("Cbocnc1").Width / 2
      Let .Visible = True
      Set Cbo1 = .Parent.Controls("Cbocnc2")
      Let .ListIndex = 0
    End With

  End With

EH:
  Call RES.ErrorHandler.ErrorHandlerDisplay(sFN)

End Sub

Private Sub View_assets_cncinvoice_bonus( _
  ByVal xCombo As MSForms.ComboBox, _
  ByVal FRM As FRMBasics)

  Dim i As Long, j As Long, k As Long, xLeft As Long

  Const sFN As String = "VIEW::-assets_cncinvoice_bonus"

  On Error GoTo EH

  With FRM

    ' Set controls layout
    Let xLeft = .InsideWidth * 0.05 ' Left for all main controls
    Let j = (.Controls("Framecnc2").InsideWidth - (xLeft * 2)) / 2
    For i = 0 To 1
      With .Controls(IIf(i = 0, "Txtcnc11", "Txtcnc12"))
        Let .ControlTipText = IIf(i = 0, "Concepto del bono", "Valor del bono")
        Let .Width = j
        Let .Left = (xLeft + (.Width * i))
        Let .Visible = True
      End With
    Next i

    ' Set Combo2 and List values
    Let k = MDL.LengthRows(eqlMdlSheetTmp, eqlMdlTblPurbonus_tmp)
    With .Controls("Lstcnc0")
      Let .ColumnHeads = False
      Let .RowSource = MDL.MGet(eqlMdlSheetTmp, eqlMdlTblPurbonus_tmp, eqlMdlStrTableName)
      Let .ColumnCount = 3
      Let .ColumnWidths = "0;;"
      Let .ColumnHeads = True
      Let .ControlTipText = "Doble click (o pulse la Barra Espaciadora) en el ítem que desee modificar"
      Let .Enabled = (k > 0)
    End With

    ' Set buttons status
    Let .Controls("Btncnc0").Enabled = (k <= GLOSNG_PUR_LIMIT)
    Let .Controls("Btncnc1").Enabled = False
    Let .Controls("Btncnc2").Enabled = False
    Let .Controls("Btncnc3").Enabled = (k <= GLOSNG_PUR_LIMIT And k > 1)
    Let .Controls("Btncnc4").Enabled = (k <= GLOSNG_PUR_LIMIT And k > 0)
    Let .Controls("Btncnc5").Enabled = (k <= GLOSNG_PUR_LIMIT)

  End With

EH:
  Call RES.ErrorHandler.ErrorHandlerDisplay(sFN)

End Sub

Private Sub View_assets_cncinvoice_dedexports( _
  ByVal xCombo As MSForms.ComboBox, _
  ByVal FRM As FRMBasics)

  Dim i As Long, j As Long, k As Long, xLeft As Long

  Const sFN As String = "VIEW::-assets_cncinvoice_bonus"

  On Error GoTo EH

  With FRM

    ' Set controls layout
    Let xLeft = .InsideWidth * 0.05 ' Left for all main controls
    Let xCombo.Width = .Controls("Cbocnc1").Width / 2
    Let xCombo.Left = xLeft
    With .Controls("Chkcnc1")
      Let .Caption = "Todos los " & xCombo.Value
      Let .Left = (xCombo.Left + xCombo.Width + 2)
      Let .Visible = True
    End With

    ' Set Combo2 and List values
    Let k = MDL.LengthRows(eqlMdlSheetPur, eqlMdlTblpurchasecncdedexp_pur)
    'Set Lst0 = Nothing
    With .Controls("Lstcnc0")
      Let .ColumnHeads = False
      Let .RowSource = MDL.MGet(eqlMdlSheetPur, eqlMdlTblpurchasecncdedexp_pur, eqlMdlStrTableName)
      Let .ColumnCount = 5
      Let .ColumnWidths = ";;;;0"
      Let .MultiSelect = fmMultiSelectMulti
      Let .ListStyle = fmListStyleOption
      Let .ColumnHeads = True
      Let .ControlTipText = "Seleccione los " & xCombo.Value & " que desee aplicar. Doble click o Barra Espaciadora para modificar"
      Let .Enabled = (k > 0)
      ' Highlight selected items (Set tag=empty)
      If k > 0 Then
        Let i = 0
        For j = 0 To k - 1
          Let .Selected(j) = VBA.CBool(.List(j, 4))
          Let i = i + IIf(.Selected(j), 1, 0)
        Next j
      End If
      Let .Parent.Controls("Chkcnc1").Value = (i = .ListCount)
    End With
    'Set Lst0 = .Controls("Lstcnc0")

    ' Set buttons status
    Let .Controls("Btncnc0").Enabled = False
    Let .Controls("Btncnc1").Enabled = False
    Let .Controls("Btncnc2").Enabled = False
    Let .Controls("Btncnc3").Enabled = False
    Let .Controls("Btncnc4").Enabled = False
    Let .Controls("Btncnc5").Enabled = False

  End With

EH:
  Call RES.ErrorHandler.ErrorHandlerDisplay(sFN)

End Sub

Private Sub View_assets_cncinvoice_dedtax( _
  ByVal xCombo As MSForms.ComboBox, _
  ByVal FRM As FRMBasics)

  Dim i As Long, j As Long, k As Long, xLeft As Long

  Const sFN As String = "VIEW::-assets_cncinvoice_bonus"

  On Error GoTo EH

  With FRM

    ' Set controls layout
    Let xLeft = .InsideWidth * 0.05 ' Left for all main controls
    Let xCombo.Width = .Controls("Cbocnc1").Width / 2
    Let xCombo.Left = xLeft
    With .Controls("Chkcnc1")
      Let .Caption = "Todas las " & xCombo.Value
      Let .Left = (xCombo.Left + xCombo.Width + 2)
      Let .Visible = True
    End With

    ' Set Combo2 and List values
    Let k = MDL.LengthRows(eqlMdlSheetTmp, eqlMdlTblConfigpurretentions_tmp)
    'Set Lst0 = Nothing
    With .Controls("Lstcnc0")
      Let .ColumnHeads = False
      Let .RowSource = MDL.MGet(eqlMdlSheetTmp, eqlMdlTblConfigpurretentions_tmp, eqlMdlStrTableName)
      Let .ColumnCount = 2
      Let .ColumnWidths = ";"
      Let .MultiSelect = fmMultiSelectMulti
      Let .ListStyle = fmListStyleOption
      Let .ColumnHeads = True
      Let .ControlTipText = "Seleccione los ítems de " & xCombo.Value & " que desee aplicar"
      Let .Enabled = (k > 0)
      ' Highlight selected items (Set tag=empty)
      If k > 0 Then
        Let i = 0
        For j = 0 To k - 1
          Let .Selected(j) = VBA.CBool(.List(j, 2))
          Let i = i + IIf(.Selected(j), 1, 0)
        Next j
      End If
      Let .Parent.Controls("Chkcnc1").Value = (i = .ListCount)
    End With
    'Set Lst0 = .Controls("Lstcnc0")

    ' Set buttons status
    Let .Controls("Btncnc0").Enabled = False
    Let .Controls("Btncnc1").Enabled = False
    Let .Controls("Btncnc2").Enabled = False
    Let .Controls("Btncnc3").Enabled = False
    Let .Controls("Btncnc4").Enabled = False
    Let .Controls("Btncnc5").Enabled = False

  End With

EH:
  Call RES.ErrorHandler.ErrorHandlerDisplay(sFN)

End Sub

Private Sub View_assets_cncinvoice_dedothertax( _
  ByVal xCombo As MSForms.ComboBox, _
  ByVal FRM As FRMBasics)

  Dim i As Long, j As Long, k As Long, xLeft As Long

  Const sFN As String = "VIEW::-assets_cncinvoice_bonus"

  On Error GoTo EH

  With FRM

    ' Set controls layout
    Let xLeft = .InsideWidth * 0.05 ' Left for all main controls
    Let xCombo.Width = .Controls("Cbocnc1").Width / 2
    Let xCombo.Left = xLeft
    With .Controls("Chkcnc1")
      Let .Caption = "Todas las " & xCombo.Value
      Let .Left = (xCombo.Left + xCombo.Width + 2)
      Let .Visible = True
    End With

    ' Set Combo2 and List values
    Let k = MDL.LengthRows(eqlMdlSheetTmp, eqlMdlTblConfigpurretentionsother_tmp)
    'Set Lst0 = Nothing
    With .Controls("Lstcnc0")
      Let .ColumnHeads = False
      Let .RowSource = MDL.MGet(eqlMdlSheetTmp, eqlMdlTblConfigpurretentionsother_tmp, eqlMdlStrTableName)
      Let .ColumnCount = 2
      Let .ColumnWidths = ";"
      Let .MultiSelect = fmMultiSelectMulti
      Let .ListStyle = fmListStyleOption
      Let .ColumnHeads = True
      Let .ControlTipText = "Seleccione los ítems de " & xCombo.Value & " que desee aplicar"
      Let .Enabled = (k > 0)
      ' Highlight selected items (Set tag=empty)
      If k > 0 Then
        Let i = 0
        For j = 0 To k - 1
          Let .Selected(j) = VBA.CBool(.List(j, 2))
          Let i = i + IIf(.Selected(j), 1, 0)
        Next j
      End If
      Let .Parent.Controls("Chkcnc1").Value = (i = .ListCount)
    End With
    'Set Lst0 = .Controls("Lstcnc0")

    ' Set buttons status
    Let .Controls("Btncnc0").Enabled = False
    Let .Controls("Btncnc1").Enabled = False
    Let .Controls("Btncnc2").Enabled = False
    Let .Controls("Btncnc3").Enabled = False
    Let .Controls("Btncnc4").Enabled = False
    Let .Controls("Btncnc5").Enabled = False

  End With

EH:
  Call RES.ErrorHandler.ErrorHandlerDisplay(sFN)

End Sub

Private Sub View_assets_cncinvoice_deddiscounts( _
  ByVal xCombo As MSForms.ComboBox, _
  ByVal FRM As FRMBasics)

  Dim i As Long, j As Long, k As Long, xLeft As Long

  Const sFN As String = "VIEW::-assets_cncinvoice_bonus"

  On Error GoTo EH

  With FRM

    ' Set controls layout
    Let xLeft = .InsideWidth * 0.05 ' Left for all main controls
    Let j = (.Controls("Framecnc2").InsideWidth - (xLeft * 2)) / 3
    Let xCombo.Width = j '.Controls("Cbocnc1").Width / 2
    Let xCombo.Left = xLeft

    For i = 0 To 1
      With .Controls(IIf(i = 0, "Txtcnc11", "Txtcnc12"))
        Let .ControlTipText = IIf(i = 0, "Concepto del descuento", "Valor del descuento")
        Let .Width = j
        Let .Left = (xCombo.Left + xCombo.Width) + (.Width * i)
        Let .Visible = True
      End With
    Next i

    ' Set Combo2 and List values
    Let k = MDL.LengthRows(eqlMdlSheetTmp, eqlMdlTblPurdiscounts_tmp)
    With .Controls("Lstcnc0")
      Let .ColumnHeads = False
      Let .RowSource = MDL.MGet(eqlMdlSheetTmp, eqlMdlTblPurdiscounts_tmp, eqlMdlStrTableName)
      Let .ColumnCount = 3
      Let .ColumnWidths = "0;;"
      Let .ColumnHeads = True
      Let .ControlTipText = "Doble click (o pulse la Barra Espaciadora) en el ítem que desee modificar"
      Let .Enabled = (k > 0)
    End With

    ' Set buttons status
    Let .Controls("Btncnc0").Enabled = (k <= GLOSNG_PUR_LIMIT)
    Let .Controls("Btncnc1").Enabled = False
    Let .Controls("Btncnc2").Enabled = False
    Let .Controls("Btncnc3").Enabled = (k <= GLOSNG_PUR_LIMIT And k > 1)
    Let .Controls("Btncnc4").Enabled = (k <= GLOSNG_PUR_LIMIT And k > 0)
    Let .Controls("Btncnc5").Enabled = (k <= GLOSNG_PUR_LIMIT)

  End With

EH:
  Call RES.ErrorHandler.ErrorHandlerDisplay(sFN)

End Sub

Private Sub View_assets_config_rm( _
  ByVal xCombo As MSForms.ComboBox, _
  ByVal FRM As FRMBasics)

  Dim VecAux0 As Variant
  Dim i As Long, j As Long, k As Long, xLeft As Long

  Const sFN As String = "VIEW::-assets_config_rm"

  On Error GoTo EH
  With FRM

    ' Set controls layout
    Let xLeft = .InsideWidth * 0.05 ' Left for all main controls
    Let j = .Controls("Lstcon1").Width / 7

    Let VecAux0 = Array("Cbocon5", "Cbocon6", "Txtcon8", "Txtcon9", "Txtcon10", "Cbocon7", "Nombre y símbolo del elemento o especie mineral", "Unidad del elemento o especie mineral enriquecido en proceso metalúrgico", "Cotización LME del elemento o especie mineral", "Regalía Minera del elemento o especie mineral para exportación", "Regalía Minera del elemento o especie mineral para venta interna", "Unidad del elemento o especie mineral")
    For i = 0 To 5
      With .Controls(VBA.CStr(VecAux0(i)))
        Let .ControlTipText = VBA.CStr(VecAux0(i + 6))
        Let .Width = j * IIf(i = 0, 2, 1)
        Let .Left = xLeft + IIf(i > 0, .Width, 0) + (.Width * i)
        Let .Visible = True
      End With
    Next i
    With .Controls("Cbocon5")
      Let .ColumnCount = 2
      Let i = (.Width / 2) * 0.97
      Let .ColumnWidths = i & ";" & i
      Let .BoundColumn = 1
      Let .TextColumn = 1
    End With
    Let VecAux0 = Empty

    ' Set Combo2 and List values
    Let k = MDL.LengthRows(eqlMdlSheetSys, eqlMdlTblConfigpurrm_sys)
    With .Controls("Lstcon1")
      Let .ColumnHeads = False
      Let .RowSource = MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigpurrm_sys, eqlMdlStrTableName)
      Let .ColumnCount = 7
      Let i = (.Width / 7) * 0.97
      Let VecAux0 = Array(i, i, i, i, i, i, i)
      Let .ColumnWidths = VBA.Join(VecAux0, ";") ' ";;;;;;"
      Let VecAux0 = Empty
      Let .ColumnHeads = True
      Let .Enabled = (k > 0)
    End With

    ' Set buttons status
    Let .Controls("Btncon10").Enabled = (k <= 50)
    Let .Controls("Btncon11").Enabled = False
    Let .Controls("Btncon12").Enabled = False
    Let .Controls("Btncon13").Enabled = False '(k <= 50 And k > 12)

  End With

EH:
  Let VecAux0 = Empty
  Call RES.ErrorHandler.ErrorHandlerDisplay(sFN)

End Sub

Private Sub View_assets_config_retentions( _
  ByVal xCombo As MSForms.ComboBox, _
  ByVal FRM As FRMBasics)

  Dim VecAux0 As Variant
  Dim i As Long, j As Long, k As Long, xLeft As Long

  Const sFN As String = "VIEW::-assets_config_retentions"

  On Error GoTo EH

  With FRM

    ' Set controls layout
    Let xLeft = .InsideWidth * 0.05 ' Left for all main controls
    Let j = .Controls("Lstcon1").Width / 2

    Let VecAux0 = Array("Txtcon8", "Txtcon9", "Concepto de la retención", "Porcentaje de retención")
    For i = 0 To 1
      With .Controls(VBA.CStr(VecAux0(i)))
        Let .ControlTipText = VBA.CStr(VecAux0(i + 2))
        Let .Width = j
        Let .Left = xLeft + (.Width * i)
        Let .Visible = True
      End With
    Next i
    Let VecAux0 = Empty

    ' Set Combo2 and List values
    Let k = MDL.LengthRows(eqlMdlSheetTmp, eqlMdlTblConfigpurretentions_tmp)
    With .Controls("Lstcon1")
      Let .ColumnHeads = False
      Let .RowSource = MDL.MGet(eqlMdlSheetTmp, eqlMdlTblConfigpurretentions_tmp, eqlMdlStrTableName)
      Let .ColumnCount = 2
      Let .ColumnWidths = ";"
      Let .ColumnHeads = True
      Let .Enabled = (k > 0)
    End With

    ' Set buttons status
    Let .Controls("Btncon10").Enabled = (k <= GLOSNG_PUR_LIMIT)
    Let .Controls("Btncon11").Enabled = False
    Let .Controls("Btncon12").Enabled = False
    Let .Controls("Btncon13").Enabled = False '(k <= GLOSNG_PUR_LIMIT And k > 3)

  End With

EH:
  Let VecAux0 = Empty
  Call RES.ErrorHandler.ErrorHandlerDisplay(sFN)

End Sub

Private Sub View_assets_config_retentionsother( _
  ByVal xCombo As MSForms.ComboBox, _
  ByVal FRM As FRMBasics)

  Dim VecAux0 As Variant
  Dim i As Long, j As Long, k As Long, xLeft As Long

  Const sFN As String = "VIEW::-assets_config_retentionsother"

  On Error GoTo EH

  With FRM

    ' Set controls layout
    Let xLeft = .InsideWidth * 0.05 ' Left for all main controls
    Let j = .Controls("Lstcon1").Width / 2

    Let VecAux0 = Array("Txtcon8", "Txtcon9", "Concepto de la retención", "Porcentaje de retención")
    For i = 0 To 1
      With .Controls(VBA.CStr(VecAux0(i)))
        Let .ControlTipText = VBA.CStr(VecAux0(i + 2))
        Let .Width = j
        Let .Left = xLeft + (.Width * i)
        Let .Visible = True
      End With
    Next i
    Let VecAux0 = Empty

    ' Set Combo2 and List values
    Let k = MDL.LengthRows(eqlMdlSheetTmp, eqlMdlTblConfigpurretentionsother_tmp)
    With .Controls("Lstcon1")
      Let .ColumnHeads = False
      Let .RowSource = MDL.MGet(eqlMdlSheetTmp, eqlMdlTblConfigpurretentionsother_tmp, eqlMdlStrTableName)
      Let .ColumnCount = 2
      Let .ColumnWidths = ";"
      Let .ColumnHeads = True
      Let .Enabled = (k > 0)
    End With

    ' Set buttons status
    Let .Controls("Btncon10").Enabled = (k <= GLOSNG_PUR_LIMIT)
    Let .Controls("Btncon11").Enabled = False
    Let .Controls("Btncon12").Enabled = False
    Let .Controls("Btncon13").Enabled = False '(k <= GLOSNG_PUR_LIMIT And k > 4)

  End With

EH:
  Let VecAux0 = Empty
  Call RES.ErrorHandler.ErrorHandlerDisplay(sFN)

End Sub

Private Sub View_assets_config_penalties( _
  ByVal xCombo As MSForms.ComboBox, _
  ByVal FRM As FRMBasics)

  Dim VecAux0 As Variant
  Dim i As Long, j As Long, k As Long, xLeft As Long

  Const sFN As String = "VIEW::-assets_config_penalties"

  On Error GoTo EH

  With FRM

    ' Set controls layout
    Let xLeft = .InsideWidth * 0.05 ' Left for all main controls
    Let j = .Controls("Lstcon1").Width / 5

    Let VecAux0 = Array("Cbocon5", "Txtcon8", "Txtcon9", "Txtcon10", "Txtcon11", "Nombre del elemento o especie penalizable", "Ley del elemento o especie penalizable", "Ley base del elemento o especie por encima del cual será penalizable", "Monto a deducir (en " & MDL.CurrencyMainSymbolForeign & ")", "Límite por cada punto de ley del elemento o especie mineral penalizable")
    For i = 0 To 4
      With .Controls(VBA.CStr(VecAux0(i)))
        Let .ControlTipText = VBA.CStr(VecAux0(i + 5))
        Let .Width = j
        Let .Left = xLeft + (.Width * i)
        Let .Visible = True
      End With
    Next i
    With .Controls("Cbocon5")
      Let .ColumnCount = 2
      Let .ColumnWidths = "0;"
      Let .BoundColumn = 2
      Let .TextColumn = 2
    End With
    Let VecAux0 = Empty

    ' Set Combo2 and List values
    Let k = MDL.LengthRows(eqlMdlSheetPur, eqlMdlTblpurchasecncpenalties_pur)
    With .Controls("Lstcon1")
      Let .ColumnHeads = False
      Let .RowSource = MDL.MGet(eqlMdlSheetPur, eqlMdlTblpurchasecncpenalties_pur, eqlMdlStrTableName)
      Let .ColumnCount = 5
      Let .ColumnWidths = ";;;;"
      Let .ColumnHeads = True
      Let .Enabled = (k > 0)
    End With

    ' Set buttons status
    Let .Controls("Btncon10").Enabled = (k <= GLOSNG_PUR_LIMIT)
    Let .Controls("Btncon11").Enabled = False
    Let .Controls("Btncon12").Enabled = False
    Let .Controls("Btncon13").Enabled = False '(k <= GLOSNG_PUR_LIMIT And k > 6)

  End With

EH:
  Let VecAux0 = Empty
  Call RES.ErrorHandler.ErrorHandlerDisplay(sFN)

End Sub

Private Sub View_assets_config_exports( _
  ByVal xCombo As MSForms.ComboBox, _
  ByVal FRM As FRMBasics)

  Dim VecAux0 As Variant
  Dim i As Long, j As Long, k As Long, xLeft As Long

  Const sFN As String = "VIEW::-assets_config_exports"

  On Error GoTo EH

  With FRM

    ' Set controls layout
    Let xLeft = .InsideWidth * 0.05 ' Left for all main controls
    Let j = .Controls("Lstcon1").Width / 4

    Let VecAux0 = Array("Txtcon8", "Txtcon9", "Txtcon10", "Chkcon7", "Concepto de la deducción en exportación (Ej. Rollback)", "Opcional. Porcentaje de descuento aplicable al valor bruto de la liquidación de concentrados", "Opcional. Valor de descuento aplicable al tonelaje bruto húmedo de la liquidación de concentrados", "Opcional. El valor que introduzca será descontado directamente en la liquidación de conentrados")
    For i = 0 To 3
      With .Controls(VBA.CStr(VecAux0(i)))
        Let .ControlTipText = VBA.CStr(VecAux0(i + 4))
        Let .Width = j
        Let .Left = xLeft + (.Width * i)
        Let .Visible = True
      End With
    Next i
    Let VecAux0 = Empty

    ' Set Combo2 and List values
    Let k = MDL.LengthRows(eqlMdlSheetPur, eqlMdlTblpurchasecncdedexp_pur)
    With .Controls("Lstcon1")
      Let .ColumnHeads = False
      Let .RowSource = MDL.MGet(eqlMdlSheetPur, eqlMdlTblpurchasecncdedexp_pur, eqlMdlStrTableName)
      Let .ColumnCount = 4
      Let .ColumnWidths = ";;;"
      Let .ColumnHeads = True
      Let .Enabled = (k > 0)
    End With

    ' Set buttons status
    Let .Controls("Btncon10").Enabled = (k <= GLOSNG_PUR_LIMIT)
    Let .Controls("Btncon11").Enabled = False
    Let .Controls("Btncon12").Enabled = False
    Let .Controls("Btncon13").Enabled = False '(k <= GLOSNG_PUR_LIMIT And k > 0)

  End With

EH:
  Let VecAux0 = Empty
  Call RES.ErrorHandler.ErrorHandlerDisplay(sFN)

End Sub





' ERROR ASSERTSIONS -----------------------------------------------------------------'
Private Sub Debug_assert(ByVal Cond As Boolean, Optional Origin As String = VBA.vbNullString)

  #If Debbuging Then
    Debug.Assert Cond
    If Not Cond Then Debug.Print "Assert on: " & Origin
  #End If

End Sub

Private Function View_core_error_handler(ByVal ControlId As String, ByVal FRM As FRMBasics) As Boolean

  Const sFN As String = "VIEW::-core_error_handler"

  On Error GoTo EH
  Let View_core_error_handler = False

  With RES.ErrorHandler
    Select Case ControlId

      Case "regression"
        Select Case .ErrorNum
          Case 0, 1: Call CMB.CMsgBox(.ErrorMsg, STRAPPTITLE, "Error", cmbError, cmbNoButtons, FRM, "Txtreg1", "BtnMain1", cmbRequired, cmbCenter, cmbMiddle, cDuration:=cmb3Seconds)
          Case 2: Call CMB.CMsgBox(.ErrorMsg, STRAPPTITLE, "Error", cmbError, cmbNoButtons, FRM, "Txtreg0", "BtnMain1", cmbRequired, cDuration:=cmb3Seconds)
          Case Else
            If Not .ErrorMsg = VBA.vbNullString Then Call CMB.CMsgBox(.ErrorMsg, STRAPPTITLE, "Error", cmbError, cmbClose, FRM): Let View_core_error_handler = True
        End Select

      Case "balmet"
        Select Case .ErrorNum
          Case 0.1: Call RES.FrmCMsgbox(.ErrorMsg, STRAPPTITLE, vbInformation, True, FRM.Controls("Txtbal3"))
          Case 1.2: Call RES.FrmCMsgbox(.ErrorMsg, STRAPPTITLE, vbInformation, True, FRM.Controls("Txtbal2"))
          Case 1.11: Call RES.FrmCMsgbox(.ErrorMsg, STRAPPTITLE, vbInformation, True, FRM.Controls("Txtbal0"))
          Case Else
            If Not .ErrorMsg = VBA.vbNullString Then Call RES.FrmCMsgbox(.ErrorMsg, STRAPPTITLE, vbInformation)
        End Select

      Case Else
        If Not .ErrorMsg = VBA.vbNullString Then Call VBA.MsgBox(.ErrorMsg, vbInformation, STRAPPTITLE)

    End Select
  End With

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function



' CORE METHODS LIST (PUBLIC) ------------------------------------------------------- '
Public Sub ViewUI(ByVal ControlId As String)

  Dim oFrm As New FRMBasics

  #If Debugging Then
    Call VBA.MsgBox("VIEW: " & ControlId)
  #End If

  On Error GoTo EH
  Call View_assets_initialize_model

  Select Case ControlId

    Case "regression": Call View_form_regression(ControlId, oFrm)
    Case "regressionselection": Call View_form_regressionselection(ControlId)

    Case "balmet": Call View_form_balmet(ControlId, oFrm)
    Case "balmetselection": Call View_form_balmetselection(ControlId)
    Case "balmeteco": Call View_form_balmeteco(ControlId, oFrm)

    Case "projection": Call View_form_projection(ControlId, oFrm) 'Call ProjectionsTest

    Case "oreinvoice": Call View_form_purchasemineral(ControlId, oFrm)

    Case "concentrateinvoice": Call View_form_purchaseconcentrate(ControlId, oFrm)

    Case "configlauncher": Call View_form_config(ControlId, oFrm)

    'Help, About, License
    Case "helprlselection", "helpbalmetselection", "helpudf": Call View_Core_execution("hlp" & VBA.Mid(ControlId, 5), Nothing, VStatusMsg:="Ayuda - " & STRAPPTITLE, VBooHelp:=True)
    Case "help": Call View_form_help
    Case "about", "license": Call View_form_about(ControlId, oFrm, (ControlId = "license"))

    Case Else: GoTo EH

  End Select

EH:
  Set oFrm = Nothing
  Call View_assets_controls_free
  Call RES.ErrorHandler.ErrorHandlerDisplay("VIEW::UI")

End Sub
' CORE METHODS LIST (PUBLIC) ------------------------------------------------------- '
' ======================== CLASS EXCLIQ LITE VIEW STRUCTURE ======================== '


