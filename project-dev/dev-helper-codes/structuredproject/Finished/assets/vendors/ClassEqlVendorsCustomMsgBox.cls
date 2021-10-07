VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassEqlVendorsCustomMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'============================= CLASS MODULES STRUCTURE =============================='
'=============================== CREDITS AND LICENSE ================================'
'
' MIT License
'
' Copyright (c) 2021 Roberto Carlos Romay Medina (Roccou)
' CustomMsgboxClass, Excel simple message box for use with macros on any Form.
' Version: 2.1.1
' GitHub: https://github.com/roccouu/
' Potosí - Bolívia
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
'=============================== CREDITS AND LICENSE ================================'



'============================= CLASS MODULES STRUCTURE =============================='
' CUSTOM EVENTS ---------------------------------------------------------------------'
' OBJECT VARIABLES (GLOOBJ_) --------------------------------------------------------'
' VARIABLES (GLOStr_) ---------------------------------------------------------------'
' VARIABLES CUSTOM ------------------------------------------------------------------'
' CONSTANTS (GLOSTR_) ---------------------------------------------------------------'
' CONSTANTS ERROR MESSAGES (GLOSTR_ERR_) --------------------------------------------'
' PROPERTIES ------------------------------------------------------------------------'
' IsActive
' CONSTRUCTOR/DESTRUCTOR ------------------------------------------------------------'
' CONTROLS EVENTS -------------------------------------------------------------------'
' CORE METHODS LIST PRIVATE ---------------------------------------------------------'
' CMsgbox_core_open
' CMsgbox_core_close
' CMsgbox_show
' CMsgbox_set_structure
' CMsgbox_set_layout
' CMsgbox_set_alignment
' CMsgbox_set_values
' CMsgbox_set_visibility
' CMsgbox_set_theme_colors
' CMsgbox_set_required_theme_color_to_target_control
' CMsgbox_restore_required_target_control
' CMsgbox_set_timer_and_required_control
' ASSETS ----------------------------------------------------------------------------'
' CMsgbox_asset_set_colors_pallete
' CMsgbox_asset_get_status_values
' CMsgbox_asset_get_align_horizontal
' CMsgbox_asset_get_align_vertical
' CMsgbox_asset_get_frm_has_control
' CMsgbox_asset_stop_frm_cancel_default_buttons
' CMsgbox_asset_restore_frm_cancel_default_buttons
' CMsgbox_asset_exists_frm
' CMsgbox_asset_required_control_setfocus
' ERRORS MANNAGEMENT AND ASSERTSIONS ------------------------------------------------'
' Debug_assert
' PUBLIC METHODS LIST (INTERFACE) ---------------------------------------------------'
' CMsgBox
'============================= CLASS MODULES STRUCTURE =============================='





'============================= CLASS MODULES STRUCTURE =============================='
' CUSTOM EVENTS ---------------------------------------------------------------------'
' Requester: could be the name of any control, Request: any value to send, Response: any value to receive, and Cancel maby to abort event
Public Event CMsgBoxEvent(ByVal CStrReqstr As String, ByVal CVarReq As Variant, ByRef CVarRes As Variant, ByRef CBooCancel As Boolean)

' OBJECT VARIABLES (GLOOBJ_) --------------------------------------------------------'
Private FmeMsgbox As MSForms.Frame
Private WithEvents BtnMsgBox0 As MSForms.CommandButton
Attribute BtnMsgBox0.VB_VarHelpID = -1
Private WithEvents LblMsgboxTitle As MSForms.Label
Attribute LblMsgboxTitle.VB_VarHelpID = -1
Private WithEvents LblMsgboxClose As MSForms.Label
Attribute LblMsgboxClose.VB_VarHelpID = -1

' VARIABLES (GLOStr_) ---------------------------------------------------------------'
Private GLOStr_ParentCancelButton As String
Private GLOStr_ParentDefaultButton As String
Private GLOBoo_IsActive As Boolean

' VARIABLES CUSTOM ------------------------------------------------------------------'
Public Enum CMB_MsgType
  cmbDefault
  cmbError
  cmbSuccess
  cmbInfo
  cmbQuestion
  cmbAlert
End Enum

Public Enum CMB_ControlRequired
  cmbNoRequired
  cmbRequired
End Enum

Public Enum CMB_ControlPositionX
  cmbCenter
  cmbLeft
  cmbRight
End Enum

Public Enum CMB_ControlPositionY
  cmbMiddle
  cmbTop
  cmbBottom
End Enum

Public Enum CMB_MsgDuration
  cmbNoTime = 2
  cmb3Seconds
  cmb4Seconds
  cmb5Seconds ' ...
End Enum

Public Enum CMB_ControlButtons
  cmbNoButtons
  cmbOKOnly
  cmbClose
  cmbCancel
  ' cmbOKCancel ' To support asyncronous capture events this class could be setted inside an independent UserFuorm (more complexity for class mannagement)
  ' cmbYesNo
End Enum

Public Enum CMB_MsgBoxAspect
  cmbPopupDefault
  cmbFluid
  cmbTagLine
End Enum

Private Type CMB_Typ_Parameters
  cmbStrMsg As String
  cmbStrTitle As String
  cmbStrSubtitle As String
  cmbEnuMsgType As CMB_MsgType
  cmbEnuMsgButtons As CMB_ControlButtons
  cmbUFrm As MSForms.UserForm
  cmbStrCtrlTargetName As String
  cmbStrCtrlSourceName As String
  cmbEnuCtrlRequired As CMB_ControlRequired
  cmbEnuPositionX As CMB_ControlPositionX
  cmbEnuPositionY As CMB_ControlPositionY
  cmbThmColor As XlRgbColor
  cmbEnuDuration As CMB_MsgDuration
  cmbEnuAspect As CMB_MsgBoxAspect
  cmbVarEnabling As Variant
End Type

Private Type CMB_Typ_ColorPallete
  cmbBackground As Long
  cmbBgFontcolor As Long
  cmbFontcolor As Long
  cmbErrorMaincolor As XlRgbColor
  cmbSuccessMaincolor As XlRgbColor
  cmbInfoMaincolor As XlRgbColor
  cmbQuestionMaincolor As XlRgbColor
  cmbAlertMaincolor As XlRgbColor
  cmbDefaultMaincolor As XlRgbColor
  cmbTempCtrlBackColor As XlRgbColor
  cmbErrorSoftcolor As XlRgbColor
  cmbSuccessSoftcolor As XlRgbColor
  cmbInfoSoftcolor As XlRgbColor
  cmbQuestionSoftcolor As XlRgbColor
  cmbAlertSoftcolor As XlRgbColor
  cmbDefaultSoftcolor As XlRgbColor
  cmbTempCtrlColor As XlRgbColor
End Type

Private cColor As CMB_Typ_ColorPallete
Private cParam As CMB_Typ_Parameters

' CONSTANTS (GLOSTR_) ---------------------------------------------------------------'

' CONSTANTS ERROR MESSAGES (GLOSTR_ERR_) --------------------------------------------'
Private Const CUSTOM_ERROR As Long = VBA.vbObjectError + 516





' PROPERTIES ------------------------------------------------------------------------'
Public Property Get IsActive() As Boolean

  Let IsActive = GLOBoo_IsActive

End Property




' CONSTRUCTOR/DESTRUCTOR ------------------------------------------------------------'
Private Sub Class_Terminate()

  Call CMsgbox_core_close

End Sub





' CONTROLS EVENTS -------------------------------------------------------------------'
' Buttons
Private Sub BtnMsgBox0_Click()

  Dim VarReq As Variant, VarRes As Variant
  Dim StrReqstr As String
  Dim BooCancel As Boolean

  RaiseEvent CMsgBoxEvent(StrReqstr, VarReq, VarRes, BooCancel)

  If BooCancel Then
    ' Do any stuff here
  End If

  Call CMsgbox_core_close

End Sub

' Labels
Private Sub LblMsgBoxClose_Click()

  Call CMsgbox_core_close

End Sub

Private Sub LblMsgBoxClose_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)

  Let LblMsgboxClose.BackColor = rgbTomato
  Let LblMsgboxClose.BorderColor = VBA.vbRed

End Sub

Private Sub LblMsgboxTitle_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)

  Let LblMsgboxClose.BackColor = cColor.cmbFontcolor
  Let LblMsgboxClose.BorderColor = rgbAntiqueWhite

End Sub





' CORE METHODS LIST PRIVATE ---------------------------------------------------------'
Private Sub CMsgbox_core_open( _
  ByVal MStrMsg As String, _
  ByVal MStrTitle As String, _
  ByVal MStrSubtitle As String, _
  ByVal MEnuMsgType As CMB_MsgType, _
  ByVal MEnuMsgButtons As CMB_ControlButtons, _
  ByVal UFrm As MSForms.UserForm, _
  ByVal MStrCtrlTargetName As String, _
  ByVal MStrCtrlSourceName As String, _
  ByVal MEnuCtrlRequired As CMB_ControlRequired, _
  ByVal MEnuPositionX As CMB_ControlPositionX, _
  ByVal MEnuPositionY As CMB_ControlPositionY, _
  ByVal MThmColor As XlRgbColor, _
  ByVal MEnuDuration As CMB_MsgDuration, _
  ByVal MEnuAspect As CMB_MsgBoxAspect)

  On Error GoTo EH

  With cParam
    Let .cmbStrMsg = MStrMsg
    Let .cmbStrTitle = MStrTitle
    Let .cmbStrSubtitle = MStrSubtitle
    Let .cmbEnuMsgType = MEnuMsgType
    Let .cmbEnuMsgButtons = MEnuMsgButtons
    Set .cmbUFrm = UFrm
    Let .cmbStrCtrlTargetName = MStrCtrlTargetName
    Let .cmbStrCtrlSourceName = MStrCtrlSourceName
    Let .cmbEnuCtrlRequired = MEnuCtrlRequired
    Let .cmbEnuPositionX = MEnuPositionX
    Let .cmbEnuPositionY = MEnuPositionY
    Let .cmbThmColor = MThmColor
    Let .cmbEnuDuration = MEnuDuration
    Let .cmbEnuAspect = MEnuAspect
  End With

  ' Validate Objects
  If Not CMsgbox_asset_exists_frm() Then Call VBA.Err.Raise(CUSTOM_ERROR, "CMsgBox::-core_open", "Se requiere un objeto UserForm")

  ' 1. Initialize colors pallete
  Call CMsgbox_asset_set_colors_pallete ' Set MsgBox colors

  ' 2. Build CMsgBox
  Call CMsgbox_set_structure ' Build CMsgBox

  ' 3. Show CMsgBox
  Call CMsgbox_show

  Call UFrm.Repaint ' Obvious, to refresh form

EH:
  If Not VBA.Err.Number = 0 Then
    Call VBA.MsgBox("CMB Error: " & VBA.Err.Description, vbInformation, "CMB")
    Call VBA.Err.Clear
  End If

End Sub

Private Sub CMsgbox_core_close()

  Dim BooExists As Boolean

  On Error GoTo EH
  If Not FmeMsgbox Is Nothing Then

    Let BooExists = CMsgbox_asset_get_frm_has_control(cParam.cmbUFrm, "FrameMsgBoxObject")

    ' Restore original colors to target control
    Call CMsgbox_restore_required_target_control

    ' Restore original enabled starus in source control(s)' AUGMENTED
    Call CMsgbox_restore_enable_source_control

    ' Restore Cancel and  Default button in main UserForm interface
    Call CMsgbox_asset_restore_frm_cancel_default_buttons

    ' Set focus to control target
    Call CMsgbox_asset_required_control_setfocus

    ' Removing and eliminating CMsgBox
    With cParam.cmbUFrm
      If BooExists Then
        Let .Controls("FrameMsgBoxObject").Visible = False
        Call .Controls.Remove("FrameMsgBoxObject")
      End If

      #If Debugging Then
        Dim cC As Object
        For Each cC In .Controls
          Debug.Print cC.Name
        Next cC
        Set cC = Nothing
      #End If
    End With

    Set cParam.cmbUFrm = Nothing

  End If

  Let GLOBoo_IsActive = False

EH:
  Set FmeMsgbox = Nothing
  Set BtnMsgBox0 = Nothing
  Set LblMsgboxTitle = Nothing
  Set LblMsgboxClose = Nothing

End Sub

Private Sub CMsgbox_show()

  ' When user send another UserForm, check if this new form already has a CMsgBox instance, else create new instance
  On Error GoTo EH

  'Let FmeMsgbox.BackColor = VBA.vbRed ' To debbuging

  ' Set CMsgBox Values
  Call CMsgbox_set_values

  ' Set CMsgBox Layout
  Call CMsgbox_set_layout

  ' Set CMsgBox controls Visibility
  Call CMsgbox_set_visibility

  ' Set CMsgBox Alignment
  Call CMsgbox_set_alignment

  ' Set CMsgBox Theme color
  Call CMsgbox_set_theme_colors

  ' Show CMsgBox
  Let FmeMsgbox.Visible = True

  ' Set Target control theme <<<====!
  Call CMsgbox_set_required_theme_color_to_target_control

  ' Disable Source control(s)
  Call CMsgbox_disable_source_control

  ' Now CMsgBox is active
  Let GLOBoo_IsActive = True

  ' Set CMsgBox timer
  If Not cParam.cmbEnuDuration = cmbNoTime Then Call CMsgbox_set_timer_and_required_control

EH:

End Sub

Private Sub CMsgbox_set_structure()

  'BUILDS A FRAME WITH ALL CONTROLS FOR CUSTOM MSGBOX
  ' [x]  A Frame main wrapper object.
  ' [x]  A Frame as structure holder.
  ' [x]  A Header Label for Title.
  ' [x]  A Label for Subtitle.
  ' [x]  A Label multirow for Content message.
  ' [x]  A Label for Close button.
  ' [x]  A Label inside frame close button for content text x.

  Const BytCloseButtonSize As Byte = 10
  Const SngIconFontSize As Single = 1.5

  ' Build CMsgBox
  On Error GoTo EH
  With cParam.cmbUFrm

    ' Prepare main object
    If CMsgbox_asset_get_frm_has_control(cParam.cmbUFrm, "FrameMsgBoxObject") Then
      Let .Controls("FrameMsgBoxObject").Visible = False
      Call .Controls.Remove("FrameMsgBoxObject")
    End If
    If Not FmeMsgbox Is Nothing Then Set FmeMsgbox = Nothing

    ' FmeMsgbox is the global object that will work as CMsgBox
    Set FmeMsgbox = .Controls.Add("Forms.Frame.1", "FrameMsgBoxObject", False)

    With FmeMsgbox

      Call .Controls.Add("Forms.Frame.1", "FrameMsgBoxContainer")
      With .Controls("FrameMsgBoxContainer")

        Set LblMsgboxTitle = .Controls.Add("Forms.Label.1", "LblMsgboxTitle")
        Call .Controls.Add("Forms.Label.1", "LblMsgboxIcon")
        Call .Controls.Add("Forms.Label.1", "LblMsgboxSubtitle")
        Call .Controls.Add("Forms.Label.1", "LblMsgboxContent")
        Set LblMsgboxClose = .Controls.Add("Forms.Label.1", "LblMsgboxClose")
        If cParam.cmbEnuDuration = cmbNoTime Then Set BtnMsgBox0 = .Controls.Add("Forms.CommandButton.1", "BtnMsgboxA")

      End With
      Call .ZOrder(0)

    End With

  End With

EH:

End Sub

Private Sub CMsgbox_set_layout()

  ' Configure th CMsgBox layout with MEnuAspect as base
  ' 1. Set dimensions
  ' 2. Set hide or unhide controls

  Dim LngWidth As Long

  Const LngTagLineMinWidth As Long = 100
  Const LngLeft As Long = 10

  On Error GoTo EH

  With FmeMsgbox.Parent ' Frame cMsgBox Object always will have a UserForm as Parent

    ' Set CMsgBox width (Containers)
    With FmeMsgbox

      Let LngWidth = (.Parent.InsideWidth) * IIf(.Parent.InsideWidth <= 200, 1, IIf(.Parent.InsideWidth > 400, 0.5, 0.85))
      Select Case cParam.cmbEnuAspect

        Case cmbFluid
          Let .Width = .Parent.Width * 1.1
          Let .Controls("FrameMsgBoxContainer").Width = LngWidth

        Case cmbTagLine
          Let LngWidth = LngTagLineMinWidth
          If CMsgbox_asset_get_frm_has_control(.Parent, cParam.cmbStrCtrlTargetName) Then
            'If .Parent.Controls(cParam.cmbStrCtrlTargetName).Width > .Width Then
            If .Parent.Controls(cParam.cmbStrCtrlTargetName).Width > LngWidth Then
              Let LngWidth = .Parent.Controls(cParam.cmbStrCtrlTargetName).Width
              'Let .Width = LngWidth
            End If
          'Else
            'Let LngWidth = LngTagLineMinWidth
            'Let .Width = LngWidth
          End If
          Let .Width = LngWidth
          Let .Controls("FrameMsgBoxContainer").Width = .Width * 1.1

        Case Else ' cmbPopupDefault
          Let .Width = LngWidth ' Fix defined CMsgBox area
          Let .Controls("FrameMsgBoxContainer").Width = LngWidth * 1.1
          If cParam.cmbEnuDuration = cmbNoTime Then Let .Controls("BtnMsgboxA").Height = .Controls("BtnMsgboxA").Height * 0.8

      End Select

    End With

    ' Set dimensions to CMsgBox controls
    With .Controls("LblMsgboxTitle") ' Title
      Let .FontSize = .FontSize * 1.1
      Let .FontBold = True
      Let .Width = .Parent.Width
      Let .Height = .FontSize * 1.7
    End With

    With LblMsgboxClose ' Close Button
      Let .Font.Size = .Font.Size * 0.8
      Let .Height = .Parent.Controls("LblMsgboxTitle").Height * 0.6
      Let .Width = .Height
      Call .ZOrder(0)
    End With

    With .Controls("LblMsgboxIcon") ' Icon
      Let .Font.Size = .Font.Size * IIf(cParam.cmbEnuAspect = cmbTagLine, 1.1, 2)
      Let .Font.Bold = True
      Let .AutoSize = True
      Let .WordWrap = False
      Let .Height = .FontSize * 1.25
      Let .Width = .Height
    End With

    With .Controls("LblMsgboxSubtitle") ' Subtitle
      Let .AutoSize = True
      Let .WordWrap = False
      Let .Font.Bold = True
      Let .Width = .Parent.InsideWidth - ((LngLeft * 4) + IIf(cParam.cmbEnuMsgType = cmbDefault, 0, .Parent.Controls("LblMsgboxIcon").Width * 1.5))
    End With

    With .Controls("LblMsgboxContent") ' Content
      Let .Font.Size = .Font.Size * 1.1

      Let LngWidth = LngWidth - ((LngLeft * 1.5) + IIf(cParam.cmbEnuMsgType = cmbDefault, 0, .Parent.Controls("LblMsgboxIcon").Width + 5) + _
      IIf(cParam.cmbEnuDuration = cmbNoTime, .Parent.Controls("LblMsgboxClose").Width, 0))

      Let .Width = LngWidth '.Parent.InsideWidth - ((LngLeft * IIf(cParam.cmbEnuAspect = cmbTagLine, 3, 4)) + IIf(cParam.cmbEnuMsgType = cmbDefault, 0, .Parent.Controls("LblMsgboxIcon").Width + 5) + _
      IIf(cParam.cmbEnuDuration = cmbNoTime, .Parent.Controls("LblMsgboxClose").Width, 0))
      Let .WordWrap = True
      Let .Caption = cParam.cmbStrMsg ' Value is here exceptionally. MS issues.
      Let .AutoSize = True
      Let .Width = LngWidth '.Parent.InsideWidth - ((LngLeft * IIf(cParam.cmbEnuAspect = cmbTagLine, 3, 4)) + IIf(cParam.cmbEnuMsgType = cmbDefault, 0, .Parent.Controls("LblMsgboxIcon").Width + 5) + _
      IIf(cParam.cmbEnuDuration = cmbNoTime, .Parent.Controls("LblMsgboxClose").Width, 0))
    End With

  End With

EH:

End Sub

Private Sub CMsgbox_set_alignment()

  ' Configure the CMsgBox alignment respect its Parent object and/or respect its TargetControl object, with MEnuAspect as base
  ' 1. Set Horizontal align
  ' 2. Set Vertical align
  ' 3. [Set alignment respect control target]

  Dim xCenter As Long, xMiddle As Long, xRight As Long, xBottom As Long, xHeight As Long

  Const sFN As String = "CMsgbox::-Set_alignment"
  Const LngLeft As Long = 10

  On Error GoTo EH

  With FmeMsgbox.Parent ' Frame cMsgBox Object always will have a UserForm as Parent

    ' Set CMsgBox height/width
    Let xHeight = .Controls("LblMsgboxTitle").Height + .Controls("LblMsgboxSubtitle").Height + .Controls("LblMsgboxContent").Height + (LngLeft * 2.5)
    If cParam.cmbEnuDuration = cmbNoTime Then Let xHeight = xHeight + (.Controls("BtnMsgboxA").Height * IIf(cParam.cmbEnuAspect = cmbFluid, 1, 1.7))

    With FmeMsgbox

      ' CMsgBox Height adjustment
      Select Case cParam.cmbEnuAspect

        Case cmbFluid
          Let .Height = xHeight
          Let xMiddle = (((.Parent.InsideHeight) * 0.5) - (.Height * 0.5))
          Let xBottom = ((.Parent.InsideHeight) - (.Height * 1.1))
          Let .Top = IIf(cParam.cmbEnuPositionY = cmbTop, 10, IIf(cParam.cmbEnuPositionY = cmbMiddle, xMiddle, xBottom))
          Let .Left = -0.5
          With .Controls("FrameMsgBoxContainer")
            Let .Height = .Height * 1.1
            Let .Left = (.Parent.InsideWidth * 0.43) - (.InsideWidth * 0.5)
          End With

        Case cmbTagLine ' Adjust both containers at main content
          Let xCenter = .Parent.Controls("LblMsgboxContent").Height
          Let .Height = .Parent.Controls("LblMsgboxContent").Height * IIf(xCenter < (.Parent.Controls("LblMsgboxContent").FontSize * 2), 1.5, 1.2)
          If CMsgbox_asset_get_frm_has_control(.Parent, cParam.cmbStrCtrlTargetName) Then
            Let .Left = CMsgbox_asset_get_align_horizontal(.Parent.Controls(cParam.cmbStrCtrlTargetName), cParam.cmbEnuPositionX)
            Let .Top = CMsgbox_asset_get_align_vertical(.Parent.Controls(cParam.cmbStrCtrlTargetName), cParam.cmbEnuPositionY)
          Else
            Let .Left = ((.Parent.InsideWidth) * 0.5) - (.Width * 0.5)
            Let .Top = ((.Parent.InsideHeight * 0.95) * 0.5) - (.Height * 0.5)
          End If

          ' Adjust controls positions & Readjust controls dimensions
          With .Controls("FrameMsgBoxContainer")
            Let .Width = .Parent.Width * 1.1
            Let .Height = .Parent.Height * 1.1
            Let .Left = -0.5
            Let .Top = -0.5
          End With

          With .Controls("LblMsgboxIcon")
            Let .Left = 5
            Let .Top = 2
          End With

          With .Controls("LblMsgboxContent")
            'Let .Width = .Parent.InsideWidth - (IIf(cParam.cmbEnuMsgType = cmbDefault, 0, .Parent.Controls("LblMsgboxIcon").Width + 5) + IIf(cParam.cmbEnuDuration = cmbNoTime, .Parent.Controls("LblMsgboxClose").Width + 5, 0) + 10) - IIf(cParam.cmbEnuAspect = cmbTagLine, 5, 0)
            Let .Left = IIf(cParam.cmbEnuMsgType = cmbDefault, 0, .Parent.Controls("LblMsgboxIcon").Left + .Parent.Controls("LblMsgboxIcon").Width) + 5
            Let .Top = 2
          End With

          With .Controls("LblMsgboxClose")
            Let .Left = .Parent.Controls("LblMsgboxContent").Left + .Parent.Controls("LblMsgboxContent").Width + 3
            Let .Top = .Parent.Controls("LblMsgboxContent").Top
            Let .TextAlign = fmTextAlignCenter
          End With

          With .Controls("LblMsgboxTitle")
            'Let .Width = FmeMsgbox.Width
            Let .Height = FmeMsgbox.Height
            Let .Left = 0
            Let .Top = 0
          End With

          If cParam.cmbEnuDuration = cmbNoTime Then Let .Controls("BtnMsgboxA").Top = -100

        Case Else ' cmbPopupDefault
          Let .Height = xHeight
          Let xCenter = (((.Parent.InsideWidth) * 0.5) - (.Width * 0.5))
          Let xMiddle = (((.Parent.InsideHeight * 0.9) * 0.5) - (.Height * 0.5))
          Let xBottom = ((.Parent.InsideHeight * 0.95) - .Height)
          Let xRight = ((.Parent.InsideWidth) - .Width - 10)
          Let .Left = IIf(cParam.cmbEnuPositionX = cmbLeft, 10, IIf(cParam.cmbEnuPositionX = cmbCenter, xCenter, xRight))
          Let .Top = IIf(cParam.cmbEnuPositionY = cmbTop, 10, IIf(cParam.cmbEnuPositionY = cmbMiddle, xMiddle, xBottom))

          Let .Controls("FrameMsgBoxContainer").Height = .Height * 1.1
          If cParam.cmbEnuDuration = cmbNoTime Then Let .Controls("BtnMsgboxA").Left = .Parent.Left - 100

      End Select

      If Not cParam.cmbEnuAspect = cmbTagLine Then

        ' Adjust controls positions
        With .Controls("FrameMsgBoxContainer") ' Container
          If Not cParam.cmbEnuAspect = cmbFluid Then Let .Left = -0.5
          Let .Top = -0.5
        End With

        With .Controls("LblMsgboxTitle") ' Title
          Let .Left = 0
          Let .Top = 0
        End With

        With .Controls("LblMsgboxClose") ' Close button
          Let .Left = .Parent.InsideWidth - (.Width + (LngLeft * IIf(cParam.cmbEnuAspect = cmbFluid, 1, 3)))
          Let .Top = 2
          Let .TextAlign = fmTextAlignCenter
        End With

        With .Controls("LblMsgboxIcon") ' Icon
          Let .Left = LngLeft
          Let .Top = .Parent.Controls("LblMsgboxTitle").Top + .Parent.Controls("LblMsgboxTitle").Height + LngLeft
        End With

        With .Controls("LblMsgboxSubtitle") ' Subtitle
          Let .Left = IIf(cParam.cmbEnuMsgType = cmbDefault, LngLeft, .Parent.Controls("LblMsgboxIcon").Left + .Parent.Controls("LblMsgboxIcon").Width + 5)
          Let .Top = .Parent.Controls("LblMsgboxIcon").Top
        End With

        With .Controls("LblMsgboxContent") ' Content
          Let .Left = IIf(cParam.cmbEnuMsgType = cmbDefault, LngLeft, .Parent.Controls("LblMsgboxIcon").Left + .Parent.Controls("LblMsgboxIcon").Width + 5)
          Let .Top = IIf(cParam.cmbStrSubtitle = VBA.vbNullString, .Parent.Controls("LblMsgboxSubtitle").Top, .Parent.Controls("LblMsgboxSubtitle").Top + .Parent.Controls("LblMsgboxSubtitle").Height + 5)
        End With

        If cParam.cmbEnuDuration = cmbNoTime Then
          With .Controls("BtnMsgboxA") ' Button
            Let .Width = .Width * 0.8
            Let .Height = (.Parent.Controls("LblMsgboxTitle").Height * 1.3)
            Let .Left = (.Parent.Controls("LblMsgboxContent").Left + .Parent.Controls("LblMsgboxContent").Width) - .Width ' .Parent.Left - 100
            Let .Top = IIf(cParam.cmbEnuAspect = cmbFluid, (-.Parent.Height * 2), (.Parent.InsideHeight - (.Height + (LngLeft * 2))))
          End With
        End If

      End If
      If (.Parent.Height * 1.15) < .Height Or (.Parent.Width * 1.15) < .Width Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "Height and width of Form do not can be less than height or width of CMsgBox.")

    End With

  End With

EH:

End Sub

Private Sub CMsgbox_set_values()

  ' Configure the CMsgBox values and its controls values with received parameters as base
  Dim StrMsgType As String
  Dim LngHeight As Long

  Const cMsgboxLeft As Byte = 10
  Const cMsgboxTop As Byte = 10
  Const BytFontSize As Byte = 1.2

  On Error GoTo EH
  ' Get Icon values
  Let StrMsgType = CMsgbox_asset_get_status_values()
  With cParam.cmbUFrm

    With .Controls("LblMsgboxTitle") ' Title
      Let .Caption = IIf(cParam.cmbEnuAspect = cmbTagLine, VBA.vbNullString, VBA.Space$(4) & cParam.cmbStrTitle)
    End With

    With LblMsgboxClose ' Close button
      Let .Caption = VBA.ChrW(&H2716)
      Let .ControlTipText = "Cerrar"
    End With

    With .Controls("LblMsgboxIcon") ' Icon
      Let .Caption = StrMsgType
    End With

    With .Controls("LblMsgboxSubtitle")
      Let .Caption = IIf(cParam.cmbEnuAspect = cmbTagLine, VBA.vbNullString, cParam.cmbStrSubtitle)
    End With

    ' Stop Cancel and Default UserForm buttons
    Call CMsgbox_asset_stop_frm_cancel_default_buttons

    If cParam.cmbEnuDuration = cmbNoTime Then
      With .Controls("BtnMsgboxA")
        Let .Caption = IIf(cParam.cmbEnuMsgButtons = cmbOKOnly, "Aceptar", IIf(cParam.cmbEnuMsgButtons = cmbCancel, "Cancelar", "Cerrar"))
        Let .Accelerator = "a"
        Let .Cancel = True
      End With
    End If

  End With

EH:

End Sub

Private Sub CMsgbox_set_visibility()

  ' Configure the CMsgBox visibility properties
  On Error GoTo EH

  With cParam.cmbUFrm

    Let .Controls("LblMsgboxTitle").Visible = True
    Let .Controls("LblMsgboxClose").Visible = (cParam.cmbEnuDuration = cmbNoTime)
    Let .Controls("LblMsgboxIcon").Visible = Not (cParam.cmbEnuMsgType = cmbDefault)
    Let .Controls("LblMsgboxSubtitle").Visible = True
    Let .Controls("LblMsgboxContent").Visible = True
    If cParam.cmbEnuDuration = cmbNoTime Then Let .Controls("BtnMsgboxA").Visible = Not (cParam.cmbEnuMsgButtons = cmbNoButtons)

  End With

EH:

End Sub

Private Sub CMsgbox_set_theme_colors()

  ' Configure the CMsgBox theme color for its controls
  Dim cMsgBoxIconColor As Long ', cMsgVBOrange As Long

  On Error GoTo EH

  ' Determine icon color
  With cColor
    Select Case cParam.cmbEnuMsgType
      Case cmbError: cMsgBoxIconColor = .cmbErrorMaincolor ' VBA.vbRed
      Case cmbSuccess: cMsgBoxIconColor = .cmbSuccessMaincolor 'VBA.vbGreen
      Case cmbInfo: cMsgBoxIconColor = .cmbInfoMaincolor 'VBA.vbBlue
      Case cmbQuestion: cMsgBoxIconColor = .cmbQuestionMaincolor 'cMsgVBOrange
      Case cmbAlert: cMsgBoxIconColor = .cmbAlertMaincolor 'VBA.vbCyan
      Case Else: cMsgBoxIconColor = .cmbDefaultMaincolor 'VBA.vbBlack ' cmbDefault
    End Select
  End With

  ' Frame cMsgBox Object always will have a UserForm as Parent
  With cParam.cmbUFrm

    With FmeMsgbox
      Let .BorderStyle = fmBorderStyleSingle
      Let .SpecialEffect = fmSpecialEffectFlat
      Let .ScrollBars = fmScrollBarsNone
      Let .BorderColor = cColor.cmbBackground
      Let .BackColor = .BackColor
      Let .Controls("FrameMsgBoxContainer").BorderStyle = fmBorderStyleNone
      Let .Controls("FrameMsgBoxContainer").SpecialEffect = fmSpecialEffectFlat
      Let .Controls("FrameMsgBoxContainer").ScrollBars = fmScrollBarsNone
    End With

    With .Controls("LblMsgboxTitle")
      Let .BorderStyle = IIf(cParam.cmbEnuAspect = cmbPopupDefault, fmBorderStyleSingle, fmBorderStyleNone)
      Let .BackStyle = IIf(cParam.cmbEnuAspect = cmbTagLine, fmBackStyleTransparent, fmBackStyleOpaque)
      Let .BackColor = cColor.cmbBackground
      Let .ForeColor = cColor.cmbBgFontcolor
    End With

    With LblMsgboxClose
      Let .BorderStyle = fmBorderStyleSingle
      Let .BackStyle = fmBackStyleOpaque
      Let .ForeColor = rgbAntiqueWhite
      Let .BackColor = cColor.cmbFontcolor
      Let .BorderColor = rgbAntiqueWhite
    End With

    With .Controls("LblMsgboxIcon") ' Icon
      'Let .BackColor = VBA.vbRed 'cMsgboxFontcolor' To debbug
      Let .ForeColor = cMsgBoxIconColor
      Let .BackStyle = fmBackStyleTransparent
    End With

    With .Controls("LblMsgboxSubtitle")
      'Let .BackColor = VBA.vbRed 'cMsgboxFontcolor' To debbug
      Let .BackStyle = fmBackStyleTransparent
      Let .ForeColor = cColor.cmbFontcolor
    End With

    With .Controls("LblMsgboxContent")
      'Let .BackColor = VBA.vbGreen ' To debbug
      Let .BackStyle = fmBackStyleTransparent
      Let .ForeColor = cColor.cmbFontcolor
    End With

    If cParam.cmbEnuDuration = cmbNoTime Then
      With .Controls("BtnMsgboxA")
        Let .ForeColor = cColor.cmbFontcolor
      End With
    End If

    Select Case cParam.cmbEnuAspect

      Case cmbFluid
        Let .Controls("FrameMsgBoxContainer").BorderStyle = fmBorderStyleNone
        Let .Controls("FrameMsgBoxContainer").SpecialEffect = fmSpecialEffectFlat
        Let .Controls("FrameMsgBoxContainer").BackColor = .Controls("FrameMsgBoxContainer").Parent.BackColor
        Let .Controls("FrameMsgBoxContainer").BorderColor = .Controls("FrameMsgBoxContainer").Parent.BackColor

      Case cmbTagLine ' Adjust both containers at just main content
        Let FmeMsgbox.BorderStyle = fmBorderStyleSingle

      Case Else ' cmbPopupDefault|cmbPopupAutoclose
        Let FmeMsgbox.BorderStyle = fmBorderStyleSingle
        Let .Controls("FrameMsgBoxContainer").BorderStyle = fmBorderStyleNone
        Let .Controls("FrameMsgBoxContainer").SpecialEffect = fmSpecialEffectFlat
        Let .Controls("FrameMsgBoxContainer").BackColor = .Controls("FrameMsgBoxContainer").Parent.BackColor
        Let .Controls("FrameMsgBoxContainer").BorderColor = cColor.cmbBackground 'cMsgboxFontcolor

    End Select

  End With

EH:

End Sub

Private Sub CMsgbox_set_required_theme_color_to_target_control()

  ' Configure the CMsgBox theme color for its controls
  ' If Exist MStrCtrlTargetName and its data is required:
  ' 1. [x] Capture MStrCtrlTargetName colors
  ' 2. [x] Change MStrCtrlTargetName colors to warn, success, etc.

  Dim CtrTgtControl As MSForms.Control

  On Error GoTo EH
  ' Work with theme colors
  If cParam.cmbEnuCtrlRequired = cmbRequired Then 'Data control is required
    If CMsgbox_asset_get_frm_has_control(cParam.cmbUFrm, cParam.cmbStrCtrlTargetName) Then ' Do it only if MStrCtrlTargetName exists in UFrm Controls

      Set CtrTgtControl = cParam.cmbUFrm.Controls(cParam.cmbStrCtrlTargetName)
      With CtrTgtControl
        Select Case VBA.TypeName(CtrTgtControl)

          Case "TextBox", "ComboBox", "Label", "ListBox", "Frame", "UserForm"
            Let cColor.cmbTempCtrlBackColor = .BackColor
            Let cColor.cmbTempCtrlColor = .BorderColor
            Select Case cParam.cmbEnuMsgType
              Case cmbError
                Let .BorderColor = cColor.cmbErrorMaincolor
                Let .BackColor = cColor.cmbErrorSoftcolor
              Case cmbSuccess
                Let .BorderColor = cColor.cmbSuccessMaincolor
                Let .BackColor = cColor.cmbSuccessSoftcolor
              Case cmbInfo
                Let .BorderColor = cColor.cmbInfoMaincolor
                Let .BackColor = cColor.cmbInfoSoftcolor
              Case cmbQuestion
                Let .BorderColor = cColor.cmbQuestionMaincolor
                Let .BackColor = cColor.cmbQuestionSoftcolor
              Case cmbAlert
                Let .BorderColor = cColor.cmbAlertMaincolor
                Let .BackColor = cColor.cmbAlertSoftcolor
              Case Else
                Let .BorderColor = cColor.cmbDefaultMaincolor
                Let .BackColor = cColor.cmbDefaultSoftcolor
            End Select

        Case "CommandButton", "ToggleButton", "CheckBox", "OptionButton", "SpinButton", "MultiPage", "SpinButton", "TabStrip"
            Let cColor.cmbTempCtrlBackColor = .BackColor
            Let cColor.cmbTempCtrlColor = .ForeColor
            Select Case cParam.cmbEnuMsgType
              Case cmbError
                Let .ForeColor = cColor.cmbErrorMaincolor
                Let .BackColor = cColor.cmbErrorSoftcolor
              Case cmbSuccess
                Let .ForeColor = cColor.cmbSuccessMaincolor
                Let .BackColor = cColor.cmbSuccessSoftcolor
              Case cmbInfo
                Let .ForeColor = cColor.cmbInfoMaincolor
                Let .BackColor = cColor.cmbInfoSoftcolor
              Case cmbQuestion
                Let .ForeColor = cColor.cmbQuestionMaincolor
                Let .BackColor = cColor.cmbQuestionSoftcolor
              Case cmbAlert
                Let .ForeColor = cColor.cmbAlertMaincolor
                Let .BackColor = cColor.cmbAlertSoftcolor
              Case Else
                Let .ForeColor = cColor.cmbDefaultMaincolor
                Let .BackColor = cColor.cmbDefaultSoftcolor
            End Select

        End Select
      End With

    End If
  End If

EH:
  Set CtrTgtControl = Nothing

End Sub

Private Sub CMsgbox_restore_required_target_control()

  ' Configure the CMsgBox theme color for its controls
  ' If Exist MStrCtrlTargetName and its data is required:
  ' 1. [x] Capture MStrCtrlTargetName colors
  ' 4. [x] Restore MStrCtrlTargetName colors

  Dim MTgtControl As MSForms.Control

  On Error GoTo EH
  ' Restore original colors to control target
  If cParam.cmbEnuCtrlRequired = cmbRequired Then 'Data control is required
    If CMsgbox_asset_get_frm_has_control(cParam.cmbUFrm, cParam.cmbStrCtrlTargetName) Then ' Do it only if MStrCtrlTargetName exists in UFrm Controls

      Set MTgtControl = cParam.cmbUFrm.Controls(cParam.cmbStrCtrlTargetName)
      With MTgtControl
        Select Case VBA.TypeName(MTgtControl)

          Case "TextBox", "ComboBox", "Label", "ListBox", "Frame", "UserForm"
            Let .BackColor = cColor.cmbTempCtrlBackColor
            Let .BorderColor = cColor.cmbTempCtrlColor

          Case "CommandButton", "ToggleButton", "CheckBox", "OptionButton", "SpinButton", "MultiPage", "SpinButton", "TabStrip"
            Let .BackColor = cColor.cmbTempCtrlBackColor
            Let .ForeColor = cColor.cmbTempCtrlColor

        End Select
      End With

    End If

  End If

EH:
  Set MTgtControl = Nothing

End Sub

Private Sub CMsgbox_disable_source_control()

  ' Disable the CMsgBox caller control and other user selected controls
  Dim MSrcControl As MSForms.Control
  Dim VecAux0 As Variant
  Dim StrAux0 As String
  Dim i As Long, j As Long

  On Error GoTo EH
  If cParam.cmbEnuCtrlRequired = cmbRequired Then ' Exist required control
    Let VecAux0 = VBA.Split(cParam.cmbStrCtrlSourceName, ";")
    If VBA.IsArray(VecAux0) Then

      ReDim cParam.cmbVarEnabling(0 To (UBound(VecAux0) - LBound(VecAux0)), 0 To 1) ' Control name, enabled status
      Let j = 0
      For i = LBound(VecAux0) To UBound(VecAux0)
        Let StrAux0 = VBA.CStr(VecAux0(i))
        If CMsgbox_asset_get_frm_has_control(cParam.cmbUFrm, StrAux0) Then ' Do it only if MStrCtrlSourceName exists in UFrm Controls
          Set MSrcControl = cParam.cmbUFrm.Controls(StrAux0)
          With MSrcControl
            Select Case VBA.TypeName(MSrcControl)
              Case "TextBox", "ComboBox", "Label", "ListBox", "Frame", "CommandButton", "ToggleButton", "CheckBox", "OptionButton", "SpinButton", "MultiPage", "SpinButton", "TabStrip"
                Let cParam.cmbVarEnabling(j, 0) = MSrcControl.Name
                Let cParam.cmbVarEnabling(j, 1) = IIf(.Enabled, 1, 0)
                Let .Enabled = False
                Let j = j + 1
            End Select
          End With
        End If
      Next i

    End If
  End If

EH:
  Let VecAux0 = Empty
  Set MSrcControl = Nothing

End Sub

Private Sub CMsgbox_restore_enable_source_control()

  ' Enable the CMsgBox caller control and other user selected controls
  Dim StrAux0 As String
  Dim BooAux0 As Boolean
  Dim i As Long

  On Error GoTo EH
  If cParam.cmbEnuCtrlRequired = cmbRequired Then ' Exist required control
    If VBA.IsArray(cParam.cmbVarEnabling) Then

      For i = LBound(cParam.cmbVarEnabling) To UBound(cParam.cmbVarEnabling)
        Let StrAux0 = VBA.CStr(cParam.cmbVarEnabling(i, 0))
        Let BooAux0 = VBA.CBool(cParam.cmbVarEnabling(i, 1))
        If CMsgbox_asset_get_frm_has_control(cParam.cmbUFrm, StrAux0) Then ' Do it only if MStrCtrlSourceName exists in UFrm Controls
          Let cParam.cmbUFrm.Controls(StrAux0).Enabled = BooAux0
        End If
      Next i

    End If
  End If

EH:
  Let cParam.cmbVarEnabling = Empty

End Sub

Private Sub CMsgbox_set_timer_and_required_control()

  ' Configure the CMsgBox theme color for its controls
  ' If Exist MStrCtrlTargetName and its data is required:
  ' 1. [x] Set timer
  ' If MEnuDuration has a time value
  ' 2. [x] Set timer

'  Dim CtrTgtControl As MSForms.Control, CtrSrcControl As MSForms.Control

  Dim mcTime As String

  On Error GoTo EH
  ' Get MStrCtrlSourceName control
  'If CMsgbox_asset_get_frm_has_control(cParam.cmbUFrm, cParam.cmbStrCtrlSourceName) Then ' Do it only if MStrCtrlSourceName exists in UFrm Controls
  '  Set CtrSrcControl = cParam.cmbUFrm.Controls(cParam.cmbStrCtrlSourceName)
  'End If

  ' Execute timer
  Let mcTime = "00:00:" & IIf(cParam.cmbEnuDuration < 10, "0" & cParam.cmbEnuDuration, "03")
  With FmeMsgbox.Parent

    If cParam.cmbEnuCtrlRequired = cmbRequired Then 'Data control is required
      If CMsgbox_asset_get_frm_has_control(cParam.cmbUFrm, cParam.cmbStrCtrlTargetName) Then ' Do it only if MStrCtrlTargetName exists in UFrm Controls

        'Set CtrTgtControl = .Controls(cParam.cmbStrCtrlTargetName)
        'If Not CtrSrcControl Is Nothing Then Let CtrSrcControl.Left = CtrSrcControl.Left + (.Width * 1.5)

        ' Wait Application by timer
        Call Application.Wait(VBA.Now + VBA.TimeValue(mcTime))

        ' Restore controls values
        'If Not CtrSrcControl Is Nothing Then Let CtrSrcControl.Left = CtrSrcControl.Left - (.Width * 1.5)

      End If

    Else

      'If Not CtrSrcControl Is Nothing Then Let CtrSrcControl.Left = CtrSrcControl.Left + (.Width * 1.5)

      ' Wait Application by timer
      Call Application.Wait(VBA.Now + VBA.TimeValue(mcTime))

      ' Restore controls values
      'If Not CtrSrcControl Is Nothing Then Let CtrSrcControl.Left = CtrSrcControl.Left - (.Width * 1.5)

    End If

  End With

  Call CMsgbox_core_close

EH:
'  Set CtrTgtControl = Nothing
'  Set CtrSrcControl = Nothing

End Sub









' ASSETS ----------------------------------------------------------------------------'
Private Sub CMsgbox_asset_set_colors_pallete()

  ' Creates the CMsgBox theme color and its Status theme colors
  Dim VecColorBase As Variant
  Dim BooExistsLight As Boolean
  Dim LngBgColor As Long, LngBgFontColor As Long, LngFontColor As Long, i As Long, LngMax As Long
  Dim xR As Integer, xG As Integer, xB As Integer

  Const LngTopColor As Long = 255
  Const LngMinLight As Long = 200
  Const DblPercent As Double = 0.8


  On Error GoTo EH
  With cParam
    Let VecColorBase = Array(.cmbThmColor \ 256 ^ 0 And LngTopColor, .cmbThmColor \ 256 ^ 1 And LngTopColor, .cmbThmColor \ 256 ^ 2 And LngTopColor)

    ' Check for dark or light colors and find its darken or lighten colors for Bg and BgFont
    For i = LBound(VecColorBase) To UBound(VecColorBase)
      Let BooExistsLight = (VBA.CLng(VecColorBase(i)) >= LngMinLight)
      If BooExistsLight Then Exit For
    Next i

    Let LngMax = Application.WorksheetFunction.Max(VecColorBase)
    Let LngBgColor = .cmbThmColor
  End With

  Let xR = VBA.CLng(VecColorBase(0))
  Let xG = VBA.CLng(VecColorBase(1))
  Let xB = VBA.CLng(VecColorBase(2))
  Let VecColorBase = Empty

  If BooExistsLight Then ' Find its darken color

    Let xR = (((100 - (IIf(xR = LngMax, 0, xR) * (100 / LngTopColor))) * (1 - DblPercent)) / 100) * LngTopColor
    Let xG = (((100 - (IIf(xG = LngMax, 0, xG) * (100 / LngTopColor))) * (1 - DblPercent)) / 100) * LngTopColor
    Let xB = (((100 - (IIf(xB = LngMax, 0, xB) * (100 / LngTopColor))) * (1 - DblPercent)) / 100) * LngTopColor

    Let LngBgFontColor = VBA.RGB(xR, xG, xB)
    Let LngFontColor = LngBgFontColor

  Else ' Find its lighten color

    Let xR = (((100 - (IIf(xR = LngMax, 0, xR) * (100 / LngTopColor))) * (DblPercent)) / 100) * LngTopColor
    Let xG = (((100 - (IIf(xG = LngMax, 0, xG) * (100 / LngTopColor))) * (DblPercent)) / 100) * LngTopColor
    Let xB = (((100 - (IIf(xB = LngMax, 0, xB) * (100 / LngTopColor))) * (DblPercent)) / 100) * LngTopColor

    Let LngBgFontColor = VBA.RGB(xR, xG, xB)
    Let LngFontColor = LngBgColor

  End If

  ' Set finded colors to main CMsgBox theme color pallete
  With cColor

    Let .cmbBackground = LngBgColor
    Let .cmbBgFontcolor = LngBgFontColor
    Let .cmbFontcolor = LngFontColor

    ' Set Status theme colors
    Let .cmbAlertMaincolor = VBA.vbCyan
    Let .cmbQuestionMaincolor = VBA.RGB(250, 100, 100) ' cMsgVBOrange
    Let .cmbInfoMaincolor = VBA.vbBlue
    Let .cmbSuccessMaincolor = VBA.vbGreen
    Let .cmbErrorMaincolor = VBA.vbRed
    Let .cmbDefaultMaincolor = LngBgColor 'VBA.vbBlack

    Let .cmbAlertSoftcolor = VBA.RGB(230, 250, 240)
    Let .cmbQuestionSoftcolor = VBA.RGB(255, 250, 240)
    Let .cmbInfoSoftcolor = VBA.RGB(240, 240, 255)
    Let .cmbSuccessSoftcolor = VBA.RGB(240, 255, 240)
    Let .cmbErrorSoftcolor = VBA.RGB(255, 240, 240)
    Let .cmbDefaultSoftcolor = LngBgFontColor 'VBA.vbGrayText

  End With

EH:

End Sub

Private Function CMsgbox_asset_get_status_values() As String

  ' Get CMsgBox status values
  Dim StrMsgType As String

  On Error GoTo EH
  With cParam

    Select Case .cmbEnuMsgType
      Case cmbError: Let StrMsgType = VBA.ChrW(&H2716)
      Case cmbSuccess: Let StrMsgType = VBA.ChrW(&H2714)
      Case cmbInfo: Let StrMsgType = "i"
      Case cmbQuestion: Let StrMsgType = "?"
      Case cmbAlert: Let StrMsgType = "!"
      Case Else: Let StrMsgType = VBA.vbNullString ' cmbDefault
    End Select

  End With

  Let CMsgbox_asset_get_status_values = StrMsgType

EH:

End Function

Private Function CMsgbox_asset_get_align_horizontal( _
  ByVal CCtlTarget As MSForms.Control, _
  ByVal CEnuHorizontalAlign As CMB_ControlPositionX) As Long

  Dim xLeft As Long, xCMBWidth As Long, CtlTargetParentLeft As Long

  On Error GoTo EH
  With FmeMsgbox.Parent

    Let CtlTargetParentLeft = IIf(VBA.TypeName(CCtlTarget.Parent) = VBA.TypeName(FmeMsgbox.Parent), 0, CCtlTarget.Parent.Left)

    Let xCMBWidth = FmeMsgbox.Width
    Select Case CEnuHorizontalAlign
      Case cmbRight: Let xLeft = (CtlTargetParentLeft + CCtlTarget.Left + CCtlTarget.Width) - xCMBWidth
      Case cmbCenter: Let xLeft = ((CtlTargetParentLeft + CCtlTarget.Left + (CCtlTarget.Width / 2))) - (xCMBWidth / 2)
      Case Else: Let xLeft = CCtlTarget.Left + CtlTargetParentLeft  ' dpcLeft
    End Select

    If (xLeft + xCMBWidth) > .InsideWidth Then Let xLeft = xLeft - ((xLeft + xCMBWidth) - .InsideWidth)
    If xLeft < 0 Then Let xLeft = 0

  End With
  Let CMsgbox_asset_get_align_horizontal = xLeft

EH:

End Function

Private Function CMsgbox_asset_get_align_vertical( _
  ByVal CCtlTarget As MSForms.Control, _
  ByVal CEnuVerticalAlign As CMB_ControlPositionY) As Long

  Dim xTop As Long, xCMBHeight As Long, xCtlParentTop As Long

  On Error GoTo EH
  With FmeMsgbox.Parent

    ' Get the top of Control Target relative to its Parent
    Let xCtlParentTop = IIf(VBA.TypeName(CCtlTarget.Parent) = VBA.TypeName(FmeMsgbox.Parent), 0, CCtlTarget.Parent.Top) + 5

    ' Calculate top to CMsgBox
    Let xCMBHeight = FmeMsgbox.Height
    Select Case CEnuVerticalAlign
      Case cmbTop: Let xTop = (CCtlTarget.Top + xCtlParentTop) - xCMBHeight
      Case cmbMiddle: Let xTop = ((CCtlTarget.Top + xCtlParentTop) + (CCtlTarget.Height * 0.5)) - (xCMBHeight * 0.5)
      Case Else: Let xTop = (CCtlTarget.Top + xCtlParentTop + CCtlTarget.Height) ' cmbBottom
    End Select

    ' CMsgBox always must be inside Form area (Just if Form area > CMsgBox area)
    If (xTop + xCMBHeight) > (.InsideHeight * 1.1) Then Let xTop = xTop - ((xTop + xCMBHeight) - .InsideHeight)
    If xTop < 0 Then Let xTop = 0
  End With

  'Return calculated top
  Let CMsgbox_asset_get_align_vertical = xTop

EH:

End Function

Private Function CMsgbox_asset_get_frm_has_control( _
  ByVal UFrm As Object, _
  ByVal CStrTarget As String) As Boolean

  Dim CtlNeedle As MSForms.Control

  On Error GoTo EH
  Let CMsgbox_asset_get_frm_has_control = False

  ' Search for all Form Controls collection to find CStrTarget
  With UFrm
    For Each CtlNeedle In UFrm.Controls
      If VBA.LCase(CtlNeedle.Name) = VBA.LCase(CStrTarget) Then Let CMsgbox_asset_get_frm_has_control = True: Exit For
    Next CtlNeedle
    Set CtlNeedle = Nothing
  End With

EH:
  Set CtlNeedle = Nothing

End Function

Private Sub CMsgbox_asset_stop_frm_cancel_default_buttons()

  Dim BtnWithCancelOrDefault As Object

  On Error GoTo EH
  ' Save cancel property of current button in UserForm: GLOStr_ParentCancelButton
  For Each BtnWithCancelOrDefault In FmeMsgbox.Parent.Controls
    If VBA.TypeName(BtnWithCancelOrDefault) = "CommandButton" Then
      If BtnWithCancelOrDefault.Cancel Then
        Let GLOStr_ParentCancelButton = BtnWithCancelOrDefault.Name
        Exit For
      End If
    End If
  Next BtnWithCancelOrDefault

  ' Save default property of current button in UserForm: GLOStr_ParentDefaultButton
  For Each BtnWithCancelOrDefault In FmeMsgbox.Parent.Controls
    If VBA.TypeName(BtnWithCancelOrDefault) = "CommandButton" Then
      If BtnWithCancelOrDefault.Default Then
        Let GLOStr_ParentDefaultButton = BtnWithCancelOrDefault.Name
        Let BtnWithCancelOrDefault.Default = False
        Exit For
      End If
    End If
  Next BtnWithCancelOrDefault

EH:
  Set BtnWithCancelOrDefault = Nothing

End Sub

Private Sub CMsgbox_asset_restore_frm_cancel_default_buttons()

  Dim BtnWithCancelOrDefault As Object

  On Error GoTo EH
  If FmeMsgbox Is Nothing Then GoTo EH

  ' Restore Cancel button in main UserForm interface
  With cParam.cmbUFrm
    If Not GLOStr_ParentCancelButton = VBA.vbNullString Then
      Let .Controls(GLOStr_ParentCancelButton).Cancel = True
    End If

    ' Restore Default button in main UserForm interface
    If Not GLOStr_ParentDefaultButton = VBA.vbNullString Then Let .Controls(GLOStr_ParentDefaultButton).Default = True
  End With

EH:
  Set BtnWithCancelOrDefault = Nothing

End Sub

Private Function CMsgbox_asset_exists_frm() As Boolean

  Dim ObjFrm As Object

  On Error GoTo EH
  With cParam

    #If Debugging Then
      Call Debug_assert(Not .cmbUFrm Is Nothing, "core_open")
    #End If

    Let CMsgbox_asset_exists_frm = False

    'Verifies if UFRM is a correct object
    If Not .cmbUFrm Is Nothing Then
      If VBA.UserForms.Count > 0 Then
        For Each ObjFrm In VBA.UserForms
          If VBA.TypeName(.cmbUFrm) = VBA.TypeName(ObjFrm) Then Let CMsgbox_asset_exists_frm = True: Exit For
        Next ObjFrm
        Set ObjFrm = Nothing
      End If
    End If

  End With

  ' Validate Objects
  If Not CMsgbox_asset_exists_frm Then Call VBA.Err.Raise(CUSTOM_ERROR, "CMsgBox::-core_open", "Se requiere un objeto UserForm")

EH:
  If Not VBA.Err.Number = 0 Then
    Call VBA.MsgBox(VBA.Err.Description & VBA.vbNewLine & "#Error: " & VBA.Err.Number & VBA.vbNewLine & "Origen: " & VBA.Err.Source, vbInformation, "CMsgBox")
    Call VBA.Err.Clear
  End If

End Function

Private Sub CMsgbox_asset_required_control_setfocus()

  On Error GoTo EH
  If CMsgbox_asset_get_frm_has_control(cParam.cmbUFrm, cParam.cmbStrCtrlTargetName) Then

    With cParam.cmbUFrm.Controls(cParam.cmbStrCtrlTargetName)
      Select Case VBA.TypeName(cParam.cmbUFrm.Controls(cParam.cmbStrCtrlTargetName))
        Case "TextBox", "ComboBox", "Label", "ListBox", "CommandButton", "ToggleButton", "CheckBox", "OptionButton", "SpinButton", "SpinButton", "RefEdit"
          If .Enabled And .Visible Then .SetFocus
        Case Else
      End Select
    End With

  End If

EH:

End Sub




' ERRORS MANNAGEMENT AND ASSERTSIONS ------------------------------------------------'
Private Sub Debug_assert(ByVal Cond As Boolean, Optional Origin As String = VBA.vbNullString)

  #If Debugging Then
    Debug.Assert Cond
  #End If

End Sub




' PUBLIC METHODS LIST (INTERFACE) ---------------------------------------------------'
Public Sub CMsgBox( _
  Optional cMsg As String = VBA.vbNullString, _
  Optional cTitle As String = "CustomMsgboxClass", _
  Optional cSubtitle As String = VBA.vbNullString, _
  Optional cMsgType As CMB_MsgType = cmbDefault, _
  Optional cMsgButtons As CMB_ControlButtons = cmbNoButtons, _
  Optional cUFrm As MSForms.UserForm, _
  Optional cCtrlTargetName As String = VBA.vbNullString, _
  Optional cCtrlSourceName As String = VBA.vbNullString, _
  Optional cCtrlRequired As CMB_ControlRequired = cmbNoRequired, _
  Optional cPositionX As CMB_ControlPositionX = cmbCenter, _
  Optional cPositionY As CMB_ControlPositionY = cmbMiddle, _
  Optional cThmColor As XlRgbColor = rgbDarkGreen, _
  Optional cDuration As CMB_MsgDuration = cmbNoTime, _
  Optional cAspect As CMB_MsgBoxAspect = cmbPopupDefault)

  On Error GoTo EH

  Call CMsgbox_core_open(cMsg, cTitle, cSubtitle, cMsgType, cMsgButtons, cUFrm, cCtrlTargetName, cCtrlSourceName, cCtrlRequired, cPositionX, cPositionY, cThmColor, cDuration, cAspect)

EH:

End Sub
'============================= CLASS MODULES STRUCTURE =============================='
