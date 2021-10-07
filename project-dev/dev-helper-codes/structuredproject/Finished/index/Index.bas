Attribute VB_Name = "Index"
Option Explicit
Option Private Module


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


' ============================= INDEX MODULE STRUCTURE ============================= '
' OBJECT VARIABLES (GLOOBJ_) ------------------------------------------------------- '
' VARIABLES (GLOStr_) -------------------------------------------------------------- '
' VARIABLES CUSTOM ----------------------------------------------------------------- '
' CONSTANTS (GLOSTR_) -------------------------------------------------------------- '
' CORE METHODS LIST PRIVATE -------------------------------------------------------- '
' auto_open
' auto_close
' Test functions ------------------------------------------------------------------- '
' test
' PUBLIC METHODS LIST (INTERFACE) -------------------------------------------------- '
' ExcliqliteRibbonBegin
' ExcliqliteRibbonGetInitialValues
' ExcliqliteRibbonGetEnabled
' ExcliqliteRibbonActionExecutor
' ExcliqliteCloseHelp_click
' ============================= INDEX MODULE STRUCTURE ============================= '




' ============================= INDEX MODULE STRUCTURE ============================= '
' OBJECT VARIABLES (GLOOBJ_) ------------------------------------------------------- '
Public exq_crt As ClassEqlMain




' CORE METHODS LIST PRIVATE -------------------------------------------------------- '
' AUTOMATION FUNCTIONS
Private Sub auto_open()

  #If Debugging Then
    Debug.Print "Auto"
  #End If
  If exq_crt Is Nothing Then
    Set exq_crt = New ClassEqlMain
    Call exq_crt.AppBegin
  End If

End Sub

Private Sub auto_close()

  On Error GoTo EH
  If Not exq_crt Is Nothing Then Set exq_crt = Nothing

EH:
  On Error GoTo -1

End Sub

' Test functions ------------------------------------------------------------------- '
Public Sub test(ByVal ix As IRibbonUI)
  Debug.Print "Ribbon"
End Sub


' PUBLIC METHODS LIST (INTERFACE) -------------------------------------------------- '
' Ribbon
Public Sub ExcliqliteRibbonBegin(ByVal IRIBBON As IRibbonUI)

  Call Application.Volatile(True)
  If Not exq_crt Is Nothing Then Call exq_crt.AppInit(IRIBBON)

End Sub

' ... Events: Status controls ...
Public Sub ExcliqliteRibbonGetInitialValues(ByVal Control As IRibbonControl, ByRef ValueControl As Variant)

  'Call VBA.MsgBox("ExcliqliteRibbonGetInitialValues: Initializing controlUX: " & Control.id)
  If Not exq_crt Is Nothing Then Let ValueControl = exq_crt.AppRibbonStatusGetter(Control.id) 'Initial status of controls

End Sub

Public Sub ExcliqliteRibbonGetEnabled(ByVal Control As IRibbonControl, ByRef StatusControl As Variant)

  'Call VBA.MsgBox("ExcliqliteRibbonGetEnabled: Initializing controlUX: " & Control.id)
  If Not exq_crt Is Nothing Then Let StatusControl = exq_crt.AppRibbonStatusSetter(Control.id, True) 'For menu controls enabled/disabled status

End Sub

Sub ExcliqliteRibbonActionExecutor(ByVal Control As IRibbonControl, Optional id As String, Optional Index As Integer)

  On Error GoTo EH
  Call Application.Volatile
  'Call VBA.MsgBox("Hi: " & id & ", my index is: " & index & ", Id in control: " & control.id)
  If Not exq_crt Is Nothing Then Call exq_crt.AppRibbonExecutorControls(Control.id, id, Index)

EH:
  On Error GoTo -1

End Sub

Public Sub ExcliqliteCloseHelp_click(ByVal StrBookName As String, ByVal StrSheetName As String)

  On Error GoTo EH
  Call Application.Volatile
  'Call VBA.MsgBox("Hi: " & id & ", my index is: " & index & ", Id in control: " & control.id)
  If Not exq_crt Is Nothing Then Call exq_crt.AppRibbonExecutorControls("btnexcliqliteclose-" & StrBookName, StrSheetName)
EH:
  On Error GoTo -1

End Sub
' PUBLIC METHODS LIST (INTERFACE) -------------------------------------------------- '
' ============================= INDEX MODULE STRUCTURE ============================= '

