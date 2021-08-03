Attribute VB_Name = "Index"
Option Explicit
Option Private Module

Public exq_crt As AppExcliq_current
'Public EHGLOBAL As AppErrorHandler

'Public exq_crt As New AppExcliq
'Public EHGlobal As New AppErrorHandler

'LICENSE & ACKNOWLEDGMENTS
'
'MIT License
'
'Copyright (c) 2019 Roberto Carlos Romay Medina
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.
'
'Acknowledgments to:
' - StarUML v2.8.0
' - Pencil v3.0.3
' - VSCode v1.41.0
' - InkScape v0.92.4
' - Gimp v2.10.2
' - Just Color Picker v4.6
' - Office RibbonX Editor | Fernando Andreu v1.5.1.418
' - MS Excel v1911
' - VBA7.1 V1091
' - Otto Javier González | www.excelymas.com
' - Ismael Romero | www.excelforo.blogspot.com
' - David Asurmendi | www.davidasurmendi.blogspot.com
' - Sergio Alejandro Campos | www.exceleinfo.com
'
'
'Roccou: I think RefEdits works very well with a good treatment.
'GitHub: https://github.com/roccouu/
'ExcliqLite's home page: https://roccouu.github.io/ExcliqLite/
'ExcliqLite's repo: https://github.com/Roccouu/ExcliqLite
'2019, POTOSÍ - BOLÍVIA


'CUSTOM ERRORS: _
APP: 514 _
RIBBON: 515 _
CLASSES: 516 _
VIEWS: 517 _
CONTROLLERS: 518 _
MODELS: 519



'Test functions
Sub test(ByVal ix As IRibbonUI)
  Debug.Print "Ribbon"
End Sub

Sub cVersion()
  Call exq_crt.CallUDFsRegistration
End Sub



'AUTOMATIZATION FUNCTIONS
Sub auto_open()
  #If Debugging Then
    Debug.Print "Auto"
  #End If
  If exq_crt Is Nothing Then Set exq_crt = New AppExcliq_current
  'If EHGLOBAL Is Nothing Then Set EHGLOBAL = New AppErrorHandler
  'Call exq_crt.CallUDFsRegistration
End Sub

Sub auto_close()
  On Error GoTo EH
  If Not exq_crt Is Nothing Then Set exq_crt = Nothing
  'If Not EHGLOBAL Is Nothing Then Set EHGLOBAL = Nothing
EH:
  On Error GoTo -1
End Sub

'Ribbon
Public Sub ExcliqliteRibbonBegin(ByVal IRIBBON As IRibbonUI)
  Call Application.Volatile
  If Not exq_crt Is Nothing Then Call exq_crt.AppInit(IRIBBON)
End Sub

'... Events: Status controls ...
Public Sub ExcliqliteRibbonGetInitialValues(ByVal Control As IRibbonControl, ByRef ValueControl As Variant)
  'Call VBA.MsgBox("ExcliqliteRibbonGetInitialValues: Initializing controlUX: " & Control.id)
  If Not exq_crt Is Nothing Then Let ValueControl = exq_crt.AppRibbonStatusGetter(Control.id) 'Initial status of controls
End Sub

Public Sub ExcliqliteRibbonGetEnabled(ByVal Control As IRibbonControl, ByRef StatusControl As Variant)
  'Call VBA.MsgBox("ExcliqliteRibbonGetEnabled: Initializing controlUX: " & Control.id)
  If Not exq_crt Is Nothing Then Let StatusControl = exq_crt.AppRibbonStatusSetter(Control.id, True) 'For menu controls enabled/disabled status
End Sub

'... Functionality subs ...
'Sub ExcliqliteRibbonActionExecutor(ByVal control As IRibbonControl)'DELETE
'  On Error GoTo EH
'  Call Application.Volatile
'  If Not exq_crt Is Nothing Then Call exq_crt.AppRibbonExecutorControls(control.id)
'EH:
'  If Not EHGlobal Is Nothing Then Call EHGlobal.ErrorHandlerDisplay("Index::Exe")
'End Sub

Sub ExcliqliteRibbonActionExecutor(ByVal Control As IRibbonControl, Optional id As String, Optional Index As Integer)
  On Error GoTo EH
  Call Application.Volatile
  'Call VBA.MsgBox("Hi: " & id & ", my index is: " & index & ", Id in control: " & control.id)
  If Not exq_crt Is Nothing Then Call exq_crt.AppRibbonExecutorControls(Control.id, id, Index)
EH:
  On Error GoTo -1
  'If Not EHGLOBAL Is Nothing Then Call EHGLOBAL.ErrorHandlerDisplay("Index::Exe")
End Sub

''Callback for drpdwns onAction'DELETE
'Sub ExcliqliteRibbonActionExecutorDrp(ByVal control As IRibbonControl, Optional id As String, Optional index As Integer)
'  Call VBA.MsgBox("Hi: " & id & ", my index is: " & index & "Id in control: " & control.id)
'End Sub
'
''CREDENTIALS BRIDGE
'Private Function EXCLIQLITE_GetCredentials(ByVal Petition As String) As String
'  Let EXCLIQLITE_GetCredentials = exq_crt.AppExcliqGetCredentials(Petition)
'End Function



