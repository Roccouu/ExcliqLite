VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const LINE_NO_TEXT As String = "Línea Número: "
Private ICONTYPE As VbMsgBoxStyle
Private FULLSRC As String
Private CUSTOMMSG As String

Private AppTitle As String

Private GLOStrMsg As String
Private GLODblNum As Double
Private GLOStyIco As VbMsgBoxStyle

' ExcelMacroMastery.com Error handling code
' https://excelmacromastery.com/vba-error-handling/

' Example of using:
'
' 1. Place DisplayError in the topmost sub at the bottom. Replace the third paramter
'    with the name of the sub.
' DisplayError Err.source, Err.Description, "Module1.Topmost", Erl
'
' 2. Place RaiseError in all the other subs at the bottom of each. Replace the third paramter
'    with the name of the sub.
' RaiseError Err.Number, Err.source, "Module1.Level1", Err.Description, Erl
'
'
' 3. The error handling in each sub should look like this
'
'  Sub subName()
'
'    On Error Goto eh
'
'    The main code of the sub here!!!!!
'
'  done:
'      Exit Sub
'  eh:
'      DisplayError Err.Source, Err.Description, "Module1.Topmost", Erl
'  End Sub
'


'PROPERTIES
Public Property Get ErrorMsg() As String
  
  Let ErrorMsg = GLOStrMsg
  
End Property

'Public Property Let PROPStrMsg(ByVal StrMsg As String)
'  Let GLOStrMsg = StrMsg
'End Property

Public Property Get ErrorNum() As Single
  
  Let ErrorNum = GLODblNum
  
End Property

'Public Property Let PROPSngNum(ByVal SngNum As Single)
'  Let GLOSngNum = SngNum
'End Property

Public Property Get PROPStyIco() As VbMsgBoxStyle
  
  Let PROPStyIco = GLOStyIco
  
End Property

Public Property Let PROPStyIco(ByVal StyIco As VbMsgBoxStyle)
  
  Let PROPStyIco = StyIco

End Property




'CONSTRUCTOR CLASS
Private Sub Class_Initialize()
  
  Let FULLSRC = VBA.vbNullString

End Sub




'METHODS PRIVATE
Private Sub Debug_assert(ByVal Cond As Boolean, Optional Origin As String = VBA.vbNullString)
  
  #If Debugging Then
    Debug.Assert Cond
    If Not Cond Then Debug.Print "Debugged from: " & Origin
  #End If

End Sub

' Reraises an error and adds line number and current procedure name
Private Sub RaiseError( _
  ByVal errorNo As Long, _
  ByVal src As String, _
  ByVal proc As String, _
  ByVal desc As String, _
  ByVal lineNo As Long, _
  Optional MsgText As String = VBA.vbNullString)

  ' If called for the first time then add line number
  If FULLSRC = VBA.vbNullString Then
    ' Add error line number if present
    If lineNo <> 0 Then
      Let FULLSRC = VBA.vbNewLine & LINE_NO_TEXT & lineNo & " "
    End If
    ' Add procedure to source
    Let CUSTOMMSG = desc & IIf(MsgText = VBA.vbNullString, VBA.vbNullString, VBA.vbNewLine & MsgText)
    Let FULLSRC = FULLSRC & IIf(FULLSRC = VBA.vbNullString, FULLSRC, VBA.vbNewLine) & proc
  Else
    ' If error has already been raised simply add on procedure name
    Let CUSTOMMSG = desc
    Let FULLSRC = proc & VBA.vbNewLine & src
  End If
  
  ' Pause the code here when debugging
  '(To Debug: "Tools->VBA Properties" from the menu.
  ' Add "Debugging=1" to the
  ' "Conditional Compilation Arguments.)
  '#If Debugging Then
  '  Debug.Assert False
  '#End If

  ' Reraise the error so it will be caught in the caller procedure
  ' (Note: If the code stops here, make sure DisplayError has been
  ' placed in the topmost procedure)
  Call VBA.Err.Raise(errorNo, FULLSRC, CUSTOMMSG)

End Sub

' Displays the error when it reaches the topmost sub
' Note: You can add a call to logging from this sub
Private Sub DisplayError( _
  ByVal src As String, _
  ByVal desc As String, _
  ByVal sProcname As String, _
  ByVal lineNo As Long)

  ' Check If the error happens in topmost sub
  If FULLSRC = VBA.vbNullString Then
    ' Reset string to remove "VBAProject" and add line number if it exists
    Let src = IIf(lineNo = 0, VBA.vbNullString, vbNewLine & LINE_NO_TEXT & lineNo)
  End If

'  ' Build the final message
  Let FULLSRC = "Ha ocurrido el siguiente error: " & VBA.vbNewLine & VBA.Err.Description & vbNewLine & vbNewLine & "Orígen del Error: " & VBA.vbNewLine
  Let FULLSRC = FULLSRC & sProcname & vbNewLine & src

  ' Display the message
  Call VBA.MsgBox(FULLSRC, ICONTYPE, AppTitle)

  ' reset the fullsrc value
  Let FULLSRC = VBA.vbNullString
  Let CUSTOMMSG = VBA.vbNullString
  On Error GoTo -1

End Sub


Private Sub RegisterError(ByVal Calleer As String, ByVal MsgIconType As VbMsgBoxStyle, Optional Display As Boolean = False, Optional MsgText As String = VBA.vbNullString)
  
  If Display Then 'Call DisplayError to show it
    Call DisplayError(VBA.Err.Source, VBA.Err.Description, Calleer, Erl)
  Else 'Raises the error
    'Sets the Msgbox icon for the error message
    'vbQuestion:32
    'vbInformation:64
    'vbExclamation:48
    'vbCritical:16
    'vbDefaultButton1=0
    'Raises the error
    If ICONTYPE = vbDefaultButton1 Then Let ICONTYPE = MsgIconType
    Debug.Print Calleer, "Error Number: " & VBA.Err.Number, "Description: ", VBA.Err.Description, "(" & VBA.Date() & "|" & VBA.Time() & ")"
    Call RaiseError(VBA.Err.Number, VBA.Err.Source, Calleer, VBA.Err.Description, Erl, MsgText)
  End If

End Sub

'PUBLIC METHODS
Public Sub ErrorHandlerRaise(ByVal Calleer As String, Optional MsgIconType As VbMsgBoxStyle = VBA.vbExclamation, Optional CustomMsgText As String = VBA.vbNullString)
  
  If VBA.Err.Number <> 0 Then
    Call RegisterError(Calleer, MsgIconType, MsgText:=CustomMsgText)
    Call ErrorHandlerSet(VBA.Err.Number, "Orígen: " & Calleer & VBA.vbNewLine & VBA.Err.Description)
  End If

End Sub

Public Sub ErrorHandlerDisplay(ByVal Calleer As String, Optional xIcontype As VBA.VbMsgBoxStyle = VBA.vbExclamation)
  
  If ICONTYPE = 0 Then ICONTYPE = xIcontype
  If VBA.Err.Number <> 0 Then Call RegisterError(Calleer, vbDefaultButton1, True)

End Sub

Public Sub ErrorHandlerClear()
  
  If Not (VBA.Err.Number = 0) Then
    On Error GoTo -1
    Call VBA.Err.Clear
  End If

End Sub

Public Sub ErrorHandlerSet(ByVal ErrNum As Double, ByVal ErrMsg As String)
  
  Let GLODblNum = ErrNum
  Let GLOStrMsg = ErrMsg

End Sub
