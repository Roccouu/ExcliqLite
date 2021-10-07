VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassEqlControllerVersioning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ===================== CLASS EXCLIQ LITE VERSIONING STRUCTURE ===================== '
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



' ===================== CLASS EXCLIQ LITE VERSIONING STRUCTURE ===================== '
' ABSTRACT IMPLEMENTATIONS --------------------------------------------------------- '
' OBJECT VARIABLES (GLOOBJ_) ------------------------------------------------------- '
' PROPERTIES ----------------------------------------------------------------------- '
' ABSTRACT INTERFACE PROPERTIES IMPLEMENTATION ------------------------------------- '
' CORE METHODS LIST (PRIVATE) ------------------------------------------------------ '
' Version_verify
' Version_call_foreign_macro
' Version_checker
' Version_check_addins
' Version_getfull
' ERROR ASSERTSIONS ---------------------------------------------------------------- '
' Debug_asset
' METHODS PUBLIC ------------------------------------------------------------------- '
' ABSTRACT INTERFACE UI METHODS LIST AND UDF METHODS LIST (PUBLIC) ----------------- '
' ClassEqlControllerInterface_GetSolution
' ClassEqlControllerInterface_GetSolutionUDF
' ===================== CLASS EXCLIQ LITE VERSIONING STRUCTURE ===================== '




' ===================== CLASS EXCLIQ LITE VERSIONING STRUCTURE ===================== '
' Allows to get ExcliqLite Version
' Based on SEMVER: www.semver.com

' ABSTRACT IMPLEMENTATIONS --------------------------------------------------------- '
Implements ClassEqlControllerInterface

' OBJECT VARIABLES (GLOOBJ_) ------------------------------------------------------- '
Private RES As ClassEqlResources





' PROPERTIES ----------------------------------------------------------------------- '
' ABSTRACT INTERFACE PROPERTIES IMPLEMENTATION ------------------------------------- '
Public Property Set ClassEqlControllerInterface_ResourcesApp(ByRef RHS As ClassEqlResources)

  Set RES = RHS

End Property




' CORE METHODS PRIVATE ------------------------------------------------------------- '
Private Function Version_verify() As Boolean

  On Error GoTo EH

  If Application.ThisWorkbook.IsAddin Then
    Let Version_verify = Version_check_addins()
  Else
    Let Version_verify = True
  End If

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise("VERSION::-verify")

End Function

Private Function Version_call_foreign_macro( _
  ByVal VStrFilePath As String, _
  ByVal VStrFileName As String, _
  ByVal VStrMacroName As String, _
  ByVal VStrAppName As String) As String

  Dim Wbk As Workbook
  Dim StrFile As String, StrBMacro As String
  Dim BooExistFile As Boolean

  Let StrFile = VBA.Dir(VStrFilePath)
  Let BooExistFile = (VBA.Len(StrFile) > 0)

  Let Version_call_foreign_macro = VBA.vbNullString
  'Debug.Print Application.ThisWorkbook.Name
  If Application.ThisWorkbook.Name = VStrFileName Then Version_call_foreign_macro = Version_getfull(): GoTo EH

  Set Wbk = Nothing
  If BooExistFile Then
    Let StrBMacro = "'" & VStrFilePath & "'" & VStrMacroName
    'Debug.Print StrBMacro
    If Not RES.ExcelExistsWorkbook(VStrFileName) Then ' Closed Book, lets open it
      Call Application.Workbooks.Open(FileName:=VStrFilePath, ReadOnly:=True)
    End If
    Set Wbk = Application.Workbooks(VStrFileName)
    If Not Wbk Is Nothing Then
      'Debug.Print StrBMacro
      Let Version_call_foreign_macro = Application.Run(StrBMacro)
      Let Version_call_foreign_macro = VBA.Replace(VBA.LCase(Version_call_foreign_macro), VBA.LCase(VStrAppName & " v"), VBA.vbNullString)
      'Debug.Print Version_call_foreign_macro
      Call Wbk.Close(SaveChanges:=False)
    End If
  End If

EH:
  Set Wbk = Nothing
  Call RES.ErrorHandler.ErrorHandlerRaise("VERSION::-call_foreign_macro")

End Function

Private Function Version_checker(ByVal VStrVersionLocal As String, ByVal v As String) As Boolean

  Dim VecVector As Variant
  Dim StrAux0 As String
  Dim IntMAJOR As Integer, IntMINOR As Integer, IntPATCH As Integer, IntMAJORLocal As Integer, IntMINORLocal As Integer, IntPATCHLocal As Integer

  Let Version_checker = False
  If v = VBA.vbNullString Then GoTo EH

  ' Get local version numbers
  Let StrAux0 = VStrVersionLocal
  Let VecVector = VBA.Split(StrAux0, ".")
  Let IntMAJORLocal = VBA.CInt(VecVector(0))
  Let IntMINORLocal = VBA.CInt(VecVector(1))
  Let IntPATCHLocal = VBA.CInt(VecVector(2))
  Let VecVector = Empty

  ' Get foreign addin version numbers
  Let VecVector = VBA.Split(v, ".")
  Let IntMAJOR = VBA.CInt(VecVector(0))
  Let IntMINOR = VBA.CInt(VecVector(1))
  Let IntPATCH = VBA.CInt(VecVector(2))
  Let VecVector = Empty

  ' 1.0.0 1.0.1 1.1.2 2.0.1
  ' Verify each version number. Major first
  Let Version_checker = (IntMAJORLocal > IntMAJOR Or IntMINORLocal > IntMINOR Or IntPATCHLocal > IntPATCH)

EH:
  Let VecVector = Empty
  Call RES.ErrorHandler.ErrorHandlerRaise("VERSION::-checker")

End Function

Private Function Version_check_addins() As Boolean

  Dim VecMinorVersions() As Integer
  Dim StrTitle As String, StrTitleRgx As String, StrVersion As String, StrVersionLocal As String, StrFullVersionLocal As String, StrGetVersionMacro As String, StrAux0 As String
  Dim BooExistsVersionOthers As Boolean, BooObsolete As Boolean, BooAux0 As Boolean
  Dim i As Integer, j As Integer, k As Integer

  On Error Resume Next
  Let Version_check_addins = False

  Let StrTitle = RES.AppStrTitle ' Excliq Lite
  Let StrVersionLocal = Version_getfull(BooJustVersion:=True)
  Let StrFullVersionLocal = Version_getfull()
  Let StrGetVersionMacro = "!IndexUI.ExcliqLiteGetVersion" 'EQL_VERSION"
  Let BooExistsVersionOthers = False
  Let BooObsolete = False

  With Application
    Let StrAux0 = VBA.StrReverse(VBA.Mid(VBA.StrReverse(.ThisWorkbook.Name), VBA.InStr(VBA.StrReverse(.ThisWorkbook.Name), ".") + 1))
    Let StrTitleRgx = VBA.LCase(VBA.Replace(RES.AppStrTitle, " ", "\s?")) ' "excliq\s?lite"
    If Not RES.REGEX.isExcliqLiteAddin(StrTitleRgx, StrAux0) Then GoTo EH ' Close this addin, it's name was changed
    Let j = 0: Let k = 0
    For i = 1 To .AddIns.Count
      Let StrAux0 = VBA.StrReverse(VBA.Mid(VBA.StrReverse(.AddIns(i).Name), VBA.InStr(VBA.StrReverse(.AddIns(i).Name), ".") + 1))
      Let BooAux0 = RES.REGEX.isExcliqLiteAddin(StrTitleRgx, StrAux0)
      If BooAux0 Then
        Let StrVersion = Version_call_foreign_macro(.AddIns(i).FullName, .AddIns(i).Name, StrGetVersionMacro, StrTitle) ' .Run(.AddIns(i).FullName & "!GetVersion")
        If StrVersion = StrVersionLocal Then
          If Not .ThisWorkbook.Name = .AddIns(i).Name Then GoTo EH ' Close this addin, just one is required
        Else
          Let BooExistsVersionOthers = Not (StrVersion = VBA.vbNullString)
          If BooExistsVersionOthers Then ' Already Exist other versions
            If Version_checker(StrVersionLocal, StrVersion) Then ' This addin is the Last Version
              Let Version_check_addins = True
              ReDim Preserve VecMinorVersions(j)
              Let VecMinorVersions(j) = i
              Let j = j + 1
            Else
              Let k = k + 1
            End If
          End If
        End If
      End If
    Next i

    ' Unistall oldiest versions or skip installation if this version is older or install this version if is LV
    If BooExistsVersionOthers Then
      If Version_check_addins Then ' Install addin, uninstall lower version addins.
        If k > 0 Then Let BooObsolete = True: GoTo EH ' This addin is obsolete, skip instalation
        For i = LBound(VecMinorVersions) To UBound(VecMinorVersions) ' Uninstall oldiest versions
          If .AddIns(VecMinorVersions(i)).Installed Then Let .AddIns(VecMinorVersions(i)).Installed = False
        Next
      Else ' Greatest version exists
        Let BooObsolete = True ' This addin is obsolete, skip instalation
      End If
    Else ' Just exist thisWorkbook into the Addins list
      Let Version_check_addins = True
    End If
  End With

EH:
  If BooObsolete Then Call VBA.MsgBox("Esta versión de Excliq Lite es obsoleta, ya tiene instalada una versión más reciente. Se recomienda que utilice esa versión.", vbInformation, StrFullVersionLocal)
  Erase VecMinorVersions
  Call RES.ErrorHandler.ErrorHandlerRaise("VERSION::-check_addins")

End Function

Private Function Version_getfull(Optional BooJustVersion As Boolean = False) As String

  Dim MDL As ClassEqlModel

  On Error GoTo EH

  Set MDL = New ClassEqlModel
  Set MDL.ResourcesApp = RES
  Let Version_getfull = IIf(BooJustVersion, VBA.vbNullString, VBA.StrConv(MDL.AppTitle, vbProperCase) & " v") & MDL.AppVersion
  Set MDL = Nothing

EH:
  Set MDL = Nothing
  On Error GoTo -1

End Function




' ERRORS MANNAGEMENT AND ASSERTSIONS ----------------------------------------------- '
Private Sub Debug_assert(ByVal Cond As Boolean, Optional Origin As String = VBA.vbNullString)

  #If Debugging Then
    Debug.Assert Cond
    If Not Cond Then Debug.Print "Debugged from: " & Origin
  #End If

End Sub





' METHODS PUBLIC ------------------------------------------------------------------- '
' ABSTRACT INTERFACE UI METHODS LIST AND UDF METHODS LIST (PUBLIC) ----------------- '
Private Function ClassEqlControllerInterface_GetSolution(ByVal CStrRequester As String, ByRef CVarRequest As Variant, ByRef CVarResponse As Variant) As Boolean

  On Error GoTo EH
  Let ClassEqlControllerInterface_GetSolution = Version_verify()

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise("VERSION::IGetSolution")

End Function

Public Function ClassEqlControllerInterface_GetSolutionUDF(ByVal CStrRequester As String, ByVal CVarRequest As Variant) As Variant

  On Error GoTo EH
  Let ClassEqlControllerInterface_GetSolutionUDF = Version_getfull()

EH:
  On Error GoTo -1

End Function
' METHODS PUBLIC ------------------------------------------------------------------- '
' ===================== CLASS EXCLIQ LITE VERSIONING STRUCTURE ===================== '


