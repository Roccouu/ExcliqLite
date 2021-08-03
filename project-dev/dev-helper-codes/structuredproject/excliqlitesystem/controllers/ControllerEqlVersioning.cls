VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppExcliqVersioning_current"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

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


'Allows to get ExcliqLite Version
'Based on SEMVER: www.semver.com

'============================= CLASS MODULES STRUCTURE =============================='
' ABSTRACT IMPLEMENTATIONS ----------------------------------------------------------'
Implements ICoreController

' OBJECT VARIABLES (GLOOBJ_) --------------------------------------------------------'
Private EHGLOBAL As AppErrorHandler

' VARIABLES (GLOStr_) ---------------------------------------------------------------'
Private GLOBOO_EH As Boolean

Private MINOR As Integer
Private MAJOR As Integer
Private PATCH As Integer

Private VERSION As String
Private APPNAME As String
Private APPPROPERNAME As String

Private AUTHOR As String
Private MAIL As String
Private SLOGAN As String
Private VERSIONYEAR As String
Private Origin As String
Private GIT As String
Private LICENSE As String
Private LICENSEINFO As String




' PROPERTIES ------------------------------------------------------------------------'
' ABSTRACT INTERFACE PROPERTIES IMPLEMENTATION --------------------------------------'
Public Property Set ICoreController_ErrorHandler(RHS As AppErrorHandler)
  Set EHGLOBAL = Nothing
  Set EHGLOBAL = RHS
  Let GLOBOO_EH = (Not RHS Is Nothing)
End Property

' PROPERTIES LOCALS -----------------------------------------------------------------'
Public Property Get VERSIONGET() As Variant
  VERSIONGET = VERSION
End Property

Public Property Get VERSIONGETNAME() As Variant
  VERSIONGETNAME = "v" & VERSION
End Property

Public Property Get VERSIONAPPNAME() As Variant
  VERSIONAPPNAME = APPNAME
End Property

Public Property Get VERSIONAUTHOR() As Variant
  VERSIONAUTHOR = AUTHOR
End Property

Public Property Get VERSIONAPPMAIL() As Variant
  VERSIONAPPMAIL = MAIL
End Property

Public Property Get VERSIONAPPSLOGAN() As Variant
  VERSIONAPPSLOGAN = SLOGAN
End Property

Public Property Get VERSIONAPPYEAR() As Variant
  VERSIONAPPYEAR = VERSIONYEAR
End Property

Public Property Get VERSIONAPPORIGIN() As Variant
  VERSIONAPPORIGIN = Origin
End Property

Public Property Get VERSIONAPPGIT() As Variant
  VERSIONAPPGIT = GIT
End Property

Public Property Get VERSIONAPPLICENSE() As Variant
  VERSIONAPPLICENSE = LICENSE
End Property

Public Property Get VERSIONAPPLICENSEINFO() As Variant
  VERSIONAPPLICENSEINFO = LICENSEINFO
End Property






' CONSTRUCTOR/DESTRUCTOR ------------------------------------------------------------'
Private Sub Class_Initialize()
  Dim MDL As Model
  
  Set EHGLOBAL = New AppErrorHandler
  Set MDL = New ModelApp
  
  MAJOR = 1
  MINOR = 0
  PATCH = 0
  VERSION = MAJOR & "." & MINOR & "." & PATCH
  APPNAME = "excliq lite"
  APPPROPERNAME = VBA.StrConv(APPNAME, vbProperCase)

  AUTHOR = "Roberto Carlos Romay Medina"
  MAIL = "<rocky.romay@gmail.com>"
  SLOGAN = "Básicos Excel para Comercialización de minerales en Bolivia."
  VERSIONYEAR = "2019 - " & VBA.Format(VBA.Date, "YYYY")
  Origin = "Potosí - Bolivia"
  GIT = "https://github.com/roccouu/"
  LICENSE = "MIT"
  LICENSEINFO = "LICENSE & ACKNOWLEDGMENTS" & VBA.vbNewLine & VBA.vbNewLine & _
  "MIT License" & VBA.vbNewLine & VBA.vbNewLine & _
  "Copyright (c) " & VERSIONYEAR & VBA.vbNewLine & "Roberto Carlos Romay Medina" & VBA.vbNewLine & VBA.vbNewLine & _
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
  " - Otto Javier González | www.excelymas.com" & VBA.vbNewLine & _
  " - Ismael Romero | www.excelforo.blogspot.com" & VBA.vbNewLine & _
  " - David Asurmendi | www.davidasurmendi.blogspot.com" & VBA.vbNewLine & _
  " - Sergio Alejandro Campos | www.exceleinfo.com"
  Set MDL = Nothing
End Sub






' CORE METHODS PRIVATE --------------------------------------------------------------'
Private Function VersionCallForeignMacro(ByVal FilePath As String, ByVal FileName As String, ByVal MacroName As String) As String
  Dim bMacro As String
  Dim Wb As Workbook
  Dim existFile As Boolean
  Dim file As String
  
  file = VBA.Dir(FilePath)
  existFile = (Len(file) > 0)
  
  VersionCallForeignMacro = VBA.vbNullString
  'Debug.Print Application.ThisWorkbook.Name
  If Application.ThisWorkbook.Name = FileName Then VersionCallForeignMacro = VERSION: GoTo VCFM
  'If "ExcliqLite_v1.0.0.xlam" = FileName Then VersionCallForeignMacro = VERSION: GoTo VCFM
  Set Wb = Nothing
  If existFile Then
    bMacro = "'" & FilePath & "'" & MacroName
    'Debug.Print bMacro
    On Error Resume Next
    Set Wb = Application.Workbooks(FileName)
    If Wb Is Nothing Then 'Closed Book, lets open it
      Application.Workbooks.Open FilePath
      Set Wb = Application.Workbooks(FileName)
    End If
    On Error GoTo 0
    
    On Error Resume Next
    If Not Wb Is Nothing Then
      'Debug.Print bMacro
      VersionCallForeignMacro = Application.Run(bMacro)
      'Debug.Print VersionCallForeignMacro
      Wb.Close SaveChanges:=False
    End If
  End If
  
VCFM:
  Set Wb = Nothing
  On Error GoTo 0
End Function

Private Function VersionVerifyNumbers(ByVal vLocal As Integer, ByVal vForeign As Integer) As Boolean
  VersionVerifyNumbers = (vLocal > vForeign)
End Function

Private Function VersionChecker(ByVal v As String) As Boolean
  Dim vVector() As String
  
  Dim vMAJOR As Variant
  Dim vMINOR As Variant
  Dim vPATCH As Variant
  
  Dim bMINOR As Boolean
  Dim bMAJOR As Boolean
  Dim bPATCH As Boolean
    
  VersionChecker = False
  vVector = VBA.Split(v, ".")
  
  vMAJOR = VBA.CInt(vVector(0))
  vMINOR = VBA.CInt(vVector(1))
  vPATCH = VBA.CInt(vVector(2))
  
  '1.0.0 1.0.1 1.1.2 2.0.1
  If v = VBA.vbNullString Then GoTo VVE
  'Verify each number of version. Major first
  If (MAJOR > vMAJOR) Then
    VersionChecker = True 'Instalar este addin y Desinstalar Addins de menor versión.
  ElseIf (MINOR > vMINOR) Then
    VersionChecker = True ': GoTo VVE 'Instalar este addin y Desinstalar Addins de menor versión.
  ElseIf (PATCH > vPATCH) Then
    VersionChecker = True ': GoTo VVE 'Instalar este addin y Desinstalar Addins de menor versión.
  End If

VVE:
End Function

Private Function VersionCheckAddins() As Boolean
  Dim i As Integer, j As Integer, k As Integer
  Dim vVersion As String
  Dim minorVersions() As Integer
  Dim minorVersionsA() As Integer
  Dim ExistsVersionOthers As Boolean
  Dim Obsolete As Boolean
  Dim StrAux As Variant
  Dim GetVersionMacro As String
  
  On Error Resume Next
  GetVersionMacro = "!IndexUI.ExcliqLiteGetVersion"
  VersionCheckAddins = False
  ExistsVersionOthers = False
  Obsolete = False
  
  With Application
    ReDim minorVersions(0 To .AddIns.Count)
    j = 0
    For i = 1 To .AddIns.Count
      StrAux = VBA.InStr(VBA.LCase(.AddIns(i).Name), "excliqlite")
      If StrAux > 0 Then
        vVersion = VersionCallForeignMacro(.AddIns(i).FullName, .AddIns(i).Name, GetVersionMacro) ' .Run(.AddIns(i).FullName & "!GetVersion")
        If vVersion <> VERSION Then
          If vVersion <> VBA.vbNullString Then
            ExistsVersionOthers = True 'Already Exist other versions
            VersionCheckAddins = VersionChecker(vVersion)
            If VersionCheckAddins Then
              minorVersions(j) = i
              j = j + 1
            Else
              k = k + 1
            End If
          End If
        End If
      End If
    Next i
        
    If ExistsVersionOthers Then
      If j > 0 Then
        j = j - 1
        ReDim minorVersionsA(0 To j)
        For i = 0 To j
          minorVersionsA(i) = minorVersions(i)
        Next i
      End If

      If VersionCheckAddins Then 'Install addin, uninstall lower version addins.
        If k > 0 Then Obsolete = True: GoTo VCAO 'Existen versiones mayores instaladas. No se instala el Addin y salir.
        j = j + 1
        For i = 1 To j 'Uninstall oldiest versions
          If .AddIns(minorVersionsA(i - 1)).Installed Then .AddIns(minorVersionsA(i)).Installed = False
        Next
      Else 'Versión superior existe
        Obsolete = True 'Existen versiones mayores instaladas. No se instala el Addin y salir.
      End If
    Else 'Just exist thisWorkbook into the Addins list
      VersionCheckAddins = True
    End If
  End With

VCAO:
  If Obsolete Then Call VBA.MsgBox( _
    "Esta versión de Excliq Lite es obsoleta, ya tiene instalada una versión más reciente. Se recomienda que utilice esa versión.", _
    vbInformation, APPPROPERNAME & " v" & VERSION)
  Erase minorVersions
  Erase minorVersionsA
  On Error GoTo 0
End Function

Private Function VersionGetExtensionFile(ByVal file As String) As String
  VersionGetExtensionFile = VBA.vbNullString
  If file <> VBA.vbNullString Then VersionGetExtensionFile = VBA.StrReverse(VBA.Left(VBA.StrReverse(file), 5))
  'VersionGetExtensionFile = ".xlam"
End Function

Private Function VersionGetExtensionFileIsAddin(ByVal file As String) As Boolean
  VersionGetExtensionFileIsAddin = (VBA.LCase(VersionGetExtensionFile(Application.ThisWorkbook.FullName)) = ".xlam")
End Function






' METHODS PUBLIC --------------------------------------------------------------------'
Public Function VersionVerify() As Boolean
  Dim FileExtension As String

  If VersionGetExtensionFileIsAddin(Application.ThisWorkbook.FullName) Then
    VersionVerify = VersionCheckAddins()
  Else
    VersionVerify = True
  End If
End Function

' ABSTRACT INTERFACE UI METHODS LIST AND UDF METHODS LIST (PUBLIC) ------------------'
Private Function ICoreController_GetSolution(ByVal CStrRequester As String, ByVal CVarRequest As Variant, CVarResponse As Variant) As Boolean
End Function

Public Function ICoreController_GetSolutionUDF(ByVal CStrRequester As String, ByVal CVarRequest As Variant) As Variant
  
  On Error GoTo EH
  Let ICoreController_GetSolutionUDF = VersionVerify()
  
EH:
  Call EHGLOBAL.ErrorHandlerRaise("APPREG::IGetSolution")

End Function
' METHODS PUBLIC --------------------------------------------------------------------'
'============================= CLASS MODULES STRUCTURE =============================='


