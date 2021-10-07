VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassEqlResourcesRegex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ================= CLASS EXCLIQ LITE RESOURCES REGEX STRUCTURE ==================== '
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



' ================= CLASS EXCLIQ LITE RESOURCES REGEX STRUCTURE ==================== '
' OBJECT VARIABLES (GLOOBJ_) ------------------------------------------------------- '
' PROPERTIES ----------------------------------------------------------------------- '
' CONSTRUCTOR/DESTRUCTOR ----------------------------------------------------------- '
' CORE METHODS LIST (PRIVATE) ------------------------------------------------------ '
' ERROR ASSERTSIONS ---------------------------------------------------------------- '
' Debug_asset
' CORE METHODS LIST (PUBLIC) ------------------------------------------------------- '
' isNumberEntireVectorCommaSeparated
' isNumberEntireAndDecimalVectorColonSeparated
' isNumberWithDecimalsSemicolonAndSeparatedSlashGrouped
' isNumberVector
' isNumberMinusReg
' isEmptyStringReg
' isStringReg
' isNumberReg
' isPercentReg
' isDateReg
' isRange
' isRangeWithName
' isLocalPath
' isNetworkPath
' isPhonenumber
' isEmail
' isWebpage
' isNotavaiable
' isMulticurrency
' isRekey
' isFinalkey
' isOreItem
' isThreeDigitFloatNumber
' CustomRegex
' isProviderType
' isProvidersTypes
' isMineralType
' isMineralComplex
' isMineralBrosa
' isMineralExcliqlite
' isMineralUnitExcliqlite
' isMineralGenericNameExcliqlite
' isValidImageFile
' isStringNoSemicolon
' isExcliqLiteAddin
' ================= CLASS EXCLIQ LITE RESOURCES REGEX STRUCTURE ==================== '



' ================= CLASS EXCLIQ LITE RESOURCES REGEX STRUCTURE ==================== '
' OBJECT VARIABLES (GLOOBJ_) ------------------------------------------------------- '
Private REGEXPROP As Object ' As New RegExp
Private EHGLOBAL As ClassEqlErrorHandler



' PROPERTIES ----------------------------------------------------------------------- '
Public Property Set ErrorHandler(ByRef ObjEH As ClassEqlErrorHandler)

  On Error GoTo EH
  Set EHGLOBAL = ObjEH

EH:
  Call EHGLOBAL.ErrorHandlerRaise("REGEX::ErrorHandler")

End Property



' CONSTRUCTOR/DESTRUCTOR ----------------------------------------------------------- '
Private Sub Class_Initialize()

  Set REGEXPROP = VBA.CreateObject("VBScript.RegExp")
  With REGEXPROP
    Let .Global = True
    Let .MultiLine = True
    Let .IgnoreCase = True
  End With

End Sub

Private Sub Class_Terminate()

  Set REGEXPROP = Nothing

End Sub



' CORE METHODS LIST (PRIVATE) ------------------------------------------------------ '
' ERROR ASSERTSIONS ---------------------------------------------------------------- '
' Debug_asset
Private Sub Debug_assert(ByVal Cond As Boolean, Optional Origin As String = VBA.vbNullString)

  #If Debbuging Then
    Debug.Assert Cond
  #End If

End Sub



' CORE METHODS LIST (PUBLIC) ------------------------------------------------------- '
Public Function isNumberEntireVectorCommaSeparated(ByVal Str As String) As Boolean

  On Error GoTo EH
  Let REGEXPROP.Pattern = "^[-?0-9]+(-?[0-9,]){3,}-?[0-9]$"
  Let isNumberEntireVectorCommaSeparated = REGEXPROP.test(Str)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("REGEX::isNumberEntireVectorCommaSeparated")

End Function

Public Function isNumberEntireAndDecimalVectorColonSeparated(ByVal Str As String) As Boolean

  On Error GoTo EH
  Let REGEXPROP.Pattern = "^((-?\d([.,]\d)?)+;?)+|(-?\d([.,]\d)?)$" '"^((-?[0-9](,[0-9])?)+;?)+(-?[0-9](,[0-9])?)$"
  Let isNumberEntireAndDecimalVectorColonSeparated = REGEXPROP.test(Str)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("REGEX::isNumberEntireAndDecimalVectorColonSeparated")

End Function

Public Function isNumberWithDecimalsSemicolonAndSeparatedSlashGrouped(ByVal Str As String) As Boolean

  On Error GoTo EH
  Let REGEXPROP.Pattern = "^((-?\d([.,]\d)?)+[/;]?)+|(-?\d([.,]\d*)?)$"
  Let isNumberWithDecimalsSemicolonAndSeparatedSlashGrouped = REGEXPROP.test(Str)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("REGEX::isNumberWithDecimalsSemicolonAndSeparatedSlashGrouped")

End Function

Public Function isNumberVector(ByVal Str As String) As Boolean

  On Error GoTo EH
  Let REGEXPROP.Pattern = "^[-?0-9]+(-?[0-9,]){3,}-?[0-9]$"
  Let isNumberVector = REGEXPROP.test(Str)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("REGEX::isNumberVector")

End Function

Public Function isNumberMinusReg(ByVal Str As String) As Boolean

  On Error GoTo EH
  Let REGEXPROP.Pattern = "(^[0-9]*(,?)[0-9]+$)|(^-?[0-9]*(.?)[0-9]+$)"
  Let isNumberMinusReg = REGEXPROP.test(Str)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("REGEX::isNumberMinusReg")

End Function



' PREDEFINED METHODS
Public Function isEmptyStringReg(ByVal Str As String) As Boolean

  On Error GoTo EH
  Let REGEXPROP.Pattern = "\S+"
  Let isEmptyStringReg = Not REGEXPROP.test(Str)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("REGEX::isEmptyStringReg")

End Function

Public Function isStringReg(ByVal Str As String, Optional Str2 As String = VBA.vbNullString) As Boolean

  '*MODIFIED
  On Error GoTo EH
  Let REGEXPROP.Pattern = "\w" & IIf(Str2 = VBA.vbNullString, VBA.vbNullString, "|" & Str2 & "+")
  Let isStringReg = REGEXPROP.test(Str)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("REGEX::isStringReg")

End Function

Public Function isNumberReg(ByVal Str As String) As Boolean

  On Error GoTo EH
  Let REGEXPROP.Pattern = "(^[-+]?[0-9]*(,?)[0-9]+$)|(^[-+]?[0-9]*(.?)[0-9]+$)" ' "(^[0-9]*(,?)[0-9]+$)|(^[0-9]*(.?)[0-9]+$)" '"\d"
  Let isNumberReg = REGEXPROP.test(Str)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("REGEX::isNumberReg")

End Function

Public Function isPercentReg(ByVal Str As String) As Boolean

  On Error GoTo EH
  Let REGEXPROP.Pattern = "(^[0-9]{1,3})(([,.][0-9]{1,2})?%?)$" '"(^[0-9]{1,3})(([,.][0-9]{1,2})?)$"
  Let isPercentReg = REGEXPROP.test(Str)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("REGEX::isPercentReg")

End Function

Public Function isDateReg(ByVal Str As String) As Boolean

  On Error GoTo EH
  Let REGEXPROP.Pattern = "(^[0-9]{1,2}-[0-9]{1,2}-([0-9]{2}|[0-9]{4})$)|(^[0-9]{1,2}/[0-9]{1,2}/([0-9]{2}|[0-9]{4})$)"
  Let isDateReg = REGEXPROP.test(Str)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("REGEX::isDateReg")

End Function

Public Function isRange(ByVal Str As String) As Boolean
  '^(.+\!)?(\$?[A-Z]+\$?[0-9]+)$|^(.+\!)?(\$?[A-Z]+\$?[0-9]+)([\:\;\,]\$?[A-Z]+\$?[0-9]+)+
  '[libro1]hoja1!$R$4;$R$2:$R$2
  '$R$4;$R$2:$R$2
  '$R$4hoja1!$R$4;$R$2:$R$2
  '$R$4;$R$2
  '$O$3
  '$R$4:$R$
  '$R$4:$
  '$R$4:$R
  '$R$4:
  '$R$4
  '$R$
  'Hoja1
  'Hoja1!
  '$R$4:$R$2
  'Hoja1!$O$3
  On Error GoTo EH
  Let isRange = CustomRegex(Str, "^(.+\!)?(\$?[A-Z]+\$?[0-9]+)$|^(.+\!)?(\$?[A-Z]+\$?[0-9]+)([\:\;\,]\$?[A-Z]+\$?[0-9]+)+", True, False, True)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("REGEX::isRange")

End Function

Public Function isRangeWithName(ByVal Str As String) As Boolean

  On Error GoTo EH
  Let REGEXPROP.Pattern = "^([\W\w]+(\!){1})"
  Let isRangeWithName = REGEXPROP.test(Str)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("REGEX::isRangeWithName")

End Function

Public Function isLocalPath(ByVal Str As String) As Boolean

  '^(\w:\\){1}(.*)[\\\w]$ 'Local Drive
  On Error GoTo EH
  Let REGEXPROP.Pattern = "^(\w:\\){1}(.*)[\\\w]$"
  Let isLocalPath = REGEXPROP.test(Str)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("REGEX::isLocalPath")

End Function

Public Function isNetworkPath(ByVal Str As String) As Boolean

  '^(\\){2}(.*)[\\\w]$ 'Red
  On Error GoTo EH
  Let REGEXPROP.Pattern = "^(\\){2}(.*)[\\\w]$"
  Let isNetworkPath = REGEXPROP.test(Str)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("REGEX::isNetworkPath")

End Function

Public Function isPhonenumber(ByVal Str As String) As Boolean

  Dim xPattern As String

  On Error GoTo EH
  Let xPattern = "^(?:(?:\(?(?:00|\+)([1-4]\d\d|[1-9]\d?)\)?)?[\-\.\ \\\/]?)?((?:\(?\d{1,}\)?[\-\.\ \\\/]?){0,})(?:[\-\.\ \\\/]?(?:#|ext\.?|extension|x)[\-\.\ \\\/]?(\d+))?$"
  Let isPhonenumber = Me.CustomRegex(Str, xPattern, i:=True)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("REGEX::isPhonenumber")

End Function

Public Function isEmail(ByVal Str As String) As Boolean

  Dim xPattern As String

  On Error GoTo EH
  Let xPattern = "^[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?$"
  Let isEmail = Me.CustomRegex(Str, xPattern, True, i:=True)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("REGEX::isEmail")

End Function

Public Function isWebpage(ByVal Str As String) As Boolean

  Dim xPattern As String
  'Let xPattern = "^\b(?:(?:https?|ftp):\/\/|www\.)?[-a-z0-9+&@#\/%?=~_|!:,.;]*[-a-z0-9+&@#\/%=~_|]"

  On Error GoTo EH
  Let xPattern = "^(((http|https|ftp)?(\:{1}\/{2}))?(w{3}\.{1})?)?([-a-z0-9+&@#\/%?=~_!:])+(\.[-a-z0-9+&@#\/%?=~_!:]+)$"
  Let isWebpage = Me.CustomRegex(Str, xPattern, i:=True)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("REGEX::isWebpage")

End Function

Public Function isNotavaiable(ByVal Str As String) As Boolean

  Dim xPattern As String

  On Error GoTo EH
  Let xPattern = "^(N\/D)+.*"
  Let isNotavaiable = Me.CustomRegex(Str, xPattern, True, i:=True)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("REGEX::isNotavaiable")

End Function

Public Function isMulticurrency(ByVal Str As String) As Boolean

  Dim xPattern As String

  On Error GoTo EH
  Let xPattern = "[\w\W]+(es|s)"
  Let isMulticurrency = Me.CustomRegex(Str, xPattern, True, i:=True)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("REGEX::isMulticurrency")

End Function

Public Function isRekey(ByVal Str As String) As Boolean

  'TODO: ELIMINAR DE EXCLIQ, SÓLO FUNCIONA EN EXCLIQREG
  Dim xPattern As String

  On Error GoTo EH
  Let xPattern = "^(\w{5,})?(\-\w{1,3})+-EXCLIQ-(\w{5,})?$"
  Let isRekey = Me.CustomRegex(Str, xPattern, True, i:=True)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("REGEX::isRekey")

End Function

Public Function isFinalkey(ByVal Str As String) As Boolean

  Dim xPattern As String
  'Let xPattern = "^[\w:;>=<?@\]\[\\\^\`\{\}\|\~\Ç\ü\é]{5}&([\w:;>=<?@\]\[\\\^\`\{\}\|\~\Ç\ü\é]{2,3}-){4}([\w:;>=<?@\]\[\\\^\`\{\}\|\~\Ç\ü\é]{2,3})&([\w:;>=<?@\]\[\\\^\`\{\}\|\~\Ç\ü\é]{2,3}-){7}([\w:;>=<?@\]\[\\\^\`\{\}\|\~\Ç\ü\é]{2,3})&([\w:;>=<?@\]\[\\\^\`\{\}\|\~\Ç\ü\é]{2,3}-)+([\w:;>=<?@\]\[\\\^\`\{\}\|\~\Ç\ü\é]{2,3})$"

  On Error GoTo EH
  Let xPattern = "^([A-Z@]+&[A-Z@]+&[A-Z@]+&[A-Z@]+)$"
  Let isFinalkey = Me.CustomRegex(Str, xPattern, True, i:=True)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("REGEX::isFinalkey")

End Function

Public Function isOreItem(ByVal Str As String) As Boolean

  Dim xPattern As String

  On Error GoTo EH
  Let xPattern = "^((\d{2,10}).(\d{2,10}).(\d{2,10}).(\d{2,10})){1}$"
  Let isOreItem = Me.CustomRegex(Str, xPattern, True, i:=True)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("REGEX::isOreItem")

End Function

Public Function isThreeDigitFloatNumber(ByVal Str As String) As Boolean

  Dim xPattern As String

  On Error GoTo EH
  Let xPattern = "^(1?\d{1,2}%?)$|^(1?\d{1,2})([,.])(\d{1,2}%?)$"
  Let isThreeDigitFloatNumber = Me.CustomRegex(Str, xPattern, True, i:=True)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("REGEX::isThreeDigitFloatNumber")

End Function

'CUSTOM METHOD
Public Function CustomRegex(ByVal Value As Variant, ByVal StrPattern As String, Optional G As Boolean, Optional m As Boolean, Optional i As Boolean) As Boolean

  On Error GoTo EH
  With REGEXPROP
    Let .Global = G
    Let .MultiLine = m
    Let .IgnoreCase = i
    Let .Pattern = StrPattern
  End With
  Let CustomRegex = REGEXPROP.test(Value)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("REGEX::CustomRegex")

End Function



' NEW METHODS IN BOTH PROJECTS
Public Function isProviderType(ByVal Str As String) As Boolean

  Dim StrPattern As String

  On Error GoTo EH
  Let StrPattern = "^(cooperativista|particular)$" '"^particular$|^cooperativista$"
  Let isProviderType = Me.CustomRegex(Str, StrPattern, True, False, True)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("REGEX::isProviderType")

End Function

Public Function isProvidersTypes(ByVal Str As String) As Boolean

  Dim StrPattern As String

  On Error GoTo EH
  Let StrPattern = "^((cooperativista|particular);?)+(cooperativista|particular)$"
  Let isProvidersTypes = Me.CustomRegex(Str, StrPattern, True, False, True)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("REGEX::isProviderType")

End Function

Public Function isMineralType(ByVal Str As String) As Boolean

  Dim StrPattern As String

  On Error GoTo EH
  Let StrPattern = "^mineral(es)?\s?(complejo(s?))?$|^complejo(s?)$|^bro[zs]a(s?)$"
  Let isMineralType = Me.CustomRegex(Str, StrPattern, True, False, True)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("REGEX::isMineralType")

End Function

Public Function isMineralComplex(ByVal Str As String) As Boolean

  Dim StrPattern As String

  On Error GoTo EH
  Let StrPattern = "((^mineral(es)?\s?)?complejo(s)?)$"
  Let isMineralComplex = Me.CustomRegex(Str, StrPattern, True, False, True)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("REGEX::isMineralComplex")

End Function

Public Function isMineralBrosa(ByVal Str As String) As Boolean

  Dim StrPattern As String

  On Error GoTo EH
  Let StrPattern = "^bro[sz]a(s?)$"
  Let isMineralBrosa = Me.CustomRegex(Str, StrPattern, True, False, True)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("REGEX::isMineralBrosa")

End Function

Public Function isMineralExcliqlite(ByVal Str As String) As Boolean

  Dim StrPattern As String

  On Error GoTo EH
  Let StrPattern = "^pb$|^ag$|^zn$|^cu$|^au$|^sn$|^sb$|^co$|^bi$|^al$|^ta$|^fe$|^ni$|^pt$|^pa$"
  Let isMineralExcliqlite = Me.CustomRegex(Str, StrPattern, True, False, True)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("REGEX::isMineralExcliqlite")

End Function

Public Function isMineralUnitExcliqlite(ByVal Str As String) As Boolean

  Dim StrPattern As String

  On Error GoTo EH
  Let StrPattern = "^\%$|^dm$|^ot$|^g\/t$|^g\/l$"
  Let isMineralUnitExcliqlite = Me.CustomRegex(Str, StrPattern, True, False, True)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("REGEX::isMineralUnitExcliqlite")

End Function

Public Function isMineralGenericNameExcliqlite(ByVal Str As String) As Boolean

  Dim StrPattern As String

  On Error GoTo EH
  Let StrPattern = "^b[0-9]+$|^bx[0-9]+$"
  Let isMineralGenericNameExcliqlite = Me.CustomRegex(Str, StrPattern, True, False, True)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("REGEX::isMineralGenericNameExcliqlite")

End Function

Public Function isValidImageFile(ByVal Str As String) As Boolean

  Dim StrPattern As String

  On Error GoTo EH
  Let StrPattern = ".+(\.jpg|\.png|\.gif|\.bmp|\.tiff)$"
  Let isValidImageFile = Me.CustomRegex(Str, StrPattern, True, False, True)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("REGEX::isValidImageFile")

End Function

Public Function isStringNoSemicolon(ByVal Str As String) As Boolean

  On Error GoTo EH
  Let isStringNoSemicolon = False
  If Me.isEmptyStringReg(Str) Then GoTo EH
  Let REGEXPROP.Pattern = "\;+"
  If REGEXPROP.test(Str) Then GoTo EH
  Let REGEXPROP.Pattern = "\w"
  Let isStringNoSemicolon = REGEXPROP.test(Str)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("REGEX::isStringNoSemicolon")

End Function

Public Function isExcliqLiteAddin(ByVal AStrMDLEQName As String, ByVal Str As String) As Boolean

  Dim StrPattern As String

  On Error GoTo EH
  Let StrPattern = "^" & AStrMDLEQName & "((_v[_]?)?(\d+.\d+.\d+)?)$"
  Let isExcliqLiteAddin = Me.CustomRegex(Str, StrPattern, True, False, True)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("REGEX::isExcliqLiteAddin")

End Function
' CORE METHODS LIST (PUBLIC) ------------------------------------------------------- '
' ================= CLASS EXCLIQ LITE RESOURCES REGEX STRUCTURE ==================== '

