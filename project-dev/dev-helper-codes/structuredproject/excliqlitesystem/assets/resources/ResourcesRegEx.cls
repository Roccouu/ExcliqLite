VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppResRegEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private REGEXPROP As New RegExp


'CONSTRUCTOR/DESTRUCTOR
Private Sub Class_Initialize()
  With REGEXPROP
    Let .Global = True
    Let .MultiLine = True
    Let .IgnoreCase = True
  End With
End Sub

Private Sub Class_Terminate()
  Set REGEXPROP = Nothing
End Sub


'METHODS PRIVATE
Private Sub Debug_assert(ByVal Cond As Boolean, Optional Origin As String = VBA.vbNullString)
  #If Debbuging Then
    Debug.Assert Cond
  #End If
End Sub


'LITE METHODS
Public Function isNumberEntireVectorCommaSeparated(ByVal Str As String) As Boolean
  On Error GoTo EH
  Let REGEXPROP.Pattern = "^[-?0-9]+(-?[0-9,]){3,}-?[0-9]$"
  Let isNumberEntireVectorCommaSeparated = REGEXPROP.test(Str)

EH:
  On Error GoTo -1
  Call VBA.Err.Clear
End Function

Public Function isNumberEntireAndDecimalVectorColonSeparated(ByVal Str As String) As Boolean
  On Error GoTo EH
  Let REGEXPROP.Pattern = "^((-?\d([.,]\d)?)+;?)+|(-?\d([.,]\d)?)$" '"^((-?[0-9](,[0-9])?)+;?)+(-?[0-9](,[0-9])?)$"
  Let isNumberEntireAndDecimalVectorColonSeparated = REGEXPROP.test(Str)

EH:
  On Error GoTo -1
  Call VBA.Err.Clear
End Function

Public Function isNumberWithDecimalsSemicolonAndSeparatedSlashGrouped(ByVal Str As String) As Boolean
  On Error GoTo EH
  Let REGEXPROP.Pattern = "^((-?\d([.,]\d)?)+[/;]?)+|(-?\d([.,]\d*)?)$"
  Let isNumberWithDecimalsSemicolonAndSeparatedSlashGrouped = REGEXPROP.test(Str)

EH:
  On Error GoTo -1
  Call VBA.Err.Clear
End Function

Public Function isNumberVector(ByVal Str As String) As Boolean
  On Error GoTo EH
  Let REGEXPROP.Pattern = "^[-?0-9]+(-?[0-9,]){3,}-?[0-9]$"
  Let isNumberVector = REGEXPROP.test(Str)

EH:
  On Error GoTo -1
  Call VBA.Err.Clear
End Function

Public Function isNumberMinusReg(ByVal Str As String) As Boolean
  On Error GoTo EH
  Let REGEXPROP.Pattern = "(^[0-9]*(,?)[0-9]+$)|(^-?[0-9]*(.?)[0-9]+$)"
  Let isNumberMinusReg = REGEXPROP.test(Str)

EH:
  On Error GoTo -1
  Call VBA.Err.Clear
End Function



'PREDEFINED METHODS
Public Function isEmptyStringReg(ByVal Str As String) As Boolean
  On Error GoTo EH
  Let REGEXPROP.Pattern = "\S+"
  Let isEmptyStringReg = Not REGEXPROP.test(Str)

EH:
  On Error GoTo -1
  Call VBA.Err.Clear
End Function

Public Function isStringReg(ByVal Str As String, Optional Str2 As String = VBA.vbNullString) As Boolean
  '*MODIFIED
  On Error GoTo EH
  Let REGEXPROP.Pattern = "\w" & IIf(Str2 = VBA.vbNullString, VBA.vbNullString, "|" & Str2 & "+")
  Let isStringReg = REGEXPROP.test(Str)

EH:
  On Error GoTo -1
  Call VBA.Err.Clear
End Function

Public Function isNumberReg(ByVal Str As String) As Boolean
  On Error GoTo EH
  Let REGEXPROP.Pattern = "(^[0-9]*(,?)[0-9]+$)|(^[0-9]*(.?)[0-9]+$)" '"\d"
  Let isNumberReg = REGEXPROP.test(Str)

EH:
  On Error GoTo -1
  Call VBA.Err.Clear
End Function

Public Function isPercentReg(ByVal Str As String) As Boolean
  On Error GoTo EH
  Let REGEXPROP.Pattern = "(^[0-9]{1,3})(([,.][0-9]{1,2})?%?)$" '"(^[0-9]{1,3})(([,.][0-9]{1,2})?)$"
  Let isPercentReg = REGEXPROP.test(Str)

EH:
  On Error GoTo -1
  Call VBA.Err.Clear
End Function

Public Function isDateReg(ByVal Str As String) As Boolean
  On Error GoTo EH
  Let REGEXPROP.Pattern = "(^[0-9]{1,2}-[0-9]{1,2}-([0-9]{2}|[0-9]{4})$)|(^[0-9]{1,2}/[0-9]{1,2}/([0-9]{2}|[0-9]{4})$)"
  Let isDateReg = REGEXPROP.test(Str)

EH:
  On Error GoTo -1
  Call VBA.Err.Clear
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
  
  Let isRange = CustomRegex(Str, "^(.+\!)?(\$?[A-Z]+\$?[0-9]+)$|^(.+\!)?(\$?[A-Z]+\$?[0-9]+)([\:\;\,]\$?[A-Z]+\$?[0-9]+)+", True, False, True)
End Function

Public Function isRangeWithName(ByVal Str As String) As Boolean
  Let REGEXPROP.Pattern = "^([\W\w]+(\!){1})"
  Let isRangeWithName = REGEXPROP.test(Str)
End Function

Public Function isLocalPath(ByVal Str As String) As Boolean
  '^(\w:\\){1}(.*)[\\\w]$ 'Local Drive
  On Error GoTo EH
  Let REGEXPROP.Pattern = "^(\w:\\){1}(.*)[\\\w]$"
  Let isLocalPath = REGEXPROP.test(Str)

EH:
  On Error GoTo -1
  Call VBA.Err.Clear
End Function

Public Function isNetworkPath(ByVal Str As String) As Boolean
  '^(\\){2}(.*)[\\\w]$ 'Red
  On Error GoTo EH
  Let REGEXPROP.Pattern = "^(\\){2}(.*)[\\\w]$"
  Let isNetworkPath = REGEXPROP.test(Str)

EH:
  On Error GoTo -1
  Call VBA.Err.Clear
End Function

Public Function isPhonenumber(ByVal Str As String) As Boolean
  Dim xPattern As String
  
  On Error GoTo EH
  Let xPattern = "^(?:(?:\(?(?:00|\+)([1-4]\d\d|[1-9]\d?)\)?)?[\-\.\ \\\/]?)?((?:\(?\d{1,}\)?[\-\.\ \\\/]?){0,})(?:[\-\.\ \\\/]?(?:#|ext\.?|extension|x)[\-\.\ \\\/]?(\d+))?$"
  Let isPhonenumber = Me.CustomRegex(Str, xPattern, i:=True)

EH:
  On Error GoTo -1
  Call VBA.Err.Clear
End Function

Public Function isEmail(ByVal Str As String) As Boolean
  Dim xPattern As String
  
  On Error GoTo EH
  Let xPattern = "^[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?$"
  Let isEmail = Me.CustomRegex(Str, xPattern, True, i:=True)

EH:
  On Error GoTo -1
  Call VBA.Err.Clear
End Function

Public Function isWebpage(ByVal Str As String) As Boolean
  Dim xPattern As String
  'Let xPattern = "^\b(?:(?:https?|ftp):\/\/|www\.)?[-a-z0-9+&@#\/%?=~_|!:,.;]*[-a-z0-9+&@#\/%=~_|]"
  
  On Error GoTo EH
  Let xPattern = "^(((http|https|ftp)?(\:{1}\/{2}))?(w{3}\.{1})?)?([-a-z0-9+&@#\/%?=~_!:])+(\.[-a-z0-9+&@#\/%?=~_!:]+)$"
  Let isWebpage = Me.CustomRegex(Str, xPattern, i:=True)

EH:
  On Error GoTo -1
  Call VBA.Err.Clear
End Function

Public Function isNotavaiable(ByVal Str As String) As Boolean
  Dim xPattern As String
  
  On Error GoTo EH
  Let xPattern = "^(N\/D)+.*"
  Let isNotavaiable = Me.CustomRegex(Str, xPattern, True, i:=True)

EH:
  On Error GoTo -1
  Call VBA.Err.Clear
End Function

Public Function isMulticurrency(ByVal Str As String) As Boolean
  Dim xPattern As String
  
  On Error GoTo EH
  Let xPattern = "[\w\W]+(es|s)"
  Let isMulticurrency = Me.CustomRegex(Str, xPattern, True, i:=True)

EH:
  On Error GoTo -1
  Call VBA.Err.Clear
End Function

Public Function isRekey(ByVal Str As String) As Boolean
  'TODO: ELIMINAR DE EXCLIQ, SÓLO FUNCIONA EN EXCLIQREG
  Dim xPattern As String
  
  On Error GoTo EH
  Let xPattern = "^(\w{5,})?(\-\w{1,3})+-EXCLIQ-(\w{5,})?$"
  Let isRekey = Me.CustomRegex(Str, xPattern, True, i:=True)

EH:
  On Error GoTo -1
  Call VBA.Err.Clear
End Function

Public Function isFinalkey(ByVal Str As String) As Boolean
  Dim xPattern As String
  'Let xPattern = "^[\w:;>=<?@\]\[\\\^\`\{\}\|\~\Ç\ü\é]{5}&([\w:;>=<?@\]\[\\\^\`\{\}\|\~\Ç\ü\é]{2,3}-){4}([\w:;>=<?@\]\[\\\^\`\{\}\|\~\Ç\ü\é]{2,3})&([\w:;>=<?@\]\[\\\^\`\{\}\|\~\Ç\ü\é]{2,3}-){7}([\w:;>=<?@\]\[\\\^\`\{\}\|\~\Ç\ü\é]{2,3})&([\w:;>=<?@\]\[\\\^\`\{\}\|\~\Ç\ü\é]{2,3}-)+([\w:;>=<?@\]\[\\\^\`\{\}\|\~\Ç\ü\é]{2,3})$"
  
  On Error GoTo EH
  Let xPattern = "^([A-Z@]+&[A-Z@]+&[A-Z@]+&[A-Z@]+)$"
  Let isFinalkey = Me.CustomRegex(Str, xPattern, True, i:=True)

EH:
  On Error GoTo -1
  Call VBA.Err.Clear
End Function

Public Function isOreItem(ByVal Str As String) As Boolean
  Dim xPattern As String
  
  On Error GoTo EH
  Let xPattern = "^((\d{2,10}).(\d{2,10}).(\d{2,10}).(\d{2,10})){1}$"
  Let isOreItem = Me.CustomRegex(Str, xPattern, True, i:=True)

EH:
  On Error GoTo -1
  Call VBA.Err.Clear
End Function

Public Function isThreeDigitFloatNumber(ByVal Str As String) As Boolean
  Dim xPattern As String
  
  On Error GoTo EH
  Let xPattern = "^(1?\d{1,2}%?)$|^(1?\d{1,2})([,.])(\d{1,2}%?)$"
  Let isThreeDigitFloatNumber = Me.CustomRegex(Str, xPattern, True, i:=True)

EH:
  On Error GoTo -1
  Call VBA.Err.Clear
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
  On Error GoTo -1
  Call VBA.Err.Clear
End Function



'NEW METHODS IN BOTH PROJECTS
'Public Function isStringWithCustom(ByVal Str As String, Optional Str2 As String = VBA.vbNullString) As Boolean
'  Let REGEXPROP.Pattern = "\w+" & IIf(Str2 = VBA.vbNullString, VBA.vbNullString, "|(" & Str2 & ")")
'  Let isStringReg = REGEXPROP.test(Str)
'End Function

Public Function isProviderType(ByVal Str As String) As Boolean
  Dim StrPattern As String
  
  On Error GoTo EH
  Let StrPattern = "^particular$|^cooperativista$"
  Let isProviderType = Me.CustomRegex(Str, StrPattern, True, False, True)

EH:
  On Error GoTo -1
  Call VBA.Err.Clear
End Function

Public Function isMineralType(ByVal Str As String) As Boolean
  Dim StrPattern As String
  
  On Error GoTo EH
  Let StrPattern = "^mineral(es)?\s?(complejo(s?))?$|^complejo(s?)$|^bro[zs]a(s?)$"
  Let isMineralType = Me.CustomRegex(Str, StrPattern, True, False, True)

EH:
  On Error GoTo -1
  Call VBA.Err.Clear
End Function

Public Function isMineralComplex(ByVal Str As String) As Boolean
  Dim StrPattern As String
  
  On Error GoTo EH
  Let StrPattern = "((^mineral(es)?\s?)?complejo(s)?)$"
  Let isMineralComplex = Me.CustomRegex(Str, StrPattern, True, False, True)

EH:
  On Error GoTo -1
  Call VBA.Err.Clear
End Function

Public Function isMineralBrosa(ByVal Str As String) As Boolean
  Dim StrPattern As String
  
  On Error GoTo EH
  Let StrPattern = "^bro[sz]a(s?)$"
  Let isMineralBrosa = Me.CustomRegex(Str, StrPattern, True, False, True)

EH:
  On Error GoTo -1
  Call VBA.Err.Clear
End Function

Public Function isMineralExcliqlite(ByVal Str As String) As Boolean
  Dim StrPattern As String
  
  On Error GoTo EH
  Let StrPattern = "^pb$|^ag$|^zn$|^cu$|^au$|^sn$|^sb$|^co$|^bi$|^al$|^ta$|^fe$|^ni$|^pt$|^pa$"
  Let isMineralExcliqlite = Me.CustomRegex(Str, StrPattern, True, False, True)

EH:
  On Error GoTo -1
  Call VBA.Err.Clear
End Function

Public Function isMineralUnitExcliqlite(ByVal Str As String) As Boolean
  Dim StrPattern As String
  
  On Error GoTo EH
  Let StrPattern = "^\%$|^dm$|^ot$|^g\/t$|^g\/l$"
  Let isMineralUnitExcliqlite = Me.CustomRegex(Str, StrPattern, True, False, True)

EH:
  On Error GoTo -1
  Call VBA.Err.Clear
End Function































'Option Explicit
'
'Private REGEXPROP As New RegExp
'
'Private Sub Class_Initialize()
'  With REGEXPROP
'    .Global = True
'    .MultiLine = True
'    .IgnoreCase = True
'  End With
'End Sub
'
''PREDEFINED METHODS
'Public Function isEmptyStringReg(ByVal str As String) As Boolean
'  REGEXPROP.Pattern = "\S"
'  isEmptyStringReg = Not REGEXPROP.test(str)
'End Function
'Public Function isStringReg(ByVal str As String) As Boolean
'  REGEXPROP.Pattern = "\w"
'  isStringReg = REGEXPROP.test(str)
'End Function
'Public Function isNumberReg(ByVal str As String) As Boolean
'  REGEXPROP.Pattern = "\d"
'  isNumberReg = REGEXPROP.test(str)
'End Function
'Public Function isRange(ByVal str As String) As Boolean
'  '^(.+\!)?(\$?[A-Z]+\$?[0-9]+)$|^(.+\!)?(\$?[A-Z]+\$?[0-9]+)([\:\;\,]\$?[A-Z]+\$?[0-9]+)+
'  '[libro1]hoja1!$R$4;$R$2:$R$2
'  '$R$4;$R$2:$R$2
'  '$R$4hoja1!$R$4;$R$2:$R$2
'  '$R$4;$R$2
'  '$O$3
'  '$R$4:$R$
'  '$R$4:$
'  '$R$4:$R
'  '$R$4:
'  '$R$4
'  '$R$
'  'Hoja1
'  'Hoja1!
'  '$R$4:$R$2
'  'Hoja1!$O$3
'
'  isRange = customRegex(str, "^(.+\!)?(\$?[A-Z]+\$?[0-9]+)$|^(.+\!)?(\$?[A-Z]+\$?[0-9]+)([\:\;\,]\$?[A-Z]+\$?[0-9]+)+", True, False, True)
'End Function
'Public Function isRangeWithName(ByVal str As String) As Boolean
'  REGEXPROP.Pattern = "^([\W\w]+(\!){1})"
'  Let isRangeWithName = REGEXPROP.test(str)
'End Function
'
''CUSTOM METHOD
'Public Function customRegex(ByVal Value As Variant, ByVal strPattern As String, G As Boolean, m As Boolean, i As Boolean) As Boolean
'  With REGEXPROP
'    .Global = G
'    .MultiLine = m
'    .IgnoreCase = i
'    .Pattern = strPattern
'  End With
'  customRegex = REGEXPROP.test(Value)
'End Function
