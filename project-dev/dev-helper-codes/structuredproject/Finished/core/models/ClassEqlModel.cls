VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassEqlModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================= CLASS EXCLIQ LITE MODEL STRUCTURE ======================== '
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


' ======================= CLASS EXCLIQ LITE MODEL STRUCTURE ======================== '
' OBJECT VARIABLES (GLOOBJ_) ------------------------------------------------------- '
' VARIABLES (GLOStr_) -------------------------------------------------------------- '
' PROPERTIES ----------------------------------------------------------------------- '
' CONSTRUCTOR/DESTRUCTOR ----------------------------------------------------------- '
' CORE METHODS LIST (PRIVATE) ------------------------------------------------------ '
' Model_start
' ERROR ASSERTSIONS ---------------------------------------------------------------- '
' Debug_assert
' CORE METHODS LIST (PUBLIC) ------------------------------------------------------- '
' MSet
' MGet
' MUpd
' MDel
' ======================= CLASS EXCLIQ LITE MODEL STRUCTURE ======================== '



' ======================= CLASS EXCLIQ LITE MODEL STRUCTURE ======================== '
' OBJECT VARIABLES (GLOOBJ_) ------------------------------------------------------- '
Private RES As ClassEqlResources
Private GLOSHT_DATA As Worksheet

' VARIABLES (GLOStr_) -------------------------------------------------------------- '
Private GLOVEC_VERSIONDATA As Variant





' PROPERTIES ----------------------------------------------------------------------- '
Public Property Set ResourcesApp(ByRef ObjRES As ClassEqlResources)

  Set RES = ObjRES
  Let GLOVEC_VERSIONDATA = Me.MGet(eqlMdlSheetSys, eqlMdlTblConfigapp_sys, eqlMdlArray, MLngIndex:=1)

End Property

Public Property Get LengthRows(ByVal MEnuSheet As EQLMDL_ENU_SHEETS, ByVal MEnuTable As EQLMDL_ENU_TABLES) As Long

  Dim MPT As ClassEqlModelParent

  On Error GoTo EH
  Call Model_start(MPT, MEnuSheet, MEnuTable)
  Let LengthRows = MPT.LengthRows(GLOSHT_DATA.ListObjects(MEnuTable))

EH:
  Set MPT = Nothing
  Call RES.ErrorHandler.ErrorHandlerRaise("MDL::LengthRows")

End Property

Public Property Get LengthCols(ByVal MEnuSheet As EQLMDL_ENU_SHEETS, ByVal MEnuTable As EQLMDL_ENU_TABLES) As Long

  Dim MPT As ClassEqlModelParent

  On Error GoTo EH
  Call Model_start(MPT, MEnuSheet, MEnuTable)
  Let LengthCols = MPT.LengthCols(GLOSHT_DATA.ListObjects(MEnuTable))

EH:
  Set MPT = Nothing
  Call RES.ErrorHandler.ErrorHandlerRaise("MDL::LengthCols")

End Property

' App data
Public Property Get AppTitle() As String
  Let AppTitle = VBA.CStr(GLOVEC_VERSIONDATA(0, 21))
End Property

Public Property Get AppVersion() As String
  Let AppVersion = VBA.CStr(GLOVEC_VERSIONDATA(0, 22)) & "." & VBA.CStr(GLOVEC_VERSIONDATA(0, 23)) & "." & VBA.CStr(GLOVEC_VERSIONDATA(0, 24))
End Property

Public Property Get AppAuthor() As String
  Let AppAuthor = VBA.CStr(GLOVEC_VERSIONDATA(0, 25))
End Property

Public Property Get AppEmail() As String
  Let AppEmail = VBA.CStr(GLOVEC_VERSIONDATA(0, 26))
End Property

Public Property Get AppSlogan() As String
  Let AppSlogan = VBA.CStr(GLOVEC_VERSIONDATA(0, 27))
End Property

Public Property Get AppYear() As String
  Let AppYear = VBA.CStr(GLOVEC_VERSIONDATA(0, 28))
End Property

Public Property Get AppOrigin() As String
  Let AppOrigin = VBA.CStr(GLOVEC_VERSIONDATA(0, 29))
End Property

Public Property Get AppGit() As String
  Let AppGit = VBA.CStr(GLOVEC_VERSIONDATA(0, 30))
End Property

Public Property Get AppLicense() As String
  Let AppLicense = VBA.CStr(GLOVEC_VERSIONDATA(0, 31))
End Property

' DB tables - Ribbon options
Public Property Get Formulas() As Boolean
  Let Formulas = VBA.CBool(GLOVEC_VERSIONDATA(0, 0))
End Property

Public Property Let Formulas(ByVal xStatus As Boolean)
  Call Me.MUpd(eqlMdlSheetSys, eqlMdlTblConfigapp_sys, Array(IIf(xStatus, 1, 0)), "formulas", 1)
End Property

Public Property Get Tables() As Byte
  Let Tables = VBA.CByte(GLOVEC_VERSIONDATA(0, 1))
End Property

Public Property Let Tables(ByVal xTable As Byte) '0: vile table, 1: plain table, 2: table
  Call Me.MUpd(eqlMdlSheetSys, eqlMdlTblConfigapp_sys, Array(xTable), "table_type", 1)
End Property

Public Property Get TableStyles() As Byte
  Let TableStyles = VBA.CByte(GLOVEC_VERSIONDATA(0, 2))
End Property

Public Property Let TableStyles(ByVal xStyle As Byte) '0: Default grey eql style, 1: green eql style, 2: excel predefined pastel style
  Call Me.MUpd(eqlMdlSheetSys, eqlMdlTblConfigapp_sys, Array(xStyle), "table_style", 1)
End Property

Public Property Get TableThemeColor() As XlThemeColor ' Tables Themes: Returns the current EQSystem theme color using as base Table Style
  Select Case Me.TableStyles
    Case 0: Let TableThemeColor = xlThemeColorDark2 ' Grey ExcliqLite
    Case 1: Let TableThemeColor = xlThemeColorAccent6 ' 10 ' Green ExcliqLite
    Case Else: Let TableThemeColor = xlThemeColorAccent1 ' Blue XL
  End Select
End Property

Public Property Get TableThemeFontColor() As XlThemeColor ' Tables Themes: Returns the current EQSystem theme color using as base Table Style
  Select Case Me.TableStyles
    Case 0: Let TableThemeFontColor = xlThemeColorDark2 ' Black ExcliqLite
    Case 1: Let TableThemeFontColor = xlThemeColorDark1 ' White ExcliqLite
    Case Else: Let TableThemeFontColor = xlThemeColorDark1 ' White XL
  End Select
End Property

Public Property Get TableRowStripes() As Boolean
  Let TableRowStripes = VBA.CBool(GLOVEC_VERSIONDATA(0, 3))
End Property

Public Property Let TableRowStripes(ByVal xStripe As Boolean) 'TRUE: Stiped rows, FALSE: Simple rows
  Call Me.MUpd(eqlMdlSheetSys, eqlMdlTblConfigapp_sys, Array(IIf(xStripe, 1, 0)), "table_rowstripe", 1)
End Property

Public Property Get TableSelectionsCell() As Boolean
  Let TableSelectionsCell = VBA.CBool(GLOVEC_VERSIONDATA(0, 4))
End Property

Public Property Let TableSelectionsCell(ByVal MBooSelections As Boolean) ' TRUE: Show Inputbox 8 enabled, FALSE: Show Inputbox 8 disabled
  Call Me.MUpd(eqlMdlSheetSys, eqlMdlTblConfigapp_sys, Array(IIf(MBooSelections, 1, 0)), "table_selections_cell", 1)
End Property

' DB tables - Config options
Public Property Get TableSelectionsCompact() As Boolean
  Let TableSelectionsCompact = VBA.CBool(GLOVEC_VERSIONDATA(0, 5))
End Property

Public Property Let TableSelectionsCompact(ByVal MBooSelectionsCompact As Boolean) ' TRUE: To print a compact or extended Balmet
  Call Me.MUpd(eqlMdlSheetSys, eqlMdlTblConfigapp_sys, Array(IIf(MBooSelectionsCompact, 1, 0)), "table_selections_compact", 1)
End Property

Public Property Get TableSelectionsHeads() As Boolean
  Let TableSelectionsHeads = VBA.CBool(GLOVEC_VERSIONDATA(0, 6))
End Property

Public Property Let TableSelectionsHeads(ByVal MBooSelectionsHeads As Boolean) ' TRUE: To print a Feed heads on Balmet
  Call Me.MUpd(eqlMdlSheetSys, eqlMdlTblConfigapp_sys, Array(IIf(MBooSelectionsHeads, 1, 0)), "table_selections_heads", 1)
End Property

' DB tables - Purchase options
Public Property Get PurAverage() As Boolean
  Let PurAverage = VBA.CBool(GLOVEC_VERSIONDATA(0, 7))
End Property

Public Property Let PurAverage(ByVal MBooPurAverage As Boolean) ' TRUE: Average RM's and other datas anyway
  Call Me.MUpd(eqlMdlSheetSys, eqlMdlTblConfigapp_sys, Array(IIf(MBooPurAverage, 1, 0)), "pur_average", 1)
End Property

Public Property Get PurRMDetailed() As Boolean
  Let PurRMDetailed = VBA.CBool(GLOVEC_VERSIONDATA(0, 8))
End Property

Public Property Let PurRMDetailed(ByVal MBooPurRMDetailed As Boolean) ' TRUE: Detailed RM's
  Call Me.MUpd(eqlMdlSheetSys, eqlMdlTblConfigapp_sys, Array(IIf(MBooPurRMDetailed, 1, 0)), "pur_rmdetail", 1)
End Property

Public Property Get PurRetentions() As Boolean
  Let PurRetentions = VBA.CBool(GLOVEC_VERSIONDATA(0, 9))
End Property

Public Property Let PurRetentions(ByVal MBooPurRetentions As Boolean)
  Call Me.MUpd(eqlMdlSheetSys, eqlMdlTblConfigapp_sys, Array(IIf(MBooPurRetentions, 1, 0)), "pur_resumed_retentions", 1)
End Property

Public Property Get PurRetentionsOther() As Boolean
  Let PurRetentionsOther = VBA.CBool(GLOVEC_VERSIONDATA(0, 10))
End Property

Public Property Let PurRetentionsOther(ByVal MBooPurRetentionsOther As Boolean)
  Call Me.MUpd(eqlMdlSheetSys, eqlMdlTblConfigapp_sys, Array(IIf(MBooPurRetentionsOther, 1, 0)), "pur_resumed_retentionsother", 1)
End Property

Public Property Get PurDiscounts() As Boolean
  Let PurDiscounts = VBA.CBool(GLOVEC_VERSIONDATA(0, 11))
End Property

Public Property Let PurDiscounts(ByVal MBooPurDiscounts As Boolean)
  Call Me.MUpd(eqlMdlSheetSys, eqlMdlTblConfigapp_sys, Array(IIf(MBooPurDiscounts, 1, 0)), "pur_resumed_discounts", 1)
End Property

Public Property Get PurBonus() As Boolean
  Let PurBonus = VBA.CBool(GLOVEC_VERSIONDATA(0, 12))
End Property

Public Property Let PurBonus(ByVal MBooPurBonus As Boolean)
  Call Me.MUpd(eqlMdlSheetSys, eqlMdlTblConfigapp_sys, Array(IIf(MBooPurBonus, 1, 0)), "pur_resumed_bonus", 1)
End Property

Public Property Get PurExports() As Boolean
  Let PurExports = VBA.CBool(GLOVEC_VERSIONDATA(0, 13))
End Property

Public Property Let PurExports(ByVal MBooPurExports As Boolean)
  Call Me.MUpd(eqlMdlSheetSys, eqlMdlTblConfigapp_sys, Array(IIf(MBooPurExports, 1, 0)), "pur_resumed_exports", 1)
End Property

Public Property Get PurDestination() As String
  Let PurDestination = VBA.CStr(GLOVEC_VERSIONDATA(0, 14))
End Property

Public Property Let PurDestination(ByVal MStrPurDestination As String)
  Call Me.MUpd(eqlMdlSheetSys, eqlMdlTblConfigapp_sys, Array(MStrPurDestination), "pur_destination", 1)
End Property

' DB tables - Company options
Public Property Get CompanyLogo() As String
  Let CompanyLogo = VBA.CStr(GLOVEC_VERSIONDATA(0, 15))
End Property

Public Property Let CompanyLogo(ByVal MStrCompanyLogo As String)
  Call Me.MUpd(eqlMdlSheetSys, eqlMdlTblConfigapp_sys, Array(MStrCompanyLogo), "company_logo", 1)
End Property

Public Property Get CompanyName() As String
  Let CompanyName = VBA.CStr(GLOVEC_VERSIONDATA(0, 16))
End Property

Public Property Let CompanyName(ByVal MStrCompanyName As String)
  Call Me.MUpd(eqlMdlSheetSys, eqlMdlTblConfigapp_sys, Array(MStrCompanyName), "company_name", 1)
End Property

Public Property Get CompanyAddress() As String
  Let CompanyAddress = VBA.CStr(GLOVEC_VERSIONDATA(0, 17))
End Property

Public Property Let CompanyAddress(ByVal MStrCompanyAddress As String)
  Call Me.MUpd(eqlMdlSheetSys, eqlMdlTblConfigapp_sys, Array(MStrCompanyAddress), "company_address", 1)
End Property

Public Property Get CompanyPhone() As String
  Let CompanyPhone = VBA.CStr(GLOVEC_VERSIONDATA(0, 18))
End Property

Public Property Let CompanyPhone(ByVal MStrCompanyPhone As String)
  Call Me.MUpd(eqlMdlSheetSys, eqlMdlTblConfigapp_sys, Array(MStrCompanyPhone), "company_phone", 1)
End Property

Public Property Get CompanyEmail() As String
  Let CompanyEmail = VBA.CStr(GLOVEC_VERSIONDATA(0, 19))
End Property

Public Property Let CompanyEmail(ByVal MstrCompanyEmail As String)
  Call Me.MUpd(eqlMdlSheetSys, eqlMdlTblConfigapp_sys, Array(MstrCompanyEmail), "company_email", 1)
End Property

Public Property Get CompanyHeader() As Boolean
  Let CompanyHeader = VBA.CBool(GLOVEC_VERSIONDATA(0, 20))
End Property

Public Property Let CompanyHeader(ByVal MBooCompanyHeader As Boolean)
  Call Me.MUpd(eqlMdlSheetSys, eqlMdlTblConfigapp_sys, Array(MBooCompanyHeader), "company_header", 1)
End Property

'Currencies
Public Property Get CurrencyMainSymbolLocal() As String
  Let CurrencyMainSymbolLocal = VBA.CStr(Me.MGet(eqlMdlSheetSys, eqlMdlTblConfigcurrencies_sys, eqlMdlArray, "symbol", MStrWhere:="currency_main=1")(0, 0))
End Property

Public Property Get CurrencyMainSymbolForeign() As String
  Let CurrencyMainSymbolForeign = VBA.CStr(Me.MGet(eqlMdlSheetSys, eqlMdlTblConfigcurrencies_sys, eqlMdlArray, "symbol", MStrWhere:="currency_foreign=1")(0, 0))
End Property



' CONSTRUCTOR/DESTRUCTOR ----------------------------------------------------------- '
Private Sub Class_Terminate()

  Let GLOVEC_VERSIONDATA = Empty
  Set GLOSHT_DATA = Nothing
  Set RES = Nothing

End Sub



' CORE METHODS LIST (PRIVATE) ------------------------------------------------------ '
Private Sub Model_start(ByRef MMPT As ClassEqlModelParent, ByVal MSheet As EQLMDL_ENU_SHEETS, ByRef MTable As EQLMDL_ENU_TABLES)

  Dim Wst As Worksheet

  On Error GoTo EH
  ' Check if Table exists
  If MSheet <= 0 Or MSheet > Application.ThisWorkbook.Worksheets.Count Then Call VBA.Err.Raise(CUSTOM_ERROR_MDL, "MDL::-start", "¡No existe el orígen de datos de índie: '" & MSheet & "'!")
  Set Wst = Application.ThisWorkbook.Worksheets(MSheet)
  With Wst
    If .ListObjects.Count <= 0 Then Call VBA.Err.Raise(CUSTOM_ERROR_MDL, "MDL::-start", "¡No existen tablas en el orígen de datos: '" & .Name & "'!")
    Let MTable = MTable - IIf(MTable < 100, 0, IIf(MTable < 200, 100, 200))
    If .ListObjects(MTable) Is Nothing Then Call VBA.Err.Raise(CUSTOM_ERROR_MDL, "MDL::-start", "¡No existe la tabla de índice '" & MTable & "' en el orígen de datos: '" & .Name & "'!")
  End With

  ' Set local sheet
  If GLOSHT_DATA Is Nothing Then
    Set GLOSHT_DATA = Application.ThisWorkbook.Worksheets(MSheet)
  Else
    If Not GLOSHT_DATA Is Application.ThisWorkbook.Worksheets(MSheet) Then
      Set GLOSHT_DATA = Nothing
      Set GLOSHT_DATA = Application.ThisWorkbook.Worksheets(MSheet)
    End If
  End If
  If Not GLOSHT_DATA Is Nothing Then Debug.Print GLOSHT_DATA.Name

  ' Set values to Abstract parent model
  If MMPT Is Nothing Then
    Set MMPT = New ClassEqlModelParent
    Set MMPT.ResourcesApp = RES
  End If

EH:
  Set Wst = Nothing
  Call RES.ErrorHandler.ErrorHandlerRaise("MDL::-set_datatables")

End Sub



' ERRORS MANNAGEMENT AND ASSERTSIONS ----------------------------------------------- '
Private Sub Debug_assert(ByVal Cond As Boolean, Optional Origin As String = VBA.vbNullString)

  #If Debbuging Then
    Debug.Assert Cond
    If Not Cond Then Debug.Print "Assert on: " & Origin
  #End If

End Sub
' CORE METHODS LIST (PRIVATE) ------------------------------------------------------ '



' CORE METHODS LIST (PUBLIC) ------------------------------------------------------- '
' PUBLIC METHODS CRUD
Public Sub MSet( _
  ByVal MEnuSheet As EQLMDL_ENU_SHEETS, _
  ByVal MEnuTable As EQLMDL_ENU_TABLES, _
  Optional MVarData As Variant = Empty, _
  Optional MLngRows As Long = 0, _
  Optional MBooEmpties As Boolean = False)

  Dim MPT As ClassEqlModelParent

  On Error GoTo EH
  Call Model_start(MPT, MEnuSheet, MEnuTable)
  Call MPT.MSet(GLOSHT_DATA.ListObjects(MEnuTable), MVarData, MLngRows, MBooEmpties)

EH:
  Set MPT = Nothing
  Call RES.ErrorHandler.ErrorHandlerRaise("MDL::MSet")

End Sub

Public Function MGet( _
  ByVal MEnuSheet As EQLMDL_ENU_SHEETS, _
  ByVal MEnuTable As EQLMDL_ENU_TABLES, _
  ByVal MEnuResAs As EQLMDL_ENU_TABLE_AS, _
  Optional MStrField As String = VBA.vbNullString, _
  Optional MLngIndex As Long = 0, _
  Optional MStrWhere As String = VBA.vbNullString) As Variant

  Dim MPT As ClassEqlModelParent

  On Error GoTo EH
  Call Model_start(MPT, MEnuSheet, MEnuTable)
  If MEnuResAs = eqlMdlArray Or MEnuResAs = eqlMdlStrTableName Then
    Let MGet = MPT.MGet(GLOSHT_DATA.ListObjects(MEnuTable), MEnuResAs, MStrField, MLngIndex, MStrWhere)
  Else
    Set MGet = MPT.MGet(GLOSHT_DATA.ListObjects(MEnuTable), MEnuResAs, MStrField, MLngIndex, MStrWhere)
  End If

EH:
  Set MPT = Nothing
  Call RES.ErrorHandler.ErrorHandlerRaise("MDL::MGet")

End Function

Public Sub MUpd( _
  ByVal MEnuSheet As EQLMDL_ENU_SHEETS, _
  ByVal MEnuTable As EQLMDL_ENU_TABLES, _
  ByVal MVarData As Variant, _
  Optional MStrField As String = VBA.vbNullString, _
  Optional MLngIndex As Long = 0, _
  Optional MStrWhere As String = VBA.vbNullString)

  Dim MPT As ClassEqlModelParent

  On Error GoTo EH
  Call Model_start(MPT, MEnuSheet, MEnuTable)
  Call MPT.MUpd(GLOSHT_DATA.ListObjects(MEnuTable), MVarData, MStrField, MLngIndex, MStrWhere)

EH:
  Set MPT = Nothing
  Call RES.ErrorHandler.ErrorHandlerRaise("MDL::MUpd")

End Sub

Public Sub MDel(ByVal MEnuSheet As EQLMDL_ENU_SHEETS, ByVal MEnuTable As EQLMDL_ENU_TABLES, Optional MLngRow As Long = 0)

  Dim MPT As ClassEqlModelParent

  On Error GoTo EH
  Call Model_start(MPT, MEnuSheet, MEnuTable)
  Call MPT.MDel(GLOSHT_DATA.ListObjects(MEnuTable), MLngRow)

EH:
  Set MPT = Nothing
  Call RES.ErrorHandler.ErrorHandlerRaise("MDL::MDel")

End Sub
' CORE METHODS LIST (PUBLIC) ------------------------------------------------------- '
' ======================= CLASS EXCLIQ LITE MODEL STRUCTURE ======================== '





