VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ModelExcliqliteDatasheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'============================= CLASS MODULES STRUCTURE =============================='
' OBJECT VARIABLES (GLOOBJ_) --------------------------------------------------------'
Private EHGLOBAL As AppErrorHandler
Private MPT As ModelParent_current
Private GLOSHT_DATA As Worksheet

' VARIABLES (GLOStr_) ---------------------------------------------------------------'
Private GLOVEC_VERSIONDATA As Variant

' VARIABLES CUSTOM (GLOEnu|Typ_) ----------------------------------------------------'
Public Enum EQLMDL_ENU_TABLESLITE
  eqlMdlTblFormulas
  eqlMdlTblCurrencies
  eqlMdlTblTables
  eqlMdlTblCompany
  eqlMdlTblOrepurchaseoptions
  eqlMdlTblConfigpurdestiny
  eqlMdlTblConfigpurrm
  eqlMdlTblConfigpurretentions
  eqlMdlTblConfigpurretentionsother
  eqlMdlTblPurchasevariousretentions
  eqlMdlTblPurchasevariousretentionsother
  eqlMdlTblPurchasebonus
  eqlMdlTblPurchasediscounts
  eqlMdlTblPurchasebasedata
  eqlMdlTblPurchasemaindata
  eqlMdlTblConpurcontents
  eqlMdlTblConpurrefination
  eqlMdlTblConpurpenalties
  eqlMdlTblConpurdedexp
  eqlMdlTblConpurdeddiscounts
  eqlMdlTblConpurbonus
  eqlMdlTblConpurbase
  eqlMdlTblChemicalelements
  eqlMdlTblChemicalelementscx
  eqlMdlTblChemicalunits
  eqlMdlTblChemicalunitscx
  eqlMdlTblVersiontable
  eqlMdlTblDatatest
End Enum

' CONSTANTS ERROR MESSAGES (GLOSTR_ERR_) --------------------------------------------'
Private Const CUSTOM_ERROR = VBA.vbObjectError + 519







' PROPERTIES ------------------------------------------------------------------------'
Public Property Set ErrorHandler(ByRef ObjEH As AppErrorHandler)

  If EHGLOBAL Is Nothing Then
    Set EHGLOBAL = ObjEH
    Set MPT.ErrorHandler = EHGLOBAL
    Call MPT.MCheckForDatasheetbasetables(Model_set_datatables())
    Let GLOVEC_VERSIONDATA = Me.MGet(eqlMdlTblVersiontable, eqlMdlArray, MLngIndex:=1)
  End If

End Property

Public Property Get LengthRows(ByVal MEnuTable As EQLMDL_ENU_TABLESLITE) As Long

  Const sFN As String = "MDL::LengthRows"
  
  On Error GoTo EH
  If EHGLOBAL Is Nothing Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "Se requiere un manejador de errores.")
  Let LengthRows = MPT.LengthRows(Model_get_for_strtable(MEnuTable))

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Property

Public Property Get LengthCols(ByVal MEnuTable As EQLMDL_ENU_TABLESLITE) As Long

  Const sFN As String = "MDL::LengthCols"
  
  On Error GoTo EH
  If EHGLOBAL Is Nothing Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "Se requiere un manejador de errores.")
  Let LengthCols = MPT.LengthCols(Model_get_for_strtable(MEnuTable))

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Property


' App data
Public Property Get AppTitle() As String
  Let AppTitle = VBA.CStr(GLOVEC_VERSIONDATA(0, 0))
End Property

Public Property Get AppVersion() As String
  Let AppVersion = "v" & VBA.CStr(GLOVEC_VERSIONDATA(0, 1)) & "." & VBA.CStr(GLOVEC_VERSIONDATA(0, 2)) & "." & VBA.CStr(GLOVEC_VERSIONDATA(0, 3))
End Property

Public Property Get AppAuthor() As String
  Let AppAuthor = VBA.CStr(GLOVEC_VERSIONDATA(0, 4))
End Property

Public Property Get AppEmail() As String
  Let AppEmail = VBA.CStr(GLOVEC_VERSIONDATA(0, 5))
End Property

Public Property Get AppSlogan() As String
  Let AppSlogan = VBA.CStr(GLOVEC_VERSIONDATA(0, 6))
End Property

Public Property Get AppYear() As String
  Let AppYear = VBA.CStr(GLOVEC_VERSIONDATA(0, 7))
End Property

Public Property Get AppOrigin() As String
  Let AppOrigin = VBA.CStr(GLOVEC_VERSIONDATA(0, 8))
End Property

Public Property Get AppGit() As String
  Let AppGit = VBA.CStr(GLOVEC_VERSIONDATA(0, 9))
End Property

Public Property Get AppLicense() As String
  Let AppLicense = VBA.CStr(GLOVEC_VERSIONDATA(0, 10))
End Property






' DB tables
Public Property Get Formulas() As Boolean
  Let Formulas = VBA.CBool(Me.MGet(eqlMdlTblFormulas, eqlMdlArray, "formulas", 1)(0, 0))
End Property

Public Property Let Formulas(ByVal xStatus As Boolean)
  Call Me.MUpd(eqlMdlTblFormulas, Array(IIf(xStatus, 1, 0)), "formulas", 1)
End Property

Public Property Get Tables() As Byte
  Let Tables = VBA.CByte(Me.MGet(eqlMdlTblTables, eqlMdlArray, "table_type", 1)(0, 0))
End Property

Public Property Let Tables(ByVal xTable As Byte) '0: vile table, 1: plain table, 2: table
  Call Me.MUpd(eqlMdlTblTables, Array(xTable), "table_type", 1)
End Property

Public Property Get TableStyles() As Byte
  Let TableStyles = VBA.CByte(Me.MGet(eqlMdlTblTables, eqlMdlArray, "table_style", 1)(0, 0))
End Property

Public Property Let TableStyles(ByVal xStyle As Byte) '0: Default grey eql style, 1: green eql style, 2: excel predefined pastel style
  Call Me.MUpd(eqlMdlTblTables, Array(xStyle), "table_style", 1)
End Property

Public Property Get TableThemeColor() As XlThemeColor ' Tables Themes: Returns the current EQSystem theme color using as base Table Style
  Select Case Me.TableStyles
    Case 0: Let TableThemeColor = xlThemeColorDark2 'Grey ExcliqLite
    Case 1: Let TableThemeColor = xlThemeColorAccent6 '10 'Green ExcliqLite
    Case Else: Let TableThemeColor = xlThemeColorAccent1 'Blue XL
  End Select
End Property

Public Property Get TableThemeFontColor() As XlThemeColor ' Tables Themes: Returns the current EQSystem theme color using as base Table Style
  Select Case Me.TableStyles
    Case 0: Let TableThemeFontColor = xlThemeColorDark2 'Black ExcliqLite
    Case 1: Let TableThemeFontColor = xlThemeColorDark1 'White ExcliqLite
    Case Else: Let TableThemeFontColor = xlThemeColorDark1 'White XL
  End Select
End Property


Public Property Get TableRowStripes() As Boolean
  Let TableRowStripes = VBA.CBool(Me.MGet(eqlMdlTblTables, eqlMdlArray, "table_rowstripe", 1)(0, 0))
End Property

Public Property Let TableRowStripes(ByVal xStripe As Boolean) 'TRUE: Stiped rows, FALSE: Simple rows
  Call Me.MUpd(eqlMdlTblTables, Array(IIf(xStripe, 1, 0)), "table_rowstripe", 1)
End Property

Public Property Get TableSelectionsCell() As Boolean
  Let TableSelectionsCell = VBA.CBool(Me.MGet(eqlMdlTblTables, eqlMdlArray, "table_selections_cell", 1)(0, 0))
End Property

Public Property Let TableSelectionsCell(ByVal MBooSelections As Boolean) 'TRUE: Show Inputbox 8 enabled, FALSE: Show Inputbox 8 disabled
  Call Me.MUpd(eqlMdlTblTables, Array(IIf(MBooSelections, 1, 0)), "table_selections_cell", 1)
End Property

'Currencies
'Public Property Get CurrencyMainSymbol() As String
'  Let CurrencyMainSymbol = VBA.CStr(MLT.MGet(eqlMdlTblCurrencies, eqlMdlArray, "symbol", MStrWhere:="currency_main=1")(0, 0))
'End Property
'
'Public Property Let CurrencyMainSymbol(ByVal MBooCurrency As Boolean)
'  Call Me.MUpd(eqlMdlTblTables, Array(IIf(MBooCurrency, 1, 0)), "symbol", 1)
'End Property








' CONSTRUCTOR/DESTRUCTOR ------------------------------------------------------------'
Private Sub Class_Initialize()

  Set GLOSHT_DATA = excliqlitedatasheet
  Set MPT = New ModelParent_current
  Set MPT.SheetData = GLOSHT_DATA

End Sub

Private Sub Class_Terminate()

  Let GLOVEC_VERSIONDATA = Empty
  Set MPT = Nothing
  Set GLOSHT_DATA = Nothing
  Set EHGLOBAL = Nothing

End Sub







' CORE METHODS PRIVATE --------------------------------------------------------------'
Private Function Model_set_datatables() As Variant

  Dim VecTblNames As Variant

  On Error GoTo EH
  ReDim VecTblNames(27)
  Let VecTblNames(0) = "excliqliteformulascon"
  Let VecTblNames(1) = "excliqlitecurrenciescon"
  Let VecTblNames(2) = "excliqlitetablescon"
  Let VecTblNames(3) = "excliqlitetablescompany"
  Let VecTblNames(4) = "excliqliteorepurchaseoptions"
  Let VecTblNames(5) = "excliqliteconfigpurdestiny"
  Let VecTblNames(6) = "excliqliteconfigpurrm"
  Let VecTblNames(7) = "excliqliteconfigpurretentions"
  Let VecTblNames(8) = "excliqliteconfigpurretentionsother"
  Let VecTblNames(9) = "excliqlitepurchasevariousretentions"
  Let VecTblNames(10) = "excliqlitepurchasevariousretentionsother"
  Let VecTblNames(11) = "excliqlitepurchasebonus"
  Let VecTblNames(12) = "excliqlitepurchasediscounts"
  Let VecTblNames(13) = "excliqlitepurchasebasedata"
  Let VecTblNames(14) = "excliqlitepurchasemaindata"
  Let VecTblNames(15) = "excliqliteconpurcontents"
  Let VecTblNames(16) = "excliqliteconpurrefination"
  Let VecTblNames(17) = "excliqliteconpurpenalties"
  Let VecTblNames(18) = "excliqliteconpurdedexp"
  Let VecTblNames(19) = "excliqliteconpurdeddiscounts"
  Let VecTblNames(20) = "excliqliteconpurbonus"
  Let VecTblNames(21) = "excliqliteconpurbase"
  Let VecTblNames(22) = "excliqlitechemicalelements"
  Let VecTblNames(23) = "excliqlitechemicalelementscx"
  Let VecTblNames(24) = "excliqlitechemicalunits"
  Let VecTblNames(25) = "excliqlitechemicalunitscx"
  Let VecTblNames(26) = "excliqliteversiontable"
  Let VecTblNames(27) = "excliqlitedatatest"

  'Return
  Let Model_set_datatables = VecTblNames

EH:
  Let VecTblNames = Empty
  Call EHGLOBAL.ErrorHandlerRaise("MDL::-set_datatables")

End Function

Private Function Model_get_for_strtable(ByVal MEnuTable As EQLMDL_ENU_TABLESLITE) As String

  Dim VecTables As Variant
  Dim StrTable As String

  On Error GoTo EH
  Let VecTables = Model_set_datatables()
  Let StrTable = VecTables(MEnuTable)
  Erase VecTables

  'Return
  Let Model_get_for_strtable = StrTable

EH:
  Let VecTables = Empty
  Call EHGLOBAL.ErrorHandlerRaise("MDL::-get_for_strtable")

End Function







' ASSETS ----------------------------------------------------------------------------'








' ERRORS MANNAGEMENT AND ASSERTSIONS ------------------------------------------------'
Private Sub Debug_assert(ByVal Cond As Boolean, Optional Origin As String = VBA.vbNullString)
  #If Debbuging Then
    Debug.Assert Cond
    If Not Cond Then Debug.Print "Assert on: " & Origin
  #End If
End Sub
' CORE METHODS PRIVATE --------------------------------------------------------------'







' METHODS PUBLIC --------------------------------------------------------------------'

Public Sub SetCurrency(ByVal MMulti As String, ByVal MSingle As String, ByVal MSymbol As String)
  Let excliqlitedatasheet.ListObjects("excliqlitecurrenciescon").ListColumns("currencies").DataBodyRange.Cells(1, 1).Value = MMulti
  Let excliqlitedatasheet.ListObjects("excliqlitecurrenciescon").ListColumns("currency").DataBodyRange.Cells(1, 1).Value = MSingle
  Let excliqlitedatasheet.ListObjects("excliqlitecurrenciescon").ListColumns("symbol").DataBodyRange.Cells(1, 1).Value = MSymbol
End Sub

Public Sub GetCurrency(ByRef MMulti As String, ByRef MSingle As String, ByRef MSymbol As String)
  Let MMulti = excliqlitedatasheet.ListObjects("excliqlitecurrenciescon").ListColumns("currencies").DataBodyRange.Cells(1, 1).Value
  Let MSingle = excliqlitedatasheet.ListObjects("excliqlitecurrenciescon").ListColumns("currency").DataBodyRange.Cells(1, 1).Value
  Let MSymbol = excliqlitedatasheet.ListObjects("excliqlitecurrenciescon").ListColumns("symbol").DataBodyRange.Cells(1, 1).Value
End Sub




Public Function GetTableSelectionsModsCompact() As Boolean
  Let GetTableSelectionsModsCompact = excliqlitedatasheet.ListObjects("excliqlitetablescon").ListColumns("table_selections_compact").DataBodyRange.Cells(1, 1).Value
End Function

Public Sub SetTableSelectionsModsCompact(ByVal RBooSelectionsCompact As Boolean)
  'TRUE: To print a compact or extended Balmet
  Let excliqlitedatasheet.ListObjects("excliqlitetablescon").ListColumns("table_selections_compact").DataBodyRange.Cells(1, 1).Value = IIf(RBooSelectionsCompact, 1, 0)
End Sub

Public Function GetTableSelectionsModsHeads() As Boolean
  Let GetTableSelectionsModsHeads = excliqlitedatasheet.ListObjects("excliqlitetablescon").ListColumns("table_selections_heads").DataBodyRange.Cells(1, 1).Value
End Function

Public Sub SetTableSelectionsModsHeads(ByVal RBooSelectionsHeads As Boolean)
  'TRUE: To print a Feed heads on Balmet
  Let excliqlitedatasheet.ListObjects("excliqlitetablescon").ListColumns("table_selections_heads").DataBodyRange.Cells(1, 1).Value = IIf(RBooSelectionsHeads, 1, 0)
End Sub

Public Function GetTablePurOprionsRMDetail() As Boolean
  'TRUE: To resumed RM in excliqliteorepurchaseoptions
  Let GetTablePurOprionsRMDetail = VBA.CBool(excliqlitedatasheet.ListObjects("excliqliteorepurchaseoptions").ListColumns("orepurchase_rmdetail").DataBodyRange.Cells(1, 1).Value)
End Function

Public Sub SetTablePurOptionsRMDetail(ByVal RBooDetail As Boolean)
  'TRUE: To resumed RM in excliqliteorepurchaseoptions
  Let excliqlitedatasheet.ListObjects("excliqliteorepurchaseoptions").ListColumns("orepurchase_rmdetail").DataBodyRange.Cells(1, 1).Value = IIf(RBooDetail, 1, 0)
End Sub

Public Function GetTablePageOptionsExists() As Boolean
  'TRUE: To resumed RM in excliqlitetablescompany
  Let GetTablePageOptionsExists = VBA.CBool(excliqlitedatasheet.ListObjects("excliqlitetablescompany").ListColumns("company_header").DataBodyRange.Cells(1, 1).Value)
End Function

Public Sub SetTablePageOptionsExists(ByVal RBooExists As Boolean)
  'TRUE: Exists data header/footer in excliqlitetablescompany
  Let excliqlitedatasheet.ListObjects("excliqlitetablescompany").ListColumns("company_header").DataBodyRange.Cells(1, 1).Value = IIf(RBooExists, 1, 0)
End Sub

Public Function GetTablePageOptionsLogo() As String
  'Logo in excliqlitetablescompany
  Let GetTablePageOptionsLogo = excliqlitedatasheet.ListObjects("excliqlitetablescompany").ListColumns("company_logo").DataBodyRange.Cells(1, 1).Value
End Function

Public Sub SetTablePageOptionsLogo(ByVal RStrLogo As String)
  'Logo in excliqlitetablescompany
  Let excliqlitedatasheet.ListObjects("excliqlitetablescompany").ListColumns("company_logo").DataBodyRange.Cells(1, 1).Value = RStrLogo
End Sub

Public Function GetTablePageOptionsCompany() As String
  'Company name in excliqlitetablescompany
  Let GetTablePageOptionsCompany = excliqlitedatasheet.ListObjects("excliqlitetablescompany").ListColumns("company_name").DataBodyRange.Cells(1, 1).Value
End Function

Public Sub SetTablePageOptionsCompany(ByVal RStrName As String)
  'Company name excliqlitetablescompany
  Let excliqlitedatasheet.ListObjects("excliqlitetablescompany").ListColumns("company_name").DataBodyRange.Cells(1, 1).Value = RStrName
End Sub

Public Function GetTablePageOptionsAddress() As String
  'Company address in excliqlitetablescompany
  Let GetTablePageOptionsAddress = excliqlitedatasheet.ListObjects("excliqlitetablescompany").ListColumns("company_address").DataBodyRange.Cells(1, 1).Value
End Function

Public Sub SetTablePageOptionsAddress(ByVal RStrAddress As String)
  'Company address excliqlitetablescompany
  Let excliqlitedatasheet.ListObjects("excliqlitetablescompany").ListColumns("company_address").DataBodyRange.Cells(1, 1).Value = RStrAddress
End Sub

Public Function GetTablePageOptionsPhone() As String
  'Company phone in excliqlitetablescompany
  Let GetTablePageOptionsPhone = excliqlitedatasheet.ListObjects("excliqlitetablescompany").ListColumns("company_phone").DataBodyRange.Cells(1, 1).Value
End Function

Public Sub SetTablePageOptionsPhone(ByVal RStrPhone As String)
  'Company phone excliqlitetablescompany
  Let excliqlitedatasheet.ListObjects("excliqlitetablescompany").ListColumns("company_phone").DataBodyRange.Cells(1, 1).Value = RStrPhone
End Sub

Public Function GetTablePageOptionsEmail() As String
  'Company email in excliqlitetablescompany
  Let GetTablePageOptionsEmail = excliqlitedatasheet.ListObjects("excliqlitetablescompany").ListColumns("company_email").DataBodyRange.Cells(1, 1).Value
End Function

Public Sub SetTablePageOptionsEmail(ByVal RStrEmail As String)
  'Company email excliqlitetablescompany
  Let excliqlitedatasheet.ListObjects("excliqlitetablescompany").ListColumns("company_email").DataBodyRange.Cells(1, 1).Value = RStrEmail
End Sub

'Exists each Page data?
Public Function GetTablePageOptionsLogoExists() As Byte
  'Logo in excliqlitetablescompany
  Let GetTablePageOptionsLogoExists = VBA.CByte(excliqlitedatasheet.ListObjects("excliqlitetablescompany").ListColumns("company_logo").DataBodyRange.Cells(2, 1).Value)
End Function

Public Sub SetTablePageOptionsLogoExists(ByVal RBytLogo As Byte)
  'Logo in excliqlitetablescompany
  Let excliqlitedatasheet.ListObjects("excliqlitetablescompany").ListColumns("company_logo").DataBodyRange.Cells(2, 1).Value = RBytLogo
End Sub

Public Function GetTablePageOptionsCompanyExists() As Byte
  'Company name in excliqlitetablescompany
  Let GetTablePageOptionsCompanyExists = VBA.CByte(excliqlitedatasheet.ListObjects("excliqlitetablescompany").ListColumns("company_name").DataBodyRange.Cells(2, 1).Value)
End Function

Public Sub SetTablePageOptionsCompanyExists(ByVal RBytName As Byte)
  'Company name excliqlitetablescompany
  Let excliqlitedatasheet.ListObjects("excliqlitetablescompany").ListColumns("company_name").DataBodyRange.Cells(2, 1).Value = RBytName
End Sub

Public Function GetTablePageOptionsAddressExists() As Byte
  'Company address in excliqlitetablescompany
  Let GetTablePageOptionsAddressExists = VBA.CByte(excliqlitedatasheet.ListObjects("excliqlitetablescompany").ListColumns("company_address").DataBodyRange.Cells(2, 1).Value)
End Function

Public Sub SetTablePageOptionsAddressExists(ByVal RBytAddress As Byte)
  'Company address excliqlitetablescompany
  Let excliqlitedatasheet.ListObjects("excliqlitetablescompany").ListColumns("company_address").DataBodyRange.Cells(2, 1).Value = RBytAddress
End Sub

Public Function GetTablePageOptionsPhoneExists() As Byte
  'Company phone in excliqlitetablescompany
  Let GetTablePageOptionsPhoneExists = VBA.CByte(excliqlitedatasheet.ListObjects("excliqlitetablescompany").ListColumns("company_phone").DataBodyRange.Cells(2, 1).Value)
End Function

Public Sub SetTablePageOptionsPhoneExists(ByVal RBytPhone As Byte)
  'Company phone excliqlitetablescompany
  Let excliqlitedatasheet.ListObjects("excliqlitetablescompany").ListColumns("company_phone").DataBodyRange.Cells(2, 1).Value = RBytPhone
End Sub

Public Function GetTablePageOptionsEmailExists() As Byte
  'Company email in excliqlitetablescompany
  Let GetTablePageOptionsEmailExists = VBA.CByte(excliqlitedatasheet.ListObjects("excliqlitetablescompany").ListColumns("company_email").DataBodyRange.Cells(2, 1).Value)
End Function

Public Sub SetTablePageOptionsEmailExists(ByVal RBytEmail As Byte)
  'Company email excliqlitetablescompany
  Let excliqlitedatasheet.ListObjects("excliqlitetablescompany").ListColumns("company_email").DataBodyRange.Cells(2, 1).Value = RBytEmail
End Sub


''Tables Themes
'Public Function GetTablesThemeColor() As XlThemeColor
'
'  'Returns the current EQSystem theme color
'  Select Case Me.TableStyles
'    Case 0: Let GetTablesThemeColor = xlThemeColorDark2 'Grey ExcliqLite
'    Case 1: Let GetTablesThemeColor = xlThemeColorAccent6 '10 'Green ExcliqLite
'    Case Else: Let GetTablesThemeColor = xlThemeColorAccent1 'Blue XL
'  End Select
'
'End Function
'
'Public Function GetTablesThemeFontColor() As XlThemeColor
'
'  'Returns the current EQSystem theme color
'  Select Case Me.TableStyles
'    Case 0: Let GetTablesThemeFontColor = xlThemeColorDark2 'Black ExcliqLite
'    Case 1: Let GetTablesThemeFontColor = xlThemeColorDark1 'White ExcliqLite
'    Case Else: Let GetTablesThemeFontColor = xlThemeColorDark1 'White XL
'  End Select
'
'End Function


'PUBLIC METHODS CRUD
Public Sub MSet(ByVal MEnuTable As EQLMDL_ENU_TABLESLITE, Optional MVarData As Variant = Empty, Optional MLngRows As Long = 0)

  Const sFN As String = "MDL::MSet"
  
  On Error GoTo EH
  If EHGLOBAL Is Nothing Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "Se requiere un manejador de errores.")
  Call MPT.MSet(Model_get_for_strtable(MEnuTable), MVarData, MLngRows)

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Sub

Public Function MGet( _
  ByVal MEnuTable As EQLMDL_ENU_TABLESLITE, _
  ByVal MEnuResAs As EQLMDL_ENU_TABLE_AS, _
  Optional MStrField As String = VBA.vbNullString, _
  Optional MLngIndex As Long = 0, _
  Optional MStrWhere As String = VBA.vbNullString) As Variant

  Const sFN As String = "MDL::MGet"
  
  On Error GoTo EH
  If EHGLOBAL Is Nothing Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "Se requiere un manejador de errores.")
  Let MGet = MPT.MGet(Model_get_for_strtable(MEnuTable), MEnuResAs, MStrField, MLngIndex, MStrWhere)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("MDL::MGet")

End Function

Public Sub MUpd( _
  ByVal MEnuTable As EQLMDL_ENU_TABLESLITE, _
  ByVal MVarData As Variant, _
  Optional MStrField As String = VBA.vbNullString, _
  Optional MLngIndex As Long = 0, _
  Optional MStrWhere As String = VBA.vbNullString)

  Const sFN As String = "MDL::MUpd"
  
  On Error GoTo EH
  If EHGLOBAL Is Nothing Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "Se requiere un manejador de errores.")
  Call MPT.MUpd(Model_get_for_strtable(MEnuTable), MVarData, MStrField, MLngIndex, MStrWhere)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("MDL::MUpd")

End Sub

Public Sub MDel(ByVal MEnuTable As EQLMDL_ENU_TABLESLITE, Optional MLngRow As Long = 0)

  Const sFN As String = "MDL::MDel"
  
  On Error GoTo EH
  If EHGLOBAL Is Nothing Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "Se requiere un manejador de errores.")
  Call MPT.MDel(Model_get_for_strtable(MEnuTable), MLngRow)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("MDL::MDel")

End Sub
' METHODS PUBLIC --------------------------------------------------------------------'
'============================= CLASS MODULES STRUCTURE =============================='



