VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppExcliqPurchaseMin_current"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ABSTRACT IMPLEMENTATIONS ----------------------------------------------------------'
Implements ICoreController

' OBJECT VARIABLES (GLOOBJ_) --------------------------------------------------------'
Private EHGLOBAL As AppErrorHandler
Private GLOSHTDATA As Worksheet

Private MLT As ModelExcliqliteDatasheet
Private MDL As Object 'FOR EXLIQ

' VARIABLES (GLOStr_) ---------------------------------------------------------------'
Private GLOBOO_EH As Boolean
Private GLOBOO_FORMULAS As Boolean
Private GLOBOO_ACCURATEANYWAY As Boolean
Private GLOBOO_RESUMED_RM As Boolean
Private GLOBOO_RESUMED_RETENTIONS As Boolean
Private GLOBOO_RESUMED_RETENTIONSOTHERS As Boolean
Private GLOBOO_RESUMED_DISCOUNTS As Boolean
Private GLOBOO_RESUMED_BONUS As Boolean
Private GLOLNG_ROWCOUNTER As Long

Private GLOSTR_CURRENCY As String
Private GLOSTR_CURRENCIES As String
Private GLOSTR_CURLOCAL As String
Private GLOSTR_CURFOREIGN  As String

' VARIABLES CUSTOM (GLOEnu|Typ_) ----------------------------------------------------'
Private Type GLOTYPPURCHASE
  PVecTitle As Variant
  PVecInvoice As Variant
  PVecProvider As Variant
  PVecOre As Variant
  PVecOrePay As Variant
  PVecGrossValue As Variant
  PVecDeductions As Variant
  PVecBonus As Variant
  PVecFinal As Variant
  PVecUnits As Variant
  PBooBonus As Boolean
  PStrTotalOreRng As String
  PDblTotalOreVal As Double
  PStrTotalPayRng As String
  PDblTotalPayVal As Double
  PStrTotalGrossRng As String
  PDblTotalGrossVal As Double
  PStrTotalDedRng As String
  PDblTotalDedVal As Double
  PStrTotalBonRng As String
  PDblTotalBonVal As Double
  PLngRowCounter As Long
End Type

Private Type GLOSTRLO
  PCurrencies As String
  PBasedata As String
  PMaindata As String
  POptions As String
  PVariousret As String
  PVariousretother As String
  PDiscounts As String
  PBonus As String
  PConfigprm As String
  PConfigpret As String
  PConfigpretother As String
End Type

'Private GLOVAR_PURCHASEMINERALRESULTS As Variant
Private GLOTYPLO As GLOSTRLO
Private GLOSTRPREFIX As String
'Private GLOBOOBONUS As Boolean

' CONSTANTS ERROR MESSAGES (GLOSTR_ERR_) --------------------------------------------'
Private Const GLOSTR_ERR_1 As String = "Faltan datos de orígen."
Private Const GLOSTR_ERR_2 As String = "Al menos un dato debe ser diferente de cero."
Private Const GLOSTR_ERR_3 As String = "Los datos independientes (X) deben tener la misma cantidad que los de orígen (Y)."

Private Const GLOSTR_ERR_9 As String = "¡Error, división entre cero!"





' ================== CLASS APPEXCLIQPURCHASEMINERAL STRUCTURE ======================'
' OBJECT VARIABLES (GLOOBJ_) --------------------------------------------------------'
' VARIABLES (GLOStr_) ---------------------------------------------------------------'
' CONSTANTS ERROR MESSAGES (GLOSTR_ERR_) --------------------------------------------'
' PROPERTIES ------------------------------------------------------------------------'
' CONSTRUCTOR/DESTRUCTOR ------------------------------------------------------------'
' CORE METHODS PRIVATE --------------------------------------------------------------'
' CORE METHODS LIST (PRIVATE) -------------------------------------------------------'
' PurchaseMineral_Core_Main
' PurchaseMineral_Core_Spreadmodule_generate_data_purchasebase
' PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal
' PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal_getoredata
' PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal_getorepay
' PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal_getgrossvalue
' PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal_getdeductions
' PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal_getbonus
' PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal_getfinalpurchase
' ASSETS ----------------------------------------------------------------------------'
' PurchaseMineral_Core_assets_get_average
' PurchaseMineral_Core_assets_get_grades
' PurchaseMineral_Core_assets_set_strlos
' PurchaseMineral_Core_assets_get_deductions_rm
' PurchaseMineral_Core_assets_get_deductions_retentions
' PurchaseMineral_Core_assets_get_deductions_retentionsothers
' PurchaseMineral_Core_assets_get_deductions_discounts
' ERRORS MANNAGEMENT AND ASSERTSIONS ------------------------------------------------'
' Debug_assert
' METHODS PUBLIC --------------------------------------------------------------------'
' INTERFACE METHODS LIST ------------------------------------------------------------'
' PurchaseMineralGet
' =================== CLASS APPEXCLIQPURCHASEMINERAL STRUCTURE ======================'

' PROPERTIES ------------------------------------------------------------------------'
' ABSTRACT INTERFACE PROPERTIES IMPLEMENTATION --------------------------------------'
Public Property Set ICoreController_ErrorHandler(RHS As AppErrorHandler)
  
  Set EHGLOBAL = Nothing
  Set EHGLOBAL = RHS
  Let GLOBOO_EH = (Not RHS Is Nothing)

End Property

' LOCAL PROPERTIES -------------------------------------------------------------------'
Public Property Set SheetData(ByVal PShtdata As Worksheet)
  
  Set GLOSHTDATA = PShtdata

End Property

Public Property Let SheetPrefix(ByVal PStrprefix As String)
  
  Let GLOSTRPREFIX = PStrprefix 'OJO

End Property

'Public Property Get Oreinvoice() As Variant
'
'  Let Oreinvoice = GLOVAR_PURCHASEMINERALRESULTS
'
'End Property

'Public Property Get Bonus() As Boolean
'
'  Let Bonus = GLOBOOBONUS
'
'End Property


' CONSTRUCTOR/DESTRUCTOR ------------------------------------------------------------'
Private Sub Class_Initialize()
  
  Set EHGLOBAL = New AppErrorHandler
  Set MLT = New ModelExcliqliteDatasheet
  Let GLOBOO_FORMULAS = MLT.Formulas
  Let GLOBOO_ACCURATEANYWAY = False 'Get this value from db
  Let GLOBOO_RESUMED_RM = False
  Let GLOBOO_RESUMED_RETENTIONS = False
  Let GLOBOO_RESUMED_RETENTIONSOTHERS = False
  Let GLOBOO_RESUMED_DISCOUNTS = False
  Let GLOBOO_RESUMED_BONUS = False
  'Let GLOLNG_ROWCOUNTER = 1
  Set MLT = Nothing

End Sub

Private Sub Class_Terminate()
  
  'Let GLOVAR_PURCHASEMINERALRESULTS = Empty
  Set GLOSHTDATA = Nothing
  If GLOBOO_EH = False Then Set EHGLOBAL = Nothing
  Set MLT = Nothing

End Sub





' CORE METHODS PRIVATE --------------------------------------------------------------'
' CORE METHODS LIST (PRIVATE) -------------------------------------------------------'
' Oreinvoice
Private Function PurchaseMineral_Core_Main(ByRef PVarRequest As Variant, ByRef PVarResponse As Variant) As Boolean
  
  Dim RES As AppResources_current
  Dim MIN As ClassPurchase_current

  Dim VecProviders As Variant, VecProvidersType As Variant, VecResponse As Variant, VecAux0 As Variant
  Dim BooVarious As Boolean
  Dim i As Long, k As Long

  Const sFN As String = "APPMIN::-core_main"

  On Error GoTo EH
  Let PurchaseMineral_Core_Main = False

  ' Request
  ' 0-StrPlace      5-StrProNames[]      10-BooNewSheet
  ' 1-StrDate       6-StrProType[]       11-RngBox
  ' 2-StrTC         7-BooAccurate        12-BooDataSheet: THIS PARAM COMES WITH THE MODELDATASHEET OBJECT AND MUST BE ASSINGNED AT MODEL GLOBAL OF THIS CONTROLLER
  ' 3-BooME         8-BooPage            13-Boo1PP: THIS BOOL SAYS THAT EACH PUR PER PAGE OR ALL PUR IN A PAGE
  ' 4-BooVarious    9-BooPrintTwo

  
  'Getting the number of ore invoices requested
  Set RES = New AppResources_current
  Let BooVarious = VBA.CBool(PVarRequest(4))
  If BooVarious Then
    Set RES.ErrorHandler = EHGLOBAL
    Let VecProviders = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(VBA.CStr(PVarRequest(5)), k, RValueType:=eqlResStrings)
    Let VecProvidersType = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(VBA.CStr(PVarRequest(6)), k, RValueType:=eqlResStrings)
    'Set RES = Nothing
  End If

  If VBA.CStr(PVarRequest(12)) = "lite" Then 'It will be a boolean
    Set MLT = New ModelExcliqliteDatasheet
    Set MLT.ErrorHandler = EHGLOBAL
    'Let GLOSTR_CURRENCY = VBA.CStr(MLT.MGet(eqlMdlTblCurrencies, eqlMdlArray, "currencies", MStrWhere:="currency_main=1")(0, 0))
    'Let GLOSTR_CURLOCAL = VBA.CStr(MLT.MGet(eqlMdlTblCurrencies, eqlMdlArray, "symbol", MStrWhere:="currency_main=1")(0, 0))
    Let VecAux0 = RES.ArrayToBaseZero(MLT.MGet(eqlMdlTblCurrencies, eqlMdlArray, "currency, currencies, symbol", MStrWhere:="currency_main=1"), eqlRes1D)
    Let GLOSTR_CURRENCY = VBA.CStr(VecAux0(0))
    Let GLOSTR_CURRENCIES = VBA.CStr(VecAux0(1))
    Let GLOSTR_CURLOCAL = VBA.CStr(VecAux0(2))
    Let GLOSTR_CURFOREIGN = VBA.CStr(MLT.MGet(eqlMdlTblCurrencies, eqlMdlArray, "symbol", MStrWhere:="currency_foreign=1")(0, 0))
  Else
    'Excliq
  End If

  'Requesting final invoices
  'Let GLOLNG_ROWCOUNTER = 1
  ReDim VecResponse(k - 1)
  For GLOLNG_ROWCOUNTER = 0 To (k - 1)
    
    If BooVarious Then
      Let PVarRequest(5) = VecProviders(GLOLNG_ROWCOUNTER) 'IF 1PROV AND 1PP VECPROVIDERS(FOO, FOO, FOO...)
      Let PVarRequest(6) = VecProvidersType(GLOLNG_ROWCOUNTER) 'IF 1PROV AND 1PP VECPROVIDERSTYPE(BAZ, BAZ, BAZ...)
    End If
    
    'Get Mineral purchase
    Set MIN = New ClassPurchase_current
    If Not PurchaseMineral_Core_Spreadmodule_generate(PVarRequest, MIN, RES) Then GoTo EH
    Set VecResponse(GLOLNG_ROWCOUNTER) = MIN
    Set MIN = Nothing
  
  Next GLOLNG_ROWCOUNTER

  'Return
  Let PVarResponse = VecResponse
  Let PurchaseMineral_Core_Main = True

EH:
  Let VecProviders = Empty
  Let VecProvidersType = Empty
  Let VecResponse = Empty
  Set MLT = Nothing
  Set MIN = Nothing
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

'Private Function PurchaseMineral_Core_Main(ByRef PVarRequest As Variant) As Boolean
'
'  Dim TypOinv As GLOTYPPURCHASE
'
'  Const sFN As String = "APPPRO::-core_main"
'
'  On Error GoTo EH
'  Let PurchaseMineral_Core_Main = False
'
'  'Resolution
'  Call PurchaseMineral_Core_assets_set_strlos
'  With TypOinv
'    'Title, purchase and provider datas
'    If Not PurchaseMineral_Core_Spreadmodule_generate_data_purchasebase(PVarRequest, .PVecTitle, .PVecInvoice, .PVecProvider) Then GoTo EH
'
'    'Ore data to ore weight
'    If Not PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal(PVarRequest, TypOinv) Then GoTo EH
'
'    'Return
'    Let GLOBOOBONUS = .PBooBonus
'    Let GLOVAR_PURCHASEMINERALRESULTS = Array(.PVecTitle, .PVecInvoice, .PVecProvider, .PVecOre, .PVecOrePay, .PVecGrossValue, .PVecDeductions, .PVecBonus, .PVecFinal, .PVecUnits)
'  End With
'
'  Let PurchaseMineral_Core_Main = True
'
'EH:
'  Call EHGLOBAL.ErrorHandlerRaise(sFN)
'End Function

Private Function PurchaseMineral_Core_Spreadmodule_generate(ByRef PVarRequest As Variant, ByRef PClsRequest As ClassPurchase_current, ByVal RES As AppResources_current) As Boolean
  
  Dim TypOinv As GLOTYPPURCHASE

  Const sFN As String = "APPMIN::-core_spreadmodule_generate"

  On Error GoTo EH
  Let PurchaseMineral_Core_Spreadmodule_generate = False

  'Resolution
  'Call PurchaseMineral_Core_assets_set_strlos
  With TypOinv
    'Title, purchase and provider datas
    If Not PurchaseMineral_Core_Spreadmodule_generate_data_purchasebase(PVarRequest, .PVecTitle, .PVecInvoice, .PVecProvider) Then GoTo EH
  
    'Ore data to ore weight
    If Not PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal(PVarRequest, TypOinv, RES) Then GoTo EH
  
    'Return
    Let PClsRequest.Title = .PVecTitle
    Let PClsRequest.Invoice = .PVecInvoice
    Let PClsRequest.Provider = .PVecProvider
    Let PClsRequest.Ore = .PVecOre
    Let PClsRequest.OrePayment = .PVecOrePay
    Let PClsRequest.GrossValue = .PVecGrossValue
    Let PClsRequest.Deductions = .PVecDeductions
    Let PClsRequest.Bonus = .PVecBonus
    Let PClsRequest.Final = .PVecFinal
    Let PClsRequest.Units = .PVecUnits
    Let PClsRequest.BooBonus = .PBooBonus
  End With
  
  Let PurchaseMineral_Core_Spreadmodule_generate = True

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)
  
End Function

Private Function PurchaseMineral_Core_Spreadmodule_generate_data_purchasebase( _
  ByRef PVarRequest As Variant, _
  ByRef VecDataTitle As Variant, _
  ByRef VecDataInvoice As Variant, _
  ByRef VecDataProvider As Variant) As Boolean

  'Dim RES As AppResources
  'Dim MDL As ModelExcliqliteDatasheet

  Dim StrAux0 As String, StrCurLocal As String, StrCurForeign As String

  Const STR_SIM As String = "Simulación"
  Const STR_DASH As String = "-"
  Const sFN As String = "APPMIN::-core_main_spreadmodule_generate_data_purchasebase"

  ' Request
  ' 0-StrPlace      5-StrProNames[]      10-BooNewSheet
  ' 1-StrDate       6-StrProType[]       11-RngBox
  ' 2-StrTC         7-BooAccurate        12-BooDataSheet: THIS PARAM COMES WITH THE MODELDATASHEET OBJECT AND MUST BE ASSINGNED AT MODEL GLOBAL OF THIS CONTROLLER
  ' 3-BooME         8-BooPage            13-Boo1PP: THIS BOOL SAYS THAT EACH PUR PER PAGE OR ALL PUR IN A PAGE
  ' 4-BooVarious    9-BooPrintTwo
  On Error GoTo EH
  Let PurchaseMineral_Core_Spreadmodule_generate_data_purchasebase = False

  'Unpack data
'  With GLOTYPLO
''    Set RES = New AppResources
''    Let StrAux0 = RES.resExcelVLockup("1", GLOSHTDATA.ListObjects(GLOTYPLO.PCurrencies).ListColumns("currency_main").Range, GLOSHTDATA.ListObjects(.PCurrencies).ListColumns("currencies").Range)
''    Let StrCurLocal = RES.resExcelVLockup("1", GLOSHTDATA.ListObjects(.PCurrencies).ListColumns("currency_main").Range, GLOSHTDATA.ListObjects(.PCurrencies).ListColumns("symbol").Range)
''    Let StrCurForeign = RES.resExcelVLockup("1", GLOSHTDATA.ListObjects(.PCurrencies).ListColumns("currency_foreign").Range, GLOSHTDATA.ListObjects(.PCurrencies).ListColumns("symbol").Range)
''    Set RES = Nothing
'  End With
  
'  If VBA.CStr(PVarRequest(12)) = "lite" Then
'    Set MLT = New ModelExcliqliteDatasheet
'    Set MLT.ErrorHandler = EHGLOBAL
'    Let StrAux0 = VBA.CStr(MLT.MGet(eqlMdlTblCurrencies, eqlMdlArray, "currencies", MStrWhere:="currency_main=1")(0, 0))
'    Let StrCurLocal = VBA.CStr(MLT.MGet(eqlMdlTblCurrencies, eqlMdlArray, "symbol", MStrWhere:="currency_main=1")(0, 0))
'    Let StrCurForeign = VBA.CStr(MLT.MGet(eqlMdlTblCurrencies, eqlMdlArray, "symbol", MStrWhere:="currency_foreign=1")(0, 0))
'    Set MLT = Nothing
'  Else
'    'Excliq
'  End If

  'Title
  ReDim VecDataTitle(0 To 1, 0 To 0)
  Let VecDataTitle(0, 0) = "LIQUIDACIÓN DE MINERALES"
  Let VecDataTitle(1, 0) = "(EXPRESADO EN " & VBA.UCase(GLOSTR_CURRENCIES) & ")"

  'Purchase data
  ReDim VecDataInvoice(0 To 6, 0 To 2)
  Let VecDataInvoice(0, 0) = "DATOS LIQUIDACIÓN"
  Let VecDataInvoice(1, 0) = "Destino:": Let VecDataInvoice(1, 1) = STR_SIM 'E
  Let VecDataInvoice(2, 0) = "Lugar:": Let VecDataInvoice(2, 1) = VBA.CStr(PVarRequest(0))
  Let VecDataInvoice(3, 0) = "Fecha:": Let VecDataInvoice(3, 1) = VBA.CDate(VBA.Format$(VBA.CStr(PVarRequest(1)), "dd/mm/yy"))
  Let VecDataInvoice(4, 0) = "Liquidación:": Let VecDataInvoice(4, 1) = STR_SIM 'E
  Let VecDataInvoice(5, 0) = "Lote:": Let VecDataInvoice(5, 1) = STR_SIM 'E
  Let VecDataInvoice(6, 0) = "T/C:": Let VecDataInvoice(6, 1) = PVarRequest(2)
  Let VecDataInvoice(6, 2) = GLOSTR_CURFOREIGN & "/" & GLOSTR_CURLOCAL

  'Provider data
  ReDim VecDataProvider(0 To 6, 0 To 2)
  Let VecDataProvider(0, 0) = "DATOS PROVEDOR"
  Let VecDataProvider(1, 0) = "Código:": Let VecDataProvider(1, 1) = STR_DASH 'E
  Let VecDataProvider(2, 0) = "Nombre:": Let VecDataProvider(2, 1) = VBA.CStr(PVarRequest(5))
  Let VecDataProvider(3, 0) = "C.I.:": Let VecDataProvider(3, 1) = STR_DASH 'E
  Let VecDataProvider(4, 0) = "Tipo:": Let VecDataProvider(4, 1) = VBA.CStr(PVarRequest(6))
  Let VecDataProvider(5, 0) = "NIT:": Let VecDataProvider(5, 1) = STR_DASH 'E
  Let VecDataProvider(6, 0) = "NIM:": Let VecDataProvider(6, 1) = STR_DASH 'E

  'Return
  Let PurchaseMineral_Core_Spreadmodule_generate_data_purchasebase = True

EH:
  'Set MDL = Nothing
  'Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal( _
  ByRef PVarRequest As Variant, _
  ByRef PTypResponse As GLOTYPPURCHASE, _
  ByVal RES As AppResources_current) As Boolean
  
  'Dim RES As AppResources

  Dim RngBox As Range
  
  Dim VecAux0 As Variant
  'Dim StrAux0 As String, GLOSTR_CURLOCAL As String, GLOSTR_CURFOREIGN As String, StrRngAux As String
  Dim BooSimple As Boolean, BooAverag As Boolean, BooME As Boolean
  Dim DblAux0 As Double, DblAux1 As Double, DblTC As Double
  Dim i As Long, j As Long, k As Long, l As Long, m As Long, n As Long, xRows As Long, xCols As Long, xRowCounter As Long

  Const STR_SIM As String = "Simulación"
  Const STR_DASH As String = "-"
  Const sFN As String = "APPMIN::-core_main_spreadmodule_generate_data_purchasebase"

  ' Request
  ' 0-StrPlace      5-StrProNames[]      10-BooNewSheet
  ' 1-StrDate       6-StrProType[]       11-RngBox
  ' 2-StrTC         7-BooAccurate        12-BooDataSheet: THIS PARAM COMES WITH THE MODELDATASHEET OBJECT AND MUST BE ASSINGNED AT MODEL GLOBAL OF THIS CONTROLLER
  ' 3-BooME         8-BooPage            13-Boo1PP: THIS BOOL SAYS THAT EACH PUR PER PAGE OR ALL PUR IN A PAGE
  ' 4-BooVarious    9-BooPrintTwo
  On Error GoTo EH
  Let PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal = False

  'Unpack data
  Let BooSimple = (VBA.CBool(PVarRequest(4)) = False) 'True: to one provider, print a purchase per sheet. False: to one provider, print all pur in one sheet
  Let BooAverag = VBA.CBool(PVarRequest(7)) 'Calculate with Averages?: Take averages per each mineral
  Let BooME = VBA.CBool(PVarRequest(3)) 'Calculate ME?: Calculates a ME column
  Let DblTC = VBA.CDbl(VBA.Replace(PVarRequest(2), ".", ","))
  
'  With GLOTYPLO
'    Set RES = New AppResources 'FIXME: THIS DATA MUST TO COME IN VECDATAINVOICE
'    Let StrAux0 = RES.resExcelVLockup("1", GLOSHTDATA.ListObjects(.PCurrencies).ListColumns("currency_main").Range, GLOSHTDATA.ListObjects(.PCurrencies).ListColumns("currencies").Range)
'    Let gloStr_CurLocal = RES.resExcelVLockup("1", GLOSHTDATA.ListObjects(.PCurrencies).ListColumns("currency_main").Range, GLOSHTDATA.ListObjects(.PCurrencies).ListColumns("symbol").Range)
'    Let gloStr_CurForeign = RES.resExcelVLockup("1", GLOSHTDATA.ListObjects(.PCurrencies).ListColumns("currency_foreign").Range, GLOSHTDATA.ListObjects(.PCurrencies).ListColumns("symbol").Range)
'    Set RES = Nothing
'  End With
  
  
  If GLOBOO_FORMULAS Then Set RngBox = PVarRequest(11)
  Let xCols = 4

  Dim StrTotalOreRng As String, StrTotalOrePayRng As String, StrTotalGrossValRng As String, StrTotalDeductionsRng As String, StrTotalBonusRng As String
  Dim DblTotalOre As Double, DblTotalOrePay As Double, DblTotalGrossVal As Double, DblTotalDeductions As Double, DblTotalBonus As Double
  
  With EHGLOBAL
    'Ore Data
    Let PTypResponse.PLngRowCounter = 3
    If Not PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal_getoredata(BooSimple, RngBox, PTypResponse, RES) Then
      Call .ErrorHandlerSet(1, "Error en la preparación de Datos del mineral."): GoTo EH
    End If
    
    'Ore payment contents
    If Not PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal_getorepay(BooSimple, BooAverag, RngBox, PTypResponse, RES) Then
      Call .ErrorHandlerSet(2, "Error en el cálculo del Peso pagable."): GoTo EH
    End If
  
    'Gross value
    If Not PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal_getgrossvalue(BooSimple, BooAverag, BooME, DblTC, GLOSTR_CURLOCAL, GLOSTR_CURFOREIGN, RngBox, PTypResponse, RES) Then
      Call .ErrorHandlerSet(3, "Error en el cálculo del Valor Bruto."): GoTo EH
    End If
  
    'Deductions
    If Not PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal_getdeductions(BooSimple, BooME, DblTC, RngBox, PTypResponse, RES) Then
      Call .ErrorHandlerSet(4, "Error en el cálculo de Deducciones."): GoTo EH
    End If
  
    'Bonus
    If Not PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal_getbonus(BooSimple, BooME, DblTC, RngBox, PTypResponse, RES) Then
      Call .ErrorHandlerSet(5, "Error en el cálculo de Bonos."): GoTo EH
    End If
  
    'Final purchase
    If Not PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal_getfinalpurchase(BooME, DblTC, GLOSTR_CURLOCAL, GLOSTR_CURFOREIGN, RngBox, PTypResponse, RES) Then
      Call .ErrorHandlerSet(6, "Error en el cálculo final de la Liquidación de Minerales."): GoTo EH
    End If
  End With
  If Not BooSimple Then Let GLOLNG_ROWCOUNTER = GLOLNG_ROWCOUNTER + 1

  'Return
  Let PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal = True

EH:
  Set RngBox = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal_getoredata( _
  ByVal PBooSimple As Boolean, _
  ByVal PRngBox As Range, _
  ByRef PTypResponse As GLOTYPPURCHASE, _
  ByVal RES As AppResources_current) As Boolean

  'Obtains the basic data to calculate the purchase
  Dim VecAux0 As Variant
  Dim StrAux0 As String
  Dim i As Long, j As Long, k As Long, xRows As Long, xCols As Long, xRowResize As Long, xRowTotals As Long

  Const sFN As String = "APPMIN::-core_main_spreadmodule_generate_data_purchasebase_getoredata"

  On Error GoTo EH
  Let PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal_getoredata = False

  With PTypResponse
    'Ore Data
    Let xCols = 4
    Let .PDblTotalOreVal = 0
    
    If PBooSimple Then 'Just one provider, [nPur[A1P[Normal|Accurate]|1/P] | 1Pur[1/P]] TODO: MÉTODO DE ROWCOUNT EN MDL
      'Let xRows = 2 + GLOSHTDATA.ListObjects(GLOTYPLO.PBasedata).ListRows.Count 'Title, Subtitles, [contents], totals
      Let xRows = 2 + MLT.LengthRows(eqlMdlTblPurchasebasedata) 'E 'Title, Subtitles, [contents], totals
      Let VecAux0 = Array(IIf(xRows > 2, "# ", VBA.vbNullString) & "Muestra", "Ingreso", "Contenidos", "Tipo", "PBH [TN]")
    Else 'n Providers
      Let xRows = 3 'Title, Subtitles, contents, totals
      Let VecAux0 = Array("Muestra", "Ingreso", "Contenidos", "Tipo", "PBH [T]")
    End If
    Let xRowResize = xRows - 2
    Let xRowTotals = xRows + 1
    
    '=
    Dim VecAux1 As Variant
    Dim LngIndex As Long
    Let LngIndex = IIf(PBooSimple, 0, GLOLNG_ROWCOUNTER + 1)
    'Fill matrix
    ReDim .PVecOre(0 To 0, 0 To xCols)
    Let .PVecOre(0, 0) = "DATOS MINERAL Y PESO BRUTO HÚMEDO"
    ' 1.  Add Subtitles at second row of main vector
    Let .PVecOre = RES.ArrayAddAtFirst(.PVecOre, VecAux0)
    ' 2.  Get all basedatatables
    Let VecAux0 = Empty
    Let VecAux0 = MLT.MGet(eqlMdlTblPurchasebasedata, eqlMdlArray, "MUESTRA, FECHA INGRESO, PESO HÚMEDO BRUTO", LngIndex) 'E
    If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(0, "Falló la lectura de datos base de la liquidación"): GoTo EH
    Let .PDblTotalOreVal = Application.WorksheetFunction.Sum(RES.ArrayToBaseZero(RES.ArrayDelIndex(VecAux0, UBound(VecAux0), RByCols:=True)))
    For i = 0 To UBound(VecAux0)
      Let VecAux0(i, 0) = IIf(PBooSimple, i & ". ", VBA.vbNullString) & VecAux0(i, 0)
    Next i
    ' 3.  Get all maindatatables
    Let VecAux1 = MLT.MGet(eqlMdlTblPurchasemaindata, eqlMdlArray, "CONTENIDOS, TIPO", LngIndex) 'E
    If Not VBA.IsArray(VecAux1) Then Call EHGLOBAL.ErrorHandlerSet(0, "Falló la lectura de datos del mineral de la liquidación"): GoTo EH
    ' 4.  Concat maindatatables in the third position of basedatatables
    Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1, 2, True)
    ' 5.  Add data to main vector
    Let .PVecOre = RES.ArrayAddAtLast(.PVecOre, VecAux0)
    ' 6.  Calculate totals
    Let VecAux0 = Empty
    ReDim VecAux0(0 To 0, 0 To xCols)
    Let VecAux0(0, 0) = "TOTAL PESO BRUTO HÚMEDO"
    If GLOBOO_FORMULAS Then Let VecAux0(0, xCols) = "=SUM(" & PRngBox.Range("I6").Resize(xRowResize, 1).Address(False, False) & ")"
    If Not GLOBOO_FORMULAS Then Let VecAux0(0, xCols) = .PDblTotalOreVal
    ' 7.  Add totals to main vector
    Let .PVecOre = RES.ArrayAddAtLast(.PVecOre, VecAux0)
    Let VecAux0 = Empty
    Let VecAux1 = Empty
    '=
    
    
'    ReDim .PVecOre(0 To xRows, 0 To xCols)
'    Let .PVecOre(0, 0) = "DATOS MINERAL Y PESO BRUTO HÚMEDO"
'    For j = 0 To UBound(.PVecOre, 2)
'      Let k = 1
'      For i = 1 To UBound(.PVecOre)
'        If i = 1 Then
'          Let .PVecOre(i, j) = VecAux0(j)
'        ElseIf i < xRows Then
'          Select Case j
'            Case 0 'Muestra
'              Let StrAux0 = IIf(PBooSimple, IIf(j = 0, IIf(xRows > 2, k & ". ", VBA.vbNullString), VBA.vbNullString), VBA.vbNullString)
'              Let .PVecOre(i, j) = StrAux0 & GLOSHTDATA.ListObjects(GLOTYPLO.PBasedata).ListColumns("MUESTRA").DataBodyRange.Cells(k, 1).Value
'              Let .PVecOre(i, j) = StrAux0 & GLOSHTDATA.ListObjects(GLOTYPLO.PBasedata).ListColumns("MUESTRA").DataBodyRange.Cells(k, 1).Value
'            Case 1 'Ingreso
'              Let .PVecOre(i, j) = GLOSHTDATA.ListObjects(GLOTYPLO.PBasedata).ListColumns("FECHA INGRESO").DataBodyRange.Cells(k, 1).Value
'            Case 2 'Contenidos
'              Let .PVecOre(i, j) = VBA.Replace(GLOSHTDATA.ListObjects(GLOTYPLO.PMaindata).ListColumns("CONTENIDOS").DataBodyRange.Cells(k, 1).Value, ";", ",")
'            Case 3 'TIPO
'              Let .PVecOre(i, j) = VBA.StrConv(GLOSHTDATA.ListObjects(GLOTYPLO.PMaindata).ListColumns("TIPO").DataBodyRange.Cells(k, 1).Value, vbProperCase)
'            Case Else 'PBH
'              Let .PVecOre(i, j) = GLOSHTDATA.ListObjects(GLOTYPLO.PBasedata).ListColumns("PESO HÚMEDO BRUTO").DataBodyRange.Cells(k, 1).Value
'              Let .PDblTotalOreVal = .PDblTotalOreVal + VBA.CDbl(.PVecOre(i, j))
'          End Select
'          Let k = k + 1
'        Else 'Totals
'          If j = 0 Then Let .PVecOre(i, j) = "TOTAL PESO BRUTO HÚMEDO"
'          If j = xCols Then
'            If GLOBOO_FORMULAS Then Let .PVecOre(i, j) = "=SUM(" & PRngBox.Range("I6").Resize(xRowResize, 1).address(False, False) & ")"
'            If Not GLOBOO_FORMULAS Then Let .PVecOre(i, j) = .PDblTotalOreVal
'          End If
'        End If
'      Next i
'    Next j
    If GLOBOO_FORMULAS Then Let .PStrTotalOreRng = "I" & (3 + xRowTotals)
    Let .PLngRowCounter = .PLngRowCounter + xRows + 1
  End With
'  Let VecAux0 = Empty

  'Return
  Let PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal_getoredata = True

EH:
  Let VecAux0 = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal_getorepay( _
  ByVal PBooSimple As Boolean, _
  ByVal PBooAverage As Boolean, _
  ByVal PRngBox As Range, _
  ByRef PTypResponse As GLOTYPPURCHASE, _
  ByVal RES As AppResources_current) As Boolean

  Dim VecAux0 As Variant
  Dim StrAux
  Dim DblAux As Double, DblAux1 As Double
  Dim i As Long, j As Long, k As Long, l As Long, xRows As Long, xCols As Long, xRowResize As Long, xRowTotals As Long, xRowFirstVal As Long

  Const sFN As String = "APPMIN::-core_main_spreadmodule_generate_data_purchasebase_getorepay"

  On Error GoTo EH
  Let PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal_getorepay = False

  With PTypResponse 'Ore Data
    Let xCols = 4
    Let .PDblTotalPayVal = 0
    Let VecAux0 = Array("#", "Humedad [%]", "PBS [T]", "Merma [%]", "PNS [TN]")
    
    'Determine matrix dimenssions
    If PBooAverage Then 'Just one provider with average in data
      Let xRows = 3 'Title, Subtitles, contents, totals
    Else 'Title, Subtitles, [contents], totals
      'Let xRows = 2 + IIf(PBooSimple, GLOSHTDATA.ListObjects(GLOTYPLO.PBasedata).ListRows.Count, 1)
      Let xRows = 2 + IIf(PBooSimple, MLT.LengthRows(eqlMdlTblPurchasebasedata), 1)
    End If
    Let xRowResize = xRows - 2
    Let xRowTotals = xRows + 1
    Let xRowFirstVal = .PLngRowCounter + 3
    
'    ReDim .PVecOrePay(0 To xRows, 0 To xCols)
'    Let .PVecOrePay(0, 0) = "DETERMINACIÓN PESO PAGABLE"
    
    ' ==
    ' Fill matrix
    Dim VecAux1 As Variant, VecAux2 As Variant, VecAux3 As Variant
    Dim DblAux0 As Double, DblAux2 As Double, DblAux3 As Double, DblAux4 As Double
    Dim LngIndex As Long
    Dim StrAux0 As String, StrAux1 As String, StrAux2 As String, StrAux3 As String, StrAux4 As String
    Let LngIndex = IIf(PBooSimple, 0, GLOLNG_ROWCOUNTER + 1)
    ReDim .PVecOrePay(0 To 0, 0 To xCols)
    Let .PVecOrePay(0, 0) = "DETERMINACIÓN PESO PAGABLE"
    Let .PVecOrePay = RES.ArrayAddAtLast(.PVecOrePay, Array("#", "Humedad [%]", "PBS [T]", "Merma [%]", "PNS [TN]"))
    ' 1.  Get PBH, H2O & LSS
    Let VecAux1 = MLT.MGet(eqlMdlTblPurchasebasedata, eqlMdlArray, "PESO HÚMEDO BRUTO, HUMEDAD, MERMA", LngIndex) 'E
    If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(0, "Falló la lectura de datos para determinar el peso pagable de la liquidación"): GoTo EH
    Let DblAux0 = 0: Let DblAux3 = 0: Let DblAux4 = 0
    If PBooAverage Then
    
      ReDim VecAux2(UBound(VecAux1))
      ReDim VecAux3(UBound(VecAux1))
      For i = 0 To UBound(VecAux1)
        Let DblAux1 = VBA.CDbl(VecAux1(i, 0)) * (1 - (VBA.CDbl(VecAux1(i, 1)) / 100)) 'PBS
        Let DblAux2 = DblAux1 * (1 - (VBA.CDbl(VecAux1(i, 2)) / 100)) 'PNS
        Let DblAux0 = DblAux0 + VBA.CDbl(VecAux1(i, 0)) 'PBH
        Let DblAux3 = DblAux3 + DblAux1 'PBS
        Let DblAux4 = DblAux4 + DblAux2 'PNS
        If GLOBOO_FORMULAS Then
          Let VecAux2(i) = PRngBox.Range(.PStrTotalOreRng).Offset(-xRows + 1 + i, 0).Address(False, False) & "*(1-(" & VBA.Replace(VecAux1(i, 1), ",", ".") & "/100))" 'PBS
          Let VecAux3(i) = PRngBox.Range("G" & xRowFirstVal + i).Address(False, False) & "*(1-(" & VBA.Replace(VecAux1(i, 2), ",", ".") & "/100))" 'PNS
        End If
      Next i
      ' 2.  Calculate Avg(H2O) & Avg(LSS)
      Let DblAux1 = (1 - (DblAux3 / DblAux0)) 'H2O
      Let DblAux2 = (1 - (DblAux4 / DblAux0)) 'LSS
      ' 3.  Fill main vector with results
      If GLOBOO_FORMULAS Then
        Let StrAux3 = "=1-(SUM(" & VBA.Join(VecAux2, ",") & ")/" & PRngBox.Range(.PStrTotalOreRng).Address(False, False) & ")" 'H2O
        Let StrAux4 = "=1-(SUM(" & VBA.Join(VecAux3, ",") & ")*100/" & PRngBox.Range(.PStrTotalOreRng).Address(False, False) & ")" 'LSS
        Let StrAux0 = "=" & PRngBox.Range(.PStrTotalOreRng).Address(False, False) & "*(1-(" & PRngBox.Range("F" & xRowFirstVal).Address(False, False) & "/100))"
        Let StrAux1 = "=" & PRngBox.Range("G" & xRowFirstVal).Address(False, False) & "*(1-(" & PRngBox.Range("H" & xRowFirstVal).Address(False, False) & "/100))"
        Let StrAux2 = "=SUM(" & PRngBox.Range("I" & xRowFirstVal).Resize(xRowResize, 1).Address(False, False) & ")"
      End If
      Let .PVecOrePay = RES.ArrayAddAtLast(.PVecOrePay, Array("Promedios", IIf(GLOBOO_FORMULAS, StrAux3, DblAux1), IIf(GLOBOO_FORMULAS, StrAux0, DblAux3), IIf(GLOBOO_FORMULAS, StrAux4, DblAux2), IIf(GLOBOO_FORMULAS, StrAux1, DblAux4)))
      Let .PVecOrePay = RES.ArrayAddAtLast(.PVecOrePay, Array("TOTAL PESO PAGABLE", Empty, Empty, Empty, IIf(GLOBOO_FORMULAS, StrAux2, DblAux4)))
    
    Else
      
      ReDim VecAux0(0 To UBound(VecAux1) + 1, 0 To xCols)
      For i = 0 To UBound(VecAux1)
        Let VecAux0(i, 0) = i + 1 & "." '#
        Let VecAux0(i, 1) = VecAux1(i, 1) 'H2O
        Let VecAux0(i, 3) = VecAux1(i, 2) 'LSS
        If GLOBOO_FORMULAS Then
          Let StrAux0 = PRngBox.Range(.PStrTotalOreRng).Offset(-xRows + 1 + i, 0).Address(False, False)
          Let VecAux0(i, 2) = "=" & StrAux & "*(1-(" & PRngBox.Range("F" & xRowFirstVal + i).Address(False, False) & "/100))" 'PBS
          Let VecAux0(i, 4) = "=" & PRngBox.Range("G" & xRowFirstVal + i).Address(False, False) & "*(1-(" & PRngBox.Range("H" & xRowFirstVal + i).Address(False, False) & "/100))" 'PNS
        Else
          Let VecAux0(i, 2) = VBA.CDbl(VecAux1(i, 0)) * (1 - (VBA.CDbl(VecAux1(i, 1)) / 100)) 'PBS
          Let VecAux0(i, 4) = VBA.CDbl(VecAux0(i, 2)) * (1 - (VBA.CDbl(VecAux1(i, 2)) / 100)) 'PNS
          Let .PDblTotalPayVal = .PDblTotalPayVal * VBA.CDbl(VecAux0(i, 4)) 'PNS TOTAL
        End If
      Next i
      Let VecAux0(UBound(VecAux0), 0) = "TOTAL PESO PAGABLE"
      Let VecAux0(UBound(VecAux0), xCols) = IIf(GLOBOO_FORMULAS, "=SUM(" & PRngBox.Range("I" & xRowFirstVal).Resize(xRowResize, 1).Address(False, False) & ")", .PDblTotalPayVal)
    
    End If
    'RangeString of Totals
    If GLOBOO_FORMULAS Then Let .PStrTotalPayRng = "I" & (.PLngRowCounter + xRowTotals)
    Let .PLngRowCounter = .PLngRowCounter + xRows + 1
    ' ==
  End With
    
    
    
'    'Fill matrix
'    If PBooAverage Then 'Just one provider with average in data
'
'      For j = 0 To UBound(.PVecOrePay, 2)
'        For i = 1 To UBound(.PVecOrePay)
'          If i = 1 Then 'Subtitles
'            Let .PVecOrePay(i, j) = VecAux0(j)
'          ElseIf i < xRows Then
'            Select Case j
'              Case 0 '#
'                Let .PVecOrePay(i, j) = "Promedios"
'              Case 1 'Humedad [%]
'                Let DblAux = Application.WorksheetFunction.Average(GLOSHTDATA.ListObjects(GLOTYPLO.PBasedata).ListColumns("HUMEDAD").DataBodyRange)
'                Let .PVecOrePay(i, j) = DblAux
'              Case 2 'PBS [T] xRowResize
'                If GLOBOO_FORMULAS Then
'                  Let .PVecOrePay(i, j) = "=" & PRngBox.Range(.PStrTotalOreRng).address(False, False) & "*(1-(" & PRngBox.Range("F" & xRowFirstVal).address(False, False) & "/100))"
'                Else
'                  Let DblAux1 = .PDblTotalOreVal * (1 - (DblAux / 100))
'                  Let .PVecOrePay(i, j) = DblAux1
'                End If
'              Case 3 'Merma [%]
'                Let DblAux = Application.WorksheetFunction.Average(GLOSHTDATA.ListObjects(GLOTYPLO.PBasedata).ListColumns("MERMA").DataBodyRange)
'                Let .PVecOrePay(i, j) = DblAux
'              Case Else 'PNS
'                If GLOBOO_FORMULAS Then
'                  Let .PVecOrePay(i, j) = "=" & PRngBox.Range("G" & xRowFirstVal).address(False, False) & "*(1-(" & PRngBox.Range("H" & xRowFirstVal).address(False, False) & "/100))"
'                Else
'                  Let .PVecOrePay(i, j) = DblAux1 * (1 - (DblAux / 100))
'                  Let .PDblTotalPayVal = .PDblTotalPayVal + VBA.CDbl(.PVecOrePay(i, j))
'                End If
'            End Select
'          Else 'totals
'            If j = 0 Then Let .PVecOrePay(i, j) = "TOTAL PESO PAGABLE"
'            If j = xCols Then
'              If GLOBOO_FORMULAS Then Let .PVecOrePay(i, j) = "=SUM(" & PRngBox.Range("I" & xRowFirstVal).Resize(xRowResize, 1).address(False, False) & ")"
'              If Not GLOBOO_FORMULAS Then Let .PVecOrePay(i, j) = .PDblTotalPayVal
'            End If
'          End If
'        Next i
'      Next j
'
'    Else 'Just one provider or multiple
'
'      For j = 0 To UBound(.PVecOrePay, 2)
'        For i = 1 To UBound(.PVecOrePay)
'          If i = 1 Then 'Subtitles
'            Let .PVecOrePay(i, j) = VecAux0(j)
'          ElseIf i < xRows Then
'            Let k = IIf(PBooSimple, i - 1, GLOLNG_ROWCOUNTER)
'            Select Case j
'              Case 0 '#
'                Let .PVecOrePay(i, j) = "'" & IIf(PBooSimple, k, 1) & "."
'              Case 1 'Humedad [%]
'                Let .PVecOrePay(i, j) = GLOSHTDATA.ListObjects(GLOTYPLO.PBasedata).ListColumns("HUMEDAD").DataBodyRange.Cells(k, 1).Value
'              Case 2 'PBS [T]
'                If GLOBOO_FORMULAS Then
'                  Let l = IIf(PBooSimple, 0, 1)
'                  Let StrAux = IIf(PBooSimple, PRngBox.Range(.PStrTotalOreRng).Offset(-xRows + 1, 0).Offset(k, 0).address(False, False), PRngBox.Range(.PStrTotalOreRng).address(False, False))
'                  Let .PVecOrePay(i, j) = "=" & StrAux & "*(1-(" & PRngBox.Range("F" & xRowFirstVal + (i - 2)).address(False, False) & "/100))"
'                Else
'                  Let DblAux = VBA.CDbl(GLOSHTDATA.ListObjects(GLOTYPLO.PBasedata).ListColumns("PESO HÚMEDO BRUTO").DataBodyRange.Cells(k, 1).Value)
'                  Let .PVecOrePay(i, j) = (DblAux * (1 - (VBA.CDbl(.PVecOrePay(i, j - 1))) / 100))
'                End If
'              Case 3 'Merma [%]
'                Let .PVecOrePay(i, j) = GLOSHTDATA.ListObjects(GLOTYPLO.PBasedata).ListColumns("MERMA").DataBodyRange.Cells(k, 1).Value
'              Case Else 'PNS
'                If GLOBOO_FORMULAS Then
'                  Let .PVecOrePay(i, j) = "=" & PRngBox.Range("G" & xRowFirstVal + (i - 2)).address(False, False) & "*(1-(" & PRngBox.Range("H" & xRowFirstVal + (i - 2)).address(False, False) & "/100))"
'                Else
'                  Let .PVecOrePay(i, j) = (VBA.CDbl(.PVecOrePay(i, j - 2)) * (1 - (VBA.CDbl(.PVecOrePay(i, j - 1))) / 100))
'                  Let .PDblTotalPayVal = .PDblTotalPayVal + VBA.CDbl(.PVecOrePay(i, j))
'                End If
'            End Select
'          Else 'totals
'            If j = 0 Then Let .PVecOrePay(i, j) = "TOTAL PESO PAGABLE"
'            If j = xCols Then
'              If GLOBOO_FORMULAS Then Let .PVecOrePay(i, j) = "=SUM(" & PRngBox.Range("I" & xRowFirstVal).Resize(xRowResize, 1).address(False, False) & ")"
'              If Not GLOBOO_FORMULAS Then Let .PVecOrePay(i, j) = .PDblTotalPayVal
'            End If
'          End If
'        Next i
'      Next j
'
'    End If
'    'RangeString of Totals
'    If GLOBOO_FORMULAS Then Let .PStrTotalPayRng = "I" & (.PLngRowCounter + xRowTotals) '+ IIf(PBooAverage, 1, 0)
'    Let .PLngRowCounter = .PLngRowCounter + xRows + 1
'  End With
'  Let VecAux0 = Empty

  'Return
  Let PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal_getorepay = True

EH:
  Let VecAux0 = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal_getgrossvalue( _
  ByVal PBooSimple As Boolean, _
  ByVal PBooAverage As Boolean, _
  ByVal PBooME As Boolean, _
  ByVal PDblTC As Double, _
  ByVal PStrLocal As String, _
  ByVal PStrForeign As String, _
  ByVal PRngBox As Range, _
  ByRef PTypResponse As GLOTYPPURCHASE, _
  ByVal RES As AppResources_current) As Boolean

  'Dim RES As AppResources
  
  Dim StrAux0 As String, StrAux1 As String, StrAux2 As String, StrAux3 As String, StrPNS As String
  Dim DblAux0 As Double, DblAux2 As Double, DblAux3 As Double, DblAux4 As Double
  Dim BooEquals As Boolean
  
  Dim VecAux0 As Variant, VecAux1 As Variant, VecAux2 As Variant, VecAux3 As Variant
  'Dim DblAux As Double
  Dim i As Long, j As Long, k As Long, l As Long, m As Long, n As Long, u As Long, xRows As Long, xCols As Long, xRowResize As Long, xRowTotals As Long, xRowFirstVal As Long

  Const sFN As String = "APPMIN::-core_main_spreadmodule_generate_data_purchasebase_getgrossvalue"

  On Error GoTo EH
  Let PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal_getgrossvalue = False

  With PTypResponse
    'Ore Data
    Let xCols = 4
    Let .PDblTotalGrossVal = 0
    'Let DblAux = 0
    Let VecAux0 = Array("#", "Ley", PStrForeign & "/PTO", IIf(PBooME, "Valor [" & PStrForeign & "]", VBA.vbNullString), "Valor [" & PStrLocal & "]")
    
    
    ' ==
    ' Fill matrix
    Dim REGEX As AppResRegEx
    'Dim VecAux1 As Variant, VecAux2 As Variant, VecAux3 As Variant
    Let xRowFirstVal = .PLngRowCounter + 3 ', xRowResize As Long, xRowTotals As Long, xRowFirstVal As Long
    Let xRows = (UBound(.PVecOrePay) - LBound(.PVecOrePay)) - 3 ', xRowResize As Long, xRowTotals As Long, xRowFirstVal As Long
    Dim LngIndex As Long
    'Dim StrAux0 As String, StrAux1 As String, StrAux2 As String, StrAux3 As String, StrAux4 As String
'  Dim i As Long, j As Long, k As Long, l As Long, u As Long
    
    Let LngIndex = IIf(PBooSimple, 0, GLOLNG_ROWCOUNTER + 1)
    'ReDim .PVecGrossValue(0 To 0, 0 To xCols)
    'Let .PVecGrossValue(0, 0) = "DETERMINACIÓN VALOR BRUTO"
    'Let .PVecGrossValue = RES.ArrayAddAtLast(.PVecGrossValue, Array("#", "Ley", PStrForeign & "/PTO", IIf(PBooME, "Valor [" & PStrForeign & "]", VBA.vbNullString), "Valor [" & PStrLocal & "]"))
    ' 1.  Get PBH, H2O, LSS; GRADE & PRICE...
      'Let StrAux0 = GLOSHTDATA.ListObjects(GLOTYPLO.PMaindata).ListColumns("TIPO").DataBodyRange.Cells(i, 1).Value
      'Let StrAux1 = GLOSHTDATA.ListObjects(GLOTYPLO.PMaindata).ListColumns("CONTENIDOS").DataBodyRange.Cells(i, 1).Value
      'Let StrAux2 = GLOSHTDATA.ListObjects(GLOTYPLO.PMaindata).ListColumns("LEYES").DataBodyRange.Cells(i, 1).Value
      'Let StrAux3 = GLOSHTDATA.ListObjects(GLOTYPLO.PMaindata).ListColumns("UNIDADES").DataBodyRange.Cells(i, 1).Value
      'Let StrAux4 = GLOSHTDATA.ListObjects(GLOTYPLO.PMaindata).ListColumns("PRECIO").DataBodyRange.Cells(i, 1).Value
      'Let StrAux5 = GLOSHTDATA.ListObjects(GLOTYPLO.PBasedata).ListColumns("PESO HÚMEDO BRUTO").DataBodyRange.Cells(i, 1).Value
    ' Get data from tables
    Let VecAux0 = MLT.MGet(eqlMdlTblPurchasebasedata, eqlMdlArray, "PESO HÚMEDO BRUTO, HUMEDAD, MERMA", LngIndex) 'E
    Let VecAux1 = MLT.MGet(eqlMdlTblPurchasemaindata, eqlMdlArray, MLngIndex:=LngIndex) 'E
    Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux0, RByCols:=True)
    If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(0, "Falló la lectura de datos para determinar el valor bruto de la liquidación"): GoTo EH
    Let VecAux1 = Empty
    
    ' 1.  Separate all complex
    Let i = 0
    Set REGEX = New AppResRegEx
    Do While i <= UBound(VecAux0)
      'PNS * GRD * PRC
      If GLOBOO_FORMULAS Then Let VecAux0(i, 0) = "(" & PRngBox.Range(.PStrTotalOreRng).Offset(-xRows + 1 + i, 0).Address(False, False) & "*(1-(" & VBA.Replace(VBA.CStr(VecAux0(i, 1)), ",", ".") & "*0.01)))*(1-(" & VBA.Replace(VBA.CStr(VecAux0(i, 2)), ",", ".") & "*0.01))"
      If REGEX.isMineralComplex(VBA.CStr(VecAux0(i, 3))) Then 'Complex
        Let VecAux2 = VBA.Split(VBA.CStr(VecAux0(i, 4)), ";")
        Let j = UBound(VecAux2) - LBound(VecAux2)
        For k = 0 To j
          Let VecAux3 = Array(VecAux0(i, 0), VecAux0(i, 1), VecAux0(i, 2), VecAux0(i, 3))
          For l = 4 To 7
            Let VecAux2 = VBA.Split(VBA.CStr(VecAux0(i, l)), ";")
            Let u = UBound(VecAux3) + 1
            ReDim Preserve VecAux3(u)
            Let VecAux3(u) = IIf(l = 4 And k > 0, "~", VBA.vbNullString) & VecAux2(k) 'Add identifier to complex materials
            Let VecAux2 = Empty
          Next l
          Let VecAux1 = RES.ArrayAddAtLast(VecAux1, VecAux3)
          Let VecAux3 = Empty
        Next k
      Else 'Normal
        Let VecAux1 = RES.ArrayAddAtLast(VecAux1, RES.ArrayDelIndex(VecAux0, i, True))
      End If
      Let i = i + 1
    Loop
    Set REGEX = Nothing
    
    ' 2.  Calculate section
    Let k = 1
    Let n = 0
    ReDim VecAux0(0 To 0, 0 To 4)
    Let VecAux0(0, 0) = "DETERMINACIÓN VALOR BRUTO"
    Let VecAux0 = RES.ArrayAddAtLast(VecAux0, Array("#", "Ley", PStrForeign & "/PTO", IIf(PBooME, "Valor [" & PStrForeign & "]", VBA.vbNullString), "Valor [" & PStrLocal & "]"))
    
    If PBooAverage Then 'Just one provider to average or calculate all its purchases
      For i = 0 To UBound(VecAux1)
        Let l = 1
        Let StrAux0 = VBA.Replace(VBA.CStr(VecAux1(i, 4)), "~", VBA.vbNullString)
        
        If GLOBOO_FORMULAS Then
          Let m = 0
          ReDim VecGrossA(0)
          ReDim VecGrossB(0)
          ReDim VecGross(0)
          Let StrPNS = VBA.CStr(VecAux1(i, 0))
          'PNS * PRC
          Let VecGrossA(m) = "(" & StrPNS & "*" & VBA.Replace(VBA.CStr(VecAux1(i, 5)), ",", ".") & "*" & VBA.Replace(VBA.CStr(VecAux1(i, 7)), ",", ".") & ")"
          Let VecGrossB(m) = "(" & StrPNS & ")"
          Let VecGross(m) = StrPNS
                       'Array(0:Name,  1:Grade,       2:Price,       3:Frgn,4:Local,      5:PNS,         6:Unit)
          Let VecAux2 = Array(StrAux0, VecAux1(i, 5), VecAux1(i, 7), Empty, "=" & StrPNS, VecAux1(i, 0), VecAux1(i, 6))
        Else
          Let DblAux0 = VBA.CDbl(VecAux1(i, 0)) * VBA.CDbl(VecAux1(i, 5)) * VBA.CDbl(VecAux1(i, 7)) * PDblTC 'PNS*GRD*PRC
                       'Array(0:Name,  1:Grade,       2:Price,       3:Frgn,4:Local,      4:PNS,         5:Unit)
          Let VecAux2 = Array(StrAux0, VecAux1(i, 5), VecAux1(i, 7), Empty, DblAux0, VecAux1(i, 0), VecAux1(i, 6))
        End If
        
        For j = k To UBound(VecAux1)
          Let BooEquals = (Not VBA.CStr(VecAux1(j, 4)) = VBA.vbNullString And Not VBA.CStr(VecAux1(i, 4)) = VBA.vbNullString) 'Contents
          If BooEquals Then
            'Content(i) = Content(j) and Unit(i) = Unit(j) ?
            Let BooEquals = ((VBA.LCase(VBA.Replace(VBA.CStr(VecAux1(i, 4)), "~", VBA.vbNullString)) = VBA.LCase(VBA.Replace(VBA.CStr(VecAux1(j, 4)), "~", VBA.vbNullString))) And (VBA.LCase(VBA.CStr(VecAux1(i, 6))) = VBA.LCase(VBA.CStr(VecAux1(j, 6)))))
            If BooEquals Then
              'PRC(i) = PRC(j) ?
              Let BooEquals = (VBA.LCase(VBA.CStr(VecAux1(i, 7))) = VBA.LCase(VBA.CStr(VecAux1(j, 7))))
              If GLOBOO_FORMULAS Then
                If BooEquals Or GLOBOO_ACCURATEANYWAY Then
                  Let StrPNS = VBA.CStr(VecAux1(j, 0))
                  Let m = m + 1
                  ReDim Preserve VecGrossA(m) 'PNS * GRD * PRC
                  Let VecGrossA(m) = "(" & StrPNS & "*" & VBA.Replace(VBA.CStr(VecAux1(j, 5)), ",", ".") & "*" & VBA.Replace(VBA.CStr(VecAux1(j, 7)), ",", ".") & ")"
                  ReDim Preserve VecGrossB(m) 'PNS * PRC
                  Let VecGrossB(m) = "(" & StrPNS & ")" '& "*" & VBA.Replace(VBA.CStr(VecAux1(j, 7)), ",", ".")
                  ReDim Preserve VecGross(m) 'SUM(PNS)
                  Let VecGross(m) = StrPNS
                End If
              Else
                Let DblAux0 = VBA.CDbl(VecAux1(j, 0)) 'PNS
                Let VecAux2(5) = VBA.CDbl(VecAux2(5)) + IIf(BooEquals Or GLOBOO_ACCURATEANYWAY, DblAux0, 0)
                Let DblAux0 = VBA.CDbl(VecAux1(j, 0)) * VBA.CDbl(VecAux1(j, 5)) * VBA.CDbl(VecAux1(j, 7)) * PDblTC 'GRS=PNS*GRD*PRC*TC
                Let VecAux2(4) = VBA.CDbl(VecAux2(4)) + IIf(BooEquals Or GLOBOO_ACCURATEANYWAY, DblAux0, 0)
              End If
              Let DblAux0 = VBA.CDbl(VecAux1(j, 7)) 'PRC
              Let VecAux2(2) = VBA.CDbl(VecAux2(2)) + IIf(BooEquals Or GLOBOO_ACCURATEANYWAY, DblAux0, 0)
              If BooEquals Or GLOBOO_ACCURATEANYWAY Then Let VecAux1(j, 4) = VBA.vbNullString
              If GLOBOO_ACCURATEANYWAY Then Let l = l + 1
              Let PBooAverage = (PBooAverage Or BooEquals Or GLOBOO_ACCURATEANYWAY)
            End If
          End If
        Next j
        
        If Not VecAux1(i, 4) = VBA.vbNullString Then 'Capture just valid items
          Let VecAux1(i, 4) = VBA.vbNullString
          If l > 1 Then Let VecAux2(2) = VBA.CDbl(VecAux2(2)) / l 'PRC
          Let VecAux2(0) = n + 1 & ". " & IIf(PBooAverage, "Promed. ", VBA.vbNullString) & VBA.StrConv(VBA.CStr(VecAux2(0)), vbProperCase) 'Content
          If GLOBOO_FORMULAS Then
            If PBooAverage Then
              Let VecAux2(1) = "=SUM(" & VBA.Join(VecGrossA, ",") & ")/(SUM(" & VBA.Join(VecGrossB, ",") & ")*" & VBA.Replace(VBA.CStr(VecAux2(2)), ",", ".") & ")" 'GRD
              Let VecAux2(4) = "=SUM(" & VBA.Join(VecGross, ",") & ")*" & PRngBox.Range("F" & xRowFirstVal + n).Address(False, False) & "*" & VBA.Replace(VBA.CStr(VecAux2(2)), ",", ".") & "*" & PRngBox.Range("B10").Address(False, False)
            Else
              Let VecAux2(4) = VBA.CStr(VecAux2(4)) & "*" & PRngBox.Range("F" & xRowFirstVal + n).Address(False, False) & "*" & PRngBox.Range("G" & xRowFirstVal + n).Address(False, False)
            End If
            If PBooME Then Let VecAux2(3) = "=" & PRngBox.Range("I" & xRowFirstVal + n).Address(False, False) & "/" & PRngBox.Range("B10").Address(False, False)
            Let VecGrossA = Empty
            Let VecGrossB = Empty
            Let VecGross = Empty
          Else
            Let VecAux2(1) = VBA.CDbl(VecAux2(4)) / (VBA.CDbl(VecAux2(2)) * VBA.CDbl(VecAux2(5)) * PDblTC)
            If PBooME Then Let VecAux2(3) = VBA.CDbl(VecAux2(4)) / PDblTC
            Let .PDblTotalGrossVal = .PDblTotalGrossVal + VBA.CDbl(VecAux2(4))
          End If
          Let n = n + 1
          Let VecAux0 = RES.ArrayAddAtLast(VecAux0, VecAux2)
        End If
        
        Let k = k + 1
      Next i
      
    Else 'One or more providers. Purchases 1:1
      
      For i = 0 To UBound(VecAux1)
        If GLOBOO_FORMULAS Then 'GROSS & GROSS ME
          Let VecAux1(i, 4) = VBA.CStr(VecAux1(i, 5)) & "*" & PRngBox.Range("F" & xRowFirstVal + n).Address(False, False) & "*" & PRngBox.Range("G" & xRowFirstVal + n).Address(False, False)
          If PBooME Then Let VecAux1(i, 3) = "=" & PRngBox.Range("I" & xRowFirstVal + n).Address(False, False) & "/" & PRngBox.Range("B10").Address(False, False)
        Else
          Let VecAux1(i, 4) = VBA.CDbl(VecAux1(i, 5)) * VBA.CDbl(VecAux1(i, 2)) * VBA.CDbl(VecAux1(i, 1)) * PDblTC
          If PBooME Then Let VecAux1(i, 3) = VBA.CDbl(VecAux1(i, 4)) / PDblTC
          Let .PDblTotalGrossVal = .PDblTotalGrossVal + VBA.CDbl(VecAux1(i, 4))
        End If
      Next i
    
    End If
    
    'Determine matrix dimenssions
    Let xRows = UBound(VecAux1) + 1 'Title, Subtitles, [contents], totals
    Let xRowResize = xRows - 2
    Let xRowTotals = xRows + 1
        
    'Totals
    If GLOBOO_FORMULAS Then
      If PBooME Then Let StrAux1 = "=SUM(" & PRngBox.Range("H" & xRowFirstVal).Resize(xRowResize, 1).Address(False, False) & ")"
      Let StrAux0 = "=SUM(" & PRngBox.Range("I" & xRowFirstVal).Resize(xRowResize, 1).Address(False, False) & ")"
    End If
    Let VecAux0 = RES.ArrayAddAtLast(VecAux0, Array("TOTAL VALOR BRUTO", Empty, Empty, IIf(PBooME, IIf(GLOBOO_FORMULAS, StrAux0, .PDblTotalGrossVal / PDblTC), Empty), IIf(GLOBOO_FORMULAS, StrAux0, .PDblTotalGrossVal), Empty, Empty))
    
    'Save Units
    Let .PVecUnits = RES.ArrayPop(VecAux0, RByCols:=True)
    Call RES.ArrayPop(VecAux0, RByCols:=True) 'Delete last column
    
    'Complete main vector
    Let .PVecGrossValue = RES.ArrayConcat(.PVecGrossValue, VecAux0)
    
    'RangeString of Totals
    If GLOBOO_FORMULAS Then Let .PStrTotalGrossRng = "I" & (.PLngRowCounter + xRowTotals)
    Let .PLngRowCounter = .PLngRowCounter + xRows + 1
    ' ==
    
    
    
    
    
    
'    'Get all unique contents and average of grades and prices
'    If PBooAverage Then 'Just one provider with average in data
'      If Not PurchaseMineral_Core_assets_get_average(.PVecUnits, VecAux1, RES) Then GoTo EH
'    Else
'      If Not PurchaseMineral_Core_assets_get_grades(PBooSimple, .PVecOrePay, VecAux1, .PVecUnits, RES) Then GoTo EH
'    End If
'
'    'Determine matrix dimenssions
'    Let xRows = 2 + UBound(VecAux1) + 1 'Title, Subtitles, [contents], totals
'    Let xRowResize = xRows - 2
'    Let xRowTotals = xRows + 1
'    Let xRowFirstVal = .PLngRowCounter + 3 ', xRowResize As Long, xRowTotals As Long, xRowFirstVal As Long
'    ReDim .PVecGrossValue(0 To xRows, 0 To xCols)
'
'    'Fill matrix
'    Let .PVecGrossValue(0, 0) = "DETERMINACIÓN VALOR BRUTO"
'    If PBooAverage Then 'Just one provider with average in data
'
'      For j = 0 To UBound(.PVecGrossValue, 2)
'        For i = 1 To UBound(.PVecGrossValue)
'          If i = 1 Then 'Subtitles
'            Let .PVecGrossValue(i, j) = VecAux0(j)
'          ElseIf i < xRows Then
'            Let k = i - 2
'            Select Case j
'              Case 0 '#
'                Let .PVecGrossValue(i, j) = "Promed. " & VecAux1(k, 3)
'              Case 1 'Leyes
'                Let .PVecGrossValue(i, j) = VBA.CDbl(VecAux1(k, 0))
'              Case 2 'Prices/Point
'                Let .PVecGrossValue(i, j) = VBA.CDbl(VecAux1(k, 1))
'              Case 3 'Value Foreign
'                If PBooME Then
'                  If GLOBOO_FORMULAS Then
'                    Let .PVecGrossValue(i, j) = "=" & PRngBox.Range("I" & xRowFirstVal + k).address(False, False) & "/" & PRngBox.Range("B10").address(False, False)
'                  Else
'                    Let .PVecGrossValue(i, j) = VBA.CDbl(VecAux1(k, 0)) * VBA.CDbl(VecAux1(k, 1)) * VBA.CDbl(VecAux1(k, 2)) '.PDblTotalPayVal
'                    Let DblAux = DblAux + .PVecGrossValue(i, j)
'                  End If
'                End If
'              Case Else 'Value Local
'                If GLOBOO_FORMULAS Then
'                  Let .PVecGrossValue(i, j) = "=" & PRngBox.Range("F" & xRowFirstVal + k).address(False, False) & "*" & PRngBox.Range("G" & xRowFirstVal + k).address(False, False) & "*" & VBA.Replace(VBA.CStr(VecAux1(k, 2)), ",", ".") & "*" & PRngBox.Range("B10").address(False, False)
'                Else
'                  Let .PVecGrossValue(i, j) = VBA.CDbl(VecAux1(k, 0)) * VBA.CDbl(VecAux1(k, 1)) * VBA.CDbl(VecAux1(k, 2)) * PDblTC '.PDblTotalPayVal * PDblTC
'                  Let .PDblTotalGrossVal = .PDblTotalGrossVal + VBA.CDbl(.PVecGrossValue(i, j))
'                End If
'            End Select
'          Else 'totals
'            If j = 0 Then Let .PVecGrossValue(i, j) = "TOTAL VALOR BRUTO"
'            If j >= (xCols - 1) Then
'              If GLOBOO_FORMULAS Then
'                If j = (xCols - 1) And PBooME Then Let .PVecGrossValue(i, j) = "=SUM(" & PRngBox.Range("H" & xRowFirstVal).Resize(xRowResize, 1).address(False, False) & ")"
'                If j = xCols Then Let .PVecGrossValue(i, j) = "=SUM(" & PRngBox.Range("I" & xRowFirstVal).Resize(xRowResize, 1).address(False, False) & ")"
'              Else
'                If j = (xCols - 1) And PBooME Then Let .PVecGrossValue(i, j) = DblAux
'                If j = xCols Then Let .PVecGrossValue(i, j) = .PDblTotalGrossVal
'              End If
'            End If
'          End If
'        Next i
'      Next j
'
'    Else 'Just one provider or multiple GLOLNG_ROWCOUNTER
'      Let l = 0
'      Let m = 2 - UBound(.PVecOrePay)
'      For j = 0 To UBound(.PVecGrossValue, 2)
'        For i = 1 To UBound(.PVecGrossValue)
'          If i = 1 Then 'Subtitles
'            Let .PVecGrossValue(i, j) = VecAux0(j)
'          ElseIf i < xRows Then
'            Let k = i - 2
'            Select Case j
'              Case 0 '#
'                If PBooSimple Then 'Let l = VBA.CLng(VecAux1(k))
'                  Let .PVecGrossValue(i, j) = VecAux1(k, 0) 'IIf(PBooSimple, l, 1)
'                Else
'                  Let .PVecGrossValue(i, j) = 1 'IIf(PBooSimple, l, 1)
'                End If
'              Case 1 'Leyes
'                Let .PVecGrossValue(i, j) = VBA.CDbl(VecAux1(k, 1))
'              Case 2 'Prices/Point
'                Let .PVecGrossValue(i, j) = VBA.CDbl(VecAux1(k, 2))
'              Case 3 'Value Foreign
'                If PBooME Then
'                  If GLOBOO_FORMULAS Then
'                    Let .PVecGrossValue(i, j) = "=" & PRngBox.Range("I" & xRowFirstVal + k).address(False, False) & "/" & PRngBox.Range("B10").address(False, False)
'                  Else
'                    Let .PVecGrossValue(i, j) = VBA.CDbl(VecAux1(k, 3)) ' * VBA.CDbl(VecAux3(k))
'                    Let DblAux = DblAux + .PVecGrossValue(i, j)
'                  End If
'                End If
'              Case Else 'Value Local
'                If GLOBOO_FORMULAS Then
'                  Let .PVecGrossValue(i, j) = "=" & PRngBox.Range("F" & xRowFirstVal + k).address(False, False) & "*" & PRngBox.Range("G" & xRowFirstVal + k).address(False, False) & "*" & PRngBox.Range(.PStrTotalPayRng).Offset(m, 0).Offset(l, 0).address(False, False) & "*" & PRngBox.Range("B10").address(False, False)
'                  If VBA.CBool(VecAux1(k, 4)) = False Then Let l = l + 1
'                Else
'                  Let .PVecGrossValue(i, j) = VBA.CDbl(VecAux1(k, 3)) * PDblTC
'                  Let .PDblTotalGrossVal = .PDblTotalGrossVal + VBA.CDbl(.PVecGrossValue(i, j))
'                End If
'            End Select
'          Else 'totals
'            If j = 0 Then Let .PVecGrossValue(i, j) = "TOTAL VALOR BRUTO"
'            If j >= (xCols - 1) Then
'              If GLOBOO_FORMULAS Then
'                If j = (xCols - 1) And PBooME Then Let .PVecGrossValue(i, j) = "=SUM(" & PRngBox.Range("H" & xRowFirstVal).Resize(xRowResize, 1).address(False, False) & ")"
'                If j = xCols Then Let .PVecGrossValue(i, j) = "=SUM(" & PRngBox.Range("I" & xRowFirstVal).Resize(xRowResize, 1).address(False, False) & ")"
'              Else
'                If j = (xCols - 1) And PBooME Then Let .PVecGrossValue(i, j) = DblAux
'                If j = xCols Then Let .PVecGrossValue(i, j) = .PDblTotalGrossVal
'              End If
'            End If
'          End If
'        Next i
'      Next j
'    End If
'    Let VecAux0 = Empty
'    Let VecAux1 = Empty
'    Let VecAux2 = Empty
'    Let VecAux3 = Empty
'
'    'RangeString of Totals
'    If GLOBOO_FORMULAS Then Let .PStrTotalGrossRng = "I" & (.PLngRowCounter + xRowTotals) '+ IIf(PBooAverage, 1, 0)
'    Let .PLngRowCounter = .PLngRowCounter + xRows + 1
  End With

  'Return
  Let PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal_getgrossvalue = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Let VecAux2 = Empty
  Let VecAux3 = Empty
  Let VecGrossA = Empty
  Let VecGrossB = Empty
  Let VecGross = Empty
  'Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal_getdeductions( _
  ByVal PBooSimple As Boolean, _
  ByVal PBooME As Boolean, _
  ByVal PDblTC As Double, _
  ByVal PRngBox As Range, _
  ByRef PTypResponse As GLOTYPPURCHASE, _
  ByVal RES As AppResources_current) As Boolean

  Dim VecAux0 As Variant, VecAux1 As Variant, VecAux2 As Variant, VecAux3 As Variant
  Dim StrAux0 As String, StrAux1 As String, StrAux2 As String, StrAux3 As String
  Dim BooRM As Boolean, BooRet As Boolean, BooRetOther As Boolean, BooDiscounts As Boolean
  Dim DblAux0 As Double, DblAux1 As Double, DblAux2 As Double
  Dim i As Long, j As Long, k As Long, xRows As Long, xCols As Long, xRowResize As Long, xRowTotals As Long, xRowFirstVal As Long

  Const sFN As String = "APPMIN::-core_main_spreadmodule_generate_data_purchasebase_getdeductions"

  On Error GoTo EH
  Let PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal_getdeductions = False
  
  With PTypResponse
    
    'Initial and base values
    Let .PDblTotalDedVal = 0
    Let xCols = 4
    Let xRows = 1 'Title, [Subtitles, contents], totals
    Let i = 0
    Let k = 0
    Let xRowFirstVal = .PLngRowCounter + 3
    
    'Calcualte discounts
    Dim LngIndex As Long
    Let LngIndex = IIf(PBooSimple, 0, GLOLNG_ROWCOUNTER + 1)
    
    'Get Contents and determine its RM's
    Let VecAux0 = MLT.MGet(eqlMdlTblPurchasemaindata, eqlMdlArray, "CONTENIDOS", LngIndex) 'E
    Let VecAux0 = RES.ArrayToBaseZero(VecAux0, eqlRes1D)
    If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(0, "¡Falló la lectura de datos de Regalías Mineras!"): GoTo EH
    Let VecAux0 = RES.ArrayUniques(VBA.Split(VBA.Join(VecAux0, ";"), ";"))
    ReDim VecAux1(0)
    For i = 0 To UBound(VecAux0)
      ReDim VecAux1(i)
      Let VecAux1(i) = MLT.MGet(eqlMdlTblConfigpurrm, eqlMdlArray, "VENTA INTERNA", MStrWhere:="ELEMENTO=" & VBA.CStr(VecAux0(i)))(0, 0)
    Next i
    'Fill matrix
    ReDim .PVecDeductions(0 To 1, 0 To xCols)
    Let .PVecDeductions(0, 0) = "(-) DEDUCCIONES"
    Let .PVecDeductions(0, 1) = "RETENCIONES DE LEY - REGALÍAS [%]:"
    Let j = IIf(GLOBOO_RESUMED_RETENTIONS, UBound(VecAux0), 0)
    For i = 0 To j
      If GLOBOO_RESUMED_RETENTIONS Then Let DblAux0 = Application.WorksheetFunction.Sum(VecAux1) '%Retentions
      If GLOBOO_RESUMED_RETENTIONS Then Let StrAux3 = "RM:"
      If Not GLOBOO_RESUMED_RETENTIONS Then Let DblAux0 = VBA.CDbl(VecAux0(i)) '%Retentions
      If Not GLOBOO_RESUMED_RETENTIONS Then Let StrAux3 = i + 1 & ". RM-" & VBA.CStr(VecAux0(i))
      Let DblAux0 = VBA.CDbl(VecAux1(i)) 'RM's
      Let StrAux0 = "=" & PRngBox.Range("I" & xRowFirstVal + i).Address(False, False) & "/" & PRngBox.Range("B10").Address(False, False) '$ME
      Let StrAux1 = "=" & PRngBox.Range(.PStrTotalGrossRng).Address(False, False) & "*" & PRngBox.Range("F" & xRowFirstVal + i).Address(False, False) & "/100" '$MN
      Let DblAux1 = (.PDblTotalGrossVal / PDblTC) * (DblAux0 / 100) '#ME
      Let DblAux2 = (.PDblTotalGrossVal) * (DblAux0 / 100) '#MN
      Let .PVecDeductions = RES.ArrayAddAtLast(.PVecDeductions, Array(StrAux3, VBA.Replace(DblAux0, ",", "."), Empty, IIf(PBooME, IIf(GLOBOO_FORMULAS, StrAux0, VBA.Replace(DblAux1, ",", ".")), Empty), IIf(GLOBOO_FORMULAS, StrAux1, VBA.Replace(DblAux2, ",", "."))))
    Next i
    Let .PDblTotalDedVal = .PDblTotalDedVal + DblAux2
    
    'Get Retentions and determine its %
    Let VecAux0 = MLT.MGet(eqlMdlTblConfigpurretentions, eqlMdlArray, "CONCEPTO, PORCENTAJE", LngIndex, "SELECTED=1") 'E
    If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(0, "¡Falló la lectura de datos de Retenciones de ley!"): GoTo EH
    'Fill matrix
    Let .PVecDeductions = RES.ArrayAddAtLast(.PVecDeductions, Array("RETENCIONES DE LEY - IMPUESTOS E INSTITUCIONALES [%]:", Empty, Empty, Empty, Empty))
    Let j = IIf(GLOBOO_RESUMED_RETENTIONS, UBound(VecAux0), 0)
    For i = 0 To j
      If GLOBOO_RESUMED_RETENTIONS Then Let DblAux0 = Application.WorksheetFunction.Sum(RES.ArrayPop(VecAux0, True, True)) '%Retentions
      If GLOBOO_RESUMED_RETENTIONS Then Let StrAux3 = "RESUMEN:"
      If Not GLOBOO_RESUMED_RETENTIONS Then Let DblAux0 = VBA.CDbl(VecAux0(i, 1)) '%Retentions
      If Not GLOBOO_RESUMED_RETENTIONS Then Let StrAux3 = i + 1 & ". " & VBA.CStr(VecAux0(i, 0))
      Let StrAux0 = "=" & PRngBox.Range("I" & xRowFirstVal + i).Address(False, False) & "/" & PRngBox.Range("B10").Address(False, False) '$ME
      Let StrAux1 = "=" & PRngBox.Range(.PStrTotalGrossRng).Address(False, False) & "*" & PRngBox.Range("F" & xRowFirstVal + i).Address(False, False) & "/100" '$MN
      Let DblAux1 = (.PDblTotalGrossVal / PDblTC) * (DblAux0 / 100) '#ME
      Let DblAux2 = (.PDblTotalGrossVal) * (DblAux0 / 100) '#MN
      Let .PVecDeductions = RES.ArrayAddAtLast(.PVecDeductions, Array(StrAux3, VBA.Replace(DblAux0, ",", "."), Empty, IIf(PBooME, IIf(GLOBOO_FORMULAS, StrAux0, VBA.Replace(DblAux1, ",", ".")), Empty), IIf(GLOBOO_FORMULAS, StrAux1, VBA.Replace(DblAux2, ",", "."))))
    Next i
    Let .PDblTotalDedVal = .PDblTotalDedVal + DblAux2
    
    'Get Retentions others and determine its %
    Let VecAux0 = MLT.MGet(eqlMdlTblConfigpurretentionsother, eqlMdlArray, "CONCEPTO, PORCENTAJE", LngIndex, "SELECTED=1") 'E
    If VBA.IsArray(VecAux0) Then
      'Fill matrix
      Let .PVecDeductions = RES.ArrayAddAtLast(.PVecDeductions, Array("RETENCIONES INSTITUCIONALES OTRAS [%]:", Empty, Empty, Empty, Empty))
      Let j = IIf(GLOBOO_RESUMED_RETENTIONSOTHERS, UBound(VecAux0), 0)
      For i = 0 To j
        If GLOBOO_RESUMED_RETENTIONSOTHERS Then Let DblAux0 = Application.WorksheetFunction.Sum(RES.ArrayPop(VecAux0, True, True)) '%Retentions
        If GLOBOO_RESUMED_RETENTIONSOTHERS Then Let StrAux3 = "RESUMEN:"
        If Not GLOBOO_RESUMED_RETENTIONSOTHERS Then Let DblAux0 = VBA.CDbl(VecAux0(i, 1)) '%Retentions
        If Not GLOBOO_RESUMED_RETENTIONSOTHERS Then Let StrAux3 = i + 1 & ". " & VBA.CStr(VecAux0(i, 0))
        Let StrAux0 = "=" & PRngBox.Range("I" & xRowFirstVal + i).Address(False, False) & "/" & PRngBox.Range("B10").Address(False, False) '$ME
        Let StrAux1 = "=" & PRngBox.Range(.PStrTotalGrossRng).Address(False, False) & "*" & PRngBox.Range("F" & xRowFirstVal + i).Address(False, False) & "/100" '$MN
        Let DblAux1 = (.PDblTotalGrossVal / PDblTC) * (DblAux0 / 100) '#ME
        Let DblAux2 = (.PDblTotalGrossVal) * (DblAux0 / 100) '#MN
        Let .PVecDeductions = RES.ArrayAddAtLast(.PVecDeductions, Array(StrAux3, VBA.Replace(DblAux0, ",", "."), Empty, IIf(PBooME, IIf(GLOBOO_FORMULAS, StrAux0, VBA.Replace(DblAux1, ",", ".")), Empty), IIf(GLOBOO_FORMULAS, StrAux1, VBA.Replace(DblAux2, ",", "."))))
      Next i
      Let .PDblTotalDedVal = .PDblTotalDedVal + DblAux2
    End If
    
    'Get Discounts and determine its value
    Let VecAux0 = MLT.MGet(eqlMdlTblConfigpurretentionsother, eqlMdlArray, MLngIndex:=LngIndex) 'E
    If VBA.IsArray(VecAux0) Then
      'Fill matrix
      Let .PVecDeductions = RES.ArrayAddAtLast(.PVecDeductions, Array("DESCUENTOS:", Empty, Empty, Empty, Empty))
      Let j = IIf(GLOBOO_RESUMED_DISCOUNTS, UBound(VecAux0), 0)
      For i = 0 To j
        If GLOBOO_RESUMED_DISCOUNTS Then Let DblAux0 = Application.WorksheetFunction.Sum(RES.ArrayPop(VecAux0, True, True)) '%Retentions
        If GLOBOO_RESUMED_DISCOUNTS Then Let StrAux3 = "Descuentos varios:"
        If Not GLOBOO_RESUMED_DISCOUNTS Then Let DblAux0 = VBA.CDbl(VecAux0(i, 1)) '%Retentions
        If Not GLOBOO_RESUMED_DISCOUNTS Then Let StrAux3 = i + 1 & ". " & VBA.CStr(VecAux0(i, 0))
        Let StrAux0 = "=" & PRngBox.Range("I" & xRowFirstVal + i).Address(False, False) & "/" & PRngBox.Range("B10").Address(False, False) '$ME
        Let DblAux1 = (DblAux0 / PDblTC) '#ME
        Let .PVecDeductions = RES.ArrayAddAtLast(.PVecDeductions, Array(StrAux3, Empty, Empty, IIf(PBooME, IIf(GLOBOO_FORMULAS, StrAux0, VBA.Replace(DblAux1, ",", ".")), Empty), VBA.Replace(DblAux0, ",", ".")))
      Next i
      Let .PDblTotalDedVal = .PDblTotalDedVal + DblAux0
    End If
    
    'Totals
    Let xRowResize = UBound(.PVecDeductions) + 1
    Let xRowTotals = xRowResize + 1
    If GLOBOO_FORMULAS Then
      If PBooME Then Let StrAux0 = "=SUM(" & PRngBox.Range("H" & xRowFirstVal).Resize(xRowResize, 1).Address(False, False) & ")"
      Let StrAux1 = "=SUM(" & PRngBox.Range("I" & xRowFirstVal).Resize(xRowResize, 1).Address(False, False) & ")"
    End If
    Let .PVecDeductions = RES.ArrayAddAtLast(.PVecDeductions, Array("TOTAL DEDUCCIONES", Empty, Empty, IIf(PBooME, IIf(GLOBOO_FORMULAS, StrAux0, VBA.Replace(.PDblTotalDedVal / PDblTC, ",", ".")), Empty), IIf(GLOBOO_FORMULAS, StrAux1, VBA.Replace(.PDblTotalDedVal, ",", "."))))
    
    'RangeString of Totals
    If GLOBOO_FORMULAS Then Let .PStrTotalDedRng = "I" & (.PLngRowCounter + 1)
    Let .PLngRowCounter = .PLngRowCounter + 1
    
    
    
    
    
    
    
'    'Initial and base values
'    Let .PDblTotalDedVal = 0
'    Let xCols = 4
'    Let xRows = 1 'Title, [Subtitles, contents], totals
'    Let i = 0
'    Let k = 0
'    Let xRowFirstVal = .PLngRowCounter + 3 ', xRowResize As Long, xRowTotals As Long, xRowFirstVal As Long
'
'    Let BooRM = PurchaseMineral_Core_assets_get_deductions_rm(PBooSimple, PBooME, PDblTC, PRngBox, xRows, PTypResponse, VecAux0, RES)
'    Let BooRet = PurchaseMineral_Core_assets_get_deductions_retentions(PBooSimple, PBooME, PDblTC, PRngBox, PTypResponse, xRows, VecAux1, RES)
'    Let BooRetOther = PurchaseMineral_Core_assets_get_deductions_retentionsothers(PBooSimple, PBooME, PDblTC, PRngBox, PTypResponse, xRows, VecAux2, RES)
'    Let BooDiscounts = PurchaseMineral_Core_assets_get_deductions_discounts(PBooSimple, PBooME, PDblTC, PRngBox, PTypResponse, xRows, VecAux3, RES)
'
'    'Possible error
'    If BooRM = False And BooRet = False And BooRetOther = False And BooDiscounts = False Then GoTo EH
'
'    'Determine matrix dimenssions
'    Let xRowResize = xRows - 2
'    Let xRowTotals = .PLngRowCounter + 1
'    ReDim .PVecDeductions(0 To xRows, 0 To xCols)
'
'    'Fill matrix
'    Let .PVecDeductions(0, 0) = "(-) DEDUCCIONES"
'
'    If BooRM Then 'RM
'      For j = 0 To UBound(.PVecDeductions, 2)
'        For i = 0 To UBound(VecAux0)
'          If VBA.IsNumeric(VecAux0(i, j)) Then
'            If Not VecAux0(i, j) = Empty Then Let .PVecDeductions(i + 1, j) = VBA.Replace(VecAux0(i, j), ",", ".")
'          Else
'            Let .PVecDeductions(i + 1, j) = VecAux0(i, j)
'          End If
'        Next i
'      Next j
'      Let i = i + 1
'    End If
'    Let VecAux0 = Empty
'
'    If BooRet Then 'Retentions
'      For j = 0 To UBound(.PVecDeductions, 2)
'        For k = i To UBound(VecAux1) + i
'          If VBA.IsNumeric(VecAux1(k - i, j)) Then
'            If Not VecAux1(k - i, j) = Empty Then Let .PVecDeductions(k, j) = VBA.Replace(VecAux1(k - i, j), ",", ".")
'          Else
'            Let .PVecDeductions(k, j) = VecAux1(k - i, j)
'          End If
'        Next k
'      Next j
'    Else
'      Let k = i
'    End If
'    Let VecAux1 = Empty
'
'    If BooRetOther Then 'Retentions others
'      For j = 0 To UBound(.PVecDeductions, 2)
'        For i = k To UBound(VecAux2) + k
'          If VBA.IsNumeric(VecAux2(i - k, j)) Then
'            If Not VecAux2(i - k, j) = Empty Then Let .PVecDeductions(i, j) = VBA.Replace(VecAux2(i - k, j), ",", ".")
'          Else
'            Let .PVecDeductions(i, j) = VecAux2(i - k, j)
'          End If
'        Next i
'      Next j
'    Else
'      Let i = k
'    End If
'    Let VecAux2 = Empty
'
'    If BooDiscounts Then 'Discounts
'      For j = 0 To UBound(.PVecDeductions, 2)
'        For k = i To UBound(VecAux3) + i
'          If VBA.IsNumeric(VecAux3(k - i, j)) Then
'            If Not VecAux3(k - i, j) = Empty Then Let .PVecDeductions(k, j) = VBA.Replace(VecAux3(k - i, j), ",", ".")
'          Else
'            Let .PVecDeductions(k, j) = VecAux3(k - i, j)
'          End If
'        Next k
'      Next j
'    Else
'      Let k = i
'    End If
'    Let VecAux3 = Empty
'
'    For j = 0 To UBound(.PVecDeductions, 2) 'totals
'      Let i = UBound(.PVecDeductions)
'      If j = 0 Then Let .PVecDeductions(i, j) = "TOTAL DEDUCCIONES"
'      If j >= (xCols - 1) Then
'        If GLOBOO_FORMULAS Then
'          If j = (xCols - 1) And PBooME Then Let .PVecDeductions(i, j) = "=SUM(" & PRngBox.Range("H" & xRowFirstVal).Resize(xRowResize, 1).address(False, False) & ")"
'          If j = xCols Then Let .PVecDeductions(i, j) = "=SUM(" & PRngBox.Range("I" & xRowFirstVal).Resize(xRowResize, 1).address(False, False) & ")"
'        Else
'          If j = (xCols - 1) And PBooME Then Let .PVecDeductions(i, j) = VBA.Replace(.PDblTotalDedVal / PDblTC, ",", ".")
'          If j = xCols Then Let .PVecDeductions(i, j) = VBA.Replace(.PDblTotalDedVal, ",", ".")
'        End If
'      End If
'    Next j
'
'    'RangeString of Totals
'    If GLOBOO_FORMULAS Then Let .PStrTotalDedRng = "I" & (.PLngRowCounter + 1)
'    Let .PLngRowCounter = .PLngRowCounter + 1
  End With
  
  'Return
  Let PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal_getdeductions = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Let VecAux2 = Empty
  Let VecAux3 = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal_getbonus( _
  ByVal PBooSimple As Boolean, _
  ByVal PBooME As Boolean, _
  ByVal PDblTC As Double, _
  ByVal PRngBox As Range, _
  ByRef PTypResponse As GLOTYPPURCHASE, _
  ByVal RES As AppResources_current) As Boolean

  'Dim RES As AppResources
  
  Dim VecAux0 As Variant, VecAux1 As Variant, VecAux2 As Variant, VecAux3 As Variant
  Dim StrAux0 As String, StrAux1 As String, StrAux2 As String, StrAux3 As String
  Dim DblAux As Double, DblAux0 As Double, DblAux1 As Double
  Dim BooBonus As Boolean
  Dim i As Long, j As Long, k As Long, xRows As Long, xCols As Long, xRowResize As Long, xRowTotals As Long, xRowFirstVal As Long

  Const sFN As String = "APPMIN::-core_main_spreadmodule_generate_data_purchasebase_getbonus"

  On Error GoTo EH
  Let PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal_getbonus = False
  
  With PTypResponse
    
    '====
    'Initial and base values
    Let .PDblTotalBonVal = 0
    Let xCols = 4
    Let xRows = 1 'Title, [Subtitles, contents], totals
    Let i = 0
    Let k = 0
    Let xRowFirstVal = .PLngRowCounter + 3
    
    'Calcualte bonus
    Dim LngIndex As Long
    Let LngIndex = IIf(PBooSimple, 0, GLOLNG_ROWCOUNTER + 1)
    
    'Fill matrix
    ReDim .PVecBonus(0 To 1, 0 To xCols)
    Let .PVecBonus(0, 0) = "(+) BONOS"
    'Get bONUS and determine its value
    Let VecAux0 = MLT.MGet(eqlMdlTblConpurbonus, eqlMdlArray, MLngIndex:=LngIndex) 'E
    Let .PBooBonus = VBA.IsArray(VecAux0)
    
    If .PBooBonus Then
      
      'Fill matrix
      Let j = IIf(GLOBOO_RESUMED_BONUS, UBound(VecAux0), 0)
      For i = 0 To j
        If GLOBOO_RESUMED_BONUS Then Let DblAux0 = Application.WorksheetFunction.Sum(RES.ArrayPop(VecAux0, True, True)) 'Bonus Values
        If GLOBOO_RESUMED_BONUS Then Let StrAux3 = "Bonos varios:"
        If Not GLOBOO_RESUMED_BONUS Then Let DblAux0 = VBA.CDbl(VecAux0(i, 1)) '#Bonus value
        If Not GLOBOO_RESUMED_BONUS Then Let StrAux3 = i + 1 & ". " & VBA.CStr(VecAux0(i, 0)) 'Bonus names
        Let StrAux0 = "=" & PRngBox.Range("I" & xRowFirstVal + i).Address(False, False) & "/" & PRngBox.Range("B10").Address(False, False) '$ME
        Let DblAux1 = (DblAux0 / PDblTC) '#ME
        Let .PVecBonus = RES.ArrayAddAtLast(.PVecBonus, Array(StrAux3, Empty, Empty, IIf(PBooME, IIf(GLOBOO_FORMULAS, StrAux0, VBA.Replace(DblAux1, ",", ".")), Empty), VBA.Replace(DblAux0, ",", ".")))
      Next i
      Let .PDblTotalBonVal = .PDblTotalBonVal + DblAux0
    
      'Totals
      Let xRows = UBound(.PVecBonus) + 1
      Let xRowResize = xRows - 2
      Let xRowTotals = xRowResize + 1
      If GLOBOO_FORMULAS Then
        If PBooME Then Let StrAux0 = "=SUM(" & PRngBox.Range("H" & xRowFirstVal).Resize(xRowResize, 1).Address(False, False) & ")"
        Let StrAux1 = "=SUM(" & PRngBox.Range("I" & xRowFirstVal).Resize(xRowResize, 1).Address(False, False) & ")"
      End If
      Let .PVecBonus = RES.ArrayAddAtLast(.PVecBonus, Array("TOTAL DEDUCCIONES", Empty, Empty, IIf(PBooME, IIf(GLOBOO_FORMULAS, StrAux0, VBA.Replace(.PDblTotalBonVal / PDblTC, ",", ".")), Empty), IIf(GLOBOO_FORMULAS, StrAux1, VBA.Replace(.PDblTotalBonVal, ",", "."))))
      
    End If
    
    'RangeString of Totals
    If GLOBOO_FORMULAS Then Let .PStrTotalBonRng = "I" & (.PLngRowCounter + xRowTotals)
    Let .PLngRowCounter = .PLngRowCounter + xRows
    '====
  
  
  
'
'
'    'Initial and base values
'    Let xCols = 4
'    Let xRows = 0
'    Let .PDblTotalBonVal = 0
'    Let DblAux = 0
'    Let .PBooBonus = False
'
'    Let BooBonus = (Not GLOSHTDATA.ListObjects(GLOTYPLO.PBonus).DataBodyRange Is Nothing)
'    If BooBonus Then
'      If PBooSimple Then
'        Set RES = New AppResources
'        Let StrAux0 = Application.WorksheetFunction.TextJoin(";", True, GLOSHTDATA.ListObjects(GLOTYPLO.PBonus).ListColumns("VALOR").DataBodyRange)
'        Let VecAux2 = 0 'RES.resArrayGetDataArrayFromVariants(StrAux0, xRows, False, True, True)
'        Let StrAux0 = Application.WorksheetFunction.TextJoin(";", True, GLOSHTDATA.ListObjects(GLOTYPLO.PBonus).ListColumns("BONO").DataBodyRange)
'        Let VecAux3 = 0 'RES.resArrayGetDataArrayFromVariants(StrAux0, xRows, True)
'        Set RES = Nothing
'      Else 'Múltiple
'        Let StrAux0 = GLOSHTDATA.ListObjects(GLOTYPLO.PBonus).ListColumns("VALOR").DataBodyRange.Cells(GLOLNG_ROWCOUNTER, 1).Value
'        If VBA.Len(StrAux0) > 0 Then
'          Let VecAux2 = VBA.Split(StrAux0, ";")
'          Let xRows = UBound(VecAux2) + 1 'Subtitle + content
'          Let StrAux0 = GLOSHTDATA.ListObjects(GLOTYPLO.PBonus).ListColumns("BONO").DataBodyRange.Cells(GLOLNG_ROWCOUNTER, 1).Value
'          Let VecAux3 = VBA.Split(StrAux0, ";")
'        End If
'      End If
'      If xRows > 0 Then
'        Let j = 0
'        ReDim VecAux0(0)
'        ReDim VecAux1(0)
'        For i = 0 To UBound(VecAux2)
'          If Not VBA.CDbl(VecAux2(i)) = 0 Then
'            ReDim Preserve VecAux0(j)
'            ReDim Preserve VecAux1(j)
'            Let VecAux0(j) = VecAux3(i) 'Concept
'            Let VecAux1(j) = VecAux2(i) 'Value
'            Let j = j + 1
'          End If
'        Next i
'        Let xRows = j  'UBound(VecAux0) + 1 'Subtitle + content
'        Let VecAux2 = Empty
'        Let VecAux3 = Empty
'      End If
'    End If
'
'    'Determine matrix dimenssions
'    If xRows <= 0 Then Let PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal_getbonus = True: GoTo EH 'If doesn't exists bonus
'    Let xRows = xRows + 1 'Title, Subtitle + content, Totals
'    Let xRowResize = xRows - 1
'    Let xRowTotals = xRows + 1
'    Let xRowFirstVal = .PLngRowCounter + 2 ', xRowResize As Long, xRowTotals As Long, xRowFirstVal As Long
'    ReDim .PVecBonus(0 To xRows, 0 To xCols)
'
'    'Fill matrix
'    Let .PVecBonus(0, 0) = "(+) BONOS"
'    For j = 0 To UBound(.PVecBonus, 2)
'      For i = 1 To xRows
'        If i < xRows Then
'          Let k = i - 1
'          Select Case j
'            Case 0 'Concept
'              Let .PVecBonus(i, j) = VBA.CStr(VecAux0(k))
'            Case 3 'Value Foreign
'              If PBooME Then
'                If GLOBOO_FORMULAS Then
'                  Let .PVecBonus(i, j) = "=" & PRngBox.Range("I" & xRowFirstVal + k).address(False, False) & "/" & PRngBox.Range("B10").address(False, False)
'                Else
'                  Let .PVecBonus(i, j) = VBA.CDbl(VecAux1(k)) / PDblTC
'                  Let DblAux = DblAux + VBA.CDbl(.PVecBonus(i, j))
'                End If
'              End If
'            Case 4 'Value Local
'              Let .PVecBonus(i, j) = VBA.CDbl(VecAux1(k))
'              Let .PDblTotalBonVal = .PDblTotalBonVal + VBA.CDbl(.PVecBonus(i, j))
'          End Select
'        Else 'totals
'          If j = 0 Then Let .PVecBonus(i, j) = "TOTAL BONOS"
'          If j >= (xCols - 1) Then
'            If GLOBOO_FORMULAS Then
'              If j = (xCols - 1) And PBooME Then Let .PVecBonus(i, j) = "=SUM(" & PRngBox.Range("H" & xRowFirstVal).Resize(xRowResize, 1).address(False, False) & ")"
'              If j = xCols Then Let .PVecBonus(i, j) = "=SUM(" & PRngBox.Range("I" & xRowFirstVal).Resize(xRowResize, 1).address(False, False) & ")"
'            Else
'              If j = (xCols - 1) And PBooME Then Let .PVecBonus(i, j) = DblAux
'              If j = xCols Then Let .PVecBonus(i, j) = .PDblTotalBonVal
'            End If
'          End If
'        End If
'      Next i
'    Next j
'    Let VecAux0 = Empty
'    Let VecAux1 = Empty
'
'    'RangeString of Totals
'    If GLOBOO_FORMULAS Then Let .PStrTotalBonRng = "I" & (.PLngRowCounter + xRowTotals)
'    Let .PLngRowCounter = .PLngRowCounter + xRows
'    Let .PBooBonus = True
  End With

  'Return
  Let PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal_getbonus = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Let VecAux2 = Empty
  Let VecAux3 = Empty
  'Set RES = Nothing
  Debug.Print VBA.Err.Description
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal_getfinalpurchase( _
  ByVal PBooME As Boolean, _
  ByVal PDblTC As Double, _
  ByVal PStrLocal As String, _
  ByVal PStrForeign As String, _
  ByVal PRngBox As Range, _
  ByRef PTypResponse As GLOTYPPURCHASE, _
  ByVal RES As AppResources_current) As Boolean
  
  Dim NTX As AppResNumberToString
  
  Dim StrBonusRng As String
  Dim DblAux0 As Double
  Dim xRows As Long, xCols As Long, xRowFirstVal As Long

  Const sFN As String = "APPMIN::-core_main_spreadmodule_generate_data_purchasebase_getbonus"

  On Error GoTo EH
  Let PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal_getfinalpurchase = False

  With PTypResponse
    
    'Initial and base values
    Let xCols = 8
    Let DblAux0 = 0
    
    'Determine matrix dimenssions
    Let xRows = 3 'Title, (Subtitles, contents), totals
    Let xRowFirstVal = .PLngRowCounter + 3 + IIf(.PDblTotalBonVal = 0, 0, 1)
    
    'Fill matrix PLngRowCounter
    Let DblAux0 = (.PDblTotalGrossVal - .PDblTotalDedVal) + .PDblTotalBonVal
    ReDim .PVecFinal(0 To xRows, 0 To xCols)
    
    Let .PVecFinal(0, 0) = "LIQUIDACIÓN FINAL (VALOR BRUTO - DEDUCCIONES" & IIf(Not .PDblTotalBonVal = 0, " + BONOS", VBA.vbNullString) & ")"
    If PBooME Then 'ME
      Let .PVecFinal(0, 7) = "VALOR " & PStrForeign
      If GLOBOO_FORMULAS Then
        Let .PVecFinal(1, 7) = "=" & PRngBox.Range("I" & xRowFirstVal).Address(False, False) & "/" & PRngBox.Range("B10").Address(False, False)
      Else
        Let .PVecFinal(1, 7) = VBA.Replace(DblAux0 / PDblTC, ",", ".")
      End If
    End If
    
    Let .PVecFinal(0, 8) = "VALOR " & PStrLocal 'MN
    If GLOBOO_FORMULAS Then
      If .PDblTotalBonVal = 0 Then
        Let StrBonusRng = VBA.vbNullString
      Else
        Let StrBonusRng = "," & PRngBox.Range(.PStrTotalBonRng).Address(False, False)
      End If
      Let .PVecFinal(1, 8) = "=SUM(" & PRngBox.Range(.PStrTotalGrossRng).Address(False, False) & ",-" & PRngBox.Range(.PStrTotalDedRng).Address(False, False) & StrBonusRng & ")"
    Else
      Let .PVecFinal(1, 8) = VBA.Replace(DblAux0, ",", ".")
    End If
    
    Let .PVecFinal(1, 0) = "Líquido pagable (" & IIf(DblAux0 = 0, "Sin saldo", IIf(DblAux0 > 0, "Saldo a favor", "Saldo en contra")) & "):"
    Let .PVecFinal(2, 0) = "Son:"
    If GLOBOO_FORMULAS Then
      Let .PVecFinal(2, 1) = "=EQL_NUMERO_A_TEXTO(" & PRngBox.Range("I" & xRowFirstVal).Address(False, False) & ",""" & VBA.UCase(GLOSTR_CURRENCY) & """,""" & VBA.UCase(GLOSTR_CURRENCIES) & """,0)"
    Else
      Set NTX = New AppResNumberToString
      Let .PVecFinal(2, 1) = NTX.NumToString(DblAux0, GLOSTR_CURRENCY, GLOSTR_CURRENCIES, False)
      Set NTX = Nothing
    End If
  End With
  
  'Return
  Let PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal_getfinalpurchase = True

EH:
  Set NTX = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)
  
  
  
  
  
  
  
  
  
  
'  'Dim RES As AppResources
'  Dim NTX As AppResNumberToString
'
'  Dim StrAux0 As String, StrAux1 As String, StrBonusRng As String
'  Dim DblAux0 As Double
'  Dim xRows As Long, xCols As Long, xRowFirstVal As Long
'
'  Const sFN As String = "APPMIN::-core_main_spreadmodule_generate_data_purchasebase_getbonus"
'
'  On Error GoTo EH
'  Let PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal_getfinalpurchase = False
'
'  With PTypResponse
'    'Initial and base values
'    Let xCols = 8
'    Let DblAux0 = 0
'    With GLOTYPLO
'      Set RES = New AppResources
'      Let StrAux0 = 0 'RES.resExcelVLockup(PStrLocal, GLOSHTDATA.ListObjects(.PCurrencies).ListColumns("symbol").DataBodyRange, GLOSHTDATA.ListObjects(.PCurrencies).ListColumns("currency").DataBodyRange)
'      Let StrAux1 = 0 'RES.resExcelVLockup(PStrLocal, GLOSHTDATA.ListObjects(.PCurrencies).ListColumns("symbol").DataBodyRange, GLOSHTDATA.ListObjects(.PCurrencies).ListColumns("currencies").DataBodyRange)
'      Set RES = Nothing
'    End With
'
'    'Determine matrix dimenssions
'    Let xRows = 3 'Title, (Subtitles, contents), totals
'    Let xRowFirstVal = .PLngRowCounter + 3 + IIf(.PDblTotalBonVal = 0, 0, 1)
'
'    'Fill matrix PLngRowCounter
'    Let DblAux0 = (.PDblTotalGrossVal - .PDblTotalDedVal) + .PDblTotalBonVal
'    ReDim .PVecFinal(0 To xRows, 0 To xCols)
'
'    Let .PVecFinal(0, 0) = "LIQUIDACIÓN FINAL (VALOR BRUTO - DEDUCCIONES" & IIf(Not .PDblTotalBonVal = 0, " + BONOS", VBA.vbNullString) & ")"
'    If PBooME Then 'ME
'      Let .PVecFinal(0, 7) = "VALOR " & PStrForeign
'      If GLOBOO_FORMULAS Then
'        Let .PVecFinal(1, 7) = "=" & PRngBox.Range("I" & xRowFirstVal).address(False, False) & "/" & PRngBox.Range("B10").address(False, False)
'      Else
'        Let .PVecFinal(1, 7) = VBA.Replace(DblAux0 / PDblTC, ",", ".")
'      End If
'    End If
'
'    Let .PVecFinal(0, 8) = "VALOR " & PStrLocal 'MN
'    If GLOBOO_FORMULAS Then
'      If .PDblTotalBonVal = 0 Then
'        Let StrBonusRng = VBA.vbNullString
'      Else
'        Let StrBonusRng = "," & PRngBox.Range(.PStrTotalBonRng).address(False, False)
'      End If
'      Let .PVecFinal(1, 8) = "=SUM(" & PRngBox.Range(.PStrTotalGrossRng).address(False, False) & ",-" & PRngBox.Range(.PStrTotalDedRng).address(False, False) & StrBonusRng & ")"
'    Else
'      Let .PVecFinal(1, 8) = VBA.Replace(DblAux0, ",", ".")
'    End If
'
'    Let .PVecFinal(1, 0) = "Líquido pagable (" & IIf(DblAux0 = 0, "Sin saldo", IIf(DblAux0 > 0, "Saldo a favor", "Saldo en contra")) & "):"
'    Let .PVecFinal(2, 0) = "Son:"
'    If GLOBOO_FORMULAS Then
'      Let .PVecFinal(2, 1) = "=EQL_NUMERO_A_TEXTO(" & PRngBox.Range("I" & xRowFirstVal).address(False, False) & ",""" & VBA.UCase(StrAux0) & """,""" & VBA.UCase(StrAux1) & """,0)"
'    Else
'      Set NTX = New AppResNumberToString
'      Let .PVecFinal(2, 1) = NTX.NumToString(DblAux0, StrAux0, StrAux1, False)
'      Set NTX = Nothing
'    End If
'  End With
'
'  'Return
'  Let PurchaseMineral_Core_Spreadmodule_generate_data_purchasefinal_getfinalpurchase = True
'
'EH:
'  Set NTX = Nothing
'  'Set RES = Nothing
'  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function









' ASSETS ----------------------------------------------------------------------------'
Private Function PurchaseMineral_Core_assets_get_average( _
  ByRef PVarUnits As Variant, _
  ByRef PVarContents As Variant, _
  ByVal RES As AppResources_current) As Boolean

  'Dim RES As AppResources
  Dim REGEX As AppResRegEx
  
  Dim VecAux0 As Variant, VecAux1 As Variant, VecAux2 As Variant, VecAux3 As Variant, VecAux4 As Variant, VecAux5 As Variant, VecAux6 As Variant
  Dim VecAux7 As Variant, VecAux8 As Variant
  Dim StrAux0 As String, StrAux1 As String, StrAux2 As String, StrAux3 As String, StrAux4 As String, StrAux5 As String
  Dim DblAux0 As Double, DblAux1 As Double
  Dim i As Long, j As Long, k As Long, l As Long, xRows As Long, xCols As Long

  Const sFN As String = "APPMIN::-core_assets_get_average"

  On Error GoTo EH
  Let PurchaseMineral_Core_assets_get_average = False
  Dim VecContent As Variant, VecGrade As Variant, VecUnit As Variant, VecPrice As Variant, VecWeight As Variant
  
  'Save in arrays each element non complex with its weights, grades and unit
  'If it is complex, split it and repeat it weight, then, save it in the array
  Set REGEX = New AppResRegEx
  ReDim VecContent(0)
  ReDim VecGrade(0)
  ReDim VecUnit(0)
  ReDim VecPrice(0)
  ReDim VecWeight(0)
  Let j = 0
  Let xRows = GLOSHTDATA.ListObjects(GLOTYPLO.PMaindata).ListRows.Count
  For i = 1 To xRows
    With GLOTYPLO
      Let StrAux0 = GLOSHTDATA.ListObjects(GLOTYPLO.PMaindata).ListColumns("TIPO").DataBodyRange.Cells(i, 1).Value
      Let StrAux1 = GLOSHTDATA.ListObjects(GLOTYPLO.PMaindata).ListColumns("CONTENIDOS").DataBodyRange.Cells(i, 1).Value
      Let StrAux2 = GLOSHTDATA.ListObjects(GLOTYPLO.PMaindata).ListColumns("LEYES").DataBodyRange.Cells(i, 1).Value
      Let StrAux3 = GLOSHTDATA.ListObjects(GLOTYPLO.PMaindata).ListColumns("UNIDADES").DataBodyRange.Cells(i, 1).Value
      Let StrAux4 = GLOSHTDATA.ListObjects(GLOTYPLO.PMaindata).ListColumns("PRECIO").DataBodyRange.Cells(i, 1).Value
      Let StrAux5 = GLOSHTDATA.ListObjects(GLOTYPLO.PBasedata).ListColumns("PESO HÚMEDO BRUTO").DataBodyRange.Cells(i, 1).Value
    End With
    ReDim Preserve VecContent(j)
    ReDim Preserve VecGrade(j)
    ReDim Preserve VecUnit(j)
    ReDim Preserve VecPrice(j)
    ReDim Preserve VecWeight(j)
    If REGEX.isMineralComplex(StrAux0) Then 'Is complex
      Let VecAux0 = VBA.Split(StrAux1, ";") 'Contents
      Let VecAux1 = VBA.Split(StrAux2, ";") 'Grades
      Let VecAux2 = VBA.Split(StrAux3, ";") 'Units
      Let VecAux3 = VBA.Split(StrAux4, ";") 'Prices
      For k = 0 To UBound(VecAux1)
        ReDim Preserve VecContent(j)
        ReDim Preserve VecGrade(j)
        ReDim Preserve VecUnit(j)
        ReDim Preserve VecPrice(j)
        ReDim Preserve VecWeight(j)
        Let VecContent(j) = VecAux0(k)
        Let VecGrade(j) = VBA.CDbl(VecAux1(k))
        Let VecUnit(j) = VBA.CStr(VecAux2(k))
        Let VecPrice(j) = VBA.CDbl(VecAux3(k))
        Let VecWeight(j) = VBA.CDbl(StrAux5)
        Let j = j + 1
      Next k
    Else 'Is mineral or broza
      Let VecContent(j) = StrAux1
      Let VecGrade(j) = VBA.CDbl(StrAux2)
      Let VecUnit(j) = VBA.CStr(StrAux3)
      Let VecPrice(j) = VBA.CDbl(StrAux4)
      Let VecWeight(j) = VBA.CDbl(StrAux5)
    End If
    Let j = j + 1
  Next i
  Set REGEX = Nothing
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Let VecAux2 = Empty
  Let VecAux3 = Empty
  
  'With the filled array, run across it and identify repeateds, then, average its weights and grades
  'Return array with all new values
  ReDim VecAux5(0) 'Grades
  ReDim VecAux6(0) 'Prices
  ReDim VecAux7(0) 'Weights
  ReDim VecAux8(0) 'Contents
  ReDim VecAux0(0) 'Grades
  ReDim VecAux1(0) 'Prices
  ReDim VecAux2(0) 'Weights
  ReDim PVarUnits(0) 'units
  Let k = 0
  Let l = 0
  For i = 0 To UBound(VecContent)
    If Not (VBA.CStr(VecContent(i)) = VBA.vbNullString) Then
      ReDim Preserve VecAux5(k)
      ReDim Preserve VecAux6(k)
      ReDim Preserve VecAux7(k)
      ReDim Preserve VecAux8(k)
      ReDim Preserve PVarUnits(k)
      Let VecAux5(k) = VecGrade(i)
      Let VecAux6(k) = VecPrice(i)
      Let VecAux7(k) = VecWeight(i)
      Let VecAux8(k) = VecContent(i)
      Let PVarUnits(k) = VecUnit(i)
      For j = 0 To UBound(VecContent)
        If VBA.LCase(VecContent(j)) = VBA.LCase(VecAux8(k)) And Not (VBA.CStr(VecContent(j)) = VBA.vbNullString) Then
          ReDim Preserve VecAux0(j)
          ReDim Preserve VecAux1(j)
          ReDim Preserve VecAux2(j)
          Let VecAux0(j) = VecGrade(j)
          Let VecAux1(j) = VecPrice(j)
          Let VecAux2(j) = VecWeight(j)
          Let VecContent(j) = VBA.vbNullString
          Let l = l + 1
        End If
      Next j
      If l > 0 Then
        Let VecAux5(k) = Application.WorksheetFunction.Average(VecAux0)
        Let VecAux6(k) = Application.WorksheetFunction.Average(VecAux1)
        Let VecAux7(k) = Application.WorksheetFunction.Average(VecAux2)
        Let VecAux0 = Empty
        Let VecAux1 = Empty
        Let VecAux2 = Empty
        ReDim VecAux0(0) 'Grades
        ReDim VecAux1(0) 'Prices
        ReDim VecAux2(0) 'Weights
        Let l = 0
      End If
      Let k = k + 1
    End If
  Next i
  Let VecContent = Empty
  Let VecGrade = Empty
  Let VecUnit = Empty
  Let VecPrice = Empty
  Let VecWeight = Empty
  
  ReDim PVarContents(0 To UBound(VecAux5), 0 To 3)
  For i = 0 To UBound(PVarContents)
    Let PVarContents(i, 0) = VecAux5(i) 'Grades
    Let PVarContents(i, 1) = VecAux6(i) 'Prices
    Let PVarContents(i, 2) = VecAux7(i) 'Weights
    Let PVarContents(i, 3) = VecAux8(i) 'Contents
  Next i
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Let VecAux2 = Empty
  Let VecAux3 = Empty
  Let VecAux4 = Empty
  Let VecAux5 = Empty
  Let VecAux6 = Empty
  Let VecAux7 = Empty
  Let VecAux8 = Empty

  'Return
  Let PurchaseMineral_Core_assets_get_average = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Let VecAux2 = Empty
  Let VecAux3 = Empty
  Let VecAux4 = Empty
  Let VecAux5 = Empty
  Let VecAux6 = Empty
  Let VecAux7 = Empty
  Let VecAux8 = Empty
  Let VecContent = Empty
  Let VecGrade = Empty
  Let VecUnit = Empty
  Let VecPrice = Empty
  Let VecWeight = Empty
  Set REGEX = Nothing
  'Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function PurchaseMineral_Core_assets_get_grades( _
  ByVal PBooSimple As Boolean, _
  ByVal PVarWeights As Variant, _
  ByRef PVarDataGross As Variant, _
  ByRef PVarUnits As Variant, _
  ByVal RES As AppResources_current) As Boolean

  'PVarWeights: 1D array containing weights of each sample
  'PVarDataGross: 2D array to return all captured data
  'Dim RES As AppResources
  Dim REGEX As AppResRegEx
  
  Dim VecAux0 As Variant, VecAux1 As Variant, VecAux2 As Variant, VecNumbers As Variant, VecGrades As Variant, VecPrices As Variant, VecAux3 As Variant, VecWeights As Variant, VecTypes As Variant
  Dim StrAux0 As String, StrAux1 As String, StrAux2 As String, StrAux3 As String, StrAux4 As String
  Dim BooComplex As Boolean, BooAux As Boolean
  Dim DblAux As Double
  Dim i As Long, j As Long, k As Long, l As Long, xRows As Long, xRowAux As Long

  Const sFN As String = "APPMIN::-core_assets_get_grades"

  On Error GoTo EH
  Let PurchaseMineral_Core_assets_get_grades = False

  'Get datas
  Let StrAux0 = GLOTYPLO.PMaindata ' "excliqlitepurchasemaindata"
  Let xRows = IIf(PBooSimple, excliqlitedatasheet.ListObjects(StrAux0).ListRows.Count, GLOLNG_ROWCOUNTER)
  
  'get weights
  ReDim VecAux3(0 To UBound(PVarWeights) - 3)
  For i = 0 To UBound(VecAux3)
    Let VecAux3(i) = PVarWeights(i + 2, 4)
  Next i
  ReDim VecNumbers(0)
  ReDim VecGrades(0)
  ReDim PVarUnits(0)
  ReDim VecPrices(0)
  ReDim VecWeights(0)
  ReDim VecTypes(0)
  Let j = 0
  Let l = 0
  Let xRowAux = IIf(PBooSimple, 1, xRows)
  For i = xRowAux To xRows
    Let StrAux1 = excliqlitedatasheet.ListObjects(StrAux0).ListColumns("TIPO").DataBodyRange.Cells(i, 1).Value
    Set REGEX = New AppResRegEx
    Let BooComplex = REGEX.isMineralComplex(StrAux1)
    Set REGEX = Nothing
    If BooComplex Then
      Let StrAux2 = excliqlitedatasheet.ListObjects(StrAux0).ListColumns("LEYES").DataBodyRange.Cells(i, 1).Value
      Let StrAux3 = excliqlitedatasheet.ListObjects(StrAux0).ListColumns("UNIDADES").DataBodyRange.Cells(i, 1).Value
      Let StrAux4 = excliqlitedatasheet.ListObjects(StrAux0).ListColumns("PRECIO").DataBodyRange.Cells(i, 1).Value
      Let VecAux0 = VBA.Split(StrAux2, ";") 'Grades
      Let VecAux1 = VBA.Split(StrAux3, ";") 'Units
      Let VecAux2 = VBA.Split(StrAux4, ";") 'Prices
      ReDim Preserve VecNumbers(j)
      Let StrAux0 = (j + 1) & ". "  '#
      Let VecNumbers(j) = "'" & StrAux0
      For k = 0 To UBound(VecAux0)
        ReDim Preserve VecNumbers(j) '# -1
        ReDim Preserve VecGrades(j)
        ReDim Preserve PVarUnits(j)
        ReDim Preserve VecPrices(j)
        ReDim Preserve VecWeights(j)
        ReDim Preserve VecTypes(j)
        If VBA.IsEmpty(VecNumbers(j)) Then Let VecNumbers(j) = "~" & StrAux0 '#
        Let VecGrades(j) = VBA.CDbl(VecAux0(k)) 'Grades
        Let PVarUnits(j) = VecAux1(k) 'Units
        Let VecPrices(j) = VBA.CDbl(VecAux2(k)) 'Prices
        Let VecWeights(j) = VecAux3(l) 'Weights
        Let VecTypes(j) = True 'Complex
        Let j = j + 1
      Next k
      Let l = l + 1
    Else
      ReDim Preserve VecNumbers(j) '#
      ReDim Preserve VecGrades(j) 'Grades
      ReDim Preserve PVarUnits(j) 'Units
      ReDim Preserve VecPrices(j) 'Prices
      ReDim Preserve VecWeights(j)
      ReDim Preserve VecTypes(j) 'Non complex
      Let VecNumbers(j) = "'" & (j + 1) & "." '#
      Let VecGrades(j) = VBA.CDbl(excliqlitedatasheet.ListObjects(StrAux0).ListColumns("LEYES").DataBodyRange.Cells(i, 1).Value)
      Let PVarUnits(j) = excliqlitedatasheet.ListObjects(StrAux0).ListColumns("UNIDADES").DataBodyRange.Cells(i, 1).Value
      Let VecPrices(j) = VBA.CDbl(excliqlitedatasheet.ListObjects(StrAux0).ListColumns("PRECIO").DataBodyRange.Cells(i, 1).Value)
      Let VecWeights(j) = VecAux3(l) 'Weights
      Let VecTypes(j) = False 'Non Complex
      Let l = l + 1
      Let j = j + 1
    End If
  Next i
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Let VecAux2 = Empty
  Let VecAux3 = Empty
  
  'Fill all captured data in main array
  ReDim PVarDataGross(0 To UBound(VecNumbers), 0 To 4) '#, Grades, Prices and weights
  For i = 0 To UBound(PVarDataGross)
    Let PVarDataGross(i, 0) = VecNumbers(i)
    Let PVarDataGross(i, 1) = VecGrades(i)
    Let PVarDataGross(i, 2) = VecPrices(i)
    If Not GLOBOO_FORMULAS Then Let PVarDataGross(i, 3) = VBA.CDbl(VecGrades(i)) * VBA.CDbl(VecPrices(i)) * VecWeights(i)
    Let PVarDataGross(i, 4) = VecTypes(i)
  Next i
  Let VecAux0 = Empty
  Let VecNumbers = Empty
  Let VecGrades = Empty
  Let VecPrices = Empty
  Let VecWeights = Empty
  Let VecTypes = Empty

  'Return
  Let PurchaseMineral_Core_assets_get_grades = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Let VecAux2 = Empty
  Let VecNumbers = Empty
  Let VecGrades = Empty
  Let VecPrices = Empty
  Let VecWeights = Empty
  Let VecTypes = Empty
  'Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

'Private Sub PurchaseMineral_Core_assets_set_strlos()
'
'  With GLOTYPLO
'    Let .PCurrencies = GLOSTRPREFIX & "currenciescon"
'    Let .PBasedata = GLOSTRPREFIX & "purchasebasedata"
'    Let .PMaindata = GLOSTRPREFIX & "purchasemaindata"
'    Let .POptions = GLOSTRPREFIX & "orepurchaseoptions"
'    Let .PVariousret = GLOSTRPREFIX & "purchasevariousretentions"
'    Let .PVariousretother = GLOSTRPREFIX & "purchasevariousretentionsother"
'    Let .PDiscounts = GLOSTRPREFIX & "purchasediscounts"
'    Let .PBonus = GLOSTRPREFIX & "purchasebonus"
'    Let .PConfigprm = GLOSTRPREFIX & "configpurrm"
'    Let .PConfigpret = GLOSTRPREFIX & "configpurretentions"
'    Let .PConfigpretother = GLOSTRPREFIX & "configpurretentionsother"
'  End With
'
'End Sub

Private Function PurchaseMineral_Core_assets_get_deductions_rm( _
  ByVal PBooSimple As Boolean, _
  ByVal PBooME As Boolean, _
  ByVal PDblTC As Double, _
  ByVal PRngBox As Range, _
  ByRef PLngRows As Long, _
  ByRef PTypResponse As GLOTYPPURCHASE, _
  ByRef PVarDataDeductionRM As Variant, _
  ByVal RES As AppResources_current) As Boolean

  'Dim RES As AppResources
  
  Dim VecAux0 As Variant
  Dim StrAux0 As String
  Dim DblAux As Double, DblRM As Double
  Dim BooResumedRM As Boolean
  Dim i As Long, j As Long, k As Long, xRows As Long, xCols As Long, xRowFirstVal  As Long

  Const sFN As String = "APPMIN::-core_assets_get_deductions_rm"

  On Error GoTo EH
  Let PurchaseMineral_Core_assets_get_deductions_rm = False

  'Initial and base values
  Let DblRM = 0
  Let xCols = 4
  Let xRows = 0
  
  With GLOTYPLO
    Set RES = New AppResources
    Let BooResumedRM = VBA.CBool(GLOSHTDATA.ListObjects(.POptions).ListColumns("orepurchase_rmdetail").DataBodyRange.Cells(1, 1).Value)
    If BooResumedRM Then 'Short
      If PBooSimple Then Let StrAux0 = Application.WorksheetFunction.TextJoin(";", True, GLOSHTDATA.ListObjects(.PMaindata).ListColumns("CONTENIDOS").DataBodyRange)
      If PBooSimple = False Then Let StrAux0 = GLOSHTDATA.ListObjects(.PMaindata).ListColumns("CONTENIDOS").DataBodyRange.Cells(GLOLNG_ROWCOUNTER, 1).Value
      Let VecAux0 = 0 'RES.resArrayUniques(VBA.Split(StrAux0, ";"))
      If VBA.IsArray(VecAux0) Then
        For i = 0 To UBound(VecAux0)
          Let DblRM = DblRM + 0 'VBA.CDbl(RES.resExcelVLockup(VBA.CStr(VecAux0(i)), GLOSHTDATA.ListObjects(.PConfigprm).ListColumns("ELEMENTO").DataBodyRange, GLOSHTDATA.ListObjects(.PConfigprm).ListColumns("VENTA INTERNA").DataBodyRange))
        Next i
        Let VecAux0 = Empty
        Let xRows = 1 'Subtitle + content
      End If
    Else 'Detailed
      If PBooSimple Then Let StrAux0 = Application.WorksheetFunction.TextJoin(";", True, GLOSHTDATA.ListObjects(.PMaindata).ListColumns("CONTENIDOS").DataBodyRange)
      If PBooSimple = False Then Let StrAux0 = GLOSHTDATA.ListObjects(.PMaindata).ListColumns("CONTENIDOS").DataBodyRange.Cells(GLOLNG_ROWCOUNTER, 1).Value
      Let VecAux0 = 0 'RES.resArrayUniques(VBA.Split(StrAux0, ";"))
      If VBA.IsArray(VecAux0) Then Let xRows = UBound(VecAux0) + 1 'Subtitle + content
    End If
  End With
  
  
  With PTypResponse
    'Determine matrix dimenssions
    If xRows <= 0 Then GoTo EH 'If exists rm's
    Let xRowFirstVal = .PLngRowCounter + 3 ', xRowFirstVal  As Long, xRowResize As Long, xRowTotalsAs Long
    ReDim PVarDataDeductionRM(0 To xRows, 0 To xCols)
    
    'Fill matrix
    For j = 0 To UBound(PVarDataDeductionRM, 2)
      For i = 0 To UBound(PVarDataDeductionRM)
        If i = 0 Then
          If j = 0 Then Let PVarDataDeductionRM(i, j) = "RETENCIONES DE LEY - REGALÍAS [%]:"
        Else
          Let k = i - 1
          Select Case j
            Case 0 'Concept
              If BooResumedRM Then Let PVarDataDeductionRM(i, j) = "RM"
              If BooResumedRM = False Then Let PVarDataDeductionRM(i, j) = "RM-" & VBA.CStr(VecAux0(k))
            Case 1 '%
              If BooResumedRM Then Let PVarDataDeductionRM(i, j) = DblRM
              If BooResumedRM = False Then Let PVarDataDeductionRM(i, j) = 0 'RES.resExcelVLockup(VBA.CStr(VecAux0(k)), GLOSHTDATA.ListObjects(GLOTYPLO.PConfigprm).ListColumns("ELEMENTO").DataBodyRange, GLOSHTDATA.ListObjects(GLOTYPLO.PConfigprm).ListColumns("VENTA INTERNA").DataBodyRange)
            Case 3 'Value Foreign
              If PBooME Then
                If GLOBOO_FORMULAS Then Let PVarDataDeductionRM(i, j) = "=" & PRngBox.Range("I" & xRowFirstVal + k).Address(False, False) & "/" & PRngBox.Range("B10").Address(False, False)
                If GLOBOO_FORMULAS = False Then Let PVarDataDeductionRM(i, j) = (.PDblTotalGrossVal / PDblTC) * (VBA.CDbl(PVarDataDeductionRM(i, 1)) / 100)
              End If
            Case 4 'Value Local
              If GLOBOO_FORMULAS Then
                Let PVarDataDeductionRM(i, j) = "=" & PRngBox.Range(.PStrTotalGrossRng).Address(False, False) & "*" & PRngBox.Range("F" & xRowFirstVal + k).Address(False, False) & "/100"
              Else
                Let DblAux = .PDblTotalGrossVal * (VBA.CDbl(PVarDataDeductionRM(i, 1)) / 100)
                Let PVarDataDeductionRM(i, j) = DblAux
                Let .PDblTotalDedVal = .PDblTotalDedVal + DblAux
              End If
          End Select
        End If
      Next i
    Next j
    Let .PLngRowCounter = xRowFirstVal + xRows - 1
  End With
  Let VecAux0 = Empty
  Set RES = Nothing

  'Return
  Let PLngRows = PLngRows + xRows + 1
  Let PurchaseMineral_Core_assets_get_deductions_rm = True

EH:
  Let VecAux0 = Empty
  'Set RES = Nothing
  Debug.Print VBA.Err.Description
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function PurchaseMineral_Core_assets_get_deductions_retentions( _
  ByVal PBooSimple As Boolean, _
  ByVal PBooME As Boolean, _
  ByVal PDblTC As Double, _
  ByVal PRngBox As Range, _
  ByRef PTypResponse As GLOTYPPURCHASE, _
  ByRef PLngRows As Long, _
  ByRef PVarDataDeductionRetentions As Variant, _
  ByVal RES As AppResources_current) As Boolean

  'Dim RES As AppResources
  
  Dim VecAux0 As Variant
  Dim StrAux0 As String
  Dim BooAllRet As Boolean
  Dim i As Long, j As Long, k As Long, xRows As Long, xCols As Long, xRowFirstVal  As Long

  Const sFN As String = "APPMIN::-core_assets_get_deductions_retentions"

  On Error GoTo EH
  Let PurchaseMineral_Core_assets_get_deductions_retentions = False

  'Initial and base values
  Let xCols = 4
  Let xRows = 0
  With GLOTYPLO
    Set RES = New AppResources
    If PBooSimple Then 'Just one provider
      Let xRows = Application.WorksheetFunction.CountIf(GLOSHTDATA.ListObjects(.PConfigpret).ListColumns("SELECTED").DataBodyRange, "=1")
      If xRows > 0 Then
        ReDim VecAux0(0)
        For i = 1 To GLOSHTDATA.ListObjects(.PConfigpret).ListRows.Count
          If VBA.CBool(GLOSHTDATA.ListObjects(.PConfigpret).ListColumns("SELECTED").DataBodyRange.Cells(i, 1).Value) Then
            ReDim Preserve VecAux0(i - 1)
            Let VecAux0(i - 1) = GLOSHTDATA.ListObjects(.PConfigpret).ListColumns("CONCEPTO").DataBodyRange.Cells(i, 1).Value
          End If
        Next i
      End If
    Else 'Multiple
      Let BooAllRet = (Not GLOSHTDATA.ListObjects(.PVariousret).DataBodyRange Is Nothing)
      If BooAllRet Then
        Let BooAllRet = VBA.CBool(GLOSHTDATA.ListObjects(.PVariousret).ListColumns("all").DataBodyRange.Cells(GLOLNG_ROWCOUNTER, 1).Value)
        If BooAllRet Then
          Let xRows = GLOSHTDATA.ListObjects(.PConfigpret).ListRows.Count 'Subtitle + content
          ReDim VecAux0(0)
          For i = 1 To GLOSHTDATA.ListObjects(.PConfigpret).ListRows.Count
            ReDim Preserve VecAux0(i - 1)
            Let VecAux0(i - 1) = GLOSHTDATA.ListObjects(.PConfigpret).ListColumns("CONCEPTO").DataBodyRange.Cells(i, 1).Value
          Next i
        Else
          Let StrAux0 = GLOSHTDATA.ListObjects(.PVariousret).ListColumns("values").DataBodyRange.Cells(GLOLNG_ROWCOUNTER, 1).Value
          If VBA.Len(StrAux0) > 0 Then
            Let VecAux0 = VBA.Split(StrAux0, ";")
            Let xRows = UBound(VecAux0) + 1 'Subtitle + content
          End If
        End If
      End If
    End If
  End With
  
  With PTypResponse
    'Determine matrix dimenssions
    If xRows <= 0 Then GoTo EH 'If doesn't exists Retentions
    Let xRowFirstVal = .PLngRowCounter + 2
    ReDim PVarDataDeductionRetentions(0 To xRows, 0 To xCols)
    
    'Fill matrix
    For j = 0 To UBound(PVarDataDeductionRetentions, 2)
      For i = 0 To UBound(PVarDataDeductionRetentions)
        If i = 0 Then
          If j = 0 Then Let PVarDataDeductionRetentions(i, j) = "RETENCIONES DE LEY - IMPUESTOS Y OTROS [%]:"
        Else
          Let k = i - 1
          Select Case j
            Case 0 'Concept
              Let PVarDataDeductionRetentions(i, j) = VBA.CStr(VecAux0(k))
            Case 1 '%
              Let PVarDataDeductionRetentions(i, j) = 0 'VBA.CDbl(RES.resExcelVLockup(VBA.CStr(VecAux0(k)), GLOSHTDATA.ListObjects(GLOTYPLO.PConfigpret).ListColumns("CONCEPTO").DataBodyRange, GLOSHTDATA.ListObjects(GLOTYPLO.PConfigpret).ListColumns("PORCENTAJE").DataBodyRange))
            Case 3 'Value Foreign
              If PBooME Then
                If GLOBOO_FORMULAS Then
                  Let PVarDataDeductionRetentions(i, j) = "=" & PRngBox.Range("I" & xRowFirstVal + k).Address(False, False) & "/" & PRngBox.Range("B10").Address(False, False)
                Else
                  Let PVarDataDeductionRetentions(i, j) = (.PDblTotalGrossVal / PDblTC) * (VBA.CDbl(PVarDataDeductionRetentions(i, 1)) / 100)
                End If
              End If
            Case 4 'Value Local
              If GLOBOO_FORMULAS Then
                Let PVarDataDeductionRetentions(i, j) = "=" & PRngBox.Range("F" & xRowFirstVal + k).Address(False, False) & "*" & PRngBox.Range(.PStrTotalGrossRng).Address(False, False) & "/100"
              Else
                Let PVarDataDeductionRetentions(i, j) = .PDblTotalGrossVal * (VBA.CDbl(PVarDataDeductionRetentions(i, 1)) / 100)
                Let .PDblTotalDedVal = .PDblTotalDedVal + VBA.CDbl(PVarDataDeductionRetentions(i, j))
              End If
          End Select
        End If
      Next i
    Next j
    Let .PLngRowCounter = xRowFirstVal + xRows - 1
  End With
  Set RES = Nothing
  Let VecAux0 = Empty
        
  'Return
  Let PLngRows = PLngRows + xRows + 1
  Let PurchaseMineral_Core_assets_get_deductions_retentions = True

EH:
  Let VecAux0 = Empty
  'Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function PurchaseMineral_Core_assets_get_deductions_retentionsothers( _
  ByVal PBooSimple As Boolean, _
  ByVal PBooME As Boolean, _
  ByVal PDblTC As Double, _
  ByVal PRngBox As Range, _
  ByRef PTypResponse As GLOTYPPURCHASE, _
  ByRef PLngRows As Long, _
  ByRef PVarDataDeductionRetentionsOthers As Variant, _
  ByVal RES As AppResources_current) As Boolean

  'Dim RES As AppResources
  
  Dim VecAux0 As Variant
  Dim StrAux0 As String
  Dim BooAllRetOther As Boolean
  Dim i As Long, j As Long, k As Long, xRows As Long, xCols As Long, xRowFirstVal  As Long

  Const sFN As String = "APPMIN::-core_assets_get_deductions_retentionsothers"

  On Error GoTo EH
  Let PurchaseMineral_Core_assets_get_deductions_retentionsothers = False

  'Initial and base values
  Let xCols = 4
  Let xRows = 0
  
  With GLOTYPLO
    If PBooSimple Then 'Just one provider
      Let xRows = Application.WorksheetFunction.CountIf(GLOSHTDATA.ListObjects(.PConfigpretother).ListColumns("SELECTED").DataBodyRange, "=1")
      If xRows > 0 Then
        ReDim VecAux0(0)
        For i = 1 To GLOSHTDATA.ListObjects(.PConfigpretother).ListRows.Count
          If VBA.CBool(GLOSHTDATA.ListObjects(.PConfigpretother).ListColumns("SELECTED").DataBodyRange.Cells(i, 1).Value) Then
            ReDim Preserve VecAux0(i - 1)
            Let VecAux0(i - 1) = GLOSHTDATA.ListObjects(.PConfigpretother).ListColumns("CONCEPTO").DataBodyRange.Cells(i, 1).Value
          End If
        Next i
      End If
    Else 'Multiple
      Let BooAllRetOther = (Not GLOSHTDATA.ListObjects(.PVariousretother).DataBodyRange Is Nothing)
      If BooAllRetOther Then
        Let BooAllRetOther = VBA.CBool(GLOSHTDATA.ListObjects(.PVariousretother).ListColumns("all").DataBodyRange.Cells(GLOLNG_ROWCOUNTER, 1).Value)
        If BooAllRetOther Then
          Let xRows = GLOSHTDATA.ListObjects(.PConfigpretother).ListRows.Count 'Subtitle + content
          ReDim VecAux0(0)
          For i = 1 To GLOSHTDATA.ListObjects(.PConfigpretother).ListRows.Count
            ReDim Preserve VecAux0(i - 1)
            Let VecAux0(i - 1) = GLOSHTDATA.ListObjects(.PConfigpretother).ListColumns("CONCEPTO").DataBodyRange.Cells(i, 1).Value
          Next i
        Else
          Let StrAux0 = GLOSHTDATA.ListObjects(.PVariousretother).ListColumns("values").DataBodyRange.Cells(GLOLNG_ROWCOUNTER, 1).Value
          If VBA.Len(StrAux0) > 0 Then
            Let VecAux0 = VBA.Split(StrAux0, ";")
            Let xRows = UBound(VecAux0) + 1 'Subtitle + content
          End If
        End If
      End If
    End If
  End With
    
  'Determine matrix dimenssions
  With PTypResponse
    If xRows <= 0 Then GoTo EH 'If doesn't exists retentions others
    Let xRowFirstVal = .PLngRowCounter + 2
    ReDim PVarDataDeductionRetentionsOthers(0 To xRows, 0 To xCols)
    
    'Fill matrix
    Set RES = New AppResources
    For j = 0 To UBound(PVarDataDeductionRetentionsOthers, 2)
      For i = 0 To UBound(PVarDataDeductionRetentionsOthers)
        If i = 0 Then
          If j = 0 Then Let PVarDataDeductionRetentionsOthers(i, j) = "RETENCIONES OTRAS [%]:"
        Else
          Let k = i - 1
          Select Case j
            Case 0 'Concept
              Let PVarDataDeductionRetentionsOthers(i, j) = VBA.CStr(VecAux0(k))
            Case 1 '%
              Let PVarDataDeductionRetentionsOthers(i, j) = 0 'VBA.CDbl(RES.resExcelVLockup(VBA.CStr(VecAux0(k)), GLOSHTDATA.ListObjects(GLOTYPLO.PConfigpretother).ListColumns("CONCEPTO").DataBodyRange, GLOSHTDATA.ListObjects(GLOTYPLO.PConfigpretother).ListColumns("PORCENTAJE").DataBodyRange))
            Case 3 'Value Foreign
              If PBooME Then
                If GLOBOO_FORMULAS Then
                  Let PVarDataDeductionRetentionsOthers(i, j) = "=" & PRngBox.Range("I" & xRowFirstVal + k).Address(False, False) & "/" & PRngBox.Range("B10").Address(False, False)
                Else
                  Let PVarDataDeductionRetentionsOthers(i, j) = (.PDblTotalGrossVal / PDblTC) * (VBA.CDbl(PVarDataDeductionRetentionsOthers(i, 1)) / 100)
                End If
              End If
            Case 4 'Value Local
              If GLOBOO_FORMULAS Then
                Let PVarDataDeductionRetentionsOthers(i, j) = "=" & PRngBox.Range("F" & xRowFirstVal + k).Address(False, False) & "*" & PRngBox.Range(.PStrTotalGrossRng).Address(False, False) & "/100"
              Else
                Let PVarDataDeductionRetentionsOthers(i, j) = .PDblTotalGrossVal * (VBA.CDbl(PVarDataDeductionRetentionsOthers(i, 1)) / 100)
                Let .PDblTotalDedVal = .PDblTotalDedVal + VBA.CDbl(PVarDataDeductionRetentionsOthers(i, j))
              End If
          End Select
        End If
      Next i
    Next j
    Let .PLngRowCounter = xRowFirstVal + xRows - 1
  End With
  Let VecAux0 = Empty
  Set RES = Nothing
  
  'Return
  Let PLngRows = PLngRows + xRows + 1
  Let PurchaseMineral_Core_assets_get_deductions_retentionsothers = True

EH:
  Let VecAux0 = Empty
  'Set RES = Nothing
  Debug.Print VBA.Err.Description
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function PurchaseMineral_Core_assets_get_deductions_discounts( _
  ByVal PBooSimple As Boolean, _
  ByVal PBooME As Boolean, _
  ByVal PDblTC As Double, _
  ByVal PRngBox As Range, _
  ByRef PTypResponse As GLOTYPPURCHASE, _
  ByRef PLngRows As Long, _
  ByRef PVarDataDeductionsDiscounts As Variant, _
  ByVal RES As AppResources_current) As Boolean

  'Dim RES As AppResources
  
  Dim VecAux0 As Variant, VecAux1 As Variant, VecAux2 As Variant, VecAux3 As Variant, xRowFirstVal  As Long
  Dim StrAux0 As String
  Dim BooDiscounts As Boolean
  Dim i As Long, j As Long, k As Long, xRows As Long, xCols As Long

  Const sFN As String = "APPMIN::-core_assets_get_deductions_discounts"

  On Error GoTo EH
  Let PurchaseMineral_Core_assets_get_deductions_discounts = False

  'Initial and base values
  Let xCols = 4
  Let xRows = 0
  
  Let BooDiscounts = (Not GLOSHTDATA.ListObjects(GLOTYPLO.PDiscounts).DataBodyRange Is Nothing)
  If BooDiscounts Then
    With GLOTYPLO
      If PBooSimple Then
        Set RES = New AppResources
        Let StrAux0 = Application.WorksheetFunction.TextJoin(";", True, GLOSHTDATA.ListObjects(.PDiscounts).ListColumns("VALOR").DataBodyRange)
        Let VecAux2 = 0 'RES.resArrayGetDataArrayFromVariants(StrAux0, xRows, False, True, True) 'Value = #
        Let StrAux0 = Application.WorksheetFunction.TextJoin(";", True, GLOSHTDATA.ListObjects(.PDiscounts).ListColumns("DESCUENTO").DataBodyRange)
        Let VecAux3 = 0 'RES.resArrayGetDataArrayFromVariants(StrAux0, xRows, True)
        Set RES = Nothing
      Else 'Múltiple
        Let StrAux0 = GLOSHTDATA.ListObjects(.PDiscounts).ListColumns("VALOR").DataBodyRange.Cells(GLOLNG_ROWCOUNTER, 1).Value
        If VBA.Len(StrAux0) > 0 Then
          Let VecAux2 = VBA.Split(StrAux0, ";")
          Let xRows = UBound(VecAux2) + 1 'Subtitle + content
          Let StrAux0 = GLOSHTDATA.ListObjects(.PDiscounts).ListColumns("DESCUENTO").DataBodyRange.Cells(GLOLNG_ROWCOUNTER, 1).Value
          Let VecAux3 = VBA.Split(StrAux0, ";")
        End If
      End If
    End With
    
    If xRows > 0 Then
      Let j = 0
      ReDim VecAux0(0)
      ReDim VecAux1(0)
      For i = 0 To UBound(VecAux2)
        If Not VBA.CDbl(VecAux2(i)) = 0 Then
          ReDim Preserve VecAux0(j)
          ReDim Preserve VecAux1(j)
          Let VecAux0(j) = VecAux3(i) 'Concept
          Let VecAux1(j) = VecAux2(i) 'Value
          Let j = j + 1
        End If
      Next i
      Let xRows = j 'UBound(VecAux0) + 1 'Subtitle + content
      Let VecAux2 = Empty
      Let VecAux3 = Empty
    End If
  End If
  
  'Determine matrix dimenssions
  With PTypResponse
    If xRows <= 0 Then GoTo EH 'If doesn't exists retentions others
    Let xRowFirstVal = .PLngRowCounter + 2
    ReDim PVarDataDeductionsDiscounts(0 To xRows, 0 To xCols)
    
    'Fill matrix
    For j = 0 To UBound(PVarDataDeductionsDiscounts, 2)
      For i = 0 To UBound(PVarDataDeductionsDiscounts)
        If i = 0 Then
          If j = 0 Then Let PVarDataDeductionsDiscounts(i, j) = "DESCUENTOS:"
        Else
          Let k = i - 1
          Select Case j
            Case 0 'Concept
              Let PVarDataDeductionsDiscounts(i, j) = VBA.CStr(VecAux0(k))
            Case 3 'Value Foreign
              If PBooME Then
                If GLOBOO_FORMULAS Then
                  Let PVarDataDeductionsDiscounts(i, j) = "=" & PRngBox.Range("I" & xRowFirstVal + k).Address(False, False) & "/" & PRngBox.Range("B10").Address(False, False)
                Else
                  Let PVarDataDeductionsDiscounts(i, j) = (VBA.CDbl(VecAux1(k)) / PDblTC)
                End If
              End If
            Case 4 'Value Local
              Let PVarDataDeductionsDiscounts(i, j) = VecAux1(k)
              Let .PDblTotalDedVal = .PDblTotalDedVal + VBA.CDbl(PVarDataDeductionsDiscounts(i, j))
          End Select
        End If
      Next i
    Next j
    Let .PLngRowCounter = xRowFirstVal + xRows - 1
  End With
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Set RES = Nothing
  
  'Return
  Let PLngRows = PLngRows + xRows + 1
  Let PurchaseMineral_Core_assets_get_deductions_discounts = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Let VecAux2 = Empty
  Let VecAux3 = Empty
  'Set RES = Nothing
  Debug.Print VBA.Err.Description
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function




' ERRORS MANNAGEMENT AND ASSERTSIONS ------------------------------------------------'
' Debug_asset                                                    All Assets in module'
Private Sub Debug_assert(ByVal Cond As Boolean, Optional Origin As String = VBA.vbNullString)
  
  #If Debbuging Then
    Debug.Assert Cond
    If Not Cond Then Debug.Print "Assert on: " & Origin
  #End If

End Sub




' METHODS PUBLIC --------------------------------------------------------------------'
' INTERFACE METHODS LIST ------------------------------------------------------------'
' ABSTRACT INTERFACE UI METHODS LIST AND UDF METHODS LIST (PUBLIC) ------------------'
Public Function ICoreController_GetSolution(ByVal CStrRequester As String, ByVal CVarRequest As Variant, CVarResponse As Variant) As Boolean
  
  On Error GoTo EH
  Let ICoreController_GetSolution = PurchaseMineral_Core_Main(CVarRequest, CVarResponse)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("APPREG::IGetSolution")

End Function

Private Function ICoreController_GetSolutionUDF(ByVal CStrRequester As String, ByVal CVarRequest As Variant) As Variant
End Function
' METHODS PUBLIC --------------------------------------------------------------------'
' ============================ CLASS MODULES STRUCTURE =============================='
