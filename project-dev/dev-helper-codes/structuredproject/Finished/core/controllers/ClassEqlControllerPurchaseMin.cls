VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassEqlControllerPurchaseMin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============ CLASS EXCLIQ LITE CONTROLLER PURCHASE MINERAL STRUCTURE ============= '
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



' ============ CLASS EXCLIQ LITE CONTROLLER PURCHASE MINERAL STRUCTURE ============= '
' OBJECT VARIABLES (GLOOBJ_) ------------------------------------------------------- '
' VARIABLES (GLOStr_) -------------------------------------------------------------- '
' CONSTANTS ERROR MESSAGES (GLOSTR_ERR_) ------------------------------------------- '
' PROPERTIES ----------------------------------------------------------------------- '
' CONSTRUCTOR/DESTRUCTOR ----------------------------------------------------------- '
' CORE METHODS PRIVATE ------------------------------------------------------------- '
' CORE METHODS LIST (PRIVATE) ------------------------------------------------------ '
' Purchasemin_core_main
' Purchasemin_core_generate_data_header
' Purchasemin_core_generate_data_content
' Purchasemin_core_generate_data_content_getoredata
' Purchasemin_core_generate_data_content_getorepay
' Purchasemin_core_generate_data_content_getgrossvalue
' Purchasemin_core_generate_data_content_getdeductions
' Purchasemin_core_generate_data_content_getbonus
' Purchasemin_core_generate_data_content_getfinalpurchase
' ASSETS --------------------------------------------------------------------------- '
' PurchaseMineral_Core_assets_get_average
' PurchaseMineral_Core_assets_get_grades
' PurchaseMineral_Core_assets_set_strlos
' PurchaseMineral_Core_assets_get_deductions_rm
' PurchaseMineral_Core_assets_get_deductions_retentions
' PurchaseMineral_Core_assets_get_deductions_retentionsothers
' PurchaseMineral_Core_assets_get_deductions_discounts
' ERRORS MANNAGEMENT AND ASSERTSIONS ----------------------------------------------- '
' Debug_assert
' METHODS PUBLIC ------------------------------------------------------------------- '
' INTERFACE METHODS LIST ----------------------------------------------------------- '
' PurchaseMineralGet
' ============ CLASS EXCLIQ LITE CONTROLLER PURCHASE MINERAL STRUCTURE ============= '




' ============ CLASS EXCLIQ LITE CONTROLLER PURCHASE MINERAL STRUCTURE ============= '
' ABSTRACT IMPLEMENTATIONS --------------------------------------------------------- '
Implements ClassEqlControllerInterface

' OBJECT VARIABLES (GLOOBJ_) ------------------------------------------------------- '
Private RES As ClassEqlResources
Private MDL As ClassEqlModel
Private GLORNG_RNGBOX As Range

' VARIABLES (GLOStr_) -------------------------------------------------------------- '
Private GLOBOO_FORMULAS As Boolean
Private GLOBOO_AVERAGEANYWAY As Boolean
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

Private GLOSTR_PLACE As String
Private GLOBOO_ME As Boolean
Private GLOBOO_VARIOUS As Boolean
Private GLOBOO_AVERAGE  As Boolean

' VARIABLES CUSTOM (GLOEnu|Typ_) --------------------------------------------------- '
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

' CONSTANTS ERROR MESSAGES (GLOSTR_ERR_) ------------------------------------------- '
Private Const GLOSTR_ERR_1 As String = "Faltan datos de orígen."
Private Const GLOSTR_ERR_2 As String = "Al menos un dato debe ser diferente de cero."
Private Const GLOSTR_ERR_3 As String = "Los datos independientes (X) deben tener la misma cantidad que los de orígen (Y)."

Private Const GLOSTR_ERR_9 As String = "¡Error, división entre cero!"






' PROPERTIES ----------------------------------------------------------------------- '
' ABSTRACT INTERFACE PROPERTIES IMPLEMENTATION ------------------------------------- '
Public Property Set ClassEqlControllerInterface_ResourcesApp(ByRef RHS As ClassEqlResources)

  Set RES = RHS

  Let GLOBOO_FORMULAS = RES.Formulas

End Property




' CONSTRUCTOR/DESTRUCTOR ------------------------------------------------------------ '
Private Sub Class_Terminate()

  Set MDL = Nothing
  Set RES = Nothing

End Sub





' CORE METHODS PRIVATE ------------------------------------------------------------- '
' CORE METHODS LIST (PRIVATE) ------------------------------------------------------ '
Private Function Purchasemin_core_main(ByRef PVarRequest As Variant, ByRef PVarResponse As Variant) As Boolean

  Dim MIN As ClassEqlControllerHelper
  Dim RngBox As Range
  Dim VecResponse As Variant, VecAux0 As Variant, VecAux1 As Variant
  Dim BooTwo As Boolean
  Dim i As Long, k As Long

  Const sFN As String = "MIN::-core_main"

  On Error GoTo EH
  Let Purchasemin_core_main = False

  ' Request: 0.StrPlace, 1.BooME, 2.BooVarious, 3.BooAverage, 4.BooPage, 5.BooPrintTwo, 6.BooNewSheet, 7.RngBox
  Set MDL = New ClassEqlModel
  With MDL
    Set .ResourcesApp = RES ' <---
    Let GLOBOO_AVERAGEANYWAY = .PurAverage ' Get this value from db
    Let GLOBOO_RESUMED_RM = .PurRMDetailed
    Let GLOBOO_RESUMED_RETENTIONS = .PurRetentions
    Let GLOBOO_RESUMED_RETENTIONSOTHERS = .PurRetentionsOther
    Let GLOBOO_RESUMED_DISCOUNTS = .PurDiscounts
    Let GLOBOO_RESUMED_BONUS = .PurBonus
  End With

  ' Getting global ore invoice basic data
  Let GLOSTR_PLACE = VBA.CStr(PVarRequest(0))
  Let GLOBOO_ME = VBA.CBool(PVarRequest(1))
  Let GLOBOO_VARIOUS = VBA.CBool(PVarRequest(2))
  Let GLOBOO_AVERAGE = VBA.CBool(PVarRequest(3))
  Set GLORNG_RNGBOX = PVarRequest(7)
  Let BooTwo = VBA.CBool(PVarRequest(5))

  ' Getting the number of ore invoices requested
  Let k = MDL.LengthRows(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur) - 1
  Let k = IIf(GLOBOO_VARIOUS, k, 0)

  Let VecAux0 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigcurrencies_sys, eqlMdlArray, "currency, currencies, symbol", MStrWhere:="currency_main=1"), eqlRes1D)
  Let GLOSTR_CURRENCY = VBA.CStr(VecAux0(0))
  Let GLOSTR_CURRENCIES = VBA.CStr(VecAux0(1))
  Let GLOSTR_CURLOCAL = VBA.CStr(VecAux0(2))
  Let GLOSTR_CURFOREIGN = MDL.CurrencyMainSymbolForeign ' VBA.CStr(MDL.MGet(eqlMdlTblConfigcurrencies, eqlMdlArray, "symbol", MStrWhere:="currency_foreign=1")(0, 0))
  Let VecAux0 = Empty

  ' Requesting final invoices
  Let VecAux1 = MDL.MGet(eqlMdlSheetPur, eqlMdlTblPurchasemin_pur, eqlMdlArray)
  ReDim VecResponse(k)
  For GLOLNG_ROWCOUNTER = 0 To k

    If GLOBOO_VARIOUS Then
      Let VecAux0 = Empty
      Let VecAux0 = RES.ArrayShift(VecAux1)
    Else
      Let VecAux0 = VecAux1
      Let VecAux1 = Empty
    End If

    ' Update RngBox coordinates
    If GLOBOO_FORMULAS Then
      If GLOBOO_VARIOUS Then
        Let i = (9 * IIf(BooTwo, 2, 1) + IIf(BooTwo, 2, 0)) * GLOLNG_ROWCOUNTER
        Set GLORNG_RNGBOX = GLORNG_RNGBOX.Range("A1").Offset(0, i)
      End If
    End If

    ' Get Mineral purchase
    Set MIN = New ClassEqlControllerHelper
    If Not Purchasemin_core_generate(VecAux0, MIN) Then GoTo EH
    Set VecResponse(GLOLNG_ROWCOUNTER) = MIN
    Set MIN = Nothing

  Next GLOLNG_ROWCOUNTER

  ' Return
  Let PVarResponse = VecResponse
  Let Purchasemin_core_main = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Let VecResponse = Empty
  Set RngBox = Nothing
  Set MIN = Nothing
  Set MDL = Nothing
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Purchasemin_core_generate(ByVal PVarRequest As Variant, ByRef PClsRequest As ClassEqlControllerHelper) As Boolean

  Dim TypOinv As GLOTYPPURCHASE

  Const sFN As String = "MIN::-core_generate"

  On Error GoTo EH
  Let Purchasemin_core_generate = False

  ' Resolution
  With TypOinv
    ' Title, purchase and provider datas
    If Not Purchasemin_core_generate_data_header(PVarRequest, .PVecTitle, .PVecInvoice, .PVecProvider) Then GoTo EH

    ' Ore data to ore weight
    If Not Purchasemin_core_generate_data_content(PVarRequest, TypOinv) Then GoTo EH

    ' Return
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

  Let Purchasemin_core_generate = True

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Purchasemin_core_generate_data_header( _
  ByRef PVarRequest As Variant, _
  ByRef VecDataTitle As Variant, _
  ByRef VecDataInvoice As Variant, _
  ByRef VecDataProvider As Variant) As Boolean

  Dim VecAux0 As Variant
  Dim StrAux0 As String
  Dim BooAux0 As Boolean

  Const STR_SIM As String = "Simulación"
  Const STR_DASH As String = "-"
  Const sFN As String = "MIN::-core_generate_data_header"

  On Error GoTo EH
  Let Purchasemin_core_generate_data_header = False

  ' Title
  ReDim VecDataTitle(0 To 1, 0 To 0)
  Let VecDataTitle(0, 0) = "LIQUIDACIÓN DE MINERALES"
  Let VecDataTitle(1, 0) = "(EXPRESADO EN " & VBA.UCase(GLOSTR_CURRENCIES) & ")"

  ' Purchase data
  Let StrAux0 = MDL.PurDestination
  ReDim VecDataInvoice(0 To 6, 0 To 2)
  Let VecDataInvoice(0, 0) = "DATOS LIQUIDACIÓN"
  Let VecDataInvoice(1, 0) = "Destino:": Let VecDataInvoice(1, 1) = IIf(StrAux0 = VBA.vbNullString, STR_SIM, StrAux0)
  Let VecDataInvoice(2, 0) = "Lugar:": Let VecDataInvoice(2, 1) = GLOSTR_PLACE
  Let VecDataInvoice(3, 0) = "Fecha:": Let VecDataInvoice(3, 1) = VBA.CDate(VBA.Format$(VBA.CStr(PVarRequest(0, 17)), "dd/mm/yy"))
  Let VecDataInvoice(4, 0) = "Liquidación:": Let StrAux0 = VBA.CStr(PVarRequest(0, 16)): Let VecDataInvoice(4, 1) = IIf(StrAux0 = VBA.vbNullString, STR_SIM, StrAux0)
  Let VecDataInvoice(5, 0) = "Lote:": Let StrAux0 = VBA.CStr(PVarRequest(0, 15)): Let VecDataInvoice(5, 1) = IIf(StrAux0 = VBA.vbNullString, STR_SIM, StrAux0)
  Let VecDataInvoice(6, 0) = "T/C:": Let VecDataInvoice(6, 1) = PVarRequest(0, 18)
  Let VecDataInvoice(6, 2) = GLOSTR_CURFOREIGN & "/" & GLOSTR_CURLOCAL

  ' Provider data
  ReDim VecDataProvider(0 To 7, 0 To 2)
  Let VecDataProvider(0, 0) = "DATOS PROVEDOR"
  Let VecDataProvider(1, 0) = "Código:": Let StrAux0 = VBA.CStr(PVarRequest(0, 19)): Let VecDataProvider(1, 1) = IIf(StrAux0 = VBA.vbNullString, STR_DASH, StrAux0)
  Let VecDataProvider(2, 0) = "Nombre:": Let VecDataProvider(2, 1) = VBA.CStr(PVarRequest(0, 20))
  Let VecDataProvider(3, 0) = "C.I.:": Let StrAux0 = VBA.CStr(PVarRequest(0, 22)): Let VecDataProvider(3, 1) = IIf(StrAux0 = VBA.vbNullString, STR_DASH, StrAux0)
  Let VecDataProvider(4, 0) = "Tipo:": Let VecDataProvider(4, 1) = VBA.CStr(PVarRequest(0, 21)): Let BooAux0 = Not (VBA.LCase(VBA.CStr(PVarRequest(0, 21))) = "particular")
  Let VecDataProvider(5, 0) = "NIT:": Let StrAux0 = VBA.CStr(PVarRequest(0, 23)): Let VecDataProvider(5, 1) = IIf(StrAux0 = VBA.vbNullString, STR_DASH, StrAux0)
  Let VecDataProvider(6, 0) = "NIM:": Let StrAux0 = VBA.CStr(PVarRequest(0, 24)): Let VecDataProvider(6, 1) = IIf(StrAux0 = VBA.vbNullString, STR_DASH, StrAux0)
  Let VecDataProvider(7, 0) = "Mina:": Let StrAux0 = VBA.CStr(PVarRequest(0, 25)): Let VecDataProvider(7, 1) = IIf(StrAux0 = VBA.vbNullString, STR_DASH, StrAux0)
  If BooAux0 Then
    Let StrAux0 = VBA.CStr(PVarRequest(0, 26)): Let StrAux0 = IIf(StrAux0 = VBA.vbNullString, STR_DASH, StrAux0)
    Let VecDataProvider = RES.ArrayConcat(VecDataProvider, Array("Cooperativa:", StrAux0))
  End If
  ReDim VecAux0(0 To 2, 0 To 2)
  Let VecAux0(0, 0) = "Transportista:": Let StrAux0 = VBA.CStr(PVarRequest(0, 27)): Let VecAux0(0, 1) = IIf(StrAux0 = VBA.vbNullString, STR_DASH, StrAux0)
  Let VecAux0(1, 0) = "Placa Mov.:": Let StrAux0 = VBA.CStr(PVarRequest(0, 28)): Let VecAux0(1, 1) = IIf(StrAux0 = VBA.vbNullString, STR_DASH, StrAux0)
  Let VecAux0(2, 0) = "Cód. Peso:": Let StrAux0 = VBA.CStr(PVarRequest(0, 29)): Let VecAux0(2, 1) = IIf(StrAux0 = VBA.vbNullString, STR_DASH, StrAux0)
  Let VecDataProvider = RES.ArrayConcat(VecDataProvider, VecAux0)

  ' Return
  Let Purchasemin_core_generate_data_header = True

EH:
  Let VecAux0 = Empty
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Purchasemin_core_generate_data_content( _
  ByRef PVarRequest As Variant, _
  ByRef PTypResponse As GLOTYPPURCHASE) As Boolean

  Dim RngBox As Range
  Dim VecAux0 As Variant
  Dim BooSimple As Boolean, BooAverag As Boolean, BooME As Boolean
  Dim DblAux0 As Double, DblAux1 As Double, DblTC As Double
  Dim i As Long, j As Long, k As Long, l As Long, m As Long, n As Long, xRows As Long, xCols As Long, xRowCounter As Long

  Dim StrTotalOreRng As String, StrTotalOrePayRng As String, StrTotalGrossValRng As String, StrTotalDeductionsRng As String, StrTotalBonusRng As String
  Dim DblTotalOre As Double, DblTotalOrePay As Double, DblTotalGrossVal As Double, DblTotalDeductions As Double, DblTotalBonus As Double

  Const STR_SIM As String = "Simulación"
  Const STR_DASH As String = "-"
  Const sFN As String = "MIN::-core_generate_data_content"

  On Error GoTo EH
  Let Purchasemin_core_generate_data_content = False

  Let xCols = 4
  With RES.ErrorHandler

    ' Ore Data
    Let PTypResponse.PLngRowCounter = 3
    If Not Purchasemin_core_generate_data_content_getoredata(PVarRequest, PTypResponse) Then Call .ErrorHandlerSet(1, "Error en la preparación de Datos del mineral."): GoTo EH

    ' Ore payment contents
    If Not Purchasemin_core_generate_data_content_getorepay(PVarRequest, PTypResponse) Then Call .ErrorHandlerSet(2, "Error en el cálculo del Peso pagable."): GoTo EH

    ' Gross value
    If Not Purchasemin_core_generate_data_content_getgrossvalue(PVarRequest, PTypResponse) Then Call .ErrorHandlerSet(3, "Error en el cálculo del Valor Bruto."): GoTo EH

    ' Deductions
    If Not Purchasemin_core_generate_data_content_getdeductions(PVarRequest, PTypResponse) Then Call .ErrorHandlerSet(4, "Error en el cálculo de Deducciones."): GoTo EH

    ' Bonus
    If Not Purchasemin_core_generate_data_content_getbonus(PVarRequest, PTypResponse) Then Call .ErrorHandlerSet(5, "Error en el cálculo de Bonos."): GoTo EH

    'Final purchase
    If Not Purchasemin_core_generate_data_content_getfinalpurchase(PVarRequest, PTypResponse) Then Call .ErrorHandlerSet(6, "Error en el cálculo final de la Liquidación de Minerales."): GoTo EH

  End With

  ' Return
  Let Purchasemin_core_generate_data_content = True

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Purchasemin_core_generate_data_content_getoredata( _
  ByVal PVarRequest As Variant, _
  ByRef PTypResponse As GLOTYPPURCHASE) As Boolean

  ' Obtains the basic data to calculate the purchase
  Dim VecAux0 As Variant, VecAux1 As Variant
  Dim StrAux0 As String
  Dim i As Long, j As Long, k As Long, xRows As Long, xCols As Long, xRowResize As Long, xRowTotals As Long

  Const sFN As String = "MIN::-core_main_spreadmodule_generate_data_purchasebase_getoredata"

  On Error GoTo EH
  Let Purchasemin_core_generate_data_content_getoredata = False

  With PTypResponse

    ' Ore Data
    Let xCols = 4
    Let .PDblTotalOreVal = 0

    If GLOBOO_VARIOUS Then ' n Providers
      Let xRows = 3 ' Title, Subtitles, contents, totals
    Else ' Just one provider, [nPur[A1P[Normal|Accurate]|1/P] | 1Pur[1/P]] TODO: MÉTODO DE ROWCOUNT EN MDL
      Let xRows = 2 + (UBound(PVarRequest) + 1) ' Title, Subtitles, [contents], totals
    End If
    Let xRowResize = xRows - 2
    Let xRowTotals = xRows + 1

    ' Fill matrix
    ReDim .PVecOre(0 To 0, 0 To xCols)
    Let .PVecOre(0, 0) = "DATOS MINERAL Y PESO BRUTO HÚMEDO"

    ' Add Subtitles at second row of main vector
    Let .PVecOre = RES.ArrayConcat(.PVecOre, Array(IIf(xRows > 3, "# ", VBA.vbNullString) & "Muestra", "Ingreso", "Contenidos", "Tipo", "PBH [TN]")) 'VecAux0)
    Let VecAux0 = Empty

    ' Get Samples
    For i = 0 To UBound(PVarRequest)
      Let PVarRequest(i, 0) = IIf(GLOBOO_VARIOUS, VBA.vbNullString, i + 1 & ". ") & PVarRequest(i, 0)
    Next i
    Let VecAux0 = RES.ArrayShift(PVarRequest, True, RByCols:=True)

    ' Get Date
    Let VecAux1 = RES.ArrayDelIndex(PVarRequest, 1, True, RByCols:=True)
    Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1, RByCols:=True)
    Let VecAux1 = Empty

    ' Get Contents
    Let VecAux1 = RES.ArrayDelIndex(PVarRequest, 6, True, RByCols:=True)
    Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1, RByCols:=True)
    Let VecAux1 = Empty

    ' Get Type
    Let VecAux1 = RES.ArrayDelIndex(PVarRequest, 5, True, RByCols:=True)
    Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1, RByCols:=True)
    Let VecAux1 = Empty

    ' Get Gross Weight
    Let VecAux1 = RES.ArrayDelIndex(PVarRequest, 4, True, RByCols:=True)
    Let .PDblTotalOreVal = Application.WorksheetFunction.Sum(VecAux1)
    Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1, RByCols:=True)
    Let VecAux1 = Empty

    ' Add data to main vector
    Let .PVecOre = RES.ArrayConcat(.PVecOre, VecAux0)
    Let VecAux0 = Empty

    ' Calculate totals
    ReDim VecAux0(0 To 0, 0 To xCols)
    Let VecAux0(0, 0) = "TOTAL PESO BRUTO HÚMEDO"
    If GLOBOO_FORMULAS Then Let VecAux0(0, xCols) = "=SUM(" & GLORNG_RNGBOX.Range("I6").Resize(xRowResize, 1).Address(False, False) & ")"
    If Not GLOBOO_FORMULAS Then Let VecAux0(0, xCols) = .PDblTotalOreVal

    ' Add totals to main vector
    Let .PVecOre = RES.ArrayAddAtLast(.PVecOre, VecAux0)
    Let VecAux0 = Empty

    If GLOBOO_FORMULAS Then Let .PStrTotalOreRng = "I" & (3 + xRowTotals)
    Let .PLngRowCounter = .PLngRowCounter + xRows + 1

  End With

  ' Return
  Let Purchasemin_core_generate_data_content_getoredata = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Purchasemin_core_generate_data_content_getorepay( _
  ByVal PVarRequest As Variant, _
  ByRef PTypResponse As GLOTYPPURCHASE) As Boolean

  Dim VecAux0 As Variant, VecAux1 As Variant, VecAux2 As Variant
  Dim StrAux0 As String, StrAux1 As String, StrAux2 As String, StrAux3 As String, StrAux4 As String
  Dim DblAux0 As Double, DblAux1 As Double, DblAux2 As Double, DblAux3 As Double, DblAux4 As Double
  Dim i As Long, xRows As Long, xCols As Long, xRowResize As Long, xRowTotals As Long, xRowFirstVal As Long

  Const sFN As String = "MIN::-core_main_spreadmodule_generate_data_purchasebase_getorepay"

  On Error GoTo EH
  Let Purchasemin_core_generate_data_content_getorepay = False

  With PTypResponse ' Ore Data

    Let xCols = 4
    Let .PDblTotalPayVal = 0
    'Let VecAux0 = Array("#", "Humedad [%]", "PBS [T]", "Merma [%]", "PNS [TN]")

    ' Determine matrix dimenssions
    If GLOBOO_VARIOUS Then ' Title, Subtitles, [contents], totals
      Let xRows = 3 ' Title, Subtitles, contents, totals
    Else ' Just one provider with average in data
      Let xRows = 2 + UBound(PVarRequest) + 1
      Let xRows = IIf(GLOBOO_AVERAGE, 3, xRows)
    End If
    Let xRowResize = xRows - 2
    Let xRowTotals = xRows + 1
    Let xRowFirstVal = .PLngRowCounter + 3

    ' Fill matrix
    ReDim .PVecOrePay(0 To 0, 0 To xCols)
    Let .PVecOrePay(0, 0) = "DETERMINACIÓN PESO PAGABLE"
    Let .PVecOrePay = RES.ArrayConcat(.PVecOrePay, Array("#", "Humedad [%]", "PBS [T]", "Merma [%]", "PNS [TN]"))

    ' Get PBH, H2O & LSS
    Let DblAux0 = 0: Let DblAux3 = 0: Let DblAux4 = 0
    If GLOBOO_AVERAGE Then

      ReDim VecAux1(UBound(PVarRequest))
      ReDim VecAux2(UBound(PVarRequest))
      For i = 0 To UBound(PVarRequest)
        Let DblAux1 = VBA.CDbl(PVarRequest(i, 4)) * (1 - (VBA.CDbl(PVarRequest(i, 2)) / 100)) ' PBS = PBH * (1 - (H2O/100))
        Let DblAux2 = DblAux1 * (1 - (VBA.CDbl(PVarRequest(i, 3)) / 100)) ' PNS = PBS * (1 - (LSS/100))
        Let DblAux0 = DblAux0 + VBA.CDbl(PVarRequest(i, 4)) ' PBH
        Let DblAux3 = DblAux3 + DblAux1 ' PBS
        Let DblAux4 = DblAux4 + DblAux2 ' PNS
        If GLOBOO_FORMULAS Then
          Let VecAux1(i) = GLORNG_RNGBOX.Range(.PStrTotalOreRng).Offset(-xRows + 1 + i, 0).Address(False, False) & "*(1-(" & VBA.Replace(PVarRequest(i, 2), ",", ".") & "/100))" ' PBS
          Let VecAux2(i) = "(" & VBA.CStr(VecAux1(i)) & ")*(1-(" & VBA.Replace(PVarRequest(i, 3), ",", ".") & "/100))"  ' PNS
        End If
      Next i
      ' Calculate Avg(H2O) & Avg(LSS)
      Let DblAux1 = (1 - (DblAux3 / DblAux0)) ' H2O*
      Let DblAux2 = (1 - (DblAux4 / DblAux0)) ' LSS*
      ' Fill main vector with results
      If GLOBOO_FORMULAS Then
        Let StrAux3 = "=(1-(SUM(" & VBA.Join(VecAux1, ",") & ")/" & GLORNG_RNGBOX.Range(.PStrTotalOreRng).Address(False, False) & "))*100" ' H2O
        Let StrAux4 = "=(1-(SUM(" & VBA.Join(VecAux2, ",") & ")/" & GLORNG_RNGBOX.Range("G" & xRowFirstVal).Address(False, False) & "))*100" ' LSS
        Let StrAux0 = "=" & GLORNG_RNGBOX.Range(.PStrTotalOreRng).Address(False, False) & "*(1-(" & GLORNG_RNGBOX.Range("F" & xRowFirstVal).Address(False, False) & "/100))"
        Let StrAux1 = "=" & GLORNG_RNGBOX.Range("G" & xRowFirstVal).Address(False, False) & "*(1-(" & GLORNG_RNGBOX.Range("H" & xRowFirstVal).Address(False, False) & "/100))"
        Let StrAux2 = "=SUM(" & GLORNG_RNGBOX.Range("I" & xRowFirstVal).Resize(xRowResize, 1).Address(False, False) & ")"
      End If
      Let .PVecOrePay = RES.ArrayAddAtLast(.PVecOrePay, Array("Promedios", IIf(GLOBOO_FORMULAS, StrAux3, DblAux1), IIf(GLOBOO_FORMULAS, StrAux0, DblAux3), IIf(GLOBOO_FORMULAS, StrAux4, DblAux2), IIf(GLOBOO_FORMULAS, StrAux1, DblAux4)))
      Let .PVecOrePay = RES.ArrayAddAtLast(.PVecOrePay, Array("TOTAL PESO PAGABLE", Empty, Empty, Empty, IIf(GLOBOO_FORMULAS, StrAux2, DblAux4)))

    Else

      ReDim VecAux0(0 To UBound(PVarRequest) + 1, 0 To xCols)
      For i = 0 To UBound(PVarRequest)
        Let VecAux0(i, 0) = i + 1 & ".." ' #
        Let VecAux0(i, 1) = PVarRequest(i, 2) ' H2O
        Let VecAux0(i, 3) = PVarRequest(i, 3) ' LSS
        If GLOBOO_FORMULAS Then
          Let StrAux0 = GLORNG_RNGBOX.Range(.PStrTotalOreRng).Offset(-xRows + 2 + i, 0).Address(False, False)
          Let VecAux0(i, 2) = "=" & StrAux0 & "*(1-(" & GLORNG_RNGBOX.Range("F" & xRowFirstVal + i).Address(False, False) & "/100))" ' PBS
          Let VecAux0(i, 4) = "=" & GLORNG_RNGBOX.Range("G" & xRowFirstVal + i).Address(False, False) & "*(1-(" & GLORNG_RNGBOX.Range("H" & xRowFirstVal + i).Address(False, False) & "/100))" ' PNS
        Else
          Let VecAux0(i, 2) = VBA.CDbl(PVarRequest(i, 4)) * (1 - (VBA.CDbl(PVarRequest(i, 2)) / 100)) ' PBS
          Let VecAux0(i, 4) = VBA.CDbl(VecAux0(i, 2)) * (1 - (VBA.CDbl(PVarRequest(i, 3)) / 100)) ' PNS
          Let .PDblTotalPayVal = .PDblTotalPayVal + VBA.CDbl(VecAux0(i, 4)) ' PNS TOTAL
        End If
      Next i
      Let VecAux0(UBound(VecAux0), 0) = "TOTAL PESO PAGABLE"
      Let VecAux0(UBound(VecAux0), xCols) = IIf(GLOBOO_FORMULAS, "=SUM(" & GLORNG_RNGBOX.Range("I" & xRowFirstVal).Resize(xRowResize, 1).Address(False, False) & ")", .PDblTotalPayVal)
      Let .PVecOrePay = RES.ArrayConcat(.PVecOrePay, VecAux0)

    End If
    ' RangeString of Totals
    If GLOBOO_FORMULAS Then Let .PStrTotalPayRng = "I" & (.PLngRowCounter + xRowTotals)
    Let .PLngRowCounter = .PLngRowCounter + xRows + 1

  End With

  ' Return
  Let Purchasemin_core_generate_data_content_getorepay = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Let VecAux2 = Empty
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Purchasemin_core_generate_data_content_getgrossvalue( _
  ByVal PVarRequest As Variant, _
  ByRef PTypResponse As GLOTYPPURCHASE) As Boolean

  Dim VecAux0 As Variant, VecAux1 As Variant, VecAux2 As Variant, VecAux3 As Variant, VecAux4 As Variant
  Dim StrAux0 As String, StrAux1 As String, StrAux2 As String, StrContent As String, StrUnit As String
  Dim BooAveraged As Boolean, BooEquals As Boolean
  Dim DblAux0 As Double, DblPNS As Double, DblGrade As Double, DblPrice As Double, DblValue As Double, DblTC As Double
  Dim i As Long, j As Long, k As Long, xRows As Long, xCols As Long, xRowResize As Long, xRowTotals As Long, xRowFirstVal As Long

  Const sFN As String = "MIN::-core_main_spreadmodule_generate_data_purchasebase_getgrossvalue"

  On Error GoTo EH
  Let Purchasemin_core_generate_data_content_getgrossvalue = False

  With PTypResponse

    ' Ore Data
    Let xCols = 4
    Let .PDblTotalGrossVal = 0

    ' Fill matrix
    Let xRowFirstVal = .PLngRowCounter + 3 ', xRowResize As Long, xRowTotals As Long, xRowFirstVal As Long
    Let xRows = (UBound(.PVecOrePay) - LBound(.PVecOrePay)) - 3 ', xRowResize As Long, xRowTotals As Long, xRowFirstVal As Long

    ' Prepare Contents, Grades, Price, [FgnValue], LocValue (Includes disglose of complex), Units and PNS
    For i = LBound(PVarRequest) To UBound(PVarRequest)

      Let DblAux0 = (VBA.CDbl(PVarRequest(i, 4)) * (1 - (VBA.CDbl(PVarRequest(i, 2)) / 100))) * (1 - (VBA.CDbl(PVarRequest(i, 3)) / 100)) ' ( PBH ((1-(H2O/100))) )*(1-(LSS/100))
      Let DblTC = VBA.CDbl(PVarRequest(i, 18))
      If RES.REGEX.isMineralComplex(VBA.CStr(PVarRequest(i, 5))) Then ' Complex
        Let VecAux1 = VBA.Split(VBA.CStr(PVarRequest(i, 6)), ";")
        Let VecAux2 = VBA.Split(VBA.CStr(PVarRequest(i, 7)), ";")
        Let VecAux3 = VBA.Split(VBA.CStr(PVarRequest(i, 9)), ";")
        Let VecAux4 = VBA.Split(VBA.CStr(PVarRequest(i, 8)), ";")
        For k = LBound(VecAux1) To UBound(VecAux1)
          Let StrContent = IIf(k = 0, VBA.vbNullString, "~") & VBA.CStr(VecAux1(k))
          Let DblGrade = VBA.CDbl(VecAux2(k))
          Let DblPrice = VBA.CDbl(VecAux3(k))
          Let StrUnit = VBA.CStr(VecAux4(k))
          Let DblValue = DblGrade * DblPrice * DblAux0 * DblTC
          Let VecAux0 = RES.ArrayConcat(VecAux0, RES.ArrayToBaseZero(Array(StrContent, DblGrade, DblPrice, Empty, DblValue, StrUnit, DblAux0), eqlRes2D))
        Next k
        Let VecAux1 = Empty
        Let VecAux2 = Empty
        Let VecAux3 = Empty
        Let VecAux4 = Empty
      Else ' Normal
        Let DblGrade = VBA.CDbl(PVarRequest(i, 7))
        Let DblPrice = VBA.CDbl(PVarRequest(i, 9))
        Let DblValue = DblGrade * DblPrice * DblAux0 * DblTC
        Let VecAux0 = RES.ArrayConcat(VecAux0, RES.ArrayToBaseZero(Array(PVarRequest(i, 6), DblGrade, DblPrice, Empty, DblValue, PVarRequest(i, 8), DblAux0), eqlRes2D))
      End If

    Next i

    ' If AVG is required, take fixed Grades (if AVG_ANYWY is active) and calculate prices averagin' values
    If GLOBOO_AVERAGE Then ' Just one provider. Average or calculate all purchases.

      For i = 0 To UBound(VecAux0)
        Let k = 1
        Let StrAux0 = VBA.LCase(VBA.Replace(VBA.CStr(VecAux0(i, 0)), "~", VBA.vbNullString))
        Let BooAveraged = False
        If Not StrAux0 = VBA.vbNullString Then
          Let DblGrade = VBA.CDbl(VecAux0(i, 1)) ' Grade
          Let DblPrice = VBA.CDbl(VecAux0(i, 2)) ' Price
          Let DblValue = VBA.CDbl(VecAux0(i, 4)) ' LocValue
          Let DblAux0 = VBA.CDbl(VecAux0(i, 6)) ' PNS
          If (i + 1) <= UBound(VecAux0) Then
            For j = (i + 1) To UBound(VecAux0)
              Let StrAux1 = VBA.LCase(VBA.Replace(VBA.CStr(VecAux0(j, 0)), "~", VBA.vbNullString))
              Let BooEquals = (Not StrAux0 = VBA.vbNullString And Not StrAux1 = VBA.vbNullString) ' Contents
              If BooEquals Then
                Let BooEquals = (StrAux0 = StrAux1) And (VBA.LCase(VBA.CStr(VecAux0(i, 5))) = VBA.LCase(VBA.CStr(VecAux0(j, 5))))
                If BooEquals Then ' Content(i) = Content(j) and Unit(i) = Unit(j) ?
                  ' Calculate Grade*, Price* and Value* average values
                  Let BooAveraged = True
                  Let StrAux2 = VBA.StrConv(VBA.CStr(VecAux0(i, 0)), vbProperCase)
                  Let VecAux0(i, 0) = VBA.vbNullString
                  Let VecAux0(j, 0) = VBA.vbNullString
                  Let DblGrade = DblGrade + (VBA.CDbl(VecAux0(j, 1)) * IIf(GLOBOO_AVERAGEANYWAY, 1, 0)) ' 1: Force Average grade
                  Let DblPrice = DblPrice + (VBA.CDbl(VecAux0(j, 2)) * IIf(GLOBOO_AVERAGEANYWAY, 0, 1)) ' 0: Force Average price
                  Let DblValue = DblValue + VBA.CDbl(VecAux0(j, 4))
                  Let DblAux0 = DblAux0 + VBA.CDbl(VecAux0(j, 6)) ' +PNS
                  Let k = k + 1
                End If
              End If
            Next j
          End If

          ' Obtains preliminar vector with: Contents, Grades, Prices, FgnValue, LocValue, Units, PNS
          If BooAveraged Then ' Capture averaged items
            'Let VecAux1 = RES.ArrayConcat(VecAux1, RES.ArrayToBaseZero(Array("Promed. " & StrAux2, DblGrade / k, DblPrice / k, (DblValue / k) / DblTC, DblValue / k, VecAux0(i, 5), DblAux0 / k), eqlRes2D))
            Let VecAux1 = RES.ArrayConcat(VecAux1, RES.ArrayToBaseZero(Array("Promed. " & StrAux2, DblGrade / k, DblPrice / k, DblValue / DblTC, DblValue, VecAux0(i, 5), DblAux0), eqlRes2D))
          Else ' Capture not averaged valid items
            Let VecAux1 = RES.ArrayConcat(VecAux1, RES.ArrayToBaseZero(Array(VBA.StrConv(StrAux0, vbProperCase), DblGrade, DblPrice, DblValue / DblTC, DblValue, VecAux0(i, 5), DblAux0), eqlRes2D))
          End If
        End If
      Next i

      ' Calculate Grades* or Prices* depending AVERAGE ANYWAY
      For i = LBound(VecAux1) To UBound(VecAux1)
        'Let VecAux1(i, IIf(GLOBOO_AVERAGEANYWAY, 1, 2)) = VBA.CDbl(VecAux1(i, 4)) / (VBA.CDbl(VecAux1(i, IIf(GLOBOO_AVERAGEANYWAY, 1, 2))) * VBA.CDbl(VecAux1(i, 6)))
        Let StrAux0 = VBA.LCase(VBA.CStr(VecAux1(i, 0)))
        Let DblValue = VBA.CDbl(VecAux1(i, 4)) ' +Gross Value
        Let DblPNS = VBA.CDbl(VecAux1(i, 6)) ' +PNS
        Let DblAux0 = VBA.CDbl(VecAux1(i, IIf(GLOBOO_AVERAGEANYWAY, 2, 1))) ' Grade or Price
        If VBA.InStr(StrAux0, "promed.") > 0 Then Let VecAux1(i, IIf(GLOBOO_AVERAGEANYWAY, 1, 2)) = DblValue / (DblPNS * DblAux0)
      Next i
      Let VecAux0 = Empty
      Let VecAux0 = VecAux1
      Let VecAux1 = Empty

    End If

    ' Determine matrix dimenssions
    Let xRows = 2 + UBound(VecAux0) + 1 'Title, Subtitles, [contents], totals
    Let xRowResize = xRows - 2
    Let xRowTotals = xRows + 1

    ' Obtains semifinal vector (and units in Type units vector) in numbers or formulas with: Contents, Grades, Prices, [FngValues], LocValue
    Let j = -1
    Let DblTC = VBA.CDbl(PVarRequest(0, 18))
    For i = LBound(VecAux0) To UBound(VecAux0)
      If GLOBOO_FORMULAS Then
        If GLOBOO_AVERAGE Then
          Let VecAux0(i, 4) = "=" & GLORNG_RNGBOX.Range("F" & xRowFirstVal + i).Address(False, False) & "*" & GLORNG_RNGBOX.Range("G" & xRowFirstVal + i).Address(False, False) & "*" & VBA.CStr(VBA.Replace(VecAux0(i, 6), ",", "."))
        Else
          Let j = j + IIf(VBA.InStr(VBA.CStr(VecAux0(i, 0)), "~") > 0, 0, 1)
          Let VecAux0(i, 4) = "=" & GLORNG_RNGBOX.Range("F" & xRowFirstVal + i).Address(False, False) & "*" & GLORNG_RNGBOX.Range("G" & xRowFirstVal + i).Address(False, False) & "*" & GLORNG_RNGBOX.Range(PTypResponse.PStrTotalOreRng).Offset(j + 3, 0).Address(False, False)
        End If
        Let VecAux0(i, 3) = IIf(GLOBOO_ME, "=" & GLORNG_RNGBOX.Range("I" & xRowFirstVal + i).Address(False, False) & "/" & GLORNG_RNGBOX.Range("B10").Address(False, False), Empty)
      Else
        Let DblAux0 = VBA.CDbl(VecAux0(i, 4))
        Let VecAux0(i, 3) = IIf(GLOBOO_ME, DblAux0 / DblTC, Empty)
        Let .PDblTotalGrossVal = .PDblTotalGrossVal + DblAux0 'VBA.CDbl(VecAux0(i, 4))
      End If
    Next i
    Let .PVecUnits = RES.ArrayToBaseZero(RES.ArrayDelIndex(VecAux0, 5, RByCols:=True), eqlRes1D)
    Call RES.ArrayPop(VecAux0, RByCols:=True)

    ' Obtains Final vector (and units in Type units vector) in numbers or formulas with: Contents, Grades, Prices, [FngValues], LocValue
    ReDim .PVecGrossValue(0 To 0, 0 To 4)
    Let .PVecGrossValue(0, 0) = "DETERMINACIÓN VALOR BRUTO"
    Let .PVecGrossValue = RES.ArrayConcat(.PVecGrossValue, Array("#", "Ley", GLOSTR_CURFOREIGN & "/PTO", IIf(GLOBOO_ME, "Valor [" & GLOSTR_CURFOREIGN & "]", Empty), "Valor [" & GLOSTR_CURLOCAL & "]"))
    Let .PVecGrossValue = RES.ArrayConcat(.PVecGrossValue, VecAux0)
    If GLOBOO_FORMULAS Then
      Let StrAux1 = IIf(GLOBOO_ME, "=SUM(" & GLORNG_RNGBOX.Range("H" & xRowFirstVal).Resize(xRowResize, 1).Address(False, False) & ")", Empty)
      Let StrAux0 = "=SUM(" & GLORNG_RNGBOX.Range("I" & xRowFirstVal).Resize(xRowResize, 1).Address(False, False) & ")"
      Let .PVecGrossValue = RES.ArrayConcat(.PVecGrossValue, Array("TOTAL VALOR BRUTO", Empty, Empty, IIf(GLOBOO_ME, StrAux1, Empty), StrAux0))
    Else
      Let .PVecGrossValue = RES.ArrayConcat(.PVecGrossValue, Array("TOTAL VALOR BRUTO", Empty, Empty, IIf(GLOBOO_ME, .PDblTotalGrossVal / DblTC, Empty), .PDblTotalGrossVal))
    End If

    ' RangeString of Totals
    Let .PStrTotalGrossRng = IIf(GLOBOO_FORMULAS, "I" & (.PLngRowCounter + xRowTotals), Empty)
    Let .PLngRowCounter = .PLngRowCounter + xRows + 1

  End With

  ' Return
  Let Purchasemin_core_generate_data_content_getgrossvalue = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Let VecAux2 = Empty
  Let VecAux3 = Empty
  Let VecAux4 = Empty
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Purchasemin_core_generate_data_content_getdeductions( _
  ByVal PVarRequest As Variant, _
  ByRef PTypResponse As GLOTYPPURCHASE) As Boolean

  Dim VecAux0 As Variant, VecAux1 As Variant, VecAux2 As Variant, VecAux3 As Variant
  Dim StrAux0 As String, StrAux1 As String, StrAux2 As String
  Dim DblAux0 As Double, DblAux1 As Double, DblAux2 As Double, DblTC As Double
  Dim i As Long, j As Long, k As Long, xRows As Long, xCols As Long, xRowResize As Long, xRowTotals As Long, xRowFirstVal As Long, xRowFirstVal2 As Long

  Const sFN As String = "MIN::-core_main_spreadmodule_generate_data_purchasebase_getdeductions"

  On Error GoTo EH
  Let Purchasemin_core_generate_data_content_getdeductions = False

  With PTypResponse

    ' Initial and base values
    Let .PDblTotalDedVal = 0
    Let xCols = 4
    Let xRows = 1 'Title, [Subtitles, contents], totals
    Let i = 0
    Let k = 0
    Let xRowFirstVal = .PLngRowCounter + 3
    Let xRowFirstVal2 = xRowFirstVal
    Let DblTC = VBA.CDbl(PVarRequest(0, 18))

    ' Calcualte deductions
    ReDim .PVecDeductions(0 To 1, 0 To xCols)
    Let .PVecDeductions(0, 0) = "(-) DEDUCCIONES"

    ' Get Contents and determine RM's
    Let VecAux1 = MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigpurrm_sys, eqlMdlArray, "ELEMENTO, VENTA INTERNA") ' RM's from db
    Let .PVecDeductions(1, 0) = "RETENCIONES DE LEY - REGALÍAS [%]:"

    If GLOBOO_VARIOUS Then ' One register arrived
      Let VecAux0 = RES.ArrayUniques(VBA.Split(PVarRequest(0, 6), ";")) ' Get elements contents from main pur vector
    Else ' One provider with possible lot of purchases and resumed or disglosed preferences
      Let VecAux0 = RES.ArrayToBaseZero(RES.ArrayDelIndex(PVarRequest, 6, True, RByCols:=True), eqlRes1D) ' Get elements contents from main pur vector
      Let VecAux0 = RES.ArrayUniques(VBA.Split(VBA.Join(VecAux0, ";"), ";"))
    End If

    Let VecAux3 = RES.ArrayToBaseZero(RES.ArrayShift(VecAux1, True, RByCols:=True), eqlRes1D)
    For i = LBound(VecAux0) To UBound(VecAux0)
      Let j = RES.ArrayGetValuePosition(VecAux3, VBA.CStr(VecAux0(i)))
      If j >= 0 Then Let VecAux2 = RES.ArrayConcat(VecAux2, RES.ArrayDelIndex(VecAux1, j, True))
    Next i
    Let VecAux0 = Empty
    Let VecAux1 = Empty
    Let VecAux3 = Empty

    Let j = UBound(VecAux2) - LBound(VecAux2)
    Let j = IIf(GLOBOO_RESUMED_RM, 0, j)
    For i = 0 To j
      If GLOBOO_RESUMED_RM Then Let DblAux0 = Application.WorksheetFunction.Sum(RES.ArrayPop(VecAux2, RByCols:=True)): Let StrAux2 = "RM" ' %Retentions
      If Not GLOBOO_RESUMED_RM Then Let DblAux0 = VBA.CDbl(VecAux2(i, 1)): Let StrAux2 = i + 1 & ". RM-" & VBA.CStr(VecAux2(i, 0))  ' %Retentions
      If GLOBOO_FORMULAS Then
        Let StrAux0 = "=" & GLORNG_RNGBOX.Range("I" & xRowFirstVal + i).Address(False, False) & "/" & GLORNG_RNGBOX.Range("B10").Address(False, False) ' $ME
        Let StrAux1 = "=" & GLORNG_RNGBOX.Range(.PStrTotalGrossRng).Address(False, False) & "*" & GLORNG_RNGBOX.Range("F" & xRowFirstVal + i).Address(False, False) & "/100" ' $MN
      Else
        Let DblAux1 = (.PDblTotalGrossVal / DblTC) * (DblAux0 / 100) ' #ME
        Let DblAux2 = (.PDblTotalGrossVal) * (DblAux0 / 100) ' #MN
        Let .PDblTotalDedVal = .PDblTotalDedVal + DblAux2
      End If
      Let .PVecDeductions = RES.ArrayAddAtLast(.PVecDeductions, Array(StrAux2, VBA.Replace(DblAux0, ",", "."), Empty, IIf(GLOBOO_ME, IIf(GLOBOO_FORMULAS, StrAux0, VBA.Replace(DblAux1, ",", ".")), Empty), IIf(GLOBOO_FORMULAS, StrAux1, VBA.Replace(DblAux2, ",", "."))))
    Next i
    Let xRowFirstVal2 = xRowFirstVal + i + 1
    Let VecAux2 = Empty

    ' Get Retentions and determine %
    Let VecAux1 = MDL.MGet(eqlMdlSheetTmp, eqlMdlTblConfigpurretentions_tmp, eqlMdlArray, "CONCEPTO, PORCENTAJE") ' Get retentions concept and %value
    Let .PVecDeductions = RES.ArrayConcat(.PVecDeductions, Array("RETENCIONES DE LEY - INSTITUCIONALES [%]:", Empty, Empty, Empty, Empty))

    Let VecAux0 = VBA.Split(PVarRequest(0, 12), ";") ' Get selected retentions from main pur vector
    For i = LBound(VecAux0) To UBound(VecAux0)
      If VBA.CDbl(VecAux0(i)) = 1 Then Let VecAux2 = RES.ArrayConcat(VecAux2, RES.ArrayDelIndex(VecAux1, i, True))
    Next i
    Let VecAux0 = Empty
    Let VecAux1 = Empty

    Let j = UBound(VecAux2) - LBound(VecAux2)
    Let j = IIf(GLOBOO_RESUMED_RETENTIONS, 0, j)
    For i = 0 To j
      If GLOBOO_RESUMED_RETENTIONS Then Let DblAux0 = Application.WorksheetFunction.Sum(RES.ArrayPop(VecAux2, RByCols:=True)): Let StrAux2 = "RESUMEN" ' %Retentions
      If Not GLOBOO_RESUMED_RETENTIONS Then Let DblAux0 = VBA.CDbl(VecAux2(i, 1)): Let StrAux2 = i + 1 & ". " & VBA.CStr(VecAux2(i, 0))  ' %Retentions
      If GLOBOO_FORMULAS Then
        Let StrAux0 = "=" & GLORNG_RNGBOX.Range("I" & xRowFirstVal + i).Address(False, False) & "/" & GLORNG_RNGBOX.Range("B10").Address(False, False) ' $ME
        Let StrAux1 = "=" & GLORNG_RNGBOX.Range(.PStrTotalGrossRng).Address(False, False) & "*" & GLORNG_RNGBOX.Range("F" & xRowFirstVal2 + i).Address(False, False) & "/100" ' $MN
      Else
        Let DblAux1 = (.PDblTotalGrossVal / DblTC) * (DblAux0 / 100) ' #ME
        Let DblAux2 = (.PDblTotalGrossVal) * (DblAux0 / 100) ' #MN
        Let .PDblTotalDedVal = .PDblTotalDedVal + DblAux2
      End If
      Let .PVecDeductions = RES.ArrayConcat(.PVecDeductions, Array(StrAux2, VBA.Replace(DblAux0, ",", "."), Empty, IIf(GLOBOO_ME, IIf(GLOBOO_FORMULAS, StrAux0, VBA.Replace(DblAux1, ",", ".")), Empty), IIf(GLOBOO_FORMULAS, StrAux1, VBA.Replace(DblAux2, ",", "."))))
    Next i
    Let xRowFirstVal2 = xRowFirstVal2 + i + 1
    Let VecAux2 = Empty

    ' Get Retentions others and determine %
    Let VecAux1 = MDL.MGet(eqlMdlSheetTmp, eqlMdlTblConfigpurretentionsother_tmp, eqlMdlArray, "CONCEPTO, PORCENTAJE") ' Get retentions others concept and %value
    Let StrAux0 = VBA.CStr(PVarRequest(0, 13)) ' Get selected retentions others from main pur vector
    If VBA.InStr(StrAux0, "1") > 0 Then
      Let VecAux0 = VBA.Split(StrAux0, ";") ' Get selected retentions others from main pur vector
      Let .PVecDeductions = RES.ArrayConcat(.PVecDeductions, Array("RETENCIONES INSTITUCIONALES OTRAS [%]:", Empty, Empty, Empty, Empty))
      For i = LBound(VecAux0) To UBound(VecAux0)
        If VBA.CDbl(VecAux0(i)) = 1 Then Let VecAux2 = RES.ArrayConcat(VecAux2, RES.ArrayDelIndex(VecAux1, i, True))
      Next i
      Let j = UBound(VecAux2) - LBound(VecAux2)
      Let j = IIf(GLOBOO_RESUMED_RETENTIONSOTHERS, 0, j)
      For i = 0 To j
        If GLOBOO_RESUMED_RETENTIONSOTHERS Then Let DblAux0 = Application.WorksheetFunction.Sum(RES.ArrayPop(VecAux2, RByCols:=True)): Let StrAux2 = "RESUMEN" ' %Retentions
        If Not GLOBOO_RESUMED_RETENTIONSOTHERS Then Let DblAux0 = VBA.CDbl(VecAux2(i, 1)): Let StrAux2 = i + 1 & ". " & VBA.CStr(VecAux2(i, 0))  ' %Retentions
        If GLOBOO_FORMULAS Then
          Let StrAux0 = "=" & GLORNG_RNGBOX.Range("I" & xRowFirstVal2 + i).Address(False, False) & "/" & GLORNG_RNGBOX.Range("B10").Address(False, False) ' $ME
          Let StrAux1 = "=" & GLORNG_RNGBOX.Range(.PStrTotalGrossRng).Address(False, False) & "*" & GLORNG_RNGBOX.Range("F" & xRowFirstVal2 + i).Address(False, False) & "/100" ' $MN
        Else
          Let DblAux1 = (.PDblTotalGrossVal / DblTC) * (DblAux0 / 100) ' #ME
          Let DblAux2 = (.PDblTotalGrossVal) * (DblAux0 / 100) ' #MN
          Let .PDblTotalDedVal = .PDblTotalDedVal + DblAux2
        End If
        Let .PVecDeductions = RES.ArrayConcat(.PVecDeductions, Array(StrAux2, VBA.Replace(DblAux0, ",", "."), Empty, IIf(GLOBOO_ME, IIf(GLOBOO_FORMULAS, StrAux0, VBA.Replace(DblAux1, ",", ".")), Empty), IIf(GLOBOO_FORMULAS, StrAux1, VBA.Replace(DblAux2, ",", "."))))
      Next i
      Let xRowFirstVal2 = xRowFirstVal2 + i + 1
    End If
    Let VecAux0 = Empty
    Let VecAux1 = Empty
    Let VecAux2 = Empty

    ' Get Discounts and determine its value
    Let StrAux0 = VBA.CStr(PVarRequest(0, 11)) ' Get discounts pipe (|) separated from main pur vector
    If Not StrAux0 = VBA.vbNullString Then
      Let VecAux0 = VBA.Split(StrAux0, "|") ' Get discounts semicolon (;) separated from main pur vector
      Let VecAux0(0) = VBA.Replace(VBA.Replace(VBA.Replace(VBA.CStr(VecAux0(0)), ";", VBA.vbTab), "[eqlpipe]", "|"), "[eqlsemicolon]", ";")
      Let VecAux2 = RES.ArrayToBaseZero(VBA.Split(VBA.CStr(VecAux0(0)), VBA.vbTab), eqlRes2D, False)
      Let VecAux2 = RES.ArrayConcat(VecAux2, RES.ArrayToBaseZero(VBA.Split(VBA.CStr(VecAux0(1)), ";"), eqlRes2D, False), RByCols:=True)
      Let .PVecDeductions = RES.ArrayConcat(.PVecDeductions, Array("DESCUENTOS:", Empty, Empty, Empty, Empty))
      Let j = UBound(VecAux2) - LBound(VecAux2)
      Let j = IIf(GLOBOO_RESUMED_DISCOUNTS, 0, j)
      For i = 0 To j
        If GLOBOO_RESUMED_DISCOUNTS Then Let DblAux0 = Application.WorksheetFunction.Sum(RES.ArrayPop(VecAux2, RByCols:=True)): Let StrAux2 = "RESUMEN" ' %Retentions
        If Not GLOBOO_RESUMED_DISCOUNTS Then Let DblAux0 = VBA.CDbl(VecAux2(i, 1)): Let StrAux2 = i + 1 & ". " & VBA.CStr(VecAux2(i, 0))  ' %Retentions
        If GLOBOO_FORMULAS Then
          Let StrAux0 = "=" & GLORNG_RNGBOX.Range("I" & xRowFirstVal2 + i).Address(False, False) & "/" & GLORNG_RNGBOX.Range("B10").Address(False, False) ' $ME
        Else
          Let DblAux1 = (DblAux0 / DblTC) ' #ME
        End If
        Let DblAux2 = DblAux0 ' #MN
        Let .PDblTotalDedVal = .PDblTotalDedVal + DblAux2
        Let .PVecDeductions = RES.ArrayConcat(.PVecDeductions, Array(StrAux2, Empty, Empty, IIf(GLOBOO_ME, IIf(GLOBOO_FORMULAS, StrAux0, VBA.Replace(DblAux1, ",", ".")), Empty), VBA.Replace(DblAux0, ",", ".")))
      Next i
    End If
    Let VecAux0 = Empty
    Let VecAux1 = Empty
    Let VecAux2 = Empty

    ' Totals
    Let xRowResize = UBound(.PVecDeductions)
    Let xRowTotals = xRowResize + 1
    If GLOBOO_FORMULAS Then
      Let StrAux0 = IIf(GLOBOO_ME, "=SUM(" & GLORNG_RNGBOX.Range("H" & xRowFirstVal - 1).Resize(xRowResize, 1).Address(False, False) & ")", Empty)
      Let StrAux1 = "=SUM(" & GLORNG_RNGBOX.Range("I" & xRowFirstVal - 1).Resize(xRowResize, 1).Address(False, False) & ")"
    End If
    Let .PVecDeductions = RES.ArrayConcat(.PVecDeductions, Array("TOTAL DEDUCCIONES", Empty, Empty, IIf(GLOBOO_ME, IIf(GLOBOO_FORMULAS, StrAux0, VBA.Replace(.PDblTotalDedVal / DblTC, ",", ".")), Empty), IIf(GLOBOO_FORMULAS, StrAux1, VBA.Replace(.PDblTotalDedVal, ",", "."))))

    ' RangeString of Totals
    Let .PStrTotalDedRng = IIf(GLOBOO_FORMULAS, "I" & (.PLngRowCounter + xRowTotals + 1), Empty)
    Let .PLngRowCounter = .PLngRowCounter + xRowTotals

  End With

  ' Return
  Let Purchasemin_core_generate_data_content_getdeductions = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Let VecAux2 = Empty
  Let VecAux3 = Empty
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Purchasemin_core_generate_data_content_getbonus( _
  ByVal PVarRequest As Variant, _
  ByRef PTypResponse As GLOTYPPURCHASE) As Boolean

  Dim VecAux0 As Variant, VecAux1 As Variant
  Dim StrAux0 As String, StrAux1 As String, StrAux2 As String
  Dim DblAux0 As Double, DblAux1 As Double, DblAux2 As Double, DblTC As Double
  Dim i As Long, j As Long, xRows As Long, xCols As Long, xRowResize As Long, xRowTotals As Long, xRowFirstVal As Long

  Const sFN As String = "MIN::-core_main_spreadmodule_generate_data_purchasebase_getbonus"

  On Error GoTo EH
  Let Purchasemin_core_generate_data_content_getbonus = False

  With PTypResponse

    ' Initial and base values
    Let .PDblTotalBonVal = 0
    Let xCols = 4
    Let xRows = 1 ' Title, [Subtitles, contents], totals
    Let xRowFirstVal = .PLngRowCounter + 3
    Let DblTC = VBA.CDbl(PVarRequest(0, 18))

    ' Get Bonus and determine its value
    Let StrAux0 = VBA.CStr(PVarRequest(0, 10)) ' Get bonus pipe (|) separated from main pur vector
    Let .PBooBonus = Not (StrAux0 = VBA.vbNullString)

    If .PBooBonus Then

      Let VecAux0 = VBA.Split(StrAux0, "|") ' Get bonus semicolon (;) separated from main pur vector
      Let VecAux0(0) = VBA.Replace(VBA.Replace(VBA.Replace(VBA.CStr(VecAux0(0)), ";", VBA.vbTab), "[eqlpipe]", "|"), "[eqlsemicolon]", ";")
      Let VecAux1 = RES.ArrayToBaseZero(VBA.Split(VBA.CStr(VecAux0(0)), VBA.vbTab), eqlRes2D, False)
      Let VecAux1 = RES.ArrayConcat(VecAux1, RES.ArrayToBaseZero(VBA.Split(VBA.CStr(VecAux0(1)), ";"), eqlRes2D, False), RByCols:=True)
      ReDim .PVecBonus(0 To 0, 0 To xCols)
      Let .PVecBonus(0, 0) = "(+) BONOS"
      Let j = UBound(VecAux1) - LBound(VecAux1)
      Let j = IIf(GLOBOO_RESUMED_BONUS, 0, j)

      For i = 0 To j
        If GLOBOO_RESUMED_BONUS Then Let DblAux0 = Application.WorksheetFunction.Sum(RES.ArrayPop(VecAux1, RByCols:=True)): Let StrAux2 = "RESUMEN" ' Bonus
        If Not GLOBOO_RESUMED_BONUS Then Let DblAux0 = VBA.CDbl(VecAux1(i, 1)): Let StrAux2 = i + 1 & ". " & VBA.CStr(VecAux1(i, 0))  ' Bonus
        If GLOBOO_FORMULAS Then
          Let StrAux0 = "=" & GLORNG_RNGBOX.Range("I" & xRowFirstVal + i).Address(False, False) & "/" & GLORNG_RNGBOX.Range("B10").Address(False, False) ' $ME
        Else
          Let DblAux1 = (DblAux0 / DblTC) ' #ME
        End If
        Let DblAux2 = DblAux0 ' #MN
        Let .PDblTotalBonVal = .PDblTotalBonVal + DblAux2
        Let .PVecBonus = RES.ArrayConcat(.PVecBonus, Array(StrAux2, Empty, Empty, IIf(GLOBOO_ME, IIf(GLOBOO_FORMULAS, StrAux0, VBA.Replace(DblAux1, ",", ".")), Empty), VBA.Replace(DblAux0, ",", ".")))
      Next i

      ' Totals
      Let xRows = UBound(.PVecBonus) + 1
      Let xRowResize = xRows - 1
      Let xRowTotals = xRowResize + 1
      If GLOBOO_FORMULAS Then
        Let StrAux0 = IIf(GLOBOO_ME, "=SUM(" & GLORNG_RNGBOX.Range("H" & xRowFirstVal).Resize(xRowResize, 1).Address(False, False) & ")", Empty)
        Let StrAux1 = "=SUM(" & GLORNG_RNGBOX.Range("I" & xRowFirstVal).Resize(xRowResize, 1).Address(False, False) & ")"
      End If
      Let .PVecBonus = RES.ArrayConcat(.PVecBonus, Array("TOTAL BONOS", Empty, Empty, IIf(GLOBOO_ME, IIf(GLOBOO_FORMULAS, StrAux0, VBA.Replace(.PDblTotalBonVal / DblTC, ",", ".")), Empty), IIf(GLOBOO_FORMULAS, StrAux1, VBA.Replace(.PDblTotalBonVal, ",", "."))))

    End If
    Let VecAux0 = Empty
    Let VecAux1 = Empty

    ' RangeString of Totals
    Let .PStrTotalBonRng = IIf(GLOBOO_FORMULAS, "I" & (.PLngRowCounter + xRowTotals + 1), Empty)
    Let .PLngRowCounter = .PLngRowCounter + xRows

  End With

  ' Return
  Let Purchasemin_core_generate_data_content_getbonus = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Purchasemin_core_generate_data_content_getfinalpurchase( _
  ByVal PVarRequest As Variant, _
  ByRef PTypResponse As GLOTYPPURCHASE) As Boolean

  Dim NTX As ClassEqlControllerNumToStr

  Dim StrBonusRng As String
  Dim DblAux0 As Double, DblTC As Double
  Dim xRows As Long, xCols As Long, xRowFirstVal As Long

  Const sFN As String = "MIN::-core_main_spreadmodule_generate_data_purchasebase_getbonus"

  On Error GoTo EH
  Let Purchasemin_core_generate_data_content_getfinalpurchase = False

  With PTypResponse

    ' Initial and base values
    Let xCols = 8

    ' Determine matrix dimenssions
    Let xRows = 3 'Title, (Subtitles, contents), totals
    Let xRowFirstVal = .PLngRowCounter + 3 + IIf(.PBooBonus, 2, 0)
    Let DblTC = VBA.CDbl(PVarRequest(0, 18))

    ' Fill matrix PLngRowCounter
    Let DblAux0 = (.PDblTotalGrossVal - .PDblTotalDedVal) + .PDblTotalBonVal
    ReDim .PVecFinal(0 To xRows, 0 To xCols)

    Let .PVecFinal(0, 0) = "LIQUIDACIÓN FINAL (VALOR BRUTO - DEDUCCIONES" & IIf(.PBooBonus, " + BONOS", VBA.vbNullString) & ")"
    If GLOBOO_ME Then ' ME
      Let .PVecFinal(0, 7) = "VALOR " & GLOSTR_CURFOREIGN
      If GLOBOO_FORMULAS Then
        Let .PVecFinal(1, 7) = "=" & GLORNG_RNGBOX.Range("I" & xRowFirstVal).Address(False, False) & "/" & GLORNG_RNGBOX.Range("B10").Address(False, False)
      Else
        Let .PVecFinal(1, 7) = VBA.Replace(DblAux0 / DblTC, ",", ".")
      End If
    End If

    Let .PVecFinal(0, 8) = "VALOR " & GLOSTR_CURLOCAL ' MN
    If GLOBOO_FORMULAS Then
      If .PDblTotalBonVal = 0 Then
        Let StrBonusRng = VBA.vbNullString
      Else
        Let StrBonusRng = "," & GLORNG_RNGBOX.Range(.PStrTotalBonRng).Address(False, False)
      End If
      Let .PVecFinal(1, 8) = "=SUM(" & GLORNG_RNGBOX.Range(.PStrTotalGrossRng).Address(False, False) & ",-" & GLORNG_RNGBOX.Range(.PStrTotalDedRng).Address(False, False) & StrBonusRng & ")"
    Else
      Let .PVecFinal(1, 8) = VBA.Replace(DblAux0, ",", ".")
    End If

    Let .PVecFinal(1, 0) = "Líquido pagable (" & IIf(DblAux0 = 0, "Sin saldo", IIf(DblAux0 > 0, "Saldo a favor", "Saldo en contra")) & "):"
    Let .PVecFinal(2, 0) = "Son:"
    If GLOBOO_FORMULAS Then
      Let .PVecFinal(2, 1) = "=EQL_NUMERO_A_TEXTO(" & GLORNG_RNGBOX.Range("I" & xRowFirstVal).Address(False, False) & ",""" & VBA.UCase(GLOSTR_CURRENCY) & """,""" & VBA.UCase(GLOSTR_CURRENCIES) & """,0)"
    Else
      Set NTX = New ClassEqlControllerNumToStr
      Set NTX.ClassEqlControllerInterface_ResourcesApp = RES
      Let .PVecFinal(2, 1) = NTX.NumToString(DblAux0, GLOSTR_CURRENCY, GLOSTR_CURRENCIES, False)
      Set NTX = Nothing
    End If

  End With

  ' Return
  Let Purchasemin_core_generate_data_content_getfinalpurchase = True

EH:
  Set NTX = Nothing
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function





' ERRORS MANNAGEMENT AND ASSERTSIONS ----------------------------------------------- '
' Debug_assert
Private Sub Debug_assert(ByVal Cond As Boolean, Optional Origin As String = VBA.vbNullString)

  #If Debbuging Then
    Debug.Assert Cond
    If Not Cond Then Debug.Print "Assert on: " & Origin
  #End If

End Sub




' METHODS PUBLIC ------------------------------------------------------------------- '
' INTERFACE METHODS LIST ----------------------------------------------------------- '
' ABSTRACT INTERFACE UI METHODS LIST AND UDF METHODS LIST (PUBLIC) ----------------- '
Public Function ClassEqlControllerInterface_GetSolution(ByVal CStrRequester As String, ByRef CVarRequest As Variant, ByRef CVarResponse As Variant) As Boolean

  On Error GoTo EH
  Let ClassEqlControllerInterface_GetSolution = Purchasemin_core_main(CVarRequest, CVarResponse)

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise("MIN::IGetSolution")

End Function

Private Function ClassEqlControllerInterface_GetSolutionUDF(ByVal CStrRequester As String, ByVal CVarRequest As Variant) As Variant
End Function
' METHODS PUBLIC --------------------------------------------------------------------'
' ============ CLASS EXCLIQ LITE CONTROLLER PURCHASE MINERAL STRUCTURE ============= '

