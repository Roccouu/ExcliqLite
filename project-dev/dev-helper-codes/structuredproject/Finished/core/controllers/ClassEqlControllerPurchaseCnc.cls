VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassEqlControllerPurchaseCnc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ========== CLASS EXCLIQ LITE CONTROLLER PURCHASE CONCENTRATE STRUCTURE =========== '
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




' ========== CLASS EXCLIQ LITE CONTROLLER PURCHASE CONCENTRATE STRUCTURE =========== '
' ABSTRACT IMPLEMENTATIONS --------------------------------------------------------- '
' OBJECT VARIABLES (GLOOBJ_) ------------------------------------------------------- '
' VARIABLES (GLOStr_) -------------------------------------------------------------- '
' PROPERTIES ----------------------------------------------------------------------- '
' ABSTRACT INTERFACE PROPERTIES IMPLEMENTATION ------------------------------------- '
' CONSTRUCTOR/DESTRUCTOR ----------------------------------------------------------- '
' CORE METHODS PRIVATE ------------------------------------------------------------- '
' Purchasecnc_core_main
' Purchasecnc_core_generate_header
' Purchasecnc_core_generate_content
' Purchasecnc_core_generate_content_weight
' Purchasecnc_core_generate_content_contents
' Purchasecnc_core_generate_content_assessment
' Purchasecnc_core_generate_content_assessment_value
' Purchasecnc_core_generate_content_assessment_treatmentcost
' Purchasecnc_core_generate_content_assessment_refinement
' Purchasecnc_core_generate_content_assessment_penalties
' Purchasecnc_core_generate_content_deductions
' Purchasecnc_core_generate_content_deductions_exports
' Purchasecnc_core_generate_content_deductions_rmtaxes
' Purchasecnc_core_generate_content_deductions_lawtaxes
' Purchasecnc_core_generate_content_deductions_othertaxes
' Purchasecnc_core_generate_content_deductions_discounts
' Purchasecnc_core_generate_content_bonus
' Purchasecnc_core_generate_content_finalpurchase
' ASSETS --------------------------------------------------------------------------- '
' Purchasecnc_core_assets_srcarray_to_dstarray
' ERRORS MANNAGEMENT AND ASSERTSIONS ----------------------------------------------- '
' Debug_assert
' METHODS PUBLIC ------------------------------------------------------------------- '
' ABSTRACT INTERFACE UI METHODS LIST AND UDF METHODS LIST (PUBLIC) ----------------- '
' ClassEqlControllerInterface_GetSolution
' ClassEqlControllerInterface_GetSolutionUDF
' ========== CLASS EXCLIQ LITE CONTROLLER PURCHASE CONCENTRATE STRUCTURE =========== '



' ========== CLASS EXCLIQ LITE CONTROLLER PURCHASE CONCENTRATE STRUCTURE =========== '
' ABSTRACT IMPLEMENTATIONS --------------------------------------------------------- '
Implements ClassEqlControllerInterface

' OBJECT VARIABLES (GLOOBJ_) ------------------------------------------------------- '
Private RES As ClassEqlResources
Private MDL As ClassEqlModel

' VARIABLES (GLOStr_) -------------------------------------------------------------- '
Private Type GLOTYPCOMPONENT
  Vector As Variant
  Exists As Boolean
  StrRngA1 As String
  LngRngA1 As Long
End Type

Private Type GLOTYPPURCHASE
  PTitle As GLOTYPCOMPONENT
  PInvoice As GLOTYPCOMPONENT
  PProvider As GLOTYPCOMPONENT
  PWeight As GLOTYPCOMPONENT ' PNS
  PContents As GLOTYPCOMPONENT
  PAssessmentMain As GLOTYPCOMPONENT ' Total/T
  PAssessment As GLOTYPCOMPONENT
  PTreatmentCost As GLOTYPCOMPONENT
  PRefinement As GLOTYPCOMPONENT
  PPenalties As GLOTYPCOMPONENT
  PDeductions As GLOTYPCOMPONENT
  PDExports As GLOTYPCOMPONENT
  PDRMRets As GLOTYPCOMPONENT
  PDLawRets As GLOTYPCOMPONENT
  PDOthRets As GLOTYPCOMPONENT
  PDDiscnts As GLOTYPCOMPONENT
  PBonus As GLOTYPCOMPONENT
  PFinal As GLOTYPCOMPONENT

  UnitsGraVec As Variant
  UnitsCtzVec As Variant
  UnitsTraVec As Variant

  TotalAssessmentRng As String
  TotalDeductionsRng As String
  TotalBonusRng As String

  TotalPBHDbl As Double
  TotalPBSDbl As Double
  TotalAssessmentDbl As Double
  TotalDeductionsDbl As Double
  TotalBonusDbl As Double

  CBOO_FORMULAS As Boolean
  CSTR_DESTINY As String

  CBOO_MN As Boolean
  CDBL_TC As Double

  CBOO_RESUMEDEXPORTS As Boolean
  CBOO_RESUMEDRM As Boolean
  CBOO_RESUMEDRETENTION As Boolean
  CBOO_RESUMEDRETENTIONOTHER As Boolean
  CBOO_RESUMEDDISCOUNTS As Boolean
  CBOO_RESUMEDBONUS As Boolean

  CSTR_LOCAL_CURRENCY As String
  CSTR_LOCAL_CURRENCIES As String
  CSTR_LOCAL_SYMBOL_CURRENCY As String
  CSTR_FOREIGN_CURRENCY As String
  CSTR_FOREIGN_CURRENCIES As String
  CSTR_FOREIGN_SYMBOL_CURRENCY As String

  TBL_PURCNC As Variant
  TBL_PURCNC_CONTENTS As Variant
  TBL_PURCNC_PENALTIES As Variant
  TBL_PURCNC_EXPORTS As Variant
  TBL_PUR_RM As Variant
  TBL_TMP_RETENTION As Variant
  TBL_TMP_RETENTION_OTHER As Variant
  TBL_TMP_DISCOUNTS As Variant
  TBL_TMP_BONUS As Variant
End Type





' PROPERTIES ----------------------------------------------------------------------- '
' ABSTRACT INTERFACE PROPERTIES IMPLEMENTATION ------------------------------------- '
Public Property Set ClassEqlControllerInterface_ResourcesApp(ByRef RHS As ClassEqlResources)

  Set RES = RHS

End Property





' CONSTRUCTOR/DESTRUCTOR ----------------------------------------------------------- '
Private Sub Class_Terminate()

  Set MDL = Nothing
  Set RES = Nothing

End Sub





' CORE METHODS PRIVATE ------------------------------------------------------------- '
Private Function Purchasecnc_core_main(ByVal PVarRequest As Variant, ByRef PVarResponse As Variant) As Boolean

  Dim PRC As ClassEqlControllerHelper
  Dim VecAux0 As Variant
  Dim BooAux0 As Boolean
  Dim TypCnc As GLOTYPPURCHASE

  Const sFN As String = "CNC::-core_main"

  On Error GoTo EH
  Let Purchasecnc_core_main = False

  ' Request: 0.BooPage, 1.BooNewSheet, 2.RngBox, 3.SIMULATION
  ' Get currencies values
  Set MDL = New ClassEqlModel
  Set MDL.ResourcesApp = RES
  With TypCnc
    Let .CBOO_RESUMEDRM = MDL.PurRMDetailed
    Let .CBOO_RESUMEDRETENTION = MDL.PurRetentions
    Let .CBOO_RESUMEDRETENTIONOTHER = MDL.PurRetentionsOther
    Let .CBOO_RESUMEDDISCOUNTS = MDL.PurDiscounts
    Let .CBOO_RESUMEDBONUS = MDL.PurBonus
    Let .CBOO_RESUMEDEXPORTS = MDL.PurExports
    Let .CBOO_FORMULAS = MDL.Formulas

    Let VecAux0 = MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigcurrencies_sys, eqlMdlArray, "currency, currencies", MStrWhere:="currency_foreign=1")
    Let .CSTR_FOREIGN_CURRENCY = VBA.CStr(VecAux0(0, 0))
    Let .CSTR_FOREIGN_CURRENCIES = VBA.CStr(VecAux0(0, 1))
    Let .CSTR_FOREIGN_SYMBOL_CURRENCY = MDL.CurrencyMainSymbolForeign
    Let VecAux0 = Empty
    Let VecAux0 = MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigcurrencies_sys, eqlMdlArray, "currency, currencies", MStrWhere:="currency_main=1")
    Let .CSTR_LOCAL_CURRENCY = VBA.CStr(VecAux0(0, 0))
    Let .CSTR_LOCAL_CURRENCIES = VBA.CStr(VecAux0(0, 1))
    Let .CSTR_LOCAL_SYMBOL_CURRENCY = MDL.CurrencyMainSymbolLocal
    Let VecAux0 = Empty

    Let .TBL_PUR_RM = MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigpurrm_sys, eqlMdlArray)
    Let .TBL_PURCNC = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetPur, eqlMdlTblpurchasecnc_pur, eqlMdlArray), eqlRes1D)
    Let .TBL_PURCNC_CONTENTS = MDL.MGet(eqlMdlSheetPur, eqlMdlTblpurchasecnccontents_pur, eqlMdlArray)
    Let .TBL_PURCNC_EXPORTS = MDL.MGet(eqlMdlSheetPur, eqlMdlTblpurchasecncdedexp_pur, eqlMdlArray)
    Let .TBL_PURCNC_PENALTIES = MDL.MGet(eqlMdlSheetPur, eqlMdlTblpurchasecncpenalties_pur, eqlMdlArray)
    Let .TBL_TMP_BONUS = MDL.MGet(eqlMdlSheetTmp, eqlMdlTblPurbonus_tmp, eqlMdlArray)
    Let .TBL_TMP_DISCOUNTS = MDL.MGet(eqlMdlSheetTmp, eqlMdlTblPurdiscounts_tmp, eqlMdlArray)
    Let .TBL_TMP_RETENTION = MDL.MGet(eqlMdlSheetTmp, eqlMdlTblConfigpurretentions_tmp, eqlMdlArray)
    Let .TBL_TMP_RETENTION_OTHER = MDL.MGet(eqlMdlSheetTmp, eqlMdlTblConfigpurretentionsother_tmp, eqlMdlArray)

    Let .CDBL_TC = VBA.CDbl(.TBL_PURCNC(4))
    Let .CBOO_MN = VBA.CBool(.TBL_PURCNC(5))

    If VBA.CBool(PVarRequest(3)) Then
      Let .CSTR_DESTINY = MDL.PurDestination 'VBA.CStr(VecAux0(0))
    Else
      Let .CSTR_DESTINY = VBA.UCase(VBA.Replace(MDL.PurDestination, "SIMULACIÓN;", VBA.vbNullString))
    End If

  End With
  Set MDL = Nothing

  ' Title, purchase and provider datas
  If Not Purchasecnc_core_generate_header(PVarRequest, TypCnc) Then GoTo EH

  ' Cnc data to ore weight
  If Not Purchasecnc_core_generate_content(PVarRequest, TypCnc) Then GoTo EH

  ' Return
  With TypCnc
    ' Free memory
    Let .TBL_PUR_RM = Empty
    Let .TBL_PURCNC = Empty
    Let .TBL_PURCNC_CONTENTS = Empty
    Let .TBL_PURCNC_EXPORTS = Empty
    Let .TBL_PURCNC_PENALTIES = Empty
    Let .TBL_TMP_BONUS = Empty
    Let .TBL_TMP_DISCOUNTS = Empty
    Let .TBL_TMP_RETENTION = Empty
    Let .TBL_TMP_RETENTION_OTHER = Empty
    ' Create the purchase cnc
    Set PRC = New ClassEqlControllerHelper
    Let PRC.Title = .PTitle.Vector: Let .PTitle.Vector = Empty
    Let PRC.Invoice = .PInvoice.Vector: Let .PInvoice.Vector = Empty
    Let PRC.Provider = .PProvider.Vector: Let .PProvider.Vector = Empty
    Let PRC.WeightCNC = .PWeight.Vector: Let .PWeight.Vector = Empty
    Let PRC.ContentsCNC = .PContents.Vector: Let .PContents.Vector = Empty
    Let PRC.AssessmentCNC = .PAssessmentMain.Vector: Let .PAssessmentMain.Vector = Empty
    Let PRC.Deductions = .PDeductions.Vector: Let .PDeductions.Vector = Empty
    Let PRC.Bonus = .PBonus.Vector: Let .PBonus.Vector = Empty
    Let PRC.Final = .PFinal.Vector: Let .PFinal.Vector = Empty
    Let PRC.UnitsCtzCNC = .UnitsCtzVec: Let .UnitsCtzVec = Empty
    Let PRC.UnitsGraCNC = .UnitsGraVec: Let .UnitsGraVec = Empty
    Let PRC.UnitsTraCNC = .UnitsTraVec: Let .UnitsTraVec = Empty
    Let PRC.BooDeductionsCNC = .PDeductions.Exists
    Let PRC.BooBonus = .PBonus.Exists
    Let PVarResponse = Array(PRC)
    Set PRC = Nothing
  End With

  Let Purchasecnc_core_main = True

EH:
  Let VecAux0 = Empty
  Set PRC = Nothing
  Set MDL = Nothing
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Purchasecnc_core_generate_header( _
  ByRef PVarRequest As Variant, _
  ByRef PTypCnc As GLOTYPPURCHASE) As Boolean

  Dim xRows As Byte
  Dim BooParticular As Boolean

  Const sFN As String = "CNC::-core_generate_header"

  ' Request: 0.BooPage, 1.BooNewSheet, 2.RngBox
  On Error GoTo EH
  Let Purchasecnc_core_generate_header = False

  ' Title
  With PTypCnc.PTitle
    ReDim .Vector(0 To 1, 0 To 0)
    Let .Vector(0, 0) = "LIQUIDACIÓN DE CONCENTRADOS MINERALES"
    Let .Vector(1, 0) = "(EN " & VBA.UCase(PTypCnc.CSTR_FOREIGN_CURRENCIES) & ")"
  End With

  ' Purchase data
  With PTypCnc.PInvoice
    ReDim .Vector(0 To 6, 0 To 3)
    Let .Vector(0, 0) = "DATOS LIQUIDACIÓN"
    Let .Vector(1, 0) = "Destino:": Let .Vector(1, 1) = PTypCnc.CSTR_DESTINY
    Let .Vector(2, 0) = "Liquidación:": Let .Vector(2, 1) = VBA.CStr(PTypCnc.TBL_PURCNC(0))
    Let .Vector(3, 0) = "Lote:": Let .Vector(3, 1) = VBA.CStr(PTypCnc.TBL_PURCNC(1))
    Let .Vector(4, 0) = "Fecha:": Let .Vector(4, 1) = VBA.CDate(VBA.Format$(VBA.CStr(PTypCnc.TBL_PURCNC(2)), "dd/mm/yy"))
    Let .Vector(5, 0) = "Lugar:": Let .Vector(5, 1) = VBA.CStr(PTypCnc.TBL_PURCNC(3))
    Let .Vector(6, 0) = "T/C:": Let .Vector(6, 1) = VBA.Replace(VBA.CStr(PTypCnc.TBL_PURCNC(4)), ",", ".")
    Let .Vector(6, 2) = PTypCnc.CSTR_LOCAL_SYMBOL_CURRENCY & "/" & PTypCnc.CSTR_FOREIGN_SYMBOL_CURRENCY
  End With

  ' Provider data
  With PTypCnc.PProvider
    Let BooParticular = (VBA.LCase(VBA.CStr(PTypCnc.TBL_PURCNC(8))) = "particular")
    Let xRows = 9 + IIf(BooParticular, -1, 0)
    ReDim .Vector(0 To xRows, 0 To 3)
    Let .Vector(0, 0) = "DATOS PROVEDOR"
    Let .Vector(1, 0) = "Provedor:": Let .Vector(1, 1) = VBA.CStr(PTypCnc.TBL_PURCNC(6))
    Let .Vector(2, 0) = "Código:": Let .Vector(2, 1) = VBA.CStr(PTypCnc.TBL_PURCNC(7))
    Let .Vector(3, 0) = "Tipo:": Let .Vector(3, 1) = VBA.CStr(PTypCnc.TBL_PURCNC(8))
    Let .Vector(4, 0) = "C.I.:": Let .Vector(4, 1) = VBA.CStr(PTypCnc.TBL_PURCNC(9))
    Let .Vector(5, 0) = "NIT:": Let .Vector(5, 1) = VBA.CStr(PTypCnc.TBL_PURCNC(10))
    Let .Vector(6, 0) = "NIM:": Let .Vector(6, 1) = VBA.CStr(PTypCnc.TBL_PURCNC(11))
    If BooParticular = False Then Let .Vector(7, 0) = "Cooperativa:": Let .Vector(7, 1) = VBA.CStr(PTypCnc.TBL_PURCNC(12))
    Let .Vector(IIf(BooParticular = False, 8, 7), 0) = "Mina:": Let .Vector(IIf(BooParticular = False, 8, 7), 1) = VBA.CStr(PTypCnc.TBL_PURCNC(13))
    Let .Vector(IIf(BooParticular = False, 9, 8), 0) = "Procedencia:": Let .Vector(IIf(BooParticular = False, 9, 8), 1) = VBA.CStr(PTypCnc.TBL_PURCNC(14))
  End With

  ' Return
  Let Purchasecnc_core_generate_header = True

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Purchasecnc_core_generate_content( _
  ByRef PVarRequest As Variant, _
  ByRef PTypResponse As GLOTYPPURCHASE) As Boolean

  Dim RngBox As Range

  Const sFN As String = "CNC::-core_generate_content"

  ' Request: 0.BooPage, 1.BooNewSheet, 2.RngBox
  On Error GoTo EH
  Let Purchasecnc_core_generate_content = False

  ' Unpack data
  If PTypResponse.CBOO_FORMULAS Then Set RngBox = PVarRequest(2)

  With RES.ErrorHandler

    ' Cnc Data
    'Let PTypResponse.PLngRowCounter = 3
    If Not Purchasecnc_core_generate_content_weight(RngBox, PTypResponse) Then Call .ErrorHandlerSet(1, "Error en la preparación de Datos del concentrado mineral."): GoTo EH

    ' Contents
    If Not Purchasecnc_core_generate_content_contents(RngBox, PTypResponse) Then Call .ErrorHandlerSet(2, "Error en la determinación de contenidos."): GoTo EH

    ' Assessment
    If Not Purchasecnc_core_generate_content_assessment(RngBox, PTypResponse) Then Call .ErrorHandlerSet(3, "Error en la valoración." & VBA.vbNewLine & .ErrorMsg): GoTo EH

    ' Deductions
    If Not Purchasecnc_core_generate_content_deductions(RngBox, PTypResponse) Then Call .ErrorHandlerSet(4, "Error en proceso de deducciones." & VBA.vbNewLine & .ErrorMsg): GoTo EH

    ' Bonus
    If Not Purchasecnc_core_generate_content_bonus(RngBox, PTypResponse) Then Call .ErrorHandlerSet(5, "Error en el cálculo de bonos."): GoTo EH

    ' Final purchase
    If Not Purchasecnc_core_generate_content_finalpurchase(RngBox, PTypResponse) Then Call .ErrorHandlerSet(6, "Error en el cálculo final de la Liquidación de Concentrados Minerales."): GoTo EH

  End With

  ' Replace commas for dots
  With PTypResponse

    If Not PTypResponse.CBOO_FORMULAS Then
      If .PWeight.Exists Then Call Purchasecnc_core_assets_srcarray_to_dstarray(.PWeight.Vector)
      If .PContents.Exists Then Call Purchasecnc_core_assets_srcarray_to_dstarray(.PContents.Vector)
      If .PAssessmentMain.Exists Then Call Purchasecnc_core_assets_srcarray_to_dstarray(.PAssessmentMain.Vector)
      If .PDeductions.Exists Then Call Purchasecnc_core_assets_srcarray_to_dstarray(.PDeductions.Vector)
      If .PBonus.Exists Then Call Purchasecnc_core_assets_srcarray_to_dstarray(.PBonus.Vector)
      If .PFinal.Exists Then Call Purchasecnc_core_assets_srcarray_to_dstarray(.PFinal.Vector)
    End If

  End With

  ' Return
  Let Purchasecnc_core_generate_content = True

EH:
  Set RngBox = Nothing
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Purchasecnc_core_generate_content_weight( _
  ByVal PRngBox As Range, _
  ByRef PTypRes As GLOTYPPURCHASE) As Boolean

  Dim DatIncoming As Date
  Dim DblH2O As Double, DblDecrease As Double, DblPBS As Double
  Dim StrPBS As String, StrPNS As String
  Dim xCols As Long

  Const sFN As String = "CNC::-core_generate_content_weight"

  On Error GoTo EH
  Let Purchasecnc_core_generate_content_weight = False

  With PTypRes.PWeight

    Let xCols = 5

    ' Get data
    Let DatIncoming = VBA.CDate(PTypRes.TBL_PURCNC(15))
    Let PTypRes.TotalPBHDbl = VBA.CDbl(PTypRes.TBL_PURCNC(16))
    Let DblH2O = VBA.CDbl(PTypRes.TBL_PURCNC(17))
    Let DblDecrease = VBA.CDbl(PTypRes.TBL_PURCNC(18))

    ' Fill main vector with values
    ReDim .Vector(0 To 0, 0 To xCols)
    Let .Vector(0, 0) = "DATOS CONCENTRADO"
    Let .Vector = RES.ArrayAddAtLast(.Vector, Array("Ingreso", "PBH [T]", "Humedad [%]", "PBS [T]", "MERMA [%]", "PNS [T]"))

    If PTypRes.CBOO_FORMULAS Then
      Let StrPBS = "=" & PRngBox.Range("G6").Address(False, False) & "*(1-(" & PRngBox.Range("H6").Address(False, False) & "/100))" ' PBS
      Let StrPNS = "=" & PRngBox.Range("I6").Address(False, False) & "*(1-(" & PRngBox.Range("J6").Address(False, False) & "/100))"   ' PNS
      Let .Vector = RES.ArrayAddAtLast(.Vector, Array(VBA.Format(VBA.CDate(DatIncoming), "dd/mm/yy"), VBA.Replace(PTypRes.TotalPBHDbl, ",", "."), VBA.Replace(DblH2O, ",", "."), StrPBS, VBA.Replace(DblDecrease, ",", "."), StrPNS))
      Let .StrRngA1 = "F6"
    Else
      Let DblPBS = (PTypRes.TotalPBHDbl * (1 - (DblH2O / 100))) ' PBS
      Let PTypRes.TotalPBSDbl = (DblPBS * (1 - (DblDecrease / 100))) ' PNS
      Let .Vector = RES.ArrayAddAtLast(.Vector, Array(VBA.CDate(DatIncoming), PTypRes.TotalPBHDbl, DblH2O, DblPBS, DblDecrease, PTypRes.TotalPBSDbl))
    End If

    If Not VBA.IsArray(.Vector) Then GoTo EH

  End With

  ' Return
  Let Purchasecnc_core_generate_content_weight = True

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Purchasecnc_core_generate_content_contents( _
  ByVal PRngBox As Range, _
  ByRef PTypRes As GLOTYPPURCHASE) As Boolean

  Dim VecAux0 As Variant, VecAux1 As Variant
  Dim StrCont As String, StrGrad As String, StrCotz As String, StrAux0 As String
  Dim DblAux0 As Double
  Dim i As Long, j As Long, xCols As Long

  Const sFN As String = "CNC::-core_generate_content_contents"

  On Error GoTo EH
  Let Purchasecnc_core_generate_content_contents = False

  With PTypRes

    ' Get data
    ReDim .UnitsGraVec(UBound(.TBL_PURCNC_CONTENTS)) ' Unit Gra
    ReDim .UnitsCtzVec(UBound(.TBL_PURCNC_CONTENTS)) ' Unit Ctz

    ' Fill matrix
    With .PContents

      Let xCols = 5
      Let .LngRngA1 = 9 + UBound(PTypRes.TBL_PURCNC_CONTENTS)
      ReDim .Vector(0 To 0, 0 To xCols)
      Let .Vector(0, 0) = "DATOS DE CONTENIDOS EN CONCENTRADO"
      Let .Vector = RES.ArrayAddAtLast(.Vector, Array("Contenidos", Empty, Empty, Empty, "Ley", "Cotización"))

      For i = 0 To UBound(PTypRes.TBL_PURCNC_CONTENTS)
        Let j = RES.ArrayGetValuePosition(PTypRes.TBL_PUR_RM, VBA.CStr(PTypRes.TBL_PURCNC_CONTENTS(i, 0)), 0)
        Let PTypRes.UnitsGraVec(i) = PTypRes.TBL_PUR_RM(j, 6) ' Unit Gra
        Let PTypRes.UnitsCtzVec(i) = VBA.Replace(VBA.CStr(PTypRes.TBL_PUR_RM(j, 2)), ".", VBA.vbNullString) ' Unit Ctz
        Let StrAux0 = "[" & PTypRes.TBL_PURCNC_CONTENTS(i, 5) & "]"
        Let DblAux0 = VBA.CDbl(VBA.Replace(VBA.CStr(PTypRes.TBL_PURCNC_CONTENTS(i, 1)), StrAux0, VBA.vbNullString))
        Let .Vector = RES.ArrayAddAtLast(.Vector, Array(VBA.StrConv(VBA.CStr(PTypRes.TBL_PUR_RM(j, 0)), vbProperCase), Empty, Empty, Empty, DblAux0, PTypRes.TBL_PUR_RM(j, 3)))
      Next i
      Let VecAux1 = Empty

      If PTypRes.CBOO_FORMULAS Then Let .StrRngA1 = "F9"

    End With

  End With

  ' Return
  Let Purchasecnc_core_generate_content_contents = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Purchasecnc_core_generate_content_assessment(ByVal PRngBox As Range, ByRef PTypRes As GLOTYPPURCHASE) As Boolean

  Dim BooAssessmetsExists As Boolean
  Dim DblTotalPerTon As Double
  Dim i As Long, j As Long, k As Long, xRows As Long, xCols As Long, xRowFirstVal As Long, xRowResize As Long, xRowTotals As Long, LngResize As Long

  Const sFN As String = "CNC::-core_generate_content_assessment"

  On Error GoTo EH
  Let Purchasecnc_core_generate_content_assessment = False

  ' Get all values from assessment methods
  With RES.ErrorHandler

    Let PTypRes.TotalAssessmentDbl = 0

    ' Assessment
    If Not Purchasecnc_core_generate_content_assessment_value(PRngBox, PTypRes) Then Call .ErrorHandlerSet(3.1, "Falló el cálculo de valoración del concentrado."): GoTo EH

    ' Treatment cost
    If Not Purchasecnc_core_generate_content_assessment_treatmentcost(PRngBox, PTypRes) Then Call .ErrorHandlerSet(3.2, "Falló el cálculo del costo de tratamiento."): GoTo EH

    ' Refinement
    If Not Purchasecnc_core_generate_content_assessment_refinement(PRngBox, PTypRes) Then Call .ErrorHandlerSet(3.3, "Falló el cálculo del costo de refinación."): GoTo EH

    ' Penalties
    If Not Purchasecnc_core_generate_content_assessment_penalties(PRngBox, PTypRes) Then Call .ErrorHandlerSet(3.4, "Falló el cálculo de penalidades."): GoTo EH

  End With

  ' Fill Assessment matrix with results
  With PTypRes

    ' Prepare main assessment vector
    Let .PAssessmentMain.Vector = RES.ArrayConcat(.PAssessmentMain.Vector, .PAssessment.Vector)
    Let LngResize = UBound(.PAssessmentMain.Vector) + 2 ' Just Assessment value

    ' Determine if other assessments exists
    Let BooAssessmetsExists = (.PTreatmentCost.Exists Or .PRefinement.Exists Or .PPenalties.Exists)
    If BooAssessmetsExists Then ' Add contents
      If .PTreatmentCost.Exists Then Let .PAssessmentMain.Vector = RES.ArrayConcat(.PAssessmentMain.Vector, .PTreatmentCost.Vector)
      If .PRefinement.Exists Then Let .PAssessmentMain.Vector = RES.ArrayConcat(.PAssessmentMain.Vector, .PRefinement.Vector)
      If .PPenalties.Exists Then Let .PAssessmentMain.Vector = RES.ArrayConcat(.PAssessmentMain.Vector, .PPenalties.Vector)
      Let LngResize = (UBound(.PAssessmentMain.Vector) + 1) - LngResize ' Assessment value
    End If

    ' Add Totals
    With .PAssessmentMain

      If PTypRes.CBOO_FORMULAS Then
        Dim StrPerTonVal As String, StrPerTonRng As String
        Dim LngFirstA1 As Long
        Let LngFirstA1 = 9 + (UBound(PTypRes.PContents.Vector) - 2) + (UBound(PTypRes.PAssessment.Vector) + 1)  ' First Main Assessment row address
        Let StrPerTonRng = "F" & LngFirstA1 + IIf(LngResize <= 0, 1, LngResize + 2) ' Main Assessment per ton address
        Let StrPerTonVal = "=" & PRngBox.Range(PTypRes.PAssessment.StrRngA1).Offset(UBound(PTypRes.PAssessment.Vector) - 2, 5).Address(False, False)
        If LngResize > 0 Then Let StrPerTonVal = StrPerTonVal & "-SUM(" & PRngBox.Range("F" & LngFirstA1).Offset(2, 5).Resize(LngResize, 1).Address(False, False) & ")"
        Let .Vector = RES.ArrayConcat(.Vector, Array("Total valor bruto por tonelada:", Empty, Empty, Empty, Empty, StrPerTonVal))
        Let StrPerTonVal = "=" & PRngBox.Range(StrPerTonRng).Offset(0, 5).Address(False, False)
        Let StrPerTonVal = StrPerTonVal & "*" & PRngBox.Range("K6").Address(False, False)
        Let .Vector = RES.ArrayConcat(.Vector, Array("TOTAL VALOR BRUTO [" & PTypRes.CSTR_FOREIGN_SYMBOL_CURRENCY & "]:", Empty, Empty, Empty, Empty, StrPerTonVal))
      Else
        Let DblTotalPerTon = PTypRes.TotalAssessmentDbl
        Let PTypRes.TotalAssessmentDbl = PTypRes.TotalAssessmentDbl * VBA.CDbl(PTypRes.PWeight.Vector(2, 5)) ' Main Assessment Value
        Let .Vector = RES.ArrayConcat(.Vector, Array("Total valor bruto por tonelada:", Empty, Empty, Empty, Empty, DblTotalPerTon))
        Let .Vector = RES.ArrayConcat(.Vector, Array("TOTAL VALOR BRUTO [" & PTypRes.CSTR_FOREIGN_SYMBOL_CURRENCY & "]:", Empty, Empty, Empty, Empty, PTypRes.TotalAssessmentDbl))
      End If
      Let .LngRngA1 = (9 + (UBound(PTypRes.PContents.Vector) - 2)) + (UBound(.Vector) + 1)

    End With
    Let .PAssessment.Vector = Empty
    Let .PTreatmentCost.Vector = Empty
    Let .PRefinement.Vector = Empty
    Let .PPenalties.Vector = Empty
    If Not VBA.IsArray(.PAssessmentMain.Vector) Then GoTo EH
    Let .TotalAssessmentRng = "K" & (9 + (UBound(.PContents.Vector) - 2)) + (UBound(.PAssessmentMain.Vector) + 1) ' Main Assessment address
    Let .PAssessmentMain.StrRngA1 = .TotalAssessmentRng

  End With

  ' Return
  Let Purchasecnc_core_generate_content_assessment = True

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Purchasecnc_core_generate_content_assessment_value(ByVal PRngBox As Range, ByRef PTypRes As GLOTYPPURCHASE) As Boolean

  Dim RngBox As Range
  Dim StrAux0 As String, StrAux1 As String, StrGrade As String, StrCotiz As String, StrValue As String, StrGradeC As String, StrCotizC As String
  Dim DblAux0 As Double, DblAux1 As Double, DblGrade As Double, DblCotiz As Double, DblValue As Double
  Dim i As Long, j As Long, xCols As Long, xRowResize As Long

  Const sFN As String = "CNC::-core_generate_content_assessment_value"

  On Error GoTo EH
  Let Purchasecnc_core_generate_content_assessment_value = False

  With PTypRes

    ReDim .UnitsTraVec(UBound(.TBL_PURCNC_CONTENTS))

    ' Calculate values
    With .PAssessment

      Let xCols = 5
      Let xRowResize = UBound(PTypRes.TBL_PURCNC_CONTENTS) + 1
      Let .LngRngA1 = PTypRes.PContents.LngRngA1 + 3
      Let .StrRngA1 = "F" & .LngRngA1
      If PTypRes.CBOO_FORMULAS Then Set RngBox = PRngBox.Range(.StrRngA1)
      Let .LngRngA1 = .LngRngA1 + xRowResize

      ReDim .Vector(0 To 0, 0 To xCols)
      Let .Vector(0, 0) = "VALORACIÓN"
      Let .Vector(0, 5) = "Valor [" & PTypRes.CSTR_FOREIGN_SYMBOL_CURRENCY & "/T]"
      Let .Vector = RES.ArrayAddAtLast(.Vector, Array("Mineral", "Ley", "Deducción", "Pagable [%]", "Cotización", "Valor"))

      For i = 0 To UBound(PTypRes.TBL_PURCNC_CONTENTS)

        Let j = RES.ArrayGetValuePosition(PTypRes.TBL_PUR_RM, VBA.CStr(PTypRes.TBL_PURCNC_CONTENTS(i, 0)), 0)

        Let StrAux0 = VBA.CStr(PTypRes.TBL_PURCNC_CONTENTS(i, 5))
        Let PTypRes.UnitsTraVec(i) = IIf(VBA.LCase(StrAux0) = "dm", "OT", StrAux0) ' Set transaction units
        If PTypRes.CBOO_FORMULAS Then
          Let StrGrade = "=" & PRngBox.Range(PTypRes.PContents.StrRngA1).Offset(i, 4).Address(False, False) & "*" & VBA.Replace(VBA.CStr(PTypRes.TBL_PUR_RM(j, 9)), ",", ".") ' 1.  Conversion of mineral units
          Let StrCotiz = "=" & PRngBox.Range(PTypRes.PContents.StrRngA1).Offset(i, 5).Address(False, False) & "*" & VBA.Replace(VBA.CStr(PTypRes.TBL_PUR_RM(j, 8)), ",", ".") ' 2.  Conversion of prices units
          ' 3.  Calculate final value
          Let StrAux0 = RngBox.Offset(i, 2).Address(False, False) ' Deduction
          Let StrAux1 = RngBox.Offset(i, 3).Address(False, False) ' %Payable
          Let StrGradeC = RngBox.Offset(i, 1).Address(False, False) ' Grade
          Let StrCotizC = RngBox.Offset(i, 4).Address(False, False) ' Cotiz
          If VBA.InStr(VBA.LCase(VBA.CStr(PTypRes.TBL_PUR_RM(j, 7))), "ot") > 0 Then ' =IF(Gra=0;Gra-ABS(Ded);IF(Ded=0;Gra*Pay/100;(Gra-ABS(Ded))))*(Ctz*Pay/100)
            Let StrValue = "=IF(" & StrGradeC & "=0," & StrGradeC & "-ABS(" & StrAux0 & "),IF(" & StrAux0 & "=0," & StrGradeC & "*" & StrAux1 & "*0.01,(" & StrGradeC & "-ABS(" & StrAux0 & ")))*" & StrAux1 & ")*(" & StrCotizC & "*0.01)"
          Else ' =IF(OR(Gra=0;Gra>=100);Gra-ABS(Ded);IF(Ded=0;Gra*Pay/100;IF(Gra>(Ded/((Pay/100)-1));Gra*Pay/100;Gra-ABS(Ded))))*(0,01)*(Ctz)
            Let StrValue = "=IF(OR(" & StrGradeC & "=0," & StrGradeC & ">=100)," & StrGradeC & "-ABS(" & StrAux0 & "),IF(" & StrAux0 & "=0," & StrGradeC & "*" & StrAux1 & "*0.01,IF(" & StrGradeC & ">(" & StrAux0 & "/((" & StrAux1 & "*0.01)-1))," & StrGradeC & "*" & StrAux1 & "*0.01," & StrGradeC & "-ABS(" & StrAux0 & "))))*(" & StrCotizC & "*0.01)"
          End If
          Let StrAux0 = VBA.Replace(VBA.CStr(PTypRes.TBL_PURCNC_CONTENTS(i, 2)), ",", ".") 'RngBox.Offset(i, 2).Address(False, False) ' Deduction
          Let StrAux1 = VBA.Replace(VBA.CStr(PTypRes.TBL_PURCNC_CONTENTS(i, 4)), ",", ".") 'RngBox.Offset(i, 3).Address(False, False) ' %Payable
          Let .Vector = RES.ArrayAddAtLast(.Vector, Array(VBA.StrConv(VBA.CStr(PTypRes.TBL_PUR_RM(j, 1)), vbProperCase), StrGrade, StrAux0, StrAux1, StrCotiz, StrValue))
        Else
          Let StrAux0 = VBA.CStr(PTypRes.TBL_PURCNC_CONTENTS(i, 1))
          Let StrAux1 = "[" & VBA.CStr(PTypRes.TBL_PURCNC_CONTENTS(i, 5)) & "]"
          Let StrAux0 = VBA.Replace(StrAux0, StrAux1, VBA.vbNullString)
          Let DblGrade = VBA.CDbl(StrAux0) * VBA.CDbl(PTypRes.TBL_PUR_RM(j, 9)) ' 1. Mineral conversion units
          Let DblCotiz = VBA.CDbl(PTypRes.TBL_PUR_RM(j, 3)) * VBA.CDbl(PTypRes.TBL_PUR_RM(j, 8)) ' 2. Units conversion of prices
          ' 3.  Calculate final value
          Let DblAux0 = VBA.CDbl(PTypRes.TBL_PURCNC_CONTENTS(i, 2)) ' Deduction
          Let DblAux1 = VBA.CDbl(PTypRes.TBL_PURCNC_CONTENTS(i, 4)) ' %Payable
          If VBA.LCase(VBA.CStr(PTypRes.TBL_PUR_RM(j, 7))) = "ot" Then ' =IF(Gra=0;Gra-ABS(Ded);IF(Ded=0;Gra*Pay/100;(Gra-ABS(Ded))))*(Ctz*Pay/100)
            Let DblValue = IIf(DblGrade = 0, DblGrade - VBA.Abs(DblAux0), IIf(DblAux0 = 0, DblGrade * DblAux1 * 0.01, (DblGrade - VBA.Abs(DblAux0))) * DblAux1) * (DblCotiz * 0.01)
          Else ' =IF(OR(Gra=0;Gra>=100);Gra-ABS(Ded);IF(Ded=0;Gra*Pay/100;IF(Gra>(Ded/((Pay/100)-1));Gra*Pay/100;Gra-ABS(Ded))))*(0,01)*(Ctz)
            Let DblValue = IIf(DblGrade = 0 Or DblGrade >= 100, DblGrade - VBA.Abs(DblAux0), IIf(DblAux0 = 0, DblGrade * DblAux1 * 0.01, IIf(DblGrade > (DblAux0 / ((DblAux1 * 0.01) - 1)), DblGrade * DblAux1 * 0.01, DblGrade - VBA.Abs(DblAux0)))) * 0.01 * DblCotiz
          End If
          Let .Vector = RES.ArrayAddAtLast(.Vector, Array(VBA.StrConv(VBA.CStr(PTypRes.TBL_PUR_RM(j, 1)), vbProperCase), DblGrade, DblAux0, DblAux1, DblCotiz, DblValue))
          Let PTypRes.TotalAssessmentDbl = PTypRes.TotalAssessmentDbl + DblValue
        End If

      Next i
      If Not VBA.IsArray(.Vector) Then GoTo EH

      ' Totals row
      If PTypRes.CBOO_FORMULAS Then Let StrValue = "=SUM(" & RngBox.Offset(0, 5).Resize(xRowResize, 1).Address(False, False) & ")"
      Let .Vector = RES.ArrayAddAtLast(.Vector, Array("Total Valoración", Empty, Empty, Empty, Empty, IIf(PTypRes.CBOO_FORMULAS, StrValue, PTypRes.TotalAssessmentDbl)))

    End With

  End With

  ' Return
  Let Purchasecnc_core_generate_content_assessment_value = True

EH:
  Set RngBox = Nothing
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Purchasecnc_core_generate_content_assessment_treatmentcost( _
  ByVal PRngBox As Range, _
  ByRef PTypRes As GLOTYPPURCHASE) As Boolean

  Dim RngBox As Range

  Const sFN As String = "CNC::-core_generate_content_assessment_treatmentcost"

  On Error GoTo EH
  Let Purchasecnc_core_generate_content_assessment_treatmentcost = False

  With PTypRes

    ' Get data
    Let .PTreatmentCost.Exists = True

    ' Fill matrix
    With .PTreatmentCost

      Let .LngRngA1 = PTypRes.PAssessment.LngRngA1
      Dim xCols As Long
      Let xCols = 5
      Let .LngRngA1 = .LngRngA1 + 2
      Let .StrRngA1 = "F" & .LngRngA1
      If PTypRes.CBOO_FORMULAS Then Set RngBox = PRngBox.Range(.StrRngA1)
      Let .LngRngA1 = .LngRngA1 + 2

      ' Fill final vector with data
      ReDim .Vector(0 To 0, 0 To xCols)
      Let .Vector(0, 0) = "COSTO TRATAMIENTO"
      Let .Vector = RES.ArrayAddAtLast(.Vector, Array("Maquila", Empty, Empty, Empty, Empty, PTypRes.TBL_PURCNC(19)))
      Let .Vector = RES.ArrayAddAtLast(.Vector, Array("Ajuste", "Actual", "Base", "Diferencia", "Escalador", Empty))

      ' Calculate results
      If PTypRes.CBOO_FORMULAS Then
        Dim StrCur As String, StrDif As String, StrVal As String
        Let StrCur = "=" & PRngBox.Range(PTypRes.PAssessment.StrRngA1).Offset(0, 4).Address(False, False) ' Actual
        Let StrDif = "=" & RngBox.Offset(2, 1).Address(False, False) & "-" & RngBox.Offset(2, 2).Address(False, False) ' Diferencia
        Let StrVal = "=IF(" & RngBox.Offset(2, 3).Address(False, False) & ">0," & RngBox.Offset(2, 3).Address(False, False) & "*" & RngBox.Offset(2, 4).Address(False, False) & ",0)"
        Let .Vector = RES.ArrayAddAtLast(.Vector, Array(Empty, StrCur, VBA.Replace(VBA.CStr(PTypRes.TBL_PURCNC(20)), ",", "."), StrDif, VBA.Replace(VBA.CStr(PTypRes.TBL_PURCNC(21)), ",", "."), StrVal))
      Else
        Dim DblCur As Double, DblDif As Double, DblVal As Double
        Let DblCur = VBA.CDbl(PTypRes.PAssessment.Vector(2, 4)) ' Actual
        Let DblDif = (DblCur - VBA.CDbl(PTypRes.TBL_PURCNC(20))) ' Diferencia
        Let DblVal = IIf(DblDif > 0, DblDif * VBA.CDbl(PTypRes.TBL_PURCNC(21)), 0) ' Result
        Let .Vector = RES.ArrayAddAtLast(.Vector, Array(Empty, DblCur, PTypRes.TBL_PURCNC(20), DblDif, PTypRes.TBL_PURCNC(21), DblVal))
        Let PTypRes.TotalAssessmentDbl = PTypRes.TotalAssessmentDbl - (VBA.CDbl(PTypRes.TBL_PURCNC(19)) + DblVal)
      End If
      If Not VBA.IsArray(.Vector) Then GoTo EH

    End With

  End With

  ' Return
  Let Purchasecnc_core_generate_content_assessment_treatmentcost = True

EH:
  Set RngBox = Nothing
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Purchasecnc_core_generate_content_assessment_refinement( _
  ByVal PRngBox As Range, _
  ByRef PTypRes As GLOTYPPURCHASE) As Boolean

  Dim RngBox As Range
  Dim StrGra As String, StrDed As String, StrPay As String, StrVal As String, StrAux0 As String
  Dim DblGra As Double, DblDed As Double, DblPay As Double, DblVal As Double
  Dim i As Long, xRows As Long, xCols As Long

  Const sFN As String = "CNC::-core_generate_content_assessment_refinement"

  On Error GoTo EH
  Let Purchasecnc_core_generate_content_assessment_refinement = False

  With PTypRes

    ' Get data (0-CONCEPTO 1-BASE 2-CONDICIÓN [%])
    Let .PRefinement.Exists = False
    Let xRows = 0
    For i = LBound(.TBL_PURCNC_CONTENTS) To UBound(.TBL_PURCNC_CONTENTS)
      Let DblDed = VBA.CDbl(.TBL_PURCNC_CONTENTS(i, 7))
      Let DblPay = VBA.CDbl(.TBL_PURCNC_CONTENTS(i, 8))
      If Not (DblDed = 0 And DblPay = 0) Then
        Let .PRefinement.Exists = True
        Let xRows = xRows + 1
      End If
    Next i

    ' Fill matrix
    With .PRefinement

      Let .LngRngA1 = PTypRes.PTreatmentCost.LngRngA1

      If .Exists Then

        Let xCols = 5
        Let .LngRngA1 = .LngRngA1 + 3
        Let .StrRngA1 = "F" & .LngRngA1
        If PTypRes.CBOO_FORMULAS Then Set RngBox = PRngBox.Range(.StrRngA1)
        Let .LngRngA1 = .LngRngA1 + (xRows - 1) 'UBound(VecAux0)

        ' Fill final vector with data
        ReDim .Vector(0 To 0, 0 To xCols)
        Let .Vector(0, 0) = "COSTO REFINACIÓN"
        Let .Vector = RES.ArrayAddAtLast(.Vector, Array("Concepto", Empty, "Base", "Condición[%]", "Pagable", Empty))

        ' Calculate results
        For i = LBound(PTypRes.TBL_PURCNC_CONTENTS) To UBound(PTypRes.TBL_PURCNC_CONTENTS)

          Let DblDed = VBA.CDbl(PTypRes.TBL_PURCNC_CONTENTS(i, 7))
          Let DblPay = VBA.CDbl(PTypRes.TBL_PURCNC_CONTENTS(i, 8))
          If Not (DblDed = 0 And DblPay = 0) Then

            Let StrAux0 = VBA.LCase(VBA.CStr(PTypRes.UnitsGraVec(i)))

            If PTypRes.CBOO_FORMULAS Then

              Let StrGra = PRngBox.Range(PTypRes.PAssessment.StrRngA1).Offset(i, 1).Address(False, False) ' Get Mineral Grades
              Let StrDed = PRngBox.Range(PTypRes.PAssessment.StrRngA1).Offset(i, 2).Address(False, False) ' Get Deduction
              Let StrPay = PRngBox.Range(PTypRes.PAssessment.StrRngA1).Offset(i, 3).Address(False, False) ' Get %Payable
              ' Calculate final value
              If StrAux0 = "ot" Then ' =IF(Gra=0;Gra-ABS(Ded);IF(Ded=0;Gra*Pay/100;(Gra-ABS(Ded))))
                Let StrPay = "=IF(" & StrGra & "=0," & StrGra & "-ABS(" & StrDed & "),IF(" & StrDed & "=0," & StrGra & "*" & StrPay & "*0.01,(" & StrGra & "-ABS(" & StrDed & ")))*" & StrPay & ")"
              Else ' =IF(OR(Gra=0;Gra>=100);Gra-ABS(Ded);IF(Ded=0;Gra*Pay/100;IF(Gra>(Ded/((Pay/100)-1));Gra*Pay/100;Gra-ABS(Ded))))*(0,01*2204.6223)
                Let StrPay = "=IF(OR(" & StrGra & "=0," & StrGra & ">=100)," & StrGra & "-ABS(" & StrDed & "),IF(" & StrDed & "=0," & StrGra & "*" & StrPay & "*0.01,IF(" & StrGra & ">(" & StrDed & "/((" & StrPay & "*0.01)-1))," & StrGra & "*" & StrPay & "*0.01," & StrGra & "-ABS(" & StrDed & "))))*(" & IIf(VBA.LCase(VBA.CStr(PTypRes.UnitsGraVec(i))) = "%", "2204.6223", "1") & "*0.01)"
              End If
              ' =IF(CTZ*CND/100>BSE;PAY*CTZ*CND/100;PAY*BSE)
              Let StrVal = "=IF(" & _
                PRngBox.Range(PTypRes.PAssessment.StrRngA1).Offset(i, 4).Address(False, False) & "*" & _
                RngBox.Offset(i, 3).Address(False, False) & "*0.01>" & _
                RngBox.Offset(i, 2).Address(False, False) & "," & _
                PRngBox.Range(PTypRes.PAssessment.StrRngA1).Offset(i, 4).Address(False, False) & "*" & _
                RngBox.Offset(i, 3).Address(False, False) & "*" & _
                RngBox.Offset(i, 4).Address(False, False) & "*0.01," & _
                RngBox.Offset(i, 2).Address(False, False) & "*" & _
                RngBox.Offset(i, 4).Address(False, False) & ")"
              Let .Vector = RES.ArrayAddAtLast(.Vector, Array(VBA.StrConv(VBA.CStr(PTypRes.TBL_PURCNC_CONTENTS(i, 6)), vbProperCase), Empty, VBA.Replace(VBA.CStr(PTypRes.TBL_PURCNC_CONTENTS(i, 7)), ",", "."), VBA.Replace(VBA.CStr(PTypRes.TBL_PURCNC_CONTENTS(i, 8)), ",", "."), StrPay, StrVal))

            Else

              Let DblGra = VBA.CDbl(PTypRes.PAssessment.Vector(i + 2, 1)) ' Get Mineral Grades
              Let DblDed = VBA.CDbl(PTypRes.PAssessment.Vector(i + 2, 2)) ' Get Deduction
              Let DblPay = VBA.CDbl(PTypRes.PAssessment.Vector(i + 2, 3)) ' Get %Payable
              If StrAux0 = "ot" Then ' =IF(Gra=0;Gra-ABS(Ded);IF(Ded=0;Gra*Pay/100;(Gra-ABS(Ded))))
                Let DblPay = IIf(DblGra = 0, DblGra - VBA.Abs(StrDed), IIf(DblDed = 0, DblGra * DblPay * 0.01, (DblGra - Abs(DblDed))) * DblPay)
              Else ' =IF(OR(Gra=0;Gra>=100);Gra-ABS(Ded);IF(Ded=0;Gra*Pay/100;IF(Gra>(Ded/((Pay/100)-1));Gra*Pay/100;Gra-ABS(Ded))))*(0,01*2204.6223)
                Let DblPay = IIf(DblGra = 0 Or DblGra >= 100, DblGra - VBA.Abs(DblDed), IIf(DblDed = 0, DblGra * DblPay * 0.01, IIf(DblGra > (DblDed / ((DblPay * 0.01) - 1)), DblGra * DblPay * 0.01, DblGra - VBA.Abs(DblDed)))) * (IIf(StrAux0 = "%", 2204.6223, 1) * 0.01)
              End If
              ' =IF(CTZ*CND/100>BSE;PAY*CTZ*CND/100;PAY*BSE)
              Let DblVal = IIf(VBA.CDbl(PTypRes.PAssessment.Vector(i + 2, 4)) * VBA.CDbl(PTypRes.TBL_PURCNC_CONTENTS(i, 8)) * 0.01 > VBA.CDbl(PTypRes.TBL_PURCNC_CONTENTS(i, 7)), VBA.CDbl(PTypRes.PAssessment.Vector(i + 2, 4)) * VBA.CDbl(PTypRes.TBL_PURCNC_CONTENTS(i, 8)) * DblPay * 0.01, VBA.CDbl(PTypRes.TBL_PURCNC_CONTENTS(i, 7)) * DblPay)
              Let .Vector = RES.ArrayAddAtLast(.Vector, Array(VBA.StrConv(VBA.CStr(PTypRes.TBL_PURCNC_CONTENTS(i, 6)), vbProperCase), Empty, PTypRes.TBL_PURCNC_CONTENTS(i, 7), PTypRes.TBL_PURCNC_CONTENTS(i, 8), DblPay, DblVal))
              Let PTypRes.TotalAssessmentDbl = PTypRes.TotalAssessmentDbl - DblVal

            End If

          End If

        Next i
        If Not VBA.IsArray(.Vector) Then GoTo EH

      End If

    End With

  End With

  ' Return
  Let Purchasecnc_core_generate_content_assessment_refinement = True

EH:
  Set RngBox = Nothing
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Purchasecnc_core_generate_content_assessment_penalties( _
  ByVal PRngBox As Range, _
  ByRef PTypRes As GLOTYPPURCHASE) As Boolean

  Dim RngBox As Range
  Dim StrCur As String, StrBse As String, StrPay As String, StrLim As String, StrVal As String
  Dim DblCur As Double, DblBse As Double, DblPay As Double, DblLim As Double, DblVal As Double
  Dim i As Long, xCols As Long

  Const sFN As String = "CNC::-core_generate_content_assessment_penalties"

  On Error GoTo EH
  Let Purchasecnc_core_generate_content_assessment_penalties = False

  With PTypRes

    ' Get data (0-ELEMENTO 1-ACTUAL 2-BASE 3-DEDUCCIÓN 4-LÍMITE 5-SELECTED)
    Let .PPenalties.Exists = VBA.IsArray(.TBL_PURCNC_PENALTIES)

    ' Fill matrix
    With .PPenalties

      Let .LngRngA1 = PTypRes.PRefinement.LngRngA1
      If .Exists Then

        Let xCols = 5
        Let .LngRngA1 = .LngRngA1 + 3
        Let .StrRngA1 = "F" & .LngRngA1
        If PTypRes.CBOO_FORMULAS Then Set RngBox = PRngBox.Range(.StrRngA1)
        Let .LngRngA1 = .LngRngA1 + UBound(PTypRes.TBL_PURCNC_PENALTIES)

        ' Fill final vector with data
        ReDim .Vector(0 To 0, 0 To xCols)
        Let .Vector(0, 0) = "PENALIDADES"
        Let .Vector = RES.ArrayAddAtLast(.Vector, Array("Elemento", "Actual", "Base", "[" & PTypRes.CSTR_FOREIGN_SYMBOL_CURRENCY & "]", "Límite", Empty))

        ' Calculate results
        For i = 0 To UBound(PTypRes.TBL_PURCNC_PENALTIES)

          If PTypRes.CBOO_FORMULAS Then

            Let StrCur = RngBox.Offset(i, 1).Address(False, False) ' Current
            Let StrBse = RngBox.Offset(i, 2).Address(False, False) ' Base
            Let StrPay = RngBox.Offset(i, 3).Address(False, False) ' Pay
            Let StrLim = RngBox.Offset(i, 4).Address(False, False) ' Limit
            ' Calculate final value: =IF((Cur-Bse)>0;(Cur-Bse)*Pay/Lim;0)
            Let StrVal = "=IF(" & StrCur & "-" & StrBse & ">0," & "(" & StrCur & "-" & StrBse & ")*" & StrPay & "/" & StrLim & ",0)"
            Let .Vector = RES.ArrayAddAtLast(.Vector, Array(VBA.StrConv(VBA.CStr(PTypRes.TBL_PURCNC_PENALTIES(i, 0)), vbProperCase), VBA.Replace(VBA.CStr(PTypRes.TBL_PURCNC_PENALTIES(i, 1)), ",", "."), VBA.Replace(VBA.CStr(PTypRes.TBL_PURCNC_PENALTIES(i, 2)), ",", "."), VBA.Replace(VBA.CStr(PTypRes.TBL_PURCNC_PENALTIES(i, 3)), ",", "."), VBA.Replace(VBA.CStr(PTypRes.TBL_PURCNC_PENALTIES(i, 4)), ",", "."), StrVal))

          Else

            Let DblCur = VBA.CDbl(PTypRes.TBL_PURCNC_PENALTIES(i, 1)) ' Current
            Let DblBse = VBA.CDbl(PTypRes.TBL_PURCNC_PENALTIES(i, 2)) ' Base
            Let DblPay = VBA.CDbl(PTypRes.TBL_PURCNC_PENALTIES(i, 3)) ' Pay
            Let DblLim = VBA.CDbl(PTypRes.TBL_PURCNC_PENALTIES(i, 4)) ' Limit
            ' Calculate final value: =IF((Cur-Bse)>0;(Cur-Bse)*Pay/Lim;0)
            Let DblVal = IIf(DblCur - DblBse > 0, (DblCur - DblBse) * DblPay / DblLim, 0)
            Let .Vector = RES.ArrayAddAtLast(.Vector, Array(VBA.StrConv(VBA.CStr(PTypRes.TBL_PURCNC_PENALTIES(i, 0)), vbProperCase), DblCur, DblBse, DblPay, DblLim, DblVal))
            Let PTypRes.TotalAssessmentDbl = PTypRes.TotalAssessmentDbl - DblVal

          End If

        Next i
        If Not VBA.IsArray(.Vector) Then GoTo EH

      End If

    End With

  End With

  ' Return
  Let Purchasecnc_core_generate_content_assessment_penalties = True

EH:
  Set RngBox = Nothing
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Purchasecnc_core_generate_content_deductions( _
  ByVal PRngBox As Range, _
  ByRef PTypRes As GLOTYPPURCHASE) As Boolean

  Dim LngResize As Long

  Const sFN As String = "CNC::-core_generate_content_deductions"

  On Error GoTo EH
  Let Purchasecnc_core_generate_content_deductions = False

  ' Get all values from deductions methods
  With RES.ErrorHandler

    Let PTypRes.TotalDeductionsDbl = 0

    ' Exports discounts
    If Not Purchasecnc_core_generate_content_deductions_exports(PRngBox, PTypRes) Then Call .ErrorHandlerSet(4.1, "Falló el cálculo de deducciones de exportación."): GoTo EH

    ' RM Tax
    If Not Purchasecnc_core_generate_content_deductions_rmtaxes(PRngBox, PTypRes) Then Call .ErrorHandlerSet(4.2, "Falló el cálculo de regalías."): GoTo EH

    ' Taxes by Law
    If Not Purchasecnc_core_generate_content_deductions_lawtaxes(PRngBox, PTypRes) Then Call .ErrorHandlerSet(4.3, "Falló el cálculo de retenciones de ley."): GoTo EH

    ' Taxes others
    If Not Purchasecnc_core_generate_content_deductions_othertaxes(PRngBox, PTypRes) Then Call .ErrorHandlerSet(4.4, "Falló el cálculo en retenciones otras."): GoTo EH

    ' Discounts others
    If Not Purchasecnc_core_generate_content_deductions_discounts(PRngBox, PTypRes) Then Call .ErrorHandlerSet(4.5, "Falló el cálculo en descuentos otros."): GoTo EH

  End With

  ' Fill deductions matrix with results
  With PTypRes

    Let .PDeductions.LngRngA1 = 0
    Let .PDeductions.Exists = (.PDExports.Exists Or .PDRMRets.Exists Or .PDLawRets.Exists Or .PDOthRets.Exists Or .PDDiscnts.Exists)
    If .PDeductions.Exists Then
      ' Create Deductions vector
      ReDim .PDeductions.Vector(0 To 0, 0 To 5)
      Let .PDeductions.Vector(0, 0) = "DEDUCCIONES"
      Let .PDeductions.Vector = RES.ArrayAddAtLast(.PDeductions.Vector, Array("CONCEPTO", Empty, Empty, "PORCENTAJE", "VALOR", Empty))
      ' Add contents
      If .PDExports.Exists Then Let .PDeductions.Vector = RES.ArrayConcat(.PDeductions.Vector, .PDExports.Vector)
      If .PDRMRets.Exists Then Let .PDeductions.Vector = RES.ArrayConcat(.PDeductions.Vector, .PDRMRets.Vector)
      If .PDLawRets.Exists Then Let .PDeductions.Vector = RES.ArrayConcat(.PDeductions.Vector, .PDLawRets.Vector)
      If .PDOthRets.Exists Then Let .PDeductions.Vector = RES.ArrayConcat(.PDeductions.Vector, .PDOthRets.Vector)
      If .PDDiscnts.Exists Then Let .PDeductions.Vector = RES.ArrayConcat(.PDeductions.Vector, .PDDiscnts.Vector)
      ' Add Totals
      With .PDeductions

        If PTypRes.CBOO_FORMULAS Then
          Let LngResize = UBound(.Vector) - 2 '3: Title, subtitle, subtitle content
          'Let PTypRes.TotalDeductionsRng = "F" & 9 + (UBound(PTypRes.PContents.Vector) - 2) + UBound(PTypRes.PAssessmentMain.Vector) + LngResize + 1 ' Main Deductions address
          'Let .StrRngA1 = "F" & 9 + (UBound(PTypRes.PContents.Vector) - 2) + UBound(PTypRes.PAssessmentMain.Vector) + 4 ' Deductions address
          Let .StrRngA1 = "K" & 7 + (UBound(PTypRes.PContents.Vector) + 1) + (UBound(PTypRes.PAssessmentMain.Vector) + 1) + 3 ' Deductions address
          'Let .StrRngA1 = PRngBox.Range(PTypRes.TotalAssessmentRng).Offset(3, 0).Resize(LngResize, 1).Address(False, False) ' Main Deductions address
          'Let .Vector = RES.ArrayConcat(.Vector, Array("TOTAL DEDUCCIONES [" & PTypRes.CSTR_FOREIGN_SYMBOL_CURRENCY & "]:", Empty, Empty, Empty, Empty, "=SUM(" & PRngBox.Range(.StrRngA1).Offset(0, 5).Resize(LngResize, 1).Address(False, False) & ")"))
          Let .Vector = RES.ArrayConcat(.Vector, Array("TOTAL DEDUCCIONES [" & PTypRes.CSTR_FOREIGN_SYMBOL_CURRENCY & "]:", Empty, Empty, Empty, Empty, "=SUM(" & PRngBox.Range(.StrRngA1).Resize(LngResize, 1).Address(False, False) & ")"))
        Else
          Let .Vector = RES.ArrayConcat(.Vector, Array("TOTAL DEDUCCIONES [" & PTypRes.CSTR_FOREIGN_SYMBOL_CURRENCY & "]:", Empty, Empty, Empty, Empty, PTypRes.TotalDeductionsDbl))
        End If
        Let .LngRngA1 = PTypRes.PAssessment.LngRngA1 + (UBound(.Vector) + 1)
        Let .LngRngA1 = PTypRes.PAssessment.LngRngA1 + (UBound(PTypRes.PAssessmentMain.Vector) + 1) + (UBound(.Vector) + 1)

      End With
      Let .PDExports.Vector = Empty
      Let .PDRMRets.Vector = Empty
      Let .PDLawRets.Vector = Empty
      Let .PDOthRets.Vector = Empty
      Let .PDDiscnts.Vector = Empty
      If Not VBA.IsArray(PTypRes.PDeductions.Vector) Then GoTo EH
      Let .TotalDeductionsRng = "K" & (9 + (UBound(.PContents.Vector))) + UBound(.PAssessmentMain.Vector) + UBound(.PDeductions.Vector)  ' Main Deductions address
    End If

  End With

  ' Return
  Let Purchasecnc_core_generate_content_deductions = True

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Purchasecnc_core_generate_content_deductions_exports( _
  ByVal PRngBox As Range, _
  ByRef PTypRes As GLOTYPPURCHASE) As Boolean

  Dim RngBox As Range
  Dim VecAux0 As Variant, VecAux1 As Variant
  Dim StrPer As String, StrCsh As String, StrTMH As String, StrAmt As String, StrVal As String
  Dim DblPer As Double, DblCsh As Double, DblTMH As Double, DblAmt As Double, DblVal As Double
  Dim BooDirect As Boolean, BooIndirect As Boolean
  Dim i As Long, j As Long, xCols As Long

  Const sFN As String = "CNC::-core_generate_content_deductions_exports"

  On Error GoTo EH
  Let Purchasecnc_core_generate_content_deductions_exports = False

  With PTypRes

    ' Get data (0.CONCEPTO 1.PORCENTAJE 2.VALOR 3.VALOR DIRECTO 4.SELECTED)
    Let .PDExports.Exists = VBA.IsArray(.TBL_PURCNC_EXPORTS)

    ' Fill matrix
    With .PDExports

      Let .LngRngA1 = PTypRes.PAssessmentMain.LngRngA1
      If .Exists Then

        Let .Exists = (RES.ArrayGetValuePosition(PTypRes.TBL_PURCNC_EXPORTS, 1, 4) >= 0) ' Exists selected rows in exports table
        If .Exists Then

          ' Check for Resume tables boolean
          If PTypRes.CBOO_RESUMEDEXPORTS Then
            Let VecAux1 = PTypRes.TBL_PURCNC_EXPORTS
            Let BooDirect = False
            Let BooIndirect = False
            ReDim VecAux0(0 To 1, 0 To UBound(VecAux1, 2))
            For i = 0 To UBound(VecAux1)
              If VBA.CBool(VecAux1(i, 4)) Then ' Selected
                If VBA.CBool(VecAux1(i, 3)) Then ' Direct value
                  Let VecAux0(0, 1) = VBA.CDbl(VecAux0(0, 1)) + VBA.CDbl(VecAux1(i, 2))
                  Let BooDirect = True
                Else ' Undirect value
                  Let VecAux0(1, 1) = VBA.CDbl(VecAux0(1, 1)) + VBA.CDbl(VecAux1(i, 1))
                  Let VecAux0(1, 2) = VBA.CDbl(VecAux0(1, 2)) + VBA.CDbl(VecAux1(i, 2))
                  Let BooIndirect = True
                End If
              End If
            Next i
            Let VecAux1 = Empty
            Let VecAux0(0, 0) = "Descuentos directos varios"
            Let VecAux0(1, 0) = "Descuentos varios"
            Let VecAux0(0, 3) = 1
            Let VecAux0(1, 3) = 0
            Let VecAux0(0, 4) = 1
            Let VecAux0(1, 4) = 1
            If Not BooDirect Then Let VecAux0 = RES.ArrayPop(VecAux0)
            If Not BooIndirect Then Let VecAux0 = RES.ArrayShift(VecAux0)
          Else
            Let VecAux0 = PTypRes.TBL_PURCNC_EXPORTS
          End If

          ' Determine Addresses and row numbers
          Let xCols = 5
          Let .LngRngA1 = .LngRngA1 + 4
          Let .StrRngA1 = "F" & .LngRngA1
          If PTypRes.CBOO_FORMULAS Then Set RngBox = PRngBox.Range(.StrRngA1)
          Let .LngRngA1 = .LngRngA1 + UBound(VecAux0)

          ' Fill final vector with data
          ReDim .Vector(0 To 0, 0 To xCols)
          Let .Vector(0, 0) = "DESCUENTOS POR EXPORTACIÓN"

          ' Calculate results
          If PTypRes.CBOO_FORMULAS Then
            Let StrTMH = PRngBox.Range("G6").Address(False, False) ' TMH
            Let StrAmt = PRngBox.Range(PTypRes.PAssessmentMain.StrRngA1).Offset(0, 5).Address(False, False) ' Assessment
          Else
            Let DblTMH = VBA.CDbl(PTypRes.PWeight.Vector(2, 1)) ' TMH
            Let DblAmt = VBA.CDbl(PTypRes.PAssessmentMain.Vector(UBound(PTypRes.PAssessmentMain.Vector), UBound(PTypRes.PAssessmentMain.Vector, 2))) ' Assessment
          End If

          For i = 0 To UBound(VecAux0)

            If VBA.CBool(VecAux0(i, 4)) Then ' Selected
              If PTypRes.CBOO_FORMULAS Then
                Let StrPer = RngBox.Offset(i, 3).Address(False, False) ' Percent
                Let StrCsh = RngBox.Offset(i, 4).Address(False, False) ' Value
                ' Calculate final value: =IF(Per>0;Amt*Per*0.01;TMH*Csh)
                Let StrVal = "=IF(" & StrPer & ">0," & "(" & StrAmt & "*" & StrPer & "*0.01)," & StrTMH & "*" & StrCsh & ")"
                Let StrVal = IIf(VBA.CBool(VecAux0(i, 3)), VBA.Replace(VBA.CStr(VecAux0(i, 2)), ",", "."), StrVal)
                Let .Vector = RES.ArrayAddAtLast(.Vector, Array(VBA.StrConv(VBA.CStr(VecAux0(i, 0)), vbProperCase), Empty, Empty, VBA.Replace(VBA.CStr(VecAux0(i, 1)), ",", "."), VBA.Replace(VBA.CStr(VecAux0(i, 2)), ",", "."), StrVal))
              Else
                Let DblPer = VBA.CDbl(VecAux0(i, 1))  ' Percent
                Let DblCsh = VBA.CDbl(VecAux0(i, 2)) ' Value
                ' Calculate final value: =IF(Per>0;Amt*Per*0.01;TMH*Csh)
                Let DblVal = IIf(DblPer > 0, (DblAmt * DblPer * 0.01), DblTMH * DblCsh)
                Let DblVal = IIf(VBA.CBool(VecAux0(i, 3)), VBA.CDbl(VecAux0(i, 2)), DblVal)
                Let .Vector = RES.ArrayAddAtLast(.Vector, Array(VBA.StrConv(VBA.CStr(VecAux0(i, 0)), vbProperCase), Empty, Empty, VecAux0(i, 1), VecAux0(i, 2), DblVal))
                Let PTypRes.TotalDeductionsDbl = PTypRes.TotalDeductionsDbl + DblVal
              End If
            End If

          Next i
          If Not VBA.IsArray(.Vector) Then GoTo EH

        End If

      End If

    End With

  End With

  ' Return
  Let Purchasecnc_core_generate_content_deductions_exports = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Set RngBox = Nothing
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Purchasecnc_core_generate_content_deductions_rmtaxes( _
  ByVal PRngBox As Range, _
  ByRef PTypRes As GLOTYPPURCHASE) As Boolean

  Dim RngBox As Range
  Dim VecAux0 As Variant, VecAux1 As Variant
  Dim StrPer As String, StrRM As String, StrAmt As String, StrVal As String
  Dim DblPer As Double, DblRM As Double, DblAmt As Double, DblVal As Double
  Dim i As Long, j As Long, xCols As Long

  Const sFN As String = "CNC::-core_generate_content_deductions_rmtaxes"

  On Error GoTo EH
  Let Purchasecnc_core_generate_content_deductions_rmtaxes = False

  With PTypRes

    ' Get data (0.CONCEPTO 1.PORCENTAJE)
    ReDim VecAux0(0 To 0, 0 To 1)
    For i = 0 To UBound(.PContents.Vector)
      If i >= 2 Then
        Let j = RES.ArrayGetValuePosition(.TBL_PUR_RM, VBA.CStr(.PContents.Vector(i, 0)), 0)
        Let VecAux0 = RES.ArrayConcat(VecAux0, Array(.TBL_PUR_RM(j, 0), .TBL_PUR_RM(j, 5)))
      End If
    Next i
    Call RES.ArrayShift(VecAux0)
    Let .PDRMRets.Exists = VBA.IsArray(VecAux0)

    ' Fill matrix
    With .PDRMRets

      Let .LngRngA1 = PTypRes.PDExports.LngRngA1
      If .Exists Then

        ' Check for Resume tables boolean
        If PTypRes.CBOO_RESUMEDRM Then
          Let VecAux1 = VecAux0
          Let VecAux0 = Empty
          ReDim VecAux0(0 To 0, 0 To UBound(VecAux1, 2))
          For i = 0 To UBound(VecAux1)
            Let VecAux0(0, 1) = VBA.CDbl(VecAux0(0, 1)) + VBA.CDbl(VecAux1(i, 1))
          Next i
          Let VecAux0(0, 0) = "RM"
          Let VecAux1 = Empty
        End If

        ' Determine Addresses and row numbers
        Let xCols = 5
        Let .LngRngA1 = .LngRngA1 + IIf(PTypRes.PDExports.Exists, 2, 4)
        Let .StrRngA1 = "F" & .LngRngA1
        If PTypRes.CBOO_FORMULAS Then Set RngBox = PRngBox.Range(.StrRngA1)
        Let .LngRngA1 = .LngRngA1 + UBound(VecAux0)

        ' Fill final vector with data
        ReDim .Vector(0 To 0, 0 To xCols)
        Let .Vector(0, 0) = "RETENCIONES DE LEY - REGALÍAS"

        ' Calculate results
        If PTypRes.CBOO_FORMULAS Then
          Let StrAmt = PRngBox.Range(PTypRes.PAssessmentMain.StrRngA1).Address(False, False) ' Assessment'5).Address(False, False) ' Assessment
        Else
          Let DblAmt = VBA.CDbl(PTypRes.PAssessmentMain.Vector(UBound(PTypRes.PAssessmentMain.Vector), UBound(PTypRes.PAssessmentMain.Vector, 2))) ' Assessment
        End If

        For i = 0 To UBound(VecAux0)

          Let StrRM = IIf(VBA.LCase(VBA.CStr(VecAux0(i, 0))) = "rm", VBA.CStr(VecAux0(i, 0)), VBA.StrConv(VBA.CStr(VecAux0(i, 0)), vbProperCase))
          If PTypRes.CBOO_FORMULAS Then
            Let StrPer = RngBox.Offset(i, 3).Address(False, False) ' RM Percent
            ' Calculate final value: =Per*Amt*/0.01
            Let StrVal = "=" & StrPer & "*" & StrAmt & "*0.01"
            Let .Vector = RES.ArrayAddAtLast(.Vector, Array(StrRM, Empty, Empty, VBA.Replace(VBA.CStr(VecAux0(i, 1)), ",", "."), 0, StrVal))
          Else
            Let DblPer = VBA.CDbl(VecAux0(i, 1)) ' RM Percent
            ' Calculate final value: =Per*Amt*/0.01
            Let DblVal = DblPer * DblAmt * 0.01
            Let .Vector = RES.ArrayAddAtLast(.Vector, Array(StrRM, Empty, Empty, VecAux0(i, 1), 0, DblVal))
            Let PTypRes.TotalDeductionsDbl = PTypRes.TotalDeductionsDbl + DblVal
          End If

        Next i
        If Not VBA.IsArray(.Vector) Then GoTo EH

      End If

    End With

  End With

  ' Return
  Let Purchasecnc_core_generate_content_deductions_rmtaxes = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Set RngBox = Nothing
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Purchasecnc_core_generate_content_deductions_lawtaxes( _
  ByVal PRngBox As Range, _
  ByRef PTypRes As GLOTYPPURCHASE) As Boolean

  Dim RngBox As Range
  Dim VecAux0 As Variant
  Dim StrPer As String, StrRet As String, StrAmt As String, StrVal As String
  Dim DblPer As Double, DblRM As Double, DblAmt As Double, DblVal As Double
  Dim i As Long, j As Long, xCols As Long

  Const sFN As String = "CNC::-core_generate_content_deductions_lawtaxes"

  On Error GoTo EH
  Let Purchasecnc_core_generate_content_deductions_lawtaxes = False

  With PTypRes

    ' Get data (0.CONCEPTO 1.PORCENTAJE)
    Let j = RES.ArrayGetValuePosition(.TBL_TMP_RETENTION, 1, 2)
    Let .PDLawRets.Exists = (j >= 0)

    ' Fill matrix
    With .PDLawRets

      Let .LngRngA1 = PTypRes.PDRMRets.LngRngA1
      If .Exists Then

        'Check for Resume tables boolean
        If PTypRes.CBOO_RESUMEDRETENTION Then
          ReDim VecAux0(0 To 0, 0 To 2)
          For i = LBound(PTypRes.TBL_TMP_RETENTION) To UBound(PTypRes.TBL_TMP_RETENTION)
            If VBA.CBool(PTypRes.TBL_TMP_RETENTION(i, 2)) Then Let VecAux0(0, 1) = VBA.CDbl(VecAux0(0, 1)) + VBA.CDbl(PTypRes.TBL_TMP_RETENTION(i, 1))
          Next i
          Let VecAux0(0, 0) = "Varias"
          Let VecAux0(0, 2) = 1
        Else
          ReDim VecAux0(0 To 0, 0 To 2)
          For i = LBound(PTypRes.TBL_TMP_RETENTION) To UBound(PTypRes.TBL_TMP_RETENTION)
            If VBA.CBool(PTypRes.TBL_TMP_RETENTION(i, 2)) Then Let VecAux0 = RES.ArrayConcat(VecAux0, Array(PTypRes.TBL_TMP_RETENTION(i, 0), PTypRes.TBL_TMP_RETENTION(i, 1), PTypRes.TBL_TMP_RETENTION(i, 2)))
          Next i
          Call RES.ArrayShift(VecAux0)
        End If

        ' Determine Addresses and row numbers
        Let xCols = 5
        Let .LngRngA1 = .LngRngA1 + IIf(PTypRes.PDRMRets.Exists, 2, 4)
        Let .StrRngA1 = "F" & .LngRngA1
        If PTypRes.CBOO_FORMULAS Then Set RngBox = PRngBox.Range(.StrRngA1)
        Let .LngRngA1 = .LngRngA1 + UBound(VecAux0)

        ' Fill final vector with data
        ReDim .Vector(0 To 0, 0 To xCols)
        Let .Vector(0, 0) = "RETENCIONES DE LEY - INSTITUCIONALES"

        ' Calculate results
        If PTypRes.CBOO_FORMULAS Then
          Let StrAmt = PRngBox.Range(PTypRes.PAssessmentMain.StrRngA1).Address(False, False) '5).Address(False, False) ' Assessment
        Else
          Let DblAmt = VBA.CDbl(PTypRes.PAssessmentMain.Vector(UBound(PTypRes.PAssessmentMain.Vector), UBound(PTypRes.PAssessmentMain.Vector, 2))) ' Assessment
        End If

        For i = 0 To UBound(VecAux0)

          Let StrRet = VBA.UCase(VBA.CStr(VecAux0(i, 0)))
          If PTypRes.CBOO_FORMULAS Then
            Let StrPer = RngBox.Offset(i, 3).Address(False, False) ' Retention Percent
            ' Calculate final value: =Per*Amt*/0.01
            Let StrVal = "=" & StrPer & "*" & StrAmt & "*0.01"
            Let .Vector = RES.ArrayAddAtLast(.Vector, Array(StrRet, Empty, Empty, VBA.Replace(VBA.CStr(VecAux0(i, 1)), ",", "."), 0, StrVal))
          Else
            Let DblPer = VBA.CDbl(VecAux0(i, 1)) ' Retention Percent
            ' Calculate final value: =Per*Amt*/0.01
            Let DblVal = DblPer * DblAmt * 0.01
            Let .Vector = RES.ArrayAddAtLast(.Vector, Array(StrRet, Empty, Empty, VecAux0(i, 1), 0, DblVal))
            Let PTypRes.TotalDeductionsDbl = PTypRes.TotalDeductionsDbl + DblVal
          End If

        Next i
        If Not VBA.IsArray(.Vector) Then GoTo EH

      End If

    End With

  End With

  ' Return
  Let Purchasecnc_core_generate_content_deductions_lawtaxes = True

EH:
  Let VecAux0 = Empty
  Set RngBox = Nothing
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Purchasecnc_core_generate_content_deductions_othertaxes( _
  ByVal PRngBox As Range, _
  ByRef PTypRes As GLOTYPPURCHASE) As Boolean

  Dim RngBox As Range
  Dim VecAux0 As Variant, VecAux1 As Variant
  Dim i As Long, j As Long, xCols As Long
  Dim StrPer As String, StrRet As String, StrAmt As String, StrVal As String
  Dim DblPer As Double, DblRM As Double, DblAmt As Double, DblVal As Double

  Const sFN As String = "CNC::-core_generate_content_deductions_othertaxes"

  On Error GoTo EH
  Let Purchasecnc_core_generate_content_deductions_othertaxes = False

  With PTypRes

    ' Get data (0.CONCEPTO 1.PORCENTAJE)
    Let j = RES.ArrayGetValuePosition(.TBL_TMP_RETENTION_OTHER, 1, 2)
    Let .PDOthRets.Exists = (j >= 0)

    ' Fill matrix
    With .PDOthRets

      Let .LngRngA1 = PTypRes.PDLawRets.LngRngA1
      If .Exists Then

        ' Check for Resume tables boolean
        If PTypRes.CBOO_RESUMEDRETENTIONOTHER Then
          ReDim VecAux0(0 To 0, 0 To 2)
          For i = LBound(PTypRes.TBL_TMP_RETENTION_OTHER) To UBound(PTypRes.TBL_TMP_RETENTION_OTHER)
            If VBA.CBool(PTypRes.TBL_TMP_RETENTION_OTHER(i, 2)) Then Let VecAux0(0, 1) = VBA.CDbl(VecAux0(0, 1)) + VBA.CDbl(PTypRes.TBL_TMP_RETENTION_OTHER(i, 1))
          Next i
          Let VecAux0(0, 0) = "Varias"
          Let VecAux0(0, 2) = 1
        Else
          ReDim VecAux0(0 To 0, 0 To 2)
          For i = LBound(PTypRes.TBL_TMP_RETENTION_OTHER) To UBound(PTypRes.TBL_TMP_RETENTION_OTHER)
            If VBA.CBool(PTypRes.TBL_TMP_RETENTION_OTHER(i, 2)) Then Let VecAux0 = RES.ArrayConcat(VecAux0, Array(PTypRes.TBL_TMP_RETENTION_OTHER(i, 0), PTypRes.TBL_TMP_RETENTION_OTHER(i, 1), PTypRes.TBL_TMP_RETENTION_OTHER(i, 2)))
          Next i
          Call RES.ArrayShift(VecAux0)
        End If

        ' Determine Addresses and row numbers
        Let xCols = 5
        Let .LngRngA1 = .LngRngA1 + IIf(PTypRes.PDRMRets.Exists, 2, 4)
        Let .StrRngA1 = "F" & .LngRngA1
        If PTypRes.CBOO_FORMULAS Then Set RngBox = PRngBox.Range(.StrRngA1)
        Let .LngRngA1 = .LngRngA1 + UBound(VecAux0)

        ' Fill final vector with data
        ReDim .Vector(0 To 0, 0 To xCols)
        Let .Vector(0, 0) = "RETENCIONES INSTITUCIONALES OTRAS"

        ' Calculate results
        If PTypRes.CBOO_FORMULAS Then
          Let StrAmt = PRngBox.Range(PTypRes.PAssessmentMain.StrRngA1).Address(False, False) '5).Address(False, False) ' Assessment
        Else
          Let DblAmt = VBA.CDbl(PTypRes.PAssessmentMain.Vector(UBound(PTypRes.PAssessmentMain.Vector), UBound(PTypRes.PAssessmentMain.Vector, 2))) ' Assessment
        End If

        For i = 0 To UBound(VecAux0)

          Let StrRet = VBA.UCase(VBA.CStr(VecAux0(i, 0)))
          If PTypRes.CBOO_FORMULAS Then
            Let StrPer = RngBox.Offset(i, 3).Address(False, False) ' Retention Percent
            ' Calculate final value: =Per*Amt*/0.01
            Let StrVal = "=" & StrPer & "*" & StrAmt & "*0.01"
            Let .Vector = RES.ArrayAddAtLast(.Vector, Array(StrRet, Empty, Empty, VBA.Replace(VBA.CStr(VecAux0(i, 1)), ",", "."), 0, StrVal))
          Else
            Let DblPer = VBA.CDbl(VecAux0(i, 1)) ' Retention Percent
            ' Calculate final value: =Per*Amt*/0.01
            Let DblVal = DblPer * DblAmt * 0.01
            Let .Vector = RES.ArrayAddAtLast(.Vector, Array(StrRet, Empty, Empty, VecAux0(i, 1), 0, DblVal))
            Let PTypRes.TotalDeductionsDbl = PTypRes.TotalDeductionsDbl + DblVal
          End If

        Next i
        If Not VBA.IsArray(.Vector) Then GoTo EH

      End If

    End With

  End With

  ' Return
  Let Purchasecnc_core_generate_content_deductions_othertaxes = True

EH:
  Let VecAux0 = Empty
  Set RngBox = Nothing
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Purchasecnc_core_generate_content_deductions_discounts( _
  ByVal PRngBox As Range, _
  ByRef PTypRes As GLOTYPPURCHASE) As Boolean

  Dim RngBox As Range
  Dim VecAux0 As Variant
  Dim StrDsc As String, StrNme As String, StrVal As String
  Dim DblDsc As Double, DblVal As Double, DblAux0 As Double
  Dim i As Long, xCols As Long

  Const sFN As String = "CNC::-core_spreadmodule_generate_data_purchasebase_getdeductions_discounts"

  On Error GoTo EH
  Let Purchasecnc_core_generate_content_deductions_discounts = False

  With PTypRes

    ' Get data (0.CONCEPTO 1.VALOR)
    Let .PDDiscnts.Exists = VBA.IsArray(.TBL_TMP_DISCOUNTS)

    ' Fill matrix
    With .PDDiscnts

      Let .LngRngA1 = PTypRes.PDOthRets.LngRngA1
      If .Exists Then

        ' Check for Resume tables boolean
        If PTypRes.CBOO_RESUMEDDISCOUNTS Then
          Let VecAux0 = RES.ArrayPop(PTypRes.TBL_TMP_DISCOUNTS, True, RByCols:=True)
          Let DblAux0 = Application.WorksheetFunction.Sum(VecAux0)
          Let VecAux0 = Empty
          ReDim VecAux0(0 To 0, 0 To 1)
          Let VecAux0(0, 0) = "Descuentos Varios"
          Let VecAux0(0, 1) = DblAux0
        Else
          Let VecAux0 = RES.ArraySlice(PTypRes.TBL_TMP_DISCOUNTS, 1, RByCols:=True)
        End If

        ' Determine Addresses and row numbers
        Let xCols = 5
        Let .LngRngA1 = .LngRngA1 + IIf(PTypRes.PDOthRets.Exists, 2, 4)
        Let .StrRngA1 = "F" & .LngRngA1
        If PTypRes.CBOO_FORMULAS Then Set RngBox = PRngBox.Range(.StrRngA1)
        Let .LngRngA1 = .LngRngA1 + UBound(VecAux0)

        ' Fill final vector with data
        ReDim .Vector(0 To 0, 0 To xCols)
        Let .Vector(0, 0) = "DESCUENTOS"

        ' Calculate results
        For i = 0 To UBound(VecAux0)

          Let StrNme = VBA.StrConv(VBA.CStr(VecAux0(i, 0)), vbProperCase)
          If PTypRes.CBOO_FORMULAS Then
            Let StrVal = VBA.Replace(VecAux0(i, 1), ",", ".")
            Let .Vector = RES.ArrayAddAtLast(.Vector, Array(StrNme, Empty, Empty, 0, StrVal, StrVal))
          Else
            Let DblVal = VBA.CDbl(VecAux0(i, 1))
            Let .Vector = RES.ArrayAddAtLast(.Vector, Array(StrNme, Empty, Empty, 0, DblVal, DblVal))
            Let PTypRes.TotalDeductionsDbl = PTypRes.TotalDeductionsDbl + DblVal
          End If

        Next i
        If Not VBA.IsArray(.Vector) Then GoTo EH

      End If

    End With

  End With

  ' Return
  Let Purchasecnc_core_generate_content_deductions_discounts = True

EH:
  Let VecAux0 = Empty
  Set RngBox = Nothing
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Purchasecnc_core_generate_content_bonus( _
  ByVal PRngBox As Range, _
  ByRef PTypRes As GLOTYPPURCHASE) As Boolean

  Dim RngBox As Range
  Dim VecAux0 As Variant
  Dim StrNme As String, StrVal As String
  Dim DblVal As Double, DblAux0 As Double
  Dim i As Long, xCols As Long

  Const sFN As String = "CNC::-core_generate_content_bonus"

  On Error GoTo EH
  Let Purchasecnc_core_generate_content_bonus = False

  With PTypRes

    ' Get data (0.CONCEPTO 1.VALOR)
    Let .PBonus.Exists = VBA.IsArray(.TBL_TMP_BONUS)
    Let .TotalBonusDbl = 0

    ' Fill matrix
    With .PBonus

      'Let .LngRngA1 = PTypRes.PAssessmentMain.LngRngA1 + PTypRes.PDeductions.LngRngA1
      Let .LngRngA1 = PTypRes.PDeductions.LngRngA1
      If .Exists Then

        ' Check for Resume tables boolean
        If PTypRes.CBOO_RESUMEDBONUS Then
          Let VecAux0 = RES.ArrayPop(PTypRes.TBL_TMP_BONUS, True, RByCols:=True)
          Let DblAux0 = Application.WorksheetFunction.Sum(VecAux0)
          Let VecAux0 = Empty
          ReDim VecAux0(0 To 0, 0 To 1)
          Let VecAux0(0, 0) = "Bonos Varios"
          Let VecAux0(0, 1) = DblAux0
        Else
          Let VecAux0 = RES.ArraySlice(PTypRes.TBL_TMP_BONUS, 0, True, True)
        End If

        ' Determine Addresses and row numbers
        Let xCols = 5
        Let .LngRngA1 = .LngRngA1 + 2
        Let .StrRngA1 = "F" & .LngRngA1
        If PTypRes.CBOO_FORMULAS Then Set RngBox = PRngBox.Range(.StrRngA1)
        Let .LngRngA1 = .LngRngA1 + UBound(VecAux0)

        ' Fill final vector with data
        ReDim .Vector(0 To 0, 0 To xCols)
        Let .Vector(0, 0) = "BONOS"

        ' Calculate results
        For i = 0 To UBound(VecAux0)

          Let StrNme = VBA.StrConv(VBA.CStr(VecAux0(i, 0)), vbProperCase)
          If PTypRes.CBOO_FORMULAS Then
            Let StrVal = VBA.Replace(VecAux0(i, 1), ",", ".")
            Let .Vector = RES.ArrayAddAtLast(.Vector, Array(StrNme, Empty, Empty, Empty, Empty, StrVal))
          Else
            Let DblVal = VBA.CDbl(VecAux0(i, 1))
            Let .Vector = RES.ArrayAddAtLast(.Vector, Array(StrNme, Empty, Empty, Empty, Empty, DblVal))
            Let PTypRes.TotalBonusDbl = PTypRes.TotalBonusDbl + DblVal
          End If

        Next i
        If Not VBA.IsArray(.Vector) Then GoTo EH

        ' Fill totals
        If PTypRes.CBOO_FORMULAS Then
          'Let .Vector = RES.ArrayAddAtLast(.Vector, Array("TOTAL BONOS [" & PTypRes.CSTR_FOREIGN_SYMBOL_CURRENCY & "]", Empty, Empty, Empty, Empty, "=SUM(" & RngBox.Offset(0, 5).Resize(UBound(VecAux0) + 1, 1).Address(False, False) & ")"))
          Let .Vector = RES.ArrayAddAtLast(.Vector, Array("TOTAL BONOS [" & PTypRes.CSTR_FOREIGN_SYMBOL_CURRENCY & "]", Empty, Empty, Empty, Empty, "=SUM(" & PRngBox.Range(PTypRes.TotalDeductionsRng).Offset(2, 0).Resize(UBound(VecAux0) + 1, 1).Address(False, False) & ")"))
        Else
          Let .Vector = RES.ArrayAddAtLast(.Vector, Array("TOTAL BONOS [" & PTypRes.CSTR_FOREIGN_SYMBOL_CURRENCY & "]", Empty, Empty, Empty, Empty, PTypRes.TotalBonusDbl))
        End If
        If Not VBA.IsArray(.Vector) Then GoTo EH

      End If

    End With
    If PTypRes.CBOO_FORMULAS Then
      Let .PBonus.LngRngA1 = PRngBox.Range("K1:" & PTypRes.TotalDeductionsRng).Rows.Count + UBound(PTypRes.PBonus.Vector) + 1
      Let .TotalBonusRng = "K" & .PBonus.LngRngA1
    End If

  End With

  ' Return
  Let Purchasecnc_core_generate_content_bonus = True

EH:
  Let VecAux0 = Empty
  Set RngBox = Nothing
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Purchasecnc_core_generate_content_finalpurchase( _
  ByVal PRngBox As Range, _
  ByRef PTypRes As GLOTYPPURCHASE) As Boolean

  Dim NTS As ClassEqlControllerNumToStr

  Dim RngBox As Range
  Dim StrAux0 As String
  Dim DblAux0 As Double
  Dim xCols As Long, xRows As Long

  Const sFN As String = "CNC::-core_generate_content_finalpurchase"

  On Error GoTo EH
  Let Purchasecnc_core_generate_content_finalpurchase = False

  With PTypRes

    ' Fill matrix PLngRowCounter
    With .PFinal

      ' Determine Addresses and row numbers
      Let .LngRngA1 = PTypRes.PBonus.LngRngA1
      Let xCols = 10
      Let xRows = 4
      Let .LngRngA1 = .LngRngA1 + 3
      Let .StrRngA1 = "A" & .LngRngA1
      If PTypRes.CBOO_FORMULAS Then Set RngBox = PRngBox.Range(.StrRngA1)

      ' Fill final vector with data
      Let DblAux0 = 0
      ReDim .Vector(0 To xRows, 0 To xCols)

      Let .Vector(0, 0) = "LIQUIDACIÓN FINAL (VALOR BRUTO" & IIf(PTypRes.PDeductions.Exists, " - DEDUCCIONES", VBA.vbNullString) & IIf(PTypRes.PBonus.Exists, " + BONOS", VBA.vbNullString) & ")"
      If PTypRes.CBOO_MN Then Let .Vector(0, 9) = "VALOR " & PTypRes.CSTR_LOCAL_SYMBOL_CURRENCY ' MN
      Let .Vector(0, 10) = "VALOR " & PTypRes.CSTR_FOREIGN_SYMBOL_CURRENCY ' ME

      If PTypRes.CBOO_FORMULAS Then

        ' Set liquid
        Let StrAux0 = RngBox.Offset(0, 10).Address(False, False)
        Let .Vector(1, 0) = "=""Líquido pagable ("" & IF(" & StrAux0 & "=0,""Sin saldo"",IF(" & StrAux0 & ">0,""Saldo a favor"",""Saldo en contra"")) & ""):"""
        If PTypRes.CBOO_MN Then Let .Vector(1, 9) = "=" & RngBox.Offset(0, 10).Address(False, False) & "*" & PRngBox.Range("B10").Address(False, False) 'MN
        Let .Vector(1, 10) = "=SUM(" & _
          PRngBox.Range(PTypRes.TotalAssessmentRng).Address(False, False) & _
          IIf(PTypRes.PDeductions.Exists, ",-" & PRngBox.Range(PTypRes.TotalDeductionsRng).Address(False, False), VBA.vbNullString) & _
          IIf(PTypRes.PBonus.Exists, "," & PRngBox.Range(PTypRes.TotalBonusRng).Address(False, False), VBA.vbNullString) & ")"
        ' Set liquid textual
        'Let .Vector(2, 1) = "=EQL_NUMERO_A_TEXTO(" & RngBox.Offset(0, 10).Address(False, False) & ",'" & PTypRes.CSTR_FOREIGN_CURRENCY & "','" & PTypRes.CSTR_FOREIGN_CURRENCIES & "',0)"
        Let .Vector(2, 1) = "=EQL_NUMERO_A_TEXTO(" & RngBox.Offset(0, 10).Address(False, False) & ",""" & PTypRes.CSTR_FOREIGN_CURRENCY & """,""" & PTypRes.CSTR_FOREIGN_CURRENCIES & """,0)"

      Else

        ' Set liquid
        Let DblAux0 = (PTypRes.TotalAssessmentDbl - PTypRes.TotalDeductionsDbl) + PTypRes.TotalBonusDbl
        Let .Vector(1, 0) = "Líquido pagable (" & IIf(DblAux0 = 0, "Sin saldo", IIf(DblAux0 > 0, "Saldo a favor", "Saldo en contra")) & "):"
        If PTypRes.CBOO_MN Then Let .Vector(1, 9) = DblAux0 * PTypRes.CDBL_TC 'MN
        Let .Vector(1, 10) = DblAux0
        ' Set liquid textual
        Set NTS = New ClassEqlControllerNumToStr
        Set NTS.ClassEqlControllerInterface_ResourcesApp = RES
        Let .Vector(2, 1) = NTS.NumToString(DblAux0, PTypRes.CSTR_FOREIGN_CURRENCY, PTypRes.CSTR_FOREIGN_CURRENCIES, False)
        Set NTS = Nothing

      End If
      Let .Vector(2, 0) = "Son:"

    End With

  End With

  ' Return
  Let Purchasecnc_core_generate_content_finalpurchase = True

EH:
  Set RngBox = Nothing
  Set NTS = Nothing
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function









' ASSETS --------------------------------------------------------------------------- '
Private Sub Purchasecnc_core_assets_srcarray_to_dstarray(ByRef PVarSrc As Variant)

  Dim i As Long, j As Long

  On Error GoTo EH
  For j = 0 To UBound(PVarSrc, 2)
    For i = 0 To UBound(PVarSrc)
      If (Not PVarSrc(i, j) = Empty) Or (PVarSrc(i, j) = 0) Then
        If VBA.IsNumeric(PVarSrc(i, j)) Then
          Let PVarSrc(i, j) = VBA.Replace(PVarSrc(i, j), ",", ".")
        End If
      End If
    Next i
  Next j

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise("CNC::-core_assets_srcarray_to_dstarray")

End Sub






' ERRORS MANNAGEMENT AND ASSERTSIONS ----------------------------------------------- '
' Debug_asset                                                    All Assets in module'
Private Sub Debug_assert(ByVal Cond As Boolean, Optional Origin As String = VBA.vbNullString)

  #If Debbuging Then
    Debug.Assert Cond
    If Not Cond Then Debug.Print "Assert on: " & Origin
  #End If

End Sub





' METHODS PUBLIC ------------------------------------------------------------------- '
' ABSTRACT INTERFACE UI METHODS LIST AND UDF METHODS LIST (PUBLIC) ----------------- '
Public Function ClassEqlControllerInterface_GetSolution(ByVal CStrRequester As String, ByRef CVarRequest As Variant, ByRef CVarResponse As Variant) As Boolean

  On Error GoTo EH
  Let ClassEqlControllerInterface_GetSolution = Purchasecnc_core_main(CVarRequest, CVarResponse)

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise("CNC::IGetSolution")

End Function

Private Function ClassEqlControllerInterface_GetSolutionUDF(ByVal CStrRequester As String, ByVal CVarRequest As Variant) As Variant
End Function
' METHODS PUBLIC ------------------------------------------------------------------- '
' ========== CLASS EXCLIQ LITE CONTROLLER PURCHASE CONCENTRATE STRUCTURE =========== '

