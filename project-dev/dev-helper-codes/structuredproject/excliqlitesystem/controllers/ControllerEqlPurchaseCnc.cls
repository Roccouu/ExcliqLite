VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppExcliqPurchaseCnc_current"
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
  PWeight As GLOTYPCOMPONENT 'PNS
  PContents As GLOTYPCOMPONENT
  PAssessmentMain As GLOTYPCOMPONENT 'Total/T
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

  GFormulasBoo As Boolean
  GRowcounterLng As Long
  
  GMNBoo As Boolean
  GTCDbl As Double
  
  GAverageAnywayBoo As Boolean
  GResumedExportsBoo As Boolean
  GResumedRMBoo As Boolean
  GResumedRetentionsBoo As Boolean
  GResumedREetentionsOthersBoo As Boolean
  GResumedDiscountsBoo As Boolean
  GResumedBonusBoo As Boolean
  
  GCurMainNameStr As String
  GCurMainNameStrs As String
  GCurMainSymStr As String
  GCurForeNameStr As String
  GCurForeNameStrs As String
  GCurForeSymStr As String
End Type

' PROPERTIES VARIABLES --------------------------------------------------------------'

' CONSTANTS ERROR MESSAGES (GLOSTR_ERR_) --------------------------------------------'
Private Const CUSTOM_ERROR = VBA.vbObjectError + 518




' ================== CLASS APPEXCLIQPURCHASEMINERAL STRUCTURE ======================='
' OBJECT VARIABLES (GLOOBJ_) --------------------------------------------------------'
' VARIABLES (GLOStr_) ---------------------------------------------------------------'
' CONSTANTS ERROR MESSAGES (GLOSTR_ERR_) --------------------------------------------'
' PROPERTIES ------------------------------------------------------------------------'
' CONSTRUCTOR/DESTRUCTOR ------------------------------------------------------------'
' CORE METHODS PRIVATE --------------------------------------------------------------'
' CORE METHODS LIST (PRIVATE) -------------------------------------------------------'
' PurchaseConcentrate_Core_Main
' PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasebase
' PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal
' PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getcncdata
' PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getcnccontents
' PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getassessment
' PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getdeductions
' PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getbonus
' PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getfinalpurchase
' ASSETS ----------------------------------------------------------------------------'
' PurchaseConcentrate_Core_assets_get_average
' PurchaseConcentrate_Core_assets_get_grades
' PurchaseConcentrate_Core_assets_set_strlos
' PurchaseConcentrate_Core_assets_get_deductions_rm
' PurchaseConcentrate_Core_assets_get_deductions_retentions
' PurchaseConcentrate_Core_assets_get_deductions_retentionsothers
' PurchaseConcentrate_Core_assets_get_deductions_discounts
' ERRORS MANNAGEMENT AND ASSERTSIONS ------------------------------------------------'
' Debug_assert
' METHODS PUBLIC --------------------------------------------------------------------'
' INTERFACE METHODS LIST ------------------------------------------------------------'
' PurchaseConcentrateGet
' =================== CLASS APPEXCLIQPURCHASEMINERAL STRUCTURE ======================'





' PROPERTIES ------------------------------------------------------------------------'
' ABSTRACT INTERFACE PROPERTIES IMPLEMENTATION --------------------------------------'
Public Property Set ICoreController_ErrorHandler(RHS As AppErrorHandler)
  
  Set EHGLOBAL = Nothing
  Set EHGLOBAL = RHS
  Let GLOBOO_EH = (Not RHS Is Nothing)

End Property










' CONSTRUCTOR/DESTRUCTOR ------------------------------------------------------------'
Private Sub Class_Initialize()
  
  Set EHGLOBAL = New AppErrorHandler

End Sub

Private Sub Class_Terminate()
  
  Set GLOSHTDATA = Nothing
  If GLOBOO_EH = False Then Set EHGLOBAL = Nothing

End Sub





' CORE METHODS PRIVATE --------------------------------------------------------------'
' CORE METHODS LIST (PRIVATE) -------------------------------------------------------'
Private Function PurchaseConcentrate_Core_Main(ByRef PVarRequest As Variant, PVarResponse As Variant) As Boolean

  Dim RES As AppResources_current
  
  Dim VecAux0 As Variant
  Dim TypCnc As GLOTYPPURCHASE
  
  Const sFN As String = "APPCNC::-core_main"

  On Error GoTo EH
  Let PurchaseConcentrate_Core_Main = False

  'Get currencies values
  With TypCnc
    If VBA.CBool(PVarRequest(4)) Then 'It will be a boolean = "lite"
      Set MLT = New ModelExcliqliteDatasheet
      Set MLT.ErrorHandler = EHGLOBAL
      Let VecAux0 = MLT.MGet(eqlMdlTblCurrencies, eqlMdlArray, "currency, currencies, symbol", MStrWhere:="currency_foreign=1")
      If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(0, "¡Tabla de Monedas sin valores iniciales!"): GoTo EH
      Let VecAux0 = RES.ArrayToBaseZero(VecAux0, eqlRes1D)
      Let .GCurForeNameStr = VBA.CStr(VecAux0(0))
      Let .GCurForeNameStrs = VBA.CStr(VecAux0(1))
      Let .GCurForeSymStr = VBA.CStr(VecAux0(2))
      Let VecAux0 = Empty
      Let VecAux0 = MLT.MGet(eqlMdlTblCurrencies, eqlMdlArray, "currency, currencies, symbol", MStrWhere:="currency_main=1")
      If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(0, "¡Tabla de Monedas sin valores iniciales!"): GoTo EH
      Let .GCurMainNameStr = VBA.CStr(VecAux0(0))
      Let .GCurMainNameStrs = VBA.CStr(VecAux0(1))
      Let .GCurMainSymStr = VBA.CStr(VecAux0(2))
      Let VecAux0 = Empty
      Let .GFormulasBoo = MLT.Formulas
    Else
      'Excliq
    End If
  End With
  
  Set RES = New AppResources_current
  Set RES.ErrorHandler = EHGLOBAL
  
  'Title, purchase and provider datas
  If Not PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasebase(PVarRequest, RES, TypCnc) Then GoTo EH

  'Cnc data to ore weight
  If Not PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal(PVarRequest, RES, TypCnc) Then GoTo EH

  'Return
  With TypCnc
    Dim PRC As ClassPurchase_current
    Set PRC = New ClassPurchase_current
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
    Set PVarResponse = PRC
    Set PRC = Nothing
  End With

  Let PurchaseConcentrate_Core_Main = True

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasebase( _
  ByRef PVarRequest As Variant, _
  ByVal RES As AppResources_current, _
  ByRef PTypCnc As GLOTYPPURCHASE) As Boolean

  Dim VecAux0 As Variant
  Dim xRows As Byte
  Dim BooParticular As Boolean

  Const STR_SIM As String = "SIMULACIÓN"
  Const STR_ND As String = "N/D"
  Const sFN As String = "APPCNC::-core_spreadmodule_generate_data_purchasebase"

  ' Request
  ' 0-BooPage
  ' 1-BooPrintTwo
  ' 2-BooNewSheet
  ' 3-RngBox
  
  On Error GoTo EH
  Let PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasebase = False

  'Title
  With PTypCnc.PTitle
    ReDim .Vector(0 To 1, 0 To 0)
    Let .Vector(0, 0) = "LIQUIDACIÓN DE CONCENTRADOS MINERALES"
    Let .Vector(1, 0) = "(EN " & VBA.UCase(PTypCnc.GCurForeNameStrs) & ")"
  End With
  
  'Get data from db (0:LIQUIDACIÓN 1:LOTE  2:FECHA 3:LUGAR 4:T/C 5:BOOTC 6:PROVEDOR  7:CÓDIGO  8:TIPO  9:CI  10:NIT 11:NIM 12:COOPERATIVA 13:MINA _
                    14:PROCEDENCIA 15:INGRESO 16:PBH 17:H2O 18:MERMA 19:Maquila 20:Base  21:Escalador)
  Let VecAux0 = MLT.MGet(eqlMdlTblPurchasebasedata, eqlMdlArray, MLngIndex:=1)
  If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(0, "¡Base de datos de liquidaciones sin información!"): GoTo EH
  Let VecAux0 = RES.ArrayToBaseZero(VecAux0, eqlRes1D)
  Let PTypCnc.GMNBoo = VBA.CBool(VecAux0(5))
  Let PTypCnc.GTCDbl = VBA.CDbl(VecAux0(4))
  
  'Purchase data
  With PTypCnc.PInvoice
    ReDim .Vector(0 To 6, 0 To 3)
    Let .Vector(0, 0) = "DATOS LIQUIDACIÓN"
    Let .Vector(1, 0) = "Destino:": Let .Vector(1, 1) = STR_SIM 'E
    Let .Vector(2, 0) = "Liquidación:": Let .Vector(2, 1) = VBA.CStr(VecAux0(0)) 'E
    Let .Vector(3, 0) = "Lote:": Let .Vector(3, 1) = VBA.CStr(VecAux0(1)) 'E
    Let .Vector(4, 0) = "Fecha:": Let .Vector(4, 1) = VBA.CDate(VBA.Format$(VBA.CStr(VecAux0(2)), "dd/mm/yy"))
    Let .Vector(5, 0) = "Lugar:": Let .Vector(5, 1) = VBA.CStr(VecAux0(3))
    Let .Vector(6, 0) = "T/C:": Let .Vector(6, 1) = VBA.Replace(VBA.CStr(VecAux0(4)), ",", ".")
    Let .Vector(6, 2) = PTypCnc.GCurMainSymStr & "/" & PTypCnc.GCurForeSymStr
  End With
  
  'Provider data
  With PTypCnc.PProvider
    Let BooParticular = (VBA.LCase(VBA.CStr(VecAux0(8))) = "particular")
    Let xRows = 9 + IIf(BooParticular, -1, 0)
    ReDim .Vector(0 To xRows, 0 To 3)
    Let .Vector(0, 0) = "DATOS PROVEDOR"
    Let .Vector(1, 0) = "Provedor:": Let .Vector(1, 1) = VBA.CStr(VecAux0(6))
    Let .Vector(2, 0) = "Código:": Let .Vector(2, 1) = VBA.CStr(VecAux0(7))
    Let .Vector(3, 0) = "Tipo:": Let .Vector(3, 1) = VBA.CStr(VecAux0(8))
    Let .Vector(4, 0) = "C.I.:": Let .Vector(4, 1) = VBA.CStr(VecAux0(9))
    Let .Vector(5, 0) = "NIT:": Let .Vector(5, 1) = VBA.CStr(VecAux0(10))
    Let .Vector(6, 0) = "NIM:": Let .Vector(6, 1) = VBA.CStr(VecAux0(11))
    If BooParticular = False Then Let .Vector(7, 0) = "Cooperativa:": Let .Vector(7, 1) = VBA.CStr(VecAux0(12))
    Let .Vector(IIf(BooParticular = False, 8, 7), 0) = "Mina:": Let .Vector(IIf(BooParticular = False, 8, 7), 1) = VBA.CStr(VecAux0(13))
    Let .Vector(IIf(BooParticular = False, 9, 8), 0) = "Procedencia:": Let .Vector(IIf(BooParticular = False, 9, 8), 1) = VBA.CStr(VecAux0(14))
  End With

  'Return
  Let PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasebase = True

EH:
  Let VecAux0 = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal( _
  ByRef PVarRequest As Variant, _
  ByVal RES As AppResources_current, _
  ByRef PTypResponse As GLOTYPPURCHASE) As Boolean

  Dim RngBox As Range

  Const sFN As String = "APPCNC::-core_spreadmodule_generate_data_purchasefinal"

  ' Request
  ' 0-BooPage
  ' 1-BooPrintTwo
  ' 2-BooNewSheet
  ' 3-RngBox
  
  On Error GoTo EH
  Let PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal = False

  'Unpack data
  If PTypResponse.GFormulasBoo Then Set RngBox = PVarRequest(3)

  With EHGLOBAL
    'Cnc Data
    'Let PTypResponse.PLngRowCounter = 3
    If Not PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getcncdata(RngBox, RES, PTypResponse) Then
      Call .ErrorHandlerSet(1, "Error en la preparación de Datos del concentrado mineral."): GoTo EH
    End If

    'Contents
    If Not PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getcnccontents(RngBox, RES, PTypResponse) Then
      Call .ErrorHandlerSet(2, "Error en la determinación de contenidos."): GoTo EH
    End If

    'Assessment
    If Not PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getassessment(RngBox, RES, PTypResponse) Then
      Call .ErrorHandlerSet(3, "Error en la valoración." & VBA.vbNewLine & .ErrorMsg): GoTo EH
    End If

    'Deductions
    If Not PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getdeductions(RngBox, RES, PTypResponse) Then
      Call .ErrorHandlerSet(4, "Error en proceso de deducciones." & VBA.vbNewLine & .ErrorMsg): GoTo EH
    End If

    'Bonus
    If Not PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getbonus(RngBox, RES, PTypResponse) Then
      Call .ErrorHandlerSet(5, "Error en el cálculo de bonos."): GoTo EH
    End If

    'Final purchase
    If Not PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getfinalpurchase(RngBox, RES, PTypResponse) Then
      Call .ErrorHandlerSet(6, "Error en el cálculo final de la Liquidación de Concentrados Minerales."): GoTo EH
    End If
  End With
  
  'Replace commas for dots
  With PTypResponse
    
    If Not PTypResponse.GFormulasBoo Then
      If .PWeight.Exists Then Call PurchaseConcentrate_Core_assets_srcarray_to_dstarray(.PWeight.Vector)
      If .PContents.Exists Then Call PurchaseConcentrate_Core_assets_srcarray_to_dstarray(.PContents.Vector)
      If .PAssessmentMain.Exists Then Call PurchaseConcentrate_Core_assets_srcarray_to_dstarray(.PAssessmentMain.Vector)
      If .PDeductions.Exists Then Call PurchaseConcentrate_Core_assets_srcarray_to_dstarray(.PDeductions.Vector)
      If .PBonus.Exists Then Call PurchaseConcentrate_Core_assets_srcarray_to_dstarray(.PBonus.Vector)
      If .PFinal.Exists Then Call PurchaseConcentrate_Core_assets_srcarray_to_dstarray(.PFinal.Vector)
    End If
  
  End With

  'Return
  Let PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal = True

EH:
  Set RngBox = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getcncdata( _
  ByVal PRngBox As Range, _
  ByVal RES As AppResources_current, _
  ByRef PTypRes As GLOTYPPURCHASE) As Boolean

  Dim VecAux0 As Variant
  Dim DatIncoming As Date
  Dim DblH2O As Double, DblDecrease As Double, DblPBS As Double
  Dim StrPBS As String, StrPNS As String
  Dim xCols As Long

  Const sFN As String = "APPCNC::-core_main_spreadmodule_generate_data_purchasebase_getcncdata"

  On Error GoTo EH
  Let PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getcncdata = False

  With PTypRes.PWeight
    
    Let xCols = 5
    
    'Get data
    Let VecAux0 = MLT.MGet(eqlMdlTblPurchasebasedata, eqlMdlArray, "INGRESO, PBH, H2O, MERMA", 1)
    If VBA.IsArray(VecAux0) Then GoTo EH
    Let VecAux0 = RES.ArrayToBaseZero(VecAux0, eqlRes1D)
    Let DatIncoming = VBA.CDate(VecAux0(0))
    Let PTypRes.TotalPBHDbl = VBA.CDbl(VecAux0(1))
    Let DblH2O = VBA.CDbl(VecAux0(2))
    Let DblDecrease = VBA.CDbl(VecAux0(3))
    Let VecAux0 = Empty
    
    'Fill main vector with values
    ReDim .Vector(0 To 0, 0 To xCols)
    Let .Vector(0, 0) = "DATOS CONCENTRADO"
    Let .Vector = RES.ArrayAddAtLast(.Vector, Array("Ingreso", "PBH [T]", "Humedad [%]", "PBS [T]", "MERMA [%]", "PNS [T]"))
    
    If PTypRes.GFormulasBoo Then
      Let StrPBS = "=" & PRngBox.Range("G6").Address(False, False) & "*(1-(" & PRngBox.Range("H6").Address(False, False) & "/100))" ' PBS
      Let StrPNS = "=" & PRngBox.Range("I6").Address(False, False) & "*(1-(" & PRngBox.Range("J6").Address(False, False) & "/100))"   ' PNS
      Let .Vector = RES.ArrayAddAtLast(.Vector, Array(VBA.Format(VBA.CDate(DatIncoming), "dd/mm/yy"), VBA.Replace(PTypRes.TotalPBHDbl, ",", "."), VBA.Replace(DblH2O, ",", "."), StrPBS, VBA.Replace(DblDecrease, ",", "."), StrPNS))
      Let .StrRngA1 = "F6"
    Else
      Let DblPBS = (PTypRes.TotalPBHDbl * (1 - (DblH2O / 100))) ' PBS
      Let PTypRes.TotalPBSDbl = (DblPBS * (1 - (DblH2O / 100))) * (1 - (DblDecrease / 100)) ' PNS
      Let .Vector = RES.ArrayAddAtLast(.Vector, Array(DatIncoming, PTypRes.TotalPBHDbl, DblH2O, DblPBS, DblDecrease, PTypRes.TotalPBSDbl))
    End If
    
    If Not VBA.IsArray(.Vector) Then GoTo EH
  
  End With
  
  'Return
  Let PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getcncdata = True
  
EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getcnccontents( _
  ByVal PRngBox As Range, _
  ByVal RES As AppResources_current, _
  ByRef PTypRes As GLOTYPPURCHASE) As Boolean

  Dim VecAux0 As Variant, VecAux1 As Variant
  Dim StrCont As String, StrGrad As String, StrCotz As String
  Dim i As Long, xCols As Long

  Const sFN As String = "APPCNC::-core_spreadmodule_generate_data_purchasebase_getcnccontents"

  On Error GoTo EH
  Let PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getcnccontents = False

  With PTypRes
    
    ' Get data
    Let VecAux0 = MLT.MGet(eqlMdlTblConpurcontents, eqlMdlArray, "CONTENIDOS, LEY, UNIDAD")
    If Not VBA.IsArray(VecAux0) Then GoTo EH
    
    ReDim .UnitsGraVec(UBound(VecAux0)) ' Unit Gra
    ReDim .UnitsCtzVec(UBound(VecAux0)) ' Unit Ctz

    ' Fill matrix
    With .PContents
      
      Let xCols = 5
      Let .LngRngA1 = 9 + UBound(VecAux0)
      ReDim .Vector(0 To 0, 0 To xCols)
      Let .Vector(0, 0) = "DATOS DE CONTENIDOS EN CONCENTRADO"
      Let .Vector = RES.ArrayAddAtLast(.Vector, Array("Contenidos", Empty, Empty, Empty, "Ley", "Cotización"))
      
      For i = 0 To UBound(VecAux0)
        Let VecAux1 = MLT.MGet(eqlMdlTblConfigpurrm, eqlMdlArray, "COTIZACIÓN, UNIDAD", MStrWhere:="MINERAL=" & StrCont)
        If Not VBA.IsArray(VecAux1) Then GoTo EH
        Let VecAux1 = RES.ArrayToBaseZero(VecAux1, eqlRes1D)
        Let PTypRes.UnitsGraVec(i) = VecAux0(2) ' Unit Gra
        Let PTypRes.UnitsCtzVec(i) = VecAux1(1) ' Unit Ctz
        Let .Vector = RES.ArrayAddAtLast(.Vector, Array(VBA.StrConv(VBA.CStr(VecAux1(0)), vbProperCase), Empty, Empty, Empty, VecAux0(1), VecAux1(0)))
        Let VecAux1 = Empty
      Next i
      
      If Not VBA.IsArray(.Vector) Then GoTo EH
      If PTypRes.GFormulasBoo Then Let .StrRngA1 = "F9"
    
    End With
  
  End With
  
  'Return
  Let PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getcnccontents = True
  
EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getassessment( _
  ByVal PRngBox As Range, _
  ByVal RES As AppResources_current, _
  ByRef PTypRes As GLOTYPPURCHASE) As Boolean

  Dim i As Long, j As Long, k As Long, xRows As Long, xCols As Long, xRowFirstVal As Long, xRowResize As Long, xRowTotals As Long

  Const sFN As String = "APPCNC::-core_spreadmodule_generate_data_purchasebase_getassessment"

  On Error GoTo EH
  Let PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getassessment = False

  'Get all values from assessment methods
  With EHGLOBAL
    
    Let PTypRes.TotalAssessmentDbl = 0
    
    'Assessment
    If Not PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getassessment_value(PRngBox, RES, PTypRes) Then
      Call .ErrorHandlerSet(3.1, "Falló el cálculo de valoración del concentrado."): GoTo EH
    End If

    'Treatment cost
    If Not PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getassessment_treatmentcost(PRngBox, RES, PTypRes) Then
      Call .ErrorHandlerSet(3.2, "Falló el cálculo del costo de tratamiento."): GoTo EH
    End If

    'Refinement
    If Not PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getassessment_refinement(PRngBox, RES, PTypRes) Then
      Call .ErrorHandlerSet(3.3, "Falló el cálculo del costo de refinación."): GoTo EH
    End If

    'Penalties
    If Not PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getassessment_penalties(PRngBox, RES, PTypRes) Then
      Call .ErrorHandlerSet(3.4, "Falló el cálculo de penalidades."): GoTo EH
    End If
  
  End With
  
  ' Fill Assessment matrix with results
  Dim DblTotalPerTon As Double
  Dim LngResize As Long
  Dim BooAssessmetsExists As Boolean
  
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
      
      If PTypRes.GFormulasBoo Then
        Dim StrPerTonVal As String, StrPerTonRng As String
        Dim LngFirstA1 As Long
        Let LngFirstA1 = 9 + (UBound(PTypRes.PContents.Vector) - 2) + (UBound(PTypRes.PAssessment.Vector) + 1)  ' First Main Assessment row address
        Let StrPerTonRng = "F" & LngFirstA1 + IIf(LngResize <= 0, 1, LngResize + 2) ' Main Assessment per ton address
        'Let PTypRes.TotalAssessmentRng = "F" & LngFirstA1 + IIf(LngResize <= 0, 2, LngResize + 3) ' Main Assessment address
        Let StrPerTonVal = "=" & PRngBox.Range(PTypRes.PAssessment.StrRngA1).Offset(UBound(PTypRes.PAssessment.Vector) - 2, 5).Address(False, False)
        If LngResize > 0 Then Let StrPerTonVal = StrPerTonVal & "-SUM(" & PRngBox.Range("F" & LngFirstA1).Offset(2, 5).Resize(LngResize, 1).Address(False, False) & ")"
        Let .Vector = RES.ArrayConcat(.Vector, Array("Total valor bruto por tonelada:", Empty, Empty, Empty, Empty, StrPerTonVal))
        Let StrPerTonVal = "=" & PRngBox.Range(StrPerTonRng).Address(False, False)
        Let StrPerTonVal = StrPerTonVal & "*" & PRngBox.Range("K6").Address(False, False)
        Let .Vector = RES.ArrayConcat(.Vector, Array("TOTAL VALOR BRUTO [" & PTypRes.GCurForeSymStr & "]:", Empty, Empty, Empty, Empty, StrPerTonVal))
      Else
        Let DblTotalPerTon = PTypRes.TotalAssessmentDbl
        Let PTypRes.TotalAssessmentDbl = PTypRes.TotalAssessmentDbl * VBA.CDbl(PTypRes.PWeight.Vector(2, 5)) ' Main Assessment Value
        Let .Vector = RES.ArrayConcat(.Vector, Array("Total valor bruto por tonelada:", Empty, Empty, Empty, Empty, DblTotalPerTon))
        Let .Vector = RES.ArrayConcat(.Vector, Array("TOTAL VALOR BRUTO [" & PTypRes.GCurForeSymStr & "]:", Empty, Empty, Empty, Empty, PTypRes.TotalAssessmentDbl))
      End If
      Let .LngRngA1 = (9 + (UBound(PTypRes.PContents.Vector) - 2)) + (UBound(.Vector) + 1)
    
    End With
    Let .PAssessment.Vector = Empty
    Let .PTreatmentCost.Vector = Empty
    Let .PRefinement.Vector = Empty
    Let .PPenalties.Vector = Empty
    If Not VBA.IsArray(.PAssessmentMain.Vector) Then GoTo EH
    Let .TotalDeductionsRng = "K" & (9 + (UBound(.PContents.Vector) - 2)) + (UBound(.PAssessmentMain.Vector) + 1) ' Main Assessment address
  
  End With
  
  'Return
  Let PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getassessment = True

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getassessment_value( _
  ByVal PRngBox As Range, _
  ByVal RES As AppResources_current, _
  ByRef PTypRes As GLOTYPPURCHASE) As Boolean
  
  Dim RngBox As Range
  
  Dim VecAux1 As Variant
  Dim StrAux0 As String, StrAux1 As String, StrGrade As String, StrCotiz As String, StrValue As String, StrGradeC As String, StrCotizC As String
  Dim DblAux0 As Double, DblAux1 As Double, DblAux2 As Double, DblGrade As Double, DblCotiz As Double, DblValue As Double
  Dim i As Long, xCols As Long, xRowResize As Long

  Const sFN As String = "APPCNC::-core_spreadmodule_generate_data_purchasebase_getassessment_value"

  On Error GoTo EH
  Let PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getassessment_value = False

  With PTypRes
    
    ' Get data (0:CONTENIDOS-1:LEY-2:DEDUC. MÍNIMA [u]-3:Y/O-4:PAGABLE AL [%]-5:UNIDAD)
    Dim VecAux0 As Variant
    Let VecAux0 = MLT.MGet(eqlMdlTblConpurbase, eqlMdlArray)
    If Not VBA.IsArray(VecAux0) Then GoTo EH
    Let VecAux0 = RES.ArrayToBaseZero(VecAux0, eqlRes1D)
    
    ReDim .UnitsTraVec(UBound(VecAux0))
    
    ' Calculate values
    With .PAssessment
      
      Let xCols = 5
      Let xRowResize = UBound(VecAux0) + 1
      Let .LngRngA1 = PTypRes.PContents.LngRngA1 + 3
      Let .StrRngA1 = "F" & .LngRngA1
      If PTypRes.GFormulasBoo Then Set RngBox = PRngBox.Range(.StrRngA1)
      Let .LngRngA1 = .LngRngA1 + xRowResize
      
      Let DblAux2 = 0
      ReDim .Vector(0 To 0, 0 To xCols)
      Let .Vector(0, 0) = "VALORACIÓN"
      Let .Vector(0, 5) = "Valor [" & PTypRes.GCurForeSymStr & "/T]"
      Let .Vector = RES.ArrayAddAtLast(.Vector, Array("Mineral", "Ley", "Deducción", "Pagable [%]", "Cotización"), "Valor")
      
      For i = 0 To UBound(VecAux0)
        ' Get all data from RM Table: (0:MINERAL-1:ELEMENTO-2:UNIDAD-3:COTIZACIÓN-4:EXPORTACIÓN-5:VENTA INTERNA-6:UNIDAD MINERAL-7:CONVERSIÓN(OJO)-8:FACTOR-9:FACTOR LOCAL)
        Let VecAux1 = MLT.MGet(eqlMdlTblConfigpurrm, eqlMdlArray, MStrWhere:="MINERAL=" & VBA.CStr(VecAux0(i)))
        If Not VBA.IsArray(VecAux1) Then GoTo EH
        Let VecAux1 = RES.ArrayToBaseZero(VecAux1, eqlRes1D)
        
        Let PTypRes.UnitsTraVec(i) = VecAux0(2) ' Set transaction units
        If PTypRes.GFormulasBoo Then
          Let StrGrade = "=" & RngBox.Offset(i, 4).Address(False, False) & "*" & VBA.Replace(VBA.CStr(VecAux1(11)), ",", ".") ' 1.  Conversion of mineral units
          Let StrCotiz = "=" & RngBox.Offset(i, 5).Address(False, False) & "*" & VBA.Replace(VBA.CStr(VecAux1(10)), ",", ".") ' 2.  Conversion of prices units
          ' 3.  Calculate final value
          Let StrAux0 = RngBox.Offset(i, 2).Address(False, False) ' Deduction
          Let StrAux1 = RngBox.Offset(i, 3).Address(False, False) ' %Payable
          Let StrGradeC = RngBox.Offset(i, 1).Address(False, False) ' Grade
          Let StrCotizC = RngBox.Offset(i, 4).Address(False, False) ' Cotiz
          If VBA.LCase(VBA.CStr(VecAux1(7))) = "ot" Then ' =IF(Gra=0;Gra-ABS(Ded);IF(Ded=0;Gra*Pay/100;(Gra-ABS(Ded))))*(Ctz*Pay/100)
            Let StrValue = "=IF(" & StrGradeC & "=0," & StrGradeC & "-ABS(" & StrAux0 & "),IF(" & StrAux0 & "=0," & StrGradeC & "*" & StrAux1 & "*0.01,(" & StrGradeC & "-ABS(" & StrAux0 & ")))*" & StrAux1 & ")*(" & StrCotizC & "*0.01)"
          Else ' =IF(OR(Gra=0;Gra>=100);Gra-ABS(Ded);IF(Ded=0;Gra*Pay/100;IF(Gra>(Ded/((Pay/100)-1));Gra*Pay/100;Gra-ABS(Ded))))*(0,01)*(Ctz)
            Let StrValue = "=IF(OR(" & StrGradeC & "=0," & StrGradeC & ">=100)," & StrGradeC & "-ABS(" & StrAux0 & "),IF(" & StrAux0 & "=0," & StrGradeC & "*" & StrAux1 & "*0.01,IF(" & StrGradeC & ">(" & StrAux0 & "/((" & StrAux1 & "*0.01)-1))," & StrGradeC & "*" & StrAux1 & "*0.01," & StrGradeC & "-ABS(" & StrAux0 & "))))*(" & StrCotizC & "*0.01)"
          End If
          Let .Vector = RES.ArrayAddAtLast(.Vector, Array(VBA.StrConv(VBA.CStr(VecAux1(1)), vbProperCase), StrGrade, StrAux0, StrAux1, StrCotiz, StrValue))
        Else
          Let DblGrade = VBA.CDbl(VecAux0(1)) * VBA.CDbl(VecAux1(11)) ' 1.  Conversion of mineral units
          Let DblCotiz = VBA.CDbl(VecAux1(3)) * VBA.CDbl(VecAux1(10)) ' 2.  Conversion of prices units
          ' 3.  Calculate final value
          Let DblAux0 = VBA.CDbl(VecAux0(2)) ' Deduction
          Let DblAux1 = VBA.CDbl(VecAux0(3)) ' %Payable
          If VBA.LCase(VBA.CStr(VecAux1(7))) = "ot" Then ' =IF(Gra=0;Gra-ABS(Ded);IF(Ded=0;Gra*Pay/100;(Gra-ABS(Ded))))*(Ctz*Pay/100)
            Let DblValue = IIf(DblGrade = 0, DblGrade - VBA.Abs(DblAux0), IIf(DblAux0 = 0, DblGrade * DblAux1 * 0.01, (DblGrade - VBA.Abs(DblAux0))) * DblAux1) * (DblCotiz * 0.01)
          Else ' =IF(OR(Gra=0;Gra>=100);Gra-ABS(Ded);IF(Ded=0;Gra*Pay/100;IF(Gra>(Ded/((Pay/100)-1));Gra*Pay/100;Gra-ABS(Ded))))*(0,01)*(Ctz)
            Let DblValue = IIf(DblGrade = 0 Or DblGrade >= 100, DblGrade - VBA.Abs(DblAux0), IIf(DblAux0 = 0, DblGrade * DblAux1 * 0.01, IIf(DblGrade > (DblAux0 / ((DblAux1 * 0.01) - 1)), DblGrade * DblAux1 * 0.01, DblGrade - VBA.Abs(DblAux0)))) * 0.01 * DblCotiz
            Let DblAux2 = DblAux2 + DblValue
          End If
          Let .Vector = RES.ArrayAddAtLast(.Vector, Array(VBA.StrConv(VBA.CStr(VecAux1(1)), vbProperCase), DblGrade, DblAux0, DblAux1, DblCotiz, DblValue))
          Let PTypRes.TotalAssessmentDbl = PTypRes.TotalAssessmentDbl + DblValue
        End If
      
      Next i
      If Not VBA.IsArray(.Vector) Then GoTo EH
      Let VecAux1 = Empty
        
      'Totals row
      If PTypRes.GFormulasBoo Then Let StrValue = "=SUM(" & RngBox.Offset(i, 5).Resize(xRowResize, 1).Address(False, False) & ")"
      Let .Vector = RES.ArrayAddAtLast(.Vector, Array("Total Valoración", Empty, Empty, Empty, Empty, IIf(PTypRes.GFormulasBoo, StrValue, DblAux2)))
    
    End With
  
  End With

  'Return
  Let PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getassessment_value = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Set RngBox = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getassessment_treatmentcost( _
  ByVal PRngBox As Range, _
  ByVal RES As AppResources_current, _
  ByRef PTypRes As GLOTYPPURCHASE) As Boolean
  
  Dim RngBox As Range
  
  Const sFN As String = "APPCNC::-core_spreadmodule_generate_data_purchasefinal_getassessment_treatmentcost"

  On Error GoTo EH
  Let PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getassessment_treatmentcost = False

  With PTypRes
    
    'Get data
    Dim VecAux0 As Variant
    Let VecAux0 = MLT.MGet(eqlMdlTblPurchasebasedata, eqlMdlArray, "Maquila, Base, Escalador", 1)
    Let .PTreatmentCost.Exists = VBA.IsArray(VecAux0)
    If .PTreatmentCost.Exists Then Let VecAux0 = RES.ArrayToBaseZero(VecAux0, eqlRes1D)
  
    ' Fill matrix
    With .PTreatmentCost
      
      Let .LngRngA1 = PTypRes.PAssessment.LngRngA1
      If .Exists Then
        Dim xCols As Long
        Let xCols = 5
        Let .LngRngA1 = .LngRngA1 + 2
        Let .StrRngA1 = "F" & .LngRngA1
        If PTypRes.GFormulasBoo Then Set RngBox = PRngBox.Range(.StrRngA1)
        Let .LngRngA1 = .LngRngA1 + 2
        
        ' Fill final vector with data
        ReDim .Vector(0 To 0, 0 To xCols)
        Let .Vector(0, 0) = "COSTO TRATAMIENTO"
        Let .Vector = RES.ArrayAddAtLast(.Vector, Array("Maquila", Empty, Empty, Empty, Empty, VecAux0(0)))
        Let .Vector = RES.ArrayAddAtLast(.Vector, Array("Ajuste", "Actual", "Base", "Diferencia", "Escalador", Empty))
  
        ' Calculate results
        If PTypRes.GFormulasBoo Then
          Dim StrCur As String, StrDif As String, StrVal As String
          Let StrCur = "=" & PRngBox.Range(PTypRes.PAssessment.StrRngA1).Offset(0, 4).Address(False, False) ' Actual
          Let StrDif = "=" & RngBox.Offset(2, 1).Address(False, False) & "-" & RngBox.Offset(2, 2).Address(False, False) ' Diferencia
          Let StrVal = "=IF(" & RngBox.Offset(2, 3).Address(False, False) & ">0," & RngBox.Offset(2, 3).Address(False, False) & "*" & RngBox.Offset(2, 4).Address(False, False) & ",0)"
          Let .Vector = RES.ArrayAddAtLast(.Vector, Array(Empty, StrCur, VBA.Replace(VBA.CStr(VecAux0(1)), ",", "."), StrDif, VBA.Replace(VBA.CStr(VecAux0(2)), ",", "."), StrVal))
        Else
          Dim DblCur As Double, DblDif As Double, DblVal As Double
          Let DblCur = VBA.CDbl(PTypRes.PAssessment.Vector(2, 4)) ' Actual
          Let DblDif = (DblCur - VBA.CDbl(VecAux0(1))) ' Diferencia
          Let DblVal = IIf(DblDif > 0, DblDif * VBA.CDbl(VecAux0(2)), 0) ' Result
          Let .Vector = RES.ArrayAddAtLast(.Vector, Array(Empty, DblCur, VecAux0(1), DblDif, VecAux0(2), DblVal))
          Let PTypRes.TotalAssessmentDbl = PTypRes.TotalAssessmentDbl - DblVal
        End If
        If Not VBA.IsArray(.Vector) Then GoTo EH
      End If

    End With
  
  End With

  'Return
  Let PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getassessment_treatmentcost = True

EH:
  Let VecAux0 = Empty
  Set RngBox = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getassessment_refinement( _
  ByVal PRngBox As Range, _
  ByVal RES As AppResources_current, _
  ByRef PTypRes As GLOTYPPURCHASE) As Boolean

  Dim RngBox As Range
  
  Const sFN As String = "APPCNC::-core_spreadmodule_generate_data_purchasefinal_getassessment_refinement"

  On Error GoTo EH
  Let PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getassessment_refinement = False

  With PTypRes
    
    'Get data (0-CONCEPTO 1-BASE 2-CONDICIÓN [%])
    Dim VecAux0 As Variant
    Let VecAux0 = MLT.MGet(eqlMdlTblConpurrefination, eqlMdlArray)
    Let .PRefinement.Exists = VBA.IsArray(VecAux0)
    If .PRefinement.Exists Then Let VecAux0 = RES.ArrayToBaseZero(VecAux0, eqlRes1D)
    
    ' Fill matrix
    With .PRefinement
      
      Let .LngRngA1 = PTypRes.PTreatmentCost.LngRngA1
      If .Exists Then
        Dim xCols As Long
        Let xCols = 5
        Let .LngRngA1 = .LngRngA1 + 3
        Let .StrRngA1 = "F" & .LngRngA1
        If PTypRes.GFormulasBoo Then Set RngBox = PRngBox.Range(.StrRngA1)
        Let .LngRngA1 = .LngRngA1 + UBound(VecAux0)
        
        ' Fill final vector with data
        ReDim .Vector(0 To 0, 0 To xCols)
        Let .Vector(0, 0) = "COSTO REFINACIÓN"
        Let .Vector = RES.ArrayAddAtLast(.Vector, Array("Concepto", Empty, "Base", "Condición[%]", "Pagable", Empty))
  
        ' Calculate results
        Dim StrGra As String, StrDed As String, StrPay As String, StrVal As String, StrAux0 As String
        Dim DblGra As Double, DblDed As Double, DblPay As Double, DblVal As Double
        Dim i As Long
        For i = 0 To UBound(VecAux0)
          Let StrAux0 = VBA.LCase(VBA.CStr(PTypRes.UnitsGraVec(i)))
          
          If PTypRes.GFormulasBoo Then
            
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
            Let .Vector = RES.ArrayAddAtLast(.Vector, Array(VBA.StrConv(VBA.CStr(VecAux0(0)), vbProperCase), Empty, VBA.Replace(VBA.CStr(VecAux0(1)), ",", "."), VBA.Replace(VBA.CStr(VecAux0(2)), ",", "."), StrPay, StrVal))
          
          Else
            
            Let DblGra = VBA.CDbl(PTypRes.PAssessment.Vector(i, 1)) ' Get Mineral Grades
            Let DblDed = VBA.CDbl(PTypRes.PAssessment.Vector(i, 2)) ' Get Deduction
            Let DblPay = VBA.CDbl(PTypRes.PAssessment.Vector(i, 3)) ' Get %Payable
            If StrAux0 = "ot" Then ' =IF(Gra=0;Gra-ABS(Ded);IF(Ded=0;Gra*Pay/100;(Gra-ABS(Ded))))
              Let DblPay = IIf(DblGra = 0, DblGra - VBA.Abs(StrDed), IIf(DblDed = 0, DblGra * DblPay * 0.01, (DblGra - Abs(DblDed))) * DblPay)
            Else ' =IF(OR(Gra=0;Gra>=100);Gra-ABS(Ded);IF(Ded=0;Gra*Pay/100;IF(Gra>(Ded/((Pay/100)-1));Gra*Pay/100;Gra-ABS(Ded))))*(0,01*2204.6223)
              Let DblPay = IIf(DblGra = 0 Or DblGra >= 100, DblGra - VBA.Abs(DblDed), IIf(DblDed = 0, DblGra * DblPay * 0.01, IIf(DblGra > (DblDed / ((DblPay * 0.01) - 1)), DblGra * DblPay * 0.01, DblGra - VBA.Abs(DblDed)))) * (IIf(StrAux0 = "%", 2204.6223, 1) * 0.01)
            End If
            ' =IF(CTZ*CND/100>BSE;PAY*CTZ*CND/100;PAY*BSE)
            Let DblVal = IIf(VBA.CDbl(PTypRes.PAssessment.Vector(i + 2, 4)) * VBA.CDbl(VecAux0(2)) * 0.01 > VBA.CDbl(VecAux0(1)), VBA.CDbl(PTypRes.PAssessment.Vector(i + 2, 4)) * VBA.CDbl(VecAux0(2)) * DblPay * 0.01, VBA.CDbl(VecAux0(1)) * DblPay)
            Let .Vector = RES.ArrayAddAtLast(.Vector, Array(VBA.StrConv(VBA.CStr(VecAux0(0)), vbProperCase), Empty, VecAux0(1), VecAux0(2), DblPay, DblVal))
            Let PTypRes.TotalAssessmentDbl = PTypRes.TotalAssessmentDbl - DblVal
          
          End If
  
        Next i
        If Not VBA.IsArray(.Vector) Then GoTo EH
      End If

    End With
  
  End With

  'Return
  Let PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getassessment_refinement = True

EH:
  Let VecAux0 = Empty
  Set RngBox = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getassessment_penalties( _
  ByVal PRngBox As Range, _
  ByVal RES As AppResources_current, _
  ByRef PTypRes As GLOTYPPURCHASE) As Boolean

  Dim RngBox As Range
  
  Const sFN As String = "APPCNC::-core_spreadmodule_generate_data_purchasefinal_getassessment_penalties"

  On Error GoTo EH
  Let PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getassessment_penalties = False

  With PTypRes
    
    ' Get data (0-ELEMENTO 1-ACTUAL 2-BASE 3-DEDUCCIÓN 4-LÍMITE 5-SELECTED)
    Dim VecAux0 As Variant
    Let VecAux0 = MLT.MGet(eqlMdlTblConpurpenalties, eqlMdlArray, "ELEMENTO, ACTUAL, BASE, DEDUCCIÓN, LÍMITE", MStrWhere:="SELECTED=1")
    Let .PRefinement.Exists = VBA.IsArray(VecAux0)
    
    ' Fill matrix
    With .PRefinement
      
      Let .LngRngA1 = PTypRes.PRefinement.LngRngA1
      If .Exists Then
        Dim xCols As Long
        Let xCols = 5
        Let .LngRngA1 = .LngRngA1 + 3
        Let .StrRngA1 = "F" & .LngRngA1
        If PTypRes.GFormulasBoo Then Set RngBox = PRngBox.Range(.StrRngA1)
        Let .LngRngA1 = .LngRngA1 + UBound(VecAux0)
        
        ' Fill final vector with data
        ReDim .Vector(0 To 0, 0 To xCols)
        Let .Vector(0, 0) = "PENALIDADES"
        Let .Vector = RES.ArrayAddAtLast(.Vector, Array("Elemento", "Actual", "Base", "[" & PTypRes.GCurForeSymStr & "]", "Límite", Empty))
        
        ' Calculate results
        Dim StrCur As String, StrBse As String, StrPay As String, StrLim As String, StrVal As String
        Dim DblCur As Double, DblBse As Double, DblPay As Double, DblLim As Double, DblVal As Double
        Dim i As Long
        For i = 0 To UBound(VecAux0)
          
          If PTypRes.GFormulasBoo Then
            
            Let StrCur = RngBox.Offset(i, 1).Address(False, False) ' Current
            Let StrBse = RngBox.Offset(i, 2).Address(False, False) ' Base
            Let StrPay = RngBox.Offset(i, 3).Address(False, False) ' Pay
            Let StrLim = RngBox.Offset(i, 4).Address(False, False) ' Limit
            ' Calculate final value: =IF((Cur-Bse)>0;(Cur-Bse)*Pay/Lim;0)
            Let StrVal = "=IF(" & StrCur & "-" & StrBse & ">0," & "(" & StrCur & "-" & StrBse & ")*" & StrPay & "/" & StrLim & ",0)"
            Let .Vector = RES.ArrayAddAtLast(.Vector, Array(VBA.StrConv(VBA.CStr(VecAux0(i, 0)), vbProperCase), VBA.Replace(VBA.CStr(VecAux0(i, 1)), ",", "."), VBA.Replace(VBA.CStr(VecAux0(i, 2)), ",", "."), VBA.Replace(VBA.CStr(VecAux0(i, 3)), ",", "."), VBA.Replace(VBA.CStr(VecAux0(i, 4)), ",", "."), StrVal))
          
          Else
            
            Let DblCur = VBA.CDbl(VecAux0(i, 1)) ' Current
            Let DblBse = VBA.CDbl(VecAux0(i, 2)) ' Base
            Let DblPay = VBA.CDbl(VecAux0(i, 3)) ' Pay
            Let DblLim = VBA.CDbl(VecAux0(i, 4)) ' Limit
            ' Calculate final value: =IF((Cur-Bse)>0;(Cur-Bse)*Pay/Lim;0)
            Let DblVal = IIf(DblCur - DblBse > 0, (DblCur - DblBse) * DblPay / DblLim, 0)
            Let .Vector = RES.ArrayAddAtLast(.Vector, Array(VBA.StrConv(VBA.CStr(VecAux0(i, 0)), vbProperCase), DblCur, DblBse, DblPay, DblLim, DblVal))
            Let PTypRes.TotalAssessmentDbl = PTypRes.TotalAssessmentDbl - DblVal
          
          End If
  
        Next i
        If Not VBA.IsArray(.Vector) Then GoTo EH
      End If

    End With
  
  End With

  'Return
  Let PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getassessment_penalties = True

EH:
  Let VecAux0 = Empty
  Set RngBox = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getdeductions( _
  ByVal PRngBox As Range, _
  ByRef RES As AppResources_current, _
  ByRef PTypRes As GLOTYPPURCHASE) As Boolean


  Const sFN As String = "APPCNC::-core_spreadmodule_generate_data_purchasebase_getdeductions"

  On Error GoTo EH
  Let PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getdeductions = False

  'Get all values from deductions methods
  With EHGLOBAL
  
    Let PTypRes.TotalDeductionsDbl = 0
    
    'Exports discounts
    If Not PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getdeductions_exports(PRngBox, RES, PTypRes) Then
      Call .ErrorHandlerSet(4.1, "Falló el cálculo de deducciones de exportación."): GoTo EH
    End If

    'RM Tax
    If Not PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getdeductions_rmtaxes(PRngBox, RES, PTypRes) Then
      Call .ErrorHandlerSet(4.2, "Falló el cálculo de regalías."): GoTo EH
    End If

    'Taxes by Law
    If Not PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getdeductions_lawtaxes(PRngBox, RES, PTypRes) Then
      Call .ErrorHandlerSet(4.3, "Falló el cálculo de retenciones de ley."): GoTo EH
    End If

    'Taxes others
    If Not PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getdeductions_othertaxes(PRngBox, RES, PTypRes) Then
      Call .ErrorHandlerSet(4.4, "Falló el cálculo en retenciones otras."): GoTo EH
    End If

    'Discounts others
    If Not PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getdeductions_discounts(PRngBox, RES, PTypRes) Then
      Call .ErrorHandlerSet(4.5, "Falló el cálculo en descuentos otros."): GoTo EH
    End If
  
  End With

  ' Fill deductions matrix with results
  Dim LngResize As Long
  
  With PTypRes
    
    Let .PDeductions.LngRngA1 = 0
    Let .PDeductions.Exists = (.PDExports.Exists Or .PDRMRets.Exists Or .PDLawRets.Exists Or .PDOthRets.Exists Or .PDDiscnts.Exists)
    If .PDeductions.Exists Then
      ' Create Deductions vector
      ReDim .PDeductions.Vector(0 To 1, 0 To 5)
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
        
        If PTypRes.GFormulasBoo Then
          Let LngResize = UBound(.Vector) - 3 '3: Title, subtitle, subtitle content
          'Let PTypRes.TotalDeductionsRng = "F" & 9 + (UBound(PTypRes.PContents.Vector) - 2) + UBound(PTypRes.PAssessmentMain.Vector) + LngResize + 1 ' Main Deductions address
          Let .StrRngA1 = "F" & 9 + (UBound(PTypRes.PContents.Vector) - 2) + UBound(PTypRes.PAssessmentMain.Vector) + 4 ' Deductions address
          Let .Vector = RES.ArrayConcat(.Vector, Array("TOTAL DEDUCCIONES [" & PTypRes.GCurForeSymStr & "]:", Empty, Empty, Empty, Empty, "=SUM(" & PRngBox.Range(.StrRngA1).Offset(0, 5).Resize(LngResize, 1).Address(False, False) & ")"))
        Else
          Let .Vector = RES.ArrayConcat(.Vector, Array("TOTAL DEDUCCIONES [" & PTypRes.GCurForeSymStr & "]:", Empty, Empty, Empty, Empty, PTypRes.TotalDeductionsDbl))
        End If
        Let .LngRngA1 = PTypRes.PAssessment.LngRngA1 + (UBound(.Vector) + 1)
      
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
  
  'Return
  Let PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getdeductions = True

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getdeductions_exports( _
  ByVal PRngBox As Range, _
  ByRef RES As AppResources_current, _
  ByRef PTypRes As GLOTYPPURCHASE) As Boolean

  Dim RngBox As Range
  
  Dim VecAux0 As Variant, VecAux1 As Variant

  Const sFN As String = "APPCNC::-core_spreadmodule_generate_data_purchasebase_getdeductions_exports"

  On Error GoTo EH
  Let PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getdeductions_exports = False

  With PTypRes

    ' Get data (0.CONCEPTO 1.PORCENTAJE 2.VALOR 3.VALOR DIRECTO 4.SELECTED)
    Let VecAux0 = MLT.MGet(eqlMdlTblConpurdedexp, eqlMdlArray, "CONCEPTO, PORCENTAJE, VALOR, VALOR DIRECTO", MStrWhere:="SELECTED=1")
    Let .PDExports.Exists = VBA.IsArray(VecAux0)
    
    ' Fill matrix
    With .PDExports
      
      Let .LngRngA1 = PTypRes.PAssessmentMain.LngRngA1
      If .Exists Then
        'Check for Resume tables boolean
        Dim i As Long
        If PTypRes.GResumedExportsBoo Then
          Dim BooDirect As Boolean, BooIndirect As Boolean
          Let VecAux1 = VecAux0
          Let VecAux0 = Empty
          Let BooDirect = False
          Let BooIndirect = False
          ReDim VecAux0(0 To 0, 1 To UBound(VecAux1, 2))
          For i = 0 To UBound(VecAux1)
            Let BooDirect = IIf(BooDirect, BooDirect, (VBA.CBool(VecAux1(i, 3)) = True))
            Let BooIndirect = IIf(BooIndirect, BooIndirect, (VBA.CBool(VecAux1(i, 3)) = False))
            If VBA.CBool(VecAux1(i, 3)) Then
              Let VecAux0(0, 1) = VBA.CDbl(VecAux0(0, 1)) + VBA.CDbl(VecAux1(i, 1))
            Else
              Let VecAux0(0, 2) = VBA.CDbl(VecAux0(0, 2)) + VBA.CDbl(VecAux1(i, 1))
            End If
          Next i
          Let VecAux1 = Empty
          Let VecAux0(0, 0) = "Descuentos directos varios"
          Let VecAux0(1, 0) = "Descuentos varios"
          Let VecAux0(0, 3) = 1
          Let VecAux0(1, 3) = 0
          If Not BooDirect Then Let VecAux0 = RES.ArrayPop(VecAux0)
          If Not BooIndirect Then Let VecAux0 = RES.ArrayShift(VecAux0)
        End If
        
        'Determine Addresses and row numbers
        Dim xCols As Long
        Let xCols = 5
        Let .LngRngA1 = .LngRngA1 + 4
        Let .StrRngA1 = "F" & .LngRngA1
        If PTypRes.GFormulasBoo Then Set RngBox = PRngBox.Range(.StrRngA1)
        Let .LngRngA1 = .LngRngA1 + UBound(VecAux0)
        
        ' Fill final vector with data
        ReDim .Vector(0 To 0, 0 To xCols)
        Let .Vector(0, 0) = "DESCUENTOS POR EXPORTACIÓN"
        
        ' Calculate results
        Dim StrPer As String, StrCsh As String, StrTMH As String, StrAmt As String, StrVal As String
        Dim DblPer As Double, DblCsh As Double, DblTMH As Double, DblAmt As Double, DblVal As Double
        
        If PTypRes.GFormulasBoo Then
          Let StrTMH = PRngBox.Range("G6").Address(False, False) ' TMH
          Let StrAmt = PRngBox.Range(PTypRes.PAssessmentMain.StrRngA1).Offset(0, 5).Address(False, False) ' Assessment
        Else
          Let DblTMH = VBA.CDbl(PTypRes.PWeight.Vector(2, 1)) ' TMH
          Let DblAmt = VBA.CDbl(PTypRes.PAssessmentMain.Vector(UBound(PTypRes.PAssessmentMain.Vector), UBound(PTypRes.PAssessmentMain.Vector, 2))) ' Assessment
        End If
        
        For i = 0 To UBound(VecAux0)
          
          If PTypRes.GFormulasBoo Then
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
            Let DblVal = IIf(DblPer > 0, (StrAmt * DblPer * 0.01), StrTMH * DblCsh)
            Let DblVal = IIf(VBA.CBool(VecAux0(i, 3)), VBA.CDbl(VecAux0(i, 2)), DblVal)
            Let .Vector = RES.ArrayAddAtLast(.Vector, Array(VBA.StrConv(VBA.CStr(VecAux0(i, 0)), vbProperCase), Empty, Empty, VecAux0(i, 1), VecAux0(i, 2), DblVal))
            Let PTypRes.TotalDeductionsDbl = PTypRes.TotalDeductionsDbl + DblVal
          End If
  
        Next i
        If Not VBA.IsArray(.Vector) Then GoTo EH
      End If

    End With

  End With

  'Return
  Let PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getdeductions_exports = True

EH:
  Let VecAux0 = Empty
  Set RngBox = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getdeductions_rmtaxes( _
  ByVal PRngBox As Range, _
  ByRef RES As AppResources_current, _
  ByRef PTypRes As GLOTYPPURCHASE) As Boolean

  Dim RngBox As Range
  
  Dim VecAux0 As Variant, VecAux1 As Variant
  
  Const sFN As String = "APPCNC::-core_spreadmodule_generate_data_purchasebase_getdeductions_rmtaxes"

  On Error GoTo EH
  Let PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getdeductions_rmtaxes = False

  With PTypRes

    ' Get data (0.CONCEPTO 1.PORCENTAJE 2.VALOR 3.VALOR DIRECTO)
    Dim i As Long
    For i = 0 To UBound(.PContents.Vector)
      If i > 2 Then
        Let VecAux1 = MLT.MGet(eqlMdlTblConfigpurrm, eqlMdlArray, "MINERAL, VENTA INTERNA", MStrWhere:="MINERAL=" & VBA.CStr(.PContents.Vector(i, 0)))
        If VBA.IsArray(VecAux1) Then Let VecAux0 = RES.ArrayAddAtLast(VecAux0, VecAux1)
        Let VecAux1 = Empty
      End If
    Next i
    Let .PDRMRets.Exists = VBA.IsArray(VecAux0)
    
    ' Fill matrix
    With .PDRMRets
      
      Let .LngRngA1 = PTypRes.PDExports.LngRngA1
      If .Exists Then
        
        'Check for Resume tables boolean
        If PTypRes.GResumedRMBoo Then
          Let VecAux1 = VecAux0
          Let VecAux0 = Empty
          ReDim VecAux0(0 To 0, 0 To UBound(VecAux1, 2))
          Let VecAux0(0, 1) = 0
          For i = 0 To UBound(VecAux1)
            Let VecAux0(0, 1) = VBA.CDbl(VecAux0(0, 1)) + VBA.CDbl(VecAux1(i, 1))
          Next i
          Let VecAux0(0, 0) = "Regalías Mineras"
          Let VecAux1 = Empty
        End If
        
        'Determine Addresses and row numbers
        Dim xCols As Long
        Let xCols = 5
        Let .LngRngA1 = .LngRngA1 + IIf(PTypRes.PDExports.Exists, 2, 4)
        Let .StrRngA1 = "F" & .LngRngA1
        If PTypRes.GFormulasBoo Then Set RngBox = PRngBox.Range(.StrRngA1)
        Let .LngRngA1 = .LngRngA1 + UBound(VecAux0)
        
        ' Fill final vector with data
        ReDim .Vector(0 To 0, 0 To xCols)
        Let .Vector(0, 0) = "RETENCIONES DE LEY - REGALÍAS"
        
        ' Calculate results
        Dim StrPer As String, StrRM As String, StrAmt As String, StrVal As String
        Dim DblPer As Double, DblRM As Double, DblAmt As Double, DblVal As Double
        
        If PTypRes.GFormulasBoo Then
          Let StrAmt = PRngBox.Range(PTypRes.PAssessmentMain.StrRngA1).Offset(0, 5).Address(False, False) ' Assessment
        Else
          Let DblAmt = VBA.CDbl(PTypRes.PAssessmentMain.Vector(UBound(PTypRes.PAssessmentMain.Vector), UBound(PTypRes.PAssessmentMain.Vector, 2))) ' Assessment
        End If
        
        For i = 0 To UBound(VecAux0)
          
          Let StrRM = IIf(VBA.LCase(VBA.CStr(VecAux0(i, 0))) = "rm", VBA.CStr(VecAux0(i, 0)), VBA.StrConv(VBA.CStr(VecAux0(i, 0)), vbProperCase))
          If PTypRes.GFormulasBoo Then
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

  'Return
  Let PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getdeductions_rmtaxes = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Set RngBox = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getdeductions_lawtaxes( _
  ByVal PRngBox As Range, _
  ByRef RES As AppResources_current, _
  ByRef PTypRes As GLOTYPPURCHASE) As Boolean

  Dim RngBox As Range
  
  Dim VecAux0 As Variant, VecAux1 As Variant

  Const sFN As String = "APPCNC::-core_spreadmodule_generate_data_purchasebase_getdeductions_lawtaxes"

  On Error GoTo EH
  Let PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getdeductions_lawtaxes = False

  With PTypRes

    ' Get data (0.CONCEPTO 1.PORCENTAJE)
    Let VecAux0 = MLT.MGet(eqlMdlTblConfigpurretentions, eqlMdlArray, "CONCEPTO, PORCENTAJE", MStrWhere:="SELECTED=1")
    Let .PDLawRets.Exists = VBA.IsArray(VecAux0)
    
    ' Fill matrix
    With .PDLawRets
      
      Let .LngRngA1 = PTypRes.PDRMRets.LngRngA1
      If .Exists Then
        
        'Check for Resume tables boolean
        Dim i As Long
        If PTypRes.GResumedRetentionsBoo Then
          Let VecAux1 = VecAux0
          Let VecAux0 = Empty
          ReDim VecAux0(0 To 0, 0 To UBound(VecAux1, 2))
          Let VecAux0(0, 1) = 0
          For i = 0 To UBound(VecAux1)
            Let VecAux0(0, 1) = VBA.CDbl(VecAux0(0, 1)) + VBA.CDbl(VecAux1(i, 1))
          Next i
          Let VecAux0(0, 0) = "Varias"
          Let VecAux1 = Empty
        End If
        
        'Determine Addresses and row numbers
        Dim xCols As Long
        Let xCols = 5
        Let .LngRngA1 = .LngRngA1 + IIf(PTypRes.PDRMRets.Exists, 2, 4)
        Let .StrRngA1 = "F" & .LngRngA1
        If PTypRes.GFormulasBoo Then Set RngBox = PRngBox.Range(.StrRngA1)
        Let .LngRngA1 = .LngRngA1 + UBound(VecAux0)
        
        ' Fill final vector with data
        ReDim .Vector(0 To 0, 0 To xCols)
        Let .Vector(0, 0) = "RETENCIONES DE LEY - INSTITUCIONALES"
        
        ' Calculate results
        Dim StrPer As String, StrRet As String, StrAmt As String, StrVal As String
        Dim DblPer As Double, DblRM As Double, DblAmt As Double, DblVal As Double
        
        If PTypRes.GFormulasBoo Then
          Let StrAmt = PRngBox.Range(PTypRes.PAssessmentMain.StrRngA1).Offset(0, 5).Address(False, False) ' Assessment
        Else
          Let DblAmt = VBA.CDbl(PTypRes.PAssessmentMain.Vector(UBound(PTypRes.PAssessmentMain.Vector), UBound(PTypRes.PAssessmentMain.Vector, 2))) ' Assessment
        End If
        
        For i = 0 To UBound(VecAux0)
          
          Let StrRet = VBA.UCase(VBA.CStr(VecAux0(i, 0)))
          If PTypRes.GFormulasBoo Then
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

  'Return
  Let PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getdeductions_lawtaxes = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Set RngBox = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getdeductions_othertaxes( _
  ByVal PRngBox As Range, _
  ByRef RES As AppResources_current, _
  ByRef PTypRes As GLOTYPPURCHASE) As Boolean

  Dim RngBox As Range
  
  Dim VecAux0 As Variant, VecAux1 As Variant
  
  Const sFN As String = "APPCNC::-core_spreadmodule_generate_data_purchasebase_getdeductions_othertaxes"

  On Error GoTo EH
  Let PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getdeductions_othertaxes = False

  With PTypRes

    ' Get data (0.CONCEPTO 1.PORCENTAJE)
    Let VecAux0 = MLT.MGet(eqlMdlTblConfigpurretentionsother, eqlMdlArray, "CONCEPTO, PORCENTAJE", MStrWhere:="SELECTED=1")
    Let .PDOthRets.Exists = VBA.IsArray(VecAux0)
    
    ' Fill matrix
    With .PDOthRets
      
      Let .LngRngA1 = PTypRes.PDLawRets.LngRngA1
      If .Exists Then
        
        'Check for Resume tables boolean
        Dim i As Long
        If PTypRes.GResumedREetentionsOthersBoo Then
          Let VecAux1 = VecAux0
          Let VecAux0 = Empty
          ReDim VecAux0(0 To 0, 0 To UBound(VecAux1, 2))
          Let VecAux0(0, 1) = 0
          For i = 0 To UBound(VecAux1)
            Let VecAux0(0, 1) = VBA.CDbl(VecAux0(0, 1)) + VBA.CDbl(VecAux1(i, 1))
          Next i
          Let VecAux0(0, 0) = "Varias"
          Let VecAux1 = Empty
        End If
        
        'Determine Addresses and row numbers
        Dim xCols As Long
        Let xCols = 5
        Let .LngRngA1 = .LngRngA1 + IIf(PTypRes.PDRMRets.Exists, 2, 4)
        Let .StrRngA1 = "F" & .LngRngA1
        If PTypRes.GFormulasBoo Then Set RngBox = PRngBox.Range(.StrRngA1)
        Let .LngRngA1 = .LngRngA1 + UBound(VecAux0)
        
        ' Fill final vector with data
        ReDim .Vector(0 To 0, 0 To xCols)
        Let .Vector(0, 0) = "RETENCIONES INSTITUCIONALES OTRAS"
        
        ' Calculate results
        Dim StrPer As String, StrRet As String, StrAmt As String, StrVal As String
        Dim DblPer As Double, DblRM As Double, DblAmt As Double, DblVal As Double
        
        If PTypRes.GFormulasBoo Then
          Let StrAmt = PRngBox.Range(PTypRes.PAssessmentMain.StrRngA1).Offset(0, 5).Address(False, False) ' Assessment
        Else
          Let DblAmt = VBA.CDbl(PTypRes.PAssessmentMain.Vector(UBound(PTypRes.PAssessmentMain.Vector), UBound(PTypRes.PAssessmentMain.Vector, 2))) ' Assessment
        End If
        
        For i = 0 To UBound(VecAux0)
          
          Let StrRet = VBA.UCase(VBA.CStr(VecAux0(i, 0)))
          If PTypRes.GFormulasBoo Then
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

  'Return
  Let PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getdeductions_othertaxes = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Set RngBox = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getdeductions_discounts( _
  ByVal PRngBox As Range, _
  ByRef RES As AppResources_current, _
  ByRef PTypRes As GLOTYPPURCHASE) As Boolean

  Dim RngBox As Range
  
  Dim VecAux0 As Variant, VecAux1 As Variant

  Const sFN As String = "APPCNC::-core_spreadmodule_generate_data_purchasebase_getdeductions_discounts"

  On Error GoTo EH
  Let PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getdeductions_discounts = False

  With PTypRes

    ' Get data (0.CONCEPTO 1.VALOR)
    Let VecAux0 = MLT.MGet(eqlMdlTblConpurdeddiscounts, eqlMdlArray)
    Let .PDOthRets.Exists = VBA.IsArray(VecAux0)
    
    ' Fill matrix
    With .PDOthRets
      
      Let .LngRngA1 = PTypRes.PDOthRets.LngRngA1
      If .Exists Then
        
        'Check for Resume tables boolean
        Dim i As Long
        If PTypRes.GResumedDiscountsBoo Then
          Dim DblAux0 As Double
          Let VecAux1 = RES.ArrayPop(VecAux0, RByCols:=True)
          Let DblAux0 = Application.WorksheetFunction.Sum(VecAux1)
          Let VecAux1 = Empty
          ReDim VecAux0(0 To 0, 0 To 1)
          Let VecAux0(0, 0) = "Descuentos Varios"
          Let VecAux0(0, 1) = DblAux0
        End If
        
        'Determine Addresses and row numbers
        Dim xCols As Long
        Let xCols = 5
        Let .LngRngA1 = .LngRngA1 + IIf(PTypRes.PDOthRets.Exists, 2, 4)
        Let .StrRngA1 = "F" & .LngRngA1
        If PTypRes.GFormulasBoo Then Set RngBox = PRngBox.Range(.StrRngA1)
        Let .LngRngA1 = .LngRngA1 + UBound(VecAux0)
        
        ' Fill final vector with data
        ReDim .Vector(0 To 0, 0 To xCols)
        Let .Vector(0, 0) = "DESCUENTOS"
        
        ' Calculate results
        Dim StrDsc As String, StrNme As String, StrVal As String
        Dim DblDsc As Double, DblVal As Double
        
        For i = 0 To UBound(VecAux0)
          
          Let StrNme = VBA.StrConv(VBA.CStr(VecAux0(i, 0)), vbProperCase)
          If PTypRes.GFormulasBoo Then
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

  'Return
  Let PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getdeductions_discounts = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Set RngBox = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getbonus( _
  ByVal PRngBox As Range, _
  ByRef RES As AppResources_current, _
  ByRef PTypRes As GLOTYPPURCHASE) As Boolean

  Dim RngBox As Range
  
  Dim VecAux0 As Variant, VecAux1 As Variant

  Const sFN As String = "APPCNC::-core_spreadmodule_generate_data_purchasebase_getbonus"

  On Error GoTo EH
  Let PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getbonus = False

  With PTypRes

    ' Get data (0.CONCEPTO 1.VALOR)
    Let VecAux0 = MLT.MGet(eqlMdlTblConpurbonus, eqlMdlArray)
    Let .PBonus.Exists = VBA.IsArray(VecAux0)
    Let .TotalBonusDbl = 0
    
    ' Fill matrix
    With .PBonus
      
      Let .LngRngA1 = PTypRes.PAssessmentMain.LngRngA1 + PTypRes.PDeductions.LngRngA1
      If .Exists Then
        
        'Check for Resume tables boolean
        Dim i As Long
        If PTypRes.GResumedBonusBoo Then
          Dim DblAux0 As Double
          Let VecAux1 = RES.ArrayPop(VecAux0, RByCols:=True)
          Let DblAux0 = Application.WorksheetFunction.Sum(VecAux1)
          Let VecAux1 = Empty
          ReDim VecAux0(0 To 0, 0 To 1)
          Let VecAux0(0, 0) = "Bonos Varios"
          Let VecAux0(0, 1) = DblAux0
        End If
        
        'Determine Addresses and row numbers
        Dim xCols As Long
        Let xCols = 5
        Let .LngRngA1 = .LngRngA1 + 2
        Let .StrRngA1 = "F" & .LngRngA1
        If PTypRes.GFormulasBoo Then Set RngBox = PRngBox.Range(.StrRngA1)
        Let .LngRngA1 = .LngRngA1 + UBound(VecAux0)
        
        ' Fill final vector with data
        ReDim .Vector(0 To 0, 0 To xCols)
        Let .Vector(0, 0) = "BONOS"
        
        ' Calculate results
        Dim StrNme As String, StrVal As String
        Dim DblVal As Double
        
        For i = 0 To UBound(VecAux0)
          
          Let StrNme = VBA.StrConv(VBA.CStr(VecAux0(i, 0)), vbProperCase)
          If PTypRes.GFormulasBoo Then
            Let StrVal = VBA.Replace(VecAux0(i, 1), ",", ".")
            Let .Vector = RES.ArrayAddAtLast(.Vector, Array(StrNme, Empty, Empty, Empty, Empty, StrVal))
          Else
            Let DblVal = VBA.CDbl(VecAux0(i, 1))
            Let .Vector = RES.ArrayAddAtLast(.Vector, Array(StrNme, Empty, Empty, Empty, Empty, DblVal))
            Let PTypRes.TotalBonusDbl = PTypRes.TotalBonusDbl + DblVal
          End If
  
        Next i
        If Not VBA.IsArray(.Vector) Then GoTo EH
        
        'Fill totals
        If PTypRes.GFormulasBoo Then
          Let .Vector = RES.ArrayAddAtLast(.Vector, Array("TOTAL BONOS [" & PTypRes.GCurForeSymStr & "]", Empty, Empty, Empty, Empty, "=SUM(" & RngBox.Offset(0, 5).Resize(UBound(VecAux0) + 1, 1).Address(False, False) & ")"))
        Else
          Let .Vector = RES.ArrayAddAtLast(.Vector, Array(StrNme, Empty, Empty, Empty, Empty, PTypRes.TotalBonusDbl))
        End If
        If Not VBA.IsArray(.Vector) Then GoTo EH
      
      End If

    End With
    Let .TotalBonusRng = "K" & .PBonus.LngRngA1
  
  End With

  'Return
  Let PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getbonus = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Set RngBox = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getfinalpurchase( _
  ByVal PRngBox As Range, _
  ByRef RES As AppResources_current, _
  ByRef PTypRes As GLOTYPPURCHASE) As Boolean

  Dim NTS As AppResNumberToString

  Dim RngBox As Range
  
  Const sFN As String = "APPCNC::-core_main_spreadmodule_generate_data_purchasebase_getfinalpurchase"

  On Error GoTo EH
  Let PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getfinalpurchase = False

  With PTypRes
    
    'Fill matrix PLngRowCounter
    With .PFinal
      
      Dim StrAux0 As String
      Dim xCols As Long, xRows As Long
      Dim DblAux0 As Double
      'Determine Addresses and row numbers
      Let .LngRngA1 = PTypRes.PBonus.LngRngA1
      Let xCols = 10
      Let xRows = 4
      Let .LngRngA1 = .LngRngA1 + 3
      Let .StrRngA1 = "A" & .LngRngA1
      If PTypRes.GFormulasBoo Then Set RngBox = PRngBox.Range(.StrRngA1)
      
      ' Fill final vector with data
      Let DblAux0 = 0
      ReDim .Vector(0 To xRows, 0 To xCols)
      
      Let .Vector(0, 0) = "LIQUIDACIÓN FINAL (VALOR BRUTO" & IIf(PTypRes.PDeductions.Exists, " - DEDUCCIONES", VBA.vbNullString) & IIf(PTypRes.PBonus.Exists, " + BONOS", VBA.vbNullString) & ")"
      If PTypRes.GMNBoo Then Let .Vector(0, 9) = "VALOR " & PTypRes.GCurMainSymStr ' MN
      Let .Vector(0, 10) = "VALOR " & PTypRes.GCurForeSymStr 'ME
  
      If PTypRes.GFormulasBoo Then 'TODO: ING!!!!
        
        'Set liquid
        Let StrAux0 = RngBox.Offset(0, 10).Address(False, False)
        Let .Vector(1, 0) = "='Líquido pagable (' & IF(" & StrAux0 & "=0,'Sin saldo',IF(" & StrAux0 & ">0,'Saldo a favor','Saldo en contra'))):"
        If PTypRes.GMNBoo Then Let .Vector(1, 9) = "=" & RngBox.Offset(0, 10).Address(False, False) & "*" & PRngBox.Range("B10").Address(False, False) 'MN
        Let .Vector(1, 10) = "=SUM(" & _
          PRngBox.Range(PTypRes.TotalAssessmentRng).Address(False, False) & _
          IIf(PTypRes.PDeductions.Exists, ",-" & PRngBox.Range(PTypRes.TotalDeductionsRng).Address(False, False), VBA.vbNullString) & _
          IIf(PTypRes.PBonus.Exists, "," & PRngBox.Range(PTypRes.TotalBonusRng).Address(False, False), VBA.vbNullString) & ")"
        'Set liquid textual
        Let .Vector(2, 1) = "=EQL_NUMERO_A_TEXTO_CURRENT(" & RngBox.Offset(0, 10).Address(False, False) & ",""" & PTypRes.GCurForeNameStr & """,""" & PTypRes.GCurForeNameStrs & """,0)"
      
      Else
          
        'Set liquid
        Let DblAux0 = (PTypRes.TotalAssessmentDbl - PTypRes.TotalDeductionsDbl) + PTypRes.TotalBonusDbl
        Let .Vector(1, 0) = "Líquido pagable (" & IIf(DblAux0 = 0, "Sin saldo", IIf(DblAux0 > 0, "Saldo a favor", "Saldo en contra")) & "):"
        If PTypRes.GMNBoo Then Let .Vector(1, 9) = DblAux0 * PTypRes.GTCDbl 'MN
        Let .Vector(1, 10) = DblAux0
        'Set liquid textual
        Set NTS = New AppResNumberToString
        Let .Vector(2, 1) = NTS.NumToString(DblAux0, PTypRes.GCurForeNameStr, PTypRes.GCurForeNameStrs, False)
        Set NTS = Nothing
      
      End If
      Let .Vector(2, 0) = "Son:"
    
    End With
  
  End With

  'Return
  Let PurchaseConcentrate_Core_Spreadmodule_generate_data_purchasefinal_getfinalpurchase = True

EH:
  Set RngBox = Nothing
  Set NTS = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function









' ASSETS ----------------------------------------------------------------------------'
Private Sub PurchaseConcentrate_Core_assets_srcarray_to_dstarray(ByRef PVarSrc As Variant)
  
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
  Call EHGLOBAL.ErrorHandlerRaise("APPCNC::-core_assets_srcarray_to_dstarray")

End Sub










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
  Let ICoreController_GetSolution = PurchaseConcentrate_Core_Main(CVarRequest, CVarResponse)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("APPCNC::IGetSolution")

End Function

Private Function ICoreController_GetSolutionUDF(ByVal CStrRequester As String, ByVal CVarRequest As Variant) As Variant
End Function
' METHODS PUBLIC --------------------------------------------------------------------'
' ============================ CLASS MODULES STRUCTURE =============================='
