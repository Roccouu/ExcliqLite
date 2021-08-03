VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppExcliqProjections_current"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ABSTRACT IMPLEMENTATIONS ----------------------------------------------------------'
Implements ICoreController

' OBJECT VARIABLES (GLOOBJ_) --------------------------------------------------------'
Private EHGLOBAL As AppErrorHandler

' VARIABLES (GLOStr_) ---------------------------------------------------------------'
Private GLOBOO_EH As Boolean
Private GLOBOO_FORMULAS As Boolean

' CONSTANTS ERROR MESSAGES (GLOSTR_ERR_) --------------------------------------------'
Private Type GLOTYP_ERRORS
  NoDataSource As String
  NoDataZeroes As String
End Type
Private GLOTypErr As GLOTYP_ERRORS

'Private Const GLOSTR_ERR_1 As String = "Faltan datos de orígen."
'Private Const GLOSTR_ERR_2 As String = "Al menos un dato debe ser diferente de cero."


' ============================= CLASS MODULES STRUCTURE ============================='
' OBJECT VARIABLES (GLOOBJ_) --------------------------------------------------------'
' VARIABLES (GLOStr_) ---------------------------------------------------------------'
' CONSTANTS ERROR MESSAGES (GLOSTR_ERR_) --------------------------------------------'
' CONSTANTS (GLOSTR_) ---------------------------------------------------------------'
' PROPERTIES ------------------------------------------------------------------------'
' CONSTRUCTOR/DESTRUCTOR ------------------------------------------------------------'
' CORE METHODS PRIVATE --------------------------------------------------------------'
' CORE METHODS LIST (PRIVATE) -------------------------------------------------------'
' 1. [ ] VALIDATIONS
' Projection_Core_UDF_validation
' 2. [ ] PREPARATIONS
' Projection_Core_UDF_resolve
' 3. [ ] RESOLVE
' Projection_Core_UDF_resolve_net_cash_flow_calculate
' 4. [ ] MODULE RESOLUTIONS
' Projection_Core_Main
' Projection_Core_Spreadmodule_generate
' Projection_Core_Spreadmodule_generate_viability
' ASSETS ----------------------------------------------------------------------------'
' Projection_Core_assets_get_ctzx
' ERRORS MANNAGEMENT AND ASSERTSIONS ------------------------------------------------'
' Debug_asset
' Debug_assert_errors_set_error
' METHODS PUBLIC --------------------------------------------------------------------'
' UI METHODS LIST UDF ---------------------------------------------------------------'
' ProjectionUDF
' INTERFACE METHODS LIST ------------------------------------------------------------'
' ProjectionGet
' ============================ CLASS MODULES STRUCTURE =============================='


' PROPERTIES ------------------------------------------------------------------------'
' ABSTRACT INTERFACE PROPERTIES IMPLEMENTATION --------------------------------------'
Public Property Set ICoreController_ErrorHandler(ByRef RHS As AppErrorHandler)
  
  Set EHGLOBAL = Nothing
  Set EHGLOBAL = RHS
  Let GLOBOO_EH = (Not RHS Is Nothing)

End Property





' CONSTRUCTOR/DESTRUCTOR ------------------------------------------------------------'
Private Sub Class_Initialize()
  
  Set EHGLOBAL = New AppErrorHandler
  Dim MDL As ModelExcliqliteDatasheet
  Set MDL = New ModelExcliqliteDatasheet
  Let GLOBOO_FORMULAS = MDL.Formulas
  Set MDL = Nothing
  
  Let GLOTypErr.NoDataSource = "Faltan datos de orígen."
  Let GLOTypErr.NoDataZeroes = "Al menos un dato debe ser diferente de cero."

End Sub

Private Sub Class_Terminate()
  
  If GLOBOO_EH = False Then Set EHGLOBAL = Nothing

End Sub







' CORE METHODS PRIVATE --------------------------------------------------------------'
' CORE METHODS LIST (PRIVATE) -------------------------------------------------------'
' 1. [x] VALIDATIONS
Private Function Projection_Core_UDF_validation(ByVal PVarNetCashFlow As Variant, ByRef PVarErrUDF As Variant) As Boolean
  
  Dim RES As AppResources_current
  
  Dim VecCash As Variant
  Dim i As Long
  
  Const sFN As String = "APPPRO::-core_udf_validation"
  
  On Error GoTo EH
  Let Projection_Core_UDF_validation = False
  
  Set RES = New AppResources_current
  Set RES.ErrorHandler = EHGLOBAL
  'Dim resq As AppResources
  'Let VecCash = resq.resArrayGetDataArrayFromVariants(PVarNetCashFlow, i, False, False, True, RBooRngRowsCols:=True)
  Let VecCash = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(PVarNetCashFlow, i, RValueType:=eqlResNumbers, RBooAbs:=True, RBooIncludeZeroes:=True)
  If i <= 0 Then Call Debug_assert_errors_set_error(PVarErrUDF, xlErrValue, 1, GLOTypErr.NoDataSource): GoTo EH 'TODO
  If Application.WorksheetFunction.Sum(VecCash) = 0 Then Call Debug_assert_errors_set_error(PVarErrUDF, xlErrValue, 2, GLOTypErr.NoDataZeroes): GoTo EH 'TODO
  Let VecCash = Empty
  Set RES = Nothing
  
  Let Projection_Core_UDF_validation = True

EH:
  Let VecCash = Empty
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function





' 3. [x] RESOLVE
Private Function Projection_Core_UDF_resolve(ByVal PVarNetCashFlow As Variant, Optional PBooLocal As Boolean = False) As Variant
  
  On Error GoTo EH
  'Validation
  If Not Projection_Core_UDF_validation(PVarNetCashFlow, Projection_Core_UDF_resolve) Then GoTo EH

  'Preparation

  'Resolve
  If Not Projection_Core_UDF_resolve_net_cash_flow_calculate(PVarNetCashFlow, Projection_Core_UDF_resolve) Then: GoTo EH

EH:
  If PBooLocal Then
    If Not VBA.Err.Number = 0 Then Call EHGLOBAL.ErrorHandlerRaise("APPPRO::-core_udf_resolve")
  Else
    #If Debugging Then
      Debug.Print VBA.Err.Description
    #End If
    On Error GoTo -1: Call VBA.Err.Clear
  End If

End Function

Private Function Projection_Core_UDF_resolve_net_cash_flow_calculate( _
  ByVal PVarNetCashFlow As Variant, _
  ByRef PVarResponse As Variant) As Boolean
  
  Dim RES As AppResources_current
  
  Dim VecCash As Variant
  Dim i As Long
  Dim DblFlow As Double
  Dim BooViability As Boolean
  
  Const sFN As String = "APPPRO::-core_udf_resolve_net_cash_flow_calculate"

  On Error GoTo EH
  Let Projection_Core_UDF_resolve_net_cash_flow_calculate = False
  
  'Operate
  Set RES = New AppResources_current
  Set RES.ErrorHandler = EHGLOBAL
  'Let VecCash = RES.resArrayGetDataArrayFromVariants(PVarNetCashFlow, i, False, False, True, RBooRngRowsCols:=True)
  Let VecCash = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(PVarNetCashFlow, i, RValueType:=eqlResNumbers, RBooAbs:=True, RBooIncludeZeroes:=True)
  Set RES = Nothing
  Let DblFlow = 0
  For i = 1 To UBound(VecCash)
    Let DblFlow = DblFlow + VecCash(i)
    Let BooViability = (DblFlow >= VBA.Abs(VecCash(0)))
    If BooViability Then Exit For
  Next i
  
  'Return
  'Utilities appears in the time i
  Let PVarResponse = IIf(BooViability, "Rentable desde " & i, "Proyecto Inviable")

  Let Projection_Core_UDF_resolve_net_cash_flow_calculate = True

EH:
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

' 4. [x] MODULE RESOLUTIONS
Private Function Projection_Core_Main(ByRef PVarRequest As Variant, ByRef PVarResponse As Variant) As Boolean
  
  Dim VecProjection As Variant, VecViability As Variant
  Dim j As Long

  Const sFN As String = "APPPRO::-core_main"

  On Error GoTo EH
  Let Projection_Core_Main = False

  'Request: [DblIo, StrW[], StrCtz[], DblVR, StrD, StrU, DblEgr, DblT, Dbl%, BooVAN, BooTIR, BooGraph, BooGraphSheet, BooNewSheet, Rng]
  
  'Resolve Sources
  If Not Projection_Core_Spreadmodule_generate(PVarRequest, VecProjection, j) Then GoTo EH

  'Resolve Regression
  If Not Projection_Core_Spreadmodule_generate_viability(PVarRequest, VecProjection, VecViability, j) Then GoTo EH

  'Return
  Let PVarResponse = Array(VecProjection, VecViability)
  Let Projection_Core_Main = True

EH:
  Let VecProjection = Empty
  Let VecViability = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Projection_Core_Spreadmodule_generate( _
  ByRef PVarRequest As Variant, _
  ByRef PVarPROJECTION As Variant, _
  ByRef PxJ As Long)
  
  Dim RES As AppResources_current
  Dim REG As AppExcliqRegression
  
  Dim RngBox As Range

  Dim VecTime As Variant, VecWeights As Variant, VecWName As Variant, VecCtz As Variant, VecUnitCtz As Variant, VecCtzX As Variant
  Dim VecHeader As Variant, VecAux1 As Variant, VecWAcum As Variant
  Dim StrDivise As String, StrTimeUnit As String, StrAux As String
  Dim DblIo As Double, DblVR As Double, DblOut As Double, DblTime As Double, DblRec As Double
  Dim i As Long, j As Long, k As Long, m As Long, n As Long, u As Long, v As Long, w As Long
  Dim xWLen As Long, xWLen2 As Long
  
  Const STR_DASH As String = "-"
  Const sFN As String = "APPPRO::-core_main_spreadmodule_generate"

  ' Request
  ' 0-DblIo      5-VecW[]           9-StrDivise         11-BooVAN           16-Rng
  ' 1-DblVR      6-VecWName[]       10-StrTimeUnit      12-BooTIR
  ' 2-DblEgr     7-VecWUnitCtz[]                        13-BooGraph
  ' 3-DblT       8-VecCtz[]                             14-BooGraphSheet
  ' 4-Dbl%                                              15-BooNewSheet
  On Error GoTo EH
  Let Projection_Core_Spreadmodule_generate = False

  'Unpack data
  Let DblIo = VBA.Abs(VBA.CDbl(PVarRequest(0))) * -1
  Let DblVR = VBA.CDbl(PVarRequest(1))
  Let DblOut = VBA.CDbl(PVarRequest(2))
  Let DblTime = VBA.CDbl(PVarRequest(3))
  Let DblRec = VBA.CDbl(PVarRequest(4))
  
  Let StrDivise = VBA.CStr(PVarRequest(9))
  Let StrTimeUnit = VBA.CStr(PVarRequest(10))

  Set RES = New AppResources_current
  Set RES.ErrorHandler = EHGLOBAL
'  Let VecWeights = RES.resArrayGetDataArrayFromVariants(VBA.CStr(PVarRequest(5)), i, False, True)
'  Let VecWName = RES.resArrayGetDataArrayFromVariants(VBA.CStr(PVarRequest(6)), i, True)
'  Let VecUnitCtz = RES.resArrayGetDataArrayFromVariants(VBA.CStr(PVarRequest(7)), i, True, RStrCustomRegex:="%")
  Let VecWeights = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(VBA.CStr(PVarRequest(5)), i, RValueType:=eqlResNumbers, RBooAbs:=True)
  Let VecWName = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(VBA.CStr(PVarRequest(6)), i, RValueType:=eqlResStrings)
  Let VecUnitCtz = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(VBA.CStr(PVarRequest(7)), i, RValueType:=eqlResStrings, RStrCustomRegex:="%")
  Let VecCtz = VBA.Split(PVarRequest(8), "/") 'Each Ctz comes as: 1;2;3/1,2;2;4,5/4;5;9 etc...
  Set RES = Nothing
  
  If GLOBOO_FORMULAS Then Set RngBox = PVarRequest(16)

  'Get X list from VecCtz
  Let VecCtzX = Projection_Core_assets_get_ctzx(VecCtz)

  'Get Time vector
  ReDim VecTime(0)
  For i = 0 To DblTime
    ReDim Preserve VecTime(i)
    Let VecTime(i) = i
  Next i

  'Calculate the final matrix dimenssion: Rows[1: Header], Cols[1: Time, Io; W*2: Wi, Ctzi; 4: Outs, ...]
  Let m = 1 + UBound(VecTime)
  Let n = 1 + ((UBound(VecWeights) + 1) * 2) + 4

  'Fill Vector of Headers
  Let VecAux1 = Array("Tiempo [" & StrTimeUnit & "]", "Inversión [" & StrDivise & "]", "Cotización", "Valor Bruto", _
                      "Egresos [" & DblOut & "%]", "Utilidad Neta [" & StrDivise & "]", "Valor Residual [" & StrDivise & "]", "Flujo Neto de Caja [" & StrDivise & "]")
  ReDim VecHeader(0)
  Let xWLen = 1 + (UBound(VecWeights) + 1) 'Ctz
  Let xWLen2 = (xWLen + (UBound(VecWeights) + 1)) 'Weights + Ctz
  For i = 0 To n
    ReDim Preserve VecHeader(i)
    If i <= 1 Then 'Time, Io
      Let VecHeader(i) = VecAux1(k)
      Let k = k + 1
    ElseIf i >= 2 And i <= xWLen Then 'Ctz
      'Cotización [OT/USD]
      Let VecHeader(i) = VecAux1(k) & "' " & VecWName(j) & "[" & StrDivise & "/" & VecUnitCtz(j) & "]"
      Let j = j + 1
      If i = xWLen Then Let k = k + 1: Let j = 0
    ElseIf i > xWLen And i <= xWLen2 Then  'W
      'Valor Bruto Ag
      Let VecHeader(i) = VecAux1(k) & " " & VecWName(j) & " [" & StrDivise & "]"
      Let j = j + 1
      If i = xWLen2 Then Let k = k + 1
    Else
      Let VecHeader(i) = VecAux1(k)
      Let k = k + 1
    End If
  Next i
  Let VecAux1 = Empty

  'Fill with final values PVarProjection  VecHeader
  Let k = 0
  Let u = 1
  Let v = 0
  ReDim PVarPROJECTION(0 To m, 0 To n)

  For j = 0 To n
    For i = 0 To m
      If i = 0 Then
        Let PVarPROJECTION(i, j) = VecHeader(j) 'Headers
      ElseIf j = 0 Then
        Let PVarPROJECTION(i, j) = VecTime(i - 1) 'Time
      ElseIf j = 1 Then
        Let PVarPROJECTION(i, j) = IIf(i = j, DblIo, STR_DASH) 'Io
      End If
    Next i
  Next j

  For j = 2 To xWLen 'Ctz
    For i = 1 To m
      If i = 1 Then
        Let PVarPROJECTION(i, j) = STR_DASH
        Let u = 1
      Else
        If GLOBOO_FORMULAS Then
          With RngBox.Range("A2")
            Let PVarPROJECTION(i, j) = "=EQL_REGRESION_LINEAL(""" & VecCtz(k) & """, """ & VecCtzX(k) & """, " & RngBox.Offset(i + 1, 0).Address & ")"
          End With
        Else
          Set REG = New AppExcliqRegression
          Set REG.ErrorHandler = EHGLOBAL
          Let PVarPROJECTION(i, j) = REG.RegressionGetY(VBA.CStr(VecCtz(k)), VBA.CStr(VecCtzX(k)), VBA.CStr(VecTime(u)))
          Set REG = Nothing
        End If
        Let u = u + 1
        If u = m Then Let k = k + 1
      End If
    Next i
  Next j

  For j = (xWLen + 1) To xWLen2 'VB
    For i = 1 To m
      If i = 1 Then
        Let PVarPROJECTION(i, j) = STR_DASH
        Let u = 0
      Else
        If GLOBOO_FORMULAS Then
          With RngBox.Range("A2")
            Let StrAux = VBA.CStr(VBA.Replace(VecWeights(u), ",", "."))
            Let PVarPROJECTION(i, j) = "=" & StrAux & "*" & RngBox.Offset(i + 1, (j - xWLen + 1)).Address
          End With
        Else
          If v = 0 Then 'Initialize just once
            ReDim VecWAcum(m - 2)
            For v = 0 To UBound(VecWAcum)
              Let VecWAcum(v) = 0 'Zero to initialize each empty space
            Next v
          End If
          Let PVarPROJECTION(i, j) = VecWeights(u) * PVarPROJECTION(i, (j - xWLen + 1))
          Let VecWAcum(i - 2) = VecWAcum(i - 2) + PVarPROJECTION(i, j)
        End If
      End If
    Next i
    Let u = u + 1
  Next j

  For j = (xWLen2 + 1) To (xWLen2 + 1) 'Outs
    For i = 1 To m
      If i = 1 Then
        Let PVarPROJECTION(i, j) = STR_DASH
        Let u = 0
      Else
        If GLOBOO_FORMULAS Then
          With RngBox.Range("A2")
            Let StrAux = VBA.CStr(VBA.Replace(DblOut, ",", "."))
            Let PVarPROJECTION(i, j) = "=SUM(" & RngBox.Offset(i + 1, (j - xWLen + 1)).Resize(1, xWLen - 1).Address & ")*" & StrAux & "/100"
          End With
        Else
          Let PVarPROJECTION(i, j) = VBA.CDbl(VecWAcum(u)) * DblOut / 100
          Let u = u + 1
        End If
      End If
    Next i
  Next j

  For j = (xWLen2 + 2) To (xWLen2 + 2) 'NU
    For i = 1 To m
      If i = 1 Then
        Let PVarPROJECTION(i, j) = STR_DASH
      Else
        If GLOBOO_FORMULAS Then
          With RngBox.Range("A2")
            Let PVarPROJECTION(i, j) = "=SUM(" & RngBox.Offset(i + 1, (j - xWLen)).Resize(1, xWLen - 1).Address & ")-" & RngBox.Offset(i + 1, (j - 1)).Address
          End With
        Else
          Let PVarPROJECTION(i, j) = VecWAcum(i - 2) - VBA.CDbl(PVarPROJECTION(i, j - 1))
        End If
      End If
    Next i
  Next j

  For j = (xWLen2 + 3) To (xWLen2 + 3) 'VR
    For i = 1 To m
      Let PVarPROJECTION(i, j) = IIf(i = 1, STR_DASH, DblVR)
    Next i
  Next j

  For j = (xWLen2 + 4) To (xWLen2 + 4) 'FNC
    For i = 1 To m
      If i = 1 Then
        Let PVarPROJECTION(i, j) = DblIo
      Else
        If GLOBOO_FORMULAS Then
          With RngBox.Range("A2")
            Let PVarPROJECTION(i, j) = "=SUM(" & .Offset(i, (j - 2)).Resize(1, 2).Address & ")"
          End With
        Else
          Let PVarPROJECTION(i, j) = PVarPROJECTION(i, j - 2) + PVarPROJECTION(i, j - 1)
        End If
      End If
    Next i
  Next j
  
  ReDim Preserve PVarRequest(UBound(PVarRequest) + 1)
  Let PVarRequest(UBound(PVarRequest)) = xWLen
  
  'Return
  Let PxJ = j
  Let Projection_Core_Spreadmodule_generate = True
  
EH:
  Let VecTime = Empty
  Let VecWeights = Empty
  Let VecWName = Empty
  Let VecCtz = Empty
  Let VecUnitCtz = Empty
  Let VecCtzX = Empty
  Let VecHeader = Empty
  Let VecAux1 = Empty
  Let VecWAcum = Empty
  Set RngBox = Nothing
  Set REG = Nothing
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Projection_Core_Spreadmodule_generate_viability( _
  ByRef PVarRequest As Variant, _
  ByVal PVarPROJECTION As Variant, _
  ByRef PVarViability As Variant, _
  ByRef PxJ As Long)
  
  Const sFN As String = "APPPRO::-core_main_spreadmodule_generate_viability"

  ' Request
  ' 0-DblIo      5-VecW[]           9-StrDivise         11-BooVAN           16-Rng
  ' 1-DblVR      6-VecWName[]       10-StrTimeUnit      12-BooTIR
  ' 2-DblEgr     7-VecWUnitCtz[]                        13-BooGraph
  ' 3-DblT       8-VecCtz[]                             14-BooGraphSheet
  ' 4-Dbl%                                              15-BooNewSheet
  On Error GoTo EH
  Let Projection_Core_Spreadmodule_generate_viability = False
  
  'Unpack data
  Dim RngBox As Range
  Dim StrDivise As String, StrTimeUnit As String
  Dim DblRat As Double
  Dim StrAux1 As String, StrAux2 As String, StrRat As String
  Dim m As Long, n As Long, mJ As Long, i As Long
  Dim BooVAN As Boolean, BooTIR As Boolean
  
  Let BooVAN = VBA.CBool(PVarRequest(11))
  Let BooTIR = VBA.CBool(PVarRequest(12))
  Let StrDivise = VBA.CStr(PVarRequest(9))
  Let StrTimeUnit = VBA.CStr(PVarRequest(10))
  Let StrRat = VBA.Replace(VBA.CStr(PVarRequest(4)), ",", ".")
  Let DblRat = VBA.CDbl(PVarRequest(4)) / 100
  Let mJ = VBA.CLng(PVarRequest(3)) + 1 'time+1
  If GLOBOO_FORMULAS Then Set RngBox = PVarRequest(16)
  
  'Requirements: m+4, n+1, rngbox
  Let m = 1 + IIf(BooVAN, 1, 0) + IIf(BooTIR, 1, 0)
  ReDim PVarViability(0 To m, 0 To 1)
  Let PVarViability(0, 0) = "Indicadores" 'Headers
  Let PVarViability(0, 1) = "Resultados"

  Let PVarViability(1, 0) = "Comportamiento" ' [" & StrTimeUnit & "]"
  If BooVAN Then Let PVarViability(2, 0) = "VAN [" & StrDivise & "]"
  If BooTIR Then Let PVarViability(3, 0) = "TIR [%]"

  If GLOBOO_FORMULAS Then
    With RngBox.Range("A3")
      Let PVarViability(1, 1) = "=EQL_PROYECCION_PROYECTO(" & .Offset(0, PxJ - 1).Resize(mJ, 1).Address(True, True) & ") & "" " & StrTimeUnit & """"
      Let StrAux1 = .Offset(0, PxJ - 1).Resize(mJ, 1).Address(True, True)
      If BooVAN Then
        Let PVarViability(2, 1) = "=+IFERROR(VNA(" & StrRat & "%," & StrAux1 & "),""¡Indeterminable!"")"
      End If
      If BooTIR Then
        Let PVarViability(3, 1) = "=+IFERROR(IRR(" & StrAux1 & "," & StrRat & "%),""¡Indeterminable!"")"
      End If
    End With
  Else
    Dim VecPro As Variant
    Dim StrPro As String
    ReDim VecPro(UBound(PVarPROJECTION) - 1)
    For i = 1 To UBound(PVarPROJECTION)
      Let VecPro(i - 1) = PVarPROJECTION(i, UBound(PVarPROJECTION, 2))
    Next i
    Let VecPro(0) = VBA.Abs(VecPro(0)) * -1
    Let StrPro = VBA.Join(VecPro, ";")
    Let PVarViability(1, 1) = Projection_Core_UDF_resolve(StrPro, True) & " " & StrTimeUnit
    If BooVAN Then Let PVarViability(2, 1) = Application.WorksheetFunction.NPV(DblRat, VecPro)
    If BooTIR Then Let PVarViability(3, 1) = Application.WorksheetFunction.IRR(VecPro, DblRat)
  End If
  
  Let Projection_Core_Spreadmodule_generate_viability = True

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)
  
End Function



' ASSETS ----------------------------------------------------------------------------'
Private Function Projection_Core_assets_get_ctzx(ByVal PVarCtz As Variant) As Variant
  
  Dim VecAux1 As Variant, VecAux2 As Variant
  Dim i As Long, j As Long, k As Long
  
  Const sFN As String = "APPPRO::-core_assets_get_ctzx"

  On Error GoTo EH
  
  ReDim VecAux2(0)
  For i = 0 To UBound(PVarCtz)
    Let VecAux1 = VBA.Split(PVarCtz(i), ";")
    Let k = UBound(VecAux1)
    ReDim VecAux1(0)
    For j = 0 To k
      ReDim Preserve VecAux1(j)
      Let VecAux1(j) = (j - k) '* -1
    Next j
    ReDim Preserve VecAux2(i)
    Let VecAux2(i) = VBA.Join(VecAux1, ";")
  Next i
  
  Let Projection_Core_assets_get_ctzx = VecAux2

EH:
  Let VecAux1 = Empty
  Let VecAux2 = Empty
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

'Errors mannagement
Private Sub Debug_assert_errors_set_error( _
  ByRef PVarErrUDF As Variant, _
  Optional PErrXL As XlCVError, _
  Optional PErrNum As Single, _
  Optional PErrMsg As String)

  Const sFN As String = "APPPRO::-core_assets_set_error"
    
  On Error GoTo EH
  
  Let PVarErrUDF = PErrXL
  Call EHGLOBAL.ErrorHandlerSet(PErrNum, PErrMsg)

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Sub






' METHODS PUBLIC --------------------------------------------------------------------'
' ABSTRACT INTERFACE UI METHODS LIST AND UDF METHODS LIST (PUBLIC) ------------------'
Public Function ICoreController_GetSolution(ByVal CStrRequester As String, ByVal CVarRequest As Variant, CVarResponse As Variant) As Boolean
  
  On Error GoTo EH
  Let ICoreController_GetSolution = Projection_Core_Main(CVarRequest, CVarResponse)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("APPPRO::IGetSolution")

End Function

Public Function ICoreController_GetSolutionUDF(ByVal CStrRequester As String, ByVal CVarRequest As Variant) As Variant
  
  On Error GoTo EH
  Let ICoreController_GetSolutionUDF = Projection_Core_UDF_resolve(CVarRequest)
  
EH:
  Call EHGLOBAL.ErrorHandlerRaise("APPPRO::IGetSolution")
  
End Function
' METHODS PUBLIC --------------------------------------------------------------------'
'============================= CLASS MODULES STRUCTURE =============================='

