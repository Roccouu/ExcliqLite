VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppExcliqRegression_current"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================== CLASS REGRESSION STRUCTURE ============================'
' OBJECT VARIABLES (GLOOBJ_) --------------------------------------------------------'
' VARIABLES (GLOStr_) ---------------------------------------------------------------'
' CONSTANTS ERROR MESSAGES (GLOSTR_ERR_) --------------------------------------------'
' CONSTANTS (GLOSTR_) ---------------------------------------------------------------'
' PROPERTIES ------------------------------------------------------------------------'
' CONSTRUCTOR/DESTRUCTOR ------------------------------------------------------------'
' CORE METHODS PRIVATE --------------------------------------------------------------'
' CORE METHODS LIST (PRIVATE) -------------------------------------------------------'
' 1. [x] VALIDATIONS
' Regression_Core_UDF_validation
' 2. [x] PREPARATIONS
' Regression_Core_UDF_preparation
' Regression_Core_UDF_preparation_set_x
' Regression_Core_UDF_preparation_set_y
' Regression_Core_UDF_preparation_set_e
' 3. [x] RESOLVE
' Regression_Core_UDF_resolve_regression
' Regression_Core_UDF_resolve_regression_calculate
' Regression_Core_UDF_resolve_regression_calculate_a
' Regression_Core_UDF_resolve_regression_calculate_b
' Regression_Core_UDF_resolve_regression_calculate_tendence_equation
' Regression_Core_UDF_resolve_regression_calculate_r
' Regression_Core_UDF_resolve_regression_calculate_r2
' Regression_Core_UDF_resolve_regression_calculate_se
' 4. [x] MODULE RESOLUTIONS
' Regression_Core_Main
' Regression_Core_Spreadmodule_generate_regression
' Regression_Core_Spreadmodule_generate_sources
' Regression_Core_Spreadmodule_generate_corelations
' ASSETS ----------------------------------------------------------------------------'
' Regression_Core_assets_getmainvectors
' ERROR ASSERTSIONS -----------------------------------------------------------------'
' Debug_asset
' METHODS PUBLIC --------------------------------------------------------------------'
' UI METHODS LIST UDF ---------------------------------------------------------------'
' RegressionUDF
' RegressionUDF_a
' RegressionUDF_b
' RegressionUDF_r
' RegressionUDF_r2
' RegressionUDF_Se
' RegressionUDF_n
' INTERFACE METHODS LIST ------------------------------------------------------------'
' RegressionGet
' RegressionGetY
' =========================== CLASS REGRESSION STRUCTURE ============================'

' =========================== CLASS REGRESSION STRUCTURE ============================'
' ABSTRACT IMPLEMENTATIONS ----------------------------------------------------------'
Implements ICoreController

' OBJECT VARIABLES (GLOOBJ_) --------------------------------------------------------'
Private EHGLOBAL As AppErrorHandler

' VARIABLES (GLOStr_) ---------------------------------------------------------------'
Private GLOBOO_EH As Boolean
Private GLOBOO_FORMULAS As Boolean
Private GLOVEC_REG As Variant

' CONSTANTS ERROR MESSAGES (GLOSTR_ERR_) --------------------------------------------'
Private Type GLOTYP_ERRORS
  DataYNo As String
  DataYless3 As String
  DataXDiferent As String
  DataDiv0 As String
End Type
Private GLOTypErr As GLOTYP_ERRORS





' PROPERTIES ------------------------------------------------------------------------'
' ABSTRACT INTERFACE PROPERTIES IMPLEMENTATION --------------------------------------'
Public Property Set ICoreController_ErrorHandler(RHS As AppErrorHandler)
  
  Set EHGLOBAL = Nothing
  Set EHGLOBAL = RHS
  Let GLOBOO_EH = (Not RHS Is Nothing)

End Property





' CONSTRUCTOR/DESTRUCTOR ------------------------------------------------------------'
Private Sub Class_Initialize()
  
  Dim MDL As ModelExcliqliteDatasheet
  
  Set EHGLOBAL = New AppErrorHandler
  Set MDL = New ModelExcliqliteDatasheet
  Set MDL.ErrorHandler = EHGLOBAL
  Let GLOBOO_FORMULAS = MDL.Formulas
  Set MDL = Nothing
  
  Let GLOTypErr.DataYNo = "Faltan los datos de orígen (Y)."
  Let GLOTypErr.DataYless3 = "Los datos de orígen (Y) deben ser como mínimo tres."
  Let GLOTypErr.DataXDiferent = "Los datos independientes (X) deben tener la misma cantidad que los de orígen (Y)."
  Let GLOTypErr.DataDiv0 = "¡Error, división entre cero!"

End Sub

Private Sub Class_Terminate()
  
  If GLOBOO_EH = False Then Set EHGLOBAL = Nothing
  Let GLOVEC_REG = Empty

End Sub





' CORE METHODS PRIVATE --------------------------------------------------------------'
' CORE METHODS LIST (PRIVATE) -------------------------------------------------------'
' 1. [x] VALIDATIONS
Private Function Regression_Core_UDF_validation( _
  ByVal RVarY As Variant, _
  ByVal RVarX As Variant, _
  ByVal RVarExtrapolation As Variant, _
  ByRef RVarErrUDF As Variant) As Boolean

  Dim RES As AppResources_current
  
  Dim i As Long, j As Long
  '1. [x]: Verify that all parameters are numbers
  '2. [x]: if X exists, its length must to have same length than Y

  On Error GoTo EH
  Let RVarErrUDF = VBA.CVErr(xlErrValue)
  Let Regression_Core_UDF_validation = False

  Set RES = New AppResources_current
  Set RES.ErrorHandler = EHGLOBAL
  'Y
  Call RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(RVarY, i, RBooIncludeZeroes:=True)
  If i = 0 Then Call EHGLOBAL.ErrorHandlerSet(0, GLOTypErr.DataYNo): GoTo EH
  If i < 3 Then Call EHGLOBAL.ErrorHandlerSet(1, GLOTypErr.DataYless3): GoTo EH

  'X
  Call RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(RVarX, j, RBooIncludeZeroes:=True)
  If j > 0 Then
    If Not i = j Then Call EHGLOBAL.ErrorHandlerSet(2, GLOTypErr.DataXDiferent): GoTo EH
  End If

  Let Regression_Core_UDF_validation = True

EH:
  Call EHGLOBAL.ErrorHandlerRaise("REG::-core_udf_validation")
End Function





' 2. [x] PREPARATIONS
Private Function Regression_Core_UDF_preparation( _
  ByRef RVarYSrc As Variant, _
  ByRef RVarXSrc As Variant, _
  ByRef RVarESrc As Variant, _
  ByRef RVarY As Variant, _
  ByRef RVarX As Variant, _
  ByRef RVarExtrapolation As Variant) As Boolean

  Dim RES As AppResources_current
  
  Dim i As Long

  On Error GoTo EH
  Let Regression_Core_UDF_preparation = False

  Set RES = New AppResources_current
  Set RES.ErrorHandler = EHGLOBAL
  Let RVarY = Regression_Core_UDF_preparation_set_y(RVarYSrc, i, RES)
  Let RVarX = Regression_Core_UDF_preparation_set_x(RVarXSrc, i, RES)
  Let RVarExtrapolation = Regression_Core_UDF_preparation_set_e(RVarESrc, i, RES)

  Let Regression_Core_UDF_preparation = True

EH:
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise("APPREG::-core_udf_preparation")
End Function

Private Function Regression_Core_UDF_preparation_set_x(ByVal RVarX As Variant, ByVal RVarLengthY As Long, ByVal RES As AppResources_current) As Variant
  Dim i As Long
  Dim VecU As Variant
  Dim LngLenghtX As Long

  On Error GoTo EH
  Let VecU = Array(0)
  Let Regression_Core_UDF_preparation_set_x = VecU
  
  Let VecU = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(RVarX, i, RValueType:=eqlResNumbers, RBooAbs:=False, RBooIncludeZeroes:=True, RRngRC:=eqlResRngCol)
  
  Let LngLenghtX = IIf(i = 0, RVarLengthY, i) - 1
  
  If i = 0 Then
    For i = 0 To LngLenghtX
      ReDim Preserve VecU(i)
      Let VecU(i) = i + 1
    Next i
  End If
  
  Let Regression_Core_UDF_preparation_set_x = VecU

EH:
  Erase VecU
  Call EHGLOBAL.ErrorHandlerRaise("APPREG::-core_udf_presparation_setx")
End Function

Private Function Regression_Core_UDF_preparation_set_y(ByVal RVarY As Variant, ByRef RLngCounter As Long, ByVal RES As AppResources_current) As Variant
  On Error GoTo EH
  Let Regression_Core_UDF_preparation_set_y = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(RVarY, RLngCounter, RValueType:=eqlResNumbers, RBooIncludeZeroes:=True)
  
EH:
End Function

Private Function Regression_Core_UDF_preparation_set_e(ByVal RVarE As Variant, ByRef RLngCounter As Long, ByVal RES As AppResources_current) As Variant
  Dim E As Variant
  
  On Error GoTo EH
  Let E = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(RVarE, RLngCounter, RBooIncludeZeroes:=True)
  Let Regression_Core_UDF_preparation_set_e = E(0)
  
EH:
End Function






' 3. [x] RESOLVE (validate and prepare itself and resolve)
Private Function Regression_Core_UDF_resolve_regression( _
  ByVal RStrWhat As String, _
  ByVal RVarY As Variant, _
  ByVal RVarX As Variant, _
  ByVal RVarExtrapolation As Variant, _
  Optional RBooLocal As Boolean = False) As Variant

  Dim mX As Variant, mY As Variant, mEx As Variant
  Dim mEXs As Double, mEYs As Double, mEX2s As Double, mEY2s As Double, mEXYs As Double, mDivAux As Double
  Dim mN As Double, mA As Double, mB As Double

  On Error GoTo EH

  'Validation
  If Not Regression_Core_UDF_validation(RVarY, RVarX, RVarExtrapolation, Regression_Core_UDF_resolve_regression) Then GoTo EH

  'Preparation
  If Not Regression_Core_UDF_preparation(RVarY, RVarX, RVarExtrapolation, mY, mX, mEx) Then GoTo EH

  'Resolve
  If Not Regression_Core_UDF_resolve_regression_calculate(mX, mY, mN, mEXs, mEYs, mEX2s, mEY2s, mEXYs, mDivAux, Regression_Core_UDF_resolve_regression) Then: GoTo EH

  Let mA = Regression_Core_UDF_resolve_regression_calculate_a(mEXs, mEYs, mEX2s, mEXYs, mDivAux)
  Let mB = Regression_Core_UDF_resolve_regression_calculate_b(mEXs, mEYs, mEXYs, mN, mDivAux)

  Select Case RStrWhat
    Case "e": Regression_Core_UDF_resolve_regression = Regression_Core_UDF_resolve_regression_calculate_tendence_equation(mA, mB, mEx)
    Case "a": Regression_Core_UDF_resolve_regression = mA
    Case "b": Regression_Core_UDF_resolve_regression = mB
    Case "r": Regression_Core_UDF_resolve_regression = Regression_Core_UDF_resolve_regression_calculate_r(mEXs, mEYs, mEX2s, mEY2s, mEXYs, mN)
    Case "r2": Regression_Core_UDF_resolve_regression = Regression_Core_UDF_resolve_regression_calculate_r2(mEXs, mEYs, mEX2s, mEY2s, mEXYs, mN)
    Case "se": Regression_Core_UDF_resolve_regression = Regression_Core_UDF_resolve_regression_calculate_se(mA, mB, mEYs, mEY2s, mEXYs, mN)
    Case "n": Regression_Core_UDF_resolve_regression = mN
    Case Else: Let Regression_Core_UDF_resolve_regression = VBA.CVErr(xlErrNA)
  End Select

EH:
  Let mX = Empty
  Let mY = Empty
  Let mEx = Empty
  If RBooLocal Then
    If Not VBA.Err.Number = 0 Then Call EHGLOBAL.ErrorHandlerRaise("APPREG::core_udf_resolve_regression_calculate_r")
  Else
    #If Debugging Then
      Debug.Print VBA.Err.Description
    #End If
    On Error GoTo -1: Call VBA.Err.Clear
  End If

End Function

Private Function Regression_Core_UDF_resolve_regression_calculate( _
  ByVal RX As Variant, _
  ByVal RY As Variant, _
  ByRef Rn As Double, _
  ByRef REXs As Double, _
  ByRef REYs As Double, _
  ByRef REX2s As Double, _
  ByRef REY2s As Double, _
  ByRef REXYs As Double, _
  ByRef RDivAux As Double, _
  ByRef RVarErrUDF As Variant) As Boolean

  Dim i As Long, j As Long

  On Error GoTo EH
  Let RVarErrUDF = VBA.CVErr(xlErrValue)
  Let Regression_Core_UDF_resolve_regression_calculate = False

  Let Rn = UBound(RY) + 1: Let REXs = 0: Let REYs = 0: Let REX2s = 0: Let REY2s = 0: Let REXYs = 0: Let j = 0
  For i = LBound(RY) To UBound(RY)
    Let REXs = REXs + RX(i)
    Let REYs = REYs + RY(i)
    Let REX2s = REX2s + (RX(i) ^ 2)
    Let REY2s = REY2s + (RY(i) ^ 2)
    Let REXYs = REXYs + (RX(i) * RY(i))
  Next

  Let RDivAux = ((Rn * REX2s) - ((REXs) ^ 2))
  If RDivAux = 0 Then Call EHGLOBAL.ErrorHandlerSet(9, GLOTypErr.DataDiv0): Let RVarErrUDF = VBA.CVErr(xlErrDiv0): GoTo EH
  
  Let Regression_Core_UDF_resolve_regression_calculate = True

EH:
  Call EHGLOBAL.ErrorHandlerRaise("APPERR::-core_udf_resolve_regression_calculate")
End Function

Private Function Regression_Core_UDF_resolve_regression_calculate_a( _
  ByVal aEXs As Double, _
  ByVal aEYs As Double, _
  ByVal aEX2s As Double, _
  ByVal aEXYs As Double, _
  ByVal divAux As Double) As Double
  
  On Error GoTo EH
  Let Regression_Core_UDF_resolve_regression_calculate_a = ((aEYs * aEX2s) - (aEXs * aEXYs)) / divAux

EH:
  Call EHGLOBAL.ErrorHandlerRaise("APPREG::core_udf_resolve_regression_calculate_a")
End Function

Private Function Regression_Core_UDF_resolve_regression_calculate_b( _
  ByVal bEXs As Double, _
  ByVal bEYs As Double, _
  ByVal bEXYs As Double, _
  ByVal bn As Double, _
  ByVal divAux As Double) As Double
  
  On Error GoTo EH
  Let Regression_Core_UDF_resolve_regression_calculate_b = ((bn * bEXYs) - (bEXs * bEYs)) / divAux

EH:
  Call EHGLOBAL.ErrorHandlerRaise("APPREG::core_udf_resolve_regression_calculate_b")
End Function

Private Function Regression_Core_UDF_resolve_regression_calculate_tendence_equation( _
  ByVal eA As Double, _
  ByVal eB As Double, _
  ByVal ext As Double) As Double

  On Error GoTo EH
  Let Regression_Core_UDF_resolve_regression_calculate_tendence_equation = eA + (eB * ext)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("APPREG::core_udf_resolve_regression_calculate_tendence_equation")

End Function

Private Function Regression_Core_UDF_resolve_regression_calculate_r( _
  ByVal REXs As Double, _
  ByVal REYs As Double, _
  ByVal REX2s As Double, _
  ByVal REY2s As Double, _
  ByVal REXYs As Double, _
  ByVal Rn As Double) As Double

  Dim rtop, rdown As Double

  On Error GoTo EH
  Let Regression_Core_UDF_resolve_regression_calculate_r = 0
  Let rtop = ((Rn * REXYs) - (REXs * REYs))
  Let rdown = ((Rn * REX2s) - (REXs ^ 2)) * ((Rn * REY2s) - (REYs ^ 2))
  If Not rdown = 0 Then
    Let rdown = VBA.Sqr(VBA.Abs(rdown))
    Let Regression_Core_UDF_resolve_regression_calculate_r = rtop / rdown
  End If

EH:
  Call EHGLOBAL.ErrorHandlerRaise("APPREG::core_udf_resolve_regression_calculate_r")
End Function

Private Function Regression_Core_UDF_resolve_regression_calculate_r2( _
  ByVal REXs As Double, _
  ByVal REYs As Double, _
  ByVal REX2s As Double, _
  ByVal REY2s As Double, _
  ByVal REXYs As Double, _
  ByVal Rn As Double) As Double

  Dim rtop, rdown As Double

  On Error GoTo EH
  Let Regression_Core_UDF_resolve_regression_calculate_r2 = 0
  Let rtop = ((Rn * REXYs) - (REXs * REYs)) ^ 2
  Let rdown = ((Rn * REX2s) - (REXs ^ 2)) * ((Rn * REY2s) - (REYs ^ 2))
  If Not rdown = 0 Then
    Let Regression_Core_UDF_resolve_regression_calculate_r2 = rtop / rdown
  End If

EH:
  Call EHGLOBAL.ErrorHandlerRaise("APPREG::core_udf_resolve_regression_calculate_r2")
End Function

Private Function Regression_Core_UDF_resolve_regression_calculate_se( _
  ByVal ra As Double, _
  ByVal rb As Double, _
  ByVal REYs As Double, _
  ByVal REY2s As Double, _
  ByVal REXYs As Double, _
  ByVal Rn As Double) As Double

  Dim rtop, rdown As Double

  On Error GoTo EH
  Let Regression_Core_UDF_resolve_regression_calculate_se = 0
  Let rtop = ((REY2s - (ra * REYs) - (rb * REXYs)))
  Let rdown = Rn - 2
  If Not rdown = 0 Then
    Let Regression_Core_UDF_resolve_regression_calculate_se = VBA.Sqr(VBA.Abs(rtop / rdown))
  End If

EH:
  Call EHGLOBAL.ErrorHandlerRaise("APPREG::core_udf_resolve_regression_calculate_se")
End Function






' 4. [ ] MODULE RESOLUTIONS
Private Function Regression_Core_Main(ByRef BVarRequest As Variant, ByRef BVarResponse As Variant) As Boolean
  
  Dim VecSources As Variant, VecRegression As Variant, VecCorelations As Variant
  Dim BooSources As Boolean, BooCorelations As Boolean
  Dim j As Long

  Const sFN As String = "APPREG::-core_main"

  On Error GoTo EH
  Let Regression_Core_Main = False

  'Resolution
  'Resolve Sources
  If Not Regression_Core_Spreadmodule_generate_sources(BVarRequest, VecSources, j, BooSources) Then GoTo EH
  
  'Resolve Regression
  If Not Regression_Core_Spreadmodule_generate_regression(BVarRequest, VecRegression, j, BooSources) Then GoTo EH
  
  'Resolve corelations
  If Not Regression_Core_Spreadmodule_generate_corelations(BVarRequest, VecCorelations, j, BooCorelations, BooSources) Then GoTo EH

  'Return
  Let BVarResponse = Array(BooSources, True, BooCorelations, VecSources, VecRegression, VecCorelations)
  Let Regression_Core_Main = True

EH:
  Let VecSources = Empty
  Let VecRegression = Empty
  Let VecCorelations = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Regression_Core_Spreadmodule_generate_regression( _
  ByRef RVarRequest As Variant, _
  ByRef RVarResponse As Variant, _
  ByRef RxJ As Long, _
  ByVal RBooSrc As Boolean) As Boolean

  'Calculates all regression excercises
  'Data required: RVarRequest
  Dim RngBox As Range
  Dim VecExtrapolation As Variant
  Dim VecX As Variant, VecY As Variant, VecE As Variant
  Dim StrX As String, StrY As String, StrE As String
  Dim Dbla As Double, Dblb As Double
  
  Dim i As Long, j As Long, m As Long, k As Long, n As Long

  Dim BooX2 As Boolean, BooY2 As Boolean, BooXY As Boolean, BooJustReg As Boolean, BooSrc As Boolean

  Const sFN As String = "APPREG::-core_spreadmodule_generate_regression"

  'Incoming data: 0-X '1-Y '2-E '3-ShowAll '4-BooJustReg '5-X2 '6-Y2 '7-XY '8-a '9-b '10-r '11-r2 '12-Se '13-Graph '14-SheetGraph '15-NewSheet '13-RngBox
  'Capture basic values (0-2): X2, Y2, XY
  On Error GoTo EH
  Let Regression_Core_Spreadmodule_generate_regression = False
  
  'Capture basic values
  Let BooSrc = (RxJ > 1) 'If j>1 then exists Source table FIXME: DELETE BOOSRC
  Let BooJustReg = VBA.CBool(RVarRequest(4)) 'If j>1 then exists Source table
  Let StrX = VBA.CStr(RVarRequest(0))  'X
  Let StrY = VBA.CStr(RVarRequest(1))  'Y
  Let StrE = VBA.CStr(RVarRequest(2))  'E
  
  'Calculate a, b
  'Let VecX = Empty: Let VecY = Empty
  Let Dbla = VBA.CDbl(Regression_Core_UDF_resolve_regression("a", StrY, StrX, StrE, True))
  Let Dblb = VBA.CDbl(Regression_Core_UDF_resolve_regression("b", StrY, StrX, StrE, True))

  'Extrapolation
  Let VecE = VBA.Split(StrE, ";")
  Let VecY = VBA.Split(StrY, ";")
  Let k = UBound(VecY) + 3  'Extrapolation A1 range table
  Let VecY = Empty
  Let m = UBound(VecE) + 1  'Extrapolation A1 range table
  Let n = IIf(BooSrc, RxJ, 0) 'Results A1 range table
  Let j = n
  ReDim VecExtrapolation(0 To m, 0 To 1)
  Let VecExtrapolation(0, 0) = "X'"
  Let VecExtrapolation(0, 1) = "Y'"

  If GLOBOO_FORMULAS Then '|If BooJustReg is false
    Set RngBox = RVarRequest(16)
    For i = 1 To m
      Let VecExtrapolation(i, 0) = VecE(i - 1)
      If BooJustReg Or RBooSrc = False Then
        Let VecExtrapolation(i, 1) = "=" & VBA.Replace(Dbla, ",", ".") & "+(" & VBA.Replace(Dblb, ",", ".") & "*" & RngBox.Range("A1").Offset(i + 1, 0).Address(False, False) & ")"
      Else
        Let VecExtrapolation(i, 1) = "=EQL_REGRESION_LINEAL(" & _
            RngBox.Range("C3:C" & k).Address & ", " & _
            RngBox.Range("B3:B" & k).Address & ", " & _
            RngBox.Range("A1").Offset(i + 1, j).Address(False, False) & ")"
      End If
    Next i
  Else
    For i = 1 To m
      Let VecExtrapolation(i, 0) = VecE(i - 1)
      Let VecExtrapolation(i, 1) = Dbla + (Dblb * VecE(i - 1))
    Next i
  End If

  'Return
  Let RxJ = RxJ + 4 '2 cols: X', Y', 1: white column, 1: Start next calculation
  Let RVarResponse = VecExtrapolation
  Let Regression_Core_Spreadmodule_generate_regression = True

EH:
  Let VecExtrapolation = Empty
  Let VecX = Empty
  Let VecY = Empty
  Let VecE = Empty
  Set RngBox = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Regression_Core_Spreadmodule_generate_sources( _
  ByRef RVarRequest As Variant, _
  ByRef RVarResponse As Variant, _
  ByRef RxJ As Long, _
  ByRef RBooExistsSrc As Boolean) As Boolean

  'Calculates all base regression excercises
  'Data required: X, Y, E, BooX^2, BooY^2, BooXY
  Dim RngBox As Range
  
  Dim VecXs As Variant, VecYs As Variant, VecEs As Variant
  Dim StrXs As String, StrYs As String, StrEs As String
  Dim vHeaders() As Variant, vAux() As Variant
  Dim i As Long, j As Long, m As Long, k As Long, n As Long
  Dim BooX2 As Boolean, BooY2 As Boolean, BooXY As Boolean
  
  Const sFN As String = "APPREG::-core_spreadmodule_generate_sources"
  
  On Error GoTo EH
  Let Regression_Core_Spreadmodule_generate_sources = False
  
  'Incoming data: 0-X '1-Y '2-E '3-ShowAll '4-JustRegression '5-X2 '6-Y2 '7-XY '8-a '9-b '10-r '11-r2 '12-Se '13-Graph '14-SheetGraph '15-NewSheet '13-RngBox
  'Capture basic values (0-2): X2, Y2, XY
  Let StrXs = VBA.CStr(RVarRequest(0))  'X
  Let StrYs = VBA.CStr(RVarRequest(1))  'Y
  Let StrEs = VBA.CStr(RVarRequest(2))  'E
  Let BooX2 = RVarRequest(5) 'X
  Let BooY2 = RVarRequest(6) 'Y
  Let BooXY = RVarRequest(7) 'E
  Let RBooExistsSrc = (BooX2 Or BooY2 Or BooXY Or RVarRequest(4) = False) 'Last: JustReg
  
  Let j = 0
  If Not Regression_Core_assets_getmainvectors(StrXs, StrYs, StrEs, VecXs, VecYs, VecEs, True) Then GoTo EH 'Error!
  'Updating values in Request
  Let RVarRequest(0) = StrXs 'X
  Let RVarRequest(1) = StrYs 'Y
  Let RVarRequest(2) = StrEs 'E
  If RBooExistsSrc Then
    Let m = UBound(VecYs) + 1 'rows
    Let n = 2 + (IIf(BooX2, 1, 0) + IIf(BooY2, 1, 0) + IIf(BooXY, 1, 0)) 'columns
  
    Let vAux = Array("N", "X", "Y", "X2", "Y2", "XY")
    For i = 0 To UBound(vAux)
      If i < 3 Then
        ReDim Preserve vHeaders(j): Let vHeaders(j) = vAux(i): Let j = j + 1
      Else
        If VBA.CBool(RVarRequest(i + 2)) Then ReDim Preserve vHeaders(j): Let vHeaders(j) = vAux(i): Let j = j + 1
      End If
    Next i
    Erase vAux
    
    'Main operations to extrapolation
    If GLOBOO_FORMULAS Then Set RngBox = RVarRequest(16)
    ReDim RVarResponse(0 To m, 0 To n)
    For j = 0 To n
      For i = 0 To m
        If i = 0 Then
          Let RVarResponse(i, j) = vHeaders(j)
        Else
          If j = 0 Then Let RVarResponse(i, j) = i
          If j = 1 Then Let RVarResponse(i, j) = VecXs(i - 1)
          If j = 2 Then Let RVarResponse(i, j) = VecYs(i - 1)
          If GLOBOO_FORMULAS Then
            If RVarResponse(0, j) = "X2" Then Let RVarResponse(i, j) = "=" & RngBox.Range("B" & i + 2).Address & "^2"
            If RVarResponse(0, j) = "Y2" Then Let RVarResponse(i, j) = "=" & RngBox.Range("C" & i + 2).Address & "^2"
            If RVarResponse(0, j) = "XY" Then Let RVarResponse(i, j) = "=" & RngBox.Range("B" & i + 2).Address & "*" & RngBox.Range("C" & i + 2).Address
          Else
            If RVarResponse(0, j) = "X2" Then Let RVarResponse(i, j) = VecXs(i - 1) ^ 2
            If RVarResponse(0, j) = "Y2" Then Let RVarResponse(i, j) = VecYs(i - 1) ^ 2
            If RVarResponse(0, j) = "XY" Then Let RVarResponse(i, j) = VecXs(i - 1) * VecYs(i - 1)
          End If
        End If
      Next i
    Next j
    Erase vHeaders
  End If
  
  Let RxJ = IIf(j = 0, 1, j + 1) 'Column zero to next table or column j+1
  Let Regression_Core_Spreadmodule_generate_sources = True
  
EH:
  Let VecXs = Empty
  Let VecYs = Empty
  Let VecEs = Empty
  Erase vAux
  Erase vHeaders
  Set RngBox = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Regression_Core_Spreadmodule_generate_corelations( _
  ByRef RVarRequest As Variant, _
  ByRef RVarResponse As Variant, _
  ByRef RxJ As Long, _
  ByRef RBooExistsCor As Boolean, _
  ByVal RBooSrc As Boolean) As Boolean

  Dim RngBox As Range
  
  Dim VecXs As Variant, VecYs As Variant, VecEs As Variant
  Dim StrXs As String, StrYs As String, StrEs As String
  Dim VecBooCor As Variant
  Dim VecResHeaders As Variant, VecResResults As Variant
  Dim VecAux1 As Variant, VecAux2 As Variant
  
  Dim StrRangeAux As String
  Dim i As Long, j As Long, m As Long, k As Long, n As Long
  Dim Booa As Boolean, Boob As Boolean, Boor As Boolean, Boor2 As Boolean, BooSe As Boolean
  
  Const sFN As String = "APPREG::-core_spreadmodule_generate_corelations"
  
  On Error GoTo EH
  Let Regression_Core_Spreadmodule_generate_corelations = False
  
  'Incoming data: 0-X '1-Y '2-E '3-ShowAll '4-JustRegression '5-X2 '6-Y2 '7-XY '8-a '9-b '10-r '11-r2 '12-Se '13-Graph '14-SheetGraph '15-NewSheet '16-RngBox
  'Capture basic values: X2, Y2, XY, a, b, r, r2, se
  Let StrXs = VBA.CStr(RVarRequest(0))  'X
  Let StrYs = VBA.CStr(RVarRequest(1))  'Y
  Let StrEs = VBA.CStr(RVarRequest(2))  'E
  Let Booa = RVarRequest(8) 'X
  Let Boob = RVarRequest(9) 'Y
  Let Boor = RVarRequest(10) 'E
  Let Boor2 = RVarRequest(11) 'X
  Let BooSe = RVarRequest(12) 'Y
  Let RBooExistsCor = (Booa Or Boob Or Boor Or Boor2 Or BooSe And VBA.CBool(RVarRequest(4)) = False) 'Last: JustReg
  
  If RBooExistsCor Then
    Let VecYs = VBA.Split(StrYs, ";")
    Let m = UBound(VecYs): Let VecYs = Empty
    Let VecBooCor = Array(Booa, Boob, Boor, Boor2, BooSe)

    'Extrapolation results
    Let VecAux1 = Array("Coeficiente a:", "Coeficiente b:", "Coeficiente de correlación (r):", "Coeficiente de determinación (r2):", "Error estándar de la determinación (Se):")
    Let VecAux2 = Array("a", "b", "r", "r2", "se")
    Set RngBox = RVarRequest(16)
    Let StrRangeAux = RngBox.Range("C3:C" & m + 3).Address & ", " & RngBox.Range("B3:B" & m + 3).Address
    ReDim VecResHeaders(0)
    ReDim VecResResults(0)
    Let i = 0
    For j = 0 To UBound(VecAux1)
      If VBA.CBool(VecBooCor(j)) Then
        ReDim Preserve VecResHeaders(i)
        ReDim Preserve VecResResults(i)
        Let VecResHeaders(i) = VecAux1(j)
        If GLOBOO_FORMULAS Then
          If RBooSrc Then Let VecResResults(i) = "=EQL_REGRESION_LINEAL_" & VecAux2(j) & "(" & StrRangeAux & ")"
          If RBooSrc = False Then Let VecResResults(i) = "=EQL_REGRESION_LINEAL_" & VecAux2(j) & "(""" & VBA.Replace(StrYs, ",", ".") & """, """ & VBA.Replace(StrXs, ",", ".") & ")"
        Else
          Let VecResResults(i) = VBA.CDbl(Regression_Core_UDF_resolve_regression(VecAux2(j), StrYs, StrXs, StrEs, True))
        End If
        Let i = i + 1
      End If
    Next j
    Let VecAux1 = Empty
    Let VecAux2 = Empty
    Let VecBooCor = Empty
    
    'Fill response vector
    ReDim RVarResponse(0 To UBound(VecResHeaders) + 1, 0 To 1)
    Let RVarResponse(0, 0) = "Coeficientes"
    Let RVarResponse(0, 1) = "Valores"
    For i = 0 To UBound(RVarResponse) - 1
      Let RVarResponse(i + 1, 0) = VecResHeaders(i)
      Let RVarResponse(i + 1, 1) = VecResResults(i)
    Next i
    Let VecResHeaders = Empty
    Let VecResResults = Empty
  End If
  
  'return
  Let Regression_Core_Spreadmodule_generate_corelations = True

EH:
  Let VecXs = Empty
  Let VecYs = Empty
  Let VecEs = Empty
  Let StrXs = Empty
  Let StrYs = Empty
  Let StrEs = Empty
  Let VecAux1 = Empty
  Let VecAux2 = Empty
  Let VecBooCor = Empty
  Let VecResHeaders = Empty
  Let VecResResults = Empty
  Set RngBox = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function





' ASSETS ----------------------------------------------------------------------------'
Private Function Regression_Core_assets_getmainvectors( _
  ByRef RStrX As String, _
  ByRef RStrY As String, _
  ByRef RStrE As String, _
  ByRef RVecX As Variant, _
  ByRef RVecY As Variant, _
  ByRef RVecE As Variant, _
  Optional RBooAsStrings As Boolean = False) As Boolean
  
  Dim RES As AppResources
  
  Dim i As Long, j As Long, k As Long
  
  Const sFN As String = "APPREG::-core_assets_getmainvectors"
  
  On Error GoTo EH
  Let Regression_Core_assets_getmainvectors = False
  
  Set RES = New AppResources
  Let RVecY = RES.resArrayGetDataXCharSeparatedStringFromVariantsAsVector(RStrY, i, RValueType:=eqlResNumbers, RBooIncludeZeroes:=True) 'Y
  Let RVecX = RES.resArrayGetDataXCharSeparatedStringFromVariantsAsVector(RStrX, j, RValueType:=eqlResNumbers, RBooIncludeZeroes:=True) 'X
  Let RVecE = RES.resArrayGetDataXCharSeparatedStringFromVariantsAsVector(RStrE, k, RValueType:=eqlResNumbers, RBooIncludeZeroes:=True) 'E
  Set RES = Nothing
  
  If i <= 0 Then GoTo EH 'There is no elements in vector!
  
  If j <= 0 Then 'Refill X
    Let RVecX = Empty
    ReDim RVecX(0)
    For i = 0 To UBound(RVecY)
      ReDim Preserve RVecX(i)
      Let RVecX(i) = i + 1
    Next i
  End If
  
  If k <= 0 Then 'Refill E
    Let RVecE = Empty
    ReDim RVecE(0)
    Let RVecE(0) = RVecX(UBound(RVecX)) + 1
  End If
  
  If RBooAsStrings Then
    Let RStrX = VBA.Join(RVecX, ";")
    Let RStrY = VBA.Join(RVecY, ";")
    Let RStrE = VBA.Join(RVecE, ";")
  End If
  
  Let Regression_Core_assets_getmainvectors = True

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function





' ERROR ASSERTSIONS -----------------------------------------------------------------'
Private Sub Debug_assert(ByVal Cond As Boolean, Optional Origin As String = VBA.vbNullString)
  #If Debbuging Then
    Debug.Assert Cond
    If Not Cond Then Debug.Print "Assert on: " & Origin
  #End If
End Sub
' CORE METHODS PRIVATE --------------------------------------------------------------'





' METHODS PUBLIC --------------------------------------------------------------------'
Public Function RegressionGetY(ByVal RStrY As String, ByVal RStrX As String, ByVal RStrExtrapolation As String) As Double
  On Error GoTo EH
  Let RegressionGetY = VBA.CDbl(Regression_Core_UDF_resolve_regression("e", RStrY, RStrX, RStrExtrapolation, True))

EH:
  If GLOBOO_EH Then
    Call EHGLOBAL.ErrorHandlerRaise("APPREG::RegressionGetY")
  Else
    Call EHGLOBAL.ErrorHandlerDisplay("APPREGRESSION::RegressionGetY")
  End If
End Function

' ABSTRACT INTERFACE UI METHODS LIST AND UDF METHODS LIST (PUBLIC) ------------------'
Public Function ICoreController_GetSolution(ByVal CStrRequester As String, ByVal CVarRequest As Variant, CVarResponse As Variant) As Boolean
  
  On Error GoTo EH
  Let ICoreController_GetSolution = Regression_Core_Main(CVarRequest, CVarResponse)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("APPREG::IGetSolution")

End Function

Public Function ICoreController_GetSolutionUDF(ByVal CStrRequester As String, ByVal CVarRequest As Variant) As Variant
  
  On Error GoTo EH
  Let ICoreController_GetSolutionUDF = Regression_Core_UDF_resolve_regression(VBA.CStr(CVarRequest(0)), CVarRequest(1), CVarRequest(2), CVarRequest(3))
  
EH:
  Call EHGLOBAL.ErrorHandlerRaise("APPREG::IGetSolution")
  
End Function
' METHODS PUBLIC --------------------------------------------------------------------'
' =========================== CLASS REGRESSION STRUCTURE ============================'

