VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppExcliqBalance_current"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================= CLASS BALMET STRUCTURE =============================='
' ABSTRACT IMPLEMENTATIONS ----------------------------------------------------------'
' OBJECT VARIABLES (GLOOBJ_) --------------------------------------------------------'
' VARIABLES (GLOStr_) ---------------------------------------------------------------'
' VARIABLES CUSTOM (GLOEnu|Typ_) ----------------------------------------------------'
' CONSTANTS (GLOSTR_) ---------------------------------------------------------------'
' CONSTANTS ERROR MESSAGES (GLOSTR_ERR_) --------------------------------------------'
' PROPERTIES ------------------------------------------------------------------------'
' CONSTRUCTOR/DESTRUCTOR ------------------------------------------------------------'
' CORE METHODS LIST PRIVATE ---------------------------------------------------------'
' UDF CORE METHODS
' Balmet_core_udf
' Balmet_core_main_udf
' Balmet_core_main_udf_validate
' Balmet_core_main_udf_prepare
' Balmet_core_main_udf_calculate
' Balmet_core_main_udf_eco
' Balmet_core_main_udf_validate_eco
' Balmet_core_main_udf_prepare_eco
' Balmet_core_main_udf_calculate_eco
' APP CORE METHODS
' Balmet_core_main
' Balmet_core_main_app
' Balmet_core_main_app_eco
' Balmet_core_main_app_eco_spreadbalmet_get_grades_data
' Balmet_core_main_app_eco_spreadbalmet_get_weights_data
' Balmet_core_main_app_spreadbalmet_get_products_names
' Balmet_core_main_app_spreadbalmet_get_weights
' Balmet_core_main_app_spreadbalmet_get_weights_generic
' Balmet_core_main_app_spreadbalmet_get_weights_conventional
' Balmet_core_main_app_spreadbalmet_get_weights_cramer
' Balmet_core_main_app_spreadbalmet_get_weights_inverse
' Balmet_core_main_app_spreadbalmet_get_percents
' Balmet_core_main_app_spreadbalmet_get_volumes
' Balmet_core_main_app_spreadbalmet_get_volumes_generic
' Balmet_core_main_app_spreadbalmet_get_volumes_percents
' Balmet_core_main_app_spreadbalmet_get_grades
' Balmet_core_main_app_spreadbalmet_get_unities
' Balmet_core_main_app_spreadbalmet_get_fines
' Balmet_core_main_app_spreadbalmet_get_recoveries
' Balmet_core_main_app_spreadbalmet_get_ratio
' ASSETS ----------------------------------------------------------------------------'
' Balmet_assets_get_grades
' Balmet_assets_get_grades_cx
' Balmet_assets_check_names
' Balmet_assets_check_names_and_units
' Balmet_assets_check_names_and_units_cx
' ERRORS MANNAGEMENT AND ASSERTSIONS ------------------------------------------------'
' Debug_assert
' METHODS LIST PUBLIC ---------------------------------------------------------------'
' INTERFACE UDF METHODS LIST --------------------------------------------------------'
' ICoreController_GetSolutionUDF
' INTERFACE METHODS LIST ------------------------------------------------------------'
' ICoreController_GetSolution
' ============================= CLASS BALMET STRUCTURE =============================='


' ============================= CLASS BALMET STRUCTURE =============================='
' ABSTRACT IMPLEMENTATIONS ----------------------------------------------------------'
Implements ICoreController

' OBJECT VARIABLES (GLOOBJ_) --------------------------------------------------------'
Private EHGLOBAL As AppErrorHandler

' VARIABLES (GLOStr_) ---------------------------------------------------------------'
Private GLOBOO_EH As Boolean
Private GLOBOO_FORMULAS As Boolean

' VARIABLES CUSTOM (GLOEnu|Typ_) ----------------------------------------------------'
Public Enum EQLBAL_ENU_METHOD
  eqlBalConventional
  eqlBalCramer
  eqlBalInverseMatrix
End Enum

Public Enum EQLBAL_ENU_RESULT
  eqlBalWeights
  eqlBalWeightPercents
  eqlBalGradesHeads
  eqlBalUnities
  eqlBalFines
  eqlBalRecoveries
  eqlBalRatio
  eqlBalGradesHeadsCx
  eqlBalUnitiesCx
  eqlBalFinesCx
  eqlBalRecoveriesCx
  eqlBalVolume
  eqlBalVolumePercents
End Enum

Public Enum EQLBAL_ENU_RESULTECO
  eqlBalHeadsGrades
  eqlBalProdsGrades
End Enum

Public Enum EQLBAL_ENU_RESULTDIRECTION
  eqlBalVertical
  eqlBalHorizontal
End Enum

Private Enum EQLBAL_ENU_TYPECONTENT
  eqlBalSolids
  eqlBalVolumes
  eqlBalBoth
End Enum

Public Enum EQLBAL_ENU_TYPERESULT
  eqlBalJustGrades
  eqlBalJustGradesCx
  eqlBalJustGradesBoth
End Enum

Public Enum EQLBAL_ENU_GRADESVECTORTYPE
  eqlBalAs1D
  eqlBalAs2D
End Enum

Private GLOVEC_BALANCE As Variant

' CONSTANTS (GLOSTR_) ---------------------------------------------------------------'
Private Const CDBL_DM_TOP As Double = 10000
Private Const CDBL_PC_TOP As Double = 100
Private Const CDBL_GT_TOP As Double = 1000000
Private Const CDBL_MAX As Double = 10000000000000#

Private Const GLOBOO_BAL_LIMIT As Single = 5
Private Const GLOBOO_BAL_LIMITCX As Single = 3

' CONSTANTS ERROR MESSAGES (GLOSTR_ERR_) --------------------------------------------'
Private Const CUSTOM_ERROR = VBA.vbObjectError + 518





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

End Sub

Private Sub Class_Terminate()
  
  If GLOBOO_EH = False Then Set EHGLOBAL = Nothing
  Let GLOVEC_BALANCE = Empty

End Sub





' CORE METHODS PRIVATE --------------------------------------------------------------'
' UDF CORE METHODS
Private Function Balmet_core_udf(ByVal CStrRequester As String, ByVal CVarRequest As Variant) As Variant
  
  On Error GoTo EH
  Select Case CStrRequester
    Case "balmetudf": Let Balmet_core_udf = Balmet_core_main_udf(CVarRequest(0), CVarRequest(1), CVarRequest(2), CVarRequest(3), CVarRequest(4), CVarRequest(5), CVarRequest(6), CVarRequest(7), CVarRequest(8))
    Case "balmetudfeco": Let Balmet_core_udf = Balmet_core_main_udf_eco(CVarRequest(0), CVarRequest(1), CVarRequest(2), CVarRequest(3), CVarRequest(4), CVarRequest(5), CVarRequest(6), CVarRequest(7), CVarRequest(8), CVarRequest(9))
    Case Else: GoTo EH
  End Select

EH:
  If VBA.Err.Number <> 0 Then
    Debug.Print VBA.Err.Description
    On Error GoTo -1
    Call VBA.Err.Clear
  End If

End Function

Private Function Balmet_core_main(ByVal CStrRequester As String, ByVal CVarRequest As Variant, CVarResponse As Variant) As Boolean
  
  On Error GoTo EH
  ' Get Balmet Entirely Array finished
  Select Case CStrRequester
    Case "balmet": Let Balmet_core_main = Balmet_core_main_app(CVarRequest)
    Case "balmeteco": Let Balmet_core_main = Balmet_core_main_app_eco(CVarRequest, CVarResponse, CVarRequest(0)) 'TODO: OJO definir 0
    Case Else: Let Balmet_core_main = False
  End Select
  
  If Balmet_core_main Then CVarResponse = GLOVEC_BALANCE

EH:
  Let GLOVEC_BALANCE = Empty
  Call EHGLOBAL.ErrorHandlerRaise("BALMET::-main")

End Function

Private Function Balmet_core_main_udf( _
  ByVal BVarFeed As Variant, _
  ByVal BVarGrad As Variant, _
  ByVal BVarGradCx As Variant, _
  ByVal BVarElem As Variant, _
  ByVal BVarUnit As Variant, _
  ByVal BMethod As EQLBAL_ENU_METHOD, _
  ByVal BResult As EQLBAL_ENU_RESULT, _
  ByVal BDirection As EQLBAL_ENU_RESULTDIRECTION, _
  ByVal BAllResult As Boolean) As Variant
  
' Resolves a mxn balmet returning an mx1 matrix of weights _
  Requirements: _
  1. BVarFeed:   Balmet Feed, the weight feed of the process in Tons _
  2. BVarGrad: Balmet Grades, the mxn matrix of head, concentrates and tails grades in %, DM, g/L or g/T. This grades can include Complex minerals _
  3. BVarUnit: Balmet Units, a vector with next possible random values %, DM, g/L or g/T. This Units can include Complex minerals _
  4. BVarDensity [optional]: Balmet Density, a number to determine the volume of Balmet weights _
  5. BMethod [optional]: Balmet Method, a single number between 0 to 2, 0 (default) conventional balmet, 1 determinants, and 2 inverse matrix _
  6. BResult [optional]: Balmet result, it could be any of next: 0: Weights, 1: WeightPercents, 2: Recups, 3: Fines and 4: Ratios. _
  7. BDirection [optional]: Balmet result direction vector, it could be any of next: 0: vertically, 1: horizontally. _
  The BalmetUDF Method returns an mx1 matrix with all weights of the Balmet.
  
  Dim VecAux0 As Variant
  Dim BooComplex As Boolean, BooExistsGL As Boolean
  Dim EnuTypeFeed As EQLBAL_ENU_TYPECONTENT
  
  On Error GoTo EH
  Let Balmet_core_main_udf = VBA.CVErr(xlErrValue)
  Let BooComplex = False
  Let BooExistsGL = False
  
  'Forming the request data
  Let VecAux0 = Array(BVarFeed, BVarGrad, BVarGradCx, BVarElem, BVarUnit)
  
  'Sets GLOVEC_BALANCE with final result
  If Not Balmet_core_main_udf_validate(VecAux0, Balmet_core_main_udf, BooComplex, BooExistsGL) Then GoTo EH
  If Not Balmet_core_main_udf_prepare(VecAux0, Balmet_core_main_udf, BooComplex, BooExistsGL, EnuTypeFeed) Then GoTo EH
  If Not Balmet_core_main_udf_calculate(VecAux0, BMethod, BResult, BDirection, Balmet_core_main_udf, BooComplex, BooExistsGL, BAllResult, EnuTypeFeed) Then GoTo EH
  
  Let Balmet_core_main_udf = GLOVEC_BALANCE
  Let VecAux0 = Empty
  Let GLOVEC_BALANCE = Empty
  
EH:
  Let VecAux0 = Empty
  
End Function

Private Function Balmet_core_main_udf_validate( _
  ByVal BVarRequest As Variant, _
  ByRef BVarErrUDF As Variant, _
  ByRef BBooComplex As Boolean, _
  ByRef BBooGL As Boolean) As Boolean
  
  Dim RES As AppResources_current
  Dim RCT As AppResourcesController_current
  
  'Data required: (BVarFeed, BVarGrad, BVarGradCx, BVarElem, BVarUnit)
  Dim VecFeed As Variant, VecGrad As Variant, VecUnit As Variant, VecName As Variant, VecGradCx As Variant
  Dim StrUnitCx As String
  Dim StrAux0 As String
  Dim DblAux0 As Double
  Dim BooGenericProds As Boolean
  Dim EnuBaltype As EQLBAL_ENU_TYPE
  Dim i As Long, j As Long, k As Long, l As Long, m As Long
  Dim LngProducts As Long, LngColumns As Long, LngRows As Long, LngElements As Long, LngElementsAndCx As Long, LngElementsCx As Long, LngColumnsCx As Long
  Dim LngColumnsArr As Long
  Dim xUnitsType As Byte

  On Error GoTo EH
  Let Balmet_core_main_udf_validate = False

  'Initialize validations
  Set RES = New AppResources_current
  Set RES.ErrorHandler = EHGLOBAL
  Set RCT = New AppResourcesController_current
  Set RCT.ErrorHandler = EHGLOBAL
  
  'Grades
  'Extract all balmet and cx grades
  Let VecGrad = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(1), j, RValueType:=eqlResNumbers, RBooIncludeZeroes:=True, RBooAbs:=True)
  Let VecGradCx = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(2), k, RValueType:=eqlResNumbers, RBooIncludeZeroes:=True, RBooAbs:=True)
  If j = 0 Then Let BVarErrUDF = VBA.CVErr(xlErrNull): GoTo EH 'Grades required
    'If Not Balmet_assets_get_is_balmet(j, k) Then Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH 'Not a balmet
  If Not RCT.ResCtrIsBalmet(j, k) Then Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH 'Not a balmet
  'Get the type of balmet to evaluate
    'Let EnuBaltype = Balmet_assets_get_type(j, k) 'Balmet type
  Let EnuBaltype = RCT.ResCtrBalGetType(j, k) 'Balmet type
  If EnuBaltype = eqlBalNothing Then Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH 'Not a balmet
  'Get the balmet products
    'If Not Balmet_assets_get_products_bounds_fulldata( _
      j, _
      LngProducts, _
      LngColumns, _
      LngRows, _
      LngElements, _
      LngElementsAndCx, _
      LngElementsCx, _
      LngColumnsCx, _
      k _
    ) Then Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH 'Not a balmet
  If Not RCT.ResCtrBalGetProductsBoundsFulldata(j, LngProducts, LngColumns, LngRows, LngElements, LngElementsAndCx, LngElementsCx, LngColumnsCx, k) Then Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH   'Not a balmet
  Set RCT = Nothing

  'Names & Units
  'Ej.: (Zn, Pb, Cu, [Ag]) | (%, %, %, [DM])
  'Elements names (always will exist min one element and never repeated items, this validation includes cx elements)
  Let VecName = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(3), i, RValueType:=eqlResStrings, RValueAs:=eqlResStrings)
  Let VecUnit = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(4), j, RValueType:=eqlResStrings, RStrCustomRegex:="%", RValueAs:=eqlResStrings)
  
  If i > (LngColumns + LngColumnsCx) Or j > (LngColumns + LngColumnsCx) Or (j = 0) Then Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH 'Names or Units exceed the limit of products
  Let LngColumnsArr = LngColumns - 1
  If j < LngColumns Then  'Refill units with default % and if balmet is type cx refill it too
    For k = j To LngColumnsArr
      ReDim Preserve VecUnit(k)
      Let VecUnit(k) = "%"
    Next k
    Let j = k
  End If 'Else All ok
  If EnuBaltype = eqlBalWithComplex And j < (LngColumns + LngColumnsCx) Then
    For k = j To (LngColumnsArr + LngColumnsCx)
      ReDim Preserve VecUnit(k)
      Let VecUnit(k) = "g/T"
    Next k
  End If
  
  'Case Names: Check if Names has repeated items
  If RES.ArrayHasRepeatedElements(VecName) Then Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH 'Exists repeated elements!
  Let BooGenericProds = (i = 0) 'generic products
  Let BBooGL = RES.ArrayHasValue(VecUnit, "g/l", False, True) 'Any elements has g/l unit
  'Ej.: (Zn, Pb, [Ag]) | (%, %, [DM])
  For k = 0 To LngColumnsArr
    If i <= LngColumns And k >= i Then  'Complete the list with generic products
      ReDim Preserve VecName(k)
      Let VecName(k) = "B" & k + 1
    Else
      If Not Balmet_assets_check_names_and_units(VBA.CStr(VecName(k)), VBA.CStr(VecUnit(k)), BVarErrUDF) Then GoTo EH
    End If
  Next k
  If EnuBaltype = eqlBalWithComplex Then 'Complete the cx list with generic products
    For k = LngColumns To (LngColumnsArr + LngColumnsCx)
      If k > UBound(VecName) And UBound(VecName) < (LngColumns + LngColumnsCx) Then
        ReDim Preserve VecName(k)
        Let VecName(k) = "BX" & (k - k) + 1
      Else
        If Not Balmet_assets_check_names_and_units_cx(VBA.CStr(VecName(k)), VBA.CStr(VecUnit(k)), BVarErrUDF) Then GoTo EH
      End If
    Next k
  End If
  
  'Get grades in 2D
  Let VecGrad = Balmet_assets_get_grades(VecGrad, k, BResponseDimension:=eqlBalAs2D)
  If k = 0 Then Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH 'Balances up to ten products not allowed, diferent at mxn neighter
  If EnuBaltype = eqlBalWithComplex Then
    Let VecGradCx = Balmet_assets_get_grades_cx(VecGradCx, LngElements, k, BResponseDimension:=eqlBalAs2D)
    If k = 0 Then Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH 'Balances up to ten products not allowed, diferent at mxn neighter
  End If
  
  'Grades (always will exist min one grade)
  Let k = 0: Let l = 0: Let m = 0
  For j = 0 To UBound(VecGrad, 2) 'Check if each Grade value is correct and correspond with its Unit
    For i = 0 To UBound(VecGrad)
      Let StrAux0 = VBA.LCase(VecUnit(j))
      Let DblAux0 = VBA.CDbl(VecGrad(i, j))
      Select Case StrAux0
        Case "%"
          If DblAux0 > CDBL_PC_TOP Then Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH 'Unreasonable Grades
          Let k = k + IIf(DblAux0 = CDBL_PC_TOP, 1, 0)
          If k > 1 Then Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH 'Unreasonable Grades
        Case "dm"
          If DblAux0 > CDBL_DM_TOP Then Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH 'Unreasonable Grades
          Let l = l + IIf(DblAux0 = CDBL_DM_TOP, 1, 0)
          If l > 1 Then Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH 'Unreasonable Grades
        Case "g/t", "g/l"
          If DblAux0 > CDBL_GT_TOP Then Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH 'Unreasonable Grades
          Let m = m + IIf(DblAux0 = CDBL_GT_TOP, 1, 0)
          If m > 1 Then Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH 'Unreasonable Grades
      End Select
    Next i
    Let k = 0: Let l = 0: Let m = 0
  Next j

  'Grades cx
  If EnuBaltype = eqlBalWithComplex Then
    Let BBooComplex = True
    Let k = 0: Let l = 0: Let m = 0
    For j = 0 To UBound(VecGradCx, 2) 'Check if each Grade value is correct and correspond with it Unit
      Let StrUnitCx = VBA.LCase(VecUnit(j + LngColumns))
      For i = LBound(VecGradCx) To UBound(VecGradCx)
        Let DblAux0 = VBA.CDbl(VecGradCx(i, j))
        Select Case StrUnitCx
          Case "dm"
            If DblAux0 > CDBL_DM_TOP Then Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH 'Unreasonable Grades
            Let l = l + IIf(DblAux0 = CDBL_DM_TOP, 1, 0)
            If l > 1 Then Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH 'Unreasonable Grades
          Case "g/t", "g/l"
            If DblAux0 > CDBL_GT_TOP Then Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH 'Unreasonable Grades
            Let m = m + IIf(DblAux0 = CDBL_GT_TOP, 1, 0)
            If m > 1 Then Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH 'Unreasonable Grades
          Case Else
            Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH
        End Select
      Next i
      Let k = 0: Let l = 0: Let m = 0
    Next j
  End If

  'Feed
  'Check for Units if it has volumes
  Let xUnitsType = 0
  If BBooGL Then
    Let j = (UBound(VecUnit) + 1)
    Let i = RES.ArrayHasValueN(VecUnit, "g/l", False, True)
    Let xUnitsType = IIf(i = j, 1, 2) 'All g/T: EW | 'At least 1g/T: (W, V); If juts W, Fines = '-'
  End If
  'Calculate Feeds
  Let VecFeed = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(0), i, RValueType:=eqlResNumbers, RBooAbs:=True)
  If i = 0 Then Let BVarErrUDF = VBA.CVErr(xlErrNull): GoTo EH 'Feed required
  Select Case xUnitsType
    Case 0, 1
      Let DblAux0 = Application.WorksheetFunction.Sum(VecFeed)
      If DblAux0 = 0 Then Let BVarErrUDF = VBA.CVErr(xlErrNull): GoTo EH 'Feed required
      If DblAux0 > CDBL_MAX Then Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH 'Feed extremely big
    Case 2
      Let DblAux0 = VBA.Abs(VBA.CDbl(VecFeed(0)))
      If DblAux0 = 0 Then Let BVarErrUDF = VBA.CVErr(xlErrNull): GoTo EH 'Feed required
      If DblAux0 > CDBL_MAX Then Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH 'Feed extremely big
      If i > 1 Then
        Let DblAux0 = VBA.Abs(VBA.Abs(Application.WorksheetFunction.Sum(VecFeed)) - DblAux0)
        If DblAux0 > CDBL_MAX Then Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH 'Feed extremely big
      End If
  End Select
  Let VecFeed = Empty

  'Return
  Let Balmet_core_main_udf_validate = True

EH:
  Let VecFeed = Empty
  Let VecGrad = Empty
  Let VecName = Empty
  Let VecUnit = Empty
  Let VecGradCx = Empty
  Set RCT = Nothing
  Set RES = Nothing
  
End Function

Private Function Balmet_core_main_udf_prepare( _
  ByRef BVarRequest As Variant, _
  ByRef BVarErrUDF As Variant, _
  ByVal BBooComplex As Boolean, _
  ByVal BBooGL As Boolean, _
  ByRef BEnuTypeFeed As EQLBAL_ENU_TYPECONTENT) As Boolean
  
  'Data required: (BVarFeed, BVarGrad, BVarUnits)
  Dim RES As AppResources_current
  Dim RCT As AppResourcesController_current

  Dim VecFeed As Variant, VecGrad As Variant, VecName As Variant, VecUnit As Variant, _
      VecGradCx As Variant, VecNameCx As Variant, VecUnitCx As Variant, VecGradGL As Variant, _
      VecAux0 As Variant, VecAux1 As Variant
  Dim DblFeed As Double, DblFeedVol As Double
  Dim i As Long, j As Long, k As Long, xUnitsType As Long
  Dim LngProducts As Long, LngColumns As Long, LngRows As Long, LngElements As Long, LngElementsAndCx As Long, LngElementsCx As Long, LngColumnsCx As Long
  Dim LngColumnsArr As Long

  On Error GoTo EH
  Let Balmet_core_main_udf_prepare = False

  'Initialize validations
  Set RES = New AppResources_current
  Set RES.ErrorHandler = EHGLOBAL
  Set RCT = New AppResourcesController_current
  Set RCT.ErrorHandler = EHGLOBAL

  'Grades
  'Extract all balmet and cx grades
  Let VecGrad = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(1), j, RValueType:=eqlResNumbers, RBooIncludeZeroes:=True, RBooAbs:=True)
  Let VecGradCx = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(2), k, RValueType:=eqlResNumbers, RBooIncludeZeroes:=True, RBooAbs:=True)
  'Get the balmet products
'  Call Balmet_assets_get_products_bounds_fulldata( _
'    j, _
'    LngProducts, _
'    LngColumns, _
'    LngRows, _
'    LngElements, _
'    LngElementsAndCx, _
'    LngElementsCx, _
'    LngColumnsCx, _
'    k _
'  )
  Call RCT.ResCtrBalGetProductsBoundsFulldata(j, LngProducts, LngColumns, LngRows, LngElements, LngElementsAndCx, LngElementsCx, LngColumnsCx, k)
  Set RCT = Nothing

  'Names & Units
  'Ej.: (Zn, Pb, Cu, [Ag]) | (%, %, %, [DM])
  'Elements names (always will exist min one element and never repeated items, this validation includes cx elements)
  Let VecName = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(3), i, RValueType:=eqlResStrings, RValueAs:=eqlResStrings)
  Let VecUnit = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(4), j, RValueType:=eqlResStrings, RStrCustomRegex:="%", RValueAs:=eqlResStrings)
  
  Let LngColumnsArr = LngColumns - 1
  If j < LngColumns Then  'Refill units with default % and if balmet is type cx refill it too
    For k = j To LngColumnsArr
      ReDim Preserve VecUnit(k)
      Let VecUnit(k) = "%"
    Next k
    Let j = k
  End If 'Else All ok
  If BBooComplex And j < (LngColumns + LngColumnsCx) Then
    For k = j To (LngColumnsArr + LngColumnsCx)
      ReDim Preserve VecUnit(k)
      Let VecUnit(k) = "g/T"
    Next k
  End If
  
  'Case Names: Check if Names has repeated items
  For k = 0 To LngColumnsArr
    If i <= LngColumns And k >= i Then  'Complete the list with generic products
      ReDim Preserve VecName(k)
      Let VecName(k) = "B" & k + 1
    End If
  Next k
  If BBooComplex Then 'Complete the cx list with generic products
    For k = LngColumns To (LngColumnsArr + LngColumnsCx)
      If k > UBound(VecName) And UBound(VecName) < (LngColumns + LngColumnsCx) Then
        ReDim Preserve VecName(k)
        Let VecName(k) = "BX" & (k - k) + 1
      End If
    Next k
  End If
  
  'Get grades in 2D
  Let VecGrad = Balmet_assets_get_grades(VecGrad, k, BResponseDimension:=eqlBalAs2D)
  If BBooComplex Then Let VecGradCx = Balmet_assets_get_grades_cx(VecGradCx, LngElements, k, BResponseDimension:=eqlBalAs2D)
  
  'Feed
  'Check for Units if it has volumes
  Let xUnitsType = 0
  If BBooGL Then
    Let VecAux0 = RES.ArrayConcat(VecUnit, VecUnitCx)
    If Not VBA.IsArray(VecAux0) Then Let BVarErrUDF = VBA.CVErr(xlErrValue): GoTo EH 'Error where concat arr
    Let VecAux1 = RES.ArrayConcat(VecGrad, VecGradCx, RByCols:=True)
    If Not VBA.IsArray(VecAux1) Then Let BVarErrUDF = VBA.CVErr(xlErrValue): GoTo EH 'Error where concat arr
    'Get xUnitsType
    Let j = (UBound(VecAux0) + 1)
    Let i = RES.ArrayHasValueN(VecAux0, "g/l", False, True)
    Let xUnitsType = IIf(i = j, 1, 2) 'All g/T: EW | 'At least 1g/T: (W, V); If juts W, Fines = '-'
  End If
  Let BEnuTypeFeed = xUnitsType
  'Calculate Feeds
  Let VecFeed = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(0), i, RValueType:=eqlResNumbers, RBooAbs:=True)
  Select Case BEnuTypeFeed 'xUnitsType
    Case eqlBalVolumes, eqlBalSolids '0, 1
      Let DblFeed = Application.WorksheetFunction.Sum(VecFeed)
      Let DblFeedVol = 0
    Case eqlBalBoth '2
      Let DblFeed = VBA.Abs(VBA.CDbl(VecFeed(0)))
      Let DblFeedVol = DblFeed
      If i > 1 Then Let DblFeedVol = VBA.Abs(VBA.Abs(Application.WorksheetFunction.Sum(VecFeed)) - DblFeed)
  End Select
  Let VecFeed = Empty
  Set RES = Nothing

  'Return
  Let BVarRequest = Empty
  Let BVarRequest = Array(DblFeed, DblFeedVol, VecGrad, VecGradCx, VecName, VecUnit, VecNameCx, VecUnitCx, VecGradGL)
  Let Balmet_core_main_udf_prepare = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Let VecFeed = Empty
  Let VecGrad = Empty
  Let VecName = Empty
  Let VecUnit = Empty
  Let VecGradCx = Empty
  Let VecNameCx = Empty
  Let VecUnitCx = Empty
  Let VecGradGL = Empty
  Set RCT = Nothing
  Set RES = Nothing

End Function

Private Function Balmet_core_main_udf_calculate( _
  ByVal BVarRequest As Variant, _
  ByVal BMethod As Byte, _
  ByVal BResult As Byte, _
  ByVal BDirection As EQLBAL_ENU_RESULTDIRECTION, _
  ByRef BVarErrUDF As Variant, _
  ByVal BBooComplex As Boolean, _
  ByVal BBooGL As Boolean, _
  ByVal BAllResult As Boolean, _
  ByVal BEnuTypeFeed As EQLBAL_ENU_TYPECONTENT) As Boolean

  Dim RES As AppResources_current

  Dim VecWeight As Variant, VecVolume As Variant
  Dim DblFeed As Double, DblFeedVol As Double
  
  On Error GoTo EH
  Let Balmet_core_main_udf_calculate = False

  Set RES = New AppResources_current
  Set RES.ErrorHandler = EHGLOBAL

  'Data entry: (DblFeed, DblFeedVol, VecGrad, VecGradCx, VecName, VecUnit, VecNameCx, VecUnitCx, VecGradGL)
  Let DblFeed = VBA.CDbl(BVarRequest(0))
  Let DblFeedVol = VBA.CDbl(BVarRequest(1))
  If Not Balmet_core_main_app_spreadbalmet_get_weights(DblFeed, BVarRequest(2), BMethod, VecWeight, RES, BFullData:=True) Then GoTo EH
  If BEnuTypeFeed = eqlBalBoth Then
    If Not Balmet_core_main_app_spreadbalmet_get_volumes(DblFeedVol, VecWeight, BVarRequest(2), BVarRequest(3), BVarRequest(5), VecVolume, RES) Then GoTo EH
  End If

  'Return
  Select Case BResult
    
    Case eqlBalWeights
      If Not BAllResult Then
        Call RES.ArrayPop(VecWeight)
        If Not VBA.IsArray(VecWeight) Then Call EHGLOBAL.ErrorHandlerSet(0.1, "¡Error al calcular los pesos del Balance Metalúrgico!"): GoTo EH 'Error
      End If
      Let GLOVEC_BALANCE = VecWeight 'eqlBalWeights [mx1]
    
    Case eqlBalWeightPercents: If Not Balmet_core_main_app_spreadbalmet_get_percents(VecWeight, GLOVEC_BALANCE, RES) Then GoTo EH 'eqlBalWeightPercents [mx1]
    
    Case eqlBalVolume 'eqlBalVolume [mx1]
      If BEnuTypeFeed = eqlBalBoth Then
        Let GLOVEC_BALANCE = VecVolume
      Else
        Let GLOVEC_BALANCE = VecWeight
      End If
    
    Case eqlBalVolumePercents 'eqlBalVolumePercents [mx1]
      If BEnuTypeFeed = eqlBalBoth Then
        If Not Balmet_core_main_app_spreadbalmet_get_volumes_percents(VecVolume, GLOVEC_BALANCE, RES) Then GoTo EH
      Else
        If Not Balmet_core_main_app_spreadbalmet_get_percents(VecWeight, GLOVEC_BALANCE, RES, BEnuTypeFeedPer:=BEnuTypeFeed) Then GoTo EH
      End If
    
    Case eqlBalGradesHeads
      If Not Balmet_core_main_app_spreadbalmet_get_grades(VecWeight, VecVolume, BVarRequest(2), _
         BVarRequest(3), BVarRequest(5), BVarRequest(4), GLOVEC_BALANCE, BAllResult, RES, _
         BBooHeads:=True, BBooJustHeads:=True) Then GoTo EH 'eqlBalGradesHeads [2xn]
    
    Case eqlBalUnities 'eqlBalUnities [mxn]
      If Not Balmet_core_main_app_spreadbalmet_get_unities(VecWeight, VecVolume, BVarRequest(2), _
         BVarRequest(3), BVarRequest(5), BVarRequest(4), GLOVEC_BALANCE, BAllResult, RES) Then GoTo EH 'eqlBalGradesHeadsCx [2xn]
    
    Case eqlBalFines 'eqlBalFines [mxn]
      If Not Balmet_core_main_app_spreadbalmet_get_fines(VecWeight, VecVolume, BVarRequest(2), _
         BVarRequest(3), BVarRequest(5), BVarRequest(4), GLOVEC_BALANCE, BAllResult, RES) Then GoTo EH 'eqlBalGradesHeadsCx [2xn]
    
    Case eqlBalRecoveries 'eqlBalRecoveries [mxn]
      If Not Balmet_core_main_app_spreadbalmet_get_recoveries(VecWeight, VecVolume, BVarRequest(2), _
         BVarRequest(3), BVarRequest(5), BVarRequest(4), GLOVEC_BALANCE, BAllResult, RES) Then GoTo EH 'eqlBalGradesHeadsCx [2xn]
    
    Case eqlBalRatio 'eqlBalRatio [mx1]
      If Not Balmet_core_main_app_spreadbalmet_get_ratio(VecWeight, VecVolume, GLOVEC_BALANCE, RES) Then GoTo EH
    
    Case eqlBalGradesHeadsCx
      If Not Balmet_core_main_app_spreadbalmet_get_grades(VecWeight, VecVolume, BVarRequest(2), _
         BVarRequest(3), BVarRequest(5), BVarRequest(4), GLOVEC_BALANCE, BAllResult, RES, _
         BBooHeads:=True, BBooJustHeads:=True, BBooCx:=True) Then GoTo EH 'eqlBalGradesHeadsCx [2xn]
    
    Case eqlBalUnitiesCx 'eqlBalUnitiesCx [mxn]
      If Not Balmet_core_main_app_spreadbalmet_get_unities(VecWeight, VecVolume, BVarRequest(2), _
         BVarRequest(3), BVarRequest(5), BVarRequest(4), GLOVEC_BALANCE, BAllResult, RES, _
         BBooCx:=True) Then GoTo EH 'eqlBalGradesHeadsCx [2xn]
    
    Case eqlBalFinesCx 'eqlBalFinesCx [mxn]
      If Not Balmet_core_main_app_spreadbalmet_get_fines(VecWeight, VecVolume, BVarRequest(2), _
         BVarRequest(3), BVarRequest(5), BVarRequest(4), GLOVEC_BALANCE, BAllResult, RES, _
         BBooCx:=True) Then GoTo EH 'eqlBalGradesHeadsCx [2xn]
    
    Case eqlBalRecoveriesCx 'eqlBalRecoveriesCx [mxn]
      If Not Balmet_core_main_app_spreadbalmet_get_recoveries(VecWeight, VecVolume, BVarRequest(2), _
         BVarRequest(3), BVarRequest(5), BVarRequest(4), GLOVEC_BALANCE, BAllResult, RES, _
         BBooCx:=True) Then GoTo EH 'eqlBalGradesHeadsCx [2xn]
    
    Case Else: GoTo EH
  
  End Select
  Let VecWeight = Empty

  If BDirection = eqlBalHorizontal Then
    Let GLOVEC_BALANCE = RES.ArrayTranspose(GLOVEC_BALANCE)
    If Not VBA.IsArray(GLOVEC_BALANCE) Then GoTo EH
  End If

  Let Balmet_core_main_udf_calculate = True

EH:
  If Not Balmet_core_main_udf_calculate Then
    Select Case EHGLOBAL.ErrorNum
      Case 0.1, 1.11, 1.2: Let BVarErrUDF = VBA.CVErr(xlErrNA)
      Case 1.111, 1.114: Let BVarErrUDF = VBA.CVErr(xlErrDiv0)
      Case 1.1: Let BVarErrUDF = VBA.CVErr(xlErrNum)
      Case Else: Let BVarErrUDF = VBA.CVErr(xlErrValue)
      ' 0.2,1.113,1.3,1.4,2.1,2.11,2.2,3.1,3.2,4.1,4.2,5.1,5.2,5.3,5.4,6.1,6.2,6.3,7.1,7.2,7.3,7.4,8.1,8.2,9.1,9.2,9.3
    End Select
  End If
  Set RES = Nothing

End Function



' APP CORE METHODS
Private Function Balmet_core_main_app(ByVal BVarRequest As Variant) As Boolean
  
  ' Data entry in BVarRequest: (0.DblFeed, 1.DblFeedVol, 2.VecGrad, 3.VecGradCx, 4.VecName, 5.VecUnit, 6.BytMethod, 7.Booleans, 8.EnuTypeFeed, 9.RngBox)
  Dim RES As AppResources_current
  
  Dim RngBox As Range
  Dim VecWeight As Variant, VecVolume As Variant, VecAux0 As Variant, VecAux1 As Variant
  Dim BooAll As Boolean, BooPercents As Boolean, BooUnities As Boolean, BooFines As Boolean, _
      BooGrams As Boolean, BooOT As Boolean, BooRatio As Boolean, BooHeads As Boolean
  Dim StrVol As String
  Dim LngAux0 As Long, LngPer As Long, LngGrad As Long, LngGrades As Long, LngCurrent As Long
  Dim DblFeed As Double, DblFeedVol As Double
  Dim EnuTypeFeed As EQLBAL_ENU_TYPECONTENT
  Dim EnuMethod As EQLBAL_ENU_METHOD
  
  Const sFN As String = "APPBALMET::-core_main_app_calculate"
  
  On Error GoTo EH
  Let Balmet_core_main_app = False
  
  Set RES = New AppResources_current
  Set RES.ErrorHandler = EHGLOBAL
  
  ' Unpack data
  Let DblFeed = VBA.CDbl(BVarRequest(0))
  Let DblFeedVol = VBA.CDbl(BVarRequest(1))
  Let EnuMethod = IIf(VBA.CByte(BVarRequest(6)) = 0, eqlBalConventional, IIf(VBA.CByte(BVarRequest(6)) = 1, eqlBalCramer, eqlBalInverseMatrix))
  Let VecAux0 = BVarRequest(7) ' Booleans Array
  Let BooAll = VBA.CBool(VecAux0(0))
  Let BooPercents = VBA.CBool(VecAux0(1))
  Let BooUnities = VBA.CBool(VecAux0(2))
  Let BooFines = VBA.CBool(VecAux0(3))
  Let BooGrams = VBA.CBool(VecAux0(4))
  Let BooOT = VBA.CBool(VecAux0(5))
  Let BooRatio = VBA.CBool(VecAux0(6))
  Let BooHeads = VBA.CBool(VecAux0(7))
  Let EnuTypeFeed = IIf(VBA.CByte(BVarRequest(8)) = 0, eqlBalSolids, IIf(VBA.CByte(BVarRequest(8)) = 1, eqlBalVolumes, eqlBalBoth))
  If GLOBOO_FORMULAS Then Set RngBox = BVarRequest(9) ' Get Rng if formulas
  
  ' Get Balmet products in main array
  Let LngAux0 = (UBound(BVarRequest(2), 2) - LBound(BVarRequest(2), 2))
  If Not Balmet_core_main_app_spreadbalmet_get_products_names(BVarRequest(4), LngAux0, VecAux0, RES, BooHeads) Then GoTo EH
  
  ' Get Balmet Contents wit its optional Percent column in two arrays and concat to main array
  If Not Balmet_core_main_app_spreadbalmet_get_weights(DblFeed, BVarRequest(2), EnuMethod, VecWeight, RES, BFullData:=True) Then GoTo EH
  If Not Balmet_core_main_app_spreadbalmet_get_weights(DblFeed, BVarRequest(2), EnuMethod, VecAux1, RES, GLOBOO_FORMULAS, BooHeads, True, RngBox, EnuTypeFeed, True) Then GoTo EH
  Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1, RByCols:=True)
  If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(1, "Fallo al unir resultados de Contenidos a la matriz de Balance Metalúrgico."): GoTo EH
  If BooPercents Then
    If Not Balmet_core_main_app_spreadbalmet_get_percents(VecWeight, VecAux1, RES, GLOBOO_FORMULAS, BooHeads, True, EnuTypeFeed, RngBox) Then GoTo EH
    Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1, RByCols:=True)
    If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(3, "Fallo al unir resultados de Porcentaje en Contenidos a la matriz de Balance Metalúrgico."): GoTo EH
  End If
  
  Let StrVol = IIf(EnuTypeFeed = eqlBalBoth, IIf(BooPercents, "E3", "D3"), "C3")
  If EnuTypeFeed = eqlBalBoth Then
    If Not Balmet_core_main_app_spreadbalmet_get_volumes(DblFeedVol, VecWeight, BVarRequest(2), BVarRequest(3), BVarRequest(5), VecVolume, RES) Then GoTo EH
    If Not Balmet_core_main_app_spreadbalmet_get_volumes(DblFeedVol, VecWeight, BVarRequest(2), BVarRequest(3), BVarRequest(5), VecAux1, RES, GLOBOO_FORMULAS, True, BooHeads, RngBox, BooPercents) Then GoTo EH
    Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1, RByCols:=True)
    If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(2, "Fallo al unir resultados de Contenidos en volumen a la matriz de Balance Metalúrgico."): GoTo EH
    If BooPercents Then
      If Not Balmet_core_main_app_spreadbalmet_get_volumes_percents(VecVolume, VecAux1, RES, GLOBOO_FORMULAS, BooHeads, True, RngBox) Then GoTo EH
      Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1, RByCols:=True)
      If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(4, "Fallo al unir resultados de Porcentaje en Volumen a la matriz de Balance Metalúrgico."): GoTo EH
    End If
  End If
  
  ' Get the rest of products and concat to main array
  Let LngGrades = IIf(EnuTypeFeed = eqlBalBoth, IIf(BooPercents, 4, 2), IIf(BooPercents, 2, 1))
  Let LngCurrent = (UBound(VecAux0, 2) - LBound(VecAux0, 2)) - 1
  If Not Balmet_core_main_app_spreadbalmet_get_grades(VecWeight, VecVolume, BVarRequest(2), BVarRequest(3), BVarRequest(5), BVarRequest(4), VecAux1, True, RES, GLOBOO_FORMULAS, True, BooHeads, RngBox, False, False, BooPercents, EnuTypeFeed, StrVol, LngGrades, LngCurrent) Then GoTo EH 'eqlBalGradesHeads [2xn]
  Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1, RByCols:=True)
  If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(5, "Fallo al unir resultados de Leyes a la matriz de Balance Metalúrgico."): GoTo EH
  Let LngCurrent = (UBound(VecAux0, 2) - LBound(VecAux0, 2)) - 1
  
  If BooUnities Then
    If Not Balmet_core_main_app_spreadbalmet_get_unities(VecWeight, VecVolume, BVarRequest(2), BVarRequest(3), BVarRequest(5), BVarRequest(4), VecAux1, True, RES, GLOBOO_FORMULAS, True, BooHeads, RngBox, False, BooPercents, EnuTypeFeed, StrVol, LngGrades, LngCurrent) Then GoTo EH 'eqlBalGradesHeadsCx [2xn]
    Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1, RByCols:=True)
    If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(6, "Fallo al unir resultados de Unidades a la matriz de Balance Metalúrgico."): GoTo EH
    Let LngCurrent = (UBound(VecAux0, 2) - LBound(VecAux0, 2)) - 1
  End If
  
  If BooFines Then
    If Not Balmet_core_main_app_spreadbalmet_get_fines(VecWeight, VecVolume, BVarRequest(2), BVarRequest(3), BVarRequest(5), BVarRequest(4), VecAux1, True, RES, GLOBOO_FORMULAS, True, BooHeads, RngBox, False, BooGrams, BooOT, BooPercents, EnuTypeFeed, StrVol, LngGrades, LngCurrent) Then GoTo EH 'eqlBalGradesHeadsCx [2xn]
    Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1, RByCols:=True)
    If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(7, "Fallo al unir resultados de Finos a la matriz de Balance Metalúrgico."): GoTo EH
    Let LngCurrent = (UBound(VecAux0, 2) - LBound(VecAux0, 2)) - 1
  End If
  
  If Not Balmet_core_main_app_spreadbalmet_get_recoveries(VecWeight, VecVolume, BVarRequest(2), BVarRequest(3), BVarRequest(5), BVarRequest(4), VecAux1, True, RES, GLOBOO_FORMULAS, True, BooHeads, RngBox, False, BooPercents, EnuTypeFeed, StrVol, LngGrades, LngCurrent) Then GoTo EH 'eqlBalGradesHeadsCx [2xn]
  Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1, RByCols:=True)
  If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(8, "Fallo al unir resultados de Recuperaciones a la matriz de Balance Metalúrgico."): GoTo EH
  Let LngCurrent = (UBound(VecAux0, 2) - LBound(VecAux0, 2)) - 1
  
  If BooRatio Then
    If Not Balmet_core_main_app_spreadbalmet_get_ratio(VecWeight, VecVolume, VecAux1, RES, GLOBOO_FORMULAS, True, BooHeads, RngBox, StrVol) Then GoTo EH
    Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1, RByCols:=True)
    If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(9, "Fallo al unir resultados de Ratio a la matriz de Balance Metalúrgico."): GoTo EH
  End If
  Let VecWeight = Empty
  Let VecVolume = Empty

  ' Return
  Let GLOVEC_BALANCE = VecAux0
  Let VecAux0 = Empty

  Let Balmet_core_main_app = True

EH:
  Let VecWeight = Empty
  Let VecVolume = Empty
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function



' APP CORE METHODS - ECO
Private Function Balmet_core_main_udf_eco( _
  ByVal BVarFeed As Variant, _
  ByVal BVarConc As Variant, _
  ByVal BVarGrad As Variant, _
  ByVal BVarReco As Variant, _
  ByVal BVarGradCx As Variant, _
  ByVal BVarRecoCx As Variant, _
  ByVal BVarUnit As Variant, _
  ByVal BResult As EQLBAL_ENU_RESULTECO, _
  ByVal BTypeResult As EQLBAL_ENU_TYPERESULT, _
  ByVal BDirection As EQLBAL_ENU_RESULTDIRECTION) As Variant
  
  'Obtains the Feed Grade from one or mxn matrix of concentrates (for more or equal than two concentrates, Feed Grade will be an accourate) _
  Requirements: _
  1. BVarFeed:   Feed, the weight feed of the process in Tons _
  2. BVarRecs:   Concentrates Recuperations, a number or mxn matrix of Doubles _
  3. BVarWeig:Concentrates Weights, a number or mxn matrix of Doubles in Tons _
  4. BVarGrad: Balmet Concentrates Grades, the mxn matrix of concentrates, %, DM, g/L or g/T _
  The Method returns a Double or a matrix of Doubles with the value or values of the Feed(s) Grade(s).

  Dim RES As AppResources_current
  
  Dim VecAux0 As Variant, VecAux1 As Variant
  Dim BooComplex As Boolean, BooError As Boolean
  Dim EnuTypeFeed As EQLBAL_ENU_TYPECONTENT
  
  On Error GoTo EH
  Let Balmet_core_main_udf_eco = VBA.CVErr(xlErrValue)
  Let BooError = False
  Set RES = New AppResources_current
  Set RES.ErrorHandler = EHGLOBAL
  
  'Forming the request data
  Let VecAux0 = Array(BVarFeed, BVarConc, BVarGrad, BVarReco, BVarGradCx, BVarRecoCx, BVarUnit)
  
  'Get final result
  If Not Balmet_core_main_udf_validate_eco(VecAux0, Balmet_core_main_udf_eco, RES) Then GoTo EH
  If Not Balmet_core_main_udf_prepare_eco(VecAux0, VecAux1, Balmet_core_main_udf_eco, RES, BooComplex) Then GoTo EH
  If Not Balmet_core_main_udf_calculate_eco(VecAux1, VecAux0, BooComplex, RES, BResult, BDirection) Then GoTo EH
  
  'Return
  If BTypeResult = eqlBalJustGrades Then
    Let Balmet_core_main_udf_eco = VecAux0(0)
  ElseIf BTypeResult = eqlBalJustGradesCx Then
    Let Balmet_core_main_udf_eco = VecAux0(1)
  Else
    Let Balmet_core_main_udf_eco = RES.ArrayConcat(VecAux0(0), VecAux0(1), RByCols:=(BDirection = eqlBalVertical))
    If Not VBA.IsArray(Balmet_core_main_udf_eco) Then GoTo EH
  End If

  Let BooError = True
  
EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Set RES = Nothing
  If Not BooError Then
    Select Case EHGLOBAL.ErrorNum
      Case 0.1: Let Balmet_core_main_udf_eco = VBA.CVErr(xlErrDiv0)
      Case 0.2: Let Balmet_core_main_udf_eco = VBA.CVErr(xlErrNum)
    End Select
  End If

End Function

Private Function Balmet_core_main_udf_validate_eco( _
  ByVal BVarRequest As Variant, _
  ByRef BVarErrUDF As Variant, _
  ByVal RES As AppResources_current) As Boolean

  'Data required: (BVarFeed, BVarConc, BVarGrad, BVarReco, BVarGradCx, BVarRecoCx, BVarUnit)
  Dim MDL As ModelExcliqliteDatasheet

  'Dim RngBox As Range
  Dim VecFeed As Variant, VecWeights As Variant, VecGrades As Variant, VecRecs As Variant, VecGradesCx As Variant, VecRecsCx As Variant, VecUnits As Variant, VecUnitsCx As Variant
  Dim BooBalCx As Boolean, BooGL As Boolean
  Dim StrAux0 As String
  
  Dim DblFeedSolid As Double, DblFeedVolume As Double, DblSum As Double, DblAux0 As Double, DblAux1 As Double
  Dim i As Long, j As Long, x1 As Long, x2 As Long, x3 As Long, x4 As Long, x5 As Long
  
  Const sFN As String = "APPBALMET::-core_main_udf_validation_eco"
  
  On Error GoTo EH
  Let Balmet_core_main_udf_validate_eco = False
  
  'All incoming main values are strings, requires RES method getVariantsAsStrings
  Let VecFeed = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(0), x1, RValueType:=eqlResNumbers, RBooAbs:=True)
  If x1 = 0 Then Let BVarErrUDF = VBA.CVErr(xlErrNull): GoTo EH 'Feeds required
  Let VecWeights = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(1), x2, RValueType:=eqlResNumbers, RBooAbs:=True)
  If x2 = 0 Then Let BVarErrUDF = VBA.CVErr(xlErrNull): GoTo EH 'Weights required
  If x2 > GLOBOO_BAL_LIMIT Then Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH 'Just 5 products as limit of balmets and One product at least
  Let VecGrades = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(2), x3, RValueType:=eqlResNumbers, RBooAbs:=True)
  If x3 = 0 Or Not x3 = x2 Then Let BVarErrUDF = VBA.CVErr(xlErrNull): GoTo EH 'Grades required
  Let VecRecs = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(3), x4, RValueType:=eqlResNumbers, RBooAbs:=True)
  Let VecGradesCx = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(4), x5, RValueType:=eqlResNumbers, RBooAbs:=True)
  If x5 > GLOBOO_BAL_LIMITCX Then Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH 'Just 3 Cx avaiable
  Let BooBalCx = (i > 0)
  Let VecRecsCx = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(5), i, RValueType:=eqlResNumbers, RBooAbs:=True)
  Let VecUnits = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(6), i, RValueType:=eqlResStrings, RStrCustomRegex:="%", RValueAs:=eqlResStrings)
  
  'Units
  If i > 0 Then
    'Set RngBox = excliqlitedatasheet.ListObjects("excliqlitechemicalunits").DataBodyRange
'    Set MDL = New ModelApp
'    Set MDL.ErrorHandler = EHGLOBAL
'    Set RngBox = MDL.MGet(eqlMdlTblChemicalunits, eqlMdlRange)
'    Set MDL = Nothing
'    For j = 0 To UBound(VecUnits)
'      If Application.WorksheetFunction.CountIf(RngBox, VBA.CStr(VecUnits(j))) = 0 Then Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH 'Unit doesn't exists
'    Next j
'    Set RngBox = Nothing
    
    Set MDL = New ModelExcliqliteDatasheet
    Set MDL.ErrorHandler = EHGLOBAL
    For i = 0 To UBound(VecUnits)
      If Not VBA.IsArray(MDL.MGet(eqlMdlTblChemicalunits, eqlMdlArray, "Unidad", MStrWhere:="Unidad=" & VBA.CStr(VecUnits(i)))) Then Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH 'Unit doesn't exists
    Next i
    Set MDL = Nothing
    
    If RES.ArrayHasValueN(VecUnits, "dm", False, True) > 1 Then GoTo EH 'just one DM's is required
    Let BooGL = RES.ArrayHasValue(VecUnits, "g/l", False, True)
  End If
  If BooBalCx Then
    If i < x3 Then 'Refill units prev check of cx
      For j = i To (x3 - 1)
        ReDim Preserve VecUnits(j)
        Let VecUnits(j) = "%"
      Next j
      Let i = j
    End If
    If i < (x3 + x5) Then 'Refill cx grades
      For j = i To ((x3 + x5) - 1)
        ReDim Preserve VecUnits(j)
        Let VecUnits(j) = "g/T"
      Next j
      Let i = j
    End If
    If i > (x3 + x5) Then 'Slice units
      Call RES.ArraySlice(VecUnits, (x3 + x5))
      If Not VBA.IsArray(VecUnits) Then GoTo EH 'Error in slice method
    End If
  End If
  
  'Feed
  Let DblFeedSolid = VBA.CDbl(VecFeed(0))
  Let DblSum = Application.WorksheetFunction.Sum(VecFeed)
  Let DblFeedVolume = IIf(BooGL, IIf((DblSum - DblFeedSolid) <= 0, DblFeedSolid, (DblSum - DblFeedSolid)), 0)
  If DblFeedSolid <= 0 Then Let BVarErrUDF = VBA.CVErr(xlErrNull): GoTo EH 'Feed required
  If DblFeedSolid > CDBL_MAX Then Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH 'Feed too big
  If BooGL Then
    If DblFeedVolume <= 0 Then Let BVarErrUDF = VBA.CVErr(xlErrNull): GoTo EH 'Volume Feed required
    If DblFeedVolume > CDBL_MAX Then Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH 'Volume Feed too big
  End If
  Let VecFeed = Empty
  
  'Weight(s)
  Let DblAux0 = 0
  Let DblAux1 = 0
  For i = 0 To UBound(VecWeights)
    If VBA.LCase(VBA.CStr(VecUnits(i))) = "g/l" Then
      Let DblAux1 = DblAux1 + VecWeights(i)
    Else
      Let DblAux0 = DblAux0 + VecWeights(i)
    End If
  Next i
  If DblAux0 > DblFeedSolid Then GoTo EH 'Weights can't be greater than Feed
  If BooGL And DblAux1 > 0 And DblAux1 > DblFeedVolume Then GoTo EH 'Volumes can't be greater than Volume Feed
  Let VecWeights = Empty
  
  'Grade(s)
  For i = 0 To UBound(VecGrades) 'Validate values for each Grade
    Let DblAux0 = VBA.Abs(VBA.CDbl(VecGrades(i)))
    Let StrAux0 = VBA.LCase(VBA.CStr(VecUnits(i)))
    Select Case StrAux0
      Case "%": If DblAux0 > CDBL_PC_TOP Then Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH 'Same Grades, same weights'No one Grade can exceed 100%
      Case "dm": If DblAux0 > CDBL_DM_TOP Then Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH 'Same Grades, same weights 'No one DM Grade can exceed 10000DM
      Case "g/t", "g/l": If DblAux0 > CDBL_GT_TOP Then Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH 'Same Grades, same weights 'No one DM Grade can exceed 1000000[g/T|g/L]
      Case Else: GoTo EH
    End Select
  Next i
  Let VecGrades = Empty
  
  'Recuperation(s)
  For i = 0 To UBound(VecRecs)
    If VecRecs(i) > CDBL_PC_TOP Then GoTo EH 'Recuperations are percentages, so its limit is 100%
  Next i
  Let VecRecs = Empty
    
  'GradesCx Units and GradesCX
  If BooBalCx Then
    'Complex units
    Let VecUnitsCx = RES.ArraySlice(VecUnits, x3)
    If Not VBA.IsArray(VecUnits) Then GoTo EH
    If RES.ArrayHasValue(VecUnitsCx, "%", False, True) Then Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH 'Unit unavaiable for cx
    'Complex Grade
    For i = 0 To UBound(VecGradesCx) 'Validate values for each Grade
      Let DblAux0 = VBA.Abs(VBA.CDbl(VecGradesCx(i)))
      Let StrAux0 = VBA.LCase(VBA.CStr(VecUnitsCx(i)))
      Select Case StrAux0
        Case "dm": If DblAux0 > CDBL_DM_TOP Then Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH 'No one DM Grade can exceed 10000DM
        Case "g/t", "g/l": If DblAux0 > CDBL_GT_TOP Then Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH 'No one DM Grade can exceed 1000000[g/T|g/L]
        Case Else: GoTo EH
      End Select
    Next i
    Let VecGradesCx = Empty
    Let VecUnitsCx = Empty
    'Complex Recs
    For i = 0 To UBound(VecRecsCx)
      If VecRecsCx(i) > CDBL_PC_TOP Then GoTo EH 'Recuperations are percentages, so its limit is 100%
    Next i
    Let VecRecsCx = Empty
  End If
  Let VecUnits = Empty
  
  Let Balmet_core_main_udf_validate_eco = True

EH:
  Let VecFeed = Empty
  Let VecWeights = Empty
  Let VecGrades = Empty
  Let VecRecs = Empty
  Let VecUnits = Empty
  Let VecGradesCx = Empty
  Let VecUnitsCx = Empty
  Let VecRecsCx = Empty
  'Set RngBox = Nothing
  Set MDL = Nothing

End Function

Private Function Balmet_core_main_udf_prepare_eco( _
  ByVal BVarRequest As Variant, _
  ByRef BVarResponse As Variant, _
  ByRef BVarErrUDF As Variant, _
  ByVal RES As AppResources_current, _
  ByRef BBooCx As Boolean) As Boolean

  'Data required: (BVarFeed, BVarConc, BVarGrad, BVarReco, BVarGradCx, BVarRecoCx, BVarUnit)
  Dim VecFeed As Variant, VecWeights As Variant, VecGrades As Variant, VecRecs As Variant, VecGradesCx As Variant, VecRecsCx As Variant, VecUnits As Variant, VecUnitsCx As Variant
  Dim BooGL As Boolean
  Dim StrAux0 As String
  
  Dim DblFeedSolid As Double, DblFeedVolume As Double, DblSum As Double, DblAux0 As Double, DblAux1 As Double
  Dim i As Long, j As Long, x1 As Long, x2 As Long, x3 As Long, x4 As Long, x5 As Long
  
  Const sFN As String = "APPBALMET::-core_main_udf_validation_eco"
  
  On Error GoTo EH
  Let Balmet_core_main_udf_prepare_eco = False
  
  'All incoming main values are variants, requires RES method getVariantsAsVector
  Let VecFeed = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(0), x1, RValueType:=eqlResNumbers, RBooAbs:=True)
  Let VecWeights = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(1), x2, RValueType:=eqlResNumbers, RBooAbs:=True)
  Let VecGrades = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(2), x3, RValueType:=eqlResNumbers, RBooAbs:=True)
  Let VecRecs = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(3), x4, RValueType:=eqlResNumbers, RBooAbs:=True)
  Let VecGradesCx = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(4), x5, RValueType:=eqlResNumbers, RBooAbs:=True)
  Let BBooCx = (i > 0)
  Let VecRecsCx = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(5), i, RValueType:=eqlResNumbers, RBooAbs:=True)
  Let VecUnits = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(6), i, RValueType:=eqlResStrings, RStrCustomRegex:="%", RValueAs:=eqlResStrings)
  
  'Units
  If i > 0 Then Let BooGL = RES.ArrayHasValue(VecUnits, "g/l", False, True)
  If i < x3 Then 'Refill units prev check of cx
    For j = i To (x3 - 1)
      ReDim Preserve VecUnits(j)
      Let VecUnits(j) = "%"
    Next j
    Let i = j
  End If
  If BBooCx Then
    If i < (x3 + x5) Then 'Refill cx grades
      For j = i To ((x3 + x5) - 1)
        ReDim Preserve VecUnits(j)
        Let VecUnits(j) = "g/T"
      Next j
      Let i = j
    End If
    If i > (x3 + x5) Then 'Slice units
      Call RES.ArraySlice(VecUnits, (x3 + x5))
      If Not VBA.IsArray(VecUnits) Then GoTo EH 'Error in slice method
    End If
  End If
  
  'Feed
  Let DblFeedSolid = VBA.CDbl(VecFeed(0))
  Let DblSum = Application.WorksheetFunction.Sum(VecFeed)
  Let DblFeedVolume = IIf(BooGL, IIf((DblSum - DblFeedSolid) <= 0, DblFeedSolid, (DblSum - DblFeedSolid)), 0)
  Let VecFeed = Empty
  
  'Weight(s)
  
  'Grade(s)
  
  'Recuperation(s)
  If x4 < x3 Then
    For i = x3 To (x3 - 1)
      ReDim Preserve VecRecs(i)
      Let VecRecs(i) = 80
    Next i
  End If
    
  'GradesCx Units and GradesCX
  If BBooCx Then
    'Complex units
    Let VecUnitsCx = RES.ArraySlice(VecUnits, x3)
    If Not VBA.IsArray(VecUnitsCx) Then GoTo EH
    'Complex Grade
    'Complex Recs
    If UBound(VecRecsCx) < UBound(VecGradesCx) Then
      Let j = UBound(VecRecsCx)
      For i = j To UBound(VecGradesCx)
        ReDim Preserve VecRecsCx(i)
        Let VecRecsCx(i) = 80
      Next i
    End If
    If UBound(VecRecsCx) > UBound(VecGradesCx) Then
      Call RES.ArraySlice(VecRecsCx, UBound(VecGradesCx))
      If Not VBA.IsArray(VecRecsCx) Then GoTo EH
    End If
  End If
  
  Let BVarResponse = Array(Array(DblFeedSolid, DblFeedVolume), VecWeights, VecGrades, VecRecs, VecGradesCx, VecRecsCx, VecUnits, VecUnitsCx)
  Let Balmet_core_main_udf_prepare_eco = True

EH:
  Let VecFeed = Empty
  Let VecWeights = Empty
  Let VecGrades = Empty
  Let VecRecs = Empty
  Let VecUnits = Empty
  Let VecGradesCx = Empty
  Let VecUnitsCx = Empty
  Let VecRecsCx = Empty
  
End Function

Private Function Balmet_core_main_udf_calculate_eco( _
  ByVal BVarRequest As Variant, _
  ByRef BVarResponse As Variant, _
  ByVal BBooCx As Boolean, _
  ByVal RES As AppResources_current, _
  ByVal BEnuTypeBalEco As EQLBAL_ENU_RESULTECO, _
  ByVal BEnuOrientation As EQLBAL_ENU_RESULTDIRECTION) As Boolean
  
  'Data required: (DblFeedSolid, DblFeedVolume), VecWeights, VecGrades, VecRecs, VecGradesCx, VecRecsCx, VecUnits, VecUnitsCx)
  Dim DblFeed As Double, DblFeedSol As Double, DblFeedVol As Double
  Dim VecWeig As Variant, VecGrad As Variant, VecRecs As Variant, VecGradCx As Variant, VecRecsCx As Variant, VecUnit As Variant, VecUnitCx As Variant
  Dim VecResu As Variant, VecResuCx As Variant, VecAux0 As Variant, VecAux1 As Variant
  Dim StrAux0 As String
  Dim i As Long, j As Long
  Dim DblAux1 As Double, DblAux2 As Double ', DblAux3 As Double
  
  Const sFN As String = "APPBALMET::-core_main_udf_calculate_eco"
  
  On Error GoTo EH
  Let Balmet_core_main_udf_calculate_eco = False
  
  'Unpacking all datas
  Let VecAux0 = BVarRequest(0)
  Let DblFeedSol = VBA.CDbl(VecAux0(0))
  Let DblFeedVol = VBA.CDbl(VecAux0(1))
  Let VecAux0 = Empty
  Let VecWeig = BVarRequest(1)
  Let VecGrad = BVarRequest(2)
  Let VecRecs = BVarRequest(3)
  Let VecGradCx = BVarRequest(4)
  Let VecRecsCx = BVarRequest(5)
  Let VecUnit = BVarRequest(6)
  Let VecUnitCx = BVarRequest(7)
  
  'RESOLVE BALMET
  'Resolve the Balmet base
  ReDim VecAux0(UBound(VecWeig))
  For i = 0 To UBound(VecWeig)
    Let StrAux0 = VBA.LCase(VBA.CStr(VecUnit(i)))
    Let DblFeed = IIf(StrAux0 = "g/l", DblFeedVol, DblFeedSol)
    If BEnuTypeBalEco = eqlBalHeadsGrades Then '= 100Kk / A%R
      Let DblAux1 = VBA.CDbl(VecWeig(i)) * VBA.CDbl(VecGrad(i)) * 100
      Let DblAux2 = DblFeed * VBA.CDbl(VecRecs(i))
    Else '= A a %R / 100K
      Let DblAux1 = VBA.CDbl(VecRecs(i)) * DblFeed * VBA.CDbl(VecGrad(i))
      Let DblAux2 = VBA.CDbl(VecWeig(i)) * 100
    End If
    If DblAux2 = 0 Then Call EHGLOBAL.ErrorHandlerSet(0.1, "División entre cero. Alimentación o Recuperaciones erróneas."): GoTo EH 'Error Div0!
    Let VecAux0(i) = DblAux1 / DblAux2
  Next i
  
  'Build the final data
  ReDim VecResu(0 To 0, 0 To UBound(VecAux0))
  For j = 0 To UBound(VecAux0) 'Just one row (horizontal)
    Let VecResu(0, j) = VecAux0(j)
  Next j
  Let VecAux0 = Empty
  
  'Resolve Complex product if it exists
  If BBooCx Then
    ReDim VecAux0(UBound(VecGradCx))
    For i = LBound(VecGradCx) To UBound(VecGradCx)
      Let StrAux0 = VBA.LCase(VBA.CStr(VecUnitCx(i)))
      Let DblFeed = IIf(StrAux0 = "g/l", DblFeedVol, DblFeedSol)
      If BEnuTypeBalEco = eqlBalHeadsGrades Then '= 100Kk / A%R
        Let DblAux1 = VBA.CDbl(VecWeig(i)) * VBA.CDbl(VecGradCx(i)) * 100
        Let DblAux2 = DblFeed * VBA.CDbl(VecRecs(i))
      Else '= A a %R / 100K
        Let DblAux1 = VBA.CDbl(VecRecs(i)) * DblFeed * VBA.CDbl(VecGradCx(i))
        Let DblAux2 = VBA.CDbl(VecWeig(i)) * 100
      End If
      If DblAux2 = 0 Then Call EHGLOBAL.ErrorHandlerSet(0.1, "División entre cero. Alimentación o Recuperaciones erróneas en valores de mineral complejo."): GoTo EH 'Error Div0!
      ReDim Preserve VecAux0(i)
      Let VecAux0(i) = DblAux1 / DblAux2
    Next i
    'Build the final Cx data
    ReDim VecResuCx(0 To 0, 0 To UBound(VecAux0))
    For i = 0 To UBound(VecAux0) 'Just one row (horizontal)
      Let VecResuCx(0, i) = VecAux0(i)
    Next i
    Let VecAux0 = Empty
  End If
  
  'Return
  If BEnuOrientation = eqlBalVertical Then
    Let VecResu = RES.ArrayTranspose(VecResu)
    If Not VBA.IsArray(VecResu) Then Call EHGLOBAL.ErrorHandlerSet(0.2, "Fallo al orientar verticalmente el resultado."): GoTo EH
    If BBooCx Then
      Let VecResuCx = RES.ArrayTranspose(VecResuCx)
      If Not VBA.IsArray(VecResu) Then Call EHGLOBAL.ErrorHandlerSet(0.2, "Fallo al orientar verticalmente el resultado."): GoTo EH
    End If
  End If
  Let BVarResponse = Array(VecResu, VecResuCx)
  Let VecResu = Empty
  Let VecResuCx = Empty
  
  Let Balmet_core_main_udf_calculate_eco = True

EH:
  Let VecWeig = Empty
  Let VecGrad = Empty
  Let VecRecs = Empty
  Let VecGradCx = Empty
  Let VecRecsCx = Empty
  Let VecUnit = Empty
  Let VecUnitCx = Empty
  Let VecAux0 = Empty
  Let VecResu = Empty
  Let VecResuCx = Empty

End Function

Private Function Balmet_core_main_app_eco(ByRef BVecRequest As Variant, ByRef BVecResponse As Variant, ByVal BEnuTypeBalEco As EQLBAL_ENU_RESULTECO) As Boolean
  
  'Data entry in BVarRequest: (DblFeed, DblFeedVol, VecGrad, VecGradCx, VecName, VecUnit, BytMethod, Booleans, EnuTypeFeed, RngBox, VecEcos)
  'VecEcos: Array(Array(DblFeedSolid, DblFeedVolume), VecWeights, VecGrades, VecRecs, VecGradesCx, VecRecsCx, VecUnits, VecUnitsCx, BEnuTypeBalEco)
  Dim RES As AppResources_current
  
  Dim VecAux0 As Variant, VecAux1 As Variant
  Dim BooCx As Boolean
  
  Const sFN As String = "APPBALMET::-core_main_app_eco"
  
  On Error GoTo EH
  Let Balmet_core_main_app_eco = False
  
  Set RES = New AppResources_current
  Set RES.ErrorHandler = EHGLOBAL
  Let VecAux1 = BVecRequest(10)
  Let BVecRequest = RES.ArrayPop(BVecRequest)
  Let BooCx = VBA.IsArray(VecAux1(4))
  'Get headsgrades or prodgrades from udf_calculation_eco
  'Request: Array(Array(DblFeedSolid, DblFeedVolume), VecWeights, VecGrades, VecRecs, VecGradesCx, VecRecsCx, VecUnits, VecUnitsCx)
  If Not Balmet_core_main_udf_calculate_eco(VecAux1, VecAux0, BooCx, RES, BEnuTypeBalEco, eqlBalHorizontal) Then GoTo EH
  'Response: Array(VecResu, VecResuCx)
  
  'Obtain vector with weights and 1D vectors of grades and head grades TODO: ING!!!
  If Not Balmet_core_main_app_eco_spreadbalmet_get_weights_data(VBA.CDbl(BVecRequest(0)), VecAux1(1), VecAux1(1), RES) Then GoTo EH
  
  'Get whole grades matrix from corresponding method
  If BEnuTypeBalEco = eqlBalHeadsGrades Then
    If Not Balmet_core_main_app_eco_spreadbalmet_get_grades_data(VBA.CDbl(BVecRequest(0)), VecAux1(1), VecAux1(2), VecAux0(0), VecAux1(4), VecAux0(1), BooCx, VecAux0, RES) Then GoTo EH
  Else
    If Not Balmet_core_main_app_eco_spreadbalmet_get_grades_data(VBA.CDbl(BVecRequest(0)), VecAux1(1), VecAux0(0), VecAux1(2), VecAux0(1), VecAux1(4), BooCx, VecAux0, RES) Then GoTo EH
  End If
  
  'Calculate Balance
  Let BVecRequest(2) = VecAux0(0)
  Let BVecRequest(3) = VecAux0(1)
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  If Not Balmet_core_main_app(BVecRequest) Then GoTo EH
  
  'Return
  Let BVecResponse = GLOVEC_BALANCE
  Let GLOVEC_BALANCE = Empty
  
  Let Balmet_core_main_app_eco = True
  
EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Balmet_core_main_app_eco_spreadbalmet_get_grades_data( _
  ByVal BDblFeed As Double, _
  ByVal BVarWeig As Variant, _
  ByVal BVarGrad As Variant, _
  ByVal BVarHead As Variant, _
  ByVal BVarGradCx As Variant, _
  ByVal BVarHeadCx As Variant, _
  ByVal BBooCx As Boolean, _
  ByRef BResponse As Variant, _
  ByVal RES As AppResources_current) As Boolean

  ' Builds the Balmet grades matrix from 1D param arrays _
    Feed: System Feed _
    Weig: Whole weights of balance, Prods, Tails and Feed _
    Grad: All main diagonal Prod Grades (GradCx too if it exists) _
    Head: All Head Grades (HeadCx too if it exists)
  'TODO: DIV/0!
  Dim VecAux0 As Variant
  Dim VecUi As Variant, VecPe As Variant, VecHeadsUi As Variant, VecUiZ As Variant 'accourate unities calculated to fill void spaces in Balmet grades
  Dim VecSemiFinalGrades As Variant, VecGradesCxSemi As Variant
  Dim StrGradeCx As String, StrHeadGradeCx As String
  Dim DblAux1 As Double
  Dim i As Long, j As Long, k As Long, m As Long, n As Long
  Dim StrFinalGrades As String, StrFinalGradesCx As String

  Const sFN = "APPBALMET::-core_main_spreadbalmet_get_grades_data_eco"
  
  On Error GoTo EH
  Let Balmet_core_main_app_eco_spreadbalmet_get_grades_data = False

  ' Objetive: _
    1. Generate %P _
    2. Generate U with %P * Grades _
    3. Get U[Feeds] - U[Grade] = U* _
    4. Generate U[i] with U*/(Products-1) _
    5. Generate Grades*
  
  'Generate %P: [Products, Tails, Feed]
  Let j = 0
  ReDim VecPe(UBound(BVarWeig) - LBound(BVarWeig))
  For i = LBound(BVarWeig) To UBound(BVarWeig)
    Let VecPe(j) = VBA.CDbl(BVarWeig(i)) / BDblFeed
    Let j = j + 1
  Next i

  'Calculate Unities
  ReDim VecUi(UBound(BVarGrad) - LBound(BVarGrad))
  For i = LBound(VecPe) To UBound(VecPe)
    Let VecUi(i) = VBA.CDbl(VecPe(i)) * VBA.CDbl(BVarGrad(i)) 'Just products 'cause Tails grades dosn't exists
  Next i

  'Calculate Head Unities
  ReDim VecHeadsUi(UBound(BVarHead) - LBound(BVarHead))
  For k = LBound(BVarHead) To UBound(BVarHead)
    Let VecHeadsUi(k) = VBA.CDbl(BVarHead(k)) * 100
  Next k
  
  'Calculate Ui* constants
  Let k = (UBound(BVarHead) - LBound(BVarHead)) + 1 'Products with no grades
  ReDim VecUiZ(UBound(BVarHead) - LBound(BVarHead))
  For i = LBound(VecUiZ) To UBound(VecUiZ)
    Let VecUiZ(i) = (VBA.CDbl(VecHeadsUi(i)) - VBA.CDbl(VecUi(i))) / k 'k is total products including Tails
  Next i

  Let m = (UBound(BVarWeig) - LBound(BVarWeig)) 'Includes: Products, Feed, Tails
  Let n = (UBound(BVarWeig) - LBound(BVarWeig))
  ReDim VecAux0(m)
  ReDim VecSemiFinalGrades(m)
  For j = 0 To n
    For i = 0 To m
      If i = j Then 'Main diagonal (Grades)
        Let VecAux0(i) = VBA.CDbl(VecUi(i)) / VBA.CDbl(VecPe(i))
      ElseIf i = m Then 'Heads
        Let VecAux0(i) = VBA.CDbl(VecHeadsUi(j)) / VBA.CDbl(VecPe(i))
      Else
        Let VecAux0(i) = VBA.CDbl(VecUiZ(j)) / VBA.CDbl(VecPe(i))
      End If
    Next i
    Let VecSemiFinalGrades(j) = VBA.Join(VecAux0, ";")
  Next j
  Let StrFinalGrades = VBA.Join(VecSemiFinalGrades, ";")
    
  'Calculate GradesCx
  If BBooCx Then
    ReDim VecGradesCxSemi(m)
    For i = 0 To UBound(VecGradesCxSemi)
      Let DblAux1 = (VBA.CDbl(BVarGradCx(i)) * VBA.CDbl(VecPe(m))) - (VBA.CDbl(BVarHeadCx(i)) * VBA.CDbl(VecPe(m))) / k 'U1*
      If i = 0 Then 'Main Grade
        Let VecGradesCxSemi(i) = BVarGradCx(i)
      ElseIf i = m Then 'HeadGrade
        Let VecGradesCxSemi(i) = BVarHeadCx(i)
      Else
        Let VecGradesCxSemi(i) = DblAux1 / VBA.CDbl(VecPe(i))
      End If
    Next i
    Let StrFinalGradesCx = VBA.Join(VecGradesCxSemi, ";")
  End If
  
  'Return
  Let BResponse = Array(StrFinalGrades, StrFinalGradesCx)
  
  Let Balmet_core_main_app_eco_spreadbalmet_get_grades_data = True

EH:
  Let BVarWeig = Empty
  Let VecUi = Empty
  Let VecPe = Empty
  Let VecHeadsUi = Empty
  Let VecUiZ = Empty
  Let VecSemiFinalGrades = Empty
  Let VecGradesCxSemi = Empty
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Balmet_core_main_app_eco_spreadbalmet_get_weights_data( _
  ByVal BDblFeed As Double, _
  ByVal BVarWeig As Variant, _
  ByRef BVarResp As Variant, _
  ByVal RES As AppResources_current)

  Dim VecAux0 As Variant
  Dim DblAux0 As Double
  
  Const sFN As String = "APPBALMET::-core_main_spreadbalmet_get_weights_data_eco"
  
  On Error GoTo EH
  Let Balmet_core_main_app_eco_spreadbalmet_get_weights_data = False
  
  Let DblAux0 = Application.WorksheetFunction.Sum(BVarWeig)
  Let DblAux0 = BDblFeed - DblAux0
  
  Let VecAux0 = BVarWeig
  Let VecAux0 = RES.ArrayAddAtLast(VecAux0, DblAux0)
  Let VecAux0 = RES.ArrayAddAtLast(VecAux0, BDblFeed)
  
  'Return
  Let BVarResp = VecAux0
  Let VecAux0 = Empty
  
  Let Balmet_core_main_app_eco_spreadbalmet_get_weights_data = True

EH:
  Let VecAux0 = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function


' CORE METHODS
Private Function Balmet_core_main_app_spreadbalmet_get_products_names( _
  ByVal BVecNames As Variant, _
  ByVal BLngRows As Long, _
  ByRef BVecResponse As Variant, _
  ByVal RES As AppResources_current, _
  Optional BBooHeads As Boolean = False) As Boolean

  'Data required: (VecNames)
  Dim MDL As ModelExcliqliteDatasheet
  
  Dim VecAux0 As Variant, VecAux1 As Variant, VecAux2 As Variant
  Dim StrAux0 As String
  'Dim BooExists As Boolean
  Dim i As Long, j As Long, xRow As Long, xCol As Long ', BLngRows As Long

  Const sFN As String = "APPBALMET::-core_spreadbalmet_get_products_names"

  On Error GoTo EH
  Let Balmet_core_main_app_spreadbalmet_get_products_names = False

  'Prepare array of names
  ''Let VecAux2 = excliqlitedatasheet.ListObjects("excliqlitechemicalelements").DataBodyRange.Value
  'Set MDL = New ModelApp
  'Set MDL.ErrorHandler = EHGLOBAL
  'Let VecAux2 = MDL.MGet(eqlMdlTblChemicalelements, eqlMdlArray)
  'Set MDL = Nothing
  ''Let i = (UBound(VecAux2) - LBound(VecAux2))
  ''ReDim VecAux0(0 To i)
  ''Let j = 0
  ''For i = LBound(VecAux2) To UBound(VecAux2)
  ''  Let VecAux0(j) = VecAux2(i, 2)
  ''  Let j = j + 1
  ''Next i
  'ReDim VecAux0(0 To UBound(VecAux2))
  'For i = LBound(VecAux2) To UBound(VecAux2)
  '  Let VecAux0(i) = VecAux2(i, 1)
  'Next i
  'ReDim VecAux1(0 To BLngRows)
  'For i = 0 To UBound(VecAux1)
  '  Let straux0 = VBA.LCase(VBA.CStr(BVecNames(i)))
  '  Let BooExists = RES.ArrayHasValue(VecAux0, straux0, False, True)
  '  If BooExists Then
  '    'Let VecAux1(i) = VecAux2(i + 1, 1)
  '    Let VecAux1(i) = VecAux2(i, 0)
  '  ElseIf (straux0 Like "b#" Or straux0 Like "bx#") Then
  '    Let VecAux1(i) = BVecNames(i)
  '  Else
  '    Call EHGLOBAL.ErrorHandlerSet(0.1, "¡Nombre de elemento: " & VBA.CStr(BVecNames(i)) & ", no existe!"): GoTo EH 'Error
  '  End If
  'Next i
  'Let VecAux0 = Empty
  'Let VecAux2 = Empty
  
  'Let VecAux2 = excliqlitedatasheet.ListObjects("excliqlitechemicalelements").DataBodyRange.Value
  Set MDL = New ModelExcliqliteDatasheet
  Set MDL.ErrorHandler = EHGLOBAL
  ReDim VecAux1(0 To BLngRows)
  For i = 0 To UBound(VecAux1)
    Let StrAux0 = VBA.LCase(VBA.CStr(BVecNames(i)))
    Let VecAux0 = MDL.MGet(eqlMdlTblChemicalelements, eqlMdlArray, "Elemento", MStrWhere:="Símbolo=" & StrAux0)
    If VBA.IsArray(VecAux0) Then
      Let VecAux1(i) = VecAux0(0, 0)
    ElseIf (StrAux0 Like "b#" Or StrAux0 Like "bx#") Then
      Let VecAux1(i) = BVecNames(i)
    Else
      Call EHGLOBAL.ErrorHandlerSet(0.1, "¡Nombre de elemento: " & VBA.CStr(BVecNames(i)) & ", no existe!"): GoTo EH 'Error
    End If
  Next i
  Let VecAux0 = Empty
  Set MDL = Nothing
  
  'Prepare main array
  Let xRow = BLngRows + 3 '3: Title, Feed, Tail
  Let xCol = 1
  ReDim VecAux0(0 To xRow, 0 To xCol)
  
  'Set Balmet main products names
  Let j = 0
  Let VecAux0(0, 0) = "#"
  Let VecAux0(0, 1) = "Producto"
  For i = 1 To UBound(VecAux0)
    If i <= (UBound(VecAux1) + 1) Then Let VecAux0(i, 1) = VBA.CStr(VecAux1(j))
    If i = (UBound(VecAux1) + 2) Then Let VecAux0(i, 1) = "C"
    If i = (UBound(VecAux1) + 3) Then Let VecAux0(i, 1) = "A"
    Let VecAux0(i, 0) = i
    Let j = j + 1
  Next i
  If BBooHeads Then
    ReDim VecAux2(0 To 1, 0 To 1)
    Let VecAux2(0, 0) = i
    Let VecAux2(1, 0) = i + 1
    Let VecAux2(0, 1) = "Cbz Ensayada"
    Let VecAux2(1, 1) = "Cbz Calculada"
    Let VecAux0 = RES.ArrayAddAtLast(VecAux0, VecAux2)
    If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(0.2, "Fallo al generar encabezados de Cabezas ensayada y calculada."): GoTo EH
  End If
  Let VecAux1 = Empty
  
  'Return
  Let BVecResponse = VecAux0
  Let VecAux0 = Empty
  Let Balmet_core_main_app_spreadbalmet_get_products_names = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Let VecAux2 = Empty
  Set MDL = Nothing
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Balmet_core_main_app_spreadbalmet_get_weights( _
  ByVal BDblFeed As Double, _
  ByRef BVarGrad As Variant, _
  ByVal BMethod As EQLBAL_ENU_METHOD, _
  ByRef BVarResponse As Variant, _
  ByVal RES As AppResources_current, _
  Optional BBooFormulas As Boolean = False, _
  Optional BBooHeads As Boolean = False, _
  Optional BBooWithHeader As Boolean = False, _
  Optional BRngBox As Range, _
  Optional BEnuTypeFeedWei As EQLBAL_ENU_TYPECONTENT = eqlBalSolids, _
  Optional BFullData As Boolean = False) As Boolean
  
  'Calculates the contents of main balmet (solids or volumes)
  Dim VecAux0 As Variant, VecAux1 As Variant
  Dim StrAux0 As String, StrAux1 As String
  Dim DblAux0 As Double, DblAux1 As Double
  Dim i As Long, j As Long, xRow As Long, xCol As Long, LngVolPos As Long

  Const sFN As String = "APPBALMET::-core_spreadbalmet_get_weights"

  On Error GoTo EH
  Let Balmet_core_main_app_spreadbalmet_get_weights = False

  'Prepare main array
  Let xRow = (UBound(BVarGrad) - LBound(BVarGrad))
  Let xCol = 0
  ReDim VecAux0(0 To xRow, 0 To xCol)
  ReDim VecAux1(xRow)
  
  'Calculate main results
  'Calculate by selected Method
  If Not Balmet_core_main_app_spreadbalmet_get_weights_generic(BDblFeed, BVarGrad, BMethod, VecAux0, RES) Then GoTo EH
  'If results has negatives, change position of head grades form top to bottom
  If Application.WorksheetFunction.MIN(VecAux0) < 0 Then
    Let VecAux0 = RES.ArrayShift(BVarGrad)
    If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(1.1, "Fallo calcular las cantidades de productos."): GoTo EH 'Error
    Let BVarGrad = RES.ArrayAddAtLast(BVarGrad, VecAux0)
    If Not VBA.IsArray(BVarGrad) Then Call EHGLOBAL.ErrorHandlerSet(1.1, "Fallo calcular las cantidades de productos."): GoTo EH 'Error
    If Not Balmet_core_main_app_spreadbalmet_get_weights_generic(BDblFeed, BVarGrad, BMethod, VecAux0, RES) Then GoTo EH
    If Application.WorksheetFunction.MIN(VecAux0) < 0 Then 'There is not valid datas to be a balmet
      Call EHGLOBAL.ErrorHandlerSet(1.2, "Leyes incorrectas. Imposible calcular el Balance Metalúrgico."): GoTo EH 'Error
    End If
  End If
  'If all is ok: get final data matrix
  If BFullData Then 'Include total
    Let VecAux0 = RES.ArrayAddAtLast(VecAux0, BDblFeed)
    If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(1.1, "Fallo calcular las cantidades de productos."): GoTo EH 'Error
  End If
  
  'Add Headers
  If BBooWithHeader Then
    Let VecAux0 = RES.ArrayAddAtFirst(VecAux0, IIf(BEnuTypeFeedWei = eqlBalVolumes, "Volumen[L]", "Peso[T]"))
    If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(1.3, "Fallo al crear encabezado del Volumen."): GoTo EH 'Error
  End If
  
  'Calculate Heads
  If BBooHeads Then
    ReDim VecAux1(0 To 1, 0 To 0)
    If BBooFormulas Then
      With BRngBox
        Let VecAux1(0, 0) = "=SUM(" & .Range("C3").Resize(xRow, 1).Address(False, False) & ")"
        Let VecAux1(1, 0) = "=SUM(" & .Range("C3").Resize(xRow, 1).Address(False, False) & ")"
      End With
    Else
      Let VecAux1(0, 0) = BDblFeed
      Let VecAux1(1, 0) = BDblFeed
    End If
    Let VecAux0 = RES.ArrayAddAtLast(VecAux0, VecAux1)
    If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(1.4, "Fallo al calcular el Volumen de cabezas calculada y ensayada."): GoTo EH 'Error
    Let VecAux1 = Empty
  End If
  
  'Return
  Let BVarResponse = VecAux0
  Let VecAux0 = Empty
  Let Balmet_core_main_app_spreadbalmet_get_weights = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Balmet_core_main_app_spreadbalmet_get_weights_generic( _
  ByVal BDblFeed As Double, _
  ByVal BVecGrad As Variant, _
  ByVal BMethod As EQLBAL_ENU_METHOD, _
  ByRef BVecResponse As Variant, _
  ByVal RES As AppResources_current) As Boolean
  
  'Data required: (StrFeed, StrGrade)
  Const sFN As String = "APPBALMET::-core_main_spreadbalmet_get_weights_generic"
  
  On Error GoTo EH
  Let Balmet_core_main_app_spreadbalmet_get_weights_generic = False
  
  'Calculate by Method selected
  Select Case BMethod
    Case eqlBalConventional 'Conventional
      If Not Balmet_core_main_app_spreadbalmet_get_weights_conventional(BDblFeed, BVecGrad, BVecResponse) Then GoTo EH
    Case eqlBalCramer 'CRAMER
      If Not Balmet_core_main_app_spreadbalmet_get_weights_cramer(BDblFeed, BVecGrad, BVecResponse, RES) Then GoTo EH
    Case Else 'Inverse matrix
      If Not Balmet_core_main_app_spreadbalmet_get_weights_inverse(BDblFeed, BVecGrad, BVecResponse, RES) Then GoTo EH
  End Select
  
  'Return 2DVec
  Let Balmet_core_main_app_spreadbalmet_get_weights_generic = True

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Balmet_core_main_app_spreadbalmet_get_weights_conventional( _
  ByVal BDblFeed As Double, _
  ByVal BVecGrad As Variant, _
  ByRef BVecResponse As Variant) As Boolean
  
  'Data required: (VecFeed, VecGrade)
  Dim VecAux As Variant
  Dim DblAux0 As Double, DblAux1 As Double, DblFeedG As Double, DblGradeG As Double, DblTailG As Double
  
  Const sFN As String = "APPBALMET::-core_main_spreadbalmet_get_weights_conventional"
  
  On Error GoTo EH
  Let Balmet_core_main_app_spreadbalmet_get_weights_conventional = False
    
  'Conventional method, only for balances with 2 products
  If (UBound(BVecGrad) + 1) > 3 Then Call EHGLOBAL.ErrorHandlerSet(1.11, "Balances de dos productos sólo pueden tener tres elementos."): GoTo EH  'Not avaiable for more than two products
  
  Let DblGradeG = VBA.Abs(VBA.CDbl(BVecGrad(0, 0))) '0
  Let DblTailG = VBA.Abs(VBA.CDbl(BVecGrad(1, 0))) '1
  Let DblFeedG = VBA.Abs(VBA.CDbl(BVecGrad(2, 0))) '2
  
  Let DblAux0 = DblGradeG - DblTailG
  
  #If Debugging Then
    Call Debug_assert((DblAux0 = 0), sFN)
  #End If
  
  If DblAux0 = 0 Then Call EHGLOBAL.ErrorHandlerSet(1.112, "El determinante es cero, el Balance no tiene solución."): GoTo EH 'Error Div0!
  
  'Resolve
  ReDim VecAux(0 To 1, 0 To 0)
  '%Rec
  Let DblAux1 = (DblGradeG * (DblFeedG - DblTailG) * 100) / (DblFeedG * DblAux0)
  'B
  Let VecAux(0, 0) = (DblAux1 * BDblFeed * DblFeedG) / (DblGradeG * 100)
  'C
  Let VecAux(1, 0) = (BDblFeed - VecAux(0, 0))
  
  'Forming the response: (Vec[mx1])
  Let BVecResponse = VecAux
  Let Balmet_core_main_app_spreadbalmet_get_weights_conventional = True

EH:
  Let VecAux = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Balmet_core_main_app_spreadbalmet_get_weights_cramer( _
  ByVal BDblFeed As Double, _
  ByVal BVecGrad As Variant, _
  ByRef BVecResponse As Variant, _
  ByVal RES As AppResources_current) As Boolean
 
  'Data required: (VecFeed, VecGrade)
  Dim VecAux As Variant
  Dim VecHeadGrades As Variant, VecProdGrades As Variant, VecAuxDeterm0 As Variant
  Dim DblDeterm As Double, DblDetermB As Double
  Dim i As Long, j As Long, k As Long, xRow As Long
  
  Const sFN As String = "APPBALMET::-core_main_spreadbalmet_get_weights_cramer"
  
  On Error GoTo EH
  Let Balmet_core_main_app_spreadbalmet_get_weights_cramer = False
  
  'Resolve Balmet
  'Get separated arrays
  Let VecProdGrades = BVecGrad
  Let VecHeadGrades = RES.ArrayPop(VecProdGrades)
  If Not VBA.IsArray(VecProdGrades) Then Call EHGLOBAL.ErrorHandlerSet(1.113, "Fallo al ordenar la matriz de leyes."): GoTo EH 'Error
  Let VecProdGrades = RES.ArrayAddAtFirst(VecProdGrades, 1, True)
  If Not VBA.IsArray(VecProdGrades) Then Call EHGLOBAL.ErrorHandlerSet(1.113, "Fallo al ordenar la matriz de leyes."): GoTo EH 'Error
  Let VecProdGrades = Application.WorksheetFunction.Transpose(VecProdGrades)
  Let VecHeadGrades = RES.ArrayAddAtFirst(VecHeadGrades, 1, True)
  If Not VBA.IsArray(VecHeadGrades) Then Call EHGLOBAL.ErrorHandlerSet(1.113, "Fallo al ordenar la matríz de leyes."): GoTo EH 'Error
  Let VecHeadGrades = Application.WorksheetFunction.Transpose(VecHeadGrades)
  Let xRow = (UBound(VecHeadGrades) - LBound(VecHeadGrades))
  
  '2. [x]: Make a matrix AuxDeterm dynamic mxn with wich calculate each final weight value
  Let DblDeterm = Application.WorksheetFunction.MDeterm(VecProdGrades)
  If DblDeterm = 0 Then Call EHGLOBAL.ErrorHandlerSet(1.114, "El determinante es cero. Balance sin solución."): GoTo EH 'Error Div0!
  
  ReDim VecAux(0 To xRow, 0 To 0)
  Let k = 0
  For i = 0 To UBound(VecAux)
    Let VecAuxDeterm0 = VecProdGrades
    For j = LBound(VecHeadGrades) To UBound(VecHeadGrades)
      Let VecAuxDeterm0(k + LBound(VecAuxDeterm0), i + LBound(VecAuxDeterm0, 2)) = VecHeadGrades(j, LBound(VecHeadGrades))
      Let k = k + 1
    Next j
    Let DblDetermB = Application.WorksheetFunction.MDeterm(VecAuxDeterm0)
    '3. [x]: Fill a matrix with Results values mx1
    Let VecAux(i, 0) = BDblFeed * (DblDetermB / DblDeterm)
    Let k = 0
  Next i
  
  'Return: Forming the response: (Vec[mxn])
  Let BVecResponse = VecAux
  Let Balmet_core_main_app_spreadbalmet_get_weights_cramer = True

EH:
  Let BVecGrad = Empty
  Let VecAux = Empty
  Let VecHeadGrades = Empty
  Let VecProdGrades = Empty
  Let VecAuxDeterm0 = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Balmet_core_main_app_spreadbalmet_get_weights_inverse( _
  ByVal BDblFeed As Double, _
  ByVal BVecGrad As Variant, _
  ByRef BVecResponse As Variant, _
  ByVal RES As AppResources_current) As Boolean
  
  'Data required: (VecFeed, VecGrade)
  Dim VecAux As Variant
  Dim VecHeadGrades As Variant, VecProdGrades As Variant, VecAuxDeterm0 As Variant
  Dim DblDeterm As Double
  Dim i As Long, j As Long, k As Long, l As Long

  Const sFN As String = "APPBALMET::-core_spreadbalmet_get_weights_inverse"

  On Error GoTo EH
  Let Balmet_core_main_app_spreadbalmet_get_weights_inverse = False

  'Resolve Balmet
  'Get separated arrays
  Let VecProdGrades = BVecGrad
  Let VecHeadGrades = RES.ArrayPop(VecProdGrades)
  If Not VBA.IsArray(VecProdGrades) Then Call EHGLOBAL.ErrorHandlerSet(1.113, "Fallo al ordenar la matriz de leyes."): GoTo EH 'Error
  Let VecProdGrades = RES.ArrayAddAtFirst(VecProdGrades, 1, True)
  If Not VBA.IsArray(VecProdGrades) Then Call EHGLOBAL.ErrorHandlerSet(1.113, "Fallo al ordenar la matriz de leyes."): GoTo EH 'Error
  Let VecProdGrades = Application.WorksheetFunction.Transpose(VecProdGrades)
  Let VecHeadGrades = RES.ArrayAddAtFirst(VecHeadGrades, 1, True)
  If Not VBA.IsArray(VecHeadGrades) Then Call EHGLOBAL.ErrorHandlerSet(1.113, "Fallo al ordenar la matríz de leyes."): GoTo EH 'Error
  Let VecHeadGrades = Application.WorksheetFunction.Transpose(VecHeadGrades)
  'Let xRow = (UBound(VecHeadGrades) - LBound(VecHeadGrades))
  
  For i = LBound(VecHeadGrades) To UBound(VecHeadGrades)
    Let VecHeadGrades(i, LBound(VecHeadGrades, 2)) = BDblFeed * VBA.CDbl(VecHeadGrades(i, LBound(VecHeadGrades, 2)))
  Next i

  'Make a matrix AuxDeterm dynamic mxn with wich calculate each final weight value
  Let DblDeterm = Application.WorksheetFunction.MDeterm(VecProdGrades)
  If DblDeterm = 0 Then Call EHGLOBAL.ErrorHandlerSet(1.114, "El determinante es cero, el Balance no tiene solución."): GoTo EH  'Error

  'Apply inverse matrix method and then fill a matrix with Results values mx1
  Let VecAuxDeterm0 = Application.WorksheetFunction.MInverse(VecProdGrades)
  Let VecAux = Application.WorksheetFunction.MMult(VecAuxDeterm0, VecHeadGrades)

  'Return
  Let k = 0
  Let l = 0
  Let i = UBound(VecAux) - LBound(VecAux)
  Let j = UBound(VecAux, 2) - LBound(VecAux, 2)
  ReDim BVecResponse(0 To i, 0 To j)
  For j = LBound(VecAux, 2) To UBound(VecAux, 2)
    For i = LBound(VecAux) To UBound(VecAux)
      Let BVecResponse(k, l) = VecAux(i, j)
      Let k = k + 1
    Next i
    Let l = l + 1
    Let k = 0
  Next j

  Let Balmet_core_main_app_spreadbalmet_get_weights_inverse = True

EH:
  Let VecAux = Empty
  Let VecHeadGrades = Empty
  Let VecProdGrades = Empty
  Let VecAuxDeterm0 = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Balmet_core_main_app_spreadbalmet_get_percents( _
  ByVal BVarContents As Variant, _
  ByRef BVarResponse As Variant, _
  ByVal RES As AppResources_current, _
  Optional BBooFormulas As Boolean = False, _
  Optional BBooHeads As Boolean = False, _
  Optional BBooWithHeader As Boolean = False, _
  Optional BEnuTypeFeedPer As EQLBAL_ENU_TYPECONTENT = eqlBalSolids, _
  Optional BRngBox As Range) As Boolean

  'Calculates a column of percentages from values of contents (numeric results or by formulas)
  Dim VecAux0 As Variant, VecAux1 As Variant
  Dim DblFeed As Double
  Dim i As Long, j As Long, xRow As Long, xCol As Long

  Const sFN As String = "APPBALMET::-core_spreadbalmet_get_weights_percents"

  On Error GoTo EH
  Let Balmet_core_main_app_spreadbalmet_get_percents = False

  'Prepare main array
  Let xRow = (UBound(BVarContents) - LBound(BVarContents))
  Let xCol = 0
  ReDim VecAux0(0 To xRow, 0 To xCol)
  
  'Calculate main results
  If BBooFormulas Then
    'Let j = (xRow + 1)
    With BRngBox
      For i = 0 To xRow
        If i < xRow Then Let VecAux0(i, 0) = "=(" & .Range("C" & (3 + i)).Address(False, False) & "/" & .Range("C" & (3 + xRow)).Address & ")*100"
        If i = xRow Then Let VecAux0(i, 0) = "=SUM(" & .Range("D3").Resize(xRow, 1).Address(False, False) & ")"
      Next i
    End With
  Else
    Let DblFeed = VBA.CDbl(BVarContents(xRow, 0))
    Let j = LBound(BVarContents)
    For i = 0 To xRow
      Let VecAux0(i, 0) = (VBA.CDbl(BVarContents(j, 0)) / DblFeed) * 100
      Let j = j + 1
    Next i
  End If
  
  'Add Headers
  If BBooWithHeader Then
    Let VecAux0 = RES.ArrayAddAtFirst(VecAux0, IIf(BEnuTypeFeedPer = eqlBalVolumes, "Volumen[%]", "Peso[%]"))
    If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(3.1, "Fallo al calcular el Porcentaje en " & IIf(BEnuTypeFeedPer = eqlBalSolids, "peso", "volumen") & "."): GoTo EH  'Error
  End If
  
  'Calculate Heads
  If BBooHeads Then
    ReDim VecAux1(0 To 1, 0 To 0)
    If BBooFormulas Then
      With BRngBox
        Let VecAux1(0, 0) = "=SUM(" & .Range("D3").Resize(xRow, 1).Address(False, False) & ")"
        Let VecAux1(1, 0) = "=SUM(" & .Range("D3").Resize(xRow, 1).Address(False, False) & ")"
      End With
    Else
      Let VecAux1(0, 0) = (DblFeed / DblFeed) * 100
      Let VecAux1(1, 0) = (DblFeed / DblFeed) * 100
    End If
    Let VecAux0 = RES.ArrayAddAtLast(VecAux0, VecAux1)
    If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(3.2, "Fallo al calcular el Porcentaje en peso de cabezas calculada y ensayada."): GoTo EH 'Error
  End If
  
  'Return
  Let BVarResponse = VecAux0
  Let VecAux0 = Empty
  Let Balmet_core_main_app_spreadbalmet_get_percents = True

EH:
  Let VecAux0 = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Balmet_core_main_app_spreadbalmet_get_volumes( _
  ByVal BDblFeedVol As Double, _
  ByVal BVarFeed As Variant, _
  ByVal BVarGrad As Variant, _
  ByVal BVarGradCx As Variant, _
  ByVal BVarUnit As Variant, _
  ByRef BVarResponse As Variant, _
  ByVal RES As AppResources_current, _
  Optional BBooFormulas As Boolean = False, _
  Optional BBooWithHeader As Boolean = False, _
  Optional BBooHeads As Boolean = False, _
  Optional BRngBox As Range, _
  Optional BBooPercent As Boolean = False) As Boolean

  'Calculates the volume contents of main balmet
  Dim VecAux0 As Variant, VecAux1 As Variant
  Dim StrAux0 As String, StrAux1 As String
  Dim DblAux0 As Double, DblAux1 As Double
  Dim i As Long, j As Long, xRow As Long, xCol As Long, LngVolPos As Long

  Const sFN As String = "APPBALMET::-core_spreadbalmet_get_volumes"

  On Error GoTo EH
  Let Balmet_core_main_app_spreadbalmet_get_volumes = False

  'Prepare main array
  Let xRow = (UBound(BVarGrad) - LBound(BVarGrad))
  Let xCol = 0
  
  'Calculate main results
  If Not Balmet_core_main_app_spreadbalmet_get_volumes_generic(BDblFeedVol, BVarFeed, BVarGrad, BVarGradCx, BVarUnit, VecAux0, RES) Then GoTo EH
  
  'Add Headers
  If BBooWithHeader Then
    Let VecAux0 = RES.ArrayAddAtFirst(VecAux0, "Volumen[L]")
    If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(2.1, "Fallo al crear encabezado del Volumen."): GoTo EH 'Error
  End If
  
  'Calculate Heads
  If BBooHeads Then
    ReDim VecAux1(0 To 1, 0 To 0)
    If BBooFormulas Then
      Let StrAux0 = IIf(BBooPercent, "E3", "D3")
      With BRngBox.Range(StrAux0)
        Let VecAux1(0, 0) = "=SUM(" & .Resize(xRow, 1).Address(False, False) & ")"
        Let VecAux1(1, 0) = "=SUM(" & .Resize(xRow, 1).Address(False, False) & ")"
      End With
    Else
      Let VecAux1(0, 0) = VecAux0(UBound(VecAux0), LBound(VecAux0, 2))
      Let VecAux1(1, 0) = VecAux0(UBound(VecAux0), LBound(VecAux0, 2))
    End If
    Let VecAux0 = RES.ArrayAddAtLast(VecAux0, VecAux1)
    If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(2.2, "Fallo al calcular el Volumen de cabezas calculada y ensayada."): GoTo EH 'Error
    Let VecAux1 = Empty
  End If
  
  'Return
  Let BVarResponse = VecAux0
  Let VecAux0 = Empty
  Let Balmet_core_main_app_spreadbalmet_get_volumes = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Balmet_core_main_app_spreadbalmet_get_volumes_generic( _
  ByVal BDblFeedVol As Double, _
  ByVal BVarFeed As Variant, _
  ByVal BVarGrad As Variant, _
  ByVal BVarGradCx As Variant, _
  ByVal BVarUnit As Variant, _
  ByRef BVarResponse As Variant, _
  ByVal RES As AppResources_current) As Boolean

  'Data required: (BDblFeedVol, BVarFeed, BVarGrad, BVarGradCx, BVarUnit, BVarResponse)
  Dim VecAux0 As Variant, VecAux1 As Variant
  Dim DblAux0 As Double, DblAux1 As Double
  Dim i As Long, j As Long, xRow As Long, xCol As Long, LngVolPos As Long

  Const sFN As String = "APPBALMET::-core_spreadbalmet_get_volumes_generic"

  On Error GoTo EH
  Let Balmet_core_main_app_spreadbalmet_get_volumes_generic = False

  'Prepare main array
  Let xRow = (UBound(BVarGrad) - LBound(BVarGrad))
  Let xCol = 0
  
  'Prepare basic data (gettin' first occurence of g/l)
  Let VecAux1 = RES.ArrayConcat(BVarGrad, BVarGradCx, RByCols:=True)
  If Not VBA.IsArray(VecAux1) Then Call EHGLOBAL.ErrorHandlerSet(2.11, "Error en el cálculo genérico del volumen."): GoTo EH
  ReDim VecAux0(0 To xRow, 0 To xCol)
  For j = LBound(VecAux1, 2) To UBound(VecAux1, 2)
    If VBA.LCase(VBA.CStr(BVarUnit(j))) = "g/l" Then
      For i = LBound(VecAux1) To UBound(VecAux1)
        Let VecAux0(i, 0) = VecAux1(i, j)
      Next i
      Exit For
    End If
  Next j
  Let VecAux1 = Empty
  
  'Calculate recoveries
  'Gettin' percents
  Let DblAux0 = VBA.CDbl(BVarFeed(xRow, 0))
  ReDim VecAux1(0 To xRow, 0 To 0)
  For i = LBound(VecAux1) To UBound(VecAux1)
    Let VecAux1(i, 0) = (VBA.CDbl(BVarFeed(i, 0)) / DblAux0) * 100
  Next i
  'Gettin' unities (% * g/l)
  Let DblAux0 = 0
  For i = LBound(VecAux1) To UBound(VecAux1)
    If i < UBound(VecAux1) Then
      Let VecAux1(i, 0) = (VBA.CDbl(VecAux1(i, 0)) * VBA.CDbl(VecAux0(i, 0)))
      Let DblAux0 = DblAux0 + VBA.CDbl(VecAux1(i, 0))
    Else
      Let VecAux1(i, 0) = DblAux0
    End If
  Next i
  'Gettin' recoveries
  For i = LBound(VecAux1) To UBound(VecAux1)
    Let VecAux1(i, 0) = (VBA.CDbl(VecAux1(i, 0)) / VBA.CDbl(VecAux1(xRow, 0))) * 100
  Next i
  
  'Calculate fines (feed vol [L] * grade vol [g/L])
  Let DblAux1 = BDblFeedVol * VBA.CDbl(VecAux0(xRow, 0))
  For i = LBound(VecAux1) To UBound(VecAux1)
    Let VecAux1(i, 0) = (VBA.CDbl(VecAux1(i, 0)) * DblAux1) / 100
  Next i
  
  'Calculate Volumes
  Let DblAux0 = 0
  For i = LBound(VecAux1) To UBound(VecAux1)
    Let VecAux0(i, 0) = (VBA.CDbl(VecAux1(i, 0)) / VBA.CDbl(VecAux0(i, 0)))
    If i < (UBound(VecAux1) - 1) Then Let DblAux0 = DblAux0 + VBA.CDbl(VecAux0(i, 0))
    If i = (UBound(VecAux1) - 1) Then Let VecAux0(i, 0) = BDblFeedVol - DblAux0 'Adjustment for Tails
    If i = (UBound(VecAux1)) Then Let VecAux0(i, 0) = BDblFeedVol
  Next i
    
  'Return
  Let BVarResponse = VecAux0
  Let VecAux0 = Empty
  Let Balmet_core_main_app_spreadbalmet_get_volumes_generic = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Balmet_core_main_app_spreadbalmet_get_volumes_percents( _
  ByVal BVarFeedsVol As Variant, _
  ByRef BVarResponse As Variant, _
  ByVal RES As AppResources_current, _
  Optional BBooFormulas As Boolean = False, _
  Optional BBooWithHeader As Boolean = False, _
  Optional BBooHeads As Boolean = False, _
  Optional BRngBox As Range) As Boolean

  'Calculates the volume percents of main balmet
  Dim VecAux0 As Variant, VecAux1 As Variant
  Dim StrAux0 As String
  Dim DblFeedVol As Double
  Dim i As Long, j As Long, xRow As Long, xCol As Long, LngVolPos As Long

  Const sFN As String = "APPBALMET::-core_spreadbalmet_get_volumes_percents"

  On Error GoTo EH
  Let Balmet_core_main_app_spreadbalmet_get_volumes_percents = False

  'Prepare main array
  Let xRow = (UBound(BVarFeedsVol) - LBound(BVarFeedsVol))
  Let xCol = 0
  ReDim VecAux0(0 To xRow, 0 To xCol)
  ReDim VecAux1(xRow)

  'Calculate main results
  If BBooFormulas Then
    With BRngBox.Range("F3")
      For i = 0 To xRow
        If i < xRow Then Let VecAux0(i, 0) = "=(" & .Offset(i, 0).Address(False, False) & "/" & .Offset(xRow, 0).Address & ")*100"
        If i = xRow Then Let VecAux0(i, 0) = "=SUM(" & .Resize(xRow, 1).Address(False, False) & ")"
      Next i
    End With
  Else
    Let DblFeedVol = VBA.CDbl(BVarFeedsVol(xRow, 0))
    Let j = LBound(BVarFeedsVol)
    For i = 0 To xRow
      Let VecAux0(i, 0) = (VBA.CDbl(BVarFeedsVol(j, 0)) / DblFeedVol) * 100
      Let j = j + 1
    Next i
  End If

  'Add Headers
  If BBooWithHeader Then
    Let VecAux0 = RES.ArrayAddAtFirst(VecAux0, "Volumen[%]")
    If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(4.1, "Fallo al crear encabezado del Volumen."): GoTo EH 'Error
  End If

  'Calculate Heads
  If BBooHeads Then
    ReDim VecAux1(0 To 1, 0 To 0)
    If BBooFormulas Then
      With BRngBox.Range("F3")
        Let VecAux1(0, 0) = "=SUM(" & .Resize(xRow, 1).Address(False, False) & ")"
        Let VecAux1(1, 0) = "=SUM(" & .Resize(xRow, 1).Address(False, False) & ")"
      End With
    Else
      Let VecAux1(0, 0) = (DblFeedVol / DblFeedVol) * 100
      Let VecAux1(1, 0) = (DblFeedVol / DblFeedVol) * 100
    End If
    Let VecAux0 = RES.ArrayAddAtLast(VecAux0, VecAux1)
    If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(4.2, "Fallo al calcular el Volumen de cabezas calculada y ensayada."): GoTo EH 'Error
  End If

  'Return
  Let BVarResponse = VecAux0
  Let VecAux0 = Empty
  Let Balmet_core_main_app_spreadbalmet_get_volumes_percents = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Balmet_core_main_app_spreadbalmet_get_grades( _
  ByVal BVarContents As Variant, _
  ByVal BVarVolumes As Variant, _
  ByVal BVarGrades As Variant, _
  ByVal BVarGradesCx As Variant, _
  ByVal BVarUnits As Variant, _
  ByVal BVarNames As Variant, _
  ByRef BVarResponse As Variant, _
  ByVal BAllResult As Boolean, _
  ByVal RES As AppResources_current, _
  Optional BBooFormulas As Boolean = False, _
  Optional BBooWithHeader As Boolean = False, _
  Optional BBooHeads As Boolean = False, _
  Optional BRngBox As Range, _
  Optional BBooJustHeads As Boolean = False, _
  Optional BBooCx As Boolean = False, _
  Optional BBooPercents As Boolean = False, _
  Optional BEnuTypeFeedPer As EQLBAL_ENU_TYPECONTENT = eqlBalSolids, _
  Optional BStrVol As String = VBA.vbNullString, _
  Optional BLngGra As Long = 0, _
  Optional BLngCur As Long = 0) As Boolean

  'Calculates the all grades of main balmet
  Dim VecAux0 As Variant, VecAux1 As Variant, VecAux2 As Variant, VecAux3 As Variant, VecAux4 As Variant
  Dim StrFeed As String, StrAux0 As String
  Dim BooVol As Boolean
  Dim DblAux0 As Double, DblFeed As Double
  Dim i As Long, j As Long, k As Long, xRow As Long, xCol As Long

  Const sFN As String = "APPBALMET::-core_spreadbalmet_get_weights_grades"

  On Error GoTo EH
  Let Balmet_core_main_app_spreadbalmet_get_grades = False

  'Calculate main results
  If Not BBooJustHeads Then
    Let VecAux0 = RES.ArrayAddAtLast(BVarGrades, BVarGradesCx, True)
    If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(5.1, "Fallo al generar matriz de Leyes."): GoTo EH 'Error
  End If
  
  'Add Headers
  If BBooWithHeader Then
    ReDim VecAux1(UBound(BVarNames))
    For i = 0 To UBound(VecAux1)
      Let VecAux1(i) = VBA.CStr(BVarNames(i)) & "[" & VBA.CStr(BVarUnits(i)) & "]"
    Next i
    Let VecAux0 = RES.ArrayAddAtFirst(VecAux0, VecAux1)
    If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(5.2, "Fallo al crear encabezados de Leyes."): GoTo EH 'Error
  End If
  
  'Calculate Heads
  If BBooHeads Then
    'Method calls two other methods named calculated and tested
    If BBooFormulas Then
      
      Let xRow = (UBound(BVarGrades) - LBound(BVarGrades))
      Let xCol = UBound(BVarNames)
      ReDim VecAux3(0 To xRow - 1)
      ReDim VecAux2(0 To 1, 0 To xCol)
      For i = 0 To xCol
        Let BooVol = (VBA.LCase(VBA.CStr(BVarUnits(i))) = "g/l")
        Let StrAux0 = IIf(BooVol, BStrVol, "C3")
        Let k = IIf(BEnuTypeFeedPer = eqlBalBoth And BooVol, IIf(BBooPercents, 2, 1), BLngGra)
        With BRngBox.Range(StrAux0)
          Let VecAux2(0, i) = "=" & .Offset(xRow, k + i).Address(False, False)
          For j = 0 To (xRow - 1)
            Let VecAux3(j) = "(" & .Offset(j, 0).Address(False, False) & "*" & .Offset(j, k + i).Address(False, False) & ")"
          Next j
          Let VecAux2(1, i) = "=SUM(" & VBA.Join(VecAux3, ";") & ")/" & .Offset(xRow, 0).Address(False, False)
        End With
      Next i
      Let VecAux1 = Empty
      Let VecAux3 = Empty
      Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux2)
      If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(5.4, "Fallo al crear encabezados de Leyes."): GoTo EH 'Error
      Let VecAux2 = Empty
    
    Else
      
      If BBooCx Then 'Just Cx
        Let VecAux1 = BVarGradesCx
        Let VecAux2 = RES.ArraySlice(BVarUnits, UBound(BVarGrades, 2))
      Else
        If BAllResult Then 'Concat both matrices
          Let VecAux1 = RES.ArrayConcat(BVarGrades, BVarGradesCx, RByCols:=True)
          Let VecAux2 = BVarUnits
        Else 'Just grades of normal balmet
          Let VecAux1 = BVarGrades
          Let VecAux2 = RES.ArraySlice(BVarUnits, UBound(BVarGrades, 2), True)
          'Let VecAux2 = BVarUnits
        End If
      End If
      If Not VBA.IsArray(VecAux1) Or Not VBA.IsArray(VecAux2) Then Call EHGLOBAL.ErrorHandlerSet(5.3, "Fallo al generar matriz de Leyes de cabeza ensayada y calculada."): GoTo EH 'Error
      'Obtain %
      Let DblFeed = VBA.CDbl(BVarContents(UBound(BVarContents), 0))
      For i = LBound(BVarContents) To UBound(BVarContents)
        Let BVarContents(i, 0) = (VBA.CDbl(BVarContents(i, 0)) / DblFeed) * 100
      Next i
      If VBA.IsArray(BVarVolumes) Then
        Let DblFeed = VBA.CDbl(BVarVolumes(UBound(BVarVolumes), 0))
        For i = LBound(BVarVolumes) To UBound(BVarVolumes)
          Let BVarVolumes(i, 0) = (VBA.CDbl(BVarVolumes(i, 0)) / DblFeed) * 100
        Next i
      End If
      'Fill final vector with tested grades
      Let VecAux3 = RES.ArrayPop(VecAux1, True)
      If Not VBA.IsArray(VecAux3) Then Call EHGLOBAL.ErrorHandlerSet(5.4, "Fallo al crear encabezados de Leyes."): GoTo EH 'Error
      'Obtain U
      For j = LBound(VecAux2) To UBound(VecAux2) 'Units
        For i = LBound(VecAux1) To UBound(VecAux1) 'Grades
          If Not VBA.LCase(VBA.CStr(VecAux2(j))) = "g/l" Then
            Let VecAux1(i, j) = (VBA.CDbl(BVarContents(i, 0)) * VBA.CDbl(VecAux1(i, j)))
          Else
            Let VecAux1(i, j) = (VBA.CDbl(BVarVolumes(i, 0)) * VBA.CDbl(VecAux1(i, j)))
          End If
        Next i
      Next j
      For j = LBound(VecAux1, 2) To UBound(VecAux1, 2) 'Unities
        For i = LBound(VecAux1) To UBound(VecAux1)
          If i < UBound(VecAux1) Then
            Let DblAux0 = DblAux0 + (VBA.CDbl(VecAux1(i, j)))
          Else
            Let VecAux1(i, j) = DblAux0 / 100
          End If
        Next i
      Next j
      'Fill final vector with calculated grades
      Let VecAux2 = Empty
      Let VecAux2 = RES.ArrayPop(VecAux1)
      If Not VBA.IsArray(VecAux2) Then Call EHGLOBAL.ErrorHandlerSet(5.4, "Fallo al crear encabezados de Leyes."): GoTo EH 'Error
      Let VecAux1 = Empty
      Let VecAux1 = RES.ArrayConcat(VecAux3, VecAux2)
      If Not VBA.IsArray(VecAux1) Then Call EHGLOBAL.ErrorHandlerSet(5.4, "Fallo al crear encabezados de Leyes."): GoTo EH 'Error
      Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1)
      If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(5.4, "Fallo al crear encabezados de Leyes."): GoTo EH 'Error
    
    End If
  End If

  'Return
  Let BVarResponse = VecAux0
  Let VecAux0 = Empty
  Let Balmet_core_main_app_spreadbalmet_get_grades = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Let VecAux2 = Empty
  Let VecAux3 = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Balmet_core_main_app_spreadbalmet_get_unities( _
  ByVal BVarContents As Variant, _
  ByVal BVarVolumes As Variant, _
  ByVal BVarGrades As Variant, _
  ByVal BVarGradesCx As Variant, _
  ByVal BVarUnits As Variant, _
  ByVal BVarNames As Variant, _
  ByRef BVarResponse As Variant, _
  ByVal BAllResult As Boolean, _
  ByVal RES As AppResources_current, _
  Optional BBooFormulas As Boolean = False, _
  Optional BBooWithHeader As Boolean = False, _
  Optional BBooHeads As Boolean = False, _
  Optional BRngBox As Range, _
  Optional BBooCx As Boolean = False, _
  Optional BBooPercents As Boolean = False, _
  Optional BEnuTypeFeedPer As EQLBAL_ENU_TYPECONTENT = eqlBalSolids, _
  Optional BStrVol As String = VBA.vbNullString, _
  Optional BLngGra As Long = 0, _
  Optional BLngCur As Long = 0) As Boolean

  'Calculates the all unities of main balmet
  Dim VecAux0 As Variant, VecAux1 As Variant, VecAux2 As Variant, VecAux3 As Variant
  Dim StrAux0 As String
  Dim BooVol As Boolean
  Dim DblAux0 As Double, DblAux1 As Double
  Dim i As Long, j As Long, k As Long, l As Long, xRow As Long, xCol As Long

  Const sFN As String = "APPBALMET::-core_spreadbalmet_get_unities"

  On Error GoTo EH
  Let Balmet_core_main_app_spreadbalmet_get_unities = False
  
  'Calculate main results
  Let xRow = (UBound(BVarGrades) - LBound(BVarGrades))
  Let xCol = UBound(BVarNames)
  If BBooFormulas Then
  
      ReDim VecAux0(0 To xRow, 0 To xCol)
      For j = 0 To xCol
        Let BooVol = (VBA.LCase(VBA.CStr(BVarUnits(j))) = "g/l")
        Let StrAux0 = IIf(BooVol, BStrVol, "C3")
        Let k = IIf(BEnuTypeFeedPer = eqlBalBoth And BooVol, IIf(BBooPercents, 2, 1), BLngGra)
        Let l = BLngCur + IIf(BEnuTypeFeedPer = eqlBalBoth And BooVol, IIf(BBooPercents, -2, -1), 0)
        With BRngBox.Range(StrAux0)
          For i = 0 To xRow
            If i < xRow Then
              Let VecAux0(i, j) = "=(" & .Offset(i, 0).Address(False, False) & "*" & .Offset(i, k + j).Address(False, False) & "*100)/" & .Offset(xRow, 0).Address(False, False)
            Else
              Let VecAux0(i, j) = "=SUM(" & .Offset(0, l + j).Resize(xRow, 1).Address(False, False) & ")"
            End If
          Next i
        End With
      Next j
        
  Else
    
    'Prepare grades and units
    If BBooCx Then 'Just Cx
      Let VecAux1 = BVarGradesCx
      Let VecAux2 = RES.ArraySlice(BVarUnits, UBound(BVarGrades, 2))
    Else
      If BAllResult Then 'Concat both matrices
        Let VecAux1 = RES.ArrayConcat(BVarGrades, BVarGradesCx, RByCols:=True)
        Let VecAux2 = BVarUnits
      Else 'Just grades of normal balmet
        Let VecAux1 = BVarGrades
        Let VecAux2 = RES.ArraySlice(BVarUnits, UBound(BVarGrades, 2))
        Let VecAux2 = BVarUnits
      End If
    End If
    If Not VBA.IsArray(VecAux1) Or Not VBA.IsArray(VecAux2) Then Call EHGLOBAL.ErrorHandlerSet(6.1, "Fallo al generar matriz de Unidades."): GoTo EH 'Error
    'Obtain Percents
    If Not Balmet_core_main_app_spreadbalmet_get_percents(BVarContents, BVarContents, RES) Then GoTo EH
    If VBA.IsArray(BVarVolumes) Then
      If Not Balmet_core_main_app_spreadbalmet_get_volumes_percents(BVarVolumes, BVarVolumes, RES) Then GoTo EH
    End If
    'Obtain U
    Let xRow = (UBound(VecAux1) - LBound(VecAux1))
    Let xCol = (UBound(VecAux1, 2) - LBound(VecAux1, 2))
    Let DblAux0 = 0
    Let DblAux1 = 0
    ReDim VecAux0(0 To xRow, 0 To xCol)
    For j = LBound(VecAux2) To UBound(VecAux2) 'Units
      For i = LBound(VecAux1) To UBound(VecAux1) 'Grades
        If VBA.LCase(VBA.CStr(VecAux2(j))) = "g/l" Then
          Let VecAux0(i, j) = (VBA.CDbl(BVarVolumes(i, 0)) * VBA.CDbl(VecAux1(i, j)))
        Else
          Let VecAux0(i, j) = (VBA.CDbl(BVarContents(i, 0)) * VBA.CDbl(VecAux1(i, j)))
        End If
      Next i
    Next j
    Let VecAux1 = Empty
    Let VecAux2 = Empty
    Let VecAux2 = RES.ArrayPop(VecAux0, True)
    If Not VBA.IsArray(VecAux2) Then Call EHGLOBAL.ErrorHandlerSet(6.1, "Fallo al generar matriz de Unidades."): GoTo EH 'Error
    Let DblAux0 = 0
    For j = LBound(VecAux0, 2) To UBound(VecAux0, 2)
      For i = LBound(VecAux0) To UBound(VecAux0)
        If i < UBound(VecAux0) Then
          Let DblAux0 = DblAux0 + VBA.CDbl(VecAux0(i, j))
        Else
          Let VecAux0(i, j) = DblAux0
        End If
      Next i
      Let DblAux0 = 0
    Next j
  
  End If
  
  'Add Headers
  If BBooWithHeader Then
    'Fill arr with names
    For i = 0 To UBound(BVarNames)
      Let BVarNames(i) = "Unidad " & VBA.CStr(BVarNames(i))
    Next i
    Let VecAux0 = RES.ArrayAddAtFirst(VecAux0, BVarNames)
    If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(6.2, "Fallo al crear encabezados de Unidades."): GoTo EH 'Error
  End If
  
  'Calculate Heads
  If BBooHeads Then
    'Method calls two other methods named calculated and tested
    If BBooFormulas Then
    
      ReDim VecAux1(0 To 1, 0 To xCol)
      For i = 0 To xCol
        Let BooVol = (VBA.LCase(VBA.CStr(BVarUnits(i))) = "g/l")
        Let StrAux0 = IIf(BooVol, BStrVol, "C3")
        Let k = IIf(BEnuTypeFeedPer = eqlBalBoth And BooVol, IIf(BBooPercents, 2, 1), BLngGra)
        Let l = BLngCur + IIf(BEnuTypeFeedPer = eqlBalBoth And BooVol, IIf(BBooPercents, -2, -1), 0)
        With BRngBox.Range(StrAux0)
          Let VecAux1(0, i) = "=(" & .Offset(xRow + 1, k + i).Address(False, False) & "*100)"
          Let VecAux1(1, i) = "=SUM(" & .Offset(0, l + i).Resize(xRow, 1).Address(False, False) & ")"
        End With
      Next i
      Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1)
      If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(6.3, "Fallo al determinar Cabezas ensayada y calculada en Unidades."): GoTo EH 'Error
      Let VecAux1 = Empty
    
    Else
      
      'Obtain U
      Let DblAux0 = 0
      ReDim VecAux1(0 To 1, 0 To UBound(VecAux0, 2))
      For j = LBound(VecAux1, 2) To UBound(VecAux1, 2)
        For i = LBound(VecAux1) To UBound(VecAux1)
          If i < UBound(VecAux1) Then
            Let VecAux1(i, j) = VecAux2(0, j)
          Else
            Let VecAux1(i, j) = VecAux0(UBound(VecAux0), j)
          End If
        Next i
      Next j
      
      'Fill final vector with calculated grades
      Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1)
      If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(6.3, "Fallo al crear valores de cabeza ensayada y calculada en Unidades."): GoTo EH 'Error
    
    End If
  End If

  'Return
  Let BVarResponse = VecAux0
  Let VecAux0 = Empty
  Let Balmet_core_main_app_spreadbalmet_get_unities = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Let VecAux2 = Empty
  Let VecAux3 = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Balmet_core_main_app_spreadbalmet_get_fines( _
  ByVal BVarContents As Variant, _
  ByVal BVarVolumes As Variant, _
  ByVal BVarGrades As Variant, _
  ByVal BVarGradesCx As Variant, _
  ByVal BVarUnits As Variant, _
  ByVal BVarNames As Variant, _
  ByRef BVarResponse As Variant, _
  ByVal BAllResult As Boolean, _
  ByVal RES As AppResources_current, _
  Optional BBooFormulas As Boolean = False, _
  Optional BBooWithHeader As Boolean = False, _
  Optional BBooHeads As Boolean = False, _
  Optional BRngBox As Range, _
  Optional BBooCx As Boolean = False, _
  Optional BBooGr As Boolean = False, _
  Optional BBooOT As Boolean = False, _
  Optional BBooPercents As Boolean = False, _
  Optional BEnuTypeFeedPer As EQLBAL_ENU_TYPECONTENT = eqlBalSolids, _
  Optional BStrVol As String = VBA.vbNullString, _
  Optional BLngGra As Long = 0, _
  Optional BLngCur As Long = 0) As Boolean
  
  'Calculates the all Fines of main balmet
  Dim VecAux0 As Variant, VecAux1 As Variant, VecAux2 As Variant, VecAux3 As Variant
  Dim StrAux0 As String, StrAux1 As String
  Dim BooVol As Boolean
  Dim DblAux0 As Double
  Dim i As Long, j As Long, k As Long, l As Long, xRow As Long, xCol As Long

  Const sFN As String = "APPBALMET::-core_spreadbalmet_get_fines"

  On Error GoTo EH
  Let Balmet_core_main_app_spreadbalmet_get_fines = False
  
  'Calculate main results
  Let xRow = (UBound(BVarGrades) - LBound(BVarGrades))
  Let xCol = UBound(BVarNames)
  If BBooFormulas Then
  
    ReDim VecAux0(0 To xRow, 0 To xCol)
    For j = 0 To xCol
      Let BooVol = (VBA.LCase(VBA.CStr(BVarUnits(j))) = "g/l")
      Let StrAux0 = IIf(BooVol, BStrVol, "C3")
      Let k = IIf(BEnuTypeFeedPer = eqlBalBoth And BooVol, IIf(BBooPercents, 2, 1), BLngGra)
      Let l = BLngCur + IIf(BEnuTypeFeedPer = eqlBalBoth And BooVol, IIf(BBooPercents, -2, -1), 0)
      Let StrAux1 = VBA.LCase(VBA.CStr(BVarUnits(j)))
      Select Case StrAux1
        Case "%": Let StrAux1 = "/100)" & IIf(BBooGr, "*(1000^2)", VBA.vbNullString)
        Case "dm": Let StrAux1 = "*100)" & IIf(BBooGr, ")", ")/(1000^2)")
        Case "g/t", "g/l": Let StrAux1 = IIf(BBooGr, ")", ")/(1000^2)")
      End Select
      With BRngBox.Range(StrAux0)
        For i = 0 To xRow
          If i < xRow Then
            Let VecAux0(i, j) = "=(" & .Offset(i, 0).Address(False, False) & "*" & .Offset(i, k + j).Address(False, False) & StrAux1
          Else
            Let VecAux0(i, j) = "=SUM(" & .Offset(0, l + j).Resize(xRow, 1).Address(False, False) & ")"
          End If
        Next i
      End With
    Next j

  Else
    
    'Prepare grades and units
    If BBooCx Then 'Just Cx
      Let VecAux1 = BVarGradesCx
      Let VecAux2 = RES.ArraySlice(BVarUnits, UBound(BVarGrades, 2))
      'Let VecAux2 = BVarUnits
    Else
      If BAllResult Then 'Concat both matrices
        Let VecAux1 = RES.ArrayConcat(BVarGrades, BVarGradesCx, RByCols:=True)
        Let VecAux2 = BVarUnits
      Else 'Just grades of normal balmet
        Let VecAux1 = BVarGrades
        Let VecAux2 = RES.ArraySlice(BVarUnits, UBound(BVarGrades, 2))
        Let VecAux2 = BVarUnits
      End If
    End If
    If Not VBA.IsArray(VecAux1) Or Not VBA.IsArray(VecAux2) Then Call EHGLOBAL.ErrorHandlerSet(7.1, "Fallo al generar matriz de Finos."): GoTo EH 'Error
    'Grades conversion
    For j = LBound(VecAux2) To UBound(VecAux2) 'Units
      For i = LBound(VecAux1) To UBound(VecAux1) 'Grades
        Let StrAux0 = VBA.LCase(VBA.CStr(VecAux2(j)))
        Select Case StrAux0
          Case "%": Let VecAux1(i, j) = VBA.CDbl(VecAux1(i, j)) / 100
          Case "dm": Let VecAux1(i, j) = (VBA.CDbl(VecAux1(i, j)) * 100) / (1000 ^ 2)
          Case "g/t": Let VecAux1(i, j) = (VBA.CDbl(VecAux1(i, j)) / (1000 ^ 2))
          Case "g/l": Let VecAux1(i, j) = (VBA.CDbl(VecAux1(i, j)) / (1000 ^ 2))
        End Select
        If BBooGr Then Let VecAux1(i, j) = VBA.CDbl(VecAux1(i, j)) * (1000 ^ 2)
      Next i
    Next j
    'Obtain F
    Let xRow = (UBound(VecAux1) - LBound(VecAux1))
    Let xCol = (UBound(VecAux1, 2) - LBound(VecAux1, 2))
    ReDim VecAux0(0 To xRow, 0 To xCol)
    For j = LBound(VecAux2) To UBound(VecAux2) 'Units
      For i = LBound(VecAux1) To UBound(VecAux1) 'Grades
        If VBA.LCase(VBA.CStr(VecAux2(j))) = "g/l" Then
          Let VecAux0(i, j) = (VBA.CDbl(BVarVolumes(i, 0)) * VBA.CDbl(VecAux1(i, j)))
        Else
          Let VecAux0(i, j) = (VBA.CDbl(BVarContents(i, 0)) * VBA.CDbl(VecAux1(i, j)))
        End If
      Next i
    Next j
    Let VecAux1 = Empty
    Let VecAux2 = Empty
    Let VecAux2 = RES.ArrayPop(VecAux0, True)
    If Not VBA.IsArray(VecAux2) Then Call EHGLOBAL.ErrorHandlerSet(7.1, "Fallo al generar matriz de Finos."): GoTo EH 'Error
    Let DblAux0 = 0
    For j = LBound(VecAux0, 2) To UBound(VecAux0, 2)
      For i = LBound(VecAux0) To UBound(VecAux0)
        If i < UBound(VecAux0) Then
          Let DblAux0 = DblAux0 + VBA.CDbl(VecAux0(i, j))
        Else
          Let VecAux0(i, j) = DblAux0
        End If
      Next i
    Next j
  
  End If
  
  'Add Headers
  If BBooWithHeader Then
    'Fill arr with names
    Let VecAux1 = BVarNames
    For i = 0 To UBound(BVarNames)
      Let VecAux1(i) = "Fino " & VBA.CStr(VecAux1(i)) & IIf(BBooGr, "[g]", "[T]")
    Next i
    Let VecAux0 = RES.ArrayAddAtFirst(VecAux0, VecAux1)
    If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(7.2, "Fallo al crear encabezados de Finos."): GoTo EH 'Error
    Let VecAux1 = Empty
  End If
  
  'Calculate Heads
  If BBooHeads Then
    'Method calls two other methods named calculated and tested
    If BBooFormulas Then
    
      ReDim VecAux1(0 To 1, 0 To xCol)
      For i = 0 To xCol
        Let BooVol = (VBA.LCase(VBA.CStr(BVarUnits(i))) = "g/l")
        Let StrAux0 = IIf(BooVol, BStrVol, "C3")
        Let k = IIf(BEnuTypeFeedPer = eqlBalBoth And BooVol, IIf(BBooPercents, 2, 1), BLngGra)
        Let l = BLngCur + IIf(BEnuTypeFeedPer = eqlBalBoth And BooVol, IIf(BBooPercents, -2, -1), 0)
        Let StrAux1 = VBA.LCase(VBA.CStr(BVarUnits(i)))
        Select Case StrAux1
          Case "%": Let StrAux1 = "/100)" & IIf(BBooGr, "*(1000^2)", VBA.vbNullString)
          Case "dm": Let StrAux1 = "*100)" & IIf(BBooGr, ")", ")/(1000^2)")
          Case "g/t", "g/l": Let StrAux1 = IIf(BBooGr, ")", ")/(1000^2)")
        End Select
        With BRngBox.Range(StrAux0)
          Let VecAux1(0, i) = "=(" & .Offset(xRow + 1, 0).Address(False, False) & "*" & .Offset(xRow + 1, k + i).Address(False, False) & StrAux1
          Let VecAux1(1, i) = "=SUM(" & .Offset(0, l + i).Resize(xRow, 1).Address(False, False) & ")"
        End With
      Next i
      Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1)
      If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(7.3, "Fallo al determinar Cabezas ensayada y calculada en Finos."): GoTo EH 'Error
      Let VecAux1 = Empty
    
    Else
      
      'Obtain U
      Let DblAux0 = 0
      ReDim VecAux1(0 To 1, 0 To UBound(VecAux0, 2))
      For j = LBound(VecAux1, 2) To UBound(VecAux1, 2)
        For i = LBound(VecAux1) To UBound(VecAux1)
          If i < UBound(VecAux1) Then
            Let VecAux1(i, j) = VecAux2(0, j)
          Else
            Let VecAux1(i, j) = VecAux0(UBound(VecAux0), j)
          End If
        Next i
      Next j

      'Fill final vector with calculated grades
      Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1)
      If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(7.3, "Fallo al crear valores de cabeza ensayada y calculada en Finos."): GoTo EH 'Error
    
    End If
  End If
  
  'If user requires gr and OT in balmet
  If BBooOT Then
    If BBooFormulas Then
      
      Let VecAux2 = VecAux0
      Let VecAux1 = Empty
      Let k = 0
      For j = 0 To xCol
        Let BooVol = (VBA.LCase(VBA.CStr(BVarUnits(j))) = "g/l")
        Let StrAux0 = IIf(BooVol, BStrVol, "C3")
        Let l = BLngCur + IIf(BEnuTypeFeedPer = eqlBalBoth And BooVol, IIf(BBooPercents, -2, -1), 0) + j
        Let StrAux1 = IIf(BBooGr, VBA.vbNullString, "*1000")
        With BRngBox.Range(StrAux0)
          Let k = k + j
          Let StrAux0 = VBA.LCase(VBA.CStr(BVarNames(j)))
          If StrAux0 = "ag" Or StrAux0 = "au" Then
            ReDim VecAux1(0 To UBound(VecAux0)) '0,0: Header, pre-last: Head-tested, last: Head-calculated
            For i = 0 To UBound(VecAux1)
              If i = 0 Then
                Let VecAux1(i) = "Fino " & VBA.CStr(BVarNames(j)) & "[OT]"
              ElseIf i <= xRow Or i = (xRow + 1) Then
                Let VecAux1(i) = "=(" & .Offset(i, l).Address(False, False) & StrAux1 & ")/31.1035"
              Else
                Let VecAux1(i) = "=SUM(" & .Offset(0, l + j).Resize(xRow, 1).Address(False, False) & ")"
              End If
            Next i
            Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1, k + 1, True)
            If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(7.4, "Fallo al crear valores de columna Finos [OT]."): GoTo EH 'Error
            Let k = k + 1
            Let VecAux1 = Empty
          End If
        End With
      Next j
    
    Else
      
      'Iterate in all elements of Names
      'If exists Ag or Au and if user requires grams or OT, calculate its values
      'Once calculated, concat the resultant matrix into auxvec0 at the Ag or Au position
      'Headers in this case, must be like: Fine Ag[OT]
      Let VecAux2 = VecAux0
      Let VecAux1 = Empty
      Let k = 0
      For j = LBound(BVarNames) To UBound(BVarNames)
        Let StrAux0 = VBA.LCase(VBA.CStr(BVarNames(j)))
        Let k = k + j
        If StrAux0 = "ag" Or StrAux0 = "au" Then
          ReDim VecAux1(0 To UBound(VecAux0), 0 To 0) '0,0: Header, pre-last: Head-tested, last: Head-calculated
          Let StrAux0 = "Fino " & VBA.CStr(BVarNames(j)) & "[OT]"
          Let VecAux1(0, 0) = "Fino " & VBA.CStr(BVarNames(j)) & "[OT]"
          For i = 1 To UBound(VecAux1)
            Let VecAux1(i, 0) = (VBA.CDbl(VecAux2(i, j)) / IIf(BBooGr, 1, (1000 ^ 2))) / 31.1035 'Conversion to OT
          Next i
          Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1, k + 1, True)
          If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(7.4, "Fallo al crear valores de columna Finos [OT]."): GoTo EH 'Error
          Let k = k + 1
          Let VecAux1 = Empty
        End If
      Next j
    
    End If
  End If

  'Return
  Let BVarResponse = VecAux0
  Let VecAux0 = Empty
  Let Balmet_core_main_app_spreadbalmet_get_fines = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Let VecAux2 = Empty
  Let VecAux3 = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Balmet_core_main_app_spreadbalmet_get_recoveries( _
  ByVal BVarContents As Variant, _
  ByVal BVarVolumes As Variant, _
  ByVal BVarGrades As Variant, _
  ByVal BVarGradesCx As Variant, _
  ByVal BVarUnits As Variant, _
  ByVal BVarNames As Variant, _
  ByRef BVarResponse As Variant, _
  ByVal BAllResult As Boolean, _
  ByVal RES As AppResources_current, _
  Optional BBooFormulas As Boolean = False, _
  Optional BBooWithHeader As Boolean = False, _
  Optional BBooHeads As Boolean = False, _
  Optional BRngBox As Range, _
  Optional BBooCx As Boolean = False, _
  Optional BBooPercents As Boolean = False, _
  Optional BEnuTypeFeedPer As EQLBAL_ENU_TYPECONTENT = eqlBalSolids, _
  Optional BStrVol As String = VBA.vbNullString, _
  Optional BLngGra As Long = 0, _
  Optional BLngCur As Long = 0) As Boolean

  'Data required: (BVarGrades, BVarNames)
  Dim VecAux0 As Variant, VecAux1 As Variant, VecAux2 As Variant, VecAux3 As Variant
  Dim StrAux0 As String, StrAux1 As String
  Dim BooVol As Boolean
  Dim DblAux0 As Double
  Dim i As Long, j As Long, k As Long, l As Long, xRow As Long, xCol As Long

  Const sFN As String = "APPBALMET::-core_spreadbalmet_get_recoveries"

  On Error GoTo EH
  Let Balmet_core_main_app_spreadbalmet_get_recoveries = False
  
  'Calculate main results
  Let xRow = (UBound(BVarGrades) - LBound(BVarGrades))
  Let xCol = UBound(BVarNames)
  If BBooFormulas Then
  
      ReDim VecAux0(0 To xRow, 0 To xCol)
      For j = 0 To xCol
        Let BooVol = (VBA.LCase(VBA.CStr(BVarUnits(j))) = "g/l")
        Let StrAux0 = IIf(BooVol, BStrVol, "C3")
        Let k = IIf(BEnuTypeFeedPer = eqlBalBoth And BooVol, IIf(BBooPercents, 2, 1), BLngGra)
        Let l = BLngCur + IIf(BEnuTypeFeedPer = eqlBalBoth And BooVol, IIf(BBooPercents, -2, -1), 0)
        With BRngBox.Range(StrAux0)
          ReDim VecAux1(xRow - 1)
          For i = 0 To (xRow - 1)
            Let VecAux1(i) = .Offset(i, 0).Address(False, False) & "*" & .Offset(i, k + j).Address(False, False)
          Next i
          Let StrAux1 = "SUM(" & VBA.Join(VecAux1, ";") & ")"
          Let VecAux1 = Empty
          For i = 0 To xRow
            If i < xRow Then
              Let VecAux0(i, j) = "=((" & .Offset(i, 0).Address(False, False) & "*" & .Offset(i, k + j).Address(False, False) & ")/" & StrAux1 & ")*100"
            Else
              Let VecAux0(i, j) = "=SUM(" & .Offset(0, l + j).Resize(xRow, 1).Address(False, False) & ")"
            End If
          Next i
        End With
      Next j
      Let VecAux1 = Empty
  
  Else
    
    'Obtain U
    If Not Balmet_core_main_app_spreadbalmet_get_unities(BVarContents, BVarVolumes, BVarGrades, BVarGradesCx, BVarUnits, BVarNames, VecAux0, BAllResult, RES, BBooCx:=BBooCx) Then GoTo EH
    'Obtain %R
    For j = LBound(VecAux0, 2) To UBound(VecAux0, 2)
      Let DblAux0 = VBA.CDbl(VecAux0(UBound(VecAux0), j))
      For i = LBound(VecAux0) To UBound(VecAux0)
        Let VecAux0(i, j) = (VBA.CDbl(VecAux0(i, j)) / DblAux0) * 100
      Next i
    Next j
  
  End If
  
  'Add Headers
  If BBooWithHeader Then
    'Fill arr with names
    Let VecAux1 = BVarNames
    For i = 0 To UBound(BVarNames)
      Let VecAux1(i) = "Rec. " & VBA.CStr(VecAux1(i)) & "[%]"
    Next i
    Let VecAux0 = RES.ArrayAddAtFirst(VecAux0, VecAux1)
    If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(8.1, "Fallo al crear encabezados de Recuperaciones."): GoTo EH 'Error
    Let VecAux1 = Empty
  End If
  
  'Calculate Heads
  If BBooHeads Then
    'Method calls two other methods named calculated and tested
    If BBooFormulas Then
    
      ReDim VecAux2(0 To 1, 0 To xCol)
      For j = 0 To xCol
        Let BooVol = (VBA.LCase(VBA.CStr(BVarUnits(j))) = "g/l")
        Let StrAux0 = IIf(BooVol, BStrVol, "C3")
        Let k = IIf(BEnuTypeFeedPer = eqlBalBoth And BooVol, IIf(BBooPercents, 2, 1), BLngGra)
        Let l = BLngCur + IIf(BEnuTypeFeedPer = eqlBalBoth And BooVol, IIf(BBooPercents, -2, -1), 0)
        With BRngBox.Range(StrAux0)
          ReDim VecAux1(xRow - 1)
          For i = 0 To (xRow - 1)
            Let VecAux1(i) = .Offset(i, 0).Address(False, False) & "*" & .Offset(i, k + j).Address(False, False)
          Next i
          Let StrAux1 = "SUM(" & VBA.Join(VecAux1, ";") & ")"
          Let VecAux1 = Empty
          Let VecAux2(0, j) = "=(" & StrAux1 & "/(" & .Offset(xRow, 0).Address(False, False) & "*" & .Offset(xRow, k + j).Address(False, False) & "))*100"
          Let VecAux2(1, j) = "=SUM(" & .Offset(0, l + j).Resize(xRow, 1).Address(False, False) & ")"
        End With
      Next j
      Let VecAux1 = Empty
      'Fill final vector with calculated recoveries
      Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux2)
      If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(8.2, "Fallo al crear valores de cabeza ensayada y calculada en Recuperaciones."): GoTo EH 'Error
      Let VecAux2 = Empty
    
    Else
      
      'Obtain U
      Let DblAux0 = 0
      ReDim VecAux1(0 To 1, 0 To UBound(VecAux0, 2))
      For j = LBound(VecAux0, 2) To UBound(VecAux0, 2) 'Unities
        For i = 1 To UBound(VecAux0)
          If i < UBound(VecAux0) Then
            Let DblAux0 = DblAux0 + (VBA.CDbl(VecAux0(i, j)))
          Else
            Let VecAux1(0, j) = VecAux0(i, j)
            Let VecAux1(1, j) = DblAux0
          End If
        Next i
        Let DblAux0 = 0
      Next j
      'Fill final vector with calculated grades
      Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1)
      If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(8.2, "Fallo al crear valores de cabeza ensayada y calculada en Recuperaciones."): GoTo EH 'Error
      Let VecAux1 = Empty
    
    End If
  End If

  'Return
  Let BVarResponse = VecAux0
  Let VecAux0 = Empty
  Let Balmet_core_main_app_spreadbalmet_get_recoveries = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Let VecAux2 = Empty
  Let VecAux3 = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Balmet_core_main_app_spreadbalmet_get_ratio( _
  ByVal BVarContents As Variant, _
  ByVal BVarVolumes As Variant, _
  ByRef BVarResponse As Variant, _
  ByVal RES As AppResources_current, _
  Optional BBooFormulas As Boolean = False, _
  Optional BBooWithHeader As Boolean = False, _
  Optional BBooHeads As Boolean = False, _
  Optional BRngBox As Range, _
  Optional BStrVol As String = VBA.vbNullString) As Boolean

  'Data required: (BVarGrades, BVarNames)
  Dim VecAux0 As Variant, VecAux1 As Variant
  Dim StrAux0 As String
  Dim DblAux0 As Double
  Dim i As Long, j As Long, xRow As Long, xCol As Long

  Const sFN As String = "APPBALMET::-core_spreadbalmet_get_ratio"

  On Error GoTo EH
  Let Balmet_core_main_app_spreadbalmet_get_ratio = False
  
  'Calculate main results
  Let VecAux1 = RES.ArrayConcat(BVarContents, BVarVolumes, RByCols:=True)
  If Not VBA.IsArray(VecAux1) Then Call EHGLOBAL.ErrorHandlerSet(9.1, "Fallo al crear resultados de Ratio."): GoTo EH 'Error
  Let xRow = (UBound(VecAux1) - LBound(VecAux1))
  Let xCol = (UBound(VecAux1, 2) - LBound(VecAux1, 2))
  ReDim VecAux0(0 To xRow, 0 To xCol)
  If BBooFormulas Then
  
      For j = 0 To xCol
        Let StrAux0 = IIf(j > 0, BStrVol, "C3")
        With BRngBox.Range(StrAux0)
          For i = 0 To xRow
            Let VecAux0(i, j) = "=" & .Offset(i, 0).Address(False, False) & "/" & .Offset(xRow, 0).Address(False, False)
          Next i
        End With
      Next j
  
  Else
    
    'Ratio
    For j = LBound(VecAux0, 2) To UBound(VecAux0, 2)
      Let DblAux0 = VBA.CDbl(VecAux1(UBound(VecAux1), j))
      For i = LBound(VecAux0) To UBound(VecAux0)
        Let VecAux0(i, j) = DblAux0 / VBA.CDbl(VecAux1(i, j))
      Next i
    Next j
  
  End If
  Let VecAux1 = Empty
  
  'Add Headers
  If BBooWithHeader Then
    'Fill arr with names
    ReDim VecAux1(0 To 0, 0 To xCol)
    Let VecAux1(0, 0) = "Ratio" & IIf(xCol > 0, " Sólido", VBA.vbNullString)
    If xCol > 0 Then Let VecAux1(0, 1) = "Ratio Volumen"
    Let VecAux0 = RES.ArrayAddAtFirst(VecAux0, VecAux1)
    If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(9.2, "Fallo al crear encabezados de Ratio."): GoTo EH 'Error
    Let VecAux1 = Empty
  End If
  
  'Calculate Heads
  If BBooHeads Then
    'Method calls two other methods named calculated and tested
    If BBooFormulas Then
      
      ReDim VecAux1(0 To 1, 0 To xCol)
      For j = 0 To xCol
        Let StrAux0 = IIf(j > 0, BStrVol, "C3")
        With BRngBox.Range(StrAux0)
          Let VecAux1(0, j) = "=" & .Offset(xRow + 1, 0).Address(False, False) & "/" & .Offset(xRow, 0).Address(False, False)
          Let VecAux1(1, j) = "=" & .Offset(xRow + 2, 0).Address(False, False) & "/" & .Offset(xRow, 0).Address(False, False)
        End With
      Next j
      'Fill final vector with calculated grades
      Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1)
      If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(9.3, "Fallo al crear valores de cabeza ensayada y calculada en Ratio."): GoTo EH 'Error
      Let VecAux1 = Empty
    
    Else
      
      'Obtain U
      ReDim VecAux1(0 To 1, 0 To UBound(VecAux0, 2))
      Let VecAux1(0, 0) = VecAux0(UBound(VecAux0), 0)
      Let VecAux1(1, 0) = VecAux0(UBound(VecAux0), 0)
      If xCol > 0 Then
        Let VecAux1(0, 1) = VecAux0(UBound(VecAux0), 1)
        Let VecAux1(1, 1) = VecAux0(UBound(VecAux0), 1)
      End If
      'Fill final vector with calculated grades
      Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1)
      If Not VBA.IsArray(VecAux0) Then Call EHGLOBAL.ErrorHandlerSet(9.3, "Fallo al crear valores de cabeza ensayada y calculada en Ratio."): GoTo EH 'Error
      Let VecAux1 = Empty
    
    End If
  End If

  'Return
  Let BVarResponse = VecAux0
  Let VecAux0 = Empty
  Let Balmet_core_main_app_spreadbalmet_get_ratio = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function





' ASSETS ----------------------------------------------------------------------------'
'Private Function Balmet_assets_get_products_bounds_OOLD( _
'  ByVal BLngSource As Long, _
'  Optional BLngSourceCx As Long = 0, _
'  Optional BWhat As EQLBAL_ENU_BOUNDS = eqlBalProducts) As Long
'
''  'Returns the number of Products, Cols, Rows, TotalProducts and TotalComplexProducts of any Balmet
''  Dim LngProd As Long, LngCol As Long, LngRow As Long, LngTotProd As Long, LngTotCx As Long, LngAux0 As Long
''  Dim BooFound As Boolean, BooTimeup As Boolean
''
''  Const LngLASTBOUND As Long = 10
''  Const sFN As String = "APPBALMET::-assets_get_products_bounds"
''
''  On Error GoTo EH
''  Let Balmet_assets_get_products_bounds = 0
''
''  'Determine if the Source number is a valid quantity of elements to be a Balmet
'''  Public Enum EQLBAL_ENU_BOUNDS
'''    eqlBalProducts
'''    eqlBalColumns
'''    eqlBalRows
'''    eqlBalFullProducts
'''    eqlBalFullProductsCx
'''    eqlBalProductsCx
'''    eqlBalColumnsCx
'''  End Enum
''
''  Let BooFound = False
''  Let BooTimeup = False
''  Let LngProd = 2
''  Do While (Not BooFound And Not BooTimeup)
''    Let LngCol = LngProd - 1
''    Let LngRow = (LngProd + 1)
''    Let LngTotProd = LngCol * LngRow
''
''    Let LngProd = LngProd + 1
''    Let BooFound = (BLngSource = LngTotProd)
''    Let BooTimeup = (LngProd = LngLASTBOUND)
''    'Exit Do
''  Loop
''  If Not BooFound Then
''    GoTo EH
''  Else 'Determine Cx elements
''    If BLngSourceCx > 0 Then
''      Let LngTotCx = IIf(BLngSourceCx = LngRow Or BLngSourceCx = (LngRow * 2) Or BLngSourceCx = (LngRow * 3), BLngSourceCx + LngTotProd, 0)
''      If LngTotCx = 0 Then GoTo EH
''    End If
''  End If
''
''  'Return
''  Select Case BWhat
''    Case eqlBalProducts: Let Balmet_assets_get_products_bounds = LngProd - 1 'Number of Balmet Products
''    Case eqlBalColumns: Let Balmet_assets_get_products_bounds = LngCol   'Number of normal Balmet Cols
''    Case eqlBalRows: Let Balmet_assets_get_products_bounds = LngRow 'Number of Balmet Rows
''    Case eqlBalFullProducts: Let Balmet_assets_get_products_bounds = LngTotProd 'Number of Full normal Balmet grades. Ej. 8 grades to Balmet of 3 Prods
''    Case eqlBalFullProductsCx: Let Balmet_assets_get_products_bounds = LngTotCx 'Number of Full normal and complex Balmet grades. Ej. 12 grades to Balmet of 3 Prods
''    Case eqlBalProductsCx: Let Balmet_assets_get_products_bounds = LngTotCx - LngTotProd 'Number of Full complex Balmet grades. Ej. 12-8 grades = 4 to Balmet of 3 Prods
''    Case eqlBalColumnsCx
''      Let LngAux0 = LngTotCx - LngTotProd
''      Let Balmet_assets_get_products_bounds = LngAux0 / LngRow 'Number of complex Balmet columns.
''    Case Else: GoTo EH
''  End Select
''
''EH:
''  Call EHGLOBAL.ErrorHandlerRaise(sFN)
'End Function

'Private Function Balmet_assets_get_products_bounds_fulldata_OOLD( _
'  ByVal BLngSource As Long, _
'  ByRef BLngProducts As Long, _
'  ByRef BLngColumns As Long, _
'  ByRef BLngRows As Long, _
'  ByRef BLngElements As Long, _
'  ByRef BLngElementsAndCx As Long, _
'  ByRef BLngElementsCx As Long, _
'  ByRef BLngColumnsCx As Long, _
'  Optional BLngSourceCx As Long = 0) As Boolean
'
''  'Returns the number of Products, Cols, Rows, TotalProducts and TotalComplexProducts of any Balmet
''  Dim LngProd As Long, LngCol As Long, LngRow As Long, LngTotProd As Long, LngTotCx As Long
''  Dim BooFound As Boolean, BooTimeup As Boolean
''
''  Const LngLASTBOUND As Long = 10
''  Const sFN As String = "APPBALMET::-assets_get_products_bounds_fulldata"
''
''  On Error GoTo EH
''  Let Balmet_assets_get_products_bounds_fulldata = False
''
''  'Determine if the Source number is a valid value of elements to be a Balmet
''  Let BooFound = False
''  Let BooTimeup = False
''  Let LngProd = 2
''  Do While (Not BooFound And Not BooTimeup)
''    Let LngCol = LngProd - 1
''    Let LngRow = (LngProd + 1)
''    Let LngTotProd = LngCol * LngRow
''
''    Let LngProd = LngProd + 1
''    Let BooFound = (BLngSource = LngTotProd)
''    Let BooTimeup = (LngProd = LngLASTBOUND)
''    'Exit Do
''  Loop
''  If Not BooFound Then
''    GoTo EH
''  Else 'Determine Cx elements (just 3 columns of Cx allowed)
''    If BLngSourceCx > 0 Then
''      Let LngTotCx = IIf(BLngSourceCx = LngRow Or BLngSourceCx = (LngRow * 2) Or BLngSourceCx = (LngRow * 3), BLngSourceCx + LngTotProd, 0)
''      If LngTotCx = 0 Then GoTo EH
''    End If
''  End If
''
''  'Return
''  Let BLngProducts = LngProd - 1 'Number of Balmet Products
''  Let BLngColumns = LngCol 'Number of normal Balmet Cols
''  Let BLngRows = LngRow 'Number of Balmet Rows
''  Let BLngElements = LngTotProd 'Number of Full normal Balmet grades. Ej. 8 grades to Balmet of 3 Prods
''  Let BLngElementsAndCx = IIf(LngTotCx = 0, LngTotProd, LngTotCx) 'Number of Full normal and complex Balmet grades. Ej. 12 grades to Balmet of 3 Prods
''  Let BLngElementsCx = BLngElementsAndCx - LngTotProd 'Number of Full complex Balmet grades. Ej. 12-8 grades = 4 to Balmet of 3 Prods
''  Let BLngColumnsCx = (BLngElementsAndCx - LngTotProd) / LngRow 'Number of complex Balmet columns.
''  'Let BLngElementsCx = LngTotCx - LngTotProd 'Number of Full complex Balmet grades. Ej. 12-8 grades = 4 to Balmet of 3 Prods
''  'Let BLngColumnsCx = (LngTotCx - LngTotProd) / LngRow 'Number of complex Balmet columns.
''
''  Let Balmet_assets_get_products_bounds_fulldata = True
''
''EH:
''  Call EHGLOBAL.ErrorHandlerRaise(sFN)
'End Function

'Private Function Balmet_assets_get_type_OOLD(ByVal BLngSource As Long, Optional BLngSourceCx As Long = 0) As EQLBAL_ENU_TYPE
'  'Return "normal" or "complex" type of any Balmet
'
''  Dim i As Long, j As Long, k As Long, l As Long, M As Long, n As Long, o As Long
''  Dim StrType As EQLBAL_ENU_TYPE
''
''  Const sFN = "APPBALMET::-assets_get_type"
''
''  On Error GoTo EH
''  Let l = 0
''  Let n = 0
''  Let StrType = eqlBalNothing
''
''  If Balmet_assets_get_products_bounds_fulldata(BLngSource, i, j, k, l, M, n, o, BLngSourceCx) Then
''    If n > 0 Then
''      Let StrType = eqlBalWithComplex
''    Else
''      Let StrType = eqlBalNormal
''    End If
''  End If
''
''  'Return
''  Let Balmet_assets_get_type = StrType
''
''EH:
''  Call EHGLOBAL.ErrorHandlerRaise(sFN)
'End Function

'Private Function Balmet_assets_get_is_balmet_OOLD( _
'  ByVal BLngSource As Long, _
'  Optional BLngSourceCx As Long = 0) As Boolean
'  'Return if source is a correct number of elements for any Balmet
'
''  Const sFN As String = "APPBALMET::-assets_get_is_balmet"
''
''  On Error GoTo EH
''
''  'Return
''  Let Balmet_assets_get_is_balmet = Not (Balmet_assets_get_type(BLngSource, BLngSourceCx) = eqlBalNothing)
''
''EH:
''  Call EHGLOBAL.ErrorHandlerRaise(sFN)
'End Function


'Private Function Balmet_assets_get_complex_datas_OOLD( _
'  ByRef BStrSource As String, _
'  ByRef BLngCol As Long, _
'  Optional BBooNamesUnits As Boolean = False, _
'  Optional BBooDelFromSource As Boolean = True) As String
'
''  'Resturn a string of Complex Minerals Names or Units
''  Dim VecSource As Variant, VecOcean As Variant
''  Dim i As Long, j As Long, kDM As Long
''  Dim StrFound As String
''
''  Const sFN As String = "APPBALMET::-assets_get_complex_datas"
''
''  On Error GoTo EH
''  Let Balmet_assets_get_complex_datas = VBA.vbNullString
''  Let kDM = 0
''  Let BLngCol = 0
''  Let StrFound = VBA.vbNullString
''  Let VecSource = VBA.Split(BStrSource, ";")
''  If BBooNamesUnits Then 'Units
''    Let VecOcean = Array("dm", "g/t", "g/l")
''  Else 'Names
''    Let VecOcean = Array("ag", "au", "pt")
''  End If
''  For i = 0 To UBound(VecSource)
''    For j = 0 To UBound(VecOcean)
''      If VBA.LCase(VecSource(i)) = VecOcean(j) Then
''        Let StrFound = VecSource(i)
''        If Not BBooNamesUnits Then
''          If BBooDelFromSource Then
''            Let BStrSource = VBA.Replace(BStrSource, StrFound, VBA.vbNullString, Count:=1)
''            Let BLngCol = BLngCol + 1
''          End If
''        Else
''          If BBooDelFromSource Then
''            If VBA.LCase(StrFound) = "dm" Then
''              Let BStrSource = VBA.Replace(BStrSource, StrFound, VBA.vbNullString, Count:=1)
''              Let kDM = kDM + 1
''            ElseIf i = UBound(VecSource) Then
''              Let BStrSource = VBA.Replace(BStrSource, StrFound, VBA.vbNullString, Count:=1)
''              Let BLngCol = BLngCol + 1
''            End If
''          End If
''        End If
''      End If
''    Next j
''  Next i
''  If Not BBooNamesUnits Then
''    If BLngCol > 1 Or BLngCol = 0 Then Let BLngCol = 0: GoTo EH
''  Else
''    If Not (kDM = 1 And BLngCol = 0) Then
''      If Not (kDM = 0 And BLngCol = 1) Then Let BLngCol = 0: GoTo EH
''    End If
''    'If (BLngCol = 0 And kDM = 0) Or (BLngCol > 1 Or kDM > 1) Or (kDM = BLngCol) Then Let BLngCol = 0: GoTo EH
''  End If
''
''  'Return
''  Let Balmet_assets_get_complex_datas = StrFound
''
''EH:
''  Let VecSource = Empty
''  Let VecOcean = Empty
''  Call EHGLOBAL.ErrorHandlerRaise(sFN)
'End Function

'Private Function Balmet_assets_get_units_OOLD( _
'  ByRef BStrSource As String, _
'  ByRef BLngCol As Long, _
'  Optional BBooDelFromSource As Boolean = True) As String
'
''  'Resturn a string with Units of grades
''  Dim VecSource As Variant, VecFound As Variant, VecOcean As Variant
''  Dim i As Long, j As Long, LngDMCounter As Long
''  Dim StrFound As String, StrAux As String
''
''  Const sFN As String = "APPBALMET::-assets_get_units"
''
''  On Error GoTo EH
''  Let Balmet_assets_get_units = VBA.vbNullString
''
''  Let BLngCol = 0
''  Let LngDMCounter = 0
''  Let StrFound = VBA.vbNullString
''  Let VecSource = VBA.Split(BStrSource, ";")
''  Let VecOcean = Array("dm", "g/t", "g/l", "%")
''  ReDim VecFound(0)
''  For i = 0 To UBound(VecSource)
''    For j = 0 To UBound(VecOcean)
''      If VBA.LCase(VecSource(i)) = VecOcean(j) Then
''        ReDim Preserve VecFound(BLngCol)
''        Let StrAux = VecSource(i)
''        Let VecFound(BLngCol) = StrAux
''        If BBooDelFromSource Then Let BStrSource = VBA.Replace(BStrSource, StrAux, VBA.vbNullString, Count:=1)
''        Let BLngCol = BLngCol + 1
''        Let LngDMCounter = LngDMCounter + IIf(VecOcean(j) = "dm", 1, 0)
''      End If
''    Next j
''  Next i
''  If LngDMCounter > 1 And BLngCol > 0 Then GoTo EH
''
''  'Return
''  Let Balmet_assets_get_units = VBA.Join(VecFound, ";")
''
''EH:
''  Let VecSource = Empty
''  Let VecFound = Empty
''  Let VecOcean = Empty
''  Call EHGLOBAL.ErrorHandlerRaise(sFN)
'End Function

'Private Function Balmet_assets_get_units_refill_OOLD(ByVal BStrUnit As String, ByRef BLngCounter As Long) As Variant
''  Dim VecSource As Variant
''  Dim i As Long
''
''  Const sFN As String = "APPBALMET::-assets_get_units_refill"
''
''  On Error GoTo EH
''  Let Balmet_assets_get_units_refill = Empty
''
''  ReDim VecSource(BLngCounter - 1)
''  For i = 0 To UBound(VecSource)
''    Let VecSource(i) = BStrUnit
''  Next i
''
''  Let BLngCounter = IIf(i > 0, i, 0)
''  Let Balmet_assets_get_units_refill = VecSource
''
''EH:
''  Let VecSource = Empty
''  Call EHGLOBAL.ErrorHandlerRaise(sFN)
'End Function

Private Function Balmet_assets_get_grades( _
  ByVal BVarVilSource As Variant, _
  ByRef BLngCounter As Long, _
  Optional BxChar As String = ";", _
  Optional BResponseDimension As EQLBAL_ENU_GRADESVECTORTYPE = eqlBalAs1D) As Variant

  'Resturn a string with grades of normal Balmet
  Dim RCT As AppResourcesController_current
  
  Dim VecSource As Variant, VecNew As Variant
  Dim i As Long, j As Long, k As Long, l As Long, xRow As Long, xCol As Long
  Dim StrFound As String, StrAux As String

  Const sFN As String = "APPBALMET::-assets_get_grades"

  On Error GoTo EH
  Let BLngCounter = 0
  
  If VBA.TypeName(BVarVilSource) = "String" Then
    Let VecSource = VBA.Split(BVarVilSource, BxChar)
  Else
    Let VecSource = BVarVilSource
  End If
  Let k = UBound(VecSource) + 1
  Set RCT = New AppResourcesController_current
  Set RCT.ErrorHandler = EHGLOBAL
  'Let xRow = Balmet_assets_get_products_bounds(k, 0, eqlBalRows) - 1
  'Let xCol = Balmet_assets_get_products_bounds(k, 0, eqlBalColumns) - 1
  Let xRow = RCT.ResCtrBalGetProductsBounds(k, 0, eqlBalRows) - 1
  Let xCol = RCT.ResCtrBalGetProductsBounds(k, 0, eqlBalColumns) - 1
  Set RCT = Nothing
  
  If BResponseDimension = eqlBalAs1D Then '1D
    ReDim VecNew(UBound(VecSource))
    For i = LBound(VecSource) To UBound(VecSource)
      Let VecNew(i) = VBA.Abs(VBA.CDbl(VecSource(i)))
    Next i
  Else '2D
    Let l = 0
    ReDim VecNew(0 To xRow, 0 To xCol)
    For j = 0 To xCol
      For i = 0 To xRow
        Let VecNew(i, j) = VBA.Abs(VBA.CDbl(VecSource(l)))
        Let l = l + 1
      Next i
    Next j
  End If
  Let VecSource = Empty

  'Return
  Let Balmet_assets_get_grades = VecNew
  If BResponseDimension = eqlBalAs1D Then
    Let BLngCounter = ((UBound(VecNew) - LBound(VecNew)) + 1)
  Else
    Let BLngCounter = ((UBound(VecNew) - LBound(VecNew)) + 1) * ((UBound(VecNew, 2) - LBound(VecNew, 2)) + 1)
  End If

EH:
  Let VecSource = Empty
  Let VecNew = Empty
  Set RCT = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Balmet_assets_get_grades_cx( _
  ByVal BVarVilSourceCx As Variant, _
  ByVal BLngVilSource As Long, _
  ByRef BLngCounter As Long, _
  Optional BxChar As String = ";", _
  Optional BResponseDimension As EQLBAL_ENU_GRADESVECTORTYPE = eqlBalAs1D) As Variant

  'Resturn a string with grades of normal Balmet
  Dim RCT As AppResourcesController_current
  
  Dim VecSource As Variant, VecNew As Variant
  Dim i As Long, j As Long, k As Long, l As Long, xRow As Long, xCol As Long
  Dim StrFound As String, StrAux As String

  Const sFN As String = "APPBALMET::-assets_get_grades_cx"

  On Error GoTo EH
  Let BLngCounter = 0
  
  If VBA.TypeName(BVarVilSourceCx) = "String" Then
    Let VecSource = VBA.Split(BVarVilSourceCx, BxChar)
  Else
    Let VecSource = BVarVilSourceCx
  End If
  Let l = UBound(VecSource) + 1
  Let k = BLngVilSource
  
  Set RCT = New AppResourcesController_current
  Set RCT.ErrorHandler = EHGLOBAL
  'Let xRow = Balmet_assets_get_products_bounds(k, l, eqlBalRows) - 1
  'Let xCol = Balmet_assets_get_products_bounds(k, l, eqlBalColumnsCx) - 1
  Let xRow = RCT.ResCtrBalGetProductsBounds(k, l, eqlBalRows) - 1
  Let xCol = RCT.ResCtrBalGetProductsBounds(k, l, eqlBalColumnsCx) - 1
  Set RCT = Nothing
  
  If BResponseDimension = eqlBalAs1D Then '1D
    Let l = 0
    ReDim VecNew((UBound(VecSource) - LBound(VecSource)))
    For i = LBound(VecSource) To UBound(VecSource)
      Let VecNew(l) = VecSource(i)
      Let l = l + 1
    Next i
  Else '2D
    ReDim VecNew(0 To xRow, 0 To xCol)
    Let l = LBound(VecSource)
    For j = 0 To xCol
      For i = 0 To xRow
        Let VecNew(i, j) = VecSource(l)
        Let l = l + 1
      Next i
    Next j
  End If
  Let VecSource = Empty

  'Return
  Let Balmet_assets_get_grades_cx = VecNew
  If BResponseDimension = eqlBalAs1D Then
    Let BLngCounter = ((UBound(VecNew) - LBound(VecNew)) + 1)
  Else
    Let BLngCounter = ((UBound(VecNew) - LBound(VecNew)) + 1) * ((UBound(VecNew, 2) - LBound(VecNew, 2)) + 1)
  End If

EH:
  Let VecSource = Empty
  Let VecNew = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)
  Set RCT = Nothing

End Function

Private Function Balmet_assets_check_names( _
  ByVal BRngBox As Range, _
  ByVal BStrNeedle As String, _
  ByRef BStrFound As String) As Boolean
  
  Dim RES As AppResources_current
  
  Dim Vec As Variant
  Dim Str1 As String, Str2 As String
  Dim i As Long, j As Long
  
  Const sFN As String = "APPBALMET::-assets_check_names"

  On Error GoTo EH
  Let Balmet_assets_check_names = False
  
  If BRngBox Is Nothing Then GoTo EH
  Let Vec = BRngBox.Value
  If VBA.IsArray(Vec) Then GoTo EH
  Set RES = New AppResources_current
  Set RES.ErrorHandler = EHGLOBAL
  If RES.ArrayLength(Vec) = 0 Then GoTo EH
  Set RES = Nothing
  
  For j = LBound(Vec, 2) To UBound(Vec, 2)
    For i = LBound(Vec) To UBound(Vec)
      Let Str1 = VBA.CStr(Vec(i, j))
      If (VBA.LCase(BStrNeedle) Like Str2) Then
        Let BStrFound = Str1
        Let Balmet_assets_check_names = True
        Exit For
      End If
    Next i
  Next j

EH:
  Let Vec = Empty
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)
  
End Function

Private Function Balmet_assets_check_names_and_units( _
  ByVal BStrName As String, _
  ByVal BStrUnit As String, _
  ByRef BVarErr As Variant) As Boolean
  
  Dim MDL As ModelExcliqliteDatasheet
  
  Dim RngBox0 As Range, RngBox1 As Range, RngBox2 As Range, RngBox3 As Range
  Dim StrAux0 As String, StrAux1 As String
  Dim i As Long, j As Long
  
  Const sFN As String = "APPBALMET::-assets_check_names_and_units"

  On Error GoTo EH
  Let Balmet_assets_check_names_and_units = False
  
  'Set RngBox0 = excliqlitedatasheet.ListObjects("excliqlitechemicalelements").DataBodyRange
  'Set RngBox1 = excliqlitedatasheet.ListObjects("excliqlitechemicalunits").DataBodyRange
  'Set RngBox2 = excliqlitedatasheet.ListObjects("excliqlitechemicalelementscx").DataBodyRange
  'Set RngBox3 = excliqlitedatasheet.ListObjects("excliqlitechemicalunitscx").DataBodyRange
  
  Set MDL = New ModelExcliqliteDatasheet 'TODO: ING!!!
  Set MDL.ErrorHandler = EHGLOBAL
  Set RngBox0 = MDL.MGet(eqlMdlTblChemicalelements, eqlMdlRange)
  Set RngBox1 = MDL.MGet(eqlMdlTblChemicalunits, eqlMdlRange)
  Set RngBox2 = MDL.MGet(eqlMdlTblChemicalelementscx, eqlMdlRange)
  Set RngBox3 = MDL.MGet(eqlMdlTblChemicalunitscx, eqlMdlRange)
  Set MDL = Nothing
  
  If RngBox0 Is Nothing Or RngBox1 Is Nothing Or RngBox2 Is Nothing Or RngBox3 Is Nothing Then Let BVarErr = VBA.CVErr(xlErrNA): GoTo EH 'Table doesn't exist!
        
  'Check if element introduced by user exists in database
  Let StrAux0 = VBA.LCase(VBA.CStr(BStrName))
  'If Not Balmet_assets_check_names(RngBox0, StrAux0, StrAux0) Then Let BVarErr = VBA.CVErr(xlErrNA): GoTo EH 'Error, invalid Unit!
  Let StrAux1 = VBA.LCase(VBA.CStr(BStrUnit))
  If Application.WorksheetFunction.CountIf(RngBox1, StrAux1) = 0 Then Let BVarErr = VBA.CVErr(xlErrNA): GoTo EH 'Error, invalid Unit!
  If Application.WorksheetFunction.CountIf(RngBox0, StrAux0) = 0 Then Let BVarErr = VBA.CVErr(xlErrNA): GoTo EH 'Error, invalid Name!
  'If element is silver group (cx), it must to have any unit of cx units table, otherwise, no other element can have DM as unit
  If Application.WorksheetFunction.CountIf(RngBox2, StrAux0) > 0 Then
    If Application.WorksheetFunction.CountIf(RngBox3, StrAux1) = 0 Then Let BVarErr = VBA.CVErr(xlErrNA): GoTo EH 'Error, invalid Unit!
    If Not (StrAux0 = "ag" Or StrAux0 = "plata") Then
      If (StrAux1 = "dm") Then Let BVarErr = VBA.CVErr(xlErrNA): GoTo EH 'Error, No DMs Unit in other elements!
    End If
  Else
    If (StrAux1 = "dm") Then Let BVarErr = VBA.CVErr(xlErrNA): GoTo EH 'Error, No DMs Unit in other elements!
  End If

  Let Balmet_assets_check_names_and_units = True
  
EH:
  Set RngBox0 = Nothing
  Set RngBox1 = Nothing
  Set RngBox2 = Nothing
  Set RngBox3 = Nothing
  Set MDL = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Balmet_assets_check_names_and_units_cx( _
  ByVal BStrName As String, _
  ByVal BStrUnit As String, _
  ByRef BVarErr As Variant) As Boolean
  
  Dim MDL As ModelExcliqliteDatasheet
  
  Dim RngBox0 As Range, RngBox1 As Range, RngBox2 As Range, RngBox3 As Range
  Dim StrAux0 As String, StrAux1 As String
  Dim i As Long, j As Long
  
  Const sFN As String = "APPBALMET::-assets_check_names_and_units_cx"

  On Error GoTo EH
  Let Balmet_assets_check_names_and_units_cx = False
      
  'Set RngBox0 = excliqlitedatasheet.ListObjects("excliqlitechemicalelementscx").DataBodyRange
  'Set RngBox1 = excliqlitedatasheet.ListObjects("excliqlitechemicalunitscx").DataBodyRange
  
  Set MDL = New ModelExcliqliteDatasheet
  Set MDL.ErrorHandler = EHGLOBAL
  Set RngBox0 = MDL.MGet(eqlMdlTblChemicalelementscx, eqlMdlRange)
  Set RngBox1 = MDL.MGet(eqlMdlTblChemicalunitscx, eqlMdlRange)
  Set MDL = Nothing
  If RngBox0 Is Nothing Or RngBox1 Is Nothing Then Let BVarErr = VBA.CVErr(xlErrNA): GoTo EH 'Table doesn't exist!
        
  'Check if element introduced by user exists in database
  Let StrAux0 = VBA.LCase(VBA.CStr(BStrName))
  Let StrAux1 = VBA.LCase(VBA.CStr(BStrUnit))
  If Application.WorksheetFunction.CountIf(RngBox1, StrAux1) = 0 Then Let BVarErr = VBA.CVErr(xlErrNA): GoTo EH 'Error, invalid Unit!
  'If element is silver group (cx), it must to have any unit of cx units table, otherwise, no other element can have DM as unit
  If Application.WorksheetFunction.CountIf(RngBox0, StrAux0) = 0 Then Let BVarErr = VBA.CVErr(xlErrNA): GoTo EH 'Error, invalid Name!
  If Not (StrAux0 = "ag" Or StrAux0 = "plata") Then
    If (StrAux1 = "dm") Then Let BVarErr = VBA.CVErr(xlErrNA): GoTo EH 'Error, No DMs Unit in other elements!
  End If

  Let Balmet_assets_check_names_and_units_cx = True
  
EH:
  Set RngBox0 = Nothing
  Set RngBox1 = Nothing
  Set MDL = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

'Private Function Balmet_assets_check_negatives_OOLD(ByVal BVarVector As Variant) As Integer
'
''  Dim i As Variant
''  Dim j As Byte
''
''  On Error GoTo EH
''  Let Balmet_assets_check_negatives = -1
''
''  If Not VBA.IsArray(BVarVector) Then GoTo EH
''  Let j = 0
''  For Each i In BVarVector
''    If VBA.CDbl(i) < 0 Then Let j = 1: Exit For
''  Next i
''
''  Let Balmet_assets_check_negatives = j
''
''EH:
''  Call EHGLOBAL.ErrorHandlerRaise("APPBALMET::-assets_check_negatives")
'
'End Function





' ERRORS MANNAGEMENT AND ASSERTSIONS ------------------------------------------------'
Private Sub Debug_assert(ByVal Cond As Boolean, Optional Origin As String = VBA.vbNullString)
  
  #If Debbuging Then
    Debug.Assert Cond
    If Not Cond Then Debug.Print "Assert on: " & Origin
  #End If

End Sub





' METHODS LIST PUBLIC ---------------------------------------------------------------'
' ABSTRACT INTERFACE UI METHODS LIST AND UDF METHODS LIST (PUBLIC) ------------------'
Public Function ICoreController_GetSolution(ByVal CStrRequester As String, ByVal CVarRequest As Variant, CVarResponse As Variant) As Boolean
  
  On Error GoTo EH
  Let ICoreController_GetSolution = Balmet_core_main(CStrRequester, CVarRequest, CVarResponse)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("BALMET::IGetSolution")

End Function

Public Function ICoreController_GetSolutionUDF(ByVal CStrRequester As String, ByVal CVarRequest As Variant) As Variant
  
  Let ICoreController_GetSolutionUDF = Balmet_core_udf(CStrRequester, CVarRequest)
  
End Function
' METHODS LIST PUBLIC ---------------------------------------------------------------'
' ============================= CLASS BALMET STRUCTURE =============================='

