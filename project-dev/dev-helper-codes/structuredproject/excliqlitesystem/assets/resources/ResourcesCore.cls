VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppResourcesController_current"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' OBJECT VARIABLES (GLOOBJ_) --------------------------------------------------------'
Private EHGLOBAL As AppErrorHandler

' VARIABLES (GLOStr_) ---------------------------------------------------------------'
Private GLOBOO_FORMULAS As Boolean
'Public Enum EQLBAL_ENU_METHOD
'  eqlBalConventional
'  eqlBalCramer
'  eqlBalInverseMatrix
'End Enum
'Public Enum EQLBAL_ENU_RESULT
'  eqlBalWeights
'  eqlBalWeightPercents
'  eqlBalGradesHeads
'  eqlBalUnities
'  eqlBalFines
'  eqlBalRecoveries
'  eqlBalRatio
'  eqlBalGradesHeadsCx
'  eqlBalUnitiesCx
'  eqlBalFinesCx
'  eqlBalRecoveriesCx
'  eqlBalVolume
'  eqlBalVolumePercents
'End Enum
'Public Enum EQLBAL_ENU_RESULTECO
'  eqlBalHeadsGrades
'  eqlBalProdsGrades
'End Enum
'Public Enum EQLBAL_ENU_RESULTDIRECTION
'  eqlBalVertical
'  eqlBalHorizontal
'End Enum
Public Enum EQLBAL_ENU_BOUNDS
  eqlBalProducts
  eqlBalColumns
  eqlBalRows
  eqlBalFullProducts
  eqlBalFullProductsCx
  eqlBalProductsCx
  eqlBalColumnsCx
End Enum
'Private Enum EQLBAL_ENU_TYPECONTENT
'  eqlBalSolids
'  eqlBalVolumes
'  eqlBalBoth
'End Enum
'Private Enum EQLBAL_ENU_TYPERESULT
'  eqlBalJustGrades
'  eqlBalJustGradesCx
'  eqlBalJustGradesBoth
'End Enum
Public Enum EQLBAL_ENU_TYPE
  eqlBalNormal
  eqlBalWithComplex
  eqlBalNothing
End Enum
'Public Enum EQLBAL_ENU_GRADESVECTORTYPE
'  eqlBalAs1D
'  eqlBalAs2D
'End Enum


' CONSTANTS (GLOSTR_) ---------------------------------------------------------------'
















' PROPERTIES ------------------------------------------------------------------------'
Public Property Set ErrorHandler(ByRef RHS As AppErrorHandler)
  
  Set EHGLOBAL = Nothing
  Set EHGLOBAL = RHS
  Dim MDL As ModelExcliqliteDatasheet
  Set MDL = New ModelExcliqliteDatasheet
  Set MDL.ErrorHandler = EHGLOBAL
  Let GLOBOO_FORMULAS = MDL.Formulas
  Set MDL = Nothing

End Property

' LOCAL PROPERTIES ------------------------------------------------------------------'






' CONSTRUCTOR/DESTRUCTOR ------------------------------------------------------------'
Private Sub Class_Terminate()
  
  Set EHGLOBAL = Nothing

End Sub






' CORE METHODS LIST PRIVATE ---------------------------------------------------------'
' ASSETS ----------------------------------------------------------------------------'
Private Function Resctr_assets_bal_get_products_bounds( _
  ByVal BLngSource As Long, _
  Optional BLngSourceCx As Long = 0, _
  Optional BWhat As EQLBAL_ENU_BOUNDS = eqlBalProducts) As Long
  
  'Returns the number of Products, Cols, Rows, TotalProducts and TotalComplexProducts of any Balmet
  Dim LngProd As Long, LngCol As Long, LngRow As Long, LngTotProd As Long, LngTotCx As Long, LngAux0 As Long
  Dim BooFound As Boolean, BooTimeup As Boolean
  
  Const LngLASTBOUND As Long = 10
  Const sFN As String = "RES::-assets_bal_get_products_bounds"
  
  On Error GoTo EH
  Let Resctr_assets_bal_get_products_bounds = 0
    
  Let BooFound = False
  Let BooTimeup = False
  Let LngProd = 2
  Do While (Not BooFound And Not BooTimeup)
    Let LngCol = LngProd - 1
    Let LngRow = (LngProd + 1)
    Let LngTotProd = LngCol * LngRow
    
    Let LngProd = LngProd + 1
    Let BooFound = (BLngSource = LngTotProd)
    Let BooTimeup = (LngProd = LngLASTBOUND)
    'Exit Do
  Loop
  If Not BooFound Then
    GoTo EH
  Else 'Determine Cx elements
    If BLngSourceCx > 0 Then
      Let LngTotCx = IIf(BLngSourceCx = LngRow Or BLngSourceCx = (LngRow * 2) Or BLngSourceCx = (LngRow * 3), BLngSourceCx + LngTotProd, 0)
      If LngTotCx = 0 Then GoTo EH
    End If
  End If
  
  'Return
  Select Case BWhat
    Case eqlBalProducts: Let Resctr_assets_bal_get_products_bounds = LngProd - 1 'Number of Balmet Products
    Case eqlBalColumns: Let Resctr_assets_bal_get_products_bounds = LngCol   'Number of normal Balmet Cols
    Case eqlBalRows: Let Resctr_assets_bal_get_products_bounds = LngRow 'Number of Balmet Rows
    Case eqlBalFullProducts: Let Resctr_assets_bal_get_products_bounds = LngTotProd 'Number of Full normal Balmet grades. Ej. 8 grades to Balmet of 3 Prods
    Case eqlBalFullProductsCx: Let Resctr_assets_bal_get_products_bounds = LngTotCx 'Number of Full normal and complex Balmet grades. Ej. 12 grades to Balmet of 3 Prods
    Case eqlBalProductsCx: Let Resctr_assets_bal_get_products_bounds = LngTotCx - LngTotProd 'Number of Full complex Balmet grades. Ej. 12-8 grades = 4 to Balmet of 3 Prods
    Case eqlBalColumnsCx
      Let LngAux0 = LngTotCx - LngTotProd
      Let Resctr_assets_bal_get_products_bounds = LngAux0 / LngRow 'Number of complex Balmet columns.
    Case Else: GoTo EH
  End Select

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Resctr_assets_bal_get_products_bounds_fulldata( _
  ByVal BLngSource As Long, _
  ByRef BLngProducts As Long, _
  ByRef BLngColumns As Long, _
  ByRef BLngRows As Long, _
  ByRef BLngElements As Long, _
  ByRef BLngElementsAndCx As Long, _
  ByRef BLngElementsCx As Long, _
  ByRef BLngColumnsCx As Long, _
  Optional BLngSourceCx As Long = 0) As Boolean
  
  'Returns the number of Products, Cols, Rows, TotalProducts and TotalComplexProducts of any Balmet
  Dim LngProd As Long, LngCol As Long, LngRow As Long, LngTotProd As Long, LngTotCx As Long
  Dim BooFound As Boolean, BooTimeup As Boolean
  
  Const LngLASTBOUND As Long = 10
  Const sFN As String = "RES::-assets_bal_get_products_bounds_fulldata"
  
  On Error GoTo EH
  Let Resctr_assets_bal_get_products_bounds_fulldata = False
  
  'Determine if the Source number is a valid value of elements to be a Balmet
  Let BooFound = False
  Let BooTimeup = False
  Let LngProd = 2
  Do While (Not BooFound And Not BooTimeup)
    Let LngCol = LngProd - 1
    Let LngRow = (LngProd + 1)
    Let LngTotProd = LngCol * LngRow
    
    Let LngProd = LngProd + 1
    Let BooFound = (BLngSource = LngTotProd)
    Let BooTimeup = (LngProd = LngLASTBOUND)
    'Exit Do
  Loop
  If Not BooFound Then
    GoTo EH
  Else 'Determine Cx elements (just 3 columns of Cx allowed)
    If BLngSourceCx > 0 Then
      Let LngTotCx = IIf(BLngSourceCx = LngRow Or BLngSourceCx = (LngRow * 2) Or BLngSourceCx = (LngRow * 3), BLngSourceCx + LngTotProd, 0)
      If LngTotCx = 0 Then GoTo EH
    End If
  End If
  
  'Return
  Let BLngProducts = LngProd - 1 'Number of Balmet Products
  Let BLngColumns = LngCol 'Number of normal Balmet Cols
  Let BLngRows = LngRow 'Number of Balmet Rows
  Let BLngElements = LngTotProd 'Number of Full normal Balmet grades. Ej. 8 grades to Balmet of 3 Prods
  Let BLngElementsAndCx = IIf(LngTotCx = 0, LngTotProd, LngTotCx) 'Number of Full normal and complex Balmet grades. Ej. 12 grades to Balmet of 3 Prods
  Let BLngElementsCx = BLngElementsAndCx - LngTotProd 'Number of Full complex Balmet grades. Ej. 12-8 grades = 4 to Balmet of 3 Prods
  Let BLngColumnsCx = (BLngElementsAndCx - LngTotProd) / LngRow 'Number of complex Balmet columns.
  
  Let Resctr_assets_bal_get_products_bounds_fulldata = True

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Resctr_assets_bal_get_type(ByVal BLngSource As Long, Optional BLngSourceCx As Long = 0) As EQLBAL_ENU_TYPE
  
  'Return "normal" or "complex" type of any Balmet
  Dim i As Long, j As Long, k As Long, l As Long, m As Long, n As Long, o As Long
  Dim StrType As EQLBAL_ENU_TYPE
  
  Const sFN = "RESCTR::-assets_bal_get_type"

  On Error GoTo EH
  Let l = 0
  Let n = 0
  Let StrType = eqlBalNothing
  
  If Resctr_assets_bal_get_products_bounds_fulldata(BLngSource, i, j, k, l, m, n, o, BLngSourceCx) Then
    If n > 0 Then
      Let StrType = eqlBalWithComplex
    Else
      Let StrType = eqlBalNormal
    End If
  End If
  
  'Return
  Let Resctr_assets_bal_get_type = StrType

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Resctr_assets_bal_is_balmet(ByVal BLngSource As Long, Optional BLngSourceCx As Long = 0) As Boolean
  
  'Return if source is a correct number of elements for any Balmet
  Const sFN As String = "RESCTR::-assets_bal_is_balmet"

  On Error GoTo EH
  'Return
  Let Resctr_assets_bal_is_balmet = Not (Resctr_assets_bal_get_type(BLngSource, BLngSourceCx) = eqlBalNothing)

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Resctr_assets_bal_get_generic_names( _
  ByVal VStrGrades As String, _
  ByVal VStrGradesCx As String, _
  ByVal VCx As Boolean, _
  ByVal VUnits As Boolean) As String
  
  Dim VecAux0 As Variant
  Dim StrAux0 As String
  Dim i As Long, j As Long, k As Long
  
  Const sFN As String = "RESCTR:-assets_bal_get_generic_names"
  
  On Error GoTo EH
  
  Let StrAux0 = VBA.vbNullString
  Let VecAux0 = VBA.Split(VStrGrades, ";")
  Let j = UBound(VecAux0) - LBound(VecAux0) + 1
  If VCx Then
    Let VecAux0 = Empty
    Let VecAux0 = VBA.Split(VStrGradesCx, ";")
    Let k = UBound(VecAux0) - LBound(VecAux0) + 1
  End If
  Let VecAux0 = Empty
  
  If Me.ResCtrIsBalmet(j) Then
    Let j = Me.ResCtrBalGetProductsBounds(j, BWhat:=eqlBalProducts) - 1
    Let k = Me.ResCtrBalGetProductsBounds(j, k, eqlBalProductsCx) - 1
    ReDim VecAux0(IIf(VCx, k, j))
    For i = 0 To UBound(VecAux0)
      If VUnits Then
        Let VecAux0(i) = IIf(VCx, "g/T", "%")
      Else
        Let VecAux0(i) = "B" & IIf(VCx, "X", VBA.vbNullString) & i + 1
      End If
    Next i
    Let StrAux0 = VBA.Join(VecAux0, ";")
    Let VecAux0 = Empty
  End If
  
  Let Resctr_assets_bal_get_generic_names = StrAux0
  
EH:
  Let VecAux0 = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function


'Private Function Resctr_assets_bal_get_complex_datas( _
'  ByRef BStrSource As String, _
'  ByRef BLngCol As Long, _
'  Optional BBooNamesUnits As Boolean = False, _
'  Optional BBooDelFromSource As Boolean = True) As String
'
'  'Resturn a string of Complex Minerals Names or Units
'  Dim VecSource As Variant, VecOcean As Variant
'  Dim i As Long, j As Long, kDM As Long
'  Dim StrFound As String
'
'  Const sFN As String = "RESCTR::-assets_bal_get_complex_datas"
'
'  On Error GoTo EH
'  Let Resctr_assets_bal_get_complex_datas = VBA.vbNullString
'  Let kDM = 0
'  Let BLngCol = 0
'  Let StrFound = VBA.vbNullString
'  Let VecSource = VBA.Split(BStrSource, ";")
'  If BBooNamesUnits Then 'Units TODO: WITH DB
'    Let VecOcean = Array("dm", "g/t", "g/l")
'  Else 'Names
'    Let VecOcean = Array("ag", "au", "pt")
'  End If
'  For i = 0 To UBound(VecSource)
'    For j = 0 To UBound(VecOcean)
'      If VBA.LCase(VecSource(i)) = VecOcean(j) Then
'        Let StrFound = VecSource(i)
'        If Not BBooNamesUnits Then
'          If BBooDelFromSource Then
'            Let BStrSource = VBA.Replace(BStrSource, StrFound, VBA.vbNullString, Count:=1)
'            Let BLngCol = BLngCol + 1
'          End If
'        Else
'          If BBooDelFromSource Then
'            If VBA.LCase(StrFound) = "dm" Then
'              Let BStrSource = VBA.Replace(BStrSource, StrFound, VBA.vbNullString, Count:=1)
'              Let kDM = kDM + 1
'            ElseIf i = UBound(VecSource) Then
'              Let BStrSource = VBA.Replace(BStrSource, StrFound, VBA.vbNullString, Count:=1)
'              Let BLngCol = BLngCol + 1
'            End If
'          End If
'        End If
'      End If
'    Next j
'  Next i
'  If Not BBooNamesUnits Then
'    If BLngCol > 1 Or BLngCol = 0 Then Let BLngCol = 0: GoTo EH
'  Else
'    If Not (kDM = 1 And BLngCol = 0) Then
'      If Not (kDM = 0 And BLngCol = 1) Then Let BLngCol = 0: GoTo EH
'    End If
'  End If
'
'  'Return
'  Let Resctr_assets_bal_get_complex_datas = StrFound
'
'EH:
'  Let VecSource = Empty
'  Let VecOcean = Empty
'  Call EHGLOBAL.ErrorHandlerRaise(sFN)
'End Function

'Private Function Resctr_assets_bal_get_units( _
'  ByRef BStrSource As String, _
'  ByRef BLngCol As Long, _
'  Optional BBooDelFromSource As Boolean = True) As String
'
'  'Resturn a string with Units of grades
'  Dim VecSource As Variant, VecFound As Variant, VecOcean As Variant
'  Dim i As Long, j As Long, LngDMCounter As Long
'  Dim StrFound As String, StrAux As String
'
'  Const sFN As String = "RESCTR::-assets_bal_get_units"
'
'  On Error GoTo EH
'  Let Resctr_assets_bal_get_units = VBA.vbNullString
'
'  Let BLngCol = 0
'  Let LngDMCounter = 0
'  Let StrFound = VBA.vbNullString
'  Let VecSource = VBA.Split(BStrSource, ";")
'  Let VecOcean = Array("dm", "g/t", "g/l", "%")
'  ReDim VecFound(0)
'  For i = 0 To UBound(VecSource)
'    For j = 0 To UBound(VecOcean)
'      If VBA.LCase(VecSource(i)) = VecOcean(j) Then
'        ReDim Preserve VecFound(BLngCol)
'        Let StrAux = VecSource(i)
'        Let VecFound(BLngCol) = StrAux
'        If BBooDelFromSource Then Let BStrSource = VBA.Replace(BStrSource, StrAux, VBA.vbNullString, Count:=1)
'        Let BLngCol = BLngCol + 1
'        Let LngDMCounter = LngDMCounter + IIf(VecOcean(j) = "dm", 1, 0)
'      End If
'    Next j
'  Next i
'  If LngDMCounter > 1 And BLngCol > 0 Then GoTo EH
'
'  'Return
'  Let Resctr_assets_bal_get_units = VBA.Join(VecFound, ";")
'
'EH:
'  Let VecSource = Empty
'  Let VecFound = Empty
'  Let VecOcean = Empty
'  Call EHGLOBAL.ErrorHandlerRaise(sFN)
'End Function

'Private Function Resctr_assets_bal_get_units_refill(ByVal BStrUnit As String, ByRef BLngCounter As Long) As Variant
'  Dim VecSource As Variant
'  Dim i As Long
'
'  Const sFN As String = "RESCTR::-assets_get_units_refill"
'
'  On Error GoTo EH
'  Let Resctr_assets_bal_get_units_refill = Empty
'
'  ReDim VecSource(BLngCounter - 1)
'  For i = 0 To UBound(VecSource)
'    Let VecSource(i) = BStrUnit
'  Next i
'
'  Let BLngCounter = IIf(i > 0, i, 0)
'  Let Resctr_assets_bal_get_units_refill = VecSource
'
'EH:
'  Let VecSource = Empty
'  Call EHGLOBAL.ErrorHandlerRaise(sFN)
'End Function









' ERRORS MANNAGEMENT AND ASSERTSIONS ------------------------------------------------'
Private Sub Debug_assert(ByVal Cond As Boolean, Optional Origin As String = VBA.vbNullString)
  #If Debbuging Then
    Debug.Assert Cond
    If Not Cond Then Debug.Print "Assert on: " & Origin
  #End If
End Sub








' INTERFACE METHODS LIST FRM - SEL --------------------------------------------------'
Public Function ResCtrBalGetProductsBounds( _
  ByVal BLngSource As Long, _
  Optional BLngSourceCx As Long = 0, _
  Optional BWhat As EQLBAL_ENU_BOUNDS = eqlBalProducts) As Long
  
  'Returns the number of Products, Cols, Rows, TotalProducts and TotalComplexProducts of any Balmet
  Const sFN As String = "RESCTR::BalGetProductsBounds"
  
  On Error GoTo EH
  Let ResCtrBalGetProductsBounds = Resctr_assets_bal_get_products_bounds(BLngSource, BLngSourceCx, BWhat)
  
EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Public Function ResCtrBalGetProductsBoundsFulldata( _
  ByVal BLngSource As Long, _
  ByRef BLngProducts As Long, _
  ByRef BLngColumns As Long, _
  ByRef BLngRows As Long, _
  ByRef BLngElements As Long, _
  ByRef BLngElementsAndCx As Long, _
  ByRef BLngElementsCx As Long, _
  ByRef BLngColumnsCx As Long, _
  Optional BLngSourceCx As Long = 0) As Boolean
  
  'Returns the number of Products, Cols, Rows, TotalProducts and TotalComplexProducts of any Balmet
  Const sFN As String = "RESCTR::BalGetProductsBoundsFulldata"
  
  On Error GoTo EH
  Let ResCtrBalGetProductsBoundsFulldata = Resctr_assets_bal_get_products_bounds_fulldata(BLngSource, BLngProducts, BLngColumns, BLngRows, BLngElements, BLngElementsAndCx, BLngElementsCx, BLngColumnsCx, BLngSourceCx)
  
EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Public Function ResCtrBalGetType(ByVal BLngSource As Long, Optional BLngSourceCx As Long = 0) As EQLBAL_ENU_TYPE
  
  'Return "normal" or "complex" type of any Balmet
  Const sFN As String = "RESCTR::BalGetType"
  
  On Error GoTo EH
  Let ResCtrBalGetType = Resctr_assets_bal_get_type(BLngSource, BLngSourceCx)
  
EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Public Function ResCtrIsBalmet(ByVal BLngSource As Long, Optional BLngSourceCx As Long = 0) As Boolean
  
  'Return if source is a correct number of elements of any Balmet
  Const sFN As String = "RESCTR::IsBalmet"
  
  On Error GoTo EH
  Let ResCtrIsBalmet = Resctr_assets_bal_is_balmet(BLngSource, BLngSourceCx)
  
EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Public Function ResCtrGetGenericNames( _
  ByVal BLngSource As String, _
  Optional BLngSourceCx As String = VBA.vbNullString, _
  Optional BBooCx As Boolean = False, _
  Optional BBooUnits As Boolean = False) As String
  
  ' Return a string of generic names or units
  Const sFN As String = "RESCTR::GetGenericNames"
  
  On Error GoTo EH
  Let ResCtrGetGenericNames = Resctr_assets_bal_get_generic_names(BLngSource, BLngSourceCx, BBooCx, BBooUnits)
  
EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

'Public Function ResCtrBalGetComplexDatas( _
'  ByRef BStrSource As String, _
'  ByRef BLngCol As Long, _
'  Optional BBooNamesUnits As Boolean = False, _
'  Optional BBooDelFromSource As Boolean = True) As String
'
'  'Returns a string of Complex Minerals Names or Units
'  Const sFN As String = "RESCTR::BalGetComplexDatas"
'
'  On Error GoTo EH
'  Let ResCtrBalGetComplexDatas = Resctr_assets_bal_get_complex_datas(BStrSource, BLngCol, BBooNamesUnits, BBooDelFromSource)
'
'EH:
'  Call EHGLOBAL.ErrorHandlerRaise(sFN)
'
'End Function

'Public Function ResCtrBalGetUnits( _
'  ByRef BStrSource As String, _
'  ByRef BLngCol As Long, _
'  Optional BBooDelFromSource As Boolean = True) As String
'
'  'Resturn a string with Units of grades
'  Const sFN As String = "RESCTR::BalGetUnits"
'
'  On Error GoTo EH
'  Let ResCtrBalGetUnits = Resctr_assets_bal_get_units(BStrSource, BLngCol, BBooDelFromSource)
'
'EH:
'  Call EHGLOBAL.ErrorHandlerRaise(sFN)
'
'End Function
' METHODS PUBLIC --------------------------------------------------------------------'




