VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassEqlControllerBalmet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ================== CLASS EXCLIQ LITE CONTROLLER BALMET STRUCTURE ================= '
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


' ================== CLASS EXCLIQ LITE CONTROLLER BALMET STRUCTURE ================= '
' ABSTRACT IMPLEMENTATIONS --------------------------------------------------------- '
' OBJECT VARIABLES (GLOOBJ_) ------------------------------------------------------- '
' VARIABLES (GLOStr_) -------------------------------------------------------------- '
' VARIABLES CUSTOM (GLOEnu|Typ_) --------------------------------------------------- '
' CONSTANTS (GLOSTR_) -------------------------------------------------------------- '
' CONSTANTS ERROR MESSAGES (GLOSTR_ERR_) ------------------------------------------- '
' PROPERTIES ----------------------------------------------------------------------- '
' CONSTRUCTOR/DESTRUCTOR ----------------------------------------------------------- '
' CORE METHODS LIST PRIVATE -------------------------------------------------------- '
' UDF CORE METHODS
' Balmet_core_udf
' Balmet_core_main
' Balmet_core_main_udf
' Balmet_core_main_udf_validate
' Balmet_core_main_udf_prepare
' Balmet_core_main_udf_calculate
' APP CORE METHODS
' Balmet_core_main_app
' UDF CORE METHODS - ECO
' Balmet_core_main_udf_eco
' Balmet_core_main_udf_validate_eco
' Balmet_core_main_udf_prepare_eco
' Balmet_core_main_udf_calculate_eco
' APP CORE METHODS - ECO
' Balmet_core_main_app_eco
' Balmet_core_main_app_eco_spreadbalmet_get_grades_hybrid_data
' Balmet_core_main_app_eco_spreadbalmet_get_percents_hybrid_data
' CORE METHODS
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
' ASSETS --------------------------------------------------------------------------- '
' ERRORS MANNAGEMENT AND ASSERTSIONS ----------------------------------------------- '
' Debug_assert
' METHODS LIST PUBLIC -------------------------------------------------------------- '
' ABSTRACT INTERFACE UI METHODS LIST AND UDF METHODS LIST (PUBLIC) ----------------- '
' ClassEqlControllerInterface_GetSolution
' ClassEqlControllerInterface_GetSolutionUDF
' ================== CLASS EXCLIQ LITE CONTROLLER BALMET STRUCTURE ================= '





' ================== CLASS EXCLIQ LITE CONTROLLER BALMET STRUCTURE ================= '
' ABSTRACT IMPLEMENTATIONS --------------------------------------------------------- '
Implements ClassEqlControllerInterface

' OBJECT VARIABLES (GLOOBJ_) ------------------------------------------------------- '
Private RES As ClassEqlResources

' VARIABLES (GLOStr_) -------------------------------------------------------------- '
Private GLOVEC_BALANCE As Variant
Private GLOBOO_FORMULAS As Boolean




' PROPERTIES ----------------------------------------------------------------------- '
Public Property Set ClassEqlControllerInterface_ResourcesApp(ByRef RHS As ClassEqlResources)

  Set RES = RHS
  Let GLOBOO_FORMULAS = RES.Formulas

End Property




' CONSTRUCTOR/DESTRUCTOR ----------------------------------------------------------- '
Private Sub Class_Terminate()

  Let GLOVEC_BALANCE = Empty
  Set RES = Nothing

End Sub





' CORE METHODS PRIVATE ------------------------------------------------------------- '
' UDF CORE METHODS
Private Function Balmet_core_udf(ByVal CStrRequester As String, ByVal CVarRequest As Variant) As Variant

  On Error GoTo EH
  Select Case CStrRequester
    Case "balmetudf": Let Balmet_core_udf = Balmet_core_main_udf(CVarRequest(0), CVarRequest(1), CVarRequest(2), CVarRequest(3), CVarRequest(4), CVarRequest(5), CVarRequest(6), CVarRequest(7), CVarRequest(8))
    Case "balmetudfeco": Let Balmet_core_udf = Balmet_core_main_udf_eco(CVarRequest(0), CVarRequest(1), CVarRequest(2), CVarRequest(3), CVarRequest(4), CVarRequest(5), CVarRequest(6), CVarRequest(7), CVarRequest(8), CVarRequest(9))
    Case Else: GoTo EH
  End Select

EH:
  If Not VBA.Err.Number = 0 Then Call RES.ErrorHandler.ErrorHandlerRaise("BALMET::-core_udf")

End Function

Private Function Balmet_core_main(ByVal CStrRequester As String, ByVal CVarRequest As Variant, CVarResponse As Variant) As Boolean

  On Error GoTo EH
  Select Case CStrRequester
    Case "balmet", "balmetselection": Let Balmet_core_main = Balmet_core_main_app(CVarRequest)
    Case "balmeteco": Let Balmet_core_main = Balmet_core_main_app_eco(CVarRequest, CVarRequest(14))
    Case Else: Let Balmet_core_main = False
  End Select

  If Balmet_core_main Then CVarResponse = GLOVEC_BALANCE

EH:
  Let GLOVEC_BALANCE = Empty
  Call RES.ErrorHandler.ErrorHandlerRaise("BALMET::-main")

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

  'Forming the request data
  Let VecAux0 = Array(BVarFeed, BVarGrad, BVarGradCx, BVarElem, BVarUnit)

  'Sets GLOVEC_BALANCE with final result
  If Not Balmet_core_main_udf_validate(VecAux0, Balmet_core_main_udf, BooComplex) Then GoTo EH
  If Not Balmet_core_main_udf_prepare(VecAux0, Balmet_core_main_udf, BooComplex, EnuTypeFeed) Then GoTo EH
  If Not Balmet_core_main_udf_calculate(VecAux0, BMethod, BResult, BDirection, Balmet_core_main_udf, BAllResult, EnuTypeFeed) Then GoTo EH

  Let Balmet_core_main_udf = GLOVEC_BALANCE
  Let VecAux0 = Empty
  Let GLOVEC_BALANCE = Empty

EH:
  Let VecAux0 = Empty

End Function

Private Function Balmet_core_main_udf_validate( _
  ByVal BVarRequest As Variant, _
  ByRef BVarErrUDF As Variant, _
  ByRef BBooComplex As Boolean) As Boolean

  ' Data required: (BVarFeed, BVarGrad, BVarGradCx, BVarElem, BVarUnit)
  Dim MDL As ClassEqlModel
  Dim VecFeed As Variant, VecGrad As Variant, VecUnit As Variant, VecName As Variant, VecGradCx As Variant, VecNameCx As Variant, VecUnitCx As Variant
  Dim VecAux0 As Variant, VecAux1 As Variant, VecAux2 As Variant, VecAux3 As Variant
  Dim StrAux0 As String
  Dim i As Long, j As Long, k As Long, l As Long, m As Long
  Dim EnuBalType As EQLBAL_ENU_TYPE

  On Error GoTo EH
  Let Balmet_core_main_udf_validate = False
  Set MDL = New ClassEqlModel
  Set MDL.ResourcesApp = RES

  ' Feed
  Let VecFeed = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(0), i, RValueType:=eqlResNumbers, RBooAbs:=True, RValueAs:=eqlResNumbers, RDblTop:=CDBL_MAX)
  If i <= 0 Then GoTo EH ' Feed is required

  ' Extract all balmet and cx grades
  Let VecGrad = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(1), j, RValueType:=eqlResNumbers, RBooAbs:=True, RValueAs:=eqlResNumbers, RDblTop:=CDBL_MAX)
  Let VecGradCx = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(2), k, RValueType:=eqlResNumbers, RBooAbs:=True, RValueAs:=eqlResNumbers, RDblTop:=CDBL_MAX)
  If j <= 0 Then Let BVarErrUDF = VBA.CVErr(xlErrNull): GoTo EH ' Grades are required
  If Not RES.RESCORE.BalIsBalmet(j, k) Then Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH ' Not a balmet
  ' Get the type of balmet to evaluate
  Let EnuBalType = RES.RESCORE.BalGetType(j, k) ' Balmet type
  If EnuBalType = eqlBalNothing Then Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH ' Not a balmet
  ' Get cols of products
  Let l = RES.RESCORE.BalGetProductsBounds(j, BWhat:=eqlBalColumns) ' Main Balmet cols
  Let m = RES.RESCORE.BalGetProductsBounds(j, k, BWhat:=eqlBalColumnsCx) ' Cx Balmet cols

  ' Names
  Let VecAux0 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalelements_sys, eqlMdlArray, "Símbolo"), eqlRes1D)
  Let VecAux1 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalelementscx_sys, eqlMdlArray, "Símbolo"), eqlRes1D)
  Let StrAux0 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(BVarRequest(3), j, RValueType:=eqlResNotNumbers, RBooUniques:=True)
  If Not RES.RESCORE.BalRefillNames2(RES, VecAux0, StrAux0, BLngGradCol:=(l - 1), BLngGradCxCol:=(m - 1), BVarName:=VecName, BVarNameCx:=VecNameCx, BStrErrMsg:=StrAux0, BBooCx:=(EnuBalType = eqlBalWithComplex), BVarDBNamesCx:=VecAux1) Then GoTo EH

  ' Units
  Let VecAux2 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalunits_sys, eqlMdlArray, "Unidad"), eqlRes1D)
  Let VecAux3 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalunitscx_sys, eqlMdlArray, "Unidad"), eqlRes1D)
  Let StrAux0 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(BVarRequest(4), j, RValueType:=eqlResNotNumbers, RStrCustomRegex:="%")
  If Not RES.RESCORE.BalRefillUnits2(RES, VecAux2, StrAux0, BLngGradCol:=(l - 1), BLngGradCxCol:=(m - 1), BVarUnit:=VecUnit, BVarUnitCx:=VecUnitCx, BStrErrMsg:=StrAux0, BBooCx:=(EnuBalType = eqlBalWithComplex), BVarDBUnitsCx:=VecAux3) Then GoTo EH
  Let VecAux2 = Empty
  Let VecAux3 = Empty

  ' Validations
  If Not RES.RESCORE.BalCheckCustomNamesDoesNotHasWrongUnits(RES, StrAux0, BVarName:=VecName, BVarUnit:=VecUnit, BEnuCheckWhat:=eqlBalAgUnit) Then GoTo EH
  If Not RES.RESCORE.BalCheckCustomNamesDoesNotHasWrongUnits(RES, StrAux0, VecAux1, VecName, VecUnit, eqlBalCxElemsUnitOnNormalBalmet) Then GoTo EH
  If Not RES.RESCORE.BalCheckCustomNamesDoesNotHasWrongUnits(RES, StrAux0, BVarName:=VecName, BVarUnit:=VecUnit, BEnuCheckWhat:=eqlBalDMOnOtherElements) Then GoTo EH

  If Not RES.RESCORE.BalCheckForMaximousValuesInGrades(StrAux0, VecGrad, VecUnit, VecGradCx, VecUnitCx, (EnuBalType = eqlBalWithComplex)) Then GoTo EH

  If EnuBalType = eqlBalWithComplex Then
    If Not RES.RESCORE.BalCheckCustomNamesDoesNotHasWrongUnits(RES, StrAux0, BVarSrcElems:=VecAux1, BVarName:=VecName, BEnuCheckWhat:=eqlBalNonCx) Then GoTo EH
    If Not RES.RESCORE.BalCheckCustomNamesDoesNotHasWrongUnits(RES, StrAux0, BVarSrcElems:=Empty, BVarName:=VecName, BVarUnit:=VecUnit, BEnuCheckWhat:=eqlBalDMOnOtherElements) Then GoTo EH
    If Not RES.RESCORE.BalCheckForMaximousValuesInGrades(StrAux0, VecGrad, VecUnit, VecGradCx, VecUnitCx, (EnuBalType = eqlBalWithComplex)) Then GoTo EH
  End If
  Let BBooComplex = (EnuBalType = eqlBalWithComplex)

  ' Return
  Let Balmet_core_main_udf_validate = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Let VecAux2 = Empty
  Let VecAux3 = Empty

  Let VecFeed = Empty
  Let VecGrad = Empty
  Let VecName = Empty
  Let VecUnit = Empty
  Let VecGradCx = Empty
  Let VecNameCx = Empty
  Let VecUnitCx = Empty

  Set MDL = Nothing

  #If Debugging Then
    Debug.Print VBA.Err.Description
  #End If

End Function

Private Function Balmet_core_main_udf_prepare( _
  ByRef BVarRequest As Variant, _
  ByRef BVarErrUDF As Variant, _
  ByVal BBooComplex As Boolean, _
  ByRef BEnuTypeFeed As EQLBAL_ENU_TYPECONTENT) As Boolean

  'Data required: (BVarFeed, BVarGrad, BVarUnits)
    'Data required: (BVarFeed, BVarGrad, BVarGradCx, BVarElem, BVarUnit)
  Dim MDL As ClassEqlModel
  Dim VecFeed As Variant, VecGrad As Variant, VecUnit As Variant, VecName As Variant, VecGradCx As Variant, VecNameCx As Variant, VecUnitCx As Variant, VecGradGL As Variant
  Dim VecAux0 As Variant, VecAux1 As Variant, VecAux2 As Variant, VecAux3 As Variant
  Dim StrAux0 As String
  Dim DblFeed As Double, DblFeedVol As Double
  Dim LngElements As Long, i As Long, j As Long, k As Long, l As Long, m As Long

  On Error GoTo EH
  Let Balmet_core_main_udf_prepare = False
  Set MDL = New ClassEqlModel
  Set MDL.ResourcesApp = RES

  ' Extract all balmet and cx grades
  Let VecGrad = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(1), j, RValueType:=eqlResNumbers, RBooAbs:=True, RValueAs:=eqlResNumbers, RDblTop:=CDBL_MAX)
  Let VecGradCx = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(2), k, RValueType:=eqlResNumbers, RBooAbs:=True, RValueAs:=eqlResNumbers, RDblTop:=CDBL_MAX)
  ' Get cols of products
  Let l = RES.RESCORE.BalGetProductsBounds(j, BWhat:=eqlBalColumns) ' Main Balmet cols
  Let m = RES.RESCORE.BalGetProductsBounds(j, k, BWhat:=eqlBalColumnsCx) ' Cx Balmet cols
  ' Get grades in 2D
  Let VecGrad = RES.RESCORE.BalGetGrades(VecGrad, j, BResponseDimension:=eqlBalAs2D)
  If BBooComplex Then
    Let LngElements = RES.RESCORE.BalGetProductsBounds(j, k, eqlBalFullProducts)
    Let VecGradCx = RES.RESCORE.BalGetGradesCx(VecGradCx, LngElements, k, BResponseDimension:=eqlBalAs2D)
  End If

  ' Names
  Let VecAux0 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalelements_sys, eqlMdlArray, "Símbolo"), eqlRes1D)
  Let VecAux1 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalelementscx_sys, eqlMdlArray, "Símbolo"), eqlRes1D)
  Let StrAux0 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(BVarRequest(3), j, RValueType:=eqlResNotNumbers, RBooUniques:=True)
  If Not RES.RESCORE.BalRefillNames2(RES, VecAux0, StrAux0, BLngGradCol:=(l - 1), BLngGradCxCol:=(m - 1), BVarName:=VecName, BVarNameCx:=VecNameCx, BStrErrMsg:=StrAux0, BBooCx:=BBooComplex, BVarDBNamesCx:=VecAux1) Then GoTo EH

  ' Units
  Let VecAux2 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalunits_sys, eqlMdlArray, "Unidad"), eqlRes1D)
  Let VecAux3 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalunitscx_sys, eqlMdlArray, "Unidad"), eqlRes1D)
  Let StrAux0 = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsString(BVarRequest(4), j, RValueType:=eqlResNotNumbers, RStrCustomRegex:="%")
  If Not RES.RESCORE.BalRefillUnits2(RES, VecAux2, StrAux0, BLngGradCol:=(l - 1), BLngGradCxCol:=(m - 1), BVarUnit:=VecUnit, BVarUnitCx:=VecUnitCx, BStrErrMsg:=StrAux0, BBooCx:=BBooComplex, BVarDBUnitsCx:=VecAux3) Then GoTo EH
  Let VecAux2 = Empty
  Let VecAux3 = Empty

  ' Feed
  Let VecFeed = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(0), i, RValueType:=eqlResNumbers, RBooAbs:=True, RValueAs:=eqlResNumbers, RDblTop:=CDBL_MAX)
  Let BEnuTypeFeed = RES.RESCORE.BalTypeFeed(RES, VBA.Join(VecUnit, ";"))
  If Not RES.RESCORE.BalGetFeed(BEnuTypeFeed, VecFeed, DblFeed, DblFeedVol) Then GoTo EH
  Let VecFeed = Empty

  ' Return
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
  Set MDL = Nothing

End Function

Private Function Balmet_core_main_udf_calculate( _
  ByVal BVarRequest As Variant, _
  ByVal BMethod As Byte, _
  ByVal BResult As Byte, _
  ByVal BDirection As EQLBAL_ENU_RESULTDIRECTION, _
  ByRef BVarErrUDF As Variant, _
  ByVal BAllResult As Boolean, _
  ByVal BEnuTypeFeed As EQLBAL_ENU_TYPECONTENT) As Boolean

  Dim VecWeight As Variant, VecVolume As Variant
  Dim DblFeed As Double, DblFeedVol As Double

  On Error GoTo EH
  Let Balmet_core_main_udf_calculate = False

  ' Data entry: (DblFeed, DblFeedVol, VecGrad, VecGradCx, VecName, VecUnit, VecNameCx, VecUnitCx, VecGradGL)
  Let DblFeed = VBA.CDbl(BVarRequest(0))
  Let DblFeedVol = VBA.CDbl(BVarRequest(1))
  If Not Balmet_core_main_app_spreadbalmet_get_weights(DblFeed, BVarRequest(2), BMethod, VecWeight, BFullData:=True) Then GoTo EH
  If BEnuTypeFeed = eqlBalBoth Then
    If Not Balmet_core_main_app_spreadbalmet_get_volumes(DblFeedVol, VecWeight, BVarRequest(2), BVarRequest(3), BVarRequest(5), VecVolume) Then GoTo EH
  End If

  ' Return
  Select Case BResult

    Case eqlBalWeights
      If Not BAllResult Then
        Call RES.ArrayPop(VecWeight)
        If Not VBA.IsArray(VecWeight) Then Call RES.ErrorHandler.ErrorHandlerSet(0.1, "¡Error al calcular los pesos del Balance Metalúrgico!"): GoTo EH 'Error
      End If
      Let GLOVEC_BALANCE = VecWeight ' eqlBalWeights [mx1]

    Case eqlBalWeightPercents: If Not Balmet_core_main_app_spreadbalmet_get_percents(VecWeight, GLOVEC_BALANCE) Then GoTo EH 'eqlBalWeightPercents [mx1]

    Case eqlBalVolume ' eqlBalVolume [mx1]
      If BEnuTypeFeed = eqlBalBoth Then
        Let GLOVEC_BALANCE = VecVolume
      Else
        Let GLOVEC_BALANCE = VecWeight
      End If

    Case eqlBalVolumePercents ' eqlBalVolumePercents [mx1]
      If BEnuTypeFeed = eqlBalBoth Then
        If Not Balmet_core_main_app_spreadbalmet_get_volumes_percents(VecVolume, GLOVEC_BALANCE) Then GoTo EH
      Else
        If Not Balmet_core_main_app_spreadbalmet_get_percents(VecWeight, GLOVEC_BALANCE, BEnuTypeFeedPer:=BEnuTypeFeed) Then GoTo EH
      End If

    Case eqlBalGradesHeads
      If Not Balmet_core_main_app_spreadbalmet_get_grades(VecWeight, VecVolume, BVarRequest(2), _
         BVarRequest(3), BVarRequest(5), BVarRequest(4), GLOVEC_BALANCE, BAllResult, _
         BBooHeads:=True, BBooJustHeads:=True) Then GoTo EH ' eqlBalGradesHeads [2xn]

    Case eqlBalUnities ' eqlBalUnities [mxn]
      If Not Balmet_core_main_app_spreadbalmet_get_unities(VecWeight, VecVolume, BVarRequest(2), _
         BVarRequest(3), BVarRequest(5), BVarRequest(4), GLOVEC_BALANCE, BAllResult) Then GoTo EH ' eqlBalGradesHeadsCx [2xn]

    Case eqlBalFines ' eqlBalFines [mxn]
      If Not Balmet_core_main_app_spreadbalmet_get_fines(VecWeight, VecVolume, BVarRequest(2), _
         BVarRequest(3), BVarRequest(5), BVarRequest(4), GLOVEC_BALANCE, BAllResult) Then GoTo EH ' eqlBalGradesHeadsCx [2xn]

    Case eqlBalRecoveries ' eqlBalRecoveries [mxn]
      If Not Balmet_core_main_app_spreadbalmet_get_recoveries(VecWeight, VecVolume, BVarRequest(2), _
         BVarRequest(3), BVarRequest(5), BVarRequest(4), GLOVEC_BALANCE, BAllResult) Then GoTo EH 'eqlBalGradesHeadsCx [2xn]

    Case eqlBalRatio ' eqlBalRatio [mx1]
      If Not Balmet_core_main_app_spreadbalmet_get_ratio(VecWeight, VecVolume, GLOVEC_BALANCE) Then GoTo EH

    Case eqlBalGradesHeadsCx
      If Not Balmet_core_main_app_spreadbalmet_get_grades(VecWeight, VecVolume, BVarRequest(2), _
         BVarRequest(3), BVarRequest(5), BVarRequest(4), GLOVEC_BALANCE, BAllResult, _
         BBooHeads:=True, BBooJustHeads:=True, BBooCx:=True) Then GoTo EH ' eqlBalGradesHeadsCx [2xn]

    Case eqlBalUnitiesCx ' eqlBalUnitiesCx [mxn]
      If Not Balmet_core_main_app_spreadbalmet_get_unities(VecWeight, VecVolume, BVarRequest(2), _
         BVarRequest(3), BVarRequest(5), BVarRequest(4), GLOVEC_BALANCE, BAllResult, _
         BBooCx:=True) Then GoTo EH ' eqlBalGradesHeadsCx [2xn]

    Case eqlBalFinesCx ' eqlBalFinesCx [mxn]
      If Not Balmet_core_main_app_spreadbalmet_get_fines(VecWeight, VecVolume, BVarRequest(2), _
         BVarRequest(3), BVarRequest(5), BVarRequest(4), GLOVEC_BALANCE, BAllResult, _
         BBooCx:=True) Then GoTo EH ' eqlBalGradesHeadsCx [2xn]

    Case eqlBalRecoveriesCx ' eqlBalRecoveriesCx [mxn]
      If Not Balmet_core_main_app_spreadbalmet_get_recoveries(VecWeight, VecVolume, BVarRequest(2), _
         BVarRequest(3), BVarRequest(5), BVarRequest(4), GLOVEC_BALANCE, BAllResult, _
         BBooCx:=True) Then GoTo EH ' eqlBalGradesHeadsCx [2xn]

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
    Select Case RES.ErrorHandler.ErrorNum
      Case 0.1, 1.11, 1.2: Let BVarErrUDF = VBA.CVErr(xlErrNA)
      Case 1.111, 1.114: Let BVarErrUDF = VBA.CVErr(xlErrDiv0)
      Case 1.1: Let BVarErrUDF = VBA.CVErr(xlErrNum)
      Case Else: Let BVarErrUDF = VBA.CVErr(xlErrValue)
      ' 0.2,1.113,1.3,1.4,2.1,2.11,2.2,3.1,3.2,4.1,4.2,5.1,5.2,5.3,5.4,6.1,6.2,6.3,7.1,7.2,7.3,7.4,8.1,8.2,9.1,9.2,9.3
    End Select
  End If

End Function

' APP CORE METHODS
Private Function Balmet_core_main_app(ByVal BVarRequest As Variant) As Boolean

  ' Data entry in BVarRequest: (0.DblFeed, 1.DblFeedVol, 2.VecGrad, 3.VecGradCx, 4.VecName, 5.VecUnit, 6.BytMethod, 7.Booleans, 8.EnuTypeFeed, 9.RngBox)
  Dim RngBox As Range
  Dim VecWeight As Variant, VecVolume As Variant, VecAux0 As Variant, VecAux1 As Variant
  Dim BooAll As Boolean, BooPercents As Boolean, BooUnities As Boolean, BooFines As Boolean, _
      BooGrams As Boolean, BooOT As Boolean, BooRatio As Boolean, BooHeads As Boolean
  Dim StrVol As String, StrName As String
  Dim LngAux0 As Long, LngPer As Long, LngGrad As Long, LngGrades As Long, LngCurrent As Long
  Dim DblFeed As Double, DblFeedVol As Double
  Dim EnuTypeFeed As EQLBAL_ENU_TYPECONTENT
  Dim EnuMethod As EQLBAL_ENU_METHOD

  Const sFN As String = "BALMET::-core_main_app_calculate"

  On Error GoTo EH
  Let Balmet_core_main_app = False

  ' Unpack data
  Let DblFeed = VBA.CDbl(VBA.Replace(VBA.CStr(BVarRequest(0)), ".", ","))
  Let DblFeedVol = VBA.CDbl(VBA.Replace(VBA.CStr(BVarRequest(1)), ".", ","))
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
  Let StrName = BVarRequest(10)
  'GoTo EH '<=====!



  ' CONVERT GRADES TO 2D VECTOR IF IT IS A STRING ******************************************************
  If VBA.TypeName(BVarRequest(2)) = "String" Then Let BVarRequest(2) = RES.RESCORE.BalGetGrades(VBA.Replace(VBA.CStr(BVarRequest(2)), ".", ","), LngAux0, BResponseDimension:=eqlBalAs2D) ' Grades
  If VBA.TypeName(BVarRequest(3)) = "String" Then
    If VBA.CStr(BVarRequest(3)) = VBA.vbNullString Then Let BVarRequest(3) = Empty
    If Not VBA.IsEmpty(BVarRequest(3)) Then Let BVarRequest(3) = RES.RESCORE.BalGetGradesCx(VBA.Replace(VBA.CStr(BVarRequest(3)), ".", ","), LngAux0, LngAux0, BResponseDimension:=eqlBalAs2D) ' GradesCx
  End If
  If VBA.TypeName(BVarRequest(4)) = "String" Then Let BVarRequest(4) = VBA.Split(BVarRequest(4), ";") ' Names
  If VBA.TypeName(BVarRequest(5)) = "String" Then Let BVarRequest(5) = VBA.Split(BVarRequest(5), ";") ' Units
  ' ****************************************************************************************************



  ' Get Balmet products in main array
  Let LngAux0 = (UBound(BVarRequest(2), 2) - LBound(BVarRequest(2), 2))
  If Not Balmet_core_main_app_spreadbalmet_get_products_names(BVarRequest(4), LngAux0, VecAux0, BooHeads, StrName) Then GoTo EH

  ' Get Balmet Contents with its optional Percent column in two arrays and concat to main array
  If Not Balmet_core_main_app_spreadbalmet_get_weights(DblFeed, BVarRequest(2), EnuMethod, VecWeight, BFullData:=True) Then GoTo EH
  If Not Balmet_core_main_app_spreadbalmet_get_weights(DblFeed, BVarRequest(2), EnuMethod, VecAux1, GLOBOO_FORMULAS, BooHeads, True, RngBox, EnuTypeFeed, True) Then GoTo EH
  Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1, RByCols:=True)
  If Not VBA.IsArray(VecAux0) Then Call RES.ErrorHandler.ErrorHandlerSet(1, "Fallo al unir resultados de Contenidos a la matriz de Balance Metalúrgico."): GoTo EH
  If BooPercents Then
    If Not Balmet_core_main_app_spreadbalmet_get_percents(VecWeight, VecAux1, GLOBOO_FORMULAS, BooHeads, True, EnuTypeFeed, RngBox) Then GoTo EH
    Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1, RByCols:=True)
    If Not VBA.IsArray(VecAux0) Then Call RES.ErrorHandler.ErrorHandlerSet(3, "Fallo al unir resultados de Porcentaje en Contenidos a la matriz de Balance Metalúrgico."): GoTo EH
  End If

  Let StrVol = IIf(EnuTypeFeed = eqlBalBoth, IIf(BooPercents, "E3", "D3"), "C3")
  If EnuTypeFeed = eqlBalBoth Then
    If Not Balmet_core_main_app_spreadbalmet_get_volumes(DblFeedVol, VecWeight, BVarRequest(2), BVarRequest(3), BVarRequest(5), VecVolume) Then GoTo EH
    If Not Balmet_core_main_app_spreadbalmet_get_volumes(DblFeedVol, VecWeight, BVarRequest(2), BVarRequest(3), BVarRequest(5), VecAux1, GLOBOO_FORMULAS, True, BooHeads, RngBox, BooPercents) Then GoTo EH
    Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1, RByCols:=True)
    If Not VBA.IsArray(VecAux0) Then Call RES.ErrorHandler.ErrorHandlerSet(2, "Fallo al unir resultados de Contenidos en volumen a la matriz de Balance Metalúrgico."): GoTo EH
    If BooPercents Then
      If Not Balmet_core_main_app_spreadbalmet_get_volumes_percents(VecVolume, VecAux1, GLOBOO_FORMULAS, BooHeads, True, RngBox) Then GoTo EH
      Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1, RByCols:=True)
      If Not VBA.IsArray(VecAux0) Then Call RES.ErrorHandler.ErrorHandlerSet(4, "Fallo al unir resultados de Porcentaje en Volumen a la matriz de Balance Metalúrgico."): GoTo EH
    End If
  End If

  ' Get the rest of products and concat to main array
  Let LngGrades = IIf(EnuTypeFeed = eqlBalBoth, IIf(BooPercents, 4, 2), IIf(BooPercents, 2, 1))
  Let LngCurrent = (UBound(VecAux0, 2) - LBound(VecAux0, 2)) - 1
  If Not Balmet_core_main_app_spreadbalmet_get_grades(VecWeight, VecVolume, BVarRequest(2), BVarRequest(3), BVarRequest(5), BVarRequest(4), VecAux1, True, GLOBOO_FORMULAS, True, BooHeads, RngBox, False, False, BooPercents, EnuTypeFeed, StrVol, LngGrades, LngCurrent) Then GoTo EH 'eqlBalGradesHeads [2xn]
  Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1, RByCols:=True)
  If Not VBA.IsArray(VecAux0) Then Call RES.ErrorHandler.ErrorHandlerSet(5, "Fallo al unir resultados de Leyes a la matriz de Balance Metalúrgico."): GoTo EH
  Let LngCurrent = (UBound(VecAux0, 2) - LBound(VecAux0, 2)) - 1

  If BooUnities Then
    If Not Balmet_core_main_app_spreadbalmet_get_unities(VecWeight, VecVolume, BVarRequest(2), BVarRequest(3), BVarRequest(5), BVarRequest(4), VecAux1, True, GLOBOO_FORMULAS, True, BooHeads, RngBox, False, BooPercents, EnuTypeFeed, StrVol, LngGrades, LngCurrent) Then GoTo EH 'eqlBalGradesHeadsCx [2xn]
    Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1, RByCols:=True)
    If Not VBA.IsArray(VecAux0) Then Call RES.ErrorHandler.ErrorHandlerSet(6, "Fallo al unir resultados de Unidades a la matriz de Balance Metalúrgico."): GoTo EH
    Let LngCurrent = (UBound(VecAux0, 2) - LBound(VecAux0, 2)) - 1
  End If

  If BooFines Then
  ' Data entry in BVarRequest: (0.DblFeed, 1.DblFeedVol, 2.VecGrad, 3.VecGradCx, 4.VecName, 5.VecUnit, 6.BytMethod, 7.Booleans, 8.EnuTypeFeed, 9.RngBox)
    If Not Balmet_core_main_app_spreadbalmet_get_fines(VecWeight, VecVolume, BVarRequest(2), BVarRequest(3), BVarRequest(5), BVarRequest(4), VecAux1, True, GLOBOO_FORMULAS, True, BooHeads, RngBox, False, BooGrams, BooOT, BooPercents, EnuTypeFeed, StrVol, LngGrades, LngCurrent) Then GoTo EH 'eqlBalGradesHeadsCx [2xn]
    Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1, RByCols:=True)
    If Not VBA.IsArray(VecAux0) Then Call RES.ErrorHandler.ErrorHandlerSet(7, "Fallo al unir resultados de Finos a la matriz de Balance Metalúrgico."): GoTo EH
    Let LngCurrent = (UBound(VecAux0, 2) - LBound(VecAux0, 2)) - 1
  End If

  If Not Balmet_core_main_app_spreadbalmet_get_recoveries(VecWeight, VecVolume, BVarRequest(2), BVarRequest(3), BVarRequest(5), BVarRequest(4), VecAux1, True, GLOBOO_FORMULAS, True, BooHeads, RngBox, False, BooPercents, EnuTypeFeed, StrVol, LngGrades, LngCurrent) Then GoTo EH 'eqlBalGradesHeadsCx [2xn]
  Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1, RByCols:=True)
  If Not VBA.IsArray(VecAux0) Then Call RES.ErrorHandler.ErrorHandlerSet(8, "Fallo al unir resultados de Recuperaciones a la matriz de Balance Metalúrgico."): GoTo EH
  Let LngCurrent = (UBound(VecAux0, 2) - LBound(VecAux0, 2)) - 1

  If BooRatio Then
    If Not Balmet_core_main_app_spreadbalmet_get_ratio(VecWeight, VecVolume, VecAux1, GLOBOO_FORMULAS, True, BooHeads, RngBox, StrVol) Then GoTo EH
    Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1, RByCols:=True)
    If Not VBA.IsArray(VecAux0) Then Call RES.ErrorHandler.ErrorHandlerSet(9, "Fallo al unir resultados de Ratio a la matriz de Balance Metalúrgico."): GoTo EH
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
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

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

  ' Obtains the Feed Grade from one or mxn matrix of concentrates (for more or equal than two concentrates, Feed Grade will be an accourate) _
    Requirements: _
    1. BVarFeed:   Feed, the weight feed of the process in Tons _
    2. BVarRecs:   Concentrates Recuperations, a number or mxn matrix of Doubles _
    3. BVarWeig:Concentrates Weights, a number or mxn matrix of Doubles in Tons _
    4. BVarGrad: Balmet Concentrates Grades, the mxn matrix of concentrates, %, DM, g/L or g/T _
    The Method returns a Double or a matrix of Doubles with the value or values of the Feed(s) Grade(s).

  Dim VecAux0 As Variant, VecAux1 As Variant
  Dim BooComplex As Boolean, BooError As Boolean
  Dim EnuTypeFeed As EQLBAL_ENU_TYPECONTENT

  On Error GoTo EH
  Let Balmet_core_main_udf_eco = VBA.CVErr(xlErrValue)
  Let BooError = False

  ' Forming the request data
  Let VecAux0 = Array(BVarFeed, BVarConc, BVarGrad, BVarReco, BVarGradCx, BVarRecoCx, BVarUnit)

  ' Get final result
  If Not Balmet_core_main_udf_validate_eco(VecAux0, Balmet_core_main_udf_eco) Then GoTo EH
  If Not Balmet_core_main_udf_prepare_eco(VecAux0, VecAux1, Balmet_core_main_udf_eco, BooComplex) Then GoTo EH
  If Not Balmet_core_main_udf_calculate_eco(VecAux1, VecAux0, BooComplex, BResult, BDirection) Then GoTo EH

  ' Return
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
  If Not BooError Then
    Select Case RES.ErrorHandler.ErrorNum
      Case 0.1: Let Balmet_core_main_udf_eco = VBA.CVErr(xlErrDiv0)
      Case 0.2: Let Balmet_core_main_udf_eco = VBA.CVErr(xlErrNum)
    End Select
  End If

End Function

Private Function Balmet_core_main_udf_validate_eco( _
  ByVal BVarRequest As Variant, _
  ByRef BVarErrUDF As Variant) As Boolean

  ' Data required: (BVarFeed, BVarConc, BVarGrad, BVarReco, BVarGradCx, BVarRecoCx, BVarUnit)
  Dim MDL As ClassEqlModel
  Dim VecFeed As Variant, VecWeights As Variant, VecGrades As Variant, VecRecs As Variant, VecGradesCx As Variant, VecRecsCx As Variant, VecUnits As Variant, VecUnitsCx As Variant, _
      VecAux0 As Variant, VecName As Variant, VecNameCx As Variant
  Dim StrAux0 As String
  Dim BooBalCx As Boolean, BooGL As Boolean

  Dim DblFeedSolid As Double, DblFeedVolume As Double, DblSum As Double, DblAux0 As Double, DblAux1 As Double
  Dim i As Long, j As Long, x1 As Long, x2 As Long, x3 As Long, x4 As Long, x5 As Long, x6 As Long

  Const sFN As String = "BALMET::-core_main_udf_validation_eco"

  On Error GoTo EH
  Let Balmet_core_main_udf_validate_eco = False
  Set MDL = New ClassEqlModel
  Set MDL.ResourcesApp = RES

  ' All incoming main values are strings, requires RES method getVariantsAsStrings
  Let VecFeed = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(0), x1, RValueType:=eqlResNumbers, RBooAbs:=True, RValueAs:=eqlResNumbers, RDblTop:=CDBL_MAX)
  Let VecWeights = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(1), x2, RValueType:=eqlResNumbers, RBooAbs:=True, RValueAs:=eqlResNumbers, RDblTop:=CDBL_MAX)
  Let VecGrades = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(2), x3, RValueType:=eqlResNumbers, RBooAbs:=True, RValueAs:=eqlResNumbers, RDblTop:=CDBL_MAX)
  Let VecRecs = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(3), x4, RValueType:=eqlResNumbers, RBooAbs:=True, RValueAs:=eqlResNumbers, RDblTop:=CDBL_MAX)
  Let VecGradesCx = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(4), x5, RValueType:=eqlResNumbers, RBooAbs:=True, RValueAs:=eqlResNumbers, RDblTop:=CDBL_MAX)
  Let VecRecsCx = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(5), x6, RValueType:=eqlResNumbers, RBooAbs:=True, RValueAs:=eqlResNumbers, RDblTop:=CDBL_MAX)
  Let VecUnits = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(6), i, RValueType:=eqlResStrings, RStrCustomRegex:="%", RValueAs:=eqlResStrings)
  If x1 = 0 Then Let BVarErrUDF = VBA.CVErr(xlErrNull): GoTo EH ' Feeds required
  If x2 = 0 Then Let BVarErrUDF = VBA.CVErr(xlErrNull): GoTo EH ' Weights required
  If x2 > CLNG_BAL_LIMIT Then Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH ' Just 5 products as limit of balmets and One product at least
  If x3 = 0 Or Not x3 = x2 Then Let BVarErrUDF = VBA.CVErr(xlErrNull): GoTo EH ' Grades required
  If x5 > CLNG_BAL_LIMITCX Then Let BVarErrUDF = VBA.CVErr(xlErrNA): GoTo EH ' Just 3 Cx avaiable
  Let BooBalCx = (x5 > 0)

  ' Names & NamesCx
  If Not RES.RESCORE.BalRefillNames(RES, VBA.vbNullString, BLngGradCol:=(x3 - 1), BLngGradCxCol:=(x5 - 1), BVarName:=VecName, BVarNameCx:=VecNameCx, BBooCx:=BooBalCx) Then GoTo EH
  ' Units & UnitsCx
  Let StrAux0 = IIf(i > 0, VBA.Join(VecUnits, ";"), VBA.vbNullString)
  If Not RES.RESCORE.BalRefillUnits(RES, StrAux0, BLngGradCol:=(x3 - 1), BLngGradCxCol:=(x5 - 1), BVarUnit:=VecUnits, BVarUnitCx:=VecUnitsCx, BBooCx:=BooBalCx) Then GoTo EH

  Let VecAux0 = RES.ArrayToBaseZero(MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalelementscx_sys, eqlMdlArray, "Símbolo"), eqlRes1D)
  If Not RES.RESCORE.BalCheckCustomNamesDoesNotHasWrongUnits(RES, StrAux0, BVarSrcElems:=VecAux0, BVarName:=VecName, BVarUnit:=VecUnits, BEnuCheckWhat:=eqlBalCxElemsUnitOnNormalBalmet) Then GoTo EH
  If Not RES.RESCORE.BalCheckCustomNamesDoesNotHasWrongUnits(RES, StrAux0, BVarName:=VecAux0, BVarUnit:=VecName, BEnuCheckWhat:=eqlBalDMOnOtherElements) Then GoTo EH
  ' Check for maximous values in grades
  If Not RES.RESCORE.BalCheckForMaximousValuesInGrades(StrAux0, RES.ArrayToBaseZero(VecGrades, eqlRes2D), VecUnits, BBooOneToOne:=True) Then GoTo EH
  ' Check for maximous values in recs
  If Application.WorksheetFunction.Max(VecRecs) >= 100 Then GoTo EH
  ' Check for volumes in feeds
  Dim EnuTypeFeed As EQLBAL_ENU_TYPECONTENT
  Let EnuTypeFeed = RES.RESCORE.BalTypeFeed(RES, VBA.Join(VecUnits, ";")) ' Feed values
  If Not RES.RESCORE.BalCheckForVolumesAndSolidsFeeds(RES, VecFeed, VecWeights, VecUnits, EnuTypeFeed, CDBL_MAX, StrAux0) Then GoTo EH

  ' Cx:
  If BooBalCx Then
    ' Grades, Names, Units & Recs lengths
    If Not x5 = x6 Then GoTo EH
    ' Namescx vs Unitscx
    If Not RES.RESCORE.BalCheckCustomNamesDoesNotHasWrongUnits(RES, StrAux0, BVarSrcElems:=Empty, BVarName:=VecNameCx, BVarUnit:=VecUnitsCx, BEnuCheckWhat:=eqlBalDMOnOtherElements) Then GoTo EH
    If Not RES.RESCORE.BalCheckCustomNamesDoesNotHasWrongUnits(RES, StrAux0, BVarName:=VecNameCx, BVarUnit:=VecUnitsCx, BEnuCheckWhat:=eqlBalNonCx) Then GoTo EH
    ' Check for maximous values in gradescx
    If Not RES.RESCORE.BalCheckForMaximousValuesInGrades(StrAux0, RES.ArrayToBaseZero(VecGrades, eqlRes2D), VecUnits, RES.ArrayToBaseZero(VecGradesCx, eqlRes2D), VecUnitsCx, BBooCx:=True, BBooOneToOne:=True) Then GoTo EH
    ' Check for maximous values in recs
    If Application.WorksheetFunction.Max(VecRecsCx) >= 100 Then GoTo EH
  End If
  Let VecAux0 = Empty

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
  Set MDL = Nothing

End Function

Private Function Balmet_core_main_udf_prepare_eco( _
  ByVal BVarRequest As Variant, _
  ByRef BVarResponse As Variant, _
  ByRef BVarErrUDF As Variant, _
  ByRef BBooCx As Boolean) As Boolean

  'Data required: (BVarFeed, BVarConc, BVarGrad, BVarReco, BVarGradCx, BVarRecoCx, BVarUnit)
  Dim VecFeed As Variant, VecWeights As Variant, VecGrades As Variant, VecRecs As Variant, VecGradesCx As Variant, VecRecsCx As Variant, VecUnits As Variant, VecUnitsCx As Variant
  Dim StrAux0 As String

  Dim DblFeedSolid As Double, DblFeedVolume As Double, DblSum As Double
  Dim i As Long, j As Long, x1 As Long, x2 As Long, x3 As Long, x4 As Long, x5 As Long, x6 As Long

  Const sFN As String = "BALMET::-core_main_udf_prepare_eco"

  On Error GoTo EH
  Let Balmet_core_main_udf_prepare_eco = False

  ' All incoming main values are strings, requires RES method getVariantsAsStrings
  Let VecFeed = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(0), x1, RValueType:=eqlResNumbers, RBooAbs:=True, RValueAs:=eqlResNumbers, RDblTop:=CDBL_MAX)
  Let VecWeights = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(1), x2, RValueType:=eqlResNumbers, RBooAbs:=True, RValueAs:=eqlResNumbers, RDblTop:=CDBL_MAX)
  Let VecGrades = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(2), x3, RValueType:=eqlResNumbers, RBooAbs:=True, RValueAs:=eqlResNumbers, RDblTop:=CDBL_MAX)
  Let VecRecs = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(3), x4, RValueType:=eqlResNumbers, RBooAbs:=True, RValueAs:=eqlResNumbers, RDblTop:=CDBL_MAX)
  Let VecGradesCx = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(4), x5, RValueType:=eqlResNumbers, RBooAbs:=True, RValueAs:=eqlResNumbers, RDblTop:=CDBL_MAX)
  Let VecRecsCx = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(5), x6, RValueType:=eqlResNumbers, RBooAbs:=True, RValueAs:=eqlResNumbers, RDblTop:=CDBL_MAX)
  Let VecUnits = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(BVarRequest(6), i, RValueType:=eqlResStrings, RStrCustomRegex:="%", RValueAs:=eqlResStrings)
  Let BBooCx = (x5 > 0)

  ' Units & UnitsCx
  Let StrAux0 = VBA.vbNullString
  If i > 0 Then Let StrAux0 = VBA.Join(VecUnits, ";")
  If Not RES.RESCORE.BalRefillUnits(RES, StrAux0, BLngGradCol:=(x3 - 1), BLngGradCxCol:=(x5 - 1), BVarUnit:=VecUnits, BVarUnitCx:=VecUnitsCx, BBooCx:=BBooCx) Then GoTo EH
  If Not RES.RESCORE.BalRefillRecs(RES, BLngGradCol:=(x3 - 1), BLngGradColcx:=(x5 - 1), BVarRec:=VecRecs, BVarRecCx:=VecRecsCx, BBooCx:=BBooCx) Then GoTo EH

  ' Feeds
  Dim EnuTypeFeed As EQLBAL_ENU_TYPECONTENT
  If VBA.IsArray(VecUnitsCx) Then Let StrAux0 = VBA.Join(VecUnitsCx, ";")
  Let StrAux0 = VBA.Join(VecUnits, ";") & IIf(BBooCx, ";" & StrAux0, VBA.vbNullString)
  Let EnuTypeFeed = RES.RESCORE.BalTypeFeed(RES, StrAux0) ' Feed values
  If Not RES.RESCORE.BalGetFeed(EnuTypeFeed, VecFeed, DblFeedSolid, DblFeedVolume) Then GoTo EH

  ' Return
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
  ByVal BEnuTypeBalEco As EQLBAL_ENU_RESULTECO, _
  ByVal BEnuOrientation As EQLBAL_ENU_RESULTDIRECTION) As Boolean

  ' Data required: (DblFeedSolid, DblFeedVolume), VecWeights, VecGrades, VecRecs, VecGradesCx, VecRecsCx, VecUnits, VecUnitsCx)
  Dim DblFeed As Double, DblFeedSol As Double, DblFeedVol As Double
  Dim VecWeig As Variant, VecGrad As Variant, VecRecs As Variant, VecGradCx As Variant, VecRecsCx As Variant, VecUnit As Variant, VecUnitCx As Variant
  Dim VecResu As Variant, VecResuCx As Variant, VecAux0 As Variant, VecAux1 As Variant
  Dim StrAux0 As String
  Dim i As Long, j As Long
  Dim DblAux1 As Double, DblAux2 As Double ', DblAux3 As Double

  Const sFN As String = "BALMET::-core_main_udf_calculate_eco"

  On Error GoTo EH
  Let Balmet_core_main_udf_calculate_eco = False

  ' Unpacking all datas
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

  ' RESOLVE BALMET
  ' Resolve the Balmet base
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
    If DblAux2 = 0 Then Call RES.ErrorHandler.ErrorHandlerSet(0.1, "División entre cero. Alimentación o Recuperaciones erróneas."): GoTo EH 'Error Div0!
    Let VecAux0(i) = DblAux1 / DblAux2
  Next i

  ' Build the final data
  ReDim VecResu(0 To 0, 0 To UBound(VecAux0))
  For j = 0 To UBound(VecAux0) ' Just one row (horizontal)
    Let VecResu(0, j) = VecAux0(j)
  Next j
  Let VecAux0 = Empty

  ' Resolve Complex product if it exists
  If BBooCx Then
    ReDim VecAux0(UBound(VecGradCx))
    For i = LBound(VecGradCx) To UBound(VecGradCx)
      Let StrAux0 = VBA.LCase(VBA.CStr(VecUnitCx(i)))
      Let DblFeed = IIf(StrAux0 = "g/l", DblFeedVol, DblFeedSol)
      If BEnuTypeBalEco = eqlBalHeadsGrades Then ' = 100Kk / A%R
        Let DblAux1 = VBA.CDbl(VecWeig(i)) * VBA.CDbl(VecGradCx(i)) * 100
        Let DblAux2 = DblFeed * VBA.CDbl(VecRecsCx(i))
      Else ' = Aa%R / 100K
        Let DblAux1 = VBA.CDbl(VecRecsCx(i)) * DblFeed * VBA.CDbl(VecGradCx(i))
        Let DblAux2 = VBA.CDbl(VecWeig(i)) * 100
      End If
      If DblAux2 = 0 Then Call RES.ErrorHandler.ErrorHandlerSet(0.1, "División entre cero. Alimentación o Recuperaciones erróneas en valores de mineral complejo."): GoTo EH 'Error Div0!
      ReDim Preserve VecAux0(i)
      Let VecAux0(i) = DblAux1 / DblAux2
    Next i
    ' Build the final Cx data
    ReDim VecResuCx(0 To 0, 0 To UBound(VecAux0))
    For i = 0 To UBound(VecAux0) ' Just one row (horizontal)
      Let VecResuCx(0, i) = VecAux0(i)
    Next i
    Let VecAux0 = Empty
  End If

  ' Return
  If BEnuOrientation = eqlBalVertical Then
    Let VecResu = RES.ArrayTranspose(VecResu)
    If Not VBA.IsArray(VecResu) Then Call RES.ErrorHandler.ErrorHandlerSet(0.2, "Fallo al orientar verticalmente el resultado."): GoTo EH
    If BBooCx Then
      Let VecResuCx = RES.ArrayTranspose(VecResuCx)
      If Not VBA.IsArray(VecResu) Then Call RES.ErrorHandler.ErrorHandlerSet(0.2, "Fallo al orientar verticalmente el resultado."): GoTo EH
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

' APP CORE METHODS - ECO
Private Function Balmet_core_main_app_eco(ByRef BVarRequest As Variant, ByVal BEnuTypeBalEco As EQLBAL_ENU_RESULTECO) As Boolean

  ' Data entry in BVarRequest: (0.DblFeed, 1.DblFeedVol, 2.VecGrad, 3.VecGradCx, 4.VecName, 5.VecUnit, 6.EnuMethod, 7.Booleans, 8.EnuTypeFeed, 9.RngBox (, 10.StrFeedName, 11.StrNameA, 12.StrUnitA, 13. BooNewSheet, 14.BEnuTypeBalEco, 15.VecEcos))
  ' VecEcos: Array(0.Array(DblFeedSolid, DblFeedVolume), 1.VecWeights, 2.VecGrades, 3.VecRecs, 4.VecGradesCx, 5.VecRecsCx, 6.VecUnits, 7.VecUnitsCx, 8.BEnuTypeBalEco)
  Dim VecEcos As Variant, VecAux0 As Variant, VecAux1 As Variant
  Dim BooCx As Boolean
  Dim DblFeed As Double, DblFeedVol As Double
  Dim EnuFeedType As EQLBAL_ENU_TYPECONTENT

  Const sFN As String = "BALMET::-core_main_app_eco"

  On Error GoTo EH
  Let Balmet_core_main_app_eco = False

  Let VecEcos = RES.ArrayPop(BVarRequest) ' FIXME:
  Let VecEcos = VecEcos(0) ' Ecos Vector
  Let DblFeed = VBA.CDbl(BVarRequest(0)) ' Feed
  Let DblFeedVol = VBA.CDbl(BVarRequest(1)) ' Feed Vol
  Let EnuFeedType = VBA.CLng(BVarRequest(8)) ' Type Feed

  Let BooCx = VBA.IsArray(VecEcos(4))

  ' Get headsgrades or prodgrades from udf_calculation_eco. Response: Array(VecGrad/Head, VecGradCx/HeadCx)
  If Not Balmet_core_main_udf_calculate_eco(BVarRequest:=VecEcos, BVarResponse:=VecAux0, BBooCx:=BooCx, BEnuTypeBalEco:=BEnuTypeBalEco, BEnuOrientation:=eqlBalHorizontal) Then GoTo EH

  ' Transform results in vil 1D arrays
  Let VecAux0(0) = RES.ArrayToBaseZero(VecAux0(0), eqlRes1D)
  Let VecAux0(1) = RES.ArrayToBaseZero(VecAux0(1), eqlRes1D)

  ' Gets hybrid percents in two vectors of Solids or Volumes or Solids and Volumes. Response: Array(VecSolids, VecVolumes)
  If Not Balmet_core_main_app_eco_spreadbalmet_get_percents_hybrid_data(DblFeed, DblFeedVol, BVarProd:=VecEcos(1), BVarUnit:=VecEcos(6), BEnuFeedType:=EnuFeedType, BVarResp:=VecAux1) Then GoTo EH

  ' Get whole grades matrix from corresponding method
  If BEnuTypeBalEco = eqlBalHeadsGrades Then
    If Not Balmet_core_main_app_eco_spreadbalmet_get_grades_hybrid_data(BVarPerc:=VecAux1(0), BVarPercVol:=VecAux1(1), BVarGrad:=VecEcos(2), BVarHead:=VecAux0(0), BVarUnit:=VecEcos(6), BVarGradCx:=VecEcos(4), BVarHeadCx:=VecAux0(1), BVarUnitCx:=VecEcos(7), BBooCx:=BooCx, BEnuFeedType:=EnuFeedType, BResponse:=VecAux0) Then GoTo EH
  Else
    If Not Balmet_core_main_app_eco_spreadbalmet_get_grades_hybrid_data(BVarPerc:=VecAux1(0), BVarPercVol:=VecAux1(1), BVarGrad:=VecAux0(0), BVarHead:=VecEcos(2), BVarUnit:=VecEcos(6), BVarGradCx:=VecAux0(1), BVarHeadCx:=VecEcos(4), BVarUnitCx:=VecEcos(7), BBooCx:=BooCx, BEnuFeedType:=EnuFeedType, BResponse:=VecAux0) Then GoTo EH
  End If

  ' Calculate Balance
  Let BVarRequest(2) = VecAux0(0) ' Grades
  Let BVarRequest(3) = VecAux0(1) ' GradesCx
  Let VecAux0 = Empty
  Let VecEcos = Empty
  If Not Balmet_core_main_app(BVarRequest) Then GoTo EH

  Let Balmet_core_main_app_eco = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Let VecEcos = Empty
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Balmet_core_main_app_eco_spreadbalmet_get_grades_hybrid_data( _
  ByVal BVarPerc As Variant, _
  ByVal BVarPercVol As Variant, _
  ByVal BVarGrad As Variant, _
  ByVal BVarHead As Variant, _
  ByVal BVarUnit As Variant, _
  ByVal BVarGradCx As Variant, _
  ByVal BVarHeadCx As Variant, _
  ByVal BVarUnitCx As Variant, _
  ByVal BBooCx As Boolean, _
  ByVal BEnuFeedType As EQLBAL_ENU_TYPECONTENT, _
  ByRef BResponse As Variant) As Boolean

  ' Builds the Balmet grades matrix from 1D param arrays _
    Feed: System Feed _
    Perc: Whole Percents of balance, Prods, Tails and Feed (Solids|Volumes or Solids&Volumes) _
    Grad: All main diagonal Prod Grades (GradCx too if it exists) _
    Head: All Head Grades (HeadCx too if it exists)

  ' Average unities calculated to fill void spaces in Balmet grades
  Dim VecAux0 As Variant, VecUi As Variant, VecHeadsUi As Variant, VecUiZ As Variant, VecSemiFinalGrades As Variant, VecGradesCxSemi As Variant
  Dim StrGradeCx As String, StrHeadGradeCx As String, StrFinalGrades As String, StrFinalGradesCx As String
  Dim BooIsVol As Boolean
  Dim DblAux1 As Double
  Dim i As Long, j As Long, k As Long, m As Long, n As Long

  Const sFN = "BALMET::-core_main_spreadbalmet_get_grades_data_eco"

  On Error GoTo EH
  Let Balmet_core_main_app_eco_spreadbalmet_get_grades_hybrid_data = False

  ' Objetive: _
    1. Generate U with %P * Grades _
    2. Generate U* = U[Feeds] - U[Grade] _
    3. Generate U[i] with U*/(Products-1) _
    4. Generate Grades*

  Let BVarUnit = RES.ArrayToBaseZero(BVarUnit, eqlRes1D)

  If BEnuFeedType = eqlBalBoth Then

    ' Calculate Unities
    ReDim VecUi(UBound(BVarGrad) - LBound(BVarGrad))
    For i = LBound(VecUi) To UBound(VecUi)
      Let BooIsVol = (VBA.LCase(VBA.CStr(BVarUnit(i))) = "g/l")
      If BooIsVol Then
        Let VecUi(i) = VBA.CDbl(BVarPercVol(i)) * VBA.CDbl(BVarGrad(i)) ' Just products 'cause Tails grades doesn't exists
      Else
        Let VecUi(i) = VBA.CDbl(BVarPerc(i)) * VBA.CDbl(BVarGrad(i)) ' Just products 'cause Tails grades doesn't exists
      End If
    Next i

    ' Calculate Head Unities
    ReDim VecHeadsUi(UBound(BVarHead) - LBound(BVarHead))
    For k = LBound(BVarHead) To UBound(BVarHead)
      Let BooIsVol = (VBA.LCase(VBA.CStr(BVarUnit(i))) = "g/l")
      If BooIsVol Then
        Let VecHeadsUi(k) = VBA.CDbl(BVarHead(k)) * VBA.CDbl(BVarPercVol(UBound(BVarPercVol)))
      Else
        Let VecHeadsUi(k) = VBA.CDbl(BVarHead(k)) * VBA.CDbl(BVarPerc(UBound(BVarPerc)))
      End If
    Next k

    ' Calculate Ui* constants
    Let k = (UBound(BVarHead) - LBound(BVarHead)) + 1 ' Products with no grades
    ReDim VecUiZ(UBound(BVarHead) - LBound(BVarHead))
    For i = LBound(VecUiZ) To UBound(VecUiZ)
      Let VecUiZ(i) = (VBA.CDbl(VecHeadsUi(i)) - VBA.CDbl(VecUi(i))) / k ' k is total products including Tails
    Next i
    Let m = (UBound(BVarPerc) - LBound(BVarPerc)) ' Includes: Products, Feed, Tails
    Let n = (UBound(BVarUnit) - LBound(BVarUnit))
    ReDim VecAux0(m)
    ReDim VecSemiFinalGrades(n) ' Save each column of grade as string semicolon separated to obtain finally a unique string of grades
    For j = 0 To n
      Let BooIsVol = (VBA.LCase(VBA.CStr(BVarUnit(j))) = "g/l")
      For i = 0 To m
        If i = j Then ' Main diagonal Grades
          If BooIsVol Then
            Let VecAux0(i) = VBA.CDbl(VecUi(i)) / VBA.CDbl(BVarPercVol(i))
          Else
            Let VecAux0(i) = VBA.CDbl(VecUi(i)) / VBA.CDbl(BVarPerc(i))
          End If
        ElseIf i = m Then ' Heads
          If BooIsVol Then
            Let VecAux0(i) = VBA.CDbl(VecHeadsUi(j)) / VBA.CDbl(BVarPercVol(i))
          Else
            Let VecAux0(i) = VBA.CDbl(VecHeadsUi(j)) / VBA.CDbl(BVarPerc(i))
          End If
        Else ' The rest of grades
          If BooIsVol Then
            Let VecAux0(i) = VBA.CDbl(VecUiZ(j)) / VBA.CDbl(BVarPercVol(i))
          Else
            Let VecAux0(i) = VBA.CDbl(VecUiZ(j)) / VBA.CDbl(BVarPerc(i))
          End If
        End If
        If VBA.CDbl(VecAux0(i)) <= 0 Then GoTo EH ' #0 in grades!
      Next i
      Let VecSemiFinalGrades(j) = VBA.Join(VecAux0, ";")
    Next j
    Let StrFinalGrades = VBA.Join(VecSemiFinalGrades, ";")

    ' Calculate GradesCx
    If BBooCx Then
      ReDim VecGradesCxSemi(m)
      For j = 0 To n
        Let BooIsVol = (VBA.LCase(VBA.CStr(BVarUnitCx(j))) = "g/l")

        For i = 0 To UBound(VecGradesCxSemi)
          If BooIsVol Then
            Let DblAux1 = (VBA.CDbl(BVarGradCx(j)) * VBA.CDbl(BVarPercVol(m))) - (VBA.CDbl(BVarHeadCx(j)) * VBA.CDbl(BVarPercVol(m))) / k ' U1*
          Else
            Let DblAux1 = (VBA.CDbl(BVarGradCx(j)) * VBA.CDbl(BVarPerc(m))) - (VBA.CDbl(BVarHeadCx(j)) * VBA.CDbl(BVarPerc(m))) / k ' U1*
          End If

          If i = 0 Then ' Main Grade
            Let VecGradesCxSemi(i) = BVarGradCx(j)
          ElseIf i = m Then ' HeadGrade
            Let VecGradesCxSemi(i) = BVarHeadCx(j)
          Else
            If BooIsVol Then
              Let VecGradesCxSemi(i) = DblAux1 / VBA.CDbl(BVarPercVol(i))
            Else
              Let VecGradesCxSemi(i) = DblAux1 / VBA.CDbl(BVarPerc(i))
            End If
          End If
          If VBA.CDbl(VecGradesCxSemi(i)) <= 0 Then GoTo EH ' #0 in gradescx!
        Next i
        Let StrFinalGradesCx = VBA.Join(VecGradesCxSemi, ";")
      Next j
    End If

  Else ' Solids|Volumes or Solids&Volumes

    ' Calculate Unities
    ReDim VecUi(UBound(BVarGrad) - LBound(BVarGrad))
    For i = LBound(VecUi) To UBound(VecUi)
      Let VecUi(i) = VBA.CDbl(BVarPerc(i)) * VBA.CDbl(BVarGrad(i)) ' Just products 'cause Tails grades doesn't exists
    Next i

    ' Calculate Head Unities
    ReDim VecHeadsUi(UBound(BVarHead) - LBound(BVarHead))
    For k = LBound(BVarHead) To UBound(BVarHead)
      Let VecHeadsUi(k) = VBA.CDbl(BVarHead(k)) * VBA.CDbl(BVarPerc(UBound(BVarPerc)))
    Next k

    ' Calculate Ui* constants
    Let k = (UBound(BVarHead) - LBound(BVarHead)) + 1 ' Products with no grades
    ReDim VecUiZ(UBound(BVarHead) - LBound(BVarHead))
    For i = LBound(VecUiZ) To UBound(VecUiZ)
      Let VecUiZ(i) = (VBA.CDbl(VecHeadsUi(i)) - VBA.CDbl(VecUi(i))) / k ' k is total products including Tails
    Next i
    Let m = (UBound(BVarPerc) - LBound(BVarPerc)) ' Includes: Products, Feed, Tails
    Let n = (UBound(BVarUnit) - LBound(BVarUnit))
    ReDim VecAux0(m)
    ReDim VecSemiFinalGrades(n) ' Save each column of grade as string semicolon separated to obtain finally a unique string of grades
    For j = 0 To n
      For i = 0 To m
        If i = j Then ' Main diagonal Grades
          Let VecAux0(i) = VBA.CDbl(VecUi(i)) / VBA.CDbl(BVarPerc(i))
        ElseIf i = m Then ' Heads
          Let VecAux0(i) = VBA.CDbl(VecHeadsUi(j)) / VBA.CDbl(BVarPerc(i))
        Else ' The rest of grades
          Let VecAux0(i) = VBA.CDbl(VecUiZ(j)) / VBA.CDbl(BVarPerc(i))
        End If
        If VBA.CDbl(VecAux0(i)) <= 0 Then GoTo EH ' #0 in grades!
      Next i
      Let VecSemiFinalGrades(j) = VBA.Join(VecAux0, ";")
    Next j
    Let StrFinalGrades = VBA.Join(VecSemiFinalGrades, ";")

    ' Calculate GradesCx
    If BBooCx Then
      ReDim VecGradesCxSemi(m)
      For j = 0 To n
        For i = 0 To UBound(VecGradesCxSemi)
          Let DblAux1 = (VBA.CDbl(BVarGradCx(j)) * VBA.CDbl(BVarPerc(m))) - (VBA.CDbl(BVarHeadCx(j)) * VBA.CDbl(BVarPerc(m))) / k ' U1*
          If i = 0 Then ' Main Grade
            Let VecGradesCxSemi(i) = BVarGradCx(j)
          ElseIf i = m Then ' HeadGrade
            Let VecGradesCxSemi(i) = BVarHeadCx(j)
          Else
            Let VecGradesCxSemi(i) = DblAux1 / VBA.CDbl(BVarPerc(i))
          End If
          If VBA.CDbl(VecGradesCxSemi(i)) <= 0 Then GoTo EH ' #0 in gradescx!
        Next i
        Let StrFinalGradesCx = VBA.Join(VecGradesCxSemi, ";")
      Next j
    End If

  End If

  ' Return
  Let BResponse = Array(StrFinalGrades, StrFinalGradesCx)

  Let Balmet_core_main_app_eco_spreadbalmet_get_grades_hybrid_data = True

EH:
  Let VecUi = Empty
  Let VecHeadsUi = Empty
  Let VecUiZ = Empty
  Let VecSemiFinalGrades = Empty
  Let VecGradesCxSemi = Empty
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Balmet_core_main_app_eco_spreadbalmet_get_percents_hybrid_data( _
  ByVal BDblFeed As Double, _
  ByVal BDblFeedVol As Double, _
  ByVal BVarProd As Variant, _
  ByVal BVarUnit As Variant, _
  ByVal BEnuFeedType As EQLBAL_ENU_TYPECONTENT, _
  ByRef BVarResp As Variant)

  ' Obtains 1D Vectors with Feed + Products Weights or FeedVol + Products Vols, or both
  ' BVarUnit: Main balmet units vector

  Dim VecWei As Variant, VecVol As Variant
  Dim BooIsVol As Boolean
  Dim DblFeed As Double, DblFeedVol As Double, DblTail As Double, DblTailVol As Double
  Dim i As Long, j As Long, k As Long

  Const sFN As String = "BALMET::-core_main_spreadbalmet_get_weights_data_eco"

  On Error GoTo EH
  Let Balmet_core_main_app_eco_spreadbalmet_get_percents_hybrid_data = False

  If BEnuFeedType = eqlBalBoth Then

    ' Separate Weights and Vols
    ReDim VecWei(0)
    ReDim VecVol(0)
    Let j = 0
    Let k = 0
    For i = LBound(BVarUnit) To UBound(BVarUnit)
      Let BooIsVol = (VBA.LCase(VBA.CStr(BVarUnit(i))) = "g/l")
      ReDim Preserve VecVol(i)
      ReDim Preserve VecWei(i)
      If BooIsVol Then
        Let VecVol(i) = BVarProd(i)
        Let DblFeedVol = DblFeedVol + VBA.CDbl(BVarProd(i))
        Let j = j + 1
      Else
        Let VecWei(i) = BVarProd(i)
        Let DblFeed = DblFeed + VBA.CDbl(BVarProd(i))
        Let k = k + 1
      End If
    Next i

    ' Obtain values for empty spaces
    Let DblFeed = (BDblFeed - DblFeed) / k
    Let DblFeedVol = (BDblFeedVol - DblFeedVol) / j

    ' Refill empties
    For i = LBound(VecVol) To UBound(VecVol)
      If VecVol(i) = Empty Then Let VecVol(i) = DblFeedVol
      If VecWei(i) = Empty Then Let VecWei(i) = DblFeed
    Next i

    ' Add Tails and Feeds
    Let DblTail = BDblFeed - Application.WorksheetFunction.Sum(VecWei)
    Let DblTailVol = BDblFeedVol - Application.WorksheetFunction.Sum(VecVol)
    Let VecWei = RES.ArrayConcat(VecWei, Array(DblTail, BDblFeed))
    Let VecVol = RES.ArrayConcat(VecVol, Array(DblTailVol, BDblFeedVol))

    ' Calculate %
    For i = LBound(VecWei) To UBound(VecWei)
      Let VecWei(i) = (VBA.CDbl(VecWei(i)) / BDblFeed) * 100
      Let VecVol(i) = (VBA.CDbl(VecVol(i)) / BDblFeedVol) * 100
    Next i

  Else ' Solids or Volumes

    ' Add Tails and Feeds
    Let DblTail = BDblFeed - Application.WorksheetFunction.Sum(BVarProd)
    Let VecWei = RES.ArrayConcat(BVarProd, Array(DblTail, BDblFeed))

    ' Calculate %
    For i = LBound(VecWei) To UBound(VecWei)
      Let VecWei(i) = (VBA.CDbl(VecWei(i)) / BDblFeed) * 100
    Next i

  End If

  ' Return
  Let BVarResp = Array(VecWei, VecVol)

  Let Balmet_core_main_app_eco_spreadbalmet_get_percents_hybrid_data = True

EH:
  Let VecWei = Empty
  Let VecVol = Empty
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

' CORE METHODS
Private Function Balmet_core_main_app_spreadbalmet_get_products_names( _
  ByVal BVecNames As Variant, _
  ByVal BLngRows As Long, _
  ByRef BVecResponse As Variant, _
  Optional BBooHeads As Boolean = False, _
  Optional BFeedName As String = VBA.vbNullString) As Boolean

  Dim MDL As ClassEqlModel
  Dim VecAux0 As Variant, VecAux1 As Variant, VecAux2 As Variant
  Dim StrAux0 As String
  Dim i As Long, j As Long, xRow As Long, xCol As Long

  Const sFN As String = "BALMET::-core_spreadbalmet_get_products_names"

  On Error GoTo EH
  Let Balmet_core_main_app_spreadbalmet_get_products_names = False
  Set MDL = New ClassEqlModel
  Set MDL.ResourcesApp = RES

  ReDim VecAux1(0 To BLngRows)
  For i = 0 To UBound(VecAux1)
    Let StrAux0 = VBA.LCase(VBA.CStr(BVecNames(i)))
    Let VecAux0 = MDL.MGet(eqlMdlSheetSys, eqlMdlTblConfigchemicalelements_sys, eqlMdlArray, "Elemento", MStrWhere:="Símbolo=" & StrAux0)
    If VBA.IsArray(VecAux0) Then
      Let VecAux1(i) = VecAux0(0, 0)
    ElseIf (StrAux0 Like "b#" Or StrAux0 Like "bx#") Then
      Let VecAux1(i) = BVecNames(i)
    Else
      Call RES.ErrorHandler.ErrorHandlerSet(0.1, "¡Nombre de elemento: " & VBA.CStr(BVecNames(i)) & ", no existe!"): GoTo EH 'Error
    End If
  Next i
  Let VecAux0 = Empty

  ' Prepare main array
  Let xRow = BLngRows + 3 ' 3: Title, Feed, Tail
  Let xCol = 1
  ReDim VecAux0(0 To xRow, 0 To xCol)

  ' Set Balmet main products names
  Let j = 0
  Let VecAux0(0, 0) = "#"
  Let VecAux0(0, 1) = "Producto"
  For i = 1 To UBound(VecAux0)
    If i <= (UBound(VecAux1) + 1) Then Let VecAux0(i, 1) = VBA.CStr(VecAux1(j))
    If i = (UBound(VecAux1) + 2) Then Let VecAux0(i, 1) = "C"
    If i = (UBound(VecAux1) + 3) Then Let VecAux0(i, 1) = IIf(BFeedName = VBA.vbNullString, "A", BFeedName)
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
    If Not VBA.IsArray(VecAux0) Then Call RES.ErrorHandler.ErrorHandlerSet(0.2, "Fallo al generar encabezados de Cabezas ensayada y calculada."): GoTo EH
  End If
  Let VecAux1 = Empty

  ' Return
  Let BVecResponse = VecAux0
  Let VecAux0 = Empty
  Let Balmet_core_main_app_spreadbalmet_get_products_names = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Let VecAux2 = Empty
  Set MDL = Nothing
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Balmet_core_main_app_spreadbalmet_get_weights( _
  ByVal BDblFeed As Double, _
  ByRef BVarGrad As Variant, _
  ByVal BMethod As EQLBAL_ENU_METHOD, _
  ByRef BVarResponse As Variant, _
  Optional BBooFormulas As Boolean = False, _
  Optional BBooHeads As Boolean = False, _
  Optional BBooWithHeader As Boolean = False, _
  Optional BRngBox As Range, _
  Optional BEnuTypeFeedWei As EQLBAL_ENU_TYPECONTENT = eqlBalSolids, _
  Optional BFullData As Boolean = False) As Boolean

  ' Calculates the contents of main balmet (solids or volumes)
  Dim VecAux0 As Variant, VecAux1 As Variant
  Dim StrAux0 As String, StrAux1 As String
  Dim DblAux0 As Double, DblAux1 As Double
  Dim i As Long, j As Long, xRow As Long, xCol As Long, LngVolPos As Long

  Const sFN As String = "BALMET::-core_spreadbalmet_get_weights"

  On Error GoTo EH
  Let Balmet_core_main_app_spreadbalmet_get_weights = False

  ' Prepare main array
  Let xRow = (UBound(BVarGrad) - LBound(BVarGrad))
  Let xCol = 0
  ReDim VecAux0(0 To xRow, 0 To xCol)
  ReDim VecAux1(xRow)

  ' Calculate main results
  ' Calculate by selected Method
  If Not Balmet_core_main_app_spreadbalmet_get_weights_generic(BDblFeed, BVarGrad, BMethod, VecAux0) Then GoTo EH
  ' If results has negatives, change position of head grades form top to bottom
  If Application.WorksheetFunction.MIN(VecAux0) < 0 Then
    Let VecAux0 = RES.ArrayShift(BVarGrad)
    If Not VBA.IsArray(VecAux0) Then Call RES.ErrorHandler.ErrorHandlerSet(1.1, "Fallo calcular las cantidades de productos."): GoTo EH 'Error
    Let BVarGrad = RES.ArrayAddAtLast(BVarGrad, VecAux0)
    If Not VBA.IsArray(BVarGrad) Then Call RES.ErrorHandler.ErrorHandlerSet(1.1, "Fallo calcular las cantidades de productos."): GoTo EH 'Error
    If Not Balmet_core_main_app_spreadbalmet_get_weights_generic(BDblFeed, BVarGrad, BMethod, VecAux0) Then GoTo EH
    If Application.WorksheetFunction.MIN(VecAux0) < 0 Then 'There is not valid datas to be a balmet
      Call RES.ErrorHandler.ErrorHandlerSet(1.2, "Leyes incorrectas. Imposible calcular el Balance Metalúrgico."): GoTo EH 'Error
    End If
  End If
  ' If all is ok: get final data matrix
  If BFullData Then 'Include total
    Let VecAux0 = RES.ArrayAddAtLast(VecAux0, BDblFeed)
    If Not VBA.IsArray(VecAux0) Then Call RES.ErrorHandler.ErrorHandlerSet(1.1, "Fallo calcular las cantidades de productos."): GoTo EH 'Error
  End If

  ' Add Headers
  If BBooWithHeader Then
    Let VecAux0 = RES.ArrayAddAtFirst(VecAux0, IIf(BEnuTypeFeedWei = eqlBalVolumes, "Volumen[L]", "Peso[T]"))
    If Not VBA.IsArray(VecAux0) Then Call RES.ErrorHandler.ErrorHandlerSet(1.3, "Fallo al crear encabezado del Volumen."): GoTo EH 'Error
  End If

  ' Calculate Heads
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
    If Not VBA.IsArray(VecAux0) Then Call RES.ErrorHandler.ErrorHandlerSet(1.4, "Fallo al calcular el Volumen de cabezas calculada y ensayada."): GoTo EH 'Error
    Let VecAux1 = Empty
  End If

  ' Return
  Let BVarResponse = VecAux0
  Let VecAux0 = Empty
  Let Balmet_core_main_app_spreadbalmet_get_weights = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Balmet_core_main_app_spreadbalmet_get_weights_generic( _
  ByVal BDblFeed As Double, _
  ByVal BVecGrad As Variant, _
  ByVal BMethod As EQLBAL_ENU_METHOD, _
  ByRef BVecResponse As Variant) As Boolean

  ' Data required: (StrFeed, StrGrade)
  Const sFN As String = "BALMET::-core_main_spreadbalmet_get_weights_generic"

  On Error GoTo EH
  Let Balmet_core_main_app_spreadbalmet_get_weights_generic = False

  ' Calculate by Method selected
  Select Case BMethod
    Case eqlBalConventional ' Conventional
      If Not Balmet_core_main_app_spreadbalmet_get_weights_conventional(BDblFeed, BVecGrad, BVecResponse) Then GoTo EH
    Case eqlBalCramer ' CRAMER
      If Not Balmet_core_main_app_spreadbalmet_get_weights_cramer(BDblFeed, BVecGrad, BVecResponse) Then GoTo EH
    Case Else ' Inverse matrix
      If Not Balmet_core_main_app_spreadbalmet_get_weights_inverse(BDblFeed, BVecGrad, BVecResponse) Then GoTo EH
  End Select

  ' Return 2DVec
  Let Balmet_core_main_app_spreadbalmet_get_weights_generic = True

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Balmet_core_main_app_spreadbalmet_get_weights_conventional( _
  ByVal BDblFeed As Double, _
  ByVal BVecGrad As Variant, _
  ByRef BVecResponse As Variant) As Boolean

  ' Data required: (VecFeed, VecGrade)
  Dim VecAux As Variant
  Dim DblAux0 As Double, DblAux1 As Double, DblFeedG As Double, DblGradeG As Double, DblTailG As Double

  Const sFN As String = "BALMET::-core_main_spreadbalmet_get_weights_conventional"

  On Error GoTo EH
  Let Balmet_core_main_app_spreadbalmet_get_weights_conventional = False

  ' Conventional method, only for balances with 2 products
  If (UBound(BVecGrad) + 1) > 3 Then Call RES.ErrorHandler.ErrorHandlerSet(1.11, "Balances de dos productos sólo pueden tener tres elementos."): GoTo EH  'Not avaiable for more than two products

  Let DblGradeG = VBA.Abs(VBA.CDbl(BVecGrad(0, 0))) ' 0
  Let DblTailG = VBA.Abs(VBA.CDbl(BVecGrad(1, 0))) ' 1
  Let DblFeedG = VBA.Abs(VBA.CDbl(BVecGrad(2, 0))) ' 2

  Let DblAux0 = DblGradeG - DblTailG

  #If Debugging Then
    Call Debug_assert((DblAux0 = 0), sFN)
  #End If

  If DblAux0 = 0 Then Call RES.ErrorHandler.ErrorHandlerSet(1.112, "El determinante es cero, el Balance no tiene solución."): GoTo EH 'Error Div0!

  ' Resolve
  ReDim VecAux(0 To 1, 0 To 0)
  ' %Rec
  Let DblAux1 = (DblGradeG * (DblFeedG - DblTailG) * 100) / (DblFeedG * DblAux0)
  ' B
  Let VecAux(0, 0) = (DblAux1 * BDblFeed * DblFeedG) / (DblGradeG * 100)
  ' C
  Let VecAux(1, 0) = (BDblFeed - VecAux(0, 0))

  ' Forming the response: (Vec[mx1])
  Let BVecResponse = VecAux
  Let Balmet_core_main_app_spreadbalmet_get_weights_conventional = True

EH:
  Let VecAux = Empty
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Balmet_core_main_app_spreadbalmet_get_weights_cramer( _
  ByVal BDblFeed As Double, _
  ByVal BVecGrad As Variant, _
  ByRef BVecResponse As Variant) As Boolean

  ' Data required: (VecFeed, VecGrade)
  Dim VecAux As Variant
  Dim VecHeadGrades As Variant, VecProdGrades As Variant, VecAuxDeterm0 As Variant
  Dim DblDeterm As Double, DblDetermB As Double
  Dim i As Long, j As Long, k As Long, xRow As Long

  Const sFN As String = "BALMET::-core_main_spreadbalmet_get_weights_cramer"

  On Error GoTo EH
  Let Balmet_core_main_app_spreadbalmet_get_weights_cramer = False

  ' Resolve Balmet
  ' Get separated arrays
  Let VecProdGrades = BVecGrad
  Let VecHeadGrades = RES.ArrayPop(VecProdGrades)
  If Not VBA.IsArray(VecProdGrades) Then Call RES.ErrorHandler.ErrorHandlerSet(1.113, "Fallo al ordenar la matriz de leyes."): GoTo EH 'Error
  Let VecProdGrades = RES.ArrayAddAtFirst(VecProdGrades, 1, True)
  If Not VBA.IsArray(VecProdGrades) Then Call RES.ErrorHandler.ErrorHandlerSet(1.113, "Fallo al ordenar la matriz de leyes."): GoTo EH 'Error
  Let VecProdGrades = Application.WorksheetFunction.Transpose(VecProdGrades)
  Let VecHeadGrades = RES.ArrayAddAtFirst(VecHeadGrades, 1, True)
  If Not VBA.IsArray(VecHeadGrades) Then Call RES.ErrorHandler.ErrorHandlerSet(1.113, "Fallo al ordenar la matríz de leyes."): GoTo EH 'Error
  Let VecHeadGrades = Application.WorksheetFunction.Transpose(VecHeadGrades)
  Let xRow = (UBound(VecHeadGrades) - LBound(VecHeadGrades))

  ' Make a matrix AuxDeterm dynamic mxn with wich calculate each final weight value
  Let DblDeterm = Application.WorksheetFunction.MDeterm(VecProdGrades)
  If DblDeterm = 0 Then Call RES.ErrorHandler.ErrorHandlerSet(1.114, "El determinante es cero. Balance sin solución."): GoTo EH 'Error Div0!

  ReDim VecAux(0 To xRow, 0 To 0)
  Let k = 0
  For i = 0 To UBound(VecAux)
    Let VecAuxDeterm0 = VecProdGrades
    For j = LBound(VecHeadGrades) To UBound(VecHeadGrades)
      Let VecAuxDeterm0(k + LBound(VecAuxDeterm0), i + LBound(VecAuxDeterm0, 2)) = VecHeadGrades(j, LBound(VecHeadGrades))
      Let k = k + 1
    Next j
    Let DblDetermB = Application.WorksheetFunction.MDeterm(VecAuxDeterm0)
    ' Fill a matrix with Results values mx1
    Let VecAux(i, 0) = BDblFeed * (DblDetermB / DblDeterm)
    Let k = 0
  Next i

  ' Return: Forming the response: (Vec[mxn])
  Let BVecResponse = VecAux
  Let Balmet_core_main_app_spreadbalmet_get_weights_cramer = True

EH:
  Let BVecGrad = Empty
  Let VecAux = Empty
  Let VecHeadGrades = Empty
  Let VecProdGrades = Empty
  Let VecAuxDeterm0 = Empty
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Balmet_core_main_app_spreadbalmet_get_weights_inverse( _
  ByVal BDblFeed As Double, _
  ByVal BVecGrad As Variant, _
  ByRef BVecResponse As Variant) As Boolean

  ' Data required: (VecFeed, VecGrade)
  Dim VecAux As Variant
  Dim VecHeadGrades As Variant, VecProdGrades As Variant, VecAuxDeterm0 As Variant
  Dim DblDeterm As Double
  Dim i As Long, j As Long, k As Long, l As Long

  Const sFN As String = "BALMET::-core_spreadbalmet_get_weights_inverse"

  On Error GoTo EH
  Let Balmet_core_main_app_spreadbalmet_get_weights_inverse = False

  ' Resolve Balmet
  ' Get separated arrays
  Let VecProdGrades = BVecGrad
  Let VecHeadGrades = RES.ArrayPop(VecProdGrades)
  If Not VBA.IsArray(VecProdGrades) Then Call RES.ErrorHandler.ErrorHandlerSet(1.113, "Fallo al ordenar la matriz de leyes."): GoTo EH 'Error
  Let VecProdGrades = RES.ArrayAddAtFirst(VecProdGrades, 1, True)
  If Not VBA.IsArray(VecProdGrades) Then Call RES.ErrorHandler.ErrorHandlerSet(1.113, "Fallo al ordenar la matriz de leyes."): GoTo EH 'Error
  Let VecProdGrades = Application.WorksheetFunction.Transpose(VecProdGrades)
  Let VecHeadGrades = RES.ArrayAddAtFirst(VecHeadGrades, 1, True)
  If Not VBA.IsArray(VecHeadGrades) Then Call RES.ErrorHandler.ErrorHandlerSet(1.113, "Fallo al ordenar la matríz de leyes."): GoTo EH 'Error
  Let VecHeadGrades = Application.WorksheetFunction.Transpose(VecHeadGrades)

  For i = LBound(VecHeadGrades) To UBound(VecHeadGrades)
    Let VecHeadGrades(i, LBound(VecHeadGrades, 2)) = BDblFeed * VBA.CDbl(VecHeadGrades(i, LBound(VecHeadGrades, 2)))
  Next i

  ' Make a matrix AuxDeterm dynamic mxn with wich calculate each final weight value
  Let DblDeterm = Application.WorksheetFunction.MDeterm(VecProdGrades)
  If DblDeterm = 0 Then Call RES.ErrorHandler.ErrorHandlerSet(1.114, "El determinante es cero, el Balance no tiene solución."): GoTo EH  'Error

  ' Apply inverse matrix method and then fill a matrix with Results values mx1
  Let VecAuxDeterm0 = Application.WorksheetFunction.MInverse(VecProdGrades)
  Let VecAux = Application.WorksheetFunction.MMult(VecAuxDeterm0, VecHeadGrades)

  ' Return
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
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Balmet_core_main_app_spreadbalmet_get_percents( _
  ByVal BVarContents As Variant, _
  ByRef BVarResponse As Variant, _
  Optional BBooFormulas As Boolean = False, _
  Optional BBooHeads As Boolean = False, _
  Optional BBooWithHeader As Boolean = False, _
  Optional BEnuTypeFeedPer As EQLBAL_ENU_TYPECONTENT = eqlBalSolids, _
  Optional BRngBox As Range) As Boolean

  ' Calculates a column of percentages from values of contents (numeric results or by formulas)
  Dim VecAux0 As Variant, VecAux1 As Variant
  Dim DblFeed As Double
  Dim i As Long, j As Long, xRow As Long, xCol As Long

  Const sFN As String = "BALMET::-core_spreadbalmet_get_weights_percents"

  On Error GoTo EH
  Let Balmet_core_main_app_spreadbalmet_get_percents = False

  ' Prepare main array
  Let xRow = (UBound(BVarContents) - LBound(BVarContents))
  Let xCol = 0
  ReDim VecAux0(0 To xRow, 0 To xCol)

  ' Calculate main results
  If BBooFormulas Then
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

  ' Add Headers
  If BBooWithHeader Then
    Let VecAux0 = RES.ArrayAddAtFirst(VecAux0, IIf(BEnuTypeFeedPer = eqlBalVolumes, "Volumen[%]", "Peso[%]"))
    If Not VBA.IsArray(VecAux0) Then Call RES.ErrorHandler.ErrorHandlerSet(3.1, "Fallo al calcular el Porcentaje en " & IIf(BEnuTypeFeedPer = eqlBalSolids, "peso", "volumen") & "."): GoTo EH  'Error
  End If

  ' Calculate Heads
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
    If Not VBA.IsArray(VecAux0) Then Call RES.ErrorHandler.ErrorHandlerSet(3.2, "Fallo al calcular el Porcentaje en peso de cabezas calculada y ensayada."): GoTo EH 'Error
  End If

  ' Return
  Let BVarResponse = VecAux0
  Let VecAux0 = Empty
  Let Balmet_core_main_app_spreadbalmet_get_percents = True

EH:
  Let VecAux0 = Empty
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Balmet_core_main_app_spreadbalmet_get_volumes( _
  ByVal BDblFeedVol As Double, _
  ByVal BVarFeed As Variant, _
  ByVal BVarGrad As Variant, _
  ByVal BVarGradCx As Variant, _
  ByVal BVarUnit As Variant, _
  ByRef BVarResponse As Variant, _
  Optional BBooFormulas As Boolean = False, _
  Optional BBooWithHeader As Boolean = False, _
  Optional BBooHeads As Boolean = False, _
  Optional BRngBox As Range, _
  Optional BBooPercent As Boolean = False) As Boolean

  ' Calculates the volume contents of main balmet
  Dim VecAux0 As Variant, VecAux1 As Variant
  Dim StrAux0 As String, StrAux1 As String
  Dim DblAux0 As Double, DblAux1 As Double
  Dim i As Long, j As Long, xRow As Long, xCol As Long, LngVolPos As Long

  Const sFN As String = "BALMET::-core_spreadbalmet_get_volumes"

  On Error GoTo EH
  Let Balmet_core_main_app_spreadbalmet_get_volumes = False

  ' Prepare main array
  Let xRow = (UBound(BVarGrad) - LBound(BVarGrad))
  Let xCol = 0

  ' Calculate main results
  If Not Balmet_core_main_app_spreadbalmet_get_volumes_generic(BDblFeedVol, BVarFeed, BVarGrad, BVarGradCx, BVarUnit, VecAux0) Then GoTo EH

  ' Add Headers
  If BBooWithHeader Then
    Let VecAux0 = RES.ArrayAddAtFirst(VecAux0, "Volumen[L]")
    If Not VBA.IsArray(VecAux0) Then Call RES.ErrorHandler.ErrorHandlerSet(2.1, "Fallo al crear encabezado del Volumen."): GoTo EH 'Error
  End If

  ' Calculate Heads
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
    If Not VBA.IsArray(VecAux0) Then Call RES.ErrorHandler.ErrorHandlerSet(2.2, "Fallo al calcular el Volumen de cabezas calculada y ensayada."): GoTo EH 'Error
    Let VecAux1 = Empty
  End If

  ' Return
  Let BVarResponse = VecAux0
  Let VecAux0 = Empty
  Let Balmet_core_main_app_spreadbalmet_get_volumes = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Balmet_core_main_app_spreadbalmet_get_volumes_generic( _
  ByVal BDblFeedVol As Double, _
  ByVal BVarFeed As Variant, _
  ByVal BVarGrad As Variant, _
  ByVal BVarGradCx As Variant, _
  ByVal BVarUnit As Variant, _
  ByRef BVarResponse As Variant) As Boolean

  ' Data required: (BDblFeedVol, BVarFeed, BVarGrad, BVarGradCx, BVarUnit, BVarResponse)
  Dim VecAux0 As Variant, VecAux1 As Variant
  Dim DblAux0 As Double, DblAux1 As Double
  Dim i As Long, j As Long, xRow As Long, xCol As Long, LngVolPos As Long

  Const sFN As String = "BALMET::-core_spreadbalmet_get_volumes_generic"

  On Error GoTo EH
  Let Balmet_core_main_app_spreadbalmet_get_volumes_generic = False

  ' Prepare main array
  Let xRow = (UBound(BVarGrad) - LBound(BVarGrad))
  Let xCol = 0

  ' Prepare basic data (gettin' first occurence of g/l)
  Let VecAux1 = RES.ArrayConcat(BVarGrad, BVarGradCx, RByCols:=True)
  If Not VBA.IsArray(VecAux1) Then Call RES.ErrorHandler.ErrorHandlerSet(2.11, "Error en el cálculo genérico del volumen."): GoTo EH
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

  ' Calculate recoveries
  ' Gettin' percents
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
  ' Gettin' recoveries
  For i = LBound(VecAux1) To UBound(VecAux1)
    Let VecAux1(i, 0) = (VBA.CDbl(VecAux1(i, 0)) / VBA.CDbl(VecAux1(xRow, 0))) * 100
  Next i

  ' Calculate fines (feed vol [L] * grade vol [g/L])
  Let DblAux1 = BDblFeedVol * VBA.CDbl(VecAux0(xRow, 0))
  For i = LBound(VecAux1) To UBound(VecAux1)
    Let VecAux1(i, 0) = (VBA.CDbl(VecAux1(i, 0)) * DblAux1) / 100
  Next i

  ' Calculate Volumes
  Let DblAux0 = 0
  For i = LBound(VecAux1) To UBound(VecAux1)
    Let VecAux0(i, 0) = (VBA.CDbl(VecAux1(i, 0)) / VBA.CDbl(VecAux0(i, 0)))
    If i < (UBound(VecAux1) - 1) Then Let DblAux0 = DblAux0 + VBA.CDbl(VecAux0(i, 0))
    If i = (UBound(VecAux1) - 1) Then Let VecAux0(i, 0) = BDblFeedVol - DblAux0 ' Adjustment for Tails
    If i = (UBound(VecAux1)) Then Let VecAux0(i, 0) = BDblFeedVol
  Next i

  ' Return
  Let BVarResponse = VecAux0
  Let VecAux0 = Empty
  Let Balmet_core_main_app_spreadbalmet_get_volumes_generic = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Balmet_core_main_app_spreadbalmet_get_volumes_percents( _
  ByVal BVarFeedsVol As Variant, _
  ByRef BVarResponse As Variant, _
  Optional BBooFormulas As Boolean = False, _
  Optional BBooWithHeader As Boolean = False, _
  Optional BBooHeads As Boolean = False, _
  Optional BRngBox As Range) As Boolean

  ' Calculates the volume percents of main balmet
  Dim VecAux0 As Variant, VecAux1 As Variant
  Dim StrAux0 As String
  Dim DblFeedVol As Double
  Dim i As Long, j As Long, xRow As Long, xCol As Long, LngVolPos As Long

  Const sFN As String = "BALMET::-core_spreadbalmet_get_volumes_percents"

  On Error GoTo EH
  Let Balmet_core_main_app_spreadbalmet_get_volumes_percents = False

  ' Prepare main array
  Let xRow = (UBound(BVarFeedsVol) - LBound(BVarFeedsVol))
  Let xCol = 0
  ReDim VecAux0(0 To xRow, 0 To xCol)
  ReDim VecAux1(xRow)

  ' Calculate main results
  If BBooFormulas Then
    With BRngBox.Range("E3")
      For i = 0 To xRow
        If i < xRow Then Let VecAux0(i, 0) = "=(" & .Offset(i, 0).Address(False, False) & "/" & .Offset(xRow, 0).Address & ")*100"
        If i = xRow Then Let VecAux0(i, 0) = "=SUM(" & .Offset(0, 1).Resize(xRow, 1).Address(False, False) & ")"
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
    If Not VBA.IsArray(VecAux0) Then Call RES.ErrorHandler.ErrorHandlerSet(4.1, "Fallo al crear encabezado del Volumen."): GoTo EH 'Error
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
    If Not VBA.IsArray(VecAux0) Then Call RES.ErrorHandler.ErrorHandlerSet(4.2, "Fallo al calcular el Volumen de cabezas calculada y ensayada."): GoTo EH 'Error
  End If

  ' Return
  Let BVarResponse = VecAux0
  Let VecAux0 = Empty
  Let Balmet_core_main_app_spreadbalmet_get_volumes_percents = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

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

  ' Calculates the all grades of main balmet
  Dim VecAux0 As Variant, VecAux1 As Variant, VecAux2 As Variant, VecAux3 As Variant, VecAux4 As Variant
  Dim StrFeed As String, StrAux0 As String
  Dim BooVol As Boolean
  Dim DblAux0 As Double, DblFeed As Double
  Dim i As Long, j As Long, k As Long, xRow As Long, xCol As Long

  Const sFN As String = "BALMET::-core_spreadbalmet_get_weights_grades"

  On Error GoTo EH
  Let Balmet_core_main_app_spreadbalmet_get_grades = False

  ' Calculate main results
  If Not BBooJustHeads Then
    'Let VecAux0 = RES.ArrayAddAtLast(BVarGrades, BVarGradesCx, True)
    Let VecAux0 = RES.ArrayConcat(BVarGrades, BVarGradesCx, RByCols:=True)
    If Not VBA.IsArray(VecAux0) Then Call RES.ErrorHandler.ErrorHandlerSet(5.1, "Fallo al generar matriz de Leyes."): GoTo EH 'Error
  End If

  ' Add Headers
  If BBooWithHeader Then
    ReDim VecAux1(UBound(BVarNames))
    For i = 0 To UBound(VecAux1)
      Let VecAux1(i) = VBA.CStr(BVarNames(i)) & "[" & VBA.CStr(BVarUnits(i)) & "]"
    Next i
    Let VecAux0 = RES.ArrayAddAtFirst(VecAux0, VecAux1)
    If Not VBA.IsArray(VecAux0) Then Call RES.ErrorHandler.ErrorHandlerSet(5.2, "Fallo al crear encabezados de Leyes."): GoTo EH 'Error
  End If

  ' Calculate Heads
  If BBooHeads Then
    ' Method calls two other methods named calculated and tested
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
          Let VecAux2(1, i) = "=SUM(" & VBA.Join(VecAux3, ",") & ")/" & .Offset(xRow, 0).Address(False, False)
        End With
      Next i
      Let VecAux1 = Empty
      Let VecAux3 = Empty
      Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux2)
      If Not VBA.IsArray(VecAux0) Then Call RES.ErrorHandler.ErrorHandlerSet(5.4, "Fallo al crear encabezados de Leyes."): GoTo EH 'Error
      Let VecAux2 = Empty

    Else

      If BBooCx Then ' Just Cx
        Let VecAux1 = BVarGradesCx
        Let VecAux2 = RES.ArraySlice(BVarUnits, UBound(BVarGrades, 2))
      Else
        If BAllResult Then ' Concat both matrices
          Let VecAux1 = RES.ArrayConcat(BVarGrades, BVarGradesCx, RByCols:=True)
          Let VecAux2 = BVarUnits
        Else ' Just grades of normal balmet
          Let VecAux1 = BVarGrades
          Let VecAux2 = RES.ArraySlice(BVarUnits, UBound(BVarGrades, 2), True)
        End If
      End If
      If Not VBA.IsArray(VecAux1) Or Not VBA.IsArray(VecAux2) Then Call RES.ErrorHandler.ErrorHandlerSet(5.3, "Fallo al generar matriz de Leyes de cabeza ensayada y calculada."): GoTo EH 'Error
      ' Obtain %
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
      ' Fill final vector with tested grades
      Let VecAux3 = RES.ArrayPop(VecAux1, True)
      If Not VBA.IsArray(VecAux3) Then Call RES.ErrorHandler.ErrorHandlerSet(5.4, "Fallo al crear encabezados de Leyes."): GoTo EH 'Error
      ' Obtain U
      For j = LBound(VecAux2) To UBound(VecAux2) ' Units
        For i = LBound(VecAux1) To UBound(VecAux1) ' Grades
          If Not VBA.LCase(VBA.CStr(VecAux2(j))) = "g/l" Then
            Let VecAux1(i, j) = (VBA.CDbl(BVarContents(i, 0)) * VBA.CDbl(VecAux1(i, j)))
          Else
            Let VecAux1(i, j) = (VBA.CDbl(BVarVolumes(i, 0)) * VBA.CDbl(VecAux1(i, j)))
          End If
        Next i
      Next j
      For j = LBound(VecAux1, 2) To UBound(VecAux1, 2) ' Unities
        For i = LBound(VecAux1) To UBound(VecAux1)
          If i < UBound(VecAux1) Then
            Let DblAux0 = DblAux0 + (VBA.CDbl(VecAux1(i, j)))
          Else
            Let VecAux1(i, j) = DblAux0 / 100
          End If
        Next i
      Next j
      ' Fill final vector with calculated grades
      Let VecAux2 = Empty
      Let VecAux2 = RES.ArrayPop(VecAux1)
      If Not VBA.IsArray(VecAux2) Then Call RES.ErrorHandler.ErrorHandlerSet(5.4, "Fallo al crear encabezados de Leyes."): GoTo EH 'Error
      Let VecAux1 = Empty
      Let VecAux1 = RES.ArrayConcat(VecAux3, VecAux2)
      If Not VBA.IsArray(VecAux1) Then Call RES.ErrorHandler.ErrorHandlerSet(5.4, "Fallo al crear encabezados de Leyes."): GoTo EH 'Error
      Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1)
      If Not VBA.IsArray(VecAux0) Then Call RES.ErrorHandler.ErrorHandlerSet(5.4, "Fallo al crear encabezados de Leyes."): GoTo EH 'Error

    End If
  End If

  ' Return
  Let BVarResponse = VecAux0
  Let VecAux0 = Empty
  Let Balmet_core_main_app_spreadbalmet_get_grades = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Let VecAux2 = Empty
  Let VecAux3 = Empty
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

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

  ' Calculates the all unities of main balmet
  Dim VecAux0 As Variant, VecAux1 As Variant, VecAux2 As Variant
  Dim StrAux0 As String
  Dim BooVol As Boolean
  Dim DblAux0 As Double
  Dim i As Long, j As Long, k As Long, l As Long, xRow As Long, xCol As Long

  Const sFN As String = "BALMET::-core_spreadbalmet_get_unities"

  On Error GoTo EH
  Let Balmet_core_main_app_spreadbalmet_get_unities = False

  ' Calculate main results
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

    ' Prepare grades and units
    If BBooCx Then ' Just Cx
      Let VecAux1 = BVarGradesCx
      Let VecAux2 = RES.ArraySlice(BVarUnits, UBound(BVarGrades, 2))
    Else
      If BAllResult Then ' Concat both matrices
        Let VecAux1 = RES.ArrayConcat(BVarGrades, BVarGradesCx, RByCols:=True)
        Let VecAux2 = BVarUnits
      Else ' Just grades of normal balmet
        Let VecAux1 = BVarGrades
        Let VecAux2 = RES.ArraySlice(BVarUnits, UBound(BVarGrades, 2))
        Let VecAux2 = BVarUnits
      End If
    End If
    If Not VBA.IsArray(VecAux1) Or Not VBA.IsArray(VecAux2) Then Call RES.ErrorHandler.ErrorHandlerSet(6.1, "Fallo al generar matriz de Unidades."): GoTo EH 'Error
    ' Obtain Percents
    If Not Balmet_core_main_app_spreadbalmet_get_percents(BVarContents, BVarContents) Then GoTo EH
    If VBA.IsArray(BVarVolumes) Then
      If Not Balmet_core_main_app_spreadbalmet_get_volumes_percents(BVarVolumes, BVarVolumes) Then GoTo EH
    End If
    ' Obtain U
    Let xRow = (UBound(VecAux1) - LBound(VecAux1))
    Let xCol = (UBound(VecAux1, 2) - LBound(VecAux1, 2))
    ReDim VecAux0(0 To xRow, 0 To xCol)
    For j = LBound(VecAux2) To UBound(VecAux2) ' Units
      For i = LBound(VecAux1) To UBound(VecAux1) ' Grades
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
    If Not VBA.IsArray(VecAux2) Then Call RES.ErrorHandler.ErrorHandlerSet(6.1, "Fallo al generar matriz de Unidades."): GoTo EH 'Error
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

  ' Add Headers
  If BBooWithHeader Then
    ' Fill arr with names
    For i = 0 To UBound(BVarNames)
      Let BVarNames(i) = "Unidad " & VBA.CStr(BVarNames(i))
    Next i
    Let VecAux0 = RES.ArrayAddAtFirst(VecAux0, BVarNames)
    If Not VBA.IsArray(VecAux0) Then Call RES.ErrorHandler.ErrorHandlerSet(6.2, "Fallo al crear encabezados de Unidades."): GoTo EH 'Error
  End If

  ' Calculate Heads
  If BBooHeads Then
    ' Method calls two other methods named calculated and tested
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
      If Not VBA.IsArray(VecAux0) Then Call RES.ErrorHandler.ErrorHandlerSet(6.3, "Fallo al determinar Cabezas ensayada y calculada en Unidades."): GoTo EH 'Error
      Let VecAux1 = Empty

    Else

      ' Obtain U
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

      ' Fill final vector with calculated grades
      Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1)
      If Not VBA.IsArray(VecAux0) Then Call RES.ErrorHandler.ErrorHandlerSet(6.3, "Fallo al crear valores de cabeza ensayada y calculada en Unidades."): GoTo EH 'Error

    End If
  End If

  ' Return
  Let BVarResponse = VecAux0
  Let VecAux0 = Empty
  Let Balmet_core_main_app_spreadbalmet_get_unities = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Let VecAux2 = Empty
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

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

  ' Calculates the all Fines of main balmet
  Dim VecAux0 As Variant, VecAux1 As Variant, VecAux2 As Variant, VecAux3 As Variant
  Dim StrAux0 As String, StrAux1 As String
  Dim BooVol As Boolean
  Dim DblAux0 As Double
  Dim i As Long, j As Long, k As Long, l As Long, xRow As Long, xCol As Long

  Const sFN As String = "BALMET::-core_spreadbalmet_get_fines"

  On Error GoTo EH
  Let Balmet_core_main_app_spreadbalmet_get_fines = False
  Let BBooGr = Not BBooGr

  ' Calculate main results
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

    ' Prepare grades and units
    If BBooCx Then 'Just Cx
      Let VecAux1 = BVarGradesCx
      Let VecAux2 = RES.ArraySlice(BVarUnits, UBound(BVarGrades, 2))
    Else
      If BAllResult Then ' Concat both matrices
        Let VecAux1 = RES.ArrayConcat(BVarGrades, BVarGradesCx, RByCols:=True)
        Let VecAux2 = BVarUnits
      Else ' Just grades of normal balmet
        Let VecAux1 = BVarGrades
        Let VecAux2 = RES.ArraySlice(BVarUnits, UBound(BVarGrades, 2))
        Let VecAux2 = BVarUnits
      End If
    End If
    If Not VBA.IsArray(VecAux1) Or Not VBA.IsArray(VecAux2) Then Call RES.ErrorHandler.ErrorHandlerSet(7.1, "Fallo al generar matriz de Finos."): GoTo EH 'Error
    ' Grades conversion
    For j = LBound(VecAux2) To UBound(VecAux2) ' Units
      For i = LBound(VecAux1) To UBound(VecAux1) ' Grades
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
    ' Obtain F
    Let xRow = (UBound(VecAux1) - LBound(VecAux1))
    Let xCol = (UBound(VecAux1, 2) - LBound(VecAux1, 2))
    ReDim VecAux0(0 To xRow, 0 To xCol)
    For j = LBound(VecAux2) To UBound(VecAux2) ' Units
      For i = LBound(VecAux1) To UBound(VecAux1) ' Grades
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
    If Not VBA.IsArray(VecAux2) Then Call RES.ErrorHandler.ErrorHandlerSet(7.1, "Fallo al generar matriz de Finos."): GoTo EH 'Error
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

  ' Add Headers
  If BBooWithHeader Then
    ' Fill arr with names
    Let VecAux1 = BVarNames
    For i = 0 To UBound(BVarNames)
      Let VecAux1(i) = "Fino " & VBA.CStr(VecAux1(i)) & IIf(BBooGr, "[g]", "[T]")
    Next i
    Let VecAux0 = RES.ArrayAddAtFirst(VecAux0, VecAux1)
    If Not VBA.IsArray(VecAux0) Then Call RES.ErrorHandler.ErrorHandlerSet(7.2, "Fallo al crear encabezados de Finos."): GoTo EH 'Error
    Let VecAux1 = Empty
  End If

  ' Calculate Heads
  If BBooHeads Then
    ' Method calls two other methods named calculated and tested
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
      If Not VBA.IsArray(VecAux0) Then Call RES.ErrorHandler.ErrorHandlerSet(7.3, "Fallo al determinar Cabezas ensayada y calculada en Finos."): GoTo EH 'Error
      Let VecAux1 = Empty

    Else

      ' Obtain U
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

      ' Fill final vector with calculated grades
      Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1)
      If Not VBA.IsArray(VecAux0) Then Call RES.ErrorHandler.ErrorHandlerSet(7.3, "Fallo al crear valores de cabeza ensayada y calculada en Finos."): GoTo EH 'Error

    End If
  End If

  ' If user requires gr and OT in balmet
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
            If Not VBA.IsArray(VecAux0) Then Call RES.ErrorHandler.ErrorHandlerSet(7.4, "Fallo al crear valores de columna Finos [OT]."): GoTo EH 'Error
            Let k = k + 1
            Let VecAux1 = Empty
          End If
        End With
      Next j

    Else

      ' Iterate in all elements of Names
      ' If exists Ag or Au and if user requires grams or OT, calculate its values
      ' Once calculated, concat the resultant matrix into auxvec0 at the Ag or Au position
      ' Headers in this case, must be like: Fine Ag[OT]
      Let VecAux2 = VecAux0
      Let VecAux1 = Empty
      Let k = 0
      For j = LBound(BVarNames) To UBound(BVarNames)
        Let StrAux0 = VBA.LCase(VBA.CStr(BVarNames(j)))
        Let k = k + j
        If StrAux0 = "ag" Or StrAux0 = "au" Then
          ReDim VecAux1(0 To UBound(VecAux0), 0 To 0) ' 0,0: Header, pre-last: Head-tested, last: Head-calculated
          Let StrAux0 = "Fino " & VBA.CStr(BVarNames(j)) & "[OT]"
          Let VecAux1(0, 0) = "Fino " & VBA.CStr(BVarNames(j)) & "[OT]"
          For i = 1 To UBound(VecAux1)
            Let VecAux1(i, 0) = (VBA.CDbl(VecAux2(i, j)) * IIf(BBooGr, 1, (1000 ^ 2))) / 31.1035 ' Conversion to OT
          Next i
          Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1, k + 1, True)
          If Not VBA.IsArray(VecAux0) Then Call RES.ErrorHandler.ErrorHandlerSet(7.4, "Fallo al crear valores de columna Finos [OT]."): GoTo EH 'Error
          Let k = k + 1
          Let VecAux1 = Empty
        End If
      Next j

    End If
  End If

  ' Return
  Let BVarResponse = VecAux0
  Let VecAux0 = Empty
  Let Balmet_core_main_app_spreadbalmet_get_fines = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Let VecAux2 = Empty
  Let VecAux3 = Empty
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

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

  ' Data required: (BVarGrades, BVarNames)
  Dim VecAux0 As Variant, VecAux1 As Variant, VecAux2 As Variant, VecAux3 As Variant
  Dim StrAux0 As String, StrAux1 As String
  Dim BooVol As Boolean
  Dim DblAux0 As Double
  Dim i As Long, j As Long, k As Long, l As Long, xRow As Long, xCol As Long

  Const sFN As String = "BALMET::-core_spreadbalmet_get_recoveries"

  On Error GoTo EH
  Let Balmet_core_main_app_spreadbalmet_get_recoveries = False

  ' Calculate main results
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
          Let StrAux1 = "SUM(" & VBA.Join(VecAux1, ",") & ")"
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

    ' Obtain U
    If Not Balmet_core_main_app_spreadbalmet_get_unities(BVarContents, BVarVolumes, BVarGrades, BVarGradesCx, BVarUnits, BVarNames, VecAux0, BAllResult, BBooCx:=BBooCx) Then GoTo EH
    ' Obtain %R
    For j = LBound(VecAux0, 2) To UBound(VecAux0, 2)
      Let DblAux0 = VBA.CDbl(VecAux0(UBound(VecAux0), j))
      For i = LBound(VecAux0) To UBound(VecAux0)
        Let VecAux0(i, j) = (VBA.CDbl(VecAux0(i, j)) / DblAux0) * 100
      Next i
    Next j

  End If

  ' Add Headers
  If BBooWithHeader Then
    ' Fill arr with names
    Let VecAux1 = BVarNames
    For i = 0 To UBound(BVarNames)
      Let VecAux1(i) = "Rec. " & VBA.CStr(VecAux1(i)) & "[%]"
    Next i
    Let VecAux0 = RES.ArrayAddAtFirst(VecAux0, VecAux1)
    If Not VBA.IsArray(VecAux0) Then Call RES.ErrorHandler.ErrorHandlerSet(8.1, "Fallo al crear encabezados de Recuperaciones."): GoTo EH 'Error
    Let VecAux1 = Empty
  End If

  ' Calculate Heads
  If BBooHeads Then
    ' Method calls two other methods named calculated and tested
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
          Let StrAux1 = "SUM(" & VBA.Join(VecAux1, ",") & ")"
          Let VecAux1 = Empty
          Let VecAux2(0, j) = "=(" & StrAux1 & "/(" & .Offset(xRow, 0).Address(False, False) & "*" & .Offset(xRow, k + j).Address(False, False) & "))*100"
          Let VecAux2(1, j) = "=SUM(" & .Offset(0, l + j).Resize(xRow, 1).Address(False, False) & ")"
        End With
      Next j
      Let VecAux1 = Empty
      ' Fill final vector with calculated recoveries
      Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux2)
      If Not VBA.IsArray(VecAux0) Then Call RES.ErrorHandler.ErrorHandlerSet(8.2, "Fallo al crear valores de cabeza ensayada y calculada en Recuperaciones."): GoTo EH 'Error
      Let VecAux2 = Empty

    Else

      ' Obtain U
      Let DblAux0 = 0
      ReDim VecAux1(0 To 1, 0 To UBound(VecAux0, 2))
      For j = LBound(VecAux0, 2) To UBound(VecAux0, 2) ' Unities
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
      ' Fill final vector with calculated grades
      Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1)
      If Not VBA.IsArray(VecAux0) Then Call RES.ErrorHandler.ErrorHandlerSet(8.2, "Fallo al crear valores de cabeza ensayada y calculada en Recuperaciones."): GoTo EH 'Error
      Let VecAux1 = Empty

    End If
  End If

  ' Return
  Let BVarResponse = VecAux0
  Let VecAux0 = Empty
  Let Balmet_core_main_app_spreadbalmet_get_recoveries = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Let VecAux2 = Empty
  Let VecAux3 = Empty
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Balmet_core_main_app_spreadbalmet_get_ratio( _
  ByVal BVarContents As Variant, _
  ByVal BVarVolumes As Variant, _
  ByRef BVarResponse As Variant, _
  Optional BBooFormulas As Boolean = False, _
  Optional BBooWithHeader As Boolean = False, _
  Optional BBooHeads As Boolean = False, _
  Optional BRngBox As Range, _
  Optional BStrVol As String = VBA.vbNullString) As Boolean

  ' Data required: (BVarGrades, BVarNames)
  Dim VecAux0 As Variant, VecAux1 As Variant
  Dim StrAux0 As String
  Dim DblAux0 As Double
  Dim i As Long, j As Long, xRow As Long, xCol As Long

  Const sFN As String = "BALMET::-core_spreadbalmet_get_ratio"

  On Error GoTo EH
  Let Balmet_core_main_app_spreadbalmet_get_ratio = False

  ' Calculate main results
  Let VecAux1 = RES.ArrayConcat(BVarContents, BVarVolumes, RByCols:=True)
  If Not VBA.IsArray(VecAux1) Then Call RES.ErrorHandler.ErrorHandlerSet(9.1, "Fallo al crear resultados de Ratio."): GoTo EH 'Error
  Let xRow = (UBound(VecAux1) - LBound(VecAux1))
  Let xCol = (UBound(VecAux1, 2) - LBound(VecAux1, 2))
  ReDim VecAux0(0 To xRow, 0 To xCol)
  If BBooFormulas Then ' Ratio in formulas

      For j = 0 To xCol
        Let StrAux0 = IIf(j > 0, BStrVol, "C3")
        With BRngBox.Range(StrAux0)
          For i = 0 To xRow
            Let VecAux0(i, j) = "=" & .Offset(xRow, 0).Address(False, False) & "/" & .Offset(i, 0).Address(False, False)
          Next i
        End With
      Next j

  Else ' Ratio

    For j = LBound(VecAux0, 2) To UBound(VecAux0, 2)
      Let DblAux0 = VBA.CDbl(VecAux1(UBound(VecAux1), j))
      For i = LBound(VecAux0) To UBound(VecAux0)
        Let VecAux0(i, j) = DblAux0 / VBA.CDbl(VecAux1(i, j))
      Next i
    Next j

  End If
  Let VecAux1 = Empty

  ' Add Headers
  If BBooWithHeader Then
    ' Fill arr with names
    ReDim VecAux1(0 To 0, 0 To xCol)
    Let VecAux1(0, 0) = "Ratio" & IIf(xCol > 0, " Sólido", VBA.vbNullString)
    If xCol > 0 Then Let VecAux1(0, 1) = "Ratio Volumen"
    Let VecAux0 = RES.ArrayAddAtFirst(VecAux0, VecAux1)
    If Not VBA.IsArray(VecAux0) Then Call RES.ErrorHandler.ErrorHandlerSet(9.2, "Fallo al crear encabezados de Ratio."): GoTo EH 'Error
    Let VecAux1 = Empty
  End If

  ' Calculate Heads
  If BBooHeads Then
    ' Method calls two other methods named calculated and tested
    If BBooFormulas Then

      ReDim VecAux1(0 To 1, 0 To xCol)
      For j = 0 To xCol
        Let StrAux0 = IIf(j > 0, BStrVol, "C3")
        With BRngBox.Range(StrAux0)
          Let VecAux1(0, j) = "=" & .Offset(xRow + 1, 0).Address(False, False) & "/" & .Offset(xRow, 0).Address(False, False)
          Let VecAux1(1, j) = "=" & .Offset(xRow + 2, 0).Address(False, False) & "/" & .Offset(xRow, 0).Address(False, False)
        End With
      Next j
      ' Fill final vector with calculated grades
      Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1)
      If Not VBA.IsArray(VecAux0) Then Call RES.ErrorHandler.ErrorHandlerSet(9.3, "Fallo al crear valores de cabeza ensayada y calculada en Ratio."): GoTo EH 'Error
      Let VecAux1 = Empty

    Else

      ' Obtain U
      ReDim VecAux1(0 To 1, 0 To UBound(VecAux0, 2))
      Let VecAux1(0, 0) = VecAux0(UBound(VecAux0), 0)
      Let VecAux1(1, 0) = VecAux0(UBound(VecAux0), 0)
      If xCol > 0 Then
        Let VecAux1(0, 1) = VecAux0(UBound(VecAux0), 1)
        Let VecAux1(1, 1) = VecAux0(UBound(VecAux0), 1)
      End If
      ' Fill final vector with calculated grades
      Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1)
      If Not VBA.IsArray(VecAux0) Then Call RES.ErrorHandler.ErrorHandlerSet(9.3, "Fallo al crear valores de cabeza ensayada y calculada en Ratio."): GoTo EH 'Error
      Let VecAux1 = Empty

    End If
  End If

  ' Return
  Let BVarResponse = VecAux0
  Let VecAux0 = Empty
  Let Balmet_core_main_app_spreadbalmet_get_ratio = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function





' ASSETS --------------------------------------------------------------------------- '






' ERRORS MANNAGEMENT AND ASSERTSIONS ----------------------------------------------- '
Private Sub Debug_assert(ByVal Cond As Boolean, Optional Origin As String = VBA.vbNullString)

  #If Debbuging Then
    Debug.Assert Cond
    If Not Cond Then Debug.Print "Assert on: " & Origin
  #End If

End Sub





' METHODS LIST PUBLIC -------------------------------------------------------------- '
' ABSTRACT INTERFACE UI METHODS LIST AND UDF METHODS LIST (PUBLIC) ----------------- '
Public Function ClassEqlControllerInterface_GetSolution(ByVal CStrRequester As String, ByRef CVarRequest As Variant, ByRef CVarResponse As Variant) As Boolean

  On Error GoTo EH
  Let ClassEqlControllerInterface_GetSolution = Balmet_core_main(CStrRequester, CVarRequest, CVarResponse)

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise("BALMET::IGetSolution")

End Function

Public Function ClassEqlControllerInterface_GetSolutionUDF(ByVal CStrRequester As String, ByVal CVarRequest As Variant) As Variant

  Let ClassEqlControllerInterface_GetSolutionUDF = Balmet_core_udf(CStrRequester, CVarRequest)

End Function
' METHODS LIST PUBLIC -------------------------------------------------------------- '
' ================== CLASS EXCLIQ LITE CONTROLLER BALMET STRUCTURE ================= '

