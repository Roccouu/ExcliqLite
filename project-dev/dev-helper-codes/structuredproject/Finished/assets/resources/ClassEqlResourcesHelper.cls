VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassEqlResourcesHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ================= CLASS EXCLIQ LITE RESOURCES HELPER STRUCTURE =================== '
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



' ================= CLASS EXCLIQ LITE RESOURCES HELPER STRUCTURE =================== '
' OBJECT VARIABLES (GLOOBJ_) ------------------------------------------------------- '
' VARIABLES (GLOStr_) -------------------------------------------------------------- '
' PROPERTIES ----------------------------------------------------------------------- '
' CONSTRUCTOR/DESTRUCTOR ----------------------------------------------------------- '
' CORE METHODS LIST (PRIVATE) ------------------------------------------------------ '
' ASSETS --------------------------------------------------------------------------- '
' Resctr_assets_get_custom_names_and_units_for_balmet
' Resctr_assets_check_custom_names_does_not_has_wrong_units
' Resctr_assets_check_for_feeds_values
' Resctr_assets_check_for_volumes_and_solids_feeds
' Resctr_assets_check_for_maximous_values_in_grades
' Resctr_assets_bal_type_feed
' Resctr_assets_bal_generate_generic_names
' Resctr_assets_bal_generate_generic_units
' Resctr_assets_bal_generate_generic_recs
' Resctr_assets_refill_names
' Resctr_assets_refill_names_2
' Resctr_assets_check_names_on_db
' Resctr_assets_refill_units
' Resctr_assets_refill_units_2
' Resctr_assets_check_units_on_db
' Resctr_assets_refill_recs
' Resctr_assets_filter_names
' Resctr_assets_filter_units
' Resctr_assets_get_feed
' Resctr_assets_bal_get_products_bounds
' Resctr_assets_bal_get_products_bounds_fulldata
' Resctr_assets_bal_get_type
' Resctr_assets_bal_is_balmet
' Resctr_assets_bal_get_generic_names
' Resctr_assets_bal_get_grades
' Resctr_assets_bal_get_grades_cx
' ERROR ASSERTSIONS ---------------------------------------------------------------- '
' Debug_asset
' METHODS PUBLIC ------------------------------------------------------------------- '
' BalGetCustomNamesAndUnits
' BalCheckCustomNamesDoesNotHasWrongUnits
' BalCheckForFeedsValues
' BalCheckForVolumesAndSolidsFeeds
' BalCheckForMaximousValuesInGrades
' BalGetProductsBounds
' BalGetProductsBoundsFulldata
' BalGetType
' BalIsBalmet
' GetGenericNamesAndUnits
' BalTypeFeed
' BalGetGenericNames
' BalGetGenericUnits
' BalGetGenericRecoveries
' BalRefillNames
' BalRefillNames2
' BalRefillUnits
' BalRefillUnits2
' BalRefillRecs
' BalFilterNames
' BalFilterUnits
' BalGetFeed
' BalGetGrades
' ================= CLASS EXCLIQ LITE RESOURCES HELPER STRUCTURE =================== '




' ================= CLASS EXCLIQ LITE RESOURCES HELPER STRUCTURE =================== '
' OBJECT VARIABLES (GLOOBJ_) ------------------------------------------------------- '
Private EHGLOBAL As ClassEqlErrorHandler

' VARIABLES (GLOStr_) -------------------------------------------------------------- '
Private GLOBOO_FORMULAS As Boolean



' PROPERTIES ----------------------------------------------------------------------- '
Public Property Set ErrorHandler(ByRef RHS As ClassEqlErrorHandler)

  Set EHGLOBAL = RHS

End Property

Public Property Let Formulas(ByVal RBooFor As Boolean)

  Let GLOBOO_FORMULAS = RBooFor

End Property




' CONSTRUCTOR/DESTRUCTOR ----------------------------------------------------------- '
Private Sub Class_Terminate()

  Set EHGLOBAL = Nothing

End Sub





' CORE METHODS LIST PRIVATE -------------------------------------------------------- '
' ASSETS --------------------------------------------------------------------------- '
' Validations balmet
Private Function Resctr_assets_get_custom_names_and_units_for_balmet( _
  ByVal BRES As ClassEqlResources, _
  ByVal BVarSrcElems As Variant, _
  ByVal BVarSrcUnits As Variant, _
  ByVal BVarSrc As Variant, _
  ByRef BVarName As Variant, _
  ByRef BVarUnit As Variant, _
  ByRef BLngM As Long, _
  ByRef BLngN As Long, _
  ByRef BSTrMsg As String) As Boolean ' m, n: #names, #units

  Dim StrAux0 As String
  Dim i As Long, j As Long

  Const sFN As String = "RESCORE:-assets_get_custom_names_and_units_for_balmet"

  On Error GoTo EH
  Let Resctr_assets_get_custom_names_and_units_for_balmet = False

  ' Initialize variables
  Let BLngM = 0
  Let BLngN = 0
  ReDim BVarName(BLngM)
  ReDim BVarUnit(BLngN)

  ' Validate elements
  For i = LBound(BVarSrc) To UBound(BVarSrc)
    Let StrAux0 = VBA.LCase(VBA.CStr(BVarSrc(i)))
    If BRES.ArrayHasValue(BVarSrcElems, StrAux0, ALCase:=True) Then

      ReDim Preserve BVarName(BLngM)
      Let BVarName(BLngM) = BVarSrc(i)
      Let BLngM = BLngM + 1

    ElseIf BRES.ArrayHasValue(BVarSrcUnits, StrAux0, ALCase:=True) Then

      ReDim Preserve BVarUnit(BLngN)
      Let BVarUnit(BLngN) = BVarSrc(i)
      Let BLngN = BLngN + 1

    Else

      Let BSTrMsg = "El valor '" & VBA.CStr(BVarSrc(i)) & "' no existe en la base de datos de elementos ni unidades.": GoTo EH

    End If
  Next i

  Let Resctr_assets_get_custom_names_and_units_for_balmet = True

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Resctr_assets_check_custom_names_does_not_has_wrong_units( _
  ByVal BRES As ClassEqlResources, _
  ByVal BVarSrcElems As Variant, _
  ByRef BVarName As Variant, _
  ByRef BVarUnit As Variant, _
  ByVal BEnuCheckWhat As EQLBAL_ENU_WRONG_UNITS, _
  ByRef BSTrMsg As String) As Boolean ' m, n: #names, #units

  Dim StrAux0 As String
  Dim i As Long, j As Long

  Const sFN As String = "RESCORE:-assets_get_custom_names_and_units_for_balmet"

  On Error GoTo EH
  Let Resctr_assets_check_custom_names_does_not_has_wrong_units = False

  Select Case BEnuCheckWhat

    Case eqlBalAgUnit
      If BRES.ArrayHasValue(BVarName, "ag", ALCase:=True) Then ' Names
        Let i = BRES.ArrayGetValuePosition(BVarName, "ag")
        If VBA.CStr(BVarUnit(i)) = "%" Then Let BSTrMsg = "El mineral de plata no puede tener % como unidad.": GoTo EH
      End If

    Case eqlBalCxElemsUnitOnNormalBalmet
      For i = LBound(BVarName) To UBound(BVarName)
        If BRES.ArrayHasValue(BVarSrcElems, VBA.LCase(VBA.CStr(BVarName(i))), ALCase:=True) Then
          Let j = BRES.ArrayGetValuePosition(BVarName, VBA.LCase(VBA.CStr(BVarName(i))))
          If VBA.CStr(BVarUnit(j)) = "%" Then Let BSTrMsg = "El mineral '" & VBA.CStr(BVarName(i)) & "' no puede tener % como unidad.": GoTo EH
        End If
      Next i

    Case eqlBalDMOnOtherElements
      If BRES.ArrayHasValue(BVarUnit, "dm", ALCase:=True) Then ' Units
        Let i = BRES.ArrayGetValuePosition(BVarUnit, "dm")
        Let StrAux0 = VBA.LCase(VBA.CStr(BVarName(i)))
        If Not BRES.REGEX.isMineralGenericNameExcliqlite(StrAux0) Then
          If Not VBA.LCase(VBA.CStr(BVarName(i))) = "ag" Then Let BSTrMsg = "Sólo el mineral de plata puede tener DM como unidad.": GoTo EH
        End If
      End If

    Case eqlBalNonCx
      For i = LBound(BVarName) To UBound(BVarName)
        Let StrAux0 = VBA.LCase(VBA.CStr(BVarName(i)))
        If Not BRES.REGEX.isMineralGenericNameExcliqlite(StrAux0) Then
          If Not BRES.ArrayHasValue(BVarSrcElems, StrAux0, ALCase:=True) Then Let BSTrMsg = "El elemento/mineral '" & VBA.CStr(BVarName(i)) & "' no puede ser de tipo complejo.": GoTo EH
        End If
      Next i

    Case Else: GoTo EH
  End Select

  Let Resctr_assets_check_custom_names_does_not_has_wrong_units = True

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Resctr_assets_check_for_feeds_values( _
  ByVal BVarFeed As Variant, _
  ByVal BEnuTypeFeed As EQLBAL_ENU_TYPECONTENT, _
  ByVal BCDBL_MAX_TOP As Double, _
  ByRef BSTrMsg As String) As Boolean

  ' Checks the correct contents for feeds
  Dim DblAux0 As Double
  Dim i As Long

  Const sFN As String = "RESCORE:-assets_get_custom_names_and_units_for_balmet"

  On Error GoTo EH
  Let Resctr_assets_check_for_feeds_values = False

  Let i = UBound(BVarFeed) + 1
  If i = 0 Then Let BSTrMsg = "¡Alimentación debe contener valores numéricos mayor a cero!": GoTo EH
  If BEnuTypeFeed = eqlBalBoth Then

    If i > 1 Then ' Just one value to both feed types
      Let DblAux0 = VBA.CDbl(BVarFeed(0))
      If DblAux0 = 0 Then Let BSTrMsg = "¡Alimentación debe contener valores numéricos mayor a cero!": GoTo EH
      If DblAux0 >= BCDBL_MAX_TOP Then Let BSTrMsg = "¡La alimentación que introdujo es un número muy grande!": GoTo EH
    Else ' Each feed types has its own value
      Let DblAux0 = VBA.CDbl(BVarFeed(0))
      If DblAux0 = 0 Then Let BSTrMsg = "¡Alimentación debe contener valores numéricos mayor a cero!": GoTo EH
      If DblAux0 >= BCDBL_MAX_TOP Then Let BSTrMsg = "¡La alimentación que introdujo es un número muy grande!": GoTo EH
      Let DblAux0 = Application.WorksheetFunction.Sum(BVarFeed) - DblAux0
      If DblAux0 = 0 Then Let BSTrMsg = "¡Alimentación en volumen debe contener valores numéricos mayor a cero!": GoTo EH
      If DblAux0 >= BCDBL_MAX_TOP Then Let BSTrMsg = "¡La alimentación en volumen es un número muy grande!": GoTo EH
    End If

  Else

    Let DblAux0 = Application.WorksheetFunction.Sum(BVarFeed)
    If DblAux0 = 0 Then Let BSTrMsg = "¡Alimentación debe contener valores numéricos mayor a cero!": GoTo EH
    If DblAux0 >= BCDBL_MAX_TOP Then Let BSTrMsg = "¡La alimentación que introdujo es un número muy grande!": GoTo EH

  End If

  Let Resctr_assets_check_for_feeds_values = True

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Resctr_assets_check_for_volumes_and_solids_feeds( _
  ByVal BRES As ClassEqlResources, _
  ByVal BVarFeeds As Variant, _
  ByVal BVarProds As Variant, _
  ByVal BVarUnits As Variant, _
  ByVal BEnuTypeFeed As EQLBAL_ENU_TYPECONTENT, _
  ByVal BCDBL_MAX_TOP As Double, _
  ByRef BSTrMsg As String) As Boolean

  ' Checks the correct contents for volume feeds
  Dim DblAux0 As Double, DblAux1 As Double, DblAux2 As Double
  Dim i As Long

  Const sFN As String = "RESCORE:-assets_check_for_volumes_and_solids_feeds"

  On Error GoTo EH
  Let Resctr_assets_check_for_volumes_and_solids_feeds = False

  Let DblAux0 = Application.WorksheetFunction.Sum(BVarFeeds)

  If BEnuTypeFeed = eqlBalBoth Then

    Let DblAux0 = DblAux0 - IIf((UBound(BVarFeeds) - LBound(BVarFeeds)) > 0, VBA.CDbl(BVarFeeds(0)), 0)
    Let DblAux1 = 0
    Let DblAux2 = 0
    For i = LBound(BVarUnits) To UBound(BVarUnits)
      If VBA.LCase(VBA.CStr(BVarUnits(i))) = "g/l" Then
        Let DblAux1 = DblAux1 + VBA.CDbl(BVarProds(i))
      Else
        Let DblAux2 = DblAux2 + VBA.CDbl(BVarProds(i))
      End If
    Next i
    If VBA.CDbl(BVarFeeds(0)) >= BCDBL_MAX_TOP Then Let BSTrMsg = "La cantidad de sólidos de alimentación es un número muy grande.": GoTo EH
    If DblAux0 >= BCDBL_MAX_TOP Then Let BSTrMsg = "El volumen de alimentación es un número muy grande.": GoTo EH
    If DblAux2 >= VBA.CDbl(BVarFeeds(0)) Then Let BSTrMsg = "La suma de los pesos de productos debe ser menor que el peso de alimentación.": GoTo EH
    If DblAux1 >= DblAux0 Then Let BSTrMsg = "La suma de los volúmenes de productos debe ser menor que el volumen de alimentación.": GoTo EH

  ElseIf BEnuTypeFeed = eqlBalVolumes Then

    Let DblAux1 = Application.WorksheetFunction.Sum(BVarProds)
    If DblAux0 >= BCDBL_MAX_TOP Then Let BSTrMsg = "El volumen de alimentación es un número muy grande.": GoTo EH
    If DblAux1 >= DblAux0 Then Let BSTrMsg = "La suma de los volúmenes de productos debe ser menor que el volumen de alimentación.": GoTo EH

  Else

    Let DblAux1 = Application.WorksheetFunction.Sum(BVarProds)
    If DblAux0 >= BCDBL_MAX_TOP Then Let BSTrMsg = "La cantidad de sólidos de alimentación es un número muy grande.": GoTo EH
    If DblAux1 >= DblAux0 Then Let BSTrMsg = "La suma de los pesos de productos debe ser menor que el peso de alimentación.": GoTo EH

  End If

  Let Resctr_assets_check_for_volumes_and_solids_feeds = True

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Resctr_assets_check_for_maximous_values_in_grades( _
  ByVal BVarVilGrades As Variant, _
  ByVal BVarUnits As Variant, _
  ByVal BVarVilGradesCx As Variant, _
  ByVal BVarUnitsCx As Variant, _
  ByVal BBooCx As Boolean, _
  ByRef BSTrMsg As String, _
  ByRef BBooOneToOne As Boolean) As Boolean

  ' Check for maximous values in grades. Vilgrades as 2D
  Dim StrAux0 As String
  Dim i As Long, j As Long, k As Long

  Const CDBL_DM_TOP As Double = 10000
  Const CDBL_PC_TOP As Double = 100
  Const CDBL_GT_TOP As Double = 1000000
  Const sFN As String = "RESCORE:-assets_check_for_maximous_values_in_grades"

  On Error GoTo EH
  Let Resctr_assets_check_for_maximous_values_in_grades = False

  If BBooCx Then

    ' One to one ***
    If Not BBooOneToOne Then
      Let BVarVilGrades = Me.BalGetGrades(BVarVilGrades, k, BResponseDimension:=eqlBalAs2D)
      Let BVarVilGradesCx = Me.BalGetGradesCx(BVarVilGradesCx, k, k, BResponseDimension:=eqlBalAs2D)
    End If
    ' One to one ***
    For j = LBound(BVarUnitsCx) To UBound(BVarUnitsCx)
      Let StrAux0 = VBA.LCase(VBA.CStr(BVarUnitsCx(j)))
      For i = LBound(BVarVilGradesCx) To UBound(BVarVilGradesCx)
        Select Case StrAux0
          Case "dm"
            If VBA.CDbl(BVarVilGradesCx(i, j)) >= CDBL_DM_TOP Then Let BSTrMsg = "Ninguna ley de mineral expresada en decimarcos puede tener un valor mayor o igual a " & CDBL_DM_TOP & "DM.": GoTo EH
          Case Else
            If VBA.CDbl(BVarVilGradesCx(i, j)) >= CDBL_GT_TOP Then Let BSTrMsg = "Ninguna ley de mineral expresada en gramos por tonelada o gramos por litro puede tener un valor mayor o igual a " & CDBL_GT_TOP & "g/T o g/L.": GoTo EH
        End Select
      Next i
    Next j

  Else

    ' One to one ***
    If Not BBooOneToOne Then Let BVarVilGrades = Me.BalGetGrades(BVarVilGrades, k, BResponseDimension:=eqlBalAs2D)
    ' One to one ***
    For j = LBound(BVarUnits) To UBound(BVarUnits)
      Let StrAux0 = VBA.LCase(VBA.CStr(BVarUnits(j)))
      For i = LBound(BVarVilGrades) To UBound(BVarVilGrades)
        Select Case StrAux0
          Case "%"
            If VBA.CDbl(BVarVilGrades(i, j)) >= CDBL_PC_TOP Then Let BSTrMsg = "Ninguna ley de mineral expresada en porcentaje puede tener un valor mayor o igual a " & CDBL_PC_TOP & "%.": GoTo EH
          Case "dm"
            If VBA.CDbl(BVarVilGrades(i, j)) >= CDBL_DM_TOP Then Let BSTrMsg = "Ninguna ley de mineral expresada en decimarcos puede tener un valor mayor o igual a " & CDBL_DM_TOP & "DM.": GoTo EH
          Case Else
            If VBA.CDbl(BVarVilGrades(i, j)) >= CDBL_GT_TOP Then Let BSTrMsg = "Ninguna ley de mineral expresada en gramos por tonelada o gramos por litro puede tener un valor mayor o igual a " & CDBL_GT_TOP & "g/T o g/L.": GoTo EH
        End Select
      Next i
    Next j

  End If

  Let Resctr_assets_check_for_maximous_values_in_grades = True

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

' Balmet common assets
Private Function Resctr_assets_bal_type_feed(ByVal BRES As ClassEqlResources, ByVal BStrUnits As String) As EQLBAL_ENU_TYPECONTENT

  ' Return type feed of Balmet: solids, volumes or both
  Dim i As Long, j As Long
  Dim VecAux0 As Variant

  Const sFN As String = "RESCORE::-assets_bal_type_feed"

  On Error GoTo EH
  Let VecAux0 = VBA.Split(BStrUnits, ";")
  Let j = UBound(VecAux0) + 1
  Let i = BRES.ArrayHasValueN(VecAux0, "g/l", ALCase:=True)
  Let VecAux0 = Empty

  ' Return
  Let Resctr_assets_bal_type_feed = IIf(i = 0, eqlBalSolids, IIf(i = j, eqlBalVolumes, eqlBalBoth))

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Resctr_assets_bal_generate_generic_names( _
  ByVal BLngLength As Long, _
  ByVal BBooCx As Boolean, _
  ByVal BBooAsVector As Boolean) As Variant

  ' Return generic names for Balmet
  Dim i As Long
  Dim VecAux0 As Variant

  Const sFN As String = "RESCORE::-assets_bal_generate_generic_names"

  On Error GoTo EH
  If BLngLength < 0 Then GoTo EH
  ReDim VecAux0(BLngLength)
  For i = 0 To BLngLength
    Let VecAux0(i) = "B" & IIf(BBooCx, "X", VBA.vbNullString) & i + 1
  Next i

  ' Return
  If BBooAsVector Then
    Let Resctr_assets_bal_generate_generic_names = VecAux0
  Else
    Let Resctr_assets_bal_generate_generic_names = VBA.Join(VecAux0, ";")
  End If

EH:
  Let VecAux0 = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Resctr_assets_bal_generate_generic_units( _
  ByVal BLngLength As Long, _
  ByVal BBooCx As Boolean, _
  ByVal BBooAsVector As Boolean, _
  ByVal BBooTFs As Boolean) As Variant

  ' Return generic units for Balmet
  Dim i As Long
  Dim VecAux0 As Variant

  Const sFN As String = "RESCORE::-assets_bal_generate_generic_units"

  On Error GoTo EH
  If BLngLength < 0 Then GoTo EH
  ReDim VecAux0(BLngLength)
  For i = 0 To BLngLength
    Let VecAux0(i) = IIf(BBooTFs, "TF", IIf(BBooCx, "g/T", "%"))
  Next i

  ' Return
  If BBooAsVector Then
    Let Resctr_assets_bal_generate_generic_units = VecAux0
  Else
    Let Resctr_assets_bal_generate_generic_units = VBA.Join(VecAux0, ";")
  End If

EH:
  Let VecAux0 = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Resctr_assets_bal_generate_generic_recs( _
  ByVal BLngLength As Long, _
  ByVal BBooAsVector As Boolean) As Variant

  ' Return generic recoveries for Balmet
  Dim i As Long
  Dim VecAux0 As Variant

  Const sFN As String = "RESCORE::-assets_bal_generate_generic_recs"

  On Error GoTo EH
  If BLngLength < 0 Then GoTo EH
  ReDim VecAux0(BLngLength)
  For i = 0 To BLngLength
    Let VecAux0(i) = 80
  Next i

  ' Return
  If BBooAsVector Then
    Let Resctr_assets_bal_generate_generic_recs = VecAux0
  Else
    Let Resctr_assets_bal_generate_generic_recs = VBA.Join(VecAux0, ";")
  End If

EH:
  Let VecAux0 = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Resctr_assets_refill_names( _
  ByVal BRES As ClassEqlResources, _
  ByVal BStrNames As String, _
  ByVal BLngGradCol As Long, _
  ByVal BLngGradCxCol As Long, _
  ByRef BVarName As Variant, _
  ByRef BVarNameCx As Variant, _
  ByVal BBooCx As Boolean, _
  ByVal BBooAsString As Boolean) As Boolean

  ' Return Names and NamesCx. If there are incomplete vectors it is refilled with generic Names for Balmet
  Dim VecAux0 As Variant, VecAux1 As Variant, VecAux2 As Variant
  Dim BooExistsCx As Boolean
  Dim i As Long

  Const sFN As String = "RESCORE::-assets_refill_names"

  On Error GoTo EH
  Let Resctr_assets_refill_names = False

  Let BooExistsCx = False
  If BStrNames = VBA.vbNullString Then ' Return generics

    Let VecAux1 = Me.BalGetGenericNames(BLngGradCol, BBooAsVector:=True) ' Names
    If BBooCx Then Let VecAux2 = Me.BalGetGenericNames(BLngGradCxCol, BBooCx:=True, BBooAsVector:=True) ' Names Cx

  Else ' Return Names or Names refilled

    Let VecAux0 = VBA.Split(BStrNames, ";")

    If UBound(VecAux0) < BLngGradCol Then ' Refill Names
      Let VecAux1 = Me.BalGetGenericNames(BLngGradCol - (UBound(VecAux0) + 1), BBooAsVector:=True)
      Let VecAux1 = BRES.ArrayConcat(VecAux0, VecAux1)
      Let VecAux0 = Empty
      If Not VBA.IsArray(VecAux1) Then GoTo EH ' Error in slice method

    ElseIf UBound(VecAux0) = BLngGradCol Then ' Just Names exists
      Let VecAux1 = VecAux0
      Let VecAux0 = Empty

    Else ' Exists Unitscx
      Let BooExistsCx = True
      Let VecAux1 = BRES.ArraySlice(VecAux0, BLngGradCol, RBooGetLeftSide:=True)
      If Not VBA.IsArray(VecAux1) Then GoTo EH ' Error in slice method

    End If

    If BBooCx Then
      If BooExistsCx Then
        If UBound(VecAux0) < BLngGradCxCol Then ' Refill unitscx
          Let VecAux2 = Me.BalGetGenericNames(BLngGradCxCol - (UBound(VecAux0) + 1), BBooCx:=True, BBooAsVector:=True)
          Let VecAux2 = BRES.ArrayConcat(VecAux0, VecAux2)
          If Not VBA.IsArray(VecAux2) Then GoTo EH ' Error in slice method

        ElseIf UBound(VecAux0) = BLngGradCxCol Then  ' Unitscx is ok
          Let VecAux2 = VecAux0

        Else ' Get Unitscx
          Let VecAux2 = BRES.ArraySlice(VecAux0, BLngGradCxCol, RBooGetLeftSide:=True)
          If Not VBA.IsArray(VecAux2) Then GoTo EH ' Error in slice method

        End If
      Else
        Let VecAux2 = Me.BalGetGenericNames(BLngGradCxCol, BBooCx:=True, BBooAsVector:=True) ' Names Cx
      End If
    End If
  End If

  If BBooAsString Then
    Let BVarName = VBA.Join(VecAux1, ";")
    If BBooCx Then Let BVarNameCx = VBA.Join(VecAux2, ";")
  Else
    Let BVarName = VecAux1
    If BBooCx Then Let BVarNameCx = VecAux2
  End If
  Let VecAux0 = Empty

  Let Resctr_assets_refill_names = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Let VecAux2 = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Resctr_assets_refill_names_2( _
  ByVal BRES As ClassEqlResources, _
  ByVal BVarDBNames As Variant, _
  ByVal BStrNames As String, _
  ByVal BLngGradCol As Long, _
  ByVal BLngGradCxCol As Long, _
  ByRef BVarName As Variant, _
  ByRef BVarNameCx As Variant, _
  ByRef BStrErrMsg As String, _
  ByVal BBooCx As Boolean, _
  ByVal BVarDBNamesCx As Variant, _
  ByVal BBooAsString As Boolean) As Boolean

  ' Return Names and NamesCx. If there are incomplete vectors it is refilled with generic Names for Balmet
  Dim VecAux0 As Variant, VecAux1 As Variant, VecAux2 As Variant
  Dim StrAux0 As String
  Dim BooExistsCx As Boolean
  Dim i As Long

  Const sFN As String = "RESCORE::-assets_refill_names"

  On Error GoTo EH
  Let Resctr_assets_refill_names_2 = False

  Let BooExistsCx = False
  If BStrNames = VBA.vbNullString Then ' Return generics

    Let VecAux1 = Me.BalGetGenericNames(BLngGradCol, BBooAsVector:=True) ' Names
    If BBooCx Then Let VecAux2 = Me.BalGetGenericNames(BLngGradCxCol, BBooCx:=True, BBooAsVector:=True) ' Names Cx

  Else ' Return Names or Names refilled

    Let VecAux0 = VBA.Split(BStrNames, ";")

    If UBound(VecAux0) < BLngGradCol Then ' Refill Names
      If Not Resctr_assets_check_names_on_db(BRES, BVarDBNames, VecAux0, BStrErrMsg) Then GoTo EH ' Check if names exists in db
      Let VecAux1 = Me.BalGetGenericNames(BLngGradCol - (UBound(VecAux0) + 1), BBooAsVector:=True)
      Let VecAux1 = BRES.ArrayConcat(VecAux0, VecAux1)
      Let VecAux0 = Empty
      If Not VBA.IsArray(VecAux1) Then GoTo EH ' Error in slice method

    ElseIf UBound(VecAux0) = BLngGradCol Then ' Just Names exists
      If Not Resctr_assets_check_names_on_db(BRES, BVarDBNames, VecAux0, BStrErrMsg) Then GoTo EH ' Check if names exists in db
      Let VecAux1 = VecAux0
      Let VecAux0 = Empty

    Else ' Exists Unitscx
      Let BooExistsCx = True
      Let VecAux1 = BRES.ArraySlice(VecAux0, BLngGradCol, RBooGetLeftSide:=True)
      If Not VBA.IsArray(VecAux1) Then GoTo EH ' Error in slice method
      If Not Resctr_assets_check_names_on_db(BRES, BVarDBNames, VecAux1, BStrErrMsg) Then GoTo EH ' Check if names exists in db

    End If

    If BBooCx Then
      If BooExistsCx Then
        If UBound(VecAux0) < BLngGradCxCol Then ' Refill unitscx
          If Not Resctr_assets_check_names_on_db(BRES, BVarDBNames, VecAux0, BStrErrMsg) Then GoTo EH ' Check if names exists in db
          Let VecAux2 = Me.BalGetGenericNames(BLngGradCxCol - (UBound(VecAux0) + 1), BBooCx:=True, BBooAsVector:=True)
          Let VecAux2 = BRES.ArrayConcat(VecAux0, VecAux2)
          If Not VBA.IsArray(VecAux2) Then GoTo EH ' Error in slice method

        ElseIf UBound(VecAux0) = BLngGradCxCol Then  ' Unitscx is ok
          If Not Resctr_assets_check_names_on_db(BRES, BVarDBNames, VecAux0, BStrErrMsg) Then GoTo EH ' Check if names exists in db
          Let VecAux2 = VecAux0

        Else ' Get Unitscx
          Let VecAux2 = BRES.ArraySlice(VecAux0, BLngGradCxCol, RBooGetLeftSide:=True)
          If Not VBA.IsArray(VecAux2) Then GoTo EH ' Error in slice method
          If Not Resctr_assets_check_names_on_db(BRES, BVarDBNames, VecAux2, BStrErrMsg) Then GoTo EH ' Check if names exists in db

        End If
      Else
        Let VecAux2 = Me.BalGetGenericNames(BLngGradCxCol, BBooCx:=True, BBooAsVector:=True) ' Names Cx
      End If
    End If
  End If

  If BBooAsString Then
    Let BVarName = VBA.Join(VecAux1, ";")
    If BBooCx Then Let BVarNameCx = VBA.Join(VecAux2, ";")
  Else
    Let BVarName = VecAux1
    If BBooCx Then Let BVarNameCx = VecAux2
  End If
  Let VecAux0 = Empty

  Let Resctr_assets_refill_names_2 = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Let VecAux2 = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Resctr_assets_check_names_on_db( _
  ByVal BRES As ClassEqlResources, _
  ByVal BVarDBNames As Variant, _
  ByVal BVarNames As Variant, _
  ByRef BStrErrMsg As String) As Boolean

  ' Return Names and NamesCx based on db.
  Dim StrAux0 As String
  Dim i As Long

  Const sFN As String = "RESCORE::-assets_check_names_on_db"

  On Error GoTo EH
  Let Resctr_assets_check_names_on_db = False

  ' Check if names exists in db
  For i = LBound(BVarNames) To UBound(BVarNames)
    Let StrAux0 = VBA.LCase(VBA.CStr(BVarNames(i)))
    If Not BRES.REGEX.isMineralGenericNameExcliqlite(StrAux0) Then
      If Not BRES.ArrayHasValue(BVarDBNames, StrAux0, ALCase:=True) Then Let BStrErrMsg = "El mineral '" & VBA.CStr(BVarNames(i)) & "' no existe en la base de datos.": GoTo EH
    End If
  Next i

  Let Resctr_assets_check_names_on_db = True

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Resctr_assets_refill_units( _
  ByVal BRES As ClassEqlResources, _
  ByVal BStrUnits As String, _
  ByVal BLngGradCol As Long, _
  ByVal BLngGradCxCol As Long, _
  ByRef BVarUnit As Variant, _
  ByRef BVarUnitCx As Variant, _
  ByVal BBooCx As Boolean, _
  ByVal BBooAsString As Boolean) As Boolean

  ' Return Units and UnitsCx. If there are incomplete vectors it is refilled with generic units for Balmet
  Dim VecAux0 As Variant, VecAux1 As Variant, VecAux2 As Variant
  Dim BooExistsCx As Boolean
  Dim i As Long

  Const sFN As String = "RESCORE::-assets_refill_units"

  On Error GoTo EH
  Let Resctr_assets_refill_units = False

  Let BooExistsCx = False
  If BStrUnits = VBA.vbNullString Then ' Return generics

    Let VecAux1 = Me.BalGetGenericUnits(BLngGradCol, BBooAsVector:=True) ' Units
    If BBooCx Then Let VecAux2 = Me.BalGetGenericUnits(BLngGradCxCol, BBooCx:=True, BBooAsVector:=True) ' Units Cx

  Else ' Return Units or Units refilled

    Let VecAux0 = VBA.Split(BStrUnits, ";")

    If UBound(VecAux0) < BLngGradCol Then ' Refill units
      Let VecAux1 = Me.BalGetGenericUnits(BLngGradCol - (UBound(VecAux0) + 1), BBooAsVector:=True)
      Let VecAux1 = BRES.ArrayConcat(VecAux0, VecAux1)
      If Not VBA.IsArray(VecAux1) Then GoTo EH ' Error in slice method

    ElseIf UBound(VecAux0) = BLngGradCol Then ' Just Units exists
      Let VecAux1 = VecAux0

    Else ' Exists Unitscx
      Let BooExistsCx = True
      Let VecAux1 = BRES.ArraySlice(VecAux0, BLngGradCol, RBooGetLeftSide:=True)
      If Not VBA.IsArray(VecAux1) Then GoTo EH ' Error in slice method

    End If

    If BBooCx Then
      If BooExistsCx Then
        If UBound(VecAux0) < BLngGradCxCol Then ' Refill unitscx
          Let VecAux2 = Me.BalGetGenericUnits(BLngGradCxCol - (UBound(VecAux0) + 1), BBooCx:=True, BBooAsVector:=True)
          Let VecAux2 = BRES.ArrayConcat(VecAux0, VecAux2)
          If Not VBA.IsArray(VecAux2) Then GoTo EH ' Error in slice method

        ElseIf UBound(VecAux0) = BLngGradCxCol Then  ' Unitscx is ok
          Let VecAux2 = VecAux0

        Else ' Get Unitscx
          Let VecAux2 = BRES.ArraySlice(VecAux0, BLngGradCxCol, RBooGetLeftSide:=True)
          If Not VBA.IsArray(VecAux2) Then GoTo EH ' Error in slice method

        End If
      Else
        Let VecAux2 = Me.BalGetGenericUnits(BLngGradCxCol, BBooCx:=True, BBooAsVector:=True) ' Units Cx
      End If
    End If
  End If

  If BBooAsString Then
    Let BVarUnit = VBA.Join(VecAux1, ";")
    If BBooCx Then Let BVarUnitCx = VBA.Join(VecAux2, ";")
  Else
    Let BVarUnit = VecAux1
    If BBooCx Then Let BVarUnitCx = VecAux2
  End If
  Let VecAux0 = Empty

  Let Resctr_assets_refill_units = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Let VecAux2 = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Resctr_assets_refill_units_2( _
  ByVal BRES As ClassEqlResources, _
  ByVal BVarDBUnits As Variant, _
  ByVal BStrUnits As String, _
  ByVal BLngGradCol As Long, _
  ByVal BLngGradCxCol As Long, _
  ByRef BVarUnit As Variant, _
  ByRef BVarUnitCx As Variant, _
  ByRef BStrErrMsg As String, _
  ByVal BBooCx As Boolean, _
  ByVal BVarDBUnitsCx As Variant, _
  ByVal BBooAsString As Boolean) As Boolean

  ' Return Units and/or UnitsCx. If there are incomplete vectors it is refilled with generic units for Balmet
  Dim VecAux0 As Variant, VecAux1 As Variant, VecAux2 As Variant
  Dim BooExistsCx As Boolean
  Dim i As Long

  Const sFN As String = "RESCORE::-assets_refill_units"

  On Error GoTo EH
  Let Resctr_assets_refill_units_2 = False

  Let BooExistsCx = False
  If BStrUnits = VBA.vbNullString Then ' Return generics

    Let VecAux1 = Me.BalGetGenericUnits(BLngGradCol, BBooAsVector:=True) ' Units
    If BBooCx Then Let VecAux2 = Me.BalGetGenericUnits(BLngGradCxCol, BBooCx:=True, BBooAsVector:=True) ' Units Cx

  Else ' Return Units or Units refilled

    Let VecAux0 = VBA.Split(BStrUnits, ";")

    If UBound(VecAux0) < BLngGradCol Then ' Refill units
      If Not Resctr_assets_check_units_on_db(BRES, BVarDBUnits, VecAux0, BStrErrMsg) Then GoTo EH ' Check if there are correct units
      Let VecAux1 = Me.BalGetGenericUnits(BLngGradCol - (UBound(VecAux0) + 1), BBooAsVector:=True)
      Let VecAux1 = BRES.ArrayConcat(VecAux0, VecAux1)
      If Not VBA.IsArray(VecAux1) Then GoTo EH ' Error in slice method

    ElseIf UBound(VecAux0) = BLngGradCol Then ' Just Units exists
      If Not Resctr_assets_check_units_on_db(BRES, BVarDBUnits, VecAux0, BStrErrMsg) Then GoTo EH ' Check if there are correct units
      Let VecAux1 = VecAux0

    Else ' Exists Unitscx
      Let BooExistsCx = True
      Let VecAux1 = BRES.ArraySlice(VecAux0, BLngGradCol, RBooGetLeftSide:=True)
      If Not VBA.IsArray(VecAux1) Then GoTo EH ' Error in slice method
      If Not Resctr_assets_check_units_on_db(BRES, BVarDBUnits, VecAux1, BStrErrMsg) Then GoTo EH ' Check if there are correct units

    End If

    If BBooCx Then
      If BooExistsCx Then
        If UBound(VecAux0) < BLngGradCxCol Then ' Refill unitscx
          If Not Resctr_assets_check_units_on_db(BRES, BVarDBUnitsCx, VecAux0, BStrErrMsg) Then GoTo EH ' Check if there are correct units
          Let VecAux2 = Me.BalGetGenericUnits(BLngGradCxCol - (UBound(VecAux0) + 1), BBooCx:=True, BBooAsVector:=True)
          Let VecAux2 = BRES.ArrayConcat(VecAux0, VecAux2)
          If Not VBA.IsArray(VecAux2) Then GoTo EH ' Error in slice method

        ElseIf UBound(VecAux0) = BLngGradCxCol Then  ' Unitscx is ok
          If Not Resctr_assets_check_units_on_db(BRES, BVarDBUnitsCx, VecAux0, BStrErrMsg) Then GoTo EH ' Check if there are correct units
          Let VecAux2 = VecAux0

        Else ' Get Unitscx
          Let VecAux2 = BRES.ArraySlice(VecAux0, BLngGradCxCol, RBooGetLeftSide:=True)
          If Not VBA.IsArray(VecAux2) Then GoTo EH ' Error in slice method
          If Not Resctr_assets_check_units_on_db(BRES, BVarDBUnitsCx, VecAux2, BStrErrMsg) Then GoTo EH ' Check if there are correct units

        End If
      Else
        Let VecAux2 = Me.BalGetGenericUnits(BLngGradCxCol, BBooCx:=True, BBooAsVector:=True) ' Units Cx
      End If
    End If
  End If

  If BBooAsString Then
    Let BVarUnit = VBA.Join(VecAux1, ";")
    If BBooCx Then Let BVarUnitCx = VBA.Join(VecAux2, ";")
  Else
    Let BVarUnit = VecAux1
    If BBooCx Then Let BVarUnitCx = VecAux2
  End If
  Let VecAux0 = Empty

  Let Resctr_assets_refill_units_2 = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Let VecAux2 = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Resctr_assets_check_units_on_db( _
  ByVal BRES As ClassEqlResources, _
  ByVal BVarDBUnits As Variant, _
  ByVal BVarUnits As Variant, _
  ByRef BStrErrMsg As String) As Boolean

  ' Return Units and UnitsCx based on db.
  Dim StrAux0 As String
  Dim i As Long

  Const sFN As String = "RESCORE::-assets_check_units_on_db"

  On Error GoTo EH
  Let Resctr_assets_check_units_on_db = False

  ' Check if names exists in db
  For i = LBound(BVarUnits) To UBound(BVarUnits)
    Let StrAux0 = VBA.LCase(VBA.CStr(BVarUnits(i)))
    If Not BRES.ArrayHasValue(BVarDBUnits, StrAux0, ALCase:=True) Then Let BStrErrMsg = "La unidad '" & VBA.CStr(BVarUnits(i)) & "' no existe en la base de datos.": GoTo EH
  Next i

  Let Resctr_assets_check_units_on_db = True

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Resctr_assets_refill_recs( _
  ByVal BRES As ClassEqlResources, _
  ByVal BLngGradCol As Long, _
  ByVal BLngGradColcx As Long, _
  ByRef BVarRec As Variant, _
  ByRef BVarRecCx As Variant, _
  ByVal BBooCx As Boolean, _
  ByVal BBooAsString As Boolean) As Boolean

  ' Return recoveries. If there are incomplete vectors it are refilled with generics
  Dim VecAux0 As Variant, VecAux1 As Variant

  Const sFN As String = "RESCORE::-assets_refill_units"

  On Error GoTo EH
  Let Resctr_assets_refill_recs = False

  If BLngGradCol < 0 Then ' Return generics

    Let VecAux0 = Me.BalGetGenericRecoveries(BLngGradCol, BBooAsVector:=True) ' Recs

  Else ' Return refilled Recs

    If UBound(BVarRec) < BLngGradCol Then
      Let VecAux0 = Me.BalGetGenericRecoveries(BLngGradCol - (UBound(BVarRec) + 1), BBooAsVector:=True)
      Let VecAux0 = BRES.ArrayConcat(VecAux0, BVarRec)
      If Not VBA.IsArray(VecAux0) Then GoTo EH ' Error in slice method
    ElseIf UBound(BVarRec) = BLngGradCol Then  ' Just exists recs
      Let VecAux0 = BVarRec
    Else ' Recs are more than grades
      Let VecAux0 = BRES.ArraySlice(BVarRec, BLngGradCol, RBooGetLeftSide:=True)
      If Not VBA.IsArray(VecAux0) Then GoTo EH ' Error in slice method
    End If

  End If

  If BBooCx Then
    If BLngGradColcx < 0 Then ' Return generics

      Let VecAux1 = Me.BalGetGenericRecoveries(BLngGradColcx, BBooAsVector:=True)  ' Recscx

    Else ' Return refilled Recscx

      If UBound(BVarRecCx) < BLngGradColcx Then
        Let VecAux1 = Me.BalGetGenericRecoveries(BLngGradColcx - (UBound(BVarRecCx) + 1), BBooAsVector:=True)
        Let VecAux1 = BRES.ArrayConcat(VecAux1, BVarRecCx)
        If Not VBA.IsArray(VecAux1) Then GoTo EH ' Error in slice method
      ElseIf UBound(BVarRecCx) = BLngGradColcx Then ' Just exists recs
        Let VecAux1 = BVarRecCx
      Else ' Recscx are more than gradescx
        Let VecAux1 = BRES.ArraySlice(BVarRecCx, BLngGradColcx, RBooGetLeftSide:=True)
        If Not VBA.IsArray(VecAux1) Then GoTo EH ' Error in slice method
      End If

    End If
  End If


  If BBooAsString Then
    Let BVarRec = VBA.Join(VecAux0, ";")
    If BBooCx Then Let BVarRecCx = VBA.Join(VecAux1, ";")
  Else
    Let BVarRec = VecAux0
    If BBooCx Then Let BVarRecCx = VecAux1
  End If

  Let Resctr_assets_refill_recs = True

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function
Private Sub Resctr_assets_filter_names( _
  ByVal BRES As ClassEqlResources, _
  ByVal BVarNamesDB As Variant, _
  ByVal BVarNamesNeedles As Variant, _
  ByRef BVarNames As Variant, _
  ByVal BBooAsString As Boolean)

  ' Return a list of correct elements names
  Dim VecAux0 As Variant
  Dim StrAux0 As String
  Dim i As Long, j As Long

  Const sFN As String = "RESCORE::-assets_filter_names"

  On Error GoTo EH
  ReDim VecAux0(0)
  Let j = 0
  For i = LBound(BVarNamesNeedles) To UBound(BVarNamesNeedles)
    Let StrAux0 = VBA.LCase(VBA.CStr(BVarNamesNeedles(i)))
    If BRES.ArrayHasValue(BVarNamesDB, StrAux0, ALCase:=True) Then
      ReDim Preserve VecAux0(j)
      Let VecAux0(j) = BVarNamesNeedles(i)
      Let j = j + 1
    End If
  Next i

  If (j > 0) Then
    If BBooAsString Then
      Let BVarNames = VBA.Join(VecAux0, ";")
    Else
      Let BVarNames = VecAux0
    End If
  End If

EH:
  Let VecAux0 = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Sub

Private Sub Resctr_assets_filter_units( _
  ByVal BRES As ClassEqlResources, _
  ByVal BVarUnitsDB As Variant, _
  ByVal BVarUnitsNeedles As Variant, _
  ByRef BVarUnits As Variant, _
  ByVal BBooAsString As Boolean)

  ' Return a list of correct elements units
  Dim VecAux0 As Variant
  Dim StrAux0 As String
  Dim i As Long, j As Long

  Const sFN As String = "RESCORE::-assets_filter_units"

  On Error GoTo EH
  ReDim VecAux0(0)
  Let j = 0
  For i = LBound(BVarUnitsNeedles) To UBound(BVarUnitsNeedles)
    Let StrAux0 = VBA.LCase(VBA.CStr(BVarUnitsNeedles(i)))
    If BRES.ArrayHasValue(BVarUnitsDB, StrAux0, ALCase:=True) Then
      ReDim Preserve VecAux0(j)
      Let VecAux0(j) = BVarUnitsNeedles(i)
      Let j = j + 1
    End If
  Next i

  If (j > 0) Then
    If BBooAsString Then
      Let BVarUnits = VBA.Join(VecAux0, ";")
    Else
      Let BVarUnits = VecAux0
    End If
  End If

EH:
  Let VecAux0 = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Sub

Private Function Resctr_assets_get_feed( _
  ByVal BEnuTypeFeed As EQLBAL_ENU_TYPECONTENT, _
  ByVal BVarFeed As Variant, _
  ByRef BDblFeed As Double, _
  ByRef BDblFeedVol As Double) As Boolean

  ' Return Feeds
  Dim DblFeedSolid As Double, DblFeedVolume As Double, DblSum As Double

  Const sFN As String = "RESCORE::-assets_get_feed"

  On Error GoTo EH
  Let Resctr_assets_get_feed = False

  If BEnuTypeFeed = eqlBalBoth Then
    Let DblFeedSolid = VBA.CDbl(BVarFeed(0))
    Let DblSum = Application.WorksheetFunction.Sum(BVarFeed)
    Let DblFeedVolume = DblSum - DblFeedSolid
  Else
    Let DblFeedSolid = Application.WorksheetFunction.Sum(BVarFeed)
    Let DblFeedVolume = 0
  End If

  Let BDblFeed = DblFeedSolid
  Let BDblFeedVol = DblFeedVolume

  Let Resctr_assets_get_feed = True

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Resctr_assets_bal_get_products_bounds( _
  ByVal BLngSource As Long, _
  Optional BLngSourceCx As Long = 0, _
  Optional BWhat As EQLBAL_ENU_BOUNDS = eqlBalProducts) As Long

  ' Returns the number of Products, Cols, Rows, TotalProducts and TotalComplexProducts of any Balmet
  Dim LngProd As Long, LngCol As Long, LngRow As Long, LngTotProd As Long, LngTotCx As Long, LngAux0 As Long
  Dim BooFound As Boolean, BooTimeup As Boolean

  Const LngLASTBOUND As Long = 10 ' BAL_LIMIT:10, BAL_LIMITCX:3
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
  Else ' Determine Cx elements
    If BLngSourceCx > 0 Then
      Let LngTotCx = IIf(BLngSourceCx = LngRow Or BLngSourceCx = (LngRow * 2) Or BLngSourceCx = (LngRow * 3), BLngSourceCx + LngTotProd, 0)
      If LngTotCx = 0 Then GoTo EH
    End If
  End If

  ' Return
  Select Case BWhat
    Case eqlBalProducts: Let Resctr_assets_bal_get_products_bounds = LngProd - 1 ' Number of Balmet Products
    Case eqlBalColumns: Let Resctr_assets_bal_get_products_bounds = LngCol   ' Number of normal Balmet Cols
    Case eqlBalRows: Let Resctr_assets_bal_get_products_bounds = LngRow ' Number of Balmet Rows
    Case eqlBalFullProducts: Let Resctr_assets_bal_get_products_bounds = LngTotProd ' Number of Full normal Balmet grades. Ej. 8 grades to Balmet of 3 Prods
    Case eqlBalFullProductsCx: Let Resctr_assets_bal_get_products_bounds = LngTotCx ' Number of Full normal and complex Balmet grades. Ej. 12 grades to Balmet of 3 Prods
    Case eqlBalProductsCx: Let Resctr_assets_bal_get_products_bounds = LngTotCx - LngTotProd ' Number of Full complex Balmet grades. Ej. 12-8 grades = 4 to Balmet of 3 Prods
    Case eqlBalColumnsCx
      Let LngAux0 = LngTotCx - LngTotProd
      Let Resctr_assets_bal_get_products_bounds = LngAux0 / LngRow ' Number of complex Balmet columns.
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

  ' Returns the number of Products, Cols, Rows, TotalProducts and TotalComplexProducts of any Balmet
  Dim LngProd As Long, LngCol As Long, LngRow As Long, LngTotProd As Long, LngTotCx As Long
  Dim BooFound As Boolean, BooTimeup As Boolean

  Const LngLASTBOUND As Long = 10 ' BAL_LIMIT:10, BAL_LIMITCX:3
  Const sFN As String = "RES::-assets_bal_get_products_bounds_fulldata"

  On Error GoTo EH
  Let Resctr_assets_bal_get_products_bounds_fulldata = False

  ' Determine if the Source number is a valid value of elements to be a Balmet
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
  Else ' Determine Cx elements (just 3 columns of Cx allowed)
    If BLngSourceCx > 0 Then
      Let LngTotCx = IIf(BLngSourceCx = LngRow Or BLngSourceCx = (LngRow * 2) Or BLngSourceCx = (LngRow * 3), BLngSourceCx + LngTotProd, 0)
      If LngTotCx = 0 Then GoTo EH
    End If
  End If

  ' Return
  Let BLngProducts = LngProd - 1 ' Number of Balmet Products
  Let BLngColumns = LngCol ' Number of normal Balmet Cols
  Let BLngRows = LngRow ' Number of Balmet Rows
  Let BLngElements = LngTotProd ' Number of Full normal Balmet grades. Ej. 8 grades to Balmet of 3 Prods
  Let BLngElementsAndCx = IIf(LngTotCx = 0, LngTotProd, LngTotCx) ' Number of Full normal and complex Balmet grades. Ej. 12 grades to Balmet of 3 Prods
  Let BLngElementsCx = BLngElementsAndCx - LngTotProd ' Number of Full complex Balmet grades. Ej. 12-8 grades = 4 to Balmet of 3 Prods
  Let BLngColumnsCx = (BLngElementsAndCx - LngTotProd) / LngRow ' Number of complex Balmet columns.

  Let Resctr_assets_bal_get_products_bounds_fulldata = True

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Resctr_assets_bal_get_type(ByVal BLngSource As Long, Optional BLngSourceCx As Long = 0) As EQLBAL_ENU_TYPE

  ' Return "normal" or "complex" type of any Balmet
  Dim i As Long, j As Long, k As Long, l As Long, m As Long, n As Long, o As Long
  Dim StrType As EQLBAL_ENU_TYPE

  Const sFN = "RESCORE::-assets_bal_get_type"

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

  ' Return
  Let Resctr_assets_bal_get_type = StrType

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Resctr_assets_bal_is_balmet(ByVal BLngSource As Long, Optional BLngSourceCx As Long = 0) As Boolean

  ' Return if source is a correct number of elements for any Balmet
  Const sFN As String = "RESCORE::-assets_bal_is_balmet"

  On Error GoTo EH
  ' Return
  Let Resctr_assets_bal_is_balmet = Not (Resctr_assets_bal_get_type(BLngSource, BLngSourceCx) = eqlBalNothing)

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Resctr_assets_bal_get_generic_names( _
  ByVal BStrGrades As String, _
  ByRef BStrColumn As String, _
  ByRef BStrColumnCx As String, _
  ByVal BStrGradesCx As String, _
  ByVal BBooCx As Boolean, _
  ByVal BBooUnits As Boolean) As String

  Dim VecAux0 As Variant
  Dim StrAux0 As String
  Dim BooIsBalmet As Boolean
  Dim i As Long, j As Long, k As Long

  Const sFN As String = "RESCORE:-assets_bal_get_generic_names"

  On Error GoTo EH

  Let StrAux0 = VBA.vbNullString
  Let VecAux0 = VBA.Split(BStrGrades, ";")
  Let j = UBound(VecAux0) - LBound(VecAux0) + 1
  If BBooCx Then
    Let VecAux0 = Empty
    Let VecAux0 = VBA.Split(BStrGradesCx, ";")
    Let k = UBound(VecAux0) - LBound(VecAux0) + 1
  End If
  Let VecAux0 = Empty

  Let BooIsBalmet = Me.BalIsBalmet(j)
  If BooIsBalmet Then
    Let i = j
    Let j = Me.BalGetProductsBounds(j, BWhat:=eqlBalColumns) - 1
    Let k = Me.BalGetProductsBounds(i, k, eqlBalColumnsCx) - 1
    ReDim VecAux0(IIf(BBooCx, k, j))
    For i = 0 To UBound(VecAux0)
      If BBooUnits Then
        Let VecAux0(i) = IIf(BBooCx, "g/T", "%")
      Else
        Let VecAux0(i) = "B" & IIf(BBooCx, "X", VBA.vbNullString) & i + 1
      End If
    Next i
    Let StrAux0 = VBA.Join(VecAux0, ";")
    Let VecAux0 = Empty
  End If

  Let BStrColumn = IIf(BooIsBalmet, j, 0)
  Let BStrColumnCx = IIf(BooIsBalmet And BBooCx, k, 0)
  Let Resctr_assets_bal_get_generic_names = StrAux0

EH:
  Let VecAux0 = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function


Private Function Resctr_assets_bal_get_grades( _
  ByVal BVarVilSource As Variant, _
  ByRef BLngCounter As Long, _
  ByVal BxChar As String, _
  ByVal BResponseDimension As EQLBAL_ENU_GRADESVECTORTYPE) As Variant

  ' Return a string with grades of normal Balmet (vilsource: vector or string, counter is byref value)
  Dim VecSource As Variant, VecNew As Variant
  Dim i As Long, j As Long, k As Long, l As Long, xRow As Long, xCol As Long
  Dim StrFound As String, StrAux As String

  Const sFN As String = "APPBALMET::-assets_bal_get_grades"

  On Error GoTo EH
  Let BLngCounter = 0

  If VBA.TypeName(BVarVilSource) = "String" Then
    Let VecSource = VBA.Split(BVarVilSource, BxChar)
  Else
    Let VecSource = BVarVilSource
  End If

  Let k = UBound(VecSource) + 1
  Let xRow = Resctr_assets_bal_get_products_bounds(k, 0, eqlBalRows) - 1
  Let xCol = Resctr_assets_bal_get_products_bounds(k, 0, eqlBalColumns) - 1

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

  ' Return
  Let Resctr_assets_bal_get_grades = VecNew
  If BResponseDimension = eqlBalAs1D Then
    Let BLngCounter = ((UBound(VecNew) - LBound(VecNew)) + 1)
  Else
    Let BLngCounter = ((UBound(VecNew) - LBound(VecNew)) + 1) * ((UBound(VecNew, 2) - LBound(VecNew, 2)) + 1)
  End If

EH:
  Let VecSource = Empty
  Let VecNew = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Function Resctr_assets_bal_get_grades_cx( _
  ByVal BVarVilSourceCx As Variant, _
  ByVal BLngVilSource As Long, _
  ByRef BLngCounter As Long, _
  ByVal BxChar As String, _
  ByVal BResponseDimension As EQLBAL_ENU_GRADESVECTORTYPE) As Variant

  ' Return a string with grades of normal Balmet (vilsource: vector or string, counter is byref value)
  Dim VecSource As Variant, VecNew As Variant
  Dim i As Long, j As Long, k As Long, l As Long, xRow As Long, xCol As Long
  Dim StrFound As String, StrAux As String

  Const sFN As String = "APPBALMET::-assets_bal_get_grades_cx"

  On Error GoTo EH
  Let BLngCounter = 0

  If VBA.TypeName(BVarVilSourceCx) = "String" Then
    Let VecSource = VBA.Split(BVarVilSourceCx, BxChar)
  Else
    Let VecSource = BVarVilSourceCx
  End If
  Let l = UBound(VecSource) + 1
  Let k = BLngVilSource ' Normal Balmet number of grades elements

  Let xRow = Resctr_assets_bal_get_products_bounds(k, l, eqlBalRows) - 1
  Let xCol = Resctr_assets_bal_get_products_bounds(k, l, eqlBalColumnsCx) - 1

  If BResponseDimension = eqlBalAs1D Then '1D
    Let l = 0
    ReDim VecNew((UBound(VecSource) - LBound(VecSource)))
    For i = LBound(VecSource) To UBound(VecSource)
      Let VecNew(l) = VecSource(i)
      Let l = l + 1
    Next i
  Else ' 2D
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

  ' Return
  Let Resctr_assets_bal_get_grades_cx = VecNew
  If BResponseDimension = eqlBalAs1D Then
    Let BLngCounter = ((UBound(VecNew) - LBound(VecNew)) + 1)
  Else
    Let BLngCounter = ((UBound(VecNew) - LBound(VecNew)) + 1) * ((UBound(VecNew, 2) - LBound(VecNew, 2)) + 1)
  End If

EH:
  Let VecSource = Empty
  Let VecNew = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function
' ASSETS --------------------------------------------------------------------------- '




' ERRORS MANNAGEMENT AND ASSERTSIONS ----------------------------------------------- '
Private Sub Debug_assert(ByVal Cond As Boolean, Optional Origin As String = VBA.vbNullString)

  #If Debbuging Then
    Debug.Assert Cond
    If Not Cond Then Debug.Print "Assert on: " & Origin
  #End If

End Sub





' CORE METHODS LIST (PUBLIC) ------------------------------------------------------- '
' INTERFACE METHODS LIST FRM - SEL ------------------------------------------------- '
' Balmet validations
Public Function BalGetCustomNamesAndUnits( _
  ByVal BRES As ClassEqlResources, _
  ByVal BVarSrcElems As Variant, _
  ByVal BVarSrcUnits As Variant, _
  ByVal BVarSrc As Variant, _
  ByRef BVarName As Variant, _
  ByRef BVarUnit As Variant, _
  ByRef BLngM As Long, _
  ByRef BLngN As Long, _
  ByRef BSTrMsg As String) As Boolean ' m, n: #names, #units

  ' Builds two vectors with correct elements and units
  Const sFN As String = "RESCORE:-BalGetCustomNamesAndUnits"

  On Error GoTo EH
  Let BalGetCustomNamesAndUnits = Resctr_assets_get_custom_names_and_units_for_balmet(BRES, BVarSrcElems, BVarSrcUnits, BVarSrc, BVarName, BVarUnit, BLngM, BLngN, BSTrMsg)

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Public Function BalCheckCustomNamesDoesNotHasWrongUnits( _
  ByVal BRES As ClassEqlResources, _
  ByRef BSTrMsg As String, _
  Optional BVarSrcElems As Variant, _
  Optional BVarName As Variant, _
  Optional BVarUnit As Variant, _
  Optional BEnuCheckWhat As EQLBAL_ENU_WRONG_UNITS = eqlBalAgUnit) As Boolean

  ' Searchs for does not exists wrong units
  Const sFN As String = "RESCORE:-BalGetCustomNamesAndUnits"

  On Error GoTo EH
  Let BalCheckCustomNamesDoesNotHasWrongUnits = Resctr_assets_check_custom_names_does_not_has_wrong_units(BRES, BVarSrcElems, BVarName, BVarUnit, BEnuCheckWhat, BSTrMsg)

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Public Function BalCheckForFeedsValues( _
  ByVal BVarFeed As Variant, _
  ByVal BEnuTypeFeed As EQLBAL_ENU_TYPECONTENT, _
  ByVal BCDBL_MAX_TOP As Double, _
  ByRef BSTrMsg As String) As Boolean

  ' Searchs for correct feeds
  Const sFN As String = "RESCORE:BalCheckForFeedsValues"

  On Error GoTo EH
  Let BalCheckForFeedsValues = Resctr_assets_check_for_feeds_values(BVarFeed, BEnuTypeFeed, BCDBL_MAX_TOP, BSTrMsg)

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Public Function BalCheckForVolumesAndSolidsFeeds( _
  ByVal BRES As ClassEqlResources, _
  ByVal BVarFeeds As Variant, _
  ByVal BVarProds As Variant, _
  ByVal BVarUnits As Variant, _
  ByVal BEnuTypeFeed As EQLBAL_ENU_TYPECONTENT, _
  ByVal BCDBL_MAX_TOP As Double, _
  ByRef BSTrMsg As String) As Boolean

  ' Searchs for correct volume feeds
  Const sFN As String = "RESCORE:BalCheckForVolumesAndSolidsFeeds"

  On Error GoTo EH
  Let BalCheckForVolumesAndSolidsFeeds = Resctr_assets_check_for_volumes_and_solids_feeds(BRES, BVarFeeds, BVarProds, BVarUnits, BEnuTypeFeed, BCDBL_MAX_TOP, BSTrMsg)

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Public Function BalCheckForMaximousValuesInGrades( _
  ByRef BSTrMsg As String, _
  ByVal BVarVilGrades As Variant, _
  Optional BVarUnits As Variant, _
  Optional BVarVilGradesCx As Variant, _
  Optional BVarUnitsCx As Variant, _
  Optional BBooCx As Boolean, _
  Optional BBooOneToOne As Boolean) As Boolean

  ' Check for maximous values in grades
  Const sFN As String = "RESCORE:BalCheckForMaximousValuesInGrades"

  On Error GoTo EH
  Let BalCheckForMaximousValuesInGrades = Resctr_assets_check_for_maximous_values_in_grades(BVarVilGrades, BVarUnits, BVarVilGradesCx, BVarUnitsCx, BBooCx, BSTrMsg, BBooOneToOne)

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

' Balmet assets
Public Function BalGetProductsBounds( _
  ByVal BLngSource As Long, _
  Optional BLngSourceCx As Long = 0, _
  Optional BWhat As EQLBAL_ENU_BOUNDS = eqlBalProducts) As Long

  'Returns the number of Products, Cols, Rows, TotalProducts and TotalComplexProducts of any Balmet
  Const sFN As String = "RESCORE::BalGetProductsBounds"

  On Error GoTo EH
  Let BalGetProductsBounds = Resctr_assets_bal_get_products_bounds(BLngSource, BLngSourceCx, BWhat)

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Public Function BalGetProductsBoundsFulldata( _
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
  Const sFN As String = "RESCORE::BalGetProductsBoundsFulldata"

  On Error GoTo EH
  Let BalGetProductsBoundsFulldata = Resctr_assets_bal_get_products_bounds_fulldata(BLngSource, BLngProducts, BLngColumns, BLngRows, BLngElements, BLngElementsAndCx, BLngElementsCx, BLngColumnsCx, BLngSourceCx)

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Public Function BalGetType(ByVal BLngSource As Long, Optional BLngSourceCx As Long = 0) As EQLBAL_ENU_TYPE

  ' Return "normal" or "complex" type of any Balmet
  Const sFN As String = "RESCORE::BalGetType"

  On Error GoTo EH
  Let BalGetType = Resctr_assets_bal_get_type(BLngSource, BLngSourceCx)

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Public Function BalIsBalmet(ByVal BLngSource As Long, Optional BLngSourceCx As Long = 0) As Boolean

  ' Return if source is a correct number of elements of any Balmet
  Const sFN As String = "RESCORE::IsBalmet"

  On Error GoTo EH
  Let BalIsBalmet = Resctr_assets_bal_is_balmet(BLngSource, BLngSourceCx)

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Public Function GetGenericNamesAndUnits( _
  ByVal BStrSource As String, _
  ByRef BStrColumn As String, _
  ByRef BStrColumnCx As String, _
  Optional BStrSourceCx As String = VBA.vbNullString, _
  Optional BBooCx As Boolean = False, _
  Optional BBooUnits As Boolean = False) As String

  ' Return a string of generic names or units
  Const sFN As String = "RESCORE::GetGenericNames"

  On Error GoTo EH
  Let GetGenericNamesAndUnits = Resctr_assets_bal_get_generic_names(BStrSource, BStrColumn, BStrColumnCx, BStrSourceCx, BBooCx, BBooUnits)

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Public Function BalTypeFeed(ByVal BRES As ClassEqlResources, ByVal BStrUnits As String) As EQLBAL_ENU_TYPECONTENT

  ' Return type feed of Balmet: solids, volumes or both
  Const sFN As String = "RESCORE::BalTypeFeed"

  On Error GoTo EH
  Let BalTypeFeed = Resctr_assets_bal_type_feed(BRES, BStrUnits)

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Public Function BalGetGenericNames( _
  ByVal BLngLength As Long, _
  Optional BBooCx As Boolean = False, _
  Optional BBooAsVector As Boolean = False) As Variant

  ' Return Balmet generics names
  Const sFN As String = "RESCORE::BalGetGenericNames"

  On Error GoTo EH
  Let BalGetGenericNames = Resctr_assets_bal_generate_generic_names(BLngLength, BBooCx, BBooAsVector)

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Public Function BalGetGenericUnits( _
  ByVal BLngLength As Long, _
  Optional BBooCx As Boolean = False, _
  Optional BBooAsVector As Boolean = False, _
  Optional BBooTFs As Boolean = False) As Variant

  ' Return Balmet generics names
  Const sFN As String = "RESCORE::BalGetGenericUnits"

  On Error GoTo EH
  Let BalGetGenericUnits = Resctr_assets_bal_generate_generic_units(BLngLength, BBooCx, BBooAsVector, BBooTFs)

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Public Function BalGetGenericRecoveries( _
  ByVal BLngLength As Long, _
  Optional BBooAsVector As Boolean = False) As Variant

  ' Return Balmet generics names
  Const sFN As String = "RESCORE::BalGetGenericRecoveries"

  On Error GoTo EH
  Let BalGetGenericRecoveries = Resctr_assets_bal_generate_generic_recs(BLngLength, BBooAsVector)

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Public Function BalRefillNames( _
  ByVal BRES As ClassEqlResources, _
  ByVal BStrNames As String, _
  ByVal BLngGradCol As Long, _
  ByVal BLngGradCxCol As Long, _
  ByRef BVarName As Variant, _
  ByRef BVarNameCx As Variant, _
  Optional BBooCx As Boolean = False, _
  Optional BBooAsString As Boolean = False) As Boolean

  ' Return Balmet correct names including generics
  Const sFN As String = "RESCORE::BalRefillNames"

  On Error GoTo EH
  Let BalRefillNames = Resctr_assets_refill_names(BRES, BStrNames, BLngGradCol, BLngGradCxCol, BVarName, BVarNameCx, BBooCx, BBooAsString)

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Public Function BalRefillNames2( _
  ByVal BRES As ClassEqlResources, _
  ByVal BVarDBNames As Variant, _
  ByVal BStrNames As String, _
  ByVal BLngGradCol As Long, _
  ByVal BLngGradCxCol As Long, _
  ByRef BVarName As Variant, _
  ByRef BVarNameCx As Variant, _
  ByRef BStrErrMsg As String, _
  Optional BBooCx As Boolean = False, _
  Optional BVarDBNamesCx As Variant = Empty, _
  Optional BBooAsString As Boolean = False) As Boolean

  ' Return Balmet correct names including generics
  Const sFN As String = "RESCORE::BalRefillNames2"

  On Error GoTo EH
  Let BalRefillNames2 = Resctr_assets_refill_names_2(BRES, BVarDBNames, BStrNames, BLngGradCol, BLngGradCxCol, BVarName, BVarNameCx, BStrErrMsg, BBooCx, BVarDBNamesCx, BBooAsString)

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Public Function BalRefillUnits( _
  ByVal BRES As ClassEqlResources, _
  ByVal BStrUnits As String, _
  ByVal BLngGradCol As Long, _
  ByVal BLngGradCxCol As Long, _
  ByRef BVarUnit As Variant, _
  ByRef BVarUnitCx As Variant, _
  Optional BBooCx As Boolean = False, _
  Optional BBooAsString As Boolean = False) As Boolean

  ' Return Balmet correct units including generics units
  Const sFN As String = "RESCORE::BalRefillUnits"

  On Error GoTo EH
  Let BalRefillUnits = Resctr_assets_refill_units(BRES, BStrUnits, BLngGradCol, BLngGradCxCol, BVarUnit, BVarUnitCx, BBooCx, BBooAsString)

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Public Function BalRefillUnits2( _
  ByVal BRES As ClassEqlResources, _
  ByVal BVarDBUnits As Variant, _
  ByVal BStrUnits As String, _
  ByVal BLngGradCol As Long, _
  ByVal BLngGradCxCol As Long, _
  ByRef BVarUnit As Variant, _
  ByRef BVarUnitCx As Variant, _
  ByRef BStrErrMsg As String, _
  Optional BBooCx As Boolean = False, _
  Optional BVarDBUnitsCx As Variant = Empty, _
  Optional BBooAsString As Boolean = False) As Boolean

  ' Return Balmet correct units including generics units
  Const sFN As String = "RESCORE::BalRefillUnits2"

  On Error GoTo EH
  Let BalRefillUnits2 = Resctr_assets_refill_units_2(BRES, BVarDBUnits, BStrUnits, BLngGradCol, BLngGradCxCol, BVarUnit, BVarUnitCx, BStrErrMsg, BBooCx, BVarDBUnitsCx, BBooAsString)

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Public Function BalRefillRecs( _
  ByVal BRES As ClassEqlResources, _
  ByVal BLngGradCol As Long, _
  ByVal BLngGradColcx As Long, _
  ByRef BVarRec As Variant, _
  ByRef BVarRecCx As Variant, _
  Optional BBooCx As Boolean = False, _
  Optional BBooAsString As Boolean) As Boolean

  ' Return Balmet correct and refilled recoveries
  Const sFN As String = "RESCORE::BalRefillRecs"

  On Error GoTo EH
  Let BalRefillRecs = Resctr_assets_refill_recs(BRES, BLngGradCol, BLngGradColcx, BVarRec, BVarRecCx, BBooCx, BBooAsString)

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Public Sub BalFilterNames( _
  ByVal BRES As ClassEqlResources, _
  ByVal BVarNamesDB As Variant, _
  ByVal BVarNamesNeedles As Variant, _
  ByRef BVarNames As Variant, _
  Optional BBooAsString As Boolean = False)

  ' Return Balmet correct and Elements names
  Const sFN As String = "RESCORE::BalFilterNames"

  On Error GoTo EH
  Call Resctr_assets_filter_names(BRES, BVarNamesDB, BVarNamesNeedles, BVarNames, BBooAsString)

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Sub

Public Sub BalFilterUnits( _
  ByVal BRES As ClassEqlResources, _
  ByVal BVarUnitsDB As Variant, _
  ByVal BVarUnitsNeedles As Variant, _
  ByRef BVarUnits As Variant, _
  Optional BBooAsString As Boolean = False)

  ' Return Balmet correct and Elements names
  Const sFN As String = "RESCORE::BalFilterUnits"

  On Error GoTo EH
  Call Resctr_assets_filter_units(BRES, BVarUnitsDB, BVarUnitsNeedles, BVarUnits, BBooAsString)

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Sub

Public Function BalGetFeed( _
  ByVal BEnuTypeFeed As EQLBAL_ENU_TYPECONTENT, _
  ByVal BVarFeed As Variant, _
  ByRef BDblFeed As Double, _
  ByRef BDblFeedVol As Double) As Boolean

  ' Return Balmet correct Feeds
  Const sFN As String = "RESCORE::BalGetFeed"

  On Error GoTo EH
  Let BalGetFeed = Resctr_assets_get_feed(BEnuTypeFeed, BVarFeed, BDblFeed, BDblFeedVol)

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Public Function BalGetGrades( _
  ByVal BVarVilSource As Variant, _
  ByRef BLngCounter As Long, _
  Optional BxChar As String = ";", _
  Optional BResponseDimension As EQLBAL_ENU_GRADESVECTORTYPE = eqlBalAs1D) As Variant

  ' Return a string with grades of normal Balmet (vilsource: vector or string, counter is byref value)
  Const sFN As String = "RESCORE::BalGetGrades"

  On Error GoTo EH
  Let BalGetGrades = Resctr_assets_bal_get_grades(BVarVilSource, BLngCounter, BxChar, BResponseDimension)

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Public Function BalGetGradesCx( _
  ByVal BVarVilSourceCx As Variant, _
  ByVal BLngVilSource As Long, _
  ByRef BLngCounter As Long, _
  Optional BxChar As String = ";", _
  Optional BResponseDimension As EQLBAL_ENU_GRADESVECTORTYPE = eqlBalAs1D) As Variant

  ' Return a string with grades of normal Balmet (vilsource: vector or string, counter is byref value)
  Const sFN As String = "RESCORE::BalGetGradesCx"

  On Error GoTo EH
  Let BalGetGradesCx = Resctr_assets_bal_get_grades_cx(BVarVilSourceCx, BLngVilSource, BLngCounter, BxChar, BResponseDimension)

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function
' CORE METHODS LIST (PUBLIC) ------------------------------------------------------- '
' ================= CLASS EXCLIQ LITE RESOURCES HELPER STRUCTURE =================== '


