VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassEqlControllerProjections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ========== CLASS EXCLIQ LITE CONTROLLER PROJECTS PROJECTIONS STRUCTURE =========== '
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



' ========== CLASS EXCLIQ LITE CONTROLLER PROJECTS PROJECTIONS STRUCTURE =========== '
' ABSTRACT IMPLEMENTATIONS --------------------------------------------------------- '
' OBJECT VARIABLES (GLOOBJ_) ------------------------------------------------------- '
' VARIABLES (GLOStr_) -------------------------------------------------------------- '
' CONSTANTS ERROR MESSAGES (GLOSTR_ERR_) ------------------------------------------- '
' CONSTANTS (GLOSTR_) -------------------------------------------------------------- '
' PROPERTIES ----------------------------------------------------------------------- '
' CONSTRUCTOR/DESTRUCTOR ----------------------------------------------------------- '
' CORE METHODS PRIVATE ------------------------------------------------------------- '
' CORE METHODS LIST (PRIVATE) ------------------------------------------------------ '
' VALIDATION
' Projection_core_udf_validation
' PREPARATIONS
' RESOLVE
' Projection_core_udf_resolve
' Projection_core_udf_resolve_net_cash_flow_calculate
' MODULE RESOLUTIONS
' Projection_core_main
' Projection_core_spreadmodule_generate
' Projection_core_spreadmodule_generate_viability
' ASSETS --------------------------------------------------------------------------- '
' Projection_core_assets_get_ctzx
' ERRORS MANNAGEMENT AND ASSERTSIONS ----------------------------------------------- '
' Debug_asset
' Debug_assert_errors_set_error
' METHODS PUBLIC ------------------------------------------------------------------- '
' ABSTRACT INTERFACE UI METHODS LIST AND UDF METHODS LIST (PUBLIC) ----------------- '
' ClassEqlControllerInterface_GetSolution
' ClassEqlControllerInterface_GetSolutionUDF
' ========== CLASS EXCLIQ LITE CONTROLLER PROJECTS PROJECTIONS STRUCTURE =========== '




' ========== CLASS EXCLIQ LITE CONTROLLER PROJECTS PROJECTIONS STRUCTURE =========== '
' ABSTRACT IMPLEMENTATIONS --------------------------------------------------------- '
Implements ClassEqlControllerInterface

' OBJECT VARIABLES (GLOOBJ_) ------------------------------------------------------- '
Private RES As ClassEqlResources

' VARIABLES (GLOStr_) -------------------------------------------------------------- '
Private GLOBOO_FORMULAS As Boolean

' CONSTANTS ERROR MESSAGES (GLOSTR_ERR_) ------------------------------------------- '
Private Type GLOTYP_ERRORS
  NoDataSource As String
  NoDataZeroes As String
End Type
Private GLOTypErr As GLOTYP_ERRORS




' PROPERTIES ----------------------------------------------------------------------- '
' ABSTRACT INTERFACE PROPERTIES IMPLEMENTATION ------------------------------------- '
Public Property Set ClassEqlControllerInterface_ResourcesApp(ByRef RHS As ClassEqlResources)

  Set RES = RHS

  Let GLOBOO_FORMULAS = RES.Formulas

End Property



' CONSTRUCTOR/DESTRUCTOR ----------------------------------------------------------- '
Private Sub Class_Initialize()

  Let GLOTypErr.NoDataSource = "Faltan datos de orígen."
  Let GLOTypErr.NoDataZeroes = "Al menos un dato debe ser diferente de cero."

End Sub

Private Sub Class_Terminate()

  Set RES = Nothing

End Sub





' CORE METHODS PRIVATE ------------------------------------------------------------- '
' CORE METHODS UDF
' VALIDATION
Private Function Projection_core_udf_validation(ByVal PVarNetCashFlow As Variant, ByRef PVarErrUDF As Variant) As Boolean

  Dim VecCash As Variant
  Dim i As Long

  Const sFN As String = "PRO::-core_udf_validation"

  On Error GoTo EH
  Let Projection_core_udf_validation = False

  Let VecCash = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(PVarNetCashFlow, i, RValueType:=eqlResNumbers, RBooAbs:=True, RBooIncludeZeroes:=True)
  If i <= 0 Then Call Debug_assert_errors_set_error(PVarErrUDF, xlErrValue, 1, GLOTypErr.NoDataSource): GoTo EH
  If Application.WorksheetFunction.Sum(VecCash) = 0 Then Call Debug_assert_errors_set_error(PVarErrUDF, xlErrValue, 2, GLOTypErr.NoDataZeroes): GoTo EH
  Let VecCash = Empty

  Let Projection_core_udf_validation = True

EH:
  Let VecCash = Empty
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

' PREPARATIONS

' RESOLVE
Private Function Projection_core_udf_resolve(ByVal PVarNetCashFlow As Variant) As Variant

  On Error GoTo EH
  ' Validation
  If Not Projection_core_udf_validation(PVarNetCashFlow, Projection_core_udf_resolve) Then GoTo EH

  ' Preparation

  ' Resolve
  If Not Projection_core_udf_resolve_net_cash_flow_calculate(PVarNetCashFlow, Projection_core_udf_resolve) Then: GoTo EH

EH:
  If Not VBA.Err.Number = 0 Then Call RES.ErrorHandler.ErrorHandlerRaise("PRO::-core_udf_resolve")

End Function

Private Function Projection_core_udf_resolve_net_cash_flow_calculate( _
  ByVal PVarNetCashFlow As Variant, _
  ByRef PVarResponse As Variant) As Boolean

  Dim VecCash As Variant
  Dim i As Long
  Dim DblFlow As Double
  Dim BooViability As Boolean

  Const sFN As String = "PRO::-core_udf_resolve_net_cash_flow_calculate"

  On Error GoTo EH
  Let Projection_core_udf_resolve_net_cash_flow_calculate = False

  ' Operate
  Let VecCash = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(PVarNetCashFlow, i, RValueType:=eqlResNumbers, RBooAbs:=True, RBooIncludeZeroes:=True)
  Let DblFlow = 0
  For i = 1 To UBound(VecCash)
    Let DblFlow = DblFlow + VecCash(i)
    Let BooViability = (DblFlow >= VBA.Abs(VecCash(0)))
    If BooViability Then Exit For
  Next i

  ' Return
  ' Utilities appears in the time i
  Let PVarResponse = IIf(BooViability, "Rentable desde " & i, "Proyecto Inviable")

  Let Projection_core_udf_resolve_net_cash_flow_calculate = True

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

' MODULE RESOLUTIONS
Private Function Projection_core_main(ByRef PVarRequest As Variant, ByRef PVarResponse As Variant) As Boolean

  Dim VecProjection As Variant, VecViability As Variant
  Dim j As Long

  Const sFN As String = "PRO::-core_main"

  On Error GoTo EH
  Let Projection_core_main = False


  ' Data in BVarRequest: (0.DblIo, 1.DblVR, 2.DblEgr, 3.DblTm, 4.Dbl%, 5.VecW, 6.VecWName, 7.VecWUnitCtz, 8.VecCtz, 9.StrDivise, 10.StrTimeUnit, 11.BooVAN, 12.BooTIR, 13.BooGraph, 14.BooGraphSheet, 15.BooNewSheet, 16.RngBox, (17.WeightsLng))
  ' Resolve Sources
  If Not Projection_core_spreadmodule_generate(PVarRequest, VecProjection, j) Then GoTo EH

  ' Resolve Regression
  If Not Projection_core_spreadmodule_generate_viability(PVarRequest, VecProjection, VecViability, j) Then GoTo EH

  ' Return
  Let PVarResponse = Array(VecProjection, VecViability)
  Let Projection_core_main = True

EH:
  Let VecProjection = Empty
  Let VecViability = Empty
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Projection_core_spreadmodule_generate( _
  ByRef PVarRequest As Variant, _
  ByRef PVarProjection As Variant, _
  ByRef PxJ As Long)

  Dim REG As ClassEqlControllerRegression

  Dim RngBox As Range
  Dim VecTime As Variant, VecWeights As Variant, VecWName As Variant, VecCtz As Variant, VecUnitCtz As Variant, VecCtzX As Variant
  Dim VecHeader As Variant, VecAux1 As Variant, VecWAcum As Variant
  Dim StrDivise As String, StrTimeUnit As String, StrAux As String
  Dim DblIo As Double, DblVR As Double, DblOut As Double, DblTime As Double, DblRec As Double
  Dim i As Long, j As Long, k As Long, m As Long, n As Long, u As Long, v As Long, w As Long
  Dim xWLen As Long, xWLen2 As Long

  Const STR_DASH As String = "-"
  Const sFN As String = "PRO::-core_main_spreadmodule_generate"

  ' Data in PVarRequest: (0.DblIo, 1.DblVR, 2.DblEgr, 3.DblTm, 4.Dbl%, 5.VecW, 6.VecWName, 7.VecWUnitCtz, 8.VecCtz, 9.StrDivise, 10.StrTimeUnit, 11.BooVAN, 12.BooTIR, 13.BooGraph, 14.BooGraphSheet, 15.BooNewSheet, 16.RngBox, (17.WeightsLng))
  On Error GoTo EH
  Let Projection_core_spreadmodule_generate = False

  ' Unpack data
  Let DblIo = VBA.Abs(VBA.CDbl(PVarRequest(0))) * -1
  Let DblVR = VBA.CDbl(PVarRequest(1))
  Let DblOut = VBA.CDbl(PVarRequest(2))
  Let DblTime = VBA.CDbl(PVarRequest(3))
  Let DblRec = VBA.CDbl(PVarRequest(4))

  Let StrDivise = VBA.CStr(PVarRequest(9))
  Let StrTimeUnit = VBA.CStr(PVarRequest(10))

  Let VecWeights = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(VBA.CStr(PVarRequest(5)), i, RValueType:=eqlResNumbers, RBooAbs:=True)
  Let VecWName = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(VBA.CStr(PVarRequest(6)), i, RValueType:=eqlResNotNumbers, RValueAs:=eqlResNotNumbers)
  Let VecUnitCtz = RES.ArrayGetDataXCharSeparatedStringFromVariantsAsVector(VBA.CStr(PVarRequest(7)), i, RValueType:=eqlResNotNumbers, RStrCustomRegex:="%", RValueAs:=eqlResNotNumbers)
  Let VecCtz = VBA.Split(PVarRequest(8), "/") ' Each Ctz comes as: 1;2;3/1,2;2;4,5/4;5;9 etc...

  If GLOBOO_FORMULAS Then Set RngBox = PVarRequest(16)

  ' Get X list from VecCtz
  Let VecCtzX = Projection_core_assets_get_ctzx(VecCtz)

  ' Get Time vector
  ReDim VecTime(0)
  For i = 0 To DblTime
    ReDim Preserve VecTime(i)
    Let VecTime(i) = i
  Next i

  ' Calculate the final matrix dimenssion: Rows[1: Header], Cols[1: Time, Io; W*2: Wi, Ctzi; 4: Outs, ...]
  Let m = 1 + UBound(VecTime)
  Let n = 1 + ((UBound(VecWeights) + 1) * 2) + 4

  ' Fill Vector of Headers
  Let VecAux1 = Array("Tiempo [" & StrTimeUnit & "]", "Inversión [" & StrDivise & "]", "Cotización", "Valor Bruto", _
                      "Egresos [" & DblOut & "%]", "Utilidad Neta [" & StrDivise & "]", "Valor Residual [" & StrDivise & "]", "Flujo Neto de Caja [" & StrDivise & "]")
  ReDim VecHeader(0)
  Let xWLen = 1 + (UBound(VecWeights) + 1) ' Ctz
  Let xWLen2 = (xWLen + (UBound(VecWeights) + 1)) ' Weights + Ctz
  For i = 0 To n
    ReDim Preserve VecHeader(i)
    If i <= 1 Then ' Time, Io
      Let VecHeader(i) = VecAux1(k)
      Let k = k + 1
    ElseIf i >= 2 And i <= xWLen Then ' Ctz
      ' Cotización [OT/USD]
      Let VecHeader(i) = VecAux1(k) & "' " & VecWName(j) & "[" & StrDivise & "/" & VecUnitCtz(j) & "]"
      Let j = j + 1
      If i = xWLen Then Let k = k + 1: Let j = 0
    ElseIf i > xWLen And i <= xWLen2 Then  ' W
      ' Valor Bruto Ag
      Let VecHeader(i) = VecAux1(k) & " " & VecWName(j) & " [" & StrDivise & "]"
      Let j = j + 1
      If i = xWLen2 Then Let k = k + 1
    Else
      Let VecHeader(i) = VecAux1(k)
      Let k = k + 1
    End If
  Next i
  Let VecAux1 = Empty

  ' Fill with final values PVarProjection  VecHeader
  Let k = 0
  Let u = 1
  Let v = 0
  ReDim PVarProjection(0 To m, 0 To n)

  For j = 0 To n
    For i = 0 To m
      If i = 0 Then
        Let PVarProjection(i, j) = VecHeader(j) ' Headers
      ElseIf j = 0 Then
        Let PVarProjection(i, j) = VecTime(i - 1) ' Time
      ElseIf j = 1 Then
        Let PVarProjection(i, j) = IIf(i = j, DblIo, STR_DASH) ' Io
      End If
    Next i
  Next j

  For j = 2 To xWLen ' Ctz
    For i = 1 To m
      If i = 1 Then
        Let PVarProjection(i, j) = STR_DASH
        Let u = 1
      Else
        If GLOBOO_FORMULAS Then
          With RngBox.Range("A2")
            Let PVarProjection(i, j) = "=EQL_REGRESION_LINEAL(""" & VecCtz(k) & """, """ & VecCtzX(k) & """, " & RngBox.Offset(i + 1, 0).Address & ")"
          End With
        Else
          Set REG = New ClassEqlControllerRegression
          Set REG.ClassEqlControllerInterface_ResourcesApp = RES
          Let PVarProjection(i, j) = REG.RegressionGetY(VBA.CStr(VecCtz(k)), VBA.CStr(VecCtzX(k)), VBA.CStr(VecTime(u)))
          Set REG = Nothing
        End If
        Let u = u + 1
        If u = m Then Let k = k + 1
      End If
    Next i
  Next j

  For j = (xWLen + 1) To xWLen2 ' VB
    For i = 1 To m
      If i = 1 Then
        Let PVarProjection(i, j) = STR_DASH
        Let u = 0
      Else
        If GLOBOO_FORMULAS Then
          With RngBox.Range("A2")
            Let StrAux = VBA.CStr(VBA.Replace(VecWeights(u), ",", "."))
            Let PVarProjection(i, j) = "=" & StrAux & "*" & RngBox.Offset(i + 1, (j - xWLen + 1)).Address
          End With
        Else
          If v = 0 Then ' Initialize just once
            ReDim VecWAcum(m - 2)
            For v = 0 To UBound(VecWAcum)
              Let VecWAcum(v) = 0 ' Zero to initialize each empty space
            Next v
          End If
          Let PVarProjection(i, j) = VecWeights(u) * PVarProjection(i, (j - xWLen + 1))
          Let VecWAcum(i - 2) = VecWAcum(i - 2) + PVarProjection(i, j)
        End If
      End If
    Next i
    Let u = u + 1
  Next j

  For j = (xWLen2 + 1) To (xWLen2 + 1) ' Outs
    For i = 1 To m
      If i = 1 Then
        Let PVarProjection(i, j) = STR_DASH
        Let u = 0
      Else
        If GLOBOO_FORMULAS Then
          With RngBox.Range("A2")
            Let StrAux = VBA.CStr(VBA.Replace(DblOut, ",", "."))
            Let PVarProjection(i, j) = "=SUM(" & RngBox.Offset(i + 1, (j - xWLen + 1)).Resize(1, xWLen - 1).Address & ")*" & StrAux & "/100"
          End With
        Else
          Let PVarProjection(i, j) = VBA.CDbl(VecWAcum(u)) * DblOut / 100
          Let u = u + 1
        End If
      End If
    Next i
  Next j

  For j = (xWLen2 + 2) To (xWLen2 + 2) ' NU
    For i = 1 To m
      If i = 1 Then
        Let PVarProjection(i, j) = STR_DASH
      Else
        If GLOBOO_FORMULAS Then
          With RngBox.Range("A2")
            Let PVarProjection(i, j) = "=SUM(" & RngBox.Offset(i + 1, (j - xWLen)).Resize(1, xWLen - 1).Address & ")-" & RngBox.Offset(i + 1, (j - 1)).Address
          End With
        Else
          Let PVarProjection(i, j) = VecWAcum(i - 2) - VBA.CDbl(PVarProjection(i, j - 1))
        End If
      End If
    Next i
  Next j

  For j = (xWLen2 + 3) To (xWLen2 + 3) ' VR
    For i = 1 To m
      Let PVarProjection(i, j) = IIf(i = 1, STR_DASH, DblVR)
    Next i
  Next j

  For j = (xWLen2 + 4) To (xWLen2 + 4) ' FNC
    For i = 1 To m
      If i = 1 Then
        Let PVarProjection(i, j) = DblIo
      Else
        If GLOBOO_FORMULAS Then
          With RngBox.Range("A2")
            Let PVarProjection(i, j) = "=SUM(" & .Offset(i, (j - 2)).Resize(1, 2).Address & ")"
          End With
        Else
          Let PVarProjection(i, j) = PVarProjection(i, j - 2) + PVarProjection(i, j - 1)
        End If
      End If
    Next i
  Next j

  ReDim Preserve PVarRequest(UBound(PVarRequest) + 1)
  Let PVarRequest(UBound(PVarRequest)) = xWLen

  ' Return
  Let PxJ = j
  Let Projection_core_spreadmodule_generate = True

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
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Projection_core_spreadmodule_generate_viability( _
  ByRef PVarRequest As Variant, _
  ByVal PVarProjection As Variant, _
  ByRef PVarViability As Variant, _
  ByRef PxJ As Long)

  Const sFN As String = "PRO::-core_main_spreadmodule_generate_viability"

  ' Data in BVarRequest: (0.DblIo, 1.DblVR, 2.DblEgr, 3.DblTm, 4.Dbl%, 5.VecW, 6.VecWName, 7.VecWUnitCtz, 8.VecCtz, 9.StrDivise, 10.StrTimeUnit, 11.BooVAN, 12.BooTIR, 13.BooGraph, 14.BooGraphSheet, 15.BooNewSheet, 16.RngBox, (17.WeightsLng))
  On Error GoTo EH
  Let Projection_core_spreadmodule_generate_viability = False

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
  Let mJ = VBA.CLng(PVarRequest(3)) + 1 ' time+1
  If GLOBOO_FORMULAS Then Set RngBox = PVarRequest(16)

  ' Requirements: m+4, n+1, rngbox
  Let m = 1 + IIf(BooVAN, 1, 0) + IIf(BooTIR, 1, 0)
  ReDim PVarViability(0 To m, 0 To 1)
  Let PVarViability(0, 0) = "Indicadores" ' Headers
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
    ReDim VecPro(UBound(PVarProjection) - 1)
    For i = 1 To UBound(PVarProjection)
      Let VecPro(i - 1) = PVarProjection(i, UBound(PVarProjection, 2))
    Next i
    Let VecPro(0) = VBA.Abs(VecPro(0)) * -1
    Let StrPro = VBA.Join(VecPro, ";")
    Let PVarViability(1, 1) = Projection_core_udf_resolve(StrPro) & " " & StrTimeUnit
    If BooVAN Then Let PVarViability(2, 1) = Application.WorksheetFunction.NPV(DblRat, VecPro)
    If BooTIR Then Let PVarViability(3, 1) = Application.WorksheetFunction.IRR(VecPro, DblRat)
  End If

  Let Projection_core_spreadmodule_generate_viability = True

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function



' ASSETS --------------------------------------------------------------------------- '
Private Function Projection_core_assets_get_ctzx(ByVal PVarCtz As Variant) As Variant

  Dim VecAux1 As Variant, VecAux2 As Variant
  Dim i As Long, j As Long, k As Long

  Const sFN As String = "PRO::-core_assets_get_ctzx"

  On Error GoTo EH

  ReDim VecAux2(0)
  For i = 0 To UBound(PVarCtz)
    Let VecAux1 = VBA.Split(PVarCtz(i), ";")
    Let k = UBound(VecAux1)
    ReDim VecAux1(0)
    For j = 0 To k
      ReDim Preserve VecAux1(j)
      Let VecAux1(j) = (j - k)
    Next j
    ReDim Preserve VecAux2(i)
    Let VecAux2(i) = VBA.Join(VecAux1, ";")
  Next i

  Let Projection_core_assets_get_ctzx = VecAux2

EH:
  Let VecAux1 = Empty
  Let VecAux2 = Empty
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function


' ERRORS MANNAGEMENT AND ASSERTSIONS ----------------------------------------------- '
Private Sub Debug_assert(ByVal Cond As Boolean, Optional Origin As String = VBA.vbNullString)

  #If Debbuging Then
    Debug.Assert Cond
    If Not Cond Then Debug.Print "Assert on: " & Origin
  #End If

End Sub

' Errors mannagement
Private Sub Debug_assert_errors_set_error( _
  ByRef PVarErrUDF As Variant, _
  Optional PErrXL As XlCVError, _
  Optional PErrNum As Single, _
  Optional PErrMsg As String)

  Const sFN As String = "PRO::-core_assets_set_error"

  On Error GoTo EH

  Let PVarErrUDF = PErrXL
  Call RES.ErrorHandler.ErrorHandlerSet(PErrNum, PErrMsg)

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Sub




' METHODS PUBLIC ------------------------------------------------------------------- '
' ABSTRACT INTERFACE UI METHODS LIST AND UDF METHODS LIST (PUBLIC) ----------------- '
Public Function ClassEqlControllerInterface_GetSolution(ByVal CStrRequester As String, ByRef CVarRequest As Variant, ByRef CVarResponse As Variant) As Boolean

  On Error GoTo EH
  Let ClassEqlControllerInterface_GetSolution = Projection_core_main(CVarRequest, CVarResponse)

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise("PRO::IGetSolution")

End Function

Public Function ClassEqlControllerInterface_GetSolutionUDF(ByVal CStrRequester As String, ByVal CVarRequest As Variant) As Variant

  On Error GoTo EH
  Let ClassEqlControllerInterface_GetSolutionUDF = Projection_core_udf_resolve(CVarRequest)

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise("PRO::IGetSolution")

End Function
' METHODS PUBLIC ------------------------------------------------------------------- '
' ========== CLASS EXCLIQ LITE CONTROLLER PROJECTS PROJECTIONS STRUCTURE =========== '

