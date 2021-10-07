VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassEqlViewPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ===================== CLASS EXCLIQ LITE VIEW PRINT STRUCTURE ===================== '
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


' ===================== CLASS EXCLIQ LITE VIEW PRINT STRUCTURE ===================== '
' CORE METHODS LIST (PRIVATE) ------------------------------------------------------ '
' Viewprint_regression
' Viewprint_balmet
' Viewprint_projection
' Viewprint_oreinvoice
' Viewprint_cncinvoice
' Viewprint_help
' ASSETS ----------------------------------------------------------------------------'
' Viewprint_core_assets_getheaderandfooter
' Viewprint_core_assets_gethelpcontent
' ERRORS MANNAGEMENT AND ASSERTSIONS ------------------------------------------------'
' Debug_assert
' CORE METHODS LIST (PUBLIC) ------------------------------------------------------- '
' ViewPrint
' ===================== CLASS EXCLIQ LITE VIEW PRINT STRUCTURE ===================== '



' ===================== CLASS EXCLIQ LITE VIEW PRINT STRUCTURE ===================== '
' CORE METHODS LIST (PRIVATE) ------------------------------------------------------ '
Private Function Viewprint_regression( _
  ByVal VRES As ClassEqlResources, _
  ByVal VMDL As ClassEqlModel, _
  ByVal BVarReq As Variant, _
  ByVal BVarRes As Variant) As Boolean

  Dim VWP As ClassEqlViewParent
  Dim RngBox As Range
  Dim ShtParentRngBox As Worksheet
  Dim LO As ListObject
  Dim RngAux As Range
  Dim VarGraphData As Variant

  Dim StrTableName As String, StrAux1 As String, StrAux2 As String
  Dim BooGraph As Boolean, BooGraphSheet As Boolean, BooSrc As Boolean, BooCor As Boolean, BooReg As Boolean, BooJustRegression As Boolean
  Dim BooTableRowStripe As Boolean
  Dim BytTables As Byte, BytTableStyle As Byte
  Dim m As Long, n As Long, i As Long, MG As Long

  Const sFN As String = "VIEWPRINT::-regression_print"

  On Error GoTo EH
  Let Viewprint_regression = False

  'Unpack data
  Let BooSrc = VBA.CBool(BVarRes(0))
  Let BooCor = VBA.CBool(BVarRes(2))
  Let BooReg = VBA.CBool(BVarReq(4))
  Set RngBox = BVarReq(16)

  Set VWP = New ClassEqlViewParent
  Let BytTables = VMDL.Tables
  Let BytTableStyle = VMDL.TableStyles
  Let BooTableRowStripe = VMDL.TableRowStripes

  'Draw tables on spreadsheet
  'Sources
  Let n = IIf(BooSrc, 1, 0)
  If BooSrc Then
    Let m = UBound(BVarRes(3)) + 1
    Let n = n + UBound(BVarRes(3), 2)
    Let RngBox.Range("A2").Resize(m, n).Value2 = BVarRes(3)
    Set ShtParentRngBox = RngBox.Parent
    With ShtParentRngBox
      Let StrTableName = RngBox.Parent.Name & "basereg"
      Set LO = .ListObjects.Add(xlSrcRange, RngBox.Range("A2").CurrentRegion, XlListObjectHasHeaders:=xlYes)

      #If Debugging Then
        Call Debug_assert(Not LO Is Nothing, sFN)
      #End If

      If LO Is Nothing Then Call VBA.Err.Raise(CUSTOM_ERROR_VWS, sFN, "Se requiere un objeto ListObject")
      With LO
        Let .Name = StrTableName
        Let .ShowTotals = True
        For i = 2 To .ListColumns.Count
          Let .ListColumns.Item(i).TotalsCalculation = xlTotalsCalculationSum
        Next i
        Let .DataBodyRange.NumberFormat = "#0.00"
        Let .ListColumns.Item(1).DataBodyRange.NumberFormat = "0"
        Let .TotalsRowRange.NumberFormat = "#0.00"
        Let .HeaderRowRange.HorizontalAlignment = xlRight
        For Each RngAux In .HeaderRowRange
          If RngAux.Value = "X2" Or RngAux.Value = "Y2" Then Let RngAux.Characters(Start:=2, length:=2).Font.Superscript = True
        Next RngAux
        Set RngAux = Nothing
        Call VWP.ViewsTableStylesLOS(VRES, ShtParentRngBox, LO.Name, False, bHasTitle:=True, VTableType:=BytTables, VTableStyle:=BytTableStyle, VTableRowStripe:=BooTableRowStripe)
      End With
    End With
    Set LO = Nothing
    Set ShtParentRngBox = Nothing
    Let RngBox.Range("A1").Value2 = "Regresión Lineal - Proceso"
  End If

  'Extrapolation
  Let n = n + IIf(BooSrc, UBound(BVarRes(4), 2), 0)
  Let RngBox.Range("A2").Offset(0, n).Resize(UBound(BVarRes(4)) + 1, 2).Value2 = BVarRes(4)
  Set ShtParentRngBox = RngBox.Parent
  With ShtParentRngBox
    Let StrTableName = RngBox.Parent.Name & "extrareg"
    Set LO = .ListObjects.Add(xlSrcRange, RngBox.Range("A2").Offset(0, n).CurrentRegion, XlListObjectHasHeaders:=xlYes)

    #If Debugging Then
      Call Debug_assert(Not LO Is Nothing, sFN)
    #End If

    If LO Is Nothing Then Call VBA.Err.Raise(CUSTOM_ERROR_VWS, sFN, "Se requiere un objeto ListObject")
    With LO
      Let .Name = StrTableName
      Let .DataBodyRange.NumberFormat = "#0.00"
      Let .HeaderRowRange.HorizontalAlignment = xlRight
      Call VWP.ViewsTableStylesLOS(VRES, ShtParentRngBox, LO.Name, True, bHasTitle:=True, VTableType:=BytTables, VTableStyle:=BytTableStyle, VTableRowStripe:=BooTableRowStripe)
    End With
  End With
  Set LO = Nothing
  Set ShtParentRngBox = Nothing
  Let RngBox.Range("A1").Offset(0, n).Value2 = "Extrapolación"
  Let n = n + 2

  'Correlations
  If BooCor Then
    Let m = UBound(BVarRes(5)) + 1
    Let n = n + UBound(BVarRes(5), 2)
    Let RngBox.Range("A2").Offset(0, n).Resize(m, 2).Value2 = BVarRes(5)
    Set ShtParentRngBox = RngBox.Parent
    With ShtParentRngBox
      Let StrTableName = RngBox.Parent.Name & "resultreg"
      Set LO = .ListObjects.Add(xlSrcRange, RngBox.Range("A2").Offset(0, n).CurrentRegion, XlListObjectHasHeaders:=xlYes)

      #If Debugging Then
        Call Debug_assert(Not LO Is Nothing, sFN)
      #End If

      If LO Is Nothing Then Call VBA.Err.Raise(CUSTOM_ERROR_VWS, sFN, "Se requiere un objeto ListObject")
      With LO
        Let .Name = StrTableName
        Let .ListColumns.Item(2).DataBodyRange.NumberFormat = "#0.00"
        For Each RngAux In .ListColumns.Item(1).DataBodyRange
          If RngAux.Value = "Coeficiente de determinación (r2):" Then Let RngAux.Characters(Start:=32, length:=1).Font.Superscript = True
          If RngAux.Value = "Error estándar de la determinación (Se):" Then Let RngAux.Characters(Start:=38, length:=1).Font.Subscript = True
        Next RngAux
        Let .ListColumns.Item(1).DataBodyRange.Font.Italic = True
        Set RngAux = Nothing
        Call VWP.ViewsTableStylesLOS(VRES, ShtParentRngBox, .Name, True, bHasTitle:=True, bColsAutofit:=True, VTableType:=BytTables, VTableStyle:=BytTableStyle, VTableRowStripe:=BooTableRowStripe)
      End With
    End With
    Set LO = Nothing
    Set ShtParentRngBox = Nothing
    Let RngBox.Range("A1").Offset(0, n).Value2 = "Ecuación de Tendencia Y' = a + b(X')"
    Let n = n + 2
  End If

  'Optional: Graphic
  'Incoming data: 0-X '1-Y '2-E '3-ShowAll '4-JustRegression '5-X2 '6-Y2 '7-XY '8-a '9-b '10-r '11-r2 '12-Se '13-Graph '14-SheetGraph '15-NewSheet '13-RngBox
  Let BooGraph = VBA.CBool(BVarReq(13))
  Let BooGraphSheet = VBA.CBool(BVarReq(14)) 'T: Same, F New GraphSheet

  If BooGraph Then
    Let BooJustRegression = (BooReg Or BooSrc = False)
    If BooJustRegression Then
      ReDim VarGraphData(1)
      Let StrAux1 = VBA.Replace(VBA.CStr(BVarReq(0)), ",", ".")
      Let StrAux1 = VBA.Replace(StrAux1, ";", ",")
      Let StrAux2 = VBA.Replace(VBA.CStr(BVarReq(1)), ",", ".")
      Let StrAux2 = VBA.Replace(StrAux2, ";", ",")
      Let VarGraphData(0) = StrAux1 'VBA.Replace(VBA.CStr(BVarReq(0)), ",", ".")
      Let VarGraphData(1) = StrAux2 'VBA.Replace(VBA.CStr(BVarReq(1)), ",", ".")
    Else
      Let MG = UBound(BVarRes(3))
      Set VarGraphData = RngBox.Range("B3").Resize(MG, 2)
    End If

    Let n = n + 1
    Call VWP.ViewsSetGraphic( _
      VRES, _
      RngBox.Parent, _
      BooGraphSheet, _
      VarGraphData, _
      RngBox.Range("A1").Offset(0, n), _
      "Gráfico Regresión", _
      "Regresión Lineal", _
      aTrendLines:=True, _
      dEquation:=True, _
      dR2:=True, _
      lineMinor:=Not BooGraphSheet, _
      lineMajor:=True, _
      VBooDataAsString:=BooJustRegression, _
      VVecStrData:=VarGraphData)
  End If

  Let Viewprint_regression = True

EH:
  Set RngAux = Nothing
  Set RngBox = Nothing
  Set LO = Nothing
  Set VWP = Nothing
  Call VRES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Viewprint_balmet( _
  ByVal VRES As ClassEqlResources, _
  ByVal VMDL As ClassEqlModel, _
  ByVal BVarReq As Variant, _
  ByVal BVarBALMET As Variant) As Boolean

  Dim VWP As ClassEqlViewParent
  Dim LO As ListObject
  Dim RngBox As Range
  Dim VecAux0 As Variant
  Dim StrNameA As String, StrUnitA As String, StrTName As String
  Dim BOOFORMULAS As Boolean, BooHeads As Boolean, BooPercents As Boolean
  Dim xProducts As Long, LngPercents As Long
  Dim SngTables As Single
  Dim BooTableRowStripe As Boolean
  Dim BytTables As Byte, BytTableStyle As Byte
  Dim EnuMethod As EQLBAL_ENU_METHOD, EnuTypeFeed As EQLBAL_ENU_TYPECONTENT

  Const sFN As String = "VIEWPRINT::-balmet"

  On Error GoTo EH
  Let Viewprint_balmet = False

  #If Debugging Then
    Call Debug_assert(VBA.IsArray(BVarBALMET), sFN)
  #End If

  Set VWP = New ClassEqlViewParent
  Let BytTables = VMDL.Tables
  Let BytTableStyle = VMDL.TableStyles
  Let BooTableRowStripe = VMDL.TableRowStripes

  ' Data required in BVarRequest: (0.DblFeed, 1.DblFeedVol, 2.VecGrad, 3.VecGradCx, 4.VecName, 5.VecUnit, 6.BytMethod, 7.Booleans, 8.EnuTypeFeed, 9.RngBox (, 10.StrFeedName, 11.StrNameA, 12.StrUnitA))
  ' Booleans values: 0.BooAll, 1.BooPercents, 2.BooUnities, 3.BooFines, 4.BooGrams, 5.BooOT, 6.BooRatio, 7.BooHeads (, 8.Formulas, 9.Tables)
  Let EnuMethod = IIf(BVarReq(6) = 0, eqlBalConventional, IIf(BVarReq(6) = 1, eqlBalCramer, eqlBalInverseMatrix))
  Let EnuTypeFeed = IIf(BVarReq(8) = 0, eqlBalSolids, IIf(BVarReq(8) = 1, eqlBalVolume, eqlBalBoth))
  Let VecAux0 = BVarReq(7)
  Let BooPercents = VBA.CBool(VecAux0(1))
  Let BooHeads = VBA.CBool(VecAux0(7))
  Set RngBox = BVarReq(9)
  Let StrNameA = VBA.CStr(BVarReq(11))
  Let StrUnitA = VBA.CStr(BVarReq(12))

  Let BOOFORMULAS = VBA.CBool(VecAux0(8))
  Let SngTables = VBA.CByte(VecAux0(9))
  Let VecAux0 = Empty

  With RngBox

    ' Set Balmet in range
    Call VRES.ArrayNumbersCommasToDots(BVarBALMET)
    Let .Range("A2").Resize(UBound(BVarBALMET) + 1, UBound(BVarBALMET, 2) + 1).Value2 = BVarBALMET

    ' Set Balmet range to LOTable
    With .Parent
      Let StrTName = "Balance"
      Set LO = .ListObjects.Add(xlSrcRange, RngBox.Range("A2").CurrentRegion, XlListObjectHasHeaders:=xlYes)
      #If Debugging Then
        Call Debug_assert(Not LO Is Nothing, sFN)
      #End If
      If LO Is Nothing Then Call VBA.Err.Raise(CUSTOM_ERROR_VWS, sFN, "Se requiere un objeto ListObject")

      With LO

        Let .Name = StrTName & "_" & .Name
        Let .DataBodyRange.NumberFormat = "#0.00"
        Let .ListColumns.Item(1).DataBodyRange.NumberFormat = "0"
        Let .HeaderRowRange.HorizontalAlignment = xlRight
        Call VWP.ViewsTableStylesLOS(VRES, RngBox.Parent, LO.Name, False, bHasTitle:=True, VTableType:=BytTables, VTableStyle:=BytTableStyle, VTableRowStripe:=BooTableRowStripe)

      End With

      Set LO = Nothing
    End With

    ' Formulas case
    Let xProducts = UBound(BVarBALMET) - 1 + IIf(BooHeads, -2, 0)
    If BOOFORMULAS Then
      With .Range("C3")
        If Not SngTables = 2 Then

          Dim StrFeed As String, StrGrades As String, StrFormulaArray As String
          Let StrFeed = .Offset(xProducts, 0).Address(False, False)
          Let LngPercents = IIf(BooPercents, 2, 1) + IIf(EnuTypeFeed = eqlBalBoth, 2, 0)
          Let StrGrades = .Offset(0, LngPercents).Resize(xProducts + 1, xProducts - 1).Address(False, False)

          ' EQL_BALMET_CURRENT( _
              ByVal Alimentacion As Variant, _
              ByVal LeyesConcentrado As Variant, _
              Optional NombresElementos_Opcional As Variant = VBA.vbNullString, _
              Optional UnidadesLey_Opcional As Variant = "%", _
              Optional Metodo_Opcional As Long = 0, _
              Optional ResultadoHorizontal_Opcional As Boolean = False, _
              Optional IncluirAlimentacion_Opcional As Boolean = False) As Variant

          Let StrFormulaArray = "=+EQL_BALMET(" & StrFeed & "," & StrGrades & ",""" & VBA.CStr(StrNameA) & """,""" & VBA.CStr(StrUnitA) & """," & EnuMethod & ")"
          Let .Resize(xProducts, 1).FormulaArray = StrFormulaArray

        End If
      End With
    End If

    ' Set final title
    Let .Range("A1").Value = "Balance Metalúrgico - " & xProducts & " Productos"
  End With

  Let Viewprint_balmet = True

EH:
  Set LO = Nothing
  Set VWP = Nothing
  Call VRES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Viewprint_projection( _
  ByVal VRES As ClassEqlResources, _
  ByVal VMDL As ClassEqlModel, _
  ByVal PVarRequest As Variant, _
  ByVal PVarProjection As Variant) As Boolean

  ' Data in BVarRequest: (0.DblIo, 1.DblVR, 2.DblEgr, 3.DblTm, 4.Dbl%, 5.VecW, 6.VecWName, 7.VecWUnitCtz, 8.VecCtz, 9.StrDivise, 10.StrTimeUnit, 11.BooVAN, 12.BooTIR, 13.BooGraph, 14.BooGraphSheet, 15.BooNewSheet, 16.RngBox, (17.WeightsLng))
  Dim VWP As ClassEqlViewParent
  Dim ShtParentRngBox As Worksheet
  Dim LO As ListObject
  Dim RngBox As Range, RngAux As Range
  Dim VarGraphData As Variant

  Dim m As Long, n As Long, i As Long, MG As Long, xLastColumn As Long
  Dim StrTableName As String, StrAux1 As String, StrAux2 As String
  Dim BooGraph As Boolean, BooGraphSheet As Boolean
  Dim BooTableRowStripe As Boolean
  Dim BytTables As Byte, BytTableStyle As Byte

  Const sFN As String = "VIEWPRINT::-projection"

  On Error GoTo EH
  Let Viewprint_projection = False

  Set VWP = New ClassEqlViewParent
  Let BytTables = VMDL.Tables
  Let BytTableStyle = VMDL.TableStyles
  Let BooTableRowStripe = VMDL.TableRowStripes

  ' Unpack data
  If Not VBA.IsArray(PVarProjection) Then GoTo EH
  If Not VBA.IsArray(PVarProjection(0)) Then GoTo EH
  If Not VBA.IsArray(PVarProjection(1)) Then GoTo EH
  Let MG = VBA.CLng(PVarRequest(UBound(PVarRequest))) + 2
  Set RngBox = PVarRequest(16)

  ' Draw tables on spreadsheet
  ' Projection
  Let m = UBound(PVarProjection(0)) + 1
  Let n = UBound(PVarProjection(0), 2) + 1
  Let xLastColumn = n
  Let RngBox.Range("A2").Resize(m, n).Value2 = PVarProjection(0)
  Set ShtParentRngBox = RngBox.Parent
  With ShtParentRngBox
    Let StrTableName = RngBox.Parent.Name & "baseproject"
    Set LO = .ListObjects.Add(xlSrcRange, RngBox.Range("A2").CurrentRegion, XlListObjectHasHeaders:=xlYes)

    #If Debugging Then
      Call Debug_assert(Not LO Is Nothing, sFN)
    #End If
    If LO Is Nothing Then Call VBA.Err.Raise(CUSTOM_ERROR_VWS, sFN, "Se requiere un objeto ListObject")

    With LO
      Let .Name = StrTableName
      Let .ShowTotals = True
      For i = MG To .ListColumns.Count
        Let .ListColumns.Item(i).TotalsCalculation = xlTotalsCalculationSum
      Next i
      Let .DataBodyRange.NumberFormat = "#0.00"
      Let .ListColumns.Item(1).DataBodyRange.NumberFormat = "0"
      Let .TotalsRowRange.NumberFormat = "#0.00"
      Let .HeaderRowRange.HorizontalAlignment = xlRight
      Let .DataBodyRange.HorizontalAlignment = xlRight
      Set RngAux = Nothing
      Call VWP.ViewsTableStylesLOS(VRES, ShtParentRngBox, LO.Name, False, bHasTitle:=True, VTableType:=BytTables, VTableStyle:=BytTableStyle, VTableRowStripe:=BooTableRowStripe)
    End With
  End With
  Set LO = Nothing
  Set ShtParentRngBox = Nothing
  Let RngBox.Range("A1").Value2 = "PROYECCIÓN PROYECTO DE NEGOCIO"

  ' Viability
  Let n = n + 1
  Let RngBox.Range("A2").Offset(0, n).Resize(4, 2).Value2 = PVarProjection(1)
  Set ShtParentRngBox = RngBox.Parent
  With ShtParentRngBox
    Let StrTableName = RngBox.Parent.Name & "viabilityprojection"
    Set LO = .ListObjects.Add(xlSrcRange, RngBox.Range("A2").Offset(0, n).CurrentRegion, XlListObjectHasHeaders:=xlYes)

    #If Debugging Then
      Call Debug_assert(Not LO Is Nothing, sFN)
    #End If
    If LO Is Nothing Then Call VBA.Err.Raise(CUSTOM_ERROR_VWS, sFN, "Se requiere un objeto ListObject")

    With LO
      Let .Name = StrTableName
      Let .DataBodyRange.NumberFormat = "#0.00"
      Call VWP.ViewsTableStylesLOS(VRES, ShtParentRngBox, LO.Name, True, bHasTitle:=True, bColsAutofit:=True, VTableType:=BytTables, VTableStyle:=BytTableStyle, VTableRowStripe:=BooTableRowStripe)
    End With
  End With
  Set LO = Nothing
  Set ShtParentRngBox = Nothing
  Let RngBox.Range("A1").Offset(0, n).Value2 = "COMPORTAMIENTO"

  ' Optional: Graphic
  Let BooGraph = VBA.CBool(PVarRequest(13))
  Let BooGraphSheet = VBA.CBool(PVarRequest(14)) ' T: Same, F New GraphSheet
  Let StrAux1 = VBA.CStr(PVarRequest(9))
  Let StrAux2 = VBA.CStr(PVarRequest(10))

  If BooGraph Then
    Set VarGraphData = RngBox.Range("A3").Offset(0, xLastColumn - 1).Resize(m - 1, 1)
    Let n = n + 3
    Call VWP.ViewsSetGraphic(VRES, RngBox.Parent, BooGraphSheet, VarGraphData, RngBox.Range("A1").Offset(0, n), "Gráfico Proyección", "Flujo Neto de Caja", StrAux1, StrAux2, xlColumnClustered)
  End If

  Let Viewprint_projection = True

EH:
  Set RngAux = Nothing
  Set RngBox = Nothing
  Set LO = Nothing
  Set VWP = Nothing
  Call VRES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Viewprint_oreinvoice( _
  ByVal VRES As ClassEqlResources, _
  ByVal VMDL As ClassEqlModel, _
  ByVal PVarRequest As Variant, _
  ByVal PVarOreinvoice As Variant) As Boolean

  Dim HPR As ClassEqlControllerHelper
  Dim VWP As ClassEqlViewParent
  Dim RngBox As Range
  Dim VecAux0 As Variant
  Dim StrAux0 As String, StrHeaderTop As String, StrHeaderBtm As String, StrLogoTop As String, StrLogoBtm As String
  Dim BooTable As Boolean, BooConfigPage As Boolean, BooTwoPage As Boolean
  Dim EnuThm As XlThemeColor, EnuThmFont As XlThemeColor
  Dim m As Long, n As Long, i As Long, j As Long

  Const sFN As String = "VIEWPRINT::-oreinvoice"

  On Error GoTo EH
  Let Viewprint_oreinvoice = False

  If Not VBA.IsArray(PVarOreinvoice) Then GoTo EH

  ' Request: 0.StrPlace, 1.BooME, 2.BooVarious, 3.BooAverage, 4.BooPage, 5.BooPrintTwo, 6.BooNewSheet, 7.RngBox
  Let BooConfigPage = VBA.CBool(PVarRequest(4))
  Let BooTwoPage = VBA.CBool(PVarRequest(5))
  Set RngBox = PVarRequest(7)
  If RngBox Is Nothing Then GoTo EH
  Set VWP = New ClassEqlViewParent
  Let BooTable = VMDL.Tables
  Let EnuThm = VMDL.TableThemeColor
  Let EnuThmFont = VMDL.TableThemeFontColor
  For i = LBound(PVarOreinvoice) To UBound(PVarOreinvoice)

    Set HPR = PVarOreinvoice(i)

    With RngBox

      ' Title
      Let .Range("E1").Resize(2, 1).Value = HPR.Title
      Call VWP.ViewsTableStaticSet(VRES, .Range("A1").Resize(2, 9), True, VVarHeaderHorizontalAlignement:=xlHAlignCenter, VVarBodyHorizontalAlignement:=xlHAlignCenter, VBooBorderLines:=True, VBooTableType:=BooTable, VThmColor:=EnuThm, VThmFontColor:=EnuThmFont)
      Let .Range("E2").Font.Size = 8

      ' Purchase Data
      Set RngBox = .Range("A4").Resize(UBound(HPR.Invoice) + 1, 3)
      Let RngBox.Value = HPR.Invoice ' Pur Data
      Call VWP.ViewsTableStaticSet(VRES, RngBox, True, False, True, False, VBooBorderLines:=True, VBooHorizontalLines:=True, VBooTableType:=BooTable, VThmColor:=EnuThm, VThmFontColor:=EnuThmFont)
      Let RngBox.Range("B4").NumberFormat = "dd/mm/yy"
      Let RngBox.Range("B7").NumberFormat = "0.00"
      With RngBox.Range("B2").Validation
        Call .Delete
        Call .Add(Type:=xlValidateList, Formula1:=VBA.Replace(RngBox.Range("B2").Value, ";", ","))
        Let StrAux0 = VBA.Left(RngBox.Range("B2").Value, VBA.InStr(RngBox.Range("B2").Value, ";") - 1)
        Let RngBox.Range("B2").Value = StrAux0
        Let .ShowError = False
      End With

      ' Provider data
      Let m = UBound(HPR.Invoice) + 2 + 4
      Set RngBox = Nothing
      Set RngBox = .Range("A" & m).Resize(UBound(HPR.Provider) + 1, 3)
      Let RngBox.Value = HPR.Provider ' Provider data
      Call VWP.ViewsTableStaticSet(VRES, RngBox, True, False, True, False, "0.00", VBooBorderLines:=True, VBooHorizontalLines:=True, VBooTableType:=BooTable, VThmColor:=EnuThm, VThmFontColor:=EnuThmFont)

      ' Ore Data
      Set RngBox = Nothing
      Set RngBox = .Range("E4").Resize(UBound(HPR.Ore) + 1, 5)
      Let RngBox.Value = HPR.Ore ' Ore Data
      Call VWP.ViewsTableStaticSet(VRES, RngBox, True, True, True, True, "0.00", VBooBorderLines:=True, VBooHorizontalLines:=True, VBooVerticalLines:=True, VInsideVerticalsJustInBody:=True, VBooTableType:=BooTable, VThmColor:=EnuThm, VThmFontColor:=EnuThmFont)
      Let RngBox.Range("B2").Resize(RngBox.Rows.Count - 2, 1).NumberFormat = "dd/mm/yy"

      ' Ore weight data payable
      Let m = UBound(HPR.Ore) + 1 + 4
      Set RngBox = Nothing
      Set RngBox = .Range("E" & m).Resize(UBound(HPR.OrePayment) + 1, 5)
      Let RngBox.Value = HPR.OrePayment ' Ore weight data
      Call VWP.ViewsTableStaticSet(VRES, RngBox, True, True, True, True, "0.00", VBooBorderLines:=True, VBooHorizontalLines:=True, VBooVerticalLines:=True, VInsideVerticalsJustInBody:=True, VBooTableType:=BooTable, VThmColor:=EnuThm, VThmFontColor:=EnuThmFont)

      ' Gross Value
      Let m = m + UBound(HPR.OrePayment) + 1
      Set RngBox = Nothing
      Set RngBox = .Range("E" & m).Resize(UBound(HPR.GrossValue) + 1, 5)
      Let RngBox.Value = HPR.GrossValue ' Gross Value
      Call VWP.ViewsTableStaticSet(VRES, RngBox, True, True, True, True, "0.00", VBooBorderLines:=True, VBooHorizontalLines:=True, VBooVerticalLines:=True, VInsideVerticalsJustInBody:=True, VBooTableType:=BooTable, VThmColor:=EnuThm, VThmFontColor:=EnuThmFont)
      With RngBox.Range("B3")
        For j = LBound(HPR.Units) To UBound(HPR.Units)
          Let .Range("A" & j + 1).NumberFormat = "0.00""" & VBA.UCase(VBA.CStr(HPR.Units(j))) & """"
        Next j
      End With

      ' Deductions
      Let m = m + UBound(HPR.GrossValue) + 1
      Set RngBox = Nothing
      Set RngBox = .Range("E" & m).Resize(UBound(HPR.Deductions) + 1, 5)
      Let RngBox.Value = HPR.Deductions ' Deductions
      Call VWP.ViewsTableStaticSet(VRES, RngBox, True, False, True, True, "0.00", VBooBorderLines:=True, VBooHorizontalLines:=True, VBooTableType:=BooTable, VThmColor:=EnuThm, VThmFontColor:=EnuThmFont)

      ' Bonus
      If HPR.BooBonus Then ' If exists bonus
        Let m = m + UBound(HPR.Deductions) + 1
        Set RngBox = Nothing
        Set RngBox = .Range("E" & m).Resize(UBound(HPR.Bonus) + 1, 5)
        Let RngBox.Value = HPR.Bonus ' Bonus
        Call VWP.ViewsTableStaticSet(VRES, RngBox, True, False, True, True, "0.00", VBooBorderLines:=True, VBooHorizontalLines:=True, VBooTableType:=BooTable, VThmColor:=EnuThm, VThmFontColor:=EnuThmFont)
      End If

      ' Final
      If HPR.BooBonus Then
        Let m = m + UBound(HPR.Bonus) + 2
      Else
        Let m = m + UBound(HPR.Deductions) + 2
      End If
      Set RngBox = Nothing
      Set RngBox = .Range("A" & m).Resize(UBound(HPR.Final) + 1, 9)
      Let RngBox.Value = HPR.Final ' Final
      Call VWP.ViewsTableStaticSet(VRES, RngBox, True, True, True, False, "0.00", VBooBorderLines:=True, VBooTableType:=BooTable, VThmColor:=EnuThm, VThmFontColor:=EnuThmFont)
      With RngBox.Range("B3").Resize(2, RngBox.Columns.Count - 1)
        Call .Merge
        Let .HorizontalAlignment = xlHAlignLeft
        Let .VerticalAlignment = xlVAlignTop
      End With

      ' Signatures
      With RngBox.Range("A7").Validation ' Signatures
        Call .Delete
        Call .Add(Type:=xlValidateList, Formula1:=VBA.UCase("dpto. compras, revisado por"))
        Let .ShowError = False
        Let .Parent.Value = VBA.UCase("dpto. compras")
      End With
      Call VWP.ViewsTableStaticSet(VRES, RngBox.Range("A7:B7"), True, VBooTableType:=BooTable, VThmColor:=EnuThm, VThmFontColor:=EnuThmFont)
      With RngBox.Range("E7").Validation
        Call .Delete
        Call .Add(Type:=xlValidateList, Formula1:=VBA.UCase("cajero, pagado por"))
        Let .ShowError = False
        Let .Parent.Value = VBA.UCase("cajero")
      End With
      Call VWP.ViewsTableStaticSet(VRES, RngBox.Range("E7:F7"), True, VVarHeaderHorizontalAlignement:=xlHAlignCenter, VBooTableType:=BooTable, VThmColor:=EnuThm, VThmFontColor:=EnuThmFont)
      Let RngBox.Range("H7").Value = "PROVEDOR"
      Call VWP.ViewsTableStaticSet(VRES, RngBox.Range("H7:I7"), True, VBooTableType:=BooTable, VThmColor:=EnuThm, VThmFontColor:=EnuThmFont)

      ' Two per page:
      If VBA.CBool(PVarRequest(5)) Then
        Let m = m + 7
        Call .Range("A1").Resize(m, 9).Select
        Call .Range("A1").Resize(m, 9).Copy
        Call .Range("J1").PasteSpecial(xlPasteAll)
        Let Application.CutCopyMode = False
        Call .Range("S1").Select
        Set RngBox = .Range("S1")
      Else
        Set RngBox = .Range("J1")
      End If

      ' Page setup
      If BooConfigPage Then
        Call Viewprint_core_assets_getheaderandfooter(VRES, VMDL, StrHeaderTop, StrHeaderBtm, StrLogoTop, StrLogoBtm)
        ' Set up the page
        Call VWP.ViewsPageSetup( _
          VRES, _
          RngBox.Parent, _
          "A1", IIf(BooTwoPage, xlLandscape, xlPortrait), xlPaperLetter, 2, 1, 2, 2, 0.5, 1, _
          ShHeaderLeft:=StrHeaderTop, _
          ShHeaderRight:=StrLogoTop, _
          ShFooterLeft:=StrHeaderTop, _
          ShFooterRight:=StrLogoBtm, _
          ShImgHeader:=2, _
          ShImgFooter:=2, _
          ShImgFileHeader:=StrLogoTop, _
          ShImgFileFooter:=StrLogoBtm, _
          ShZoom:=90)
      End If

    End With

  Next i
  Set VWP = Nothing

  ' Return
  Let Viewprint_oreinvoice = True

EH:
  Set RngBox = Nothing
  Set VWP = Nothing
  Set HPR = Nothing
  Call VRES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Function Viewprint_cncinvoice( _
  ByVal VRES As ClassEqlResources, _
  ByVal VMDL As ClassEqlModel, _
  ByVal PVarRequest As Variant, _
  ByVal PVarCncinvoice As Variant) As Boolean

  Dim HPR As ClassEqlControllerHelper
  Dim VWP As ClassEqlViewParent
  Dim RngBox As Range, RngBos As Range
  Dim StrAux0 As String, StrHeaderTop As String, StrHeaderBtm As String, StrLogoTop As String, StrLogoBtm As String
  Dim BooTable As Boolean, BooConfigPage As Boolean
  Dim m As Long, n As Long, i As Long, j As Long
  Dim EnuThm As XlThemeColor, EnuThmFont As XlThemeColor

  Const PUR_WIDTH As Byte = 4
  Const INV_WIDTH As Byte = 6
  Const DOC_WIDTH As Byte = 11

  Const sFN As String = "VIEWPRINT::-cncinvoice"

  On Error GoTo EH
  Let Viewprint_cncinvoice = False

  ' Request: 0.BooPage, 1.BooNewSheet, 2.RngBox, 3.Simulation
  If Not VBA.IsArray(PVarCncinvoice) Then GoTo EH
  Let BooConfigPage = VBA.CBool(PVarRequest(0))
  Set RngBox = PVarRequest(2)
  If RngBox Is Nothing Then GoTo EH

  Set VWP = New ClassEqlViewParent
  Let BooTable = VMDL.Tables
  Let EnuThm = VMDL.TableThemeColor
  Let EnuThmFont = VMDL.TableThemeFontColor

  With RngBox

    'GoTo EH
    Set HPR = PVarCncinvoice(0)

    ' Title
    Let .Range("F1").Resize(2, 1).Value = HPR.Title
    Call VWP.ViewsTableStaticSet(VRES, .Range("A1").Resize(2, DOC_WIDTH), True, VVarHeaderHorizontalAlignement:=xlHAlignCenter, VVarBodyHorizontalAlignement:=xlHAlignCenter, VBooBorderLines:=True, VBooTableType:=BooTable, VThmColor:=EnuThm, VThmFontColor:=EnuThmFont)
    Let .Range("F2").Font.Size = 8

    ' Purchase Data
    Set RngBos = .Range("A4").Resize(UBound(HPR.Invoice) + 1, PUR_WIDTH)
    Let RngBos.Value = HPR.Invoice ' Pur Data
    Call VWP.ViewsTableStaticSet(VRES, RngBos, True, False, True, False, VBooBorderLines:=True, VBooHorizontalLines:=True, VBooTableType:=BooTable, VThmColor:=EnuThm, VThmFontColor:=EnuThmFont)
    Let RngBos.Range("B5").NumberFormat = "dd/mm/yy"
    Let RngBos.Range("B7").NumberFormat = "0.00"
    With RngBos.Range("B2").Validation
      Call .Delete
      Call .Add(Type:=xlValidateList, Formula1:=VBA.Replace(RngBos.Range("B2").Value, ";", ","))
      Let StrAux0 = VBA.Left(RngBos.Range("B2").Value, VBA.InStr(RngBos.Range("B2").Value, ";") - 1)
      Let RngBos.Range("B2").Value = StrAux0
      Let .ShowError = False
    End With

    ' Provider data
    Let m = UBound(HPR.Invoice) + 2 + 4
    Set RngBos = Nothing
    Set RngBos = .Range("A" & m).Resize(UBound(HPR.Provider) + 1, PUR_WIDTH)
    Let RngBos.Value = HPR.Provider ' Provider data
    Call VWP.ViewsTableStaticSet(VRES, RngBos, True, False, True, False, "0.00", VBooBorderLines:=True, VBooHorizontalLines:=True, VBooTableType:=BooTable, VThmColor:=EnuThm, VThmFontColor:=EnuThmFont)

    ' Cnc Data (weight)
    Set RngBos = Nothing
    Set RngBos = .Range("F4").Resize(UBound(HPR.WeightCNC) + 1, INV_WIDTH)
    Let RngBos.Value = HPR.WeightCNC ' Cnc Data
    Call VWP.ViewsTableStaticSet(VRES, RngBos, True, True, True, False, "0.00", VVarBodyHorizontalAlignement:=xlHAlignRight, VBooBorderLines:=True, VBooHorizontalLines:=True, VBooVerticalLines:=True, VInsideVerticalsJustInBody:=True, VBooTableType:=BooTable, VThmColor:=EnuThm, VThmFontColor:=EnuThmFont)
    Let RngBos.Range("B6").Resize(RngBos.Rows.Count, 1).NumberFormat = "dd/mm/yy"

    ' Cnc contents data
    Let m = UBound(HPR.WeightCNC) + 1 + 4
    Set RngBos = Nothing
    Set RngBos = .Range("F" & m).Resize(UBound(HPR.ContentsCNC) + 1, INV_WIDTH)
    Let RngBos.Value = HPR.ContentsCNC ' Cnc weight data
    Call VWP.ViewsTableStaticSet(VRES, RngBos, True, True, True, False, "0.00", VVarBodyHorizontalAlignement:=xlHAlignRight, VBooBorderLines:=True, VBooHorizontalLines:=True, VBooVerticalLines:=True, VInsideVerticalsJustInBody:=True, VBooTableType:=BooTable, VThmColor:=EnuThm, VThmFontColor:=EnuThmFont)
    With RngBos.Range("E3")
      For i = 0 To UBound(HPR.UnitsGraCNC)
        Let .Range("A" & i + 1).NumberFormat = "0.00""" & VBA.UCase(VBA.CStr(HPR.UnitsGraCNC(i))) & """"
        Let .Range("B" & i + 1).NumberFormat = "0.00""" & VBA.UCase(VBA.CStr(HPR.UnitsCtzCNC(i))) & """"
      Next i
    End With

    ' Asessment (Gross Value)
    Let m = m + UBound(HPR.ContentsCNC) + 1
    Set RngBos = Nothing
    Set RngBos = .Range("F" & m).Resize(UBound(HPR.AssessmentCNC) + 1, INV_WIDTH)
    Let RngBos.Value = HPR.AssessmentCNC ' Gross Value
    Call VWP.ViewsTableStaticSet(VRES, RngBos, True, True, True, True, "0.00", VVarBodyHorizontalAlignement:=xlHAlignRight, VBooBorderLines:=True, VBooHorizontalLines:=True, VBooTableType:=BooTable, VThmColor:=EnuThm, VThmFontColor:=EnuThmFont)

    Set RngBos = Nothing
    Set RngBos = .Range("F" & m).Resize(UBound(HPR.ContentsCNC) + 2, INV_WIDTH)
    Call VWP.ViewsTableStaticSet(VRES, RngBos, True, True, True, True, "0.00", VVarBodyHorizontalAlignement:=xlHAlignRight, VBooTableType:=BooTable, VThmColor:=EnuThm, VThmFontColor:=EnuThmFont)

    With RngBos.Range("B3")
      For i = 0 To UBound(HPR.UnitsTraCNC)
        Let .Range("A" & i + 1).NumberFormat = "0.00""" & VBA.UCase(VBA.CStr(HPR.UnitsTraCNC(i))) & """"
      Next i
    End With

    ' Deductions
    If HPR.BooDeductionsCNC Then ' If exists deductions
      Let m = m + UBound(HPR.AssessmentCNC) + 1
      Set RngBos = Nothing
      Set RngBos = .Range("F" & m).Resize(UBound(HPR.Deductions) + 1, INV_WIDTH)
      Let RngBos.Value = HPR.Deductions ' Deductions
      Call VWP.ViewsTableStaticSet(VRES, RngBos, True, True, True, True, "0.00", VVarBodyHorizontalAlignement:=xlHAlignRight, VBooBorderLines:=True, VBooHorizontalLines:=True, VBooTableType:=BooTable, VThmColor:=EnuThm, VThmFontColor:=EnuThmFont)
    End If

    ' Bonus
    If HPR.BooBonus Then ' If exists bonus
      Let m = m + UBound(HPR.Deductions) + 1
      Set RngBos = Nothing
      Set RngBos = .Range("F" & m).Resize(UBound(HPR.Bonus) + 1, INV_WIDTH)
      Let RngBos.Value = HPR.Bonus ' Bonus
      Call VWP.ViewsTableStaticSet(VRES, RngBos, True, False, True, True, "0.00", VBooBorderLines:=True, VBooHorizontalLines:=True, VBooTableType:=BooTable, VThmColor:=EnuThm, VThmFontColor:=EnuThmFont)
    End If

    ' Final
    If HPR.BooBonus Then
      Let m = m + UBound(HPR.Bonus) + 2
    Else
      If HPR.BooDeductionsCNC Then
        Let m = m + UBound(HPR.Deductions) + 2
      Else
        Let m = m + UBound(HPR.AssessmentCNC) + 2
      End If
    End If
    Set RngBos = Nothing
    Set RngBos = .Range("A" & m).Resize(UBound(HPR.Final) + 1, DOC_WIDTH)
    Let RngBos.Value = HPR.Final ' Final
    Call VWP.ViewsTableStaticSet(VRES, RngBos, True, True, True, False, "0.00", VBooBorderLines:=True, VBooTableType:=BooTable, VThmColor:=EnuThm, VThmFontColor:=EnuThmFont)
    With RngBos.Range("B3").Resize(2, RngBos.Columns.Count - 1)
      Call .Merge
      Let .HorizontalAlignment = xlHAlignLeft
      Let .VerticalAlignment = xlVAlignTop
    End With

    ' Signatures
    With RngBos.Range("A7").Validation ' Signatures
      Call .Delete
      Call .Add(Type:=xlValidateList, Formula1:=VBA.UCase("dpto. compras, revisado por"))
      Let .ShowError = False
      Let .Parent.Value = VBA.UCase("dpto. compras")
      Call VWP.ViewsTableStaticSet(VRES, .Parent.Range("A1:B1"), True, VBooTableType:=BooTable, VThmColor:=EnuThm, VThmFontColor:=EnuThmFont)
    End With
    With RngBos.Range("F7").Validation
      Call .Delete
      Call .Add(Type:=xlValidateList, Formula1:=VBA.UCase("cajero, pagado por"))
      Let .ShowError = False
      Let .Parent.Value = VBA.UCase("cajero")
    End With
    Call VWP.ViewsTableStaticSet(VRES, RngBos.Range("E7:G7"), True, VVarHeaderHorizontalAlignement:=xlHAlignCenter, VBooTableType:=BooTable, VThmColor:=EnuThm, VThmFontColor:=EnuThmFont)
    With RngBos.Range("J7")
      Let .Value = "PROVEDOR"
      Call VWP.ViewsTableStaticSet(VRES, .Range("A1:B1"), True, VBooTableType:=BooTable, VThmColor:=EnuThm, VThmFontColor:=EnuThmFont)
    End With

    ' Autosetup page
    If BooConfigPage Then
      Call Viewprint_core_assets_getheaderandfooter(VRES, VMDL, StrHeaderTop, StrHeaderBtm, StrLogoTop, StrLogoBtm)
      ' Set up the page
      Call VWP.ViewsPageSetup( _
        VRES, _
        RngBox.Parent, _
        "A1", xlPortrait, xlPaperLetter, 2, 1, 2, 2, 0.5, 1, _
        ShHeaderLeft:=StrHeaderTop, _
        ShHeaderRight:=StrLogoTop, _
        ShFooterLeft:=StrHeaderTop, _
        ShFooterRight:=StrLogoBtm, _
        ShImgHeader:=2, _
        ShImgFooter:=2, _
        ShImgFileHeader:=StrLogoTop, _
        ShImgFileFooter:=StrLogoBtm)
    End If

  End With
  Set VWP = Nothing

  ' Return
  Let Viewprint_cncinvoice = True

EH:
  Set RngBos = Nothing
  Set RngBox = Nothing
  Set HPR = Nothing
  Set VWP = Nothing
  Call VRES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function


Private Function Viewprint_help( _
  ByVal VRES As ClassEqlResources, _
  ByVal VMDL As ClassEqlModel, _
  ByVal PVarRequest As Variant, _
  ByVal PStrId As String) As Boolean

  Dim ShtParentRngBox As Worksheet
  Dim RngBox As Range
  Dim ShpCurrent As Shape
  Dim StrAux0 As String
  Dim EnuThm As XlThemeColor, EnuThmFont As XlThemeColor

  Const sFN As String = "VIEWPRINT::-help"

  On Error GoTo EH
  Let Viewprint_help = False

  ' Request: RngBox
  If Not VBA.IsArray(PVarRequest) Then GoTo EH
  Set RngBox = PVarRequest(0)
  If RngBox Is Nothing Then GoTo EH

  Let EnuThm = VMDL.TableThemeColor
  Let EnuThmFont = VMDL.TableThemeFontColor

  ' Clean active sheet
  Set ShtParentRngBox = RngBox.Parent
  Call ShtParentRngBox.Shapes.SelectAll
  Call Selection.Delete

  ' Choice the current module help requester and fill main vector data
  Select Case PStrId

    Case "regression", "balmet", "balmeteco", "projection", "oreinvoice", "concentrateinvoice": Let StrAux0 = "eqlimg" & PStrId
    Case "rlselection", "balmetselection", "udf": Let StrAux0 = VBA.vbNullString
    Case Else: GoTo EH

  End Select

  ' Copy help img ribbon to active sheet
  If Not PStrId = "udf" Then ' Udf's help hasn't images to show
    With excliqlites.Shapes.Item("eqlimgribbonmenu")
      Call .Copy
      Call ShtParentRngBox.Range("A1").PasteSpecial(xlPasteAll)
    End With
    Set ShpCurrent = ShtParentRngBox.Shapes.Item(1)
    With ShpCurrent
      Let .Name = "Img_RibbonMenu"
      Let .LockAspectRatio = msoTrue
      Let .Width = ShtParentRngBox.Range("A1:Q1").Width
    End With
    Set ShpCurrent = Nothing

    ' Copy help frm img to active sheet
    If Not StrAux0 = VBA.vbNullString Then
      With excliqlites.Shapes.Item(StrAux0)
        Call .Copy
        Call ShtParentRngBox.Range("A9").PasteSpecial(xlPasteAll)
      End With

      Set ShpCurrent = ShtParentRngBox.Shapes.Item(2)
      With ShpCurrent
        Let .Name = "Img_Frm" & PStrId
        Let .LockAspectRatio = msoTrue
        Let .Width = ShtParentRngBox.Range("A:G").Width
      End With
      Set ShpCurrent = Nothing
    End If
  End If

  With ShtParentRngBox

    ' Set help ribbon img highlighter
    If Not PStrId = "udf" Then ' Udf's help hasn't images to show
      Select Case PStrId
        Case "regression": Set ShpCurrent = .Shapes.AddShape(msoShapeRectangle, 1.2, 43.2, 43.2, 66)
        Case "rlselection": Set ShpCurrent = .Shapes.AddShape(msoShapeRectangle, 92, 43.2, 91, 20)
        Case "balmet": Set ShpCurrent = .Shapes.AddShape(msoShapeRectangle, 45.5, 43.2, 43.2, 66)
        Case "balmetselection": Set ShpCurrent = .Shapes.AddShape(msoShapeRectangle, 92, 61, 91, 18)
        Case "balmeteco": Set ShpCurrent = .Shapes.AddShape(msoShapeRectangle, 92, 76, 91, 18)
        Case "projection": Set ShpCurrent = .Shapes.AddShape(msoShapeRectangle, 185.5, 43.2, 43.2, 66)
        Case "oreinvoice": Set ShpCurrent = .Shapes.AddShape(msoShapeRectangle, 230.5, 43.2, 43.2, 66)
        Case "concentrateinvoice": Set ShpCurrent = .Shapes.AddShape(msoShapeRectangle, 276, 43.2, 50, 66)
      End Select
      Let ShpCurrent.Name = "ShpHiglighter"
      With ShpCurrent.Fill
        Let .ForeColor.ObjectThemeColor = EnuThm
        Let .Transparency = 0.7
        Call .Solid
      End With
      With ShpCurrent.Line
        Let .ForeColor.ObjectThemeColor = EnuThm
        Let .ForeColor.Brightness = -0.25
        Let .Weight = 1.5
      End With
      Set ShpCurrent = Nothing
    End If

    ' Set help data
    With .Range(IIf(PStrId = "udf", "B2:Q2", IIf(StrAux0 = VBA.vbNullString, "D9:N9", "I9:Q9")))
    'With .Range(IIf(StrAux0 = VBA.vbNullString, "D9:N9", IIf(Not PStrId = "udf", "I9:Q9", "A9:Q9")))
      Set ShpCurrent = ShtParentRngBox.Shapes.AddTextbox(msoTextOrientationHorizontal, .Left, .Top, .Width, .Height)
    End With
    With ShpCurrent
      Let .Name = "ShpContent"
      Let .Fill.ForeColor.ObjectThemeColor = EnuThm
      Let .Fill.Transparency = 0.9
      Let .Line.ForeColor.ObjectThemeColor = EnuThm
      Let .TextFrame.Characters.Text = Viewprint_core_assets_gethelpcontent(VRES, PStrId)
      Let .TextFrame2.AutoSize = msoAutoSizeShapeToFitText
    End With
    Set ShpCurrent = Nothing

    ' Set help close button
    With .Range(IIf(PStrId = "udf", "P2:Q3", IIf(StrAux0 = VBA.vbNullString, "M9:N10", "P9:Q10")))
      Set ShpCurrent = ShtParentRngBox.Shapes.AddFormControl(xlButtonControl, .Left, .Top, .Width, .Height)
    End With
    With ShpCurrent
      Let .TextFrame.Characters.Text = "Cerrar Ayuda"
      Let .OnAction = "'ExcliqliteCloseHelp_click """ & ShtParentRngBox.Parent.Name & """, """ & ShtParentRngBox.Name & """'"
      #If Debugging Then
        Debug.Print "'ExcliqliteCloseHelp_click """ & ShtParentRngBox.Parent.Name & """, """ & ShtParentRngBox.Name & """'"
      #End If
    End With
    Set ShpCurrent = Nothing

    Let .Application.ActiveWindow.DisplayHeadings = False
    Let .Application.ActiveWindow.DisplayGridlines = False
    Call .Range("R1").Select

  End With

  Let Viewprint_help = True

EH:
  Set RngBox = Nothing
  Set ShtParentRngBox = Nothing
  Call VRES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function



' ASSETS ----------------------------------------------------------------------------'
Private Sub Viewprint_core_assets_getheaderandfooter( _
  ByVal VRES As ClassEqlResources, _
  ByVal VMDL As ClassEqlModel, _
  ByRef VStrHeader As String, _
  ByRef VStrFooter As String, _
  ByRef VStrImgHeader As String, _
  ByRef VStrImgFooter As String)

  Dim StrAux0 As String, StrAux1 As String, StrAux2 As String

  Const sFN As String = "VIEWPRINT::-core_assets_getheaderandfooter"

  On Error GoTo EH

  Let VStrHeader = VMDL.CompanyName ' Company
  Let StrAux0 = VBA.Right(VStrHeader, 4) ' [0|1|2]
  Let VStrHeader = VBA.Replace(VStrHeader, StrAux0, VBA.vbNullString)
  Let VStrHeader = IIf(StrAux0 = "[0]" Or StrAux0 = "[2]", VStrHeader & Chr(10), VBA.vbNullString)
  Let VStrFooter = IIf(StrAux0 = "[1]" Or StrAux0 = "[2]", VStrHeader & Chr(10), VBA.vbNullString)

  Let StrAux0 = VMDL.CompanyAddress ' Company Address
  Let StrAux1 = VBA.Right(StrAux0, 4) ' [0|1|2]
  Let StrAux0 = VBA.Replace(StrAux0, StrAux1, VBA.vbNullString)
  Let VStrHeader = VStrHeader & IIf(StrAux1 = "[0]" Or StrAux1 = "[2]", StrAux0 & Chr(10), VBA.vbNullString)
  Let VStrFooter = VStrFooter & IIf(StrAux1 = "[1]" Or StrAux1 = "[2]", StrAux0 & Chr(10), VBA.vbNullString)

  Let StrAux0 = VMDL.CompanyPhone ' Company Phone
  Let StrAux1 = VBA.Right(StrAux0, 4) ' [0|1|2]
  Let StrAux0 = VBA.Replace(StrAux0, StrAux1, VBA.vbNullString)
  Let VStrHeader = VStrHeader & IIf(StrAux1 = "[0]" Or StrAux1 = "[2]", StrAux0, VBA.vbNullString)
  Let VStrFooter = VStrFooter & IIf(StrAux1 = "[1]" Or StrAux1 = "[2]", StrAux0, VBA.vbNullString)

  Let StrAux2 = IIf(StrAux0 = VBA.vbNullString, VBA.vbNullString, " | ")
  Let StrAux0 = VMDL.CompanyEmail ' Company Email
  Let StrAux1 = VBA.Right(StrAux0, 4) ' [0|1|2]
  Let StrAux0 = VBA.Replace(StrAux0, StrAux1, VBA.vbNullString)
  Let VStrHeader = VStrHeader & IIf(StrAux1 = "[0]" Or StrAux1 = "[2]", StrAux2 & StrAux0, VBA.vbNullString)
  Let VStrFooter = VStrFooter & IIf(StrAux1 = "[1]" Or StrAux1 = "[2]", StrAux2 & StrAux0, VBA.vbNullString)

  Let StrAux0 = VMDL.CompanyLogo ' Company Logo
  Let StrAux1 = VBA.Right(StrAux0, 4) ' [0|1|2]
  Let StrAux0 = VBA.Replace(StrAux0, StrAux1, VBA.vbNullString)
  Let VStrImgHeader = IIf(StrAux1 = "[0]" Or StrAux1 = "[2]", StrAux0, VBA.vbNullString)
  Let VStrImgFooter = IIf(StrAux1 = "[1]" Or StrAux1 = "[2]", StrAux0, VBA.vbNullString)

EH:
  Call VRES.ErrorHandler.ErrorHandlerRaise(sFN)

End Sub

Private Function Viewprint_core_assets_gethelpcontent(ByVal VRES As ClassEqlResources, ByVal VStrRequester As String) As String

  Dim StrAux0 As String, STRAPPTITLE As String

  On Error GoTo EH

  Let Viewprint_core_assets_gethelpcontent = VBA.vbNullString

  Let STRAPPTITLE = VRES.AppStrTitle
  Select Case VStrRequester

    ' REGRESSION
    Case "regression"
      Let StrAux0 = _
        "AYUDA RÁPIDA: REGRESIÓN LINEAL" & Chr(13) & Chr(13) & _
        "DESCRIPCIÓN:" & Chr(13) & _
        "DESCRIPCIÓN:" & Chr(13) & _
        "Permite obtener el ajuste lineal para cada punto de una variable dependiente Y' proyectada, todo, a partir de una colección datos dispersos X e Y." & Chr(13) & Chr(13) & _
        "COMPONENTES:" & Chr(13) & _
        "1. X: Conjunto de números separados por punto y coma, Ej. 1;2;3. requiere la misma cantidad de datos que Y. Al ser opcional puede dejar vacío el cuadro de texto." & Chr(13) & _
        "2. Y: Conjunto de números separados por punto y coma, Ej. 1;2;3. Como mínimo se requiere tres datos." & Chr(13) & _
        "3. Extrapolar: Conjunto de números separados por punto y coma, Ej. 1;2;3. requiere como mínimo un número. Por defecto, recibe el valor cero o el número siguiente al último número de la variable Y. Al ser opcional puede dejar vacío el cuadro de texto." & Chr(13) & _
        "4. Opciones varias que el algoritmo incluirá en el resultado." & Chr(13) & _
        "5. Opciones de gráfico: si activa la casilla, el sistema generará el gráfico del ejercicio en la hoja de resultado actual o en una hoja de gráfico." & Chr(13) & _
        "6. Opciones de impresión de resultado: Si selecciona Hoja nueva, el resultado se mostrará en una nueva hoja de excel, si selecciona Hoja actual, el resultado se mostrará en la celda de la hoja que seleccione mediante el botón del cuadro de texto." & Chr(13) & _
        "7. Resolver: tras pulsar el botón, el sistema resolverá la regresión lineal, mostrará el resultado y cerrará el formulario."

    ' REGRESSION SELECTION
    Case "rlselection"
      Let StrAux0 = _
        "AYUDA RÁPIDA: REGRESIÓN LINEAL SELECCIÓN" & Chr(13) & Chr(13) & _
        "DESCRIPCIÓN:" & Chr(13) & _
        "Permite obtener el ajuste lineal para cada punto de una variable dependiente Y' proyectada, todo, a partir de una colección datos dispersos X e Y correspondientes a una, dos o tres columnas de una hoja Excel." & Chr(13) & Chr(13) & _
        "COMPONENTES:" & Chr(13) & _
        "1. X: Un rango de tres o más datos numéricos seleccionados en una hoja Excel. Requiere la misma cantidad de datos que Y. Al ser opcional puede obviar esta selección." & Chr(13) & _
        "2. Y: Un rango de tres o más datos numéricos seleccionados en una hoja Excel." & Chr(13) & _
        "3. Extrapolar: Un rango de uno o más datos numéricos seleccionados en una hoja Excel. Por defecto, recibe el valor cero o el número siguiente al último número de la variable Y. Al ser opcional puede dejar obviar esta selección." & Chr(13) & _
        "4. Opciones varias: Si la opción Resultados compactos en selecciones del formulario de Configuración está marcado, el resultado final mostrará sólo la información básica de la regresión lineal." & Chr(13) & _
        "5. Opciones de impresión de resultado: Si la opción Pedir celda RL y Balmet del grupo Configurción en la cinta de opciones de Excel está marcada, el sistema pedirá que seleccione una celda en una hoja del libro que desee donde imprimirá el resultado; de otro modo, será creada una nueva hoja en el libro actual y el resultado será impreso en la primera celda."

    ' BALMET
    Case "balmet"
      Let StrAux0 = _
        "AYUDA RÁPIDA: BALANCE METALÚRGICO" & Chr(13) & Chr(13) & _
        "DESCRIPCIÓN:" & Chr(13) & _
        "Permite obtener el balance de un proceso metalúrgico de hasta diez productos mediante uno de tres posibles métodos y para productos directos e indirectos (minerales complejos)." & Chr(13) & Chr(13) & _
        "COMPONENTES:" & Chr(13) & _
        "1. Método: Si el proceso tiene más de dos productos, debe seleccionar CRAMER o Matriz inversa. (Por defecto CRAMER)." & Chr(13) & _
        "2. Peso Alimentación: un número o varios separados por punto y coma. Si introduce más de un número y el proceso contiene leyes volumétricas (g/L), tomará el primer valor como peso seco en toneladas, los demás valores corresponderán al volumen de alimentación." & Chr(13) & _
        "3. Leyes de minerales y productos: Conjunto de números separados por punto y coma, Ej. 1;2;3. Para un balance de dos productos se requieren tres leyes (tres filas: producto valioso o Concentrado, Relave y Alimentación, una columna: 3x1), para un balance de tres productos son necesarias ocho leyes (4x2: dos concentrados, relave y alimentación), para cuatro productos quince leyes (5x3), etc. El sistema resuelve balances de hasta diez productos." & Chr(13) & _
        "4. Seleccionar minerales: Opcional, por defecto, el sistema trabaja con minerales genéricos, pulse la casilla si desea trabajar con minerales específicos y asígnelos mediante la lista desplegable." & Chr(13) & _
        "5. Seleccionar unidades: Opcional, por defecto, el sistema trabaja con unidades de porcentaje (%), pulse la casilla si desea trabajar con unidades específicas y asígnelas mediante la lista desplegable." & Chr(13) & _
        "6. Complejos de metales o minerales preciosos: Opcional, si el proceso obtiene productos por consecuencia (como en un proceso por lixiviación o quizá en un proceso por flotación), puede calcular sus resultados asignando en los controles los valores de leyes obtenidos; en este caso, cada conjunto de leyes debe corresponder con el número de filas del proceso principal, es decir, para dos productos: tres filas de leyes (3x1), para tres productos: cuatro filas (4x1), para cuatro productos: (5x1), etc. El sistema soporta hasta tres productos consecuencia, esto es: dos productos 3x3, tres productos: 4x3, cuatro productos: 5x3, etc." & Chr(13) & _
        "7. Opciones varias que el algoritmo incluirá en el resultado." & Chr(13) & _
        "8. Opciones de impresión de resultado: Si selecciona Hoja nueva, el resultado se mostrará en una nueva hoja de excel, si selecciona Hoja actual, el resultado se mostrará en la celda de la hoja que seleccione mediante el botón del cuadro de texto." & Chr(13) & _
        "9. Resolver: tras pulsar el botón, el sistema resolverá el balance metalúrgico, mostrará el resultado y cerrará el formulario."

    ' BALMETSELECTION
    Case "balmetselection"
      Let StrAux0 = _
        "AYUDA RÁPIDA: BALANCE METALÚRGICO SELECCIÓN" & Chr(13) & Chr(13) & _
        "DESCRIPCIÓN:" & Chr(13) & _
        "Permite obtener el balance de un proceso metalúrgico de hasta diez productos mediante el método de CRAMER y para productos directos e indirectos (minerales complejos). Requiere la selección de uno o dos rangos de datos." & Chr(13) & Chr(13) & _
        "COMPONENTES:" & Chr(13) & _
        "1. Peso Alimentación: un número o la suma de varios, siempre será el primer número de la selección. Para balances volumétricos se toma los siguientes números del rango y si todo el proeso es volumétrico, la alimentación es la suma del primer rango." & Chr(13) & _
        "2. Leyes de minerales y productos: Rango de números. Para un balance de dos productos se requieren tres leyes (tres filas: producto valioso o Concentrado, Relave y Alimentación, una columna: 3x1), para un balance de tres productos son necesarias ocho leyes (4x2: dos concentrados, relave y alimentación), para cuatro productos quince leyes (5x3), etc. El sistema resuelve balances de hasta diez productos. Puede introducir en este rango las unidades y nombres de los productos del balance, Ej. Zn, %, de otro modo se aplicarán nombres genéricos y % como unidad." & Chr(13) & _
        "3. Complejos de metales o minerales preciosos: Opcional, si el proceso obtiene productos por consecuencia (como en un proceso por lixiviación o quizá en un proceso por flotación), puede calcular sus resultados seleccionando un terer rango con los valores de leyes obtenidos; en este caso, cada conjunto de leyes debe corresponder con el número de filas del proceso principal, es decir, para dos productos: tres filas de leyes (3x1), para tres productos: cuatro filas (4x1), para cuatro productos: (5x1), etc. El sistema soporta hasta tres productos consecuencia, esto es: dos productos 3x3, tres productos: 4x3, cuatro productos: 5x3, etc. Puede introducir en este rango las unidades y nombres de los productos complejos del balance, Ej. Ag, DM" & Chr(13) & _
        "4. Opciones varias: Si la opción Resultados compactos en selecciones del formulario de Configuración está marcado, el resultado final mostrará sólo la información básica del balance metalúrgico." & Chr(13) & _
        "5. Opciones de impresión de resultado: Si la opción Pedir celda RL y Balmet del grupo Configurción en la cinta de opciones de Excel está marcada, el sistema pedirá que seleccione una celda en una hoja del libro que desee donde imprimirá el resultado; de otro modo, será creada una nueva hoja en el libro actual y el resultado será impreso en la primera celda."

    ' BALMETECO
    Case "balmeteco"
      Let StrAux0 = _
        "AYUDA RÁPIDA: BALANCE METALÚRGICO (ANÁLISIS ECONÓMICO)" & Chr(13) & Chr(13) & _
        "DESCRIPCIÓN:" & Chr(13) & _
        "Permite obtener el balance de un proceso metalúrgico de hasta diez productos mediante uno de tres posibles métodos y para productos directos e indirectos (minerales complejos). El objetivo es calcular pesos o leyes de productos y alimentación que permitan un análisis económico de la operación." & Chr(13) & Chr(13) & _
        "COMPONENTES:" & Chr(13) & _
        "1. Objetivo: Ley de cabeza, que se obtiene a partir de leyes y pesos de productos. Ley de productos, que se obtienen a partir de leyes y pesos de cabeza." & Chr(13) & _
        "2. Método: Si el proceso tiene más de dos productos, debe seleccionar CRAMER o Matriz inversa. (Por defecto CRAMER)." & Chr(13) & _
        "3. Peso Alimentación: un número o varios separados por punto y coma. Si introduce más de un número y el proceso contiene leyes volumétricas (g/L), tomará el primer valor como peso seco en toneladas, los demás valores corresponderán al volumen de alimentación." & Chr(13) & _
        "4. Peso Producto: un número o varios separados por punto y coma, cada uno representa un producto del proceso." & Chr(13) & _
        "5. Leyes de cabeza/productos: Conjunto de números separados por punto y coma, Ej. 1;2;3. cada ley se corresponde con cada peso de producto." & Chr(13) & _
        "6. Recuperación: Opcional. Se trata de la calidad de producto que se desea obtener en el proceso. Cada recuperación debe corresponder a cada peso de producto. Por defecto, el sistema trabaja con 80%." & Chr(13) & _
        "7. Complejos de metales o minerales preciosos: Opcional, si el proceso obtiene productos por consecuencia (como en un proceso por lixiviación o quizá en un proceso por flotación), puede calcular sus resultados asignando en los controles los valores de leyes obtenidos; en este caso, cada conjunto de leyes debe corresponder con el número de pesos de productos del proceso principal, tal cual la recuperación." & Chr(13) & _
        "8. Opciones varias que el algoritmo incluirá en el resultado." & Chr(13) & _
        "9. Opciones de impresión de resultado: Si selecciona Hoja nueva, el resultado se mostrará en una nueva hoja de excel, si selecciona Hoja actual, el resultado se mostrará en la celda de la hoja que seleccione mediante el botón del cuadro de texto." & Chr(13) & _
        "10. Resolver: tras pulsar el botón, el sistema resolverá el balance metalúrgico, mostrará el resultado y cerrará el formulario."

    ' PROJECTS PROYECTIONS
    Case "projection"
      Let StrAux0 = _
        "AYUDA RÁPIDA: PROYECCIÓN PROYECTOS" & Chr(13) & Chr(13) & _
        "DESCRIPCIÓN:" & Chr(13) & _
        "Permite obtener la proyección de un proyecto empresarial minero-metalúrgico en N días, semanas, meses o años. Este módulo utiliza el valor residual del proyecto como constante. El objetivo es conocer la viabilidad del proyecto mediante los criterios VAN y TIR." & Chr(13) & Chr(13) & _
        "COMPONENTES:" & Chr(13) & _
        "1. Inversión inicial: un número, siempre diferente de cero. SE trata del monto de inversión inicial que requiere el proyecto." & Chr(13) & _
        "2. Peso fino mineral: Corresponde al peso en toneladas del mineral o metal rico resultado de cualquier proceso metalúrgico. Puede introducir hasta diez pesos de diferentes productos a la vez, números separados por punto y coma, Ej.: 1;2;3." & Chr(13) & _
        "3. Cotización mineral (o metal): Opcional, corresponde a la cotización de cada producto, cada cotización debe ser un grupo mínimo de tres números, puede seleccionarlos desde una hoja Excel o introducirlos en el cuadro de texto separando cada número con punto y coma y cada grupo con una barra diagonal, Ej.: 1;1;1/2;1;2/3;3;1" & Chr(13) & _
        "4. Valor residual: un número aproximado que corresponde al las depreciaciones que se obtendrían de activos y similares durante la vida del proyecto." & Chr(13) & _
        "5. Descuento en porcentaje sobre ingreso bruto: un valor aproximado al conjunto de pagos por conceptos impositivos, egresos, sueldos y salarios, etc. aplicables durante la vida del proyecto. Valor en porcentaje." & Chr(13) & _
        "6. Parámetros de proyección: el tiempo estimado de vida del proyecto y la tasa de ganancia mínima requerida estimada. Permitirán conocer la capacidad y el tiempo en que el proyecto recupera la inversión y se convierte en autosustentable." & Chr(13) & _
        "7. Opciones varias: Mostrará o no, según las opciones estén marcadas: el Valor Actual Neto, la Tasa interna de Retorno y el gráfico del proyecto en una hoja nueva de gráfico o en la misma hoja de resultados." & Chr(13) & _
        "8. Opciones de impresión de resultado: Si selecciona Hoja nueva, el resultado se mostrará en una nueva hoja de excel, si selecciona Hoja actual, el resultado se mostrará en la celda de la hoja que seleccione mediante el botón del cuadro de texto." & Chr(13) & _
        "9. Resolver: tras pulsar el botón, el sistema resolverá el ejercicio de proyección, mostrará el resultado y cerrará el formulario."

    ' MINERAL PURCHASE
    Case "oreinvoice"
      Let StrAux0 = _
        "AYUDA RÁPIDA: LIQUIDACIÓN DE MINERALES" & Chr(13) & Chr(13) & _
        "DESCRIPCIÓN:" & Chr(13) & _
        "Permite obtener hasta diez liquidaciones compra-venta de minerales basado en la comercialización de minerales boliviana (adaptable a comercialización de otros países)." & Chr(13) & Chr(13) & _
        "COMPONENTES:" & Chr(13) & _
        "1. Datos base: datos iniciales de la liquidación." & Chr(13) & _
        "2. Datos provedor: si la casilla Varios provedores está seleccionada, podrá introducir una lista de hasta diez provedores de tipo particular o cooperativista, esta opción permitirá generar una liquidación por provedor. Si se encuentra desactivada y tiene varios datos de liquidaciones en la lista, será generada una liquidación que incluya todos los registros que figuran en la lista correspondientes a un solo provedor." & Chr(13) & _
        "3. Datos liquidación: Cada liquidación requiere datos del mineral (fecha de internación a planta, peso, número de muestra, etc.), datos de retenciones impositivas que gravan a la transacción, así como otros descuentos y bonos acordados entre comprador y provedor; toda esa información debe ser introducida en este apartado. Si la casilla Promedia valores está seleccionada y varias liquidaciones corresponden a un solo provedor, todos los registros de todas las cargas minerles de la liquidación se resumirán en uno solo en el documento final. A su vez, la casilla Autoconfigurar página dará formato de impresión final a la o las páginas generadas, opcionalmente, cada página puede tener una copia incluida, esto ocurrirá si la casilla Incluir copia está activada; el resultado irá en una hoja en disposición horizontal." & Chr(13) & _
        "4. Dónde mostrar resultado: si selecciona Hoja nueva, el resultado se mostrará en una nueva hoja de excel, si selecciona Hoja actual, el resultado se mostrará en la celda de la hoja que seleccione mediante el botón del cuadro de texto." & Chr(13) & _
        "5. Generar: tras pulsar el botón, el sistema generará el o los documentos finales de liquidación de minerales, mostrará el resultado y cerrará el formulario."

    ' CONCENTRATE PURCHASE
    Case "concentrateinvoice"
      Let StrAux0 = _
        "AYUDA RÁPIDA: LIQUIDACIÓN DE CONCENTRADOS MINERALES" & Chr(13) & Chr(13) & _
        "DESCRIPCIÓN:" & Chr(13) & _
        "Permite obtener una liquidación de compra-venta de concentrados minerales basado en la comercialización de minerales boliviana (adaptable a comercialización de otros países)." & Chr(13) & Chr(13) & _
        "COMPONENTES:" & Chr(13) & _
        "1. Datos base y de provedor: datos iniciales de la liquidación. Todo el cálculo es realizado en moneda extranjera asumiendo que ésta es el dólar americano (USD), si la casilla Imprimir moneda local está seleccionada, se generará una columna con el valor de moneda local en la liquidación final" & Chr(13) & _
        "2. Datos concentrado y costo tratamiento: datos fijados en un documento denominado contrato de compra-venta de concentrados minerales entre el comprador y provedor." & Chr(13) & _
        "3. Datos liquidación: una liquidación de cocentrados minerales requiere datos del mineral (contenidos, peso, etc.), datos de retenciones impositivas que gravan a la transacción, así como otros descuentos y bonos acordados entre comprador y provedor; toda esa información debe ser introducida en este apartado. Si la casilla Autoconfigurar página está seleccionada, se dará formato de impresión final a la página generada." & Chr(13) & _
        "4. Dónde mostrar resultado: si selecciona Hoja nueva, el resultado se mostrará en una nueva hoja de excel, si selecciona Hoja actual, el resultado se mostrará en la celda de la hoja que seleccione mediante el botón del cuadro de texto." & Chr(13) & _
        "5. Generar: tras pulsar el botón, el sistema generará el documento final de liquidación de concentrados minerales, mostrará el resultado y cerrará el formulario."

    ' HELP UDF's
    Case "udf"
      Let StrAux0 = Chr(13) & _
        "AYUDA RÁPIDA: FUNCIONES DE EXCLIQ LITE EN EXCEL" & Chr(13) & Chr(13) & _
        STRAPPTITLE & " contiene veinticinco (25) funciones similares a las funciones nativas de " & Application.Name & " como =SUMA(num1; num2; ...). Estas funciones corresponden a los tres primeros módulos del sistema: Regresión lineal, Balance metalúrgico y Proyección de proyectos; y dos complementarios: Versión y Número a texto." & Chr(13) & _
        "Su aplicación permite obtener resultados rápidos de cada ejercicio directamente en la celda desde donde son invocados. Para llamar a cualquier función, deber escribir al principio el signo igual (=) seguido del nombre de la función. Todas las funciones de " & STRAPPTITLE & " comienzan con el prefijo 'EQL_', lo que le permite encontrar rápidamente la función que desea. Los datos pueden ser referencias a otras celdas de la hoja actual del libro actual o de una hoja correspondiente a cualquier otro libro de " & Application.Name & " que se encuentre abierto; o puede introducir también una lista de números separados por punto y coma y encerrados entre comillas, Ej. '1;2;3'." & Chr(13) & Chr(13) & _
        "NOTA: Como las aplicaciones creadas en el lenguaje VBA de " & Application.Name & " carecen de un método que permita mostrar ayudas flotantes tras llamar a una función desde una celda, tal como sí ocurre con las funciones nativas de " & Application.Name & ", se recomienda que al trabajar con las funciones de " & STRAPPTITLE & ", las llame desde la barra de fórmulas que despliega el formulario de fórmulas donde tendrá una guía de utilización de cada fórmula y sus parámetros; o presione la combinación de teclas Ctrl + Shift + A, que le mostrará una guía flotante de los parámetros de la función que utilice." & Chr(13) & Chr(13)
      Let StrAux0 = StrAux0 & _
        "COMPONENTES - FÓRMULA PARA VERSIÓN DEL SISTEMA " & VBA.UCase(STRAPPTITLE) & ":" & Chr(13) & _
        "1. EQL_VERSION: Mostrará la versión del complemento de " & VBA.UCase(STRAPPTITLE) & " que esté utilizando en " & Application.Name & ". No tiene parámetros." & Chr(13) & Chr(13)
      Let StrAux0 = StrAux0 & _
        "COMPONENTES - FÓRMULAS PARA REGRESIÓN LINEAL:" & Chr(13) & _
        "1. EQL_REGRESION_LINEAL(Y, Opcional_X, Opcional_Extrapolacion = 0): Permite obtener la regresión lineal de un grupo mayor a tres números. Requiere del parámetro obligatorio Y que representa a la nube de puntos en el plano cartesiano. El parámetro Opcional_X representa las imágenes de los puntos Y, deberá tener la misma cantidad de números que Y; si omite este parámetro, el sistema generará números correlativos a partir de 1. El parámetro Opcional_Extrapolacion representa a los números a extrapolar, es independiente de X e Y, si lo omite, el sistema generará un número correlativo al último valor de Y para hallar la extrapolación de la Regresión lineal." & Chr(13) & _
        "2. EQL_REGRESION_LINEAL_a(Y, Opcional_X): Halla el coeficiente 'a' de la ecuación lineal Y' = a + bX' de la Regresión lineal; el parámetro Y es obligatorio." & Chr(13) & _
        "3. EQL_REGRESION_LINEAL_b(Y, Opcional_X): Halla el coeficiente 'b' de la ecuación lineal Y' = a + bX' de la Regresión lineal; el parámetro Y es obligatorio." & Chr(13) & _
        "4. EQL_REGRESION_LINEAL_r(Y, Opcional_X): Halla el coeficiente 'r' de correlación de la Regresión lineal; el parámetro Y es obligatorio." & Chr(13) & _
        "5. EQL_REGRESION_LINEAL_r2(Y, Opcional_X): Halla el coeficiente 'r^2' de determinación de la Regresión lineal; el parámetro Y es obligatorio." & Chr(13) & _
        "6. EQL_REGRESION_LINEAL_Se(Y, Opcional_X): Halla el error estándar 'S(e)' de determinación de la Regresión lineal; el parámetro Y es obligatorio." & Chr(13) & _
        "7. EQL_REGRESION_LINEAL_n(Y, Opcional_X): Permite obtener la cantidad de datos involucrados en la Regresión lineal; el parámetro Y es obligatorio." & Chr(13) & Chr(13)
      Let StrAux0 = StrAux0 & _
        "COMPONENTES - FÓRMULAS PARA BALANCE METALÚRGICO:" & Chr(13) & _
        "1. EQL_BALMET(Alimentacion, LeyesConcentrado, NombresElementos_Opcional, UnidadesLey_Opcional, Metodo_Opcional, ResultadoHorizontal_Opcional, IncluirAlimentacion_Opcional): Todas las funciones de balance metalúrgico retornan datos matriciales como respuesta, por lo que deberá primero seleccionar un número  de celdas correspondiente a los datos esperados, introducir la fórmula y pulsar la combinación de teclas Ctrl + Shift + Enter para ejecutar la función. Permite obtener los pesos de los productos de un proceso metalúrgico a partir de los parámetros obligatorios: peso de Alimentación, que debe ser un número y las leyes de dichos productos que deben ser un grupo de números que formen una matriz según la cantidad de productos del proceso, es decir, para balances metalúrgicos de dos productos se requieren tres datos que forman tres filas y una columna (3x1); para tres productos cuatro filas y dos columnas (4x2), etc., puede resolver balances metalúrgicos de hasta diez productos." & _
        "Los parámetros opcionales: NombresElementos_Opcional y UnidadesLey_Opcional permiten hacer conocer al sistema a qué elementos o especies minerales corresponden las leyes de los productos, por defecto, el sistema trabaja con elementos genéricos (B1, B2, etc.); la lista de elementos que tiene incorpodada el sistema puede ser encontrada en la ficha Elementos/Especies minerales del formulario de configuración en la cinta de opciones de " & Application.Name & "; las unidades que puede introducir son: porcentaje (%), decimarcos (DM), gramos por litro o gramos por tonelada (g/T ó g/L), el sistema trabaja por defecto con porcentajes (%). El parámetro opcional Metodo_Opcional, indica al sistema qué metodo aplicar, existen tres: 0: Convencional (sólo para balances metalúrgicos de dos productos), 1: CRAMER o determinantes y 2: Matriz inversa." & _
        "El parámetro opcional ResultadoHorizontal_Opcional recibe uno de dos posibles valores, 0 o 1, (ó verdadero o falso), cuando es 0, la matriz respuesta que contiene los pesos del balance metalúrgico se imprimirán en columnas; en el valor 1 tenrán disposición de filas. Finalmente, la fórmula retorna por defecto sólo los pesos de los productos del proceso, pero es posible pedirle que entre esos pesos también retorne el peso de alimentación mediante el parámetro opcional IncluirAlimentacion_Opcional cuyos posibles valores deben ser 0 o 1." & Chr(13) & _
        "2. EQL_BALMET_PORCENTAJE_EN_PESO: Permite encontrar los porcentajes en peso del proceso correspondientes a los pesos de los productos." & Chr(13) & _
        "3. EQL_BALMET_VOLUMEN: Permite encontrar encontrar los volumenes de los productos en caso de que el medio del proceso tenga volúmenes como base." & Chr(13) & _
        "4. EQL_BALMET_PORCENTAJE_EN_VOLUMEN: Permite encontrar los porcentajes en volumen del proceso correspondientes a los volumenes de los productos." & Chr(13) & _
        "5. EQL_BALMET_CABEZAS_ENSAYADA_CALCULADA: Permite encontrar un conjunto de valores correspondientes a las leyes de cabeza ensayada y calculada del proceso." & Chr(13) & _
        "6. EQL_BALMET_UNIDADES: Permite encontrar la columna de datos del balance metalúrgico denominada Unidades que expresa el tanto por uno de los pesos de los productos." & Chr(13) & _
        "7. EQL_BALMET_FINOS: Permite encontrar la columna de datos del balance metalúrgico denominada Finos que expresa la cantidad de material fino rico obtenido en los pesos de los productos." & Chr(13) & _
        "8. EQL_BALMET_RECUPERACION: Permite encontrar la columna de datos del balance metalúrgico denominada Recuperación que expresa la calidad del material obtenido en el proceso." & Chr(13) & _
        "9. EQL_BALMET_RATIO: Permite encontrar la columna de datos del balance metalúrgico denominada Radio de recuperación que expresa la cantidad alimentación requerida por cada tonelada de producto obtenido en el proceso." & Chr(13) & _
        "10. EQL_BALMET_CABEZAS_ENSAYADA_CALCULADA_COMPLEJO: Permite encontrar un conjunto de valores correspondientes a las leyes de cabeza ensayada y calculada de productos consecuencia (complejos) del proceso." & Chr(13) & _
        "11. EQL_BALMET_UNIDADES_COMPLEJO: Permite encontrar la columna de datos del balance metalúrgico denominada Unidades para productos consecuencia (complejos) que expresa el tanto por uno de los pesos de los productos." & Chr(13) & _
        "12. EQL_BALMET_FINOS_COMPLEJO: Permite encontrar la columna de datos del balance metalúrgico denominada Finos para productos consecuencia (complejos) que expresa la cantidad de material fino rico obtenido en los pesos de los productos." & Chr(13) & _
        "13. EQL_BALMET_RECUPERACION_COMPLEJO: Permite encontrar la columna de datos del balance metalúrgico denominada Recuperación para productos consecuencia (complejos) que expresa la calidad del material obtenido en el proceso." & Chr(13) & _
        "14. EQL_BALMET_ECONOMICO_LEY_CABEZA: Con esta función obtendrá el valor aproximado de las leyes de cabeza requeridas para un proceso con leyes de productos y pesos que inicialmente dados." & Chr(13) & _
        "15. EQL_BALMET_ECONOMICO_LEY_PRODUCTOS: Con esta función obtendrá el valor aproximado de las leyes de productos requeridas para un proceso con leyes de productos y pesos que inicialmente dados." & Chr(13) & Chr(13)
      Let StrAux0 = StrAux0 & _
        "COMPONENTES - FÓRMULAS PARA PROYECCIÓN DE PROYECTOS MINERO-METALÚRGICOS:" & Chr(13) & _
        "1. EQL_PROYECCION_PROYECTO(FlujoNetoDeCaja): A partir del flujo neto de caja que consiste en un grupo de valores numéricos, esta función responderá si el proyecto es viable, en ceros o inviable." & Chr(13) & Chr(13)
      Let StrAux0 = StrAux0 & _
        "COMPONENTES - FÓRMULA PARA CONVERTIR UN NÚMERO EN SU EQUIVALENTE LITERAL:" & Chr(13) & _
        "1. EQL_NUMERO_A_TEXTO: Tras introducir un número en el primer parámetro de la función y opcionalmente los nombres en singular y plural de su moneda en sus parámetros siguientes, esta función retornará como respuesta el mismo número pero en literal, Ej. el número 1 trendrpa como respuesta: 'UNO'."

    Case Else: GoTo EH
  End Select

  Let Viewprint_core_assets_gethelpcontent = StrAux0

EH:
  Call VRES.ErrorHandler.ErrorHandlerRaise("VIEWPRINT::-core_assets_gethelpcontent")

End Function






' ERRORS MANNAGEMENT AND ASSERTSIONS ------------------------------------------------'
Private Sub Debug_assert(ByVal Cond As Boolean, Optional Origin As String = VBA.vbNullString)

  #If Debbuging Then
    Debug.Assert Cond
    If Not Cond Then Debug.Print "Assert on: " & Origin
  #End If

End Sub
' CORE METHODS LIST (PRIVATE) ------------------------------------------------------ '



' CORE METHODS LIST (PUBLIC) ------------------------------------------------------- '
Public Function ViewPrint( _
  ByVal VRES As ClassEqlResources, _
  ByVal VMDL As ClassEqlModel, _
  ByVal VStrRequester As String, _
  Optional VVarRequest As Variant, _
  Optional VVarResponse As Variant) As Boolean

  On Error GoTo EH
  Call VBA.MsgBox("Will print a " & VStrRequester & "!")
  Select Case VStrRequester
    Case "regression", "regressionselection": Let ViewPrint = Viewprint_regression(VRES, VMDL, VVarRequest, VVarResponse)
    Case "balmet", "balmetselection", "balmeteco": Let ViewPrint = Viewprint_balmet(VRES, VMDL, VVarRequest, VVarResponse)
    Case "projection": Let ViewPrint = Viewprint_projection(VRES, VMDL, VVarRequest, VVarResponse)

    Case "oreinvoice": Let ViewPrint = Viewprint_oreinvoice(VRES, VMDL, VVarRequest, VVarResponse)
    Case "concentrateinvoice": Let ViewPrint = Viewprint_cncinvoice(VRES, VMDL, VVarRequest, VVarResponse)

    Case "hlpregression", "hlprlselection", "hlpbalmet", "hlpbalmetselection", "hlpbalmeteco", "hlpprojection", "hlporeinvoice", "hlpconcentrateinvoice", "hlpudf"
      Let ViewPrint = Viewprint_help(VRES, VMDL, VVarRequest, VBA.Mid(VStrRequester, 4))
    Case Else: GoTo EH
  End Select

EH:
  Call VRES.ErrorHandler.ErrorHandlerRaise("VIEWPRINT::ViewPrint")

End Function
' CORE METHODS LIST (PUBLIC) ------------------------------------------------------- '
' ===================== CLASS EXCLIQ LITE VIEW PRINT STRUCTURE ===================== '

