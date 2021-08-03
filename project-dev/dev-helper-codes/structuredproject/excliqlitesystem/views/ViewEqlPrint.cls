VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ViewPrint_current"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'============================ VIEWPRINT CLASS STRUCTURE ============================='
' OBJECT VARIABLES (GLOOBJ_) --------------------------------------------------------'
Private EHGLOBAL As AppErrorHandler

' CONSTANTS ERROR MESSAGES (GLOSTR_ERR_) --------------------------------------------'
Private Const CUSTOM_ERROR As Long = VBA.vbObjectError + 517





' PROPERTIES ------------------------------------------------------------------------'
Public Property Set ErrorHandler(ByRef ObjEH As AppErrorHandler)
  Set EHGLOBAL = ObjEH
End Property





' CONSTRUCTOR/DESTRUCTOR ------------------------------------------------------------'
Private Sub Class_Terminate()
  
  Set EHGLOBAL = Nothing

End Sub





' METHODS PRIVATE -------------------------------------------------------------------'
' CORE CONTROLS METHODS LIST (PRIVATE) ----------------------------------------------'
Private Function Viewprint_regression(ByVal BVarReq As Variant, ByVal BVarRes As Variant) As Boolean
  
  Dim VWP As ViewsParent_current

  Dim RngBox As Range
  Dim ShtParentRngBox As Worksheet
  Dim LO As ListObject
  Dim RngAux As Range
  Dim VarGraphData As Variant
  
  Dim StrTableName As String, StrAux1 As String, StrAux2 As String
  Dim BooGraph As Boolean, BooGraphSheet As Boolean, BooSrc As Boolean, BooCor As Boolean, BooReg As Boolean, BooJustRegression As Boolean
  Dim m As Long, n As Long, i As Long, MG As Long

  Const sFN As String = "VIEWPRINT::-regression_print"
  
  On Error GoTo EH
  Let Viewprint_regression = False
  
  'Unpack data
  Let BooSrc = VBA.CBool(BVarRes(0))
  Let BooCor = VBA.CBool(BVarRes(2))
  Let BooReg = VBA.CBool(BVarReq(4))
  Set RngBox = BVarReq(16)
  
  Set VWP = New ViewsParent_current
  Set VWP.ErrorHandler = EHGLOBAL
  
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
      
      If LO Is Nothing Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "Se requiere un objeto ListObject")
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
        Call VWP.ViewsTableStylesLOS(ShtParentRngBox, LO.Name, False, bHasTitle:=True)
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
  
    If LO Is Nothing Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "Se requiere un objeto ListObject")
    With LO
      Let .Name = StrTableName
      Let .DataBodyRange.NumberFormat = "#0.00"
      Let .HeaderRowRange.HorizontalAlignment = xlRight
      Call VWP.ViewsTableStylesLOS(ShtParentRngBox, LO.Name, True, bHasTitle:=True)
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
    
      If LO Is Nothing Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "Se requiere un objeto ListObject")
      With LO
        Let .Name = StrTableName
        Let .ListColumns.Item(2).DataBodyRange.NumberFormat = "#0.00"
        For Each RngAux In .ListColumns.Item(1).DataBodyRange
          If RngAux.Value = "Coeficiente de determinación (r2):" Then Let RngAux.Characters(Start:=32, length:=1).Font.Superscript = True
          If RngAux.Value = "Error estándar de la determinación (Se):" Then Let RngAux.Characters(Start:=38, length:=1).Font.Subscript = True
        Next RngAux
        Let .ListColumns.Item(1).DataBodyRange.Font.Italic = True
        Set RngAux = Nothing
        Call VWP.ViewsTableStylesLOS(ShtParentRngBox, .Name, True, bHasTitle:=True, bColsAutofit:=True)
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
    'RngBox.Range("B3").Resize(m, 2), 'data
    Call VWP.ViewsSetGraphic( _
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
    'Let Application.CutCopyMode = False
  End If

  Let Viewprint_regression = True

EH:
  Set RngAux = Nothing
  Set RngBox = Nothing
  Set LO = Nothing
  Set VWP = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function






' CORE VIEW ASSETS ------------------------------------------------------------------'





' ERRORS MANNAGEMENT AND ASSERTSIONS ------------------------------------------------'
Private Sub Debug_assert(ByVal Cond As Boolean, Optional Origin As String = VBA.vbNullString)
  
  #If Debbuging Then
    Debug.Assert Cond
    If Not Cond Then Debug.Print "Assert on: " & Origin
  #End If

End Sub
' METHODS PRIVATE -------------------------------------------------------------------'





' METHODS PUBLIC --------------------------------------------------------------------'
' PUBLIC METHODS LIST (INTERFACE) ---------------------------------------------------'
Public Function ViewPrint(ByVal VStrRequester As String, Optional VVarRequest As Variant, Optional VVarResponse As Variant) As Boolean
  
  On Error GoTo EH
  Select Case VStrRequester
    Case "regression", "regressionselection": Let ViewPrint = Viewprint_regression(VVarRequest, VVarResponse)
    
    Case "balmet"
    Case "balmeteco"
    Case "balmetselection"
    
    Case "projection"
    Case "oreinvoice"
    Case "concentrateinvoice"
    
    Case Else: GoTo EH
  End Select

EH:
  Call EHGLOBAL.ErrorHandlerRaise("VIEWPRINT::ViewPrint")

End Function
' METHODS PUBLIC --------------------------------------------------------------------'
'============================ VIEWPRINT CLASS STRUCTURE ============================='

