VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassEqlViewParent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ===================== CLASS EXCLIQ LITE VIEW PARENT STRUCTURE ==================== '
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


' ===================== CLASS EXCLIQ LITE VIEW PARENT STRUCTURE ==================== '
' ERRORS MANNAGEMENT AND ASSERTSIONS ----------------------------------------------- '
' Debug_asset
' CORE METHODS LIST (PUBLIC) ------------------------------------------------------- '
' ViewsCaptureRange
' ViewFormCleanner
' ViewFormAddControl
' ViewFormStyles
' ViewsTableStaticSet
' ViewsTableStylesLOS
' viewsSetGraphic
' ViewsPageSetup
' ===================== CLASS EXCLIQ LITE VIEW PARENT STRUCTURE ==================== '



' ===================== CLASS EXCLIQ LITE VIEW PARENT STRUCTURE ==================== '
' ERRORS MANNAGEMENT AND ASSERTSIONS ----------------------------------------------- '
' Debug_assert
Private Sub Debug_assert(ByVal Cond As Boolean, Optional Origin As String = VBA.vbNullString)

  #If Debbuging Then
    Debug.Assert Cond
    If Not Cond Then Debug.Print "Assert on: " & Origin
  #End If

End Sub



' CORE METHODS LIST (PUBLIC) ------------------------------------------------------- '
'ooooooooooooooooooooooo---  FORMS AND RANGES MANNAGEMENT  ---ooooooooooooooooooooooo'
Public Function ViewsCaptureRange( _
  ByVal VRES As ClassEqlResources, _
  ByVal FRM As FRMBasics, _
  Optional sPrompt As String = "Seleccione un rango", _
  Optional sTitle As String = "Excel") As Range

  Dim RngBox As Range
  Dim i As Long, xPos As Long

  On Error GoTo EH
  If Not FRM Is Nothing Then 'To Forms

    With FRM
      Let i = .Height
      Let xPos = .Top
      Let .Height = 0
      Let .Top = -xPos * 100
      On Error Resume Next
      Set RngBox = Application.InputBox(sPrompt, sTitle, Type:=8)
      On Error GoTo -1
      Let .Top = xPos
      Let .Height = i
      Set ViewsCaptureRange = RngBox
    End With

  Else 'For selections

    On Error Resume Next
    Set RngBox = Application.InputBox(sPrompt, sTitle, Type:=8)
    On Error GoTo -1
    Set ViewsCaptureRange = RngBox

  End If

EH:
  Set RngBox = Nothing
  Call VRES.ErrorHandler.ErrorHandlerRaise("VIEWSPARENT::CaptureRange", vbInformation)

End Function
'ooooooooooooooooooooooo---  FORMS AND RANGES MANNAGEMENT  ---ooooooooooooooooooooooo'



'ooooooooooooooooooooooooooooo---  FORMS CUSTOMIZING  ---oooooooooooooooooooooooooooo'
Public Sub ViewFormCleanner(ByVal VRES As ClassEqlResources, ByVal FRM As MSForms.UserForm)

  Dim CtrlA As MSForms.Control

  If FRM Is Nothing Then GoTo EH

  With FRM 'First step: Clean controls if it exists in form
    For Each CtrlA In FRM.Controls
      If Not CtrlA Is Nothing Then Call FRM.Controls.Remove(CtrlA.Name)
    Next CtrlA
  End With

EH:
  Set CtrlA = Nothing
  Call VRES.ErrorHandler.ErrorHandlerRaise("VIEWSPARENT::FormCleanner", vbInformation)

End Sub

Public Sub ViewFormAddControl( _
  ByVal VRES As ClassEqlResources, _
  ByVal CtlParent As Object, _
  ByVal CtlType As String, _
  ByVal CtlName As String, _
  Optional CtlOccurences As Long = 0, _
  Optional CtlCombinedName As Boolean = True, _
  Optional CtlFromX As Long = 0, _
  Optional CtlVisible As Boolean = True)

  Dim i As Long

  On Error GoTo EH

  With CtlParent
    For i = CtlFromX To CtlOccurences
      Call .Controls.Add("Forms." & CtlType & ".1", CtlName & IIf(CtlCombinedName, i, VBA.vbNullString), CtlVisible)
    Next i
  End With

EH:
  Call VRES.ErrorHandler.ErrorHandlerRaise("VIEWSPARENT::FormAddControl", vbInformation)

End Sub

Public Sub ViewFormStyles(ByVal VRES As ClassEqlResources, ByVal FRM As MSForms.UserForm)

  Dim Ctrl As MSForms.Control

  On Error GoTo EH
  If Not FRM Is Nothing Then
    With FRM
      For Each Ctrl In .Controls
        If TypeName(Ctrl) = "ComboBox" Then Ctrl.Style = fmStyleDropDownList ': If Ctrl.ListIndex > 0 Then Ctrl.ListIndex = 0
        If TypeName(Ctrl) = "Label" Or TypeName(Ctrl) = "CheckBox" Or TypeName(Ctrl) = "OptionButton" Then Ctrl.BackStyle = 0
        If TypeName(Ctrl) = "Frame" Or TypeName(Ctrl) = "ComboBox" Or TypeName(Ctrl) = "TextBox" Or TypeName(Ctrl) = "RefEdit" Or TypeName(Ctrl) = "ListBox" Then Ctrl.BorderStyle = 1
        If TypeName(Ctrl) = "CheckBox" Or TypeName(Ctrl) = "ListBox" Then Ctrl.SpecialEffect = 0
      Next Ctrl
    End With
  End If

EH:
  Call VRES.ErrorHandler.ErrorHandlerRaise("VIEWSPARENT::FormStyles", vbInformation)

End Sub
'ooooooooooooooooooooooooooooo---  FORMS CUSTOMIZING  ---oooooooooooooooooooooooooooo'



'ooooooooooooooooooooooooo---  STATIC TABLES CUSTOMIZING  ---oooooooooooooooooooooooo'
Public Sub ViewsTableStaticSet( _
  ByVal VRES As ClassEqlResources, _
  ByVal VRngBox As Range, _
  Optional VBooRowHeader As Boolean = False, _
  Optional VBooRowSubHeader As Boolean = False, _
  Optional VBooFirstCol As Boolean = False, _
  Optional VBooRowResults As Boolean, _
  Optional VStrBodyNumberFormat As String = "General", _
  Optional VVarHeaderHorizontalAlignement As XlHAlign = xlHAlignGeneral, _
  Optional VVarBodyHorizontalAlignement As XlHAlign = xlHAlignGeneral, _
  Optional VBooBorderLines As Boolean = False, _
  Optional VBooHorizontalLines As Boolean = False, _
  Optional VBooVerticalLines As Boolean = False, _
  Optional VBooFirstColShadow As Boolean = False, _
  Optional VInsideVerticalsJustInBody As Boolean = False, _
  Optional VBooTableType As Boolean = False, _
  Optional VThmColor As XlThemeColor, _
  Optional VThmFontColor As XlThemeColor)

  Dim ThmColor As XlThemeColor, ThmFontColor As XlThemeColor
  Dim BooTableType As Boolean
  Dim xRows As Long, xCols As Long

  Const sFN As String = "VIEWSPARENT::TableStaticSet"

  On Error GoTo EH

  #If Debugging Then
    Call Debug_assert(VRngBox Is Nothing, sFN)
  #End If

  Let BooTableType = VBooTableType 'VBA.CBool(xMDL.Tables)
  Let ThmColor = VThmColor 'xMDL.TableThemeColor
  Let ThmFontColor = VThmFontColor 'xMDL.TableThemeFontColor

  If VBooBorderLines And BooTableType Then Call VRngBox.BorderAround(xlContinuous, xlThin, ThemeColor:=ThmColor)
  If VBooHorizontalLines And BooTableType Then
    With VRngBox.Borders(xlInsideHorizontal)
      Let .LineStyle = xlContinuous
      Let .Weight = xlHairline
      Let .ThemeColor = ThmColor
    End With
  End If
  If VBooVerticalLines And BooTableType Then
    If VInsideVerticalsJustInBody Then
      With VRngBox.Range("A2").Resize(VRngBox.Rows.Count - 2, VRngBox.Columns.Count).Borders(xlInsideVertical)
        Let .LineStyle = xlContinuous
        Let .Weight = xlHairline
        Let .ThemeColor = ThmColor
      End With
    Else
      With VRngBox.Borders(xlInsideVertical)
        Let .LineStyle = xlContinuous
        Let .Weight = xlHairline
        Let .ThemeColor = ThmColor
      End With
    End If
  End If

  If VBooRowHeader Then
    With VRngBox.Range("A1").Resize(1, VRngBox.Columns.Count)
      If BooTableType Then
        Let .Interior.Pattern = xlSolid
        Let .Interior.ThemeColor = ThmColor
        Let .Font.ThemeColor = ThmFontColor
      End If
      Let .Font.Bold = True
      Let .HorizontalAlignment = VVarHeaderHorizontalAlignement
    End With
  End If
  If VBooRowSubHeader Then
    With VRngBox.Range("A2").Resize(1, VRngBox.Columns.Count)
      Let .Font.Bold = True
      Let .HorizontalAlignment = VVarHeaderHorizontalAlignement
    End With
  End If
  If VBooFirstCol Then
    With VRngBox.Range("A2").Resize(VRngBox.Rows.Count - 1, 1)
      Let .Font.Bold = True
      If VBooFirstColShadow And BooTableType Then
        Let .Interior.Pattern = xlSolid
        Let .Interior.ThemeColor = ThmColor
        Let .Font.ThemeColor = ThmFontColor
      End If
    End With
  End If
  If VBooRowResults Then
    With VRngBox.Range("A1").Offset(VRngBox.Rows.Count - 1, 0).Resize(1, VRngBox.Columns.Count)
      If BooTableType Then
        Let .Borders(xlEdgeTop).Weight = xlThin
        Let .Borders(xlEdgeBottom).LineStyle = xlDouble
      End If
      Let .Font.Bold = True
    End With
  End If

  Let xRows = VRngBox.Rows.Count - IIf(VBooRowHeader, 1, 0)
  Let xCols = VRngBox.Columns.Count - IIf(VBooFirstCol, 1, 0)
  With VRngBox.Range("A" & IIf(VBooRowHeader, 2, 1)).Offset(0, IIf(VBooFirstCol, 1, 0)).Resize(IIf(xRows <= 0, 1, xRows), IIf(xCols <= 0, 1, xCols))
    Let .NumberFormat = VStrBodyNumberFormat
    Let .HorizontalAlignment = VVarBodyHorizontalAlignement
  End With

EH:
  Call VRES.ErrorHandler.ErrorHandlerRaise(sFN, vbInformation)

End Sub
' ooooooooooooooooooooooooo---  STATIC TABLES CUSTOMIZING  ---oooooooooooooooooooooooo'


' AUGMENTED From eq
' oooooooooooooooooooooooooo---  LISTOPBJECTS CUSTOMIZING  ---oooooooooooooooooooooooo'
Public Sub ViewsTableStylesLOS( _
  ByVal VRES As ClassEqlResources, _
  ByVal wSheet As Worksheet, _
  ByVal sTable As String, _
  ByVal bFirstCol As Boolean, _
  Optional bSetIndex As Boolean = False, _
  Optional bHasTitle As Boolean = False, _
  Optional bColsAutofit As Boolean = False, _
  Optional VTableType As Byte, _
  Optional VTableStyle As Byte, _
  Optional VTableRowStripe As Boolean)

  Dim StrTableStyle As String
  Dim BooRowStripes As Boolean
  Dim xStyles As Byte, xType As Byte

  On Error GoTo EH
  Let xType = VTableType 'xMDL.Tables
  Let xStyles = VTableStyle 'xMDL.TableStyles
  Let BooRowStripes = VTableRowStripe 'xMDL.TableRowStripes

  Let StrTableStyle = VBA.vbNullString
  Let StrTableStyle = IIf(xStyles = 0, "TableStyleMedium4", IIf(xStyles = 1, "TableStyleMedium7", "TableStyleMedium2")) 'Se debe obtener de excliqprintoptions (Con o sin estilos)

  If wSheet Is Nothing Or sTable = VBA.vbNullString Then GoTo EH
  With wSheet
    With .ListObjects(sTable)
      If bHasTitle Then Let .HeaderRowRange.Range("A1").Offset(-1, 0).Font.Bold = True
      If bSetIndex Then
        If .DataBodyRange.Rows.Count > 0 Then
          Let .DataBodyRange(1, 1).Value = 1
          If .DataBodyRange.Rows.Count > 1 Then Let .DataBodyRange(2, 1).Value = 2
        End If
        If .DataBodyRange.Rows.Count > 2 Then
          Call .DataBodyRange.Range("A1:A2").AutoFill(Destination:=.DataBodyRange.Range("A1:A" & .DataBodyRange.Rows.Count), Type:=xlFillDefault)
        End If
      End If
      'Let .TableStyle = StrTableStyle
      If bColsAutofit Then Call .Range.Columns.AutoFit
      If bFirstCol Then .ShowTableStyleFirstColumn = True
      Let .ShowTableStyleRowStripes = BooRowStripes
      'If bTotals Then .ShowTotals = True
      Select Case xType
        Case 0: Let .TableStyle = VBA.vbNullString: Call .Unlist
        Case 1: Let .TableStyle = StrTableStyle: Call .Unlist
        Case 2: Let .TableStyle = StrTableStyle
      End Select
    End With
  End With

EH:
  Call VRES.ErrorHandler.ErrorHandlerRaise("VIEWSPARENT::ViewsTableStylesLOS", vbInformation)
  'Call Views_ErrorHandler("VIEWSPARENT::ViewsTableStylesLOS", vbInformation)

End Sub
' oooooooooooooooooooooooooo---  LISTOPBJECTS CUSTOMIZING  ---oooooooooooooooooooooooo'





' oooooooooooooooooooooooooo---  XL GRAPHICS CUSTOMIZING  ---ooooooooooooooooooooooooo'
Public Sub ViewsSetGraphic( _
  ByVal VRES As ClassEqlResources, _
  ByVal Ws As Worksheet, _
  ByVal selfSheet As Boolean, _
  ByVal rangeData As Variant, _
  Optional snapRange As Range, _
  Optional nameGraph As String = "Gráfico", _
  Optional titleGraph As String = VBA.vbNullString, _
  Optional titleX As String = "X", _
  Optional titleY As String = "Y", _
  Optional aGraphType As Variant = xlXYScatter, _
  Optional aTrendLines As Boolean, _
  Optional dEquation As Boolean, _
  Optional dR2 As Boolean, _
  Optional lineMinor As Boolean, _
  Optional lineMajor As Boolean, _
  Optional VBooDataAsString As Boolean, _
  Optional VVecStrData As Variant)

  Dim aGraph As Object 'Variant 'ChartObject

  On Error GoTo EH
  If selfSheet Then
    Set aGraph = Ws.ChartObjects.Add( _
      Left:=IIf(Not VBA.IsMissing(snapRange), snapRange.Left, 100), _
      Width:=400, _
      Top:=IIf(Not VBA.IsMissing(snapRange), snapRange.Top, 100), _
      Height:=250).Chart
  Else
    Set aGraph = VRES.ExcelCreateChartSheet(nameGraph, Ws.Parent) 'ActiveWorkbook.Charts.Add2
  End If

  With aGraph
    Let .ChartType = aGraphType 'xlXYScatter

    If VBooDataAsString Then 'Data incoming is type string
      Call .SeriesCollection.NewSeries
      Let .SeriesCollection(1).XValues = "={" & VBA.CStr(VVecStrData(0)) & "}" '"={4,3,5,3,2}"
      Let .SeriesCollection(1).Values = "={" & VBA.CStr(VVecStrData(1)) & "}"
    Else
      Call .SetSourceData(rangeData) 'ws.Range(rangeData)
    End If

    If .HasLegend Then Call .Legend.Delete
    If .HasTitle Then
      Let .ChartTitle.Text = titleGraph
    Else
      Call .SetElement(msoElementChartTitleAboveChart)
      Let .ChartTitle.Text = titleGraph
    End If
    With .Axes.Item(xlCategory)
      Let .HasTitle = True
      Let .AxisTitle.Caption = titleX
    End With
    With .Axes.Item(xlValue)
      Let .HasTitle = True
      Let .AxisTitle.Caption = titleY
    End With
    If aTrendLines Then
      Call .FullSeriesCollection(1).Trendlines.Add
      With .FullSeriesCollection(1).Trendlines(1)
        Let .DisplayEquation = dEquation 'True
        Let .DisplayRSquared = dR2 'True
        With .Format.Line
          Let .ForeColor.ObjectThemeColor = msoThemeColorAccent6
          Let .DashStyle = msoLineSolid
          Let .Weight = 1.55
        End With
        With .DataLabel
          Let .Left = 240
          Let .Top = 155
          Let .HorizontalAlignment = xlHAlignLeft
        End With
      End With
      Let .FullSeriesCollection(1).MarkerStyle = 8
      Let .FullSeriesCollection(1).MarkerSize = 3.5
    End If
    If lineMinor Then Call .SetElement(msoElementPrimaryValueGridLinesMinorMajor)
    If lineMajor Then Call .SetElement(msoElementPrimaryCategoryGridLinesMinorMajor)
  End With

EH:
  Set aGraph = Nothing
  Call VRES.ErrorHandler.ErrorHandlerRaise("VIEWSPARENT::viewsSetGraphic", vbInformation)
  'Call Views_ErrorHandler("VIEWSPARENT::viewsSetGraphic", vbInformation)

End Sub
' oooooooooooooooooooooooooo---  XL GRAPHICS CUSTOMIZING  ---ooooooooooooooooooooooooo'


' Sheets
' ooooooooooooooooooooooooooo---  XL PAGES CUSTOMIZING  ---ooooooooooooooooooooooooooo'
Public Sub ViewsPageSetup( _
  ByVal VRES As ClassEqlResources, _
  ByVal Sh As Worksheet, _
  Optional shRange As String = "A1", _
  Optional ShOrientation As XlPageOrientation = xlPortrait, _
  Optional ShPaperSize As XlPaperSize = xlPaperLetter, _
  Optional ShMarginTop As Double = -1, _
  Optional ShMarginRight As Double = -1, _
  Optional ShMarginBottom As Double = -1, _
  Optional ShMarginLeft As Double = -1, _
  Optional ShMarginHeader As Double = -1, _
  Optional ShMarginFooter As Double = -1, _
  Optional ShHeaderLeft As String = VBA.vbNullString, _
  Optional ShHeaderCenter As String = VBA.vbNullString, _
  Optional ShHeaderRight As String = VBA.vbNullString, _
  Optional ShFooterLeft As String = VBA.vbNullString, _
  Optional ShFooterCenter As String = VBA.vbNullString, _
  Optional ShFooterRight As String = VBA.vbNullString, _
  Optional ShImgHeader As Single = -1, _
  Optional ShImgFooter As Single = -1, _
  Optional ShImgFileHeader As String = VBA.vbNullString, _
  Optional ShImgFileFooter As String = VBA.vbNullString, _
  Optional ShZoom As Integer = 100)

  On Error GoTo EH

  If Sh Is Nothing Then Set Sh = ActiveSheet

  Let Application.PrintCommunication = False
  With Sh.PageSetup
    Let .Orientation = ShOrientation
    Let .PaperSize = ShPaperSize
    If ShMarginTop >= 0 Then Let .TopMargin = Application.CentimetersToPoints(ShMarginTop)
    If ShMarginRight >= 0 Then Let .RightMargin = Application.CentimetersToPoints(ShMarginRight)
    If ShMarginBottom >= 0 Then Let .BottomMargin = Application.CentimetersToPoints(ShMarginBottom)
    If ShMarginLeft >= 0 Then Let .LeftMargin = Application.CentimetersToPoints(ShMarginLeft)

    If ShMarginHeader >= 0 Then Let .HeaderMargin = Application.CentimetersToPoints(ShMarginHeader)
    If ShMarginFooter >= 0 Then Let .FooterMargin = Application.CentimetersToPoints(ShMarginFooter)

    ' Header & Footer Configurations
    Let .LeftHeader = ShHeaderLeft
    Let .CenterHeader = ShHeaderCenter
    Let .RightHeader = ShHeaderRight
    Let .LeftFooter = ShFooterLeft
    Let .CenterFooter = ShFooterCenter
    Let .RightFooter = ShFooterRight

    ' Image header
    If Not ShImgFileHeader = VBA.vbNullString Then
      If ShImgHeader = 0 Then ' LeftHeader
        With .LeftHeaderPicture
          Let .FileName = ShImgFileHeader
          Let .Height = 100
          Let .Width = 100
        End With
        Let .LeftHeader = IIf(VBA.IsEmpty(ShHeaderLeft), "&G", "&G" & .LeftHeader)
      End If
      If ShImgHeader = 1 Then ' CenterHeader
        With .CenterHeaderPicture
          Let .FileName = ShImgFileHeader
          Let .Height = 100
          Let .Width = 100
        End With
        Let .CenterHeader = IIf(VBA.IsEmpty(ShHeaderCenter), "&G", "&G" & .CenterHeader)
      End If
      If ShImgHeader = 2 Then ' RightHeader
        With .RightHeaderPicture
          Let .FileName = ShImgFileHeader
          Let .Height = 100
          Let .Width = 100
        End With
        Let .RightHeader = IIf(VBA.IsEmpty(ShHeaderRight), "&G", "&G" & .RightHeader)
      End If
    End If

    ' Image footer
    If ShImgFileFooter = VBA.vbNullString Then
      If ShImgFooter = 0 Then ' LeftFooter
        With .LeftFooterPicture
          Let .FileName = ShImgFileFooter
          Let .Height = 100
          Let .Width = 100
        End With
        Let .LeftFooter = IIf(VBA.IsEmpty(ShFooterLeft), "&G", "&G" & .LeftFooter)
      End If
      If ShImgFooter = 1 Then ' CenterFooter
        With .CenterFooterPicture
          Let .FileName = ShImgFileFooter
          Let .Height = 100
          Let .Width = 100
        End With
        Let .CenterFooter = IIf(VBA.IsEmpty(ShFooterCenter), "&G", "&G" & .CenterFooter)
      End If
      If ShImgFooter = 2 Then ' RightFooter
        With .RightFooterPicture
          Let .FileName = ShImgFileFooter
          Let .Height = 100
          Let .Width = 100
        End With
        Let .RightFooter = IIf(VBA.IsEmpty(ShFooterRight), "&G", "&G" & .RightFooter)
      End If
    End If

    Let .PrintArea = shRange
    If Not ShZoom = 100 Then Let .Zoom = ShZoom
  End With

EH:
  Let Application.PrintCommunication = True
  Call VRES.ErrorHandler.ErrorHandlerRaise("VIEWSPARENT::PageSetup", vbInformation)

End Sub
' ooooooooooooooooooooooooooo---  XL PAGES CUSTOMIZING  ---ooooooooooooooooooooooooooo'
' CORE METHODS LIST (PUBLIC) ------------------------------------------------------- '
' ===================== CLASS EXCLIQ LITE VIEW PARENT STRUCTURE ==================== '

