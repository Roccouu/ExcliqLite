VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassEqlResources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================== CLASS EXCLIQ LITE RES STRUCTURE ========================= '
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


' ======================== CLASS EXCLIQ LITE RES STRUCTURE ========================= '
' OBJECT VARIABLES (GLOOBJ_) ------------------------------------------------------- '
' VARIABLES (GLOStr_) -------------------------------------------------------------- '
' PROPERTIES ----------------------------------------------------------------------- '
' CONSTRUCTOR/DESTRUCTOR ----------------------------------------------------------- '
' CORE METHODS LIST (PRIVATE) ------------------------------------------------------ '
' ERRORS MANNAGEMENT AND ASSERTSIONS ----------------------------------------------- '
' Debug_assert
' CORE METHODS LIST (PUBLIC) ------------------------------------------------------- '
' EXCEL's IMPROVED FOR HIGH PERFORMANCE METHODS ------------------------------------ '
' ExcelCurrentOrNewSheet
' ExcelSearchSheet
' ExcelCountSheetNameCoincidences
' ExcelCreateSheet
' ExcelExistSheet
' ExcelExistsWorkbook
' ExcelDelSheet
' ExcelCreateChartSheet
' ExcelEdges
' ExcelIsProtectedSheet
' FORMS's IMPROVED FOR HIGH PERFORMANCE METHODS ------------------------------------ '
' FrmCMsgbox
' FrmSelTextControl
' ARRAY's IMPROVED FOR HIGH PERFORMANCE METHODS ------------------------------------ '
' ArrayEmpty
' ArrayLength
' ArrayToBaseZero
' ArrayGetDataFromRangesByRows
' ArrayGetDataFromRangesByCols
' ArrayGetDataAsStringFromRanges
' ArrayGetDataColonSeparatedFromString
' ArrayGetDataSemicolonSeparatedSlashGroupedStringFromVariants
' ArrayGetDataXCharSeparatedStringFromRanges
' ArrayGetDataXCharSeparatedStringFromVariantsAsString
' ArrayGetDataXCharSeparatedStringFromVariantsAsVector
' ArrayHasValue
' ArrayHasValue2D
' ArrayHasValueN
' ArrayHasValueN2D
' ArrayGetValuePosition
' ArrayHasRepeatedElements
' ArrayShift
' ArrayPop
' ArraySlice
' ArrayAddAtFirst
' ArrayAddAtLast
' ArrayDelIndex
' ArrayAddElement
' ArrayConcat
' ArrayTranspose
' ArrayUniques
' ArrayDimension
' ArrayNumbersCommasToDots
' ======================== CLASS EXCLIQ LITE RES STRUCTURE ========================= '



' OBJECT VARIABLES (GLOOBJ_) ------------------------------------------------------- '
Private EHGLOBAL As ClassEqlErrorHandler
Public REGEX As ClassEqlResourcesRegex
Public RESCORE As ClassEqlResourcesHelper

' VARIABLES (GLOStr_) -------------------------------------------------------------- '
Private GLOBOO_FORMULAS As Boolean
' System main data
Private GLOStr_AppTitle As String



' PROPERTIES ----------------------------------------------------------------------- '
' Error handler
Public Property Set ErrorHandler(ByRef ObjEH As ClassEqlErrorHandler)

  Set EHGLOBAL = ObjEH

  Set REGEX = New ClassEqlResourcesRegex
  Set REGEX.ErrorHandler = EHGLOBAL
  Set RESCORE = New ClassEqlResourcesHelper
  Set RESCORE.ErrorHandler = EHGLOBAL

End Property

Public Property Get ErrorHandler() As ClassEqlErrorHandler

  Set ErrorHandler = EHGLOBAL

End Property

' System main data
Public Property Let AppStrTitle(ByVal RStrTitle As String)

  Let GLOStr_AppTitle = RStrTitle

End Property

Public Property Get AppStrTitle() As String

  Let AppStrTitle = GLOStr_AppTitle

End Property

' System configuration
Public Property Let Formulas(ByVal RBooFormulas As Boolean)

  Let GLOBOO_FORMULAS = RBooFormulas

  If Not RESCORE Is Nothing Then Let RESCORE.Formulas = RBooFormulas

End Property

Public Property Get Formulas() As Boolean

  Let Formulas = GLOBOO_FORMULAS

End Property


' CONSTRUCTOR/DESTRUCTOR ----------------------------------------------------------- '
Private Sub Class_Terminate()

  Set RESCORE = Nothing
  Set REGEX = Nothing
  Set EHGLOBAL = Nothing

End Sub


' CORE METHODS LIST (PRIVATE) ------------------------------------------------------ '
' ERRORS MANNAGEMENT AND ASSERTSIONS ----------------------------------------------- '
' Debug_assert
Private Sub Debug_assert(ByVal bDeb As Boolean, Optional sCaller As String = VBA.vbNullString)

  #If Debugging Then
    Debug.Assert bDeb
    If Not bDeb Then Debug.Print sCaller
  #End If

End Sub


' CORE METHODS LIST (PUBLIC) ------------------------------------------------------- '
' EXCEL's IMPROVED FOR HIGH PERFORMANCE METHODS ------------------------------------ '
Public Function ExcelCurrentOrNewSheet(ByVal NameSheet As String, Optional Wbk As Workbook) As Worksheet

  ' Check if Worksheet named NameSheet exists, if doesn't creates it, else creates the new sheet with NameSheet_x.
  Dim Wb As Workbook
  Dim newSheetName As String

  Const sFN As String = "RES::ExcelCreateChartSheet"

  On Error GoTo EH
  If Wbk Is Nothing Then Set Wb = Application.ActiveWorkbook
  If Not Wbk Is Nothing Then Set Wb = Wbk

  If Not Wb.ProtectStructure Then

    ' Creates the new worksheet
    Let newSheetName = Me.ExcelCountSheetNameCoincidences(NameSheet, Wb)
    If Me.ExcelCreateSheet(newSheetName, Wbk:=Wb) Then
      Call Wb.Sheets(newSheetName).Activate
      Set ExcelCurrentOrNewSheet = Wb.Sheets(newSheetName)
    End If

  End If

EH:
  Set Wb = Nothing
  Call EHGLOBAL.ErrorHandlerRaise("RES::ExcelCurrentOrNewSheet")

End Function

Public Function ExcelSearchSheet(ByVal NameSheet As String, Optional Wbk As Workbook) As Boolean

  ' FIXME: REBUILD
  Dim Wb As Workbook

  On Error Resume Next
  If Wbk Is Nothing Then Set Wb = ActiveWorkbook
  If Not Wbk Is Nothing Then Set Wb = Wbk
  Let ExcelSearchSheet = (Wb.Worksheets(NameSheet).Name <> VBA.vbNullString)

  Set Wb = Nothing
  On Error GoTo 0

End Function

Public Function ExcelCountSheetNameCoincidences(ByVal ShName As String, Optional Wbk As Workbook) As String

  ' FIXME: REBUILD
  Dim i As Integer
  Dim Wb As Workbook
  Dim ShNameLocal As String, ShNameAux As String

  Const sFN As String = "RES::ExcelCountSheetNameCoincidences"

  On Error GoTo EH
  If Wbk Is Nothing Then Set Wb = Application.ActiveWorkbook
  If Not Wbk Is Nothing Then Set Wb = Wbk
  Let ExcelCountSheetNameCoincidences = ShName
  Let i = 0

  With Wb
    If Me.ExcelExistSheet(ShName, Wbk) Then
      Do
        Let i = i + 1
        Let ShNameLocal = ShName & "_" & i
      Loop While (Me.ExcelExistSheet(ShNameLocal, Wbk))
      Let ExcelCountSheetNameCoincidences = ShNameLocal
    End If
  End With

EH:
  Set Wb = Nothing
  If VBA.Err.Number <> 0 Then
    Let ShNameAux = "Sheet" & Wb.Sheets.Count + 1
    Let ExcelCountSheetNameCoincidences = ShNameAux
    'Call VBA.MsgBox("Imposible asignar el nombre: " & ShName & " a la nueva hoja, se creó la hoja: " & ShNameAux & " en su lugar.", vbInformation, GLOStr_AppTitle)
    Call EHGLOBAL.ErrorHandlerRaise("RES::ExcelCountSheetNameCoincidences", vbInformation)
  End If

End Function

Public Function ExcelCreateSheet(ByVal NameSheet As String, Optional AfterSheet As String = VBA.vbNullString, Optional BeforeSheet As String = VBA.vbNullString, Optional Wbk As Workbook) As Boolean

  Dim afS As Boolean, bfS As Boolean
  Dim Wb As Workbook

  If Wbk Is Nothing Then Set Wb = ActiveWorkbook
  If Not Wbk Is Nothing Then Set Wb = Wbk

  Let afS = IIf(AfterSheet <> VBA.vbNullString, ExcelSearchSheet(AfterSheet, Wb), False)
  Let bfS = IIf(BeforeSheet <> VBA.vbNullString, ExcelSearchSheet(BeforeSheet, Wb), False)

  On Error Resume Next
  If ExcelSearchSheet(NameSheet, Wb) Then
    Let ExcelCreateSheet = False
  Else
    If afS Then
      Let Wb.Worksheets.Add(After:=Wb.Worksheets(AfterSheet)).Name = NameSheet
    ElseIf bfS Then
      Let Wb.Worksheets.Add(before:=Wb.Worksheets(BeforeSheet)).Name = NameSheet
    Else
      Let Wb.Worksheets.Add(After:=Wb.Worksheets(Wb.Worksheets.Count)).Name = NameSheet
    End If
    Let ExcelCreateSheet = True
  End If

  Set Wb = Nothing
  On Error GoTo 0

End Function

Public Function ExcelExistSheet(ByVal NameSheet As String, Optional Wbk As Workbook) As Boolean

  Dim Wb As Workbook
  Dim i As Long

  Const sFN As String = "RES::ExcelExistSheet"

  On Error GoTo EH
  If Wbk Is Nothing Then Set Wb = ActiveWorkbook
  If Not Wbk Is Nothing Then Set Wb = Wbk

  Let ExcelExistSheet = False
  With Wb
    For i = 1 To .Worksheets.Count
      If (VBA.LCase(.Worksheets(i).Name) = VBA.LCase(NameSheet)) Then Let ExcelExistSheet = True: Exit For
    Next i
  End With

EH:
  Set Wb = Nothing
  Call EHGLOBAL.ErrorHandlerRaise("RES::ExcelExistSheet")

End Function

Public Function ExcelExistsWorkbook(ByVal RStrWbkName As String) As Boolean

  ' NEW on both projects
  Dim BooExists As Boolean
  Dim i As Long

  Const sFN As String = "RES::ExcelExistsWorkbook"

  On Error GoTo EH

  Let BooExists = False
  For i = 1 To Application.Workbooks.Count
    If VBA.LCase(Application.Workbooks(i).Name) = VBA.LCase(RStrWbkName) Then Let BooExists = True: Exit For
  Next i

  Let ExcelExistsWorkbook = BooExists

EH:
  Call EHGLOBAL.ErrorHandlerRaise("RES::ExcelExistsWorkbook")

End Function

Public Sub ExcelDelSheet(ByVal RStrNameSheet As String, Optional RWbk As Workbook = Nothing, Optional RStrWbk As String = VBA.vbNullString)

  ' NEW on both projects
  Dim Wbk As Workbook
  Dim Sht As Worksheet
  Dim i As Long

  Const sFN As String = "RES::ExcelDelSheet"

  On Error GoTo EH

  If RWbk Is Nothing Then
    If RStrWbk = VBA.vbNullString Then
      Set Wbk = Application.ActiveWorkbook
    Else
      If ExcelExistsWorkbook(RStrWbk) Then Set Wbk = Application.Workbooks(RStrWbk)
    End If
  Else
    Set Wbk = RWbk
  End If

  ' Del sheet
  Let Application.DisplayAlerts = False
  Let Application.ScreenUpdating = False
  If Not Wbk Is Nothing Then

    If Me.ExcelExistSheet(RStrNameSheet, Wbk) Then
      Set Sht = Wbk.Worksheets(RStrNameSheet)
      Call Sht.Delete
      Set Sht = Nothing
      Set Wbk = Nothing
    End If

  End If

EH:
  Set Sht = Nothing
  Set Wbk = Nothing
  Let Application.ScreenUpdating = True
  Let Application.DisplayAlerts = True
  Call EHGLOBAL.ErrorHandlerRaise("RES::ExcelDelSheet")

End Sub

Public Function ExcelCreateChartSheet(ByVal NameSheet As String, Optional Wbk As Workbook, Optional AfterSheet As String = VBA.vbNullString, Optional BeforeSheet As String = VBA.vbNullString) As Chart

  ' FIXME: REBUILD
  Dim Wb As Workbook
  Dim chartSheet As Chart

  If Wbk Is Nothing Then Set Wb = ActiveWorkbook
  If Not Wbk Is Nothing Then Set Wb = Wbk

  On Error Resume Next
  Let NameSheet = Me.ExcelCountSheetNameCoincidences(NameSheet, Wb)
  Set chartSheet = Wbk.Charts.Add2
  Let chartSheet.Name = NameSheet
  Set ExcelCreateChartSheet = chartSheet
  Set chartSheet = Nothing
  Set Wb = Nothing
  On Error GoTo 0

End Function

Public Function ExcelEdges(ByVal xRng As Range, ByVal xRow As Long, ByVal xCol As Long) As Boolean

  Dim Sh As Worksheet
  Dim Rn As Range

  On Error GoTo EH
  Set Sh = xRng.Parent
  Set Rn = xRng.Range("A1")
  Let ExcelEdges = (((Rn.row + xRow) > (Sh.Rows.Count)) Or ((Rn.Column + xCol) > (Sh.Columns.Count)))
  Set Rn = Nothing
  Set Sh = Nothing

EH:
  Set Rn = Nothing
  Set Sh = Nothing
  Call EHGLOBAL.ErrorHandlerRaise("RES::ExcelEdges")

End Function

Public Function ExcelIsProtectedSheet(ByVal RSht As Worksheet) As Boolean

  ' Check if worksheet is protected
  On Error GoTo EH

  With RSht
    Let ExcelIsProtectedSheet = (.ProtectContents Or .ProtectDrawingObjects Or .ProtectScenarios)
  End With

EH:
  Call EHGLOBAL.ErrorHandlerRaise("RES::ExcelIsProtectedSheet")

End Function

' FORMS's IMPROVED FOR HIGH PERFORMANCE METHODS ------------------------------------ '
Public Sub FrmCMsgbox( _
  ByVal RMsg As String, _
  Optional RTitle As String = VBA.vbNullString, _
  Optional RMsgType As VbMsgBoxStyle = vbCritical, _
  Optional RCtrlReq As Boolean, _
  Optional RCtrl As MSForms.Control, _
  Optional RCtrlColor As VbMsgBoxStyle = vbCritical, _
  Optional RCtrlFocus As Boolean = False, _
  Optional RSelTxt As Boolean = False)

  Dim XlsColor As XlRgbColor, XlsTempColor As XlRgbColor
  Dim XlsBackColor As XlRgbColor, XlsTempBackColor As XlRgbColor

  ' Set FrmCMsgBox colors
  On Error GoTo EH
  Select Case RCtrlColor
    Case vbInformation
      Let XlsTempBackColor = rgbLightBlue
      Let XlsTempColor = rgbSteelBlue
    Case vbQuestion
      Let XlsTempBackColor = rgbBisque
      Let XlsTempColor = rgbDarkOrange
    Case vbExclamation
      Let XlsTempBackColor = rgbGold
      Let XlsTempColor = rgbDarkGoldenrod
    Case Else
      Let XlsTempBackColor = VBA.RGB(255, 230, 230)
      Let XlsTempColor = rgbTomato
  End Select

  ' Exists a control and it is required
  If RCtrlReq And Not RCtrl Is Nothing Then
    With RCtrl
      ' Get the control colors
      Select Case VBA.TypeName(RCtrl)
        Case "TextBox", "ComboBox", "Label", "ListBox", "Frame", "UserForm"
          Let XlsBackColor = .BackColor
          Let XlsColor = .BorderColor
          Let .BackColor = XlsTempBackColor
          Let .BorderColor = XlsTempColor
        Case "CommandButton", "ToggleButton", "CheckBox", "OptionButton", "SpinButton", "MultiPage", "SpinButton", "TabStrip"
          Let XlsBackColor = .BackColor
          Let XlsColor = .ForeColor
          Let .BackColor = XlsTempBackColor
          Let .ForeColor = XlsTempColor
      End Select
    End With
  End If

  ' Show the message
  Call VBA.MsgBox(RMsg, RMsgType, IIf(RTitle = VBA.vbNullString, Application.Name, RTitle))

  If RCtrlReq And Not RCtrl Is Nothing Then
    With RCtrl
      ' Reset the control colors
      Select Case VBA.TypeName(RCtrl)
        Case "TextBox", "ComboBox", "Label", "ListBox", "Frame", "UserForm"
          Let .BackColor = XlsBackColor
          Let .BorderColor = XlsColor
        Case "CommandButton", "ToggleButton", "CheckBox", "OptionButton", "SpinButton", "MultiPage", "SpinButton", "TabStrip"
          Let .BackColor = XlsBackColor
          Let .ForeColor = XlsColor
      End Select
    End With
  End If

  ' Focus on control
  If Not RCtrl Is Nothing Then
    With RCtrl
      If .Enabled And .Visible Then
        Select Case VBA.TypeName(RCtrl)
          Case "TextBox", "ComboBox", "Label", "ListBox"
            If RSelTxt Then Call Me.FrmSelTextControl(RCtrl)
        End Select
        If RCtrlFocus Then Call .SetFocus
      End If
    End With
  End If

EH:
  Call EHGLOBAL.ErrorHandlerRaise("RES::FrmCMsgbox")

End Sub

Public Sub FrmSelTextControl(ByVal Control As Object)

  On Error GoTo EH
  With Control
    Let .SelStart = 0
    Let .SelLength = VBA.Len(Control.Text)
    If .Enabled And .Visible Then Call .SetFocus
  End With

EH:
  Call EHGLOBAL.ErrorHandlerRaise("RES::FrmSelTextControl")

End Sub

' ARRAY's IMPROVED FOR HIGH PERFORMANCE METHODS ------------------------------------ '
Function ArrayEmpty(ByRef arr As Variant) As Boolean
  On Error GoTo EH

  Dim i As Long, length As Long, Locallength As Long
  Let Locallength = 1
  Let length = 1
  Let ArrayEmpty = False

  ' Loop until no more dimensions
  Do While True
    Let i = i + 1
    ' If the array has no items then this line will throw an error
    Let length = Locallength * (UBound(arr, i) - LBound(arr, i) + 1)
    ' Set ArrayLength here to avoid returing 1 for an empty array
    Let ArrayEmpty = (length > 0)
    If ArrayEmpty Then Exit Do
  Loop

Done:
  Exit Function
EH:
  If VBA.Err.Number = 13 Then ' Type Mismatch Error
    Call EHGLOBAL.ErrorHandlerRaise("RES::ArrayLength", vbInformation)
  Else
    Call EHGLOBAL.ErrorHandlerClear
  End If

End Function

Function ArrayLength(ByRef arr As Variant) As Long

  'https://excelmacromastery.com/
  On Error GoTo EH

  ' Loop is used for multidimensional arrays. The Loop will terminate when a
  ' "Subscript out of Range" error occurs i.e. there are no more dimensions.
  Dim i As Long, length As Long
  Let length = 1

  If Not VBA.IsArray(arr) Then Call VBA.Err.Raise(CUSTOM_ERROR_RES, "RES::ArrayLength", "¡Se esperaba un array!")

  ' Loop until no more dimensions
  Do While True
    Let i = i + 1
    ' If the array has no items then this line will throw an error
    Let length = length * (UBound(arr, i) - LBound(arr, i) + 1)
    ' Set ArrayLength here to avoid returing 1 for an empty array
    Let ArrayLength = length
  Loop

'Done:
'  Exit Function
EH:
  If VBA.Err.Number = 13 Then ' Type Mismatch Error
    Call EHGLOBAL.ErrorHandlerRaise("RES::ArrayLength", vbInformation)
  Else
    Call EHGLOBAL.ErrorHandlerClear
  End If

End Function

Public Function ArrayToBaseZero( _
  ByVal AVarVector As Variant, _
  Optional AEnuResultDimension As EQLRES_ENU_DIMENSIONARRAY = eqlResDefaultArray, _
  Optional AByCols As Boolean = True) As Variant

  Dim VecVector As Variant, VecVector0 As Variant
  Dim i As Long, j As Long, k As Long, l As Long, LngDimension As Long

  Const sFN As String = "RES::ArrayToBaseZero"

  On Error GoTo EH
  'Call Debug_assert(Not VBA.IsArray(AVarVector), sFN)

  If VBA.IsArray(AVarVector) Then
    Let i = Me.ArrayDimension(AVarVector)
    Let LngDimension = i
    If i = -1 Then
      Call VBA.Err.Raise(CUSTOM_ERROR_RES, sFN, "Se esperaba un array.")
    ElseIf i > 2 Then
      Call VBA.Err.Raise(CUSTOM_ERROR_RES, sFN, "Se esperaba un array de dimensión 1 o 2.")
    ElseIf i = 1 Then
      Let k = 0
      ReDim VecVector(0 To (UBound(AVarVector) - LBound(AVarVector)))
      For i = LBound(AVarVector) To UBound(AVarVector)
        Let VecVector(k) = AVarVector(i)
        Let k = k + 1
      Next i
    Else
      Let k = 0
      Let l = 0
      ReDim VecVector(0 To (UBound(AVarVector) - LBound(AVarVector)), 0 To (UBound(AVarVector, 2) - LBound(AVarVector, 2)))
      For j = LBound(AVarVector, 2) To UBound(AVarVector, 2)
        For i = LBound(AVarVector) To UBound(AVarVector)
          Let VecVector(k, l) = AVarVector(i, j)
          Let k = k + 1
        Next i
        Let l = l + 1
        Let k = 0
      Next j
    End If
  Else
    If VBA.IsNull(AVarVector) Then
      Call VBA.Err.Raise(CUSTOM_ERROR_RES, sFN, "Se esperaba un array de dimensión 1 o 2 sin contenidos nulos.")
    Else
      ReDim VecVector(0 To 0, 0 To 0)
      Let VecVector(0, 0) = AVarVector
    End If
  End If

  Let k = 0
  If AEnuResultDimension = eqlRes1D Then ' Response AS 1D
    If LngDimension = 2 Then ' Just if 2D
      Let i = Me.ArrayLength(VecVector) - 1
      ReDim VecVector0(0 To i)
      If AByCols Then ' Cols
        For j = LBound(VecVector, 2) To UBound(VecVector, 2)
          For i = LBound(VecVector) To UBound(VecVector)
            Let VecVector0(k) = VecVector(i, j)
            Let k = k + 1
          Next i
        Next j
      Else ' Rows
        For i = LBound(VecVector) To UBound(VecVector)
          For j = LBound(VecVector, 2) To UBound(VecVector, 2)
            Let VecVector0(k) = VecVector(i, j)
            Let k = k + 1
          Next j
        Next i
      End If
      Let VecVector = Empty
      Let VecVector = VecVector0
    End If
    'Let VecVector = Empty
    'Let VecVector = VecVector0
  ElseIf AEnuResultDimension = eqlRes2D Then ' Response as 2D
    If LngDimension = 1 Then ' Just if 1D
      If AByCols Then ' Cols
        ReDim VecVector0(0 To 0, 0 To UBound(VecVector))
        For i = LBound(VecVector) To UBound(VecVector)
          Let VecVector0(0, i) = VecVector(i)
        Next i
      Else ' Rows
        ReDim VecVector0(0 To UBound(VecVector), 0 To 0)
        For i = LBound(VecVector) To UBound(VecVector)
          Let VecVector0(i, 0) = VecVector(i)
        Next i
      End If
      Let VecVector = Empty
      Let VecVector = VecVector0
    End If
    'Let VecVector = Empty
    'Let VecVector = VecVector0
  End If

  Let ArrayToBaseZero = VecVector

EH:
  Let VecVector = Empty
  Let VecVector0 = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

' ALERT
'Public Function ArrayVectorOfNumbersFromColonSeparatedString(ByVal sVector As String, Optional bAbs As Boolean = False) As Variant
'
'  Dim vVector As Variant
'  Dim vFinal() As Variant
'  Dim i As Long
'  Dim sVectorLocal As String
'  Const sFN As String = "RES::ArrayVectorOfNumbersFromColonSeparatedString"
'
'  On Error GoTo EH
'  Call Debug_assert(Not sVector = VBA.vbNullString, sFN)
'  If sVector = VBA.vbNullString Then Call VBA.Err.Raise(CUSTOM_ERROR_RES, sFN, "Se requiere una cadena con números.")
'  Let ArrayVectorOfNumbersFromColonSeparatedString = Array()
'  Let sVectorLocal = VBA.Replace(sVector, ",", ".")
'  Let vVector = VBA.Split(sVectorLocal, ";")
'  If Me.ArrayEmpty(vVector) Then
'    ReDim vFinal(0 To UBound(vVector))
'    For i = LBound(vVector) To UBound(vVector)
'      If bAbs Then
'        Let vFinal(i) = VBA.Abs(VBA.CDbl(vVector(i)))
'      Else
'        Let vFinal(i) = VBA.CDbl(vVector(i))
'      End If
'    Next i
'  End If
'  Erase ArrayVectorOfNumbersFromColonSeparatedString
'  Let ArrayVectorOfNumbersFromColonSeparatedString = vFinal
'
'EH:
'  Erase vFinal
'  Let vVector = Empty
'  Call EHGLOBAL.ErrorHandlerRaise(sFN)
'
'End Function

Function ArrayGetDataFromRangesByRows( _
  ByVal Rng As Range, _
  ByRef RCounter As Long, _
  Optional RType As Boolean = False, _
  Optional RAbs As Boolean = False, _
  Optional RIncludeZeroes As Boolean = False, _
  Optional RIncludeEmptyStrings As Boolean = False, _
  Optional RCustomRegex As String = VBA.vbNullString) As Variant

  'Gets all data number/strings from ranges in order of Rows
  Dim r As Range
  Dim VData As Variant
  Dim i As Long, j As Long

  Let ArrayGetDataFromRangesByRows = Empty

  On Error GoTo EH
  If Rng Is Nothing Then GoTo EH

  Let i = 0
  ReDim VData(0)
  Let VData(0) = VBA.vbNullString
  If RType Then 'Strings
    If (Application.WorksheetFunction.CountA(Rng) - Application.WorksheetFunction.Count(Rng)) <= 0 Then
      Let ArrayGetDataFromRangesByRows = VData: GoTo EH
    End If
    For j = 1 To Rng.Areas.Count
      For Each r In Rng.Areas(j)
        If Not REGEX.isNumberReg(r.Value) Then
          If RIncludeEmptyStrings Then
            ReDim Preserve VData(i)
            Let VData(i) = r.Value
            Let i = i + 1
          Else
            If REGEX.isStringReg(r.Value, RCustomRegex) Then
              ReDim Preserve VData(i)
              Let VData(i) = r.Value
              Let i = i + 1
            End If
          End If
        End If
      Next r
    Next j
  Else 'Numbers
    If Application.WorksheetFunction.Count(Rng) <= 0 Then
      Let ArrayGetDataFromRangesByRows = VData: GoTo EH
    End If
    For j = 1 To Rng.Areas.Count
      For Each r In Rng.Areas(j)
        If REGEX.isNumberReg(r.Value) Then
          If RIncludeZeroes Then
            ReDim Preserve VData(i)
            Let VData(i) = IIf(RAbs, VBA.Abs(r.Value), r.Value)
            Let i = i + 1
          Else
            If VBA.Abs(r.Value) > 0 Then
              ReDim Preserve VData(i)
              Let VData(i) = IIf(RAbs, VBA.Abs(r.Value), r.Value)
              Let i = i + 1
            End If
          End If
        End If
      Next r
    Next j
  End If

  Let RCounter = i
  Let ArrayGetDataFromRangesByRows = VData

EH:
  Let VData = Empty
  Set r = Nothing
  Call EHGLOBAL.ErrorHandlerRaise("RES::ArrayGetDataFromRanges")

End Function

Function ArrayGetDataFromRangesByCols( _
  ByVal Rng As Range, _
  ByRef RCounter As Long, _
  Optional RType As Boolean = False, _
  Optional RAbs As Boolean = False, _
  Optional RIncludeZeroes As Boolean = False, _
  Optional RIncludeEmptyStrings As Boolean = False, _
  Optional RCustomRegex As String = VBA.vbNullString) As Variant

  'Gets all data number/strings from ranges in order of Rows
  Dim r As Range
  Dim VData As Variant, VecRng As Variant
  Dim i As Long, kAreas As Long, kRow As Variant, kCol As Variant
  Dim DblAux As Double
  Dim StrAux As String

  Let ArrayGetDataFromRangesByCols = Empty

  On Error GoTo EH
  If Rng Is Nothing Then GoTo EH

  Let i = 0
  ReDim VData(0)
  Let VData(0) = VBA.vbNullString
  If RType Then 'Strings
    If (Application.WorksheetFunction.CountA(Rng) - Application.WorksheetFunction.Count(Rng)) <= 0 Then
      Let ArrayGetDataFromRangesByCols = VData: GoTo EH
    End If
    ReDim VData(0)
    For kAreas = 1 To Rng.Areas.Count
      If Rng.Areas(kAreas).Cells.Count > 1 Then
        Let VecRng = Rng.Areas(kAreas).Value
        For kCol = LBound(VecRng, 2) To UBound(VecRng, 2)
          For kRow = LBound(VecRng) To UBound(VecRng)
            If Not REGEX.isNumberReg(VecRng(kRow, kCol)) Then
              Let StrAux = VecRng(kRow, kCol)
              If RIncludeEmptyStrings Then
                ReDim Preserve VData(i)
                Let VData(i) = StrAux
                Let i = i + 1
              Else
                If REGEX.isStringReg(StrAux, RCustomRegex) Then
                  ReDim Preserve VData(i)
                  Let VData(i) = StrAux
                  Let i = i + 1
                End If
              End If
            End If
          Next kRow
        Next kCol
      Else
        If Not REGEX.isNumberReg(Rng.Value) Then
          If RIncludeEmptyStrings Then
            ReDim Preserve VData(i)
            Let VData(i) = Rng.Areas(kAreas).Cells(1, 1).Value
            Let i = i + 1
          Else
            If REGEX.isStringReg(Rng.Value, RCustomRegex) Then
              ReDim Preserve VData(i)
              Let VData(i) = Rng.Areas(kAreas).Cells(1, 1).Value
              Let i = i + 1
            End If
          End If
        End If
      End If
    Next kAreas
  Else 'Numbers
    If Application.WorksheetFunction.Count(Rng) <= 0 Then
      Let ArrayGetDataFromRangesByCols = VData: GoTo EH
    End If
    ReDim VData(0)
    For kAreas = 1 To Rng.Areas.Count
      If Rng.Areas(kAreas).Cells.Count > 1 Then
        Let VecRng = Rng.Areas(kAreas).Value
        For kCol = LBound(VecRng, 2) To UBound(VecRng, 2)
          For kRow = LBound(VecRng) To UBound(VecRng)
            If REGEX.isNumberReg(VecRng(kRow, kCol)) Then
              Let DblAux = VecRng(kRow, kCol)
              If RIncludeZeroes Then
                ReDim Preserve VData(i)
                Let VData(i) = IIf(RAbs, VBA.Abs(DblAux), DblAux)
                Let i = i + 1
              Else
                If VBA.Abs(VecRng(kRow, kCol)) > 0 Then
                  ReDim Preserve VData(i)
                  Let VData(i) = IIf(RAbs, VBA.Abs(DblAux), DblAux)
                  Let i = i + 1
                End If
              End If
            End If
          Next kRow
        Next kCol
      Else
        If REGEX.isNumberReg(Rng.Value) Then
          If RIncludeZeroes Then
            ReDim Preserve VData(i)
            Let VData(i) = IIf(RAbs, VBA.Abs(Rng.Areas(kAreas).Cells(1, 1).Value), Rng.Areas(kAreas).Cells(1, 1).Value)
            Let i = i + 1
          Else
            If VBA.Abs(Rng.Value) > 0 Then
              ReDim Preserve VData(i)
              Let VData(i) = IIf(RAbs, VBA.Abs(Rng.Areas(kAreas).Cells(1, 1).Value), Rng.Areas(kAreas).Cells(1, 1).Value)
              Let i = i + 1
            End If
          End If
        End If
      End If
    Next kAreas
  End If

  Let RCounter = i
  Let ArrayGetDataFromRangesByCols = VData

EH:
  Let VData = Empty
  Set r = Nothing
  Call EHGLOBAL.ErrorHandlerRaise("RES::ArrayGetDataFromRangesByCols")

End Function

Function ArrayGetDataAsStringFromRanges( _
  ByVal Rng As Range, _
  ByRef RCounter As Long, _
  Optional RType As Boolean = False, _
  Optional RAbs As Boolean = False, _
  Optional RIncludeZeroes As Boolean = False, _
  Optional RIncludeEmptyStrings As Boolean = False, _
  Optional RCustomRegex As String = VBA.vbNullString, _
  Optional RBooRngRowCol As Boolean = False) As String

  Dim VData As Variant

  Let ArrayGetDataAsStringFromRanges = Empty

  On Error GoTo EH
  If Rng Is Nothing Then GoTo EH

  If RBooRngRowCol Then
    Let VData = ArrayGetDataFromRangesByCols(Rng, RCounter, RType, RAbs, RIncludeZeroes, RIncludeEmptyStrings, RCustomRegex)
  Else
    Let VData = ArrayGetDataFromRangesByRows(Rng, RCounter, RType, RAbs, RIncludeZeroes, RIncludeEmptyStrings, RCustomRegex)
  End If

  'Let RCounter = i
  Let ArrayGetDataAsStringFromRanges = VBA.Join(VData, ";")

EH:
  Let VData = Empty
  Call EHGLOBAL.ErrorHandlerRaise("RES::ArrayGetDataAsStringFromRanges")

End Function

Public Function ArrayGetDataColonSeparatedFromString( _
  ByVal StrVector As String, _
  ByRef StrCounter As Long, _
  Optional StrValueType As Boolean = False, _
  Optional StrAbs As Boolean = False, _
  Optional StrIncludeZeroes As Boolean = False, _
  Optional StrIncludeEmptyStrings As Boolean = False, _
  Optional StrCustomRegex As String = VBA.vbNullString) As Variant

  Dim VData As Variant, vAux As Variant
  Dim i As Long, j As Long

  Let ArrayGetDataColonSeparatedFromString = Empty

  On Error GoTo EH
  Let i = 0
  Let j = 0
  ReDim VData(0)
  If REGEX.isEmptyStringReg(StrVector) Then
    Let VData(0) = VBA.vbNullString
    Let ArrayGetDataColonSeparatedFromString = VData
    GoTo EH
  End If
  If StrValueType Then 'Strings
    Let vAux = VBA.Split(StrVector, ";")
    For i = 0 To UBound(vAux)
      If REGEX.isStringReg(vAux(i), StrCustomRegex) Then 'If Not REGEX.isNumberReg(vAux(i)) Then
        If StrIncludeEmptyStrings Then
          ReDim Preserve VData(j)
          Let VData(j) = vAux(i)
          Let j = j + 1
        Else
          ReDim Preserve VData(j)
          Let VData(j) = vAux(i)
          Let j = j + 1
        End If
      End If
    Next i
    If j = 0 Then GoTo EH
  Else 'Numbers
    If Not REGEX.isNumberEntireAndDecimalVectorColonSeparated(StrVector) Then
      Let VData(0) = VBA.vbNullString
      Let ArrayGetDataColonSeparatedFromString = VData
      GoTo EH
    End If
    Let vAux = VBA.Split(VBA.Replace(StrVector, ".", ","), ";")
    For i = 0 To UBound(vAux)
      If REGEX.isNumberReg(vAux(i)) And Not REGEX.isEmptyStringReg(vAux(i)) Then
        If StrIncludeZeroes Then
          ReDim Preserve VData(j)
          Let VData(j) = IIf(StrAbs, VBA.Abs(VBA.CDbl(vAux(i))), VBA.CDbl(vAux(i)))
          Let j = j + 1
        Else
          If VBA.Abs(VBA.CDbl(vAux(i))) > 0 Then
            ReDim Preserve VData(j)
            Let VData(j) = IIf(StrAbs, VBA.Abs(VBA.CDbl(vAux(i))), VBA.CDbl(vAux(i)))
            Let j = j + 1
          End If
        End If
      End If
    Next i
    If j = 0 Then GoTo EH
  End If

  Let StrCounter = j
  Let ArrayGetDataColonSeparatedFromString = VData

EH:
  Let VData = Empty
  Let vAux = Empty
  Call EHGLOBAL.ErrorHandlerRaise("RES::ArrayGetDataColonSeparatedFromString")

End Function

' ALERT!
'Public Function ArrayGetDataColonSeparatedStringFromVariants( _
'  ByVal RVarVector As Variant, _
'  ByRef RLngCounter As Long, _
'  Optional RBooValueType As Boolean = False, _
'  Optional RBooAbs As Boolean = False, _
'  Optional RBooIncludeZeroes As Boolean = False, _
'  Optional RBooIncludeEmptyStrings As Boolean = False, _
'  Optional RStrCustomRegex As String = VBA.vbNullString, _
'  Optional RBooRngRowsCols As Boolean = False) As String
'
'  'Extract all values numbers or strings from any variable as Range or Strings, each separated by semicolon
'  'Just works with Strings, Ranges and Numbers
'  Dim VecAux As Variant
'  Dim StrAux As String, StrResponse As String
'
'  Const sFN As String = "RES::ArrayGetDataColonSeparatedStringFromVariants"
'
'  On Error GoTo EH
'  Let ArrayGetDataColonSeparatedStringFromVariants = VBA.vbNullString
'
'  'Process
'  Let StrAux = VBA.TypeName(RVarVector)
'  Select Case StrAux
'    Case "Double", "Integer", "Long", "Single"
'      Let StrResponse = VBA.CStr(RVarVector)
'    Case "String"
'      Let VecAux = ArrayGetDataColonSeparatedFromString(RVarVector, RLngCounter, RBooValueType, RBooAbs, RBooIncludeZeroes, RBooIncludeEmptyStrings, RStrCustomRegex)
'      Let StrResponse = VBA.Join(VecAux, ";")
'    Case "Range"
'      Let StrResponse = ArrayGetDataAsStringFromRanges(RVarVector, RLngCounter, RBooValueType, RBooAbs, RBooIncludeZeroes, RBooIncludeEmptyStrings, RStrCustomRegex, RBooRngRowsCols)
'    Case Else: GoTo EH
'  End Select
'
'  'Return
'  Let ArrayGetDataColonSeparatedStringFromVariants = StrResponse
'
'EH:
'  Let VecAux = Empty
'  Call EHGLOBAL.ErrorHandlerRaise(sFN)
'
'End Function

Public Function ArrayGetDataSemicolonSeparatedSlashGroupedStringFromVariants( _
  ByVal RVarVector As Variant, _
  ByRef RLngCounter As Long, _
  Optional RBooValueType As Boolean = False, _
  Optional RBooAbs As Boolean = False, _
  Optional RBooIncludeZeroes As Boolean = False, _
  Optional RBooIncludeEmptyStrings As Boolean = False, _
  Optional RStrCustomRegex As String = VBA.vbNullString, _
  Optional RBooRngRowsCols As Boolean = False) As String

  'Extract all values numbers or strings from any variable as Range or Strings, each separated by semicolon
  'Just works with Strings, Ranges and Numbers
  Dim VecAux As Variant, VecAux1 As Variant, VecAux2 As Variant
  Dim StrAux As String, StrResponse As String
  Dim i As Long, j As Long

  Const sFN As String = "RES::ArrayGetDataSemicolonSeparatedSlashGroupedStringFromVariants"

  On Error GoTo EH
  Let ArrayGetDataSemicolonSeparatedSlashGroupedStringFromVariants = VBA.vbNullString

  'Process
  Let StrAux = VBA.TypeName(RVarVector)
  Select Case StrAux
    Case "Double", "Integer", "Long", "Single"
      Let StrResponse = VBA.CStr(RVarVector)

    Case "String"
      Let VecAux1 = VBA.Split(RVarVector, "/")
      ReDim VecAux2(0)
      Let j = 0
      For i = 0 To UBound(VecAux1)
        If Not REGEX.isEmptyStringReg(VBA.CStr(VecAux1(i))) Then
          Let VecAux = ArrayGetDataColonSeparatedFromString(VBA.CStr(VecAux1(i)), RLngCounter, RBooValueType, RBooAbs, RBooIncludeZeroes, RBooIncludeEmptyStrings, RStrCustomRegex)
          If RLngCounter > 0 Then
            ReDim Preserve VecAux2(j)
            Let VecAux2(j) = VBA.Join(VecAux, ";")
            Let j = j + 1
          End If
        End If
      Next i
      Let VecAux = Empty
      Let VecAux1 = Empty
      Let RLngCounter = j
      Let StrResponse = VBA.Join(VecAux2, "/")
      Let VecAux2 = Empty

    Case "Range"
      Dim RngBox As Range, RngArea As Range
      Set RngBox = RVarVector
      If Not RngBox Is Nothing Then
        ReDim VecAux(0)
        ReDim VecAux1(0)
        Let j = 0
        Let i = 0
        If RngBox.Areas.Count = 1 Then
          For Each RngArea In RngBox.Columns
            ReDim Preserve VecAux(j)
            Let VecAux(j) = ArrayGetDataAsStringFromRanges(RngArea, RLngCounter, RBooValueType, RBooAbs, RBooIncludeZeroes, RBooIncludeEmptyStrings, RStrCustomRegex, RBooRngRowsCols)
            If RLngCounter > 0 Then
              ReDim Preserve VecAux1(i)
              Let VecAux1(i) = VecAux(j)
              Let i = i + 1
            End If
            Let j = j + 1
          Next RngArea
          Let StrResponse = VBA.Join(VecAux1, "/")
        Else
          For Each RngArea In RngBox.Areas
            ReDim Preserve VecAux(j)
            Let VecAux(j) = ArrayGetDataAsStringFromRanges(RngArea, RLngCounter, RBooValueType, RBooAbs, RBooIncludeZeroes, RBooIncludeEmptyStrings, RStrCustomRegex, RBooRngRowsCols)
            If RLngCounter > 0 Then
              ReDim Preserve VecAux1(i)
              Let VecAux1(i) = VecAux(j)
              Let i = i + 1
            End If
            Let j = j + 1
          Next RngArea
          Let StrResponse = VBA.Join(VecAux1, "/")
        End If
        Let RLngCounter = i
        Set RngArea = Nothing
        Set RngBox = Nothing
      End If

    Case Else: GoTo EH
  End Select

  'Return
  Let ArrayGetDataSemicolonSeparatedSlashGroupedStringFromVariants = StrResponse

EH:
  Let VecAux = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

' ALERT!
'Public Function ArrayGetDataArrayFromVariants( _
'  ByVal RVarVector As Variant, _
'  ByRef RLngCounter As Long, _
'  Optional RBooValueType As Boolean = False, _
'  Optional RBooAbs As Boolean = False, _
'  Optional RBooIncludeZeroes As Boolean = False, _
'  Optional RBooIncludeEmptyStrings As Boolean = False, _
'  Optional RStrCustomRegex As String = VBA.vbNullString, _
'  Optional RBooRngRowsCols As Boolean = False) As Variant
'
'  'Extract all values numbers or strings from any variable as Range or Strings, each separated by semicolon
'  'Just works with Strings, Ranges and Numbers
'  Dim i As Long
'  Dim VecAux As Variant, VecAux1 As Variant
'  Dim StrAux As String
'
'  Const sFN As String = "RES::ArrayGetDataArrayFromVariants"
'
'  On Error GoTo EH
'  Let ArrayGetDataArrayFromVariants = Array(0)
'  Let RLngCounter = 0
'
'  'Process
'  Let StrAux = VBA.TypeName(RVarVector)
'  Select Case StrAux
'    Case "Double", "Integer", "Long", "Single"
'      If RBooIncludeZeroes Then
'        Let StrAux = VBA.CStr(IIf(RBooAbs, VBA.Abs(RVarVector), RVarVector)): Let RLngCounter = 1
'      Else
'        Let StrAux = VBA.CStr(IIf(RBooAbs, VBA.Abs(RVarVector), RVarVector)): Let RLngCounter = IIf(RBooIncludeZeroes And RVarVector = 0, 0, 1)
'      End If
'    Case "String"
'      Let VecAux = ArrayGetDataColonSeparatedFromString(RVarVector, RLngCounter, RBooValueType, RBooAbs, RBooIncludeZeroes, RBooIncludeEmptyStrings, RStrCustomRegex)
'      Let StrAux = VBA.Join(VecAux, ";") 'TODO: CORREGIR EN CASO DE CADENA VACÍA
'    Case "Range"
'      Let StrAux = ArrayGetDataAsStringFromRanges(RVarVector, RLngCounter, RBooValueType, RBooAbs, RBooIncludeZeroes, RBooIncludeEmptyStrings, RStrCustomRegex, RBooRngRowsCols)
'    Case Else: GoTo EH
'  End Select
'
'  Let VecAux = VBA.Split(StrAux, ";") 'strings
'  If Not RBooValueType Then 'numbers
'    If RLngCounter = 0 Then
'      Let VecAux1 = Array(0)
'    Else
'      ReDim VecAux1(UBound(VecAux))
'      For i = 0 To UBound(VecAux)
'        Let VecAux1(i) = VBA.CDbl(VecAux(i))
'      Next i
'    End If
'    Let VecAux = Empty
'    Let VecAux = VecAux1
'    Let VecAux1 = Empty
'  End If
'
'  'Return
'  Let ArrayGetDataArrayFromVariants = VecAux
'
'EH:
'  Let VecAux = Empty
'  Debug.Print VBA.Err.Description
'  Call EHGLOBAL.ErrorHandlerRaise(sFN)
'
'End Function

Public Function ArrayGetDataXCharSeparatedStringFromRanges( _
  ByVal RRngVector As Range, _
  Optional RXChar As String = ";", _
  Optional RBooIncludeEmptyStrings As Boolean = False, _
  Optional RRngRC As EQLRES_ENU_RNGRC = eqlResRngCol) As String

  'Extract all values numbers, strings or dates from any variable like Ranges or Strings, each separated by semicolon
  Dim RngArea As Range
  Dim VecAux As Variant
  Dim StrAux As String, StrResponse As String, StrAux0 As String
  Dim i As Long, j As Long

  Const sFN As String = "RES::ArrayGetDataXCharSeparatedStringFromRanges"

  On Error GoTo EH
  Let ArrayGetDataXCharSeparatedStringFromRanges = VBA.vbNullString

  'Process
  Let StrAux = VBA.TypeName(RRngVector)
  If RRngVector Is Nothing Then GoTo EH
  ReDim VecAux(0)
  Let j = 0
  For Each RngArea In RRngVector.Areas
    If RRngRC = eqlResRngCol Then 'Cols
      For i = 1 To RngArea.Columns.Count
        ReDim Preserve VecAux(j)
        Let VecAux(j) = Application.WorksheetFunction.TextJoin(RXChar, RBooIncludeEmptyStrings, RngArea.Columns.Item(i))
        Let j = j + 1
      Next i
    Else 'Rows
      ReDim Preserve VecAux(j)
      Let VecAux(j) = Application.WorksheetFunction.TextJoin(RXChar, RBooIncludeEmptyStrings, RngArea)
      Let j = j + 1
    End If
  Next RngArea
  Set RngArea = Nothing
  Let StrResponse = VBA.Join(VecAux, RXChar)

  'Return
  Let ArrayGetDataXCharSeparatedStringFromRanges = StrResponse

EH:
  Let VecAux = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Public Function ArrayGetDataXCharSeparatedStringFromVariantsAsString( _
  ByVal RVarVector As Variant, _
  ByRef RLngCounter As Long, _
  Optional RXChar As String = ";", _
  Optional RValueType As EQLRES_ENU_VALUETYPE = eqlResNumbers, _
  Optional RBooAbs As Boolean = False, _
  Optional RBooIncludeZeroes As Boolean = False, _
  Optional RBooIncludeEmptyStrings As Boolean = False, _
  Optional RStrCustomRegex As String = VBA.vbNullString, _
  Optional RRngRC As EQLRES_ENU_RNGRC = eqlResRngCol, _
  Optional RDblTop As Double = -1, _
  Optional RBooUniques As Boolean = False) As String

  Dim Rng As Range
  Dim VData As Variant, vAux As Variant
  Dim StrData As String
  Dim i As Long, j As Long

  Const sFN = "RES::ArrayGetDataXCharSeparatedStringFromVariantsAsString"

  On Error GoTo EH
  Let ArrayGetDataXCharSeparatedStringFromVariantsAsString = VBA.vbNullString
  Let RLngCounter = 0

  'Get the target string
  If VBA.TypeName(RVarVector) = "Range" Then
    Set Rng = RVarVector
    Let StrData = ArrayGetDataXCharSeparatedStringFromRanges(Rng, RXChar, RBooIncludeEmptyStrings, RRngRC)
    Set Rng = Nothing
  ElseIf VBA.TypeName(RVarVector) = "String" Or VBA.IsNumeric(RVarVector) Then
    Let StrData = RVarVector
  Else
    GoTo EH
  End If

  'Filter data
  Let i = 0
  Let j = 0
  ReDim VData(0)
  If RValueType = eqlResStrings Then 'Strings

    Let vAux = VBA.Split(StrData, RXChar)
    For i = 0 To UBound(vAux)
      If REGEX.isStringReg(vAux(i), RStrCustomRegex) Then
        If REGEX.isEmptyStringReg(VBA.CStr(vAux(i))) Then
          If RBooIncludeEmptyStrings Then
            ReDim Preserve VData(j)
            Let VData(j) = vAux(i)
            Let j = j + 1
          End If
        Else
          ReDim Preserve VData(j)
          Let VData(j) = vAux(i)
          Let j = j + 1
        End If
      End If
    Next i

  ' *********
  ElseIf RValueType = eqlResNotNumbers Then ' Strings not numbers
    Let vAux = VBA.Split(StrData, RXChar)
    For i = 0 To UBound(vAux)
      If Not VBA.IsNumeric(vAux(i)) Then
        If REGEX.isEmptyStringReg(VBA.CStr(vAux(i))) Then
          If RBooIncludeEmptyStrings Then
            ReDim Preserve VData(j)
            Let VData(j) = vAux(i)
            Let j = j + 1
          End If
        Else
          ReDim Preserve VData(j)
          Let VData(j) = vAux(i)
          Let j = j + 1
        End If
      End If
    Next i
  ' *********

  ElseIf RValueType = eqlResNumbers Then 'Numbers

    If REGEX.isNumberEntireAndDecimalVectorColonSeparated(StrData) Then
      Let vAux = VBA.Split(VBA.Replace(StrData, ".", ","), RXChar)
      For i = 0 To UBound(vAux)
        If REGEX.isNumberReg(vAux(i)) And VBA.IsNumeric(vAux(i)) Then
          If VBA.CDbl(vAux(i)) = 0 Then ' Filter by zeroes
            If RBooIncludeZeroes Then
              ReDim Preserve VData(j)
              Let VData(j) = IIf(RBooAbs, VBA.Abs(VBA.CDbl(VBA.Replace(VBA.CStr(vAux(i)), ".", ","))), VBA.CDbl(VBA.Replace(VBA.CStr(vAux(i)), ".", ",")))
              Let j = j + 1
            End If
          Else
            If RDblTop >= 0 Then ' Filter by top number
              If VBA.CDbl(vAux(i)) <= RDblTop Then
                ReDim Preserve VData(j)
                Let VData(j) = IIf(RBooAbs, VBA.Abs(VBA.CDbl(VBA.Replace(VBA.CStr(vAux(i)), ".", ","))), VBA.CDbl(VBA.Replace(VBA.CStr(vAux(i)), ".", ",")))
                Let j = j + 1
              End If
            Else ' Normal numbers
              ReDim Preserve VData(j)
              Let VData(j) = IIf(RBooAbs, VBA.Abs(VBA.CDbl(VBA.Replace(VBA.CStr(vAux(i)), ".", ","))), VBA.CDbl(VBA.Replace(VBA.CStr(vAux(i)), ".", ",")))
              Let j = j + 1
            End If
          End If
        End If
      Next i
    End If

  Else 'Dates

    Let vAux = VBA.Split(StrData, RXChar)
    For j = 0 To UBound(vAux)
      If VBA.IsDate(vAux(j)) Then
        ReDim Preserve VData(j)
        Let VData(j) = vAux(j)
      End If
    Next j

  End If
  If j = 0 Then GoTo EH

  If RBooUniques Then Let VData = Me.ArrayUniques(VData)

  Let RLngCounter = (UBound(VData) - LBound(VData)) + 1 'j
  Let ArrayGetDataXCharSeparatedStringFromVariantsAsString = VBA.Join(VData, RXChar)

EH:
  Let VData = Empty
  Let vAux = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Public Function ArrayGetDataXCharSeparatedStringFromVariantsAsVector( _
  ByVal RVarVector As Variant, _
  ByRef RLngCounter As Long, _
  Optional RXChar As String = ";", _
  Optional RValueType As EQLRES_ENU_VALUETYPE = eqlResNumbers, _
  Optional RBooAbs As Boolean = False, _
  Optional RBooIncludeZeroes As Boolean = False, _
  Optional RBooIncludeEmptyStrings As Boolean = False, _
  Optional RStrCustomRegex As String = VBA.vbNullString, _
  Optional RRngRC As EQLRES_ENU_RNGRC = eqlResRngCol, _
  Optional RValueAs As EQLRES_ENU_VALUETYPE = eqlResNumbers, _
  Optional RStrDateFormat As String = VBA.vbNullString, _
  Optional RDblTop As Double = -1, _
  Optional RBooUniques As Boolean = False) As Variant

  Dim VData As Variant, VData1 As Variant
  Dim StrAux0 As String
  Dim i As Long

  Const sFN = "RES::ArrayGetDataXCharSeparatedStringFromVariantsAsVector"

  On Error GoTo EH

  Let StrAux0 = ArrayGetDataXCharSeparatedStringFromVariantsAsString(RVarVector, RLngCounter, RXChar, RValueType, RBooAbs, RBooIncludeZeroes, RBooIncludeEmptyStrings, RStrCustomRegex, RRngRC, RDblTop)

  If RLngCounter = 0 Then
    Let ArrayGetDataXCharSeparatedStringFromVariantsAsVector = VData
  Else
    Let VData1 = VBA.Split(StrAux0, RXChar) ' Strings
    If RValueAs = eqlResStrings Or RValueAs = eqlResNotNumbers Then
      Let VData = VData1
    ElseIf RValueAs = eqlResNumbers Then ' Numbers
      ReDim VData(UBound(VData1))
      For i = LBound(VData) To UBound(VData)
        Let VData(i) = VBA.CDbl(VData1(i))
      Next i
    ElseIf RValueAs = eqlResDates Then ' Dates
      ReDim VData(UBound(VData1))
      For i = LBound(VData) To UBound(VData)
        Let VData(i) = VBA.CDate((VBA.Format(VBA.CDate(VData1(i)), RStrDateFormat)))
      Next i
    End If
  End If

  ' Return
  Let ArrayGetDataXCharSeparatedStringFromVariantsAsVector = VData

EH:
  Let VData = Empty
  Let VData1 = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Public Function ArrayHasValue( _
  ByVal AVector As Variant, _
  ByVal ANeedle As Variant, _
  Optional AStrOrNum As Boolean = False, _
  Optional ALCase As Boolean = False) As Boolean

  Dim i As Long
  Dim sAux As String

  On Error GoTo EH
  Let ArrayHasValue = False
  For i = LBound(AVector) To UBound(AVector)
    If AStrOrNum Then Let sAux = VBA.CDbl(AVector(i))
    If Not AStrOrNum Then Let sAux = IIf(ALCase, VBA.LCase(AVector(i)), AVector(i))
    If ANeedle = sAux Then Let ArrayHasValue = True: Exit For
  Next i

EH:
  Call EHGLOBAL.ErrorHandlerRaise("RES::ArrayHasValue")

End Function

Public Function ArrayHasValue2D( _
  ByVal AVector As Variant, _
  ByVal ANeedle As Variant, _
  Optional ALngColToSearch As Long = -1) As Boolean

  Dim i As Long, j As Long
  Dim StrAux0 As String, StrAux1 As String

  ' Searching by SEARCHABLE data in 1D or 2D vector, return True if needle is founded or false if doesn't.
  On Error GoTo EH
  Let ArrayHasValue2D = -1

  If VBA.IsArray(AVector) Then
    Let i = Me.ArrayDimension(AVector)
    If i = 1 Then ' 1D vectors
      Let StrAux1 = VBA.LCase(VBA.CStr(ANeedle))
      For i = LBound(AVector) To UBound(AVector)
        Let StrAux0 = VBA.LCase(VBA.CStr(AVector(i)))
        Let ArrayHasValue2D = (StrAux0 = StrAux1)
        If ArrayHasValue2D Then Exit For
      Next i

    ElseIf i = 2 Then ' 2D vectors
      Let StrAux1 = VBA.LCase(VBA.CStr(ANeedle))
      If ALngColToSearch < LBound(AVector, 2) Or ALngColToSearch > UBound(AVector, 2) Then ' Searching in each col (by default)
        For j = LBound(AVector, 2) To UBound(AVector, 2)
          For i = LBound(AVector) To UBound(AVector)
            Let StrAux0 = VBA.LCase(VBA.CStr(AVector(i, j)))
            Let ArrayHasValue2D = (StrAux0 = StrAux1)
            If ArrayHasValue2D Then Exit For
          Next i
          If ArrayHasValue2D Then Exit For
        Next j
      Else ' Search in specified col
        For i = LBound(AVector) To UBound(AVector)
          Let StrAux0 = VBA.LCase(VBA.CStr(AVector(i, ALngColToSearch)))
          Let ArrayHasValue2D = (StrAux0 = StrAux1)
          If ArrayHasValue2D Then Exit For
        Next i
      End If

    End If
  End If

EH:
  Call EHGLOBAL.ErrorHandlerRaise("RES::ArrayHasValue2D")

End Function

Public Function ArrayHasValueN( _
  ByVal AVector As Variant, _
  ByVal ANeedle As Variant, _
  Optional AStrOrNum As Boolean = False, _
  Optional ALCase As Boolean = False) As Long

  Dim i As Long
  Dim sAux As Variant

  On Error GoTo EH
  Let ArrayHasValueN = 0
  For i = LBound(AVector) To UBound(AVector)
    If AStrOrNum Then Let sAux = VBA.CDbl(AVector(i)) 'Num
    If Not AStrOrNum Then Let sAux = IIf(ALCase, VBA.LCase(AVector(i)), AVector(i)) 'Str
    If ANeedle = sAux Then Let ArrayHasValueN = ArrayHasValueN + 1
  Next i

EH:
  Call EHGLOBAL.ErrorHandlerRaise("RES::ArrayHasValueN")

End Function

Public Function ArrayHasValueN2D( _
  ByVal AVector As Variant, _
  ByVal ANeedle As Variant, _
  Optional ALngColToSearch As Long = -1) As Long

  Dim i As Long, j As Long, k As Long
  Dim StrAux0 As String, StrAux1 As String

  ' Obtains number of needle occurences in vector searching by SEARCHABLE data in 1D or 2D vector, return number of occurences founded or -1 if it doesn't found.
  On Error GoTo EH
  Let ArrayHasValueN2D = -1
  Let k = 0

  If VBA.IsArray(AVector) Then
    Let i = Me.ArrayDimension(AVector)
    If i = 1 Then ' 1D vectors
      Let StrAux1 = VBA.LCase(VBA.CStr(ANeedle))
      For i = LBound(AVector) To UBound(AVector)
        Let StrAux0 = VBA.LCase(VBA.CStr(AVector(i)))
        Let k = k + IIf(StrAux0 = StrAux1, 1, 0)
      Next i

    ElseIf i = 2 Then ' 2D vectors
      Let StrAux1 = VBA.LCase(VBA.CStr(ANeedle))
      If ALngColToSearch < LBound(AVector, 2) Or ALngColToSearch > UBound(AVector, 2) Then ' Searching in each col (by default)
        For j = LBound(AVector, 2) To UBound(AVector, 2)
          For i = LBound(AVector) To UBound(AVector)
            Let StrAux0 = VBA.LCase(VBA.CStr(AVector(i, j)))
            Let k = k + IIf(StrAux0 = StrAux1, 1, 0)
          Next i
        Next j
      Else ' Search in specified col
        For i = LBound(AVector) To UBound(AVector)
          Let StrAux0 = VBA.LCase(VBA.CStr(AVector(i, ALngColToSearch)))
          Let k = k + IIf(StrAux0 = StrAux1, 1, 0)
        Next i
      End If

    End If
  End If

  ' Return
  Let ArrayHasValueN2D = IIf(k > 0, k, ArrayHasValueN2D)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("RES::ArrayHasValueN2D")

End Function

Public Function ArrayGetValuePosition( _
  ByVal AVector As Variant, _
  ByVal ANeedle As Variant, _
  Optional ALngColToSearch As Long = -1) As Long

  Dim i As Long, j As Long
  Dim StrAux0 As String, StrAux1 As String

  ' Search a searchable data in 1D or 2D vector, return index of position value founded or -1 if it doesn't found.
  On Error GoTo EH
  Let ArrayGetValuePosition = -1

  If VBA.IsArray(AVector) Then
    Let i = Me.ArrayDimension(AVector)
    If i = 1 Then
      Let StrAux1 = VBA.LCase(VBA.CStr(ANeedle))
      For i = LBound(AVector) To UBound(AVector)
        Let StrAux0 = VBA.LCase(VBA.CStr(AVector(i)))
        If StrAux0 = StrAux1 Then Let ArrayGetValuePosition = i: Exit For
      Next i

    ElseIf i = 2 Then
      If ALngColToSearch < LBound(AVector, 2) Or ALngColToSearch > UBound(AVector, 2) Then ' Search in every col
        Let StrAux1 = VBA.LCase(VBA.CStr(ANeedle))
        For j = LBound(AVector, 2) To UBound(AVector, 2)
          For i = LBound(AVector) To UBound(AVector)
            Let StrAux0 = VBA.LCase(VBA.CStr(AVector(i, j)))
            If StrAux0 = StrAux1 Then Let ArrayGetValuePosition = i: Exit For
          Next i
          If ArrayGetValuePosition > -1 Then Exit For
        Next j
      Else ' Search in specified col
        Let StrAux1 = VBA.LCase(VBA.CStr(ANeedle))
        For i = LBound(AVector) To UBound(AVector)
          Let StrAux0 = VBA.LCase(VBA.CStr(AVector(i, ALngColToSearch)))
          If StrAux0 = StrAux1 Then Let ArrayGetValuePosition = i: Exit For
        Next i
      End If

    End If
  End If

EH:
  Call EHGLOBAL.ErrorHandlerRaise("RES::ArrayGetValuePosition")

End Function

Public Function ArrayHasRepeatedElements(ByVal Vec As Variant) As Boolean

  Dim StrNeedle As String, StrCurrent As String
  Dim i As Long, j As Long

  ' Counts repeated elements casted as strings in array 1D
  Const sFN As String = "RES::ArrayHasRepeatedElements"

  On Error GoTo EH
  Let ArrayHasRepeatedElements = False

  If Not VBA.IsArray(Vec) Then Call VBA.Err.Raise(CUSTOM_ERROR_RES, sFN, "¡Se requiere un array!")
  If (UBound(Vec) < 0) Then Call VBA.Err.Raise(CUSTOM_ERROR_RES, sFN, "¡Se requiere un array!")
  If UBound(Vec) = LBound(Vec) Then GoTo EH

  Let i = LBound(Vec)
  Let j = LBound(Vec) + 1

  Do While (ArrayHasRepeatedElements = False)
    Let StrNeedle = VBA.CStr(Vec(i))
    Let StrCurrent = VBA.CStr(Vec(j))
    Let ArrayHasRepeatedElements = (StrNeedle = StrCurrent)
    Let i = i + IIf(j = UBound(Vec), 1, 0)
    Let j = IIf(j = UBound(Vec), i + 1, j + 1)
    If i = UBound(Vec) Then Exit Do
  Loop

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Public Function ArrayShift( _
  ByRef RVarSource As Variant, _
  Optional RReturnSrc As Boolean = False, _
  Optional RByCols As Boolean = False) As Variant

  ' Deletes the first element of any 1D array. In 2D arrays can delete the first row or the first col _
    althrough, uses the VecSource index base and the returned array will have base 0. _
    Arrays that has just one element are returned as empty

  Dim xIndex As Long

  Const sFN As String = "RES::ArrayShift"

  On Error GoTo EH
  If VBA.IsArray(RVarSource) Then
    Let xIndex = Me.ArrayDimension(RVarSource)
    If xIndex = 1 Then
      Let xIndex = LBound(RVarSource)
    ElseIf xIndex = 2 Then
      If RByCols Then
        Let xIndex = LBound(RVarSource, 2)
      Else
        Let xIndex = LBound(RVarSource)
      End If
    Else
      Let xIndex = -1
    End If
  End If

  Let ArrayShift = Me.ArrayDelIndex(RVarSource, xIndex, RReturnSrc, RByCols)

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Public Function ArrayPop( _
  ByRef RVarSource As Variant, _
  Optional RReturnSrc As Boolean = False, _
  Optional RByCols As Boolean = False) As Variant

  ' Deletes the last element of any 1D array. In 2D arrays can delete the last row or the last col _
    althrough, uses the VecSource index base and the returned array will have base 0. _
    Arrays that has just one element are returned as empty
  Dim xIndex As Long

  Const sFN As String = "RES::ArrayPop"

  On Error GoTo EH

  If VBA.IsArray(RVarSource) Then
    Let xIndex = Me.ArrayDimension(RVarSource)
    If xIndex = 1 Then
      Let xIndex = UBound(RVarSource)
    ElseIf xIndex = 2 Then
      If RByCols Then
        Let xIndex = UBound(RVarSource, 2)
      Else
        Let xIndex = UBound(RVarSource)
      End If
    Else
      Let xIndex = -1
    End If
  End If

  Let ArrayPop = Me.ArrayDelIndex(RVarSource, xIndex, RReturnSrc, RByCols)

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Public Function ArraySlice( _
  ByRef RVarSource As Variant, _
  ByRef RLngIndex As Long, _
  Optional RReturnSrc As Boolean = False, _
  Optional RByCols As Boolean = False, _
  Optional RBooGetLeftSide As Boolean = False) As Variant

  ' Slices the any array (1D or 2D) at index x and return the first part as array _
    and the VecSource sliced, both has base 0.
  Dim VecAux0 As Variant, VecAux1 As Variant
  Dim i As Long, j As Long, k As Long, l As Long, m As Long, xCol As Long, xRow  As Long, LngIndex As Long
  Dim EnuDimensionSrc As EQLRES_ENU_DIMENSIONARRAY

  Const sFN As String = "RES::ArraySlice"

  On Error GoTo EH

  If Not VBA.IsArray(RVarSource) Then Call VBA.Err.Raise(CUSTOM_ERROR_RES, sFN, "¡Se requiere un array!")
  Let i = Me.ArrayDimension(RVarSource)
  Let EnuDimensionSrc = IIf(i = 1, eqlRes1D, IIf(i = 2, eqlRes2D, eqlResNoArray))
  If i > 2 Then GoTo EH
  If EnuDimensionSrc = eqlResNoArray Then Call VBA.Err.Raise(CUSTOM_ERROR_RES, sFN, "¡Se requiere un array!")

  Let j = 0
  Let k = 0
  Let l = 0
  Let m = 0
  Let RLngIndex = VBA.Abs(RLngIndex)
  Select Case EnuDimensionSrc

    Case eqlRes1D
      If (LBound(RVarSource) = UBound(RVarSource)) Or RLngIndex < LBound(RVarSource) Or RLngIndex > UBound(RVarSource) Then
        ReDim VecAux0(0)
      ElseIf RLngIndex = UBound(RVarSource) Then ' Same total bound and slice index - AUGMENTED
        Let VecAux0 = RVarSource
      Else
        ReDim VecAux0(0)
        ReDim VecAux1(0)
        For i = LBound(RVarSource) To UBound(RVarSource)
          If i <= RLngIndex Then
            ReDim Preserve VecAux0(j)
            Let VecAux0(j) = RVarSource(i)
            Let j = j + 1
          Else
            ReDim Preserve VecAux1(k)
            Let VecAux1(k) = RVarSource(i)
            Let k = k + 1
          End If
        Next i
      End If

    Case eqlRes2D

      If RByCols Then 'Cols

        If (LBound(RVarSource, 2) = UBound(RVarSource, 2)) Or RLngIndex < LBound(RVarSource, 2) Or RLngIndex > UBound(RVarSource, 2) Then
          Let ArraySlice = Empty
          GoTo EH
        ElseIf RLngIndex = UBound(RVarSource, 2) Then ' Same total bound and slice index - AUGMENTED
          Let VecAux0 = RVarSource
        Else
          Let xCol = (UBound(RVarSource, 2) - LBound(RVarSource, 2))
          Let xRow = UBound(RVarSource) - LBound(RVarSource)
          Let LngIndex = (RLngIndex - LBound(RVarSource, 2))
          ReDim VecAux0(0 To xRow, 0 To LngIndex)
          ReDim VecAux1(0 To xRow, 0 To (xCol - (LngIndex + 1)))
          For i = LBound(RVarSource) To UBound(RVarSource)
            For j = LBound(RVarSource, 2) To UBound(RVarSource, 2)
              If j <= RLngIndex Then
                Let VecAux0(k, l) = RVarSource(i, j)
              Else
                Let VecAux1(k, m) = RVarSource(i, j)
              End If
              Let l = l + IIf(j <= RLngIndex, 1, 0)
              Let m = m + IIf(j <= RLngIndex, 0, 1)
            Next j
            Let k = k + 1
            Let l = 0
            Let m = 0
          Next i
        End If

      Else 'Rows

        If (LBound(RVarSource) = UBound(RVarSource)) Or RLngIndex < LBound(RVarSource) Or RLngIndex > UBound(RVarSource) Then
          Let ArraySlice = Empty
          GoTo EH
        ElseIf RLngIndex = UBound(RVarSource) Then ' Same total bound and slice index - AUGMENTED
          Let VecAux0 = RVarSource
        Else
          Let xCol = UBound(RVarSource, 2) - LBound(RVarSource, 2)
          Let xRow = (UBound(RVarSource) - LBound(RVarSource))
          Let LngIndex = (RLngIndex - LBound(RVarSource, 2))
          ReDim VecAux0(0 To LngIndex, 0 To xCol)
          ReDim VecAux1(0 To (xRow - (LngIndex + 1)), 0 To xCol)
          For i = LBound(RVarSource) To UBound(RVarSource)
            For j = LBound(RVarSource, 2) To UBound(RVarSource, 2)
              If i <= RLngIndex Then
                Let VecAux0(l, k) = RVarSource(i, j)
              Else
                Let VecAux1(m, k) = RVarSource(i, j)
              End If
              Let k = k + 1
            Next j
            Let l = l + IIf(i <= RLngIndex, 1, 0)
            Let m = m + IIf(i <= RLngIndex, 0, 1)
            Let k = 0
          Next i
        End If

      End If

  End Select

  ' Return
  If RBooGetLeftSide Then
    Let ArraySlice = VecAux0
  Else
    Let ArraySlice = VecAux1
  End If

  If Not RReturnSrc Then
    Let RVarSource = Empty
    If RBooGetLeftSide Then
      Let RVarSource = VecAux1
    Else
      Let RVarSource = VecAux0
    End If
    Let VecAux0 = Empty
    Let VecAux1 = Empty
  End If

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Public Function ArrayAddAtFirst( _
  ByVal RVarSource As Variant, _
  ByVal RWithThisValue As Variant, _
  Optional RByCols As Boolean = False) As Variant

  ' Inserts one element at first position of any 1D array. In 2D arrays, the first row or the first col _
    althrough, uses the VecSource index base and the returned array will have base 0. _
    If RWithThisValue is some valid value, this value is used to fill the new element space.
  Dim xIndex As Long

  Const sFN As String = "RES::ArrayAddAtFirst"

  On Error GoTo EH
  If VBA.IsArray(RVarSource) Then
    Let xIndex = Me.ArrayDimension(RVarSource)
    If xIndex = 1 Then
      Let xIndex = LBound(RVarSource)
    ElseIf xIndex = 2 Then
      If RByCols Then
        Let xIndex = LBound(RVarSource, 2)
      Else
        Let xIndex = LBound(RVarSource)
      End If
    Else
      Let xIndex = -1
    End If
  End If

  Let ArrayAddAtFirst = ArrayAddElement(RVarSource, RWithThisValue, xIndex, RByCols)

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Public Function ArrayAddAtLast( _
  ByVal RVarSource As Variant, _
  ByVal RWithThisValue As Variant, _
  Optional RByCols As Boolean = False) As Variant

  ' Inserts one element at last position of any 1D array. In 2D arrays, the last row or the last col _
    althrough, uses the VecSource index base and the returned array will have base 0. _
    If RWithThisValue is some valid value, this value is used to fill the new element space.
  Dim xIndex As Long

  Const sFN As String = "RES::ArrayAddAtLast"

  On Error GoTo EH
  If VBA.IsArray(RVarSource) Then
    Let xIndex = Me.ArrayDimension(RVarSource)
    If xIndex = 1 Then
      Let xIndex = UBound(RVarSource) + 1
    ElseIf xIndex = 2 Then
      If RByCols Then
        Let xIndex = UBound(RVarSource, 2) + 1
      Else
        Let xIndex = UBound(RVarSource) + 1
      End If
    Else
      Let xIndex = -1
    End If
  End If

  Let ArrayAddAtLast = ArrayAddElement(RVarSource, RWithThisValue, xIndex, RByCols)

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Public Function ArrayDelIndex( _
  ByRef RVarSource As Variant, _
  ByVal RLngIndex As Long, _
  Optional RReturnSrc As Boolean = False, _
  Optional RByCols As Boolean = False) As Variant

  ' Deletes the last element of any 1D array. In 2D arrays can delete the last row or the last col _
    althrough, uses the VecSource index base and the returned array will have base 0. _
    Arrays that has just one element are returned as empty

  Dim VecAux0 As Variant, VecAux1 As Variant
  Dim i As Long, j As Long, k As Long, l As Long, m As Long, xCol As Long, xRow  As Long
  Dim EnuDimensionSrc As EQLRES_ENU_DIMENSIONARRAY

  Const sFN As String = "RES::ArrayDelIndex"

  On Error GoTo EH

  If Not VBA.IsArray(RVarSource) Then Call VBA.Err.Raise(CUSTOM_ERROR_RES, sFN, "¡Se requiere un array!")
  Let i = Me.ArrayDimension(RVarSource)
  Let EnuDimensionSrc = IIf(i = 1, eqlRes1D, IIf(i = 2, eqlRes2D, eqlResNoArray))
  If i > 2 Then GoTo EH
  If EnuDimensionSrc = eqlResNoArray Then Call VBA.Err.Raise(CUSTOM_ERROR_RES, sFN, "¡Se requiere un array!")

  Let j = 0
  Let k = 0
  Let l = 0
  Let m = 0
  Select Case EnuDimensionSrc
    Case eqlRes1D

      If (LBound(RVarSource) = UBound(RVarSource)) Then
        ReDim VecAux1(0 To 0)
        Let VecAux1(0, 0) = RVarSource(UBound(RVarSource))
        'ReDim VecAux0(0)
      Else
        ReDim VecAux0(0)
        ReDim VecAux1(0)
        For i = LBound(RVarSource) To UBound(RVarSource)
          If i = RLngIndex Then
            Let VecAux1(0) = RVarSource(i)
          Else
            ReDim Preserve VecAux0(j)
            Let VecAux0(j) = RVarSource(i)
            Let j = j + 1
          End If
        Next i
      End If

    Case eqlRes2D

      If RByCols Then 'Cols

        If (LBound(RVarSource, 2) = UBound(RVarSource, 2)) Then
          Let k = 0
          Let j = UBound(RVarSource, 2)
          ReDim VecAux1(0 To (UBound(RVarSource) - LBound(RVarSource)), 0 To 0)
          For i = LBound(RVarSource) To UBound(RVarSource)
            Let VecAux1(k, 0) = RVarSource(i, j)
            Let k = k + 1
          Next i
        Else
          Let xCol = (UBound(RVarSource, 2) - LBound(RVarSource, 2)) - 1
          Let xRow = UBound(RVarSource) - LBound(RVarSource)
          ReDim VecAux0(0 To xRow, 0 To xCol)
          ReDim VecAux1(0 To xRow, 0 To 0)
          For j = LBound(RVarSource, 2) To UBound(RVarSource, 2)
            For i = LBound(RVarSource) To UBound(RVarSource)
              If j = RLngIndex Then
                Let VecAux1(m, 0) = RVarSource(i, j)
              Else
                Let VecAux0(m, l) = RVarSource(i, j)
              End If
              Let m = m + 1
            Next i
            Let m = 0
            If Not j = RLngIndex Then Let l = l + 1
          Next j
        End If

      Else 'Rows

        If (LBound(RVarSource) = UBound(RVarSource)) Then
          Let k = 0
          Let i = UBound(RVarSource)
          ReDim VecAux1(0 To 0, 0 To (UBound(RVarSource, 2) - LBound(RVarSource, 2)))
          For j = LBound(RVarSource, 2) To UBound(RVarSource, 2)
            Let VecAux1(0, k) = RVarSource(i, j)
            Let k = k + 1
          Next j
        Else
          Let xCol = UBound(RVarSource, 2) - LBound(RVarSource, 2)
          Let xRow = (UBound(RVarSource) - LBound(RVarSource)) - 1
          ReDim VecAux0(0 To xRow, 0 To xCol)
          ReDim VecAux1(0 To 0, 0 To xCol)
          For i = LBound(RVarSource) To UBound(RVarSource)
            For j = LBound(RVarSource, 2) To UBound(RVarSource, 2)
              If i = RLngIndex Then
                Let VecAux1(0, k) = RVarSource(i, j)
              Else
                Let VecAux0(m, k) = RVarSource(i, j)
              End If
              Let k = k + 1
            Next j
            Let k = 0
            If Not i = RLngIndex Then Let m = m + 1
          Next i
        End If

      End If

  End Select

  Let ArrayDelIndex = VecAux1
  If Not RReturnSrc Then
    Let RVarSource = Empty
    Let RVarSource = VecAux0
    Let VecAux1 = Empty
  End If

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Public Function ArrayAddElement( _
  ByVal RVarSource As Variant, _
  ByVal RWithThisValue As Variant, _
  ByVal RLngPositionAt As Long, _
  Optional RByCols As Boolean = False) As Variant

  ' Adds a new element of any 1D array. In 2D arrays adds a row or col _
    at VecSource index base and the returned array will have base 0.

  Dim VecAux0 As Variant
  Dim i As Long, xRow As Long, xCol As Long
  Dim EnuDimension As EQLRES_ENU_DIMENSIONARRAY

  Const sFN As String = "RES::ArrayAddElement"

  On Error GoTo EH

  If Not VBA.IsArray(RVarSource) Then
    If VBA.IsArray(RWithThisValue) Then
      Let VecAux0 = RWithThisValue
    Else
      If VBA.TypeName(RWithThisValue) = "String" Or _
         VBA.IsNumeric(RWithThisValue) Or _
         VBA.IsDate(RWithThisValue) Or _
         VBA.IsEmpty(RWithThisValue) Then
        ReDim VecAux0(0)
        Let VecAux0(0) = RWithThisValue
      Else
        Call VBA.Err.Raise(CUSTOM_ERROR_RES, sFN, "¡Se requiere un array!")
      End If
    End If

  Else

    Let i = ArrayDimension(RVarSource)
    Let EnuDimension = IIf(i = 1, eqlRes1D, IIf(i = 2, eqlRes2D, eqlResNoArray))
    If EnuDimension = eqlResNoArray Then GoTo EH

    If EnuDimension = eqlRes1D Then '1D
      If VBA.IsArray(RWithThisValue) Then
        Let VecAux0 = RWithThisValue
      ElseIf VBA.TypeName(RWithThisValue) = "String" Or _
        VBA.IsNumeric(RWithThisValue) Or _
        VBA.IsDate(RWithThisValue) Or _
        VBA.IsEmpty(RWithThisValue) Then
        Let VecAux0 = Array(RWithThisValue)
      Else
        Call VBA.Err.Raise(CUSTOM_ERROR_RES, sFN, "¡Se requiere un valor u otro array para adicionar al array!")
      End If
    Else '2D
      If VBA.IsArray(RWithThisValue) Then
        Let VecAux0 = RWithThisValue
      ElseIf VBA.TypeName(RWithThisValue) = "String" Or _
        VBA.IsNumeric(RWithThisValue) Or _
        VBA.IsDate(RWithThisValue) Or _
        VBA.IsEmpty(RWithThisValue) Then
        Let xRow = (UBound(RVarSource) - LBound(RVarSource))
        Let xCol = (UBound(RVarSource, 2) - LBound(RVarSource, 2))
        Let i = 0
        ReDim VecAux0(IIf(RByCols, xRow, xCol))
        For i = LBound(VecAux0) To UBound(VecAux0)
          Let VecAux0(i) = RWithThisValue
        Next i
      Else
        Call VBA.Err.Raise(CUSTOM_ERROR_RES, sFN, "¡Se requiere un valor u otro array para adicionar al array!")
      End If
    End If

    Let VecAux0 = Me.ArrayConcat(RVarSource, VecAux0, RLngPositionAt, RByCols)

  End If

  Let ArrayAddElement = VecAux0

EH:
  Let VecAux0 = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN, vbInformation)

End Function

Public Function ArrayConcat( _
  ByVal RVarSrc As Variant, _
  ByVal RVarAdd As Variant, _
  Optional RPositionAt As Long = -1, _
  Optional RByCols As Boolean = False) As Variant

  Dim VecAux0 As Variant
  Dim BooOneDimension As Boolean
  Dim i As Long, j As Long, k As Long, l As Long, xRow As Long, xCol As Long
  Dim EnuDimension As EQLRES_ENU_DIMENSIONARRAY, EnuDimensionVecAdd As EQLRES_ENU_DIMENSIONARRAY

  Const sFN As String = "RES::ArrayConcat"

  On Error GoTo EH

  If Not VBA.IsArray(RVarSrc) And Not VBA.IsArray(RVarAdd) Then
    Call VBA.Err.Raise(CUSTOM_ERROR_RES, sFN, "¡Se requiere un array base y un valor u otro array para adicionar al array base!")
  ElseIf Not VBA.IsArray(RVarSrc) And VBA.IsArray(RVarAdd) Then
    Let VecAux0 = RVarAdd
  ElseIf VBA.IsArray(RVarSrc) And Not VBA.IsArray(RVarAdd) Then
    Let VecAux0 = RVarSrc
  Else
    Let i = Me.ArrayDimension(RVarSrc)
    Let EnuDimension = IIf(i = 1, eqlRes1D, IIf(i = 2, eqlRes2D, eqlResNoArray))
    Let i = Me.ArrayDimension(RVarAdd)
    Let EnuDimensionVecAdd = IIf(i = 1, eqlRes1D, IIf(i = 2, eqlRes2D, eqlResNoArray))
    If EnuDimension = eqlResNoArray Or EnuDimensionVecAdd = eqlResNoArray Then GoTo EH
    Let k = 0

    Select Case EnuDimension
      Case eqlRes1D
        Let xRow = (Me.ArrayLength(RVarSrc) + Me.ArrayLength(RVarAdd)) - 1
        ReDim VecAux0(xRow)
        For i = LBound(RVarSrc) To UBound(RVarSrc)
          Let VecAux0(k) = RVarSrc(i)
          Let k = k + 1
        Next i

        Select Case EnuDimensionVecAdd
          Case eqlRes1D

            If RPositionAt < LBound(RVarSrc) Or RPositionAt > UBound(RVarSrc) Then
              For i = LBound(RVarAdd) To UBound(RVarAdd)
                Let VecAux0(k) = RVarAdd(i)
                Let k = k + 1
              Next i
            Else
              Let k = 0
              For i = LBound(RVarSrc) To UBound(RVarSrc)
                If i = RPositionAt Then
                  For j = LBound(RVarAdd) To UBound(RVarAdd)
                    Let VecAux0(k) = RVarAdd(j)
                    Let k = k + 1
                  Next j
                End If
                Let VecAux0(k) = RVarSrc(i)
                Let k = k + 1
              Next i
            End If

          Case eqlRes2D

            If RPositionAt < LBound(RVarSrc) Or RPositionAt > UBound(RVarSrc) Then
              For j = LBound(RVarAdd, 2) To UBound(RVarAdd, 2)
                For i = LBound(RVarAdd) To UBound(RVarAdd)
                  Let VecAux0(k) = RVarAdd(i, j)
                  Let k = k + 1
                Next i
              Next j
            Else
              For i = LBound(RVarSrc) To UBound(RVarSrc)
                If i = RPositionAt Then
                  For l = LBound(RVarAdd, 2) To UBound(RVarAdd, 2)
                    For j = LBound(RVarAdd) To UBound(RVarAdd)
                      Let VecAux0(k) = RVarAdd(j, l)
                      Let k = k + 1
                    Next j
                  Next l
                End If
                Let VecAux0(k) = RVarSrc(i)
                Let k = k + 1
              Next i
            End If

        End Select

      Case eqlRes2D

        If RByCols Then 'By Cols
          If EnuDimensionVecAdd = eqlRes1D Then
            Let xCol = ((UBound(RVarSrc, 2) - LBound(RVarSrc, 2)) + 1) '+ Me.ArrayLength(RVarSrc)
          Else '2D
            Let xCol = ((UBound(RVarSrc, 2) - LBound(RVarSrc, 2)) + (UBound(RVarAdd, 2) - LBound(RVarAdd, 2))) + 1
          End If
          Let xRow = IIf((UBound(RVarSrc) - LBound(RVarSrc)) > (UBound(RVarAdd) - LBound(RVarAdd)), (UBound(RVarSrc) - LBound(RVarSrc)), (UBound(RVarAdd) - LBound(RVarAdd)))

          Let k = 0
          Let l = 0
          ReDim VecAux0(0 To xRow, 0 To xCol)
          If RPositionAt < LBound(RVarSrc, 2) Or RPositionAt > UBound(RVarSrc, 2) Then

            For j = LBound(RVarSrc, 2) To UBound(RVarSrc, 2)
              For i = LBound(RVarSrc) To UBound(RVarSrc)
                Let VecAux0(k, l) = RVarSrc(i, j)
                Let k = k + 1
              Next i
              Let k = 0
              Let l = l + 1
            Next j
            If EnuDimensionVecAdd = eqlRes1D Then
              For i = LBound(RVarAdd) To UBound(RVarAdd)
                Let VecAux0(k, UBound(VecAux0, 2)) = RVarAdd(i)
                Let k = k + 1
              Next i
            Else
              For j = LBound(RVarAdd, 2) To UBound(RVarAdd, 2)
                For i = LBound(RVarAdd) To UBound(RVarAdd)
                  Let VecAux0(k, l) = RVarAdd(i, j)
                  Let k = k + 1
                Next i
                Let k = 0
                Let l = l + 1
              Next j
            End If

          Else 'At Position n

            Let RPositionAt = RPositionAt - 1
            For j = LBound(RVarSrc, 2) To RPositionAt
              For i = LBound(RVarSrc) To UBound(RVarSrc)
                Let VecAux0(k, l) = RVarSrc(i, j)
                Let k = k + 1
              Next i
              Let k = 0
              Let l = l + 1
            Next j

            If EnuDimensionVecAdd = eqlRes1D Then
              Let k = 0
              For i = LBound(RVarAdd) To UBound(RVarAdd)
                Let VecAux0(k, l) = RVarAdd(i)
                Let k = k + 1
              Next i
              Let l = l + 1
            Else
              For j = LBound(RVarAdd, 2) To UBound(RVarAdd, 2)
                For i = LBound(RVarAdd) To UBound(RVarAdd)
                  Let VecAux0(k, l) = RVarAdd(i, j)
                  Let k = k + 1
                Next i
                Let k = 0
                Let l = l + 1
              Next j
            End If

            Let k = 0
            For j = (RPositionAt + 1) To UBound(RVarSrc, 2)
              For i = LBound(RVarSrc) To UBound(RVarSrc)
                Let VecAux0(k, l) = RVarSrc(i, j)
                Let k = k + 1
              Next i
              Let k = 0
              Let l = l + 1
            Next j
          End If

        Else 'By rows

          If EnuDimensionVecAdd = eqlRes1D Then '1D
            Let xRow = ((UBound(RVarSrc) - LBound(RVarSrc)) + 1)
            Let xCol = IIf((UBound(RVarSrc, 2) - LBound(RVarSrc, 2)) > (UBound(RVarAdd) - LBound(RVarAdd)), (UBound(RVarSrc, 2) - LBound(RVarSrc, 2)), (UBound(RVarAdd) - LBound(RVarAdd)))
          Else '2D
            Let xRow = ((UBound(RVarSrc) - LBound(RVarSrc)) + (UBound(RVarAdd) - LBound(RVarAdd))) + 1
            Let xCol = IIf((UBound(RVarSrc, 2) - LBound(RVarSrc, 2)) > (UBound(RVarAdd, 2) - LBound(RVarAdd, 2)), (UBound(RVarSrc, 2) - LBound(RVarSrc, 2)), (UBound(RVarAdd, 2) - LBound(RVarAdd, 2)))
          End If

          Let k = 0
          Let l = 0
          ReDim VecAux0(0 To xRow, 0 To xCol)
          If RPositionAt < LBound(RVarSrc) Or RPositionAt > UBound(RVarSrc) Then

            For i = LBound(RVarSrc) To UBound(RVarSrc)
              For j = LBound(RVarSrc, 2) To UBound(RVarSrc, 2)
                Let VecAux0(k, l) = RVarSrc(i, j)
                Let l = l + 1
              Next j
              Let l = 0
              Let k = k + 1
            Next i
            If EnuDimensionVecAdd = eqlRes1D Then
              For i = LBound(RVarAdd) To UBound(RVarAdd)
                Let VecAux0(UBound(VecAux0), l) = RVarAdd(i)
                Let l = l + 1
              Next i
            Else
              Let l = 0
              For i = LBound(RVarAdd) To UBound(RVarAdd)
                For j = LBound(RVarAdd, 2) To UBound(RVarAdd, 2)
                  Let VecAux0(k, l) = RVarAdd(i, j)
                  Let l = l + 1
                Next j
                Let l = 0
                Let k = k + 1
              Next i
            End If

          Else

            Let RPositionAt = RPositionAt - 1
            For i = LBound(RVarSrc) To RPositionAt
              For j = LBound(RVarSrc, 2) To UBound(RVarSrc, 2)
                Let VecAux0(l, k) = RVarSrc(i, j)
                Let k = k + 1
              Next j
              Let k = 0
              Let l = l + 1
            Next i

            If EnuDimensionVecAdd = eqlRes1D Then
              Let k = 0
              For i = LBound(RVarAdd) To UBound(RVarAdd)
                Let VecAux0(l, k) = RVarAdd(i)
                Let k = k + 1
              Next i
              Let l = l + 1
            Else
              For i = LBound(RVarAdd) To UBound(RVarAdd)
                For j = LBound(RVarAdd, 2) To UBound(RVarAdd, 2)
                  Let VecAux0(l, k) = RVarAdd(i, j)
                  Let k = k + 1
                Next j
                Let l = l + 1
              Next i
            End If

            Let k = 0
            For i = (RPositionAt + 1) To UBound(RVarSrc)
              For j = LBound(RVarSrc, 2) To UBound(RVarSrc, 2)
                Let VecAux0(l, k) = RVarSrc(i, j)
                Let k = k + 1
              Next j
              Let k = 0
              Let l = l + 1
            Next i
          End If
        End If

      Case Else: GoTo EH
    End Select
  End If

  Let ArrayConcat = VecAux0

EH:
  Let VecAux0 = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN, vbInformation)

End Function

Public Function ArrayTranspose(ByVal RVarSource As Variant) As Variant

  ' Transposes any 1D array or 2D arrays and return a new array. _
    1D arrays: 1xn are converted in 2D arrays mx1 but not viceversa. _
    2D arrays: mxn are converted in 2D arrays nxm and vicesersa. _
               If 2D array has more than TRANSPOSE_ARRAY_CONSTANT [~5200 items], _
               conversion is executed using for - next loop.

  Dim VecAux0 As Variant
  Dim i As Long, j As Long, xCol As Long, xRow  As Long
  Dim EnuDimension As EQLRES_ENU_DIMENSIONARRAY

  Const TRANSPOSE_ARRAY_CONSTANT As Long = 5200
  Const sFN As String = "RES::ArrayTranspose"

  On Error GoTo EH

  If Not VBA.IsArray(RVarSource) Then Call VBA.Err.Raise(CUSTOM_ERROR_RES, sFN, "¡Se requiere un array!")
  Let i = Me.ArrayDimension(RVarSource)
  Let EnuDimension = IIf(i = 1, eqlRes1D, IIf(i = 2, eqlRes2D, eqlResNoArray))
  If EnuDimension = eqlResNoArray Then Call VBA.Err.Raise(CUSTOM_ERROR_RES, sFN, "¡Se requiere un array!")

  If EnuDimension = eqlRes1D Then '1D

    Let xRow = UBound(RVarSource)
    ReDim VecAux0(0 To xRow, 0 To 0)
    For i = 0 To UBound(VecAux0)
      Let VecAux0(i, 0) = RVarSource(i)
    Next i

  Else '2D

    Let xCol = UBound(RVarSource, 2) - LBound(RVarSource, 2)
    Let xRow = UBound(RVarSource) - LBound(RVarSource)
    If xCol >= TRANSPOSE_ARRAY_CONSTANT Or xRow >= TRANSPOSE_ARRAY_CONSTANT Then 'Apply loop
      ReDim VecAux0(LBound(RVarSource, 2) To UBound(RVarSource, 2), LBound(RVarSource) To UBound(RVarSource))
      For j = LBound(RVarSource, 2) To UBound(VecAux0, 2)
        For i = LBound(RVarSource) To UBound(RVarSource)
          Let VecAux0(i, j) = RVarSource(j, i)
        Next i
      Next j

    Else 'Apply Transpose Application method
      Let VecAux0 = Application.WorksheetFunction.Transpose(RVarSource)
    End If

  End If

  Let ArrayTranspose = VecAux0

EH:
  Let VecAux0 = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN, vbInformation)

End Function

Public Function ArrayUniques(ByVal RVarSrc As Variant) As Variant

  Dim VecAux0 As Variant, VecAux1 As Variant
  Dim StrNeedle As String
  Dim i As Long, j As Long, k As Long

  Const sFN As String = "RES::ArrayUniques"

  On Error GoTo EH

  If Not VBA.IsArray(RVarSrc) Then Call VBA.Err.Raise(CUSTOM_ERROR_RES, sFN, "¡Se requiere un array!")
  If (UBound(RVarSrc) = 0) Then Let ArrayUniques = RVarSrc: GoTo EH

  Let i = 1
  Let VecAux0 = RVarSrc

  For i = 0 To UBound(VecAux0)
    Let StrNeedle = VBA.LCase(VBA.CStr(VecAux0(i)))
    Let j = Me.ArrayHasValueN(RVarSrc, StrNeedle, False, True)
    If j > 1 Then
      If i < UBound(VecAux0) Then
        For k = i + 1 To UBound(VecAux0)
          If StrNeedle = VBA.LCase(VBA.CStr(VecAux0(k))) Then Let VecAux0(k) = VBA.vbNullString
        Next k
      End If
    End If
  Next i

  Let j = 0
  ReDim VecAux1(0)
  For i = 0 To UBound(VecAux0)
    If Not VBA.CStr(VecAux0(i)) = VBA.vbNullString Then
      ReDim Preserve VecAux1(j)
      Let VecAux1(j) = VecAux0(i)
      Let j = j + 1
    End If
  Next i

  If j = 0 Then GoTo EH
  Let ArrayUniques = VecAux1

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Call EHGLOBAL.ErrorHandlerRaise(sFN, vbInformation)

End Function

Public Function ArrayDimension(ByVal RVarSrc As Variant) As Long
  Dim i As Long, j As Long, xLength As Long

  Const sFN As String = "RES::ArrayDimension"

  On Error GoTo EH

  'Assuming that Src is not array
  Let xLength = -1
  If Not VBA.IsArray(RVarSrc) Then Call VBA.Err.Raise(CUSTOM_ERROR_RES, sFN, "¡Se requiere un array!")

  'If is array, calculate its dimensions
  Let i = 1
  Do While True
    Let j = (UBound(RVarSrc, i) - LBound(RVarSrc, i)) + 1
    Let xLength = i
    Let i = i + 1
  Loop

EH:
  Let ArrayDimension = xLength
  Call EHGLOBAL.ErrorHandlerClear

End Function

Public Sub ArrayNumbersCommasToDots(ByRef RVarSrc As Variant)

  ' Changes every number with decimal comma separated to dot separated (MS:#$%@)
  Dim i As Long, j As Long

  On Error GoTo EH
  For j = 0 To UBound(RVarSrc, 2)
    For i = 0 To UBound(RVarSrc)
      If (Not RVarSrc(i, j) = Empty) Or (RVarSrc(i, j) = 0) Then
        If VBA.IsNumeric(RVarSrc(i, j)) Then
          Let RVarSrc(i, j) = VBA.Replace(RVarSrc(i, j), ",", ".")
        End If
      End If
    Next i
  Next j

EH:
  Call EHGLOBAL.ErrorHandlerRaise("RES::ArrayNumbersCommasToDots")

End Sub
' CORE METHODS LIST (PUBLIC) ------------------------------------------------------- '
' ======================== CLASS EXCLIQ LITE RES STRUCTURE ========================= '


