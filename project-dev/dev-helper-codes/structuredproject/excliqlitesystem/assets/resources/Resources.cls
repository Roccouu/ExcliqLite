VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppResources_current"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Controlador que brinda funcionalidades básicas al sistema (Recursos)
Option Explicit

' OBJECT VARIABLES (GLOOBJ_) --------------------------------------------------------'
Private EHGLOBAL As AppErrorHandler

' VARIABLES (GLOStr_) ---------------------------------------------------------------'
Private GLOBOO_EH As Boolean
'Public Enum EQLRES_ENU_VALUETYPE
'  eqlResNumbers
'  eqlResStrings
'  eqlResDates
'  eqlResRanges
'End Enum
'Public Enum EQLRES_ENU_RNGRC
'  eqlResRngRow
'  eqlResRngCol
'End Enum
'Public Enum EQLRES_ENU_DIMENSIONARRAY
'  eqlRes1D
'  eqlRes2D
'  eqlResNoArray
'  eqlResDefaultArray
'End Enum

'ERROR CONTROL
Private Const CUSTOM_ERROR As Long = VBA.vbObjectError + 514

Private AppHeader As String
Private AppRangeName As String

Private ELST As Variant

Private EQX As String
Private EQY As String
Private EQXRND As String
Private EQYRND As Long
Private CELLX As String
Private CELLY As String
Private DBCELL As String
Private DBKEYCELL As String
Private APPNAME As String
Private AppTitle As String
Private APPEXT As String
Private DBEXT As String
Private MAINPATH As String
Private OLDSTATUSBAR As Variant

Private pcName As String
Private DBPATH As String
Private DBKEY As String

'PROPERTIES
'Error handler
Public Property Set ErrorHandler(ByRef ObjEH As AppErrorHandler)
  Set EHGLOBAL = Nothing
  Set EHGLOBAL = ObjEH
  Let GLOBOO_EH = (Not ObjEH Is Nothing)
End Property

'Sys references
Public Property Get ERRORList() As Variant
  ERRORList = ELST
End Property

'DB references
Public Property Get EQXRES() As String
  EQXRES = EQX
End Property
Public Property Get EQYRES() As String
  EQYRES = EQY
End Property
Public Property Get EQXRNDRES() As String
  EQXRNDRES = EQXRND
End Property
Public Property Get EQYRNDRES() As String
  EQYRNDRES = EQYRND
End Property
Public Property Get CELLXRES() As String
  CELLXRES = CELLX
End Property
Public Property Get CELLYRES() As String
  CELLYRES = CELLY
End Property
Public Property Get APPNAMERES() As String
  APPNAMERES = APPNAME
End Property
Public Property Get APPTITLERES() As String
  APPTITLERES = AppTitle
End Property
Public Property Get pRESTITLE() As String
  pRESTITLE = AppTitle
End Property
Public Property Get APPEXTRES() As String
  APPEXTRES = APPEXT
End Property
Public Property Get DBEXTRES() As String
  DBEXTRES = DBEXT
End Property
Public Property Get MAINPATHRES() As String
  MAINPATHRES = MAINPATH
End Property

Public Property Get DBCELLRES() As String
  DBCELLRES = DBCELL
End Property
Public Property Get DBKEYCELLRES() As String
  DBKEYCELLRES = DBKEYCELL
End Property




'CONSTRUCTOR/DESTRUCTOR
Private Sub Class_Initialize()
  Set EHGLOBAL = New AppErrorHandler
  APPEXT = "xlsx"
  MAINPATH = ThisWorkbook.Path
End Sub

Private Sub Class_Terminate()
  If GLOBOO_EH = False Then Set EHGLOBAL = Nothing
End Sub




'PRIVATE EXCLIQ LITE RES METHODS
Private Sub RES_ErrorHandler( _
  ByVal sFN As String, _
  Optional MsgTypeIcon As VbMsgBoxStyle = vbExclamation, _
  Optional RClear As Boolean = False)
  
  If RClear Then
    Call EHGLOBAL.ErrorHandlerClear
  Else
    If GLOBOO_EH Then
      Call EHGLOBAL.ErrorHandlerRaise(sFN, MsgTypeIcon)
    Else
      Call EHGLOBAL.ErrorHandlerDisplay(sFN, MsgTypeIcon)
    End If
  End If
End Sub


'PUBLIC EXCLIQ LITE RES METHODS
'Excel appication methods
Public Function ExcelStrIsRangeOfEmpties(ByVal rCells As Range) As Boolean
  Dim i As Range
  For Each i In rCells.Cells
    Let ExcelStrIsRangeOfEmpties = (Not VBA.IsEmpty(i.Value)): If Not ExcelStrIsRangeOfEmpties Then Exit For
  Next i
  Set i = Nothing
End Function

Public Function ExcelSelfOrNewSheet(ByVal refRange As String, ByVal NameSheet As String, ByVal customNameSheet As String, Optional Wbk As Workbook) As Range
  'Crea una nueva hoja si la referencia entregada es incorrecta o new dándole el nombre personalizado. Retorna el rango referenciado de refRange o A1 de new
  Dim Sh As Worksheet
  Dim Wb As Workbook
  
  If Wbk Is Nothing Then Set Wb = ActiveWorkbook
  If Not Wbk Is Nothing Then Set Wb = Wbk
  
  On Error GoTo EH
  If refRange = "new" Then 'crear nueva hoja de cálculo y pasar las coordenadas A1 como rango
    Set Sh = Me.ExcelCurrentOrNewSheet(customNameSheet, Wb)
    Set ExcelSelfOrNewSheet = Sh.Range("A1")
  Else 'refRange es la referencia a un rango, se debe validar el rango y si es correcto obtenerlo y devolverlo. Crear nuevo
    If Me.ExcelIsRange(refRange) And Me.ExcelSearchSheet(NameSheet, Wbk) Then
      Wb.Worksheets(NameSheet).Activate
      Set ExcelSelfOrNewSheet = Wb.Worksheets(NameSheet).Range(refRange)
    Else
      Set Sh = Me.ExcelCurrentOrNewSheet(customNameSheet, Wb)
      Set ExcelSelfOrNewSheet = Sh.Range("A1")
    End If
  End If
  Set Wb = Nothing: Set Sh = Nothing
EH:
  If VBA.Err.Number > 0 Then Call Me.SysErrorHandler("¡Error al crear una nueva hoja de cálculo: " & customNameSheet & "!", "Resources: RESoNSheet")
End Function

Public Function ExcelIsRange(ByVal Address As String) As Boolean
  On Error GoTo EH
  Let ExcelIsRange = (Range(Address).Cells.Count >= 1)
EH:
  On Error GoTo 0
End Function

Public Function ExcelGetUniqueRange(ByVal Holder As Object, ByVal rName As String) As String
  Dim i As Integer
  Dim RgNameLocal As String
  Dim RgNameAux As String
    
  Let i = 0
  Let RgNameLocal = rName
  On Error GoTo EH
  If ExcelPageSetupRangeNameExists(Holder, rName) Then
    Do
      Let i = i + 1
      Let RgNameLocal = rName & "_" & i
    Loop While (ExcelPageSetupRangeNameExists(Holder, RgNameLocal))
  End If
  Let ExcelGetUniqueRange = RgNameLocal

EH:
  If Err.Number > 0 Then
    RgNameAux = "CustomRange" & Holder.Names.Count + 1
    ExcelGetUniqueRange = RgNameAux
    Call Me.SysErrorHandler("Imposible asignar el nombre: " & RgNameLocal & " al rango de la hoja actual, se creó el nombre de rango: " & RgNameAux & " en su lugar.", "Res: ExcelGetUniqueRange")
  End If
End Function

Public Function ExcelCurrentOrNewSheet(ByVal NameSheet As String, Optional Wbk As Workbook) As Worksheet
  
  ' Check if Worksheet named NameSheet exists, if doesn't creates it, else creates the new sheet with NameSheet_x.
  Dim Wb As Workbook
  Dim newSheetName As String

  Const sFN As String = "RES::ExcelCreateChartSheet"
  
  On Error GoTo EH
  If Wbk Is Nothing Then Set Wb = Application.ActiveWorkbook
  If Not Wbk Is Nothing Then Set Wb = Wbk

  ' Creates the new worksheet
  Let newSheetName = Me.ExcelCountSheetNameCoincidences(NameSheet, Wb)
  If Me.ExcelCreateSheet(newSheetName, Wbk:=Wb) Then
    Call Wb.Sheets(newSheetName).Activate
    Set ExcelCurrentOrNewSheet = Wb.Sheets(newSheetName)
  End If

EH:
  Set Wb = Nothing
  Call RES_ErrorHandler(sFN, vbInformation)

End Function

Public Function ExcelSearchSheet(ByVal NameSheet As String, Optional Wbk As Workbook) As Boolean
  Dim Wb As Workbook
  
  On Error Resume Next
  If Wbk Is Nothing Then Set Wb = ActiveWorkbook
  If Not Wbk Is Nothing Then Set Wb = Wbk
  Let ExcelSearchSheet = (Wb.Worksheets(NameSheet).Name <> VBA.vbNullString)
  
  Set Wb = Nothing
  On Error GoTo 0
End Function

Private Function ExcelPageSetupRangeNameExists(ByVal Sh As Object, Optional rName As String = VBA.vbNullString) As Boolean
  'Allows to know if Sh has a Nm range name
  Dim n As Name
  
  With Sh
    For Each n In Sh.Names
      If VBA.LCase(n.Name) = VBA.LCase(IIf(VBA.IsEmpty(rName), AppRangeName, rName)) Then ExcelPageSetupRangeNameExists = True: Exit For
    Next n
  End With
  Set n = Nothing
End Function

Public Function ExcelCountSheetNameCoincidences(ByVal ShName As String, Optional Wbk As Workbook) As String
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
    Call Me.SysErrorHandler("Imposible asignar el nombre: " & ShName & " a la nueva hoja, se creó la hoja: " & ShNameAux & " en su lugar.", "Res: NameCoincidences")
    'Call EHGLOBAL.ErrorHandlerRaise("RES::ExcelCountSheetNameCoincidences", vbInformation)
    Call RES_ErrorHandler(sFN, vbInformation)
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
  Call RES_ErrorHandler(sFN, vbInformation)
End Function

Public Function ExcelCreateChartSheet(ByVal NameSheet As String, Optional Wbk As Workbook, Optional AfterSheet As String = VBA.vbNullString, Optional BeforeSheet As String = VBA.vbNullString) As Chart
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

'Excel validation methods
Public Function ExcelNumRangeDoubleValidation(ByVal rRD As Variant) As Boolean
  On Error GoTo EH
  Let ExcelNumRangeDoubleValidation = (Not (TypeName(rRD) = "Range" Or TypeName(rRD) = "Double")): If ExcelNumRangeDoubleValidation Then GoTo EH
  Let ExcelNumRangeDoubleValidation = IIf(TypeName(rRD) = "Range", (Not ExcelNumIsRangeOfNumbers(rRD)), (Not VBA.IsNumeric(rRD) Or VBA.IsEmpty(rRD)))
EH:
End Function

'Excel Date types
Public Function ExcelDateSys(ByVal DateType As Single, Optional Place As String = VBA.vbNullString, Optional Separator As String = "/") As String
  If DateType = 0 Then ExcelDateSys = Place & ", " & VBA.StrConv(VBA.Format(VBA.Date, "DDDD"), vbProperCase) & ", " & VBA.Format(VBA.Date, "DD") & " de " & VBA.StrConv(VBA.Format(VBA.Date, "MMMM"), vbProperCase) & " de " & VBA.Format(VBA.Date, "YYYY")
  If DateType = 1 Then ExcelDateSys = Place & ", " & VBA.StrConv(VBA.Format(VBA.Date, "DDDD"), vbProperCase) & ", " & VBA.Format(VBA.Date, "DD") & " de " & VBA.StrConv(VBA.Format(VBA.Date, "MMMM"), vbProperCase) & " de " & VBA.Format(VBA.Date, "YYYY") & " - " & Now
  If DateType = 2 Then ExcelDateSys = VBA.Format(VBA.Date, "DD") & " de " & VBA.StrConv(VBA.Format(VBA.Date, "MMMM"), vbProperCase) & " de " & VBA.Format(VBA.Date, "YYYY")
  If DateType = 3 Then ExcelDateSys = VBA.StrConv(VBA.Format(VBA.Date, "MMMM"), vbProperCase) & ", " & VBA.Format(VBA.Date, "DD") & " de " & VBA.Format(VBA.Date, "YYYY")
  If DateType = 4 Then ExcelDateSys = VBA.Format(VBA.Date, "DD") & Separator & VBA.Format(VBA.Date, "MM") & Separator & VBA.Format(VBA.Date, "YY")
End Function

Public Sub ExcelSetRangeName(ByVal Holder As Object, ByVal rName As String, ByVal refRng As Range, Optional v As Boolean = True)
  'Generic, inserts a range name
  'MsgBox "Insertando nombre de rango: " & rName & " en: " & Holder.Name
  With Holder
    If Not ExcelPageSetupRangeNameExists(Holder, rName) Then Holder.Names.Add Name:=rName, RefersTo:=refRng, Visible:=v
  End With
End Sub

'Excel validation methods
Public Function ExcelNumIsRangeOfNumbers(ByVal rNumbers As Range, Optional rPositive As Boolean = False, Optional rZero As Boolean = False) As Boolean
  Dim i As Range
  
  For Each i In rNumbers.Cells
    Let ExcelNumIsRangeOfNumbers = (VBA.IsNumeric(i.Value) And Not VBA.IsEmpty(i.Value))
    If ExcelNumIsRangeOfNumbers Then ExcelNumIsRangeOfNumbers = IIf(rPositive, (i.Value >= 0), IIf(rZero, (i.Value <> 0), ExcelNumIsRangeOfNumbers))
    If Not ExcelNumIsRangeOfNumbers Then Exit For
  Next i
  
  Set i = Nothing
End Function





'System management Methods
Public Sub SysErrorHandler( _
  Optional Msg As String = "Ha ocurrido un error inesperado, consulte al Proveedor.", _
  Optional Origin As String = VBA.vbNullString, _
  Optional addMsg As Boolean = False, _
  Optional nLineMsg As Boolean, Optional MsgType As Variant = vbCritical)
  
  Dim primaryMsg As String
  primaryMsg = "Ha ocurrido un error inesperado, consulte al Proveedor."
  
  MsgBox _
    IIf(addMsg, primaryMsg & IIf(nLineMsg, VBA.vbNewLine & Msg, " " & Msg), Msg) & VBA.vbNewLine & _
    IIf(Origin = VBA.vbNullString, VBA.vbNullString, VBA.vbNewLine & "Orígen: " & Origin) & VBA.vbNewLine & _
    "Error: " & VBA.Err.Number & VBA.vbNewLine & _
    VBA.Err.Description, vbCritical, AppTitle
End Sub

'Forms management methods
Public Function FrmValidateRefEdit(ByVal refEd As Object, Optional refMsg As String = VBA.vbNullString, Optional emptyMsg As String = VBA.vbNullString) As Boolean
  Dim REGEX As AppResRegEx
  
  Set REGEX = New AppResRegEx
  If Not REGEX.isEmptyStringReg(Trim(refEd.Text)) Then
    If Not REGEX.isRange(Trim(refEd.Text)) Then
      FrmCustomMsgboxRequiredControl refEd, "La referencia: " & refEd.Text & " es incorrecta.", vbExclamation: Call FrmSelTextControl(refEd)
    Else
      If Not ExcelIsRange(Trim(refEd.Text)) Then
        If Not ExcelIsRange(Trim(Me.StrReplaceCommas(refEd.Text))) Then
          FrmCustomMsgboxRequiredControl refEd, "La referencia: " & refEd.Text & " es incorrecta.", vbExclamation: Call FrmSelTextControl(refEd)
        Else
          refEd.Text = Trim(Me.StrReplaceCommas(refEd.Text))
          FrmValidateRefEdit = True
        End If
      Else
        FrmValidateRefEdit = True
      End If
    End If
  Else
    FrmCustomMsgboxRequiredControl refEd, "Seleccione un rango de celdas por favor.", vbExclamation: Call FrmSelTextControl(refEd)
  End If
  Set REGEX = Nothing
End Function

Public Sub FrmCustomMsgboxRequiredControl( _
  ByVal xControl As Object, _
  ByVal StrMsg As String, _
  Optional TYPEMSG As VbMsgBoxStyle = vbCritical, _
  Optional SelTxt As Boolean = False)
  
  Dim Color As String
  Dim cBackColor As Variant
  
  With xControl
    Let cBackColor = .BackColor
    Select Case VBA.TypeName(xControl)
      Case "TextBox", "ComboBox", "Label", "ListBox"
        Let Color = .BorderColor
        
        Let .BorderColor = VBA.vbRed
        Let .BackColor = VBA.RGB(255, 230, 230)
        Call VBA.MsgBox(StrMsg, TYPEMSG, AppTitle)
        Let .BorderColor = Color
        Let .BackColor = cBackColor
        If .Enabled And .Visible Then Call .SetFocus
        If SelTxt Then Call Me.FrmSelTextControl(xControl)

      Case "CommandButton", "ToggleButton", "CheckBox", "OptionButton", "SpinButton"
        Let Color = .ForeColor
        
        Let .ForeColor = VBA.vbRed
        Let .BackColor = VBA.RGB(255, 230, 230)
        Call VBA.MsgBox(StrMsg, TYPEMSG, AppTitle)
        Let .ForeColor = Color
        Let .BackColor = cBackColor
      
      Case Else: GoTo EH
    End Select
  End With

EH:
End Sub

Public Sub FrmSelTextControl(ByVal Control As Variant)
  With Control
    .SelStart = 0
    .SelLength = Len(Control.Text)
    If .Enabled Then .SetFocus
  End With
End Sub

Public Sub FrmCustomMsgboxRequiredLabel(ByVal Control As Object, ByVal Msg As String, Optional TYPEMSG As Variant = vbCritical)
  Dim Color As String
  Dim cBackColor As Variant
  Dim cBackStyle As Variant
  Dim cBorderStyle As Variant
  Dim cForeColor As Variant
  With Control
    cBackStyle = .BackStyle
    cBorderStyle = .BorderStyle
    cForeColor = .ForeColor
    Color = .BorderColor
    
    cBackColor = .BackColor
    .BackStyle = 1
    .BorderStyle = 1
    .ForeColor = VBA.vbRed
    .BorderColor = VBA.vbRed
    .BackColor = VBA.RGB(255, 230, 230)
    MsgBox Msg, TYPEMSG, AppTitle
    .BorderColor = Color
    .BackColor = cBackColor
    .ForeColor = cForeColor
    .BackStyle = cBackStyle
    .BorderStyle = cBorderStyle
  End With
End Sub

Public Function FrmCountListboxSelectedItems(ByVal LBox As Object, Optional WithHeader As Single = 0) As Integer
  Dim i As Single
  FrmCountListboxSelectedItems = 0
  If WithHeader >= LBox.ListCount Then GoTo RFLSI
  For i = WithHeader To LBox.ListCount - 1
    FrmCountListboxSelectedItems = FrmCountListboxSelectedItems + IIf(LBox.Selected(i), 1, 0)
  Next i
RFLSI:
End Function

Public Function FrmValidateTextNum(ByVal txt As Object, Optional txtMsg As String = VBA.vbNullString, Optional emptyMsg As String = VBA.vbNullString) As Boolean
  'Valida que un control textbox contenga números
  Dim REGEX As AppResRegEx
  Dim E As Boolean
  
  On Error GoTo RFVTNE
  Let E = False
  Set REGEX = New AppResRegEx
  If Not REGEX.isEmptyStringReg(Trim(txt.Text)) Then
    If Not VBA.IsNumeric(Trim(txt.Text)) Then E = True: GoTo RFVTNE
    txt.Text = Me.StrReplaceCommas(txt.Text, ".")
    If Not VBA.IsNumeric(Trim(txt.Text)) Then E = True: GoTo RFVTNE
  Else
    Call Me.FrmCustomMsgboxRequiredControl(txt, "Escriba un número.", vbExclamation): Call Me.FrmSelTextControl(txt): GoTo RFVTNE
  End If
  FrmValidateTextNum = True
RFVTNE:
  If E Then Call Me.FrmCustomMsgboxRequiredControl(txt, "El valor: " & txt.Text & " es incorrecto.", vbExclamation): Call FrmSelTextControl(txt)
  If VBA.Err.Number <> 0 Then Call Me.SysErrorHandler(Origin:="Recursos, validación números txt.")
  Set REGEX = Nothing
  On Error GoTo 0
End Function

Public Function FrmNumberTextBox(ByVal numData As MSForms.ReturnInteger, ByRef txtData As Object, Optional Shft As Integer, Optional eventType As Boolean = False, Optional TxtOrNum As Boolean = False) As Variant
  'Permite la introducción de sólo números en un textbox
  'FrmNumberTextBox = IIf((numData >= 48 And numData <= 57) Or numData = 44, numData, 0)
  Dim commas() As String
  Dim i As Integer
  Dim td As String
  
  If eventType Then
    If TxtOrNum Then 'Text
    '36($),37(%),46(.),47(/),64-90(@-Z),97-122(a-z)
      FrmNumberTextBox = IIf((numData >= 64 And numData < 90) Or (numData >= 97 And numData < 122) Or (numData = 32 Or numData = 36 Or numData = 37 Or numData = 46 Or numData = 47), numData, 0)
    Else 'Numbers
      td = txtData.Text
      If td <> VBA.vbNullString Then
        commas = VBA.Split(td, ",")
        i = UBound(commas)
      End If
      FrmNumberTextBox = IIf((numData >= 48 And numData <= 57) Or (numData = 46 And (i = 0)), IIf(numData = 46, 44, numData), 0)
    End If
  Else
    If (Shft = 2 And numData = 118) Or (Shft = 2 And numData = 86) Then txtData = Null
  End If
End Function

'App methods
Public Function AppGetColumnAlphaSemi26Base(ByVal nNum As Long, Optional nBase As Long = 25, Optional nAcum As String = VBA.vbNullString) As String
  'Permite obtener letras A-Z|AA-AL en base a un semi sistema de base 25
  If VBA.Abs(nNum) <= nBase Then
    AppGetColumnAlphaSemi26Base = VBA.Chr(nNum + 64) & VBA.StrReverse(nAcum)
  Else
    AppGetColumnAlphaSemi26Base = AppGetColumnAlphaSemi26Base(nNum / nBase, nBase, nAcum & VBA.Chr((nNum Mod nBase) + 64))
  End If
End Function

'String managemente methods
Public Function StrReplaceCommas(ByVal Str As String, Optional sChar As String = ";") As String
  StrReplaceCommas = VBA.Replace(Str, sChar, ",")
End Function

Public Function StrUCaseWord(s As String) As String
  If Len(s) > 0 Then
    Dim l As String
    Dim half As String
    l = Left(s, 1)
    half = Mid(s, 2)
    l = UCase(l)
    StrUCaseWord = l & half
  End If
End Function













'AUGMENTED METHODS
'Config









'AUGMENTED METHODS
'New on both projects
Private Sub Debug_assert(ByVal bDeb As Boolean, Optional sCaller As String = VBA.vbNullString)
  Debug.Assert bDeb
  If Not bDeb Then Debug.Print sCaller
End Sub

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
  Call RES_ErrorHandler("RES::ExcelEdges", vbInformation)
End Function

Public Function ExcelNumIsRangeOfNumbersAndEmpties( _
  ByVal rNumbers As Range, _
  Optional rZero As Boolean = False) As Boolean
  
  Dim i As Range
  
  For Each i In rNumbers.Cells
    Let ExcelNumIsRangeOfNumbersAndEmpties = VBA.IsNumeric(i.Value)
    If ExcelNumIsRangeOfNumbersAndEmpties Then Let ExcelNumIsRangeOfNumbersAndEmpties = IIf(rZero, (i.Value <> 0), ExcelNumIsRangeOfNumbersAndEmpties)
    If Not ExcelNumIsRangeOfNumbersAndEmpties Then Exit For
  Next i
  
  Set i = Nothing
End Function

Public Function ExcelVLockup( _
  ByVal VarNeedle As String, _
  ByVal RngSource As Range, _
  ByVal RngValue As Range) As String
  
  Dim StrNeedle As String
  Dim i As Long, j As Long
  
  On Error GoTo EH
  Let ExcelVLockup = VBA.vbNullString
    
  For i = 1 To RngSource.Rows.Count
    If VBA.LCase(VarNeedle) = VBA.LCase(VBA.CStr(RngSource.Cells(i, 1).Value)) Then Let ExcelVLockup = VBA.CStr(RngValue.Cells(i, 1).Value)
  Next i
  
EH:
  Call RES_ErrorHandler("RES::ExcelVLockup", vbInformation)
End Function






'Arrays
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
    Call RES_ErrorHandler("RES::ArrayLength", vbInformation)
    'Call EHGLOBAL.ErrorHandlerRaise("RES::ArrayLength", vbInformation)
  ElseIf Not (VBA.Err.Number = 0) Then
    On Error GoTo -1
    Call VBA.Err.Clear
  End If

End Function

Function ArrayLength(ByRef arr As Variant) As Long
  'https://excelmacromastery.com/
  On Error GoTo EH
  
  ' Loop is used for multidimensional arrays. The Loop will terminate when a
  ' "Subscript out of Range" error occurs i.e. there are no more dimensions.
  Dim i As Long, length As Long
  Let length = 1
  
  If Not VBA.IsArray(arr) Then Call VBA.Err.Raise(CUSTOM_ERROR, "RES::ArrayLength", "¡Se esperaba un array!")
  
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
  Call RES_ErrorHandler("RES::ArrayLength", vbInformation, Not (VBA.Err.Number = 13))
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
      Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "Se esperaba un array.")
    ElseIf i > 2 Then
      Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "Se esperaba un array de dimensión 1 o 2.")
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
      Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "Se esperaba un array de dimensión 1 o 2 sin contenidos nulos.")
    Else
      ReDim VecVector(0 To 0, 0 To 0)
      Let VecVector(0, 0) = AVarVector
    End If
  End If
  
  Let k = 0
  If AEnuResultDimension = eqlRes1D Then 'Response AS 1D
    If LngDimension = 2 Then 'Just if 2D
      Let i = Me.ArrayLength(VecVector) - 1
      ReDim VecVector0(0 To i)
      If AByCols Then 'Cols
        For j = LBound(VecVector, 2) To UBound(VecVector, 2)
          For i = LBound(VecVector) To UBound(VecVector)
            Let VecVector0(k) = VecVector(i, j)
            Let k = k + 1
          Next i
        Next j
      Else 'Rows
        For i = LBound(VecVector) To UBound(VecVector)
          For j = LBound(VecVector, 2) To UBound(VecVector, 2)
            Let VecVector0(k) = VecVector(i, j)
            Let k = k + 1
          Next j
        Next i
      End If
    End If
    Let VecVector = Empty
    Let VecVector = VecVector0
  ElseIf AEnuResultDimension = eqlRes2D Then 'Response as 2D
    If LngDimension = 1 Then 'Just if 1D
      If AByCols Then 'Cols
        ReDim VecVector0(0 To 0, 0 To UBound(VecVector))
        For i = LBound(VecVector) To UBound(VecVector)
          Let VecVector0(0, i) = VecVector(i)
        Next i
      Else 'Rows
        ReDim VecVector0(0 To UBound(VecVector), 0 To 0)
        For i = LBound(VecVector) To UBound(VecVector)
          Let VecVector0(i, 0) = VecVector(i)
        Next i
      End If
    End If
    Let VecVector = Empty
    Let VecVector = VecVector0
  End If
  
  Let ArrayToBaseZero = VecVector

EH:
  Let VecVector = Empty
  Let VecVector0 = Empty
  Call RES_ErrorHandler(sFN, vbInformation)

End Function

Public Function ArrayVectorOfNumbersFromColonSeparatedString(ByVal sVector As String, Optional bAbs As Boolean = False) As Variant
  Dim vVector As Variant
  Dim vFinal() As Variant
  Dim i As Long
  Dim sVectorLocal As String
  Const sFN As String = "RES::ArrayVectorOfNumbersFromColonSeparatedString"
  
  On Error GoTo EH
  Call Debug_assert(Not sVector = VBA.vbNullString, sFN)
  If sVector = VBA.vbNullString Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "Se requiere una cadena con números.")
  Let ArrayVectorOfNumbersFromColonSeparatedString = Array()
  Let sVectorLocal = VBA.Replace(sVector, ",", ".")
  Let vVector = VBA.Split(sVectorLocal, ";")
  If Me.ArrayEmpty(vVector) Then
    ReDim vFinal(0 To UBound(vVector))
    For i = LBound(vVector) To UBound(vVector)
      If bAbs Then
        Let vFinal(i) = VBA.Abs(VBA.CDbl(vVector(i)))
      Else
        Let vFinal(i) = VBA.CDbl(vVector(i))
      End If
    Next i
  End If
  Erase ArrayVectorOfNumbersFromColonSeparatedString
  Let ArrayVectorOfNumbersFromColonSeparatedString = vFinal

EH:
  Erase vFinal
  Let vVector = Empty
  Call RES_ErrorHandler(sFN, vbInformation)
End Function

Public Function ArraySumFromVector(ByVal sVector As Variant, Optional bAbs As Boolean = False) As Double
  Dim i As Long
  Dim bLength As Boolean
  Const sFN As String = "RES::ArraySumFromVector"
  
  On Error GoTo EH
  Let bLength = Me.ArrayEmpty(sVector)
  Call Debug_assert(Not bLength, sFN)
  If Not bLength Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "Se requiere un vector.")
  Let ArraySumFromVector = 0
  For i = LBound(sVector) To UBound(sVector)
    If bAbs Then
      Let ArraySumFromVector = ArraySumFromVector + VBA.Abs(VBA.CDbl(sVector(i)))
    Else
      Let ArraySumFromVector = ArraySumFromVector + VBA.CDbl(sVector(i))
    End If
  Next i

EH:
  Call RES_ErrorHandler(sFN, vbInformation)
End Function

Public Function ArrayFromRangeJustNumbers(ByVal NumRange As Range) As Variant
  Dim REGEX As AppResRegEx
  Dim vVector As Variant
  Dim xArea As Range
  Dim i As Long, j As Long, k As Long
  Const sFN As String = "RES::ArrayFromRangeJustNumbers"
  
  On Error GoTo EH
  ReDim vVector(0 To 0)
  Call Debug_assert(Not NumRange Is Nothing, sFN)
  If NumRange Is Nothing Then GoTo EH
  Erase vVector
  
  Set REGEX = New AppResRegEx
  Let k = 0
  If NumRange.Areas.Count > 1 Then
    For Each xArea In NumRange
      For i = 1 To xArea.Columns.Count
        For j = 1 To xArea.Rows.Count
          Debug.Print xArea.Cells(j, i).Value
          If REGEX.isNumberMinusReg(xArea.Cells(j, i).Value) Then
            ReDim Preserve vVector(k)
            Let vVector(k) = xArea.Cells(j, i).Value
            Let k = k + 1
          End If
        Next j
      Next i
    Next xArea
  Else
    For i = 1 To NumRange.Columns.Count
      For j = 1 To NumRange.Rows.Count
        Debug.Print NumRange.Cells(j, i).Value
        If REGEX.isNumberMinusReg(NumRange.Cells(j, i).Value) Then
          ReDim Preserve vVector(k)
          Let vVector(k) = NumRange.Cells(j, i).Value
          Let k = k + 1
        End If
      Next j
    Next i
  End If
  Set REGEX = Nothing
  Let ArrayFromRangeJustNumbers = vVector

EH:
  Erase vVector
  Set xArea = Nothing
  Set REGEX = Nothing
  Call RES_ErrorHandler(sFN, vbInformation)
End Function

Public Function ExcelQuantityOfNumbersFromRange( _
  ByVal rRange As Range, _
  Optional bAbs As Boolean = False, _
  Optional bZeros As Boolean = False, _
  Optional BPercents As Boolean = False) As Long
  
  Dim REGEX As AppResRegEx
  Dim vVector As Variant
  Dim i As Long, j As Long
  Dim bVECTOR As Boolean
  Const sFN As String = "ExcelQuantityOfNumbersFromRange"
  
  On Error GoTo EH
  Call Debug_assert(Not rRange Is Nothing, sFN)
  If rRange Is Nothing Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "Se requiere un objeto Range.")
  
  Let vVector = rRange.Value
  Let bVECTOR = Me.ArrayEmpty(vVector)
  Let ExcelQuantityOfNumbersFromRange = 0
  If Not Me.ArrayEmpty(vVector) Then
    Set REGEX = New AppResRegEx
    For i = LBound(vVector) To UBound(vVector)
      For j = LBound(vVector, 2) To UBound(vVector, 2)
        If bAbs Then
          If REGEX.isNumberMinusReg(vVector(i, j)) Then
            If bZeros Then
              'If Not VBA.CDbl(vVECTOR(i, j)) = 0 Then
            End If
          End If
          Let ExcelQuantityOfNumbersFromRange = ExcelQuantityOfNumbersFromRange + VBA.Abs(VBA.CDbl(vVector(i)))
        Else
          Let ExcelQuantityOfNumbersFromRange = ExcelQuantityOfNumbersFromRange + VBA.CDbl(vVector(i))
        End If
      Next j
    Next i
    Set REGEX = Nothing
  End If

EH:
  Let vVector = Empty
  Call RES_ErrorHandler(sFN, vbInformation)
End Function

Function ArrayGetDataFromRangesByRows( _
  ByVal Rng As Range, _
  ByRef RCounter As Long, _
  Optional RType As Boolean = False, _
  Optional RAbs As Boolean = False, _
  Optional RIncludeZeroes As Boolean = False, _
  Optional RIncludeEmptyStrings As Boolean = False, _
  Optional RCustomRegex As String = VBA.vbNullString) As Variant
  
  'Gets all data number/strings from ranges in order of Rows
  Dim REGEX As AppResRegEx
  Dim r As Range
  Dim VData As Variant
  Dim i As Long, j As Long
  
  Let ArrayGetDataFromRangesByRows = Empty
  
  On Error GoTo EH
  If Rng Is Nothing Then GoTo EH
  
  Set REGEX = New AppResRegEx
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
  Set REGEX = Nothing
  
  Let RCounter = i
  Let ArrayGetDataFromRangesByRows = VData
  
EH:
  Let VData = Empty
  Set r = Nothing
  Set REGEX = Nothing
  Call RES_ErrorHandler("RES::ArrayGetDataFromRanges", vbInformation)
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
  Dim REGEX As AppResRegEx
  Dim r As Range
  Dim VData As Variant, VecRng As Variant
  Dim i As Long, kAreas As Long, kRow As Variant, kCol As Variant
  Dim DblAux As Double
  Dim StrAux As String
  
  Let ArrayGetDataFromRangesByCols = Empty
  
  On Error GoTo EH
  If Rng Is Nothing Then GoTo EH
  
  Set REGEX = New AppResRegEx
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
  Set REGEX = Nothing
  
  Let RCounter = i
  Let ArrayGetDataFromRangesByCols = VData
  
EH:
  Let VData = Empty
  Set r = Nothing
  Set REGEX = Nothing
  Call RES_ErrorHandler("RES::ArrayGetDataFromRanges", vbInformation)
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
  
  Dim REGEX As AppResRegEx
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
  Call RES_ErrorHandler("RES::ArrayGetDataAsStringFromRanges", vbInformation)
End Function

Public Function ArrayGetDataColonSeparatedFromString( _
  ByVal StrVector As String, _
  ByRef StrCounter As Long, _
  Optional StrValueType As Boolean = False, _
  Optional StrAbs As Boolean = False, _
  Optional StrIncludeZeroes As Boolean = False, _
  Optional StrIncludeEmptyStrings As Boolean = False, _
  Optional StrCustomRegex As String = VBA.vbNullString) As Variant
  
  Dim REGEX As AppResRegEx
  Dim VData As Variant, vAux As Variant
  Dim i As Long, j As Long
  
  Let ArrayGetDataColonSeparatedFromString = Empty
  
  On Error GoTo EH
  Set REGEX = New AppResRegEx
  
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
  Set REGEX = Nothing
  
  Let StrCounter = j
  Let ArrayGetDataColonSeparatedFromString = VData
  
EH:
  Let VData = Empty
  Let vAux = Empty
  Set REGEX = Nothing
  Call RES_ErrorHandler("RES::ArrayGetDataColonSeparatedFromString", vbInformation)
End Function

Public Function ArrayGetDataColonSeparatedStringFromVariants( _
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
  Dim VecAux As Variant
  Dim StrAux As String, StrResponse As String
  
  Const sFN As String = "RES::ArrayGetDataColonSeparatedStringFromVariants"
  
  On Error GoTo EH
  Let ArrayGetDataColonSeparatedStringFromVariants = VBA.vbNullString
  
  'Process
  Let StrAux = VBA.TypeName(RVarVector)
  Select Case StrAux
    Case "Double", "Integer", "Long", "Single"
      Let StrResponse = VBA.CStr(RVarVector)
    Case "String"
      Let VecAux = ArrayGetDataColonSeparatedFromString(RVarVector, RLngCounter, RBooValueType, RBooAbs, RBooIncludeZeroes, RBooIncludeEmptyStrings, RStrCustomRegex)
      Let StrResponse = VBA.Join(VecAux, ";")
    Case "Range"
      Let StrResponse = ArrayGetDataAsStringFromRanges(RVarVector, RLngCounter, RBooValueType, RBooAbs, RBooIncludeZeroes, RBooIncludeEmptyStrings, RStrCustomRegex, RBooRngRowsCols)
    Case Else: GoTo EH
  End Select
  
  'Return
  Let ArrayGetDataColonSeparatedStringFromVariants = StrResponse
  
EH:
  Let VecAux = Empty
  Call RES_ErrorHandler(sFN, vbInformation)
End Function

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
  Dim REGEX As AppResRegEx
  
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
      Set REGEX = New AppResRegEx
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
      Set REGEX = Nothing
      Let RLngCounter = j
      Let StrResponse = VBA.Join(VecAux2, "/")
      Let VecAux2 = Empty
    
    Case "Range"
      Dim RngBox As Range, RngArea As Range
      Set RngBox = RVarVector
      If Not RngBox Is Nothing Then
        Set REGEX = New AppResRegEx
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
      Set REGEX = Nothing
    
    Case Else: GoTo EH
  End Select
  
  'Return
  Let ArrayGetDataSemicolonSeparatedSlashGroupedStringFromVariants = StrResponse
  
EH:
  Let VecAux = Empty
  Call RES_ErrorHandler(sFN, vbInformation)
End Function

Public Function ArrayGetDataArrayFromVariants( _
  ByVal RVarVector As Variant, _
  ByRef RLngCounter As Long, _
  Optional RBooValueType As Boolean = False, _
  Optional RBooAbs As Boolean = False, _
  Optional RBooIncludeZeroes As Boolean = False, _
  Optional RBooIncludeEmptyStrings As Boolean = False, _
  Optional RStrCustomRegex As String = VBA.vbNullString, _
  Optional RBooRngRowsCols As Boolean = False) As Variant
  
  'Extract all values numbers or strings from any variable as Range or Strings, each separated by semicolon
  'Just works with Strings, Ranges and Numbers
  Dim i As Long
  Dim VecAux As Variant, VecAux1 As Variant
  Dim StrAux As String
  
  Const sFN As String = "RES::ArrayGetDataArrayFromVariants"
  
  On Error GoTo EH
  Let ArrayGetDataArrayFromVariants = Array(0)
  Let RLngCounter = 0
  
  'Process
  Let StrAux = VBA.TypeName(RVarVector)
  Select Case StrAux
    Case "Double", "Integer", "Long", "Single"
      If RBooIncludeZeroes Then
        Let StrAux = VBA.CStr(IIf(RBooAbs, VBA.Abs(RVarVector), RVarVector)): Let RLngCounter = 1
      Else
        Let StrAux = VBA.CStr(IIf(RBooAbs, VBA.Abs(RVarVector), RVarVector)): Let RLngCounter = IIf(RBooIncludeZeroes And RVarVector = 0, 0, 1)
      End If
    Case "String"
      Let VecAux = ArrayGetDataColonSeparatedFromString(RVarVector, RLngCounter, RBooValueType, RBooAbs, RBooIncludeZeroes, RBooIncludeEmptyStrings, RStrCustomRegex)
      Let StrAux = VBA.Join(VecAux, ";") 'TODO: CORREGIR EN CASO DE CADENA VACÍA
    Case "Range"
      Let StrAux = ArrayGetDataAsStringFromRanges(RVarVector, RLngCounter, RBooValueType, RBooAbs, RBooIncludeZeroes, RBooIncludeEmptyStrings, RStrCustomRegex, RBooRngRowsCols)
    Case Else: GoTo EH
  End Select
  
  Let VecAux = VBA.Split(StrAux, ";") 'strings
  If Not RBooValueType Then 'numbers
    If RLngCounter = 0 Then
      Let VecAux1 = Array(0)
    Else
      ReDim VecAux1(UBound(VecAux))
      For i = 0 To UBound(VecAux)
        Let VecAux1(i) = VBA.CDbl(VecAux(i))
      Next i
    End If
    Let VecAux = Empty
    Let VecAux = VecAux1
    Let VecAux1 = Empty
  End If
  
  'Return
  Let ArrayGetDataArrayFromVariants = VecAux
  
EH:
  Let VecAux = Empty
  Debug.Print VBA.Err.Description
  Call RES_ErrorHandler(sFN, vbInformation)
End Function

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
  Call RES_ErrorHandler(sFN, vbInformation)
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
  Optional RRngRC As EQLRES_ENU_RNGRC = eqlResRngCol) As String

  Dim REGEX As AppResRegEx
  
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
  Set REGEX = New AppResRegEx
  Let i = 0
  Let j = 0
  ReDim VData(0)
  If RValueType = eqlResStrings Then 'Strings
    Let vAux = VBA.Split(StrData, RXChar)
    For i = 0 To UBound(vAux)
      If REGEX.isStringReg(vAux(i), RStrCustomRegex) Then
        ReDim Preserve VData(j)
        If RBooIncludeEmptyStrings Then
          Let VData(j) = vAux(i)
        Else
          Let VData(j) = vAux(i)
        End If
        Let j = j + 1
      End If
    Next i
  ElseIf RValueType = eqlResNumbers Then 'Numbers
    If REGEX.isNumberEntireAndDecimalVectorColonSeparated(StrData) Then
      Let vAux = VBA.Split(VBA.Replace(StrData, ".", ","), RXChar)
      For i = 0 To UBound(vAux)
        If REGEX.isNumberReg(vAux(i)) And Not REGEX.isEmptyStringReg(vAux(i)) Then
          If RBooIncludeZeroes Then
            ReDim Preserve VData(j)
            Let VData(j) = IIf(RBooAbs, VBA.Abs(VBA.CDbl(vAux(i))), VBA.CDbl(vAux(i)))
            Let j = j + 1
          Else
            If VBA.Abs(VBA.CDbl(vAux(i))) > 0 Then
              ReDim Preserve VData(j)
              Let VData(j) = IIf(RBooAbs, VBA.Abs(VBA.CDbl(vAux(i))), VBA.CDbl(vAux(i)))
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
  Set REGEX = Nothing
  
  Let RLngCounter = j
  Let ArrayGetDataXCharSeparatedStringFromVariantsAsString = VBA.Join(VData, RXChar)
  
EH:
  Let VData = Empty
  Let vAux = Empty
  Set REGEX = Nothing
  Call RES_ErrorHandler(sFN, vbInformation)
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
  Optional RStrDateFormat As String = VBA.vbNullString) As Variant

  Dim VData As Variant, VData1 As Variant
  Dim StrAux0 As String
  Dim i As Long
  
  Const sFN = "RES::ArrayGetDataXCharSeparatedStringFromVariantsAsVector"
  
  On Error GoTo EH
  
  Let StrAux0 = ArrayGetDataXCharSeparatedStringFromVariantsAsString(RVarVector, RLngCounter, RXChar, RValueType, RBooAbs, RBooIncludeZeroes, RBooIncludeEmptyStrings, RStrCustomRegex, RRngRC)
  
  If RLngCounter = 0 Then
    ReDim VData(0)
    Let VData(0) = VBA.vbNullString
    Let ArrayGetDataXCharSeparatedStringFromVariantsAsVector = VData
  Else
    Let VData1 = VBA.Split(StrAux0, RXChar) 'Strings
    If RValueAs = eqlResStrings Then
      Let VData = VData1
    ElseIf RValueAs = eqlResNumbers Then 'Numbers
      ReDim VData(UBound(VData1))
      For i = LBound(VData) To UBound(VData)
        Let VData(i) = VBA.CDbl(VData1(i))
      Next i
    ElseIf RValueAs = eqlResDates Then 'Dates
      ReDim VData(UBound(VData1))
      For i = LBound(VData) To UBound(VData)
        Let VData(i) = VBA.CDate((VBA.Format(VBA.CDate(VData1(i)), RStrDateFormat)))
      Next i
    End If
  End If
  
  'Return
  Let ArrayGetDataXCharSeparatedStringFromVariantsAsVector = VData
  
EH:
  Let VData = Empty
  Let VData1 = Empty
  Call RES_ErrorHandler(sFN, vbInformation)
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
  Call RES_ErrorHandler("RES::ArrayHasValue", vbInformation)

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
  Call RES_ErrorHandler("RES::ArrayHasValueN", vbInformation)

End Function

Public Function ArrayGetValuePosition( _
  ByVal AVector As Variant, _
  ByVal ANeedle As Variant) As Long
  
  Dim i As Long
  Dim StrAux0 As String, StrAux1 As String

  On Error GoTo EH
  Let ArrayGetValuePosition = -1
  
  If VBA.IsArray(AVector) Then
    Let StrAux1 = VBA.LCase(VBA.CStr(ANeedle))
    For i = LBound(AVector) To UBound(AVector)
      Let StrAux0 = VBA.LCase(VBA.CStr(AVector(i)))
      If StrAux0 = StrAux1 Then Let ArrayGetValuePosition = i: Exit For
    Next i
  End If

EH:
  Call RES_ErrorHandler("RES::ArrayGetValuePosition", vbInformation)

End Function

Public Function ArrayHasRepeatedElements(ByVal Vec As Variant) As Boolean
  Dim StrNeedle As String
  Dim i As Long, j As Long
  
  Const sFN As String = "RES::ArrayHasRepeatedElements"
  
  On Error GoTo EH
  Let ArrayHasRepeatedElements = False
  
  If Not VBA.IsArray(Vec) Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "¡Se requiere un array!")
  If (UBound(Vec) = 0) Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "¡Se requiere un array!")
  
  Let i = 1
  Let j = 1
  Let StrNeedle = VBA.CStr(Vec(0))
  Do While (ArrayHasRepeatedElements = False)
    Let ArrayHasRepeatedElements = (StrNeedle = VBA.CStr(Vec(i)))
    Let i = i + 1
    If i > UBound(Vec) Then
      Let StrNeedle = VBA.CStr(Vec(j))
      Let j = j + 1
      Let i = j
    End If
    If j - 1 = UBound(Vec) Then Exit Do
  Loop
EH:
  Call RES_ErrorHandler(sFN, vbInformation)

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
  Call RES_ErrorHandler(sFN, vbInformation)

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
  Call RES_ErrorHandler(sFN, vbInformation)

End Function

Public Function ArraySlice( _
  ByRef RVarSource As Variant, _
  ByRef RLngIndex As Long, _
  Optional RReturnSrc As Boolean = False, _
  Optional RByCols As Boolean = False) As Variant
  
  ' Slices the any array (1D or 2D) at index x and return the first part as array _
    and the VecSource sliced, both has base 0.
  Dim VecAux0 As Variant, VecAux1 As Variant
  Dim i As Long, j As Long, k As Long, l As Long, m As Long, xCol As Long, xRow  As Long, LngIndex As Long
  Dim EnuDimensionSrc As EQLRES_ENU_DIMENSIONARRAY

  Const sFN As String = "RES::ArraySlice"

  On Error GoTo EH

  If Not VBA.IsArray(RVarSource) Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "¡Se requiere un array!")
  Let i = Me.ArrayDimension(RVarSource)
  Let EnuDimensionSrc = IIf(i = 1, eqlRes1D, IIf(i = 2, eqlRes2D, eqlResNoArray))
  If i > 2 Then GoTo EH
  If EnuDimensionSrc = eqlResNoArray Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "¡Se requiere un array!")

  Let j = 0
  Let k = 0
  Let l = 0
  Let m = 0
  Let RLngIndex = VBA.Abs(RLngIndex)
  Select Case EnuDimensionSrc
    
    Case eqlRes1D
      If (LBound(RVarSource) = UBound(RVarSource)) Or RLngIndex < LBound(RVarSource) Or RLngIndex > UBound(RVarSource) Then
        ReDim VecAux0(0)
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

  Let ArraySlice = VecAux1
  If Not RReturnSrc Then
    Let RVarSource = Empty
    Let RVarSource = VecAux0
    Let VecAux1 = Empty
  End If

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Call RES_ErrorHandler(sFN, vbInformation)

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
  Call RES_ErrorHandler(sFN, vbInformation)

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
  Call RES_ErrorHandler(sFN, vbInformation)

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
  
  If Not VBA.IsArray(RVarSource) Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "¡Se requiere un array!")
  Let i = Me.ArrayDimension(RVarSource)
  Let EnuDimensionSrc = IIf(i = 1, eqlRes1D, IIf(i = 2, eqlRes2D, eqlResNoArray))
  If i > 2 Then GoTo EH
  If EnuDimensionSrc = eqlResNoArray Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "¡Se requiere un array!")
  
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
          'Let ArrayDelIndex = Empty
          'GoTo EH
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
          'Let ArrayDelIndex = Empty
          'GoTo EH
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
  Call RES_ErrorHandler(sFN, vbInformation)

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
        Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "¡Se requiere un array!")
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
        Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "¡Se requiere un valor u otro array para adicionar al array!")
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
        Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "¡Se requiere un valor u otro array para adicionar al array!")
      End If
    End If
    
    Let VecAux0 = Me.ArrayConcat(RVarSource, VecAux0, RLngPositionAt, RByCols)
  
  End If
  
  Let ArrayAddElement = VecAux0
  
EH:
  Let VecAux0 = Empty
  Call RES_ErrorHandler(sFN, vbInformation)
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
    Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "¡Se requiere un array base y un valor u otro array para adicionar al array base!")
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
  Call RES_ErrorHandler(sFN, vbInformation)

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
  
  If Not VBA.IsArray(RVarSource) Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "¡Se requiere un array!")
  Let i = Me.ArrayDimension(RVarSource)
  Let EnuDimension = IIf(i = 1, eqlRes1D, IIf(i = 2, eqlRes2D, eqlResNoArray))
  If EnuDimension = eqlResNoArray Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "¡Se requiere un array!")
  
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
  Call RES_ErrorHandler(sFN, vbInformation)

End Function

Public Function ArrayUniques(ByVal Vec As Variant) As Variant
  Dim VecAux As Variant, VecAux1 As Variant
  Dim StrNeedle As String
  Dim i As Long, j As Long, k As Long
  
  Const sFN As String = "RES::ArrayUniques"
  
  On Error GoTo EH
  
  If Not VBA.IsArray(Vec) Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "¡Se requiere un array!")
  If (UBound(Vec) = 0) Then Let ArrayUniques = Vec: GoTo EH
  
  Let i = 1
  Let VecAux = Vec
  
  For i = 0 To UBound(VecAux)
    Let StrNeedle = VBA.LCase(VBA.CStr(VecAux(i)))
    Let j = Me.ArrayHasValueN(Vec, StrNeedle, False, True)
    If j > 1 Then
      If i < UBound(VecAux) Then
        For k = i + 1 To UBound(VecAux)
          If StrNeedle = VBA.LCase(VBA.CStr(VecAux(k))) Then Let VecAux(k) = VBA.vbNullString
        Next k
      End If
    End If
  Next i
  
  Let j = 0
  ReDim VecAux1(0)
  For i = 0 To UBound(VecAux)
    If Not VBA.CStr(VecAux(i)) = VBA.vbNullString Then
      ReDim Preserve VecAux1(j)
      Let VecAux1(j) = VecAux(i)
      Let j = j + 1
    End If
  Next i
  
  If j = 0 Then GoTo EH
  Let ArrayUniques = VecAux1
  
EH:
  Let VecAux = Empty
  Let VecAux1 = Empty
  Call RES_ErrorHandler(sFN, vbInformation)

End Function

Public Function ArrayDimension(ByVal RVarSrc As Variant) As Long
  Dim i As Long, j As Long, xLength As Long
  
  Const sFN As String = "RES::ArrayDimension"
  
  On Error GoTo EH
  
  'Assuming that Src is not array
  Let xLength = -1
  If Not VBA.IsArray(RVarSrc) Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "¡Se requiere un array!")
  
  'If is array, calculate its dimensions
  Let i = 1
  Do While True
    Let j = (UBound(RVarSrc, i) - LBound(RVarSrc, i)) + 1
    Let xLength = i
    Let i = i + 1
  Loop

EH:
  Let ArrayDimension = xLength
  Call RES_ErrorHandler(sFN, vbInformation, True)

End Function


