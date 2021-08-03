VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ModelParent_current"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'============================= CLASS MODULES STRUCTURE =============================='
' OBJECT VARIABLES (GLOOBJ_) --------------------------------------------------------'
Private EHGLOBAL As AppErrorHandler
Private GLOSHT_DATA As Worksheet

' VARIABLES (GLOStr_) ---------------------------------------------------------------'

' VARIABLES CUSTOM (GLOEnu|Typ_) ----------------------------------------------------'
Public Enum EQLMDL_ENU_TABLE_AS
  eqlMdlArray
  eqlMdlRange
  eqlMdlListObject
End Enum

' CONSTANTS ERROR MESSAGES (GLOSTR_ERR_) --------------------------------------------'
Private Const CUSTOM_ERROR = VBA.vbObjectError + 519







' PROPERTIES ------------------------------------------------------------------------'
Public Property Set ErrorHandler(ByRef ObjEH As AppErrorHandler)
  
  Set EHGLOBAL = ObjEH

End Property

Public Property Set SheetData(ByVal ObjSheet As Worksheet)
  
  Set GLOSHT_DATA = ObjSheet

End Property

Public Property Get LengthRows(ByVal MStrTable As String) As Long
  
  Const sFN As String = "MDLPARENT::LengthRows"
  
  On Error GoTo EH
  If GLOSHT_DATA Is Nothing Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "Se requiere una fuente de datos.")
  Let LengthRows = GLOSHT_DATA.ListObjects(MStrTable).ListRows.Count

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Property

Public Property Get LengthCols(ByVal MStrTable As String) As Long

  Const sFN As String = "MDLPARENT::LengthCols"
  
  On Error GoTo EH
  If GLOSHT_DATA Is Nothing Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "Se requiere una fuente de datos.")
  Let LengthCols = GLOSHT_DATA.ListObjects(MStrTable).ListColumns.Count

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Property











' CONSTRUCTOR/DESTRUCTOR ------------------------------------------------------------'
Private Sub Class_Terminate()
  
  Set EHGLOBAL = Nothing
  Set GLOSHT_DATA = Nothing

End Sub







' CORE METHODS PRIVATE --------------------------------------------------------------'
Private Sub Model_check_for_datasheetbasetables(ByVal MVarDataTables As Variant)

  Dim LO As ListObject
  Dim i As Long
  
  On Error GoTo EH
  For i = LBound(MVarDataTables) To UBound(MVarDataTables)
    Set LO = GLOSHT_DATA.ListObjects(VBA.CStr(MVarDataTables(i)))
    If LO Is Nothing Then Call VBA.Err.Raise(CUSTOM_ERROR, "MDLPARENT", "Fallo al leer la tabla " & VBA.CStr(MVarDataTables(i)) & ". ¡No existe!")
    Set LO = Nothing
  Next i

EH:
  Set LO = Nothing
  Call EHGLOBAL.ErrorHandlerRaise("MDLPARENT::-check_for_datasheetbasetables")

End Sub


'CRUD
Private Sub Model_set(ByVal MStrTable As String, ByVal MVarData As Variant, ByVal MLngRows As Long)

  'Use MLngRows when desire insert empty rows
  Const sFN As String = "MDLPARENT::-set"
  
  On Error GoTo EH
  If GLOSHT_DATA Is Nothing Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "Se requiere una fuente de datos.")
  Call Model_table_add_row_from_array(GLOSHT_DATA.ListObjects(MStrTable), MVarData, MLngRows)

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Sub

Private Function Model_get( _
  ByVal MStrTable As String, _
  ByVal MEnuResAs As EQLMDL_ENU_TABLE_AS, _
  ByVal MStrField As String, _
  ByVal MLngIndex As Long, _
  ByVal MStrWhere As String) As Variant

  Dim RES As AppResources_current
  
  Dim LO As ListObject
  Dim VecData As Variant, VecHead As Variant, VecField As Variant
  Dim VecOne As Variant, VecResult As Variant, VecAux0 As Variant
  Dim StrTable As String, StrNameA As String, StrNameB As String, StrCriteria As String
  Dim i As Long, j As Long
  
  Const sFN As String = "MDLPARENT::-get"
  
  On Error GoTo EH
  If GLOSHT_DATA Is Nothing Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "Se requiere una fuente de datos.")
  Set LO = GLOSHT_DATA.ListObjects(MStrTable)
  
  Select Case MEnuResAs
    Case eqlMdlArray
      'Get headsrange
      Let VecHead = LO.HeaderRowRange.Value
      
      'Get data[index]
      If MLngIndex <= 0 Then
        Let VecData = LO.DataBodyRange.Value
      ElseIf MLngIndex > 0 Then
        If MLngIndex > LO.ListRows.Count Then Call VBA.Err.Raise(CUSTOM_ERROR, "MDLPARENT::-get", "El índice de la tabla es mayor al número de registros.")
        Let VecData = LO.ListRows(MLngIndex).Range.Value
      End If
      Set RES = New AppResources_current
      Set RES.ErrorHandler = EHGLOBAL
      Let VecHead = RES.ArrayToBaseZero(VecHead)
      Let VecData = RES.ArrayToBaseZero(VecData)
      Let VecResult = VecData
      
      'Filter data by fields (If MStrField exists)
      If Not (MStrField = VBA.vbNullString) Then
        Let VecResult = Empty
        Let VecField = RES.ArrayToBaseZero(VBA.Split(MStrField, ","))
        For j = LBound(VecField) To UBound(VecField) 'Filter by cols
          Let StrNameA = VBA.LCase(VBA.Trim(VBA.CStr(VecField(j))))
          For i = 0 To UBound(VecHead, 2)
            Let StrNameB = VBA.LCase(VBA.CStr(VecHead(0, i)))
            If StrNameA = StrNameB Then
              Let VecOne = RES.ArrayDelIndex(VecData, i, True, True)
              Let VecResult = RES.ArrayConcat(VecResult, VecOne, RByCols:=True)
              Exit For
            End If
          Next i
        Next j
      End If
      
      'Filter data by rows (If Where exists. At the moment just works with = operator)
      If Not (MStrWhere = VBA.vbNullString) Then
        Let VecAux0 = RES.ArrayToBaseZero(VBA.Split(MStrWhere, "="))
        Let StrNameA = VBA.LCase(VBA.Trim(VBA.CStr(VecAux0(0))))
        Let StrCriteria = VBA.LCase(VBA.Trim(VBA.CStr(VecAux0(1))))
        Let VecAux0 = VecResult
        Let VecResult = Empty
        For j = 0 To UBound(VecHead, 2) 'Filter by rows
          Let StrNameB = VBA.LCase(VBA.CStr(VecHead(0, j)))
          If StrNameA = StrNameB Then
            For i = 0 To UBound(VecData)
              Let StrNameB = VBA.LCase(VBA.Trim(VBA.CStr(VecData(i, j))))
              If StrNameB = StrCriteria Then
                Let VecOne = RES.ArrayDelIndex(VecAux0, i, True)
                Let VecResult = RES.ArrayConcat(VecResult, VecOne)
                'Exit For
              End If
            Next i
          End If
        Next j
      End If
      Let VecData = Empty
      Set RES = Nothing
      
      'Return data
      Let Model_get = VecResult
    
    Case eqlMdlRange
      Set Model_get = LO.DataBodyRange
      
    Case eqlMdlListObject
      Set Model_get = LO
  End Select

EH:
  Let VecAux0 = Empty
  Let VecOne = Empty
  Let VecResult = Empty
  Let VecHead = Empty
  Let VecField = Empty
  Let VecData = Empty
  Set LO = Nothing
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Sub Model_upd( _
  ByVal MStrTable As String, _
  ByVal MVarData As Variant, _
  ByVal MStrField As String, _
  ByVal MLngIndex As Long, _
  ByVal MStrWhere As String)

  Dim RES As AppResources_current
  
  Dim LO As ListObject
  Dim VecFields As Variant
  Dim VecAux0 As Variant
  Dim StrTable As String
  Dim i As Long, j As Long, k As Long, l As Long
  
  Const sFN As String = "MDLPARENT::-upd"
  
  On Error GoTo EH
  If GLOSHT_DATA Is Nothing Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "Se requiere una fuente de datos.")
  Set LO = GLOSHT_DATA.ListObjects(MStrTable)
  
  Let i = LO.ListRows.Count
  Let j = LO.ListColumns.Count
  Set RES = New AppResources_current
  Set RES.ErrorHandler = EHGLOBAL
  If Not (MStrField = VBA.vbNullString) Then 'Update Specific fields
    Let VecFields = VBA.Split(MStrField, ",")
    If MLngIndex <= 0 Then
      If Not (MStrWhere = VBA.vbNullString) Then
        Let VecAux0 = VBA.Split(MStrWhere, "=")
        Dim StrAux0 As String, StrAux1 As String, StrAux2 As String, StrAux3 As String
        Let StrAux0 = VBA.LCase(VBA.Trim(VBA.CStr(VecAux0(0)))) 'Field
        Let StrAux1 = VBA.LCase(VBA.Trim(VBA.CStr(VecAux0(1)))) 'Value
        'Search by field
        Let VecAux0 = RES.ArrayToBaseZero(LO.HeaderRowRange.Value)
        Let l = UBound(VecAux0, 2)
        For k = LBound(VecAux0, 2) To UBound(VecAux0, 2)
          Let StrAux3 = VBA.LCase(VBA.Trim(VBA.CStr(VecAux0(0, k)))) 'Table Field
          Let StrAux2 = VBA.CStr(VecAux0(0, k)) 'Table Field
          If StrAux3 = StrAux0 Then Exit For
        Next k
        If k > l Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "Índice incorrecto, El nombre: " & VBA.UCase(StrAux2) & " no existe entre los campos de la tabla.")
        'Search by row
        Let VecAux0 = RES.ArrayToBaseZero(LO.ListColumns(StrAux2).DataBodyRange.Value)
        Let l = UBound(VecAux0)
        For k = LBound(VecAux0) To UBound(VecAux0)
          Let StrAux3 = VBA.LCase(VBA.Trim(VBA.CStr(VecAux0(k, 0)))) 'Table row
          If StrAux3 = StrAux1 Then Exit For
        Next k
        If k > l Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "Índice incorrecto, El valor: " & VBA.UCase(StrAux1) & " no existe entre los registros de la tabla.")
        Let l = k + 1
        Let k = 0
        'Update Table
        Let VecAux0 = RES.ArrayToBaseZero(LO.HeaderRowRange.Value)
        For i = LBound(VecFields) To UBound(VecFields)
          Let StrAux0 = VBA.LCase(VBA.Trim(VBA.CStr(VecFields(i))))
          For j = LBound(VecAux0, 2) To UBound(VecAux0, 2)
            Let StrAux1 = VBA.LCase(VBA.Trim(VBA.CStr(VecAux0(0, j))))
            If StrAux0 = StrAux1 Then
              Let LO.ListColumns(VecAux0(0, j)).DataBodyRange.Cells(l, 1).Value = MVarData(k)
              Let k = k + 1
            End If
          Next j
        Next i
      End If
    Else
      If MLngIndex > i Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "Índice incorrecto, es mayor a la cantidad de registros de la tabla.")
      Let k = 0
      Let VecAux0 = RES.ArrayToBaseZero(LO.HeaderRowRange.Value)
      For i = LBound(VecFields) To UBound(VecFields)
        Let StrAux0 = VBA.LCase(VBA.Trim(VBA.CStr(VecFields(i))))
        For j = LBound(VecAux0, 2) To UBound(VecAux0, 2)
          Let StrAux1 = VBA.LCase(VBA.Trim(VBA.CStr(VecAux0(0, j))))
          If StrAux0 = StrAux1 Then
            Let LO.ListColumns(VecAux0(0, j)).DataBodyRange.Cells(MLngIndex, 1).Value = MVarData(k)
            Let k = k + 1
            Exit For
          End If
        Next j
      Next i
    End If
  Else 'Update entire row or entire table
    Let VecAux0 = RES.ArrayToBaseZero(MVarData)
    Let k = RES.ArrayDimension(VecAux0)
    If MLngIndex > 0 Then
      If Not k = 1 Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "La dimensión de los datos es incorrecta, no coincide con la dimensión de la tabla.")
      If Not j = ((UBound(VecAux0) - LBound(VecAux0)) + 1) Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "Números de campos de datos incorrecto, no coincide con la cantidad de campos de la tabla.")
      Let LO.ListRows(MLngIndex) = MVarData
    Else
      If Not k = 2 Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "La dimensión de los datos es incorrecta, no coincide con la dimensión de la tabla.")
      If Not j = ((UBound(VecAux0, 2) - LBound(VecAux0, 2)) + 1) Or Not i = ((UBound(VecAux0) - LBound(VecAux0)) + 1) Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "Dimensión de datos incorrecto, no coincide con la dimensión de la tabla.")
      Let LO.DataBodyRange = MVarData
    End If
  End If

EH:
  Let VecAux0 = Empty
  Let VecFields = Empty
  Set LO = Nothing
  Set RES = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Sub

Private Sub Model_del(ByVal MStrTable As String, ByVal MLngIndex As Long)

  Dim LO As ListObject
  Dim StrTable As String
  
  Const sFN As String = "MDLPARENT::-del"
  
  On Error GoTo EH
  If GLOSHT_DATA Is Nothing Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "Se requiere una fuente de datos.")
  Set LO = GLOSHT_DATA.ListObjects(MStrTable)
  
  If MLngIndex <= 0 Then
    Call LO.DataBodyRange.Delete
  Else
    If MLngIndex > LO.ListRows.Count Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "Índice incorrecto, es mayor a la cantidad de registros de la tabla.")
    Call LO.ListRows(MLngIndex).Delete
  End If

EH:
  Set LO = Nothing
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Sub





' ASSETS ----------------------------------------------------------------------------'
'LO's mannagement
Private Sub Model_table_add_row_from_array(ByVal MLOTable As ListObject, ByRef MVarData As Variant, ByVal MLngRows As Long)
  
  'Method Used as CREATE and UPDATE from CRUD, inserts data 1 to n rows of any table. Id data is empty inserts a empty row
  Dim RES As AppResources_current
  
  Dim TableRange As Range 'MODIFIED AND OPTIMIZED 30/04/21
  Dim i As Long, j As Long, k As Long

  Const sFN As String = "MDLPARENT::-table_add_row_from_array"
  
  On Error GoTo EH
  If VBA.IsArray(MVarData) Then
    Set RES = New AppResources_current
    Set RES.ErrorHandler = EHGLOBAL
    Let MVarData = RES.ArrayToBaseZero(MVarData)
    Let i = RES.ArrayDimension(MVarData)
    If i = 1 Then
      'Let i = 1 'MVarData would be a 1D array base 0
      If Not MLOTable.ListColumns.Count = ((UBound(MVarData) - LBound(MVarData)) + 1) Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "Dimensión de datos incompatible con dimensión de tabla.")
    Else
      Let i = UBound(MVarData) 'MVarData must be an 2D array base 1
      If Not MLOTable.ListColumns.Count = ((UBound(MVarData, 2) - LBound(MVarData, 2)) + 1) Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "Dimensión de datos incompatible con dimensión de tabla.")
    End If
  
    If (MLOTable.ListRows.Count <= 0 And i <= 1) Then
      Let MLOTable.ListRows.Add(AlwaysInsert:=True).Range.Value = MVarData
    Else
      Let j = IIf(MLOTable.ListRows.Count = 0, 2, MLOTable.Range.Rows.Count + 1)
      Let k = IIf(MLOTable.ListRows.Count = 0, i - 1, i)
      Set TableRange = MLOTable.Range
      Call MLOTable.Resize(TableRange.Resize(TableRange.Rows.Count + k, TableRange.Columns.Count))
      Let TableRange.Range("A" & j).Resize(i, TableRange.Columns.Count).Value = MVarData
    End If
  Else
    Call Model_table_add_rows_empties(MLOTable, MLngRows)
  End If

EH:
  Set TableRange = Nothing
  Call EHGLOBAL.ErrorHandlerRaise("MDLPARENT::AddRowFromArray", vbInformation)

End Sub

Private Sub Model_table_add_rows_empties(ByVal MLOTable As ListObject, ByVal MLngRows As Long)
  
  'Method Used as CREATE from CRUD, inserts empty rows to any table
  Dim TableRange As Range 'CREATED AT 04/05/21
  Dim i As Long, j As Long, k As Long

  On Error GoTo EH
  Let i = MLngRows

  If (MLOTable.ListRows.Count <= 0 And i <= 1) Then
    Call MLOTable.ListRows.Add(AlwaysInsert:=True)
  Else
    Let j = IIf(MLOTable.ListRows.Count = 0, 2, MLOTable.Range.Rows.Count + 1)
    Let k = IIf(MLOTable.ListRows.Count = 0, i - 1, i)
    Set TableRange = MLOTable.Range
    Call MLOTable.Resize(TableRange.Resize(TableRange.Rows.Count + k, TableRange.Columns.Count))
  End If

EH:
  Set TableRange = Nothing
  Call EHGLOBAL.ErrorHandlerRaise("MDLPARENT::TableAddRowsEmpties", vbInformation)

End Sub






' ERRORS MANNAGEMENT AND ASSERTSIONS ------------------------------------------------'
Private Sub Debug_assert(ByVal Cond As Boolean, Optional Origin As String = VBA.vbNullString)
  #If Debbuging Then
    Debug.Assert Cond
    If Not Cond Then Debug.Print "Assert on: " & Origin
  #End If
End Sub
' CORE METHODS PRIVATE --------------------------------------------------------------'







' METHODS PUBLIC --------------------------------------------------------------------'
' CRUD
Public Sub MSet(ByVal MStrTable As String, Optional MVarData As Variant = Empty, Optional MLngRows As Long = 0)

  On Error GoTo EH
  Call Model_set(MStrTable, MVarData, MLngRows)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("MDLPARENT::MSet")

End Sub

Public Function MGet( _
  ByVal MStrTable As String, _
  ByVal MEnuResAs As EQLMDL_ENU_TABLE_AS, _
  Optional MStrField As String = VBA.vbNullString, _
  Optional MLngIndex As Long = 0, _
  Optional MStrWhere As String = VBA.vbNullString) As Variant

  On Error GoTo EH
  Select Case MEnuResAs
    Case eqlMdlArray: Let MGet = Model_get(MStrTable, MEnuResAs, MStrField, MLngIndex, MStrWhere)
    Case eqlMdlRange, eqlMdlListObject: Set MGet = Model_get(MStrTable, MEnuResAs, MStrField, MLngIndex, MStrWhere)
    Case Else: Call VBA.Err.Raise(CUSTOM_ERROR, "MDLPARENT::MGet")
  End Select

EH:
  Call EHGLOBAL.ErrorHandlerRaise("MDLPARENT::MGet")

End Function

Public Sub MUpd( _
  ByVal MStrTable As String, _
  ByVal MVarData As Variant, _
  Optional MStrField As String = VBA.vbNullString, _
  Optional MLngIndex As Long = 0, _
  Optional MStrWhere As String = VBA.vbNullString)

  On Error GoTo EH
  Call Model_upd(MStrTable, MVarData, MStrField, MLngIndex, MStrWhere)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("MDLPARENT::MUpd")

End Sub

Public Sub MDel(ByVal MStrTable As String, Optional MLngRow As Long = 0)

  On Error GoTo EH
  Call Model_del(MStrTable, MLngRow)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("MDLPARENT::MDel")

End Sub

Public Sub MCheckForDatasheetbasetables(ByVal MVarDataTables As Variant)

  On Error GoTo EH
  Call Model_check_for_datasheetbasetables(MVarDataTables)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("MDLPARENT::MCheckForDatasheetbasetables")

End Sub
' METHODS PUBLIC --------------------------------------------------------------------'
'============================= CLASS MODULES STRUCTURE =============================='



