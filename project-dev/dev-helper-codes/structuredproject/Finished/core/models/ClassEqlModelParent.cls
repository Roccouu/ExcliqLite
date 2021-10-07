VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassEqlModelParent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==================== CLASS EXCLIQ LITE MODEL PARENT STRUCTURE ==================== '
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


' ==================== CLASS EXCLIQ LITE MODEL PARENT STRUCTURE ==================== '
' OBJECT VARIABLES (GLOOBJ_) ------------------------------------------------------- '
' PROPERTIES ----------------------------------------------------------------------- '
' CONSTRUCTOR/DESTRUCTOR ----------------------------------------------------------- '
' CORE METHODS LIST (PRIVATE) ------------------------------------------------------ '
' Model_set
' Model_get
' Model_upd
' Model_del
' ASSETS --------------------------------------------------------------------------- '
' Model_table_add_row_from_array
' ERROR ASSERTSIONS ---------------------------------------------------------------- '
' Debug_assert
' CORE METHODS LIST (PUBLIC) ------------------------------------------------------- '
' MSet
' MGet
' MUpd
' MDel
' ==================== CLASS EXCLIQ LITE MODEL PARENT STRUCTURE ==================== '



' ==================== CLASS EXCLIQ LITE MODEL PARENT STRUCTURE ==================== '
' OBJECT VARIABLES (GLOOBJ_) ------------------------------------------------------- '
Private RES As ClassEqlResources



' PROPERTIES ----------------------------------------------------------------------- '
Public Property Set ResourcesApp(ByRef ObjRES As ClassEqlResources)

  Set RES = ObjRES

End Property

Public Property Get LengthRows(ByVal MLOTable As ListObject) As Long

  Const sFN As String = "MDLPARENT::LengthRows"

  On Error GoTo EH
  Let LengthRows = MLOTable.ListRows.Count

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Property

Public Property Get LengthCols(ByVal MLOTable As ListObject) As Long

  Const sFN As String = "MDLPARENT::LengthCols"

  On Error GoTo EH
  Let LengthCols = MLOTable.ListColumns.Count

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Property



' CONSTRUCTOR/DESTRUCTOR ----------------------------------------------------------- '
Private Sub Class_Terminate()

  Set RES = Nothing

End Sub



' CORE METHODS LIST (PRIVATE) ------------------------------------------------------ '
' CRUD
Private Sub Model_set(ByVal MLOTable As ListObject, ByVal MVarData As Variant, ByVal MLngRows As Long, ByVal MBooEmpties As Boolean)

  ' Use MLngRows when desire insert empty rows
  Const sFN As String = "MDLPARENT::-set"

  On Error GoTo EH
  Call Model_table_add_row_from_array(MLOTable, MVarData, MLngRows, MBooEmpties)

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Sub

Private Function Model_get( _
  ByVal MLOTable As ListObject, _
  ByVal MEnuResAs As EQLMDL_ENU_TABLE_AS, _
  ByVal MStrField As String, _
  ByVal MLngIndex As Long, _
  ByVal MStrWhere As String) As Variant

  Dim VecData As Variant, VecHead As Variant, VecField As Variant
  Dim VecOne As Variant, VecResult As Variant, VecAux0 As Variant
  Dim StrNameA As String, StrNameB As String, StrCriteria As String
  Dim i As Long, j As Long

  Const sFN As String = "MDLPARENT::-get"

  On Error GoTo EH
  With MLOTable
    If .DataBodyRange Is Nothing Then GoTo EH
    Select Case MEnuResAs

      Case eqlMdlArray
        ' Get headsrange
        Let VecHead = .HeaderRowRange.Value

        ' Get data[index]
        If MLngIndex <= 0 Then
          Let VecData = .DataBodyRange.Value
        ElseIf MLngIndex > 0 Then
          If MLngIndex > .ListRows.Count Then Call VBA.Err.Raise(CUSTOM_ERROR_MDL, "MDLPARENT::-get", "El índice de la tabla es mayor al número de registros.")
          Let VecData = .ListRows(MLngIndex).Range.Value
        End If
        Let VecHead = RES.ArrayToBaseZero(VecHead)
        Let VecData = RES.ArrayToBaseZero(VecData)
        Let VecResult = VecData

        ' Filter data by fields (If MStrField exists)
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

        ' Filter data by rows (If Where exists. At the moment just works with = operator)
        If Not (MStrWhere = VBA.vbNullString) Then
          Let VecAux0 = RES.ArrayToBaseZero(VBA.Split(MStrWhere, "="))
          Let StrNameA = VBA.LCase(VBA.Trim(VBA.CStr(VecAux0(0))))
          Let StrCriteria = VBA.LCase(VBA.Trim(VBA.CStr(VecAux0(1))))
          Let VecAux0 = VecResult
          Let VecResult = Empty
          For j = 0 To UBound(VecHead, 2) ' Filter by rows
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

        ' Return data
        Let Model_get = VecResult

      Case eqlMdlRange
        Set Model_get = .DataBodyRange

      Case eqlMdlListObject
        Set Model_get = MLOTable

    End Select
  End With

EH:
  Let VecAux0 = Empty
  Let VecOne = Empty
  Let VecResult = Empty
  Let VecHead = Empty
  Let VecField = Empty
  Let VecData = Empty
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Function

Private Sub Model_upd( _
  ByVal MLOTable As ListObject, _
  ByVal MVarData As Variant, _
  ByVal MStrField As String, _
  ByVal MLngIndex As Long, _
  ByVal MStrWhere As String)

  Dim VecFields As Variant, VecAux0 As Variant, VecAux1 As Variant
  Dim StrAux0 As String, StrAux1 As String, StrAux2 As String, StrAux3 As String
  Dim i As Long, j As Long, k As Long, l As Long

  Const sFN As String = "MDLPARENT::-upd"

  On Error GoTo EH

  Call Debug_assert(Not MLOTable Is Nothing, sFN)

  With MLOTable
    If .DataBodyRange Is Nothing Then GoTo EH
    Let i = .ListRows.Count
    Let j = .ListColumns.Count
    If Not (MStrField = VBA.vbNullString) Then ' Update Specific fields
      Let VecFields = VBA.Split(MStrField, ",")
      If Not (UBound(VecFields) - LBound(VecFields)) = (UBound(MVarData) - LBound(MVarData)) Then Call VBA.Err.Raise(CUSTOM_ERROR_MDL, sFN, "Cantidad de campos y valores incorrecta, se requiere la misma cantidad pares campo=valor."): GoTo EH
      If MLngIndex <= 0 Then ' Update specific row by conditional or all rows by cols

        If MStrWhere = VBA.vbNullString Then ' Conditional doesn't exists, update all column

          ' Update Table by entire columns with just one value or a 2D vector (by cols)
          If i > 0 Then ' Required min 1 row
            Let VecAux0 = RES.ArrayToBaseZero(.HeaderRowRange.Value, eqlRes1D)
            For i = LBound(VecFields) To UBound(VecFields)
              Let StrAux0 = VBA.LCase(VBA.Trim(VBA.CStr(VecFields(i))))
              Let j = RES.ArrayGetValuePosition(VecAux0, StrAux0) + 1
              If j > 0 Then
                Let VecAux1 = MVarData(i)
                If VBA.IsArray(VecAux1) Then
                  Let k = RES.ArrayDimension(VecAux1)
                  If k > 2 Or k <= 0 Then Call VBA.Err.Raise(CUSTOM_ERROR_MDL, sFN, "Dimensión incorrecta de vector.")
                  If k = 1 Then
                    Let VecAux1 = RES.ArrayToBaseZero(VecAux1, eqlRes2D, False)
                  Else
                    If UBound(VecAux1) > 0 Then Let VecAux1 = RES.ArraySlice(VecAux1, 0, RByCols:=True, RBooGetLeftSide:=True)
                  End If
                  Let .ListColumns(j).DataBodyRange.Value = VecAux1
                  Let VecAux1 = Empty
                Else
                  Let .ListColumns(j).DataBodyRange.Value = MVarData(i)
                End If
              End If
            Next i
          End If

        Else ' Conditional exists: WHERE A = B

          Let VecAux0 = VBA.Split(MStrWhere, "=")
          Let StrAux0 = VBA.LCase(VBA.Trim(VBA.CStr(VecAux0(0)))) ' Field
          Let StrAux1 = VBA.LCase(VBA.Trim(VBA.CStr(VecAux0(1)))) ' Value
          ' Search by field
          Let VecAux0 = RES.ArrayToBaseZero(.HeaderRowRange.Value, eqlRes1D)
          Let k = RES.ArrayGetValuePosition(VecAux0, StrAux0) + 1 ' Table Field
          If k = -1 Then Call VBA.Err.Raise(CUSTOM_ERROR_MDL, sFN, "Índice incorrecto, El nombre: " & VBA.UCase(StrAux2) & " no existe entre los campos de la tabla.")
          ' Search by row
          Let VecAux1 = RES.ArrayToBaseZero(.ListColumns(k).DataBodyRange.Value, eqlRes1D) ' Value
          Let k = RES.ArrayGetValuePosition(VecAux1, StrAux1) + 1 ' Row position
          If k = -1 Then Call VBA.Err.Raise(CUSTOM_ERROR_MDL, sFN, "Índice incorrecto, El valor: " & VBA.UCase(StrAux1) & " no existe entre los registros de la tabla.")
          ' Update just one row in Table
          For i = LBound(VecFields) To UBound(VecFields)
            Let j = RES.ArrayGetValuePosition(VecAux0, VBA.LCase(VBA.Trim(VBA.CStr(VecFields(i))))) + 1
            If j > 0 Then Let .ListColumns(j).DataBodyRange.Cells(k, 1).Value = MVarData(i)
          Next i

        End If

      Else ' Update a row by index

        If MLngIndex > i Then Call VBA.Err.Raise(CUSTOM_ERROR_MDL, sFN, "Índice incorrecto, es mayor a la cantidad de registros de la tabla.")
        Let VecAux0 = RES.ArrayToBaseZero(.HeaderRowRange.Value, eqlRes1D)
        For i = LBound(VecFields) To UBound(VecFields)
          Let j = RES.ArrayGetValuePosition(VecAux0, VBA.LCase(VBA.Trim(VBA.CStr(VecFields(i))))) + 1
          If j > 0 Then
            ' When MVarData has single value, updates a cell, when there are a vector, updates with entire vector's size starting in the first target cell
            If VBA.IsArray(MVarData(i)) Then
              Let k = RES.ArrayDimension(MVarData(i))
              Select Case k ' Vector 1D: Update the current row. Vector 2D updates from current cell to vector's size (row & col).
                Case 1: Let .ListColumns(j).DataBodyRange.Range(MLngIndex, 1).Resize(1, UBound(MVarData(i)) + 1).Value = MVarData(i)
                Case 2: Let .ListColumns(j).DataBodyRange.Range(MLngIndex, 1).Resize(UBound(MVarData(i)) + 1, UBound(MVarData(i), 2) + 1).Value = MVarData(i)
              End Select
            Else
              Let .ListColumns(j).DataBodyRange.Cells(MLngIndex, 1).Value = MVarData(i)
            End If
          End If
        Next i

      End If

    Else ' Fields names are unavaiable, then: Update entire row or entire table

      Let VecAux0 = RES.ArrayToBaseZero(MVarData)
      Let k = RES.ArrayDimension(VecAux0)
      If MLngIndex > 0 Then
        If Not k = 1 Then Call VBA.Err.Raise(CUSTOM_ERROR_MDL, sFN, "La dimensión de los datos es incorrecta, no coincide con la dimensión de la tabla.")
        If Not j = ((UBound(VecAux0) - LBound(VecAux0)) + 1) Then Call VBA.Err.Raise(CUSTOM_ERROR_MDL, sFN, "Números de campos de datos incorrecto, no coincide con la cantidad de campos de la tabla.")
        Let .ListRows(MLngIndex).Range.Value = MVarData
      Else
        If Not k = 2 Then Call VBA.Err.Raise(CUSTOM_ERROR_MDL, sFN, "La dimensión de los datos es incorrecta, no coincide con la dimensión de la tabla.")
        If Not j = ((UBound(VecAux0, 2) - LBound(VecAux0, 2)) + 1) Or Not i = ((UBound(VecAux0) - LBound(VecAux0)) + 1) Then Call VBA.Err.Raise(CUSTOM_ERROR_MDL, sFN, "Dimensión de datos incorrecto, no coincide con la dimensión de la tabla.")
        Let .DataBodyRange.Value = MVarData
      End If

    End If
  End With

EH:
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Let VecFields = Empty
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Sub

Private Sub Model_del(ByVal MLOTable As ListObject, ByVal MLngIndex As Long)

  Const sFN As String = "MDLPARENT::-del"

  On Error GoTo EH

  With MLOTable
    If MLngIndex <= 0 Then
      If .ListRows.Count > 0 Then
        Call .DataBodyRange.ClearContents
        Call .DataBodyRange.Delete
      End If
    Else
      If MLngIndex > .ListRows.Count Then Call VBA.Err.Raise(CUSTOM_ERROR_MDL, sFN, "Índice incorrecto, es mayor a la cantidad de registros de la tabla.")
      Call .ListRows(MLngIndex).Delete
    End If
  End With

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise(sFN)

End Sub



' ASSETS --------------------------------------------------------------------------- '
' LO's mannagement
Private Sub Model_table_add_row_from_array(ByVal MLOTable As ListObject, ByRef MVarData As Variant, ByVal MLngRows As Long, ByVal MBooEmpties As Boolean)

  ' Method Used as CREATE and UPDATE from CRUD, inserts data 1 to n rows of any table. Id data is empty inserts a empty row
  Dim i As Long, j As Long ' MODIFIED AND OPTIMIZED 30/04/21

  Const sFN As String = "MDLPARENT::-table_add_row_from_array"

  On Error GoTo EH
  If VBA.IsArray(MVarData) Then

    ' Check for dimension and columns of table and data to entry
    Let MVarData = RES.ArrayToBaseZero(MVarData)
    Let i = RES.ArrayDimension(MVarData)
    If i = 1 Then ' MVarData would be a 1D array base 0
      If Not MLOTable.ListColumns.Count = ((UBound(MVarData) - LBound(MVarData)) + 1) Then Call VBA.Err.Raise(CUSTOM_ERROR_MDL, sFN, "Dimensión de datos incompatible con dimensión de tabla.")
    Else
      Let i = UBound(MVarData) + 1 ' MVarData must be an 2D array base 1
      If Not MLOTable.ListColumns.Count = ((UBound(MVarData, 2) - LBound(MVarData, 2)) + 1) Then Call VBA.Err.Raise(CUSTOM_ERROR_MDL, sFN, "Dimensión de datos incompatible con dimensión de tabla.")
    End If

    ' Set data entry to table
    With MLOTable
      If (.ListRows.Count <= 0 And i <= 1) Then
        Let .ListRows.Add(AlwaysInsert:=True).Range.Value = MVarData
      Else
        Let j = .ListRows.Count
        Call .Resize(.Range.Resize(.ListRows.Count + IIf(.HeaderRowRange Is Nothing, 0, 1) + i, .ListColumns.Count))
        Let .Range.Range("A" & j + IIf(.HeaderRowRange Is Nothing, 1, 2)).Resize(i, .ListColumns.Count).Value = MVarData
      End If
    End With

  Else

    If MBooEmpties Then
      With MLOTable
        If (.ListRows.Count <= 0 And MLngRows <= 1) Then
          Call .ListRows.Add(AlwaysInsert:=True)
        Else
          Call .Resize(.Range.Resize(.ListRows.Count + IIf(.HeaderRowRange Is Nothing, 0, 1) + MLngRows, .ListColumns.Count))
        End If
      End With
    End If

  End If

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise("MDLPARENT::AddRowFromArray", vbInformation)

End Sub



' ERRORS MANNAGEMENT AND ASSERTSIONS ----------------------------------------------- '
Private Sub Debug_assert(ByVal Cond As Boolean, Optional Origin As String = VBA.vbNullString)

  #If Debugging Then
    Debug.Assert Cond
    If Not Cond Then Debug.Print "Assert on: " & Origin
  #End If

End Sub
' CORE METHODS LIST (PRIVATE) ------------------------------------------------------ '



' CORE METHODS LIST (PUBLIC) ------------------------------------------------------- '
' CRUD
Public Sub MSet(ByVal MLOTable As ListObject, Optional MVarData As Variant = Empty, Optional MLngRows As Long = 0, Optional MBooEmpties As Boolean = False)

  On Error GoTo EH
  Call Model_set(MLOTable, MVarData, MLngRows, MBooEmpties)

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise("MDLPARENT::MSet")

End Sub

Public Function MGet( _
  ByVal MLOTable As ListObject, _
  ByVal MEnuResAs As EQLMDL_ENU_TABLE_AS, _
  Optional MStrField As String = VBA.vbNullString, _
  Optional MLngIndex As Long = 0, _
  Optional MStrWhere As String = VBA.vbNullString) As Variant

  On Error GoTo EH
  Select Case MEnuResAs
    Case eqlMdlArray: Let MGet = Model_get(MLOTable, MEnuResAs, MStrField, MLngIndex, MStrWhere)
    Case eqlMdlRange, eqlMdlListObject: Set MGet = Model_get(MLOTable, MEnuResAs, MStrField, MLngIndex, MStrWhere)
    Case eqlMdlStrTableName: Let MGet = MLOTable.Parent.Name & "!" & MLOTable.Name
    Case Else: Call VBA.Err.Raise(CUSTOM_ERROR_MDL, "MDLPARENT::MGet")
  End Select

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise("MDLPARENT::MGet")

End Function

Public Sub MUpd( _
  ByVal MLOTable As ListObject, _
  ByVal MVarData As Variant, _
  Optional MStrField As String = VBA.vbNullString, _
  Optional MLngIndex As Long = 0, _
  Optional MStrWhere As String = VBA.vbNullString)

  On Error GoTo EH
  Call Model_upd(MLOTable, MVarData, MStrField, MLngIndex, MStrWhere)

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise("MDLPARENT::MUpd")

End Sub

Public Sub MDel(ByVal MLOTable As ListObject, Optional MLngRow As Long = 0)

  On Error GoTo EH
  Call Model_del(MLOTable, MLngRow)

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise("MDLPARENT::MDel")

End Sub
' CORE METHODS LIST (PUBLIC) ------------------------------------------------------- '
' ==================== CLASS EXCLIQ LITE MODEL PARENT STRUCTURE ==================== '

