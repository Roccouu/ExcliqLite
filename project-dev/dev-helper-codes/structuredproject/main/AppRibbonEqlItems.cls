VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RibbonExcliqItems_current"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' OBJECT VARIABLES (GLOOBJ_) --------------------------------------------------------'
Private EHGLOBAL As AppErrorHandler

' VARIABLES (GLOStr_) ---------------------------------------------------------------'
Private GLOBOO_EH As Boolean
Private GLORIBBON_ITEM As Variant

' CONSTANTS ERRORS AND ERROR MESSAGES (GLOSTR_ERR_) ---------------------------------'
Private Const CUSTOM_ERROR As Long = VBA.vbObjectError + 514





'=========================== CLASS RIBBONITEMS STRUCTURE ============================'
' OBJECT VARIABLES (GLOOBJ_) --------------------------------------------------------'
' VARIABLES (GLOStr_) ---------------------------------------------------------------'
' CONSTANTS ERRORS AND ERROR MESSAGES (GLOSTR_ERR_) ---------------------------------'
' CONSTANTS (GLOSTR_) ---------------------------------------------------------------'
' PROPERTIES ------------------------------------------------------------------------'
' CONSTRUCTOR/DESTRUCTOR ------------------------------------------------------------'
' CORE METHODS LIST PRIVATE ---------------------------------------------------------'
' RIBBONITEM_Core_initialize
' RIBBONITEM_Core_get
' RIBBONITEM_Core_set
' RIBBONITEM_Core_enabling
' ASSETS ----------------------------------------------------------------------------'
' ERRORS MANNAGEMENT AND ASSERTSIONS ------------------------------------------------'
' Debug_assert
' INTERFACE METHODS LIST ------------------------------------------------------------'
' RIBBONITEMINITIAL
' RIBBONITEMGET
' RIBBONITEMSET
' RIBBONITEMENABLINGALL
'=========================== CLASS RIBBONITEMS STRUCTURE ============================'


' PROPERTIES ------------------------------------------------------------------------'
Public Property Set PROPObjEH(ByRef ObjEH As AppErrorHandler)
  Set EHGLOBAL = Nothing
  Set EHGLOBAL = ObjEH
  Let GLOBOO_EH = (Not ObjEH Is Nothing)
End Property

Public Property Get PROPVecItems() As Variant
  Let PROPVecItems = GLORIBBON_ITEM
End Property







' CONSTRUCTOR/DESTRUCTOR ------------------------------------------------------------'
Private Sub Class_Initialize()
  Set EHGLOBAL = New AppErrorHandler
End Sub

Private Sub Class_Terminate()
  If GLOBOO_EH = False Then Set EHGLOBAL = Nothing
  Let GLORIBBON_ITEM = Empty
End Sub







' CORE METHODS LIST PRIVATE ---------------------------------------------------------'
Private Function RIBBONITEM_Core_initialize() As Boolean
  Const sFN As String = "RIBBONITEM::-core_initialize"
  
  On Error GoTo EH
  Let RIBBONITEM_Core_initialize = False
  
  Let GLORIBBON_ITEM = Empty
  ReDim GLORIBBON_ITEM(0 To 16, 0 To 1)
  Let GLORIBBON_ITEM(0, 0) = ("regression")
  Let GLORIBBON_ITEM(1, 0) = ("balmet")
  Let GLORIBBON_ITEM(2, 0) = ("regressionselection")
  Let GLORIBBON_ITEM(3, 0) = ("balmetselection")
  Let GLORIBBON_ITEM(4, 0) = ("balmeteco") 'NEW
  Let GLORIBBON_ITEM(5, 0) = ("projection")
  Let GLORIBBON_ITEM(6, 0) = ("oreinvoice")
  Let GLORIBBON_ITEM(7, 0) = ("concentrateinvoice")
  Let GLORIBBON_ITEM(8, 0) = ("configtabletype") 'NEW
  Let GLORIBBON_ITEM(9, 0) = ("configtablestyle") 'NEW
  Let GLORIBBON_ITEM(10, 0) = ("configformulas") 'NEW
  Let GLORIBBON_ITEM(11, 0) = ("configcellforselections") 'NEW
  Let GLORIBBON_ITEM(12, 0) = ("configtablerowshadows") 'NEW
  Let GLORIBBON_ITEM(13, 0) = ("configlauncher") 'NEW
  Let GLORIBBON_ITEM(14, 0) = ("help")
  Let GLORIBBON_ITEM(15, 0) = ("about")
  Let GLORIBBON_ITEM(16, 0) = ("license")

  Let GLORIBBON_ITEM(0, 1) = (True)
  Let GLORIBBON_ITEM(1, 1) = (True)
  Let GLORIBBON_ITEM(2, 1) = (True)
  Let GLORIBBON_ITEM(3, 1) = (True)
  Let GLORIBBON_ITEM(4, 1) = (True) 'NEW
  Let GLORIBBON_ITEM(5, 1) = (True)
  Let GLORIBBON_ITEM(6, 1) = (True)
  Let GLORIBBON_ITEM(7, 1) = (True)
  Let GLORIBBON_ITEM(8, 1) = (True) 'NEW
  Let GLORIBBON_ITEM(9, 1) = (True) 'NEW
  Let GLORIBBON_ITEM(10, 1) = (True) 'NEW
  Let GLORIBBON_ITEM(11, 1) = (True) 'NEW
  Let GLORIBBON_ITEM(12, 1) = (True) 'NEW
  Let GLORIBBON_ITEM(13, 1) = (True) 'NEW
  Let GLORIBBON_ITEM(14, 1) = (True)
  Let GLORIBBON_ITEM(15, 1) = (True)
  Let GLORIBBON_ITEM(16, 1) = (True)

  Let RIBBONITEM_Core_initialize = True

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)
End Function

Private Function RIBBONITEM_Core_get(ByVal RItem As String) As Variant
  
  Dim BooFound As Boolean
  Dim i As Long

  Const sFN As String = "RIBBONITEM::-core_get"
  
  On Error GoTo EH
  Let BooFound = False
  For i = LBound(GLORIBBON_ITEM) To UBound(GLORIBBON_ITEM)
    If VBA.LCase(RItem) = VBA.LCase(VBA.CStr(GLORIBBON_ITEM(i, 0))) Then
      Let RIBBONITEM_Core_get = GLORIBBON_ITEM(i, 1)
      Let BooFound = True
    End If
  Next i
  
  If Not BooFound Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "Ítem: " & RItem & " de la cinta de opciones inválido.")

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Function

Private Sub RIBBONITEM_Core_set(ByVal RItem As String, ByVal RValue As Variant)
  Dim BooFound As Boolean
  Dim i As Long

  Const sFN As String = "RIBBONITEM::-core_set"
  
  On Error GoTo EH
  Let BooFound = False
  For i = 0 To UBound(GLORIBBON_ITEM)
    If VBA.LCase(RItem) = VBA.LCase(GLORIBBON_ITEM(i, 0)) Then
      Let GLORIBBON_ITEM(i, 1) = RValue
      Let BooFound = True
    End If
  Next i
  
  If BooFound = False Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "No se encontró el ítem: " & RItem & " de la cinta de opciones")
EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)
End Sub

Private Sub RIBBONITEM_Core_enabling(ByVal RValue As Boolean)
  Dim i As Long

  Const sFN As String = "RIBBONITEM::-core_enabling"
  
  On Error GoTo EH
  For i = 0 To UBound(GLORIBBON_ITEM)
    Let GLORIBBON_ITEM(i, 1) = RValue
  Next i
  
EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)
End Sub






' ASSETS ----------------------------------------------------------------------------'
' ERRORS MANNAGEMENT AND ASSERTSIONS ------------------------------------------------'
' Debug_assert
' INTERFACE METHODS LIST ------------------------------------------------------------'
Public Function RIBBONITEMINITIAL() As Boolean
  'Sets the initial values for stand by Ribbon
  On Error GoTo EH
  Let RIBBONITEMINITIAL = RIBBONITEM_Core_initialize()

EH:
  If GLOBOO_EH Then
    Call EHGLOBAL.ErrorHandlerRaise("RIBBONITEM::RIBBONITEMINITIAL")
  Else
    Call EHGLOBAL.ErrorHandlerDisplay("RIBBONITEM::RIBBONITEMINITIAL")
  End If
End Function

Public Function RIBBONITEMGET(ByVal RItem As String) As Variant
  'Get the current value of any Ribbon item
  On Error GoTo EH
  Let RIBBONITEMGET = RIBBONITEM_Core_get(RItem)

EH:
  If GLOBOO_EH Then
    Call EHGLOBAL.ErrorHandlerRaise("RIBBONITEM::RIBBONITEMGET")
  Else
    Call EHGLOBAL.ErrorHandlerDisplay("RIBBONITEM::RIBBONITEMGET")
  End If
End Function

Public Sub RIBBONITEMSET(ByVal RItem As String, ByVal RValue As Variant)
  'Set a specific value to specific Ribbon item
  On Error GoTo EH
  Call RIBBONITEM_Core_set(RItem, RValue)

EH:
  If GLOBOO_EH Then
    Call EHGLOBAL.ErrorHandlerRaise("RIBBONITEM::RIBBONITEMSET")
  Else
    Call EHGLOBAL.ErrorHandlerDisplay("RIBBONITEM::RIBBONITEMSET")
  End If
End Sub

Public Sub RIBBONITEMENABLINGALL(ByVal RValue As Boolean)
  'Set value to Ribbon all items
  On Error GoTo EH
  Call RIBBONITEM_Core_enabling(RValue)

EH:
  If GLOBOO_EH Then
    Call EHGLOBAL.ErrorHandlerRaise("RIBBONITEM::RIBBONITEMENABLINGALL")
  Else
    Call EHGLOBAL.ErrorHandlerDisplay("RIBBONITEM::RIBBONITEMENABLINGALL")
  End If
End Sub
'=========================== CLASS RIBBONITEMS STRUCTURE ============================'


