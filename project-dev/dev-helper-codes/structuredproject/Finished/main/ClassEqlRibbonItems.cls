VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassEqlRibbonItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ================== CLASS EXCLIQ LITE APP RIBBON ITEMS STRUCTURE ================== '
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


' ================== CLASS EXCLIQ LITE APP RIBBON ITEMS STRUCTURE ================== '
' OBJECT VARIABLES (GLOOBJ_) --------------------------------------------------------'
' VARIABLES (GLOStr_) ---------------------------------------------------------------'
' PROPERTIES ------------------------------------------------------------------------'
' CONSTRUCTOR/DESTRUCTOR ------------------------------------------------------------'
' CORE METHODS LIST (PRIVATE) ------------------------------------------------------ '
' RIBBONITEM_Core_initialize
' RIBBONITEM_Core_get
' RIBBONITEM_Core_set
' RIBBONITEM_Core_enabling
' ASSETS ----------------------------------------------------------------------------'
' ERRORS MANNAGEMENT AND ASSERTSIONS ------------------------------------------------'
' Debug_assert
' CORE METHODS LIST (PUBLIC) ------------------------------------------------------- '
' RIBBONITEMINITIAL
' RIBBONITEMGET
' RIBBONITEMSET
' RIBBONITEMENABLINGALL
' ================== CLASS EXCLIQ LITE APP RIBBON ITEMS STRUCTURE ================== '



' ================== CLASS EXCLIQ LITE APP RIBBON ITEMS STRUCTURE ================== '
' OBJECT VARIABLES (GLOOBJ_) ------------------------------------------------------- '
Private EHGLOBAL As ClassEqlErrorHandler

' VARIABLES (GLOStr_) -------------------------------------------------------------- '
Private GLORIBBON_ITEM As Variant



' PROPERTIES ----------------------------------------------------------------------- '
Public Property Set ErrorHandler(ByRef ObjEH As ClassEqlErrorHandler)

  Set EHGLOBAL = ObjEH

End Property

Public Property Get Items() As Variant

  Let Items = GLORIBBON_ITEM

End Property



' CONSTRUCTOR/DESTRUCTOR ----------------------------------------------------------- '
Private Sub Class_Terminate()

  Let GLORIBBON_ITEM = Empty
  Set EHGLOBAL = Nothing

End Sub



' CORE METHODS LIST (PRIVATE) ------------------------------------------------------ '
Private Function RIBBONITEM_Core_initialize() As Boolean

  Const sFN As String = "RIBBONITEM::-core_initialize"

  On Error GoTo EH
  Let RIBBONITEM_Core_initialize = False

  Let GLORIBBON_ITEM = Empty
  ReDim GLORIBBON_ITEM(0 To 19, 0 To 1)
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
  Let GLORIBBON_ITEM(15, 0) = ("helprlselection")
  Let GLORIBBON_ITEM(16, 0) = ("helpbalmetselection")
  Let GLORIBBON_ITEM(17, 0) = ("helpudf")
  Let GLORIBBON_ITEM(18, 0) = ("about")
  Let GLORIBBON_ITEM(19, 0) = ("license")

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
  Let GLORIBBON_ITEM(17, 1) = (True)
  Let GLORIBBON_ITEM(18, 1) = (True)
  Let GLORIBBON_ITEM(19, 1) = (True)

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

  If Not BooFound Then Call VBA.Err.Raise(CUSTOM_ERROR_RIBBON, sFN, "Ítem: " & RItem & " de la cinta de opciones inválido.")

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

  If BooFound = False Then Call VBA.Err.Raise(CUSTOM_ERROR_RIBBON, sFN, "No se encontró el ítem: " & RItem & " de la cinta de opciones")
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



' ERRORS MANNAGEMENT AND ASSERTSIONS ----------------------------------------------- '
' Debug_assert
Private Sub Debug_assert(ByVal Cond As Boolean, Optional Origin As String = VBA.vbNullString)

  #If Debugging Then
    Debug.Assert Cond
    If Not Cond Then Debug.Print "Debugged from: " & Origin
  #End If

End Sub



' CORE METHODS LIST (PUBLIC) ------------------------------------------------------- '
Public Function RIBBONITEMINITIAL() As Boolean

  ' Sets the initial values for stand by Ribbon
  On Error GoTo EH
  Let RIBBONITEMINITIAL = RIBBONITEM_Core_initialize()

EH:
  Call EHGLOBAL.ErrorHandlerRaise("RIBBONITEM::RIBBONITEMINITIAL")

End Function

Public Function RIBBONITEMGET(ByVal RItem As String) As Variant

  ' Get the current value of any Ribbon item
  On Error GoTo EH
  Let RIBBONITEMGET = RIBBONITEM_Core_get(RItem)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("RIBBONITEM::RIBBONITEMGET")

End Function

Public Sub RIBBONITEMSET(ByVal RItem As String, ByVal RValue As Variant)

  ' Set a specific value to specific Ribbon item
  On Error GoTo EH
  Call RIBBONITEM_Core_set(RItem, RValue)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("RIBBONITEM::RIBBONITEMSET")

End Sub

Public Sub RIBBONITEMENABLINGALL(ByVal RValue As Boolean)

  ' Set value to Ribbon all items
  On Error GoTo EH
  Call RIBBONITEM_Core_enabling(RValue)

EH:
  Call EHGLOBAL.ErrorHandlerRaise("RIBBONITEM::RIBBONITEMENABLINGALL")

End Sub
' CORE METHODS LIST (PUBLIC) ------------------------------------------------------- '
' ================== CLASS EXCLIQ LITE APP RIBBON ITEMS STRUCTURE ================== '

