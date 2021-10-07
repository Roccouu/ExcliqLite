VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassEqlControllerNumToStr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============ CLASS EXCLIQ LITE CONTROLLER NUMBER TO STRING STRUCTURE ============= '
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


' ============ CLASS EXCLIQ LITE CONTROLLER NUMBER TO STRING STRUCTURE ============= '
' ABSTRACT IMPLEMENTATIONS --------------------------------------------------------- '
' OBJECT VARIABLES (GLOOBJ_) ------------------------------------------------------- '
' PROPERTIES ----------------------------------------------------------------------- '
' ABSTRACT INTERFACE PROPERTIES IMPLEMENTATION ------------------------------------- '
' CONSTRUCTOR/DESTRUCTOR ----------------------------------------------------------- '
' CORE METHODS LIST (PRIVATE) ------------------------------------------------------ '
' aletra
' aConvierteCifra
' NumToString_local
' NumToString_UDF
' ERRORS MANNAGEMENT AND ASSERTSIONS ----------------------------------------------- '
' Debug_assert
' CORE METHODS LIST (PUBLIC) ------------------------------------------------------- '
' NumToString
' ABSTRACT INTERFACE UI METHODS LIST AND UDF METHODS LIST (PUBLIC) ------------------'
' ClassEqlControllerInterface_GetSolution
' ClassEqlControllerInterface_GetSolutionUDF
' ============ CLASS EXCLIQ LITE CONTROLLER NUMBER TO STRING STRUCTURE ============= '




' ============ CLASS EXCLIQ LITE CONTROLLER NUMBER TO STRING STRUCTURE ============= '
' ABSTRACT IMPLEMENTATIONS --------------------------------------------------------- '
Implements ClassEqlControllerInterface

' OBJECT VARIABLES (GLOOBJ_) ------------------------------------------------------- '
Private RES As ClassEqlResources
' Unidad                                           x
' Decena                                          xx
' Centena                                        xxx
' Unidad de Mil                                x.xxx
' Decena de Mil                               xx.xxx
' Centena de Mil                             xxx.xxx
' Unidad de Millón                         x.xxx.xxx
' Decena de Millón                        xx.xxx.xxx
' Centena de Millón                      xxx.xxx.xxx
' Unidad de Mil Millones               x.xxx.xxx.xxx
' Decena de Mil Millones              xx.xxx.xxx.xxx
' Centena de Mil Millones            xxx.xxx.xxx.xxx
' Unidad de Billón                 x.xxx.xxx.xxx.xxx
' Decena de Billón                xx.xxx.xxx.xxx.xxx
' Centena de Billón              xxx.xxx.xxx.xxx.xxx
' Unidad de Mil Billones       x.xxx.xxx.xxx.xxx.xxx
' Decena de Mil Billones      xx.xxx.xxx.xxx.xxx.xxx
' Centena de Mil Billones    xxx.xxx.xxx.xxx.xxx.xxx
' Unidad de Trillón        x.xxx.xxx.xxx.xxx.xxx.xxx
' Un Billón 1.000.000.000.000
' Currency: -922.337.203.685.477,5808 a 922.337.203.685.477,5807, Char type: @





' PROPERTIES ----------------------------------------------------------------------- '
' ABSTRACT INTERFACE PROPERTIES IMPLEMENTATION ------------------------------------- '
Public Property Set ClassEqlControllerInterface_ResourcesApp(ByRef RHS As ClassEqlResources)

  Set RES = RHS

End Property




' CONSTRUCTOR/DESTRUCTOR ----------------------------------------------------------- '
Private Sub Class_Terminate()

  Set RES = Nothing

End Sub







' CORE METHODS PRIVATE ------------------------------------------------------------- '
Private Function aletra(ByVal Numero As Currency, Optional literalDecimals As Boolean) As String

  Dim Texto, Billones, MilMillones, Millones, Miles, Cientos, Decimales, Cadena, CadBillones, CadMilMillones, CadMillones, CadMiles, CadCientos, CadDecimales As String

  On Error GoTo EH
  Let Numero = VBA.Abs(Numero)
  Let Texto = Round(Numero, 2)
  Let Texto = FormatNumber(Texto, 2)
  Let Texto = Right(Space(22) & Texto, 22) '14 Caracteres desde la derecha de los(14 espacios a la izq + Texto), eso es para contar hasta decena de billón
  'Centenea de Billón                xxx.xxx.xxx.xxx.xxx,yy
  Let Billones = Mid(Texto, 1, 3)
  Let MilMillones = Mid(Texto, 5, 3)
  Let Millones = Mid(Texto, 9, 3)
  Let Miles = Mid(Texto, 13, 3)
  Let Cientos = Mid(Texto, 17, 3)
  Let Decimales = Mid(Texto, 21)

  Let CadBillones = aConvierteCifra(Billones, False)
  Let CadMilMillones = aConvierteCifra(MilMillones, False)
  Let CadMillones = aConvierteCifra(Millones, False)
  Let CadMiles = aConvierteCifra(Miles, False)
  Let CadCientos = aConvierteCifra(Cientos, True)

  If Trim(CadBillones) > VBA.vbNullString Then
    Let Cadena = Cadena & Trim(CadBillones) & " BILL" & IIf(Trim(CadBillones) = "UN", "ÓN ", "ONES ")
  End If
  If Trim(CadBillones & CadMilMillones) > VBA.vbNullString Then
    Let Cadena = Cadena & Trim(CadMilMillones) & " MIL "
  End If
  If Trim(CadMilMillones & CadMillones) > VBA.vbNullString Then
    Let Cadena = Cadena & Trim(CadMillones) & " MILL" & IIf(Trim(CadMillones) = "UN", "ÓN ", "ONES ")
  End If
  If Trim(CadMiles) > VBA.vbNullString Then
    Let Cadena = Cadena & Trim(CadMiles) & " MIL "
  End If
  If Not RES.REGEX.isEmptyStringReg(CadCientos) Then
    Let Cadena = Cadena & Trim(CadCientos)
  End If

  Let CadDecimales = Decimales
  Let Cadena = Cadena & IIf(literalDecimals, " CON ", " ") & Trim(CadDecimales) & "/100 "
  Let aletra = Trim(Cadena)

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise("NUMTXT::-aletra")

End Function

Private Function aConvierteCifra(Texto, IsCientos As Boolean)

  Dim Centena, Decena, Unidad, atxtCentena, atxtDecena, atxtUnidad As String
  Dim aCentenas, aDecenas, aUnidades, adecena As Variant

  On Error GoTo EH
  Let aCentenas = Array(VBA.vbNullString, "CIEN", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS", "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS", "OCHOCIENTOS", "NOVECIENTOS")
  Let aDecenas = Array(VBA.vbNullString, "DIEZ", "VEINT", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", "SETENTA", "OCHENTA", "NOVENTA")
  Let aUnidades = Array("CERO", "UN", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE")
  Let adecena = Array("DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", "QUINCE", "DIECISEIS", "DIECISIETE", "DIECIOCHO", "DIECINUEVE")

  Let Centena = Mid(Texto, 1, 1)
  Let Decena = Mid(Texto, 2, 1)
  Let Unidad = Mid(Texto, 3, 1)
  Let atxtCentena = VBA.vbNullString: atxtDecena = VBA.vbNullString: atxtUnidad = VBA.vbNullString

  If Not RES.REGEX.isEmptyStringReg(Centena) Then 'Centenas
    Let atxtCentena = aCentenas(Centena) & IIf(Centena = "1" And Decena & Unidad <> "00", "TO", VBA.vbNullString)
  End If

  If Not RES.REGEX.isEmptyStringReg(Decena) Then 'Decenas
    If Decena <> "0" Then
      If Decena = "1" Then
        If Unidad = "0" Then
          Let atxtDecena = aDecenas(Decena)
        Else
          Let atxtDecena = adecena(Unidad)
        End If
      Else
        Let atxtDecena = aDecenas(Decena)
        Let atxtDecena = atxtDecena & IIf(Decena = "2", IIf(Unidad = "0", "E", "I"), IIf(Unidad = "0", VBA.vbNullString, " Y "))
      End If
    End If
  End If

  If Not RES.REGEX.isEmptyStringReg(Unidad) Then 'Unidades
    If Decena <> "1" Then
      Let atxtUnidad = aUnidades(Unidad) '& IIf(Unidad = "1", IIf(IsCientos, "O", vba.vbnullstring), vba.vbnullstring)
    End If
  End If

  If Trim(Centena & Decena) <> VBA.vbNullString And Unidad = "0" Then
    Let atxtUnidad = VBA.vbNullString
  End If

  Let aConvierteCifra = atxtCentena & " " & atxtDecena & atxtUnidad

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise("NUMTXT::-aConvierteCifra")

End Function

Private Function NumToString_local( _
  ByVal DblNum As Double, _
  Optional StrCursingle As String = VBA.vbNullString, _
  Optional StrCurmultiple As String = VBA.vbNullString, _
  Optional BooLiteralcents As Boolean) As String

  Let NumToString_local = aletra(DblNum, BooLiteralcents) & " " & IIf(Round(DblNum, 0) = 1, StrCursingle, StrCurmultiple)

End Function

Private Function NumToString_UDF( _
  ByVal VarNum As Variant, _
  Optional StrCursingle As String = VBA.vbNullString, _
  Optional StrCurmultiple As String = VBA.vbNullString, _
  Optional BooLiteralcents As Boolean) As Variant

  Dim BooError As Boolean

  On Error GoTo EH

  If TypeName(VarNum) = "Range" Then
    Let VarNum = Application.WorksheetFunction.Sum(VarNum)
  ElseIf VBA.TypeName(VarNum) = "Double" Then
  Else
    Let BooError = True: GoTo EH
  End If
  If VBA.TypeName(VarNum) = "Double" Then
    If VBA.Len(VBA.CStr(VarNum)) > 19 Then BooError = True: GoTo EH
  End If

  Let NumToString_UDF = aletra(VarNum, BooLiteralcents) & " " & IIf(Round(VarNum, 0) = 1, StrCursingle, StrCurmultiple)

EH:
  If BooError Then NumToString_UDF = VBA.CVErr(xlErrValue)

End Function



' ERROR ASSERTSIONS ---------------------------------------------------------------- '
Private Sub Debug_assert(ByVal Cond As Boolean, Optional Origin As String = VBA.vbNullString)

  #If Debbuging Then
    Debug.Assert Cond
    If Not Cond Then Debug.Print "Assert on: " & Origin
  #End If

End Sub
' CORE METHODS PRIVATE ------------------------------------------------------------- '









' METHODS PUBLIC ------------------------------------------------------------------- '
Public Function NumToString( _
  ByVal NDblNum As Double, _
  Optional NStrCursingle As String = VBA.vbNullString, _
  Optional NStrCurmultiple As String = VBA.vbNullString, _
  Optional NBooLiteralcents As Boolean) As String

  On Error GoTo EH
  Let NumToString = NumToString_local(NDblNum, NStrCursingle, NStrCurmultiple, NBooLiteralcents)

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise("NTS::NumToString")

End Function

' ABSTRACT INTERFACE UI METHODS LIST AND UDF METHODS LIST (PUBLIC) ----------------- '
Private Function ClassEqlControllerInterface_GetSolution(ByVal CStrRequester As String, ByRef CVarRequest As Variant, ByRef CVarResponse As Variant) As Boolean

  On Error GoTo EH
  Let CVarResponse = NumToString_local(VBA.CDbl(CVarRequest(0)), VBA.CStr(CVarRequest(1)), VBA.CStr(CVarRequest(2)), VBA.CBool(CVarRequest(3)))

EH:
  Call RES.ErrorHandler.ErrorHandlerRaise("NTS::IGetSolution")

End Function

Public Function ClassEqlControllerInterface_GetSolutionUDF(ByVal CStrRequester As String, ByVal CVarRequest As Variant) As Variant

  Let ClassEqlControllerInterface_GetSolutionUDF = NumToString_UDF(VBA.CDbl(CVarRequest(0)), VBA.CStr(CVarRequest(1)), VBA.CStr(CVarRequest(2)), VBA.CBool(CVarRequest(3)))

End Function
' METHODS PUBLIC ------------------------------------------------------------------- '
' ============ CLASS EXCLIQ LITE CONTROLLER NUMBER TO STRING STRUCTURE ============= '

