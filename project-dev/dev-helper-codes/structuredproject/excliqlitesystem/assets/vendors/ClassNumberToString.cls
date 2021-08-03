VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppExcliqNumToStr_current"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'============================= CLASS MODULES STRUCTURE =============================='
' ABSTRACT IMPLEMENTATIONS ----------------------------------------------------------'
Implements ICoreController

' OBJECT VARIABLES (GLOOBJ_) --------------------------------------------------------'
Private EHGLOBAL As AppErrorHandler
Private REGEX As New AppResRegEx
  'Unidad                                           x
  'Decena                                          xx
  'Centena                                        xxx
  'Unidad de Mil                                x.xxx
  'Decena de Mil                               xx.xxx
  'Centena de Mil                             xxx.xxx
  'Unidad de Millón                         x.xxx.xxx
  'Decena de Millón                        xx.xxx.xxx
  'Centena de Millón                      xxx.xxx.xxx
  'Unidad de Mil Millones               x.xxx.xxx.xxx
  'Decena de Mil Millones              xx.xxx.xxx.xxx
  'Centena de Mil Millones            xxx.xxx.xxx.xxx
  'Unidad de Billón                 x.xxx.xxx.xxx.xxx
  'Decena de Billón                xx.xxx.xxx.xxx.xxx
  'Centena de Billón              xxx.xxx.xxx.xxx.xxx
  'Unidad de Mil Billones       x.xxx.xxx.xxx.xxx.xxx
  'Decena de Mil Billones      xx.xxx.xxx.xxx.xxx.xxx
  'Centena de Mil Billones    xxx.xxx.xxx.xxx.xxx.xxx
  'Unidad de Trillón        x.xxx.xxx.xxx.xxx.xxx.xxx
  'Un Billón 1.000.000.000.000
  'Currency: -922.337.203.685.477,5808 a 922.337.203.685.477,5807, Char type: @






' PROPERTIES ------------------------------------------------------------------------'
' ABSTRACT INTERFACE PROPERTIES IMPLEMENTATION --------------------------------------'
Public Property Set ICoreController_ErrorHandler(RHS As AppErrorHandler)
  Set EHGLOBAL = Nothing
  Set EHGLOBAL = RHS
End Property






' CONSTRUCTOR/DESTRUCTOR ------------------------------------------------------------'
Private Sub Class_Initialize()
  Set EHGLOBAL = New AppErrorHandler
End Sub

Private Sub Class_Terminate()
  Set EHGLOBAL = Nothing
End Sub







' CORE METHODS PRIVATE --------------------------------------------------------------'
Private Function aletra(ByVal Numero As Currency, Optional literalDecimals As Boolean) As String
  Dim Texto, Billones, MilMillones, Millones, Miles, Cientos, Decimales, Cadena, CadBillones, CadMilMillones, CadMillones, CadMiles, CadCientos, CadDecimales As String
  
  'VBA.Round(VBA.Abs(VBA.CDbl(Sh.Range("I" & auxRow - 1).value)), 2)
  Numero = VBA.Abs(Numero)
  Texto = Round(Numero, 2)
  Texto = FormatNumber(Texto, 2)
  Texto = Right(Space(22) & Texto, 22) '14 Caracteres desde la derecha de los(14 espacios a la izq + Texto), eso es para contar hasta decena de billón
  'Centenea de Billón                xxx.xxx.xxx.xxx.xxx,yy
  Billones = Mid(Texto, 1, 3)
  MilMillones = Mid(Texto, 5, 3)
  Millones = Mid(Texto, 9, 3)
  Miles = Mid(Texto, 13, 3)
  Cientos = Mid(Texto, 17, 3)
  Decimales = Mid(Texto, 21)
  
  CadBillones = aConvierteCifra(Billones, False)
  CadMilMillones = aConvierteCifra(MilMillones, False)
  CadMillones = aConvierteCifra(Millones, False)
  CadMiles = aConvierteCifra(Miles, False)
  CadCientos = aConvierteCifra(Cientos, True)
  
  If Trim(CadBillones) > VBA.vbNullString Then
    Cadena = Cadena & Trim(CadBillones) & " BILL" & IIf(Trim(CadBillones) = "UN", "ÓN ", "ONES ")
  End If
  If Trim(CadBillones & CadMilMillones) > VBA.vbNullString Then
    Cadena = Cadena & Trim(CadMilMillones) & " MIL "
  End If
  If Trim(CadMilMillones & CadMillones) > VBA.vbNullString Then
    Cadena = Cadena & Trim(CadMillones) & " MILL" & IIf(Trim(CadMillones) = "UN", "ÓN ", "ONES ")
  End If
  If Trim(CadMiles) > VBA.vbNullString Then
    Cadena = Cadena & Trim(CadMiles) & " MIL "
  End If
  If Not REGEX.isEmptyStringReg(CadCientos) Then
    Cadena = Cadena & Trim(CadCientos)
  End If
  
  CadDecimales = Decimales
  Cadena = Cadena & IIf(literalDecimals, " CON ", " ") & Trim(CadDecimales) & "/100 "
  aletra = Trim(Cadena)
End Function

Private Function aConvierteCifra(Texto, IsCientos As Boolean)
  Dim Centena, Decena, Unidad, atxtCentena, atxtDecena, atxtUnidad As String
  Dim aCentenas, aDecenas, aUnidades, adecena As Variant
  
  aCentenas = Array(VBA.vbNullString, "CIEN", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS", "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS", "OCHOCIENTOS", "NOVECIENTOS")
  aDecenas = Array(VBA.vbNullString, "DIEZ", "VEINT", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", "SETENTA", "OCHENTA", "NOVENTA")
  aUnidades = Array("CERO", "UN", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE")
  adecena = Array("DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", "QUINCE", "DIECISEIS", "DIECISIETE", "DIECIOCHO", "DIECINUEVE")
  
  Centena = Mid(Texto, 1, 1)
  Decena = Mid(Texto, 2, 1)
  Unidad = Mid(Texto, 3, 1)
  atxtCentena = VBA.vbNullString: atxtDecena = VBA.vbNullString: atxtUnidad = VBA.vbNullString
  
  If Not REGEX.isEmptyStringReg(Centena) Then 'Centenas
    atxtCentena = aCentenas(Centena) & IIf(Centena = "1" And Decena & Unidad <> "00", "TO", VBA.vbNullString)
  End If
    
  If Not REGEX.isEmptyStringReg(Decena) Then 'Decenas
    If Decena <> "0" Then
      If Decena = "1" Then
        If Unidad = "0" Then
          atxtDecena = aDecenas(Decena)
        Else
          atxtDecena = adecena(Unidad)
        End If
      Else
        atxtDecena = aDecenas(Decena)
        atxtDecena = atxtDecena & IIf(Decena = "2", IIf(Unidad = "0", "E", "I"), IIf(Unidad = "0", VBA.vbNullString, " Y "))
      End If
    End If
  End If
                        
  If Not REGEX.isEmptyStringReg(Unidad) Then 'Unidades
    If Decena <> "1" Then
      atxtUnidad = aUnidades(Unidad) '& IIf(Unidad = "1", IIf(IsCientos, "O", vba.vbnullstring), vba.vbnullstring)
    End If
  End If
    
  If Trim(Centena & Decena) <> VBA.vbNullString And Unidad = "0" Then
    atxtUnidad = VBA.vbNullString
  End If
  
  aConvierteCifra = atxtCentena & " " & atxtDecena & atxtUnidad
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






' ASSETS ----------------------------------------------------------------------------'







' ERROR ASSERTSIONS -----------------------------------------------------------------'
Private Sub Debug_assert(ByVal Cond As Boolean, Optional Origin As String = VBA.vbNullString)
  #If Debbuging Then
    Debug.Assert Cond
    If Not Cond Then Debug.Print "Assert on: " & Origin
  #End If
End Sub
' CORE METHODS PRIVATE --------------------------------------------------------------'









' METHODS PUBLIC --------------------------------------------------------------------'
' ABSTRACT INTERFACE UI METHODS LIST AND UDF METHODS LIST (PUBLIC) ------------------'
Public Function ICoreController_GetSolution(ByVal CStrRequester As String, ByVal CVarRequest As Variant, CVarResponse As Variant) As Boolean
  
  On Error GoTo EH
  Let CVarResponse = NumToString_local(VBA.CDbl(CVarRequest(0)), VBA.CStr(CVarRequest(1)), VBA.CStr(CVarRequest(2)), VBA.CBool(CVarRequest(3)))

EH:
  Call EHGLOBAL.ErrorHandlerRaise("APPNTS::IGetSolution")

End Function

Public Function ICoreController_GetSolutionUDF(ByVal CStrRequester As String, ByVal CVarRequest As Variant) As Variant
  
  Let ICoreController_GetSolutionUDF = NumToString_UDF(VBA.CDbl(CVarRequest(0)), VBA.CStr(CVarRequest(1)), VBA.CStr(CVarRequest(2)), VBA.CBool(CVarRequest(3)))
  
End Function
' METHODS PUBLIC --------------------------------------------------------------------'
'============================= CLASS MODULES STRUCTURE =============================='


