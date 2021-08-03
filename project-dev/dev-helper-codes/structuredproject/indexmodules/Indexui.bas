Attribute VB_Name = "IndexUI"
Option Explicit



'LICENSE & ACKNOWLEDGMENTS
'
'MIT License
'
'Copyright (c) 2019 Roberto Carlos Romay Medina
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.
'
'Acknowledgments to:
' - StarUML v2.8.0
' - Pencil v3.0.3
' - VSCode v1.41.0
' - InkScape v0.92.4
' - Gimp v2.10.2
' - Just Color Picker v4.6
' - Office RibbonX Editor | Fernando Andreu v1.5.1.418
' - MS Excel v1911
' - VBA7.1 V1091
' - Otto Javier González | www.excelymas.com
' - Ismael Romero | www.excelforo.blogspot.com
' - David Asurmendi | www.davidasurmendi.blogspot.com
' - Sergio Alejandro Campos | www.exceleinfo.com
'
'
'Roccou: I think RefEdits works very well with a good treatment.
'GitHub: https://github.com/roccouu/
'ExcliqLite's home page: https://roccouu.github.io/ExcliqLite/
'ExcliqLite's repo: https://github.com/Roccouu/ExcliqLite
'2019, POTOSÍ - BOLÍVIA


'VERSIONIG
Public Function ExcliqLiteGetVersion() As String
  'Let ExcliqLiteGetVersion = exq_crt.APPVERSION
End Function




'--- RIBBON FUNCTIONS ---
'... Functionality subs ...
Public Function ExcliqLiteTEST() As Variant
  Dim a(0 To 2, 0 To 0) As Variant
  'Let ExcliqLiteTEST = Array(1, 2, 3)
  Let a(0, 0) = 1
  Let a(1, 0) = 2
  Let a(2, 0) = 3
  Let ExcliqLiteTEST = a
  Erase a
End Function

'--- UDF'S FUNCTIONS ---
'Linear Regression UDF'S
Function EQL_REGRESION_LINEAL(ByVal Y As Variant, Optional Opcional_X As Variant, Optional Opcional_Extrapolacion As Variant = 0) As Double
  
  Call Application.Volatile
  Let EQL_REGRESION_LINEAL = exq_crt.LinearRegression("e", Y, Opcional_X, Opcional_Extrapolacion)

End Function

Function EQL_REGRESION_LINEAL_a(ByVal Y As Variant, Optional Opcional_X As Variant) As Double
  
  Call Application.Volatile
  Let EQL_REGRESION_LINEAL_a = exq_crt.LinearRegression("a", Y, Opcional_X, False)

End Function

Function EQL_REGRESION_LINEAL_b(ByVal Y As Variant, Optional Opcional_X As Variant) As Double
  
  Call Application.Volatile
  Let EQL_REGRESION_LINEAL_b = exq_crt.LinearRegression("b", Y, Opcional_X, False)

End Function

Function EQL_REGRESION_LINEAL_r(ByVal Y As Variant, Optional Opcional_X As Variant) As Double
  
  Call Application.Volatile
  Let EQL_REGRESION_LINEAL_r = exq_crt.LinearRegression("r", Y, Opcional_X, False)

End Function

Function EQL_REGRESION_LINEAL_r2(ByVal Y As Variant, Optional Opcional_X As Variant) As Double
  
  Call Application.Volatile
  Let EQL_REGRESION_LINEAL_r2 = exq_crt.LinearRegression("r2", Y, Opcional_X, False)

End Function

Function EQL_REGRESION_LINEAL_Se(ByVal Y As Variant, Optional Opcional_X As Variant) As Double
  
  Call Application.Volatile
  Let EQL_REGRESION_LINEAL_Se = exq_crt.LinearRegression("se", Y, Opcional_X, False)

End Function

Function EQL_REGRESION_LINEAL_n(ByVal Y As Variant, Optional Opcional_X As Variant) As Double
  
  Call Application.Volatile
  Let EQL_REGRESION_LINEAL_n = exq_crt.LinearRegression("n", Y, Opcional_X, False)

End Function

'Balmet UDF's
Public Function EQL_BALMET_CURRENT(ByVal Alimentacion As Variant, ByVal LeyesConcentrado As Variant, Optional LeyesConcentradoComplejo_Opcional As Variant = Empty, Optional NombresElementos_Opcional As Variant = VBA.vbNullString, Optional UnidadesLey_Opcional As Variant = "%", Optional Metodo_Opcional As Long = 0, Optional ResultadoHorizontal_Opcional As Boolean = False, Optional IncluirAlimentacion_Opcional As Boolean = False) As Variant

  Call Application.Volatile(True)
  Let EQL_BALMET_CURRENT = exq_crt.Balmet("balmetudf", Alimentacion, LeyesConcentrado, LeyesConcentradoComplejo_Opcional, NombresElementos_Opcional, UnidadesLey_Opcional, IIf(Metodo_Opcional = 0, eqlBalConventional, IIf(Metodo_Opcional = 1, eqlBalCramer, eqlBalInverseMatrix)), eqlBalWeights, IIf(ResultadoHorizontal_Opcional, eqlBalHorizontal, eqlBalVertical), IncluirAlimentacion_Opcional)

End Function

Public Function EQL_BALMET_PORCENTAJE_EN_PESO_CURRENT(ByVal Alimentacion As Variant, ByVal LeyesConcentrado As Variant, Optional LeyesConcentradoComplejo_Opcional As Variant = Empty, Optional NombresElementos_Opcional As Variant = VBA.vbNullString, Optional UnidadesLey_Opcional As Variant = "%", Optional Metodo_Opcional As Long = 0, Optional ResultadoHorizontal_Opcional As Boolean = False) As Variant

  Call Application.Volatile(True)
  Let EQL_BALMET_PORCENTAJE_EN_PESO_CURRENT = exq_crt.Balmet("balmetudf", Alimentacion, LeyesConcentrado, LeyesConcentradoComplejo_Opcional, NombresElementos_Opcional, UnidadesLey_Opcional, IIf(Metodo_Opcional = 0, eqlBalConventional, IIf(Metodo_Opcional = 1, eqlBalCramer, eqlBalInverseMatrix)), eqlBalWeightPercents, IIf(ResultadoHorizontal_Opcional, eqlBalHorizontal, eqlBalVertical), False)

End Function

Public Function EQL_BALMET_VOLUMEN_CURRENT(ByVal Alimentacion As Variant, ByVal LeyesConcentrado As Variant, Optional LeyesConcentradoComplejo_Opcional As Variant = Empty, Optional NombresElementos_Opcional As Variant = VBA.vbNullString, Optional UnidadesLey_Opcional As Variant = "g/L", Optional Metodo_Opcional As Byte = 0, Optional ResultadoHorizontal_Opcional As Boolean = False) As Variant

  Call Application.Volatile(True)
  Let EQL_BALMET_VOLUMEN_CURRENT = exq_crt.Balmet("balmetudf", Alimentacion, LeyesConcentrado, LeyesConcentradoComplejo_Opcional, NombresElementos_Opcional, UnidadesLey_Opcional, IIf(Metodo_Opcional = 0, eqlBalConventional, IIf(Metodo_Opcional = 1, eqlBalCramer, eqlBalInverseMatrix)), eqlBalVolume, IIf(ResultadoHorizontal_Opcional, eqlBalHorizontal, eqlBalVertical), False)

End Function

Public Function EQL_BALMET_PORCENTAJE_EN_VOLUMEN_CURRENT(ByVal Alimentacion As Variant, ByVal LeyesConcentrado As Variant, Optional LeyesConcentradoComplejo_Opcional As Variant = Empty, Optional NombresElementos_Opcional As Variant = VBA.vbNullString, Optional UnidadesLey_Opcional As Variant = "g/L", Optional Metodo_Opcional As Byte = 0, Optional ResultadoHorizontal_Opcional As Boolean = False) As Variant

  Call Application.Volatile(True)
  Let EQL_BALMET_PORCENTAJE_EN_VOLUMEN_CURRENT = exq_crt.Balmet("balmetudf", Alimentacion, LeyesConcentrado, LeyesConcentradoComplejo_Opcional, NombresElementos_Opcional, UnidadesLey_Opcional, IIf(Metodo_Opcional = 0, eqlBalConventional, IIf(Metodo_Opcional = 1, eqlBalCramer, eqlBalInverseMatrix)), eqlBalVolumePercents, IIf(ResultadoHorizontal_Opcional, eqlBalHorizontal, eqlBalVertical), False)
    
End Function

Public Function EQL_BALMET_CABEZAS_ENSAYADA_CALCULADA_CURRENT(ByVal Alimentacion As Variant, ByVal LeyesConcentrado As Variant, Optional LeyesConcentradoComplejo_Opcional As Variant = Empty, Optional NombresElementos_Opcional As Variant = VBA.vbNullString, Optional UnidadesLey_Opcional As Variant = "%", Optional Metodo_Opcional As Byte = 0, Optional ResultadoHorizontal_Opcional As Boolean = False, Optional IncluirLeyesDeComplejos_Opcional As Boolean = False) As Variant

  Call Application.Volatile(True)
  Let EQL_BALMET_CABEZAS_ENSAYADA_CALCULADA_CURRENT = exq_crt.Balmet("balmetudf", Alimentacion, LeyesConcentrado, LeyesConcentradoComplejo_Opcional, NombresElementos_Opcional, UnidadesLey_Opcional, IIf(Metodo_Opcional = 0, eqlBalConventional, IIf(Metodo_Opcional = 1, eqlBalCramer, eqlBalInverseMatrix)), eqlBalGradesHeads, IIf(ResultadoHorizontal_Opcional, eqlBalHorizontal, eqlBalVertical), IncluirLeyesDeComplejos_Opcional)
  
End Function

Public Function EQL_BALMET_UNIDADES_CURRENT(ByVal Alimentacion As Variant, ByVal LeyesConcentrado As Variant, Optional LeyesConcentradoComplejo_Opcional As Variant = Empty, Optional NombresElementos_Opcional As Variant = VBA.vbNullString, Optional UnidadesLey_Opcional As Variant = "%", Optional Metodo_Opcional As Byte = 0, Optional ResultadoHorizontal_Opcional As Boolean = False, Optional IncluirUnidadesDeComplejos_Opcional As Boolean = False) As Variant

  Call Application.Volatile(True)
  Let EQL_BALMET_UNIDADES_CURRENT = exq_crt.Balmet("balmetudf", Alimentacion, LeyesConcentrado, LeyesConcentradoComplejo_Opcional, NombresElementos_Opcional, UnidadesLey_Opcional, IIf(Metodo_Opcional = 0, eqlBalConventional, IIf(Metodo_Opcional = 1, eqlBalCramer, eqlBalInverseMatrix)), eqlBalUnities, IIf(ResultadoHorizontal_Opcional, eqlBalHorizontal, eqlBalVertical), IncluirUnidadesDeComplejos_Opcional)
  
End Function

Public Function EQL_BALMET_FINOS_CURRENT(ByVal Alimentacion As Variant, ByVal LeyesConcentrado As Variant, Optional LeyesConcentradoComplejo_Opcional As Variant = Empty, Optional NombresElementos_Opcional As Variant = VBA.vbNullString, Optional UnidadesLey_Opcional As Variant = "%", Optional Metodo_Opcional As Byte = 0, Optional ResultadoHorizontal_Opcional As Boolean = False, Optional IncluirFinosDeComplejos_Opcional As Boolean = False) As Variant

  Call Application.Volatile(True)
  Let EQL_BALMET_FINOS_CURRENT = exq_crt.Balmet("balmetudf", Alimentacion, LeyesConcentrado, LeyesConcentradoComplejo_Opcional, NombresElementos_Opcional, UnidadesLey_Opcional, IIf(Metodo_Opcional = 0, eqlBalConventional, IIf(Metodo_Opcional = 1, eqlBalCramer, eqlBalInverseMatrix)), eqlBalFines, IIf(ResultadoHorizontal_Opcional, eqlBalHorizontal, eqlBalVertical), IncluirFinosDeComplejos_Opcional)
  
End Function

Public Function EQL_BALMET_RECUPERACION_CURRENT(ByVal Alimentacion As Variant, ByVal LeyesConcentrado As Variant, Optional LeyesConcentradoComplejo_Opcional As Variant = Empty, Optional NombresElementos_Opcional As Variant = VBA.vbNullString, Optional UnidadesLey_Opcional As Variant = "%", Optional Metodo_Opcional As Byte = 0, Optional ResultadoHorizontal_Opcional As Boolean = False, Optional IncluirRecuperacionDeComplejos_Opcional As Boolean = False) As Variant

  Call Application.Volatile(True)
  Let EQL_BALMET_RECUPERACION_CURRENT = exq_crt.Balmet("balmetudf", Alimentacion, LeyesConcentrado, LeyesConcentradoComplejo_Opcional, NombresElementos_Opcional, UnidadesLey_Opcional, IIf(Metodo_Opcional = 0, eqlBalConventional, IIf(Metodo_Opcional = 1, eqlBalCramer, eqlBalInverseMatrix)), eqlBalRecoveries, IIf(ResultadoHorizontal_Opcional, eqlBalHorizontal, eqlBalVertical), IncluirRecuperacionDeComplejos_Opcional)
  
End Function

Public Function EQL_BALMET_RATIO_CURRENT(ByVal Alimentacion As Variant, ByVal LeyesConcentrado As Variant, Optional LeyesConcentradoComplejo_Opcional As Variant = Empty, Optional NombresElementos_Opcional As Variant = VBA.vbNullString, Optional UnidadesLey_Opcional As Variant = "%", Optional Metodo_Opcional As Byte = 0, Optional ResultadoHorizontal_Opcional As Boolean = False) As Variant

  Call Application.Volatile(True)
  Let EQL_BALMET_RATIO_CURRENT = exq_crt.Balmet("balmetudf", Alimentacion, LeyesConcentrado, LeyesConcentradoComplejo_Opcional, NombresElementos_Opcional, UnidadesLey_Opcional, IIf(Metodo_Opcional = 0, eqlBalConventional, IIf(Metodo_Opcional = 1, eqlBalCramer, eqlBalInverseMatrix)), eqlBalRatio, IIf(ResultadoHorizontal_Opcional, eqlBalHorizontal, eqlBalVertical), False)
  
End Function

Public Function EQL_BALMET_CABEZAS_ENSAYADA_CALCULADA_COMPLEJO_CURRENT(ByVal Alimentacion As Variant, ByVal LeyesConcentrado As Variant, Optional LeyesConcentradoComplejo_Opcional As Variant = Empty, Optional NombresElementos_Opcional As Variant = VBA.vbNullString, Optional UnidadesLey_Opcional As Variant = "%", Optional Metodo_Opcional As Byte = 0, Optional ResultadoHorizontal_Opcional As Boolean = False) As Variant

  Call Application.Volatile(True)
  Let EQL_BALMET_CABEZAS_ENSAYADA_CALCULADA_COMPLEJO_CURRENT = exq_crt.Balmet("balmetudf", Alimentacion, LeyesConcentrado, LeyesConcentradoComplejo_Opcional, NombresElementos_Opcional, UnidadesLey_Opcional, IIf(Metodo_Opcional = 0, eqlBalConventional, IIf(Metodo_Opcional = 1, eqlBalCramer, eqlBalInverseMatrix)), eqlBalGradesHeadsCx, IIf(ResultadoHorizontal_Opcional, eqlBalHorizontal, eqlBalVertical), True)
  
End Function

Public Function EQL_BALMET_UNIDADES_COMPLEJO_CURRENT(ByVal Alimentacion As Variant, ByVal LeyesConcentrado As Variant, Optional LeyesConcentradoComplejo_Opcional As Variant = Empty, Optional NombresElementos_Opcional As Variant = VBA.vbNullString, Optional UnidadesLey_Opcional As Variant = "%", Optional Metodo_Opcional As Byte = 0, Optional ResultadoHorizontal_Opcional As Boolean = False) As Variant

  Call Application.Volatile(True)
  Let EQL_BALMET_UNIDADES_COMPLEJO_CURRENT = exq_crt.Balmet("balmetudf", Alimentacion, LeyesConcentrado, LeyesConcentradoComplejo_Opcional, NombresElementos_Opcional, UnidadesLey_Opcional, IIf(Metodo_Opcional = 0, eqlBalConventional, IIf(Metodo_Opcional = 1, eqlBalCramer, eqlBalInverseMatrix)), eqlBalUnitiesCx, IIf(ResultadoHorizontal_Opcional, eqlBalHorizontal, eqlBalVertical), True)
  
End Function

Public Function EQL_BALMET_FINOS_COMPLEJO_CURRENT(ByVal Alimentacion As Variant, ByVal LeyesConcentrado As Variant, Optional LeyesConcentradoComplejo_Opcional As Variant = Empty, Optional NombresElementos_Opcional As Variant = VBA.vbNullString, Optional UnidadesLey_Opcional As Variant = "%", Optional Metodo_Opcional As Byte = 0, Optional ResultadoHorizontal_Opcional As Boolean = False) As Variant

  Call Application.Volatile(True)
  Let EQL_BALMET_FINOS_COMPLEJO_CURRENT = exq_crt.Balmet("balmetudf", Alimentacion, LeyesConcentrado, LeyesConcentradoComplejo_Opcional, NombresElementos_Opcional, UnidadesLey_Opcional, IIf(Metodo_Opcional = 0, eqlBalConventional, IIf(Metodo_Opcional = 1, eqlBalCramer, eqlBalInverseMatrix)), eqlBalFinesCx, IIf(ResultadoHorizontal_Opcional, eqlBalHorizontal, eqlBalVertical), True)
  
End Function

Public Function EQL_BALMET_RECUPERACION_COMPLEJO_CURRENT(ByVal Alimentacion As Variant, ByVal LeyesConcentrado As Variant, Optional LeyesConcentradoComplejo_Opcional As Variant = Empty, Optional NombresElementos_Opcional As Variant = VBA.vbNullString, Optional UnidadesLey_Opcional As Variant = "%", Optional Metodo_Opcional As Byte = 0, Optional ResultadoHorizontal_Opcional As Boolean = False) As Variant

  Call Application.Volatile(True)
  Let EQL_BALMET_RECUPERACION_COMPLEJO_CURRENT = exq_crt.Balmet("balmetudf", Alimentacion, LeyesConcentrado, LeyesConcentradoComplejo_Opcional, NombresElementos_Opcional, UnidadesLey_Opcional, IIf(Metodo_Opcional = 0, eqlBalConventional, IIf(Metodo_Opcional = 1, eqlBalCramer, eqlBalInverseMatrix)), eqlBalRecoveriesCx, IIf(ResultadoHorizontal_Opcional, eqlBalHorizontal, eqlBalVertical), True)
  
End Function

'Economics balmets
Public Function EQL_BALMET_ECONOMICO_LEY_CABEZA_CURRENT(ByVal Alimentacion As Variant, ByVal PesoConcentrado As Variant, ByVal LeyConcentrado As Variant, ByVal Recuperacion As Variant, Optional UnidadesLey_Opcional As Variant = "%", Optional LeyComplejo_Opcional As Variant = 0, Optional RecuperacionComplejo_Opcional As Variant = 0, Optional Resultado_Base_Complejo_Ambos_Opcional As Long = 0, Optional Resultado_en_Vertical_Opcional As Boolean = False) As Variant
  
  Call Application.Volatile(True)
  Let EQL_BALMET_ECONOMICO_LEY_CABEZA_CURRENT = exq_crt.BalmetEco("balmetudfeco", Alimentacion, PesoConcentrado, LeyConcentrado, Recuperacion, LeyComplejo_Opcional, RecuperacionComplejo_Opcional, UnidadesLey_Opcional, eqlBalHeadsGrades, IIf(Resultado_Base_Complejo_Ambos_Opcional <= 0, eqlBalJustGrades, IIf(Resultado_Base_Complejo_Ambos_Opcional = 1, eqlBalJustGradesCx, eqlBalJustGradesBoth)), IIf(Resultado_en_Vertical_Opcional, eqlBalVertical, eqlBalHorizontal))

End Function

Public Function EQL_BALMET_ECONOMICO_LEY_PRODUCTOS_CURRENT(ByVal Alimentacion As Variant, ByVal PesoConcentrado As Variant, ByVal LeyCabezas As Variant, ByVal Recuperacion As Variant, Optional UnidadesLey_Opcional As Variant = "%", Optional LeyCabezasComplejo_Opcional As Variant = 0, Optional RecuperacionComplejo_Opcional As Variant = 0, Optional Resultado_Base_Complejo_Ambos_Opcional As Long = 0, Optional Resultado_en_Vertical_Opcional As Boolean = False) As Variant
  
  Call Application.Volatile
  Let EQL_BALMET_ECONOMICO_LEY_PRODUCTOS_CURRENT = exq_crt.BalmetEco("balmetudfeco", Alimentacion, PesoConcentrado, LeyCabezas, Recuperacion, LeyCabezasComplejo_Opcional, RecuperacionComplejo_Opcional, UnidadesLey_Opcional, eqlBalProdsGrades, IIf(Resultado_Base_Complejo_Ambos_Opcional <= 0, eqlBalJustGrades, IIf(Resultado_Base_Complejo_Ambos_Opcional = 1, eqlBalJustGradesCx, eqlBalJustGradesBoth)), IIf(Resultado_en_Vertical_Opcional, eqlBalVertical, eqlBalHorizontal))

End Function

''Projections
'Public Function EQL_PROYECCION_PROYECTO(ByVal FlujoNetoDeCaja As Variant) As Variant
'  Call Application.Volatile
'  Let EQL_PROYECCION_PROYECTO = exq_crt.ProjectionCalculate(FlujoNetoDeCaja)
'End Function
'
''Miscellanious
'Public Function EQL_NUMERO_A_TEXTO(ByVal Numero As Variant, Optional Divisa_Singular As String = VBA.vbNullString, Optional Divisa_Multiple As String = VBA.vbNullString, Optional Centavos_Literales As Boolean = False) As Variant
'  Call Application.Volatile
'  Let EQL_NUMERO_A_TEXTO = exq_crt.NumberToStringGet(Numero, Divisa_Singular, Divisa_Multiple, Centavos_Literales)
'End Function
'--- UDF'S FUNCTIONS ---


