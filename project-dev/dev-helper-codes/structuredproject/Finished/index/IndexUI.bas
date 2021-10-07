Attribute VB_Name = "IndexUI"
Option Explicit


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


' =========================== INDEX UI MODULE STRUCTURE ============================ '
' CORE METHODS LIST PRIVATE -------------------------------------------------------- '
' PUBLIC METHODS LIST (INTERFACE) -------------------------------------------------- '
' ExcliqLiteTEST
' UDF'S functions ------------------------------------------------------------------ '
' EQL_VERSION
' ExcliqLiteGetVersion
' EQL_REGRESION_LINEAL
' EQL_REGRESION_LINEAL_a
' EQL_REGRESION_LINEAL_b
' EQL_REGRESION_LINEAL_r
' EQL_REGRESION_LINEAL_r2
' EQL_REGRESION_LINEAL_Se
' EQL_REGRESION_LINEAL_n
' EQL_BALMET
' EQL_BALMET_PORCENTAJE_EN_PESO
' EQL_BALMET_VOLUMEN
' EQL_BALMET_PORCENTAJE_EN_VOLUMEN
' EQL_BALMET_CABEZAS_ENSAYADA_CALCULADA
' EQL_BALMET_UNIDADES
' EQL_BALMET_FINOS
' EQL_BALMET_RECUPERACION
' EQL_BALMET_RATIO
' EQL_BALMET_CABEZAS_ENSAYADA_CALCULADA_COMPLEJO
' EQL_BALMET_UNIDADES_COMPLEJO
' EQL_BALMET_FINOS_COMPLEJO
' EQL_BALMET_RECUPERACION_COMPLEJO
' EQL_BALMET_ECONOMICO_LEY_CABEZA
' EQL_BALMET_ECONOMICO_LEY_PRODUCTOS
' EQL_PROYECCION_PROYECTO
' EQL_NUMERO_A_TEXTO
' =========================== INDEX UI MODULE STRUCTURE ============================ '




' =========================== INDEX UI MODULE STRUCTURE ============================ '
' PUBLIC METHODS LIST (INTERFACE) -------------------------------------------------- '
' ... Functionality subs ...
Public Function ExcliqLiteTEST() As Variant
  Dim a(0 To 2, 0 To 0) As Variant
  'Let ExcliqLiteTEST = Array(1, 2, 3)
  Let a(0, 0) = 1
  Let a(1, 0) = 2
  Let a(2, 0) = 3
  Let ExcliqLiteTEST = a
  Erase a
End Function



' UDF'S functions ------------------------------------------------------------------ '
' VERSIONING
Public Function EQL_VERSION() As String
Attribute EQL_VERSION.VB_Description = "Devuelve la versión actual de Excliq Lite."
Attribute EQL_VERSION.VB_ProcData.VB_Invoke_Func = " \n21"

  If Not exq_crt Is Nothing Then Let EQL_VERSION = exq_crt.EqlVersion()

End Function

Public Function ExcliqLiteGetVersion() As String

  Let ExcliqLiteGetVersion = EQL_VERSION()

End Function


' Linear Regression UDF'S
Function EQL_REGRESION_LINEAL(ByVal y As Variant, Optional Opcional_X As Variant, Optional Opcional_Extrapolacion As Variant = 0) As Double
Attribute EQL_REGRESION_LINEAL.VB_Description = "Obtiene la regresión lineal de tres o más datos."
Attribute EQL_REGRESION_LINEAL.VB_ProcData.VB_Invoke_Func = " \n21"

  Call Application.Volatile
  If Not exq_crt Is Nothing Then Let EQL_REGRESION_LINEAL = exq_crt.LinearRegression("e", y, Opcional_X, Opcional_Extrapolacion)

End Function

Function EQL_REGRESION_LINEAL_a(ByVal y As Variant, Optional Opcional_X As Variant) As Double
Attribute EQL_REGRESION_LINEAL_a.VB_Description = "Devuelve el Coeficiente a de una regresión."
Attribute EQL_REGRESION_LINEAL_a.VB_ProcData.VB_Invoke_Func = " \n21"

  Call Application.Volatile
  If Not exq_crt Is Nothing Then Let EQL_REGRESION_LINEAL_a = exq_crt.LinearRegression("a", y, Opcional_X, False)

End Function

Function EQL_REGRESION_LINEAL_b(ByVal y As Variant, Optional Opcional_X As Variant) As Double
Attribute EQL_REGRESION_LINEAL_b.VB_Description = "Devuelve el Coeficiente b de una regresión."
Attribute EQL_REGRESION_LINEAL_b.VB_ProcData.VB_Invoke_Func = " \n21"

  Call Application.Volatile
  If Not exq_crt Is Nothing Then Let EQL_REGRESION_LINEAL_b = exq_crt.LinearRegression("b", y, Opcional_X, False)

End Function

Function EQL_REGRESION_LINEAL_r(ByVal y As Variant, Optional Opcional_X As Variant) As Double
Attribute EQL_REGRESION_LINEAL_r.VB_Description = "Devuelve el Coeficiente de correlación de una regresión."
Attribute EQL_REGRESION_LINEAL_r.VB_ProcData.VB_Invoke_Func = " \n21"

  Call Application.Volatile
  If Not exq_crt Is Nothing Then Let EQL_REGRESION_LINEAL_r = exq_crt.LinearRegression("r", y, Opcional_X, False)

End Function

Function EQL_REGRESION_LINEAL_r2(ByVal y As Variant, Optional Opcional_X As Variant) As Double
Attribute EQL_REGRESION_LINEAL_r2.VB_Description = "Devuelve el Coeficiente de determinación de una regresión."
Attribute EQL_REGRESION_LINEAL_r2.VB_ProcData.VB_Invoke_Func = " \n21"

  Call Application.Volatile
  If Not exq_crt Is Nothing Then Let EQL_REGRESION_LINEAL_r2 = exq_crt.LinearRegression("r2", y, Opcional_X, False)

End Function

Function EQL_REGRESION_LINEAL_Se(ByVal y As Variant, Optional Opcional_X As Variant) As Double
Attribute EQL_REGRESION_LINEAL_Se.VB_Description = "Devuelve el Error estándar de una regresión."
Attribute EQL_REGRESION_LINEAL_Se.VB_ProcData.VB_Invoke_Func = " \n21"

  Call Application.Volatile
  If Not exq_crt Is Nothing Then Let EQL_REGRESION_LINEAL_Se = exq_crt.LinearRegression("se", y, Opcional_X, False)

End Function

Function EQL_REGRESION_LINEAL_n(ByVal y As Variant, Optional Opcional_X As Variant) As Double
Attribute EQL_REGRESION_LINEAL_n.VB_Description = "Devuelve el total de datos que intervienen en la regresión."
Attribute EQL_REGRESION_LINEAL_n.VB_ProcData.VB_Invoke_Func = " \n21"

  Call Application.Volatile
  If Not exq_crt Is Nothing Then Let EQL_REGRESION_LINEAL_n = exq_crt.LinearRegression("n", y, Opcional_X, False)

End Function

' Balmet UDF's
Public Function EQL_BALMET(ByVal Alimentacion As Variant, ByVal LeyesConcentrado As Variant, Optional NombresElementos_Opcional As Variant = VBA.vbNullString, Optional UnidadesLey_Opcional As Variant = "%", Optional Metodo_Opcional As Long = 0, Optional ResultadoHorizontal_Opcional As Boolean = False, Optional IncluirAlimentacion_Opcional As Boolean = False) As Variant
Attribute EQL_BALMET.VB_Description = "Obtiene el o los pesos de productos de un balance metalúrgico."
Attribute EQL_BALMET.VB_ProcData.VB_Invoke_Func = " \n21"

  Call Application.Volatile(True)
  If Not exq_crt Is Nothing Then Let EQL_BALMET = exq_crt.Balmet("balmetudf", Alimentacion, LeyesConcentrado, Empty, NombresElementos_Opcional, UnidadesLey_Opcional, IIf(Metodo_Opcional = 0, eqlBalConventional, IIf(Metodo_Opcional = 1, eqlBalCramer, eqlBalInverseMatrix)), eqlBalWeights, IIf(ResultadoHorizontal_Opcional, eqlBalHorizontal, eqlBalVertical), IncluirAlimentacion_Opcional)

End Function

Public Function EQL_BALMET_PORCENTAJE_EN_PESO(ByVal Alimentacion As Variant, ByVal LeyesConcentrado As Variant, Optional NombresElementos_Opcional As Variant = VBA.vbNullString, Optional UnidadesLey_Opcional As Variant = "%", Optional Metodo_Opcional As Long = 0, Optional ResultadoHorizontal_Opcional As Boolean = False) As Variant
Attribute EQL_BALMET_PORCENTAJE_EN_PESO.VB_Description = "Obtiene el Porcentaje de en peso de un balance metalúrgico."
Attribute EQL_BALMET_PORCENTAJE_EN_PESO.VB_ProcData.VB_Invoke_Func = " \n21"

  Call Application.Volatile(True)
  If Not exq_crt Is Nothing Then Let EQL_BALMET_PORCENTAJE_EN_PESO = exq_crt.Balmet("balmetudf", Alimentacion, LeyesConcentrado, Empty, NombresElementos_Opcional, UnidadesLey_Opcional, IIf(Metodo_Opcional = 0, eqlBalConventional, IIf(Metodo_Opcional = 1, eqlBalCramer, eqlBalInverseMatrix)), eqlBalWeightPercents, IIf(ResultadoHorizontal_Opcional, eqlBalHorizontal, eqlBalVertical), False)

End Function

Public Function EQL_BALMET_VOLUMEN(ByVal Alimentacion As Variant, ByVal LeyesConcentrado As Variant, Optional NombresElementos_Opcional As Variant = VBA.vbNullString, Optional UnidadesLey_Opcional As Variant = "g/L", Optional Metodo_Opcional As Byte = 0, Optional ResultadoHorizontal_Opcional As Boolean = False) As Variant
Attribute EQL_BALMET_VOLUMEN.VB_Description = "Obtiene el o los volúmenes de productos de un balance metalúrgico."
Attribute EQL_BALMET_VOLUMEN.VB_ProcData.VB_Invoke_Func = " \n21"

  Call Application.Volatile(True)
  If Not exq_crt Is Nothing Then Let EQL_BALMET_VOLUMEN = exq_crt.Balmet("balmetudf", Alimentacion, LeyesConcentrado, Empty, NombresElementos_Opcional, UnidadesLey_Opcional, IIf(Metodo_Opcional = 0, eqlBalConventional, IIf(Metodo_Opcional = 1, eqlBalCramer, eqlBalInverseMatrix)), eqlBalVolume, IIf(ResultadoHorizontal_Opcional, eqlBalHorizontal, eqlBalVertical), False)

End Function

Public Function EQL_BALMET_PORCENTAJE_EN_VOLUMEN(ByVal Alimentacion As Variant, ByVal LeyesConcentrado As Variant, Optional NombresElementos_Opcional As Variant = VBA.vbNullString, Optional UnidadesLey_Opcional As Variant = "g/L", Optional Metodo_Opcional As Byte = 0, Optional ResultadoHorizontal_Opcional As Boolean = False) As Variant
Attribute EQL_BALMET_PORCENTAJE_EN_VOLUMEN.VB_Description = "Obtiene el Porcentaje de en volumen de un balance metalúrgico."
Attribute EQL_BALMET_PORCENTAJE_EN_VOLUMEN.VB_ProcData.VB_Invoke_Func = " \n21"

  Call Application.Volatile(True)
  If Not exq_crt Is Nothing Then Let EQL_BALMET_PORCENTAJE_EN_VOLUMEN = exq_crt.Balmet("balmetudf", Alimentacion, LeyesConcentrado, Empty, NombresElementos_Opcional, UnidadesLey_Opcional, IIf(Metodo_Opcional = 0, eqlBalConventional, IIf(Metodo_Opcional = 1, eqlBalCramer, eqlBalInverseMatrix)), eqlBalVolumePercents, IIf(ResultadoHorizontal_Opcional, eqlBalHorizontal, eqlBalVertical), False)

End Function

Public Function EQL_BALMET_CABEZAS_ENSAYADA_CALCULADA(ByVal Alimentacion As Variant, ByVal LeyesConcentrado As Variant, Optional NombresElementos_Opcional As Variant = VBA.vbNullString, Optional UnidadesLey_Opcional As Variant = "%", Optional Metodo_Opcional As Byte = 0, Optional ResultadoHorizontal_Opcional As Boolean = False) As Variant ', Optional IncluirLeyesDeComplejos_Opcional As Boolean = False) As Variant
Attribute EQL_BALMET_CABEZAS_ENSAYADA_CALCULADA.VB_Description = "Obtiene leyes de cabeza calculada y ensayada de un balance metalúrgico."
Attribute EQL_BALMET_CABEZAS_ENSAYADA_CALCULADA.VB_ProcData.VB_Invoke_Func = " \n21"

  Call Application.Volatile(True)
  If Not exq_crt Is Nothing Then Let EQL_BALMET_CABEZAS_ENSAYADA_CALCULADA = exq_crt.Balmet("balmetudf", Alimentacion, LeyesConcentrado, Empty, NombresElementos_Opcional, UnidadesLey_Opcional, IIf(Metodo_Opcional = 0, eqlBalConventional, IIf(Metodo_Opcional = 1, eqlBalCramer, eqlBalInverseMatrix)), eqlBalGradesHeads, IIf(ResultadoHorizontal_Opcional, eqlBalHorizontal, eqlBalVertical), False) 'IncluirLeyesDeComplejos_Opcional)

End Function

Public Function EQL_BALMET_UNIDADES(ByVal Alimentacion As Variant, ByVal LeyesConcentrado As Variant, Optional NombresElementos_Opcional As Variant = VBA.vbNullString, Optional UnidadesLey_Opcional As Variant = "%", Optional Metodo_Opcional As Byte = 0, Optional ResultadoHorizontal_Opcional As Boolean = False) As Variant ', Optional IncluirUnidadesDeComplejos_Opcional As Boolean = False) As Variant
Attribute EQL_BALMET_UNIDADES.VB_Description = "Obtiene las unidades de un balance metalúrgico."
Attribute EQL_BALMET_UNIDADES.VB_ProcData.VB_Invoke_Func = " \n21"

  Call Application.Volatile(True)
  If Not exq_crt Is Nothing Then Let EQL_BALMET_UNIDADES = exq_crt.Balmet("balmetudf", Alimentacion, LeyesConcentrado, Empty, NombresElementos_Opcional, UnidadesLey_Opcional, IIf(Metodo_Opcional = 0, eqlBalConventional, IIf(Metodo_Opcional = 1, eqlBalCramer, eqlBalInverseMatrix)), eqlBalUnities, IIf(ResultadoHorizontal_Opcional, eqlBalHorizontal, eqlBalVertical), False) ', IncluirUnidadesDeComplejos_Opcional)

End Function

Public Function EQL_BALMET_FINOS(ByVal Alimentacion As Variant, ByVal LeyesConcentrado As Variant, Optional NombresElementos_Opcional As Variant = VBA.vbNullString, Optional UnidadesLey_Opcional As Variant = "%", Optional Metodo_Opcional As Byte = 0, Optional ResultadoHorizontal_Opcional As Boolean = False) As Variant ', Optional IncluirFinosDeComplejos_Opcional As Boolean = False) As Variant
Attribute EQL_BALMET_FINOS.VB_Description = "Obtiene los finos de un balance metalúrgico."
Attribute EQL_BALMET_FINOS.VB_ProcData.VB_Invoke_Func = " \n21"

  Call Application.Volatile(True)
  If Not exq_crt Is Nothing Then Let EQL_BALMET_FINOS = exq_crt.Balmet("balmetudf", Alimentacion, LeyesConcentrado, Empty, NombresElementos_Opcional, UnidadesLey_Opcional, IIf(Metodo_Opcional = 0, eqlBalConventional, IIf(Metodo_Opcional = 1, eqlBalCramer, eqlBalInverseMatrix)), eqlBalFines, IIf(ResultadoHorizontal_Opcional, eqlBalHorizontal, eqlBalVertical), False) ', IncluirFinosDeComplejos_Opcional)

End Function

Public Function EQL_BALMET_RECUPERACION(ByVal Alimentacion As Variant, ByVal LeyesConcentrado As Variant, Optional NombresElementos_Opcional As Variant = VBA.vbNullString, Optional UnidadesLey_Opcional As Variant = "%", Optional Metodo_Opcional As Byte = 0, Optional ResultadoHorizontal_Opcional As Boolean = False) As Variant ', Optional IncluirRecuperacionDeComplejos_Opcional As Boolean = False) As Variant
Attribute EQL_BALMET_RECUPERACION.VB_Description = "Obtiene los porcentajes de recuperación de un balance metalúrgico."
Attribute EQL_BALMET_RECUPERACION.VB_ProcData.VB_Invoke_Func = " \n21"

  Call Application.Volatile(True)
  If Not exq_crt Is Nothing Then Let EQL_BALMET_RECUPERACION = exq_crt.Balmet("balmetudf", Alimentacion, LeyesConcentrado, Empty, NombresElementos_Opcional, UnidadesLey_Opcional, IIf(Metodo_Opcional = 0, eqlBalConventional, IIf(Metodo_Opcional = 1, eqlBalCramer, eqlBalInverseMatrix)), eqlBalRecoveries, IIf(ResultadoHorizontal_Opcional, eqlBalHorizontal, eqlBalVertical), False) ', IncluirRecuperacionDeComplejos_Opcional)

End Function

Public Function EQL_BALMET_RATIO(ByVal Alimentacion As Variant, ByVal LeyesConcentrado As Variant, Optional NombresElementos_Opcional As Variant = VBA.vbNullString, Optional UnidadesLey_Opcional As Variant = "%", Optional Metodo_Opcional As Byte = 0, Optional ResultadoHorizontal_Opcional As Boolean = False) As Variant
Attribute EQL_BALMET_RATIO.VB_Description = "Obtiene el Radio de Concentración de un balance metalúrgico."
Attribute EQL_BALMET_RATIO.VB_ProcData.VB_Invoke_Func = " \n21"

  Call Application.Volatile(True)
  If Not exq_crt Is Nothing Then Let EQL_BALMET_RATIO = exq_crt.Balmet("balmetudf", Alimentacion, LeyesConcentrado, Empty, NombresElementos_Opcional, UnidadesLey_Opcional, IIf(Metodo_Opcional = 0, eqlBalConventional, IIf(Metodo_Opcional = 1, eqlBalCramer, eqlBalInverseMatrix)), eqlBalRatio, IIf(ResultadoHorizontal_Opcional, eqlBalHorizontal, eqlBalVertical), False)

End Function

Public Function EQL_BALMET_CABEZAS_ENSAYADA_CALCULADA_COMPLEJO(ByVal Alimentacion As Variant, ByVal LeyesConcentrado As Variant, ByVal LeyesConcentradoComplejo As Variant, Optional NombresElementos_Opcional As Variant = VBA.vbNullString, Optional UnidadesLey_Opcional As Variant = "%", Optional Metodo_Opcional As Byte = 0, Optional ResultadoHorizontal_Opcional As Boolean = False) As Variant
Attribute EQL_BALMET_CABEZAS_ENSAYADA_CALCULADA_COMPLEJO.VB_Description = "Obtiene leyes de cabeza calculada y ensayada de material complejo en un balance metalúrgico."
Attribute EQL_BALMET_CABEZAS_ENSAYADA_CALCULADA_COMPLEJO.VB_ProcData.VB_Invoke_Func = " \n21"

  Call Application.Volatile(True)
  If Not exq_crt Is Nothing Then Let EQL_BALMET_CABEZAS_ENSAYADA_CALCULADA_COMPLEJO = exq_crt.Balmet("balmetudf", Alimentacion, LeyesConcentrado, LeyesConcentradoComplejo, NombresElementos_Opcional, UnidadesLey_Opcional, IIf(Metodo_Opcional = 0, eqlBalConventional, IIf(Metodo_Opcional = 1, eqlBalCramer, eqlBalInverseMatrix)), eqlBalGradesHeadsCx, IIf(ResultadoHorizontal_Opcional, eqlBalHorizontal, eqlBalVertical), True)

End Function

Public Function EQL_BALMET_UNIDADES_COMPLEJO(ByVal Alimentacion As Variant, ByVal LeyesConcentrado As Variant, ByVal LeyesConcentradoComplejo As Variant, Optional NombresElementos_Opcional As Variant = VBA.vbNullString, Optional UnidadesLey_Opcional As Variant = "%", Optional Metodo_Opcional As Byte = 0, Optional ResultadoHorizontal_Opcional As Boolean = False) As Variant
Attribute EQL_BALMET_UNIDADES_COMPLEJO.VB_Description = "Obtiene las unidades de material complejo en un balance metalúrgico."
Attribute EQL_BALMET_UNIDADES_COMPLEJO.VB_ProcData.VB_Invoke_Func = " \n21"

  Call Application.Volatile(True)
  If Not exq_crt Is Nothing Then Let EQL_BALMET_UNIDADES_COMPLEJO = exq_crt.Balmet("balmetudf", Alimentacion, LeyesConcentrado, LeyesConcentradoComplejo, NombresElementos_Opcional, UnidadesLey_Opcional, IIf(Metodo_Opcional = 0, eqlBalConventional, IIf(Metodo_Opcional = 1, eqlBalCramer, eqlBalInverseMatrix)), eqlBalUnitiesCx, IIf(ResultadoHorizontal_Opcional, eqlBalHorizontal, eqlBalVertical), True)

End Function

Public Function EQL_BALMET_FINOS_COMPLEJO(ByVal Alimentacion As Variant, ByVal LeyesConcentrado As Variant, ByVal LeyesConcentradoComplejo As Variant, Optional NombresElementos_Opcional As Variant = VBA.vbNullString, Optional UnidadesLey_Opcional As Variant = "%", Optional Metodo_Opcional As Byte = 0, Optional ResultadoHorizontal_Opcional As Boolean = False) As Variant
Attribute EQL_BALMET_FINOS_COMPLEJO.VB_Description = "Obtiene los finos de material complejo en un balance metalúrgico."
Attribute EQL_BALMET_FINOS_COMPLEJO.VB_ProcData.VB_Invoke_Func = " \n21"

  Call Application.Volatile(True)
  If Not exq_crt Is Nothing Then Let EQL_BALMET_FINOS_COMPLEJO = exq_crt.Balmet("balmetudf", Alimentacion, LeyesConcentrado, LeyesConcentradoComplejo, NombresElementos_Opcional, UnidadesLey_Opcional, IIf(Metodo_Opcional = 0, eqlBalConventional, IIf(Metodo_Opcional = 1, eqlBalCramer, eqlBalInverseMatrix)), eqlBalFinesCx, IIf(ResultadoHorizontal_Opcional, eqlBalHorizontal, eqlBalVertical), True)

End Function

Public Function EQL_BALMET_RECUPERACION_COMPLEJO(ByVal Alimentacion As Variant, ByVal LeyesConcentrado As Variant, ByVal LeyesConcentradoComplejo As Variant, Optional NombresElementos_Opcional As Variant = VBA.vbNullString, Optional UnidadesLey_Opcional As Variant = "%", Optional Metodo_Opcional As Byte = 0, Optional ResultadoHorizontal_Opcional As Boolean = False) As Variant
Attribute EQL_BALMET_RECUPERACION_COMPLEJO.VB_Description = "Obtiene los porcentajes de recuperación de material complejo en un balance metalúrgico."
Attribute EQL_BALMET_RECUPERACION_COMPLEJO.VB_ProcData.VB_Invoke_Func = " \n21"

  Call Application.Volatile(True)
  If Not exq_crt Is Nothing Then Let EQL_BALMET_RECUPERACION_COMPLEJO = exq_crt.Balmet("balmetudf", Alimentacion, LeyesConcentrado, LeyesConcentradoComplejo, NombresElementos_Opcional, UnidadesLey_Opcional, IIf(Metodo_Opcional = 0, eqlBalConventional, IIf(Metodo_Opcional = 1, eqlBalCramer, eqlBalInverseMatrix)), eqlBalRecoveriesCx, IIf(ResultadoHorizontal_Opcional, eqlBalHorizontal, eqlBalVertical), True)

End Function

' Economics balmets
Public Function EQL_BALMET_ECONOMICO_LEY_CABEZA(ByVal Alimentacion As Variant, ByVal PesoConcentrado As Variant, ByVal LeyConcentrado As Variant, ByVal Recuperacion As Variant, Optional UnidadesLey_Opcional As Variant = "%", Optional LeyComplejo_Opcional As Variant = 0, Optional RecuperacionComplejo_Opcional As Variant = 0, Optional Resultado_Base_Complejo_Ambos_Opcional As Long = 0, Optional Resultado_en_Vertical_Opcional As Boolean = False) As Variant
Attribute EQL_BALMET_ECONOMICO_LEY_CABEZA.VB_Description = "Obtiene las leyes de cabeza de un balance metalúrgico."
Attribute EQL_BALMET_ECONOMICO_LEY_CABEZA.VB_ProcData.VB_Invoke_Func = " \n21"

  Call Application.Volatile(True)
  If Not exq_crt Is Nothing Then Let EQL_BALMET_ECONOMICO_LEY_CABEZA = exq_crt.BalmetEco("balmetudfeco", Alimentacion, PesoConcentrado, LeyConcentrado, Recuperacion, LeyComplejo_Opcional, RecuperacionComplejo_Opcional, UnidadesLey_Opcional, eqlBalHeadsGrades, IIf(Resultado_Base_Complejo_Ambos_Opcional <= 0, eqlBalJustGrades, IIf(Resultado_Base_Complejo_Ambos_Opcional = 1, eqlBalJustGradesCx, eqlBalJustGradesBoth)), IIf(Resultado_en_Vertical_Opcional, eqlBalVertical, eqlBalHorizontal))

End Function

Public Function EQL_BALMET_ECONOMICO_LEY_PRODUCTOS(ByVal Alimentacion As Variant, ByVal PesoConcentrado As Variant, ByVal LeyCabezas As Variant, ByVal Recuperacion As Variant, Optional UnidadesLey_Opcional As Variant = "%", Optional LeyCabezasComplejo_Opcional As Variant = 0, Optional RecuperacionComplejo_Opcional As Variant = 0, Optional Resultado_Base_Complejo_Ambos_Opcional As Long = 0, Optional Resultado_en_Vertical_Opcional As Boolean = False) As Variant
Attribute EQL_BALMET_ECONOMICO_LEY_PRODUCTOS.VB_Description = "Obtiene las leyes de productos de un balance metalúrgico."
Attribute EQL_BALMET_ECONOMICO_LEY_PRODUCTOS.VB_ProcData.VB_Invoke_Func = " \n21"

  Call Application.Volatile
  If Not exq_crt Is Nothing Then Let EQL_BALMET_ECONOMICO_LEY_PRODUCTOS = exq_crt.BalmetEco("balmetudfeco", Alimentacion, PesoConcentrado, LeyCabezas, Recuperacion, LeyCabezasComplejo_Opcional, RecuperacionComplejo_Opcional, UnidadesLey_Opcional, eqlBalProdsGrades, IIf(Resultado_Base_Complejo_Ambos_Opcional <= 0, eqlBalJustGrades, IIf(Resultado_Base_Complejo_Ambos_Opcional = 1, eqlBalJustGradesCx, eqlBalJustGradesBoth)), IIf(Resultado_en_Vertical_Opcional, eqlBalVertical, eqlBalHorizontal))

End Function

' Projections
Public Function EQL_PROYECCION_PROYECTO(ByVal FlujoNetoDeCaja As Variant) As Variant
Attribute EQL_PROYECCION_PROYECTO.VB_Description = "Devuelve 'Proyecto rentable' o 'Proyecto inviable' acorde a la tendencia del flujo neto de caja."
Attribute EQL_PROYECCION_PROYECTO.VB_ProcData.VB_Invoke_Func = " \n21"

  Call Application.Volatile
  If Not exq_crt Is Nothing Then Let EQL_PROYECCION_PROYECTO = exq_crt.Projections(FlujoNetoDeCaja)

End Function

' Miscellanious
Public Function EQL_NUMERO_A_TEXTO(ByVal Numero As Variant, Optional Divisa_Singular As String = VBA.vbNullString, Optional Divisa_Multiple As String = VBA.vbNullString, Optional Centavos_Literales As Boolean = False) As Variant
Attribute EQL_NUMERO_A_TEXTO.VB_Description = "Obtiene una expresión literal del número dado."
Attribute EQL_NUMERO_A_TEXTO.VB_ProcData.VB_Invoke_Func = " \n21"

  Call Application.Volatile
  If Not exq_crt Is Nothing Then Let EQL_NUMERO_A_TEXTO = exq_crt.NumberToStringGet(Numero, Divisa_Singular, Divisa_Multiple, Centavos_Literales)

End Function
' UDF'S functions ------------------------------------------------------------------ '
' PUBLIC METHODS LIST (INTERFACE) -------------------------------------------------- '
' =========================== INDEX UI MODULE STRUCTURE ============================ '


