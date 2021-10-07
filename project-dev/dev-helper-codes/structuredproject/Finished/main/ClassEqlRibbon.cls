VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassEqlRibbon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ===================== CLASS EXCLIQ LITE APP RIBBON STRUCTURE ===================== '
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


' ===================== CLASS EXCLIQ LITE APP RIBBON STRUCTURE ===================== '
' OBJECT VARIABLES (GLOOBJ_) ------------------------------------------------------- '
' VARIABLES (GLOStr_) -------------------------------------------------------------- '
' CONSTANTS ERRORS AND ERROR MESSAGES (GLOSTR_ERR_) -------------------------------- '
' CONSTANTS (GLOSTR_) -------------------------------------------------------------- '
' PROPERTIES ----------------------------------------------------------------------- '
' CONSTRUCTOR/DESTRUCTOR ----------------------------------------------------------- '
' CORE METHODS LIST PRIVATE -------------------------------------------------------- '
' Ribbon_initial_values
' Ribbon_SetStandBy
' Ribbon_Commons_Reset
' Ribbon_Enabling
' APP_NewOpenWorkbook
' APP_DisableEnableRibbon
' CONTROLS EVENTS ------------------------------------------------------------------ '
' ASSETS --------------------------------------------------------------------------- '
' ERRORS MANNAGEMENT AND ASSERTSIONS ----------------------------------------------- '
' Debug_assert
' CORE METHODS LIST (PUBLIC) ------------------------------------------------------- '
' RibbonCommonsInit
' RibbonCommonsStatusSetter
' RibbonCommonsStatusGetter
' RibbonCommonsActionsExecutorControls
' ===================== CLASS EXCLIQ LITE APP RIBBON STRUCTURE ===================== '



' ===================== CLASS EXCLIQ LITE APP RIBBON STRUCTURE ===================== '
' OBJECT VARIABLES (GLOOBJ_) ------------------------------------------------------- '
' App
Private WithEvents APP As Application
Attribute APP.VB_VarHelpID = -1
' Errors
Private EHGLOBAL As ClassEqlErrorHandler
' System Ribbon
Private EQRIBBON As IRibbonUI
' Ribbon
Private GLOOBJRIBBON As ClassEqlRibbonItems

' VARIABLES (GLOStr_) -------------------------------------------------------------- '
' System name
Private GLOStr_AppTitle As String
' Ribbon status controls
Private RIBBONINITIALIZED As Boolean
Private RIBBONSTATUS As Boolean
Private RIBBON_ENABLED_STATUS As Boolean

' CONSTANTS ERRORS AND ERROR MESSAGES (GLOSTR_ERR_) -------------------------------- '
Private Const EQPREFIX As String = "btnexcliqlite"



' PROPERTIES ----------------------------------------------------------------------- '
Public Property Set ErrorHandler(ByRef ObjEH As ClassEqlErrorHandler)

  Set EHGLOBAL = ObjEH
  Set GLOOBJRIBBON = New ClassEqlRibbonItems
  Set GLOOBJRIBBON.ErrorHandler = EHGLOBAL

End Property

Public Property Let AppTitle(ByVal PStrTitle As String)

  Let GLOStr_AppTitle = PStrTitle

End Property

Public Property Get APPRIBBONINITIALIZED() As Boolean

  Let APPRIBBONINITIALIZED = RIBBONINITIALIZED

End Property



' CONSTRUCTOR/DESTRUCTOR ----------------------------------------------------------- '
Private Sub Class_Initialize()

  If APP Is Nothing Then Set APP = Application

End Sub

Private Sub Class_Terminate()

  Set APP = Nothing
  Set GLOOBJRIBBON = Nothing
  Set EHGLOBAL = Nothing

End Sub



' CORE METHODS LIST PRIVATE -------------------------------------------------------- '
' Initialize ribbon values
Private Function Ribbon_initial_values() As Boolean

  On Error GoTo EH
  Let Ribbon_initial_values = GLOOBJRIBBON.RIBBONITEMINITIAL()

EH:
  Call EHGLOBAL.ErrorHandlerRaise("RIBBON:-initial_values")

End Function

' RIBBON StandBy
Private Sub Ribbon_SetStandBy()

  Dim VecRi As Variant
  Dim i As Long

  On Error GoTo EH
  Let VecRi = GLOOBJRIBBON.Items

  For i = 0 To UBound(VecRi)
    #If Debugging Then
      Call Debug_assert(Not VecRi(i, 0) = VBA.vbNullString, "standby")
    #End If
    Call Ribbon_Commons_Reset(EQPREFIX & VBA.StrConv(VBA.CStr(VecRi(i, 0)), vbProperCase), VecRi(i, 1))
  Next i

  Call Ribbon_udfs_register

EH:
  Let VecRi = Empty
  Call EHGLOBAL.ErrorHandlerRaise("RIBBON::-setstandby")

End Sub

' Commons
Private Sub Ribbon_Commons_Reset(ByVal ControlId As String, Optional Status As Variant)

  On Error GoTo EH
  #If Debugging Then
    Debug.Print "RESET EQRIBBON STATUS: ", ControlId, Status '(EQRIBBON Is Nothing)
  #End If

  If RIBBONSTATUS Then
    #If Debugging Then
      Debug.Print "Reseteo el control: ", ControlId & " A", Status
    #End If
    Call EQRIBBON.InvalidateControl(ControlId)
  End If

EH:
  Call EHGLOBAL.ErrorHandlerRaise("RIBBON::-commons_reset")

End Sub

' Ribbon enabling method
Private Sub Ribbon_Enabling(Optional BooEnabling As Boolean = False)

  Dim VecRi As Variant
  Dim i As Long

  On Error GoTo EH
  Let VecRi = GLOOBJRIBBON.Items

  For i = LBound(VecRi) To UBound(VecRi)
    Call GLOOBJRIBBON.RIBBONITEMENABLINGALL(BooEnabling)
  Next i

EH:
  Let VecRi = Empty
  Call EHGLOBAL.ErrorHandlerRaise("RIBBON::-enabling")

End Sub


' UDF's register table
Private Sub Ribbon_udfs_register()

  Dim MacroName() As Variant
  Dim MacroDescription() As Variant
  Dim MacroParameters() As Variant
  Dim ECategory As String
  Dim MacroNameS As String
  Dim MacroDescriptionS As String
  Dim i As Single

  On Error GoTo EH

  #If Debugging Then
    Call Debug_assert(1 = 1)
  #End If

  Let MacroName = Array( _
    "EQL_VERSION", "EQL_REGRESION_LINEAL", "EQL_REGRESION_LINEAL_a", "EQL_REGRESION_LINEAL_b", "EQL_REGRESION_LINEAL_r", "EQL_REGRESION_LINEAL_r2", "EQL_REGRESION_LINEAL_Se", "EQL_REGRESION_LINEAL_n", _
    "EQL_BALMET", "EQL_BALMET_PORCENTAJE_EN_PESO", "EQL_BALMET_VOLUMEN", "EQL_BALMET_PORCENTAJE_EN_VOLUMEN", "EQL_BALMET_CABEZAS_ENSAYADA_CALCULADA", "EQL_BALMET_UNIDADES", "EQL_BALMET_FINOS", "EQL_BALMET_RECUPERACION", "EQL_BALMET_RATIO", "EQL_BALMET_CABEZAS_ENSAYADA_CALCULADA_COMPLEJO", "EQL_BALMET_UNIDADES_COMPLEJO", "EQL_BALMET_FINOS_COMPLEJO", "EQL_BALMET_RECUPERACION_COMPLEJO", "EQL_BALMET_ECONOMICO_LEY_CABEZA", "EQL_BALMET_ECONOMICO_LEY_PRODUCTOS", _
    "EQL_PROYECCION_PROYECTO", _
    "EQL_NUMERO_A_TEXTO")

  Let MacroDescription = Array( _
     "Devuelve la versión actual de " & GLOStr_AppTitle & ".", "Obtiene la regresión lineal de tres o más datos.", "Devuelve el Coeficiente a de una regresión.", "Devuelve el Coeficiente b de una regresión.", "Devuelve el Coeficiente de correlación de una regresión.", "Devuelve el Coeficiente de determinación de una regresión.", "Devuelve el Error estándar de una regresión.", "Devuelve el total de datos que intervienen en la regresión.", _
     "Obtiene el o los pesos de productos de un balance metalúrgico.", "Obtiene el Porcentaje de en peso de un balance metalúrgico.", "Obtiene el o los volúmenes de productos de un balance metalúrgico.", _
     "Obtiene el Porcentaje de en volumen de un balance metalúrgico.", "Obtiene leyes de cabeza calculada y ensayada de un balance metalúrgico.", "Obtiene las unidades de un balance metalúrgico.", _
     "Obtiene los finos de un balance metalúrgico.", "Obtiene los porcentajes de recuperación de un balance metalúrgico.", "Obtiene el Radio de Concentración de un balance metalúrgico.", _
     "Obtiene leyes de cabeza calculada y ensayada de material complejo en un balance metalúrgico.", "Obtiene las unidades de material complejo en un balance metalúrgico.", "Obtiene los finos de material complejo en un balance metalúrgico.", "Obtiene los porcentajes de recuperación de material complejo en un balance metalúrgico.", "Obtiene las leyes de cabeza de un balance metalúrgico.", "Obtiene las leyes de productos de un balance metalúrgico.", _
     "Devuelve 'Proyecto rentable' o 'Proyecto inviable' acorde a la tendencia del flujo neto de caja.", _
     "Obtiene una expresión literal del número dado.")

  Let MacroParameters = Array(Array("Ningún parámetro."), _
      Array("Requerido. Rango de más de tres celdas con números.", "Opcional. Rango de más de tres celdas con números.", "Opcional. Número o lista de números."), _
      Array("Requerido. Rango de más de tres celdas con números.", "Opcional. Rango de más de tres celdas con números."), _
      Array("Requerido. Rango de más de tres celdas con números.", "Opcional. Rango de más de tres celdas con números."), _
      Array("Requerido. Rango de más de tres celdas con números.", "Opcional. Rango de más de tres celdas con números."), _
      Array("Requerido. Rango de más de tres celdas con números.", "Opcional. Rango de más de tres celdas con números."), _
      Array("Requerido. Rango de más de tres celdas con números.", "Opcional. Rango de más de tres celdas con números."), _
      Array("Requerido. Rango de más de tres celdas con números.", "Opcional. Rango de más de tres celdas con números."), _
    Array("Requerido. Cantidad de alimentación.", "Requerido. Leyes de productos y alimentación.", "Opcional. Nombres de minerales. Ej. Zn.", "Opcional. Unidades de ley. Ej %.", "Opcional. Número: 0: Convencional, 1: CRAMER, 2: Matriz inversa.", "Opcional. Número: 0: Vertical, 1: Horizontal.", "Opcional. Número: 0: Sólo pesos de productons, 1: Incluye peso de alimentación."), _
    Array("Requerido. Cantidad de alimentación.", "Requerido. Leyes de productos y alimentación.", "Opcional. Nombres de minerales. Ej. Zn.", "Opcional. Unidades de ley. Ej %.", "Opcional. Número: 0: Convencional, 1: CRAMER, 2: Matriz inversa.", "Opcional. Número: 0: Vertical, 1: Horizontal."), _
    Array("Requerido. Cantidad de alimentación.", "Requerido. Leyes de productos y alimentación.", "Opcional. Nombres de minerales. Ej. Zn.", "Opcional. Unidades de ley. Ej %.", "Opcional. Número: 0: Convencional, 1: CRAMER, 2: Matriz inversa.", "Opcional. Número: 0: Vertical, 1: Horizontal."), _
    Array("Requerido. Cantidad de alimentación.", "Requerido. Leyes de productos y alimentación.", "Opcional. Nombres de minerales. Ej. Zn.", "Opcional. Unidades de ley. Ej g/L.", "Opcional. Número: 0: Convencional, 1: CRAMER, 2: Matriz inversa.", "Opcional. Número: 0: Vertical, 1: Horizontal."), _
    Array("Requerido. Cantidad de alimentación.", "Requerido. Leyes de productos y alimentación.", "Opcional. Nombres de minerales. Ej. Zn.", "Opcional. Unidades de ley. Ej %.", "Opcional. Número: 0: Convencional, 1: CRAMER, 2: Matriz inversa.", "Opcional. Número: 0: Vertical, 1: Horizontal."), _
    Array("Requerido. Cantidad de alimentación.", "Requerido. Leyes de productos y alimentación.", "Opcional. Nombres de minerales. Ej. Zn.", "Opcional. Unidades de ley. Ej %.", "Opcional. Número: 0: Convencional, 1: CRAMER, 2: Matriz inversa.", "Opcional. Número: 0: Vertical, 1: Horizontal."), _
    Array("Requerido. Cantidad de alimentación.", "Requerido. Leyes de productos y alimentación.", "Opcional. Nombres de minerales. Ej. Zn.", "Opcional. Unidades de ley. Ej %.", "Opcional. Número: 0: Convencional, 1: CRAMER, 2: Matriz inversa.", "Opcional. Número: 0: Vertical, 1: Horizontal."), _
    Array("Requerido. Cantidad de alimentación.", "Requerido. Leyes de productos y alimentación.", "Opcional. Nombres de minerales. Ej. Zn.", "Opcional. Unidades de ley. Ej %.", "Opcional. Número: 0: Convencional, 1: CRAMER, 2: Matriz inversa.", "Opcional. Número: 0: Vertical, 1: Horizontal."), _
    Array("Requerido. Cantidad de alimentación.", "Requerido. Leyes de productos y alimentación.", "Opcional. Nombres de minerales. Ej. Zn.", "Opcional. Unidades de ley. Ej %.", "Opcional. Número: 0: Convencional, 1: CRAMER, 2: Matriz inversa.", "Opcional. Número: 0: Vertical, 1: Horizontal."), _
    Array("Requerido. Cantidad de alimentación.", "Requerido. Leyes de productos y alimentación.", "Requerido. Leyes de productos y alimentación mineral complejo.", "Opcional. Nombres de minerales. Ej. Zn;Ag.", "Opcional. Unidades de ley. Ej %;DM.", "Opcional. Número: 0: Convencional, 1: CRAMER, 2: Matriz inversa.", "Opcional. Número: 0: Vertical, 1: Horizontal."), _
    Array("Requerido. Cantidad de alimentación.", "Requerido. Leyes de productos y alimentación.", "Requerido. Leyes de productos y alimentación mineral complejo.", "Opcional. Nombres de minerales. Ej. Zn;Ag.", "Opcional. Unidades de ley. Ej %;DM.", "Opcional. Número: 0: Convencional, 1: CRAMER, 2: Matriz inversa.", "Opcional. Número: 0: Vertical, 1: Horizontal."), _
    Array("Requerido. Cantidad de alimentación.", "Requerido. Leyes de productos y alimentación.", "Requerido. Leyes de productos y alimentación mineral complejo.", "Opcional. Nombres de minerales. Ej. Zn;Ag.", "Opcional. Unidades de ley. Ej %;DM.", "Opcional. Número: 0: Convencional, 1: CRAMER, 2: Matriz inversa.", "Opcional. Número: 0: Vertical, 1: Horizontal."), _
    Array("Requerido. Cantidad de alimentación.", "Requerido. Leyes de productos y alimentación.", "Requerido. Leyes de productos y alimentación mineral complejo.", "Opcional. Nombres de minerales. Ej. Zn;Ag.", "Opcional. Unidades de ley. Ej %;DM.", "Opcional. Número: 0: Convencional, 1: CRAMER, 2: Matriz inversa.", "Opcional. Número: 0: Vertical, 1: Horizontal."), _
    Array("Requerido. Cantidad de alimentación.", "Requerido. Pesos aproximados de productos.", "Requerido. Leyes aproximadas de productos.", "Requerido. Porcentajes aproximados de recuperación.", "Opcional. Unidades de ley. Ej %;DM.", "Opcional. Ley aproximada de producto mineral complejo.", "Opcional. Porcentaje aproximado de recuperació mineral complejo.", "Opcional. Número: 0: Retorna resultado principal, 1: Incluye resultado de mineral complejo.", "Opcional. Número: 0: Vertical, 1: Horizontal."), _
    Array("Requerido. Cantidad de alimentación.", "Requerido. Pesos aproximados de productos.", "Requerido. Leyes aproximadas de cabeza.", "Requerido. Porcentajes aproximados de recuperación.", "Opcional. Unidades de ley. Ej %;DM.", "Opcional. Ley aproximada de producto mineral complejo.", "Opcional. Porcentaje aproximado de recuperació mineral complejo.", "Opcional. Número: 0: Retorna resultado principal, 1: Incluye resultado de mineral complejo.", "Opcional. Número: 0: Vertical, 1: Horizontal."), _
      Array("Rango de celdas con números"), _
    Array("Número", "Opcional. Nombre de divisa. Ej. Boliviano", "Opcional. Nombre de divisa en plural. Ej. Bolivianos", "Opcional. Número: 0: Expresión numérica de decimales, 1: Expresión literal de decimales."))

  Let ECategory = GLOStr_AppTitle '"ExcliqLite"

  For i = LBound(MacroName) To UBound(MacroName)
    Let MacroNameS = VBA.CStr(MacroName(i))
    Let MacroDescriptionS = VBA.CStr(MacroDescription(i))
    ' Application.MacroOptions MacroName(i), MacroDescription(i), ECategory:=ECategory, argumentdescriptions:=MacroParameters
    Call Application.MacroOptions(macro:=MacroNameS, Description:=MacroDescriptionS, Category:=ECategory, ArgumentDescriptions:=MacroParameters(i))
  Next i

EH:
  Erase MacroParameters
  Erase MacroDescription
  Erase MacroName
  Call EHGLOBAL.ErrorHandlerRaise("RIBBON::-udfs_register")

End Sub


' APPLICATION MAIN EVENTS TO SETTING BASIC AND MOST IMPORTANT SYSTEM RESOURCES
' Methods Private
Private Sub APP_NewOpenWorkbook(ByVal Wb As Workbook, Optional NewOpen As Boolean = False)

  Const sFN As String = "RIBBON::-new_open_workbook"

  #If Debugging Then
    Call VBA.MsgBox(IIf(NewOpen, "Open", "New") & " Workbook is: " & Wb.Name)
  #End If

  On Error GoTo EH
  If Not EQRIBBON Is Nothing Then
    If Not RIBBONSTATUS And Application.Workbooks.Count > 0 Then
      #If Debugging Then
        Call VBA.MsgBox("Nuevos libros abiertos, habilitar ribbon")
      #End If
      Call Ribbon_Enabling(BooEnabling:=True)
      Call APP_DisableEnableRibbon(DisableEnable:=True)
    End If
  Else
    Call VBA.Err.Raise(CUSTOM_ERROR_RIBBON, sFN, "No existe Objeto Ribbon para iniciar Sistema.")
  End If

EH:
  Call EHGLOBAL.ErrorHandlerRaise(sFN)

End Sub

Private Sub APP_DisableEnableRibbon(Optional DisableEnable As Boolean = False)

  Dim vRibbon As Variant
  Dim i As Long

  On Error GoTo EH
  Let vRibbon = GLOOBJRIBBON.Items
  For i = LBound(vRibbon) To UBound(vRibbon)
    Call Ribbon_Commons_Reset(vRibbon(i, 0), vRibbon(i, 1))
  Next i
  Let RIBBONSTATUS = DisableEnable
  Let RIBBON_ENABLED_STATUS = DisableEnable

EH:
  Let vRibbon = Empty
  Call EHGLOBAL.ErrorHandlerRaise("RIBBON::-disable_enable_ribbon")

End Sub



' CONTROLS EVENTS ------------------------------------------------------------------ '
Private Sub APP_NewWorkbook(ByVal Wb As Workbook)

  'Check for Ribbon
  Call APP_NewOpenWorkbook(Wb)

End Sub

Private Sub APP_WorkbookOpen(ByVal Wb As Workbook)

  ' Check for Ribbon
  Call APP_NewOpenWorkbook(Wb, True)

End Sub

Private Sub APP_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)

  Const sFN As String = "RIBBON::APP:-BeforeClose"

  On Error GoTo EH
  If Not EQRIBBON Is Nothing And Not Wb Is Nothing Then

    With Wb
      If .IsAddin Then
        If .Name = ThisWorkbook.Name Then
          ' Save this addin when closing
          If Not Cancel Then
            #If Debugging Then
              Call VBA.MsgBox("Guardando cambios en " & .Name & " antes de cerrar")
            #End If
            Let Application.ScreenUpdating = False
            Let Application.DisplayAlerts = False
            Call .Save
            Let Application.ScreenUpdating = True
            Let Application.DisplayAlerts = True
          End If
        End If
      Else
        If Not Cancel Then
          #If Debugging Then
            Call VBA.MsgBox("Existen " & Application.Workbooks.Count & " libros por cerrar.")
          #End If
          If Application.Workbooks.Count <= 0 And RIBBON_ENABLED_STATUS Then
            Call Ribbon_Enabling(BooEnabling:=False)
            Call APP_DisableEnableRibbon(DisableEnable:=False)
          End If
        End If
      End If
    End With

  Else

    Call VBA.Err.Raise(CUSTOM_ERROR_RIBBON, sFN, "No existe Objeto Ribbon para iniciar Sistema.")

  End If

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
' Commons
Public Sub RibbonCommonsInit(ByVal Ribbon As IRibbonUI)

  Const sFN As String = "RIBBON::Init"

  On Error GoTo EH
  Let RIBBONSTATUS = (Not Ribbon Is Nothing)
  If RIBBONSTATUS Then
    Set EQRIBBON = Ribbon

    #If Debugging Then
      Call VBA.MsgBox("Aplicación" & IIf(RIBBONINITIALIZED, " iniciada.", " iniciando."), vbInformation, Application.Name)
    #End If

    ' Initialize System included Ribbon just if System is started for first time. This is to work with only one instance.
    If Not RIBBONINITIALIZED Then

      #If Debugging Then
        Call VBA.MsgBox(IIf(RIBBONINITIALIZED, "Continúo la ", "Inicio la ") & "Aplicación", vbInformation, Application.Name)
      #End If

      Let RIBBONINITIALIZED = Ribbon_initial_values()
      Let RIBBON_ENABLED_STATUS = RIBBONINITIALIZED

      #If Debugging Then
        Call VBA.MsgBox("Estableciendo Ribbon en modo de espera.", vbInformation, GLOStr_AppTitle)
      #End If

      Call Ribbon_SetStandBy 'Setting Ribbon values to Stand by.
    End If
  Else
    Call VBA.Err.Raise(CUSTOM_ERROR_RIBBON, sFN, "No existe Objeto Ribbon para iniciar Sistema.")
  End If

  #If Debugging Then
    Debug.Print "Primero"
  #End If

EH:
  Call EHGLOBAL.ErrorHandlerDisplay(sFN)

End Sub

Public Function RibbonCommonsStatusSetter(ByVal ControlId As String, Optional GetVisibleEnable As Boolean = False) As Variant

  On Error GoTo EH
  #If Debugging Then
    Debug.Print "Segundo", ControlId
  #End If

  ' Ribbon controls enabling status setter
  If RIBBONINITIALIZED Then
    If GetVisibleEnable Then ' Enable/Disable
      Let RibbonCommonsStatusSetter = GLOOBJRIBBON.RIBBONITEMGET(ControlId)
    Else ' Visibility
      'Let RibbonCommonsStatusSetter = uRIBBONCOL.Item(ControlId)
    End If
  End If
EH:
  Call EHGLOBAL.ErrorHandlerDisplay("RIBBON::CommonsStatusSetter")

End Function

Public Function RibbonCommonsNoWorkbooksOpenend() As Boolean

  Dim BooAux0 As Boolean
  Dim i As Long, j As Long

  On Error GoTo EH

  ' Check for opened workbooks in collection
  With Application

    Let j = 0
    For i = 1 To .Workbooks.Count
      Let j = j + IIf(.Workbooks(i).IsAddin, 0, 1)
    Next i
    Let BooAux0 = (j > 0)

    If BooAux0 Then ' Exists opened workbooks
      If Not RIBBON_ENABLED_STATUS Then ' Enable
        ' If ribbon is disabled, enable all
        Call Ribbon_Enabling(BooEnabling:=True)
        Call APP_DisableEnableRibbon(DisableEnable:=True)
      End If
    Else
      If RIBBON_ENABLED_STATUS Then ' Disable
        ' Disable ribbon
        Call Ribbon_Enabling(BooEnabling:=False)
        Call APP_DisableEnableRibbon(DisableEnable:=False)
      End If
    End If

  End With

  Let RibbonCommonsNoWorkbooksOpenend = BooAux0

EH:
  Call EHGLOBAL.ErrorHandlerDisplay("RIBBON::CommonsNoWorkbooksOpenend")

End Function
' CORE METHODS LIST (PUBLIC) ------------------------------------------------------- '
' ===================== CLASS EXCLIQ LITE APP RIBBON STRUCTURE ===================== '


