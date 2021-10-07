VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassEqlControllerHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ================= CLASS EXCLIQ LITE CONTROLLER HELPER STRUCTURE ================== '
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


' ================= CLASS EXCLIQ LITE CONTROLLER HELPER STRUCTURE ================== '
' OBJECT VARIABLES (GLOOBJ_) ------------------------------------------------------- '
' PROPERTIES ----------------------------------------------------------------------- '
' ================= CLASS EXCLIQ LITE CONTROLLER HELPER STRUCTURE ================== '



' ================= CLASS EXCLIQ LITE CONTROLLER HELPER STRUCTURE ================== '
' OBJECT VARIABLES (GLOOBJ_) ------------------------------------------------------- '
Private PVecTitle As Variant
Private PVecInvoice As Variant
Private PVecProvider As Variant
Private PVecOre As Variant
Private PVecOrePayment As Variant
Private PVecGrossValue As Variant
Private PVecDeductions As Variant
Private PVecBonus As Variant
Private PVecFinal As Variant
Private PVecUnits As Variant
Private PBooBonus As Boolean

'CNC
Private CNCWeightVec As Variant
Private CNCContentsVec As Variant
Private CNCAssessmentVec As Variant
Private CNCUnitsCtzVec As Variant
Private CNCUnitsGraVec As Variant
Private CNCUnitsTraVec As Variant
Private CNCBooDeductions As Boolean



' PROPERTIES ----------------------------------------------------------------------- '
Public Property Let Title(ByVal VecTitle As Variant)

  Let PVecTitle = VecTitle

End Property

Public Property Get Title() As Variant

  Let Title = PVecTitle

End Property


Public Property Let Invoice(ByVal VecInvoice As Variant)

  Let PVecInvoice = VecInvoice

End Property

Public Property Get Invoice() As Variant

  Let Invoice = PVecInvoice

End Property


Public Property Let Provider(ByVal VecProvider As Variant)

  Let PVecProvider = VecProvider

End Property

Public Property Get Provider() As Variant

  Let Provider = PVecProvider

End Property


Public Property Let Ore(ByVal VecOre As Variant)

  Let PVecOre = VecOre

End Property

Public Property Get Ore() As Variant

  Let Ore = PVecOre

End Property


Public Property Let OrePayment(ByVal VecOrePayment As Variant)

  Let PVecOrePayment = VecOrePayment

End Property

Public Property Get OrePayment() As Variant

  Let OrePayment = PVecOrePayment

End Property


Public Property Let GrossValue(ByVal VecGrossValue As Variant)

  Let PVecGrossValue = VecGrossValue

End Property

Public Property Get GrossValue() As Variant

  Let GrossValue = PVecGrossValue

End Property


Public Property Let Deductions(ByVal VecDeductions As Variant)

  Let PVecDeductions = VecDeductions

End Property

Public Property Get Deductions() As Variant

  Let Deductions = PVecDeductions

End Property


Public Property Let Bonus(ByVal VecBonus As Variant)

  Let PVecBonus = VecBonus

End Property

Public Property Get Bonus() As Variant

  Let Bonus = PVecBonus

End Property


Public Property Let Final(ByVal VecFinal As Variant)

  Let PVecFinal = VecFinal

End Property

Public Property Get Final() As Variant

  Let Final = PVecFinal

End Property


Public Property Let Units(ByVal VecUnits As Variant)

  Let PVecUnits = VecUnits

End Property

Public Property Get Units() As Variant

  Let Units = PVecUnits

End Property


Public Property Let BooBonus(ByVal cBooBonus As Boolean)

  Let PBooBonus = cBooBonus

End Property

Public Property Get BooBonus() As Boolean

  Let BooBonus = PBooBonus

End Property

'CNC
Public Property Let WeightCNC(ByVal VecWeightCNC As Variant)

  Let CNCWeightVec = VecWeightCNC

End Property

Public Property Get WeightCNC() As Variant

  Let WeightCNC = CNCWeightVec

End Property

Public Property Let ContentsCNC(ByVal VecContentsCNC As Variant)

  Let CNCContentsVec = VecContentsCNC

End Property

Public Property Get ContentsCNC() As Variant

  Let ContentsCNC = CNCContentsVec

End Property

Public Property Let AssessmentCNC(ByVal VecAssessmentCNC As Variant)

  Let CNCAssessmentVec = VecAssessmentCNC

End Property

Public Property Get AssessmentCNC() As Variant

  Let AssessmentCNC = CNCAssessmentVec

End Property

Public Property Let UnitsCtzCNC(ByVal VecUnitsCtzCNC As Variant)

  Let CNCUnitsCtzVec = VecUnitsCtzCNC

End Property

Public Property Get UnitsCtzCNC() As Variant

  Let UnitsCtzCNC = CNCUnitsCtzVec

End Property

Public Property Let UnitsGraCNC(ByVal VecUnitsGraCNC As Variant)

  Let CNCUnitsGraVec = VecUnitsGraCNC

End Property

Public Property Get UnitsGraCNC() As Variant

  Let UnitsGraCNC = CNCUnitsGraVec

End Property

Public Property Let UnitsTraCNC(ByVal VecUnitsTraCNC As Variant)

  Let CNCUnitsTraVec = VecUnitsTraCNC

End Property

Public Property Get UnitsTraCNC() As Variant

  Let UnitsTraCNC = CNCUnitsTraVec

End Property

Public Property Let BooDeductionsCNC(ByVal cBooDed As Boolean)

  Let CNCBooDeductions = cBooDed

End Property

Public Property Get BooDeductionsCNC() As Boolean

  Let BooDeductionsCNC = CNCBooDeductions

End Property
' ================= CLASS EXCLIQ LITE CONTROLLER HELPER STRUCTURE ================== '

