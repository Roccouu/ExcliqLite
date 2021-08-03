VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassPurchase_current"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ========================= CLASS ORE PURCHASE STRUCTURE ============================'
' VARIABLES (GLOStr_) ---------------------------------------------------------------'
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

'    x Let .PTitle.Vector = Empty
'    x Let .PInvoice.Vector = Empty
'    x Let .PProvider.Vector = Empty
'    Let .PWeight.Vector = Empty
'    Let .PContents.Vector = Empty
'    Let .PAssessmentMain.Vector = Empty
'    x Let .PDeductions.Vector = Empty
'    x Let .PBonus.Vector = Empty
'    x Let .PFinal.Vector = Empty
'    Let .UnitsCtzVec = Empty
'    Let .UnitsGraVec = Empty
'    Let .UnitsTraVec = Empty



' ========================= CLASS ORE PURCHASE STRUCTURE ============================'
' VARIABLES (GLOStr_) ---------------------------------------------------------------'
' PROPERTIES ------------------------------------------------------------------------'
' ========================= CLASS ORE PURCHASE STRUCTURE ============================'




' PROPERTIES ------------------------------------------------------------------------'
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

' ========================= CLASS ORE PURCHASE STRUCTURE ============================'

