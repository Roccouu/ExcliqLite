Attribute VB_Name = "IndexVars"
Option Explicit
Option Private Module


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

' ========================= INDEX VARS MODULE STRUCTURE ============================ '
' VARIABLES CUSTOM ----------------------------------------------------------------- '
' CONSTANTS (GLOSTR_) -------------------------------------------------------------- '
' ========================= INDEX VARS MODULE STRUCTURE ============================ '




' ========================= INDEX VARS MODULE STRUCTURE ============================ '
' CONSTANTS (CSTR_) -----------------------------------------------------------------'
' CONSTANTS ERRORS ------------------------------------------------------------------'
Public Const CUSTOM_ERROR_APP As Long = VBA.vbObjectError + 514
Public Const CUSTOM_ERROR_RIBBON As Long = VBA.vbObjectError + 515
Public Const CUSTOM_ERROR_MDL As Long = VBA.vbObjectError + 516
Public Const CUSTOM_ERROR_VWS As Long = VBA.vbObjectError + 517
Public Const CUSTOM_ERROR_CTR As Long = VBA.vbObjectError + 518
Public Const CUSTOM_ERROR_RES As Long = VBA.vbObjectError + 519
Public Const CUSTOM_ERROR_HLP As Long = VBA.vbObjectError + 520
Public Const CUSTOM_ERROR_VND As Long = VBA.vbObjectError + 521

'Public Enum EQLENU_CUSTOM_ERRORS ' FIXME: OPTIMIZE TO THIS WAY OF ERROR MANNAGEMENT
'  eqlCUSTOM_ERROR_APP = VBA.vbObjectError + 514
'  eqlCUSTOM_ERROR_RIBBON
'  eqlCUSTOM_ERROR_MDL
'  eqlCUSTOM_ERROR_VWS
'  eqlCUSTOM_ERROR_CTR
'  eqlCUSTOM_ERROR_RES
'  eqlCUSTOM_ERROR_HLP
'  eqlCUSTOM_ERROR_VND
'End Enum

' Balmets so on
Public Const CDBL_DM_TOP As Double = 10000
Public Const CDBL_PC_TOP As Double = 100
Public Const CDBL_GT_TOP As Double = 1000000
Public Const CDBL_MAX As Double = 1000000000
Public Const CLNG_BAL_LIMIT As Single = 10
Public Const CLNG_BAL_LIMITCX As Single = 3
Public Const CSNG_BONUS_DISCOUNT_LIMIT As Single = 5
Public Const GLOSNG_PUR_LIMIT As Single = 10
Public Const GLOSNG_PUR_PEN_LIMIT As Single = 15
Public Const GLOLNG_MAIN_COLOR As Long = 3693849 ' VBA.RGB(25,93,56)

' CUSTOM VARIABLES (EQL_TYP_|_ENU_) -------------------------------------------------'
' Eql data sheets
Public Enum EQLMDL_ENU_SHEETS
  eqlMdlSheetSys = 2 '= "excliqlitedatasheetsys"
  eqlMdlSheetPur '= "excliqlitedatasheetpur"
  eqlMdlSheetTmp '= "excliqlitedatasheettmp"
End Enum

Public Enum EQLMDL_ENU_TABLES
  eqlMdlTblConfigapp_sys = 1
  eqlMdlTblConfigcurrencies_sys
  eqlMdlTblConfigpurrm_sys
  eqlMdlTblConfigchemicalelements_sys
  eqlMdlTblConfigchemicalelementscx_sys
  eqlMdlTblConfigchemicalunits_sys
  eqlMdlTblConfigchemicalunitscx_sys
  eqlMdlTblDatatest_sys
  eqlMdlTblPurchasemin_pur = 101
  eqlMdlTblpurchasecnc_pur
  eqlMdlTblpurchasecnccontents_pur
  eqlMdlTblpurchasecncpenalties_pur
  eqlMdlTblpurchasecncdedexp_pur
  eqlMdlTblConfigpurretentions_tmp = 201
  eqlMdlTblConfigpurretentionsother_tmp
  eqlMdlTblPurbonus_tmp
  eqlMdlTblPurdiscounts_tmp
End Enum

' Main Model table response
Public Enum EQLMDL_ENU_TABLE_AS
  eqlMdlArray
  eqlMdlRange
  eqlMdlListObject
  eqlMdlStrTableName
End Enum

' Main controller
Public Enum EQLCTR_ENU_TABLE_AS
  eqlCtrGet
  eqlCtrSet
  eqlCtrShw
  eqlCtrHlp
  eqlCtrVer
End Enum

' Balmet
Public Enum EQLBAL_ENU_RESULT
  eqlBalWeights
  eqlBalWeightPercents
  eqlBalGradesHeads
  eqlBalUnities
  eqlBalFines
  eqlBalRecoveries
  eqlBalRatio
  eqlBalGradesHeadsCx
  eqlBalUnitiesCx
  eqlBalFinesCx
  eqlBalRecoveriesCx
  eqlBalVolume
  eqlBalVolumePercents
End Enum

Public Enum EQLBAL_ENU_RESULTECO
  eqlBalHeadsGrades
  eqlBalProdsGrades
End Enum

Public Enum EQLBAL_ENU_RESULTDIRECTION
  eqlBalVertical
  eqlBalHorizontal
End Enum

Public Enum EQLBAL_ENU_TYPERESULT
  eqlBalJustGrades
  eqlBalJustGradesCx
  eqlBalJustGradesBoth
End Enum

' RES
Public Enum EQLRES_ENU_VALUETYPE
  eqlResNumbers
  eqlResNotNumbers
  eqlResStrings
  eqlResDates
  eqlResRanges
  eqlResJustRanges ' Just for use in views for ranges capturing
End Enum

Public Enum EQLRES_ENU_RNGRC
  eqlResRngRow
  eqlResRngCol
End Enum

Public Enum EQLRES_ENU_DIMENSIONARRAY
  eqlRes1D
  eqlRes2D
  eqlResNoArray
  eqlResDefaultArray
End Enum

' VARIABLES CUSTOM (GLOEnu|Typ_) --------------------------------------------------- '
Public Enum EQLBAL_ENU_METHOD
  eqlBalConventional
  eqlBalCramer
  eqlBalInverseMatrix
End Enum

Public Enum EQLBAL_ENU_BOUNDS
  eqlBalProducts
  eqlBalColumns
  eqlBalRows
  eqlBalFullProducts
  eqlBalFullProductsCx
  eqlBalProductsCx
  eqlBalColumnsCx
End Enum

Public Enum EQLBAL_ENU_TYPECONTENT
  eqlBalSolids
  eqlBalVolumes
  eqlBalBoth
End Enum

Public Enum EQLBAL_ENU_TYPE
  eqlBalNormal
  eqlBalWithComplex
  eqlBalNothing
End Enum

Public Enum EQLBAL_ENU_GRADESVECTORTYPE
  eqlBalAs1D
  eqlBalAs2D
End Enum

Public Enum EQLBAL_ENU_WRONG_UNITS
  eqlBalAgUnit
  eqlBalCxElemsUnitOnNormalBalmet
  eqlBalDMOnOtherElements
  eqlBalNonCx
End Enum
' ========================= INDEX VARS MODULE STRUCTURE ============================ '

