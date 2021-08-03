Attribute VB_Name = "Módulo1"
Option Explicit


Sub Curget_Testsimple()
  Dim MLT As ModelExcliqliteDatasheet
  Dim ErrH As AppErrorHandler
  Dim RES As AppResources_current
  
  Dim VecAux0 As Variant
  Dim GLOSTR_CURFOREIGN As String
  
  Set ErrH = New AppErrorHandler
  
  Set RES = New AppResources_current
  Set RES.ErrorHandler = ErrH
  
  Set MLT = New ModelExcliqliteDatasheet
  Set MLT.ErrorHandler = ErrH
  
  Let VecAux0 = RES.ArrayToBaseZero(MLT.MGet(eqlMdlTblCurrencies, eqlMdlArray, "currency, currencies, symbol", MStrWhere:="currency_main=1"), eqlRes1D)
  Let GLOSTR_CURFOREIGN = VBA.CStr(MLT.MGet(eqlMdlTblCurrencies, eqlMdlArray, "symbol", MStrWhere:="currency_foreign=1")(0, 0))
  
  Let VecAux0 = Empty
  Set ErrH = Nothing
  Set MLT = Nothing
End Sub







Sub OrepurAux_Testsimple()
  Dim RES As AppResources_current
  Dim REGEX As AppResRegEx
  Dim Err As AppErrorHandler
  Dim MLT As ModelExcliqliteDatasheet
  Dim VecAux0 As Variant, VecAux1 As Variant, VecAux2 As Variant, VecAux3 As Variant, VecGrossA As Variant, VecGrossB As Variant, VecGross As Variant
  
  Dim i As Long, j As Long, k As Long, l As Long, u As Long, v As Long, w As Long
  Dim m As Long, n As Long
  
  Set Err = New AppErrorHandler
  Set RES = New AppResources_current
  Set RES.ErrorHandler = Err
  Set MLT = New ModelExcliqliteDatasheet
  Set MLT.ErrorHandler = Err
  
  Let VecAux0 = MLT.MGet(eqlMdlTblPurchasebasedata, eqlMdlArray, "PESO HÚMEDO BRUTO, HUMEDAD, MERMA", 0) 'E
  Let VecAux1 = MLT.MGet(eqlMdlTblPurchasemaindata, eqlMdlArray) 'E
  Let VecAux0 = RES.ArrayConcat(VecAux0, VecAux1, RByCols:=True)
  If Not VBA.IsArray(VecAux0) Then Debug.Print "Error": Exit Sub 'Call EHGLOBAL.ErrorHandlerSet(0, "Falló la lectura de datos para determinar el valor bruto de la liquidación"): GoTo EH
  Let VecAux1 = Empty
  
  ' 1.  Separate all complex
  'Dim VecCont As Variant, VecGrad As Variant, VecUnit As Variant, VecPrce As Variant
  'ReDim VecCont(0): ReDim VecGrad(0): ReDim VecUnit(0): ReDim VecPrce(0)

  Set REGEX = New AppResRegEx
'  Dim i As Long, j As Long, k As Long, l As Long, u As Long
  Dim BooAccurateAnyway As Boolean, BooEquals As Boolean, BooAccurate As Boolean, BOOFORMULAS As Boolean
  Dim StrAux0 As String, StrAux1 As String, StrAux2 As String, StrAux3 As String
  Let BOOFORMULAS = True
  
  Let i = 0
  Do While i <= UBound(VecAux0)
    If BOOFORMULAS Then
      Let StrAux0 = "(" & Range("I" & 416 + i).Address(False, False) & "*(1-(" & VBA.Replace(VBA.CStr(VecAux0(i, 1)), ",", ".") & "*0.01)))*(1-(" & VBA.Replace(VBA.CStr(VecAux0(i, 2)), ",", ".") & "*0.01))" 'PNS*GRD*PRC
      Debug.Print StrAux0
      Let VecAux0(i, 0) = StrAux0 '"(" & Range("I" & 416 + i).address(False, False) & "*(1-(" & VBA.Replace(VBA.CStr(VecAux0(i, 1)), ",", ".") & "*0.01))*(1-(" & VBA.Replace(VBA.CStr(VecAux1(i, 2)), ",", ".") & "*0.01)" 'PNS*GRD*PRC
    End If
    If REGEX.isMineralComplex(VBA.CStr(VecAux0(i, 3))) Then
      Let VecAux2 = VBA.Split(VBA.CStr(VecAux0(i, 4)), ";")
      Let j = UBound(VecAux2) - LBound(VecAux2)
      For k = 0 To j
        If BOOFORMULAS And k > 0 Then Debug.Print StrAux0
        Let VecAux3 = Array(VecAux0(i, 0), VecAux0(i, 1), VecAux0(i, 2), VecAux0(i, 3))
        For l = 4 To 7
          Let VecAux2 = VBA.Split(VBA.CStr(VecAux0(i, l)), ";")
          Let u = UBound(VecAux3) + 1
          ReDim Preserve VecAux3(u)
          Let VecAux3(u) = IIf(l = 4 And k > 0, "~", VBA.vbNullString) & VecAux2(k)
          Let VecAux2 = Empty
        Next l
        Let VecAux1 = RES.ArrayAddAtLast(VecAux1, VecAux3)
        Let VecAux3 = Empty
      Next k
    Else
      Let VecAux1 = RES.ArrayAddAtLast(VecAux1, RES.ArrayDelIndex(VecAux0, i, True))
    End If
    Let i = i + 1
  Loop
    
    ' 2.  Calculate PNS
    For i = 0 To UBound(VecAux1) 'PBH*H2O*LSS
      If Not BOOFORMULAS Then Let VecAux1(i, 0) = ((VBA.CDbl(VecAux1(i, 0)) * (1 - (VBA.CDbl(VecAux1(i, 1)) / 100))) * (1 - (VBA.CDbl(VecAux1(i, 2)) / 100))) '* (VBA.CDbl(VecAux1(i, 5))) * (VBA.CDbl(VecAux1(i, 7)))
    Next i
    Let VecAux0 = Empty
'    Exit Sub
  
    ' 2.  Filter by Content, Unit & [Price] (Includes formulas)
    'Dim BooAccurateAnyway As Boolean, BooEquals As Boolean, BooAccurate As Boolean, BOOFORMULAS As Boolean
    Dim DblAux0 As Double
    Dim DblTC As Double
    Dim StrPNS As String, StrGRSo As String, StrGRD As String, StrGrs As String
    'Dim StrAux0 As String, StrAux1 As String, StrAux2 As String, StrAux3 As String
    Let BooAccurateAnyway = True
    'Let BOOFORMULAS = True
    Let DblTC = 6.96
    Let k = 1
    Let n = 0
    
    
    ReDim VecAux3(0 To 0, 0 To 4)
    If BOOFORMULAS Then
      For i = 0 To UBound(VecAux1)
        Let BooAccurate = False
        Let l = 1
        
        Let StrAux0 = VBA.Replace(VBA.CStr(VecAux1(i, 4)), "~", VBA.vbNullString)
        
        If BOOFORMULAS Then
          Let m = 0
          ReDim VecGrossA(0)
          ReDim VecGrossB(0)
          ReDim VecGross(0)
          Let StrPNS = VBA.CStr(VecAux1(i, 0)) 'Range("I" & 416 + i).address(False, False) & "*" & VBA.Replace(VBA.CStr(VecAux1(i, 5)), ",", ".") & "*" & VBA.Replace(VBA.CStr(VecAux1(i, 7)), ",", ".") 'PNS*GRD*PRC
          'PNS * PRC
          Let VecGrossA(m) = "(" & StrPNS & "*" & VBA.Replace(VBA.CStr(VecAux1(i, 5)), ",", ".") & "*" & VBA.Replace(VBA.CStr(VecAux1(i, 7)), ",", ".") & ")"
          Let VecGrossB(m) = "(" & StrPNS & ")" '*" & VBA.Replace(VBA.CStr(VecAux1(i, 7)), ",", ".") & ")" 'SUM(PNS)
          Let VecGross(m) = StrPNS
                       'Array(Name,    Grade,         Price,         Foreign, Local, PNS,           Unit)
          Let VecAux2 = Array(StrAux0, VecAux1(i, 5), VecAux1(i, 7), Empty, "=" & StrPNS, VecAux1(i, 0), VecAux1(i, 6))
        Else
          Let DblAux0 = VBA.CDbl(VecAux1(i, 0)) * VBA.CDbl(VecAux1(i, 5)) * VBA.CDbl(VecAux1(i, 7)) * DblTC 'PNS*GRD*PRC
                       'Array(Name,    Grade,         Price,         Foreign, Local, PNS,           Unit)
          Let VecAux2 = Array(StrAux0, VecAux1(i, 5), VecAux1(i, 7), Empty, DblAux0, VecAux1(i, 0), VecAux1(i, 6))
        End If
        
        For j = k To UBound(VecAux1)
          'Let BooEquals = (Not VecAux1(j, 0) = VBA.vbNullString And Not VecAux1(i, 0) = VBA.vbNullString)
          Let StrAux0 = VBA.CStr(VecAux1(j, 4)) 'Content
          Let StrAux1 = VBA.CStr(VecAux1(i, 4)) 'Content
          Let BooEquals = (Not StrAux0 = VBA.vbNullString And Not StrAux1 = VBA.vbNullString)
          If BooEquals Then
            'Let BooEquals = (VBA.LCase(VBA.CStr(VecAux1(i, 4))) = VBA.LCase(VBA.CStr(VecAux1(j, 4))) And VBA.LCase(VBA.CStr(VecAux1(i, 6))) = VBA.LCase(VBA.CStr(VecAux1(j, 6))))
            Let StrAux0 = VBA.LCase(VBA.Replace(VBA.CStr(VecAux1(i, 4)), "~", VBA.vbNullString)) 'Content
            Let StrAux1 = VBA.LCase(VBA.Replace(VBA.CStr(VecAux1(j, 4)), "~", VBA.vbNullString)) 'Content
            Let StrAux2 = VBA.LCase(VBA.CStr(VecAux1(i, 6))) 'Unit
            Let StrAux3 = VBA.LCase(VBA.CStr(VecAux1(j, 6))) 'Unit
            Let BooEquals = ((StrAux0 = StrAux1) And (StrAux2 = StrAux3))
            If BooEquals Then
              Let BooEquals = (VBA.LCase(VBA.CStr(VecAux1(i, 7))) = VBA.LCase(VBA.CStr(VecAux1(j, 7)))) 'PRCi=PRCj
              
              If BOOFORMULAS Then
                'Let BooEquals = (VBA.LCase(VBA.CStr(VecAux1(i, 7))) = VBA.LCase(VBA.CStr(VecAux1(j, 7)))) 'PRCi=PRCj
                  'Let DblAux0 = VBA.CDbl(VecAux1(j, 0)) 'PNS
                If BooEquals Or BooAccurateAnyway Then
                  Let StrPNS = VBA.CStr(VecAux1(j, 0))
                  Let m = m + 1
                  ReDim Preserve VecGrossA(m) 'PNS * GRD * PRC
                  Let VecGrossA(m) = "(" & StrPNS & "*" & VBA.Replace(VBA.CStr(VecAux1(j, 5)), ",", ".") & "*" & VBA.Replace(VBA.CStr(VecAux1(j, 7)), ",", ".") & ")"
                  ReDim Preserve VecGrossB(m) 'PNS * PRC
                  Let VecGrossB(m) = "(" & StrPNS & ")" '& "*" & VBA.Replace(VBA.CStr(VecAux1(j, 7)), ",", ".")
                  ReDim Preserve VecGross(m) 'SUM(PNS)
                  Let VecGross(m) = StrPNS
                End If
                'IIf(BooEquals Or BooAccurateAnyway, DblAux0, 0) & IIf(j = UBound(VecAux1), ")", VBA.vbNullString)
                'Let DblAux0 = VBA.CDbl(VecAux1(j, 0)) * VBA.CDbl(VecAux1(j, 5)) * VBA.CDbl(VecAux1(j, 7)) * DblTC 'GRS=PNS*GRD*PRC*TC
                'Let VecAux2(4) = VBA.CDbl(VecAux2(4)) + IIf(BooEquals Or BooAccurateAnyway, DblAux0, 0)
                Let DblAux0 = VBA.CDbl(VecAux1(j, 7)) 'PRC
                Let VecAux2(2) = VBA.CDbl(VecAux2(2)) + IIf(BooEquals Or BooAccurateAnyway, DblAux0, 0)
                'If BooEquals Or BooAccurateAnyway Then Let VecAux1(j, 4) = VBA.vbNullString
                'If BooAccurateAnyway Then Let l = l + 1
                'Let BooAccurate = (BooAccurate Or BooEquals Or BooAccurateAnyway)
              Else
                'Let BooEquals = (VBA.LCase(VBA.CStr(VecAux1(i, 7))) = VBA.LCase(VBA.CStr(VecAux1(j, 7)))) 'PRCi=PRCj
                Let DblAux0 = VBA.CDbl(VecAux1(j, 0)) 'PNS
                Let VecAux2(5) = VBA.CDbl(VecAux2(5)) + IIf(BooEquals Or BooAccurateAnyway, DblAux0, 0)
                Let DblAux0 = VBA.CDbl(VecAux1(j, 0)) * VBA.CDbl(VecAux1(j, 5)) * VBA.CDbl(VecAux1(j, 7)) * DblTC 'GRS=PNS*GRD*PRC*TC
                Let VecAux2(4) = VBA.CDbl(VecAux2(4)) + IIf(BooEquals Or BooAccurateAnyway, DblAux0, 0)
                Let DblAux0 = VBA.CDbl(VecAux1(j, 7)) 'PRC
                Let VecAux2(2) = VBA.CDbl(VecAux2(2)) + IIf(BooEquals Or BooAccurateAnyway, DblAux0, 0)
                'If BooEquals Or BooAccurateAnyway Then Let VecAux1(j, 4) = VBA.vbNullString
                'If BooAccurateAnyway Then Let l = l + 1
                'Let BooAccurate = (BooAccurate Or BooEquals Or BooAccurateAnyway)
              End If
            
              If BooEquals Or BooAccurateAnyway Then Let VecAux1(j, 4) = VBA.vbNullString
              If BooAccurateAnyway Then Let l = l + 1
              Let BooAccurate = (BooAccurate Or BooEquals Or BooAccurateAnyway)
            
            End If
          End If
        Next j
        If Not VecAux1(i, 4) = VBA.vbNullString Then
          Let VecAux1(i, 4) = VBA.vbNullString
          If l > 1 Then Let VecAux2(2) = VBA.CDbl(VecAux2(2)) / l 'PRC
          Let VecAux2(0) = n + 1 & ". " & IIf(BooAccurate, "Promed. ", VBA.vbNullString) & VBA.StrConv(VBA.CStr(VecAux2(0)), vbProperCase) 'Content
          
          If BOOFORMULAS Then
            'Let VecAux1(i, 4) = VBA.vbNullString
            'If l > 1 Then Let VecAux2(2) = VBA.CDbl(VecAux2(2)) / l 'PRC
            'Let VecAux2(0) = IIf(BooAccurate, "Promed. ", VBA.vbNullString) & VBA.StrConv(VBA.CStr(VecAux2(0)), vbProperCase) 'Content
            If BooAccurate Then
              Let VecAux2(1) = "=SUM(" & VBA.Join(VecGrossA, ",") & ")/(SUM(" & VBA.Join(VecGrossB, ",") & ")*" & VBA.Replace(VBA.CStr(VecAux2(2)), ",", ".") & ")" 'GRD
              Let VecAux2(4) = "=SUM(" & VBA.Join(VecGross, ",") & ")*" & Range("F" & 426 + i).Address(False, False) & "*" & VBA.Replace(VBA.CStr(VecAux2(2)), ",", ".") & "*" & Range("C420").Address(False, False)
              'Let VecAux3 = RES.ArrayAddAtLast(VecAux3, VecAux2)
            Else
              'Let VecAux2(1) = "=SUM(" & VBA.Join(VecGrossA, ",") & ")/(SUM(" & VBA.Join(VecGrossB, ",") & ")*" & VBA.Replace(VBA.CStr(VecAux2(2)), ",", ".") & ")" 'GRD
              Let VecAux2(4) = VBA.CStr(VecAux2(4)) & "*" & Range("F" & 426 + i).Address(False, False) & "*" & Range("G" & 426 + i).Address(False, False)
            End If
            Let VecGrossA = Empty
            Let VecGrossB = Empty
            Let VecGross = Empty
          Else
            'Let VecAux1(i, 4) = VBA.vbNullString
            'If l > 1 Then Let VecAux2(2) = VBA.CDbl(VecAux2(2)) / l
            'Let VecAux2(0) = IIf(BooAccurate, "Promed. ", VBA.vbNullString) & VBA.StrConv(VBA.CStr(VecAux2(0)), vbProperCase)
            Let VecAux2(1) = VBA.CDbl(VecAux2(4)) / (VBA.CDbl(VecAux2(2)) * VBA.CDbl(VecAux2(5)) * DblTC)
            'Let VecAux3 = RES.ArrayAddAtLast(VecAux3, VecAux2)
          End If
          Let n = n + 1
          Let VecAux3 = RES.ArrayAddAtLast(VecAux3, VecAux2)
        End If
        Let k = k + 1
      Next i
    Else
      For i = 0 To UBound(VecAux1)
        Let BooAccurate = False
        Let l = 1
        Let DblAux0 = VBA.CDbl(VecAux1(i, 0)) * VBA.CDbl(VecAux1(i, 5)) * VBA.CDbl(VecAux1(i, 7)) * DblTC 'PNS*GRD*PRC
                     'Array(Name,          Grade,         Price,         Foreign, Local, PNS, Unit)
        Let StrAux0 = VBA.Replace(VBA.CStr(VecAux1(i, 4)), "~", VBA.vbNullString)
        Let VecAux2 = Array(StrAux0, VecAux1(i, 5), VecAux1(i, 7), Empty, DblAux0, VecAux1(i, 0), VecAux1(i, 6))
        For j = k To UBound(VecAux1)
          'Let BooEquals = (Not VecAux1(j, 0) = VBA.vbNullString And Not VecAux1(i, 0) = VBA.vbNullString)
          Let StrAux0 = VBA.CStr(VecAux1(j, 4)) 'Content
          Let StrAux1 = VBA.CStr(VecAux1(i, 4)) 'Content
          Let BooEquals = (Not StrAux0 = VBA.vbNullString And Not StrAux1 = VBA.vbNullString)
          If BooEquals Then
            'Let BooEquals = (VBA.LCase(VBA.CStr(VecAux1(i, 4))) = VBA.LCase(VBA.CStr(VecAux1(j, 4))) And VBA.LCase(VBA.CStr(VecAux1(i, 6))) = VBA.LCase(VBA.CStr(VecAux1(j, 6))))
            Let StrAux0 = VBA.LCase(VBA.Replace(VBA.CStr(VecAux1(i, 4)), "~", VBA.vbNullString)) 'Content
            Let StrAux1 = VBA.LCase(VBA.Replace(VBA.CStr(VecAux1(j, 4)), "~", VBA.vbNullString)) 'Content
            Let StrAux2 = VBA.LCase(VBA.CStr(VecAux1(i, 6))) 'Unit
            Let StrAux3 = VBA.LCase(VBA.CStr(VecAux1(j, 6))) 'Unit
            Let BooEquals = ((StrAux0 = StrAux1) And (StrAux2 = StrAux3))
            If BooEquals Then
              Let BooEquals = (VBA.LCase(VBA.CStr(VecAux1(i, 7))) = VBA.LCase(VBA.CStr(VecAux1(j, 7)))) 'PRCi=PRCj
              Let DblAux0 = VBA.CDbl(VecAux1(j, 0)) 'PNS
              Let VecAux2(5) = VBA.CDbl(VecAux2(5)) + IIf(BooEquals Or BooAccurateAnyway, DblAux0, 0)
              Let DblAux0 = VBA.CDbl(VecAux1(j, 0)) * VBA.CDbl(VecAux1(j, 5)) * VBA.CDbl(VecAux1(j, 7)) * DblTC 'GRS=PNS*GRD*PRC*TC
              Let VecAux2(4) = VBA.CDbl(VecAux2(4)) + IIf(BooEquals Or BooAccurateAnyway, DblAux0, 0)
              Let DblAux0 = VBA.CDbl(VecAux1(j, 7)) 'PRC
              Let VecAux2(2) = VBA.CDbl(VecAux2(2)) + IIf(BooEquals Or BooAccurateAnyway, DblAux0, 0)
              If BooEquals Or BooAccurateAnyway Then Let VecAux1(j, 4) = VBA.vbNullString
              If BooAccurateAnyway Then Let l = l + 1
              Let BooAccurate = (BooAccurate Or BooEquals Or BooAccurateAnyway)
            End If
          End If
        Next j
        If Not VecAux1(i, 4) = VBA.vbNullString Then
          Let VecAux1(i, 4) = VBA.vbNullString
          If l > 1 Then Let VecAux2(2) = VBA.CDbl(VecAux2(2)) / l
          Let VecAux2(0) = IIf(BooAccurate, "Promed. ", VBA.vbNullString) & VBA.StrConv(VBA.CStr(VecAux2(0)), vbProperCase)
          Let VecAux2(1) = VBA.CDbl(VecAux2(4)) / (VBA.CDbl(VecAux2(2)) * VBA.CDbl(VecAux2(5)) * DblTC)
          Let VecAux3 = RES.ArrayAddAtLast(VecAux3, VecAux2)
        End If
        Let k = k + 1
      Next i
    End If
  
  Debug.Print "Success!!!"










'0000 works ok
'  Set REGEX = New AppResRegEx
''  Dim i As Long, j As Long, k As Long, l As Long, u As Long
'  Let i = 0
'  Do While i <= UBound(VecAux0)
'    If REGEX.isMineralComplex(VBA.CStr(VecAux0(i, 3))) Then
'      Let VecAux2 = VBA.Split(VBA.CStr(VecAux0(i, 4)), ";")
'      Let j = UBound(VecAux2) - LBound(VecAux2)
'      For k = 0 To j
'        Let VecAux3 = Array(VecAux0(i, 0), VecAux0(i, 1), VecAux0(i, 2), VecAux0(i, 3))
'        For l = 4 To 7
'          Let VecAux2 = VBA.Split(VBA.CStr(VecAux0(i, l)), ";")
'          Let u = UBound(VecAux3) + 1
'          ReDim Preserve VecAux3(u)
'          Let VecAux3(u) = VecAux2(k)
'          Let VecAux2 = Empty
'        Next l
'        Let VecAux1 = RES.ArrayAddAtLast(VecAux1, VecAux3)
'        Let VecAux3 = Empty
'      Next k
'    Else
'      Let VecAux1 = RES.ArrayAddAtLast(VecAux1, RES.ArrayDelIndex(VecAux0, i, True))
'    End If
'    Let i = i + 1
'  Loop
'
'    ' 2.  Calculate PNS
'    For i = 0 To UBound(VecAux1) 'PBH*H2O*LSS
'      Let VecAux1(i, 0) = ((VBA.CDbl(VecAux1(i, 0)) * (1 - (VBA.CDbl(VecAux1(i, 1)) / 100))) * (1 - (VBA.CDbl(VecAux1(i, 2)) / 100))) '* (VBA.CDbl(VecAux1(i, 5))) * (VBA.CDbl(VecAux1(i, 7)))
'    Next i
'    Let VecAux0 = Empty
'
'
'    ' 2.  Filter by Content, Unit & [Price]
'    Dim BooAccurateAnyway As Boolean, BooEquals As Boolean, BooAccurate As Boolean
'    Dim DblAux0 As Double
'    Dim DblTC As Double
'    Dim StrAux0 As String, StrAux1 As String, StrAux2 As String, StrAux3 As String
'    Let BooAccurateAnyway = True
'    Let DblTC = 6.96
'    Let k = 1
'    ReDim VecAux3(0 To 0, 0 To 4)
'    For i = 0 To UBound(VecAux1)
'      Let BooAccurate = False
'      Let l = 1
'      Let DblAux0 = VBA.CDbl(VecAux1(i, 0)) * VBA.CDbl(VecAux1(i, 5)) * VBA.CDbl(VecAux1(i, 7)) * DblTC 'PNS*GRD*PRC
'                   'Array(Name,          Grade,         Price,         Foreign, Local, PNS, Unit)
'      Let VecAux2 = Array(VecAux1(i, 4), VecAux1(i, 5), VecAux1(i, 7), Empty, DblAux0, VecAux1(i, 0), VecAux1(i, 6))
'      For j = k To UBound(VecAux1)
'        'Let BooEquals = (Not VecAux1(j, 0) = VBA.vbNullString And Not VecAux1(i, 0) = VBA.vbNullString)
'        Let StrAux0 = VBA.CStr(VecAux1(j, 4))
'        Let StrAux1 = VBA.CStr(VecAux1(i, 4))
'        Let BooEquals = (Not StrAux0 = VBA.vbNullString And Not StrAux1 = VBA.vbNullString)
'        If BooEquals Then
'          'Let BooEquals = (VBA.LCase(VBA.CStr(VecAux1(i, 4))) = VBA.LCase(VBA.CStr(VecAux1(j, 4))) And VBA.LCase(VBA.CStr(VecAux1(i, 6))) = VBA.LCase(VBA.CStr(VecAux1(j, 6))))
'          Let StrAux0 = VBA.LCase(VBA.CStr(VecAux1(i, 4)))
'          Let StrAux1 = VBA.LCase(VBA.CStr(VecAux1(j, 4)))
'          Let StrAux2 = VBA.LCase(VBA.CStr(VecAux1(i, 6)))
'          Let StrAux3 = VBA.LCase(VBA.CStr(VecAux1(j, 6)))
'          Let BooEquals = ((StrAux0 = StrAux1) And (StrAux2 = StrAux3))
'          If BooEquals Then
'            Let BooEquals = (VBA.LCase(VBA.CStr(VecAux1(i, 7))) = VBA.LCase(VBA.CStr(VecAux1(j, 7)))) 'PRCi=PRCj
'            Let DblAux0 = VBA.CDbl(VecAux1(j, 0)) 'PNS
'            Let VecAux2(5) = VBA.CDbl(VecAux2(5)) + IIf(BooEquals Or BooAccurateAnyway, DblAux0, 0)
'            Let DblAux0 = VBA.CDbl(VecAux1(j, 0)) * VBA.CDbl(VecAux1(j, 5)) * VBA.CDbl(VecAux1(j, 7)) * DblTC 'GRS=PNS*GRD*PRC*TC
'            Let VecAux2(4) = VBA.CDbl(VecAux2(4)) + IIf(BooEquals Or BooAccurateAnyway, DblAux0, 0)
'            Let DblAux0 = VBA.CDbl(VecAux1(j, 7)) 'PRC
'            Let VecAux2(2) = VBA.CDbl(VecAux2(2)) + IIf(BooEquals Or BooAccurateAnyway, DblAux0, 0)
'            If BooEquals Or BooAccurateAnyway Then Let VecAux1(j, 4) = VBA.vbNullString
'            If BooAccurateAnyway Then Let l = l + 1
'            Let BooAccurate = (BooAccurate Or BooEquals Or BooAccurateAnyway)
'          End If
'        End If
'      Next j
'      If Not VecAux1(i, 4) = VBA.vbNullString Then
'        Let VecAux1(i, 4) = VBA.vbNullString
'        If l > 1 Then Let VecAux2(2) = VBA.CDbl(VecAux2(2)) / l
'        Let VecAux2(0) = IIf(BooAccurate, "Promed. ", VBA.vbNullString) & VBA.StrConv(VBA.CStr(VecAux2(0)), vbProperCase)
'        Let VecAux2(1) = VBA.CDbl(VecAux2(4)) / (VBA.CDbl(VecAux2(2)) * VBA.CDbl(VecAux2(5)) * DblTC)
'        Let VecAux3 = RES.ArrayAddAtLast(VecAux3, VecAux2)
'      End If
'      Let k = k + 1
'    Next i
'
'  Debug.Print "Success!!!"
'000000 works ok




'  Let i = 0
'  Let u = 0
'  Do While i <= UBound(VecAux0)
'    If REGEX.isMineralComplex(VBA.CStr(VecAux0(i, 3))) Then
'      Let VecAux2 = VBA.Split(VBA.CStr(VecAux0(i, 4)), ";")
'      Let w = UBound(VecAux2) - LBound(VecAux2)
'      For u = 0 To w
'        Let VecAux3 = Array(VecAux0(i, 0), VecAux0(i, 1), VecAux0(i, 2), VecAux0(i, 3))
'        For l = 4 To 7
'          Let VecAux2 = VBA.Split(VBA.CStr(VecAux0(i, l)), ";")
'          Let v = UBound(VecAux3) + 1
'          ReDim Preserve VecAux3(v)
'          Let VecAux3(v) = VecAux2(u)
'          Let VecAux2 = Empty
'        Next l
'        Let VecAux1 = RES.ArrayAddAtLast(VecAux1, VecAux3)
'        Let VecAux3 = Empty
'      Next u
'    Else
'      Let VecAux1 = RES.ArrayAddAtLast(VecAux1, RES.ArrayDelIndex(VecAux0, i, True))
'    End If
'    Let i = i + 1
'  Loop
'  Debug.Print "Success!!!"



'  Let i = 0
'  Let u = 0
'  ReDim VecAux1(0 To j, 0 To UBound(VecAux0, 2))
''  Exit Sub
'  Do While u <= j 'UBound(VecAux0)
'  'Exit Do
'    'Let VecAux1 = RES.ArrayConcat(VecAux1, RES.ArrayDelIndex(VecAux0, i))
'    For k = 0 To UBound(VecAux0, 2) '-2v
'      Let VecAux1(u, k) = VecAux0(i, k)
'    Next k '-2v
''    Let VecAux1(u, 0) = VecAux0(i, 0)'-1v
''    Let VecAux1(u, 1) = VecAux0(i, 1)
''    Let VecAux1(u, 2) = VecAux0(i, 2)
''    Let VecAux1(u, 3) = VecAux0(i, 3)
''    Let VecAux1(u, 4) = VecAux0(i, 4)
''    Let VecAux1(u, 5) = VecAux0(i, 5)
''    Let VecAux1(u, 6) = VecAux0(i, 6)
''    Let VecAux1(u, 7) = VecAux0(i, 7)'-1v
'    If REGEX.isMineralComplex(VBA.CStr(VecAux0(i, 3))) Then
'      For l = 4 To 7
'        Let VecAux2 = VBA.Split(VBA.CStr(VecAux0(i, l)), ";")
'        For k = 0 To UBound(VecAux2)
'          Let VecAux1(u + k, 0) = VecAux0(i, 0)
'          Let VecAux1(u + k, 1) = VecAux0(i, 1)
'          Let VecAux1(u + k, 2) = VecAux0(i, 2)
'          Let VecAux1(u + k, 3) = VecAux0(i, 3)
'          Let VecAux1(u + k, l) = VecAux2(k)
'        Next k
'        Let VecAux2 = Empty
'      Next l
'      Let u = u + (k - 1)
'    End If
'    Let u = u + 1
'    Let i = i + 1
'  Loop
'  Debug.Print "Success!!!"
  
End Sub








Sub Objects_Testsimple()
'  Dim a As Worksheet, b As AppExcliq
'
'  Set a = excliqlites
'  Set b = New AppExcliq
'  If a Is excliqlites Then Debug.Print "Hi: ", a Is excliqlites
'  If Not b Is excliqlites Then Debug.Print "Hi: ", Not b Is excliqlites
'
'  Set a = Nothing
'  Set b = Nothing
  Dim a As Variant
  'Dim a(0) As Variant
  'Let a(0) = 1
  'ReDim Preserve a(1)
  'Let a(1) = 2
  'Let a = Array("TOTAL PESO PAGABLE", Empty, Empty, Empty, 4) '[{"3A",,;"4A","4B","4C"}]
  'Let a = [{"DATOS MINERAL Y PESO BRUTO HÚMEDO",empty,empty,empty,empty;"Muestra", "Ingreso", "Contenidos", "Tipo", "PBH [T]"}]
  Let a = Empty
End Sub








Sub ModelAppfilter_Testsimple()
  Dim RngA As Range, RngB As Range, RngCriteria As Range
  Dim VecData As Variant
  
  Set RngA = Hoja6.ListObjects("eqlfilteretst_a").Range.CurrentRegion
  Debug.Print RngA.Address(False, False)
  Set RngCriteria = Hoja6.Range("A1").CurrentRegion
  Debug.Print RngCriteria.Address(False, False)
  
  Call RngA.AdvancedFilter(xlFilterInPlace, RngCriteria)
  Let VecData = RngA.CurrentRegion
  
  Set RngA = Nothing
  Set RngCriteria = Nothing
  Let VecData = Empty
End Sub

Sub ModelApp_Test()
  Dim EHGLOBAL As AppErrorHandler
  Dim MDL As ModelExcliqliteDatasheet
  
  Dim LO As ListObject
  Dim RngBox As Range
  Dim VecAux As Variant
  Dim i As Long
  
  On Error GoTo EH
  Set EHGLOBAL = New AppErrorHandler
  Set MDL = New ModelExcliqliteDatasheet
  Set MDL.ErrorHandler = EHGLOBAL
  
  'SET
  'Call MDL.MSet(eqlMdlTblDatatest, Array("Bolis", "Bs"))
  'Call MDL.MSet(eqlMdlTblDatatest, [{"Bolis","Bs";"Lunas","Lu"}])
  'Call MDL.MSet(eqlMdlTblDatatest, Array("Sol", "Soles", "S/.", 0, 0))
  'Call MDL.MSet(eqlMdlTblDatatest, MLngRows:=1)
  
  'GET
  'Set RngBox = MDL.MGet(eqlMdlTblCurrencies, eqlMdlRange)
  'Set LO = MDL.MGet(eqlMdlTblCurrencies, eqlMdlListObject)
  'Let VecAux = MDL.MGet(eqlMdlTblCurrencies, eqlMdlArray, "symbol")
  'Let VecAux = MDL.MGet(eqlMdlTblCurrencies, eqlMdlArray, Empty, 1)
  'Let VecAux = MDL.MGet(eqlMdlTblCurrencies, eqlMdlArray, "currency, symbol", 3)
  'Let VecAux = MDL.MGet(eqlMdlTblCurrencies, eqlMdlArray, "currency, symbol", MStrWhere:="symbol=$UY")
  'Let VecAux = MDL.MGet(eqlMdlTblCurrencies, eqlMdlArray, "currency", MStrWhere:="symbol=$")
  
  'Dim straux0 As String, StrCurLocal As String, StrCurForeign As String
  'Let straux0 = VBA.CStr(MDL.MGet(eqlMdlTblCurrencies, eqlMdlArray, "currencies", MStrWhere:="currency_main=1")(0, 0))
  'Let StrCurLocal = VBA.CStr(MDL.MGet(eqlMdlTblCurrencies, eqlMdlArray, "symbol", MStrWhere:="currency_main=1")(0, 0))
  'Let StrCurForeign = VBA.CStr(MDL.MGet(eqlMdlTblCurrencies, eqlMdlArray, "symbol", MStrWhere:="currency_foreign=1")(0, 0))
  'Debug.Print MDL.Formulas
  'Let MDL.Formulas = True
  'Debug.Print MDL.Formulas
  'Let MDL.Formulas = False
  'Debug.Print MDL.Formulas
  'Debug.Print MDL.Tables
  'Let MDL.Tables = 0
  'Debug.Print MDL.Tables
  'Let MDL.Tables = 1
  'Debug.Print MDL.Tables
  
  'Set RngBox = MDL.MGet(eqlMdlTblChemicalunits, eqlMdlRange)
'  Let VecAux = Array("%", "ppm")
'  For i = 0 To UBound(VecAux)
'    If Not VBA.IsArray(MDL.MGet(eqlMdlTblChemicalunits, eqlMdlArray, "Unidad", MStrWhere:="Unidad=" & VBA.CStr(VecAux(i)))) Then
'      Debug.Print "Unidad: " & VBA.CStr(VecAux(i)) & " no existe"
'    Else
'      Debug.Print "Unidad: " & VBA.CStr(VecAux(i)) & " existe"
'    End If
'  Next i
'  Set RngBox = Nothing
  
  Dim VecAux1 As Variant, VecAux0 As Variant, BVecNames As Variant
  Dim StrAux0 As String
  Let BVecNames = Array("Cu", "W", "Ir", "Op")
  ReDim VecAux1(0 To 3)
  For i = 0 To UBound(VecAux1)
    Let StrAux0 = VBA.LCase(VBA.CStr(BVecNames(i)))
    Let VecAux0 = MDL.MGet(eqlMdlTblChemicalelements, eqlMdlArray, "Elemento", MStrWhere:="Símbolo=" & StrAux0)
    If VBA.IsArray(VecAux0) Then
      Let VecAux1(i) = VecAux0(0, 0)
    ElseIf (StrAux0 Like "b#" Or StrAux0 Like "bx#") Then
      Let VecAux1(i) = BVecNames(i)
    Else
      Debug.Print "¡Nombre de elemento: " & VBA.CStr(BVecNames(i)) & ", no existe!"
    End If
  Next i
  Let VecAux0 = Empty
  Let VecAux1 = Empty
  Let BVecNames = Empty

  
  'UPD
  'Call MDL.MUpd(eqlMdlTblDatatest, Array("Bolis", "Bs"), "A, C", 1)
  'Call MDL.MUpd(eqlMdlTblDatatest, Array("Boliviano", "BOB"), "A, C", MStrWhere:="B=Bolivianos")
  
  'DEL
  'Call MDL.MDel(eqlMdlTblDatatest, 6)
  
EH:
  Let VecAux = Empty
  Set RngBox = Nothing
  Set LO = Nothing
  Set MDL = Nothing
  Call EHGLOBAL.ErrorHandlerDisplay("Me")
End Sub


Sub Separatorsapp_Testsimple()
  Dim StrI As String, Sepa As String, StrIx As String, Deci As String
  Dim i As Double
  
  Let Deci = Application.International(xlDecimalSeparator)
  Let Sepa = Application.International(xlListSeparator)
   

  
  Let i = 0.25
  If Deci = "," Then
    Let StrIx = VBA.Format(i, "0.00")
  Else
    Let StrIx = VBA.CStr(i)
  End If
  Let StrI = "=SUMA(" & StrIx & Sepa & StrIx & ")"
  
  Debug.Print StrI
  'Let ActiveCell.FormulaLocal = StrI

End Sub




Sub ListObjectsToArray_Testsimple()
  Dim RngBox As Range, RngFin As Range
  Dim VarX As Variant, VarY As Variant
  Dim i As Long
  
  Let VarY = Array("currency", "symbol", "currency_foreign")
  
  'Let VarX = excliqlitedatasheet.ListObjects("excliqlitecurrenciescon").ListColumns(1).Range.address
  Set RngFin = excliqlitedatasheet.ListObjects("excliqlitecurrenciescon").ListColumns(VBA.CStr(VarY(0))).DataBodyRange
  For i = 1 To UBound(VarY)
    Set RngBox = excliqlitedatasheet.ListObjects("excliqlitecurrenciescon").ListColumns(VBA.CStr(VarY(i))).DataBodyRange
    Set RngFin = Application.Union(RngBox, RngFin)
  Next i
  Let VarX = RngFin.Value
  
  Let VarX = Empty
  Set RngBox = Nothing
  Set RngFin = Nothing
End Sub






Sub Balmet_current_Test()
  Dim EHGLOBAL As AppErrorHandler
  Dim Bmt As AppExcliqBalance_current
  
  On Error GoTo EH
  'Data entry in BVarRequest: (DblFeed, DblFeedVol, VecGrad, VecGradCx, VecName, VecUnit, BytMethod, Booleans, TypeBalmet, RngBox)
  
  Dim VecAux0 As Variant, VecName As Variant, VecUnit As Variant, VecBools As Variant, VecGrade As Variant, VecGradeCx As Variant
  Dim VecProducts As Variant
  'BooAll, BooPercents, BooUnits, BooFines, BooGrams, BooOT, BooRatio, BooHeads
  
  Set EHGLOBAL = New AppErrorHandler
  Set Bmt = New AppExcliqBalance_current
  Set Bmt.ICoreController_ErrorHandler = EHGLOBAL
  
  ReDim VecGrade(0 To 3, 0 To 1) '3P
  Let VecGrade(0, 0) = 60
  Let VecGrade(1, 0) = 2.57
  Let VecGrade(2, 0) = 0.41
  Let VecGrade(3, 0) = 3
  Let VecGrade(0, 1) = 18.08
  Let VecGrade(1, 1) = 50
  Let VecGrade(2, 1) = 0.84
  Let VecGrade(3, 1) = 8
  
  ReDim VecGradeCx(0 To 3, 0 To 0) '2P
  Let VecGradeCx(0, 0) = 359.69
  Let VecGradeCx(1, 0) = 5.67
  Let VecGradeCx(2, 0) = 0.9
  Let VecGradeCx(3, 0) = 26.34

'  ReDim VecGrade(0 To 2, 0 To 0) '2P
'  Let VecGrade(0, 0) = 8#
'  Let VecGrade(1, 0) = 1.2
'  Let VecGrade(2, 0) = 3#
'  ReDim VecGradeCx(0 To 2, 0 To 0) '2P
'  Let VecGradeCx(0, 0) = 45#
'  Let VecGradeCx(1, 0) = 0.6
'  Let VecGradeCx(2, 0) = 2.5

  Let VecBools = Array(True, True, True, True, True, True, True, True)
  Let VecName = Array("Zn", "Pb", "Ag") '"Pb")
  Let VecUnit = Array("%", "%", "g/L")
'  Let VecName = Array("Zn", "Ag") '"Pb")
'  Let VecUnit = Array("%", "g/L")
  
  Let VecAux0 = Array(400, 500, VecGrade, VecGradeCx, VecName, VecUnit, 2, VecBools, 2, Selection) ' Nothing)
  Call Bmt.ICoreController_GetSolution("balmet", VecAux0, VecAux0)
EH:
  Set Bmt = Nothing
  Call EHGLOBAL.ErrorHandlerDisplay("Me")
  Set EHGLOBAL = Nothing
End Sub




Sub Min_Test()
  Dim arr As Variant
  Dim i As Long
  
  Let arr = [{1,-33,4,6,-4,0}]
  Let i = Application.WorksheetFunction.MIN(arr)
  Debug.Print i
  
  Let arr = Empty
End Sub




Sub ArrayMethods_Test()
  Dim RES As AppResources
  Dim Vec0 As Variant, Vec1 As Variant, Vec2 As Variant
  Dim BooError As Boolean
  
  Set RES = New AppResources

'  'Slice 1D array
'  Let Vec1 = [{"A1","B1","C1"}]
'  Let Vec2 = RES.resArraySlice(Vec1, 1, BooError)
  
  ''Slice 2D array by rows
  'Let Vec1 = [{"3A","3B","3C";"4A","4B","4C"}]
  'Let Vec2 = RES.resArraySlice(Vec1, 1, BooError)
  '
  ''Slice 2D array by cols
  'Let Vec1 = [{"3A","3B","3C";"4A","4B","4C"}]
  'Let Vec2 = RES.resArraySlice(Vec1, 1, BooError, , True)
  
  ''Slice 2D array by rows, return whole src
  'Let Vec1 = [{"3A","3B","3C";"4A","4B","4C"}]
  'Let Vec2 = RES.resArraySlice(Vec1, 1, BooError, True)
  
  ''Slice 2D array by cols, return whole src
  'Let Vec1 = [{"3A","3B","3C";"4A","4B","4C"}]
  'Let Vec2 = RES.resArraySlice(Vec1, 1, BooError, True, True)
  '
  'Debug.Print "Success!"
  ''-------------
  
  '''AddLast Str or Num at no array
  'Let Vec0 = "XA"
  ''Let Vec1 = [{"A1","B1","C1"}]
  'Let Vec2 = RES.resArrayAddAtLast(Vec1, Vec0, BooError)
  
  ''AddLast 1D Vec at no array
  ''Let Vec0 = "XA"
  'Let Vec1 = [{"A1","B1","C1"}]
  'Let Vec2 = RES.resArrayAddAtLast(Vec0, Vec1, BooError)
  
  ''AddLast 2D Vec at no array
  ''Let Vec0 = "XA"
  'Let Vec1 = [{"3A","3B","3C";"4A","4B","4C"}]
  'Let Vec2 = RES.resArrayAddAtLast(Vec0, Vec1, BooError)
  
  'Debug.Print "Success!"
  '-------------

  ''AddLast Str or Num at no array in opt 1D
  ''Let Vec0 = "XA"
  'Let Vec1 = [{"A1","B1","C1"}]
  'Let Vec2 = RES.resArrayAddAtLast(Vec1, Vec0, BooError)
  
  ''AddLast 1D Vec at no array in opt 2D byrows
  'Let Vec0 = "XA"
  'Let Vec1 = [{"A1","B1","C1"}]
  'Let Vec2 = RES.resArrayAddAtLast(Vec1, Vec0, BooError)
  
  ''AddLast constant at 2D Vec array in opt 2D byrows
  'Let Vec0 = "XA"
  'Let Vec1 = [{"A1","B1","C1";"A2","B2","C2"}]
  'Let Vec2 = RES.resArrayAddAtLast(Vec1, Vec0, BooError)
  
  ''AddLast constant at 2D Vec array in opt 2D bycols
  'Let Vec0 = "XA"
  'Let Vec1 = [{"A1","B1","C1";"A2","B2","C2"}]
  'Let Vec2 = RES.resArrayAddAtLast(Vec1, Vec0, BooError, True)

  ''AddLast constant at 2D Vec array in opt 2D bycols
  'Let Vec0 = [{"A1","B1","C1";"A2","B2","C2"}]
  'Let Vec1 = [{"A3","B3","C3";"A4","B4","C4"}]
  'Let Vec2 = RES.resArrayAddAtLast(Vec1, Vec0, BooError, True)
  '
  'Debug.Print "Success!"
  '-------------
  
  
  
  ''AddFirst Str or Num at no array
  'Let Vec0 = "XA"
  ''Let Vec1 = [{"A1","B1","C1"}]
  'Let Vec2 = RES.resArrayAddAtFirst(Vec1, Vec0, BooError)
  
  ''AddFirst 1D Vec at no array
  ''Let Vec0 = "XA"
  'Let Vec1 = [{"A1","B1","C1"}]
  'Let Vec2 = RES.resArrayAddAtFirst(Vec0, Vec1, BooError)
  
  ''AddFirst 2D Vec at no array
  ''Let Vec0 = "XA"
  'Let Vec1 = [{"3A","3B","3C";"4A","4B","4C"}]
  'Let Vec2 = RES.resArrayAddAtFirst(Vec0, Vec1, BooError)
  
  'Debug.Print "Success!"
  '-------------

  ''AddFirst Str or Num at no array in opt 1D
  ''Let Vec0 = "XA"
  'Let Vec1 = [{"A1","B1","C1"}]
  'Let Vec2 = RES.resArrayAddAtFirst(Vec1, Vec0, BooError)
  
  ''AddFirst 1D Vec at no array in opt 2D byrows
  'Let Vec0 = "XA"
  'Let Vec1 = [{"A1","B1","C1"}]
  'Let Vec2 = RES.resArrayAddAtFirst(Vec1, Vec0, BooError)
  
  ''AddFirst constant at 2D Vec array in opt 2D byrows
  'Let Vec0 = "XA"
  'Let Vec1 = [{"A1","B1","C1";"A2","B2","C2"}]
  'Let Vec2 = RES.resArrayAddAtFirst(Vec1, Vec0, BooError)
  
  ''AddFirst constant at 2D Vec array in opt 2D bycols
  'Let Vec0 = "XA"
  'Let Vec1 = [{"A1","B1","C1";"A2","B2","C2"}]
  'Let Vec2 = RES.resArrayAddAtFirst(Vec1, Vec0, BooError, True)

  'Debug.Print "Success!"
  '-------------
  
  ''Concat 2D By rows
  'Let Vec0 = [{"1A","1B","1C";"2A","2B","2C"}]
  'Let Vec1 = [{"3A","3B","3C";"4A","4B","4C"}]
  'Let Vec2 = RES.resArrayConcat(Vec0, Vec1, BooError, 0)
  
  ''Concat 2D By rows at n position and Add vector 1D
  'Let Vec0 = [{"1A","1B","1C";"2A","2B","2C"}]
  'Let Vec1 = [{"3A","3B","3C","4A","4B","4C"}]
  'Let Vec2 = RES.resArrayConcat(Vec0, Vec1, BooError, 2)
  
  'Debug.Print "Success!"
  '-------------
  
  ''Shift 1D
  'Let Vec0 = [{"1A","1B","1C","2A","2B","3B"}] '[{"1A","1B","1C";"2A","2B","3B"}]
  'Let Vec1 = RES.resArrayShift(Vec0, BooError, True)
  
  ''Shift 2D By Cols
  'Let Vec0 = [{"1A","1B","1C";"2A","2B","3B"}]
  'Let Vec1 = RES.resArrayShift(Vec0, BooError, True)
  
  ''Shift 2D By Rows
  'Let Vec0 = [{"1A","1B","1C";"2A","2B","2C";"3A","3B","3C"}]
  'Let Vec1 = RES.resArrayShift(Vec0, BooError, False, True)
  
  'Debug.Print "Success!"
  '-------------
  
  ''Pop 1D
  'Let Vec0 = [{"1A","1B","1C","2A","2B","3B"}] '[{"1A","1B","1C";"2A","2B","3B"}]
  'Let Vec1 = RES.resArrayPop(Vec0, BooError, True, eqlRes1D)
  
  ''Pop 2D By Cols
  'Let Vec0 = [{"1A","1B","1C";"2A","2B","3C"}]
  'Let Vec1 = RES.resArrayPop(Vec0, BooError, True, eqlRes2D, True)
  
  ''Pop 2D By Rows
  'Let Vec0 = [{"1A","1B","1C";"2A","2B","2C";"3A","3B","3C"}]
  'Let Vec1 = RES.resArrayPop(Vec0, BooError)
  
  'Debug.Print "Success!"
  '-------------
  
  ''Del x Index 1D
  'Let Vec0 = [{"1A","1B","1C","2A","2B","3B"}] '[{"1A","1B","1C";"2A","2B","3B"}]
  'Let Vec1 = RES.resArrayDelIndex(Vec0, 2, BooError, True)
  
  ''Del x index 2D By Cols
  'Let Vec0 = [{"1A","1B","1C";"2A","2B","3C"}]
  'Let Vec1 = RES.resArrayDelIndex(Vec0, 2, BooError, True, True)
  
  ''Del x index 2D By Rows
  'Let Vec0 = [{"1A","1B","1C";"2A","2B","2C";"3A","3B","3C"}]
  'Let Vec1 = RES.resArrayDelIndex(Vec0, 2, BooError, True, eqlRes2D)
  
  'Debug.Print "Success!"
  '-------------
  
  ''Concat 1D
  'Let Vec0 = [{"1A","1B","1C"}]
  'Let Vec1 = [{"2A","2B","2C"}]
  'Let Vec2 = RES.resArrayConcat(Vec0, Vec1, BooError, 2, eqlRes1D)
  
  ''Concat 1D at 2D By cols
  'Let Vec0 = [{"1A","1B","1C";"2A","2B","2C"}]
  'Let Vec1 = [{"2A","2B","2C"}]
  'Let Vec2 = RES.resArrayConcat(Vec0, Vec1, BooError, 0, eqlRes2D, True)
  
  ''Concat 2D By cols
  'Let Vec0 = [{"1A","1B","1C";"2A","2B","2C"}]
  'Let Vec1 = [{"3A","3B","3C";"4A","4B","4C"}]
  'Let Vec2 = RES.resArrayConcat(Vec0, Vec1, BooError, 0, eqlRes2D, True)
  
  ''Concat 2D By cols at n position and Add vector 1D
  'Let Vec0 = [{"1A","1B","1C";"2A","2B","2C"}]
  'Let Vec1 = [{"3A","3B","3C","4A","4B","4C"}]
  'Let Vec2 = RES.resArrayConcat(Vec0, Vec1, BooError, 2, eqlRes2D, True)
  
  ''Concat 2D By cols at n position and Add vector 2D
  'Let Vec0 = [{"1A","1B","1C";"2A","2B","2C"}]
  'Let Vec1 = [{"3A","3B","3C";"4A","4B","4C"}]
  'Let Vec2 = RES.resArrayConcat(Vec0, Vec1, BooError, 2, eqlRes2D, True)
  
  
  ''Concat 1D at 2D By rows
  'Let Vec0 = [{"1A","1B","1C";"2A","2B","2C"}]
  'Let Vec1 = [{"3A","3B","3C"}]
  'Let Vec2 = RES.resArrayConcat(Vec0, Vec1, BooError, 0, eqlRes2D)
  
  ''Concat 2D By rows
  'Let Vec0 = [{"1A","1B","1C";"2A","2B","2C"}]
  'Let Vec1 = [{"3A","3B","3C";"4A","4B","4C"}]
  'Let Vec2 = RES.resArrayConcat(Vec0, Vec1, BooError, 0, eqlRes2D)
  
  ''Concat 2D By rows at n position and Add vector 1D
  'Let Vec0 = [{"1A","1B","1C";"2A","2B","2C"}]
  'Let Vec1 = [{"3A","3B","3C","4A","4B","4C"}]
  'Let Vec2 = RES.resArrayConcat(Vec0, Vec1, BooError, 2, eqlRes2D)
        
  Let Vec0 = Empty
  Let Vec1 = Empty
  Let Vec2 = Empty
  Set RES = Nothing
End Sub






Sub CountIfInArrays_Test()
  Dim arr As Variant
  
  Let arr = Array(100)
  Debug.Print Application.WorksheetFunction.Sum(arr)
  
  
  Let arr = Array("Na (1)", "Na (2)", "Na (3)", "Na (4)")
  
  'Debug.Print Application.WorksheetFunction.CountIf(arr, "na")
  
  Let arr = Empty
End Sub


Sub TextsComparison_Test()
  Dim Rng As Range, Rng1 As Range
  Dim Vec As Variant, Vec2 As Variant
  Dim Str As String, Str2 As String
  'Dim i As Double
  Dim i As Long, j As Long
  
  'Let Str = "Non Float Bi (1)"
  Let Str = "Na NF(1)"
  'Set Rng1 = excliqlitedatasheet.ListObjects("excliqlitechemicalelements").DataBodyRange
  Let Vec = excliqlitedatasheet.ListObjects("excliqlitechemicalelements").DataBodyRange.Value
  'Let Str2 = Application.WorksheetFunction.TextJoin("* ", True, Rng1)
  
'  Debug.Print Str2
'  Debug.Print Str
'  Debug.Print (Str Like Str2)
'  For Each Vec2 In Vec
'    'Debug.Print Rng.Value, "Like " & "*" & Str & "*", (Rng.Value Like "*" & Str & "*")
'    Let Str2 = "*" & VBA.CStr(Vec2) & "*"
'    Debug.Print Str2, "Like " & Str, (Str Like Str2)
'  Next Vec2
  
  For j = LBound(Vec, 2) To UBound(Vec, 2)
    For i = LBound(Vec) To UBound(Vec)
      Let Str2 = VBA.LCase(VBA.CStr(Vec(i, j)) & " *")
      If (VBA.LCase(Str) Like Str2) Then
        Debug.Print Str2, "Like " & Str, (VBA.LCase(Str) Like Str2)
        Exit For
      End If
    Next i
  Next j
  
  'Let i = Application.WorksheetFunction.CountIf(Rng, Str)
'  For Each Rng In Rng1
'    'Debug.Print Rng.Value, "Like " & "*" & Str & "*", (Rng.Value Like "*" & Str & "*")
'    Debug.Print "*" & Rng.Value & "*", "Like " & Str, (Str Like "*" & Rng.Value & "*")
'  Next Rng
  
  'Debug.Print i
  
  Set Rng = Nothing
  Set Rng1 = Nothing
End Sub



Sub ProductBounds_Test()
  Dim BalTe As AppExcliqBalance_current
  
  Set BalTe = New AppExcliqBalance_current
  
'  Debug.Print BalTe.BalmetAssetsGetProductsBounds(8, 16, eqlBalColumnsCx)
  
  Set BalTe = Nothing
  
End Sub




Sub ArrDimensions_Test()
  Dim arr As Variant
  
  'ReDim arr(0)
  
  Debug.Print ArrayDimension_Test(arr)
  
  Let arr = Empty
End Sub

Public Function ArrayConcat_Test( _
  ByVal RVarSrc As Variant, _
  ByVal RVarAdd As Variant, _
  ByRef RError As Boolean, _
  Optional RPositionAt As Long = 0, _
  Optional RDimension As EQLRES_ENU_DIMENSIONARRAY = eqlRes1D, _
  Optional RByCols As Boolean = False, _
  Optional RES As AppResources) As Variant
  
  Dim VecAux0 As Variant
  Dim BooOneDimension As Boolean
  Dim i As Long, j As Long, k As Long, l As Long, xRow As Long, xCol As Long
  
  On Error GoTo EH
  Let RError = True
  
  Select Case RDimension
    Case eqlRes1D
      Let xRow = RES.resArrayLength(RVarSrc) + RES.resArrayLength(RVarAdd)
      Let k = 0
      ReDim VecAux0(0)
      For i = LBound(RVarSrc) To UBound(RVarSrc)
        If i = RPositionAt Then
          For j = LBound(RVarAdd) To UBound(RVarAdd)
            Let VecAux0(k) = RVarAdd(j)
            Let k = k + 1
          Next j
        Else
          Let VecAux0(k) = RVarSrc(i)
        End If
        Let k = k + 1
      Next i
    
    Case eqlRes2D
      If RByCols Then
        Let i = RES.resArrayDimension(RVarAdd)
        Let BooOneDimension = (i = 1)
        If i = 1 Then
          Let xCol = ((UBound(RVarSrc, 2) - LBound(RVarSrc, 2)) + 1) '+ Me.resArrayLength(RVarSrc)
          Let xRow = (UBound(RVarSrc) - LBound(RVarSrc))  '+ Me.resArrayLength(RVarSrc)
        Else '2D
          Let xCol = ((UBound(RVarSrc, 2) - LBound(RVarSrc, 2)) + 1) + ((UBound(RVarAdd, 2) - LBound(RVarSrc, 2)) + 1)
          Let xRow = (UBound(RVarSrc) - LBound(RVarSrc))  '+ Me.resArrayLength(RVarSrc)
        End If
        
        Let k = 0
        Let l = 0
        ReDim VecAux0(0 To xRow, 0 To xCol)
        For j = LBound(RVarSrc, 2) To RPositionAt
          For i = LBound(RVarSrc) To UBound(RVarSrc)
            Let VecAux0(k, l) = RVarSrc(i, j)
            Let k = k + 1
          Next i
          Let l = l + 1
        Next j
        
        If BooOneDimension Then
          For i = LBound(RVarAdd) To UBound(RVarAdd)
            Let VecAux0(k, l) = RVarAdd(i)
            Let k = k + 1
          Next i
          Let l = l + 1
        Else
          For j = LBound(RVarAdd, 2) To UBound(RVarAdd, 2)
            For i = LBound(RVarAdd) To UBound(RVarAdd)
              Let VecAux0(k, l) = RVarAdd(i, j)
              Let k = k + 1
            Next i
            Let l = l + 1
          Next j
        End If
      
        For j = RPositionAt To UBound(RVarSrc, 2)
          For i = LBound(RVarSrc) To UBound(RVarSrc)
            Let VecAux0(k, l) = RVarSrc(i, j)
            Let k = k + 1
          Next i
          Let l = l + 1
        Next j
      
      Else 'By rows
        Let i = RES.resArrayDimension(RVarAdd)
        Let BooOneDimension = (i = 1)
        If i = 1 Then
          Let xCol = (UBound(RVarSrc, 2) - LBound(RVarSrc, 2)) '+ Me.resArrayLength(RVarSrc)
          Let xRow = (UBound(RVarSrc) - LBound(RVarSrc)) + 1 '+ Me.resArrayLength(RVarSrc)
        Else '2D
          Let xCol = (UBound(RVarSrc, 2) - LBound(RVarSrc, 2))
          Let xRow = ((UBound(RVarSrc) - LBound(RVarSrc)) + 1) + ((UBound(RVarAdd) - LBound(RVarSrc)) + 1)
        End If
        
        Let k = 0
        Let l = 0
        ReDim VecAux0(0 To xRow, 0 To xCol)
        For i = LBound(RVarSrc) To RPositionAt
          For j = LBound(RVarSrc, 2) To UBound(RVarSrc, 2)
            Let VecAux0(k, l) = RVarSrc(i, j)
            Let k = k + 1
          Next j
          Let l = l + 1
        Next i
        
        If BooOneDimension Then
          For j = LBound(RVarAdd) To UBound(RVarAdd)
            Let VecAux0(k, l) = RVarAdd(j)
            Let k = k + 1
          Next j
          Let l = l + 1
        Else
          For i = LBound(RVarAdd) To UBound(RVarAdd)
            For j = LBound(RVarAdd, 2) To UBound(RVarAdd, 2)
              Let VecAux0(k, l) = RVarAdd(i, j)
              Let k = k + 1
            Next j
            Let l = l + 1
          Next i
        End If
        
        For i = RPositionAt To UBound(RVarSrc)
          For j = LBound(RVarSrc, 2) To UBound(RVarSrc, 2)
            Let VecAux0(k, l) = RVarSrc(i, j)
            Let k = k + 1
          Next j
          Let l = l + 1
        Next i
      End If
    
    Case Else: GoTo EH
  End Select
  
  Let RError = False
  Let ArrayConcat_Test = VecAux0

EH:
  Let VecAux0 = Empty
  'Call RES_ErrorHandler(sFN, vbInformation)
  On Error GoTo -1
End Function

Public Function ArrayDimension_Test(ByVal RVarSrc As Variant) As Long
  Dim i As Long, j As Long, xLength As Long
  
  On Error GoTo EH
  
  'Assuming that Src is not an array
  Let xLength = -1
  If Not VBA.IsArray(RVarSrc) Then GoTo EH
  
  'Is is an array, calculate its dimensions
  Let i = 1
  Do While True
    Let j = (UBound(RVarSrc, i) - LBound(RVarSrc, i)) + 1
    Let xLength = i
    Let i = i + 1
  Loop

EH:
  Let ArrayDimension_Test = xLength
  'Call EHGLOBAL.ErrorHandlerClear
  On Error GoTo -1
  Call VBA.Err.Clear
End Function




Sub CVErr_Test()
  'Dim CVErrx As Variant
  'Let CVErrx = VBA.CVErr(xlErrNA)
  'Let CVErrx = CVErrx & " ?"
  With ActiveCell
    Call .Clear
    'Let .Value = "Whiiiii"
    'Let .Value = VBA.CVErr(xlErrNA) '& "?"'X
  End With
End Sub


Sub Cncinvoice_Test()
  ' Request
  ' 0-BooPage
  ' 1-BooPrintTwo
  ' 2-BooNewSheet
  ' 3-RngBox
'  Dim cnc As AppExcliqPurchaseConcentrate
'
'  Set cnc = New AppExcliqPurchaseConcentrate
'
'  Let cnc.SheetPrefix = "excliqlite"
'  Set cnc.SheetData = excliqlitedatasheet
'  Debug.Print cnc.PurchaseConcentrateGet(Array(0, 0, 0, Selection.Range("A1")))
'
'  Set cnc = Nothing

  Dim cnc As Controller
  
  Set cnc = New Controller
  
  'Set cnc.SheetData = excliqlitedatasheet
  Debug.Print cnc.Concentrateinvoice(Array(0, 0, 0, Selection.Range("A1")))
  
  Set cnc = Nothing
End Sub




Sub RngAlignmentsTest()
  Dim r As Range
  Dim h As XlHAlign
  
  Set r = Selection
  Let h = xlHAlignCenter
  Let r.HorizontalAlignment = h
  
  Set r = Nothing
End Sub


Sub FormatNumbersTest()
'  Debug.Print excliqlitedatasheet.ListObjects.Item(1).Name
'  Debug.Print excliqlitedatasheet.ListObjects.Item(1).DisplayName
  'Let excliqlitedatasheet.ListObjects("excliqliteformulascon").DisplayName = False
'  Let excliqlitedatasheet.ListObjects("excliqliteformulascon").Name.Visible = False
'  Dim n As Name
'  For Each n In excliqlitedatasheet.Names
'    Debug.Print n.Name
'  Next n
'  Set n = Nothing
  Dim Rng As Range

  Set Rng = Selection

  'Let rng.Value = VBA.Format(2, "0.00DM")
  Let Rng.NumberFormat = "0.00 ""DM"""
  'Let rng.Value = VBA.Format$(2 / 100, 2, vbFalse, vbTrue, vbFalse)


  Set Rng = Nothing
End Sub

Sub OreinvoiceTest()
  ' Request
  ' 0-StrPlace      5-StrProNames[]      10-BooNewSheet
  ' 1-StrDate       6-StrProType[]       11-RngBox
  ' 2-StrTC         7-BooAccurate
  ' 3-BooME         8-BooPage
  ' 4-BooVarious    9-BooPrintTwo
  Dim ot As AppExcliqPurchaseMineral
  Set ot = New AppExcliqPurchaseMineral
  Dim Vec As Variant, Vec1 As Variant, Vec2 As Variant
  Let Vec2 = Array(0, VBA.vbNullString)
  'Let Vec = Array("Potosí", VBA.Date(), 6.96, False, False, "Roccou", "Particular", False, False, False, False, Selection) 'simple
  'Let Vec = Array("Potosí", VBA.Date(), 6.96, False, True, "Roccou", "Particular", False, False, False, False, Selection) 'multiple
  Let Vec = Array("Potosí", VBA.Date(), 6.96, True, True, "Roccou", "Particular", True, False, False, False, Selection) 'average - simple
  Call ot.PurchaseMineralGet(Vec)
  
'  Debug.Print ot.ProjectionUDF(Selection)
  
  Set ot = Nothing
End Sub




Sub UniquesTest()
  Dim RES As New AppResources
  Dim v As Variant
  
  Let v = RES.resArrayUniques(Array("k", "zn", "ag", "Zn", "pb", "U", "ZN", "sb", "zN", "Pb", "Ag"))
  Set RES = Nothing
End Sub



Sub BaseArrays()
  Dim Vec As Variant
  
  ReDim Vec(1 To 1, 0 To 1)
  Let Vec = Empty
End Sub




Private Function Factorial(ByVal DblX As Double) As Double
  'Let Factorial = IIf(DblX = 1, 1, DblX * Factorial(DblX - 1))
  If DblX = 1 Then
    Let Factorial = 1
  Else
    Let Factorial = DblX * Factorial(DblX - 1)
  End If
End Function

Sub Factorial_Test()
  Dim i As Double
  Let i = Factorial(5)
End Sub


Function ArrayHasRepeatedElements(ByVal Vec As Variant) As Boolean
  Dim StrNeedle As String
  Dim i As Long, j As Long
  
  On Error GoTo EH
  If Not VBA.IsArray(Vec) Then GoTo EH
  Let i = 1
  Let j = 1
  Let StrNeedle = VBA.CStr(Vec(0))
  Let ArrayHasRepeatedElements = False
  Do While (ArrayHasRepeatedElements = False)
    Let ArrayHasRepeatedElements = (StrNeedle = VBA.CStr(Vec(i)))
    Let i = i + 1
    If i > UBound(Vec) Then
      Let StrNeedle = VBA.CStr(Vec(j))
      Let j = j + 1
      Let i = j
    End If
    If j - 1 = UBound(Vec) Then Exit Do
  Loop
EH:
End Function

Sub ArrayHasRepeatedElements_Test()
  Dim Boo As Boolean
  Let Boo = ArrayHasRepeatedElements(Array(2, 1, 7, 3, 1, 1))
End Sub


Sub ViewsTableStatic_Test()
  'Dim VIEWS As New ViewsApp
  Dim v As Variant
  Dim i As Long, j As Long
  
  Let v = Selection.Value
'  For j = LBound(v, 2) To UBound(v, 2)
'    For i = LBound(v) To UBound(v)
'      If VBA.IsEmpty(v(i, j)) Then Let v(i, j) = VBA.vbNullString
'    Next i
'  Next j
  
  'Call VIEWS.ViewsTableStaticSet(Selection, True, True, True)
  Range("Q20").Resize(UBound(v), UBound(v, 2)).Value2 = v
  
  Let v = Empty
  'Set VIEWS = Nothing
End Sub






Sub VecErrorTest2(Optional VErr As String = "Ho")
  Debug.Print VErr
End Sub


Sub VecErrorTest()
  Const VecErr As String = "Hi"
  Call VecErrorTest2(VecErr)
End Sub



Sub ProjectionsTest()
  ' Request
  ' 0-DblIo      5-VecW[]           9-StrDivise         11-BooVAN           16-Rng
  ' 1-DblVR      6-VecWName[]       10-StrTimeUnit      12-BooTIR
  ' 2-DblEgr     7-VecWUnitCtz[]                        13-BooGraph
  ' 3-DblT       8-VecCtz[]                             14-BooGraphSheet
  ' 4-Dbl%                                              15-BooNewSheet
  Dim pt As New AppExcliqProjections
  Dim Vec As Variant, Vec1 As Variant, Vec2 As Variant
  Let Vec2 = Array(0, VBA.vbNullString)
  Let Vec = Array(1000, 120, 2, 2, 51, "20;30", "Ag;Zn", "OT;%", "1;2;3/2;5;6;7", "USD", "Años", True, True, False, False, False, Selection)
  'Call pt.ProjectionGet(Vec, vec1, vec2)
  
'  Debug.Print pt.ProjectionUDF(Selection)
  
  Set pt = Nothing
End Sub





'------------- INTERNAL ITERFACE METHODS -------------'
Public Function BalmetGetProductsBounds(ByVal BLngSource As Long, Optional BStrWhat As String = "prod") As Long
  'Return the number of Products, Cols, Rows, TotalProducts and TotalComplexProducts of any Balmet
  Dim i As Long, LngProd As Long, LngCol As Long, LngRow As Long, LngTotProd As Long, LngTotCx As Long
  Dim BooFound As Boolean, BooTimeup As Boolean
  
  Const LngLASTBOUND As Long = 10
  Const sFN As String = "APPBALMET::GetProductsBounds"
  
  On Error GoTo EH
  Let BalmetGetProductsBounds = 0
  
  'Determine if the Source number is a valid quantity of elements to be a Balmet
  Let BooFound = False
  Let BooTimeup = False
  Let LngProd = 2
  Do While (Not BooFound And Not BooTimeup)
    Let LngCol = LngProd - 1
    Let LngRow = (LngProd + 1)
    Let LngTotProd = ((LngProd * (LngProd + 1)) - (LngProd + 1))
    Let LngTotCx = (LngProd * (LngProd + 1))
    
    Let LngProd = LngProd + 1
    Let BooFound = ((BLngSource = LngTotProd) Or (BLngSource = LngTotCx))
    Let BooTimeup = (LngProd = LngLASTBOUND)
    'Exit Do
  Loop
  If Not BooFound Then GoTo EH
  
  'Return
  Select Case BStrWhat
    Case "prod": Let BalmetGetProductsBounds = LngProd - 1
    Case "cols": Let BalmetGetProductsBounds = LngCol
    Case "rows": Let BalmetGetProductsBounds = LngRow
    Case "full": Let BalmetGetProductsBounds = LngTotProd
    Case "fulx": Let BalmetGetProductsBounds = LngTotCx
    Case Else: GoTo EH
  End Select

EH:
'  Call EHGlobal.ErrorHandlerRaise(sFN)
End Function

Public Function BalmetGetType(ByVal BLngSource As Long) As String
  'Return "normal" or "complex" type of any Balmet
  Dim LngProd As Long, LngTotProd As Long, LngTotCx As Long
  Dim BooFound As Boolean, BooTimeup As Boolean
  Dim StrType As String
  
  Const LngLASTBOUND As Long = 10
  Const sFN As String = "APPBALMET::GetType"
  
  On Error GoTo EH
  Let BalmetGetType = VBA.vbNullString
  
  'Determine if the Source number is a valid quantity of elements to be a Balmet
  Let BooFound = False
  Let BooTimeup = False
  Let StrType = VBA.vbNullString
  Let LngProd = 2
  Do While (Not BooFound And Not BooTimeup)
    Let LngTotProd = ((LngProd * (LngProd + 1)) - (LngProd + 1))
    Let LngTotCx = (LngProd * (LngProd + 1))
    
    Let LngProd = LngProd + 1
    Let BooFound = ((BLngSource = LngTotProd) Or (BLngSource = LngTotCx))
    Let BooTimeup = (LngProd = LngLASTBOUND)
    Let StrType = IIf((BLngSource = LngTotProd), "normal", IIf((BLngSource = LngTotCx), "complex", VBA.vbNullString))
    'Exit Do
  Loop
  If Not BooFound Then GoTo EH
  
  'Return
  Let BalmetGetType = StrType

EH:
'  Call EHGlobal.ErrorHandlerRaise(sFN)
End Function

Public Function BalmetGetProductsBoundsExist(ByVal BLngSource As Long) As Boolean
  'Return if source is a correct number of elements of any Balmet
  Dim i As Long
  Dim BooFound As Boolean
  
  Const sFN As String = "APPBALMET::GetProductsBoundsExist"
  
  On Error GoTo EH
  Let i = BalmetGetProductsBounds(BLngSource, "full")
  
  'Return
  Let BalmetGetProductsBoundsExist = Not (i = 0)

EH:
'  Call EHGlobal.ErrorHandlerRaise(sFN)
End Function
'------------- INTERNAL ITERFACE METHODS -------------'

Sub Test_balmetexists()
  Const a As Long = 12
  Debug.Print BalmetGetProductsBoundsExist(a), "type: " & BalmetGetType(a)
End Sub

Sub balmetproducts()
  Dim i As Long
  
  Debug.Print "P", "R", "G", "X"
  For i = 2 To 10
    Debug.Print i, i + 1, ((i * (i + 1)) - (i + 1)), (i * (i + 1))
  Next i
End Sub

Public Function resArrayHasValue( _
  ByVal AVector As Variant, _
  ByVal ANeedle As Variant, _
  Optional AStrOrNum As Boolean = False, _
  Optional ALCase As Boolean = False) As Boolean
  
  Dim i As Long
  Dim sAux As String

  On Error GoTo EH
  Let resArrayHasValue = False
  For i = 0 To UBound(AVector)
    If AStrOrNum Then Let sAux = IIf(ALCase, VBA.LCase(AVector(i)), AVector(i))
    If Not AStrOrNum Then Let sAux = VBA.CDbl(AVector(i))
    If ANeedle = sAux Then Let resArrayHasValue = True: Exit For
  Next i

EH:
  'Call EHGlobal.ErrorHandlerRaise("RES::InArray", vbInformation)
End Function

Public Function resArrayHasValueN( _
  ByVal AVector As Variant, _
  ByVal ANeedle As Variant, _
  Optional AStrOrNum As Boolean = False, _
  Optional ALCase As Boolean = False) As Long
  
  Dim i As Long
  Dim sAux As Variant

  On Error GoTo EH
  Let resArrayHasValueN = 0
  For i = 0 To UBound(AVector)
    If AStrOrNum Then Let sAux = IIf(ALCase, VBA.LCase(AVector(i)), AVector(i))
    If Not AStrOrNum Then Let sAux = VBA.CDbl(AVector(i))
    If ANeedle = sAux Then Let resArrayHasValueN = resArrayHasValueN + 1
  Next i

EH:
  'Call EHGlobal.ErrorHandlerRaise("RES::InArray", vbInformation)
End Function

Sub Test_resArrayHasValue_N()
  Dim i As Long
  Dim b As Boolean
  Dim a As Variant
  
  Let a = Array(1, 2, 3, "3")
  'Let i = resArrayHasValueN(a, 2)
  Let b = resArrayHasValue(a, 3)
  
  Let a = Empty
End Sub

Public Function resArrayGetDataColonSeparatedFromString( _
  ByVal StrVector As String, _
  ByRef StrCounter As Long, _
  Optional StrValueType As Boolean = False, _
  Optional StrAbs As Boolean = False, _
  Optional StrIncludeZeroes As Boolean = False, _
  Optional StrIncludeEmptyStrings As Boolean = False) As Variant
  
  Dim REGEX As AppResRegEx
  Dim VData As Variant, vAux As Variant
  Dim i As Long, j As Long, k As Long
  
  Let resArrayGetDataColonSeparatedFromString = Empty
  
  On Error GoTo EH
  Set REGEX = New AppResRegEx
  If REGEX.isEmptyStringReg(StrVector) Then GoTo EH
  
  Let i = 0
  Let j = 0
  Let k = 0
  If StrValueType Then 'Strings
    Let vAux = VBA.Split(StrVector, ";")
    ReDim VData(0)
    For i = 0 To UBound(vAux)
      If Not REGEX.isNumberReg(vAux(i)) Then
        If StrIncludeEmptyStrings Then
          ReDim Preserve VData(j)
          Let VData(j) = vAux(i)
          Let j = j + 1
        Else
          If REGEX.isStringReg(vAux(i)) Then
            ReDim Preserve VData(j)
            Let VData(j) = vAux(i)
            Let j = j + 1
          End If
        End If
      End If
    Next i
    If j = 0 Then GoTo EH
  Else 'Numbers
    'Let vAux = VBA.Split(StrVector, ";")
    If Not REGEX.isNumberEntireAndDecimalVectorColonSeparated(StrVector) Then GoTo EH
    Let vAux = VBA.Split(VBA.Replace(StrVector, ".", ","), ";")
    ReDim VData(0)
    For i = 0 To UBound(vAux)
      If REGEX.isNumberReg(vAux(i)) And Not REGEX.isEmptyStringReg(vAux(i)) Then
        If StrIncludeZeroes Then
          ReDim Preserve VData(j)
          Let VData(j) = IIf(StrAbs, VBA.Abs(VBA.CDbl(vAux(i))), VBA.CDbl(vAux(i)))
          Let j = j + 1
        Else
          If VBA.Abs(VBA.CDbl(vAux(i))) > 0 Then
            ReDim Preserve VData(j)
            Let VData(j) = IIf(StrAbs, VBA.Abs(VBA.CDbl(vAux(i))), VBA.CDbl(vAux(i)))
            Let j = j + 1
          End If
        End If
      End If
    Next i
    If j = 0 Then GoTo EH
  End If
  Set REGEX = Nothing
  
  Let StrCounter = j
  Let resArrayGetDataColonSeparatedFromString = VData
  
EH:
  Let VData = Empty
  Set REGEX = Nothing
End Function

Sub Test_resArrayGetDataColonSeparatedFromString()
  Dim VData As Variant
  Dim i As Long
  
  'Let vData = resArrayGetDataColonSeparatedFromString("a;1,2;c;0;5;4.52;-3", i, False, False)
  'Let vData = resArrayGetDataColonSeparatedFromString("a;1,2;c;0;;5;4.5;-3;%;g/T", i, True, StrIncludeEmptyStrings:=True)
  'Let vData = resArrayGetDataColonSeparatedFromString("%", i, True)
  Let VData = resArrayGetDataColonSeparatedFromString("a;1,2;c;0;;5;4.5;-3;%;g/T", i, True)
  
  Let VData = Empty
End Sub

Function resArrayGetDataFromRangesByRows( _
  ByVal Rng As Range, _
  ByRef RCounter As Long, _
  Optional RType As Boolean = False, _
  Optional RAbs As Boolean = False, _
  Optional RIncludeZeroes As Boolean = False, _
  Optional RIncludeEmptyStrings As Boolean = False) As Variant
  
  Dim REGEX As AppResRegEx
  Dim r As Range
  Dim VData As Variant
  Dim i As Long, j As Long, kRow As Long, kCol As Long
  
  Let resArrayGetDataFromRangesByRows = Empty
  
  On Error GoTo EH
  If Rng Is Nothing Then GoTo EH
  
  Set REGEX = New AppResRegEx
  Let i = 0
  If RType Then 'Strings
    If (Application.WorksheetFunction.CountA(Rng) - Application.WorksheetFunction.Count(Rng)) <= 0 Then GoTo EH
    For j = 1 To Rng.Areas.Count
      ReDim VData(0)
      For Each r In Rng.Areas(j)
        If Not REGEX.isNumberReg(r.Value) Then
          If RIncludeEmptyStrings Then
            ReDim Preserve VData(i)
            Let VData(i) = r.Value
            Let i = i + 1
          Else
            If REGEX.isStringReg(r.Value) Then
              ReDim Preserve VData(i)
              Let VData(i) = r.Value
              Let i = i + 1
            End If
          End If
        End If
      Next r
    Next j
  Else 'Numbers
    If Application.WorksheetFunction.Count(Rng) <= 0 Then GoTo EH
    Dim VecRng As Variant
    'Let VecRng = Rng.Value
    'Let VecRng = Empty
    ReDim VData(0)
    For j = 1 To Rng.Areas.Count
      If Rng.Areas(j).Cells.Count > 1 Then
        Let VecRng = Rng.Areas(j).Value
        For kCol = LBound(VecRng, 2) To UBound(VecRng, 2)
          For kRow = LBound(VecRng) To UBound(VecRng)
            If REGEX.isNumberReg(VecRng(kRow, kCol)) Then
              If RIncludeZeroes Then
                ReDim Preserve VData(i)
                If RAbs Then
                  Let VData(i) = VBA.Abs(VecRng(kRow, kCol))
                Else
                  Let VData(i) = VecRng(kRow, kCol)
                End If
                Let i = i + 1
              Else
                If VBA.Abs(VecRng(kRow, kCol)) > 0 Then
                  ReDim Preserve VData(i)
                  If RAbs Then
                    Let VData(i) = VBA.Abs(VecRng(kRow, kCol))
                  Else
                    Let VData(i) = VecRng(kRow, kCol)
                  End If
                  Let i = i + 1
                End If
              End If
            End If
          Next kRow
        Next kCol
      Else
        If REGEX.isNumberReg(Rng.Value) Then
          If RIncludeZeroes Then
            ReDim Preserve VData(i)
            Let VData(i) = IIf(RAbs, VBA.Abs(Rng.Areas(j).Cells(1, 1).Value), Rng.Areas(j).Cells(1, 1).Value)
            Let i = i + 1
          Else
            If VBA.Abs(Rng.Value) > 0 Then
              ReDim Preserve VData(i)
              Let VData(i) = IIf(RAbs, VBA.Abs(Rng.Areas(j).Cells(1, 1).Value), Rng.Areas(j).Cells(1, 1).Value)
              Let i = i + 1
            End If
          End If
        End If
      End If
    Next j

'    For j = 1 To Rng.Areas.Count
'      ReDim vData(0)
'      For Each r In Rng.Areas(j)
'        If REGEX.isNumberReg(r.Value) Then
'          If RIncludeZeroes Then
'            ReDim Preserve vData(i)
'            Let vData(i) = IIf(RAbs, VBA.Abs(r.Value), r.Value)
'            Let i = i + 1
'          Else
'            If VBA.Abs(r.Value) > 0 Then
'              ReDim Preserve vData(i)
'              Let vData(i) = IIf(RAbs, VBA.Abs(r.Value), r.Value)
'              Let i = i + 1
'            End If
'          End If
'        End If
'      Next r
'    Next j
  End If
  Set REGEX = Nothing
  
  Let RCounter = i
  Let resArrayGetDataFromRangesByRows = VData
  
EH:
  Let VData = Empty
  Set r = Nothing
  Set REGEX = Nothing
End Function

Sub Test_resArrayGetDataFromRangesByRows()
  Dim VData As Variant
  Dim i As Long
  
  Let VData = resArrayGetDataFromRangesByRows(Application.Selection, i, False, True)
  'Let vData = resArrayGetDataFromRangesByRows(Application.Selection, i, True, RIncludeEmptyStrings:=True)
  
  Let VData = Empty
End Sub

Sub countifarrays()
  Dim ab As String
  Dim va As Variant
  
  Let ab = "dm;st;xy"
  Let va = VBA.Split(ab, ";")
  Debug.Print Application.WorksheetFunction.CountIf(va, "dm") 'NO FUNCIONA CON ARREGLOS, SÓLO CON RANGOS
  Let va = Empty
End Sub

Sub printdots()
  Dim a(0, 0) As Variant
  
  Let a(0, 0) = "=" & Range("N1").Address & "/31.100005522"
  Selection.Value2 = a
End Sub

Public Function arrTester2(a As Range, b As Range) As Variant
  'Let arrTester2 = Array(1, 2, 3)
  Dim C(1, 0) As Variant
  
  Let C(0, 0) = a.Value2 + b.Value2
  Let C(1, 0) = a.Value2 * b.Value2
  
  Let arrTester2 = C
  
  Erase C
End Function

Public Function arrTester() As Variant
  'Let arrTester = Array(1, 2, 3)
  Dim r As Range
  Dim xProducts As Single, cPer As Single
  Dim a As String, b As String, C As String
  
  Let xProducts = 2
  Let cPer = 1
  Set r = Range("M1")
  With r.Range("C3")
    'Call .Range("C3").Resize(xProducts, 1).Select
    'Call .Range("C3").Offset(xProducts, 0).Select
    'Call .Range("C3").Offset(0, cPer + 1).Resize(xProducts + 1, xProducts - 1).Select
    'Let a = .Range("C3").Offset(xProducts, 0).address(True, True) & ";"
    'Let b = .Range("C3").Offset(0, cPer + 1).Resize(xProducts + 1, xProducts - 1).address(True, True) & ";"
    'Let c = "=EQL_BALMET(" & a & b & 1 & ";1)"
    'Let a = .Offset(xProducts, 0).address(True, True)
    'Let b = .Offset(0, cPer + 1).address(True, True)
    'Let c = "=arrTester2(" & a & ";" & b & ")"
    'Let .Range("C3").FormulaArray = c
    'Call .Range("C3").Resize(xProducts, 1).FillDown
    Call .Resize(xProducts, 1).Select
    'Let .Resize(xProducts, 1).FormulaArray = "=arrTester2(" & .Offset(xProducts, 0).address(True, True) & "," & .Offset(0, cPer + 1).address(True, True) & ")"
    'Let .Range("C3").Resize(xProducts, 1).FormulaArray = c
    Let .Resize(xProducts, 1).FormulaArray = _
        "=EQL_BALMET(" & _
           .Offset(xProducts, 0).Address(True, True) & "," & _
           .Offset(0, cPer + 1).Resize(xProducts + 1, xProducts - 1).Address(True, True) & "," & _
           1 & ",1)"
    'Let .Range("C3").Resize(xProducts, 1).Value2 = _
        "=+EQ_BALMET(" & _
           .Range("C3").Offset(xProducts, 0).address(True, True) & ";" & _
           .Range("C3").Offset(0, cPer + 1).Resize(xProducts + 1, xProducts - 1).address(True, True) & ";" & _
           1 & ";1)"
  End With
  Set r = Nothing
End Function

Sub arrPrint()
  Dim X As Range
  Dim arr(0, 2) As Variant
  Dim i As Long, j As Long
  
  For j = 0 To 2
    For i = 0 To 0
      Let arr(i, j) = "{=+arrtester()}"
    Next i
  Next j
  
  Set X = Range("M7")
  With X
    Let .Range("A1").Resize(1, 3).Value2 = arr
  End With
  
  Erase arr
  Set X = Nothing
End Sub

Sub testoffsets()
  Dim r As Range
  
  Set r = Selection
  With r
    Debug.Print r.Range("C3").Offset(3, 0).Resize(1, 1).Select
  End With
  
  Set r = Nothing
End Sub
Sub testinbalparts()
'  Dim x As AppExcliqBalance
'  Dim a() As Variant
'  Dim s() As Variant
'  Dim r As String
'
'  ReDim a(2, 0)
'  Let a(0, 0) = 1
'  Let a(1, 0) = 2
'  Let a(2, 0) = 3
'  Set x = New AppExcliqBalance
'  Debug.Print x.BalmetGetPercents(a, s, r)
'  Set x = Nothing
'  Debug.Print r
'
'  Erase a
'  Erase s
'  Set x = Nothing
End Sub
Function balancetestrancx(ByVal r1 As Range, ByVal r2 As Range, ByVal r3 As Range) As Variant
  Dim appbal As New AppExcliqBalance
  Dim bRES As Variant
  
  'Let bres = appbal.BalmetComplexUDF(r1, r2, r3, 1)
  'Let bres = appbal.BalmetUDFREC(r1, r2, 1)
'  Let bRES = appbal.BalmetUDFRECComplex(r1, r2, r3, 1)
  Let balancetestrancx = bRES
  
  Let bRES = Empty
  Set appbal = Nothing
End Function

Function balancetestran(ByVal r1 As Range, ByVal r2 As Range) As Variant
  Dim appbal As New AppExcliqBalance
  Dim bRES As Variant
  
  'Let bres = appbal.BalmetUDF(r1, r2)
  'Let bres = appbal.BalmetUDF(r1, r2, 1)
'  Let bRES = appbal.BalmetUDF(r1, r2, 2)
  Let balancetestran = bRES
  
  Let bRES = Empty
  Set appbal = Nothing
End Function

Sub balancetest()
  Dim appbal As New AppExcliqBalance
  Dim bRES As Variant
  
  'Let bres = appbal.BalmetUDF(Range("J8"), Selection)
  'Let bres = appbal.BalmetUDF(Range("C28"), Selection, 1, True)
  'Let bres = appbal.BalmetUDF(Range("D41"), Selection, 1)
  'Let bres = appbal.BalmetUDF(Range("D41"), Selection, 2)
'  Let bRES = appbal.BalmetUDF(Range("C28"), Selection, 2, True)

  'Let bres = appbal.BalmetUDFComplex(Range("D41"), Selection, Range("G41:G44"), 1)
  'Let bres = appbal.BalmetUDFComplex(Range("C28"), Selection, Range("G25:G28"), 1, True)
  'Let bres = appbal.BalmetUDFREC(Range("J2"), Selection)
  'Let bres = appbal.BalmetUDFREC(Range("D41"), Selection, 1)
  'Let bres = appbal.BalmetUDFREC(Range("C28"), Selection, 1, True)
  'Let bres = appbal.BalmetUDFRECCOMPLEX(Range("D41"), Selection, Range("G41:G44"), 1)
  'Let bres = appbal.BalmetUDFRECCOMPLEX(Range("C28"), Selection, Range("G25:G28"), 1, True)

  Let bRES = Empty
  Set appbal = Nothing
End Sub
Sub determinantshome()
  Dim a(0 To 2, 0 To 2) As Variant
  Dim i As Long, j As Long
  Dim d As Double
  
  For i = 0 To 2
    For j = 0 To 2
      Let a(i, j) = (VBA.Abs(VBA.Rnd(10)) * i) + 1
    Next j
  Next i
  Let d = Application.WorksheetFunction.MDeterm(a)
  Debug.Print d
  
  Erase a
End Sub
Sub rangeselements()
  Dim X As Range, Y As Range
  Dim i As Long, j As Long
  
  Set Y = Selection
  For i = 1 To Y.Columns.Count
    For j = 1 To Y.Rows.Count
      Debug.Print Y.Cells(j, i).Value
    Next j
  Next i
  
  Set X = Nothing
  Set Y = Nothing
End Sub

Sub sumabs()
  Dim a As Double, b As Double
  'Debug.Print Application.WorksheetFunction.Sum(Selection)
  Let a = Application.WorksheetFunction.SumIf( _
              Selection, ">0")
  Let b = Application.WorksheetFunction.SumIf( _
              Selection, "<0")
  Debug.Print a + VBA.Abs(b)
  
  Let a = Empty
  Let b = Empty
End Sub

Sub nums()
'Arrays
  Dim a As Variant
  Dim b As Double
  
'  ReDim a(1, 1)
'  Let a(0, 0) = 1
'  Let a(1, 0) = 2
'  Let a(1, 1) = 0
'  Let a(1, 1) = 1
  Let a = Array(1, 2, 3, 4)
  'Let b = Application.WorksheetFunction.CountIf(a, ">0")'do not works
  Let b = Application.WorksheetFunction.Sum(a)
  Debug.Print b
  
  Erase a

''Ranges
'  'Debug.Print VBA.IsNumeric(Selection.Value)
'  'Debug.Print VBA.IsEmpty(Selection.Value)
'  Dim r As Variant
'  Dim i As Long, j As Long, k As Long
'
'  Let r = Selection.Value
'  'Let k = Selection.Rows.Count - 1
'' Print the array values
'  Debug.Print "i", "j", "Value"
'  For i = LBound(r) To UBound(r)
'    For j = LBound(r, 2) To UBound(r, 2)
'      Debug.Print i, j, r(i, j)
'    Next j
'  Next i
'
'  Debug.Print "i", "j", "Value" 'OK
'  For j = LBound(r, 2) To UBound(r, 2)
'    For i = LBound(r) To UBound(r)
'      Debug.Print i, j, r(i, j)
'    Next i
'  Next j
'
'  Set r = Nothing
End Sub
Sub ViewsTableStylesLOStest()
  Dim rbox As Range
  Dim LO As ListObject
  On Error GoTo vEH
  
  Set rbox = Range("A23")
  Set LO = rbox.Parent.ListObjects.Add(xlSrcRange, rbox.Range("A1").CurrentRegion, XlListObjectHasHeaders:=xlYes)
  
  With LO
    Let .ShowTableStyleFirstColumn = True
    Let .ShowTotals = True
    'Select Case xType
    '  Case 0: .TableStyle = VBA.vbNullString: Call .Unlist
    '  Case 1: Call .Unlist
    'End Select
  End With

vEH:
  Set LO = Nothing
  Set rbox = Nothing
End Sub

Sub resExcelRangeHasDatas()
  Dim xRng As Range
  Dim VData As Variant
  Dim i As Long, j As Long
  Dim resExcelRangeHasData As Boolean
  
  On Error GoTo EH
  Set xRng = Range("B1:F14")
  Let resExcelRangeHasData = False
  If xRng.Rows.Count = 1 And xRng.Columns.Count = 1 Then
    Let resExcelRangeHasData = Not (xRng.Value = Empty)
  Else
    Let VData = xRng.Value
    For i = LBound(VData) To UBound(VData)
      For j = LBound(VData, 2) To UBound(VData, 2)
        If Not (VData(i, j) = Empty) Then Let resExcelRangeHasData = True: Exit For
      Next j
      If resExcelRangeHasData Then Exit For
    Next i
    Erase VData
  End If
  Debug.Print resExcelRangeHasData

EH:
  Erase VData
  'Call EHGlobal.ErrorHandlerRaise("RES::ExcelEdges")
End Sub

Sub Ncells()
  Dim Sh As Worksheet
  Dim Rn As Range
  
  Set Sh = ThisWorkbook.ActiveSheet
  Set Rn = Sh.Range("A8")
  
  Debug.Print Sh.Rows.Count, Sh.Columns.Count
  Debug.Print Rn.row, Rn.Column
  Debug.Print Rn.row + 25, Rn.Column + 20
  Debug.Print Rn.row + 25 > Sh.Rows.Count, Rn.Column + 20 > Sh.Columns.Count
  
  Set Rn = Nothing
  Set Sh = Nothing
End Sub

Sub arrs()
  Dim a(0 To 2, 0 To 0) As Variant
  Dim r As Range
  
  Let a(0, 0) = 1
  Let a(1, 0) = 2
  Let a(2, 0) = 3
  
  'Let excliqlites.Range("A1:A3").Value2 = a
  Set r = excliqlites.Range("A8")
  
  Call r.Range("B2").Select
  Call r.Range("A1").Select
  Call r.Range("A1").Offset(2, 7).Select
  Debug.Print "FORMULA(" & _
              r.Range("C3:C6").Address & ", " & _
              r.Range("B3:B6").Address & ", " & _
              r.Range("A1").Offset(2, 7).Address(False, False) & ")"
  Erase a
  Set r = Nothing
End Sub


'Option Explicit
'Option Private Module
'
''Para todas las interfaces ocultas
''Una sola función para Action
''Una sola función para Visibility
''Una sola función para Enabling
''Una sola función para Reset
'
''El concepto del sistema indica la lectura de  valores de la DB a partir de métodos que requieren los datos desde el Controlador
''a partir de los resultados generados se construyen las vistas finales que se entregan al usuario en forma de reportes e información terminal.
''El método CRUD es aplicado específicamente al Modelo.
''Las lecturas de cada tabla de la DB se harán con esticto CRUD
''Las lecturas de COMBINACIONES DE TABLAS DE LA DB se harán con CRUD para el Modelo específico de esa COMBINACIÓN
'
''ERRORES PERSONALIZADOS: _
'APP: 514 _
'RIBBON: 515 _
'CLASSES: 516 _
'VIEWS: 517 _
'CONTROLLERS: 518 _
'MODELS: 519
'
'Private exq As New AppExcliq 'Dim exq As New AppExcliq
''Public EHQ_SRC As String
'Public EHGlobal As New AppErrorHandler
'
''Declaraciones para pruebas
''Dim ntos As New AppResNumberToString
'
''MÉTODOS DE PRUEBA
''Sub test()
''End Sub
'
'
'
''METHODS
''Ribbon activa el sistema llamando a excliqRibbonBegin
'Private Sub auto_open()
'  If exq Is Nothing Then Set exq = New AppExcliq
'End Sub
'
'Private Sub auto_close()
'  If Application.ActiveWorkbook.Saved Then
'    MsgBox "Cerrando Excliq."
'    Set exq = Nothing
'  End If
'End Sub
'
''--- RIBBON FUNCTIONS ---
'Private Sub ExcliqRibbonBeginTest()
'  Application.Volatile
'  Call exq.AppInitTest
''ESM:
'End Sub
'
'Public Sub ExcliqRibbonBegin(ByVal IRIBBON As IRibbonUI)
'  Application.Volatile
'    Call exq.AppInit(IRIBBON)
'ERB:
'End Sub
'
''... Events: Status controls ...
'Public Sub ExcliqRibbonGetEnabled(ByVal Control As IRibbonControl, ByRef StatusControl As Variant)
'  'Let StatusControl = True 'exq.AppRibbonStatusSetter(control.ID, true) 'Para estado habilitado/deshabilitado del menú
'  Let StatusControl = exq.AppRibbonStatusSetter(Control.ID, True) 'Para estado habilitado/deshabilitado del menú
'End Sub
'
''Public Sub ExcliqRibbonStatus(ByVal control As IRibbonControl, ByRef StatusControl As Variant)
'Public Sub ExcliqRibbonGetVisible(ByVal Control As IRibbonControl, ByRef StatusControl As Variant)
'  'MsgBox "Ribbon: Fui llamado Por: " & Control.id & " | Valor: " & VBA.Mid(Control.id, 3) 'uRIBBONCOLLECTION.Item(Control.ID)
'  'StatusControl = uRIBBONCOLLECTION.Item(Control.ID)
'    'StatusControl = True
'    StatusControl = exq.AppRibbonStatusSetter(Control.ID)
'End Sub
'
'Public Sub ExcliqRibbonMessage(ByVal Control As IRibbonControl, ByRef Msg As Variant)
'  Msg = exq.AppRibbonStatusSetter(Control.ID)
'End Sub
'
''... Functionality subs ...
'Sub ExcliqRibbonActionExecutor(ByVal Control As IRibbonControl)
'  Application.Volatile
'  MsgBox "Control: " & VBA.LCase(VBA.Mid(Control.ID, 7)) & ", " & VBA.LCase(VBA.Left(Control.ID, 3))
'  'Call exq.AppRibbonExecutor(VBA.LCase(VBA.Mid(Control.ID, 7)), VBA.LCase(VBA.Left(Control.ID, 3)))
'  Call exq.AppRibbonExecutorControls(VBA.LCase(VBA.Mid(Control.ID, 7)), VBA.LCase(VBA.Left(Control.ID, 3)))
'End Sub
'
'Sub ExcliqRibbonActionExecutorTest()
'  Dim Order As String
'  Application.Volatile
'  'SysBtnOpenSession
'  'SysBtnCloseSession
'  'SysBtnNewUser
'  'SysBtnNewTerminal
'
'  'ConBtnProfile
'  'ConBtnPrivacity
'  'ConBtnUsers
'  Order = "ConBtnUsers"
'  Call exq.AppRibbonExecutorControls(VBA.LCase(VBA.Mid(Order, 7)), VBA.LCase(VBA.Left(Order, 3)))
'End Sub
'
'
'
'
'
''System Tab
''Sub ExcliqRibbonSystem(ByVal Control As IRibbonControl)
''  Application.Volatile
''  Call exq.CallViewSystem(Control.ID)
''ESM:
''End Sub
''Sub ExcliqRibbonSystemTest()
''  Application.Volatile
''  Call exq.CallViewSystem("btnNewUser") 'btnNewTerminal") 'btnNewUser") 'btnOpenSession")
''ESM:
''End Sub
''
'''Simulations Tab
''Sub ExcliqRibbonSimulations(ByVal Control As IRibbonControl)
''  Application.Volatile
''  Call exq.CallViewSimulations(Control.ID)
'''  Select Case Control.ID
'''    Case "btnNewRegressionSelection"
'''      Call exq.moduleRegressionSelection(Selection)
'''    Case "btnNewBalmetSelection"
'''      Call exq.moduleBalmetSelection(Selection)
'''    Case Else
'''      Call exq.simulationsView(Control.ID)
'''  End Select
''ESM:
''End Sub
'
''Gestios Tab
''Config Tab
''Help Tab
''--- RIBBON FUNCTIONS ---
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
''
''Sub begin(ByVal control As IRibbonControl, ByRef Enable)
''  'Call exq.init
''  Dim Enable As Boolean
''  If Not exq.APPINITIALIZED Then
''    Call exq.init
''  End If
''  Select Case control.ID
''    Case "tabSystem"
''      Enable = exq.APPTABSYS
''    Case "tabGestion"
''      Enable = exq.APPTABGES
''    Case "tabConfig"
''      Enable = exq.APPTABCON
''    Case "btnOpenSession"
''      Enable = exq.APPSESSIONSTATUS
''    Case "btnCloseSession"
''      Enable = (Not exq.APPSESSIONSTATUS)
''    Case "btnNewUsers"
''      Enable = exq.APPNEWUSERS
''  End Select
''End Sub
'
'
'
'
'
''SYSTEM SUBROUTINES
''Sub beginSession()
''  Call exq.sessionStartView
'''  If excliq.sessionStart(usr, key) Then
'''    userMenu (excliq.CURRENTUSER)
'''  Else
'''    MsgBox "Datos incorrectos.", , "Excliq"
'''  End If
''End Sub
'
''Sub sessionBtnAction(ByVal action As String, frm As FRMUsers)
''  If action = "Iniciar" Then
''    If exq.sessionStart(frm.USRNameTextBox.value, frm.USRPassTextBox.value) Then
''      frm.USRNameTextBox.SetFocus
''    Else
''      Unload frm
''    End If
''  ElseIf action = "Registrar Usuario" Then
''    If NEWUSERS(frm.USRNameTextBox.value, frm.USRPassTextBox.value, Mid(frm.USRRoleComboBox.value, 1, 1), Mid(frm.USRJobComboBox.value, 1, 1)) Then
''      frm.USRNameTextBox.SetFocus
''    Else
''      Unload frm
''    End If
''  End If
''End Sub
'
''Sub usersRegisterFrm()
''  Call exq.enableUserView
''End Sub
''
''Sub usersRegister()
''  'exq.enableUserView
''End Sub
''
''Function NEWUSERS(ByVal usrStr As String, ByVal UserKey As String, ByVal UserRole As String, ByVal userJob As String) As Boolean
''  NEWUSERS = exq.NEWUSERS(usrStr, UserKey, UserRole, userJob)
''End Function
''
''Sub closeSession()
''  If MsgBox("¿Seguro que desea cerrar la sesión?", vbYesNo, "Excliq") = vbYes Then
''    Call exq.sessionStop(exq.CURRENTUSER)
''  End If
''End Sub
'''Optional FullID As String = vba.vbnullstring,
''Sub systemRegister()
''  'Registrar nuevo usuario
''  'exq.enableExcliq
''End Sub
'''FORMS SUBROUTINES
''Public Function moduleSimulationTextNumbers(ByVal numData As MSForms.ReturnInteger, ByRef txtData As Object, Optional Shft As Integer, Optional eventType As Boolean = False, Optional TxtOrNum As Boolean = False) As Variant
''  moduleSimulationTextNumbers = exq.moduleSimulationNumbers(numData, txtData, Shft, eventType, TxtOrNum)
''End Function
''
'Sub ExcliqFormsActionExecutor(ByVal eAction As String, ByVal eFrm As Object, Optional eCtrl As Object)
'  'Call exq.moduleRegressionAction(action, frm)
'  'Call exq.AppFrmCtrlAction(action, frm, ctrl)
'  'Call exq.AppRibbonExecutor(action:=VBA.LCase(eAction), frm:=eFrm, ctrl:=eCtrl)
'
'  'Call exq.AppRibbonExecutorForms(VBA.LCase(eAction), eFrm, eCtrl)
'End Sub
'
''CREDENTIALS BRIDGE
'Private Function EXCLIQ_GetCredentials(ByVal Petition As String) As String
'  Let EXCLIQ_GetCredentials = exq.AppExcliqGetCredentials(Petition)
'End Function
'''SIMULATION SUBROUTINES
'
'
'
'
'
''Function moduleSimulationBtnAction(ByVal action As String, ByVal frm As FRMBasics, Optional Ctrl As Object)
''  'Call exq.moduleRegressionAction(action, frm)
''  Call exq.moduleSimulationAction(action, frm, Ctrl)
''End Function
'''Regression
''Function moduleSimulationRegressionBtnAction(ByVal action As String, frm As Variant)
''  Call exq.moduleRegressionAction(action, frm)
''End Function
''Sub moduleSimulationRegressionBtnResultsOLD(ByVal allResults As Boolean, ByVal x As Variant, ByVal y As Variant, ByVal XA As Variant, ByVal coord As String, ByVal aSheet As String, ByVal graph As String, datashow() As Variant)
''  Debug.Print UBound(datashow)
''  Dim i As Integer
''  For i = 0 To UBound(datashow)
''    Debug.Print i & ". " & datashow(i) & " - " & UBound(datashow)
''  Next i
''  exq.moduleRegressionResults allResults, x, y, XA, coord, aSheet, graph, datashow
''End Sub
'''Balance
''Sub balanceEvents(ByVal frm As FRMBalance, Optional action As String = "Cancel", Optional controlA As Object, Optional controlB As Object)
''  Call exq.moduleBalanceEvents(frm, action, controlA, controlB)
''End Sub
''
'''GESTION SUBROUTINES
'''CONFIG SUBROUTINES
'''APP SUBROUTINES
''Sub formsMannager(frm As Object)
''  'Dim frmReg As New EnterprisesUserForm 'UsersUserForm
''  'Dim dataReg() As Variant
''  Load frm
''  frm.Show
''
''  'mostrar frmSession y pasar parámetros
'''  If exq.sessionStart(usr, key) Then
'''    userMenu (exq.CURRENTUSER)
'''  Else
'''    MsgBox "Datos incorrectos.", , "Excliq"
'''  End If
''End Sub
''
''Private Sub userMenu(ByVal usr As Integer)
'''  Dim userfrm As New UsersUserForm
'''  Dim a As Variant
'''  a = Array(1, 2, 3)
'''  For i = 0 To UBound(a)
'''    userfrm.USRJobComboBox.AddItem a(i), i
'''    userfrm.USRJobComboBox.ControlTipText = userfrm.USRJobComboBox.Value
'''  Next i
''  If usr = 0 Then
''    'Súper Usuario
''  Else
''    'Usuarios normales
''  End If
''End Sub
''
''Sub enableSystem(ByVal kc As String)
''  exq.enableExcliq kc
''End Sub
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'''El concepto del sistema indica la lectura de  valores de la DB a partir de métodos que requieren los datos desde el Controlador
'''a partir de los resultados generados se construyen las vistas finales que se entregan al usuario en forma de reportes e información terminal.
'''El método CRUD es aplicado específicamente al Modelo.
'''Las lecturas de cada tabla de la DB se harán con esticto CRUD
'''Las lecturas de COMBINACIONES DE TABLAS DE LA DB se harán con CRUD para el Modelo específico de esa COMBINACIÓN
''Option Explicit
''Option Private Module
''Dim exq As New AppExcliq
''Dim ress As New AppResources
'''Declaraciones para pruebas
''Dim ntos As New AppResNumberToString
''
'''MÉTODOS DE PRUEBA
''Sub test()
''  Call BALANCE
''  'exq.moduleRegressionView exq.SESSIONSTARTED
''  'ress.checkAppEdges Sheets("Hoja1").Range("A1")
''  'ModuleSimulationRegressionForm.Show
'''  Dim num As Double
'''  num = CDbl(InputBox("Un número:", "Excliq", "123456789101112"))
'''  MsgBox FormatNumber(Round(num, 2), 2) & vbNewLine & vbNewLine & ntos.numToString(num, "Boliviano.", "Bolivianos.", True)
''  'MsgBox ntos.numACadena(num)
''End Sub
'''Public Function sayHello(ByVal site As String) As String
'''Public Function sayHello(site)
''''Función que ayuda a saludar
'''  sayHello = "Hello World! " & site
'''End Function
''
''
'''METHODS
''Sub begin()
''  exq.init
''End Sub
''
'''SYSTEM SUBROUTINES
''Sub beginSession()
'''  Dim usr, key As String
'''  Dim newSession As Object
'''  Set newSession = New UsersUserForm
'''  'mostrar frmSession y pasar parámetros
'''  newSession.Caption = "Excliq - Iniciar Sesión"
'''  newSession.USRJobLabel.Visible = False
'''  newSession.USRJobComboBox.Visible = False
'''  newSession.USRRoleLabel.Visible = False
'''  newSession.USRTypeComboBox.Visible = False
'''  newSession.USRRegCommandButton.Caption = "Iniciar"
'''  newSession.USRCancelCommandButton.Caption = "Cancelar"
'''  formsMannager newSession
''
''  exq.sessionStartView
''
'''  If exq.sessionStart(usr, key) Then
'''    userMenu (exq.CURRENTUSER)
'''  Else
'''    MsgBox "Datos incorrectos.", , "Excliq"
'''  End If
''End Sub
''
''Sub sessionBtnAction(ByVal action As String, frm As FRMUsers)
''  If action = "Iniciar" Then
''    If exq.sessionStart(frm.USRNameTextBox.value, frm.USRPassTextBox.value) Then
''      frm.USRNameTextBox.SetFocus
''    Else
''      Unload frm
''    End If
''  ElseIf action = "Registrar Usuario" Then
''    If newUsers(frm.USRNameTextBox.value, frm.USRPassTextBox.value, Mid(frm.USRRoleComboBox.value, 1, 1), Mid(frm.USRJobComboBox.value, 1, 1)) Then
''      frm.USRNameTextBox.SetFocus
''    Else
''      Unload frm
''    End If
''  End If
''End Sub
''
''Sub usersRegisterFrm()
''  exq.enableUserView
''End Sub
''
''Sub usersRegister()
''  'exq.enableUserView
''End Sub
''
''Function newUsers(ByVal usrStr As String, ByVal UserKey As String, ByVal userRole As String, ByVal userJob As String) As Boolean
''  newUsers = exq.newUsers(usrStr, UserKey, userRole, userJob)
''End Function
''
''Sub closeSession()
''  If MsgBox("¿Seguro que desea cerrar la sesión?", vbYesNo, "Excliq") = vbYes Then
''    exq.sessionStop (exq.CURRENTUSER)
''  End If
''End Sub
''
''Sub systemRegister()
''  'Registrar nuevo usuario
''  'exq.enableExcliq
''End Sub
''
'''SIMULATION SUBROUTINES
'''Regression
''Sub REGRESSION()
''  exq.moduleRegressionView 'exq.SESSIONSTARTED
''End Sub
''Function moduleSimulationRegressionBtnAction(ByVal action As String, frm As Variant) ' As Boolean
''  'moduleSimulationRegressionBtnAction = exq.moduleRegressionAction(action, frm)
''  Call exq.moduleRegressionAction(action, frm)
''End Function
''
''Sub moduleSimulationRegressionBtnResultsOLD(ByVal allResults As Boolean, ByVal x As Variant, ByVal y As Variant, ByVal XA As Variant, ByVal coord As String, ByVal aSheet As String, ByVal graph As String, datashow() As Variant)
'''Sub moduleSimulationRegressionBtnResults(ByVal allResults As Boolean, ByVal X As Variant, ByVal Y As Variant, ByVal XA As Variant, ByVal coord As String, ByVal aSheet As String, ByVal graph As String, ParamArray datashow() As Variant)
''  Debug.Print UBound(datashow)
''  Dim i As Integer
''  For i = 0 To UBound(datashow)
''    Debug.Print i & ". " & datashow(i) & " - " & UBound(datashow)
''  Next i
''  exq.moduleRegressionResults allResults, x, y, XA, coord, aSheet, graph, datashow
''End Sub
'''Balance
''Sub BALANCE()
''  Call exq.moduleBalanceView
''End Sub
''Sub balanceEvents(ByVal frm As FRMBalance, Optional action As String = "Cancel", Optional controlA As Object, Optional controlB As Object)
''  Call exq.moduleBalanceEvents(frm, action, controlA, controlB)
''End Sub
''
'''GESTION SUBROUTINES
'''CONFIG SUBROUTINES
'''APP SUBROUTINES
''Sub formsMannager(frm As Object)
''  'Dim frmReg As New EnterprisesUserForm 'UsersUserForm
''  'Dim dataReg() As Variant
''  Load frm
''  frm.Show
''
''  'mostrar frmSession y pasar parámetros
'''  If exq.sessionStart(usr, key) Then
'''    userMenu (exq.CURRENTUSER)
'''  Else
'''    MsgBox "Datos incorrectos.", , "Excliq"
'''  End If
''End Sub
''
''Private Sub userMenu(ByVal usr As Integer)
'''  Dim userfrm As New UsersUserForm
'''  Dim a As Variant
'''  a = Array(1, 2, 3)
'''  For i = 0 To UBound(a)
'''    userfrm.USRJobComboBox.AddItem a(i), i
'''    userfrm.USRJobComboBox.ControlTipText = userfrm.USRJobComboBox.Value
'''  Next i
''  If usr = 0 Then
''    'Súper Usuario
''  Else
''    'Usuarios normales
''  End If
''End Sub
''
''Sub enableSystem(ByVal kc As String)
''  exq.enableExcliq kc
''End Sub
'
'
