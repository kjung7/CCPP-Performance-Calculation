VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGas_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Gas As New clsStream_
Public Fuel As New clsStream_

Public Composition As New clsComposition_

Public w As New Dictionary
Public LHV As New Dictionary

Public T_fuel, P_fuel, m_fuel As Variant
Public LHV_fuel As Variant
Public mode As Integer

Public Power_de, Power_map As Variant
Public Thermal_effi_de, Thermal_effi_map As Variant

Function Gas_in() As Variant

    Dim HRSG_input As Worksheet
    Dim Property_fuel As Variant
    Dim r As Double
    Dim key As Variant
    
    Set HRSG_input = ThisWorkbook.Worksheets("HRSG_input")
    
    Gas.Composition.H2O = HRSG_input.Cells(7, "X")
    Gas.Composition.O2 = HRSG_input.Cells(8, "X")
    Gas.Composition.CO2 = HRSG_input.Cells(9, "X")
    Gas.Composition.N2 = HRSG_input.Cells(10, "X")
    Gas.Composition.Ar = HRSG_input.Cells(11, "X")
    Gas.Composition.H2 = HRSG_input.Cells(12, "X")
    Gas.Composition.CO = HRSG_input.Cells(13, "X")
    Gas.Composition.CH4 = HRSG_input.Cells(14, "X")
    Gas.Composition.C2H6 = HRSG_input.Cells(15, "X")
    Gas.Composition.C3H8 = HRSG_input.Cells(16, "X")
    Gas.Composition.n_C4H10 = HRSG_input.Cells(17, "X")
    Gas.Composition.i_C4H10 = HRSG_input.Cells(18, "X")
    Gas.Composition.n_C5H12 = HRSG_input.Cells(19, "X")
    Gas.Composition.i_C5H12 = HRSG_input.Cells(20, "X")
    Gas.Composition.n_C6H14 = HRSG_input.Cells(21, "X")
    Gas.Composition.n_C7H16 = HRSG_input.Cells(22, "X")
    Gas.Composition.n_C8H18 = HRSG_input.Cells(23, "X")
    Gas.Composition.He = HRSG_input.Cells(24, "X")
    Gas.Composition.H2S = HRSG_input.Cells(25, "X")
    Gas.Composition.SO2 = HRSG_input.Cells(26, "X")
    Gas.Composition.Comp_save
    
    
     'Molecular weight
    If w.Exists("H2O") Then
       w("H2O") = 18.01528
       w("O2") = 31.9988
       w("CO2") = 44.0095
       w("N2") = 28.0134
       w("Ar") = 39.948
       w("H2") = 2.01588
       w("CO") = 28.0101
       w("CH4") = 16.04246
       w("C2H6") = 30.06904
       w("C3H8") = 44.09562
       w("n_C4H10") = 58.1222
       w("i_C4H10") = 58.1222
       w("n_C5H12") = 72.14878
       w("i_C5H12") = 72.14878
       w("n_C6H14") = 86.17536
       w("n_C7H16") = 100.20194
       w("n_C8H18") = 114.22852
       w("He") = 4.0026
       w("H2S") = 34.08088
       w("SO2") = 64.064
    Else
    w.Add "H2O", 18.01528
    w.Add "O2", 31.9988
    w.Add "CO2", 44.0095
    w.Add "N2", 28.0134
    w.Add "Ar", 39.948
    w.Add "H2", 2.01588
    w.Add "CO", 28.0101
    w.Add "CH4", 16.04246
    w.Add "C2H6", 30.06904
    w.Add "C3H8", 44.09562
    w.Add "n_C4H10", 58.1222
    w.Add "i_C4H10", 58.1222
    w.Add "n_C5H12", 72.14878
    w.Add "i_C5H12", 72.14878
    w.Add "n_C6H14", 86.17536
    w.Add "n_C7H16", 100.20194
    w.Add "n_C8H18", 114.22852
    w.Add "He", 4.0026
    w.Add "H2S", 34.08088
    w.Add "SO2", 64.064
    End If
    
    'LHV
    If LHV.Exists("H2") Then
       LHV("H2O") = 0
       LHV("O2") = 0
       LHV("CO2") = 0
       LHV("N2") = 0
       LHV("Ar") = 0
       LHV("H2") = 119914.646208418
       LHV("CO") = 10100.3201975624
       LHV("CH4") = 50033.3833698394
       LHV("C2H6") = 47518.6701404084
       LHV("C3H8") = 46340.3229417385
       LHV("n_C4H10") = 45725.2111392059
       LHV("i_C4H10") = 45567.2649949597
       LHV("n_C5H12") = 45351.6910185806
       LHV("i_C5H12") = 45255.6703753042
       LHV("n_C6H14") = 45109.0362769398
       LHV("n_C7H16") = 44927.3348363418
       LHV("n_C8H18") = 44789.2869896708
       LHV("He") = 0
       LHV("H2S") = 15198.8830187702
       LHV("SO2") = 0
    Else
    LHV.Add "H2O", 0
    LHV.Add "O2", 0
    LHV.Add "CO2", 0
    LHV.Add "N2", 0
    LHV.Add "Ar", 0
    LHV.Add "H2", 119914.646208418
    LHV.Add "CO", 10100.3201975624
    LHV.Add "CH4", 50033.3833698394
    LHV.Add "C2H6", 47518.6701404084
    LHV.Add "C3H8", 46340.3229417385
    LHV.Add "n_C4H10", 45725.2111392059
    LHV.Add "i_C4H10", 45567.2649949597
    LHV.Add "n_C5H12", 45351.6910185806
    LHV.Add "i_C5H12", 45255.6703753042
    LHV.Add "n_C6H14", 45109.0362769398
    LHV.Add "n_C7H16", 44927.3348363418
    LHV.Add "n_C8H18", 44789.2869896708
    LHV.Add "He", 0
    LHV.Add "H2S", 15198.8830187702
    LHV.Add "SO2", 0
    End If

    r = 8.31451 '기체상수

    Gas.Properties.P = HRSG_input.Cells(4, "X")
    Gas.Properties.T = HRSG_input.Cells(5, "X")
 
    Gas.Properties.MW = 0
    
    For Each key In w
        Gas.Properties.MW = Gas.Properties.MW + Gas.Composition.Comp(key) * w(key)
    Next key
    
    '물성치 계산
    Property_fuel = idealgas("hsck", "p", Gas.Properties.P, "t", Gas.Properties.T, Gas.Composition.Composition)

    Gas.Properties.h = Property_fuel(0)
    Gas.Properties.s = Property_fuel(1)
    Gas.Properties.C = Property_fuel(2)
    Gas.Properties.k = Property_fuel(3)

    'LHV 계산
    LHV_fuel = 0
    For Each key In LHV
        LHV_fuel = LHV_fuel + LHV(key) * w(key) * Gas.Composition.Comp(key)
    Next key

    LHV_fuel = LHV_fuel / Gas.Properties.MW

    '기체상수
    Gas.Properties.r = r / Gas.Properties.MW

    Gas.Properties.Prop_save
    
    Gas.Massflow.m = HRSG_input.Cells(3, "X") * 10 / 36
    Gas.Massflow.Mass_save

End Function

Function Gas_in_off() As Variant
    Dim HRSG_input_off As Worksheet

    Set HRSG_input_off = ThisWorkbook.Worksheets("HRSG_input_off")

    Dim Property_fuel As Variant
    Dim r As Double
    Dim key As Variant
'    Dim dummy As Variant
    
    Gas.Composition.H2O = HRSG_input_off.Cells(7, "N")
    Gas.Composition.O2 = HRSG_input_off.Cells(8, "N")
    Gas.Composition.CO2 = HRSG_input_off.Cells(9, "N")
    Gas.Composition.N2 = HRSG_input_off.Cells(10, "N")
    Gas.Composition.Ar = HRSG_input_off.Cells(11, "N")
    Gas.Composition.H2 = HRSG_input_off.Cells(12, "N")
    Gas.Composition.CO = HRSG_input_off.Cells(13, "N")
    Gas.Composition.CH4 = HRSG_input_off.Cells(14, "N")
    Gas.Composition.C2H6 = HRSG_input_off.Cells(15, "N")
    Gas.Composition.C3H8 = HRSG_input_off.Cells(16, "N")
    Gas.Composition.n_C4H10 = HRSG_input_off.Cells(17, "N")
    Gas.Composition.i_C4H10 = HRSG_input_off.Cells(18, "N")
    Gas.Composition.n_C5H12 = HRSG_input_off.Cells(19, "N")
    Gas.Composition.i_C5H12 = HRSG_input_off.Cells(20, "N")
    Gas.Composition.n_C6H14 = HRSG_input_off.Cells(21, "N")
    Gas.Composition.n_C7H16 = HRSG_input_off.Cells(22, "N")
    Gas.Composition.n_C8H18 = HRSG_input_off.Cells(23, "N")
    Gas.Composition.He = HRSG_input_off.Cells(24, "N")
    Gas.Composition.H2S = HRSG_input_off.Cells(25, "N")
    Gas.Composition.SO2 = HRSG_input_off.Cells(26, "N")
    Gas.Composition.Comp_save
    'Molecular weight
    If w.Exists("H2O") Then
       w("H2O") = 18.01528
       w("O2") = 31.9988
       w("CO2") = 44.0095
       w("N2") = 28.0134
       w("Ar") = 39.948
       w("H2") = 2.01588
       w("CO") = 28.0101
       w("CH4") = 16.04246
       w("C2H6") = 30.06904
       w("C3H8") = 44.09562
       w("n_C4H10") = 58.1222
       w("i_C4H10") = 58.1222
       w("n_C5H12") = 72.14878
       w("i_C5H12") = 72.14878
       w("n_C6H14") = 86.17536
       w("n_C7H16") = 100.20194
       w("n_C8H18") = 114.22852
       w("He") = 4.0026
       w("H2S") = 34.08088
       w("SO2") = 64.064
    Else
    w.Add "H2O", 18.01528
    w.Add "O2", 31.9988
    w.Add "CO2", 44.0095
    w.Add "N2", 28.0134
    w.Add "Ar", 39.948
    w.Add "H2", 2.01588
    w.Add "CO", 28.0101
    w.Add "CH4", 16.04246
    w.Add "C2H6", 30.06904
    w.Add "C3H8", 44.09562
    w.Add "n_C4H10", 58.1222
    w.Add "i_C4H10", 58.1222
    w.Add "n_C5H12", 72.14878
    w.Add "i_C5H12", 72.14878
    w.Add "n_C6H14", 86.17536
    w.Add "n_C7H16", 100.20194
    w.Add "n_C8H18", 114.22852
    w.Add "He", 4.0026
    w.Add "H2S", 34.08088
    w.Add "SO2", 64.064
    End If
    
    'LHV
    If LHV.Exists("H2") Then
       LHV("H2O") = 0
       LHV("O2") = 0
       LHV("CO2") = 0
       LHV("N2") = 0
       LHV("Ar") = 0
       LHV("H2") = 119914.646208418
       LHV("CO") = 10100.3201975624
       LHV("CH4") = 50033.3833698394
       LHV("C2H6") = 47518.6701404084
       LHV("C3H8") = 46340.3229417385
       LHV("n_C4H10") = 45725.2111392059
       LHV("i_C4H10") = 45567.2649949597
       LHV("n_C5H12") = 45351.6910185806
       LHV("i_C5H12") = 45255.6703753042
       LHV("n_C6H14") = 45109.0362769398
       LHV("n_C7H16") = 44927.3348363418
       LHV("n_C8H18") = 44789.2869896708
       LHV("He") = 0
       LHV("H2S") = 15198.8830187702
       LHV("SO2") = 0
    Else
    LHV.Add "H2O", 0
    LHV.Add "O2", 0
    LHV.Add "CO2", 0
    LHV.Add "N2", 0
    LHV.Add "Ar", 0
    LHV.Add "H2", 119914.646208418
    LHV.Add "CO", 10100.3201975624
    LHV.Add "CH4", 50033.3833698394
    LHV.Add "C2H6", 47518.6701404084
    LHV.Add "C3H8", 46340.3229417385
    LHV.Add "n_C4H10", 45725.2111392059
    LHV.Add "i_C4H10", 45567.2649949597
    LHV.Add "n_C5H12", 45351.6910185806
    LHV.Add "i_C5H12", 45255.6703753042
    LHV.Add "n_C6H14", 45109.0362769398
    LHV.Add "n_C7H16", 44927.3348363418
    LHV.Add "n_C8H18", 44789.2869896708
    LHV.Add "He", 0
    LHV.Add "H2S", 15198.8830187702
    LHV.Add "SO2", 0
    End If
    r = 8.31451 '기체상수
    
    Gas.Properties.P = HRSG_input_off.Cells(4, "N")
    Gas.Properties.T = HRSG_input_off.Cells(5, "N")
 
    Gas.Properties.MW = 0
    
    For Each key In w
        Gas.Properties.MW = Gas.Properties.MW + Gas.Composition.Comp(key) * w(key)
    Next key

    '물성치 계산
    Property_fuel = idealgas("hsck", "p", Gas.Properties.P, "t", Gas.Properties.T, Gas.Composition.Composition)

    Gas.Properties.h = Property_fuel(0)
    Gas.Properties.s = Property_fuel(1)
    Gas.Properties.C = Property_fuel(2)
    Gas.Properties.k = Property_fuel(3)

    'LHV 계산
    LHV_fuel = 0
    For Each key In LHV
        LHV_fuel = LHV_fuel + LHV(key) * w(key) * Gas.Composition.Comp(key)
    Next key

    LHV_fuel = LHV_fuel / Gas.Properties.MW

    '기체상수
    Gas.Properties.r = r / Gas.Properties.MW

    Gas.Properties.Prop_save
    
    Gas.Massflow.m = HRSG_input_off.Cells(3, "N") * 10 / 36
    Gas.Massflow.Mass_save
    
'    dummy = idealgas("h", "p", Gas.Properties.P, "t", 156.9, Gas.Composition.Composition)
'    MsgBox dummy(0)

End Function

Function Fuel_in() As Variant
    
    Dim HRSG_input As Worksheet

    Set HRSG_input = ThisWorkbook.Worksheets("HRSG_input")

    Dim Property_fuel As Variant
    Dim r As Double
    Dim key As Variant
    
    Fuel.Composition.H2O = HRSG_input.Cells(8 + 42, "X")
    Fuel.Composition.O2 = HRSG_input.Cells(9 + 42, "X")
    Fuel.Composition.CO2 = HRSG_input.Cells(10 + 42, "X")
    Fuel.Composition.N2 = HRSG_input.Cells(11 + 42, "X")
    Fuel.Composition.Ar = HRSG_input.Cells(12 + 42, "X")
    Fuel.Composition.H2 = HRSG_input.Cells(13 + 42, "X")
    Fuel.Composition.CO = HRSG_input.Cells(14 + 42, "X")
    Fuel.Composition.CH4 = HRSG_input.Cells(15 + 42, "X")
    Fuel.Composition.C2H6 = HRSG_input.Cells(16 + 42, "X")
    Fuel.Composition.C3H8 = HRSG_input.Cells(17 + 42, "X")
    Fuel.Composition.n_C4H10 = HRSG_input.Cells(18 + 42, "X")
    Fuel.Composition.i_C4H10 = HRSG_input.Cells(19 + 42, "X")
    Fuel.Composition.n_C5H12 = HRSG_input.Cells(20 + 42, "X")
    Fuel.Composition.i_C5H12 = HRSG_input.Cells(21 + 42, "X")
    Fuel.Composition.n_C6H14 = HRSG_input.Cells(22 + 42, "X")
    Fuel.Composition.n_C7H16 = HRSG_input.Cells(23 + 42, "X")
    Fuel.Composition.n_C8H18 = HRSG_input.Cells(24 + 42, "X")
    Fuel.Composition.He = HRSG_input.Cells(25, "X")
    Fuel.Composition.H2S = HRSG_input.Cells(26, "X")
    Fuel.Composition.SO2 = HRSG_input.Cells(27, "X")
    Fuel.Composition.Comp_save
    r = 8.31451 '기체상수
    
    'Molecular weight
    If w.Exists("H2O") Then
       w("H2O") = 18.01528
       w("O2") = 31.9988
       w("CO2") = 44.0095
       w("N2") = 28.0134
       w("Ar") = 39.948
       w("H2") = 2.01588
       w("CO") = 28.0101
       w("CH4") = 16.04246
       w("C2H6") = 30.06904
       w("C3H8") = 44.09562
       w("n_C4H10") = 58.1222
       w("i_C4H10") = 58.1222
       w("n_C5H12") = 72.14878
       w("i_C5H12") = 72.14878
       w("n_C6H14") = 86.17536
       w("n_C7H16") = 100.20194
       w("n_C8H18") = 114.22852
       w("He") = 4.0026
       w("H2S") = 34.08088
       w("SO2") = 64.064
    Else
    w.Add "H2O", 18.01528
    w.Add "O2", 31.9988
    w.Add "CO2", 44.0095
    w.Add "N2", 28.0134
    w.Add "Ar", 39.948
    w.Add "H2", 2.01588
    w.Add "CO", 28.0101
    w.Add "CH4", 16.04246
    w.Add "C2H6", 30.06904
    w.Add "C3H8", 44.09562
    w.Add "n_C4H10", 58.1222
    w.Add "i_C4H10", 58.1222
    w.Add "n_C5H12", 72.14878
    w.Add "i_C5H12", 72.14878
    w.Add "n_C6H14", 86.17536
    w.Add "n_C7H16", 100.20194
    w.Add "n_C8H18", 114.22852
    w.Add "He", 4.0026
    w.Add "H2S", 34.08088
    w.Add "SO2", 64.064
    End If
    
    'LHV
    If LHV.Exists("H2") Then
       LHV("H2O") = 0
       LHV("O2") = 0
       LHV("CO2") = 0
       LHV("N2") = 0
       LHV("Ar") = 0
       LHV("H2") = 119914.646208418
       LHV("CO") = 10100.3201975624
       LHV("CH4") = 50033.3833698394
       LHV("C2H6") = 47518.6701404084
       LHV("C3H8") = 46340.3229417385
       LHV("n_C4H10") = 45725.2111392059
       LHV("i_C4H10") = 45567.2649949597
       LHV("n_C5H12") = 45351.6910185806
       LHV("i_C5H12") = 45255.6703753042
       LHV("n_C6H14") = 45109.0362769398
       LHV("n_C7H16") = 44927.3348363418
       LHV("n_C8H18") = 44789.2869896708
       LHV("He") = 0
       LHV("H2S") = 15198.8830187702
       LHV("SO2") = 0
    Else
    LHV.Add "H2O", 0
    LHV.Add "O2", 0
    LHV.Add "CO2", 0
    LHV.Add "N2", 0
    LHV.Add "Ar", 0
    LHV.Add "H2", 119914.646208418
    LHV.Add "CO", 10100.3201975624
    LHV.Add "CH4", 50033.3833698394
    LHV.Add "C2H6", 47518.6701404084
    LHV.Add "C3H8", 46340.3229417385
    LHV.Add "n_C4H10", 45725.2111392059
    LHV.Add "i_C4H10", 45567.2649949597
    LHV.Add "n_C5H12", 45351.6910185806
    LHV.Add "i_C5H12", 45255.6703753042
    LHV.Add "n_C6H14", 45109.0362769398
    LHV.Add "n_C7H16", 44927.3348363418
    LHV.Add "n_C8H18", 44789.2869896708
    LHV.Add "He", 0
    LHV.Add "H2S", 15198.8830187702
    LHV.Add "SO2", 0
    End If
    
    Fuel.Properties.P = HRSG_input.Cells(4 + 42, "x")
    Fuel.Properties.T = HRSG_input.Cells(5 + 42, "x")
    
    Fuel.Properties.MW = 0
    
    For Each key In w
        Fuel.Properties.MW = Fuel.Properties.MW + Fuel.Composition.Comp(key) * w(key)
    Next key
   
    '물성치 계산
    Property_fuel = idealgas("hsck", "p", Fuel.Properties.P, "t", Fuel.Properties.T, Fuel.Composition.Composition)

    Fuel.Properties.h = Property_fuel(0)
    Fuel.Properties.s = Property_fuel(1)
    Fuel.Properties.C = Property_fuel(2)
    Fuel.Properties.k = Property_fuel(3)

    'LHV 계산
    LHV_fuel = 0
    For Each key In LHV
        LHV_fuel = LHV_fuel + LHV(key) * w(key) * Fuel.Composition.Comp(key)
    Next key

    LHV_fuel = LHV_fuel / Fuel.Properties.MW

    '기체상수
    Fuel.Properties.r = r / Fuel.Properties.MW

    Fuel.Properties.Prop_save
    
    If mode = 0 And IsEmpty(Power_de) Then
        Fuel.Massflow.m = 0
      ElseIf mode = 0 And Not IsEmpty(Power_de) Then
        Fuel.Massflow.m = Power_de / (Thermal_effi_de * LHV_fuel)
      ElseIf mode = 2 Then
        Fuel.Massflow.m = Power_map / (Thermal_effi_map * LHV_fuel)
      End If
    Fuel.Massflow.Mass_save

End Function



