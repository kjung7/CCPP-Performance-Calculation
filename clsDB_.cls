VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDB_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Inlet_h As New clsStream_
Public Outlet_h As New clsStream_
Public Inlet_c As New clsStream_steam_
Public Inlet_c_valve As New clsStream_steam_
Public Outlet_c As New clsStream_steam_

Public Fuel2 As New clsStream_

Public mode As Integer
Public name As Variant

Public Effectiveness, Velocity, UA, Check_Mass As Variant

Public P_loss_hot, delta_P_hot As Variant
Public Calc_option, Input_value As Variant
Public Method_P_loss_hot As Variant

Public LHV, Efficiency As Double
Public Mass_h, Temp_h_in, Press_h_in As Double

Public Stop_cal_value As Double

Sub calDB()

Select Case mode

Case 0

    Dim dummy As Variant
    Dim Property_DB As Variant
    Dim w As New Dictionary
    
    Dim m_fuel As Double
    Dim iter As Integer
    Dim Y, y_for As Double
    
    Dim key As Variant
    Dim composition_gas, composition_gas_2 As Variant
    
    If Calc_option = "No Duct Firing" Then
        Fuel2.Massflow.m = 0
        Set Outlet_h.Massflow = Inlet_h.Massflow
        If P_loss_hot = 0 Then
            Set Outlet_h.Properties = Inlet_h.Properties
        Else
            If Method_P_loss_hot = "P_loss_hot" Then
                Outlet_h.Properties.P = Inlet_h.Properties.P * (1 - P_loss_hot)
                delta_P_hot = Inlet_h.Properties.P - Outlet_h.Properties.P
            ElseIf Method_P_loss_hot = "delta_P_hot" Then
                delta_P_hot = P_loss_hot
                Outlet_h.Properties.P = Inlet_h.Properties.P - P_loss_hot
            End If
            Outlet_h.Properties.T = Inlet_h.Properties.T
            dummy = idealgas("hsck", "p", Outlet_h.Properties.P, "t", Outlet_h.Properties.T, Inlet_h.Composition.Composition)
            
            Outlet_h.Properties.h = dummy(0)
            Outlet_h.Properties.s = dummy(1)
            Outlet_h.Properties.C = dummy(2)
            Outlet_h.Properties.k = dummy(3)
            Outlet_h.Properties.MW = Inlet_h.Properties.MW
            Outlet_h.Properties.r = Inlet_h.Properties.r
            Outlet_h.Properties.Prop_save
        End If
        
        Set Outlet_h.Composition = Inlet_h.Composition
        Mass_h = Inlet_h.Massflow.m
        Temp_h_in = Inlet_h.Properties.T
        Press_h_in = Inlet_h.Properties.P
        Exit Sub
    End If
    
    
  ''HOT side ''''''''''''''''''''''''
Select Case Calc_option
    Case "Burner Exit Temperature", "Temperature Change"
   
   m_fuel = 10 / 36
   
  If Calc_option = "Burner Exit Temperature" Then
        Outlet_h.Properties.T = Input_value
  ElseIf Calc_option = "Temperature Change" Then
        Outlet_h.Properties.T = Inlet_h.Properties.T + Input_value
  End If

  iter = 0
      
  Do
  
  iter = iter + 1

  Fuel2.Massflow.m = m_fuel
  
    '연소가스 조성 계산
'    If Cells(22, "N") = 1 Then
        composition_gas = cal_Composition_gas(Inlet_h.Properties.Properties, Inlet_h.Composition.Composition, Inlet_h.Massflow.m, Fuel2.Properties.Properties, Fuel2.Composition.Composition, Fuel2.Massflow.m)
        Outlet_h.Composition.Composition = composition_gas
        Outlet_h.Composition.Comp_in (Outlet_h.Composition.Composition)
        Outlet_h.Composition.Comp_save
'    ElseIf Cells(22, "N") = 0 Then
'        composition_gas_2 = cal_Composition_gas_2(Inlet_h.Properties.Properties, Inlet_h.Composition.Composition, Inlet_h.Massflow.m, Fuel2.Properties.Properties, Fuel2.Composition.Composition, Fuel2.Massflow.m)
'        Outlet_h.Composition.Composition = composition_gas_2
'        Outlet_h.Composition.Comp_in (Outlet_h.Composition.Composition)
'        Outlet_h.Composition.Comp_save
'    End If
    
    '연소가스 유량 계산
    Outlet_h.Massflow.m = Inlet_h.Massflow.m + Fuel2.Massflow.m
    Outlet_h.Massflow.Mass_save

    '연소가스 엔탈피 계산
    Outlet_h.Properties.h = ((Efficiency * Fuel2.Massflow.m * (Fuel2.Properties.h + LHV)) + (Inlet_h.Massflow.m * Inlet_h.Properties.h)) / Outlet_h.Massflow.m


    '연소가스 압력 계산
    If Method_P_loss_hot = "P_loss_hot" Then
        Outlet_h.Properties.P = Inlet_h.Properties.P * (1 - P_loss_hot)
        delta_P_hot = Inlet_h.Properties.P - Outlet_h.Properties.P
    ElseIf Method_P_loss_hot = "delta_P_hot" Then
        delta_P_hot = P_loss_hot
        Outlet_h.Properties.P = Inlet_h.Properties.P - P_loss_hot
    End If
    
    '연소가스 온도 계산
    dummy = idealgas("t", "p", Outlet_h.Properties.P, "h", Outlet_h.Properties.h, Outlet_h.Composition.Composition)

    Y = Outlet_h.Properties.T - dummy(0)

    If Abs(Y) < 10 ^ (-8) Then
        Exit Do
    ElseIf iter > 100 Then
        MsgBox "반복계산 횟수가 100번 이상이 됨."
        Error 1
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Fuel2.Massflow.m = m_fuel + 10 ^ (-6)
    
    '연소가스 조성 계산
        composition_gas = cal_Composition_gas(Inlet_h.Properties.Properties, Inlet_h.Composition.Composition, Inlet_h.Massflow.m, Fuel2.Properties.Properties, Fuel2.Composition.Composition, Fuel2.Massflow.m)
        Outlet_h.Composition.Composition = composition_gas
        Outlet_h.Composition.Comp_in (Outlet_h.Composition.Composition)
        Outlet_h.Composition.Comp_save
        
    '연소가스 유량 계산
    Outlet_h.Massflow.m = Inlet_h.Massflow.m + Fuel2.Massflow.m
    Outlet_h.Massflow.Mass_save

    '연소가스 엔탈피 계산
    Outlet_h.Properties.h = ((Efficiency * Fuel2.Massflow.m * (Fuel2.Properties.h + LHV)) + (Inlet_h.Massflow.m * Inlet_h.Properties.h)) / Outlet_h.Massflow.m

    '연소가스 압력 계산
    If Method_P_loss_hot = "P_loss_hot" Then
        Outlet_h.Properties.P = Inlet_h.Properties.P * (1 - P_loss_hot)
        delta_P_hot = Inlet_h.Properties.P - Outlet_h.Properties.P
    ElseIf Method_P_loss_hot = "delta_P_hot" Then
        delta_P_hot = P_loss_hot
        Outlet_h.Properties.P = Inlet_h.Properties.P - P_loss_hot
    End If
    
    '연소가스 온도 계산
    dummy = idealgas("t", "p", Outlet_h.Properties.P, "h", Outlet_h.Properties.h, Outlet_h.Composition.Composition)
    
    y_for = Outlet_h.Properties.T - dummy(0)

    m_fuel = m_fuel - (Y / ((y_for - Y) / 10 ^ (-6)))
       
    Loop
    
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case "Specify Fuel Flow"
    
    Fuel2.Massflow.m = Input_value * 10 / 36


    '연소가스 조성 계산
'    MsgBox Inlet_h.Properties.k`
    composition_gas = cal_Composition_gas(Inlet_h.Properties.Properties, Inlet_h.Composition.Composition, Inlet_h.Massflow.m, Fuel2.Properties.Properties, Fuel2.Composition.Composition, Fuel2.Massflow.m)

    Outlet_h.Composition.Composition = composition_gas
    Outlet_h.Composition.Comp_in (Outlet_h.Composition.Composition)
    
    '연소가스 유량 계산
    Outlet_h.Massflow.m = Inlet_h.Massflow.m + Fuel2.Massflow.m
    Outlet_h.Massflow.Mass_save
    '연소가스 엔탈피 계산
    Outlet_h.Properties.h = ((Efficiency * Fuel2.Massflow.m * (Fuel2.Properties.h + LHV)) + (Inlet_h.Massflow.m * Inlet_h.Properties.h)) / Outlet_h.Massflow.m


    '연소가스 압력 계산
    If Method_P_loss_hot = "P_loss_hot" Then
        Outlet_h.Properties.P = Inlet_h.Properties.P * (1 - P_loss_hot)
        delta_P_hot = Inlet_h.Properties.P - Outlet_h.Properties.P
    ElseIf Method_P_loss_hot = "delta_P_hot" Then
        delta_P_hot = P_loss_hot
        Outlet_h.Properties.P = Inlet_h.Properties.P - P_loss_hot
    End If
    
    '연소가스 온도 계산
    dummy = idealgas("t", "p", Outlet_h.Properties.P, "h", Outlet_h.Properties.h, Outlet_h.Composition.Composition)
    Outlet_h.Properties.T = dummy(0)
    
    End Select
    
    '연소가스 물성치 계산
    Property_DB = idealgas("hsck", "p", Outlet_h.Properties.P, "t", Outlet_h.Properties.T, Outlet_h.Composition.Composition)

    Outlet_h.Properties.h = Property_DB(0)
    Outlet_h.Properties.s = Property_DB(1)
    Outlet_h.Properties.C = Property_DB(2)
    Outlet_h.Properties.k = Property_DB(3)
    
  'Molecular weight
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
    w.Add "H2S", 34.08088
    w.Add "He", 4.0026
    w.Add "SO2", 64.064
    
'    For Each key In w
'        Outlet_h.Properties.mw = Outlet_h.Properties.mw + w(key) * Outlet_h.Composition.Comp(key)
'    Next key

    Outlet_h.Properties.MW = Inlet_h.Properties.MW
            
    Outlet_h.Properties.r = 8.31451 / Outlet_h.Properties.MW

    Outlet_h.Properties.Prop_save
    Mass_h = Inlet_h.Massflow.m
    Temp_h_in = Inlet_h.Properties.T
    Press_h_in = Inlet_h.Properties.P
    
Stop_cal_value = Outlet_h.Properties.T ''수렴성

Case 1

 Dim delta_P_hot_off  As Double
 
    If Calc_option = "No Duct Firing" Then
        Fuel2.Massflow.m = 0
        Set Outlet_h.Massflow = Inlet_h.Massflow
        If P_loss_hot = 0 Then
            Set Outlet_h.Properties = Inlet_h.Properties
        Else
            If Method_P_loss_hot = "P_loss_hot" Then
                Outlet_h.Properties.P = Inlet_h.Properties.P * (1 - P_loss_hot)
                delta_P_hot = Inlet_h.Properties.P - Outlet_h.Properties.P
            ElseIf Method_P_loss_hot = "delta_P_hot" Then
                delta_P_hot = P_loss_hot
                Outlet_h.Properties.P = Inlet_h.Properties.P - P_loss_hot
            End If
            Outlet_h.Properties.T = Inlet_h.Properties.T
            dummy = idealgas("hsck", "p", Outlet_h.Properties.P, "t", Outlet_h.Properties.T, Inlet_h.Composition.Composition)
            
            Outlet_h.Properties.h = dummy(0)
            Outlet_h.Properties.s = dummy(1)
            Outlet_h.Properties.C = dummy(2)
            Outlet_h.Properties.k = dummy(3)
            Outlet_h.Properties.MW = Inlet_h.Properties.MW
            Outlet_h.Properties.r = Inlet_h.Properties.r
            Outlet_h.Properties.Prop_save
        End If
        
        Set Outlet_h.Composition = Inlet_h.Composition
        Mass_h = Inlet_h.Massflow.m
        Temp_h_in = Inlet_h.Properties.T
        Press_h_in = Inlet_h.Properties.P
        Exit Sub
    End If
 
  ''HOT side ''''''''''''''''''''''''
Select Case Calc_option

    Case "Burner Exit Temperature", "Temperature Change"
    m_fuel = 10 / 36
    If Calc_option = "Burner Exit Temperature" Then
        Outlet_h.Properties.T = Input_value
    ElseIf Calc_option = "Temperature Change" Then
        Outlet_h.Properties.T = Inlet_h.Properties.T + Input_value
    End If

    iter = 0
    Do
    iter = iter + 1

    Fuel2.Massflow.m = m_fuel
    composition_gas = cal_Composition_gas(Inlet_h.Properties.Properties, Inlet_h.Composition.Composition, Inlet_h.Massflow.m, Fuel2.Properties.Properties, Fuel2.Composition.Composition, Fuel2.Massflow.m)
    Outlet_h.Composition.Composition = composition_gas
    Outlet_h.Composition.Comp_in (Outlet_h.Composition.Composition)
    Outlet_h.Composition.Comp_save
    
    '연소가스 유량 계산
    Outlet_h.Massflow.m = Inlet_h.Massflow.m + Fuel2.Massflow.m
    Outlet_h.Massflow.Mass_save

    '연소가스 엔탈피 계산
    Outlet_h.Properties.h = ((Efficiency * Fuel2.Massflow.m * (Fuel2.Properties.h + LHV)) + (Inlet_h.Massflow.m * Inlet_h.Properties.h)) / Outlet_h.Massflow.m

    '연소가스 압력 계산
    delta_P_hot_off = delta_P_hot * ((Inlet_h.Massflow.m / Mass_h) ^ 1.84) * ((Inlet_h.Properties.T + 273.15) / (Temp_h_in + 273.15)) * (Press_h_in / Inlet_h.Properties.P)  'gas side P_loss_off
    Outlet_h.Properties.P = Inlet_h.Properties.P - delta_P_hot_off
    
    '연소가스 온도 계산
    dummy = idealgas("t", "p", Outlet_h.Properties.P, "h", Outlet_h.Properties.h, Outlet_h.Composition.Composition)

    Y = Outlet_h.Properties.T - dummy(0)

    If Abs(Y) < 10 ^ (-8) Then
        Exit Do
    ElseIf iter > 100 Then
        MsgBox "반복계산 횟수가 100번 이상이 됨."
        Error 1
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Fuel2.Massflow.m = m_fuel + 10 ^ (-6)
    composition_gas = cal_Composition_gas(Inlet_h.Properties.Properties, Inlet_h.Composition.Composition, Inlet_h.Massflow.m, Fuel2.Properties.Properties, Fuel2.Composition.Composition, Fuel2.Massflow.m)
    Outlet_h.Composition.Composition = composition_gas
    Outlet_h.Composition.Comp_in (Outlet_h.Composition.Composition)
    Outlet_h.Composition.Comp_save
    
    '연소가스 유량 계산
    Outlet_h.Massflow.m = Inlet_h.Massflow.m + Fuel2.Massflow.m
    Outlet_h.Massflow.Mass_save

    '연소가스 엔탈피 계산
    Outlet_h.Properties.h = ((Efficiency * Fuel2.Massflow.m * (Fuel2.Properties.h + LHV)) + (Inlet_h.Massflow.m * Inlet_h.Properties.h)) / Outlet_h.Massflow.m

    '연소가스 압력 계산
    delta_P_hot_off = delta_P_hot * ((Inlet_h.Massflow.m / Mass_h) ^ 1.84) * ((Inlet_h.Properties.T + 273.15) / (Temp_h_in + 273.15)) * (Press_h_in / Inlet_h.Properties.P) 'gas side P_loss_off
    Outlet_h.Properties.P = Inlet_h.Properties.P - delta_P_hot_off
    
    '연소가스 온도 계산
    dummy = idealgas("t", "p", Outlet_h.Properties.P, "h", Outlet_h.Properties.h, Outlet_h.Composition.Composition)
    
    y_for = Outlet_h.Properties.T - dummy(0)

    m_fuel = m_fuel - (Y / ((y_for - Y) / 10 ^ (-6)))
       
    Loop

  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case "Specify Fuel Flow"
    
    Fuel2.Massflow.m = Input_value
      
    composition_gas = cal_Composition_gas(Inlet_h.Properties.Properties, Inlet_h.Composition.Composition, Inlet_h.Massflow.m, Fuel2.Properties.Properties, Fuel2.Composition.Composition, Fuel2.Massflow.m)
    Outlet_h.Composition.Composition = composition_gas
    Outlet_h.Composition.Comp_in (Outlet_h.Composition.Composition)
    Outlet_h.Composition.Comp_save
    
    '연소가스 유량 계산
    Outlet_h.Massflow.m = Inlet_h.Massflow.m + Fuel2.Massflow.m

    Outlet_h.Massflow.Mass_save

    '연소가스 엔탈피 계산
    Outlet_h.Properties.h = ((Efficiency * Fuel2.Massflow.m * (Fuel2.Properties.h + LHV)) + (Inlet_h.Massflow.m * Inlet_h.Properties.h)) / Outlet_h.Massflow.m

    '연소가스 압력 계산
    delta_P_hot_off = delta_P_hot * ((Inlet_h.Massflow.m / Mass_h) ^ 1.84) * ((Inlet_h.Properties.T + 273.15) / (Temp_h_in + 273.15)) * (Press_h_in / Inlet_h.Properties.P) 'gas side P_loss_off
    Outlet_h.Properties.P = Inlet_h.Properties.P - delta_P_hot_off
    
    '연소가스 온도 계산
    dummy = idealgas("t", "p", Outlet_h.Properties.P, "h", Outlet_h.Properties.h, Outlet_h.Composition.Composition)
    Outlet_h.Properties.T = dummy(0)

    End Select

    '연소가스 물성치 계산
    Property_DB = idealgas("hsck", "p", Outlet_h.Properties.P, "t", Outlet_h.Properties.T, Outlet_h.Composition.Composition)
    
    Outlet_h.Properties.h = Property_DB(0)
    Outlet_h.Properties.s = Property_DB(1)
    Outlet_h.Properties.C = Property_DB(2)
    Outlet_h.Properties.k = Property_DB(3)
    
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
    
'    For Each key In w
'        Outlet_h.Properties.mw = Outlet_h.Properties.mw + w(key) * Outlet_h.Composition.Comp(key)
'    Next key
    Outlet_h.Properties.MW = Inlet_h.Properties.MW
    
    Outlet_h.Properties.r = 8.31451 / Outlet_h.Properties.MW
    Outlet_h.Properties.Prop_save

Stop_cal_value = Outlet_h.Properties.T ''수렴성

Case Else
    MsgBox "DuctBurner : Case error"
End Select
End Sub

Function cal_Composition_gas(Property_air As Variant, Composition_air As Variant, Massflow_air As Variant, Property_fuel As Variant, Composition_fuel As Variant, Massflow_fuel As Variant) As Variant '연소 후 조성계산
    
    Dim Comp_array(19, 19) As Variant
    Dim section0, section1, section2, section3, section4, section5, section6 As Variant
    Dim section7, section8, section9, section10, section11, section12 As Variant
    Dim section13, section14, section15, section16, section17, section18, section19 As Variant
    Dim section_data As Variant
    Dim k, z As Double

    section0 = Array(1, 0, 0, 0, 0, 1, 0, 2, 3, 4, 5, 5, 6, 6, 7, 8, 9, 0, 1, 0)
    section1 = Array(0, 1, 0, 0, 0, -0.5, -0.5, -2, -3.5, -5, -6.5, -6.5, -8, -8, -9.5, -11, -12.5, 0, -1.5, 0)
    section2 = Array(0, 0, 1, 0, 0, 0, 1, 1, 2, 3, 4, 4, 5, 5, 6, 7, 8, 0, 0, 0)
    section3 = Array(0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    section4 = Array(0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    section5 = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    section6 = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    section7 = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    section8 = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    section9 = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    section10 = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    section11 = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    section12 = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    section13 = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    section14 = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    section15 = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    section16 = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    section17 = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0)
    section18 = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    section19 = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0)

    For k = 0 To 19
        For z = 0 To 19
            If k = 0 Then section_data = section0
            If k = 1 Then section_data = section1
            If k = 2 Then section_data = section2
            If k = 3 Then section_data = section3
            If k = 4 Then section_data = section4
            If k = 5 Then section_data = section5
            If k = 6 Then section_data = section6
            If k = 7 Then section_data = section7
            If k = 8 Then section_data = section8
            If k = 9 Then section_data = section9
            If k = 10 Then section_data = section10
            If k = 11 Then section_data = section11
            If k = 12 Then section_data = section12
            If k = 13 Then section_data = section13
            If k = 14 Then section_data = section14
            If k = 15 Then section_data = section15
            If k = 16 Then section_data = section16
            If k = 17 Then section_data = section17
            If k = 18 Then section_data = section18
            If k = 19 Then section_data = section18
            Comp_array(k, z) = section_data(z)
        Next z
    Next k
    
    Dim total_mol_Air, total_mol_Fuel, total_mol_prod As Double
    Dim n_Air(19), n_Fuel(19) As Variant
    Dim Fuel_coeff(0 To 19, 0 To 19)
    Dim Prod(19) As Double
    Dim n_prod(19) As Double
    Dim i, j, q As Integer
    Dim Comb_outlet_Comp(30) As Variant
    
    '총 몰수 계산
    total_mol_Air = Massflow_air / Property_air(6)
    total_mol_Fuel = Massflow_fuel / Property_fuel(6)

    '각 Composition의 몰수(몰분율*총 몰수)
    For i = 0 To 19
        n_Air(i) = Composition_air(i) * total_mol_Air
        n_Fuel(i) = Composition_fuel(i) * total_mol_Fuel
    Next

    For i = 0 To 19 Step 1
        For j = 0 To 19 Step 1
        Fuel_coeff(i, j) = Comp_array(i, j)
        Next
    Next
    
    For i = 0 To 19 Step 1
        For j = 0 To 19 Step 1
            Prod(i) = Prod(i) + Fuel_coeff(i, j) * n_Fuel(j)
        Next
    Next
    
    For i = 0 To 19 Step 1
        n_prod(i) = Prod(i) + n_Air(i)
    Next
    
    total_mol_prod = 0
    For i = 0 To 19 Step 1
        total_mol_prod = total_mol_prod + n_prod(i)
    Next
    
    For i = 0 To 19 Step 1
        Comb_outlet_Comp(i) = n_prod(i) / total_mol_prod
    Next

    cal_Composition_gas = Comb_outlet_Comp

End Function

Function cal_Composition_gas_2(Property_air As Variant, Composition_air As Variant, Massflow_air As Variant, Property_fuel As Variant, Composition_fuel As Variant, Massflow_fuel As Variant) As Variant '연소 후 조성계산
    Dim total_mol_Air, total_mol_Fuel, total_mol_prod As Double
    Dim n_Air(18), n_Fuel(4) As Variant
    Dim Fuel_coeff(0 To 18, 0 To 18)
    Dim Prod(18) As Double
    Dim n_prod(18) As Double
    Dim i, j, q As Integer
    Dim Comb_outlet_Comp(30) As Variant
    
    Worksheets("Combustion_2").Activate
        
    '총 몰수 계산
    total_mol_Air = Massflow_air / Property_air(6)
    total_mol_Fuel = Massflow_fuel / Property_fuel(6)
    
    '각 Composition의 몰수(몰분율*총 몰수)
    For i = 0 To 18
        n_Air(i) = Composition_air(i) * total_mol_Air
    Next
    
    For i = 0 To 4
        n_Fuel(i) = Composition_fuel(i) * total_mol_Fuel
    Next
    
    For i = 0 To 18 Step 1
        For j = 0 To 4 Step 1
        Fuel_coeff(i, j) = Cells(i + 2, j + 2)
        Next
    Next
    
    For i = 0 To 18 Step 1
        For j = 0 To 4 Step 1
            Prod(i) = Prod(i) + Fuel_coeff(i, j) * n_Fuel(j)
        Next
    Next
    
    For i = 0 To 18 Step 1
        n_prod(i) = Prod(i) + n_Air(i)
    Next
    
    total_mol_prod = 0
    For i = 0 To 18 Step 1
        total_mol_prod = total_mol_prod + n_prod(i)
    Next
    
    For i = 0 To 18 Step 1
        Comb_outlet_Comp(i) = n_prod(i) / total_mol_prod
    Next

    cal_Composition_gas_2 = Comb_outlet_Comp
    
    'Worksheets("Input").Activate
    
End Function






