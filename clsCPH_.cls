VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCPH_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Inlet_h As New clsStream_
Public Outlet_h As New clsStream_
Public Cir As New clsStream_steam_ 'Circulaiton flow
Public Inlet_c As New clsStream_steam_
Public Inlet_c_valve As New clsStream_steam_
Public Inlet_c_cir As New clsStream_steam_
Public Outlet_c_cir As New clsStream_steam_
Public Outlet_c As New clsStream_steam_

'HRSG_read'''''''''''''''''
Public mode, j As Integer
Public name As Variant

Public delta_T_bypass, Gas_ratio As Double 'bypass인 경우 hotside temperature drop/

Public E_loss, P_loss_hot, P_loss_valve, P_loss_cold As Double
Public Method_P_loss_hot, Method_P_loss_valve, Method_P_loss_cold As Variant
Public Calc_option, Input_value As Variant
Public D_o, T_tube, Material, L_tube, n_tube, x_tube As Variant
Public P_EVAP As Double
''''''''''''''''''''''''''''
Public CPH_mode As Variant
Public Temp_in_c As Double

Public delta_P_hot, delta_P_valve, delta_P_cold As Variant
Public delta_P_valve_off, delta_P_cold_off As Double

Private Q_c, Q_h As Double
Private Cp_c, Cp_h, C_c, C_h, C_min, C_r As Double
Public Effectiveness As Double

Public D_i, rho_c, Velocity, A_o As Double 'Steam velocity & A_o
Public f_fin, Fin_efficiency As Double
Public NTU, Effectiveness_NTU, UA, U As Double 'U
Public Press_c, Temp_c, Press_SI, Temp_SI As Double
Public Temp_tube, k_tube As Double
Public Temp_h As Double
Public my_c, tc_c, Re, PR, R_i As Double 'R_i
Public R_w As Double  'R_w
Public R_o As Double 'R_o

Public Stop_cal_value As Double

Public h_c_assume As Double
Public Mass_h, Press_h_in, Temp_h_in As Double
Public Mass_in_c, Mass_in_c_cir As Double

Public my_h, Heat_capa_h, tc_h As Double
Public Heat_capa_c As Double

Public R_i_off, R_w_off, R_o_off, U_off, UA_off As Double 'off thermal resistance
Public Check_Mass As Variant 'EVAP Mass 계산 완료
Public BFP_Cal_option As Variant

Sub calCPH()

Select Case mode

Case 0

    Dim dummy As Variant
    Dim del_T_2, del_T_1, LMTD As Double
    
'     Cold side inlet의 stream 이 연결되지 않을 경우 hot side stream bypass
'     동일 pressure level의 ECON 계산 완료 이후에 stream이 연결됨
    If IsEmpty(Inlet_c.Properties.h) Then
        
        Outlet_h.Properties.T = Inlet_h.Properties.T - delta_T_bypass
        Outlet_h.Properties.P = Inlet_h.Properties.P
        dummy = idealgas("hsck", "p", Outlet_h.Properties.P, "t", Outlet_h.Properties.T, Inlet_h.Composition.Composition)
        Outlet_h.Properties.h = dummy(0)
        Outlet_h.Properties.s = dummy(1)
        Outlet_h.Properties.C = dummy(2)
        Outlet_h.Properties.k = dummy(3)
        Outlet_h.Properties.MW = Inlet_h.Properties.MW
        Outlet_h.Properties.r = Inlet_h.Properties.r

        Outlet_h.Properties.Prop_save

        Set Outlet_h.Massflow = Inlet_h.Massflow
        Set Outlet_h.Composition = Inlet_h.Composition
        Exit Sub
        
    End If
    
    f_fin = 10
    Fin_efficiency = 1
    
    If Not Gas_ratio = 0 Then
        Inlet_h.Massflow.m = Inlet_h.Massflow.m * Gas_ratio
    End If
        
    If Method_P_loss_valve = "P_loss_valve" Then
        Inlet_c_valve.Properties.P = Inlet_c.Properties.P * (1 - P_loss_valve)
        delta_P_valve = Inlet_c.Properties.P - Inlet_c_valve.Properties.P
    ElseIf Method_P_loss_valve = "delta_P_valve" Then
        delta_P_valve = P_loss_valve
        Inlet_c_valve.Properties.P = Inlet_c.Properties.P - P_loss_valve
    End If
    Inlet_c_valve.Properties.h = Inlet_c.Properties.h
    Inlet_c_valve.Properties.T = T_ph(Inlet_c.Properties.P, Inlet_c.Properties.h)
    
    Set Outlet_h.Massflow = Inlet_h.Massflow
    Set Outlet_h.Composition = Inlet_h.Composition
    Set Outlet_c.Massflow = Inlet_c.Massflow

    If CPH_mode = "Bypass" Then 'CPH bypass
        Set Outlet_c.Properties = Inlet_c.Properties
        Set Outlet_c.Massflow = Inlet_c.Massflow
        Set Outlet_h.Properties = Inlet_h.Properties
        Exit Sub
    End If
    
    'inlet_c_cir properties
    If Temp_in_c = 0 Then 'circulation X
        Set Inlet_c_cir.Properties = Inlet_c_valve.Properties
        Inlet_c_cir.Properties.Prop_save
    Else 'circulation
        Inlet_c_cir.Properties.P = Inlet_c_valve.Properties.P
        Inlet_c_cir.Properties.T = Temp_in_c
        Inlet_c_cir.Properties.h = h_pT(Inlet_c_cir.Properties.P, Inlet_c_cir.Properties.T)
        Inlet_c_cir.Properties.s = s_pT(Inlet_c_cir.Properties.P, Inlet_c_cir.Properties.T)
        Inlet_c_cir.Properties.cp = Cp_pT(Inlet_c_cir.Properties.P, Inlet_c_cir.Properties.T)
        Inlet_c_cir.Properties.Cv = Cv_pT(Inlet_c_cir.Properties.P, Inlet_c_cir.Properties.T)
        Inlet_c_cir.Properties.k = Inlet_c_cir.Properties.cp / Inlet_c_cir.Properties.Cv
        Inlet_c_cir.Properties.x = x_ph(Inlet_c_cir.Properties.P, Inlet_c_cir.Properties.h)
        Inlet_c_cir.Properties.Prop_save
    End If
    
    'outlet_c_cirproperties -> cir properties & massflow
    If Method_P_loss_cold = "P_loss_cold" Then
        Outlet_c_cir.Properties.P = Inlet_c_cir.Properties.P * (1 - P_loss_cold)
        delta_P_cold = Inlet_c_cir.Properties.P - Outlet_c_cir.Properties.P
    ElseIf Method_P_loss_cold = "delta_P_cold" Then
        delta_P_cold = P_loss_cold
        Outlet_c_cir.Properties.P = Inlet_c_cir.Properties.P - P_loss_cold
    End If
    
    If Calc_option = "Water outlet enthalpy" Then
        Outlet_c_cir.Properties.h = Input_value
        Outlet_c_cir.Properties.T = T_ph(Outlet_c_cir.Properties.P, Outlet_c_cir.Properties.h)
    ElseIf Calc_option = "Water outlet temperature" Then 'User input : water outlet temperature
        Outlet_c_cir.Properties.T = Input_value
        Outlet_c_cir.Properties.h = h_pT(Outlet_c_cir.Properties.P, Outlet_c_cir.Properties.T)
    ElseIf Calc_option = "Approach temperature (Gas in - water out)" Then 'User input : Approach temperature
        Outlet_c_cir.Properties.T = Inlet_h.Properties.T - Input_value
        Outlet_c_cir.Properties.h = h_pT(Outlet_c_cir.Properties.P, Outlet_c_cir.Properties.T)
    ElseIf Calc_option = "Exit subcooling" Then
        Outlet_c_cir.Properties.T = Tsat_p(Outlet_c_cir.Properties.P) - Input_value
        If Input_value = 0 Then
            Outlet_c_cir.Properties.h = hL_p(Outlet_c_cir.Properties.P)
        Else
            Outlet_c_cir.Properties.h = h_pT(Outlet_c_cir.Properties.P, Outlet_c_cir.Properties.T)
        End If
    End If
    Outlet_c_cir.Properties.s = s_ph(Outlet_c_cir.Properties.P, Outlet_c_cir.Properties.h)
'    Outlet_c_cir.Properties.Cp = Cp_pT(Outlet_c_cir.Properties.P, Outlet_c_cir.Properties.T)
'    Outlet_c_cir.Properties.Cv = Cv_pT(Outlet_c_cir.Properties.P, Outlet_c_cir.Properties.T)
'    Outlet_c_cir.Properties.k = Outlet_c_cir.Properties.Cp / Outlet_c_cir.Properties.Cv
'    Outlet_c_cir.Properties.x = x_ph(Outlet_c_cir.Properties.P, Outlet_c_cir.Properties.h)
    Outlet_c_cir.Properties.Prop_save
    Press_c = (Inlet_c_cir.Properties.P + Outlet_c_cir.Properties.P) / 2
    Temp_c = (Inlet_c_cir.Properties.T + Outlet_c_cir.Properties.T) / 2
        
    If Temp_in_c = 0 Then 'circulation X
        Set Inlet_c_cir.Massflow = Inlet_c.Massflow
    Else 'circulation
        Set Cir.Properties = Outlet_c_cir.Properties
        Cir.Massflow.m = Inlet_c.Massflow.m * (Inlet_c_cir.Properties.h - Inlet_c_valve.Properties.h) / (Cir.Properties.h - Inlet_c_cir.Properties.h)
        Inlet_c_cir.Massflow.m = Inlet_c.Massflow.m + Cir.Massflow.m
        Inlet_c_cir.Massflow.Mass_save
    End If
        Set Outlet_c_cir.Massflow = Inlet_c_cir.Massflow
        
    'Heat exchange start''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Q_c = Outlet_c_cir.Massflow.m * (Outlet_c_cir.Properties.h - Inlet_c_cir.Properties.h)
    Q_h = Q_c / (1 - E_loss)

    ' Hot side outlet property calculation
    If Method_P_loss_hot = "P_loss_hot" Then
        Outlet_h.Properties.P = Inlet_h.Properties.P * (1 - P_loss_hot)
        delta_P_hot = Inlet_h.Properties.P - Outlet_h.Properties.P
    ElseIf Method_P_loss_hot = "delta_P_hot" Then
        delta_P_hot = P_loss_hot
        Outlet_h.Properties.P = Inlet_h.Properties.P - P_loss_hot
    End If
    Outlet_h.Properties.h = Inlet_h.Properties.h - (Q_h / Outlet_h.Massflow.m)
    dummy = idealgas("t", "p", Outlet_h.Properties.P, "h", Outlet_h.Properties.h, Inlet_h.Composition.Composition)
    Outlet_h.Properties.T = dummy(0)
    dummy = idealgas("sck", "p", Outlet_h.Properties.P, "t", Outlet_h.Properties.T, Inlet_h.Composition.Composition)
    Outlet_h.Properties.s = dummy(0)
    Outlet_h.Properties.C = dummy(1)
    Outlet_h.Properties.k = dummy(2)
    Outlet_h.Properties.MW = Inlet_h.Properties.MW
    Outlet_h.Properties.r = Inlet_h.Properties.r
    Outlet_h.Properties.Prop_save
    Temp_h = (Inlet_h.Properties.T + Outlet_h.Properties.T) / 2
    'LMTD 구하기''''''''''''''''''''''''''''''''''''''''''''''''''
    del_T_2 = Inlet_h.Properties.T - Outlet_c_cir.Properties.T
    del_T_1 = Outlet_h.Properties.T - Inlet_c_cir.Properties.T

    
    If del_T_2 < 0 Or del_T_1 < 0 Then
    MsgBox "Error - " & name & " : Temperature cross"
    Error 1
    End If
    
    LMTD = (del_T_2 - del_T_1) / Application.WorksheetFunction.Ln(del_T_2 / del_T_1)
    
    'Effectiveness'''''''''''''''''''''''''''''''''''''''''''''''''''''
    Cp_c = (Outlet_c_cir.Properties.h - Inlet_c_cir.Properties.h) / (Outlet_c_cir.Properties.T - Inlet_c_cir.Properties.T)
    C_c = Cp_c * Outlet_c_cir.Massflow.m
    Cp_h = (Inlet_h.Properties.h - Outlet_h.Properties.h) / (Inlet_h.Properties.T - Outlet_h.Properties.T)
    C_h = Cp_h * Outlet_h.Massflow.m
    If C_c < C_h Then
        C_r = C_c / C_h
        C_min = C_c
        Effectiveness = (Outlet_c_cir.Properties.T - Inlet_c_cir.Properties.T) / (Inlet_h.Properties.T - Inlet_c_cir.Properties.T)
    ElseIf C_c > C_h Then
        C_r = C_h / C_c
        C_min = C_h
        Effectiveness = (Inlet_h.Properties.T - Outlet_h.Properties.T) / (Inlet_h.Properties.T - Inlet_c_cir.Properties.T)
    End If
    
    If D_o = 0 Then
        NTU = Application.WorksheetFunction.Ln((Effectiveness - 1) / (Effectiveness * C_r - 1)) / (C_r - 1)
        UA = NTU * C_min
    Else
        'Steam velocity & A_o'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        D_i = D_o - 2 * T_tube
        rho_c = rho_pT(Press_c, Temp_c)
        Velocity = 4 * Inlet_c_cir.Massflow.m / (n_tube * rho_c * Application.WorksheetFunction.pi * D_i ^ 2)
        A_o = f_fin * Application.WorksheetFunction.pi * D_o * n_tube * L_tube * x_tube
    
        'U''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        NTU = Application.WorksheetFunction.Ln((Effectiveness - 1) / (Effectiveness * C_r - 1)) / (C_r - 1)
        UA = NTU * C_min
        U = UA / A_o
    
        'R_i 구하기'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        my_c = my_pT(Press_c, Temp_c)
        tc_c = tc_pT(Press_c, Temp_c) * 10 ^ -3
        Re = (rho_c * Velocity * D_i) / my_c
        PR = (Cp_c * my_c) / tc_c
        R_i = D_o / (0.023 * Re ^ 0.8 * PR ^ 0.4 * tc_c)
        
        'R_w 구하기'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Temp_tube = (Temp_c + Temp_h) / 2
        If Material = "CS" Then
            k_tube = 46.73 * 0.001 + (Temp_tube - 260) * -0.0249 * 0.001 'kW/m-C basis
        ElseIf Material = "SS" Then
            k_tube = 18.34 * 0.001 + (Temp_tube - 260) * 0.015 * 0.001 'kW/m-C basis
        ElseIf Material = "T91" Then
            k_tube = 27 * 0.001 + (Temp_tube - 260) * 0.0053 * 0.001 'kW/m-C basis
        ElseIf Material = "T11" Then
            k_tube = 37.55 * 0.001 + (Temp_tube - 260) * -0.0137 * 0.001 'kW/m-C basis
        ElseIf Material = "T22" Then
            k_tube = 36.86 * 0.001 + (Temp_tube - 260) * -0.0109 * 0.001 'kW/m-C basis
        End If
        R_w = Application.WorksheetFunction.Ln(D_o / D_i) * D_o / (2 * k_tube)
        
        'R_o 구하기'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        R_o = (1 / U - f_fin * R_w - f_fin * R_i) * Fin_efficiency
    End If

    'outlet_c properties and massflow
    Set Outlet_c.Properties = Outlet_c_cir.Properties
    Outlet_c.Massflow.m = Outlet_c_cir.Massflow.m - Cir.Massflow.m
    

    h_c_assume = Outlet_c_cir.Properties.h

    'Off-design의 reference'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Mass_h = Inlet_h.Massflow.m
    Press_h_in = Inlet_h.Properties.P
    Temp_h_in = Inlet_h.Properties.T
    
    Mass_in_c = Inlet_c.Massflow.m
    Mass_in_c_cir = Inlet_c_cir.Massflow.m
    
    my_h = mu_t_gas(Temp_h, Inlet_h.Composition.Composition)
    Heat_capa_h = Cp_h
    tc_h = tc_t_gas(Temp_h, Inlet_h.Composition.Composition)
    
    Heat_capa_c = Cp_c

    
    Stop_cal_value = Outlet_h.Properties.T
'OFF-DESIGN''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Case 1
    
    Dim P_in_cal, delta_P_hot_off As Double
    Dim my_c_off, tc_c_off, my_h_off, tc_h_off As Double
    Dim iter As Integer
    Dim Y, y_for As Double
        Dim step_size As Double
    Dim Toler_Adj As Double
    
        
    Toler_Adj = 10 ^ (-8)

    
    If Not Gas_ratio = 0 Then
    Inlet_h.Massflow.m = Inlet_h.Massflow.m * Gas_ratio
    End If
    Set Outlet_h.Massflow = Inlet_h.Massflow
    Set Outlet_h.Composition = Inlet_h.Composition
    
    If CPH_mode = "Bypass" Then 'CPH bypass
        Set Outlet_c.Properties = Inlet_c.Properties
        Set Outlet_c.Massflow = Inlet_c.Massflow
        Set Outlet_h.Properties = Inlet_h.Properties
        Exit Sub
    End If
    
    
     'Cold side 입구 압력 계산 Do~Loop''''''''''''''''
'    Do
'    If BFP_Cal_option = "Curve Method" Then Exit Do
'    If delta_P_valve = 0 Then
'        delta_P_valve_off = 0
'    Else
'        delta_P_valve_off = delta_P_valve * ((Inlet_c.Massflow.m / Mass_in_c) ^ 1.98)
'    End If
'    Inlet_c_valve.Properties.P = Inlet_c.Properties.P - delta_P_valve_off
'    Inlet_c_valve.Properties.h = Inlet_c.Properties.h
'    Inlet_c_valve.Properties.T = T_ph(Inlet_c_valve.Properties.P, Inlet_c_valve.Properties.h)
'
'    If delta_P_cold = 0 Then
'        delta_P_cold_off = 0
'    Else
'        delta_P_cold_off = delta_P_cold * ((Inlet_c_cir.Massflow.m / Mass_in_c_cir) ^ 1.98)
'    End If
'    P_in_cal = Inlet_c.Properties.P
'    Inlet_c.Properties.P = Outlet_c.Properties.P + delta_P_valve_off + delta_P_cold_off
'
'    If Abs(P_in_cal - Inlet_c.Properties.P) < 0.1 Then Exit Do
'    Inlet_c.Properties.T = T_ph(Inlet_c.Properties.P, Inlet_c.Properties.h)
'    Loop
    ''''''''''''''''''''''''''''''''''''''''''''''
    
    If BFP_Cal_option = "Press Cal" Then
    
        If delta_P_valve = 0 Then
            delta_P_valve_off = 0
        Else
            delta_P_valve_off = delta_P_valve * ((Inlet_c.Massflow.m / Mass_in_c) ^ 1.98)
        End If
        If delta_P_cold = 0 Then
            delta_P_cold_off = 0
        Else
            delta_P_cold_off = delta_P_cold * ((Inlet_c_cir.Massflow.m / Mass_in_c_cir) ^ 1.98)
        End If
        Inlet_c.Properties.P = Outlet_c.Properties.P + delta_P_valve_off + delta_P_cold_off
        Inlet_c.Properties.T = T_ph(Inlet_c.Properties.P, Inlet_c.Properties.h)
        
    End If
    
    delta_P_valve_off = delta_P_valve * (Inlet_c.Massflow.m / Mass_in_c) ^ 1.98
    Inlet_c_valve.Properties.P = Inlet_c.Properties.P - delta_P_valve_off
    Inlet_c_valve.Properties.h = Inlet_c.Properties.h
    Inlet_c_valve.Properties.T = T_ph(Inlet_c_valve.Properties.P, Inlet_c_valve.Properties.h)
    
    'coldside inlet properties
    If Temp_in_c = 0 Then 'circulation X
        Set Inlet_c_cir.Properties = Inlet_c_valve.Properties
        Inlet_c_cir.Properties.Prop_save
    Else 'circulation
        Inlet_c_cir.Properties.P = Inlet_c_valve.Properties.P
        Inlet_c_cir.Properties.T = Temp_in_c
        Inlet_c_cir.Properties.h = h_pT(Inlet_c_cir.Properties.P, Inlet_c_cir.Properties.T)
        Inlet_c_cir.Properties.s = s_pT(Inlet_c_cir.Properties.P, Inlet_c_cir.Properties.T)
        Inlet_c_cir.Properties.cp = Cp_pT(Inlet_c_cir.Properties.P, Inlet_c_cir.Properties.T)
        Inlet_c_cir.Properties.Cv = Cv_pT(Inlet_c_cir.Properties.P, Inlet_c_cir.Properties.T)
        Inlet_c_cir.Properties.k = Inlet_c_cir.Properties.cp / Inlet_c_cir.Properties.Cv
        Inlet_c_cir.Properties.x = x_ph(Inlet_c_cir.Properties.P, Inlet_c_cir.Properties.h)
        Inlet_c_cir.Properties.Prop_save
    End If
    
    iter = 0
    Do
    iter = iter + 1
    
    Outlet_c_cir.Properties.h = h_c_assume
    
    'cir enthalpy & Massflow
    If Temp_in_c = 0 Then
        Cir.Massflow.m = 0
    Else
        Cir.Properties.h = Outlet_c_cir.Properties.h
        Cir.Massflow.m = Inlet_c.Massflow.m * (Inlet_c_cir.Properties.h - Inlet_c_valve.Properties.h) / (Cir.Properties.h - Inlet_c_cir.Properties.h)
    End If
    Inlet_c_cir.Massflow.m = Inlet_c.Massflow.m + Cir.Massflow.m
    Inlet_c_cir.Massflow.Mass_save
    
    ' Cold side outlet property calculation
    delta_P_cold_off = delta_P_cold * ((Inlet_c_cir.Massflow.m / Mass_in_c_cir) ^ 1.98)
    Outlet_c_cir.Properties.P = Inlet_c_cir.Properties.P - delta_P_cold_off
    Outlet_c_cir.Properties.T = T_ph(Outlet_c_cir.Properties.P, Outlet_c_cir.Properties.h)
    Outlet_c_cir.Properties.s = s_ph(Outlet_c_cir.Properties.P, Outlet_c_cir.Properties.h)
'    Outlet_c_cir.Properties.Cp = Cp_pT(Outlet_c_cir.Properties.P, Outlet_c_cir.Properties.T)
'    Outlet_c_cir.Properties.Cv = Cv_pT(Outlet_c_cir.Properties.P, Outlet_c_cir.Properties.T)
'    Outlet_c_cir.Properties.k = Outlet_c_cir.Properties.Cp / Outlet_c_cir.Properties.Cv
'    Outlet_c_cir.Properties.x = x_ph(Outlet_c_cir.Properties.P, Outlet_c_cir.Properties.h)
    Outlet_c_cir.Properties.Prop_save
    Press_c = (Inlet_c_cir.Properties.P + Outlet_c_cir.Properties.P) / 2
    Temp_c = (Inlet_c_cir.Properties.T + Outlet_c_cir.Properties.T) / 2

    Set Outlet_c_cir.Massflow = Inlet_c_cir.Massflow
    
    'Heat exchange start''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Q_c = Outlet_c_cir.Massflow.m * (Outlet_c_cir.Properties.h - Inlet_c_cir.Properties.h)
    Q_h = Q_c / (1 - E_loss)
    ' Hot side outlet property calculation
    delta_P_hot_off = delta_P_hot * ((Inlet_h.Massflow.m / Mass_h) ^ 1.84) * ((Inlet_h.Properties.T + 273.15) / (Temp_h_in + 273.15)) * (Press_h_in / Inlet_h.Properties.P) 'gas side P_loss_off
    Outlet_h.Properties.P = Inlet_h.Properties.P - delta_P_hot_off
    Outlet_h.Properties.h = Inlet_h.Properties.h - (Q_h / Outlet_h.Massflow.m)
    dummy = idealgas("t", "p", Outlet_h.Properties.P, "h", Outlet_h.Properties.h, Inlet_h.Composition.Composition)
    Outlet_h.Properties.T = dummy(0)
    dummy = idealgas("sck", "p", Outlet_h.Properties.P, "t", Outlet_h.Properties.T, Inlet_h.Composition.Composition)
    Outlet_h.Properties.s = dummy(0)
    Outlet_h.Properties.C = dummy(1)
    Outlet_h.Properties.k = dummy(2)
    Outlet_h.Properties.MW = Inlet_h.Properties.MW
    Outlet_h.Properties.r = Inlet_h.Properties.r
    Outlet_h.Properties.Prop_save
    Temp_h = (Inlet_h.Properties.T + Outlet_h.Properties.T) / 2
    
     'C_cold & C_hot
    Cp_c = (Outlet_c_cir.Properties.h - Inlet_c_cir.Properties.h) / (Outlet_c_cir.Properties.T - Inlet_c_cir.Properties.T)
    C_c = Cp_c * Outlet_c_cir.Massflow.m
    Cp_h = (Inlet_h.Properties.h - Outlet_h.Properties.h) / (Inlet_h.Properties.T - Outlet_h.Properties.T)
    C_h = Cp_h * Outlet_h.Massflow.m

    If C_c < C_h Then
        C_r = C_c / C_h
        C_min = C_c
        Effectiveness = (Outlet_c_cir.Properties.T - Inlet_c_cir.Properties.T) / (Inlet_h.Properties.T - Inlet_c_cir.Properties.T)
    ElseIf C_c > C_h Then
        C_r = C_h / C_c
        C_min = C_h
        Effectiveness = (Inlet_h.Properties.T - Outlet_h.Properties.T) / (Inlet_h.Properties.T - Inlet_c_cir.Properties.T)
    End If
    
    If D_o = 0 Then
        tc_h_off = tc_t_gas(Temp_h, Inlet_h.Composition.Composition)
        my_h_off = mu_t_gas(Temp_h, Inlet_h.Composition.Composition)
        UA_off = UA * (Inlet_h.Massflow.m / Mass_h) ^ 0.65 * (Cp_h / Heat_capa_h) ^ 0.33 * (tc_h_off / tc_h) ^ 0.67 * (my_h_off / my_h) ^ -0.32
        NTU = UA_off / C_min
    Else
        'R_i_off 구하기''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        my_c_off = my_pT(Press_c, Temp_c)
        tc_c_off = tc_pT(Press_c, Temp_c) * 10 ^ -3
        R_i_off = R_i * (Mass_in_c / Inlet_c_cir.Massflow.m) ^ 0.8 * (my_c / my_c_off) ^ 0.4 * (Heat_capa_c / Cp_c) ^ 0.4 * (tc_c / tc_c_off) ^ 0.6
        
        'R_w_off 구하기''''''''''''''''''''''''''''''''''''''''
        Temp_tube = (Temp_c + Temp_h) / 2
        If Material = "CS" Then
            k_tube = 46.73 * 0.001 + (Temp_tube - 260) * -0.0249 * 0.001 'kW/m-C basis
        ElseIf Material = "SS" Then
            k_tube = 18.34 * 0.001 + (Temp_tube - 260) * 0.015 * 0.001 'kW/m-C basis
        ElseIf Material = "T91" Then
            k_tube = 27 * 0.001 + (Temp_tube - 260) * 0.0053 * 0.001 'kW/m-C basis
        ElseIf Material = "T11" Then
            k_tube = 37.55 * 0.001 + (Temp_tube - 260) * -0.0137 * 0.001 'kW/m-C basis
        ElseIf Material = "T22" Then
            k_tube = 36.86 * 0.001 + (Temp_tube - 260) * -0.0109 * 0.001 'kW/m-C basis
        End If
        R_w_off = Application.WorksheetFunction.Ln(D_o / D_i) * D_o / (2 * k_tube)
        
        'R_o_off 구하기''''''''''''''''''''''''''''''''''''''''''''''
        my_h_off = mu_t_gas(Temp_h, Inlet_h.Composition.Composition)
        tc_h_off = tc_t_gas(Temp_h, Inlet_h.Composition.Composition)
        R_o_off = R_o * (Mass_h / Inlet_h.Massflow.m) * (Heat_capa_h / Cp_h) ^ 0.33 * (my_h / my_h_off) ^ -0.67 * (tc_h / tc_h_off) ^ 0.67
        
        'U_off 구하기 '''''''''''''''''''''''''''''''''
        U_off = 1 / (R_o_off / Fin_efficiency + f_fin * R_w_off + f_fin * R_i_off)
        NTU = U_off * A_o / C_min
    End If
    
    Y = Effectiveness - (1 - Exp(-NTU * (1 - C_r))) / (1 - C_r * Exp(-NTU * (1 - C_r)))

    If Abs(Y) < Toler_Adj Then
       Exit Do
    ElseIf iter > 100 And Abs(Y) > 0.005 Then
       MsgBox "반복계산 횟수가 100번 이상이 됨."
       Error 1
    ElseIf iter > 100 And Abs(Y) <= 0.005 Then
      Toler_Adj = 0.0051
       
    End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Cold side Q 계산
    Outlet_c.Properties.h = h_c_assume + 10 ^ (-6)
    
    'cir enthalpy & Massflow
    If Temp_in_c = 0 Then
        Cir.Massflow.m = 0
    Else
        Cir.Properties.h = Outlet_c_cir.Properties.h
        Cir.Massflow.m = Inlet_c.Massflow.m * (Inlet_c_cir.Properties.h - Inlet_c_valve.Properties.h) / (Cir.Properties.h - Inlet_c_cir.Properties.h)
    End If
    Inlet_c_cir.Massflow.m = Inlet_c.Massflow.m + Cir.Massflow.m
    Inlet_c_cir.Massflow.Mass_save
    
    ' Cold side outlet property calculation
    delta_P_cold_off = delta_P_cold * ((Inlet_c_cir.Massflow.m / Mass_in_c_cir) ^ 1.98)
    Outlet_c_cir.Properties.P = Inlet_c_cir.Properties.P - delta_P_cold_off
    Outlet_c_cir.Properties.T = T_ph(Outlet_c_cir.Properties.P, Outlet_c_cir.Properties.h)
    Outlet_c_cir.Properties.s = s_ph(Outlet_c_cir.Properties.P, Outlet_c_cir.Properties.h)
'    Outlet_c_cir.Properties.Cp = Cp_pT(Outlet_c_cir.Properties.P, Outlet_c_cir.Properties.T)
'    Outlet_c_cir.Properties.Cv = Cv_pT(Outlet_c_cir.Properties.P, Outlet_c_cir.Properties.T)
'    Outlet_c_cir.Properties.k = Outlet_c_cir.Properties.Cp / Outlet_c_cir.Properties.Cv
'    Outlet_c_cir.Properties.x = x_ph(Outlet_c_cir.Properties.P, Outlet_c_cir.Properties.h)
    Outlet_c_cir.Properties.Prop_save
    Press_c = (Inlet_c_cir.Properties.P + Outlet_c_cir.Properties.P) / 2
    Temp_c = (Inlet_c_cir.Properties.T + Outlet_c_cir.Properties.T) / 2

    Set Outlet_c_cir.Massflow = Inlet_c_cir.Massflow
    
    'Heat exchange start''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Q_c = Outlet_c_cir.Massflow.m * (Outlet_c_cir.Properties.h - Inlet_c_cir.Properties.h)
    Q_h = Q_c / (1 - E_loss)
    ' Hot side outlet property calculation
    delta_P_hot_off = delta_P_hot * ((Inlet_h.Massflow.m / Mass_h) ^ 1.84) * ((Inlet_h.Properties.T + 273.15) / (Temp_h_in + 273.15)) * (Press_h_in / Inlet_h.Properties.P) 'gas side P_loss_off
    Outlet_h.Properties.P = Inlet_h.Properties.P - delta_P_hot_off
    Outlet_h.Properties.h = Inlet_h.Properties.h - (Q_h / Outlet_h.Massflow.m)
    dummy = idealgas("t", "p", Outlet_h.Properties.P, "h", Outlet_h.Properties.h, Inlet_h.Composition.Composition)
    Outlet_h.Properties.T = dummy(0)
    dummy = idealgas("sck", "p", Outlet_h.Properties.P, "t", Outlet_h.Properties.T, Inlet_h.Composition.Composition)
    Outlet_h.Properties.s = dummy(0)
    Outlet_h.Properties.C = dummy(1)
    Outlet_h.Properties.k = dummy(2)
    Outlet_h.Properties.MW = Inlet_h.Properties.MW
    Outlet_h.Properties.r = Inlet_h.Properties.r
    Outlet_h.Properties.Prop_save
    Temp_h = (Inlet_h.Properties.T + Outlet_h.Properties.T) / 2
    
     'C_cold & C_hot
    Cp_c = (Outlet_c_cir.Properties.h - Inlet_c_cir.Properties.h) / (Outlet_c_cir.Properties.T - Inlet_c_cir.Properties.T)
    C_c = Cp_c * Outlet_c_cir.Massflow.m
    Cp_h = (Inlet_h.Properties.h - Outlet_h.Properties.h) / (Inlet_h.Properties.T - Outlet_h.Properties.T)
    C_h = Cp_h * Outlet_h.Massflow.m

    If C_c < C_h Then
        C_r = C_c / C_h
        C_min = C_c
        Effectiveness = (Outlet_c_cir.Properties.T - Inlet_c_cir.Properties.T) / (Inlet_h.Properties.T - Inlet_c_cir.Properties.T)
    ElseIf C_c > C_h Then
        C_r = C_h / C_c
        C_min = C_h
        Effectiveness = (Inlet_h.Properties.T - Outlet_h.Properties.T) / (Inlet_h.Properties.T - Inlet_c_cir.Properties.T)
    End If
    
    If D_o = 0 Then
        tc_h_off = tc_t_gas(Temp_h, Inlet_h.Composition.Composition)
        my_h_off = mu_t_gas(Temp_h, Inlet_h.Composition.Composition)
        UA_off = UA * (Inlet_h.Massflow.m / Mass_h) ^ 0.65 * (Cp_h / Heat_capa_h) ^ 0.33 * (tc_h_off / tc_h) ^ 0.67 * (my_h_off / my_h) ^ -0.32
        NTU = UA_off / C_min
    Else
        'R_i_off 구하기''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        my_c_off = my_pT(Press_c, Temp_c)
        tc_c_off = tc_pT(Press_c, Temp_c) * 10 ^ -3
        R_i_off = R_i * (Mass_in_c / Inlet_c_cir.Massflow.m) ^ 0.8 * (my_c / my_c_off) ^ 0.4 * (Heat_capa_c / Cp_c) ^ 0.4 * (tc_c / tc_c_off) ^ 0.6
        
        'R_w_off 구하기''''''''''''''''''''''''''''''''''''''''
        Temp_tube = (Temp_c + Temp_h) / 2
        If Material = "CS" Then
            k_tube = 46.73 * 0.001 + (Temp_tube - 260) * -0.0249 * 0.001 'kW/m-C basis
        ElseIf Material = "SS" Then
            k_tube = 18.34 * 0.001 + (Temp_tube - 260) * 0.015 * 0.001 'kW/m-C basis
        ElseIf Material = "T91" Then
            k_tube = 27 * 0.001 + (Temp_tube - 260) * 0.0053 * 0.001 'kW/m-C basis
        ElseIf Material = "T11" Then
            k_tube = 37.55 * 0.001 + (Temp_tube - 260) * -0.0137 * 0.001 'kW/m-C basis
        ElseIf Material = "T22" Then
            k_tube = 36.86 * 0.001 + (Temp_tube - 260) * -0.0109 * 0.001 'kW/m-C basis
        End If
        R_w_off = Application.WorksheetFunction.Ln(D_o / D_i) * D_o / (2 * k_tube)
        
        'R_o_off 구하기''''''''''''''''''''''''''''''''''''''''''''''
        my_h_off = mu_t_gas(Temp_h, Inlet_h.Composition.Composition)
        tc_h_off = tc_t_gas(Temp_h, Inlet_h.Composition.Composition)
        R_o_off = R_o * (Mass_h / Inlet_h.Massflow.m) * (Heat_capa_h / Cp_h) ^ 0.33 * (my_h / my_h_off) ^ -0.67 * (tc_h / tc_h_off) ^ 0.67
        
        'U_off 구하기 '''''''''''''''''''''''''''''''''
        U_off = 1 / (R_o_off / Fin_efficiency + f_fin * R_w_off + f_fin * R_i_off)
        NTU = U_off * A_o / C_min
    End If
    
    y_for = Effectiveness - (1 - Exp(-NTU * (1 - C_r))) / (1 - C_r * Exp(-NTU * (1 - C_r)))

        ' Adaptive relaxation
    If Abs(y_for) > 0.01 Then
        step_size = 0.3  ' Decrease step size if error is large
    ElseIf Abs(y_for) <= 0.01 Then
        step_size = 1  ' Increase step size for faster convergence
    End If
    
    h_c_assume = h_c_assume - step_size * (Y / ((y_for - Y) / 10 ^ (-6)))
    

    Loop
    
    'outlet properties
    Set Outlet_c.Properties = Outlet_c_cir.Properties
    Set Outlet_c.Massflow = Inlet_c.Massflow
    
    If BFP_Cal_option = "Press Cal" Then
        Stop_cal_value = Inlet_c.Properties.P
    Else
        Stop_cal_value = Outlet_c.Properties.T
    End If
    

 End Select

End Sub



