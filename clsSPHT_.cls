VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSPHT_"
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

Public Steam_tc As New clsSteam_thermal_conductivity_
'HRSG_read'''''''''''''''''
Public mode, j As Integer
Public name, Component_name As Variant

Public delta_T_bypass, Gas_ratio As Double 'bypass인 경우 hotside temperature drop/

Public E_loss, P_loss_hot, P_loss_valve, P_loss_cold As Double
Public Method_P_loss_hot, Method_P_loss_valve, Method_P_loss_cold As Variant
Public Calc_option, Input_value As Variant
Public D_o, T_tube, Material, L_tube, n_tube, x_tube As Variant
Public Check As Integer
''''''''''''''''''''''''''''
Public delta_P_hot, delta_P_valve, delta_P_cold As Variant
Public delta_P_valve_off, delta_P_cold_off As Double

Private Q_c, Q_h As Double
Private Cp_c, Cp_h, C_c, C_h, C_min, C_r As Double
Public Effectiveness As Double

Public Press_c, Temp_c, Press_SI, Temp_SI As Double
Public Temp_tube, k_tube As Double
Public Temp_h As Double
Public D_i, rho_c, Velocity, A_o As Double 'Steam velocity & A_o
Public f_fin, Fin_efficiency As Double
Public NTU, UA, U As Double 'U
Public my_c, tc_c, Re, PR, R_i As Double 'R_i
Public R_w As Double  'R_w
Public R_o As Double 'R_o

Public Stop_cal_value As Double

Public T_c_assume As Double
Public Mass_h, Press_h_in, Temp_h_in As Double
Public Mass_c, sv_in_c, sv_in_c_valve As Double

Public my_h, Heat_capa_h, tc_h As Double
Public Heat_capa_c As Double

Public R_i_off, R_w_off, R_o_off, U_off, UA_off As Double 'off thermal resistance
Public Check_Mass As Variant 'EVAP Mass 계산 완료

Sub calSPHT()

Select Case mode

Case 0

    Dim dummy As Variant
    Dim del_T_2, del_T_1, LMTD As Double
    
    Dim iter As Integer
    Dim Y, y_for As Double 'calc_option=effectiveness
    Dim T_cold_out As Double
    
'     Cold side inlet의 stream 이 연결되지 않을 경우 hot side stream bypass
'     동일 pressure level의 ECON 계산 완료 이후에 stream이 연결됨
    
    If IsEmpty(Inlet_c.Properties.T) Then
        
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
    Inlet_c_valve.Properties.T = T_ph(Inlet_c_valve.Properties.P, Inlet_c_valve.Properties.h)
    
    Set Outlet_h.Massflow = Inlet_h.Massflow
    Set Outlet_c.Massflow = Inlet_c.Massflow
    Set Outlet_h.Composition = Inlet_h.Composition

    'Calculation option check
    Select Case Calc_option
    Case "Steam outlet temperature", "Approach temperature (Gas in - steam out)", "Degrees of superheat"
 
    ' Cold side outlet temperature
    If Calc_option = "Steam outlet temperature" Then 'User input : Steam outlet temperature
        Outlet_c.Properties.T = Input_value
    ElseIf Calc_option = "Approach temperature (Gas in - steam out)" Then 'User input : Approach temperature
        Outlet_c.Properties.T = Inlet_h.Properties.T - Input_value
    ElseIf Calc_option = "Degrees of superheat" Then 'User input : Degrees of superheat
        Outlet_c.Properties.T = Tsat_p(Inlet_c.Properties.P) + Input_value
    End If
    
    ' Cold side Q 계산
    If Method_P_loss_cold = "P_loss_cold" Then
        Outlet_c.Properties.P = Inlet_c_valve.Properties.P * (1 - P_loss_cold)
        delta_P_cold = Inlet_c_valve.Properties.P - Outlet_c.Properties.P
    ElseIf Method_P_loss_cold = "delta_P_cold" Then
        delta_P_cold = P_loss_cold
        Outlet_c.Properties.P = Inlet_c_valve.Properties.P - P_loss_cold
    End If
    
    Outlet_c.Properties.h = h_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
    Q_c = Outlet_c.Massflow.m * (Outlet_c.Properties.h - Inlet_c_valve.Properties.h)
   
    ' Hot side Q
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
    
    ' Cold side outlet property calculation
    Outlet_c.Properties.s = s_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
    Outlet_c.Properties.cp = Cp_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
    Outlet_c.Properties.Cv = Cv_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
    Outlet_c.Properties.k = Outlet_c.Properties.cp / Outlet_c.Properties.Cv
    Outlet_c.Properties.x = x_ph(Outlet_c.Properties.P, Outlet_c.Properties.h)
    Outlet_c.Properties.Prop_save
    Press_c = (Inlet_c_valve.Properties.P + Outlet_c.Properties.P) / 2
    Temp_c = (Inlet_c_valve.Properties.T + Outlet_c.Properties.T) / 2
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'LMTD 구하기''''''''''''''''''''''''''''''''''''''''''''''''''
    del_T_2 = Inlet_h.Properties.T - Outlet_c.Properties.T
    del_T_1 = Outlet_h.Properties.T - Inlet_c_valve.Properties.T
    
    If del_T_2 < 0 Or del_T_1 < 0 Then
    MsgBox "Error - " & name & " : Temperature cross"
    Error 1
    End If
    
    LMTD = (del_T_2 - del_T_1) / Application.WorksheetFunction.Ln(del_T_2 / del_T_1)
    
    'Effectiveness_Temp'''''''''''''''''''''''''''''''''''''''''''''''''''''
    Cp_c = (Outlet_c.Properties.h - Inlet_c_valve.Properties.h) / (Outlet_c.Properties.T - Inlet_c_valve.Properties.T)
    C_c = Cp_c * Outlet_c.Massflow.m
    Cp_h = (Inlet_h.Properties.h - Outlet_h.Properties.h) / (Inlet_h.Properties.T - Outlet_h.Properties.T)
    C_h = Cp_h * Outlet_h.Massflow.m
    If C_c < C_h Then
        C_r = C_c / C_h
        C_min = C_c
        Effectiveness = (Outlet_c.Properties.T - Inlet_c_valve.Properties.T) / (Inlet_h.Properties.T - Inlet_c_valve.Properties.T)
    ElseIf C_c > C_h Then
        C_r = C_h / C_c
        C_min = C_h
        Effectiveness = (Inlet_h.Properties.T - Outlet_h.Properties.T) / (Inlet_h.Properties.T - Inlet_c_valve.Properties.T)
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case "Effectiveness" 'User input : Effectiveness

    T_cold_out = Inlet_h.Properties.T - 20 'cold side outlet 온도 초기값 설정
    
    iter = 0
    Effectiveness = Input_value
    Do
    iter = iter + 1
        
    ' Cold side Q 계산
    Outlet_c.Properties.T = T_cold_out
    If Method_P_loss_cold = "P_loss_cold" Then
        Outlet_c.Properties.P = Inlet_c_valve.Properties.P * (1 - P_loss_cold)
        delta_P_cold = Inlet_c_valve.Properties.P - Outlet_c.Properties.P
    ElseIf Method_P_loss_cold = "delta_P_cold" Then
        delta_P_cold = P_loss_cold
        Outlet_c.Properties.P = Inlet_c_valve.Properties.P - P_loss_cold
    End If
    Outlet_c.Properties.h = h_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
    
    Q_c = Outlet_c.Massflow.m * (Outlet_c.Properties.h - Inlet_c_valve.Properties.h)
    
    ' Hot side Q
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
    
    ' Cold side outlet property calculation
    Outlet_c.Properties.s = s_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
    Outlet_c.Properties.cp = Cp_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
    Outlet_c.Properties.Cv = Cv_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
    Outlet_c.Properties.k = Outlet_c.Properties.cp / Outlet_c.Properties.Cv
    Outlet_c.Properties.x = x_ph(Outlet_c.Properties.P, Outlet_c.Properties.h)
    
    Outlet_c.Properties.Prop_save
    
    'Q = U*A*LMTD
    del_T_2 = Inlet_h.Properties.T - Outlet_c.Properties.T
    del_T_1 = Outlet_h.Properties.T - Inlet_c_valve.Properties.T
    
    If del_T_2 < 0 Or del_T_1 < 0 Then
    MsgBox "Error - " & name & " : Temperature cross"
    Error 1
    End If
    
    LMTD = (del_T_2 - del_T_1) / Application.WorksheetFunction.Ln(del_T_2 / del_T_1)
    UA = Q_c / LMTD
    
    'C_cold & C_hot
    Cp_c = (Outlet_c.Properties.h - Inlet_c_valve.Properties.h) / (Outlet_c.Properties.T - Inlet_c_valve.Properties.T)
    C_c = Cp_c * Outlet_c.Massflow.m
    Cp_h = (Inlet_h.Properties.h - Outlet_h.Properties.h) / (Inlet_h.Properties.T - Outlet_h.Properties.T)
    C_h = Cp_h * Outlet_h.Massflow.m
    
    If C_c < C_h Then
        C_r = C_c / C_h
        C_min = C_c
    ElseIf C_c > C_h Then
        C_r = C_h / C_c
        C_min = C_h
    End If
    
    NTU = UA / C_min
    Y = Effectiveness - (1 - Exp(-NTU * (1 - C_r))) / (1 - C_r * Exp(-NTU * (1 - C_r)))

    If Abs(Y) < 10 ^ (-8) Then
       Exit Do
    ElseIf iter > 100 Then
       MsgBox "반복계산 횟수가 100번 이상이 됨."
       Error 1
    End If
        
    Outlet_c.Properties.T = T_cold_out + 10 ^ (-6)
    If Method_P_loss_cold = "P_loss_cold" Then
        Outlet_c.Properties.P = Inlet_c_valve.Properties.P * (1 - P_loss_cold)
        delta_P_cold = Inlet_c_valve.Properties.P - Outlet_c.Properties.P
    ElseIf Method_P_loss_cold = "delta_P_cold" Then
        delta_P_cold = P_loss_cold
        Outlet_c.Properties.P = Inlet_c_valve.Properties.P - P_loss_cold
    End If
    Outlet_c.Properties.h = h_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
    
    Q_c = Outlet_c.Massflow.m * (Outlet_c.Properties.h - Inlet_c_valve.Properties.h)
    
    ' Hot side Q
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
    
    ' Cold side outlet property calculation
    Outlet_c.Properties.s = s_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
    Outlet_c.Properties.cp = Cp_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
    Outlet_c.Properties.Cv = Cv_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
    Outlet_c.Properties.k = Outlet_c.Properties.cp / Outlet_c.Properties.Cv
    Outlet_c.Properties.x = x_ph(Outlet_c.Properties.P, Outlet_c.Properties.h)
    
    Outlet_c.Properties.Prop_save
    
    'Q = U*A*LMTD
    del_T_2 = Inlet_h.Properties.T - Outlet_c.Properties.T
    del_T_1 = Outlet_h.Properties.T - Inlet_c_valve.Properties.T
    
    If del_T_2 < 0 Or del_T_1 < 0 Then
    MsgBox "Error - " & name & " : Temperature cross"
    Error 1
    End If
    
    LMTD = (del_T_2 - del_T_1) / Application.WorksheetFunction.Ln(del_T_2 / del_T_1)
    UA = Q_c / LMTD
    
    'C_cold & C_hot
    Cp_c = (Outlet_c.Properties.h - Inlet_c_valve.Properties.h) / (Outlet_c.Properties.T - Inlet_c_valve.Properties.T)
    C_c = Cp_c * Outlet_c.Massflow.m
    Cp_h = (Inlet_h.Properties.h - Outlet_h.Properties.h) / (Inlet_h.Properties.T - Outlet_h.Properties.T)
    C_h = Cp_h * Outlet_h.Massflow.m
    
    If C_c < C_h Then
        C_r = C_c / C_h
        C_min = C_c
    ElseIf C_c > C_h Then
        C_r = C_h / C_c
        C_min = C_h
    End If
    
    NTU = UA / C_min
    y_for = Effectiveness - (1 - Exp(-NTU * (1 - C_r))) / (1 - C_r * Exp(-NTU * (1 - C_r)))
    
    T_cold_out = T_cold_out - (Y / ((y_for - Y) / 10 ^ (-6)))
    
    Loop
    End Select
    
    If D_o = 0 Then
        NTU = Application.WorksheetFunction.Ln((Effectiveness - 1) / (Effectiveness * C_r - 1)) / (C_r - 1)
        UA = NTU * C_min
    Else
        'Steam velocity & A_o'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        D_i = D_o - 2 * T_tube
        rho_c = rho_pT(Press_c, Temp_c)
        Velocity = 4 * Inlet_c.Massflow.m / (n_tube * rho_c * Application.WorksheetFunction.pi * D_i ^ 2)
        A_o = f_fin * Application.WorksheetFunction.pi * D_o * n_tube * L_tube * x_tube
    
        'U''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        NTU = Application.WorksheetFunction.Ln((Effectiveness - 1) / (Effectiveness * C_r - 1)) / (C_r - 1)
        UA = NTU * C_min
        U = UA / A_o
    
        'R_i 구하기'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        my_c = my_pT(Press_c, Temp_c)
        'steam_tc 구하기'''''''''''''''''''''''''''''
        '모듈의 함수 사용가능 여부 판단
        Press_SI = Press_c * 10 ^ -3
        Temp_SI = Temp_c + 273.15
        If Temp_SI < 0 Or Press_SI < 0.000611657 Or Temp_SI > 800 Or Press_SI > 400 Or Not ((Press_SI <= 100 And Temp_SI <= 100 + 273.15) Or (Press_SI <= 150 And Temp_SI <= 400 + 273.15) Or (Press_SI <= 200 And Temp_SI <= 250 + 273.15) Or (Press_SI <= 400 And Temp_SI <= 125 + 273.15)) Then
            Steam_tc.Steam_tc_read
            With Steam_tc
                .Press = Press_c
                .Temp = Temp_c
            End With
            Steam_tc.Steam_tc_cal
            tc_c = Steam_tc.tc
        Else '가능
            tc_c = tc_pT(Press_c, Temp_c) * 10 ^ -3
        End If
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
    
    Stop_cal_value = Outlet_h.Properties.T

    'Off-design의 reference'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    T_c_assume = Outlet_c.Properties.T 'Off-design 초기값
    
    'dela_P_off'''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Mass_h = Inlet_h.Massflow.m
    Press_h_in = Inlet_h.Properties.P
    Temp_h_in = Inlet_h.Properties.T
    
    Mass_c = Outlet_c.Massflow.m
    If Check = 1 Then 'EVAP 다음 SPHT
        sv_in_c = vV_p(Inlet_c.Properties.P)
    Else
        sv_in_c = v_pT(Inlet_c.Properties.P, Inlet_c.Properties.T)
    End If
    
    If Check = 1 And delta_P_valve > 1000 Then
        sv_in_c_valve = v_pT(Inlet_c_valve.Properties.P, Inlet_c_valve.Properties.T)
    Else
        sv_in_c_valve = sv_in_c
    End If
    'R_off'''''''''''''''''''''''''''''''''''''''''''''''''''''''
    my_h = mu_t_gas(Temp_h, Inlet_h.Composition.Composition)
    Heat_capa_h = Cp_h
    tc_h = tc_t_gas(Temp_h, Inlet_h.Composition.Composition)
    Heat_capa_c = Cp_c
    
'OFF-DESIGN''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Case 1

    Dim delta_P_hot_off, P_in_cal As Double
    Dim sv_in_c_off, sv_in_c_valve_off As Double
    Dim my_c_off, tc_c_off, my_c_valve_off, tc_c_valve_off, my_h_off, tc_h_off As Double
    
    Dim M_cold As Double
    Dim step_size As Double
        Dim Toler_Adj As Double
    
        
    Toler_Adj = 10 ^ (-8)
    
    If Not Gas_ratio = 0 Then
    Inlet_h.Massflow.m = Inlet_h.Massflow.m * Gas_ratio
    End If
    
    Set Outlet_h.Massflow = Inlet_h.Massflow
    Set Outlet_h.Composition = Inlet_h.Composition
    Set Outlet_c.Massflow = Inlet_c.Massflow
    
    
    If Not Component_name = "SPHT_RH" Then
        If delta_P_valve = 0 And delta_P_cold = 0 Then
            Inlet_c.Properties.P = Outlet_c.Properties.P
            Inlet_c.Properties.T = T_ph(Inlet_c.Properties.P, Inlet_c.Properties.h)
        Else
            'Cold side 입구 압력 계산 Do~Loop''''''''''''''''
            Do
            
            If delta_P_valve = 0 Then
                delta_P_valve_off = 0
            Else
                If Check = 1 Then 'EVAP 다음 SPHT
                    sv_in_c_off = vV_p(Inlet_c.Properties.P)
                Else
                    sv_in_c_off = v_pT(Inlet_c.Properties.P, Inlet_c.Properties.T)
                End If
                delta_P_valve_off = delta_P_valve * ((Outlet_c.Massflow.m / Mass_c) ^ 1.98) * (sv_in_c_off / sv_in_c)
            End If
            Inlet_c_valve.Properties.P = Inlet_c.Properties.P - delta_P_valve_off
            Inlet_c_valve.Properties.h = Inlet_c.Properties.h
            Inlet_c_valve.Properties.T = T_ph(Inlet_c_valve.Properties.P, Inlet_c_valve.Properties.h)
            
            If delta_P_cold = 0 Then
                delta_P_cold_off = 0
            Else
                If delta_P_valve_off < 1000 Then
                    sv_in_c_valve_off = sv_in_c_valve
                Else
                    sv_in_c_valve_off = v_ph(Inlet_c_valve.Properties.P, Inlet_c_valve.Properties.h)
                End If
                delta_P_cold_off = delta_P_cold * ((Outlet_c.Massflow.m / Mass_c) ^ 1.98) * (sv_in_c_valve_off / sv_in_c_valve)
'                delta_P_cold_off = delta_P_cold
            End If
            P_in_cal = Inlet_c.Properties.P
            Inlet_c.Properties.P = Outlet_c.Properties.P + delta_P_valve_off + delta_P_cold_off
         
            If Abs(P_in_cal - Inlet_c.Properties.P) < 0.1 Then Exit Do
            Inlet_c.Properties.T = T_ph(Inlet_c.Properties.P, Inlet_c.Properties.h)
            Loop
            '''''''''''''''''''''''''''''''''''''''''''''''''
        End If
    End If
    If Check = 1 Then 'EVAP 다음 SPHT
        sv_in_c_off = vV_p(Inlet_c.Properties.P)
    Else
        sv_in_c_off = v_pT(Inlet_c.Properties.P, Inlet_c.Properties.T)
    End If
    delta_P_valve_off = delta_P_valve * ((Outlet_c.Massflow.m / Mass_c) ^ 1.98) * (sv_in_c_off / sv_in_c)

    Inlet_c_valve.Properties.P = Inlet_c.Properties.P - delta_P_valve_off
    Inlet_c_valve.Properties.h = Inlet_c.Properties.h
    Inlet_c_valve.Properties.T = T_ph(Inlet_c_valve.Properties.P, Inlet_c_valve.Properties.h)

    If delta_P_valve_off < 1000 Then
        sv_in_c_valve_off = sv_in_c_valve
    Else
        sv_in_c_valve_off = v_ph(Inlet_c_valve.Properties.P, Inlet_c_valve.Properties.h)
    End If
    delta_P_cold_off = delta_P_cold * ((Outlet_c.Massflow.m / Mass_c) ^ 1.98) * (sv_in_c_valve_off / sv_in_c_valve)
'    delta_P_cold_off = delta_P_cold
    Outlet_c.Properties.P = Inlet_c_valve.Properties.P - delta_P_cold_off

    iter = 0
    'Cold side 출구 온도 계산 Do~Loop''''''''''''''''
    Do
    iter = iter + 1
    Outlet_c.Properties.T = T_c_assume
        
    Press_c = (Inlet_c_valve.Properties.P + Outlet_c.Properties.P) / 2
    Temp_c = (Inlet_c_valve.Properties.T + Outlet_c.Properties.T) / 2
    
    ' Cold side Q 계산
    Outlet_c.Properties.h = h_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
    Q_c = Outlet_c.Massflow.m * (Outlet_c.Properties.h - Inlet_c_valve.Properties.h)

    ' Hot side Q
    Q_h = Q_c / (1 - E_loss)
    
    ' Hot side outlet property calculation
    delta_P_hot_off = delta_P_hot * ((Inlet_h.Massflow.m / Mass_h) ^ 1.84) * ((Inlet_h.Properties.T + 273.15) / (Temp_h_in + 273.15)) * (Press_h_in / Inlet_h.Properties.P) 'gas side P_loss_off
    Outlet_h.Properties.P = Inlet_h.Properties.P - delta_P_hot_off
    Outlet_h.Properties.h = Inlet_h.Properties.h - (Q_h / Outlet_h.Massflow.m)
    
    dummy = idealgas("t", "p", Outlet_h.Properties.P, "h", Outlet_h.Properties.h, Inlet_h.Composition.Composition)
    Outlet_h.Properties.T = dummy(0)
    Temp_h = (Inlet_h.Properties.T + Outlet_h.Properties.T) / 2

    'C_cold & C_hot
    Cp_c = (Outlet_c.Properties.h - Inlet_c_valve.Properties.h) / (Outlet_c.Properties.T - Inlet_c_valve.Properties.T)
    C_c = Cp_c * Outlet_c.Massflow.m
    Cp_h = (Inlet_h.Properties.h - Outlet_h.Properties.h) / (Inlet_h.Properties.T - Outlet_h.Properties.T)
    C_h = Cp_h * Outlet_h.Massflow.m
    
    If C_c < C_h Then
        C_r = C_c / C_h
        C_min = C_c
        Effectiveness = (Outlet_c.Properties.T - Inlet_c_valve.Properties.T) / (Inlet_h.Properties.T - Inlet_c_valve.Properties.T)
    ElseIf C_c > C_h Then
        C_r = C_h / C_c
        C_min = C_h
        Effectiveness = (Inlet_h.Properties.T - Outlet_h.Properties.T) / (Inlet_h.Properties.T - Inlet_c_valve.Properties.T)
    End If
    
    If D_o = 0 Then
        tc_h_off = tc_t_gas(Temp_h, Inlet_h.Composition.Composition)
        my_h_off = mu_t_gas(Temp_h, Inlet_h.Composition.Composition)
        UA_off = UA * (Inlet_h.Massflow.m / Mass_h) ^ 0.65 * (Cp_h / Heat_capa_h) ^ 0.33 * (tc_h_off / tc_h) ^ 0.67 * (my_h_off / my_h) ^ -0.32 * (Outlet_c.Massflow.m / Mass_c) ^ 0.15
        NTU = UA_off / C_min
    Else
        'R_i_off 구하기''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        my_c_off = my_pT(Press_c, Temp_c)
        'steam_tc_off 구하기'''''''''''''''''''''''''''''
        '모듈의 함수 사용가능 여부 판단
        Press_SI = Press_c * 10 ^ -3
        Temp_SI = Temp_c + 273.15
        If Temp_SI < 0 Or Press_SI < 0.000611657 Or Temp_SI > 800 Or Press_SI > 400 Or Not ((Press_SI <= 100 And Temp_SI <= 100 + 273.15) Or (Press_SI <= 150 And Temp_SI <= 400 + 273.15) Or (Press_SI <= 200 And Temp_SI <= 250 + 273.15) Or (Press_SI <= 400 And Temp_SI <= 125 + 273.15)) Then
            Steam_tc.Steam_tc_read
            With Steam_tc
                .Press = Press_c
                .Temp = Temp_c
            End With
            Steam_tc.Steam_tc_cal
            tc_c_off = Steam_tc.tc
        Else '가능
            tc_c_off = tc_pT(Press_c, Temp_c) * 10 ^ -3
        End If
        R_i_off = R_i * (Mass_c / Outlet_c.Massflow.m) ^ 0.8 * (my_c / my_c_off) ^ 0.4 * (Heat_capa_c / Cp_c) ^ 0.4 * (tc_c / tc_c_off) ^ 0.6
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'R_o_off 구하기''''''''''''''''''''''''''''''''''''''''''''''
        my_h_off = mu_t_gas(Temp_h, Inlet_h.Composition.Composition)
        tc_h_off = tc_t_gas(Temp_h, Inlet_h.Composition.Composition)
        R_o_off = R_o * (Mass_h / Inlet_h.Massflow.m) * (Heat_capa_h / Cp_h) ^ 0.33 * (my_h / my_h_off) ^ -0.67 * (tc_h / tc_h_off) ^ 0.67
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
    Outlet_c.Properties.T = T_c_assume + 10 ^ (-6)
        
    Press_c = (Inlet_c_valve.Properties.P + Outlet_c.Properties.P) / 2
    Temp_c = (Inlet_c_valve.Properties.T + Outlet_c.Properties.T) / 2
    
    ' Cold side Q 계산
    Outlet_c.Properties.h = h_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
    Q_c = Outlet_c.Massflow.m * (Outlet_c.Properties.h - Inlet_c_valve.Properties.h)
    
    ' Hot side Q
    Q_h = Q_c / (1 - E_loss)
    
    ' Hot side outlet property calculation
    delta_P_hot_off = delta_P_hot * ((Inlet_h.Massflow.m / Mass_h) ^ 1.84) * ((Inlet_h.Properties.T + 273.15) / (Temp_h_in + 273.15)) * (Press_h_in / Inlet_h.Properties.P) 'gas side P_loss_off
    Outlet_h.Properties.P = Inlet_h.Properties.P - delta_P_hot_off
    Outlet_h.Properties.h = Inlet_h.Properties.h - (Q_h / Outlet_h.Massflow.m)
    dummy = idealgas("t", "p", Outlet_h.Properties.P, "h", Outlet_h.Properties.h, Inlet_h.Composition.Composition)
    Outlet_h.Properties.T = dummy(0)
    Temp_h = (Inlet_h.Properties.T + Outlet_h.Properties.T) / 2

    'C_cold & C_hot
    Cp_c = (Outlet_c.Properties.h - Inlet_c_valve.Properties.h) / (Outlet_c.Properties.T - Inlet_c_valve.Properties.T)
    C_c = Cp_c * Outlet_c.Massflow.m
    Cp_h = (Inlet_h.Properties.h - Outlet_h.Properties.h) / (Inlet_h.Properties.T - Outlet_h.Properties.T)
    C_h = Cp_h * Outlet_h.Massflow.m
    
    If C_c < C_h Then
        C_r = C_c / C_h
        C_min = C_c
        Effectiveness = (Outlet_c.Properties.T - Inlet_c_valve.Properties.T) / (Inlet_h.Properties.T - Inlet_c_valve.Properties.T)
    ElseIf C_c > C_h Then
        C_r = C_h / C_c
        C_min = C_h
        Effectiveness = (Inlet_h.Properties.T - Outlet_h.Properties.T) / (Inlet_h.Properties.T - Inlet_c_valve.Properties.T)
    End If
    
    If D_o = 0 Then
        tc_h_off = tc_t_gas(Temp_h, Inlet_h.Composition.Composition)
        my_h_off = mu_t_gas(Temp_h, Inlet_h.Composition.Composition)
        UA_off = UA * (Inlet_h.Massflow.m / Mass_h) ^ 0.65 * (Cp_h / Heat_capa_h) ^ 0.33 * (tc_h_off / tc_h) ^ 0.67 * (my_h_off / my_h) ^ -0.32 * (Outlet_c.Massflow.m / Mass_c) ^ 0.15
        NTU = UA_off / C_min
    Else
        'R_i_off 구하기''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        my_c_off = my_pT(Press_c, Temp_c)
        'steam_tc_off 구하기'''''''''''''''''''''''''''''
        '모듈의 함수 사용가능 여부 판단
        Press_SI = Press_c * 10 ^ -3
        Temp_SI = Temp_c + 273.15
        If Temp_SI < 0 Or Press_SI < 0.000611657 Or Temp_SI > 800 Or Press_SI > 400 Or Not ((Press_SI <= 100 And Temp_SI <= 100 + 273.15) Or (Press_SI <= 150 And Temp_SI <= 400 + 273.15) Or (Press_SI <= 200 And Temp_SI <= 250 + 273.15) Or (Press_SI <= 400 And Temp_SI <= 125 + 273.15)) Then
            Steam_tc.Steam_tc_read
            With Steam_tc
                .Press = Press_c
                .Temp = Temp_c
            End With
            Steam_tc.Steam_tc_cal
            tc_c_off = Steam_tc.tc
        Else '가능
            tc_c_off = tc_pT(Press_c, Temp_c) * 10 ^ -3
        End If
        R_i_off = R_i * (Mass_c / Outlet_c.Massflow.m) ^ 0.8 * (my_c / my_c_off) ^ 0.4 * (Heat_capa_c / Cp_c) ^ 0.4 * (tc_c / tc_c_off) ^ 0.6
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'R_o_off 구하기''''''''''''''''''''''''''''''''''''''''''''''
        my_h_off = mu_t_gas(Temp_h, Inlet_h.Composition.Composition)
        tc_h_off = tc_t_gas(Temp_h, Inlet_h.Composition.Composition)
        R_o_off = R_o * (Mass_h / Inlet_h.Massflow.m) * (Heat_capa_h / Cp_h) ^ 0.33 * (my_h / my_h_off) ^ -0.67 * (tc_h / tc_h_off) ^ 0.67
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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

    T_c_assume = T_c_assume - step_size * (Y / ((y_for - Y) / 10 ^ (-6)))
    

    

    Loop
    '''''''''''''''''''''''''''''''''''''''''''''''''
    
    dummy = idealgas("sck", "p", Outlet_h.Properties.P, "t", Outlet_h.Properties.T, Inlet_h.Composition.Composition)
    Outlet_h.Properties.s = dummy(0)
    Outlet_h.Properties.C = dummy(1)
    Outlet_h.Properties.k = dummy(2)
    Outlet_h.Properties.MW = Inlet_h.Properties.MW
    Outlet_h.Properties.r = Inlet_h.Properties.r
    Outlet_h.Properties.Prop_save
    
    Outlet_c.Properties.s = s_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
    Outlet_c.Properties.cp = Cp_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
    Outlet_c.Properties.Cv = Cv_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
    Outlet_c.Properties.k = Outlet_c.Properties.cp / Outlet_c.Properties.Cv
    Outlet_c.Properties.x = x_ph(Outlet_c.Properties.P, Outlet_c.Properties.h)
    Outlet_c.Properties.Prop_save
    
    Stop_cal_value = Inlet_c.Properties.P
    
      If D_o = 0 Then
    UA = UA_off
  Else
   UA = U_off * A_o
   Velocity = 4 * Inlet_c.Massflow.m / (n_tube * (1 / sv_in_c_off) * Application.WorksheetFunction.pi * D_i ^ 2)
  End If

    
 End Select

End Sub
