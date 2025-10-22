VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEVAP_"
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
Public Inlet_cir As New clsStream_steam_
Public Outlet_cir As New clsStream_steam_
Public Outlet_c_BD As New clsStream_steam_

Public index_EVAP As Variant
'HRSG_read'''''''''''''''''
Public mode, j As Integer
Public name, Component_name, Pressure_level As Variant

Public delta_T_bypass, Gas_ratio As Double 'bypass인 경우 hotside temperature drop/
Public P_drum As Double

Public E_loss, P_loss_hot, P_loss_valve As Double
Public Method_P_loss_hot, Method_P_loss_valve As Variant
Public Calc_option, Input_value As Variant
Public D_o, T_tube, Material, L_tube, n_tube, x_tube As Variant
Public Mass_BD As Double 'EVAP_BD HP+IP mass
''''''''''''''''''''''''''''
Public delta_P_hot, delta_P_valve As Double
Public delta_P_valve_off As Double

Private Q_c, Q_h As Double
Private Cp_h, C_h, C_min As Double
Public Effectiveness As Double

Public Velocity As Double
Public D_i, A_tube, Mass_flux, A_o As Double 'Steam velocity & A_o
Public f_fin, Fin_efficiency As Double
Public NTU, UA, U As Double 'U
Public Press_tube, Temp_tube, k_tube As Double
Public Cir_ratio, Quality, Head, P_Static As Double 'circulation
Public rho_l, my_l, tc_l, Cp_l, Re, PR, h_sp As Double 'h_sp
Public rho_g, Enthalpy_g, my_g, X_tt, coeff_F As Double 'coeff_F
Public st, h_lg, delta_P, delta_T, h_nb As Double 'h_nb
Public coeff_S  As Double 'coeff_S
Public Temp_h As Double
Public R_i, R_w, R_o As Double 'R_i,R_w,R_o

Public Stop_cal_value As Double

Public Mass_h, Press_h_in, Temp_h_in, Enthalpy_l As Double
Public Mass_in_c, Mass_out_c, sv_in_c As Double
Public my_h, Heat_capa_h, tc_h As Double
Public R_i_off, R_w_off, R_o_off, U_off, UA_off As Double 'off thermal resistance
Public Check_Mass As Variant 'EVAP Mass 계산 완료
Public BFP_Cal_option As Variant


Sub calEVAP()

Select Case mode

Case 0
  
    Dim dummy As Variant
    Dim del_T_2, del_T_1, LMTD As Double
      
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
    
    If Inlet_h.Properties.T < Inlet_c.Properties.T Then
        MsgBox "Error - " & name & " : Inlet water temperature is too high" & Chr(13) & "Water temperature = " & Round(Inlet_c.Properties.T, 3) & " oC" & Chr(13) & "Gas temperature = " & Round(Inlet_h.Properties.T, 3) & " oC"
        Error 1
    End If
    
    f_fin = 10
    Fin_efficiency = 1
    Quality = 1 / Cir_ratio
    P_Static = rho_ph(Inlet_c.Properties.P, Inlet_c.Properties.h) * 9.81 * Head / 1000 'kPa

    If Not Gas_ratio = 0 Then
        Inlet_h.Massflow.m = Inlet_h.Massflow.m * Gas_ratio
    End If
    
      
    Set Outlet_h.Massflow = Inlet_h.Massflow
    Set Outlet_h.Composition = Inlet_h.Composition

    If P_drum = 0 Then     'P_drum only exists for LPB/DA or Evap_BD / Stand alone DA, Direct Feeding, HP/IP Eva -> P=0
        If P_loss_valve = 0 Then
            delta_P_valve = 0
            If P_Static = 0 Then
                Set Inlet_c_valve.Properties = Inlet_c.Properties
            Else
                Inlet_c_valve.Properties.P = Inlet_c.Properties.P - P_Static
                Inlet_c_valve.Properties.h = Inlet_c.Properties.h
                Inlet_c_valve.Properties.T = T_ph(Inlet_c_valve.Properties.P, Inlet_c_valve.Properties.h)
            End If
        Else
            If Method_P_loss_valve = "P_loss_valve" Then
                Inlet_c_valve.Properties.P = (Inlet_c.Properties.P - P_Static) * (1 - P_loss_valve)
            ElseIf Method_P_loss_valve = "delta_P_valve" Then
                Inlet_c_valve.Properties.P = Inlet_c.Properties.P - P_Static - P_loss_valve
            End If
            delta_P_valve = (Inlet_c.Properties.P - P_Static) - Inlet_c_valve.Properties.P
            Inlet_c_valve.Properties.h = Inlet_c.Properties.h
            Inlet_c_valve.Properties.T = T_ph(Inlet_c_valve.Properties.P, Inlet_c_valve.Properties.h)
        End If
    Else
        If P_drum = Inlet_c.Properties.P Then  'NO valve loss exists
            Set Inlet_c_valve.Properties = Inlet_c.Properties
        Else
            Inlet_c_valve.Properties.P = P_drum
            Inlet_c_valve.Properties.h = Inlet_c.Properties.h
            Inlet_c_valve.Properties.T = T_ph(Inlet_c_valve.Properties.P, Inlet_c_valve.Properties.h)
        End If
    End If
    
    ' Cold side outlet property
    Outlet_c.Properties.P = Inlet_c_valve.Properties.P
    Outlet_c.Properties.T = Tsat_p(Outlet_c.Properties.P)
    Outlet_c.Properties.h = h_px(Outlet_c.Properties.P, 1)
'    Outlet_c.Properties.Cp = Cp_ph(Outlet_c.Properties.P, Outlet_c.Properties.h)
'    Outlet_c.Properties.Cv = Cv_ph(Outlet_c.Properties.P, Outlet_c.Properties.h)
'    Outlet_c.Properties.k = Outlet_c.Properties.Cp / Outlet_c.Properties.Cv
'    Outlet_c.Properties.x = 1
    Outlet_c.Properties.Prop_save
    
    Outlet_c_BD.Properties.P = Inlet_c_valve.Properties.P
    Outlet_c_BD.Properties.T = Tsat_p(Outlet_c_BD.Properties.P)
    Outlet_c_BD.Properties.h = h_px(Outlet_c_BD.Properties.P, 0)                'Saturated Liquid State
    
    
   Select Case Calc_option
    Case "Desired steam production" 'User input = Steam_massflow

        Inlet_c.Massflow.m = Input_value * 10 / 36 + Mass_BD                    'Desired steam production + Blowdown = Inlet Water Flow

        Outlet_c.Massflow.m = Input_value * 10 / 36
        Inlet_c.Massflow.Mass_save
        Outlet_c.Massflow.Mass_save

        ' Cold side Q
        Q_c = (Outlet_c.Massflow.m * Outlet_c.Properties.h + Mass_BD * Outlet_c_BD.Properties.h) - Inlet_c.Massflow.m * Inlet_c_valve.Properties.h

        ' Hot side Q
        Q_h = Q_c / (1 - E_loss)
        
        ' Hot side property calculation
        If Method_P_loss_hot = "P_loss_hot" Then
            Outlet_h.Properties.P = Inlet_h.Properties.P * (1 - P_loss_hot)
            delta_P_hot = Inlet_h.Properties.P - Outlet_h.Properties.P
        ElseIf Method_P_loss_hot = "delta_P_hot" Then
            delta_P_hot = P_loss_hot
            Outlet_h.Properties.P = Inlet_h.Properties.P - P_loss_hot
        End If
        Outlet_h.Properties.h = Inlet_h.Properties.h - Q_h / Inlet_h.Massflow.m
        dummy = idealgas("t", "p", Outlet_h.Properties.P, "h", Outlet_h.Properties.h, Outlet_h.Composition.Composition)
        Outlet_h.Properties.T = dummy(0)

        ' Contraint check (Effectiveness maximum = 0.999)
        If Outlet_h.Properties.T <= Outlet_c.Properties.T Then
            Outlet_h.Properties.T = Outlet_c.Properties.T + (Inlet_h.Properties.T - Outlet_c.Properties.T) * 0.001
            dummy = idealgas("h", "p", Outlet_h.Properties.P, "t", Outlet_h.Properties.T, Outlet_h.Composition.Composition)
            Outlet_h.Properties.h = dummy(0)
            
            Q_h = Outlet_h.Massflow.m * (Inlet_h.Properties.h - Outlet_h.Properties.h)
            Q_c = Q_h * (1 - E_loss)
            
            Outlet_c.Massflow.m = (Q_c - Mass_BD * (Outlet_c_BD.Properties.h - Inlet_c_valve.Properties.h)) / (Outlet_c.Properties.h - Inlet_c_valve.Properties.h)
            Inlet_c.Massflow.m = Outlet_c.Massflow.m + Mass_BD
'            MsgBox name & " : WARNING - EVAP cannot raise user-requested vapor" & Chr(13) & "requested = " & Input_value & " kg/s" & "    limit = " & Round(Outlet_c.Massflow.m, 3) & " kg/s"
        End If
        
        dummy = idealgas("sck", "p", Outlet_h.Properties.P, "t", Outlet_h.Properties.T, Outlet_h.Composition.Composition)
        Outlet_h.Properties.s = dummy(0)
        Outlet_h.Properties.C = dummy(1)
        Outlet_h.Properties.k = dummy(2)
        Outlet_h.Properties.MW = Inlet_h.Properties.MW
        Outlet_h.Properties.r = Inlet_h.Properties.r
        Temp_h = (Inlet_h.Properties.T + Outlet_h.Properties.T) / 2
        Outlet_h.Properties.Prop_save

    Case "Pinch point temperature difference", "Effectiveness"
        ' Hot side outlet
        
        If Method_P_loss_hot = "P_loss_hot" Then
            Outlet_h.Properties.P = Inlet_h.Properties.P * (1 - P_loss_hot)
            delta_P_hot = Inlet_h.Properties.P - Outlet_h.Properties.P
        ElseIf Method_P_loss_hot = "delta_P_hot" Then
            delta_P_hot = P_loss_hot
            Outlet_h.Properties.P = Inlet_h.Properties.P - P_loss_hot
        End If
        If Calc_option = "Pinch point temperature difference" Then
            Outlet_h.Properties.T = Outlet_c.Properties.T + Input_value
        ElseIf Calc_option = "Effectiveness" Then
            Outlet_h.Properties.T = Outlet_c.Properties.T + (Inlet_h.Properties.T - Outlet_c.Properties.T) * (1 - Input_value)
        End If
        
        dummy = idealgas("hsck", "p", Outlet_h.Properties.P, "t", Outlet_h.Properties.T, Outlet_h.Composition.Composition)
        Outlet_h.Properties.h = dummy(0)
        Outlet_h.Properties.s = dummy(1)
        Outlet_h.Properties.C = dummy(2)
        Outlet_h.Properties.k = dummy(3)
        Outlet_h.Properties.MW = Inlet_h.Properties.MW
        Outlet_h.Properties.r = Inlet_h.Properties.r
        Temp_h = (Inlet_h.Properties.T + Outlet_h.Properties.T) / 2
        Outlet_h.Properties.Prop_save
        
        ' Hot side Q
        Q_h = Outlet_h.Massflow.m * (Inlet_h.Properties.h - Outlet_h.Properties.h)
        
        ' Cold side Q
        Q_c = Q_h * (1 - E_loss)
        
        Outlet_c.Massflow.m = (Q_c - Mass_BD * (Outlet_c_BD.Properties.h - Inlet_c_valve.Properties.h)) / (Outlet_c.Properties.h - Inlet_c_valve.Properties.h)
        Inlet_c.Massflow.m = Outlet_c.Massflow.m + Mass_BD


        Inlet_c.Massflow.Mass_save
        Outlet_c.Massflow.Mass_save
    
    Case Else
        MsgBox "EVAP : Input error"
        Error 1
    End Select
    Check_Mass = "Done"
    
     'Circulation'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Inlet_cir.Massflow.m = Inlet_c.Massflow.m * Cir_ratio
    Outlet_cir.Massflow.m = Inlet_cir.Massflow.m

    Inlet_cir.Properties.P = Inlet_c_valve.Properties.P
    Inlet_cir.Properties.T = Tsat_p(Inlet_cir.Properties.P)
    Inlet_cir.Properties.h = h_px(Inlet_cir.Properties.P, 0)
    
    Outlet_cir.Properties.P = Inlet_cir.Properties.P
    Outlet_cir.Properties.h = Inlet_cir.Properties.h + Quality * (Outlet_c.Properties.h - Inlet_c_valve.Properties.h)
    Outlet_cir.Properties.T = T_ph(Outlet_cir.Properties.P, Outlet_cir.Properties.h) ' inlet_cir과 같은 온도
    
    'LMTD 구하기''''''''''''''''''''''''''''''''''''''''''''''''''
    del_T_2 = Inlet_h.Properties.T - Outlet_cir.Properties.T
    del_T_1 = Outlet_h.Properties.T - Inlet_cir.Properties.T
    
    If del_T_2 < 0 Or del_T_1 < 0 Then
    MsgBox "Error - " & name & " : Temperature cross"
    Error 1
    End If
    
    LMTD = (del_T_2 - del_T_1) / Application.WorksheetFunction.Ln(del_T_2 / del_T_1)
    
    ' Effectiveness_temp calculation''''''''''''''''''''''''''''''''''''''''''''''''
    Cp_h = (Outlet_h.Properties.h - Inlet_h.Properties.h) / (Outlet_h.Properties.T - Inlet_h.Properties.T)
    C_h = Cp_h * Outlet_h.Massflow.m
    C_min = C_h ' EVAP에서는 C_h가 C_min이 됨(cold side는 상변화때문에 C_c가 무한대에 가까워짐)
    Effectiveness = (Inlet_h.Properties.T - Outlet_h.Properties.T) / (Inlet_h.Properties.T - Inlet_cir.Properties.T)
    
    If D_o = 0 Then
        NTU = -Application.WorksheetFunction.Ln(1 - Effectiveness)
        UA = NTU * C_min
    Else
        'steam_mass_flux & A_o''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        D_i = D_o - 2 * T_tube
        A_tube = (Application.WorksheetFunction.pi * D_i ^ 2) / 4
        Mass_flux = Inlet_cir.Massflow.m / (A_tube * n_tube)
        A_o = f_fin * Application.WorksheetFunction.pi * D_o * n_tube * L_tube * x_tube

        'U''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        NTU = -Application.WorksheetFunction.Ln(1 - Effectiveness)
        UA = NTU * C_min
        U = UA / A_o
    
        'R_w'''''''''''''''''''''''''''''''''''''''''''''''''
        Temp_tube = (Temp_h + Inlet_cir.Properties.T) / 2
        If Temp_tube > 373 Then '오류 제거
            Temp_tube = 373
        End If
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
        
        'R_i 구하기'''''''''''''''''''''''''''''''''''''
        'hsp.l 구하기'''''''''''''''''''
        If Inlet_cir.Properties.P > 16530 Then
             Enthalpy_l = Inlet_cir.Properties.h - 1
             my_l = my_ph(Inlet_cir.Properties.P, Enthalpy_l)
        Else
             my_l = my_ph(Inlet_cir.Properties.P, Inlet_cir.Properties.h)
        End If
             
        Re = ((1 - Quality) * Mass_flux * D_i) / my_l
      
        Cp_l = CpL_p(Inlet_cir.Properties.P) * 1000
        tc_l = tcL_p(Inlet_cir.Properties.P)
        PR = Cp_l * my_l / tc_l
      
        h_sp = 0.023 * Re ^ 0.8 * PR ^ 0.4 * tc_l / D_i
      
        'F구하기'''''''''''''''''''''''''''''''''''''''''''''''''''''''
        rho_l = rhoL_p(Inlet_cir.Properties.P)
        rho_g = rhoV_p(Inlet_cir.Properties.P)
        If Inlet_cir.Properties.P > 6031 Then
            Enthalpy_g = h_px(6031, 1)
            my_g = my_ph(6031, Enthalpy_g)
        Else
            Enthalpy_g = h_px(Inlet_cir.Properties.P, 1)
            my_g = my_ph(Inlet_cir.Properties.P, Enthalpy_g)
        End If
        X_tt = (Quality / (1 - Quality)) ^ 0.9 * (rho_l / rho_g) ^ 0.5 * (my_g / my_l) ^ 0.1
        coeff_F = 2.35 * (X_tt + 0.213) ^ 0.736
      
        'hnb 구하기'''''''''''''''''''''''''''''''''''''''''''''''
        st = st_p(Inlet_cir.Properties.P)
        h_lg = (h_px(Inlet_cir.Properties.P, 1) - h_px(Inlet_cir.Properties.P, 0)) * 1000
        Press_tube = psat_T(Temp_tube)
        delta_P = (Press_tube - Inlet_cir.Properties.P) * 1000
        delta_T = Temp_tube - Inlet_cir.Properties.T
        h_nb = (0.00122 * tc_l ^ 0.79 * Cp_l ^ 0.45 * rho_l ^ 0.49 * delta_T ^ 0.24 * delta_P ^ 0.75) / (st ^ 0.5 * my_l ^ 0.29 * h_lg ^ 0.24 * rho_g ^ 0.24)
        'S 구하기'''''''''''''''''''''''''''''''''''''''''''''''''''''''
        coeff_S = 1 / (1 + 2.53 * 10 ^ -6 * (Re * coeff_F ^ 1.25) ^ 1.17)
        R_i = D_o / (D_i * 0.001 * (coeff_F * h_sp + coeff_S * h_nb))
        
        'R_o 구하기'''''''''''''''''''''''''''''''''''''
        R_o = (1 / U - f_fin * R_w - f_fin * R_i) * Fin_efficiency
    End If
    
    Stop_cal_value = Outlet_c.Massflow.m
    
    'Off-design의 reference'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Mass_in_c = Inlet_c.Massflow.m
    Mass_out_c = Outlet_c.Massflow.m 'Off-design 초기값
    'dela_P_off'''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Mass_h = Inlet_h.Massflow.m
    Press_h_in = Inlet_h.Properties.P
    Temp_h_in = Inlet_h.Properties.T
   
    'R_off'''''''''''''''''''''''''''''''''''''''''''''''''''''''
    my_h = mu_t_gas(Temp_h, Inlet_h.Composition.Composition)
    Heat_capa_h = Cp_h
    tc_h = tc_t_gas(Temp_h, Inlet_h.Composition.Composition)
    
Case 1 '(inlet_c_valve=outlet_c)

    Dim delta_P_hot_off, P_in_cal As Double
    Dim my_h_off, tc_h_off As Double
    
    Dim M_cold As Double
    Dim iter As Integer
    Dim Y, y_for As Double
    Dim step_size As Double
    Dim Toler_Adj As Double
    
        
    Toler_Adj = 10 ^ (-8)
    

    If Inlet_h.Properties.T < Inlet_c.Properties.T Then
        MsgBox "Error - " & name & " : Inlet water temperature is too high" & Chr(13) & "Water temperature = " & Round(Inlet_c.Properties.T, 3) & " oC" & Chr(13) & "Gas temperature = " & Round(Inlet_h.Properties.T, 3) & " oC"
        Error 1
    End If

    If Not Gas_ratio = 0 Then
    Inlet_h.Massflow.m = Inlet_h.Massflow.m * Gas_ratio
    End If
    
    Set Outlet_h.Massflow = Inlet_h.Massflow
    Set Outlet_h.Composition = Inlet_h.Composition

    If P_drum = 0 Then
        Inlet_c_valve.Properties.P = Outlet_c.Properties.P
    Else
        Inlet_c_valve.Properties.P = P_drum
        Outlet_c.Properties.P = Inlet_c_valve.Properties.P
    End If
    
    
    Outlet_c.Properties.T = Tsat_p(Outlet_c.Properties.P)
    Outlet_c.Properties.h = h_px(Outlet_c.Properties.P, 1)
'    Outlet_c.Properties.Cp = Cp_ph(Outlet_c.Properties.P, Outlet_c.Properties.h)
'    Outlet_c.Properties.Cv = Cv_ph(Outlet_c.Properties.P, Outlet_c.Properties.h)
'    Outlet_c.Properties.k = Outlet_c.Properties.Cp / Outlet_c.Properties.Cv
'    Outlet_c.Properties.x = 1
    Outlet_c.Properties.Prop_save
    
    Outlet_c_BD.Properties.P = Outlet_c.Properties.P
    Outlet_c_BD.Properties.T = Tsat_p(Outlet_c_BD.Properties.P)
    Outlet_c_BD.Properties.h = h_px(Outlet_c_BD.Properties.P, 0)

    M_cold = Mass_out_c 'cold side 질량 초기값 설정
    
    iter = 0
    Do
    iter = iter + 1

    Outlet_c.Massflow.m = M_cold
    
    Inlet_c.Massflow.m = Outlet_c.Massflow.m + Mass_BD
    Inlet_c.Massflow.Mass_save
    Outlet_c.Massflow.Mass_save
    
    ' Cold side Q
    Q_c = (Outlet_c.Massflow.m * Outlet_c.Properties.h + Mass_BD * Outlet_c_BD.Properties.h) - Inlet_c.Massflow.m * Inlet_c.Properties.h

    Q_h = Q_c / (1 - E_loss)

    ' Hot side property calculation
    delta_P_hot_off = delta_P_hot * ((Inlet_h.Massflow.m / Mass_h) ^ 1.84) * ((Inlet_h.Properties.T + 273.15) / (Temp_h_in + 273.15)) * (Press_h_in / Inlet_h.Properties.P) 'gas side P_loss_off
    Outlet_h.Properties.P = Inlet_h.Properties.P - delta_P_hot_off
    Outlet_h.Properties.h = Inlet_h.Properties.h - Q_h / Inlet_h.Massflow.m
    dummy = idealgas("t", "p", Outlet_h.Properties.P, "h", Outlet_h.Properties.h, Outlet_h.Composition.Composition)
    Outlet_h.Properties.T = dummy(0)
    dummy = idealgas("sck", "p", Outlet_h.Properties.P, "t", Outlet_h.Properties.T, Outlet_h.Composition.Composition)
        Outlet_h.Properties.s = dummy(0)
        Outlet_h.Properties.C = dummy(1)
        Outlet_h.Properties.k = dummy(2)
        Outlet_h.Properties.MW = Inlet_h.Properties.MW
        Outlet_h.Properties.r = Inlet_h.Properties.r
        Outlet_h.Properties.Prop_save
    Temp_h = (Inlet_h.Properties.T + Outlet_h.Properties.T) / 2
    'Circulation'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Inlet_cir.Massflow.m = Inlet_c.Massflow.m * Cir_ratio
    Outlet_cir.Massflow.m = Inlet_cir.Massflow.m
    
    Inlet_cir.Properties.P = Outlet_c.Properties.P
    Inlet_cir.Properties.T = Tsat_p(Inlet_cir.Properties.P)
    Inlet_cir.Properties.h = h_px(Inlet_cir.Properties.P, 0)
    
    Outlet_cir.Properties.P = Inlet_cir.Properties.P
    Outlet_cir.Properties.h = Inlet_cir.Properties.h + Quality * (Outlet_c.Properties.h - Inlet_c.Properties.h)
    Outlet_cir.Properties.T = T_ph(Outlet_cir.Properties.P, Outlet_cir.Properties.h) ' inlet_cir과 같은 온도
    
    'Effectiveness 구하기(Temperature)
    Cp_h = (Outlet_h.Properties.h - Inlet_h.Properties.h) / (Outlet_h.Properties.T - Inlet_h.Properties.T)
    C_h = Cp_h * Outlet_h.Massflow.m
    C_min = C_h ' EVAP에서는 C_h가 C_min이 됨(cold side는 상변화때문에 C_c가 무한대에 가까워짐)
    Effectiveness = (Inlet_h.Properties.T - Outlet_h.Properties.T) / (Inlet_h.Properties.T - Inlet_cir.Properties.T)
    
    If D_o = 0 Then
        tc_h_off = tc_t_gas(Temp_h, Inlet_h.Composition.Composition)
        my_h_off = mu_t_gas(Temp_h, Inlet_h.Composition.Composition)
        UA_off = UA * (Inlet_h.Massflow.m / Mass_h) ^ 0.65 * (Cp_h / Heat_capa_h) ^ 0.33 * (tc_h_off / tc_h) ^ 0.67 * (my_h_off / my_h) ^ -0.32
'        UA_off = UA * (Inlet_h.Massflow.m / Mass_h) ^ 0.8
        NTU = UA_off / C_min
    Else
        'steam_mass_flux''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Mass_flux = Inlet_cir.Massflow.m / (A_tube * n_tube)
        'R_w_off'''''''''''''''''''''''''''''''''''''''''''''''''
        Temp_tube = (Temp_h + Inlet_cir.Properties.T) / 2
        If Temp_tube > 373 Then '오류 제거
            Temp_tube = 373
        End If
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

        'R_i_off 구하기'''''''''''''''''''''''''''''''''''''
        'hsp.l 구하기'''''''''''''''''''
        If Inlet_cir.Properties.P > 16530 Then
             Enthalpy_l = Inlet_cir.Properties.h - 1
             my_l = my_ph(Inlet_cir.Properties.P, Enthalpy_l)
        Else
             my_l = my_ph(Inlet_cir.Properties.P, Inlet_cir.Properties.h)
        End If
        Re = ((1 - Quality) * Mass_flux * D_i) / my_l
      
        Cp_l = CpL_p(Inlet_cir.Properties.P) * 1000
        tc_l = tcL_p(Inlet_cir.Properties.P)
        PR = Cp_l * my_l / tc_l
      
        h_sp = 0.023 * Re ^ 0.8 * PR ^ 0.4 * tc_l / D_i

        'F구하기'''''''''''''''''''''''''''''''''''''''''''''''''''''''
        rho_l = rhoL_p(Inlet_cir.Properties.P)
        rho_g = rhoV_p(Inlet_cir.Properties.P)
        If Inlet_cir.Properties.P > 6031 Then
            Enthalpy_g = h_px(6031, 1)
            my_g = my_ph(6031, Enthalpy_g)
        Else
            Enthalpy_g = h_px(Inlet_cir.Properties.P, 1)
            my_g = my_ph(Inlet_cir.Properties.P, Enthalpy_g)
        End If
        X_tt = (Quality / (1 - Quality)) ^ 0.9 * (rho_l / rho_g) ^ 0.5 * (my_g / my_l) ^ 0.1
        coeff_F = 2.35 * (X_tt + 0.213) ^ 0.736

        'hnb 구하기'''''''''''''''''''''''''''''''''''''''''''''''
        st = st_p(Inlet_cir.Properties.P)

        h_lg = (h_px(Inlet_cir.Properties.P, 1) - h_px(Inlet_cir.Properties.P, 0)) * 1000
        Press_tube = psat_T(Temp_tube)
        delta_P = (Press_tube - Inlet_cir.Properties.P) * 1000
        delta_T = Temp_tube - Inlet_cir.Properties.T

        h_nb = (0.00122 * tc_l ^ 0.79 * Cp_l ^ 0.45 * rho_l ^ 0.49 * delta_T ^ 0.24 * delta_P ^ 0.75) / (st ^ 0.5 * my_l ^ 0.29 * h_lg ^ 0.24 * rho_g ^ 0.24)

        'S 구하기'''''''''''''''''''''''''''''''''''''''''''''''''''''''
        coeff_S = 1 / (1 + 2.53 * 10 ^ -6 * (Re * coeff_F ^ 1.25) ^ 1.17)
        R_i_off = D_o / (D_i * 0.001 * (coeff_F * h_sp + coeff_S * h_nb))
        
        'R_o_off 구하기''''''''''''''''''''''''''''''''''''''''''''''
        my_h_off = mu_t_gas(Temp_h, Inlet_h.Composition.Composition)
        tc_h_off = tc_t_gas(Temp_h, Inlet_h.Composition.Composition)
        R_o_off = R_o * (Mass_h / Inlet_h.Massflow.m) * (Heat_capa_h / Cp_h) ^ 0.33 * (my_h / my_h_off) ^ -0.67 * (tc_h / tc_h_off) ^ 0.67
        
        'U_off 구하기 '''''''''''''''''''''''''''''''''
        U_off = 1 / (R_o_off / Fin_efficiency + f_fin * R_w_off + f_fin * R_i_off)
        NTU = U_off * A_o / C_min
    End If
    
    Y = Effectiveness - (1 - Exp(-NTU))

    If Abs(Y) < Toler_Adj Then
       Exit Do
    ElseIf iter > 100 And Abs(Y) > 0.005 Then
       MsgBox "반복계산 횟수가 100번 이상이 됨."
       Error 1
    ElseIf iter > 100 And Abs(Y) <= 0.005 Then
      Toler_Adj = 0.0051
       
    End If

    Outlet_c.Massflow.m = M_cold + 10 ^ (-6)
    
    Inlet_c.Massflow.m = Outlet_c.Massflow.m + Mass_BD
    Inlet_c.Massflow.Mass_save
    Outlet_c.Massflow.Mass_save
    
    ' Cold side Q
    Q_c = (Outlet_c.Massflow.m * Outlet_c.Properties.h + Mass_BD * Outlet_c_BD.Properties.h) - Inlet_c.Massflow.m * Inlet_c.Properties.h

    Q_h = Q_c / (1 - E_loss)

    ' Hot side property calculation
    delta_P_hot_off = delta_P_hot * ((Inlet_h.Massflow.m / Mass_h) ^ 1.84) * ((Inlet_h.Properties.T + 273.15) / (Temp_h_in + 273.15)) * (Press_h_in / Inlet_h.Properties.P) 'gas side P_loss_off
    Outlet_h.Properties.P = Inlet_h.Properties.P - delta_P_hot_off
    Outlet_h.Properties.h = Inlet_h.Properties.h - Q_h / Inlet_h.Massflow.m
    dummy = idealgas("t", "p", Outlet_h.Properties.P, "h", Outlet_h.Properties.h, Outlet_h.Composition.Composition)
    Outlet_h.Properties.T = dummy(0)
    dummy = idealgas("sck", "p", Outlet_h.Properties.P, "t", Outlet_h.Properties.T, Outlet_h.Composition.Composition)
        Outlet_h.Properties.s = dummy(0)
        Outlet_h.Properties.C = dummy(1)
        Outlet_h.Properties.k = dummy(2)
        Outlet_h.Properties.MW = Inlet_h.Properties.MW
        Outlet_h.Properties.r = Inlet_h.Properties.r
        Outlet_h.Properties.Prop_save
    Temp_h = (Inlet_h.Properties.T + Outlet_h.Properties.T) / 2
    'Circulation'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Inlet_cir.Massflow.m = Inlet_c.Massflow.m * Cir_ratio
    Outlet_cir.Massflow.m = Inlet_cir.Massflow.m
    
    Inlet_cir.Properties.P = Outlet_c.Properties.P
    Inlet_cir.Properties.T = Tsat_p(Inlet_cir.Properties.P)
    Inlet_cir.Properties.h = h_px(Inlet_cir.Properties.P, 0)
    
    Outlet_cir.Properties.P = Inlet_cir.Properties.P
    Outlet_cir.Properties.h = Inlet_cir.Properties.h + Quality * (Outlet_c.Properties.h - Inlet_c.Properties.h)
    Outlet_cir.Properties.T = T_ph(Outlet_cir.Properties.P, Outlet_cir.Properties.h) ' inlet_cir과 같은 온도
    
    'Effectiveness 구하기(Temperature)
    Cp_h = (Outlet_h.Properties.h - Inlet_h.Properties.h) / (Outlet_h.Properties.T - Inlet_h.Properties.T)
    C_h = Cp_h * Outlet_h.Massflow.m
    C_min = C_h ' EVAP에서는 C_h가 C_min이 됨(cold side는 상변화때문에 C_c가 무한대에 가까워짐)
    Effectiveness = (Inlet_h.Properties.T - Outlet_h.Properties.T) / (Inlet_h.Properties.T - Inlet_cir.Properties.T)
    
    If D_o = 0 Then
        tc_h_off = tc_t_gas(Temp_h, Inlet_h.Composition.Composition)
        my_h_off = mu_t_gas(Temp_h, Inlet_h.Composition.Composition)
        UA_off = UA * (Inlet_h.Massflow.m / Mass_h) ^ 0.65 * (Cp_h / Heat_capa_h) ^ 0.33 * (tc_h_off / tc_h) ^ 0.67 * (my_h_off / my_h) ^ -0.32
'        UA_off = UA * (Inlet_h.Massflow.m / Mass_h) ^ 0.8
        NTU = UA_off / C_min
    Else
        'steam_mass_flux''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Mass_flux = Inlet_cir.Massflow.m / (A_tube * n_tube)
        'R_w_off'''''''''''''''''''''''''''''''''''''''''''''''''
        Temp_tube = (Temp_h + Inlet_cir.Properties.T) / 2
        If Temp_tube > 373 Then '오류 제거
            Temp_tube = 373
        End If
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

        'R_i_off 구하기'''''''''''''''''''''''''''''''''''''
        'hsp.l 구하기'''''''''''''''''''
        If Inlet_cir.Properties.P > 16530 Then
             Enthalpy_l = Inlet_cir.Properties.h - 1
             my_l = my_ph(Inlet_cir.Properties.P, Enthalpy_l)
        Else
             my_l = my_ph(Inlet_cir.Properties.P, Inlet_cir.Properties.h)
        End If
        Re = ((1 - Quality) * Mass_flux * D_i) / my_l
      
        Cp_l = CpL_p(Inlet_cir.Properties.P) * 1000
        tc_l = tcL_p(Inlet_cir.Properties.P)
        PR = Cp_l * my_l / tc_l
      
        h_sp = 0.023 * Re ^ 0.8 * PR ^ 0.4 * tc_l / D_i

        'F구하기'''''''''''''''''''''''''''''''''''''''''''''''''''''''
        rho_l = rhoL_p(Inlet_cir.Properties.P)
        rho_g = rhoV_p(Inlet_cir.Properties.P)
        Enthalpy_g = h_px(Inlet_cir.Properties.P, 1)
        If Inlet_cir.Properties.P > 6031 Then
            Enthalpy_g = h_px(6031, 1)
            my_g = my_ph(6031, Enthalpy_g)
        Else
            Enthalpy_g = h_px(Inlet_cir.Properties.P, 1)
            my_g = my_ph(Inlet_cir.Properties.P, Enthalpy_g)
        End If
        X_tt = (Quality / (1 - Quality)) ^ 0.9 * (rho_l / rho_g) ^ 0.5 * (my_g / my_l) ^ 0.1
        coeff_F = 2.35 * (X_tt + 0.213) ^ 0.736

        'hnb 구하기'''''''''''''''''''''''''''''''''''''''''''''''
        st = st_p(Inlet_cir.Properties.P)

        h_lg = (h_px(Inlet_cir.Properties.P, 1) - h_px(Inlet_cir.Properties.P, 0)) * 1000
        Press_tube = psat_T(Temp_tube)
        delta_P = (Press_tube - Inlet_cir.Properties.P) * 1000
        delta_T = Temp_tube - Inlet_cir.Properties.T

        h_nb = (0.00122 * tc_l ^ 0.79 * Cp_l ^ 0.45 * rho_l ^ 0.49 * delta_T ^ 0.24 * delta_P ^ 0.75) / (st ^ 0.5 * my_l ^ 0.29 * h_lg ^ 0.24 * rho_g ^ 0.24)

        'S 구하기'''''''''''''''''''''''''''''''''''''''''''''''''''''''
        coeff_S = 1 / (1 + 2.53 * 10 ^ -6 * (Re * coeff_F ^ 1.25) ^ 1.17)
        R_i_off = D_o / (D_i * 0.001 * (coeff_F * h_sp + coeff_S * h_nb))
        
        'R_o_off 구하기''''''''''''''''''''''''''''''''''''''''''''''
        my_h_off = mu_t_gas(Temp_h, Inlet_h.Composition.Composition)
        tc_h_off = tc_t_gas(Temp_h, Inlet_h.Composition.Composition)
        R_o_off = R_o * (Mass_h / Inlet_h.Massflow.m) * (Heat_capa_h / Cp_h) ^ 0.33 * (my_h / my_h_off) ^ -0.67 * (tc_h / tc_h_off) ^ 0.67
        
        'U_off 구하기 '''''''''''''''''''''''''''''''''
        U_off = 1 / (R_o_off / Fin_efficiency + f_fin * R_w_off + f_fin * R_i_off)
        NTU = U_off * A_o / C_min
    End If
    
    y_for = Effectiveness - (1 - Exp(-NTU))
    
            ' Adaptive relaxation
    If Abs(y_for) > 0.01 Then
        step_size = 0.3  ' Decrease step size if error is large
    ElseIf Abs(y_for) <= 0.01 Then
        step_size = 1  ' Increase step size for faster convergence
    End If

    M_cold = M_cold - step_size * (Y / ((y_for - Y) / 10 ^ (-6)))
    

    
    
    Loop
    
    If BFP_Cal_option = "Press Cal" Then 'Cal inlet pressure (ECON outlet)
    
        'Cold side 입구 압력 계산 Do~Loop''''''''''''''''
        Do
        If delta_P_valve = 0 Then
            delta_P_valve_off = 0
        Else
            delta_P_valve_off = delta_P_valve * ((Inlet_c.Massflow.m / Mass_in_c) ^ 1.98)
        End If
        P_Static = rho_ph(Inlet_c.Properties.P, Inlet_c.Properties.h) * 9.81 * Head / 1000 'kPa
        P_in_cal = Inlet_c.Properties.P
        Inlet_c.Properties.P = Inlet_c_valve.Properties.P + P_Static + delta_P_valve_off
        If Abs(P_in_cal - Inlet_c.Properties.P) < 0.1 Then Exit Do
        Inlet_c.Properties.T = T_ph(Inlet_c.Properties.P, Inlet_c.Properties.h)
        Loop
        
        Stop_cal_value = Inlet_c.Properties.P ''수렴성
    Else
        Stop_cal_value = Inlet_c.Massflow.m
    End If
    
       
  If D_o = 0 Then
    UA = UA_off
  Else
   UA = U_off * A_o
 '  Velocity = 4 * Inlet_c.Massflow.m / (n_tube * (1 / sv_in_c_off) * Application.WorksheetFunction.pi * D_i ^ 2)
  End If
    
Case Else
    MsgBox "EVAP : Case error"

End Select
End Sub







