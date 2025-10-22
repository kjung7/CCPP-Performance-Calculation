VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMIX_"
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
Public Sec As New clsStream_steam_

'HRSG_read'''''''''''''''''
Public e, mode, j, i As Integer
Public name, Component_name As Variant

Public E_loss, P_loss_cold As Variant
Public Method_P_loss_cold As Variant
Public Calc_option, Input_value, Valve As Variant
''''''''''''''''''''''''''''
Public delta_P_cold As Variant
Public delta_P_cold_off As Double

Public Stop_cal_value As Double

Public Mass_in_c, sv_in_c As Double 'off넘기기

Public Check_Mass, Effectiveness, Velocity, UA As Variant
Public BFP_Cal_option As Variant

Sub calMIX()

Select Case mode

Case 0

    Dim dummy As Variant
    If IsEmpty(Inlet_c.Properties.Properties) Then
        'Hotside bypass''''''''''''''''''''''''''''''
        Set Outlet_h.Properties = Inlet_h.Properties
        Set Outlet_h.Massflow = Inlet_h.Massflow
        Set Outlet_h.Composition = Inlet_h.Composition
        Exit Sub
    End If
    
    If IsEmpty(Sec.Properties.h) Then
        'Hotside bypass''''''''''''''''''''''''''''''
        Set Outlet_h.Properties = Inlet_h.Properties
        Set Outlet_h.Massflow = Inlet_h.Massflow
        Set Outlet_h.Composition = Inlet_h.Composition
        Exit Sub
    End If
    
    ''hot side bypass''''''''''''''''''''''''
    Set Outlet_h.Properties = Inlet_h.Properties
    Set Outlet_h.Massflow = Inlet_h.Massflow
    Set Outlet_h.Composition = Inlet_h.Composition
    
    'Sec steam
    If Valve = "Yes" Then 'valve가 있을 경우만 secondary flow 물성치 다시 계산
        Sec.Properties.T = T_ph(Sec.Properties.P, Sec.Properties.h)
        Sec.Properties.s = s_pT(Sec.Properties.P, Sec.Properties.T)
        Sec.Properties.cp = Cp_pT(Sec.Properties.P, Sec.Properties.T)
        Sec.Properties.Cv = Cv_pT(Sec.Properties.P, Sec.Properties.T)
        Sec.Properties.k = Sec.Properties.cp / Sec.Properties.Cv
        Sec.Properties.x = x_ph(Sec.Properties.P, Sec.Properties.h)
        Sec.Properties.Prop_save
    End If
    
    Select Case Calc_option
    Case "Outlet temperature", "Temperature change"

  ''outlet temperature
    If Calc_option = "Outlet temperature" Then 'User input : Steam outlet temperature
        Outlet_c.Properties.T = Input_value
    ElseIf Calc_option = "Temperature change" Then 'User input : Temperature change
        Outlet_c.Properties.T = Inlet_c.Properties.T - Input_value
    End If

  ''outlet properties
    If Outlet_c.Properties.T = Inlet_c.Properties.T And P_loss_cold = 0 Then
        Set Outlet_c.Properties = Inlet_c.Properties
        delta_P_cold = 0
    Else
        If Method_P_loss_cold = "P_loss_cold" Then
            Outlet_c.Properties.P = Inlet_c.Properties.P * (1 - P_loss_cold)
            delta_P_cold = Inlet_c.Properties.P - Outlet_c.Properties.P
        ElseIf Method_P_loss_cold = "delta_P_cold" Then
            delta_P_cold = P_loss_cold
            Outlet_c.Properties.P = Inlet_c.Properties.P - P_loss_cold
        End If
        Outlet_c.Properties.h = h_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
        Outlet_c.Properties.s = s_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
        Outlet_c.Properties.cp = Cp_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
        Outlet_c.Properties.Cv = Cv_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
        Outlet_c.Properties.k = Outlet_c.Properties.cp / Outlet_c.Properties.Cv
        Outlet_c.Properties.x = x_ph(Outlet_c.Properties.P, Outlet_c.Properties.h)
        Outlet_c.Properties.Prop_save
    End If

    If i < e Then
        Sec.Massflow.m = Inlet_c.Massflow.m * (Inlet_c.Properties.h - Outlet_c.Properties.h / (1 - E_loss)) / (Outlet_c.Properties.h / (1 - E_loss) - Sec.Properties.h)
        Sec.Massflow.Mass_save
        'outlet massflow
        Outlet_c.Massflow.m = Inlet_c.Massflow.m + Sec.Massflow.m
        Outlet_c.Massflow.Mass_save
    Else
        If IsEmpty(Outlet_c.Massflow.m) Then
            Sec.Massflow.m = Inlet_c.Massflow.m * (Inlet_c.Properties.h - Outlet_c.Properties.h / (1 - E_loss)) / (Outlet_c.Properties.h / (1 - E_loss) - Sec.Properties.h)
            Sec.Massflow.Mass_save
            Outlet_c.Massflow.m = Inlet_c.Massflow.m + Sec.Massflow.m
            Outlet_c.Massflow.Mass_save
        Else
            Sec.Massflow.m = Outlet_c.Massflow.m * (Outlet_c.Properties.h / (1 - E_loss) - Inlet_c.Properties.h) / (Sec.Properties.h - Inlet_c.Properties.h)
            Sec.Massflow.Mass_save
            Inlet_c.Massflow.m = Outlet_c.Massflow.m - Sec.Massflow.m
            Inlet_c.Massflow.Mass_save
        End If
    End If
    Sec.Massflow.Mass_save
    
    Case "Receive massflow" 'Sec massflow가 input인 경우
    
    If Sec.Massflow.m = 0 And P_loss_cold = 0 And E_loss = 0 Then
        If i < e Or IsEmpty(Outlet_c.Massflow.m) Then
            Outlet_c.Massflow.m = Inlet_c.Massflow.m
        Else
            Inlet_c.Massflow.m = Outlet_c.Massflow.m
        End If
        Set Outlet_c.Properties = Inlet_c.Properties
        delta_P_cold = 0
    Else
        If i < e Or IsEmpty(Outlet_c.Massflow.m) Then
            Outlet_c.Massflow.m = Inlet_c.Massflow.m + Sec.Massflow.m
            Outlet_c.Massflow.Mass_save
        Else
            Inlet_c.Massflow.m = Outlet_c.Massflow.m - Sec.Massflow.m
        End If
        'outlet properties
        If Method_P_loss_cold = "P_loss_cold" Then
            Outlet_c.Properties.P = Inlet_c.Properties.P * (1 - P_loss_cold)
            delta_P_cold = Inlet_c.Properties.P - Outlet_c.Properties.P
        ElseIf Method_P_loss_cold = "delta_P_cold" Then
            delta_P_cold = P_loss_cold
            Outlet_c.Properties.P = Inlet_c.Properties.P - P_loss_cold
        End If
        Outlet_c.Properties.h = (Inlet_c.Massflow.m * Inlet_c.Properties.h + Sec.Massflow.m * Sec.Properties.h) * (1 - E_loss) / Outlet_c.Massflow.m
        Outlet_c.Properties.T = T_ph(Outlet_c.Properties.P, Outlet_c.Properties.h)
'        Outlet_c.Properties.s = s_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
'        Outlet_c.Properties.Cp = Cp_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
'        Outlet_c.Properties.Cv = Cv_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
'        Outlet_c.Properties.k = Outlet_c.Properties.Cp / Outlet_c.Properties.Cv
'        Outlet_c.Properties.x = x_ph(Outlet_c.Properties.P, Outlet_c.Properties.h)
        Outlet_c.Properties.Prop_save
    End If
    
    End Select
    Stop_cal_value = Outlet_c.Properties.T ''수렴성

    'delta_P_off''''''''''''''''''''''''''''''''''''''''''''''''''
    Mass_in_c = Inlet_c.Massflow.m
    If i < e Then 'EVAP 앞에 있는 MIX
        If Inlet_c.Properties.T = Tsat_p(Inlet_c.Properties.P) Then
            sv_in_c = vV_p(Inlet_c.Properties.P)
        Else
            sv_in_c = v_pT(Inlet_c.Properties.P, Inlet_c.Properties.T)
        End If
    End If
    
Case 1
    
    Dim P_in_cal As Double
    Dim sv_in_c_off As Double
    
    ''hot side bypass''''''''''''''''''''''''
    Set Outlet_h.Properties = Inlet_h.Properties
    Set Outlet_h.Massflow = Inlet_h.Massflow
    Set Outlet_h.Composition = Inlet_h.Composition
    
    Select Case Calc_option
    Case "Outlet temperature", "Temperature change"

  ''outlet temperature
    If Calc_option = "Outlet temperature" Then 'User input : Steam outlet temperature
        Outlet_c.Properties.T = Input_value
    ElseIf Calc_option = "Temperature change" Then 'User input : Temperature change
        Outlet_c.Properties.T = Inlet_c.Properties.T - Input_value
    End If
    
    If i < e Or BFP_Cal_option = "Press Cal" Then 'EVAP 앞의 MIX이거나 Press Cal인 경우 -> Cold side 입구 압력 계산
        If delta_P_cold = 0 Then
            Inlet_c.Properties.P = Outlet_c.Properties.P
        Else
            If i < e Then
                Do 'Cold side 입구 압력 계산 Do~Loop
                    If Inlet_c.Properties.T = Tsat_p(Inlet_c.Properties.P) Then
                        sv_in_c_off = vV_p(Inlet_c.Properties.P) 'EVAP 앞에 있는 SPL는 steam
                    Else
                        sv_in_c_off = v_pT(Inlet_c.Properties.P, Inlet_c.Properties.T)
                    End If
                    delta_P_cold_off = delta_P_cold * ((Inlet_c.Massflow.m / Mass_in_c) ^ 1.98) * (sv_in_c_off / sv_in_c)
                P_in_cal = Inlet_c.Properties.P
                Inlet_c.Properties.P = Outlet_c.Properties.P + delta_P_cold_off
    
                If Abs(P_in_cal - Inlet_c.Properties.P) < 1 Then Exit Do
                Inlet_c.Properties.T = T_ph(Inlet_c.Properties.P, Inlet_c.Properties.h)
                Loop
            Else
                delta_P_cold_off = delta_P_cold * ((Inlet_c.Massflow.m / Mass_in_c) ^ 1.98)
                Inlet_c.Properties.P = Outlet_c.Properties.P + delta_P_cold_off
            End If
        End If
        Inlet_c.Properties.T = T_ph(Inlet_c.Properties.P, Inlet_c.Properties.h)
    End If
    
    ''Sec steam
    If Valve = "Yes" Then 'Valve "Yes"
        Sec.Properties.P = Inlet_c.Properties.P
        Sec.Properties.T = T_ph(Sec.Properties.P, Sec.Properties.h)
        Sec.Properties.s = s_pT(Sec.Properties.P, Sec.Properties.T)
        Sec.Properties.cp = Cp_pT(Sec.Properties.P, Sec.Properties.T)
        Sec.Properties.Cv = Cv_pT(Sec.Properties.P, Sec.Properties.T)
        Sec.Properties.k = Sec.Properties.cp / Sec.Properties.Cv
        Sec.Properties.x = x_ph(Sec.Properties.P, Sec.Properties.h)
        Sec.Properties.Prop_save
    End If
    
    'delta_P_cold_off
    If delta_P_cold = 0 Then
        delta_P_cold_off = 0
    Else
        If i < e Then
            If Inlet_c.Properties.T = Tsat_p(Inlet_c.Properties.P) Then
                sv_in_c_off = vV_p(Inlet_c.Properties.P)
            Else
                sv_in_c_off = v_pT(Inlet_c.Properties.P, Inlet_c.Properties.T)
            End If
            delta_P_cold_off = delta_P_cold * ((Inlet_c.Massflow.m / Mass_in_c) ^ 1.98) * (sv_in_c_off / sv_in_c)
        Else
            delta_P_cold_off = delta_P_cold * ((Inlet_c.Massflow.m / Mass_in_c) ^ 1.98)
        End If
    End If
    
  ''outlet properties
    If Outlet_c.Properties.T = Inlet_c.Properties.T And delta_P_cold_off = 0 Then
        Set Outlet_c.Properties = Inlet_c.Properties
    Else
        Outlet_c.Properties.P = Inlet_c.Properties.P - delta_P_cold_off
        Outlet_c.Properties.h = h_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
        Outlet_c.Properties.s = s_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
        Outlet_c.Properties.cp = Cp_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
        Outlet_c.Properties.Cv = Cv_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
        Outlet_c.Properties.k = Outlet_c.Properties.cp / Outlet_c.Properties.Cv
        Outlet_c.Properties.x = x_ph(Outlet_c.Properties.P, Outlet_c.Properties.h)
        Outlet_c.Properties.Prop_save
    End If

    If i < e Then
        Sec.Massflow.m = Inlet_c.Massflow.m * (Inlet_c.Properties.h - Outlet_c.Properties.h / (1 - E_loss)) / (Outlet_c.Properties.h / (1 - E_loss) - Sec.Properties.h)
        Sec.Massflow.Mass_save
        Outlet_c.Massflow.m = Inlet_c.Massflow.m + Sec.Massflow.m
        Outlet_c.Massflow.Mass_save
    Else
        Sec.Massflow.m = Outlet_c.Massflow.m * (Outlet_c.Properties.h / (1 - E_loss) - Inlet_c.Properties.h) / (Sec.Properties.h - Inlet_c.Properties.h)
        Sec.Massflow.Mass_save
        Inlet_c.Massflow.m = Outlet_c.Massflow.m - Sec.Massflow.m
        Inlet_c.Massflow.Mass_save
    End If
    Sec.Massflow.Mass_save
    Stop_cal_value = Sec.Massflow.m ''수렴성
    
    Case "Receive massflow" 'Sec massflow가 input인 경우
    
    If i < e Then
        Outlet_c.Massflow.m = Inlet_c.Massflow.m + Sec.Massflow.m                   'If the component is ahead of EVAP, Outlet Mass = Inlet Mass + Secondary Mass
        Outlet_c.Massflow.Mass_save
    Else
        Inlet_c.Massflow.m = Outlet_c.Massflow.m - Sec.Massflow.m                   'If the component is behind of EVAP, Inlet Mass = Outlet Mass - Secondary Mass

    End If
    'EVAP 앞의 MIX이거나 Press Cal인 경우 -> Cold side 입구 압력 계산
    If i < e Or BFP_Cal_option = "Press Cal" Then
        If delta_P_cold = 0 Then
            Inlet_c.Properties.P = Outlet_c.Properties.P
        Else
            If name = "IP_MIX_RH" Then
                sv_in_c_off = v_pT(Inlet_c.Properties.P, Inlet_c.Properties.T)
                
                If Mass_in_c = 0 Then 'Mass_in_c가 0인 경우 추가
                    delta_P_cold_off = delta_P_cold
                Else
                    delta_P_cold_off = delta_P_cold * ((Inlet_c.Massflow.m / Mass_in_c) ^ 1.98) * (sv_in_c_off / sv_in_c)
                
                End If
                If delta_P_cold_off = 0 Then 'inlet_c mass=0 인 경우
                    delta_P_cold_off = delta_P_cold
                Else
                    If delta_P_cold_off > delta_P_cold And delta_P_cold / delta_P_cold_off < 0.55 Then delta_P_cold_off = delta_P_cold / 0.55
                    If delta_P_cold_off < delta_P_cold And delta_P_cold_off / delta_P_cold < 0.55 Then delta_P_cold_off = delta_P_cold * 0.55
                End If
                Inlet_c.Properties.P = Outlet_c.Properties.P + delta_P_cold_off
            ElseIf i < e Then
                Do 'Cold side 입구 압력 계산 Do~Loop
                    If Inlet_c.Properties.T = Tsat_p(Inlet_c.Properties.P) Then
                        sv_in_c_off = vV_p(Inlet_c.Properties.P) 'EVAP 앞에 있는 SPL는 steam
                    Else
                        sv_in_c_off = v_pT(Inlet_c.Properties.P, Inlet_c.Properties.T)
                    End If
                    delta_P_cold_off = delta_P_cold * ((Inlet_c.Massflow.m / Mass_in_c) ^ 1.98) * (sv_in_c_off / sv_in_c)
                P_in_cal = Inlet_c.Properties.P
                Inlet_c.Properties.P = Outlet_c.Properties.P + delta_P_cold_off
    
                If Abs(P_in_cal - Inlet_c.Properties.P) < 1 Then Exit Do
                Inlet_c.Properties.T = T_ph(Inlet_c.Properties.P, Inlet_c.Properties.h)
                Loop
            Else
                delta_P_cold_off = delta_P_cold * ((Inlet_c.Massflow.m / Mass_in_c) ^ 1.98)
                Inlet_c.Properties.P = Outlet_c.Properties.P + delta_P_cold_off
            End If
        End If
        Inlet_c.Properties.T = T_ph(Inlet_c.Properties.P, Inlet_c.Properties.h)
    End If
    
   ''Sec steam
    If Valve = "Yes" Then 'Valve "Yes"
        Sec.Properties.P = Inlet_c.Properties.P
        Sec.Properties.T = T_ph(Sec.Properties.P, Sec.Properties.h)
        Sec.Properties.s = s_pT(Sec.Properties.P, Sec.Properties.T)
        Sec.Properties.cp = Cp_pT(Sec.Properties.P, Sec.Properties.T)
        Sec.Properties.Cv = Cv_pT(Sec.Properties.P, Sec.Properties.T)
        Sec.Properties.k = Sec.Properties.cp / Sec.Properties.Cv
        Sec.Properties.x = x_ph(Sec.Properties.P, Sec.Properties.h)
        Sec.Properties.Prop_save
    End If
    
'    'delta_P_cold_off
    If delta_P_cold = 0 Then
        delta_P_cold_off = 0
    Else
        If name = "IP_MIX_RH" Then
            sv_in_c_off = v_pT(Inlet_c.Properties.P, Inlet_c.Properties.T)
            If Mass_in_c = 0 Then 'Mass_in_c가 0인 경우 추가
                delta_P_cold_off = delta_P_cold
            Else
                delta_P_cold_off = delta_P_cold * ((Inlet_c.Massflow.m / Mass_in_c) ^ 1.98) * (sv_in_c_off / sv_in_c)
            End If
            If delta_P_cold_off = 0 Then
                delta_P_cold_off = delta_P_cold
            Else
                If delta_P_cold_off > delta_P_cold And delta_P_cold / delta_P_cold_off < 0.55 Then delta_P_cold_off = delta_P_cold / 0.55
                If delta_P_cold_off < delta_P_cold And delta_P_cold_off / delta_P_cold < 0.55 Then delta_P_cold_off = delta_P_cold * 0.55
            End If
            Inlet_c.Properties.P = Outlet_c.Properties.P + delta_P_cold_off
        Else
        If i < e Then
            If Inlet_c.Properties.T = Tsat_p(Inlet_c.Properties.P) Then
                sv_in_c_off = vV_p(Inlet_c.Properties.P)
            Else
                sv_in_c_off = v_pT(Inlet_c.Properties.P, Inlet_c.Properties.T)
            End If
            delta_P_cold_off = delta_P_cold * ((Inlet_c.Massflow.m / Mass_in_c) ^ 1.98) * (sv_in_c_off / sv_in_c)
        Else
            delta_P_cold_off = delta_P_cold * ((Inlet_c.Massflow.m / Mass_in_c) ^ 1.98)
        End If
        End If
    End If
    
    If Sec.Massflow.m = 0 And delta_P_cold_off = 0 And E_loss = 0 Then
        Set Outlet_c.Properties = Inlet_c.Properties
    Else
        Outlet_c.Properties.P = Inlet_c.Properties.P - delta_P_cold_off
        Outlet_c.Properties.h = (Inlet_c.Massflow.m * Inlet_c.Properties.h + Sec.Massflow.m * Sec.Properties.h) * (1 - E_loss) / Outlet_c.Massflow.m

        Outlet_c.Properties.T = T_ph(Outlet_c.Properties.P, Outlet_c.Properties.h)

'        Outlet_c.Properties.s = s_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
'        Outlet_c.Properties.Cp = Cp_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
'        Outlet_c.Properties.Cv = Cv_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
'        Outlet_c.Properties.k = Outlet_c.Properties.Cp / Outlet_c.Properties.Cv
'        Outlet_c.Properties.x = x_ph(Outlet_c.Properties.P, Outlet_c.Properties.h)
        Outlet_c.Properties.Prop_save
    End If
    
    If BFP_Cal_option = "Press Cal" Then
        Stop_cal_value = Inlet_c.Properties.P
    Else
        Stop_cal_value = Outlet_c.Properties.T
    End If

    End Select
    

End Select

End Sub

    






