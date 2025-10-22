VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSPL_"
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
Public Ter As New clsStream_steam_

'HRSG_read'''''''''''''''''
Public e, mode, j, i As Integer
Public name, Pressure_level As Variant

Public E_loss, P_loss_cold As Variant
Public Method_P_loss_cold As Variant
Public Calc_option, Input_value As Variant
''''''''''''''''''''''''''''
Public Check As Double

Public delta_P_cold As Variant
Public delta_P_cold_off As Double

Public Stop_cal_value As Double

Public Mass_in_c, sv_in_c As Double 'off넘기기

Public Check_Mass, Effectiveness, Velocity, UA As Variant
Public BFP_Cal_option As Variant

Sub calSPL()

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
    
    ''hot side bypass''''''''''''''''''''''''
    Set Outlet_h.Properties = Inlet_h.Properties
    Set Outlet_h.Massflow = Inlet_h.Massflow
    Set Outlet_h.Composition = Inlet_h.Composition
    
    ''outlet massflow
    If Calc_option = "DA Pegging STM" And Check = 1 Then 'Ter=0
        Sec.Massflow.m = Inlet_c.Massflow.m            'Whole steam flow to DA
        Sec.Massflow.Mass_save
        Outlet_c.Massflow.m = 0
        Outlet_c.Massflow.Mass_save
        Set Sec.Properties = Inlet_c.Properties
        Exit Sub
    End If
    
    If Pressure_level = "RH" Or IsEmpty(Outlet_c.Massflow.m) Then              'RH_SPL
        Outlet_c.Massflow.m = Inlet_c.Massflow.m - Sec.Massflow.m - Ter.Massflow.m
    Else
        If i < e Then                                                          'If it is downstream of Evaporator
            Outlet_c.Massflow.m = Inlet_c.Massflow.m - Sec.Massflow.m - Ter.Massflow.m         'Outlet Mass Flow = Inlet - Secondary Flow
                        If Outlet_c.Massflow.m < 0 Then Sec.Massflow.m = Inlet_c.Massflow.m - Ter.Massflow.m
            If Outlet_c.Massflow.m < 0 Then Outlet_c.Massflow.m = Inlet_c.Massflow.m - Sec.Massflow.m - Ter.Massflow.m
        Else
            Inlet_c.Massflow.m = Outlet_c.Massflow.m + Sec.Massflow.m + Ter.Massflow.m         'Inlet = Outlet + Secondary
        End If
    End If
    
    If E_loss = 0 And P_loss_cold = 0 Then
        Set Outlet_c.Properties = Inlet_c.Properties
        delta_P_cold = 0
    Else
        'outlet properties
        If Method_P_loss_cold = "P_loss_cold" Then
            Outlet_c.Properties.P = Inlet_c.Properties.P * (1 - P_loss_cold)
            delta_P_cold = Inlet_c.Properties.P - Outlet_c.Properties.P
        ElseIf Method_P_loss_cold = "delta_P_cold" Then
            delta_P_cold = P_loss_cold
            Outlet_c.Properties.P = Inlet_c.Properties.P - P_loss_cold
        End If
        
        Outlet_c.Properties.h = ((Inlet_c.Massflow.m * Inlet_c.Properties.h) * (1 - E_loss) - Sec.Massflow.m * Sec.Properties.h - Ter.Massflow.m * Ter.Properties.h) / Outlet_c.Massflow.m
        Outlet_c.Properties.T = T_ph(Outlet_c.Properties.P, Outlet_c.Properties.h)
        Outlet_c.Properties.s = s_ph(Outlet_c.Properties.P, Outlet_c.Properties.h)
        Outlet_c.Properties.cp = Cp_ph(Outlet_c.Properties.P, Outlet_c.Properties.h)
        Outlet_c.Properties.Cv = Cv_ph(Outlet_c.Properties.P, Outlet_c.Properties.h)
        Outlet_c.Properties.k = Outlet_c.Properties.cp / Outlet_c.Properties.Cv
        Outlet_c.Properties.x = x_ph(Outlet_c.Properties.P, Outlet_c.Properties.h)
    End If
    Outlet_c.Properties.Prop_save
    Stop_cal_value = Outlet_c.Properties.T ''수렴성
    
    'delta_P_loss'''''''''''''''''''''''''''''''''''''''''''''''
    Mass_in_c = Inlet_c.Massflow.m
    If Pressure_level = "RH" Then
        sv_in_c = v_pT(Inlet_c.Properties.P, Inlet_c.Properties.T)
    Else
        If i < e Then
            If Inlet_c.Properties.T = Tsat_p(Inlet_c.Properties.P) Then
                sv_in_c = vV_p(Inlet_c.Properties.P) 'EVAP 앞 SPL steam
            Else
                sv_in_c = v_pT(Inlet_c.Properties.P, Inlet_c.Properties.T)
            End If
        End If
    End If

Case 1

    
    Dim P_in_cal As Double
    Dim sv_in_c_off As Double
    
    ''hot side bypass''''''''''''''''''''''''
    Set Outlet_h.Properties = Inlet_h.Properties
    Set Outlet_h.Massflow = Inlet_h.Massflow
    Set Outlet_h.Composition = Inlet_h.Composition
    
    If Calc_option = "DA Pegging STM" And Check = 1 Then 'inlet_c=sec
        Sec.Massflow.m = Inlet_c.Massflow.m              'Whole mass flow going out to deaerator pegging steam
        Sec.Massflow.Mass_save
        Outlet_c.Massflow.m = 0
        Outlet_c.Massflow.Mass_save
        Set Outlet_c.Properties = Inlet_c.Properties
        Set Sec.Properties = Inlet_c.Properties
        Exit Sub
    End If
    
    If Pressure_level = "RH" Then                                              'RH_SPL
        Outlet_c.Massflow.m = Inlet_c.Massflow.m - Sec.Massflow.m - Ter.Massflow.m
    Else
        If i < e Then                                                          'When Splitter is placed ahead of EVAP, Oulet Flow = Inlet Mass Flow - Secondary Flow
            Outlet_c.Massflow.m = Inlet_c.Massflow.m - Sec.Massflow.m - Ter.Massflow.m
            If Outlet_c.Massflow.m < 0 Then Sec.Massflow.m = Inlet_c.Massflow.m - Ter.Massflow.m
            If Outlet_c.Massflow.m < 0 Then Outlet_c.Massflow.m = Inlet_c.Massflow.m - Sec.Massflow.m - Ter.Massflow.m
        Else
            Inlet_c.Massflow.m = Outlet_c.Massflow.m + Sec.Massflow.m + Ter.Massflow.m         'When Splitter is placed behind of EVAP, Inlet Mass Flow = Outlet Flow + Secondary Flow
        End If
    End If
    
    'Cold side 입구 압력''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Not Pressure_level = "RH" Then
        If i < e Or BFP_Cal_option = "Press Cal" Then 'EVAP 앞의SPL이거나 Press Cal인 경우 -> Cold side 입구 압력 계산
            If delta_P_cold = 0 Then
                Inlet_c.Properties.P = Outlet_c.Properties.P
                Inlet_c.Properties.T = T_ph(Inlet_c.Properties.P, Inlet_c.Properties.h)
            Else
                'Cold side 입구 압력 계산 Do~Loop''''''''''''''''
                If i < e Then
                    Do
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
                    Inlet_c.Properties.T = T_ph(Inlet_c.Properties.P, Inlet_c.Properties.h)
                End If
                '''''''''''''''''''''''''''''''''''''''''''''''''
            End If
        End If
    End If
    
    Set Sec.Properties = Inlet_c.Properties
    Set Ter.Properties = Inlet_c.Properties
    
    'outlet properties
    If delta_P_cold = 0 Then
        Set Outlet_c.Properties = Inlet_c.Properties
    Else
        If Pressure_level = "RH" Then
            sv_in_c_off = v_pT(Inlet_c.Properties.P, Inlet_c.Properties.T)
            delta_P_cold_off = delta_P_cold * ((Inlet_c.Massflow.m / Mass_in_c) ^ 1.98) * (sv_in_c_off / sv_in_c)
        Else
            If i < e Then
                If Inlet_c.Properties.T = Tsat_p(Inlet_c.Properties.P) Then
                    sv_in_c_off = vV_p(Inlet_c.Properties.P) 'EVAP 앞 SPL steam
                Else
                    sv_in_c_off = v_pT(Inlet_c.Properties.P, Inlet_c.Properties.T)
                End If
                delta_P_cold_off = delta_P_cold * ((Inlet_c.Massflow.m / Mass_in_c) ^ 1.98) * (sv_in_c_off / sv_in_c)
            Else
                delta_P_cold_off = delta_P_cold * ((Inlet_c.Massflow.m / Mass_in_c) ^ 1.98)
            End If
        End If
        Outlet_c.Properties.P = Inlet_c.Properties.P - delta_P_cold_off
        Outlet_c.Properties.h = ((Inlet_c.Massflow.m * Inlet_c.Properties.h) * (1 - E_loss) - Sec.Massflow.m * Sec.Properties.h - Ter.Massflow.m * Ter.Properties.h) / Outlet_c.Massflow.m
        Outlet_c.Properties.T = T_ph(Outlet_c.Properties.P, Outlet_c.Properties.h)
        Outlet_c.Properties.s = s_ph(Outlet_c.Properties.P, Outlet_c.Properties.h)
'       Outlet_c.Properties.Cp = Cp_ph(Outlet_c.Properties.P, Outlet_c.Properties.h)
'       Outlet_c.Properties.Cv = Cv_ph(Outlet_c.Properties.P, Outlet_c.Properties.h)
'       Outlet_c.Properties.k = Outlet_c.Properties.Cp / Outlet_c.Properties.Cv
'       Outlet_c.Properties.x = x_ph(Outlet_c.Properties.P, Outlet_c.Properties.h)
        Outlet_c.Properties.Prop_save

    End If
    
    If BFP_Cal_option = "Press Cal" Then
        Stop_cal_value = Inlet_c.Properties.P
    Else
        Stop_cal_value = Outlet_c.Properties.T
    End If

End Select

End Sub

    







