VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGLD_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Inlet_h As New clsStream_
Public Outlet_h As New clsStream_
Public Inlet_c As New clsStream_steam_
Public Outlet_c As New clsStream_steam_
Public Inlet_GLD As New clsStream_steam_
Public Outlet_GLD As New clsStream_steam_
'HRSG_read'''''''''''''''''
Public mode, j As Integer
Public name As Variant

Public E_loss, P_loss_cold As Double
Public Method_P_loss_cold As Variant
Public Calc_option_P_GLD, Input_value_P_GLD As Variant
''''''''''''''''''''''''''''
Public delta_P_cold As Double
Public delta_P_cold_off As Double

Private Q_c, Q_GLD As Double
Private Cp_c, Cp_GLD, C_c, C_GLD As Double
Public C_min, C_r As Double
Public Press_c, Temp_c As Double

Public Stop_cal_value As Double
Public Mass_c, T_c_assume As Double

Public Effectiveness, UA, Velocity As Double
Public BFP_Cal_option As Variant

Sub calGLD()

Select Case mode

Case 0

    Dim dummy As Variant
    Dim del_T_2, del_T_1, LMTD As Double
    
    Dim iter As Integer
    Dim Y, y_for As Double
    
'     Cold side inlet의 stream 이 연결되지 않을 경우 hot side stream bypass
'     동일 pressure level의 ECON 계산 완료 이후에 stream이 연결됨
    If IsEmpty(Inlet_c.Properties.Properties) Then
        'Hotside bypass''''''''''''''''''''''''''''''
        Set Outlet_h.Properties = Inlet_h.Properties
        Set Outlet_h.Massflow = Inlet_h.Massflow
        Set Outlet_h.Composition = Inlet_h.Composition
        Exit Sub
        
    End If
    
    'Hotside bypass
    Set Outlet_h.Massflow = Inlet_h.Massflow
    Set Outlet_h.Properties = Inlet_h.Properties
    Set Outlet_h.Composition = Inlet_h.Composition
    
    Set Outlet_c.Massflow = Inlet_c.Massflow
    Set Outlet_GLD.Massflow = Inlet_GLD.Massflow
    
    'GLD hot side properties calculation
    Inlet_GLD.Properties.h = h_pT(Inlet_GLD.Properties.P, Inlet_GLD.Properties.T)
    If Calc_option_P_GLD = "No pressure drop" Or Calc_option_P_GLD = "Operating pressure" Then
        Outlet_GLD.Properties.P = Input_value_P_GLD * 100
    ElseIf Calc_option_P_GLD = "Pressure drop" Then
        Outlet_GLD.Properties.P = Inlet_GLD.Properties.P * (1 - Input_value_P_GLD) ''check
    End If
    Outlet_GLD.Properties.T = Tsat_p(Outlet_GLD.Properties.P)
    Outlet_GLD.Properties.h = h_px(Outlet_GLD.Properties.P, 0)
    Q_GLD = Outlet_GLD.Massflow.m * (Inlet_GLD.Properties.h - Outlet_GLD.Properties.h)
    ' GLD cold side properties calculation
    Q_c = Q_GLD
    Outlet_c.Properties.h = Inlet_c.Properties.h + Q_c / Outlet_c.Massflow.m
    If Method_P_loss_cold = "P_loss_cold" Then
        Outlet_c.Properties.P = Inlet_c.Properties.P * (1 - P_loss_cold)
        delta_P_cold = Inlet_c.Properties.P - Outlet_c.Properties.P
    ElseIf Method_P_loss_cold = "delta_P_cold" Then
        delta_P_cold = P_loss_cold
        Outlet_c.Properties.P = Inlet_c.Properties.P - P_loss_cold
    End If
    Outlet_c.Properties.T = T_ph(Outlet_c.Properties.P, Outlet_c.Properties.h)
    Outlet_c.Properties.s = s_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
    Outlet_c.Properties.cp = Cp_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
    Outlet_c.Properties.Cv = Cv_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
    Outlet_c.Properties.k = Outlet_c.Properties.cp / Outlet_c.Properties.Cv
    Outlet_c.Properties.x = x_ph(Outlet_c.Properties.P, Outlet_c.Properties.h)
    Outlet_c.Properties.Prop_save

    Stop_cal_value = Outlet_c.Properties.T
    
    'Off-design의 reference'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    T_c_assume = Outlet_c.Properties.T
    Mass_c = Outlet_c.Massflow.m

'OFF-DESIGN''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Case 1

    'Hotside bypass
    Set Outlet_h.Massflow = Inlet_h.Massflow
    Set Outlet_h.Properties = Inlet_h.Properties
    Set Outlet_h.Composition = Inlet_h.Composition

    Set Outlet_c.Massflow = Inlet_c.Massflow
    Set Outlet_GLD.Massflow = Inlet_GLD.Massflow
    
    'GLD hot side properties calculation
    Inlet_GLD.Properties.h = h_pT(Inlet_GLD.Properties.P, Inlet_GLD.Properties.T)
    If Calc_option_P_GLD = "No pressure drop" Or Calc_option_P_GLD = "Operating pressure" Then
        Outlet_GLD.Properties.P = Input_value_P_GLD * 100
    ElseIf Calc_option_P_GLD = "Pressure drop" Then
        Outlet_GLD.Properties.P = Inlet_GLD.Properties.P * (1 - Input_value_P_GLD)
    End If
    Outlet_GLD.Properties.T = Tsat_p(Outlet_GLD.Properties.P)
    Outlet_GLD.Properties.h = h_px(Outlet_GLD.Properties.P, 0)
    Q_GLD = Outlet_GLD.Massflow.m * (Inlet_GLD.Properties.h - Outlet_GLD.Properties.h)
     
    If delta_P_cold = 0 Then
        delta_P_cold_off = 0
    Else
        delta_P_cold_off = delta_P_cold * ((Outlet_c.Massflow.m / Mass_c) ^ 1.98)
    End If

    If BFP_Cal_option = "Press Cal" Then
        Inlet_c.Properties.P = Outlet_c.Properties.P + delta_P_cold_off
        Inlet_c.Properties.T = T_ph(Inlet_c.Properties.P, Inlet_c.Properties.h)
    End If
    
    ' GLD cold side properties calculation
    Q_c = Q_GLD
    Outlet_c.Properties.h = Inlet_c.Properties.h + Q_c / Outlet_c.Massflow.m
    Outlet_c.Properties.P = Inlet_c.Properties.P - delta_P_cold_off
    ' Cold side Q 계산
    Outlet_c.Properties.T = T_ph(Outlet_c.Properties.P, Outlet_c.Properties.h)
    Outlet_c.Properties.s = s_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
    Outlet_c.Properties.cp = Cp_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
    Outlet_c.Properties.Cv = Cv_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
    Outlet_c.Properties.k = Outlet_c.Properties.cp / Outlet_c.Properties.Cv
    Outlet_c.Properties.x = x_ph(Outlet_c.Properties.P, Outlet_c.Properties.h)
    Outlet_c.Properties.Prop_save

    If BFP_Cal_option = "Press Cal" Then
        Stop_cal_value = Inlet_c.Properties.P
    Else
        Stop_cal_value = Outlet_c.Properties.T
    End If
    
 End Select

End Sub





