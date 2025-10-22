VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDESH_"
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
Public mode, j As Integer
Public name As Variant

Public E_loss, P_loss_valve, P_loss_cold As Double
Public Method_P_loss_valve, Method_P_loss_cold As Variant
Public Calc_option, Input_value, Valve As Variant
''''''''''''''''''''''''''''
Public delta_P_valve, delta_P_cold As Variant
Public delta_P_valve_off, delta_P_cold_off As Variant

Public Stop_cal_value As Double

Public Mass_in_c, sv_in_c, sv_in_c_valve As Double 'off넘기기

Public Check_Mass, Effectiveness, Velocity, UA As Variant

Sub calDESH()

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
    
        Outlet_h.Properties.T = Inlet_h.Properties.T
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
    Case "Outlet temperature", "Temperature change", "Degrees of superheat"

  ''outlet temperature
    If Calc_option = "Outlet temperature" Then 'User input : Steam outlet temperature
        Outlet_c.Properties.T = Input_value
    ElseIf Calc_option = "Temperature change" Then 'User input : Temperature change
        Outlet_c.Properties.T = Inlet_c.Properties.T - Input_value
    ElseIf Calc_option = "Degrees of superheat" Then 'User input : Degrees of superheat
        Outlet_c.Properties.T = Tsat_p(Inlet_c.Properties.P) + Input_value
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
    
    'Sec massflow
    Sec.Massflow.m = Inlet_c.Massflow.m * (Inlet_c.Properties.h - Outlet_c.Properties.h / (1 - E_loss)) / (Outlet_c.Properties.h / (1 - E_loss) - Sec.Properties.h)
    Sec.Massflow.Mass_save
    
   'outlet massflow
    Outlet_c.Massflow.m = Inlet_c.Massflow.m + Sec.Massflow.m
    Outlet_c.Massflow.Mass_save
    
    Case "Receive massflow" 'Sec massflow가 input인 경우
    
    If Sec.Massflow.m = 0 And P_loss_cold = 0 And E_loss = 0 Then
        Outlet_c.Massflow.m = Inlet_c.Massflow.m
        Set Outlet_c.Properties = Inlet_c.Properties
        delta_P_cold = 0
    Else
       'outlet massflow
        Outlet_c.Massflow.m = Inlet_c.Massflow.m + Sec.Massflow.m
        Outlet_c.Massflow.Mass_save
    
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
        Outlet_c.Properties.s = s_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
        Outlet_c.Properties.cp = Cp_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
        Outlet_c.Properties.Cv = Cv_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
        Outlet_c.Properties.k = Outlet_c.Properties.cp / Outlet_c.Properties.Cv
        Outlet_c.Properties.x = x_ph(Outlet_c.Properties.P, Outlet_c.Properties.h)
        Outlet_c.Properties.Prop_save
    End If
    End Select
    
    'delta_P_loss''''''''''''''''''''''''''''''''''''''''''''
    Mass_in_c = Inlet_c.Massflow.m
    If Inlet_c.Properties.T = Tsat_p(Inlet_c.Properties.P) Then
        sv_in_c = vV_p(Inlet_c.Properties.P)
    Else
        sv_in_c = v_pT(Inlet_c.Properties.P, Inlet_c.Properties.T)
    End If
    
     
Stop_cal_value = Outlet_c.Properties.T ''수렴성

Case 1
    
    Dim P_in_cal As Double
    Dim sv_in_c_off As Double
  
    ''hot side bypass''''''''''''''''''''''''
    Set Outlet_h.Properties = Inlet_h.Properties
    Set Outlet_h.Massflow = Inlet_h.Massflow
    Set Outlet_h.Composition = Inlet_h.Composition
    
    Select Case Calc_option
    Case "Outlet temperature", "Temperature change", "Degrees superheat"
    
   'outlet temperature
    If Calc_option = "Outlet temperature" Then 'User input : Steam outlet temperature
        Outlet_c.Properties.T = Input_value
    ElseIf Calc_option = "Temperature change" Then 'User input : Temperature change
        Outlet_c.Properties.T = Inlet_c.Properties.T - Input_value
    ElseIf Calc_option = "Degrees superheat" Then 'User input : Degrees superheat
        Outlet_c.Properties.T = Tsat_p(Inlet_c.Properties.P) + Input_value
    End If
    
    'Cold side 입구 압력 계산''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If delta_P_cold = 0 Then
        Inlet_c.Properties.P = Outlet_c.Properties.P
        Inlet_c.Properties.T = T_ph(Inlet_c.Properties.P, Inlet_c.Properties.h)
    Else
        Do
        If Inlet_c.Properties.T = Tsat_p(Inlet_c.Properties.P) Then
            sv_in_c_off = vV_p(Inlet_c.Properties.P)
        Else
            sv_in_c_off = v_pT(Inlet_c.Properties.P, Inlet_c.Properties.T)
        End If
        delta_P_cold_off = delta_P_cold * ((Inlet_c.Massflow.m / Mass_in_c) ^ 1.98) * (sv_in_c_off / sv_in_c)
'        delta_P_cold_off = delta_P_cold
        
        P_in_cal = Inlet_c.Properties.P
        Inlet_c.Properties.P = Outlet_c.Properties.P + delta_P_cold_off
    
        If Abs(P_in_cal - Inlet_c.Properties.P) < 1 Then Exit Do
        Inlet_c.Properties.T = T_ph(Inlet_c.Properties.P, Inlet_c.Properties.h)
        Loop
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

  ''outlet properties
    If delta_P_cold = 0 Then
        delta_P_cold_off = 0
    Else
        If Inlet_c.Properties.T = Tsat_p(Inlet_c.Properties.P) Then
            sv_in_c_off = vV_p(Inlet_c.Properties.P)
        Else
            sv_in_c_off = v_pT(Inlet_c.Properties.P, Inlet_c.Properties.T)
        End If
        delta_P_cold_off = delta_P_cold * ((Inlet_c.Massflow.m / Mass_in_c) ^ 1.98) * (sv_in_c_off / sv_in_c)
'        delta_P_cold_off = delta_P_cold
        
    End If
    
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
    
    'Sec massflow
    Sec.Massflow.m = Inlet_c.Massflow.m * (Inlet_c.Properties.h - Outlet_c.Properties.h / (1 - E_loss)) / (Outlet_c.Properties.h / (1 - E_loss) - Sec.Properties.h)
    Sec.Massflow.Mass_save
    
  ''outlet massflow
    Outlet_c.Massflow.m = Inlet_c.Massflow.m + Sec.Massflow.m
    Outlet_c.Massflow.Mass_save
    
    Stop_cal_value = Sec.Massflow.m ''수렴성
    
    Case "Receive massflow", "Maximum steam temperaure" 'Sec massflow가 input인 경우
      
    Outlet_c.Massflow.m = Inlet_c.Massflow.m + Sec.Massflow.m
    Outlet_c.Massflow.Mass_save
    
    'Cold side 입구 압력 계산''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If delta_P_cold = 0 Then
        Inlet_c.Properties.P = Outlet_c.Properties.P
        Inlet_c.Properties.T = T_ph(Inlet_c.Properties.P, Inlet_c.Properties.h)
    Else
        Do
        If Not delta_P_cold = 0 Then
            If Inlet_c.Properties.T = Tsat_p(Inlet_c.Properties.P) Then
                sv_in_c_off = vV_p(Inlet_c.Properties.P)
            Else
                sv_in_c_off = v_pT(Inlet_c.Properties.P, Inlet_c.Properties.T)
            End If
        Else
        End If
        delta_P_cold_off = delta_P_cold * ((Inlet_c.Massflow.m / Mass_in_c) ^ 1.98) * (sv_in_c_off / sv_in_c)
'        delta_P_cold_off = delta_P_cold
        
        P_in_cal = Inlet_c.Properties.P
        Inlet_c.Properties.P = Outlet_c.Properties.P + delta_P_cold_off
    
        If Abs(P_in_cal - Inlet_c.Properties.P) < 1 Then Exit Do
        Inlet_c.Properties.T = T_ph(Inlet_c.Properties.P, Inlet_c.Properties.h)
        Loop
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Sec steam
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
    
    'outlet properties
    If delta_P_cold = 0 Then
        delta_P_cold_off = 0
    Else
        If Inlet_c.Properties.T = Tsat_p(Inlet_c.Properties.P) Then
            sv_in_c_off = vV_p(Inlet_c.Properties.P)
        Else
            sv_in_c_off = v_pT(Inlet_c.Properties.P, Inlet_c.Properties.T)
        End If
        delta_P_cold_off = delta_P_cold * ((Inlet_c.Massflow.m / Mass_in_c) ^ 1.98) * (sv_in_c_off / sv_in_c)
'        delta_P_cold_off = delta_P_cold
        
    End If

    If Sec.Massflow.m = 0 And delta_P_cold_off = 0 And E_loss = 0 Then
        Set Outlet_c.Properties = Inlet_c.Properties
    Else
        Outlet_c.Properties.P = Inlet_c.Properties.P - delta_P_cold_off
        
        Outlet_c.Properties.h = (Inlet_c.Massflow.m * Inlet_c.Properties.h + Sec.Massflow.m * Sec.Properties.h) * (1 - E_loss) / Outlet_c.Massflow.m
        
        Outlet_c.Properties.T = T_ph(Outlet_c.Properties.P, Outlet_c.Properties.h)
        Outlet_c.Properties.s = s_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
        Outlet_c.Properties.cp = Cp_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
        Outlet_c.Properties.Cv = Cv_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
        Outlet_c.Properties.k = Outlet_c.Properties.cp / Outlet_c.Properties.Cv
        Outlet_c.Properties.x = x_ph(Outlet_c.Properties.P, Outlet_c.Properties.h)
        Outlet_c.Properties.Prop_save
    End If

    Stop_cal_value = Inlet_c.Properties.P ''수렴성
    
    End Select
  
End Select
End Sub

    




