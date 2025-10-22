VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDA_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public Inlet_h As New clsStream_
Public Outlet_h As New clsStream_
Public Inlet_c As New clsStream_steam_
Public Inlet_c_valve As New clsStream_steam_ '= BFWIN
Public Outlet_c As New clsStream_steam_

Public BFWIN As New clsStream_steam_
Public MNSTM As New clsStream_steam_
Public AUXSTM As New clsStream_steam_
Public MNBFW As New clsStream_steam_
Public AUXBFW As New clsStream_steam_

Public mode, j As Integer
Public name As Variant
Public Check_MN, Check_AUX As Variant
Public DA_type, P_drum, Head, P_Static As Double
Public MNSTM_Con, AUXSTM_Con As Variant
Public Pegging_STM As Variant
Public P_loss_valve As Double
Public Method_P_loss_valve As Variant

Public delta_P_valve As Double
Public delta_P_valve_off As Double
Public Mass_in_c As Double

Public Stop_cal_value As Double
Public Calc_option As Variant
Public Check_Mass As Variant 'EVAP Mass 계산 완료
Public Effectiveness, Velocity, UA As Variant
Public BFP_Cal_option As Variant

Sub calDA()

Select Case mode

Case 0

Dim dummy As Variant
Dim HRSG_input As Worksheet
Dim HRSG_input_off As Worksheet

    Set HRSG_input = ThisWorkbook.Worksheets("HRSG_input")
    If IsEmpty(Inlet_c.Properties.h) Or MNBFW.Massflow.m = 0 Then

        'Hotside bypass''''''''''''''''''''''''''''''
        Set Outlet_h.Properties = Inlet_h.Properties
        Set Outlet_h.Massflow = Inlet_h.Massflow
        Set Outlet_h.Composition = Inlet_h.Composition
        P_Static = rhoL_p(P_drum) * 9.81 * Head / 1000
        Exit Sub
    End If

    ''hot side bypass''''''''''''''''''''''''
    Set Outlet_h.Properties = Inlet_h.Properties
    Set Outlet_h.Massflow = Inlet_h.Massflow
    Set Outlet_h.Composition = Inlet_h.Composition
    
    P_Static = rho_ph(Inlet_c.Properties.P, Inlet_c.Properties.h) * 9.81 * Head / 1000 'kPa
    ''inlet_c_valve Properties''''''''''''''''''''''
    If DA_type = "Standalone DA" Then 'Cal inlet_c_valve P
        If P_loss_valve = 0 Then
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
            Inlet_c_valve.Properties.h = Inlet_c.Properties.h
            Inlet_c_valve.Properties.T = T_ph(Inlet_c_valve.Properties.P, Inlet_c_valve.Properties.h)
        End If
        
        '' Standalone DA" Pegging STM 존재 여부 판단''''
        If Abs(P_drum - psat_T(Inlet_c_valve.Properties.T)) < 100 Then 'pegging STM 없는 경우
            Pegging_STM = "No"
            HRSG_input.Cells(10, "AA").value = psat_T(Inlet_c_valve.Properties.T) / 100     'Pressure Writing
            Inlet_c_valve.Properties.P = psat_T(Inlet_c_valve.Properties.T)
        Else '-> inlet_c_valve P=P_drum
            Pegging_STM = "Yes"
            Inlet_c_valve.Properties.P = P_drum
        End If
        Inlet_c_valve.Properties.h = Inlet_c.Properties.h
        Inlet_c_valve.Properties.T = T_ph(Inlet_c_valve.Properties.P, Inlet_c_valve.Properties.h)
        
        delta_P_valve = Inlet_c.Properties.P - P_Static - Inlet_c_valve.Properties.P 'Off design에서 사용
    Else  'DA_type = "Integral DA" ->inlet_c_valve P=P_drum
        Pegging_STM = "Yes"
        Inlet_c_valve.Properties.P = P_drum
        Inlet_c_valve.Properties.h = Inlet_c.Properties.h
        Inlet_c_valve.Properties.T = T_ph(Inlet_c_valve.Properties.P, Inlet_c_valve.Properties.h)
    End If

    'BFWin 물성치 계산
    Set BFWIN.Properties = Inlet_c_valve.Properties
    
    'MNBFW properties
    MNBFW.Properties.P = Inlet_c_valve.Properties.P
    MNBFW.Properties.T = Tsat_p(MNBFW.Properties.P)
    MNBFW.Properties.h = h_px(MNBFW.Properties.P, 0)
    MNBFW.Properties.s = s_ph(MNBFW.Properties.P, MNBFW.Properties.h)
    MNBFW.Properties.cp = Cp_ph(MNBFW.Properties.P, MNBFW.Properties.h)
    MNBFW.Properties.Cv = Cv_ph(MNBFW.Properties.P, MNBFW.Properties.h)
    MNBFW.Properties.k = MNBFW.Properties.cp / MNBFW.Properties.Cv
    MNBFW.Properties.x = 0
    MNBFW.Properties.Prop_save
    Set Outlet_c.Properties = MNBFW.Properties

    Select Case DA_type
    
    Case "Standalone DA"
    
    'AUXBFW out 없음
    Set AUXBFW.Properties = MNBFW.Properties
    AUXBFW.Properties.Prop_save
    AUXBFW.Massflow.m = 0

    If Pegging_STM = "Yes" Then
    'Cal Pegging steam''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If AUXSTM_Con = "N/A" Or IsEmpty(AUXSTM_Con) Then 'MNSTM만 존재
            If MNSTM_Con = "Yes" Then 'MNSTM clsDA_에서 계산
                MNSTM.Massflow.m = MNBFW.Massflow.m * (MNBFW.Properties.h - BFWIN.Properties.h) / (MNSTM.Properties.h - BFWIN.Properties.h)
            ElseIf MNSTM_Con = "No (MAX)" Then  'MNSTM 질량유량 control 하지 않음(SPL에서 전부 SEC로 빠지는 경우)
                Check_MN = 1
            End If
            BFWIN.Massflow.m = MNBFW.Massflow.m - MNSTM.Massflow.m
            
        Else 'MNSTM & AUXSTM 둘다 있음 (MNSTM 질량 유량 control 하지 않음)
     Check_MN = 1
            If AUXSTM_Con = "Yes" Then 'AUXSTM clsDA_에서 계산
                AUXSTM.Massflow.m = (MNSTM.Massflow.m * (BFWIN.Properties.h - MNSTM.Properties.h) + MNBFW.Massflow.m * (MNBFW.Properties.h - BFWIN.Properties.h)) / (AUXSTM.Properties.h - BFWIN.Properties.h)
                AUXSTM.Massflow.Mass_save
            ElseIf AUXSTM_Con = "No (MAX)" Then
                Check_AUX = 1
            End If
            BFWIN.Massflow.m = MNBFW.Massflow.m - (MNSTM.Massflow.m + AUXSTM.Massflow.m)        'DA Inlet Flow = Total Feedwater (BFP DISCH) - Main Pegging - Aux Peggging
            
        End If
    Else 'Pegging steam=0(Pegging_STM=' ')
        MNSTM.Massflow.m = 0
        AUXSTM.Massflow.m = 0
        BFWIN.Massflow.m = MNBFW.Massflow.m
        
    End If
    Inlet_c.Massflow.m = BFWIN.Massflow.m               'DA Inlet Flow = Inlet Colde Side Flow Rate
    Inlet_c.Massflow.Mass_save
    Outlet_c.Massflow.m = MNBFW.Massflow.m
    Outlet_c.Massflow.Mass_save
    
    Case "Integral DA"

    'AUXBFW out=EVAP in
    Set AUXBFW.Properties = MNBFW.Properties
    AUXBFW.Properties.Prop_save

    'Cal Pegging steam''''''''''''''''''''''''
    If AUXSTM_Con = "N/A" Or IsEmpty(AUXSTM_Con) Then 'MNSTM만 존재
        If MNSTM_Con = "Yes" Then 'MNSTM clsDA_에서 계산
            MNSTM.Massflow.m = (MNBFW.Massflow.m * (MNBFW.Properties.h - BFWIN.Properties.h) + AUXBFW.Massflow.m * (AUXBFW.Properties.h - BFWIN.Properties.h)) / (MNSTM.Properties.h - BFWIN.Properties.h)
        ElseIf MNSTM_Con = "No (MAX)" Then  'MNSTM 질량유량 control 하지 않음(SPL에서 전부 SEC로 빠지는 경우)
            Check_MN = 1
        End If
        BFWIN.Massflow.m = (MNBFW.Massflow.m + AUXBFW.Massflow.m) - MNSTM.Massflow.m

    Else 'MNSTM & AUXSTM 둘다 있음 (MNSTM 질량 유량 control 하지 않음)
        Check_MN = 1
        If AUXSTM_Con = "Yes" Then 'AUXSTM clsDA_에서 계산
            AUXSTM.Massflow.m = (MNBFW.Massflow.m * (MNBFW.Properties.h - BFWIN.Properties.h) + AUXBFW.Massflow.m * (AUXBFW.Properties.h - BFWIN.Properties.h) - MNSTM.Massflow.m * (MNSTM.Properties.h - BFWIN.Properties.h)) / (AUXSTM.Properties.h - BFWIN.Properties.h)
            AUXSTM.Massflow.Mass_save
        ElseIf AUXSTM_Con = "No" Then
            Check_AUX = 1
        End If
        BFWIN.Massflow.m = (MNBFW.Massflow.m + AUXBFW.Massflow.m) - (MNSTM.Massflow.m + AUXSTM.Massflow.m) 'DA Inlet Flow = (Integral EVA Flow + all BFP Disch Flow) - (Main Pegging + Aux Pegging)

    End If
    Inlet_c.Massflow.m = BFWIN.Massflow.m 'DA Inlet Flow = Inlet Colde Side Flow Rate
    Inlet_c.Massflow.Mass_save
    Outlet_c.Massflow.m = MNBFW.Massflow.m + AUXBFW.Massflow.m
    Outlet_c.Massflow.Mass_save

    End Select
    
    Mass_in_c = Inlet_c.Massflow.m
    Stop_cal_value = Inlet_c_valve.Properties.P
    
Case 1
    
    Dim P_in_cal As Double
    
    Set HRSG_input_off = ThisWorkbook.Worksheets("HRSG_input_off")
    
    ''hot side bypass''''''''''''''''''''''''
    Set Outlet_h.Properties = Inlet_h.Properties
    Set Outlet_h.Massflow = Inlet_h.Massflow
    Set Outlet_h.Composition = Inlet_h.Composition
    
    Inlet_c_valve.Properties.P = P_drum
    Inlet_c_valve.Properties.h = Inlet_c.Properties.h
    Inlet_c_valve.Properties.T = T_ph(Inlet_c_valve.Properties.P, Inlet_c_valve.Properties.h)

    ''Standalone DA" Pegging STM 존재 여부 판단''''''''''''''''''''''''
    If DA_type = "Standalone DA" Then
        If Abs(P_drum - psat_T(Inlet_c_valve.Properties.T)) < 200 Then 'pegging STM 없는 경우
            Pegging_STM = "No"
        Else
            Pegging_STM = "Yes"
        End If
    Else
        Pegging_STM = "Yes"
    End If
    
    'BFWin 물성치 계산
    Set BFWIN.Properties = Inlet_c_valve.Properties
    'MNBFW properties
    MNBFW.Properties.P = Inlet_c_valve.Properties.P
    MNBFW.Properties.T = Tsat_p(MNBFW.Properties.P)
    MNBFW.Properties.h = h_px(MNBFW.Properties.P, 0)
    MNBFW.Properties.s = s_ph(MNBFW.Properties.P, MNBFW.Properties.h)
    MNBFW.Properties.cp = Cp_ph(MNBFW.Properties.P, MNBFW.Properties.h)
    MNBFW.Properties.Cv = Cv_ph(MNBFW.Properties.P, MNBFW.Properties.h)
    MNBFW.Properties.k = MNBFW.Properties.cp / MNBFW.Properties.Cv
    MNBFW.Properties.x = 0
    MNBFW.Properties.Prop_save
    Set Outlet_c.Properties = MNBFW.Properties

    Select Case DA_type
    
    Case "Standalone DA"
    
    'AUXBFW out 없음
    Set AUXBFW.Properties = MNBFW.Properties
    AUXBFW.Properties.Prop_save
    AUXBFW.Massflow.m = 0

    If Pegging_STM = "Yes" Then
    'Cal Pegging steam''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If AUXSTM_Con = "N/A" Or IsEmpty(AUXSTM_Con) Then 'MNSTM만 존재
            If MNSTM_Con = "Yes" Then 'MNSTM clsDA_에서 계산
                MNSTM.Massflow.m = MNBFW.Massflow.m * (MNBFW.Properties.h - BFWIN.Properties.h) / (MNSTM.Properties.h - BFWIN.Properties.h)
            ElseIf MNSTM_Con = "No (MAX)" Then  'MNSTM 질량유량 control 하지 않음(SPL에서 전부 SEC로 빠지는 경우)
                Check_MN = 1
            End If
            BFWIN.Massflow.m = MNBFW.Massflow.m - MNSTM.Massflow.m
            
        Else 'MNSTM & AUXSTM 둘다 있음 (MNSTM 질량 유량 control 하지 않음)
     Check_MN = 1
            If AUXSTM_Con = "Yes" Then 'AUXSTM clsDA_에서 계산
                AUXSTM.Massflow.m = (MNSTM.Massflow.m * (BFWIN.Properties.h - MNSTM.Properties.h) + MNBFW.Massflow.m * (MNBFW.Properties.h - BFWIN.Properties.h)) / (AUXSTM.Properties.h - BFWIN.Properties.h)
                AUXSTM.Massflow.Mass_save
            ElseIf AUXSTM_Con = "No (MAX)" Then
                Check_AUX = 1
            End If
            BFWIN.Massflow.m = MNBFW.Massflow.m - (MNSTM.Massflow.m + AUXSTM.Massflow.m)        'DA Inlet Flow = Total Feedwater (BFP DISCH) - Main Pegging - Aux Peggging
            
        End If
    Else 'Pegging steam=0(Pegging_STM=' ')
        MNSTM.Massflow.m = 0
        AUXSTM.Massflow.m = 0
        BFWIN.Massflow.m = MNBFW.Massflow.m
        
    End If
    Inlet_c.Massflow.m = BFWIN.Massflow.m               'DA Inlet Flow = Inlet Colde Side Flow Rate
    Inlet_c.Massflow.Mass_save
    Outlet_c.Massflow.m = MNBFW.Massflow.m
    Outlet_c.Massflow.Mass_save
    
    Case "Integral DA"

    'AUXBFW out=EVAP in
    Set AUXBFW.Properties = MNBFW.Properties
    AUXBFW.Properties.Prop_save

    'Cal Pegging steam''''''''''''''''''''''''
    If AUXSTM_Con = "N/A" Or IsEmpty(AUXSTM_Con) Then 'MNSTM만 존재
        If MNSTM_Con = "Yes" Then 'MNSTM clsDA_에서 계산
            MNSTM.Massflow.m = (MNBFW.Massflow.m * (MNBFW.Properties.h - BFWIN.Properties.h) + AUXBFW.Massflow.m * (AUXBFW.Properties.h - BFWIN.Properties.h)) / (MNSTM.Properties.h - BFWIN.Properties.h)
        ElseIf MNSTM_Con = "No (MAX)" Then  'MNSTM 질량유량 control 하지 않음(SPL에서 전부 SEC로 빠지는 경우)
            Check_MN = 1
        End If
        BFWIN.Massflow.m = (MNBFW.Massflow.m + AUXBFW.Massflow.m) - MNSTM.Massflow.m

    Else 'MNSTM & AUXSTM 둘다 있음 (MNSTM 질량 유량 control 하지 않음)
        Check_MN = 1
        If AUXSTM_Con = "Yes" Then 'AUXSTM clsDA_에서 계산
            AUXSTM.Massflow.m = (MNBFW.Massflow.m * (MNBFW.Properties.h - BFWIN.Properties.h) + AUXBFW.Massflow.m * (AUXBFW.Properties.h - BFWIN.Properties.h) - MNSTM.Massflow.m * (MNSTM.Properties.h - BFWIN.Properties.h)) / (AUXSTM.Properties.h - BFWIN.Properties.h)
            AUXSTM.Massflow.Mass_save
        ElseIf AUXSTM_Con = "No" Then
            Check_AUX = 1
        End If
        BFWIN.Massflow.m = (MNBFW.Massflow.m + AUXBFW.Massflow.m) - (MNSTM.Massflow.m + AUXSTM.Massflow.m) 'DA Inlet Flow = (Integral EVA Flow + all BFP Disch Flow) - (Main Pegging + Aux Pegging)

    End If
    Inlet_c.Massflow.m = BFWIN.Massflow.m 'DA Inlet Flow = Inlet Colde Side Flow Rate
    Inlet_c.Massflow.Mass_save
    Outlet_c.Massflow.m = MNBFW.Massflow.m + AUXBFW.Massflow.m
    Outlet_c.Massflow.Mass_save

    End Select
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
 
End Select

End Sub




