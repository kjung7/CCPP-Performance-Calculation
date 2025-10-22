VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHRSG_call_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public HP_Water As New clsSteam_
Public RH_Steam As New clsSteam_
Public IP_Water As New clsSteam_
Public LP_Water As New clsSteam_
Public Sec As New clsSteam_

Public Hot_side_inlet As clsStream_
Public Hot_side_inlet_1 As clsStream_
Public Hot_side_inlet_2 As clsStream_
Public Fuel2 As New clsGas_

Public mode, Number As Integer
Public dict As New Dictionary

Public index, Pressure_level, Component_name As Variant
Public E_loss, P_loss_hot As Variant
Public P_loss_valve, P_loss_cold As Variant
Public Calc_option, Input_value As Variant
Public index_S, Stream_S, Valve As Variant
Public Calc_option_T, Input_value_T, index_T, Stream_T As Variant
Public D_o, T_tube, Material, L_tube, n_tube, x_tube As Variant

Public Gas_ratio As Variant

Public DA_type As Variant
Public P_HP_DSN, P_IP_DSN, P_LP_DSN, EVAP_Cir_ratio, EVAP_Head, DA_Head As Double

Public P_STM_Level, P_Static, P_DRUM_Oper As Double
Public n_HP_end, n_IP_end, n_LP_end, n_None_end, e_EVAP_BD, n_HX As Integer
Public P_HP_off, P_IP_off, P_LP_off As Double
Public n_HP_start, n_IP_start, n_LP_start, n_HX_start As Integer
Public Method_P_loss_hot, Method_P_loss_valve, Method_P_loss_cold As Variant

Public Number_PL As Variant 'Number of pressure level
Public Mass_sec As Variant
Public i, j, e, V, n_DESH As Integer
Public BFP_Cal_option As Variant



Sub HRSG_call()

    Dim HRSG_input As Worksheet
    Dim HRSG_input_off As Worksheet
    
    Dim index_BFP, index_MNBFW, index_AUXBFW, index_MNSTM, index_AUXSTM As Variant
    Dim index_HP_BFP, index_IP_BFP, index_LP_BFP As Variant
    
    Dim k, q As Integer
    Dim e_HP, e_IP, e_LP, e_None As Integer
    
    Set HRSG_input = ThisWorkbook.Worksheets("HRSG_input")
    Set HRSG_input_off = ThisWorkbook.Worksheets("HRSG_input_off")

    For k = i + 1 To Number 'Same Pressure Number Forwarding
    If Pressure_level(k) = Pressure_level(i) Then
        Exit For
    End If
    Next k
          
    For q = i - 1 To 1 Step -1 'Same Pressure Number Backwarding
    If Pressure_level(q) = Pressure_level(i) Then
        Exit For
    End If
    Next q
    
    For e_HP = dict.Count To 1 Step -1  'Component Number of HP Evaporator
        If index(e_HP) = "HP_EVAP" Then Exit For
    Next e_HP
    
    For e_IP = dict.Count To 1 Step -1 'Component Number of IP Evaporator
        If index(e_IP) = "IP_EVAP" Then Exit For
    Next e_IP
    
    For e_LP = dict.Count To 1 Step -1 'Component Number of LP Evaporator / DA+LPB
        If index(e_LP) = "LP_EVAP" Or index(e_LP) = "LP_EVAP_BD" Then Exit For
    Next e_LP
    
    For e_None = dict.Count To 1 Step -1 'Componen Number of Deaerator -> DA의 존재 여부를 확인에 사용
        If index(e_None) = "None_DA" Then Exit For
    Next e_None
    
    'BFP Inlet Source
    If mode = 0 Then
        index_HP_BFP = HRSG_input.Cells(22, "AA") & "_" & HRSG_input.Cells(22, "AB")
        index_IP_BFP = HRSG_input.Cells(23, "AA") & "_" & HRSG_input.Cells(23, "AB")
        index_LP_BFP = HRSG_input.Cells(24, "AA") & "_" & HRSG_input.Cells(24, "AB")
    ElseIf mode = 1 Then
        index_HP_BFP = HRSG_input_off.Cells(22, "Q") & "_" & HRSG_input_off.Cells(22, "R")
        index_IP_BFP = HRSG_input_off.Cells(23, "Q") & "_" & HRSG_input_off.Cells(23, "R")
        index_LP_BFP = HRSG_input_off.Cells(24, "Q") & "_" & HRSG_input_off.Cells(24, "R")
    End If
    
    'Pressure Level Definition
    If Not e_HP = 0 And e_IP = 0 And e_LP = 0 Then
        Number_PL = "Single-Pressure" 'HP_EVAP -> Single pressure
    ElseIf Not e_HP = 0 And e_IP = 0 And Not e_LP = 0 Then
        Number_PL = "Double-Pressure" 'HP_EVAP & LP_EVAP -> Double-Pressure
    ElseIf Not e_HP = 0 And Not e_IP = 0 And Not e_LP = 0 Then
        Number_PL = "Triple-Pressure" 'HP_EVAP & IP_EVAP & LP_EVAP -> Triple-Pressure
    End If
    
    If n_HP_start < n_IP_start Then
        n_HX_start = n_HP_start
    ElseIf n_HP_start > n_IP_start Then
        n_HX_start = n_IP_start
    End If
    
Select Case mode
Case 0
    
    Select Case Component_name(i)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case "DB"
    '' Fuel2 ''''''''''''
    With Fuel2
        .mode = 0
    End With
    Fuel2.Fuel_in
    '''''''' '''''''''''''''''''''''"
    With dict(index(i))
        .mode = 0
        
        Set .Inlet_h.Properties = Hot_side_inlet.Properties
        Set .Inlet_h.Massflow = Hot_side_inlet.Massflow
        Set .Inlet_h.Composition = Hot_side_inlet.Composition

        Set .Fuel2.Properties = Fuel2.Fuel.Properties
        Set .Fuel2.Composition = Fuel2.Fuel.Composition
        
        .name = index(i)
        
        .P_loss_hot = P_loss_hot(i)
        .Method_P_loss_hot = Method_P_loss_hot
        
        .Calc_option = Calc_option(i)
        .Input_value = Input_value(i)

        .LHV = Fuel2.LHV_fuel
        .Efficiency = 1
        
    End With
    dict(index(i)).calDB
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case "SCR"
    With dict(index(i))
        .mode = 0
        
        Set .Inlet_h.Properties = Hot_side_inlet.Properties
        Set .Inlet_h.Massflow = Hot_side_inlet.Massflow
        Set .Inlet_h.Composition = Hot_side_inlet.Composition

        .name = index(i)
        
        .P_loss_hot = P_loss_hot(i)
        .Method_P_loss_hot = Method_P_loss_hot
        
        .Calc_option = Calc_option(i)
        .Input_value = Input_value(i)
        
    End With
    dict(index(i)).calSCR
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case "SPHT", "SPHT_RH", "SPHT_1", "SPHT_2", "SPHT_3", "HRH", "CRH", "CRH_1", "CRH_2"
     With dict(index(i))
     
        If Component_name(i) = "SPHT_RH" Then
            With RH_Steam
                .Massflow.m = HRSG_input.Cells(39, "AA") * 10 / 36
                .P = HRSG_input.Cells(39, "AB") * 100
                .T = HRSG_input.Cells(39, "AC")
            End With
            RH_Steam.calSteam
            Set .Inlet_c.Properties = RH_Steam.Properties
            Set .Inlet_c.Massflow = RH_Steam.Massflow
        Else                'Inlet conditions = From outlet conditions from the next equipment
            Set .Inlet_c.Properties = dict(index(k)).Outlet_c.Properties
            Set .Inlet_c.Massflow = dict(index(k)).Outlet_c.Massflow
        End If
        
        Set .Inlet_h.Properties = Hot_side_inlet.Properties
        .Inlet_h.Massflow.m = Hot_side_inlet.Massflow.m
        Set .Inlet_h.Composition = Hot_side_inlet.Composition
        
        'Component Number of Evaporators / DA
        If Pressure_level(i) = "HP" Then
            e = e_HP
        ElseIf Pressure_level(i) = "RH" Or Pressure_level(i) = "IP" Then
            e = e_IP
        ElseIf Pressure_level(i) = "LP" Then
            e = e_LP
        End If
        
        'delta bypass temperature calculation
        If i = n_HX_start Then
            If Number_PL = "Single-Pressure" Then
                .delta_T_bypass = (.Inlet_h.Properties.T - 180) / n_HX
            Else
                .delta_T_bypass = (.Inlet_h.Properties.T - 150) / n_HX
'                .delta_T_bypass = 0
            End If
            
        Else        'If previous component is calculated with pinch point difference and its calculation is done or it is a duct burner, delta bypass temp = 0
            If dict(index(i - 1)).Calc_option = "Pinch point temperature difference" And dict(index(i - 1)).Check_Mass = "Done" Then
                .delta_T_bypass = 0
            ElseIf dict(index(i - 1)).name = "N/A_DB" Then
                .delta_T_bypass = 0
            Else
                .delta_T_bypass = dict(index(n_HX_start)).delta_T_bypass
            End If
        End If
        
        If e > 0 Then
            If dict(index(e)).Check_Mass = "Done" Then .delta_T_bypass = 0
        End If
        
        .mode = mode
        .j = j
        .name = index(i)
        
        .Gas_ratio = Gas_ratio
        
        .E_loss = E_loss(i)
        .P_loss_hot = P_loss_hot(i)
        .Method_P_loss_hot = Method_P_loss_hot
        .P_loss_valve = P_loss_valve(i)
        .Method_P_loss_valve = Method_P_loss_valve
        .P_loss_cold = P_loss_cold(i)
        .Method_P_loss_cold = Method_P_loss_cold

        .Calc_option = Calc_option(i)
        .Input_value = Input_value(i)
 
        .D_o = D_o(i)
        .T_tube = T_tube(i)
        .Material = Material(i)
        .L_tube = L_tube(i)
        .x_tube = x_tube(i)
        .n_tube = n_tube(i)
        
        'Distinction for saturation temperature
        If Not IsEmpty(.Inlet_c.Properties.T) Then
        If .Inlet_c.Properties.T = Tsat_p(.Inlet_c.Properties.P) Then
            .Check = 1
        Else
            .Check = 0
        End If
        End If
      
     End With
     dict(index(i)).calSPHT
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case "EVAP", "EVAP_BD"
    With dict(index(i))
       
         'Evaporator Mode Selection
        .index_EVAP = HRSG_input.Cells(21, "AA") & "_" & HRSG_input.Cells(21, "AB")            'Setting Evaporator inlet
        If Pressure_level(i) = HRSG_input.Cells(18, "AA") Or Component_name(i) = "EVAP_BD" Then    'Integral LP/DA or Evaporator Blowdown Component exists
            If .index_EVAP = "None_DA" Then                                                        'If Evaporator Inlet is from DA
                .P_drum = HRSG_input.Cells(10, "AA") * 100                                          'Drum Operating Pressure = DA Operating Pressure by Pressure Equalizing
                .Inlet_c.Properties.P = .P_drum
                .Inlet_c.Properties.h = hL_p(.P_drum)
                .Inlet_c.Properties.T = Tsat_p(.P_drum)                                             'Inlet P = DA Operating Press, H = Sat Liquid Enthalpy
                .Inlet_c.Properties.s = s_ph(.P_drum, .Inlet_c.Properties.h)
            Else
                .P_drum = HRSG_input.Cells(21, "AD") * 100
                Set .Inlet_c.Properties = dict(.index_EVAP).Outlet_c.Properties
                Set .Inlet_c.Massflow = dict(.index_EVAP).Outlet_c.Massflow
            End If
        ElseIf DA_type = "no DA" Or IsEmpty(DA_type) Then
            If Number_PL = "Single-Pressure" Then 'index_EVAP와 연결되는 EVAP: HP_EVAP
                If Pressure_level(i) = "HP" Then
                
                    Set .Inlet_c.Properties = dict(.index_EVAP).Outlet_c.Properties
                    Set .Inlet_c.Massflow = dict(.index_EVAP).Outlet_c.Massflow
                End If
            Else
                If Pressure_level(i) = "LP" Then 'index_EVAP와 연결되는 EVAP: LP_EVAP
                    Set .Inlet_c.Properties = dict(.index_EVAP).Outlet_c.Properties
                    Set .Inlet_c.Massflow = dict(.index_EVAP).Outlet_c.Massflow
                Else
                    .index_EVAP = "N/A"                                                                    'Not setting Evaporator inlet
                    Set .Inlet_c.Properties = dict(index(k)).Outlet_c.Properties                      'Previous Component Outlet Stream Properties
                    Set .Inlet_c.Massflow = dict(index(k)).Outlet_c.Massflow
                End If
            End If
            .P_drum = 0
        Else
            .index_EVAP = "N/A"                                                                    'Not setting Evaporator inlet
            Set .Inlet_c.Properties = dict(index(k)).Outlet_c.Properties                      'Previous Component Outlet Stream Properties
            Set .Inlet_c.Massflow = dict(index(k)).Outlet_c.Massflow                               'Previous Component Outlet Stream Mass Flow
            .P_drum = 0
        End If
        
        Set .Inlet_h.Properties = Hot_side_inlet.Properties
        .Inlet_h.Massflow.m = Hot_side_inlet.Massflow.m
        Set .Inlet_h.Composition = Hot_side_inlet.Composition
        
        If dict(index(i - 1)).Calc_option = "Pinch point temperature difference" And dict(index(i - 1)).Check_Mass = "Done" Then
            .delta_T_bypass = 0
        Else
            .delta_T_bypass = dict(index(n_HX_start)).delta_T_bypass
        End If
        
        If dict(index(i)).Check_Mass = "Done" Then .delta_T_bypass = 0
        
        .mode = mode
        .j = j
        .name = index(i)
        .Component_name = Component_name(i)
        .Pressure_level = Pressure_level(i)
        
        .Gas_ratio = Gas_ratio
        
        .E_loss = E_loss(i)
        .P_loss_hot = P_loss_hot(i)
        .Method_P_loss_hot = Method_P_loss_hot
        .P_loss_valve = P_loss_valve(i)
        .Method_P_loss_valve = Method_P_loss_valve
        
        .Calc_option = Calc_option(i)
        .Input_value = Input_value(i)
        
        .D_o = D_o(i)
        .T_tube = T_tube(i)
        .Material = Material(i)
        .L_tube = L_tube(i)
        .x_tube = x_tube(i)
        .n_tube = n_tube(i)
        
        .Cir_ratio = EVAP_Cir_ratio
        .Head = EVAP_Head
        
        'Blowdown Mass Setting
        If Component_name(i) = "EVAP_BD" Then
            If index_IP_BFP = "LP_EVAP_BD" And index_HP_BFP = "LP_EVAP_BD" Then
                If Not dict(index(n_IP_end)).Inlet_c.Massflow.m = 15 And Not dict(index(n_HP_end)).Inlet_c.Massflow.m = 15 Then
                    .Mass_BD = dict(index(n_IP_end)).Inlet_c.Massflow.m + dict(index(n_HP_end)).Inlet_c.Massflow.m
                Else
                    .Mass_BD = 0
                End If
            ElseIf Not index_IP_BFP = "LP_EVAP_BD" And index_HP_BFP = "LP_EVAP_BD" Then
                If Not dict(index(n_HP_end)).Inlet_c.Massflow.m = 15 Then
                    .Mass_BD = dict(index(n_HP_end)).Inlet_c.Massflow.m
                Else
                    .Mass_BD = 0
                End If
            End If
        Else
            .Mass_BD = 0
        End If
        
     End With
     dict(index(i)).calEVAP
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case "ECON", "ECON_1", "ECON_2", "ECON_3", "ECON_4"
    With dict(index(i))
    
        Set .Inlet_c.Properties = dict(index(k)).Outlet_c.Properties
        Set .Inlet_c.Massflow = dict(index(k)).Outlet_c.Massflow
        Set .Inlet_h.Properties = Hot_side_inlet.Properties
        .Inlet_h.Massflow.m = Hot_side_inlet.Massflow.m
        Set .Inlet_h.Composition = Hot_side_inlet.Composition

        'Component Number of Evaporators / DA
        If Pressure_level(i) = "HP" Then
            e = e_HP
        ElseIf Pressure_level(i) = "RH" Or Pressure_level(i) = "IP" Then
            e = e_IP
        ElseIf Pressure_level(i) = "LP" Then
            e = e_LP
        End If
        
        If dict(index(i - 1)).Calc_option = "Pinch point temperature difference" And dict(index(i - 1)).Check_Mass = "Done" Then
            .delta_T_bypass = 0
        Else
            .delta_T_bypass = dict(index(n_HX_start)).delta_T_bypass
        End If
        
        If e > 0 Then
            If dict(index(e)).Check_Mass = "Done" Then .delta_T_bypass = 0
        End If
        
        .mode = mode
        .j = j
        .name = index(i)
        
        .Gas_ratio = Gas_ratio
        
        .E_loss = E_loss(i)
        .P_loss_hot = P_loss_hot(i)
        .Method_P_loss_hot = Method_P_loss_hot
        .P_loss_valve = P_loss_valve(i)
        .Method_P_loss_valve = Method_P_loss_valve
        .P_loss_cold = P_loss_cold(i)
        .Method_P_loss_cold = Method_P_loss_cold
        
        .Calc_option = Calc_option(i)
        .Input_value = Input_value(i)
        
        .D_o = D_o(i)
        .T_tube = T_tube(i)
        .Material = Material(i)
        .L_tube = L_tube(i)
        .x_tube = x_tube(i)
        .n_tube = n_tube(i)

        
     End With
     dict(index(i)).calECON
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case "ECON_DH"
    
    With dict(index(i))
    
       .mode = mode
       
        .Inlet_c.Properties.P = HRSG_input.Cells(65, "AA").value * 100 'KPA
         .Inlet_c.Properties.T = HRSG_input.Cells(66, "AA").value
        .Inlet_c.Properties.h = h_pT(HRSG_input.Cells(65, "AA").value * 100, HRSG_input.Cells(66, "AA").value)
        .Inlet_c.Massflow = HRSG_input.Cells(67, "AA").value / 3.6 'KG/S

        Set .Inlet_h.Properties = Hot_side_inlet.Properties
        .Inlet_h.Massflow.m = Hot_side_inlet.Massflow.m
        Set .Inlet_h.Composition = Hot_side_inlet.Composition

        .name = index(i)
        
        .D_o = D_o(i)
        .T_tube = T_tube(i)
        .Material = Material(i)
        .L_tube = L_tube(i)
        .x_tube = x_tube(i)
        .n_tube = n_tube(i)
        
        .Calc_option = Calc_option(i)
        .Input_value = Input_value(i)
        
        .E_loss = E_loss(i)
        .P_loss_hot = P_loss_hot(i)
        .Method_P_loss_hot = Method_P_loss_hot
        .P_loss_cold = P_loss_cold(i)
        .Method_P_loss_cold = Method_P_loss_cold
        
        .Gas_ratio = Gas_ratio
        
     End With
     dict(index(i)).calECON
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case "DESH"
    With dict(index(i))
    
        Set .Inlet_c.Properties = dict(index(k)).Outlet_c.Properties
        Set .Inlet_c.Massflow = dict(index(k)).Outlet_c.Massflow
        Set .Inlet_h.Properties = Hot_side_inlet.Properties
        .Inlet_h.Massflow.m = Hot_side_inlet.Massflow.m
        Set .Inlet_h.Composition = Hot_side_inlet.Composition
        
        .mode = mode
        .name = index(i)
        
        .E_loss = E_loss(i)
'        .P_loss_valve = P_loss_valve(i)
'        .Method_P_loss_valve = Method_P_loss_valve
        .P_loss_cold = P_loss_cold(i)
        .Method_P_loss_cold = Method_P_loss_cold
        
        .Calc_option = Calc_option(i)
        .Input_value = Input_value(i)
        .Valve = Valve(i)
        
        'Secondary flow m, P, h''''''''''''''''''
        
        If Stream_S(i) = "SEC" Then 'Secondary Flow
            If Calc_option(i) = "Receive massflow" Then 'm
                .Sec.Massflow.m = dict(index_S(i)).Sec.Massflow.m
'           Else: clsDESH_에서 계산
            End If
            
            If Valve(i) = "Yes" Then 'Valve Controlling the pressure to its destination
                .Sec.Properties.P = dict(index(k)).Outlet_c.Properties.P
                .Sec.Properties.h = dict(index_S(i)).Sec.Properties.h 'Isenthalpyic Process
            Else
                Set .Sec.Properties = dict(index_S(i)).Sec.Properties 'Holding its inlet pressure
            End If
            
        ElseIf Stream_S(i) = "OUT" Then
            If Calc_option(i) = "Receive massflow" Then 'm
                .Sec.Massflow.m = dict(index_S(i)).Outlet_c.Massflow.m
'           Else: clsDESH_에서 계산
            End If
            
            If Valve(i) = "Yes" Then 'P
                .Sec.Properties.P = dict(index(k)).Outlet_c.Properties.P
                .Sec.Properties.h = dict(index_S(i)).Outlet_c.Properties.h
            Else
                Set .Sec.Properties = dict(index_S(i)).Outlet_c.Properties
            End If
          
        ElseIf Stream_S(i) = "TER" Then
            If Calc_option(i) = "Receive massflow" Then 'm
                .Sec.Massflow.m = dict(index_S(i)).Ter.Massflow.m      'Whole flow from the corresponding component
'           Else: clsDESH_에서 계산
            End If
            
            If Valve(i) = "Yes" Then 'Valve Controlling the pressure to its destination
                .Sec.Properties.P = dict(index(k)).Outlet_c.Properties.P
                .Sec.Properties.h = dict(index_S(i)).Ter.Properties.h 'Isenthalpyic Process
            Else
                Set .Sec.Properties = dict(index_S(i)).Ter.Properties 'Holding its inlet pressure
            End If
        End If
        
     End With
     dict(index(i)).calDESH
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case "MIX", "MIX_1", "MIX_2", "MIX_RH"
    With dict(index(i))
    
        Set .Inlet_c.Properties = dict(index(k)).Outlet_c.Properties
        Set .Inlet_c.Massflow = dict(index(k)).Outlet_c.Massflow

        Set .Inlet_h.Properties = Hot_side_inlet.Properties
        .Inlet_h.Massflow.m = Hot_side_inlet.Massflow.m
        Set .Inlet_h.Composition = Hot_side_inlet.Composition
        
        'Component Number of Evaporators / DA
        If Pressure_level(i) = "HP" Then
            .e = e_HP
        ElseIf Pressure_level(i) = "RH" Or Pressure_level(i) = "IP" Then
            .e = e_IP
        ElseIf Pressure_level(i) = "LP" Then
            .e = e_LP
        ElseIf Pressure_level(i) = "None" Then
            .e = e_None
        End If
        
        .mode = mode
        .j = j
        .i = i
        .name = index(i)

        .E_loss = E_loss(i)
        .P_loss_cold = P_loss_cold(i)
        .Method_P_loss_cold = Method_P_loss_cold
        
        .Calc_option = Calc_option(i)
        .Input_value = Input_value(i)
        .Valve = Valve(i)
 
        'Secondary flow m, P, h''''''''''''''''''
        If index_S(i) = "1" Or index_S(i) = "2" Or index_S(i) = "3" Or index_S(i) = "4" Or index_S(i) = "5" Or index_S(i) = "6" Or Component_name(i) = "MIX_RH" Then
            If index_S(i) = "1" Then
            With Sec
                .Massflow.m = HRSG_input.Cells(41, "AA") * 10 / 36
                .P = HRSG_input.Cells(41, "AB") * 100
                .T = HRSG_input.Cells(41, "AC")
            End With
            Sec.calSteam
            ElseIf index_S(i) = "2" Then
            With Sec
                .Massflow.m = HRSG_input.Cells(42, "AA") * 10 / 36
                .P = HRSG_input.Cells(42, "AB") * 100
                .T = HRSG_input.Cells(42, "AC")
            End With
            Sec.calSteam
            ElseIf index_S(i) = "3" Then
            With Sec
                .Massflow.m = HRSG_input.Cells(43, "AA") * 10 / 36
                .P = HRSG_input.Cells(43, "AB") * 100
                .T = HRSG_input.Cells(43, "AC")
            End With
            Sec.calSteam
            ElseIf index_S(i) = "4" Then
            With Sec
                .Massflow.m = HRSG_input.Cells(44, "AA") * 10 / 36
                .P = HRSG_input.Cells(44, "AB") * 100
                .T = HRSG_input.Cells(44, "AC")
            End With
            Sec.calSteam
            ElseIf index_S(i) = "5" Then
            With Sec
                .Massflow.m = HRSG_input.Cells(45, "AA") * 10 / 36
                .P = HRSG_input.Cells(45, "AB") * 100
                .T = HRSG_input.Cells(45, "AC")
            End With
            Sec.calSteam
            ElseIf index_S(i) = "6" Then
            With Sec
                .Massflow.m = HRSG_input.Cells(46, "AA") * 10 / 36
                .P = HRSG_input.Cells(46, "AB") * 100
                .T = HRSG_input.Cells(46, "AC")
            End With
            Sec.calSteam
            ElseIf Component_name(i) = "MIX_RH" Then
            With Sec
                .Massflow.m = HRSG_input.Cells(39, "AA") * 10 / 36
                .P = HRSG_input.Cells(39, "AB") * 100
                .T = HRSG_input.Cells(39, "AC")
            End With
            Sec.calSteam
            End If
        Set .Sec.Properties = Sec.Properties
        Set .Sec.Massflow = Sec.Massflow
        End If
        
        If Stream_S(i) = "SEC" Then
            If Calc_option(i) = "Receive massflow" Then 'm
                .Sec.Massflow.m = dict(index_S(i)).Sec.Massflow.m
'           Else: clsMIX_에서 계산
            End If
            
            If Valve(i) = "Yes" Then  'Valve Controlling the pressure to its destination
                .Sec.Properties.P = dict(index(k)).Outlet_c.Properties.P
                .Sec.Properties.h = dict(index_S(i)).Sec.Properties.h
            Else
                Set .Sec.Properties = dict(index_S(i)).Sec.Properties 'Holding its inlet pressure
            End If
            
        ElseIf Stream_S(i) = "OUT" Then
            If Calc_option(i) = "Receive massflow" Then 'm
                .Sec.Massflow.m = dict(index_S(i)).Outlet_c.Massflow.m      'Whole flow from the corresponding component
'           Else: clsMIX_에서 계산
            End If
            
            If Valve(i) = "Yes" Then 'P
                .Sec.Properties.P = dict(index(k)).Outlet_c.Properties.P
                .Sec.Properties.h = dict(index_S(i)).Outlet_c.Properties.h
            Else
                Set .Sec.Properties = dict(index_S(i)).Outlet_c.Properties
            End If
            
           
        ElseIf Stream_S(i) = "TER" Then
            If Calc_option(i) = "Receive massflow" Then 'm
                .Sec.Massflow.m = dict(index_S(i)).Ter.Massflow.m      'Whole flow from the corresponding component
'           Else: clsMIX_에서 계산
            End If
            
            If Valve(i) = "Yes" Then  'Valve Controlling the pressure to its destination
                .Sec.Properties.P = dict(index(k)).Outlet_c.Properties.P
                .Sec.Properties.h = dict(index_S(i)).Ter.Properties.h
            Else
                Set .Sec.Properties = dict(index_S(i)).Ter.Properties 'Holding its inlet pressure
            End If
        End If
'
     End With
     dict(index(i)).calMIX
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case "SPL", "SPL_RH", "SPL_1", "SPL_2", "SPL_3"
    With dict(index(i))
    
        If Pressure_level(i) = "RH" Then
            With RH_Steam
                .Massflow.m = HRSG_input.Cells(39, "AA") * 10 / 36
                .P = HRSG_input.Cells(39, "AB") * 100
                .T = HRSG_input.Cells(39, "AC")
            End With
            RH_Steam.calSteam
            Set .Inlet_c.Properties = RH_Steam.Properties
            Set .Inlet_c.Massflow = RH_Steam.Massflow
            Set .Sec.Properties = RH_Steam.Properties
            Set .Ter.Properties = RH_Steam.Properties
        Else
            Set .Inlet_c.Properties = dict(index(k)).Outlet_c.Properties
            Set .Inlet_c.Massflow = dict(index(k)).Outlet_c.Massflow
            Set .Sec.Properties = dict(index(k)).Outlet_c.Properties
            Set .Ter.Properties = dict(index(k)).Outlet_c.Properties
            
            'Evaporator / DA Component Number
            If Pressure_level(i) = "HP" Then
                .e = e_HP
            ElseIf Pressure_level(i) = "IP" Then
                .e = e_IP
            ElseIf Pressure_level(i) = "LP" Then
                .e = e_LP
            ElseIf Pressure_level(i) = "None" Then
                .e = e_None
            End If
        End If
    
        Set .Inlet_h.Properties = Hot_side_inlet.Properties
        .Inlet_h.Massflow.m = Hot_side_inlet.Massflow.m
        Set .Inlet_h.Composition = Hot_side_inlet.Composition
        
        .mode = mode
        .j = j
        .i = i
        .name = index(i)
        .Pressure_level = Pressure_level(i)
      
        .E_loss = E_loss(i)
        .P_loss_cold = P_loss_cold(i)
        .Method_P_loss_cold = Method_P_loss_cold
        
        .Calc_option = Calc_option(i)
        .Input_value = Input_value(i)
        
        'Secondary flow massflow
        If Calc_option(i) = "Specify Flow" Then
            .Sec.Massflow.m = Input_value(i) * 10 / 36
        ElseIf Calc_option(i) = "Specify Flow Fraction" Then
            .Sec.Massflow.m = dict(index(k)).Outlet_c.Massflow.m * Input_value(i)
        ElseIf Calc_option(i) = "Receive massflow" Then
        
            If Stream_S(i) = "SEC" Then
                .Sec.Massflow.m = dict(index_S(i)).Sec.Massflow.m 'Secondary Flow from The Corresponding Component
            ElseIf Stream_S(i) = "OUT" Then
                .Sec.Massflow.m = dict(index_S(i)).Outlet_c.Massflow.m 'Whole Flow from the Corresponding Component
            ElseIf Stream_S(i) = "IN" Then
                .Sec.Massflow.m = dict(index_S(i)).Inlet_c.Massflow.m 'Whole Flow from the Corresponding Component
            End If
        ElseIf Calc_option(i) = "DA Pegging STM" Then 'DA에서 0이라 하면 0 / DA에서 계산안하면 전부 넘기기
        
            If dict(index_S(i)).Check_MN = 1 And Stream_S(i) = "MNSTM" Then 'In case of whole steam supplying to DA and it is a main source
                .Check = 1
            ElseIf dict(index_S(i)).Check_AUX = 1 And Stream_S(i) = "AUXSTM" Then 'In case of whole steam supplying to DA and it is a Aux source
                .Check = 1
            Else 'DA에서 계산되서 넘어옴
                If Stream_S(i) = "MNSTM" Then
                    .Sec.Massflow.m = dict(index_S(i)).MNSTM.Massflow.m 'Calculated steam flow in DA Component
                ElseIf Stream_S(i) = "AUXSTM" Then
                    .Sec.Massflow.m = dict(index_S(i)).AUXSTM.Massflow.m 'Calculated steam flow in DA Component
                End If
                .Check = 0
            End If
            
        End If
        
        'Tertiary flow massflow
        If Calc_option_T(i) = "Specify Flow" Then
            .Ter.Massflow.m = Input_value_T(i) * 10 / 36
        ElseIf Calc_option_T(i) = "Specify Flow Fraction" Then
            .Ter.Massflow.m = dict(index(k)).Outlet_c.Massflow.m * Input_value_T(i)
        ElseIf Calc_option_T(i) = "Receive massflow" Then
            If Stream_T(i) = "SEC" Then
                .Ter.Massflow.m = dict(index_T(i)).Sec.Massflow.m 'Secondary Flow from The Corresponding Component
            ElseIf Stream_T(i) = "OUT" Then
                .Ter.Massflow.m = dict(index_T(i)).Outlet_c.Massflow.m 'Whole Flow from the Corresponding Component
            ElseIf Stream_T(i) = "IN" Then
                .Ter.Massflow.m = dict(index_T(i)).Inlet_c.Massflow.m 'Whole Flow from the Corresponding Component
            End If
        End If
        
     End With
     
     dict(index(i)).calSPL
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case "MIX_gas"
     With dict(index(i))
         
        Set .Inlet_h_1.Properties = Hot_side_inlet_1.Properties
        Set .Inlet_h_2.Properties = Hot_side_inlet_2.Properties
        .Inlet_h_1.Massflow.m = Hot_side_inlet_1.Massflow.m
        .Inlet_h_2.Massflow.m = Hot_side_inlet_2.Massflow.m
        Set .Inlet_h_1.Composition = Hot_side_inlet_1.Composition

        .mode = mode
        .name = index(i)
      
     End With
     
     dict(index(i)).calMIX_gas
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case "BFP"
    With dict(index(i))
    
        'PUMP 입구 물성치''''''''''''''''''''
        If Pressure_level(i) = "HP" Then
            .index_BFP = index_HP_BFP
        ElseIf Pressure_level(i) = "IP" Then
            .index_BFP = index_IP_BFP
        ElseIf Pressure_level(i) = "LP" Then
            .index_BFP = index_LP_BFP
        End If
        
        If .index_BFP = "None_DA" Then
            P_DRUM_Oper = HRSG_input.Cells(10, "AA") * 100              'Deaerator Operating Pressure
            If DA_type = "Standalone DA" Then
                P_Static = rhoL_p(P_DRUM_Oper) * 9.81 * DA_Head / 1000     'Deaerator Static Head
            ElseIf DA_type = "Integral DA" Then
                P_Static = rhoL_p(P_DRUM_Oper) * 9.81 * EVAP_Head / 1000  '*** In Case of Integral DA, Static Head = Drum Static
            End If
            .Inlet_c.Properties.P = P_DRUM_Oper + P_Static         'Inlet Pressure = Vessel Operating Press + Static Head
            .Inlet_c.Properties.h = hL_p(P_DRUM_Oper)               'Inlet Enthalpy = Saturated Liquid Enthalpy corresponding to Vessel Operating Pressure
        ElseIf .index_BFP = "LP_EVAP_BD" Then
            P_DRUM_Oper = HRSG_input.Cells(21, "AD") * 100              'EVAP_BD Operating Pressure
            P_Static = rhoL_p(P_DRUM_Oper) * 9.81 * EVAP_Head / 1000
            .Inlet_c.Properties.P = P_DRUM_Oper + P_Static         'Inlet Pressure = Vessel Operating Press + Static Head
            .Inlet_c.Properties.h = hL_p(P_DRUM_Oper)               'Inlet Enthalpy = Saturated Liquid Enthalpy corresponding to Vessel Operating Pressure
        Else
            Set .Inlet_c.Properties = dict(.index_BFP).Outlet_c.Properties
        End If
        
        'PUMP 입구 질량 유량''''''''''''''''''''
        If IsEmpty(dict(index(q)).Inlet_c.Massflow.m) Then          'If mass flow from next level component is empty, Initial value = 15 kg/s (54 t/h)
            .Inlet_c.Massflow.m = 15
        Else
            .Inlet_c.Massflow.m = dict(index(q)).Inlet_c.Massflow.m 'Inlet mass flow = Oulet mass flow from the next component
        End If
        ''''''''''''''''''''''''''''''
        Set .Inlet_h.Properties = Hot_side_inlet.Properties
        .Inlet_h.Massflow.m = Hot_side_inlet.Massflow.m
        Set .Inlet_h.Composition = Hot_side_inlet.Composition
        
        'PUMP 출구 압력''''''''''''''''''''
        'BFP Suction Flow Point Setting
        If Pressure_level(i) = "HP" Then
            P_STM_Level = P_HP_DSN
        ElseIf Pressure_level(i) = "IP" Then
            P_STM_Level = P_IP_DSN
        ElseIf Pressure_level(i) = "LP" Then
            P_STM_Level = P_LP_DSN
        End If
        
        If Method_P_loss_cold = "P_loss_cold" Then
            For V = 1 To dict.Count
            If Pressure_level(V) = "HP" And Pressure_level(i) = "HP" Then
                If V = e_HP Then
                    P_STM_Level = (P_STM_Level / ((1 - P_loss_cold(V)) * (1 - P_loss_valve(V)))) + dict("HP_EVAP").P_Static / 100
                Else
                    P_STM_Level = P_STM_Level / ((1 - P_loss_cold(V)) * (1 - P_loss_valve(V)))
                End If
            ElseIf Pressure_level(V) = "IP" And Pressure_level(i) = "IP" Then
                If V = e_IP Then
                    P_STM_Level = (P_STM_Level / ((1 - P_loss_cold(V)) * (1 - P_loss_valve(V)))) + dict("IP_EVAP").P_Static / 100
                Else
                    P_STM_Level = P_STM_Level / ((1 - P_loss_cold(V)) * (1 - P_loss_valve(V)))
                End If
            ElseIf Pressure_level(V) = "LP" And Pressure_level(i) = "LP" Then
                If V = e_LP Then
                    P_STM_Level = (P_STM_Level / ((1 - P_loss_cold(V)) * (1 - P_loss_valve(V)))) + dict("LP_EVAP").P_Static / 100
                Else
                    P_STM_Level = P_STM_Level / ((1 - P_loss_cold(V)) * (1 - P_loss_valve(V)))
                End If
            End If
            Next V
            
        ElseIf Method_P_loss_cold = "delta_P_cold" Then
            For V = 1 To dict.Count
            If Pressure_level(V) = "HP" And Pressure_level(i) = "HP" Then
                If V = e_HP Then
                    P_STM_Level = P_STM_Level + (P_loss_cold(V) + P_loss_valve(V) + dict("HP_EVAP").P_Static) / 100
                Else
                    P_STM_Level = P_STM_Level + (P_loss_cold(V) + P_loss_valve(V)) / 100
                End If
            ElseIf Pressure_level(V) = "IP" And Pressure_level(i) = "IP" Then
                If V = e_IP Then
                    P_STM_Level = P_STM_Level + (P_loss_cold(V) + P_loss_valve(V) + dict("IP_EVAP").P_Static) / 100
                Else
                    P_STM_Level = P_STM_Level + (P_loss_cold(V) + P_loss_valve(V)) / 100
                End If
            ElseIf Pressure_level(V) = "LP" And Pressure_level(i) = "LP" Then
                If V = e_LP Then
                    P_STM_Level = P_STM_Level + (P_loss_cold(V) + P_loss_valve(V) + dict("LP_EVAP").P_Static) / 100
                Else
                    P_STM_Level = P_STM_Level + (P_loss_cold(V) + P_loss_valve(V)) / 100
                End If
            End If
            Next V
        End If
        .Outlet_c.Properties.P = P_STM_Level * 100
        
        .mode = mode
        .j = j
        .name = index(i)
        
        .Calc_option = Calc_option(i)
        .Input_value = Input_value(i)
        .index_S = index_S(i)

     End With
     dict(index(i)).calPUMP
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case "COP"
    With dict(index(i))

        If IsEmpty(dict(index(q)).Inlet_c.Massflow.m) Then
            .Inlet_c.Properties.P = HRSG_input.Cells(40, "X") * 100
            .Inlet_c.Properties.h = HRSG_input.Cells(42, "X")
            .Inlet_c.Massflow.m = 15
        Else
            .Inlet_c.Properties.P = HRSG_input.Cells(40, "X") * 100
            .Inlet_c.Properties.h = HRSG_input.Cells(42, "X")
            .Inlet_c.Massflow.m = dict(index(q)).Inlet_c.Massflow.m
        End If
        
        Set .Inlet_h.Properties = Hot_side_inlet.Properties
        .Inlet_h.Massflow.m = Hot_side_inlet.Massflow.m
        Set .Inlet_h.Composition = Hot_side_inlet.Composition
        
        'PUMP 출구 압력 구하기'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If e_None = 0 Then 'DA component 없음
            If e_EVAP_BD = 0 Then 'EVAP_BD 없음
                If Number_PL = "Single-Pressure" Then 'single pressure 인 경우 마지막 EVAP->HP_EVAP
                    P_STM_Level = P_HP_DSN
                Else 'double / triple pressure 인 경우 마지막 EVAP->LP_EVAP
                    P_STM_Level = P_LP_DSN
                End If
            Else
                P_STM_Level = HRSG_input.Cells(21, "AD") 'LP_EVAP_BD Operating Pressure
            End If
        Else
            P_STM_Level = HRSG_input.Cells(10, "AA") 'Deaerator or LP/DA Integral Operating Pressure
        End If
        
        If Method_P_loss_cold = "P_loss_cold" Then
            For V = 1 To dict.Count
            If e_None = 0 Then 'DA component 없음
                If e_EVAP_BD = 0 Then 'EVAP_BD 없음
                    If Number_PL = "Single-Pressure" Then 'HP_EVAP inlet_c 압력 계산 (HP_SPHT ~ HP_EVAP)
                        If V = e_HP Then
                            P_STM_Level = (P_STM_Level / ((1 - P_loss_cold(V)) * (1 - P_loss_valve(V)))) + dict("HP_EVAP").P_Static / 100
                        ElseIf Pressure_level(V) = "HP" Then
                            P_STM_Level = P_STM_Level / ((1 - P_loss_cold(V)) * (1 - P_loss_valve(V)))
                        End If
                    Else
                        If V = e_LP Then 'LP_EVAP inlet_c 압력 계산(LP_SPHT ~ LP_EVAP)
                            P_STM_Level = (P_STM_Level / ((1 - P_loss_cold(V)) * (1 - P_loss_valve(V)))) + dict("LP_EVAP").P_Static / 100
                        ElseIf Pressure_level(V) = "LP" Then
                            P_STM_Level = P_STM_Level / ((1 - P_loss_cold(V)) * (1 - P_loss_valve(V)))
                        End If
                    End If
                Else 'LP_EVAP_BD inlet_c 압력 계산
                    If V = e_EVAP_BD Then P_STM_Level = (P_STM_Level / ((1 - P_loss_cold(e_LP)) * (1 - P_loss_valve(e_LP)))) + dict("LP_EVAP_BD").P_Static / 100
                End If
                
                If Pressure_level(V) = "None" Then 'EVAP ~ COP 출구
                    P_STM_Level = P_STM_Level / ((1 - P_loss_cold(V)) * (1 - P_loss_valve(V)))
                End If
            Else
                If Pressure_level(V) = "None" Then
                    If V = e_None Then
                        P_STM_Level = (P_STM_Level / ((1 - P_loss_cold(V)) * (1 - P_loss_valve(V)))) + dict("None_DA").P_Static / 100
                    Else
                        P_STM_Level = P_STM_Level / ((1 - P_loss_cold(V)) * (1 - P_loss_valve(V)))
                    End If
                End If
            End If
            Next V
            
        ElseIf Method_P_loss_cold = "delta_P_cold" Then
        
            For V = 1 To dict.Count
            If e_None = 0 Then 'DA component 없음
                If e_EVAP_BD = 0 Then 'EVAP_BD 없음
                    If Number_PL = "Single-Pressure" Then 'HP_EVAP inlet_c 압력 계산
                        If V = e_HP Then
                            P_STM_Level = P_STM_Level + (P_loss_cold(V) + P_loss_valve(V) + dict("HP_EVAP").P_Static) / 100
                        ElseIf Pressure_level(V) = "HP" Then
                            P_STM_Level = P_STM_Level + (P_loss_cold(V) + P_loss_valve(V)) / 100
                        End If
                    Else
                        If V = e_LP Then 'LP_EVAP inlet_c 압력 계산
                            P_STM_Level = P_STM_Level + (P_loss_cold(V) + P_loss_valve(V) + dict("LP_EVAP").P_Static) / 100
                        ElseIf Pressure_level(V) = "LP" Then
                            P_STM_Level = P_STM_Level + (P_loss_cold(V) + P_loss_valve(V)) / 100
                        End If
                    End If
                Else 'LP_EVAP_BD inlet_c 압력 계산
                    If V = e_EVAP_BD Then P_STM_Level = P_STM_Level + (P_loss_cold(e_LP) + P_loss_valve(e_LP) + dict("LP_EVAP_BD").P_Static) / 100
                End If
                
                If Pressure_level(V) = "None" Then 'EVAP ~ COP 출구
                    P_STM_Level = P_STM_Level + (P_loss_cold(V) + P_loss_valve(V)) / 100
                End If
            Else
                If Pressure_level(V) = "None" Then  'DA exists
                    If V = e_None Then
                        P_STM_Level = P_STM_Level + (P_loss_cold(V) + P_loss_valve(V) + dict("None_DA").P_Static) / 100
                    Else
                        P_STM_Level = P_STM_Level + (P_loss_cold(V) + P_loss_valve(V)) / 100
                    End If
                End If
            End If
            Next V
            
        End If
        
        .Outlet_c.Properties.P = P_STM_Level * 100
        
        .mode = mode
        .j = j
        .name = index(i)
        
        .Calc_option = Calc_option(i)
        .Input_value = Input_value(i)
        .index_S = index_S(i)

     End With
     dict(index(i)).calPUMP
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case "DA"
    With dict(index(i))
    
        Set .Inlet_c.Properties = dict(index(k)).Outlet_c.Properties
        Set .Inlet_c.Massflow = dict(index(k)).Outlet_c.Massflow
        Set .Inlet_h.Properties = Hot_side_inlet.Properties
        Set .Inlet_h.Massflow = Hot_side_inlet.Massflow
        Set .Inlet_h.Composition = Hot_side_inlet.Composition
        
        .j = j
        .mode = 0
        .name = index(i)
        
        .P_loss_valve = P_loss_valve(i)
        .Method_P_loss_valve = Method_P_loss_valve
        
        .DA_type = DA_type
        .Head = DA_Head
        
        .P_drum = HRSG_input.Cells(10, "AA") * 100
        
        'Pegging_steam'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        index_MNSTM = HRSG_input.Cells(13, "AA") & "_" & HRSG_input.Cells(13, "AB")
        If HRSG_input.Cells(13, "AC") = "SEC" Then
            Set .MNSTM = dict(index_MNSTM).Sec                                      'Secondary Stream
        ElseIf HRSG_input.Cells(13, "AC") = "OUT" Then
            Set .MNSTM.Properties = dict(index_MNSTM).Outlet_c.Properties
            Set .MNSTM.Massflow = dict(index_MNSTM).Outlet_c.Massflow               'Whole steam flow from the pre-defined stream
        End If
        .MNSTM_Con = HRSG_input.Cells(13, "AD")
        
        index_AUXSTM = HRSG_input.Cells(14, "AA") & "_" & HRSG_input.Cells(14, "AB")
        If HRSG_input.Cells(14, "AC") = "SEC" Then
            Set .AUXSTM = dict(index_AUXSTM).Sec
        ElseIf HRSG_input.Cells(14, "AC") = "OUT" Then
            Set .AUXSTM.Properties = dict(index_AUXSTM).Outlet_c.Properties
            Set .AUXSTM.Massflow = dict(index_AUXSTM).Outlet_c.Massflow              'Whole steam flow from the pre-defined stream
        End If
        .AUXSTM_Con = HRSG_input.Cells(14, "AD")
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        index_MNBFW = HRSG_input.Cells(17, "AA")  'No of Pressure Level _ Component Name

        If index_MNBFW = "HP & IP & LP" Then
            If Not dict(index(n_LP_end)).Inlet_c.Massflow.m = 15 And Not dict(index(n_IP_end)).Inlet_c.Massflow.m = 15 And Not dict(index(n_HP_end)).Inlet_c.Massflow.m = 15 Then
                .MNBFW.Massflow.m = dict(index(n_LP_end)).Inlet_c.Massflow.m + dict(index(n_IP_end)).Inlet_c.Massflow.m + dict(index(n_HP_end)).Inlet_c.Massflow.m              'Whole BFP Flow = HP Q + IP Q + LP Q
            Else
                .MNBFW.Massflow.m = 0
            End If
        ElseIf index_MNBFW = "HP & IP" Then
            If Not dict(index(n_IP_end)).Inlet_c.Massflow.m = 15 And Not dict(index(n_HP_end)).Inlet_c.Massflow.m = 15 Then
                .MNBFW.Massflow.m = dict(index(n_IP_end)).Inlet_c.Massflow.m + dict(index(n_HP_end)).Inlet_c.Massflow.m                                                         'Whole BFP Flow = HP Q + LP Q
            Else
                .MNBFW.Massflow.m = 0
            End If
        ElseIf index_MNBFW = "HP & LP" Then
            If Not dict(index(n_LP_end)).Inlet_c.Massflow.m = 15 And Not dict(index(n_HP_end)).Inlet_c.Massflow.m = 15 Then
                .MNBFW.Massflow.m = dict(index(n_LP_end)).Inlet_c.Massflow.m + dict(index(n_HP_end)).Inlet_c.Massflow.m                                                         'Whole BFP Flow = HP Q + LP Q
            Else
                .MNBFW.Massflow.m = 0
            End If
        ElseIf index_MNBFW = "HP" Then
            If Not dict(index(n_HP_end)).Inlet_c.Massflow.m = 15 Then
                .MNBFW.Massflow.m = dict(index(n_HP_end)).Inlet_c.Massflow.m                                                                                                    'Whole BFP Flow = HP Q
            Else
                .MNBFW.Massflow.m = 0
            End If
        End If
        
        If DA_type = "Integral DA" Then
            index_AUXBFW = HRSG_input.Cells(18, "AA") & "_" & HRSG_input.Cells(18, "AB")                                                                                        'Auxiliary Feedwater Flow Discharge Flow Destination
            Set .AUXBFW.Massflow = dict(index_AUXBFW).Outlet_c.Massflow                                                                                                         'Auxiliary Feedwater Flow Discharge Flow Rate
        End If
       
        
    End With
    dict(index(i)).calDA
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case "CPH", "CPH_1", "CPH_2"
    With dict(index(i))
    
'        If Component_name(i) = "CPH" And number = i Or Component_name(i) = "CPH_1" And number = i Then 'Feed water in
        If i = n_None_end Then 'None의 마지막 component가 CPH인 경우 sheet값 읽기
             If IsEmpty(dict(index(q)).Inlet_c.Massflow.m) Then
                .Inlet_c.Properties.P = HRSG_input.Cells(40, "X") * 100
                .Inlet_c.Properties.T = HRSG_input.Cells(41, "X")
                .Inlet_c.Properties.h = HRSG_input.Cells(42, "X")
                .Inlet_c.Massflow.m = 15
            Else
                .Inlet_c.Properties.P = HRSG_input.Cells(40, "X") * 100
                .Inlet_c.Properties.T = HRSG_input.Cells(41, "X")
                .Inlet_c.Properties.h = HRSG_input.Cells(42, "X")
                .Inlet_c.Massflow.m = dict(index(q)).Inlet_c.Massflow.m
            End If
        Else
            Set .Inlet_c.Properties = dict(index(k)).Outlet_c.Properties            'Properties from the next equipment
            Set .Inlet_c.Massflow = dict(index(k)).Outlet_c.Massflow
        End If
        
        Set .Inlet_h.Properties = Hot_side_inlet.Properties
        .Inlet_h.Massflow.m = Hot_side_inlet.Massflow.m
        Set .Inlet_h.Composition = Hot_side_inlet.Composition
        
        If dict(index(i - 1)).Calc_option = "Pinch point temperature difference" And dict(index(i - 1)).Check_Mass = "Done" Then
            .delta_T_bypass = 0
        Else
            .delta_T_bypass = dict(index(n_HX_start)).delta_T_bypass
        End If
        
        
        .mode = mode
        .name = index(i)
        .j = j
        
        .Gas_ratio = Gas_ratio
        
        .E_loss = E_loss(i)
        .P_loss_hot = P_loss_hot(i)
        .Method_P_loss_hot = Method_P_loss_hot
        .P_loss_valve = P_loss_valve(i)
        .Method_P_loss_valve = Method_P_loss_valve
        .P_loss_cold = P_loss_cold(i)
        .Method_P_loss_cold = Method_P_loss_cold
  
        .Calc_option = Calc_option(i)
        .Input_value = Input_value(i)
        
        .D_o = D_o(i)
        .T_tube = T_tube(i)
        .Material = Material(i)
        .L_tube = L_tube(i)
        .x_tube = x_tube(i)
        .n_tube = n_tube(i)
        
        If Component_name(i) = "CPH" Or Component_name(i) = "CPH_1" Then
            .CPH_mode = HRSG_input.Cells(27, "AA")
            If .CPH_mode = "Circulation" Then
                .Temp_in_c = HRSG_input.Cells(28, "AA")
            Else
                .Temp_in_c = 0
            End If
        ElseIf Component_name(i) = "CPH_2" Then
            .CPH_mode = HRSG_input.Cells(27, "AC")
            If .CPH_mode = "Circulation" Then
                .Temp_in_c = HRSG_input.Cells(28, "AC")
            Else
                .Temp_in_c = 0
            End If
        End If
        
     End With
     
     dict(index(i)).calCPH
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case "GLD"
    
    With dict(index(i))

        .mode = mode
        .name = index(i)
        .j = j
        
        Set .Inlet_c.Properties = dict(index(k)).Outlet_c.Properties
        Set .Inlet_c.Massflow = dict(index(k)).Outlet_c.Massflow

        Set .Inlet_h.Properties = Hot_side_inlet.Properties
        .Inlet_h.Massflow.m = Hot_side_inlet.Massflow.m
        Set .Inlet_h.Composition = Hot_side_inlet.Composition
        
        .Inlet_GLD.Massflow.m = HRSG_input.Cells(40, "AA") * 10 / 36
        .Inlet_GLD.Properties.P = HRSG_input.Cells(40, "AB") * 100
        .Inlet_GLD.Properties.T = HRSG_input.Cells(40, "AC")
        
        .E_loss = E_loss(i)
        .P_loss_cold = P_loss_cold(i)
        .Method_P_loss_cold = Method_P_loss_cold
        
        .Calc_option_P_GLD = HRSG_input.Cells(31, "AA")
        .Input_value_P_GLD = HRSG_input.Cells(32, "AA")
        

    End With
    dict(index(i)).calGLD
    End Select
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
Case 1

    Select Case Component_name(i)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case "DB"
    '' Fuel2 ''''''''''''
    With Fuel2
        .mode = 0
    End With
    Fuel2.Fuel_in
    '''''''' '''''''''''''''''''''''"
    With dict(index(i))
        .mode = 1
        
        Set .Inlet_h.Properties = Hot_side_inlet.Properties
        Set .Inlet_h.Massflow = Hot_side_inlet.Massflow
        Set .Inlet_h.Composition = Hot_side_inlet.Composition

        Set .Fuel2.Properties = Fuel2.Fuel.Properties
        Set .Fuel2.Composition = Fuel2.Fuel.Composition
        
        .name = index(i)
    
        .Calc_option = Calc_option(i)
        .Input_value = Input_value(i)

        .LHV = Fuel2.LHV_fuel
        .Efficiency = 1
    End With
    dict(index(i)).calDB
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case "SCR"
    With dict(index(i))
        .mode = 1
        
        Set .Inlet_h.Properties = Hot_side_inlet.Properties
        Set .Inlet_h.Massflow = Hot_side_inlet.Massflow
        Set .Inlet_h.Composition = Hot_side_inlet.Composition

        .name = index(i)
        
        .P_loss_hot = P_loss_hot(i)
        .Method_P_loss_hot = Method_P_loss_hot
        
        .Calc_option = Calc_option(i)
        .Input_value = Input_value(i)
        
    End With
    dict(index(i)).calSCR
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case "SPHT", "SPHT_RH", "SPHT_1", "SPHT_2", "SPHT_3", "HRH", "CRH", "CRH_1", "CRH_2"

     With dict(index(i))
        If Component_name(i) = "SPHT_RH" Then
            With RH_Steam
                .Massflow.m = HRSG_input_off.Cells(39, "Q") * 10 / 36
                .P = HRSG_input_off.Cells(39, "R") * 100
                .T = HRSG_input_off.Cells(39, "S")
            End With
            RH_Steam.calSteam
            Set .Inlet_c.Properties = RH_Steam.Properties
            Set .Inlet_c.Massflow = RH_Steam.Massflow
        Else
            'Terminal Point Pressure Setting
            If i = n_HP_start Then
                .Outlet_c.Properties.P = P_HP_off * 100
            ElseIf i = n_IP_start Then
                .Outlet_c.Properties.P = P_IP_off * 100
            ElseIf i = n_LP_start Then
                .Outlet_c.Properties.P = P_LP_off * 100
            Else
                .Outlet_c.Properties.P = dict(index(q)).Inlet_c.Properties.P  'Oulet Press = Inlet pressure of next component
            End If
            Set .Inlet_c.Properties = dict(index(k)).Outlet_c.Properties
            Set .Inlet_c.Massflow = dict(index(k)).Outlet_c.Massflow
        End If
        
        Set .Inlet_h.Properties = Hot_side_inlet.Properties
        .Inlet_h.Massflow.m = Hot_side_inlet.Massflow.m
        Set .Inlet_h.Composition = Hot_side_inlet.Composition
 
        .mode = 1
        .j = j
        .name = index(i)
        .Component_name = Component_name(i)
        
        .E_loss = E_loss(i)
        .Gas_ratio = Gas_ratio
      
     End With
     dict(index(i)).calSPHT
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case "EVAP", "EVAP_BD"
    With dict(index(i))
    
        .mode = mode
        .name = index(i)
        .j = j
        
        .index_EVAP = HRSG_input_off.Cells(21, "Q") & "_" & HRSG_input_off.Cells(21, "R")      'Component Stream which feeds evaporator
        If Pressure_level(i) = HRSG_input_off.Cells(18, "Q") Or Component_name(i) = "EVAP_BD" Then
            If .index_EVAP = "None_DA" Then
                .P_drum = HRSG_input_off.Cells(10, "Q") * 100
                .Inlet_c.Properties.P = .P_drum
                .Inlet_c.Properties.h = hL_p(.P_drum)
                .Inlet_c.Properties.T = Tsat_p(.P_drum)
                .Inlet_c.Properties.s = s_ph(.P_drum, .Inlet_c.properteis.h)
            Else
                .P_drum = HRSG_input_off.Cells(21, "T") * 100
                Set .Inlet_c.Properties = dict(.index_EVAP).Outlet_c.Properties
                Set .Inlet_c.Massflow = dict(.index_EVAP).Outlet_c.Massflow
            End If
        ElseIf DA_type = "no DA" Or IsEmpty(DA_type) Then
            If Number_PL = "Single-Pressure" Then 'index_EVAP와 연결되는 EVAP: HP_EVAP
                If Pressure_level(i) = "HP" Then
                    Set .Inlet_c.Properties = dict(.index_EVAP).Outlet_c.Properties
                    Set .Inlet_c.Massflow = dict(.index_EVAP).Outlet_c.Massflow
                End If
            Else
                If Pressure_level(i) = "LP" Then 'index_EVAP와 연결되는 EVAP: LP_EVAP
                    Set .Inlet_c.Properties = dict(.index_EVAP).Outlet_c.Properties
                    Set .Inlet_c.Massflow = dict(.index_EVAP).Outlet_c.Massflow
                Else
                    .index_EVAP = "N/A"
                    Set .Inlet_c.Properties = dict(index(k)).Outlet_c.Properties
                    Set .Inlet_c.Massflow = dict(index(k)).Outlet_c.Massflow
                End If
            End If
            .P_drum = 0
            .Outlet_c.Properties.P = dict(index(q)).Inlet_c.Properties.P                      'Inlet pressure from the component ahead of the EVAP
        Else
            .index_EVAP = "N/A"
            Set .Inlet_c.Properties = dict(index(k)).Outlet_c.Properties
            Set .Inlet_c.Massflow = dict(index(k)).Outlet_c.Massflow
            .P_drum = 0
            .Outlet_c.Properties.P = dict(index(q)).Inlet_c.Properties.P                      'Inlet pressure from the component ahead of the EVAP
        End If
        
        
        Set .Inlet_h.Properties = Hot_side_inlet.Properties
        .Inlet_h.Massflow.m = Hot_side_inlet.Massflow.m
        Set .Inlet_h.Composition = Hot_side_inlet.Composition
        
        .E_loss = E_loss(i)
        
        If Component_name(i) = "EVAP_BD" Then
            If index_IP_BFP = "LP_EVAP_BD" And index_HP_BFP = "LP_EVAP_BD" Then
                If Not dict(index(n_IP_end)).Inlet_c.Massflow.m = 15 And Not dict(index(n_HP_end)).Inlet_c.Massflow.m = 15 Then
                    .Mass_BD = dict(index(n_IP_end)).Inlet_c.Massflow.m + dict(index(n_HP_end)).Inlet_c.Massflow.m
                Else
                    .Mass_BD = 0
                End If
            ElseIf Not index_IP_BFP = "LP_EVAP_BD" And index_HP_BFP = "LP_EVAP_BD" Then
                If Not dict(index(n_HP_end)).Inlet_c.Massflow.m = 15 Then
                    .Mass_BD = dict(index(n_HP_end)).Inlet_c.Massflow.m
                Else
                    .Mass_BD = 0
                End If
            End If
        Else
            .Mass_BD = 0
        End If
        
        .Gas_ratio = Gas_ratio
        .BFP_Cal_option = BFP_Cal_option
        
     End With
     
     dict(index(i)).calEVAP
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case "ECON", "ECON_1", "ECON_2", "ECON_3", "ECON_4"
    
    With dict(index(i))
    
       .mode = mode
       .j = j
       
        Set .Inlet_c.Properties = dict(index(k)).Outlet_c.Properties
        Set .Inlet_c.Massflow = dict(index(k)).Outlet_c.Massflow
        Set .Inlet_h.Properties = Hot_side_inlet.Properties
        .Inlet_h.Massflow.m = Hot_side_inlet.Massflow.m
        Set .Inlet_h.Composition = Hot_side_inlet.Composition
        .Outlet_c.Properties.P = dict(index(q)).Inlet_c.Properties.P
        
        .name = index(i)
       
        .E_loss = E_loss(i)
        .Gas_ratio = Gas_ratio
        .BFP_Cal_option = BFP_Cal_option
     
     End With
     dict(index(i)).calECON
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case "ECON_DH"
    
    With dict(index(i))
    
       .mode = mode
       .j = j
       
        .Inlet_c.Properties.P = HRSG_input_off.Cells(60, "Q").value * 100 'KPA
         .Inlet_c.Properties.T = HRSG_input_off.Cells(61, "Q").value
        .Inlet_c.Properties.h = h_pT(HRSG_input_off.Cells(60, "Q").value * 100, HRSG_input_off.Cells(61, "Q").value)
        .Inlet_c.Massflow = HRSG_input_off.Cells(62, "Q").value / 3.6 'KG/S
        
        Set .Inlet_h.Properties = Hot_side_inlet.Properties
        .Inlet_h.Massflow.m = Hot_side_inlet.Massflow.m
        Set .Inlet_h.Composition = Hot_side_inlet.Composition
     
        .name = index(i)
       
        .E_loss = E_loss(i)
        .Gas_ratio = Gas_ratio
        .BFP_Cal_option = BFP_Cal_option
     
     End With
     dict(index(i)).calECON
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case "DESH"
    With dict(index(i))
    
        .mode = mode
        .name = index(i)
       
        Set .Inlet_c.Properties = dict(index(k)).Outlet_c.Properties
        Set .Inlet_c.Massflow = dict(index(k)).Outlet_c.Massflow
        .Outlet_c.Properties.P = dict(index(q)).Inlet_c.Properties.P 'Outlet Press = Next Component Inlet Press
        Set .Inlet_h.Properties = Hot_side_inlet.Properties
        .Inlet_h.Massflow.m = Hot_side_inlet.Massflow.m
        Set .Inlet_h.Composition = Hot_side_inlet.Composition

        .Calc_option = Calc_option(i)
        .Input_value = Input_value(i)
        .E_loss = E_loss(i)
        .Valve = Valve(i)
        .j = j
        
        ''Secondary flow m, P, h''''''''''''''''''
        If Stream_S(i) = "SEC" Then
            If Calc_option(i) = "Receive massflow" Then 'm
                .Sec.Massflow.m = dict(index_S(i)).Sec.Massflow.m 'If Stream is Secondary, Receiving secondary flow
            ElseIf Calc_option(i) = "Maximum steam temperaure" Then 'm
                If n_DESH = 1 Then
                    .Sec.Massflow.m = Mass_sec(0)
                ElseIf n_DESH = 2 Then
                    If Pressure_level(i) = "HP" Then
                        .Sec.Massflow.m = Mass_sec(0)
                    ElseIf Pressure_level(i) = "IP" Then
                        .Sec.Massflow.m = Mass_sec(1)
                    End If
                End If
'           Else: clsDESH_에서 계산
            End If
            
            If Valve(i) = "Yes" Then 'P
'                .Sec.Properties.P = 1                                 'Inlet pressure unknown, it would be calculated in the corresponding class
                .Sec.Properties.h = dict(index_S(i)).Sec.Properties.h
            Else
                Set .Sec.Properties = dict(index_S(i)).Sec.Properties 'Meaning receiving pressure from the stream
            End If
            
        ElseIf Stream_S(i) = "OUT" Then
            If Calc_option(i) = "Receive massflow" Then 'm
                .Sec.Massflow.m = dict(index_S(i)).Outlet_c.Massflow.m 'Receiving whole mass flow from the secondary stream
            ElseIf Calc_option(i) = "Maximum steam temperaure" Then 'm
                If n_DESH = 1 Then
                    .Sec.Massflow.m = Mass_sec(0)
                ElseIf n_DESH = 2 Then
                    If Pressure_level(i) = "HP" Then
                        .Sec.Massflow.m = Mass_sec(0)
                    ElseIf Pressure_level(i) = "IP" Then
                        .Sec.Massflow.m = Mass_sec(1)
                    End If
                End If
'           Else: clsDESH_에서 계산
            End If
            
            If Valve(i) = "Yes" Then 'P
'                .Sec.Properties.P = 1                                   'Inlet pressure unknown, it would be calculated in the corresponding class
                .Sec.Properties.h = dict(index_S(i)).Outlet_c.Properties.h
            Else
                Set .Sec.Properties = dict(index_S(i)).Outlet_c.Properties 'Meaning receiving pressure from the stream
            End If
        
        ElseIf Stream_S(i) = "TER" Then
            If Calc_option(i) = "Receive massflow" Then 'm
                .Sec.Massflow.m = dict(index_S(i)).Ter.Massflow.m 'If Stream is Secondary, Receiving secondary flow
            ElseIf Calc_option(i) = "Maximum steam temperaure" Then 'm
                If n_DESH = 1 Then
                    .Sec.Massflow.m = Mass_sec(0)
                ElseIf n_DESH = 2 Then
                    If Pressure_level(i) = "HP" Then
                        .Sec.Massflow.m = Mass_sec(0)
                    ElseIf Pressure_level(i) = "IP" Then
                        .Sec.Massflow.m = Mass_sec(1)
                    End If
                End If
'           Else: clsDESH_에서 계산
            End If
            
            If Valve(i) = "Yes" Then 'P
                .Sec.Properties.h = dict(index_S(i)).Ter.Properties.h
            Else
                Set .Sec.Properties = dict(index_S(i)).Ter.Properties 'Meaning receiving pressure from the stream
            End If
        
        End If
     
     End With
     dict(index(i)).calDESH
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case "MIX", "MIX_1", "MIX_2", "MIX_RH"
    With dict(index(i))
    
        .mode = mode
        .j = j
        .i = i
        .name = index(i)
        .Component_name = Component_name(i)
        
        Set .Inlet_c.Properties = dict(index(k)).Outlet_c.Properties
        Set .Inlet_c.Massflow = dict(index(k)).Outlet_c.Massflow
        
        .Outlet_c.Properties.P = dict(index(q)).Inlet_c.Properties.P
        
        Set .Inlet_h.Properties = Hot_side_inlet.Properties
        .Inlet_h.Massflow.m = Hot_side_inlet.Massflow.m
        Set .Inlet_h.Composition = Hot_side_inlet.Composition
     
        .Calc_option = Calc_option(i)
        .Input_value = Input_value(i)
        .E_loss = E_loss(i)
        .Valve = Valve(i)
        .BFP_Cal_option = BFP_Cal_option
        
         'Secondary flow m, P, T''''''''''''''''''''''''''''''''''''''''''''''''''''
        If index_S(i) = "1" Or index_S(i) = "2" Or index_S(i) = "3" Or index_S(i) = "4" Or index_S(i) = "5" Or index_S(i) = "6" Or Component_name(i) = "MIX_RH" Then
            If index_S(i) = "1" Then
            With Sec
                .Massflow.m = HRSG_input_off.Cells(41, "q") * 10 / 36
                .P = HRSG_input_off.Cells(41, "r") * 100
                .T = HRSG_input_off.Cells(41, "S")
            End With
            Sec.calSteam
            ElseIf index_S(i) = "2" Then
            With Sec
                .Massflow.m = HRSG_input_off.Cells(42, "q") * 10 / 36
                .P = HRSG_input_off.Cells(42, "r") * 100
                .T = HRSG_input_off.Cells(42, "S")
            End With
            Sec.calSteam
            ElseIf index_S(i) = "3" Then
            With Sec
                .Massflow.m = HRSG_input_off.Cells(43, "q") * 10 / 36
                .P = HRSG_input_off.Cells(43, "r") * 100
                .T = HRSG_input_off.Cells(43, "S")
            End With
            Sec.calSteam
            ElseIf index_S(i) = "4" Then
            With Sec
                .Massflow.m = HRSG_input_off.Cells(44, "q") * 10 / 36
                .P = HRSG_input_off.Cells(44, "r") * 100
                .T = HRSG_input_off.Cells(44, "S")
            End With
            Sec.calSteam
            ElseIf index_S(i) = "5" Then
            With Sec
                .Massflow.m = HRSG_input_off.Cells(45, "q") * 10 / 36
                .P = HRSG_input_off.Cells(45, "r") * 100
                .T = HRSG_input_off.Cells(45, "S")
            End With
            Sec.calSteam
            ElseIf index_S(i) = "6" Then
            With Sec
                .Massflow.m = HRSG_input_off.Cells(46, "q") * 10 / 36
                .P = HRSG_input_off.Cells(46, "r") * 100
                .T = HRSG_input_off.Cells(46, "S")
            End With
            Sec.calSteam
            ElseIf Component_name(i) = "MIX_RH" Then
            With Sec
                .Massflow.m = HRSG_input_off.Cells(39, "q") * 10 / 36
                .P = HRSG_input_off.Cells(39, "r") * 100
                .T = HRSG_input_off.Cells(39, "S")
            End With
            Sec.calSteam
            End If
        Set .Sec.Properties = Sec.Properties
        Set .Sec.Massflow = Sec.Massflow
        End If
         'Secondary flow m, P, h''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Stream_S(i) = "SEC" Then
            If Calc_option(i) = "Receive massflow" Then 'm
                .Sec.Massflow.m = dict(index_S(i)).Sec.Massflow.m
'           Else clsDESH_에서 계산
            End If
            
            If Valve(i) = "Yes" Then 'P
'                .Sec.Properties.P = 1
                .Sec.Properties.h = dict(index_S(i)).Sec.Properties.h
            Else
                Set .Sec.Properties = dict(index_S(i)).Sec.Properties
            End If
            
        ElseIf Stream_S(i) = "OUT" Then
            If Calc_option(i) = "Receive massflow" Then 'm
                .Sec.Massflow.m = dict(index_S(i)).Outlet_c.Massflow.m
'           Else clsMIX에서 계산
            End If
            
            If Valve(i) = "Yes" Then 'P
'                .Sec.Properties.P = 1
                .Sec.Properties.h = dict(index_S(i)).Outlet_c.Properties.h
            Else
                Set .Sec.Properties = dict(index_S(i)).Outlet_c.Properties
            End If
        
        ElseIf Stream_S(i) = "TER" Then
            If Calc_option(i) = "Receive massflow" Then 'm
                .Sec.Massflow.m = dict(index_S(i)).Ter.Massflow.m
'           Else clsDESH_에서 계산
            End If
            
            If Valve(i) = "Yes" Then 'P
                .Sec.Properties.h = dict(index_S(i)).Ter.Properties.h
            Else
                Set .Sec.Properties = dict(index_S(i)).Ter.Properties
            End If
            
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     End With
     dict(index(i)).calMIX
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case "SPL", "SPL_RH", "SPL_1", "SPL_2", "SPL_3"
    With dict(index(i))
       .mode = mode
       .i = i
       
        If Pressure_level(i) = "RH" Then
            With RH_Steam
                .Massflow.m = HRSG_input_off.Cells(39, "Q") * 10 / 36
                .P = HRSG_input_off.Cells(39, "R") * 100
                .T = HRSG_input_off.Cells(39, "S")
            End With
            RH_Steam.calSteam
            Set .Inlet_c.Properties = RH_Steam.Properties
            Set .Inlet_c.Massflow = RH_Steam.Massflow
            Set .Sec.Properties = RH_Steam.Properties
            Set .Ter.Properties = RH_Steam.Properties
        Else
            Set .Inlet_c.Properties = dict(index(k)).Outlet_c.Properties
            Set .Inlet_c.Massflow = dict(index(k)).Outlet_c.Massflow
            Set .Sec.Properties = dict(index(k)).Outlet_c.Properties
            Set .Ter.Properties = dict(index(k)).Outlet_c.Properties
            
            If i = n_HP_start Then
                .Outlet_c.Properties.P = P_HP_off * 100
            ElseIf i = n_IP_start Then
                .Outlet_c.Properties.P = P_IP_off * 100
            ElseIf i = n_LP_start Then
                .Outlet_c.Properties.P = P_LP_off * 100
            ElseIf q = 0 Then
                .Outlet_c.Properties.P = dict(index(i - 1)).Inlet_c.Properties.P
            Else
                .Outlet_c.Properties.P = dict(index(q)).Inlet_c.Properties.P
            End If
            
        End If
        
        Set .Inlet_h.Properties = Hot_side_inlet.Properties
        .Inlet_h.Massflow.m = Hot_side_inlet.Massflow.m
        Set .Inlet_h.Composition = Hot_side_inlet.Composition
      
        .name = index(i)
       .Pressure_level = Pressure_level(i)
        
        .E_loss = E_loss(i)
        .Calc_option = Calc_option(i)
        .Input_value = Input_value(i)
        .BFP_Cal_option = BFP_Cal_option
        
        'Secondary flow massflow
        If Calc_option(i) = "Specify Flow" Then
            .Sec.Massflow.m = Input_value(i) * 10 / 36
        ElseIf Calc_option(i) = "Specify Flow Fraction" Then
            .Sec.Massflow.m = dict(index(k)).Outlet_c.Massflow.m * Input_value(i)
        ElseIf Calc_option(i) = "Receive massflow" Then
        
            If Stream_S(i) = "SEC" Then
                .Sec.Massflow.m = dict(index_S(i)).Sec.Massflow.m
                
            ElseIf Stream_S(i) = "MNSTM" Then
                .Sec.Massflow.m = dict(index_S(i)).MNSTM.Massflow.m
            ElseIf Stream_S(i) = "AUXSTM" Then
                .Sec.Massflow.m = dict(index_S(i)).AUXSTM.Massflow.m
            ElseIf Stream_S(i) = "IN" Then
                .Sec.Massflow.m = dict(index_S(i)).Inlet_c.Massflow.m
            End If
        ElseIf Calc_option(i) = "DA Pegging STM" Then 'DA에서 0이라 하면 0 / DA에서 계산안하면 전부 넘기기
        
            If dict(index_S(i)).Check_MN = 1 And Stream_S(i) = "MNSTM" Then     'In case of Supplying Whole Steam as main source
                .Check = 1
            ElseIf dict(index_S(i)).Check_AUX = 1 And Stream_S(i) = "AUXSTM" Then    'In case of Supplying Whole Steam as aux source
                .Check = 1
            Else 'DA에서 계산되서 넘어옴
                If Stream_S(i) = "MNSTM" Then
                    .Sec.Massflow.m = dict(index_S(i)).MNSTM.Massflow.m
                ElseIf Stream_S(i) = "AUXSTM" Then
                    .Sec.Massflow.m = dict(index_S(i)).AUXSTM.Massflow.m
                End If
                .Check = 0
            End If
            
        End If
        
        'Tertiary flow massflow
        If Calc_option_T(i) = "Specify Flow" Then
            .Ter.Massflow.m = Input_value_T(i) * 10 / 36
        ElseIf Calc_option_T(i) = "Specify Flow Fraction" Then
            .Ter.Massflow.m = dict(index(k)).Outlet_c.Massflow.m * Input_value_T(i)
        ElseIf Calc_option_T(i) = "Receive massflow" Then
        
            If Stream_T(i) = "SEC" Then
                .Ter.Massflow.m = dict(index_T(i)).Sec.Massflow.m 'Secondary Flow from The Corresponding Component
            ElseIf Stream_T(i) = "OUT" Then
                .Ter.Massflow.m = dict(index_T(i)).Outlet_c.Massflow.m 'Whole Flow from the Corresponding Component
            ElseIf Stream_T(i) = "IN" Then
                .Ter.Massflow.m = dict(index_T(i)).Inlet_c.Massflow.m 'Whole Flow from the Corresponding Component
            End If
            
        End If
        
     End With
     dict(index(i)).calSPL
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case "MIX_gas"
     With dict(index(i))
         .mode = mode
         
        Set .Inlet_h_1.Properties = Hot_side_inlet_1.Properties
        Set .Inlet_h_2.Properties = Hot_side_inlet_2.Properties
        .Inlet_h_1.Massflow.m = Hot_side_inlet_1.Massflow.m
        .Inlet_h_2.Massflow.m = Hot_side_inlet_2.Massflow.m
        Set .Inlet_h_1.Composition = Hot_side_inlet_1.Composition

        .name = index(i)
      
     End With
     
     dict(index(i)).calMIX_gas
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case "BFP"
    With dict(index(i))
        .mode = mode
    
    If DA_type = "Integral DA" And HRSG_input_off.Cells(18, "Q") = "LP" Then
    
        P_STM_Level = P_LP_off * 100
        For V = 1 To e_LP - 1
            If Pressure_level(V) = "LP" Then P_STM_Level = P_STM_Level + dict(index(V)).delta_P_valve_off + dict(index(V)).delta_P_cold_off
        Next V
        
        HRSG_input_off.Cells(10, "Q").value = P_STM_Level / 100        'Pressure Overriding to the input sheet
        
    ElseIf Not e_EVAP_BD = 0 Then
    
        P_STM_Level = P_LP_off * 100
        For V = 1 To e_LP - 1
        
            If Pressure_level(V) = "LP" Then P_STM_Level = P_STM_Level + dict(index(V)).delta_P_valve_off + dict(index(V)).delta_P_cold_off

        Next V
        HRSG_input_off.Cells(21, "T").value = P_STM_Level / 100
        
    End If
    
        If Pressure_level(i) = "HP" Then
            index_BFP = index_HP_BFP
        ElseIf Pressure_level(i) = "IP" Then
            index_BFP = index_IP_BFP
        ElseIf Pressure_level(i) = "LP" Then
            index_BFP = index_LP_BFP
        End If
        
        'PUMP 입구 물성치''''''''''''''''''''
        If index_BFP = "None_DA" Then
            P_DRUM_Oper = HRSG_input_off.Cells(10, "Q") * 100
            If DA_type = "Standalone DA" Then
                P_Static = rhoL_p(P_DRUM_Oper) * 9.81 * DA_Head / 1000
            ElseIf DA_type = "Integral DA" Then
                P_Static = rhoL_p(P_DRUM_Oper) * 9.81 * EVAP_Head / 1000  '*** In Case of Integral DA, Static Head = Drum Static
            End If
            .Inlet_c.Properties.P = P_DRUM_Oper + P_Static
            .Inlet_c.Properties.h = hL_p(P_DRUM_Oper)
        ElseIf index_BFP = "LP_EVAP_BD" Then
            P_DRUM_Oper = HRSG_input_off.Cells(21, "T") * 100              'EVAP_BD Operating Pressure
            P_Static = rhoL_p(P_DRUM_Oper) * 9.81 * EVAP_Head / 1000
            .Inlet_c.Properties.P = P_DRUM_Oper + P_Static
            .Inlet_c.Properties.h = hL_p(P_DRUM_Oper)
        Else
            Set .Inlet_c.Properties = dict(.index_BFP).Outlet_c.Properties
        End If
        
        If BFP_Cal_option = "Press Cal" Then .Outlet_c.Properties.P = dict(index(q)).Inlet_c.Properties.P
        
        'PUMP 입구 질량 유량''''''''''''''''''''
        
        If IsEmpty(dict(index(q)).Inlet_c.Massflow.m) Then
            .Inlet_c.Massflow.m = 15
        Else
            .Inlet_c.Massflow.m = dict(index(q)).Inlet_c.Massflow.m                 'Mass flow from the component ahead of BFPs
'            If index(i) = "HP_BFP" Then MsgBox .Inlet_c.Massflow.m
        End If
        ''''''''''''''''''''''''''''''
        Set .Inlet_h.Properties = Hot_side_inlet.Properties
        .Inlet_h.Massflow.m = Hot_side_inlet.Massflow.m
        Set .Inlet_h.Composition = Hot_side_inlet.Composition
        
        .name = index(i)
        .Calc_option = Calc_option(i)
'        .Input_value = Input_value(i)
        .index_S = index_S(i)
        .Component_name = Component_name(i)
        .BFP_Cal_option = BFP_Cal_option
        .j = j
        
     End With
     dict(index(i)).calPUMP
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case "COP"
    With dict(index(i))
    
        .mode = mode
        .j = j
         
        If IsEmpty(dict(index(q)).Inlet_c.Massflow.m) Then
            .Inlet_c.Properties.P = HRSG_input_off.Cells(30, "N") * 100
            .Inlet_c.Properties.h = HRSG_input_off.Cells(32, "N")
            .Inlet_c.Massflow.m = 15
        Else
            .Inlet_c.Properties.P = HRSG_input_off.Cells(30, "N") * 100
            .Inlet_c.Properties.h = HRSG_input_off.Cells(32, "N")
            .Inlet_c.Massflow.m = dict(index(q)).Inlet_c.Massflow.m
        End If
        
        If BFP_Cal_option = "Press Cal" Then .Outlet_c.Properties.P = dict(index(q)).Inlet_c.Properties.P
          
        Set .Inlet_h.Properties = Hot_side_inlet.Properties
        .Inlet_h.Massflow.m = Hot_side_inlet.Massflow.m
        Set .Inlet_h.Composition = Hot_side_inlet.Composition
       
        .name = index(i)
        .Calc_option = Calc_option(i)
'        .Input_value = Input_value(i)
        .index_S = index_S(i)
        .Component_name = Component_name(i)
        .BFP_Cal_option = BFP_Cal_option

     End With
     dict(index(i)).calPUMP
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case "DA"
    
    With dict(index(i))
    
        .mode = 1
        .name = index(i)
        .j = j
        .Head = DA_Head
        
        Set .Inlet_c.Properties = dict(index(k)).Outlet_c.Properties
        Set .Inlet_c.Massflow = dict(index(k)).Outlet_c.Massflow
        
        Set .Inlet_h.Properties = Hot_side_inlet.Properties
        Set .Inlet_h.Massflow = Hot_side_inlet.Massflow
        Set .Inlet_h.Composition = Hot_side_inlet.Composition

        .DA_type = DA_type
        .P_drum = HRSG_input_off.Cells(10, "Q") * 100
        
        'Pegging_steam'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        index_MNSTM = HRSG_input_off.Cells(13, "Q") & "_" & HRSG_input_off.Cells(13, "R")
        If HRSG_input_off.Cells(13, "S") = "SEC" Then
            Set .MNSTM = dict(index_MNSTM).Sec
        ElseIf HRSG_input_off.Cells(13, "S") = "OUT" Then
            Set .MNSTM.Properties = dict(index_MNSTM).Outlet_c.Properties
            Set .MNSTM.Massflow = dict(index_MNSTM).Outlet_c.Massflow               'Whole steam flow from the pre-defined stream
        End If
        .MNSTM_Con = HRSG_input_off.Cells(13, "T")
        
        index_AUXSTM = HRSG_input_off.Cells(14, "Q") & "_" & HRSG_input_off.Cells(14, "R")
        If HRSG_input_off.Cells(14, "S") = "SEC" Then
            Set .AUXSTM = dict(index_AUXSTM).Sec
        ElseIf HRSG_input_off.Cells(14, "S") = "OUT" Then
            Set .AUXSTM.Properties = dict(index_AUXSTM).Outlet_c.Properties
            Set .AUXSTM.Massflow = dict(index_AUXSTM).Outlet_c.Massflow               'Whole steam flow from the pre-defined stream
        End If
        .AUXSTM_Con = HRSG_input_off.Cells(14, "T")
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        index_MNBFW = HRSG_input_off.Cells(17, "Q")
        
        If index_MNBFW = "HP & IP & LP" Then
            If Not dict(index(n_LP_end)).Inlet_c.Massflow.m = 15 And Not dict(index(n_IP_end)).Inlet_c.Massflow.m = 15 And Not dict(index(n_HP_end)).Inlet_c.Massflow.m = 15 Then
                .MNBFW.Massflow.m = dict(index(n_LP_end)).Inlet_c.Massflow.m + dict(index(n_IP_end)).Inlet_c.Massflow.m + dict(index(n_HP_end)).Inlet_c.Massflow.m              'Whole BFP Flow = HP Q + IP Q + LP Q
            Else
                .MNBFW.Massflow.m = 0
            End If
        ElseIf index_MNBFW = "HP & IP" Then
            If Not dict(index(n_IP_end)).Inlet_c.Massflow.m = 15 And Not dict(index(n_HP_end)).Inlet_c.Massflow.m = 15 Then
                .MNBFW.Massflow.m = dict(index(n_IP_end)).Inlet_c.Massflow.m + dict(index(n_HP_end)).Inlet_c.Massflow.m                                                         'Whole BFP Flow = HP Q + LP Q
            Else
                .MNBFW.Massflow.m = 0
            End If
        ElseIf index_MNBFW = "HP & LP" Then
            If Not dict(index(n_LP_end)).Inlet_c.Massflow.m = 15 And Not dict(index(n_HP_end)).Inlet_c.Massflow.m = 15 Then
                .MNBFW.Massflow.m = dict(index(n_LP_end)).Inlet_c.Massflow.m + dict(index(n_HP_end)).Inlet_c.Massflow.m                                                         'Whole BFP Flow = HP Q + LP Q
            Else
                .MNBFW.Massflow.m = 0
            End If
        ElseIf index_MNBFW = "HP" Then
            If Not dict(index(n_HP_end)).Inlet_c.Massflow.m = 15 Then
                .MNBFW.Massflow.m = dict(index(n_HP_end)).Inlet_c.Massflow.m                                                                                                    'Whole BFP Flow = HP Q
            Else
                .MNBFW.Massflow.m = 0
            End If
        End If
        
        If DA_type = "Integral DA" Then
            index_AUXBFW = HRSG_input_off.Cells(18, "Q") & "_" & HRSG_input_off.Cells(18, "R")
            Set .AUXBFW.Massflow = dict(index_AUXBFW).Outlet_c.Massflow
        End If
        .BFP_Cal_option = BFP_Cal_option
        
    End With
    dict(index(i)).calDA
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case "CPH", "CPH_1", "CPH_2"

    With dict(index(i))
        .mode = mode
        .name = index(i)
        .j = j
        
'        If Component_name(i) = "CPH" And number = i Or Component_name(i) = "CPH_1" And number = i Then 'Feed water in
        If i = n_None_end Then 'None의 마지막 component가 CPH인 경우 sheet값 읽기
             If IsEmpty(dict(index(q)).Inlet_c.Massflow.m) Then
                .Inlet_c.Properties.P = HRSG_input_off.Cells(30, "N") * 100
                .Inlet_c.Properties.T = HRSG_input_off.Cells(31, "N")
                .Inlet_c.Properties.h = HRSG_input_off.Cells(32, "N")
                .Inlet_c.Massflow.m = 15
            Else
                .Inlet_c.Properties.P = HRSG_input_off.Cells(30, "N") * 100
                .Inlet_c.Properties.T = HRSG_input_off.Cells(31, "N")
                .Inlet_c.Properties.h = HRSG_input_off.Cells(32, "N")
                .Inlet_c.Massflow.m = dict(index(q)).Inlet_c.Massflow.m
            End If
        Else
            Set .Inlet_c.Properties = dict(index(k)).Outlet_c.Properties
            Set .Inlet_c.Massflow = dict(index(k)).Outlet_c.Massflow
        End If
        
     If Not q = 0 Then
        .Outlet_c.Properties.P = dict(index(q)).Inlet_c.Properties.P
     Else
        .Outlet_c.Properties.P = dict(index(i - 1)).Inlet_c.Properties.P
     End If
        .BFP_Cal_option = BFP_Cal_option
         
        Set .Inlet_h.Properties = Hot_side_inlet.Properties
        .Inlet_h.Massflow.m = Hot_side_inlet.Massflow.m
        Set .Inlet_h.Composition = Hot_side_inlet.Composition
        
        If Component_name(i) = "CPH" Or Component_name(i) = "CPH_1" Then
            .CPH_mode = HRSG_input_off.Cells(27, "Q")
            If .CPH_mode = "Circulation" Then
                .Temp_in_c = HRSG_input_off.Cells(28, "Q")
            Else
                .Temp_in_c = 0
            End If
        ElseIf Component_name(i) = "CPH_2" Then

            .CPH_mode = HRSG_input_off.Cells(27, "S")
            If .CPH_mode = "Circulation" Then
                .Temp_in_c = HRSG_input_off.Cells(28, "S")
            Else
                .Temp_in_c = 0
            End If
        End If
        
        .E_loss = E_loss(i)
        .Gas_ratio = Gas_ratio

     End With
     
     dict(index(i)).calCPH
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case "GLD"
    With dict(index(i))
    
        .j = j
        .mode = mode
        .name = index(i)
        
        Set .Inlet_c.Properties = dict(index(k)).Outlet_c.Properties
        Set .Inlet_c.Massflow = dict(index(k)).Outlet_c.Massflow
        
        Set .Inlet_h.Properties = Hot_side_inlet.Properties
        .Inlet_h.Massflow.m = Hot_side_inlet.Massflow.m
        Set .Inlet_h.Composition = Hot_side_inlet.Composition
        
        .Inlet_GLD.Massflow.m = HRSG_input_off.Cells(40, "Q") * 10 / 36
        .Inlet_GLD.Properties.P = HRSG_input_off.Cells(40, "R") * 100
        .Inlet_GLD.Properties.T = HRSG_input_off.Cells(40, "S")
        
        .Calc_option_P_GLD = HRSG_input_off.Cells(31, "Q")
        .Input_value_P_GLD = HRSG_input_off.Cells(32, "Q")
        
        
        .E_loss = E_loss(i)
        .Outlet_c.Properties.P = dict(index(q)).Inlet_c.Properties.P
        .BFP_Cal_option = BFP_Cal_option
          
    End With
    dict(index(i)).calGLD
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     
    End Select
    
End Select


End Sub






