VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHRSG_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public mode As Double
Public Gas As New clsGas_
Public HRSG_read As New clsHRSG_read_
Public HRSG_call As New clsHRSG_call_
Public Y As Variant

Function calHRSG(mode As Integer)

    
    Select Case mode
    
    Case 0
    
    Dim HRSG_input As Worksheet
    Dim i, j As Integer
    Dim dummy(), ea() As Variant
    Dim ea_max As Variant
    
    Set HRSG_input = ThisWorkbook.Worksheets("HRSG_input")

    Gas.Gas_in
    HRSG_read.read_data
    
    ReDim dummy(HRSG_read.Number)
    ReDim ea(HRSG_read.Number)
    
    For j = 1 To 100
    For i = 1 To HRSG_read.Number
        If j = 1 Then
            dummy(i) = 0
        ElseIf Not j = 1 Then
            dummy(i) = HRSG_call.dict(HRSG_read.index(i)).Stop_cal_value
        End If
'        If i > 20 Then
'        MsgBox i
'        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        With HRSG_call
        .mode = mode
        .i = i
        .j = j
        .Component_name = HRSG_read.index_c
        .Pressure_level = HRSG_read.index_p
        .Number = HRSG_read.Number
        .index = HRSG_read.index
        
        .D_o = HRSG_read.D_o
        .T_tube = HRSG_read.T_tube
        .Material = HRSG_read.Material
        .L_tube = HRSG_read.L_tube
        .n_tube = HRSG_read.n_tube
        .x_tube = HRSG_read.x_tube
        
        .E_loss = HRSG_read.E_loss
        .P_loss_hot = HRSG_read.P_loss_hot
        .P_loss_valve = HRSG_read.P_loss_valve
        .P_loss_cold = HRSG_read.P_loss_cold
        
        .Calc_option = HRSG_read.Calc_option
        .Input_value = HRSG_read.Input_value
        
        .index_S = HRSG_read.index_S
        .Stream_S = HRSG_read.Stream_S
        .Valve = HRSG_read.Valve
        
        .Calc_option_T = HRSG_read.Calc_option_T
        .Input_value_T = HRSG_read.Input_value_T
        .index_T = HRSG_read.index_T
        .Stream_T = HRSG_read.Stream_T
        
        .DA_type = HRSG_read.DA_type
        .P_HP_DSN = HRSG_read.P_HP_DSN
        .P_IP_DSN = HRSG_read.P_IP_DSN
        .P_LP_DSN = HRSG_read.P_LP_DSN
        .EVAP_Cir_ratio = HRSG_read.EVAP_Cir_ratio
        .EVAP_Head = HRSG_read.EVAP_Head
        .DA_Head = HRSG_read.DA_Head
        
        .n_HP_end = HRSG_read.n_HP_end
        .n_IP_end = HRSG_read.n_IP_end
        .n_LP_end = HRSG_read.n_LP_end
        .n_None_end = HRSG_read.n_None_end
        .e_EVAP_BD = HRSG_read.e_EVAP_BD
        .n_HX = HRSG_read.n_HX
        
        .n_HP_start = HRSG_read.n_HP_start
        .n_IP_start = HRSG_read.n_IP_start
        .n_LP_start = HRSG_read.n_LP_start
        
        .Method_P_loss_hot = HRSG_read.Method_P_loss_hot
        .Method_P_loss_valve = HRSG_read.Method_P_loss_valve
        .Method_P_loss_cold = HRSG_read.Method_P_loss_cold
        .BFP_Cal_option = HRSG_read.BFP_Cal_option 'BFP head calculation option
        
        Set .dict = HRSG_read.dict
        
        If i = 1 Then
            Set .Hot_side_inlet = Gas.Gas
            .Gas_ratio = 0
        Else
'            If IsEmpty(HRSG_read.SPLorMIX(i)) Then
'                Set .Hot_side_inlet = HRSG_call.dict(HRSG_call.index(i - 1)).Outlet_h
'            Else
'                If HRSG_read.SPLorMIX_ratio(i) = 1 Then
'                    Set .Hot_side_inlet_1 = HRSG_call.dict(HRSG_call.index(i - 1)).Outlet_h
'                    Set .Hot_side_inlet_2 = HRSG_call.dict(HRSG_read.SPLorMIX(i)).Outlet_h
'                Else
'                    Set .Hot_side_inlet = HRSG_call.dict(HRSG_read.SPLorMIX(i)).Outlet_h
'                End If
'            End If

            If HRSG_read.index(i) = HRSG_read.index_Para_HX_1 Then
                Set .Hot_side_inlet = HRSG_call.dict(HRSG_call.index(HRSG_read.i_Gas_SPL)).Outlet_h 'Applying Gas Split Ratio for Gas Source
                .Gas_ratio = HRSG_read.Gas_ratio_1
            ElseIf HRSG_read.index(i) = HRSG_read.index_Para_HX_2 Then
                Set .Hot_side_inlet = HRSG_call.dict(HRSG_call.index(HRSG_read.i_Gas_SPL)).Outlet_h 'Applying Gas Split Ratio for Gas Source
                .Gas_ratio = HRSG_read.Gas_ratio_2
            ElseIf HRSG_read.index_c(i) = "MIX_gas" Then
                Set .Hot_side_inlet_1 = HRSG_call.dict(HRSG_call.index(i - 1)).Outlet_h
                Set .Hot_side_inlet_2 = HRSG_call.dict(HRSG_call.index(HRSG_read.i_Para_HX_1)).Outlet_h
            Else
                Set .Hot_side_inlet = HRSG_call.dict(HRSG_call.index(i - 1)).Outlet_h   'No SPlit or Mix for Gas Side, Gas source from the previous component
                .Gas_ratio = 0
            End If

        End If
        
        End With
        HRSG_call.HRSG_call
        ea(i) = Abs(dummy(i) - HRSG_call.dict(HRSG_read.index(i)).Stop_cal_value)

    Next i

    ea_max = Application.WorksheetFunction.Max(ea)
    
        If ea_max < 10 ^ (-8) Then
            Exit For
        End If
'       If j > 13 Then MsgBox j
    Next j

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
Case 1

    Dim HRSG_input_off As Worksheet

    Set HRSG_input_off = ThisWorkbook.Worksheets("HRSG_input_off")
    Gas.Gas_in_off
    HRSG_read.read_data_off
    
    ReDim dummy(HRSG_read.Number)
    ReDim ea(HRSG_read.Number)

    For j = 1 To 150
    For i = 1 To HRSG_read.Number
        If j = 1 Then
            dummy(i) = 0
        ElseIf Not j = 1 Then
            dummy(i) = HRSG_call.dict(HRSG_read.index(i)).Stop_cal_value
        End If
'        If j = 3 And i > 5 Then
'        MsgBox i
'        End If

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        With HRSG_call
        
        .mode = mode
        .i = i
        .j = j
        .Component_name = HRSG_read.index_c
        .Pressure_level = HRSG_read.index_p
        .Number = HRSG_read.Number
        .index = HRSG_read.index

        .E_loss = HRSG_read.E_loss
        .Calc_option = HRSG_read.Calc_option
        .Input_value = HRSG_read.Input_value

        .index_S = HRSG_read.index_S
        .Stream_S = HRSG_read.Stream_S

        .Calc_option_T = HRSG_read.Calc_option_T
        .Input_value_T = HRSG_read.Input_value_T
        .index_T = HRSG_read.index_T
        .Stream_T = HRSG_read.Stream_T

        .P_HP_off = HRSG_read.P_HP_off
        .P_IP_off = HRSG_read.P_IP_off
        .P_LP_off = HRSG_read.P_LP_off
        .DA_type = HRSG_read.DA_type

        .Mass_sec = HRSG_read.Mass_sec
        .n_DESH = HRSG_read.n_DESH
        
        Set .dict = HRSG_read.dict

        If i = 1 Then
            Set .Hot_side_inlet = Gas.Gas
            .Gas_ratio = 0
        Else

            If HRSG_read.index(i) = HRSG_read.index_Para_HX_1 Then
                Set .Hot_side_inlet = HRSG_call.dict(HRSG_call.index(HRSG_read.i_Gas_SPL)).Outlet_h
                .Gas_ratio = HRSG_read.Gas_ratio_1
            ElseIf HRSG_read.index(i) = HRSG_read.index_Para_HX_2 Then
                Set .Hot_side_inlet = HRSG_call.dict(HRSG_call.index(HRSG_read.i_Gas_SPL)).Outlet_h
                .Gas_ratio = HRSG_read.Gas_ratio_2
            ElseIf HRSG_read.index_c(i) = "MIX_gas" Then
                Set .Hot_side_inlet_1 = HRSG_call.dict(HRSG_call.index(i - 1)).Outlet_h
                Set .Hot_side_inlet_2 = HRSG_call.dict(HRSG_call.index(HRSG_read.i_Para_HX_1)).Outlet_h
            Else
                Set .Hot_side_inlet = HRSG_call.dict(HRSG_call.index(i - 1)).Outlet_h
                .Gas_ratio = 0
            End If
        End If
        End With
        
        HRSG_call.HRSG_call
        '''''''''''''''''''''''''''''''''''''''''''''''''''''
        ea(i) = Abs(dummy(i) - HRSG_call.dict(HRSG_read.index(i)).Stop_cal_value)

    Next i
        ea_max = Application.WorksheetFunction.Max(ea)
        If ea_max < 10 ^ (-3) Then
        Exit For
        End If
'      MsgBox j
    
    
        Worksheets("HRSG_input_off").Cells(50, "Q").value = HRSG_call.dict(HRSG_read.index(HRSG_read.n_HP_start)).Outlet_c.Properties.P / 100
        Worksheets("HRSG_input_off").Cells(51, "Q").value = HRSG_call.dict(HRSG_read.index(HRSG_read.n_HP_start)).Outlet_c.Properties.T
        Worksheets("HRSG_input_off").Cells(52, "Q").value = HRSG_call.dict(HRSG_read.index(HRSG_read.n_HP_start)).Outlet_c.Massflow.m * 3.6
                
        Worksheets("HRSG_input_off").Cells(50, "S").value = HRSG_call.dict(HRSG_read.index(HRSG_read.n_IP_start)).Outlet_c.Properties.P / 100
        Worksheets("HRSG_input_off").Cells(51, "S").value = HRSG_call.dict(HRSG_read.index(HRSG_read.n_IP_start)).Outlet_c.Properties.T
        Worksheets("HRSG_input_off").Cells(52, "S").value = HRSG_call.dict(HRSG_read.index(HRSG_read.n_IP_start)).Outlet_c.Massflow.m * 3.6

    Next j
    
    If Not HRSG_read.n_DESH = 0 And Worksheets("HRSG_input_off").Cells(56, "Q").value = "HRSG Outlet" Then
        ReDim Y(HRSG_read.n_DESH - 1)
        If HRSG_read.n_DESH = 1 Then 'Unkown 1 개
            If HRSG_read.i_IP_DESH = 0 Then 'HP_DESH의 Calc_option이 "Maximum steam temperaure"
                Y(0) = HRSG_read.Input_value(HRSG_read.i_HP_DESH) - HRSG_call.dict(HRSG_read.index(HRSG_read.n_HP_start)).Outlet_c.Properties.T
            ElseIf HRSG_read.i_HP_DESH = 0 Then 'IP_DESH의 Calc_option이 "Maximum steam temperaure"
                Y(0) = HRSG_read.Input_value(HRSG_read.i_IP_DESH) - HRSG_call.dict(HRSG_read.index(HRSG_read.n_IP_start)).Outlet_c.Properties.T
            End If
        ElseIf HRSG_read.n_DESH = 2 Then 'Unkown 2 개/ HP_DESH와 IP_DESH의 Calc_option이 "Maximum steam temperaure"
            Y(0) = HRSG_read.Input_value(HRSG_read.i_HP_DESH) - HRSG_call.dict(HRSG_read.index(HRSG_read.n_HP_start)).Outlet_c.Properties.T
            Y(1) = HRSG_read.Input_value(HRSG_read.i_IP_DESH) - HRSG_call.dict(HRSG_read.index(HRSG_read.n_IP_start)).Outlet_c.Properties.T
        End If
        calHRSG = Y
    End If
    
    
        'STG inlet Conditions Case
    If Not HRSG_read.n_DESH = 0 And Worksheets("HRSG_input_off").Cells(56, "Q").value = "STG Inlet" Then 'Unkown이 있는 경우 Y 정의
        ReDim Y(HRSG_read.n_DESH - 1)
        If HRSG_read.n_DESH = 1 Then 'Unkown 1 개
            If HRSG_read.i_IP_DESH = 0 Then 'HP_DESH의 Calc_option이 "Maximum steam temperaure"
                Y(0) = HRSG_call.dict(HRSG_read.index(HRSG_read.n_HP_start)).Outlet_c.Properties.T - Worksheets("HRSG_input_off").Cells(55, "Q").value
            ElseIf HRSG_read.i_HP_DESH = 0 Then 'IP_DESH의 Calc_option이 "Maximum steam temperaure"
                Y(0) = HRSG_call.dict(HRSG_read.index(HRSG_read.n_IP_start)).Outlet_c.Properties.T - Worksheets("HRSG_input_off").Cells(55, "S").value
            End If
        ElseIf HRSG_read.n_DESH = 2 Then 'Unkown 2 개/ HP_DESH와 IP_DESH의 Calc_option이 "Maximum steam temperaure"
            Y(0) = HRSG_call.dict(HRSG_read.index(HRSG_read.n_HP_start)).Outlet_c.Properties.T - Worksheets("HRSG_input_off").Cells(55, "Q").value
            Y(1) = HRSG_call.dict(HRSG_read.index(HRSG_read.n_IP_start)).Outlet_c.Properties.T - Worksheets("HRSG_input_off").Cells(55, "S").value
        End If
        calHRSG = Y
    End If
    
    
End Select

End Function






