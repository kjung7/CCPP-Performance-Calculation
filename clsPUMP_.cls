VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPUMP_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Inlet_c As New clsStream_steam_
Public Inlet_c_valve As New clsStream_steam_
Public Outlet_c_s As New clsStream_steam_
Public Outlet_c As New clsStream_steam_

Public Inlet_h As New clsStream_
Public Outlet_h As New clsStream_
Public Effectiveness, Velocity, UA As Variant

Public mode, j As Double
Public name As Variant

Public Calc_option, Input_value, index_S, Component_name, index_BFP As Variant

Public Efficiency As Double
Public Stop_cal_value As Double
Public Check_Mass As Variant 'EVAP Mass 계산 완료
Public PUMP_Q_Design, PUMP_H_Design, PUMP_Q_OffDesign, PUMP_H_OffDesign, PUMP_Eff_Off As Double
Dim PUMP_H(3), PUMP_Q1(1, 3), Coeff_H(2) As Double
Dim PUMP_Eff(4), PUMP_Q2(2, 4), Coeff_Eff(3) As Double
Public BFP_Cal_option As Variant

Sub calPUMP()
    
    Select Case mode
    
    Case 0
    
    If Not index_BFP = "None_SPL" Then     'In case of Suction from Splitter (Direct Feeding Case), NO pressure and Enthalpy have been defined for initial calculation
        
        'If IsEmpty(Inlet_c.Properties.P) Then Inlet_c.Properties.P = 100
      '  If IsEmpty(Inlet_c.Properties.h) Then Inlet_c.Properties.h = 300
        
        Inlet_c.Properties.T = T_ph(Inlet_c.Properties.P, Inlet_c.Properties.h)
        Inlet_c.Properties.s = s_ph(Inlet_c.Properties.P, Inlet_c.Properties.h)
    End If
    
    If IsEmpty(Inlet_c.Properties.h) Then
        'Hotside bypass''''''''''''''''''''''''''''''
        Set Outlet_h.Properties = Inlet_h.Properties
        Set Outlet_h.Massflow = Inlet_h.Massflow
        Set Outlet_h.Composition = Inlet_h.Composition
        Exit Sub
    End If
    'Hotside bypass''''''''''''''''''''''''
    Set Outlet_h.Properties = Inlet_h.Properties
    Set Outlet_h.Massflow = Inlet_h.Massflow
    Set Outlet_h.Composition = Inlet_h.Composition
    Set Outlet_c.Massflow = Inlet_c.Massflow
    
    Outlet_c_s.Properties.P = Outlet_c.Properties.P
    Outlet_c_s.Properties.s = Inlet_c.Properties.s
    Outlet_c_s.Properties.h = h_ps(Outlet_c_s.Properties.P, Outlet_c_s.Properties.s)
    If Calc_option = "Input efficiency" Then
        Efficiency = Input_value
        Outlet_c.Properties.h = Inlet_c.Properties.h + (Outlet_c_s.Properties.h - Inlet_c.Properties.h) / Efficiency
        Outlet_c.Properties.T = T_ph(Outlet_c.Properties.P, Outlet_c.Properties.h)
    ElseIf Calc_option = "Feed water outlet temperature" Then
        Outlet_c.Properties.T = Input_value
        Outlet_c.Properties.h = h_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
        Efficiency = (Outlet_c_s.Properties.h - Inlet_c.Properties.h) / (Outlet_c.Properties.h - Inlet_c.Properties.h)
    End If
    
    Outlet_c.Properties.s = s_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
    Outlet_c.Properties.cp = Cp_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
    Outlet_c.Properties.Cv = Cv_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
    Outlet_c.Properties.k = Outlet_c.Properties.cp / Outlet_c.Properties.Cv
    Outlet_c.Properties.x = x_ph(Outlet_c.Properties.P, Outlet_c.Properties.h)
    Outlet_c.Properties.Prop_save
    
    PUMP_Q_Design = Inlet_c.Massflow.m * v_ph(Inlet_c.Properties.P, Inlet_c.Properties.h) * 3600 'm3/h
    PUMP_H_Design = (Outlet_c.Properties.P - Inlet_c.Properties.P) * 1000 / (9.81 * 1 / (v_ph(Inlet_c.Properties.P, Inlet_c.Properties.h))) 'm
    
    Stop_cal_value = Outlet_c.Properties.h
    
Case 1

    If Not index_BFP = "None_SPL" Then
        Inlet_c.Properties.T = T_ph(Inlet_c.Properties.P, Inlet_c.Properties.h)
        Inlet_c.Properties.s = s_ph(Inlet_c.Properties.P, Inlet_c.Properties.h)
    End If
    
    'Hotside bypass''''''''''''''''''''''''
    Set Outlet_h.Properties = Inlet_h.Properties
    Set Outlet_h.Massflow = Inlet_h.Massflow
    Set Outlet_h.Composition = Inlet_h.Composition
    Set Outlet_c.Massflow = Inlet_c.Massflow

If BFP_Cal_option = "Curve Method" Then
    PUMP_Q_OffDesign = Inlet_c.Massflow.m * v_ph(Inlet_c.Properties.P, Inlet_c.Properties.h) * 3600 'm3/h

'   PUMP Off Design Head Calculation'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
     If Component_name = "BFP" Then
        PUMP_Q_Design = PUMP_Q_Design * (1 + Worksheets("HRSG_input_off").Range("N41").value + Worksheets("HRSG_input_off").Range("N42").value)
     ElseIf Component_name = "COP" Then
        PUMP_Q_Design = PUMP_Q_Design * (1 + Worksheets("HRSG_input_off").Range("N57").value + Worksheets("HRSG_input_off").Range("N58").value)
     End If
    
    PUMP_Q1(0, 0) = 0 * PUMP_Q_Design
    PUMP_Q1(1, 0) = (0 * PUMP_Q_Design) ^ 2
    PUMP_Q1(0, 1) = 0.3 * PUMP_Q_Design
    PUMP_Q1(1, 1) = (0.3 * PUMP_Q_Design) ^ 2
    PUMP_Q1(0, 2) = 1 * PUMP_Q_Design
    PUMP_Q1(1, 2) = (1 * PUMP_Q_Design) ^ 2

    If Component_name = "BFP" Then
        PUMP_Q1(0, 3) = Worksheets("HRSG_input_off").Range("N40").value * PUMP_Q_Design
        PUMP_Q1(1, 3) = (Worksheets("HRSG_input_off").Range("N40").value * PUMP_Q_Design) ^ 2
        PUMP_H(0) = (Worksheets("HRSG_input_off").Range("N36").value * PUMP_H_Design)
        PUMP_H(1) = (Worksheets("HRSG_input_off").Range("N37").value * PUMP_H_Design)
        PUMP_H(2) = (Worksheets("HRSG_input_off").Range("N38").value * PUMP_H_Design)
        PUMP_H(3) = (Worksheets("HRSG_input_off").Range("N39").value * PUMP_H_Design)
    ElseIf Component_name = "COP" Then
        PUMP_Q1(0, 3) = Worksheets("HRSG_input_off").Range("N56").value * PUMP_Q_Design
        PUMP_Q1(1, 3) = (Worksheets("HRSG_input_off").Range("N56").value * PUMP_Q_Design) ^ 2
        PUMP_H(0) = (Worksheets("HRSG_input_off").Range("N52").value * PUMP_H_Design)
        PUMP_H(1) = (Worksheets("HRSG_input_off").Range("N53").value * PUMP_H_Design)
        PUMP_H(2) = (Worksheets("HRSG_input_off").Range("N54").value * PUMP_H_Design)
        PUMP_H(3) = (Worksheets("HRSG_input_off").Range("N55").value * PUMP_H_Design)
    End If
    Coeff_H(0) = WorksheetFunction.index(WorksheetFunction.LinEst(PUMP_H, PUMP_Q1), 1)
    Coeff_H(1) = WorksheetFunction.index(WorksheetFunction.LinEst(PUMP_H, PUMP_Q1), 2)
    Coeff_H(2) = WorksheetFunction.index(WorksheetFunction.LinEst(PUMP_H, PUMP_Q1), 3)


    PUMP_H_OffDesign = Coeff_H(0) * PUMP_Q_OffDesign ^ 2 + Coeff_H(1) * PUMP_Q_OffDesign ^ 1 + Coeff_H(2)
    Outlet_c.Properties.P = Inlet_c.Properties.P + PUMP_H_OffDesign * 9.81 * (rho_ph(Inlet_c.Properties.P, Inlet_c.Properties.h)) / 1000

    
   

'    'PUMP Off Design Efficiency Calculation'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    PUMP_Q2(0, 0) = 0 * PUMP_Q_Design
    PUMP_Q2(1, 0) = (0 * PUMP_Q_Design) ^ 2
    PUMP_Q2(2, 0) = (0 * PUMP_Q_Design) ^ 3

    PUMP_Q2(0, 1) = 0.3 * PUMP_Q_Design
    PUMP_Q2(1, 1) = (0.3 * PUMP_Q_Design) ^ 2
    PUMP_Q2(2, 1) = (0.3 * PUMP_Q_Design) ^ 3

    PUMP_Q2(0, 2) = 0.9 * PUMP_Q_Design
    PUMP_Q2(1, 2) = (0.9 * PUMP_Q_Design) ^ 2
    PUMP_Q2(2, 2) = (0.9 * PUMP_Q_Design) ^ 3

    PUMP_Q2(0, 3) = 1 * PUMP_Q_Design
    PUMP_Q2(1, 3) = (1 * PUMP_Q_Design) ^ 2
    PUMP_Q2(2, 3) = (1 * PUMP_Q_Design) ^ 3

    If Component_name = "BFP" Then
        PUMP_Q2(0, 4) = Worksheets("HRSG_input_off").Range("N40").value * PUMP_Q_Design
        PUMP_Q2(1, 4) = (Worksheets("HRSG_input_off").Range("N40").value * PUMP_Q_Design) ^ 2
        PUMP_Q2(2, 4) = (Worksheets("HRSG_input_off").Range("N40").value * PUMP_Q_Design) ^ 3
        PUMP_Eff(0) = Worksheets("HRSG_input_off").Range("N45").value * Efficiency
        PUMP_Eff(1) = Worksheets("HRSG_input_off").Range("N46").value * Efficiency
        PUMP_Eff(2) = Worksheets("HRSG_input_off").Range("N47").value * Efficiency
        PUMP_Eff(3) = Worksheets("HRSG_input_off").Range("N48").value * Efficiency
        PUMP_Eff(4) = Worksheets("HRSG_input_off").Range("N49").value * Efficiency
    ElseIf Component_name = "COP" Then
        PUMP_Q2(0, 4) = Worksheets("HRSG_input_off").Range("N56").value * PUMP_Q_Design
        PUMP_Q2(1, 4) = (Worksheets("HRSG_input_off").Range("N56").value * PUMP_Q_Design) ^ 2
        PUMP_Q2(2, 4) = (Worksheets("HRSG_input_off").Range("N56").value * PUMP_Q_Design) ^ 3
       PUMP_Eff(0) = Worksheets("HRSG_input_off").Range("N61").value * Efficiency
        PUMP_Eff(1) = Worksheets("HRSG_input_off").Range("N62").value * Efficiency
        PUMP_Eff(2) = Worksheets("HRSG_input_off").Range("N63").value * Efficiency
        PUMP_Eff(3) = Worksheets("HRSG_input_off").Range("N64").value * Efficiency
        PUMP_Eff(4) = Worksheets("HRSG_input_off").Range("N65").value * Efficiency
    End If

    Coeff_Eff(0) = WorksheetFunction.index(WorksheetFunction.LinEst(PUMP_Eff, PUMP_Q2), 1)
    Coeff_Eff(1) = WorksheetFunction.index(WorksheetFunction.LinEst(PUMP_Eff, PUMP_Q2), 2)
    Coeff_Eff(2) = WorksheetFunction.index(WorksheetFunction.LinEst(PUMP_Eff, PUMP_Q2), 3)
    Coeff_Eff(3) = WorksheetFunction.index(WorksheetFunction.LinEst(PUMP_Eff, PUMP_Q2), 4)

    PUMP_Eff_Off = Coeff_Eff(0) * PUMP_Q_OffDesign ^ 3 + Coeff_Eff(1) * PUMP_Q_OffDesign ^ 2 + Coeff_Eff(2) * PUMP_Q_OffDesign ^ 1 + Coeff_Eff(3)

End If
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    
    Outlet_c_s.Properties.P = Outlet_c.Properties.P
    Outlet_c_s.Properties.s = Inlet_c.Properties.s
    Outlet_c_s.Properties.h = h_ps(Outlet_c_s.Properties.P, Outlet_c_s.Properties.s)

    If PUMP_Eff_Off > 0.1 Then Efficiency = PUMP_Eff_Off
'    Outlet_c.Properties.h = Inlet_c.Properties.h + (Outlet_c_s.Properties.h - Inlet_c.Properties.h) / PUMP_Eff_Off
    Outlet_c.Properties.h = Inlet_c.Properties.h + (Outlet_c_s.Properties.h - Inlet_c.Properties.h) / Efficiency

    Outlet_c.Properties.T = T_ph(Outlet_c.Properties.P, Outlet_c.Properties.h)

    Outlet_c.Properties.s = s_ph(Outlet_c.Properties.P, Outlet_c.Properties.h)
'    Outlet_c.Properties.Cp = Cp_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
'    Outlet_c.Properties.Cv = Cv_pT(Outlet_c.Properties.P, Outlet_c.Properties.T)
'    Outlet_c.Properties.k = Outlet_c.Properties.Cp / Outlet_c.Properties.Cv
'    Outlet_c.Properties.x = x_ph(Outlet_c.Properties.P, Outlet_c.Properties.h)
    Outlet_c.Properties.Prop_save

 
    Stop_cal_value = Outlet_c.Properties.h

    End Select
End Sub


