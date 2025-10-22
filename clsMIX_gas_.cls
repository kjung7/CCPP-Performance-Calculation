VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMIX_gas_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public Inlet_h_1 As New clsStream_
Public Inlet_h_2 As New clsStream_
Public Outlet_h As New clsStream_

Public mode As Integer
Public name As Variant

Public Inlet_h As New clsStream_
Public Inlet_c As New clsStream_steam_
Public Inlet_c_valve As New clsStream_steam_
Public Outlet_c As New clsStream_steam_
Public Outlet_c_pipe As New clsStream_steam_
Public Effectiveness_Temp, Velocity, UA, Effectiveness As Variant

Public Stop_cal_value As Double
Public Check_PP, Check_Mass As Variant 'pinch point/ Mass 계산 완료

Sub calMIX_gas()

Select Case mode

Case 0

    Dim dummy As Variant
    
    Set Outlet_h.Composition = Inlet_h_1.Composition
    Outlet_h.Massflow.m = Inlet_h_1.Massflow.m + Inlet_h_2.Massflow.m
    Outlet_h.Massflow.Mass_save
    
    Outlet_h.Properties.P = Inlet_h_1.Properties.P
    Outlet_h.Properties.h = (Inlet_h_1.Massflow.m * Inlet_h_1.Properties.h + Inlet_h_2.Massflow.m * Inlet_h_2.Properties.h) / Outlet_h.Massflow.m
    
    dummy = idealgas("t", "p", Outlet_h.Properties.P, "h", Outlet_h.Properties.h, Outlet_h.Composition.Composition)
    Outlet_h.Properties.T = dummy(0)
    
    dummy = idealgas("sck", "p", Outlet_h.Properties.P, "t", Outlet_h.Properties.T, Outlet_h.Composition.Composition)
    Outlet_h.Properties.s = dummy(0)
    Outlet_h.Properties.C = dummy(1)
    Outlet_h.Properties.k = dummy(2)
    
    Outlet_h.Properties.MW = Inlet_h.Properties.MW
    Outlet_h.Properties.r = Inlet_h.Properties.r
    Outlet_h.Properties.Prop_save
    
    Stop_cal_value = Outlet_h.Properties.T

Case 1

    Set Outlet_h.Composition = Inlet_h_1.Composition
    Outlet_h.Massflow.m = Inlet_h_1.Massflow.m + Inlet_h_2.Massflow.m
    Outlet_h.Massflow.Mass_save
    
    Outlet_h.Properties.P = Inlet_h_1.Properties.P
    Outlet_h.Properties.h = (Inlet_h_1.Massflow.m * Inlet_h_1.Properties.h + Inlet_h_2.Massflow.m * Inlet_h_2.Properties.h) / Outlet_h.Massflow.m
   
    dummy = idealgas("t", "p", Outlet_h.Properties.P, "h", Outlet_h.Properties.h, Outlet_h.Composition.Composition)
    Outlet_h.Properties.T = dummy(0)
    
    dummy = idealgas("sck", "p", Outlet_h.Properties.P, "t", Outlet_h.Properties.T, Outlet_h.Composition.Composition)
    Outlet_h.Properties.s = dummy(0)
    Outlet_h.Properties.C = dummy(1)
    Outlet_h.Properties.k = dummy(2)
    
    Outlet_h.Properties.MW = Inlet_h.Properties.MW
    Outlet_h.Properties.r = Inlet_h.Properties.r
    Outlet_h.Properties.Prop_save
    
    Stop_cal_value = Outlet_h.Properties.T
    
End Select

    
End Sub


