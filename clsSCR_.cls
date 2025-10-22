VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSCR_"
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

Public mode As Integer
Public name As Variant

Public P_loss_hot, delta_P_hot As Variant
Public Calc_option, Input_value As Variant
Public Method_P_loss_hot As Variant
Public Stop_cal_value As Double

Public Effectiveness, Velocity, UA, Check_Mass As Variant

Sub calSCR()

Dim dummy As Variant

Select Case mode

Case 0, 1

    Outlet_h.Properties.T = Inlet_h.Properties.T - Input_value
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
    Stop_cal_value = Outlet_h.Properties.T
    
End Select


End Sub

