VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMixer_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Inlet As New clsStream_GT
Public Coolant As New clsStream_GT
Public Outlet As New clsStream_GT

Private Outlet_Prop(7) As Variant

Public mode As Integer

Sub cal_Mixer_de()

Select Case mode

Case 0, 11, 12, 2, 31, 32
    Dim total_mol_main, total_mol_coolant, total_mol_mix As Double
    Dim n_mix(19), w(19) As Double
    Dim i As Integer
    Dim dummy As Variant
    Dim r As Double
    Dim Property_Mixer As Variant
    Dim cp(19) As Double 'Composition 저장용
    
    total_mol_main = Inlet.Massflow / Inlet.Properties(6)
    total_mol_coolant = Coolant.Massflow / Coolant.Properties(6)

    For i = 0 To 19
        n_mix(i) = 0
        n_mix(i) = (Inlet.Composition_GT(i) * total_mol_main) + (Coolant.Composition_GT(i) * total_mol_coolant)
    Next

    total_mol_mix = 0
    For i = 0 To 19
        total_mol_mix = total_mol_mix + n_mix(i)
    Next
    
    For i = 0 To 19
        cp(i) = 0
        cp(i) = n_mix(i) / total_mol_mix
        
    Next
    
    Outlet.Composition_GT = Array(cp(0), cp(1), cp(2), cp(3), cp(4), cp(5), cp(6), cp(7), cp(8), cp(9), cp(10), cp(11), cp(12), cp(13), cp(14), cp(15), cp(16), cp(17), cp(18), cp(19))
    
    '상수 정의
    w(0) = 18.01528   'H2O
    w(1) = 31.9988    'O2
    w(2) = 44.0095    'CO2
    w(3) = 28.0134    'N2
    w(4) = 39.948     'Ar
    w(5) = 2.01588    'H2
    w(6) = 28.0101    'CO
    w(7) = 16.04246   'CH4
    w(8) = 30.06904   'C2H6
    w(9) = 44.09562   'C3H8
    w(10) = 58.1222   'n-C4H10
    w(11) = 58.1222   'i-C4H10
    w(12) = 72.14878  'n-C5H12
    w(13) = 72.14878  'i-C5H12
    w(14) = 86.17536  'n-C6H14
    w(15) = 100.20194 'n-C7H16
    w(16) = 114.22852 'n-C8H18
    w(17) = 4.0026    'He
    w(18) = 34.08088  'H2S
    w(19) = 64.064 'SO2
    r = 8.31451
    
    Outlet_Prop(6) = 0
    'total molecular weight
    For i = 0 To 19
        Outlet_Prop(6) = Outlet_Prop(6) + w(i) * Outlet.Composition_GT(i)
    Next
    
    'Gas constant
    Outlet_Prop(7) = r / Outlet_Prop(6)
    
    'Mass flow rate
    Outlet.Massflow = Inlet.Massflow + Coolant.Massflow
    
    'Enthalpy
    Outlet_Prop(2) = ((Inlet.Massflow * Inlet.Properties(2)) + (Coolant.Massflow * Coolant.Properties(2))) / Outlet.Massflow
    
    'Pressure
    Outlet_Prop(1) = Inlet.Properties(1)
    
    'Temperature
    dummy = idealgas("t", "p", Outlet_Prop(1), "h", Outlet_Prop(2), Outlet.Composition_GT)
    Outlet_Prop(0) = dummy(0)
    
    'Property
    Property_Mixer = idealgas("hsck", "p", Outlet_Prop(1), "t", Outlet_Prop(0), Outlet.Composition_GT)
    Outlet_Prop(3) = Property_Mixer(1)
    Outlet_Prop(4) = Property_Mixer(2)
    Outlet_Prop(5) = Property_Mixer(3)
    
    Outlet.Properties = Array(Outlet_Prop(0), Outlet_Prop(1), Outlet_Prop(2), Outlet_Prop(3), Outlet_Prop(4), Outlet_Prop(5), Outlet_Prop(6), Outlet_Prop(7))
    
Case Else
    MsgBox "Mixer : Case error"
End Select
End Sub


