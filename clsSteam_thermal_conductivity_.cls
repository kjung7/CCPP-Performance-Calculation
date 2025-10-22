VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSteam_thermal_conductivity_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Steam_tc_curve As Variant
Public row, col As Integer

Public Press As Variant
Public Temp As Variant
Public tc As Variant

Sub Steam_tc_read()

Dim TC_array(4, 42) As Variant
Dim section1 As Variant
Dim section2 As Variant
Dim section3 As Variant
Dim section4 As Variant
Dim section_data As Variant
Dim i, z As Double

    section1 = Array(1, 304.085, 316.395, 327.337, 341.015, 351.957, 364.271, 377.951, 388.896, 402.574, 416.258, 428.572, 440.883, 453.195, 462.774, 476.454, 488.768, 503.818, 514.764, 527.078, 539.391, 551.705, 564.019, 577.701, 587.28, 602.329, 614.64, 626.954, 640.636, 651.582, 663.893, 677.573, 689.887, 702.202, 714.514, 725.459, 740.507, 751.45, 762.395, 774.707, 788.389, 802.069)
    section2 = Array(5000, 0.0000530927, 0.0000533749, 0.0000536596, 0.0000542397, 0.000054529, 0.0000557257, 0.000056637, 0.0000575653, 0.0000581876, 0.0000597896, 0.0000611017, 0.0000617634, 0.0000627746, 0.0000641547, 0.0000652038, 0.0000666348, 0.0000680946, 0.0000695903, 0.0000711176, 0.0000722819, 0.0000738682, 0.0000754894, 0.0000771447, 0.0000788406, 0.0000805679, 0.0000814403, 0.0000832276, 0.0000850526, 0.0000869208, 0.000087862, 0.0000892988, 0.0000912586, 0.0000932614, 0.0000947883, 0.000096342, 0.0000979156, 0.0000989778, 0.0001006, 0.000102247, 0.000104489, 0.000106198)
    section3 = Array(10000, 0, 0.0000733014, 0.0000701501, 0.0000675038, 0.0000663941, 0.000065662, 0.0000652916, 0.0000652819, 0.0000656264, 0.0000663358, 0.0000670566, 0.0000677827, 0.0000685167, 0.0000696398, 0.00007078, 0.0000719361, 0.0000731125, 0.0000739056, 0.0000751155, 0.0000767655, 0.0000784473, 0.0000797331, 0.000080598, 0.0000823668, 0.0000841713, 0.000086017, 0.0000874269, 0.0000888566, 0.0000903131, 0.0000922934, 0.0000938044, 0.0000953402, 0.0000974326, 0.0000984876, 0.000099556, 0.000101182, 0.000102841, 0.000103956, 0.000105658, 0.000106802, 0.000108547)
    section4 = Array(15000, 0, 0, 0, 0, 0.0000984352, 0.0000896803, 0.0000835113, 0.0000808032, 0.00007818, 0.0000768948, 0.000077302, 0.0000768675, 0.0000768545, 0.0000772659, 0.0000781011, 0.000078088, 0.0000793635, 0.0000797883, 0.0000806508, 0.0000819712, 0.0000837686, 0.0000842154, 0.0000855942, 0.0000874727, 0.0000893907, 0.000090357, 0.0000918363, 0.0000928308, 0.0000943524, 0.0000958954, 0.0000979999, 0.0000990611, 0.000101235, 0.000101775, 0.000103441, 0.00010456, 0.000106271, 0.000107422, 0.000109183, 0.000110363, 0.000112168)
    
    For i = 1 To 4
        For z = 0 To 41
            If i = 1 Then section_data = section1
            If i = 2 Then section_data = section2
            If i = 3 Then section_data = section3
            If i = 4 Then section_data = section4
                                    
            TC_array(i, z + 1) = section_data(z)
        Next z
    Next i

    Steam_tc_curve = TC_array
    

End Sub

Sub Steam_tc_cal()

    Dim i, j As Integer
    Dim inter_1, inter_2, inter As Double
    
     For i = 2 To 4 - 1 Step 1
        If Steam_tc_curve(i, 1) <= Press And Press < Steam_tc_curve(i + 1, 1) Then
            For j = 2 To 42 - 1 Step 1
                If Steam_tc_curve(1, j) <= Temp And Temp < Steam_tc_curve(1, j + 1) Then
                    inter_1 = (Steam_tc_curve(i, j + 1) - Steam_tc_curve(i, j)) / (Steam_tc_curve(1, j + 1) - Steam_tc_curve(1, j)) * (Temp - Steam_tc_curve(1, j)) + Steam_tc_curve(i, j)
                    inter_2 = (Steam_tc_curve(i + 1, j + 1) - Steam_tc_curve(i + 1, j)) / (Steam_tc_curve(1, j + 1) - Steam_tc_curve(1, j)) * (Temp - Steam_tc_curve(1, j)) + Steam_tc_curve(i + 1, j)
                    inter = (inter_2 - inter_1) / (Steam_tc_curve(i + 1, 1) - Steam_tc_curve(i, 1)) * (Press - Steam_tc_curve(i, 1)) + inter_1
                ElseIf Temp <= Steam_tc_curve(1, 2) Then
                    inter_1 = (Steam_tc_curve(i, 3) - Steam_tc_curve(i, 2)) / (Steam_tc_curve(1, 3) - Steam_tc_curve(1, 2)) * (Temp - Steam_tc_curve(1, 2)) + Steam_tc_curve(i, 2)
                    inter_2 = (Steam_tc_curve(i + 1, 3) - Steam_tc_curve(i + 1, 2)) / (Steam_tc_curve(1, 3) - Steam_tc_curve(1, 2)) * (Temp - Steam_tc_curve(1, 2)) + Steam_tc_curve(i + 1, 2)
                    inter = (inter_2 - inter_1) / (Steam_tc_curve(i + 1, 1) - Steam_tc_curve(i, 1)) * (Press - Steam_tc_curve(i, 1)) + inter_1
                ElseIf Steam_tc_curve(1, 42) <= Temp Then
                    inter_1 = (Steam_tc_curve(i, 42) - Steam_tc_curve(i, 42 - 1)) / (Steam_tc_curve(1, 42) - Steam_tc_curve(1, 42 - 1)) * (Temp - Steam_tc_curve(1, 42 - 1)) + Steam_tc_curve(i, 42 - 1)
                    inter_2 = (Steam_tc_curve(i + 1, 42) - Steam_tc_curve(i + 1, 42 - 1)) / (Steam_tc_curve(1, 42) - Steam_tc_curve(1, 42 - 1)) * (Temp - Steam_tc_curve(1, 42 - 1)) + Steam_tc_curve(i + 1, 42 - 1)
                    inter = (inter_2 - inter_1) / (Steam_tc_curve(i + 1, 1) - Steam_tc_curve(i, 1)) * (Press - Steam_tc_curve(i, 1)) + inter_1
                End If
            Next
        ElseIf Press <= Steam_tc_curve(2, 1) Then
            For j = 2 To 42 - 1 Step 1
                If Steam_tc_curve(1, j) <= Temp And Temp < Steam_tc_curve(1, j + 1) Then
                    inter_1 = (Steam_tc_curve(2, j + 1) - Steam_tc_curve(2, j)) / (Steam_tc_curve(1, j + 1) - Steam_tc_curve(1, j)) * (Temp - Steam_tc_curve(1, j)) + Steam_tc_curve(2, j)
                    inter_2 = (Steam_tc_curve(3, j + 1) - Steam_tc_curve(3, j)) / (Steam_tc_curve(1, j + 1) - Steam_tc_curve(1, j)) * (Temp - Steam_tc_curve(1, j)) + Steam_tc_curve(3, j)
                    inter = (inter_2 - inter_1) / (Steam_tc_curve(3, 1) - Steam_tc_curve(2, 1)) * (Press - Steam_tc_curve(2, 1)) + inter_1
                ElseIf Temp <= Steam_tc_curve(1, 2) Then
                    inter_1 = (Steam_tc_curve(i, 3) - Steam_tc_curve(i, 2)) / (Steam_tc_curve(1, 3) - Steam_tc_curve(1, 2)) * (Temp - Steam_tc_curve(1, 2)) + Steam_tc_curve(i, 2)
                    inter_2 = (Steam_tc_curve(i + 1, 3) - Steam_tc_curve(i + 1, 2)) / (Steam_tc_curve(1, 3) - Steam_tc_curve(1, 2)) * (Temp - Steam_tc_curve(1, 2)) + Steam_tc_curve(i + 1, 2)
                    inter = (inter_2 - inter_1) / (Steam_tc_curve(i + 1, 1) - Steam_tc_curve(i, 1)) * (Press - Steam_tc_curve(i, 1)) + inter_1
                ElseIf Steam_tc_curve(1, 42) <= Temp Then
                    inter_1 = (Steam_tc_curve(i, 42) - Steam_tc_curve(i, 42 - 1)) / (Steam_tc_curve(1, 42) - Steam_tc_curve(1, 42 - 1)) * (Temp - Steam_tc_curve(1, 42 - 1)) + Steam_tc_curve(i, 42 - 1)
                    inter_2 = (Steam_tc_curve(i + 1, 42) - Steam_tc_curve(i + 1, 42 - 1)) / (Steam_tc_curve(1, 42) - Steam_tc_curve(1, 42 - 1)) * (Temp - Steam_tc_curve(1, 42 - 1)) + Steam_tc_curve(i + 1, 42 - 1)
                    inter = (inter_2 - inter_1) / (Steam_tc_curve(i + 1, 1) - Steam_tc_curve(i, 1)) * (Press - Steam_tc_curve(i, 1)) + inter_1
                End If
            Next
        ElseIf Steam_tc_curve(4, 1) <= Press Then
            For j = 2 To 42 - 1 Step 1
                If Steam_tc_curve(1, j) <= Temp And Temp < Steam_tc_curve(1, j + 1) Then
                    inter_1 = (Steam_tc_curve(4 - 1, j + 1) - Steam_tc_curve(4 - 1, j)) / (Steam_tc_curve(1, j + 1) - Steam_tc_curve(1, j)) * (Temp - Steam_tc_curve(1, j)) + Steam_tc_curve(4 - 1, j)
                    inter_2 = (Steam_tc_curve(4, j + 1) - Steam_tc_curve(4, j)) / (Steam_tc_curve(1, j + 1) - Steam_tc_curve(1, j)) * (Temp - Steam_tc_curve(1, j)) + Steam_tc_curve(4, j)
                    inter = (inter_2 - inter_1) / (Steam_tc_curve(4, 1) - Steam_tc_curve(4 - 1, 1)) * (Press - Steam_tc_curve(4 - 1, 1)) + inter_1
                ElseIf Temp <= Steam_tc_curve(1, 2) Then
                    inter_1 = (Steam_tc_curve(i, 3) - Steam_tc_curve(i, 2)) / (Steam_tc_curve(1, 3) - Steam_tc_curve(1, 2)) * (Temp - Steam_tc_curve(1, 2)) + Steam_tc_curve(i, 2)
                    inter_2 = (Steam_tc_curve(i + 1, 3) - Steam_tc_curve(i + 1, 2)) / (Steam_tc_curve(1, 3) - Steam_tc_curve(1, 2)) * (Temp - Steam_tc_curve(1, 2)) + Steam_tc_curve(i + 1, 2)
                    inter = (inter_2 - inter_1) / (Steam_tc_curve(i + 1, 1) - Steam_tc_curve(i, 1)) * (Press - Steam_tc_curve(i, 1)) + inter_1
                ElseIf Steam_tc_curve(1, 42) <= Temp Then
                    inter_1 = (Steam_tc_curve(i, 42) - Steam_tc_curve(i, 42 - 1)) / (Steam_tc_curve(1, 42) - Steam_tc_curve(1, 42 - 1)) * (Temp - Steam_tc_curve(1, 42 - 1)) + Steam_tc_curve(i, 42 - 1)
                    inter_2 = (Steam_tc_curve(i + 1, 42) - Steam_tc_curve(i + 1, 42 - 1)) / (Steam_tc_curve(1, 42) - Steam_tc_curve(1, 42 - 1)) * (Temp - Steam_tc_curve(1, 42 - 1)) + Steam_tc_curve(i + 1, 42 - 1)
                    inter = (inter_2 - inter_1) / (Steam_tc_curve(i + 1, 1) - Steam_tc_curve(i, 1)) * (Press - Steam_tc_curve(i, 1)) + inter_1
                End If
            Next
        End If
    Next
    
    tc = inter

End Sub
