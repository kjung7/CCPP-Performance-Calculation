Attribute VB_Name = "Gas_property_calculation"
Option Explicit

Function idealgas(req_all As String, known1 As String, ByVal val1 As Double, known2 As String, ByVal val2 As Double, Comp As Variant) As Variant
    
    Dim Result() As Variant
    Dim mw_comp(19) As Double
    Dim MW As Double
    Dim i, n As Integer
    Dim req As String
    
    ' Molar mass 계산
    mw_comp(0) = 18.01528   'H2O
    mw_comp(1) = 31.9988    'O2
    mw_comp(2) = 44.0095    'CO2
    mw_comp(3) = 28.0134    'N2
    mw_comp(4) = 39.948     'Ar
    mw_comp(5) = 2.01588    'H2
    mw_comp(6) = 28.0101    'CO
    mw_comp(7) = 16.04246   'CH4
    mw_comp(8) = 30.06904   'C2H6
    mw_comp(9) = 44.09562   'C3H8
    mw_comp(10) = 58.1222   'n-C4H10
    mw_comp(11) = 58.1222   'i-C4H10
    mw_comp(12) = 72.14878  'n-C5H12
    mw_comp(13) = 72.14878  'i-C5H12
    mw_comp(14) = 86.17536  'n-C6H14
    mw_comp(15) = 100.20194 'n-C7H16
    mw_comp(16) = 114.22852 'n-C8H18
    mw_comp(17) = 4.0026    'He
    mw_comp(18) = 34.08088  'H2S
    mw_comp(19) = 64.064
    
    For i = 0 To 19 Step 1
         MW = MW + Comp(i) * mw_comp(i)
    Next

    ' 계산
    n = Len(req_all)
    ReDim Result(n - 1)
    
    For i = 0 To n - 1 Step 1
        req = Mid(req_all, i + 1, 1)
    
        Select Case req
            Case "t"
                Select Case known2
                    Case "h"
                        Result(i) = t_h_gas(val2, Comp, MW)
                    Case "s"
                        Result(i) = T_ps_gas(val1, val2, Comp, MW)
                    Case Else
                        MsgBox "온도를 계산하기 위하여 필요한 물성값이 아님. (known2 = " & known2 & ")"
                        Error 1
                End Select
            Case "p"
                If known1 = "t" And known2 = "s" Then
                    Result(i) = p_ts_gas(val1, val2, Comp, MW)
                ElseIf known1 = "s" And known2 = "t" Then
                    Result(i) = p_ts_gas(val2, val1, Comp, MW)
                Else
                    MsgBox "압력을 계산하기 위하여 필요한 물성값이 아님. (known1 = " & known1 & ", known2 = " & known2 & ")"
                    Error 1
                End If
            Case "h"
                Select Case known2
                    Case "t"
                        Result(i) = h_t_gas(val2, Comp, MW)
                    Case "s"
                        Result(i) = h_ps_gas(val1, val2, Comp, MW)
                    Case Else
                        MsgBox "엔탈피를 계산하기 위하여 필요한 물성값이 아님. (known2 = " & known2 & ")"
                        Error 1
                End Select
            Case "s"
                Select Case known2
                    Case "t"
                        Result(i) = s_pT_gas(val1, val2, Comp, MW)
                    Case "h"
                        Result(i) = s_ph_gas(val1, val2, Comp, MW)
                    Case Else
                        MsgBox "엔트로피를 계산하기 위하여 필요한 물성값이 아님. (known2 = " & known2 & ")"
                        Error 1
                End Select
            Case "c"
                Select Case known2
                    Case "t"
                        Result(i) = c_t_gas(val2, Comp, MW)
                    Case "s"
                        Result(i) = c_ps_gas(val1, val2, Comp, MW)
                    Case "h"
                        Result(i) = c_h_gas(val2, Comp, MW)
                    Case Else
                        MsgBox "등압비열을 계산하기 위하여 필요한 물성값이 아님. (known2 = " & known2 & ")"
                        Error 1
                End Select
            Case "k"
                Select Case known2
                    Case "t"
                        Result(i) = k_t_gas(val2, Comp, MW)
                    Case "s"
                        Result(i) = k_ps_gas(val1, val2, Comp, MW)
                    Case "h"
                        Result(i) = k_h_gas(val2, Comp, MW)
                    Case "c"
                        Result(i) = k_c_gas(val2, Comp, MW)
                    Case Else
                        MsgBox "등압비열을 계산하기 위하여 필요한 물성값이 아님. (known2 = " & known2 & ")"
                        Error 1
                End Select
            Case Else
                MsgBox req & "은 계산을 지원하는 물성이 아님."
                Error 1
        End Select
        
    Next
    
    idealgas = Result
    
End Function

Function k_t_gas(ByVal T As Double, Comp As Variant, MW As Double) As Double
    
    Dim r, C As Double
    
    r = 8.31451         '[J/(mol-K)]
    
    C = c_t_gas(T, Comp, MW)
    k_t_gas = C / (C - r / MW)

End Function

Function k_ps_gas(ByVal P As Double, ByVal s As Double, Comp As Variant, MW As Double) As Double
    
    Dim r As Double
    
    r = 8.31451         '[J/(mol-K)]
    
    C = c_ps_gas(P, s, Comp, MW)
    k_ps_gas = C / (C - r / MW)

End Function

Function k_h_gas(ByVal h As Double, Comp As Variant, MW As Double) As Double
    
    Dim r As Double
    
    r = 8.31451         '[J/(mol-K)]
    
    C = c_h_gas(h, Comp, MW)
    k_h_gas = C / (C - r / MW)
    
End Function

Function k_c_gas(ByVal C As Double, Comp As Variant, MW As Double) As Double

    Dim r As Double
    
    r = 8.31451         '[J/(mol-K)]
    
    k_c_gas = C / (C - r / MW)
    
End Function


Function t_h_gas(ByVal h As Double, Comp As Variant, MW As Double) As Double

    Dim T As Double
    Dim Y, y_for As Double
    Dim iter As Integer
                 
    T = 25  '온도 초기값 설정
    iter = 0

    Do
        iter = iter + 1

        Y = h - h_t_gas(T, Comp, MW)

        If Abs(Y) < 10 ^ (-8) Then
            Exit Do
        ElseIf iter > 100 Then
            MsgBox "t_h_gas 함수의 반복계산 횟수가 100번 이상이 됨."
            Error 1
        End If
        
        y_for = h - h_t_gas(T + 10 ^ (-4), Comp, MW)
        
        T = T - (Y / ((y_for - Y) / 10 ^ (-4)))

    Loop

    t_h_gas = T

End Function

Function T_ps_gas(ByVal P As Double, ByVal s As Double, Comp As Variant, MW As Double) As Double

    Dim T As Double
    Dim Y, y_for As Double
    Dim iter As Integer
                 
    T = 25  '온도 초기값 설정
    iter = 0

    Do
        iter = iter + 1

        Y = s - s_pT_gas(P, T, Comp, MW)

        If Abs(Y) < 10 ^ (-8) Then
            Exit Do
        ElseIf iter > 100 Then
            MsgBox "t_ps_gas 함수의 반복계산 횟수가 100번 이상이 됨."
            Error 1
        End If
        
        y_for = s - s_pT_gas(P, T + 10 ^ (-4), Comp, MW)
    
        T = T - (Y / ((y_for - Y) / 10 ^ (-4)))
    Loop

    T_ps_gas = T

End Function

Function p_ts_gas(ByVal T As Double, ByVal s As Double, Comp As Variant, MW As Double) As Double

    Dim P As Double
    Dim Y, y_for As Double
    Dim iter As Integer
                 
    P = 101.325  '온도 초기값 설정
    iter = 0

    Do
        iter = iter + 1

        Y = s - s_pT_gas(P, T, Comp, MW)

        If Abs(Y) < 10 ^ (-8) Then
            Exit Do
        ElseIf iter > 100 Then
            MsgBox "p_ts_gas 함수의 반복계산 횟수가 100번 이상이 됨."
            Error 1
        End If
        
        y_for = s - s_pT_gas(P + 10 ^ (-4), T, Comp, MW)
    
        P = P - (Y / ((y_for - Y) / 10 ^ (-4)))

    Loop

    p_ts_gas = P

End Function

Function h_ps_gas(ByVal P As Double, ByVal s As Double, Comp As Variant, MW As Double) As Double

    Dim T As Double
    
    T = T_ps_gas(P, s, Comp, MW)
    h_ps_gas = h_t_gas(T, Comp, MW)
    
End Function

Function s_ph_gas(ByVal P As Double, ByVal h As Double, Comp As Variant, MW As Double) As Double

    Dim T As Double
    
    T = t_h_gas(h, Comp, MW)
    s_ph_gas = s_pT_gas(P, T, Comp, MW)
    
End Function

Function c_ps_gas(ByVal P As Double, ByVal s As Double, Comp As Variant, MW As Double) As Double

    Dim T As Double
    
    T = T_ps_gas(P, s, Comp, MW)
    c_ps_gas = c_t_gas(T, Comp, MW)
    
End Function

Function c_h_gas(ByVal h As Double, Comp As Variant, MW As Double) As Double

    Dim T As Double
    
    T = t_h_gas(h, Comp, MW)
    c_h_gas = c_t_gas(T, Comp, MW)
    
End Function

Function h_t_gas(ByVal T As Double, Comp As Variant, MW As Double) As Double
    Dim i, j As Integer
    Dim t_vec(8), h_comp(19) As Double
    Dim r, h_f As Double
    Dim coeff As Variant
    
    T = (T + 273.15)    '[C to K]
    r = 8.31451         '[J/(mol-K)]
    
    t_vec(0) = -(T ^ (-2))
    t_vec(1) = Log(T) / T
    t_vec(2) = 1
    t_vec(3) = T / 2
    t_vec(4) = T ^ 2 / 3
    t_vec(5) = T ^ 3 / 4
    t_vec(6) = T ^ 4 / 5
    t_vec(7) = 1 / T

If T >= 200 And T < 1000 Then '[K]
    For i = 0 To 19 Step 1
        If Not Comp(i) = 0 Then
            Select Case i
                Case 0 'H2O
                    h_f = -241.826 ' [kJ/mol]
                    coeff = Array(-39479.6083, 575.573102, 0.931782653, 0.00722271286, -0.00000734255737, 4.95504349E-09, -1.336933246E-12, -33039.7431)
                Case 1 'O2
                    h_f = 0 ' [kJ/mol]
                    coeff = Array(-34255.6342, 484.700097, 1.119010961, 0.00429388924, -0.000000683630052, -2.0233727E-09, 1.039040018E-12, -3391.45487)
                Case 2 'CO2
                    h_f = -393.51  ' [kJ/mol]
                    coeff = Array(49436.5054, -626.411601, 5.30172524, 0.002503813816, -2.127308728E-07, -7.68998878E-10, 2.849677801E-13, -45281.9846)
                Case 3 'N2
                    h_f = 0 ' [kJ/mol]
                    coeff = Array(22103.71497, -381.846182, 6.08273836, -0.00853091441, 0.00001384646189, -9.62579362E-09, 2.519705809E-12, 710.846086)
                Case 4 'Ar
                    h_f = 0 ' [kJ/mol]
                    coeff = Array(0#, 0#, 2.5, 0#, 0#, 0#, 0#, -745.375)
                Case 5 'H2
                    h_f = 0 ' [kJ/mol]
                    coeff = Array(40783.2321, -800.918604, 8.21470201, -0.01269714457, 0.00001753605076, -1.20286027E-08, 3.36809349E-12, 2682.484665)
                Case 6 'CO
                    h_f = -110.535196 ' [kJ/mol]
                    coeff = Array(14890.45326, -292.2285939, 5.72452717, -0.00817623503, 0.00001456903469, -1.087746302E-08, 3.027941827E-12, -13031.31878)
                Case 7 'CH4
                    h_f = -74.6   ' [kJ/mol]
                    coeff = Array(-176685.0998, 2786.18102, -12.0257785, 0.0391761929, -0.0000361905443, 2.026853043E-08, -4.97670549E-12, -23313.1436)
                Case 8 'C2H6
                    h_f = -83.851544 ' [kJ/mol]
                    coeff = Array(-186204.4161, 3406.19186, -19.51705092, 0.0756583559, -0.0000820417322, 0.000000050611358, -1.319281992E-11, -27029.3289)
                Case 9 'C3H8
                    h_f = -104.68  ' [kJ/mol]
                    coeff = Array(-243314.4337, 4656.27081, -29.39466091, 0.1188952745, -0.0001376308269, 8.81482391E-08, -2.342987994E-11, -35403.3527)
                Case 10 'C4H10
                    h_f = -125.79  ' [kJ/mol]
                    coeff = Array(-317587.254, 6176.33182, -38.9156212, 0.1584654284, -0.0001860050159, 1.199676349E-07, -3.20167055E-11, -45403.6339)
                Case 11 'i-C4H10
                    h_f = -134.99 ' [kJ/mol]
                    coeff = Array(-383446.933, 7000.03964, -44.400269, 0.1746183447, -0.0002078195348, 1.339792433E-07, -3.55168163E-11, -50340.1889)
                Case 12 'n-C5H12
                    h_f = -146.76 ' [kJ/mol]
                    coeff = Array(-276889.4625, 5834.28347, -36.1754148, 0.1533339707, -0.0001528395882, 0.00000008191092, -1.792327902E-11, -46653.7525)
                Case 13 'i-C5H12
                    h_f = -153.7 ' [kJ/mol]
                    coeff = Array(-423190.339, 6497.1891, -36.8112697, 0.1532424729, -0.0001548790714, 8.74989712E-08, -2.07054771E-11, -51554.1659)
                Case 14 'n-C6H14
                    h_f = -166.92 ' [kJ/mol]
                    coeff = Array(-581592.67, 10790.97724, -66.3394703, 0.2523715155, -0.0002904344705, 1.802201514E-07, -4.61722368E-11, -72715.4457)
                Case 15 'n-C7H16
                    h_f = -187.78 ' [kJ/mol]
                    coeff = Array(-612743.289, 11840.85437, -74.871886, 0.2918466052, -0.000341679549, 2.159285269E-07, -5.65585273E-11, -80134.0894)
                Case 16 'n-C8H18
                    h_f = -208.75 ' [kJ/mol]
                    coeff = Array(-698664.715, 13385.01096, -84.1516592, 0.327193666, -0.000377720959, 2.339836988E-07, -6.01089265E-11, -90262.2325)
                Case 17 'He
                    h_f = 0 ' [kJ/mol]
                    coeff = Array(0#, 0#, 2.5, 0#, 0#, 0#, 0#, -745.375)
                Case 18 'H2S
                    h_f = -20.6 ' [kJ/mol]
                    coeff = Array(9543.80881, -68.7517508, 4.05492196, -0.0003014557336, 0.00000376849775, -2.239358925E-09, 3.086859108E-13, -3278.45728)
                Case 19 'SO2
                    h_f = -296.81 ' [kJ/mol]
                    coeff = Array(-53108.4214, 909.031167, -2.356891244, 0.0220445, -0.0000251078, 0.000000014463, -3.36907E-12, -41137.5212)
                     
            
            
            End Select
            
            For j = 0 To 7 Step 1
                h_comp(i) = h_comp(i) + t_vec(j) * coeff(j)
            Next
            h_comp(i) = (h_comp(i) * r * T - h_f * 1000) * Comp(i)

        End If
    Next

ElseIf T >= 1000 And T < 6000 Then '[K]
    For i = 0 To 19 Step 1
        If Not Comp(i) = 0 Then
            Select Case i
                Case 0 ' 'H2O'
                    h_f = -241.826 ' [kJ/mol]
                    coeff = Array(1034972.096, -2412.698562, 4.64611078, 0.002291998307, -0.000000683683048, 9.42646893E-11, -4.82238053E-15, -13842.86509)
                Case 1 ' 'O2'
                    h_f = 0 ' [kJ/mol]
                    coeff = Array(-1037939.022, 2344.830282, 1.819732036, 0.001267847582, -2.188067988E-07, 2.053719572E-11, -8.19346705E-16, -16890.10929)
                Case 2 ' 'CO2'
                    h_f = -393.51  ' [kJ/mol]
                    coeff = Array(117696.2419, -1788.791477, 8.29152319, -0.0000922315678, 4.86367688E-09, -1.891053312E-12, 6.33003659E-16, -39083.5059)
                Case 3 ' 'N2'
                    h_f = 0 ' [kJ/mol]
                    coeff = Array(587712.406, -2239.249073, 6.06694922, -0.00061396855, 1.491806679E-07, -1.923105485E-11, 1.061954386E-15, 12832.10415)
                Case 4 ' 'Ar'
                    h_f = 0 ' [kJ/mol]
                    coeff = Array(20.10538475, -0.0599266107, 2.500069401, -3.99214116E-08, 1.20527214E-11, -1.819015576E-15, 1.078576636E-19, -744.993961)
                Case 5 ' 'H2'
                    h_f = 0 ' [kJ/mol]
                    coeff = Array(560812.801, -837.150474, 2.975364532, 0.001252249124, -0.000000374071619, 5.9366252E-11, -3.6069941E-15, 5339.82441)
                Case 6 ' 'CO'
                    h_f = -110.535196 ' [kJ/mol]
                    coeff = Array(461919.725, -1944.704863, 5.91671418, -0.000566428283, 0.000000139881454, -1.787680361E-11, 9.62093557E-16, -2466.261084)
                Case 7 ' 'CH4'
                    h_f = -74.6   ' [kJ/mol]
                    coeff = Array(3730042.76, -13835.01485, 20.49107091, -0.001961974759, 0.000000472731304, -3.72881469E-11, 1.623737207E-15, 75320.6691)
                Case 8 ' 'C2H6'
                    h_f = -83.851544 ' [kJ/mol]
                    coeff = Array(5025782.13, -20330.22397, 33.2255293, -0.00383670341, 0.000000723840586, -7.3191825E-11, 3.065468699E-15, 111596.395)
                Case 9 ' 'C3H8'
                    h_f = -104.68  ' [kJ/mol]
                    coeff = Array(6420731.68, -26597.91134, 45.3435684, -0.00502066392, 0.000000947121694, -9.57540523E-11, 4.00967288E-15, 145558.2459)
                Case 10 ' 'C4H10'
                    h_f = -125.79  ' [kJ/mol]
                    coeff = Array(7682322.45, -32560.5151, 57.3673275, -0.00619791681, 0.000001180186048, -1.221893698E-10, 5.25063525E-15, 177452.656)
                Case 11 'i-C4H10
                    h_f = -134.99 ' [kJ/mol]
                    coeff = Array(7528018.92, -32025.1706, 57.00161, -0.00606001309, 0.000001143975809, -1.157061835E-10, 4.84604291E-15, 172850.0802)
                Case 12 'n-C5H12
                    h_f = -146.76 ' [kJ/mol]
                    coeff = Array(-2530779.286, -8972.59326, 45.3622326, -0.002626989916, 0.000003135136419, -5.31872894E-10, 2.886896868E-14, 14846.16529)
                Case 13 'i-C5H12
                    h_f = -153.7 ' [kJ/mol]
                    coeff = Array(11568885.94, -45562.4687, 74.9544363, -0.00784541558, 0.000001444393314, -1.464370213E-10, 6.230285E-15, 254492.7135)
                Case 14 'n-C6H14
                    h_f = -166.92 ' [kJ/mol]
                    coeff = Array(-3106625.684, -7346.08792, 46.9413176, 0.001693963977, 0.000002068996667, -4.21214168E-10, 2.452345845E-14, 523.750312)
                Case 15 'n-C7H16
                    h_f = -187.78 ' [kJ/mol]
                    coeff = Array(9135632.47, -39233.1969, 78.8978085, -0.00465425193, 0.000002071774142, -3.4425393E-10, 1.976834775E-14, 205070.8295)
                Case 16 'n-C8H18
                    h_f = -208.75 ' [kJ/mol]
                    coeff = Array(6365406.95, -31053.64657, 69.6916234, 0.01048059637, -0.00000412962195, 5.54322632E-10, -2.651436499E-14, 150096.8785)
                Case 17 'He
                    h_f = 0 ' [kJ/mol]
                    coeff = Array(560812.801, -837.150474, 2.975364532, 0.001252249124, -0.000000374071619, 5.9366252E-11, -3.6069941E-15, 5339.82441)
                Case 18 'H2S
                    h_f = -20.6 ' [kJ/mol]
                    coeff = Array(1430040.22, -5284.02865, 10.16182124, -0.000970384996, 2.154003405E-07, -2.1696957E-11, 9.31816307E-16, 29086.96214)
                Case 19 'SO2
                    h_f = -296.81 ' [kJ/mol]
                    coeff = Array(-112764.0116, -825.226138, 7.61617863, -0.000199932761, 5.65563143E-08, -5.45431661E-12, 2.918294102E-16, -33513.0869)
            
            
            End Select
            
            For j = 0 To 7 Step 1
                h_comp(i) = h_comp(i) + t_vec(j) * coeff(j)
            Next
            h_comp(i) = (h_comp(i) * r * T - h_f * 1000) * Comp(i)

        End If
    Next
Else
    MsgBox "200~6000[K] 사이의 값을 입력하시오. (입력된 온도 = " & CStr(T) & " [K])"
    Error 1
End If

h_t_gas = 0
For i = 0 To 19 Step 1
    h_t_gas = h_t_gas + h_comp(i)
Next
h_t_gas = h_t_gas / MW

End Function

Function s_pT_gas(P As Double, ByVal T As Double, Comp As Variant, MW As Double) As Double
    Dim i, j As Integer
    Dim t_vec(8), s_comp(19) As Double
    Dim r, p_ref As Double
    Dim coeff As Variant
    
    T = (T + 273.15)    '[C to K]
    r = 8.31451         '[J/(mol-K)]
    p_ref = 101.325     '대기압 ISO condition
    
    t_vec(0) = (-T ^ (-2)) / 2
    t_vec(1) = -T ^ (-1)
    t_vec(2) = Log(T)
    t_vec(3) = T
    t_vec(4) = T ^ 2 / 2
    t_vec(5) = T ^ 3 / 3
    t_vec(6) = T ^ 4 / 4
    t_vec(7) = 1

If T >= 200 And T < 1000 Then '[K]
    For i = 0 To 19 Step 1
        If Not Comp(i) = 0 Then
            Select Case i
                Case 0 'H2O
                    coeff = Array(-39479.6083, 575.573102, 0.931782653, 0.00722271286, -0.00000734255737, 4.95504349E-09, -1.336933246E-12, 17.24205775)
                Case 1 'O2
                    coeff = Array(-34255.6342, 484.700097, 1.119010961, 0.00429388924, -0.000000683630052, -2.0233727E-09, 1.039040018E-12, 18.4969947)
                Case 2 'CO2
                    coeff = Array(49436.5054, -626.411601, 5.30172524, 0.002503813816, -2.127308728E-07, -7.68998878E-10, 2.849677801E-13, -7.04827944)
                Case 3 'N2
                    coeff = Array(22103.71497, -381.846182, 6.08273836, -0.00853091441, 0.00001384646189, -9.62579362E-09, 2.519705809E-12, -10.76003744)
                Case 4 'Ar
                    coeff = Array(0#, 0#, 2.5, 0#, 0#, 0#, 0#, 4.37967491)
                Case 5 'H2
                    coeff = Array(40783.2321, -800.918604, 8.21470201, -0.01269714457, 0.00001753605076, -1.20286027E-08, 3.36809349E-12, -30.43788844)
                Case 6 'CO
                    coeff = Array(14890.45326, -292.2285939, 5.72452717, -0.00817623503, 0.00001456903469, -1.087746302E-08, 3.027941827E-12, -7.85924135)
                Case 7 'CH4
                    coeff = Array(-176685.0998, 2786.18102, -12.0257785, 0.0391761929, -0.0000361905443, 2.026853043E-08, -4.97670549E-12, 89.0432275)
                Case 8 'C2H6
                    coeff = Array(-186204.4161, 3406.19186, -19.51705092, 0.0756583559, -0.0000820417322, 0.000000050611358, -1.319281992E-11, 129.8140496)
                Case 9 'C3H8
                    coeff = Array(-243314.4337, 4656.27081, -29.39466091, 0.1188952745, -0.0001376308269, 8.81482391E-08, -2.342987994E-11, 184.1749277)
                Case 10 'C4H10
                    coeff = Array(-317587.254, 6176.33182, -38.9156212, 0.1584654284, -0.0001860050159, 1.199676349E-07, -3.20167055E-11, 237.9488665)
                Case 11 'i-C4H10
                    coeff = Array(-383446.933, 7000.03964, -44.400269, 0.1746183447, -0.0002078195348, 1.339792433E-07, -3.55168163E-11, 265.8966497)
                Case 12 'n-C5H12
                    coeff = Array(-276889.4625, 5834.28347, -36.1754148, 0.1533339707, -0.0001528395882, 0.00000008191092, -1.792327902E-11, 226.5544053)
                Case 13 'i-C5H12
                    coeff = Array(-423190.339, 6497.1891, -36.8112697, 0.1532424729, -0.0001548790714, 8.74989712E-08, -2.07054771E-11, 230.9518218)
                Case 14 'n-C6H14
                    coeff = Array(-581592.67, 10790.97724, -66.3394703, 0.2523715155, -0.0002904344705, 1.802201514E-07, -4.61722368E-11, 393.828354)
                Case 15 'n-C7H16
                    coeff = Array(-612743.289, 11840.85437, -74.871886, 0.2918466052, -0.000341679549, 2.159285269E-07, -5.65585273E-11, 440.721332)
                Case 16 'n-C8H18
                    coeff = Array(-698664.715, 13385.01096, -84.1516592, 0.327193666, -0.000377720959, 2.339836988E-07, -6.01089265E-11, 493.922214)
                Case 17 'He
                    coeff = Array(0#, 0#, 2.5, 0#, 0#, 0#, 0#, 0.928723974)
                Case 18 'H2S
                    coeff = Array(9543.80881, -68.7517508, 4.05492196, -0.0003014557336, 0.00000376849775, -2.239358925E-09, 3.086859108E-13, 1.415194691)
                Case 19 'SO2
                    coeff = Array(-53108.4214, 909.031167, -2.356891244, 0.0220445, -0.0000251078, 0.000000014463, -3.36907E-12, 40.45512519)
            
            End Select
            
            For j = 0 To 7 Step 1
                s_comp(i) = s_comp(i) + (t_vec(j) * coeff(j))
            Next
            s_comp(i) = (s_comp(i) * r - r * Log(P * Comp(i) / p_ref)) * Comp(i)

        End If
    Next

ElseIf T >= 1000 And T < 6000 Then '[K]
    For i = 0 To 19 Step 1
        If Not Comp(i) = 0 Then
            Select Case i
                Case 0 'H2O
                    coeff = Array(1034972.096, -2412.698562, 4.64611078, 0.002291998307, -0.000000683683048, 9.42646893E-11, -4.82238053E-15, -7.97814851)
                Case 1 'O2
                    coeff = Array(-1037939.022, 2344.830282, 1.819732036, 0.001267847582, -2.188067988E-07, 2.053719572E-11, -8.19346705E-16, 17.38716506)
                Case 2 'CO2
                    coeff = Array(117696.2419, -1788.791477, 8.29152319, -0.0000922315678, 4.86367688E-09, -1.891053312E-12, 6.33003659E-16, -26.52669281)
                Case 3 'N2
                    coeff = Array(587712.406, -2239.249073, 6.06694922, -0.00061396855, 1.491806679E-07, -1.923105485E-11, 1.061954386E-15, -15.86640027)
                Case 4 'Ar
                    coeff = Array(20.10538475, -0.0599266107, 2.500069401, -3.99214116E-08, 1.20527214E-11, -1.819015576E-15, 1.078576636E-19, 4.37918011)
                Case 5 'H2
                    coeff = Array(560812.801, -837.150474, 2.975364532, 0.001252249124, -0.000000374071619, 5.9366252E-11, -3.6069941E-15, -2.202774769)
                Case 6 'CO
                    coeff = Array(461919.725, -1944.704863, 5.91671418, -0.000566428283, 0.000000139881454, -1.787680361E-11, 9.62093557E-16, -13.87413108)
                Case 7 'CH4
                    coeff = Array(3730042.76, -13835.01485, 20.49107091, -0.001961974759, 0.000000472731304, -3.72881469E-11, 1.623737207E-15, -121.9124889)
                Case 8 'C2H6
                    coeff = Array(5025782.13, -20330.22397, 33.2255293, -0.00383670341, 0.000000723840586, -7.3191825E-11, 3.065468699E-15, -203.9410584)
                Case 9 'C3H8
                    coeff = Array(6420731.68, -26597.91134, 45.3435684, -0.00502066392, 0.000000947121694, -9.57540523E-11, 4.00967288E-15, -281.8374734)
                Case 10 'C4H10
                    coeff = Array(7682322.45, -32560.5151, 57.3673275, -0.00619791681, 0.000001180186048, -1.221893698E-10, 5.25063525E-15, -358.791876)
                Case 11 'i-C4H10
                    coeff = Array(7528018.92, -32025.1706, 57.00161, -0.00606001309, 0.000001143975809, -1.157061835E-10, 4.84604291E-15, -357.617689)
                Case 12 'n-C5H12
                    coeff = Array(-2530779.286, -8972.59326, 45.3622326, -0.002626989916, 0.000003135136419, -5.31872894E-10, 2.886896868E-14, -251.6550384)
                Case 13 'i-C5H12
                    coeff = Array(11568885.94, -45562.4687, 74.9544363, -0.00784541558, 0.000001444393314, -1.464370213E-10, 6.230285E-15, -480.198578)
                Case 14 'n-C6H14
                    coeff = Array(-3106625.684, -7346.08792, 46.9413176, 0.001693963977, 0.000002068996667, -4.21214168E-10, 2.452345845E-14, -254.9967718)
                Case 15 'n-C7H16
                    coeff = Array(9135632.47, -39233.1969, 78.8978085, -0.00465425193, 0.000002071774142, -3.4425393E-10, 1.976834775E-14, -485.110402)
                Case 16 'n-C8H18
                    coeff = Array(6365406.95, -31053.64657, 69.6916234, 0.01048059637, -0.00000412962195, 5.54322632E-10, -2.651436499E-14, -416.989565)
                Case 17 'He
                    coeff = Array(560812.801, -837.150474, 2.975364532, 0.001252249124, -0.000000374071619, 5.9366252E-11, -3.6069941E-15, 0.928723974)
                Case 18 'H2S
                    coeff = Array(1430040.22, -5284.02865, 10.16182124, -0.000970384996, 2.154003405E-07, -2.1696957E-11, 9.31816307E-16, -43.49160391)
                                Case 19 'SO2
                    coeff = Array(-112764.0116, -825.226138, 7.61617863, -0.000199932761, 5.65563143E-08, -5.45431661E-12, 2.918294102E-16, -16.55776085)
            
            End Select
            
            For j = 0 To 7 Step 1
                s_comp(i) = s_comp(i) + (t_vec(j) * coeff(j))
            Next
            s_comp(i) = (s_comp(i) * r - r * Log(P * Comp(i) / p_ref)) * Comp(i)

        End If
    Next
Else
    MsgBox "200~6000[K] 사이의 값을 입력하시오. (입력된 온도 = " & CStr(T) & " [K])"
    Error 1
End If

s_pT_gas = 0
For i = 0 To 19 Step 1
    s_pT_gas = s_pT_gas + s_comp(i)
Next
s_pT_gas = s_pT_gas / MW

End Function

Function c_t_gas(ByVal T As Double, Comp As Variant, MW As Double) As Double
    Dim i, j As Integer
    Dim t_vec(7), c_comp(19) As Double
    Dim r As Double
    Dim coeff As Variant
    
    T = (T + 273.15)    '[C to K]
    r = 8.31451         '[J/(mol-K)]
    
    t_vec(0) = T ^ (-2)
    t_vec(1) = T ^ (-1)
    t_vec(2) = 1
    t_vec(3) = T
    t_vec(4) = T ^ 2
    t_vec(5) = T ^ 3
    t_vec(6) = T ^ 4

If T >= 200 And T < 1000 Then '[K]
    For i = 0 To 19 Step 1
        If Not Comp(i) = 0 Then
            Select Case i
                Case 0 'H2O
                    coeff = Array(-39479.6083, 575.573102, 0.931782653, 0.00722271286, -0.00000734255737, 4.95504349E-09, -1.336933246E-12)
                Case 1 'O2
                    coeff = Array(-34255.6342, 484.700097, 1.119010961, 0.00429388924, -0.000000683630052, -2.0233727E-09, 1.039040018E-12)
                Case 2 'CO2
                    coeff = Array(49436.5054, -626.411601, 5.30172524, 0.002503813816, -2.127308728E-07, -7.68998878E-10, 2.849677801E-13)
                Case 3 'N2
                    coeff = Array(22103.71497, -381.846182, 6.08273836, -0.00853091441, 0.00001384646189, -9.62579362E-09, 2.519705809E-12)
                Case 4 'Ar
                    coeff = Array(0#, 0#, 2.5, 0#, 0#, 0#, 0#, -745.375)
                Case 5 'H2
                    coeff = Array(40783.2321, -800.918604, 8.21470201, -0.01269714457, 0.00001753605076, -1.20286027E-08, 3.36809349E-12)
                Case 6 'CO
                    coeff = Array(14890.45326, -292.2285939, 5.72452717, -0.00817623503, 0.00001456903469, -1.087746302E-08, 3.027941827E-12)
                Case 7 'CH4
                    coeff = Array(-176685.0998, 2786.18102, -12.0257785, 0.0391761929, -0.0000361905443, 2.026853043E-08, -4.97670549E-12)
                Case 8 'C2H6
                    coeff = Array(-186204.4161, 3406.19186, -19.51705092, 0.0756583559, -0.0000820417322, 0.000000050611358, -1.319281992E-11)
                Case 9 'C3H8
                    coeff = Array(-243314.4337, 4656.27081, -29.39466091, 0.1188952745, -0.0001376308269, 8.81482391E-08, -2.342987994E-11)
                Case 10 'n-C4H10
                    coeff = Array(-317587.254, 6176.33182, -38.9156212, 0.1584654284, -0.0001860050159, 1.199676349E-07, -3.20167055E-11)
                Case 11 'i-C4H10
                    coeff = Array(-383446.933, 7000.03964, -44.400269, 0.1746183447, -0.0002078195348, 1.339792433E-07, -3.55168163E-11)
                Case 12 'n-C5H12
                    coeff = Array(-276889.4625, 5834.28347, -36.1754148, 0.1533339707, -0.0001528395882, 0.00000008191092, -1.792327902E-11)
                Case 13 'i-C5H12
                    coeff = Array(-423190.339, 6497.1891, -36.8112697, 0.1532424729, -0.0001548790714, 8.74989712E-08, -2.07054771E-11)
                Case 14 'n-C6H14
                    coeff = Array(-581592.67, 10790.97724, -66.3394703, 0.2523715155, -0.0002904344705, 1.802201514E-07, -4.61722368E-11)
                Case 15 'n-C7H16
                    coeff = Array(-612743.289, 11840.85437, -74.871886, 0.2918466052, -0.000341679549, 2.159285269E-07, -5.65585273E-11)
                Case 16 'n-C8H18
                    coeff = Array(-698664.715, 13385.01096, -84.1516592, 0.327193666, -0.000377720959, 2.339836988E-07, -6.01089265E-11)
                Case 17 'He
                    coeff = Array(0#, 0#, 2.5, 0#, 0#, 0#, 0#)
                Case 18 'H2S
                    coeff = Array(9543.80881, -68.7517508, 4.05492196, -0.0003014557336, 0.00000376849775, -2.239358925E-09, 3.086859108E-13)
                Case 19 'SO2
                    coeff = Array(-53108.4214, 909.031167, -2.356891244, 0.0220445, -0.0000251078, 0.000000014463, -3.36907E-12)
            
            End Select
            
            For j = 0 To 6 Step 1
                c_comp(i) = c_comp(i) + t_vec(j) * coeff(j)
            Next
            c_comp(i) = c_comp(i) * r * Comp(i)

        End If
    Next

ElseIf T >= 1000 And T < 6000 Then '[K]
    For i = 0 To 19 Step 1
        If Not Comp(i) = 0 Then
            Select Case i
                Case 0 ' 'H2O'
                    coeff = Array(1034972.096, -2412.698562, 4.64611078, 0.002291998307, -0.000000683683048, 9.42646893E-11, -4.82238053E-15)
                Case 1 ' 'O2'
                    coeff = Array(-1037939.022, 2344.830282, 1.819732036, 0.001267847582, -2.188067988E-07, 2.053719572E-11, -8.19346705E-16)
                Case 2 ' 'CO2'
                    coeff = Array(117696.2419, -1788.791477, 8.29152319, -0.0000922315678, 4.86367688E-09, -1.891053312E-12, 6.33003659E-16)
                Case 3 ' 'N2'
                    coeff = Array(587712.406, -2239.249073, 6.06694922, -0.00061396855, 1.491806679E-07, -1.923105485E-11, 1.061954386E-15)
                Case 4 ' 'Ar'
                    coeff = Array(20.10538475, -0.0599266107, 2.500069401, -3.99214116E-08, 1.20527214E-11, -1.819015576E-15, 1.078576636E-19)
                Case 5 ' 'H2'
                    coeff = Array(560812.801, -837.150474, 2.975364532, 0.001252249124, -0.000000374071619, 5.9366252E-11, -3.6069941E-15)
                Case 6 ' 'CO'
                    coeff = Array(461919.725, -1944.704863, 5.91671418, -0.000566428283, 0.000000139881454, -1.787680361E-11, 9.62093557E-16)
                Case 7 ' 'CH4'
                    coeff = Array(3730042.76, -13835.01485, 20.49107091, -0.001961974759, 0.000000472731304, -3.72881469E-11, 1.623737207E-15)
                Case 8 ' 'C2H6'
                    coeff = Array(5025782.13, -20330.22397, 33.2255293, -0.00383670341, 0.000000723840586, -7.3191825E-11, 3.065468699E-15)
                Case 9 ' 'C3H8'
                    coeff = Array(6420731.68, -26597.91134, 45.3435684, -0.00502066392, 0.000000947121694, -9.57540523E-11, 4.00967288E-15)
                Case 10 ' 'n-C4H10'
                    coeff = Array(7682322.45, -32560.5151, 57.3673275, -0.00619791681, 0.000001180186048, -1.221893698E-10, 5.25063525E-15)
                Case 11 'i-C4H10
                    coeff = Array(7528018.92, -32025.1706, 57.00161, -0.00606001309, 0.000001143975809, -1.157061835E-10, 4.84604291E-15)
                Case 12 'n-C5H12
                    coeff = Array(-2530779.286, -8972.59326, 45.3622326, -0.002626989916, 0.000003135136419, -5.31872894E-10, 2.886896868E-14)
               Case 13 'i-C5H12
                    coeff = Array(11568885.94, -45562.4687, 74.9544363, -0.00784541558, 0.000001444393314, -1.464370213E-10, 6.230285E-15)
                Case 14 'n-C6H14
                    coeff = Array(-3106625.684, -7346.08792, 46.9413176, 0.001693963977, 0.000002068996667, -4.21214168E-10, 2.452345845E-14)
                Case 15 'n-C7H16
                    coeff = Array(9135632.47, -39233.1969, 78.8978085, -0.00465425193, 0.000002071774142, -3.4425393E-10, 1.976834775E-14)
                Case 16 'n-C8H18
                    coeff = Array(6365406.95, -31053.64657, 69.6916234, 0.01048059637, -0.00000412962195, 5.54322632E-10, -2.651436499E-14)
                Case 17 'He
                    coeff = Array(0#, 0#, 2.5, 0#, 0#, 0#, 0#)
                Case 18 'H2S
                    coeff = Array(1430040.22, -5284.02865, 10.16182124, -0.000970384996, 2.154003405E-07, -2.1696957E-11, 9.31816307E-16)
                Case 19 'SO2
                    coeff = Array(-112764.0116, -825.226138, 7.61617863, -0.000199932761, 5.65563143E-08, -5.45431661E-12, 2.918294102E-16)
            
            End Select
            
            For j = 0 To 6 Step 1
                c_comp(i) = c_comp(i) + t_vec(j) * coeff(j)
            Next
            c_comp(i) = c_comp(i) * r * Comp(i)

        End If
    Next
Else
    MsgBox "200~6000[K] 사이의 값을 입력하시오. (입력된 온도 = " & CStr(T) & " [K])"
    Error 1
End If

c_t_gas = 0
For i = 0 To 19 Step 1
    c_t_gas = c_t_gas + c_comp(i)
Next
c_t_gas = c_t_gas / MW

End Function
'viscosity calculation'''''''''''''''''''''''''''''''''''''''''''''''''''
Function mu_t_gas(ByVal T As Double, Comp As Variant) As Double
    
    Dim dummy1, dummy2, bunmo(5) As Variant
    Dim i, j As Integer
    
    dummy1 = pi(T)
    dummy2 = mu_t(T)
    For i = 0 To 5 Step 1
        For j = 0 To 5 Step 1
            bunmo(i) = bunmo(i) + Comp(j) * dummy1(i, j)
        Next
        mu_t_gas = mu_t_gas + Comp(i) * dummy2(i) / bunmo(i)
    Next
End Function
'thermal conductivity calculation'''''''''''''''''''''''''''''''''''''''''''''''''''
Function tc_t_gas(ByVal T As Double, Comp As Variant) As Double
    
    Dim dummy1, dummy2, bunmo(5) As Variant
    Dim i, j As Integer
    
    dummy1 = pi(T)
    dummy2 = tc_t(T)
    For i = 0 To 5 Step 1
        For j = 0 To 5 Step 1
            bunmo(i) = bunmo(i) + Comp(j) * dummy1(i, j)
        Next
        tc_t_gas = tc_t_gas + Comp(i) * dummy2(i) / bunmo(i)
    Next
End Function

Function tc_t(ByVal T As Double) As Variant

    Dim tc(5) As Variant
     
'H2O conductivity

    tc(0) = -5.4481917E-11 * (T) ^ 4 - 0.000000036087773 * (T) ^ 3 + 0.00010085922 * (T) ^ 2 + 0.064492661 * (T) ^ 1 + 16.232049

'O2 conductivity
 
    tc(1) = -1.2013725E-11 * (T) ^ 4 + 0.000000032072343 * (T) ^ 3 - 0.000039989704 * (T) ^ 2 + 0.082277624 * (T) ^ 1 + 24.717432
   
'CO2 conductivity

    tc(2) = 5.9123502E-11 * (T) ^ 4 - 0.0000001447979 * (T) ^ 3 + 0.000089571773 * (T) ^ 2 + 0.062993467 * (T) ^ 1 + 14.619913
    
'N2 conductivity

    tc(3) = -1.2994786E-11 * (T) ^ 4 + 0.000000035592206 * (T) ^ 3 - 0.000047728976 * (T) ^ 2 + 0.07682124 * (T) ^ 1 + 23.211893
    
'Ar conductivity

    tc(4) = -8.4018735E-11 * (T) ^ 4 + 0.00000016062633 * (T) ^ 3 - 0.00010003051 * (T) ^ 2 + 0.058607988 * (T) ^ 1 + 16.118886
    
'H2 conductivity
    
    tc(5) = -5.7986824E-10 * (T) ^ 4 + 0.0000011836847 * (T) ^ 3 - 0.00078397288 * (T) ^ 2 + 0.54872465 * (T) ^ 1 + 162.65257
    
    tc_t = tc
    
End Function

Function mu_t(ByVal T As Double) As Variant

    ReDim mu(5) As Variant
   
'H2O Viscosity
    
    mu(0) = -0.00000000297983 * (T) ^ (3) + 0.0000089545 * (T) ^ (2) + 0.0370732 * (T) ^ (1) + 11.0734
    
'O2 Viscosity

    mu(1) = 0.0000000179218 * (T) ^ (3) - 0.0000343466 * (T) ^ (2) + 0.0561867 * (T) ^ (1) + 18.6644
    
'CO2 Viscosity
    
    mu(2) = 0.0000000125012 * (T) ^ (3) - 0.0000273942 * (T) ^ (2) + 0.0475277 * (T) ^ (1) + 13.6821

'N2 Viscosity

    mu(3) = 0.00000000407019 * (T) ^ (3) - 0.0000163843 * (T) ^ (2) + 0.0414597 * (T) ^ (1) + 16.8188
    
'Ar Viscosity

    mu(4) = 0.0000000185642 * (T) ^ (3) - 0.0000381661 * (T) ^ (2) + 0.062455 * (T) ^ (1) + 21.2913
    
'H2 Viscosity

    mu(5) = 0.00000000265911 * (T) ^ (3) - 0.0000080425 * (T) ^ (2) + 0.019696 * (T) ^ (1) + 8.4725
    
    mu_t = mu
    
End Function

Function pi(ByVal T As Double) As Variant
    
    Dim i, j As Integer
    Dim dummy, MW(5), Result(5, 5) As Variant
    
    dummy = mu_t(T)
    MW(0) = 18.01528
    MW(1) = 31.9988
    MW(2) = 44.0095
    MW(3) = 28.0134
    MW(4) = 39.948
    MW(5) = 2.01588
    
    For i = 0 To 5 Step 1
        For j = 0 To 5 Step 1
            If i = j Then
                Result(i, j) = 1
            Else
                Result(i, j) = (1 + ((dummy(i) / dummy(j)) ^ 0.5 * (MW(i) / MW(j)) ^ 0.25)) ^ 2 / (8 * (1 + (MW(j) / MW(i)))) ^ 0.5
            End If
        Next
    Next
    
    pi = Result
    
End Function




