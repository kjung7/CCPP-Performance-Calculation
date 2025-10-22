VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHRSG_read_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public dict As New Dictionary

Public DB As New clsDB_
Public SCR As New clsSCR_
Public SPHT As New clsSPHT_
Public EVAP As New clsEVAP_
Public ECON As New clsECON_

Public DESH As New clsDESH_
Public MIX As New clsMIX_
Public SPL As New clsSPL_
Public MIX_gas As New clsMIX_gas_

Public PUMP As New clsPUMP_
Public DA As New clsDA_
Public CPH As New clsCPH_
Public GLD As New clsGLD_

Public Number As Integer
Public index, index_p, index_c As Variant
Public E_loss, P_loss_hot As Variant
Public P_loss_valve, P_loss_cold As Variant
Public Calc_option, Input_value As Variant

Public index_S, Stream_S, Valve As Variant
Public Calc_option_T, Input_value_T, index_T, Stream_T As Variant

Public D_o, T_tube, Material, L_tube, n_tube, x_tube As Variant

Public DA_type As Variant
Public P_HP_DSN, P_IP_DSN, P_LP_DSN, EVAP_Cir_ratio, EVAP_Head, DA_Head As Double
Public i_Gas_SPL, i_Para_HX_1 As Integer
Public index_Para_HX_1, index_Para_HX_2 As Variant
Public Gas_ratio_1, Gas_ratio_2 As Double

Public n_HP_end, n_IP_end, n_LP_end, n_None_end, e_LP, n_HX As Integer
Public P_HP_off, P_IP_off, P_LP_off As Double
Public n_HP_start, n_IP_start, n_LP_start, e_EVAP_BD, n_DA As Integer
Public Method_P_loss_hot, Method_P_loss_valve, Method_P_loss_cold, BFP_Cal_option As Variant

Public Mass_sec As Variant
Public n_DESH, i_HP_DESH, i_IP_DESH As Integer

Function read_data() As Dictionary

    Dim HRSG_input As Worksheet
    Dim HRSG_input_off As Worksheet
    
    Dim rg As Range
    
    Dim i As Integer
    Dim skey As String
    Dim key_p As String
    Dim key_c As String
    Dim key As Variant
    Dim P_STM_Level As Double
    
    Set HRSG_input = ThisWorkbook.Worksheets("HRSG_input")
    Set HRSG_input_off = ThisWorkbook.Worksheets("HRSG_input_off")
    
    Set rg = HRSG_input.Range(HRSG_input.Range("A2"), HRSG_input.Range("A2").End(xlDown))

    Number = rg.Rows.Count

    ReDim index(Number)
    ReDim index_c(Number)
    ReDim index_p(Number)
    
    ReDim E_loss(Number)
    ReDim P_loss_hot(Number)
    ReDim P_loss_valve(Number)
    ReDim P_loss_cold(Number)
    
    ReDim Calc_option(Number)
    ReDim Input_value(Number)
    
    ReDim index_S(Number)
    ReDim Stream_S(Number)
    ReDim Valve(Number)
    
    ReDim Calc_option_T(Number)
    ReDim Input_value_T(Number)
    ReDim index_T(Number)
    ReDim Stream_T(Number)
    
    ReDim D_o(Number)
    ReDim T_tube(Number)
    ReDim Material(Number)
    ReDim L_tube(Number)
    ReDim n_tube(Number)
    ReDim x_tube(Number)
    
    For i = 1 To Number

        skey = rg(i, 1) & "_" & rg(i, 2)
        key_p = rg(i, 1)
        key_c = rg(i, 2)

        index(i) = skey
        index_c(i) = key_c
        index_p(i) = key_p

        E_loss(i) = rg(i, 3)
        
        If HRSG_input.Cells(31, "X") = "-" Then
            P_loss_hot(i) = rg(i, 4)
        ElseIf HRSG_input.Cells(31, "X") = "mbar" Then
            P_loss_hot(i) = rg(i, 4) * 0.1 'mbar to kPa
        End If
        
        If HRSG_input.Cells(32, "X") = "-" Then
            P_loss_valve(i) = rg(i, 5)
        ElseIf HRSG_input.Cells(32, "X") = "bar" Then
            P_loss_valve(i) = rg(i, 5) * 100 'bar to kPa
        End If
        
        If HRSG_input.Cells(33, "X") = "-" Then
            P_loss_cold(i) = rg(i, 6)
        ElseIf HRSG_input.Cells(33, "X") = "bar" Then
            P_loss_cold(i) = rg(i, 6) * 100 'bar to kPa
        End If
        
        Calc_option(i) = rg(i, 7)
        Input_value(i) = rg(i, 8)
        
        index_S(i) = rg(i, 9) 'Secondary Flow Pressure Level _ Component Name
        Stream_S(i) = rg(i, 10) 'Secondary Flow Stream
        Valve(i) = rg(i, 11) 'Secondary Flow Pressure Control
        
        Calc_option_T(i) = rg(i, 12)
        Input_value_T(i) = rg(i, 13)
        index_T(i) = rg(i, 14) 'Tertiary Flow Pressure Level _ Component Name
        Stream_T(i) = rg(i, 15) 'Tertiary Flow Stream
        
        
        'Tube diameter, meter basis
         If HRSG_input.Cells(35, "X") = "mm" Then
             D_o(i) = rg(i, 16) * 0.001
        ElseIf HRSG_input.Cells(35, "X") = "inch" Then
             D_o(i) = rg(i, 16) * 25.4 * 0.001
        End If
              
        'Tube thcikness, meter basis
         If HRSG_input.Cells(36, "X") = "mm" Then
             T_tube(i) = rg(i, 17) * 0.001
        ElseIf HRSG_input.Cells(36, "X") = "inch" Then
             T_tube(i) = rg(i, 17) * 25.4 * 0.001
        End If
                            
        'Tube length, meter basis
         If HRSG_input.Cells(37, "X") = "m" Then
             L_tube(i) = rg(i, 19)
        ElseIf HRSG_input.Cells(37, "X") = "inch" Then
             L_tube(i) = rg(i, 19) * 0.0254
        End If
                       
        
      '  D_o(i) = rg(i, 16) * 0.001
      '  T_tube(i) = rg(i, 17) * 0.001
        Material(i) = rg(i, 18)
       ' L_tube(i) = rg(i, 19)
        n_tube(i) = rg(i, 20)
        x_tube(i) = rg(i, 21)

        If dict.Exists(skey) = True Then
            
            If key_c = "DB" Then 'DuctBurner
                Set DB = New clsDB_
                Set dict(skey) = DB
            
            ElseIf key_c = "SCR" Then 'DuctBurner
                Set SCR = New clsSCR_
                Set dict(skey) = SCR
                
            ElseIf key_c = "SPHT" Or key_c = "SPHT_RH" Or key_c = "SPHT_1" Or key_c = "SPHT_2" Or key_c = "SPHT_3" Or key_c = "HRH" Or key_c = "CRH" Or key_c = "CRH_1" Or key_c = "CRH_2" Then
                Set SPHT = New clsSPHT_
                Set dict(skey) = SPHT
                
            ElseIf key_c = "EVAP" Or key_c = "EVAP_BD" Then
                Set EVAP = New clsEVAP_
                Set dict(skey) = EVAP
                
            ElseIf key_c = "ECON" Or key_c = "ECON_1" Or key_c = "ECON_2" Or key_c = "ECON_3" Or key_c = "ECON_4" Or key_c = "ECON_DH" Then
                Set ECON = New clsECON_
                Set dict(skey) = ECON
                
            ElseIf key_c = "DESH" Then
                Set DESH = New clsDESH_
                Set dict(skey) = DESH
                
            ElseIf key_c = "MIX" Or key_c = "MIX_RH" Or key_c = "MIX_1" Or key_c = "MIX_2" Then
                Set MIX = New clsMIX_
                Set dict(skey) = MIX
                
            ElseIf key_c = "SPL" Or key_c = "SPL_RH" Or key_c = "SPL_1" Or key_c = "SPL_2" Or key_c = "SPL_3" Then
                Set SPL = New clsSPL_
                Set dict(skey) = SPL
                
            ElseIf key_c = "MIX_gas" Then
                Set MIX_gas = New clsMIX_gas_
                Set dict(skey) = MIX_gas
                
            ElseIf key_c = "BFP" Or key_c = "COP" Then
                Set PUMP = New clsPUMP_
                Set dict(skey) = PUMP
                
            ElseIf key_c = "DA" Then
                Set DA = New clsDA_
                Set dict(skey) = DA
                
            ElseIf key_c = "CPH" Or key_c = "CPH_1" Or key_c = "CPH_2" Then
                Set CPH = New clsCPH_
                Set dict(skey) = CPH
            
            ElseIf key_c = "GLD" Then
                Set GLD = New clsGLD_
                Set dict(skey) = GLD
            
            End If
        Else
        
            If key_c = "DB" Then
                Set DB = New clsDB_
                dict.Add skey, DB
                
            ElseIf key_c = "SCR" Then 'DuctBurner
                Set SCR = New clsSCR_
                dict.Add skey, SCR
                
            ElseIf key_c = "SPHT" Or key_c = "SPHT_RH" Or key_c = "SPHT_1" Or key_c = "SPHT_2" Or key_c = "SPHT_3" Or key_c = "HRH" Or key_c = "CRH" Or key_c = "CRH_1" Or key_c = "CRH_2" Then
                Set SPHT = New clsSPHT_
                dict.Add skey, SPHT
                
            ElseIf key_c = "EVAP" Or key_c = "EVAP_BD" Then
                Set EVAP = New clsEVAP_
                dict.Add skey, EVAP
                
            ElseIf key_c = "ECON" Or key_c = "ECON_1" Or key_c = "ECON_2" Or key_c = "ECON_3" Or key_c = "ECON_4" Or key_c = "ECON_DH" Then
                Set ECON = New clsECON_
                dict.Add skey, ECON
                
            ElseIf key_c = "DESH" Then
                Set DESH = New clsDESH_
                dict.Add skey, DESH
                
            ElseIf key_c = "MIX" Or key_c = "MIX_RH" Or key_c = "MIX_1" Or key_c = "MIX_2" Then
                Set MIX = New clsMIX_
                dict.Add skey, MIX
                
            ElseIf key_c = "SPL" Or key_c = "SPL_RH" Or key_c = "SPL_1" Or key_c = "SPL_2" Or key_c = "SPL_3" Then
                Set SPL = New clsSPL_
                dict.Add skey, SPL
                
            ElseIf key_c = "MIX_gas" Then
                Set MIX_gas = New clsMIX_gas_
                dict.Add skey, MIX_gas
                
            ElseIf key_c = "BFP" Or key_c = "COP" Then
                Set PUMP = New clsPUMP_
                dict.Add skey, PUMP
                
            ElseIf key_c = "DA" Then
                Set DA = New clsDA_
                dict.Add skey, DA
            
            ElseIf key_c = "CPH" Or key_c = "CPH_1" Or key_c = "CPH_2" Then
                Set CPH = New clsCPH_
                dict.Add skey, CPH
                
            ElseIf key_c = "GLD" Then
                Set GLD = New clsGLD_
                dict.Add skey, GLD
                
            End If
        End If
        
    Next i
 
    index_Para_HX_1 = HRSG_input.Cells(50, "Z") & "_" & HRSG_input.Cells(50, "AA")
    index_Para_HX_2 = HRSG_input.Cells(51, "Z") & "_" & HRSG_input.Cells(51, "AA")
    Gas_ratio_1 = HRSG_input.Cells(50, "AB")
    Gas_ratio_2 = HRSG_input.Cells(51, "AB")
    
    If Not IsEmpty(index_Para_HX_1) Then
        For i = 1 To Number
            If index(i) = index_Para_HX_1 Then
                i_Para_HX_1 = i                     'Component number for Split exchanger
                Exit For
            End If
        Next

        For i = i_Para_HX_1 - 1 To 1 Step -1
        If index_c(i) = "SPHT" Or index_c(i) = "SPHT_RH" Or index_c(i) = "SPHT_1" Or index_c(i) = "SPHT_2" Or index_c(i) = "SPHT_3" Or index_c(i) = "HRH" Or index_c(i) = "CRH" Or index_c(i) = "CRH_1" Or index_c(i) = "CRH_2" Or index_c(i) = "EVAP" Or index_c(i) = "EVAP_BD" Or index_c(i) = "ECON" Or index_c(i) = "ECON_1" Or index_c(i) = "ECON_2" Or index_c(i) = "ECON_3" Or index_c(i) = "ECON_4" Or index_c(i) = "ECON_DH" Or index_c(i) = "CPH" Then
            i_Gas_SPL = i   'Component number for gas split source
            Exit For
        End If
    Next
    End If
    
    DA_type = HRSG_input.Cells(9, "AA")
    P_HP_DSN = HRSG_input.Cells(4, "AA")
    P_IP_DSN = HRSG_input.Cells(5, "AA")
    P_LP_DSN = HRSG_input.Cells(6, "AA")
    EVAP_Cir_ratio = HRSG_input.Cells(4, "AD")
    EVAP_Head = HRSG_input.Cells(5, "AD")
    DA_Head = HRSG_input.Cells(6, "AD")
    
    For n_HP_end = Number To 1 Step -1 'Last Component Number with HP
    If index_p(n_HP_end) = "HP" Then Exit For
    Next n_HP_end
    
    For n_IP_end = Number To 1 Step -1 'Last Component Number with IP
    If index_p(n_IP_end) = "IP" Then Exit For
    Next n_IP_end
    
    For n_LP_end = Number To 1 Step -1 'Last Component Number with LP
    If index_p(n_LP_end) = "LP" Then Exit For
    Next n_LP_end
    
    For n_None_end = Number To 1 Step -1 'Last Component Number with LP
    If index_p(n_None_end) = "None" Then Exit For
    Next n_None_end
    
    For e_EVAP_BD = Number To 1 Step -1 'EVAP_BD
    If index_c(e_EVAP_BD) = "EVAP_BD" Then Exit For
    Next e_EVAP_BD
    
'    For n_DA = number To 1 Step -1 'n_DA
'    If index_c(n_DA) = "DA" Then Exit For
'    Next n_DA
    
    For i = 1 To Number 'Number of Heat Exchanger
    If index_c(i) = "SPHT" Or index_c(i) = "SPHT_RH" Or index_c(i) = "SPHT_1" Or index_c(i) = "SPHT_2" Or index_c(i) = "SPHT_3" Or index_c(i) = "HRH" Or index_c(i) = "CRH" Or index_c(i) = "CRH_1" Or index_c(i) = "CRH_2" Or index_c(i) = "EVAP" Or index_c(i) = "EVAP_BD" Or index_c(i) = "ECON" Or index_c(i) = "ECON_1" Or index_c(i) = "ECON_2" Or index_c(i) = "ECON_3" Or index_c(i) = "ECON_4" Or index_c(i) = "ECON_DH" Or index_c(i) = "CPH" Then
        n_HX = n_HX + 1
    End If
    Next i
    
    For n_HP_start = 1 To Number            'Starting component number with HP
    If index_p(n_HP_start) = "HP" Then
        Exit For
    End If
    Next n_HP_start
    
    For n_IP_start = 1 To Number            'Starting component number with IP
    If index_p(n_IP_start) = "IP" Then
        Exit For
    End If
    Next n_IP_start
    
    For n_LP_start = 1 To Number            'Starting component number with LP
    If index_p(n_LP_start) = "LP" Then
        Exit For
    End If
    Next n_LP_start
    
    Method_P_loss_hot = HRSG_input.Cells(31, "W")
    Method_P_loss_valve = HRSG_input.Cells(32, "W")
    Method_P_loss_cold = HRSG_input.Cells(33, "W")
    BFP_Cal_option = HRSG_input.Cells(3, "AD")          'BFP Calculation Option
    
    'Drum 압력 구하기''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If DA_type = "Integral DA" And HRSG_input.Cells(18, "AA") = "LP" Then 'EVAP, DA componenet / DA의 pegging STM에 LP_EVAP만 있는 경우
        
        For e_LP = 1 To Number
        If index(e_LP) = "LP_EVAP" Then Exit For
        Next e_LP
        P_STM_Level = P_LP_DSN
        
        If Method_P_loss_cold = "P_loss_cold" Then
            For i = 1 To e_LP - 1
                If index_p(i) = "LP" Then P_STM_Level = P_STM_Level / ((1 - P_loss_cold(i)) * (1 - P_loss_valve(i)))
            Next i
            
        ElseIf Method_P_loss_cold = "delta_P_cold" Then
            For i = 1 To e_LP - 1
                If index_p(i) = "LP" Then P_STM_Level = P_STM_Level + (P_loss_cold(i) + P_loss_valve(i)) / 100
            Next i
        End If
        
        HRSG_input.Cells(10, "AA").value = P_STM_Level 'Design DA/LPB Pressure Writing
        HRSG_input_off.Cells(10, "Q").value = P_STM_Level 'Off Design DA/LPB Initial Pressure Value
        
    ElseIf Not e_EVAP_BD = 0 Then
    
        P_STM_Level = P_LP_DSN 'bar
        
        If Method_P_loss_cold = "P_loss_cold" Then
            
            For i = 1 To e_EVAP_BD - 1
                If index_p(i) = "LP" Then P_STM_Level = P_STM_Level / ((1 - P_loss_cold(i)) * (1 - P_loss_valve(i)))
            Next i
            
        ElseIf Method_P_loss_cold = "delta_P_cold" Then
        
            For i = 1 To e_EVAP_BD - 1
                If index_p(i) = "LP" Then P_STM_Level = P_STM_Level + (P_loss_cold(i) + P_loss_valve(i)) / 100
            Next i
        End If
        
        HRSG_input.Cells(21, "AD").value = P_STM_Level 'Design DA/LPB Pressure Writing
        HRSG_input_off.Cells(21, "T").value = P_STM_Level 'Off Design DA/LPB Initial Pressure Value
        
    End If
    

    
End Function

Function read_data_off() As Dictionary

    Dim HRSG_input_off As Worksheet
    Dim dict As New Dictionary
    Dim rg As Range
    Dim i As Integer
    Dim skey As String
    Dim key_p As String
    Dim key_c As String
    Dim key As Variant
    
    Set HRSG_input_off = ThisWorkbook.Worksheets("HRSG_input_off")
    
    Set rg = HRSG_input_off.Range(HRSG_input_off.Range("A2"), HRSG_input_off.Range("A2").End(xlDown))

    Number = rg.Rows.Count

    ReDim index(Number)
    ReDim index_c(Number)
    ReDim index_p(Number)

    ReDim E_loss(Number)
    ReDim Calc_option(Number)
    ReDim Input_value(Number)
    
    ReDim index_S(Number)
    ReDim Stream_S(Number)
    
    ReDim Calc_option_T(Number)
    ReDim Input_value_T(Number)
    ReDim index_T(Number)
    ReDim Stream_T(Number)
    
    For i = 1 To Number
    
        skey = rg(i, 1) & "_" & rg(i, 2)
        key_p = rg(i, 1)
        key_c = rg(i, 2)
        
        index(i) = skey
        index_c(i) = key_c
        index_p(i) = key_p
        
        E_loss(i) = rg(i, 3)
        Calc_option(i) = rg(i, 4)
        Input_value(i) = rg(i, 5)
        
        index_S(i) = rg(i, 6)
        Stream_S(i) = rg(i, 7)
        
        Calc_option_T(i) = rg(i, 8)
        Input_value_T(i) = rg(i, 9)
        index_T(i) = rg(i, 10)
        Stream_T(i) = rg(i, 11)
        
        If Calc_option(i) = "Maximum steam temperaure" And index_p(i) = "HP" Then
            i_HP_DESH = i
        ElseIf Calc_option(i) = "Maximum steam temperaure" And index_p(i) = "IP" Then
            i_IP_DESH = i
        End If
        
            If key_c = "DB" Then
                Set DB = New clsDB_
                dict.Add skey, DB
                
            ElseIf key_c = "SCR" Then
                Set SCR = New clsSCR_
                dict.Add skey, SCR
                
            ElseIf key_c = "SPHT" Or key_c = "SPHT_RH" Or key_c = "SPHT_1" Or key_c = "SPHT_2" Or key_c = "SPHT_3" Or key_c = "HRH" Or key_c = "CRH" Or key_c = "CRH_1" Or key_c = "CRH_2" Then
                Set SPHT = New clsSPHT_
                dict.Add skey, SPHT
                
            ElseIf key_c = "EVAP" Or key_c = "EVAP_BD" Then
                Set EVAP = New clsEVAP_
                dict.Add skey, EVAP
                
            ElseIf key_c = "ECON" Or key_c = "ECON_1" Or key_c = "ECON_2" Or key_c = "ECON_3" Or key_c = "ECON_4" Or key_c = "ECON_DH" Then
                Set ECON = New clsECON_
                dict.Add skey, ECON
                
            ElseIf key_c = "DESH" Then
                Set DESH = New clsDESH_
                dict.Add skey, DESH
                
            ElseIf key_c = "MIX" Or key_c = "MIX_RH" Or key_c = "MIX_1" Or key_c = "MIX_2" Then
                Set MIX = New clsMIX_
                dict.Add skey, MIX
                
            ElseIf key_c = "SPL" Or key_c = "SPL_RH" Or key_c = "SPL_1" Or key_c = "SPL_2" Or key_c = "SPL_3" Then
                Set SPL = New clsSPL_
                dict.Add skey, SPL
            
            ElseIf key_c = "MIX_gas" Then
                Set MIX_gas = New clsMIX_gas_
                dict.Add skey, MIX_gas
            
            ElseIf key_c = "BFP" Or key_c = "COP" Then
                Set PUMP = New clsPUMP_
                dict.Add skey, PUMP
                
            ElseIf key_c = "DA" Then
                Set DA = New clsDA_
                dict.Add skey, DA
                
            ElseIf key_c = "CPH" Or key_c = "CPH_1" Or key_c = "CPH_2" Then
                Set CPH = New clsCPH_
                dict.Add skey, CPH
                
            ElseIf key_c = "GLD" Then
                Set GLD = New clsGLD_
                dict.Add skey, GLD
                
            End If

    Next i
    
    DA_type = HRSG_input_off.Cells(9, "Q")
    P_HP_off = HRSG_input_off.Cells(4, "Q")
    P_IP_off = HRSG_input_off.Cells(5, "Q")
    P_LP_off = HRSG_input_off.Cells(6, "Q")
    
    If i_HP_DESH = 0 And i_IP_DESH = 0 Then
        n_DESH = 0 'Unkown 없음 -> solver 돌지 않음
    ElseIf i_HP_DESH = 0 Or i_IP_DESH = 0 Then
        n_DESH = 1 'Unkown 1 개
    Else
        n_DESH = 2 'Unkown 2 개
    End If
    
    If Not n_DESH = 0 Then
        ReDim Mass_sec(n_DESH - 1)
    
        If n_DESH = 1 Then 'Unkown 1 개
            If i_IP_DESH = 0 Then 'HP_DESH의 Calc_option이 "Maximum steam temperaure"
                Mass_sec(0) = HRSG_input_off.Cells(5, "T")
            ElseIf i_HP_DESH = 0 Then 'IP_DESH의 Calc_option이 "Maximum steam temperaure"
                Mass_sec(0) = HRSG_input_off.Cells(6, "T")
            End If
        ElseIf n_DESH = 2 Then 'Unkown 2 개/ HP_DESH와 IP_DESH의 Calc_option이 "Maximum steam temperaure"
            Mass_sec(0) = HRSG_input_off.Cells(5, "T")
            Mass_sec(1) = HRSG_input_off.Cells(6, "T")
        End If
    End If
    
End Function







