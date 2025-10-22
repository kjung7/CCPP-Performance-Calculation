VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsResultPrint_HRSG_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Result As New clsSolverHRSG_

Public col As Integer


Sub ResultPrint(mode)

    Dim i, k As Integer
    
    If mode = 0 Then
    
    worksheet_name = "HRSG_Output"
    Worksheet_name_off = "HRSG_input_off"
    
    Worksheets(worksheet_name).Range("B3:L71").ClearContents
    Worksheets(worksheet_name).Range("Q4:S8").ClearContents
    Worksheets(worksheet_name).Range("Q11:S11").ClearContents
    
    For i = 1 To Result.HRSG_read.Number
    
    Worksheets(worksheet_name).Cells(i + 2, "B") = Result.HRSG_read.index(i)
    Worksheets(worksheet_name).Cells(i + 2, "C") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).Outlet_h.Massflow.m, 4)
    Worksheets(worksheet_name).Cells(i + 2, "D") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).Inlet_c.Massflow.m, 4)
    Worksheets(worksheet_name).Cells(i + 2, "E") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).Outlet_c.Massflow.m, 4)
    
    Worksheets(worksheet_name).Cells(i + 2, "F") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).Inlet_h.Properties.T, 4)
    Worksheets(worksheet_name).Cells(i + 2, "G") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).Outlet_h.Properties.T, 4)

    Worksheets(worksheet_name).Cells(i + 2, "H") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).Inlet_c.Properties.T, 4)
    Worksheets(worksheet_name).Cells(i + 2, "I") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).Outlet_c.Properties.T, 4)
    
    If Result.HRSG_read.index_c(i) = "BFP" Or Result.HRSG_read.index_c(i) = "COP" Then
        Worksheets(worksheet_name).Cells(i + 2, "J") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).Efficiency, 4)
    Else
        Worksheets(worksheet_name).Cells(i + 2, "J") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).Effectiveness, 4)
    End If
        
    Worksheets(worksheet_name).Cells(i + 2, "K") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).Inlet_c.Properties.P, 4) / 100
    Worksheets(worksheet_name).Cells(i + 2, "L") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).Outlet_c.Properties.P, 4) / 100
   
   Worksheets(worksheet_name).Cells(i + 2, "M") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).Velocity, 4)
    Worksheets(worksheet_name).Cells(i + 2, "N") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).UA, 4)
    
    
    
    
    If Result.HRSG_read.index(i) = "None_DA" Then
        Worksheets(worksheet_name).Cells(4, "Q") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).BFWIN.Massflow.m, 4)
        Worksheets(worksheet_name).Cells(4, "R") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).BFWIN.Properties.h, 4)
        Worksheets(worksheet_name).Cells(4, "S") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).BFWIN.Properties.P, 4)
        Worksheets(worksheet_name).Cells(5, "Q") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).MNSTM.Massflow.m, 4)
        Worksheets(worksheet_name).Cells(5, "R") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).MNSTM.Properties.h, 4)
        Worksheets(worksheet_name).Cells(5, "S") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).MNSTM.Properties.P, 4)
        Worksheets(worksheet_name).Cells(6, "Q") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).AUXSTM.Massflow.m, 4)
        Worksheets(worksheet_name).Cells(6, "R") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).AUXSTM.Properties.h, 4)
        Worksheets(worksheet_name).Cells(6, "S") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).AUXSTM.Properties.P, 4)
        Worksheets(worksheet_name).Cells(7, "Q") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).MNBFW.Massflow.m, 4)
        Worksheets(worksheet_name).Cells(7, "R") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).MNBFW.Properties.h, 4)
        Worksheets(worksheet_name).Cells(7, "S") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).MNBFW.Properties.P, 4)
        Worksheets(worksheet_name).Cells(8, "Q") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).AUXBFW.Massflow.m, 4)
        Worksheets(worksheet_name).Cells(8, "R") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).AUXBFW.Properties.h, 4)
        Worksheets(worksheet_name).Cells(8, "S") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).AUXBFW.Properties.P, 4)
    ElseIf Result.HRSG_read.index(i) = "None_CPH_1" Or Result.HRSG_read.index(i) = "None_CPH" Then
        Worksheets(worksheet_name).Cells(10, "P") = Result.HRSG_read.index(i)
        Worksheets(worksheet_name).Cells(11, "Q") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).Cir.Massflow.m, 4)
        Worksheets(worksheet_name).Cells(11, "R") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).Cir.Properties.T, 4)
        Worksheets(worksheet_name).Cells(11, "S") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).Cir.Properties.P, 4)
    End If
    
    If IsEmpty(Worksheets(Worksheet_name_off).Cells(5, "T")) Then 'cell이 비어있는 경우만 design의 값을 sheet에 입력
        If Result.HRSG_read.index(i) = "HP_DESH" Then Worksheets(Worksheet_name_off).Cells(5, "T") = Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).Sec.Massflow.m
    End If
    
    If IsEmpty(Worksheets(Worksheet_name_off).Cells(6, "T")) Then
        If Result.HRSG_read.index(i) = "IP_DESH" Then Worksheets(Worksheet_name_off).Cells(6, "T") = Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).Sec.Massflow.m
    End If
    
    
     'HP STM Outlet Conditions
    
    Worksheets("HRSG_input").Cells(57, "AA").value = Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(Result.HRSG.HRSG_call.n_HP_start)).Outlet_c.Properties.P / 100
    Worksheets("HRSG_input").Cells(58, "AA").value = Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(Result.HRSG.HRSG_call.n_HP_start)).Outlet_c.Properties.T
    Worksheets("HRSG_input").Cells(59, "AA").value = Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(Result.HRSG.HRSG_call.n_HP_start)).Outlet_c.Massflow.m * 3.6
    
    'IP STM Outlet Conditions
    Worksheets("HRSG_input").Cells(57, "AC").value = Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(Result.HRSG.HRSG_call.n_IP_start)).Outlet_c.Properties.P / 100
    Worksheets("HRSG_input").Cells(58, "AC").value = Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(Result.HRSG.HRSG_call.n_IP_start)).Outlet_c.Properties.T
    Worksheets("HRSG_input").Cells(59, "AC").value = Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(Result.HRSG.HRSG_call.n_IP_start)).Outlet_c.Massflow.m * 3.6
    
    
    Next
    
    ElseIf mode = 1 Then
    
    worksheet_name = "HRSG_output_off"
        Worksheets(worksheet_name).Range("B3:L71").ClearContents
            Worksheets(worksheet_name).Range("Q4:S8").ClearContents
    Worksheets(worksheet_name).Range("Q11:S11").ClearContents
        
    For i = 1 To Result.HRSG_read.Number
    

    
    Worksheets(worksheet_name).Cells(i + 2, "B") = Result.HRSG_read.index(i)
    
    Worksheets(worksheet_name).Cells(i + 2, "C") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).Outlet_h.Massflow.m, 4)
    Worksheets(worksheet_name).Cells(i + 2, "D") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).Inlet_c.Massflow.m, 4)
    Worksheets(worksheet_name).Cells(i + 2, "E") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).Outlet_c.Massflow.m, 4)
    
    Worksheets(worksheet_name).Cells(i + 2, "F") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).Inlet_h.Properties.T, 4)
    Worksheets(worksheet_name).Cells(i + 2, "G") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).Outlet_h.Properties.T, 4)
    
    Worksheets(worksheet_name).Cells(i + 2, "H") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).Inlet_c.Properties.T, 4)
    Worksheets(worksheet_name).Cells(i + 2, "I") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).Outlet_c.Properties.T, 4)

    If Result.HRSG_read.index_c(i) = "BFP" Or Result.HRSG_read.index_c(i) = "COP" Then
        Worksheets(worksheet_name).Cells(i + 2, "J") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).PUMP_Eff_Off, 4)
    Else
        Worksheets(worksheet_name).Cells(i + 2, "J") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).Effectiveness, 4)
    End If
    
   
    Worksheets(worksheet_name).Cells(i + 2, "K") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).Inlet_c.Properties.P, 4) / 100
    Worksheets(worksheet_name).Cells(i + 2, "L") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).Outlet_c.Properties.P, 4) / 100
    
           Worksheets(worksheet_name).Cells(i + 2, "M") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).Velocity, 4)
    Worksheets(worksheet_name).Cells(i + 2, "N") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).UA, 4)
     
    
    
    If Result.HRSG_read.index(i) = "None_DA" Then
        Worksheets(worksheet_name).Cells(4, "q") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).BFWIN.Massflow.m, 4)
        Worksheets(worksheet_name).Cells(4, "r") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).BFWIN.Properties.h, 4)
        Worksheets(worksheet_name).Cells(4, "s") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).BFWIN.Properties.P, 4)
        Worksheets(worksheet_name).Cells(5, "Q") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).MNSTM.Massflow.m, 4)
        Worksheets(worksheet_name).Cells(5, "R") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).MNSTM.Properties.h, 4)
        Worksheets(worksheet_name).Cells(5, "S") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).MNSTM.Properties.P, 4)
        Worksheets(worksheet_name).Cells(6, "Q") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).AUXSTM.Massflow.m, 4)
        Worksheets(worksheet_name).Cells(6, "R") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).AUXSTM.Properties.h, 4)
        Worksheets(worksheet_name).Cells(6, "S") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).AUXSTM.Properties.P, 4)
        Worksheets(worksheet_name).Cells(7, "Q") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).MNBFW.Massflow.m, 4)
        Worksheets(worksheet_name).Cells(7, "R") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).MNBFW.Properties.h, 4)
        Worksheets(worksheet_name).Cells(7, "S") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).MNBFW.Properties.P, 4)
        Worksheets(worksheet_name).Cells(8, "Q") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).AUXBFW.Massflow.m, 4)
        Worksheets(worksheet_name).Cells(8, "R") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).AUXBFW.Properties.h, 4)
        Worksheets(worksheet_name).Cells(8, "S") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).AUXBFW.Properties.P, 4)
    ElseIf Result.HRSG_read.index(i) = "None_CPH_1" Or Result.HRSG_read.index(i) = "None_CPH" Then
        Worksheets(worksheet_name).Cells(10, "P") = Result.HRSG_read.index(i)
        Worksheets(worksheet_name).Cells(11, "Q") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).Cir.Massflow.m, 4)
        Worksheets(worksheet_name).Cells(11, "R") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).Cir.Properties.T, 4)
        Worksheets(worksheet_name).Cells(11, "S") = Round(Result.HRSG.HRSG_call.dict(Result.HRSG_read.index(i)).Cir.Properties.P, 4)
    End If

    
    Next
    
    
    End If
    
 
End Sub
    
    







