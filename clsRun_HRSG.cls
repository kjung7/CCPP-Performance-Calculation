VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRun_HRSG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public HRSG As New clsHRSG_
Public HRSG_result As New clsResultPrint_HRSG_
Public Solver As New clsSolverHRSG_

Sub HRSGExe()

    Dim k As Long
    k = Timer
    
        'Screen value definition
    screenUpdateState = Application.ScreenUpdating
    statusBarState = Application.DisplayStatusBar
    calcState = Application.Calculation
    eventsState = Application.EnableEvents

    'Turning off screen update setting
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    
    Set HRSG = New clsHRSG_
    Set HRSG_result = New clsResultPrint_HRSG_
    Set Solver = New clsSolverHRSG_

    Solver.NR_solver_HRSG (0)
    With HRSG_result
        Set .Result = Solver
    End With
    HRSG_result.ResultPrint (0)
    MsgBox "Design 계산 완료" & Chr(13) & "경과 시간 : " & Format(Abs(Timer - k) & " sec")

    Solver.NR_solver_HRSG (1)
    With HRSG_result
        Set .Result = Solver
    End With
    HRSG_result.ResultPrint (1)
    MsgBox "Off-design 계산 완료" & Chr(13) & "경과 시간 : " & Format(Abs(Timer - k) & " sec")
    
    
        'Recovering screen update setting
    Application.ScreenUpdating = screenUpdateState
    Application.DisplayStatusBar = statusBarState
    Application.Calculation = calcState
    Application.EnableEvents = eventsState

End Sub




