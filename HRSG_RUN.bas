Attribute VB_Name = "HRSG_RUN"
Option Explicit

Public Run As New clsRun_HRSG

'Public Solver As New clsSolver_
'Public Read As New clsRead_

Sub HRSGExe()
    
    Set Run = New clsRun_HRSG
   
    
    
    
    Run.HRSGExe
    
End Sub
