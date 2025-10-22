VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSolverHRSG_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 Option Explicit

Public HRSG As New clsHRSG_
Public HRSG_read As New clsHRSG_read_

Function NR_solver_HRSG(mode As Integer) As Variant
    
Select Case mode
Case 0
    
    Set HRSG = New clsHRSG_
    Set HRSG_read = New clsHRSG_read_
    
    Dim n, i, j, k, iter, q, r As Integer
    Dim x() As Variant
    Dim ea, delta As Double
    Dim Y_old() As Variant
    Dim Y_new() As Variant
    Dim delta_X() As Variant
    Dim U() As Variant
    Dim del_X() As Variant
    Dim Y_matrix() As Variant
    Dim Jacobi_matrix() As Variant
    Dim Inverse() As Variant
    Dim X_p As Double
    Dim X_new() As Variant
    Dim err() As Variant
    Dim max_err As Double
    Dim Y() As Variant
    Dim y_max As Variant
    
    
    HRSG.calHRSG (0)
    HRSG_read.read_data
    

Case 1

    Dim HRSG_input_off As Worksheet
    Set HRSG_read = New clsHRSG_read_
    
    HRSG_read.read_data_off
    
    If HRSG_read.n_DESH = 0 Then 'Unkown이 없는 경우 solver를 돌지 않음
        HRSG.calHRSG (1)
        HRSG_read.read_data_off
    Else

        Set HRSG_input_off = ThisWorkbook.Worksheets("HRSG_input_off")

        n = HRSG_read.n_DESH
        iter = 1
        ea = 1
'        y_max = 1
        delta = 0.0001

        ReDim x(n - 1)
        ReDim Y_old(n - 1)
        ReDim Y_new(n - 1)
        ReDim delta_X(n - 1, n - 1)
        ReDim U(n - 1, n - 1)
        ReDim del_X(n - 1)
        ReDim Y_matrix(n - 1, n - 1)
        ReDim Jacobi_matrix(n - 1, n - 1)
        ReDim Inverse(n - 1, n - 1)
        ReDim X_new(n - 1)
        ReDim err(n - 1)

        'Identity matrix 생성
        For i = 0 To n - 1
            For j = 0 To n - 1
                If i = j Then
                    U(i, j) = 1
                Else
                    U(i, j) = 0
                End If
            Next
        Next

        'Do~Loop 문''''''''''''''''''''''''''''''
'        Do While y_max >= 0.01
        Do While ea >= 0.0001
        
        If iter = 1 Then
            For i = 0 To n - 1
                x(i) = HRSG_read.Mass_sec(i)
            Next
            Y_old = HRSG.calHRSG(mode)
        End If

        If iter >= 2 Then
            Y_old = HRSG.calHRSG(mode)
        End If

        For i = 0 To n - 1
            For j = 0 To n - 1
                delta_X(i, j) = x(j) - U(i, j) * delta '초기값에서 delta를 빼준 매트릭스 생성
            Next
        Next

        'X에서 delta을 뺀 X를 이용하여 새로운 Y값을 구하여 매트릭스 형태로 만듦
        For i = 0 To n - 1
    
            For j = 0 To n - 1
                del_X(j) = delta_X(i, j)
            Next

            If n = 1 Then
                If HRSG_read.i_IP_DESH = 0 Then
                    HRSG_input_off.Cells(5, "T").value = del_X(0)
                ElseIf HRSG_read.i_HP_DESH = 0 Then
                    HRSG_input_off.Cells(6, "T").value = del_X(0)
                End If
            ElseIf n = 2 Then
                HRSG_input_off.Cells(5, "T").value = del_X(0)
                HRSG_input_off.Cells(6, "T").value = del_X(1)
            End If

            Y_new = HRSG.calHRSG(mode)
            
            For k = 0 To n - 1
                Y_matrix(k, i) = Y_old(k) - Y_new(k) 'Y_old - Y_new 값은 delta Y이다. delta Y를 매트릭스형태로 저장함
            Next
        Next

        For i = 0 To n - 1
            For j = 0 To n - 1
                Jacobi_matrix(i, j) = Y_matrix(i, j) / delta ' delta Y / delta X 형태의 매트릭스를 만들어 Jacobi 매트릭스를 만듦
            Next
        Next

        '역행렬 구함
        Inverse = Inverse_matrix_HRSG(Jacobi_matrix, mode)

        'Jacobi 매트릭스를 이용하여 새로운 X값 구함
        For i = 0 To n - 1
            X_p = 0
            For j = 0 To n - 1
                X_p = X_p + Inverse(i, j) * Y_old(j)
            Next
            X_new(i) = x(i) - X_p
        Next

        'X_new이 허수면 반복문 멈추고 X_new값을 전 Step의 X값인 X_old로 바꿈
        For i = 0 To n - 1
            If Not Application.WorksheetFunction.Imaginary(X_new(i)) = 0 Then
                X_new(i) = x(i)
                Exit Do
            End If
        Next

        For i = 0 To n - 1
            err(i) = Abs(X_new(i) - x(i)) '초기 가정값X-새로 구한 X로 차이를 구함
            x(i) = X_new(i) 'new 값을 기존값(결과)에 대입
            If err(i) > max_err Then
                max_err = err(i)
            End If
        Next

        ea = max_err
'         y_max = Application.WorksheetFunction.Max(Abs(Y_old(n - 3)), Abs(Y_old(n - 2)), Abs(Y_old(n - 1)))
        max_err = 0
        iter = iter + 1

        If n = 1 Then
            If HRSG_read.i_IP_DESH = 0 Then
                HRSG_input_off.Cells(5, "T").value = x(0)
            ElseIf HRSG_read.i_HP_DESH = 0 Then
                HRSG_input_off.Cells(6, "T").value = x(0)
            End If
        ElseIf n = 2 Then
            HRSG_input_off.Cells(5, "T").value = x(0)
            HRSG_input_off.Cells(6, "T").value = x(1)
        End If
        Loop
        '''''''''''''''''''''''''''''''''''''''''''''''
    End If
    
End Select
End Function

