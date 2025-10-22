Attribute VB_Name = "Inverse"
Option Explicit

Function Inverse_matrix_HRSG(Matrix As Variant, mode As Integer) As Variant

    Dim n As Integer
    Dim i As Integer, j As Integer, k As Integer, max_index As Integer
    Dim pivot_column As Integer
    'Dim original(10, 10) As Variant, Inverse(10, 10) As Variant
    Dim original() As Variant, Inverse() As Variant
    Dim max_value As Double, ftemp(2) As Double, pivot_value As Double
    

    '������� ����
    For i = 2 To 10 'DESH ��ġ�� 10�̻����� Ŀ���� ���� ������ ����
    
        If Worksheets("HRSG_input_off").Cells(i, "D") = "Maximum steam temperaure" Then
            n = n + 1
        End If
        
        If n > 1 Then Exit For 'n=2 �� ��쿡�� for�� stop ���� (n=1�̸� i=10���� ���ƾ���)
    Next
       
    ReDim original(n - 1, n - 1)
    ReDim Inverse(n - 1, n - 1)
    
    
    '��������� 10���� ũ�� ��� �����ϰ� ����.
    '���� ��������� 10���� ũ���ϰ��� �Ѵٸ� �Ʒ��� ��� ��ºθ� �����Ͽ��� �մϴ�.
    If n > 10 Then
        MsgBox "Max.10"
        Exit Function
    End If
    
    'm�� ���簢����� �ҷ����� m�� ���簢 ��������� ������.
    For i = 0 To n - 1
        For j = 0 To n - 1
            
            'm�� ���簢���
            original(i, j) = Matrix(i, j)
            
            'm�� �������
            If i = j Then
                Inverse(i, j) = 1
            Else
                Inverse(i, j) = 0
            End If
        Next j
    Next i
    
    '����� ���
    For pivot_column = 0 To (n - 1)
    
        max_index = pivot_column '0
        max_value = 0
            
        For i = pivot_column To (n - 1)
        
            If (original(i, pivot_column) ^ 2) > (max_value ^ 2) Then
                max_index = i
                max_value = original(i, pivot_column)
            End If

            If (pivot_column <> max_index) Then
            
                For j = 0 To (n - 1)
                
                    ftemp(0) = original(pivot_column, j)
                    ftemp(1) = Inverse(pivot_column, j)
                    
                    original(pivot_column, j) = original(max_index, j)
                    Inverse(pivot_column, j) = Inverse(max_index, j)
                    
                    original(max_index, j) = ftemp(0)
                    Inverse(max_index, j) = ftemp(1)
                
                Next j

            End If

        Next i
       
        pivot_value = original(pivot_column, pivot_column)

        For j = 0 To (n - 1)
            original(pivot_column, j) = original(pivot_column, j) / pivot_value
        
            Inverse(pivot_column, j) = Inverse(pivot_column, j) / pivot_value
        Next j
       
        For i = 0 To (n - 1)
            
            If (i <> pivot_column) Then
            
                ftemp(0) = original(i, pivot_column)
            
            
                For j = 0 To (n - 1)
                    original(i, j) = original(i, j) - ftemp(0) * original(pivot_column, j)
                    Inverse(i, j) = Inverse(i, j) - ftemp(0) * Inverse(pivot_column, j)
                Next j
            
            End If
            
        Next i
        
    Next pivot_column
    
    Inverse_matrix_HRSG = Inverse

End Function





