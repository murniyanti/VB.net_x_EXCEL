Attribute VB_Name = "Module1"
Sub Multiple_Server()
Attribute Multiple_Server.VB_ProcData.VB_Invoke_Func = "j\n14"

Dim i As Integer
Dim n As Integer
Dim m As Integer

'Cw is Costumer Waiting, used to calculate Average of costumer who wait
Dim Cw As Integer
'St is simulation time, used to calculate percentage time server busy
Dim St As Integer



'
' for customer 1, must be fixed, manually count
'
'
'

Cw = 0
St = 0
n = 0

'for customer 1
Cells(3, 2) = 0
Cells(3, 3) = 0
Cells(3, 4) = 0
Cells(3, 5) = "=RandBetween(1, 100)"
If Cells(3, 5) <= 50 Then Cells(3, 6) = 1 Else Cells(3, 6) = 2

Cells(3, 7) = "=randbetween(1,100)"

'time start service if Cashier 1
If Cells(3, 6) = 1 Then
    Cells(3, 8) = Cells(3, 4)
If Cells(3, 7) <= 50 Then
    Cells(3, 9) = 2
    Else
        If Cells(3, 7) >= 51 And Cells(3, 7) <= 85 Then
            Cells(3, 9) = 3
            Else
                If Cells(3, 7) >= 86 And Cells(3, 7) <= 100 Then
                    Cells(3, 9) = 4
                    End If
            End If
         End If

'Cells(3, 8) = Cells(3, 4).Value

'time service end
Cells(3, 10) = Cells(3, 8).Value + Cells(3, 9).Value
'waiting time
Cells(3, 14) = 0

n = Cells(3, 10).Value
    Cells(2, 18).Value = n
End If

'Time start service if Cashier 2
If Cells(3, 6) = 2 Then
    Cells(3, 11) = Cells(3, 4).Value
If Cells(3, 7) <= 35 Then
    Cells(3, 12) = 3
    Else
        If Cells(3, 7) >= 36 And Cells(3, 7) <= 80 Then
            Cells(3, 12) = 4
            Else
                If Cells(3, 7) >= 81 And Cells(3, 7) <= 100 Then
                    Cells(3, 12) = 5
                    End If
            End If
         End If

'Cells(3, 8) = Cells(3, 4).Value
'time service end
Cells(3, 13) = Cells(3, 11).Value + Cells(3, 12).Value
'waiting time
Cells(3, 14) = 0
m = Cells(3, 13).Value
    Cells(2, 19).Value = m
End If




'
'
'For Customer 2 and the rest
'
'
'
'

For i = 4 To 24

Cells(i, 2).Select
ActiveCell.FormulaR1C1 = "=RANDBETWEEN(1,100)"

'Random number for Interarrival time

If Cells(i, 2) <= 25 Then
        Cells(i, 3) = 1

Else
    If Cells(i, 2) >= 26 And Cells(i, 2) <= 50 Then
        Cells(i, 3) = 2
    Else
        If Cells(i, 2) >= 51 And Cells(i, 2) <= 75 Then
            Cells(i, 3) = 3
        Else
            If Cells(i, 2) >= 76 Then
                Cells(i, 3) = 4
            End If
        End If
    End If
End If

'Arrival time of customer
Cells(i, 4).Select
ActiveCell.FormulaR1C1 = "=R[-1]C+RC[-1]"


If Cells(i - 1, 10).Value = Empty And Cells(i - 1, 13).Value = Empty Then
    Cells(i, 5) = "=RANDBETWEEN(1,100)"
    'Else
     '   If n = m And Cells(i, 4) < n Then
      '  Cells(i, 5) = "=Randbetween(1,100)"
        'Else
        'If n = m And Cells(i, 4) = m Then
        'Cells(i, 5) = "=Randbetween(1,100)"
            Else
            If Cells(i, 4) >= n And Cells(i, 4) >= m Then
            Cells(i, 5) = "=Randbetween(1,100)"
            'Else
            'If Cells(i, 4) <= n And m Then
            'Cells(i, 5) = "=randbetween(1,100)"
        'End If
        'End If
        End If
        End If
        
        

'
'
'To determine which Cashier the Costumer get.
'
'Cells(i, 6).Select
'
If Cells(i, 5) >= 0 And Cells(i, 5) <= 50 Then
    Cells(i, 6) = "1"
    Else
        If Cells(i, 5) >= 51 And Cells(i, 5) <= 100 Then
            Cells(i, 6) = "2"
            Else
                If Cells(i, 5).Value = Empty Then
                    Cells(i, 6).Value = Empty
                    
                    
                End If
                    
        End If
    End If
    
' If cells Cashier empty, we need to compare Cashier 1 and 2 service end time

If Cells(i, 5).Value = Empty And n < m Then
Cells(i, 6) = 1
Else
    If Cells(i, 5).Value = Empty And m < n Then
    Cells(i, 6) = 2
    
    
    End If
End If
    


    
' to assign rand# for service time
'if cashier_1
'   rand 1-50= 2
'   rand 51-85=3
'   rand 86-100=4
'
'if cashier_2
'   rand 1-35=3
'   rand 36-80=4
'   rand 81-100=5

Cells(i, 7).Select
ActiveCell.FormulaR1C1 = "=RANDBETWEEN(1,100)"

'
'if Cashier_1
'
'time service start

If Cells(i, 6) = 1 Then
    Cells(i, 8).Select
    If Cells(i, 4).Value > n Then ActiveCell.FormulaR1C1 = Cells(i, 4).Value Else ActiveCell.FormulaR1C1 = n
    'waiting time
    Cells(i, 14).Select
    If Cells(i, 4) < n Then Cells(i, 14) = Cells(2, 18).Value - Cells(i, 4).Value Else Cells(i, 14) = 0
    If Cells(i, 14).Value = 0 Then Cw = Cw + 1
    
    Cells(i, 9).Select
    'To determine the service time
    If Cells(i, 7) >= 1 And Cells(i, 7) <= 50 Then
    Cells(i, 9) = 2
    Else
        If Cells(i, 7) >= 51 And Cells(i, 7) <= 85 Then
            Cells(i, 9) = 3
            Else
                If Cells(i, 7) >= 86 And Cells(i, 7) <= 100 Then
                    Cells(i, 9) = 4
                    End If
            End If
         End If
                    
    'Time service ends
    Cells(i, 10) = Cells(i, 8) + Cells(i, 9)
    Cells(i, 16).Value = Cells(i, 10).Value
    n = Cells(i, 10).Value
    Cells(2, 18).Value = n
    
    
End If
'
'
'Cashier 2
'
'
'if Cashier_2
'
'time service start

If Cells(i, 6) = 2 Then
    Cells(i, 11).Select
    If Cells(i, 4).Value > m Then ActiveCell.FormulaR1C1 = Cells(i, 4).Value Else ActiveCell.FormulaR1C1 = m
    'waiting time
    Cells(i, 14).Select
    If Cells(i, 4) < m Then Cells(i, 14) = Cells(2, 19).Value - Cells(i, 4).Value Else Cells(i, 14) = 0
    If Cells(i, 14).Value = 0 Then Cw = Cw + 1
    
    Cells(i, 12).Select
    'To determine the service time
    If Cells(i, 7) >= 1 And Cells(i, 7) <= 35 Then
        Cells(i, 12) = 3
    Else
        If Cells(i, 7) >= 36 And Cells(i, 7) <= 80 Then
            Cells(i, 12) = 4
            Else
                If Cells(i, 7) >= 81 And Cells(i, 7) <= 100 Then
                    Cells(i, 12) = 5
                    End If
            End If
         End If
                    
    'Time service ends
    Cells(i, 13) = Cells(i, 11) + Cells(i, 12)
    Cells(i, 17).Value = Cells(i, 13).Value
    m = Cells(i, 13).Value
    Cells(2, 19).Value = m
    
    'Cells(i, 8) = ""
    'Cells(i, 9) = ""
    'Cells(i, 10) = ""
    
End If

If Cells(i, 10).Value < Cells(i, 13).Value Then St = Cells(i, 13).Value Else St = Cells(i, 10).Value


Next i

'calculation, probability etc.
'
'Total Service Time Cashier 1
Cells(25, 9).Select
ActiveCell.FormulaR1C1 = "=sum(R[-22]C:R[-1]C)"

'Total Service Time Cashier 2
Cells(25, 12).Select
ActiveCell.FormulaR1C1 = "=sum(R[-22]C:R[-1]C)"

'Total Waiting Time
Cells(25, 14).Select
ActiveCell.FormulaR1C1 = "=sum(R[-22]C:R[-1]C)"

'Percentage of time server busy- each server
Cells(29, 2) = (Cells(25, 9).Value / St) * 100
Cells(30, 2) = (Cells(25, 12).Value / St) * 100

'Average Waiting time for Every Customer
Cells(33, 2) = Cells(25, 14).Value / 22

'Probability that a customer has to wait
If Cw = 0 Then Cells(36, 2).Value = 0 Else Cells(36, 2) = (21 - Cw) / 22

'Average Waiting time for Customer who wait
Cells(39, 2) = Cells(25, 14).Value / (21 - Cw)

End Sub

Sub Single_Server()

Dim x As Integer
'For 1st Customer
Cells(3, 2) = "=randbetween(1,100)"
If Cells(3, 2) <= 30 Then
Cells(3, 3) = 3
Else
    If Cells(3, 2) >= 31 And Cells(3, 2) <= 50 Then
    Cells(3, 3) = 5
    Else
        If Cells(3, 2) >= 51 And Cells(3, 2) <= 80 Then
        Cells(3, 3) = 7
        Else
            If Cells(3, 2) >= 81 Then
            Cells(3, 3) = 9
            End If
        End If
    End If
End If

'Arrival time
Cells(3, 4) = Cells(3, 3)

'rand# for service time
Cells(3, 5) = "=RandBetween(1, 100)"

'service time
If Cells(3, 5) <= 40 Then
Cells(3, 6) = 4
Else
    If Cells(3, 5) >= 41 And Cells(3, 5) <= 60 Then
    Cells(3, 6) = 6
    Else
        If Cells(3, 5) >= 61 And Cells(3, 5) <= 80 Then
        Cells(3, 6) = 8
        Else
            If Cells(3, 5) >= 81 Then
            Cells(3, 6) = 10
            End If
        End If
      End If
End If

'service start time
Cells(3, 7) = 0
'Cells(3, 4) + Cells(3, 5)
'waiting time
Cells(3, 8) = 0
'time service ends
Cells(3, 9) = Cells(3, 6)
'(3, 7) + Cells(3, 6)

'Time costummer spends in system
Cells(3, 10) = Cells(3, 8) + Cells(3, 6)

'Server idle time of server= waiting for customer
Cells(3, 11) = 0
'
'
'For Costumer 2 and so on
'
'


For x = 4 To 9

Cells(x, 2) = "=randbetween(1,100)"
If Cells(x, 2) <= 30 Then
Cells(x, 3) = 3
Else
    If Cells(x, 2) >= 31 And Cells(x, 2) <= 50 Then
    Cells(x, 3) = 5
    Else
        If Cells(x, 2) >= 51 And Cells(x, 2) <= 80 Then
        Cells(x, 3) = 7
        Else
            If Cells(x, 2) >= 81 Then
            Cells(x, 3) = 9
            End If
        End If
    End If
End If

'Arrival time
Cells(x, 4) = Cells(x, 3)

'rand# for service time
Cells(x, 5) = "=RandBetween(1, 100)"

'service time
If Cells(x, 5) <= 40 Then
Cells(x, 6) = 4
Else
    If Cells(x, 5) >= 41 And Cells(x, 5) <= 60 Then
    Cells(x, 6) = 6
    Else
        If Cells(x, 5) >= 61 And Cells(x, 5) <= 80 Then
        Cells(x, 6) = 8
        Else
            If Cells(x, 5) >= 81 Then
            Cells(x, 6) = 10
            End If
        End If
      End If
End If

'service start time
If Cells(x, 4).Value < Cells(x - 1, 9).Value Then Cells(x, 7) = Cells(x, 4).Value Else Cells(x, 7) = Cells(x - 1, 9)


'waiting time
If Cells(x, 4) > Cells(x - 1, 9) Then Cells(x, 8) = 0 Else Cells(x, 8) = Cells(x - 1, 9) - Cells(x, 4)

'time service ends
Cells(x, 9) = Cells(x, 6) + Cells(x, 7)

'Time costummer spends in system
Cells(x, 10) = Cells(x, 8) + Cells(x, 6)

'Server idle time of server= waiting for customer
If Cells(x - 1, 9) < Cells(x, 4) Then Cells(x, 11) = Cells(x, 4) - Cells(x - 1, 9) Else Cells(x, 11) = 0

Next x
'total
Cells(12, 6).Select
ActiveCell.FormulaR1C1 = "=sum(R[-9]C:R[-3]C)"

'TOTAL IDLE TIME OF SERVER
Cells(12, 11).Select
ActiveCell.FormulaR1C1 = "=sum(R[-9]C:R[-3]C)"

End Sub


