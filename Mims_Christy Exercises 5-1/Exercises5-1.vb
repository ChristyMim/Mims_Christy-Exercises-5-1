'Program: Chapter 5 Exercise 1 
'Programmer : Christy Mims
'Date: 03/08/2017
'Description: This program allows the user to view various scenarios for Chapter 5 exercise 1 problems 
Public Class Form1
    Dim amount As Decimal = 0
    Private Sub btnCalculateBMI_1_Click(sender As Object, e As EventArgs) Handles btnCalculateBMI_1.Click
        Dim weight, height As Double
        weight = CDbl(txtweight.Text)
        height = CDbl(txtHeight.Text)
        txtBMI.Text = BMI(weight, height)
    End Sub
    Function BMI(w As Double, h As Double) As Double
        Return Math.Round((703 * w) / h ^ 2)
    End Function


    Private Sub btnTrainingHeartRate_Click(sender As Object, e As EventArgs) Handles btnTrainingHeartRate.Click
        Dim Age As Integer
        Dim rest_rate As Double
        Age = CInt(txtAge.Text)
        rest_rate = CDbl(txtRestingHeartRate.Text)
        txtTHR.Text = CalculateTHR(Age, rest_rate)
    End Sub
    Function CalculateTHR(A As Integer, R As Double) As Double
        Return ((220 - A) - R) * 0.6D + R
    End Function


    Private Sub btnDetermineSemesterGrade_Click(sender As Object, e As EventArgs) Handles btnDetermineSemesterGrade.Click
        Dim midterm, final_exam As Double
        Dim Averaged_Amount As String
        midterm = CDbl(txtMidtermExam.Text)
        final_exam = CDbl(txtFinalExam.Text)
        Averaged_Amount = CalculateFinal(midterm, final_exam)
        txtSemesterGrade.Text = Averaged_Amount
    End Sub
    Function CalculateFinal(midterm_grade As Decimal, final_test As Decimal) As String
        CalculateFinal = Math.Round((final_test * 2 + midterm_grade) / 2)
        CalculateFinal = Ceil(CalculateFinal)

        Dim Averaged_Amount As String
        If CalculateFinal >= 100 And CalculateFinal >= 90 Then
            Averaged_Amount = "A"
        End If
        If CalculateFinal <= 89 And CalculateFinal >= 80 Then
            Averaged_Amount = "B"
        End If
        If CalculateFinal <= 79 And CalculateFinal >= 70 Then
            Averaged_Amount = "C"
        End If
        If CalculateFinal <= 69 And CalculateFinal >= 60 Then
            Averaged_Amount = "D"
        End If
        If CalculateFinal <= 59 Then
            Averaged_Amount = "F"
        End If
        Return Averaged_Amount
    End Function
    Function Ceil(num As Double) As Integer
        Return -Int(-num)
    End Function

    Private Sub btnDetermineAverage_Click(sender As Object, e As EventArgs) Handles btnDetermineAverage.Click
        Dim Average_1, Average_2, Average_3 As Integer
        Average_1 = (txtAverageValue1.Text)
        Average_2 = (txtAverageValue2.Text)
        Average_3 = (txtAverageValue3.Text)
        txtAverage.Text = Math.Round(Average(Average_1, Average_2, Average_3), 2)
    End Sub
    Function Average(Average_1 As Integer, Average_2 As Integer, Average_3 As Integer) As Double
        Return (Average_1 + Average_2 + Average_3) / 3
    End Function


    Private Sub btnDetermineMinimum_Click(sender As Object, e As EventArgs) Handles btnDetermineMinimum.Click
        Dim MinValue_1, MinValue_2, MinValue_3 As Double
        MinValue_1 = CDbl(txtMinimumValue1.Text)
        MinValue_2 = CDbl(txtMinimumValue2.Text)
        MinValue_3 = CDbl(txtMinimumValue3.Text)
        txtMinimum.Text = CStr(Minimum(MinValue_1, MinValue_2, MinValue_3))
    End Sub
    Function Minimum(Min_1 As Double, Min_2 As Double, Min_3 As Double) As Double
        Dim total_Min As Double
        total_Min = Min_1
        If Min_2 < total_Min Then
            total_Min = Min_2
        End If
        If Min_3 < total_Min Then
            total_Min = Min_3
        End If
        Return total_Min
    End Function

    Private Sub btnDisplayMonthlyBalance_Click(sender As Object, e As EventArgs) Handles btnDisplayMonthlyBalance.Click
        Dim Month_1, Month_2, Month_3, Month_4 As Decimal
        amount = CDec(txtDeposit.Text)
        Month_1 = amount_1(Month_1, Month_2, Month_3, Month_4)
        Month_2 = amount_2(Month_1, Month_2, Month_3, Month_4)
        Month_3 = amount_3(Month_1, Month_2, Month_3, Month_4)
        Month_4 = amount_4(Month_1, Month_2, Month_3, Month_4)

        lstResults.Items.Add("Month 1: " & Month_1.ToString("C"))
        lstResults.Items.Add("Month 2: " & Month_2.ToString("C"))
        lstResults.Items.Add("Month 3: " & Month_3.ToString("C"))
        lstResults.Items.Add("Month 4: " & Month_4.ToString("C"))
    End Sub
    Function amount_1(First_month As Decimal, second_month As Decimal, third_month As Decimal, fourth_month As Decimal) As Decimal
        First_month = 1.005D * 0 + amount
        Return First_month
    End Function
    Function amount_2(First_month As Decimal, second_month As Decimal, third_month As Decimal, fourth_month As Decimal) As Decimal
        second_month = 1.005D * First_month + amount
        Return second_month
    End Function
    Function amount_3(First_month As Decimal, second_month As Decimal, third_month As Decimal, fourth_month As Decimal) As Decimal
        third_month = 1.005D * second_month + amount
        Return third_month
    End Function
    Function amount_4(First_month As Decimal, second_month As Decimal, third_month As Decimal, fourth_month As Decimal) As Decimal
        fourth_month = 1.005D * third_month + amount
        Return fourth_month
    End Function

    Private Sub btnCalculateCost_Click(sender As Object, e As EventArgs) Handles btnCalculateCost.Click
        Dim residents, days As Double
        days = txtNumberofDays.Text
        txtCost.Text = (CalculateCostPerDay(days, residents)).ToString("C")
    End Sub
    Function CalculateCostPerDay(numDays As Integer, resident As Boolean) As Decimal
        If radYes.Checked Then
            If numDays <= 6 Then
                Return 65 * numDays
            ElseIf numDays > 7 Then
                Return ((numDays - 6) * 50) + 65 * 6
            End If
        End If
        If radNo.Checked Then
            If numDays > 6 Then
                Return (75 * 6) + ((numDays - 6) * 50)
            End If
            If numDays <= 6 Then
                Return 75 * numDays
            End If
        End If
    End Function
End Class
