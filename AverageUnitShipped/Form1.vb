Option Strict On
Public Class Form1

    Const minimumShipped As Integer = 0
    Const maximumShipped As Integer = 1000

    Dim day As Integer = 1
    Dim empId As Integer = 0

    Dim units(2, 6) As Integer


    Sub ResetForm()
        ' reset variables
        Dim Control As Control
        Dim rt As TextBox
        For Each Control In Me.Controls
            If TypeOf (Control) Is TextBox Then
                rt = CType(Control, TextBox)
                rt.Clear()
            End If

        Next Control
        tbAvg.Text = ""
        day = 1
        lblday.Text = "Day " + day.ToString
        empId = 0
        btnEnter.Enabled = True
        tbInput.ReadOnly = False
        LblEmp1.Font = New Font(LblEmp1.Font, FontStyle.Bold)
        LblEmp2.Font = New Font(LblEmp2.Font, FontStyle.Regular)
        LblEmp3.Font = New Font(LblEmp3.Font, FontStyle.Regular)

    End Sub

    Function arrayAverage(ByVal arrayToAverage As Integer(,)) As String
        Dim Sum As Double = 0
        Dim Avg As Double = 0
        For ColumnIndex As Integer = 0 To units.GetUpperBound(1)
            Sum = 0
            For RowIndex As Integer = 0 To units.GetUpperBound(0)
                Sum += units(RowIndex, ColumnIndex)
            Next
            Avg = Avg + (Sum / 7)
        Next
        Return Math.Round(Avg / 3, 2).ToString
    End Function
    Function DayInc(ByVal days As Integer) As String
        Dim average As Double = 0
        If (day < 7) Then
            day = day + 1

            lblday.Text = "Day " + day.ToString

        ElseIf day = 7 Then
            average = arrayAverage(units, empId)


            empId = empId + 1
            day = 1
            lblday.Text = "Day " + day.ToString
            Return "Average: " + average.ToString




        End If
        Return String.Empty
    End Function


    Function arrayAverage(ByVal arrayToAverage As Integer(,), emp As Integer) As Double
        Dim Sum As Double = 0
        For ColumnIndex As Integer = 0 To units.GetUpperBound(1)
            Sum += units(emp, ColumnIndex)
        Next

        Return Math.Round(Sum / 7, 2)
    End Function

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        Call ResetForm()
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Application.Exit()
    End Sub

    Private Sub btnEnter_Click(sender As Object, e As EventArgs) Handles btnEnter.Click
        Dim UnitShipped As Integer
        Dim userInput As String = tbInput.Text
        Dim formattedOutput As String = ""


        If Not Integer.TryParse(userInput, UnitShipped) Then
            MessageBox.Show("Please ensure you enter a number for units Shipped!")

            tbInput.SelectionStart = 0
            tbInput.SelectionLength = tbInput.Text.Length

        ElseIf (UnitShipped < minimumShipped OrElse UnitShipped > maximumShipped) Then

            MessageBox.Show("Please ensure the value must be between 0 to 1000!")

        Else
            units(empId, day - 1) = Convert.ToInt32(userInput)

            For counter As Integer = 1 To day
                formattedOutput += units(empId, counter - 1).ToString + vbCrLf
                tbInput.Text = ""

            Next

            If empId = 0 Then

                tbstore1.Text = formattedOutput
                formattedOutput = String.Empty
                tbAvgEmp1.Text = DayInc(day)
            ElseIf empId = 1 Then
                tbstore2.Text = formattedOutput
                formattedOutput = String.Empty
                tbAvgEmp2.Text = DayInc(day)
            Else

                tbstore3.Text = formattedOutput
                formattedOutput = String.Empty

                tbAvgEmp3.Text = DayInc(day)

            End If

        End If

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LblEmp1.Font = New Font(LblEmp1.Font, FontStyle.Bold)
    End Sub

    Private Sub tbAvgEmp1_TextChanged(sender As Object, e As EventArgs) Handles tbAvgEmp1.TextChanged
        LblEmp1.Font = New Font(LblEmp1.Font, FontStyle.Regular)
        LblEmp2.Font = New Font(LblEmp2.Font, FontStyle.Bold)

    End Sub

    Private Sub tbAvgEmp2_TextChanged(sender As Object, e As EventArgs) Handles tbAvgEmp2.TextChanged
        LblEmp2.Font = New Font(LblEmp2.Font, FontStyle.Regular)
        LblEmp3.Font = New Font(LblEmp3.Font, FontStyle.Bold)
    End Sub

    Private Sub tbAvgEmp3_TextChanged(sender As Object, e As EventArgs) Handles tbAvgEmp3.TextChanged
        btnEnter.Enabled = False
        tbInput.ReadOnly = True
        tbAvg.Text = "Average per day= " + arrayAverage(units)
        lblday.Text = "Done"
        ' tbAvg.Text = (Math.Round((Convert.ToDouble(tbAvgEmp1.Text) + Convert.ToDouble(tbAvgEmp2.Text) + Convert.ToDouble(tbAvgEmp3.Text)) / 3, 2)).ToString
    End Sub

    Private Sub tbstore1_TextChanged(sender As Object, e As EventArgs) Handles tbstore1.TextChanged

    End Sub

    Private Sub lblunits_Click(sender As Object, e As EventArgs) Handles lblunits.Click

    End Sub

    Private Sub tbInput_TextChanged(sender As Object, e As EventArgs) Handles tbInput.TextChanged

    End Sub

    Private Sub lblday_Click(sender As Object, e As EventArgs) Handles lblday.Click

    End Sub

    Private Sub LblEmp1_Click(sender As Object, e As EventArgs) Handles LblEmp1.Click

    End Sub

    Private Sub LblEmp2_Click(sender As Object, e As EventArgs) Handles LblEmp2.Click

    End Sub

    Private Sub tbstore2_TextChanged(sender As Object, e As EventArgs) Handles tbstore2.TextChanged

    End Sub

    Private Sub LblEmp3_Click(sender As Object, e As EventArgs) Handles LblEmp3.Click

    End Sub

    Private Sub tbstore3_TextChanged(sender As Object, e As EventArgs) Handles tbstore3.TextChanged

    End Sub

    Private Sub tbAvg_TextChanged(sender As Object, e As EventArgs) Handles tbAvg.TextChanged

    End Sub
End Class
