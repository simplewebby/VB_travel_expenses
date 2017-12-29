'Tsagan Garyaeva
'Comp-185
'Lab 7 Travel Expenses



Public Class Form1
    'Constants for Reimbursement Rates/Prices    
    Const decMEAL_REIMBURSEMENT As Decimal = 37D
    Const decPARKING_REIMBURSEMENT As Decimal = 10D
    Const decTAXI_REIMBURSEMENT As Decimal = 20D
    Const decLODGING_REIMBURSEMENT As Decimal = 95D
    Const decMILES_REIMBURSEMENT As Decimal = 0.27D


    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub btnCalc_Click(sender As Object, e As EventArgs) Handles btnCalc.Click

        InputEmpty()
        'Begin nested if statement by first checking numeric validation     
        If InputNumeric() Then
            If InputPositive() Then

            End If
            'check functions for positive math and otherwise   
            ' display nothing in the display labels for negative math      
            'usin conversion of str data types and then the tostring method to   
            'format currency in displayed fields          
            If CalcTotalReimbursement() > 0 Then lblTotalAllow.Text = CStr(CalcTotalReimbursement().ToString("c"))
        Else lblTotalAllow.Text = String.Empty
        End If
        If CalcTotal() > 0 Then lblTotalExp.Text = CStr(CalcTotal().ToString("c")) Else lblTotalExp.Text = String.Empty

        If CalcUnallowed() > 0 Then lblExcess.Text = CStr(CalcUnallowed().ToString("c")) Else lblExcess.Text = String.Empty

        If CalcSaved() > 0 Then lblSvedAmount.Text = CStr(CalcSaved().ToString("c")) Else lblSvedAmount.Text = String.Empty
        MessageBox.Show("Enter a Positive Valid Amount")
        MessageBox.Show("Enter a Numeric Amount")

    End Sub




    Function CalcLodging() As Decimal
        Dim decLodgingReimbursement As Decimal
        decLodgingReimbursement = (CDec(txtNumOfDays.Text) * decLODGING_REIMBURSEMENT)
        Return decLodgingReimbursement
    End Function
    Function CalcMeals() As Decimal
        Dim decMealReimbursement As Decimal
        decMealReimbursement = (CDec(txtNumOfDays.Text) * decMEAL_REIMBURSEMENT)
        Return decMealReimbursement
    End Function
    Function CalcMileage() As Decimal
        Dim decMileageReimbursement As Decimal
        decMileageReimbursement = (CDec(txtMilesDriven.Text) * decMILES_REIMBURSEMENT)
        Return decMileageReimbursement
    End Function
    Function CalcParkingFees() As Decimal
        Dim decParkingReimbursement As Decimal
        decParkingReimbursement = CDec(txtNumOfDays.Text) * decPARKING_REIMBURSEMENT
        Return decParkingReimbursement
    End Function
    Function CalcTaxiFees() As Decimal
        Dim decTaxiReimbursement As Decimal
        decTaxiReimbursement = decTAXI_REIMBURSEMENT * CDec(txtNumOfDays.Text)
        Return decTaxiReimbursement
    End Function
    'Function Total Reimbursement is adding the previous functions from above 
    Function CalcTotalReimbursement() As Decimal
        Dim decTotalReimbursement As Decimal
        decTotalReimbursement = CalcLodging() + CalcTaxiFees() + CalcMeals() + CalcParkingFees()
        Return decTotalReimbursement
    End Function
    Function CalcUnallowed() As Decimal
        Dim decUnallowed As Decimal
        decUnallowed = ((CDec(txtParkFee.Text)) - (CalcParkingFees())) +
            (CDec(txtTaxi.Text) - CalcTaxiFees()) + (CDec(txtLodge.Text) - CalcLodging()) + (CDec(txtMeal.Text) - CalcMeals())
        Return decUnallowed
    End Function
    Function CalcSaved() As Decimal
        Dim decSaved As Decimal
        decSaved = (CDec(txtNumOfDays.Text) * decMEAL_REIMBURSEMENT - CDec(txtMeal.Text)) +
            (CDec(txtNumOfDays.Text) * decPARKING_REIMBURSEMENT - CDec(txtParkFee.Text)) +
            (CDec(txtNumOfDays.Text) * decTAXI_REIMBURSEMENT - CDec(txtTaxi.Text)) +
              (CDec(txtNumOfDays.Text) * decLODGING_REIMBURSEMENT - CDec(txtLodge.Text))
        Return decSaved
    End Function
    'Calculation Total was added as a function    
    Function CalcTotal() As Decimal
        Dim decTotal As Decimal
        decTotal = CDec(txtAir.Text) + CDec(txtRegFee.Text) + CDec(txtMeal.Text) + (CDec(txtMilesDriven.Text) * decMILES_REIMBURSEMENT) + CDec(txtCarRental.Text) + CDec(txtLodge.Text) + CDec(txtParkFee.Text) + CDec(txtTaxi.Text)
        Return decTotal
    End Function
    'Function for Numeric Validation, returns True or false for all keyed fields   
    Function InputNumeric() As Boolean
        Dim blnNumeric As Boolean
        If IsNumeric(txtAir.Text) And IsNumeric(txtCarRental.Text) And IsNumeric(txtNumOfDays.Text) And IsNumeric(txtLodge.Text) And IsNumeric(txtMeal.Text) And IsNumeric(txtMilesDriven.Text) And IsNumeric(txtParkFee.Text) And IsNumeric(txtRegFee.Text) And IsNumeric(txtTaxi.Text) Then blnNumeric = True

        Return blnNumeric
    End Function
    'Similar to the Numeric Validation except checks all fields for Positive numbers   
    Function InputPositive() As Boolean
        Dim blnPositive As Boolean
        If CDbl(txtAir.Text) >= 0 And
            CDbl(txtCarRental.Text) >= 0 And
            CDbl(txtNumOfDays.Text) >= 0 And
            CDbl(txtLodge.Text) >= 0 And
            CDbl(txtMeal.Text) >= 0 And
            CDbl(txtMilesDriven.Text) >= 0 And
            CDbl(txtParkFee.Text) >= 0 And
            CDbl(txtRegFee.Text) >= 0 And
            CDbl(txtCarRental.Text) >= 0 Then blnPositive = True

        Return blnPositive
    End Function


    Sub InputEmpty()
        If txtAir.Text = String.Empty Then txtAir.Text = "0"

        If txtCarRental.Text = String.Empty Then txtCarRental.Text = "0"

        If txtNumOfDays.Text = String.Empty Then txtNumOfDays.Text = "0"

        If txtLodge.Text = String.Empty Then txtLodge.Text = "0"

        If txtMeal.Text = String.Empty Then txtMeal.Text = "0"

        If txtMilesDriven.Text = String.Empty Then txtMilesDriven.Text = "0"

        If txtParkFee.Text = String.Empty Then txtParkFee.Text = "0"

        If txtRegFee.Text = String.Empty Then txtRegFee.Text = "0"

        If txtTaxi.Text = String.Empty Then txtTaxi.Text = "0"

        If lblExcess.Text = String.Empty Then lblExcess.Text = "0"

        If lblSvedAmount.Text = String.Empty Then lblSvedAmount.Text = "0"

        If lblTotalExp.Text = String.Empty Then lblTotalExp.Text = "0"

        If lblTotalAllow.Text = String.Empty Then lblTotalAllow.Text = "0"

    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        txtAir.Text = String.Empty
        txtCarRental.Text = String.Empty
        txtNumOfDays.Text = String.Empty
        txtLodge.Text = String.Empty
        txtMeal.Text = String.Empty
        txtMilesDriven.Text = String.Empty
        txtParkFee.Text = String.Empty
        txtRegFee.Text = String.Empty
        txtTaxi.Text = String.Empty
        lblTotalExp.Text = String.Empty
        lblTotalAllow.Text = String.Empty
        lblExcess.Text = String.Empty
        lblSvedAmount.Text = String.Empty


    End Sub

    Private Sub Label10_Click(sender As Object, e As EventArgs) Handles Label10.Click

    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
End Class
