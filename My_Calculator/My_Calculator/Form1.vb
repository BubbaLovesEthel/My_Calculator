﻿
Public Class Form1
    Dim Operand_1 As Decimal
    Dim Operand_2 As Decimal
    Dim Memory As Decimal
    Dim Calc_Type As String
    Dim Missing_Op_1 As Boolean
    Dim Missing_Op_2 As Boolean
    Dim Start_New_Number As Boolean
    Dim Enabled_Operands = True

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox_Display.Text = "0"
        KeyPreview = True
        Missing_Op_1 = True
        Missing_Op_2 = True
        Start_New_Number = True
        Me.Show()
        Button_Equals.Focus()
    End Sub

    Private Sub TextBox_Display_TextChanged(sender As Object, e As EventArgs) Handles TextBox_Display.TextChanged
        If (Missing_Op_1) Then
            TextBox_OP1_Building.Text = TextBox_Display.Text
        ElseIf (Missing_Op_2) Then
            TextBox_OP2_Building.Text = TextBox_Display.Text
        ElseIf (Not Missing_Op_1 And Start_New_Number) Then
            TextBox_OP1_Building.Text = TextBox_Display.Text
        End If
    End Sub

    Private Sub Button_Number_Click(sender As Object, e As EventArgs) Handles Button_Zero.Click, Button_One.Click, Button_Two.Click, Button_Three.Click, Button_Four.Click, Button_Five.Click, Button_Six.Click, Button_Seven.Click, Button_Eight.Click, Button_Nine.Click
        If (TextBox_Display.Text.Equals("0") Or Start_New_Number Or (Not IsNumeric(TextBox_Display.Text) Or (TextBox_Display.Text = "NaN"))) Then
            TextBox_Display.Text = Microsoft.VisualBasic.Strings.Right((CType(sender, Button).Text), 1)
            If (Start_New_Number And Not Missing_Op_1 And Not Missing_Op_2) Then
                Missing_Op_1 = True
            End If
            Start_New_Number = False
        Else
            TextBox_Display.Text = TextBox_Display.Text + Microsoft.VisualBasic.Strings.Right((CType(sender, Button).Text), 1)
        End If
        If (Not Enabled_Operands) Then
            Enabled_Operands = True
            Enable_Operands()
        End If
        Button_Equals.Focus()
    End Sub

    Private Sub Button_Decimal_Click(sender As Object, e As EventArgs) Handles Button_Decimal.Click
        If (TextBox_Display.Text.IndexOf(".") = -1) Then
            TextBox_Display.Text = TextBox_Display.Text + "."
        End If
        Button_Equals.Focus()
    End Sub

    Private Sub Button_Negation_Click(sender As Object, e As EventArgs) Handles Button_Negation.Click
        If Not ((Not IsNumeric(TextBox_Display.Text) Or (TextBox_Display.Text = "NaN"))) Then
            If (TextBox_Display.Text.StartsWith("-")) Then
                TextBox_Display.Text = TextBox_Display.Text.Remove(0, 1)
            Else
                If Not (TextBox_Display.Text.Equals("0")) Then
                    TextBox_Display.Text = "-" + TextBox_Display.Text
                End If
            End If
        End If
        Button_Equals.Focus()
    End Sub

    Private Sub Button_Erase_Click(sender As Object, e As EventArgs) Handles Button_Erase.Click
        If TextBox_Display.Text.Length > 0 Then
            TextBox_Display.Text = TextBox_Display.Text.Substring(0, TextBox_Display.Text.Length - 1)
        End If
        If TextBox_Display.Text.Length = 0 Or TextBox_Display.Text.Equals("-") Or TextBox_Display.Text = "-0" Or (Not IsNumeric(TextBox_Display.Text) Or (TextBox_Display.Text = "NaN")) Then
            TextBox_Display.Text = "0"
        End If
        If (Not Enabled_Operands) Then
            Enabled_Operands = True
            Enable_Operands()
        End If
        Button_Equals.Focus()
    End Sub

    Private Sub Button_Binary_Calculation_Click(sender As Object, e As EventArgs) Handles Button_Addition.Click, Button_Division.Click, Button_Multiplication.Click, Button_Subtraction.Click
        TextBox_Calculation_Type_Building.Text = (CType(sender, Button).Text)
        If Missing_Op_2 And Not Start_New_Number And Not Missing_Op_1 Then
            If Not ((Not IsNumeric(TextBox_Display.Text) Or (TextBox_Display.Text = "NaN"))) Then
                Operand_2 = Val(TextBox_Display.Text)
                '                TextBox_OP2_Building.Text = Operand_2.ToString ' CHECK THIS
                TextBox_OP2.Text = Operand_2.ToString
                TextBox_Calculation_Type.Text = Calc_Type
                TextBox_OP1.Text = Operand_1.ToString
                Missing_Op_2 = False
                Binary_Calculation()
                Calc_Type = (CType(sender, Button).Text)
            End If
        End If
        If Not ((Not IsNumeric(TextBox_Display.Text) Or (TextBox_Display.Text = "NaN"))) Then
            Operand_1 = Val(TextBox_Display.Text)
            TextBox_OP1_Building.Text = Operand_1.ToString
            Start_New_Number = True
            Calc_Type = (CType(sender, Button).Text)
            Missing_Op_1 = False
            TextBox_OP2_Building.Text = ""
            Missing_Op_2 = True
        End If
        Button_Equals.Focus()
    End Sub

    Private Sub Button_Equals_Click(sender As Object, e As EventArgs) Handles Button_Equals.Click
        If (Not IsNumeric(TextBox_Display.Text) Or (TextBox_Display.Text = "NaN")) Then
            TextBox_Display.Text = "0"
            TextBox_OP1.Text = ""
            TextBox_Calculation_Type.Text = ""
            TextBox_OP2.Text = ""
            If (Not Enabled_Operands) Then
                Enabled_Operands = True
                Enable_Operands()
            End If
            Missing_Op_2 = True
            Missing_Op_1 = True
            Return
        End If
        If (Missing_Op_1) Then
            Operand_1 = Val(TextBox_Display.Text)
            Missing_Op_1 = False
        ElseIf Missing_Op_2 Then
            Operand_2 = Val(TextBox_Display.Text)
            TextBox_OP2.Text = Operand_2.ToString
            Missing_Op_2 = False
        End If
        If Not Missing_Op_1 And Not Missing_Op_2 Then
            Binary_Calculation()
            Start_New_Number = True
            TextBox_OP2.Text = Operand_2.ToString
            TextBox_OP2_Building.Text = Operand_2.ToString
            TextBox_Calculation_Type.Text = Calc_Type
            TextBox_OP1.Text = Operand_1.ToString
            Operand_1 = Val(TextBox_Display.Text)
            Missing_Op_1 = False
        End If
    End Sub

    Private Sub Binary_Calculation()
        Try
            Select Case (Calc_Type)
                Case "+"
                    TextBox_Display.Text = (Operand_1 + Operand_2).ToString
                Case "-"
                    TextBox_Display.Text = (Operand_1 - Operand_2).ToString
                Case "*"
                    TextBox_Display.Text = (Operand_1 * Operand_2).ToString
                Case "/"
                    TextBox_Display.Text = (Operand_1 / Operand_2).ToString
            End Select
        Catch ex As Exception
            TextBox_Display.Text = ex.Message()
            Disable_Operands()
        End Try
        If (Not IsNumeric(TextBox_Display.Text) Or (TextBox_Display.Text = "NaN")) Then
            Disable_Operands()
        End If
    End Sub

    Private Sub Form1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MyBase.KeyPress
        Select Case e.KeyChar
            Case CType("-", Char)
                Button_Subtraction.PerformClick()
            Case CType("+", Char)
                Button_Addition.PerformClick()
            Case CType("*", Char)
                Button_Multiplication.PerformClick()
            Case CType("/", Char)
                Button_Division.PerformClick()
            Case Microsoft.VisualBasic.ChrW(Keys.Return)
                Button_Equals.PerformClick()
            Case Convert.ToChar(Keys.Back)
                Button_Erase.PerformClick()
        End Select
    End Sub

    Private Sub Button_Clear_Click(sender As Object, e As EventArgs) Handles Button_Clear.Click
        TextBox_Display.Text = "0"
        Missing_Op_1 = True
        Missing_Op_2 = True
        Operand_2 = 0
        TextBox_OP2.Text = ""
        TextBox_OP2_Building.Text = ""
        Operand_1 = 0
        TextBox_OP1.Text = ""
        TextBox_OP1_Building.Text = ""
        Calc_Type = ""
        TextBox_Calculation_Type.Text = ""
        TextBox_Calculation_Type_Building.Text = ""
        If (Not Enabled_Operands) Then
            Enabled_Operands = True
            Enable_Operands()
        End If
        Button_Equals.Focus()
    End Sub

    Private Sub Button_Clear_Entry_Click(sender As Object, e As EventArgs) Handles Button_Clear_Entry.Click
        TextBox_Display.Text = "0"
        If (Not Enabled_Operands) Then
            Enabled_Operands = True
            Enable_Operands()
        End If
        Button_Equals.Focus()
    End Sub


    Private Sub Unary_Calculations(sender As Object, e As EventArgs) Handles Button_Square_Root.Click, Button_Squared.Click, Button_One_Over.Click, Button_Percent.Click
        TextBox_OP2.Text = ""
        TextBox_OP1.Text = TextBox_Display.Text
        Try
            Select Case CType(sender, Button).Name
                Case "Button_Square_Root"
                    TextBox_Calculation_Type.Text = "Square Root"
                    TextBox_Display.Text = Math.Sqrt(TextBox_Display.Text).ToString
                Case "Button_Squared"
                    TextBox_Calculation_Type.Text = "Squared"
                    TextBox_Display.Text = (Val(TextBox_Display.Text) * Val(TextBox_Display.Text)).ToString
                Case "Button_One_Over"
                    TextBox_Calculation_Type.Text = "1/X"
                    TextBox_Display.Text = (1 / Val(TextBox_Display.Text)).ToString
                Case "Button_Percent"
                    TextBox_Calculation_Type.Text = "Percentage"
                    TextBox_Display.Text = (Val(TextBox_Display.Text) * 0.01).ToString
            End Select
        Catch ex As Exception
            TextBox_Display.Text = ex.Message()
            Disable_Operands()
        End Try
        If (Not IsNumeric(TextBox_Display.Text) Or (TextBox_Display.Text = "NaN")) Then
            Disable_Operands()
        ElseIf (Not Missing_Op_1 And Start_New_Number) Then
            Operand_1 = Val(TextBox_Display.Text)
        End If
    End Sub

    Private Sub Disable_Operands()
        Enabled_Operands = False
        Button_Division.Enabled = False
        Button_Multiplication.Enabled = False
        Button_Subtraction.Enabled = False
        Button_Addition.Enabled = False
        Button_Square_Root.Enabled = False
        Button_Squared.Enabled = False
        Button_One_Over.Enabled = False
        Button_Percent.Enabled = False
        Button_Negation.Enabled = False
        Button_Decimal.Enabled = False
    End Sub
    Private Sub Enable_Operands()
        Enabled_Operands = True
        Button_Division.Enabled = True
        Button_Multiplication.Enabled = True
        Button_Subtraction.Enabled = True
        Button_Addition.Enabled = True
        Button_Square_Root.Enabled = True
        Button_Squared.Enabled = True
        Button_One_Over.Enabled = True
        Button_Percent.Enabled = True
        Button_Negation.Enabled = True
        Button_Decimal.Enabled = True
    End Sub

    Private Sub Button_MC_Click(sender As Object, e As EventArgs) Handles Button_MC.Click
        Memory = 0
        TextBox_Memory.Text = Memory.ToString
    End Sub

    Private Sub Button_MR_Click(sender As Object, e As EventArgs) Handles Button_MR.Click
        TextBox_Display.Text = Memory.ToString
    End Sub

    Private Sub Button_MS_Click(sender As Object, e As EventArgs) Handles Button_MS.Click
        If Not ((Not IsNumeric(TextBox_Display.Text) Or (TextBox_Display.Text = "NaN"))) Then
            Memory = Val(TextBox_Display.Text)
            TextBox_Memory.Text = Memory.ToString
        End If
    End Sub

    Private Sub Button_M_Plus_Click(sender As Object, e As EventArgs) Handles Button_M_Plus.Click
        If Not ((Not IsNumeric(TextBox_Display.Text) Or (TextBox_Display.Text = "NaN"))) Then
            Memory = Memory + Val(TextBox_Display.Text)
            TextBox_Memory.Text = Memory.ToString
        End If
    End Sub

    Private Sub Button_M_Minus_Click(sender As Object, e As EventArgs) Handles Button_M_Minus.Click
        If Not ((Not IsNumeric(TextBox_Display.Text) Or (TextBox_Display.Text = "NaN"))) Then
            Memory = Memory - Val(TextBox_Display.Text)
            TextBox_Memory.Text = Memory.ToString
        End If
    End Sub

End Class
