Attribute VB_Name = "Module3"
Public temp_string As String

Sub CouponForm(Number_Of_People_In_Party As Integer, total_cost As Variant, Party_Names() As String)

    Dim k As Integer
    Dim o As Integer

    Dim Coupon_User_Form As Object
    
    Dim OkButton As MSForms.CommandButton
    Dim CancelButton As MSForms.CommandButton
    
    Dim Pre_Tax_Total_Label As MSForms.Label
    Dim Number_of_People_in_Party_Label As MSForms.Label
    Dim Coupon_Cost_Label As MSForms.Label
    Dim Coupon_Cost_TextBox As MSForms.TextBox
    Dim Coupon_Redeem_Value_Label As MSForms.Label
    Dim Coupon_Redeem_Value_TextBox As MSForms.TextBox
    Dim Party_Names_Holder_Label As MSForms.Label
    
    Dim Party_Names_Labels_Array() As MSForms.Label
    Dim Tax_Amount_Per_Person_Labels() As MSForms.Label
    Dim Tax_Amount_Per_Person_TextBoxes() As MSForms.TextBox
    Dim Meal_Cost_Per_Person_Labels() As MSForms.Label
    Dim Meal_Cost_Per_Person_TextBoxes() As MSForms.TextBox
    Dim Coupon_Payment_Per_Person_Labels() As MSForms.Label
    Dim Coupon_Payment_Per_Person_TextBoxes() As MSForms.TextBox
    
    Dim Tax_Amount_TextBox_Name_Property() As String
    Dim Meal_Cost_TextBox_Name_Property() As String
    Dim Coupon_Payment_TextBox_Name_Property() As String
    
    Set Coupon_User_Form = ThisWorkbook.VBProject.VBComponents.Add(vbext_ct_MSForm)
    
    With Coupon_User_Form
        .Properties("Caption") = "Payments and Costs of Party Members"
        .Properties("Width") = 789.75
        .Properties("Height") = (156 + ((Number_Of_People_In_Party - 1) * 24)) + 198
    End With
    
    Set Pre_Tax_Total_Label = Coupon_User_Form.Designer.Controls.Add("Forms.Label.1")
        
    With Pre_Tax_Total_Label
        .Name = "Pre_Tax_Total_Label"
        .Caption = "Pre-Tax Total:   $" + CStr(total_cost)
        .Top = 18
        .Left = 18
        .Width = 240
        .Height = 18
        .Font.Size = 12
        .Font.Name = "Times New Roman"
    End With
    
    Set Number_of_People_in_Party_Label = Coupon_User_Form.Designer.Controls.Add("Forms.Label.1")
    
    With Number_of_People_in_Party_Label
        .Name = "Number_of_People_in_Party_Label"
        .Caption = "Number People in Party:   " + CStr(Number_Of_People_In_Party)
        .Top = 42
        .Left = 18
        .Width = 270
        .Height = 18
        .Font.Size = 12
        .Font.Name = "Times New Roman"
    End With
    
    Set Coupon_Cost_Label = Coupon_User_Form.Designer.Controls.Add("Forms.Label.1")
    
    With Coupon_Cost_Label
        .Name = "Coupon_Cost_Label"
        .Caption = "Coupon Cost:   "
        .Top = 96
        .Left = 18
        .Width = 102
        .Height = 18
        .Font.Size = 12
        .Font.Name = "Times New Roman"
    End With
    
    Set Coupon_Cost_TextBox = Coupon_User_Form.Designer.Controls.Add("Forms.TextBox.1")
    
    With Coupon_Cost_TextBox
        .Name = "Coupon_Cost_TextBox"
        .Top = 96
        .Left = 132
        .Width = 90
        .Height = 18
        .Font.Size = 12
        .Font.Name = "Times New Roman"
    End With
    
    Set Coupon_Redeem_Value_Label = Coupon_User_Form.Designer.Controls.Add("Forms.Label.1")
    
    With Coupon_Redeem_Value_Label
        .Name = "Coupon_Redeem_Value_Label"
        .Caption = "Coupon Redeem Value:"
        .Top = 120
        .Left = 18
        .Width = 120
        .Height = 18
        .Font.Size = 12
        .Font.Name = "Times New Roman"
    End With
    
    Set Coupon_Redeem_Value_TextBox = Coupon_User_Form.Designer.Controls.Add("Forms.TextBox.1")
    
    With Coupon_Redeem_Value_TextBox
        .Name = "Coupon_Redeem_Value_TextBox"
        .Top = 120
        .Left = 150
        .Width = 90
        .Height = 18
        .Font.Size = 12
        .Font.Name = "Times New Roman"
    End With
    
    k = 0
    
    ReDim Party_Names_Labels_Array(0 To (Val(Number_Of_People_In_Party)) - 1)
    ReDim Tax_Amount_Per_Person_Labels(0 To (Val(Number_Of_People_In_Party)) - 1)
    ReDim Tax_Amount_Per_Person_TextBoxes(0 To (Val(Number_Of_People_In_Party)) - 1)
    ReDim Meal_Cost_Per_Person_Labels(0 To (Val(Number_Of_People_In_Party)) - 1)
    ReDim Meal_Cost_Per_Person_TextBoxes(0 To (Val(Number_Of_People_In_Party)) - 1)
    ReDim Coupon_Payment_Per_Person_Labels(0 To (Val(Number_Of_People_In_Party)) - 1)
    ReDim Coupon_Payment_Per_Person_TextBoxes(0 To (Val(Number_Of_People_In_Party)) - 1)
    
    ReDim Tax_Amount_TextBox_Name_Property(0 To (Number_Of_People_In_Party - 1))
    ReDim Meal_Cost_TextBox_Name_Property(0 To (Number_Of_People_In_Party - 1))
    ReDim Coupon_Payment_TextBox_Name_Property(0 To (Number_Of_People_In_Party - 1))
    
    For k = 0 To Val(Number_Of_People_In_Party - 1)
    
        Set Party_Names_Labels_Array(k) = Coupon_User_Form.Designer.Controls.Add("Forms.Label.1")
        
        With Party_Names_Labels_Array(k)
            .Name = Party_Names(k)
            .Caption = Party_Names(k) + "--->"
            .Top = 156 + k * 24
            .Left = 18
            .Width = 174
            .Height = 18
            .Font.Size = 12
            .Font.Name = "Times New Roman"
        End With
        
        Set Tax_Amount_Per_Person_Labels(k) = Coupon_User_Form.Designer.Controls.Add("Forms.Label.1")
        
        With Tax_Amount_Per_Person_Labels(k)
            .Name = Party_Names(k) + "_Tax_Amount_Label"
            .Caption = "Meal Tax:"
            .Top = 156 + k * 24
            .Left = 378
            .Width = 60
            .Height = 18
            .Font.Size = 12
            .Font.Name = "Times New Roman"
        End With
        
        Set Tax_Amount_Per_Person_TextBoxes(k) = Coupon_User_Form.Designer.Controls.Add("Forms.TextBox.1")
        
        Tax_Amount_TextBox_Name_Property(k) = Party_Names(k) + "_Tax_Amount_TextBox"
        
        With Tax_Amount_Per_Person_TextBoxes(k)
            .Name = Tax_Amount_TextBox_Name_Property(k)
            .Top = 156 + k * 24
            .Left = 450
            .Width = 90
            .Height = 18
            .Font.Size = 12
            .Font.Name = "Times New Roman"
        End With
        
        Set Meal_Cost_Per_Person_Labels(k) = Coupon_User_Form.Designer.Controls.Add("Forms.Label.1")
        
        With Meal_Cost_Per_Person_Labels(k)
            .Name = CStr(Party_Names(k)) + "_Meal_Cost_Label"
            .Caption = "Meal Cost:"
            .Top = 156 + k * 24
            .Left = 204
            .Width = 60
            .Height = 18
            .Font.Size = 12
            .Font.Name = "Times New Roman"
        End With
        
        Set Meal_Cost_Per_Person_TextBoxes(k) = Coupon_User_Form.Designer.Controls.Add("Forms.TextBox.1")
        
        Meal_Cost_TextBox_Name_Property(k) = Party_Names(k) + "_Meal_Cost_TextBox"
        
        With Meal_Cost_Per_Person_TextBoxes(k)
            .Name = Meal_Cost_TextBox_Name_Property(k)
            .Top = 156 + k * 24
            .Left = 276
            .Width = 90
            .Height = 18
            .Font.Size = 12
            .Font.Name = "Times New Roman"
        End With
        
        Set Coupon_Payment_Per_Person_Labels(k) = Coupon_User_Form.Designer.Controls.Add("Forms.Label.1")
        
        With Coupon_Payment_Per_Person_Labels(k)
            .Name = Party_Names(k) + "_Coupon_Payment_Label"
            .Caption = "Coupon Payment:"
            .Top = 156 + k * 24
            .Left = 552
            .Width = 102
            .Height = 18
            .Font.Size = 12
            .Font.Name = "Times New Roman"
        End With
        
        Set Coupon_Payment_Per_Person_TextBoxes(k) = Coupon_User_Form.Designer.Controls.Add("Forms.TextBox.1")
        
        Coupon_Payment_TextBox_Name_Property(k) = Party_Names(k) + "_Coupon_Payment_TextBox"
        
        With Coupon_Payment_Per_Person_TextBoxes(k)
            .Name = Coupon_Payment_TextBox_Name_Property(k)
            .Top = 156 + k * 24
            .Left = 666
            .Width = 90
            .Height = 18
            .Font.Size = 12
            .Font.Name = "Times New Roman"
        End With
        
    Next k
    
    Set OkButton = Coupon_User_Form.Designer.Controls.Add("Forms.CommandButton.1")
    
    With OkButton
        .Name = "Ok_Button"
        .Caption = "Ok"
        .Top = (156 + ((Number_Of_People_In_Party - 1) * 24)) + 120
        .Left = 528
        .Width = 108
        .Height = 24
        .Font.Size = 12
        .Font.Name = "Times New Roman"
    End With
    
    Set CancelButton = Coupon_User_Form.Designer.Controls.Add("Forms.CommandButton.1")
    
    With CancelButton
        .Name = "Cancel_Button"
        .Caption = "Cancel"
        .Top = (156 + ((Number_Of_People_In_Party - 1) * 24)) + 120
        .Left = 648.05
        .Width = 108
        .Height = 24
        .Font.Size = 12
        .Font.Name = "Times New Roman"
    End With

'    o = 0
'
    temp_string = Party_Names_Labels_Array(0).Caption
    
    If (Number_Of_People_In_Party > 1) Then

        For o = 1 To (Number_Of_People_In_Party - 1)

            temp_string = temp_string + ";" + Party_Names_Labels_Array(o).Caption

        Next o
        
    End If
    
    Set Party_Names_Holder_Label = Coupon_User_Form.Designer.Controls.Add("Forms.Label.1")
    
    With Party_Names_Holder_Label
        .Name = "Party_Names_Holder_Label"
        .Caption = temp_string
        .Top = 18
        .Left = 648.05
        .Width = 0
        .Height = 0
        .Font.Size = 12
        .Font.Name = "Times New Roman"
    End With
    
    With Coupon_User_Form
    
        .CodeModule.InsertLines 1, "Function Check_Value(Value As Double) As String"
        .CodeModule.InsertLines 2, ""
        .CodeModule.InsertLines 3, "    Dim Value_Into_String As String"
        .CodeModule.InsertLines 4, ""
        .CodeModule.InsertLines 5, "    Value_Into_String = CStr(Value)"
        .CodeModule.InsertLines 6, ""
        .CodeModule.InsertLines 7, "    If InStr(Value_Into_String, " + """.""" + ") = 0 Then"
        .CodeModule.InsertLines 8, ""
        .CodeModule.InsertLines 9, "        Check_Value = " + """$""" + " + Value_Into_String + " + """.00"""
        .CodeModule.InsertLines 10, ""
        .CodeModule.InsertLines 11, "   ElseIf Len(Value_Into_String) - InStr(Value_Into_String, " + """.""" + ") = 1 Then"
        .CodeModule.InsertLines 12, ""
        .CodeModule.InsertLines 13, "       Check_Value = " + """$""" + " + Value_Into_String + " + """0"""
        .CodeModule.InsertLines 14, ""
        .CodeModule.InsertLines 15, "   Else"
        .CodeModule.InsertLines 16, ""
        .CodeModule.InsertLines 17, "       Check_Value = " + """$""" + " + Value_Into_String"
        .CodeModule.InsertLines 18, ""
        .CodeModule.InsertLines 19, "   End If"
        .CodeModule.InsertLines 20, ""
        .CodeModule.InsertLines 21, "End Function"
        .CodeModule.InsertLines 22, ""
        .CodeModule.InsertLines 23, "Private Sub Ok_Button_Click()"
        .CodeModule.InsertLines 24, ""
        .CodeModule.InsertLines 25, "   Dim this_workbook As Workbook"
        .CodeModule.InsertLines 26, "   Dim this_worksheet As Worksheet"
        .CodeModule.InsertLines 27, ""
        .CodeModule.InsertLines 28, "   Dim n As Integer"
        .CodeModule.InsertLines 29, "   Dim q As Integer"
        .CodeModule.InsertLines 30, "   Dim r As Integer"
        .CodeModule.InsertLines 31, "   Dim s As Integer"
        .CodeModule.InsertLines 32, "   Dim t As Integer"
        .CodeModule.InsertLines 33, "   Dim u As Integer"
        .CodeModule.InsertLines 34, "   Dim v As Integer"
        .CodeModule.InsertLines 35, ""
        .CodeModule.InsertLines 36, "   Dim temp_string As String"
        .CodeModule.InsertLines 37, "   Dim altered_temp_string As String"
        .CodeModule.InsertLines 38, "   Dim flag_text_prompt As String"
        .CodeModule.InsertLines 39, ""
        .CodeModule.InsertLines 40, "   Dim Party_Names_Split_Array_Caption() As String"
        .CodeModule.InsertLines 41, "   Dim Party_Number_Split_Array() As String"
        .CodeModule.InsertLines 42, "   Dim Party_Names_Split_Array_Name() As String"
        .CodeModule.InsertLines 43, ""
        .CodeModule.InsertLines 44, "   Dim Meal_Cost_TextBox_Calculations_Array() As Double"
        .CodeModule.InsertLines 45, "   Dim Meal_Tax_TextBox_Calculations_Array() As Double"
        .CodeModule.InsertLines 46, "   Dim Coupon_Payment_TextBox_Calculations_Array() As Double"
        .CodeModule.InsertLines 47, "   Dim Coupon_Cost_Redeem_Value_TextBox_Calculations_Array() As Double"
        .CodeModule.InsertLines 48, ""
        .CodeModule.InsertLines 50, "   Dim Control_Looper As Control"
        .CodeModule.InsertLines 51, ""
        .CodeModule.InsertLines 52, "   Set this_worksheet = Excel.Application.ThisWorkbook.Worksheets.Add"
        .CodeModule.InsertLines 53, ""
        .CodeModule.InsertLines 54, "   this_worksheet.Name = " + """Data_Summary"""
        .CodeModule.InsertLines 55, ""
        .CodeModule.InsertLines 56, "   Worksheets(" + """Data_Summary""" + ").Range(" + """A2""" + ").Value = Me.Pre_Tax_Total_Label.Caption"
        .CodeModule.InsertLines 57, "   Worksheets(" + """Data_Summary""" + ").Range(" + """A3""" + ").Value = Me.Number_of_People_in_Party_Label.Caption"
        .CodeModule.InsertLines 58, "   Worksheets(" + """Data_Summary""" + ").Range(" + """A5""" + ").Value = Me.Coupon_Cost_Label.Caption"
        .CodeModule.InsertLines 59, "   Worksheets(" + """Data_Summary""" + ").Range(" + """A6""" + ").Value = Me.Coupon_Redeem_Value_Label.Caption"
        .CodeModule.InsertLines 60, ""
        .CodeModule.InsertLines 61, "   If InStr(Me.Party_Names_Holder_Label.Caption, " + """;""" + ") = 0 Or InStr(Me.Party_Names_Holder_Label.Caption, " + """;""" + ") <> 0 Then"
        .CodeModule.InsertLines 62, ""
        .CodeModule.InsertLines 63, "       temp_string = Me.Party_Names_Holder_Label.Caption + " + """;"""
        .CodeModule.InsertLines 64, ""
        .CodeModule.InsertLines 65, "   End If"
        .CodeModule.InsertLines 66, ""
        .CodeModule.InsertLines 67, "   altered_temp_string = Replace(temp_string, " + """--->;""" + ", " + """,""" + ")"
        .CodeModule.InsertLines 68, ""
        .CodeModule.InsertLines 69, "   Party_Names_Split_Array_Caption = Split(Me.Party_Names_Holder_Label.Caption, " + """;""" + ")"
        .CodeModule.InsertLines 70, "   Party_Number_Split_Array = Split(Me.Number_of_People_in_Party_Label.Caption, " + """:   """ + ")"
        .CodeModule.InsertLines 71, "   Party_Names_Split_Array_Name = Split(altered_temp_string, " + """,""" + ")"
        .CodeModule.InsertLines 72, ""
        .CodeModule.InsertLines 73, "   For n = 0 To (Val(Party_Number_Split_Array(1)) - 1)"
        .CodeModule.InsertLines 74, ""
        .CodeModule.InsertLines 75, "       Worksheets(" + """Data_Summary""" + ").Range(" + """A""" + " + CStr(n + 9)).Value = Party_Names_Split_Array_Caption(n)"
        .CodeModule.InsertLines 76, ""
        .CodeModule.InsertLines 77, "   Next n"
        .CodeModule.InsertLines 78, ""
        .CodeModule.InsertLines 79, "   With Worksheets(" + """Data_Summary""" + ").Columns(" + """A""" + ")"
        .CodeModule.InsertLines 80, "       .ColumnWidth = 26.14"
        .CodeModule.InsertLines 81, "   End With"
        .CodeModule.InsertLines 82, ""
        .CodeModule.InsertLines 83, "   With Worksheets(" + """Data_Summary""" + ").Columns(" + """C""" + ")"
        .CodeModule.InsertLines 84, "       .ColumnWidth = 10"
        .CodeModule.InsertLines 85, "   End With"
        .CodeModule.InsertLines 86, ""
        .CodeModule.InsertLines 87, "   With Worksheets(" + """Data_Summary""" + ").Columns(" + """F""" + ")"
        .CodeModule.InsertLines 88, "       .ColumnWidth = 8.43"
        .CodeModule.InsertLines 89, "   End With"
        .CodeModule.InsertLines 90, ""
        .CodeModule.InsertLines 91, "   With Worksheets(" + """Data_Summary""" + ").Columns(" + """I""" + ")"
        .CodeModule.InsertLines 92, "       .ColumnWidth = 16.86"
        .CodeModule.InsertLines 93, "   End With"
        .CodeModule.InsertLines 94, ""
        .CodeModule.InsertLines 95, "   ReDim Meal_Cost_TextBox_Calculations_Array(0 To Val(Party_Number_Split_Array(1)) - 1)"
        .CodeModule.InsertLines 96, "   ReDim Meal_Tax_TextBox_Calculations_Array(0 To Val(Party_Number_Split_Array(1)) - 1)"
        .CodeModule.InsertLines 97, "   ReDim Coupon_Payment_TextBox_Calculations_Array(0 To Val(Party_Number_Split_Array(1)) - 1)"
        .CodeModule.InsertLines 98, "   ReDim Coupon_Cost_Redeem_Value_TextBox_Calculations_Array(0 To 1)"
        .CodeModule.InsertLines 99, ""
        .CodeModule.InsertLines 100, "  r = 0"
        .CodeModule.InsertLines 101, ""
        .CodeModule.InsertLines 102, "  s = 0"
        .CodeModule.InsertLines 103, ""
        .CodeModule.InsertLines 104, "  t = 0"
        .CodeModule.InsertLines 105, ""
        .CodeModule.InsertLines 106, "  u = 0"
        .CodeModule.InsertLines 107, ""
        .CodeModule.InsertLines 108, "  numeric_value_counter = 0"
        .CodeModule.InsertLines 109, ""
        .CodeModule.InsertLines 110, "  flag_text_prompt = " + """One/Several of the user inputs is/are not acceptable: """
        .CodeModule.InsertLines 111, ""
        .CodeModule.InsertLines 112, "  For Each Control_Looper In Me.Controls"
        .CodeModule.InsertLines 113, ""
        .CodeModule.InsertLines 114, "      If TypeName(Control_Looper) = " + """TextBox""" + " Then"
        .CodeModule.InsertLines 115, ""
        .CodeModule.InsertLines 116, "          If Control_Looper.Name = Me.Coupon_Cost_TextBox.Name Then"
        .CodeModule.InsertLines 117, ""
        .CodeModule.InsertLines 118, "              If Not IsNumeric(Control_Looper.Text) Then"
        .CodeModule.InsertLines 119, ""
        .CodeModule.InsertLines 120, "                  flag_text_prompt = flag_text_prompt + " + """Coupon Cost TextBox contains a non-numeric value, """
        .CodeModule.InsertLines 121, ""
        .CodeModule.InsertLines 122, "              ElseIf Len(Control_Looper.Text) - InStr(Control_Looper.Text, " + """.""" + ") = 1 And InStr(Control_Looper.Text, " + """.""" + ") <> 0 And InStr(Control_Looper.Text, " + """$""" + ") = 0 Then"
        .CodeModule.InsertLines 123, ""
        .CodeModule.InsertLines 124, "                  Coupon_Cost_Redeem_Value_TextBox_Calculations_Array(u) = Control_Looper.Text"
        .CodeModule.InsertLines 125, ""
        .CodeModule.InsertLines 126, "                  Worksheets(" + """Data_Summary""" + ").Range(" + """B5""" + ").Value = " + """$""" + " + Control_Looper.Text + " + """0"""
        .CodeModule.InsertLines 127, ""
        .CodeModule.InsertLines 128, "                  u = u + 1"
        .CodeModule.InsertLines 129, ""
        .CodeModule.InsertLines 130, "                  numeric_value_counter = numeric_value_counter + 1"
        .CodeModule.InsertLines 131, ""
        .CodeModule.InsertLines 132, "              ElseIf Len(Control_Looper.Text) - InStr(Control_Looper.Text, " + """.""" + ") = 0 And InStr(Control_Looper.Text, " + """.""" + ") <> 0 And InStr(Control_Looper.Text, " + """$""" + ") = 0 Then"
        .CodeModule.InsertLines 133, ""
        .CodeModule.InsertLines 134, "                  Coupon_Cost_Redeem_Value_TextBox_Calculations_Array(u) = Control_Looper.Text"
        .CodeModule.InsertLines 135, ""
        .CodeModule.InsertLines 136, "                  Worksheets(" + """Data_Summary""" + ").Range(" + """B5""" + ").Value = " + """$""" + " + Control_Looper.Text + " + """00"""
        .CodeModule.InsertLines 137, ""
        .CodeModule.InsertLines 138, "                  u = u + 1"
        .CodeModule.InsertLines 139, ""
        .CodeModule.InsertLines 140, "                  numeric_value_counter = numeric_value_counter + 1"
        .CodeModule.InsertLines 141, ""
        .CodeModule.InsertLines 142, "              ElseIf InStr(Control_Looper, " + """.""" + ") = 0 And InStr(Control_Looper, " + """$""" + ") = 0 Then"
        .CodeModule.InsertLines 143, ""
        .CodeModule.InsertLines 144, "                  Coupon_Cost_Redeem_Value_TextBox_Calculations_Array(u) = Control_Looper.Text"
        .CodeModule.InsertLines 145, ""
        .CodeModule.InsertLines 146, "                  Worksheets(" + """Data_Summary""" + ").Range(" + """B5""" + ").Value = " + """$""" + " + Control_Looper.Text + " + """.00"""
        .CodeModule.InsertLines 147, ""
        .CodeModule.InsertLines 148, "                  u = u + 1"
        .CodeModule.InsertLines 149, ""
        .CodeModule.InsertLines 150, "                  numeric_value_counter = numeric_value_counter + 1"
        .CodeModule.InsertLines 151, ""
        .CodeModule.InsertLines 152, "              ElseIf InStr(Control_Looper.Text, " + """$""" + ") > 0 Then"
        .CodeModule.InsertLines 152, ""
        .CodeModule.InsertLines 153, "                  flag_text_prompt = flag_text_prompt + " + """Coupon Cost TextBox contains a $, """
        .CodeModule.InsertLines 154, ""
        .CodeModule.InsertLines 155, "                  u = u + 1"
        .CodeModule.InsertLines 156, ""
        .CodeModule.InsertLines 157, "              ElseIf InStr(Control_Looper.Text, " + """/""" + ") > 0 Then"
        .CodeModule.InsertLines 158, ""
        .CodeModule.InsertLines 159, "                  flag_text_prompt = flag_text_prompt + " + """Coupon Cost TextBox contains a /, """
        .CodeModule.InsertLines 160, ""
        .CodeModule.InsertLines 161, "                  u = u + 1"
        .CodeModule.InsertLines 162, ""
        .CodeModule.InsertLines 163, "              Else"
        .CodeModule.InsertLines 164, ""
        .CodeModule.InsertLines 165, "                  Coupon_Cost_Redeem_Value_TextBox_Calculations_Array(u) = Control_Looper.Text"
        .CodeModule.InsertLines 166, ""
        .CodeModule.InsertLines 167, "                  Worksheets(" + """Data_Summary""" + ").Range(" + """B5""" + ").Value = " + """$""" + " + CStr(Round(Val(Control_Looper.Text) + 0.00001, 2))"
        .CodeModule.InsertLines 168, ""
        .CodeModule.InsertLines 169, "                  u = u + 1"
        .CodeModule.InsertLines 170, ""
        .CodeModule.InsertLines 171, "                  numeric_value_counter = numeric_value_counter + 1"
        .CodeModule.InsertLines 172, ""
        .CodeModule.InsertLines 173, "              End If"
        .CodeModule.InsertLines 174, ""
        .CodeModule.InsertLines 175, "          ElseIf Control_Looper.Name = Me.Coupon_Redeem_Value_TextBox.Name Then"
        .CodeModule.InsertLines 176, ""
        .CodeModule.InsertLines 177, "              If Not IsNumeric(Control_Looper.Text) Then"
        .CodeModule.InsertLines 178, ""
        .CodeModule.InsertLines 179, "                  flag_text_prompt = flag_text_prompt + " + """Coupon Redeem Value TextBox contains a non-numeric value, """
        .CodeModule.InsertLines 180, ""
        .CodeModule.InsertLines 181, "              ElseIf Len(Control_Looper.Text) - InStr(Control_Looper.Text, " + """.""" + ") = 1 And InStr(Control_Looper.Text, " + """.""" + ") <> 0 And InStr(Control_Looper.Text, " + """$""" + ") = 0 Then"
        .CodeModule.InsertLines 182, ""
        .CodeModule.InsertLines 183, "                  Coupon_Cost_Redeem_Value_TextBox_Calculations_Array(u) = Control_Looper.Text"
        .CodeModule.InsertLines 184, ""
        .CodeModule.InsertLines 185, "                  Worksheets(" + """Data_Summary""" + ").Range(" + """B6""" + ").Value = " + """$""" + " + Control_Looper.Text + " + """0"""
        .CodeModule.InsertLines 186, ""
        .CodeModule.InsertLines 187, "                  u = u + 1"
        .CodeModule.InsertLines 188, ""
        .CodeModule.InsertLines 189, "                  numeric_value_counter = numeric_value_counter + 1"
        .CodeModule.InsertLines 190, ""
        .CodeModule.InsertLines 191, "              ElseIf Len(Control_Looper.Text) - InStr(Control_Looper.Text, " + """.""" + ") = 0 And InStr(Control_Looper.Text, " + """.""" + ") <> 0 And InStr(Control_Looper.Text, " + """$""" + ") = 0 Then"
        .CodeModule.InsertLines 192, ""
        .CodeModule.InsertLines 193, "                  Coupon_Cost_Redeem_Value_TextBox_Calculations_Array(u) = Control_Looper.Text"
        .CodeModule.InsertLines 194, ""
        .CodeModule.InsertLines 195, "                  Worksheets(" + """Data_Summary""" + ").Range(" + """B6""" + ").Value = " + """$""" + " + Control_Looper.Text + " + """00"""
        .CodeModule.InsertLines 196, ""
        .CodeModule.InsertLines 197, "                  u = u + 1"
        .CodeModule.InsertLines 198, ""
        .CodeModule.InsertLines 199, "                  numeric_value_counter = numeric_value_counter + 1"
        .CodeModule.InsertLines 200, ""
        .CodeModule.InsertLines 201, "              ElseIf InStr(Control_Looper, " + """.""" + ") = 0 And InStr(Control_Looper, " + """$""" + ") = 0 Then"
        .CodeModule.InsertLines 202, ""
        .CodeModule.InsertLines 203, "                  Coupon_Cost_Redeem_Value_TextBox_Calculations_Array(u) = Control_Looper.Text"
        .CodeModule.InsertLines 204, ""
        .CodeModule.InsertLines 205, "                  Worksheets(" + """Data_Summary""" + ").Range(" + """B6""" + ").Value = " + """$""" + " + Control_Looper.Text + " + """.00"""
        .CodeModule.InsertLines 206, ""
        .CodeModule.InsertLines 207, "                  u = u + 1"
        .CodeModule.InsertLines 208, ""
        .CodeModule.InsertLines 209, "                  numeric_value_counter = numeric_value_counter + 1"
        .CodeModule.InsertLines 210, ""
        .CodeModule.InsertLines 211, "              ElseIf InStr(Control_Looper.Text, " + """$""" + ") > 0 Then"
        .CodeModule.InsertLines 212, ""
        .CodeModule.InsertLines 213, "                  flag_text_prompt = flag_text_prompt + " + """Coupon Redeem Value TextBox contains a $, """
        .CodeModule.InsertLines 214, ""
        .CodeModule.InsertLines 215, "                  u = u + 1"
        .CodeModule.InsertLines 216, ""
        .CodeModule.InsertLines 217, "              ElseIf InStr(Control_Looper.Text, " + """/""" + ") > 0 Then"
        .CodeModule.InsertLines 218, ""
        .CodeModule.InsertLines 219, "                  flag_text_prompt = flag_text_prompt + " + """Coupon Redeem Value TextBox contains a /, """
        .CodeModule.InsertLines 220, ""
        .CodeModule.InsertLines 221, "                  u = u + 1"
        .CodeModule.InsertLines 222, ""
        .CodeModule.InsertLines 223, "              Else"
        .CodeModule.InsertLines 224, ""
        .CodeModule.InsertLines 225, "                  Coupon_Cost_Redeem_Value_TextBox_Calculations_Array(u) = Control_Looper.Text"
        .CodeModule.InsertLines 226, ""
        .CodeModule.InsertLines 227, "                  Worksheets(" + """Data_Summary""" + ").Range(" + """B6""" + ").Value = " + """$""" + " + CStr(Round(Val(Control_Looper.Text) + 0.00001, 2))"
        .CodeModule.InsertLines 228, ""
        .CodeModule.InsertLines 229, "                  u = u + 1"
        .CodeModule.InsertLines 230, ""
        .CodeModule.InsertLines 231, "                  numeric_value_counter = numeric_value_counter + 1"
        .CodeModule.InsertLines 232, ""
        .CodeModule.InsertLines 233, "              End If"
        .CodeModule.InsertLines 234, ""
        .CodeModule.InsertLines 235, "          ElseIf Control_Looper.Name = Party_Names_Split_Array_Name(r) + " + """_Meal_Cost_TextBox""" + " Then"
        .CodeModule.InsertLines 236, ""
        .CodeModule.InsertLines 237, "              If Not IsNumeric(Control_Looper.Text) Then"
        .CodeModule.InsertLines 238, ""
        .CodeModule.InsertLines 239, "                  flag_text_prompt = flag_text_prompt + Party_Names_Split_Array_Name(r) + " + """'s """ + " + " + """Meal Cost input""" + " + " + """ contains a non-numeric value, """
        .CodeModule.InsertLines 240, ""
        .CodeModule.InsertLines 241, "                  Worksheets(" + """Data_Summary""" + ").Range(" + """C""" + " + CStr(r + 9)).Value = " + """Meal Cost:"""
        .CodeModule.InsertLines 242, ""
        .CodeModule.InsertLines 243, "                  r = r + 1"
        .CodeModule.InsertLines 244, ""
        .CodeModule.InsertLines 245, "              ElseIf Len(Control_Looper.Text) - InStr(Control_Looper.Text, " + """.""" + ") = 1 And InStr(Control_Looper.Text, " + """.""" + ") <> 0 And InStr(Control_Looper.Text, " + """$""" + ") = 0 Then"
        .CodeModule.InsertLines 246, ""
        .CodeModule.InsertLines 247, "                  Worksheets(" + """Data_Summary""" + ").Range(" + """C""" + " + CStr(r + 9)).Value = " + """Meal Cost:"""
        .CodeModule.InsertLines 248, ""
        .CodeModule.InsertLines 249, "                  Meal_Cost_TextBox_Calculations_Array(r) = Control_Looper.Text"
        .CodeModule.InsertLines 250, ""
        .CodeModule.InsertLines 251, "                  Worksheets(" + """Data_Summary""" + ").Range(" + """D""" + " + CStr(r + 9)).Value = " + """$""" + " + Control_Looper.Text + " + """0"""
        .CodeModule.InsertLines 252, ""
        .CodeModule.InsertLines 253, "                  r = r + 1"
        .CodeModule.InsertLines 254, ""
        .CodeModule.InsertLines 255, "                  numeric_value_counter = numeric_value_counter + 1"
        .CodeModule.InsertLines 256, ""
        .CodeModule.InsertLines 257, "              ElseIf Len(Control_Looper.Text) - InStr(Control_Looper.Text, " + """.""" + ") = 0 And InStr(Control_Looper.Text, " + """.""" + ") <> 0 And InStr(Control_Looper.Text, " + """$""" + ") = 0 Then"
        .CodeModule.InsertLines 258, ""
        .CodeModule.InsertLines 259, "                  Worksheets(" + """Data_Summary""" + ").Range(" + """C""" + " + CStr(r + 9)).Value = " + """Meal Cost:"""
        .CodeModule.InsertLines 260, ""
        .CodeModule.InsertLines 261, "                  Meal_Cost_TextBox_Calculations_Array(r) = Control_Looper.Text"
        .CodeModule.InsertLines 262, ""
        .CodeModule.InsertLines 263, "                  Worksheets(" + """Data_Summary""" + ").Range(" + """D""" + " + CStr(r + 9)).Value = " + """$""" + " + Control_Looper.Text + " + """00"""
        .CodeModule.InsertLines 264, ""
        .CodeModule.InsertLines 265, "                  r = r + 1"
        .CodeModule.InsertLines 266, ""
        .CodeModule.InsertLines 267, "                  numeric_value_counter = numeric_value_counter + 1"
        .CodeModule.InsertLines 268, ""
        .CodeModule.InsertLines 269, "              ElseIf InStr(Control_Looper, " + """.""" + ") = 0 And InStr(Control_Looper, " + """$""" + ") = 0 Then"
        .CodeModule.InsertLines 270, ""
        .CodeModule.InsertLines 271, "                  Worksheets(" + """Data_Summary""" + ").Range(" + """C""" + " + CStr(r + 9)).Value = " + """Meal Cost:"""
        .CodeModule.InsertLines 272, ""
        .CodeModule.InsertLines 273, "                  Meal_Cost_TextBox_Calculations_Array(r) = Control_Looper.Text"
        .CodeModule.InsertLines 274, ""
        .CodeModule.InsertLines 275, "                  Worksheets(" + """Data_Summary""" + ").Range(" + """D""" + " + CStr(r + 9)).Value = " + """$""" + " + Control_Looper.Text + " + """.00"""
        .CodeModule.InsertLines 276, ""
        .CodeModule.InsertLines 277, "                  r = r + 1"
        .CodeModule.InsertLines 278, ""
        .CodeModule.InsertLines 279, "                  numeric_value_counter = numeric_value_counter + 1"
        .CodeModule.InsertLines 280, ""
        .CodeModule.InsertLines 281, "              ElseIf InStr(Control_Looper.Text, " + """$""" + ") > 0 Then"
        .CodeModule.InsertLines 282, ""
        .CodeModule.InsertLines 283, "                  flag_text_prompt = flag_text_prompt + Party_Names_Split_Array_Name(r) + " + """'s """ + " + " + """Meal Cost input""" + " + " + """ contains a $, """
        .CodeModule.InsertLines 284, ""
        .CodeModule.InsertLines 285, "                  r = r + 1"
        .CodeModule.InsertLines 286, ""
        .CodeModule.InsertLines 287, "              ElseIf InStr(Control_Looper.Text, " + """/""" + ") > 0 Then"
        .CodeModule.InsertLines 288, ""
        .CodeModule.InsertLines 289, "                  flag_text_prompt = flag_text_prompt + Party_Names_Split_Array_Name(r) + " + """'s """ + " + " + """Meal Cost input""" + " + " + """ contains a /, """
        .CodeModule.InsertLines 290, ""
        .CodeModule.InsertLines 291, "                  r = r + 1"
        .CodeModule.InsertLines 292, ""
        .CodeModule.InsertLines 293, "              Else"
        .CodeModule.InsertLines 294, ""
        .CodeModule.InsertLines 295, "                  Worksheets(" + """Data_Summary""" + ").Range(" + """C""" + " + CStr(r + 9)).Value = " + """Meal Cost:"""
        .CodeModule.InsertLines 296, ""
        .CodeModule.InsertLines 297, "                  Meal_Cost_TextBox_Calculations_Array(r) = Control_Looper.Text"
        .CodeModule.InsertLines 298, ""
        .CodeModule.InsertLines 299, "                  Worksheets(" + """Data_Summary""" + ").Range(" + """D""" + " + CStr(r + 9)).Value = " + """$""" + " + CStr(Round(Val(Control_Looper.Text) + 0.00001, 2))"
        .CodeModule.InsertLines 300, ""
        .CodeModule.InsertLines 301, "                  r = r + 1"
        .CodeModule.InsertLines 302, ""
        .CodeModule.InsertLines 303, "                  numeric_value_counter = numeric_value_counter + 1"
        .CodeModule.InsertLines 304, ""
        .CodeModule.InsertLines 305, "              End If"
        .CodeModule.InsertLines 306, ""
        .CodeModule.InsertLines 307, "          ElseIf Control_Looper.Name = Party_Names_Split_Array_Name(s) + " + """_Tax_Amount_TextBox""" + " Then"
        .CodeModule.InsertLines 308, ""
        .CodeModule.InsertLines 309, "              If Not IsNumeric(Control_Looper.Text) Then"
        .CodeModule.InsertLines 310, ""
        .CodeModule.InsertLines 311, "                  flag_text_prompt = flag_text_prompt + Party_Names_Split_Array_Name(s) + " + """'s """ + " + " + """Meal Tax input""" + " + " + """ contains a non-numeric value, """
        .CodeModule.InsertLines 312, ""
        .CodeModule.InsertLines 313, "                  Worksheets(" + """Data_Summary""" + ").Range(" + """F""" + " + CStr(s + 9)).Value = " + """Meal Tax:"""
        .CodeModule.InsertLines 314, ""
        .CodeModule.InsertLines 315, "                  s = s + 1"
        .CodeModule.InsertLines 316, ""
        .CodeModule.InsertLines 317, "              ElseIf Len(Control_Looper.Text) - InStr(Control_Looper.Text, " + """.""" + ") = 1 And InStr(Control_Looper.Text, " + """.""" + ") <> 0 And InStr(Control_Looper.Text, " + """$""" + ") = 0 Then"
        .CodeModule.InsertLines 318, ""
        .CodeModule.InsertLines 319, "                  Worksheets(" + """Data_Summary""" + ").Range(" + """F""" + " + CStr(s + 9)).Value = " + """Meal Tax:"""
        .CodeModule.InsertLines 320, ""
        .CodeModule.InsertLines 321, "                  Meal_Tax_TextBox_Calculations_Array(s) = Control_Looper.Text"
        .CodeModule.InsertLines 322, ""
        .CodeModule.InsertLines 323, "                  Worksheets(" + """Data_Summary""" + ").Range(" + """G""" + " + CStr(s + 9)).Value = " + """$""" + " + Control_Looper.Text + " + """0"""
        .CodeModule.InsertLines 324, ""
        .CodeModule.InsertLines 325, "                  s = s + 1"
        .CodeModule.InsertLines 326, ""
        .CodeModule.InsertLines 327, "                  numeric_value_counter = numeric_value_counter + 1"
        .CodeModule.InsertLines 328, ""
        .CodeModule.InsertLines 329, "              ElseIf Len(Control_Looper.Text) - InStr(Control_Looper.Text, " + """.""" + ") = 0 And InStr(Control_Looper.Text, " + """.""" + ") <> 0 And InStr(Control_Looper.Text, " + """$""" + ") = 0 Then"
        .CodeModule.InsertLines 330, ""
        .CodeModule.InsertLines 331, "                  Worksheets(" + """Data_Summary""" + ").Range(" + """F""" + " + CStr(s + 9)).Value = " + """Meal Tax:"""
        .CodeModule.InsertLines 332, ""
        .CodeModule.InsertLines 333, "                  Meal_Tax_TextBox_Calculations_Array(s) = Control_Looper.Text"
        .CodeModule.InsertLines 334, ""
        .CodeModule.InsertLines 335, "                  Worksheets(" + """Data_Summary""" + ").Range(" + """G""" + " + CStr(s + 9)).Value = " + """$""" + " + Control_Looper.Text + " + """00"""
        .CodeModule.InsertLines 336, ""
        .CodeModule.InsertLines 337, "                  s = s + 1"
        .CodeModule.InsertLines 338, ""
        .CodeModule.InsertLines 339, "                  numeric_value_counter = numeric_value_counter + 1"
        .CodeModule.InsertLines 340, ""
        .CodeModule.InsertLines 341, "              ElseIf InStr(Control_Looper, " + """.""" + ") = 0 And InStr(Control_Looper, " + """$""" + ") = 0 Then"
        .CodeModule.InsertLines 342, ""
        .CodeModule.InsertLines 343, "                  Worksheets(" + """Data_Summary""" + ").Range(" + """F""" + " + CStr(s + 9)).Value = " + """Meal Tax:"""
        .CodeModule.InsertLines 344, ""
        .CodeModule.InsertLines 345, "                  Meal_Tax_TextBox_Calculations_Array(s) = Control_Looper.Text"
        .CodeModule.InsertLines 346, ""
        .CodeModule.InsertLines 347, "                  Worksheets(" + """Data_Summary""" + ").Range(" + """G""" + " + CStr(s + 9)).Value = " + """$""" + " + Control_Looper.Text + " + """.00"""
        .CodeModule.InsertLines 348, ""
        .CodeModule.InsertLines 349, "                  s = s + 1"
        .CodeModule.InsertLines 350, ""
        .CodeModule.InsertLines 351, "                  numeric_value_counter = numeric_value_counter + 1"
        .CodeModule.InsertLines 352, ""
        .CodeModule.InsertLines 353, "              ElseIf InStr(Control_Looper.Text, " + """$""" + ") > 0 Then"
        .CodeModule.InsertLines 354, ""
        .CodeModule.InsertLines 355, "                  flag_text_prompt = flag_text_prompt + Party_Names_Split_Array_Name(s) + " + """'s """ + " + " + """Meal Tax input""" + " + " + """ contains a $, """
        .CodeModule.InsertLines 356, ""
        .CodeModule.InsertLines 357, "                  s = s + 1"
        .CodeModule.InsertLines 358, ""
        .CodeModule.InsertLines 359, "              ElseIf InStr(Control_Looper.Text, " + """/""" + ") > 0 Then"
        .CodeModule.InsertLines 360, ""
        .CodeModule.InsertLines 361, "                  flag_text_prompt = flag_text_prompt + Party_Names_Split_Array_Name(s) + " + """'s """ + " + " + """Meal Tax input""" + " + " + """ contains a /, """
        .CodeModule.InsertLines 362, ""
        .CodeModule.InsertLines 363, "                  s = s + 1"
        .CodeModule.InsertLines 364, ""
        .CodeModule.InsertLines 365, "              Else"
        .CodeModule.InsertLines 366, ""
        .CodeModule.InsertLines 367, "                  Worksheets(" + """Data_Summary""" + ").Range(" + """F""" + " + CStr(s + 9)).Value = " + """Meal Tax:"""
        .CodeModule.InsertLines 368, ""
        .CodeModule.InsertLines 369, "                  Meal_Tax_TextBox_Calculations_Array(s) = Control_Looper.Text"
        .CodeModule.InsertLines 370, ""
        .CodeModule.InsertLines 371, "                  Worksheets(" + """Data_Summary""" + ").Range(" + """G""" + " + CStr(s + 9)).Value = " + """$""" + " + CStr(Round(Val(Control_Looper.Text) + 0.00001, 2))"
        .CodeModule.InsertLines 372, ""
        .CodeModule.InsertLines 373, "                  s = s + 1"
        .CodeModule.InsertLines 374, ""
        .CodeModule.InsertLines 375, "                  numeric_value_counter = numeric_value_counter + 1"
        .CodeModule.InsertLines 376, ""
        .CodeModule.InsertLines 377, "              End If"
        .CodeModule.InsertLines 378, ""
        .CodeModule.InsertLines 379, "          ElseIf Control_Looper.Name = Party_Names_Split_Array_Name(t) + " + """_Coupon_Payment_TextBox""" + " Then"
        .CodeModule.InsertLines 380, ""
        .CodeModule.InsertLines 381, "              If Not IsNumeric(Control_Looper.Text) Then"
        .CodeModule.InsertLines 382, ""
        .CodeModule.InsertLines 383, "                  flag_text_prompt = flag_text_prompt + Party_Names_Split_Array_Name(t) + " + """'s """ + " + " + """Coupon Payment input""" + " + " + """ contains a non-numeric value, """
        .CodeModule.InsertLines 384, ""
        .CodeModule.InsertLines 385, "                  Worksheets(" + """Data_Summary""" + ").Range(" + """I""" + " + CStr(t + 9)).Value = " + """Coupon Payment:"""
        .CodeModule.InsertLines 386, ""
        .CodeModule.InsertLines 387, "                  t = t + 1"
        .CodeModule.InsertLines 388, ""
        .CodeModule.InsertLines 389, "              ElseIf Len(Control_Looper.Text) - InStr(Control_Looper.Text, " + """.""" + ") = 1 And InStr(Control_Looper.Text, " + """.""" + ") <> 0 And InStr(Control_Looper.Text, " + """$""" + ") = 0 Then"
        .CodeModule.InsertLines 390, ""
        .CodeModule.InsertLines 391, "                  Worksheets(" + """Data_Summary""" + ").Range(" + """I""" + " + CStr(t + 9)).Value = " + """Coupon Payment:"""
        .CodeModule.InsertLines 392, ""
        .CodeModule.InsertLines 393, "                  Coupon_Payment_TextBox_Calculations_Array(t) = Control_Looper.Text"
        .CodeModule.InsertLines 394, ""
        .CodeModule.InsertLines 395, "                  Worksheets(" + """Data_Summary""" + ").Range(" + """J""" + " + CStr(t + 9)).Value = " + """$""" + " + Control_Looper.Text + " + """0"""
        .CodeModule.InsertLines 396, ""
        .CodeModule.InsertLines 397, "                  t = t + 1"
        .CodeModule.InsertLines 398, ""
        .CodeModule.InsertLines 399, "                  numeric_value_counter = numeric_value_counter + 1"
        .CodeModule.InsertLines 400, ""
        .CodeModule.InsertLines 401, "              ElseIf Len(Control_Looper.Text) - InStr(Control_Looper.Text, " + """.""" + ") = 0 And InStr(Control_Looper.Text, " + """.""" + ") <> 0 And InStr(Control_Looper.Text, " + """$""" + ") = 0 Then"
        .CodeModule.InsertLines 402, ""
        .CodeModule.InsertLines 403, "                  Worksheets(" + """Data_Summary""" + ").Range(" + """I""" + " + CStr(t + 9)).Value = " + """Coupon Payment:"""
        .CodeModule.InsertLines 404, ""
        .CodeModule.InsertLines 405, "                  Coupon_Payment_TextBox_Calculations_Array(t) = Control_Looper.Text"
        .CodeModule.InsertLines 406, ""
        .CodeModule.InsertLines 407, "                  Worksheets(" + """Data_Summary""" + ").Range(" + """J""" + " + CStr(t + 9)).Value = " + """$""" + " + Control_Looper.Text + " + """00"""
        .CodeModule.InsertLines 408, ""
        .CodeModule.InsertLines 409, "                  t = t + 1"
        .CodeModule.InsertLines 410, ""
        .CodeModule.InsertLines 411, "                  numeric_value_counter = numeric_value_counter + 1"
        .CodeModule.InsertLines 412, ""
        .CodeModule.InsertLines 413, "              ElseIf InStr(Control_Looper, " + """.""" + ") = 0 And InStr(Control_Looper, " + """$""" + ") = 0 Then"
        .CodeModule.InsertLines 414, ""
        .CodeModule.InsertLines 415, "                  Worksheets(" + """Data_Summary""" + ").Range(" + """I""" + " + CStr(t + 9)).Value = " + """Coupon Payment:"""
        .CodeModule.InsertLines 416, ""
        .CodeModule.InsertLines 417, "                  Coupon_Payment_TextBox_Calculations_Array(t) = Control_Looper.Text"
        .CodeModule.InsertLines 418, ""
        .CodeModule.InsertLines 419, "                  Worksheets(" + """Data_Summary""" + ").Range(" + """J""" + " + CStr(t + 9)).Value = " + """$""" + " + Control_Looper.Text + " + """.00"""
        .CodeModule.InsertLines 420, ""
        .CodeModule.InsertLines 421, "                  t = t + 1"
        .CodeModule.InsertLines 422, ""
        .CodeModule.InsertLines 423, "                  numeric_value_counter = numeric_value_counter + 1"
        .CodeModule.InsertLines 424, ""
        .CodeModule.InsertLines 425, "              ElseIf InStr(Control_Looper.Text, " + """$""" + ") > 0 Then"
        .CodeModule.InsertLines 426, ""
        .CodeModule.InsertLines 427, "                  flag_text_prompt = flag_text_prompt + Party_Names_Split_Array_Name(t) + " + """'s """ + " + " + """Coupon Payment input""" + " + " + """ contains a $, """
        .CodeModule.InsertLines 428, ""
        .CodeModule.InsertLines 429, "                  t = t + 1"
        .CodeModule.InsertLines 430, ""
        .CodeModule.InsertLines 431, "              ElseIf InStr(Control_Looper.Text, " + """/""" + ") > 0 Then"
        .CodeModule.InsertLines 432, ""
        .CodeModule.InsertLines 433, "                  flag_text_prompt = flag_text_prompt + Party_Names_Split_Array_Name(t) + " + """'s """ + " + " + """Coupon Payment input""" + " + " + """ contains a /, """
        .CodeModule.InsertLines 434, ""
        .CodeModule.InsertLines 435, "                  t = t + 1"
        .CodeModule.InsertLines 436, ""
        .CodeModule.InsertLines 437, "              Else"
        .CodeModule.InsertLines 438, ""
        .CodeModule.InsertLines 439, "                  Worksheets(" + """Data_Summary""" + ").Range(" + """I""" + " + CStr(t + 9)).Value = " + """Coupon Payment:"""
        .CodeModule.InsertLines 440, ""
        .CodeModule.InsertLines 441, "                  Coupon_Payment_TextBox_Calculations_Array(t) = Control_Looper.Text"
        .CodeModule.InsertLines 442, ""
        .CodeModule.InsertLines 443, "                  Worksheets(" + """Data_Summary""" + ").Range(" + """J""" + " + CStr(t + 9)).Value = " + """$""" + " + CStr(Round(Val(Control_Looper.Text) + 0.00001, 2))"
        .CodeModule.InsertLines 444, ""
        .CodeModule.InsertLines 445, "                  t = t + 1"
        .CodeModule.InsertLines 446, ""
        .CodeModule.InsertLines 447, "                  numeric_value_counter = numeric_value_counter + 1"
        .CodeModule.InsertLines 448, ""
        .CodeModule.InsertLines 449, "              End If"
        .CodeModule.InsertLines 450, ""
        .CodeModule.InsertLines 451, "          End If"
        .CodeModule.InsertLines 452, ""
        .CodeModule.InsertLines 453, "      End If"
        .CodeModule.InsertLines 454, ""
        .CodeModule.InsertLines 455, "  Next Control_Looper"
        .CodeModule.InsertLines 456, ""
        .CodeModule.InsertLines 457, "  If numeric_value_counter <> (3 * Val(Party_Number_Split_Array(1)) + 2) Then"
        .CodeModule.InsertLines 458, ""
        .CodeModule.InsertLines 459, "      MsgBox (flag_text_prompt + " + """ so the input(s) cannot be accepted. Please try again.""" + ")"
        .CodeModule.InsertLines 460, ""
        .CodeModule.InsertLines 461, "      this_worksheet.Delete"
        .CodeModule.InsertLines 462, ""
        .CodeModule.InsertLines 463, "      Unload Me"
        .CodeModule.InsertLines 464, ""
        .CodeModule.InsertLines 465, "  Else"
        .CodeModule.InsertLines 466, ""
        .CodeModule.InsertLines 467, "      v = 0"
        .CodeModule.InsertLines 468, ""
        .CodeModule.InsertLines 469, "      For v = 0 To (Val(Party_Number_Split_Array(1)) - 1)"
        .CodeModule.InsertLines 470, ""
        .CodeModule.InsertLines 471, "          Worksheets(" + """Data_Summary""" + ").Range(" + """A""" + " + CStr(10 + Val(Party_Number_Split_Array(1)) + v)).Value = Party_Names_Split_Array_Caption(v)"
        .CodeModule.InsertLines 472, ""
        .CodeModule.InsertLines 473, "          Worksheets(" + """Data_Summary""" + ").Range(" + """C""" + " + CStr(10 + Val(Party_Number_Split_Array(1)) + v)).Value = " + """Prorated Redeem Value:"""
        .CodeModule.InsertLines 474, ""
        .CodeModule.InsertLines 475, "          Worksheets(" + """Data_Summary""" + ").Range(" + """D""" + " + CStr(10 + Val(Party_Number_Split_Array(1)) + v)).Value = Me.Check_Value(Round((Val(Coupon_Payment_TextBox_Calculations_Array(v)) / Val(Coupon_Cost_Redeem_Value_TextBox_Calculations_Array(0))) * Val(Coupon_Cost_Redeem_Value_TextBox_Calculations_Array(1)) + 0.00001, 2))"
        .CodeModule.InsertLines 476, ""
        .CodeModule.InsertLines 477, "          Worksheets(" + """Data_Summary""" + ").Range(" + """F""" + " + CStr(10 + Val(Party_Number_Split_Array(1)) + v)).Value = " + """Reduced Payment:"""
        .CodeModule.InsertLines 478, ""
        .CodeModule.InsertLines 479, "          Worksheets(" + """Data_Summary""" + ").Range(" + """G""" + " + CStr(10 + Val(Party_Number_Split_Array(1)) + v)).Value = Me.Check_Value(Round(Val(Meal_Cost_TextBox_Calculations_Array(v)) - Val(Worksheets(" + """Data_Summary""" + ").Range(" + """D""" + " + CStr(10 + Val(Party_Number_Split_Array(1)) + v)).Value) + 0.00001, 2))"
        .CodeModule.InsertLines 480, ""
        .CodeModule.InsertLines 481, "          Worksheets(" + """Data_Summary""" + ").Range(" + """I""" + " + CStr(10 + Val(Party_Number_Split_Array(1)) + v)).Value = " + """Total Payment Due:"""
        .CodeModule.InsertLines 482, ""
        .CodeModule.InsertLines 483, "          Worksheets(" + """Data_Summary""" + ").Range(" + """J""" + " + CStr(10 + Val(Party_Number_Split_Array(1)) + v)).Value = Me.Check_Value(Round(Val(Worksheets(" + """Data_Summary""" + ").Range(" + """G""" + " + CStr(10 + Val(Party_Number_Split_Array(1)) + v)).Value) + Val(Meal_Tax_TextBox_Calculations_Array(v)) + 0.00001, 2))"
        .CodeModule.InsertLines 484, ""
        .CodeModule.InsertLines 485, "      Next v"
        .CodeModule.InsertLines 486, ""
        .CodeModule.InsertLines 487, "      Unload Me"
        .CodeModule.InsertLines 488, ""
        .CodeModule.InsertLines 489, "  End If"
        .CodeModule.InsertLines 490, ""
        .CodeModule.InsertLines 491, "End Sub"

    End With
    
    VBA.UserForms.Add(Coupon_User_Form.Name).Show
    
End Sub

'Worksheets(" + """Data_Summary""" + ").Range(" + """B6""" + ").Value = Me.Groupon_Redeem_Value_TextBox.Value"
