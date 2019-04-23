Attribute VB_Name = "Module2"
Sub Coupon()
    Dim Number_Of_People_In_Party As Variant
    Dim response As VbMsgBoxResult
    Dim key As Integer
    Dim response2 As VbMsgBoxResult
    Dim total_cost As Variant
    Dim response3 As VbMsgBoxResult
    Dim key2 As Integer
    Dim cents_counter As Integer
    Dim response4 As VbMsgBoxResult
    Dim Party_Names() As String
    Dim Party_Names_Variable As Variant
    Dim i As Integer
    Dim response5 As VbMsgBoxResult
    Dim j As Integer
    Dim Total_Party_Names_String As String
    Dim response6 As VbMsgBoxResult
    Dim Break_Party_Names_New_Line() As String
    Dim comma_counter As Integer
    Dim response7 As VbMsgBoxResult
    
    Do
    
        Number_Of_People_In_Party = InputBox("Please enter the number of people in your party.")
    
        If Not IsNumeric(Number_Of_People_In_Party) Then
     
            response = MsgBox("This is not a number. Would you like to enter again?", vbYesNo)
        
        ElseIf Number_Of_People_In_Party = 0 Then
            
            response = MsgBox(Number_Of_People_In_Party + " is not a valid party size. Would you like to enter again?", vbYesNo)
        
        ElseIf Not Val(Number_Of_People_In_Party) = Int(Val(Number_Of_People_In_Party)) Then
        
            response = MsgBox(Number_Of_People_In_Party + " is not a natural number greater than zero. Would you like to enter again?", vbYesNo)
            
        Else
        
            MsgBox "Your party size is " + Number_Of_People_In_Party + "."
            
            key = 1
            
        End If
        
    Loop Until response = vbNo Or key = 1
        
    If key = 1 Then
        
        response2 = MsgBox("Would you like to continue?", vbYesNo)
            
        If response2 = vbYes Then
            
            Do
                
                total_cost = InputBox("Please enter the total cost on your bill, not including tax, prior to any deductions being made.")
                    
                If Not IsNumeric(total_cost) Then
                    
                    response3 = MsgBox(total_cost + " is not a rational number. Would you like to try again?", vbYesNo)
                        
                ElseIf Len(total_cost) - InStr(total_cost, ".") = 1 And InStr(total_cost, ".") <> 0 And InStr(total_cost, "$") = 0 Then
                    
                    MsgBox ("The total cost on the bill, including tax, prior to any deductions is $" + total_cost + "0.")
                    
                    total_cost = CStr(total_cost) + "0"
                        
                    key2 = 1
                    
                ElseIf Len(total_cost) - InStr(total_cost, ".") = 0 And InStr(total_cost, ".") <> 0 And InStr(total_cost, "$") = 0 Then
                
                    MsgBox ("The total cost on the bill, including tax, prior to any deductions is $" + total_cost + "00.")
                    
                    total_cost = CStr(total_cost) + "00"
                        
                    key2 = 1
                    
                ElseIf InStr(total_cost, ".") = 0 And InStr(total_cost, "$") = 0 Then
                
                    MsgBox ("The total cost on the bill, including tax, prior to any deductions is $" + total_cost + ".00.")
                    
                    total_cost = CStr(total_cost) + ".00"
                
                    key2 = 1
                    
                ElseIf InStr(total_cost, "$") > 0 Then
                
                    response3 = MsgBox("you entered " + total_cost + "." + " Please enter the total amount without the dollar sign. Would you like to continue?", vbYesNo)
                    
                Else
                
                    MsgBox ("The total cost on the bill, including tax, prior to any deductions is $" + total_cost + ".")
                    
                    key2 = 1
                       
                End If
                    
            Loop Until response3 = vbNo Or key2 = 1
                
        End If
            
    End If
    
    If key2 = 1 Then
    
        Do
    
            response4 = MsgBox("Would you like to continue?", vbYesNo)
    
            If response4 = vbYes Then
            
                Party_Names_Variable = Party_Names()
        
                ReDim Party_Names(0 To (Val(Number_Of_People_In_Party)) - 1)
        
                i = 1
        
                Do
                
                    If response4 = vbYes And i < Val(Number_Of_People_In_Party) Then
            
                        Party_Names(i - 1) = InputBox("Please enter the name of person " + CStr(i) + " in your party.")
                
                        response4 = MsgBox("Would you like to continue?", vbYesNo)
                
                        i = i + 1
                        
                    ElseIf response4 = vbYes And i = Val(Number_Of_People_In_Party) Then
                    
                        Party_Names(i - 1) = InputBox("Please enter the name of person " + CStr(i) + " in your party.")
                        
                        i = i + 1
                        
                    End If
                
                Loop While i < Val(Number_Of_People_In_Party) + 1 And response4 = vbYes
                
                j = 0
                
                Total_Party_Names_String = "The party members are: "
                
                For j = 0 To Val(Number_Of_People_In_Party) - 1
                
                    If j = Val(Number_Of_People_In_Party) - 1 Then
                    
                        Total_Party_Names_String = Total_Party_Names_String + Party_Names(j)
                        
                    Else
                    
                        Total_Party_Names_String = Total_Party_Names_String + Party_Names(j) + ", "
                        
                    End If
                    
                Next
                
                Break_Party_Names_New_Line() = Split(Total_Party_Names_String, ":")
                
                response6 = MsgBox(Break_Party_Names_New_Line(0) + ":" + Chr$(10) + Chr$(10) + Break_Party_Names_New_Line(1) + Chr$(10) + Chr$(10) + "Is this correct?", vbYesNo)
                
            End If
        
        Loop Until response6 = vbYes Or response4 = vbNo
        
    End If
    
    response7 = MsgBox("Would you like to continue?", vbYesNo)
    
    If response7 = vbYes Then
    
        Do
        
            Call Module3.CouponForm(Int(Number_Of_People_In_Party), total_cost, Party_Names)
            
        Loop Until response7 = vbYes
        
    End If
    
End Sub
