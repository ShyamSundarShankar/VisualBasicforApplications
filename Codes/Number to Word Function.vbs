Public Numbers As Variant, Tens As Variant

Sub SetNums()
    Numbers = Array("", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen")
    Tens = Array("", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety")
End Sub

Function WordNum(MyNumber As Double) As String
    Dim DecimalPosition As Integer, ValNo As Variant, StrNo As String
    Dim NumStr As String, n As Integer, Temp1 As String, Temp2 As String
    ' This macro was written by Chris Mead - www.MeadInKent.co.uk
    If Abs(MyNumber) > 999999999 Then
        WordNum = "Value too large"
        Exit Function
    End If
    SetNums
    ' String representation of amount (excl decimals)
    NumStr = Right("000000000" & Trim(Str(Int(Abs(MyNumber)))), 9)
    ValNo = Array(0, Val(Mid(NumStr, 1, 3)), Val(Mid(NumStr, 4, 3)), Val(Mid(NumStr, 7, 3)))
    For n = 3 To 1 Step -1    'analyse the absolute number as 3 sets of 3 digits
        StrNo = Format(ValNo(n), "000")
        If ValNo(n) > 0 Then
            Temp1 = GetTens(Val(Right(StrNo, 2)))
            If Left(StrNo, 1) <> "0" Then
                Temp2 = Numbers(Val(Left(StrNo, 1))) & " hundred"
                If Temp1 <> "" Then Temp2 = Temp2 & " and "
            Else
                Temp2 = ""
            End If
            If n = 3 Then
                If Temp2 = "" And ValNo(1) + ValNo(2) > 0 Then Temp2 = "and "
                WordNum = Trim(Temp2 & Temp1)
            End If
            If n = 2 Then WordNum = Trim(Temp2 & Temp1 & " thousand " & WordNum)
            If n = 1 Then WordNum = Trim(Temp2 & Temp1 & " million " & WordNum)
        End If
    Next n
    NumStr = Trim(Str(Abs(MyNumber)))
    ' Values after the decimal place
    DecimalPosition = InStr(NumStr, ".")
    Numbers(0) = "Zero"
    If DecimalPosition > 0 And DecimalPosition < Len(NumStr) Then
        Temp1 = " point"
        For n = DecimalPosition + 1 To Len(NumStr)
            Temp1 = Temp1 & " " & Numbers(Val(Mid(NumStr, n, 1)))
        Next n
        WordNum = WordNum & Temp1
    End If
    If Len(WordNum) = 0 Or Left(WordNum, 2) = " p" Then
        WordNum = "Zero" & WordNum
    End If
End Function

Function GetTens(TensNum As Integer) As String
' Converts a number from 0 to 99 into text.
    If TensNum <= 19 Then
        GetTens = Numbers(TensNum)
    Else
        Dim MyNo As String
        MyNo = Format(TensNum, "00")
        GetTens = Tens(Val(Left(MyNo, 1))) & " " & Numbers(Val(Right(MyNo, 1)))
    End If
End Function
