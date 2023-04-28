Function WordToNumberRange1(inputString As String) As Double
        
            Dim allowedStrings As Variant
            Dim splittedParts As Variant
            Dim str As Variant
            Dim finalResult As Double
            Dim result As Double
            Dim isValidInput As Boolean
            
            ' Define allowed strings list
            allowedStrings = Array("zero", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety", "hundred", "thousand", "million", "billion", "trillion", "dollars", "cents", "")
            
            ' Reset variables
            result = 0
            finalResult = 0
            isValidInput = False
            
            ' Validate input
            If Not inputString = vbNullString And Len(inputString) > 0 Then
                inputString = Replace(inputString, "-", " ")
                inputString = Replace(inputString, "$", " ")
                inputString = LCase(Replace(inputString, " and", " "))
                splittedParts = Split(Trim(inputString), " ")
                
                ' Check if all the parts are allowed strings
                For Each str In splittedParts
                    If Not IsError(Application.Match(str, allowedStrings, 0)) Then
                        isValidInput = True
                    Else
                        isValidInput = False
                        Debug.Print "Invalid word found: " & str
                        MsgBox "Invalid word found: " & str, vbExclamation + vbOKOnly, "Input Error"
                        Exit For
                    End If
                Next str
                
                If isValidInput Then
                    ' Convert words to numbers
                    For Each str In splittedParts
                        Select Case str
                            Case "zero"
                                result = result + 0
                                
                            Case "one"
                                result = result + 1
                            Case "two"
                                result = result + 2
                            Case "three"
                                result = result + 3
                            Case "four"
                                result = result + 4
                            Case "five"
                                result = result + 5
                            Case "six"
                                result = result + 6
                            Case "seven"
                                result = result + 7
                            Case "eight"
                                result = result + 8
                            Case "nine"
                                result = result + 9
                            Case "ten"
                                result = result + 10
                            Case "eleven"
                                result = result + 11
                            Case "twelve"
                                result = result + 12
                            Case "thirteen"
                                result = result + 13
                            Case "fourteen"
                                result = result + 14
                            Case "fifteen"
                                result = result + 15
                            Case "sixteen"
                                result = result + 16
                            Case "seventeen"
                                result = result + 17
                            Case "eighteen"
                                result = result + 18
                            Case "nineteen"
                                result = result + 19
                            Case "twenty"
                                result = result + 20
                            Case "thirty"
                                result = result + 30
                            Case "forty"
                                result = result + 40
                            Case "fifty"
                                result = result + 50
                            Case "sixty"
                                        result = result + 60
                                    Case "seventy"
                                        result = result + 70
                                    Case "eighty"
                                        result = result + 80
                                    Case "ninety"
                                        result = result + 90
                                    Case "hundred"
                                        result = result * 100
                                    Case "thousand"
                                        result = result * 1000
                                        finalResult = finalResult + result
                                        result = 0
                                    Case "million"
                                        result = result * 1000000
                                        finalResult = finalResult + result
                                        result = 0
                                    Case "billion"
                                        result = result * 1000000000
                                        finalResult = finalResult + result
                                        result = 0
                                    Case "trillion"
                                        result = result * 1000000000000#
                                        finalResult = finalResult + result
                                        result = 0
                                    Case "dollars"
                                        finalResult = finalResult + result
                                        result = 0
                                    Case "cents"
                                        result = result / 100
                                        finalResult = finalResult + result
                                        result = 0
                                End Select
                            Next str
                        End If
                        finalResult = finalResult + result
                       
                    End If
              WordToNumberRange1 = finalResult + 0
               
            
            End Function
   
