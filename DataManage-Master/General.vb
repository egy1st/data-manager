Module General

    Public aEncrypt() As Integer = {7, 13, 27, 1, 18, 3, 29, 33, 5, 12, 17, 17, 18, 4, 4, 11, 9, 23, 19, 5, 2, 14, 13, 13, 13}

    Public Function ZeroPad(ByVal str_String As String, ByVal int_Count As Byte) As String
        If str_String <> "" Then
            Return (New String("0", int_Count - Len(Trim(str_String))) & Trim(str_String))
        End If
    End Function

    Public Function CheckSum(ByVal strNum As String) As Boolean
        Dim intCheckSum, blnDoubleFlag, X, intDigit As Integer

        For X = Len(strNum) To 1 Step -1
            intDigit = Asc(Mid$(strNum, X, 1))
            If intDigit > 47 Then
                If intDigit < 58 Then
                    intDigit = intDigit - 48

                    If blnDoubleFlag Then
                        intDigit = intDigit + intDigit
                        If intDigit > 9 Then
                            intDigit = intDigit - 9
                        End If
                    End If
                    blnDoubleFlag = Not blnDoubleFlag
                    intCheckSum = intCheckSum + intDigit
                    If intCheckSum > 9 Then
                        intCheckSum = intCheckSum - 10
                    End If
                End If
            End If
        Next
        If intCheckSum = 0 Then Return True
    End Function

    Public Function MakeMe(ByVal OldStr As String) As String
        Dim NewStr As String = ""
        Dim Num As Integer
        Dim MyChar As Char

        For Num = 1 To OldStr.Length
            MyChar = Mid(OldStr, Num, 1)
            NewStr += ZeroPad((Asc(MyChar) + aEncrypt(Num Mod 24)).ToString, 3)
        Next Num
        Return NewStr
    End Function

    Public Sub MakeMeFromFile(ByVal FileRead As String, ByVal FileWrite As String)
        Dim NewStr As String = ""
        Dim Num As Integer
        Dim MyChar As Char
        Dim FileNum1 As Integer
        Dim FileNum2 As Integer

        FileNum1 = FreeFile()
        FileOpen(FileNum1, FileRead, OpenMode.Binary, OpenAccess.Read)
        FileNum2 = FreeFile()
        FileOpen(FileNum2, FileWrite, OpenMode.Binary, OpenAccess.Write)

        For Num = 1 To FileLen(FileRead)
            FileGet(FileNum1, MyChar, Num)
            NewStr += ZeroPad((Asc(MyChar) + aEncrypt(Num Mod 24)).ToString, 3)
        Next Num
        FilePut(FileNum2, NewStr, 1)

        FileClose(FileNum1)
        FileClose(FileNum2)
    End Sub

    Public Function SolveMeFromFile(ByVal FileRead As String, Optional ByVal FileWrite As String = "") As String
        Dim NewStr As String
        Dim Num As Integer
        Dim MyChar As String = "   "
        Dim FileNum1 As Integer
        Dim FileNum2 As Integer


        FileNum1 = FreeFile()
        FileOpen(FileNum1, FileRead, OpenMode.Binary, OpenAccess.Read)
        If FileWrite <> "" Then
            FileNum2 = FreeFile()
            FileOpen(FileNum2, FileWrite, OpenMode.Binary, OpenAccess.Write)
        End If
        Dim nLoop As Integer = 0

        For Num = 1 To FileLen(FileRead) Step 3
            nLoop += 1
            FileGet(FileNum1, MyChar, Num, True)
            NewStr += Chr(CInt(MyChar) - aEncrypt(nLoop Mod 24))
        Next Num

        If FileWrite <> "" Then
            FilePut(FileNum2, NewStr)
            FileClose(FileNum2)
            FileClose(FileNum1)
        Else
            Return NewStr
            FileClose(FileNum1)
        End If

    End Function


    Public Function SolveMe(ByVal NewStr As String) As String
        Dim OldStr As String
        Dim Num As Integer
        Dim MyChar As Char
        Dim nLoop As Integer = 0
        For Num = 1 To NewStr.Length Step 3
            nLoop += 1
            MyChar = Chr(CInt(Mid(NewStr, Num, 3)) - aEncrypt(nLoop Mod 24))
            OldStr += MyChar
        Next Num
        Return OldStr
    End Function

    Public Function EncryptMe(ByVal OldStr As String) As String
        Dim NewStr As String = ""
        Dim Num As Integer
        Dim MyChar As Char

        For Num = 1 To OldStr.Length
            MyChar = Mid(OldStr, Num, 1)
            NewStr += Chr(Asc(MyChar) - 33)
        Next Num
        Return NewStr
    End Function

    Public Function DecryptMe(ByVal NewStr As String) As String
        Dim OldStr As String
        Dim Num As Integer
        Dim MyChar As Char
        Dim nLoop As Integer = 0
        nLoop += 1
        For Num = 1 To NewStr.Length
            MyChar = Chr(Asc(Mid(NewStr, Num, 1)) + 33)
            OldStr += MyChar
        Next Num
        Return OldStr
    End Function

End Module
