'<ComClass(AppProtector.ClassId, AppProtector.InterfaceId, AppProtector.EventsId)> _
Friend Class MyProtection

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    'Public Const ClassId As String = "91D83C50-88C6-431D-AE10-274C275E3883"
    'Public Const InterfaceId As String = "5A5EC27D-DADF-4B72-A214-DDB5F4B84BD1"
    'Public Const EventsId As String = "6DA345A7-EB62-4E16-B6AF-E28F0F534B27"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.

    Private Company As String = "Egyfirst Software, Inc."
    Private AuthorName As String = "Mohamed Aly Abbas Mohamed "
    Private AuthorNation As String = "Egypt-Alexandria-Smouha "
    Private AuthorPhone As String = "20-107500290 "
    Private AuthorMail As String = "mohamed@alyabbas.com "
    Private prod_website As String = "http://www.mygoldensoft.com"

    Public NotLicensed As Boolean = False
    Private KeyAlgorithm1 As Integer
    Private KeyAlgorithm2 As Integer
    Private KeyAlgorithm3 As Integer
    Private LicenseFile As String
    Private Ended As Boolean = False
    Private BuyNowURL As String

    Private AuthorForm As New System.Windows.Forms.Form()
    Private AuthorLabel1 As New System.Windows.Forms.RichTextBox()
    Private AuthorLabel2 As New System.Windows.Forms.Label()
    Private AuthorLabel3 As New System.Windows.Forms.Label()
    Private AuthorText2 As New System.Windows.Forms.TextBox()
    Private AuthorPrgress As New System.Windows.Forms.ProgressBar()
    Private AuthorButton1 As New System.Windows.Forms.Button()
    Private AuthorButton2 As New System.Windows.Forms.Button()
    Private AuthorBuyNow As New System.Windows.Forms.LinkLabel()
    Private ProductID As String = "maaat05"
    Private WinSys As String
    Private ProductName As String
    Private CompanyInfo As String
    Private DaysLimit As Integer = 30
    Private SourceFile As String
    Private SourceFile2 As String = Space(33)
    'sourcefile "windows_maaat01.dll" 	

    Public Sub New()
        MyBase.New()

    End Sub

    Public Sub SetInformation(ByVal str_ProductName As String)
        On Error Resume Next

        Dim WshShell As New Object()


        ProductName = str_ProductName
        CompanyInfo = "Product Name: " + ProductName + vbCrLf + "Company Name: " + Company + vbCrLf + "Home Page: " + prod_website + vbCrLf + "License: Free 30 Days Trial Version"


        WshShell = CreateObject("WScript.Shell")

        BuyNowURL = WshShell.RegRead("HKEY_LOCAL_MACHINE\Software\Digital River\SoftwarePassport\" + Company + ProductName + "\0\BuyURL")
        If BuyNowURL = "" Then
            BuyNowURL = WshShell.RegRead("HKEY_CURRENT_USER\Software\Digital River\SoftwarePassport\" + Company + ProductName + "\0\BuyURL")
        End If
        If BuyNowURL = "" Then
            BuyNowURL = "http://mygoldensoft.com/ordernow.htm"
        End If
        
    End Sub

    Public Sub SetLicense(ByVal int_DaysLimit As Integer)
        DaysLimit = int_DaysLimit
    End Sub

    Public Sub SetAlgorithms(ByVal int_Algorithms1 As Integer, ByVal int_Algorithms2 As Integer, ByVal int_Algorithms3 As Integer, ByVal str_Algorithms4 As String)
        Dim X As Byte
        str_Algorithms4 = str_Algorithms4.ToLower

        If str_Algorithms4.Length <> 7 Then
            MsgBox("str_Algorithms4 must a string of 7 characters")
            Exit Sub
        End If

        'For X = 1 To 7
        'If Not (Asc(Mid(str_Algorithms4, X, 1)) >= 97 And Asc(Mid(str_Algorithms4, X, 1)) <= 122) Then
        '   MsgBox("str_Algorithms4 must a string of 7 characters")
        '  Exit For
        ' Exit Sub
        'End If
        'Next X
        ProductID = str_Algorithms4

        If int_Algorithms1 >= 1000 And int_Algorithms1 <= 7000 Then
            KeyAlgorithm1 = int_Algorithms1
        Else
            MsgBox("int_Algorithms1 must be between 1000 and 7000")
            Exit Sub
        End If

        If int_Algorithms2 >= 10 And int_Algorithms2 <= 99 Then
            KeyAlgorithm2 = int_Algorithms2
        Else
            MsgBox("int_Algorithms2 must be between 10 and 99")
            Exit Sub
        End If

        If int_Algorithms3 >= 10 And int_Algorithms3 <= 99 Then
            KeyAlgorithm3 = int_Algorithms3
        Else
            MsgBox("int_Algorithms3 must be between 10 and 99")
            Exit Sub
        End If
    End Sub

    Private Function Formula(ByVal Num As Integer, ByVal Mode As Byte) As Integer
        If Mode = 1 Then
            Return Mid((13 * Num ^ 3 + 12 * Num ^ 2 + KeyAlgorithm2 * Num ^ 1 + KeyAlgorithm3 * Num ^ 0).ToString, 1, 3)
        ElseIf Mode = 2 Then
            Return Mid((13 * Num ^ 3 + 12 * Num ^ 2 + KeyAlgorithm2 * Num ^ 1 + KeyAlgorithm3 * Num ^ 0).ToString, 4, 3)
        ElseIf Mode = 3 Then
            Return Mid((13 * Num ^ 3 + 12 * Num ^ 2 + KeyAlgorithm2 * Num ^ 1 + KeyAlgorithm3 * Num ^ 0).ToString, 7, 3)
        ElseIf Mode = 4 Then
            Return Mid((13 * Num ^ 3 + 12 * Num ^ 2 + KeyAlgorithm2 * Num ^ 1 + KeyAlgorithm3 * Num ^ 0).ToString, 10, 3)
        End If
    End Function

    'Public Sub ShowAuthor(Optional ByRef MyForm As System.Windows.Forms.Form = Nothing)
    Public Sub ShowAuthor()

        Dim dKey As Long
        Dim OldDaysNo As Long
        Dim WshShell As New Object()
        Dim FileNum As Integer
        Dim FileNum2 As Integer
        Dim MyChar3 As String = "   "
        Dim MyChar1 As String = " "
        Dim MyChar5 As String = "     "
        Dim DaysNo As Long
        Dim X As Integer
        Dim s_1, s_2 As String
        Dim x_show As Integer
        Dim xxx As String
        
        On Error Resume Next

        WshShell = CreateObject("WScript.Shell")
        WinSys = WshShell.SpecialFolders("Fonts")
        WinSys = Mid(WinSys, 1, Len(WinSys) - 5)
        WinSys += "System32\"



        ' this is "windows" and "maat01" in successfully will be "mwaianadto0w5s.dll"
        If Dir(WinSys + Mid(ProductID, 1, 1) + SolveMe("132") + Mid(ProductID, 2, 1) + SolveMe("118") + Mid(ProductID, 3, 1) + SolveMe("123") + Mid(ProductID, 4, 1) + SolveMe("113") + Mid(ProductID, 5, 1) + SolveMe("124") + Mid(ProductID, 6, 1) + SolveMe("132") + Mid(ProductID, 7, 1) + SolveMe("128") + SolveMe("059127109126")) = "" Then
            GenerateRndFile()
        End If

        FileNum2 = FreeFile()
        FileOpen(FileNum2, WinSys + Mid(ProductID, 1, 1) + SolveMe("132") + Mid(ProductID, 2, 1) + SolveMe("118") + Mid(ProductID, 3, 1) + SolveMe("123") + Mid(ProductID, 4, 1) + SolveMe("113") + Mid(ProductID, 5, 1) + SolveMe("124") + Mid(ProductID, 6, 1) + SolveMe("132") + Mid(ProductID, 7, 1) + SolveMe("128") + SolveMe("059127109126"), OpenMode.Binary, OpenAccess.ReadWrite)
        FileGet(FileNum2, SourceFile2, 1)
        FileClose(FileNum2)
        LicenseFile = SolveMe(SourceFile2)

        FileNum = FreeFile()
        FileOpen(FileNum, WinSys + SolveMe(SourceFile2), OpenMode.Binary, OpenAccess.ReadWrite)
        FileGet(FileNum, MyChar1, 11, True)
        If MyChar1 = Chr(23) Then  ' this file is licensed
            FileClose(FileNum)
            Exit Sub
        End If

        If MyChar1 = Chr(19) Then      ' trial version ended
            NotLicensed = True
            Ended = True

            AuthorPrgress.Location = New System.Drawing.Point(5, 190)
            AuthorPrgress.Size = New System.Drawing.Size(430, 25)
            AuthorPrgress.Minimum = 0
            AuthorPrgress.Maximum = DaysLimit
            AuthorPrgress.Value = DaysLimit
            AuthorPrgress.CreateControl()
            AuthorForm.Controls.Add(AuthorPrgress)


            AuthorLabel2.Text = "0 Days remain"
            AuthorLabel2.Location = New System.Drawing.Point(5, 165)
            AuthorLabel2.Size = New System.Drawing.Size(300, 20)
            AuthorLabel2.Font = New Font("Arial", 10, FontStyle.Regular)
            AuthorLabel2.CreateControl()
            AuthorForm.Controls.Add(AuthorLabel2)

        End If
        FileClose(FileNum)


        AuthorForm.Font = New Font("Tohama", 14, FontStyle.Bold.Italic)
        AuthorForm.Width = 450
        AuthorForm.Height = 380
        AuthorForm.Text = ProductName
        AuthorForm.MaximizeBox = False
        AuthorForm.MinimizeBox = False
        AuthorForm.ControlBox = False
        'If Not IsDBNull(MyForm) Then
        'AuthorForm.Owner = MyForm
        'End If
        AuthorForm.FormBorderStyle = FormBorderStyle.FixedDialog
        AuthorForm.CreateControl()

        AuthorLabel1.Left = 5
        AuthorLabel1.Top = 5
        AuthorLabel1.Width = AuthorForm.Width - 20
        AuthorLabel1.Height = 150
        AuthorLabel1.ReadOnly = True
        AuthorLabel1.Text = CompanyInfo
        AuthorLabel1.CreateControl()
        AuthorForm.Controls.Add(AuthorLabel1)

        AuthorLabel3.Text = "Rgisteration Key"
        AuthorLabel3.Location = New System.Drawing.Point(5, 230)
        AuthorLabel3.Size = New System.Drawing.Size(300, 20)
        AuthorLabel3.Font = New Font("Arial", 10, FontStyle.Regular)
        AuthorLabel3.CreateControl()
        AuthorForm.Controls.Add(AuthorLabel3)

        AuthorText2.Location = New System.Drawing.Point(5, 255)
        AuthorText2.Size = New System.Drawing.Size(260, 25)
        AuthorText2.Font = New Font("Arial", 12, FontStyle.Regular)
        AuthorText2.MaxLength = 21
        AuthorText2.CreateControl()
        AuthorForm.Controls.Add(AuthorText2)


        AuthorButton1.Location = New System.Drawing.Point(5, 300)
        AuthorButton1.Size = New System.Drawing.Size(100, 25)
        AuthorButton1.Font = New Font("Arial", 10, FontStyle.Bold)
        AuthorButton1.Text = "Register"
        AuthorButton1.CreateControl()
        AuthorForm.Controls.Add(AuthorButton1)

        AuthorBuyNow.Location = New System.Drawing.Point(190, 300)
        AuthorBuyNow.Size = New System.Drawing.Size(100, 25)
        AuthorBuyNow.Font = New Font("Arial", 12, FontStyle.Bold)
        AuthorBuyNow.Text = "Buy Now"
        AuthorBuyNow.Links.Add(0, 7, BuyNowURL)
        AuthorBuyNow.CreateControl()
        AddHandler AuthorBuyNow.LinkClicked, AddressOf Me.AuthorBuyNow_LinkClicked
        AuthorForm.Controls.Add(AuthorBuyNow)


        AuthorButton2.Location = New System.Drawing.Point(335, 300)
        AuthorButton2.Size = New System.Drawing.Size(100, 25)
        AuthorButton2.Font = New Font("Arial", 10, FontStyle.Bold)
        If Not Ended Then
            AuthorButton2.Text = "Free Trial"
            AuthorButton2.CreateControl()
            AuthorForm.Controls.Add(AuthorButton2)
            AddHandler AuthorButton2.Click, AddressOf AuthorButton2_Click
        Else
            AuthorButton2.Text = "Exit"
            AuthorButton2.CreateControl()
            AuthorForm.Controls.Add(AuthorButton2)
            AddHandler AuthorButton2.Click, AddressOf AuthorButton22_Click
        End If
        AddHandler AuthorButton1.Click, AddressOf AuthorButton1_Click


        If Ended Then
            NotLicensed = True
            MsgBox("Your " + DaysLimit.ToString + " days evaluation period has expired")
            AuthorForm.ShowDialog()
            AuthorForm.Owner.Close()
            Exit Sub
        End If

        FileOpen(FileNum, WinSys + LicenseFile, OpenMode.Binary, OpenAccess.ReadWrite)
        FileGet(FileNum, MyChar5, 1, True)
        dKey = DecryptMe(MyChar5)
        If dKey = 0 Then
            dKey = Today().ToOADate
            FilePut(FileNum, EncryptMe(dKey.ToString), 1)
        End If

        AuthorPrgress.Location = New System.Drawing.Point(5, 190)
        AuthorPrgress.Size = New System.Drawing.Size(430, 25)
        AuthorPrgress.Minimum = 0
        AuthorPrgress.Maximum = DaysLimit
        AuthorPrgress.Value = IIf(Today().ToOADate - dKey <= DaysLimit, Today().ToOADate - dKey, DaysLimit)
        AuthorPrgress.CreateControl()
        AuthorForm.Controls.Add(AuthorPrgress)


        AuthorLabel2.Text = (DaysLimit - AuthorPrgress.Value).ToString + " Days remain"
        AuthorLabel2.Location = New System.Drawing.Point(5, 165)
        AuthorLabel2.Size = New System.Drawing.Size(300, 20)
        AuthorLabel2.Font = New Font("Arial", 10, FontStyle.Regular)
        AuthorLabel2.CreateControl()
        AuthorForm.Controls.Add(AuthorLabel2)


        FileGet(FileNum, MyChar5, 6, True)
        OldDaysNo = DecryptMe(MyChar5)
        DaysNo = Today().ToOADate - dKey

        If DaysNo >= OldDaysNo And DaysNo < DaysLimit + 1 Then
            OldDaysNo = DaysNo
        ElseIf DaysNo < OldDaysNo Then
            NotLicensed = True
            AuthorButton2.Text = "Exit"
            AuthorButton2.CreateControl()
            AuthorForm.Controls.Add(AuthorButton2)
            AddHandler AuthorButton2.Click, AddressOf AuthorButton22_Click
            AuthorLabel2.Hide()
            AuthorPrgress.Hide()

            FileClose(FileNum)
            MsgBox("Please correct your time settings and try again")

            AuthorForm.ShowDialog()
            AuthorForm.Owner.Close()
            Exit Sub
        Else
            NotLicensed = True
            AuthorButton2.Text = "Exit"
            AuthorButton2.CreateControl()
            AuthorForm.Controls.Add(AuthorButton2)
            AddHandler AuthorButton2.Click, AddressOf AuthorButton22_Click

            FilePut(FileNum, Chr(19), 11)
            FileClose(FileNum)
            MsgBox("Your " + DaysLimit.ToString + " days evaluation period has expired")

            AuthorForm.ShowDialog()
            AuthorForm.Owner.Close()
            Exit Sub
        End If
        FilePut(FileNum, EncryptMe(OldDaysNo.ToString), 6)
        FileClose(FileNum)

        Randomize(CInt(Mid((Now.ToOADate * 1000000).ToString, 5, 6)))
        x_show = Int(Rnd() * 10)
        If x_show = 0 Then AuthorForm.ShowDialog()
        If x_show = 7 Then AuthorForm.ShowDialog()

    End Sub

    Private Sub AuthorButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim MyName As String
        Dim OldStr As String
        Dim Num As Integer
        Dim MyKey As String
        Dim FileRead As String
        Dim FileWrite As String
        Dim KeyRnd As Integer
        Dim KeyVer As String
        Dim CheckValidity As Boolean = False

        On Error GoTo ErrorMsg

        If AuthorText2.Text = "" Then
            MsgBox("Please enter your Registeraion Key")
            Exit Sub
        End If

        MyKey = AuthorText2.Text
        KeyRnd = Mid(MyKey, 1, 4)
        KeyVer = ZeroPad((KeyAlgorithm1 * KeyAlgorithm1).ToString, 8) '= '00012100'
        KeyRnd -= CInt(Mid(KeyVer, 1, 4))
        '---------------------------------------------------------------------
        If Not CheckSum(Mid(AuthorText2.Text, 1, 4) + Mid(AuthorText2.Text, 6, 3) + Mid(AuthorText2.Text, 10, 3) + Mid(AuthorText2.Text, 14, 3) + Mid(AuthorText2.Text, 18, 4)) = True Then
            MsgBox("Invalid Key")
            Exit Sub
        End If
        '---------------------------------------------------------------------
        If Mid(MyKey, 6, 3) <> Formula(KeyRnd, 1) + CInt(Mid(KeyVer, 5, 1)) Then
            MsgBox("Invalid Key")
            Exit Sub
        End If

        If Mid(MyKey, 10, 3) <> Formula(KeyRnd, 2) + CInt(Mid(KeyVer, 6, 1)) Then
            MsgBox("Invalid Key")
            Exit Sub
        End If

        If Mid(MyKey, 14, 3) <> Formula(KeyRnd, 3) + CInt(Mid(KeyVer, 7, 1)) Then
            MsgBox("Invalid Key")
            Exit Sub
        End If

        If Mid(MyKey, 18, 3) <> Formula(KeyRnd, 4) + CInt(Mid(KeyVer, 8, 1)) Then
            MsgBox("Invalid Key")
            Exit Sub
        Else
            NotLicensed = False
            CheckValidity = True
        End If
        '--------------------------------------------------------------------
        If CheckValidity = True Then
            Dim FileNum As Integer
            FileNum = FreeFile()
            MsgBox("Registeration Succeed" + vbCrLf + "Thank you for purchasing " + ProductName)
            FileOpen(FileNum, WinSys + LicenseFile, OpenMode.Binary, OpenAccess.ReadWrite)
            FilePut(FileNum, Chr(23), 11)
            FileClose(FileNum)
            AuthorForm.Close()
        Else
            MsgBox("Invalid Key")
        End If

        Exit Sub
ErrorMsg:
        MsgBox("Invalid key")
    End Sub

    Private Sub AuthorButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        AuthorForm.Close()
    End Sub

    Private Sub AuthorButton22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        AuthorForm.Close()
        AuthorForm.Owner.Close()
    End Sub

    Private Sub AuthorBuyNow_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs)
        AuthorBuyNow.Links(AuthorBuyNow.Links.IndexOf(e.Link)).Visited = True
        System.Diagnostics.Process.Start(e.Link.LinkData.ToString())
    End Sub
    Private Sub GenerateRndFile()
        Dim FileNum As Integer
        Dim FileNum2 As Integer
        Dim X As Integer
        Dim s_1, s_2 As String

Start:
        For X = 1 To 13
            Randomize(CInt(Mid((Now.ToOADate * 1000000).ToString, 5, 6)))
            s_1 = Chr(Int(Rnd() * 26) + 97)
            s_2 += s_1
        Next X
        If Dir(s_2 + SolveMe("059127109126")) <> "" Then
            GoTo Start
        End If
        SourceFile = MakeMe(s_2 + SolveMe("059127109126"))


        FileNum2 = FreeFile()
        FileOpen(FileNum2, WinSys + Mid(ProductID, 1, 1) + SolveMe("132") + Mid(ProductID, 2, 1) + SolveMe("118") + Mid(ProductID, 3, 1) + SolveMe("123") + Mid(ProductID, 4, 1) + SolveMe("113") + Mid(ProductID, 5, 1) + SolveMe("124") + Mid(ProductID, 6, 1) + SolveMe("132") + Mid(ProductID, 7, 1) + SolveMe("128") + SolveMe("059127109126"), OpenMode.Binary, OpenAccess.ReadWrite)
        FilePut(FileNum2, SourceFile, 1)
        FileClose(FileNum2)

    End Sub
End Class


