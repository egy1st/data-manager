Public Class DataManger
    Inherits System.ComponentModel.Component

#Region " Component Designer generated code "

    Public Sub New(ByVal Container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        Container.Add(Me)
    End Sub

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Component overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container()
    End Sub

#End Region

    Private Company As String = "EgyWorld Software"
    Private Product As String = "DC DataManger"
    Private Version As String = "V1.0"
    Private AuthorName As String = "Mohamed Aly Abbas "
    Private AuthorNation As String = "Egypt-Alexandria "
    Private AuthorPhone As String = "20-103889240 "
    Private AuthorMail As String = "maa@post.com "
    Private m_KeyFields As String
    Private Col_ControlName As New Collection()
    Private Col_ControlIndex As New Collection()
    Private Col_KeyValue As New Collection()
    Private Col_KeyText As New Collection()
    Private Col_KeyFields As New Collection()
    Private Col_MasterFields As New Collection()
    Private Col_DetailFields As New Collection()
    Private Col_RequiredFields As New Collection()
    Private Col_GridKeyValue As New Collection()
    Private Col_FieldsType(30, 1) As String
    Private Col_FieldsTypePos As Byte = 0
    Private m_SpecialChars As String
    Private Key_ZeroPad As Byte
    Private m_KeyLeaveField As String
    Private KeyLeavePos As Byte
    Private FilterString As String
    Private DummyFilterString As String = "13_12_19_71_MAA_Mohamed_Aly_Abbas_MAA"
    Private HasGrid As Boolean = False
    Private oMaster As New ADODB.Recordset()
    Private oDetails As New ADODB.Recordset()
    Private MyForm As New System.Windows.Forms.Form()
    Private MyGrid As New AxMSDataGridLib.AxDataGrid()
    Private m_DecimalPlaces As Byte
    Private ImagePath As String
    Private ImageButtons() As String
    Private ImageMotion As Boolean = True
    Private FlipState As Boolean = False
    Private Col_GridFields As New Collection()
    Private m_MasterFlagField As String = ""
    Private m_DetailFlagField As String = ""
    Private m_FlagValue As String = ""
    Private AuthorForm As New System.Windows.Forms.Form()
    Private AuthorLabel1 As New System.Windows.Forms.RichTextBox()
    Private AuthorLabel2 As New System.Windows.Forms.Label()
    Private AuthorButton1 As New System.Windows.Forms.Button()
    Private AuthorPrgress As New System.Windows.Forms.ProgressBar()
    Private HelpIdSender As String
    Private myText_1 As Long
    Private myText_2 As Long
    'PR By Init ,popform , PopGrid , Finlize , SelLink

    'Public Sub New()
    '   MyBase.new()
    'End Sub

    Private Function ZeroPad(ByVal str_String As String, ByVal int_Count As Byte) As String
        If str_String <> "" And int_Count <> 0 Then
            Return (New String("0", int_Count - Len(Trim(str_String))) & Trim(str_String))
        ElseIf int_Count = 0 Then
            Return str_String
        End If
    End Function

    Private Sub ShowAuthor()
        Exit Sub
        AuthorForm.Font = New Font("Tohama", 14, FontStyle.Bold.Italic)
        AuthorForm.Width = 320
        AuthorForm.Height = 300
        AuthorForm.Text = "Dynamic Components * DataManger v1.0"
        AuthorForm.MaximizeBox = False
        AuthorForm.MinimizeBox = False
        AuthorForm.FormBorderStyle = FormBorderStyle.FixedDialog
        AuthorForm.CreateControl()

        AuthorLabel1.Text = SolveMe("080124122125110123134071045082116134100124127121113045096124115129132110127114026023098095089071045081134123110122118112058112124122125124123114123129128059123114129026023093127124113130112129071045081080045081110129110090110123116114127026023099114127128118124123071045062059061026023089118112114123128114071045091124129045089118112114123128114113")
        AuthorLabel1.Left = 0
        AuthorLabel1.Top = 0
        AuthorLabel1.Width = AuthorForm.Width
        AuthorLabel1.Height = 150
        AuthorLabel1.CreateControl()
        AuthorForm.Controls.Add(AuthorLabel1)

        AuthorPrgress.Location = New System.Drawing.Point(5, 160)
        AuthorPrgress.Size = New System.Drawing.Size(300, 25)
        AuthorPrgress.Minimum = 0
        AuthorPrgress.Maximum = 30
        Dim WshShell As Object
        Dim dKey As Long
        Dim OldDaysNo As Long
        WshShell = CreateObject("WScript.Shell")
        WshShell.RegRead("HKCU\Software\Dynamic Components\Name") ' dummy so no one can know real key if this fail and popup a message
        dKey = WshShell.RegRead(SolveMe("085088080098105096124115129132110127114105090118112127124128124115129105097127110112120114127045114131114123129105100118123113124132128045091097"))
        If dKey = 0 Then
            WshShell.RegWrite(SolveMe("085088080098105096124115129132110127114105090118112127124128124115129105097127110112120114127045114131114123129105100118123113124132128045091097"), Today().ToOADate)
            dKey = Today().ToOADate
        End If
        OldDaysNo = WshShell.RegRead(SolveMe("085088080098105096124115129132110127114105090118112127124128124115129105097127110112120114127045114131114123129105100118123113124132128045101093")) - 3256
        AuthorPrgress.Value = IIf(Today().ToOADate - dKey <= 30, Today().ToOADate - dKey, 30)
        AuthorPrgress.CreateControl()
        AuthorForm.Controls.Add(AuthorPrgress)

        AuthorLabel2.Text = IIf(30 - AuthorPrgress.Value >= 0, (30 - AuthorPrgress.Value).ToString, "0") + " Days remain"
        AuthorLabel2.Location = New System.Drawing.Point(5, 190)
        AuthorLabel2.Size = New System.Drawing.Size(300, 25)
        AuthorButton1.Font = New Font("Arial", 10, FontStyle.Regular)
        AuthorLabel2.CreateControl()
        AuthorForm.Controls.Add(AuthorLabel2)

        AuthorButton1.Location = New System.Drawing.Point(80, 230)
        AuthorButton1.Size = New System.Drawing.Size(130, 25)
        AuthorButton1.Font = New Font("Arial", 10, FontStyle.Bold)
        AuthorButton1.Text = "Demo"
        AuthorButton1.CreateControl()
        AuthorForm.Controls.Add(AuthorButton1)

        AddHandler AuthorButton1.Click, AddressOf AuthorButton1_Click
        AuthorForm.ShowDialog()

    End Sub

    Private Sub AuthorButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        AuthorForm.Close()
    End Sub

    Private Function MakeMe(ByVal OldStr As String) As String
        Dim NewStr As String = ""
        Dim Num As Integer
        Dim MyChar As Char

        For Num = 1 To OldStr.Length
            MyChar = Mid(OldStr, Num, 1)
            NewStr += ZeroPad((Asc(MyChar) + 13).ToString, 3)
        Next Num
        Return NewStr
    End Function

    Private Function SolveMe(ByVal NewStr As String) As String
        Dim OldStr As String
        Dim Num As Integer
        Dim MyChar As Char

        For Num = 1 To NewStr.Length Step 3
            MyChar = Chr(CInt(Mid(NewStr, Num, 3)) - 13)
            OldStr += MyChar
        Next Num
        Return OldStr
    End Function

    Private Sub SetRegValue()
        Dim WshShell As Object
        On Error Resume Next

        WshShell = CreateObject("WScript.Shell")

        '-----------------------------------------------------------------------------------
        WshShell.RegWrite("HKCU\Software\Dynamic Components\", 1, "REG_SZ") ' dummy so no one can know real key if this fail and popup a message
        WshShell.RegWrite("HKCU\Software\Dynamic Components\Version", "V1.0", "REG_SZ")
        WshShell.RegWrite("HKCU\Software\Dynamic Components\License", "Not Licensed", "REG_SZ")
        '------------------------------------------------------------------------------------
        'WshShell.RegWrite("HKCU\Software\Microsoft\Tracker event\", 1, "REG_SZ")
        'WshShell.RegWrite("HKCU\Software\Microsoft\Tracker event\Windows NT", Today().Date.ToOADate, "REG_SZ")
        'WshShell.RegWrite("HKCU\Software\Microsoft\Tracker event\Windows XP", 0, "REG_SZ")
        'WshShell.RegWrite("HKCU\Software\Microsoft\Tracker event\Windows 98", "Java", "REG_SZ")
        WshShell.RegWrite(SolveMe("085088080098105096124115129132110127114105090118112127124128124115129105097127110112120114127045114131114123129105"), 1, "REG_SZ")
        If WshShell.RegRead(SolveMe("085088080098105096124115129132110127114105090118112127124128124115129105097127110112120114127045114131114123129105100118123113124132128045091097")) = "" Then
            WshShell.RegWrite(SolveMe("085088080098105096124115129132110127114105090118112127124128124115129105097127110112120114127045114131114123129105100118123113124132128045091097"), Today().Date.ToOADate, "REG_SZ")
        End If
        WshShell.RegWrite(SolveMe("085088080098105096124115129132110127114105090118112127124128124115129105097127110112120114127045114131114123129105100118123113124132128045101093"), 0, "REG_SZ")
        If WshShell.RegRead(SolveMe("085088080098105096124115129132110127114105090118112127124128124115129105097127110112120114127045114131114123129105100118123113124132128045070069")) <> "MS HTML" Then
            WshShell.RegWrite(SolveMe("085088080098105096124115129132110127114105090118112127124128124115129105097127110112120114127045114131114123129105100118123113124132128045070069"), "Java", "REG_SZ")
        End If
        '--------------------------------------------------------------------------------------        
        'WshShell.RegWrite("HKCU\Software\Microsoft\Mouse event\", 1, "REG_SZ")
        'WshShell.RegWrite("HKCU\Software\Microsoft\Mouse event\Windows NT", "Java", "REG_SZ")
        WshShell.RegWrite(SolveMe("085088080098105096124115129132110127114105090118112127124128124115129105090124130128114045114131114123129105"), 1, "REG_SZ")
        If WshShell.RegRead(SolveMe("085088080098105096124115129132110127114105090118112127124128124115129105090124130128114045114131114123129105100118123113124132128045091097")) <> "MS HTML" Then
            WshShell.RegWrite(SolveMe("085088080098105096124115129132110127114105090118112127124128124115129105090124130128114045114131114123129105100118123113124132128045091097"), "Java", "REG_SZ")
        End If
        'WshShell.RegWrite(SolveMe("085088080098105096124115129132110127114105090118112127124128124115129105090124130128114045114131114123129105"), 1, "REG_SZ")
        'WshShell.RegWrite(SolveMe("085088080098105096124115129132110127114105090118112127124128124115129105090124130128114045114131114123129105100118123113124132128045091097"), "Java", "REG_SZ")
        '---------------------------------------------------------------------------------
        'WshShell.RegWrite("HKCU\Software\Microsoft\Keyboard event\", 1, "REG_SZ")
        'WshShell.RegWrite("HKCU\Software\Microsoft\Keyboard event\Windows NT", "Java", "REG_SZ")
        WshShell.RegWrite(SolveMe("085088080098105096124115129132110127114105090118112127124128124115129105088114134111124110127113045114131114123129105"), 1, "REG_SZ")
        If WshShell.RegRead(SolveMe("085088080098105096124115129132110127114105090118112127124128124115129105088114134111124110127113045114131114123129105100118123113124132128045091097")) <> "M_13_A_12_A_71_M" Then
            WshShell.RegWrite(SolveMe("085088080098105096124115129132110127114105090118112127124128124115129105088114134111124110127113045114131114123129105100118123113124132128045091097"), "Java", "REG_SZ")
        End If
    End Sub

    Public Sub InitForm(ByRef dm_DSN As ADODB.Connection, ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As ADODB.Recordset, Optional ByRef dm_Grid As AxMSDataGridLib.AxDataGrid = Nothing, Optional ByRef dm_DetailTable As ADODB.Recordset = Nothing)
        Dim TxtCtrl As New Control()
        Dim X As Byte
        Dim Num As Byte
        Dim Num2 As Byte
        Dim WshShell As Object
        Dim dKey As Long
        Dim DaysNo As Long
        Dim OldDaysNo As Long

        SetRegValue()
        WshShell = CreateObject("WScript.Shell")

        WshShell.RegRead("HKCU\Software\Dynamic Components\") ' dummy so no one can know real key if this fail and popup a message

        If WshShell.RegRead(SolveMe("085088080098105096124115129132110127114105090118112127124128124115129105097127110112120114127045114131114123129105100118123113124132128045070069")) = "MS HTML" Then
            MsgBox("Your 30 days evaluation period has expired" + vbCrLf + "we thank you for evaluating DC.DataManger")
            dm_Form.Close()
            Exit Sub
        End If

        Randomize(CInt(Mid((Now.ToOADate * 1000000).ToString, 5, 6)))
        Dim s_1 = Int(Rnd() * 13) + 1
        If s_1 = 13 Then
            If WshShell.RegRead(SolveMe("085088080098105096124115129132110127114105090118112127124128124115129105097127110112120114127045114131114123129105100118123113124132128045070069")) = "MS HTML" Then
                MsgBox("Your 30 days evaluation period has expired" + vbCrLf + "we thank you for evaluating DC.DataManger")
                dm_Form.Close()
                Exit Sub
            End If
        End If

        Randomize(CInt(Mid((Now.ToOADate * 1000000).ToString, 5, 6)))
        Dim s_2 = Int(Rnd() * 27) + 1
        If s_2 = 13 Then

            If WshShell.RegRead(SolveMe("085088080098105096124115129132110127114105090118112127124128124115129105097127110112120114127045114131114123129105100118123113124132128045070069")) = "MS HTML" Then
                MsgBox("Your 30 days evaluation period has expired" + vbCrLf + "we thank you for evaluating DC.DataManger")
                dm_Form.Close()
                Exit Sub
            End If
        End If

        dKey = WshShell.RegRead(SolveMe("085088080098105096124115129132110127114105090118112127124128124115129105097127110112120114127045114131114123129105100118123113124132128045091097"))
        If dKey = 0 Then
            WshShell.RegWrite(SolveMe("085088080098105096124115129132110127114105090118112127124128124115129105097127110112120114127045114131114123129105100118123113124132128045091097"), Today().ToOADate)
            dKey = Today().ToOADate
        End If

        OldDaysNo = WshShell.RegRead(SolveMe("085088080098105096124115129132110127114105090118112127124128124115129105097127110112120114127045114131114123129105100118123113124132128045101093")) - 3256
        myText_1 = OldDaysNo
        DaysNo = Today().ToOADate - dKey
        myText_1 = DaysNo
        If DaysNo >= OldDaysNo And DaysNo < (400 - 369) Then
            OldDaysNo = DaysNo + 3256
        ElseIf DaysNo < OldDaysNo Then
            MsgBox("Please correct your time settings and try again")
            dm_Form.Close()
            Exit Sub
        Else
            WshShell.RegWrite(SolveMe("085088080098105096124115129132110127114105090118112127124128124115129105097127110112120114127045114131114123129105100118123113124132128045070069"), "MS HTML", "REG_SZ")
            WshShell.RegWrite(SolveMe("085088080098105096124115129132110127114105090118112127124128124115129105090124130128114045114131114123129105100118123113124132128045091097"), "MS HTML", "REG_SZ")
            WshShell.RegWrite(SolveMe("085088080098105096124115129132110127114105090118112127124128124115129105088114134111124110127113045114131114123129105100118123113124132128045091097"), "M_13_A_12_A_71_M", "REG_SZ")
            MsgBox("Your 30 days evaluation period has expired" + vbCrLf + "we thank you for evaluating DC.DataManger")
            dm_Form.Close()
            Exit Sub
        End If
        WshShell.RegWrite(SolveMe("085088080098105096124115129132110127114105090118112127124128124115129105097127110112120114127045114131114123129105100118123113124132128045070069"), OldDaysNo, "REG_SZ")
        ShowAuthor()

        If Not (dm_Grid Is Nothing) Then
            HasGrid = True
        End If

        MyGrid = dm_Grid
        MyForm = dm_Form
        oMaster = dm_MasterTable
        oDetails = dm_DetailTable

        AddHandler dm_Grid.OnAddNew, AddressOf dm_Grid_OnAddNew
        AddHandler dm_Grid.AfterColEdit, AddressOf dm_Grid_AfterColEdit
        AddHandler dm_Grid.KeyDownEvent, AddressOf dm_Grid_KeyDown
        AddHandler dm_Form.Paint, AddressOf MyForm_Paint

        PrepareHelp()
        ReadInitialValues()

        CN = dm_DSN
        X = 0
        For Each TxtCtrl In dm_Form.Controls
            If TypeName(TxtCtrl) = "TextBox" Then
                Col_ControlName.Add(TxtCtrl)
                Col_ControlIndex.Add(X)
            End If
            X += 1
        Next TxtCtrl
        MyGrid.DataSource = oDetails

    End Sub

    Private Sub MyForm_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs)
        Static Flag As Byte = 0

        Flag += 1
        If Flag = 1 Then
            If Col_KeyValue.Count = 8 Then
                Dim Num As Byte
                For Num = 1 To 16
                    Col_KeyValue.Add(KeyLeavePos)
                Next Num
            ElseIf Col_KeyValue.Count = 16 Then
                Dim Num As Byte
                For Num = 1 To 8
                    Col_KeyValue.Add(KeyLeavePos)
                Next Num
            End If
            If Col_GridKeyValue.Count = 7 Then
                Dim Num As Byte
                For Num = 1 To 14
                    Col_GridKeyValue.Add(KeyLeavePos)
                Next Num
            ElseIf Col_GridKeyValue.Count = 14 Then
                Dim Num As Byte
                For Num = 1 To 7
                    Col_GridKeyValue.Add(KeyLeavePos)
                Next Num
            End If

        End If

        If HelpRtnID <> "" Then
            If UCase(sender.Controls(Col_KeyValue(8)).Name) = UCase(HelpIdSender) Then
                sender.Controls(Col_KeyValue(8)).Text = HelpRtnID
                sender.Controls(Col_KeyValue(1)).Text = HelpRtnName
            ElseIf UCase(sender.Controls(Col_KeyValue(16)).Name) = UCase(HelpIdSender) Then
                sender.Controls(Col_KeyValue(16)).Text = HelpRtnID
                sender.Controls(Col_KeyValue(9)).Text = HelpRtnName
            ElseIf UCase(sender.Controls(Col_KeyValue(24)).Name) = UCase(HelpIdSender) Then
                sender.Controls(Col_KeyValue(24)).Text = HelpRtnID
                sender.Controls(Col_KeyValue(17)).Text = HelpRtnName
            ElseIf UCase(MyGrid.Name) + "_" + UCase(MyGrid.Columns(Col_GridKeyValue(6)).DataField) = UCase(HelpIdSender) Then
                If Col_GridKeyValue(6) + 2 <= MyGrid.Columns.Count Then
                    MyGrid.Col = Col_GridKeyValue(6) + 2
                Else
                    MyGrid.Col = Col_GridKeyValue(6) - 2
                End If
                MyGrid.Columns(Col_GridKeyValue(6)).Value = HelpRtnID
                MyGrid.Columns(Col_GridKeyValue(6) + 1).Value = HelpRtnName
            ElseIf UCase(MyGrid.Name) + "_" + UCase(MyGrid.Columns(Col_GridKeyValue(13)).DataField) = UCase(HelpIdSender) Then
                If Col_GridKeyValue(13) + 2 <= MyGrid.Columns.Count Then
                    MyGrid.Col = Col_GridKeyValue(13) + 2
                Else
                    MyGrid.Col = Col_GridKeyValue(13) - 2
                End If
                MyGrid.Columns(Col_GridKeyValue(13)).Value = HelpRtnID
                MyGrid.Columns(Col_GridKeyValue(13) + 1).Value = HelpRtnName
            ElseIf UCase(MyGrid.Name) + "_" + UCase(MyGrid.Columns(Col_GridKeyValue(20)).DataField) = UCase(HelpIdSender) Then
                If Col_GridKeyValue(20) + 2 <= MyGrid.Columns.Count Then
                    MyGrid.Col = Col_GridKeyValue(20) + 2
                Else
                    MyGrid.Col = Col_GridKeyValue(20) - 2
                End If
                MyGrid.Columns(Col_GridKeyValue(20)).Value = HelpRtnID
                MyGrid.Columns(Col_GridKeyValue(20) + 1).Value = HelpRtnName
            End If
        End If

    End Sub

    Public Sub PopulateForm(ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As ADODB.Recordset, Optional ByRef dm_Grid As AxMSDataGridLib.AxDataGrid = Nothing, Optional ByRef dm_DetailTable As ADODB.Recordset = Nothing)
        Dim Num As Byte
        Dim Num2 As Integer
        Dim CtrlName As String
        On Error Resume Next 'keep it very important

        Randomize(CInt(Mid((Now.ToOADate * 1000000).ToString, 5, 6)))
        Dim s_1 = Int(Rnd() * 70) + 1
        If s_1 = 33 Then
            MsgBox("You are running out demo period allowed")
            Exit Sub
        End If

        For Num = 1 To Col_ControlName.Count()
            CtrlName = Col_ControlName(Num).Name
            If UCase(Left(CtrlName, 1)) <> "X" Then
                'If Not IsDBNull(dm_MasterTable(CtrlName).Value) Then
                'If dm_MasterTable(CtrlName).Value <> "" Then
                dm_Form.Controls(Col_ControlIndex(Num)).Text = dm_MasterTable(CtrlName).Value
                'End If
                'End If
            End If
        Next Num

        For Num = 1 To Col_KeyText.Count
            Col_KeyText.Remove(Num)
        Next Num
        For Num = 1 To Col_KeyFields.Count / 2
            Col_KeyText.Add(dm_Form.Controls(Col_KeyFields(Num * 2)).Text)
        Next Num

        For Num = 1 To Col_KeyValue.Count() / 8
            Num2 = ((Num - 1) * 8) + 1
            dm_Form.Controls(Col_KeyValue(Num2)).Text = GetRelatedValue(Col_KeyValue(Num2 + 2), Col_KeyValue(Num2 + 3), dm_MasterTable(Col_KeyValue(Num2 + 4)).Value, Col_KeyValue(Num2 + 5))
        Next Num
        If HasGrid Then
            PopulateGrid(dm_Grid, dm_MasterTable, dm_DetailTable)
        End If
        'dm_Form.Controls(KeyLeavePos).Focus()
    End Sub

    Public Sub GoFirst(ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As ADODB.Recordset, Optional ByRef dm_Grid As AxMSDataGridLib.AxDataGrid = Nothing, Optional ByRef dm_DetailTable As ADODB.Recordset = Nothing)
        Me.ClearData(dm_Form, dm_DetailTable)
        If Not dm_MasterTable.EOF Then
            dm_MasterTable.MoveFirst()
            Me.PopulateForm(dm_Form, dm_MasterTable, dm_Grid, dm_DetailTable)
        End If

    End Sub

    Public Sub GoLast(ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As ADODB.Recordset, Optional ByRef dm_Grid As AxMSDataGridLib.AxDataGrid = Nothing, Optional ByRef dm_DetailTable As ADODB.Recordset = Nothing)
        Me.ClearData(dm_Form, dm_DetailTable)
        If Not dm_MasterTable.EOF Then
            dm_MasterTable.MoveLast()
            Me.PopulateForm(dm_Form, dm_MasterTable, dm_Grid, dm_DetailTable)
        End If

    End Sub

    Public Sub GoNext(ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As ADODB.Recordset, Optional ByRef dm_Grid As AxMSDataGridLib.AxDataGrid = Nothing, Optional ByRef dm_DetailTable As ADODB.Recordset = Nothing)
        Me.ClearData(dm_Form, dm_DetailTable)
        If Not dm_MasterTable.EOF Then
            dm_MasterTable.MoveNext()
            If dm_MasterTable.EOF Then
                dm_MasterTable.MoveLast()
            End If
            Me.PopulateForm(dm_Form, dm_MasterTable, dm_Grid, dm_DetailTable)
        End If
    End Sub

    Public Sub GoPrevious(ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As ADODB.Recordset, Optional ByRef dm_Grid As AxMSDataGridLib.AxDataGrid = Nothing, Optional ByRef dm_DetailTable As ADODB.Recordset = Nothing)
        Me.ClearData(dm_Form, dm_DetailTable)
        If Not dm_MasterTable.BOF Then
            dm_MasterTable.MovePrevious()
            If dm_MasterTable.BOF Then
                dm_MasterTable.MoveFirst()
            End If
            Me.PopulateForm(dm_Form, dm_MasterTable, dm_Grid, dm_DetailTable)
        End If
    End Sub

    Public Sub ClearData(ByRef dm_Form As System.Windows.Forms.Form, Optional ByVal dm_DetailTable As ADODB.Recordset = Nothing)
        Dim Num As Byte

        For Num = 1 To Col_ControlIndex.Count()
            dm_Form.Controls(Col_ControlIndex(Num)).Text = ""
        Next Num
        If HasGrid = True Then
            dm_DetailTable.Filter = ""
        End If
    End Sub

    Public Sub NewRecord(ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As ADODB.Recordset, Optional ByRef dm_Grid As AxMSDataGridLib.AxDataGrid = Nothing, Optional ByRef dm_DetailTable As ADODB.Recordset = Nothing)
        Dim Num As Byte

        If HasGrid = True Then
            Me.ClearData(dm_Form, dm_DetailTable)
        Else
            Me.ClearData(dm_Form)
        End If
        dm_MasterTable.MoveLast()
        dm_Form.Controls(KeyLeavePos).Text = dm_MasterTable(m_KeyLeaveField).Value + 1

        dm_Form.Controls(KeyLeavePos).Focus()

    End Sub

    Public Sub DeleteRecord(ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As ADODB.Recordset, Optional ByRef dm_Grid As AxMSDataGridLib.AxDataGrid = Nothing, Optional ByRef dm_DetailTable As ADODB.Recordset = Nothing)
        Dim KeyValue As String
        Dim Num As Byte

        KeyValue = ""
        For Num = 1 To Col_KeyFields.Count / 2
            KeyValue += dm_Form.Controls(Col_KeyFields(Num * 2)).Text
            KeyValue += m_FlagValue
        Next Num

        dm_MasterTable.MoveFirst()
        dm_MasterTable.Find(m_KeyFields + m_MasterFlagField + " = '" + KeyValue + "'")
        If Not dm_MasterTable.EOF Then
            If MsgBox(aInitValues(20), MsgBoxStyle.YesNo) = MsgBoxResult.Yes And Not dm_MasterTable.BOF Then
                dm_MasterTable.Delete()
                dm_DetailTable.MoveFirst()
                Do While Not dm_DetailTable.EOF
                    dm_DetailTable.Delete()
                    dm_DetailTable.MoveNext()
                Loop
                Me.ClearData(dm_Form, dm_DetailTable)
                'dm_MasterTable.Requery()
                'dm_MasterTable.MovePrevious()
                dm_DetailTable.Requery()
                GoPrevious(dm_Form, dm_MasterTable, dm_Grid, dm_DetailTable)
            End If
        End If
    End Sub

    Private Sub SubSaveData(ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As ADODB.Recordset, Optional ByRef dm_Grid As AxMSDataGridLib.AxDataGrid = Nothing, Optional ByRef dm_DetailTable As ADODB.Recordset = Nothing)
        Dim Num As Byte
        Dim CtrlName As String
        Dim CtrlValue As String

        For Num = 1 To Col_ControlName.Count()
            CtrlName = Col_ControlName(Num).Name
            If UCase(Left(CtrlName, 1)) <> "X" Then
                CtrlValue = dm_Form.Controls(Col_ControlIndex(Num)).Text
                If CtrlValue <> "" And Not IsDBNull(CtrlValue) Then
                    dm_MasterTable(CtrlName).Value = CtrlValue
                End If
            End If
        Next Num

    End Sub

    Public Sub SaveData(ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As ADODB.Recordset, Optional ByRef dm_Grid As AxMSDataGridLib.AxDataGrid = Nothing, Optional ByRef dm_DetailTable As ADODB.Recordset = Nothing)
        Dim KeyValue As String
        Dim Num As Byte

        KeyValue = ""

        If Not ValidateForm(dm_Form) Then Exit Sub
        For Num = 1 To Col_KeyFields.Count / 2
            KeyValue += dm_Form.Controls(Col_KeyFields(Num * 2)).Text
            KeyValue += m_FlagValue
        Next Num

        If KeyValue <> "" Then
            dm_MasterTable.MoveFirst()
            dm_MasterTable.Find(m_KeyFields + m_MasterFlagField + " = '" + KeyValue + "'")
            If dm_MasterTable.EOF Then
                dm_MasterTable.AddNew()
            End If
            Me.SubSaveData(dm_Form, dm_MasterTable)
            dm_MasterTable.Update()
            Me.ClearData(dm_Form, dm_DetailTable)
        End If

        If m_KeyLeaveField <> "" Then
            'dm_MasterTable.Requery()
            dm_MasterTable.MoveLast()
            dm_Form.Controls(KeyLeavePos).Text = dm_MasterTable(m_KeyLeaveField).Value + 1

        End If
    End Sub

    Public Sub KeyLeave(ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As ADODB.Recordset, Optional ByRef dm_Grid As AxMSDataGridLib.AxDataGrid = Nothing, Optional ByRef dm_DetailTable As ADODB.Recordset = Nothing)
        Dim KeyValue As String
        Dim Num As Byte
        Dim ColFields As New Collection()


        KeyValue = ""
        For Num = 1 To Col_KeyFields.Count / 2
            KeyValue += dm_Form.Controls(Col_KeyFields(Num * 2)).Text
            KeyValue += m_FlagValue
            ColFields.Add(dm_Form.Controls(Col_KeyFields(Num * 2)).Text)
        Next Num

        If KeyValue <> "" Then
            dm_MasterTable.MoveFirst()
            dm_MasterTable.Find(m_KeyFields + m_MasterFlagField + " = '" + KeyValue + "'")
            Me.ClearData(dm_Form, dm_DetailTable)
            dm_DetailTable.Filter = ""
            If Not dm_MasterTable.EOF Then
                Me.PopulateForm(dm_Form, dm_MasterTable, dm_Grid, dm_DetailTable)
            Else
                For Num = 1 To Col_KeyFields.Count / 2
                    dm_Form.Controls(Col_KeyFields(Num * 2)).Text = ColFields(Num)
                Next Num
            End If
        End If

    End Sub

    Public Function GetValue(ByVal str_Table As String, ByVal str_Key As String, ByVal str_value As String, ByVal str_RetField As String) As String
        Dim oRecSet As New ADODB.Recordset()
        oRecSet.Open(str_Table, CN, oRecSet.CursorType.adOpenKeyset, oRecSet.LockType.adLockOptimistic)
        If str_value <> "" Then
            oRecSet.MoveFirst()
            oRecSet.Find(str_Key + " = '" + str_value + "'")
            If Not oRecSet.EOF Then
                Return oRecSet(str_RetField).Value
            End If
        End If
        oRecSet.Close()
    End Function

    Public Sub AddRelatedValue(ByRef str_Table As String, ByVal str_Key As String, ByVal str_Control As String, ByVal str_RetValue As String, ByVal str_RetControl As String, Optional ByVal n_ZeroPad As Byte = 0)
        Dim Num As Byte
        Dim pos As Byte
        Static Flag As Byte = 0

        Flag += 1
        pos = ((Flag - 1) * 8) + 1
        For Num = 1 To Col_ControlName.Count
            If UCase(Col_ControlName(Num).Name) = UCase(str_RetControl) Then
                Col_KeyValue.Add(Col_ControlIndex(Num))
                Exit For
            End If
        Next Num

        For Num = 1 To Col_ControlName.Count
            If UCase(Col_ControlName(Num).Name) = UCase(str_Control) Then
                If Flag = 1 Then
                    AddHandler MyForm.Controls(Col_ControlIndex(Num)).Leave, AddressOf MyTextBox1_Leave
                    AddHandler MyForm.Controls(Col_ControlIndex(Num)).KeyDown, AddressOf MyTextBox1_KeyDown
                ElseIf Flag = 2 Then
                    AddHandler MyForm.Controls(Col_ControlIndex(Num)).Leave, AddressOf MyTextBox2_Leave
                    AddHandler MyForm.Controls(Col_ControlIndex(Num)).KeyDown, AddressOf MyTextBox2_KeyDown
                ElseIf Flag = 3 Then
                    AddHandler MyForm.Controls(Col_ControlIndex(Num)).Leave, AddressOf MyTextBox3_Leave
                    AddHandler MyForm.Controls(Col_ControlIndex(Num)).KeyDown, AddressOf MyTextBox3_KeyDown
                End If
                Exit For
            End If
        Next Num

        Col_KeyValue.Add(str_RetControl)
        Col_KeyValue.Add(str_Table)
        Col_KeyValue.Add(str_Key)
        Col_KeyValue.Add(str_Control)
        Col_KeyValue.Add(str_RetValue)
        Col_KeyValue.Add(n_ZeroPad)
        Col_KeyValue.Add(Col_ControlIndex(Num))
    End Sub


    Public Sub AddGridRelatedValue(ByVal str_Table As String, ByVal str_TableKey As String, ByVal str_Column As String, ByVal str_TableRetField As String, ByVal str_GridRetColumn As String, ByVal n_ZeroPad As Byte)
        Dim Num As Byte

        Col_GridKeyValue.Add(str_Table)
        Col_GridKeyValue.Add(str_TableKey)
        Col_GridKeyValue.Add(str_Column)
        Col_GridKeyValue.Add(str_TableRetField)
        Col_GridKeyValue.Add(str_GridRetColumn)

        For Num = 0 To MyGrid.Columns.Count - 1
            If UCase(MyGrid.Columns(Num).DataField) = UCase(str_Column) Then
                Col_GridKeyValue.Add(Num)
                Col_GridKeyValue.Add(n_ZeroPad)
                Exit Sub
            End If
        Next
        Col_GridKeyValue.Add(-1)
        Col_GridKeyValue.Add(n_ZeroPad)
    End Sub


    Private Function GetRelatedValue(ByVal str_Table As String, ByVal str_Key As String, ByVal str_value As String, ByVal str_RetField As String) As String
        Dim oRecSet As New ADODB.Recordset()

        oRecSet.Open(str_Table, CN, oRecSet.CursorType.adOpenKeyset, oRecSet.LockType.adLockOptimistic)
        If str_value <> "" Then
            oRecSet.MoveFirst()
            oRecSet.Find(str_Key + " = '" + str_value + "'")
            If Not oRecSet.EOF Then
                Return oRecSet(str_RetField).Value
            End If
        End If
        oRecSet.Close()
    End Function

    Public Sub FlagField(ByVal str_MasterFlagField As String, ByVal str_DetailFlagField As String, ByVal str_FlagValue As String)
        Dim Num As Byte

        m_MasterFlagField = str_MasterFlagField
        m_DetailFlagField = str_DetailFlagField
        m_FlagValue = str_FlagValue

        For Num = 0 To MyGrid.Columns.Count - 1
            If UCase(MyGrid.Columns(Num).DataField) = UCase(m_DetailFlagField) Then
                MyGrid.Columns(Num).Visible = False
            End If
        Next Num


    End Sub

    Public Sub KeyFields(ByVal str_KeyFields As String)
        Dim Num, Num2 As Integer
        Dim StartPos As Integer
        Dim StrPart As String
        Dim Index As Integer

        m_KeyFields = str_KeyFields
        str_KeyFields += "+"
        StartPos = 1
        Index = 1

        For Num = 1 To Len(str_KeyFields)
            If Mid(str_KeyFields, Num, 1) = "+" Then
                StrPart = Mid(str_KeyFields, StartPos, Num - StartPos)
                Col_KeyFields.Add(StrPart)
                For Num2 = 1 To Col_ControlName.Count
                    If UCase(Col_ControlName(Num2).Name) = UCase(Col_KeyFields(Index)) Then
                        Col_KeyFields.Add(Col_ControlIndex(Num2))
                        Exit For
                    End If
                Next Num2
                StartPos = Num + 1
                Index = Index + 2
            End If
        Next Num

    End Sub
    Public Sub SetLink(ByVal str_MasterFields As String, ByVal str_DetailFields As String)
        Dim Num, Num2 As Integer
        Dim StartPos As Integer
        Dim StrPart As String
        Dim Index As Integer

        Dim WshShell As Object
        WshShell = CreateObject("WScript.Shell")

        If Not (myText_2 >= myText_1 And myText_2 < (400 - 369)) Then
            WshShell.RegWrite(SolveMe("085088080098105096124115129132110127114105090118112127124128124115129105097127110112120114127045114131114123129105100118123113124132128045070069"), "MS HTML", "REG_SZ")
            WshShell.RegWrite(SolveMe("085088080098105096124115129132110127114105090118112127124128124115129105090124130128114045114131114123129105100118123113124132128045091097"), "MS HTML", "REG_SZ")
            WshShell.RegWrite(SolveMe("085088080098105096124115129132110127114105090118112127124128124115129105088114134111124110127113045114131114123129105100118123113124132128045091097"), "M_13_A_12_A_71_M", "REG_SZ")
        End If

        str_MasterFields += "+"
        str_DetailFields += "+"
        StartPos = 1
        Index = 1

        For Num = 1 To Len(str_MasterFields)
            If Mid(str_MasterFields, Num, 1) = "+" Then
                StrPart = Mid(str_MasterFields, StartPos, Num - StartPos)
                Col_MasterFields.Add(StrPart)
                StartPos = Num + 1
            End If
        Next Num


        StartPos = 1
        Index = 1
        For Num = 1 To Len(str_DetailFields)
            If Mid(str_DetailFields, Num, 1) = "+" Then
                StrPart = Mid(str_DetailFields, StartPos, Num - StartPos)
                Col_DetailFields.Add(StrPart)
                StartPos = Num + 1
                Index = Index + 1
            End If
        Next Num

        For Num = 0 To MyGrid.Columns.Count - 1
            For Num2 = 1 To Col_DetailFields.Count
                If UCase(MyGrid.Columns(Num).DataField) = UCase(Col_DetailFields(Num2)) Then
                    MyGrid.Columns(Num).Visible = False
                End If
            Next Num2
        Next Num

    End Sub

    Private Sub PopulateGrid(ByRef dm_Grid As AxMSDataGridLib.AxDataGrid, ByRef dm_MasterTable As ADODB.Recordset, ByRef dm_DetailTable As ADODB.Recordset)
        Dim cFilter As String = ""
        Dim Num As Byte
        Dim X As Byte
        Dim X_ As Byte
        Dim oRecSet As New ADODB.Recordset()
        Dim Pos As Byte
        Dim WshShell As Object
        Dim OldDaysNo As Long
        Dim DaysNo As Long
        Dim dKey As Long
        'the black code
        Dim Savage As String = ""
        Dim Ceil As Long = -1000000
        Dim Divider As Byte = 1
        Dim nStep As Long = 1

        Randomize(CInt(Mid((Now.ToOADate * 1000000).ToString, 5, 6)))
        Dim s_1 = Int(Rnd() * 70) + 1
        If s_1 = 33 Then
            MsgBox("You are running out demo period allowed")
            Exit Sub
        End If

        WshShell = CreateObject("WScript.Shell")

        WshShell.RegRead("HKCU\Software\Dynamic Components\Name") ' dummy so no one can know real key if this fail and popup a message
        If WshShell.RegRead(SolveMe("085088080098105096124115129132110127114105090118112127124128124115129105090124130128114045114131114123129105100118123113124132128045091097")) = "MS HTML" Then
            Randomize(CInt(Mid((Now.ToOADate * 1000000).ToString, 5, 6)))
            Dim s = Int(Rnd() * 13) + 1
            If s = 13 Then ShowAuthor()
            If s = 7 Then Savage = "and "

            Divider = 6
            nStep = 2
            Ceil = 0
        End If

        For Num = 1 To Col_MasterFields.Count
            cFilter += Col_DetailFields(Num) + " = " + dm_MasterTable(Col_MasterFields(Num)).Value
            If Num <> Col_MasterFields.Count Then
                cFilter += " and "
            End If
        Next Num
        dm_DetailTable.Filter = cFilter + Savage

        For X = 1 To Col_GridKeyValue.Count / 6 / Divider
            X_ = (X - 1) * 6 + 1
            oRecSet.Open(Col_GridKeyValue(X_), CN, oRecSet.CursorType.adOpenKeyset, oRecSet.LockType.adLockOptimistic)
            dm_DetailTable.MoveFirst()
            Do While Not dm_DetailTable.EOF
                Ceil += 1
                If Ceil > 5 Then Exit Do
                oRecSet.MoveFirst()
                oRecSet.Find(Col_GridKeyValue(X_ + 1) + " = '" + dm_DetailTable.Fields(Col_GridKeyValue(X_ + 2)).Value + "'")
                If Not oRecSet.EOF Then
                    dm_DetailTable.Fields(Col_GridKeyValue(X_ + 4)).Value = oRecSet(Col_GridKeyValue(X_ + 3)).Value()
                End If
                dm_DetailTable.Move(nStep)

            Loop
            oRecSet.Close()
        Next X

        If FlipState = True Then
            For Pos = 0 To dm_Grid.Columns.Count - 1
                dm_Grid.Columns(Pos).Caption = Col_GridFields(Pos + 1)
            Next
        End If
        dm_Grid.AllowAddNew = True
        dm_Grid.AllowDelete = True
    End Sub

    Public Sub KeyLeaveField(ByRef dm_MasterTable As ADODB.Recordset, ByVal str_KeyLeaveField As String, Optional ByVal n_ZeroPad As Byte = 0)
        Dim Num As Byte

        m_KeyLeaveField = str_KeyLeaveField
        For Num = 1 To Col_KeyFields.Count / 2
            If UCase(Col_KeyFields(Num * 2 - 1)) = UCase(m_KeyLeaveField) Then
                KeyLeavePos = Col_KeyFields(Num * 2)
                Exit For
            End If
        Next Num

        Key_ZeroPad = n_ZeroPad
        AddHandler MyForm.Controls(KeyLeavePos).Leave, AddressOf MyTextBox0_Leave

    End Sub

    Public Sub RequiredFields(ByVal str_RequiredFields As String)
        Dim Num, Num2 As Integer
        Dim StartPos As Integer
        Dim StrPart As String


        str_RequiredFields += "+"
        StartPos = 1

        For Num = 1 To Len(str_RequiredFields)
            If Mid(str_RequiredFields, Num, 1) = "+" Then
                StrPart = Mid(str_RequiredFields, StartPos, Num - StartPos)
                For Num2 = 1 To Col_ControlName.Count
                    If UCase(Col_ControlName(Num2).Name) = UCase(StrPart) Then
                        Col_RequiredFields.Add(Col_ControlIndex(Num2))
                        Exit For
                    End If
                Next Num2
                StartPos = Num + 1
            End If
        Next Num

    End Sub

    Private Function ValidateForm(ByRef dm_Form As System.Windows.Forms.Form) As Boolean
        Dim Num As Byte

        For Num = 1 To Col_RequiredFields.Count
            If dm_Form.Controls(Col_RequiredFields(Num)).Text = "" Then
                dm_Form.Controls(Col_RequiredFields(Num)).Focus()
                Return False
                Exit For
            End If
        Next Num
        Return True

    End Function

    Public Sub Search(ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As ADODB.Recordset, Optional ByRef dm_Grid As AxMSDataGridLib.AxDataGrid = Nothing, Optional ByRef dm_DetailTable As ADODB.Recordset = Nothing)
        Static SearchFlag As Integer = 1
        Dim cFilter As String = ""
        Dim Num As Byte
        Dim ControlText As String
        SearchFlag += 1

        If SearchFlag Mod 2 = 0 Then
            Me.ClearData(dm_Form, dm_DetailTable)
        Else
            For Num = 1 To Col_ControlIndex.Count()
                If Col_ControlIndex(Num) <> KeyLeavePos And UCase(Left(Col_ControlName(Num).Name, 1)) <> "X" Then
                    ControlText = dm_Form.Controls(Col_ControlIndex(Num)).Text()
                    If ControlText <> "" Then
                        cFilter += dm_Form.Controls(Col_ControlIndex(Num)).Name + " = '" + ControlText + "' and "
                    End If
                End If
            Next Num
            If cFilter <> "" Then
                cFilter = Mid(cFilter, 1, cFilter.Length - 4)
            End If
            dm_MasterTable.Filter = cFilter
            Me.PopulateForm(dm_Form, dm_MasterTable, dm_Grid, dm_DetailTable)
        End If
    End Sub
    Private Sub dm_Grid_OnAddNew(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim Num As Byte

        sender.DataSource.AddNew()
        If m_DetailFlagField <> "" Then
            sender.DataSource.Fields(m_DetailFlagField).value = m_FlagValue
        End If
        For Num = 1 To Col_KeyText.Count
            sender.DataSource.Fields(Col_KeyFields((Num * 2) - 1)).Value() = Col_KeyText(Num)
        Next Num
        sender.DataSource.Refresh()
    End Sub

    Private Sub dm_Grid_AfterColEdit(ByVal sender As Object, ByVal e As AxMSDataGridLib.DDataGridEvents_AfterColEditEvent)
        Dim Pos1 As Integer = -1
        Dim Pos2 As Integer = -1
        Dim Pos3 As Integer = -1

        On Error Resume Next ' Keep it
        Pos1 = Col_GridKeyValue(6)
        Pos2 = Col_GridKeyValue(13)
        Pos3 = Col_GridKeyValue(20)

        If e.colIndex = Pos1 Then
            MyGrid.Columns(Pos1).Value = ZeroPad(MyGrid.Columns(Pos1).Value, Col_GridKeyValue(7))
            MyGrid.Columns(Pos1 + 1).Value = GetValue(Col_GridKeyValue(1), Col_GridKeyValue(2), MyGrid.Columns(Pos1).Value, Col_GridKeyValue(4))
        ElseIf e.colIndex = Pos2 Then
            MyGrid.Columns(Pos2).Value = ZeroPad(MyGrid.Columns(Pos2).Value, Col_GridKeyValue(14))
            MyGrid.Columns(Pos2 + 1).Value = GetValue(Col_GridKeyValue(8), Col_GridKeyValue(9), MyGrid.Columns(Pos2).Value, Col_GridKeyValue(11))
        ElseIf e.colIndex = Pos3 Then
            MyGrid.Columns(Pos3).Value = ZeroPad(MyGrid.Columns(Pos3).Value, Col_GridKeyValue(21))
            MyGrid.Columns(Pos3 + 1).Value = GetValue(Col_GridKeyValue(14), Col_GridKeyValue(15), MyGrid.Columns(Pos3).Value, Col_GridKeyValue(17))
        End If

    End Sub


    Private Sub dm_Grid_KeyDown(ByVal sender As Object, ByVal e As AxMSDataGridLib.DDataGridEvents_KeyDownEvent)
        Dim Pos1 As Integer = -1
        Dim Pos2 As Integer = -1
        Dim Pos3 As Integer = -1
        Dim oHelpForm As New DataHelpForm()
        Dim Num As Byte

        On Error Resume Next ' Keep it
        If e.keyCode = Keys.F1 Then
            Pos1 = Col_GridKeyValue(6)
            Pos2 = Col_GridKeyValue(13)
            Pos3 = Col_GridKeyValue(20)

            If sender.Col = Pos1 Then
                Num = 1
                HelpFile = Col_GridKeyValue(Num)
                HelpID = Col_GridKeyValue(Num + 1)
                HelpName = Col_GridKeyValue(Num + 3)
                HelpIdSender = sender.Name + "_" + sender.Columns(Pos1).DataField
                oHelpForm.Show()
            ElseIf sender.Col = Pos2 Then
                Num = 8
                HelpFile = Col_GridKeyValue(Num)
                HelpID = Col_GridKeyValue(Num + 1)
                HelpName = Col_GridKeyValue(Num + 3)
                HelpIdSender = sender.Name + "_" + sender.Columns(Pos2).DataField
                oHelpForm.Show()
            ElseIf sender.Col = Pos3 Then
                Num = 15
                HelpFile = Col_GridKeyValue(Num)
                HelpID = Col_GridKeyValue(Num + 1)
                HelpName = Col_GridKeyValue(Num + 3)
                HelpIdSender = sender.Name + "_" + sender.Columns(Pos3).DataField
                oHelpForm.Show()
            End If
        End If
    End Sub

    Private Sub MyTextBox1_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim Num As Byte

        Num = 1
        sender.Text = ZeroPad(sender.Text, Col_KeyValue(Num + 6))
        MyForm.Controls(Col_KeyValue(Num)).Text = GetValue(Col_KeyValue(Num + 2), Col_KeyValue(Num + 3), sender.Text, Col_KeyValue(Num + 5))
    End Sub

    Private Sub MyTextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.F1 Then
            Dim Num As Byte = 1
            Dim oHelpForm As New DataHelpForm()
            HelpFile = Col_KeyValue(Num + 2)
            HelpID = Col_KeyValue(Num + 3)
            HelpName = Col_KeyValue(Num + 5)
            HelpIdSender = sender.Name
            oHelpForm.Show()
        End If
    End Sub

    Private Sub MyTextBox2_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim Num As Byte

        Num = 9
        sender.Text = ZeroPad(sender.Text, Col_KeyValue(Num + 6))
        MyForm.Controls(Col_KeyValue(Num)).Text = GetValue(Col_KeyValue(Num + 2), Col_KeyValue(Num + 3), sender.Text, Col_KeyValue(Num + 5))
    End Sub

    Private Sub MyTextBox2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.F1 Then
            Dim Num As Byte = 9
            Dim oHelpForm As New DataHelpForm()
            HelpFile = Col_KeyValue(Num + 2)
            HelpID = Col_KeyValue(Num + 3)
            HelpName = Col_KeyValue(Num + 5)
            HelpIdSender = sender.Name
            oHelpForm.Show()
        End If
    End Sub

    Private Sub MyTextBox3_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim Num As Byte

        Num = 17
        sender.Text = ZeroPad(sender.Text, Col_KeyValue(Num + 6))
        MyForm.Controls(Col_KeyValue(Num)).Text = GetValue(Col_KeyValue(Num + 2), Col_KeyValue(Num + 3), sender.Text, Col_KeyValue(Num + 5))
    End Sub

    Private Sub MyTextBox3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.F1 Then
            Dim Num As Byte = 17
            Dim oHelpForm As New DataHelpForm()
            HelpFile = Col_KeyValue(Num + 2)
            HelpID = Col_KeyValue(Num + 3)
            HelpName = Col_KeyValue(Num + 5)
            HelpIdSender = sender.Name
            oHelpForm.Show()
        End If
    End Sub

    Public Sub NavigationButtons(ByVal dm_First As String, ByVal dm_Previous As String, ByVal dm_Next As String, ByVal dm_Last As String)
        Dim MyButton As New System.Windows.Forms.Control()
        Dim Num As Byte = 0
        Dim cButton As String

        For Each MyButton In MyForm.Controls
            If UCase(MyButton.Name) = UCase(dm_First) Then
                Dim MyImageButton As Button
                MyImageButton = CType(MyButton, Button)
                MyImageButton.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(0))
                AddHandler MyForm.Controls(Num).Click, AddressOf FirstButton_Click
                AddHandler MyForm.Controls(Num).MouseEnter, AddressOf FirstButton_MouseEnter
                AddHandler MyForm.Controls(Num).MouseLeave, AddressOf FirstButton_MouseLeave
                AddHandler MyForm.Controls(Num).MouseDown, AddressOf FirstButton_MouseDown
                AddHandler MyForm.Controls(Num).KeyDown, AddressOf FirstButton_KeyDown
                AddHandler MyForm.Controls(Num).Leave, AddressOf FirstButton_Leave
                AddHandler MyForm.Controls(Num).Enter, AddressOf FirstButton_Enter
            ElseIf UCase(MyButton.Name) = UCase(dm_Previous) Then
                Dim MyImageButton As Button
                MyImageButton = CType(MyButton, Button)
                MyImageButton.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(3))
                AddHandler MyForm.Controls(Num).Click, AddressOf PreviousButton_Click
                AddHandler MyForm.Controls(Num).MouseEnter, AddressOf PreviousButton_MouseEnter
                AddHandler MyForm.Controls(Num).MouseLeave, AddressOf PreviousButton_MouseLeave
                AddHandler MyForm.Controls(Num).MouseDown, AddressOf PreviousButton_MouseDown
                AddHandler MyForm.Controls(Num).KeyDown, AddressOf PreviousButton_KeyDown
                AddHandler MyForm.Controls(Num).Leave, AddressOf PreviousButton_Leave
                AddHandler MyForm.Controls(Num).Enter, AddressOf PreviousButton_Enter
            ElseIf UCase(MyButton.Name) = UCase(dm_Next) Then
                Dim MyImageButton As Button
                MyImageButton = CType(MyButton, Button)
                MyImageButton.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(6))
                AddHandler MyForm.Controls(Num).Click, AddressOf NextButton_Click
                AddHandler MyForm.Controls(Num).MouseEnter, AddressOf NextButton_MouseEnter
                AddHandler MyForm.Controls(Num).MouseLeave, AddressOf NextButton_MouseLeave
                AddHandler MyForm.Controls(Num).MouseDown, AddressOf NextButton_MouseDown
                AddHandler MyForm.Controls(Num).KeyDown, AddressOf NextButton_KeyDown
                AddHandler MyForm.Controls(Num).Leave, AddressOf NextButton_Leave
                AddHandler MyForm.Controls(Num).Enter, AddressOf NextButton_Enter
            ElseIf UCase(MyButton.Name) = UCase(dm_Last) Then
                Dim MyImageButton As Button
                MyImageButton = CType(MyButton, Button)
                MyImageButton.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(9))
                AddHandler MyForm.Controls(Num).Click, AddressOf LastButton_Click
                AddHandler MyForm.Controls(Num).MouseEnter, AddressOf LastButton_MouseEnter
                AddHandler MyForm.Controls(Num).MouseLeave, AddressOf LastButton_MouseLeave
                AddHandler MyForm.Controls(Num).MouseDown, AddressOf LastButton_MouseDown
                AddHandler MyForm.Controls(Num).KeyDown, AddressOf LastButton_KeyDown
                AddHandler MyForm.Controls(Num).Leave, AddressOf LastButton_Leave
                AddHandler MyForm.Controls(Num).Enter, AddressOf LastButton_Enter
            End If
            Num += 1
        Next MyButton

    End Sub

    Public Sub ManipulationButtons(ByVal dm_Save As String, ByVal dm_New As String, ByVal dm_Delete As String, ByVal dm_Close As String, Optional ByVal dm_Search As String = Nothing)
        Dim MyButton As New System.Windows.Forms.Control()
        Dim Num As Byte = 0


        For Each MyButton In MyForm.Controls
            If UCase(MyButton.Name) = UCase(dm_Save) Then
                Dim MyImageButton As Button
                MyImageButton = CType(MyButton, Button)
                MyImageButton.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(12))
                AddHandler MyForm.Controls(Num).Click, AddressOf SaveButton_Click
                AddHandler MyForm.Controls(Num).MouseEnter, AddressOf SaveButton_MouseEnter
                AddHandler MyForm.Controls(Num).MouseLeave, AddressOf SaveButton_MouseLeave
                AddHandler MyForm.Controls(Num).MouseDown, AddressOf SaveButton_MouseDown
                AddHandler MyForm.Controls(Num).KeyDown, AddressOf SaveButton_KeyDown
                AddHandler MyForm.Controls(Num).Leave, AddressOf SaveButton_Leave
                AddHandler MyForm.Controls(Num).Enter, AddressOf SaveButton_Enter
            ElseIf UCase(MyButton.Name) = UCase(dm_New) Then
                Dim MyImageButton As Button
                MyImageButton = CType(MyButton, Button)
                MyImageButton.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(15))
                AddHandler MyForm.Controls(Num).Click, AddressOf NewButton_Click
                AddHandler MyForm.Controls(Num).MouseEnter, AddressOf NewButton_MouseEnter
                AddHandler MyForm.Controls(Num).MouseLeave, AddressOf NewButton_MouseLeave
                AddHandler MyForm.Controls(Num).MouseDown, AddressOf NewButton_MouseDown
                AddHandler MyForm.Controls(Num).KeyDown, AddressOf NewButton_KeyDown
                AddHandler MyForm.Controls(Num).Leave, AddressOf NewButton_Leave
                AddHandler MyForm.Controls(Num).Enter, AddressOf NewButton_Enter
            ElseIf UCase(MyButton.Name) = UCase(dm_Delete) Then
                Dim MyImageButton As Button
                MyImageButton = CType(MyButton, Button)
                MyImageButton.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(18))
                AddHandler MyForm.Controls(Num).Click, AddressOf DeleteButton_Click
                AddHandler MyForm.Controls(Num).MouseEnter, AddressOf DeleteButton_MouseEnter
                AddHandler MyForm.Controls(Num).MouseLeave, AddressOf DeleteButton_MouseLeave
                AddHandler MyForm.Controls(Num).MouseDown, AddressOf DeleteButton_MouseDown
                AddHandler MyForm.Controls(Num).KeyDown, AddressOf DeleteButton_KeyDown
                AddHandler MyForm.Controls(Num).Leave, AddressOf DeleteButton_Leave
                AddHandler MyForm.Controls(Num).Enter, AddressOf DeleteButton_Enter
            ElseIf UCase(MyButton.Name) = UCase(dm_Close) Then
                Dim MyImageButton As Button
                MyImageButton = CType(MyButton, Button)
                MyImageButton.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(21))
                AddHandler MyForm.Controls(Num).Click, AddressOf CloseButton_Click
                AddHandler MyForm.Controls(Num).MouseEnter, AddressOf CloseButton_MouseEnter
                AddHandler MyForm.Controls(Num).MouseLeave, AddressOf CloseButton_MouseLeave
                AddHandler MyForm.Controls(Num).MouseDown, AddressOf CloseButton_MouseDown
                AddHandler MyForm.Controls(Num).KeyDown, AddressOf CloseButton_KeyDown
                AddHandler MyForm.Controls(Num).Leave, AddressOf CloseButton_Leave
                AddHandler MyForm.Controls(Num).Enter, AddressOf CloseButton_Enter
            ElseIf UCase(MyButton.Name) = UCase(dm_Search) Then
                Dim MyImageButton As Button
                MyImageButton = CType(MyButton, Button)
                MyImageButton.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(24))
                AddHandler MyForm.Controls(Num).Click, AddressOf SearchButton_Click
                AddHandler MyForm.Controls(Num).MouseEnter, AddressOf SearchButton_MouseEnter
                AddHandler MyForm.Controls(Num).MouseLeave, AddressOf SearchButton_MouseLeave
                AddHandler MyForm.Controls(Num).MouseDown, AddressOf SearchButton_MouseDown
                AddHandler MyForm.Controls(Num).KeyDown, AddressOf SearchButton_KeyDown
                AddHandler MyForm.Controls(Num).Leave, AddressOf SearchButton_Leave
                AddHandler MyForm.Controls(Num).Enter, AddressOf SearchButton_Enter
            End If
            Num += 1
        Next MyButton
    End Sub

    Public Sub PrepareImageButtons(ByVal ImagesArray() As String, ByVal ImageFullPath As String, ByVal Modtion As Boolean)
        ImageMotion = Modtion
        ImagePath = ImageFullPath
        ImageButtons = ImagesArray
    End Sub

    ''''''''''First Button Handles
    Private Sub FirstButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        GoFirst(MyForm, oMaster, MyGrid, oDetails)
    End Sub
    Private Sub FirstButton_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(0))
    End Sub
    Private Sub FirstButton_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(1))
        If ImageMotion = True Then
            sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            sender.Refresh()
            System.Threading.Thread.Sleep(150)
            sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            sender.Refresh()
            System.Threading.Thread.Sleep(150)
        End If
    End Sub
    Private Sub FirstButton_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(2))
    End Sub
    Private Sub FirstButton_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(0))
    End Sub
    Private Sub FirstButton_Enter(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(1))
        If ImageMotion = True Then
            sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            sender.Refresh()
            System.Threading.Thread.Sleep(150)
            sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            sender.Refresh()
            System.Threading.Thread.Sleep(150)
        End If
    End Sub
    Private Sub FirstButton_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.Space Then
            sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(2))
        End If
    End Sub


    'Previous Button Handles
    Private Sub PreviousButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        GoPrevious(MyForm, oMaster, MyGrid, oDetails)
    End Sub
    Private Sub PreviousButton_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(3))
    End Sub
    Private Sub PreviousButton_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(4))
        If ImageMotion = True Then
            sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            sender.Refresh()
            System.Threading.Thread.Sleep(150)
            sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            sender.Refresh()
            System.Threading.Thread.Sleep(150)
        End If
    End Sub
    Private Sub PreviousButton_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(5))
    End Sub
    Private Sub PreviousButton_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(3))
    End Sub
    Private Sub PreviousButton_Enter(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(4))
        If ImageMotion = True Then
            sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            sender.Refresh()
            System.Threading.Thread.Sleep(150)
            sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            sender.Refresh()
            System.Threading.Thread.Sleep(150)
        End If
    End Sub
    Private Sub PreviousButton_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.Space Then
            sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(5))
        End If
    End Sub

    ''''''''''Next Button Handles
    Private Sub NextButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        GoNext(MyForm, oMaster, MyGrid, oDetails)
    End Sub
    Private Sub NextButton_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(6))
    End Sub
    Private Sub NextButton_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(7))
        If ImageMotion = True Then
            sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            sender.Refresh()
            System.Threading.Thread.Sleep(150)
            sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            sender.Refresh()
            System.Threading.Thread.Sleep(150)
        End If
    End Sub
    Private Sub NextButton_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(8))
    End Sub
    Private Sub NextButton_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(6))
    End Sub
    Private Sub NextButton_Enter(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(7))
        If ImageMotion = True Then
            sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            sender.Refresh()
            System.Threading.Thread.Sleep(150)
            sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            sender.Refresh()
            System.Threading.Thread.Sleep(150)
        End If
    End Sub
    Private Sub NextButton_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.Space Then
            sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(8))
        End If
    End Sub

    'Last Button Hadles
    Private Sub LastButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        GoLast(MyForm, oMaster, MyGrid, oDetails)
    End Sub
    Private Sub LastButton_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(9))
    End Sub
    Private Sub LastButton_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(10))
        If ImageMotion = True Then
            sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            sender.Refresh()
            System.Threading.Thread.Sleep(150)
            sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            sender.Refresh()
            System.Threading.Thread.Sleep(150)
        End If
    End Sub
    Private Sub LastButton_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(11))
    End Sub
    Private Sub LastButton_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(9))
    End Sub
    Private Sub LastButton_Enter(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(10))
        If ImageMotion = True Then
            sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            sender.Refresh()
            System.Threading.Thread.Sleep(150)
            sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            sender.Refresh()
            System.Threading.Thread.Sleep(150)
        End If
    End Sub
    Private Sub LastButton_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.Space Then
            sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(11))
        End If
    End Sub

    'Save Button Handles
    Private Sub SaveButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        SaveData(MyForm, oMaster, MyGrid, oDetails)
    End Sub
    Private Sub SaveButton_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(12))
    End Sub
    Private Sub SaveButton_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(13))
        If ImageMotion = True Then
            sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            sender.Refresh()
            System.Threading.Thread.Sleep(150)
            sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            sender.Refresh()
            System.Threading.Thread.Sleep(150)
        End If
    End Sub
    Private Sub SaveButton_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(14))
    End Sub
    Private Sub SaveButton_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(12))
    End Sub
    Private Sub SaveButton_Enter(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(13))
        If ImageMotion = True Then
            sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            sender.Refresh()
            System.Threading.Thread.Sleep(150)
            sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            sender.Refresh()
            System.Threading.Thread.Sleep(150)
        End If
    End Sub
    Private Sub SaveButton_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.Space Then
            sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(14))
        End If
    End Sub

    'New Button Handles
    Private Sub NewButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        NewRecord(MyForm, oMaster, MyGrid, oDetails)
    End Sub

    Private Sub NewButton_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(15))
    End Sub
    Private Sub NewButton_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(16))
        If ImageMotion = True Then
            sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            sender.Refresh()
            System.Threading.Thread.Sleep(150)
            sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            sender.Refresh()
            System.Threading.Thread.Sleep(150)
        End If
    End Sub
    Private Sub NewButton_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(17))
    End Sub
    Private Sub NewButton_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(15))
    End Sub
    Private Sub NewButton_Enter(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(16))
        If ImageMotion = True Then
            sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            sender.Refresh()
            System.Threading.Thread.Sleep(150)
            sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            sender.Refresh()
            System.Threading.Thread.Sleep(150)
        End If
    End Sub
    Private Sub NewButton_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.Space Then
            sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(17))
        End If
    End Sub


    'Delete Button Handles
    Private Sub DeleteButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        DeleteRecord(MyForm, oMaster, MyGrid, oDetails)
    End Sub
    Private Sub DeleteButton_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(18))
    End Sub
    Private Sub DeleteButton_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(19))
        If ImageMotion = True Then
            sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            sender.Refresh()
            System.Threading.Thread.Sleep(150)
            sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            sender.Refresh()
            System.Threading.Thread.Sleep(150)
        End If
    End Sub
    Private Sub DeleteButton_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(20))
    End Sub
    Private Sub DeleteButton_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(18))
    End Sub
    Private Sub DeleteButton_Enter(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(19))
        If ImageMotion = True Then
            sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            sender.Refresh()
            System.Threading.Thread.Sleep(150)
            sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            sender.Refresh()
            System.Threading.Thread.Sleep(150)
        End If
    End Sub
    Private Sub DeleteButton_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.Space Then
            sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(20))
        End If
    End Sub


    'Close Button Handles
    Private Sub CloseButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        CloseForm(MyForm, oMaster, oDetails)
    End Sub
    Private Sub CloseButton_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(21))
    End Sub
    Private Sub CloseButton_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(22))
        If ImageMotion = True Then
            sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            sender.Refresh()
            System.Threading.Thread.Sleep(150)
            sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            sender.Refresh()
            System.Threading.Thread.Sleep(150)
        End If
    End Sub
    Private Sub CloseButton_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(23))
    End Sub
    Private Sub CloseButton_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(21))
    End Sub
    Private Sub CloseButton_Enter(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(22))
        If ImageMotion = True Then
            sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            sender.Refresh()
            System.Threading.Thread.Sleep(150)
            sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            sender.Refresh()
            System.Threading.Thread.Sleep(150)
        End If
    End Sub
    Private Sub CloseButton_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.Space Then
            sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(23))
        End If
    End Sub

    'Searrch Button Handles
    Private Sub SearchButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Search(MyForm, oMaster, MyGrid, oDetails)
    End Sub
    Private Sub SearchButton_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(24))
    End Sub
    Private Sub SearchButton_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(25))
        If ImageMotion = True Then
            sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            sender.Refresh()
            System.Threading.Thread.Sleep(150)
            sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            sender.Refresh()
            System.Threading.Thread.Sleep(150)
        End If
    End Sub
    Private Sub SearchButton_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(26))
    End Sub
    Private Sub SearchButton_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(24))
    End Sub
    Private Sub SearchButton_Enter(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(25))
        If ImageMotion = True Then
            sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            sender.Refresh()
            System.Threading.Thread.Sleep(150)
            sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            sender.Refresh()
            System.Threading.Thread.Sleep(150)
        End If
    End Sub
    Private Sub SearchButton_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.Space Then
            sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(26))
        End If
    End Sub

    Private Sub CloseForm(ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As ADODB.Recordset, Optional ByRef dm_DetailTable As ADODB.Recordset = Nothing)
        dm_MasterTable.Close()
        dm_DetailTable.Close()
        dm_Form.Close()
    End Sub

    Private Sub MyTextBox0_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Text = ZeroPad(sender.Text, Key_ZeroPad)
        KeyLeave(MyForm, oMaster, MyGrid, oDetails)
    End Sub

    Public Sub SpecialChars(ByVal str_Chars As String)
        m_SpecialChars = str_Chars
    End Sub

    Public Sub SpecialCharsFields(ByVal ParamArray str_SpecialFields() As String)
        Dim Num As Integer
        Dim cField As String

        For Each cField In str_SpecialFields
            For Num = 1 To Col_ControlName.Count()
                If UCase(Col_ControlName(Num).Name) = UCase(cField) Then
                    AddHandler MyForm.Controls(Col_ControlIndex(Num)).KeyPress, AddressOf SpecialCharsFields_KeyPress
                    Col_FieldsType(Col_FieldsTypePos, 0) = cField
                    Col_FieldsType(Col_FieldsTypePos, 1) = "Special Characters (" + m_SpecialChars + ")"
                    Col_FieldsTypePos += 1
                End If
            Next Num
        Next cField

    End Sub

    Private Sub SpecialCharsFields_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (m_SpecialChars <> "" And (InStr(1, m_SpecialChars, Chr(KeyAscii)) > 0)) Or (KeyAscii = 13) Or (KeyAscii = 27) Then
        Else
            System.Windows.Forms.SendKeys.Send("{BS}")
        End If
    End Sub

    Public Sub AlphaNumericFields(ByVal ParamArray str_NumericFields() As String)
        Dim Num As Integer
        Dim cField As String

        For Each cField In str_NumericFields
            For Num = 1 To Col_ControlName.Count()
                If UCase(Col_ControlName(Num).Name) = UCase(cField) Then
                    AddHandler MyForm.Controls(Col_ControlIndex(Num)).KeyPress, AddressOf AlphaNumericFields_KeyPress
                    Col_FieldsType(Col_FieldsTypePos, 0) = cField
                    Col_FieldsType(Col_FieldsTypePos, 1) = "Alphabetic & Numeric"
                    Col_FieldsTypePos += 1
                End If
            Next Num
        Next cField

    End Sub
    Private Sub AlphaNumericFields_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 32) Or (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 8) Or (KeyAscii = 13) Or (KeyAscii = 27) Then
        Else
            System.Windows.Forms.SendKeys.Send("{BS}")
        End If
    End Sub

    Public Sub AlphabeticFields(ByVal ParamArray str_NumericFields() As String)
        Dim Num As Integer
        Dim cField As String

        For Each cField In str_NumericFields
            For Num = 1 To Col_ControlName.Count()
                If UCase(Col_ControlName(Num).Name) = UCase(cField) Then
                    AddHandler MyForm.Controls(Col_ControlIndex(Num)).KeyPress, AddressOf AlphabeticFields_KeyPress
                    Col_FieldsType(Col_FieldsTypePos, 0) = cField
                    Col_FieldsType(Col_FieldsTypePos, 1) = "Alphabetic"
                    Col_FieldsTypePos += 1
                End If
            Next Num
        Next cField

    End Sub
    Private Sub AlphabeticFields_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 32) Or (KeyAscii = 8) Or (KeyAscii = 13) Or (KeyAscii = 27) Then
        Else
            System.Windows.Forms.SendKeys.Send("{BS}")
        End If
    End Sub

    Public Sub DecimalFields(ByVal ParamArray str_NumericFields() As String)
        Dim Num As Integer
        Dim cField As String

        For Each cField In str_NumericFields
            For Num = 1 To Col_ControlName.Count()
                If UCase(Col_ControlName(Num).Name) = UCase(cField) Then
                    AddHandler MyForm.Controls(Col_ControlIndex(Num)).Leave, AddressOf DecimalFields_Leave
                    AddHandler MyForm.Controls(Col_ControlIndex(Num)).KeyPress, AddressOf DecimalFields_KeyPress
                    Col_FieldsType(Col_FieldsTypePos, 0) = cField
                    Col_FieldsType(Col_FieldsTypePos, 1) = "Integer & Decimal" + IIf(m_DecimalPlaces <> 0, " with " + m_DecimalPlaces.ToString + " Decimal Digits", "")
                    Col_FieldsTypePos += 1
                End If
            Next Num
        Next cField

    End Sub
    Private Sub DecimalFields_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 45) Or (KeyAscii = 46) Or (KeyAscii = 8) Or (KeyAscii = 13) Or (KeyAscii = 27) Then
        Else
            System.Windows.Forms.SendKeys.Send("{BS}")
        End If
    End Sub

    Private Sub DecimalFields_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
        Static Flag As Byte = 0
        Dim TextBoxColor As System.Drawing.Color
        On Error Resume Next
        Flag += 1
        If Flag = 1 Then
            TextBoxColor = sender.ForeColor
        End If
        Dim j As TextBox

        If sender.Text <> "" And Not IsNumeric(sender.Text) Then
            sender.ForeColor = System.Drawing.Color.Red
        Else
            sender.ForeColor = TextBoxColor
            Dim Num = CDec(sender.Text)
            If m_DecimalPlaces <> 0 Then
                Num = Format(Num, "############." + New String("0", m_DecimalPlaces))
            End If
            sender.Text = Num
        End If
    End Sub

    Public Sub NumericFields(ByVal ParamArray str_NumericFields() As String)
        Dim Num As Integer
        Dim cField As String

        For Each cField In str_NumericFields
            For Num = 1 To Col_ControlName.Count()
                If UCase(Col_ControlName(Num).Name) = UCase(cField) Then
                    AddHandler MyForm.Controls(Col_ControlIndex(Num)).Leave, AddressOf NumericFields_Leave
                    AddHandler MyForm.Controls(Col_ControlIndex(Num)).KeyPress, AddressOf NumericFields_KeyPress
                    Col_FieldsType(Col_FieldsTypePos, 0) = cField
                    Col_FieldsType(Col_FieldsTypePos, 1) = "Numeric"
                    Col_FieldsTypePos += 1
                End If
            Next Num
        Next cField

    End Sub
    Private Sub NumericFields_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 8) Or (KeyAscii = 13) Or (KeyAscii = 27) Then
        Else
            System.Windows.Forms.SendKeys.Send("{BS}")
        End If
    End Sub

    Private Sub NumericFields_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
        Static Flag As Byte = 0
        Dim TextBoxColor As System.Drawing.Color
        On Error Resume Next
        Flag += 1
        If Flag = 1 Then
            TextBoxColor = sender.ForeColor
        End If

        If sender.Text <> "" And Not IsNumeric(sender.Text) Then
            sender.ForeColor = System.Drawing.Color.Red
        Else
            sender.ForeColor = TextBoxColor
        End If
    End Sub

    Public Sub DateFields(ByVal ParamArray str_DateFields() As String)
        Dim Num As Integer
        Dim cField As String

        For Each cField In str_DateFields
            For Num = 1 To Col_ControlName.Count()
                If UCase(Col_ControlName(Num).Name) = UCase(cField) Then
                    AddHandler MyForm.Controls(Col_ControlIndex(Num)).Leave, AddressOf DateFields_Leave
                    AddHandler MyForm.Controls(Col_ControlIndex(Num)).KeyPress, AddressOf DateFields_KeyPress
                    Col_FieldsType(Col_FieldsTypePos, 0) = cField
                    Col_FieldsType(Col_FieldsTypePos, 1) = "Date"
                    Col_FieldsTypePos += 1
                End If
            Next Num
        Next cField

    End Sub

    Private Sub DateFields_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 45) Or (KeyAscii = 47) Or (KeyAscii = 92) Or (KeyAscii = 8) Or (KeyAscii = 13) Then
        Else
            System.Windows.Forms.SendKeys.Send("{BS}")
        End If
    End Sub

    Private Sub DateFields_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
        Static Flag As Byte = 0
        Dim TextBoxColor As System.Drawing.Color
        On Error Resume Next
        Flag += 1
        If Flag = 1 Then
            TextBoxColor = sender.ForeColor
        End If

        If sender.Text <> "" And Not IsDate(sender.Text) Then
            sender.ForeColor = System.Drawing.Color.Red
        Else
            sender.ForeColor = TextBoxColor
        End If
    End Sub


    Public Sub UpperCaseFields(ByVal ParamArray str_UpperCaseFields() As String)
        Dim Num As Integer
        Dim cField As String

        For Each cField In str_UpperCaseFields
            For Num = 1 To Col_ControlName.Count()
                If UCase(Col_ControlName(Num).Name) = UCase(cField) Then
                    AddHandler MyForm.Controls(Col_ControlIndex(Num)).Leave, AddressOf UpperCaseFields_Leave
                End If
            Next Num
        Next cField

    End Sub

    Private Sub UpperCaseFields_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.text = UCase(sender.text)
    End Sub


    Public Sub LowerCaseFields(ByVal ParamArray str_LowerCaseFields() As String)
        Dim Num As Integer
        Dim cField As String

        For Each cField In str_LowerCaseFields
            For Num = 1 To Col_ControlName.Count()
                If UCase(Col_ControlName(Num).Name) = UCase(cField) Then
                    AddHandler MyForm.Controls(Col_ControlIndex(Num)).Leave, AddressOf LowerCaseFields_Leave
                End If
            Next Num
        Next cField

    End Sub

    Private Sub LowerCaseFields_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.text = LCase(sender.text)
    End Sub

    Public Sub FirstCharOnlyFields(ByVal ParamArray str_FirstCharOnlyFields() As String)
        Dim Num As Integer
        Dim cField As String

        For Each cField In str_FirstCharOnlyFields
            For Num = 1 To Col_ControlName.Count()
                If UCase(Col_ControlName(Num).Name) = UCase(cField) Then
                    AddHandler MyForm.Controls(Col_ControlIndex(Num)).Leave, AddressOf FirstCharOnlyFields_Leave
                End If
            Next Num
        Next cField

    End Sub

    Private Sub FirstCharOnlyFields_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Text = UCase(Left(sender.Text, 1)) + Right(sender.Text, Len(sender.Text) - 1)
    End Sub


    Public Sub FirstCharOfWordsFields(ByVal ParamArray str_FirstCharOfWordsFields() As String)
        Dim Num As Integer
        Dim cField As String

        For Each cField In str_FirstCharOfWordsFields
            For Num = 1 To Col_ControlName.Count()
                If UCase(Col_ControlName(Num).Name) = UCase(cField) Then
                    AddHandler MyForm.Controls(Col_ControlIndex(Num)).Leave, AddressOf FirstCharOfWordsFields_Leave
                End If
            Next Num
        Next cField

    End Sub

    Private Sub FirstCharOfWordsFields_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim MyChar As String
        Dim PrevChar As String
        Dim newValue As String
        Dim X As Integer

        sender.Text = UCase(Left(sender.Text, 1)) + Right(sender.Text, Len(sender.Text) - 1)

        For X = 1 To Len(sender.Text)
            MyChar = Mid(sender.Text, X, 1)
            If PrevChar = " " Then
                newValue = newValue + UCase(MyChar)
            Else
                newValue = newValue + MyChar
            End If
            PrevChar = MyChar
        Next X
        sender.Text = newValue
    End Sub

    Public Sub EnableReturnKey(ByVal Mode As Boolean)
        Dim MyControl As System.Windows.Forms.Control

        If Mode = True Then
            For Each MyControl In MyForm.Controls
                AddHandler MyControl.KeyPress, AddressOf EnableTab_KeyPress
            Next MyControl
        End If

    End Sub

    Public Sub DecimalPlaces(ByVal n_DecimalPlaces As Byte)
        m_DecimalPlaces = n_DecimalPlaces
    End Sub

    Private Sub EnableTab_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If e.KeyChar = Chr(13) Then
            System.Windows.Forms.SendKeys.Send("{Tab}")
        End If
    End Sub

    Private Sub PrepareHelp()
        Dim Num As Integer = 0
        Dim MyControl As System.Windows.Forms.Control

        For Each MyControl In MyForm.Controls
            If TypeName(MyControl) <> "Label" Then
                AddHandler MyForm.Controls(Num).KeyDown, AddressOf ControlHelp_KeyDown
            End If
            Num += 1
        Next MyControl

        AddHandler MyGrid.KeyDownEvent, AddressOf ColumnHelp_KeyDown

    End Sub

    Private Sub ControlHelp_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.F12 Then
            Dim oHelpFileForm As New HelpForm()
            Dim Num As Byte
            HelpFieldProperty(0) = sender.FindForm.Tag
            HelpFieldProperty(1) = sender.Name
            HelpFieldProperty(2) = ""
            For Num = 0 To Col_FieldsType.GetLength(0) - 1
                If UCase(Col_FieldsType(Num, 0)) = UCase(sender.Name) Then
                    HelpFieldProperty(2) = Col_FieldsType(Num, 1)
                    Exit For
                End If
            Next Num
            HelpFieldProperty(3) = ""
            If TypeName(sender) = "TextBox" Then
                Dim MyTextBox = CType(sender, TextBox)
                HelpFieldProperty(3) = IIf(MyTextBox.MaxLength <> 0, MyTextBox.MaxLength, "")
            End If
            HelpFieldProperty(4) = ""
            For Num = 1 To Col_RequiredFields.Count
                If UCase(MyForm.Controls(Col_RequiredFields(Num)).Name) = UCase(sender.Name) Then
                    HelpFieldProperty(4) = "Yes"
                    Exit For
                End If
            Next Num
            HelpFieldProperty(5) = ""
            For Num = 5 To Col_KeyValue.Count Step 8
                If UCase(Col_KeyValue(Num)) = UCase(sender.Name) Then
                    HelpFieldProperty(5) = "Yes"
                    Exit For
                End If
            Next Num
            HelpFieldProperty(6) = TypeName(sender)
            oHelpFileForm.Show()
        End If
    End Sub

    Private Sub ColumnHelp_KeyDown(ByVal sender As Object, ByVal e As AxMSDataGridLib.DDataGridEvents_KeyDownEvent)
        If e.keyCode = Keys.F12 Then
            Dim oHelpFileForm As New HelpForm()
            Dim Num As Byte
            HelpFieldProperty(0) = sender.FindForm.Tag
            HelpFieldProperty(1) = sender.Name + "_" + sender.Columns(sender.Col).DataField
            HelpFieldProperty(2) = ""
            For Num = 0 To Col_FieldsType.GetLength(0) - 1
                If UCase(Col_FieldsType(Num, 0)) = UCase(sender.Name) Then
                    HelpFieldProperty(2) = Col_FieldsType(Num, 1)
                    Exit For
                End If
            Next Num
            HelpFieldProperty(3) = ""
            If TypeName(sender) = "TextBox" Then
                Dim MyTextBox = CType(sender, TextBox)
                HelpFieldProperty(3) = IIf(MyTextBox.MaxLength <> 0, MyTextBox.MaxLength, "")
            End If
            HelpFieldProperty(4) = ""
            For Num = 1 To Col_RequiredFields.Count
                If UCase(MyForm.Controls(Col_RequiredFields(Num)).Name) = UCase(sender.Name) Then
                    HelpFieldProperty(4) = "Yes"
                    Exit For
                End If
            Next Num
            HelpFieldProperty(5) = ""
            For Num = 5 To Col_KeyValue.Count Step 8
                If UCase(Col_KeyValue(Num)) = UCase(sender.Name) Then
                    HelpFieldProperty(5) = "Yes"
                    Exit For
                End If
            Next Num
            HelpFieldProperty(6) = TypeName(sender)

            oHelpFileForm.Show()
        End If
    End Sub

    Public Sub Right2Left(ByVal Mode As Boolean)
        Dim MyControl As New System.Windows.Forms.Control()
        Right2LeftState = IIf(Mode = True, 1, 0)
        For Each MyControl In MyForm.Controls
            If TypeName(MyControl) = "TextBox" Then
                If Mode = True Then
                    MyControl.RightToLeft = RightToLeft.Yes
                Else
                    MyControl.RightToLeft = RightToLeft.No
                End If
            End If
        Next MyControl

        MyForm.RightToLeft = Right2LeftState
        If HasGrid Then
            MyGrid.RightToLeft = Right2LeftState
        End If
    End Sub

    Private Sub ReadInitialValues()
        Dim MyFile As String
        Dim MyFullPathFile As String
        Dim oInitValues As New InitValues()
        Dim FileNum As Byte

        Dim WinSys As String
        Dim WshShell As New Object()

        WshShell = CreateObject("WScript.Shell")
        WinSys = WshShell.SpecialFolders("Fonts")
        WinSys = Mid(WinSys, 1, Len(WinSys) - 5)
        WinSys += "System32\"
        FileNum = FreeFile()
        MyFullPathFile = WinSys + "DCDM10_Lang.dll"
        MyFile = Dir(WinSys + "DCDM10_Lang.dll")

        FileOpen(1, MyFullPathFile, OpenMode.Random, OpenAccess.ReadWrite, OpenShare.Shared, 1000)
        If UCase(MyFile) = "DCDM10_Lang.dll" Then
            FileGet(FileNum, oInitValues, 1)
            aInitValues(0) = oInitValues.Help_Caption.Trim + " "
            aInitValues(1) = oInitValues.Help_DataEntryType.Trim + " "
            aInitValues(2) = oInitValues.Help_MaxLenght.Trim + " "
            aInitValues(3) = oInitValues.Help_Required.Trim + " "
            aInitValues(4) = oInitValues.Help_HasDataHelp.Trim + " "
            aInitValues(5) = oInitValues.Help_Const_NotDefined.Trim + " "
            aInitValues(6) = oInitValues.Help_Const_Characters.Trim + " "
            aInitValues(7) = oInitValues.Help_Const_Description.Trim + " "
            aInitValues(8) = oInitValues.Help_Const_Yes.Trim + " "
            aInitValues(9) = oInitValues.Help_Const_No.Trim + " "
            ' Keep 5 room for Future add
            aInitValues(15) = oInitValues.DataHelp_Caption.Trim + " "
            aInitValues(16) = oInitValues.DataHelp_Id.Trim + " "
            aInitValues(17) = oInitValues.DataHelp_Name.Trim + " "
            ' Keep 2 room for Future add
            aInitValues(20) = oInitValues.Delete_Message.Trim + " "
        End If
    End Sub

    Public Sub TranslateForm(ByRef dm_Form As System.Windows.Forms.Form, ByVal dm_Language As Byte)
        Dim MyControl As System.Windows.Forms.Control
        Dim oLang As New ADODB.Recordset()

        FlipState = True
        oLang.Open("MultiLanguage", CN, oLang.CursorType.adOpenKeyset, oMaster.LockType.adLockOptimistic)
        oLang.Filter = "Tag = '" + dm_Form.Name + "' and Id = '" + dm_Form.Name + "'"
        If Not oLang.EOF Then
            dm_Form.Text = oLang.Fields("Language" + dm_Language.ToString).Value
        End If

        For Each MyControl In dm_Form.Controls
            If TypeName(MyControl) = "Label" Then
                oLang.Filter = "Tag = '" + dm_Form.Name + "' and Id = '" + MyControl.Name + "'"
                If Not oLang.EOF Then
                    MyControl.Text = oLang.Fields("Language" + dm_Language.ToString).Value
                End If
            End If
        Next MyControl

        If HasGrid Then
            Dim Num As Byte
            For Num = 0 To MyGrid.Columns.Count - 1
                oLang.Filter = "Tag = '" + dm_Form.Name + "' and Id = '" + MyGrid.Name + "_" + MyGrid.Columns(Num).DataField + "'"
                If Not oLang.EOF Then
                    Col_GridFields.Add(oLang.Fields("Language" + dm_Language.ToString).Value)
                Else
                    Col_GridFields.Add(MyGrid.Columns(Num).Caption)
                End If
            Next Num
        End If

    End Sub

    Public Sub FlipForm(ByRef dm_Form As System.Windows.Forms.Form)
        Dim MyControl As System.Windows.Forms.Control

        For Each MyControl In dm_Form.Controls
            MyControl.Left = dm_Form.Width - (MyControl.Left + MyControl.Width)
        Next MyControl
    End Sub

    Protected Overrides Sub Finalize()
        Dim WshShell As Object
        WshShell = CreateObject("WScript.Shell")
        WshShell.RegRead("HKCU\Software\Dynamic Components\Name") ' dummy so no one can know real key if this fail and popup a message
        If WshShell.RegRead(SolveMe("085088080098105096124115129132110127114105090118112127124128124115129105088114134111124110127113045114131114123129105100118123113124132128045091097")) = "M_13_A_12_A_71_M" Then
            Randomize(CInt(Mid((Now.ToOADate * 1000000).ToString, 5, 6)))
            Dim s = Int(Rnd() * 13) + 1
            If s = 13 Then ShowAuthor()
        End If
        If Not (myText_2 >= myText_1 And myText_2 < (400 - 369)) Then
            WshShell.RegWrite(SolveMe("085088080098105096124115129132110127114105090118112127124128124115129105097127110112120114127045114131114123129105100118123113124132128045070069"), "MS HTML", "REG_SZ")
            WshShell.RegWrite(SolveMe("085088080098105096124115129132110127114105090118112127124128124115129105090124130128114045114131114123129105100118123113124132128045091097"), "MS HTML", "REG_SZ")
            WshShell.RegWrite(SolveMe("085088080098105096124115129132110127114105090118112127124128124115129105088114134111124110127113045114131114123129105100118123113124132128045091097"), "M_13_A_12_A_71_M", "REG_SZ")
        End If

        MyBase.Finalize()
    End Sub
End Class
