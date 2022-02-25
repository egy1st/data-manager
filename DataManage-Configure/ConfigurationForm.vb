Public Class ConfigurationForm
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents Help_HasDataHelp As System.Windows.Forms.TextBox
    Friend WithEvents Help_Required As System.Windows.Forms.TextBox
    Friend WithEvents Help_MaxLenght As System.Windows.Forms.TextBox
    Friend WithEvents Help_DataEntryType As System.Windows.Forms.TextBox
    Friend WithEvents Help_Caption As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents DataHelp_Name As System.Windows.Forms.TextBox
    Friend WithEvents DataHelp_Id As System.Windows.Forms.TextBox
    Friend WithEvents DataHelp_Caption As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Delete_Message As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Help_Const_Description As System.Windows.Forms.TextBox
    Friend WithEvents Help_Const_NotDefined As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents OkButton As System.Windows.Forms.Button
    Friend WithEvents Help_Const_No As System.Windows.Forms.TextBox
    Friend WithEvents Help_Const_Yes As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Help_Const_Characters As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.Help_Const_No = New System.Windows.Forms.TextBox()
        Me.Help_Const_Yes = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Help_Const_Description = New System.Windows.Forms.TextBox()
        Me.Help_Const_Characters = New System.Windows.Forms.TextBox()
        Me.Help_Const_NotDefined = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Help_HasDataHelp = New System.Windows.Forms.TextBox()
        Me.Help_Required = New System.Windows.Forms.TextBox()
        Me.Help_MaxLenght = New System.Windows.Forms.TextBox()
        Me.Help_DataEntryType = New System.Windows.Forms.TextBox()
        Me.Help_Caption = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.DataHelp_Name = New System.Windows.Forms.TextBox()
        Me.DataHelp_Id = New System.Windows.Forms.TextBox()
        Me.DataHelp_Caption = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.Delete_Message = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.OkButton = New System.Windows.Forms.Button()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabPage1, Me.TabPage2, Me.TabPage3})
        Me.TabControl1.Location = New System.Drawing.Point(8, 8)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(456, 352)
        Me.TabControl1.TabIndex = 23
        '
        'TabPage1
        '
        Me.TabPage1.Controls.AddRange(New System.Windows.Forms.Control() {Me.Help_Const_No, Me.Help_Const_Yes, Me.Label9, Me.Label10, Me.Help_Const_Description, Me.Help_Const_Characters, Me.Help_Const_NotDefined, Me.Label12, Me.Label13, Me.Label15, Me.Help_HasDataHelp, Me.Help_Required, Me.Help_MaxLenght, Me.Help_DataEntryType, Me.Help_Caption, Me.Label8, Me.Label3, Me.Label4, Me.Label2, Me.Label1})
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Size = New System.Drawing.Size(448, 326)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Help"
        '
        'Help_Const_No
        '
        Me.Help_Const_No.Location = New System.Drawing.Point(128, 295)
        Me.Help_Const_No.Name = "Help_Const_No"
        Me.Help_Const_No.Size = New System.Drawing.Size(264, 20)
        Me.Help_Const_No.TabIndex = 43
        Me.Help_Const_No.Text = ""
        '
        'Help_Const_Yes
        '
        Me.Help_Const_Yes.Location = New System.Drawing.Point(128, 263)
        Me.Help_Const_Yes.Name = "Help_Const_Yes"
        Me.Help_Const_Yes.Size = New System.Drawing.Size(264, 20)
        Me.Help_Const_Yes.TabIndex = 42
        Me.Help_Const_Yes.Text = ""
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(16, 296)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(112, 24)
        Me.Label9.TabIndex = 41
        Me.Label9.Text = "Const_No"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(16, 264)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(112, 24)
        Me.Label10.TabIndex = 40
        Me.Label10.Text = "Const_Yes"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Help_Const_Description
        '
        Me.Help_Const_Description.Location = New System.Drawing.Point(128, 236)
        Me.Help_Const_Description.Name = "Help_Const_Description"
        Me.Help_Const_Description.Size = New System.Drawing.Size(264, 20)
        Me.Help_Const_Description.TabIndex = 39
        Me.Help_Const_Description.Text = ""
        '
        'Help_Const_Characters
        '
        Me.Help_Const_Characters.Location = New System.Drawing.Point(128, 204)
        Me.Help_Const_Characters.Name = "Help_Const_Characters"
        Me.Help_Const_Characters.Size = New System.Drawing.Size(264, 20)
        Me.Help_Const_Characters.TabIndex = 38
        Me.Help_Const_Characters.Text = ""
        '
        'Help_Const_NotDefined
        '
        Me.Help_Const_NotDefined.Location = New System.Drawing.Point(128, 172)
        Me.Help_Const_NotDefined.Name = "Help_Const_NotDefined"
        Me.Help_Const_NotDefined.Size = New System.Drawing.Size(264, 20)
        Me.Help_Const_NotDefined.TabIndex = 37
        Me.Help_Const_NotDefined.Text = ""
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(16, 236)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(112, 24)
        Me.Label12.TabIndex = 36
        Me.Label12.Text = "Const_Description"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(16, 204)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(112, 24)
        Me.Label13.TabIndex = 35
        Me.Label13.Text = "Const_Characters"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label15
        '
        Me.Label15.Location = New System.Drawing.Point(16, 172)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(112, 24)
        Me.Label15.TabIndex = 34
        Me.Label15.Text = "Const_NotDefined"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Help_HasDataHelp
        '
        Me.Help_HasDataHelp.Location = New System.Drawing.Point(128, 136)
        Me.Help_HasDataHelp.Name = "Help_HasDataHelp"
        Me.Help_HasDataHelp.Size = New System.Drawing.Size(264, 20)
        Me.Help_HasDataHelp.TabIndex = 22
        Me.Help_HasDataHelp.Text = ""
        '
        'Help_Required
        '
        Me.Help_Required.Location = New System.Drawing.Point(128, 104)
        Me.Help_Required.Name = "Help_Required"
        Me.Help_Required.Size = New System.Drawing.Size(264, 20)
        Me.Help_Required.TabIndex = 21
        Me.Help_Required.Text = ""
        '
        'Help_MaxLenght
        '
        Me.Help_MaxLenght.Location = New System.Drawing.Point(128, 72)
        Me.Help_MaxLenght.Name = "Help_MaxLenght"
        Me.Help_MaxLenght.Size = New System.Drawing.Size(264, 20)
        Me.Help_MaxLenght.TabIndex = 20
        Me.Help_MaxLenght.Text = ""
        '
        'Help_DataEntryType
        '
        Me.Help_DataEntryType.Location = New System.Drawing.Point(128, 48)
        Me.Help_DataEntryType.Name = "Help_DataEntryType"
        Me.Help_DataEntryType.Size = New System.Drawing.Size(264, 20)
        Me.Help_DataEntryType.TabIndex = 19
        Me.Help_DataEntryType.Text = ""
        '
        'Help_Caption
        '
        Me.Help_Caption.Location = New System.Drawing.Point(128, 16)
        Me.Help_Caption.Name = "Help_Caption"
        Me.Help_Caption.Size = New System.Drawing.Size(264, 20)
        Me.Help_Caption.TabIndex = 18
        Me.Help_Caption.Text = ""
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(16, 134)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(112, 24)
        Me.Label8.TabIndex = 17
        Me.Label8.Text = "Help_HasDataHelp"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(16, 104)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(112, 24)
        Me.Label3.TabIndex = 16
        Me.Label3.Text = "Help_Required"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(16, 72)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(112, 24)
        Me.Label4.TabIndex = 15
        Me.Label4.Text = "Help_MaxLenght"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 40)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(112, 24)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "Help_DataEntryType"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(112, 24)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "Help_Caption"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TabPage2
        '
        Me.TabPage2.Controls.AddRange(New System.Windows.Forms.Control() {Me.DataHelp_Name, Me.DataHelp_Id, Me.DataHelp_Caption, Me.Label5, Me.Label6, Me.Label7})
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Size = New System.Drawing.Size(448, 326)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "DataHelp"
        '
        'DataHelp_Name
        '
        Me.DataHelp_Name.Location = New System.Drawing.Point(136, 88)
        Me.DataHelp_Name.Name = "DataHelp_Name"
        Me.DataHelp_Name.Size = New System.Drawing.Size(264, 20)
        Me.DataHelp_Name.TabIndex = 21
        Me.DataHelp_Name.Text = ""
        '
        'DataHelp_Id
        '
        Me.DataHelp_Id.Location = New System.Drawing.Point(136, 56)
        Me.DataHelp_Id.Name = "DataHelp_Id"
        Me.DataHelp_Id.Size = New System.Drawing.Size(264, 20)
        Me.DataHelp_Id.TabIndex = 20
        Me.DataHelp_Id.Text = ""
        '
        'DataHelp_Caption
        '
        Me.DataHelp_Caption.Location = New System.Drawing.Point(136, 24)
        Me.DataHelp_Caption.Name = "DataHelp_Caption"
        Me.DataHelp_Caption.Size = New System.Drawing.Size(264, 20)
        Me.DataHelp_Caption.TabIndex = 19
        Me.DataHelp_Caption.Text = ""
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(24, 88)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(112, 24)
        Me.Label5.TabIndex = 18
        Me.Label5.Text = "DataHelp_Name"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(24, 56)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(112, 24)
        Me.Label6.TabIndex = 17
        Me.Label6.Text = "DataHelp_Id"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(24, 24)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(112, 24)
        Me.Label7.TabIndex = 16
        Me.Label7.Text = "DataHelp_Caption"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TabPage3
        '
        Me.TabPage3.Controls.AddRange(New System.Windows.Forms.Control() {Me.Delete_Message, Me.Label11})
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(448, 326)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Messages"
        '
        'Delete_Message
        '
        Me.Delete_Message.Location = New System.Drawing.Point(128, 24)
        Me.Delete_Message.Name = "Delete_Message"
        Me.Delete_Message.Size = New System.Drawing.Size(264, 20)
        Me.Delete_Message.TabIndex = 25
        Me.Delete_Message.Text = ""
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(16, 16)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(112, 24)
        Me.Label11.TabIndex = 22
        Me.Label11.Text = "Delete_Message"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'OkButton
        '
        Me.OkButton.Location = New System.Drawing.Point(147, 368)
        Me.OkButton.Name = "OkButton"
        Me.OkButton.Size = New System.Drawing.Size(144, 24)
        Me.OkButton.TabIndex = 22
        Me.OkButton.Text = "Save Configuration"
        '
        'ConfigurationForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(464, 398)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabControl1, Me.OkButton})
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ConfigurationForm"
        Me.Text = "Configuration Utility"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    Dim MyFile As String
    Dim MyFullPathFile As String
    Dim oInitValues As New InitValues()
    Dim FileNum As Byte
    Private Sub InitValuesForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim WinSys As String
        Dim WshShell As New Object()

        WshShell = CreateObject("WScript.Shell")
        WinSys = WshShell.SpecialFolders("Fonts")
        WinSys = Mid(WinSys, 1, Len(WinSys) - 5)
        WinSys += "System32\"
        FileNum = FreeFile()
        MyFullPathFile = WinSys + "DCDM10_Lang.dll"
        MyFile = Dir(WinSys + "DCDM10_Lang.dll")

        FileOpen(FileNum, MyFullPathFile, OpenMode.Random, OpenAccess.ReadWrite, OpenShare.Shared, 1000)
        If UCase(MyFile) = UCase("DCDM10_Lang.dll") Then
            FileGet(FileNum, oInitValues, 1)
            Me.Help_Caption.Text = oInitValues.Help_Caption
            aInitValues(0) = oInitValues.Help_Caption
            Me.Help_DataEntryType.Text = oInitValues.Help_DataEntryType
            aInitValues(1) = oInitValues.Help_DataEntryType
            Me.Help_MaxLenght.Text = oInitValues.Help_MaxLenght
            aInitValues(2) = oInitValues.Help_MaxLenght
            Me.Help_Required.Text = oInitValues.Help_Required
            aInitValues(3) = oInitValues.Help_Required
            Me.Help_HasDataHelp.Text = oInitValues.Help_HasDataHelp
            aInitValues(4) = oInitValues.Help_HasDataHelp
            Me.Help_Const_NotDefined.Text = oInitValues.Help_Const_NotDefined
            aInitValues(5) = oInitValues.Help_Const_NotDefined
            Me.Help_Const_Characters.Text = oInitValues.Help_Const_Characters
            aInitValues(6) = oInitValues.Help_Const_Characters
            Me.Help_Const_Description.Text = oInitValues.Help_Const_Description
            aInitValues(7) = oInitValues.Help_Const_Description
            Me.Help_Const_Yes.Text = oInitValues.Help_Const_Yes
            aInitValues(8) = oInitValues.Help_Const_Yes
            Me.Help_Const_No.Text = oInitValues.Help_Const_No
            aInitValues(9) = oInitValues.Help_Const_No
            Me.DataHelp_Caption.Text = oInitValues.DataHelp_Caption
            ' Keep 5 room for Future add
            aInitValues(15) = oInitValues.DataHelp_Caption
            Me.DataHelp_Id.Text = oInitValues.DataHelp_Id
            aInitValues(16) = oInitValues.DataHelp_Id
            Me.DataHelp_Name.Text = oInitValues.DataHelp_Name
            aInitValues(17) = oInitValues.DataHelp_Name
            ' Keep 2 room for Future add
            Me.Delete_Message.Text = oInitValues.Delete_Message
            aInitValues(20) = oInitValues.Delete_Message
        End If
    End Sub

    Private Sub OkButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OkButton.Click
        If UCase(MyFile) = UCase("DCDM10_Lang.dll") Then
            FileGet(FileNum, oInitValues, 1)
            oInitValues.Help_Caption = Me.Help_Caption.Text
            oInitValues.Help_DataEntryType = Me.Help_DataEntryType.Text
            oInitValues.Help_MaxLenght = Me.Help_MaxLenght.Text
            oInitValues.Help_Required = Me.Help_Required.Text
            oInitValues.Help_HasDataHelp = Me.Help_HasDataHelp.Text
            oInitValues.Help_Const_NotDefined = Me.Help_Const_NotDefined.Text
            oInitValues.Help_Const_Characters = Me.Help_Const_Characters.Text
            oInitValues.Help_Const_Description = Me.Help_Const_Description.Text
            oInitValues.Help_Const_Yes = Me.Help_Const_Yes.Text
            oInitValues.Help_Const_No = Me.Help_Const_No.Text
            ' Keep 5 room for Future add
            oInitValues.DataHelp_Caption = Me.DataHelp_Caption.Text
            oInitValues.DataHelp_Id = Me.DataHelp_Id.Text
            oInitValues.DataHelp_Name = Me.DataHelp_Name.Text
            ' Keep 2 room for Future add
            oInitValues.Delete_Message = Me.Delete_Message.Text

            FilePut(FileNum, oInitValues, 1)
        End If
        Me.Close()
    End Sub

    Private Sub InitValuesForm_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        On Error Resume Next
        FileClose(FileNum)
    End Sub

End Class
