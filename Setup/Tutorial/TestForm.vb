Public Class TestForm
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
    Friend WithEvents AxDataGrid1 As AxMSDataGridLib.AxDataGrid
    Public WithEvents Freight As System.Windows.Forms.TextBox
    Public WithEvents Freight_Label As System.Windows.Forms.Label
    Public WithEvents xCompanyName As System.Windows.Forms.TextBox
    Public WithEvents Orderdate_Label As System.Windows.Forms.Label
    Public WithEvents OrderDate As System.Windows.Forms.TextBox
    Public WithEvents ShipVia As System.Windows.Forms.TextBox
    Public WithEvents CustomerID As System.Windows.Forms.TextBox
    Public WithEvents SearchButton As System.Windows.Forms.Button
    Public WithEvents CustomerID_Label As System.Windows.Forms.Label
    Public WithEvents ShipVia_Label As System.Windows.Forms.Label
    Public WithEvents xCustomerName As System.Windows.Forms.TextBox
    Public WithEvents DeleteButton As System.Windows.Forms.Button
    Public WithEvents OkButton As System.Windows.Forms.Button
    Public WithEvents ExitButton As System.Windows.Forms.Button
    Public WithEvents NewButton As System.Windows.Forms.Button
    Public WithEvents NextButton As System.Windows.Forms.Button
    Public WithEvents FirstButton As System.Windows.Forms.Button
    Public WithEvents PrevButton As System.Windows.Forms.Button
    Public WithEvents LastButton As System.Windows.Forms.Button
    Public WithEvents OrderId_Label As System.Windows.Forms.Label
    Public WithEvents OrderID As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(TestForm))
        Me.AxDataGrid1 = New AxMSDataGridLib.AxDataGrid()
        Me.Freight = New System.Windows.Forms.TextBox()
        Me.Freight_Label = New System.Windows.Forms.Label()
        Me.xCompanyName = New System.Windows.Forms.TextBox()
        Me.Orderdate_Label = New System.Windows.Forms.Label()
        Me.OrderDate = New System.Windows.Forms.TextBox()
        Me.ShipVia = New System.Windows.Forms.TextBox()
        Me.CustomerID = New System.Windows.Forms.TextBox()
        Me.SearchButton = New System.Windows.Forms.Button()
        Me.CustomerID_Label = New System.Windows.Forms.Label()
        Me.ShipVia_Label = New System.Windows.Forms.Label()
        Me.xCustomerName = New System.Windows.Forms.TextBox()
        Me.DeleteButton = New System.Windows.Forms.Button()
        Me.OkButton = New System.Windows.Forms.Button()
        Me.ExitButton = New System.Windows.Forms.Button()
        Me.NewButton = New System.Windows.Forms.Button()
        Me.NextButton = New System.Windows.Forms.Button()
        Me.FirstButton = New System.Windows.Forms.Button()
        Me.PrevButton = New System.Windows.Forms.Button()
        Me.LastButton = New System.Windows.Forms.Button()
        Me.OrderId_Label = New System.Windows.Forms.Label()
        Me.OrderID = New System.Windows.Forms.TextBox()
        CType(Me.AxDataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'AxDataGrid1
        '
        Me.AxDataGrid1.DataSource = Nothing
        Me.AxDataGrid1.Location = New System.Drawing.Point(24, 146)
        Me.AxDataGrid1.Name = "AxDataGrid1"
        Me.AxDataGrid1.OcxState = CType(resources.GetObject("AxDataGrid1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxDataGrid1.Size = New System.Drawing.Size(456, 120)
        Me.AxDataGrid1.TabIndex = 68
        '
        'Freight
        '
        Me.Freight.AcceptsReturn = True
        Me.Freight.AutoSize = False
        Me.Freight.BackColor = System.Drawing.SystemColors.Window
        Me.Freight.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Freight.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Freight.Location = New System.Drawing.Point(132, 105)
        Me.Freight.MaxLength = 10
        Me.Freight.Name = "Freight"
        Me.Freight.Size = New System.Drawing.Size(84, 25)
        Me.Freight.TabIndex = 67
        Me.Freight.Text = ""
        '
        'Freight_Label
        '
        Me.Freight_Label.BackColor = System.Drawing.SystemColors.Control
        Me.Freight_Label.Cursor = System.Windows.Forms.Cursors.Default
        Me.Freight_Label.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Freight_Label.ForeColor = System.Drawing.Color.Blue
        Me.Freight_Label.Location = New System.Drawing.Point(20, 105)
        Me.Freight_Label.Name = "Freight_Label"
        Me.Freight_Label.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Freight_Label.Size = New System.Drawing.Size(112, 25)
        Me.Freight_Label.TabIndex = 84
        Me.Freight_Label.Text = "Freight"
        Me.Freight_Label.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'xCompanyName
        '
        Me.xCompanyName.AcceptsReturn = True
        Me.xCompanyName.AutoSize = False
        Me.xCompanyName.BackColor = System.Drawing.SystemColors.Window
        Me.xCompanyName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.xCompanyName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.xCompanyName.Location = New System.Drawing.Point(216, 73)
        Me.xCompanyName.MaxLength = 0
        Me.xCompanyName.Name = "xCompanyName"
        Me.xCompanyName.ReadOnly = True
        Me.xCompanyName.Size = New System.Drawing.Size(264, 25)
        Me.xCompanyName.TabIndex = 83
        Me.xCompanyName.Text = ""
        '
        'Orderdate_Label
        '
        Me.Orderdate_Label.BackColor = System.Drawing.SystemColors.Control
        Me.Orderdate_Label.Cursor = System.Windows.Forms.Cursors.Default
        Me.Orderdate_Label.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Orderdate_Label.ForeColor = System.Drawing.Color.Blue
        Me.Orderdate_Label.Location = New System.Drawing.Point(280, 10)
        Me.Orderdate_Label.Name = "Orderdate_Label"
        Me.Orderdate_Label.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Orderdate_Label.Size = New System.Drawing.Size(112, 25)
        Me.Orderdate_Label.TabIndex = 82
        Me.Orderdate_Label.Text = "Order Date"
        Me.Orderdate_Label.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'OrderDate
        '
        Me.OrderDate.AcceptsReturn = True
        Me.OrderDate.AutoSize = False
        Me.OrderDate.BackColor = System.Drawing.SystemColors.Window
        Me.OrderDate.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.OrderDate.ForeColor = System.Drawing.SystemColors.WindowText
        Me.OrderDate.Location = New System.Drawing.Point(392, 10)
        Me.OrderDate.MaxLength = 10
        Me.OrderDate.Name = "OrderDate"
        Me.OrderDate.Size = New System.Drawing.Size(88, 25)
        Me.OrderDate.TabIndex = 64
        Me.OrderDate.Text = ""
        '
        'ShipVia
        '
        Me.ShipVia.AcceptsReturn = True
        Me.ShipVia.AutoSize = False
        Me.ShipVia.BackColor = System.Drawing.SystemColors.Window
        Me.ShipVia.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.ShipVia.ForeColor = System.Drawing.SystemColors.WindowText
        Me.ShipVia.Location = New System.Drawing.Point(132, 72)
        Me.ShipVia.MaxLength = 2
        Me.ShipVia.Name = "ShipVia"
        Me.ShipVia.Size = New System.Drawing.Size(84, 25)
        Me.ShipVia.TabIndex = 66
        Me.ShipVia.Text = ""
        '
        'CustomerID
        '
        Me.CustomerID.AcceptsReturn = True
        Me.CustomerID.AutoSize = False
        Me.CustomerID.BackColor = System.Drawing.SystemColors.Window
        Me.CustomerID.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.CustomerID.ForeColor = System.Drawing.SystemColors.WindowText
        Me.CustomerID.Location = New System.Drawing.Point(132, 40)
        Me.CustomerID.MaxLength = 3
        Me.CustomerID.Name = "CustomerID"
        Me.CustomerID.Size = New System.Drawing.Size(84, 25)
        Me.CustomerID.TabIndex = 65
        Me.CustomerID.Text = ""
        '
        'SearchButton
        '
        Me.SearchButton.BackColor = System.Drawing.SystemColors.Control
        Me.SearchButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.SearchButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.SearchButton.Image = CType(resources.GetObject("SearchButton.Image"), System.Drawing.Bitmap)
        Me.SearchButton.Location = New System.Drawing.Point(358, 282)
        Me.SearchButton.Name = "SearchButton"
        Me.SearchButton.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.SearchButton.Size = New System.Drawing.Size(44, 41)
        Me.SearchButton.TabIndex = 76
        Me.SearchButton.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'CustomerID_Label
        '
        Me.CustomerID_Label.BackColor = System.Drawing.SystemColors.Control
        Me.CustomerID_Label.Cursor = System.Windows.Forms.Cursors.Default
        Me.CustomerID_Label.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.CustomerID_Label.ForeColor = System.Drawing.Color.Blue
        Me.CustomerID_Label.Location = New System.Drawing.Point(20, 42)
        Me.CustomerID_Label.Name = "CustomerID_Label"
        Me.CustomerID_Label.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.CustomerID_Label.Size = New System.Drawing.Size(112, 25)
        Me.CustomerID_Label.TabIndex = 81
        Me.CustomerID_Label.Text = "Customer Id"
        Me.CustomerID_Label.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ShipVia_Label
        '
        Me.ShipVia_Label.BackColor = System.Drawing.SystemColors.Control
        Me.ShipVia_Label.Cursor = System.Windows.Forms.Cursors.Default
        Me.ShipVia_Label.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.ShipVia_Label.ForeColor = System.Drawing.Color.Blue
        Me.ShipVia_Label.Location = New System.Drawing.Point(20, 74)
        Me.ShipVia_Label.Name = "ShipVia_Label"
        Me.ShipVia_Label.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShipVia_Label.Size = New System.Drawing.Size(112, 25)
        Me.ShipVia_Label.TabIndex = 80
        Me.ShipVia_Label.Text = "Ship Via"
        Me.ShipVia_Label.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'xCustomerName
        '
        Me.xCustomerName.AcceptsReturn = True
        Me.xCustomerName.AutoSize = False
        Me.xCustomerName.BackColor = System.Drawing.SystemColors.Window
        Me.xCustomerName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.xCustomerName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.xCustomerName.Location = New System.Drawing.Point(217, 41)
        Me.xCustomerName.MaxLength = 0
        Me.xCustomerName.Name = "xCustomerName"
        Me.xCustomerName.ReadOnly = True
        Me.xCustomerName.Size = New System.Drawing.Size(263, 25)
        Me.xCustomerName.TabIndex = 78
        Me.xCustomerName.Text = ""
        '
        'DeleteButton
        '
        Me.DeleteButton.BackColor = System.Drawing.SystemColors.Control
        Me.DeleteButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.DeleteButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.DeleteButton.Image = CType(resources.GetObject("DeleteButton.Image"), System.Drawing.Bitmap)
        Me.DeleteButton.Location = New System.Drawing.Point(313, 282)
        Me.DeleteButton.Name = "DeleteButton"
        Me.DeleteButton.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.DeleteButton.Size = New System.Drawing.Size(44, 41)
        Me.DeleteButton.TabIndex = 75
        Me.DeleteButton.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'OkButton
        '
        Me.OkButton.BackColor = System.Drawing.SystemColors.Control
        Me.OkButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.OkButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.OkButton.Image = CType(resources.GetObject("OkButton.Image"), System.Drawing.Bitmap)
        Me.OkButton.Location = New System.Drawing.Point(223, 282)
        Me.OkButton.Name = "OkButton"
        Me.OkButton.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.OkButton.Size = New System.Drawing.Size(44, 41)
        Me.OkButton.TabIndex = 73
        Me.OkButton.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'ExitButton
        '
        Me.ExitButton.BackColor = System.Drawing.SystemColors.Control
        Me.ExitButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.ExitButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.ExitButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ExitButton.Image = CType(resources.GetObject("ExitButton.Image"), System.Drawing.Bitmap)
        Me.ExitButton.Location = New System.Drawing.Point(436, 285)
        Me.ExitButton.Name = "ExitButton"
        Me.ExitButton.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ExitButton.Size = New System.Drawing.Size(44, 41)
        Me.ExitButton.TabIndex = 77
        Me.ExitButton.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'NewButton
        '
        Me.NewButton.BackColor = System.Drawing.SystemColors.Control
        Me.NewButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.NewButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.NewButton.Image = CType(resources.GetObject("NewButton.Image"), System.Drawing.Bitmap)
        Me.NewButton.Location = New System.Drawing.Point(268, 282)
        Me.NewButton.Name = "NewButton"
        Me.NewButton.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.NewButton.Size = New System.Drawing.Size(44, 41)
        Me.NewButton.TabIndex = 74
        Me.NewButton.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'NextButton
        '
        Me.NextButton.BackColor = System.Drawing.SystemColors.Control
        Me.NextButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.NextButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.NextButton.Image = CType(resources.GetObject("NextButton.Image"), System.Drawing.Bitmap)
        Me.NextButton.Location = New System.Drawing.Point(104, 282)
        Me.NextButton.Name = "NextButton"
        Me.NextButton.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.NextButton.Size = New System.Drawing.Size(44, 41)
        Me.NextButton.TabIndex = 71
        Me.NextButton.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'FirstButton
        '
        Me.FirstButton.BackColor = System.Drawing.SystemColors.Control
        Me.FirstButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.FirstButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.FirstButton.Image = CType(resources.GetObject("FirstButton.Image"), System.Drawing.Bitmap)
        Me.FirstButton.Location = New System.Drawing.Point(16, 282)
        Me.FirstButton.Name = "FirstButton"
        Me.FirstButton.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.FirstButton.Size = New System.Drawing.Size(44, 41)
        Me.FirstButton.TabIndex = 69
        Me.FirstButton.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'PrevButton
        '
        Me.PrevButton.BackColor = System.Drawing.SystemColors.Control
        Me.PrevButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.PrevButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.PrevButton.Image = CType(resources.GetObject("PrevButton.Image"), System.Drawing.Bitmap)
        Me.PrevButton.Location = New System.Drawing.Point(60, 282)
        Me.PrevButton.Name = "PrevButton"
        Me.PrevButton.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.PrevButton.Size = New System.Drawing.Size(44, 41)
        Me.PrevButton.TabIndex = 70
        Me.PrevButton.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'LastButton
        '
        Me.LastButton.BackColor = System.Drawing.SystemColors.Control
        Me.LastButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.LastButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LastButton.Image = CType(resources.GetObject("LastButton.Image"), System.Drawing.Bitmap)
        Me.LastButton.Location = New System.Drawing.Point(149, 282)
        Me.LastButton.Name = "LastButton"
        Me.LastButton.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.LastButton.Size = New System.Drawing.Size(44, 41)
        Me.LastButton.TabIndex = 72
        Me.LastButton.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'OrderId_Label
        '
        Me.OrderId_Label.BackColor = System.Drawing.SystemColors.Control
        Me.OrderId_Label.Cursor = System.Windows.Forms.Cursors.Default
        Me.OrderId_Label.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.OrderId_Label.ForeColor = System.Drawing.Color.Blue
        Me.OrderId_Label.Location = New System.Drawing.Point(20, 10)
        Me.OrderId_Label.Name = "OrderId_Label"
        Me.OrderId_Label.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.OrderId_Label.Size = New System.Drawing.Size(112, 25)
        Me.OrderId_Label.TabIndex = 79
        Me.OrderId_Label.Text = "Order Id"
        Me.OrderId_Label.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'OrderID
        '
        Me.OrderID.AcceptsReturn = True
        Me.OrderID.AutoSize = False
        Me.OrderID.BackColor = System.Drawing.SystemColors.Window
        Me.OrderID.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.OrderID.ForeColor = System.Drawing.SystemColors.WindowText
        Me.OrderID.Location = New System.Drawing.Point(132, 8)
        Me.OrderID.MaxLength = 5
        Me.OrderID.Name = "OrderID"
        Me.OrderID.Size = New System.Drawing.Size(84, 26)
        Me.OrderID.TabIndex = 63
        Me.OrderID.Text = ""
        '
        'TestForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(496, 334)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.AxDataGrid1, Me.Freight, Me.Freight_Label, Me.xCompanyName, Me.Orderdate_Label, Me.OrderDate, Me.ShipVia, Me.CustomerID, Me.SearchButton, Me.CustomerID_Label, Me.ShipVia_Label, Me.xCustomerName, Me.DeleteButton, Me.OkButton, Me.ExitButton, Me.NewButton, Me.NextButton, Me.FirstButton, Me.PrevButton, Me.LastButton, Me.OrderId_Label, Me.OrderID})
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "TestForm"
        Me.Tag = "Orders Form"
        Me.Text = "TestForm"
        CType(Me.AxDataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim DM As New DynamicComponents.DataManger()
    Dim aImage(26) As String
    Dim CN As New ADODB.Connection()
    Dim oMaster As New ADODB.Recordset()
    Dim oDetails As New ADODB.Recordset()
    Dim oAccess As New Access.Application()
    Dim DAO_DBEngine As New DAO.DBEngine()

    Private Sub TestForm_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        'establish DSN
        oAccess.DBEngine.RegisterDatabase("DCDM_Nwind", "Microsoft Access Driver (*.mdb)", True, "DBQ=" & VB6.GetPath & "\Nwind.mdb")
        PopulateaImage() ' define images used with buttons
        CN.Open("DSN=DCDM_NWind")
        oMaster.Open("Orders", CN, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        oDetails.Open("OrderDetails", CN, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        DM.InitForm(CN, Me, oMaster, AxDataGrid1, oDetails) 'Must Be your first declaration
        DM.PrepareImageButtons(aImage, VB6.GetPath + "\icons\", False)
        DM.NavigationButtons("FirstButton", "PrevButton", "NextButton", "LastButton")
        DM.ManipulationButtons("OkButton", "NewButton", "DeleteButton", "ExitButton", "SearchButton")
        DM.KeyFields("OrderId")
        DM.SetLink("OrderId", "OrderId")
        DM.AddRelatedValue("Customers", "CustomerID", "CustomerID", "CustomerName", "xCustomerName", 3)
        DM.AddRelatedValue("Shippers", "ShipperId", "ShipVia", "CompanyName", "xCompanyName", 2)
        DM.AddGridRelatedValue("Products", "ProductID", "ProductID", "ProductName", "ProductName", 2)
        DM.KeyLeaveField(oMaster, "OrderId", 5)
        DM.RequiredFields("OrderId+OrderDate+CustomerId")
        DM.NumericFields("CustomerID", "OrderId", "ShipVia")
        DM.DecimalFields("Freight")
        DM.DateFields("OrderDate")
        DM.DecimalPlaces(2)
        DM.EnableReturnKey(True)
        'DM.Right2Left(True)      'For Eastern Languages Support 
        'DM.FlipForm(Me)          'For Eastern Languages Support 
        'DM.TranslateForm(Me, 1)  'For MultiLanguages application Support 
        DM.PopulateForm(Me, oMaster, AxDataGrid1, oDetails) 'Must be your last declaration
    End Sub

    Private Sub PopulateaImage()
        aImage(0) = "First.ico"
        aImage(1) = "FirstOver.ico"
        aImage(2) = "FirstDown.ico"
        aImage(3) = "Previous.ico"
        aImage(4) = "PreviousOver.ico"
        aImage(5) = "PreviousDown.ico"
        aImage(6) = "Next.ico"
        aImage(7) = "NextOver.ico"
        aImage(8) = "NextDown.ico"
        aImage(9) = "Last.ico"
        aImage(10) = "LastOver.ico"
        aImage(11) = "LastDown.ico"
        aImage(12) = "Ok.ico"
        aImage(13) = "OkOver.ico"
        aImage(14) = "OkDown.ico"
        aImage(15) = "New.ico"
        aImage(16) = "NewOver.ico"
        aImage(17) = "NewDown.ico"
        aImage(18) = "Delete.ico"
        aImage(19) = "DeleteOver.ico"
        aImage(20) = "DeleteDown.ico"
        aImage(21) = "Exit.ico"
        aImage(22) = "ExitOver.ico"
        aImage(23) = "ExitDown.ico"
        aImage(24) = "Search.ico"
        aImage(25) = "SearchOver.ico"
        aImage(26) = "SearchDown.ico"
    End Sub

    Private Sub ExitButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitButton.Click
        oAccess.Quit()
    End Sub
End Class
