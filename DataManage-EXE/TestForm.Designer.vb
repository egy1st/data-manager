<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TestForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(TestForm))
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
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Freight
        '
        Me.Freight.AcceptsReturn = True
        Me.Freight.BackColor = System.Drawing.SystemColors.Window
        Me.Freight.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Freight.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Freight.Location = New System.Drawing.Point(206, 105)
        Me.Freight.MaxLength = 10
        Me.Freight.Name = "Freight"
        Me.Freight.Size = New System.Drawing.Size(84, 20)
        Me.Freight.TabIndex = 67
        '
        'Freight_Label
        '
        Me.Freight_Label.BackColor = System.Drawing.SystemColors.Control
        Me.Freight_Label.Cursor = System.Windows.Forms.Cursors.Default
        Me.Freight_Label.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Freight_Label.ForeColor = System.Drawing.Color.Blue
        Me.Freight_Label.Location = New System.Drawing.Point(94, 105)
        Me.Freight_Label.Name = "Freight_Label"
        Me.Freight_Label.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Freight_Label.Size = New System.Drawing.Size(112, 25)
        Me.Freight_Label.TabIndex = 83
        Me.Freight_Label.Text = "Freight"
        Me.Freight_Label.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'xCompanyName
        '
        Me.xCompanyName.AcceptsReturn = True
        Me.xCompanyName.BackColor = System.Drawing.SystemColors.Window
        Me.xCompanyName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.xCompanyName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.xCompanyName.Location = New System.Drawing.Point(290, 73)
        Me.xCompanyName.MaxLength = 0
        Me.xCompanyName.Name = "xCompanyName"
        Me.xCompanyName.ReadOnly = True
        Me.xCompanyName.Size = New System.Drawing.Size(264, 20)
        Me.xCompanyName.TabIndex = 82
        '
        'Orderdate_Label
        '
        Me.Orderdate_Label.BackColor = System.Drawing.SystemColors.Control
        Me.Orderdate_Label.Cursor = System.Windows.Forms.Cursors.Default
        Me.Orderdate_Label.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Orderdate_Label.ForeColor = System.Drawing.Color.Blue
        Me.Orderdate_Label.Location = New System.Drawing.Point(354, 10)
        Me.Orderdate_Label.Name = "Orderdate_Label"
        Me.Orderdate_Label.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Orderdate_Label.Size = New System.Drawing.Size(112, 25)
        Me.Orderdate_Label.TabIndex = 81
        Me.Orderdate_Label.Text = "Order Date"
        Me.Orderdate_Label.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'OrderDate
        '
        Me.OrderDate.AcceptsReturn = True
        Me.OrderDate.BackColor = System.Drawing.SystemColors.Window
        Me.OrderDate.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.OrderDate.ForeColor = System.Drawing.SystemColors.WindowText
        Me.OrderDate.Location = New System.Drawing.Point(466, 10)
        Me.OrderDate.MaxLength = 10
        Me.OrderDate.Name = "OrderDate"
        Me.OrderDate.Size = New System.Drawing.Size(88, 20)
        Me.OrderDate.TabIndex = 64
        '
        'ShipVia
        '
        Me.ShipVia.AcceptsReturn = True
        Me.ShipVia.BackColor = System.Drawing.SystemColors.Window
        Me.ShipVia.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.ShipVia.ForeColor = System.Drawing.SystemColors.WindowText
        Me.ShipVia.Location = New System.Drawing.Point(206, 72)
        Me.ShipVia.MaxLength = 2
        Me.ShipVia.Name = "ShipVia"
        Me.ShipVia.Size = New System.Drawing.Size(84, 20)
        Me.ShipVia.TabIndex = 66
        '
        'CustomerID
        '
        Me.CustomerID.AcceptsReturn = True
        Me.CustomerID.BackColor = System.Drawing.SystemColors.Window
        Me.CustomerID.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.CustomerID.ForeColor = System.Drawing.SystemColors.WindowText
        Me.CustomerID.Location = New System.Drawing.Point(206, 40)
        Me.CustomerID.MaxLength = 3
        Me.CustomerID.Name = "CustomerID"
        Me.CustomerID.Size = New System.Drawing.Size(84, 20)
        Me.CustomerID.TabIndex = 65
        '
        'SearchButton
        '
        Me.SearchButton.BackColor = System.Drawing.SystemColors.Control
        Me.SearchButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.SearchButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.SearchButton.Image = CType(resources.GetObject("SearchButton.Image"), System.Drawing.Image)
        Me.SearchButton.Location = New System.Drawing.Point(419, 343)
        Me.SearchButton.Name = "SearchButton"
        Me.SearchButton.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.SearchButton.Size = New System.Drawing.Size(44, 41)
        Me.SearchButton.TabIndex = 75
        Me.SearchButton.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.SearchButton.UseVisualStyleBackColor = False
        '
        'CustomerID_Label
        '
        Me.CustomerID_Label.BackColor = System.Drawing.SystemColors.Control
        Me.CustomerID_Label.Cursor = System.Windows.Forms.Cursors.Default
        Me.CustomerID_Label.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.CustomerID_Label.ForeColor = System.Drawing.Color.Blue
        Me.CustomerID_Label.Location = New System.Drawing.Point(94, 42)
        Me.CustomerID_Label.Name = "CustomerID_Label"
        Me.CustomerID_Label.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.CustomerID_Label.Size = New System.Drawing.Size(112, 25)
        Me.CustomerID_Label.TabIndex = 80
        Me.CustomerID_Label.Text = "Customer Id"
        Me.CustomerID_Label.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ShipVia_Label
        '
        Me.ShipVia_Label.BackColor = System.Drawing.SystemColors.Control
        Me.ShipVia_Label.Cursor = System.Windows.Forms.Cursors.Default
        Me.ShipVia_Label.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.ShipVia_Label.ForeColor = System.Drawing.Color.Blue
        Me.ShipVia_Label.Location = New System.Drawing.Point(94, 74)
        Me.ShipVia_Label.Name = "ShipVia_Label"
        Me.ShipVia_Label.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShipVia_Label.Size = New System.Drawing.Size(112, 25)
        Me.ShipVia_Label.TabIndex = 79
        Me.ShipVia_Label.Text = "Ship Via"
        Me.ShipVia_Label.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'xCustomerName
        '
        Me.xCustomerName.AcceptsReturn = True
        Me.xCustomerName.BackColor = System.Drawing.SystemColors.Window
        Me.xCustomerName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.xCustomerName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.xCustomerName.Location = New System.Drawing.Point(291, 41)
        Me.xCustomerName.MaxLength = 0
        Me.xCustomerName.Name = "xCustomerName"
        Me.xCustomerName.ReadOnly = True
        Me.xCustomerName.Size = New System.Drawing.Size(263, 20)
        Me.xCustomerName.TabIndex = 77
        '
        'DeleteButton
        '
        Me.DeleteButton.BackColor = System.Drawing.SystemColors.Control
        Me.DeleteButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.DeleteButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.DeleteButton.Image = CType(resources.GetObject("DeleteButton.Image"), System.Drawing.Image)
        Me.DeleteButton.Location = New System.Drawing.Point(374, 343)
        Me.DeleteButton.Name = "DeleteButton"
        Me.DeleteButton.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.DeleteButton.Size = New System.Drawing.Size(44, 41)
        Me.DeleteButton.TabIndex = 74
        Me.DeleteButton.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.DeleteButton.UseVisualStyleBackColor = False
        '
        'OkButton
        '
        Me.OkButton.BackColor = System.Drawing.SystemColors.Control
        Me.OkButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.OkButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.OkButton.Image = CType(resources.GetObject("OkButton.Image"), System.Drawing.Image)
        Me.OkButton.Location = New System.Drawing.Point(284, 343)
        Me.OkButton.Name = "OkButton"
        Me.OkButton.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.OkButton.Size = New System.Drawing.Size(44, 41)
        Me.OkButton.TabIndex = 72
        Me.OkButton.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.OkButton.UseVisualStyleBackColor = False
        '
        'ExitButton
        '
        Me.ExitButton.BackColor = System.Drawing.SystemColors.Control
        Me.ExitButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.ExitButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.ExitButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ExitButton.Image = CType(resources.GetObject("ExitButton.Image"), System.Drawing.Image)
        Me.ExitButton.Location = New System.Drawing.Point(497, 346)
        Me.ExitButton.Name = "ExitButton"
        Me.ExitButton.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ExitButton.Size = New System.Drawing.Size(44, 41)
        Me.ExitButton.TabIndex = 76
        Me.ExitButton.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ExitButton.UseVisualStyleBackColor = False
        '
        'NewButton
        '
        Me.NewButton.BackColor = System.Drawing.SystemColors.Control
        Me.NewButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.NewButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.NewButton.Image = CType(resources.GetObject("NewButton.Image"), System.Drawing.Image)
        Me.NewButton.Location = New System.Drawing.Point(329, 343)
        Me.NewButton.Name = "NewButton"
        Me.NewButton.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.NewButton.Size = New System.Drawing.Size(44, 41)
        Me.NewButton.TabIndex = 73
        Me.NewButton.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.NewButton.UseVisualStyleBackColor = False
        '
        'NextButton
        '
        Me.NextButton.BackColor = System.Drawing.SystemColors.Control
        Me.NextButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.NextButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.NextButton.Image = CType(resources.GetObject("NextButton.Image"), System.Drawing.Image)
        Me.NextButton.Location = New System.Drawing.Point(165, 343)
        Me.NextButton.Name = "NextButton"
        Me.NextButton.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.NextButton.Size = New System.Drawing.Size(44, 41)
        Me.NextButton.TabIndex = 70
        Me.NextButton.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.NextButton.UseVisualStyleBackColor = False
        '
        'FirstButton
        '
        Me.FirstButton.BackColor = System.Drawing.SystemColors.Control
        Me.FirstButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.FirstButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.FirstButton.Image = CType(resources.GetObject("FirstButton.Image"), System.Drawing.Image)
        Me.FirstButton.Location = New System.Drawing.Point(77, 343)
        Me.FirstButton.Name = "FirstButton"
        Me.FirstButton.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.FirstButton.Size = New System.Drawing.Size(44, 41)
        Me.FirstButton.TabIndex = 68
        Me.FirstButton.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.FirstButton.UseVisualStyleBackColor = False
        '
        'PrevButton
        '
        Me.PrevButton.BackColor = System.Drawing.SystemColors.Control
        Me.PrevButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.PrevButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.PrevButton.Image = CType(resources.GetObject("PrevButton.Image"), System.Drawing.Image)
        Me.PrevButton.Location = New System.Drawing.Point(121, 343)
        Me.PrevButton.Name = "PrevButton"
        Me.PrevButton.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.PrevButton.Size = New System.Drawing.Size(44, 41)
        Me.PrevButton.TabIndex = 69
        Me.PrevButton.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.PrevButton.UseVisualStyleBackColor = False
        '
        'LastButton
        '
        Me.LastButton.BackColor = System.Drawing.SystemColors.Control
        Me.LastButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.LastButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LastButton.Image = CType(resources.GetObject("LastButton.Image"), System.Drawing.Image)
        Me.LastButton.Location = New System.Drawing.Point(210, 343)
        Me.LastButton.Name = "LastButton"
        Me.LastButton.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.LastButton.Size = New System.Drawing.Size(44, 41)
        Me.LastButton.TabIndex = 71
        Me.LastButton.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.LastButton.UseVisualStyleBackColor = False
        '
        'OrderId_Label
        '
        Me.OrderId_Label.BackColor = System.Drawing.SystemColors.Control
        Me.OrderId_Label.Cursor = System.Windows.Forms.Cursors.Default
        Me.OrderId_Label.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.OrderId_Label.ForeColor = System.Drawing.Color.Blue
        Me.OrderId_Label.Location = New System.Drawing.Point(94, 10)
        Me.OrderId_Label.Name = "OrderId_Label"
        Me.OrderId_Label.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.OrderId_Label.Size = New System.Drawing.Size(112, 25)
        Me.OrderId_Label.TabIndex = 78
        Me.OrderId_Label.Text = "Order Id"
        Me.OrderId_Label.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'OrderID
        '
        Me.OrderID.AcceptsReturn = True
        Me.OrderID.BackColor = System.Drawing.SystemColors.Window
        Me.OrderID.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.OrderID.ForeColor = System.Drawing.SystemColors.WindowText
        Me.OrderID.Location = New System.Drawing.Point(206, 8)
        Me.OrderID.MaxLength = 5
        Me.OrderID.Name = "OrderID"
        Me.OrderID.Size = New System.Drawing.Size(84, 20)
        Me.OrderID.TabIndex = 63
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(77, 144)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(477, 178)
        Me.DataGridView1.TabIndex = 84
        '
        'TestForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(644, 419)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.Freight)
        Me.Controls.Add(Me.Freight_Label)
        Me.Controls.Add(Me.xCompanyName)
        Me.Controls.Add(Me.Orderdate_Label)
        Me.Controls.Add(Me.OrderDate)
        Me.Controls.Add(Me.ShipVia)
        Me.Controls.Add(Me.CustomerID)
        Me.Controls.Add(Me.SearchButton)
        Me.Controls.Add(Me.CustomerID_Label)
        Me.Controls.Add(Me.ShipVia_Label)
        Me.Controls.Add(Me.xCustomerName)
        Me.Controls.Add(Me.DeleteButton)
        Me.Controls.Add(Me.OkButton)
        Me.Controls.Add(Me.ExitButton)
        Me.Controls.Add(Me.NewButton)
        Me.Controls.Add(Me.NextButton)
        Me.Controls.Add(Me.FirstButton)
        Me.Controls.Add(Me.PrevButton)
        Me.Controls.Add(Me.LastButton)
        Me.Controls.Add(Me.OrderId_Label)
        Me.Controls.Add(Me.OrderID)
        Me.Name = "TestForm"
        Me.Text = "Form1"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
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
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
End Class
