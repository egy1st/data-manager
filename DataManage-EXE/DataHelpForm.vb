Imports System.Data
Friend Class DataHelpForm
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
    Friend WithEvents ID As System.Windows.Forms.TextBox
    Friend WithEvents Name0 As System.Windows.Forms.TextBox
    Friend WithEvents AxDataGrid1 As DataGridView
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(DataHelpForm))
        Me.ID = New System.Windows.Forms.TextBox()
        Me.Name0 = New System.Windows.Forms.TextBox()
        Me.AxDataGrid1 = New DataGridView
        CType(Me.AxDataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ID
        '
        Me.ID.Name = "ID"
        Me.ID.Size = New System.Drawing.Size(104, 20)
        Me.ID.TabIndex = 1
        Me.ID.Text = ""
        '
        'Name0
        '
        Me.Name0.Location = New System.Drawing.Point(103, 0)
        Me.Name0.Name = "Name0"
        Me.Name0.Size = New System.Drawing.Size(105, 20)
        Me.Name0.TabIndex = 2
        Me.Name0.Text = ""
        '
        'AxDataGrid1
        '
        Me.AxDataGrid1.DataSource = Nothing
        Me.AxDataGrid1.Location = New System.Drawing.Point(0, 24)
        Me.AxDataGrid1.Name = "AxDataGrid1"
        'Me.AxDataGrid1.OcxState = CType(resources.GetObject("AxDataGrid1.OcxState"), System.Windows.Forms.AxHost.State)  // maa
        Me.AxDataGrid1.Size = New System.Drawing.Size(224, 224)
        Me.AxDataGrid1.TabIndex = 3
        '
        'DataHelpForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(224, 254)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.AxDataGrid1, Me.Name0, Me.ID})
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "DataHelpForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        CType(Me.AxDataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim oRecSet As New DataTable
    Dim ds As New DataSet
    Dim da_rec As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter
    Private Sub HelpForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim cSelect As String
        Me.Text = aInitValues(15)
        Me.AxDataGrid1.RightToLeft = Right2LeftState

        cSelect = "Select " + HelpID + " , " + HelpName + " From " + HelpFile
        da_rec.SelectCommand = New OleDb.OleDbCommand(cSelect, CN)
        da_rec.Fill(ds, HelpFile)
        oRecSet = ds.Tables(HelpFile)
        'oRecSet.Open(cSelect, CN)
        Me.AxDataGrid1.DataSource = oRecSet
        'Me.ID.MaxLength = oRecSet(HelpID).DefinedSize // maavip
        'Me.Name0.MaxLength = oRecSet(HelpName).DefinedSize // maavip
        Me.ID.Width = Me.AxDataGrid1.Columns(0).Width
        'Me.Name0.Left = Me.AxDataGrid1.Columns(1).Left // maa
        Me.Name0.Width = Me.AxDataGrid1.Columns(1).Width
        Me.AxDataGrid1.Width = Me.AxDataGrid1.Columns(0).Width + Me.AxDataGrid1.Columns(1).Width + 40
        Me.Width = Me.AxDataGrid1.Width + 10
        Me.AxDataGrid1.Columns(0).HeaderText = aInitValues(16)
        Me.AxDataGrid1.Columns(1).HeaderText = aInitValues(17)
        If Right2LeftState = 1 Then
            Me.ID.Left = Me.Width - Me.ID.Width - 10
            Me.Name0.Left = Me.ID.Left - Me.Name0.Width
        End If
        Me.RightToLeft = Right2LeftState 'keep it at end 
    End Sub

    Private Sub ID_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ID.TextChanged
        If Me.ID.Text <> "" Then
            oRecSet.Select(HelpID + " Like '" + Me.ID.Text + "*'")
        Else
            oRecSet.Select(HelpID + " <> 'MAA_13_12_19_71_MAA'")
        End If
    End Sub

    Private Sub Name0_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Name0.TextChanged
        If Me.Name0.Text <> "" Then
            oRecSet.Select(HelpName + " Like '" + Me.Name0.Text + "*'")
        Else
            oRecSet.Select(HelpName + " <> 'MAA_13_12_19_71_MAA'")
        End If
    End Sub

    Private Sub AxDataGrid1_DblClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles AxDataGrid1.DoubleClick
        HelpRtnID = ""
        HelpRtnName = ""
        'HelpRtnID = AxDataGrid1.Columns(0).Value   // maa
        'HelpRtnName = AxDataGrid1.Columns(1).Value //maa
        Me.Close()
    End Sub

End Class
