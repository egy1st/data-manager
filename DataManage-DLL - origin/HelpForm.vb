Friend Class HelpForm
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
    Friend WithEvents TreeView1 As System.Windows.Forms.TreeView
    Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents Label1 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(HelpForm))
        Me.TreeView1 = New System.Windows.Forms.TreeView()
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox()
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'TreeView1
        '
        Me.TreeView1.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TreeView1.ImageIndex = -1
        Me.TreeView1.Name = "TreeView1"
        Me.TreeView1.SelectedImageIndex = -1
        Me.TreeView1.Size = New System.Drawing.Size(264, 320)
        Me.TreeView1.TabIndex = 0
        '
        'RichTextBox1
        '
        Me.RichTextBox1.Location = New System.Drawing.Point(288, 0)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.ReadOnly = True
        Me.RichTextBox1.Size = New System.Drawing.Size(368, 320)
        Me.RichTextBox1.TabIndex = 1
        Me.RichTextBox1.Text = ""
        '
        'ImageList1
        '
        Me.ImageList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.ImageList1.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        '
        'Label1
        '
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label1.Location = New System.Drawing.Point(264, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(24, 320)
        Me.Label1.TabIndex = 2
        '
        'HelpForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(656, 326)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label1, Me.RichTextBox1, Me.TreeView1})
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "HelpForm"
        Me.Text = "Help"
        Me.ResumeLayout(False)

    End Sub

#End Region
    Dim oHelp As New ADODB.Recordset()
    Dim HelpControlName As String
    Dim HelpControlType As String
    Dim HelpTag As String

    Private Sub TreeView1_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeView1.AfterSelect
        On Error Resume Next
        Dim cFilter As String

        Static Flag As Byte = 0
        Flag += 1
        If Flag > 1 Then
            oHelp.Filter = "id <> '13_12_19_71_MAA_Mohamed_Aly_Abbas_MAA'"
            cFilter = " Tag = '" + e.Node.Parent.Text + "' and Description = '" + e.Node.Text + "'"
            oHelp.Filter = cFilter
            RichTextBox1.Text = oHelp.Fields("Description").Value + Chr(13) + oHelp.Fields("Contents").Value
            RichTextBox1.Select(0, Len(oHelp.Fields("Description").Value))
            RichTextBox1.SelectionFont = New Font("Tahoma", 18, FontStyle.Underline)
            RichTextBox1.SelectionColor = System.Drawing.Color.Red
            RichTextBox1.Select(Len(oHelp.Fields("Description").Value) + 1, 10000)
            RichTextBox1.SelectionFont = New Font("Tahoma", 14, FontStyle.Bold, FontStyle.Italic)
            RichTextBox1.SelectionColor = System.Drawing.Color.Blue
        End If
    End Sub

    Private Sub Form4_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim NewVal As String = ""
        Dim Oldval As String = ""
        Dim RootNode As TreeNode
        Dim SubNode As TreeNode
        Dim ChildNode As TreeNode
        Dim Cfilter As String
        Dim cProperties As String
        Dim Num As Integer = 0
        Dim RecCount As Integer
        On Error GoTo EndMe

        Me.Text = aInitValues(0)
        Me.RightToLeft = Right2LeftState
        If Right2LeftState = 1 Then
            Me.RichTextBox1.Left = Me.Width - (Me.RichTextBox1.Left + Me.RichTextBox1.Width)
            Me.TreeView1.Left = Me.Width - (Me.TreeView1.Left + Me.TreeView1.Width)
            Me.Label1.Left = Me.Width - (Me.Label1.Left + Me.Label1.Width)
        End If
        HelpTag = HelpFieldProperty(0)
        HelpControlName = HelpFieldProperty(1)
        HelpControlType = HelpFieldProperty(6)
        oHelp.Open("Select * From Help order by Tag", CN, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)

        RecCount = oHelp.RecordCount
        TreeView1.ImageList = ImageList1
        RootNode = New TreeNode(aInitValues(0))

        TreeView1.Nodes.Add(RootNode)
        RootNode.ImageIndex = 0
        Do While Not oHelp.EOF
            NewVal = oHelp.Fields("Tag").Value
            If NewVal <> Oldval Then
                SubNode = New TreeNode(oHelp.Fields("Tag").Value)
                RootNode.Nodes.Add(SubNode)
                SubNode.ImageIndex = 1
            End If
            Oldval = NewVal
            ChildNode = New TreeNode(oHelp.Fields("Description").Value)
            SubNode.Nodes.Add(ChildNode)
            ChildNode.ImageIndex = 2
            oHelp.MoveNext()
        Loop

        TreeView1.Nodes(0).Expand()
        oHelp.Filter = "id <> '13_12_19_71_MAA_Mohamed_Aly_Abbas_MAA'"
        Cfilter = " Tag = '" + HelpTag + "' and id = '" + HelpControlName + "'"
        oHelp.Filter = Cfilter
        If Not oHelp.EOF Then
            If HelpControlType = "TextBox" Then
                cProperties = aInitValues(1) + IIf(HelpFieldProperty(2) <> "", HelpFieldProperty(2), aInitValues(5)) + Chr(13)
                cProperties += aInitValues(2) + IIf(HelpFieldProperty(3) <> "", HelpFieldProperty(3).ToString + " " + aInitValues(6), aInitValues(5)) + Chr(13)
                cProperties += aInitValues(3) + IIf(HelpFieldProperty(4) <> "", aInitValues(8), aInitValues(9)) + Chr(13)
                cProperties += aInitValues(4) + IIf(HelpFieldProperty(5) <> "", aInitValues(8), aInitValues(9)) + Chr(13) + Chr(13)
                cProperties += aInitValues(7) + Chr(13)
            End If
            RichTextBox1.Text = oHelp.Fields("Description").Value + Chr(13) + Chr(13) + cProperties + oHelp.Fields("Contents").Value
            RichTextBox1.Select(0, Len(oHelp.Fields("Description").Value))
            RichTextBox1.SelectionFont = New Font("Tahoma", 18, FontStyle.Underline)
            RichTextBox1.SelectionColor = System.Drawing.Color.Red
            RichTextBox1.Select(Len(oHelp.Fields("Description").Value) + 1, 65000)
            RichTextBox1.SelectionFont = New Font("Tahoma", 14, FontStyle.Bold, FontStyle.Italic)
            RichTextBox1.SelectionColor = System.Drawing.Color.Blue

            For Num = 0 To RecCount
                If TreeView1.Nodes(0).Nodes(Num).Text = HelpTag Then
                    TreeView1.Nodes(0).Nodes(Num).Expand()
                    Exit For
                End If
            Next Num
        End If
EndMe:
    End Sub


End Class
