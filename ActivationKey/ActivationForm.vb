Public Class ActivationForm
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ProductKey As System.Windows.Forms.TextBox
    Friend WithEvents ActivationKey As System.Windows.Forms.TextBox
    Friend WithEvents Generator As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.ProductKey = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ActivationKey = New System.Windows.Forms.TextBox()
        Me.Generator = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'ProductKey
        '
        Me.ProductKey.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.ProductKey.Location = New System.Drawing.Point(96, 24)
        Me.ProductKey.MaxLength = 4
        Me.ProductKey.Name = "ProductKey"
        Me.ProductKey.Size = New System.Drawing.Size(72, 22)
        Me.ProductKey.TabIndex = 0
        Me.ProductKey.Text = ""
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(88, 24)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Product Key"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 72)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(88, 24)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Activation Key"
        '
        'ActivationKey
        '
        Me.ActivationKey.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.ActivationKey.Location = New System.Drawing.Point(96, 64)
        Me.ActivationKey.MaxLength = 21
        Me.ActivationKey.Name = "ActivationKey"
        Me.ActivationKey.Size = New System.Drawing.Size(208, 22)
        Me.ActivationKey.TabIndex = 2
        Me.ActivationKey.Text = ""
        '
        'Generator
        '
        Me.Generator.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Generator.Location = New System.Drawing.Point(96, 104)
        Me.Generator.Name = "Generator"
        Me.Generator.Size = New System.Drawing.Size(136, 24)
        Me.Generator.TabIndex = 4
        Me.Generator.Text = "Generate New Key"
        '
        'ActivationForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(312, 150)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Generator, Me.Label2, Me.ActivationKey, Me.Label1, Me.ProductKey})
        Me.Name = "ActivationForm"
        Me.Text = "Activation Key"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Function ZeroPad(ByVal str_String As String, ByVal int_Count As Byte) As String
        If str_String <> "" Then
            Return (New String("0", int_Count - Len(Trim(str_String))) & Trim(str_String))
        End If
    End Function

    Private Function CheckSum(ByVal strNum As String) As Byte
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
        Return intCheckSum
    End Function

    Private Sub Generator_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Generator.Click
        Dim KeyRnd As Integer
        Dim MyKey As String
        Dim myChecksumKey As String
        Dim Rest As String
        Dim ProductID As String
        Dim Part1, Part2, Part3, Part4, Part5 As String

        Me.ActivationKey.Text = ""
        If Me.ProductKey.Text = "" Then
            MsgBox("Please enter Product key")
            Exit Sub
        End If
Start:
        Randomize(CInt(Mid((Now.ToOADate * 1000000).ToString, 5, 6)))
        KeyRnd = Rnd() * 2000 + 2000
        MyKey = 13 * KeyRnd ^ 3 + 12 * KeyRnd ^ 2 + 19 * KeyRnd ^ 1 + 71 * KeyRnd ^ 0

        ProductID = ZeroPad(Me.ProductKey.Text * Me.ProductKey.Text, 8)
        Part1 = KeyRnd + CInt((Mid(ProductID, 1, 4)))
        Part1 = ZeroPad(Part1, 4) ' no must  , it must be 4 digit
        If Part1 > 9999 Then GoTo Start

        Part2 = CInt(Mid(MyKey, 1, 3)) + CInt(Mid(ProductID, 5, 1))
        Part2 = ZeroPad(Part2, 3)
        If Part2 > 999 Then GoTo Start
        Part3 = CInt(Mid(MyKey, 4, 3)) + CInt(Mid(ProductID, 6, 1))
        Part3 = ZeroPad(Part3, 3)
        If Part3 > 999 Then GoTo Start
        Part4 = CInt(Mid(MyKey, 7, 3)) + CInt(Mid(ProductID, 7, 1))
        Part4 = ZeroPad(Part4, 3)
        If Part4 > 999 Then GoTo Start
        Part5 = CInt(Mid(MyKey, 10, 3)) + CInt(Mid(ProductID, 8, 1))
        Part5 = ZeroPad(Part5, 3)
        If Part5 > 9999 Then GoTo Start

        myChecksumKey = Part1 + Part2 + Part3 + Part4 + Part5 + "0"
        Rest = CheckSum(myChecksumKey)
        myChecksumKey = Mid(myChecksumKey, 1, 16) + (CInt(Mid(myChecksumKey, 17, 1)) + 10 - Rest).ToString

        Me.ActivationKey.Text = Mid(myChecksumKey, 1, 4) + "-" + Mid(myChecksumKey, 5, 3) + "-" + Mid(myChecksumKey, 8, 3) + "-" + Mid(myChecksumKey, 11, 3) + "-" + Mid(myChecksumKey, 14, 4)

    End Sub

    Private Sub ProductKey_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProductKey.Leave
        If sender.Text.Length <> 4 Or Not IsNumeric(sender.Text) Then
            MsgBox("Product length must be 4 Digits")
            Exit Sub
        End If

    End Sub

    Private Sub ActivationForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
