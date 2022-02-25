Imports System.Data

Public Class TestForm

    Dim DM As New DynamicComponents.DataManager()
    Dim aImage(26) As String
    Dim CN As New OleDb.OleDbConnection
    Dim oMaster As New DataTable
    Dim oDetails As New DataTable
    Dim AccessConnect As String = ""



    Private Sub TestForm_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        'Establishing a connection
        AccessConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "\nWind.accdb; Persist Security Info=False;"



        PopulateaImage() ' define images used with buttons
        CN.ConnectionString = AccessConnect
        CN.Open()
        Dim ds As New DataSet
        Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter

        da.SelectCommand = New OleDb.OleDbCommand("Select * from orders", CN)
        da.Fill(ds, "Orders")
        oMaster = ds.Tables("Orders")

        da.SelectCommand = New OleDb.OleDbCommand("Select * from OrderDetails", CN)
        da.Fill(ds, "OrderDetails")
        oDetails = ds.Tables("OrderDetails")

        DM.InitForm(da, CN, ds, Me, oMaster, DataGridView1, oDetails) 'Must Be your first declaration
        DM.PrepareImageButtons(aImage, Application.StartupPath + "\icons\", False)
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
        DM.PopulateForm(Me, oMaster, DataGridView1, oDetails) 'Must be your last declaration
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
        CN.Close()
    End Sub

End Class
