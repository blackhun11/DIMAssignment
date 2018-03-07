Imports System.Data.SqlClient
Imports System.Data.OleDb


Public Class DataAccess
    Private myConn As SqlConnection = New SqlConnection("server = '.' ;database = 'DianaFreshFood';Trusted_Connection = yes")
    Private myCmd As SqlCommand
    Private myReader As SqlDataReader
    Private results As String
    Private tempNama As String
    Private adapter As SqlDataAdapter

    Dim cmd As New SqlCommand
    Dim rdr As SqlDataReader
    Dim dt As New DataTable
    Dim ds As New DataSet

    Public Sub closeConnection()
        myConn.Close()
    End Sub

    Public Sub testConnection()
        myConn.Open()
        Dim CMD As New SqlCommand
        CMD = myConn.CreateCommand()
        CMD.CommandText = "INSERT INTO MSSTAFF VALUES ('S000000001','ASD','ASD','ASD','ASD','ASD','ASD','ASD')"
        CMD.ExecuteNonQuery()
        myConn.Close()
    End Sub

    Public Function validateLogin(ByVal StaffEmail As String, ByVal StaffPassword As String)

        myConn.Open()
        cmd = myConn.CreateCommand()
        cmd.CommandText = "validateLogin"
        cmd.CommandType = CommandType.StoredProcedure
        'cmd.CommandText = "select * from msStaff where StaffID = @StaffID and StaffPassword = @StaffPassword"
        cmd.Parameters.AddWithValue("@StaffEmail", StaffEmail)
        cmd.Parameters.AddWithValue("@StaffPassword", StaffPassword)
        rdr = cmd.ExecuteReader()
        rdr.Read()

        If (rdr.HasRows) Then
            login.Hide()
            mainForm.Show()
            Return "welcome, " + String.Format("{0}", rdr(0))
            myConn.Close()
        Else
            myConn.Close()
            Return "gagal"

        End If

    End Function

    '************************************************
    'viewStaff
    '************************************************

    Public Sub selectAllStaff(ByRef tempData As DataGridView)
        myConn.Open()

        cmd = myConn.CreateCommand()
        cmd.CommandText = "selectAllStaff"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.ExecuteNonQuery()
        adapter = New SqlDataAdapter(cmd)
        Dim datatable As New DataTable
        adapter.Fill(datatable)
        tempData.DataSource = datatable
        tempData.Columns(0).Visible = False
        tempData.Columns(8).Visible = False
        tempData.Columns(9).Visible = False
        tempData.Columns(10).Visible = False
        tempData.Columns(11).Visible = False
        tempData.Columns(12).Visible = False
        myConn.Close()
    End Sub

    Public Sub insertStaff(ByVal StaffName As String, ByVal StaffPassword As String,
                                ByVal StaffAddress As String, ByVal StaffEmail As String, ByVal StaffAltEmail As String,
                                ByVal StaffPhone As String, ByVal StaffAltPhone As String)
        myConn.Open()
        cmd = myConn.CreateCommand()
        cmd.CommandText = "insertStaff"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@StaffName", StaffName)
        cmd.Parameters.AddWithValue("@StaffPassword", StaffPassword)
        cmd.Parameters.AddWithValue("@StaffAddress", StaffAddress)
        cmd.Parameters.AddWithValue("@StaffEmail", StaffEmail)
        cmd.Parameters.AddWithValue("@StaffAltEmail", StaffAltEmail)
        cmd.Parameters.AddWithValue("@StaffPhone", StaffPhone)
        cmd.Parameters.AddWithValue("@StaffAltPhone", StaffAltPhone)
        cmd.ExecuteNonQuery()
        myConn.Close()
    End Sub

    Public Sub updateStaff(ByVal StaffID As String, ByVal StaffName As String, ByVal StaffPassword As String,
                            ByVal StaffAddress As String, ByVal StaffEmail As String, ByVal StaffAltEmail As String,
                            ByVal StaffPhone As String, ByVal StaffAltPhone As String)
        myConn.Open()
        cmd = myConn.CreateCommand()
        cmd.CommandText = "updateStaff"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@StaffID", StaffID)
        cmd.Parameters.AddWithValue("@StaffName", StaffName)
        cmd.Parameters.AddWithValue("@StaffPassword", StaffPassword)
        cmd.Parameters.AddWithValue("@StaffAddress", StaffAddress)
        cmd.Parameters.AddWithValue("@StaffEmail", StaffEmail)
        cmd.Parameters.AddWithValue("@StaffAltEmail", StaffAltEmail)
        cmd.Parameters.AddWithValue("@StaffPhone", StaffPhone)
        cmd.Parameters.AddWithValue("@StaffAltPhone", StaffAltPhone)
        cmd.ExecuteNonQuery()
        myConn.Close()
    End Sub

    Public Sub deleteStaff(ByVal StaffID As String)
        myConn.Open()
        cmd = myConn.CreateCommand()
        cmd.CommandText = "deleteStaff"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@StaffID", StaffID)
        cmd.ExecuteNonQuery()
        myConn.Close()

    End Sub

    '************************************************
    'viewMarket
    '************************************************
    Public Sub selectAllMarket(ByRef tempData As DataGridView)
        myConn.Open()

        cmd = myConn.CreateCommand()
        cmd.CommandText = "selectAllMarket"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.ExecuteNonQuery()
        adapter = New SqlDataAdapter(cmd)
        Dim datatable As New DataTable
        'dt.Clear()
        adapter.Fill(datatable)
        tempData.DataSource = datatable
        tempData.Columns(0).Visible = False
        tempData.Columns(3).Visible = False
        tempData.Columns(4).Visible = False
        myConn.Close()
    End Sub

    Public Sub insertMarket(ByVal MarketName As String, ByVal MarketAddress As String)
        myConn.Open()
        cmd = myConn.CreateCommand()
        cmd.CommandText = "insertMarket"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@MarketName", MarketName)
        cmd.Parameters.AddWithValue("@MarketAddress", MarketAddress)
        cmd.ExecuteNonQuery()
        myConn.Close()
    End Sub

    Public Sub updateMarket(ByVal MarketID As String, ByVal MarketName As String, ByVal MarketAddress As String)
        myConn.Open()
        cmd = myConn.CreateCommand()
        cmd.CommandText = "updateMarket"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@MarketID", MarketID)
        cmd.Parameters.AddWithValue("@MarketName", MarketName)
        cmd.Parameters.AddWithValue("@MarketAddress", MarketAddress)
        cmd.ExecuteNonQuery()
        myConn.Close()
    End Sub

    Public Sub deleteMarket(ByVal MarketID As String)
        myConn.Open()
        cmd = myConn.CreateCommand()
        cmd.CommandText = "deleteMarket"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@MarketID", MarketID)
        cmd.ExecuteNonQuery()
        myConn.Close()
    End Sub


    '************************************************
    'viewProduct
    '************************************************
    Public Sub selectAllProduct(ByRef tempData As DataGridView)
        myConn.Open()

        cmd = myConn.CreateCommand()
        cmd.CommandText = "selectAllProduct"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.ExecuteNonQuery()
        adapter = New SqlDataAdapter(cmd)
        Dim datatable As New DataTable
        'dt.Clear()
        adapter.Fill(datatable)
        tempData.DataSource = datatable
        tempData.Columns(0).Visible = False

        myConn.Close()
    End Sub

    Public Sub insertProduct(ByVal ProductName As String, ByVal ProductPriceforSale As String, ByVal ProductPriceforPurchase As String)
        myConn.Open()
        cmd = myConn.CreateCommand()
        cmd.CommandText = "insertProduct"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@ProductName", ProductName)
        cmd.Parameters.AddWithValue("@ProductPriceforSale", ProductPriceforSale)
        cmd.Parameters.AddWithValue("@ProductPriceforPurchase", ProductPriceforPurchase)
        cmd.ExecuteNonQuery()
        myConn.Close()
    End Sub

    Public Sub updateProduct(ByVal ProductID As String, ByVal ProductName As String, ByVal ProductPriceforSale As Integer,
                             ByVal ProductPriceforPurchase As Integer)
        myConn.Open()
        cmd = myConn.CreateCommand()
        cmd.CommandText = "updateProduct"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@ProductID", ProductID)
        cmd.Parameters.AddWithValue("@ProductName", ProductName)
        cmd.Parameters.AddWithValue("@ProductPriceforSale", ProductPriceforSale)
        cmd.Parameters.AddWithValue("@ProductPriceforPurchase", ProductPriceforPurchase)
        cmd.ExecuteNonQuery()
        myConn.Close()
    End Sub

    Public Sub deleteProduct(ByVal ProductID As String)
        myConn.Open()
        cmd = myConn.CreateCommand()
        cmd.CommandText = "deleteProduct"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@ProductID", ProductID)
        cmd.ExecuteNonQuery()
        myConn.Close()


    End Sub

    '************************************************
    'viewSupplier
    '************************************************

    Public Sub selectAllSupplier(ByRef tempData As DataGridView)
        myConn.Open()

        cmd = myConn.CreateCommand()
        cmd.CommandText = "selectAllSupplier"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.ExecuteNonQuery()
        adapter = New SqlDataAdapter(cmd)
        Dim datatable As New DataTable
        'dt.Clear()
        adapter.Fill(datatable)
        tempData.DataSource = datatable
        tempData.Columns(0).Visible = False
        tempData.Columns(7).Visible = False
        tempData.Columns(8).Visible = False
        tempData.Columns(9).Visible = False
        myConn.Close()
    End Sub

    Public Sub insertSupplier(ByVal SupplierName As String, ByVal SupplierAddress As String, ByVal SupplierEmail As String, ByVal SupplierAltEmail As String,
                              ByVal SupplierPhone As String, ByVal SupplierAltPhone As String, ByVal ProductName As String)
        myConn.Open()
        cmd = myConn.CreateCommand()
        cmd.CommandText = "insertSupplier"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@SupplierName", SupplierName)
        cmd.Parameters.AddWithValue("@SupplierAddress", SupplierAddress)
        cmd.Parameters.AddWithValue("@SupplierEmail", SupplierEmail)
        cmd.Parameters.AddWithValue("@SupplierAltEmail", SupplierAltEmail)
        cmd.Parameters.AddWithValue("@SupplierPhone", SupplierPhone)
        cmd.Parameters.AddWithValue("@SupplierAltPhone", SupplierAltPhone)
        cmd.Parameters.AddWithValue("@ProductName", ProductName)
        cmd.ExecuteNonQuery()
        myConn.Close()
    End Sub

    Public Sub updateSupplier(ByVal SupplierID As String, ByVal SupplierName As String, ByVal SupplierAddress As String, ByVal SupplierEmail As String, ByVal SupplierAltEmail As String,
                              ByVal SupplierPhone As String, ByVal SupplierAltPhone As String, ByVal ProductName As String)
        myConn.Open()
        cmd = myConn.CreateCommand()
        cmd.CommandText = "updateSupplier"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("SupplierID", SupplierID)
        cmd.Parameters.AddWithValue("@SupplierName", SupplierName)
        cmd.Parameters.AddWithValue("@SupplierAddress", SupplierAddress)
        cmd.Parameters.AddWithValue("@SupplierEmail", SupplierEmail)
        cmd.Parameters.AddWithValue("@SupplierAltEmail", SupplierAltEmail)
        cmd.Parameters.AddWithValue("@SupplierPhone", SupplierPhone)
        cmd.Parameters.AddWithValue("@SupplierAltPhone", SupplierAltPhone)
        cmd.Parameters.AddWithValue("@ProductName", ProductName)
        cmd.ExecuteNonQuery()
        myConn.Close()
    End Sub

    Public Sub deleteSupplier(ByVal SupplierID As String)
        myConn.Open()
        cmd = myConn.CreateCommand()
        cmd.CommandText = "deleteSupplier"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@SupplierID", SupplierID)
        cmd.ExecuteNonQuery()
        myConn.Close()


    End Sub
    '************************************************
    'Combo Box
    '************************************************
    Public Sub getProductforSupplier(ByRef tempData As ComboBox)
        myConn.Open()
        cmd = myConn.CreateCommand()
        cmd.CommandText = "ddlProductName"
        cmd.CommandType = CommandType.StoredProcedure
        rdr = cmd.ExecuteReader()
        tempData.Items.Clear()
        While rdr.Read()
            tempData.Items.Add(rdr.GetString(0))
        End While

        myConn.Close()
    End Sub

    Public Sub getMarketforSales(ByRef tempData As ComboBox)
        myConn.Open()
        cmd = myConn.CreateCommand()
        cmd.CommandText = "ddlMarketName"
        cmd.CommandType = CommandType.StoredProcedure
        rdr = cmd.ExecuteReader()
        tempData.Items.Clear()
        While rdr.Read()
            tempData.Items.Add(rdr.GetString(0))
        End While

        myConn.Close()
    End Sub

    Public Sub getStaffNameforSales(ByRef tempData As ComboBox)
        myConn.Open()
        cmd = myConn.CreateCommand()
        cmd.CommandText = "selectAllStaff"
        cmd.CommandType = CommandType.StoredProcedure
        rdr = cmd.ExecuteReader()
        tempData.Items.Clear()
        While rdr.Read()
            tempData.Items.Add(rdr.GetString(1))
        End While

        myConn.Close()
    End Sub

    '************************************************
    'Beban Awal
    '************************************************
    Public Sub selectnotFinalizeExpense(ByRef tempData As DataGridView)
        myConn.Open()
        cmd = myConn.CreateCommand()
        cmd.CommandText = "selectnotFinalizeExpense"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.ExecuteNonQuery()
        adapter = New SqlDataAdapter(cmd)
        Dim datatable As New DataTable
        'dt.Clear()
        adapter.Fill(datatable)
        tempData.DataSource = datatable
        tempData.Columns(0).Visible = False

        myConn.Close()
    End Sub

    Public Sub insertNewExpenses(ByVal ExpenseName As String, ByVal Cost As Integer, ByVal qty As Integer)
        myConn.Open()
        cmd = myConn.CreateCommand()
        cmd.CommandText = "insertNewExpenses"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@ExpenseName", ExpenseName)
        cmd.Parameters.AddWithValue("@Cost", Cost)
        cmd.Parameters.AddWithValue("@Qty", qty)
        cmd.ExecuteNonQuery()
        myConn.Close()
    End Sub

    '************************************************
    'Pembelian Awal
    '************************************************

    Public Sub selectnotFinalizePurchase(ByRef tempData As DataGridView)
        myConn.Open()
        cmd = myConn.CreateCommand()
        cmd.CommandText = "selectnotFinalizePurchase"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.ExecuteNonQuery()
        adapter = New SqlDataAdapter(cmd)
        Dim datatable As New DataTable
        ' dt.Clear()
        adapter.Fill(datatable)
        tempData.DataSource = datatable


        myConn.Close()
    End Sub



    Public Sub insertPurchase(ByVal StaffName As String,
                                 ByVal ProductName1 As String, ByVal ProductName2 As String,
                                 ByVal ProductName3 As String, ByVal ProductName4 As String,
                                 ByVal ProductName5 As String, ByVal Qty1 As Integer,
                                 ByVal Qty2 As Integer, ByVal Qty3 As Integer,
                                 ByVal Qty4 As Integer, ByVal Qty5 As Integer)
        myConn.Open()
        cmd = myConn.CreateCommand()
        cmd.CommandText = "insertPurchase"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@StaffName", StaffName)
        cmd.Parameters.AddWithValue("@ProductName1", ProductName1)
        cmd.Parameters.AddWithValue("@ProductName2", ProductName2)
        cmd.Parameters.AddWithValue("@ProductName3", ProductName3)
        cmd.Parameters.AddWithValue("@ProductName4", ProductName4)
        cmd.Parameters.AddWithValue("@ProductName5", ProductName5)
        cmd.Parameters.AddWithValue("@Qty1", Qty1)
        cmd.Parameters.AddWithValue("@Qty2", Qty2)
        cmd.Parameters.AddWithValue("@Qty3", Qty3)
        cmd.Parameters.AddWithValue("@Qty4", Qty4)
        cmd.Parameters.AddWithValue("@Qty5", Qty5)

        cmd.ExecuteNonQuery()
        myConn.Close()
    End Sub

    '************************************************
    'Penjualan Awal
    '************************************************

    Public Sub selectnotFinalizeSales(ByRef tempData As DataGridView)
        myConn.Open()
        cmd = myConn.CreateCommand()
        cmd.CommandText = "selectnotFinalizeSales"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.ExecuteNonQuery()
        adapter = New SqlDataAdapter(cmd)
        Dim datatable As New DataTable

        adapter.Fill(datatable)
        tempData.DataSource = datatable


        myConn.Close()
    End Sub

    Public Sub insertSales(ByVal StaffName As String, ByVal MarketName As String,
                                ByVal ProductName1 As String, ByVal ProductName2 As String,
                                ByVal ProductName3 As String, ByVal ProductName4 As String,
                                ByVal ProductName5 As String, ByVal Qty1 As Integer,
                                ByVal Qty2 As Integer, ByVal Qty3 As Integer,
                                ByVal Qty4 As Integer, ByVal Qty5 As Integer)
        myConn.Open()
        cmd = myConn.CreateCommand()
        cmd.CommandText = "insertSales"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@StaffName", StaffName)
        cmd.Parameters.AddWithValue("@MarketName", MarketName)
        cmd.Parameters.AddWithValue("@ProductName1", ProductName1)
        cmd.Parameters.AddWithValue("@ProductName2", ProductName2)
        cmd.Parameters.AddWithValue("@ProductName3", ProductName3)
        cmd.Parameters.AddWithValue("@ProductName4", ProductName4)
        cmd.Parameters.AddWithValue("@ProductName5", ProductName5)
        cmd.Parameters.AddWithValue("@Qty1", Qty1)
        cmd.Parameters.AddWithValue("@Qty2", Qty2)
        cmd.Parameters.AddWithValue("@Qty3", Qty3)
        cmd.Parameters.AddWithValue("@Qty4", Qty4)
        cmd.Parameters.AddWithValue("@Qty5", Qty5)

        cmd.ExecuteNonQuery()
        myConn.Close()
    End Sub

    Public Sub isFinalized()
        myConn.Open()
        cmd = myConn.CreateCommand()
        cmd.CommandText = "isFinalized"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.ExecuteNonQuery()

        myConn.Close()
    End Sub

    Public Sub monthlyReport(ByVal Month As Integer, ByVal Year As Integer, ByVal tempData As DataGridView)
        myConn.Open()

        cmd = myConn.CreateCommand()
        cmd.CommandText = "monthlyReport"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Month", Month)
        cmd.Parameters.AddWithValue("@Year", Year)
        cmd.ExecuteNonQuery()
        adapter = New SqlDataAdapter(cmd)
        Dim datatable As New DataTable

        adapter.Fill(datatable)
        tempData.DataSource = datatable

        myConn.Close()

    End Sub

    Public Sub viewBiaya(ByVal Month As Integer, ByVal Year As Integer, ByVal tempData As DataGridView)
        myConn.Open()

        cmd = myConn.CreateCommand()
        cmd.CommandText = "detailExpense"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Month", Month)
        cmd.Parameters.AddWithValue("@Year", Year)
        cmd.ExecuteNonQuery()
        adapter = New SqlDataAdapter(cmd)
        Dim datatable As New DataTable

        adapter.Fill(datatable)
        tempData.DataSource = datatable

        myConn.Close()

    End Sub

    'Public Sub getTotalTransaction(ByVal Month As Integer, ByVal Year As Integer
    '                              )
    '    myConn.Open()

    '    cmd = myConn.CreateCommand()
    '    cmd.CommandText = "getTotalTransaction"
    '    cmd.CommandType = CommandType.StoredProcedure
    '    cmd.Parameters.AddWithValue("@Month", Month)
    '    cmd.Parameters.AddWithValue("@Year", Year)
    '    rdr = cmd.ExecuteReader()

    '    With rdr
    '        .Read()
    '        mainForm.txtPenjualanperMonth.Text = rdr.GetValue(1)
    '        mainForm.txtPembelianperMonth.Text = rdr.GetValue(2)
    '        mainForm.txtTotalCost.Text = rdr.GetValue(3)
    '        myConn.Close()
    '    End With
    'End Sub

    'Public Sub stockPermonth(ByVal Month As Integer, ByVal Year As Integer
    '                              )
    '    myConn.Open()

    '    cmd = myConn.CreateCommand()
    '    cmd.CommandText = "stockPerMonth"
    '    cmd.CommandType = CommandType.StoredProcedure
    '    cmd.Parameters.AddWithValue("@Month", Month)
    '    cmd.Parameters.AddWithValue("@Year", Year)
    '    rdr = cmd.ExecuteReader()

    '    With rdr
    '        .Read()
    '        mainForm.txtStock.Text = rdr.GetValue(1)
    '        myConn.Close()
    '    End With
    'End Sub

    'Public Sub netPurchase(ByVal Month As Integer, ByVal Year As Integer
    '                              )
    '    myConn.Open()

    '    cmd = myConn.CreateCommand()
    '    cmd.CommandText = "netPurchase"
    '    cmd.CommandType = CommandType.StoredProcedure
    '    cmd.Parameters.AddWithValue("@Month", Month)
    '    cmd.Parameters.AddWithValue("@Year", Year)
    '    rdr = cmd.ExecuteReader()

    '    With rdr
    '        .Read()
    '        mainForm.txtHasil.Text = rdr.GetValue(1)
    '        myConn.Close()
    '    End With
    'End Sub

    'Public Sub grossProfit(ByVal Month As Integer, ByVal Year As Integer
    '                             )
    '    myConn.Open()

    '    cmd = myConn.CreateCommand()
    '    cmd.CommandText = "grossProfit"
    '    cmd.CommandType = CommandType.StoredProcedure
    '    cmd.Parameters.AddWithValue("@Month", Month)
    '    cmd.Parameters.AddWithValue("@Year", Year)
    '    rdr = cmd.ExecuteReader()

    '    With rdr
    '        .Read()
    '        mainForm.txtTotalLabaKotor.Text = rdr.GetValue(1)
    '        myConn.Close()
    '    End With
    'End Sub

    'Public Sub incomeStatement(ByVal Month As Integer, ByVal Year As Integer
    '                         )
    '    myConn.Open()

    '    cmd = myConn.CreateCommand()
    '    cmd.CommandText = "incomeStatement"
    '    cmd.CommandType = CommandType.StoredProcedure
    '    cmd.Parameters.AddWithValue("@Month", Month)
    '    cmd.Parameters.AddWithValue("@Year", Year)
    '    rdr = cmd.ExecuteReader()

    '    With rdr
    '        .Read()
    '        mainForm.txtLnR.Text = rdr.GetValue(1)
    '        myConn.Close()
    '    End With
    'End Sub

    Public Sub incomeStatementPage(ByVal Month As Integer, ByVal Year As Integer
                          )
        myConn.Open()

        cmd = myConn.CreateCommand()
        cmd.CommandText = "incomeStatementPage"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Month", Month)
        cmd.Parameters.AddWithValue("@Year", Year)
        rdr = cmd.ExecuteReader()

        With rdr
            .Read()
            mainForm.txtPenjualanperMonth.Text = rdr.GetValue(1)
            mainForm.txtPembelianperMonth.Text = rdr.GetValue(2)
            mainForm.txtTotalCost.Text = rdr.GetValue(3)
            mainForm.txtStock.Text = rdr.GetValue(4)
            mainForm.txtHasil.Text = rdr.GetValue(5)
            mainForm.txtTotalLabaKotor.Text = rdr.GetValue(6)
            mainForm.txtLnR.Text = rdr.GetValue(7)
            myConn.Close()
        End With
    End Sub


    Public Sub shipmentHistory(ByVal Day As Integer, ByVal Month As Integer, ByVal Year As Integer, ByVal tempData As DataGridView)
        myConn.Open()

        cmd = myConn.CreateCommand()
        cmd.CommandText = "shipmentHistory"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Day", Day)
        cmd.Parameters.AddWithValue("@Month", Month)
        cmd.Parameters.AddWithValue("@Year", Year)
        cmd.ExecuteNonQuery()
        adapter = New SqlDataAdapter(cmd)
        Dim datatable As New DataTable

        adapter.Fill(datatable)
        tempData.DataSource = datatable

        myConn.Close()

    End Sub
End Class