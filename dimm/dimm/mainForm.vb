Public Class mainForm

    Const MAX_MENU_ITEM = 11

    Private SQLconn As New DataAccess
    Private Login As New login
    Private allPanels(10) As Panel



    Private Sub mainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '*********************************
        ' Database initialization
        '*********************************
        ' SQLconn.selectAllStaff(dataGridStaff)
        ' SQLconn.selectAllMarket(dataGridMarket)
        '*********************************
        ' In form object initialization
        '*********************************
        allPanels(0) = panelStaff
        allPanels(1) = panelPembelian
        allPanels(2) = panelPenjualan
        allPanels(3) = panelFinalize
        allPanels(4) = panelMarket
        allPanels(5) = panelProduct
        allPanels(6) = panelSupplier
        allPanels(7) = panelBeban
        allPanels(8) = panelMonthlyReport
        allPanels(9) = panelIncomeStatement
        allPanels(10) = panelShipmentHistory

        Dim ctr As Integer
        For ctr = 0 To 10
            allPanels(ctr).Location = New Point(188, 97)
        Next
        lblTittle.Location = New Point(313, 56)

    End Sub



    '************************************************
    'menuItemClick
    '************************************************

    Private Sub StaffToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StaffToolStripMenuItem.Click
        showMenuItem(0)
        dataGridStaff.DataSource = Nothing
        SQLconn.selectAllStaff(dataGridStaff)
        dataGridStaff.Refresh()
    End Sub
    Private Sub PembelianToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PembelianToolStripMenuItem.Click
        SQLconn.getStaffNameforSales(ComboBox2)
        showMenuItem(1)
        txtGurame.Text = 0
        txtNila.Text = 0
        txtPatin.Text = 0
        txtMas.Text = 0
        txtLele.Text = 0
    End Sub

    Private Sub PenjualanToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PenjualanToolStripMenuItem.Click
        showMenuItem(2)
        SQLconn.getStaffNameforSales(ComboBox1)
        SQLconn.getMarketforSales(comboMarket)
        txtGuramePenjualan.Text = 0
        txtNilaPenjualan.Text = 0
        txtPatinPenjualan.Text = 0
        txtMasPenjualan.Text = 0
        txtLelePenjualan.Text = 0
    End Sub

    Private Sub FinalizeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FinalizeToolStripMenuItem.Click
        showMenuItem(3)
        SQLconn.selectnotFinalizeExpense(dataGridBeban)

        SQLconn.selectnotFinalizePurchase(datagridPembelian)

        SQLconn.selectnotFinalizeSales(dataGridPenjualan)

    End Sub
    Private Sub MarketToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MarketToolStripMenuItem.Click
        dataGridMarket.DataSource = Nothing
        showMenuItem(4)
        SQLconn.selectAllMarket(dataGridMarket)
    End Sub

    Private Sub ProductToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProductToolStripMenuItem.Click
        dataGridProduct.DataBindings.Clear()
        showMenuItem(5)
        SQLconn.selectAllProduct(dataGridProduct)

    End Sub
    Private Sub SupplierToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SupplierToolStripMenuItem.Click
        showMenuItem(6)
        SQLconn.getProductforSupplier(comboProduct)
        comboProduct.Refresh()
        SQLconn.selectAllSupplier(dataGridSupplier)
        dataGridSupplier.Refresh()
    End Sub

    Private Sub BebanToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BebanToolStripMenuItem.Click
        If txtQty.Text = "" Then
            txtQty.Text = 0
        End If
        showMenuItem(7)
    End Sub



    Private Sub MonthlyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MonthlyToolStripMenuItem.Click
        showMenuItem(8)
    End Sub

    Private Sub IncomeStatementToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles IncomeStatementToolStripMenuItem.Click
        showMenuItem(9)
    End Sub

    Private Sub showMenuItem(ByVal index As Integer)
        Dim ctr As Integer
        For ctr = 0 To MAX_MENU_ITEM - 1
            allPanels(ctr).Visible = False
        Next
        allPanels(index).Visible = True
    End Sub

    '************************************************
    'viewStaff
    '************************************************

    Private Sub dataGridStaff_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dataGridStaff.CellContentClick
        txtStaffID.Text = dataGridStaff.Rows(e.RowIndex).Cells(0).Value.ToString()
        txtStaffName.Text = dataGridStaff.Rows(e.RowIndex).Cells(1).Value().ToString()
        txtPassword.Text = dataGridStaff.Rows(e.RowIndex).Cells(2).Value().ToString()
        txtAddress.Text = dataGridStaff.Rows(e.RowIndex).Cells(3).Value().ToString()
        txtEmailAddress.Text = dataGridStaff.Rows(e.RowIndex).Cells(4).Value().ToString()
        txtAltEmail.Text = dataGridStaff.Rows(e.RowIndex).Cells(5).Value().ToString()
        txtPhoneNumber.Text = dataGridStaff.Rows(e.RowIndex).Cells(6).Value().ToString()
        txtAltPhone.Text = dataGridStaff.Rows(e.RowIndex).Cells(7).Value().ToString()
    End Sub

    Private Sub btnInsert_Click(sender As Object, e As EventArgs) Handles btnInsert.Click
        dataGridStaff.DataBindings.Clear()
        SQLconn.insertStaff(txtStaffName.Text, txtPassword.Text, txtAddress.Text, txtEmailAddress.Text,
                            txtAltEmail.Text, txtPhoneNumber.Text, txtAltPhone.Text)
        MessageBox.Show("Insert Sukses")
        SQLconn.selectAllStaff(dataGridStaff)

    End Sub

    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        dataGridStaff.DataBindings.Clear()
        SQLconn.updateStaff(txtStaffID.Text, txtStaffName.Text, txtPassword.Text, txtAddress.Text, txtEmailAddress.Text,
                            txtAltEmail.Text, txtPhoneNumber.Text, txtAltPhone.Text)
        MessageBox.Show("Update Sukses")
        SQLconn.selectAllStaff(dataGridStaff)

    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        dataGridStaff.DataBindings.Clear()
        SQLconn.deleteStaff(txtStaffID.Text)

        MessageBox.Show("Delete Sukses")
        SQLconn.selectAllStaff(dataGridStaff)



    End Sub

    '************************************************
    'viewMarket
    '************************************************

    Private Sub dataGridMarket_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dataGridMarket.CellContentClick
        txtMarketID.Text = dataGridMarket.Rows(e.RowIndex).Cells(0).Value.ToString()
        txtMarketName.Text = dataGridMarket.Rows(e.RowIndex).Cells(1).Value().ToString()
        txtMarketAddress.Text = dataGridMarket.Rows(e.RowIndex).Cells(2).Value().ToString()
    End Sub

    Private Sub btnInsertMarket_Click(sender As Object, e As EventArgs) Handles btnInsertMarket.Click
        dataGridMarket.DataBindings.Clear()
        SQLconn.insertMarket(txtMarketName.Text, txtMarketAddress.Text)
        MessageBox.Show("Insert Sukses")
        SQLconn.selectAllMarket(dataGridMarket)

    End Sub

    Private Sub btnUpdateMarket_Click(sender As Object, e As EventArgs) Handles btnUpdateMarket.Click
        dataGridMarket.DataBindings.Clear()
        SQLconn.updateMarket(txtMarketID.Text, txtMarketName.Text, txtMarketAddress.Text)
        MessageBox.Show("Update Sukses")
        SQLconn.selectAllMarket(dataGridMarket)

    End Sub

    Private Sub btnDeleteMarket_Click(sender As Object, e As EventArgs) Handles btnDeleteMarket.Click
        dataGridMarket.DataBindings.Clear()
        SQLconn.deleteMarket(txtMarketID.Text)

        MessageBox.Show("Delete Sukses")
        SQLconn.selectAllMarket(dataGridMarket)

    End Sub

    '************************************************
    'viewProduct
    '************************************************

    Private Sub dataGridProduct_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dataGridProduct.CellContentClick
        txtProductID.Text = dataGridProduct.Rows(e.RowIndex).Cells(0).Value.ToString()
        txtProductName.Text = dataGridProduct.Rows(e.RowIndex).Cells(1).Value().ToString()
        txtProductSale.Text = dataGridProduct.Rows(e.RowIndex).Cells(2).Value().ToString()
        txtProductPurchase.Text = dataGridProduct.Rows(e.RowIndex).Cells(3).Value().ToString()
    End Sub

    Private Sub btnInsertProduct_Click(sender As Object, e As EventArgs) Handles btnInsertProduct.Click
        SQLconn.insertProduct(txtProductName.Text, txtProductSale.Text, txtProductPurchase.Text)
        MessageBox.Show("Insert Sukses")
        SQLconn.selectAllProduct(dataGridProduct)
        dataGridProduct.Refresh()
    End Sub

    Private Sub btnUpdateProduct_Click(sender As Object, e As EventArgs) Handles btnUpdateProduct.Click
        SQLconn.updateProduct(txtProductID.Text, txtProductName.Text, txtProductSale.Text, txtProductPurchase.Text)
        MessageBox.Show("Update Sukses")
        SQLconn.selectAllProduct(dataGridProduct)
        dataGridProduct.Refresh()
    End Sub

    Private Sub btnDeleteProduct_Click(sender As Object, e As EventArgs) Handles btnDeleteProduct.Click
        SQLconn.deleteProduct(txtProductID.Text)
        MessageBox.Show("Delete Sukses")
        SQLconn.selectAllProduct(dataGridProduct)
        dataGridProduct.Refresh()
    End Sub

    '************************************************
    'viewSupplier
    '************************************************
    Private Sub dataGridSupplier_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dataGridSupplier.CellContentClick
        txtSupplierID.Text = dataGridSupplier.Rows(e.RowIndex).Cells(0).Value.ToString()
        txtSupplierName.Text = dataGridSupplier.Rows(e.RowIndex).Cells(1).Value().ToString()
        txtSupplierAddress.Text = dataGridSupplier.Rows(e.RowIndex).Cells(2).Value().ToString()
        txtSupplierEmail.Text = dataGridSupplier.Rows(e.RowIndex).Cells(3).Value().ToString()
        txtSupplierAltEmail.Text = dataGridSupplier.Rows(e.RowIndex).Cells(4).Value().ToString()
        txtSupplierPhone.Text = dataGridSupplier.Rows(e.RowIndex).Cells(5).Value().ToString()
        txtSupplierAltPhone.Text = dataGridSupplier.Rows(e.RowIndex).Cells(6).Value().ToString()
        comboProduct.SelectedItem = dataGridSupplier.Rows(e.RowIndex).Cells(10).Value().ToString()
    End Sub

    Private Sub btnInsertSupplier_Click(sender As Object, e As EventArgs) Handles btnInsertSupplier.Click
        SQLconn.insertSupplier(txtSupplierName.Text, txtSupplierAddress.Text, txtSupplierEmail.Text, txtSupplierAltEmail.Text, txtSupplierPhone.Text, txtSupplierAltPhone.Text,
                               comboProduct.text)
        MessageBox.Show("Insert Sukses")
        SQLconn.selectAllSupplier(dataGridSupplier)
    End Sub

    Private Sub btnUpdateSupplier_Click(sender As Object, e As EventArgs) Handles btnUpdateSupplier.Click
        SQLconn.updateSupplier(txtSupplierID.Text, txtSupplierName.Text, txtSupplierAddress.Text, txtSupplierEmail.Text, txtSupplierAltEmail.Text, txtSupplierPhone.Text, txtSupplierAltPhone.Text,
                               comboProduct.Text)
        MessageBox.Show("Update Sukses")
        SQLconn.selectAllSupplier(dataGridSupplier)
    End Sub

    Private Sub btnDeleteSupplier_Click(sender As Object, e As EventArgs) Handles btnDeleteSupplier.Click
        SQLconn.deleteSupplier(txtSupplierID.Text)
        MessageBox.Show("Delete Sukses")
        SQLconn.selectAllSupplier(dataGridSupplier)
    End Sub
    '************************************************
    'BebanAwal
    '************************************************
    Private Sub btnSaveExpense_Click(sender As Object, e As EventArgs) Handles btnSaveExpense.Click
        SQLconn.insertNewExpenses(txtPerihal.Text, txtBiaya.Text, txtQty.Text)
        dataGridBebanAwal.Visible = True
        SQLconn.selectnotFinalizeExpense(dataGridBebanAwal)

    End Sub


    '*********************************
    ' CheckBox Pembelian
    '*********************************
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles checkGurame.CheckedChanged
        If (checkGurame.Checked = True) Then
            txtGurame.Enabled = True
            'tempProductID(productCount) = "cb1"
            'productCount = productCount + 1
            ' checkGurame.Tag = productCount
        Else
            txtGurame.Enabled = False
            txtGurame.Text = 0
            'productCount = productCount - 1
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles checkNila.CheckedChanged
        If (checkNila.Checked = True) Then
            txtNila.Enabled = True
            'tempProductID(productCount) = "cb2"
            'productCount = productCount + 1
            'checkNila.Tag = productCount
        Else
            txtNila.Enabled = False
            txtNila.Text = 0
            'productCount = productCount - 1
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles checkPatin.CheckedChanged
        If (checkPatin.Checked = True) Then
            txtPatin.Enabled = True
            'tempProductID(productCount) = "cb3"
            'productCount = productCount + 1
            ' checkPatin.Tag = productCount
        Else
            txtPatin.Enabled = False
            txtPatin.Text = 0
            'productCount = productCount - 1
        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles checkMas.CheckedChanged
        If (checkMas.Checked = True) Then
            txtMas.Enabled = True
            'tempProductID(productCount) = "cb4"
            'productCount = productCount + 1
            'checkMas.Tag = productCount
        Else
            txtMas.Enabled = False
            txtMas.Text = 0
            'productCount = productCount - 1
        End If
    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles checkLele.CheckedChanged
        If (checkLele.Checked = True) Then
            txtLele.Enabled = True
            'tempProductID(productCount) = "cb5"
            'productCount = productCount + 1
            'checkLele.Tag = productCount
        Else
            txtLele.Enabled = False
            txtLele.Text = 0
            'productCount = productCount - 1
        End If
    End Sub
    '************************************************
    'Pembelian Awal
    '************************************************
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        SQLconn.insertPurchase(ComboBox2.Text, checkGurame.Text, checkNila.Text, checkPatin.Text,
                               checkMas.Text, checkLele.Text,
                               txtGurame.Text, txtNila.Text,
                               txtPatin.Text, txtMas.Text,
                               txtLele.Text)
        If (ComboBox2.Text = Nothing) Then
            MessageBox.Show("Please Select Staff")
        Else
            MessageBox.Show("Success")
            dataGridPembelianAwal.Visible = True
            SQLconn.selectnotFinalizePurchase(dataGridPembelianAwal)

        End If

    End Sub



    '*********************************
    ' CheckBox Penjualan
    '*********************************

    Private Sub checkGuramePenjualan_CheckedChanged(sender As Object, e As EventArgs) Handles checkGuramePenjualan.CheckedChanged
        If (checkGuramePenjualan.Checked = True) Then
            txtGuramePenjualan.Enabled = True
        Else
            txtGuramePenjualan.Enabled = False
            txtGuramePenjualan.Text = 0
        End If


    End Sub

    Private Sub checkNilaPenjualan_CheckedChanged(sender As Object, e As EventArgs) Handles checkNilaPenjualan.CheckedChanged
        If (checkNilaPenjualan.Checked = True) Then
            txtNilaPenjualan.Enabled = True
        Else
            txtNilaPenjualan.Enabled = False
            txtNilaPenjualan.Text = 0
        End If
    End Sub

    Private Sub checkPatinPenjualan_CheckedChanged(sender As Object, e As EventArgs) Handles checkPatinPenjualan.CheckedChanged
        If (checkPatinPenjualan.Checked = True) Then
            txtPatinPenjualan.Enabled = True
        Else
            txtPatinPenjualan.Enabled = False
            txtPatinPenjualan.Text = 0
        End If
    End Sub

    Private Sub checkMasPenjualan_CheckedChanged(sender As Object, e As EventArgs) Handles checkMasPenjualan.CheckedChanged
        If (checkMasPenjualan.Checked = True) Then
            txtMasPenjualan.Enabled = True
        Else
            txtMasPenjualan.Enabled = False
            txtMasPenjualan.Text = 0
        End If
    End Sub

    Private Sub checkLelePenjualan_CheckedChanged(sender As Object, e As EventArgs) Handles checkLelePenjualan.CheckedChanged
        If (checkLelePenjualan.Checked = True) Then
            txtLelePenjualan.Enabled = True
        Else
            txtLelePenjualan.Enabled = False
            txtLelePenjualan.Text = 0
        End If
    End Sub

    '************************************************
    'Penjualan Awal 
    '************************************************


    Private Sub btnSavePenjualan_Click(sender As Object, e As EventArgs) Handles btnSavePenjualan.Click
        SQLconn.insertSales(ComboBox1.Text, comboMarket.Text,
                               checkGuramePenjualan.Text, checkNilaPenjualan.Text,
                               checkPatinPenjualan.Text,
                               checkMasPenjualan.Text, checkLelePenjualan.Text,
                               txtGuramePenjualan.Text, txtNilaPenjualan.Text,
                               txtPatinPenjualan.Text, txtMasPenjualan.Text,
                               txtLelePenjualan.Text)
        If (ComboBox1.Text = Nothing Or comboMarket.Text = Nothing) Then

            MessageBox.Show("Your StaffName or Market Destination is empty")
        Else
            MessageBox.Show("Success")
            dataGridPenjualanAwal.Visible = True
            SQLconn.selectnotFinalizePurchase(dataGridPenjualanAwal)

        End If
    End Sub



    Private Sub btnFinalize_Click(sender As Object, e As EventArgs) Handles btnFinalize.Click
        SQLconn.isFinalized()
        MessageBox.Show("Success")
        SQLconn.selectnotFinalizeExpense(dataGridBeban)

        SQLconn.selectnotFinalizePurchase(datagridPembelian)

        SQLconn.selectnotFinalizeSales(dataGridPenjualan)


    End Sub

    Private Sub btnViewMonthlyReport_Click(sender As Object, e As EventArgs) Handles btnViewMonthlyReport.Click
        If txtBulan.Text = Nothing Or txtTahun.Text = Nothing Then
            txtBulan.Text = DateTime.Now.Month.ToString()
            txtTahun.Text = DateTime.Now.Year.ToString()
        End If
        SQLconn.monthlyReport(txtBulan.Text, txtTahun.Text, dataGridMonthlyReport)
        dataGridMonthlyReport.Visible = True
    End Sub

 
   
    Private Sub CheckBox1_CheckedChanged_1(sender As Object, e As EventArgs) Handles checkLnR.CheckedChanged

        If checkLnR.Checked = True Then
            dataGridLnR.Visible = True
            txtLnRBulan.Enabled = False
            If txtLnRBulan.Text = Nothing Or txtLnRTahun.Text = Nothing Then
                txtLnRBulan.Text = DateTime.Now.Month.ToString()
                txtLnRTahun.Text = DateTime.Now.Year.ToString()
            End If
            txtLnRTahun.Enabled = False
            SQLconn.incomeStatementPage(txtLnRBulan.Text, txtLnRTahun.Text)
            ' SQLconn.stockPermonth(txtLnRBulan.Text, txtLnRTahun.Text)
            'SQLconn.netPurchase(txtLnRBulan.Text, txtLnRTahun.Text)
            SQLconn.viewBiaya(txtLnRBulan.Text, txtLnRTahun.Text, dataGridLnR)
            ' SQLconn.grossProfit(txtLnRBulan.Text, txtLnRTahun.Text)
            ' SQLconn.incomeStatement(txtLnRBulan.Text, txtLnRTahun.Text)

        Else
            dataGridLnR.Visible = False
            txtLnRBulan.Enabled = True
            txtLnRTahun.Enabled = True
        End If
    End Sub

    Private Sub ToolStripLabel2_Click(sender As Object, e As EventArgs) Handles ToolStripLabel2.Click
        Me.Close()
        MessageBox.Show("Logout Sukses")
        Login.Show()

    End Sub

   
    Private Sub ToolStripLabel3_Click(sender As Object, e As EventArgs) Handles ToolStripLabel3.Click
        txtDayShip.Text = DateTime.Now.Day.ToString()
        txtMonthShip.Text = DateTime.Now.Month.ToString()
        txtYearShip.Text = DateTime.Now.Year.ToString()
        SQLconn.shipmentHistory(txtDayShip.Text, txtMonthShip.Text, txtYearShip.Text, dataGridShip)
        showMenuItem(10)
    End Sub

    Private Sub btnSearchShip_Click(sender As Object, e As EventArgs) Handles btnSearchShip.Click
  
        SQLconn.shipmentHistory(txtDayShip.Text, txtMonthShip.Text, txtYearShip.Text, dataGridShip)
    End Sub

  

End Class