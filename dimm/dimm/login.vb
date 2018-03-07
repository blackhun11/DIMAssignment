Public Class login

    Private SQLconn As New DataAccess

    Private Sub btn_Login_Click(sender As Object, e As EventArgs) Handles btn_Login.Click
        MessageBox.Show(SQLconn.validateLogin(txtUserEmail.Text, txtPassword.Text))

    End Sub

  
End Class
