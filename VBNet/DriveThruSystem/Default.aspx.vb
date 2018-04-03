
Partial Class _Default
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            UpdateProgress1.DisplayAfter = 100
            UpdateProgress1.Visible = True
            UpdateProgress1.DynamicLayout = True
        End If
    End Sub


    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click

        System.Threading.Thread.Sleep(5000)
        Label1.Text = "Welcome To ThaiCreate.Com : " & DateTime.Now.ToLongTimeString()

    End Sub

End Class
