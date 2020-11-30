<%@ Page Language = "VB" %>
<%@ Import Namespace = "System.Data.OleDb" %>
<!DOCTYPE html>
<html xmlns = "http://www.w3.org/1999/xhtml">
<head id = "Head1" runat = "server">
    <title>Main Menu</title>
    <script runat = "server">
        Sub GoTo_Click1(Src As Object, E As EventArgs)
            Response.Redirect("hotel1.aspx")
        End Sub
        Sub GoTo_Click2(Src As Object, E As EventArgs)
            Response.Redirect("hotel2.aspx")
        End Sub
        Sub GoTo_Click3(Src As Object, E As EventArgs)
            Response.Redirect("hotel3.aspx")
        End Sub
        Sub GoTo_Click4(Src As Object, E As EventArgs)
            Response.Redirect("hotel4.aspx")
        End Sub
        Sub GoTo_Click5(Src As Object, E As EventArgs)
            Response.Redirect("hotel5.aspx")
        End Sub
    </script>
</head>
    <body style = "font-family:Tahoma;">
        <h3>Higgins Hotel Web App</h3>
        <h3>Main Menu</h3>
        <form runat = "server" id = "form1">
            <div>
                <p>Click Option 1 to Create Guest Table</p>
                <asp:Button Text = "Option 1" OnClick = "GoTo_Click1" runat = "server" ID = "Button1" />
            </div>
            <div>
                <p>Click Option 2 to Insert Records</p>
                <asp:Button Text = "Option 2" OnClick = "GoTo_Click2" runat = "server" ID = "Button2" />
            </div>
            <div>
                <p>Click Option 3 to Retrieve Guest Records</p>
                <asp:Button Text = "Option 3" OnClick = "GoTo_Click3" runat = "server" ID = "Button3" />
            </div>
            <div>
                <p>Click Option 4 to Create Staff Table</p>
                <asp:Button Text = "Option 4" OnClick = "GoTo_Click4" runat = "server" ID = "Button4" />
            </div>
            <div>
                <p>Click Option 5 to View Staff Records</p>
                <asp:Button Text = "Option 5" OnClick = "GoTo_Click5" runat = "server" ID = "Button5" />
            </div>
        </form>
    </body>
</html>
