<%@ Page Language = "VB" %>
<%@ Import Namespace = "System.Data.OleDb" %>
<!DOCTYPE html>
<html xmlns = "http://www.w3.org/1999/xhtml">
<head id = "Head1" runat = "server">
    <title>Insert Staff Records</title>
    <script runat = "server">
        Sub Create_Click(Src As Object, E As EventArgs)
            Try
                'Connect to the Database
                Dim cnAccess As New OleDbConnection(
                "Provider = Microsoft.ACE.OLEDB.12.0;" &
                "Data Source = C:\Users\icent\Documents\Yulia's files\Fall 2020\ITMD 422\Labs\Lab 6\HigginsHotelSystems.accdb")
                Dim sSelectSQL As String = "CREATE TABLE Staff"
                sSelectSQL &= "([StaffID] Number Primary Key, [FName] Text(20),"
                sSelectSQL &= "[LName] Text(20), [HireDate] Date,"
                sSelectSQL &= "[StaffPic] Text)"

                Dim cmdSelect As New OleDbCommand(sSelectSQL, cnAccess)
                cnAccess.Open()
                cmdSelect.ExecuteNonQuery()
                cnAccess.Close()
                msg.Text = "Table Created!"

            Catch ex As Exception
                msg.Text = ex.Message
                ' Response.Write("Table Exists or Connection Failed")
            End Try
        End Sub

        Sub Insert_Click(Src As Object, E As EventArgs)
            Try
                'Connect to the Database
                Dim cnAccess As New OleDbConnection(
                "Provider = Microsoft.ACE.OLEDB.12.0;" &
                "Data Source = C:\Users\icent\Documents\Yulia's files\Fall 2020\ITMD 422\Labs\Lab 6\HigginsHotelSystems.accdb")

                cnAccess.Open()
                Dim sID, sFName, sLName, sDate, sInsertSQL As String
                sID = StaffID.Text
                sFName = FName.Text
                sLName = LName.Text
                sDate = HireDate.Text

                'Construct the insert statement
                sInsertSQL = "INSERT INTO Staff(" &
                    "[StaffID], [FName], [LName], [HireDate]) VALUES" &
                    "(" & sID & ",'" & sFName & "','" & sLName & "','" & sDate & "');"

                'Construct the OleDbCommand object
                Dim cmdInsert As New OleDbCommand(sInsertSQL, cnAccess)

                'since this is not a query, we do not expect to return data 
                cmdInsert.ExecuteNonQuery()

                Response.Write("Data Recorded!")
            Catch ex As Exception
                Response.Write(ex.Message)
                Response.Write("Connection Failed")
            End Try
        End Sub

        Sub GoTo_Click(Src As Object, E As EventArgs)
            Response.Redirect("menu.aspx")
        End Sub
    </script>
</head>
    <body style = "font-family:Tahoma;">
        <h3>Higgins Hotel Systems</h3>
        <form runat = "server" id = "form1">
            <div>
                <asp:Button Text = "Create Table" OnClick = "Create_Click" runat = "server" ID = "Button1" />
                <p><asp:Label id = "msg" runat = "server" /></p>
            </div>
            <div>
                <h3>Enter Staff Member Details</h3>
                <table>
                    <tr>
                        <td>Staff ID: </td>
                        <td><asp:Textbox id = "StaffID" runat="server" /></td>
                    </tr>
                    <tr>
                        <td>First Name: </td>
                        <td><asp:Textbox id = "FName" runat = "server" /></td>
                    </tr>
                    <tr>
                        <td>Last Name: </td>
                        <td><asp:Textbox id = "LName" runat = "server" /></td>
                    </tr>
                    <tr>
                        <td>Hire Date: </td>
                        <td><asp:Textbox id = "HireDate" runat = "server" /></td>
                    </tr>
                </table>
            </div>
            <div>
                <p><asp:Button Text = "Insert" OnClick = "Insert_Click" runat = "server" ID = "Button2" /></p>
                <p><asp:Button Text = "Back to Main Menu" OnClick = "GoTo_Click" runat = "server" ID = "Button3" /></p>
            </div>
        </form>
    </body>
</html>
