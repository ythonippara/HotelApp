<%@ Page Language = "VB" %>
<%@ Import Namespace = "System.Data.OleDb" %>
<!DOCTYPE html>
<html xmlns = "http://www.w3.org/1999/xhtml">
    <head id="Head1" runat = "server">
        <title>Connection</title>
        <script runat = "server">
            Sub Insert_Click(Src As Object, E As EventArgs)
                Try
                    'Connect to the Database
                    Dim cnAccess As New OleDbConnection(
                    "Provider = Microsoft.ACE.OLEDB.12.0;" &
                    "Data Source = C:\Users\icent\Documents\Yulia's files\Fall 2020\ITMD 422\Labs\Lab 6\HigginsHotelSystems.accdb")

                    cnAccess.Open()
                    Dim sID, sFName, sLName, sZip, sState, sCardNo, sInsertSQL As String
                    sID = GuestID.Text
                    sFName = FName.Text
                    sLName = LName.Text
                    sZip = Zip.Text
                    sState = State.Text
                    sCardNo = CardNo.Text

                    'Construct the insert statement
                    sInsertSQL = "INSERT INTO Guests(" &
                        "[GuestID], [LName], [FName], [ZipCode], [StateIDCard], [CreditCardNo]) VALUES" &
                        "(" & sID & ",'" & sLName & "','" & sFName & "'," & sZip & ",'" & sState & "', " & sCardNo & ");"

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
                Response.Redirect("hotel3.aspx")
            End Sub
        </script>
    </head>
    <body style = "font-family:Tahoma;">
        <h3>Enter Guest Details</h3>
        <form runat = "server" id = "form1">
            <div>
                <table>
                    <tr>
                        <td>Guest ID: </td>
                        <td><asp:Textbox id = "GuestID" runat="server" /></td>
                    </tr>
                    <tr>
                        <td>Last Name: </td>
                        <td>
                            <asp:TextBox ID="LName" runat="server" /></td>
                    </tr>
                    <tr>
                        <td>First Name: </td>
                        <td><asp:Textbox id = "FName" runat = "server" /></td>
                    </tr>
                    <tr>
                        <td>Zip Code: </td>
                        <td><asp:Textbox id = "Zip" runat = "server" /></td>
                    </tr>
                    <tr>
                        <td>State ID: </td>
                        <td><asp:Textbox id = "State" runat = "server" /></td>
                    </tr>
                    <tr>
                        <td>Credit Card: </td>
                        <td><asp:Textbox id = "CardNo" runat = "server" /></td>
                    </tr>
                </table>
            </div>
            <br />
            <div>
                <asp:Button Text = "Insert" OnClick = "Insert_Click" runat = "server" ID = "Button1" />
                <p>
                    <asp:Label id = "msg" runat = "server" />
                </p>
            </div>
            <br />
            <div>
                <asp:Button Text = "Retrieve Records" OnClick = "GoTo_Click" runat = "server" ID = "Button2" />
            </div>
        </form>
    </body>
</html>