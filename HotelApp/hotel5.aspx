<%@ Page Language = "VB" %>
<%@ Import Namespace = "System.Data.OleDb" %>
<%@ Import Namespace = "System.IO" %>
<%@ Import Namespace = "System.Drawing" %>

<!DOCTYPE html>
<html xmlns = "http://www.w3.org/1999/xhtml">
    <head id = "Head1" runat = "server">
        <title>View Staff Info</title>
        <script runat = "server">
            Public Class PictureBox
                Property Image As Drawing.Bitmap
            End Class

            Sub Search_Click(Src As Object, E As EventArgs)
                Try
                    'Connect to the Database
                    Dim cnAccess As New OleDbConnection(
                    "Provider = Microsoft.ACE.OLEDB.12.0;" &
                    "Data Source = C:\Users\icent\Documents\Yulia's files\Fall 2020\ITMD 422\Labs\Lab 6\HigginsHotelSystems.accdb")

                    cnAccess.Open()

                    Dim sLName As String
                    sLName = LName.Text.Trim

                    'Construct the SELECT statement
                    Dim sSelectSQL As String

                    'Create the SQL Select Statement
                    sSelectSQL = "SELECT * FROM Staff WHERE ([LName] Like '" & sLName & "')"

                    'Create the OleDbCommand object
                    Dim cmdSelect As New OleDbCommand(sSelectSQL, cnAccess)
                    Dim drEmp As OleDbDataReader, sbResults As New StringBuilder()
                    drEmp = cmdSelect.ExecuteReader()
                    sbResults.Append("<table>")
                    Do While drEmp.Read()
                        sbResults.Append("<table>")
                        sbResults.Append("<tr><td><b>Staff ID: </b>")
                        sbResults.Append(drEmp.GetValue(0).ToString)
                        sbResults.Append("</td></tr><tr><td><b>  Name: </b>")
                        sbResults.Append(drEmp.GetString(1))
                        sbResults.Append("</td></tr><tr><td><b>  Last Name: </b>")
                        sbResults.Append(drEmp.GetString(2))
                        sbResults.Append("</td></tr><tr><td><b>  Hire Date: </b>")
                        sbResults.Append(drEmp.GetValue(3).ToString)
                        sbResults.Append("</td></tr><img id='image' width='200' height='300' src=' ")
                        sbResults.Append(drEmp.GetString(4))
                        sbResults.Append("' />")
                        sbResults.Append("</table>")
                        sbResults.Append("<br></br>")
                    Loop

                    sbResults.Append("</table>")

                    msg.Text = sbResults.ToString()

                    cnAccess.Close()
                    cnAccess = Nothing

                    Response.Write("Data Found!")
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
        <h3>Enter Staff Name</h3>
        <form runat = "server" id = "form1">
            <div>
                <table>
                    <tr>
                        <td>Last Name: </td>
                        <td><asp:Textbox id = "LName" runat = "server" /></td>
                    </tr>
                </table>
            </div>
            <div>
                <p><asp:Button Text = "Search" OnClick = "Search_Click" runat = "server" ID = "Button1" /></p>
                <p><asp:Label id = "msg" runat = "server" /></p>
            </div>
            <div>
                <p><asp:Button Text = "Back to Main Menu" OnClick = "GoTo_Click" runat = "server" ID = "Button2" /></p>
            </div>
        </form>
    </body>
</html>
