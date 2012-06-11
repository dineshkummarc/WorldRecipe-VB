<%--
'++++++++++++++++++++++++++++++++++++++++++++
'+ World Recipe Directory v2.7 ASP .NET
'+ Programmer: Dexter Zafra, Norwalk, CA. USA
'+ Website: www.Ex-designz.net & www.myasp-net.com
'+ Creation Date: June 25, 2005 
'+   
'+Purpose: This ASP.NET application was developed by (Dexter Zafra) of www.ex-designz.net and www.myasp-net.com to help classic 
'+ASP coders and beginners to learn ASP.NET in VB without having to use or rely too much on Visual Studio.NET, and applying  the skills 
'+and tricks they've learned from classic ASP. The advantage of hand coding is the total control of the HTML,JavaScript and most of all CSS-P layout. This does not mean to completly eliminate/ignore the use of VS.NET. For a large enterprise project, VS.NET is the best tools to use. 
'+
'++++++++++++++++++++++++++++++++++++++++++++
--%>

<%@ Page Language="VB" Debug="True" %>
<%@ import Namespace="System.Data.Oledb" %>

<script runat="server">

   'When button is click check login and password with database users table
   'If login and password is correct starts a user session

   Sub btnlogin_Click(Sender as Object, e as EventArgs)

        'Declared Login fields variables
        Dim strUname as string
        Dim strPass as string
        Dim strSQLQuery as string
        Dim CheckUname as string
        Dim Checkpass as string

        CheckUname = uname.Text
        Checkpass = password.Text
                
        'Display the login error message
        lblinvalid.text = "Your username and / or password were incorrect. Please try again."

        'Database SQL query
        strSQLQuery = "SELECT * FROM users"
        DataBase_Connect(strSQLQuery)

        Do while (objDataReader.Read())

            strUname=objDataReader("uname")
            strPass=objDataReader("password")

            if CheckUname = "" And Checkpass = "" Then
                lblinvalid.text = "You must enter a username and password. Please try again."
                lblinvalid.Visible = True
            elseif CheckUname = "" Then
                lblinvalid.text = "You must enter a user name. Please try again."
                lblinvalid.Visible = True
            elseif Checkpass = "" Then
                lblinvalid.text = "You must enter a password. Please try again."
                lblinvalid.Visible = True
            elseif (strUname <> Request.Form("uname")) Then
                lblinvalid.text = "Your username is incorrect. Please try again."
                lblinvalid.Visible = True
            elseif (strPass <> Request.Form("password")) Then
                lblinvalid.text = "Your password is incorrect. Please try again."
                lblinvalid.Visible = True
            else
                Session("userid") = strUname
                Response.Redirect("recipemanager.aspx")
            end if

        loop

        'Close the database connection
        DataBase_Disconnect()

   End Sub

</script>

<!--#include file="inc_admindbconn.aspx"-->

<!--Powered By www.Ex-designz.net Recipe Cookbook ASP.NET version - Author: Dexter Zafra, Norwalk,CA-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Recipe Manager - www.ex-designz.net</title>
<style type="text/css" media="screen">@import "../css/cssreciaspx.css";</style>
</head>
<body>
<br />
<br />
<div align="center">
<h3>Demo Recipe Manager</h3>
<span class="content2">
<br />
<br />
<b>Username:</b> admin
<br />
<b>Password:</b> admin
</span>
</div>
<br />
    <form runat="server">
        <div style="text-align: center; margin-bottom: 12px;"><asp:Label runat="server" class="cred" id="lblinvalid" Visible=false /></div>
        <table cellspacing="1" cellpadding="2" width="260" bgcolor="#ffffff" border="0" align="center">
                <tr>
                    <td align="left" bgcolor="#6898d0" colspan="2">
                        <span class="content3">Recipe Admin Login</span>
                    </td>
                </tr>
                <tr>
                    <td width="80" bgcolor="#f7f7f7">
                        <span class="content2">User Name:</span>
                    </td>
                    <td bgcolor="#fbfbfb">
                       <asp:TextBox runat="server" id="uname" class="textbox" size="20" />
                    </td>
                </tr>
                <tr>
                    <td width="80" bgcolor="#f7f7f7">
                        <span class="content2">Password</span>
                    </td>
                    <td bgcolor="#fbfbfb">
                        <asp:TextBox runat="server" id="password" class="textbox" size="20" textmode="password" />
                    </td>
                </tr>
                <tr>
                    <td valign="bottom" align="middle" bgcolor="#f7f7f7" colspan="2">
                        <asp:Button runat="server" class="submit" OnClick="btnlogin_Click" Text="Login"/>
                    </td>
                </tr>
        </table>
    </form>

</body>
</html>
