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

<%@ Page Language="VB" Debug="true" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.Oledb" %>

<script runat="server">

'++++++++++++++++++++++++++++++++++++++++++++
'Handles page load events
   Sub Page_Load()
    
           Dim strSQL as string
           Dim totalcomments as integer
        
           'Call check user function - Check if user has started a session 
           Check_User()         
       
            'SQL display details and rating value
            strSQL = "SELECT *, (RATING/NO_RATES) AS Rates FROM Recipes WHERE id=" & Request.QueryString("id")
    
            'Connect to the database
            DataBase_Connect(strSQL)
    
            'Read data
            objDataReader.Read()

            Dim strAproveStat as Integer
            strAproveStat = objDataReader("LINK_APPROVED")

            If strAproveStat > 0 Then
          
               approvebutton.visible = False
               lblapprovalstatus.text = "Viewing Recipe"

            Else

            lblapprovalstatus.text = "Unapprove - This recipe is waiting for approval"

            End If
    
            lblname.text = objDataReader("Name")
            lblauthor.text = objDataReader("Author")
            lblhits.Text = objDataReader("Hits")
            lbldate.Text = objDataReader("Date")    
            lblCatName.Text = objDataReader("Category")           
            Ingredients.text = objDataReader("Ingredients")
            Instructions.text = objDataReader("Instructions")
            
            'Close database connection
            DataBase_Disconnect()
    
  End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
 Sub Approve_Recipe(sender as object, e as System.EventArgs)

            Dim strSQL as string

            strSQL = "Update Recipes set LINK_APPROVED = 1 where id=" & Request.QueryString("id")

            'Call Open database - connect to the database
            objConnection = New OledbConnection(strConnection)
            objCommand = New OledbCommand(strSQL, objConnection)

            objConnection.Open()
            objCommand.ExecuteNonQuery()
            
            'Close the db connection and free up memory
            objCommand = nothing
            objConnection.Close()
            objConnection = nothing

            Dim strSQL2 as string

           'SQL insert rating   
           strSQL2 = "Update Recipes  SET RATING = RATING + " & 5 & ", NO_RATES = NO_RATES + 1 WHERE ID =" & Request.QueryString("id")
        
           'Call Open database - connect to the database
            'Call Open database - connect to the database
            objConnection = New OledbConnection(strConnection)
            objCommand = New OledbCommand(strSQL2, objConnection)

           objConnection.Open()
           objCommand.ExecuteNonQuery()

           objConnection.Close()
           objConnection = nothing

            Dim address as string

            address = "recipemanager.aspx?tab=1"
            Server.Transfer(address)

  End Sub
'++++++++++++++++++++++++++++++++++++++++++++

</script>

<!--#include file="inc_admindbconn.aspx"-->

<!--Powered By www.Ex-designz.net Recipe Cookbook ASP.NET version - Author: Dexter Zafra, Norwalk,CA-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Admin Reviewing Recipe - www.ex-designz.net</title>
<style type="text/css" media="screen">@import "../css/cssreciaspx.css";</style>
</head>
<body>
       <form runat="server">
           <table width=100% height=100%>
               <tr>
	        <td valign="middle">
                       <table width=40% border="0" cellpadding="0" cellspacing="0" align="center">
                           <tr>
            	                   <td colspan=2  bgcolor="#6898d0">
                                    <div class="roundcont2">
                                      <div class="roundtop">
                                       <img src="../images/hleft.gif" height="5" width="5" alt="" class="corner">
                                     <div style="text-align: left; padding-left: 6px;padding-bottom: 2px;"><asp:Label id="lblapprovalstatus" cssclass="content3" runat="server" />
                 </div> 
             </div>
        </div>
    </td>
</tr>
       <tr>
             <td bgcolor="#ffffff">
            	     <table width="100%" border=0 cellpadding=3 cellspacing=0>       

            		                  <tr>
          <td bgcolor="#F4F9FF" class="content2">Name:</td>   		
         <td bgcolor="#FBFDFF">
<asp:Label runat="server" id="lblname" class="content2" />
            			                 </td>
            		                  </tr>
<tr>
          <td bgcolor="#F4F9FF" class="content2">Category:</td>   		
         <td bgcolor="#FBFDFF">
<asp:Label runat="server" id="lblCatName" class="content2" />
            			                 </td>
            		                  </tr>
                     <tr>
                       <td bgcolor="#F4F9FF" class="content2">Author:</td>   		
                       <td bgcolor="#FBFDFF">
      <asp:Label runat="server" id="lblauthor" class="content2" />
            			 </td>        		                  
                                   </tr>

            		                  <tr>
          <td bgcolor="#F4F9FF" class="content2">Hits:</td>   		
         <td bgcolor="#FBFDFF">
<asp:Label runat="server" id="lblhits" class="content2" />
            			                 </td>
            		                  </tr>
                     <tr>
                       <td bgcolor="#F4F9FF" class="content2">Date:</td>   		
                       <td bgcolor="#FBFDFF">
                     <asp:Label runat="server" id="lbldate" class="content2" />
            			 </td>          		                  
                                    </tr>

            		                  <tr>
            			         <td valign="top" bgcolor="#F4F9FF" class="content2">Ingredients:</td>
            			            <td bgcolor="#FBFDFF">
 <asp:TextBox runat="server" id="Ingredients" Class="textbox" textmode="multiline" columns="70" rows="14" readonly />
            			                 </td>
            		                  </tr>
                                           <tr>
            			            <td valign="top" bgcolor="#F4F9FF" class="content2">Instructions:</td>  		
            			            <td bgcolor="#FBFDFF">
 <asp:TextBox runat="server" id="Instructions" Class="textbox" textmode="multiline" columns="70" rows="14" readonly />
<br />
<div style="text-align: left;" class="content2"><asp:Button runat="server" Text="Approve This Recipe" id="approvebutton" class="submit" onclick="Approve_Recipe"/></div>
            			                 </td>
            		                  </tr>          		                  
            	                   </table>
<br />
<div style="text-align: center;" class="content2"><asp:HyperLink runat="server" NavigateUrl="JavaScript:onClick= window.close()" class="content2">Close Window</asp:HyperLink></div>
                                </td>
		                    </tr>
		               </table>
	               </td>
               </tr>
           </table>
    </form>

    </body>
</html>
