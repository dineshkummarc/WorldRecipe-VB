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
    
            Check_User()

           'Call recipe count
            DisplayRecipeCount()

            'Call count unpprove recipes
            UnApproveRecipe()

            DisplayCategoryCount()

            'Call count total comments
            DisplayCommentsCount()


            GetDropdownCatName()

            GetDropdownCatID()

            'Display admin user name
            lblusername.Text = "Welcome Admin:&nbsp;" & session("userid")
    
    	    'Check which action were selected, edit a recipe or delete a recipe
            strSQL = "SELECT * FROM Recipes WHERE id=" & Request.QueryString("id") 
    
            DataBase_Connect(strSQL)   
            objDataReader.Read()

            'This will be the value to be populated into the textboxes
            Name.text = objDataReader("Name")
            Author.text = objDataReader("Author")
            Hits.text = objDataReader("Hits")
            Ingredients.text = objDataReader("Ingredients")
            Instructions.text = objDataReader("Instructions")
    
            DataBase_Disconnect()
    
 End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
  Sub UnApproveRecipe()

Dim CmdCount As New OleDbCommand("Select Count(ID) From Recipes Where LINK_APPROVED = 0", New OleDbConnection(strConnection))
        CmdCount.Connection.Open()
        lblunapproved.Text = "Recipe Approval:&nbsp;" & CmdCount.ExecuteScalar() 
        CmdCount.Connection.Close()

  End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
  Sub DisplayCommentsCount()

        Dim CmdCount As New OleDbCommand("Select Count(ID) From COMMENTS_RECIPE", New OleDbConnection(strConnection))
        CmdCount.Connection.Open()
        lbCountComments.Text = "Total Comments:&nbsp;" & CmdCount.ExecuteScalar()
        CmdCount.Connection.Close()

  End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
   Sub DisplayRecipeCount()

        Dim CmdCount As New OleDbCommand("Select Count(ID) From Recipes", New OleDbConnection(strConnection))
        CmdCount.Connection.Open()
        lbCountRecipe.Text = "Total Recipes:&nbsp;" & CmdCount.ExecuteScalar()
        CmdCount.Connection.Close()

   End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
  'Get the number of categoy
  Sub DisplayCategoryCount()

        Dim CmdCount As New OleDbCommand("Select Count(CAT_ID) From RECIPE_CAT", New OleDbConnection(strConnection))
        CmdCount.Connection.Open()
        lbCountCat.Text = "Total Category:&nbsp;" & CmdCount.ExecuteScalar()
        CmdCount.Connection.Close()

  End Sub
'++++++++++++++++++++++++++++++++++++++++++++
    
   


'++++++++++++++++++++++++++++++++++++++++++++
    'Update recipes, name, ingredients, instructions, author 
   Sub Change_Recipes(sender As Object, e As System.EventArgs)
    
        objConnection = New OledbConnection(strConnection)
        objConnection.Open()
    
      If request.form("CategoryName") = "" AND request.form("CategoryID") = "" Then
        strSQL = "update Recipes set Name='" & replace(request("Name"),"'","''")
        strSQL += "', Ingredients='" & replace(request("Ingredients"),"'","''")
        strSQL += "', Instructions='" & replace(request("Instructions"),"'","''")
        strSQL += "', Author='" & replace(request("Author"),"'","''")
        strSQL += "', Hits='" & replace(request("Hits"),"'","''")
        strSQL += "' where ID = " & request("id")

      Else
        
        strSQL = "update Recipes set Name='" & replace(request("Name"),"'","''")
        strSQL += "', Category='" & replace(request("CategoryName"),"'","''")
        strSQL += "', CAT_ID='" & replace(request("CategoryID"),"'","''")
        strSQL += "', Ingredients='" & replace(request("Ingredients"),"'","''")
        strSQL += "', Instructions='" & replace(request("Instructions"),"'","''")
        strSQL += "', Author='" & replace(request("Author"),"'","''")
        strSQL += "', Hits='" & replace(request("Hits"),"'","''")
        strSQL += "' where ID = " & request("id")

      End If

        objCommand = New OledbCommand(strSQL,objConnection)
        objCommand.ExecuteNonQuery()
    
        objCommand = nothing
        objConnection.Close()
        objConnection = nothing
        
        strURLRedirect = "confirmdel.aspx?catname=" & request("Name") & "&mode=update"
        Server.Transfer(strURLRedirect)
    
 End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
 'Display category name in the dropdownlist
 Sub GetDropdownCatName()

   Dim myConnection as New OledbConnection(strConnection)

   strSQL = "SELECT CAT_ID, CAT_TYPE From RECIPE_CAT Order by CAT_TYPE ASC"
                             
    Dim myCommand as New OledbCommand(strSQL, myConnection)

	myConnection.Open()
	
	Dim objDR as OledbDataReader
	objDR = myCommand.ExecuteReader(CommandBehavior.CloseConnection)
	
	'Databind the DataReader to the listbox Web control
	CategoryName.DataSource = objDR
	CategoryName.DataBind()
	
	'Add a new listitem to the beginning of the listitemcollection
        CategoryName.Items.Insert(0, new ListItem(""))

 End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
 Sub GetDropdownCatID()

   Dim myConnection as New OledbConnection(strConnection)

   strSQL = "SELECT CAT_ID, CAT_TYPE From RECIPE_CAT Order by CAT_TYPE ASC"
                             
    Dim myCommand as New OledbCommand(strSQL, myConnection)

	myConnection.Open()
	
	Dim objDR as OledbDataReader
	objDR = myCommand.ExecuteReader(CommandBehavior.CloseConnection)
	
	'Databind the DataReader to the listbox Web control
	CategoryID.DataSource = objDR
	CategoryID.DataBind()
	
	'Add a new listitem to the beginning of the listitemcollection
        CategoryID.Items.Insert(0, new ListItem(""))

 End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++    
    'Event Back to recipe manager page
    Sub BackToManager(sender as object, e as System.EventArgs)
    
        Server.Transfer("recipemanager.aspx")
    
 End Sub
'++++++++++++++++++++++++++++++++++++++++++++


'++++++++++++++++++++++++++++++++++++++++++++
  Private strSQL as string
  Private strURLRedirect as string
'++++++++++++++++++++++++++++++++++++++++++++

</script>

<!--#include file="inc_admindbconn.aspx"-->

<!--Powered By www.Ex-designz.net Recipe Cookbook ASP.NET version - Author: Dexter Zafra, Norwalk,CA-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Edit - Delete Page - www.ex-designz.net</title>
<style type="text/css" media="screen">@import "../css/cssreciaspx.css";</style>
</head>
<body>
<form style="margin-top: 16px; margin-bottom: 1px;" runat="server">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr>
    <td width="100%" colspan="2">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr>
    <td width="50%"><div style="padding-left: 20px;"><h3>Recipe Manager</h3></div>
<div style="padding-left: 20px;"><asp:Label font-name="verdana" font-size="9" ID="lblusername" runat="server" /></div>
</td>
  </tr>
</table>
<br />
</td>
  </tr>
  <tr>
    <td width="21%" align="left" valign="top">
<!--Begin Admin Task Panel-->
<div style="margin-right: 14px;">
 <div class="roundcont2">
<div class="roundtop">
<img src="../images/hleft.gif" height="5" width="5" alt="" class="corner">
<div class="dcnt"><span class="content3">Admin Task</span></div> 
</div>
</div>
<div class="contentdisplay">
<div class="contentdis5">
<div class="divmenu2">
<asp:Label ID="lbCountRecipe" runat="server" />
<br />
<br />
<span class="bluearrow2">»</span>&nbsp;<asp:HyperLink runat="server" ID="lblmangermainpagelink" tooltip="Back to recipe manager home" NavigateUrl="recipemanager.aspx">Recipe Manager Home</asp:HyperLink>
<br />
<br />
    <asp:Image ID="img1" ImageAlign="AbsBottom" ImageURL="../images/adminapproval_icon.gif" AlternateTex="Aprroval Manager" runat="server" />
    <asp:HyperLink tooltip="Recipes waiting for approval" runat="server" ID="approvallink" NavigateUrl="recipemanager.aspx?tab=1"><asp:Label ID="lblunapproved" runat="server" /></asp:HyperLink>
<br />
<asp:Image ID="img2" ImageAlign="AbsBottom" ImageURL="../images/commentadmin_icon.gif" AlternateTex="Comment Manager" runat="server" />
<asp:HyperLink tooltip="Click this link to edit/delete recipe comments" runat="server" ID="countcommentlink" NavigateUrl="commentsmanager.aspx"><asp:Label ID="lbCountComments" runat="server" /></asp:HyperLink>
<br />
<asp:Image ID="img3" ImageAlign="AbsBottom" ImageURL="../images/admincategory_icon.gif" AlternateTex="Category Manager" runat="server" />
<asp:HyperLink tooltip="Click this link to edit/delete and add a recipe category" runat="server" ID="editcat" NavigateUrl="categorymanager.aspx"><asp:Label ID="lbCountCat" runat="server" /></asp:HyperLink>  
</div>
</div>
</div>
</div>
<!--End Admin Task Panel-->
</td>
    <td width="79%" valign="top">
      <table width="70%" border="0" cellpadding="0" cellspacing="1" align="left">
          <tr>
              <td colspan="2"  bgcolor="#6898d0">
               <div class="roundcont2">
                <div class="roundtop">
                <img src="../images/hleft.gif" height="5" width="5" alt="" class="corner">
               <div style="text-align: left; padding-left: 6px;padding-bottom: 2px;"><span class="content3">Updating Recipe ID #: <%=request("id")%></span></div> 
             </div>
        </div>
    </td>
</tr>
      <tr>
          <td bgcolor="#F4F9FF" class="content2">Name:</td>   		
             <td bgcolor="#FBFDFF">
              <asp:TextBox runat="server" id="Name" class="textbox" size="30" maxlenght="30" />
       </td>
   </tr>
<tr>
    <td bgcolor="#F4F9FF" class="content2">Category:</td>   		
        <td bgcolor="#FBFDFF">
<span class="content8"><strong>Note:</strong> If you want to move the recipe to a different category, make sure you match the left field (Category Name) to the right field (Category Name), i.e. Barbque has to match Barbeque. If you don't want to move it, don't do nothing, leave it blank.</span>
<br />
   <asp:listbox id="CategoryName" runat="server" Rows="1" DataTextField="CAT_TYPE" DataValueField="CAT_TYPE" /> 
   <asp:listbox id="CategoryID" runat="server" Rows="1" DataTextField="CAT_TYPE" DataValueField="CAT_ID" /> 
  </td>
</tr>
      <tr>
           <td bgcolor="#F4F9FF" class="content2">Author:</td>   		
                <td bgcolor="#FBFDFF">
                 <asp:TextBox runat="server" id="Author" class="textbox" size="25" maxlenght="25" />
        </td>
    </tr>
<tr>
     <td bgcolor="#F4F9FF" class="content2">Hits:</td>   		
         <td bgcolor="#FBFDFF">
            <asp:TextBox runat="server" id="Hits" class="textbox" size="6" maxlenght="6" />
       </td>
</tr>
     <tr>
         <td valign="top" bgcolor="#F4F9FF" class="content2">Ingredients:</td>
            	 <td bgcolor="#FBFDFF">
                <asp:TextBox runat="server" id="Ingredients" Class="textbox" textmode="multiline" columns="70" rows="14" />
       </td>
 </tr>
         <tr>
               <td valign="top" bgcolor="#F4F9FF" class="content2">Instructions:</td>  		
                 <td bgcolor="#FBFDFF">
                    <asp:TextBox runat="server" id="Instructions" Class="textbox" textmode="multiline" columns="70" rows="14" />
         </td>
    </tr>
<tr>
       <td align=left colspan=2 bgcolor="#ffffff">
         <div style="padding-left: 75px;">
            <asp:Button runat="server" Text="Update" id="updatebutton" class="submit" tooltip="Click to update" onclick="Change_Recipes"/>
            <asp:Button runat="server" Text="Cancel" id="cancelbutton" class="submit" tooltip="Click to cancel" onclick="BackToManager"/>
       </div>
     </td>
   </tr>
   </table>
  </td>
</tr>
</table>
</form>
<div style="text-align: center; margin-top: 30px; margin-bottom: 20px;">
<a href="http://www.ex-designz.net" class="hlink" title="Visit our website">Powered By Ex-designz.net World Recipe</a>
</div>
<br />
</body>
</html>