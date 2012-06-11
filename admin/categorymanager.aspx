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
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.Oledb" %>

<script runat="server">

'++++++++++++++++++++++++++++++++++++++++++++
  'Handle page load events
  Sub Page_Load(Sender As Object, E As EventArgs)

          'Call recipe count
          DisplayCommentsCount()

          'Call recipe count
          DisplayRecipeCount()

          'Call count unpprove recipes
          UnApproveRecipe()

          'Call check user function - Check if user has started a session 
          Check_User()

          'Display admin user name
          lblusername.Text = "Welcome Admin:&nbsp;" & session("userid")
          catmanagerlink.enabled = False
    
          Panel1.visible = False

       If Not Page.IsPostBack then

            GetRecipes("CAT_ID ASC")

      End If

 End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
  'Bind data - show data
  Sub GetRecipes(strSQL as string)

        'Creates the SQL statement
         strSQL = "SELECT * FROM RECIPE_CAT Order By CAT_ID ASC"
         objConnection = New OledbConnection(strConnection)
         objCommand = New OledbCommand(strSQL, objConnection)

         Dim RecipeAdapter as New OledbDataAdapter(objCommand)
         Dim dts as New DataSet()
         RecipeAdapter.Fill(dts)

         'Display the total number of categories in the left panel menu
         lbCountCat.Text = "Total Category:&nbsp;" & CStr(dts.Tables(0).Rows.Count)

         Recipes_table.DataSource = dts  
         Recipes_table.DataBind()

         'close database connection
         objConnection.Close()

 End Sub 
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
  'Handles update category event
   Sub Update_Category(sender As Object, e As System.EventArgs)

       If Page.IsPostBack Then
    
        objConnection = New OledbConnection(strConnection)
        objConnection.Open()
        
        strSQL = "update RECIPE_CAT set CAT_TYPE='" & replace(request("CategoryName"),"'","''")
        strSQL += "' where CAT_ID = " & request("CategoryID")

        objCommand = New OledbCommand(strSQL,objConnection)
        objCommand.ExecuteNonQuery()
    
        objCommand = nothing
        objConnection.Close()
        objConnection = nothing

        'Update Recipes table Category name field
        Dim strSQL3 as string
        objConnection = New OledbConnection(strConnection)
        objConnection.Open()
    
        strSQL3 = "update Recipes set Category='" & replace(request("CategoryName"),"'","''")
        strSQL3 += "' where CAT_ID = " & request("CategoryID")

        objCommand = New OledbCommand(strSQL3,objConnection)
        objCommand.ExecuteNonQuery()

        objCommand = nothing
        objConnection.Close()
        objConnection = nothing

        strURLRedirect = "confirmcatedit.aspx?catname=" & request("CategoryName") & "&mode=update"
        Server.Transfer(strURLRedirect)
        
       End If
    
 End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
  'Handles add new category event
   Sub Add_Category(sender As Object, e As System.EventArgs)
    
        If Page.IsPostBack Then
    
        objConnection = New OledbConnection(strConnection)
        objConnection.Open()
    
        strSQL = "insert into RECIPE_CAT (CAT_TYPE) values ('" & replace(request("CategoryName"),"'","''") & "')"
    
        objCommand = New OledbCommand(strSQL,objConnection)
        objCommand.ExecuteNonQuery()
    
        objCommand = nothing
        objConnection.Close()
        objConnection = nothing

        strURLRedirect = "confirmcatedit.aspx?catname=" & request("CategoryName") & "&mode=add"
        Server.Transfer(strURLRedirect)
        
    End If
    
 End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
   'Handle delete the selected category
   Sub Delete_Category(sender As Object, e As System.EventArgs)
    
      If Page.IsPostBack Then
    
        objConnection = New OledbConnection(strConnection)
        objConnection.Open()
    
        strSQL = "delete * from RECIPE_CAT where CAT_ID = " & request("CategoryID")
        objCommand = New OledbCommand(strSQL,objConnection)
        objCommand.ExecuteNonQuery()

        objCommand = nothing
        objConnection.Close()
        objConnection = nothing


        'Delete all related recipes in the Recipes Table
        Dim strSQL2 as string
        objConnection = New OledbConnection(strConnection)
        objConnection.Open()
    
        strSQL2 = "delete * from Recipes where CAT_ID = " & request("CategoryID")
        objCommand = New OledbCommand(strSQL2,objConnection)
        objCommand.ExecuteNonQuery()

        objCommand = nothing
        objConnection.Close()
        objConnection = nothing

        strURLRedirect = "confirmcatedit.aspx?catname=" & request("CategoryName") & "&mode=del"
        Server.Transfer(strURLRedirect)
        
     End If
    
  End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
  'Display total number of unapprove recipes
  Sub UnApproveRecipe()

Dim CmdCount As New OleDbCommand("Select Count(ID) From Recipes Where LINK_APPROVED = 0", New OleDbConnection(strConnection))
        CmdCount.Connection.Open()
        lblunapproved.Text = "Waiting For Approval:&nbsp;" & CmdCount.ExecuteScalar() 
        CmdCount.Connection.Close()

  End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
  'Display total number of comments
  Sub DisplayCommentsCount()

        Dim CmdCount As New OleDbCommand("Select Count(ID) From COMMENTS_RECIPE", New OleDbConnection(strConnection))
        CmdCount.Connection.Open()
        lbCountComments.Text = "Total Comments:&nbsp;" & CmdCount.ExecuteScalar()
        CmdCount.Connection.Close()

  End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
   'Display total number of recipes
   Sub DisplayRecipeCount()

        Dim CmdCount As New OleDbCommand("Select Count(ID) From Recipes", New OleDbConnection(strConnection))
        CmdCount.Connection.Open()
        lbCountRecipe.Text = "Total Recipes:&nbsp;" & CmdCount.ExecuteScalar()
        CmdCount.Connection.Close()

     End Sub
'++++++++++++++++++++++++++++++++++++++++++++


'++++++++++++++++++++++++++++++++++++++++++++
  Sub Edit_Handle(sender as Object, e As DataGridCommandEventArgs)

       If (e.CommandName="edit") then

            Dim iIdNumber as TableCell = e.Item.Cells(0)
            Dim iCatName as TableCell = e.Item.Cells(1)
            Dim address as string

            Panel1.visible = True
            Panel3.visible = false
            AddNewCat.visible = False
            lblheaderform.text = "Editing Category #:&nbsp;" & iIdNumber.text
            lblnamedis2.text = "Category Name:"
            updatebutton.visible = true
            CategoryID.visible = true

            'This will be the value to be populated into the textboxes
            CategoryName.text = iCatName.Text
            CategoryID.value = iIdNumber.text

             e.item.BackColor = System.Drawing.ColorTranslator.FromHtml("#F0E68C")

        ElseIf (e.CommandName="delete") then

            Dim iIdNumber as TableCell = e.Item.Cells(0)
            Dim iCatName as TableCell = e.Item.Cells(1)
            Dim address as string

            Panel1.visible = True
            Panel3.visible = True
            AddNewCat.visible = False
            lblheaderform.text = "Deleting Category #:&nbsp;" & iIdNumber.text
            lblnamedis2.text = "Category Name:"
            updatebutton.visible = false
            CategoryID.visible = true

            'This will be the value to be populated into the textboxes
            CategoryName.text = iCatName.Text
            CategoryID.value = iIdNumber.text

            e.item.BackColor = System.Drawing.ColorTranslator.FromHtml("#F0E68C")

          strSQL = "SELECT Count(CAT_ID) FROM Recipes WHERE CAT_ID = " & iIdNumber.text

         'Open database - connect to the database      
         objConnection = New OledbConnection(strConnection)
         objCommand = New OledbCommand(strSQL, objConnection)

         objCommand.Connection.Open()
         lblrcdcount2.Text = "The number of recipes belong to&nbsp;" & iCatName.Text & "&nbsp;category:&nbsp; " & objCommand.ExecuteScalar() 
         objCommand.Connection.Close()

     End if

  End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
 Sub dgCat_ItemDataBound(sender as Object, e as DataGridItemEventArgs)

    'First, make sure we're not dealing with a Header or Footer row
    If e.Item.ItemType <> ListItemType.Header AND e.Item.ItemType <> ListItemType.Footer then

       'Data row mouseover changecolor
       e.Item.Attributes.Add("onmouseover", "this.style.backgroundColor='#F4F9FF'")
       e.Item.Attributes.Add("onmouseout", "this.style.backgroundColor='#ffffff'")

       'Display cell tooltip in the grid
       e.Item.Cells(0).ToolTip = "Category # " & DataBinder.Eval(e.Item.DataItem, "CAT_ID")
       e.Item.Cells(1).ToolTip = DataBinder.Eval(e.Item.DataItem, "CAT_TYPE") & " Category"

      Dim editButton as LinkButton = e.Item.Cells(2).Controls(0)
      Dim deleteButton as LinkButton = e.Item.Cells(3).Controls(0)

     deleteButton.ToolTip = "Delete category (" & DataBinder.Eval(e.Item.DataItem, "CAT_TYPE") & ") CAT ID #:" & DataBinder.Eval(e.Item.DataItem, "CAT_ID")  
     editButton.ToolTip = "Edit category (" & DataBinder.Eval(e.Item.DataItem, "CAT_TYPE") & ") CAT ID #:" & DataBinder.Eval(e.Item.DataItem, "CAT_ID")  
   End If

      'Display the pagecount in the top and footer
      Dim pageindex As Integer = Recipes_table.CurrentPageIndex + 1
      LblPageInfo.Text = "Showing Page " + pageindex.ToString() + " of " + Recipes_table.PageCount.ToString()

 End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
 'Switch to Add Category mode
  Sub ChangeToAddCat(ByVal s As Object, ByVal e As EventArgs)
        
        CategoryName.text = ""
        CategoryID.visible = false
        Panel1.visible = True
        Panel3.visible = false
        AddNewCat.visible = True
        updatebutton.visible = False
        lblheaderform.text = "Adding New Category"
        lblnamedis2.text = "Category Name:"

 End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
  'Handles page change links - paging system
  Sub New_Page(sender As Object, e As DataGridPageChangedEventArgs)

         Recipes_table.CurrentPageIndex = e.NewPageIndex
         GetRecipes("CAT_ID ASC")

  End Sub
'++++++++++++++++++++++++++++++++++++++++++++


'+++++++++++++++++++++++++++++++++++++++++++++++++
'Here we declare our module-level variables 

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
<title>Recipe Manager - www.ex-designz.net</title>
<style type="text/css" media="screen">@import "../css/cssreciaspx.css";</style>
</head>
<body>
<form style="margin-top: 16px; margin-bottom: 1px;" runat="server">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr>
    <td width="100%" colspan="2">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr>
    <td width="50%"><div style="padding-left: 20px;"><h3>Recipe Category Manager</h3></div>
<div style="padding-left: 20px;"><asp:Label font-name="verdana" font-size="9" ID="lblusername" runat="server" /></div>
<br />
</td>
  </tr>
</table>
</td>
  </tr> 
  <tr>
    <td width="21%" align="left" valign="top">
<!--Begin Admin Task Panel-->
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
<span class="bluearrow2">»</span>&nbsp;<a title="Back to recipe manager main page" href="recipemanager.aspx">Recipe Manager Main</a>
<br />
<br />
       <asp:Image ID="img1" ImageAlign="AbsBottom" ImageURL="../images/adminapproval_icon.gif" AlternateTex="Aprroval Manager" runat="server" />
       <asp:HyperLink tooltip="Recipes waiting for approval" runat="server" ID="approvallink" NavigateUrl="recipemanager.aspx?tab=1"><asp:Label ID="lblunapproved" runat="server" /></asp:HyperLink>
<br />
<asp:Image ID="img2" ImageAlign="AbsBottom" ImageURL="../images/commentadmin_icon.gif" AlternateTex="Comment Manager" runat="server" />
<asp:HyperLink tooltip="Click this link to edit/delete recipe comments" runat="server" ID="countcommentlink" NavigateUrl="commentsmanager.aspx"><asp:Label ID="lbCountComments" runat="server" /></asp:HyperLink>
<br />
<asp:Image ID="img3" ImageAlign="AbsBottom" ImageURL="../images/admincategory_icon.gif" AlternateTex="Category Manager" runat="server" />
<asp:HyperLink tooltip="Click this link to edit/delete and add a recipe category" runat="server" ID="catmanagerlink" NavigateUrl="categorymanager.aspx"><asp:Label ID="lbCountCat" runat="server" /></asp:HyperLink> 
<br />
<asp:Image ID="img4" ImageAlign="AbsBottom" ImageURL="../images/addnewcategoryimg.gif" AlternateTex="Click this link to Add a New Category" runat="server" />
<a title="Click This link to add new category" href="categorymanager.aspx#this"
  ID="AddNewCategory"
  onserverclick="ChangeToAddCat"
  runat="server">Add New Category</a> 
</div>
</div>
</div>
<br />
<!--End Admin Task Panel-->
</td>
   <td width="79%" valign="top">
  <!--Begin update edit form-->
<asp:Panel ID="Panel1" runat="server">
<div style="padding-left: 21px; width:47%;">
<div class="roundcont2">
<div class="roundtop">
<img src="../images/hleft.gif" height="5" width="5" alt="" class="corner">
<div style="text-align: left; padding-left: 6px;padding-bottom: 2px;"><asp:Label ID="lblheaderform" cssClass="content3" runat="server" /></div> 
</div>
</div>
<div class="contentdisplay3">
<div class="contentdis5">
<asp:Label ID="lblnamedis2" cssClass="content2" runat="server" />
<asp:TextBox runat="server" id="CategoryName" class="textbox" size="18" maxlenght="18" />
        <asp:RequiredFieldValidator runat="server"
        id="authorname" ControlToValidate="CategoryName"
        cssClass="cred2" errormessage="<br />* Enter a Category Name"
        display="Dynamic" />
<asp:Button runat="server" Text="Update" id="updatebutton" tooltip="Click to update" class="submit" OnClick="Update_Category" />
<asp:Button runat="server" Text="Add" tooltip="Click to add new category" id="AddNewCat" class="submit" onclick="Add_Category" />
<input type="text"  runat="server" id="CategoryID" name="CategoryID" class="textbox" size="2" maxlenght="2" readOnly="True" style="visibility:hidden;">
<br />
<asp:Panel ID="Panel3" runat="server">
<div style="padding-top: 4px; padding-bottom: 6px;">
<span class="content2">
<asp:Label ID="lblrcdcount2" runat="server" />
<br />
<span class="cred">Are you sure you want to delete this category?
<br />
Note: all recipes belong to this category will be deleted as well.
</span>
<br />
<a title="Click This link to permanently delete this category" class="cred" href="categorymanager.aspx#this"
  ID="DelCategory"
  onserverclick="Delete_Category"
  runat="server">Finalize Delete Category</a>
</span>
</div>
</asp:Panel>
</div>
</div>
<br />
</asp:Panel>
<!--End update edit form-->
</div>
<!--End display edit category name form-->
<table width="100%" border="0" cellspacing="1" align="left">
  <tr>
    <th scope="row"><div align="left">
     <asp:DataGrid runat="server" id="Recipes_table" cssclass="hlink" AutoGenerateColumns="False" 
     Backcolor="#ffffff" BorderStyle="none" BorderColor="#E1EDFF" cellpadding="5" Width="95%" HorizontalAlign="Center" PageSize="10" OnItemDataBound="dgCat_ItemDataBound" AllowPaging="True" OnPageIndexChanged="New_Page" onItemCommand="Edit_Handle"> 
  <HeaderStyle Font-Bold="True" BackColor="#6898d0" ForeColor="#ffffff" cssclass="header" />
  <AlternatingItemStyle BackColor="White" />                                   
  <Columns>
  <asp:BoundColumn DataField="CAT_ID" HeaderText="CAT_ID" SortExpression="CAT_ID ASC" />
  <asp:BoundColumn DataField="CAT_TYPE" HeaderText="Category Name" SortExpression="CAT_TYPE ASC" />
  <asp:ButtonColumn Text='<img border="0" src="../images/icon_edit.gif">' HeaderText="Edit" CommandName="edit" />
  <asp:ButtonColumn Text='<img border="0" src="../images/icon_delete.gif">' HeaderText="Delete" CommandName="delete" />
  </Columns>
   <PagerStyle Mode="NumericPages" BackColor="#fcfcfc" HorizontalAlign="left" />
    </asp:DataGrid>  
<div style="padding-left:20px; padding-top:3px;"><asp:Label ID="LblPageInfo" cssClass="content2" runat="server" />
</div>                                                                                           
   </div></th>
 </tr>
</table>
</td>
  </tr>
</table>
</form>
<div style="text-align: center; margin-top: 15px;">
<a href="http://www.ex-designz.net" class="hlink" title="Visit our website">Powered By Ex-designz.net World Recipe</a>
</div>
</body>
</html>