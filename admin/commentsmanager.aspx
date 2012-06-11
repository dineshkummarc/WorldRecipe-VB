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
          DisplayRecipeCount()

          'Call count unpprove recipes
          UnApproveRecipe()

          DisplayCategoryCount()

          'Call check user function - Check if user has started a session 
          Check_User()

          Panel1.visible = False
          countcommentlink.enabled = false

          'Display admin user name
          lblusername.Text = "Welcome Admin:&nbsp;" & session("userid")

       If Not Page.IsPostBack then

           GetRecipes("DATE DESC")

       End If

 End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
  'Display number of categories
  Sub DisplayCategoryCount()

        Dim CmdCount As New OleDbCommand("Select Count(CAT_ID) From RECIPE_CAT", New OleDbConnection(strConnection))
        CmdCount.Connection.Open()
        lbCountCat.Text = "Total Category:&nbsp;" & CmdCount.ExecuteScalar()
        CmdCount.Connection.Close()

   End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
  'Display number of unapprove recipes
  Sub UnApproveRecipe()

Dim CmdCount As New OleDbCommand("Select Count(ID) From Recipes Where LINK_APPROVED = 0", New OleDbConnection(strConnection))
        CmdCount.Connection.Open()
        lblunapproved.Text = "Waiting For Approval:&nbsp;" & CmdCount.ExecuteScalar() 
        CmdCount.Connection.Close()

  End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
   'Display number of recipes
   Sub DisplayRecipeCount()

        Dim CmdCount As New OleDbCommand("Select Count(ID) From Recipes", New OleDbConnection(strConnection))
        CmdCount.Connection.Open()
        lbCountRecipe.Text = "Total Recipes:&nbsp;" & CmdCount.ExecuteScalar()
        CmdCount.Connection.Close()

     End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
  'Display recipe to the datagrid
  Sub GetRecipes(strSQL as string)

        'Creates the SQL statement
         strSQL = "SELECT * FROM COMMENTS_RECIPE Order By COM_ID DESC"
         objConnection = New OledbConnection(strConnection)
         objCommand = New OledbCommand(strSQL, objConnection)

         Dim RecipeAdapter as New OledbDataAdapter(objCommand)
         Dim dts as New DataSet()
         RecipeAdapter.Fill(dts)

         'Display the total number of comments in the left panel menu
         lbCountComments.Text = "Total Comments:&nbsp;" & CStr(dts.Tables(0).Rows.Count)

         Recipes_table.DataSource = dts  
         Recipes_table.DataBind()

        objConnection.Close()

 End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
    'Handle update comment
   Sub Update_Comments(sender As Object, e As System.EventArgs)
   
      If Page.IsPostBack Then

        objConnection = New OledbConnection(strConnection)
        objConnection.Open()
        
        strSQL = "update COMMENTS_RECIPE set AUTHOR='" & replace(request("Author"),"'","''")
        strSQL += "', EMAIL='" & replace(request("Email"),"'","''")
        strSQL += "', COMMENTS='" & replace(request("Comments"),"'","''")
        strSQL += "' where COM_ID = " & request("KeyIDs")

        objCommand = New OledbCommand(strSQL,objConnection)
        objCommand.ExecuteNonQuery()
    
        objCommand = nothing
        objConnection.Close()
        objConnection = nothing
        
        'Redirect to confirm update page
        strURLRedirect = "confirmcommentupdate.aspx?mode=update"
        Server.Transfer(strURLRedirect)

    End If
    
 End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
 'Confirm delete comment - show popup dialog box
 Sub dgComment_ItemDataBound(sender as Object, e as DataGridItemEventArgs)

    'First, make sure we're not dealing with a Header or Footer row
    If e.Item.ItemType <> ListItemType.Header AND e.Item.ItemType <> ListItemType.Footer then

      Dim editButton as LinkButton = e.Item.Cells(0).Controls(0)
      Dim deleteButton as LinkButton = e.Item.Cells(1).Controls(0)

      'We can now add the onclick event handler
      deleteButton.Attributes("onclick") = "javascript:return confirm('Are you sure you want to delete Comment ID # " & _
       DataBinder.Eval(e.Item.DataItem, "COM_ID") & "?')"  

      editButton.ToolTip = "Update comment ID #: " & DataBinder.Eval(e.Item.DataItem, "COM_ID") & " author: " & DataBinder.Eval(e.Item.DataItem, "Author")
      deleteButton.ToolTip = "Delete comment ID #: " & DataBinder.Eval(e.Item.DataItem, "COM_ID") & " author: " & DataBinder.Eval(e.Item.DataItem, "Author")

      'Data row mouseover changecolor
      e.Item.Attributes.Add("onmouseover", "this.style.backgroundColor='#F4F9FF'")
      e.Item.Attributes.Add("onmouseout", "this.style.backgroundColor='#ffffff'")
  
    End If

 End Sub
'++++++++++++++++++++++++++++++++++++++++++++  



'++++++++++++++++++++++++++++++++++++++++++++
  'Delete the selected comment
  Sub Delete_Comment(sender as Object, e As DataGridCommandEventArgs)


        If (e.CommandName="Delete") then
        Dim iIdKeyNumber as TableCell = e.Item.Cells(2) 
        Dim iIdCommNumber as TableCell = e.Item.Cells(3)
    
        objConnection = New OledbConnection(strConnection)
        objConnection.Open()
    
        strSQL = "delete * from COMMENTS_RECIPE where COM_ID = " & iIdKeyNumber.text
        objCommand = New OledbCommand(strSQL,objConnection)
        objCommand.ExecuteNonQuery()

        objCommand = nothing
        objConnection.Close()
        objConnection = nothing
    
       'This part will decriment 1 to the total comments field in the Recipes table   
       Dim strSQL2 as string
       objConnection = New OledbConnection(strConnection)
       objConnection.Open()
        
       strSQL2 = "Update Recipes set TOTAL_COMMENTS = TOTAL_COMMENTS - 1 where ID=" & iIdCommNumber.text

        objCommand = New OledbCommand(strSQL2,objConnection)
        objCommand.ExecuteNonQuery()

        objCommand = nothing
        objConnection.Close()
        objConnection = nothing

        'Redirect to confirm delete page
        strURLRedirect = "confirmcommentupdate.aspx?mode=del"
        Server.Transfer(strURLRedirect)

   End If

 End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
'Handle edit command events
  Sub Edit_Handle(sender as Object, e As DataGridCommandEventArgs)

        If (e.CommandName="edit") then

            Dim iComIDNum as TableCell = e.Item.Cells(2)      
            Dim iAuthorName as TableCell = e.Item.Cells(4)
            Dim iAuthorEmail as TableCell = e.Item.Cells(5)
            Dim iComment as TableCell = e.Item.Cells(7)

           Panel1.visible = True

           'This will be the value to be populated into the textboxes
            Author.text = iAuthorName.Text
            Email.text = iAuthorEmail.Text
            Comments.text = iComment.text
            KeyIDs.value = iComIDNum.text
            lblheaderform.text = "Updating Comment #:&nbsp;" & iComIDNum.text 

        End if

  End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
  'Handles page change links - paging system
  Sub New_Page(sender As Object, e As DataGridPageChangedEventArgs)

         Recipes_table.CurrentPageIndex = e.NewPageIndex
         GetRecipes("DATE DESC")

  End Sub
'++++++++++++++++++++++++++++++++++++++++++++


'++++++++++++++++++++++++++++++++++++++++++++
 'Module-level variables

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
    <td width="50%"><div style="padding-left: 20px;"><h3>Recipe Comments Manager</h3></div>
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
<asp:HyperLink tooltip="Click this link to edit/delete and add a recipe category" runat="server" ID="editcat" NavigateUrl="categorymanager.aspx"><asp:Label ID="lbCountCat" runat="server" /></asp:HyperLink>  
</div>
</div>
</div>
<br />
<!--End Admin Task Panel-->
</td>
    <td width="79%" valign="top">
<!--Begin update edit form-->
<asp:Panel ID="Panel1" runat="server">
<div style="padding-left: 21px; width:46%;">
<div class="roundcont2">
<div class="roundtop">
<img src="../images/hleft.gif" height="5" width="5" alt="" class="corner">
<div style="text-align: left; padding-left: 6px;padding-bottom: 2px;"><asp:Label ID="lblheaderform" cssClass="content3" runat="server" /></div> 
</div>
</div>
<div class="contentdisplay3">
<div class="contentdis5">
<span class="content2">Author:</span>
<asp:TextBox runat="server" id="Author" class="textbox" size="20" maxlenght="20" />
<input type="text" runat="server" id="KeyIDs" name="KeyIDs" class="textbox" size="3" maxlenght="3" readOnly="True" style="visibility:hidden;">     
<br />
<span class="content2">Email:</span>
&nbsp;&nbsp;<asp:TextBox runat="server" id="Email" class="textbox" size="30" maxlenght="30" />
<br />
<span class="content2">Comment:</span>
<br />
<asp:TextBox runat="server" id="Comments" Class="textbox" textmode="multiline" columns="46" rows="5" />
<br />
<asp:Button runat="server" Text="Update" id="updatebutton" class="submit" onclick="Update_Comments" />
</div>
</div>
</div>
<br />
</asp:Panel>
<!--End update edit form-->
<table width="100%" border="0" cellspacing="1">
  <tr>
    <th scope="row"><div align="left">
     <asp:DataGrid runat="server" id="Recipes_table" cssclass="hlink" AutoGenerateColumns="False" 
     Backcolor="#ffffff" BorderStyle="none" BorderColor="#E1EDFF" cellpadding="5" Width="95%" HorizontalAlign="Center" PageSize="30" AllowPaging="True" OnPageIndexChanged="New_Page" OnItemDataBound="dgComment_ItemDataBound" DataKeyField="ID" OnDeleteCommand="Delete_Comment" onItemCommand="Edit_Handle"> 
     <HeaderStyle Font-Bold="True" BackColor="#6898d0" ForeColor="#ffffff" cssclass="header" />
     <AlternatingItemStyle BackColor="White" />                                   
     <Columns>
     <asp:ButtonColumn Text='<img border="0" src="../images/icon_edit.gif">' HeaderText="Edit" CommandName="edit" />
     <asp:ButtonColumn Text='<img border="0" src="../images/icon_delete.gif">' HeaderText="Delete" CommandName="Delete" />
     <asp:BoundColumn DataField="COM_ID" HeaderText="Key" SortExpression="COM_ID DESC" />  
     <asp:BoundColumn DataField="ID" HeaderText="ID" SortExpression="id ASC" />  
     <asp:BoundColumn DataField="Author" HeaderText="Author" SortExpression="Author ASC" />  
<asp:BoundColumn DataField="EMAIL" HeaderText="Email" SortExpression="EMAIL ASC" /> 
 <asp:BoundColumn DataField="Date" DataFormatString="{0:d}" HeaderText="Date" SortExpression="Date" />
<asp:BoundColumn DataField="COMMENTS" HeaderText="Comments" SortExpression="COMMENTS ASC" /> 
<asp:HyperLinkColumn HeaderText="Details" Text="View" DataNavigateUrlField="ID" DataNavigateUrlFormatString="../recipedetail.aspx?id={0}" target="_blank"/> 
     </Columns>
     <PagerStyle Mode="NumericPages" BackColor="#fcfcfc" HorizontalAlign="left" />
    </asp:DataGrid>                                                                                             
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