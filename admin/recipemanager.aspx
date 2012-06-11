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
 'Handles page load events
  Sub Page_Load(Sender As Object, E As EventArgs)

          'Call the total number of recipes count
          DisplayRecipeCount()

          'Call count unpprove recipes
          UnApproveRecipe()

          DisplayCategoryCount()

          'Call count total comments
          DisplayCommentsCount()

          GetDropdownlistCategory()

          Display_Letter()

          Display_LetterCat()

          'Call check user function - Check if user has started a session 
          Check_User()

          'Display admin user name
          lblusername.Text = "Welcome Admin:&nbsp;" & session("userid")
          lblunapproved2.visible = false
          lblmangermainpage.text = "Default View"
          lblmangermainpagelink.ToolTip = "Back to Default View"

       If Not Page.IsPostBack then

          lblrecordperpage.text = "Default 20 records per page"
          lblrecordperpageFooter.text = "Showing default 20 records per page"
          lblrecordperpageTop.text = "- 20 records per page" 

          dgrd_recipe.PageSize = 20
          SortExpression = "ID"  'Set the default column to sort by

          BindData()

       End If

 End Sub
'++++++++++++++++++++++++++++++++++++++++++++




'++++++++++++++++++++++++++++++++++++++++++++
  'Display the alphabetical letter listing 
  Sub Display_Letter()

      Dim i as Integer
      lblletterlegend.text = "Recipe A-Z:"
      lblalphaletter.Text = string.empty
      for i = 65 to 90	
         lblalphaletter.text = lblalphaletter.text & "<a href=""recipemanager.aspx?l=" & chr(i) & chr(34) & _
" class=""dlet"" title="& chr(i) & ">" & chr(i) &  "</a>&nbsp;&nbsp;"
      next

 End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
  'Display the alphabetical letter listing for category sorting
  Sub Display_LetterCat()

      Dim i as Integer
      Dim cid as integer = request("catid")
      lblletterlegend.text = "Recipe A-Z:"
      lblalphalettercat.Text = string.empty
      for i = 65 to 90	
lblalphalettercat.text = lblalphalettercat.text & "<a href=""recipemanager.aspx?catid=" & cid & "&lc=" & chr(i) & chr(34) & " class=""dlet"" title="& chr(i) & ">" & chr(i) &  "</a>&nbsp;&nbsp;"
      next

 End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'+++++++++++++++++++++++++++++++++++++++++++++
  'Show data in the datagrid
  Sub BindData()
     
         BindData_SQL_Statements()

         objConnection = New OledbConnection(strConnection)
         objCommand = New OledbCommand(strSQL, objConnection)

         Dim RecipeAdapter as New OledbDataAdapter(objCommand)
         Dim dts as New DataSet()
         RecipeAdapter.Fill(dts)

         'Display total number of categories sorted
         If request("catid") <> "" Then
           lblrcdCatcount.Text = CStr(dts.Tables(0).Rows.Count) & "&nbsp;recipes"
           lblrcdCatcountfooter.Text = CStr(dts.Tables(0).Rows.Count) & "&nbsp;records"
         End If

         'Display total number of recipes in sorted categories starting with letter
         If request("lc") <> "" Then
     lblrcdCatcount.Text = CStr(dts.Tables(0).Rows.Count) & "&nbsp;recipes starting with letter&nbsp;<b>" & request("lc") & "</b>"
     lblrcdCatcountfooter.Text = CStr(dts.Tables(0).Rows.Count) & "&nbsp;recipes starting with letter&nbsp;" & request("lc")
         End If

         'Display total number of recipes sorted by letter
         If request("l") <> "" Then
 lblSortedCat.Text = "Showing&nbsp;" & CStr(dts.Tables(0).Rows.Count) & "&nbsp;recipes starting with letter&nbsp;<b>" & request("l") & "</b>"
           lblrcdalphaletterfooter.Text = CStr(dts.Tables(0).Rows.Count) & "&nbsp;recipes starting with letter&nbsp;<b>" & request("l") & "</b>&nbsp;-" 
           lbCountRecipeFooter.visible = false
         End If

         If request("find") <> "" AND request("s") = "1" Then

 lblSortedCat.Text = "Your search for recipe name&nbsp;" & "(<b>" & request("find") & "</b>) return:&nbsp;" & CStr(dts.Tables(0).Rows.Count) & "&nbsp;records"
           lbCountRecipeFooter.visible = false
         
         elseif request("find") <> "" AND request("s") = "2" Then

 lblSortedCat.Text = "Your search for recipe author&nbsp;" & "(<b>" & request("find") & "</b>) return:&nbsp;" & CStr(dts.Tables(0).Rows.Count) & "&nbsp;records"
           lbCountRecipeFooter.visible = false

         elseif request("find") <> "" AND request("s") = "3" Then

 lblSortedCat.Text = "Your search for recipe ID#&nbsp;" & "(<b>" & request("find") & "</b>) return:&nbsp;" & CStr(dts.Tables(0).Rows.Count) & "&nbsp;records"
           lbCountRecipeFooter.visible = false

         End If

         'Update the column headers
         UpdateColumnHeaders(dgrd_recipe)
         dgrd_recipe.DataSource = dts

         'Finally, let's bind the data
         dgrd_recipe.DataBind() 

         EnabledDisabled_PagerButtons()

         'Let's close database connection
         objConnection.Close()

         If Request("l") = "" AND Request("catid") = "" Then

         pan2.visible = True
         pan1.visible = false

         End If
    
 End Sub
'+++++++++++++++++++++++++++++++++++++++++++++



'+++++++++++++++++++++++++++++++++++++++++++++
'Handles BindData SQL statements
Sub BindData_SQL_Statements()

        'Check if it is a sort category from the dropdownlist or not
         If Request("catid") <> "" then

           'Lets hide some unneeded control
           pan1.visible = True
           pan2.visible = False
           lblrecordperpageFooter.visible = False
           lbCountRecipeFooter.visible = False
           lblrecordperpageTop.visible = False

             'Carry on our SQL statement
             strSearchSQL = " Where CAT_ID =" & Replace(Request("catid"),"'","''")

         End if

         If Request("lc") <> "" then

           'Lets hide some unneeded control
           pan1.visible = True
           pan2.visible = False       
           lblrecordperpageFooter.visible = False
           lbCountRecipeFooter.visible = False
           lblrecordperpageTop.visible = False

             'Carry on our SQL statement
             strSearchSQL = " Where CAT_ID =" & Replace(Request("catid"),"'","''") & " AND Name LIKE '" & Replace(Request("lc"),"'","''") & "%" & "'"

         End if

         'Check if it is a sort alphabet letter, then carry on SQL
         If Request("l") <> "" then

              pan1.visible = False
              pan2.visible = True

             'Carry on our SQL statement
             strSearchSQL = " WHERE Name LIKE '" & Replace(Request("l"),"'","''") & "%" & "'"

         End if

         'Search parameters
          if Request("find") <> "" AND Request("s") = "1" then

             strSearchSQL = " WHERE Name LIKE '%" & Replace(Request("find"),"'","''") & "%'"

          elseif Request("find") <> "" AND Request("s") = "2" then

             strSearchSQL = " WHERE Author LIKE '%" & Replace(Request("find"),"'","''") & "%'"

          elseif Request("find") <> "" AND Request("s") = "3" then

             strSearchSQL = " WHERE ID LIKE '%" & Replace(Request("find"),"'","''") & "%'"

         end if

         'Display the approval tab
         If Request("tab") = "1" Then

            strSearchSQL = " Where LINK_APPROVED = 0" 

            'Hide panel
            Panel1.visible = False
            pan1.visible = False
            pan2.visible = False
            
           'Hide the footer total records count
           LblPageInfo.visible = false
           lbCountRecipeFooter.visible = False
           lblrecordperpageFooter.visible = False
           lblrecordperpageTop.visible = False
           approvallink.enabled = false
           lblSortedCat.Text = "How To? - To approve a recipe, click the Recipe Name link inside the grid."
           lblthese.text = "There are&nbsp;"
           lblthese2.text = "&nbsp;recipe(s) waiting for approval." 
           lblmangermainpage.text = "Recipe Manager Main"
           lblmangermainpagelink.ToolTip = "Back to Main Recipe Manager Page"
           lblunapproved2.visible = True

         End if

        'Creates the SQL statement
         strSQL = "SELECT * FROM Recipes" & strSearchSQL 

       'Add the ORDER BY clause, if necessary
       If SortExpression.Length > 0 Then 
         strSQL &= " ORDER BY " & SortExpression
      
          If SortAscending Then
             strSQL &= " ASC"
          Else
             strSQL &= " DESC"
          End If

       End If

 End Sub
'+++++++++++++++++++++++++++++++++++++++++++++



'+++++++++++++++++++++++++++++++++++++++++++++
'Handle pager buttons enabled and disabled
 Sub EnabledDisabled_PagerButtons()

      'Disabled and enabled footer pager button
      If dgrd_recipe.CurrentPageIndex <> 0 Then 

          Prev_Buttons() 
          Firstbutton.enabled = true 
          Prevbutton.enabled = true 
          FirstButton.ToolTip = "Go back to first page"

   else

          Firstbutton.enabled = false 
          Prevbutton.enabled = false 
          FirstButton.ToolTip = ""

   End if 

   If dgrd_recipe.CurrentPageIndex <> (dgrd_recipe.PageCount-1) then 

        Next_Buttons() 
        NextButton.enabled = true 
        Lastbutton.enabled = true 
        Lastbutton.Tooltip = "Go to the last page" 

   else
        
        LstRecpage.enabled = false  'if there are less than 20 records, we're not going to allow page size change, so disabled the dropdownlist
        NextButton.enabled = false 
        Lastbutton.enabled = False 
        NextButton.ToolTip = ""
        Lastbutton.Tooltip = "" 

  End if

 End Sub
'+++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
 'Handle PageSize Option
 Sub DisplayPagerecord_Click(sender As Object, e As EventArgs)

   Dim lstRecPerPage as string = LstRecpage.SelectedItem.Value

    'Check how many records per page to display     
     Select lstRecPerPage  'lstRecPerpage is the ID of the Dropdownlist control

        Case "10" 

          dgrd_recipe.PageSize = 10
          lblrecordperpage.text = "Showing 10 records per page" 
          lblrecordperpageFooter.text = "Showing 10 records per page"
          lblrecordperpageTop.text = "- 10 records per page"   

       Case "20" 

          dgrd_recipe.PageSize = 20
          lblrecordperpage.text = "Default 20 records per page"
          lblrecordperpageFooter.text = "Showing default 20 records per page"
          lblrecordperpageTop.text = "- 20 records per page"            

       Case "30" 

          dgrd_recipe.PageSize = 30
          lblrecordperpage.text = "Showing 30 records per page" 
          lblrecordperpageFooter.text = "Showing 30 records per page"
          lblrecordperpageTop.text = "- 30 records per page"    

       Case "40"

          dgrd_recipe.PageSize = 40
          lblrecordperpage.text = "Showing 40 records per page"   
          lblrecordperpageFooter.text = "Showing 40 records per page" 
          lblrecordperpageTop.text = "- 40 records per page" 

       Case "50" 

          dgrd_recipe.PageSize = 50
          lblrecordperpage.text = "Showing 50 records per page" 
          lblrecordperpageFooter.text = "Showing 50 records per page" 
          lblrecordperpageTop.text = "- 50 records per page" 

       Case "60"

          dgrd_recipe.PageSize = 60
          lblrecordperpage.text = "Showing 60 records per page" 
          lblrecordperpageFooter.text = "Showing 60 records per page"
          lblrecordperpageTop.text = "- 60 records per page" 

       Case "80" 

          dgrd_recipe.PageSize = 80
          lblrecordperpage.text = "Showing 80 records per page" 
          lblrecordperpageFooter.text = "Showing 80 records per page"
          lblrecordperpageTop.text = "- 80 records per page"  

       Case "100"

          dgrd_recipe.PageSize = 100
          lblrecordperpage.text = "Showing 100 records per page"
          lblrecordperpageFooter.text = "Showing 100 records per page"
          lblrecordperpageTop.text = "- 100 records per page" 

       Case Else

          dgrd_recipe.PageSize = 20
          lblrecordperpage.text = "Default 20 records per page"
          lblrecordperpageFooter.text = "Showing default 20 records per page"
          lblrecordperpageTop.text = "- 20 records per page"      

       End Select

       'Now, bind the data and get the page size from the dropdownlist
       BindData()

 End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
 'Handles search admin
 Sub AdminSearch_Click(sender as object, e as EventArgs)

    strURLRedirect = "recipemanager.aspx?find=" & find.Text & "&s=" & sopt.SelectedItem.Value
    Server.Transfer(strURLRedirect)

 End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
  'Display total number of recipes in the database
  Sub DisplayRecipeCount()

        strSQL = "Select Count(ID) From Recipes"
        DBconnect()
        objConnection.Open()
        lbCountRecipe.Text = "Total Recipes:&nbsp;" & objCommand.ExecuteScalar()
        lbCountRecipeFooter.Text = objCommand.ExecuteScalar() & "&nbsp;records"
        objConnection.Close()

  End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
  'Display total number of unapprove recipes
  Sub UnApproveRecipe()

        strSQL = "Select Count(ID) From Recipes Where LINK_APPROVED = 0"
        DBconnect()
        objConnection.Open()
        lblunapproved.Text = "Waiting For Approval:&nbsp;" & objCommand.ExecuteScalar() 
        lblunapproved2.Text = objCommand.ExecuteScalar() 
        objConnection.Close()

  End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
  'Display total number of comments
  Sub DisplayCommentsCount()

        strSQL = "Select Count(ID) From COMMENTS_RECIPE"
        DBconnect()
        objConnection.Open()
        lbCountComments.Text = "Total Comments:&nbsp;" & objCommand.ExecuteScalar()
        objConnection.Close()

  End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
  'Display total number of categories in the category table
  Sub DisplayCategoryCount()

        strSQL = "Select Count(CAT_ID) From RECIPE_CAT"
        DBconnect()
        objConnection.Open()
        lbCountCat.Text = "Total Category:&nbsp;" & objCommand.ExecuteScalar()
        objConnection.Close()

  End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
   'Handle sort category selection redirect
   Sub GetCatName(sender as object, e as EventArgs)

         strURLRedirect = "recipemanager.aspx?catid=" & Request("CategoryName")
         Server.Transfer(strURLRedirect)

   End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
  'Previous footer button
  Sub Prev_Buttons()

     Dim PrevSet As String
     If dgrd_recipe.CurrentPageIndex+1 <> 1 and ResultCount <> -1 Then

         PrevSet = dgrd_recipe.PageSize
         PrevButton.ToolTip = "Go back to previous page: (" & PrevSet & ") records per page"
	
   End If

 End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
  'Next footer button
  Sub Next_Buttons()

    Dim NextSet As String
    If dgrd_recipe.CurrentPageIndex+1 < dgrd_recipe.PageCount Then

        NextSet = dgrd_recipe.PageSize
        NextButton.ToolTip = "Go to next page: (" & NextSet & ") records per page"

     End If

  End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
  'Handles footer button pager click event
  Sub FooterPager_ButtonClick(sender As Object, e As EventArgs)

  'Used by external paging
  Dim arg As String = sender.CommandArgument

  Select arg
     Case "next":  'The next Button was Clicked
        If (dgrd_recipe.CurrentPageIndex < (dgrd_recipe.PageCount - 1)) Then
            dgrd_recipe.CurrentPageIndex += 1
        End If 

     Case "prev":   'The prev button was clicked
         If (dgrd_recipe.CurrentPageIndex > 0) Then
             dgrd_recipe.CurrentPageIndex -= 1
         End If

     Case "last":   'The Last Page button was clicked
         dgrd_recipe.CurrentPageIndex = (dgrd_recipe.PageCount - 1)

     Case Else:     'The First Page button was clicked
         dgrd_recipe.CurrentPageIndex = Convert.ToInt32(arg)
	End Select

    'Now, bind the data!
    BindData()

  End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
  'Handle edit databound
  Sub Edit_Handle(sender as Object, e As DataGridCommandEventArgs)

        If (e.CommandName="edit") then
            Dim iIdNumber as TableCell = e.Item.Cells(0)     

            strURLRedirect = "editing.aspx?id=" & iIdNumber.Text
            Server.Transfer(strURLRedirect)

        End if

  End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
 'Display category name in the dropdownlist
 Sub GetDropdownlistCategory()

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
            CategoryName.Items.Insert(0, new ListItem("Choose Category"))

 End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
  'Handle the deletebutton click event
  Sub Delete_Recipes(sender as Object, e As DataGridCommandEventArgs)

        If (e.CommandName="Delete") then
        Dim iIdNumber2 as TableCell = e.Item.Cells(0)
        Dim iIRecipename as TableCell = e.Item.Cells(1)     
    
        objConnection = New OledbConnection(strConnection)
        objConnection.Open()
    
        strSQL = "delete * from Recipes where ID = " & iIdNumber2.Text
        objCommand = New OledbCommand(strSQL,objConnection)
        objCommand.ExecuteNonQuery()
    
        objCommand = nothing
        objConnection.Close()
        objConnection = nothing

        'Redirect to confirm delete page
        strURLRedirect = "confirmdel.aspx?catname=" & iIRecipename.Text & "&mode=del"
        Server.Transfer(strURLRedirect)
    
   End If

 End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
 Sub dgRecipe_ItemDataBound(sender as Object, e as DataGridItemEventArgs)

     Dim strIDcell as integer = DataBinder.Eval(e.Item.DataItem, "ID")

    'First, make sure we're not dealing with a Header or Footer row
    If e.Item.ItemType <> ListItemType.Header AND e.Item.ItemType <> ListItemType.Footer then

      Dim deleteButton as LinkButton = e.Item.Cells(7).Controls(0)
      Dim editButton as LinkButton = e.Item.Cells(6).Controls(0)

     'We can now add the onclick event handler
     deleteButton.Attributes("onclick") = "javascript:return confirm('Are you sure you want to delete Recipe ID # " & _
     DataBinder.Eval(e.Item.DataItem, "ID") & "?')" 
     deleteButton.ToolTip = "Delete recipe (" & DataBinder.Eval(e.Item.DataItem, "Name") & ") ID #:" & DataBinder.Eval(e.Item.DataItem, "ID")  
     editButton.ToolTip = "Edit recipe (" & DataBinder.Eval(e.Item.DataItem, "Name") & ") ID #:" & DataBinder.Eval(e.Item.DataItem, "ID")  


    'Data row mouseover changecolor
    e.Item.Attributes.Add("onmouseover", "this.style.backgroundColor='#ECF5FF'")
    e.Item.Attributes.Add("onmouseout", "this.style.backgroundColor='#ffffff'")

    'handle cell 1 Recipe name change cell and font color
    e.Item.Cells(1).Attributes.Add("onmouseover", "this.style.backgroundColor='#F0E68C';this.style.cursor='pointer';this.style.color='#ff3e3e'")
    e.Item.Cells(1).Attributes.Add("onmouseout", "this.style.backgroundColor='#fff';this.style.cursor='pointer';this.style.color='#048'")
    e.Item.Cells(1).ForeColor = System.Drawing.ColorTranslator.FromHtml("#048")

    'Handle cell 1 - Recipe name click event
    e.Item.Cells(1).Attributes.Add("Onclick", "javascript:window.open('viewing.aspx?id=" & strIDcell & "'," & "'','height=690,width=700')")
    
    'Display cell tooltip in the grid
    e.Item.Cells(0).ToolTip = "Recipe # " & DataBinder.Eval(e.Item.DataItem, "ID")
    e.Item.Cells(1).ToolTip = "Click to view: " & DataBinder.Eval(e.Item.DataItem, "Name") & " recipe"
    e.Item.Cells(2).ToolTip = "Category: " & DataBinder.Eval(e.Item.DataItem, "Category")
    e.Item.Cells(3).ToolTip = "Recipe author: " & DataBinder.Eval(e.Item.DataItem, "Author")
    e.Item.Cells(4).ToolTip = "Submitted on: " & DataBinder.Eval(e.Item.DataItem, "Date") 
    e.Item.Cells(5).ToolTip = "This recipe has been viewed: " & DataBinder.Eval(e.Item.DataItem, "Hits")    

   'If we're in the approval tab, then we change the tooltip
   If request("tab") = "1" Then

     e.Item.Cells(1).ToolTip = "Waiting for approval, click to approve: " & DataBinder.Eval(e.Item.DataItem, "Name") & " recipe"

  End if

  'Display the sorted category name from the dropdownlist
  If request("catid") <> "" Then

        lblSortedCat.Text = "Sorted Category: " & DataBinder.Eval(e.Item.DataItem, "Category")
        lblletterlegendcat.text = DataBinder.Eval(e.Item.DataItem, "Category") & "&nbsp;Recipes A-Z:"

  End If

  End If

    'Handles the header link tooltip
    'First, make sure we're dealing with a Header
    If e.Item.ItemType = ListItemType.Header Then

        'Display cell header tooltip
        e.Item.Cells(0).ToolTip = "Sort by ID - ASC or DESC"
        e.Item.Cells(1).ToolTip = "Sort by Recipe Name"
        e.Item.Cells(2).ToolTip = "Sort by Recipe Category"
        e.Item.Cells(3).ToolTip = "Sort by Author"
        e.Item.Cells(4).ToolTip = "Sort by Submitted date"
        e.Item.Cells(5).ToolTip = "Sort by Most Popular - Hits"

        'handle clickable and change color cell header on mouseover
        e.Item.Cells(0).Attributes.Add("onmouseover","this.style.backgroundColor='#428DFF';this.style.cursor='pointer'")
        e.Item.Cells(0).Attributes.Add("onmouseout","this.style.backgroundColor='#79AEFF';this.style.cursor='pointer'")
        e.Item.Cells(0).Attributes.Add("Onclick", "javascript:__doPostBack('dgrd_recipe$_ctl2$_ctl0')")
        e.Item.Cells(1).Attributes.Add("onmouseover","this.style.backgroundColor='#428DFF';this.style.cursor='pointer'")
        e.Item.Cells(1).Attributes.Add("onmouseout","this.style.backgroundColor='#79AEFF';this.style.cursor='pointer'")
        e.Item.Cells(1).Attributes.Add("Onclick", "javascript:__doPostBack('dgrd_recipe$_ctl2$_ctl1')")
        e.Item.Cells(2).Attributes.Add("onmouseover","this.style.backgroundColor='#428DFF';this.style.cursor='pointer'")
        e.Item.Cells(2).Attributes.Add("onmouseout","this.style.backgroundColor='#79AEFF';this.style.cursor='pointer'")
        e.Item.Cells(2).Attributes.Add("Onclick", "javascript:__doPostBack('dgrd_recipe$_ctl2$_ctl2')")
        e.Item.Cells(3).Attributes.Add("onmouseover","this.style.backgroundColor='#428DFF';this.style.cursor='pointer'")
        e.Item.Cells(3).Attributes.Add("onmouseout","this.style.backgroundColor='#79AEFF';this.style.cursor='pointer'")
        e.Item.Cells(3).Attributes.Add("Onclick", "javascript:__doPostBack('dgrd_recipe$_ctl2$_ctl3')")
        e.Item.Cells(4).Attributes.Add("onmouseover","this.style.backgroundColor='#428DFF';this.style.cursor='pointer'")
        e.Item.Cells(4).Attributes.Add("onmouseout","this.style.backgroundColor='#79AEFF';this.style.cursor='pointer'")
        e.Item.Cells(4).Attributes.Add("Onclick", "javascript:__doPostBack('dgrd_recipe$_ctl2$_ctl4')")
        e.Item.Cells(5).Attributes.Add("onmouseover","this.style.backgroundColor='#428DFF';this.style.cursor='pointer'")
        e.Item.Cells(5).Attributes.Add("onmouseout","this.style.backgroundColor='#79AEFF';this.style.cursor='pointer'")
        e.Item.Cells(5).Attributes.Add("Onclick", "javascript:__doPostBack('dgrd_recipe$_ctl2$_ctl5')")

    End if

   'Change the color of the column search result, depending on the filter criteria
    Dim strSR as integer = request("s")
     
      Select Case strSR

          Case "1"

             if e.Item.ItemType <> ListItemType.Header

            e.Item.Cells(1).BackColor = System.Drawing.Color.Ivory
            e.Item.Cells(1).Attributes.Add("onmouseover", "this.style.backgroundColor='#F0E68C';this.style.cursor='pointer'")
            e.Item.Cells(1).Attributes.Add("onmouseout", "this.style.backgroundColor='Ivory'")
            e.Item.Cells(1).Attributes.Add("onmouseover", "this.style.backgroundColor='#F0E68C';this.style.cursor='pointer';this.style.color='#ff3e3e'")
            e.Item.Cells(1).Attributes.Add("onmouseout", "this.style.backgroundColor='Ivory';this.style.cursor='pointer';this.style.color='#048'")

           else

                 dgrd_recipe.HeaderStyle.BackColor = System.Drawing.ColorTranslator.FromHtml("#79AEFF")

           end if
     
          Case "2"
             
             if e.Item.ItemType <> ListItemType.Header

                 e.Item.Cells(3).BackColor = System.Drawing.Color.Ivory
                 e.Item.Cells(3).Attributes.Add("onmouseover", "this.style.backgroundColor='Snow'")
                 e.Item.Cells(3).Attributes.Add("onmouseout", "this.style.backgroundColor='Ivory'")

           else

                 dgrd_recipe.HeaderStyle.BackColor = System.Drawing.ColorTranslator.FromHtml("#79AEFF")

           end if

         Case "3"

           if e.Item.ItemType <> ListItemType.Header

                e.Item.Cells(0).BackColor = System.Drawing.Color.Ivory
                e.Item.Cells(0).Attributes.Add("onmouseover", "this.style.backgroundColor='Snow'")
                e.Item.Cells(0).Attributes.Add("onmouseout", "this.style.backgroundColor='Ivory'")

           else

                dgrd_recipe.HeaderStyle.BackColor = System.Drawing.ColorTranslator.FromHtml("#79AEFF")

           end if

    End Select

   'Change the color of the recipe name column when sorted by alpha letter
   if request("l") <> "" Then

         if e.Item.ItemType <> ListItemType.Header

            e.Item.Cells(1).BackColor = System.Drawing.Color.Ivory
            e.Item.Cells(1).Attributes.Add("onmouseover", "this.style.backgroundColor='#F0E68C';this.style.cursor='pointer'")
            e.Item.Cells(1).Attributes.Add("onmouseout", "this.style.backgroundColor='Ivory'")
            e.Item.Cells(1).Attributes.Add("onmouseover", "this.style.backgroundColor='#F0E68C';this.style.cursor='pointer';this.style.color='#ff3e3e'")
            e.Item.Cells(1).Attributes.Add("onmouseout", "this.style.backgroundColor='Ivory';this.style.cursor='pointer';this.style.color='#048'")

          else

            dgrd_recipe.HeaderStyle.BackColor = System.Drawing.ColorTranslator.FromHtml("#79AEFF")

         end if

   end if

   'Change the color of the recipe name column when sorted by alpha letter
   if request("lc") <> "" Then

        if e.Item.ItemType <> ListItemType.Header

            e.Item.Cells(1).BackColor = System.Drawing.Color.Ivory
            e.Item.Cells(1).Attributes.Add("onmouseover", "this.style.backgroundColor='#F0E68C';this.style.cursor='pointer'")
            e.Item.Cells(1).Attributes.Add("onmouseout", "this.style.backgroundColor='Ivory'")
            e.Item.Cells(1).Attributes.Add("onmouseover", "this.style.backgroundColor='#F0E68C';this.style.cursor='pointer';this.style.color='#ff3e3e'")
            e.Item.Cells(1).Attributes.Add("onmouseout", "this.style.backgroundColor='Ivory';this.style.cursor='pointer';this.style.color='#048'")

          else

            dgrd_recipe.HeaderStyle.BackColor = System.Drawing.ColorTranslator.FromHtml("#79AEFF")

       end if

   end if

   'Change the color of the category column when sorted category
   if request("catid") <> "" Then

       if e.Item.ItemType <> ListItemType.Header

            e.Item.Cells(2).BackColor = System.Drawing.Color.FloralWhite
            e.Item.Cells(2).Attributes.Add("onmouseover", "this.style.backgroundColor='#F0E68C'")
            e.Item.Cells(2).Attributes.Add("onmouseout", "this.style.backgroundColor='#FFFAF0'")

         else

            dgrd_recipe.HeaderStyle.BackColor = System.Drawing.ColorTranslator.FromHtml("#79AEFF")

      end if

   end if

    'Display the pagecount in the top and footer
    Dim pageindex As Integer = dgrd_recipe.CurrentPageIndex + 1
    LblPageInfo.Text = "Page " + pageindex.ToString() + " of " + dgrd_recipe.PageCount.ToString() + "&nbsp;-&nbsp;"
    LblPageInfoTop.Text = "Showing Page " + pageindex.ToString() + " of " + dgrd_recipe.PageCount.ToString() 

 End Sub
'++++++++++++++++++++++++++++++++++++++++++++

  

'++++++++++++++++++++++++++++++++++++++++++++
 'The SortCommand event handler
  Sub Recipes_SortCommand(sender as Object, e as DataGridSortCommandEventArgs)
    'Toggle SortAscending if the column that the data was sorted by has
    'been clicked again...
    If e.SortExpression = Me.SortExpression Then 
      SortAscending = Not SortAscending
    Else
      SortAscending = True
    End If

    'Set the SortExpression property to the SortExpression passed in
    Me.SortExpression = e.SortExpression

    BindData()  'rebind the DataGrid data
  End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
 'The Page-level properties that write to ViewState
 Private Property SortExpression() As String
    Get
        Dim o As Object = viewstate("SortExpression")
        If o Is Nothing Then
            Return String.Empty
        Else
            Return o.ToString
        End If
    End Get
    Set(ByVal Value As String)
        viewstate("SortExpression") = Value
    End Set
 End Property
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
 Private Property SortAscending() As Boolean
    Get
        Dim o As Object = viewstate("SortAscending")
        If o Is Nothing Then
            Return True
        Else
            Return Convert.ToBoolean(o)
        End If
    End Get
    Set(ByVal Value As Boolean)
        viewstate("SortAscending") = Value
    End Set
 End Property
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
  Sub UpdateColumnHeaders(ByVal dgrid As DataGrid)

    Dim c As DataGridColumn
    For Each c In dgrid.Columns
        c.HeaderText = Regex.Replace(c.HeaderText, "\s<.*>", String.Empty)  'Clear any <img> tags that might be present
        
        If c.SortExpression = SortExpression Then
            If SortAscending Then
                c.HeaderText &= " <img src=""../images/arrow_up2.gif"" title=""Sort By Descending Order"" border=""0"">"
                lblsortorder.text = "Sorted By Ascending" 
                orderimage.ImageUrl = "../images/arrow_down2.gif"
                orderimage.visible = True
            Else
                c.HeaderText &= " <img src=""../images/arrow_down2.gif"" title=""Sort By Ascending Order"" border=""0"">"
                lblsortorder.text = "Sorted By Descending"
                orderimage.ImageUrl = "../images/arrow_up2.gif"
                orderimage.visible = True
            End If
        End If
    Next

 End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
   'Handles page change links - paging system
   Sub New_Page(sender As Object, e As DataGridPageChangedEventArgs)

         dgrd_recipe.CurrentPageIndex = e.NewPageIndex
         BindData()

  End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'++++++++++++++++++++++++++++++++++++++++++++
'Database connection string
 Sub DBconnect()

     objConnection = New OledbConnection(strConnection)
     objCommand = New OledbCommand(strSQL, objConnection)

 End Sub
'++++++++++++++++++++++++++++++++++++++++++++



'+++++++++++++++++++++++++++++++++++++++++++++++++
'Here we declare our module-level variables 

  Private strSQL as string
  Private strURLRedirect as string
  Private ResultCount as Integer
  Private strSearchSQL as string

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
    <td width="100%" colspan="2" align="left">
<div style="padding-left: 20px;"><h3>Recipe Manager</h3></div>
</td>
  </tr>
  <tr>
    <td width="23%" align="left">
<div style="padding-left: 20px;"><asp:Label font-name="verdana" font-size="9" ID="lblusername" runat="server" /></div>
</td>
    <td width="79%" align="left">
<div style="padding-left: 2px;"><asp:Label  ID="lblSortedCat" font-name="verdana" font-size="9" runat="server" /> <asp:Label ID="lblrcdCatcount" font-name="verdana" font-size="9" runat="server" /> <asp:Label ID="lblsortorder" font-name="verdana" font-size="9" runat="server" /> <asp:Image id="orderimage" runat="server" visible="false" />
</td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <table border="0" cellpadding="0" cellspacing="0" width="100%">
        <tr>
          <td width="23%" align="left"></td>
          <td width="41%" align="left"><asp:Label ID="LblPageInfoTop" cssClass="content2" runat="server" />
<asp:Label ID="lblrecordperpageTop" cssClass="content2" runat="server" />
</td>
          <td width="38%" align="right">
<div style="padding-right: 25px; padding-bottom: 4px;">
<asp:Panel id="panel1" runat="server">
<span class="content2"><b>Sort Category:</b></span><asp:listbox id="CategoryName" cssClass="cselect" runat="server" Rows="1" 
               DataTextField="CAT_TYPE" DataValueField="CAT_ID" /> 
<asp:Button runat="server" ID="GO" OnClick="GetCatName" cssclass="submit" Text="Display"/>
</asp:Panel>
<asp:Label  ID="lblthese" font-name="verdana" font-size="9" runat="server" />
<asp:Label ID="lblunapproved2" font-name="verdana" font-size="9" runat="server" />
<asp:Label ID="lblthese2" font-name="verdana" font-size="9" runat="server" />
 </div>
   </td>
       </tr>
      </table>
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
<span class="bluearrow2">»</span>&nbsp;<asp:HyperLink runat="server" ID="lblmangermainpagelink" NavigateUrl="recipemanager.aspx"><asp:Label id="lblmangermainpage" runat="server" /></asp:HyperLink>
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

<!--Begin Admin Page Record Panel-->
 <div class="roundcont2">
<div class="roundtop">
<img src="../images/hleft.gif" height="5" width="5" alt="" class="corner">
<div class="dcnt"><span class="content3">Page Options</span></div> 
</div>
</div>
<div class="contentdisplay">
<div class="contentdis5">
<div style="text-align: center;">
<asp:Dropdownlist runat="server" ID="LstRecpage" AutoPostBack="True" OnSelectedIndexChanged="DisplayPagerecord_Click" class="cselect" width="150px">
<asp:Listitem Value="20" selected>Default 20 Records</asp:Listitem>
<asp:Listitem Value="10">10 Records Per Page</asp:Listitem>
<asp:Listitem Value="30">30 Records Per Page</asp:Listitem>
<asp:Listitem Value="40">40 Records Per Page</asp:Listitem>
<asp:Listitem Value="50">50 Records Per Page</asp:Listitem>
<asp:Listitem Value="60">60 Records Per Page</asp:Listitem>
<asp:Listitem Value="80">80 Records Per Page</asp:Listitem>
<asp:Listitem Value="100">100 Records Per Page</asp:Listitem>
</asp:Dropdownlist>
<br />
<asp:Label ID="lblrecordperpage" cssClass="content2" runat="server" />
</div>
</div>
</div>
<br />
<!--End Admin Page Record Panel-->

<!--Begin search panel-->
 <div class="roundcont2">
<div class="roundtop">
<img src="../images/hleft.gif" height="5" width="5" alt="" class="corner">
<div class="dcnt"><span class="content3">Admin Search</span></div> 
</div>
</div>
<div class="contentdisplay">
<div class="contentdis5">
<img src="../images/search.gif" border="0" alt="Search recipe" align="absmiddle">
<asp:TextBox runat="server" id="find" class="textbox" size="15" />
<br />
<asp:radiobuttonlist id="sopt" repeatdirection="Vertical" font-name="verdana" font-size="9"  runat="server">
<asp:listitem value="1" text="Recipe Name" selected="true" />
<asp:listitem value="2" text="Recipe Author" />
<asp:listitem value="3" text="Recipe ID" />
</asp:radiobuttonlist>
&nbsp;<asp:Button runat="server" class="submit" tooltip="Go find it!" OnClick="AdminSearch_Click" Text="Search"/>
</div>
</div>
<!--End search panel-->
</td>
    <td width="79%" valign="top">
<!--Begin Alphabet Letter links-->
<div class="divtoplet">
<img src="../images/hleft.gif" height="5" width="5" alt="" class="corner">
<div style="text-align: left; background-color: #6898D0; padding-left: 10px; padding-top: 0px; padding-bottom: 3px; height: 19px;">
<asp:Panel ID="pan1" runat="server">
<asp:Label Id="lblletterlegendcat" class="corange" runat="server" />&nbsp;
<asp:Label id="lblalphalettercat" runat="server" />
</asp:Panel>
<asp:Panel ID="pan2" runat="server">
<asp:Label id="lblletterlegend" cssClass="corange" runat="server" />&nbsp;
<asp:Label id="lblalphaletter" runat="server" />
</asp:Panel>
</div>
</div>
<!--End Alphabet Letter links-->
<!--Begin display datagrid-->
<asp:DataGrid runat="server" id="dgrd_recipe" cssclass="hlink" AutoGenerateColumns="False" AllowSorting="true"
     Backcolor="#ffffff" BorderStyle="none" BorderColor="#E1EDFF" cellpadding="5" Width="95%" HorizontalAlign="Center"  onSortCommand="Recipes_SortCommand" AllowPaging="True" OnPageIndexChanged="New_Page"  OnItemDataBound="dgRecipe_ItemDataBound" DataKeyField="ID" OnDeleteCommand="Delete_Recipes" onItemCommand="Edit_Handle"> 
     <HeaderStyle Font-Bold="True" BackColor="#79AEFF" cssclass="header" />
     <AlternatingItemStyle BackColor="White" />                                  
     <Columns>    
     <asp:BoundColumn DataField="ID" HeaderText="ID" SortExpression="ID" />   
     <asp:BoundColumn DataField="Name" HeaderText="Recipe Name" SortExpression="Name" />
     <asp:BoundColumn DataField="Category" HeaderText="Category" SortExpression="Category" />
     <asp:BoundColumn DataField="Author" HeaderText="Author" SortExpression="Author" />
     <asp:BoundColumn DataField="Date" DataFormatString="{0:d}" HeaderText="Date" SortExpression="Date" />
     <asp:BoundColumn DataField="Hits" HeaderText="Hits" SortExpression="Hits" />
     <asp:ButtonColumn Text='<img border="0" src="../images/icon_edit.gif">' HeaderText="Edit" CommandName="edit" />
     <asp:ButtonColumn Text='<img border="0" src="../images/icon_delete.gif">' HeaderText="Delete" CommandName="Delete" />
     </Columns>
     <PagerStyle Mode="NumericPages" BackColor="#fcfcfc" HorizontalAlign="left" />
    </asp:DataGrid>
<!--End display datagrid-->

<!--Begin display pager button-->
<div style="margin-left: 20px; margin-right: 20px; margin-top: 5px; background-color: #fff; padding-left: 2px; padding-top: 1px; border: #fff 1px solid;">
<asp:Linkbutton id="Firstbutton" Text='<img border="0" src="../images/firstpage.gif" align="absmiddle">' CommandArgument="0" runat="server" onClick="FooterPager_ButtonClick"/>
<asp:linkbutton id="Prevbutton" Text='<img border="0" src="../images/prevpage.gif" align="absmiddle">' CommandArgument="prev" runat="server" onClick="FooterPager_ButtonClick"/>
<asp:linkbutton id="Nextbutton" Text='<img border="0" src="../images/nextpage.gif" align="absmiddle">' CommandArgument="next" runat="server" onClick="FooterPager_ButtonClick"/>
<asp:linkbutton id="Lastbutton" Text='<img border="0" src="../images/lastpage.gif" align="absmiddle">' CommandArgument="last" runat="server" onClick="FooterPager_ButtonClick"/>&nbsp;&nbsp;&nbsp;
<asp:Label ID="LblPageInfo" cssClass="content2" runat="server" />
<asp:Label ID="lblrcdCatcountfooter" cssClass="content2" runat="server" /> <asp:Label ID="lbCountRecipeFooter" cssClass="content2" runat="server" /> <asp:Label ID="lblrcdalphaletterfooter" cssClass="content2" runat="server" />&nbsp;&nbsp;<asp:Label ID="lblrecordperpageFooter" cssClass="content2" runat="server" />
</div>
<!--End display pager button-->         
</td>
  </tr>
</table>
</form>
<div style="text-align: center; margin-top: 25px; padding-bottom: 30px;">
<a href="http://www.ex-designz.net" class="hlink" title="Visit our website">Powered By Ex-designz.net World Recipe</a>
</div>
</body>
</html>