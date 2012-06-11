<script runat="server">
    
    'Get the database server Map Path
    Function DB_Path()
        
       if instr(Context.Request.ServerVariables("PATH_TRANSLATED"),"Recipes") then
          DB_Path = System.Web.HttpContext.Current.Server.MapPath("/db/recipedb.mdb")
       else
          DB_Path = System.Web.HttpContext.Current.Server.MapPath("/db/recipedb.mdb")
       end if

    End Function


    'Database connection
    Sub DataBase_Connect(strSQL)
    
        objConnection = New OledbConnection(strConnection)
        objConnection.Open()
        objCommand = New OledbCommand(strSQL, objConnection)
        objDataReader  = objCommand.ExecuteReader()
    
    End Sub
    
    'Database disconnect
    Sub DataBase_Disconnect()
    
        objDataReader.Close()
        objConnection.Close()
    
    End Sub
    
    'Check if user has started a session
    Function Check_User()
   
	if session("userid") = "" then
    	  Server.Transfer("Adminlogin.aspx")
    	end if
    
    End function

    Private strDBLocation = DB_Path()
    Private strConnection as string = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBLocation
    Private objConnection
    Private objCommand
    Private objDataReader as OledbDataReader

</script>

