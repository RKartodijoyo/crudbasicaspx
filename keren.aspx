<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDB" %>
<%@ Import Namespace="System.Data.OleDb.OleDbCommand" %>
<html><head><title>Data Mahasiswa</title>

<script language="VB" runat="server">
    Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        Dim strConn As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("database.mdb") & ";"
        Dim MySQL As String = "SELECT Name, NIK, Alamat, Jurusan, MatkulBind, MatkulMatdas ,MatkulAgama, MatkulKewarganegaraan FROM mahasiswa"
        Dim MyConn As New OleDbConnection(strConn)
        Dim Cmd As New OleDbCommand(MySQL, MyConn)
        MyConn.Open()
        rptGuestbook.DataSource = Cmd.ExecuteReader(System.Data.CommandBehavior.CloseConnection)
        rptGuestbook.DataBind()
    End Sub
    Sub OnBtnSendClicked(ByVal s As Object, ByVal e As EventArgs)
      
      
        If txtNIK.Text = String.Empty Then
            
        
       
        Else
            Dim strConn As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("database.mdb") & ";"
            Dim checkNIK As String = "SELECT * FROM mahasiswa WHERE= NIK'" & txtNIK.Text & "'"
            
            If checkNIK Is Nothing Then
            
            
            
                
            
                Dim MySQL As String = "INSERT INTO mahasiswa (Name, NIK, Alamat, Jurusan, MatkulBind, MatkulMatdas ,MatkulAgama, MatkulKewarganegaraan) VALUES ('" & txtName.Text & "','" & txtNIK.Text & "','" & txtAlamat.Text & "','" & txtJurusan.Text & "','" & ab.Text & "','" & ac.Text & "','" & ad.Text & "','" & ae.Text & "')"
                Dim MyConn As New OleDbConnection(strConn)
                Dim cmd As New OleDbCommand(MySQL, MyConn)
                MyConn.Open()
                cmd.ExecuteNonQuery()
                MyConn.Close()
                Response.Redirect("keren.aspx")
            Else
                
            End If
            
        End If
    End Sub

    Sub OnBtnClearClicked(ByVal s As Object, ByVal e As EventArgs)
        txtName.Text = ""
        txtNIK.Text = ""
        txtAlamat.Text = ""
		        
    End Sub
    Sub OnBtnDeleteClicked(ByVal s As Object, ByVal e As EventArgs)
       
        Dim strConn As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("database.mdb") & ";"
	
        Dim MySQL As String = "DELETE FROM mahasiswa " & " WHERE NIK = '" & txtNIK.Text & "'"
        Dim MyConn As New OleDbConnection(strConn)
        Dim cmd As New OleDbCommand(MySQL, MyConn)
        MyConn.Open()
        cmd.ExecuteNonQuery()
        MyConn.Close()
        Response.Redirect("keren.aspx")
      
    End Sub
    Sub OnBtnUpdateClicked(ByVal s As Object, ByVal e As EventArgs)
        Dim strConn As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("database.mdb") & ";"
        Dim MySQL As String = "UPDATE mahasiswa SET Name = '" & txtName.Text & "', Alamat ='" & txtAlamat.Text & "',Jurusan ='" & txtJurusan.Text & "',MatkulBind ='" & ab.Text & "',MatkulMatdas ='" & ac.Text & "',MatkulAgama ='" & ad.Text & "',MatkulKewarganegaraan ='" & ae.Text & "' WHERE NIK = '" & txtNIK.Text & "'"
        Dim MyConn As New OleDbConnection(strConn)
        Dim cmd As New OleDbCommand(MySQL, MyConn)
        MyConn.Open()
        cmd.ExecuteNonQuery()
        MyConn.Close()
        Response.Redirect("keren.aspx")
    End Sub
</script>

<style type="text/css">
BODY {
		scrollbar-3dlight-color:black;
		scrollbar-arrow-color:white;
		scrollbar-base-color:RGB(51,102,204);
		scrollbar-track-color:white;
		scrollbar-darkshadow-color:white;
		scrollbar-face-color:RGB(51,102,204);
		scrollbar-highlight-color:RGB(51,102,204);
		scrollbar-shadow-color:black	}
</style>

</head>

<body>
<form id="Form1" method="post" runat="server">
			<asp:label id="Label1" runat="server" font-names="Tahoma" font-size="X-Large">Program Data Mahasiswa</asp:label>
			<br><br>
			<asp:label id="Label2" runat="server" font-names="Tahoma" font-size="X-Small">Name:</asp:label>
			<asp:textbox id="txtName" runat="server" font-size="X-Small" font-names="Tahoma" width="174px" style="LEFT: 54px; POSITION: absolute;"></asp:textbox>
			<br><br>
			<asp:label id="Label3" runat="server" font-names="Tahoma" font-size="X-Small">NIK:</asp:label>
			<asp:textbox id="txtNIK" runat="server" font-size="X-Small" font-names="Tahoma" width="174px" style="LEFT: 54px; POSITION: absolute;"></asp:textbox>
			<br><br>
			<asp:label id="Label4" runat="server" font-names="Tahoma" font-size="X-Small">Alamat:</asp:label>
			<asp:textbox id="txtAlamat" runat="server" font-size="X-Small" font-names="Tahoma" width="174px" style="LEFT: 54px; POSITION: absolute;"></asp:textbox>
			<br><br>
			
		<asp:label id="Label5" runat="server" font-names="Tahoma" font-size="X-Small">Jurusan:</asp:label>
			
		<asp:DropDownList ID="txtJurusan" runat="server" style="LEFT: 54px; POSITION: absolute;">
    <asp:ListItem Enabled="true" Text="Elektro" Value="Elektro"></asp:ListItem>
    <asp:ListItem Text="Informatika" Value="Informatika"></asp:ListItem>
  
</asp:DropDownList>
	<br><br>
	
	<asp:label id="Label7" runat="server" font-names="Tahoma" font-size="X-Small">B.ind:</asp:label>
			
		<asp:DropDownList ID="ab" runat="server" style="LEFT: 54px; POSITION: absolute;">
    <asp:ListItem Enabled="true" Text="A" Value="A"></asp:ListItem>
    <asp:ListItem Text="B" Value="B"></asp:ListItem>
    <asp:ListItem Text="C" Value="C"></asp:ListItem>
      <asp:ListItem Text="D" Value="D"></asp:ListItem>
   <asp:ListItem Text="E " Value="E"></asp:ListItem>
  
</asp:DropDownList>
	<br><br>
	<asp:label id="Label8" runat="server" font-names="Tahoma" font-size="X-Small">Matdas:</asp:label>
			
		<asp:DropDownList ID="ac" runat="server" style="LEFT: 54px; POSITION: absolute;">
    <asp:ListItem Enabled="true" Text="A" Value="A"></asp:ListItem>
    <asp:ListItem Text="B" Value="B"></asp:ListItem>
    <asp:ListItem Text="C" Value="C"></asp:ListItem>
      <asp:ListItem Text="D" Value="D"></asp:ListItem>
   <asp:ListItem Text="E " Value="E"></asp:ListItem>
  
</asp:DropDownList>
	<br><br>
	<asp:label id="Label9" runat="server" font-names="Tahoma" font-size="X-Small">Agama:</asp:label>
			
		<asp:DropDownList ID="ad" runat="server" style="LEFT: 54px; POSITION: absolute;">
    <asp:ListItem Enabled="true" Text="A" Value="A"></asp:ListItem>
    <asp:ListItem Text="B" Value="B"></asp:ListItem>
    <asp:ListItem Text="C" Value="C"></asp:ListItem>
      <asp:ListItem Text="D" Value="D"></asp:ListItem>
   <asp:ListItem Text="E " Value="E"></asp:ListItem>
  
</asp:DropDownList>
	<br><br>
	<asp:label id="Label10" runat="server" font-names="Tahoma" font-size="X-Small">KWG:</asp:label>
			
		<asp:DropDownList ID="ae" runat="server" style="LEFT: 54px; POSITION: absolute;">
    <asp:ListItem Enabled="true" Text="A" Value="A"></asp:ListItem>
    <asp:ListItem Text="B" Value="B"></asp:ListItem>
    <asp:ListItem Text="C" Value="C"></asp:ListItem>
      <asp:ListItem Text="D" Value="D"></asp:ListItem>
   <asp:ListItem Text="E " Value="E"></asp:ListItem>
  
</asp:DropDownList>
	<br><br>
			<asp:button id="btnSend" onclick="OnBtnSendClicked" runat="server" font-names="Tahoma" text="Insert" backcolor="Navy" tabindex="5" font-size="X-Small" forecolor="White"></asp:button>
			<asp:button id="btnClear" onclick="OnBtnClearClicked" runat="server" text="Clear" backcolor="Navy" tabindex="6" font-names="Tahoma" font-size="X-Small" forecolor="White" style="LEFT: 50px; POSITION: absolute;"></asp:button>
			<asp:button id="btnUpdate" onclick="OnBtnUpdateClicked" runat="server" text="Update" backcolor="Navy" tabindex="6" font-names="Tahoma" font-size="X-Small" forecolor="White" style="LEFT: 90px; POSITION: absolute;"></asp:button>
			<asp:button id="btnDelete" onclick="OnBtnDeleteClicked" runat="server" text="Delete" backcolor="Navy" tabindex="6" font-names="Tahoma" font-size="X-Small" forecolor="White" style="LEFT: 140px; POSITION: absolute;"></asp:button>
		
		</form>
		<br>
<asp:label id="Label6" runat="server" font-size="X-Large" font-names="Tahoma">List mahasiswa</asp:label>
<br><br>

<asp:Repeater ID="rptGuestbook" Runat="Server">
	<HeaderTemplate>
		<table border=0 cellpadding=2    cellspacing=1 width=100%>
	</HeaderTemplate>

	<ItemTemplate>
		<tr bgcolor=RGB(51,102,204) height=1><td colspan="2"><font face=Tahoma size=1></font></td>
    </tr>
    <tr>		
			<td width=20%><b><font face=Tahoma size=2>Nama:</font></b></td> 
				<td><font face=Tahoma size=2><%#Container.DataItem("Name")%></td></tr>			
		<tr>
		<tr>		
			<td width=20%><b><font face=Tahoma size=2>NIK:</font></b></td> 
				<td><font face=Tahoma size=2><%#Container.DataItem("NIK")%></td></tr>		
		<tr>
			<td width=20%><b><font face=Tahoma size=2>Alamat:</td> 
				<td><font face=Tahoma size=2><%#Container.DataItem("Alamat")%></td></tr>				
		<tr>
			<td width=20%><b><font face=Tahoma size=2>Jurusan:</td> 
			
				<td><font face=Tahoma size=2><%#Container.DataItem("Jurusan")%></td></tr>
				<tr>
			<td width=20%><b><font face=Tahoma size=2></td> 
			
				<td><font face=Tahoma size=2>|KODE MATA KULIAH_|_NAMA MATA KULIAH_|_NILAI|</td></tr>
				<tr>
			<td width=20%><b><font face=Tahoma size=2></td> 
			
				<td><font face=Tahoma size=2>|UBHARA0000000001_|_BAHASA INDONESIA_|___<%#Container.DataItem("MatkulBind")%>__|</td></tr>	
				<tr>
			<td width=20%><b><font face=Tahoma size=2></td> 
			
				<td><font face=Tahoma size=2>|UBHARA0000000002_|_MATEMATIKA DASAR_|___<%#Container.DataItem("MatkulMatdas")%>__|</td></tr>		
	<tr>
			<td width=20%><b><font face=Tahoma size=2></td> 
			
				<td><font face=Tahoma size=2>|UBHARA0000000003_|_AGAMA____________|___<%#Container.DataItem("MatkulAgama")%>__|</td></tr>	
				<tr>
			<td width=20%><b><font face=Tahoma size=2></td> 
			
				<td><font face=Tahoma size=2>|UBHARA0000000004_|_KEWARGANEGARAAN_|___<%#Container.DataItem("MatkulKewarganegaraan")%>__|</td></tr>		
	
	</ItemTemplate>
		
	<SeparatorTemplate>
		<tr height=10><td></td></tr>
	</SeparatorTemplate>
	
	<FooterTemplate>
		</table>
	</FooterTemplate>

</asp:Repeater>
	

	
</body>
</html>
