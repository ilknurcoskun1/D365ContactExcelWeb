<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="CRMContact.aspx.cs" Inherits="D365ContactExcelWep.CRMContact" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
       <asp:FileUpload ID="fileUpload" runat="server" />
        <asp:Button ID="uploadButton" runat="server" Text="Upload File" OnClick="UploadButton_Click" />
        <br />
        <asp:Label ID="resultLabel" runat="server" Text=""></asp:Label>
        <br />
        <asp:GridView ID="contactGridView" runat="server" AutoGenerateColumns="false">
            <Columns>
                <asp:BoundField DataField="firstname" HeaderText="First Name" />
                <asp:BoundField DataField="lastname" HeaderText="Last Name" />
                <asp:BoundField DataField="emailaddress1" HeaderText="Email" />
                <asp:BoundField DataField="createdon" HeaderText="Creation Date" DataFormatString="{0:yyyy-MM-dd}" />
            </Columns>
        </asp:GridView>
    </form>
</body>
</html>
