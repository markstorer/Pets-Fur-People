<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="FeaturedPet.aspx.vb" Inherits="UglyMoneyAuction.FeaturedPet" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:Table ID="FeaturedPetTable" runat="server" >
            <asp:TableRow class="" ID="FeaturedPetRow" runat="server">
                <asp:TableCell class="" ID="FeaturedPetCell" runat="server" HorizontalAlign="Left" BorderStyle="Solid" BorderWidth="1" BorderColor="Blue">
                </asp:TableCell>
            </asp:TableRow>
        </asp:Table>
    </div>
    </form>
</body>
</html>
