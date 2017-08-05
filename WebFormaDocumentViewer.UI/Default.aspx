<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="WebFormsDocumentViewer.UI.Default" %>

<%@ Register Assembly="WebFormsDocumentViewer" Namespace="WebFormsDocumentViewer" TagPrefix="cc" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <cc:DocumentViewer runat="server" Width="500" Height="500" FilePath="~/sample.pdf" />
        </div>
    </form>
</body>
</html>
