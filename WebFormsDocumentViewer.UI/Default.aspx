<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="WebFormsDocumentViewer.UI.Default" %>

<%@ Register Assembly="WebFormsDocumentViewer" Namespace="WebFormsDocumentViewer" TagPrefix="cc" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Document Viewer</title>
    <link rel="stylesheet" href="Content/bootstrap.min.css" />
    <link rel="stylesheet" href="Content/Site.css" />
</head>
<body>
    <form id="form1" runat="server">
        <div class="jumbotron text-center">
            <h2>Document Viewer custom control for Web Forms</h2>
            <p>
                This currently supports PDF and Word documents. For more details check the
                <a href="https://github.com/HajbokRobert/WebFormsDocumentViewer" target="_blank">Github repository</a>.
            </p>
        </div>
        <div class="container">
            <div class="row">
                <div class="col col-md-6">
                    <div class="card text-center">
                        <div class="card-block">
                            <h4 class="card-title">PDF Viewer</h4>
                            <p class="card-text">This renders a PDF document in an iframe.</p>
                            <cc:DocumentViewer runat="server" Width="500" Height="500" FilePath="~/Samples/sample.pdf" />
                        </div>
                    </div>
                </div>
                <div class="col col-md-6">
                    <div class="card text-center">
                        <div class="card-block">
                            <h4 class="card-title">Word Viewer</h4>
                            <p class="card-text">
                                This converts the Word document to PDF and renders it in an iframe. 
                                You should have Microsoft Office installed for this.
                            </p>
                            <cc:DocumentViewer runat="server" Width="500" Height="500" FilePath="~/Samples/sample.docx" TempDirectoryPath="~/TempFiles" />
                        </div>
                    </div>
                </div>
            </div>
            <br />
            <div class="row">                
                <div class="col col-md-6">
                    <div class="card text-center">
                        <div class="card-block">
                            <h4 class="card-title">PowerPoint Viewer</h4>
                            <p class="card-text">
                                This converts the PowerPoint document to PDF and renders it in an iframe. 
                                You should have Microsoft Office installed for this.
                            </p>
                            <cc:DocumentViewer runat="server" Width="500" Height="500" FilePath="~/Samples/sample.pptx" TempDirectoryPath="~/TempFiles" />
                        </div>
                    </div>
                </div>
                <div class="col col-md-6">
                    <div class="card text-center">
                        <div class="card-block">
                            <h4 class="card-title">Excel Viewer</h4>
                            <p class="card-text">
                                This converts the Excel document to HTML and renders it in an iframe. 
                                You should have Microsoft Office installed for this.
                            </p>
                            <cc:DocumentViewer runat="server" Width="500" Height="500" FilePath="~/Samples/sample.xlsx" TempDirectoryPath="~/TempFiles" />
                        </div>
                    </div>
                </div>
            </div>
        </div>
        <br />
    </form>
</body>
</html>
