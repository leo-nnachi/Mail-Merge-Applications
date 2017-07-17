<%@ Page Title="Home Page" Language="C#" AutoEventWireup="true" CodeBehind="Home.aspx.cs"
    Inherits="MailMerger.Home" MasterPageFile="~/Site.master" %>

<asp:Content ID="HeaderContent" runat="server" ContentPlaceHolderID="HeadContent">
    <script type="text/javascript">
           
        function MergeMail() {
            var dataDoc = document.getElementById("txtDataDoc");
            var mergeDoc = document.getElementById("txtMergeDoc");
            var email = document.getElementById("email");
            var queryString = "format=" + mergeDoc.value + '&email=' + email.value + "&source=" + dataDoc.value + "&file_type=Word document&split=No&prefix=&field1=&field2=&field3=&field4=&field5=&suffix=&delimiter=";

            //window.location = "MailMerge.aspx?" + queryString;

            var win = window.open('MailMerge.aspx?' + queryString, 'Download', 'width=350,height=500,resizable=no,scrollbars=no,toolbar=no,location=no,status=no');

    
            //            //http://localhost/MailMerger/docs/penmast.doc
            //            //"MailMerge.aspx?format=C:\\docs\\benefitstatement_ibm_language.doc&source=C:\\docs\\penmast.doc";
            //            //format=http://192.168.3.60/MailMerger/MailMergeDocs/benefitstatement_ibm_language.doc
            //            //&source=http://localhost/MailMerger/MailMergeDocs/penmast.doc";
        }

        function CreateTemplate() {

            var filePath = document.getElementById("txtDocumentPath");
            var queryString = "path=" + filePath.value; // + "&source=" + dataDoc.value;
            
            window.location = "CreateTemplate.aspx?" + queryString;

        }


        function MergeMailApp() {
            var source1 = document.getElementById("source1");
            var format1 = document.getElementById("format1");
            var email1 = document.getElementById("email1");
            var file_type1 = document.getElementById("file_type1");
            var split1 = document.getElementById("split1");
            var queryString = "format=" + format1.value + "&email=" + email1.value + "&source=" + source1.value + "&file_type=" + file_type1.value + "&split=" + split1.value + "&prefix=&field1=pm_forename&field2=&field3=&field4=&field5=&suffix=&delimiter=" + "&database=abbott8";

            var win = window.open('MailMerge.aspx?' + queryString, 'Download', 'width=350,height=500,resizable=no,scrollbars=no,toolbar=no,location=no,status=no');


        }
          
    </script>
    <style type="text/css">
        .InputBox
        {
            width: 300px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" runat="server" ContentPlaceHolderID="MainContent">
    <h2>
    </h2>
    <p>
        Here we will be merging the datasource into the Mail format.
    </p>
    <table style="width: 610px">
        <tr>
            <td>
                <asp:Label ID="Label2" runat="server" Text="Email Address: "></asp:Label>
            </td>
            <td>
                <input id="email" class="InputBox" type="text" value="imran.qaiser@itsbettertogether.co.uk" /><%--rachael.yates@lpsystems.com,imran.qaiser@itsbettertogether.co.uk--%>
            </td>
        </tr>
        <tr>
            <td>
                <asp:Label ID="lblMergeDoc" runat="server" Text="Template Document: "></asp:Label>
            </td>
            <td>
                <input id="txtMergeDoc" type="text" class="InputBox"
                    value='<%= Server.MapPath("~/MailMergeDocs/1.doc") %>' />
                <%--northerntruststatement,withsorp--%>
            </td>
        </tr>
        <tr>
            <td>
                <asp:Label ID="lblDataDoc" runat="server" Text="Data Document: "></asp:Label>
            </td>
            <td>
                <input id="txtDataDoc" class="InputBox" type="text" value='<%= Server.MapPath("~/MailMergeDocs/1.txt") %>'
                    />
                <%--penmast_200,datasource_copy--%>
            </td>
        </tr>
    </table>
    <p>
        Click here to merge and download
        <input type="button" class="login" title=" Mail Merge " value=" Mail Merge " onclick="javascript:MergeMail();" />
    </p>
    <hr />
    <p>
        Here we will be merging the datasource into the Mail format.
    </p>
    <table>
        <tr>
            <td>
                <asp:Label ID="Label1" runat="server" Text="Template Document: "></asp:Label>
            </td>
            <td>
                <input id="txtDocumentPath" class="InputBox"  type="text" value='<%= Server.MapPath("~/MailMergeDocs/abc_-2013.doc") %>' />
            </td>
        </tr>
    </table>
    <p>
        Click here to merge and download
        <input type="button" class="login" title=" Create Template " value=" Create Template "
            onclick="javascript:CreateTemplate();" />
    </p>


      <p>
        Here we will be merging the datasource into the Mail format using custom tests.
    </p>
       <table style="width: 610px">
        <tr>
            <td>
                <asp:Label ID="Label3" runat="server" Text="Email Address: "></asp:Label>
            </td>
            <td>
                <input id="email1" class="InputBox" type="text" value="imran.qaiser@itsbettertogether.co.uk" />
            </td>
        </tr>
        <tr>
            <td>
                <asp:Label ID="Label4" runat="server" Text="Template Document: "></asp:Label>
            </td>
            <td>
                <input id="format1" type="text" class="InputBox"
                    value='<%= Server.MapPath("~/MailMergeDocs/1.docx") %>' />
              
            </td>
        </tr>
        <tr>
            <td>
                <asp:Label ID="Label5" runat="server" Text="Data Document: "></asp:Label>
            </td>
            <td>
                <input id="source1" class="InputBox" type="text" value='<%= Server.MapPath("~/MailMergeDocs/1.txt") %>'
                    />
              
            </td>
        </tr>
        <tr>
            <td>
                <asp:Label ID="Label6" runat="server" Text="File Type: "></asp:Label>
            </td>
            <td>
                <select id="file_type1" class="InputBox">
                    <option>Word document</option>
                    <option>PDF document</option>
                    <option> </option>
                </select>
               
              
            </td>
        </tr>
        <tr>
            <td>
                <asp:Label ID="Label7" runat="server" Text="Split File: "></asp:Label>
            </td>
            <td>
             <select id="split1" class="InputBox">
                    <option>No</option>
                    <option>Yes</option>
                    <option> </option>
                </select>
              
              
            </td>
        </tr>
    </table>

     
    <p>
        Click here to merge and download
        <input type="button" class="login" title=" Mail Merge " value=" Mail Merge " onclick="javascript:MergeMailApp();" />
        <asp:Button ID="Button1" runat="server" onclick="Button1_Click" Text="Check File Copy" />
    </p>
    <hr />
</asp:Content>
