

<%@ Page Title="Home Page" Language="VB" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.vb" Inherits="TestPDF._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">


   
        <div class="col-md-4">
            <h2>Getting Started</h2>
                 
        </div>
                   <p>

                <button onclick="downloadPDF()">הורד PDF</button>
            </p>
       
        <script>
            function downloadPDF() {
                $.ajax({
                    url: "http://localhost:63313/Default/sendPdf", 
                     method: "GET",
                    xhrFields: {
                    responseType: "blob" // קבלת תגובה כ-BLOB
                },
                success: function (data) {
                 // יצירת URL מה-Blob
                 const blob = new Blob([data], { type: "application/pdf" });
                 const url = window.URL.createObjectURL(blob);

                // פתיחת ה-PDF בלשונית חדשה או הורדתו
                 window.open(url, "_blank");
                const a = document.createElement("a");
                 a.href = url;
                 a.download = "Report.pdf"; // שם הקובץ להורדה
                 a.click();
                 window.URL.revokeObjectURL(url); // ניקוי ה-URL הזמני
            },
            error: function (xhr, status, error) {
                 console.error("Error generating PDF:", error);
            }
            });
            }

    </script>

</asp:Content>
