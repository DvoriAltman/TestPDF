

<%@ Page Title="Home Page" Language="VB" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.vb" Inherits="TestPDF._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">


   
        <div class="col-md-4">
            <h2>Getting Started</h2>
                 
        </div>
                   <p>

                <button onclick="downloadPDF()">הורד PDF</button>
            </p>
  <script>
        function downloadPdf() {
            $.ajax({
                url: '/Detault/GeneratePdf', // נקודת הקצה ליצירת הדוח
                type: 'POST',
                xhrFields: {
                    responseType: 'blob' // מקבלים את התגובה כ-Blob
                },
                success: function (data, status, xhr) {
                    // יצירת URL להורדה
                    const blob = new Blob([data], { type: 'application/pdf' });
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.style.display = 'none';
                    a.href = url;
                    a.download = 'Report.pdf'; // שם הקובץ שיורד
                    document.body.appendChild(a);
                    a.click();
                    window.URL.revokeObjectURL(url);
                },
                error: function (xhr, status, error) {
                    console.error('Error:', error);
                    alert('שגיאה ביצירת הדוח');
                }
            });
        }
    </script>
<%--    <script>
           function downloadPDF() {
                $.ajax({
                    url: "/CreatePDF/CreatePDFReport", 
                     method: "GET",
                    xhrFields: {
                    responseType: "blob" // קבלת תגובה כ-BLOB
                },
                success: function (response) {
                 // יצירת URL מה-Blob
                    //console.log(The responce ${response})
                    //var blob = new Blob([response], { type: "application/octet-stream" });
                    //const url = window.URL.createObjectURL(blob);
                    //var link = document.createElement('a');
                    //link.href = URL.createObjectURL(response);
                    //link.download = "Report.pdf"; // שם הקובץ שיורד
                    //link.click();
                    //URL.revokeObjectURL(link.href);

                //// פתיחת ה-PDF בלשונית חדשה או הורדתו
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

    </script>--%>


</asp:Content>
