

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
                url: '/Default/GeneratePdf', // נקודת הקצה ליצירת הדוח
                type: 'GET',
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


</asp:Content>
