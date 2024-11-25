<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TestPDF.aspx.vb" Inherits="TestPDF.TestPDF" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <button id="testButton">TestPdf</button>
        </div>
    </form>
            <script>
        const tButton = document.getElementById("testButton");
        tButton.addEventListener("click", function generatePDFReport() {
            $.ajax({
                type: "POST",
                url: 'CreatePDF.vb/CreatePDFReport', 
                data: JSON.stringify({}),  
                contentType: "application/json; charset=utf-8",
                dataType: "json",
                success: function(response) {
                // הצלחה - ניתן להוסיף הודעה או לעבור לדף אחר
                alert("הדוח נוצר בהצלחה!");
                //window.location.href = 'YourPath/ClientReport.pdf';  
                },
                error: function(xhr, status, error) {
                // טיפול בשגיאה
                alert("שגיאה ביצירת הדוח: " + error);
                }
            });
        }       
    </script>
</body>
</html>
