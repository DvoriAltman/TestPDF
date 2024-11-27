Imports Microsoft.VisualBasic
'Imports Aspose.Pdf
Imports iText.Kernel.Font
Imports iText.Kernel.Pdf
Imports iText.Layout
Imports iText.Layout.Element
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Data
Imports System.Diagnostics
Imports System.Web.Services
Imports System.Configuration.ConfigurationSettings
'Imports System.Web.UI.DataVisualization.Charting
Imports System.Resources
Imports System.Collections
'Imports Aspose.Pdf.Text

Public Class CreatePDF


    <WebMethod()>
    Public Sub GeneratePdf()
        '        'יצירת PDF חדש בזיכרון
        Using MemoryStream As New MemoryStream()
            Using pdfWriter As New PdfWriter(MemoryStream)

                Dim pdfDoc As New PdfDocument(pdfWriter)
                Dim document As New Document(pdfDoc)

                'Dim fontPath As String = "C:\Windows\Fonts\Arial.ttf"
                document.Add(New Paragraph("Title"))
                document.Add(New Paragraph("Subtitle"))
                Dim title As New Paragraph("title")
                Dim subtitle As New Paragraph("iText 7")
                document.Add(title)
                document.Add(subtitle)

                ' ובדיקת הקונסול לוג עם גודל המידע לראות אם נשמר באמת סגירת המסמך
                document.Close()
                MemoryStream.Close()
                Console.WriteLine("PDF size: " & MemoryStream.Length)

                Console.WriteLine("Sending PDF to client")
            End Using

            HttpContext.Current.Response.Clear()
            HttpContext.Current.Response.ContentType = "application/pdf"
            HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment; filename=Report.pdf")
            HttpContext.Current.Response.BinaryWrite(MemoryStream.ToArray())  ' שליחת קובץ ה-PDF
            HttpContext.Current.Response.Flush()
            HttpContext.Current.Response.Close()
            HttpContext.Current.Response.End()
        End Using

    End Sub




    '       


    '    Catch ex As Exception
    '        Throw New ApplicationException("Error generating PDF", ex)
    '    End Try
    'End Sub

    'Dim pdfDocument As New Aspose.Pdf.Document()
    'Dim text As New TextFragment("שלום הנה דוח בסיעתא דשמיא")
    'Dim page As Page
    'page = pdfDocument.Pages.Add()
    'page.Paragraphs.Add(text)

    'Using ms As New MemoryStream()
    'pdfDocument.Save()
    'Using fileStream As New FileStream(templatePath, FileMode.Open, FileAccess.Read)
    '    fileStream.CopyTo(memoryStream)






    'Public Function CreatePDFReport() As Byte()
    '    Dim ms As New MemoryStream()
    '    Dim userId As String = 12

    '    Dim firstName = "dvori"


    '    ' יצירת המסמך
    '    Dim pdfDoc As New Document(PageSize.A4, 50, 50, 50, 50) 'הגדרת גודל עמוד ורוחב שוליים לפי פיקסלים

    '    Try
    '        Dim writer As PdfWriter = PdfWriter.GetInstance(pdfDoc, ms)
    '        ' של מספור עמודים אוטומטי  Page Event הוספת  

    '        'writer.PageEvent = New PageNumber()

    '        'writer.PageEvent = New PageNumber()


    '        'הוספת שילוב של דוח פידיאף קיים
    '        Dim existingPdfPath As String = "C:\DvoriSpace\Reports\oldReport.pdf"   ' הקובץ המקורי

    '        '   HTML פתיחת המסמך והגדרת שימוש בתגיות
    '        pdfDoc.Open()

    '        Dim reader As New PdfReader(existingPdfPath) 'הוספת קורא לקובץ פידיאף שלמעלה
    '        ' הוספת עמודים ספציפיים מהקובץ 
    '        Dim totalPages As Integer = reader.NumberOfPages
    '        Dim cb As PdfContentByte = writer.DirectContent

    '        'הגדרת הפונטים במסמך
    '        Dim baseFont As BaseFont = BaseFont.CreateFont(fontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED) 'לעדכן לאריאל או תומך עברית אחר
    '        Dim titleFont As Font = New Font(baseFont, 22, Font.BOLD, BaseColor.BLACK) 'פונט מודגש לכותרות 
    '        Dim customFont As Font = New Font(baseFont, 16, Font.NORMAL, BaseColor.BLACK)  'פונט רגיל לשאר המסמך

    '        'יצירת סימניות בדוח
    '        Dim bookmarks As New List(Of Dictionary(Of String, Object))
    '        ' סימניה ראשית
    '        Dim mainBookmark As New Dictionary(Of String, Object)
    '        mainBookmark("Title") = "mainBookMark" 'Resources.ResourceForPDF.mainBookmark
    '        mainBookmark("Action") = "GoTo"
    '        mainBookmark("Page") = "1 Fit" ' עמוד 1
    '        ' סימניה לפרק 1
    '        Dim chapter1Bookmark As New Dictionary(Of String, Object)
    '        chapter1Bookmark("Title") = "1bookmark" 'Resources.ResourceForPDF.chapter1Bookmark
    '        chapter1Bookmark("Action") = "GoTo"
    '        chapter1Bookmark("Page") = "2 Fit" ' עמוד 2
    '        ' סימניה לפרק 2
    '        Dim chapter2Bookmark As New Dictionary(Of String, Object)
    '        chapter2Bookmark("Title") = "2bookmark" 'Resources.ResourceForPDF.chapter2Bookmark
    '        chapter2Bookmark("Action") = "GoTo"
    '        chapter2Bookmark("Page") = "3 Fit" ' עמוד 3
    '        Dim chapter3Bookmark As New Dictionary(Of String, Object)
    '        chapter3Bookmark("Title") = "3bookmark" 'Resources.ResourceForPDF.chapter3Bookmark
    '        chapter3Bookmark("Action") = "GoTo"
    '        chapter3Bookmark("Page") = "4 Fit" ' עמוד 4
    '        Dim chapter4Bookmark As New Dictionary(Of String, Object)
    '        chapter4Bookmark("Title") = "4bookmark" 'Resources.ResourceForPDF.chapter4Bookmark
    '        chapter4Bookmark("Action") = "GoTo"
    '        chapter4Bookmark("Page") = "5 Fit" ' עמוד 5
    '        ' הוספת הסימניות לרשימה
    '        bookmarks.Add(mainBookmark)
    '        bookmarks.Add(chapter1Bookmark)
    '        bookmarks.Add(chapter2Bookmark)
    '        bookmarks.Add(chapter3Bookmark)
    '        bookmarks.Add(chapter4Bookmark)
    '        ' הצמדת הסימניות למסמך
    '        writer.Outlines = bookmarks

    '        'page 1
    '        ' -תוכן עניינים דינאמי עם קישורים לעמודים במסמך לעדכן מספרי עמודים בסוף
    '        pdfDoc.Add(New PdfDate(Date.Now.ToString("dd/MM/yyyy")))   'תאריך של עכשיו
    '        pdfDoc.Add(New Paragraph(Environment.NewLine))
    '        pdfDoc.Add(New Paragraph("title", titleFont)) 'Resources.ResourceForPDF.chavat & firstName, titleFont)) 
    '        pdfDoc.Add(New Paragraph(Environment.NewLine))
    '        Dim list As New List(List.UNORDERED) 'הוספת רשימה עם נקודות
    '        list.Add(New ListItem("xx")) 'Resources.ResourceForPDF.list1))
    '        Dim anchor1 As New Anchor(Resources.ResourceForPDF.kishur)
    '        anchor1.Reference = "#Page2"   '  יעד
    '        pdfDoc.Add(anchor1)   'הוספת הקישור לבסוף
    '        list.Add(New ListItem(Resources.ResourceForPDF.list2))
    '        Dim anchor2 As New Anchor(Resources.ResourceForPDF.kishur)
    '        anchor2.Reference = "#Page3"
    '        pdfDoc.Add(anchor2)
    '        list.Add(New ListItem(Resources.ResourceForPDF.list3))
    '        Dim anchor3 As New Anchor(Resources.ResourceForPDF.kishur)
    '        anchor3.Reference = "#Page4"
    '        pdfDoc.Add(anchor3)
    '        pdfDoc.NewPage()
    '        list.Add(New ListItem(Resources.ResourceForPDF.list4))
    '        Dim anchor4 As New Anchor(Resources.ResourceForPDF.kishur)
    '        anchor3.Reference = "#Page5"
    '        pdfDoc.Add(anchor3)
    '        pdfDoc.NewPage()
    '        list.Add(New ListItem(Resources.ResourceForPDF.list5))
    '        Dim anchor5 As New Anchor(Resources.ResourceForPDF.kishur)
    '        anchor3.Reference = "#Page5"
    '        pdfDoc.Add(anchor3)
    '        'pdfDoc.Add(New Paragraph(Resources.Resource.footerLine, titleFont))
    '        pdfDoc.NewPage()
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''page 2
    '' ייבוא עמוד 2 מהקובץ הקיים והוספתו בעמדה זו
    ''Dim page2 As PdfImportedPage = writer.GetImportedPage(reader, 2)
    ''cb.AddTemplate(page2, 0, 0)
    ''או הוספת התוכן החדש מהסורסס
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page2Header, titleFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.mabadoch, titleFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.echhechanti, titleFont))
    'pdfDoc.Add(New Paragraph(Resources.Resource.footerLine, titleFont))
    'pdfDoc.NewPage()

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''page 3
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.perekishiyut & " " & firstName, titleFont))
    'pdfDoc.Add(New Paragraph(Environment.NewLine))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.perekishiyut2, customFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.perekishiyut1, customFont))

    ''כאן להוסיף את הגרף האופקי עם אחוזי טיפוס ומתחת את הטיפוס
    ''את הגרף עכביש להעביר לעמוד נפרד  עם פירוט תוכן והגדלת הגרף
    'pdfDoc.Add(New Paragraph(Resources.Resource.footerLine, titleFont))
    'pdfDoc.NewPage()
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''page 4
    '' ייבוא עמוד 4 מהקובץ הקיים והוספתו בעמדה זו
    ''Dim page4 As PdfImportedPage = writer.GetImportedPage(reader, 4)
    ''cb.AddTemplate(page4, 0, 0)
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page4header, titleFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page4cont1, customFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page4tipusOmanuti1, customFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page4tipusOmanuti2, customFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page4tipusAnashim1, customFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page4tipusAnashim2, customFont))
    'pdfDoc.Add(New Paragraph(Resources.Resource.footerLine, titleFont))
    'pdfDoc.NewPage()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''page 5 'להוסיף מס אחוזים
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page5Hanetiyot & " " & firstName, titleFont))
    'pdfDoc.Add(New Paragraph(firstName & " " & Resources.ResourceForPDF.page5header, customFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page5sadot1 & "%", customFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page5Yrgun, customFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page5sadot2 & "%", customFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page5Tech, customFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page5sadot3 & "%", customFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page5Asakim, customFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page5sadot4 & "%", customFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page5Omanut, customFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page5sadot5 & "%", customFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page5Tarbut, customFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page5sadot6 & "%", customFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page5Anashim, customFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page5sadot7 & "%", customFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page5Mada, customFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page5sadot8 & "%", customFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page5Chutz, customFont))
    ''Dim page5 As PdfImportedPage = writer.GetImportedPage(reader, 5) ' להוריד את זה?
    'pdfDoc.Add(New Paragraph(Resources.Resource.footerLine, titleFont))
    'pdfDoc.NewPage()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''page 6  התחלת פרק הלימודים
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.perekHalimudim, titleFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page6header, customFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page6cont1 & " " & firstName, customFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page6cont2, customFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page6cont3, customFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page6cont4beforeGraph, customFont))
    'pdfDoc.Add(New Paragraph(Environment.NewLine))
    '''פה הגרף
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page6Ramatlimud1, customFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page6Ramatlimudperut1, customFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page6Ramatlimud2, customFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page6Ramatlimudperut2, customFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page6Ramatlimud3, customFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page6Ramatlimudperut3, customFont))
    'pdfDoc.Add(New Paragraph(Resources.Resource.footerLine, titleFont))
    'pdfDoc.NewPage()
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''page 7
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page7hemshechLimudim, titleFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page7Hesber, customFont))
    ''כאן להכניס את הכותרות מהריסורס לתוך טבלה אליה שולפים את נתוני הלקוח
    ''המלצות בטבלה לפי מוסד,מסלול,סוג ואחוז התאמה
    'pdfDoc.Add(New Paragraph(Environment.NewLine))
    'pdfDoc.Add(New Paragraph(Resources.Resource.footerLine, titleFont))
    'pdfDoc.NewPage()
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''page 8 
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page7hemshechLimudim, titleFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page8Header1, customFont))
    '''טבלת המלצות יועץ המערכת 
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page8Header2, customFont))
    'pdfDoc.Add(New Paragraph(Resources.Resource.footerLine, titleFont))
    ''טבלת הבחירות האישיות של הלקוח
    'pdfDoc.NewPage()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''Page 9
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page7hemshechLimudim, titleFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page9Header, customFont))
    '''טבלת המלצות המערכת 
    'pdfDoc.Add(New Paragraph(Resources.Resource.footerLine, titleFont))
    'pdfDoc.NewPage()
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''page 10 התחלת פרק התעסוקה
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.perekHataasuka, titleFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.perekHataasukaCont, customFont))
    ''  של היועץ- טבלת המלצת למקצועות לפי מקצוע, איפה לומדים ומסלול לימודים נדרש
    'pdfDoc.Add(New Paragraph(Resources.Resource.footerLine, titleFont))
    'pdfDoc.NewPage()
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''page 11 
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page11HemshechTaasuka, titleFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page11Hesber, titleFont))
    ''טבלת המלצות יועץ המערכת
    'pdfDoc.Add(New Paragraph(Resources.Resource.footerLine, titleFont))
    'pdfDoc.NewPage()
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''page 12
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page11HemshechTaasuka, titleFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page12Hesber, titleFont))
    ''טבלה בחירות הלקוח לתעסוקה
    'pdfDoc.Add(New Paragraph(Resources.Resource.footerLine, titleFont))
    'pdfDoc.NewPage()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''page 13
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page11HemshechTaasuka, titleFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.page9Header, titleFont))
    ''טבלה עד עשר בחירות המערכת 
    'pdfDoc.Add(New Paragraph(Resources.Resource.footerLine, titleFont))
    'pdfDoc.NewPage()
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''Dim page12 As PdfImportedPage = writer.GetImportedPage(reader, 12) 'ייבוא עמוד 12 
    ''cb.AddTemplate(page12, 0, 0)    '  בעיה רק עם המלל של סוף המסמך, התאמה לפי שם לקוח
    ''cb.BeginText()
    ''cb.SetTextMatrix(50, 750) ' מיקום השכבה בעמוד
    ''cb.ShowText("טקסט דינמי נוסף לעמוד: {userName}")
    ''cb.EndText()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''page 14 סיכום צידה לדרך
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.sicum, titleFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.sicumCont1, customFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.sicumCont2, customFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.sicumCont3, customFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.sicumCont4, customFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.sicumCont5, customFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.sicumCont6, customFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.sicumCont7, customFont))
    'pdfDoc.Add(New Paragraph(Resources.ResourceForPDF.sicumCont8, customFont))
    'pdfDoc.Add(New Paragraph(Resources.Resource.footerLine, titleFont))
    ' סגירת המסמך
    '        pdfDoc.Close()
    '        Console.WriteLine("הדוח נוצר בהצלחה ב-" & outputPath)
    '        Process.Start(outputPath)

    '    Catch ex As Exception
    '        Console.WriteLine("שגיאה ביצירת הדוח: " & ex.Message)
    '    Finally

    '        If pdfDoc.IsOpen Then
    '            pdfDoc.Close()
    '        End If

    '    End Try


    '    Return ms.ToArray()
    'End Function



End Class
