Imports Microsoft.VisualBasic
Imports iTextSharp.text
Imports System.IO
Imports iTextSharp.text.pdf
Imports System.Data
Imports System.Diagnostics
Imports System.Web.Services
Imports System.Configuration.ConfigurationSettings
'Imports System.Web.UI.DataVisualization.Charting
Imports iTextSharp.tool.xml
Imports iTextSharp.text.pdf.parser
Imports System.Resources
Imports System.Collections
Imports PdfSharp.Pdf
Imports PdfSharp.Drawing

Public Class CreatePDF

    <WebMethod()>
    Public Sub sendPdf()
        Dim pdfData As Byte() = CreatePDFReport()
        HttpContext.Current.Response.Clear()
        HttpContext.Current.Response.ContentType = "application/pdf"
        HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment; filename=Report.pdf")
        HttpContext.Current.Response.OutputStream.Write(pdfData, 0, pdfData.Length)
        HttpContext.Current.Response.BinaryWrite(pdfData)
        HttpContext.Current.Response.Flush()
        HttpContext.Current.Response.End()
    End Sub

    <WebMethod()>
    Public Sub CreatePDFReport()
        ' יצירת מסמך PDF חדש
        Dim doc As New PdfDocument()
        doc.info.Title = "הדוח שלי"

        ' יצירת דף חדש במסמך
        Dim page As PdfPage = doc.AddPage()

        ' יצירת גרפיקה לדף
        Dim gfx As XGraphics = XGraphics.FromPdfPage(page)

        ' הגדרת פונטים לציור
        Dim font As New XFont("Verdana", 20, XFontStyle.Bold)

        ' הוספת טקסט לדף
        gfx.DrawString("שלום עולם, זהו דוח PDF", font, XBrushes.Black, New XPoint(100, 100))

        ' שמירת ה-PDF
        doc.Save("Report.pdf")
    End Sub


    'Public Function CreatePDFReport() As Byte()
    '    Try
    '        ' יצירת PDF חדש בזיכרון
    '        Using memoryStream As New MemoryStream()
    '            Dim document As New Document(PageSize.A4, 50, 50, 25, 25)
    '            memoryStream.Flush()
    '            memoryStream.Position = 0

    '            ' הגדרת writer שיכתוב ל-MemoryStream
    '            Dim writer As PdfWriter = PdfWriter.GetInstance(document, memoryStream)

    '            ' פתיחת מסמך ה-PDF
    '            document.Open()
    '            'הגדרת הפונטים במסמך
    '            Dim fontPath As String = "C:\Windows\Fonts\Arial.ttf"
    '            Dim baseFont As BaseFont = baseFont.CreateFont(fontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED) 'לעדכן לאריאל או תומך עברית אחר
    '            Dim titleFont As Font = New Font(baseFont, 22, Font.BOLD, BaseColor.BLACK) 'פונט מודגש לכותרות 
    '            Dim customFont As Font = New Font(baseFont, 16, Font.NORMAL, BaseColor.BLACK)  'פונט רגיל לשאר המסמך
    '            ' הוספת תוכן למסמך
    '            document.Add(New PdfDate(Date.Now.ToString("dd/MM/yyyy")))   'תאריך של עכשיו
    '            document.Add(New Paragraph(Environment.NewLine))
    '            document.Add(New Paragraph("title", titleFont))
    '            document.Add(New Paragraph("iTextSharp", titleFont))
    '            document.Add(New Paragraph("text", customFont))

    '            ' סגירת המסמך
    '            document.Close()

    '            ' החזרת הדאטה כ-Byte Array
    '            Return memoryStream.ToArray()
    '        End Using
    '    Catch ex As Exception
    '        ' טיפול בשגיאות - תוכל להוסיף לוגים או לזרוק חריגה
    '        Throw New ApplicationException("Error generating PDF", ex)
    '    End Try
    'End Function







    'Public Function CreatePDFReport() As Byte()
    '    Dim ms As New MemoryStream()
    '    Dim userId As String = 12
    '    'להוסיף בדיקת שפת לקוח כדי להתאים את הדוח גנרית
    '    Dim ds As Data.DataSet
    '    Dim dat As New DataTable   'qCandy.qCandy(AppSettings("sDSN"))
    '    '' מיקום קובץ הפונט
    '    Dim fontPath As String = "C:\Windows\Fonts\Arial.ttf" ' לעדכן את הנתיב לפי המיקום הנכון ופונט רצוי מהתיקייה
    '    '' מיקום שמירת ה-PDF לעצמי לבדיקות תוצאה- אחכ להוריד
    '    Dim outputPath As String = "C:\DvoriSpace\Reports\ClientReport.pdf"

    '    '''הקצאת משתנים לוקליים
    '    'ds = dat.getManById(userId.ToString)
    '    'Dim firstName = ds.Tables(0).Rows(0)("first_name")    'שם המשתמש
    '    ''בדיקות לשפת הלקוח- להמשיך בדיקה בהתאם

    '    Dim firstName = "dvori"

    '    ''שם התאמה ותיאור ההתאמה 
    '    'ds = dat.getUserCompetenciesById(userId.ToString, "T", lang)
    '    Dim compName1 = ds.Tables(0).Rows(0)("compentencyName")
    '    Dim compDesc1 = ds.Tables(0).Rows(0)("compDescription")
    '    Dim compName2 = ds.Tables(0).Rows(1)("compentencyName")
    '    Dim compDesc2 = ds.Tables(0).Rows(1)("compDescription")
    '    Dim compName3 = ds.Tables(0).Rows(2)("compentencyName")
    '    Dim compDesc3 = ds.Tables(0).Rows(2)("compDescription")
    '    Dim compName4 = ds.Tables(0).Rows(3)("compentencyName")
    '    Dim compDesc4 = ds.Tables(0).Rows(3)("compDescription")
    '    Dim compName5 = ds.Tables(0).Rows(4)("compentencyName")
    '    Dim compDesc5 = ds.Tables(0).Rows(4)("compDescription")
    '    Dim compName6 = ds.Tables(0).Rows(5)("compentencyName")
    '    Dim compDesc6 = ds.Tables(0).Rows(5)("compDescription")
    '    Dim compName7 = ds.Tables(0).Rows(6)("compentencyName")
    '    Dim compDesc7 = ds.Tables(0).Rows(6)("compDescription")

    '    'ds = dat.GetMarkFitForMan(userId.ToString)
    '    Dim tbl = ds.Tables(0)
    '    Dim tbl1 = createDataTable("קטגוריה_System.String:תיאור_System.String:סהכ_System.String")
    '    Dim caption As String = "עסקים:ארגון:תרבות:אנשים:אמנות:חוץ:מדע:טכנולוגיה"

    '    'For x = 0 To caption.Split(":").Length - 1
    '    '    Dim tr As DataRow = tbl1.NewRow
    '    '    tr.Item(0) = caption.Split(":")(x)
    '    '    tbl1.Rows.Add(tr)
    '    'Next

    '    'For x = 0 To 7
    '    '    Dim tr As DataRow = tbl1.NewRow
    '    '    ' tr.Item(1) = tbl.Rows(0)("textMark").Split("I")(3).Split(",")(x - 1)
    '    '    tbl1.Rows(x)(2) = tbl.Rows(0)("textMark").Split("I")(3).Split(",")(x)
    '    '    ds = dat.GetFitFields(x + 1, lang)

    '    'Next

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
