Public Class _Default
    Inherits Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
        Dim cpdf As CreatePDF = New CreatePDF()
        cpdf.GeneratePdf()
    End Sub
End Class