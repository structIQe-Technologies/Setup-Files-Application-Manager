Imports System.Diagnostics
Imports System.IO

Public Class OutlookMailHelper

    Public Shared Sub ComposeMailWithDefaultClient(toAddress As String, subject As String, body As String, attachmentFile As String)

        Try
            ' Encode the subject and body to make them URI-safe
            Dim subjectEncoded As String = Uri.EscapeDataString(subject)
            Dim bodyEncoded As String = Uri.EscapeDataString(body)

            ' Build the mailto URI
            Dim mailtoUri As String = $"mailto:{toAddress}?subject={subjectEncoded}&body={bodyEncoded}"

            ' Open default mail app
            Process.Start(New ProcessStartInfo(mailtoUri) With {
                .UseShellExecute = True
            })



        Catch ex As Exception
            MsgBox("Error preparing email: " & ex.Message, MsgBoxStyle.Critical, "App_Name")
        End Try

    End Sub

End Class
