'Imports Microsoft.Office.Interop.Outlook
'Imports System.Diagnostics
'Imports System.IO

'Public Class OutlookMailHelper

'    Public Shared Sub SendSmartEmail(toAddress As String, subject As String, body As String, attachmentPath As String)

'        Try
'            ' Try creating Outlook instance
'            Dim outlookApp As New Application()

'            Dim mail As MailItem = CType(outlookApp.CreateItem(OlItemType.olMailItem), MailItem)
'            mail.To = toAddress
'            mail.Subject = subject

'            If File.Exists(attachmentPath) Then
'                mail.Attachments.Add(attachmentPath)
'            End If

'            mail.Display()
'            Threading.Thread.Sleep(500)

'            ' Prepend your HTML body above Outlook's signature
'            Dim customBodyHtml As String = "<p style='margin:0;'>Dear Team,<br><br>" &
'                                           body.Trim().Replace(vbCrLf, "<br>") &
'                                           "<br><br>Regards,<br>Aditya</p>"

'            mail.HTMLBody = customBodyHtml & mail.HTMLBody

'        Catch ex As system.Exception
'            ' 📨 Fallback to mailto
'            Dim safeSubject As String = Uri.EscapeDataString(subject)
'            Dim safeBody As String = Uri.EscapeDataString(body & vbCrLf & vbCrLf & "(Please attach the license file manually.)")
'            Dim mailtoUri As String = $"mailto:{toAddress}?subject={safeSubject}&body={safeBody}"
'            MsgBox("hello")
'            Process.Start(New ProcessStartInfo(mailtoUri) With {.UseShellExecute = True})

'            ' 📁 Show the folder so user can attach the file
'            If File.Exists(attachmentPath) Then
'                Process.Start("explorer.exe", $"/select,""{attachmentPath}""")
'            End If
'        End Try

'    End Sub

'End ClassImports System.Diagnostics


Imports System.IO
Imports System.Diagnostics
Imports Microsoft.Office.Interop.Outlook

Public Class OutlookMailHelper

    Public Shared Sub SmartEmail(toAddress As String, subject As String, body As String, attachmentPath As String)

        If IsOutlookInstalled() Then
            Try
                MsgBox("Please ensure Outlook is closed", MsgBoxStyle.Exclamation, "Outlook Warning")
                ' Try using Outlook Interop
                Dim outlookApp As New Application
                Dim mail As MailItem = CType(outlookApp.CreateItem(OlItemType.olMailItem), MailItem)

                mail.To = toAddress
                mail.Subject = subject

                If File.Exists(attachmentPath) Then
                    mail.Attachments.Add(attachmentPath)
                End If


                Dim outlookPath As String = GetOutlookPath()
                If Not String.IsNullOrEmpty(outlookPath) AndAlso File.Exists(outlookPath) Then
                    Process.Start(outlookPath) ' Launch Outlook if not already open
                    Threading.Thread.Sleep(1000) ' Optional: give it a moment to launch
                End If
                mail.Display()
                Threading.Thread.Sleep(500)
                Dim cleanBody As String = body.Trim().
                            Replace(vbCrLf, "<br>").
                            Replace(vbLf, "<br>").
                            Replace("\n", "<br>")
                '  Dim customBodyHtml As String =                body.Trim().Replace(vbCrLf, "\n")

                Dim customBodyHtml As String = "<div style='font-family:Segoe UI, sans-serif; font-size:11pt;'>" &
                                cleanBody &
                                "</div>"


                mail.HTMLBody = customBodyHtml & mail.HTMLBody
                Return

            Catch ex As System.Exception
                ' Outlook failed – fall back to mailto
            End Try



        End If

        ' 📨 Fallback to default mail client
        Try
            Dim safeSubject As String = Uri.EscapeDataString(subject)
            Dim safeBody As String = Uri.EscapeDataString(body & vbCrLf & vbCrLf & "*** IMPORTANT: Please attach the license file (.c2v) manually from your download folder. ***")
            Dim mailtoUri As String = $"mailto:{toAddress}?subject={safeSubject}&body={safeBody}"

            Process.Start(New ProcessStartInfo(mailtoUri) With {.UseShellExecute = True})

            If File.Exists(attachmentPath) Then
                Process.Start("explorer.exe", $"/select,""{attachmentPath}""")
            End If

        Catch ex2 As System.Exception
            MsgBox("Failed to launch email client: " & ex2.Message, MsgBoxStyle.Critical, "Email Error")
        End Try

    End Sub

    Private Shared Function IsOutlookInstalled() As Boolean
        Try
            Dim outlookPath As String = GetOutlookPath()
            Return Not String.IsNullOrEmpty(outlookPath) AndAlso File.Exists(outlookPath)
        Catch
            Return False
        End Try
    End Function

    Private Shared Function GetOutlookPath() As String
        ' Try both 64-bit and 32-bit registry keys
        Dim outlookKeyPaths As String() = {
            "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE",
            "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE"
        }
        For Each keyPath In outlookKeyPaths
            Dim path As String = CStr(Microsoft.Win32.Registry.GetValue(keyPath, "", Nothing))
            If Not String.IsNullOrEmpty(path) Then
                Return path
            End If
        Next

        Return String.Empty
    End Function



End Class
