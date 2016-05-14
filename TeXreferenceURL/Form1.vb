Imports System.Net
Imports System.Text

Public Class Form1
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.Control AndAlso e.Shift AndAlso e.KeyCode = Keys.Enter Then
            TextBox2.Clear()
            For Each escapedURL In TextBox1.Lines
                Dim URL As String = System.Web.HttpUtility.UrlDecode(escapedURL)

                'エスケープする
                URL = URL.Replace("\", "\\") 'エスケープに使うから最初にする
                URL = URL.Replace("_", "\_")
                URL = URL.Replace("%", "\%")
                URL = URL.Replace("#", "\#")
                URL = URL.Replace("$", "\$")
                URL = URL.Replace("&", "\&")
                URL = URL.Replace("{", "\{")
                URL = URL.Replace("}", "\}")
                URL = URL.Replace("<", "\<")
                URL = URL.Replace(">", "\>")
                URL = URL.Replace("^", "\^")
                URL = URL.Replace("|", "\|")
                URL = URL.Replace("~", "\~")
				Try
					Dim html_byte() As Byte
					Dim wc As New WebClient
					html_byte = wc.DownloadData(escapedURL)
					Dim ContentType As String = wc.ResponseHeaders.Item(HttpResponseHeader.ContentType)
					Dim httpheader_charset_begin As Integer = ContentType.IndexOf("charset=")

					Dim html As String

					If httpheader_charset_begin > -1 Then
						httpheader_charset_begin += "charset=".Length
						Dim htmlenc As String = ContentType.Substring(httpheader_charset_begin, ContentType.Length - httpheader_charset_begin)
						html = Encoding.GetEncoding(htmlenc).GetString(html_byte)
					Else
						html = Encoding.ASCII.GetString(html_byte)

						Dim html_charset_begin As Integer = html.IndexOf("charset=")
						If html_charset_begin > -1 Then
							html_charset_begin += "charset=".Length
							Dim htmlenc As String = html.Substring(html_charset_begin, html.IndexOf("""", html_charset_begin) - html_charset_begin)
							html = Encoding.GetEncoding(htmlenc).GetString(html_byte)
						Else
							html = Encoding.UTF8.GetString(html_byte)
						End If
					End If

					Dim title As String
					Dim title_begin As Integer = html.IndexOf("<title>", StringComparison.CurrentCultureIgnoreCase)
					If title_begin = -1 Then
						title = ""
					Else
						title_begin += "<title>".Length
						Dim title_end As Integer = html.IndexOf("</title>", title_begin, StringComparison.CurrentCultureIgnoreCase)
						title = html.Substring(title_begin, title_end - title_begin)
					End If
					TextBox2.Text &= "\bibitem{}「" & title & "」\\" & URL & " 閲覧日" & Now.ToString("yyyy年M月d日") & vbCrLf
				Catch ex As Exception
					TextBox2.Text &= vbCrLf & escapedURL & "について" & vbCrLf & ex.ToString & vbCrLf
				End Try
            Next
            TextBox2.Focus()
            TextBox2.SelectAll()
        End If
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.Control AndAlso e.KeyCode = Keys.A Then
            TextBox2.SelectAll()
        End If
    End Sub
End Class