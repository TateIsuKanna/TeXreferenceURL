Imports System.Net

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
                    Dim html_byte(10000 - 1) As Byte
                    Using sr As IO.Stream = DirectCast(WebRequest.Create(escapedURL), HttpWebRequest).GetResponse.GetResponseStream
                        sr.Read(html_byte, 0, 10000)
                    End Using
                    Dim html As String = System.Text.Encoding.UTF8.GetString(html_byte)

                    'HACK:エンコーディングをきちんと見るべき
                    If html.IndexOf("UTF-8", StringComparison.CurrentCultureIgnoreCase) = -1 Then
                        html = System.Text.Encoding.GetEncoding("Shift_JIS").GetString(html_byte)
                    End If

                    Dim title As String
                    Dim title_start As Integer = html.IndexOf("<title>", StringComparison.CurrentCultureIgnoreCase) + "<title>".Length
                    If title_start = -1 + "<title>".Length Then
                        title = ""
                    Else
                        Dim title_end As Integer = html.IndexOf("</title>", StringComparison.CurrentCultureIgnoreCase)
                        title = html.Substring(title_start, title_end - title_start)
                    End If
                    TextBox2.Text &= "\bibitem{}「" & title & "」\\" & URL & " 閲覧日" & Now.ToString("yyyy年M月d日") & vbCrLf
                Catch ex As Exception
                    TextBox2.Text &= vbCrLf & escapedURL & "について" & vbCrLf & ex.Message & vbCrLf
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