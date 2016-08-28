'正規表現全部見直してもいいかな












Imports System.Net
Imports System.Text
Imports System.Text.RegularExpressions

Public Class Form1
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.Control AndAlso e.KeyCode = Keys.A Then
            TextBox2.SelectAll()
        End If
    End Sub

    Private Sub ListBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ListBox1.KeyDown
        If e.Control AndAlso e.KeyCode = Keys.V Then
            For Each url In Split(Clipboard.GetText, vbCrLf)
                ListBox1.Items.Add(url.Trim)
            Next
        End If

        If e.Control AndAlso e.KeyCode = Keys.A Then
            Dim url As String = InputBox("URLか書名を入力して下さい．")
            If url <> "" Then
                ListBox1.Items.Add(url)
            End If
        End If

        If ListBox1.SelectedIndex > -1 Then
            If e.KeyCode = Keys.Delete Then
                ListBox1.Items.RemoveAt(ListBox1.SelectedIndex)
            End If
        End If

        If e.Control AndAlso e.Shift AndAlso e.KeyCode = Keys.Enter Then
            TextBox2.Clear()
            TextBox2.BackColor = SystemColors.Control
            TextBox2.AppendText("\begin{thebibliography}{" & CStr(10 ^ (Math.Floor(Math.Log10(ListBox1.Items.Count)) + 1) - 1) & "}" & vbCrLf)
            For Each urlescapedURL As String In ListBox1.Items
                If RegularExpressions.Regex.IsMatch(urlescapedURL, "^https?://", RegularExpressions.RegexOptions.IgnoreCase) Then
                    Try
                        Dim html_byte() As Byte
                        Dim wc As New WebClient
                        wc.Proxy = Nothing '要らないかもね
                        html_byte = wc.DownloadData(urlescapedURL)
                        Dim ContentType As String = wc.ResponseHeaders.Item(HttpResponseHeader.ContentType)
                        wc.Dispose()
                        Dim httpheader_charset As String = RegularExpressions.Regex.Match(ContentType, "(?<=charset=).+", RegularExpressions.RegexOptions.IgnoreCase).Value

                        Dim html As String
                        If httpheader_charset <> "" Then
                            html = Encoding.GetEncoding(httpheader_charset).GetString(html_byte)
                        Else
                            html = Encoding.ASCII.GetString(html_byte)
                            Dim charset As String = RegularExpressions.Regex.Match(html, "(?<=charset=""?)[^""]+(?="")", RegularExpressions.RegexOptions.IgnoreCase).Value
                            If charset <> "" Then
                                html = Encoding.GetEncoding(charset).GetString(html_byte)
                            Else
                                html = Encoding.UTF8.GetString(html_byte)
                            End If
                        End If

                        'HACK:html全体をdecodeしてもいいかも
                        Dim title As String = Web.HttpUtility.HtmlDecode(RegularExpressions.Regex.Match(html, "(?<=<title>).+?(?=</title>)", RegularExpressions.RegexOptions.IgnoreCase).Value)
                        'TODO:url.styの仕様詳細の確認した方がより良いね
                        TextBox2.AppendText(vbTab & "\bibitem{}" & TeXescape(title) & "\\" & "\url{" & Web.HttpUtility.UrlDecode(urlescapedURL).Replace("}", "") & "} " & Now.ToString("yyyy/M/d") & "閲覧" & vbCrLf)
                    Catch ex As Exception
                        TextBox2.BackColor = Color.Red
                        TextBox2.AppendText(vbCrLf & "エラー " & urlescapedURL & "について" & vbCrLf & ex.ToString & vbCrLf)
                    End Try
                Else
                    Try
                        Dim book As ndl.book_info = ndl.download_book_info(urlescapedURL)
                        TextBox2.AppendText(vbTab & "\bibitem{}" & TeXescape(book.author) & ":" & TeXescape(book.title) & "," & book.publisher & ",p.(" & book.release_year & ")" & vbCrLf)
                    Catch ex As Exception
                        TextBox2.BackColor = Color.Red
                        TextBox2.AppendText(vbTab & vbTab & urlescapedURL & ":" & ex.Message & vbCrLf)
                    End Try
                End If
            Next
            TextBox2.AppendText("\end{thebibliography}")
            TextBox2.Focus()
            TextBox2.SelectAll()
        End If
    End Sub
    Private Function TeXescape(ByRef unescapedText As String) As String
        Dim escapedText As String = unescapedText

        escapedText = escapedText.Replace(" \ ", " $\backslash$ ") 'エスケープに使うから最初にエスケープ
        escapedText = escapedText.Replace("_", "\_")
        escapedText = escapedText.Replace("%", "\%")
        escapedText = escapedText.Replace("#", "\#")
        escapedText = escapedText.Replace("$", "\$")
        escapedText = escapedText.Replace("&", "\&")
        escapedText = escapedText.Replace("{", "\{")
        escapedText = escapedText.Replace("}", "\}")
        escapedText = escapedText.Replace("^", "\^")
        escapedText = escapedText.Replace("~", "\~")

        escapedText = escapedText.Replace("<", "$<$")
        escapedText = escapedText.Replace(">", "$>$")
        escapedText = escapedText.Replace("|", "$|$")

        Return escapedText
    End Function
End Class
Public Class ndl
    Shared cookie As New CookieContainer

    Public Shared Function download_book_info(ByRef book_name As String) As book_info
        Dim search_html As String
        Dim HTTPreq As HttpWebRequest = WebRequest.CreateHttp("http://iss.ndl.go.jp/books?op_id=1&any=" & book_name & "&mediatype=1")
        HTTPreq.CookieContainer = cookie
        Using APIsr As New IO.StreamReader(HTTPreq.GetResponse.GetResponseStream(), Encoding.UTF8)
            search_html = APIsr.ReadToEnd
        End Using

        Dim first_item_URL As String = RegularExpressions.Regex.Match(search_html, "http://iss.ndl.go.jp/books/.+?(?="")").Value

        If first_item_URL = "" Then
            Throw New ArgumentException("書籍が見つかりませんでした．")
        End If

        Dim HTTPreq2 As HttpWebRequest = WebRequest.CreateHttp(first_item_URL)
        HTTPreq2.CookieContainer = cookie
        Dim book_info_html As String
        Using APIsr As New IO.StreamReader(HTTPreq2.GetResponse.GetResponseStream(), Encoding.UTF8)
            'HACK:emタグだけへの対処はちょっと・・・やっぱりWebClientのDownloadString使いたい
            book_info_html = Web.HttpUtility.HtmlDecode(RegularExpressions.Regex.Replace(APIsr.ReadToEnd, "</?em.*?>", ""))
        End Using

        Dim res_book_info As book_info
        res_book_info.title = RegularExpressions.Regex.Match(book_info_html, "(?<=contenttitle"">\n\s{2}<h1>\n\s{3}).+?(?=\n\s{2}</h1>)").Value

        res_book_info.author = RegularExpressions.Regex.Match(book_info_html, "(?<=著者[\s\S]+?>).+?(?=</a>)").Value
        res_book_info.author = RegularExpressions.Regex.Replace(res_book_info.author, "\s(共?著|編)$", "")
        res_book_info.author = res_book_info.author.Replace(",", "")

        res_book_info.publisher = RegularExpressions.Regex.Match(book_info_html, "(?<=出版社</th><td>).+?(?=</td>)").Value
        res_book_info.release_year = RegularExpressions.Regex.Match(book_info_html, "(?<=出版年[\s\S]+?<span>).+?(?=</span>)").Value

        Return res_book_info
    End Function

    Public Structure book_info
        Public title As String
        Public author As String
        Public publisher As String
        Public release_year As String
    End Structure
End Class
