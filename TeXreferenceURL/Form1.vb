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
                If Regex.IsMatch(urlescapedURL, "^https?://", RegexOptions.IgnoreCase) Then
                    Try
                        'TODO:エンコーディング確認はもう少し上手い事!
                        Dim html_byte() As Byte
                        Dim wc As New WebClient
                        wc.Proxy = Nothing '要らないかもね
                        html_byte = wc.DownloadData(urlescapedURL)
                        Dim ContentType As String = wc.ResponseHeaders.Item(HttpResponseHeader.ContentType)
                        wc.Dispose()
                        Dim httpheader_charset As String = Regex.Match(ContentType, "(?<=charset=).+", RegexOptions.IgnoreCase).Value

                        Dim html As String
                        If httpheader_charset <> "" Then
                            html = Encoding.GetEncoding(httpheader_charset).GetString(html_byte)
                        Else
                            html = Encoding.ASCII.GetString(html_byte)
                            Dim charset As String = Regex.Match(html, "(?<=charset=""?)[^""]+(?="")", RegexOptions.IgnoreCase).Value
                            If charset <> "" Then
                                html = Encoding.GetEncoding(charset).GetString(html_byte)
                            Else
                                html = Encoding.UTF8.GetString(html_byte)
                            End If
                        End If

                        'HACK:html全体をdecodeしてもいいかも
                        Dim title As String = Web.HttpUtility.HtmlDecode(Regex.Match(html, "(?<=<title>).+?(?=</title>)", RegexOptions.IgnoreCase).Value)
                        'TODO:url.styの仕様詳細の確認した方がより良いね
                        TextBox2.AppendText(vbTab & "\bibitem{}" & TeXescape(title) & "\\" & "\url{" & Web.HttpUtility.UrlDecode(urlescapedURL).Replace("}", "") & "} " & Now.ToString("yyyy/M/d") & "閲覧" & vbCrLf)
                    Catch ex As Exception
                        TextBox2.BackColor = Color.Red
                        TextBox2.AppendText(vbCrLf & "エラー " & urlescapedURL & "について" & vbCrLf & ex.ToString & vbCrLf)
                    End Try
                Else
                    Try
                        Dim book As ndl.book_info = ndl.download_bookinfo_usingAPI(urlescapedURL)
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

    Public Shared Function download_bookinfo_usingAPI(ByRef book_name As String) As book_info
        Dim API_HTML As String
        Dim HTTPreq As HttpWebRequest = WebRequest.CreateHttp("http://iss.ndl.go.jp/api/sru?operation=searchRetrieve&query=title=" & book_name & " AND sortBy=issued_date/sort.ascending&recordSchema=dcndl_simple")
        HTTPreq.CookieContainer = cookie
        Using APIsr As New IO.StreamReader(HTTPreq.GetResponse.GetResponseStream(), Encoding.UTF8)
            API_HTML = APIsr.ReadToEnd
        End Using

        If API_HTML.Contains("<numberOfRecords>0</numberOfRecords>") OrElse Not API_HTML.Contains("numberOfRecords") Then
            Throw New ArgumentException("書籍が見つかりませんでした．")
        End If

        Dim res_book_info As book_info
        res_book_info.title = Regex.Match(API_HTML, "(?<=<dc:title>).+(?=</dc:title>)").Value
        res_book_info.title &= Regex.Match(API_HTML, "(?<=<dcndl:volume>).+(?=</dcndl:volume>)").Value

        res_book_info.author = Regex.Match(API_HTML, "(?<=<dc:creator>).+(?=</dc:creator>)").Value
        res_book_info.author = Regex.Replace(res_book_info.author, "\s(共?著|編)$", "")
        res_book_info.author = res_book_info.author.Replace(",", "")

        res_book_info.publisher = Regex.Match(API_HTML, "(?<=<dc:publisher>).+(?=</dc:publisher>)").Value
        res_book_info.release_year = Regex.Match(API_HTML, "(?<=<dcterms:issued xsi:type=""dcterms:W3CDTF"">).+(?=</dcterms:issued>)").Value

        Return res_book_info
    End Function

    Public Structure book_info
        Public title As String
        Public author As String
        Public publisher As String
        Public release_year As String
    End Structure
End Class
