﻿Imports System.Net
Imports System.Text

Public Class Form1
	Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
		If e.Control AndAlso e.KeyCode = Keys.A Then
			TextBox2.SelectAll()
		End If
	End Sub

	Private Sub ListBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ListBox1.KeyDown
		If e.Control AndAlso e.KeyCode = Keys.V Then
			ListBox1.Items.Add(Clipboard.GetText)
		End If

		If e.Control AndAlso e.KeyCode = Keys.A Then
			ListBox1.Items.Add(InputBox("URLか書名を入力して下さい．"))
		End If

		If ListBox1.SelectedIndex > -1 Then
			If e.KeyCode = Keys.Delete Then
				ListBox1.Items.RemoveAt(ListBox1.SelectedIndex)
			End If
		End If

		If e.Control AndAlso e.Shift AndAlso e.KeyCode = Keys.Enter Then
			TextBox2.Clear()
			TextBox2.AppendText("\begin{thebibliography}{" & CStr(10 ^ (Math.Floor(Math.Log10(ListBox1.Items.Count)) + 1) - 1) & "}" & vbCrLf)
			For Each escapedURL As String In ListBox1.Items
				If RegularExpressions.Regex.IsMatch(escapedURL, "^https?") Then
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
							title = Web.HttpUtility.HtmlDecode(html.Substring(title_begin, title_end - title_begin))
						End If

						TextBox2.AppendText(vbTab & "\bibitem{}" & TeXescape(title) & "\\" & TeXescape(Web.HttpUtility.UrlDecode(escapedURL)) & " " & Now.ToString("yyyy/M/d") & "閲覧" & vbCrLf)
					Catch ex As Exception
						TextBox2.AppendText(vbCrLf & "エラー " & escapedURL & "について" & vbCrLf & ex.ToString & vbCrLf)
					End Try
				Else
					Dim book As New ndl_book(escapedURL)
					TextBox2.AppendText(vbTab & "\bibitem{}" & TeXescape(book.author) & ":" & TeXescape(book.title) & "," & book.publisher & ",p.(" & book.release_year & ")" & vbCrLf)
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
Public Class ndl_book
	Public title As String
	Public author As String
	Public publisher As String
	Public release_year As String

	Sub New(ByRef book_name As String)
		Dim wc As New WebClient
		wc.Encoding = Encoding.UTF8
		Dim search_html As String = wc.DownloadString("http://iss.ndl.go.jp/books?ar=4e1f&any=" & book_name & "&display=&op_id=1&mediatype=1")
		Dim first_item_URL As String = RegularExpressions.Regex.Match(search_html, "http://iss.ndl.go.jp/books/.+?(?="")").Value

		Dim book_info_html As String = wc.DownloadString(first_item_URL)
		'HACK:ここまで書かなくても[\s\S]で済ませる?
		title = RegularExpressions.Regex.Match(book_info_html, "(?<=contenttitle"">\n\s{2}<h1>\n\s{3}).+?(?=\n\s{2}</h1>)").Value

		author = RegularExpressions.Regex.Match(book_info_html, "(?<=著者[\s\S]+?>).+?(?=</a>)").Value
		author = RegularExpressions.Regex.Replace(author, "\s(共?著|編)$", "")
		author = author.Replace(",", "")

		publisher = RegularExpressions.Regex.Match(book_info_html, "(?<=出版社</th><td>).+?(?=</td>)").Value
		release_year = RegularExpressions.Regex.Match(book_info_html, "(?<=出版年[\s\S]+?<span>).+?(?=</span>)").Value
	End Sub
End Class