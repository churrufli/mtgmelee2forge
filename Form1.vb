Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Text.RegularExpressions
Public Class Form1

    Public Function GetHtmlDocument(ByVal html As String) As System.Windows.Forms.HtmlDocument
        Dim browser As WebBrowser = New WebBrowser()
        browser.ScriptErrorsSuppressed = True
        browser.DocumentText = html
        browser.Document.OpenNew(True)
        browser.Document.Write(html)
        browser.Refresh()
        Return browser.Document
    End Function


    Public Shared Function ReadWeb(MyUrl As String)
        Try
            Dim client As WebClient = New WebClient()
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
            Dim reply As String = client.DownloadString(MyUrl)
            Return reply
        Catch
        End Try

    End Function
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim url = "https://mtgmelee.com/Decklist/View/140939"
        url = TextBox1.Text

        Dim t = ReadWeb(url)
        Dim doc As HtmlDocument = GetHtmlDocument(t)
        Dim links As New List(Of String)
        Dim tit As String
        Dim deck As String
        For Each Element As HtmlElement In doc.GetElementsByTagName("button")
            'MsgBox(Element.OuterHtml)
            If Element.OuterHtml.Contains("Copy for MTG") Then
                deck = Element.GetAttribute("data-clipboard-text")
            End If
        Next

        For Each Element As HtmlElement In doc.GetElementsByTagName("a")
            If Element.OuterHtml.Contains("decklist-card-title mr-3") Then
                tit = Element.OuterHtml

            End If
        Next

        deck = deck
        Dim folder = Directory.GetCurrentDirectory() & "/result/"
        folder = folder
        Dim number As String = Directory.GetFiles(folder, "*.dck").Count + 1
        If Len(number) = 1 Then number = "0" & number

        tit = Split(tit, ">")(1).ToString
        tit = Split(tit, "<")(0).ToString
        'nombre del creador del mazo
        Dim creat As String
        For Each Element As HtmlElement In doc.GetElementsByTagName("span")
            If Element.OuterHtml.Contains("decklist-card-title-author") Then
                creat = Element.OuterHtml
            End If
        Next

        creat = Split(creat, "_blank>")(1).ToString
        creat = Split(creat, "</A")(0).ToString

        tit = "#" & number & " " & tit & " by " & creat

        deck = Replace(deck, "Deck" & vbCrLf, Nothing)
        deck = Replace(deck, "Sideboard", "[sideboard]")
        deck = "[metadata]" & vbCrLf & "Name=" & tit & vbCrLf & "[Main]" & vbCrLf & deck
        Dim file As System.IO.StreamWriter
        file = My.Computer.FileSystem.OpenTextFileWriter(folder & tit & ".dck", True)
        file.WriteLine(deck)
        file.Close()
        MsgBox("Saved " & tit)

    End Sub
End Class
