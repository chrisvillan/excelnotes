Attribute VB_Name = "HTML"
Private Sub parsehtml()
    Dim http As Object, html As New HTMLDocument, topics As Object, titleElem As Object, detailsElem As Object, topic As HTMLHtmlElement
    Dim i As Integer
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", "https://news.ycombinator.com/", False
    http.send
    html.body.innerHTML = http.responseText
    
    Set topics = html.getElementsByClassName("athing")
    i = 2
    
    For Each topic In topics
        Set titleElem = topic.getElementsByTagName("td")(2)
        Sheets(1).Cells(i, 1).Value = titleElem.getElementsByTagName("a")(0).innerText
        Sheets(1).Cells(i, 2).Value = titleElem.getElementsByTagName("a")(0).href
        Set detailsElem = topic.NextSibling.getElementsByTagName("td")(1)
        Sheets(1).Cells(i, 3).Value = detailsElem.getElementsByTagName("span")(0).innerText
        Sheets(1).Cells(i, 4).Value = detailsElem.getElementsByTagName("a")(0).innerText
        i = i + 1
    Next
End Sub

