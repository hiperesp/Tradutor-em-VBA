Attribute VB_Name = "Translator"
Private isTranslating As Boolean

Public Function Translate(sourceText As String, sourceLang As String, destLang As String) As String
    
    If sourceText <> "" Then
        sourceLang = parseLang(sourceLang)
        destLang = parseLang(destLang)
    
        Dim req As Object
        Dim JSON As Dictionary
        Dim data As String
        
        Set req = CreateObject("MSXML2.XMLHTTP")
        req.Open "POST", "https://translate.yandex.net/api/v1.5/tr.json/translate?key=trnsl.1.1.20191202T002434Z.3a41f3e25b844206.ba01cb199038c87c92d6d8ecfc5574c8f60afcec&lang=" & sourceLang & "-" & destLang, False
        req.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=ISO-8851-1"
        req.Send "text=" & sourceText
        
        Dim texto As String
        texto = req.responseText
        
        Set JSON = JsonConverter.ParseJson(texto)
        
        Translate = JSON("text")(1)
    Else
        Translate = sourceText
    End If
    
End Function

Private Function parseLang(lang As String) As String
    If lang = "Inglês" Or lang = "English" Or lang = "en" Then
        parseLang = "en"
    ElseIf lang = "Espanhol" Or lang = "Spanish" Or lang = "Español" Or lang = "Espanol" Or lang = "Castellano" Or lang = "es" Then
        parseLang = "es"
    ElseIf lang = "Italiano" Or lang = "Italian" Or lang = "it" Then
        parseLang = "it"
    Else
        parseLang = "pt"
    End If
End Function



