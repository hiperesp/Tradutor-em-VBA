VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Tela 
   Caption         =   "Hiper Tradutor"
   ClientHeight    =   5655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11130
   OleObjectBlob   =   "Tela.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Tela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    Dim sText As String
    Dim sLang As String
    Dim dLang As String
    sText = sourceText.value
    sLang = sourceLang.text
    dLang = destLang.text
    destText.value = Translator.Translate(sText, sLang, dLang)

End Sub


Private Sub sourceText_Change()

End Sub

Private Sub UserForm_Initialize()
    addComboBoxLanguages sourceLang
    addComboBoxLanguages destLang
End Sub

Private Sub addComboBoxLanguages(languageComboBox As comboBox)
    With languageComboBox
        .AddItem "Português"
        .AddItem "Inglês"
        .AddItem "Espanhol"
        .AddItem "Italiano"
        .text = "Português"
    End With
End Sub

