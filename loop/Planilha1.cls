VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Ex01()

    Dim nome As String
    nome = InputBox("Informe o seu Nome")
    
End Sub

Sub Ex02()
    
    Dim nome As String
    Dim sobrenome As String
    nome = InputBox("Informe o seu Nome")
    sobrenome = InputBox("Informe seu Sobrenome")
        
            MsgBox "Seu nome completo é " & nome & " " & sobrenome, vbInformation
                   
End Sub

Sub Ex03()

   Dim nome As String
   Dim sobrenome As String
   Dim idade As Byte
   
   On Error GoTo erro
   
        nome = InputBox("Informe o seu Nome")
        sobrenome = InputBox("Informe o seu Sobrenome")
        idade = InputBox("Informe a sua idade")
        MsgBox "Seu nome completo é " & nome & " " & sobrenome & " " & "com idade de " & idade & " " & "anos" & ".", vbInformation
        Exit Sub
    
erro:
   MsgBox ("Digite um valor válido"), vbCritical
   
    
        
             
             
        
End Sub

