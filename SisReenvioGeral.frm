VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSisReenvio 
   Caption         =   "Form1"
   ClientHeight    =   1155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   ScaleHeight     =   1155
   ScaleWidth      =   4380
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3510
      Top             =   315
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmSisReenvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Recibido As Boolean
Dim Mensaje As String
Dim sckError

Private Sub Form_Load()
   '
   ' Autor
   ' Herley Nicolas Ramos Sanchez
   ' e-mail: Nicolas.MCP@gmail.com
   '
   ' Licença GNU GPL (Software Livre)
   '
   Dim objFSO As Scripting.FileSystemObject
   Dim objFile
   Dim lSize As Long
   Dim x As Integer
   Dim ArqLeido As String
   
   If Dir(App.Path & "\leitura.log") = "" Then
      lSize = 100000000
   Else
      Set objFSO = New Scripting.FileSystemObject
      Set objFile = objFSO.GetFile(App.Path & "\leitura.log")
      lSize = objFile.Size
      Set objFile = Nothing
      Set objFSO = Nothing
   End If

   If lSize > 5000000 Then
      Open App.Path & "\leitura.log" For Output As #LOG
   Else
      Open App.Path & "\leitura.log" For Append As #LOG
   End If
   If Dir(App.Path & "\e_geral*.eml", vbArchive) = "" _
      Then
      ' Nenhum e-mail baixado
      
      Print #LOG, "-------------------------------------------------------------"
      Print #LOG, "xxxxxxxxx Começo da rotina de baixar e-mails GERAIS xxxxxxxxx"
      Call Leer("192.168.3.42", 110, "e_geral", "ge5rm72nr")
'      Call Leer("192.168.136.130", 110, "nicolas", "senha_kp223*(U", 2)
      DoEvents
      Print #LOG, "------------ Fim da rotina de baixar e-mails GERAIS ------------"
   Else
      Print #LOG, "==== Já existem e-mails GERAIS baixados, devem ser processados primeiro " & Date & " " & Time & " ===="
      For x = 1 To 10
         ArqLeido = Dir(App.Path & "\e_geral*.eml", vbArchive)
         If ArqLeido <> "" Then
            Envio (App.Path & "\" & ArqLeido)
            DoEvents
            Print #LOG, "] Enviado " & App.Path & "\" & ArqLeido & " >>>------> " & Date & " " & Time
            Kill (App.Path & "\" & ArqLeido)
            ArqLeido = Dir(App.Path & "\e_geral*.eml", vbArchive)
            Call Sleep(3000) 'espera 3 segundos
         End If
      Next x
   End If
   Close #LOG
   ' Fim Rotina baixar e-mails de E_GERAL
     
   Unload Me
End Sub

' se reciben los datos
''''''''''''''''''''''''''''
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
  
  Winsock1.GetData Mensaje
    
  Select Case Winsock1.Tag
  Case "RETR"
      ' escribe en disco el archivo eml
      Put #MAIL, , Mensaje
                
      If InStr(Mensaje, vbLf + "." + vbCrLf) Then
          Close #MAIL
          Recibido = True
      End If
  Case Else
      sckError = (Left$(Mensaje, 3) = "-ER")
      Recibido = True
  End Select
End Sub

Private Sub Winsock1_Close()
  Winsock1.Close
End Sub

Private Sub Leer(Pop3 As String, _
                 Puerto As Integer, _
                 Cuenta As String, _
                 Password As String)

   Dim x, b, Messages, i As Variant
   Dim ind As String
   
   ' Conecta con el Pop3
   Winsock1.Connect Pop3, Puerto
     
    Do Until Recibido
       DoEvents
    Loop
 
    If sckError Then
       Print #LOG, Date & " " & Time & " Erro de conexão com POP3"
       Exit Sub
    End If

  ' usuario
  SendCommand "USER " & Cuenta
  
  If sckError Then
    Print #LOG, Date & " " & Time & " ERRO: Nome de usuário incorreto!  "
    Exit Sub
  End If
  
  ' Envía el passowrd
  SendCommand "PASS " & Password
  
  If sckError Then
    Print #LOG, Date & " " & Time & " ERRO: Senha incorreta!"
    Exit Sub
  End If

' Get Number of Messages and total size in bytes
  SendCommand "STAT"
  x = InStr(Mensaje, " ")
  b = InStrRev(Mensaje, " ")
  Messages = Val(Mid$(Mensaje, x + 1, b - x))
  
' Recorre la cantidad de mensajes
  For i = 1 To Messages
      
      Winsock1.Tag = "RETR"
      
      'crea un archivo
      ind = Format(i, "0000")
      Open App.Path & "\" & Cuenta & ind & ".eml" For Binary Access Write As #MAIL
        
      Call SendCommand("RETR " & i)
      Print #LOG, Cuenta & ind & ".eml"
      
      Winsock1.SendData "DELE " & i & vbCrLf
      Print #LOG, "Deletado da Caixa de Entrada: " & Cuenta & ind & ".eml"
  Next
  Winsock1.Tag = ""
  
  If i > 1 Then
     i = i - 1
     Print #LOG, "Total de e-mails baixados " & ind & " ***************** " & Date & " " & Time
     ' Ha bajado algun mail, reenviar-lo
     For x = 1 To i
         ind = Format(x, "0000")
         Envio (App.Path & "\" & Cuenta & ind & ".eml")
         DoEvents
         Print #LOG, "Enviado " & App.Path & "\" & Cuenta & ind & ".eml >>>------> " & Date & " " & Time
         DoEvents
         Kill (App.Path & "\" & Cuenta & ind & ".eml")
     Next x
     Print #LOG, "*** Fim de REENVIO ***"
  Else
     Print #LOG, "Nenhum e-mail baixado ***************** " & Date & " " & Time
  End If
  
End Sub

Private Sub Envio(Anexo As String)
   Dim objMensaje As CDO.Message
   Dim Destinatarios As String
   
   On Error GoTo Errores
   
   Set objMensaje = New CDO.Message
   With objMensaje
   .From = "sis_reenvio@empresa.com.br"
'   .To = "nicolas@empresa.com.br;nicolas.mcp@gmail.com"
   Open App.Path & "\to.dat" For Input As #DEST
   Input #DEST, Destinatarios
   Close #DEST
   .BCC = Destinatarios
   .Subject = "SisReenvio"
   .AddAttachment Anexo
   .HTMLBody = "<p><font color=""#003300"" size=""+5"">Sistema de Reenvio</font></p><p>&nbsp;</p><p><font color=""#000033"">Verifique o Arquivo Anexo.</font></p><p><font color=""#663300"">Não responda a este e-mail</font></p><p><font color=""#663300"">Utilize o seu catálogos de endereços</font></p><p><font color=""#663300"">se desejar responder a alguem.</font></p>"
   
   With objMensaje.Configuration.Fields
      .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
      .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "200.98.199.2"
      .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = True 'cdoBasic
      .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "sis_reenvio@empresa.com.br"
      .Item("http://schemas.Microsoft.com/CDO/Configuration/sendpassword") = "z9k4w0vy"
      .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 587
      .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
      .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
      .Update
   End With

   .Send
   End With
   Set objMensaje = Nothing
   
   GoTo Fin
   

Errores:
   Print #LOG, "\\\\\ Erro ao enviar /////"
   Close #LOG
   DoEvents
   End
   
Fin:
End Sub

Sub SendCommand(Comando As String)
   Winsock1.SendData Comando + vbCrLf
    
   Recibido = False
   Do Until Recibido
       DoEvents
   Loop
End Sub

