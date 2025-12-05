' =============================================================
' Teste de Interação e Comportamento do Sistema
' Objetivo: validar se o ambiente permite a execução de scripts,
' exibição de diálogos e loops controlados.
' Uso restrito a testes autorizados.
' =============================================================

Option Explicit

Dim resposta
Dim contador
Dim logFile, fso, logPath

' ===== Configuração do Log =====
Set fso = CreateObject("Scripting.FileSystemObject")
logPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\teste_interacao.log"
Set logFile = fso.OpenTextFile(logPath, 8, True)

logFile.WriteLine "==============================================="
logFile.WriteLine "Teste iniciado em: " & Now
logFile.WriteLine "==============================================="

contador = 0

Do Until contador = 5
    logFile.WriteLine Now & " - Exibindo diálogo de teste (" & contador+1 & "/5)..."

    resposta = MsgBox("Este é um teste autorizado de interação com o sistema." & vbCrLf & _
                      "Deseja continuar o teste?", _
                      vbQuestion + vbYesNo, _
                      "Teste de Segurança")

    If resposta = vbNo Then
        logFile.WriteLine Now & " - Usuário encerrou o teste manualmente."
        Exit Do
    End If

    contador = contador + 1
Loop

MsgBox "O teste foi concluído com sucesso." & vbCrLf & _
       "Relatório gerado em: " & logPath, _
       vbInformation, _
       "Finalizado"

logFile.WriteLine Now & " - Teste finalizado."
logFile.WriteLine "==============================================="
logFile.Close

