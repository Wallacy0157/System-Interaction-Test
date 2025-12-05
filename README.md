# 🖥️ VBScript - Teste de Interação e Comportamento do Sistema

Este repositório contém um script em **VBScript (.vbs)** projetado para testar a capacidade de interação do ambiente Windows com caixas de diálogo, loops controlados e geração de logs.

O objetivo principal é validar se o sistema permite a execução de scripts, exibição de prompts e monitoramento de respostas do usuário.

---

## 📌 Funcionalidades

O script executa os seguintes testes:

1. **Exibição de diálogos interativos**
   Exibe até 5 caixas de diálogo pedindo confirmação do usuário.

2. **Controle de loop com interrupção manual**
   O teste para imediatamente caso o usuário clique em **"Não"**.

3. **Geração de log detalhado**
   Todas as ações são registradas no arquivo:

   ```
   teste_interacao.log
   ```

   O log é criado no mesmo diretório do script.

4. **Mensagem final informando conclusão e caminho do relatório**

---

## 🧠 Como o script funciona

Trechos principais:

### 📝 Registro do início do teste

```vb
logFile.WriteLine "Teste iniciado em: " & Now
```

### 🔁 Loop de interação com o usuário

```vb
resposta = MsgBox("Este é um teste autorizado...", vbQuestion + vbYesNo, "Teste de Segurança")
```

### 📤 Finalização manual

```vb
If resposta = vbNo Then
    logFile.WriteLine Now & " - Usuário encerrou o teste manualmente."
    Exit Do
End If
```

### ✅ Mensagem final

```vb
MsgBox "O teste foi concluído com sucesso...", vbInformation, "Finalizado"
```

---

## ▶️ Como usar

1. Salve o script com a extensão **.vbs** (ex: `teste_interacao.vbs`).
2. Execute clicando duas vezes ou via CMD com:

   ```bash
   cscript teste_interacao.vbs
   ```
3. Interaja com as caixas de diálogo.
4. Consulte o log **teste_interacao.log** para ver o registro completo.

---

## 📄 Estrutura do log

Exemplo:

```
===============================================
Teste iniciado em: 05/12/2025 10:30:15
===============================================
05/12/2025 10:30:15 - Exibindo diálogo de teste (1/5)...
05/12/2025 10:30:20 - Exibindo diálogo de teste (2/5)...
...
05/12/2025 10:31:02 - Teste finalizado.
===============================================
```

---

## 🔤 Nome sugerido para o script

```
teste_interacao.vbs
```

Outras opções:

* system_interaction_test.vbs
* vbs_interaction_checker.vbs
* dialog_behavior_test.vbs
