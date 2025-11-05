# HOSPITAL ROBOT 2

Este robô automatiza o processo de download, processamento e envio de boletos em aberto por email.

## 📋 Funcionamento do Robô

### Fluxo Principal:

1. **Login no Outlook** - Acessa automaticamente o email corporativo
2. **Download de Anexos** - Busca emails não lidos com assunto específico e baixa planilhas de boletos (Bradesco e Itaú) e extrai arquivos ZIP se necessário
3. **Processamento de Dados** - Processa as planilhas e gera PDFs individuais para cada hospital
4. **Agrupamento Inteligente** - Agrupa os PDFs por hospital usando matching inteligente de nomes
5. **Envio de Emails** - Envia cada conjunto de PDFs para os emails correspondentes do hospital (com suporte a múltiplos destinatários e CCs)
6. **Relatório Final** - Gera e envia relatório com status de todos os envios

## 🛠️ Pré-requisitos para Testar

### 1. Arquivos Necessários

Coloque estes arquivos na pasta `assets/` (ou ajuste o caminho no código):

* `infos do robo.xlsx` - Configurações do robô
* `Relação de e-mails TESTE.xlsx` - Lista de emails dos hospitais

### 2. Estrutura de Pastas

**text**

```
BI03/
├── assets/
│   ├── infos do robo.xlsx
│   └── Relação de e-mails TESTE.xlsx
├── downloads/ (criada automaticamente)
├── boletos_pdf/ (criada automaticamente)
└── rpa.py
```

## ⚙️ Configuração

### 1. Arquivo `infos do robo.xlsx`

Preencha com estas informações:

| Coluna A                     | Coluna B                               |
| ---------------------------- | -------------------------------------- |
| assunto do email             | Assunto do email que contém os anexos |
| caminho para faturas         | Pasta onde salvar os PDFs              |
| email de relatorio           | Email para receber relatórios         |
| caminho dos emails hospitais | Caminho da planilha de emails          |
| email_user                   | Email para login no Outlook            |
| email_pass                   | Senha do email                         |

### 2. Arquivo `Relação de e-mails TESTE.xlsx`

Estruture com estas colunas:

| Hospital       | Email              | Cc 1                 | Cc 2            |
| -------------- | ------------------ | -------------------- | --------------- |
| Hospital Alpha | alpha@hospital.com | financeiro@alpha.com | admin@alpha.com |
| Hospital Beta  | beta@hospital.com  | cobranca@beta.com    |                 |

### 3. Planilhas de Boletos (Bradesco e Itaú)

O robô espera planilhas de boletos dos bancos Bradesco e Itaú. Essas planilhas devem ser anexadas em um email com o assunto configurado e estar em formato Excel (xlsx ou xls).

**📋 ESTRUTURA BRADESCO:**

| Coluna A         | Coluna B          | Coluna C           | Coluna D             | Coluna E              | Coluna F        |
| ---------------- | ----------------- | ------------------ | -------------------- | --------------------- | --------------- |
| **Status** | **Pagador** | **Nº Nota** | **Nº Boleto** | **Data Vencim** | **Valor** |
| VENCIDO          | HOSPITAL1         | 123                | 456                  | 2025-03-10            | 1500.00         |

**📋 ESTRUTURA ITAÚ:**

| Coluna A          | Coluna B             | Coluna C         | Coluna D             | Coluna E           | Coluna F             |
| ----------------- | -------------------- | ---------------- | -------------------- | ------------------ | -------------------- |
| **Pagador** | **Vencimento** | **ValorR** | **Nº Boleto** | **Nº Nota** | **Observacao** |
| HOSPITAL2         | 2025-03-15           | 2000.00          | 789                  | 124                | VENCIDO              |

**📌 Observações Importantes:**

* O robô é flexível e tenta mapear as colunas automaticamente, mas é melhor seguir a estrutura acima.
* O agrupamento é feito pela coluna  **Pagador** .
* O robô processa múltiplas planilhas (Bradesco e Itaú) e agrupa todos os boletos de um mesmo hospital, independentemente do banco.

## 📧 Processo de Envio das Planilhas

### **⚠️ ETAPA CRÍTICA:**

Para o robô funcionar, você **DEVE enviar por email** as planilhas de boletos (Bradesco e Itaú):

1. **Destinatário** : O mesmo email configurado em `email_user` no `infos do robo.xlsx`
2. **Assunto** : **Exatamente igual** ao configurado em `assunto do email` no `infos do robo.xlsx`
3. **Anexo** : As planilhas de boletos (pode ser um arquivo ZIP contendo as planilhas ou as planilhas soltas)
4. **Status do Email** : Deve estar **NÃO LIDO** na caixa de entrada

### **Exemplo de Email:**

**text**

```
Para: robot.boletos@empresa.com
Assunto: Boletos em Aberto
Anexo: planilhas_boletos.zip (ou planilhas soltas)
Corpo: (pode estar vazio ou com qualquer texto)
```

## 🧪 Como Testar o Robô

### 1. Preparação do Ambiente

**bash**

```
# Instale as dependências
pip install -r requirements.txt

# Verifique se todos os arquivos estão no lugar
python rpa.py
```

### 2. Teste Passo a Passo

**Passo 1 - Configuração:**

* Verifique se `infos do robo.xlsx` está preenchido corretamente
* Confirme que as pastas de destino existem
* Teste o login manual no Outlook Web

**Passo 2 - Envio das Planilhas:**

* Envie um email para a conta do robô com:
  * **Assunto** : Exatamente igual ao configurado em "assunto do email"
  * **Anexo** : As planilhas de boletos (Bradesco e Itaú) ou um ZIP contendo elas
  * **Status** : Não lido

**Passo 3 - Execução:**

**bash**

```
python rpa.py
```

**Passo 4 - Monitoramento:**

* Observe os logs no console
* Verifique a pasta `downloads/` para os arquivos baixados
* Confira a pasta `boletos_pdf/` para os PDFs gerados
* Aguarde o email de relatório final

## 🔍 O que Observar Durante o Teste

### Comportamentos Esperados:

* ✅ Navegador abre automaticamente
* ✅ Login no Outlook realizado
* ✅ Email com anexo é encontrado e marcado como lido
* ✅ Planilhas são baixadas para `downloads/` (e extraídas se for ZIP)
* ✅ PDFs são gerados em `boletos_pdf/` (um para cada hospital, contendo todos os boletos do hospital)
* ✅ Emails são enviados para os hospitais (com múltiplos anexos se houver mais de um PDF para o mesmo hospital)
* ✅ PDFs são excluídos após envio
* ✅ Relatório é enviado para o email configurado

### Possíveis Problemas:

* ❌ Credenciais incorretas no Excel de configuração
* ❌ Email não encontrado (verificar assunto exato)
* ❌ Planilha de emails com formato incorreto
* ❌ Planilhas de boletos com estrutura muito diferente do esperado
* ❌ Problemas de permissão nas pastas
* ❌ Timeout durante o processo

## 📊 Resultados do Teste

Após a execução, verifique:

1. **Console** : Logs detalhados de cada etapa
2. **Pasta boletos_pdf** : PDFs gerados para cada hospital (apenas durante o processamento, são excluídos após envio)
3. **Email de relatório** : Status de todos os envios
4. **Caixa de saída** : Emails enviados para os hospitais (cada email contém todos os PDFs do hospital)
5. **Email original** : Deve estar marcado como "LIDO"

## 🚨 Solução de Problemas Comuns

### Erro de Login:

* Verifique `email_user` e `email_pass` no Excel
* Teste o login manualmente no Outlook Web

### Email Não Encontrado:

* Confirme o assunto **EXATAMENTE IGUAL** no `infos do robo.xlsx`
* Verifique se o email está na caixa de entrada e **NÃO LIDO**
* Confirme que o anexo é uma planilha Excel ou ZIP

### Problemas com PDFs:

* Verifique permissões de escrita na pasta `boletos_pdf`
* Confirme que as planilhas têm dados válidos na coluna **Pagador**
* Valide a estrutura das planilhas de boletos

### Erros de Envio de Email:

* Valide os emails na planilha `Relação de e-mails TESTE.xlsx`
* Verifique conexão com internet

## 📝 Notas Importantes

* O robô **marca emails como lidos** após processamento
* PDFs são **excluídos automaticamente** após envio
* Em caso de erro, o processo **continua** com os próximos hospitais
* Um **relatório detalhado** é sempre gerado ao final
* Pastas `downloads` e `boletos_pdf` são **limpas** no início de cada execução
* **As planilhas DEEM ser enviadas por email** - não funciona com arquivo local
* O robô agrupa automaticamente os boletos por hospital, mesmo que venham de planilhas diferentes (Bradesco e Itaú)

---

**Pronto para testar!** Execute `python rpa.py` e monitore o processo pelo console.
