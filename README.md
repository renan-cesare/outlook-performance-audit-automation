# Outlook Performance Audit Automation

> Projeto profissional **sanitizado**: automação de auditoria de **desempenho de carteiras** via Outlook + Excel (envio em massa, histórico e follow-up).

## Contexto
Em rotinas de operações / risco & compliance, é comum existir um processo formal de auditoria quando determinados **sinais de atenção** aparecem em carteiras de clientes.

Neste projeto, o critério interno original envolvia o cruzamento de métricas (ex.: desempenho do cliente x indicadores do assessor).  
Para fins de portfólio e segurança, este repositório utiliza o nome **“Desempenho de Carteiras”** e abstrai detalhes sensíveis, mantendo a **lógica real do processo**.

> ⚠️ Versão sanitizada:
> - Sem nomes reais
> - Sem e-mails reais
> - Sem dados internos
> - Sem planilhas reais
> - Preserva o fluxo e a arquitetura de produção

---

## O que o sistema faz

### 1) Dispatch (envio em massa)
- Lê uma planilha Excel com a lista de casos a auditar
- Lê uma base Excel de profissionais (assessores / líderes)
- Normaliza códigos e valida dados obrigatórios
- Envia e-mails em massa via **Outlook clássico (COM)**
- Insere **token único** por envio para rastreabilidade
- Captura IDs do Outlook:
  - `ConversationID`
  - `InternetMessageID`
  - `EntryID`
- Registra tudo em um **histórico de auditoria** (Excel)

### 2) Follow-up (acompanhamento)
- Lê o histórico de envios
- Verifica se houve resposta (por `ConversationID`)
- Se não houver, dispara **cobrança automática** (Reply) e atualiza histórico
- Se houver, registra data e conteúdo da resposta

---

## Estrutura do projeto

outlook-performance-audit-automation/
│
├── main.py
├── config.example.json
├── requirements.txt
├── README.md
├── LICENSE
│
└── src/
└── performance_audit/
├── init.py
├── config.py
├── logging_utils.py
├── file_lock.py
├── outlook_client.py
├── normalize.py
├── dispatch.py
├── history_store.py
└── followup.py


---

## Configuração

1) Crie um `config.json` no seu PC a partir de `config.example.json`  
2) Ajuste os caminhos das planilhas e parâmetros.

> `config.json` não deve ser versionado.

---

## Como rodar

Instalar dependências:

```bash
pip install -r requirements.txt
Teste seguro (não envia e-mail):

python main.py dispatch --dry-run
Mostrar e-mails antes de enviar:

python main.py dispatch --display-only
Rodar acompanhamento:

python main.py followup --display-only
Segurança e boas práticas
Bloqueia execução se planilhas estiverem abertas

Registra histórico para rastreabilidade

Token único por envio

Sanitização para evitar exposição de dados corporativos

Licença
MIT.
