# Outlook Performance Audit Automation

Automação em Python para **auditoria de desempenho de carteiras via Outlook e Excel**, com **envio em massa de e-mails**, **histórico de execuções**, **controle de follow-up** e **registro consolidado para acompanhamento operacional**.

> **English (short):** Python automation for performance audit workflows using Outlook and Excel, including bulk email dispatch, execution history, and automated follow-up.

---

## Principais recursos

* Automação de envio de e-mails via **Outlook (COM automation)**
* Geração dinâmica de mensagens a partir de **templates HTML**
* Processamento de bases em **Excel**
* Envio em massa com:

  * controle por assessor/carteira
  * histórico de execuções
  * identificador único por envio
* **Follow-up automático** baseado em histórico (e-mails sem resposta)
* Registro consolidado para auditoria e rastreabilidade
* Separação clara entre:

  * dados
  * templates
  * lógica de negócio

---

## Contexto

Em rotinas de **risco, compliance e backoffice**, auditorias de desempenho exigem:

* contato recorrente com assessores ou responsáveis
* envio estruturado de informações
* controle de quem respondeu ou não
* histórico auditável das interações

Este projeto automatiza esse fluxo operacional, reduzindo esforço manual e garantindo **padronização, rastreabilidade e controle**.

---

## Aviso importante (uso autorizado)

Este repositório é apresentado como **exemplo técnico/portfólio**.

* Utilize apenas **ambientes e contas autorizadas**
* Não publique dados reais, e-mails corporativos ou informações sensíveis
* Respeite políticas internas, LGPD e regras de uso do Outlook

---

## Estrutura do projeto

```text
.
├─ src/
│  └─ performance_audit/
│     ├─ __init__.py
│     ├─ app.py
│     ├─ followup.py
│     └─ outlook_client.py
├─ templates/
│  └─ email_body.html
├─ config.example.json
├─ main.py
├─ requirements.txt
├─ LICENSE
└─ README.md
```

---

## Requisitos

* Python 3.10+
* **Windows**
* Microsoft Outlook instalado e configurado

> Este projeto utiliza automação COM, sendo compatível apenas com ambiente Windows.

---

## Instalação

```bash
python -m venv .venv

# Windows
.venv\Scripts\activate

pip install -r requirements.txt
```

---

## Configuração

Crie um arquivo local de configuração:

```bash
copy config.example.json config.json
```

Campos principais do `config.json` incluem:

* caminhos de arquivos Excel
* parâmetros de envio
* controles de follow-up
* opções de execução

> O arquivo `config.json` deve permanecer fora do versionamento.

---

## Execução

```bash
python main.py
```

O processo:

* lê a base em Excel
* envia os e-mails conforme regras definidas
* registra o histórico de envios
* executa follow-up automático quando aplicável

---

## Saídas geradas

* Histórico consolidado de envios
* Controle de follow-up
* Evidências para auditoria operacional

---

## Sanitização de dados

Este repositório **não contém dados reais**.

* Bases Excel reais devem permanecer fora do Git
* Templates HTML podem ser versionados normalmente
* Identificadores sensíveis são gerados apenas em tempo de execução

---

## Licença

MIT
