# Outlook Performance Audit Automation

Automação (sanitizada) para **envio em massa** e **acompanhamento (follow-up)** de auditorias de desempenho de carteiras utilizando **Microsoft Outlook (COM / pywin32)** e **Excel**.

> Este projeto é uma adaptação profissional e sanitizada de uma automação real utilizada em ambiente corporativo.
> Não contém dados reais, e-mails reais, clientes reais ou regras proprietárias.

---

## 📌 Visão Geral

Em muitos ambientes corporativos, processos de auditoria e acompanhamento dependem de:

* Envio manual de e-mails
* Controle manual de quem respondeu e quem não respondeu
* Reenvio manual de cobranças
* Atualização manual de planilhas de controle

Este projeto resolve esse problema fornecendo:

* Envio em massa automatizado via Outlook
* Geração de token único por registro (rastreabilidade)
* Registro centralizado de histórico em Excel
* Rotina de follow-up para identificar respostas e sinalizar pendências

---

## 🎯 O que o projeto faz

* Envia e-mails de auditoria em massa via Microsoft Outlook
* Gera um token único por registro auditado
* Registra cada envio em uma planilha de histórico (Excel), incluindo:

  * Data/hora
  * Cliente
  * Assessor
  * E-mails
  * Token
  * Status
  * IDs do Outlook (quando disponíveis)
* Possui rotina de follow-up que:

  * Varre a caixa de entrada
  * Procura respostas pelo token
  * Marca registros como **RESPONDIDO** ou **COBRADO**

---

## 🧱 Estrutura do Projeto

```text
outlook-structured-operations-audit-automation/
  main.py
  config.example.json
  requirements.txt
  README.md
  .gitignore
  templates/
    email_body.html
  src/
    outlook_audit/
      __init__.py
      config.py
      dispatch.py
      followup.py
      outlook_client.py
      history_store.py
      file_lock.py
      logging_utils.py
```

---

## ⚙️ Como o Processo Funciona (Visão Conceitual)

1. O sistema carrega:

   * Uma planilha com os clientes/operações a serem auditados
   * Uma planilha com a base de profissionais (assessores e líderes)

2. Para cada registro:

   * Um token único é gerado
   * Um e-mail é montado e enviado (ou exibido para conferência)
   * O envio é registrado na planilha de histórico

3. No modo de follow-up:

   * O sistema varre a Inbox do Outlook
   * Procura respostas contendo o token
   * Atualiza o histórico:

     * Marcando como **RESPONDIDO**
     * Ou como **COBRADO** quando não há resposta

---

## 📄 Configuração

Toda a configuração é feita via arquivo JSON.

Use o arquivo de exemplo:

```bash
config.example.json
```

Crie uma cópia local (não versionada):

```bash
config.json
```

E ajuste:

* Caminhos das planilhas
* E-mail remetente do Outlook
* Modo de envio (`display` ou `send`)

> ⚠️ O repositório não inclui arquivos reais de dados nem planilhas reais.

---

## ▶️ Como Executar

### 1) Instalar dependências

```bash
pip install -r requirements.txt
```

### 2) Envio das auditorias (modo seguro primeiro)

```bash
python main.py --config config.json dispatch
```

> Recomenda-se começar com `send_mode = "display"` para validar os e-mails antes do envio real.

### 3) Rodar o follow-up

```bash
python main.py --config config.json followup
```

---

## 📊 Arquivos de Dados

O projeto espera planilhas Excel contendo:

* Base de clientes/operações a serem auditadas
* Base de profissionais (assessores / líderes)
* Base de histórico (gerada automaticamente)

Esses arquivos **não fazem parte do repositório** por motivos de confidencialidade.

---

## 🔐 Segurança e Privacidade

* Nenhuma credencial é armazenada no projeto
* A integração com Outlook é feita via cliente local (COM)
* Este repositório não contém:

  * Dados reais de clientes
  * Dados operacionais reais
  * Estruturas internas de empresas

Este código é destinado a **portfólio, estudo e referência técnica**.

---

## ⚠️ Limitações

* Funciona apenas em Windows
* Requer Microsoft Outlook instalado e configurado
* Utiliza Excel como base de persistência (não usa banco de dados)
* A identificação de respostas depende da consistência da caixa de e-mail

---

## 🧠 Filosofia do Projeto

Este projeto foi desenhado para:

* Refletir restrições reais de ambientes corporativos
* Priorizar robustez e rastreabilidade
* Integrar-se ao ecossistema existente (Outlook + Excel)
* Ser evoluído no futuro para banco de dados e dashboards, se necessário

---

## 📌 Aviso Legal

Este projeto é uma adaptação sanitizada de uma automação corporativa real.
Ele não representa nenhuma empresa, cliente, produto ou processo específico.

---

## 📜 Licença

MIT
