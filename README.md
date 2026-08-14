# Disparador de E-mails com Outlook

Aplicação em Python para automatizar envios de e-mails personalizados via Microsoft Outlook Desktop, com interface em Streamlit.

O sistema permite:

- carregar uma planilha Excel com dados dos destinatários;
- marcar quais linhas devem ser enviadas;
- validar e-mails, anexos e conteúdo da mensagem;
- personalizar o texto com campos da planilha;
- usar templates HTML ou editor simples;
- enviar mensagens pelo Outlook com suporte a anexos fixos e personalizados;
- gerar uma pré-visualização antes do envio.

---

## Visão geral do projeto

Este projeto foi pensado para facilitar campanhas e envios em massa, mantendo um fluxo simples e visual para o usuário final.

A aplicação é executada por meio do arquivo principal:

- [app_streamlit.py](app_streamlit.py)

A lógica de processamento e validações fica em:

- [core.py](core.py)

O atalho para inicializar o sistema está em:

- [iniciar_sistema.bat](iniciar_sistema.bat)

---

## Estrutura da pasta

A raiz do projeto contém:

- [app_streamlit.py](app_streamlit.py): interface gráfica do sistema em Streamlit;
- [core.py](core.py): funções utilitárias para processamento, validação e envio;
- [requirements.txt](requirements.txt): dependências do projeto;
- [iniciar_sistema.bat](iniciar_sistema.bat): script para criar ambiente virtual e iniciar a aplicação;
- [Arquivo](Arquivo): pasta central de dados e arquivos do sistema.

### Análise da pasta [Arquivo](Arquivo)

A pasta [Arquivo](Arquivo) é o coração operacional do sistema. Ela organiza todos os arquivos usados em execução e deve ser tratada como área de dados locais.

#### [Arquivo/Planilhas](Arquivo/Planilhas)
- local para arquivos Excel usados como base de dados;
- a aplicação procura por arquivos `.xlsx` aqui;
- normalmente contém as listas de destinatários e a coluna de controle de envio.

#### [Arquivo/Anexos_Personalizados](Arquivo/Anexos_Personalizados)
- área para anexos diferentes por pessoa;
- os nomes dos arquivos vêm da planilha;
- a aplicação valida se cada anexo existe antes do envio.

#### [Arquivo/Anexos_Fixos](Arquivo/Anexos_Fixos)
- arquivos que serão enviados para todos os destinatários;
- ideal para materiais gerais, como PDFs, imagens ou documentos institucionais.

#### [Arquivo/Templates](Arquivo/Templates)
- pasta para modelos HTML dos e-mails;
- pode conter layouts prontos com branding, assinatura ou estrutura visual.

#### [Arquivo/Logs](Arquivo/Logs)
- histórico de processamento, validação e resultados por linha;
- geralmente guarda saídas geradas durante execução.

> Observação: essas pastas são criadas automaticamente quando a aplicação inicia, caso ainda não existam.

---

## Como funciona o fluxo

1. O usuário inicia o sistema pelo script [iniciar_sistema.bat](iniciar_sistema.bat) ou executa o Streamlit manualmente.
2. A aplicação abre a interface no navegador.
3. O usuário seleciona a planilha Excel e a aba correta.
4. A aplicação identifica colunas como:
   - envio;
   - e-mail do destinatário;
   - anexos personalizados (quando houver).
5. O usuário escolhe o tipo de mensagem:
   - texto simples;
   - HTML personalizado;
   - com anexos fixos ou por pessoa.
6. O sistema valida:
   - e-mails;
   - presença de anexos;
   - template vazio;
   - linhas marcadas para envio.
7. A mensagem pode ser pré-visualizada antes do envio real.
8. Quando habilitado, a aplicação integra com Outlook Desktop para disparar os e-mails.

---

## Requisitos

- Windows;
- Python 3 instalado e disponível no PATH;
- Microsoft Outlook Desktop instalado e aberto;
- conta de e-mail configurada no Outlook;
- dependências listadas em [requirements.txt](requirements.txt).

---

## Instalação e execução

### 1. Clonar o projeto

```bash
git clone <url-do-repositorio>
cd disparador-emails-microsoft
```

### 2. Instalar dependências

```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

### 3. Rodar a aplicação

```bash
streamlit run app_streamlit.py
```

Ou, no Windows, execute:

```bat
iniciar_sistema.bat
```

---

## Dependências principais

O projeto usa bibliotecas como:

- `streamlit`: interface web do sistema;
- `pandas`: manipulação de planilhas e dados em tabela;
- `openpyxl`: leitura de arquivos Excel;
- `pywin32` / integração com Outlook via `win32com`.

---

## Observações importantes

- A aplicação foi desenhada para uso em ambiente Windows, principalmente por causa da integração com Outlook.
- O processo usa a conta ativa do Outlook do usuário, então a validação da conta correta é parte essencial do fluxo.
- Planilhas, logs e anexos são arquivos locais e, em geral, não devem ser versionados no Git.
- O sistema não deve ser usado sem revisar a mensagem antes do envio real.

---

## Boas práticas

- manter a planilha em [Arquivo/Planilhas](Arquivo/Planilhas);
- deixar anexos personalizados em [Arquivo/Anexos_Personalizados](Arquivo/Anexos_Personalizados);
- colocar modelos reutilizáveis em [Arquivo/Templates](Arquivo/Templates);
- revisar logs em [Arquivo/Logs](Arquivo/Logs) após cada lote de envio;
- manter o [iniciar_sistema.bat](iniciar_sistema.bat) simples e genérico.

---

## Licença

Este projeto foi criado para uso interno e operacional local. Ajustes e melhorias podem ser feitos conforme a necessidade da organização ou do usuário responsável.
