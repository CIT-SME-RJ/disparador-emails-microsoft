# Disparador de E-mails com Outlook Desktop

Aplicação em Python com interface em Streamlit para auxiliar o envio controlado de e-mails personalizados pelo Microsoft Outlook Desktop.

O sistema permite carregar uma planilha Excel, personalizar mensagens com dados da tabela, validar e-mails e anexos, revisar a pré-visualização e escolher entre validação, visualização no Outlook ou envio real.

## Objetivo

Este sistema foi criado para facilitar disparos de e-mails em lote com mais segurança operacional.

Ele permite:

- carregar uma planilha Excel com dados dos destinatários;
- escolher quais linhas serão processadas;
- indicar a coluna que contém o e-mail do destinatário;
- personalizar assunto e corpo da mensagem com tags da planilha;
- usar editor simples ou HTML avançado;
- incluir assinatura em imagem;
- anexar arquivos fixos para todos os destinatários;
- anexar arquivos personalizados por linha;
- validar dados antes do envio real;
- abrir e-mails no Outlook para conferência;
- enviar e-mails reais pelo Outlook Desktop;
- gerar logs e resultados do processamento.

## Estrutura do projeto

A pasta principal do sistema deve ficar organizada assim:

```text
disparador-emails/
├── iniciar_sistema.bat
├── README.md
├── Arquivo/
│   ├── Planilhas/
│   ├── Anexos_Personalizados/
│   ├── Anexos_Fixos/
│   ├── Assinatura/
│   └── Logs/
└── sistema/
    ├── app_streamlit.py
    ├── core.py
    ├── config.py
    ├── excel_modelo.py
    ├── style.py
    ├── ui_helpers.py
    └── requirements.txt
```

## Arquivos principais

### iniciar_sistema.bat

Arquivo usado pelo usuário para abrir o sistema.

Ele verifica o Python, cria ou reutiliza o ambiente virtual, instala os pacotes necessários, tenta abrir o Outlook Classic e inicia a interface do Streamlit no navegador.

### README.md

Manual de uso e documentação geral do sistema.

### sistema/app_streamlit.py

Arquivo principal da interface visual.

Ele organiza os passos do sistema, exibe os campos para o usuário, carrega a planilha, apresenta a pré-visualização e chama as funções de processamento.

### sistema/core.py

Arquivo responsável pela lógica principal de processamento.

Ele concentra validações, leitura de anexos, renderização de tags, integração com Outlook e geração das informações de resultado.

### sistema/config.py

Arquivo com os caminhos principais do projeto.

Ele define onde ficam as pastas de planilhas, anexos, assinatura e logs.

### sistema/excel_modelo.py

Arquivo responsável por criar ou copiar a planilha modelo.

Se `Planilha_Modelo.xlsx` ainda não existir, o sistema cria uma nova planilha modelo.

Se ela já existir, o sistema cria uma cópia numerada, como:

```text
Planilha_Modelo (2).xlsx
Planilha_Modelo (3).xlsx
```

### sistema/style.py

Arquivo reservado para o estilo visual global da interface.

### sistema/ui_helpers.py

Arquivo reservado para futuras funções auxiliares de interface.

### sistema/requirements.txt

Arquivo com os pacotes Python necessários para execução do sistema.

## Pasta Arquivo

A pasta `Arquivo` é a área operacional do sistema.

Ela é criada automaticamente, caso ainda não exista.

### Arquivo/Planilhas

Local onde devem ser colocadas as planilhas Excel usadas como base de dados.

O sistema procura arquivos `.xlsx` nessa pasta.

A planilha deve conter, no mínimo:

- uma coluna que indique se a linha deve ser processada;
- uma coluna com o e-mail do destinatário.

Exemplos de colunas esperadas:

```text
Enviar
Email_Destino
Nome
Setor
Anexo_Personalizado
```

### Arquivo/Anexos_Personalizados

Local para anexos diferentes por pessoa.

O nome do arquivo deve estar informado na planilha.

Exemplo:

```text
relatorio_maria.pdf
```

Para mais de um anexo na mesma linha, separe os nomes com ponto e vírgula:

```text
relatorio_maria.pdf; comprovante_maria.pdf
```

### Arquivo/Anexos_Fixos

Local para arquivos que serão enviados para todos os destinatários processados.

Exemplos:

- comunicado geral;
- PDF institucional;
- manual;
- imagem;
- documento padrão.

### Arquivo/Assinatura

Local para imagem de assinatura.

Formatos recomendados:

```text
.png
.jpg
.jpeg
.gif
```

### Arquivo/Logs

Local onde o sistema salva os arquivos de resultado e logs após o processamento.

Esses arquivos ajudam a verificar:

- linhas processadas;
- e-mails validados;
- e-mails exibidos no Outlook;
- e-mails enviados;
- anexos utilizados;
- anexos ausentes;
- erros encontrados;
- data e hora do processamento.

## Modos de processamento

O sistema possui três modos principais.

### Modo Validação

Apenas simula e valida no sistema.

Não interage com o Outlook.

Esse modo é recomendado para testar planilha, e-mails, anexos e conteúdo antes de abrir ou enviar mensagens.

### Modo Visualização

Abre a janela do e-mail no Outlook Desktop para conferência.

Não envia automaticamente.

Esse modo é recomendado para revisar visualmente poucas mensagens antes do disparo real.

### Modo Envio real

Realiza o envio oficial pelo Outlook Desktop.

Use somente depois de conferir:

- conta ativa do Outlook;
- assunto da mensagem;
- corpo da mensagem;
- pré-visualização;
- quantidade de linhas marcadas;
- anexos fixos;
- anexos personalizados;
- confirmação de envio real.

## Fluxo recomendado de uso

1. Abra o arquivo `iniciar_sistema.bat`.
2. Aguarde o sistema abrir no navegador.
3. Mantenha aberta a janela preta do inicializador enquanto estiver usando o sistema.
4. Confirme que o Outlook Classic está aberto.
5. Confirme que a conta correta está logada no Outlook.
6. Coloque a planilha Excel em `Arquivo/Planilhas`.
7. Se necessário, clique em `Criar planilha modelo`.
8. Escolha a planilha e a aba.
9. Selecione a coluna de envio.
10. Selecione a coluna de e-mail.
11. Configure anexos, se houver.
12. Revise o assunto.
13. Escreva ou revise a mensagem.
14. Atualize a pré-visualização.
15. Marque a mensagem como revisada.
16. Escolha o modo de processamento.
17. Execute o processo.
18. Confira os logs gerados.

## Atenções importantes durante o uso

### Não feche a janela preta do inicializador

Enquanto o sistema estiver em uso, mantenha aberta a janela preta do inicializador.

Essa janela mantém o servidor local do Streamlit em execução. Se ela for fechada, o aplicativo no navegador será encerrado ou perderá conexão.

Para encerrar o sistema corretamente, finalize o uso no navegador e depois feche a janela preta.

### Não edite a planilha enquanto ela estiver aberta no sistema

Evite editar, renomear, mover ou deixar aberta no Excel a planilha que está sendo usada pelo aplicativo.

Se a planilha estiver aberta no Excel, na pré-visualização do Windows ou sendo editada ao mesmo tempo em que o sistema tenta ler, copiar ou salvar dados, podem ocorrer erros de acesso ao arquivo.

Recomendação:

1. Feche a planilha no Excel antes de carregar no sistema.
2. Faça alterações na planilha antes de iniciar o processamento.
3. Se precisar editar novamente, pare o fluxo no sistema, feche a planilha no app se necessário, edite no Excel, salve, feche o Excel e depois clique em atualizar no sistema.

## Uso de tags na mensagem

As tags permitem personalizar a mensagem com dados da planilha.

Exemplo de planilha:

| Nome | Setor | Email_Destino |
|---|---|---|
| Maria | Financeiro | maria@dominio.com |

Exemplo de mensagem:

```text
Olá {Nome},

Esta mensagem é referente ao setor {Setor}.

Atenciosamente,
Equipe responsável
```

Resultado esperado:

```text
Olá Maria,

Esta mensagem é referente ao setor Financeiro.

Atenciosamente,
Equipe responsável
```

## Planilha modelo

O sistema pode criar uma planilha modelo para facilitar o preenchimento.

A planilha modelo contém colunas básicas:

```text
Enviar
Email_Destino
Nome
Setor
Anexo_Personalizado
```

A coluna `Enviar` indica quais linhas serão processadas.

Valores aceitos como marcação positiva:

```text
Sim
S
X
1
True
Enviar
OK
```

## Requisitos

- Windows.
- Python 3 instalado e disponível no PATH.
- Microsoft Outlook Desktop Classic instalado.
- Conta de e-mail configurada no Outlook.
- Acesso à internet para instalação inicial dos pacotes.
- Permissão para executar arquivos `.bat`.
- Pacotes listados em `sistema/requirements.txt`.

## Como iniciar o sistema

Na pasta principal, dê duplo clique em:

```text
iniciar_sistema.bat
```

O inicializador realiza as seguintes etapas:

1. verifica os arquivos principais;
2. verifica se o Python está instalado;
3. cria ou reutiliza o ambiente virtual;
4. instala ou atualiza os pacotes;
5. tenta abrir o Outlook Classic;
6. abre a interface Streamlit no navegador.

## Como rodar manualmente

Caso seja necessário executar manualmente a partir da pasta principal:

```bash
python -m venv sistema\venv
sistema\venv\Scripts\activate
pip install -r sistema\requirements.txt
streamlit run sistema\app_streamlit.py
```

## Segurança antes do envio real

Antes de executar envio real, confira:

- se o Outlook Classic está aberto;
- se a conta correta está ativa;
- se a planilha correta foi selecionada;
- se a coluna de envio está correta;
- se a coluna de e-mail está correta;
- se o assunto foi revisado;
- se o corpo da mensagem foi revisado;
- se a pré-visualização está correta;
- se os anexos foram encontrados;
- se o modo selecionado é realmente `Modo Envio real`;
- se a confirmação de envio real foi marcada.

O sistema reduz riscos, mas a revisão final é responsabilidade de quem opera o disparo.

## Boas práticas

- Teste primeiro em `Modo Validação`.
- Depois use `Modo Visualização` com poucas linhas.
- Use `Modo Envio real` apenas após revisar tudo.
- Mantenha nomes de arquivos simples.
- Evite acentos e caracteres especiais em nomes de anexos.
- Não altere nomes de colunas depois de montar a mensagem.
- Feche a planilha no Excel antes de carregar ou processar no sistema.
- Confira os logs após cada processamento.
- Não versionar planilhas, logs ou anexos com dados reais em repositórios públicos.

## Observações técnicas

A aplicação foi desenhada para execução local em Windows.

A integração com envio depende do Outlook Desktop Classic e da biblioteca `pywin32`.

O envio utiliza a conta ativa configurada no Outlook do computador.

## Licença e uso

Este projeto foi criado para uso interno e operacional.

Ajustes e melhorias podem ser feitos conforme a necessidade da equipe responsável.