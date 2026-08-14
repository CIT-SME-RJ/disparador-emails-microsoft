import html
import json
import re
from pathlib import Path
from datetime import datetime

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

from core import (
    processar_envios,
    valor_verdadeiro,
    renderizar_html
)


st.set_page_config(
    page_title="Disparador Outlook",
    page_icon="📧",
    layout="wide"
)


# =========================
# CONFIGURAÇÃO DAS PASTAS
# =========================

BASE_DIR = Path.cwd() / "Arquivo"

PASTA_PLANILHAS = BASE_DIR / "Planilhas"
PASTA_ANEXOS_DINAMICOS = BASE_DIR / "Anexos_Personalizados"
PASTA_ANEXOS_FIXOS = BASE_DIR / "Anexos_Fixos"
PASTA_TEMPLATES = BASE_DIR / "Templates"
PASTA_LOGS = BASE_DIR / "Logs"

for pasta in [
    PASTA_PLANILHAS,
    PASTA_ANEXOS_DINAMICOS,
    PASTA_ANEXOS_FIXOS,
    PASTA_TEMPLATES,
    PASTA_LOGS
]:
    pasta.mkdir(parents=True, exist_ok=True)


def encontrar_coluna_preferida(colunas, opcoes):
    """
    Tenta encontrar uma coluna pelo nome mais provável.
    Caso não encontre, retorna o índice 0.
    """
    colunas_normalizadas = {
        str(col).strip().lower(): i
        for i, col in enumerate(colunas)
    }

    for opcao in opcoes:
        chave = opcao.strip().lower()

        if chave in colunas_normalizadas:
            return colunas_normalizadas[chave]

    return 0


def carregar_template_html(caminho_template):
    """
    Lê um arquivo HTML em UTF-8.
    """
    with open(caminho_template, "r", encoding="utf-8") as arquivo:
        return arquivo.read()


def texto_simples_para_html(texto):
    """
    Converte texto simples em HTML básico.

    Permite:
    - parágrafos com linha em branco
    - quebra de linha simples
    - negrito com **texto**
    - variáveis no formato {NomeDaColuna}

    Exemplo:
    Olá **{Nome}**
    """
    if not texto or not texto.strip():
        return ""

    def aplicar_formatacao_basica(texto_linha):
        texto_linha = html.escape(texto_linha)

        texto_linha = re.sub(
            r"\*\*(.+?)\*\*",
            r"<strong>\1</strong>",
            texto_linha
        )

        return texto_linha

    blocos = texto.strip().split("\n\n")
    paragrafos_html = []

    for bloco in blocos:
        linhas = bloco.splitlines()

        linhas_formatadas = [
            aplicar_formatacao_basica(linha)
            for linha in linhas
        ]

        conteudo_paragrafo = "<br>".join(linhas_formatadas)
        paragrafos_html.append(f"<p>{conteudo_paragrafo}</p>")

    return "\n".join(paragrafos_html)


def montar_preview_html(html_preview):
    """
    Monta uma pré-visualização com fundo branco.
    Isso evita que o modo escuro do Streamlit deixe a prévia ilegível.
    """
    return f"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <style>
        html, body {{
            background-color: #ffffff !important;
            color: #111827 !important;
            margin: 0;
            padding: 0;
            font-family: Arial, Helvetica, sans-serif;
        }}

        .email-preview-container {{
            background-color: #ffffff !important;
            color: #111827 !important;
            padding: 24px;
            border: 1px solid #e5e7eb;
            border-radius: 12px;
            line-height: 1.5;
            font-size: 15px;
            box-sizing: border-box;
            width: 100%;
        }}

        .email-preview-container * {{
            max-width: 100%;
            box-sizing: border-box;
        }}

        .email-preview-container p {{
            margin-top: 0;
            margin-bottom: 12px;
        }}

        .email-preview-container a {{
            color: #2563eb !important;
        }}

        .email-preview-container table {{
            border-collapse: collapse;
            width: 100%;
        }}

        .email-preview-container td,
        .email-preview-container th {{
            border: 1px solid #d1d5db;
            padding: 6px;
        }}
    </style>
</head>
<body>
    <div class="email-preview-container">
        {html_preview}
    </div>
</body>
</html>
"""


def mostrar_pastas():
    st.code(
        f"""Arquivo principal do sistema:
{BASE_DIR}

Planilhas:
{PASTA_PLANILHAS}

Anexos personalizados:
{PASTA_ANEXOS_DINAMICOS}

Anexos fixos:
{PASTA_ANEXOS_FIXOS}

Templates HTML:
{PASTA_TEMPLATES}

Logs:
{PASTA_LOGS}
""",
        language="text"
    )

def painel_tags_compacto(colunas, contexto="mensagem", mostrar_negrito=True):
    """
    Mostra as tags disponíveis em um painel compacto, com botão de copiar.
    O usuário copia a tag e cola no ponto desejado da mensagem.
    """
    itens = []

    for coluna in colunas:
        tag = f"{{{coluna}}}"
        itens.append({
            "texto": tag,
            "tipo": "tag"
        })

    if mostrar_negrito:
        itens.append({
            "texto": "**texto em negrito**",
            "tipo": "negrito"
        })

    itens_json = json.dumps(itens, ensure_ascii=False)

    altura_base = 120
    altura_por_linha = 34
    quantidade_linhas = max(1, (len(itens) + 1) // 2)
    altura_total = altura_base + (quantidade_linhas * altura_por_linha)

    html_painel = f"""
    <div class="painel-tags">
        <div class="titulo">Tags disponíveis</div>
        <div class="subtitulo">Clique no ícone para copiar. Depois cole no ponto desejado da {contexto}.</div>

        <div id="grid-tags" class="grid-tags"></div>
    </div>

    <script>
        const itens = {itens_json};

        function copiarTexto(texto, botao) {{
            navigator.clipboard.writeText(texto).then(function() {{
                const textoOriginal = botao.innerText;
                botao.innerText = "✓";
                botao.classList.add("copiado");

                setTimeout(function() {{
                    botao.innerText = textoOriginal;
                    botao.classList.remove("copiado");
                }}, 900);
            }});
        }}

        const grid = document.getElementById("grid-tags");

        itens.forEach(function(item) {{
            const chip = document.createElement("div");
            chip.className = item.tipo === "negrito" ? "chip chip-negrito" : "chip";

            const texto = document.createElement("span");
            texto.className = "chip-texto";
            texto.innerText = item.texto;

            const botao = document.createElement("button");
            botao.className = "botao-copiar";
            botao.innerText = "📋";
            botao.title = "Copiar";
            botao.onclick = function() {{
                copiarTexto(item.texto, botao);
            }};

            chip.appendChild(texto);
            chip.appendChild(botao);
            grid.appendChild(chip);
        }});
    </script>

    <style>
        html, body {{
            margin: 0;
            padding: 0;
            background: transparent;
            font-family: Arial, Helvetica, sans-serif;
        }}

        /* =========================
        TEMA CLARO - padrão
        ========================= */

        .painel-tags {{
            width: 100%;
            box-sizing: border-box;
        }}

        .titulo {{
            font-size: 18px;
            font-weight: 700;
            color: #111827;
            margin-bottom: 4px;
        }}

        .subtitulo {{
            font-size: 12px;
            color: #4b5563;
            line-height: 1.25;
            margin-bottom: 10px;
        }}

                .grid-tags {{
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 6px 8px;
            align-items: start;
        }}

        .chip {{
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 4px;
            min-width: 0;
            background-color: #f9fafb;
            border: 1px solid #d1d5db;
            border-radius: 7px;
            padding: 3px 4px 3px 6px;
            box-sizing: border-box;
        }}

        .chip-negrito {{
                    grid-column: 1 / -1;
                    margin-top: 8px;
                }}

        .chip-texto {{
            min-width: 0;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
            font-family: Consolas, "Courier New", monospace;
            font-size: 11px;
            line-height: 1.2;
            color: #111827;
            user-select: text;
        }}

        .botao-copiar {{
            border: none;
            background: transparent;
            color: #374151;
            cursor: pointer;
            font-size: 12px;
            line-height: 1;
            padding: 1px 2px;
            border-radius: 4px;
            flex-shrink: 0;
        }}

        .botao-copiar:hover {{
            background-color: rgba(148, 163, 184, 0.22);
            color: #111827;
        }}

        .botao-copiar.copiado {{
            color: #16a34a;
            font-weight: 700;
        }}

        @media (prefers-color-scheme: dark) {{
            .titulo {{
                color: #f9fafb;
            }}

            .subtitulo {{
                color: #d1d5db;
            }}

            .chip {{
                background-color: rgba(255, 255, 255, 0.06);
                border: 1px solid rgba(255, 255, 255, 0.18);
            }}

            .chip-texto {{
                color: #f3f4f6;
            }}

            .botao-copiar {{
                color: #d1d5db;
            }}

            .botao-copiar:hover {{
                background-color: rgba(255, 255, 255, 0.12);
                color: #ffffff;
            }}

            .botao-copiar.copiado {{
                color: #22c55e;
            }}
        }}
    </style>
    """

    components.html(
        html_painel,
        height=altura_total,
        scrolling=False
    )


st.title("📧 Disparador de E-mails com Outlook Desktop")
st.markdown("Siga o passo a passo abaixo para realizar seus envios com mais segurança.")
st.divider()


with st.expander("📁 Ver estrutura de pastas usada pelo sistema", expanded=False):
    mostrar_pastas()


# =========================
# PASSO 1: VERIFICAÇÃO
# =========================

st.header("Passo 1: Verificação de Segurança")

st.write("Antes de começar, confirme as condições abaixo:")

check_outlook_aberto = st.checkbox(
    "O aplicativo Outlook Classic Desktop está aberto no meu computador."
)

check_conta_correta = st.checkbox(
    "A conta logada no Outlook é a conta oficial que deve enviar essas mensagens."
)

if not (check_outlook_aberto and check_conta_correta):
    st.warning("⚠️ Marque as duas verificações acima para desbloquear o próximo passo.")
    st.stop()

st.success("✅ Verificações concluídas.")
st.divider()


# =========================
# PASSO 2: PLANILHA
# =========================

st.header("Passo 2: Base de Dados")

st.info(
    f"📁 Coloque sua planilha Excel na pasta abaixo:\n\n`{PASTA_PLANILHAS}`"
)

planilhas = sorted(list(PASTA_PLANILHAS.glob("*.xlsx")))

if not planilhas:
    st.warning("⚠️ Nenhuma planilha .xlsx encontrada.")
    if st.button("🔄 Atualizar pasta"):
        st.rerun()
    st.stop()

arquivo_excel = st.selectbox(
    "Selecione a planilha:",
    planilhas,
    format_func=lambda caminho: caminho.name
)

try:
    xls = pd.ExcelFile(arquivo_excel, engine="openpyxl")

    aba = st.selectbox(
        "Selecione a aba da planilha:",
        xls.sheet_names
    )

    df = pd.read_excel(
        arquivo_excel,
        sheet_name=aba,
        engine="openpyxl"
    )

except Exception as e:
    st.error(f"Erro ao ler a planilha: {e}")
    st.stop()

if df.empty:
    st.warning("A aba selecionada está vazia.")
    st.stop()

colunas = list(df.columns)

if not colunas:
    st.error("Não foram encontradas colunas na planilha.")
    st.stop()

st.subheader("Prévia da planilha")
st.dataframe(df.head(10), use_container_width=True)

coluna_1, coluna_2 = st.columns(2)

with coluna_1:
    indice_col_enviar = encontrar_coluna_preferida(
        colunas,
        ["Enviar", "Envio", "Disparar", "Mandar"]
    )

    col_enviar = st.selectbox(
        "Qual coluna indica se a linha deve ser enviada?",
        colunas,
        index=indice_col_enviar
    )

with coluna_2:
    indice_col_email = encontrar_coluna_preferida(
        colunas,
        ["Email", "E-mail", "E_mail", "Destinatario", "Destinatário"]
    )

    col_email = st.selectbox(
        "Qual coluna contém o e-mail do destinatário?",
        colunas,
        index=indice_col_email
    )

total_marcados = int(df[col_enviar].apply(valor_verdadeiro).sum())

if total_marcados == 0:
    st.error(
        "⚠️ Nenhuma linha está marcada para envio ou validação. "
        "Verifique a coluna escolhida como indicador de envio."
    )
    st.stop()

st.success(f"✅ Planilha carregada. Há {total_marcados} linha(s) marcada(s) para processamento.")
st.divider()


# =========================
# PASSO 3: ANEXOS
# =========================

st.header("Passo 3: Anexos")

tipo_anexo = st.radio(
    "Essa mensagem terá anexos?",
    [
        "Não, apenas texto.",
        "Anexos fixos: o mesmo arquivo para todos.",
        "Anexos personalizados: arquivos diferentes por pessoa.",
        "Anexos fixos + personalizados."
    ]
)

col_anexos = None
usar_anexos_fixos = False
usar_anexos_personalizados = False

if tipo_anexo in [
    "Anexos fixos: o mesmo arquivo para todos.",
    "Anexos fixos + personalizados."
]:
    usar_anexos_fixos = True

    st.info(
        f"📁 Coloque os arquivos que todos receberão na pasta abaixo:\n\n`{PASTA_ANEXOS_FIXOS}`"
    )

    arquivos_fixos = sorted(list(PASTA_ANEXOS_FIXOS.glob("*.*")))

    if not arquivos_fixos:
        st.warning("⚠️ Você escolheu anexos fixos, mas nenhum arquivo foi encontrado nessa pasta.")
        if st.button("🔄 Atualizar anexos fixos"):
            st.rerun()
        st.stop()

    st.write("**Arquivos fixos encontrados:**")

    for arquivo in arquivos_fixos:
        st.write(f"- {arquivo.name}")

if tipo_anexo in [
    "Anexos personalizados: arquivos diferentes por pessoa.",
    "Anexos fixos + personalizados."
]:
    usar_anexos_personalizados = True

    st.info(
        f"📁 Coloque todos os arquivos personalizados na pasta abaixo:\n\n`{PASTA_ANEXOS_DINAMICOS}`"
    )

    st.markdown(
        """
        **Regra da planilha para anexos personalizados:**

        A planilha precisa ter uma coluna com o nome exato do arquivo, incluindo a extensão.

        Exemplos:

        - `relatorio_joao.pdf`
        - `boleto.pdf; recibo.pdf`
        """
    )

    opcoes_anexo = ["Selecione..."] + colunas

    escolha_col_anexos = st.selectbox(
        "Qual coluna da planilha contém os nomes dos anexos personalizados?",
        opcoes_anexo
    )

    if escolha_col_anexos == "Selecione...":
        st.warning("Selecione a coluna de anexos personalizados para continuar.")
        st.stop()

    col_anexos = escolha_col_anexos

st.success("✅ Configuração de anexos registrada.")
st.divider()


# =========================
# PASSO 4: CORPO DO E-MAIL
# =========================

st.header("Passo 4: Corpo do E-mail")

st.caption(
    "Use o editor simples para mensagens comuns. Use HTML avançado apenas para modelos com layout, tabelas, cores ou assinatura formatada."
)

modo_editor = st.radio(
    "Formato da mensagem:",
    [
        "Editor simples, sem HTML",
        "HTML avançado"
    ],
    index=0
)

with st.expander("📌 Dicas de personalização", expanded=False):
    st.markdown(
        """
        Use tags da planilha para personalizar a mensagem.

        **Exemplo:**

        ```text
        Olá {Nome},

        Estamos entrando em contato sobre a unidade {Unidade}.
        Atenção aos **prazos**.
        ```

        **Regras rápidas:**

        - `{Nome}` será trocado pelo valor da coluna `Nome`.
        - `{Unidade}` será trocado pelo valor da coluna `Unidade`.
        - `**texto**` aparecerá em negrito no editor simples.
        """
    )

# =========================
# CONTROLE DE PRÉ-VISUALIZAÇÃO E MENSAGEM PRONTA
# =========================

if "template_html_preview" not in st.session_state:
    st.session_state["template_html_preview"] = ""

if "template_html_aprovado" not in st.session_state:
    st.session_state["template_html_aprovado"] = ""

if "mensagem_pronta_anterior" not in st.session_state:
    st.session_state["mensagem_pronta_anterior"] = False

if "mensagem_pronta_check" not in st.session_state:
    st.session_state["mensagem_pronta_check"] = False

# =========================
# EDITOR DA MENSAGEM
# =========================

if modo_editor == "Editor simples, sem HTML":
    texto_padrao = """
Olá {Nome},

Esta é uma mensagem automática.

Atenciosamente,
Equipe responsável
""".strip()

    col_editor, col_preview = st.columns([1.15, 1])

    with col_editor:
        st.markdown("### ✍️ Digite o corpo do e-mail:")

        texto_corpo = st.text_area(
            "Corpo do e-mail",
            label_visibility="collapsed",
            value=texto_padrao,
            height=450,
            help="Use variáveis como {Nome}, {Unidade}, {Email}. Para negrito, use **texto**."
        )

    template_html = texto_simples_para_html(texto_corpo)

    with col_preview:
        st.markdown("### 👀 Pré-visualização")
        area_preview = st.empty()

    col_tags, col_botao_preview = st.columns([2, 1])

    with col_tags:
        with st.expander("📌 Tags disponíveis", expanded=True):
            painel_tags_compacto(
                colunas=colunas,
                contexto="mensagem",
                mostrar_negrito=True
            )

    with col_botao_preview:
        st.markdown("#### Pré-visualização")
        st.caption("Depois de alterar o texto, atualize a prévia.")

        if st.button("🔄 Atualizar pré-visualização", use_container_width=True):
            st.session_state["template_html_preview"] = template_html
            st.success("Pré-visualização atualizada.")

    try:
        df_marcados = df[df[col_enviar].apply(valor_verdadeiro)]

        if not df_marcados.empty:
            linha_preview = df_marcados.iloc[0]

            html_base_preview = st.session_state["template_html_preview"] or template_html

            html_preview = renderizar_html(
                html_base_preview,
                linha_preview
            )

            with area_preview:
                components.html(
                    montar_preview_html(html_preview),
                    height=450,
                    scrolling=True
                )

    except Exception as e:
        st.error(f"Erro ao gerar prévia: {e}")

else:
    st.info(
        f"📁 Coloque seu arquivo `.html` na pasta abaixo:\n\n`{PASTA_TEMPLATES}`"
    )

    templates = sorted(list(PASTA_TEMPLATES.glob("*.html")))

    html_padrao = """
<p>Olá <strong>{Nome}</strong>,</p>

<p>Esta é uma mensagem automática.</p>

<p>Atenciosamente,</p>
<p>Equipe responsável</p>
""".strip()

    if not templates:
        st.warning(
            "⚠️ Nenhum arquivo HTML foi encontrado. "
            "Será usado um modelo temporário editável abaixo."
        )

        template_html_inicial = html_padrao

    else:
        template_selecionado = st.selectbox(
            "Selecione o template HTML:",
            templates,
            format_func=lambda caminho: caminho.name
        )

        try:
            template_html_inicial = carregar_template_html(template_selecionado)
        except Exception as e:
            st.error(f"Erro ao ler o template HTML: {e}")
            st.stop()

    col_editor, col_variaveis = st.columns([4, 1])

    with col_editor:
        template_html = st.text_area(
            "Revise ou ajuste o HTML do e-mail:",
            value=template_html_inicial,
            height=320
        )

    with col_variaveis:
        st.markdown("**Tags**")
        st.caption("Copie e cole no HTML.")

        with st.container(height=240):
            for coluna in colunas:
                st.code(f"{{{coluna}}}", language="text")


if not template_html.strip():
    st.error("O corpo do e-mail está vazio. Preencha o conteúdo antes de continuar.")
    st.stop()

if not st.session_state["template_html_preview"]:
    st.session_state["template_html_preview"] = template_html

# =========================
# CONFERÊNCIA DA MENSAGEM
# =========================

st.markdown("### Conferência da mensagem")

st.caption(
    "Revise a mensagem na pré-visualização. "
    "Ao marcar como pronta, a pré-visualização será atualizada com o conteúdo atual."
)

mensagem_pronta = st.checkbox(
    "Mensagem revisada e pronta para validação ou envio.",
    key="mensagem_pronta_check"
)

mudou_para_pronta = mensagem_pronta and not st.session_state["mensagem_pronta_anterior"]

if mudou_para_pronta:
    st.session_state["template_html_preview"] = template_html
    st.session_state["template_html_aprovado"] = template_html
    st.success("✅ Mensagem marcada como pronta. Revise a pré-visualização.")

mensagem_liberada = False

if mensagem_pronta:
    if st.session_state["template_html_aprovado"] == template_html:
        mensagem_liberada = True
        st.success("✅ Mensagem pronta. O próximo passo foi liberado.")
    else:
        st.warning(
            "⚠️ A mensagem foi alterada depois de ter sido marcada como pronta. "
            "Para liberar o próximo passo, desmarque e marque novamente a confirmação."
        )
        mensagem_liberada = False
else:
    st.info(
        "Marque a confirmação acima para liberar o próximo passo. "
        "Ao marcar, a pré-visualização será atualizada com o conteúdo atual."
    )

st.session_state["mensagem_pronta_anterior"] = mensagem_pronta

st.divider()

# =========================
# PASSO 5: DISPARO
# =========================

if not mensagem_liberada:
    st.info("🔒 O Passo 5 será liberado depois que a mensagem for revisada e marcada como pronta.")
    st.stop()

st.header("Passo 5: Validação ou Disparo")

assunto = st.text_input(
    "Assunto do e-mail:",
    value="Mensagem Importante"
)

modo = st.radio(
    "O que deseja fazer agora?",
    [
        "Somente validar",
        "Enviar de verdade pelo Outlook"
    ],
    index=0
)

with st.expander("⚙️ Configurações avançadas", expanded=True):
    coluna_adv_1, coluna_adv_2 = st.columns(2)

    with coluna_adv_1:
        intervalo = st.number_input(
            "Intervalo entre e-mails, em segundos",
            min_value=0.0,
            value=1.5,
            step=0.5,
            help="Ajuda a evitar disparos muito rápidos pelo Outlook."
        )

        limite = st.number_input(
            "Limite de linhas processadas",
            min_value=0,
            value=0,
            step=1,
            help="Use 0 para processar todas as linhas marcadas."
        )

    with coluna_adv_2:
        permitir_envio_sem_anexo = st.checkbox(
            "Permitir envio mesmo se algum anexo personalizado estiver faltando",
            value=False,
            help="Por segurança, o padrão é bloquear o envio quando algum anexo personalizado estiver ausente."
        )

enviar_real = modo == "Enviar de verdade pelo Outlook"
somente_validar = not enviar_real

confirmacao_envio = True

if enviar_real:
    st.error(
        "🚨 Atenção: os e-mails serão enviados de verdade pelo Outlook. "
        "Revise planilha, assunto, corpo do e-mail e anexos antes de continuar."
    )

    confirmacao_envio = st.checkbox(
        "Tenho certeza que revisei tudo e autorizo o disparo real."
    )

botao_processar = st.button(
    "🚀 INICIAR PROCESSO",
    type="primary",
    use_container_width=True
)

if botao_processar:
    if enviar_real and not confirmacao_envio:
        st.warning("Marque a confirmação para autorizar o envio real.")
        st.stop()

    area_status = st.empty()
    barra = st.progress(0)
    logs_streamlit = []

    total_para_barra = total_marcados

    if limite and limite > 0:
        total_para_barra = min(total_marcados, int(limite))

    def atualizar_tela(log):
        logs_streamlit.append(log)

        progresso = min(len(logs_streamlit) / total_para_barra, 1.0)

        barra.progress(progresso)

        area_status.info(
            f"Processando {len(logs_streamlit)} de {total_para_barra} | "
            f"{log['STATUS']} | {log['EMAIL']}"
        )

    try:
        df_resultado, logs = processar_envios(
            df=df,
            template_html=template_html,
            pasta_anexos_dinamicos=str(PASTA_ANEXOS_DINAMICOS) if usar_anexos_personalizados else None,
            pasta_anexos_fixos=str(PASTA_ANEXOS_FIXOS) if usar_anexos_fixos else None,
            assunto=assunto,
            col_enviar=col_enviar,
            col_email=col_email,
            col_anexos=col_anexos,
            enviar_real=enviar_real,
            somente_validar=somente_validar,
            intervalo=float(intervalo),
            limite=int(limite),
            permitir_envio_sem_anexo=permitir_envio_sem_anexo,
            exigir_anexo_personalizado=usar_anexos_personalizados,
            callback=atualizar_tela
        )

        st.success("🎉 Processo finalizado.")

        df_logs = pd.DataFrame(logs)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        arquivo_resultado = PASTA_LOGS / f"Resultado_Envio_{timestamp}.xlsx"
        arquivo_logs = PASTA_LOGS / f"Logs_Envio_{timestamp}.xlsx"

        df_resultado.to_excel(arquivo_resultado, index=False)

        if not df_logs.empty:
            df_logs.to_excel(arquivo_logs, index=False)

        st.write(
            f"📁 Planilha atualizada com status, erros, horários e anexos salva em:\n\n`{arquivo_resultado}`"
        )

        if not df_logs.empty:
            st.write(
                f"📁 Log detalhado do processamento salvo em:\n\n`{arquivo_logs}`"
            )

        if df_logs.empty:
            st.warning("Nenhuma linha foi processada.")
            st.stop()

        quantidade_sucesso = df_logs["STATUS"].isin(["Enviado", "Validado"]).sum()
        quantidade_erros = len(df_logs) - quantidade_sucesso

        metrica_1, metrica_2, metrica_3 = st.columns(3)

        metrica_1.metric("Total processado", len(df_logs))
        metrica_2.metric("Sucesso", int(quantidade_sucesso))
        metrica_3.metric("Erros ou bloqueios", int(quantidade_erros))

        st.subheader("Resumo dos logs")
        st.dataframe(df_logs, use_container_width=True)

        st.subheader("Prévia da planilha atualizada")
        st.dataframe(df_resultado, use_container_width=True)

    except Exception as e:
        st.error(f"Ocorreu um erro crítico no processamento: {e}")