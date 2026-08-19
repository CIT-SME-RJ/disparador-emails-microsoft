import html
import json
import re
import base64
import mimetypes
from pathlib import Path
from datetime import datetime

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

from sistema.excel_modelo import criar_planilha_modelo
from sistema.style import aplicar_estilo_global

from sistema.core import (
    processar_envios,
    valor_verdadeiro,
    renderizar_html
)

from sistema.config import (
    PASTA_PROJETO,
    BASE_DIR,
    PASTA_PLANILHAS,
    PASTA_ANEXOS_DINAMICOS,
    PASTA_ANEXOS_FIXOS,
    PASTA_ASSINATURA,
    PASTA_LOGS,
    garantir_pastas
)

st.set_page_config(
    page_title="Disparador Outlook",
    page_icon="📧",
    layout="wide"
)

garantir_pastas()
aplicar_estilo_global()

def encontrar_coluna_preferida(colunas, opcoes):
    colunas_normalizadas = {
        str(col).strip().lower(): i
        for i, col in enumerate(colunas)
    }

    for opcao in opcoes:
        chave = opcao.strip().lower()

        if chave in colunas_normalizadas:
            return colunas_normalizadas[chave]

    return 0


def texto_simples_para_html(texto):
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


def preparar_html_preview_assinatura(
    template_html,
    caminho_assinatura
):
    if not caminho_assinatura:
        return template_html

    caminho_assinatura = Path(caminho_assinatura)

    tipo_mime, _ = mimetypes.guess_type(caminho_assinatura.name)

    if not tipo_mime:
        tipo_mime = "image/png"

    imagem_base64 = base64.b64encode(
        caminho_assinatura.read_bytes()
    ).decode("utf-8")

    imagem_data_uri = (
        f"data:{tipo_mime};base64,{imagem_base64}"
    )

    return template_html.replace(
        "cid:assinatura_img",
        imagem_data_uri
    )


def adicionar_assinatura_imagem_ao_html(
    template_html,
    usar_assinatura_imagem,
    caminho_assinatura,
    largura_imagem=320
):
    if not usar_assinatura_imagem or not caminho_assinatura:
        return template_html

    img_assinatura = (
        f'<img src="cid:assinatura_img" alt="Assinatura" width="{largura_imagem}" '
        f'style="display:block; border:0; outline:none; text-decoration:none;" />'
    )

    if '<img' in template_html and 'cid:assinatura_img' in template_html:
        return template_html

    if 'cid:assinatura_img' in template_html:
        return template_html.replace('cid:assinatura_img', img_assinatura)

    html_assinatura = f'<p style="margin-top:16px; margin-bottom:0;">\n    {img_assinatura}\n</p>'

    return f"{template_html.rstrip()}\n{html_assinatura}"


def adicionar_banner_teste_ao_html(template_html, modo_teste=True):
    if not modo_teste:
        return template_html

    banner_html = """
<div style="background-color: #fff3cd; border: 2px dashed #ffc107; color: #856404; padding: 12px 16px; margin-bottom: 20px; border-radius: 8px; font-family: Arial, sans-serif; text-align: center;">
    <strong style="font-size: 15px; text-transform: uppercase;">⚠️ MODO DE TESTE / RASCUNHO SIMULADO</strong><br>
    <span style="font-size: 12px; color: #664d03;">Este e-mail é apenas uma validação/teste e não representa um disparo oficial ao cliente.</span>
</div>
""".strip()

    return f"{banner_html}\n{template_html}"


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

Assinatura em imagem:
{PASTA_ASSINATURA}

Logs:
{PASTA_LOGS}
""",
        language="text"
    )


def painel_tags_compacto(colunas, contexto="mensagem", mostrar_negrito=True):
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
            .titulo {{ color: #f9fafb; }}
            .subtitulo {{ color: #d1d5db; }}
            .chip {{
                background-color: rgba(255, 255, 255, 0.06);
                border: 1px solid rgba(255, 255, 255, 0.18);
            }}
            .chip-texto {{ color: #f3f4f6; }}
            .botao-copiar {{ color: #d1d5db; }}
            .botao-copiar:hover {{
                background-color: rgba(255, 255, 255, 0.12);
                color: #ffffff;
            }}
            .botao-copiar.copiado {{ color: #22c55e; }}
        }}
    </style>
    """

    components.html(
        html_painel,
        height=altura_total,
        scrolling=False
    )


st.title("📧 Disparador de E-mails com Outlook Desktop")
st.markdown("##### Siga o passo a passo abaixo para realizar seus envios com mais segurança.")
st.divider()

README_PATH = PASTA_PROJETO / "README.md"

with st.expander("📘 Abrir manual do sistema / README", expanded=False):
    if README_PATH.exists():
        conteudo_readme = README_PATH.read_text(encoding="utf-8")
        st.markdown(conteudo_readme)
    else:
        st.warning(
            "⚠️ O arquivo README.md não foi encontrado na pasta principal do sistema."
        )

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
    f"📁 Coloque sua planilha Excel nesta pasta:\n\n`{PASTA_PLANILHAS}`"
)

st.caption(
    "💡 Dica: essa pasta é criada automaticamente dentro da pasta do programa. "
    "Procure por Arquivo > Planilhas."
)

if st.button("📄 Criar planilha modelo"):
    try:
        resultado_modelo = criar_planilha_modelo()

        if resultado_modelo["tipo"] == "nova":
            st.success(f"✅ {resultado_modelo['mensagem']}")
        else:
            st.info(f"📄 {resultado_modelo['mensagem']}")

        st.rerun()

    except PermissionError:
        st.warning(
            "⚠️ Não foi possível criar ou copiar a planilha modelo. "
            "Feche o arquivo no Excel ou na pré-visualização do Windows e tente novamente."
        )

    except Exception as e:
        st.error(f"Erro ao criar a planilha modelo: {e}")

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
        ["Email_Destino", "Email", "E-mail", "E_mail", "email", "e-mail", "e_mail", "Destinatario", "Destinatário"]
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

st.header("Passo 4: Montagem do E-mail")

# ASSUNTO DA MENSAGEM
st.subheader("Assunto da mensagem")

assunto = st.text_input(
    "Digite o assunto do e-mail:",
    value="Mensagem Importante",
    help="Você também pode usar tags da planilha, como {Nome}, {Setor} ou {Email_Destino}."
)

st.caption(
    "💡 Dica: o assunto também pode ser personalizado com as tags da planilha."
)

assunto_conferido = st.checkbox(
    "Conferi o assunto da mensagem e ele está correto."
)

if not assunto.strip():
    st.error("O assunto do e-mail está vazio.")
    st.stop()

if not assunto_conferido:
    st.warning("⚠️ Confira o assunto da mensagem antes de continuar.")
    st.stop()

st.divider()


# ASSINATURA EM IMAGEM
st.subheader("Assinatura")

usar_assinatura_imagem = st.checkbox(
    "Desejo usar uma assinatura em imagem."
)

caminho_assinatura = None

if usar_assinatura_imagem:
    col_assinatura_config, col_assinatura_preview = st.columns([1.15, 1])

    with col_assinatura_config:
        st.markdown("###### Configuração da assinatura")

        extensoes_assinatura = {".png", ".jpg", ".jpeg", ".gif"}

        imagens_assinatura = sorted(
            arquivo
            for arquivo in PASTA_ASSINATURA.iterdir()
            if arquivo.is_file() and arquivo.suffix.lower() in extensoes_assinatura
        )

        if not imagens_assinatura:
            st.info(
                f"📁 Coloque a imagem da assinatura nesta pasta:\n\n"
                f"`{PASTA_ASSINATURA}`"
            )

            st.caption(
                "💡 Dica: use uma imagem em formato .png, .jpg, .jpeg ou .gif."
            )

            st.warning("⚠️ Nenhuma imagem encontrada na pasta de assinatura.")

            if st.button("🔄 Atualizar pasta de assinatura"):
                st.rerun()

            st.stop()

        if len(imagens_assinatura) == 1:
            caminho_assinatura = imagens_assinatura[0]

            st.success(
                f"✅ Imagem de assinatura encontrada: `{caminho_assinatura.name}`"
            )

        else:
            caminho_assinatura = st.selectbox(
                "Selecione a imagem de assinatura:",
                imagens_assinatura,
                format_func=lambda caminho: caminho.name
            )

            if caminho_assinatura:
                st.success(
                    f"✅ Imagem de assinatura selecionada: `{caminho_assinatura.name}`"
                )

        if st.button("🔄 Atualizar pasta de assinatura"):
            st.rerun()

    with col_assinatura_preview:
        st.markdown("###### Prévia da assinatura")

        if caminho_assinatura:
            st.image(
                str(caminho_assinatura),
                caption=caminho_assinatura.name,
                use_container_width=True
            )

else:
    st.caption("Nenhuma assinatura em imagem selecionada.")

st.divider()

modo_editor = st.radio(
    "Formato da mensagem:",
    [
        "Editor simples, sem HTML",
        "HTML avançado"
    ],
    index=0
)

if "template_html_preview" not in st.session_state:
    st.session_state["template_html_preview"] = ""

if "assunto_preview" not in st.session_state:
    st.session_state["assunto_preview"] = ""

if "template_html_aprovado" not in st.session_state:
    st.session_state["template_html_aprovado"] = ""

if "mensagem_pronta_anterior" not in st.session_state:
    st.session_state["mensagem_pronta_anterior"] = False

if "mensagem_pronta_check" not in st.session_state:
    st.session_state["mensagem_pronta_check"] = False


# APOIO PARA PREENCHIMENTO DA MENSAGEM
col_tags, col_dicas = st.columns([1, 1])

with col_dicas:
    with st.expander("💡 Dicas de preenchimento da mensagem", expanded=False):
        st.markdown("""
Dicas rápidas

- Use as tags da planilha para personalizar a mensagem.
- Exemplo: `Olá {Nome},` será trocado pelo nome da pessoa na linha da planilha.
- Cada coluna da planilha pode virar uma tag, desde que esteja escrita entre chaves.
- Exemplos de tags possíveis: `{Nome}`, `{Setor}`, `{Email_Destino}`.
- Evite alterar manualmente os nomes das tags. Elas precisam bater exatamente com o nome da coluna.
- Se usar anexos personalizados, informe o nome do arquivo na coluna `Anexo_Personalizado`.
- Para mais de um anexo personalizado na mesma linha, separe os nomes com ponto e vírgula: `arquivo1.pdf; arquivo2.pdf`.
- Para colocar uma palavra em negrito, use dois asteriscos antes e depois da palavra.
""")

with col_tags:
    with st.expander("📌 Tags disponíveis", expanded=True):
        painel_tags_compacto(colunas=colunas)


# EDITOR DA MENSAGEM
if modo_editor == "Editor simples, sem HTML":
    texto_padrao = """
Olá {Nome},

Esta é uma mensagem automática.

Atenciosamente,
**Equipe responsável**
""".strip()

    col_editor, col_preview = st.columns([1.15, 1])

    with col_editor:
        st.markdown("### ✍️ Digite o corpo do e-mail:")

        texto_corpo = st.text_area(
            "Corpo do e-mail",
            label_visibility="collapsed",
            value=texto_padrao,
            height=450
        )

    with col_preview:
        st.markdown("### 👀 Pré-visualização")

        area_assunto_preview = st.empty()
        area_preview = st.empty()
        area_botao_preview = st.empty()
        area_feedback_preview = st.empty()

        with area_botao_preview:
            botao_atualizar_preview = st.button(
                "🔄 Atualizar pré-visualização",
                use_container_width=True,
                key="btn_atualizar_preview_simples"
            )

    template_html = texto_simples_para_html(texto_corpo)

    template_html = adicionar_assinatura_imagem_ao_html(
        template_html=template_html,
        usar_assinatura_imagem=usar_assinatura_imagem,
        caminho_assinatura=caminho_assinatura,
        largura_imagem=320
    )

    if botao_atualizar_preview:
        st.session_state["assunto_preview"] = assunto
        st.session_state["template_html_preview"] = template_html

        with area_feedback_preview:
            st.success("Pré-visualização atualizada.")

    try:
        df_marcados = df[df[col_enviar].apply(valor_verdadeiro)]

        if not df_marcados.empty:
            linha_preview = df_marcados.iloc[0]

            assunto_base_preview = (
                st.session_state["assunto_preview"]
                or assunto
            )

            assunto_renderizado = renderizar_html(
                assunto_base_preview,
                linha_preview
            )

            with area_assunto_preview:
                st.info(f"Assunto: {assunto_renderizado}")

            html_base_preview = (
                st.session_state["template_html_preview"]
                or template_html
            )

            html_base_preview = preparar_html_preview_assinatura(
                html_base_preview,
                caminho_assinatura
            )

            html_preview = renderizar_html(
                html_base_preview,
                linha_preview
            )

            with area_preview:
                components.html(
                    montar_preview_html(html_preview),
                    height=390,
                    scrolling=True
                )

    except Exception as e:
        st.error(f"Erro ao gerar prévia: {e}")

else:
    template_html_inicial = """
<p>Olá {Nome},</p>
<p>Mensagem de teste.</p>
""".strip()

    col_html, col_preview = st.columns([1.15, 1])

    with col_html:
        st.markdown("### 🧩 Revise o HTML:")

        template_html = st.text_area(
            "Revise o HTML:",
            value=template_html_inicial,
            height=450,
            label_visibility="collapsed"
        )

    with col_preview:
        st.markdown("### 👀 Pré-visualização")

        area_assunto_preview = st.empty()
        area_preview = st.empty()
        area_botao_preview = st.empty()
        area_feedback_preview = st.empty()

        with area_botao_preview:
            botao_atualizar_preview = st.button(
                "🔄 Atualizar pré-visualização",
                use_container_width=True,
                key="btn_atualizar_preview_html"
            )

    template_html = adicionar_assinatura_imagem_ao_html(
        template_html=template_html,
        usar_assinatura_imagem=usar_assinatura_imagem,
        caminho_assinatura=caminho_assinatura,
        largura_imagem=320
    )

    if botao_atualizar_preview:
        st.session_state["assunto_preview"] = assunto
        st.session_state["template_html_preview"] = template_html

        with area_feedback_preview:
            st.success("Pré-visualização atualizada.")

    try:
        df_marcados = df[df[col_enviar].apply(valor_verdadeiro)]

        if not df_marcados.empty:
            linha_preview = df_marcados.iloc[0]

            assunto_base_preview = (
                st.session_state["assunto_preview"]
                or assunto
            )

            assunto_renderizado = renderizar_html(
                assunto_base_preview,
                linha_preview
            )

            with area_assunto_preview:
                st.info(f"Assunto: {assunto_renderizado}")

            html_base_preview = (
                st.session_state["template_html_preview"]
                or template_html
            )

            html_base_preview = preparar_html_preview_assinatura(
                html_base_preview,
                caminho_assinatura
            )

            html_preview = renderizar_html(
                html_base_preview,
                linha_preview
            )

            with area_preview:
                components.html(
                    montar_preview_html(html_preview),
                    height=390,
                    scrolling=True
                )

    except Exception as e:
        st.error(f"Erro ao gerar prévia: {e}")


if not template_html.strip():
    st.error("O corpo do e-mail está vazio.")
    st.stop()

if not assunto.strip():
    st.error("O assunto do e-mail está vazio.")
    st.stop()

if not st.session_state["template_html_preview"]:
    st.session_state["template_html_preview"] = template_html

if not st.session_state["assunto_preview"]:
    st.session_state["assunto_preview"] = assunto
    
## FIM EDITOR DE MENSAGEM

st.divider()

st.markdown("### Conferência da mensagem")
mensagem_pronta = st.checkbox(
    "Mensagem revisada e pronta para validação ou envio.",
    key="mensagem_pronta_check"
)

mudou_para_pronta = mensagem_pronta and not st.session_state["mensagem_pronta_anterior"]

if mudou_para_pronta:
    st.session_state["template_html_preview"] = template_html
    st.session_state["template_html_aprovado"] = template_html

mensagem_liberada = mensagem_pronta and (st.session_state["template_html_aprovado"] == template_html)
st.session_state["mensagem_pronta_anterior"] = mensagem_pronta

st.divider()


# =========================
# PASSO 5: AÇÃO NO OUTLOOK
# =========================

if not mensagem_liberada:
    st.info("🔒 O Passo 5 será liberado depois que a mensagem for revisada e marcada como pronta.")
    st.stop()

st.header("Passo 5: Ação no Outlook")

modo_acao = st.radio(
    "O que você deseja fazer?",
    [
        "🔍 Modo Validação: Apenas simular/validar no sistema (sem interagir com o Outlook)",
        "👁️ Modo Visualização: Abrir janela do e-mail na tela no Outlook (para conferir)",
        "🚀 Modo Envio real: disparo de e-mails da tabela validada pelo Outlook"
    ],
    index=0
)

with st.expander("⚙️ Configurações avançadas", expanded=True):
    coluna_adv_1, coluna_adv_2 = st.columns(2)

    with coluna_adv_1:
        intervalo = st.number_input(
            "Intervalo entre e-mails (segundos)",
            min_value=0.0,
            value=1.5,
            step=0.5
        )

        limite = st.number_input(
            "Limite de linhas processadas",
            min_value=0,
            value=1 if "Modo Visualização" in modo_acao else 0,
            step=1,
            help="Em Modo Visualização, recomenda-se processar poucas linhas por vez (ex: 1 a 3)."
        )

    with coluna_adv_2:
        permitir_envio_sem_anexo = st.checkbox(
            "Permitir mesmo se anexo personalizado estiver faltando",
            value=False
        )

enviar_real = ("Envio real" in modo_acao)
modo_display = ("Modo Visualização" in modo_acao)
somente_validar = ("Modo Validação" in modo_acao)

confirmacao_envio = True

if enviar_real:
    st.error("🚨 ATENÇÃO: O disparo dos e-mails está em MODO REAL e será realizado!")
    confirmacao_envio = st.checkbox("Tenho certeza e autorizo o disparo oficial.")
elif modo_display:
    st.info("👁️ **Modo Visualização Ativo:** O Outlook vai abrir a janela com o e-mail pronto na sua tela.")
else:
    st.info("🟡 **Modo Simulação:** Validação apenas na tela do sistema.")

botao_processar = st.button("🚀 INICIAR PROCESSO", type="primary", use_container_width=True)

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

    template_html_envio = adicionar_banner_teste_ao_html(
        template_html,
        modo_teste=(not enviar_real)
    )

    try:
        df_resultado, logs = processar_envios(
            df=df,
            template_html=template_html_envio,
            pasta_anexos_dinamicos=str(PASTA_ANEXOS_DINAMICOS) if usar_anexos_personalizados else None,
            pasta_anexos_fixos=str(PASTA_ANEXOS_FIXOS) if usar_anexos_fixos else None,
            assunto=f"[TESTE] {assunto}" if not enviar_real else assunto,
            col_enviar=col_enviar,
            col_email=col_email,
            col_anexos=col_anexos,
            enviar_real=enviar_real,
            modo_display=modo_display,
            somente_validar=somente_validar,
            intervalo=float(intervalo),
            limite=int(limite),
            permitir_envio_sem_anexo=permitir_envio_sem_anexo,
            exigir_anexo_personalizado=usar_anexos_personalizados,
            usar_assinatura_imagem=usar_assinatura_imagem,
            caminho_assinatura=(str(caminho_assinatura) if caminho_assinatura else None),
            callback=atualizar_tela
        )

        st.success("🎉 Processo finalizado com sucesso!")

        df_logs = pd.DataFrame(logs)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        arquivo_resultado = PASTA_LOGS / f"Resultado_Envio_{timestamp}.xlsx"
        arquivo_logs = PASTA_LOGS / f"Logs_Envio_{timestamp}.xlsx"

        df_resultado.to_excel(arquivo_resultado, index=False)

        if not df_logs.empty:
            df_logs.to_excel(arquivo_logs, index=False)
            st.dataframe(df_logs, use_container_width=True)

    except Exception as e:
        st.error(f"Ocorreu um erro no processamento: {e}")