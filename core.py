import os
import re
import time
from pathlib import Path
from datetime import datetime

import pandas as pd


COL_RESULTADO_STATUS = "Resultado_Status"
COL_RESULTADO_ERRO = "Resultado_Erro"
COL_RESULTADO_DATAHORA = "Resultado_DataHora"
COL_RESULTADO_ANEXOS_USADOS = "Resultado_Anexos_Usados"
COL_RESULTADO_ANEXOS_FALTANDO = "Resultado_Anexos_Faltando"


def valor_verdadeiro(valor):
    """
    Interpreta valores comuns como verdadeiro:
    TRUE, True, 1, Sim, S, X, Enviar, Yes, OK.
    """
    if pd.isna(valor):
        return False

    if isinstance(valor, bool):
        return valor

    texto = str(valor).strip().lower()

    return texto in [
        "true",
        "verdadeiro",
        "1",
        "sim",
        "s",
        "x",
        "enviar",
        "yes",
        "y",
        "ok"
    ]


def limpar_valor(valor):
    """
    Converte NaN em string vazia e remove espaços extras.
    """
    if pd.isna(valor):
        return ""

    return str(valor).strip()


def email_basico_valido(email):
    """
    Validação simples de e-mail.

    Não valida se a caixa postal existe.
    Apenas evita valores vazios ou formatos claramente inválidos.
    """
    email = limpar_valor(email)

    if not email:
        return False

    padrao = r"^[^@\s]+@[^@\s]+\.[^@\s]+$"

    return re.match(padrao, email) is not None


def renderizar_html(template_html, row):
    """
    Substitui variáveis no formato {NomeDaColuna}
    usando os valores da linha da planilha.

    Exemplo:
    Coluna: Nome
    HTML: Olá {Nome}
    """
    dados = {
        str(coluna).strip(): limpar_valor(valor)
        for coluna, valor in row.items()
    }

    padrao = re.compile(r"\{([^{}]+)\}")

    def substituir(match):
        chave = match.group(1).strip()

        if chave in dados:
            return dados[chave]

        return match.group(0)

    return padrao.sub(substituir, template_html)


def obter_lista_anexos(row, col_anexos):
    """
    Lê a coluna de anexos personalizados.

    Aceita:
    arquivo.pdf

    Ou:
    arquivo1.pdf; arquivo2.docx; arquivo3.xlsx
    """
    if not col_anexos:
        return []

    if col_anexos not in row.index:
        return []

    valor = row.get(col_anexos, "")

    if pd.isna(valor):
        return []

    texto = str(valor).strip()

    if not texto or texto.lower() == "nan":
        return []

    return [
        item.strip()
        for item in texto.split(";")
        if item.strip()
    ]


def listar_anexos_fixos(pasta_anexos_fixos):
    """
    Lista todos os arquivos existentes na pasta de anexos fixos.
    """
    if not pasta_anexos_fixos:
        return []

    pasta = Path(pasta_anexos_fixos)

    if not pasta.exists():
        return []

    return [
        str(arquivo)
        for arquivo in pasta.glob("*.*")
        if arquivo.is_file()
    ]


def validar_anexos_dinamicos(lista_anexos, pasta_anexos_dinamicos):
    """
    Verifica quais anexos personalizados existem e quais estão ausentes.
    """
    encontrados = []
    ausentes = []

    if not lista_anexos:
        return encontrados, ausentes

    if not pasta_anexos_dinamicos:
        return encontrados, lista_anexos

    for nome_arquivo in lista_anexos:
        caminho = os.path.join(str(pasta_anexos_dinamicos), nome_arquivo)

        if os.path.exists(caminho):
            encontrados.append(caminho)
        else:
            ausentes.append(nome_arquivo)

    return encontrados, ausentes


def inicializar_colunas_resultado(df):
    """
    Cria ou limpa as colunas de resultado em uma cópia da planilha.
    A planilha original não é alterada.
    """
    df_resultado = df.copy()

    colunas_resultado = [
        COL_RESULTADO_STATUS,
        COL_RESULTADO_ERRO,
        COL_RESULTADO_DATAHORA,
        COL_RESULTADO_ANEXOS_USADOS,
        COL_RESULTADO_ANEXOS_FALTANDO
    ]

    for coluna in colunas_resultado:
        df_resultado[coluna] = ""

    return df_resultado


def processar_envios(
    df,
    template_html,
    pasta_anexos_dinamicos,
    pasta_anexos_fixos,
    assunto,
    col_enviar,
    col_email,
    col_anexos=None,
    enviar_real=False,
    somente_validar=True,
    intervalo=1.5,
    limite=0,
    permitir_envio_sem_anexo=False,
    exigir_anexo_personalizado=False,
    usar_assinatura_imagem=False,
    caminho_assinatura=None,
    callback=None
):
    """
    Processa os registros da planilha.

    Modos:
    - somente_validar=True:
        valida e-mails, HTML e anexos, mas NÃO abre Outlook e NÃO envia.

    - enviar_real=True e somente_validar=False:
        envia automaticamente pelo Outlook.

    Regras importantes:
    - Se exigir_anexo_personalizado=True, cada linha precisa ter pelo menos um anexo informado.
    - Se houver anexo informado e ele não existir, o envio é bloqueado quando permitir_envio_sem_anexo=False.
    """
    df_resultado = inicializar_colunas_resultado(df)
    logs = []

    if col_enviar not in df_resultado.columns:
        raise ValueError(f"Coluna de envio não encontrada: {col_enviar}")

    if col_email not in df_resultado.columns:
        raise ValueError(f"Coluna de e-mail não encontrada: {col_email}")

    if col_anexos and col_anexos not in df_resultado.columns:
        raise ValueError(f"Coluna de anexos não encontrada: {col_anexos}")

    if exigir_anexo_personalizado and not col_anexos:
        raise ValueError(
            "Anexos personalizados foram selecionados, mas nenhuma coluna de anexos foi informada."
        )

    if not template_html or not str(template_html).strip():
        raise ValueError("O corpo do e-mail está vazio.")

    if usar_assinatura_imagem:
        if not caminho_assinatura:
            raise ValueError(
                "A assinatura em imagem foi habilitada, mas nenhuma imagem foi selecionada."
            )

        caminho_assinatura_path = Path(caminho_assinatura)

        if not caminho_assinatura_path.is_file():
            raise FileNotFoundError(
                f"A imagem da assinatura não foi encontrada: {caminho_assinatura}"
            )
    else:
        caminho_assinatura_path = None

    df_filtrado = df_resultado[
        df_resultado[col_enviar].apply(valor_verdadeiro)
    ]

    usar_outlook = enviar_real and not somente_validar

    outlook = None
    pythoncom = None

    if usar_outlook:
        try:
            import pythoncom as pc
            import win32com.client as win32

            pythoncom = pc
            pythoncom.CoInitialize()
            outlook = win32.Dispatch("Outlook.Application")

        except Exception as erro_outlook:
            if pythoncom:
                pythoncom.CoUninitialize()

            raise RuntimeError(
                "Não foi possível inicializar o Outlook. "
                "Verifique se está em Windows, com Outlook Classic instalado, aberto e configurado. "
                f"Detalhe técnico: {erro_outlook}"
            )

    lista_anexos_fixos_paths = listar_anexos_fixos(pasta_anexos_fixos)

    total_tentativas = 0

    try:
        for index, row in df_filtrado.iterrows():
            if limite and limite > 0 and total_tentativas >= limite:
                break

            email = limpar_valor(row[col_email])
            data_hora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            status = ""
            erro = ""
            anexos_dinamicos_ok = []
            anexos_dinamicos_faltando = []

            try:
                if not email_basico_valido(email):
                    status = "Não processado - e-mail inválido"
                    erro = "E-mail vazio ou em formato inválido"
                    raise ValueError(erro)

                corpo_html = renderizar_html(template_html, row)

                if not corpo_html.strip():
                    status = "Não processado - HTML vazio"
                    erro = "O corpo do e-mail ficou vazio após a renderização"
                    raise ValueError(erro)

                lista_anexos_dinamicos = obter_lista_anexos(row, col_anexos)

                if exigir_anexo_personalizado and not lista_anexos_dinamicos:
                    status = "Não enviado - anexo não informado"
                    erro = (
                        "A opção de anexos personalizados foi selecionada, "
                        "mas esta linha não possui nenhum arquivo informado na coluna de anexos."
                    )
                    raise FileNotFoundError(erro)

                anexos_dinamicos_ok, anexos_dinamicos_faltando = validar_anexos_dinamicos(
                    lista_anexos=lista_anexos_dinamicos,
                    pasta_anexos_dinamicos=pasta_anexos_dinamicos
                )

                if anexos_dinamicos_faltando and not permitir_envio_sem_anexo:
                    status = "Não enviado - anexo ausente"
                    erro = "Anexo(s) não encontrado(s): " + "; ".join(anexos_dinamicos_faltando)
                    raise FileNotFoundError(erro)

                todos_anexos_para_enviar = anexos_dinamicos_ok + lista_anexos_fixos_paths

                if usar_outlook:
                    mail = outlook.CreateItem(0)
                    mail.To = email
                    mail.Subject = assunto

                    for caminho_anexo in todos_anexos_para_enviar:
                        mail.Attachments.Add(str(caminho_anexo))

                    if usar_assinatura_imagem and caminho_assinatura_path:
                        anexo_assinatura = mail.Attachments.Add(
                            str(caminho_assinatura_path)
                        )

                        anexo_assinatura.PropertyAccessor.SetProperty(
                            "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
                            "assinatura_img"
                        )

                        anexo_assinatura.PropertyAccessor.SetProperty(
                            "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B",
                            True
                        )

                    mail.HTMLBody = corpo_html
                    
                    mail.Send()
                    status = "Enviado"

                    if intervalo and intervalo > 0:
                        time.sleep(intervalo)

                else:
                    status = "Validado"

            except Exception as e:
                if not erro:
                    erro = str(e)

                if not status:
                    status = "Erro"

            finally:
                total_tentativas += 1

                todos_anexos_usados = anexos_dinamicos_ok + lista_anexos_fixos_paths

                nomes_anexos_usados = [
                    os.path.basename(anexo)
                    for anexo in todos_anexos_usados
                ]

                str_anexos_usados = "; ".join(nomes_anexos_usados) if nomes_anexos_usados else "Nenhum"
                str_anexos_faltando = "; ".join(anexos_dinamicos_faltando)

                df_resultado.at[index, COL_RESULTADO_STATUS] = status
                df_resultado.at[index, COL_RESULTADO_ERRO] = erro
                df_resultado.at[index, COL_RESULTADO_DATAHORA] = data_hora
                df_resultado.at[index, COL_RESULTADO_ANEXOS_USADOS] = str_anexos_usados
                df_resultado.at[index, COL_RESULTADO_ANEXOS_FALTANDO] = str_anexos_faltando

                log = {
                    "DATA_HORA": data_hora,
                    "LINHA_PLANILHA": index + 2,
                    "EMAIL": email,
                    "STATUS": status,
                    "ANEXOS_USADOS": str_anexos_usados,
                    "ANEXOS_FALTANDO": str_anexos_faltando,
                    "ERRO": erro
                }

                logs.append(log)

                if callback:
                    callback(log)

        return df_resultado, logs

    finally:
        if usar_outlook and pythoncom:
            pythoncom.CoUninitialize()