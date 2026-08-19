import streamlit as st

def aplicar_estilo_global():
    st.markdown(
        """
        <style>
            :root {
                --cor-primaria: #003B5C;
                --cor-secundaria: #0072CE;
                --cor-texto: #1F2937;
                --cor-texto-suave: #4B5563;
                --cor-codigo-texto: #B91C1C;
                --cor-codigo-fundo: #FEF2F2;
            }

            @media (prefers-color-scheme: dark) {
                :root {
                    --cor-primaria: #7DD3FC;
                    --cor-secundaria: #38BDF8;
                    --cor-texto: #F3F4F6;
                    --cor-texto-suave: #CBD5E1;
                    --cor-codigo-texto: #FCA5A5;
                    --cor-codigo-fundo: #3F1D1D;
                }
            }

            html, body {
                font-size: 17px;
                color: var(--cor-texto);
            }

            [data-testid="stMarkdownContainer"] p {
                font-size: 17px !important;
                line-height: 1.55 !important;
                color: var(--cor-texto) !important;
            }

            h1 {
                font-size: 36px !important;
                font-weight: 700 !important;
                color: var(--cor-primaria) !important;
            }

            h2 {
                font-size: 30px !important;
                font-weight: 700 !important;
                color: var(--cor-primaria) !important;
            }

            h3 {
                font-size: 24px !important;
                font-weight: 700 !important;
                color: var(--cor-secundaria) !important;
            }

            label, [data-testid="stWidgetLabel"] p {
                font-size: 17px !important;
                color: var(--cor-texto) !important;
                font-weight: 600 !important;
            }

            [data-testid="stAlert"] p {
                font-size: 17px !important;
                line-height: 1.5 !important;
            }

            button {
                font-size: 17px !important;
                font-weight: 600 !important;
                border-radius: 8px !important;
            }

            input, textarea {
                font-size: 17px !important;
                color: var(--cor-texto) !important;
            }

            div[data-baseweb="select"] * {
                font-size: 17px !important;
            }

            div[role="radiogroup"] label p {
                font-size: 17px !important;
                color: var(--cor-texto) !important;
            }

            details summary p {
                font-size: 17px !important;
                font-weight: 700 !important;
                color: var(--cor-primaria) !important;
            }

            [data-testid="stCaptionContainer"] p {
                font-size: 15px !important;
                color: var(--cor-texto-suave) !important;
                line-height: 1.4 !important;
            }

            code {
                color: var(--cor-codigo-texto) !important;
                /* background-color: var(--cor-codigo-fundo) !important; */
                border-radius: 4px !important;
                padding: 2px 5px !important;
            }
        </style>
        """,
        unsafe_allow_html=True
    )