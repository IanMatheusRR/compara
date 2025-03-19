import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
from io import BytesIO

# DEVE SER A PRIMEIRA INSTRUÇÃO DO SCRIPT:
st.set_page_config(
    page_title="Sistema de Controle e Comparação de Preços",
    page_icon="logo-eqtl-app-teste2.png",
    layout="wide"
)

# Inicializa variáveis de sessão para controlar o estado do botão e da mensagem
if "show_info" not in st.session_state:
    st.session_state.show_info = False
if "last_button_value" not in st.session_state:
    st.session_state.last_button_value = None

# Componente personalizado para o botão de dúvidas com ícone (sem key)
custom_button_html = """
<html>
  <head>
    <style>
      #custom-doubt-button {
        background-color: #fff;
        border: 2px solid #003a63;
        border-radius: 8px;
        padding: 10px;
        display: flex;
        align-items: center;
        justify-content: center;
        width: 150px;
        height: 60px;
        cursor: pointer;
      }
      #custom-doubt-button img {
        height: 40px;
        width: 40px;
        margin-right: 10px;
      }
      #custom-doubt-button span {
        font-size: 16px;
        color: #003a63;
        font-weight: bold;
      }
    </style>
  </head>
  <body>
    <div id="custom-doubt-button" onclick="handleClick()">
      <img src="https://via.placeholder.com/40?text=%3F" alt="Ícone">
      <span>Dúvidas</span>
    </div>
    <script>
      function handleClick() {
        Streamlit.setComponentValue(new Date().getTime());
      }
      Streamlit.setFrameHeight(document.documentElement.scrollHeight);
    </script>
  </body>
</html>
"""

try:
    button_value = components.html(custom_button_html, height=120)
except Exception as e:
    st.error(f"Erro ao renderizar o componente personalizado: {e}")
    button_value = None

if button_value and button_value != st.session_state.last_button_value:
    st.session_state.last_button_value = button_value
    st.session_state.show_info = not st.session_state.show_info

if st.session_state.show_info:
    try:
        with st.modal("Instruções de Uso"):
            st.write(
                "Para otimizar o uso das funcionalidades, por favor, carregue o arquivo CJI3 "
                "extraído do SAP com o layout BRP_RAW, utilizando o campo 'Drag and drop file here'. "
                "Certifique-se de que o formato das colunas permaneça inalterado e remova a linha amarela "
                "localizada na última linha do arquivo extraído da CJI3."
            )
    except Exception:
        st.info(
            "Para otimizar o uso das funcionalidades, por favor, carregue o arquivo CJI3 "
            "extraído do SAP com o layout BRP_RAW, utilizando o campo 'Drag and drop file here'. "
            "Certifique-se de que o formato das colunas permaneça inalterado e remova a linha amarela "
            "localizada na última linha do arquivo extraído da CJI3."
        )

# ... Restante do código do aplicativo ...

