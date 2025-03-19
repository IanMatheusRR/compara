import streamlit as st

# DEVE SER A PRIMEIRA INSTRUÇÃO DO SCRIPT:
st.set_page_config(
    page_title="Sistema de Controle e Comparação de Preços",
    page_icon="❓",
    layout="wide"
)

# Inicializa a variável de sessão para controlar a exibição da mensagem, se ainda não existir.
if "show_info" not in st.session_state:
    st.session_state.show_info = False

# Campo na barra lateral para personalização do ícone do botão de dúvidas.
custom_icon_path = st.sidebar.text_input(bater-papo.png)

# Define o rótulo do botão com base no caminho informado.
# Como st.button não aceita HTML, usamos um rótulo textual.
if custom_icon_path:
    button_label = "Dúvidas"  # O usuário pode customizar o ícone visualmente fora do botão.
else:
    button_label = "❓"

# Exibe o botão logo acima do título.
if st.button(button_label, key="toggle_info"):
    st.session_state.show_info = not st.session_state.show_info

# Se show_info for True, exibe a mensagem de orientação.
if st.session_state.show_info:
    st.info(
        "Para otimizar o uso das funcionalidades, por favor, carregue o arquivo CJI3 extraído do SAP com o layout BRP_RAW, "
        "utilizando o campo 'Drag and drop file here'. Certifique-se de que o formato das colunas permaneça inalterado e remova a "
        "linha amarela presente na última linha do arquivo extraído da CJI3."
    )

st.title("Sistema de Controle e Comparação de Preços")

# ... (restante do código do app)


