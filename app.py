import streamlit as st
import pandas as pd
from datetime import datetime
import io

# --- BANCO DE SENHAS (Simulando Inscrições) ---
# Aqui você coloca os códigos que você vai vender para os clientes
CHAVES_ATIVAS = ["Siape2024", "MestreHOD", "Conv123"]

# --- FUNÇÕES DE LÓGICA (O seu motor) ---
def formatar_mes_referencia(mes_ref):
    meses_abreviados = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"]
    try:
        dt = pd.to_datetime(mes_ref)
        return f"{meses_abreviados[dt.month - 1]}{dt.year}"
    except:
        return str(mes_ref).upper().strip()

def gerar_conteudo_macro(df):
    agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    tela_base = 1
    ultima_unidade = None
    is_first_screen_of_macro = True
    macro = f'<HAScript name="MOVFIN_AUTOMACRO" author="AUTOMACRO" creationdate="{agora}">\n'
    
    total_rows = len(df)
    for index, row in df.iterrows():
        if pd.isna(row.get("MATRICULA")): continue
        
        acao = str(row.get("CMD I/A/E", "I")).upper()
        matricula = str(row.get("MATRICULA")).split('.')[0]
        rubrica_completa = f"{row.get('R ou D')}{str(row.get('RUBRICA')).zfill(5)}{row.get('SEQUÊNCIA')}{row.get('CMD I/A/E')}"
        unidade_atual = str(row.get("UPAG"))
        
        blocos = []
        # (A lógica de blocos que você já tem permanece aqui...)
        # Vou resumir para o código não ficar gigante, mas você mantém o seu loop completo.
        
        # Exemplo simplificado de tela:
        blocos.append({"v": f"{matricula}[enter]", "d": "Entrada Matricula"})
        
        for i, bloco in enumerate(blocos):
            macro += f'    <screen name="Tela{tela_base}">\n'
            macro += f'      <actions><input value="{bloco["v"]}" /></actions>\n'
            macro += f'    </screen>\n'
            tela_base += 1
            
    macro += '</HAScript>'
    return macro

# --- INTERFACE ---
st.set_page_config(page_title="HOD Automacro", layout="centered")

# Área de Login na Barra Lateral
st.sidebar.title("🔐 Acesso Restrito")
senha_digitada = st.sidebar.text_input("Digite sua Chave de Acesso:", type="password")

if senha_digitada in CHAVES_ATIVAS:
    st.sidebar.success("Acesso Liberado!")
    
    st.title("🚀 Gerador FPATMOVFIN - Seq 6")
    st.info("Utilize a planilha padrão para gerar seus scripts .mac")

    uploaded_file = st.file_uploader("Suba o Excel aqui", type=["xlsx"])

    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        if st.button("Gerar Arquivo .mac"):
            mac_output = gerar_conteudo_macro(df)
            st.download_button(
                label="📥 Baixar MOVFIN_FINAL.mac",
                data=mac_output.encode('latin1'),
                file_name="MOVFIN_FINAL.mac"
            )
else:
    st.title("Acesso Bloqueado")
    st.warning("Para utilizar esta ferramenta, você precisa de uma chave de ativação.")
    st.write("Deseja adquirir uma licença? Clique no botão abaixo:")
    st.link_button("Falar com Suporte no WhatsApp", "https://wa.me/seunumeroaqui")
