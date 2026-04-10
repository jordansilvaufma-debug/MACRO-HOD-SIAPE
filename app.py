import streamlit as st
import pandas as pd
from datetime import datetime
import io

# --- CONFIGURAÇÃO DE ACESSO ---
CHAVES_ATIVAS = ["Siape2024", "MestreHOD"] # Adicione novas senhas aqui

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
    
    # Cabeçalho idêntico ao seu original
    macro = f'<HAScript name="MOVFIN_AUTOMACRO" description="" timeout="60000" pausetime="14000" promptall="true" blockinput="false" author="AUTOMACRO" creationdate="{agora}" supressclearevents="false" usevars="false" ignorepauseforenhancedtn="true" delayifnotenhancedtn="0" ignorepausetimeforenhancedtn="true">\n'

    total_rows = len(df)
    for index, row in df.iterrows():
        if pd.isna(row.get("MATRICULA")) and pd.isna(row.get("VALOR")):
            continue

        acao = str(row.get("CMD I/A/E", "I")).upper()
        matricula = str(row.get("MATRICULA")).split('.')[0]
        rubrica_completa = f"{row.get('R ou D')}{str(row.get('RUBRICA')).zfill(5)}{row.get('SEQUÊNCIA')}{row.get('CMD I/A/E')}"
        unidade_atual = str(row.get("UPAG"))
        
        blocos = []

        # 1. TROCA DE UNIDADE (Lógica exata do seu arquivo .py)
        if ultima_unidade is not None and unidade_atual != ultima_unidade:
            blocos.extend([
                {"v": "[pf12]", "d": '<oia status="NOTINHIBITED" />'},
                {"v": "[tab][tab][tab][tab][tab][tab][tab][tab][tab]trocahab[enter]", "d": '<oia status="NOTINHIBITED" />\n<numfields number="76" />\n<numinputfields number="10" />'},
                {"v": "x[enter]", "d": '<oia status="NOTINHIBITED" />\n<numfields number="177" />\n<numinputfields number="2" />'},
                {"v": "s[enter]", "d": '<oia status="NOTINHIBITED" />\n<numfields number="181" />\n<numinputfields number="1" />'},
                {"v": "[tab][enter]", "d": '<oia status="NOTINHIBITED" />\n<numfields number="76" />\n<numinputfields number="10" />'}
            ])

        # 2. INSERIR OU ALTERAR
        if acao in ['I', 'A']:
            mes_ref = formatar_mes_referencia(row.get("MÊS/REFERENCIA"))
            valor = float(row.get("VALOR", 0))
            valor_fmt = f"{valor:.2f}".replace('.', ',')
            assunto = str(row.get("ASSUNTO", ""))
            doc_legal = str(row.get("DOC LEGAL", ""))
            justificativa = str(row.get("JUSTIFICATIVA", ""))

            blocos.append({"v": f"{matricula}[enter]", "d": '<oia status="NOTINHIBITED" />'})
            cmd_rubrica = f"{rubrica_completa}[tab]X[enter]" if acao == 'I' else f"{rubrica_completa}[enter]"
            blocos.append({"v": cmd_rubrica, "d": '<oia status="NOTINHIBITED" />\n<numfields number="82" />\n<numinputfields number="6" />'})
            blocos.append({"v": f"{mes_ref}{valor_fmt}[tab]{assunto}[enter]", "d": '<oia status="NOTINHIBITED" />\n<numfields number="134" />\n<numinputfields number="17" />'})

            if valor > 943.74:
                blocos.append({"v": "[enter]", "d": '<oia status="NOTINHIBITED" />\n<numfields number="136" />\n<numinputfields number="0" />'})
                desc_justif = '<oia status="NOTINHIBITED" />\n<numfields number="136" />\n<numinputfields number="7" />'
            else:
                desc_justif = '<oia status="NOTINHIBITED" />\n<numfields number="134" />\n<numinputfields number="7" />'

            blocos.append({"v": f"PROCESSO {doc_legal}[tab]{justificativa}[enter]", "d": desc_justif})
            blocos.append({"v": "C[enter]", "d": '<oia status="NOTINHIBITED" />\n<numfields number="52" />\n<numinputfields number="1" />'})
            blocos.append({"v": "[enter]", "d": '<oia status="NOTINHIBITED" />\n<numfields number="53" />\n<numinputfields number="0" />'})
            blocos.append({"v": "[pf12]", "d": '<oia status="NOTINHIBITED" />\n<numfields number="82" />\n<numinputfields number="6" />'})

        # 3. EXCLUIR
        elif acao == 'E':
            blocos.extend([
                {"v": f"{matricula}[enter]", "d": '<oia status="NOTINHIBITED" />'},
                {"v": f"{rubrica_completa}[enter]", "d": '<oia status="NOTINHIBITED" />\n<numfields number="82" />\n<numinputfields number="6" />'},
                {"v": "[enter]", "d": '<oia status="NOTINHIBITED" />\n<numfields number="130" />\n<numinputfields number="0" />'},
                {"v": "C[enter]", "d": '<oia status="NOTINHIBITED" />\n<numfields number="52" />\n<numinputfields number="1" />'},
                {"v": "[enter]", "d": '<oia status="NOTINHIBITED" />\n<numfields number="53" />\n<numinputfields number="0" />'},
                {"v": "[pf12]", "d": '<oia status="NOTINHIBITED" />\n<numfields number="82" />\n<numinputfields number="6" />'}
            ])

        # GERAÇÃO DAS TELAS (Fiel ao seu original)
        for i, bloco in enumerate(blocos):
            is_last_row_of_all = (index == total_rows - 1)
            is_last_screen_of_row = (i == len(blocos) - 1)
            macro += f'    <screen name="Tela{tela_base}" entryscreen="{"true" if is_first_screen_of_macro else "false"}" exitscreen="{"true" if (is_last_row_of_all and is_last_screen_of_row) else "false"}" transient="false">\n'
            macro += f'      <description>\n        {bloco["d"]}\n      </description>\n'
            macro += f'      <actions><input value="{bloco["v"]}" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" /></actions>\n'
            if not (is_last_row_of_all and is_last_screen_of_row):
                macro += f'      <nextscreens timeout="0"><nextscreen name="Tela{tela_base + 1}" /></nextscreens>\n'
            macro += f'    </screen>\n'
            is_first_screen_of_macro = False
            tela_base += 1
        ultima_unidade = unidade_atual

    macro += '</HAScript>'
    return macro

# --- INTERFACE STREAMLIT ---
st.set_page_config(page_title="HOD Macro Pro", layout="centered")

st.sidebar.title("🔐 Autenticação")
senha = st.sidebar.text_input("Chave de Acesso", type="password")

if senha in CHAVES_ATIVAS:
    st.title("🚀 Conversor MOVFIN Seq 6")
    uploaded_file = st.file_uploader("Selecione a planilha Excel", type=["xlsx"])
    
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        if st.button("Gerar Macro"):
            conteudo = gerar_conteudo_macro(df)
            st.download_button(
                label="📥 Baixar MOVFIN_FINAL.mac",
                data=conteudo.encode('latin1'),
                file_name="MOVFIN_FINAL.mac",
                mime="text/plain"
            )
else:
    st.warning("Por favor, insira uma chave ativa na barra lateral para liberar o conversor.")
