import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import unicodedata
import re

st.set_page_config(page_title="Cherry Picking Natura", page_icon="💵", layout="wide")

st.title("💵 Agente de Sourcing: Cherry Picking Flexible")
st.write("Consolidação automática com seleção inteligente de colunas.")

arquivos_carregados = st.file_uploader("Suba as cotações aqui (.xlsx):", type=["xlsx"], accept_multiple_files=True)

def padronizar(txt):
    if pd.isna(txt) or txt is None: return ""
    txt = str(txt).upper().strip()
    return re.sub(r'\s+', ' ', ''.join(c for c in unicodedata.normalize('NFD', txt) if unicodedata.category(c) != 'Mn'))

def limpar_num(val):
    if pd.isna(val) or val is None: return None
    try:
        if isinstance(val, (int, float)): return float(val) if val > 0 else None
        num = re.findall(r'[0-9\.,\-]+', str(val).strip())
        if not num: return None
        t = num[0]
        if "." in t and "," in t: t = t.replace(".", "").replace(",", ".")
        elif "," in t: t = t.replace(",", ".")
        v = float(t)
        return v if v > 0 else None
    except: return None

if arquivos_carregados:
    st.info(f"📁 {len(arquivos_carregados)} planilhas carregadas.")
    
    # Dicionário para armazenar as tabelas extraídas de cada arquivo
    dados_extraidos = []
    
    for arq in arquivos_carregados:
        nome_fornecedor = arq.name.split(".")[0].replace("Template_para_Precificação_-_", "").replace("Cópia_de_", "").replace("1°_Rodada_", "").strip()
        try:
            xl = pd.ExcelFile(arq)
            # Tenta pegar a aba de precificação ou a primeira aba
            aba = next((s for s in xl.sheet_names if any(k in s.lower() for k in ["modelo", "precifica", "frota", "murici", "rotas"])), xl.sheet_names[0])
            
            # Lê primeiras 40 linhas para achar a tabela
            df_raw = pd.read_excel(arq, sheet_name=aba, header=None)
            
            # Encontra linha com mais textos válidos
            r_head = 0
            for r in range(min(30, len(df_raw))):
                vals = [padronizar(v) for v in df_raw.iloc[r].values if pd.notna(v)]
                if any(k in v for v in vals for k in ["ROTA", "ORIGEM", "DESTINO", "LINEHAUL", "CD", "ITEM", "VALOR", "CUSTO", "PRECO"]):
                    r_head = r
                    break
            
            df_table = pd.read_excel(arq, sheet_name=aba, header=r_head).dropna(how='all')
            df_table = df_table.loc[:, ~df_table.columns.astype(str).str.contains('^Unnamed')]
            
            if not df_table.empty:
                dados_extraidos.append({"fornecedor": nome_fornecedor, "df": df_table, "arquivo": arq.name})
        except Exception as e:
            st.warning(f"Erro ao abrir {arq.name}: {e}")

    if dados_extraidos:
        st.subheader("⚙️ Mapeamento de Colunas")
        st.caption("Se o sistema não identificar a coluna de valor automaticamente, selecione no menu abaixo:")
        
        propostas_prontas = []
        
        for item in dados_extraidos:
            forn = item["fornecedor"]
            df = item["df"]
            cols = list(df.columns)
            
            # Sugere colunas
            col_chave_sug = next((c for c in cols if any(k in str(c).lower() for k in ["rota", "linehaul", "destino", "cd", "item"])), cols[0] if cols else None)
            col_val_sug = next((c for c in cols if any(k in str(c).lower() for k in ["faturamento", "total", "valor", "preço", "preco", "custo", "r$"])), cols[-1] if cols else None)
            
            c1, c2 = st.columns([2, 2])
            with c1:
                c_chave = st.selectbox(f"🔑 Chave/Rota ({forn}):", cols, index=cols.index(col_chave_sug) if col_chave_sug in cols else 0, key=f"k_{forn}")
            with c2:
                c_val = st.selectbox(f"💵 Valor Rota ({forn}):", cols, index=cols.index(col_val_sug) if col_val_sug in cols else len(cols)-1, key=f"v_{forn}")
                
            df_sub = pd.DataFrame()
            df_sub["Item / Rota"] = df[c_chave].apply(padronizar)
            df_sub[forn] = df[c_val].apply(limpar_num)
            df_sub = df_sub.dropna(subset=["Item / Rota", forn])
            df_sub = df_sub[~df_sub["Item / Rota"].isin(["", "NAN", "TOTAL", "SUBTOTAL"])]
            
            if not df_sub.empty:
                propostas_prontas.append(df_sub)
        
        st.divider()
        
        if st.button("🚀 Gerar Consolidação Final Cherry Picking"):
            if propostas_prontas:
                df_consolidado = propostas_prontas[0]
                for p in propostas_prontas[1:]:
                    df_consolidado = pd.merge(df_consolidado, p, on="Item / Rota", how="outer")
                
                cols_forn = [list(p.columns)[-1] for p in propostas_prontas]
                
                df_consolidado["Melhor Preço"] = df_consolidado[cols_forn].min(axis=1)
                
                def calc_vencedor(r):
                    m = r["Melhor Preço"]
                    if pd.isna(m): return None
                    for c in cols_forn:
                        if r[c] == m: return c
                    return None
                
                df_consolidado["Fornecedor Vencedor"] = df_consolidado.apply(calc_vencedor, axis=1)
                
                st.balloons()
                st.success("✨ Cherry Picking Consolidado com Sucesso!")
                st.dataframe(df_consolidado, use_container_width=True)
                
                # Buffer Excel
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_consolidado.to_excel(writer, index=False, sheet_name="Cherry Picking")
                buffer.seek(0)
                
                st.download_button(
                    "📥 Baixar Consolidação em Excel (.xlsx)",
                    data=buffer,
                    file_name="Cherry_Picking_Consolidado_Final.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
