import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import unicodedata
import re

# =========================================================================
# CONFIGURAÇÃO DA PÁGINA
# =========================================================================
st.set_page_config(page_title="Cherry Picking - Frota Dedicada Natura", page_icon="🚚", layout="wide")

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&display=swap');
    html, body, [data-testid="stAppViewContainer"], .stApp { font-family: 'Poppins', sans-serif !important; }
    </style>
    """, unsafe_allow_html=True
)

st.title("🚚 Sourcing Agent: Cherry Picking - Frota Dedicada (Murici)")
st.write("Consolidação de Propostas e Formação de Preço por Rota/LineHaul.")

arquivos_carregados = st.file_uploader(
    "Suba as planilhas dos fornecedores (.xlsx ou .csv):", 
    type=["xlsx", "csv"], 
    accept_multiple_files=True
)

def padronizar_texto(texto):
    if pd.isna(texto) or texto is None:
        return ""
    texto = str(texto).upper().strip()
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    return re.sub(r'\s+', ' ', texto)

def limpar_valor(val):
    if pd.isna(val) or val is None:
        return None
    try:
        if isinstance(val, (int, float)):
            return float(val) if val > 0 else None
        
        texto = str(val).strip().upper()
        numeros = re.findall(r'[0-9\.,\-]+', texto)
        if not numeros: return None
        
        texto_limpo = numeros[0]
        if "." in texto_limpo and "," in texto_limpo:
            texto_limpo = texto_limpo.replace(".", "").replace(",", ".")
        elif "," in texto_limpo:
            texto_limpo = texto_limpo.replace(",", ".")
            
        val_float = float(texto_limpo)
        return val_float if val_float > 0 else None
    except Exception:
        return None

# =========================================================================
# LEITURA E PROCESSAMENTO DA TABELA DE ROTAS (FROTA DEDICADA)
# =========================================================================
def processar_aba_frota_dedicada(file_obj, nome_arquivo):
    dfs_extraidos = []
    
    if nome_arquivo.endswith(".csv"):
        df_raw = pd.read_csv(file_obj)
        dfs_extraidos.append((nome_arquivo, df_raw))
    else:
        xl = pd.ExcelFile(file_obj)
        abas_alvo = [s for s in xl.sheet_names if any(k in s.lower() for k in ["modelo", "precifica", "frota", "murici", "rotas", "template"])]
        if not abas_alvo:
            abas_alvo = xl.sheet_names
            
        for sheet in abas_alvo:
            try:
                df_raw = pd.read_excel(file_obj, sheet_name=sheet, header=None)
            except Exception:
                continue
                
            if df_raw.empty: continue

            # Procura a linha que contém a tabela de Rotas (CD Origem, Rota_LineHaul, Km mês, etc)
            header_row = None
            for r_idx in range(len(df_raw)):
                row_vals = [padronizar_texto(v) for v in df_raw.iloc[r_idx].values if pd.notna(v)]
                if any("ROTA" in v for v in row_vals) and any("ORIGEM" in v or "KM" in v or "LINEHAUL" in v for v in row_vals):
                    header_row = r_idx
                    break
            
            if header_row is None:
                header_row = 0

            df_sheet = pd.read_excel(file_obj, sheet_name=sheet, header=header_row)
            label = f"{nome_arquivo} ({sheet.strip()})" if len(xl.sheet_names) > 1 else nome_arquivo
            dfs_extraidos.append((label, df_sheet))

    dfs_normalizados = []

    for label, df in dfs_extraidos:
        df = df.loc[:, ~df.columns.astype(str).str.contains('^Unnamed')].dropna(how='all')
        if df.empty: continue
            
        col_rota, col_origem, col_dest, col_valor = None, None, None, None
        
        for col in df.columns:
            c_lower = str(col).strip().lower()
            if any(k in c_lower for k in ["rota_linehaul", "linehaul", "destino", "rota"]):
                col_dest = col
            elif any(k in c_lower for k in ["origem", "cd origem", "cd"]):
                col_origem = col

        # Identifica a coluna de Faturamento / Preço Final
        for col in df.columns:
            c_lower = str(col).strip().lower()
            if any(k in c_lower for k in ["faturamento", "preço final", "preco final", "valor total", "total rota"]):
                col_valor = col
                break

        if not col_valor:
            for col in df.columns:
                c_lower = str(col).strip().lower()
                if any(k in c_lower for k in ["margem", "imposto", "custo fixo", "custo variavel"]):
                    continue
                if any(k in c_lower for k in ["valor", "preço", "preco", "total", "r$"]):
                    col_valor = col
                    break

        if col_dest and col_valor:
            df_sub = pd.DataFrame()
            df_sub["CD Origem"] = df[col_origem].apply(padronizar_texto) if col_origem else "MURICI"
            df_sub["Rota / Destino"] = df[col_dest].apply(padronizar_texto)
            df_sub[label] = df[col_valor].apply(limpar_valor)
            
            df_sub = df_sub.dropna(subset=["Rota / Destino", label])
            df_sub = df_sub[~df_sub["Rota / Destino"].isin(["", "NAN", "TOTAL", "ROTA_LINEHAUL", "DESTINO"])]
            
            if not df_sub.empty:
                dfs_normalizados.append(df_sub)

    return dfs_normalizados

# =========================================================================
# EXPORTAÇÃO EXCEL COM FORMATAÇÃO NATURA
# =========================================================================
def estilizar_planilha_excel(df, fornecedores):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Cherry Picking Rotas"
    
    colunas_chave = ["CD Origem", "Rota / Destino"]
    colunas_completas = colunas_chave + fornecedores + ["Melhor Preço", "Fornecedor Vencedor"]
    ws.append(colunas_completas)
    
    idx_min_col = len(colunas_chave) + len(fornecedores) + 1
    idx_winner_col = idx_min_col + 1
    total_cols = idx_winner_col
    
    header_fill = PatternFill(start_color="9E472A", end_color="9E472A", fill_type="solid")
    header_font = Font(name="Poppins", size=10, bold=True, color="FFFFFF")
    border_cinza = Border(left=Side(style='thin', color='E0D8D3'), right=Side(style='thin', color='E0D8D3'), top=Side(style='thin', color='E0D8D3'), bottom=Side(style='thin', color='E0D8D3'))
    
    for col_idx in range(1, total_cols + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_cinza

    row_num = 2
    for _, row in df.iterrows():
        ws.cell(row=row_num, column=1, value=str(row.get("CD Origem", "")).title())
        ws.cell(row=row_num, column=2, value=str(row.get("Rota / Destino", "")).title())
        
        for col_idx, forn in enumerate(fornecedores, start=3):
            ws.cell(row=row_num, column=col_idx, value=row.get(forn, None))
        
        col_let_start = get_column_letter(3)
        col_let_end = get_column_letter(2 + len(fornecedores))
        col_let_min = get_column_letter(idx_min_col)
        
        ws.cell(row=row_num, column=idx_min_col, value=f"=MIN({col_let_start}{row_num}:{col_let_end}{row_num})")
        ws.cell(row=row_num, column=idx_winner_col, value=f'=INDEX(${col_let_start}$1:${col_let_end}$1, MATCH({col_let_min}{row_num}, {col_let_start}{row_num}:{col_let_end}{row_num}, 0))')
        
        coluna_vencedora_python = row.get("Fornecedor Vencedor")
        col_idx_vencedora = fornecedores.index(coluna_vencedora_python) + 3 if coluna_vencedora_python in fornecedores else None
        
        winner_fill = PatternFill(start_color="E6F0EA", end_color="E6F0EA", fill_type="solid")
        winner_font = Font(name="Poppins", size=10, bold=True, color="1E3D2F")
        normal_font = Font(name="Poppins", size=10, color="000000")
        
        for col_idx in range(1, total_cols + 1):
            cell = ws.cell(row=row_num, column=col_idx)
            cell.border = border_cinza
            cell.font = normal_font
            
            if col_idx == col_idx_vencedora:
                cell.fill = winner_fill
                cell.font = winner_font
                
            if col_idx >= 3:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if col_idx <= idx_min_col:
                    cell.number_format = 'R$ #,##0.00'
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
        row_num += 1

    for col in ws.columns:
        ws.column_dimensions[get_column_letter(col[0].column)].width = 25
        
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# =========================================================================
# INTERFACE STREAMLIT
# =========================================================================
if arquivos_carregados:
    st.info(f"📁 **{len(arquivos_carregados)} arquivo(s)** carregado(s).")
    
    if st.button("🚀 Processar Cherry Picking por Rota"):
        dfs_processados = []
        relatorio_erros = []
        
        progress_bar = st.progress(0)
        
        for idx, arquivo in enumerate(arquivos_carregados):
            nome_clean = arquivo.name.split(".")[0].replace("Template_para_Precificação_-_", "").replace("Cópia_de_", "").strip()
            try:
                sub_dfs = processar_aba_frota_dedicada(arquivo, nome_clean)
                if sub_dfs:
                    dfs_processados.extend(sub_dfs)
                else:
                    relatorio_erros.append((arquivo.name, "Tabela de rotas não localizada."))
            except Exception as e:
                relatorio_erros.append((arquivo.name, f"Erro: {str(e)}"))
                
            progress_bar.progress((idx + 1) / len(arquivos_carregados))
            
        progress_bar.empty()
        
        if dfs_processados:
            colunas_fornecedores = [list(d.columns)[-1] for d in dfs_processados]
            
            df_consolidado = dfs_processados[0]
            for df_prox in dfs_processados[1:]:
                df_consolidado = pd.merge(
                    df_consolidado, df_prox, 
                    on=["CD Origem", "Rota / Destino"], 
                    how="outer"
                )

            def calc_melhor_preco(row):
                vals = [row[c] for c in colunas_fornecedores if pd.notna(row[c]) and row[c] is not None]
                return min(vals) if vals else None

            def calc_fornecedor_vencedor(row):
                best = calc_melhor_preco(row)
                if best is None: return None
                for c in colunas_fornecedores:
                    if row[c] == best: return c
                return None

            df_consolidado["Melhor Preço"] = df_consolidado.apply(calc_melhor_preco, axis=1)
            df_consolidado["Fornecedor Vencedor"] = df_consolidado.apply(calc_fornecedor_vencedor, axis=1)

            buffer_excel = estilizar_planilha_excel(df_consolidado, colunas_fornecedores)

            st.balloons()
            st.success("✨ Cherry Picking de Rotas gerado com sucesso!")
            st.dataframe(df_consolidado, use_container_width=True)

            st.download_button(
                label="📥 Baixar Planilha Consolidada (.xlsx)",
                data=buffer_excel,
                file_name="Cherry_Picking_Frota_Dedicada_Murici.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
