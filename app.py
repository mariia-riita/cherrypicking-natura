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
st.set_page_config(page_title="Cherry Picking - Natura", page_icon="💵", layout="wide")

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&display=swap');
    html, body, [data-testid="stAppViewContainer"], .stApp { font-family: 'Poppins', sans-serif !important; }
    </style>
    """, unsafe_allow_html=True
)

st.title("💵 Agente de Sourcing: Cherry Picking")
st.write("Consolidação automática de cotações em lote.")

arquivos_carregados = st.file_uploader(
    "Suba todas as cotações aqui (.xlsx ou .csv):", 
    type=["xlsx", "csv"], 
    accept_multiple_files=True
)

# =========================================================================
# FUNÇÕES DE HIGIENIZAÇÃO DE DADOS
# =========================================================================
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
# PROCESSAMENTO DE ARQUIVOS
# =========================================================================
def processar_modelo_precificacao(file_obj, nome_arquivo):
    dfs_extraidos = []
    
    if nome_arquivo.endswith(".csv"):
        df_raw = pd.read_csv(file_obj)
        dfs_extraidos.append((nome_arquivo, df_raw))
    else:
        xl = pd.ExcelFile(file_obj)
        abas_alvo = [s for s in xl.sheet_names if any(k in s.lower() for k in ["modelo", "precifica", "template", "cotacao", "cotação"])]
        if not abas_alvo:
            abas_alvo = xl.sheet_names
            
        for sheet in abas_alvo:
            df_preview = pd.read_excel(file_obj, sheet_name=sheet, header=None)
            
            header_row = None
            for r_idx in range(min(25, len(df_preview))):
                vals = [str(v).strip().lower() for v in df_preview.iloc[r_idx].values if pd.notna(v)]
                tem_chave = any(k in v for v in vals for k in ["regi", "cd", "local", "rota", "origem", "veiculo", "cargo", "item"])
                tem_valor = any(k in v for v in vals for k in ["valor", "preço", "preco", "custo", "diaria", "total"])
                if tem_chave and tem_valor:
                    header_row = r_idx
                    break
            
            if header_row is None:
                header_row = 0
                
            df_sheet = pd.read_excel(file_obj, sheet_name=sheet, header=header_row)
            label = f"{nome_arquivo} ({sheet.strip()})" if len(xl.sheet_names) > 1 else nome_arquivo
            dfs_extraidos.append((label, df_sheet))

    dfs_normalizados = []
    for label, df in dfs_extraidos:
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')].dropna(how='all')
        col_reg, col_item, col_turno, col_valor = None, None, None, None
        
        for col in df.columns:
            c_lower = str(col).strip().lower()
            if any(k in c_lower for k in ["regi", "cd", "local", "filial", "origem"]): col_reg = col
            elif any(k in c_lower for k in ["carg", "funç", "func", "veic", "veículo", "item", "descri"]): col_item = col
            elif any(k in c_lower for k in ["turn", "horar", "tipo"]): col_turno = col

        for col in df.columns:
            c_lower = str(col).strip().lower()
            if "total" in c_lower or "final" in c_lower:
                col_valor = col
                break
                
        if not col_valor:
            for col in df.columns:
                c_lower = str(col).strip().lower()
                if any(g in c_lower for g in ["valor", "preço", "preco", "custo", "diaria"]) and not any(bad in c_lower for bad in ["%", "taxa", "imposto", "pis", "iss"]):
                    col_valor = col
                    break

        if col_reg and col_valor:
            df_sub = pd.DataFrame()
            df_sub["Região / Local"] = df[col_reg].apply(padronizar_texto)
            df_sub["Item / Cargo / Veículo"] = df[col_item].apply(padronizar_texto) if col_item else "PADRÃO"
            df_sub["Turno / Detalhe"] = df[col_turno].apply(padronizar_texto) if col_turno else "GERAL"
            df_sub[label] = df[col_valor].apply(limpar_valor)
            
            df_sub = df_sub.dropna(subset=["Região / Local", label])
            df_sub = df_sub[~df_sub["Região / Local"].isin(["", "NAN", "TOTAL", "REGIAO", "LOCAL"])]
            if not df_sub.empty:
                dfs_normalizados.append(df_sub)
            
    return dfs_normalizados

# =========================================================================
# EXPORTAÇÃO EXCEL
# =========================================================================
def estilizar_planilha_excel(df, fornecedores):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Cherry Picking"
    
    colunas_chave = ["Região / Local", "Item / Cargo / Veículo", "Turno / Detalhe"]
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
        ws.cell(row=row_num, column=1, value=str(row.get("Região / Local", "")).title())
        ws.cell(row=row_num, column=2, value=str(row.get("Item / Cargo / Veículo", "")).title())
        ws.cell(row=row_num, column=3, value=str(row.get("Turno / Detalhe", "")).title())
        
        for col_idx, forn in enumerate(fornecedores, start=4):
            ws.cell(row=row_num, column=col_idx, value=row.get(forn, None))
        
        col_let_start = get_column_letter(4)
        col_let_end = get_column_letter(3 + len(fornecedores))
        col_let_min = get_column_letter(idx_min_col)
        
        ws.cell(row=row_num, column=idx_min_col, value=f"=MIN({col_let_start}{row_num}:{col_let_end}{row_num})")
        ws.cell(row=row_num, column=idx_winner_col, value=f'=INDEX(${col_let_start}$1:${col_let_end}$1, MATCH({col_let_min}{row_num}, {col_let_start}{row_num}:{col_let_end}{row_num}, 0))')
        
        coluna_vencedora_python = row.get("Fornecedor Vencedor")
        col_idx_vencedora = fornecedores.index(coluna_vencedora_python) + 4 if coluna_vencedora_python in fornecedores else None
        
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
                
            if col_idx > 3:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if col_idx <= idx_min_col:
                    cell.number_format = 'R$ #,##0.00'
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
        row_num += 1

    for col in ws.columns:
        ws.column_dimensions[get_column_letter(col[0].column)].width = 20
        
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# =========================================================================
# INTERFACE E FLUXO DE EXECUÇÃO
# =========================================================================
if arquivos_carregados:
    st.info(f"📁 **{len(arquivos_carregados)} arquivo(s)** selecionado(s) para processamento.")
    
    # BOTÃO VISÍVEL IMEDIATAMENTE
    if st.button("🚀 Gerar Consolidação Cherry Picking"):
        dfs_processados = []
        relatorio_erros = []
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for idx, arquivo in enumerate(arquivos_carregados):
            nome_clean = arquivo.name.split(".")[0].replace("Template_para_Precificação_-_", "").replace("Cópia_de_", "").strip()
            status_text.text(f"⏳ Lendo arquivo {idx+1} de {len(arquivos_carregados)}: {arquivo.name}...")
            
            try:
                sub_dfs = processar_modelo_precificacao(arquivo, nome_clean)
                if sub_dfs:
                    dfs_processados.extend(sub_dfs)
                else:
                    relatorio_erros.append((arquivo.name, "Nenhuma coluna de Região ou Valor Válido foi identificada."))
            except Exception as e:
                relatorio_erros.append((arquivo.name, f"Erro de leitura: {str(e)}"))
                
            progress_bar.progress((idx + 1) / len(arquivos_carregados))
            
        status_text.empty()
        progress_bar.empty()
        
        if relatorio_erros:
            with st.expander("⚠️ Detalhes de arquivos ignorados ou não identificados"):
                for arq, msg in relatorio_erros:
                    st.warning(f"**{arq}**: {msg}")

        if dfs_processados:
            colunas_fornecedores = [list(d.columns)[-1] for d in dfs_processados]
            st.success(f"✅ {len(colunas_fornecedores)} proposta(s)/aba(s) extraída(s) com sucesso!")
            
            with st.spinner("⚡ Unificando bases e realizando o Cherry Picking..."):
                df_consolidado = dfs_processados[0]
                for df_prox in dfs_processados[1:]:
                    df_consolidado = pd.merge(
                        df_consolidado, df_prox, 
                        on=["Região / Local", "Item / Cargo / Veículo", "Turno / Detalhe"], 
                        how="outer"
                    )

                df_consolidado = df_consolidado.sort_values(by=["Região / Local", "Item / Cargo / Veículo"]).reset_index(drop=True)

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
                st.write("📊 **Prévia dos Dados Consolidados:**")
                st.dataframe(df_consolidado, use_container_width=True)

                st.download_button(
                    label="📥 Baixar Consolidação Final (.xlsx)",
                    data=buffer_excel,
                    file_name="Cherry_Picking_Consolidado_22_Propostas.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.error("❌ Nenhum dado pôde ser extraído das planilhas enviadas. Verifique a estrutura das abas.")
