import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import unicodedata
import re

# =========================================================================
# CONFIGURAÇÃO DA PÁGINA E DESIGN POPPINS
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

st.title("💵 Agente de Sourcing: Cherry Picking (Modelo de Precificação)")
st.write("Consolidação automática focada nas abas de **Modelo de Precificação** das cotações.")

arquivos_carregados = st.file_uploader(
    "Suba as planilhas dos fornecedores (.xlsx ou .csv):", 
    type=["xlsx", "csv"], accept_multiple_files=True
)

# =========================================================================
# FUNÇÕES DE PADRONIZAÇÃO E HIGIENIZAÇÃO DE DADOS
# =========================================================================
def padronizar_texto(texto):
    """Remove acentos, espaços duplos e padroniza em maiúsculas para o merge perfeito."""
    if pd.isna(texto) or texto is None:
        return ""
    texto = str(texto).upper().strip()
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    texto = re.sub(r'\s+', ' ', texto)
    return texto

def limpar_valor(val):
    """Extrai valores numéricos float de strings financeiras."""
    if pd.isna(val) or val is None:
        return None
    try:
        if isinstance(val, (int, float)):
            return float(val)
        
        texto = str(val).strip().upper()
        # Captura apenas números, pontos, vírgulas e traços
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
# LEITURA E PROCESSAMENTO DA ABA "MODELO DE PRECIFICAÇÃO"
# =========================================================================
def processar_modelo_precificacao(file_obj, nome_arquivo):
    dfs_extraidos = []
    
    if nome_arquivo.endswith(".csv"):
        df_raw = pd.read_csv(file_obj)
        dfs_extraidos.append((nome_arquivo, df_raw))
    else:
        xl = pd.ExcelFile(file_obj)
        
        # Prioriza abas que contenham "modelo", "precifica", "template" ou "cotação"
        abas_alvo = [s for s in xl.sheet_names if any(k in s.lower() for k in ["modelo", "precifica", "template", "cotacao", "cotação"])]
        if not abas_alvo:
            abas_alvo = xl.sheet_names  # Fallback caso não ache pelo nome
            
        for sheet in abas_alvo:
            df_preview = pd.read_excel(file_obj, sheet_name=sheet, header=None)
            
            # Localiza a linha do cabeçalho varrendo as primeiras 25 linhas
            header_row = None
            for r_idx in range(min(25, len(df_preview))):
                vals = [str(v).strip().lower() for v in df_preview.iloc[r_idx].values if pd.notna(v)]
                # Procura palavras-chave típicas de tabelas de precificação
                tem_chave = any(k in v for v in vals for k in ["regi", "cd", "local", "rota", "origem", "veiculo", "cargo", "item"])
                tem_valor = any(k in v for v in vals for k in ["valor", "preço", "preco", "custo", "diaria", "total"])
                
                if tem_chave and tem_valor:
                    header_row = r_idx
                    break
            
            if header_row is None:
                header_row = 0
                
            df_sheet = pd.read_excel(file_obj, sheet_name=sheet, header=header_row)
            
            label = nome_arquivo
            if len(xl.sheet_names) > 1:
                label = f"{nome_arquivo} ({sheet.strip()})"
                
            dfs_extraidos.append((label, df_sheet))

    dfs_normalizados = []
    for label, df in dfs_extraidos:
        # Limpa colunas sem nome
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')].dropna(how='all')
        
        col_reg, col_item, col_turno, col_valor = None, None, None, None
        
        # Mapeamento dinâmico de colunas
        for col in df.columns:
            c_lower = str(col).strip().lower()
            if any(k in c_lower for k in ["regi", "cd", "local", "filial", "origem"]):
                col_reg = col
            elif any(k in c_lower for k in ["carg", "funç", "func", "veic", "veículo", "item", "descri"]):
                col_item = col
            elif any(k in c_lower for k in ["turn", "horar", "tipo"]):
                col_turno = col

        # Identificação da coluna de preço/custo final
        # Prioridade 1: Valor Total / Custo Total
        for col in df.columns:
            c_lower = str(col).strip().lower()
            if "total" in c_lower or "final" in c_lower:
                col_valor = col
                break
                
        # Prioridade 2: Valor Diária / Preço Unitário (sem ser taxa isolada)
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
            
            # Filtra linhas vazias e de totais agregados
            df_sub = df_sub.dropna(subset=["Região / Local", label])
            df_sub = df_sub[~df_sub["Região / Local"].isin(["", "NAN", "TOTAL", "REGIAO", "LOCAL"])]
            dfs_normalizados.append(df_sub)
            
    return dfs_normalizados

# =========================================================================
# FORMATADOR EXCEL COM ESTILO NATURA E FÓRMULAS
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
    
    header_fill = PatternFill(start_color="9E472A", end_color="9E472A", fill_type="solid") # Terracota Natura
    header_font = Font(name="Poppins", size=10, bold=True, color="FFFFFF")
    border_cinza = Border(
        left=Side(style='thin', color='E0D8D3'), right=Side(style='thin', color='E0D8D3'),
        top=Side(style='thin', color='E0D8D3'), bottom=Side(style='thin', color='E0D8D3')
    )
    
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    for col_idx in range(1, total_cols + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = border_cinza

    row_num = 2
    for _, row in df.iterrows():
        ws.cell(row=row_num, column=1, value=str(row.get("Região / Local", "")).title())
        ws.cell(row=row_num, column=2, value=str(row.get("Item / Cargo / Veículo", "")).title())
        ws.cell(row=row_num, column=3, value=str(row.get("Turno / Detalhe", "")).title())
        
        for col_idx, forn in enumerate(fornecedores, start=4):
            valor = row.get(forn, None)
            ws.cell(row=row_num, column=col_idx, value=valor)
        
        col_let_forn_start = get_column_letter(4)
        col_let_forn_end = get_column_letter(3 + len(fornecedores))
        col_let_min = get_column_letter(idx_min_col)
        
        # Fórmulas nativas do Excel compatíveis com todas as versões
        ws.cell(row=row_num, column=idx_min_col, value=f"=MIN({col_let_forn_start}{row_num}:{col_let_forn_end}{row_num})")
        ws.cell(row=row_num, column=idx_winner_col, value=f'=INDEX(${col_let_forn_start}$1:${col_let_forn_end}$1, MATCH({col_let_min}{row_num}, {col_let_forn_start}{row_num}:{col_let_forn_end}{row_num}, 0))')
        
        coluna_vencedora_python = row.get("Fornecedor Vencedor")
        col_idx_vencedora = fornecedores.index(coluna_vencedora_python) + 4 if coluna_vencedora_python in fornecedores else None
        
        winner_fill = PatternFill(start_color="E6F0EA", end_color="E6F0EA", fill_type="solid") # Verde suave
        winner_font = Font(name="Poppins", size=10, bold=True, color="1E3D2F")
        normal_font = Font(name="Poppins", size=10, color="000000")
        
        for col_idx in range(1, total_cols + 1):
            cell = ws.cell(row=row_num, column=col_idx)
            cell.border = border_cinza
            cell.font = normal_font
            
            if col_idx == col_idx_vencedora:
                cell.fill = winner_fill
                cell.font = winner_font
                
            if col_idx in [1, 2, 3]:
                cell.alignment = left_align
            else:
                cell.alignment = center_align
                if 4 <= col_idx <= idx_min_col:
                    cell.number_format = 'R$ #,##0.00'
        row_num += 1

    for col in ws.columns:
        ws.column_dimensions[get_column_letter(col[0].column)].width = 20
        
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# =========================================================================
# EXECUÇÃO DO APLICATIVO STREAMLIT
# =========================================================================
if arquivos_carregados:
    dfs_processados = []
    for arquivo in arquivos_carregados:
        nome_clean = arquivo.name.split(".")[0].replace("Template_para_Precificação_-_", "").replace("Cópia_de_", "").strip()
        try:
            sub_dfs = processar_modelo_precificacao(arquivo, nome_clean)
            dfs_processados.extend(sub_dfs)
        except Exception as e:
            st.error(f"Erro ao ler o arquivo {arquivo.name}: {e}")

    if dfs_processados:
        colunas_fornecedores = [list(d.columns)[-1] for d in dfs_processados]
        st.success(f"🤖 Agente de Sourcing: {len(colunas_fornecedores)} proposta(s) / aba(s) extraída(s) com sucesso!")

        if st.button("🚀 Gerar Consolidação Cherry Picking"):
            with st.spinner("⚡ Unificando itens e comparando cotações..."):
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
                st.success("✨ Cherry Picking gerado!")
                st.dataframe(df_consolidado, use_container_width=True)

                st.download_button(
                    label="📥 Baixar Consolidação em Excel (.xlsx)",
                    data=buffer_excel,
                    file_name="Consolidado_Cherry_Picking_Modelo_Precificacao.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
