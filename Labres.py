import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import re
import io

def ajustar_nome_aba(nome):
    """Garante que o nome da aba não exceda 31 caracteres e remove caracteres inválidos."""
    if nome is None:
        raise ValueError("Nome não pode ser nulo")
    
    nome = re.sub(r'[\/:*?"<>|]', '', nome)
    return nome[:31]

def detectar_tabela(df_ref):
    """Detecta automaticamente a tabela no DataFrame baseado em critérios como a presença de dados válidos em várias colunas."""
    for i, row in df_ref.iterrows():
        if row.notna().sum() >= 2 and row.apply(lambda x: isinstance(x, (str, int, float))).all():
            return i
    return None

def preparar_aba(df_ref, tipo_amostra):
    """Processa uma aba do arquivo de referência, detectando a tabela e o cabeçalho automaticamente."""
    if df_ref is None or df_ref.empty:
        return pd.DataFrame(columns=['Parametros', 'VI', 'Unidade', 'Tipo de Amostra', 'Fonte'])

    df_ref = df_ref.reset_index(drop=True)
    cabecalho_index = detectar_tabela(df_ref)

    if cabecalho_index is not None:
        header_row = df_ref.iloc[cabecalho_index]
        if len(header_row) == len(df_ref.columns):
            df_ref.columns = header_row
        else:
            raise ValueError("Número de colunas e valores na linha de cabeçalho não coincidem.")

        df_ref = df_ref.iloc[cabecalho_index + 1:].reset_index(drop=True).dropna(axis=1, how='all')

    df_ref['Tipo de Amostra'] = tipo_amostra

    if 'Fonte' not in df_ref.columns:
        df_ref['Fonte'] = None

    df_ref = converter_colunas_para_numeros(df_ref, ['VI'])
    return df_ref

def processar_referencias(ref_abas):
    """Processa todas as abas do arquivo de referência em um único DataFrame consolidado."""
    if not ref_abas or len(ref_abas) == 0:
        return pd.DataFrame(columns=['Parametros', 'VI', 'Unidade', 'Tipo de Amostra', 'Fonte', 'Grupo'])

    referencias = []
    for aba, df_ref in ref_abas.items():
        tipo_amostra = aba
        df_preparado = preparar_aba(df_ref, tipo_amostra)
        referencias.append(df_preparado)

    return pd.concat(referencias, ignore_index=True)

def converter_colunas_para_numeros(df, colunas):
    """Converte colunas especificadas para números decimais, substituindo valores não numéricos por NaN."""
    for coluna in colunas:
        if coluna in df.columns:
            df[coluna] = df[coluna].astype(str).str.replace('[^0-9.,-]', '', regex=True)
            df[coluna] = df[coluna].str.replace(',', '.')
            df[coluna] = pd.to_numeric(df[coluna], errors='coerce')
    return df

def comparar_resultados(df_base, df_ref):
    """Compara os valores da coluna 'Resultado' com os valores de referência e preenche a coluna 'Grupo' com base na tabela de referência."""
    if df_base is None or df_base.empty:
        raise ValueError("O DataFrame principal está vazio.")

    if df_ref is None or df_ref.empty:
        raise ValueError("O DataFrame de referência está vazio.")

    df_base = converter_colunas_para_numeros(df_base, ['Resultado'])
    df_ref = converter_colunas_para_numeros(df_ref, ['VI'])

    df_base['Comparação'] = None
    df_base['Valor de Referência'] = None
    df_base['Unidade Divergente'] = None
    df_base['Grupo'] = None
    df_base['Fonte'] = None

    for i, row in df_base.iterrows():
        analise = row['Análise']
        resultado = row['Resultado']
        unidade_base = row.get('Unidade', None)

        ref_compatível = df_ref[df_ref['Parametros'] == analise]

        if not ref_compatível.empty:
            for _, ref_row in ref_compatível.iterrows():
                valor_ref = ref_row['VI']
                unidade_ref = ref_row.get('Unidade', None)
                grupo_ref = ref_row.get('Grupo', None)
                fonte_ref = ref_row.get('Fonte', None)

                unidade_divergente = None
                if unidade_base and unidade_ref and unidade_base != unidade_ref:
                    unidade_divergente = f"{unidade_base} != {unidade_ref}"

                if pd.notna(valor_ref):
                    if resultado < valor_ref:
                        comparacao = "Abaixo"
                    elif resultado > valor_ref:
                        comparacao = "Acima"
                    else:
                        comparacao = "Igual"

                    df_base.at[i, 'Comparação'] = comparacao
                    df_base.at[i, 'Valor de Referência'] = valor_ref
                    df_base.at[i, 'Unidade Divergente'] = unidade_divergente
                    df_base.at[i, 'Grupo'] = grupo_ref
                    df_base.at[i, 'Fonte'] = fonte_ref
                    break
            else:
                df_base.at[i, 'Comparação'] = "Sem referência"
        else:
            df_base.at[i, 'Comparação'] = "Sem referência"
            df_base.at[i, 'Grupo'] = "Sem grupo"

    colunas_para_remover = [
        "Relatório de Análises", "Nº Amostra", "Proposta Comercial", 
        "Data do Recebimento", "Data da Publicação", "Previsão de Entrega", "Situação"
    ]
    df_base = df_base.drop(columns=[col for col in colunas_para_remover if col in df_base.columns], errors='ignore')

    return df_base

def salvar_em_abas(df_comparado):
    """Salva cada grupo de 'Grupo' em uma aba separada do Excel com a coluna 'Fonte' mantida."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for grupo, grupo_df in df_comparado.groupby('Grupo'):
            if pd.notna(grupo):
                nome_aba = ajustar_nome_aba(str(grupo))
                grupo_df.to_excel(writer, sheet_name=nome_aba, index=False)
                worksheet = writer.sheets[nome_aba]

                for cell in worksheet["A1:K1"]:
                    for c in cell:
                        c.font = Font(bold=True, color="FFFFFF")
                        c.fill = PatternFill(start_color="3C5082", end_color="3C5082", fill_type="solid")
                        c.alignment = Alignment(horizontal="center", vertical="center")

                for col in worksheet.columns:
                    max_length = 0
                    column = col[0].column_letter
                    for cell in col:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(cell.value)
                        except:
                            pass
                    adjusted_width = (max_length + 2)
                    worksheet.column_dimensions[column].width = adjusted_width

                thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                for row in worksheet.iter_rows(min_row=2, min_col=1, max_row=worksheet.max_row, max_col=worksheet.max_column):
                    for cell in row:
                        cell.border = thin_border

    output.seek(0)
    return output

# Interface Streamlit
st.set_page_config(
    page_title="Conversor de Laudos",
    page_icon="https://github.com/TiagoGOliveira/labconvert/blob/main/L%20-%20Lead%20(1).png?raw=true",
    layout="wide",
    initial_sidebar_state="expanded",
)

logo_url = "https://lead-sa.com.br/wp-content/uploads/elementor/thumbs/lead-sa-logo-qq9mfkoiu8l57gz69xhgrmnl3z2xrwh799bnd1xibk.png"
st.image(logo_url, width=200)

st.title("Conversor de Laudos Laboratoriais MyLims")
st.subheader("Carregue o arquivo exportado do MyLims e dos Valores de Referência para começar.")

uploaded_file = st.file_uploader("Arquivo Excel Resultados (export) (.xls ou .xlsx)", type=["xls", "xlsx"])
ref_file = st.file_uploader("Arquivo de referência (.xlsx)", type=["xlsx"])

if uploaded_file and ref_file:
    try:
        df_principal = pd.read_excel(uploaded_file)
        if df_principal.empty:
            st.error("O arquivo de resultados está vazio ou não foi carregado corretamente.")
            st.stop()

        st.write("Prévia dos dados do laudo laboratorial:")
        st.dataframe(df_principal.head())

        ref_abas = pd.read_excel(ref_file, sheet_name=None)
        if not ref_abas:
            st.error("O arquivo de referência não contém abas válidas.")
            st.stop()

        st.write("Abas do arquivo de referência:")
        st.write(list(ref_abas.keys()))

        df_referencias = processar_referencias(ref_abas)
        if df_referencias.empty:
            st.error("Nenhuma referência válida encontrada.")
            st.stop()

        st.write("Prévia dos Valores de Referência:")
        st.dataframe(df_referencias.head())

        df_comparado = comparar_resultados(df_principal, df_referencias)

        st.write("Resultados Tabelados Acima dos Valores de Referência:")
        df_comparado_acima = df_comparado[df_comparado['Comparação'] == 'Acima']
        st.dataframe(df_comparado_acima)

        output = salvar_em_abas(df_comparado)

        st.download_button(
            label="Baixar Planilha Formatada",
            data=output,
            file_name="Resultados_Projeto XXX.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Erro: {e}")
