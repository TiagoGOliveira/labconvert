import pandas as pd
import streamlit as st


def ajustar_nome_aba(nome):
    """Garante que o nome da aba não exceda 31 caracteres."""
    if nome is None:
        raise ValueError("nome não pode ser nulo")
    return nome[:31]


def salvar_em_abas(df):
    if df is None or df.empty:
        raise ValueError("DataFrame não pode ser nulo ou vazio")

    if 'Análise' not in df.columns:
        raise ValueError("Coluna 'Análise' não encontrada no DataFrame")

    grupos = {
        'Amostragem': [
            'Amostragem-hora', 'Amostragem-Nível Dinâmico', 'Amostragem-Temperatura Ambiente',
            'Amostragem-pH', 'Amostragem-Temperatura', 'Amostragem-Potencial de Oxidação/Redução',
            'Amostragem-Condutividade', 'Amostragem-Oxigênio Dissolvido', 'Amostragem-Turbidez',
            'Presença de fase Livre', 'Tamanho da Fase Móvel', 'Modelo da Bamba', 'Nível Estático',
            'Temperatura da Amostra na Saida', 'Temperatura da Amostra na Chegada', '1-Oxigenio Dissolvido',
            '1-Potencial de Oxidaçao/Reduçao', 'Condutividade Elétrica', 'Corantes Artificiais',
            'Materiais Flutuantes', 'Resíduos Sólidos Objetáveis', 'Aspecto',
            'Cloro Total/Cloro Residual Total (mg/L)', 'Cloro Livre/Cloro Residual Livre (mg/L)',
            'Cloro Combinado/Cloraminas Totais (mg/L)', 'Diluição da Amostra para leitura de análise e de Série Clorada',
            'Turbidez Visual', 'Monocloramina em Campo (mg/L)'
        ],
        'Físico-Químico': [
            'pH', 'Temperatura', 'Turbidez', 'Oxigênio Dissolvido', 'ORP (Potencial de Oxi-Redução)',
            'Condições Meteorológicas'
        ],
        'PAH': [
            'Antraceno', 'Benzo (a) Antraceno', 'Benzo (a) Pireno', 'Benzo (b) Fluoranteno',
            'Benzo (g,h,i) Perileno', 'Benzo (k) Fluoranteno', 'Criseno', 'Dibenzo (a,h) Antraceno',
            'Fenantreno', 'Fluoreno', 'Indeno (1,2,3) Pireno', 'Pireno', 'Fluoranteno', 'Total PAH\'s'
        ],
        'Purga': [
        'Modo de Operação', 'Demais Condições Ambientais', 'Diâmetro Tubo Descarga', 'Volume do Sistema', 
        'Poço', 'Diâmetro do Poço', 'Captação da Bomba', 'Fundo do Poço', 'Vazão da Purga', 
        'Tempo de Estabilização', 'Volume da Purga', 'Turbidez Anterior à Purga', 'Turbidez Posterior à Purga', 
        'Referência do Nível', '1-Temperatura', '1-Condutividade', '1-pH', '1-Vazão', '1-Temperatura Ambiente', 
        '1-Nível de Água', 'Nível Estático com Equipamento', 'Nível Dinâmico', 'pH do Preservante', 
        'Nomenclatura', 'Inicio da Seção Filtrante', 'Observação Final'
        ],  
        'TPH': [
        'C10 a C12 (Alifáticos)', 'C12 a C16 (Alifáticos)', 'C16 a C21 (Alifáticos)', 'C21 a C32 (Alifáticos)', 
        'C10 a C12 (Aromáticos)', 'C12 a C16 (Aromáticos)', 'C16 a C21 (Aromáticos)', 'C21 a C32 (Aromáticos)', 
        '1,2-Diclorobenzeno-d4 (TPH Fracionado Voláteis (L))', '4-Bromofluorobenzeno (TPH Fracionado Voláteis (L))', 
        'C5 a C8  (Alifáticos)', 'C9 a C18  (Alifáticos)', 'C19 a C32  (Alifáticos)', 'C6 a C8 (Aromáticos)', 
        'C9 a C10 (Aromáticos)', 'C10 a C32 (Aromáticos)'
        ],
        'VOC': [
        'Benzeno', 'Tolueno', 'Etilbenzeno', 'o-Xileno', 'm,p-Xileno', 'Xileno Total', 'BTEX Total', 
        'Acenafteno', 'Acenaftileno', 'Naftaleno', '2-Fluorobifenil', 'p-Terfenil-D14', 
        '1,2-Diclorobenzeno-d4 (BTEX (L))', '4-Bromofluorobenzeno (BTEX (L))', 'o-Terfenil', 
        '1,2-Diclorobenzeno-d4', '4-Bromofluorobenzeno'
        ],
        'Inorgânicos': [
        'Alumínio', 'Antimônio', 'Arsênio', 'Bário', 'Berílio', 'Boro', 'Cádmio', 'Chumbo', 'Cianeto', 'Cloreto',
        'Cloro', 'Cobalto', 'Cobre', 'Cromo', 'Ferro', 'Fluoreto', 'Fósforo', 'Lítio', 'Manganês', 'Mercúrio', 
        'Níquel', 'Nitrato', 'Nitrito', 'Nitrogênio', 'Selênio', 'Sulfato', 'Sulfeto', 'Urânio', 'Vanádio'
        ]}


    output_file = "dados_processados_grupos.xlsx"

    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for grupo, palavras_chave in grupos.items():
                filtro = df['Análise'].str.contains('|'.join(palavras_chave), case=False, na=False)
                df_grupo = df.loc[filtro].reset_index(drop=True)  # Corrigir índices duplicados
                if not df_grupo.empty:
                    aba = ajustar_nome_aba(grupo)
                    df_grupo.to_excel(writer, sheet_name=aba, index=False)

            filtro_outros = ~df['Análise'].str.contains('|'.join([item for sublist in grupos.values() for item in sublist]), case=False, na=False)
            outras_analises = df.loc[filtro_outros].reset_index(drop=True)  # Corrigir índices duplicados
            if not outras_analises.empty:
                outras_analises.to_excel(writer, sheet_name="Outras Análises", index=False)

        return output_file
    except Exception as e:
        raise RuntimeError(f"Erro ao salvar dados em abas: {e}")

def detectar_tabela(df_ref):
    """Detecta automaticamente a tabela no DataFrame baseado em critérios como a presença de dados válidos em várias colunas."""
    for i, row in df_ref.iterrows():
        # Verifica se a linha contém uma combinação de texto e valores numéricos, indicando que pode ser o cabeçalho
        if row.notna().sum() >= 2 and row.apply(lambda x: isinstance(x, (str, int, float))).all():
            return i  # Retorna o índice da linha que parece ser o cabeçalho
    return None  # Se não encontrar nenhuma linha válida, retorna None

def preparar_aba(df_ref, tipo_amostra):
    """Processa as abas do arquivo de referência, detectando a tabela e o cabeçalho automaticamente."""
    if df_ref is None or df_ref.empty:
        return pd.DataFrame(columns=['Parametros', 'Tipo de Amostra'])

    # Remover duplicatas de índice
    df_ref = df_ref.reset_index(drop=True)

    # Tornar nomes de colunas únicos se necessário
    if not df_ref.columns.is_unique:
        df_ref.columns = [f"{col}_{i}" if df_ref.columns.duplicated()[i] else col for i, col in enumerate(df_ref.columns)]

    # Detectar a linha que pode ser o cabeçalho da tabela
    cabecalho_index = detectar_tabela(df_ref)

    if cabecalho_index is not None:
        # Definir o cabeçalho com base na linha detectada
        header_row = df_ref.iloc[cabecalho_index]
        if len(header_row) == len(df_ref.columns):
            df_ref.columns = header_row
        else:
            raise ValueError(
                f"Número de colunas ({len(df_ref.columns)}) e de valores na linha de cabeçalho ({len(header_row)}) não coincidem."
            )

        # Remover as linhas acima do cabeçalho e ajustar os índices
        df_ref = df_ref.iloc[cabecalho_index + 1:].reset_index(drop=True).dropna(axis=1, how='all')

    df_ref['Tipo de Amostra'] = tipo_amostra
    return df_ref


def mesclar_dados(df_base, ref_abas):
    """Mescla os dados do arquivo principal com o de referência, detectando cabeçalhos em várias abas."""
    df_ag = preparar_aba(ref_abas.get('Água', pd.DataFrame()), 'Água Subterrânea')
    df_tp = preparar_aba(ref_abas.get('TPH Fracionado', pd.DataFrame()), 'TPH')
    df_sl = preparar_aba(ref_abas.get('Solo', pd.DataFrame()), 'Solo')

    # Adicionar mais abas conforme necessário, por exemplo:
    # df_outros = preparar_aba(ref_abas.get('Outra Aba', pd.DataFrame()), 'Outro Tipo de Amostra')

    if df_ag.empty and df_tp.empty and df_sl.empty:
        st.warning("Nenhuma aba válida foi encontrada no arquivo de referência.")
        return df_base

    return pd.concat([df_base, df_ag, df_tp, df_sl], ignore_index=True)


# Interface Streamlit

# Configurar título da página e ícone
st.set_page_config(
    page_title="Organizador e Mesclador de Dados", 
    page_icon="https://github.com/TiagoGOliveira/labconvert/blob/main/L%20-%20Lead%20(1).png?raw=true", 
    layout="wide", 
    initial_sidebar_state="expanded"
)

# Exibir logo
logo_url = "https://lead-sa.com.br/wp-content/uploads/elementor/thumbs/lead-sa-logo-qq9mfkoiu8l57gz69xhgrmnl3z2xrwh799bnd1xibk.png"
st.image(logo_url, width=200)

# Estilo personalizado com CSS
st.markdown(
    """
    <style>
        body {
            background-image: url('https://lead-sa.com.br/wp-content/uploads/2024/11/2-1-scaled.jpg');  # Caminho para a imagem de fundo
            background-size: cover;  # Ajusta a imagem para cobrir toda a tela
            background-repeat: no-repeat;  # Impede que a imagem se repita
            background-position: center center;  # Centraliza a imagem
            color: #222224;  # Texto com cor primária
            font-family: 'Open Sans', sans-serif;  # Fonte personalizada
        }

        .stButton>button {
            background-color: #222224;  # Cor de fundo do botão
            color: white;  # Cor do texto do botão
        }

        .stButton>button:hover {
            background-color: #13AB00;  # Cor de fundo do botão ao passar o mouse
        }

        h1, h2, h3, h4, h5, h6 {
            color: #001646;  # Cor dos títulos
        }

        .sidebar .sidebar-content {
            background-color: #222224;  # Cor de fundo da sidebar
            color: white;  # Cor do texto na sidebar
        }

        .sidebar .sidebar-content a {
            color: white;  # Cor dos links na sidebar
        }

        .sidebar .sidebar-content a:hover {
            color: #13AB00;  # Cor dos links na sidebar ao passar o mouse
        }
    </style>
    """, unsafe_allow_html=True
)

# Título e conteúdo do app
st.title("Conversor de Laudos Laboratoriais")
st.subheader("Carregue os arquivos gerados no MyLims e adicione os valores de referência")

uploaded_file = st.file_uploader("Faça upload do arquivo Excel principal (.xls ou .xlsx)", type=["xls", "xlsx"])
ref_file = st.file_uploader("Faça upload do arquivo de referência (.xlsx)", type=["xlsx"])

if uploaded_file and ref_file:
    try:
        # Carregar arquivo principal
        if uploaded_file.name.endswith('.xls'):
            df = pd.read_excel(uploaded_file, engine='xlrd')
        elif uploaded_file.name.endswith('.xlsx'):
            df = pd.read_excel(uploaded_file)

        if df is None or df.empty:
            st.error("O arquivo principal está vazio ou não foi carregado corretamente.")
            st.stop()

        st.write("Prévia dos dados do arquivo principal:")
        st.dataframe(df.head())

        # Carregar abas do arquivo de referência
        ref_abas = pd.read_excel(ref_file, sheet_name=None)

        if not isinstance(ref_abas, dict) or len(ref_abas) == 0:
            st.error("O arquivo de referência não contém abas válidas.")
            st.stop()

        st.write("Abas do arquivo de referência:")
        st.write(list(ref_abas.keys()))

        # Mesclar os dados
        df_mesclado = mesclar_dados(df, ref_abas)

        if df_mesclado is None or df_mesclado.empty:
            st.error("A mesclagem dos dados resultou em um DataFrame vazio.")
            st.stop()

        st.write("Prévia dos dados mesclados:")
        st.dataframe(df_mesclado.head())

        # Salvar os dados processados em abas
        output_file = salvar_em_abas(df_mesclado)

        # Exibir link para download
        with open(output_file, "rb") as f:
            st.download_button(
                label="Baixar arquivo processado",
                data=f,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"Erro ao processar os arquivos: {e}")
else:
    st.info("Por favor, carregue os dois arquivos para começar.")


