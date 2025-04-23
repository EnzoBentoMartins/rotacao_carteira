import streamlit as st
import os
import pyodbc
import sqlite3
import numpy as np
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta

import warnings
warnings.filterwarnings('ignore')

# ---------- CONFIGURAÇÕES INICIAIS ----------
st.set_page_config(page_title="Rotação de Carteiras", layout="wide")
st.title("🔁 Sistema de Rotação de Carteiras")

# ---------- CONEXÃO COM BANCO DE DADOS ----------
@st.cache_data
def carregar_dados_sql():
    server = st.secrets["DB_SERVER"]
    database = st.secrets["DB_NAME"]
    username = st.secrets["DB_USER"]
    password = st.secrets["DB_PASSWORD"]

    connection_string = (
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={server};DATABASE={database};UID={username};PWD={password}"
    )

    conn = pyodbc.connect(connection_string)

    query = """SELECT 
    a.[id] AS Conta_ID,
    a.tipo_conta,
    b.razao_social AS Razao_Social_Pessoas,
    b.cpf_cnpj AS CNPJ,
    LEFT(b.cpf_cnpj, 8) AS Raiz_CNPJ,
    c.[grupo_id] AS Grupo_Econômico_ID,
    c.[grupo_nome] AS Grupo_Econômico_Nome,
    v.razao_social AS Nome_Vendedor,
    b.[data_ultima_venda] AS Data_Ultima_Venda_Individual,
    a.[data_abertura] AS Data_Abertura_Conta,
    CASE 
        WHEN c.[grupo_id] IS NOT NULL THEN 
            ISNULL(
                (SELECT MAX(p.data_ultima_venda) 
                 FROM [dbo].[pessoas] p
                 INNER JOIN [dbo].[rel_pessoas] r ON p.id = r.id
                 WHERE r.grupo_id = c.[grupo_id] AND p.data_ultima_venda IS NOT NULL),
                b.[data_ultima_venda]
            )
        ELSE b.[data_ultima_venda]
    END AS Data_Ultima_Venda_Grupo_CNPJ,
    a.classificacao_id AS Classificacao_Conta,
    b.classificacao_id AS Classificacao_Pessoa,
    a.porte_id AS Porte_Empresa,
    (SELECT TOP 1 d.id FROM [dbo].[rel_crm_orcamentos] AS d WHERE d.pessoa_cliente_id = b.id ORDER BY d.data_emissao DESC) AS Orcamento_ID,
    (SELECT TOP 1 d.data_emissao FROM [dbo].[rel_crm_orcamentos] AS d WHERE d.pessoa_cliente_id = b.id ORDER BY d.data_emissao DESC) AS Data_Emissao_Ultimo_Orcamento
FROM
    [grupofort].[dbo].[crm_contas] AS a
    INNER JOIN [dbo].[pessoas] AS b ON a.cliente_id = b.id
    INNER JOIN [dbo].[rel_pessoas] AS c ON b.id = c.id
    INNER JOIN [dbo].[pessoas] AS v ON a.vendedor_id = v.id
WHERE
    a.tipo_conta = 2
    AND a.excluido = 0
    AND a.status_conta = 0
    AND b.classificacao_id <> 1
    AND a.classificacao_id <> 1;"""
    df = pd.read_sql(query, conn)
    conn.close()
    return df

# ---------- SELEÇÃO DE GRUPO DE VENDEDORES ----------

vendedores_ativos_helder = ['Bryan Casarotto',
'Daniele Schmitz',
'GILBERTO LIMA DE PINHO JUNIOR',
'Laura Vitoria Da Silveira Trindade',
'Leonardo Bianchi',
'Letícia Eduarda Cruz',
'LUCAS VASCONCELOS BATTAGLIA KRAUSE',
'RONALDO DA COSTA BARRIOS',
'Ruan da Silva',
'TALIA LINS RAMOS',
'Willian Luiz Pereira'
]

vendedores_ativos_karen = ['Amanda Dias do Amaral',
'CRISTIAN RHEINHEIMER',
'Eduardo Dutra de Lima',
'GUSTAVO BALBINOTT VENASSI',
'Guilherme Rafael Hartmann Soares',
'Gustavo Cesar Burnier',
'JOAO PEDRO MOCELIN',
'Joao Gustavo Santian Da Silva',
'Kaylane Victoria Sousa Sa',
'LUCAS SIMOES BERNART',
'MARJANA  KUHN',
'MAURICIO HENRIQUE CESCO',
'MURILO DUARTE DA SILVA',
'Sidinei Da Silva Dias',
'TIAGO PEDROSO DA SILVA'
]

opcao = st.selectbox("Escolha o grupo de vendedores:", ["Distribuição (Helder)", "Corporativo (Karen)"])
vendedores_ativos = vendedores_ativos_helder if "Helder" in opcao else vendedores_ativos_karen
pasta_relatorios = 'Relatorio_Vendedores_Helder' if "Helder" in opcao else 'Relatorio_Vendedores_Karen'

# ---------- LEITURA DA REFERÊNCIA ----------
arquivo_referencia = st.file_uploader("📤 Envie o arquivo de referência (.xlsx):", type=["xlsx"])

if arquivo_referencia:
    df = carregar_dados_sql()
    df = df.drop_duplicates(subset='Raiz_CNPJ')
    referencia = pd.read_excel(arquivo_referencia, sheet_name='Planilha1')

    df['Raiz_CNPJ'] = df['Raiz_CNPJ'].astype(str).str.strip().str.zfill(14)
    referencia['Raiz_CNPJ Irão entrar'] = referencia['Raiz_CNPJ Irão entrar'].astype(str).str.strip().str.zfill(14)

    dict_transferencia = dict(zip(referencia['Raiz_CNPJ Irão entrar'], referencia['Nome_Vendedor']))

    # Atualiza o Nome_Vendedor do df conforme a referência
    df['Nome_Vendedor'] = df.apply(
        lambda row: dict_transferencia[row['Raiz_CNPJ']] if row['Raiz_CNPJ'] in dict_transferencia else row['Nome_Vendedor'],
        axis=1
    )

    # Agora você pode adicionar a data de entrada
    df['Data_Entrou_Carteira'] = np.where(
        df['Raiz_CNPJ'].isin(referencia['Raiz_CNPJ Irão entrar']),
        pd.Timestamp('2025-03-20'),
        pd.NaT
    )

    # Lógica de status
    data_limite = datetime.today() - timedelta(days=6*30)
    df['Status_Cliente'] = df['Data_Ultima_Venda_Grupo_CNPJ'].apply(
        lambda x: 'Nao Compra' if pd.isna(x) or x < data_limite else 'Compra'
    )

    df_historico = df[['Raiz_CNPJ', 'Nome_Vendedor']].dropna().drop_duplicates().reset_index(drop=True)

    df_filtrado = df[df['Nome_Vendedor'].isin(vendedores_ativos_helder + vendedores_ativos_karen)].reset_index(drop=True)

    contas_vao_rotacionar = df[
        (df['Status_Cliente'] == 'Nao Compra') &
        (df['Data_Abertura_Conta'] < data_limite) &
        ((df['Data_Entrou_Carteira'] < data_limite) | (df['Data_Entrou_Carteira'].isnull())) &
        (df['Grupo_Econômico_ID'].isnull())
    ]

    if "Helder" in opcao:
        contas_filtradas = contas_vao_rotacionar[contas_vao_rotacionar['Classificacao_Conta'].isin([5, 7])]
    else:
        contas_filtradas = contas_vao_rotacionar[~contas_vao_rotacionar['Classificacao_Conta'].isin([5, 7])]

    st.success(f"{len(contas_filtradas)} contas selecionadas para rotação.")

    def criar_tabela_historico():
        conn = sqlite3.connect('historico_rotacao.db')
        c = conn.cursor()
        c.execute('''
        CREATE TABLE IF NOT EXISTS historico_rotacao (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome_vendedor TEXT,
            conta_id INTEGER,
            tipo_rotacao TEXT,
            data_rotacao TEXT
        )
        ''')
        conn.commit()
        conn.close()

    def registrar_historico_rotacao(nome_vendedor, conta_id, tipo_rotacao, data_rotacao):
        conn = sqlite3.connect('historico_rotacao.db')
        c = conn.cursor()
        c.execute('''
        INSERT INTO historico_rotacao (nome_vendedor, conta_id, tipo_rotacao, data_rotacao)
        VALUES (?, ?, ?, ?)
        ''', (nome_vendedor, conta_id, tipo_rotacao, data_rotacao))
        conn.commit()
        conn.close()

    # ---------- FUNÇÃO DE ROTAÇÃO ----------
    def rotacionar_contas(df_contas, lista_vendedores, df_historico, limite_por_vendedor=50):
        contagem_vendedores = {v: 0 for v in lista_vendedores}
        novos_nomes = []
        indices_sobras = []

        for idx, row in df_contas.iterrows():
            cnpj = row['Raiz_CNPJ']
            vendedores_antigos = df_historico[df_historico['Raiz_CNPJ'] == cnpj]['Nome_Vendedor'].tolist()
            candidatos = [v for v in lista_vendedores if v not in vendedores_antigos and contagem_vendedores[v] < limite_por_vendedor]
            if candidatos:
                escolhido = np.random.choice(candidatos)
                contagem_vendedores[escolhido] += 1
                novos_nomes.append((idx, escolhido))
            else:
                indices_sobras.append(idx)

        df_resultado = df_contas.copy()
        data_hoje = pd.Timestamp.today().normalize()

        for idx, novo_vendedor in novos_nomes:
            df_resultado.at[idx, 'Nome_Vendedor'] = novo_vendedor
            df_resultado.at[idx, 'Data_Entrou_Carteira'] = data_hoje

        df_rotacionadas = df_resultado.loc[[idx for idx, _ in novos_nomes]].reset_index(drop=True)
        df_sobras = df_resultado.loc[indices_sobras].reset_index(drop=True)

        return df_rotacionadas, df_sobras
    
    def salvar_historico_rotacao(df_rotacionadas, nome_banco='historico_rotacao.db'):
        conn = sqlite3.connect(nome_banco)
        
        # Garante que a tabela exista
        conn.execute('''
            CREATE TABLE IF NOT EXISTS historico_rotacao (
                Raiz_CNPJ TEXT,
                Nome_Vendedor TEXT,
                Data_Entrou_Carteira DATE
            )
        ''')

        # Insere os dados novos
        df_rotacionadas[['Raiz_CNPJ', 'Nome_Vendedor', 'Data_Entrou_Carteira']].to_sql(
            'historico_rotacao',
            conn,
            if_exists='append',
            index=False
        )
        
        conn.close()

    # Botão de rotação
if st.button("🔁 Rodar contas agora"):
    contas_rotacionadas, contas_sobras = rotacionar_contas(contas_filtradas, vendedores_ativos, df_historico)

    st.success(f"{len(contas_rotacionadas)} contas rotacionadas com sucesso.")
    st.write("Contas rotacionadas:")
    st.dataframe(contas_rotacionadas)
    st.session_state["contas_rotacionadas"] = contas_rotacionadas
    st.session_state["contas_sobras"] = contas_sobras

    # Histórico
    historico_path = "historico_rotacoes_completo.xlsx"
    if os.path.exists(historico_path):
        historico_existente = pd.read_excel(historico_path)
        df_novos_historicos = pd.concat([historico_existente, contas_rotacionadas], ignore_index=True)
    else:
        df_novos_historicos = contas_rotacionadas.copy()

    df_novos_historicos = df_novos_historicos.drop_duplicates(subset=["Raiz_CNPJ", "Data_Entrou_Carteira"], keep="last")
    df_novos_historicos.to_excel(historico_path, index=False)

    st.write("Contas sem rotação (sem vendedor disponível):")
    st.dataframe(contas_sobras)

# Gerar downloads fora do if
def gerar_excel_download(df):
    from io import BytesIO
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Dados')
    return output.getvalue()

if "contas_rotacionadas" in st.session_state:
    st.download_button(
        "📥 Baixar contas rotacionadas",
        data=gerar_excel_download(st.session_state["contas_rotacionadas"]),
        file_name=f"historico_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
    )

if "contas_sobras" in st.session_state:
    st.download_button(
        "📥 Baixar contas sem rotação",
        data=gerar_excel_download(st.session_state["contas_sobras"]),
        file_name="contas_sobras.xlsx"
    )


else:
    st.info("Envie o arquivo de referência para continuar.")
    
st.markdown("---")
st.subheader("📊 Gerar Relatórios por Vendedor")

if st.button("📄 Gerar Relatórios Mensais"):
    if "contas_rotacionadas" not in st.session_state:
        st.warning("⚠️ Você precisa realizar a rotação antes de gerar o relatório.")
    else:
        def gerar_relatorios(df_atual, df_anterior, data_limite, data_rotacao, pasta_destino='Relatorio_Rotação'):
            os.makedirs(pasta_destino, exist_ok=True)

            data_rotacao = pd.to_datetime(data_rotacao).normalize()
            data_limite = pd.to_datetime(data_limite).normalize()

            for df in [df_atual, df_anterior]:
                df['Data_Entrou_Carteira'] = pd.to_datetime(df['Data_Entrou_Carteira'], errors='coerce')
                df['Data_Ultima_Venda_Grupo_CNPJ'] = pd.to_datetime(df['Data_Ultima_Venda_Grupo_CNPJ'], errors='coerce')

            writer = pd.ExcelWriter(f'{pasta_destino}/relatorio_mensal_completo.xlsx', engine='xlsxwriter')
            vendedores = df_atual['Nome_Vendedor'].dropna().unique()

            for vendedor in vendedores:
                atual_vend = df_atual[df_atual['Nome_Vendedor'] == vendedor].copy()
                anterior_vend = df_anterior[df_anterior['Nome_Vendedor'] == vendedor].copy()

                def montar_bloco(df, status):
                    bloco = df[['Nome_Vendedor', 'Razao_Social_Pessoas', 'Raiz_CNPJ',
                                'Data_Ultima_Venda_Grupo_CNPJ', 'Data_Entrou_Carteira']].copy()
                    bloco.insert(0, 'Status', status)
                    return bloco

                usados = set()
                blocos = []

                ativas = anterior_vend[
                    (anterior_vend['Data_Ultima_Venda_Grupo_CNPJ'] >= data_limite) &
                    (~anterior_vend['Raiz_CNPJ'].isin(usados))
                ]
                usados.update(ativas['Raiz_CNPJ'])
                blocos.append(montar_bloco(ativas, 'Ativa'))

                seis_meses_atras = data_rotacao - pd.DateOffset(months=6)
                recentes = anterior_vend[
                    (anterior_vend['Data_Entrou_Carteira'] >= seis_meses_atras) &
                    (anterior_vend['Data_Entrou_Carteira'] != data_rotacao) &
                    (~anterior_vend['Raiz_CNPJ'].isin(usados))
                ]
                usados.update(recentes['Raiz_CNPJ'])
                blocos.append(montar_bloco(recentes, 'Entraram Recentemente'))

                novas = atual_vend[
                    (atual_vend['Data_Entrou_Carteira'] == data_rotacao) &
                    (~atual_vend['Raiz_CNPJ'].isin(usados))
                ]
                usados.update(novas['Raiz_CNPJ'])
                blocos.append(montar_bloco(novas, 'Novas Recebidas'))

                cadastradas_recente = anterior_vend[
                    (anterior_vend['Data_Abertura_Conta'] >= seis_meses_atras) &
                    (~anterior_vend['Raiz_CNPJ'].isin(usados))
                ]
                usados.update(cadastradas_recente['Raiz_CNPJ'])
                blocos.append(montar_bloco(cadastradas_recente, 'Cadastrado Recentemente'))

                retiradas = anterior_vend[
                    (~anterior_vend['Raiz_CNPJ'].isin(atual_vend['Raiz_CNPJ'])) &
                    (~anterior_vend['Raiz_CNPJ'].isin(usados))
                ]
                usados.update(retiradas['Raiz_CNPJ'])
                blocos.append(montar_bloco(retiradas, 'Retiradas'))

                df_relatorio = pd.concat(blocos, ignore_index=True)
                df_relatorio = df_relatorio.drop_duplicates(subset='Raiz_CNPJ', keep='first')
                df_relatorio = df_relatorio.sort_values(['Status', 'Razao_Social_Pessoas']).reset_index(drop=True)

                if not df_relatorio.empty:
                    aba = vendedor[:31]
                    df_relatorio.to_excel(writer, sheet_name=aba, index=False)
                    nome_arquivo = f"{pasta_destino}/relatorio_{vendedor.replace(' ', '_')}_{data_rotacao.strftime('%Y-%m-%d')}.xlsx"
                    df_relatorio.to_excel(nome_arquivo, index=False)

            writer.close()
            return f"✅ Relatórios salvos em '{pasta_destino}' com sucesso!"

        resultado_relatorio = gerar_relatorios(
    		st.session_state["contas_rotacionadas"],
    		df_filtrado,
    		data_limite=data_limite,
    		data_rotacao=pd.Timestamp.today().normalize(),
			pasta_destino=pasta_relatorios
		)
        st.success(resultado_relatorio)
