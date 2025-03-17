import pandas as pd
import streamlit as st
import numpy as np
from io import BytesIO

# Exemplo de DataFrame de vendedores
vendedores_ativos_helder = ['Bryan Casarotto',
'Daniel Nunes de Paula Junior',
'Daniele Schmitz',
'GILBERTO LIMA DE PINHO JUNIOR',
'Kelli de Almeida Ferreira',
'Laura Vitoria Da Silveira Trindade',
'Leonardo Bianchi',
'Letícia Eduarda Cruz',
'LUCAS VASCONCELOS BATTAGLIA KRAUSE',
'RONALDO DA COSTA BARRIOS',
'Ruan da Silva',
'TALIA LINS RAMOS',
'Willian Luiz Pereira'
]
data = {'Nome_Vendedor': vendedores_ativos_helder}
df_vendedores_helder = pd.DataFrame(data)

# Carregar o DataFrame de clientes cadastrados
df_churn_revenda = pd.read_excel('./data/churn_por_class_5_7.xlsx', sheet_name='Sheet1')
df_churn_revenda = df_churn_revenda[['Raiz_CNPJ', 'Conta_ID', 'tipo_conta', 'Razao_Social_Pessoas', 'CNPJ', 'Grupo_Econômico_ID', 'Grupo_Econômico_Nome', 'Classificacao_Pessoa', 'Categoria_Porte']].drop_duplicates(subset='Raiz_CNPJ')

# Função para distribuir clientes entre os vendedores
def distribuir_clientes(vendedores, clientes, clientes_por_vendedor=50):
    # Número de vendedores
    num_vendedores = len(vendedores)
    
    # Número total de clientes
    num_clientes = len(clientes)
    
    # Verificar se temos clientes suficientes para distribuir
    if num_clientes < clientes_por_vendedor * num_vendedores:
        st.warning(f"Temos apenas {num_clientes} clientes e não conseguimos distribuir {clientes_por_vendedor} por vendedor.")
        # Ajustar o número de clientes por vendedor para não ultrapassar os clientes disponíveis
        clientes_por_vendedor = num_clientes // num_vendedores
    
    # Criar uma lista de clientes para distribuição
    clientes_distribuidos = np.array_split(clientes['Razao_Social_Pessoas'], num_vendedores)
    
    # Limitar a distribuição para exatamente 'clientes_por_vendedor' clientes por vendedor
    tabela_distribuicao = []
    for i, vendedor in enumerate(vendedores['Nome_Vendedor']):
        # Garantir que não ultrapasse o limite de clientes por vendedor
        clientes_vendedor = clientes_distribuidos[i][:clientes_por_vendedor]
        for cliente in clientes_vendedor:
            tabela_distribuicao.append([vendedor, cliente])
    
    tabela_distribuicao = pd.DataFrame(tabela_distribuicao, columns=['Vendedor', 'Cliente'])

    # Remover os clientes que já foram distribuídos
    clientes_distribuidos_ids = pd.concat([df_churn_revenda[df_churn_revenda['Razao_Social_Pessoas'] == cliente] for cliente in tabela_distribuicao['Cliente']])['Raiz_CNPJ'].unique()
    clientes_sobrando = clientes[~clientes['Raiz_CNPJ'].isin(clientes_distribuidos_ids)]

    return tabela_distribuicao, clientes_sobrando

# Função para criar um arquivo Excel a partir de um DataFrame
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Clientes Distribuídos')
    return output.getvalue()

# Streamlit UI
st.title('Distribuição de Clientes para Vendedores')

# Exibir os DataFrames de vendedores e clientes
st.subheader('Vendedores')
st.write(df_vendedores_helder)

st.subheader('Clientes Inativos Cadastrados')
st.write(df_churn_revenda)

# Variáveis de estado para persistir entre execuções
if 'tabela_distribuicao' not in st.session_state:
    st.session_state.tabela_distribuicao = pd.DataFrame()
    st.session_state.clientes_sobrando = pd.DataFrame()

# Função que será chamada para atualizar a interface com os dados
def mostrar_clientes(vendedor_selecionado, tabela_distribuicao):
    # Filtrando os clientes do vendedor selecionado
    clientes_vendedor = tabela_distribuicao[tabela_distribuicao['Vendedor'] == vendedor_selecionado]
    
    st.subheader(f'Clientes Distribuídos para {vendedor_selecionado}')
    st.write(clientes_vendedor)

    # Gerar o arquivo Excel para o vendedor selecionado
    excel_data = to_excel(clientes_vendedor)
    
    # Botão de download do arquivo Excel
    st.download_button(
        label="Baixar tabela de clientes",
        data=excel_data,
        file_name=f"clientes_{vendedor_selecionado}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Botão para distribuir os clientes
if st.button('Distribuir Clientes'):
    tabela_distribuicao, clientes_sobrando = distribuir_clientes(df_vendedores_helder, df_churn_revenda)
    
    # Armazenando a tabela de distribuição e os clientes sobrando no estado da sessão
    st.session_state.tabela_distribuicao = tabela_distribuicao
    st.session_state.clientes_sobrando = clientes_sobrando
    
    # Exibir a tabela de distribuição
    st.subheader('Tabela de Distribuição de Clientes')
    st.write(tabela_distribuicao)
    
    # Exibir clientes remanescentes
    st.subheader('Clientes Remanescentes (Próxima Rodada)')
    st.write(clientes_sobrando.reset_index(drop=True))

    # Gerar o arquivo Excel para os clientes remanescentes
    excel_data_sobrando = to_excel(clientes_sobrando)
    
    # Botão de download do arquivo Excel para clientes remanescentes
    st.download_button(
        label="Baixar clientes remanescentes",
        data=excel_data_sobrando,
        file_name="clientes_remanescentes.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
# Filtro de Vendedor para visualização separada
if len(st.session_state.tabela_distribuicao) > 0:
    vendedor_selecionado = st.selectbox('Escolha um Vendedor:', df_vendedores_helder['Nome_Vendedor'])
    
    # Chama a função para mostrar os clientes do vendedor
    mostrar_clientes(vendedor_selecionado, st.session_state.tabela_distribuicao)