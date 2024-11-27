from datetime import datetime

import pandas as pd
import pyodbc
import streamlit as st


# Função para conectar ao banco de dados
def conectar_banco():
    try:
        conn = pyodbc.connect(
            'DRIVER={SQL Server};SERVER=192.168.2.10;DATABASE=MOINHO;UID=mrs;PWD=100881*Mr')
        # conn = pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.1.22;DATABASE=TESTE;UID=mrs;PWD=100881*Mr')
        return conn
    except Exception as e:
        st.error(f"Erro ao conectar ao banco de dados: {e}")
        return None

# Função para buscar a empresa do vendedor


def buscar_empresa(cod_vend):
    conn = conectar_banco()
    if conn:
        query = "SELECT cd_emp FROM dbo.vendedor WHERE cd_vend = ?"
        cursor = conn.cursor()
        cursor.execute(query, cod_vend)
        result = cursor.fetchone()
        conn.close()
        return result[0] if result else None
    return None


# Função para verificar se a meta já existe na tabela prev_vda
def verificar_meta_existente(cd_emp, dt_ini_modelo):
    conn = conectar_banco()
    if conn:
        query = """
            SELECT cd_prev_vda
            FROM dbo.prev_vda
            WHERE cd_emp = ? AND mes_ref = ? AND cd_tp_prev = 'FATUVL'
        """
        cursor = conn.cursor()
        cursor.execute(query, cd_emp, dt_ini_modelo)
        result = cursor.fetchone()
        conn.close()
        return result[0] if result else None
    return None

# Função para inserir um novo registro na tabela prev_vda


def inserir_meta(cd_emp, dt_ini_modelo, dt_fim_modelo):
    conn = conectar_banco()
    if conn:
        try:
            cursor = conn.cursor()

            # Obter o próximo SeqMeta
            seq_query = "SELECT ISNULL(MAX(cd_prev_vda), 0) + 1 FROM dbo.prev_vda"
            cursor.execute(seq_query)
            seq_meta = cursor.fetchone()[0]

            # Inserir na tabela prev_vda, mas só se não houver duplicata
            insert_prev_vda = """
                IF NOT EXISTS (
                    SELECT 1 FROM dbo.prev_vda 
                    WHERE cd_emp = ? AND mes_ref = ? AND cd_tp_prev = 'FATUVL'
                )
                BEGIN
                    INSERT INTO dbo.prev_vda (cd_prev_vda, mes_ref, dt_ini_modelo, dt_fim_modelo, cd_tp_prev, cd_emp, TipoPeriodoMetaID, DataInicio, DataFim)
                    VALUES (?, ?, ?, ?, 'FATUVL', ?, 1, ?, ?)
                END
            """
            cursor.execute(insert_prev_vda, cd_emp, dt_ini_modelo, seq_meta, dt_ini_modelo,
                           dt_ini_modelo, dt_ini_modelo, cd_emp, dt_ini_modelo, dt_fim_modelo)

            conn.commit()
            conn.close()
            return seq_meta
        except Exception as e:
            conn.rollback()
            st.error(f"Erro ao inserir a meta: {e}")
            conn.close()
            return None


# Função para atualizar a tabela it_prev_vda
def atualizar_it_prev_vda(seq_meta, cod_vend, valor_meta):
    conn = conectar_banco()
    if conn:
        try:
            cursor = conn.cursor()

            # Atualizar ou inserir na tabela it_prev_vda, garantindo que use o cd_prev_vda correto
            update_it_prev_vda = """
                IF EXISTS (SELECT 1 FROM dbo.it_prev_vda WHERE cd_prev_vda = ? AND cd_vend = ?)
                BEGIN
                    UPDATE dbo.it_prev_vda
                    SET valor = ?
                    WHERE cd_prev_vda = ? AND cd_vend = ?
                END
                ELSE
                BEGIN
                    INSERT INTO dbo.it_prev_vda (cd_prev_vda, cd_vend, cd_tp_prev_det, valor)
                    VALUES (?, ?, 'FATUVL', ?)
                END
            """
            cursor.execute(update_it_prev_vda, seq_meta, cod_vend, valor_meta,
                           seq_meta, cod_vend, seq_meta, cod_vend, valor_meta)

            conn.commit()
            st.success(
                f"Meta atualizada ou inserida com sucesso para o vendedor {cod_vend}")
        except Exception as e:
            conn.rollback()
            st.error(f"Erro ao atualizar os dados: {e}")
        finally:
            conn.close()

# Função principal para inserir ou atualizar dados


def inserir_ou_atualizar_dados(cd_emp, cod_vend, valor_meta, dt_ini_modelo, dt_fim_modelo):
    # Verifica se a meta já existe na tabela prev_vda
    seq_meta = verificar_meta_existente(cd_emp, dt_ini_modelo)

    if seq_meta:
        # Se a meta já existir, usar o cd_prev_vda existente para atualizar it_prev_vda
        atualizar_it_prev_vda(seq_meta, cod_vend, valor_meta)
    else:
        # Se a meta não existir, inserir nova meta e depois usar o novo cd_prev_vda para atualizar it_prev_vda
        seq_meta = inserir_meta(cd_emp, dt_ini_modelo, dt_fim_modelo)
        if seq_meta is not None:
            atualizar_it_prev_vda(seq_meta, cod_vend, valor_meta)


# Título do App
st.title('App de Metas por Vendedor')

# Upload da planilha Excel
uploaded_file = st.file_uploader("Escolha o arquivo Excel", type=['xlsx'])

# Campo para selecionar o mês e o ano da meta (seleção de uma data qualquer no mês desejado)
data_meta = st.date_input(
    "Selecione uma data no mês de referência", value=datetime.now())

if uploaded_file:
    # Carregar a planilha
    df = pd.read_excel(uploaded_file, engine='openpyxl')
    st.write("Dados carregados:")
    st.dataframe(df)

    # Limpar formatação e converter valores numéricos
    df['META GERAL'] = df['META GERAL'].replace(
        {'R$': '', ',': ''}, regex=True).astype(float)

    # Extrair o mês e o ano da data selecionada
    # mes_ref = data_meta.strftime('%Y-%m')  # Definir o ano e o mês
    # print('mes_ref: ',mes_ref)
    dt_ini_modelo = data_meta.strftime('%Y-%m') + '-01'
    print('dt_ini_modelo: ', dt_ini_modelo)
    # Definir o último dia do mês
    if data_meta.month == 2:  # Fevereiro (considera ano bissexto)
        if data_meta.year % 4 == 0 and (data_meta.year % 100 != 0 or data_meta.year % 400 == 0):
            dt_fim_modelo = f"{data_meta.year}-02-29"
        else:
            dt_fim_modelo = f"{data_meta.year}-02-28"
    elif data_meta.month in [4, 6, 9, 11]:  # Meses com 30 dias
        dt_fim_modelo = f"{data_meta.year}-{data_meta.month:02d}-30"
    else:  # Meses com 31 dias
        dt_fim_modelo = f"{data_meta.year}-{data_meta.month:02d}-31"

    print('dt_fim_modelo: ', dt_fim_modelo)

    # Botão para processar os dados e inserir ou atualizar no banco
    if st.button('Inserir ou Atualizar dados no banco'):
        for index, row in df.iterrows():
            vendedor = row['Codigo Vendedor']
            valor_meta = row['META GERAL']

            # Buscar a empresa do vendedor
            empresa = buscar_empresa(vendedor)

            if empresa:
                # Inserir ou atualizar os dados no banco
                inserir_ou_atualizar_dados(
                    empresa, vendedor, valor_meta, dt_ini_modelo, dt_fim_modelo)
            else:
                st.error(f"Empresa não encontrada para o vendedor {vendedor}")


# poetry shell
# streamlit run app.py
