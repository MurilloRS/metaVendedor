import calendar
import locale
from datetime import datetime, timedelta

import pandas as pd
import pyodbc
import streamlit as st

locale.setlocale(locale.LC_TIME, 'Portuguese_Brazil.1252')


def conectar_banco():
    try:
        conn = pyodbc.connect(
            'DRIVER={SQL Server};'
            'SERVER=192.168.1.22;'
            'DATABASE=TESTE;'
            'UID=sa;'
            'PWD=Moitgt2526'
        )
        # print("Conexão com o banco de dados bem-sucedida!")
        return conn
    except Exception as e:
        # print(f"Erro ao conectar ao banco de dados: {e}")
        return None


# # Testar a conexão
# if __name__ == "__main__":
#     conectar_banco()
# dt_inicio = '2024-11-01'
# dt_fim = '2024-11-30'
# cd_vend = 'LN5358'
# cd_tp_meta = 4
# valor_meta = 60000


def info_vend(cd_vend):
    conn = conectar_banco()
    if conn:
        query = f"SELECT cd_emp, cd_equipe FROM dbo.vendedor WHERE cd_vend = '{
            cd_vend}'"
        cursor = conn.cursor()
        cursor.execute(query)
        result = cursor.fetchone()
        conn.close()
        cd_emp = result[0]
        cd_equipe = result[1]
        return cd_emp, cd_equipe if result else None
    return None


# empresa, equipe = info_vend(cd_vend)
# print(empresa, equipe)


def insert_update_meta_grupo(dt_inicio, dt_fim, meta_grupo):
    conn = conectar_banco()
    if conn:
        try:
            cursor = conn.cursor()
            query = f"""
                IF EXISTS (SELECT 1 FROM dbo.meta
                            WHERE cd_tp_meta = 1 AND dt_inicio = '{dt_inicio}' AND dt_fim = '{dt_fim}')
                BEGIN
                    UPDATE dbo.meta
                    SET valor = {meta_grupo}
                    WHERE cd_tp_meta = 1 AND dt_inicio = '{dt_inicio}' AND dt_fim = '{dt_fim}'
                END
                ELSE
                BEGIN
                    INSERT  INTO dbo.meta (cd_tp_meta,dt_inicio, dt_fim ,valor)
                    VALUES (1,'{dt_inicio}','{dt_fim}',{meta_grupo})
                END"""
            cursor.execute(query)
            conn.commit()
            st.success(
                f"Meta atualizada ou inserida com sucesso para o Grupo MTF")
        except Exception as e:
            conn.rollback()
            st.error(f"Erro ao atualizar os dados: {e}")
        finally:
            conn.close()


def insert_update_meta_empresa(cd_emp, valor_meta, dt_inicio, dt_fim):
    conn = conectar_banco()
    if conn:
        try:
            cursor = conn.cursor()
            query_update_insert = f"""
                IF EXISTS (SELECT 1 FROM dbo.meta WHERE cd_emp = '{cd_emp}' AND dt_inicio = '{dt_inicio}' AND dt_fim = '{dt_fim}')
                BEGIN
                    UPDATE dbo.meta
                    SET valor = {valor_meta}
                    WHERE cd_emp = '{cd_emp}' AND dt_inicio = '{dt_inicio}' AND dt_fim = '{dt_fim}'
                END
                ELSE
                BEGIN
                    INSERT INTO dbo.meta (cd_emp, valor, dt_inicio, dt_fim)
                    VALUES ('{cd_emp}', {valor_meta},
                            '{dt_inicio}', '{dt_fim}')
                END
            """
            cursor.execute(query_update_insert)
            conn.commit()
            # st.success(
            #     f"Meta atualizada ou inserida com sucesso para a empresa {cd_emp}")
        except Exception as e:
            conn.rollback()
            st.error(f"Erro ao atualizar os dados da empresa {cd_emp}: {e}")
        finally:
            conn.close()


def verificar_meta_existente(dt_inicio, dt_fim, cd_tp_meta, cd_emp, cd_equipe, cd_vend):
    conn = conectar_banco()
    if conn:
        cursor = conn.cursor()
        query = f"""
            SELECT cd_meta
            FROM dbo.meta
            WHERE dt_inicio = '{dt_inicio}' AND dt_fim = '{dt_fim}'
                AND cd_tp_meta = {cd_tp_meta} --cod meta vendedor
                AND cd_emp = {cd_emp}
                AND cd_equipe = '{cd_equipe}'
                AND cd_vend = '{cd_vend}'
        """
        cursor.execute(query)
        result = cursor.fetchone()
        conn.close()
        return result[0] if result else None
    return None


# print(verificar_meta_existente('2024-11-01',
    #   '2024-11-30', cd_tp_meta, empresa, equipe, cd_vend))


def uptade_insert_meta_vendedor(dt_inicio, dt_fim, cd_tp_meta, cd_vend, cd_emp, cd_equipe, valor_meta):
    conn = conectar_banco()
    if conn:
        try:
            cursor = conn.cursor()

            query_update_insert = f"""
                    IF EXISTS (SELECT 1 FROM dbo.meta
                                WHERE dt_inicio = '{dt_inicio}' AND dt_fim = '{dt_fim}'
                                AND cd_tp_meta = {cd_tp_meta} AND cd_emp = {cd_emp}
                                AND cd_vend = '{cd_vend}' AND cd_equipe = '{cd_equipe}'
                                AND tipo_periodo = '{tipo_periodo}' )
                    BEGIN
                        UPDATE dbo.meta
                        SET valor = {valor_meta}
                        WHERE cd_tp_meta = {cd_tp_meta} AND cd_emp = {cd_emp}
                            AND cd_equipe = '{cd_equipe}'
                            AND dt_inicio = '{dt_inicio}' AND dt_fim = '{dt_fim}'
                    END
                    ELSE
                    BEGIN
                        INSERT INTO dbo.meta (tipo_periodo, dt_inicio, dt_fim, cd_tp_meta, cd_emp, cd_equipe, cd_vend, valor)
                        VALUES ('{tipo_periodo}','{dt_inicio}','{dt_fim}',{cd_tp_meta},{
                cd_emp},'{cd_equipe}', '{cd_vend}',{valor_meta})
                    END
            """
            cursor.execute(query_update_insert)
            conn.commit()
            st.success(
                f"Meta atualizada ou inserida com sucesso para o vendedor {cd_vend}")
        except Exception as e:
            conn.rollback()
            st.error(f"Erro ao atualizar os dados: {e}")
        finally:
            conn.close()


def DatePicker():
    # Seleção do mês e ano para o tipo "Grupo MTF"
    col1, col2 = st.columns(2)

    with col1:
        # Exibe os meses com seus nomes em português
        # Ignora o primeiro elemento vazio
        meses = list(calendar.month_name)[1:]
        mes_atual = datetime.now().month  # Índice do mês atual (1 a 12)
        mes_selecionado = st.selectbox(
            "Selecione o mês:", meses, index=mes_atual - 1)
        # Obter o índice do mês (1 a 12)
        mes_index = meses.index(mes_selecionado) + 1

    with col2:
        # Exibe anos em um intervalo de -5 a +5 anos em relação ao atual
        ano_atual = datetime.now().year
        anos = [ano_atual + i for i in range(-5, 6)]
        ano_selecionado = st.selectbox(
            # Ano atual como padrão
            "Selecione o ano:", anos, index=anos.index(ano_atual))

    # Calcular `dt_inicio` e `dt_fim` com base no mês/ano
    dt_inicio = f"{ano_selecionado}-{mes_index:02d}-01"
    ultimo_dia_mes = (datetime(int(ano_selecionado), mes_index, 1) +
                      timedelta(days=31)).replace(day=1) - timedelta(days=1)
    dt_fim = ultimo_dia_mes.strftime('%Y-%m-%d')
    return dt_inicio, dt_fim


def tela_meta_grupo():
    dt_inicio, dt_fim = DatePicker()

    # Campo para o valor da meta do grupo
    meta_grupo = st.number_input(
        "Insira o valor da meta do grupo:",
        min_value=0.0,
        format="%.2f"
    )

    # Exibição do valor formatado com separador de milhares
    meta_grupo_formatado = f"{meta_grupo:,.2f}".replace(
        ',', 'X').replace('.', ',').replace('X', '.')
    st.write(f"Meta do Grupo MTF: R$ {
        meta_grupo_formatado}, Período: {dt_inicio} a {dt_fim}")

    if st.button('Inserir ou Atualizar Metas'):
        if meta_grupo > 0:
            insert_update_meta_grupo(dt_inicio, dt_fim, meta_grupo)
        else:
            st.warning(
                "Meta não foi inserida ou atualizada. Verifique se o valor é maior que zero.")


def tela_meta_empresa():
    dt_inicio, dt_fim = DatePicker()

    st.write("")
    st.write("")

    conn = conectar_banco()
    query = f"SELECT cd_emp, nome_fant FROM dbo.empresa WHERE cd_emp IN (1,4,6,7,8,9,10)"
    cursor = conn.cursor()
    cursor.execute(query)
    empresas = cursor.fetchall()
    conn.close()

    metas = {}
    total_metas = 0.0

    # Exibindo o campo para inserir a meta para cada empresa
    for cd_emp, nome_fant in empresas:
        # Buscar o valor da meta já inserida para a empresa (se houver)
        conn = conectar_banco()
        query_meta = f"""
        SELECT valor
        FROM dbo.meta
        WHERE cd_emp = {cd_emp}
        AND dt_inicio = '{dt_inicio}'
        AND dt_fim = '{dt_fim}'
        """
        cursor = conn.cursor()
        cursor.execute(query_meta)
        meta_existente = cursor.fetchone()
        conn.close()

        # Se houver um valor de meta, preencher no campo de inserção
        valor_meta_inicial = float(
            meta_existente[0]) if meta_existente else 0.0
        # Exibe o nome da empresa e o campo de inserção de meta
        valor_meta = st.number_input(
            f"Meta para {nome_fant} ({cd_emp}):", min_value=0.0, value=valor_meta_inicial, format="%.2f")

        # Exibir o nome da empresa e o campo para inserir a meta
        # valor_meta = st.number_input(
        # f"Meta para {nome_fant} ({cd_emp}):", min_value=0.0, format="%.2f")
        metas[cd_emp] = valor_meta
        total_metas += valor_meta

    st.write("")
    st.write("")

    # Exibição do valor formatado com separador de milhares
    total_metas_formatado = f"{total_metas:,.2f}".replace(
        ',', 'X').replace('.', ',').replace('X', '.')
    st.write(f"Meta total: R$ {
        total_metas_formatado}, Período: {dt_inicio} a {dt_fim}")

    # Botão para inserir ou atualizar as metas
    if st.button("Confirmar Metas por Empresa"):
        # Contador para saber quantas empresas tiveram a meta atualizada
        empresas_atualizadas = 0
        for cd_emp, nome_fant in empresas:
            # Recupera o valor da meta para a empresa
            valor_meta = metas[cd_emp]
            if valor_meta > 0:  # Só insere ou atualiza a meta se o valor for maior que zero
                insert_update_meta_empresa(
                    cd_emp, valor_meta, dt_inicio, dt_fim)
                empresas_atualizadas += 1  # Incrementa o contador de empresas atualizadas
                st.success(f"Meta para a empresa {
                           nome_fant} inserida ou atualizada com sucesso. Valor: R$ {valor_meta:,.2f}")

        if empresas_atualizadas > 0:
            st.success(
                f"{empresas_atualizadas} metas foram inseridas ou atualizadas com sucesso.")
        else:
            st.warning(
                "Nenhuma meta foi inserida ou atualizada. Verifique se todos os valores são maiores que zero.")

    # # Campo para o valor da meta do grupo
    # meta_empresa = st.number_input(
    #     "Insira o valor da meta do grupo:",
    #     min_value=0.0,
    #     format="%.2f"
    # )

    # # Exibição do valor formatado com separador de milhares
    # meta_grupo_formatado = f"{meta_grupo:,.2f}".replace(
    #     ',', 'X').replace('.', ',').replace('X', '.')
    # st.write(f"Meta do Grupo MTF: R$ {
    #     meta_grupo_formatado}, Período: {dt_inicio} a {dt_fim}")

    # if st.button('Inserir ou Atualizar Metas'):
    #     insert_update_meta_grupo(dt_inicio, dt_fim, meta_grupo)


    # Título do App
st.title('App de Metas')


# Seleção do tipo de meta
tipos_de_meta = ['Vendedor', 'Vendedor por Fabricante',
                 'Equipe', 'Equipe por Fabricante', 'Grupo MTF', 'Empresa']
tipo_meta = st.selectbox("Selecione o tipo de meta:", tipos_de_meta)

# Variáveis para armazenar os valores finais
dt_inicio, dt_fim, meta_grupo = None, None, None
if tipo_meta == 'Grupo MTF':
    tela_meta_grupo()
elif tipo_meta == 'Empresa':
    tela_meta_empresa()

# elif tipo_meta == 'Vendedor':

#     opcao_periodo = st.radio("Selecione o tipo de período:", options=[
#         'Mensal', 'Trimestral'], index=0)

#     tipo_periodo = 'M' if opcao_periodo == 'Mensal' else 'T'
#     # Upload da planilha Excel
#     uploaded_file = st.file_uploader("Escolha o arquivo Excel", type=['xlsx'])

#     # Campos para selecionar a data de início e data de fim
#     col1, col2 = st.columns(2)
#     if tipo_periodo == 'M':
#         with col1:
#             dt_inicio = st.date_input(
#                 "Selecione a data de início do período",
#                 value=datetime.now().replace(day=1),  # Início do mês atual como padrão
#             )

#         with col2:
#             dt_fim = st.date_input(
#                 "Selecione a data de fim do período",
#                 value=datetime.now().replace(day=1).replace(month=datetime.now().month + 1) -
#                 timedelta(days=1),  # Fim do mês atual como padrão
#             )
#     else:
#         # Seleção de trimestre e ano
#         col1, col2 = st.columns(2)
#         with col1:
#             trimestre = st.selectbox(
#                 "Selecione o trimestre:",
#                 options=["1º Trimestre", "2º Trimestre",
#                          "3º Trimestre", "4º Trimestre"],
#                 index=0,
#             )
#         with col2:
#             ano = st.selectbox(
#                 "Selecione o ano:",
#                 # Anos disponíveis
#                 options=list(range(datetime.now().year,
#                              datetime.now().year + 5)),
#                 index=0,
#             )

#     # Definir dt_inicio e dt_fim com base no trimestre selecionado
#     if trimestre == "1º Trimestre":
#         dt_inicio = f"{ano}-01-01"
#         dt_fim = f"{ano}-03-31"
#     elif trimestre == "2º Trimestre":
#         dt_inicio = f"{ano}-04-01"
#         dt_fim = f"{ano}-06-30"
#     elif trimestre == "3º Trimestre":
#         dt_inicio = f"{ano}-07-01"
#         dt_fim = f"{ano}-09-30"
#     elif trimestre == "4º Trimestre":
#         dt_inicio = f"{ano}-10-01"
#         dt_fim = f"{ano}-12-31"

#     # Exibir as datas calculadas
#     st.write(f"**Período Selecionado:** {dt_inicio} até {dt_fim}")

# # Validação opcional para garantir que dt_inicio seja anterior a dt_fim
# if dt_inicio > dt_fim:
#     st.error("A data de início deve ser anterior ou igual à data de fim.")

# if uploaded_file:
#     df = pd.read_excel(uploaded_file, engine='openpyxl')
#     st.write("Dados Carregados")
#     st.dataframe(df)

#     df['META GERAL'] = df['META GERAL'].replace(
#         {'R$': '', ',': ''}, regex=True).astype(float)

#     if st.button('Inserir ou Atualizar Metas'):
#         for index, row in df.iterrows():
#             try:
#                 cd_vend = row['Codigo Vendedor']
#                 valor_meta = row['META GERAL']

#                 empresa, equipe = info_vend(cd_vend)

#                 if empresa and equipe:
#                     # Inserir ou atualizar os dados no banco
#                     uptade_insert_meta(
#                         dt_inicio, dt_fim, cd_tp_meta, cd_vend, empresa, equipe, valor_meta)
#                 else:
#                     st.warning(
#                         f"Vendedor {cd_vend} não encontrado no banco de dados.")
#             except Exception as e:
#                 st.error(f"Vendedor não encontrado {
#                          row['Codigo Vendedor']}")
