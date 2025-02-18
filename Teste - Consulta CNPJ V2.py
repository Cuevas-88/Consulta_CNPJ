import pandas as pd
import requests
import time
import random  # Para tempos de espera aleatórios
import streamlit as st

# Função para limpar o CNPJ (remover caracteres indesejados)
def limpar_cnpj(cnpj):
    cnpj = str(cnpj).strip().replace(".", "").replace("-", "").replace("/", "")
    return cnpj.zfill(14)  # Garante que tenha 14 dígitos

# Função para consultar o CNPJ
def consultar_cnpj(cnpj):
    cnpj = limpar_cnpj(cnpj)

    if len(cnpj) != 14:
        st.error(f"CNPJ inválido: {cnpj}")
        return None

    url = f"https://www.receitaws.com.br/v1/cnpj/{cnpj}"

    tentativas = 3  # Número máximo de tentativas antes de desistir

    for tentativa in range(tentativas):
        try:
            response = requests.get(url)

            if response.status_code == 200:
                dados_cnpj = response.json()

                if 'erro' in dados_cnpj:
                    st.error(f"Erro na API para o CNPJ {cnpj}: {dados_cnpj['erro']}")
                    return None

                return {
                    'CNPJ': dados_cnpj.get('cnpj', ''),
                    'Nome': dados_cnpj.get('nome', ''),
                    'Nome Fantasia': dados_cnpj.get('fantasia', ''),
                    'Natureza Jurídica': dados_cnpj.get('natureza_juridica', ''),
                    'Endereço': f"{dados_cnpj.get('logradouro', '')}, {dados_cnpj.get('numero', '')} - {dados_cnpj.get('bairro', '')}, {dados_cnpj.get('municipio', '')} - {dados_cnpj.get('uf', '')}",
                    'Telefone': dados_cnpj.get('telefone', ''),
                    'Email': dados_cnpj.get('email', ''),
                    'Atividade Principal': dados_cnpj.get('atividade_principal', [{}])[0].get('text', ''),
                    'Situação Cadastral': dados_cnpj.get('situacao', ''),
                    'Data de Abertura': dados_cnpj.get('abertura', ''),
                    'Quadro Societário': dados_cnpj['qsa'][0]['nome'] if 'qsa' in dados_cnpj and dados_cnpj['qsa'] else 'Não disponível'
                }

            elif response.status_code == 429:
                st.warning(f"🚨 Bloqueado! Tentando novamente em alguns segundos... (Tentativa {tentativa + 1}/{tentativas})")
                time.sleep(random.uniform(10, 20))  # Espera entre 10 e 20 segundos antes de tentar de novo

            else:
                st.error(f"Erro HTTP {response.status_code} para o CNPJ {cnpj}")
                return None

        except requests.exceptions.RequestException as e:
            st.error(f"Erro de conexão para o CNPJ {cnpj}: {e}")
            return None

    st.error(f"❌ Não foi possível consultar o CNPJ {cnpj} após {tentativas} tentativas.")
    return None

# Função para processar os CNPJs da planilha e gerar o DataFrame
def processar_cnpjs(arquivo_excel):
    df = pd.read_excel(arquivo_excel, dtype=str)  # Carregar como string para evitar problemas de formatação
    
    if "CNPJ" not in df.columns:
        st.error("Erro: A planilha deve conter uma coluna chamada 'CNPJ'.")
        return None

    dados = []
    cnpjs_com_erro = []

    st.write("Prévia dos dados da planilha:")
    st.write(df.head())

    for cnpj in df["CNPJ"].dropna():  # Remove valores vazios
        cnpj_limpo = limpar_cnpj(cnpj)
        dados_cnpj = consultar_cnpj(cnpj_limpo)

        if dados_cnpj:
            dados.append(dados_cnpj)
        else:
            cnpjs_com_erro.append(cnpj_limpo)  # Guardar CNPJs que falharam
        
        time.sleep(random.uniform(3, 7))  # Evita bloqueio da API

    if cnpjs_com_erro:
        st.write("\n🔄 Tentando novamente os CNPJs que deram erro...")
        time.sleep(30)  # Espera antes de nova tentativa

        for cnpj in cnpjs_com_erro:
            dados_cnpj = consultar_cnpj(cnpj)
            if dados_cnpj:
                dados.append(dados_cnpj)
            else:
                st.warning(f"⚠️ Falha final na consulta do CNPJ {cnpj}")
            time.sleep(random.uniform(3, 7))  # Intervalo entre requisições

    df_resultado = pd.DataFrame(dados)
    df_erros = pd.DataFrame({'CNPJ com Erro': cnpjs_com_erro})

    return df_resultado, df_erros

# Função para fazer o download da planilha com os resultados
def download_planilha(df_resultado, df_erros):
    if df_resultado is None or df_resultado.empty:
        st.error("Nenhum dado para salvar.")
        return
    
    arquivo_saida = "dados_cnpjs.xlsx"
    with pd.ExcelWriter(arquivo_saida) as writer:
        df_resultado.to_excel(writer, sheet_name='Dados CNPJs', index=False)
        df_erros.to_excel(writer, sheet_name='CNPJs com Erro', index=False)
    
    st.download_button(label="Download da Planilha", data=open(arquivo_saida, "rb"), file_name=arquivo_saida, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Interface Streamlit
st.title("Consulta de CNPJs")
st.write("Faça o upload de uma planilha com a coluna 'CNPJ' para realizar a consulta.")

# Widget para upload do arquivo Excel
arquivo_upload = st.file_uploader("Escolha um arquivo Excel", type=['xlsx'])

if arquivo_upload is not None:
    df_resultado, df_erros = processar_cnpjs(arquivo_upload)

    if df_resultado is not None:
        st.write(df_resultado)
        download_planilha(df_resultado, df_erros)
