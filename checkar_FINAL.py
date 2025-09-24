# -*- coding: UTF-8 -*-
#autor: Mikhael Prajna Nunes Vieira Rocha Jorge (Mikhael Jorge - DSV) / Última alteração: 18.09.2025

import pandas as pd
import re
import numpy as np

# --- 1. CONFIGURAÇÃO GERAL ---
ARQUIVO_ENTRADA = 'Catálogo Henkel.xlsx'
ARQUIVO_SAIDA = 'Catálogo Henkel - Processado.xlsx'
NOME_DA_ABA = 'OSGT'

# Nomes das colunas usados em todos os processos
COL_VALOR = 'S_VALOR'
COL_OBRIGATORIO = 'S_OBRIGATORIO'
COL_NOME_ATRIBUTO = 'S_NOME_ATRIBUTO'
COL_LISTA_ESTATICA = 'S_LISTA_ESTATICA'
COL_TIPO_CAMPO = 'S_TIPO_CAMPO'

UNIDADES_DE_MEDIDA = ['°C', 'KG/L', 'MM', 'CM³', 'L/MIN', 'KW', 'A']
UNIDADES_DE_MEDIDA_UPPER = [u.upper() for u in UNIDADES_DE_MEDIDA]

# --- 2. FUNÇÕES DE PROCESSAMENTO E VALIDAÇÃO ---

def eh_booleano_valido(valor):
    """Verifica se o valor é 'SIM' ou 'NÃO'."""
    if pd.isna(valor):
        return False
    return str(valor).strip() in ['SIM', 'NÃO']

def eh_inteiro_valido(valor):
    """Verifica se o valor pode ser convertido para um inteiro (aceita ',' como decimal)."""
    if pd.isna(valor):
        return False
    try:
        val_str = str(valor).strip().replace(',', '.')
        return float(val_str) == int(float(val_str))
    except (ValueError, TypeError):
        return False

def eh_real_valido(valor):
    """Verifica se o valor pode ser convertido para um número real (aceita ',' como decimal)."""
    if pd.isna(valor):
        return False
    try:
        float(str(valor).strip().replace(',', '.'))
        return True
    except (ValueError, TypeError):
        return False
    
def processar_e_limpar_valor(valor, tipo_campo):
    """Função aprimorada que lida com vírgulas e unidades coladas."""
    if not isinstance(valor, str):
        return valor
    match = re.match(r'^\s*([0-9.,]+)(\s*)?([A-Z°/³]+)\s*$', valor, re.IGNORECASE)
    if not match:
        return valor
    numero_str = match.group(1).strip()
    unidade_str = match.group(3).strip()
    if unidade_str.upper() not in UNIDADES_DE_MEDIDA_UPPER:
        return valor
    numero_str_ponto = numero_str.replace(',', '.')
    if tipo_campo == 'NUMERO_INTEIRO':
        if eh_inteiro_valido(numero_str_ponto):
            return numero_str
    elif tipo_campo == 'NUMERO_REAL':
        if eh_real_valido(numero_str_ponto):
            return numero_str
    return valor

def destacar_erros(row):
    """
    Função FINAL para aplicar no DataFrame e retornar o estilo da célula,
    seguindo a lógica de múltiplos cenários de erro e cores.
    """
    # Define os nomes das colunas para clareza
    COL_OBRIGATORIO = 'S_OBRIGATORIO'
    
    # Pega os valores da linha atual
    nome_atributo = str(row[COL_NOME_ATRIBUTO])
    tipo_campo = row[COL_TIPO_CAMPO]
    valor_cel = row[COL_VALOR]
    eh_obrigatorio = row[COL_OBRIGATORIO] == True
    
    styles = [''] * len(row)
    
    # --- INÍCIO DA NOVA LÓGICA DE VALIDAÇÃO ---

    # Cenário 1: Erros que devem ser pintados de VERMELHO
    erro_vermelho = False
    
    # 1a: Atributo obrigatório ('0-') que está vazio.
    # Esta regra agora inclui LISTA_ESTATICA novamente.
    if nome_atributo.strip().startswith('0-') and eh_obrigatorio and pd.isna(valor_cel):
        erro_vermelho = True
    
    # 1b: Qualquer atributo que foi preenchido, mas com um valor inválido.
    # A exceção é LISTA_ESTATICA, que não tem validação de tipo de dado aqui.
    elif pd.notna(valor_cel) and tipo_campo != 'LISTA_ESTATICA': 
        if pd.notna(tipo_campo):
            if tipo_campo == 'BOOLEANO' and not eh_booleano_valido(valor_cel):
                erro_vermelho = True
            elif tipo_campo == 'NUMERO_REAL' and not eh_real_valido(valor_cel):
                erro_vermelho = True
            elif tipo_campo == 'NUMERO_INTEIRO' and not eh_inteiro_valido(valor_cel):
                erro_vermelho = True

    # Aplica a cor vermelha se um dos erros acima for encontrado
    if erro_vermelho:
        idx = row.index.get_loc(COL_VALOR)
        styles[idx] = 'background-color: red'
        return styles # Retorna imediatamente após encontrar um erro vermelho

    # --- NOVO CENÁRIO 2: Erros que devem ser pintados de LARANJA ---
    erro_laranja = False
    
    # Atributo obrigatório que NÃO começa com '0-' e está vazio.
    if not nome_atributo.strip().startswith('0-') and eh_obrigatorio and pd.isna(valor_cel):
        erro_laranja = True

    # Aplica a cor laranja se o erro acima for encontrado
    if erro_laranja:
        idx = row.index.get_loc(COL_VALOR)
        styles[idx] = 'background-color: orange'

    return styles

# --- 3. EXECUÇÃO PRINCIPAL ---

# --- ETAPA 0: LEITURA DO ARQUIVO ---
print(f'Lendo o arquivo {ARQUIVO_ENTRADA}...')
try:
    df = pd.read_excel(ARQUIVO_ENTRADA, sheet_name=NOME_DA_ABA, engine='openpyxl')
    print('Arquivo carregado com sucesso!')
except FileNotFoundError:
    print(f'ERRO: Arquivo "{ARQUIVO_ENTRADA}" não encontrado.')
    exit()
except ValueError as e:
    print(f'ERRO: A aba "{NOME_DA_ABA}" não foi encontrada ou está com erro. Detalhes: {e}')
    exit()

# --- ETAPA 1: REESTRUTURAÇÃO INICIAL ---
print("\n--- ETAPA 1: Reestruturando as linhas de especificação '99' e '999'... ---")
# (O código da Etapa 1 permanece o mesmo)
regex_split = re.compile(
    r'^(999|99)'
    r'('
    r'(?:'
    r'[\s.:\-–—\(\)]+'
    r'|'
    r'outros|especifique|demais|s[oó]lido|etc\.?'
    r')+'
    r')?'
    r'(.*)',
    re.IGNORECASE
)
linhas_processadas = []
i = 0
while i < len(df):
    linha_atual = df.iloc[i]
    linha_seguinte = df.iloc[i + 1] if (i + 1) < len(df) else None
    valor_bruto = str(linha_atual[COL_VALOR])
    valor_val = valor_bruto.strip().replace(u'\xa0', ' ')
    match_split = regex_split.match(valor_val)
    if match_split:
        numero = match_split.group(1)
        especificacao = match_split.group(3).strip().strip('().:;-–—= ')
        nome_atributo_seguinte = str(linha_seguinte[COL_NOME_ATRIBUTO]) if linha_seguinte is not None else ''
        if especificacao and linha_seguinte is not None and nome_atributo_seguinte.strip().startswith('1'):
            linha_atual_modificada = linha_atual.to_dict()
            linha_atual_modificada[COL_VALOR] = numero
            linhas_processadas.append(linha_atual_modificada)
            linha_seguinte_modificada = linha_seguinte.to_dict()
            valor_existente = str(linha_seguinte_modificada[COL_VALOR])
            if valor_existente and valor_existente.lower() != 'nan':
                linha_seguinte_modificada[COL_VALOR] = f'{especificacao} | {valor_existente}'
            else:
                linha_seguinte_modificada[COL_VALOR] = especificacao
            linhas_processadas.append(linha_seguinte_modificada)
            i += 2
            continue
        else:
            linha_atual_modificada = linha_atual.to_dict()
            linha_atual_modificada[COL_VALOR] = numero
            linhas_processadas.append(linha_atual_modificada)
            i += 1
            continue
    linhas_processadas.append(linha_atual.to_dict())
    i += 1
df_processado = pd.DataFrame(linhas_processadas)
print("Análise estrutural concluída.")

# --- ETAPA 2: PADRONIZAÇÃO E PENTE FINO ---
print("\n--- ETAPA 2: Aplicando padronizações e correções... ---")
# Pente Fino
regex_fonte_da_verdade = re.compile(r'(999|99)[\s\-–—]*?(?:outros|especifique)', re.IGNORECASE)
correcoes_feitas = 0
for index, row in df_processado.iterrows():
    valor_atual = str(row[COL_VALOR]).strip()
    if valor_atual in ['99', '999']:
        lista_estatica_val = str(row[COL_LISTA_ESTATICA])
        numero_correto = None
        match_verdade = regex_fonte_da_verdade.search(lista_estatica_val)
        if match_verdade:
            numero_correto = match_verdade.group(1)
        if numero_correto and numero_correto != valor_atual:
            df_processado.at[index, COL_VALOR] = numero_correto
            correcoes_feitas += 1
print(f"Pente fino concluído! {correcoes_feitas} correções foram aplicadas.")
# Padronizações gerais
val_nulos = ['#N/A', 'NAN']
df_processado[COL_VALOR] = df_processado[COL_VALOR].astype(str).str.strip().str.upper()
df_processado[COL_VALOR] = df_processado[COL_VALOR].replace(val_nulos, np.nan)
print("Valores padronizados para maiúsculo e nulos tratados.")
# Limpeza de Booleanos '0'
print("Limpando valores '0' de campos booleanos...")
filtro_booleano_zero = (df_processado[COL_TIPO_CAMPO] == 'BOOLEANO') & (df_processado[COL_VALOR] == '0')
df_processado.loc[filtro_booleano_zero, COL_VALOR] = np.nan
print("Limpeza de booleanos concluída.")


# --- ETAPA 3: IDENTIFICAÇÃO DE ERROS ---
print("\n--- ETAPA 3: Identificando células com erro de validação... ---")
df_processado['TEM_ERRO'] = False 
for index, row in df_processado.iterrows():
    tipo_campo = row[COL_TIPO_CAMPO]
    valor_cel = row[COL_VALOR]
    erro_encontrado = False
    if pd.isna(valor_cel):
        if row[COL_OBRIGATORIO] == True:
            erro_encontrado = True
    elif pd.notna(tipo_campo):
        if tipo_campo == 'BOOLEANO' and not eh_booleano_valido(valor_cel):
            erro_encontrado = True
        elif tipo_campo == 'NUMERO_REAL' and not eh_real_valido(valor_cel):
            erro_encontrado = True
        elif tipo_campo == 'NUMERO_INTEIRO' and not eh_inteiro_valido(valor_cel):
            erro_encontrado = True
    if erro_encontrado:
        df_processado.at[index, 'TEM_ERRO'] = True
print("Identificação de erros concluída.")


# --- ETAPA 3.5: TENTATIVA DE CORREÇÃO AUTOMÁTICA ---
print("\n--- ETAPA 3.5: Tentando corrigir erros de unidade em campos numéricos... ---")
updates_count = 0
filtro_para_correcao = (df_processado['TEM_ERRO'] == True) & \
                       (df_processado[COL_TIPO_CAMPO].isin(['NUMERO_REAL', 'NUMERO_INTEIRO']))
df_para_correcao = df_processado[filtro_para_correcao]
for index, row in df_para_correcao.iterrows():
    valor_original = row[COL_VALOR]
    tipo_campo = row[COL_TIPO_CAMPO]
    valor_processado = processar_e_limpar_valor(valor_original, tipo_campo)
    if valor_processado != valor_original:
        df_processado.at[index, COL_VALOR] = valor_processado
        updates_count += 1
        df_processado.at[index, 'TEM_ERRO'] = False
print(f"Tentativa de correção concluída. {updates_count} células foram corrigidas.")
df_processado = df_processado.drop(columns=['TEM_ERRO'])


# --- ETAPA 4: FORMATAÇÃO FINAL E SALVAMENTO ---
print("\n--- ETAPA 4: Aplicando formatação final e salvando o arquivo... ---")
styler = df_processado.style.apply(destacar_erros, axis=1)
print("Células com erros remanescentes marcadas para serem coloridas.")
try:
    styler.to_excel(ARQUIVO_SAIDA, engine='openpyxl', index=False)
    print(f'Processo concluído com sucesso! O arquivo "{ARQUIVO_SAIDA}" foi criado.')
except Exception as e:
    print(f'ERRO: Não foi possível salvar o arquivo de saída. Detalhes: {e}')
    print("DICA: Verifique se a biblioteca 'jinja2' está instalada ('pip install jinja2').")