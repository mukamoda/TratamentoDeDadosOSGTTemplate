import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
# DOCUMENTAAAAAAAAAAAAAR"!!!!!!
print("carregando...")
#path pros arquivos
arquivoHenkel = 'Catálogo Henkel.xlsx'
output_path='Catálogo Henkel - Relatório.xlsx'


try:
    planilha = pd.read_excel(arquivoHenkel, sheet_name="OSGT", engine="openpyxl")
    henkel_baseDeDados = pd.read_excel(arquivoHenkel, sheet_name="Base de Dados", engine="openpyxl")

    print("Abas carregadas!")

except FileNotFoundError:
    print(f"Erro: Arquivo não encontrado. Verifique se o caminho {arquivoHenkel} está correto")
    exit()

except ValueError as v:
    print(f"Erro: Umas das abas não foi encontrada no arquivo. Detalhes {v}")
    exit()

print('Otimizando...')
filtro_partnumbers=set(henkel_baseDeDados["PartNumber(IDH)"].dropna().unique())
print(f"Base de dados otimizada com {len(filtro_partnumbers)} Partnumbers únicos.")

print("Verificando os part numbers na aba OSGT")
planilha['Encontrado_na_base'] = planilha['S_PARTNUMBER'].isin(filtro_partnumbers)

planilha['S_OBRIGATORIO'] = planilha['S_OBRIGATORIO'].astype(str).str.strip().str.upper()

planilha['Obrigatório preenchido?'] = ''

planilha.loc[
   (planilha['S_OBRIGATORIO'] == 'TRUE') & planilha['S_VALOR'].notna(),
   'Obrigatório preenchido?'] = 'Sim'

planilha.loc[
   (planilha['S_OBRIGATORIO'] == 'TRUE') & planilha['S_VALOR'].isnull(),
   'Obrigatório preenchido?'] = 'Não'

print("\n----Resumo da verificação----")
print(planilha['Obrigatório preenchido?'].value_counts(dropna=False))
print("-"*25)

relatorio_planilha = planilha[(planilha['Obrigatório preenchido?'] == 'Não') & (planilha['S_NOME_ATRIBUTO'].astype(str).str.strip().str.startswith('0'))].copy()

relatorio_final = relatorio_planilha[['S_PARTNUMBER', 'S_NOME_ATRIBUTO']].copy()

relatorio_final.rename(columns={
    'S_PARTNUMBER': 'Partnumber',
    'S_NOME_ATRIBUTO': 'Atributos não preenchidos'
}, inplace=True)

relatorio_final.drop_duplicates(inplace=True)

idh_encontrados_df = planilha[planilha['Encontrado_na_base'] == True][['S_PARTNUMBER']].drop_duplicates()
idh_encontrados_df.rename(columns={'S_PARTNUMBER': 'IDHs NÃO ENCONTRADOS'}, inplace=True)
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    relatorio_final.to_excel(writer, sheet_name='Atributo', index=False)
    idh_encontrados_df.to_excel(writer, sheet_name='IDH', index=False)

try:
    print('Fazendo o bgl ficar bonitin...')
    wb = load_workbook(output_path)
   
    ws_atributo = wb['Atributo']
    ws_atributo.column_dimensions['A'].width = 15
    ws_atributo.column_dimensions['B'].width = 50

    tab_ref_atributo = f'A1:B{len(relatorio_final) + 1}'
    table_name_atributo = 'AtributoTabela'
    tab_atributo = Table(displayName=table_name_atributo, ref=tab_ref_atributo)
    style_atributo = TableStyleInfo(name='TableStyleMedium9', showFirstColumn=False, showLastColumn=False, showRowStripes=False, showColumnStripes=False)
    tab_atributo.tableStyleInfo = style_atributo
    ws_atributo.add_table(tab_atributo)

    ws_idh = wb['IDH']
    ws_idh.column_dimensions['A'].width = 20

    tab_ref_idh = f'A1:A{len(idh_encontrados_df) + 1}'
    table_name_idh = 'IDHÑEncontrados'
    tab_idh = Table(displayName=table_name_idh, ref=tab_ref_idh)
    style_idh = TableStyleInfo(name='TableStyleMedium9', showFirstColumn=False, showLastColumn=False, showRowStripes=False, showColumnStripes=False)
    tab_idh.tableStyleInfo = style_idh
    ws_idh.add_table(tab_idh)


    wb.save(output_path)
    print("feito certin a formatations")

except Exception as e:
    print(f"Erro na aplicancia: {e}")

print(f"\n Resultado salvo em {output_path}")
print("Processo concluído")

