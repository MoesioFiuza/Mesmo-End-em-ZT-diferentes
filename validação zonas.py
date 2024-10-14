import pandas as pd

planilha = r'C:\Users\moesios\Desktop\qwe\BANCO DE DADOS_ATUALIZADO REV01_v2.xlsx'
endereços = r'C:\Users\moesios\Desktop\AJUSTES P14\endereços.xlsx'

df_deslocamentos = pd.read_excel(planilha, sheet_name='DESLOCAMENTO')
df_endereços = pd.read_excel(endereços, sheet_name='ENDEREÇOS')

df_merged_o = df_deslocamentos.merge(df_endereços, left_on='END_O', right_on='ENDS', how='left')
df_merged_d = df_deslocamentos.merge(df_endereços, left_on='END_D_B', right_on='ENDS', how='left')

df_merged_o['ZONA'] = df_merged_o['ZONA_O']
df_merged_d['ZONA'] = df_merged_d['ZONA_D']

df_merged = pd.concat([df_merged_o, df_merged_d], ignore_index=True)
df_merged = df_merged[['ENDS', 'ZONA']].drop_duplicates()

df_result = df_endereços.merge(df_merged, on='ENDS', how='left')

df_zonas_diferentes = df_result.groupby('ENDS')['ZONA'].apply(lambda x: ', '.join(map(str, x.dropna().unique()))).reset_index()
df_zonas_diferentes.columns = ['ENDS', 'ZONAS']

df_result = df_result.merge(df_zonas_diferentes, on='ENDS', how='left')

df_result['ZONA_1'] = df_result['ZONAS'].apply(lambda x: x.split(', ')[0] if pd.notna(x) else None)
df_result['ZONA_2'] = df_result['ZONAS'].apply(lambda x: x.split(', ')[1] if pd.notna(x) and len(x.split(', ')) > 1 else None)

df_result = df_result.drop(columns=['ZONAS'])

df_result.to_excel(r'C:\Users\moesios\Desktop\AJUSTES P14\VERIFICAÇÃO.xlsx', index=False)