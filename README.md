# dio-dashboard-vendas-excel
Dashboard de Vendas em Excel



# Dashboard de Vendas em Excel

Descrição
Este projeto gera um arquivo Excel `dashboard_sales.xlsx` com um dashboard de vendas completo, dados sintéticos e visualizações para análise de desempenho.

Arquivos no repositório
- `dashboard_sales.xlsx` — arquivo Excel gerado com RawData, PivotData, Dashboard, Charts e Instructions.
- `gerar_dashboard_com_graficos.py` — script Python que gera o arquivo Excel.
- `README.md` — este arquivo.

Dependências
- Python 3.10 ou superior
- Bibliotecas Python: pandas, numpy, xlsxwriter, openpyxl
Instalação:
```bash
pip install pandas numpy xlsxwriter openpyxl

# gerar_dashboard.py
# Requer: Python 3.10+, pip install pandas numpy xlsxwriter openpyxl
import pandas as pd, numpy as np, random
from datetime import datetime
import calendar

np.random.seed(42)
start = pd.to_datetime('2023-01-01')
end = pd.to_datetime('2024-12-31')
dates = pd.date_range(start, end, freq='D')

regions = ['Norte','Sul','Leste','Oeste']
countries = ['Brasil','Argentina','Chile','Peru','Colombia','Uruguai','Paraguai','Bolívia']
categories = ['Eletrônicos','Móveis','Acessórios','Software','Serviços','Eletrodomésticos']
products = [f'Produto {i:02d}' for i in range(1,51)]
sales_reps = [f'Vendedor {i:02d}' for i in range(1,31)]

rows=[]
order_id=100000
for _ in range(2400):
    d = np.random.choice(dates)
    year=d.year; month=d.month
    region=np.random.choice(regions)
    country=np.random.choice(countries)
    state=f'State_{random.randint(1,20)}'
    city=f'City_{random.randint(1,200)}'
    rep=np.random.choice(sales_reps)
    cat=np.random.choice(categories, p=[0.25,0.15,0.2,0.1,0.15,0.15])
    product=np.random.choice(products)
    base_price = {'Eletrônicos':1200,'Móveis':800,'Acessórios':80,'Software':400,'Serviços':250,'Eletrodomésticos':600}[cat]
    unit_price = round(max(10, np.random.normal(base_price, base_price*0.15)),2)
    units = max(1,int(np.random.poisson(3)))
    # sazonalidade
    if month in [11,12]:
        units = int(units * 1.6)
    if month==5:
        units = int(units * 1.3)
    discount = float(np.random.choice([0,0.05,0.1,0.15,0.2], p=[0.6,0.15,0.15,0.07,0.03]))
    revenue = round(units * unit_price,2)
    net = round(revenue * (1-discount),2)
    rows.append([order_id,d,year,month,region,country,state,city,rep,cat,product,units,unit_price,revenue,discount,net])
    order_id+=1

df = pd.DataFrame(rows, columns=['OrderID','OrderDate','Year','Month','Region','Country','State','City','SalesRep','ProductCategory','Product','UnitsSold','UnitPrice','Revenue','Discount','NetRevenue'])
df['MonthName'] = df['OrderDate'].dt.month.apply(lambda x: calendar.month_abbr[x])
# PivotData
monthly = df.groupby(['Year','Month']).agg({'NetRevenue':'sum','UnitsSold':'sum'}).reset_index().sort_values(['Year','Month'])
monthly_pivot = monthly.pivot(index='Month', columns='Year', values='NetRevenue').reindex(range(1,13)).fillna(0).reset_index()
region_q = df.copy(); region_q['Quarter'] = df['OrderDate'].dt.to_period('Q').astype(str)
region_q = region_q.groupby(['Quarter','Region']).agg({'NetRevenue':'sum'}).reset_index()
top_products = df.groupby('Product').agg({'NetRevenue':'sum'}).reset_index().sort_values('NetRevenue',ascending=False).head(10)
top_reps = df.groupby('SalesRep').agg({'NetRevenue':'sum'}).reset_index().sort_values('NetRevenue',ascending=False).head(10)
category_share = df.groupby('ProductCategory').agg({'NetRevenue':'sum'}).reset_index()

# Escrever Excel com xlsxwriter
with pd.ExcelWriter('dashboard_sales.xlsx', engine='xlsxwriter', datetime_format='dd/mm/yyyy') as writer:
    df.to_excel(writer, sheet_name='RawData', index=False)
    monthly.to_excel(writer, sheet_name='PivotData', index=False, startrow=0)
    monthly_pivot.to_excel(writer, sheet_name='PivotData', index=False, startrow=monthly.shape[0]+4)
    region_q.to_excel(writer, sheet_name='PivotData', index=False, startrow=monthly.shape[0]+monthly_pivot.shape[0]+8)
    top_products.to_excel(writer, sheet_name='PivotData', index=False, startrow=monthly.shape[0]+monthly_pivot.shape[0]+region_q.shape[0]+12)
    top_reps.to_excel(writer, sheet_name='PivotData', index=False, startrow=monthly.shape[0]+monthly_pivot.shape[0]+region_q.shape[0]+top_products.shape[0]+16)
    category_share.to_excel(writer, sheet_name='PivotData', index=False, startrow=monthly.shape[0]+monthly_pivot.shape[0]+region_q.shape[0]+top_products.shape[0]+top_reps.shape[0]+20)

    workbook = writer.book
    fmt_money = workbook.add_format({'num_format':'R$ #,##0.00','bold':True})
    fmt_title = workbook.add_format({'bold':True,'font_size':14})
    # Criar sheet Dashboard
    dash = workbook.add_worksheet('Dashboard')
    dash.write('A1','Dashboard de Vendas', fmt_title)
    # Inserir dropdowns (validação) e KPIs com fórmulas que referenciam RawData (tbl)
    # Criar tabela Excel para RawData via openpyxl não é trivial com xlsxwriter; em vez disso, instruções no README explicam como transformar em tabela no Excel.
    dash.write('A3','Filtro Região:')
    dash.data_validation('B3', {'validate':'list','source':regions})
    dash.write('A4','Filtro Categoria:')
    dash.data_validation('B4', {'validate':'list','source':categories})
    # KPIs (células com fórmulas que usam SUMIFS sobre a sheet RawData)
    dash.write('A6','Total Net Revenue (R$):')
    dash.write_formula('B6',"=SUM(RawData!P:P)", fmt_money)
    dash.write('A7','Total Units Sold:')
    dash.write_formula('B7',"=SUM(RawData!L:L)")
    dash.write('A8','Average Order Value (R$):')
    dash.write_formula('B8',"=IF(COUNTA(RawData!A:A)=0,0,SUM(RawData!P:P)/COUNTA(RawData!A:A))", fmt_money)
    dash.write('A9','Average Discount:')
    dash.write_formula('B9',"=AVERAGE(RawData!O:O)")
    # Observação: gráficos e formatação mais avançada podem ser adicionados manualmente no Excel; o arquivo contém os dados e tabelas PivotData.
    dash.write('A12','Observações:')
    dash.write('A13','1) Use Atualizar Tudo no Excel para recalcular pivôs e gráficos se substituir RawData.')
    dash.write('A14','2) Para botões VBA, siga instruções na sheet Instructions.')
    # Instructions sheet
    instr = workbook.add_worksheet('Instructions')
    instr.write('A1','Instruções de uso')
    instr.write('A3','- Abra dashboard_sales.xlsx no Excel.')
    instr.write('A4','- RawData contém os dados brutos. Substitua ou adicione linhas mantendo as colunas.')
    instr.write('A5','- Vá em Dados > Atualizar Tudo para atualizar pivôs e gráficos.')
    instr.write('A6','- Para criar botões VBA: Developer > Insert > Button e atribua macro que execute ActiveWorkbook.RefreshAll().')
    writer.save()
# Gerar README.md
readme = """
# Dashboard de Vendas (Excel)

Descrição curta: Dashboard de vendas gerado sinteticamente para análise de desempenho, KPIs e visualizações.

Arquivos:
- dashboard_sales.xlsx : arquivo Excel com RawData, PivotData, Dashboard, Instructions.
- gerar_dashboard.py : script Python que gera o arquivo (se desejar regenerar).

Dependências:
Python 3.10+, pip install pandas numpy xlsxwriter openpyxl

Como reproduzir:
1. Instale dependências: pip install pandas numpy xlsxwriter openpyxl
2. Execute: python gerar_dashboard.py
3. Abra dashboard_sales.xlsx no Excel.

Planilhas:
- RawData: dados brutos (OrderID, OrderDate, Year, Month, Region, Country, State, City, SalesRep, ProductCategory, Product, UnitsSold, UnitPrice, Revenue, Discount, NetRevenue).
- PivotData: tabelas agregadas pré-calculadas (mensal, por região por trimestre, top produtos, top vendedores, participação por categoria).
- Dashboard: KPIs e filtros (dropdowns) com fórmulas que referenciam RawData.
- Instructions: instruções rápidas e como adicionar botões VBA.

Limitações e melhorias:
- O script cria dados e fórmulas; gráficos mais avançados e botões VBA podem ser adicionados manualmente no Excel. Instruções estão na sheet Instructions.
- Melhorias: adicionar slicers nativos, macros para regenerar dados, dashboards interativos com Power Query/Power Pivot.

"""
with open('README.md','w', encoding='utf-8') as f:
    f.write(readme)
print('Arquivos gerados: dashboard_sales.xlsx e README.md')
