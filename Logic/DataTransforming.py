import pandas as pd
import numpy as np
import regex as re

# Dataset load
df=pd.read_excel("C:\\Users\\Joel Ramirez\\PycharmProjects\\pptautomation\\Input\\Analysis_Sheet\\Nexeo_SN_Analysis.xlsm",sheet_name="Analysis - Sales",header=5)

#Obtaining columns from Dataset
columnslist=df.columns
columnslist=map(str, columnslist)

# Removing columns
start_with_unnamed = 'Unnamed'
start_with_0 = '0'
cols_without_unnamed= list(filter(lambda x: not x.startswith(start_with_unnamed), columnslist))
cols_without_0= list(filter(lambda x: not x.startswith(start_with_0), cols_without_unnamed))

#Columns to filter
Templatecolumns_df=['Quick', 'Standard', 'Full', 'What is better?', 'UoM', 'Count', 'Average', 'Worst Percentile', 'Bottom Quartile', 'Median', 'Top Quartile', 'Best Percentile', 'Count.1', 'Average.1', 'Worst Percentile.1', 'Bottom Quartile.1', 'Median.1', 'Top Quartile.1', 'Best Percentile.1', 'Count.2', 'Average.2', 'Worst Percentile.2', 'Bottom Quartile.2', 'Median.2', 'Top Quartile.2', 'Best Percentile.2', 'Count.3', 'Average.3', 'Worst Percentile.3', 'Bottom Quartile.3', 'Median.3', 'Top Quartile.3', 'Best Percentile.3', 'UoM.1', 'Average.4', 'Worst Percentile.4', 'Bottom Quartile.4', 'Median.4', 'Top Quartile.4', 'Best Percentile.4', 'Average.5', 'Worst Percentile.5', 'Bottom Quartile.5', 'Median.5', 'Top Quartile.5', 'Best Percentile.5', 'Average.6', 'Worst Percentile.6', 'Bottom Quartile.6', 'Median.6', 'Top Quartile.6', 'Best Percentile.6', 'Average.7', 'Worst Percentile.7', 'Bottom Quartile.7', 'Median.7', 'Top Quartile.7', 'Best Percentile.7', 'Row Num', 'Q1', 'Q2', 'Q3', 'COUNT', 'Q1.1', 'Q2.1', 'Q3.1', 'Q4', 'SUM']

Benchmark_cols=['KPI ID', 'KPI','UoM', 'Count', 'Average', 'Worst Percentile', 'Bottom Quartile', 'Median', 'Top Quartile', 'Best Percentile', 'Count.1', 'Average.1', 'Worst Percentile.1', 'Bottom Quartile.1', 'Median.1', 'Top Quartile.1', 'Best Percentile.1', 'Count.2', 'Average.2', 'Worst Percentile.2', 'Bottom Quartile.2', 'Median.2', 'Top Quartile.2', 'Best Percentile.2', 'Count.3', 'Average.3', 'Worst Percentile.3', 'Bottom Quartile.3', 'Median.3', 'Top Quartile.3', 'Best Percentile.3', 'UoM.1', 'Average.4', 'Worst Percentile.4', 'Bottom Quartile.4', 'Median.4', 'Top Quartile.4', 'Best Percentile.4', 'Average.5', 'Worst Percentile.5', 'Bottom Quartile.5', 'Median.5', 'Top Quartile.5', 'Best Percentile.5', 'Average.6', 'Worst Percentile.6', 'Bottom Quartile.6', 'Median.6', 'Top Quartile.6', 'Best Percentile.6', 'Average.7', 'Worst Percentile.7', 'Bottom Quartile.7', 'Median.7', 'Top Quartile.7', 'Best Percentile.7']


df_without_0=df[cols_without_0]


#Client Data
df_client=df_without_0.drop(columns=Templatecolumns_df)
df_client.dropna(subset = ['KPI ID'], inplace=True)
#Spliting the data in eur and usd
endswith_1=".1"
clientcols=df_client.columns
#Cleaning Client Eur Data
clientcols_eur= list(filter(lambda x: not x.endswith(endswith_1), clientcols))
df_client_eur=df[clientcols_eur]

#Spliting Bu and names
df_client_eur_total_company_col=df_client_eur.columns[0:3]
df_client_eur_total_company=df_client_eur[df_client_eur_total_company_col]
df_client_eur_by_BU= df_client_eur.drop(columns= df_client_eur.columns[2])




#Cleaning df_client_usd cols
df_client_usd=df_client.drop(columns=clientcols_eur)
df_client_usd.insert(0,"KPI ID",df_client_eur["KPI ID"])
df_client_usd.insert(0,"KPI",df_client_eur["KPI"])
old_client_usd_cols=df_client_usd.columns
new_client_usd_cols = [re.sub("\.[^.]*$","", s) for s in old_client_usd_cols]
rename_cols_mapping = dict(zip(old_client_usd_cols, new_client_usd_cols))
df_client_usd.rename(columns=rename_cols_mapping, inplace=True)


#Spliting Bu and names
df_client_usd_total_company_col=df_client_usd.columns[0:3]
df_client_usd_total_company=df_client_usd[df_client_usd_total_company_col]
df_client_usd_by_BU= df_client_usd.drop(columns= df_client_usd.columns[2])



#Benchmark Data

df_benchmark= df[Benchmark_cols]
df_benchmark_cols=df_benchmark.columns
startswith="Count"
df_benchmark_cols_without_count=list(filter(lambda x: not x.startswith(startswith), df_benchmark_cols))
df_benchmark=df_benchmark[df_benchmark_cols_without_count]

#Splitting benchmark reference set in eur
referenceset_eur1_cols=['KPI ID', 'KPI', 'Average', 'Worst Percentile','Bottom Quartile',
                        'Median', 'Top Quartile', 'Best Percentile']
referenceset_eur1=df_benchmark[referenceset_eur1_cols]


referenceset_eur2_cols=['KPI ID', 'KPI','Average.1', 'Worst Percentile.1', 'Bottom Quartile.1',
                        'Median.1','Top Quartile.1', 'Best Percentile.1']
referenceset_eur2=df_benchmark[referenceset_eur2_cols]
rename_cols_eur2_mapping = dict(zip(referenceset_eur2_cols, referenceset_eur1_cols))
referenceset_eur2.rename(columns=rename_cols_eur2_mapping,inplace=True)



referenceset_eur3_cols=['KPI ID', 'KPI','Average.3','Worst Percentile.2', 'Bottom Quartile.2',
                        'Median.2', 'Top Quartile.2','Best Percentile.2']

referenceset_eur3=df_benchmark[referenceset_eur3_cols]

rename_cols_eur3_mapping = dict(zip(referenceset_eur3_cols, referenceset_eur1_cols))
referenceset_eur3.rename(columns=rename_cols_eur3_mapping,inplace=True)

referenceset_eur4_cols=['KPI ID', 'KPI','Average.3', 'Worst Percentile.3','Bottom Quartile.3',
                        'Median.3', 'Top Quartile.3', 'Best Percentile.3']

referenceset_eur4=df_benchmark[referenceset_eur4_cols]
rename_cols_eur4_mapping = dict(zip(referenceset_eur4_cols, referenceset_eur1_cols))
referenceset_eur4.rename(columns=rename_cols_eur4_mapping,inplace=True)

#Splitting benchmark reference set in usd
referenceset_usd1_cols=['KPI ID', 'KPI', 'Average.4', 'Worst Percentile.4', 'Bottom Quartile.4',
                        'Median.4', 'Top Quartile.4', 'Best Percentile.4']
referenceset_usd1= df_benchmark[referenceset_usd1_cols]

referenceset_usd2_cols=['KPI ID', 'KPI','Average.5','Worst Percentile.5', 'Bottom Quartile.5',
                        'Median.5', 'Top Quartile.5','Best Percentile.5']
referenceset_usd2= df_benchmark[referenceset_usd2_cols]

rename_cols_usd2_mapping = dict(zip(referenceset_usd2_cols, referenceset_usd1_cols))
referenceset_usd2.rename(columns=rename_cols_usd2_mapping,inplace=True)

referenceset_usd3_cols=['KPI ID', 'KPI','Average.6', 'Worst Percentile.6','Bottom Quartile.6',
                        'Median.6', 'Top Quartile.6', 'Best Percentile.6']
referenceset_usd3= df_benchmark[referenceset_usd3_cols]
rename_cols_usd3_mapping = dict(zip(referenceset_usd3_cols, referenceset_usd1_cols))
referenceset_usd3.rename(columns=rename_cols_usd3_mapping,inplace=True)

referenceset_usd4_cols=['KPI ID', 'KPI','Average.7', 'Worst Percentile.7', 'Bottom Quartile.7',
                        'Median.7','Top Quartile.7', 'Best Percentile.7']
referenceset_usd4= df_benchmark[referenceset_usd4_cols]
rename_cols_usd4_mapping = dict(zip(referenceset_usd4_cols, referenceset_usd1_cols))
referenceset_usd4.rename(columns=rename_cols_usd4_mapping,inplace=True)