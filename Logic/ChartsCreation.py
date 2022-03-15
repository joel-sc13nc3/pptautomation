import numpy as np
import pandas as pd

def bar_chart_data_frame(kpi_id, company_data, company_name, bu_data, bu_name, reference_set_data, reference_set_name, all_metrics=False):
    company_data=company_data[company_data["KPI ID"] == kpi_id]
    company_data= company_data.iloc[:, [2]]
    company_data["item"]= company_name
    company_data.insert(0, 'item', company_data.pop("item"))

    bu_data=bu_data[bu_data["KPI ID"] == kpi_id]
    bu_data = bu_data.iloc[:, 2:]
    bu_data["item"] = bu_name
    bu_data.insert(0, 'item', bu_data.pop("item"))

    reference_set_data=reference_set_data[reference_set_data["KPI ID"] == kpi_id]
    if all_metrics ==False:
        reference_set_data=reference_set_data[['Bottom Quartile', 'Median', 'Top Quartile']]
    else: reference_set_data=reference_set_data[['Average', 'Worst Percentile', 'Bottom Quartile',
       'Median', 'Top Quartile', 'Best Percentile']]

    reference_set_data["item"] = reference_set_name
    reference_set_data.insert(0, 'item', reference_set_data.pop("item"))
    company_bu=company_data.append(bu_data)
    company_bu_reference =company_bu.append(reference_set_data)
    company_bu_reference=company_bu_reference.replace(np.nan,"N/A")

    return company_bu_reference



def referencetranformation(df,df1,rowname1,rowname2):
    df=pd.DataFrame(df)
    df=df.transpose()
    columnslist=df.columns
    newcols=["Bottom Quartile","Median","Top Quartile"]
    colsdict=dict(zip(columnslist,newcols))
    df=df.rename(columns=colsdict)
    df=df.reset_index()
    df = pd.concat([df1, df])
    df=df.replace(np.nan,"N/A")
    df=df.drop(columns="index")
    df["item"]=[rowname1,rowname2]
    df.insert(0, "item", df.pop("item"))
    return df






def bar_chart_emp_share(df,bu,kpi_id):

    df=df[df["KPI ID"].isin(kpi_id)]
    bu=bu[bu["KPI ID"].isin(kpi_id)]
    data=pd.merge(df,bu,on="KPI")
    data=data.drop(columns=["KPI ID_x","KPI ID_y"])
    data=data.replace(np.nan,0)
    datacols=np.array(data.columns[1:])
    for x in datacols:
        sum=data[x].sum()
        data[x]=data[x]/sum
        data[x] = data[x] *100
    data['Categories'] =["Account Owner","Customer Facing","Customer Facing","Non Customer Facing","Non Customer Facing","Non Customer Facing","Sales Management"]
    data.insert(0, "Categories", data.pop("Categories"))
    data=data.groupby(by=["Categories"]).sum()
    data["Top Performers"]=[0.57,0.20,0.09,0.14]
    return data