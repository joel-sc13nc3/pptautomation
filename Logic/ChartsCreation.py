import numpy as np



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