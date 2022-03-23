from thinkcell import Thinkcell
from Logic import ChartsCreation as chart
from Logic import Analysis_Sheet
import os

analysis=Analysis_Sheet.analysis_sheet()

pptfilename = os.listdir("Output")[0]

company_name=analysis.company_name
bu_name= company_name+ " by " + analysis.bu_naming
reference_set_name=analysis.reference_set_naming
currency=analysis.currency

#BM Page

reference_rev=analysis.reference_rev_1
reference_sales=analysis.reference_sales_1
region=analysis.bm_region
channel=analysis.bm_channels


if currency == "$":
    company_data=analysis.company_usd
    bu_data=analysis.bu_usd
    reference_set_data=analysis.reference_set_usd1
elif currency == "€" :
    company_data = analysis.company_eur
    bu_data = analysis.bu_eur
    reference_set_data = analysis.reference_set_eur1
else:
    print("Please select a valid option in the Selection excel file")


#KPI sBar Charts Creation

KPI1= chart.bar_chart_data_frame("KPI1", company_data, company_name, bu_data, bu_name, reference_set_data, reference_set_name)
KPI2= chart.bar_chart_data_frame("KPI2", company_data, company_name, bu_data, bu_name, reference_set_data, reference_set_name)
KPI3= chart.bar_chart_data_frame("KPI3", company_data, company_name, bu_data, bu_name, reference_set_data, reference_set_name)
KPI4= chart.bar_chart_data_frame("KPI4", company_data, company_name, bu_data, bu_name, reference_set_data, reference_set_name)
KPI5= chart.bar_chart_data_frame("KPI5", company_data, company_name, bu_data, bu_name, reference_set_data, reference_set_name)
KPI6= chart.bar_chart_data_frame("KPI6", company_data, company_name, bu_data, bu_name, reference_set_data, reference_set_name)
KPI7= chart.bar_chart_data_frame("KPI7", company_data, company_name, bu_data, bu_name, reference_set_data, reference_set_name)
KPI9= chart.bar_chart_data_frame("KPI9", company_data, company_name, bu_data, bu_name, reference_set_data, reference_set_name)
KPI10= chart.bar_chart_data_frame("KPI10", company_data, company_name, bu_data, bu_name, reference_set_data, reference_set_name)
KPI11= chart.bar_chart_data_frame("KPI11", company_data, company_name, bu_data, bu_name, reference_set_data, reference_set_name)
KPI12= chart.bar_chart_data_frame("KPI12", company_data, company_name, bu_data, bu_name, reference_set_data, reference_set_name)
KPI17= chart.bar_chart_data_frame("KPI17", company_data, company_name, bu_data, bu_name, reference_set_data, reference_set_name)
KPI18= chart.bar_chart_data_frame("KPI18", company_data, company_name, bu_data, bu_name, reference_set_data, reference_set_name)
KPI19= chart.bar_chart_data_frame("KPI19", company_data, company_name, bu_data, bu_name, reference_set_data, reference_set_name)
KPI25= chart.bar_chart_data_frame("KPI25", company_data, company_name, bu_data, bu_name, reference_set_data, reference_set_name)
KPI26= chart.bar_chart_data_frame("KPI26", company_data, company_name, bu_data, bu_name, reference_set_data, reference_set_name)
KPI32= chart.bar_chart_data_frame("KPI32", company_data, company_name, bu_data, bu_name, reference_set_data, reference_set_name)

KPI43= chart.bar_chart_data_frame("KPI43", company_data, company_name, bu_data, bu_name, reference_set_data, reference_set_name)
KPI45= chart.bar_chart_data_frame("KPI45", company_data, company_name, bu_data, bu_name, reference_set_data, reference_set_name)
KPI51= chart.bar_chart_data_frame("KPI51", company_data, company_name, bu_data, bu_name, reference_set_data, reference_set_name)
KPI57= chart.bar_chart_data_frame("KPI57", company_data, company_name, bu_data, bu_name, reference_set_data, reference_set_name)


# Emp Share

empshare=chart.bar_chart_emp_share(company_data,bu_data,["KPI25", "KPI27", "KPI28", "KPI29", "KPI30", "KPI31", "KPI32"],reference_set_data)
empshare2=chart.bar_chart_emp_share(company_data,bu_data,["KPI25", "KPI26", "KPI32"],reference_set_data)



template_name = pptfilename
filename = "Output\\template.ppttc"


tc = Thinkcell() # create thinkcell object
tc.add_template(template_name)

# add your template
tc.add_chart_from_dataframe(
    template_name=template_name,
    chart_name="KPI1",
    dataframe=KPI1,
) # add your dataframe

tc.add_chart_from_dataframe(
    template_name=template_name,
    chart_name="KPI2",
    dataframe=KPI2,
) # add your dataframe

tc.add_chart_from_dataframe(
    template_name=template_name,
    chart_name="KPI3",
    dataframe=KPI3,
) # add your dataframe

tc.add_chart_from_dataframe(
    template_name=template_name,
    chart_name="KPI4",
    dataframe=KPI4,
) # add your dataframe

tc.add_chart_from_dataframe(
    template_name=template_name,
    chart_name="KPI5",
    dataframe=KPI5,
) # add your dataframe

tc.add_chart_from_dataframe(
    template_name=template_name,
    chart_name="KPI6",
    dataframe=KPI6,
) # add your dataframe

# add your template
tc.add_chart_from_dataframe(
    template_name=template_name,
    chart_name="KPI7",
    dataframe=KPI7,
) # add your dataframe

tc.add_chart_from_dataframe(
    template_name=template_name,
    chart_name="KPI9",
    dataframe=KPI9,
) # add your dataframe

tc.add_chart_from_dataframe(
    template_name=template_name,
    chart_name="KPI10",
    dataframe=KPI10,
) # add your dataframe

tc.add_chart_from_dataframe(
    template_name=template_name,
    chart_name="KPI11",
    dataframe=KPI11,
) # add your dataframe

tc.add_chart_from_dataframe(
    template_name=template_name,
    chart_name="KPI12",
    dataframe=KPI12,
) # add your dataframe

tc.add_chart_from_dataframe(
    template_name=template_name,
    chart_name="KPI17",
    dataframe=KPI17,
) # add your dataframe

# add your template
tc.add_chart_from_dataframe(
    template_name=template_name,
    chart_name="KPI18",
    dataframe=KPI18,
) # add your dataframe

tc.add_chart_from_dataframe(
    template_name=template_name,
    chart_name="KPI19",
    dataframe=KPI19,
) # add your dataframe

tc.add_chart_from_dataframe(
    template_name=template_name,
    chart_name="KPI25",
    dataframe=KPI25,
) # add your dataframe

tc.add_chart_from_dataframe(
    template_name=template_name,
    chart_name="KPI26",
    dataframe=KPI26,
) # add your dataframe

tc.add_chart_from_dataframe(
    template_name=template_name,
    chart_name="KPI43",
    dataframe=KPI43,
) # add your dataframe

tc.add_chart_from_dataframe(
    template_name=template_name,
    chart_name="KPI45",
    dataframe=KPI45,
) # add your dataframe

tc.add_chart_from_dataframe(
    template_name=template_name,
    chart_name="KPI51",
    dataframe=KPI51,
) # add your dataframe

tc.add_chart_from_dataframe(
    template_name=template_name,
    chart_name="KPI57",
    dataframe=KPI57,
) # add your dataframe


tc.add_chart_from_dataframe(
    template_name=template_name,
    chart_name="BMREGION",
    dataframe=region,
) # add your dataframe

tc.add_chart_from_dataframe(
    template_name=template_name,
    chart_name="BMCHANNEL",
    dataframe=channel,
) # add your dataframe

tc.add_chart_from_dataframe(
    template_name=template_name,
    chart_name="BMREV",
    dataframe=reference_rev,
) # add your dataframe

tc.add_chart_from_dataframe(
    template_name=template_name,
    chart_name="BMSALES",
    dataframe=reference_sales,
) # add your dataframe

tc.add_chart_from_dataframe(
    template_name=template_name,
    chart_name="EMPSHARE",
    dataframe=empshare,
) # add your dataframe

tc.add_chart_from_dataframe(
    template_name=template_name,
    chart_name="EMPSHARE2",
    dataframe=empshare2,
) # add your dataframe


tc.save_ppttc(filename=filename)

print("The ppttc file has been succesfully created")
print("Please go the output folder and double click the pptc file")

