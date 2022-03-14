from thinkcell import Thinkcell
from Logic import ChartsCreation as chart
from Logic import Analysis_Sheet

analysis=Analysis_Sheet.analysis_sheet()

company_name="CES Total"
bu_name= company_name+ " by " + "business unit "
reference_set_name="B2B business services"


company_usd=analysis.company_usd
bu_usd=analysis.bu_usd
reference_set_usd=analysis.reference_set_usd1

#KPI sBar Charts Creation

KPI1= chart.bar_chart_data_frame("KPI1",company_usd,company_name,bu_usd,bu_name,reference_set_usd,reference_set_name)
KPI2= chart.bar_chart_data_frame("KPI2",company_usd,company_name,bu_usd,bu_name,reference_set_usd,reference_set_name)
KPI3= chart.bar_chart_data_frame("KPI3",company_usd,company_name,bu_usd,bu_name,reference_set_usd,reference_set_name)
KPI4= chart.bar_chart_data_frame("KPI4",company_usd,company_name,bu_usd,bu_name,reference_set_usd,reference_set_name)
KPI5= chart.bar_chart_data_frame("KPI5",company_usd,company_name,bu_usd,bu_name,reference_set_usd,reference_set_name)
KPI6= chart.bar_chart_data_frame("KPI6",company_usd,company_name,bu_usd,bu_name,reference_set_usd,reference_set_name)
KPI7= chart.bar_chart_data_frame("KPI7",company_usd,company_name,bu_usd,bu_name,reference_set_usd,reference_set_name)
KPI9= chart.bar_chart_data_frame("KPI9",company_usd,company_name,bu_usd,bu_name,reference_set_usd,reference_set_name)
KPI10= chart.bar_chart_data_frame("KPI10",company_usd,company_name,bu_usd,bu_name,reference_set_usd,reference_set_name)
KPI11= chart.bar_chart_data_frame("KPI11",company_usd,company_name,bu_usd,bu_name,reference_set_usd,reference_set_name)
KPI12= chart.bar_chart_data_frame("KPI12",company_usd,company_name,bu_usd,bu_name,reference_set_usd,reference_set_name)
KPI17= chart.bar_chart_data_frame("KPI17",company_usd,company_name,bu_usd,bu_name,reference_set_usd,reference_set_name)
KPI18= chart.bar_chart_data_frame("KPI18",company_usd,company_name,bu_usd,bu_name,reference_set_usd,reference_set_name)
KPI19= chart.bar_chart_data_frame("KPI19",company_usd,company_name,bu_usd,bu_name,reference_set_usd,reference_set_name)
KPI25= chart.bar_chart_data_frame("KPI25",company_usd,company_name,bu_usd,bu_name,reference_set_usd,reference_set_name)
KPI26= chart.bar_chart_data_frame("KPI26",company_usd,company_name,bu_usd,bu_name,reference_set_usd,reference_set_name)
KPI32= chart.bar_chart_data_frame("KPI32",company_usd,company_name,bu_usd,bu_name,reference_set_usd,reference_set_name)

KPI43= chart.bar_chart_data_frame("KPI43",company_usd,company_name,bu_usd,bu_name,reference_set_usd,reference_set_name)
KPI45= chart.bar_chart_data_frame("KPI45",company_usd,company_name,bu_usd,bu_name,reference_set_usd,reference_set_name)
KPI51= chart.bar_chart_data_frame("KPI51",company_usd,company_name,bu_usd,bu_name,reference_set_usd,reference_set_name)
KPI57= chart.bar_chart_data_frame("KPI57",company_usd,company_name,bu_usd,bu_name,reference_set_usd,reference_set_name)

#BM Page

reference_rev=analysis.reference_rev_1
reference_sales=analysis.reference_sales_1
region=analysis.bm_region
channel=analysis.bm_channels

# Emp Share

empshare=chart.bar_chart_emp_share(company_usd,bu_usd,["KPI25","KPI27","KPI28","KPI29","KPI30","KPI31","KPI32"])


template_name = "032020 - Sales Navigator Report Template.pptx"
filename = "C:\\Users\\Joel Ramirez\\PycharmProjects\\pptautomation\\Output\\template.ppttc"


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



tc.save_ppttc(filename=filename)