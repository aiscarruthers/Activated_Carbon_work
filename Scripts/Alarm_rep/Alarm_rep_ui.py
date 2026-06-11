import panel as pn
import Alarm_report_ui as ar
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import win32com.client
import json
n_alarms = 0
longest_tags = []
long = 0
av_long = 0
T_long = 0

def process_file(file_contents):

    df = ar.cleanFile(file_contents)
    df.index = df.DTime
    # print(df)
    topten = ar.topN(df,10)
    # print(df)
    Dfs = []
    
    for t in topten["Tag"]:
        inter = df.query('Status !="OK"& Tag == @t').resample('5T').count()
        series = inter.DTime
        Dfs.append(series)
    
    toptenDf = pd.concat(Dfs, axis=1)
    toptenDf.columns = topten["Tag"]
    toptenDf.copy().iloc[:,0:6]
    graph = px.line(toptenDf)
    af = ar.cleanFile(file_contents)
    # print(topten)
    return topten, graph, af


def file_input_callback(event):
    file_contents = file_input.value
    # print(type(file_contents))
    dfs = []
    for f in file_contents:
        file_lines = f.decode('ISO-8859-1').splitlines()
        init_frame = ar.cleanFile(file_lines)
        dfs.append(init_frame)
    global n_alarms
    global longest_tags
    global long
    global av_long
    global T_long
    final_df = pd.concat(dfs,axis=0,ignore_index=True)
    topten = ar.topN(final_df.copy(),20)
    av_long, long, T_long = ar.longestStats(final_df.copy())
    n_alarms = topten.n_alarms
    topten = topten.reset_index(drop=True)
    topten = topten.rename(columns={"DTime":"Number of Alarms"})
    topten.to_html("tag_table.html", index=False)
    plot_df = final_df.copy()
    plot_df.index = plot_df.DTime
    tag_plots = []
    
    for t in topten["Tag"]: 
        inter = plot_df.query('Status !="OK"& Tag == @t').resample('5min').count()
        series = inter.DTime
        tag_plots.append(series)

    topTenDf = pd.concat(tag_plots, axis=1)
    topTenDf.columns = topten["Tag"]
    fig = px.line(topTenDf)
    fig.write_html('interactive_alarm_trends.html')

    dataframe_widget.value = topten
    tag_plot.object = fig
    tab_2.value = final_df

def format_desc(tag_list, dictionary):
    tag_desc = []
    rm = []
    for i , j in enumerate(tag_list):
        try:
            tag_desc.append(dictionary[j])
        except KeyError:
            print(i, " " , j)
            rm.append(i)
            continue
    
    for i in sorted(rm, reverse=True):
        tag_list.pop(i)

    tag_df = pd.DataFrame({"Tag":tag_list,"Alarm Description":tag_desc})
    return tag_df.to_html()

def create_msg_file(event):
    # Create a sample .msg file
    msg_file_path = 'sample.msg'
    with open(msg_file_path, 'w') as f:
        f.write('This is a sample message.')
    
    html_table_path = "tag_table.html"
    with open(html_table_path, 'r') as html_table:
        tag_table = html_table.read()

    global long
    global av_long
    global T_long

    with open(file="tag_dict.json", mode="r" )as f:

        tag_dict = json.load(f)
 
    long_html = format_desc(long, tag_dict)
    av_long_html = format_desc(av_long, tag_dict)
    T_long_html = format_desc(T_long, tag_dict)

    F_longest = os.path.abspath("fiveLongest.png")
    Av_longest = os.path.abspath("longestAverage.png")
    Total = os.path.abspath("Total.png")
    style = os.path.abspath("style.css")
    print(Av_longest)
    longest_df = pd.DataFrame
    # Attach the Plotly HTML file to the .msg file
    outlook = win32com.client.Dispatch('Outlook.Application')
    msg = outlook.CreateItem(0)
    msg.To = 'Alan.Thompson@norit.com; David.Pepper@norit.com '
    msg.CC = 'Mairi.NicolWood@norit.com; Colin.Timmins@norit.com'
    msg.Subject = 'Alarm Management Report'
 
    attachment = os.path.abspath('interactive_alarm_trends.html')
    msg.Attachments.Add(attachment)
    msg.Attachments.Add(F_longest)
    msg.Attachments.Add(Av_longest)
    msg.Attachments.Add(Total)
    css_content = ""
    with open(style, 'r') as file:
        css_content= file.read()
    

    msg.HTMLBody = f'''
    <html>
    <head>
    <style type="text/css">
    {css_content}
    </style
    <head>
    <body>
    <p>Hi All,</p>
    <p>During the Period analysed there were {n_alarms} alarms total.</p>
    <p>10 Longest single times spent in Alarm:</p>
    <div><img src="cid:fiveLongest.png" style="border: 5px solid #ffc;"></div>
    <div>
    {long_html}
    </div>
    <p>10 Longest Alarms on average:</p>
    <div><img src="cid:longestAverage.png"></div> 
    <div>
    {av_long_html}
    </div>
    <p>10 Longest Total time spent in alarm:</p>
    <div><img src=cid:Total.png"></div>
    <div>
    {T_long_html}
    </div>
    <p><strong>Table of 20 Most Common alarms:</strong></p>
    <div>
    {tag_table}
    </div>
    <p>Kind regards,</p>
    </body>
    </html>
    '''


    # Save the .msg file
    
    msg.display()  

    # Open the .msg file with Outlook
    # os.system(f'start outlook.exe /c "IPM.Note /a {msg_file_path}/a interactive_alarm_trends.html"')

def pareto_gen(df): 
    fig = go.Figure([go.Bar(x=df["Tag"],y=df["DTime"], yaxis="y1",name="count")])


# Create FileInput widget
file_input = pn.widgets.FileInput(accept=".ALM", multiple=True)

email = pn.widgets.Button(name='Create Email', button_type='primary')
email.on_click(create_msg_file)

# Create Button to trigger file processing
process_button = pn.widgets.Button(name="Process File", button_type="primary")
process_button.on_click(file_input_callback)

place_holder = pd.DataFrame()
dataframe_widget = pn.widgets.DataFrame(widths={"index":10,"Tag":60,"DTime":10,"Desc":150})
tag_plot = pn.pane.Plotly()
pareto = pn.pane.Plotly()

template = pn.template.GoldenTemplate(
    title = "Alarm viewer",
    sidebar = [file_input,process_button,email],
    theme = "default"
)

tab_1 = pn.Column(tag_plot, dataframe_widget)
tab_2 = pn.widgets.DataFrame(height=900,width=900,widths={"index":10,"DTime":40,"Tag":40,"Status":30,"Value":10,"Units":10,"Desc":100})
template.main.append(tab_1)
template.main.append(tab_2)

# Create a Panel layout
layout = pn.Column(
    "## File Input Example",
    file_input,
    process_button,
    dataframe_widget
)

# Show the Panel app
template.servable()
