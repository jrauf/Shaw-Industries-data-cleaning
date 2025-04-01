#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import ipywidgets as widgets
from IPython.display import FileLink, display


df = pd.read_excel('6-month Number of Scans per Location - Copy.xlsx', sheet_name="Sheet1")

df['Group'] = df['Row'].astype(str).str[:2]

group_avg = df.groupby('Group', as_index=False)['Frequency'].mean()
group_avg = group_avg.sort_values('Frequency', ascending=False)
fig_group = px.bar(group_avg, x='Group', y='Frequency',
                   title="Average Frequency by Group",
                   labels={'Frequency': 'Avg Frequency'})
fig_group.update_traces(marker_color='lightslategray')
fig_group.show()
groups = group_avg['Group'].tolist()
fig_breakdown = go.Figure()

for group in groups:
    df_group = df[df['Group'] == group].sort_values('Frequency', ascending=False)
    fig_breakdown.add_trace(go.Bar(
        x=df_group['Row'],
        y=df_group['Frequency'],
        name=group,
        visible=(group == groups[0]) 
    ))
buttons = []
for i, group in enumerate(groups):
    visibility = [False] * len(groups)
    visibility[i] = True
    buttons.append(dict(
        label=group,
        method='update',
        args=[{'visible': visibility},
              {'title': f"Breakdown for Group {group}",
               'xaxis': {'title': 'Row'},
               'yaxis': {'title': 'Frequency'}}]
    ))

fig_breakdown.update_layout(
    updatemenus=[dict(
        active=0,
        buttons=buttons,
        x=1.05, 
        y=1,
        xanchor='left',
        yanchor='top'
    )],
    title=f"Breakdown for Group {groups[0]}",
    xaxis_title="Row",
    yaxis_title="Frequency"
)

fig_breakdown.show()


min_freq = df['Frequency'].min()
max_freq = df['Frequency'].max()

def frequency_to_hex(freq, min_freq, max_freq):
    t = (freq - min_freq) / (max_freq - min_freq) if max_freq > min_freq else 0
    r = int(255 * (1 - t))
    g = int(255 * (1 - t))
    b = 255
    return f"#{r:02x}{g:02x}{b:02x}"

df['hex_color'] = df['Frequency'].apply(lambda x: frequency_to_hex(x, min_freq, max_freq))

display(df[['Row', 'Frequency', 'hex_color']])

def download_excel(b):
    file_name = "output.xlsx"
    df.to_excel(file_name, index=False)
    display(FileLink(file_name, result_html_prefix="Download the Excel file: "))


download_button = widgets.Button(description="Download Excel file")
download_button.on_click(download_excel)
display(download_button)

