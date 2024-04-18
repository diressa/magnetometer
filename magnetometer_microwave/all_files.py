from matplotlib import pyplot as plt
import os
import plotly.graph_objects as go
from classifier import s_nothing, s_opened, s_closed, workbook_run


def showplt(x, y):
    plt.plot(x, y, linewidth=2, color='r')
    plt.title('Y-Direction of Microteslas')
    plt.xlabel('Time in Seconds')
    plt.ylabel('Microteslas')
    plt.show()
    # plt.savefig(file+'.png')


for file in os.listdir('xlsx_files'):
    col_x, col_y, sheet_ranges, data_length, wb, ws = workbook_run(file)
    workbook_run(file)
    showplt(col_x, col_y)
    print(file)

fig = go.Figure(data=[go.Table(header=dict(values=['Opened', 'Closed', 'Nothing']),
                cells=dict(values=[list(s_opened()), list(s_closed()), list(s_nothing())]))])

fig.show()
