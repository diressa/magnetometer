import os
from classifier import max_var, workbook_run
from matplotlib import pyplot as plt


def showplt_var(x, y):  # time_tracker.py
    print(file)
    max_variance, max_index, sample_start = max_var(col_y, 100, 50)  # max_var(data, window length, step size)
    print(max_variance, max_index)

    markers = sample_start  # return just row (sample #, *50)

    # markers in the x and y-axis lists (start, end) of window subarray.
    markers_x = (sheet_ranges['A' + str(markers)].value, sheet_ranges['A' + str(markers + 100)].value)
    markers_y = (sheet_ranges['C' + str(markers)].value, sheet_ranges['C' + str(markers + 100)].value)
    print(markers, markers_x, markers_y)

    plt.plot(x, y, linewidth=2, color='r')
    plt.plot(markers_x, markers_y, linewidth=2, color='b', marker='x')
    plt.title('Y-Direction of Microteslas')
    plt.xlabel('Time in Seconds')
    plt.ylabel('Microteslas')
    plt.show()


for file in os.listdir('xlsx_files'):
    col_x, col_y, sheet_ranges, data_length, wb, ws = workbook_run(file)
    workbook_run(file)
    showplt_var(col_x, col_y)

