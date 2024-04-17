from openpyxl import load_workbook
from statistics import pvariance, mean
from matplotlib import pyplot as plt
import os


def showplt(col_x, col_y):  # all_files.py
    plt.plot(col_x, col_y, linewidth=2, color='r')
    plt.title('Y-Direction of Microteslas')
    plt.xlabel('Time in Seconds')
    plt.ylabel('Microteslas')
    plt.show()
    # plt.savefig(file+'.png')


def workbook_run(file):  # file loops through each file
    wb = load_workbook(filename='xlsx_files/' + file)
    ws = wb.active
    sheet_ranges = wb['Raw Data']
    data_length = len(sheet_ranges['A'])
    col_x = 0           # initialized to avoid possible reference of these local variables before an assignment
    col_y = 0

    for col_x in ws.iter_cols(min_col=1, min_row=2, max_col=1, max_row=data_length, values_only=True):
        list(col_x)     # print(col_x) # x-axis, time in seconds
    for col_y in ws.iter_cols(min_col=3, min_row=2, max_col=3, max_row=data_length, values_only=True):
        list(col_y)     # y-axis, magnitude of y-direction microteslas μT

    return col_x, col_y, sheet_ranges, data_length, wb, ws


def classify():
    stack_classify = []  # confusion_matrix.py
    stack_opened = []    # display for allfiles.py
    stack_closed = []    # allfiles.py
    stack_nothing = []   # allfiles.py
    tesla_avg1 = 0
    tesla_avg2 = 0

    for file in os.listdir('xlsx_files'):
        col_x, col_y, sheet_ranges, data_length, wb, ws = workbook_run(file)
        workbook_run(file)

        for first_col_y in ws.iter_cols(min_col=3, min_row=2, max_col=3, max_row=101, values_only=True):
            list(first_col_y)               # collects first second of samples from row 2-101, excluding the header
            tesla_avg1 = mean(first_col_y)  # finds mean average of the first 100 samples
            # print(first_col_y)            # print here to see first 100 samples for all files (20)
        for last_col_y in ws.iter_cols(min_col=3, min_row=(data_length - 100), max_col=3, max_row=data_length,
                                       values_only=True):
            list(last_col_y)                # collects last second of samples: 100 before the last row, until the end
            tesla_avg2 = mean(last_col_y)   # finds mean average of the last 100 samples

        nothing = 'Nothing'
        opened = 'Close to open'
        closed = 'Open to close'

        if pvariance(col_y) < 0.5:
            classification = nothing
            # print(classification)
            stack_nothing.append(file)
            stack_classify.append(classification)
        elif pvariance(col_y) >= 0.5:
            if tesla_avg1 > tesla_avg2:
                classification = opened
                # print(classification)
                stack_opened.append(file)
                stack_classify.append(classification)
            else:
                classification = closed
                # print(classification)
                stack_closed.append(file)
                stack_classify.append(classification)
    return stack_classify, stack_nothing, stack_opened, stack_closed


stack_classify2, stack_nothing2, stack_opened2, stack_closed2 = classify()


def s_classify():   # confusion_matrix.py
    return stack_classify2


def s_nothing():    # all_files.py
    return stack_nothing2


def s_opened():     # all_files.py
    return stack_opened2


def s_closed():     # all_files.py
    return stack_closed2


def max_var(data, window, step):
    num_window = (len(data) - window)                # iterate until window size is too small
    variances = []                                   # initialize

    # (start, end, iteration) 100 for one-second window, 50 for 0.5 second
    for window_start in range(0, num_window, step):  # increases window's start index by step size 50
        window_end = window_start + window
        window_data = data[window_start:window_end]

        if len(window_data) > 1:                     # does not check (last) window if there's less than 2 data values
            window_variance = pvariance(window_data)
            variances.append(window_variance)

    # print(window_start) to check step sizes
    max_variance = max(variances)
    time_s = 0.5 * variances.index(max_variance)     # only returns the first occurrence of the max value
    sample_start = 50 * variances.index(max_variance)
    # below is a list comprehension, checks for duplicate indices
    max_index = [window_start for window_start, var in enumerate(variances) if var == max_variance]
    print("Time in seconds: ", time_s)               # time_s is (*50 for step size, /100 for sample size), = *0.5

    return max_variance, max_index, sample_start
