from openpyxl import load_workbook
from statistics import mean
from matplotlib import pyplot as plt

sample = random.choice(os.listdir('random_xlsx_files')) #random file sample
#print(sample) #remove as comment to see which file is used
wb = load_workbook(filename='random_xlsx_files/'+sample)
ws = wb.active

sheet_ranges = wb['Raw Data']
data_length = len(sheet_ranges['A'])


for col_x in ws.iter_cols(min_col=1, min_row=2, max_col=1, max_row=data_length, values_only=True):
    list(col_x)#print(col_x) # x-axis, time in seconds

for col_y in ws.iter_cols(min_col=3, min_row=2, max_col=3, max_row=data_length, values_only=True):
    list(col_y) # y-axis, magnitude of y-direction microteslas

classification = ''
#print(pvariance(col_y))

if pvariance(col_y)<0.5:
    classification = 'Nothing'
    print(classification)  # later can group files or label?
elif pvariance(col_y)>=0.5:
    if sheet_ranges['C2'].value > sheet_ranges['C' + str(data_length)].value:
        classification = 'Close to open'
        print(classification)
    else:
        classification = 'Open to close'
        print(classification)


plt.plot(col_x, col_y, linewidth=2, color='r')
plt.title('Y-Direction of Microteslas')
plt.xlabel('Time in Seconds')
plt.ylabel('Microteslas')
plt.show()

