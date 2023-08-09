import pandas as pd
import numpy as np

filename = 'LPE Plus Mistake Notes'

data = []

with open(filename + '.txt') as f:
    [data.append(line.rstrip('\n')) for line in f.readlines()]

newData = []
for i in data:
    row = i.split('->')
    newData.append(row)

df = pd.DataFrame(np.array(newData), columns=['File', 'Error'])

df.to_excel(filename + '.xlsx', 'Sheet1', index=False)