import pandas as pd
from sklearn.model_selection import train_test_split
str='project.xlsx'
read=pd.read_excel(str)
print(read)
X=read.loc[0:,[	'age',	'gender',	'weight',	'lipcolor',	'FEV1',	'smoking intensity',	'temperature']]
print(X)
Y=read.loc[0:,['label']]
#print(Y)
X_train,X_test,Y_train,Y_test=train_test_split(X,Y,test_size=0.4,random_state=42)
#print(len(X_train))
#print(len(Y_train))
#print(len(X_test))
#print(len(Y_test))


from sklearn.tree import DecisionTreeClassifier   #random forests
from sklearn.metrics import accuracy_score
model=DecisionTreeClassifier()

X_test = X_test.fillna(X_train.mean())
Y_test = Y_test.fillna(Y_train.mean())
model.fit(X_train,Y_train.values.ravel())
prediction=model.predict(X_test)
print("Accuracy of the Algorithm",accuracy_score(Y_test,prediction)*100)
print('Tree Prediction',model.predict([[40,	1,	60,	1,	0.3,	0.7,	102	]]))

predi=model.predict([[40,	1,	60,	1,	0.3,	0.7,	102	]])

import openpyxl
excel_document=openpyxl.load_workbook('treatments copd.xlsx')

sheet=excel_document.get_sheet_by_name('Sheet1')

if predi==['mild']:
    multiple_cells=sheet['B2':'E2']
    for row in multiple_cells:
        for cell in row:
            print('Treatement for Mild Case is',cell.value)
elif predi==['moderate']:
    multiple_cells = sheet['B3':'E3']
    for row in multiple_cells:
        for cell in row:
            print('Treatement for moderate Case is', cell.value)
elif predi==['severe']:
    multiple_cells = sheet['B4':'E4']
    for row in multiple_cells:
        for cell in row:
            print('Treatement for Severe Case is', cell.value)
else:
    print("couldn't able to find the stage")