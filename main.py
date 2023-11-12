from docxtpl import DocxTemplate
import pandas

data = pandas.read_csv('book1.csv')
name = data['Name'].to_list()
dept = data['Dept'].to_list()

doc = DocxTemplate('kepada.docx')

tempat=[]
for loop in dept:
    if loop == 'AT':
        tempat.append('Room A')
    elif loop == 'Ad':
        tempat.append('Room B')
    elif loop == 'AC':
        tempat.append('Room C')
    elif loop == 'MK' or loop == 'MF':
        tempat.append('Room d')
    else:
        tempat.append('Room E')


for l in name:
    num = name.index(l)
    context = {'nama':f'{l}','lokasi':tempat[num],'jabatan':dept[num]}
    doc.render(context)
    doc.save(f'C:\Documents\pythonProject\contoh output\kepada {l}.docx')





