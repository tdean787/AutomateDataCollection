import pandas

donortestinfo = pandas.read_excel('testsheet.xlsx')
donors = donortestinfo.to_dict('records')

for value in donors:
    print(value)

for person in donors:
    for k,v in person.items():
        for a in v:
            print(a)

for person in donors:
    for k,v in person.items():
        print(k)