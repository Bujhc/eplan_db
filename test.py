

a_list = ['ABB']
b_list = [
    'ABB', 'ADS', 'AXIS', 'Beckhoff', 'Belimo', 'Carel', 'Danfoss prices',\
    'Delta Electronics', 'desktop.ini', 'EBM', 'EKF', 'Evrotehna', 'FINDER', 'Gira',\
    'GSS', 'Honeywell', 'iEK', 'Johnson Controls', 'Jung', 'Legrand', 'LG', 'LSIS',\
    'ONI', 'Ostec', 'Phoenix Contact', 'RITALL', 'sabiana', 'Sauter', 'Schneider Electric',\
    'Segnetics', 'Siemens', 'TDM', 'Thermokon', 'Vadiart', 'Wago', 'York', 'ДКС', 'Инкотекс',\
    'кассетные фанкойлы', 'КЭА3', 'ОВЕН', 'ФКУ ИК-1', 'Электрокабель'
    ]

for i in b_list:
    #print(i)
    if a_list[0] == i:
        print(i)