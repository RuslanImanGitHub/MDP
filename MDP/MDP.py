# Загрузка библиотек
import numpy as np
import pandas as pd
import sys
import win32com.client
rastr = win32com.client.Dispatch('Astra.Rastr')
shablon_path = 'C:\Program Files (x86)\RastrWin3\RastrWin3\SHABLON'
regime_path = shablon_path + '\режим.rg2'
rastr.Load(1, 'regimeMDP.rg2', regime_path)
rastr.rgm('p')


# Загрузка траектории утяжеления
# Подготовка данных к загрузке в Растр
rastr.Save('Trajectory.ut2', 'C:\Program Files (x86)\RastrWin3\RastrWin3\SHABLON\траектория утяжеления.ut2')
rastr.Load(1, 'Trajectory.ut2', 'C:\Program Files (x86)\RastrWin3\RastrWin3\SHABLON\траектория утяжеления.ut2')
Trajectory = pd.read_csv('vector.csv')
# Выделение траектории для нагрузки
LoadTrajectory = Trajectory[Trajectory['variable'] == 'pn']
LoadTrajectory = LoadTrajectory.rename(columns = {'variable':'pn', 'value':'pn_value', 'tg':'pn_tg'})
#LoadTrajectory.to_csv('LoadTrajectory.csv', index=False)
# Выделение траектории для генерации
GenTrajectory = Trajectory[Trajectory['variable'] == 'pg']
GenTrajectory = GenTrajectory.rename(columns = {'variable':'pg', 'value':'pg_value', 'tg':'pg_tg'})
#GenTrajectory.to_csv('GenTrajectory.csv', index=False)
# Создаем единый датарейм для исключения ошибок повторения узлов в траектории утяжеления
FinishedTrajectory = pd.merge(left = GenTrajectory, right = LoadTrajectory,
                              left_on = 'node', right_on = 'node', how = 'outer')
FinishedTrajectory = FinishedTrajectory.fillna(0)
# Загрузка траектории в Растр итерациями
i = 0
for index, row in FinishedTrajectory.iterrows():
    rastr.Tables('ut_node').AddRow()
    rastr.Tables('ut_node').Cols('ny').SetZ(i, row['node'])
    if pd.notnull(row['pg']):
      rastr.Tables('ut_node').Cols('pg').SetZ(i, row['pg_value'])
      rastr.Tables('ut_node').Cols('tg').SetZ(i, row['pg_tg'])
      if pd.notnull(row['pn']):
         rastr.Tables('ut_node').Cols('pn').SetZ(i, row['pn_value'])
         rastr.Tables('ut_node').Cols('tg').SetZ(i, row['pn_tg'])
    i = i + 1

# Код для проверки заполнения таблицы
#New = rastr.Save('Trajectory.ut2', 'C:\Program Files (x86)\RastrWin3\RastrWin3\SHABLON\траектория утяжеления.ut2')

# Загрузка сечения
flowgate = pd.read_json('flowgate.json')
flowgate = flowgate.T
flowgate_path = shablon_path + '\сечения.sch'
rastr.Save('Flowgate.sch', flowgate_path)
rastr.Load(1, 'Flowgate.sch', flowgate_path)
i = 0
serial_number_of_flowgate = 1
position_of_flowgate = 0
rastr.Tables('sechen').AddRow()
rastr.Tables('sechen').Cols('ns').SetZ(position_of_flowgate, serial_number_of_flowgate)
for index, row in flowgate.iterrows():
    rastr.Tables('grline').AddRow()
    rastr.Tables('grline').Cols('ns').SetZ(i, serial_number_of_flowgate)
    rastr.Tables('grline').Cols('ip').SetZ(i, row['ip'])
    rastr.Tables('grline').Cols('iq').SetZ(i, row['iq'])
    i = i + 1
    
# Код для проверки заполнения таблицы
#New = rastr.Save('Flowgate.sch', 'C:\Program Files (x86)\RastrWin3\RastrWin3\SHABLON\сечения.sch')

# Загрузка нормативных возмущений
faults = pd.read_json('faults.json')
faults = faults.T

# Создаем датафрейм с тангенсами
trajectory = FinishedTrajectory
node_table = rastr.Tables('node')
node_ny = node_table.Cols('ny')
node_pn = node_table.Cols('pn')
node_qn = node_table.Cols('qn')
i = 0
ger = {}
ger['node'] = []
ger['tg'] = []
while(i < node_table.Size):
    current_ny = node_ny.Z(i)
    for index, row in trajectory.iterrows():
        if (current_ny == row['node']):
            if (row['pn_tg'] == 1):
                tg = node_qn.Z(i) / node_pn.Z(i)
                ger['node'].append(i)
                ger['tg'].append(tg)
    i = i + 1
tg_dataframe = pd.DataFrame.from_dict(ger)

# Из-за отсутствия пояснений к тому 
# как работать с утяжеление через АстраЛиб 
# напишем функцию утяжеления осуществляющую утяжеление через таблицу узлов
def ut(trajectory, constant_tg):
    # Определяем режим
    node_table = rastr.Tables('node')
    node_ny = node_table.Cols('ny')
    node_pn = node_table.Cols('pn')
    node_qn = node_table.Cols('qn')
    node_pg = node_table.Cols('pg')
    node_qg = node_table.Cols('qg')
    # Цикл выполняет итерацию по таблице растра, в процессе которой идет итерация по
    # датафрейму с траекторией и перезапись значений для шага утяжеления
    i = 0
    while(i < node_table.Size):
        current_ny = node_ny.Z(i)
        for index, row in trajectory.iterrows():
            if (current_ny == row['node']):
                prev_pn = node_pn.Z(i)
                prev_pg = node_pg.Z(i)
                prev_qn = node_qn.Z(i)
                if (row['pn_tg'] == 1):
                    for index2, row2 in constant_tg.iterrows():
                        if (current_ny == row2['node']):
                            rastr.Tables('node').Cols('qn').SetZ(i, prev_pn + (row['pn_value'] * constant_tg.loc[constant_tg['node'] == current_ny, 'tg'].values[0]))
                #if (row['pg_tg'] == 1):
                #    prev_qg = node_qg.Z(i)
                #    new_qg = prev_qg * (prev_pg + row['pg_value']) / prev_pg
                #    rastr.Tables('node').Cols('qg').SetZ(i, new_qg)
                rastr.Tables('node').Cols('pn').SetZ(i, prev_pn + row['pn_value'])
                rastr.Tables('node').Cols('pg').SetZ(i, prev_pg + row['pg_value'])
        i = i + 1
    rastr.rgm('p')

# Расчет МДП по критерию 1
# Коэффициент запаса статичекой апериодической устойчивости в нормальной схеме
result_data = pd.DataFrame(columns = ['Criteria', 'MDP'])
# Проводим утяжеление по одному шагу
rastr.Load(1, 'regimeMDP.rg2', 'C:\Program Files (x86)\RastrWin3\RastrWin3\SHABLON\режим.rg2')
status = 0
while status == 0:
    ut(FinishedTrajectory, tg_dataframe)
    P_limit = rastr.Tables('sechen').Cols('psech').Z(position_of_flowgate)
    status = rastr.rgm('p')
    
mdp_1 = abs(P_limit) * 0.8 - 30
result_criteria_1 = {'Criteria':'20% запас в норм схеме', 'MDP':mdp_1}
result_data = result_data.append(result_criteria_1, ignore_index = True)

# Расчет МДП по критерию 2
# Коэффициент запаса по напряжению в узлах нагрузки в нормальной схеме
rastr.Load(1, 'regimeMDP.rg2', 'C:\Program Files (x86)\RastrWin3\RastrWin3\SHABLON\режим.rg2')
status = 0
voltage_dip = False
while (status == 0 and voltage_dip == False):
    ut(FinishedTrajectory, tg_dataframe)
    P_limit_2 = rastr.Tables('sechen').Cols('psech').Z(position_of_flowgate)
    k = 0
    while(i < rastr.Tables('node').Size):
        if (rastr.Tables('node').Cols('vras').Z(k) <= (rastr.Tables('node').Cols('uhom').Z(k) * 0.7) * 1.15):
            voltage_dip = True
        k = k + 1
    status = rastr.rgm('p')
    
mdp_2 = abs(P_limit_2) - 30
result_criteria_2 = {'Criteria':'запас по напряжению в норм схеме', 'MDP':mdp_2}
result_data = result_data.append(result_criteria_2, ignore_index = True)

# Расчет МДП по критерию 3
# Коэффициент запаса статичекой апериодической устойчивости в послеаварийном режиме
rastr.Load(1, 'regimeMDP.rg2', 'C:\Program Files (x86)\RastrWin3\RastrWin3\SHABLON\режим.rg2')
vetv_table = rastr.Tables('vetv')
prelim_data_3 = pd.DataFrame(columns = ['Fault-node_index', 'MDP'])
i = 0
while(i < vetv_table.Size):
    current_ip = vetv_table.Cols('ip').Z(i)
    current_iq = vetv_table.Cols('iq').Z(i)
    current_np = vetv_table.Cols('np').Z(i)
    for index, row in faults.iterrows():
        if (current_ip == row['ip'] and
            current_iq == row['iq'] and
            current_np == row['np']):
            rastr.Load(1, 'regimeMDP.rg2', 'C:\Program Files (x86)\RastrWin3\RastrWin3\SHABLON\режим.rg2')
            vetv_table = rastr.Tables('vetv')
            vetv_table.Cols('sta').SetZ(i, row['sta'])
            rastr.rgm('p')
            status = 0
            while status == 0:
                ut(FinishedTrajectory, tg_dataframe)
                status = rastr.rgm('p')
            vetv_table.Cols('sta').SetZ(i, 0)
            rastr.rgm('p')
            P_limit_3_prelim = rastr.Tables('sechen').Cols('psech').Z(position_of_flowgate)
            mdp_3_prelim = abs(P_limit_3_prelim) * 0.92 - 30
            prelim_criteria_3 = {'Fault-node_index':i, 'MDP':mdp_3_prelim}
            prelim_data_3 = prelim_data_3.append(prelim_criteria_3, ignore_index = True)
    i = i + 1

mdp_3 = abs(prelim_data_3['MDP'].min())
result_criteria_3 = {'Criteria':'8% запас в послеаварийной схеме', 'MDP':mdp_3}
result_data = result_data.append(result_criteria_3, ignore_index = True)

# Расчет МДП по критерию 4
# Коэффициент запаса по напряжению в узлах нагрузки в послеаварийном режиме
rastr.Load(1, 'regimeMDP.rg2', 'C:\Program Files (x86)\RastrWin3\RastrWin3\SHABLON\режим.rg2')
prelim_data_4 = pd.DataFrame(columns = ['Fault-node_index', 'MDP'])
j = 0
while(j < rastr.Tables('vetv').Size):
    current_ip = rastr.Tables('vetv').Cols('ip').Z(j)
    current_iq = rastr.Tables('vetv').Cols('iq').Z(j)
    current_np = rastr.Tables('vetv').Cols('np').Z(j)
    for index, row in faults.iterrows():
        if (current_ip == row['ip'] and
            current_iq == row['iq'] and
            current_np == row['np']):
            rastr.Load(1, 'regimeMDP.rg2', 'C:\Program Files (x86)\RastrWin3\RastrWin3\SHABLON\режим.rg2')
            rastr.Tables('vetv').Cols('sta').SetZ(j, row['sta'])
            rastr.rgm('p')
            status = 0
            voltage_dip = False
            while (status == 0 and voltage_dip == False):
                ut(FinishedTrajectory, tg_dataframe)
                status = rastr.rgm('p')
                k = 0
                while(k < rastr.Tables('node').Size):
                    if (rastr.Tables('node').Cols('vras').Z(k) <= (rastr.Tables('node').Cols('uhom').Z(k) * 0.7) * 1.1):
                        voltage_dip = True
                    k = k + 1

                #print(j, voltage_dip, round(rastr.Tables('sechen').Cols('psech').Z(position_of_flowgate), 2), status)
            P_limit_4_prelim = rastr.Tables('sechen').Cols('psech').Z(position_of_flowgate)
            mdp_4_prelim = abs(P_limit_4_prelim) - 30
            prelim_criteria_4 = {'Fault-node_index':j, 'MDP':mdp_4_prelim}
            prelim_data_4 = prelim_data_4.append(prelim_criteria_4, ignore_index = True)
    j = j + 1

mdp_4 = abs(prelim_data_4['MDP'].min())
result_criteria_4 = {'Criteria':'запас по напряжению в послеаварийной схеме', 'MDP':mdp_4}
result_data = result_data.append(result_criteria_4, ignore_index = True)

# Расчет МДП по критерию 5
# Допустимая токовая нагрузка в нормальной
rastr.Load(1, 'regimeMDP.rg2', 'C:\Program Files (x86)\RastrWin3\RastrWin3\SHABLON\режим.rg2')
prelim_data_5 = pd.DataFrame(columns = ['line_index', 'MDP'])
max_reached = False
status = 0
while (status == 0 and max_reached == False):
    status = rastr.rgm('p')
    j = 0
    while(j < rastr.Tables('vetv').Size):
        ut(FinishedTrajectory, tg_dataframe)
        if (round(rastr.Tables('vetv').Cols('zag_i').Z(j), 2) >= 0.10):
            P_limit_5 = rastr.Tables('sechen').Cols('psech').Z(position_of_flowgate)
            prelim_criteria_5 = {'line_index':j, 'MDP':P_limit_5}
            prelim_data_5 = prelim_data_5.append(prelim_criteria_5, ignore_index = True)
            max_reached = True
        j = j + 1
    
P_limit_5_final = abs(prelim_data_5['MDP'].abs().min())    
mdp_5 = abs(P_limit_5_final) - 30
result_criteria_5 = {'Criteria':'токовая загрузка в норм схеме', 'MDP':mdp_5}
result_data = result_data.append(result_criteria_5, ignore_index = True)

# Расчет МДП по критерию 6
# Допустимая токовая нагрузка в послеаварийной схеме
rastr.Load(1, 'regimeMDP.rg2', 'C:\Program Files (x86)\RastrWin3\RastrWin3\SHABLON\режим.rg2')
prelim_data_6 = pd.DataFrame(columns = ['line_index', 'MDP'])
j = 0
while(j < rastr.Tables('vetv').Size):
    current_ip = rastr.Tables('vetv').Cols('ip').Z(j)
    current_iq = rastr.Tables('vetv').Cols('iq').Z(j)
    current_np = rastr.Tables('vetv').Cols('np').Z(j)
    for index, row in faults.iterrows():
        if (current_ip == row['ip'] and
            current_iq == row['iq'] and
            current_np == row['np']):
            rastr.Load(1, 'regimeMDP.rg2', 'C:\Program Files (x86)\RastrWin3\RastrWin3\SHABLON\режим.rg2')
            rastr.Tables('vetv').Cols('sta').SetZ(j, row['sta'])
            max_reached = False
            status = 0
            while (status == 0 and max_reached == False):
                ut(FinishedTrajectory, tg_dataframe)
                status = rastr.rgm('p')
                k = 0
                while(k < rastr.Tables('vetv').Size):
                    if (round(rastr.Tables('vetv').Cols('zag_i_av').Z(k), 2) >= 0.10):
                        P_limit_6 = rastr.Tables('sechen').Cols('psech').Z(position_of_flowgate)
                        prelim_criteria_6 = {'line_index':k, 'MDP':P_limit_6}
                        prelim_data_6 = prelim_data_5.append(prelim_criteria_6, ignore_index = True)
                        max_reached = True
                    k = k + 1
    j = j + 1            
    
P_limit_6_final = abs(prelim_data_6['MDP'].abs().min())    
mdp_6 = abs(P_limit_6_final) - 30
result_criteria_6 = {'Criteria':'токовая загрузка в послеаварийной схеме', 'MDP':mdp_6}
result_data = result_data.append(result_criteria_6, ignore_index = True)
result_data.head(6)
print(result_data)