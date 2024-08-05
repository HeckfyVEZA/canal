from itertools import dropwhile, takewhile
from re import findall, IGNORECASE
from docx2txt import process


def l_g(spis):
    spis = list(sorted(spis, key=lambda x: x[1]))
    ispis = [spis[0]]
    for i in range(1, len(spis)):
        if spis[i][1] != spis[i-1][1]:
            ispis.append(spis[i])
        else:
            ispis[-1][2]+=spis[i][2]
    return ispis

def sashQUA(main_system_name:str, the_mask = r'\d?[A-Za-zА-Яа-яЁё]\D?\D?\D?\D?\S?\S?\S?\S?\d{1,}[А-Яа-яЁё]?'):
    all_system_names = []
    for system_name in main_system_name.replace("-", " - ").split(','):
        system_name = system_name.strip()
        system_name_all = findall(the_mask, system_name, IGNORECASE)
        system_name_all = system_name_all if system_name_all else [main_system_name]
        if len(system_name_all) > 1:
            all_positions_in_system_name = tuple(tuple(filter(None, findall(r'(\D?|\d+)', sys_name, IGNORECASE))) for sys_name in system_name_all)
            system_condition = lambda x: x[0] == x[1]
            before_changing_part = ''.join(el[0] for el in takewhile(system_condition, zip(all_positions_in_system_name[0], all_positions_in_system_name[1])))
            after_changing_part = dropwhile(system_condition, zip(all_positions_in_system_name[0], all_positions_in_system_name[1]))
            changing_part_start, changing_part_thend = next(after_changing_part)
            after_changing_part = ''.join(el[0] for el in after_changing_part)
            all_system_names += [f"{before_changing_part}{'0' * (min(len(changing_part_start), len(changing_part_thend)) - len(str(i)))}{i}{after_changing_part}" for i in range(int(changing_part_start), int(changing_part_thend) + 1)]
        else:
            all_system_names += system_name_all
    return len(all_system_names)

def c_str(oborud):
    oborud = oborud.replace("M24-SR", "M24SR").replace("F24-S", "F24S")
    mosh_slov = {"9":"12", "17":"18", "23":"24", "27":"30"}
    if "ГЕРМИК" in oborud:
        oborud = "-".join(oborud.split("-")[:-1] + ["Н", oborud.split("-")[-1]]) if len(oborud.split("-")) < 7 else oborud
    if "ЭКВ-К" in oborud and (not "," in oborud):
        oborud+=',0'
    elif "ЭКВ" in oborud:
        oborud = "-".join(oborud.split("-")[:-1] + [mosh_slov[oborud.split("-")[-1]]]) if oborud.split("-")[-1] in list(mosh_slov.keys()) else oborud
    groups = ["Канал-Регуляр", 'Канал-БОБ','Канал-ВЕНТ','Канал-ЕС','Канал-КВАРК','Канал-КВАРК-ФУД','Канал-ПКВ','Канал-КВ','Канал-ЭКВ','Канал-ВКО','Канал-ФКО','Канал-КП','Канал-КВН','Канал-ФУД-Р-КОЖ','Канал-козырек','Канал-ФУД-козырек','Канал-ВИБР','Канал-ФУД-вибр','Канал-ФУД-Р-МК','Канал-крыша','Канал-КВАРК-ФУД-РКА','Канал-КВАРК-ФУД-РКО','Канал-РВК','Канал-РВС','Канал-РКН','Канал-РПВС','Канал-сетка','Канал-ФУД-сетка','Канал-C-PKT','Канал-ПКТ','Канал-ФКК','Канал-ФКП','Канал-МК','Канал-ГКД','Канал-ГКК','Канал-ГКП','КЛАБ', 'Канал-Гермик', "Канал-ГКВ", "Канал-ФУД-ГКВ", "Канал-ФУД-Регуляр", "Канал-КОЛ", "Канал-ФУД-Тюльпан", "ГЕРМИК", 'Канал-КВАРК-П', 'К-']
    apendix = ["Клапан ", 'Блок обеззараживания ','Вентилятор ','Вентилятор ','Вентилятор ','Вентилятор ','Вентилятор ','Клапан ','Воздухонагреватель ','Воздухоохладитель ','Воздухоохладитель ','Каплеуловитель ','Воздухонагреватель ','Кожух ','Козырек ','Козырек ','Комплект основы виброизолирующий ','Комплект основы виброизолирующий ','Кронштейн ','Крыша ','Решетка ','Решетка ','Решетка ','Решетка ','Решетка ','Решетка ','Сетка ','Сетка ','Теплоутилизатор ','Теплоутилизатор ','Фильтр ','Фильтр ','Хомут ','Шумоглушитель ','Шумоглушитель ','Шумоглушитель ','Клапан ', 'Клапан ', 'Гибкая вставка ', 'Гибкая вставка ', 'Клапан ', 'Клапан ', 'Клапан ','Клапан ', 'Вентилятор ', 'Адаптер Канал-']
    try:
        indexes = [groups.index(gr) for gr in groups if oborud.startswith(gr)]
        return apendix[max(indexes)] + oborud
    except:
        return oborud

def main_device(m_d):
    headers = [findall(r"\d+\. .+", item)[0] for item in m_d if len(findall(r"\d+\. .+", item))]
    return headers

def reg_chast(strochka, vent):
    choice = {220: [[0.22, 1, 'Регулятор скорости СРМ1-230В 1А IP20'],[0.44, 2, 'Регулятор скорости СРМ2-230В 2А IP20'],[0.55, 2.5, 'Регулятор скорости СРМ2,5Щ-230В 2,5А DIN IP20'],[0.66, 3, 'Регулятор скорости СРМ3-230В 3А IP20'],[0.88, 4, 'Регулятор скорости СРМ4-230В 4А IP20'],[0.88, 5, 'Регулятор скорости СРМ5Щ-230В 5А DIN IP20'],[1.1, 5, 'Регулятор скорости СРМ5-230В 5А IP20'],[1.5, 7, 'Регулятор скорости СРМ7-230В 7А IP20']],380: [[0.37, 1.5, 'Преобразователь частоты 0,37 кВт'],[0.75, 2.3, 'Преобразователь частоты 0,75 кВт'],[1.1, 2.3, 'Преобразователь частоты 1,1 кВт'],[1.5, 3.7, 'Преобразователь частоты 1,5 кВт'],[2.2, 5, 'Преобразователь частоты 2,2 кВт'],[4, 8.5, 'Преобразователь частоты 4 кВт'],[5.5, 12, 'Преобразователь частоты 5,5 кВт'],[7.5, 16, 'Преобразователь частоты 7,5 кВт'],[11, 24, 'Преобразователь частоты 11 кВт'],[15, 30, 'Преобразователь частоты 15 кВт'],[18.5, 37, 'Преобразователь частоты 18,5 кВт'],[22, 45, 'Преобразователь частоты 22 кВт'],[30, 60, 'Преобразователь частоты 30 кВт'],[37, 75, 'Преобразователь частоты 37 кВт']]}
    ch = choice[380] if "кварк" in vent.lower() else choice[220] if 'канал-вент' in vent.lower() else choice[int(findall(r'Uпит=~(\d+) В', strochka.replace(",","."))[0])]
    for item in ch:
        if item[0] >= float(findall(r'Ny=(\d+\.?\d*) кВт', strochka.replace(",","."))[0]) and item[1] > float(findall(r'Iпот=(\d+\.?\d*) A', strochka.replace(",","."))[0]):
            return item[2]
    return "Подбор невозможен"

def main_devs(headers, main_devices):
    new_headers = []
    for head in headers:
        if len(findall(r"\d+ ?шт.", head)):
            new_headers.append(int(findall(r"(\d+) ?шт.", head)[0]))
        elif "(основной и резервный)" in head:
            new_headers.append(2)
        else:
            new_headers.append(1)

    alls = [(item+"; ").split("; ")[0] for item in main_devices if (bool(len(findall(r"Индекс: .+", item))) or item in headers)]
    nalls = [[]]
    l = 0
    for item in alls:
        try:
            nalls[l].append([c_str(item.split("декс: ")[1].replace(".",'')), new_headers[l-1]])
        except:
            nalls.append([])
            l+=1
    nalls = list(map(lambda x: x[0], nalls[1:]))
    return nalls

def addon(addons):
    ado = {'addons':[]}
    for item in addons:
        if 'вент' in item and 'да' in item:
            ado['частотник/регулятор'] = True
        elif ': ' in item and len(findall(r"(\d+) ?шт\.?", item)):
            ado['addons'].append([c_str(item.split(': ')[1].split()[0]), findall(r"(\d+) ?шт\.?", item)[0] if len(findall(r"(\d+) ?шт\.?", item)) else 1])
        # else:
        #     ado['частотник/регулятор'] = False
    return ado

def new_itog(full):
    for item in full:
        # st.write(item)
        if "ВКО" in item[1] or "ФКО" in item[1]:
            ooo = item[1].split('КО-')[1]
            if len(ooo.split('-'))==2:
                item[1]+='-4'
    return full

def infos(file):
    i = [item for item in process(file).split('\n') if len(item)]
    # st.write(i)
    aflag = False
    nindex = i.index("Название:")+1
    dindex = i.index("Дополнительное оборудование:")+1
    try:
        nah = i.index("Габаритная схема")
    except:
        nah = i.index('Габаритные размеры')
        # print(file.name)
    try:
        number = {'number': i[nindex], "amount": sashQUA(i[nindex])} # Номер системы
    except:
        try:
            number = {'number': i[nindex], "amount": len(i[nindex].split(','))}
        except:
            number = {'number': i[nindex], "amount": 1}
    main_devices = i[nindex+1:dindex]
    main_devices =  main_devs(main_device(main_devices), main_devices)
    addons = addon(i[dindex:nah])
    itogi = {'number': number, "main_devices":main_devices, 'addons':addons}
    try:
        if addons["частотник/регулятор"]:
            aflag = True
            itogi['подобранные частотники'] = []
            vents = [ve for ve in main_devices if 'вентилятор' in ve[0].lower()]
            harki = [har for har in i if "Эл.двиг: " in har]
            harki = [[harki[s], vents[s]] for s in range(len(harki))]
            for item in harki:
                itogi['подобранные частотники'].append([reg_chast(item[0], item[1][0]), item[1][1]])
    except:
        pass
    koef = itogi['number']['amount']
    name = number['number']
    citog = [[name, item[0], int(item[1])*koef] for item in itogi['main_devices']]
    citog+= [[name, item[0], int(item[1])*koef] for item in itogi['addons']['addons']]
    if aflag:
        citog+= [[name, item[0], int(item[1])*koef] for item in itogi["подобранные частотники"]]
    citog = l_g(citog)
    return new_itog(citog)

