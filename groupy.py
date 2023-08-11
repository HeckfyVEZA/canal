def grouping(spisok):
    all_nomenclature_names = list(set(list(map(lambda x: x[1], spisok))))
    itog = [[None, None, None]]
    for item in all_nomenclature_names:
        alsu = sum([su[-1] for su in spisok if item in su])
        namsu = ", ".join([su[0] for su in spisok if item in su])
        itog += [[item, alsu, namsu]]
    return list(sorted(itog[1:], key=lambda x: x[0]))