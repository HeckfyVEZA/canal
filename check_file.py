def check(all_noms):
    header = all_noms[0][0]
    chlist = [[header]]
    for noms in all_noms:
        if noms[0] == header:
            chlist[-1].append(f"{noms[1]} | {noms[2]} шт.")
        else:
            header = noms[0]
            chlist.append([header])
            chlist[-1].append(f"{noms[1]} | {noms[2]} шт.")
    max_len = max([len(cist) for cist in chlist])
    chlist = [cist+[None]*(max_len - len(cist)) for cist in chlist]
    chlist = list(zip(*chlist))
    return chlist