"""Atualiza os dados embutidos no index.html a partir da Base Geral (planilha exportada como .xlsx).

Uso:
    python3 ferramentas/atualizar_site.py "REDONDETA — Base Geral do Acervo Vivo 2021–2026.xlsx" index.html

Requer: pip install openpyxl

O que o script faz:
- Lê cada aba da planilha e monta o objeto REDONDETA_DATA no formato que o site usa (todos os valores como texto).
- Mantém do index.html atual o que não vem da planilha: textos da Rosa (meta), textos dos núcleos centrais,
  Campanhas e Arcos, e as seleções editoriais de termos de Mundo e Cosmologia e de Aventuraulas.
- Não publica registros com Status documental "Superado".
- Troca nomes de alunos conhecidos pelas iniciais (rede de segurança; o certo é corrigir na planilha).
- Grava o novo REDONDETA_DATA no index.html e atualiza o número de registros mostrado antes do JavaScript rodar.
- Mostra avisos de qualidade (IDs duplicados, nomes de alunos encontrados, termo "canônico" em textos).
"""
import collections
import copy
import datetime
import json
import re
import sys

import openpyxl

# Alunos reais aparecem somente por iniciais (00_Resumo, princípio 3). Acrescente aqui novos casos encontrados.
REDACOES = [
    (r'Nick Hisamori Curti', 'N.C.'), (r'Nick Curti', 'N.C.'), (r'\bNICK\b', 'N.C.'), (r'\bNick\b', 'N.C.'),
    (r'BRUNA CAMARGO', 'B.C.'), (r'\bBruna\b', 'B.C.'),
]
SEM_COLUNAS_MESTRES = ('Imagem', 'Drive_Imagem_ID', 'URL_Imagem_Drive')


def texto(v):
    """Valor de célula no formato do site: sempre texto."""
    if v is None:
        return ''
    if isinstance(v, bool):
        return 'SIM' if v else 'NÃO'
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else str(v)
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime('%d/%m/%Y')
    return str(v).strip()


def ler_aba(wb, nome, linha_cabecalho=1, sem=()):
    linhas = list(wb[nome].iter_rows(min_row=linha_cabecalho, values_only=True))
    cab = [(str(h).strip() if h is not None else None) for h in linhas[0]]
    registros = []
    for r in linhas[1:]:
        if not any(c not in (None, '') for c in r):
            continue
        registros.append({h: texto(v) for h, v in zip(cab, r) if h and h not in sem})
    return registros


def id_de(r):
    return r.get('ID Universal') or r.get('ID') or r.get('ID Personagem') or ''


def ler_dados_do_site(html):
    ini = html.index('const REDONDETA_DATA = ') + len('const REDONDETA_DATA = ')
    fim = html.index(';\n', ini)
    return json.loads(html[ini:fim])


def montar(wb, atual):
    novo = copy.deepcopy(atual)
    C = novo['compartments']
    antigo = atual['compartments']

    # Termos: Mundo e Cosmologia (N) e Aventuraulas (S) são seleções editoriais; os demais ficam só para a busca
    termos = ler_aba(wb, '11_Termos_Conceitos')
    por_id = {r['ID']: r for r in termos}
    ids_n = [r['ID'] for r in antigo['mundo_cosmologia']]
    ids_s = [r['ID'] for r in antigo['aventuraulas']]
    C['mundo_cosmologia'] = [por_id[i] for i in ids_n if i in por_id]
    C['aventuraulas'] = [por_id[i] for i in ids_s if i in por_id]
    C['termos_complementares'] = [r for r in termos if r['ID'] not in set(ids_n) | set(ids_s)]

    # Locais: Continentes = Tipo "Continente/região" (+ exceções já no site, ex.: LOC-028); Lugares = os demais tipos
    locais = ler_aba(wb, '12_Locais_Territorios')
    extras = {r['ID'] for r in antigo['continentes']}
    C['continentes'] = [r for r in locais if r['Tipo'] == 'Continente/região' or r['ID'] in extras]
    C['lugares'] = [r for r in locais if r['Tipo'] != 'Continente/região']

    C['povos'] = ler_aba(wb, '14_Povos_Etnias')
    npcs = ler_aba(wb, '03_NPCs')
    for r in npcs:
        r['Categoria_Personagem'] = 'NPC / Habitante'
    viloes = ler_aba(wb, '04_Viloes_Capangas')
    for r in viloes:
        r['Categoria_Personagem'] = 'Força Antagônica / Vilão'
    C['personagens_redondeta'] = npcs + viloes
    C['alunos_herois'] = ler_aba(wb, '01_Alunos_Herois')
    C['mestres_mestrinhos'] = ler_aba(wb, '02_Mestres_Mestrinhos', sem=SEM_COLUNAS_MESTRES)
    C['instituicoes_organizacoes'] = {
        'organizacoes': ler_aba(wb, '06_Organizacoes'),
        'mazelas': ler_aba(wb, '05_Mazelas'),
        'orbes': ler_aba(wb, '05B_Orbes'),
    }
    C['eventos'] = ler_aba(wb, '10_Eventos_Campanhas')
    C['percursos_autorais'] = ler_aba(wb, '19_Percursos_Autorais')
    C['magias'] = ler_aba(wb, '08_Magias_Poderes')
    C['saberes_referencias'] = {
        'referencias': ler_aba(wb, '15_Referencias_Cientificas'),
        'fontes': ler_aba(wb, '13_Fontes_Referencias'),
    }
    C['itens'] = ler_aba(wb, '09_Itens_Artefatos')
    if '22_Quests_Tarefas' in wb.sheetnames:
        C['quests_tarefas'] = ler_aba(wb, '22_Quests_Tarefas')
    # campanhas_arcos não tem aba própria: continua como está no site

    novo['relations'] = ler_aba(wb, '16_Relacoes')
    novo['indice'] = ler_aba(wb, '18_Indice_Entidades')
    novo['fichas_enciclopedicas'] = ler_aba(wb, '17_Fichas_Enciclopedicas')
    novo['acervo_imagens'] = ler_aba(wb, '20_Acervo_Imagens')
    intros = ler_aba(wb, '21_Introducoes_Compartimentos')
    for r in intros:
        r['Exibir_Introducao'] = str(r.get('Exibir_Introducao', '')).strip().upper() not in ('NÃO', 'NAO', 'FALSE', 'N')
    novo['introducoes_compartimentos'] = intros

    # Lacunas do núcleo Arquivo: linhas do painel com ID e prioridade
    painel = wb['00_Painel_de_Lacunas']
    linha_id = next(i for i, r in enumerate(painel.iter_rows(values_only=True), 1) if r and r[0] == 'ID')
    lacunas = [r for r in ler_aba(wb, '00_Painel_de_Lacunas', linha_cabecalho=linha_id) if r.get('ID') and r.get('Prioridade')]

    # Registros "Superado" não são publicados
    superados = {r['ID Universal'] for r in novo['indice'] if r.get('Status documental', '').lower().startswith('superad')}

    def publicado(r):
        return id_de(r) not in superados and not r.get('Status documental', '').lower().startswith('superad')
    for k, v in list(C.items()):
        C[k] = {kk: [r for r in vv if publicado(r)] for kk, vv in v.items()} if isinstance(v, dict) else [r for r in v if publicado(r)]
    novo['relations'] = [r for r in novo['relations'] if publicado(r)]
    novo['indice'] = [r for r in novo['indice'] if publicado(r)]

    # Núcleos centrais
    H = novo['central_hubs']
    H['CENTRO_ENGRENAGEM']['relations'] = novo['relations']
    H['CENTRO_ARQUIVO']['fichas_enciclopedicas'] = novo['fichas_enciclopedicas']
    H['CENTRO_ARQUIVO']['lacunas'] = lacunas

    # Privacidade (rede de segurança)
    def limpa(x):
        if isinstance(x, dict):
            return {k: limpa(v) for k, v in x.items()}
        if isinstance(x, list):
            return [limpa(v) for v in x]
        if isinstance(x, str):
            for a, b in REDACOES:
                x = re.sub(a, b, x)
        return x
    novo = limpa(novo)
    C, H, ind = novo['compartments'], novo['central_hubs'], novo['indice']

    # Números citados nos textos da Rosa acompanham os dados
    n_ev, n_pe, n_it = len(C['eventos']), len(C['personagens_redondeta']), len(C['itens'])
    n_fo, n_re = len(C['saberes_referencias']['fontes']), len(C['saberes_referencias']['referencias'])
    for m in novo['meta']:
        for campo in ('subtitle', 'desc'):
            t = m.get(campo, '')
            t = re.sub(r'\b\d+ Marcos Históricos', f'{n_ev} Marcos Históricos', t)
            t = re.sub(r'\b\d+ acontecimentos', f'{n_ev} acontecimentos', t)
            t = re.sub(r'\b\d+ personagens do mundo', f'{n_pe} personagens do mundo', t)
            t = re.sub(r'\b\d+ itens concebidos', f'{n_it} itens concebidos', t)
            t = re.sub(r'\b\d+ fontes documentais', f'{n_fo} fontes documentais', t)
            t = re.sub(r'\b\d+ referências científicas', f'{n_re} referências científicas', t)
            m[campo] = t

    status = lambda k: sum(1 for r in ind if r['Status documental'] == k)
    H['CENTRO_ARQUIVO']['stats'] = {
        'total_entities': len(ind), 'confirmed': status('Confirmado'), 'probable': status('Provável'),
        'unknown': status('Desconhecido'), 'integrity_rate': f"{status('Confirmado') / len(ind) * 100:.1f}%",
        'relations_count': len(novo['relations']), 'sources_count': n_fo,
    }
    novo['summary'] = {
        'registros_indexados': len(ind), 'entidades_total': len(ind), 'confirmadas': status('Confirmado'),
        'relacoes': len(novo['relations']), 'alunos_herois': len(C['alunos_herois']), 'mestres': len(C['mestres_mestrinhos']),
        'npcs': sum(1 for r in C['personagens_redondeta'] if id_de(r).startswith('NPC-')),
        'viloes': sum(1 for r in C['personagens_redondeta'] if id_de(r).startswith('VIL-')),
        'magias': len(C['magias']), 'itens': n_it, 'eventos': n_ev, 'locais': len(locais),
        'organizacoes': len(C['instituicoes_organizacoes']['organizacoes']), 'povos': len(C['povos']),
        'mazelas': len(C['instituicoes_organizacoes']['mazelas']), 'orbes': len(C['instituicoes_organizacoes']['orbes']),
        'ciencias': n_re, 'fontes': n_fo, 'fichas_enciclopedicas': len(novo['fichas_enciclopedicas']),
    }
    return novo, sorted(superados)


def avisos(novo):
    out = []
    for nome, lista, chave in (('20_Acervo_Imagens', novo['acervo_imagens'], 'ID Ativo'), ('18_Indice_Entidades', novo['indice'], 'ID Universal'),
                               ('16_Relacoes', novo['relations'], 'ID')):
        dup = [k for k, v in collections.Counter(r.get(chave) for r in lista).items() if k and v > 1]
        if dup:
            out.append(f'ID duplicado em {nome}: {dup}')
    tudo = json.dumps(novo, ensure_ascii=False)
    canonico = [t for t in re.findall(r'[^"]{0,40}can[ôo]nic[^"]{0,30}', tudo) if 'Nome do Arquivo' not in t]
    if canonico:
        out.append(f'Termo "canônico" em textos: {canonico[:5]}')
    sem_registro = [e['ID Universal'] for e in novo['indice'] if not e['ID Universal'].startswith('REL-')
                    and e['ID Universal'] not in {id_de(r) for v in novo['compartments'].values()
                                                  for r in (sum(v.values(), []) if isinstance(v, dict) else v)}]
    if sem_registro:
        out.append(f'IDs do índice sem registro completo no site ({len(sem_registro)}): {sem_registro[:10]}')
    return out


def gravar(html, novo):
    js = json.dumps(novo, ensure_ascii=False, separators=(',', ':')).replace('</', '<\\/')
    ini = html.index('const REDONDETA_DATA = ')
    fim = html.index(';\n', ini)
    html = html[:ini] + 'const REDONDETA_DATA = ' + js + html[fim:]
    n = f"{len(novo['indice']):,}".replace(',', '.')
    return re.sub(r'(<span data-acervo="registros">)[\d.]+(</span>)', r'\g<1>' + n + r'\g<2>', html)


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    planilha = sys.argv[1]
    caminho_html = sys.argv[2] if len(sys.argv) > 2 else 'index.html'
    html = open(caminho_html, encoding='utf-8').read()
    wb = openpyxl.load_workbook(planilha, data_only=True)
    novo, superados = montar(wb, ler_dados_do_site(html))
    open(caminho_html, 'w', encoding='utf-8').write(gravar(html, novo))
    C = novo['compartments']
    print(f"index.html atualizado: {len(novo['indice'])} registros, {len(novo['relations'])} relações.")
    print('Compartimentos:', {k: (sum(len(x) for x in v.values()) if isinstance(v, dict) else len(v)) for k, v in C.items()})
    print('Não publicados (Superado):', superados or 'nenhum')
    for a in avisos(novo):
        print('AVISO:', a)


if __name__ == '__main__':
    main()
