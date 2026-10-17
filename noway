"""Sugestões de termos a partir do Tesauro da Câmara (TECAD) - arquivo tecadTermos.xlsx
dos Dados Abertos. Funciona offline: os dados ficam em tecad_termos.json.gz.
Para atualizar, coloque um tecadTermos.xlsx novo na pasta do app (a base é reconstruída sozinha)."""
import gzip
import json
import re
import unicodedata
from pathlib import Path

PASTA = Path(__file__).parent
GZ = PASTA / "tecad_termos.json.gz"
XLSX = PASTA / "tecadTermos.xlsx"

STOP = {"de", "da", "do", "das", "dos", "a", "o", "as", "os", "e", "em", "no", "na", "nos", "nas", "para",
        "por", "com", "contra", "sobre", "ao", "aos", "um", "uma"}


def norm(s):
    s = unicodedata.normalize("NFKD", s or "").encode("ascii", "ignore").decode().lower()
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9 ]+", " ", s)).strip()


def stem(p):
    """Radical simples para o português (plural/gênero), suficiente para casar idoso/idosa/idosos."""
    p = norm(p)
    if len(p) <= 3:
        return p
    for suf, rep in (("oes", "ao"), ("aes", "ao"), ("ais", "al"), ("eis", "el"), ("ois", "ol"), ("res", "r"),
                     ("zes", "z"), ("ns", "m"), ("s", "")):
        if p.endswith(suf) and len(p) - len(suf) >= 3:
            p = p[: len(p) - len(suf)] + rep
            break
    if len(p) > 4 and p[-1] in "aoe":
        p = p[:-1]
    return p


def chave(texto):
    """Chave de comparação: radicais das palavras significativas, em ordem."""
    return " ".join(stem(w) for w in norm(texto).split() if w not in STOP)


def _lista(valor):
    return re.findall(r'"([^"]*)"', valor) if valor else []


def ler_xlsx(fonte):
    """Lê o tecadTermos.xlsx (caminho ou arquivo em memória) e devolve a lista de termos."""
    import openpyxl

    ws = openpyxl.load_workbook(fonte, read_only=True).active
    ws.reset_dimensions()
    it = ws.iter_rows(values_only=True)
    cab = list(next(it))
    ix = {c: i for i, c in enumerate(cab)}
    if "termo" not in ix:
        raise ValueError("o arquivo não parece ser o tecadTermos.xlsx (coluna 'termo' não encontrada)")
    dados = []
    for r in it:
        termo = (r[ix["termo"]] or "").strip()
        if not termo:
            continue
        g = lambda c: _lista(r[ix[c]]) if c in ix and r[ix[c]] else []
        nota = (r[ix["notasExplicativas"]] or "").strip() if "notasExplicativas" in ix else ""
        dados.append([termo, g("use"), g("usadoPara"), g("termosEspecificos"), g("termosGenericos"),
                      g("termosRelacionados"), nota[:400]])
    return dados


def salvar_base(dados, destino=GZ):
    with gzip.open(destino, "wt", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, separators=(",", ":"))


def construir_de_xlsx(xlsx=XLSX, destino=GZ):
    dados = ler_xlsx(xlsx)
    salvar_base(dados, destino)
    return len(dados)


class Tesauro:
    def __init__(self, dados=None):
        if dados is None:
            if XLSX.exists() and (not GZ.exists() or XLSX.stat().st_mtime > GZ.stat().st_mtime):
                construir_de_xlsx()
            with gzip.open(GZ, "rt", encoding="utf-8") as f:
                dados = json.load(f)
        self.itens = dados
        self.por_termo = {}
        for it in self.itens:
            self.por_termo.setdefault(norm(it[0]), it)
        import math
        self.por_chave = {}
        self.tokens = []
        df = {}
        for k, it in enumerate(self.itens):
            self.por_chave.setdefault(chave(it[0]), []).append(k)
            toks = set(chave(it[0]).split())
            for u in it[2]:
                toks |= set(chave(u).split())
            self.tokens.append(toks)
            for w in toks:
                df[w] = df.get(w, 0) + 1
        N = len(self.itens)
        self.idf = {w: math.log((N + 1) / (c + 1)) + 1 for w, c in df.items()}
        # índice inverso: quem cita o termo (TE de X <-> TG de Y; TR e UP são bidirecionais)
        self.inverso = {}
        for it in self.itens:
            for pos, rel in ((2, "UP"), (3, "TE"), (4, "TG"), (5, "TR")):
                for alvo in it[pos]:
                    self.inverso.setdefault(norm(alvo), []).append((it[0], rel))

    def _resolve(self, item):
        """Se o termo é remissivo (USE), devolve o(s) descritor(es) preferido(s)."""
        if item[1]:
            out = [self.por_termo.get(norm(u)) for u in item[1]]
            out = [o for o in out if o]
            if out:
                return out
        return [item]

    def buscar(self, tema, limite=25):
        """Termos do tesauro que correspondem ao tema (descritores), do mais ao menos relevante."""
        t = norm(tema)
        if not t:
            return []
        palavras = [p for p in t.split() if p not in STOP]
        pont = {}
        for it in self.itens:
            n = norm(it[0])
            if n == t:
                s = 0
            elif n.startswith(t):
                s = 1
            elif t in n:
                s = 2
            elif palavras and all(p in n.split() for p in palavras):
                s = 3
            else:
                continue
            for d in self._resolve(it):
                k = d[0]
                pont[k] = (min(s, pont[k][0]) if k in pont else s, len(k), d)
        ordem = sorted(pont.values(), key=lambda x: (x[0], x[1]))
        return [x[2][0] for x in ordem[:limite]]

    def exatos(self, tema):
        it = self.por_termo.get(norm(tema))
        return [d[0] for d in self._resolve(it)] if it else []

    def relacoes(self, termos_base):
        """{'base': [...], 'sinonimos': [...], 'especificos': [...], 'genericos': [...], 'relacionados': [...]}"""
        saida = {"base": [], "sinonimos": [], "especificos": [], "genericos": [], "relacionados": []}
        for nome in termos_base:
            it = self.por_termo.get(norm(nome))
            if not it:
                continue
            saida["base"].append(it[0])
            for chave, pos in (("sinonimos", 2), ("especificos", 3), ("genericos", 4), ("relacionados", 5)):
                saida[chave] += it[pos]
        vistos = set()
        for k in saida:                                    # sem repetição entre as listas
            novo = []
            for t in saida[k]:
                if norm(t) not in vistos:
                    vistos.add(norm(t))
                    novo.append(t)
            saida[k] = novo
        return saida


ARESTAS = {"UP": 2.0, "TE": 0.9, "TR": 0.7, "TG": 0.4}
NOMES = {"UP": "Sinônimo (UP)", "TE": "Específico (TE)", "TR": "Relacionado (TR)", "TG": "Genérico (TG)",
         "LEX": "Contém o tema", "LEXP": "Contém parte do tema", "NOTA": "Citado na nota explicativa"}


def _vizinhos(self, nome, incluir_tg):
    """(termo, tipo de relação) a partir de um termo: relações diretas e inversas."""
    it = self.por_termo.get(norm(nome))
    saida = []
    if it:
        for pos, rel in ((2, "UP"), (3, "TE"), (5, "TR")):
            saida += [(x, rel) for x in it[pos]]
        if incluir_tg:
            saida += [(x, "TG") for x in it[4]]
    for origem, rel in self.inverso.get(norm(nome), []):
        # origem lista 'nome' como... rel_de_origem -> relação vista de 'nome'
        if rel == "TG":
            saida.append((origem, "TE"))       # origem tem 'nome' como genérico => origem é específico de 'nome'
        elif rel == "TR":
            saida.append((origem, "TR"))
        elif rel == "TE" and incluir_tg:
            saida.append((origem, "TG"))
    return saida


def _conceitos(self, textos):
    """Descritores do tesauro que correspondem a partes do texto (palavras e trechos contínuos),
    p.ex. 'violência contra o idoso' -> Violência, Idoso (e 'pessoa idosa' -> Idoso)."""
    achados = {}
    for texto in textos:
        pal = [w for w in norm(texto).split() if w not in STOP]
        for a in range(len(pal)):
            for b in range(a + 1, len(pal) + 1):
                trecho = pal[a:b]
                if len(trecho) == 1 and self.idf.get(stem(trecho[0]), 9) < 4.2:
                    continue                      # palavra comum demais (nacional, programa...)
                for k in self.por_chave.get(" ".join(stem(w) for w in trecho), []):
                    for d in self._resolve(self.itens[k]):
                        peso = 0.85 if len(trecho) > 1 else 0.5
                        achados[d[0]] = max(achados.get(d[0], 0), peso)
    return achados


def conceitos_do_tema(self, tema, bases=()):
    """Conceitos (descritores) que compõem o tema, p.ex. ['Idoso', 'Violência']."""
    base_n = {norm(b) for b in bases}
    c = {k: p for k, p in _conceitos(self, [tema] + list(bases)).items() if norm(k) not in base_n}
    return [k for k, _ in sorted(c.items(), key=lambda x: -x[1])]


def expandir(self, bases, amplitude=2, max_termos=80, incluir_tg=False, incluir_lexico=True, tema="",
             conceitos_sel=None):
    """Analisa o tesauro inteiro em torno do tema:
    1) sementes = termos-base + conceitos que compõem o tema (Violência, Idoso...);
    2) expansão pelo grafo (sinônimos, específicos, relacionados, nos dois sentidos) até 'amplitude' saltos;
    3) varredura de todos os termos e notas explicativas por aderência lexical (radicais, ponderados por IDF).
    A pontuação soma as evidências; termos alcançados por vários caminhos sobem no ranking."""
    if not hasattr(self, "_notas"):
        self._notas = [set(chave(it[6]).split()) if len(it) > 6 and it[6] else set() for it in self.itens]
    bases = list(bases)
    base_n = {norm(b) for b in bases}
    conceitos = {c: p for c, p in _conceitos(self, [tema] + bases).items() if norm(c) not in base_n}
    if conceitos_sel is not None:
        conceitos = {c: p for c, p in conceitos.items() if c in conceitos_sel}

    score, info = {}, {}
    sementes = {norm(b): (b, 1.0) for b in bases}
    for c, p in conceitos.items():
        sementes[norm(c)] = (c, p)
    ja = set(sementes)

    def somar(termo, peso, rel, via):
        n = norm(termo)
        if not n or n in ja:
            return
        score[n] = score.get(n, 0) + peso
        if n not in info or peso > info[n]["peso"]:
            info[n] = {"termo": termo, "rel": rel, "via": via, "peso": peso}

    fronteira = dict(sementes)
    vistos = set(ja)
    for salto in range(1, int(amplitude) + 1):
        decaimento = 0.45 ** (salto - 1)
        novos = {}
        for n, (nome, w) in fronteira.items():
            for alvo, rel in _vizinhos(self, nome, incluir_tg and salto == 1):
                it = self.por_termo.get(norm(alvo))
                if it and it[1] and rel != "UP":      # remissivo só interessa como sinônimo (UP)
                    continue
                p = (max(w, 0.85) if rel == "UP" else w) * ARESTAS[rel] * decaimento   # sinônimo = muito aderente
                somar(alvo, p, rel, nome if (salto > 1 or norm(nome) not in base_n) else "")
                if norm(alvo) not in vistos:
                    novos[norm(alvo)] = (alvo, w * ARESTAS[rel] * 0.6)
        vistos |= set(novos)
        fronteira = dict(sorted(novos.items(), key=lambda kv: -score.get(kv[0], 0))[:120])

    if incluir_lexico:
        C = {}
        for texto in [tema] + bases:
            for w in norm(texto).split():
                if w not in STOP:
                    s = stem(w)
                    C[s] = self.idf.get(s, 6.0)
        total = sum(C.values()) or 1.0
        for k, it in enumerate(self.itens):
            if it[1]:
                continue
            n = norm(it[0])
            if n in ja:
                continue
            inter = self.tokens[k] & C.keys()
            cob = sum(C[w] for w in inter) / total if inter else 0.0
            nota = self._notas[k] & C.keys()
            cob_nota = sum(C[w] for w in nota) / total if nota else 0.0
            if cob >= 0.99:
                somar(it[0], 0.9, "LEX", "")
            elif cob >= 0.45:
                somar(it[0], 0.45 * cob, "LEXP", "")
            elif cob_nota >= 0.99:
                somar(it[0], 0.3, "NOTA", "")

    ordem = sorted(score, key=lambda n: -score[n])[: int(max_termos)]
    lista = [{"termo": b, "relacao": "Termo-base", "via": "", "pontuacao": 10.0} for b in bases]
    lista += [{"termo": c, "relacao": "Conceito do tema", "via": "", "pontuacao": 9.0 + p} for c, p in
              sorted(conceitos.items(), key=lambda x: -x[1])]
    return lista + [{"termo": info[n]["termo"], "relacao": NOMES[info[n]["rel"]], "via": info[n]["via"],
                     "pontuacao": round(score[n], 2)} for n in ordem]


Tesauro.expandir = expandir
Tesauro.conceitos_do_tema = conceitos_do_tema


def montar_or(termos):
    """Junta com OR; expressões com mais de uma palavra ficam entre aspas."""
    vistos, saida = set(), []
    for t in termos:
        t = t.strip().strip('"')
        if not t or t.lower() in vistos:
            continue
        vistos.add(t.lower())
        saida.append(f'"{t}"' if " " in t else t)
    return " OR ".join(saida)
