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
        dados.append([termo, g("use"), g("usadoPara"), g("termosEspecificos"), g("termosGenericos"),
                      g("termosRelacionados")])
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
