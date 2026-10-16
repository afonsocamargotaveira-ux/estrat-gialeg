"""Sugestão de termos de busca (separados por OR) sobre o tema, via API da Anthropic."""
import json
import re

MODELO_IA = "claude-sonnet-5-5"


def _prompt(tema, n_termos, ja_usados):
    evitar = f"\nNão repita estes termos, já usados em outras sequências: {', '.join(ja_usados)}." if ja_usados else ""
    return (
        "Você é um analista legislativo da Câmara dos Deputados do Brasil que monta estratégias de busca "
        "por proposições no Sileg (pesquisa textual em ementa, explicação da ementa e indexação).\n"
        f"Tema: {tema}\n"
        f"Sugira {n_termos} termos ou expressões curtas, em português do Brasil, que apareceriam na ementa ou na "
        "indexação de proposições sobre esse tema. Faça uma análise semântica: inclua sinônimos, variações de "
        "gênero/número, expressões usadas na legislação, leis e institutos relacionados (ex.: nomes de leis "
        "conhecidas), tipos específicos do fenômeno e assuntos diretamente conectados. Não se limite a repetir "
        f"o tema. Evite termos genéricos que tragam muito ruído.{evitar}\n"
        'Responda APENAS com um array JSON de strings, ex.: ["termo um", "termo dois"].'
    )


def _parse(texto):
    """Pega o último array JSON de strings encontrado no texto."""
    achados = re.findall(r"\[[^\[\]]*\]", texto, re.S)
    for cand in reversed(achados):
        try:
            termos = json.loads(cand)
        except Exception:
            continue
        termos = [t.strip() for t in termos if isinstance(t, str) and t.strip()]
        if termos:
            return termos
    raise ValueError("nenhuma lista de termos encontrada na resposta")


def sugerir_termos(tema, api_key, n_termos=12, ja_usados=()):
    """Via API da Anthropic (Claude)."""
    import anthropic

    client = anthropic.Anthropic(api_key=api_key)
    msg = client.messages.create(model=MODELO_IA, max_tokens=1200,
                                 messages=[{"role": "user", "content": _prompt(tema, n_termos, ja_usados)}])
    return _parse("".join(b.text for b in msg.content if getattr(b, "type", "") == "text"))


def _rodar_cli(cmd, nome, instalar):
    import shutil
    import subprocess

    if not shutil.which(cmd[0]):
        raise RuntimeError(f"o comando '{cmd[0]}' não foi encontrado ({instalar})")
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=180, stdin=subprocess.DEVNULL)
    except subprocess.TimeoutExpired:
        raise RuntimeError(f"{nome} demorou demais para responder")
    if r.returncode != 0:
        raise RuntimeError(f"{nome} retornou erro: {(r.stderr or r.stdout).strip()[:300]}")
    return r.stdout


def sugerir_termos_claude_cli(tema, n_termos=12, ja_usados=()):
    """Claude pela assinatura (Claude Code em modo não interativo) - sem chave de API."""
    saida = _rodar_cli(["claude", "-p", _prompt(tema, n_termos, ja_usados)], "Claude Code",
                       "instale o Claude Code e faça login com sua conta Claude")
    return _parse(saida)


def sugerir_termos_chatgpt_cli(tema, n_termos=12, ja_usados=()):
    """ChatGPT pela assinatura (Codex CLI em modo não interativo) - sem chave de API."""
    saida = _rodar_cli(["codex", "exec", "--skip-git-repo-check", _prompt(tema, n_termos, ja_usados)],
                       "Codex CLI", "instale o Codex CLI e faça login com sua conta ChatGPT")
    return _parse(saida)


def prompt_para_claude(tema, n_termos=20, ja_usados=()):
    """Texto para colar manualmente no Claude.ai (alternativa sem API)."""
    return _prompt(tema, n_termos, ja_usados)


def interpretar_resposta(texto):
    """Aceita a resposta do Claude em JSON, lista com marcadores/linhas ou termos separados por vírgula/OR."""
    try:
        return _parse(texto)
    except Exception:
        pass
    partes = re.split(r"\n|;|,|\s+OR\s+", texto)
    termos = [re.sub(r'^[\s\-\*\u2022\d\.\)]+|["\s]+$', "", p).strip('"').strip() for p in partes]
    return [t for t in termos if t]


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


# ---------- Alternativa gratuita: indexação real das proposições da Câmara ----------
STOP = {"contra", "sobre", "para", "com", "sem", "dos", "das", "pela", "pelo", "entre"}
API_CAMARA = "https://dadosabertos.camara.leg.br/api/v2"
GENERICOS = {
    "alteração", "alteracao", "lei federal", "legislação", "legislacao", "norma jurídica", "criação",
    "medida", "definição", "competência", "regulamentação", "regulamentacao", "inclusão", "inclusao",
    "prazo", "obrigatoriedade", "direito", "dever", "penalidade", "multa", "revogação", "revogacao",
    "requerimento", "pedido", "providência", "providencia", "projeto de lei", "poder executivo",
    "união federal", "uniao federal", "administração pública", "administracao publica", "política pública",
    "politica publica", "código penal", "codigo penal", "estatuto", "competência legislativa",
}


def _plural(palavra):
    p = palavra.lower()
    if p.endswith("ão"):
        return p[:-2] + "ões"
    if p.endswith("m"):
        return p[:-1] + "ns"
    if p.endswith("l"):
        return p[:-1] + "is"
    if p.endswith(("r", "z")):
        return p + "es"
    if p.endswith("s"):
        return p
    return p + "s"


def variantes(tema):
    """Tema e sua forma no plural (se o tema tiver uma única palavra)."""
    t = tema.strip()
    saida = [t]
    if t and " " not in t:
        pl = _plural(t)
        if pl != t.lower():
            saida.append(pl)
    return saida


def sugerir_termos_camara(tema, n_termos=12, ja_usados=(), n_proposicoes=200):
    """Busca proposições sobre o tema na API de Dados Abertos da Câmara e devolve
    os termos de indexação (keywords) mais frequentes. Gratuito, sem chave."""
    import requests
    from collections import Counter
    from concurrent.futures import ThreadPoolExecutor

    h = {"Accept": "application/json"}
    ids = []
    palavras = [p for p in re.findall(r"\w+", tema.lower()) if len(p) > 3 and p not in STOP]
    for consulta in [tema] + (palavras if len(palavras) > 1 else []):
        for pagina in range(1, (n_proposicoes + 99) // 100 + 1):
            r = requests.get(f"{API_CAMARA}/proposicoes", headers=h, timeout=30, params={
                "keywords": consulta, "itens": 100, "pagina": pagina, "ordem": "DESC", "ordenarPor": "id"})
            r.raise_for_status()
            dados = r.json().get("dados", [])
            ids += [d["id"] for d in dados]
            if len(dados) < 100:
                break
    ids = list(dict.fromkeys(ids))[:n_proposicoes]

    def kw(i):
        try:
            x = requests.get(f"{API_CAMARA}/proposicoes/{i}", headers=h, timeout=30).json()
            return [k.strip() for k in (x.get("dados", {}).get("keywords") or "").split(",") if k.strip()]
        except Exception:
            return []

    with ThreadPoolExecutor(16) as ex:
        listas = list(ex.map(kw, ids))

    excluir = {t.lower() for t in GENERICOS} | {t.lower() for t in ja_usados} | {v.lower() for v in variantes(tema)}
    cont = Counter()
    for lista in listas:
        for k in set(lista):
            if k.lower() not in excluir:
                cont[k] += 1
    # mais frequentes primeiro; se faltar, completa com os que aparecem só uma vez
    freq = [k for k, c in cont.most_common()]
    return variantes(tema) + freq[:n_termos]
