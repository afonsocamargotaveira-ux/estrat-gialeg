"""Complementa as sugestões do Tesauro com o próprio Claude, sem chave de API: usa o programa
`claude` (Claude Code) já instalado e com login feito neste computador (assinatura Claude)."""
import json
import re
import os
import shutil
import subprocess
from pathlib import Path


def encontrar(caminho_manual=""):
    """Localiza o programa 'claude': caminho informado, PATH e locais comuns de instalação (Windows/Mac/Linux)."""
    if caminho_manual and Path(caminho_manual.strip('" ')).exists():
        return caminho_manual.strip('" ')
    achado = shutil.which("claude")
    if achado:
        return achado
    # Windows: o PATH deste processo pode estar desatualizado (programa instalado depois de abrir o terminal);
    # lê o PATH atual direto do registro do Windows.
    if os.name == "nt":
        try:
            import winreg
            pastas = []
            for raiz, sub in ((winreg.HKEY_CURRENT_USER, "Environment"),
                              (winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment")):
                try:
                    with winreg.OpenKey(raiz, sub) as k:
                        pastas += os.path.expandvars(winreg.QueryValueEx(k, "Path")[0]).split(";")
                except OSError:
                    pass
            for p in pastas:
                for nome in ("claude.exe", "claude.cmd"):
                    f = Path(p.strip()) / nome
                    if p.strip() and f.is_file():
                        return str(f)
        except Exception:
            pass
    home = Path.home()
    candidatos = [
        home / ".local" / "bin" / "claude.exe", home / ".local" / "bin" / "claude",
        Path(os.environ.get("APPDATA", "")) / "npm" / "claude.cmd",
        Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft" / "WinGet" / "Links" / "claude.exe",
        Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "claude" / "claude.exe",
        Path("/usr/local/bin/claude"), Path("/opt/homebrew/bin/claude"),
    ]
    for c in candidatos:
        if str(c) and c.is_file():
            return str(c)
    return None


def disponivel(caminho_manual=""):
    return encontrar(caminho_manual) is not None


def _prompt(tema, bases, ja_sugeridos, n):
    return (
        "Você é analista legislativo da Câmara dos Deputados do Brasil e monta estratégias de busca por "
        "proposições no Sileg (pesquisa textual em ementa, explicação da ementa e indexação).\n"
        f"Tema da pesquisa: {tema}\n"
        f"Termos do Tesauro já escolhidos como base: {', '.join(bases) or '(nenhum)'}\n"
        f"Termos que já foram sugeridos a partir do Tesauro (NÃO repita): {', '.join(ja_sugeridos[:150])}\n\n"
        f"Faça uma análise semântica do tema e liste até {n} termos ou expressões curtas, em português do Brasil, "
        "que costumam aparecer em ementas e na indexação de proposições legislativas sobre esse tema e que ainda "
        "NÃO estão na lista acima: sinônimos e variações de gênero/número, expressões jurídicas e nomes de leis "
        "relacionadas, institutos, públicos e situações específicas, programas e conceitos vizinhos. "
        "Evite termos genéricos que tragam muito ruído.\n"
        'Responda APENAS com um array JSON de strings, por exemplo: ["termo um", "termo dois"].'
    )


def sugerir(tema, bases, ja_sugeridos, n=30, timeout=240, caminho=""):
    exe = encontrar(caminho)
    if not exe:
        raise RuntimeError("o programa 'claude' (Claude Code) não foi encontrado neste computador. Veja as instruções de instalação abaixo do botão")
    try:
        # remove chaves de API do ambiente: assim o Claude Code usa o login da assinatura, não crédito de API
        env = {k: v for k, v in os.environ.items() if k not in ("ANTHROPIC_API_KEY", "ANTHROPIC_AUTH_TOKEN")}
        r = subprocess.run([exe, "-p"], input=_prompt(tema, bases, ja_sugeridos, n), capture_output=True,
                           text=True, encoding="utf-8", timeout=timeout, env=env)
    except subprocess.TimeoutExpired:
        raise RuntimeError("o Claude demorou demais para responder")
    if r.returncode != 0:
        msg = (r.stderr or r.stdout or "erro desconhecido").strip()[:300]
        if "credit balance" in msg.lower():
            msg += (" — o Claude Code está logado com uma conta de API (Console) em vez da assinatura. "
                    "No PowerShell, rode `claude`, digite /logout e depois /login escolhendo "
                    "'Claude account with subscription'.")
        raise RuntimeError(msg)
    for cand in reversed(re.findall(r"\[[^\[\]]*\]", r.stdout, re.S)):
        try:
            termos = json.loads(cand)
        except Exception:
            continue
        termos = [t.strip() for t in termos if isinstance(t, str) and t.strip()]
        if termos:
            return termos
    raise RuntimeError("não encontrei uma lista de termos na resposta do Claude")
