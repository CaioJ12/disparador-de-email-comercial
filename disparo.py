#!/usr/bin/env python3
"""
Disparo de e-mails personalizados com TRAVA de revisão manual.

Fluxo (3 etapas):
  1) python disparo.py preparar Empresas_Prospectadas.xlsx
       -> lê a planilha original, extrai TODOS os e-mails (mesmo misturados com
          telefone, separados por ';' etc.) e gera  revisao_envio.xlsx  com uma
          sugestão de tratamento (Grupo / Pessoal), nome e saudação.

  2) Revisão MANUAL (a trava). Duas opções:
       a) Abrir revisao_envio.xlsx no Excel e preencher/corrigir as colunas
          TIPO, NOME, SAUDACAO, FRASE_PERSONALIZADA e marcar APROVADO = SIM; ou
       b) python disparo.py revisar      (modo interativo no terminal, linha a linha)

     Enquanto houver linha com ENVIAR=SIM sem APROVADO=SIM ou sem SAUDACAO,
     o envio é BLOQUEADO.

  3) python disparo.py enviar            (simulação: mostra os e-mails, não envia)
     python disparo.py enviar --real     (envia de verdade, 8–10 s entre cada)

Só usa a biblioteca padrão + openpyxl  (pip install openpyxl).
"""

import argparse
import csv
import getpass
import os
import random
import re
import smtplib
import ssl
import sys
import time
import unicodedata
from datetime import datetime
from email.message import EmailMessage
from email.utils import formataddr, make_msgid
from html import escape
from pathlib import Path

try:
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.worksheet.datavalidation import DataValidation
except ImportError:
    sys.exit("Instale a dependência:  pip install openpyxl")

# ╔══════════════════════════════════════════════════════════════════════════╗
# ║                    ✏️  EDITE AQUI O TEXTO DO E-MAIL                       ║
# ║  Abra este arquivo no Bloco de Notas (ou VS Code), altere SÓ o que está   ║
# ║  entre as aspas triplas (três aspas duplas) e salve.                     ║
# ║                                                                          ║
# ║  Campos automáticos (trocados para cada contato):                        ║
# ║    {saudacao}            "Olá, Maria,"    /  "Olá, equipe de Compras..." ║
# ║    {nome} {empresa_curta} {a_empresa} {da_empresa} {frase_personalizada}  ║
# ║    {voces} -> você/vocês   {podem} -> pode/podem                         ║
# ║    {bloco_catalogo}      -> quadro do catálogo (ver mais abaixo)          ║
# ║    {botao_whatsapp}      -> botão verde do WhatsApp                       ║
# ║  Formatação:  **negrito**   •  linhas começando com "- " viram lista ✓    ║
# ╚══════════════════════════════════════════════════════════════════════════╝

ASSUNTO_PADRAO = "Catálogo de produtos [Sua Empresa] – {empresa_curta}"

TEXTO_EMAIL = """
{saudacao}

Meu nome é **[Seu nome]**, sou [seu cargo] da **[Sua Empresa]**, empresa de [sua cidade – UF] que atende [sua região ou segmento].

{frase_personalizada}

Gostaria de apresentar o nosso **catálogo de produtos** e me colocar à disposição como fornecedor {da_empresa}. Nele estão as principais linhas que oferecemos:

- **[Linha de produto 1]:** [breve descrição]
- **[Linha de produto 2]:** [breve descrição]
- **[Linha de produto 3]:** [breve descrição]

Trabalhamos com atendimento próximo e retorno rápido de cotações.

{bloco_catalogo}

Se fizer sentido, {voces} {podem} me enviar uma lista dos itens que costumam comprar para que eu prepare uma cotação — ou, se preferir, é só me chamar no WhatsApp:

{botao_whatsapp}

Fico à disposição e agradeço a atenção.

Atenciosamente,
"""

# ---- Quadro do CATÁLOGO (aparece onde estiver {bloco_catalogo}) -------------
CATALOGO_TITULO = "Catálogo [Sua Empresa]"
CATALOGO_TEXTO = "com nossas linhas de produtos, medidas e especificações."
# Destaques opcionais numa linha pequena, ex.: ["Rolamentos", "EPIs"]  (deixe [] para não mostrar)
CATALOGO_DESTAQUES = []
# Link público do PDF (OneDrive, site...). Com ele, a faixa fica clicável e ganha o botão de download.
# Deixe "" se não tiver: um e-mail não consegue apontar para o próprio anexo.
CATALOGO_LINK = ""
CATALOGO_BOTAO = "⬇ Baixar PDF"

# ---- Botão do WHATSAPP ------------------------------------------------------
WHATSAPP_BOTAO = "Falar no WhatsApp"
# Mensagem que já aparece digitada quando o cliente clica no botão:
WHATSAPP_MENSAGEM = "Olá, {remetente}! Recebi o seu e-mail ({empresa_curta}) e gostaria de mais informações."
# OBS.: se você editar pelo PAINEL (navegador), os textos ficam salvos em textos.json
#       e passam a valer no lugar destes. Para voltar a usar estes, apague textos.json.
# ═══════════════════════════ FIM DA ÁREA DE EDIÇÃO ═══════════════════════════


# Windows: garante acentos/emojis no terminal
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:  # noqa: BLE001
        pass

PASTA = Path(__file__).resolve().parent
TEXTOS_ARQ = PASTA / "textos.json"   # textos salvos pelo PAINEL (têm prioridade sobre a área acima)
CHAVES_TEXTO = ["ASSUNTO_PADRAO", "TEXTO_EMAIL", "CATALOGO_TITULO", "CATALOGO_TEXTO", "CATALOGO_DESTAQUES",
                "CATALOGO_LINK", "CATALOGO_BOTAO", "WHATSAPP_BOTAO", "WHATSAPP_MENSAGEM"]


def textos_atuais() -> dict:
    g = globals()
    return {k: (g[k].strip("\n") if isinstance(g[k], str) else g[k]) for k in CHAVES_TEXTO}


def aplicar_textos(d: dict):
    g = globals()
    for k in CHAVES_TEXTO:
        if k in d:
            v = d[k]
            if k == "CATALOGO_DESTAQUES" and isinstance(v, str):
                v = [x.strip() for x in v.split(",") if x.strip()]
            g[k] = v


def carregar_textos_salvos():
    if TEXTOS_ARQ.exists():
        import json
        try:
            aplicar_textos(json.loads(TEXTOS_ARQ.read_text(encoding="utf-8")))
        except Exception as e:  # noqa: BLE001
            print(f"(aviso: textos.json inválido, usando os textos do código: {e})")


def salvar_textos(d: dict):
    import json
    aplicar_textos(d)
    TEXTOS_ARQ.write_text(json.dumps(textos_atuais(), ensure_ascii=False, indent=2), encoding="utf-8")


REVISAO = PASTA / "revisao_envio.xlsx"
LOG = PASTA / "log_envios.csv"
CONFIG = PASTA / "config.txt"
CONFIG_EXEMPLO = PASTA / "config.exemplo.txt"

COLUNAS = [
    "ENVIAR",               # SIM / NAO
    "APROVADO",             # SIM só depois de revisar (TRAVA)
    "EMPRESA",
    "EMAIL",
    "TELEFONE",             # telefones da empresa achados na planilha (só consulta)
    "TIPO",                 # GRUPO / PESSOAL
    "NOME",                 # nome da pessoa (PESSOAL) ou do setor/equipe (GRUPO)
    "SAUDACAO",             # linha de abertura exata, ex.: "Olá, Maria,"
    "FRASE_PERSONALIZADA",  # opcional: frase específica p/ esta empresa
    "ASSUNTO",              # opcional: sobrescreve o assunto padrão
    "OBS_ORIGINAL",         # contexto vindo da planilha original
    "STATUS",               # preenchido pelo script (ENVIADO / ERRO)
]

EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")

# Prefixos que indicam caixa de setor (=> GRUPO)
GENERICOS = {
    "compras", "compra", "suprimentos", "sac", "ouvidoria", "contato", "comercial",
    "vendas", "financeiro", "fornecedor", "fornecedores", "central", "atendimento",
    "rh", "faleconosco", "info", "adm", "administrativo", "logistica", "cadastro",
    "suporte", "diretoria", "juridico", "marketing", "procurement", "purchasing",
}
SETOR_LEGIVEL = {
    "compras": "Compras", "compra": "Compras", "suprimentos": "Suprimentos",
    "sac": "Atendimento", "ouvidoria": "Ouvidoria", "fornecedor": "Fornecedores",
    "fornecedores": "Fornecedores", "comercial": "Comercial", "vendas": "Vendas",
    "financeiro": "Financeiro", "logistica": "Logística", "cadastro": "Cadastro",
}


# ----------------------------------------------------------------------------
# utilidades
# ----------------------------------------------------------------------------
def sem_acento(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")


def limpa_empresa(nome: str) -> str:
    nome = re.sub(r"\s+", " ", str(nome or "")).strip()
    return nome


def empresa_curta(nome: str) -> str:
    """'Empresa Exemplo Indústria S.A.' -> 'Empresa Exemplo Indústria'"""
    n = re.sub(r"\b(S/?A|S\.A\.?|LTDA\.?|EIRELI|ME)\b\.?", "", nome, flags=re.I)
    n = re.sub(r"\s*[-–(].*$", "", n)  # corta ' - Unidade 1', '(fábrica)'
    n = re.sub(r"\s+", " ", n).strip(" ,.-") or nome
    if n.isupper() or n.islower():
        n = n.title()
    return n


def da_do(emp: str) -> str:
    """'do Grupo Exemplo' / 'da Empresa Exemplo'"""
    masc = ("grupo ", "estaleiro ", "instituto ", "laboratório ", "moinho ", "frigorífico ")
    return "do" if emp.lower().startswith(masc) else "da"


def sugerir(email: str, empresa: str):
    """Retorna (tipo, nome, saudacao) sugeridos — SEMPRE revisar!"""
    local = email.split("@")[0].lower()
    base = re.sub(r"\d+$", "", re.split(r"[._\-+]", local)[0])
    partes = [p for p in re.split(r"[._\-+]", local) if p and not p.isdigit()]
    emp = empresa_curta(empresa)

    dd = da_do(emp)
    limpos = [re.sub(r"\d+$", "", sem_acento(x)) for x in partes]
    if any(x in GENERICOS for x in limpos):
        setor = next((SETOR_LEGIVEL[x] for x in limpos if x in SETOR_LEGIVEL), "")
        nome = f"Equipe de {setor}" if setor else "Equipe"
        saud = f"Olá, equipe de {setor} {dd} {emp}," if setor else f"Olá, equipe {dd} {emp},"
        return "GRUPO", nome, saud

    if len(partes) >= 1 and partes[0].isalpha() and len(partes[0]) > 2:
        primeiro = partes[0].capitalize()
        return "PESSOAL", primeiro, f"Olá, {primeiro},"

    return "GRUPO", "Equipe", f"Olá, equipe {dd} {emp},"


def sim(v) -> bool:
    return sem_acento(str(v or "")).strip().upper() in {"SIM", "S", "X", "OK", "YES", "Y", "1"}


def ler_config():
    cfg = {}
    if not CONFIG.exists() and CONFIG_EXEMPLO.exists():
        CONFIG.write_text(CONFIG_EXEMPLO.read_text(encoding="utf-8"), encoding="utf-8")
    if CONFIG.exists():
        for linha in CONFIG.read_text(encoding="utf-8").splitlines():
            linha = linha.strip()
            if linha and not linha.startswith("#") and "=" in linha:
                k, v = linha.split("=", 1)
                cfg[k.strip().upper()] = v.strip()
    return cfg


def lista_emails(valor) -> list:
    """Separa e-mails escritos com vírgula, ponto e vírgula ou espaço, sem repetidos."""
    vistos, saida = set(), []
    for e in re.split(r"[,;\s]+", str(valor or "")):
        if e and e.lower() not in vistos:
            vistos.add(e.lower())
            saida.append(e)
    return saida


def copias_envio(cfg, para="") -> tuple:
    """-> (cc, cco) sem repetir o destinatário nem quem já está em cópia."""
    para = para.lower()
    cc = [e for e in lista_emails(cfg.get("EMAILS_COPIA")) if e.lower() != para]
    ja = {para, *(e.lower() for e in cc)}
    cco = [e for e in lista_emails(cfg.get("EMAILS_COPIA_OCULTA")) if e.lower() not in ja]
    return cc, cco


def link_download(url) -> str:
    """No OneDrive empresarial (SharePoint), download=1 baixa o PDF em vez de abrir a visualização."""
    from urllib.parse import urlsplit
    url = (url or "").strip()
    host = (urlsplit(url).hostname or "").lower()
    if (host == "sharepoint.com" or host.endswith(".sharepoint.com")) and "download=1" not in url:
        return url + ("&" if "?" in url else "?") + "download=1"
    return url


def ler_modelo():
    return ASSUNTO_PADRAO, TEXTO_EMAIL.strip("\n")


# ----------------------------------------------------------------------------
# telefones / WhatsApp
# ----------------------------------------------------------------------------
TEL_RE = re.compile(r"(?:\+?55[\s-]*)?(?:0?800[\s.-]*\d{3}[\s.-]*\d{4}|\(?\d{2}\)?[\s.-]*9?[\s.-]*\d{4}[\s.-]*\d{4})")


def classificar_telefone(bruto: str):
    """-> (formatado, tipo, numero_whatsapp ou '')"""
    d = re.sub(r"\D", "", bruto)
    if d.startswith("55") and len(d) in (12, 13):
        d = d[2:]
    if d.startswith(("0800", "800")):
        d = d.lstrip("0")
        return f"0800 {d[3:6]} {d[6:]}", "0800", ""
    if len(d) == 11 and d[2] == "9":
        return f"({d[:2]}) {d[2:7]}-{d[7:]}", "Celular", "55" + d
    if len(d) == 10:
        ddd, loc = d[:2], d[2:]
        if loc[0] in "6789":  # celular antigo, sem o 9 na frente
            return f"({ddd}) {loc[:4]}-{loc[4:]}", "Celular (sem o 9?)", f"55{ddd}9{loc}"
        return f"({ddd}) {loc[:4]}-{loc[4:]}", "Fixo", ""
    return bruto.strip(), "?", ""


def link_whatsapp(numero: str, texto: str = "") -> str:
    from urllib.parse import quote
    n = re.sub(r"\D", "", numero)
    if n and not n.startswith("55"):
        n = "55" + n
    if len(n) == 12 and n[4] in "6789":  # celular sem o 9
        n = n[:4] + "9" + n[4:]
    return f"https://wa.me/{n}" + (f"?text={quote(texto)}" if texto else "")


def nome_de_linkedin(url: str) -> str:
    slug = url.rstrip("/").split("/in/")[-1]
    partes = [p for p in slug.split("-") if p.isalpha()]
    return " ".join(p.capitalize() for p in partes[:3])


def ja_enviados():
    if not LOG.exists():
        return set()
    with LOG.open(encoding="utf-8") as f:
        return {r["email"].lower() for r in csv.DictReader(f) if r.get("status") == "ENVIADO"}


def registrar(email, empresa, status, detalhe=""):
    novo = not LOG.exists()
    with LOG.open("a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if novo:
            w.writerow(["data_hora", "email", "empresa", "status", "detalhe"])
        w.writerow([datetime.now().isoformat(timespec="seconds"), email, empresa, status, detalhe])


# ----------------------------------------------------------------------------
# planilha de revisão
# ----------------------------------------------------------------------------
def carregar_revisao():
    if not REVISAO.exists():
        sys.exit("revisao_envio.xlsx não existe. Rode primeiro:  python disparo.py preparar <planilha.xlsx>")
    wb = openpyxl.load_workbook(REVISAO)
    ws = wb.active
    cab = [str(c.value or "").strip().upper() for c in ws[1]]
    falt = [c for c in COLUNAS if c not in cab]
    if falt:
        sys.exit(f"Colunas faltando em revisao_envio.xlsx: {falt}")
    idx = {c: cab.index(c) for c in COLUNAS}
    linhas = []
    for row in ws.iter_rows(min_row=2):
        d = {c: (row[idx[c]].value if row[idx[c]].value is not None else "") for c in COLUNAS}
        if not str(d["EMAIL"]).strip():
            continue
        d = {k: str(v).strip() for k, v in d.items()}
        d["_row"] = row[0].row
        linhas.append(d)
    return wb, ws, idx, linhas


def salvar_celula(ws, idx, row, coluna, valor):
    ws.cell(row=row, column=idx[coluna] + 1, value=valor)


STATUS_PALAVRAS = ("contat", "respond", "resposta", "portaria", "conseguimos")


def ler_origem(origem):
    """Lê a planilha original e devolve 1 registro por empresa com tudo que foi identificado."""
    wb = openpyxl.load_workbook(origem, data_only=True)
    ws = wb.active
    cab_row, cols = None, {}
    for r in ws.iter_rows(min_row=1, max_row=15):
        for c in r:
            v = sem_acento(str(c.value or "")).upper().strip()
            if "RAZAO SOCIAL" in v or v == "EMPRESA":
                cab_row = c.row
        if cab_row:
            for c in r:
                v = sem_acento(str(c.value or "")).upper().strip()
                if v:
                    cols[v] = c.column - 1
            break
    if not cab_row:
        sys.exit("Não achei a coluna 'RAZÃO SOCIAL' / 'EMPRESA' na planilha.")
    col_emp = next(i for k, i in cols.items() if "RAZAO SOCIAL" in k or k == "EMPRESA")
    col_resp = next((i for k, i in cols.items() if "RESPOSTA" in k), None)

    regs = []
    for row in ws.iter_rows(min_row=cab_row + 1, values_only=True):
        empresa = limpa_empresa(row[col_emp] if col_emp < len(row) else "")
        if not empresa:
            continue
        r = {"empresa": empresa, "emails": [], "telefones": [], "nomes": [], "linkedin": [], "situacao": []}
        for i, v in enumerate(row):
            if v is None or i == col_emp:
                continue
            txt = str(v)
            # status / situação (colunas de anotação)
            if any(p in sem_acento(txt).lower() for p in STATUS_PALAVRAS) and not EMAIL_RE.search(txt) \
                    and not TEL_RE.search(txt):
                r["situacao"].append(txt.strip())
                continue
            if i == col_resp:
                r["situacao"].append(txt.strip())
                continue
            for e in EMAIL_RE.findall(txt):
                e = e.strip(".").lower()
                if e not in r["emails"]:
                    r["emails"].append(e)
            for u in re.findall(r"https?://\S*linkedin\.com/in/[^\s;,]+", txt):
                u = u.rstrip(";,")
                r["linkedin"].append(u)
                n = nome_de_linkedin(u)
                if n and n not in r["nomes"]:
                    r["nomes"].append(n + " (LinkedIn)")
            resto = re.sub(r"https?://\S+", " ", EMAIL_RE.sub(" ", txt))
            for t in TEL_RE.findall(resto):
                if len(re.sub(r"\D", "", t)) < 10:
                    continue
                fmt, tipo, wa = classificar_telefone(t)
                obs = " (portaria)" if "portaria" in txt.lower() else ""
                if fmt not in [x[0] for x in r["telefones"]]:
                    r["telefones"].append((fmt, tipo + obs, wa))
            # nome solto (ex.: "Maria")
            sobra = re.sub(r"\(.*?\)", " ", TEL_RE.sub(" ", resto))
            for pedaco in re.split(r"[;/|,\n]", sobra):
                pedaco = pedaco.strip(" -_.")
                if pedaco and re.fullmatch(r"[A-Za-zÀ-ÿ' ]{3,40}", pedaco) and len(pedaco.split()) <= 4:
                    if pedaco.title() not in r["nomes"]:
                        r["nomes"].append(pedaco.title())
        # nomes vindos de e-mails pessoais
        for e in r["emails"]:
            tipo, nome, _ = sugerir(e, empresa)
            if tipo == "PESSOAL":
                completo = " ".join(p.capitalize() for p in re.split(r"[._\-]", e.split("@")[0]) if p.isalpha())
                if completo not in r["nomes"]:
                    r["nomes"].append(completo)
        regs.append(r)
    return regs


def escrever_aba_contatos(wb, regs):
    if "Contatos_Telefones" in wb.sheetnames:
        del wb["Contatos_Telefones"]
    if "Sem_email" in wb.sheetnames:
        del wb["Sem_email"]
    s = wb.create_sheet("Contatos_Telefones")
    cab = ["EMPRESA", "NOMES IDENTIFICADOS", "E-MAILS", "TELEFONE", "TIPO", "WHATSAPP", "LINKEDIN", "SITUAÇÃO"]
    s.append(cab)
    head = PatternFill("solid", fgColor="1F4E78")
    for c in s[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = head
        c.alignment = Alignment(horizontal="center", vertical="center")
    sem_mail = PatternFill("solid", fgColor="FCE4E4")
    zebra = PatternFill("solid", fgColor="F3F6FA")
    for n_emp, r in enumerate(regs):
        tels = r["telefones"] or [("", "", "")]
        for k, (fmt, tipo, wa) in enumerate(tels):
            s.append([
                r["empresa"] if k == 0 else "",
                "\n".join(r["nomes"]) if k == 0 else "",
                "\n".join(r["emails"]) if k == 0 else "",
                fmt, tipo,
                "Abrir WhatsApp" if wa else "",
                "\n".join(r["linkedin"]) if k == 0 else "",
                " | ".join(r["situacao"]) if k == 0 else "",
            ])
            linha = s.max_row
            if wa:
                c = s.cell(row=linha, column=6)
                c.hyperlink = link_whatsapp(wa)
                c.font = Font(color="128C4A", underline="single", bold=True)
            for col in range(1, len(cab) + 1):
                cel = s.cell(row=linha, column=col)
                cel.alignment = Alignment(vertical="top", wrap_text=True)
                if not r["emails"]:
                    cel.fill = sem_mail
                elif n_emp % 2:
                    cel.fill = zebra
    for col, w in zip("ABCDEFGH", [34, 26, 34, 18, 20, 16, 40, 34]):
        s.column_dimensions[col].width = w
    s.freeze_panes = "B2"
    s.auto_filter.ref = s.dimensions
    s.cell(row=s.max_row + 2, column=1,
           value="Linhas em vermelho = empresa sem e-mail na planilha (só telefone/LinkedIn).").font = Font(italic=True, color="999999")


def cmd_contatos(args):
    regs = ler_origem(Path(args.planilha))
    # terminal
    print(f"{'EMPRESA':32} {'NOMES':28} {'TELEFONE':18} TIPO")
    print("─" * 100)
    for r in regs:
        tels = r["telefones"] or [("—", "", "")]
        for k, (fmt, tipo, wa) in enumerate(tels):
            print(f"{(r['empresa'] if k == 0 else '')[:31]:32} "
                  f"{(', '.join(r['nomes']) if k == 0 else '')[:27]:28} {fmt:18} {tipo}{'  [WhatsApp]' if wa else ''}")
    if REVISAO.exists():
        wb = openpyxl.load_workbook(REVISAO)
        escrever_aba_contatos(wb, regs)
        try:
            wb.save(REVISAO)
            print(f"\n✔ Aba 'Contatos_Telefones' atualizada em {REVISAO.name} (sua revisão foi mantida).")
        except PermissionError:
            print("\n(feche o Excel para atualizar a aba Contatos_Telefones)")


def cmd_preparar(args):
    origem = Path(args.planilha)
    if REVISAO.exists() and not args.forcar:
        sys.exit(f"{REVISAO.name} já existe (pode conter sua revisão!). Use --forcar para sobrescrever.")
    regs = ler_origem(origem)

    saida, vistos = [], set()
    for r in regs:
        tels = "  /  ".join(f"{t[0]} ({t[1]})" for t in r["telefones"])
        situ = " | ".join(r["situacao"])
        for e in r["emails"]:
            if e in vistos:
                continue
            vistos.add(e)
            tipo, nome, saud = sugerir(e, r["empresa"])
            saida.append({
                "ENVIAR": "NAO" if "respondido" in situ.lower() else "SIM",
                "APROVADO": "",
                "EMPRESA": r["empresa"],
                "EMAIL": e,
                "TELEFONE": tels,
                "TIPO": tipo,
                "NOME": nome,
                "SAUDACAO": saud,
                "FRASE_PERSONALIZADA": "",
                "ASSUNTO": "",
                "OBS_ORIGINAL": situ[:300],
                "STATUS": "",
            })
    sem_email = [r for r in regs if not r["emails"]]

    out = openpyxl.Workbook()
    o = out.active
    o.title = "Revisao"
    o.append(COLUNAS)
    for d in saida:
        o.append([d[c] for c in COLUNAS])

    head = PatternFill("solid", fgColor="1F4E78")
    trava = PatternFill("solid", fgColor="FFF2CC")
    for c in o[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = head
        c.alignment = Alignment(horizontal="center")
    larg = {"ENVIAR": 9, "APROVADO": 11, "EMPRESA": 34, "EMAIL": 36, "TELEFONE": 28, "TIPO": 11, "NOME": 20,
            "SAUDACAO": 44, "FRASE_PERSONALIZADA": 50, "ASSUNTO": 30, "OBS_ORIGINAL": 34, "STATUS": 22}
    for i, c in enumerate(COLUNAS, 1):
        o.column_dimensions[openpyxl.utils.get_column_letter(i)].width = larg[c]
    for r in range(2, o.max_row + 1):
        o.cell(row=r, column=COLUNAS.index("APROVADO") + 1).fill = trava
    o.freeze_panes = "C2"
    o.auto_filter.ref = o.dimensions

    n = max(o.max_row, 2)
    L = lambda c: openpyxl.utils.get_column_letter(COLUNAS.index(c) + 1)
    dv_sn = DataValidation(type="list", formula1='"SIM,NAO"', allow_blank=True)
    dv_tp = DataValidation(type="list", formula1='"GRUPO,PESSOAL"', allow_blank=True)
    o.add_data_validation(dv_sn)
    o.add_data_validation(dv_tp)
    dv_sn.add(f"{L('ENVIAR')}2:{L('ENVIAR')}{n + 200}")
    dv_sn.add(f"{L('APROVADO')}2:{L('APROVADO')}{n + 200}")
    dv_tp.add(f"{L('TIPO')}2:{L('TIPO')}{n + 200}")

    escrever_aba_contatos(out, regs)
    out.save(REVISAO)
    n_tel = sum(len(r["telefones"]) for r in regs)
    n_wa = sum(1 for r in regs for t in r["telefones"] if t[2])
    print(f"✔ {len(saida)} e-mails extraídos  →  {REVISAO}")
    print(f"✔ {n_tel} telefones identificados ({n_wa} com WhatsApp provável) → aba 'Contatos_Telefones'")
    print(f"  {len(sem_email)} empresas sem e-mail (em vermelho na aba de contatos).")
    print("\nPRÓXIMO PASSO (trava): revise cada linha e marque APROVADO = SIM.")
    print("  • no Excel, ou")
    print("  • python disparo.py revisar")


# ----------------------------------------------------------------------------
# revisão interativa
# ----------------------------------------------------------------------------
def pergunta(rotulo, atual):
    r = input(f"  {rotulo} [{atual}]: ").strip()
    return atual if r == "" else r


def cmd_revisar(args):
    wb, ws, idx, linhas = carregar_revisao()
    pend = [l for l in linhas if sim(l["ENVIAR"]) and (args.todas or not sim(l["APROVADO"]))]
    if not pend:
        print("Nada pendente de revisão. (use --todas para revisar de novo)")
        return
    print(f"{len(pend)} linha(s) para revisar. ENTER mantém o valor; 'p' pula; 'n' = não enviar; 'q' salva e sai.\n")
    for i, l in enumerate(pend, 1):
        print("─" * 70)
        print(f"[{i}/{len(pend)}]  {l['EMPRESA']}  <{l['EMAIL']}>")
        if l["OBS_ORIGINAL"]:
            print(f"  obs: {l['OBS_ORIGINAL']}")
        acao = input("  Revisar esta? (ENTER=sim / p=pular / n=não enviar / q=sair): ").strip().lower()
        if acao == "q":
            break
        if acao == "p":
            continue
        if acao == "n":
            salvar_celula(ws, idx, l["_row"], "ENVIAR", "NAO")
            continue

        tipo = pergunta("Tipo (G=grupo / P=pessoal)", l["TIPO"]).upper()
        tipo = {"G": "GRUPO", "P": "PESSOAL"}.get(tipo, tipo)
        nome = pergunta("Nome (pessoa ou setor)", l["NOME"])
        emp = empresa_curta(l["EMPRESA"])
        if tipo == "PESSOAL":
            sug = f"Olá, {nome},"
        else:
            eq = nome[0].lower() + nome[1:] if nome.lower().startswith("equipe") else "equipe de " + nome
            sug = f"Olá, {eq} {da_do(emp)} {emp},"
        base_saud = l["SAUDACAO"] if (tipo == l["TIPO"] and nome == l["NOME"]) else sug
        saud = pergunta("Saudação (linha de abertura)", base_saud)
        frase = pergunta("Frase personalizada (opcional)", l["FRASE_PERSONALIZADA"])
        assunto = pergunta("Assunto específico (opcional, ENTER=padrão)", l["ASSUNTO"])

        for col, val in [("TIPO", tipo), ("NOME", nome), ("SAUDACAO", saud),
                         ("FRASE_PERSONALIZADA", frase), ("ASSUNTO", assunto)]:
            salvar_celula(ws, idx, l["_row"], col, val)
        if input("  Aprovar esta linha para envio? (s/N): ").strip().lower() == "s":
            salvar_celula(ws, idx, l["_row"], "APROVADO", "SIM")
        wb.save(REVISAO)  # salva a cada linha — não perde nada se fechar
    wb.save(REVISAO)
    print(f"\n✔ Revisão salva em {REVISAO.name}")


# ----------------------------------------------------------------------------
# envio
# ----------------------------------------------------------------------------
class Vazio(dict):
    def __missing__(self, k):
        return ""


LOGO = PASTA / "logo.png"
AZUL = "#0D4DA1"


def _inline(t: str) -> str:
    """escapa HTML e converte **negrito**"""
    return re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", escape(t))


def bloco_catalogo_html(cfg) -> str:
    """Faixa discreta do catálogo (ícone de anexo + 1 linha de texto)."""
    anexo = arquivo_catalogo(cfg)
    tags = " · ".join(escape(t) for t in CATALOGO_DESTAQUES)
    detalhe = ""
    if anexo:
        tam = anexo.stat().st_size / 1024 / 1024
        detalhe = f"{escape(anexo.name)} · {tam:.1f} MB"
    url = escape(link_download(CATALOGO_LINK))
    titulo = escape(CATALOGO_TITULO)
    botao = ""
    if url:
        titulo = f'<a href="{url}" target="_blank" style="color:{AZUL};text-decoration:underline">{titulo}</a>'
        botao = (f'<td valign="middle" align="right" style="padding:10px 14px 10px 0;white-space:nowrap">'
                 f'<a href="{url}" target="_blank" style="display:inline-block;background:{AZUL};color:#ffffff;'
                 f'font-family:Arial,Helvetica,sans-serif;font-size:13px;font-weight:bold;padding:8px 12px;'
                 f'border-radius:4px;text-decoration:none;white-space:nowrap">{escape(CATALOGO_BOTAO)}</a></td>')
    return f"""
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin:0 0 18px;background:#F5F8FC;border:1px solid #E1E8F2;border-radius:6px">
  <tr>
    <td valign="middle" style="padding:12px 0 12px 14px;width:26px;font-size:18px;line-height:1">&#128206;</td>
    <td valign="middle" style="padding:10px 14px 10px 8px;font-family:Arial,Helvetica,sans-serif;font-size:14px;line-height:1.45;color:#333">
      <strong style="color:{AZUL}">{titulo}</strong> — {escape(CATALOGO_TEXTO)}
      {f'<div style="font-size:12px;color:#888;margin-top:2px">{tags}</div>' if tags else ''}
      {f'<div style="font-size:12px;color:#888;margin-top:2px">{detalhe}</div>' if detalhe else ''}
    </td>{botao}
  </tr>
</table>"""


def estilo_simples(cfg) -> bool:
    """DESIGN (padrão) = visual com faixas e botões. SIMPLES = parece um e-mail escrito à mão."""
    return (cfg.get("ESTILO_EMAIL") or "DESIGN").strip().upper() == "SIMPLES"


def bloco_catalogo_simples(cfg) -> str:
    url = escape(link_download(CATALOGO_LINK))
    link = f' <a href="{url}">{escape(CATALOGO_BOTAO)}</a>' if url else ""
    anexo = " (arquivo em anexo)" if arquivo_catalogo(cfg) else ""
    return f"<p style='margin:0 0 16px'>{escape(CATALOGO_TITULO)} — {escape(CATALOGO_TEXTO)}{anexo}{link}</p>"


def whatsapp_simples(cfg) -> str:
    num = cfg.get("WHATSAPP") or cfg.get("TELEFONE", "")
    if not num:
        return ""
    return f"<p style='margin:0 0 16px'>WhatsApp: <a href=\"{escape(link_whatsapp(num))}\">{escape(num)}</a></p>"


def layout_simples(conteudo, cfg, logo_src, assunto):
    partes = [escape(cfg.get("CARGO", "")), escape(cfg.get("NOME_EMPRESA", ""))]
    linhas = [f"<strong>{escape(cfg.get('NOME_REMETENTE', ''))}</strong>", " · ".join(p for p in partes if p),
              escape(cfg.get("TELEFONE", "")), escape(cfg.get("EMAIL_ASSINATURA") or cfg.get("EMAIL_REMETENTE", "")),
              escape(cfg.get("ENDERECO", ""))]
    rodape = "".join(f"<p style='margin:0 0 8px;font-size:12px;color:#666'>{escape(cfg[k])}</p>"
                     for k in ("TEXTO_DESCADASTRO", "AVISO_LEGAL") if cfg.get(k))
    logo = (f'<p style="margin:0 0 16px"><img src="{logo_src}" width="120" alt="{escape(cfg.get("NOME_EMPRESA", ""))}" '
            'style="border:0;width:120px;height:auto"></p>') if logo_src else ""
    return f"""<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8"><title>{escape(assunto)}</title></head>
<body style="font-family:Arial,Helvetica,sans-serif;font-size:15px;line-height:1.5;color:#222">
{conteudo}
<p style="margin:0 0 10px">{'<br>'.join(x for x in linhas if x)}</p>
{logo}
{rodape}
</body></html>"""


def botao_whatsapp_html(cfg, campos) -> str:
    num = cfg.get("WHATSAPP") or cfg.get("TELEFONE", "")
    if not num:
        return ""
    url = link_whatsapp(num, WHATSAPP_MENSAGEM.format_map(campos))
    return f"""
<table role="presentation" cellpadding="0" cellspacing="0" style="margin:4px 0 22px"><tr>
  <td style="border-radius:28px;background:#25D366">
    <a href="{escape(url)}" target="_blank" style="display:inline-block;padding:10px 22px;font-family:Arial,Helvetica,sans-serif;font-size:14px;font-weight:bold;color:#ffffff;text-decoration:none;border-radius:28px">
      &#128172;&nbsp; {escape(WHATSAPP_BOTAO)}
    </a>
  </td>
  <td style="padding-left:12px;font-family:Arial,sans-serif;font-size:13px;color:#666">{escape(cfg.get("WHATSAPP") or cfg.get("TELEFONE", ""))}</td>
</tr></table>"""


def arquivo_catalogo(cfg):
    nome = cfg.get("ANEXO_CATALOGO", "").strip()
    if not nome:
        return None
    f = Path(nome) if Path(nome).is_absolute() else PASTA / nome
    return f if f.exists() else None


def corpo_html(texto: str, especiais=None, simples=False) -> str:
    especiais = especiais or {}
    blocos = []
    for p in texto.split("\n\n"):
        if p.strip() in especiais:
            blocos.append(especiais[p.strip()])
            continue
        linhas = p.split("\n")
        if all(x.strip().startswith(("- ", "• ")) for x in linhas):  # lista
            if simples:
                itens = "".join(f"<li>{_inline(x.strip()[2:])}</li>" for x in linhas)
                blocos.append(f"<ul style='margin:0 0 16px;padding-left:22px'>{itens}</ul>")
                continue
            itens = "".join(
                f"<tr><td valign='top' style='padding:0 8px 6px 0;color:{AZUL};font-weight:bold'>&#10003;</td>"
                f"<td style='padding:0 0 6px 0'>{_inline(x.strip()[2:])}</td></tr>" for x in linhas)
            blocos.append(f"<table role='presentation' cellpadding='0' cellspacing='0' style='margin:0 0 16px'>{itens}</table>")
        else:
            blocos.append(f"<p style='margin:0 0 16px'>{'<br>'.join(_inline(x) for x in linhas)}</p>")
    return "".join(blocos)


def assinatura_html(cfg, logo_src):
    nome = escape(cfg.get("NOME_REMETENTE", ""))
    cargo = escape(cfg.get("CARGO", ""))
    tel = cfg.get("TELEFONE", "")
    email = cfg.get("EMAIL_ASSINATURA") or cfg.get("EMAIL_REMETENTE", "")
    end = escape(cfg.get("ENDERECO", ""))
    tel_link = re.sub(r"[^\d+]", "", tel)
    rot = f"color:{AZUL};font-weight:bold;width:22px;vertical-align:top"
    linha = lambda r, v: f"<tr><td style='{rot}'>{r}</td><td style='color:#333'>{v}</td></tr>" if v else ""
    logo = (f'<td valign="middle" style="padding-right:14px;border-right:3px solid {AZUL}">'
            f'<img src="{logo_src}" width="120" alt="{escape(cfg.get("NOME_EMPRESA", ""))}" '
            'style="display:block;border:0;width:120px;height:auto"></td>') if logo_src else ""
    return f"""
<table role="presentation" cellpadding="0" cellspacing="0" style="font-family:Arial,Helvetica,sans-serif;font-size:13px;line-height:1.5">
  <tr>
    {logo}
    <td valign="middle" style="padding-left:{14 if logo else 0}px">
      <div style="font-size:16px;font-weight:bold;color:{AZUL}">{nome}</div>
      <div style="color:#555;margin-bottom:6px">{cargo}</div>
      <table role="presentation" cellpadding="0" cellspacing="0" style="font-size:13px">
        {linha("M:", f'<a href="tel:{tel_link}" style="color:#333;text-decoration:none">{escape(tel)}</a>' if tel else "")}
        {linha("W:", f'<a href="{link_whatsapp(cfg.get("WHATSAPP"))}" style="color:#128C4A;text-decoration:none;font-weight:bold">WhatsApp</a>' if cfg.get("WHATSAPP") else "")}
        {linha("E:", f'<a href="mailto:{escape(email)}" style="color:{AZUL};text-decoration:none">{escape(email)}</a>' if email else "")}
        {linha("A:", end)}
      </table>
    </td>
  </tr>
</table>"""


def layout_html(conteudo, cfg, logo_src, assunto):
    aviso = escape(cfg.get("AVISO_LEGAL", ""))
    descad = escape(cfg.get("TEXTO_DESCADASTRO", ""))
    topo = ('<tr><td align="center" style="padding:26px 32px 18px;border-bottom:1px solid #E6EBF2">'
            f'<img src="{logo_src}" width="210" alt="{escape(cfg.get("NOME_EMPRESA", ""))}" '
            'style="display:block;border:0;width:210px;height:auto"></td></tr>') if logo_src else ""
    return f"""<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{escape(assunto)}</title></head>
<body style="margin:0;padding:0;background:#EEF2F7">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#EEF2F7">
<tr><td align="center" style="padding:24px 12px">
  <table role="presentation" width="600" cellpadding="0" cellspacing="0" style="width:100%;max-width:600px;background:#ffffff;border-radius:8px;overflow:hidden;border:1px solid #DCE3EC">
    <tr><td style="height:6px;background:{AZUL};font-size:0;line-height:0">&nbsp;</td></tr>
    {topo}
    <tr><td style="padding:28px 36px 8px;font-family:Arial,Helvetica,sans-serif;font-size:15px;line-height:1.6;color:#222">
      {conteudo}
    </td></tr>
    <tr><td style="padding:8px 36px 28px">
      <table role="presentation" cellpadding="0" cellspacing="0" width="100%"><tr><td style="border-top:1px solid #E6EBF2;padding-top:20px">
        {assinatura_html(cfg, logo_src)}
      </td></tr></table>
    </td></tr>
    <tr><td style="height:4px;background:{AZUL};font-size:0;line-height:0">&nbsp;</td></tr>
  </table>
  <table role="presentation" width="600" cellpadding="0" cellspacing="0" style="width:100%;max-width:600px">
    <tr><td style="padding:14px 8px;font-family:Arial,Helvetica,sans-serif;font-size:11px;line-height:1.5;color:#8A94A3;text-align:justify">
      {f'<p style="margin:0 0 8px">{descad}</p>' if descad else ''}{aviso}
    </td></tr>
  </table>
</td></tr></table>
</body></html>"""


def assinatura_texto(cfg):
    partes = [cfg.get("NOME_REMETENTE", ""), cfg.get("CARGO", ""), cfg.get("NOME_EMPRESA", ""), "",
              f"M: {cfg.get('TELEFONE', '')}",
              f"E: {cfg.get('EMAIL_ASSINATURA') or cfg.get('EMAIL_REMETENTE', '')}",
              f"A: {cfg.get('ENDERECO', '')}"]
    rodape = "\n\n".join(x for x in [cfg.get("TEXTO_DESCADASTRO", ""), cfg.get("AVISO_LEGAL", "")] if x)
    return "\n".join(partes) + ("\n\n" + rodape if rodape else "")


def montar(l, assunto_pad, corpo_mod, cfg, logo_src="cid:logo"):
    if not LOGO.exists():  # logo.png é opcional
        logo_src = ""
    nome = l["NOME"]
    grupo = l["TIPO"] == "GRUPO"
    emp = empresa_curta(l["EMPRESA"])
    campos = Vazio(
        saudacao=l["SAUDACAO"],
        nome=nome,
        primeiro_nome=nome.split()[0] if nome else "",
        empresa=l["EMPRESA"],
        empresa_curta=emp,
        frase_personalizada=l["FRASE_PERSONALIZADA"],
        tipo=l["TIPO"],
        voces="vocês" if grupo else "você",
        voces_trabalham="vocês trabalham" if grupo else "você trabalha",
        podem="podem" if grupo else "pode",
        lhes="lhes" if grupo else "lhe",
        da_empresa=f"{da_do(emp)} {emp}",
        a_empresa=("o " if da_do(emp) == "do" else "a ") + emp,
        remetente=cfg.get("NOME_REMETENTE", ""),
        bloco_catalogo="@@CATALOGO@@",
        botao_whatsapp="@@WHATSAPP@@",
    )
    assunto = (l["ASSUNTO"] or assunto_pad).format_map(campos)
    texto = corpo_mod.format_map(campos)
    texto = re.sub(r"\n{3,}", "\n\n", texto).strip()  # some linhas vazias se frase em branco
    simples = estilo_simples(cfg)
    if simples:
        especiais = {"@@CATALOGO@@": bloco_catalogo_simples(cfg), "@@WHATSAPP@@": whatsapp_simples(cfg)}
        html = layout_simples(corpo_html(texto, especiais, simples=True), cfg, logo_src, assunto)
    else:
        especiais = {"@@CATALOGO@@": bloco_catalogo_html(cfg),
                     "@@WHATSAPP@@": botao_whatsapp_html(cfg, campos)}
        html = layout_html(corpo_html(texto, especiais), cfg, logo_src, assunto)
    anexo = arquivo_catalogo(cfg)
    num = cfg.get("WHATSAPP") or cfg.get("TELEFONE", "")
    texto_puro = texto.replace(
        "@@CATALOGO@@", f"{CATALOGO_TITULO} — {CATALOGO_TEXTO}"
        + (f" (arquivo: {anexo.name})" if anexo else "") + (f"\nBaixar o PDF: {link_download(CATALOGO_LINK)}" if CATALOGO_LINK.strip() else "")
    ).replace("@@WHATSAPP@@", ("WhatsApp: " + link_whatsapp(num, "" if simples else WHATSAPP_MENSAGEM.format_map(campos)))
              if num else "")
    texto_puro = re.sub(r"\*\*(.+?)\*\*", r"\1", texto_puro) + "\n\n--\n" + assinatura_texto(cfg) + "\n"
    return assunto, texto_puro, html


def cmd_previa(args):
    """Gera arquivos HTML para ver no navegador exatamente como o e-mail vai ficar."""
    cfg = ler_config()
    assunto_pad, corpo_mod = ler_modelo()
    _, _, _, linhas = carregar_revisao()
    alvo = [l for l in linhas if sim(l["ENVIAR"])] or linhas
    pasta = PASTA / "previa"
    pasta.mkdir(exist_ok=True)
    if LOGO.exists():
        (pasta / "logo.png").write_bytes(LOGO.read_bytes())
    for i, l in enumerate(alvo, 1):
        assunto, _, html = montar(l, assunto_pad, corpo_mod, cfg, logo_src="logo.png")
        nomearq = f"{i:02d}_{re.sub(r'[^a-z0-9]+', '_', l['EMAIL'].lower())}.html"
        (pasta / nomearq).write_text(html, encoding="utf-8")
    print(f"✔ {len(alvo)} prévias em {pasta}  — abra qualquer .html no navegador.")
    if alvo and getattr(args, "abrir", False):
        import webbrowser
        primeiro = sorted(pasta.glob("*.html"))[0]
        webbrowser.open(primeiro.resolve().as_uri())
        print(f"  (abrindo {primeiro.name} no navegador)")


def validar_trava(linhas):
    a_enviar = [l for l in linhas if sim(l["ENVIAR"])]
    problemas = []
    for l in a_enviar:
        if not sim(l["APROVADO"]):
            problemas.append(f"linha {l['_row']}: {l['EMAIL']} — APROVADO ≠ SIM")
        elif not l["SAUDACAO"]:
            problemas.append(f"linha {l['_row']}: {l['EMAIL']} — SAUDACAO vazia")
        elif not EMAIL_RE.fullmatch(l["EMAIL"]):
            problemas.append(f"linha {l['_row']}: e-mail inválido '{l['EMAIL']}'")
        elif re.search(r"[{}]", l["SAUDACAO"] + l["FRASE_PERSONALIZADA"]):
            problemas.append(f"linha {l['_row']}: há '{{' ou '}}' na saudação/frase")
    return a_enviar, problemas


def checar_textos(cfg):
    """Retorna mensagem de bloqueio (str) ou ''."""
    if re.search(r"\[[^\]]+\]", ASSUNTO_PADRAO + TEXTO_EMAIL + CATALOGO_TEXTO + CATALOGO_TITULO) \
            or any(re.fullmatch(r"Linha \d", t) for t in CATALOGO_DESTAQUES):
        return "O texto ainda tem campos [ENTRE COLCHETES] do modelo. Ajuste o texto do e-mail."
    if not cfg.get("EMAIL_REMETENTE") or "seu.email@" in cfg.get("EMAIL_REMETENTE", ""):
        return "Preencha o e-mail remetente nas configurações."
    if CATALOGO_LINK.strip() and not re.fullmatch(r"https?://\S+", CATALOGO_LINK.strip(), re.I):
        return "O link do catálogo precisa começar com https:// e não pode ter espaços."
    for campo, rotulo in (("EMAILS_COPIA", "Cópia (Cc)"), ("EMAILS_COPIA_OCULTA", "Cópia oculta (Cco)")):
        ruins = [e for e in lista_emails(cfg.get(campo)) if not EMAIL_RE.fullmatch(e)]
        if ruins:
            return f"{rotulo}: e-mail inválido: {', '.join(ruins)}"
    return ""


def preparar_anexo(cfg):
    """-> (anexo_path|None, bytes|None, subtipo, aviso). Levanta ValueError se bloqueado."""
    anexo = arquivo_catalogo(cfg)
    if not cfg.get("ANEXO_CATALOGO", "").strip():
        return None, None, "octet-stream", ""
    if not anexo:
        raise ValueError(f"Catálogo '{cfg['ANEXO_CATALOGO']}' não encontrado na pasta. "
                         "Envie o PDF (Configurações) ou deixe o campo do anexo vazio.")
    dados = anexo.read_bytes()
    mb = len(dados) / 1024 / 1024
    if mb > 15:
        raise ValueError(f"Catálogo tem {mb:.1f} MB (máx. 15 MB). Comprima o PDF ou use um link on-line.")
    aviso = f"Catálogo com {mb:.1f} MB — anexos grandes aumentam a chance de cair no spam." if mb > 5 else ""
    sub = {"pdf": "pdf", "zip": "zip"}.get(anexo.suffix.lower().lstrip("."), "octet-stream")
    return anexo, dados, sub, aviso


PROVEDORES_MX = [  # (trecho do MX, nome, [(host, porta)])
    ("google", "Google (Gmail / Workspace)", [("smtp.gmail.com", 587), ("smtp.gmail.com", 465)]),
    ("outlook", "Microsoft 365 / Outlook", [("smtp.office365.com", 587), ("smtp-mail.outlook.com", 587)]),
    ("hostinger", "Hostinger", [("smtp.hostinger.com", 465), ("smtp.hostinger.com", 587)]),
    ("titan.email", "Titan (Hostgator/Hostinger)", [("smtp.titan.email", 465), ("smtp.titan.email", 587)]),
    ("locaweb", "Locaweb", [("email-ssl.com.br", 465), ("smtplw.com.br", 587)]),
    ("zoho", "Zoho", [("smtp.zoho.com", 465), ("smtp.zoho.com", 587)]),
    ("umbler", "Umbler", [("smtp.umbler.com", 587)]),
    ("kinghost", "KingHost", [("smtp.kinghost.net", 587), ("smtp.kinghost.net", 465)]),
    ("uhserver", "UOL Host", [("smtps.uhserver.com", 465)]),
    ("hostgator", "HostGator", [("mail.{d}", 465), ("mail.{d}", 587)]),
    ("websitewelcome", "HostGator", [("mail.{d}", 465), ("mail.{d}", 587)]),
]


def buscar_mx(dominio):
    """Descobre os servidores de e-mail (MX) do domínio. Retorna lista de nomes."""
    import json
    import subprocess
    import urllib.request
    try:
        url = f"https://dns.google/resolve?name={dominio}&type=MX"
        with urllib.request.urlopen(url, timeout=8) as r:
            dados = json.loads(r.read().decode())
        mx = [x["data"].split()[-1].rstrip(".").lower() for x in dados.get("Answer", []) if x.get("type") == 15]
        if mx:
            return mx
    except Exception:  # noqa: BLE001
        pass
    try:  # Windows / Linux sem internet para DoH
        out = subprocess.run(["nslookup", "-type=mx", dominio], capture_output=True, text=True, timeout=10).stdout
        return [m.rstrip(".").lower() for m in re.findall(r"mail exchanger = (?:\d+ )?(\S+)", out)]
    except Exception:  # noqa: BLE001
        return []


def explicar_erro(e) -> str:
    t = f"{type(e).__name__}: {e}"
    if isinstance(e, RuntimeError):
        return str(e)
    if isinstance(e, smtplib.SMTPAuthenticationError) and ("basic authentication" in str(e).lower()
                                                           or "5.7.139" in str(e) or "5.7.3" in str(e)):
        return ("A Microsoft bloqueou o login por senha (SMTP) nesta conta. Solução: em Configurações, mude "
                "'Modo de envio' para OUTLOOK (envia pelo Outlook instalado), ou peça ao TI para liberar "
                "'SMTP AUTH'. " + f"{type(e).__name__}: {e}")
    if isinstance(e, smtplib.SMTPAuthenticationError):
        return ("Login recusado pelo servidor. Se a senha está certa, a Microsoft 365 provavelmente bloqueou o "
                "login por senha nesta conta → em Configurações mude 'Modo de envio' para OUTLOOK, ou peça ao TI "
                "para liberar 'SMTP AUTH'. " + t)
    if isinstance(e, smtplib.SMTPServerDisconnected):
        return ("O servidor fechou a conexão — normalmente servidor/porta errados para o seu provedor, "
                "ou antivírus/firewall bloqueando. " + t)
    if isinstance(e, ssl.SSLError):
        return "Erro de criptografia (porta não combina com o modo: 465=SSL, 587=STARTTLS) ou antivírus. " + t
    if isinstance(e, (TimeoutError, OSError)) and "timed out" in str(e).lower():
        return "Tempo esgotado — porta bloqueada pela rede/antivírus ou servidor errado. " + t
    if "getaddrinfo" in str(e) or "Name or service" in str(e) or "nodename" in str(e):
        return "Servidor não existe (nome errado). " + t
    if isinstance(e, ConnectionRefusedError):
        return "Conexão recusada nessa porta. " + t
    return t


def _abrir(host, porta, timeout=20):
    ctx = ssl.create_default_context()
    if porta == 465:
        s = smtplib.SMTP_SSL(host, porta, context=ctx, timeout=timeout)
        s.ehlo()
    else:
        s = smtplib.SMTP(host, porta, timeout=timeout)
        s.ehlo()
        s.starttls(context=ctx)
        s.ehlo()
    return s


def conectar_smtp(cfg, senha):
    """Conecta e faz login. Se a porta configurada cair, tenta a outra (465 <-> 587) automaticamente."""
    host = cfg.get("SMTP_HOST", "smtp.gmail.com").strip()
    porta = int(cfg.get("SMTP_PORTA", "587") or 587)
    tentativas = [porta] + [p for p in (587, 465) if p != porta]
    ultimo = None
    for p in tentativas:
        try:
            s = _abrir(host, p)
        except (smtplib.SMTPServerDisconnected, ssl.SSLError, ConnectionError, TimeoutError,
                smtplib.SMTPNotSupportedError) as e:
            ultimo = e
            continue
        s.login(cfg.get("EMAIL_REMETENTE", ""), senha)  # erro de senha sobe direto (não adianta trocar porta)
        return s
    raise ultimo


def diagnosticar_smtp(cfg, senha="", avisar=print):
    """Testa servidor/porta (e login, se houver senha). Retorna {'ok','host','porta','provedor','passos'}."""
    email = cfg.get("EMAIL_REMETENTE", "")
    dominio = email.split("@")[-1].lower() if "@" in email else ""
    passos = []

    def p(txt):
        passos.append(txt)
        avisar(txt)

    provedor, candidatos = "", []
    if dominio:
        p(f"🔎 Procurando o provedor de e-mail de {dominio} ...")
        mx = buscar_mx(dominio)
        if mx:
            p(f"   Servidores de recebimento (MX): {', '.join(mx[:3])}")
            for chave, nome, lista in PROVEDORES_MX:
                if any(chave in m for m in mx):
                    provedor = nome
                    candidatos += [(h.format(d=dominio), pt) for h, pt in lista]
                    break
            p(f"   ➜ Provedor identificado: {provedor}" if provedor else "   ➜ Provedor próprio/desconhecido")
        else:
            p("   (não consegui consultar o MX — sem internet ou DNS bloqueado)")
    atual = (cfg.get("SMTP_HOST", "").strip(), int(cfg.get("SMTP_PORTA", "587") or 587))
    if atual[0] and atual not in candidatos:
        candidatos.insert(0 if not provedor else len(candidatos), atual)
    for h in ([f"mail.{dominio}", f"smtp.{dominio}"] if dominio else []):
        for pt in (465, 587):
            if (h, pt) not in candidatos:
                candidatos.append((h, pt))
    vistos = []
    for c in candidatos:
        if c not in vistos:
            vistos.append(c)

    conectou = None
    for host, porta in vistos[:10]:
        p(f"🔌 Testando {host}:{porta} ({'SSL' if porta == 465 else 'STARTTLS'}) ...")
        try:
            s = _abrir(host, porta, timeout=12)
        except Exception as e:  # noqa: BLE001
            p("   ✖ " + explicar_erro(e))
            continue
        auth = s.has_extn("auth")
        p(f"   ✔ Conectou{'' if auth else ' (mas o servidor não oferece login AUTH)'}")
        if not senha:
            s.quit()
            conectou = conectou or (host, porta)
            p("   (sem senha informada — login não testado)")
            break
        try:
            s.login(email, senha)
            s.quit()
            p(f"   ✔ LOGIN OK como {email}")
            return {"ok": True, "host": host, "porta": porta, "provedor": provedor, "passos": passos}
        except Exception as e:  # noqa: BLE001
            p("   ✖ " + explicar_erro(e))
            try:
                s.quit()
            except Exception:  # noqa: BLE001
                pass
            if isinstance(e, smtplib.SMTPAuthenticationError):
                conectou = (host, porta)
                break  # servidor certo, senha/permissão errada
    if conectou:
        return {"ok": False, "host": conectou[0], "porta": conectou[1], "provedor": provedor, "passos": passos,
                "parcial": True}
    p("✖ Nenhuma combinação conectou. Verifique antivírus (proteção de e-mail), rede da empresa "
      "ou pergunte ao TI qual é o servidor SMTP.")
    return {"ok": False, "host": "", "porta": 0, "provedor": provedor, "passos": passos}


def construir_msg(l, cfg, anexo_info, destino=None, prefixo_assunto=""):
    assunto_pad, corpo_mod = ler_modelo()
    assunto, texto, html = montar(l, assunto_pad, corpo_mod, cfg)
    usuario = cfg.get("EMAIL_REMETENTE", "")
    nome_rem = cfg.get("NOME_REMETENTE", "")
    msg = EmailMessage()
    msg["From"] = formataddr((nome_rem, usuario)) if nome_rem else usuario
    msg["To"] = destino or l["EMAIL"]
    msg["Subject"] = prefixo_assunto + assunto
    msg["Message-ID"] = make_msgid(domain=usuario.split("@")[-1] or "localhost")
    if cfg.get("RESPONDER_PARA") and cfg["RESPONDER_PARA"].lower() != usuario.lower():
        msg["Reply-To"] = cfg["RESPONDER_PARA"]
    if not destino:  # envio de teste não copia a equipe
        cc, cco = copias_envio(cfg, l["EMAIL"])
        if cc:
            msg["Cc"] = ", ".join(cc)
        if cco:
            msg["Bcc"] = ", ".join(cco)  # o smtplib entrega e remove este cabeçalho
    msg.set_content(texto)
    msg.add_alternative(html, subtype="html")
    if LOGO.exists():  # logo embutida (aparece mesmo sem "carregar imagens")
        msg.get_payload()[1].add_related(LOGO.read_bytes(), maintype="image", subtype="png",
                                         cid="<logo>", filename="logo.png")
    anexo, dados, sub, _ = anexo_info
    if dados:
        msg.add_attachment(dados, maintype="application", subtype=sub, filename=anexo.name)
    return msg


def modo_envio(cfg) -> str:
    return (cfg.get("MODO_ENVIO") or "SMTP").strip().upper()


class EnvioSMTP:
    """Envia pelo servidor de e-mail (login com e-mail + senha)."""
    precisa_senha = True

    def __init__(self, cfg, senha):
        self.cfg, self.senha = cfg, senha
        self.smtp = conectar_smtp(cfg, senha)

    def enviar(self, l, anexo_info, destino=None, prefixo=""):
        msg = construir_msg(l, self.cfg, anexo_info, destino=destino, prefixo_assunto=prefixo)
        try:
            self.smtp.send_message(msg)
        except (smtplib.SMTPServerDisconnected, OSError):
            self.smtp = conectar_smtp(self.cfg, self.senha)  # reconecta 1x
            self.smtp.send_message(msg)

    def fechar(self):
        try:
            self.smtp.quit()
        except Exception:  # noqa: BLE001
            pass


class EnvioOutlook:
    """Envia pelo Outlook (clássico) instalado no Windows — sem senha; fica em 'Itens Enviados'."""
    precisa_senha = False
    PR_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
    PR_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"

    def __init__(self, cfg, senha=""):
        self.cfg = cfg
        try:
            import pythoncom
            import win32com.client
        except ImportError:
            raise RuntimeError("Para enviar pelo Outlook, instale o complemento uma vez no Prompt de Comando:  "
                               "pip install pywin32   (depois feche e abra o painel de novo)") from None
        self._pythoncom = pythoncom
        pythoncom.CoInitialize()
        try:
            self.ol = win32com.client.Dispatch("Outlook.Application")
            sessao = self.ol.Session
        except Exception as e:  # noqa: BLE001
            raise RuntimeError("Não consegui abrir o Outlook. Use o Outlook CLÁSSICO (o 'Novo Outlook' não "
                               f"permite automação) e deixe-o configurado com sua conta. ({e})") from None
        self.conta, self.contas = None, []
        alvo = cfg.get("EMAIL_REMETENTE", "").strip().lower()
        for i in range(1, sessao.Accounts.Count + 1):
            acc = sessao.Accounts.Item(i)
            end = (getattr(acc, "SmtpAddress", "") or "").lower()
            self.contas.append(end)
            if end == alvo:
                self.conta = acc

    def enviar(self, l, anexo_info, destino=None, prefixo=""):
        assunto_pad, corpo_mod = ler_modelo()
        assunto, _, html = montar(l, assunto_pad, corpo_mod, self.cfg, logo_src="cid:logo")
        m = self.ol.CreateItem(0)
        m.To = destino or l["EMAIL"]
        m.Subject = prefixo + assunto
        if not destino:  # envio de teste não copia a equipe
            cc, cco = copias_envio(self.cfg, l["EMAIL"])
            if cc:
                m.CC = "; ".join(cc)
            if cco:
                m.BCC = "; ".join(cco)
        resp = self.cfg.get("RESPONDER_PARA", "")
        if resp and resp.lower() != self.cfg.get("EMAIL_REMETENTE", "").lower():
            m.ReplyRecipients.Add(resp)
        if LOGO.exists():
            att = m.Attachments.Add(str(LOGO.resolve()))
            att.PropertyAccessor.SetProperty(self.PR_CONTENT_ID, "logo")
            try:
                att.PropertyAccessor.SetProperty(self.PR_HIDDEN, True)
            except Exception:  # noqa: BLE001
                pass
        m.HTMLBody = html
        anexo = anexo_info[0]
        if anexo:
            m.Attachments.Add(str(Path(anexo).resolve()))
        if self.conta is not None:
            try:
                m.SendUsingAccount = self.conta
            except Exception:  # noqa: BLE001
                m._oleobj_.Invoke(*(64209, 0, 8, 0, self.conta))
        m.Send()

    def fechar(self):
        try:
            self._pythoncom.CoUninitialize()
        except Exception:  # noqa: BLE001
            pass


def criar_envio(cfg, senha=""):
    return EnvioOutlook(cfg) if modo_envio(cfg) == "OUTLOOK" else EnvioSMTP(cfg, senha)


def diagnosticar_outlook(cfg):
    passos = ["🔎 Modo OUTLOOK: procurando o Outlook instalado ..."]
    try:
        o = EnvioOutlook(cfg)
    except Exception as e:  # noqa: BLE001
        passos.append("   ✖ " + str(e))
        return {"ok": False, "host": "", "porta": 0, "provedor": "Outlook", "passos": passos}
    passos.append("   ✔ Outlook encontrado. Contas: " + (", ".join(o.contas) or "(nenhuma)"))
    alvo = cfg.get("EMAIL_REMETENTE", "")
    if o.conta is not None:
        passos.append(f"   ✔ Vai enviar pela conta {alvo}")
    else:
        passos.append(f"   ⚠ A conta {alvo} não está no Outlook — vai usar a conta padrão do Outlook.")
    o.fechar()
    return {"ok": True, "host": "", "porta": 0, "provedor": "Outlook", "passos": passos}


def cmd_enviar(args):
    cfg = ler_config()
    assunto_pad, corpo_mod = ler_modelo()
    wb, ws, idx, linhas = carregar_revisao()
    a_enviar, problemas = validar_trava(linhas)

    if problemas:
        print("🔒 ENVIO BLOQUEADO — revise antes de disparar:\n")
        for p in problemas[:40]:
            print("  •", p)
        if len(problemas) > 40:
            print(f"  ... e mais {len(problemas) - 40}")
        print("\nCorrija no Excel (APROVADO = SIM) ou use a opção 2 do menu (Revisar).")
        sys.exit(1)

    enviados = ja_enviados()
    fila = [l for l in a_enviar if l["EMAIL"].lower() not in enviados]
    limite = args.limite or int(cfg.get("LIMITE_POR_EXECUCAO", "30"))
    pulados = len(a_enviar) - len(fila)
    fila = fila[:limite]

    print(f"Aprovados: {len(a_enviar)}  |  já enviados antes: {pulados}  |  nesta execução: {len(fila)} (limite {limite})")
    if not fila:
        print("Nada a enviar.")
        return

    # prévia
    mostrar = fila if not args.real else fila[:3]
    for l in mostrar:
        assunto, texto, _ = montar(l, assunto_pad, corpo_mod, cfg)
        print("═" * 70)
        print(f"PARA:    {l['EMAIL']}   ({l['TIPO']} — {l['EMPRESA']})")
        cc, cco = copias_envio(cfg, l["EMAIL"])
        if cc:
            print(f"CC:      {', '.join(cc)}")
        if cco:
            print(f"CCO:     {', '.join(cco)}")
        print(f"ASSUNTO: {assunto}")
        print("─" * 70)
        print(texto)
    if not args.real:
        print("═" * 70)
        print("SIMULAÇÃO — nada foi enviado. Para enviar:  python disparo.py enviar --real")
        return

    bloqueio = checar_textos(cfg)
    if bloqueio:
        sys.exit("🔒 BLOQUEADO: " + bloqueio)
    try:
        anexo_info = preparar_anexo(cfg)
    except ValueError as e:
        sys.exit(f"🔒 BLOQUEADO: {e}")
    if anexo_info[3]:
        print("⚠ ", anexo_info[3])
    if anexo_info[0]:
        print(f"📎 Anexo: {anexo_info[0].name}")

    usuario = cfg.get("EMAIL_REMETENTE")
    senha = ""
    if modo_envio(cfg) != "OUTLOOK":
        senha = os.environ.get("SMTP_SENHA") or getpass.getpass(f"Senha de app SMTP de {usuario}: ")
    dmin = float(cfg.get("INTERVALO_MIN", "8"))
    dmax = float(cfg.get("INTERVALO_MAX", "10"))
    max_erros = int(cfg.get("PARAR_APOS_ERROS_SEGUIDOS", "3"))

    print("═" * 70)
    conf = input(f"⚠  Enviar {len(fila)} e-mails REAIS de {usuario}? Digite ENVIAR {len(fila)} para confirmar: ")
    if conf.strip().upper() != f"ENVIAR {len(fila)}":
        sys.exit("Cancelado.")

    try:
        envio = criar_envio(cfg, senha)
    except Exception as e:  # noqa: BLE001
        sys.exit("✖ " + explicar_erro(e) + "\nDica: rode  python disparo.py diagnosticar")

    ok = erros_seguidos = 0
    try:
        for i, l in enumerate(fila, 1):
            try:
                envio.enviar(l, anexo_info)
                ok += 1
                erros_seguidos = 0
                registrar(l["EMAIL"], l["EMPRESA"], "ENVIADO")
                salvar_celula(ws, idx, l["_row"], "STATUS", f"ENVIADO {datetime.now():%d/%m %H:%M}")
                print(f"[{i}/{len(fila)}] ✔ {l['EMAIL']}")
            except Exception as e:  # noqa: BLE001
                erros_seguidos += 1
                registrar(l["EMAIL"], l["EMPRESA"], "ERRO", str(e)[:200])
                salvar_celula(ws, idx, l["_row"], "STATUS", f"ERRO: {str(e)[:60]}")
                print(f"[{i}/{len(fila)}] ✖ {l['EMAIL']}  — {e}")
                if erros_seguidos >= max_erros:
                    print(f"\n⛔ {max_erros} erros seguidos — parando por segurança.")
                    break
            try:
                wb.save(REVISAO)
            except PermissionError:
                pass  # planilha aberta no Excel; o log CSV continua valendo
            if i < len(fila):
                time.sleep(random.uniform(dmin, dmax))
    except KeyboardInterrupt:
        print("\nInterrompido. O que já foi enviado está no log e não será repetido.")
    finally:
        envio.fechar()
        try:
            wb.save(REVISAO)
        except PermissionError:
            print("(feche o Excel para gravar a coluna STATUS; o log_envios.csv está atualizado)")
    print(f"\nConcluído: {ok} enviados. Log: {LOG.name}")


class _Args:
    def __init__(self, **kw):
        self.__dict__.update(kw)


def planilha_origem():
    padrao = PASTA / "Empresas_Prospectadas.xlsx"
    if padrao.exists():
        return str(padrao)
    opcoes = [f for f in PASTA.glob("*.xlsx") if f.name != REVISAO.name and not f.name.startswith("~$")]
    if len(opcoes) == 1:
        return str(opcoes[0])
    for i, f in enumerate(opcoes, 1):
        print(f"  {i}) {f.name}")
    r = input("Número da planilha de empresas (ou caminho completo): ").strip().strip('"')
    return str(opcoes[int(r) - 1]) if r.isdigit() and 0 < int(r) <= len(opcoes) else r


def menu():
    """Menu usado quando o script é rodado sem comando (botão ▶ do VS Code, duplo clique...)."""
    os.chdir(PASTA)
    while True:
        print("\n" + "═" * 56)
        print("   DISPARADOR DE E-MAIL COMERCIAL")
        print("═" * 56)
        print("  1) Preparar planilha de revisão (a partir do Excel)")
        print("  2) Revisar contatos no terminal (a trava)")
        print("  3) Ver prévia dos e-mails no navegador")
        print("  4) Ver relação empresa – nome – telefone")
        print("  5) Simular envio (não envia nada)")
        print("  6) ENVIAR DE VERDADE")
        print("  0) Sair")
        op = input("\nEscolha: ").strip()
        try:
            if op == "1":
                forcar = False
                if REVISAO.exists():
                    r = input(f"⚠  {REVISAO.name} já existe e pode ter sua revisão. Apagar e refazer? (s/N): ")
                    if r.strip().lower() != "s":
                        continue
                    forcar = True
                cmd_preparar(_Args(planilha=planilha_origem(), forcar=forcar))
            elif op == "2":
                cmd_revisar(_Args(todas=False))
            elif op == "3":
                cmd_previa(_Args(abrir=True))
            elif op == "4":
                cmd_contatos(_Args(planilha=planilha_origem()))
            elif op == "5":
                cmd_enviar(_Args(real=False, limite=None))
            elif op == "6":
                cmd_enviar(_Args(real=True, limite=None))
            elif op in ("0", "q", "sair", ""):
                break
            else:
                print("Opção inválida.")
        except SystemExit as e:  # mensagens de bloqueio/erro: mostra e volta ao menu
            if e.code not in (None, 0):
                print(e.code if isinstance(e.code, str) else "")
        except KeyboardInterrupt:
            print("\n(cancelado)")
        input("\n[ENTER para voltar ao menu]")


def main():
    if len(sys.argv) == 1:
        import painel
        painel.iniciar()
        return
    if sys.argv[1] == "diagnosticar":
        cfg = ler_config()
        if modo_envio(cfg) == "OUTLOOK":
            print("\n".join(diagnosticar_outlook(cfg)["passos"]))
            return
        r = diagnosticar_smtp(cfg, getpass.getpass("Senha (ENTER para testar só a conexão): "))
        if r["host"]:
            print(f"\nSugestão: SMTP_HOST = {r['host']}   SMTP_PORTA = {r['porta']}")
        return
    if sys.argv[1] == "menu":
        menu()
        return
    if sys.argv[1] == "painel":
        import painel
        painel.iniciar()
        return
    p = argparse.ArgumentParser(description="Disparo de e-mails personalizados com trava de revisão.")
    sub = p.add_subparsers(dest="cmd", required=True)
    a = sub.add_parser("preparar", help="gera revisao_envio.xlsx a partir da planilha original")
    a.add_argument("planilha")
    a.add_argument("--forcar", action="store_true", help="sobrescreve revisao_envio.xlsx existente")
    b = sub.add_parser("revisar", help="revisão interativa linha a linha (a trava)")
    b.add_argument("--todas", action="store_true", help="inclui linhas já aprovadas")
    c = sub.add_parser("enviar", help="simula (padrão) ou envia com --real")
    c.add_argument("--real", action="store_true", help="envia de verdade")
    c.add_argument("--limite", type=int, help="máximo de e-mails nesta execução")
    e = sub.add_parser("previa", help="gera arquivos .html para ver o e-mail no navegador")
    e.add_argument("--abrir", action="store_true", help="abre a 1ª prévia no navegador")
    d = sub.add_parser("contatos", help="mostra/atualiza a relação empresa - nome - telefone")
    d.add_argument("planilha")
    args = p.parse_args()
    {"preparar": cmd_preparar, "revisar": cmd_revisar, "enviar": cmd_enviar,
     "previa": cmd_previa, "contatos": cmd_contatos}[args.cmd](args)


carregar_textos_salvos()

if __name__ == "__main__":
    main()
