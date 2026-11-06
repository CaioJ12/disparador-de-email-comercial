#!/usr/bin/env python3
"""
PAINEL do disparo de e-mails — abre no navegador (http://127.0.0.1:8765).
Roda só no seu computador (não fica acessível pela internet).

Para abrir:  dois cliques em ABRIR_PAINEL.bat   ou   python painel.py
"""
import json
import random
import re
import smtplib
import sys
import threading
import time
import traceback
import webbrowser
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

sys.path.insert(0, str(Path(__file__).resolve().parent))
import disparo as d  # noqa: E402

PASTA = d.PASTA
HTML = PASTA / "painel.html"
EDITAVEIS = ["ENVIAR", "APROVADO", "EMPRESA", "EMAIL", "TIPO", "NOME", "SAUDACAO",
             "FRASE_PERSONALIZADA", "ASSUNTO"]
CHAVES_CONFIG = ["NOME_EMPRESA", "EMAIL_REMETENTE", "NOME_REMETENTE", "RESPONDER_PARA", "EMAILS_COPIA", "EMAILS_COPIA_OCULTA",
                 "CARGO", "TELEFONE", "WHATSAPP",
                 "EMAIL_ASSINATURA", "ENDERECO", "TEXTO_DESCADASTRO", "AVISO_LEGAL", "ANEXO_CATALOGO",
                 "SMTP_HOST", "SMTP_PORTA", "INTERVALO_MIN", "INTERVALO_MAX", "LIMITE_POR_EXECUCAO",
                 "PARAR_APOS_ERROS_SEGUIDOS", "MODO_ENVIO", "ESTILO_EMAIL"]

ENVIO = {"rodando": False, "parar": False, "total": 0, "feitos": 0, "ok": 0, "erros": 0, "log": []}
TRAVA = threading.Lock()


# ----------------------------------------------------------------------------
# dados
# ----------------------------------------------------------------------------
def origem_padrao():
    p = PASTA / "Empresas_Prospectadas.xlsx"
    if p.exists():
        return p
    xs = [f for f in PASTA.glob("*.xlsx") if f.name != d.REVISAO.name and not f.name.startswith("~$")]
    return xs[0] if xs else None


def salvar_config(novos: dict):
    linhas = d.CONFIG.read_text(encoding="utf-8").splitlines() if d.CONFIG.exists() else []
    feitos = set()
    for i, ln in enumerate(linhas):
        m = re.match(r"\s*([A-Z_]+)\s*=", ln)
        if m and m.group(1) in novos:
            k = m.group(1)
            linhas[i] = f"{k} = {str(novos[k]).replace(chr(10), ' ').strip()}"
            feitos.add(k)
    for k, v in novos.items():
        if k not in feitos and k in CHAVES_CONFIG:
            linhas.append(f"{k} = {str(v).replace(chr(10), ' ').strip()}")
    d.CONFIG.write_text("\n".join(linhas) + "\n", encoding="utf-8")


def ler_contatos():
    if not d.REVISAO.exists():
        return []
    _, _, _, linhas = d.carregar_revisao()
    enviados = d.ja_enviados()
    for l in linhas:
        l["JA_ENVIADO"] = l["EMAIL"].lower() in enviados
    return linhas


def salvar_contatos(rows):
    wb, ws, idx, linhas = d.carregar_revisao()
    existentes = {l["_row"] for l in linhas}
    for r in rows:
        n = r.get("_row")
        if not n or n not in existentes:
            n = ws.max_row + 1
            d.salvar_celula(ws, idx, n, "OBS_ORIGINAL", "adicionado pelo painel")
        for c in EDITAVEIS:
            if c in r:
                v = r[c]
                if c in ("ENVIAR", "APROVADO"):
                    v = "SIM" if v in (True, "SIM", "sim", 1) else ("NAO" if c == "ENVIAR" else "")
                d.salvar_celula(ws, idx, n, c, str(v).strip())
    wb.save(d.REVISAO)


def ler_telefones():
    o = origem_padrao()
    if not o:
        return []
    regs = d.ler_origem(o)
    return [{"empresa": r["empresa"], "nomes": r["nomes"], "emails": r["emails"], "linkedin": r["linkedin"],
             "situacao": " | ".join(r["situacao"]),
             "telefones": [{"numero": t[0], "tipo": t[1], "wa": d.link_whatsapp(t[2]) if t[2] else ""}
                           for t in r["telefones"]]} for r in regs]


def linha_ficticia():
    return {"EMPRESA": "Empresa Exemplo S.A.", "EMAIL": "exemplo@empresa.com", "TIPO": "PESSOAL",
            "NOME": "Maria", "SAUDACAO": "Olá, Maria,", "FRASE_PERSONALIZADA": "", "ASSUNTO": "",
            "TELEFONE": "", "ENVIAR": "SIM", "APROVADO": "SIM", "OBS_ORIGINAL": "", "STATUS": ""}


def achar_linha(email, rascunho=None):
    if rascunho:
        base = linha_ficticia()
        base.update({k: str(v) for k, v in rascunho.items() if k in base})
        return base
    for l in ler_contatos():
        if l["EMAIL"].lower() == (email or "").lower():
            return l
    ls = ler_contatos()
    return ls[0] if ls else linha_ficticia()


def resumo():
    cfg = d.ler_config()
    linhas = ler_contatos()
    a_enviar, problemas = d.validar_trava(linhas) if linhas else ([], [])
    enviados = d.ja_enviados()
    pend = [l for l in a_enviar if l["EMAIL"].lower() not in enviados]
    bloqueios = list(problemas)
    b = d.checar_textos(cfg)
    if b:
        bloqueios.insert(0, b)
    aviso = ""
    try:
        anexo, _, _, aviso = d.preparar_anexo(cfg)
    except ValueError as e:
        anexo = None
        bloqueios.insert(0, str(e))
    if anexo and not d.CATALOGO_LINK.strip():
        aviso = (aviso + " " if aviso else "") + ("A faixa do catálogo não está clicável: cole o link do PDF "
                                                  "em Texto do e-mail. O PDF continua indo em anexo.")
    cc, cco = d.copias_envio(cfg)
    return {
        "total": len(linhas),
        "marcados": len(a_enviar),
        "aprovados": sum(1 for l in a_enviar if d.sim(l["APROVADO"])),
        "ja_enviados": len(a_enviar) - len(pend),
        "pendentes": len(pend),
        "limite": int(cfg.get("LIMITE_POR_EXECUCAO", "30") or 30),
        "bloqueios": bloqueios,
        "aviso": aviso,
        "copias": {"cc": cc, "cco": cco},
        "anexo": anexo.name if anexo else "",
        "anexo_mb": round(anexo.stat().st_size / 1024 / 1024, 2) if anexo else 0,
        "fila": [{"email": l["EMAIL"], "empresa": l["EMPRESA"], "saudacao": l["SAUDACAO"]} for l in pend],
    }


# ----------------------------------------------------------------------------
# envio em segundo plano
# ----------------------------------------------------------------------------
def log(msg):
    ENVIO["log"].append(f"{datetime.now():%H:%M:%S}  {msg}")
    ENVIO["log"] = ENVIO["log"][-400:]


def thread_envio(senha, limite):
    cfg = d.ler_config()
    try:
        wb, ws, idx, linhas = d.carregar_revisao()
        a_enviar, _ = d.validar_trava(linhas)
        enviados = d.ja_enviados()
        fila = [l for l in a_enviar if l["EMAIL"].lower() not in enviados][:limite]
        ENVIO.update(total=len(fila), feitos=0, ok=0, erros=0)
        anexo_info = d.preparar_anexo(cfg)
        dmin = float(cfg.get("INTERVALO_MIN", "8"))
        dmax = float(cfg.get("INTERVALO_MAX", "10"))
        max_erros = int(cfg.get("PARAR_APOS_ERROS_SEGUIDOS", "3"))
        if d.modo_envio(cfg) == "OUTLOOK":
            log("Abrindo o Outlook instalado ...")
        else:
            log(f"Conectando em {cfg.get('SMTP_HOST')}:{cfg.get('SMTP_PORTA')} ...")
        envio = d.criar_envio(cfg, senha)
        log("Conectado. Iniciando envio.")
        seguidos = 0
        for i, l in enumerate(fila, 1):
            if ENVIO["parar"]:
                log("⏹ Parado por você. O que já foi enviado não será repetido.")
                break
            try:
                envio.enviar(l, anexo_info)
                ENVIO["ok"] += 1
                seguidos = 0
                d.registrar(l["EMAIL"], l["EMPRESA"], "ENVIADO")
                d.salvar_celula(ws, idx, l["_row"], "STATUS", f"ENVIADO {datetime.now():%d/%m %H:%M}")
                log(f"✔ [{i}/{len(fila)}] {l['EMAIL']}")
            except Exception as e:  # noqa: BLE001
                ENVIO["erros"] += 1
                seguidos += 1
                d.registrar(l["EMAIL"], l["EMPRESA"], "ERRO", str(e)[:200])
                d.salvar_celula(ws, idx, l["_row"], "STATUS", f"ERRO: {str(e)[:60]}")
                log(f"✖ [{i}/{len(fila)}] {l['EMAIL']} — {e}")
                if seguidos >= max_erros:
                    log(f"⛔ {max_erros} erros seguidos — parando por segurança.")
                    break
            ENVIO["feitos"] = i
            try:
                wb.save(d.REVISAO)
            except PermissionError:
                pass
            if i < len(fila) and not ENVIO["parar"]:
                espera = random.uniform(dmin, dmax)
                time.sleep(espera)
        envio.fechar()
        try:
            wb.save(d.REVISAO)
        except PermissionError:
            log("(feche o Excel para gravar a coluna STATUS; o log_envios.csv está atualizado)")
        log(f"Concluído: {ENVIO['ok']} enviados, {ENVIO['erros']} erros.")
    except Exception as e:  # noqa: BLE001
        log("✖ " + d.explicar_erro(e))
        log("   Dica: Configurações → Testar conexão (descobre o servidor certo).")
    finally:
        ENVIO["rodando"] = False


# ----------------------------------------------------------------------------
# HTTP
# ----------------------------------------------------------------------------
class H(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _send(self, code, body, ctype="application/json; charset=utf-8"):
        if isinstance(body, (dict, list)):
            body = json.dumps(body, ensure_ascii=False)
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _ok_origem(self):
        o = self.headers.get("Origin")
        return not o or o == f"http://{self.headers.get('Host')}"

    def _json(self):
        n = int(self.headers.get("Content-Length") or 0)
        return json.loads(self.rfile.read(n) or b"{}")

    def do_GET(self):
        u = urlparse(self.path)
        q = {k: v[0] for k, v in parse_qs(u.query).items()}
        try:
            if u.path in ("/", "/index.html"):
                return self._send(200, HTML.read_text(encoding="utf-8"), "text/html; charset=utf-8")
            if u.path == "/logo.png":
                if not d.LOGO.exists():
                    return self._send(404, "sem logo", "text/plain; charset=utf-8")
                return self._send(200, d.LOGO.read_bytes(), "image/png")
            if u.path == "/catalogo":
                a = d.arquivo_catalogo(d.ler_config())
                if not a:
                    return self._send(404, "sem catálogo", "text/plain; charset=utf-8")
                return self._send(200, a.read_bytes(), "application/pdf")
            if u.path == "/api/estado":
                cfg = d.ler_config()
                t = d.textos_atuais()
                return self._send(200, {
                    "config": {k: cfg.get(k, "") for k in CHAVES_CONFIG},
                    "textos": t,
                    "contatos": ler_contatos(),
                    "colunas": d.COLUNAS,
                    "resumo": resumo(),
                    "origem": origem_padrao().name if origem_padrao() else "",
                    "revisao_existe": d.REVISAO.exists(),
                })
            if u.path == "/api/telefones":
                return self._send(200, ler_telefones())
            if u.path == "/api/resumo":
                return self._send(200, resumo())
            if u.path == "/api/progresso":
                return self._send(200, ENVIO)
            return self._send(404, {"erro": "não encontrado"})
        except PermissionError:
            return self._send(409, {"erro": "A planilha está aberta no Excel. Feche o Excel e tente de novo."})
        except Exception as e:  # noqa: BLE001
            traceback.print_exc()
            return self._send(500, {"erro": str(e)})

    def do_POST(self):
        if not self._ok_origem():
            return self._send(403, {"erro": "origem não permitida"})
        u = urlparse(self.path)
        try:
            if u.path in ("/api/upload_catalogo", "/api/upload_planilha"):
                nome = re.sub(r"[^\w.\- ]", "_", self.headers.get("X-Nome", "arquivo"))
                n = int(self.headers.get("Content-Length") or 0)
                if n > 30 * 1024 * 1024:
                    return self._send(413, {"erro": "arquivo grande demais (máx. 30 MB)"})
                dados = self.rfile.read(n)
                if u.path == "/api/upload_catalogo":
                    (PASTA / nome).write_bytes(dados)
                    salvar_config({"ANEXO_CATALOGO": nome})
                    return self._send(200, {"ok": True, "nome": nome})
                (PASTA / "Empresas_Prospectadas.xlsx").write_bytes(dados)
                return self._send(200, {"ok": True})

            dados = self._json()
            if u.path == "/api/textos":
                d.salvar_textos(dados)
                return self._send(200, {"ok": True})
            if u.path == "/api/textos/restaurar":
                if d.TEXTOS_ARQ.exists():
                    d.TEXTOS_ARQ.unlink()
                import importlib
                importlib.reload(d)
                return self._send(200, {"ok": True, "textos": d.textos_atuais()})
            if u.path == "/api/config":
                salvar_config({k: v for k, v in dados.items() if k in CHAVES_CONFIG})
                return self._send(200, {"ok": True})
            if u.path == "/api/contatos":
                salvar_contatos(dados.get("linhas", []))
                return self._send(200, {"ok": True, "contatos": ler_contatos()})
            if u.path == "/api/previa":
                cfg = d.ler_config()
                cfg.update({k: v for k, v in (dados.get("config") or {}).items() if k in CHAVES_CONFIG})
                antigos = d.textos_atuais()
                if dados.get("textos"):
                    d.aplicar_textos(dados["textos"])
                try:
                    l = achar_linha(dados.get("email"), dados.get("linha"))
                    assunto_pad, corpo_mod = d.ler_modelo()
                    assunto, texto, html = d.montar(l, assunto_pad, corpo_mod, cfg, logo_src="/logo.png")
                finally:
                    d.aplicar_textos(antigos)
                return self._send(200, {"assunto": assunto, "html": html, "texto": texto,
                                        "para": l["EMAIL"], "empresa": l["EMPRESA"]})
            if u.path == "/api/preparar":
                o = origem_padrao()
                if not o:
                    return self._send(400, {"erro": "Nenhuma planilha de empresas (.xlsx) na pasta."})
                try:
                    d.cmd_preparar(d_args(planilha=str(o), forcar=True))
                except SystemExit as e:
                    return self._send(400, {"erro": str(e.code)})
                return self._send(200, {"ok": True})
            if u.path == "/api/diagnostico":
                cfg = d.ler_config()
                cfg.update({k: v for k, v in (dados.get("config") or {}).items() if k in CHAVES_CONFIG and v})
                if d.modo_envio(cfg) == "OUTLOOK":
                    return self._send(200, d.diagnosticar_outlook(cfg))
                r = d.diagnosticar_smtp(cfg, dados.get("senha", ""), avisar=lambda x: None)
                return self._send(200, r)
            if u.path == "/api/teste":
                cfg = d.ler_config()
                senha, para = dados.get("senha", ""), dados.get("para", "").strip()
                precisa = d.modo_envio(cfg) != "OUTLOOK"
                if (precisa and not senha) or not d.EMAIL_RE.fullmatch(para):
                    return self._send(400, {"erro": "Informe a senha e um e-mail de destino válido."})
                anexo_info = d.preparar_anexo(cfg)
                l = achar_linha(dados.get("email"))
                envio = d.criar_envio(cfg, senha)
                try:
                    envio.enviar(l, anexo_info, destino=para, prefixo="[TESTE] ")
                finally:
                    envio.fechar()
                return self._send(200, {"ok": True, "msg": f"Teste enviado para {para} (personalizado como {l['EMAIL']})."})
            if u.path == "/api/enviar":
                with TRAVA:
                    if ENVIO["rodando"]:
                        return self._send(409, {"erro": "Já existe um envio em andamento."})
                    r = resumo()
                    if r["bloqueios"]:
                        return self._send(400, {"erro": "Envio bloqueado:\n• " + "\n• ".join(r["bloqueios"][:15])})
                    n = min(r["pendentes"], r["limite"])
                    if n == 0:
                        return self._send(400, {"erro": "Nada a enviar (nenhum aprovado pendente)."})
                    if dados.get("confirmacao", "").strip().upper() != f"ENVIAR {n}":
                        return self._send(400, {"erro": f"Digite exatamente: ENVIAR {n}"})
                    if not dados.get("senha") and d.modo_envio(d.ler_config()) != "OUTLOOK":
                        return self._send(400, {"erro": "Informe a senha do e-mail."})
                    ENVIO.update(rodando=True, parar=False, log=[], total=n, feitos=0, ok=0, erros=0)
                    threading.Thread(target=thread_envio, args=(dados["senha"], n), daemon=True).start()
                return self._send(200, {"ok": True})
            if u.path == "/api/parar":
                ENVIO["parar"] = True
                return self._send(200, {"ok": True})
            return self._send(404, {"erro": "não encontrado"})
        except PermissionError:
            return self._send(409, {"erro": "A planilha está aberta no Excel. Feche o Excel e tente de novo."})
        except (smtplib.SMTPException, OSError, RuntimeError) as e:
            return self._send(400, {"erro": d.explicar_erro(e)})
        except ValueError as e:
            return self._send(400, {"erro": str(e)})
        except Exception as e:  # noqa: BLE001
            traceback.print_exc()
            return self._send(500, {"erro": f"{type(e).__name__}: {e}"})


def d_args(**kw):
    return d._Args(**kw)


def iniciar(abrir=True):
    srv = None
    for porta in [8765, 8766, 8767, 0]:
        try:
            srv = ThreadingHTTPServer(("127.0.0.1", porta), H)
            break
        except OSError:
            continue
    url = f"http://127.0.0.1:{srv.server_address[1]}/"
    print("═" * 60)
    print("  DISPARADOR DE E-MAIL COMERCIAL")
    print(f"  Aberto em: {url}")
    print("  (deixe esta janela aberta; para encerrar, feche-a ou Ctrl+C)")
    print("═" * 60)
    if abrir:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\nPainel encerrado.")


if __name__ == "__main__":
    iniciar()
