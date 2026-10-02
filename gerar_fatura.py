#!/usr/bin/env python3
"""
gerar_fatura.py — gera faturas fictícias (layout DANFE NF3e, distribuidora fictícia) (HTML + PDF) a partir de
template-fatura-distribuidora.html e cenarios.json, para os testes de OCR/mudança contratual (UCH).

Uso:
  python3 gerar_fatura.py                       # gera todos os cenários de cenarios.json
  python3 gerar_fatura.py uch-base-convencional # gera só um cenário (nome do "arquivo")
  python3 gerar_fatura.py --listar-placeholders # lista os {{PLACEHOLDERS}} do template
  python3 gerar_fatura.py --sem-pdf             # só HTML

PDF: usa o Chrome/Chromium/Edge em modo headless (mesma renderização do navegador,
sem perda de formatação). Sem navegador, cai para o WeasyPrint (pip install weasyprint);
sem nenhum dos dois, abra o HTML no Chrome e imprima em PDF (margens: nenhuma;
escala: 100 %; sem cabeçalho/rodapé).
"""
import argparse
import calendar
import html
import json
import re
import shutil
import subprocess
import sys
import tempfile
import time
from datetime import date, timedelta
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

AQUI = Path(__file__).resolve().parent
MESES = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"]
CNPJ_EMITENTE_NUM = "99999999000191"  # CNPJ fictício com DV válido

# --------------------------------------------------------------------------- utilidades

def r2(x):
    return float(Decimal(str(x)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def fmt(v, dec=2):
    """1234.5 -> '1.234,50' (padrão brasileiro)."""
    s = f"{float(v):,.{dec}f}"
    return s.replace(",", "\x00").replace(".", ",").replace("\x00", ".")


def fmt_int(v):
    return str(int(round(float(v))))


def dmy(d):
    return d.strftime("%d/%m/%Y")


def ultimo_dia(ano, mes):
    return date(ano, mes, calendar.monthrange(ano, mes)[1])


def soma_meses(ano, mes, k):
    m = mes - 1 + k
    return ano + m // 12, m % 12 + 1


def dv_mod11(num43):
    pesos = [2, 3, 4, 5, 6, 7, 8, 9]
    s = sum(int(ch) * pesos[i % 8] for i, ch in enumerate(reversed(num43)))
    r = s % 11
    return "0" if r < 2 else str(11 - r)


def chave_acesso(emissao, nf_numero, codigo_numerico):
    # a chave reserva 9 dígitos para o nº da NF; números maiores entram truncados
    base = ("29" + emissao.strftime("%y%m") + CNPJ_EMITENTE_NUM + "66" + "000"
            + f"{int(nf_numero) % 10**9:09d}" + "1" + f"{int(codigo_numerico) % 10**8:08d}")
    assert len(base) == 43
    chave = base + dv_mod11(base)
    return " ".join(chave[i:i + 4] for i in range(0, 44, 4))


def medidor_ficticio(seq):
    """Nº de medidor fictício de 10 dígitos, um por fatura: 1 -> '9900000001'."""
    return f"99{int(seq) % 10**8:08d}"


def fator_vencimento(venc):
    """Fator de vencimento do boleto (base 22/02/2025 = 1000, após o reinício da contagem)."""
    return 1000 + (venc - date(2025, 2, 22)).days


def linha_digitavel(nf_numero, venc, valor):
    n = f"{int(nf_numero):09d}"
    valor10 = f"{int(round(valor * 100)):010d}"
    return (f"99919.09{n[0:3]} {n[3:8]}.5{n[:5]}9 00001.000019 {int(n[-1]) % 10} "
            f"{fator_vencimento(venc):04d}{valor10}")


def eixo_meses(mes_ref, ano_ref, top_mes, top_ano):
    """13 rótulos de mês/ano (4 pt) sob os gráficos da página 2, terminando no mês de referência."""
    out = []
    for k in range(13):
        a, m = soma_meses(ano_ref, mes_ref, k - 12)
        x0 = 391.5 + 14.2 * k
        out.append(f'<div class="t s4 c" style="left:{x0:.1f}pt;top:{top_mes}pt;width:14.2pt">{MESES[m - 1]}</div>')
        out.append(f'<div class="t s4 c" style="left:{x0:.1f}pt;top:{top_ano}pt;width:14.2pt">{a % 100:02d}</div>')
    return "\n  ".join(out)


# --------------------------------------------------------------------------- cálculo do cenário

def montar_contexto(cfg, cen):
    fixos = dict(cfg["fixos"])
    fixos.update(cen.get("fixos", {}))
    tar = {g: dict(v) for g, v in cfg["tarifas"].items() if isinstance(v, dict)}
    for g, v in cen.get("tarifas", {}).items():
        tar.setdefault(g, {}).update(v)
    modalidade = cen["modalidade"].upper()
    if modalidade not in ("CONVENCIONAL", "VERDE", "AZUL"):
        sys.exit(f"modalidade inválida: {modalidade}")
    grupo_a = modalidade in ("VERDE", "AZUL")

    # ---- datas do ciclo (podem ser sobrescritas no JSON)
    mes, ano = (int(x) for x in cen["mes_referencia"].split("/"))
    leit_atual = ultimo_dia(ano, mes)
    a_ant, m_ant = soma_meses(ano, mes, -1)
    a_prox, m_prox = soma_meses(ano, mes, 1)
    leit_ant = ultimo_dia(a_ant, m_ant)
    prox = ultimo_dia(a_prox, m_prox)
    def data_br(s):
        return date(*reversed([int(x) for x in s.split("/")]))

    if "leitura_atual" in cen:
        leit_atual = data_br(cen["leitura_atual"])
    if "leitura_anterior" in cen:
        leit_ant = data_br(cen["leitura_anterior"])
    if "proxima_leitura" in cen:
        prox = data_br(cen["proxima_leitura"])
    emissao = leit_atual + timedelta(days=cen.get("dias_ate_emissao", 27))
    venc = leit_atual + timedelta(days=cen.get("dias_ate_vencimento", 59))
    if "vencimento" in cen:
        venc = date(*reversed([int(x) for x in cen["vencimento"].split("/")]))
    if "data_emissao" in cen:
        emissao = date(*reversed([int(x) for x in cen["data_emissao"].split("/")]))
    dias = (leit_atual - leit_ant).days

    # ---- tributos
    pis = float(cen.get("pis", 1.03))
    cofins = float(cen.get("cofins", 4.77))
    trib = (pis + cofins) / 100.0
    aliq_icms = float(cen.get("aliquota_icms", 20.50))

    def com_trib(tarifa):
        return tarifa / (1.0 - trib)

    itens = []  # dicts: desc, qtd, tarifa, tipo ('demanda' | 'energia')

    def item(chave, qtd, tarifa, tipo):
        preco = com_trib(tarifa)
        valor = r2(qtd * preco)
        itens.append(dict(chave=chave, qtd=qtd, tarifa=tarifa, preco=preco, valor=valor,
                          piscof=r2(valor * trib), tipo=tipo))

    ctx = {}
    ctx.update({k: str(v) for k, v in fixos.items()})
    ctx["ALIQ_ICMS"] = fmt(aliq_icms)
    ctx["ALIQ_PIS"] = fmt(pis)
    ctx["ALIQ_COFINS"] = fmt(cofins)
    ctx["BANDEIRA"] = cen.get("bandeira", "VERDE").upper()
    ctx["BANDEIRA_NOME"] = cen.get("bandeira", "Verde").capitalize()
    ctx["MODALIDADE"] = modalidade.capitalize()
    ctx["SUBGRUPO"] = cen.get("subgrupo", "A4 Horo-sazonal" if grupo_a else "B3")
    ctx["MES_REFERENCIA"] = f"{mes:02d}/{ano}"
    ctx["LEITURA_ANTERIOR"] = dmy(leit_ant)
    ctx["LEITURA_ATUAL"] = dmy(leit_atual)
    ctx["PROXIMA_LEITURA"] = dmy(prox)
    ctx["NUM_DIAS"] = str(dias)
    ctx["DATA_DE"] = leit_ant.strftime("%d%m%y")
    ctx["DATA_ATE"] = leit_atual.strftime("%d%m%y")
    ctx["VENCIMENTO"] = dmy(venc)
    ctx["NF_EMISSAO"] = dmy(emissao)
    ctx["NF_SERIE"] = "000"

    leituras = dict(cfg.get("leituras_iniciais_padrao", {}))
    leituras.update(cen.get("leituras_iniciais", {}))

    if grupo_a:
        cons_p = float(cen["consumo_ponta_kwh"])
        cons_fp = float(cen["consumo_fora_ponta_kwh"])
        dem_med_p = float(cen["demanda_medida_ponta_kw"])
        dem_med_fp = float(cen["demanda_medida_fora_ponta_kw"])
        reat_p = float(cen.get("reativo_exc_ponta_kvarh", 0))
        reat_fp = float(cen.get("reativo_exc_fora_ponta_kvarh", 0))
        creat_p = float(cen.get("reativo_ponta_kvarh", 0))
        creat_fp = float(cen.get("reativo_fora_ponta_kvarh", 0))
        t = tar["grupo_a"]

        if modalidade == "VERDE":
            dem_contr = float(cen["demanda_contratada_kw"])
            # demanda a faturar: da planilha se informada; senão sem ultrapassagem
            dem_fat = float(cen.get("demanda_faturada_kw", max(dem_contr, dem_med_p, dem_med_fp)))
            item("DEM", dem_fat, t["demanda_verde"], "demanda")
            item("TUSD_P", cons_p, t["tusd_ponta_verde"], "energia")
            item("TUSD_FP", cons_fp, t["tusd_fora_ponta"], "energia")
            ctx["DEMANDA_KW"] = fmt_int(dem_contr)
        else:  # AZUL
            dem_contr_p = float(cen["demanda_contratada_ponta_kw"])
            dem_contr_fp = float(cen["demanda_contratada_fora_ponta_kw"])
            dem_fat_p = float(cen.get("demanda_faturada_ponta_kw", max(dem_contr_p, dem_med_p)))
            dem_fat_fp = float(cen.get("demanda_faturada_fora_ponta_kw", max(dem_contr_fp, dem_med_fp)))
            item("DEM_P", dem_fat_p, t["demanda_azul_ponta"], "demanda")
            item("DEM_FP", dem_fat_fp, t["demanda_azul_fora_ponta"], "demanda")
            item("TUSD_P", cons_p, t["tusd_ponta_azul"], "energia")
            item("TUSD_FP", cons_fp, t["tusd_fora_ponta"], "energia")
            ctx["DEMANDA_PONTA_KW"] = fmt_int(dem_contr_p)
            ctx["DEMANDA_FORA_PONTA_KW"] = fmt_int(dem_contr_fp)
        item("TE_P", cons_p, t["te_ponta"], "energia")
        item("TE_FP", cons_fp, t["te_fora_ponta"], "energia")
        item("REAT_P", reat_p, t["reativo_excedente"], "energia")
        item("REAT_FP", reat_fp, t["reativo_excedente"], "energia")

        # medidor / demonstrativo
        const_p, const_fp, const_dem = 0.6, 60.0, 2.4
        ctx["CONST_EA_P"], ctx["CONST_EA_FP"], ctx["CONST_DEM"] = fmt(const_p, 5), fmt(const_fp, 5), fmt(const_dem, 5)
        ctx["CONS_P_QTD"], ctx["CONS_FP_QTD"] = fmt(cons_p), fmt(cons_fp)
        ctx["REAT_P_QTD"], ctx["REAT_FP_QTD"] = fmt(reat_p), fmt(reat_fp)
        ctx["CONS_REAT_P_QTD"], ctx["CONS_REAT_FP_QTD"] = fmt(creat_p), fmt(creat_fp)
        for chave, ant, cons, const in (("EA_P", leituras["ea_p"], cons_p, const_p),
                                        ("EA_FP", leituras["ea_fp"], cons_fp, const_fp),
                                        ("REAT_P", leituras["reat_p"], creat_p, const_p),
                                        ("REAT_FP", leituras["reat_fp"], creat_fp, const_fp),
                                        ("REATEXC_P", leituras["reatexc_p"], reat_p, const_p),
                                        ("REATEXC_FP", leituras["reatexc_fp"], reat_fp, const_fp)):
            ctx[f"LEIT_ANT_{chave}"] = fmt(ant)
            ctx[f"LEIT_ATU_{chave}"] = fmt(ant + cons / const)
        ctx["LEIT_ATU_DEM_P"] = fmt(dem_med_p / const_dem)
        ctx["LEIT_ATU_DEM_FP"] = fmt(dem_med_fp / const_dem)
        ctx["DEM_MED_P"], ctx["DEM_MED_FP"] = fmt(dem_med_p), fmt(dem_med_fp)
        corr_p, corr_fp = r2(dem_med_p * 1.10), r2(dem_med_fp * 1.07)
        ctx["LEIT_ATU_DEMCORR_P"], ctx["DEMCORR_P"] = fmt(round(corr_p / const_p)), fmt(corr_p)
        ctx["LEIT_ATU_DEMCORR_FP"], ctx["DEMCORR_FP"] = fmt(round(corr_fp / const_fp)), fmt(corr_fp)
        dias_uteis = round(dias * 5 / 7)
        horas_p, horas_fp = dem_med_p * 3 * dias_uteis, dem_med_fp * (24 * dias - 3 * dias_uteis)
        ctx["FATOR_CARGA_P"] = fmt(min(1.0, cons_p / horas_p) if horas_p else 0)
        ctx["FATOR_CARGA_FP"] = fmt(min(1.0, cons_fp / horas_fp) if horas_fp else 0)
    else:
        cons = float(cen["consumo_kwh"])
        t = tar["grupo_b"]
        item("TUSD", cons, t["tusd"], "energia")
        item("TE", cons, t["te"], "energia")
        const = 1.0
        ctx["CONS_QTD"] = fmt(cons)
        ctx["CONST_EA"] = fmt(const, 5)
        ctx["LEIT_ANT_EA"] = fmt(leituras["ea"])
        ctx["LEIT_ATU_EA"] = fmt(leituras["ea"] + cons / const)

    # ---- valores por item, tributos, IRRF e total
    for it in itens:
        k = it["chave"]
        ctx[f"{k}_QTD"] = fmt(it["qtd"])
        ctx[f"{k}_PRECO"] = fmt(it["preco"], 8)
        ctx[f"{k}_VALOR"] = fmt(it["valor"])
        ctx[f"{k}_PISCOF"] = fmt(it["piscof"])
        ctx[f"{k}_TARIFA"] = fmt(it["tarifa"], 8)
    ctx["REAT_PRECO"] = fmt(com_trib(tar["grupo_a"]["reativo_excedente"]), 8)
    ctx["REAT_TARIFA"] = fmt(tar["grupo_a"]["reativo_excedente"], 8)
    base = r2(sum(it["valor"] for it in itens))
    energia = sum(it["valor"] for it in itens if it["tipo"] == "energia")
    demanda = sum(it["valor"] for it in itens if it["tipo"] == "demanda")
    irrf12, irrf48 = r2(energia * 0.012), r2(demanda * 0.048)
    total = r2(base - irrf12 - irrf48)
    ctx["BASE_PISCOF"] = fmt(base)
    ctx["VALOR_PIS"] = fmt(r2(base * pis / 100))
    ctx["VALOR_COFINS"] = fmt(r2(base * cofins / 100))
    ctx["BASE_ICMS"], ctx["VALOR_ICMS"] = fmt(0), fmt(0)
    ctx["IRRF_12"], ctx["IRRF_48"] = fmt(irrf12), fmt(irrf48)
    ctx["VALOR_TOTAL"] = fmt(total)

    # ---- identificadores fictícios da NF3e / boleto
    if "nf_numero" in cen:
        nf = int(cen["nf_numero"])
    else:
        nf = 990000001 + cfg["cenarios"].index(cen)
    ctx["NF_NUMERO"] = str(nf)
    # nº do medidor: um diferente por fatura, salvo se o cenário fixar o seu em "fixos"
    if "MEDIDOR" not in cen.get("fixos", {}):
        ctx["MEDIDOR"] = medidor_ficticio(cfg["cenarios"].index(cen) + 1)
    ctx["NF_CHAVE"] = chave_acesso(emissao, nf, cen.get("codigo_numerico", 1352197 + nf % 1000))
    ctx["NF_PROTOCOLO"] = f"3292600{nf:09d}"
    ctx["NF_PROTOCOLO_DATAHORA"] = f"{dmy(emissao)} às 23:49:05"
    ctx["DOC_PGTO"] = f"6100{nf:08d}"
    ctx["NOSSO_NUMERO"] = f"1094{nf % 100000000:08d}"
    ctx["NUM_DOCUMENTO"] = f"52{nf % 100000000:08d}"
    ctx["LINHA_DIGITAVEL"] = linha_digitavel(nf, venc, total)

    # ---- eixos dos gráficos (HTML pronto, não escapar)
    ctx["GRAF_MESES_CONSUMO"] = eixo_meses(mes, ano, 297.5, 301.5)
    ctx["GRAF_MESES_DEMANDA"] = eixo_meses(mes, ano, 476, 480)

    flags = {"VERDE": modalidade == "VERDE", "AZUL": modalidade == "AZUL",
             "CONVENCIONAL": modalidade == "CONVENCIONAL", "GRUPO_A": grupo_a,
             "MEDICAO": grupo_a}  # página de medição só existe em alta tensão (grupo A)
    return ctx, flags


# --------------------------------------------------------------------------- template

SEC_POS = re.compile(r"<!--\s*\{\{#(\w+)\}\}\s*-->(.*?)<!--\s*\{\{/\1\}\}\s*-->", re.S)
SEC_NEG = re.compile(r"<!--\s*\{\{\^(\w+)\}\}\s*-->(.*?)<!--\s*\{\{/\1\}\}\s*-->", re.S)
PLACEHOLDER = re.compile(r"\{\{(\w+)\}\}")
NAO_ESCAPAR = {"GRAF_MESES_CONSUMO", "GRAF_MESES_DEMANDA"}


def renderizar(template, ctx, flags):
    # aplica as seções até estabilizar: seções aninhadas (nomes diferentes) são
    # resolvidas uma camada por passada
    out = template
    while True:
        novo = SEC_POS.sub(lambda m: m.group(2) if flags.get(m.group(1)) else "", out)
        novo = SEC_NEG.sub(lambda m: "" if flags.get(m.group(1)) else m.group(2), novo)
        if novo == out:
            break
        out = novo

    def sub(m):
        k = m.group(1)
        if k not in ctx:
            return m.group(0)
        v = ctx[k]
        return v if k in NAO_ESCAPAR else html.escape(v, quote=False)

    out = PLACEHOLDER.sub(sub, out)
    faltando = sorted(set(PLACEHOLDER.findall(out)))
    return out, faltando


NAVEGADORES = [
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/Applications/Chromium.app/Contents/MacOS/Chromium",
    "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
    "google-chrome", "chromium", "chromium-browser", "msedge",
]


def achar_navegador():
    for cand in NAVEGADORES:
        p = cand if Path(cand).exists() else shutil.which(cand)
        if p:
            return p
    return None


def gerar_pdf_chrome(navegador, html_path, pdf_path):
    pdf_path = Path(pdf_path).resolve()
    if pdf_path.exists():
        pdf_path.unlink()
    with tempfile.TemporaryDirectory() as perfil:
        cmd = [navegador, "--headless=new", "--disable-gpu", "--no-first-run",
               "--no-default-browser-check", f"--user-data-dir={perfil}",
               # necessários em contêineres Linux (ex.: Streamlit Cloud); inofensivos no macOS
               "--no-sandbox", "--disable-dev-shm-usage",
               "--no-pdf-header-footer", f"--print-to-pdf={pdf_path}",
               Path(html_path).resolve().as_uri()]
        proc = subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        # em alguns macOS o Chrome imprime mas não encerra: espera o PDF ficar
        # estável em disco e finaliza o processo
        limite = time.time() + 60
        tamanho = -1
        while time.time() < limite:
            if proc.poll() is not None:
                break
            if pdf_path.exists():
                atual = pdf_path.stat().st_size
                if atual > 0 and atual == tamanho:
                    proc.terminate()
                    break
                tamanho = atual
            time.sleep(0.5)
        else:
            proc.kill()
        try:
            proc.wait(timeout=10)
        except subprocess.TimeoutExpired:
            proc.kill()
    if pdf_path.exists() and pdf_path.stat().st_size > 0:
        return True
    print("  ! falha no Chrome headless — tentando WeasyPrint.")
    return False


def gerar_pdf(html_path, pdf_path):
    navegador = achar_navegador()
    if navegador and gerar_pdf_chrome(navegador, html_path, pdf_path):
        return True
    try:
        from weasyprint import HTML
    except ImportError:
        if not navegador:
            print("  ! nem Chrome/Chromium/Edge nem WeasyPrint disponíveis — só o HTML foi gerado.")
        return False
    HTML(filename=str(html_path)).write_pdf(str(pdf_path))
    return True


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("cenarios", nargs="*", help="nomes (campo 'arquivo') dos cenários a gerar; vazio = todos")
    ap.add_argument("--json", default=AQUI / "cenarios.json")
    ap.add_argument("--template", default=AQUI / "template-fatura-distribuidora.html")
    ap.add_argument("--saida", default=AQUI / "saida")
    ap.add_argument("--sem-pdf", action="store_true")
    ap.add_argument("--listar-placeholders", action="store_true")
    args = ap.parse_args()

    template = Path(args.template).read_text(encoding="utf-8")
    if args.listar_placeholders:
        for p in sorted(set(PLACEHOLDER.findall(template))):
            print(p)
        return

    cfg = json.loads(Path(args.json).read_text(encoding="utf-8"))
    saida = Path(args.saida)
    saida.mkdir(parents=True, exist_ok=True)
    selecionados = [c for c in cfg["cenarios"] if not args.cenarios or c["arquivo"] in args.cenarios]
    if not selecionados:
        sys.exit("nenhum cenário encontrado com esse nome")

    for cen in selecionados:
        ctx, flags = montar_contexto(cfg, cen)
        conteudo, faltando = renderizar(template, ctx, flags)
        html_path = saida / f"{cen['arquivo']}.html"
        html_path.write_text(conteudo, encoding="utf-8")
        print(f"{cen['arquivo']}: {ctx['SUBGRUPO']} {ctx['MODALIDADE']} | ref {ctx['MES_REFERENCIA']} | "
              f"total R$ {ctx['VALOR_TOTAL']} | venc {ctx['VENCIMENTO']}")
        if faltando:
            print("  ! placeholders sem valor:", ", ".join(faltando))
        if not args.sem_pdf:
            if gerar_pdf(html_path, saida / f"{cen['arquivo']}.pdf"):
                print(f"  -> {saida / (cen['arquivo'] + '.pdf')}")


if __name__ == "__main__":
    main()
