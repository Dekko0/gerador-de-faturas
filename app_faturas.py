#!/usr/bin/env python3
"""
app_faturas.py — interface Streamlit para gerar faturas fictícias em lote.

O usuário importa o relatório .xlsx exportado pelo sistema de leitura de faturas
(cada linha = uma fatura). O app descarta as colunas de dados do sistema
(A, B, C, E, G, J, K, L, V, W, AF), mapeia as demais para os campos do gerador
(gerar_fatura.py + template-fatura-distribuidora.html) e exporta um ZIP com
todas as faturas em PDF.

Rodar:  streamlit run app_faturas.py
"""
import io
import re
import tempfile
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime
from pathlib import Path

import openpyxl
import streamlit as st

import gerar_fatura as gf

AQUI = Path(__file__).resolve().parent

# ----------------------------------------------------------------- configuração
MAX_FATURAS = 100  # limite de faturas por importação
COLUNAS_IGNORADAS = {"A", "B", "C", "E", "G", "J", "K", "L", "V", "W", "AF"}
PDF_WORKERS = 4    # instâncias de Chrome headless em paralelo


# ----------------------------------------------------------------- leitura do xlsx

def letra_coluna(i):
    """0 -> A, 1 -> B, ..., 26 -> AA (índice de coluna em letra)."""
    letra = ""
    i += 1
    while i:
        i, resto = divmod(i - 1, 26)
        letra = chr(65 + resto) + letra
    return letra


def ler_planilha(arquivo_bytes):
    """Lê o .xlsx e devolve (linhas, colunas_descartadas). Cada linha é um dict
    cabeçalho -> valor só com as colunas consideradas; linhas vazias são ignoradas."""
    wb = openpyxl.load_workbook(io.BytesIO(arquivo_bytes), read_only=True, data_only=True)
    ws = wb.active
    iterador = ws.iter_rows(values_only=True)
    try:
        cabecalho = next(iterador)
    except StopIteration:
        return [], []

    consideradas, descartadas = [], []
    for i, nome in enumerate(cabecalho):
        (descartadas if letra_coluna(i) in COLUNAS_IGNORADAS else consideradas).append((i, nome))
    descartadas = [f"{letra_coluna(i)} ({nome})" for i, nome in descartadas]

    linhas = []
    for valores in iterador:
        linha = {}
        vazia = True
        for i, nome in consideradas:
            v = valores[i] if i < len(valores) else None
            if v not in (None, ""):
                vazia = False
            linha[nome] = v
        if not vazia:
            linhas.append(linha)
    wb.close()
    return linhas, descartadas


# ----------------------------------------------------------------- linha -> cenário

def num(v, padrao=0.0):
    """Converte para float; devolve `padrao` para vazio/erro de fórmula ('#NUM!' etc.)."""
    try:
        return float(v)
    except (TypeError, ValueError):
        return padrao


def data_br(v):
    """datetime/date/texto -> 'dd/mm/aaaa' (ou None)."""
    if isinstance(v, (datetime, date)):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, str) and re.match(r"\d{2}/\d{2}/\d{4}$", v.strip()):
        return v.strip()
    return None


def dividir_tarifa(total_sem_trib, parte_a, parte_b):
    """Divide uma tarifa combinada (ex.: consumo = TUSD + TE) na proporção das
    tarifas de referência do cenarios.json."""
    soma = parte_a + parte_b
    if soma <= 0 or total_sem_trib <= 0:
        return None
    return total_sem_trib * parte_a / soma, total_sem_trib * parte_b / soma


def linha_para_cenario(linha, cfg, seq):
    """Mapeia uma linha da planilha para um cenário no formato de cenarios.json."""
    classe = str(linha.get("nome_classe") or linha.get("descricao_classe") or "").lower()
    if "azul" in classe:
        modalidade = "AZUL"
    elif "verde" in classe:
        modalidade = "VERDE"
    else:
        modalidade = "CONVENCIONAL"

    ref = linha.get("periodo_referencia")
    if not isinstance(ref, (datetime, date)):
        raise ValueError("período de referência ausente ou inválido")

    pis = num(linha.get("aliquota_pis"), 0.0122) * 100
    cofins = num(linha.get("aliquota_cofins"), 0.0562) * 100
    trib = (pis + cofins) / 100.0

    def tarifa_sem_trib(chave):
        """Tarifa da planilha (com tributos) -> tarifa sem tributos (coluna TARIFA UNIT)."""
        t = num(linha.get(chave))
        return t * (1.0 - trib) if t > 0 else 0.0

    cen = {
        "arquivo": f"fatura-{seq:04d}",
        "modalidade": modalidade,
        "mes_referencia": f"{ref.month:02d}/{ref.year}",
        "pis": round(pis, 2),
        "cofins": round(cofins, 2),
        "aliquota_icms": round(num(linha.get("aliquota_icms"), 0.205) * 100, 2),
    }

    nome_anexo = str(linha.get("arquivo_anexo") or "").strip()
    if nome_anexo:
        base = re.sub(r"\.pdf$", "", nome_anexo, flags=re.I)
        base = re.sub(r"[^\w.-]+", "_", base)[:80]
        if base:
            cen["arquivo"] = f"{seq:04d}-{base}"

    for chave_cen, coluna in (("data_emissao", "data_emissao"), ("vencimento", "data_vencimento"),
                              ("leitura_atual", "leitura_atual_fatura"),
                              ("leitura_anterior", "ultima_leitura_fatura"),
                              ("proxima_leitura", "proxima_leitura_fatura")):
        d = data_br(linha.get(coluna))
        if d:
            cen[chave_cen] = d

    nf = re.sub(r"\D", "", str(linha.get("numero_nota_fiscal") or ""))
    cen["nf_numero"] = int(nf) if nf else 990000000 + seq

    # dados do titular / instalação que vêm da planilha; o medidor é um por fatura
    fixos = {"MEDIDOR": gf.medidor_ficticio(seq)}
    if linha.get("codigo_instalacao_uc") or linha.get("codigo_instalacao"):
        fixos["CODIGO_INSTALACAO"] = str(linha.get("codigo_instalacao_uc") or linha.get("codigo_instalacao"))
    if linha.get("codigo_do_cliente_uc") or linha.get("codigo_cliente"):
        fixos["CODIGO_CLIENTE"] = str(linha.get("codigo_do_cliente_uc") or linha.get("codigo_cliente"))
    if linha.get("tipo_fornecimento"):
        fixos["TIPO_FORNECIMENTO"] = str(linha["tipo_fornecimento"])
    if linha.get("nome_uc"):
        fixos["TITULAR_LINHA2"] = str(linha["nome_uc"])
    descricao = str(linha.get("descricao_classe") or "").strip()
    if descricao:
        fixos["CLASSE"] = re.sub(r"^(B\d+|A\d+(\s+Horo-sazonal)?(\s+\w+)?)\s+", "", descricao)
    cen["fixos"] = fixos

    # tarifas da planilha (com tributos) -> tarifas sem tributos do gerador
    ref_a, ref_b = cfg["tarifas"]["grupo_a"], cfg["tarifas"]["grupo_b"]
    tarifas = {"grupo_a": {}, "grupo_b": {}}
    if modalidade == "CONVENCIONAL":
        cen["consumo_kwh"] = num(linha.get("consumo_faturado_total")) or num(linha.get("consumo_total"))
        partes = dividir_tarifa(tarifa_sem_trib("tarifa_consumo_convencional"), ref_b["tusd"], ref_b["te"])
        if partes:
            tarifas["grupo_b"]["tusd"], tarifas["grupo_b"]["te"] = partes
    else:
        cen["consumo_ponta_kwh"] = num(linha.get("consumo_faturado_ponta")) or num(linha.get("consumo_ponta"))
        cen["consumo_fora_ponta_kwh"] = (num(linha.get("consumo_faturado_fora_ponta"))
                                         or num(linha.get("consumo_fora_ponta")))
        cen["demanda_medida_ponta_kw"] = num(linha.get("demanda_maxima_ponta"))
        cen["demanda_medida_fora_ponta_kw"] = num(linha.get("demanda_maxima_fora_ponta"))
        for chave_cen, coluna in (("demanda_faturada_kw", "demanda_faturada"),
                                  ("demanda_faturada_ponta_kw", "demanda_faturada_ponta"),
                                  ("demanda_faturada_fora_ponta_kw", "demanda_faturada_fora_ponta")):
            if linha.get(coluna) is not None:
                cen[chave_cen] = num(linha[coluna])
        cen["reativo_ponta_kvarh"] = num(linha.get("consumo_reativo_ponta"))
        cen["reativo_fora_ponta_kvarh"] = num(linha.get("consumo_reativo_fora_ponta"))
        cen["reativo_exc_ponta_kvarh"] = num(linha.get("consumo_reativo_excedente_ponta"))
        cen["reativo_exc_fora_ponta_kvarh"] = num(linha.get("consumo_reativo_excedente_fora_ponta"))
        chave_tusd_p = "tusd_ponta_verde" if modalidade == "VERDE" else "tusd_ponta_azul"
        partes = dividir_tarifa(tarifa_sem_trib("tarifa_consumo_ponta"), ref_a[chave_tusd_p], ref_a["te_ponta"])
        if partes:
            tarifas["grupo_a"][chave_tusd_p], tarifas["grupo_a"]["te_ponta"] = partes
        partes = dividir_tarifa(tarifa_sem_trib("tarifa_consumo_fora_ponta"),
                                ref_a["tusd_fora_ponta"], ref_a["te_fora_ponta"])
        if partes:
            tarifas["grupo_a"]["tusd_fora_ponta"], tarifas["grupo_a"]["te_fora_ponta"] = partes
        t_reat = tarifa_sem_trib("tarifa_consumo_reativo_excedente")
        if t_reat:
            tarifas["grupo_a"]["reativo_excedente"] = t_reat
        if modalidade == "VERDE":
            cen["demanda_contratada_kw"] = num(linha.get("demanda_contratada"))
            t_dem = tarifa_sem_trib("tarifa_demanda_maxima")
            if t_dem:
                tarifas["grupo_a"]["demanda_verde"] = t_dem
        else:  # AZUL
            cen["demanda_contratada_ponta_kw"] = num(linha.get("demanda_contratada_ponta"))
            cen["demanda_contratada_fora_ponta_kw"] = num(linha.get("demanda_contratada_fora_ponta"))
            t_dem_p = tarifa_sem_trib("tarifa_demanda_maxima_ponta")
            t_dem_fp = tarifa_sem_trib("tarifa_demanda_maxima_fora_ponta")
            if t_dem_p:
                tarifas["grupo_a"]["demanda_azul_ponta"] = t_dem_p
            if t_dem_fp:
                tarifas["grupo_a"]["demanda_azul_fora_ponta"] = t_dem_fp
    tarifas = {g: v for g, v in tarifas.items() if v}
    if tarifas:
        cen["tarifas"] = tarifas

    # calibra o total da fatura para bater com valor_total_em_reais da planilha
    alvo = linha.get("valor_total_em_reais")
    if alvo is not None:
        cen["_total_planilha"] = round(num(alvo), 2)
        cen["_divergencia"] = round(calibrar_total(cen, cfg, cen["_total_planilha"]), 2)
    return cen


# chaves de tarifa de energia por modalidade (escaladas na calibração do total;
# as tarifas de demanda ficam fiéis à planilha)
CHAVES_ENERGIA = {
    "CONVENCIONAL": [("grupo_b", "tusd"), ("grupo_b", "te")],
    "VERDE": [("grupo_a", "tusd_ponta_verde"), ("grupo_a", "tusd_fora_ponta"),
              ("grupo_a", "te_ponta"), ("grupo_a", "te_fora_ponta"), ("grupo_a", "reativo_excedente")],
    "AZUL": [("grupo_a", "tusd_ponta_azul"), ("grupo_a", "tusd_fora_ponta"),
             ("grupo_a", "te_ponta"), ("grupo_a", "te_fora_ponta"), ("grupo_a", "reativo_excedente")],
}


def total_calculado(cfg, cen):
    ctx, _ = gf.montar_contexto(cfg, cen)
    return float(ctx["VALOR_TOTAL"].replace(".", "").replace(",", "."))


def _secante(cen, cfg, alvo, chaves):
    """Escala as tarifas indicadas por um fator comum até o total bater com o alvo
    (método da secante). Devolve o total obtido."""
    tarifas = cen.setdefault("tarifas", {})
    base = {ch: tarifas.get(ch[0], {}).get(ch[1], cfg["tarifas"][ch[0]][ch[1]]) for ch in chaves}

    def aplicar(k):
        for (grupo, chave), v in base.items():
            tarifas.setdefault(grupo, {})[chave] = v * k
        return total_calculado(cfg, cen)

    k0, t0 = 1.0, aplicar(1.0)
    if abs(t0 - alvo) <= 0.005 or t0 <= 0:
        return t0
    k1 = max(alvo / t0, 0.0)
    t1 = aplicar(k1)
    for _ in range(8):
        if abs(t1 - alvo) <= 0.005 or t1 == t0:
            break
        k2 = max(k1 + (alvo - t1) * (k1 - k0) / (t1 - t0), 0.0)
        k0, t0, k1 = k1, t1, k2
        t1 = aplicar(k1)
    return t1


def calibrar_total(cen, cfg, alvo):
    """Ajusta as tarifas de energia da fatura até o total calculado bater com o
    da planilha (a planilha embute componentes que não são reconstruíveis, como
    acréscimo de bandeira). Devolve a divergência residual (0 = bateu)."""
    total = _secante(cen, cfg, alvo, CHAVES_ENERGIA[cen["modalidade"]])
    if abs(total - alvo) > 0.005:
        # refino: só a tarifa TE do posto com consumo, para andar centavo a centavo
        if cen["modalidade"] == "CONVENCIONAL":
            refino = ("grupo_b", "te") if cen.get("consumo_kwh") else None
        elif cen.get("consumo_fora_ponta_kwh"):
            refino = ("grupo_a", "te_fora_ponta")
        elif cen.get("consumo_ponta_kwh"):
            refino = ("grupo_a", "te_ponta")
        else:
            refino = None
        if refino:
            total = _secante(cen, cfg, alvo, [refino])
    return total - alvo


# ----------------------------------------------------------------- geração

def gerar_uma(template, cfg, cen, navegador, pasta):
    ctx, flags = gf.montar_contexto(cfg, cen)
    conteudo, faltando = gf.renderizar(template, ctx, flags)
    html_path = pasta / f"{cen['arquivo']}.html"
    pdf_path = pasta / f"{cen['arquivo']}.pdf"
    html_path.write_text(conteudo, encoding="utf-8")
    if not gf.gerar_pdf_chrome(navegador, html_path, pdf_path):
        raise RuntimeError("falha ao converter para PDF no Chrome headless")
    return pdf_path, ctx


# ----------------------------------------------------------------- interface

def main():
    st.set_page_config(page_title="Gerador de Faturas Fictícias", page_icon="", layout="wide")
    st.title("Gerador de Faturas Fictícias")
    st.caption("Importe o relatório de faturas (.xlsx) do sistema — cada linha vira uma fatura "
               "fictícia em PDF (layout DANFE NF3e), entregues juntas em um ZIP.")

    with st.expander("Como funciona / colunas descartadas"):
        st.markdown(f"""
    1. Importe o relatório `.xlsx` exportado pelo sistema de leitura de faturas Poupenergia.
    2. As colunas de dados do sistema (**{", ".join(sorted(COLUNAS_IGNORADAS, key=lambda c: (len(c), c)))}**)
       são descartadas automaticamente — apenas os dados da fatura são usados.
    3. Cada linha preenchida vira uma fatura fictícia (Convencional, Verde ou Azul,
       conforme a classe).
    4. Ao final, baixe o **ZIP com todos os PDFs**.

    Limite: **{MAX_FATURAS} faturas por importação**.
    """)

    navegador = gf.achar_navegador()
    if not navegador:
        st.error("Chrome/Chromium/Edge não encontrado — necessário para gerar os PDFs.")
        st.stop()

    arquivo = st.file_uploader("Relatório de faturas (.xlsx)", type=["xlsx"])
    if not arquivo:
        st.stop()

    try:
        linhas, descartadas = ler_planilha(arquivo.getvalue())
    except Exception as e:
        st.error(f"Não foi possível ler a planilha: {e}")
        st.stop()

    if not linhas:
        st.error("A planilha não tem nenhuma linha de fatura preenchida.")
        st.stop()

    if len(linhas) > MAX_FATURAS:
        st.error(f"Quantidade máxima de faturas excedida: o arquivo tem **{len(linhas)}** faturas "
                 f"e o limite é **{MAX_FATURAS}** por importação. "
                 "Divida a planilha em arquivos menores e importe novamente.")
        st.stop()

    cfg = gf.json.loads((AQUI / "cenarios.json").read_text(encoding="utf-8"))
    template = (AQUI / "template-fatura-distribuidora.html").read_text(encoding="utf-8")

    cenarios, erros_mapeamento = [], []
    for n, linha in enumerate(linhas, start=1):
        try:
            cenarios.append(linha_para_cenario(linha, cfg, n))
        except Exception as e:
            erros_mapeamento.append(f"linha {n + 1} da planilha: {e or type(e).__name__}")

    st.success(f"{len(linhas)} fatura(s) encontrada(s) na planilha "
               f"({len(descartadas)} colunas de sistema descartadas).")
    if erros_mapeamento:
        st.warning("Linhas ignoradas por dados incompletos:\n\n- " + "\n- ".join(erros_mapeamento))

    st.dataframe(
        [{"arquivo": c["arquivo"], "modalidade": c["modalidade"], "referência": c["mes_referencia"],
          "vencimento": c.get("vencimento", "—"), "nº NF": c["nf_numero"],
          "total R$ (planilha)": c.get("_total_planilha", "—")} for c in cenarios],
        use_container_width=True, height=280,
    )
    divergentes = [c for c in cenarios if abs(c.get("_divergencia", 0)) > 0.009]
    if divergentes:
        st.warning("Faturas cujo total não pôde ser calibrado para o valor da planilha:\n\n- "
                   + "\n- ".join(f"{c['arquivo']}: diferença de R$ {c['_divergencia']:.2f}"
                                 for c in divergentes))

    if st.button(f"Gerar {len(cenarios)} fatura(s) em PDF", type="primary", disabled=not cenarios):
        barra = st.progress(0.0, text="Gerando faturas…")
        erros_geracao = []
        zip_buffer = io.BytesIO()
        with tempfile.TemporaryDirectory() as tmp, \
                zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as z:
            pasta = Path(tmp)
            feitos = 0
            with ThreadPoolExecutor(max_workers=PDF_WORKERS) as pool:
                tarefas = {pool.submit(gerar_uma, template, cfg, c, navegador, pasta): c for c in cenarios}
                for tarefa in as_completed(tarefas):
                    cen = tarefas[tarefa]
                    try:
                        pdf_path, _ = tarefa.result()
                        z.write(pdf_path, pdf_path.name)
                    except Exception as e:
                        erros_geracao.append(f"{cen['arquivo']}: {e}")
                    feitos += 1
                    barra.progress(feitos / len(cenarios), text=f"Gerando faturas… {feitos}/{len(cenarios)}")
        barra.empty()
        geradas = len(cenarios) - len(erros_geracao)
        if erros_geracao:
            st.warning("Faturas que falharam:\n\n- " + "\n- ".join(erros_geracao))
        if geradas:
            st.success(f"{geradas} fatura(s) gerada(s).")
            st.session_state["zip_faturas"] = zip_buffer.getvalue()
        else:
            st.error("Nenhuma fatura pôde ser gerada.")

    if st.session_state.get("zip_faturas"):
        st.download_button("⬇️ Baixar ZIP com as faturas em PDF", st.session_state["zip_faturas"],
                           file_name="faturas-ficticias.zip", mime="application/zip")


if __name__ == "__main__":
    main()
