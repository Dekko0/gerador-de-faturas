# Fixtures de fatura fictícia — layout DANFE NF3e (testes de OCR / UCH)

Template HTML que reproduz a geometria de um DANFE NF3e de distribuidora (3 páginas), com
coordenadas medidas em fatura real via pdfplumber, fontes Arial/Helvetica nos tamanhos originais e
camada de texto selecionável. **Toda a identidade institucional é fictícia** (precaução contra
falsificação de documento e LGPD):

| Campo | Valor fictício |
| --- | --- |
| Nome fantasia | DISTRIBUIDORA DE ENERGIA - BAHIA |
| Razão social | DISTRIBUIDORA DE ENERGIA DA BAHIA S.A. |
| CNPJ / IE | 99.999.999/0001-91 (DV válido, série inexistente) / 001234567 |
| Endereço | AV. FICTICIA, 1000, CENTRO, SALVADOR, BAHIA CEP 40000-001 |
| Site / consulta NF3e | www.deb-energia.com.br / https://consulta-nf3e.deb-energia.com.br |
| Telefones | 0800 116, 0800 000 8080, 0800 000 0155, 0800 000 7676, 0800 000 0167 |
| Banco do boleto | BANCO FICTICIO, código 999, agência/cedente 0001/00001 |
| Chave de acesso | usa o CNPJ fictício; código de barras é padrão decorativo (não codifica nada) |

## Arquivos

| Arquivo | O que é |
| --- | --- |
| `template-fatura-distribuidora.html` | Template com `{{PLACEHOLDERS}}` e seções por modalidade |
| `cenarios.json` | Dados fixos (titular, CNPJ, INST01, TESTE-0001…), tarifas e os 4 cenários UCH |
| `gerar_fatura.py` | Preenche o template, calcula os valores coerentes e gera HTML + PDF (Chrome headless; WeasyPrint como alternativa) |
| `saida/uch-base-convencional.pdf` | Primeiro cenário já gerado, para validar no OCR |

## Interface web (importação de planilha)

`app_faturas.py` é uma interface Streamlit para gerar faturas em lote a partir do
relatório `.xlsx` exportado pelo sistema de leitura de faturas (cada linha = uma fatura):

```bash
.venv/bin/streamlit run app_faturas.py
```

- As colunas de dados do sistema (A, B, C, E, G, J, K, L, V, W, AF) são descartadas;
  linhas vazias são ignoradas.
- Limite de **100 faturas por importação** (constante `MAX_FATURAS` no app); acima
  disso o app mostra erro de quantidade máxima excedida.
- As tarifas da planilha (com tributos) são convertidas para tarifa sem tributos e
  divididas em TUSD/TE na proporção do `cenarios.json`; datas de leitura, vencimento,
  emissão, nº da NF, códigos de cliente/instalação e nome da UC vêm da planilha.
- Resultado: ZIP com todas as faturas em PDF. `exemplo-20-faturas.xlsx` é uma amostra
  de 20 linhas para testar.

## Como gerar

O PDF é gerado com o Chrome/Chromium/Edge em modo headless (mesma renderização do
navegador, sem perda de formatação) — basta ter um deles instalado. Sem navegador,
o script cai para o WeasyPrint (`pip install weasyprint`).

```bash
python3 gerar_fatura.py uch-base-convencional   # só o cenário base
python3 gerar_fatura.py                          # todos os cenários de cenarios.json -> saida/
python3 gerar_fatura.py --sem-pdf                # só HTML (imprima no Chrome: margens "Nenhuma", escala 100 %)
python3 gerar_fatura.py --listar-placeholders
```

O que muda entre versões é só o bloco do cenário em `cenarios.json` (modalidade, mês, demanda,
consumo, PIS/COFINS). Titular, CNPJ, endereço, `INST01` e `TESTE-0001` ficam em `fixos`.

## Seções condicionais do template

```html
<!-- {{#VERDE}} --> … <!-- {{/VERDE}} -->                 só Verde
<!-- {{#AZUL}} --> … <!-- {{/AZUL}} -->                   só Azul
<!-- {{#CONVENCIONAL}} --> … <!-- {{/CONVENCIONAL}} -->   só grupo B
<!-- {{#GRUPO_A}} --> … <!-- {{/GRUPO_A}} -->             Verde ou Azul
```

## Placeholders principais (132 no total; `--listar-placeholders` mostra todos)

| Grupo | Placeholders |
| --- | --- |
| Identificação (fixos) | `TITULAR`, `TITULAR_LINHA2`, `CNPJ`, `ENDERECO_LINHA1..4`, `CODIGO_INSTALACAO`, `CODIGO_CLIENTE`, `CLASSE`, `TIPO_FORNECIMENTO`, `CONTA_CONTRATO_COLETIVA`, `BANCO` |
| Contrato (variam) | `SUBGRUPO` + `MODALIDADE` (ex.: "A4 Horo-sazonal" + "Verde", "B3" + "Convencional"), `DEMANDA_KW` (Verde), `DEMANDA_PONTA_KW` / `DEMANDA_FORA_PONTA_KW` (Azul) |
| Ciclo | `MES_REFERENCIA`, `VENCIMENTO`, `VALOR_TOTAL`, `LEITURA_ANTERIOR`, `LEITURA_ATUAL`, `NUM_DIAS`, `PROXIMA_LEITURA`, `DATA_DE`, `DATA_ATE` |
| NF3e | `NF_NUMERO`, `NF_SERIE`, `NF_EMISSAO`, `NF_CHAVE`, `NF_PROTOCOLO`, `NF_PROTOCOLO_DATAHORA`, `DOC_PGTO` |
| Itens grupo A | `DEM_*` (Verde), `DEM_P_*` / `DEM_FP_*` (Azul), `TUSD_P_*`, `TUSD_FP_*`, `TE_P_*`, `TE_FP_*`, `REAT_P_*`, `REAT_FP_*`, `CONS_P_QTD`, `CONS_FP_QTD`, `IRRF_12`, `IRRF_48` — sufixos `_QTD`, `_PRECO`, `_VALOR`, `_PISCOF`, `_TARIFA` |
| Itens grupo B | `CONS_QTD`, `TUSD_*`, `TE_*`, `IRRF_12` |
| Tributos | `BASE_PISCOF`, `ALIQ_PIS`, `VALOR_PIS`, `ALIQ_COFINS`, `VALOR_COFINS`, `BASE_ICMS`, `ALIQ_ICMS`, `VALOR_ICMS`, `BANDEIRA`, `BANDEIRA_NOME` |
| Medidor / página 2 | `MEDIDOR` (um por fatura: 9900000001, 9900000002…; fixe um em `fixos` do cenário se precisar), `LEIT_ANT_*`, `LEIT_ATU_*`, `CONST_*`, `DEM_MED_P/FP`, `DEMCORR_P/FP`, `CONS_REAT_*`, `FATOR_CARGA_P/FP`, `GRAF_MESES_CONSUMO/DEMANDA` (HTML gerado pelo script) |
| Boleto | `LINHA_DIGITAVEL`, `NOSSO_NUMERO`, `NUM_DOCUMENTO` |

## Cenários em `cenarios.json`

| Arquivo | Classificação | Demanda contratada | Ref. |
| --- | --- | --- | --- |
| uch-base-convencional | B3 Convencional | — | 05/2026 |
| uch10-mudanca-verde-105kw | A4 Horo-sazonal Verde | 105 | 06/2026 |
| uch09-mudanca-azul-ponta-forponta | A4 Horo-sazonal Azul | 80 ponta / 120 fora ponta | 07/2026 |
| uch02-mudanca-demanda-145-para-105 | A4 Horo-sazonal Verde | 105 (seed com 145) | 08/2026 |

## O que NÃO foi reproduzido com fidelidade (revisar manualmente)

1. **Identidade da distribuidora 100 % fictícia.** Se o parser/OCR identifica a concessionária
   pelo literal "COELBA"/"NEOENERGIA", ele não vai reconhecer estes PDFs — aponte a detecção de
   distribuidora do sandbox para "DISTRIBUIDORA DE ENERGIA" ou desative essa checagem. A linha
   digitável usa banco 999 (inexistente) e dígitos não bancários; validadores de boleto vão
   rejeitá-la — proposital.
2. **Layout de grupo B.** A referência é A4 Verde; o cenário Convencional usa a mesma estrutura
   com `CLASSIFICAÇÃO: B3 Convencional`, itens `Consumo-TUSD` / `Consumo-TE`, sem demanda, quadro
   GRANDEZAS CONTRATADAS vazio e coluna POSTOS HORÁRIOS em branco. Confira com uma fatura B3 real.
3. **Rótulos da modalidade Azul** (`Demanda Ativa Ponta` / `Demanda Ativa Fora Ponta`,
   `Demanda Contratada Ponta` / `Demanda Contratada Fora Ponta`) são inferidos do padrão da fatura
   Verde — confirmar com uma fatura Azul real.
4. **Logotipo** substituído pelo texto da marca fictícia em duas linhas; QR codes (chave de acesso
   e PIX) não reproduzidos; código de barras é padrão decorativo; gráficos da página 2 têm barras
   fixas (só os meses do eixo acompanham o cenário).
5. **CNPJ do titular** impresso completo (`35.343.832/0001-08`); na fatura real ele sai mascarado
   (`13.927.***/****-**`). Se o OCR espera máscara, ajuste `CNPJ` em `fixos`.
6. `INST01` e `TESTE-0001` não são numéricos (real: 8 e 10 dígitos). Se o extrator usa regex
   numérica para código da instalação/cliente, esses valores não vão bater.
7. Classe `PODER PUBLICO -- MUNICIPAL`, ICMS zerado e linhas `TRIBF-IRRF` mantidos como na
   referência, apesar do titular fictício ser "LTDA".
8. Tarifas do grupo A copiadas da referência (fev/2026); demanda ponta Azul (62,50) e tarifas B3
   são estimativas. PIS/COFINS variam por cenário para parecer real.
9. Na fatura real, leituras ≥ 100.000 aparecem sem separador de milhar na página 1
   (`290649,00`); os cenários usam leituras menores, então a peculiaridade não aparece.
10. Fonte: Liberation Sans no lugar da Helvetica — larguras iguais em dígitos; textos longos
    deslocam até ~4 pt no fim da linha.
