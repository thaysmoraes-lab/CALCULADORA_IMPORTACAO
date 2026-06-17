"""
Parser do espelho da Nota Fiscal de Importação (Protheus).

Extrai dados estruturados a partir do PDF do espelho da NF de importação,
gerada pela rotina de Documento de Entrada do TOTVS Protheus.

Saída esperada do método principal:
    {
        "cabecalho": {
            "processo": str,
            "duimp": str,
            "fornecedor": str,
            "modal": str,
            "taxa_cambio": float,
            "desembaraco": str | None,
        },
        "itens": [
            {
                "codigo": str,
                "descricao": str,
                "ncm": str,             # com pontuação (8443.99.23)
                "cfop": str,
                "item": int,
                "qtd": float,
                "vunit": float,
                "vtot": float,
                "cif": float,
                "vii": float,
                "vipi": float,
                "vpis": float,
                "vcofins": float,
                "desp_ac": float,       # despesas acessórias (frequentemente NÃO entra na base do ICMS)
                "tx_sisc": float,       # Siscomex rateada (entra na base do ICMS)
                "bc_icms": float,
                "vicms": float,
                "aliq_icms": float,     # carga efetiva (decimal: 0.18 = 18%)
                "aliq_ipi": float,
                "aliq_cofins": float,
                "aliq_pis": float,
                "aliq_ii": float,
            },
            ...
        ],
        "totais": {
            "vprod": float,
            "vcif": float,
            "vipi": float,
            "vpis": float,
            "vcofins": float,
            "bc_icms": float,
            "vicms": float,
            "vfrete": float,
            "outras_despesas": float,
            "afrmm": float,
            "sisc_total": float,
            "vii_total": float,
            "vnf": float,
        },
    }
"""

import re
from dataclasses import dataclass

import pdfplumber


# Regex auxiliares
_NUM = r"[\d\.]+,\d+"  # ex.: 1.234,56 / 0,00 / 130,7092
_NCM_CFOP = re.compile(r"\b(\d{8})\s+(3102|3120|3122|3127|3130|3551|3556)\s+(\d+)\s+")
_TRIBUTOS_LINE = re.compile(
    r"TRIBUTOS:\s*II\s*R\$:\s*([\d\.,]+)\s*/\s*IPI\s*R\$:\s*([\d\.,]+)\s*/\s*"
    r"PIS\s*R\$:\s*([\d\.,]+)\s*/\s*COFINS\s*R\$:\s*([\d\.,]+)\s*/\s*"
    r"TAXA\s+SISCOMEX\s*R\$:\s*([\d\.,]+)\s*/\s*ICMS\s*R\$:\s*([\d\.,]+)\s*/\s*"
    r"THC\s*R\$:\s*([\d\.,]+)\s*/\s*AFRMM\s*R\$:\s*([\d\.,]+)\s*/\s*"
    r"TAXA\s+CAMBIAL\s+USD:\s*([\d\.,]+)",
    re.IGNORECASE,
)
_DUIMP_LINE = re.compile(r"DUIMP:\s*([0-9A-Za-z\-/]+)\s*DE\s*(\d{2}/\d{2}/\d{4})", re.IGNORECASE)
_PROC_LINE = re.compile(r"PROCESSO:\s*([0-9A-Za-z\-/]+)", re.IGNORECASE)
_DESEMB_LINE = re.compile(r"DESEMBARACO:\s*([^/\n]+)", re.IGNORECASE)


def _br_num(s: str) -> float:
    """Converte número em formato brasileiro (1.234,56) para float."""
    if not s:
        return 0.0
    s = s.strip().replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _format_ncm(ncm8: str) -> str:
    """8443.99.23 a partir de 84439923."""
    if not ncm8 or len(ncm8) != 8:
        return ncm8 or ""
    return f"{ncm8[0:4]}.{ncm8[4:6]}.{ncm8[6:8]}"


def _join_broken_item_lines(text: str) -> list[str]:
    """
    Junta linhas quebradas dos itens. No espelho do Protheus, a descrição
    pode quebrar em duas linhas (ex.: 'CARTUCHO BASE SOLVENTE\\nVERMELHA').
    Estratégia: qualquer linha que comece (após número/espaço) com algo que
    não seja NCM+CFOP, mas que continue uma linha de item iniciada acima,
    é juntada à anterior.

    Algoritmo: percorre as linhas; quando uma linha contém o padrão 'NCM CFOP item',
    e a linha imediatamente anterior parece ser uma continuação de descrição
    (sem números ao final), juntamos as duas.
    """
    raw_lines = [ln.rstrip() for ln in text.split("\n")]
    out = []
    i = 0
    while i < len(raw_lines):
        ln = raw_lines[i]
        m = _NCM_CFOP.search(ln)
        if m:
            # Verifica se o início da linha (antes do NCM) é só palavras (sem código de produto)
            prefix = ln[: m.start()].strip()
            # Linha bem formada de item começa com o código do produto (ex.: "910 CARTUCHO...")
            # Se o prefixo não começa com dígitos, provavelmente a descrição quebrou
            starts_with_code = bool(re.match(r"^\d", prefix))
            if not starts_with_code and out:
                # Juntamos com a linha anterior (que tinha o código + parte da descrição)
                out[-1] = out[-1].rstrip() + " " + ln.lstrip()
            else:
                out.append(ln)
        else:
            out.append(ln)
        i += 1
    return out


def _parse_item_line(ln: str) -> dict | None:
    """
    Faz o parsing de uma linha de item já consolidada.
    Formato (após NCM e CFOP):
      <codigo> <descricao> <NCM> <CFOP> <item> <qtd> <vunit> <vtot>
      <%red> <bc_icms> <vicms> <bc_icms_st> <vicms_st>
      <bc_ipi> <vipi> <bc_pis_cof> <vpis> <vcofins>
      <cif> <vii> <desp_ac> <tx_sisc><aliq_icms> <aliq_ipi> <aliq_cof> <aliq_pis> <aliq_ii>

    A particularidade do espelho do Protheus é que tx_sisc e aliq_icms vêm
    coladas (ex.: '122,5118,00') — precisamos quebrar manualmente.
    """
    m = _NCM_CFOP.search(ln)
    if not m:
        return None
    ncm8, cfop, item_num = m.group(1), m.group(2), m.group(3)

    # Tudo antes do NCM: código + descrição
    cabec = ln[: m.start()].strip()
    parts = cabec.split(None, 1)  # quebra na primeira palavra
    if len(parts) == 2:
        codigo, descricao = parts
    else:
        codigo, descricao = (parts[0] if parts else ""), ""
    descricao = descricao.strip()

    # Tudo depois do CFOP+item: valores numéricos
    resto = ln[m.end():].strip()

    # ── Correção 1: qtd e vUnit podem vir colados quando o vUnit começa
    # com ponto de milhar (ex.: "11,000005.635,560" = qtd 11,00000 + vUnit 5.635,560).
    # A quantidade no espelho do Protheus sempre tem 5 casas decimais,
    # então identificamos o padrão ",ddddd<dígito>" e inserimos espaço.
    resto = re.sub(r"(,\d{5})(\d)", r"\1 \2", resto)

    # ── Correção 2: As últimas colunas do espelho (Tx Sisc + Alíq ICMS + Alíq IPI +
    # Alíq COFINS + Alíq PIS + Alíq II) podem vir várias colunas coladas
    # (ex.: "135,1818,0015,00" → Tx Sisc 135,18 + Alíq ICMS 18,00 + Alíq IPI 15,00).
    # Aplicamos a separação repetidamente até estabilizar: qualquer ",dd" seguido
    # imediatamente por "<1-2 dígitos>,dd" recebe um espaço de separação.
    while True:
        novo = re.sub(r"(,\d{2})(\d{1,2},\d{2})", r"\1 \2", resto)
        if novo == resto:
            break
        resto = novo

    # Agora extraímos todos os números
    nums = re.findall(_NUM, resto)
    # Esperado para importação (sem ICMS ST): 20 números
    # qtd, vunit, vtot, %red, bc_icms, vicms, bc_ipi, vipi, bc_pis_cof, vpis, vcofins,
    # cif, vii, desp_ac, tx_sisc, aliq_icms, aliq_ipi, aliq_cof, aliq_pis, aliq_ii
    if len(nums) < 20:
        return None  # linha mal formada

    # Detectar se há ICMS ST (raro mas possível) → 22 nums
    tem_st = (len(nums) >= 22)
    offset = 2 if tem_st else 0

    qtd       = _br_num(nums[0])
    vunit     = _br_num(nums[1])
    vtot      = _br_num(nums[2])
    # nums[3] = %red base
    bc_icms   = _br_num(nums[4])
    vicms     = _br_num(nums[5])
    # se tem_st: nums[6], nums[7] = ICMS ST
    bc_ipi    = _br_num(nums[6 + offset])
    vipi      = _br_num(nums[7 + offset])
    # nums[8+offset] = bc pis/cof
    vpis      = _br_num(nums[9 + offset])
    vcofins   = _br_num(nums[10 + offset])
    cif       = _br_num(nums[11 + offset])
    vii       = _br_num(nums[12 + offset])
    desp_ac   = _br_num(nums[13 + offset])
    tx_sisc   = _br_num(nums[14 + offset])
    aliq_icms = _br_num(nums[15 + offset]) / 100.0
    aliq_ipi  = _br_num(nums[16 + offset]) / 100.0
    aliq_cof  = _br_num(nums[17 + offset]) / 100.0
    aliq_pis  = _br_num(nums[18 + offset]) / 100.0
    aliq_ii   = _br_num(nums[19 + offset]) / 100.0

    return {
        "codigo": codigo,
        "descricao": descricao,
        "ncm": _format_ncm(ncm8),
        "cfop": cfop,
        "item": int(item_num),
        "qtd": qtd,
        "vunit": vunit,
        "vtot": vtot,
        "cif": cif,
        "vii": vii,
        "vipi": vipi,
        "vpis": vpis,
        "vcofins": vcofins,
        "desp_ac": desp_ac,
        "tx_sisc": tx_sisc,
        "bc_icms": bc_icms,
        "vicms": vicms,
        "aliq_icms": aliq_icms,
        "aliq_ipi": aliq_ipi,
        "aliq_cofins": aliq_cof,
        "aliq_pis": aliq_pis,
        "aliq_ii": aliq_ii,
    }


def _parse_cabecalho(text: str) -> dict:
    """Extrai os dados dos 'Dados Adicionais' (rodapé): processo, DUIMP, etc."""
    cab = {
        "processo": "",
        "duimp": "",
        "fornecedor": "",
        "modal": "AÉREO",  # default; usuário ajusta
        "taxa_cambio": 0.0,
        "desembaraco": None,
    }

    m = _PROC_LINE.search(text)
    if m:
        cab["processo"] = m.group(1).strip()

    m = _DUIMP_LINE.search(text)
    if m:
        cab["duimp"] = m.group(1).strip()

    m = _DESEMB_LINE.search(text)
    if m:
        desemb = m.group(1).strip()
        cab["desembaraco"] = desemb
        # Inferir modal pelo local de desembaraço
        up = desemb.upper()
        if "AEROPORTO" in up or "VIRACOPOS" in up or "GUARULHOS" in up:
            cab["modal"] = "AÉREO"
        elif "PORTO" in up or "SANTOS" in up or "ITAJAÍ" in up or "PARANAGUÁ" in up:
            cab["modal"] = "MARÍTIMO"

    m = _TRIBUTOS_LINE.search(text)
    if m:
        cab["taxa_cambio"] = _br_num(m.group(9))

    # Fornecedor: linha que começa com "Nome/Razao" + valor — pego a primeira
    for ln in text.split("\n"):
        if ln.startswith("Nome/Razao "):
            # remove o label e o que vier depois de "CNPJ"
            val = ln.replace("Nome/Razao ", "", 1)
            val = re.split(r"\s+CNPJ\b", val)[0].strip()
            if val:
                cab["fornecedor"] = val
                break

    return cab


def _parse_totais(text: str) -> dict:
    """
    Extrai os totais do quadro de rodapé do espelho.
    Estratégia: usa a linha 'TRIBUTOS:' do bloco "DADOS ADICIONAIS", que
    contém todos os valores de forma estruturada e confiável.
    """
    totais = dict.fromkeys(
        ["vprod", "vcif", "vipi", "vpis", "vcofins", "bc_icms", "vicms",
         "vfrete", "outras_despesas", "afrmm", "sisc_total", "vii_total", "vnf"],
        0.0,
    )

    m = _TRIBUTOS_LINE.search(text)
    if m:
        totais["vii_total"] = _br_num(m.group(1))
        totais["vipi"] = _br_num(m.group(2))
        totais["vpis"] = _br_num(m.group(3))
        totais["vcofins"] = _br_num(m.group(4))
        totais["sisc_total"] = _br_num(m.group(5))
        totais["vicms"] = _br_num(m.group(6))
        totais["afrmm"] = _br_num(m.group(8))

    # Vl Total Nota, Vl Total Produtos, VI Total CIF, Base ICMS, Valor Frete, Outras Despesas
    # — esses aparecem em blocos com label seguido de valor na linha de baixo.
    # Vou procurar por padrões específicos do espelho.
    # Linha tipo: "14.034,87 2.526,28 0,00 0,00 9.873,36 207,34 9.873,36 0,00 A.F.R.M.M. (R$) 0,00"
    # contém: Base ICMS, Valor ICMS, BC ST, V ST, Base PIS/COFINS, V PIS, Vl Total Produtos, Valor Seguro
    # A próxima linha tem: Valor Frete, Frete Territ. Nac., Outras Despesas, Vl Total CIF, Vlr Tot Base IPI, Vlr Tot IPI, Valor COFINS, Vl Total Nota

    lines = text.split("\n")
    for i, ln in enumerate(lines):
        if "Base ICMS Valor ICMS" in ln and i + 1 < len(lines):
            nums = re.findall(_NUM, lines[i + 1])
            if len(nums) >= 7:
                if not totais["bc_icms"]: totais["bc_icms"] = _br_num(nums[0])
                if not totais["vicms"]:   totais["vicms"]   = _br_num(nums[1])
                if not totais["vprod"]:   totais["vprod"]   = _br_num(nums[6])
        if "Valor Frete" in ln and "Vl Total Nota" in ln and i + 1 < len(lines):
            nums = re.findall(_NUM, lines[i + 1])
            if len(nums) >= 8:
                totais["vfrete"] = _br_num(nums[0])
                totais["outras_despesas"] = _br_num(nums[2])
                totais["vcif"] = _br_num(nums[3])
                # nums[4] = Vlr Tot Base IPI
                if not totais["vipi"]:    totais["vipi"]    = _br_num(nums[5])
                if not totais["vcofins"]: totais["vcofins"] = _br_num(nums[6])
                totais["vnf"] = _br_num(nums[7])

    return totais


def parse_pdf(pdf_path_or_bytes) -> dict:
    """
    Ponto de entrada principal: recebe path para PDF ou bytes/file-like
    e retorna o dicionário estruturado.
    """
    if isinstance(pdf_path_or_bytes, (str, bytes)):
        opener = pdfplumber.open(pdf_path_or_bytes)
    else:
        # file-like (BytesIO ou UploadedFile do Streamlit)
        opener = pdfplumber.open(pdf_path_or_bytes)

    with opener as pdf:
        text = "\n".join((p.extract_text() or "") for p in pdf.pages)

    consolidated_lines = _join_broken_item_lines(text)
    itens = []
    for ln in consolidated_lines:
        item = _parse_item_line(ln)
        if item:
            itens.append(item)

    cabecalho = _parse_cabecalho(text)
    totais = _parse_totais(text)

    # ── Rateio do AFRMM por item proporcional ao CIF (modal marítimo)
    # O espelho mostra AFRMM apenas no total; precisamos distribuir nos itens
    # para que entre corretamente na base do ICMS de cada um.
    # No espelho do Protheus, o AFRMM rateado JÁ está embutido em "Desp Ac"
    # do item — por isso, para o vNF não ser duplicado, subtraímos AFRMM
    # do Desp Ac quando compomos "outras_sem_icms".
    afrmm_total = totais.get("afrmm", 0.0)
    cif_total = sum(it["cif"] for it in itens) or 1.0
    for it in itens:
        if afrmm_total > 0:
            it["afrmm_item"] = round(afrmm_total * it["cif"] / cif_total, 2)
        else:
            it["afrmm_item"] = 0.0
        # Outras despesas integrantes da base do ICMS (Siscomex + AFRMM rateado)
        it["outras_base_icms"] = round(it["tx_sisc"] + it["afrmm_item"], 2)
        # Outras despesas SEM influência no ICMS (Desp Ac líquido de AFRMM,
        # já que AFRMM está embutido no Desp Ac do espelho do Protheus)
        it["outras_sem_icms"] = round(max(0.0, it["desp_ac"] - it["afrmm_item"]), 2)

    return {
        "cabecalho": cabecalho,
        "itens": itens,
        "totais": totais,
        "_texto_bruto": text,  # para debug
    }
