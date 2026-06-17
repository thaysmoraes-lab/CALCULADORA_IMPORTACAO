"""
Replica a lógica de cálculo dos impostos de importação que está na planilha
Excel gerada, usada para conferência rápida no app antes da geração do arquivo.

Para cada item:
  CIF (= VMLD)
  II         = CIF × alíq_II
  Base IPI   = CIF + II
  IPI        = Base IPI × alíq_IPI
  PIS        = CIF × alíq_PIS
  COFINS     = CIF × alíq_COFINS
  Numerador  = CIF + II + IPI + PIS + COFINS + Outras desp. integrantes
  Base ICMS  = Numerador ÷ (1 − alíq_ICMS_cheia)        ← gross-up "por dentro"
  ICMS       = Base ICMS × carga_efetiva
"""

from dataclasses import dataclass


@dataclass
class ItemCalc:
    cif: float
    ii: float
    base_ipi: float
    ipi: float
    pis: float
    cofins: float
    outras_base_icms: float        # Sisc + AFRMM + outras integrantes
    outras_sem_icms: float         # vOutro que não entra na base do ICMS
    numerador: float
    base_icms: float
    icms_destacado: float          # base × carga efetiva (= ICMS a recolher)
    icms_devido_portal: float      # base × alíquota cheia (informativo do extrato)
    fator_reducao: float           # 1 − carga/cheia
    vprod: float                   # = CIF + II
    vnf: float                     # = vProd + outras_base + outras_sem + IPI + ICMS


def calcular_item(
    cif: float,
    aliq_ii: float,
    aliq_ipi: float,
    aliq_pis: float,
    aliq_cofins: float,
    aliq_icms_cheia: float,
    carga_efetiva_icms: float,
    outras_base_icms: float,
    outras_sem_icms: float = 0.0,
) -> ItemCalc:
    """Calcula todos os tributos e a base do ICMS para um item."""
    ii       = round(cif * aliq_ii, 2)
    base_ipi = cif + ii
    ipi      = round(base_ipi * aliq_ipi, 2)
    pis      = round(cif * aliq_pis, 2)
    cofins   = round(cif * aliq_cofins, 2)

    numerador = cif + ii + ipi + pis + cofins + outras_base_icms

    if aliq_icms_cheia >= 1.0 or aliq_icms_cheia <= 0:
        base_icms = 0.0
    else:
        base_icms = round(numerador / (1 - aliq_icms_cheia), 2)

    icms_destacado    = round(base_icms * carga_efetiva_icms, 2)
    icms_devido_portal = round(base_icms * aliq_icms_cheia, 2)

    if aliq_icms_cheia > 0:
        fator_reducao = 1 - (carga_efetiva_icms / aliq_icms_cheia)
    else:
        fator_reducao = 0.0

    vprod = cif + ii
    vnf   = round(vprod + outras_base_icms + outras_sem_icms + ipi + icms_destacado, 2)

    return ItemCalc(
        cif=cif, ii=ii, base_ipi=base_ipi, ipi=ipi, pis=pis, cofins=cofins,
        outras_base_icms=outras_base_icms, outras_sem_icms=outras_sem_icms,
        numerador=numerador, base_icms=base_icms,
        icms_destacado=icms_destacado, icms_devido_portal=icms_devido_portal,
        fator_reducao=fator_reducao, vprod=vprod, vnf=vnf,
    )


def calcular_todos(itens_input: list[dict]) -> list[dict]:
    """
    Recebe a lista de itens (no formato do parser) e devolve a mesma lista com
    os campos calculados anexados, prontos para preencher o template.

    Espera, por item:
      cif, aliq_ii, aliq_ipi, aliq_pis, aliq_cofins, aliq_icms (= carga efetiva),
      tx_sisc (Siscomex), e opcionalmente:
        outras_base_icms (default = tx_sisc)
        outras_sem_icms (default = desp_ac)
        aliq_icms_cheia (default = 0.18)
    """
    saida = []
    for it in itens_input:
        carga = it.get("aliq_icms", 0.0)
        cheia = it.get("aliq_icms_cheia", 0.18)
        outras_base = it.get("outras_base_icms", it.get("tx_sisc", 0.0))
        outras_sem  = it.get("outras_sem_icms", it.get("desp_ac", 0.0))

        c = calcular_item(
            cif=it["cif"],
            aliq_ii=it.get("aliq_ii", 0.0),
            aliq_ipi=it.get("aliq_ipi", 0.0),
            aliq_pis=it.get("aliq_pis", 0.021),
            aliq_cofins=it.get("aliq_cofins", 0.0965),
            aliq_icms_cheia=cheia,
            carga_efetiva_icms=carga,
            outras_base_icms=outras_base,
            outras_sem_icms=outras_sem,
        )
        saida.append({**it, "calc": c.__dict__, "aliq_icms_cheia": cheia,
                      "outras_base_icms": outras_base, "outras_sem_icms": outras_sem})
    return saida


def totais_calculados(itens_calc: list[dict]) -> dict:
    """Soma os totais a partir dos itens já calculados."""
    if not itens_calc:
        return {}
    chaves = ["cif", "ii", "ipi", "pis", "cofins",
              "outras_base_icms", "outras_sem_icms",
              "base_icms", "icms_destacado", "icms_devido_portal",
              "vprod", "vnf"]
    tot = {k: 0.0 for k in chaves}
    for it in itens_calc:
        c = it["calc"]
        for k in chaves:
            tot[k] += c.get(k, 0.0)
    return {k: round(v, 2) for k, v in tot.items()}
