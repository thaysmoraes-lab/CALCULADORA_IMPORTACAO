# 📦 Calculadora de Importação — Protheus

App Streamlit que recebe o **PDF do espelho da NF de importação** (TOTVS Protheus) e gera uma planilha Excel completa com a memória de cálculo, as fórmulas prontas para o Configurador de Tributos (FISA170) e a conferência contra os totais do espelho.



## 🆕 Versão 1.4

- **Correção crítica**: as fórmulas M1 e M2 geradas na aba "3. Fórmulas Protheus" agora incluem `VAL:EPIS01` e `VAL:ECOF01` no numerador (PIS e COFINS-importação) — antes esses dois tributos estavam faltando, gerando base incorreta no Configurador.
- **Robustez**: substituído `TEXT(...,"0,0000")` por uma construção matemática que não depende do locale do Excel, garantindo que o fator de redução saia sempre com vírgula como separador decimal (ex.: `0,4889`).

## 🆕 Versão 1.3

- Adicionada linha **"Divisor do ICMS (1 − alíq. cheia)"** na aba "2. Cálculo" — mostra explicitamente o `82%` (ou outro valor conforme a UF), tornando a fórmula da BASE auto-documentada (`Numerador ÷ Divisor`).

## 🆕 Versão 1.2

- Adicionada linha **"Valor unitário (R$)"** na aba "1. Inputs" (entre Quantidade e CIF), preenchida com o vUnit extraído do espelho. Útil para rastrear o preço unitário do produto sem precisar fazer a divisão manual.

## 🆕 Versão 1.1 — correções importantes

- **Parser**: corrigido o caso de quantidade + valor unitário colados no PDF (quando vUnit começa com ponto de milhar, ex.: `5.635,560`).
- **Parser**: corrigido o caso de três alíquotas coladas (Tx Sisc + Alíq ICMS + Alíq IPI) — antes só lidava com duas, perdendo itens com IPI tributado.
- **Cálculo**: AFRMM agora é rateado automaticamente por item proporcional ao CIF (modal marítimo). Vai para a base do ICMS e é descontado do "Desp Ac" do vNF para evitar dupla contagem.
- **Conferência**: rótulos das linhas atualizados para refletir o novo modelo (Sisc + AFRMM na base / Desp Ac − AFRMM fora).

## ✨ O que ele faz

1. **Extrai automaticamente** do PDF: cabeçalho do processo, DUIMP, taxa de câmbio, fornecedor, itens (descrição, NCM, quantidade, CIF, II, IPI, PIS, COFINS, Siscomex, despesas e todas as alíquotas) e os totais.
2. **Detecta cenários tributários distintos no mesmo processo**: alíquota cheia × redução de base (Conv. ICMS 52/91), COFINS-importação padrão × majorada, modal aéreo × marítimo (AFRMM).
3. **Gera Excel com 5 abas**:
   - **1. Inputs** — dados do espelho preenchidos (editáveis em laranja)
   - **2. Cálculo** — memória completa item a item, com gross-up "por dentro"
   - **3. Fórmulas Protheus** — fórmulas geradas para colar no Configurador (M1 sem redução / M2 com redução)
   - **4. Conferência** — calculado × espelho da NF com status OK / DIVERGE
   - **5. Guia** — identificadores, encadeamento do cálculo e armadilhas comuns

## 🚀 Como usar localmente

```bash
git clone https://github.com/<seu-usuario>/<seu-repo>.git
cd <seu-repo>
pip install -r requirements.txt
streamlit run app.py
```

O app abre no navegador em `http://localhost:8501`.

## ☁️ Deploy no Streamlit Cloud

1. Suba o repositório no GitHub (este código já está pronto).
2. Acesse [share.streamlit.io](https://share.streamlit.io) → **New app**.
3. Selecione o repositório, branch `main` e arquivo principal `app.py`.
4. Clique em **Deploy**.

Não há variáveis de ambiente nem segredos — o app processa tudo localmente no navegador do usuário (o PDF não sai do ambiente do Streamlit Cloud).

## 📂 Estrutura

```
.
├── app.py                       # Interface Streamlit
├── core/
│   ├── __init__.py
│   ├── parser.py                # Parser do PDF do espelho
│   ├── calculator.py            # Lógica de cálculo dos impostos
│   └── excel_builder.py         # Gerador do XLSX (5 abas)
├── .streamlit/
│   └── config.toml              # Tema (paleta TOTVS-like)
├── requirements.txt
├── .gitignore
└── README.md
```

## 🧪 Fluxo de cálculo (encadeamento natural da importação)

```
CIF (= VMLD)
  ↓
II        = CIF × alíq. II
  ↓
Base IPI  = CIF + II
  ↓
IPI       = Base IPI × alíq. IPI
  ↓
PIS       = CIF × 2,10%
COFINS    = CIF × 9,65% (ou 10,25% majorada)
  ↓
Numerador = CIF + II + IPI + PIS + COFINS + Siscomex + AFRMM + outras integrantes
  ↓
Base ICMS = Numerador ÷ (1 − alíq. interna cheia)   ← gross-up "por dentro"
  ↓
ICMS      = Base ICMS × carga efetiva
```

Quando **carga efetiva = alíquota cheia** → não há redução (Método 1 no Configurador).
Quando **carga < cheia** → há redução de base (Método 2, com multiplicador = carga ÷ cheia).

## 🛠 Fórmulas que o app gera para o Configurador (FISA170)

**Método 1 — Sem redução (alíquota cheia):**

```
( O:VAL_MERCADORIA - O:DESCONTO + O:DESPESAS + O:FRETE + O:SEGURO
  + VAL:EII001 + VAL:EIPI01 ) / ( 1 - A:AIC001 )
```

**Método 2 — Com redução de base (Conv. 52/91 etc.):**

```
( ( O:VAL_MERCADORIA - O:DESCONTO + O:DESPESAS + O:FRETE + O:SEGURO
    + VAL:EII001 ) / ( 1 - A:AIC001 ) ) * 0,4889
```

Onde `0,4889 = carga ÷ cheia` (ex.: 8,8% ÷ 18% para máquinas industriais — NCMs 8443.39.10, 8443.91.99, 8428.33.00 etc.).

## 📝 Limitações e ressalvas

- O parser foi testado em espelhos do Protheus com CFOP 3102 (importação direta). Outros CFOPs estão suportados na regex, mas convém validar.
- Algumas descrições muito longas no espelho podem quebrar em duas linhas no PDF; o parser trata isso, mas em casos extremos pode ser necessário ajustar manualmente o campo "Descrição" na tabela do app antes de gerar o Excel.
- O fator de redução de base na aba "Fórmulas Protheus" usa o Item 1 como referência. Para múltiplas reduções coexistindo no mesmo processo (ex.: alguns itens com redução para 8,8% e outros para 12%), replique a regra de base no Configurador alterando apenas o multiplicador final.
- A planilha gerada usa fórmulas Excel vivas (recalculadas ao abrir). Os valores são consistentes com cálculo manual feito a partir do espelho.

## 📬 Suporte

Sugestões e ajustes: abra uma issue no GitHub.
