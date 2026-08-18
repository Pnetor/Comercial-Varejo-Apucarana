"""
Atualiza o index.html do Estoque Apucarana a partir da Planilha Google
("banco de dados" do estoque), publicada na web como CSV.

Requer 1 GitHub Secret:
  GSHEET_CSV_URL -> link de "Publicar na Web" da aba de estoque, em CSV
                     (Planilha Google -> Arquivo -> Compartilhar ->
                      Publicar na Web -> escolher a aba -> formato CSV)

Este script pode ser disparado de duas formas (ambas já configuradas no
workflow):
  1) De hora em hora (agendado), como rede de segurança
  2) Instantaneamente, via "repository_dispatch", quando o Apps Script da
     planilha avisa o GitHub assim que alguém edita uma célula

A planilha precisa ter as mesmas colunas do Excel original, na mesma ordem:
CODIGO, TP, GRUPO, Descricao, U.M., FL, ARMZ, SALDO_EM_ESTOQUE,
EMPENHO_PARA_REQ_PV_RESERVA, ESTOQUE_DISPONIVEL, DT_ULT_MOV
(cabeçalho na linha 2, dados a partir da linha 3 - igual ao Excel).
"""

import os
import re
import csv
import json
import sys
from datetime import datetime, timedelta, timezone

import requests
from io import StringIO

CSV_URL = os.environ["GSHEET_CSV_URL"]
VALIDADES_CSV_URL = os.environ.get("GSHEET_VALIDADES_CSV_URL")
PRECOS_CSV_URL = os.environ.get("GSHEET_PRECOS_CSV_URL")
PEDIDOS_CSV_URL = os.environ.get("GSHEET_PEDIDOS_CSV_URL")

REPO_INDEX_PATH = "index.html"
REPO_PRECOS_PATH = "tabela-precos.html"
REPO_PEDIDOS_PATH = "pedidos-em-aberto.html"

# ── Nomes de exibição fixos por código ──
# Quando um card precisa ser (re)criado (item saiu e voltou pra planilha,
# ou é genuinamente novo), o nome mostrado vem DAQUI se o código estiver
# na lista - nunca da descrição bruta do ERP (que é longa/inconsistente).
# Só cai para a descrição do ERP quando o código não está mapeado aqui.
NOME_CANONICO = {
    "FRINCANMI000002": "Frango Inteiro",
    "FRASCANMI000037": "Asa",
    "FRASCANMI000027": "Coxinha",
    "FRCSCANMI000018": "Coxa e Sobrecoxa",
    "FRCSCANMI000016": "Sobrecoxa",
    "FRPTCANMI000065": "Filé de Peito",
    "FRPTCANMI000018": "Peito com osso",
    "FRMDCANMI000020": "Coração",
    "FRMDCANMI000008": "Moela",
    "FRCSCANMI000051": "Sobrecoxa",
    "FRPTCANMI000062": "Meio Filé",
    "FRPTCANMI000061": "Sassami",
    "FRASCANMI000047": "Coxinha",
    "FRASCANMI000048": "Meio da Asa",
    "FRCSCANMI000052": "Coxa Pilão",
    "FRMDBELME000007": "Fígado- Bandeja- BELLAVES",
    "FRMDBELME000008": "Moela- Bandeja- BELLAVES",
    "FRINCANMI000017": "Frango Carcaça Temperado",
    "FRINCANMI000003": "Frango inteiro",
    "FRINCANMI000004": "Frango temperado 18kg",
    "FRINCANMI000005": "Frango temperado 20kg",
    "FRINCANMI000006": "Frango temperado PV",
    "FRINMIFMI000009": "Carcaça 1,6 Mister",
    "FRINCANMI000011": "Carcaça 1,650/ 1,750",
    "FRINCANMI000012": "Carcaça 1,750/ 1,850",
    "FRINCANMI000013": "Carcaça 1,850/ 1,950",
    "FRINCANMI000014": "Carcaça 1,950/ 2,050",
    "FRINCANMI000015": "Carcaça 2,050/ 2,150",
    "FRINCANMI000016": "Carcaça 2,150/ 2,250",
    "COMPRO000010": "Galinha Inteira Congelada 22Kg",
    "COMPRO000011": "Galo Inteiro Congelado 22Kg",
    "COMPRO000083": "Galinha Inteira Congelada 20Kg",
    "COMPRO000084": "Galo Inteiro Congelado 20Kg",
    "FRASCANMI000039": "Asa- Interfolhada",
    "FRASCANMI000028": "Asa",
    "FRASCANMI000001": "Asa 18Kg",
    "FRCSCANMI000040": "Coxinha da Asa Cong. Bandeja 1kg",
    "FRASCANMI000041": "Coxinha da asa- Interfolhada",
    "FRASCANMI000029": "Coxinha de asa",
    "FRASCANMI000043": "Meio da asa c/osso) Pct 10 KG",
    "FRASCANMI000040": "Meio da asa - Interfolhado",
    "FRASCANMI000005": "Meio da Asa Interfolhado 18Kg",
    "FRASCANMI000007": "Meio da asa - Individual",
    "FRCSCANMI000006": "Coxa Pilão - Interfolhada",
    "FRCSCANMI000039": "Coxa Pilão Cong. Bandeja 1kg",
    "FRCSCANMI000041": "Sobrecoxa Cong. Bandeja 1kg",
    "FRCSCANMI000002": "Coxa e sobrecoxa- Interfolhada",
    "FRCSCANME000013": "Coxa e sobrecoxa- Interfolhado ME",
    "FRCSCANME000024": "Coxa e Sobrecoxa Interfolhado ME 15kg",
    "FRCSCANMI000001": "Coxa e sobrecoxa 18Kg",
    "FRCSCANMI000003": "Coxa e sobrecoxa",
    "FRCSCANMI000010": "Sobrecoxa IQF 6 Kg",
    "FRCSCANMI000019": "Filé de Coxa c/ Sobrecoxa Pct 2kg",
    "FRCSCANMI000053": "Filé de Coxa c/ Sobrecoxa Pct 1kg",
    "FRCSCANMI000027": "Coxa e sobrecoxa a passarinho Pct 10kg",
    "FRCSCANMI000023": "Coxa dessosada 800g",
    "FRCSCANMI000021": "Coxa dorsal",
    "FRCSCANMI000034": "Filé de coxa",
    "FRCSCANMI000005": "Sobrecoxa - Interfolhada",
    "FRCSCANME000001": "Filé de Sobrecoxa s/ Pele - Pct 2kg",
    "FRPTCANMI000033": "Filé de Peito - Individual",
    "FRPTCANMI000060": "Meio Filé - Pct 1Kg",
    "FRPTCANMI000053": "Filé de Peito \"Grade B\" - Pct 10kg",
    "FRPTCANMI000027": "Filé de Peito - Interfolhado",
    "FRPTCANMI000064": "Filé de Peito - Individual",
    "FRPTCANMI000008": "Meio Filé - Interfolhado",
    "FRPTCANMI000051": "Meio Filé - Interfolhado.",
    "FRPTCANME000004": "Meio Filé ME",
    "FRPTCANMI000059": "Meio Peito Cong. Bandeja 1kg",
    "FRPTCANMI000004": "Peito com osso",
    "FRPTCANMI000050": "Peito- Interfolhado",
    "FRPTCANMI000058": "Sassami Cong. Bandeja 1Kg",
    "FRPTCANMI000026": "Sassami - Interfolhado",
    "FRPTCANMI000043": "Sassami - Individual",
    "FRPTCANMI000046": "Sassami - Pct 1kg",
    "FRPTCANMI000041": "Sassami 6kg",
    "FRPTCANMI000048": "Sassami 800g",
    "FRPTCANMI000052": "Sassami \"Grade B\"- Pct 10kg",
    "FRMDCANMI000015": "Coração 1kg",
    "FRMDCANMI000029": "Coração 2kg",
    "FRMDCANMI000034": "Coração Cong. Bandeja 600g",
    "FRMDCANMI000016": "Fígado 1kg",
    "FRMDCANME000003": "Figado 1kg ME",
    "COMPRO000089": "Moela PCT 1 Kg Cx 18kg Mister Frango",
    "FRMDCANMI000022": "Moela 1kg",
    "FRPEMIFME000003": "Pés Grade Pct 7,5",
    "FRPECANME000003": "Pés Pct 7,5",
    "FRPECANMI000002": "Pés Carijo Pct 7,5",
    "COMPRO000088": "Pés Mister Frango 7,5kg",
    "FRMDCANMI000030": "Pescoço 1kg",
    "FRSACANMI000001": "Sambiquira 1kg",
    "FRSAMIFMI000001": "Sambiquira-Mister",
    "FRMRCANMI000002": "Recorte Filé Peito + Filé Coxa",
    "FRMSCANMI000005": "CMS",
    "FRMSCANME000006": "CMS Congelada - Bloco",
    "FRMSCANMI000008": "CMS- Bloco",
    "FRCSCANMI000009": "Coxa Pilão IQF 1kg",
    "FRCSCANMI000032": "Coxa e Sobrecoxa a Passarinho IQF 800g",
    "FRCSCANMI000059": "Coxa e Sobrecoxa a Passarinho IQF 1Kg",
    "FRASCANMI000022": "Coxinha IQF 800g",
    "FRASCANMI000010": "Coxinha IQF 1kg",
    "FRASCANMI000008": "Coxinha IQF 2kg",
    "FRASCANMI000023": "Meio da Asa IQF 800g",
    "FRASCANMI000011": "Meio da Asa IQF 1kg",
    "FRPTCANMI000054": "Filé de Peito em Bifes IQF 800g",
    "FRPTCANMI000069": "Filé de peito em Bifes IQF 1Kg",
    "FRPTCANMI000021": "Meio Filé IQF 1kg",
    "FRPTCANMI000024": "Sassami IQF 800g",
    "FRPTCANMI000010": "Sassami IQF 1kg",
    "FRPTCANMI000013": "Sassami IQF 6kg",
    "FRCSCANMI000055": "Filé de Sobrecoxa IQF 1Kg",
    "FRCSCANMI000011": "Sobrecoxa IQF 1kg",
    "PEFLCANMI000002": "Tilápia 400g",
    "PEFLCANMI000001": "Tilápia 800g",
    "PEFLCANMI000011": "Tilápia em pedaços",
    "PEFLCANMI000013": "Isca de Tilápia 400g",
    "PEFLCANMI000012": "Filé Tilápia Empanado 600g",
    "COMPRO000008": "Sassami Empanado 700g",
    "COMPRO000064": "Sassami Empanado 1kg",
    "COMPRO000065": "Chicken 300g",
    "COMPRO000066": "Chicken 1kg",
    "COMPRO000099": "Fut Chicken 300g",
    "COMPRO000100": "Fut Chicken 1kg",
    "COMPRO000121": "Coxinha da Asa empanada tradicional (Fut Wings)",
    "COMPRO000123": "Coxinha da Asa empanada apimentada (Fut Wings)",
    "COMPRO000074": "Anel de cebola 6,6kg",
    "COMPRO000009": "Anel de cebola 11kg",
    "COMPRO000068": "Batata 400g",
    "COMPRO000069": "Batata 1,1kg",
    "COMPRO000067": "Batata 2kg",
    "COMPRO000098": "Batata Noisette 400G",
    "COMPRO000003": "Mandioca",
    "COMPRO000081": "Pão de Queijo 300g",
    "COMPRO000082": "Pão de Queijo 800g",
    "COMPRO000059": "Polenta",
    "COMPRO000094": "Lasanha Bolonhesa 600g",
    "COMPRO000095": "Lasanha Quatro Queijos 600g",
    "COMPRO000096": "Lasanha Frango 600g",
    "COMPRO000097": "Bolinho de Tilápia 300g",
    "COMPRO000101": "Small Fish 400g",
    "COMPRO000102": "Small Fish 1,5kg",
    "COMPRO000116": "Fut Burguer c/ Molho de Picles",
    "COMPRO000117": "Fut Burguer Bacon c/ Requeijão",
    "COMPRO000118": "Fut Burguer Maionese Grill",
    "COMPRO000119": "Fut Chicken c/ Recheio de Queijo",
}


# ── Código Canção fixo por código ──
# Mesma lógica do NOME_CANONICO: quando um card precisa ser (re)criado,
# o Código Canção vem daqui (não fica '—') se o código estiver mapeado.
COD_CANCAO_FIXO = {
    "FRASCANMI000047": "91396",
    "FRASCANMI000048": "91402",
    "FRCSCANMI000052": "91401",
    "FRCSCANMI000051": "91395",
    "FRPTCANMI000062": "91397",
    "FRPTCANMI000061": "91399",
    "FRINCANMI000002": "69964",
    "FRASCANMI000037": "91321",
    "FRASCANMI000027": "91192",
    "FRCSCANMI000018": "84374",
    "FRPTCANMI000065": "91424",
    "FRCSCANMI000016": "71315",
    "FRPTCANMI000018": "84376",
    "FRMDCANMI000020": "91247",
    "FRMDCANMI000008": "71313",
    "FRINCANMI000017": "91417",
    "COMPRO000083": "91357",
    "COMPRO000010": "86590",
    "FRINCANMI000004": "91077",
    "FRINCANMI000003": "69963",
    "FRINCANMI000005": "91078",
    "FRINCANMI000006": "91110",
    "COMPRO000084": "91359",
    "COMPRO000011": "86591",
    "FRINCANMI000011": "91346",
    "FRINCANMI000012": "91347",
    "FRINCANMI000013": "91348",
    "FRINCANMI000014": "91349",
    "FRINCANMI000015": "91350",
    "FRINCANMI000016": "91351",
    "FRASCANMI000029": "91308",
    "FRASCANMI000040": "91315",
    "FRASCANMI000043": "91391",
    "FRASCANMI000028": "91172",
    "FRASCANMI000007": "10420",
    "FRCSCANMI000021": "27441",
    "FRCSCANMI000003": "2364",
    "FRCSCANMI000027": "91404",
    "FRCSCANMI000006": "8143",
    "FRCSCANMI000023": "91030",
    "FRCSCANMI000001": "2359",
    "FRCSCANMI000053": "91429",
    "FRCSCANMI000019": "79563",
    "FRCSCANME000001": "12719",
    "FRCSCANMI000005": "2398",
    "FRPTCANMI000053": "91386",
    "FRPTCANMI000033": "91194",
    "FRPTCANMI000060": "91427",
    "FRPTCANMI000004": "2392",
    "FRMRCANMI000002": "54176",
    "FRPTCANMI000043": "91309",
    "FRPTCANMI000052": "91385",
    "FRPTCANMI000064": "91425",
    "FRPTCANMI000027": "91339",
    "FRPTCANMI000051": "91372",
    "FRPTCANME000004": "15776",
    "FRPTCANMI000041": "91300",
    "FRPTCANMI000046": "91317",
    "FRMSCANMI000005": "64076",
    "FRMDCANMI000015": "91263",
    "FRMDCANMI000016": "91254",
    "FRMDCANMI000022": "91289",
    "FRPEMIFME000003": "91272",
    "FRMDCANMI000030": "91299",
    "FRSACANMI000001": "70142",
    "FRMSCANME000006": "91389",
    "FRMSCANMI000008": "69069",
    "FRMDCANMI000029": "91296",
    "FRMDCANMI000003": "2369",
    "FRPECANMI000002": "91392",
    "FRPECANME000003": "91294",
    "FRMDCANMI000034": "91428",
    "FRCSCANMI000039": "91400",
    "FRCSCANMI000040": "91432",
    "FRPTCANMI000059": "91434",
    "FRPTCANMI000058": "91439",
    "FRCSCANMI000041": "91440",
    "FRCSCANMI000059": "91436",
    "FRCSCANMI000009": "62160",
    "FRASCANMI000011": "62158",
    "FRASCANMI000010": "62157",
    "FRPTCANMI000054": "91382",
    "FRPTCANMI000069": "91444",
    "FRPTCANMI000021": "88846",
    "FRPTCANMI000010": "62159",
    "FRCSCANMI000055": "91438",
    "FRCSCANMI000011": "64347",
    "FRCSCANMI000032": "91335",
    "FRASCANMI000023": "91319",
    "FRPTCANMI000013": "64346",
    "FRCSCANMI000010": "64345",
    "COMPRO000074": "91373",
    "COMPRO000009": "67983",
    "COMPRO000068": "91329",
    "COMPRO000069": "91330",
    "COMPRO000067": "91328",
    "COMPRO000098": "91408",
    "COMPRO000065": "91326",
    "COMPRO000066": "91327",
    "COMPRO000099": "91409",
    "COMPRO000100": "91410",
    "COMPRO000008": "67772",
    "COMPRO000064": "91325",
    "COMPRO000094": "91415",
    "COMPRO000096": "91414",
    "COMPRO000095": "91416",
    "COMPRO000003": "53245",
    "COMPRO000059": "91268",
    "COMPRO000081": "91375",
    "COMPRO000082": "91374",
    "PEFLCANMI000002": "91128",
    "PEFLCANMI000001": "91127",
    "PEFLCANMI000011": "91288",
    "PEFLCANMI000013": "91353",
    "PEFLCANMI000012": "91352",
    "COMPRO000097": "91413",
    "COMPRO000101": "91411",
    "COMPRO000102": "91412",
    "COMPRO000116": "91452",
    "COMPRO000117": "91453",
    "COMPRO000118": "91454",
    "COMPRO000119": "91455",
    "COMPRO000121": "91450",
    "COMPRO000123": "91451",
}


# ── Onde criar um card novo quando o código não existe ainda no HTML ──
# Prefixos sem ambiguidade: sempre vão para o mesmo lugar, não importa o texto.
# Não existe prefixo de código que garanta a seção sozinho - até COMPRO
# tem itens de Alimentos Preparados (ex: bolinhos, empanados) E itens de
# Congelados (ex: galinha/galo inteiro). A classificação é sempre feita
# lendo a descrição do produto, em classificar_produto_novo() abaixo.


def _sem_acento(txt):
    import unicodedata
    return "".join(
        c for c in unicodedata.normalize("NFD", txt)
        if unicodedata.category(c) != "Mn"
    )


def classificar_produto_novo(nome_produto):
    """Decide em qual seção/grupo do painel um produto novo deve entrar,
    lendo palavras-chave do nome/descrição do ERP. A descrição é o único
    sinal confiável - o prefixo/código NÃO indica a seção (ex: COMPRO tem
    tanto "Galinha Inteira Congelada" -> Congelados/Inteiros quanto
    "Sassami Empanado" -> Alimentos Preparados).
    Retorna (secao, grupo_label) ou None se não conseguir decidir com
    segurança (nesse caso o script só avisa no log, não cria o card)."""
    nome = _sem_acento(nome_produto or "").upper()

    # Alimentos Preparados: itens processados/prontos, mesmo que a
    # descrição também diga "CONGELADO" (ex: batata congelada, empanados).
    # Esse sinal tem prioridade sobre o corte de frango, porque um "Fut
    # Chicken" ou uma "Lasanha" não são pedaços de frango in natura.
    sinais_alimentos = (
        "EMPANAD", "CHICKEN", "BATATA", "MANDIOCA", "POLENTA",
        "ANEL DE CEBOLA", "PAO DE QUEIJO", "LASANHA", "TILAPIA", "PEIXE",
        "FISH", "BURGUER", "BOLINHO", "WINGS", "NUGGET", "ISCA DE",
        "HAMBURGUER", "STEAK", "NOISETTE",
    )
    if any(k in nome for k in sinais_alimentos):
        return ("alimentos", "ALIMENTOS")

    if "IQF" in nome:
        return ("iqf", "IQF")

    if "RESFR" in nome:
        grupo = "BANDEJA" if ("BANDEJA" in nome or "BDJ" in nome) else "PACOTE"
        return ("resfriados", grupo)

    if "CONG" in nome:
        eh_bandeja = "BANDEJA" in nome or "BDJ" in nome
        if eh_bandeja:
            return ("congelados", "GRUPO BANDEJA CONGELADA")
        if any(k in nome for k in ("INTEIR", "GALINHA", "GALO", "CARCACA", "TEMPERAD")):
            return ("congelados", "GRUPO INTEIROS")
        if "ASA" in nome:
            return ("congelados", "GRUPO ASA")
        if any(k in nome for k in ("COXA", "SOBRECOXA", "PERNA", "DORSAL", "LEG QUARTER")):
            return ("congelados", "GRUPO PERNA")
        if any(k in nome for k in ("PEITO", "FILE", "FILEZINHO", "SASSAMI")):
            return ("congelados", "GRUPO PEITO")
        if any(k in nome for k in ("CORACAO", "FIGADO", "MOELA", "PESCOCO", "CMS",
                                    "PES ", "SAMBIQUIRA", "MIUDO", "CARTILAGEM")):
            return ("congelados", "GRUPO MIUDOS")
        # é congelado mas o corte especifico nao bateu com nenhuma palavra-chave
        return None

    return None



def baixar_csv(url):
    resp = requests.get(url)
    resp.raise_for_status()
    resp.encoding = "utf-8"
    return resp.text


def carregar_validades(csv_text):
    """Lê a aba de validades (Cód. Prod. Cliente, Peso Líquido, Validade)
    e monta o mesmo formato usado no VALIDADES do HTML:
    { codigo: { entries: [{d, p, v}], hv: bool } }
    """
    reader = csv.reader(StringIO(csv_text))
    rows = list(reader)
    hoje = datetime.now(timezone.utc) - timedelta(hours=3)
    hoje = hoje.replace(hour=0, minute=0, second=0, microsecond=0, tzinfo=None)

    validades = {}
    for r in rows[2:]:
        if not r or not r[0]:
            continue
        codigo = r[0].strip()
        try:
            peso = float(str(r[1]).replace(".", "").replace(",", ".")) if r[1] else 0
        except (ValueError, IndexError):
            peso = 0
        data_str = r[2].strip() if len(r) > 2 else ""
        if not data_str:
            continue
        try:
            data_venc = datetime.strptime(data_str, "%d/%m/%Y")
        except ValueError:
            continue
        vencido = data_venc < hoje
        entry = {"d": data_str, "p": peso, "v": vencido}
        validades.setdefault(codigo, {"entries": [], "hv": False})
        validades[codigo]["entries"].append(entry)
        if vencido:
            validades[codigo]["hv"] = True
    return validades


def carregar_saldos(csv_text):
    """Retorna tanto os saldos quanto os dados brutos (nome, tp) de cada
    código, para permitir criar cards novos quando necessário."""
    reader = csv.reader(StringIO(csv_text))
    rows = list(reader)
    data = {}
    linhas_brutas = {}
    # linha 0 = título da planilha, linha 1 = cabeçalho, dados a partir da linha 2
    for r in rows[2:]:
        if not r or not r[0]:
            continue
        codigo = r[0].strip()
        try:
            disponivel = float(str(r[9]).replace(".", "").replace(",", ".")) if r[9] else 0
        except (ValueError, IndexError):
            disponivel = 0
        data[codigo] = disponivel
        linhas_brutas[codigo] = r
    return data, linhas_brutas


def fmt_kg(v):
    if v is None:
        v = 0
    neg = v < 0
    v = abs(v)
    if v == int(v):
        s = f"{int(v):,}".replace(",", ".")
    else:
        s = f"{v:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return ("-" if neg else "") + s + " kg"


def slug_nome(nome):
    """Nome do produto em Title Case simples, a partir da Descricao do ERP
    (que normalmente vem toda em maiúsculas)."""
    nome = nome.strip()
    palavras_minusculas = {"de", "da", "do", "das", "dos", "e", "c/", "s/", "a", "com"}
    partes = nome.split(" ")
    out = []
    for i, p in enumerate(partes):
        if i > 0 and p.lower() in palavras_minusculas:
            out.append(p.lower())
        else:
            out.append(p.capitalize() if p.isupper() or p.islower() else p)
    return " ".join(out)


_TAG_RE = re.compile(r"<div\b|</div>")


def _find_matching_div_end(html, div_start):
    """Dado o índice onde começa uma tag <div ...>, encontra o índice logo
    após o </div> que fecha ela (contando divs aninhadas). Retorna -1 se
    não achar (HTML malformado)."""
    primeiro_fechamento_tag = html.find(">", div_start)
    if primeiro_fechamento_tag == -1:
        return -1
    pos = primeiro_fechamento_tag + 1
    profundidade = 1
    for m in _TAG_RE.finditer(html, pos):
        if m.group(0) == "<div":
            profundidade += 1
        else:
            profundidade -= 1
            if profundidade == 0:
                return m.end()
    return -1


def montar_card_novo_html(codigo, nome_produto, peso_cx, cod_cancao, saldo, secao):
    """Monta o HTML de um card novo, no mesmo formato dos existentes.
    O nome exibido prioriza NOME_CANONICO (fixado manualmente); só usa a
    descrição bruta do ERP, formatada, quando o código não está mapeado."""
    if codigo in NOME_CANONICO:
        nome_fmt = NOME_CANONICO[codigo]
    else:
        nome_fmt = slug_nome(nome_produto)
    nome_esc = nome_fmt.replace('"', "&quot;")
    status = "available" if saldo and saldo > 0 else "no-stock"
    stock_html = (
        f'<div class="card-stock"><div class="stock-kg">{fmt_kg(saldo)}</div></div>'
        if status == "available" else
        '<div class="card-stock"><div class="no-stock-tag">SEM ESTOQUE</div></div>'
    )
    peso_txt = f"Cx {peso_cx} kg" if peso_cx else ""
    cancao_txt = cod_cancao if cod_cancao else "—"
    return (
        f'<div class="card {status}" data-name="{nome_esc}" data-section="{secao}" data-status="{status}">'
        f'<div class="card-info"><div class="card-name">{nome_esc}</div>'
        f'<div class="card-code">{codigo} · {peso_txt}</div>'
        f'<div class="card-cancao">Cód. Canção: {cancao_txt}</div></div>'
        f'{stock_html}</div>\n      '
    )


def inserir_cards_novos(html, data, linhas_brutas):
    """Para códigos que estão na planilha mas não têm card nenhum no HTML,
    cria o card automaticamente. Prefixos fixos (COMPRO, PEFLCANMI) vão
    direto para Alimentos Preparados; os demais são classificados pelo
    texto da descrição do produto (RESFRIADA/CONGELADA/IQF + tipo de
    corte). Só fica sem criar quando a descrição não dá pra classificar
    com segurança - nesse caso o script avisa no log."""
    codigos_existentes = set(re.findall(r'card-code">([^<· ]+)', html))

    codigos_novos = [c for c in data.keys() if c not in codigos_existentes]
    if not codigos_novos:
        return html, []

    avisos = []
    inseridos = []

    for codigo in codigos_novos:
        linha = linhas_brutas.get(codigo, [])
        nome_produto = linha[3] if len(linha) > 3 else codigo

        destino = classificar_produto_novo(nome_produto)

        if destino is None:
            avisos.append(
                f"código novo '{codigo}' ({nome_produto}) não tem card e a "
                f"descrição não deu pra classificar com segurança - card "
                f"NÃO foi criado automaticamente, precisa adicionar manualmente uma vez."
            )
            continue

        secao, grupo_label = destino
        m_peso = re.search(r"CX\s+([\d,.]+)\s*KG", nome_produto, re.I)
        peso_cx = m_peso.group(1).replace(".", ",") if m_peso else ""
        saldo = data.get(codigo, 0)

        cod_cancao = COD_CANCAO_FIXO.get(codigo, "")
        card_html = montar_card_novo_html(codigo, nome_produto, peso_cx, cod_cancao, saldo, secao)

        # Acha a seção certa, depois o group-label certo DENTRO dela, depois
        # a div "cards" logo em seguida - usando contagem de divs balanceada
        # (regex simples falha com divs aninhadas do mesmo nome).
        inserido_ok = False
        secao_start = html.find(f'data-section="{secao}"')
        if secao_start != -1:
            # recua até o "<div class="section"" que contém esse data-section
            secao_div_start = html.rfind('<div class="section"', 0, secao_start)
            secao_end = _find_matching_div_end(html, secao_div_start)
            if secao_end != -1:
                bloco_secao = html[secao_div_start:secao_end]
                grupo_idx = bloco_secao.find(f'<div class="group-label">{grupo_label}</div>')
                if grupo_idx != -1:
                    cards_div_start = bloco_secao.find('<div class="cards">', grupo_idx)
                    if cards_div_start != -1:
                        cards_div_end = _find_matching_div_end(bloco_secao, cards_div_start)
                        if cards_div_end != -1:
                            # cards_div_end aponta logo apos o "</div>" de fechamento;
                            # insere o novo card imediatamente antes desse "</div>".
                            ponto_insercao = cards_div_end - len("</div>")
                            novo_bloco_secao = (
                                bloco_secao[:ponto_insercao]
                                + card_html
                                + bloco_secao[ponto_insercao:]
                            )
                            html = html[:secao_div_start] + novo_bloco_secao + html[secao_end:]
                            inserido_ok = True

        if inserido_ok:
            inseridos.append(codigo)
        else:
            avisos.append(
                f"código novo '{codigo}' seria criado em '{secao}/{grupo_label}' "
                f"mas não encontrei esse grupo no HTML - card NÃO foi criado."
            )

    for a in avisos:
        print("AVISO -", a)
    if inseridos:
        print(f"OK - {len(inseridos)} card(s) novo(s) criado(s) automaticamente: {', '.join(inseridos)}")

    return html, inseridos


def atualizar_html(html, data, validades=None, linhas_brutas=None):
    validades = validades or {}
    linhas_brutas = linhas_brutas or {}

    # ── Cria cards novos ANTES de atualizar os existentes, para que o saldo
    # do card recém-criado já seja processado no mesmo ciclo. ──
    html, _ = inserir_cards_novos(html, data, linhas_brutas)

    card_starts = [m.start() for m in re.finditer(r'<div class="card[^"]*" data-name=', html)]
    card_starts.append(len(html))

    segments = []
    for i in range(len(card_starts) - 1):
        start, end = card_starts[i], card_starts[i + 1]
        # O "block" entre um card e o próximo pode conter, depois do
        # fechamento do card, HTML estrutural (fechamento de grupo/seção,
        # cabeçalho da próxima seção etc). Isolamos só o HTML do card em
        # si (via contagem balanceada de divs) do "rastro" que vem depois,
        # pra nunca apagar estrutura junto quando um card é removido.
        card_own_end = _find_matching_div_end(html, start)
        if card_own_end == -1 or card_own_end > end:
            card_own_end = end
        block = html[start:card_own_end]
        rastro = html[card_own_end:end]

        m = re.search(r'card-code">([^<]+)</div>', block)
        if not m:
            segments.append(block)
            segments.append(rastro)
            continue
        code = m.group(1).split("·")[0].strip()

        # ── Autocorreção: nome e Cód. Canção sempre seguem as listas fixas
        # (NOME_CANONICO / COD_CANCAO_FIXO), mesmo em cards que já existiam
        # antes dessas listas serem criadas - corrige uma vez e pra sempre. ──
        if code in NOME_CANONICO:
            nome_certo = NOME_CANONICO[code]
            nome_certo_esc = nome_certo.replace('"', "&quot;")
            block = re.sub(
                r'(<div class="card[^"]*" data-name=")[^"]*(")',
                lambda m2: m2.group(1) + nome_certo_esc + m2.group(2),
                block, count=1,
            )
            block = re.sub(
                r'(<div class="card-name">)[^<]*(</div>)',
                lambda m2: m2.group(1) + nome_certo_esc + m2.group(2),
                block, count=1,
            )
        if code in COD_CANCAO_FIXO:
            cancao_certo = COD_CANCAO_FIXO[code]
            block = re.sub(
                r'(<div class="card-cancao">Cód\. Canção: )[^<]*(</div>)',
                lambda m2: m2.group(1) + cancao_certo + m2.group(2),
                block, count=1,
            )

        # ── Gerencia o botão "ver validades" independente do estoque ──
        tem_validade = code in validades and validades[code].get("entries")
        tem_botao = bool(re.search(r'<button class="val-btn[^"]*"[^>]*data-code="' + re.escape(code) + r'"', block))
        if tem_validade and not tem_botao:
            hv = validades[code].get("hv", False)
            if hv:
                btn = f' <button class="val-btn val-btn-alert" onclick="toggleVal(this)" data-code="{code}">⚠️ ver validades</button>'
            else:
                btn = f' <button class="val-btn" onclick="toggleVal(this)" data-code="{code}">📅 ver validades</button>'
            novo_block = re.sub(
                r'(<div class="card-cancao">Cód\. Canção: [^<]*</div>)</div>',
                r'\1' + btn + '</div>', block, count=1,
            )
            if novo_block != block:
                block = novo_block
        elif not tem_validade and tem_botao:
            block = re.sub(r'\s*<button class="val-btn[^"]*"[^>]*>[^<]*</button>', '', block, count=1)
        elif tem_validade and tem_botao:
            # atualiza o estilo do botão (normal <-> alerta) se o hv mudou
            hv = validades[code].get("hv", False)
            nova_classe = "val-btn val-btn-alert" if hv else "val-btn"
            novo_texto = "⚠️ ver validades" if hv else "📅 ver validades"
            block = re.sub(
                r'<button class="val-btn[^"]*"(\s+onclick="toggleVal\(this\)" data-code="' + re.escape(code) + r'">)[^<]*</button>',
                f'<button class="{nova_classe}"' + r'\1' + novo_texto + '</button>',
                block, count=1,
            )

        if code not in data:
            # Regra: se o código não aparece mais na planilha de estoque,
            # o produto foi descontinuado/removido do ERP - o card some do
            # painel (não fica com o saldo antigo parado). Se o código
            # reaparecer numa planilha futura, o card é recriado automaticamente.
            # O "rastro" (fechamentos de grupo/seção, próximo cabeçalho)
            # é sempre preservado, mesmo quando o card é removido.
            segments.append(rastro)
            continue
        val = data[code]
        was_nostock = 'data-status="no-stock"' in block
        classe_atual_m = re.search(r'<div class="card ([^"]*)" data-name=', block)
        tinha_has_obs = classe_atual_m and "has-obs" in classe_atual_m.group(1)
        if val and val > 0:
            block = re.sub(
                r'(<div class="card )[^"]*(" data-name="[^"]*" data-section="[^"]*" data-status=")[^"]*(")',
                r"\1available\g<2>available\3", block, count=1,
            )
            block = re.sub(
                r'<div class="card-stock">.*?</div>\s*</div>',
                f'<div class="card-stock"><div class="stock-kg">{fmt_kg(val)}</div></div>',
                block, count=1, flags=re.S,
            )
            if was_nostock:
                block = re.sub(r'\s*<div class="obs-wrap">.*?</div>\s*', "\n", block, flags=re.S)
                block = block.replace(" has-obs", "").replace("has-obs ", "").replace("has-obs", "")
        else:
            classe_nova = "no-stock has-obs" if tinha_has_obs else "no-stock"
            block = re.sub(
                r'(<div class="card )[^"]*(" data-name="[^"]*" data-section="[^"]*" data-status=")[^"]*(")',
                rf"\1{classe_nova}\g<2>no-stock\3", block, count=1,
            )
            block = re.sub(
                r'<div class="card-stock">.*?</div>\s*</div>',
                '<div class="card-stock"><div class="no-stock-tag">SEM ESTOQUE</div></div>',
                block, count=1, flags=re.S,
            )
        segments.append(block)
        segments.append(rastro)

    new_html = html[: card_starts[0]] + "".join(segments)

    # Atualiza os dados de validade usados pelos selos automáticos (vence em Xd / vencido)
    new_html = re.sub(
        r"var VALIDADES = (\{.*?\})\s*;",
        "var VALIDADES = " + json.dumps(validades, ensure_ascii=False) + ";",
        new_html, count=1, flags=re.S,
    )

    # Atualiza o histórico de tendência (mantém últimos 4 ciclos)
    m = re.search(r"var HISTORICO = (\{.*?\})\s*;", new_html, re.S)
    if m:
        hist = json.loads(m.group(1))
        for code, val in data.items():
            arr = hist.get(code, [])
            arr.append(float(val))
            hist[code] = arr[-4:]
        new_hist_str = json.dumps(hist, ensure_ascii=False)
        new_html = new_html[: m.start(1)] + new_hist_str + new_html[m.end(1):]

    # Atualiza o horário exibido como "Atualizado hoje às HHhMM" (horário de Brasília, UTC-3)
    agora_brt = datetime.now(timezone.utc) - timedelta(hours=3)
    novo_timestamp = agora_brt.strftime("%Y-%m-%dT%H:%M:%S")
    new_html = re.sub(
        r'var ULTIMA_ATUALIZACAO = "[^"]*";',
        f'var ULTIMA_ATUALIZACAO = "{novo_timestamp}";',
        new_html,
    )

    # Atualiza também o selo de data no cabeçalho (badge-date), que é o campo
    # que decide se o painel mostra "hoje"/"ontem"/"há X dias"
    nova_data_badge = agora_brt.strftime("%d/%m/%Y")
    new_html = re.sub(
        r'<span class="badge-date">[^<]*</span>',
        f'<span class="badge-date">{nova_data_badge}</span>',
        new_html,
    )
    new_html = re.sub(
        r'<title>Estoque Apucarana – [^<]*</title>',
        f'<title>Estoque Apucarana – {nova_data_badge}</title>',
        new_html,
    )

    return new_html


def fmt_preco(v):
    v = (v or "").strip()
    if not v:
        return None
    # valores não numéricos (ex: "—") são mantidos como estão
    limpo = v.replace(".", "").replace(",", ".")
    try:
        num = float(limpo)
    except ValueError:
        return v
    return f"{num:.2f}".replace(".", ",")


def carregar_precos(csv_text):
    """Lê a aba Precos (NOME, REFERENCIA, PRECO_1, PRECO_2, PRECO_3) e
    devolve { nome: [preco1, preco2, preco3] }, com None nas posições vazias
    (posição vazia = não mexe no valor que já está no site)."""
    reader = csv.reader(StringIO(csv_text))
    rows = list(reader)
    precos = {}
    for r in rows[2:]:
        if not r or not r[0]:
            continue
        nome = r[0].strip()
        valores = []
        for i in (2, 3, 4):  # colunas PRECO_1, PRECO_2, PRECO_3
            bruto = r[i].strip() if len(r) > i else ""
            valores.append(fmt_preco(bruto))
        precos[nome] = valores
    return precos


def atualizar_precos_html(html, precos):
    card_pattern = re.compile(r'(<div class="card[^"]*" data-sec="[^"]*" data-name="([^"]+)">)(.*?)(</div>)', re.S)

    trend_m = re.search(r"var PRICE_TREND = (\{.*?\})\s*;", html, re.S)
    try:
        trend = json.loads(trend_m.group(1)) if trend_m else {}
    except json.JSONDecodeError:
        trend = {}

    def processa_card(m):
        abertura, nome, inner, fechamento = m.groups()
        if nome not in precos:
            return m.group(0)

        spans = list(re.finditer(r'<span class="p([^"]*)">([^<]*)</span>', inner))
        if not spans:
            return m.group(0)

        novos_valores = precos[nome]
        preco_principal_antigo = spans[0].group(2)
        preco_principal_novo = novos_valores[0] if novos_valores[0] else preco_principal_antigo

        novo_inner = inner
        # aplica de trás pra frente pra não bagunçar os índices das outras substituições
        for i in reversed(range(len(spans))):
            if i >= len(novos_valores) or novos_valores[i] is None:
                continue
            span = spans[i]
            classes = span.group(1)
            novo_span = f'<span class="p{classes}">{novos_valores[i]}</span>'
            novo_inner = novo_inner[: span.start()] + novo_span + novo_inner[span.end():]

        # tendência: compara o preço principal (primeira coluna) antigo vs novo
        try:
            antigo_f = float(preco_principal_antigo.replace(".", "").replace(",", "."))
            novo_f = float(preco_principal_novo.replace(".", "").replace(",", "."))
            if nome not in trend:
                trend[nome] = "new"
            elif novo_f > antigo_f:
                trend[nome] = "up"
            elif novo_f < antigo_f:
                trend[nome] = "down"
            else:
                trend[nome] = "same"
        except ValueError:
            pass

        return abertura + novo_inner + fechamento

    new_html = card_pattern.sub(processa_card, html, count=0)

    new_trend_str = json.dumps(trend, ensure_ascii=False)
    new_html, n_subs = re.subn(
        r"var PRICE_TREND = \{.*?\}\s*;",
        "var PRICE_TREND = " + new_trend_str.replace("\\", "\\\\") + ";",
        new_html, count=1, flags=re.S,
    )

    return new_html


def carregar_pedidos(csv_text):
    """Lê a planilha bruta de pedidos (mesmas colunas do relatório do ERP:
    Lj Embarque, Cod. Vendedor, Vendedor, Pedido, Data Pedido, Liberado,
    Cliente/Loja, Razao Social, Bairro, CEP, Cidade, UF, Obs. Pedido,
    Dt Carregamento, Dt Entrega, Cod. Produto, Produto, Qtd Liberada,
    Qtd Liberada Total, Peso Cx, Vl Total Pedido, Viagem, Carga,
    Gestão de Pedidos, Saldo sem Gestão) e agrupa em
    { vendedor: { cod_vendedor, pedidos: [ {..., produtos: [...]} ] } }
    (sem campo de senha - isso é preservado à parte, do HTML já existente)."""
    reader = csv.reader(StringIO(csv_text))
    rows = list(reader)

    vendedores = {}
    pedidos_por_vendedor = {}

    for r in rows[2:]:
        if not r or not r[0] or len(r) < 25:
            continue
        (lj_embarque, cod_vendedor, vendedor, pedido, data_pedido, liberado,
         cliente_loja, razao_social, bairro, cep, cidade, uf, obs_pedido,
         dt_carregamento, dt_entrega, cod_produto, nome_produto, qtd_liberada,
         qtd_liberada_total, peso_cx, vl_total_pedido, viagem, carga,
         gestao_pedidos, saldo_sem_gestao) = r[:25]

        if not vendedor:
            continue
        vendedores.setdefault(vendedor, cod_vendedor)
        chave_pedido = (vendedor, pedido)
        lista = pedidos_por_vendedor.setdefault(vendedor, {})
        if chave_pedido not in lista:
            lista[chave_pedido] = {
                "pedido": pedido, "lj_embarque": lj_embarque, "data_pedido": data_pedido,
                "liberado": liberado, "cliente_loja": cliente_loja, "razao_social": razao_social,
                "bairro": bairro, "cep": cep, "cidade": cidade, "uf": uf,
                "obs_pedido": obs_pedido, "dt_carregamento": dt_carregamento,
                "dt_entrega": dt_entrega, "vl_total_pedido": vl_total_pedido,
                "viagem": viagem, "carga": carga, "gestao_pedidos": gestao_pedidos,
                "produtos": [],
            }
        if cod_produto:
            try:
                ql = float(qtd_liberada.replace(".", "").replace(",", ".")) if qtd_liberada else 0
                qlt = float(qtd_liberada_total.replace(".", "").replace(",", ".")) if qtd_liberada_total else 0
            except ValueError:
                ql = qlt = 0
            saldo_norm = saldo_sem_gestao
            if saldo_sem_gestao:
                try:
                    saldo_val = float(str(saldo_sem_gestao).replace(".", "").replace(",", "."))
                    saldo_norm = str(int(saldo_val)) if saldo_val == int(saldo_val) else str(saldo_val).replace(".", ",")
                except ValueError:
                    saldo_norm = saldo_sem_gestao

            produtos = lista[chave_pedido]["produtos"]
            existente = next((p for p in produtos if p["cod"] == cod_produto), None)
            if existente is None:
                produtos.append({
                    "cod": cod_produto, "nome": nome_produto,
                    "qtd_liberada": ql, "qtd_liberada_total": qlt,
                    "peso_cx": peso_cx, "saldo_sem_gestao": saldo_norm,
                    "_ql_vistos": {ql},
                })
            else:
                # Mesmo produto já apareceu neste pedido:
                # - se essa qtd_liberada já foi vista antes, é linha duplicada
                #   (erro de exportação do ERP) -> ignora, não soma de novo
                # - se é uma qtd_liberada nova, é uma liberação parcial
                #   adicional -> soma na quantidade liberada exibida
                #   (qtd_liberada_total NÃO muda, é o mesmo total do pedido)
                if ql not in existente["_ql_vistos"]:
                    existente["_ql_vistos"].add(ql)
                    existente["qtd_liberada"] += ql

    resultado = {}
    for vendedor, cod_vendedor in vendedores.items():
        pedidos_lista = list(pedidos_por_vendedor[vendedor].values())
        for p in pedidos_lista:
            for prod in p["produtos"]:
                prod.pop("_ql_vistos", None)
        resultado[vendedor] = {
            "cod_vendedor": cod_vendedor,
            "pedidos": pedidos_lista,
        }
    return resultado


def atualizar_pedidos_html(html, novos_dados):
    """Reconstrói o DATA do painel de pedidos, preservando as senhas já
    cadastradas no arquivo (a planilha não tem senha, isso é fixo no HTML)."""
    m = re.search(r"const DATA = (\{.*?\});", html, re.S)
    if not m:
        return html
    data_atual = json.loads(m.group(1))

    novo_data = {}
    for vendedor, info in novos_dados.items():
        senha_existente = data_atual.get(vendedor, {}).get("senha", "0000")
        novo_data[vendedor] = {
            "cod_vendedor": info["cod_vendedor"],
            "senha": senha_existente,
            "pedidos": info["pedidos"],
        }
    # mantém vendedores antigos que não vieram na planilha nova (sem pedidos novos)
    for vendedor, info in data_atual.items():
        if vendedor not in novo_data:
            novo_data[vendedor] = {**info, "pedidos": []}

    novo_data_str = json.dumps(novo_data, ensure_ascii=False)
    new_html, n = re.subn(
        r"const DATA = \{.*?\};",
        "const DATA = " + novo_data_str.replace("\\", "\\\\") + ";",
        html, count=1, flags=re.S,
    )
    return new_html


def main():
    csv_text = baixar_csv(CSV_URL)
    data, linhas_brutas = carregar_saldos(csv_text)

    validades = {}
    if VALIDADES_CSV_URL:
        validades_csv_text = baixar_csv(VALIDADES_CSV_URL)
        validades = carregar_validades(validades_csv_text)

    with open(REPO_INDEX_PATH, encoding="utf-8") as f:
        html = f.read()

    new_html = atualizar_html(html, data, validades, linhas_brutas)

    with open(REPO_INDEX_PATH, "w", encoding="utf-8") as f:
        f.write(new_html)

    print(f"OK - {len(data)} códigos de estoque, {len(validades)} códigos com validade processados.")

    if PRECOS_CSV_URL:
        try:
            precos_csv_text = baixar_csv(PRECOS_CSV_URL)
            precos = carregar_precos(precos_csv_text)

            with open(REPO_PRECOS_PATH, encoding="utf-8") as f:
                precos_html = f.read()

            novo_precos_html = atualizar_precos_html(precos_html, precos)

            with open(REPO_PRECOS_PATH, "w", encoding="utf-8") as f:
                f.write(novo_precos_html)

            print(f"OK - {len(precos)} produtos de preço processados.")
        except Exception as e:
            # Uma falha aqui não deve apagar a atualização de estoque que já
            # rodou com sucesso acima - só avisa e segue (o commit do estoque
            # ainda acontece normalmente).
            print(f"AVISO - falha ao atualizar preços, estoque foi salvo mesmo assim: {e}")

    if PEDIDOS_CSV_URL:
        try:
            pedidos_csv_text = baixar_csv(PEDIDOS_CSV_URL)
            pedidos = carregar_pedidos(pedidos_csv_text)

            with open(REPO_PEDIDOS_PATH, encoding="utf-8") as f:
                pedidos_html = f.read()

            novo_pedidos_html = atualizar_pedidos_html(pedidos_html, pedidos)

            with open(REPO_PEDIDOS_PATH, "w", encoding="utf-8") as f:
                f.write(novo_pedidos_html)

            print(f"OK - {len(pedidos)} vendedores / pedidos processados.")
        except Exception as e:
            print(f"AVISO - falha ao atualizar pedidos: {e}")


if __name__ == "__main__":
    main()
