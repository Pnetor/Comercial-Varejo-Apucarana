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

# ── Onde criar um card novo quando o código não existe ainda no HTML ──
# Mapeia prefixo do código -> (data-section, texto do group-label onde entra)
# Só prefixos sem ambiguidade entram aqui (ver NOVOS_CARDS_AMBIGUOS abaixo).
NOVO_CARD_DESTINO = {
    "COMPRO": ("alimentos", "ALIMENTOS"),
}
# Prefixos que aparecem em mais de uma seção/grupo no site (ex: FRASCANMI
# tanto em Resfriados quanto em Congelados) - não criamos card sozinho
# aqui, só avisamos no log para decidir manualmente uma vez.
PREFIXOS_AMBIGUOS = (
    "FRASCANMI", "FRCSCANMI", "FRPTCANMI", "FRINCANMI", "FRMDCANMI",
    "FRMSCANMI", "FRPEMIFME", "FRSACANMI", "FRMRCANMI", "PEFLCANMI",
)


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


def montar_card_novo_html(codigo, nome_produto, peso_cx, cod_cancao, saldo):
    """Monta o HTML de um card novo, no mesmo formato dos existentes."""
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
        f'<div class="card {status}" data-name="{nome_esc}" data-section="alimentos" data-status="{status}">'
        f'<div class="card-info"><div class="card-name">{nome_esc}</div>'
        f'<div class="card-code">{codigo} · {peso_txt}</div>'
        f'<div class="card-cancao">Cód. Canção: {cancao_txt}</div></div>'
        f'{stock_html}</div>\n      '
    )


def inserir_cards_novos(html, data, linhas_brutas):
    """Para códigos que estão na planilha mas não têm card nenhum no HTML,
    cria o card automaticamente (só para prefixos sem ambiguidade, hoje
    apenas COMPRO -> Alimentos Preparados / grupo ALIMENTOS). Prefixos
    ambíguos (FR...) só geram um aviso no log."""
    codigos_existentes = set(re.findall(r'card-code">([^<· ]+)', html))

    codigos_novos = [c for c in data.keys() if c not in codigos_existentes]
    if not codigos_novos:
        return html, []

    avisos = []
    inseridos = []

    for codigo in codigos_novos:
        destino = None
        for prefixo, dest in NOVO_CARD_DESTINO.items():
            if codigo.startswith(prefixo):
                destino = dest
                break

        if destino is None:
            if codigo.startswith(PREFIXOS_AMBIGUOS):
                avisos.append(
                    f"código novo '{codigo}' não tem card e o prefixo é ambíguo "
                    f"(pode ser Resfriados ou Congelados) - card NÃO foi criado "
                    f"automaticamente, precisa adicionar manualmente uma vez."
                )
            else:
                avisos.append(
                    f"código novo '{codigo}' não tem card e o prefixo não é "
                    f"reconhecido - card NÃO foi criado automaticamente."
                )
            continue

        secao, grupo_label = destino
        linha = linhas_brutas.get(codigo, [])
        nome_produto = linha[3] if len(linha) > 3 else codigo
        # tenta achar o "peso de caixa" a partir da descrição, ex: "... CX 4,2 KG"
        m_peso = re.search(r"CX\s+([\d,.]+)\s*KG", nome_produto, re.I)
        peso_cx = m_peso.group(1).replace(".", ",") if m_peso else ""
        saldo = data.get(codigo, 0)

        card_html = montar_card_novo_html(codigo, nome_produto, peso_cx, "", saldo)

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
        block = html[start:end]
        m = re.search(r'card-code">([^<]+)</div>', block)
        if not m:
            segments.append(block)
            continue
        code = m.group(1).split("·")[0].strip()

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
            segments.append(block)
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
