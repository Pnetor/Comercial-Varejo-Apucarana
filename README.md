[README.md](https://github.com/user-attachments/files/31646274/README.md)
[README.md](https://github.com/user-attachments/files/31436092/README.md)
# Comercial-Varejo-Apucarana

## Classificação automática de itens novos

Quando surge um código novo na planilha de estoque, o script tenta descobrir
sozinho em qual seção do painel ele deve entrar (Resfriados/Congelados/IQF/
Alimentos Preparados), nesta ordem:

1. **Palavras-chave** na descrição do produto (ex: "CONGELADO", "IQF",
   "EMPANADO"). Cobre a maioria dos casos, sem custo e sem depender de nada
   externo.
2. **IA (Claude)**, só quando as palavras-chave não bastam. Requer a secret
   `ANTHROPIC_API_KEY` configurada no repositório (Settings → Secrets and
   variables → Actions). Sem essa secret, essa etapa é simplesmente pulada.
3. **"Não Classificados"**: se nada acima resolver, o card ainda assim é
   criado, numa seção própria do painel - nunca é descartado silenciosamente.
   Basta mover manualmente pra seção certa quando der.

## Como o painel é atualizado

O script `scripts/atualizar_estoque.py` roda automaticamente todo hora via
GitHub Actions (`.github/workflows/atualizar-estoque.yml`), lê os dados
publicados no Google Sheets (CSV, endereço na secret `GSHEET_CSV_URL`) e
regenera o `index.html` (estoque), `tabela-precos.html` e
`pedidos-em-aberto.html` direto no repositório. Não existe backend/servidor
por trás do site - é tudo HTML estático gerado por esse script e publicado
pelo GitHub Pages.

## Trava de validação (evita dado ruim ir pro ar)

Antes de sobrescrever cada uma das três páginas, o script confere se o
resultado faz sentido perto do que já estava publicado:

- **Estoque**: se o número de cards cair demais em relação ao anterior
  (mais de 50% de queda, e só quando já existiam pelo menos 5 cards), a
  atualização daquela página é **bloqueada** - fica valendo a versão
  anterior, pra não sumir com o painel inteiro por causa de um erro na
  planilha.
- **Preços**: mesma lógica, comparando os preços atuais com os anteriores;
  um salto absurdo (mais de 50%) trava a atualização da tabela de preços.
- **Pedidos em aberto**: mesma ideia, comparando a quantidade de
  vendedores/pedidos.

Quando uma dessas travas aciona, o job do GitHub Actions termina com erro
(fica vermelho), mas as páginas que passaram na validação são commitadas
normalmente - só a página com problema fica intacta, esperando a planilha
ser corrigida.

## Itens novos: destaque, edição e pendências

Quando um código novo aparece (estoque, preço ou vendedor de pedidos), o
card correspondente nasce com um selo **"🆕 NOVO"** e botões de ação, pra
facilitar o ajuste fino que a automação sozinha não sabe fazer:

- **Estoque**: renomear o nome de exibição do item e definir o Cód. Canção
  dele.
- **Tabela de preços**: renomear o nome de exibição do item.
- **Pedidos em aberto**: definir a senha do vendedor.

Qualquer uma dessas ações pede a **senha de administrador** antes de deixar
editar (senha padrão `0000`, definida em `SENHA_ADMIN` no `<script>` de
cada página - troque lá se quiser outra). Ao usar uma ação, só o botão
daquela ação some do card; o selo "NOVO" inteiro só desaparece quando não
sobra mais nenhuma ação pendente (ou quando "Dispensar destaque" é clicado
direto).

Essas edições ficam guardadas só no navegador (localStorage) até serem
aplicadas de verdade - o site é estático e não tem como salvar sozinho.
Por isso cada página tem um painel de pendências (chip flutuante) com um
botão de **copiar** que gera um bloco de texto pronto (`NOME_CANONICO` /
`COD_CANCAO_FIXO` etc.) pra colar direto no topo do `atualizar_estoque.py`
e commitar - assim a correção vale pra sempre, pra qualquer pessoa que
abrir o painel, e não só no navegador de quem editou.

## Botão "🚚 programação" (programação de cargas)

Cada card do estoque pode mostrar um botão **"🚚 programação"** com as
cargas em trânsito/programadas daquele produto especificamente (data de
carregamento, status, NF, placa, pedido, Kg programado/faturado). Os dados
vêm de uma segunda planilha (a de programação de cargas, publicada como CSV
na secret opcional `GSHEET_PROGRAMACAO_CSV_URL`), cruzada com o estoque pela
coluna **Código Protheus** (o mesmo código que já aparece no card, ex:
`FRCSCANMI000011`). O botão só aparece nos produtos que realmente têm
alguma linha de programação - sem essa secret configurada, o painel
funciona normalmente, só sem esse botão.

## E-mails Padrão

A aba **"📧 Emails"** (`emails-padrao.html`) reúne modelos prontos de
e-mail que os vendedores usam no dia a dia (prioridade de cadastro,
bonificação, cadastro de forma de pagamento, baixa de título, troca de
produto/reclamação, abertura de ocorrência, programação de cargas). O
vendedor escolhe o modelo, preenche um formulário curto na própria página e
o e-mail já abre pronto no cliente de e-mail dele, com destinatários,
assunto, corpo e cópia (geralmente o supervisor) preenchidos - sem precisar
digitar nada manualmente. Pra adicionar um novo modelo, basta incluir um
objeto na lista `TEMPLATES` dentro do `<script>` de `emails-padrao.html`.

## PWA (instalável)

`index.html`, `tabela-precos.html` e `emails-padrao.html` registram um
service worker (`sw.js`) e referenciam o `manifest.json`, permitindo
instalar o painel como app (ícone na tela, abre sem barra de navegador).
As páginas HTML são sempre buscadas da rede primeiro (nunca mostra
estoque/preço desatualizado só porque tem internet); o cache só entra em
ação se a conexão cair no meio de uma consulta - só ícones e manifest
ficam em cache-first. `pedidos-em-aberto.html` não registra service
worker (mas continua abrindo normalmente pelas abas).

**Atenção ao nome do arquivo**: o service worker precisa estar no
repositório como **`sw.js`** exatamente. Se ele for enviado com outro
nome (ex: `sw.txt`), o navegador toma 404 ao registrar e o painel deixa
de ser instalável - o site continua funcionando, mas sem o comportamento
de app.

Todas as páginas usam `viewport-fit=cover` no `<meta name="viewport">` e
somam `env(safe-area-inset-top)` no `padding` do cabeçalho. Isso é
obrigatório: sem o `viewport-fit=cover`, o `env()` vale sempre 0 e o
cabeçalho (com as abas de navegação) fica embaixo do relógio/bateria do
celular quando o painel abre como app instalado.
