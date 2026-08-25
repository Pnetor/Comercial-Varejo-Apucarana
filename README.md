[README.md](https://github.com/user-attachments/files/31432123/README.md)
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

## PWA (instalável)

`index.html` e `tabela-precos.html` registram um service worker
(`sw.js`) e referenciam um `manifest.json`, permitindo instalar o painel
como app (ícone na tela, abre sem barra de navegador). As páginas HTML
são sempre buscadas da rede primeiro (nunca mostra estoque/preço
desatualizado só porque tem internet); o cache só entra em ação se a
conexão cair no meio de uma consulta. `pedidos-em-aberto.html` não
registra service worker.
