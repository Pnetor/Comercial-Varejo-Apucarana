[README.md](https://github.com/user-attachments/files/31418863/README.md)
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
