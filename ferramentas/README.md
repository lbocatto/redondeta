# Ferramentas

## Atualizar o site a partir da Base Geral

`atualizar_site.py` lê a planilha "REDONDETA — Base Geral do Acervo Vivo 2021–2026" e grava os dados no `index.html`.

1. No Google Sheets, abra a Base Geral e use **Arquivo → Fazer download → Microsoft Excel (.xlsx)**.
2. Na pasta do repositório, rode:

   ```
   pip install openpyxl
   python3 ferramentas/atualizar_site.py "caminho/da/planilha.xlsx" index.html
   ```

3. Confira o resumo e os avisos que o script mostra, abra o `index.html` no navegador e suba o arquivo como de costume.

### O que vem da planilha e o que fica no site

| Vem da planilha | Fica como está no `index.html` |
|---|---|
| Todas as abas de entidades, relações, índice, fichas, imagens, introduções e lacunas | Textos da Rosa dos Ventos (`meta`) e dos núcleos centrais |
| Quests (aba `22_Quests_Tarefas`) | Quais termos aparecem em Mundo e Cosmologia e entre os fundamentos de Aventuraulas (seleção editorial) |
| Campanhas e Arcos (aba `23_Campanhas_Arcos`) | |
| Aventuraulas dissecadas em 8 etapas (aba `24_Aventuraulas`) | |

Regras aplicadas:

- Registros com Status documental "Superado" não são publicados.
- Aventuraulas (S) mostram primeiro as aventuraulas da aba 24 e depois os termos de fundamentos do método.
- Eventos e Cronologia (SO) saem na ordem da coluna "Ordem no mundo" (acontecimentos no mundo de Redondeta, não a data das aulas).
- Legados e Transformações (OSO) saem na ordem do ID do herói.
- Continentes são os locais do tipo "Continente/região"; os demais tipos vão para Lugares.
- Termos fora de Mundo e Cosmologia e de Aventuraulas não aparecem na Rosa, mas a busca abre a ficha completa.
- Os números citados nos subtítulos (eventos, personagens, itens, fontes) são recalculados.
- Nomes de alunos listados em `REDACOES` são trocados pelas iniciais. Isso é só uma proteção: o certo é corrigir na planilha.

### Avisos

O script avisa quando encontra ID duplicado, o termo "canônico" em textos, ou ID do índice sem registro completo no site. Nenhum aviso impede a gravação; corrija na planilha e rode de novo.
