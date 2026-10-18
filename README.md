# Educação Financeira Familiar — Windows 11 x64

Programa desktop em português do Brasil baseado na planilha **Organização Financeira Para Família**: mesmo método (divisão da receita em 55 / 20 / 10 / 10 / 5), mesmas 74 categorias e mesmas regras de cálculo, com gráficos, alertas e conteúdo educativo.

## Telas
| Tela | O que faz |
|---|---|
| **Painel** | Mês selecionado: receitas, despesas, investimentos e saldo; barras dos últimos 6 meses; pizza por bloco; rosca por categoria; pizza da origem das receitas; orçado x realizado; alertas. |
| **Painel do ano** | Equivale à aba *Painel* da planilha: receita, despesas, investimentos, saldo, taxa de poupança, gasto médio; barras mês a mês; pizza dos gastos por grupo; **REAL x IDEAL**; orçado x realizado no ano. |
| **Planilha** | Equivale à aba *Orçamento Mensal*: as 9 seções, 12 meses + total do ano, subtotais, sobra automática, taxa de poupança e controle orçado x realizado. **Duplo clique** em uma célula de categoria para digitar o valor do mês. |
| **Lançamentos** | Registro do dia a dia: tipo → bloco → categoria da planilha. **Editar** (botão ou duplo clique) reescreve um lançamento errado; **Excluir** aceita vários de uma vez; **Limpar lançamentos do mês**. |
| **Orçamento** | Percentuais editáveis (soma 100%) com pizza em tempo real. |
| **Metas** | Progresso, valor guardado por meta e aporte mensal sugerido. **Editar meta**, atualizar valor guardado, excluir e limpar todas. |
| **Investimentos** | Aportes por classe e evolução. |
| **Relatórios** | Resultado do ano, exportação Excel (com gráficos) e CSV, backup e restauração, **limpeza seletiva** e **Restaurar do zero**. |
| **Aprender** | Passo a passo, explicação do método, simuladores (divisão da renda, reserva de emergência, juros compostos) e lições. |

## Corrigir, limpar e recomeçar
- **Corrigir:** Lançamentos e Metas têm *Editar* (selecione a linha, corrija e salve). Na Planilha, duplo clique na célula e digite o valor certo; em branco zera.
- **Limpar:** mês, ano ou só o que veio do Excel (Planilha e Relatórios); todas as metas; percentuais de volta ao 55/20/10/10/5 (Orçamento e Relatórios). Toda limpeza pede confirmação.
- **Restaurar do zero:** botão vermelho no menu lateral e em Relatórios. Pede duas confirmações (inclusive digitar `APAGAR TUDO`) e cria antes uma cópia de segurança `backup_automatico_antes_de_restaurar_*.db` na pasta de dados. Para desfazer, use *Relatórios → Restaurar backup* e escolha esse arquivo.
- Limpezas do ano e dos valores importados também criam cópia automática (ficam as 10 mais recentes).

## Anos ilimitados
Não há limite de anos. No topo da janela, **◀ ▶** mudam o mês e **« »** mudam o ano (12 meses); você também pode digitar qualquer mês no formato `AAAA-MM` (ex.: `2045-07`). A sobra do mês anterior é calculada desde o primeiro mês com dados, por quantos anos forem necessários.

Dados locais em `Documentos\EducacaoFinanceiraFamiliar\financas.db`.

## Regras iguais às da planilha
- A divisão 55/20/10/10/5 usa a **receita do mês**; a sobra do mês anterior entra no total disponível, mas não aumenta o orçamento.
- O saldo de um mês só vira sobra do mês seguinte se for **positivo**. Em janeiro a sobra é informada por você (duplo clique na linha *Sobra do mês anterior*).
- Despesas = essenciais + metas + fazer o que quiser + educação; saídas = despesas + investimentos.
- Importação da aba **Orçamento Mensal** (`.xlsx`/`.xlsm`): só valores digitados; totais e fórmulas são ignorados.

## Gerar o instalador sem instalar ferramentas no Windows
1. Crie um repositório no GitHub e envie todos os arquivos e pastas, inclusive `.github/workflows/build.yml` e `tests`.
2. Abra **Actions** e execute **Compilar instalador Windows x64**, ou faça push para `main`/`master`.
3. Baixe o artefato `EducacaoFinanceiraFamiliar-Setup-Windows-x64`.
4. Execute `EducacaoFinanceiraFamiliar-Setup-x64.exe`.

## Validação
- A planilha-modelo foi preenchida com dados de exemplo (incluindo um mês no vermelho), recalculada no LibreOffice e importada no programa: receita, sobra, os cinco blocos, investimentos, patrimônio, despesas, saídas, saldo e orçado de cada um dos 12 meses coincidem com o Excel, sem diferença.
- Os testes automatizados cobrem valores brasileiros, estrutura de categorias, cálculo do ano, regra da sobra, edição de células, correção de lançamentos e metas, limpezas, reset com cópia de segurança, simuladores e funções dos gráficos.
- A janela completa ainda não foi testada interativamente em uma sessão Windows; confira o layout no primeiro uso.

## Limitações
- Valores importados recebem o dia 01 do mês (a planilha guarda totais mensais por categoria).
- Se você editar uma célula na tela **Planilha** e depois reimportar a planilha do mesmo ano, os ajustes feitos na tela para aquele ano são substituídos.
- Lançamentos de versões anteriores com categorias que não existem na planilha aparecem na linha *Outros lançamentos (sem categoria da planilha)*.
- Editar um lançamento que veio da planilha o transforma em lançamento manual (ele deixa de ser substituído numa nova importação).
- A restauração de backup substitui o banco atual; faça backup antes.
- Programa educativo: não fornece recomendação individual de compra de ativos nem promete rentabilidade. Os simuladores são ilustrativos.
