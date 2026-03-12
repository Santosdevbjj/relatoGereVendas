# 📊 Relatório Gerencial de Vendas com Excel e Power BI

![KlabinBiExcel003](https://github.com/user-attachments/assets/13395871-7fb0-4f11-96e2-b3c4cc3c7c1e)

**Bootcamp Klabin — Excel e Power BI Dashboards**


> 🔗 Repositório original do desafio: [power_bi_analyst — julianazanelatto](https://github.com/julianazanelatto/power_bi_analyst)

---

## 1. Problema de Negócio

Equipes de vendas e gestores precisam monitorar KPIs como Receita Total, Lucro e Desconto Médio — segmentados por país, segmento de mercado e período — para embasar decisões comerciais e identificar onde o negócio está performando abaixo do esperado.

O desafio é que nem toda organização tem o Power BI licenciado ou disponível para todos os usuários. A ausência da ferramenta não pode paralisar a análise gerencial.

**A pergunta central deste projeto é:**
> *Como entregar um relatório gerencial interativo, com navegação entre páginas e filtros dinâmicos, usando apenas o Excel — sem depender do Power BI?*

---

## 2. Contexto

Este projeto foi desenvolvido durante o **Bootcamp Klabin — Excel e Power BI Dashboards** como uma solução alternativa ao desafio original proposto em Power BI.

A base de dados utilizada é o **Financial Sample** da Microsoft — um conjunto público com dados de vendas internacionais por produto, segmento, país, mês e ano — amplamente usado para treinar análises de BI.

A solução replica a experiência de um dashboard de Power BI diretamente no Excel, usando macros VBA para simular a navegação entre páginas e a alternância de visuais, e fórmulas dinâmicas (`SUMIFS`, `AVERAGEIFS`) para calcular os KPIs com base nos filtros selecionados.

O projeto demonstra que **a ferramenta é secundária — o raciocínio analítico e a capacidade de entregar valor são o que importa.**

---

## 3. Premissas da Análise

- Os dados do **Financial Sample** são públicos e fornecidos pela Microsoft para fins educacionais — não há PII.
- As métricas de Vendas, Lucro e Desconto foram calculadas com base nas colunas originais do dataset, sem transformações que alterem a semântica dos dados.
- Os segmentadores (Segmento, País, Ano) funcionam via **Validação de Dados** do Excel — não há tabelas dinâmicas como motor principal, o que garante compatibilidade com versões mais antigas.
- As macros VBA foram escritas para funcionar no **Excel 2016 ou superior** com macros habilitadas.
- O relatório foi projetado como **prova de conceito**: demonstra a lógica gerencial antes da migração para o Power BI.

---

## 4. Estratégia da Solução

**Etapa 1 — Entendimento dos Dados**
Análise exploratória do Financial Sample: colunas disponíveis, tipos de dados, granularidade (linha = uma venda), países, segmentos e período coberto.

**Etapa 2 — Definição dos KPIs**
Seleção das três métricas centrais do relatório gerencial: Total de Vendas (`SUM`), Total de Lucro (`SUM`) e Média de Desconto (`AVERAGEIFS`). Cada KPI é calculado dinamicamente com base nos filtros ativos.

**Etapa 3 — Construção dos Visuais**
Três gráficos nativos do Excel foram implementados:
- **Vendas por Segmento** (barras horizontais)
- **Top Produtos por Lucro** (barras ordenadas)
- **Evolução das Vendas por Mês** (linha temporal)

**Etapa 4 — Interatividade com VBA**
Macros em `macros.bas` foram escritas para:
- Alternar entre **Página 1 (Visão Geral)** e **Página 2 (Detalhes)** via botões de navegação
- Trocar o visual exibido em um mesmo espaço gráfico, simulando o comportamento de bookmarks do Power BI

**Etapa 5 — Validação**
Conferência manual dos totais calculados pelas fórmulas contra os valores brutos do dataset, garantindo consistência dos KPIs para todas as combinações de filtro.

---

## 5. Insights da Análise

A construção do relatório revelou padrões relevantes sobre os dados do Financial Sample:

- O segmento **Government** concentra o maior volume de vendas, mas nem sempre o maior lucro — o desconto médio elevado nesse segmento comprime a margem.
- Os meses de **outubro a dezembro** apresentam pico de vendas consistente entre os anos disponíveis, o que sugere sazonalidade relevante para planejamento comercial.
- O produto **Paseo** lidera em lucro absoluto, enquanto produtos com desconto acima de 20% frequentemente operam no limite da lucratividade.
- A segmentação por **país** revela desempenhos assimétricos: mercados como Canadá e França apresentam margens mais saudáveis que outros com alto volume, mas alto desconto.

Esses padrões mostram que **volume de vendas e lucratividade contam histórias diferentes** — e um relatório gerencial precisa expor ambas simultaneamente.

---

## 6. Resultados

| Entregável | Descrição |
|---|---|
| `Dashboard.xlsm` | Arquivo Excel com dashboard interativo, KPIs, gráficos e navegação via VBA |
| `macros.bas` | Código VBA exportado para versionamento — navegação entre páginas e alternância de visuais |
| `instrucoes.txt` | Guia rápido de uso para qualquer usuário abrir e operar o dashboard |

O projeto entrega um relatório gerencial funcional e documentado que pode ser usado imediatamente por qualquer pessoa com Excel 2016+, sem necessidade de instalação adicional ou licença de BI.

---

## 7. Decisões Técnicas

**Por que Excel e não Power BI diretamente?**
O desafio original era em Power BI, mas optei por desenvolver primeiro no Excel para validar a lógica analítica e a estrutura do relatório de forma mais acessível. O Excel não tem restrição de licença e permite que qualquer pessoa reproduza, audite e entenda os cálculos diretamente nas células — o que é valioso em ambientes corporativos com restrições de TI.

**Por que VBA para navegação e não só abas do Excel?**
Múltiplas abas criam uma experiência fragmentada e pouco profissional. As macros VBA permitem simular a navegação fluida de um dashboard real, com botões visuais e transições entre "páginas" — o que aproxima a experiência do usuário de um relatório de BI sem sair do Excel.

**Por que `SUMIFS` e `AVERAGEIFS` e não Tabelas Dinâmicas?**
Tabelas Dinâmicas exigem atualização manual e têm limitações de layout. As fórmulas dinâmicas calculam em tempo real conforme os filtros mudam, dão mais controle sobre o posicionamento visual dos KPIs e são mais legíveis para auditoria.

**Trade-off aceito:** VBA reduz a portabilidade (macros podem ser bloqueadas por políticas de TI). A mitigação está documentada em `instrucoes.txt` e a migração para Power BI está prevista nos próximos passos.

---

## 8. Tecnologias Utilizadas

| Ferramenta | Uso no projeto |
|---|---|
| Excel 2016+ (XLSM) | Plataforma principal do dashboard |
| VBA (macros.bas) | Navegação entre páginas e alternância de visuais |
| SUMIFS / AVERAGEIFS | Cálculo dinâmico de KPIs com base nos filtros |
| Gráficos nativos do Excel | Visualizações de barras e linha |
| Validação de Dados | Segmentadores (Segmento, País, Ano) |
| Git & GitHub | Versionamento e portfólio público |

---

## 9. Como Executar o Projeto

```bash
# 1. Clone o repositório
git clone https://github.com/Santosdevbjj/relatoGereVendas.git
```

2. Abra o arquivo `Dashboard.xlsm` no **Microsoft Excel 2016 ou superior**.
3. Ao abrir, clique em **"Habilitar Conteúdo"** (ou "Enable Macros") na barra de aviso amarela — isso é necessário para que a navegação entre páginas funcione.
4. Use os **segmentadores** (dropdowns) para filtrar por Segmento, País e Ano.
5. Navegue entre as páginas usando os **botões do dashboard**.
6. Consulte `instrucoes.txt` para dúvidas rápidas de operação.

> **Atenção:** Se as macros forem bloqueadas pela política de TI da sua organização, consulte `instrucoes.txt` para a alternativa manual.

---

## 10. Aprendizados

O maior desafio foi fazer os KPIs recalcularem instantaneamente ao mudar os filtros sem usar Tabelas Dinâmicas. A solução foi combinar `SUMIFS` com referências absolutas aos intervalos de filtro, o que exigiu cuidado com a estrutura da planilha de dados para evitar erros de referência.

Aprendi também que **VBA precisa de documentação igual ao código de produção**: cada macro tem comentários explicando o que faz e por que, o que facilita manutenção e demonstra responsabilidade sobre o que foi produzido — algo que recrutadores técnicos valorizam ao revisar repositórios.

Por fim, construir no Excel antes do Power BI me fez entender melhor a lógica por trás dos cálculos — quando migrar para DAX, a transição foi muito mais natural.

---

## 11. Próximos Passos

- [ ] **Migrar para Power BI Desktop** — versão original do desafio, com modelagem estrela e medidas DAX
- [ ] **Publicar no Power BI Service** para compartilhamento online sem necessidade de instalação local
- [ ] **Adicionar Margem de Lucro (%)** como KPI complementar — `Lucro / Vendas * 100`
- [ ] **Criar análise de tendência** com linha de meta para comparar performance real vs. objetivo
- [ ] **Automatizar atualização de dados** via Power Query para eliminar a dependência de atualização manual

---

## Prints do Dashboard

<img width="1080" height="2400" alt="Dashboard visão mobile" src="https://github.com/user-attachments/assets/08dca4dc-e082-406d-a73a-6eea24c38fb8" />

<img width="2400" height="1080" alt="Dashboard Página 1 - Visão Geral" src="https://github.com/user-attachments/assets/9bc12cc2-8dde-4758-83c5-a08a81799540" />

<img width="2400" height="1080" alt="Dashboard Página 2 - Detalhes" src="https://github.com/user-attachments/assets/0d0a1aaa-8ab0-49bc-9e18-6eead7b83173" />

<img width="2400" height="1080" alt="Dashboard com filtros ativos" src="https://github.com/user-attachments/assets/022ab167-0bb1-4158-b9fd-2eba46022b72" />

---

## Estrutura do Repositório

```
relatoGereVendas/
├── Dashboard.xlsm      # Dashboard interativo com VBA
├── macros.bas          # Código VBA exportado para versionamento
├── instrucoes.txt      # Guia rápido de uso
└── README.md           # Este documento
```

---

**Autor:** Sérgio Santos

[![Portfólio Sérgio Santos](https://img.shields.io/badge/Portfólio-Sérgio_Santos-111827?style=for-the-badge&logo=githubpages&logoColor=00eaff)](https://portfoliosantossergio.vercel.app)
[![LinkedIn Sérgio Santos](https://img.shields.io/badge/LinkedIn-Sérgio_Santos-0A66C2?style=for-the-badge&logo=linkedin&logoColor=white)](https://linkedin.com/in/santossergioluiz)
