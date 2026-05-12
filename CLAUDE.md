# CLAUDE.md — Bem Bacana: Dashboard Financeiro e Comercial

> Documento central do projeto. Deve ser lido no início de toda sessão de trabalho relacionada ao dashboard.
> Última atualização: 2026-04-28 (rev. 5)

---

## 1. VISÃO GERAL DO PROJETO

**Empresa:** Bem Bacana Locações  
**Responsável:** Nico Prochaska (COO)  
**Objetivo:** Construir e manter um dashboard financeiro e comercial executivo, com dados consolidados de todas as verticais de negócio, orientado à tomada de decisão estratégica.

O dashboard é a principal ferramenta de controle financeiro e comercial da empresa. Deve ser simples de atualizar, rápido de ler e preciso nos números.

---

## 2. VERTICAIS DE NEGÓCIO

| Vertical         | Modelo de receita              | Nome na planilha        |
|------------------|-------------------------------|-------------------------|
| Casa             | Assinatura mensal             | Locação Casa            |
| Eventos          | Diárias por contrato          | Locação Eventos         |
| Escritório       | Contratos por período         | Locação Escritório      |
| Concierge        | Serviços avulsos/recorrentes  | Concierge               |
| Seminovos        | Venda direta (e-commerce)     | Seminovos               |

---

## 3. ARQUITETURA DE DADOS

### 3.1 Fontes de dados

O projeto combina três origens distintas, cada uma com papel específico:

| Fonte              | Arquivo / Aba              | Cobertura temporal       | Papel                                               |
|--------------------|----------------------------|--------------------------|-----------------------------------------------------|
| Planilha2          | `Dash BB…xlsx` → `Planilha2` | 2023 a out/2025        | Base histórica de faturamento (recibos pagos)       |
| ERP_CRM (antigo)   | `Dash BB…xlsx` → `ERP_CRM`  | 2024 a out/2025        | Pipeline comercial anterior (contratos/status)      |
| Eloca (novo ERP)   | `Relatorio…xlsx` → `Eloca`  | nov/2025 em diante     | Pipeline e faturamento no sistema atual             |
| Plan CAC           | `Dash BB…xlsx` → `Plan CAC` | 2024 a 2025            | Investimento em marketing e CAC mensal              |
| MRR ARR            | `Dash BB…xlsx` → `MRR ARR`  | 2024 a 2026            | Receita recorrente mensal (vertical Casa)           |

**Estratégia de corte temporal:**
- Dados anteriores a novembro de 2025 → Planilha2 + ERP_CRM (dashboard legado)
- Dados a partir de novembro de 2025 → Eloca (novo ERP)
- Zona de sobreposição (nov/2025): priorizar Eloca em caso de conflito

- **Frequência de atualização esperada:** semanal (com evolução para diária)
- **Múltiplos arquivos:** podem chegar exports diferentes com dados sobrepostos. O script consolida e deduplica antes de calcular KPIs.

### 3.2 Estrutura real — Planilha2 (faturamento histórico)

Fonte primária para faturamento. Uma única aba com todos os registros de 2023 em diante.

| Coluna real         | Mapeamento interno  | Tipo      | Observações                                         |
|---------------------|---------------------|-----------|-----------------------------------------------------|
| `DATA`              | `data_recibo`       | Data      | Data do recibo/pagamento                            |
| `CONTRATOS`         | `cliente_nome`      | Texto     | Nome do cliente                                     |
| `VALOR DO RECIBO`   | `valor_recebido`    | Numérico  | Valor efetivamente recebido                         |
| `VERTICAL`          | `vertical`          | Texto     | Ver domínio de verticais abaixo                     |
| `CLIENTE NOVO`      | `cliente_novo`      | Booleano  | True/False. Tratar "VERDADAIRO" como True (typo)    |
| `TIPO`              | `tipo_servico`      | Texto     | CORPORATIVO, SOCIAL, CASA, RENOVAÇÃO CASA, etc.     |
| `DURAÇÃO DO CONTRATO` | `duracao`         | Texto     | Campo nem sempre preenchido                         |
| `PAGAMENTO`         | `status_pagamento`  | Texto     | "OK" = pago, "Inadimplente" = inadimplente          |

**Domínio do campo `VERTICAL` na Planilha2:**

| Valor no arquivo          | Categoria no dashboard    | Observação                             |
|---------------------------|---------------------------|----------------------------------------|
| `EVENTOS`                 | Locação Eventos           | Vertical principal                     |
| `CASA`                    | Locação Casa              | Vertical de assinaturas                |
| `ESCRITÓRIO`              | Locação Escritório        | Vertical corporativa                   |
| `CONCIERGE`               | Concierge                 | Serviços                               |
| `SEMINOVOS`               | Seminovos                 | E-commerce                             |
| `COMISSÃO`                | Outras Receitas           | Agrupado em categoria auxiliar         |
| `REPOSIÇÃO/AVARIAS`       | Outras Receitas           | Agrupado em categoria auxiliar         |

> "Outras Receitas" aparece no dashboard como categoria separada, não misturada nas verticais principais.

**Faturamento histórico extraído da Planilha2:**

| Ano  | Faturamento total |
|------|-------------------|
| 2023 | R$ 759.937,67     |
| 2024 | R$ 963.071,57     |
| 2025 | R$ 1.027.579,16   |
| 2026 | R$ 388.405,67 (parcial) |

### 3.3 Estrutura real — ERP_CRM (pipeline comercial)

Fonte oficial de contratos e status do pipeline. Cobre 2024 em diante.

| Coluna real            | Mapeamento interno   | Tipo      | Observações                                        |
|------------------------|----------------------|-----------|----------------------------------------------------|
| `Nº Contrato`          | `id_contrato`        | Inteiro   | Chave única do contrato                            |
| `Cliente Faturamento`  | `cliente_nome`       | Texto     | Nome do cliente                                    |
| `Tipo de Serviço`      | `vertical`           | Texto     | Ver normalização abaixo                            |
| `Total`                | `valor_contrato`     | Numérico  | Valor total do contrato                            |
| `Status`               | `status_proposta`    | Texto     | Ver mapeamento de status abaixo                    |
| `Dt. Criação`          | `data_contrato`      | Data      | Data de criação do contrato                        |
| `Data Evento`          | `data_evento`        | Data      | Data do evento (relevante para Eventos)            |
| `Funcionário`          | —                    | Texto     | Não usado nos KPIs desta fase                      |

**Mapeamento de status ERP_CRM → dashboard:**

| Status no ERP    | Status no dashboard  |
|------------------|----------------------|
| `Fechada`        | Aprovada             |
| `Cancelada`      | Cancelada            |
| `Aguardando`     | Em negociação        |

**Normalização do campo `Tipo de Serviço`:**
O ERP tem inconsistências de grafia que devem ser normalizadas no script:
- `Vertical Eventos` / `VERTICAL EVENTOS` → EVENTOS
- `Vertical Casa` → CASA
- `Vertical Escritório` / `Vertical Escritórios` → ESCRITÓRIO
- `CONCIERGE` → CONCIERGE
- `SEMINOVOS` → SEMINOVOS
- `REPOSIÇÃO/AVARIAS` → Outras Receitas

### 3.4 Estrutura real — Eloca (novo ERP, nov/2025 em diante)

Fonte oficial do pipeline a partir da implantação do novo sistema. Exportado como arquivo XLS com uma única aba chamada `Eloca`.

| Coluna real       | Mapeamento interno   | Tipo      | Observações                                              |
|-------------------|----------------------|-----------|----------------------------------------------------------|
| `Proposta`        | `id_proposta`        | Inteiro   | ID sequencial da proposta (começa em 1000)               |
| `Nome do Cliente` | `cliente_nome`       | Texto     | Nome do cliente                                          |
| `Contrato`        | `num_contrato`       | Texto     | Número do contrato (ex: "2025-001"). Ausente se não fechado |
| `Tipo`            | `vertical`           | Texto     | Ver normalização abaixo                                  |
| `Fase Negociação` | `status_proposta`    | Texto     | Ver mapeamento de status abaixo                          |
| `Data Cadastro`   | `data_proposta`      | Texto     | Formato DD/MM/AAAA — converter para date no script       |
| `Data Início`     | `data_inicio`        | Texto     | Início do serviço/evento                                 |
| `Data Fim`        | `data_fim`           | Texto     | Fim do serviço/evento                                    |
| `Val. Serviços`   | `valor_servicos`     | Numérico  | Componente de serviços/mão de obra internos              |
| `Val. Desconto`   | `valor_desconto`     | Numérico  | Desconto aplicado                                        |
| `Val. Frete`      | `valor_frete`        | Numérico  | Mão de obra logística + frete propriamente dito          |
| `Val. Proposta`   | `valor_locacao`      | Numérico  | Valor dos produtos locados (faturamento de locação)      |

**Lógica dos campos de valor no dashboard:**

| Componente               | Campo fonte          | O que representa                        |
|--------------------------|----------------------|-----------------------------------------|
| Receita de Locação       | `Val. Proposta`      | Valor dos produtos locados              |
| Receita de Serviços/Frete| `Val. Frete`         | Mão de obra + transporte                |
| Desconto concedido       | `Val. Desconto`      | Abatimento sobre a proposta             |
| **Total faturado**       | Val. Proposta + Val. Frete − Val. Desconto | Receita bruta total   |

> O dashboard deve exibir os dois componentes separadamente (locação vs serviços/frete) para permitir análise de mix de receita.

**Verticais no Eloca — já padronizadas:**

| Valor no arquivo       | Categoria no dashboard |
|------------------------|------------------------|
| `LOCAÇÃO EVENTOS`      | Eventos                |
| `LOCAÇÃO CASA`         | Casa                   |
| `LOCAÇÃO ESCRITÓRIO`   | Escritório             |
| `VENDA DE SEMINOVOS`   | Seminovos              |

**Mapeamento de Fase Negociação → status dashboard:**

| Fase no Eloca                    | Status no dashboard | Observação                              |
|----------------------------------|---------------------|-----------------------------------------|
| `PROPOSTA APROVADA`              | Aprovada            |                                         |
| `PROPOSTA FINALIZADA`            | Aprovada            | Equivalente a Aprovada no processo      |
| `PROPOSTA RENOVADA (CASA)`       | Aprovada            | Marcada também como renovação (flag)    |
| `PROPOSTA ENVIADA`               | Enviada             |                                         |
| `EM NEGOCIAÇÃO & AJUSTES`        | Em negociação       |                                         |
| `EM ORÇAMENTO`                   | Em negociação       |                                         |
| `PROPOSTA REPROVADA`             | Reprovada           |                                         |
| `CLIENTE NÃO DEU CONTINUIDADE`   | Sem continuidade    |                                         |
| `CANCELADO`                      | Cancelada           |                                         |

> "PROPOSTA RENOVADA (CASA)" entra no pipeline como aprovada E é marcada com flag `renovacao=True` para análise separada de base recorrente da vertical Casa.

### 3.5 Entradas manuais (dados não exportados pelo ERP)

Alguns indicadores não existem na exportação do ERP e precisam ser informados manualmente pelo usuário. O agente deve receber esses dados via conversa e persistí-los no arquivo `dados/manual_input.json`.

| Dado                        | Frequência     | Como informar                                              |
|-----------------------------|----------------|------------------------------------------------------------|
| Investimento em marketing   | Mensal         | "Marketing de [mês/ano]: R$ X"                             |
| Inadimplência               | Sob demanda    | "Inadimplência de [mês/ano]: R$ X" ou "cliente X, R$ Y"   |
| Outros ajustes pontuais     | Sob demanda    | Usuário descreve o dado e o agente pergunta onde inserir   |

> Quando o usuário informar um dado manualmente, o agente deve: (1) confirmar o entendimento, (2) registrar em `manual_input.json`, (3) recalcular os KPIs afetados, (4) atualizar o dashboard.

**Estrutura do `manual_input.json`:**

```json
{
  "marketing": {
    "2026-01": 5000.00,
    "2026-02": 5000.00
  },
  "inadimplencia": {
    "2026-01": { "total": 3200.00, "detalhes": "Cliente X: R$ 1.500, Cliente Y: R$ 1.700" },
    "2026-02": { "total": 0, "detalhes": "" }
  }
}
```

---

## 4. KPIs E MÉTRICAS

### 4.1 Receita

| KPI                              | Cálculo                                              | Periodicidade |
|----------------------------------|------------------------------------------------------|---------------|
| Faturamento mensal total         | SUM(valor_recebido) no mês                          | Mensal        |
| Faturamento acumulado no ano     | SUM(valor_recebido) desde jan do ano corrente       | Acumulado     |
| Faturamento por vertical (mês)   | SUM(valor_recebido) agrupado por vertical           | Mensal        |
| Faturamento por vertical (ano)   | SUM(valor_recebido) YTD agrupado por vertical       | Acumulado     |

### 4.2 Meta

| KPI                              | Valor / Cálculo                                    | Observação              |
|----------------------------------|----------------------------------------------------|-------------------------|
| Meta mensal fixa                 | R$ 150.000,00                                      | Revisável conforme crescimento |
| % de atingimento                 | (faturamento_mês / meta) × 100                    | Exibir com indicador visual |
| Desvio da meta                   | faturamento_mês - meta                             | Positivo = superou      |

### 4.3 Clientes

| KPI                              | Cálculo                                            |
|----------------------------------|----------------------------------------------------|
| Clientes novos no mês            | COUNT(cliente) WHERE cliente_tipo = 'Novo'         |
| Clientes recorrentes no mês      | COUNT(cliente) WHERE cliente_tipo = 'Recorrente'   |
| Total de clientes ativos         | COUNT(clientes distintos com receita no mês)       |

### 4.4 Comercial (pipeline de propostas)

| KPI                              | Cálculo                                            |
|----------------------------------|----------------------------------------------------|
| Propostas aprovadas (contratos)  | COUNT WHERE status = 'Aprovada'                    |
| Propostas enviadas               | COUNT WHERE status = 'Enviada'                     |
| Propostas em negociação          | COUNT WHERE status = 'Em negociação'               |
| Propostas reprovadas             | COUNT WHERE status = 'Reprovada'                   |
| Taxa de conversão                | aprovadas / (aprovadas + reprovadas) × 100         |

### 4.5 CAC (Custo de Aquisição de Cliente)

| KPI        | Cálculo                                            | Dado externo necessário       |
|------------|----------------------------------------------------|-------------------------------|
| CAC mensal | investimento_marketing / clientes_novos_mes        | Investimento em marketing (manual) |

> O investimento em marketing deve ser informado manualmente em uma célula de input no dashboard ou em uma aba auxiliar da planilha.

### 4.6 Receita por Cliente

| KPI                  | Cálculo                                              |
|----------------------|------------------------------------------------------|
| Ranking top clientes | SUM(valor_recebido) por cliente, ordenado DESC       |
| Ticket médio global  | faturamento_total / total_clientes_ativos            |

### 4.7 Ticket Médio por Vertical

| KPI                                   | Cálculo                                                              | Periodicidade |
|---------------------------------------|----------------------------------------------------------------------|---------------|
| Ticket médio por vertical (mês)       | SUM(valor_recebido) por vertical / COUNT(clientes ativos) no mês    | Mensal        |
| Ticket médio por vertical (ano acum.) | SUM(valor_recebido) YTD por vertical / COUNT(clientes únicos) YTD   | Acumulado     |

> Exibir no dashboard como tabela ou cards comparativos por vertical, com evolução mês a mês.

### 4.8 Inadimplência

| KPI                          | Cálculo / Fonte                                                      | Periodicidade |
|------------------------------|----------------------------------------------------------------------|---------------|
| Inadimplência total (mês)    | Informada manualmente via `manual_input.json`                        | Mensal        |
| % inadimplência sobre fat.   | inadimplencia_mes / faturamento_mes × 100                           | Mensal        |
| Detalhamento por cliente     | Texto descritivo informado pelo usuário (opcional)                   | Sob demanda   |

> Se inadimplência não for informada para o mês, o campo exibe "não informado" no dashboard. Nunca inferir ou estimar esse valor.

### 4.9 Comparativos

| KPI                  | Cálculo                                              |
|----------------------|------------------------------------------------------|
| YoY (ano vs ano)     | (fat_mes_atual - fat_mesmo_mes_ano_anterior) / fat_ano_anterior × 100 |
| MoM (mês vs mês)     | (fat_mes_atual - fat_mes_anterior) / fat_mes_anterior × 100 |

---

## 5. STACK TÉCNICA

| Componente           | Tecnologia           | Observação                                      |
|----------------------|----------------------|-------------------------------------------------|
| Formato do dashboard | HTML interativo      | Arquivo único .html abrindo no navegador        |
| Processamento dados  | Python               | Script para ler XLS e gerar JSON ou HTML        |
| Gráficos             | Chart.js ou Plotly.js| Biblioteca JS embutida no HTML                  |
| Planilhas de entrada | XLS/CSV (export ERP) | Processadas por script Python                   |
| Armazenamento        | Arquivos locais      | Sem banco de dados por enquanto                 |

### 5.1 Fluxo de atualização

```
ERP (exporta XLS/CSV)
       ↓
Usuário entrega arquivo(s) ao agente
       ↓
Script Python (consolida múltiplos arquivos + deduplica + calcula KPIs)
       ↓
Merge com manual_input.json (marketing, inadimplência, ajustes)
       ↓
Gera dashboard.html atualizado
       ↓
Nico abre dashboard.html no navegador
```

### 5.2 Arquivos do projeto

```
BEM BACANA - DASH FINANCEIRO/
├── CLAUDE.md                          ← este arquivo (cérebro do projeto)
├── dados/
│   ├── Dash BB 23_24_25_26_v2 (5).xlsx  ← planilha histórica (Planilha2 + ERP_CRM + Plan CAC + MRR ARR)
│   ├── Relatorio (14).xlsx               ← export Eloca mais recente (detecção automática por nome)
│   ├── ftp601a1.xlsx                     ← export de faturas emitidas do Eloca
│   └── manual_input.json                 ← dados manuais (marketing, inadimplência, propostas preliminares)
├── scripts/
│   └── gerar_dashboard.py               ← script único: lê todas as fontes, calcula KPIs, gera HTML
├── output/
│   └── dashboard.html                   ← arquivo final entregue ao usuário
└── historico/
    └── dados_AAAA-MM.json               ← snapshots históricos (uso futuro)
```

---

## 6. IDENTIDADE VISUAL DO DASHBOARD

| Elemento          | Definição                                                     |
|-------------------|---------------------------------------------------------------|
| Fundo             | `#FFFFFF` branco puro                                                 |
| Logo              | Logo horizontal Bem Bacana (PNG fornecido), sobre fundo branco        |
| Cor primária      | `#076A76` — teal escuro (barra superior do logo, texto da marca)      |
| Cor secundária    | `#41A8B9` — teal claro (barra inferior do logo)                       |
| Cor de destaque   | `#FBAE4B` — âmbar/laranja (barra central do logo, alertas e ênfases) |
| Cor de texto      | `#61605B` — cinza escuro (corpo de texto)                             |
| Preto             | `#000000` — uso pontual                                               |
| Tipografia        | Avenir (Black / Roman / Book). Web fallback: Nunito, DM Sans, sans-serif |
| KPIs destacados   | Cards grandes, número central, variação verde (#28A745) / vermelho (#DC3545) |
| Gráficos          | Linha para série temporal, barra para comparativos verticais          |
| Alertas visuais   | Abaixo da meta: `#FBAE4B` (âmbar) ou vermelho `#DC3545`              |
| KPIs destacados   | Cards grandes, número central, variação em cor (verde/vermelho)|
| Gráficos          | Linha para série temporal, barra para comparativos verticais  |
| Alertas visuais   | Abaixo da meta: destaque em vermelho ou laranja               |

---

## 7. REGRAS DE NEGÓCIO E CONVENÇÕES

1. **Competência vs caixa:** usar a coluna `DATA` da Planilha2 (data do recibo) como base de faturamento mensal. Para pipeline, usar `Dt. Criação` da ERP_CRM.
2. **Meta:** R$ 150.000/mês é a meta fixa inicial. Quando revisada, atualizar tanto aqui quanto no script.
3. **Vertical Seminovos:** o faturamento é a receita de venda bruta. Margem só calculada se custo estiver disponível.
4. **Cliente novo vs recorrente:** lógica cronológica por nome. Um cliente é NOVO se seu nome nunca apareceu em nenhum mês anterior (Planilha2 ou Eloca). É RECORRENTE se já apareceu. A base `vistos_global` é carregada em ordem cronológica unificada, portanto clientes da Planilha2 2023-2025 entram no histórico antes dos clientes Eloca 2026. Nota: discrepâncias de formatação de nome entre sistemas podem classificar incorretamente clientes repetidos — dado de qualidade conhecida.
5. **Consolidação de múltiplos arquivos:** chave de deduplicação na Planilha2: `(cliente_nome, data_recibo, vertical, valor_recebido)`. Na ERP_CRM: `Nº Contrato` é chave única. Em conflito, prevalece o arquivo mais recente.
6. **Categorias de faturamento:** as verticais principais são EVENTOS, CASA, ESCRITÓRIO, CONCIERGE e SEMINOVOS. Os lançamentos de COMISSÃO e REPOSIÇÃO/AVARIAS são agrupados como "Outras Receitas" — aparecem no total consolidado mas em categoria separada nos gráficos por vertical.
7. **Status ERP_CRM:** Fechada = Aprovada, Cancelada = Cancelada, Aguardando = Em negociação. Nunca inverter esse mapeamento.
8. **Normalização de verticais no ERP_CRM:** o campo Tipo de Serviço tem grafias inconsistentes. Sempre normalizar antes de agrupar (ver seção 3.3).
9. **CAC:** só calcular se o investimento em marketing do mês estiver preenchido. Se ausente, exibir "dado indisponível" no dashboard. A aba Plan CAC da planilha histórica é referência para 2024/2025.
10. **Inadimplência:** nunca inferir ou calcular automaticamente. O campo `PAGAMENTO = Inadimplente` na Planilha2 é informativo. Valor consolidado deve ser confirmado pelo usuário e gravado em `manual_input.json`.
11. **Dados manuais:** ao receber qualquer dado manual (marketing, inadimplência, ajuste), confirmar o entendimento antes de gravar. Sempre registrar em `manual_input.json` com chave `AAAA-MM`.
12. **Ticket médio por vertical:** calcular para todos os meses disponíveis. Permitir visualização histórica da evolução por vertical.
13. **Performance por vendedor:** não incluída nesta fase. Campo `Funcionário` da ERP_CRM reservado para fase futura.

---

## 8. HISTÓRICO DE DECISÕES

| Data       | Decisão                                              | Motivo                                   |
|------------|------------------------------------------------------|------------------------------------------|
| 2026-04-27 | Dashboard em HTML interativo (não Power BI)          | Zero dependência externa, abre em qualquer máquina |
| 2026-04-27 | Uma aba por vertical no XLS                          | Estrutura já existente no ERP            |
| 2026-04-27 | Meta fixa em R$ 150.000/mês                          | Referência inicial de gestão             |
| 2026-04-27 | Suporte a múltiplos XLS do mesmo período             | ERP pode gerar exports diferentes com dados sobrepostos |
| 2026-04-27 | Dados manuais (marketing, inadimplência) via conversa | Esses dados não existem no ERP; usuário informa quando necessário |
| 2026-04-27 | Ticket médio calculado por vertical (mês e ano)      | Necessidade de análise de performance granular por negócio |
| 2026-04-27 | Inadimplência informada manualmente, nunca estimada  | Dado sensível; só exibir quando confirmado pelo usuário |
| 2026-04-27 | Duas fontes de dados: Planilha2 (faturamento) + ERP_CRM (pipeline) | Estruturas diferentes com papéis complementares |
| 2026-04-27 | COMISSÃO e REPOSIÇÃO/AVARIAS agrupados em "Outras Receitas" | Não são verticais de negócio, mas têm impacto financeiro |
| 2026-04-27 | Mapeamento status ERP: Fechada=Aprovada, Cancelada=Cancelada, Aguardando=Em negociação | Nomenclatura do ERP diverge do padrão do dashboard |
| 2026-04-27 | Performance por vendedor não incluída nesta fase     | Foco nos KPIs de negócio; campo Funcionário reservado para fase futura |
| 2026-04-27 | Novo ERP (Eloca) com estrutura diferente do anterior | Migração de sistema. Corte em nov/2025: dados anteriores da Planilha2, novos do Eloca |
| 2026-04-27 | Val. Proposta = locação, Val. Frete = serviços/logística | Dashboard exibe ambos separadamente para análise de mix de receita |
| 2026-04-27 | PROPOSTA RENOVADA (CASA) = Aprovada + flag renovação | Computa como fechada mas é rastreável como renovação de longo prazo |
| 2026-04-27 | Fundo branco + cores e logo do brand guide Bem Bacana | Identidade visual definida para o dashboard HTML |
| 2026-04-28 | Mix de receita: Total = Val.Proposta; Locação = Val.Proposta − Val.Serviços; Serviço/Frete = Val.Serviços | Val.Serviços já está embutido em Val.Proposta — não somar os dois |
| 2026-04-28 | Relatorio (14).xlsx com 9 colunas (sem Contrato, Val.Desconto, Val.Frete) | Script detecta formato automaticamente por presença da coluna 'contrato' no header |
| 2026-04-28 | ftp601a1.xlsx adicionado como fonte de Faturas Emitidas | Export de faturamento do Eloca; complementado por lançamentos manuais (NFe/DANFe/Recibo) |
| 2026-04-28 | Propostas Preliminares: entrada manual por mês/vertical/valor com status 'Preliminar' | Somam no total de propostas abertas mas não entram no faturamento |
| 2026-04-28 | Corte dupla contagem: registros Eloca com ano < 2026 ignorados no faturamento | Planilha2 cobre até dez/2025; Eloca entra apenas a partir de jan/2026 |
| 2026-04-28 | Script detecta arquivo Relatorio mais recente por ordem alfabética reversa | Evita uso acidental de versão antiga quando múltiplos exports existem na pasta dados/ |

---

## 9. PRÓXIMOS PASSOS (ROADMAP)

### Fase 1 — Fundação (atual)
- [x] Criar CLAUDE.md com arquitetura do projeto
- [x] Receber planilha histórica e mapear colunas reais (Planilha2 + ERP_CRM)
- [x] Receber export do novo ERP (Eloca) e mapear estrutura
- [x] Registrar faturamento histórico 2023-2026 no CLAUDE.md
- [x] Definir identidade visual: fundo branco + brand guide Bem Bacana + logo
- [x] Receber logo e cores do brand guide para incorporar ao dashboard
- [x] Criar script Python para processar e consolidar Planilha2 + Eloca
- [x] Criar primeira versão do dashboard HTML com KPIs principais

### Fase 2 — Consolidação
- [x] Adicionar série histórica (comparativos YoY e MoM)
- [x] Implementar ranking de clientes
- [x] Adicionar módulo de CAC com input de marketing
- [x] Refinar identidade visual com paleta Bem Bacana
- [x] Análise de entregas: concentração semanal, duração média, simultâneas, antecedência
- [x] Visão Macro de propostas (todas as fases, quantidade e valor por mês)
- [x] Faturas Emitidas por tipo de receita (ftp601a1 + entrada manual NFe/DANFe/Recibo)
- [x] Propostas Preliminares: entrada manual, status separado, soma no total geral
- [x] Mix de receita: Locação vs Serviço/Frete por mês

### Fase 3 — Evolução
- [ ] Metas por vertical
- [ ] Análise de pipeline comercial detalhada (funil visual, tempo médio de conversão)
- [ ] Performance por vendedor (campo Funcionário do ERP_CRM)
- [ ] Exportação automática do ERP via script agendado
- [ ] Análise preditiva de faturamento
- [ ] Integração direta com Eloca via API (se disponível)

---

## 10. INSTRUÇÕES PARA O AGENTE

Ao iniciar qualquer sessão de trabalho neste projeto:

1. Ler este CLAUDE.md completo antes de qualquer ação.
2. Verificar em qual fase do roadmap o projeto está (seção 9).
3. Se uma planilha XLS for fornecida, inspecionar as colunas reais antes de processar e atualizar a seção 3.3 se houver divergência.
4. Nunca inventar dados. Se um campo não existir na planilha, registrar a ausência e solicitar ao usuário.
5. Sempre que uma decisão relevante for tomada durante a sessão, adicionar ao histórico da seção 8.
6. O output final sempre deve ser salvo em `output/dashboard.html` e apresentado ao usuário com link direto.
7. Ao calcular KPIs, seguir rigorosamente as fórmulas da seção 4. Qualquer desvio deve ser documentado.
