WaterERP SAE - Gestão de Caminhão Pipa
O WaterERP SAE é uma solução Single-Page Application (SPA) de alta performance desenvolvida para o controle total de operações de transporte de água potável/reuso. Focado em ativos únicos (caminhão pipa), o sistema centraliza o controle de horímetros, quilometragem, insumos e faturamento em uma interface moderna com padrão visual Glassmorphism.
1. Visão Geral do Projeto

Objetivo: Digitalizar a operação de caminhão pipa, eliminando formulários em papel e planilhas desconexas.

Stack: Frontend (Vue 3 via CDN, Tailwind CSS) | Backend (Google Apps Script) | Banco de Dados (Google Sheets).

Arquitetura: Padrão SAE (Single File WebApp) com lógica de CRUD e normalização centralizada no servidor.

2. Requisitos Funcionais (RF)

RF01 - Gestão Operacional

Lançamento de Viagens: Registro de data, local de entrega, volume (m³).

Controle de Movimentação: Registro obrigatório de KM inicial/final e Horímetro inicial/final por viagem.

Cálculo Automático: O sistema deve calcular automaticamente o KM percorrido e as horas trabalhadas (delta).

RF02 - Gestão de Custos (Financeiro)

Abastecimentos: Registro de litros, valor por litro, posto e odômetro no momento da carga.

Despesas do Caminhão: Manutenções preventivas e corretivas (troca de óleo, pneus, filtros).

Despesas do Contrato: Custos administrativos vinculados à operação.

RF03 - Dashboards e KPIs

Consumo Médio: Cálculo de KM/L e L/Hora (Eficiência Energética).

Produtividade: Total de m³ transportados e total de viagens no mês.

Financeiro: Gráfico de composição de custos e evolução de despesas vs. volume entregue.

RF04 - Relatórios Empresariais

Fechamento Mensal: Geração de documento formatado para impressão (PDF) contendo:

Resumo de KM e Horímetro (Inicial/Final/Total).

Listagem detalhada de entregas por cliente/local.

Demonstrativo financeiro consolidado.

3. Requisitos Não Funcionais (RNF)

RNF01 - Confiabilidade (Backend-First): Toda validação e normalização de dados deve ocorrer no Codigo.js. O frontend não deve ser confiado para cálculos críticos.

RNF02 - Persistência: Utilização de Google Sheets como banco de dados NoSQL-like (serializado), com suporte a logs de auditoria (quem inseriu e quando).

RNF03 - Interface (UI/UX): Padrão SAE: Glassmorphism dark mode, responsivo para dispositivos móveis (mobile-first).

RNF04 - Performance: Tempo de carregamento inferior a 2 segundos (uso de Skeleton Loading e otimização de chamadas google.script.run).

RNF05 - Auditoria: Cada registro deve possuir um UUID único e rastreabilidade por e-mail do operador.

4. Estrutura de Dados (Google Sheets)

Aba

Campos Principais

entregas

uuid, timestamp, data, local, volume_m3, km_inic, km_fim, h_inic, h_fim, user

abastecimentos

uuid, timestamp, data, litros, valor_total, km_atual, posto, user

despesas

uuid, timestamp, data, categoria, descricao, valor, user

SYS_LOGS

timestamp, acao, detalhes, user

5. Instruções de Instalação

Crie uma nova Planilha Google.

Acesse Extensões > Apps Script.

Cole o conteúdo de Codigo.js no editor de script.

Crie um novo arquivo HTML chamado index.html e cole o código do frontend.

Clique em Implantar > Nova Implantação como App da Web.

Configure para "Executar como: Eu" e "Quem tem acesso: Qualquer pessoa com conta Google".

6. Padrão SAE (Apollo Enterprise)

Este sistema segue as diretrizes de desenvolvimento para baixo custo de manutenção e alta escalabilidade dentro do ecossistema Google Workspace.
