// Gera as páginas do dashboard bi_voos (PBIR). Uso: node gerar_paginas_bi_voos.js
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const REPORT = path.join(__dirname, 'case claude', 'bi_voos.Report', 'definition');
const PAGES_DIR = path.join(REPORT, 'pages');
const SCHEMA_PAGE = 'https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.1.0/schema.json';
const SCHEMA_VISUAL = 'https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.9.0/schema.json';

// Recorte do LOGO_AZUL_LINHAS_AEREAS.png (sem margens), registrado em StaticResources/RegisteredResources.
const LOGO = 'LOGO_AZUL_LINHAS_AEREAS17591500000000000.png';

const F = 'airline_passenger_satisfaction';
const A = 'Avaliações';
const MT = '_Medidas';

const COR = {
  titulo: '#0B2D5B',
  primaria: '#1565C0',
  secundaria: '#8FB4E3',
  destaque: '#E0703A',
};

// ---------- helpers ----------
const hex = (n) => crypto.createHash('md5').update(n).digest('hex').slice(0, 20); // IDs estáveis entre execuções
const lit = (v) => ({ expr: { Literal: { Value: v } } });
const str = (s) => lit(`'${s.replace(/'/g, "''")}'`);
const num = (n) => lit(`${n}D`);
const bool = (b) => lit(b ? 'true' : 'false');
const cor = (c) => ({ solid: { color: str(c) } });

const M = (name) => ({
  field: { Measure: { Expression: { SourceRef: { Entity: MT } }, Property: name } },
  queryRef: `${MT}.${name}`,
  nativeQueryRef: name,
});
const C = (table, col) => ({
  field: { Column: { Expression: { SourceRef: { Entity: table } }, Property: col } },
  queryRef: `${table}.${col}`,
  nativeQueryRef: col,
});
const sortMeasure = (name, direction = 'Descending') => ({
  sort: [{ field: M(name).field, direction }],
  isDefaultSort: true,
});
const sortColumn = (table, col, direction = 'Ascending') => ({
  sort: [{ field: C(table, col).field, direction }],
  isDefaultSort: true,
});

function vco({ title, background = true } = {}) {
  const o = {
    padding: [{ properties: { top: num(8), bottom: num(8), left: num(8), right: num(8) } }],
    border: [{ properties: { show: bool(true), color: cor('#E1E6EE'), radius: num(6) } }],
    dropShadow: [{ properties: { show: bool(false) } }],
  };
  if (background) o.background = [{ properties: { show: bool(true), color: cor('#FFFFFF'), transparency: num(0) } }];
  if (title) {
    o.title = [{ properties: { show: bool(true), text: str(title), fontSize: num(11), bold: bool(true), fontColor: cor(COR.titulo) } }];
  } else {
    o.title = [{ properties: { show: bool(false) } }];
  }
  return o;
}

let pageVisuals;
let z;
function add(key, pos, visual) {
  z += 1000;
  const name = hex(`${key}`);
  pageVisuals.push({
    $schema: SCHEMA_VISUAL,
    name,
    position: { x: pos[0], y: pos[1], z, height: pos[3], width: pos[2], tabOrder: z },
    visual,
  });
  return name;
}

// ---------- componentes ----------
function header(pageKey, titulo, subtitulo) {
  add(`${pageKey}-header`, [0, 0, 1280, 48], {
    visualType: 'textbox',
    objects: {
      general: [{
        properties: {
          paragraphs: [{
            textRuns: [
              { value: titulo, textStyle: { fontFamily: 'Segoe UI Semibold', fontSize: '18px', color: '#FFFFFF' } },
              { value: `   |   ${subtitulo}`, textStyle: { fontFamily: 'Segoe UI', fontSize: '12px', color: '#C9D6EA' } },
            ],
            horizontalTextAlignment: 'left',
          }],
        },
      }],
    },
    visualContainerObjects: {
      background: [{ properties: { show: bool(true), color: cor(COR.titulo), transparency: num(0) } }],
      border: [{ properties: { show: bool(false) } }],
      padding: [{ properties: { top: num(10), bottom: num(6), left: num(176), right: num(16) } }],
    },
  });
  // Logo sobre placa branca (o texto do logo é azul-marinho e sumiria no fundo do cabeçalho).
  add(`${pageKey}-logo`, [8, 4, 152, 40], {
    visualType: 'image',
    objects: {
      general: [{
        properties: {
          imageUrl: { expr: { ResourcePackageItem: { PackageName: 'RegisteredResources', PackageType: 1, ItemName: LOGO } } },
        },
      }],
      image: [{ properties: { fit: str('Fit'), altText: str('Logo Azul Linhas Aéreas') } }],
    },
    visualContainerObjects: {
      background: [{ properties: { show: bool(true), color: cor('#FFFFFF'), transparency: num(0) } }],
      border: [{ properties: { show: bool(false), radius: num(6) } }],
      padding: [{ properties: { top: num(6), bottom: num(6), left: num(10), right: num(10) } }],
      title: [{ properties: { show: bool(false) } }],
    },
    drillFilterOtherVisuals: true,
  });
}

const SLICERS = [
  [F, 'Class', 'Classe'],
  [F, 'Type of Travel', 'Tipo de Viagem'],
  [F, 'Customer Type', 'Tipo de Cliente'],
  [F, 'Gender', 'Gênero'],
  [F, 'Faixa Etária', 'Faixa Etária'],
];
function slicers(pageKey) {
  SLICERS.forEach(([t, c, label], i) => {
    add(`${pageKey}-slicer-${c}`, [16 + i * 252, 56, 240, 80], {
      visualType: 'slicer',
      syncGroup: { groupName: `sync_${c.replace(/\W/g, '')}`, fieldChanges: true, filterChanges: true },
      query: { queryState: { Values: { projections: [C(t, c)] } } },
      objects: {
        data: [{ properties: { mode: str('Dropdown') } }],
        header: [{ properties: { show: bool(true), text: str(label) } }],
      },
      visualContainerObjects: {
        padding: [{ properties: { top: num(8), bottom: num(8), left: num(8), right: num(8) } }],
      },
    });
  });
}

function kpiCard(key, pos, measures, { valueSize = 20, labelSize = 11, wrap = false } = {}) {
  const valueProps = { fontSize: num(valueSize), fontColor: cor(COR.titulo) };
  if (wrap) valueProps.textWrap = bool(true);
  add(key, pos, {
    visualType: 'cardVisual',
    query: { queryState: { Data: { projections: measures.map(M) } } },
    objects: {
      value: [{ properties: valueProps, selector: { id: 'default' } }],
      label: [{ properties: { fontSize: num(labelSize) }, selector: { id: 'default' } }],
      padding: [{ properties: { paddingUniform: num(8) }, selector: { id: 'default' } }],
      layout: [{ properties: { paddingUniform: num(0) }, selector: { id: 'default' } }],
      accentBar: [{ properties: { show: bool(true), color: cor(COR.primaria) }, selector: { id: 'default' } }],
    },
    visualContainerObjects: {
      padding: [{ properties: { top: num(8), bottom: num(8), left: num(8), right: num(8) } }],
      spacing: [{ properties: { customizeSpacing: bool(true), verticalSpacing: num(2) }, selector: { id: 'default' } }],
      border: [{ properties: { show: bool(true), color: cor('#E1E6EE'), radius: num(6) } }],
    },
  });
}

function barChart(key, pos, type, { category, series, y, title, sort, color = COR.primaria, labels = true, seriesColors }) {
  const qs = { Category: { projections: [category] }, Y: { projections: y.map(M) } };
  if (series) qs.Series = { projections: [series] };
  const objects = {
    labels: [{ properties: { show: bool(labels) } }],
  };
  if (!series && y.length === 1) objects.dataPoint = [{ properties: { defaultColor: cor(color) } }];
  if (series && seriesColors) {
    // Uma cor por valor da série (legenda).
    objects.dataPoint = Object.entries(seriesColors).map(([value, c]) => ({
      properties: { fill: cor(c) },
      selector: {
        data: [{
          scopeId: {
            Comparison: { ComparisonKind: 0, Left: series.field, Right: { Literal: { Value: `'${value}'` } } },
          },
        }],
      },
    }));
  }
  const visual = { visualType: type, query: { queryState: qs }, objects, visualContainerObjects: vco({ title }) };
  if (sort) visual.query.sortDefinition = sort;
  add(key, pos, visual);
}

function comboChart(key, pos, { category, col, line, title, sort }) {
  add(key, pos, {
    visualType: 'lineClusteredColumnComboChart',
    query: {
      queryState: {
        Category: { projections: [category] },
        Y: { projections: [M(col)] },
        Y2: { projections: [M(line)] },
      },
      sortDefinition: sort,
    },
    objects: {
      dataPoint: [
        { properties: { fill: cor(COR.secundaria) }, selector: { metadata: `${MT}.${col}` } },
        { properties: { fill: cor(COR.destaque) }, selector: { metadata: `${MT}.${line}` } },
      ],
      labels: [{ properties: { show: bool(true) } }],
    },
    visualContainerObjects: vco({ title }),
  });
}

function matrix(key, pos, { rows, columns, values, title }) {
  const qs = { Rows: { projections: rows }, Values: { projections: values.map(M) } };
  if (columns) qs.Columns = { projections: columns };
  add(key, pos, {
    visualType: 'pivotTable',
    query: { queryState: qs },
    objects: {
      columnHeaders: [{ properties: { columnAdjustment: str('growToFit'), autoSizeColumnWidth: bool(true) } }],
    },
    visualContainerObjects: vco({ title }),
  });
}

// ---------- páginas ----------
const PAGES = [];
function page(key, displayName, subtitulo, build) {
  pageVisuals = [];
  z = 0;
  header(key, displayName, subtitulo);
  slicers(key);
  build(key);
  PAGES.push({ name: hex(`page-${key}`), displayName, visuals: pageVisuals });
}

// Grade: cabeçalho 0-48, segmentações 56-136, KPIs 144-226, área principal 236-712.
const ROW1 = 236, ROW2 = 478, H_HALF = 234, H_FULL = 476;

page('visao-geral', 'Visão Geral', 'Satisfação dos passageiros: indicadores-chave', (k) => {
  kpiCard(`${k}-kpis`, [16, 144, 1248, 82], [
    'Total Passageiros', '% Satisfação', 'Δ Satisfação vs Geral (p.p.)',
    '% Clientes Recorrentes', '% Pontualidade Chegada (OTP)', 'Nota Média',
  ]);
  add(`${k}-donut`, [16, ROW1, 304, H_HALF], {
    visualType: 'donutChart',
    query: { queryState: { Category: { projections: [C(F, 'Status Satisfação')] }, Y: { projections: [M('Total Passageiros')] } } },
    objects: { labels: [{ properties: { show: bool(true), labelStyle: str('Percent of total') } }], legend: [{ properties: { show: bool(true), position: str('Bottom') } }] },
    visualContainerObjects: vco({ title: 'Distribuição da satisfação' }),
  });
  barChart(`${k}-sat-viagem`, [328, ROW1, 460, H_HALF], 'clusteredColumnChart', {
    category: C(F, 'Type of Travel'), series: C(F, 'Customer Type'), y: ['% Satisfação'],
    title: '% Satisfação por tipo de viagem e de cliente',
  });
  barChart(`${k}-sat-classe`, [796, ROW1, 468, H_HALF], 'clusteredBarChart', {
    category: C(F, 'Class'), y: ['% Satisfação'], title: '% Satisfação por classe', sort: sortMeasure('% Satisfação'),
  });
  kpiCard(`${k}-insights`, [16, ROW2, 620, H_HALF], [
    'Resumo Satisfação', 'Principal Driver de Satisfação', 'Maior Oportunidade de Melhoria', 'Serviço Pior Avaliado',
  ], { valueSize: 12, labelSize: 10, wrap: true });
  comboChart(`${k}-sat-idade`, [644, ROW2, 620, H_HALF], {
    category: C(F, 'Faixa Etária'), col: 'Total Passageiros', line: '% Satisfação',
    title: 'Passageiros e % Satisfação por faixa etária', sort: sortColumn(F, 'Faixa Etária'),
  });
});

page('perfil', 'Perfil do Passageiro', 'Quem são os passageiros e como se distribui a satisfação', (k) => {
  kpiCard(`${k}-kpis`, [16, 144, 1248, 82], [
    'Idade Média', 'Idade Mediana', 'Distância Média (mi)', 'Distância Média (km)', '% Viagens a Negócios', '% Classe Business',
  ]);
  comboChart(`${k}-idade`, [16, ROW1, 616, H_HALF], {
    category: C(F, 'Faixa Etária'), col: 'Total Passageiros', line: '% Satisfação',
    title: 'Faixa etária: volume e % Satisfação', sort: sortColumn(F, 'Faixa Etária'),
  });
  comboChart(`${k}-distancia`, [640, ROW1, 624, H_HALF], {
    category: C(F, 'Faixa Distância'), col: 'Total Passageiros', line: '% Satisfação',
    title: 'Distância do voo: volume e % Satisfação', sort: sortColumn(F, 'Faixa Distância'),
  });
  matrix(`${k}-matriz`, [16, ROW2, 616, H_HALF], {
    rows: [C(F, 'Class')], columns: [C(F, 'Type of Travel')], values: ['% Satisfação'],
    title: '% Satisfação: classe x tipo de viagem',
  });
  barChart(`${k}-cliente`, [640, ROW2, 624, H_HALF], 'hundredPercentStackedBarChart', {
    category: C(F, 'Customer Type'), series: C(F, 'Status Satisfação'), y: ['Total Passageiros'],
    title: 'Composição da satisfação por tipo de cliente',
  });
});

page('pontualidade', 'Pontualidade', 'Atrasos de partida e chegada e seu efeito na satisfação', (k) => {
  kpiCard(`${k}-kpis`, [16, 144, 1248, 82], [
    '% Pontualidade Partida (OTP)', '% Pontualidade Chegada (OTP)', 'Atraso Médio Chegada (min)',
    'Atraso P90 Chegada (min)', '% Atrasos Severos na Chegada (> 60 min)', 'Recuperação Média de Atraso (min)',
  ]);
  comboChart(`${k}-faixa`, [16, ROW1, 780, H_HALF], {
    category: C(F, 'Faixa Atraso Partida'), col: 'Total Passageiros', line: '% Satisfação',
    title: 'Faixa de atraso na partida: volume e % Satisfação', sort: sortColumn(F, 'Faixa Atraso Partida'),
  });
  kpiCard(`${k}-impacto`, [804, ROW1, 460, H_HALF], [
    'Gap Satisfação Pontual vs Atrasado (p.p.)', 'Correlação Atraso Chegada x Satisfação', '% Voos que Recuperaram Atraso',
  ], { valueSize: 18, labelSize: 10 });
  matrix(`${k}-classe`, [16, ROW2, 780, H_HALF], {
    rows: [C(F, 'Class')],
    values: ['Atraso Médio Partida (min)', 'Atraso Médio Chegada (min)', '% Pontualidade Chegada (OTP)', '% Satisfação - Chegada Pontual', '% Satisfação - Chegada Atrasada'],
    title: 'Pontualidade e satisfação por classe',
  });
  barChart(`${k}-status`, [804, ROW2, 460, H_HALF], 'clusteredBarChart', {
    category: C(F, 'Status Pontualidade Chegada'), y: ['% Satisfação'], title: '% Satisfação por pontualidade na chegada',
    sort: sortMeasure('% Satisfação'),
  });
});

page('avaliacoes', 'Avaliações de Serviço', 'Notas de 1 a 5 por serviço (0 = não aplicável, excluído)', (k) => {
  kpiCard(`${k}-kpis`, [16, 144, 1248, 82], [
    'Nota Média', '% Top-2-Box (Notas 4-5)', '% Bottom-2-Box (Notas 1-2)', 'Saldo Líquido de Avaliação',
    '% Não Aplicável', '% Passageiros Encantados',
  ]);
  barChart(`${k}-nota`, [16, ROW1, 400, H_FULL], 'clusteredBarChart', {
    category: C(A, 'Serviço'), y: ['Nota Média'], title: 'Nota média por serviço', sort: sortMeasure('Nota Média'),
  });
  barChart(`${k}-distribuicao`, [424, ROW1, 440, H_FULL], 'hundredPercentStackedBarChart', {
    category: C(A, 'Serviço'), series: C(A, 'Classificação Nota'), y: ['Avaliações Válidas'],
    title: 'Distribuição das notas por serviço', labels: false, sort: sortColumn(A, 'Serviço'),
    seriesColors: {
      'Negativa (1-2)': '#D64545',
      'Neutra (3)': '#A6A6A6',
      'Positiva (4-5)': '#2E9E5B',
    },
  });
  matrix(`${k}-categoria`, [872, ROW1, 392, H_FULL], {
    rows: [C(A, 'Categoria')],
    values: ['Nota Média', '% Top-2-Box (Notas 4-5)', 'Saldo Líquido de Avaliação', '% Não Aplicável'],
    title: 'Resumo por categoria de serviço',
  });
});

page('drivers', 'Drivers de Satisfação', 'Quais serviços mais explicam a satisfação geral', (k) => {
  kpiCard(`${k}-insights`, [16, 144, 1248, 82], [
    'Principal Driver de Satisfação', 'Maior Oportunidade de Melhoria', 'Serviço Melhor Avaliado', 'Serviço Pior Avaliado',
  ], { valueSize: 13, labelSize: 10, wrap: true });
  add(`${k}-ipa`, [16, ROW1, 620, H_FULL], {
    visualType: 'scatterChart',
    query: {
      queryState: {
        Category: { projections: [C(A, 'Serviço')] },
        X: { projections: [M('Correlação Nota x Satisfação')] },
        Y: { projections: [M('Nota Média')] },
      },
    },
    objects: {
      categoryLabels: [{ properties: { show: bool(true), fontSize: num(8) } }],
      dataPoint: [{ properties: { defaultColor: cor(COR.primaria) } }],
    },
    visualContainerObjects: vco({ title: 'Importância (correlação) x Desempenho (nota média)' }),
  });
  matrix(`${k}-tabela`, [644, ROW1, 620, 260], {
    rows: [C(A, 'Serviço')],
    values: ['Correlação Nota x Satisfação', 'Importância Relativa do Serviço', 'Gap de Nota Satisfeitos vs Insatisfeitos', 'Lift de Satisfação Top vs Bottom (p.p.)', 'Quadrante Importância x Desempenho'],
    title: 'Drivers por serviço',
  });
  add(`${k}-curva`, [644, ROW1 + 268, 620, H_FULL - 268], {
    visualType: 'lineChart',
    query: {
      queryState: {
        Category: { projections: [C(A, 'Nota')] },
        Series: { projections: [C(A, 'Categoria')] },
        Y: { projections: [M('% Satisfação dos Avaliadores')] },
      },
      sortDefinition: sortColumn(A, 'Nota'),
    },
    objects: { legend: [{ properties: { show: bool(true), position: str('Right') } }] },
    visualContainerObjects: vco({ title: '% Satisfação geral conforme a nota dada ao serviço' }),
  });
});

page('qualidade', 'Qualidade dos Dados', 'Completude e consistência da base', (k) => {
  kpiCard(`${k}-kpis`, [16, 144, 1248, 82], [
    'Total Passageiros', 'IDs Duplicados', 'Registros sem Atraso de Chegada', '% Registros sem Atraso de Chegada',
    'Passageiros com Alguma Resposta N/A', '% Completude das Avaliações',
  ]);
  barChart(`${k}-na`, [16, ROW1, 620, H_FULL], 'clusteredBarChart', {
    category: C(A, 'Serviço'), y: ['% Não Aplicável'], title: '% de respostas "não aplicável" por serviço',
    sort: sortMeasure('% Não Aplicável'), color: COR.destaque,
  });
  matrix(`${k}-tabela`, [644, ROW1, 620, H_FULL], {
    rows: [C(A, 'Serviço')],
    values: ['Avaliações Válidas', 'Respostas Não Aplicáveis', '% Não Aplicável'],
    title: 'Respostas por serviço',
  });
});

// ---------- escrita ----------
const reportJsonPath = path.join(REPORT, 'report.json');
const reportJson = JSON.parse(fs.readFileSync(reportJsonPath, 'utf8'));
let registered = reportJson.resourcePackages.find((p) => p.type === 'RegisteredResources');
if (!registered) {
  registered = { name: 'RegisteredResources', type: 'RegisteredResources', items: [] };
  reportJson.resourcePackages.push(registered);
}
if (!registered.items.some((i) => i.name === LOGO)) {
  registered.items.push({ name: LOGO, path: LOGO, type: 'Image' });
  fs.writeFileSync(reportJsonPath, JSON.stringify(reportJson, null, 2));
}

const pagesJsonPath = path.join(PAGES_DIR, 'pages.json');
const pagesJson = JSON.parse(fs.readFileSync(pagesJsonPath, 'utf8'));
const generated = new Set(PAGES.map((p) => p.name));

// Remove a página vazia padrão ("Página 1") e versões anteriores das páginas geradas.
for (const dir of fs.readdirSync(PAGES_DIR)) {
  const full = path.join(PAGES_DIR, dir);
  if (!fs.statSync(full).isDirectory()) continue;
  const pj = JSON.parse(fs.readFileSync(path.join(full, 'page.json'), 'utf8'));
  const empty = !fs.existsSync(path.join(full, 'visuals'));
  if (generated.has(dir) || (pj.displayName === 'Página 1' && empty)) fs.rmSync(full, { recursive: true });
}

for (const p of PAGES) {
  const dir = path.join(PAGES_DIR, p.name);
  fs.mkdirSync(path.join(dir, 'visuals'), { recursive: true });
  fs.writeFileSync(path.join(dir, 'page.json'), JSON.stringify({
    $schema: SCHEMA_PAGE,
    name: p.name,
    displayName: p.displayName,
    displayOption: 'FitToPage',
    height: 720,
    width: 1280,
    objects: { background: [{ properties: { color: cor('#F3F5F9'), transparency: num(0) } }] },
  }, null, 2));
  for (const v of p.visuals) {
    const vdir = path.join(dir, 'visuals', v.name);
    fs.mkdirSync(vdir, { recursive: true });
    fs.writeFileSync(path.join(vdir, 'visual.json'), JSON.stringify(v, null, 2));
  }
}

const remaining = fs.readdirSync(PAGES_DIR).filter((d) => fs.statSync(path.join(PAGES_DIR, d)).isDirectory());
pagesJson.pageOrder = [...PAGES.map((p) => p.name), ...remaining.filter((d) => !generated.has(d))];
pagesJson.activePageName = PAGES[0].name;
fs.writeFileSync(pagesJsonPath, JSON.stringify(pagesJson, null, 2));

console.log(PAGES.map((p) => `${p.displayName}: ${p.visuals.length} visuais (${p.name})`).join('\n'));
