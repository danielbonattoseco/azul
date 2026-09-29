// Gera o relatório one page bi_voos_onepage (PBIP + PBIR). Uso: node gerar_onepage_bi_voos.js
// Reaproveita o modelo semântico bi_voos.SemanticModel.
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const ROOT = path.join(__dirname, 'case claude');
const SRC_REPORT = path.join(ROOT, 'bi_voos.Report');
const REPORT = path.join(ROOT, 'bi_voos_onepage.Report');
const DEF = path.join(REPORT, 'definition');
const SCHEMA_PAGE = 'https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.1.0/schema.json';
const SCHEMA_VISUAL = 'https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.9.0/schema.json';
const LOGO = 'LOGO_AZUL_LINHAS_AEREAS17591500000000000.png';

const A = 'Avaliações';
const MT = '_Medidas';

const COR = {
  titulo: '#0B2D5B',
  texto: '#252423',
  suave: '#605E5C',
  negativo: '#D64545',
  neutro: '#A6A6A6',
  positivo: '#2E9E5B',
  excesso: '#8FB4E3',
  alerta: '#FDECEA',
  aviso: '#FFF4E5',
};

// ---------- helpers ----------
const hex = (n) => crypto.createHash('md5').update(`onepage-${n}`).digest('hex').slice(0, 20);
const lit = (v) => ({ expr: { Literal: { Value: v } } });
const str = (s) => lit(`'${s.replace(/'/g, "''")}'`);
const num = (n) => lit(`${n}D`);
const bool = (b) => lit(b ? 'true' : 'false');
const cor = (c) => ({ solid: { color: str(c) } });
const measureExpr = (name) => ({ Measure: { Expression: { SourceRef: { Entity: MT } }, Property: name } });
const corMedida = (name) => ({ solid: { color: { expr: measureExpr(name) } } });

const M = (name) => ({ field: measureExpr(name), queryRef: `${MT}.${name}`, nativeQueryRef: name });
const C = (table, col) => ({
  field: { Column: { Expression: { SourceRef: { Entity: table } }, Property: col } },
  queryRef: `${table}.${col}`,
  nativeQueryRef: col,
});
const sortMeasure = (name, direction) => ({ sort: [{ field: measureExpr(name), direction }], isDefaultSort: true });
const WILDCARD = { data: [{ dataViewWildcard: { matchingOption: 1 } }] };

const pad = (t, r = t, b = t, l = r) => [{ properties: { top: num(t), right: num(r), bottom: num(b), left: num(l) } }];

function vco({ title, bg = '#FFFFFF', border = '#E1E6EE' } = {}) {
  return {
    padding: pad(8),
    background: [{ properties: { show: bool(true), color: cor(bg), transparency: num(0) } }],
    border: [{ properties: { show: bool(true), color: cor(border), radius: num(6) } }],
    dropShadow: [{ properties: { show: bool(false) } }],
    title: title
      ? [{ properties: { show: bool(true), text: str(title), fontSize: num(11), bold: bool(true), fontColor: cor(COR.titulo) } }]
      : [{ properties: { show: bool(false) } }],
  };
}

const visuals = [];
let z = 0;
function add(key, [x, y, width, height], visual) {
  z += 1000;
  visuals.push({ $schema: SCHEMA_VISUAL, name: hex(key), position: { x, y, z, height, width, tabOrder: z }, visual });
}

// Texto rico: cada parágrafo é uma lista de runs [texto, estilo].
const run = (value, { size = 11, color = COR.texto, bold = false, font } = {}) => ({
  value,
  textStyle: {
    fontFamily: font || (bold ? 'Segoe UI Semibold' : 'Segoe UI'),
    fontSize: `${size}px`,
    color,
    ...(bold ? { fontWeight: 'bold' } : {}),
  },
});
const para = (runs, extra = {}) => ({ textRuns: runs, horizontalTextAlignment: 'left', ...extra });
const bullet = (runs) => para(runs, { listType: 'bullet', indent: 1 });
const numbered = (runs) => para(runs, { listType: 'number', indent: 1 });
const blank = () => para([run(' ', { size: 5 })]);

function textbox(key, pos, paragraphs, { bg = '#FFFFFF', border = '#E1E6EE', padding = pad(12, 16) } = {}) {
  add(key, pos, {
    visualType: 'textbox',
    objects: { general: [{ properties: { paragraphs } }] },
    visualContainerObjects: {
      background: [{ properties: { show: bool(true), color: cor(bg), transparency: num(0) } }],
      border: [{ properties: { show: bool(border !== null), ...(border ? { color: cor(border), radius: num(6) } : {}) } }],
      padding,
    },
  });
}

function card(key, pos, measure, { valueSize = 22, valueColor = COR.titulo, labelColor = COR.suave, bg = '#FFFFFF', border = '#E1E6EE', accent, wrap = false }) {
  const value = { fontSize: num(valueSize), fontColor: cor(valueColor), bold: bool(true) };
  if (wrap) value.textWrap = bool(true);
  const objects = {
    value: [{ properties: value, selector: { id: 'default' } }],
    label: [{ properties: { fontSize: num(10), fontColor: cor(labelColor) }, selector: { id: 'default' } }],
    padding: [{ properties: { paddingUniform: num(6) }, selector: { id: 'default' } }],
    layout: [{ properties: { paddingUniform: num(0) }, selector: { id: 'default' } }],
    accentBar: [{ properties: { show: bool(!!accent), ...(accent ? { color: cor(accent) } : {}) }, selector: { id: 'default' } }],
  };
  add(key, pos, {
    visualType: 'cardVisual',
    query: { queryState: { Data: { projections: [M(measure)] } } },
    objects,
    visualContainerObjects: {
      padding: pad(6),
      background: [{ properties: { show: bool(true), color: cor(bg), transparency: num(0) } }],
      border: [{ properties: { show: bool(true), color: cor(border), radius: num(8) } }],
    },
  });
}

function barByService(key, pos, measure, { title, direction }) {
  add(key, pos, {
    visualType: 'clusteredBarChart',
    query: {
      queryState: { Category: { projections: [C(A, 'Serviço')] }, Y: { projections: [M(measure)] } },
      sortDefinition: sortMeasure(measure, direction),
    },
    objects: {
      // Cor pela prioridade (medida Cor Quadrante).
      dataPoint: [{ properties: { fill: corMedida('Cor Quadrante') }, selector: WILDCARD }],
      labels: [{ properties: { show: bool(true), fontSize: num(9) } }],
      categoryAxis: [{ properties: { fontSize: num(9), showAxisTitle: bool(false) } }],
      valueAxis: [{ properties: { show: bool(false) } }],
    },
    visualContainerObjects: vco({ title }),
  });
}

// ---------- página ----------
const W = 1280;
const H = 960;

// Cabeçalho
add('header', [0, 0, W, 56], {
  visualType: 'textbox',
  objects: {
    general: [{
      properties: {
        paragraphs: [para([
          run('Serviços críticos da experiência do passageiro', { size: 20, color: '#FFFFFF', bold: true }),
          run('   |   one page para decisão: onde a avaliação é pior, quanto isso pesa e o que fazer', { size: 12, color: '#C9D6EA' }),
        ])],
      },
    }],
  },
  visualContainerObjects: {
    background: [{ properties: { show: bool(true), color: cor(COR.titulo), transparency: num(0) } }],
    border: [{ properties: { show: bool(false) } }],
    padding: pad(12, 16, 6, 184),
  },
});
add('logo', [8, 6, 164, 44], {
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
    padding: pad(6, 10),
    title: [{ properties: { show: bool(false) } }],
  },
  drillFilterOtherVisuals: true,
});

// Mensagem-chave (destaque)
textbox('mensagem', [16, 64, 1248, 60], [
  para([
    run('Mensagem-chave  ', { size: 13, color: COR.negativo, bold: true }),
    run('Wi-Fi a Bordo e a jornada digital (reserva e embarque online) concentram as piores notas e o maior potencial de ganho de satisfação. ', { size: 13, color: COR.texto, bold: true }),
    run('Priorizar esses serviços antes de itens com nota baixa, mas sem impacto na satisfação (localização do portão e conveniência de horários).', { size: 12, color: COR.texto }),
  ]),
], { bg: COR.aviso, border: '#F5C38B', padding: pad(10, 16) });

// KPIs: o mais importante (% Satisfação) em destaque escuro; alertas em vermelho.
const KY = 132, KH = 100;
card('kpi-satisfacao', [16, KY, 180, KH], '% Satisfação', { valueSize: 30, valueColor: '#FFFFFF', labelColor: '#C9D6EA', bg: COR.titulo, border: COR.titulo });
card('kpi-nota', [204, KY, 160, KH], 'Nota Média', { valueSize: 26, accent: COR.titulo });
card('kpi-bottom', [372, KY, 180, KH], '% Bottom-2-Box (Notas 1-2)', { valueSize: 26, valueColor: COR.negativo, accent: COR.negativo });
card('kpi-pior', [560, KY, 236, KH], 'Serviço Pior Avaliado', { valueSize: 15, valueColor: COR.negativo, accent: COR.negativo, wrap: true });
card('kpi-driver', [804, KY, 236, KH], 'Principal Driver de Satisfação', { valueSize: 15, accent: COR.titulo, wrap: true });
card('kpi-potencial', [1048, KY, 216, KH], 'Serviço com Maior Potencial de Ganho', { valueSize: 15, valueColor: COR.negativo, bg: COR.alerta, border: '#F2B8B5', accent: COR.negativo, wrap: true });

// Legenda das cores de prioridade (compartilhada pelos gráficos e pela tabela)
const sq = (c) => run('■ ', { size: 14, color: c, bold: true });
textbox('legenda', [16, 236, 808, 32], [
  para([
    run('Prioridade (Importância x Desempenho):   ', { size: 10, color: COR.suave, bold: true }),
    sq(COR.negativo), run('Prioridade de melhoria     ', { size: 10 }),
    sq(COR.positivo), run('Manter (ponto forte)     ', { size: 10 }),
    sq(COR.neutro), run('Baixa prioridade     ', { size: 10 }),
    sq(COR.excesso), run('Possível excesso de investimento', { size: 10 }),
  ]),
], { bg: '#F3F5F9', border: null, padding: pad(4, 8) });

barByService('bar-nota', [16, 272, 400, 368], 'Nota Média', { title: 'Nota média por serviço (1 a 5), piores no topo', direction: 'Ascending' });
barByService('bar-potencial', [424, 272, 400, 368], 'Potencial de Ganho de Satisfação (p.p.)', { title: 'Potencial de ganho de satisfação (p.p.)', direction: 'Descending' });

// Diagnóstico qualitativo (números da base completa)
const h = (t) => para([run(t, { size: 13, color: COR.titulo, bold: true })]);
const b = (t) => run(t, { size: 11, bold: true });
const t = (s) => run(s, { size: 11 });
const neg = (s) => run(s, { size: 11, color: COR.negativo, bold: true });
textbox('diagnostico', [832, 240, 432, 400], [
  h('Diagnóstico'),
  para([run('Serviços pior avaliados', { size: 11, color: COR.negativo, bold: true })]),
  bullet([b('Wi-Fi a Bordo: '), neg('nota 2,81'), t(' e 43% de notas 1-2. Com nota 5, 99% ficam satisfeitos; com notas 1-3, apenas 25% a 33%.')]),
  bullet([b('Reserva Online: '), neg('nota 2,88'), t(' e 42% de notas 1-2. Quem avalia bem tem 28 p.p. a mais de satisfação.')]),
  bullet([b('Portão (2,98) e Horários (3,22): '), t('notas baixas, mas sem relação com a satisfação (r ≈ 0). Não são prioridade.')]),
  blank(),
  para([run('Onde está o impacto', { size: 11, color: COR.titulo, bold: true })]),
  bullet([b('Embarque Online'), t(' é o principal driver (r = 0,57): satisfação de 12% a 14% com notas 1-3 contra '), b('87% com nota 5'), t('.')]),
  bullet([b('Entretenimento'), t(' (r = 0,40) e '), b('Wi-Fi'), t(' (r = 0,39) completam o trio de maior influência.')]),
  bullet([b('Viagens a lazer: '), neg('~10% de satisfação'), t(' (vs. 72% em negócios na classe Business), com nota de Wi-Fi ≈ 2,6.')]),
]);

// Tabela de apoio à decisão
add('tabela', [16, 648, 620, 296], {
  visualType: 'tableEx',
  query: {
    queryState: {
      Values: {
        projections: [
          C(A, 'Serviço'),
          M('Nota Média'),
          M('% Bottom-2-Box (Notas 1-2)'),
          M('Correlação Nota x Satisfação'),
          M('Potencial de Ganho de Satisfação (p.p.)'),
          M('Quadrante Importância x Desempenho'),
        ],
      },
    },
    sortDefinition: sortMeasure('Potencial de Ganho de Satisfação (p.p.)', 'Descending'),
  },
  objects: {
    columnHeaders: [{ properties: { columnAdjustment: str('growToFit'), autoSizeColumnWidth: bool(true), fontSize: num(9), bold: bool(true) } }],
    values: [
      { properties: { fontSize: num(9) } },
      {
        properties: { fontColor: corMedida('Cor Quadrante'), bold: bool(true) },
        selector: { ...WILDCARD, metadata: `${MT}.Quadrante Importância x Desempenho` },
      },
    ],
    total: [{ properties: { totals: bool(false) } }],
  },
  visualContainerObjects: vco({ title: 'Serviços ordenados por potencial de ganho' }),
});

// Recomendações
textbox('recomendacoes', [644, 648, 620, 296], [
  h('Possibilidades de melhoria (em ordem de prioridade)'),
  numbered([b('Conectividade: '), t('ampliar cobertura e velocidade do Wi-Fi e oferecer franquia gratuita de mensageria. '), neg('Maior potencial: ≈ +20 p.p.')]),
  numbered([b('Jornada digital: '), t('simplificar reserva e embarque online (menos etapas, check-in automático, cartão de embarque na carteira digital). '), neg('Embarque ≈ +17 p.p.; reserva ≈ +12 p.p.')]),
  numbered([b('Entretenimento: '), t('ampliar catálogo e streaming no dispositivo do passageiro, em sinergia com o Wi-Fi. '), neg('≈ +13 p.p.')]),
  numbered([b('Viajante a lazer na Economy: '), t('pacote de conforto (espaço para pernas, limpeza) e comunicação de valor para o segmento menos satisfeito.')]),
  numbered([b('Portão e horários: '), t('apenas monitorar; melhorar a nota não deve mover a satisfação.')]),
  blank(),
  para([run('Base: 129.880 passageiros. Potencial = % de notas 1-2 × lift de satisfação: indica associação, não causalidade, e não é somável entre serviços. '
    + 'Cartões, gráficos e tabela são dinâmicos; os textos refletem a base completa.', { size: 9, color: COR.suave })]),
]);

// ---------- escrita ----------
fs.rmSync(REPORT, { recursive: true, force: true });
fs.mkdirSync(DEF, { recursive: true });

// Tema e logo: copiados do relatório principal.
fs.cpSync(path.join(SRC_REPORT, 'StaticResources'), path.join(REPORT, 'StaticResources'), { recursive: true });
fs.copyFileSync(path.join(SRC_REPORT, 'definition', 'report.json'), path.join(DEF, 'report.json'));
fs.copyFileSync(path.join(SRC_REPORT, 'definition', 'version.json'), path.join(DEF, 'version.json'));

const write = (file, obj) => fs.writeFileSync(file, JSON.stringify(obj, null, 2));
write(path.join(REPORT, 'definition.pbir'), {
  $schema: 'https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json',
  version: '4.0',
  datasetReference: { byPath: { path: '../bi_voos.SemanticModel' } },
});
write(path.join(ROOT, 'bi_voos_onepage.pbip'), {
  $schema: 'https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json',
  version: '1.0',
  artifacts: [{ report: { path: 'bi_voos_onepage.Report' } }],
  settings: { enableAutoRecovery: true },
});

write(path.join(REPORT, '.platform'), {
  $schema: 'https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json',
  metadata: { type: 'Report', displayName: 'bi_voos_onepage' },
  config: { version: '2.0', logicalId: '3f6b2c1e-8d4a-4b7e-9a51-6c0e2d9f4a17' },
});

const pageName = hex('page');
const pageDir = path.join(DEF, 'pages', pageName);
fs.mkdirSync(path.join(pageDir, 'visuals'), { recursive: true });
write(path.join(DEF, 'pages', 'pages.json'), {
  $schema: 'https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.1.0/schema.json',
  pageOrder: [pageName],
  activePageName: pageName,
});
write(path.join(pageDir, 'page.json'), {
  $schema: SCHEMA_PAGE,
  name: pageName,
  displayName: 'One Page - Serviços Críticos',
  displayOption: 'FitToWidth',
  height: H,
  width: W,
  objects: { background: [{ properties: { color: cor('#F3F5F9'), transparency: num(0) } }] },
});
for (const v of visuals) {
  const dir = path.join(pageDir, 'visuals', v.name);
  fs.mkdirSync(dir, { recursive: true });
  write(path.join(dir, 'visual.json'), v);
}
console.log(`bi_voos_onepage: ${visuals.length} visuais (${pageName})`);
