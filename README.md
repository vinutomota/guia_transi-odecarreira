const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  LevelFormat, PageBreak, BorderStyle, ShadingType, TableRow, TableCell,
  Table, WidthType
} = require('docx');
const fs = require('fs');

// ─── COLORS ───────────────────────────────────────────────────────────────────
const BLUE   = "1A56A0";
const TEAL   = "0D8A7C";
const ACCENT = "E87722";
const LIGHT_BLUE_BG = "EBF4FB";
const LIGHT_TEAL_BG = "E6F5F3";
const LIGHT_ORANGE_BG = "FEF3E9";
const GRAY_BG  = "F5F6F8";
const DARK_TEXT = "1A1A2E";
const MID_TEXT  = "4A4A6A";

// ─── HELPERS ──────────────────────────────────────────────────────────────────
function pgBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

function spacer(pts = 120) {
  return new Paragraph({ spacing: { before: pts, after: 0 }, children: [] });
}

function h1(text, color = BLUE) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 180 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: color, space: 6 } },
    children: [new TextRun({ text, color, bold: true, font: "Arial", size: 36 })]
  });
}

function h2(text, color = TEAL) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 120 },
    children: [new TextRun({ text, color, bold: true, font: "Arial", size: 28 })]
  });
}

function h3(text, color = ACCENT) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 200, after: 80 },
    children: [new TextRun({ text, color, bold: true, font: "Arial", size: 24 })]
  });
}

function body(runs, centered = false) {
  const children = Array.isArray(runs) ? runs : [new TextRun({ text: runs, font: "Arial", size: 22, color: DARK_TEXT })];
  return new Paragraph({
    alignment: centered ? AlignmentType.CENTER : AlignmentType.JUSTIFIED,
    spacing: { before: 80, after: 80, line: 360 },
    children
  });
}

function bold(text, color = DARK_TEXT) {
  return new TextRun({ text, bold: true, font: "Arial", size: 22, color });
}

function run(text, color = DARK_TEXT) {
  return new TextRun({ text, font: "Arial", size: 22, color });
}

function italicRun(text, color = MID_TEXT) {
  return new TextRun({ text, italics: true, font: "Arial", size: 22, color });
}

function bullet(runs, reference = "bullets") {
  const children = Array.isArray(runs) ? runs : [new TextRun({ text: runs, font: "Arial", size: 22, color: DARK_TEXT })];
  return new Paragraph({
    numbering: { reference, level: 0 },
    spacing: { before: 60, after: 60, line: 320 },
    children
  });
}

function numberedBullet(runs) {
  return bullet(runs, "numbers");
}

function callout(runs, bg = LIGHT_BLUE_BG, borderColor = BLUE) {
  const children = Array.isArray(runs) ? runs : [new TextRun({ text: runs, font: "Arial", size: 22, color: DARK_TEXT })];
  const border = { style: BorderStyle.SINGLE, size: 8, color: borderColor };
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 140, after: 140, line: 340 },
    indent: { left: 600, right: 400 },
    border: { left: border },
    shading: { fill: bg, type: ShadingType.CLEAR },
    children
  });
}

function sectionDivider(color = BLUE) {
  return new Paragraph({
    spacing: { before: 200, after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 2, color, space: 1 } },
    children: []
  });
}

// ─── COVER PAGE ───────────────────────────────────────────────────────────────
function coverPage() {
  return [
    spacer(800),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 40 },
      children: [new TextRun({ text: "DO ZERO AOS DADOS", font: "Arial", size: 64, bold: true, color: BLUE })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 60 },
      children: [new TextRun({ text: "O Mapa para Profissionais que Desejam se Reinventar", font: "Arial", size: 30, color: TEAL })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 60 },
      children: [new TextRun({ text: "na Era da Inteligência Artificial", font: "Arial", size: 30, color: TEAL })]
    }),
    spacer(80),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 40 },
      border: {
        top: { style: BorderStyle.SINGLE, size: 2, color: ACCENT, space: 4 },
        bottom: { style: BorderStyle.SINGLE, size: 2, color: ACCENT, space: 4 }
      },
      children: [new TextRun({ text: "  Guia de Transição de Carreira para a Área de Dados  ", font: "Arial", size: 22, color: ACCENT, italics: true })]
    }),
    spacer(600),
    pgBreak()
  ];
}

// ─── OPTIONS PAGE ─────────────────────────────────────────────────────────────
function titlesPage() {
  return [
    h1("Opções de Título para Este E-book"),
    spacer(),
    h3("Opção 1"),
    body([bold("A Nova Fronteira: "), run("O Guia Definitivo para Sua Transição de Carreira para a Área de Dados")]),
    spacer(60),
    h3("Opção 2"),
    body([bold("Decodificando o Futuro: "), run("Como Migrar de Setores Tradicionais para a Tecnologia dos Dados")]),
    spacer(60),
    h3("Opção 3 (Sugerido)"),
    body([bold("Do Zero aos Dados: "), run("O Mapa para Profissionais que Desejam se Reinventar na Era da Inteligência Artificial")]),
    spacer(60),
    callout([
      bold("Por que este título funciona? "),
      run("Ele é direto, inclusivo e transmite movimento — a sensação de uma jornada real, não apenas de conteúdo técnico. O termo 'zero' elimina a barreira de entrada e convida quem ainda não tem experiência em tecnologia a se sentir acolhido.")
    ], LIGHT_TEAL_BG, TEAL),
    pgBreak()
  ];
}

// ─── INTRO ────────────────────────────────────────────────────────────────────
function introChapter() {
  return [
    h1("Introdução: O Despertar para os Dados"),
    spacer(),
    body([
      run("Vivemos em uma era onde a informação é frequentemente descrita como o "),
      bold('"novo petróleo"'),
      run(". Mas, ao contrário do petróleo bruto, o valor real dos dados não está na sua posse — está na capacidade humana de transformá-los em insights estratégicos, decisões inteligentes e produtos que mudam vidas.")
    ]),
    spacer(),
    body("Pense no volume de dados que uma grande rede de supermercados gera em um único dia: quais produtos foram vendidos, em quais horários, para quais perfis de consumidores, sob quais condições climáticas. Esse volume imenso de informação é inútil sem alguém que saiba lê-lo, interpretá-lo e, mais importante, transformá-lo em ação. É exatamente esse alguém que o mercado está buscando — e pagando muito bem para encontrar."),
    spacer(),
    body("Se você está lendo este livro, provavelmente sente que algo está mudando ao seu redor. Talvez você trabalhe em saúde, educação, varejo, logística ou finanças, e perceba que cada vez mais decisões são tomadas com base em dados. Talvez tenha ouvido falar em salários atrativos na área de tecnologia e se perguntado se esse caminho também é possível para você. A resposta curta é: sim, é possível — e mais do que isso, é provável que você já tenha mais do que precisa para começar."),
    spacer(),
    callout([
      bold("A grande verdade que ninguém te conta: "),
      run("54,2% dos profissionais que atuam com dados no Brasil hoje não tinham experiência prévia em TI antes de migrar. O setor de dados é, estruturalmente, um setor de profissionais em transição.")
    ], LIGHT_BLUE_BG, BLUE),
    spacer(),
    body("Este livro foi escrito para você que vem de outro campo, outro setor, outra história. Não é um manual técnico cheio de jargões que pressupõe que você já sabe tudo. É um mapa — desenhado por quem conhece o território — para te ajudar a entender onde você está, para onde pode ir e qual caminho faz mais sentido para o seu perfil."),
    spacer(),
    body([
      run("Ao longo destas páginas, você vai conhecer as três principais carreiras da área de dados, entender o mercado real (com números, não com promessas vazias), descobrir como aproveitar o conhecimento que você já tem — o seu "),
      bold("conhecimento de domínio"),
      run(" — e receber um guia prático para dar os primeiros passos com inteligência e sem desperdiçar tempo em cursos que não levam a lugar nenhum.")
    ]),
    spacer(),
    body("Prepare-se. O futuro pertence a quem sabe fazer as perguntas certas para os dados certos. E esse futuro começa agora."),
    pgBreak()
  ];
}

// ─── CHAPTER 1 ────────────────────────────────────────────────────────────────
function chapter1() {
  return [
    h1("Capítulo 1: O Mapa do Mercado"),
    new Paragraph({
      spacing: { before: 0, after: 180 },
      children: [new TextRun({ text: "Quem está a migrar — e por quê agora?", font: "Arial", size: 26, italics: true, color: MID_TEXT })]
    }),
    body("Antes de qualquer passo prático, é fundamental que você entenda o terreno em que está pisando. A área de dados não é uma moda passageira do mercado de trabalho — ela é a resposta estrutural de uma economia que se tornou digitalmente dependente e que, agora, não funciona mais sem profissionais capazes de transformar informação em inteligência."),
    spacer(),
    h2("O Cenário Brasileiro: Dados que Convencem"),
    body([
      run("O Brasil vive um momento singular. A digitalização acelerada pós-pandemia, combinada com a expansão do mercado de startups e a transformação digital das grandes corporações, criou uma demanda por profissionais de dados que "),
      bold("supera em muito a oferta disponível."),
      run(" Isso significa, em termos práticos, que quem se qualifica hoje encontra oportunidades rápidas — mesmo sem anos de experiência formal na área.")
    ]),
    spacer(),
    callout([
      run("Segundo dados do setor, "),
      bold("mais da metade dos profissionais de dados ativos no Brasil"),
      run(" vêm de outras áreas de formação. Isso não é exceção — é a regra. O mercado já normalizou e, mais do que isso, valoriza a diversidade de bagagem que esses profissionais trazem.")
    ], LIGHT_TEAL_BG, TEAL),
    spacer(),
    body("Áreas como Varejo, Educação e Saúde estão contratando massivamente para modernizar suas operações. Uma rede de hospitais precisa de analistas que entendam tanto de SQL quanto de protocolos médicos. Uma rede de escolas precisa de cientistas de dados que compreendam pedagogia. Um marketplace de moda precisa de pessoas que falem a língua tanto do negócio quanto dos números — e é exatamente aí que profissionais em transição têm vantagem competitiva."),
    spacer(),
    h2("Quem está fazendo essa transição?"),
    body("O perfil de quem migra para dados é muito mais diverso do que a maioria das pessoas imagina. Não é só o jovem de 22 anos recém-formado em Ciências da Computação. Vejamos os grupos mais comuns:"),
    spacer(),
    h3("Profissionais de Negócios (Administração, Economia, Marketing)"),
    body("São os que mais rapidamente se adaptam à função de Analista de Dados. Já conhecem o vocabulário do negócio, sabem o que uma diretoria precisa ver em um dashboard e entendem as métricas que movem uma empresa. O que precisam aprender é o lado técnico: SQL, Excel avançado e uma ferramenta de BI como Power BI ou Tableau."),
    spacer(),
    h3("Profissionais de Saúde e Ciências Exatas"),
    body("Médicos, enfermeiros, biólogos, químicos e farmacêuticos possuem uma base analítica sólida e estão habituados a trabalhar com dados clínicos e experimentais. Muitos encontram na Ciência de Dados uma extensão natural de sua carreira, especialmente em empresas de healthtech, laboratórios e gestoras de planos de saúde."),
    spacer(),
    h3("Professores e Profissionais de Educação"),
    body("A capacidade de explicar conceitos complexos de forma acessível é um diferencial enorme em dados. Além disso, o setor de EdTech está em plena expansão, e profissionais que dominam tanto pedagogia quanto análise de dados têm um perfil raríssimo e muito disputado."),
    spacer(),
    h3("Profissionais de Logística, Operações e Engenharia"),
    body([
      run("Engenheiros e gestores de operações já pensam em sistemas, fluxos e otimização — o que é essencialmente o que um Engenheiro de Dados faz. A transição para a área de dados costuma ser "),
      bold("a mais natural e rápida"),
      run(" para esse perfil, especialmente para funções de engenharia de dados e automação de pipelines.")
    ]),
    spacer(),
    h2("A Janela de Oportunidade"),
    body("Existe um fenômeno que os economistas chamam de 'janela de oportunidade estrutural' — um período em que uma transformação de mercado cria vagas em ritmo mais acelerado do que a formação tradicional consegue suprir. A área de dados está nessa janela agora."),
    spacer(),
    body("Universidades levam quatro a cinco anos para formar um profissional. Bootcamps e cursos online já reduziram esse ciclo para seis a doze meses. Quem entrar agora ainda pega a crista da onda — com salários mais altos, menos concorrência e mais oportunidades de crescimento rápido do que haverá daqui a cinco anos, quando o mercado estiver mais saturado."),
    spacer(),
    callout([
      bold("Conclusão do Capítulo: "),
      run("O momento é agora. O mercado brasileiro de dados é jovem, aquecido e carente de profissionais com conhecimento de negócio — que é exatamente o que você tem. O próximo passo é entender para qual carreira dentro de dados o seu perfil se encaixa melhor.")
    ], LIGHT_ORANGE_BG, ACCENT),
    pgBreak()
  ];
}

// ─── CHAPTER 2 ────────────────────────────────────────────────────────────────
function chapter2() {
  return [
    h1("Capítulo 2: A Bússola das Carreiras"),
    new Paragraph({
      spacing: { before: 0, after: 180 },
      children: [new TextRun({ text: "Analista, Engenheiro e Cientista de Dados — quem faz o quê?", font: "Arial", size: 26, italics: true, color: MID_TEXT })]
    }),
    body("Um dos maiores pontos de confusão para quem começa a pesquisar a área de dados é a multiplicidade de títulos e funções. 'Analista de Dados', 'Engenheiro de Dados', 'Cientista de Dados', 'Engenheiro de Machine Learning', 'Analista de BI'... a lista parece interminável e as fronteiras entre elas, nebulosas."),
    spacer(),
    body("Mas existe uma lógica clara por trás dessas nomenclaturas. Pense na área de dados como uma cadeia produtiva: alguém precisa construir a fábrica (infraestrutura), alguém precisa operar as máquinas (análise), e alguém precisa criar produtos inovadores (ciência). Cada uma dessas funções tem um perfil, uma rotina e um conjunto de habilidades distintos."),
    spacer(),
    h2("1. Analista de Dados — O Sommelier Estratégico"),
    body([
      run("Se os dados de uma empresa fossem vinhos, o Analista de Dados seria o sommelier: alguém que conhece profundamente o produto, sabe onde ele foi produzido, quais são suas características e — o mais importante — "),
      bold("sabe recomendar o vinho certo para cada ocasião."),
    ]),
    spacer(),
    h3("O que faz no dia a dia?"),
    bullet("Extrai dados de bancos de dados usando SQL para responder perguntas específicas do negócio."),
    bullet("Constrói dashboards e relatórios visuais em ferramentas como Power BI, Tableau ou Looker."),
    bullet("Analisa tendências de vendas, comportamento de clientes, performance de campanhas ou eficiência operacional."),
    bullet("Apresenta os resultados para gestores e diretores em reuniões de negócio — traduzindo números em decisões."),
    bullet("Trabalha em ciclos rápidos: uma semana pode estar analisando o por quê de as vendas terem caído, na outra, mapeando oportunidades de expansão."),
    spacer(),
    h3("Com quem trabalha?"),
    body("O Analista de Dados vive na interface entre tecnologia e negócio. Ele se reúne tanto com o time de marketing quanto com o time de TI. Precisa falar a língua do negócio para entender o problema e a linguagem técnica para resolvê-lo."),
    spacer(),
    h3("Habilidades essenciais"),
    bullet([bold("SQL: "), run("A ferramenta mais importante. É a linguagem universal dos dados.")]),
    bullet([bold("Excel/Google Sheets avançado: "), run("Ainda muito utilizado em médias e grandes empresas.")]),
    bullet([bold("Ferramenta de BI: "), run("Power BI (mais comum no Brasil), Tableau ou Looker.")]),
    bullet([bold("Storytelling com dados: "), run("Saber apresentar insights de forma clara e persuasiva.")]),
    bullet([bold("Pensamento analítico: "), run("A habilidade de decompor problemas complexos em partes menores.")]),
    spacer(),
    callout([
      bold("Salário médio no Brasil: "),
      run("R$ 4.000 a R$ 9.000 (júnior a pleno). Sênior e liderança podem ultrapassar R$ 15.000.")
    ], LIGHT_BLUE_BG, BLUE),
    spacer(),
    h2("2. Engenheiro de Dados — O Arquiteto dos Alicerces"),
    body("Se o Analista de Dados é o sommelier, o Engenheiro de Dados é quem construiu a vinícola: as instalações, os barris, os canos de transporte, os sistemas de controle de temperatura. Sem ele, não há vinho para analisar."),
    spacer(),
    body("O Engenheiro de Dados é responsável pela infraestrutura que sustenta todo o trabalho analítico de uma empresa. Enquanto o Analista pergunta 'o que os dados dizem?', o Engenheiro garante que os dados existam, sejam confiáveis e cheguem ao lugar certo no momento certo."),
    spacer(),
    h3("O que faz no dia a dia?"),
    bullet("Constrói e mantém pipelines de dados — os 'canos' que movem informações de um sistema para outro de forma automatizada."),
    bullet("Desenvolve processos de ETL (Extração, Transformação e Carregamento de dados)."),
    bullet("Garante a qualidade, consistência e disponibilidade dos dados para os outros times."),
    bullet("Trabalha com bancos de dados, sistemas de armazenamento em nuvem (AWS, Google Cloud, Azure) e ferramentas como Apache Spark, Kafka e Airflow."),
    bullet("Monitora e otimiza a performance dos sistemas de dados."),
    spacer(),
    h3("Habilidades essenciais"),
    bullet([bold("SQL avançado: "), run("Indispensável, com conhecimento profundo de otimização de queries.")]),
    bullet([bold("Python ou Scala: "), run("Para automatização e processamento de dados em larga escala.")]),
    bullet([bold("Plataformas de nuvem: "), run("AWS, GCP ou Azure — pelo menos uma delas em nível intermediário.")]),
    bullet([bold("Ferramentas de pipeline: "), run("Apache Airflow, dbt, Spark.")]),
    bullet([bold("Pensamento sistêmico: "), run("Capacidade de enxergar como os dados fluem por toda a organização.")]),
    spacer(),
    callout([
      bold("Salário médio no Brasil: "),
      run("R$ 7.000 a R$ 14.000 (júnior a pleno). Sênior e especialistas em nuvem podem ultrapassar R$ 20.000.")
    ], LIGHT_TEAL_BG, TEAL),
    spacer(),
    h2("3. Cientista de Dados — O Alquimista do Futuro"),
    body([
      run("Se o Analista olha para o passado (o que aconteceu?) e o Engenheiro garante o presente (os dados estão fluindo?), o Cientista de Dados olha para o "),
      bold("futuro"),
      run(": o que vai acontecer? O que podemos criar com esses dados que nunca foi feito antes?")
    ]),
    spacer(),
    body("O Cientista de Dados combina estatística avançada, programação e conhecimento de negócio para criar modelos preditivos, algoritmos de recomendação e sistemas de Machine Learning. É a função mais técnica das três — e também a que exige uma base matemática mais sólida."),
    spacer(),
    h3("O que faz no dia a dia?"),
    bullet("Formula hipóteses sobre comportamentos e padrões nos dados."),
    bullet("Treina e avalia modelos de Machine Learning para prever churn de clientes, recomendar produtos, detectar fraudes ou otimizar preços."),
    bullet("Realiza análises exploratórias profundas para descobrir padrões que não são óbvios."),
    bullet("Colabora com times de engenharia para colocar modelos em produção."),
    bullet("Comunica resultados técnicos para stakeholders não-técnicos."),
    spacer(),
    h3("Habilidades essenciais"),
    bullet([bold("Python (avançado): "), run("Com domínio de bibliotecas como Pandas, NumPy, Scikit-learn e TensorFlow.")]),
    bullet([bold("Estatística e Probabilidade: "), run("Fundamentos sólidos são indispensáveis.")]),
    bullet([bold("SQL: "), run("Para extração e manipulação de dados.")]),
    bullet([bold("Machine Learning: "), run("Algoritmos supervisionados, não-supervisionados e técnicas de validação.")]),
    bullet([bold("Curiosidade científica: "), run("A mentalidade de quem formula hipóteses e as testa com rigor.")]),
    spacer(),
    callout([
      bold("Salário médio no Brasil: "),
      run("R$ 8.000 a R$ 16.000 (júnior a pleno). Especialistas em IA e Deep Learning podem ultrapassar R$ 25.000.")
    ], LIGHT_ORANGE_BG, ACCENT),
    spacer(),
    h2("Qual caminho é para mim?"),
    body("A resposta depende do seu perfil atual, da sua tolerância a aspectos técnicos e do tipo de problema que mais te instiga resolver. Use as perguntas abaixo como um guia de autorreflexão:"),
    spacer(),
    bullet([bold("Você gosta de comunicar resultados e entender o negócio? "), run("→ Analista de Dados")]),
    bullet([bold("Você tem perfil de engenharia, pensa em sistemas e adora automatizar processos? "), run("→ Engenheiro de Dados")]),
    bullet([bold("Você tem base em exatas, adora matemática e quer criar coisas novas com algoritmos? "), run("→ Cientista de Dados")]),
    spacer(),
    callout([
      bold("Dica de ouro: "),
      run("Para quem está começando, o caminho mais acessível e com retorno mais rápido costuma ser o de Analista de Dados. A curva de aprendizado é mais suave, a demanda é enorme e a função serve como uma excelente porta de entrada para as demais.")
    ], LIGHT_BLUE_BG, BLUE),
    pgBreak()
  ];
}

// ─── CHAPTER 3 ────────────────────────────────────────────────────────────────
function chapter3() {
  return [
    h1("Capítulo 3: O Guia de Sobrevivência do Iniciante"),
    new Paragraph({
      spacing: { before: 0, after: 180 },
      children: [new TextRun({ text: "Desafios reais e como superá-los com inteligência", font: "Arial", size: 26, italics: true, color: MID_TEXT })]
    }),
    body("Até aqui, você conheceu o mercado e as carreiras. Agora vamos falar sobre o que ninguém te conta nos anúncios de cursos: os desafios reais de quem está começando — e, mais importante, como superá-los."),
    spacer(),
    body("Transições de carreira são empolgantes, mas não são lineares. Há momentos de euforia e momentos de dúvida. Entender os obstáculos com antecedência é a diferença entre desistir e adaptar."),
    spacer(),
    h2("Os 5 Desafios Mais Comuns — e Como Enfrentá-los"),
    spacer(),
    h3("Desafio 1: A Paralisia do Excesso de Opções"),
    body("O mercado de cursos de dados explodiu. São dezenas de plataformas, centenas de cursos, milhares de tutoriais no YouTube. O resultado? Muita gente passa meses assistindo aulas sem nunca colocar a mão na massa — e sem construir nada concreto para mostrar."),
    spacer(),
    body([bold("Como superar: "), run("Escolha UM trilha de aprendizado e siga até o final antes de explorar alternativas. Para Análise de Dados, uma trilha eficiente começa com SQL básico → Excel/Power BI → Python básico → Projeto prático. Simples, sequencial e com entregáveis claros em cada etapa.")]),
    spacer(),
    h3("Desafio 2: Focar Demais em Ferramentas, de Menos em Problemas"),
    body("'Aprendi Python, aprendi R, aprendi Spark, agora estou aprendendo Julia...' É um erro clássico. As ferramentas mudam — a lógica analítica, não. Empresas não contratam pessoas que 'sabem Python'. Contratam pessoas que 'resolvem problemas com dados'."),
    spacer(),
    body([bold("Como superar: "), run("Para cada nova ferramenta que aprender, aplique-a imediatamente em um projeto real — de preferência de uma área que você já conhece. Um nutricionista que analisa dados de consumo alimentar com Python é muito mais contratável do que alguém que sabe Python mas não tem contexto nenhum.")]),
    spacer(),
    h3("Desafio 3: A Síndrome do Impostor"),
    body("'Eu não sou formado em computação. Não sou desenvolvedor. Quem vai me contratar?' Este é, provavelmente, o maior inimigo dos profissionais em transição. A síndrome do impostor faz você se comparar com quem já está no mercado — esquecendo que todos eles também tiveram um primeiro dia."),
    spacer(),
    body([bold("Como superar: "), run("Documente publicamente a sua jornada. Um perfil ativo no LinkedIn, um repositório no GitHub ou um blog onde você compartilha o que está aprendendo constroem credibilidade de forma contínua. Você não precisa saber tudo para começar a ser visível — precisa mostrar que está progredindo.")]),
    spacer(),
    h3("Desafio 4: Não Ter Portfólio"),
    body("Recrutadores de dados não fazem perguntas do tipo 'em que empresa você trabalhou?'. Eles perguntam: 'me mostra um projeto seu'. Sem portfólio, você é invisível — mesmo com muito conhecimento."),
    spacer(),
    body([bold("Como superar: "), run("Construa pelo menos 3 projetos antes de começar a se candidatar a vagas. Fontes de dados públicos são perfeitas para isso: dados do IBGE, do DATASUS, do INEP, do Kaggle. Escolha temas que você conhece bem — um professor que analisa dados de evasão escolar conta uma história muito mais poderosa do que alguém que replica um projeto genérico de dataset de vinhos.")]),
    spacer(),
    h3("Desafio 5: Ignorar a Inteligência Artificial como Aliada"),
    body("Em 2024 e 2025, o mercado mudou de forma irreversível: a IA generativa tornou-se uma ferramenta de produtividade obrigatória para profissionais de dados. Quem ainda não usa IA para acelerar análises, gerar código e interpretar resultados está trabalhando com uma mão amarrada nas costas."),
    spacer(),
    body([bold("Como superar: "), run("Incorpore o uso de ferramentas como Claude, ChatGPT e GitHub Copilot no seu fluxo de trabalho desde o início. Use IA para tirar dúvidas de código, gerar explicações de conceitos estatísticos, criar templates de relatórios e depurar erros. A competência não é 'saber tudo' — é saber usar os melhores recursos disponíveis.")]),
    spacer(),
    h2("Os 7 Mandamentos do Profissional de Dados Iniciante"),
    spacer(),
    numberedBullet([bold("Construa antes de aprender mais: "), run("Cada novo conceito deve gerar um entregável.")]),
    numberedBullet([bold("Use seu conhecimento de domínio: "), run("Sua experiência anterior é vantagem competitiva, não bagagem desnecessária.")]),
    numberedBullet([bold("Seja público: "), run("Compartilhe sua jornada — erros incluídos. Transparência gera confiança.")]),
    numberedBullet([bold("Faça perguntas antes de aprender a linguagem: "), run("'Que problema de negócio eu quero resolver?' vem antes de 'que linguagem devo aprender?'")]),
    numberedBullet([bold("Construa uma rede: "), run("A comunidade de dados no Brasil (meetups, Discord, LinkedIn) é generosa e acessível.")]),
    numberedBullet([bold("Domine o básico muito bem: "), run("SQL bem escrito, análise exploratória cuidadosa e visualização clara valem mais do que 10 algoritmos de ML mal aplicados.")]),
    numberedBullet([bold("Seja paciente com a curva: "), run("A maioria das transições bem-sucedidas leva de 8 a 18 meses. Isso é normal, não falha.")]),
    pgBreak()
  ];
}

// ─── CONCLUSION ───────────────────────────────────────────────────────────────
function conclusion() {
  return [
    h1("Conclusão: O Primeiro Passo é Sempre o Mais Difícil"),
    spacer(),
    body("Você chegou até aqui. Isso já diz muito sobre você."),
    spacer(),
    body("A maioria das pessoas que pensa em mudar de carreira pára no primeiro obstáculo — seja o medo da mudança, a dúvida sobre a própria capacidade ou simplesmente a sobrecarga de informações. Você leu este livro até o fim, o que significa que está disposto a ir além do pensamento e em direção à ação."),
    spacer(),
    body("A área de dados não é perfeita. Tem suas complexidades, suas curvas de aprendizado, seus momentos de frustração. Mas é, hoje, um dos poucos setores que combina três coisas raramente encontradas juntas: crescimento sustentado, remuneração atrativa e impacto real no mundo."),
    spacer(),
    body("Cada vez que um algoritmo ajuda a detectar um câncer mais cedo, a distribuir vacinas de forma mais eficiente ou a personalizar a experiência de um estudante em dificuldade — há um profissional de dados por trás. Você pode ser esse profissional."),
    spacer(),
    callout([
      bold("O mapa que você recebeu neste livro: "),
      run("1) O mercado existe e está aquecido. 2) Seu background é um ativo, não um obstáculo. 3) Há três caminhos possíveis — escolha o que combina com você. 4) Os desafios são reais, mas superáveis com método. 5) O primeiro projeto é mais importante do que o décimo curso.")
    ], LIGHT_TEAL_BG, TEAL),
    spacer(),
    h2("Seus Próximos Passos Concretos"),
    spacer(),
    numberedBullet([bold("Esta semana: "), run("Escolha uma trilha (Analista, Engenheiro ou Cientista) e inscreva-se em um curso de SQL gratuito.")]),
    numberedBullet([bold("Em 30 dias: "), run("Complete seu primeiro projeto com dados públicos de uma área que você conhece bem.")]),
    numberedBullet([bold("Em 60 dias: "), run("Publique seu projeto no GitHub e escreva um post no LinkedIn sobre o que aprendeu.")]),
    numberedBullet([bold("Em 90 dias: "), run("Candidate-se às suas primeiras vagas — mesmo que ainda não se sinta 'pronto'. A preparação final acontece no processo.")]),
    numberedBullet([bold("Em 12 meses: "), run("Com consistência, você terá um portfólio, uma rede e as habilidades necessárias para sua primeira posição formal na área.")]),
    spacer(),
    body("A área de dados é o epicentro da revolução tecnológica atual. E a melhor notícia é: ela ainda tem espaço — muito espaço — para quem chega com vontade, método e a coragem de recomeçar."),
    spacer(),
    body([italicRun("O mapa está nas suas mãos. Agora é hora de começar a caminhar.")]),
    pgBreak()
  ];
}

// ─── APPENDIX ─────────────────────────────────────────────────────────────────
function appendix() {
  return [
    h1("Apêndice: Ferramentas de Bolso"),
    spacer(),
    h2("Glossário do Iniciante — Os Termos que Você Precisa Conhecer"),
    spacer(),
    bullet([bold("BI (Business Intelligence): "), run("O processo de transformar dados brutos em informações visuais e acionáveis para a tomada de decisão empresarial. Ferramentas populares: Power BI, Tableau, Looker.")]),
    spacer(40),
    bullet([bold("SQL (Structured Query Language): "), run("A linguagem universal para buscar, filtrar, agrupar e manipular informações em bancos de dados relacionais. É a habilidade número 1 mais solicitada em vagas de dados no Brasil.")]),
    spacer(40),
    bullet([bold("ETL (Extract, Transform, Load): "), run("O processo de extrair dados de uma fonte, transformá-los (limpar, padronizar, enriquecer) e carregá-los em um destino — geralmente um data warehouse ou data lake.")]),
    spacer(40),
    bullet([bold("Data Warehouse: "), run("Um repositório centralizado e estruturado de dados históricos de uma empresa, otimizado para análises. Exemplos: BigQuery, Snowflake, Redshift.")]),
    spacer(40),
    bullet([bold("Data Lake: "), run("Um repositório de dados brutos em seu formato original, estruturados ou não. Mais flexível que um data warehouse, permite análises exploratórias sem estrutura predefinida.")]),
    spacer(40),
    bullet([bold("Machine Learning: "), run("Área da inteligência artificial que desenvolve algoritmos capazes de aprender padrões a partir de dados históricos e fazer previsões ou classificações sobre dados novos.")]),
    spacer(40),
    bullet([bold("Pipeline de Dados: "), run("Uma sequência automatizada de processos que movem dados de um sistema a outro — desde a coleta até a disponibilização para análise.")]),
    spacer(40),
    bullet([bold("Dashboard: "), run("Um painel visual interativo que consolida as principais métricas de um negócio em um único lugar, facilitando o monitoramento e a tomada de decisão.")]),
    spacer(40),
    bullet([bold("KPI (Key Performance Indicator): "), run("Indicadores-chave de desempenho — as métricas mais importantes para avaliar se uma empresa ou equipe está alcançando seus objetivos.")]),
    spacer(40),
    bullet([bold("LGPD (Lei Geral de Proteção de Dados): "), run("A lei brasileira que regulamenta como empresas e profissionais devem coletar, armazenar, tratar e compartilhar dados pessoais. Conhecê-la é obrigatório para qualquer profissional de dados.")]),
    spacer(40),
    bullet([bold("Pandas: "), run("Biblioteca Python amplamente usada para manipulação e análise de dados tabulares. É a ferramenta mais utilizada por analistas e cientistas de dados no dia a dia.")]),
    spacer(40),
    bullet([bold("API (Application Programming Interface): "), run("Uma interface que permite que dois sistemas se comuniquem. Profissionais de dados frequentemente usam APIs para coletar dados de fontes externas como redes sociais, plataformas de e-commerce e sistemas de terceiros.")]),
    spacer(),
    sectionDivider(TEAL),
    spacer(),
    h2("Prompts Interativos — Seu Consultor de IA em Dados"),
    body("Os prompts abaixo foram criados para que você use com ferramentas de IA generativa (como Claude ou ChatGPT) e obtenha orientações personalizadas para sua transição de carreira. Copie, adapte e use!"),
    spacer(),
    h3("Prompt 1: Entrevista Vocacional em Dados"),
    callout([
      run("Atue como meu Conselheiro de Carreira especializado em transição para a área de dados. Conduza comigo uma entrevista vocacional estruturada para descobrir qual das três carreiras — Analista de Dados, Engenheiro de Dados ou Cientista de Dados — é mais adequada ao meu perfil atual. Faça perguntas sobre: 1) minha formação e área de experiência atual; 2) minhas habilidades técnicas existentes; 3) meus objetivos de carreira de curto e longo prazo; 4) minha tolerância a aspectos técnicos profundos; 5) o tipo de problema que mais me motiva resolver. Ao final, apresente uma recomendação de carreira com justificativa, uma trilha de estudos personalizada para os primeiros 90 dias e três tipos de projeto de portfólio que eu devo construir.")
    ], GRAY_BG, MID_TEXT),
    spacer(),
    h3("Prompt 2: Tradutor de Carreira"),
    callout([
      run("Eu venho da área de [INSIRA SUA ÁREA AQUI — ex: Saúde, Educação, Logística, Finanças]. Quais habilidades transferíveis do meu setor atual são mais valorizadas na área de dados? Me ajude a: 1) identificar as 5 principais competências da minha área que têm alto valor em dados; 2) traduzir meu currículo atual para o vocabulário de dados; 3) sugerir tipos de projeto de portfólio que aproveitem meu conhecimento de domínio; 4) criar um discurso de apresentação (elevator pitch) para usar em entrevistas, explicando por que minha transição agrega valor único.")
    ], GRAY_BG, MID_TEXT),
    spacer(),
    h3("Prompt 3: Gerador de Projetos de Portfólio"),
    callout([
      run("Sou um profissional em transição para dados, vindo da área de [SUA ÁREA]. Meu nível técnico atual é [iniciante/intermediário]. Quero construir um projeto de portfólio que: 1) use dados públicos disponíveis gratuitamente; 2) aplique o conhecimento da minha área de origem; 3) seja relevante para empresas do setor de [SUA ÁREA OU SETOR ALVO]; 4) demonstre as habilidades de [SQL/Python/Power BI — escolha uma]. Me proponha 3 ideias de projeto com título, objetivo, fonte de dados sugerida, ferramentas a usar e o que o projeto vai demonstrar para recrutadores.")
    ], GRAY_BG, MID_TEXT),
    spacer(),
    h3("Prompt 4: Simulador de Entrevista Técnica"),
    callout([
      run("Atue como um recrutador técnico de uma empresa de médio porte que está contratando um Analista de Dados Júnior. Conduza comigo uma entrevista simulada com 10 perguntas — misturando questões técnicas (SQL, visualização de dados, métricas) e comportamentais (resolução de problemas, comunicação de resultados). Após cada resposta minha, avalie minha resposta numa escala de 1 a 5, aponte pontos fortes e áreas de melhoria, e sugira como eu poderia ter respondido de forma mais forte. Ao final, dê um feedback geral sobre minha prontidão para o mercado.")
    ], GRAY_BG, MID_TEXT),
    spacer(),
    sectionDivider(ACCENT),
    spacer(),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 200, after: 200 },
      children: [
        new TextRun({ text: "Boa jornada. Os dados esperam por você.", font: "Arial", size: 26, bold: true, color: BLUE, italics: true })
      ]
    })
  ];
}

// ─── BUILD DOCUMENT ───────────────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "•",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      },
      {
        reference: "numbers",
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: "%1.",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      }
    ]
  },
  styles: {
    default: {
      document: { run: { font: "Arial", size: 22, color: DARK_TEXT } }
    },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: "Arial", color: BLUE },
        paragraph: { spacing: { before: 360, after: 180 }, outlineLevel: 0 }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: TEAL },
        paragraph: { spacing: { before: 280, after: 120 }, outlineLevel: 1 }
      },
      {
        id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Arial", color: ACCENT },
        paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 2 }
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 }, // A4
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    children: [
      ...coverPage(),
      ...titlesPage(),
      ...introChapter(),
      ...chapter1(),
      ...chapter2(),
      ...chapter3(),
      ...conclusion(),
      ...appendix()
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('/mnt/user-data/outputs/Ebook_Do_Zero_aos_Dados.docx', buffer);
  console.log('Done!');
}).catch(err => {
  console.error(err);
  process.exit(1);
});
// ... continuação do apêndice
    bullet([bold("Dashboard: "), run("Um painel visual interativo que consolida métricas e indicadores-chave de desempenho (KPIs) para monitoramento rápido e tomada de decisão.")]),
    spacer()
  ];
}

// ─── GERAR ARQUIVO FINAL ──────────────────────────────────────────────────────
const doc = new Document({
  sections: [{
    properties: {},
    children: [
      ...coverPage(),
      ...titlesPage(),
      ...introChapter(),
      ...chapter1(),
      ...chapter2(),
      ...chapter3(),
      ...conclusion(),
      ...appendix()
    ],
  }],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("Do_Zero_Aos_Dados.docx", buffer);
  console.log("E-book e Guia gerados com sucesso!");
});
