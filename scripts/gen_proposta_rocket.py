# -*- coding: utf-8 -*-
"""Gera a Proposta Comercial da Máquina de Venda de Franquias para a Rocket SF."""
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ---- Paleta ----
NAVY = RGBColor(0x0B, 0x1F, 0x3A)
ORANGE = RGBColor(0xE4, 0x57, 0x1B)
GREY = RGBColor(0x5A, 0x63, 0x72)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT = "EEF1F6"
HEADER_FILL = "0B1F3A"
ACCENT_FILL = "E4571B"

doc = Document()

# ---- Base style ----
normal = doc.styles["Normal"]
normal.font.name = "Calibri"
normal.font.size = Pt(10.5)
normal.font.color.rgb = RGBColor(0x1A, 0x1A, 0x1A)
normal.paragraph_format.space_after = Pt(6)
normal.paragraph_format.line_spacing = 1.12

for hstyle, sz, col in [("Heading 1", 16, NAVY), ("Heading 2", 12.5, NAVY), ("Heading 3", 11, ORANGE)]:
    st = doc.styles[hstyle]
    st.font.name = "Calibri"
    st.font.size = Pt(sz)
    st.font.color.rgb = col
    st.font.bold = True
    st.paragraph_format.space_before = Pt(12)
    st.paragraph_format.space_after = Pt(4)


def shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    sh = OxmlElement("w:shd")
    sh.set(qn("w:val"), "clear")
    sh.set(qn("w:fill"), fill)
    tcPr.append(sh)


def set_cell_text(cell, text, bold=False, color=None, size=9.5, align=None):
    cell.text = ""
    p = cell.paragraphs[0]
    if align:
        p.alignment = align
    for j, line in enumerate(str(text).split("\n")):
        if j > 0:
            r = p.add_run()
            r.add_break()
        else:
            r = p.add_run(line)
            continue
        r = p.add_run(line)
    # simpler: rebuild
    cell.text = ""
    p = cell.paragraphs[0]
    if align:
        p.alignment = align
    parts = str(text).split("\n")
    for j, line in enumerate(parts):
        run = p.add_run(line if j == 0 else line)
        if j < len(parts) - 1:
            run.add_break()
        run.font.size = Pt(size)
        run.font.bold = bold
        if color:
            run.font.color.rgb = color
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.space_before = Pt(2)


def add_table(headers, rows, widths=None, header_fill=HEADER_FILL, first_col_bold=False, caption=None):
    if caption:
        cp = doc.add_paragraph()
        cr = cp.add_run(caption)
        cr.font.size = Pt(9)
        cr.font.italic = True
        cr.font.color.rgb = GREY
        cp.paragraph_format.space_after = Pt(2)
    ncols = len(headers)
    t = doc.add_table(rows=1, cols=ncols)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = True
    hdr = t.rows[0].cells
    for i, h in enumerate(headers):
        set_cell_text(hdr[i], h, bold=True, color=WHITE, size=9.5)
        shade(hdr[i], header_fill)
    for r, row in enumerate(rows):
        cells = t.add_row().cells
        for i, val in enumerate(row):
            bold = first_col_bold and i == 0
            set_cell_text(cells[i], val, bold=bold, size=9.5)
            if r % 2 == 1:
                shade(cells[i], LIGHT)
    if widths:
        for i, w in enumerate(widths):
            for row in t.rows:
                row.cells[i].width = Cm(w)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def para(text, size=10.5, italic=False, color=None, bold=False, space_after=6, align=None):
    p = doc.add_paragraph()
    if align:
        p.alignment = align
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.font.italic = italic
    r.font.bold = bold
    if color:
        r.font.color.rgb = color
    p.paragraph_format.space_after = Pt(space_after)
    return p


def bullets(items, style="List Bullet"):
    for it in items:
        p = doc.add_paragraph(style=style)
        if isinstance(it, tuple):
            r = p.add_run(it[0])
            r.font.bold = True
            r.font.size = Pt(10.5)
            r2 = p.add_run(it[1])
            r2.font.size = Pt(10.5)
        else:
            r = p.add_run(it)
            r.font.size = Pt(10.5)
        p.paragraph_format.space_after = Pt(2)


def callout(title, text, fill=ACCENT_FILL):
    t = doc.add_table(rows=1, cols=1)
    t.style = "Table Grid"
    c = t.rows[0].cells[0]
    c.text = ""
    p = c.paragraphs[0]
    r = p.add_run(title)
    r.font.bold = True
    r.font.color.rgb = WHITE
    r.font.size = Pt(10.5)
    p2 = c.add_paragraph()
    r2 = p2.add_run(text)
    r2.font.color.rgb = WHITE
    r2.font.size = Pt(9.5)
    shade(c, fill)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


# ============================================================
# CAPA
# ============================================================
def cover_line(text, size, color, bold=False, italic=False, after=2, before=0):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.font.color.rgb = color
    r.font.bold = bold
    r.font.italic = italic
    r.font.name = "Calibri"
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.space_before = Pt(before)
    return p

for _ in range(2):
    doc.add_paragraph()
cover_line("PROPOSTA COMERCIAL", 13, ORANGE, bold=True, after=4)
cover_line("Máquina de Venda de Franquias", 30, NAVY, bold=True, after=2)
cover_line("Estruturação e operação ponta a ponta do Go-to-Market de expansão de franquias",
           13, GREY, italic=True, after=18)
cover_line("Estratégia → Oferta → Aquisição → Conversão → CRM → Pré-vendas → Vendas → Gestão → Otimização → Escala",
           11, NAVY, bold=True, after=24)

cover_line("PREPARADO PARA", 11, GREY, bold=True, after=2, before=10)
cover_line("Rocket Soluções Financeiras (Rocket SF)", 15, NAVY, bold=True, after=1)
cover_line("Crédito inteligente e consórcio para o transporte rodoviário · Chapecó — SC", 10, GREY, after=18)

cover_line("APRESENTADO POR", 11, GREY, bold=True, after=2)
cover_line("[Sua Empresa] — Parceiro de Go-to-Market & Operação Comercial", 13, NAVY, bold=True, after=1)
cover_line("Responsável: [Nome]  ·  [e-mail]  ·  [telefone]", 10, GREY, after=24)

cover_line("Documento confidencial  ·  Versão 1.0  ·  Agosto de 2026", 9.5, GREY, italic=True, after=2)

doc.add_page_break()

# ============================================================
# SUMÁRIO
# ============================================================
doc.add_heading("Sumário", level=1)
sumario = [
    "1. Sumário Executivo",
    "2. Enquadramento — o que estamos (e o que não estamos) assumindo",
    "3. Estratégia e Go-to-Market",
    "4. Estrutura de Marketing",
    "5. Mídia e Aquisição",
    "6. Forecasting de Aquisição e Vendas",
    "7. Infraestrutura Comercial — CRM dedicado",
    "8. Construção do Time Comercial",
    "9. Processo de Venda",
    "10. Gestão e Otimização da Operação",
    "11. Árvore de Indicadores",
    "12. Cronograma de Implantação",
    "13. Matriz de Responsabilidades",
    "14. Custos da Operação",
    "15. Premissas e Limites",
    "16. Governança",
    "17. Plano de Escala",
    "18. Perguntas-chave respondidas",
    "19. Próximos Passos",
    "20. Pontos que ainda precisam ser definidos antes do envio da proposta",
]
for s in sumario:
    p = doc.add_paragraph()
    r = p.add_run(s)
    r.font.size = Pt(10.5)
    r.font.color.rgb = NAVY
    p.paragraph_format.space_after = Pt(2)

doc.add_page_break()

# ============================================================
# 1. SUMÁRIO EXECUTIVO
# ============================================================
doc.add_heading("1. Sumário Executivo", level=1)
para("A Rocket Soluções Financeiras (Rocket SF) construiu um modelo de negócio validado em crédito "
     "inteligente e consórcio para o setor de transporte rodoviário e decidiu expandir sua operação por "
     "meio da venda de franquias. O desafio dessa expansão não é apenas gerar interesse: é transformar, de "
     "forma previsível e escalável, investimento em mídia em novos franqueados assinados, prontos para "
     "operar a marca.")
para("Esta proposta apresenta a terceirização da construção e da operação dessa máquina de expansão. "
     "Não entramos como agência de marketing, consultor pontual ou fornecedor de leads. Assumimos o "
     "Go-to-Market completo da venda de franquias — da estratégia à escala — construindo e operando cada "
     "nó da jornada que separa o investimento em mídia da assinatura de um novo franqueado.")

callout("Proposta de valor em uma frase",
        "Uma máquina de expansão de franquias previsível, mensurável e governada — construída, operada e "
        "otimizada de ponta a ponta por um parceiro dedicado — que converte investimento em mídia em "
        "franqueados assinados, com custo de aquisição conhecido e capacidade de escala.")

doc.add_heading("O ciclo que assumimos", level=2)
para("Estratégia → Oferta → Aquisição → Conversão → CRM → Pré-vendas → Vendas → Gestão → Otimização → Escala.",
     bold=True, color=NAVY)
para("Cada elo desse ciclo é um entregável concreto desta operação, descrito nas seções a seguir. O ponto "
     "de partida do investimento em mídia é de aproximadamente R$ 10.000/mês, escalado progressivamente "
     "conforme os resultados do funil e a capacidade comercial instalada.")

add_table(
    ["Indicador de referência", "Situação inicial", "Meta"],
    [
        ["Investimento inicial em mídia", "≈ R$ 10.000 / mês", "Escala progressiva orientada a resultado"],
        ["Time comercial inicial", "0 dedicado à franquia", "2 pré-vendas + 1 closer"],
        ["CRM de franquias", "Inexistente / compartilhado", "CRM dedicado e governado"],
        ["Máquina de aquisição", "A construir", "Campanhas no ar em ~[X] semanas"],
        ["Metas de funil (leads → franquias)", "A calibrar", "Ver Seção 6 — Forecasting"],
    ],
    widths=[6.5, 5.0, 6.0],
    caption="Metas quantitativas de funil e de vendas serão consolidadas no forecasting (Seção 6), a partir "
            "dos dados de oferta, ticket e capacidade fornecidos pela Rocket SF. Nenhum número de resultado "
            "é assumido sem essa calibração.",
)

# ============================================================
# 2. ENQUADRAMENTO
# ============================================================
doc.add_heading("2. Enquadramento — o que estamos (e o que não estamos) assumindo", level=1)
para("Para que não reste ambiguidade sobre a natureza do vínculo, o quadro abaixo delimita o escopo antes "
     "de detalhá-lo. A percepção correta é: a Rocket SF terceiriza para este parceiro a construção e a "
     "operação da sua máquina de expansão de franquias.")

add_table(
    ["Estamos assumindo", "Não estamos assumindo"],
    [
        ["Desenho completo do Go-to-Market da expansão de franquias",
         "A concessão jurídica da franquia e a titularidade da marca (permanecem com a Rocket SF)"],
        ["Construção e gestão da estrutura de marketing (copy, design, tráfego)",
         "Definição do modelo de negócio, taxas e regras da franquia (competência da Rocket SF)"],
        ["Operação de mídia e aquisição de candidatos",
         "Elaboração da COF e dos instrumentos jurídicos (advogado/jurídico da Rocket SF)"],
        ["Implantação e operação do CRM dedicado de franquias",
         "Aprovação final de cada candidato a franqueado (decisão da Rocket SF)"],
        ["Construção, gestão e performance do time de pré-vendas e vendas",
         "Implantação, treinamento operacional e onboarding do franqueado após a assinatura"],
        ["Gestão e otimização contínua de todo o funil comercial",
         "Garantia de um número específico de vendas independente das variáveis da Seção 15"],
    ],
    widths=[8.75, 8.75],
)
callout("Posicionamento",
        "Isto não é uma prestação de serviço de marketing avulsa. É a operação terceirizada de uma área "
        "comercial inteira — a máquina de expansão de franquias da Rocket SF — com responsabilidade sobre a "
        "construção, a implantação, a gestão e a otimização do sistema que gera franqueados.",
        fill=HEADER_FILL)

# ============================================================
# 3. ESTRATÉGIA E GTM
# ============================================================
doc.add_heading("3. Estratégia e Go-to-Market", level=1)
para("O primeiro bloco de trabalho é estratégico: construímos a fundação comercial completa da expansão. "
     "Sem essa camada, mídia vira custo; com ela, mídia vira franquias vendidas. Construiremos todos os nós "
     "da jornada necessários para transformar investimento em mídia em franquias assinadas.")

doc.add_heading("3.1. Entregáveis da estratégia", level=2)
add_table(
    ["Bloco", "O que construímos"],
    [
        ["ICP do franqueado", "Perfil ideal do franqueado (capacidade de investimento, perfil empreendedor, "
         "região, aderência ao setor de transporte/financeiro)"],
        ["Personas prioritárias", "Personas de decisão e seus contextos (investidor, operador, empresário "
         "do setor, profissional em transição de carreira)"],
        ["Gatilhos de compra", "Principais motivadores que levam alguém a comprar uma franquia da Rocket SF"],
        ["Proposta de valor da franquia", "Articulação clara do porquê investir na Rocket SF versus "
         "alternativas de investimento e outras franquias"],
        ["Argumentos comerciais", "Repositório de argumentos por persona e por etapa do funil"],
        ["Objeções", "Mapa das principais objeções (preço, retorno, risco, território) e respostas"],
        ["Estruturação da oferta", "Formatação da oferta de franquia para maximizar conversão (condições, "
         "ancoragem, urgência legítima, prova)"],
        ["Canais de aquisição", "Definição e priorização dos canais (Meta, Google, LinkedIn, portais de "
         "franquia, indicação, outbound)"],
        ["Jornada do candidato", "Desenho completo do caminho do primeiro clique à assinatura"],
        ["Funil comercial", "Estágios do funil, critérios de passagem e metas por etapa"],
        ["Metas e indicadores", "Metas por etapa e árvore de indicadores (Seção 11)"],
        ["Estratégia de escala", "Regras objetivas para crescer mídia e time (Seção 17)"],
    ],
    widths=[5.0, 12.5],
    first_col_bold=True,
)

doc.add_heading("3.2. BoltAI na construção e operação da jornada", level=2)
para("Utilizaremos o BoltAI como camada de apoio à construção, automação, análise e otimização da jornada: "
     "aceleração da produção de ativos (ângulos, copies, variações de criativo e páginas), automação de "
     "etapas do funil e das cadências, análise de dados de conversão e geração de hipóteses de otimização. "
     "O BoltAI amplia a velocidade e a profundidade da operação — não substitui a governança humana das "
     "decisões comerciais e de compliance.")

# ============================================================
# 4. ESTRUTURA DE MARKETING
# ============================================================
doc.add_heading("4. Estrutura de Marketing", level=1)
para("Montamos e gerimos a estrutura de marketing necessária para a aquisição de candidatos. Isso inclui a "
     "contratação e a gestão dos profissionais e a responsabilidade por toda a produção de aquisição.")

doc.add_heading("4.1. Profissionais e responsabilidades", level=2)
add_table(
    ["Função", "Responsabilidades"],
    [
        ["Copywriter", "Conceitos, ângulos, copies de anúncios, páginas e cadências"],
        ["Designer", "Criativos estáticos e em vídeo, identidade das campanhas, layout das landing pages"],
        ["Gestor de tráfego", "Estruturação, veiculação, otimização e escala das campanhas de mídia paga e "
         "remarketing"],
    ],
    widths=[4.5, 13.0],
    first_col_bold=True,
)
para("Entregáveis dessa frente: criação das campanhas; desenvolvimento de conceitos e ângulos; copies; "
     "criativos; landing pages; páginas/formulários de captação; campanhas de aquisição; remarketing; "
     "testes de ofertas; testes de criativos; e otimização contínua.")

doc.add_heading("4.2. Quem contrata, gere e paga", level=2)
add_table(
    ["Item", "Contratação / Gestão", "Pagamento"],
    [
        ["Copywriter", "[Sua Empresa]", "[a definir — ver Seção 14]"],
        ["Designer", "[Sua Empresa]", "[a definir — ver Seção 14]"],
        ["Gestor de tráfego", "[Sua Empresa]", "[a definir — ver Seção 14]"],
    ],
    widths=[5.0, 6.0, 6.5],
    first_col_bold=True,
    caption="A responsabilidade pela seleção e gestão desses profissionais é de [Sua Empresa]. A definição "
            "de quem arca com a remuneração de cada um — se está incluída na remuneração da operação ou é "
            "custo à parte da Rocket SF — é consolidada na Seção 14 (Custos da Operação).",
)

# ============================================================
# 5. MÍDIA E AQUISIÇÃO
# ============================================================
doc.add_heading("5. Mídia e Aquisição", level=1)
para("A operação começa com aproximadamente R$ 10.000/mês de investimento em mídia. Esse valor é o ponto "
     "de partida e será escalado progressivamente conforme os resultados do funil e a capacidade do time "
     "comercial de absorver o volume — nunca antes de o funil comprovar eficiência.")
para("A lógica de aquisição segue a cadeia abaixo. Ela conecta cada real investido ao resultado final e é "
     "o que permite decidir, com dado, quando escalar:", bold=True, color=NAVY)
add_table(
    ["Etapa da cadeia", "O que mede"],
    [
        ["Investimento em mídia", "Capital alocado em aquisição no período"],
        ["Leads", "Candidatos captados"],
        ["Leads qualificados", "Candidatos com fit de perfil e capacidade de investimento"],
        ["Reuniões", "Reuniões de apresentação realizadas"],
        ["Propostas / aprovações", "Candidatos que avançam para proposta e aprovação"],
        ["Franquias vendidas", "Contratos de franquia assinados"],
        ["CAC por franquia", "Custo total de aquisição dividido pelas franquias vendidas"],
        ["Receita gerada", "Receita de taxa de franquia e recorrências decorrentes"],
    ],
    widths=[6.0, 11.5],
    first_col_bold=True,
)
para("A modelagem quantitativa dessa cadeia — volumes, taxas de conversão e valores — é apresentada no "
     "forecasting (Seção 6), a ser fornecido separadamente. Nenhum número é assumido nesta proposta além do "
     "investimento inicial de mídia informado pela Rocket SF.")

# ============================================================
# 6. FORECASTING
# ============================================================
doc.add_heading("6. Forecasting de Aquisição e Vendas", level=1)
para("Esta seção é o espaço reservado para incorporar o forecasting de aquisição e vendas, fornecido "
     "separadamente. A estrutura abaixo espelha a cadeia da Seção 5 e será preenchida com os números "
     "calibrados a partir da oferta, do ticket da franquia e da capacidade comercial.")
add_table(
    ["Métrica", "Mês 1", "Mês 2", "Mês 3", "Mês 4", "Mês 5", "Mês 6"],
    [
        ["Investimento em mídia (R$)", "", "", "", "", "", ""],
        ["Leads", "", "", "", "", "", ""],
        ["Leads qualificados", "", "", "", "", "", ""],
        ["Reuniões realizadas", "", "", "", "", "", ""],
        ["Propostas / aprovações", "", "", "", "", "", ""],
        ["Franquias vendidas", "", "", "", "", "", ""],
        ["CAC por franquia (R$)", "", "", "", "", "", ""],
        ["Receita gerada (R$)", "", "", "", "", "", ""],
    ],
    widths=[4.7, 2.13, 2.13, 2.13, 2.13, 2.13, 2.13],
    caption="Campos a preencher com o forecasting fornecido pela Rocket SF / [Sua Empresa]. Os valores são "
            "premissas de planejamento, revisadas mensalmente contra o realizado (Seção 16 — Governança).",
)
callout("Como leremos o forecasting",
        "O forecasting é um instrumento de gestão, não uma promessa de resultado. Ele define as premissas "
        "de conversão por etapa; a operação existe justamente para perseguir e, quando possível, superar "
        "essas taxas — e para recalibrá-las quando o mercado responder diferente do previsto.",
        fill=HEADER_FILL)

# ============================================================
# 7. CRM
# ============================================================
doc.add_heading("7. Infraestrutura Comercial — CRM dedicado", level=1)
para("Implantamos um CRM separado e dedicado exclusivamente à venda de franquias — isolado das demais "
     "operações da Rocket SF —, garantindo que o pipeline de expansão tenha dados limpos, rastreáveis e "
     "próprios. O projeto contempla:")
bullets([
    "Definição das etapas do pipeline;",
    "Configuração do CRM;",
    "Integração dos canais de aquisição (formulários, landing pages, mídia);",
    "Distribuição dos leads (roteamento e rodízio entre pré-vendas);",
    "Automações;",
    "Cadências comerciais;",
    "Tarefas e SLAs;",
    "Motivos de perda;",
    "Dashboards;",
    "Acompanhamento de conversão por etapa;",
    "Acompanhamento da produtividade comercial.",
])
doc.add_heading("7.1. Pipeline de referência", level=2)
add_table(
    ["Etapa do pipeline", "Definição / gatilho de avanço"],
    [
        ["Novo lead", "Candidato captado, ainda não trabalhado"],
        ["Em qualificação", "Pré-vendas em contato para validar fit e capacidade"],
        ["Qualificado (MQL→SQL)", "Fit e capacidade confirmados; apto a reunião"],
        ["Reunião agendada", "Reunião de apresentação marcada"],
        ["Reunião realizada", "Apresentação da franquia concluída"],
        ["Em análise/proposta", "Candidato em análise e proposta apresentada"],
        ["Negociação / documentação", "Negociação e trâmite de COF/instrumentos, quando aplicável"],
        ["Fechado — ganho", "Contrato assinado; handoff para implantação"],
        ["Fechado — perdido", "Registro com motivo de perda estruturado"],
    ],
    widths=[5.5, 12.0],
    first_col_bold=True,
    caption="As etapas são calibradas na Fase 1 conforme o ciclo real de venda da franquia Rocket SF.",
)

# ============================================================
# 8. TIME COMERCIAL
# ============================================================
doc.add_heading("8. Construção do Time Comercial", level=1)
para("Estruturamos inicialmente um time dedicado composto por 2 profissionais de pré-vendas (SDRs) + 1 "
     "closer, responsável por conduzir o candidato da qualificação ao fechamento.")
doc.add_heading("8.1. Nossa responsabilidade sobre o time", level=2)
bullets([
    "Definição do perfil das vagas;",
    "Apoio e condução do recrutamento;",
    "Seleção;",
    "Onboarding;",
    "Treinamento;",
    "Scripts;",
    "Playbook comercial;",
    "Cadências;",
    "Definição de metas;",
    "Definição dos KPIs;",
    "Rotina de gestão;",
    "Acompanhamento da performance.",
])
para("Nós gerimos esse time no dia a dia. A definição de quem arca com a remuneração (salário/variável) "
     "de cada profissional é consolidada na Seção 14.")

doc.add_heading("8.2. Estrutura inicial e crescimento", level=2)
add_table(
    ["Momento", "Estrutura comercial", "Gatilho de expansão"],
    [
        ["Lançamento", "2 pré-vendas + 1 closer", "—"],
        ["Escala 1", "+1 pré-venda", "Volume de leads qualificados acima da capacidade de contato/SLA"],
        ["Escala 2", "+1 closer", "Reuniões qualificadas acima da capacidade de fechamento do closer atual"],
        ["Escala 3", "Coordenação/liderança comercial", "Time atinge tamanho que exige gestão dedicada em camada"],
    ],
    widths=[3.5, 6.0, 8.0],
    first_col_bold=True,
    caption="A estrutura cresce puxada pelo funil: só se adiciona capacidade comercial quando o gargalo "
            "comprovadamente migra para a etapa correspondente. Critérios objetivos na Seção 17.",
)

# ============================================================
# 9. PROCESSO DE VENDA
# ============================================================
doc.add_heading("9. Processo de Venda", level=1)
para("Estruturamos a jornada comercial completa, do primeiro contato à assinatura do novo franqueado, com "
     "início e fim de responsabilidade explícitos.")
add_table(
    ["Etapa", "Descrição", "Responsável"],
    [
        ["Lead", "Entrada do candidato pelos canais de aquisição", "Marketing"],
        ["Qualificação", "Validação de fit, perfil e capacidade de investimento", "Pré-vendas"],
        ["Contato", "Abordagem e engajamento dentro do SLA de primeira resposta", "Pré-vendas"],
        ["Reunião", "Agendamento e realização da reunião", "Pré-vendas → Closer"],
        ["Apresentação da franquia", "Apresentação do modelo, retorno e condições", "Closer"],
        ["Análise do candidato", "Avaliação de aderência e capacidade", "Closer + Rocket SF"],
        ["Follow-up", "Cadência de acompanhamento até a decisão", "Closer"],
        ["Negociação", "Ajuste de condições e alinhamento final", "Closer"],
        ["Documentação / COF", "Entrega da COF e instrumentos jurídicos, quando aplicável", "Rocket SF (jurídico)"],
        ["Fechamento", "Aprovação final do candidato", "Rocket SF"],
        ["Assinatura", "Assinatura do contrato de franquia", "Rocket SF + Closer"],
        ["Handoff", "Passagem para implantação/onboarding do franqueado", "[Sua Empresa] → Rocket SF"],
    ],
    widths=[3.8, 9.7, 4.0],
    first_col_bold=True,
)
callout("Onde começa e onde termina a operação comercial",
        "Começa no lead (primeiro contato gerado pela máquina de aquisição). Termina no handoff — a "
        "passagem do franqueado assinado para a equipe de implantação/onboarding da Rocket SF. A entrega "
        "operacional da franquia (treinamento, sistemas, abertura da unidade) é responsabilidade da "
        "franqueadora.", fill=HEADER_FILL)

# ============================================================
# 10. GESTÃO
# ============================================================
doc.add_heading("10. Gestão e Otimização da Operação", level=1)
para("A operação não termina na implantação. Depois de colocar a máquina no ar, assumimos a gestão e a "
     "otimização contínua do funil. A rotina de gestão inclui:")
add_table(
    ["Frente de gestão", "O que fazemos"],
    [
        ["Ritmo de acompanhamento", "Reuniões de acompanhamento e análise semanal do funil"],
        ["Marketing", "Gestão da estrutura de mídia, criativos e páginas"],
        ["Pré-vendas", "Gestão dos SDRs, SLAs, cadências e qualificação"],
        ["Vendas", "Gestão do closer, taxa de fechamento e ciclo de venda"],
        ["Conversão", "Análise de conversão por etapa e identificação de gargalos"],
        ["Objeções", "Análise das principais objeções e ajuste de argumentos"],
        ["Scripts e campanhas", "Revisão de scripts e revisão das campanhas"],
        ["Eficiência", "Otimização de CAC e das taxas de conversão"],
        ["Decisões de mídia", "Decisões de aumento/redução de investimento por dado"],
        ["Planejamento", "Forecasting e plano de ação periódico"],
    ],
    widths=[4.5, 13.0],
    first_col_bold=True,
)

# ============================================================
# 11. INDICADORES
# ============================================================
doc.add_heading("11. Árvore de Indicadores", level=1)
para("A operação é gerida por uma árvore de indicadores que conecta mídia, pré-vendas, vendas e economia. "
     "É por ela que sabemos se o Go-to-Market está funcionando.")

doc.add_heading("11.1. Marketing", level=2)
add_table(["Indicador", "Descrição"],
          [["Investimento", "Capital aplicado em mídia"],
           ["Impressões", "Volume de exibições"],
           ["CTR", "Taxa de cliques"],
           ["CPC", "Custo por clique"],
           ["Leads", "Candidatos captados"],
           ["CPL", "Custo por lead"],
           ["Conversão das páginas", "Taxa de conversão das landing pages"]],
          widths=[5.0, 12.5], first_col_bold=True)

doc.add_heading("11.2. Pré-vendas", level=2)
add_table(["Indicador", "Descrição"],
          [["Leads trabalhados", "Volume de leads efetivamente trabalhados"],
           ["Tempo de primeira resposta", "Velocidade do primeiro contato"],
           ["Taxa de contato", "% de leads contatados com sucesso"],
           ["Taxa de qualificação", "% de leads que se tornam qualificados"],
           ["Reuniões agendadas", "Volume de reuniões marcadas"],
           ["Show rate", "% de comparecimento às reuniões"]],
          widths=[5.0, 12.5], first_col_bold=True)

doc.add_heading("11.3. Vendas", level=2)
add_table(["Indicador", "Descrição"],
          [["Reuniões realizadas", "Reuniões efetivamente concluídas"],
           ["Taxa de avanço", "% de reuniões que avançam no funil"],
           ["Propostas / candidatos aprovados", "Volume de propostas e aprovações"],
           ["Taxa de fechamento", "% de fechamento sobre oportunidades"],
           ["Ciclo médio de venda", "Tempo médio do lead à assinatura"],
           ["Franquias vendidas", "Contratos assinados no período"]],
          widths=[5.0, 12.5], first_col_bold=True)

doc.add_heading("11.4. Economia", level=2)
add_table(["Indicador", "Descrição"],
          [["CAC por franquia", "Custo de aquisição por franquia vendida"],
           ["Investimento total de aquisição", "Soma de mídia, time e ferramentas alocados"],
           ["Receita de taxa de franquia", "Receita gerada pelas assinaturas"],
           ["Payback da aquisição", "Tempo para recuperar o investimento de aquisição"],
           ["ROI / ROAS", "Retorno sobre investimento/mídia, quando fizer sentido"]],
          widths=[5.0, 12.5], first_col_bold=True)

# ============================================================
# 12. CRONOGRAMA
# ============================================================
doc.add_heading("12. Cronograma de Implantação", level=1)
para("O cronograma abaixo detalha as primeiras semanas, com dependências, entregáveis e critério de "
     "conclusão por fase. Os prazos são estimativas de implantação; a geração de resultado (primeiras "
     "vendas) depende do ciclo de venda da franquia e das variáveis da Seção 15 — e por isso é diferenciada "
     "explicitamente do prazo de montagem.")

doc.add_heading("12.1. Visão geral", level=2)
add_table(
    ["Fase", "Prazo estimado", "Foco"],
    [
        ["Fase 1 — Estratégia e preparação", "Semanas 1–3", "Diagnóstico, GTM, ICP, oferta, jornada, funil, forecasting, metas"],
        ["Fase 2 — Construção", "Semanas 2–6", "Contratações, CRM, integrações, landing pages, criativos, scripts, playbook"],
        ["Fase 3 — Go Live", "Semanas 5–7", "Campanhas no ar, primeiros leads, operação dos SDRs, primeiras reuniões"],
        ["Fase 4 — Otimização e escala", "Semana 8 em diante", "Análise de conversão, ajustes, treinamento, escala do investimento"],
    ],
    widths=[5.5, 3.2, 8.8],
    first_col_bold=True,
    caption="Prazos com sobreposição intencional entre fases para antecipar valor. Datas exatas dependem da "
            "velocidade de resposta e disponibilização de informações pela Rocket SF (Seção 13).",
)

def phase_table(title, rows):
    doc.add_heading(title, level=3)
    add_table(
        ["Entregável", "Responsável", "Dependências", "Critério de conclusão"],
        rows,
        widths=[4.6, 3.4, 4.5, 5.0],
    )

doc.add_heading("12.2. Detalhamento por fase", level=2)
phase_table("Fase 1 — Estratégia e preparação (Semanas 1–3)", [
    ["Diagnóstico", "[Sua Empresa]", "Acesso a dados e sócios da Rocket SF", "Diagnóstico consolidado e validado"],
    ["GTM + ICP + personas", "[Sua Empresa]", "Diagnóstico", "GTM aprovado pela Rocket SF"],
    ["Oferta da franquia", "[Sua Empresa] + Rocket SF", "Modelo, taxas e condições da franquia", "Oferta homologada"],
    ["Jornada + funil", "[Sua Empresa]", "GTM e oferta", "Jornada e funil aprovados"],
    ["Forecasting + metas", "[Sua Empresa] + Rocket SF", "Ticket e capacidade", "Forecasting e metas homologados"],
])
phase_table("Fase 2 — Construção (Semanas 2–6)", [
    ["Copywriter, designer, gestor de tráfego", "[Sua Empresa]", "GTM aprovado", "Profissionais contratados e integrados"],
    ["SDRs + closer", "[Sua Empresa]", "Perfil de vagas + orçamento", "Time contratado e em onboarding"],
    ["CRM + integrações", "[Sua Empresa]", "Definição de pipeline", "CRM configurado e integrado aos canais"],
    ["Landing pages + criativos", "[Sua Empresa]", "Oferta e conceitos aprovados", "Páginas e criativos aprovados"],
    ["Scripts + playbook", "[Sua Empresa]", "Jornada e oferta", "Playbook e scripts prontos e treinados"],
])
phase_table("Fase 3 — Go Live (Semanas 5–7)", [
    ["Campanhas no ar", "[Sua Empresa]", "Criativos, páginas, verba de mídia liberada", "Campanhas ativas veiculando"],
    ["Primeiros leads", "[Sua Empresa]", "Campanhas no ar", "Leads entrando no CRM"],
    ["Operação dos SDRs", "[Sua Empresa]", "Time treinado + CRM", "SDRs trabalhando leads dentro do SLA"],
    ["Primeiras reuniões", "[Sua Empresa] + Rocket SF", "Leads qualificados", "Reuniões de apresentação realizadas"],
])
phase_table("Fase 4 — Otimização e escala (Semana 8 em diante)", [
    ["Análise de conversão", "[Sua Empresa]", "Volume mínimo de dados", "Gargalos identificados por etapa"],
    ["Ajustes de mídia/oferta/copy", "[Sua Empresa]", "Análise de conversão", "Ajustes implementados e medidos"],
    ["Ajustes comerciais + treinamento", "[Sua Empresa]", "Análise do time", "Scripts e cadências revisados"],
    ["Escala do investimento", "[Sua Empresa] + Rocket SF", "Eficiência comprovada (Seção 17)", "Mídia escalada com CAC controlado"],
])
callout("Prazo de implantação ≠ prazo de resultado",
        "A montagem da máquina segue os prazos acima. A maturação do resultado (primeiras vendas de "
        "franquia) acompanha o ciclo de venda — que envolve análise, aprovação e documentação — e será "
        "estimada no forecasting. Não prometemos datas de venda; comprometemo-nos com a construção, a "
        "operação e a otimização diligente da máquina que as produz.", fill=HEADER_FILL)

# ============================================================
# 13. RESPONSABILIDADES
# ============================================================
doc.add_heading("13. Matriz de Responsabilidades", level=1)
para("O resultado depende de um conjunto de obrigações de ambas as partes. A matriz abaixo elimina "
     "dependências implícitas.")
add_table(
    ["Frente", "[Sua Empresa]", "Rocket SF"],
    [
        ["Estratégia e GTM", "Responsável", "Aprova / fornece informações"],
        ["Estrutura de marketing", "Responsável (contrata e gere)", "Aprova materiais"],
        ["Mídia e aquisição", "Responsável (opera)", "Disponibiliza verba de mídia"],
        ["CRM e ferramentas", "Implanta e opera", "Custeia ferramentas / aprova stack"],
        ["Time comercial", "Recruta, treina e gere", "Aprova estrutura / custeia remuneração*"],
        ["Processo de venda", "Conduz até o handoff", "Aprova candidatos e assina contratos"],
        ["Documentação da franquia", "Apoia o fluxo", "Fornece COF e instrumentos jurídicos"],
        ["Informações financeiras/modelo", "Consome para a oferta", "Disponibiliza dados da franquia"],
        ["Aprovação de candidatos", "Encaminha qualificados", "Disponibilidade para aprovar em SLA"],
        ["Onboarding do franqueado", "Faz o handoff", "Estrutura para receber novos franqueados"],
        ["Governança", "Conduz e reporta", "Participação de sócios/lideranças"],
    ],
    widths=[5.0, 6.5, 6.0],
    first_col_bold=True,
    caption="* A definição de custeio da remuneração do time é consolidada na Seção 14.",
)
doc.add_heading("13.1. Compromissos da Rocket SF", level=2)
bullets([
    "Aprovação tempestiva de materiais e campanhas;",
    "Disponibilização de informações do modelo, oferta e mercado;",
    "Participação dos sócios/lideranças na governança e em reuniões-chave;",
    "Documentação da franquia (COF e instrumentos jurídicos);",
    "Informações financeiras necessárias à construção da oferta;",
    "Disponibilidade para aprovação de candidatos quando necessário;",
    "Verba de mídia;",
    "Custos das ferramentas e do CRM;",
    "Velocidade de resposta compatível com o ritmo da operação;",
    "Estrutura necessária para receber e implantar os novos franqueados.",
])

# ============================================================
# 14. CUSTOS
# ============================================================
doc.add_heading("14. Custos da Operação", level=1)
para("Os custos são apresentados de forma segregada, para que não haja sobreposição entre remuneração do "
     "trabalho, investimento em mídia e custos de estrutura. Os valores ainda não fornecidos ficam como "
     "campos a preencher.")
add_table(
    ["#", "Item de custo", "Natureza", "Quem paga", "Valor"],
    [
        ["1", "Remuneração pelo nosso trabalho", "Fee da operação", "Rocket SF", "R$ ________ [a definir]"],
        ["2", "Investimento em mídia", "Verba de aquisição", "Rocket SF", "≈ R$ 10.000/mês (inicial) → escala"],
        ["3", "Salários/remuneração do time comercial", "SDRs + closer", "[a definir]", "R$ ________ [a definir]"],
        ["4", "Ferramentas e CRM", "Stack comercial", "[a definir]", "R$ ________ [a definir]"],
        ["5", "Copywriter", "Produção", "[a definir]", "R$ ________ [a definir]"],
        ["6", "Designer", "Produção", "[a definir]", "R$ ________ [a definir]"],
        ["7", "Gestor de tráfego", "Aquisição", "[a definir]", "R$ ________ [a definir]"],
        ["8", "Outros fornecedores necessários", "Sob demanda", "[a definir]", "R$ ________ [a definir]"],
    ],
    widths=[0.8, 5.5, 3.2, 3.0, 5.0],
    caption="Os campos [a definir] serão preenchidos com os valores fornecidos pela Rocket SF / [Sua "
            "Empresa]. Nenhum valor é presumido. A única referência numérica assumida é o investimento "
            "inicial de mídia (≈ R$ 10.000/mês), informado pela Rocket SF.",
)

# ============================================================
# 15. PREMISSAS E LIMITES
# ============================================================
doc.add_heading("15. Premissas e Limites", level=1)
para("Assumimos responsabilidade integral pela construção, implantação, gestão e otimização da máquina "
     "comercial. O volume de franquias vendidas, contudo, é um resultado de sistema: depende também de "
     "fatores que estão fora do alcance da operação comercial e que compõem, em conjunto, a competitividade "
     "da oferta no mercado. Reconhecê-los explicitamente é o que torna as metas realistas e a parceria "
     "sustentável.")
add_table(
    ["Fator", "Por que impacta o resultado"],
    [
        ["Competitividade da franquia", "Atratividade do modelo frente a outras oportunidades de investimento"],
        ["Oferta e preço", "Condições comerciais e ticket influenciam diretamente a conversão"],
        ["Capacidade de investimento dos candidatos", "Perfil financeiro do público disponível no período"],
        ["Disponibilidade de territórios", "Regiões abertas para novas unidades"],
        ["Velocidade de aprovação", "Agilidade da Rocket SF em analisar e aprovar candidatos"],
        ["Capacidade operacional da franqueadora", "Capacidade de implantar e sustentar novos franqueados"],
        ["Investimento em mídia", "Volume de topo de funil disponível"],
        ["Condições de mercado", "Cenário econômico e de crédito do período"],
        ["Cumprimento das responsabilidades", "Entrega das obrigações de ambas as partes (Seção 13)"],
    ],
    widths=[5.5, 12.0],
    first_col_bold=True,
)

# ============================================================
# 16. GOVERNANÇA
# ============================================================
doc.add_heading("16. Governança", level=1)
para("A operação é conduzida sob um modelo de governança simples e disciplinado, que mantém a Rocket SF "
     "informada e no controle das decisões estratégicas.")
add_table(
    ["Instância", "Frequência", "Conteúdo"],
    [
        ["Acompanhamento operacional", "Semanal", "Análise do funil, KPIs da semana e plano de ação"],
        ["Dashboard / relatório", "Contínuo + fechamento semanal", "Indicadores de marketing, pré-vendas, vendas e economia"],
        ["Reunião executiva", "Periódica (mensal)", "Revisão do forecasting, resultados e prioridades"],
        ["Decisões de escala", "Sob gatilho", "Aumento de mídia e/ou do time comercial (Seção 17)"],
    ],
    widths=[4.5, 4.5, 8.5],
    first_col_bold=True,
)

# ============================================================
# 17. PLANO DE ESCALA
# ============================================================
doc.add_heading("17. Plano de Escala", level=1)
para("A proposta não termina no lançamento. Depois de encontrar eficiência, a máquina é escalada segundo "
     "uma lógica objetiva:")
para("Validar → Encontrar gargalos → Corrigir → Aumentar conversão → Aumentar mídia → Aumentar capacidade "
     "comercial → Escalar vendas de franquias.", bold=True, color=NAVY)
add_table(
    ["Decisão", "Critério objetivo para acionar"],
    [
        ["Aumentar investimento em mídia", "CAC por franquia dentro do alvo e capacidade comercial disponível "
         "para absorver o volume adicional"],
        ["Contratar +1 pré-venda", "Volume de leads qualificados excede a capacidade de contato dentro do SLA"],
        ["Contratar +1 closer", "Reuniões qualificadas excedem a capacidade de fechamento do closer atual"],
        ["Abrir novo canal de aquisição", "Canal atual próximo da saturação com CAC ainda eficiente"],
        ["Adicionar camada de gestão", "Time comercial atinge tamanho que exige coordenação dedicada"],
    ],
    widths=[5.5, 12.0],
    first_col_bold=True,
    caption="Escala é consequência de eficiência comprovada, nunca aposta. Só se aumenta mídia quando o "
            "funil converte; só se aumenta time quando o gargalo migra para a etapa correspondente.",
)

# ============================================================
# 18. PERGUNTAS-CHAVE
# ============================================================
doc.add_heading("18. Perguntas-chave respondidas", level=1)
para("Síntese das definições que não podem ficar ambíguas.")
add_table(
    ["Pergunta", "Resposta"],
    [
        ["O que exatamente estamos assumindo?", "O Go-to-Market completo da venda de franquias: estratégia, "
         "marketing, mídia, CRM, pré-vendas, vendas, gestão e otimização."],
        ["O que não estamos assumindo?", "A concessão jurídica da franquia, a COF, a aprovação final de "
         "candidatos e a implantação/onboarding do franqueado após a assinatura."],
        ["O que será construído?", "GTM, oferta, jornada, funil, estrutura de marketing, CRM dedicado, time "
         "comercial, scripts e playbook."],
        ["O que será operado continuamente?", "Mídia, pré-vendas, vendas, CRM, gestão do funil e otimização."],
        ["Quem será contratado?", "Copywriter, designer, gestor de tráfego, 2 pré-vendas e 1 closer."],
        ["Quem gerencia essas pessoas?", "[Sua Empresa]."],
        ["Quem paga essas pessoas?", "A definir na Seção 14 (campos abertos)."],
        ["Quem paga mídia?", "A Rocket SF (verba de aquisição, ≈ R$ 10.000/mês inicial)."],
        ["Quem paga ferramentas?", "A definir na Seção 14 (custo de ferramentas/CRM)."],
        ["Quanto tempo leva para montar a operação?", "Aproximadamente [X] semanas (Fases 1–2 do cronograma)."],
        ["Quando as campanhas entram no ar?", "Na Fase 3 (Go Live), por volta das semanas 5–7."],
        ["Quando o time começa a trabalhar?", "Após onboarding na Fase 2, operando na Fase 3."],
        ["Quais são as dependências da Rocket SF?", "As obrigações listadas na Seção 13."],
        ["Quais indicadores serão acompanhados?", "A árvore completa da Seção 11."],
        ["Como saberemos se o GTM está funcionando?", "Pelos indicadores de funil e economia versus "
         "forecasting, revisados semanalmente."],
        ["Em qual momento aumentamos mídia?", "Quando o CAC está no alvo e há capacidade comercial (Seção 17)."],
        ["Em qual momento aumentamos o time?", "Quando o gargalo migra para contato ou fechamento (Seção 17)."],
        ["Como acontece o handoff de uma franquia vendida?", "Passagem estruturada do franqueado assinado "
         "para a implantação/onboarding da Rocket SF."],
        ["Quais entregáveis ficam com a Rocket SF?", "GTM, oferta, jornada, funil, CRM configurado, "
         "playbook, scripts, dashboards e forecasting."],
        ["Como será a governança?", "Acompanhamento semanal + reunião executiva periódica + dashboards "
         "(Seção 16)."],
    ],
    widths=[6.0, 11.5],
    first_col_bold=True,
)

# ============================================================
# 19. PRÓXIMOS PASSOS
# ============================================================
doc.add_heading("19. Próximos Passos", level=1)
bullets([
    "Alinhamento executivo e assinatura de NDA;",
    "Preenchimento das informações pendentes (Seção 20);",
    "Homologação do escopo, do modelo de remuneração e dos custos;",
    "Kickoff da Fase 1 (Estratégia e preparação);",
    "Início da construção da máquina de expansão.",
])

# ============================================================
# 20. PENDÊNCIAS
# ============================================================
doc.add_heading("20. Pontos que ainda precisam ser definidos antes do envio da proposta", level=1)
para("Para transformar este documento em uma proposta comercial pronta para assinatura, é necessário "
     "definir/fornecer:")
add_table(
    ["Tema", "O que falta definir"],
    [
        ["Remuneração da operação", "Modelo (fee fixo, variável, success fee) e valores"],
        ["Custeio do time comercial", "Quem paga salários/variável dos SDRs e do closer"],
        ["Custeio de produção", "Quem paga copywriter, designer e gestor de tráfego"],
        ["Ferramentas e CRM", "Qual stack e quem custeia as licenças"],
        ["Forecasting", "Volumes e taxas de conversão por etapa (Seção 6)"],
        ["Oferta da franquia", "Ticket, taxas, royalties e condições comerciais"],
        ["Territórios", "Regiões disponíveis e regras de exclusividade"],
        ["COF e jurídico", "Status da documentação e instrumentos legais"],
        ["Capacidade de aprovação", "SLA da Rocket SF para analisar/aprovar candidatos"],
        ["Capacidade de implantação", "Quantos franqueados a Rocket SF absorve por mês"],
        ["Duração e vigência", "Prazo do contrato, ciclo mínimo e condições de renovação/saída"],
        ["Metas contratuais", "Se haverá metas pactuadas e como se relacionam ao variável"],
        ["Prazo exato de implantação", "Confirmação do [X] semanas conforme disponibilidade da Rocket SF"],
        ["Dados institucionais", "CNPJ, razão social e responsáveis para o cabeçalho da proposta"],
    ],
    widths=[5.0, 12.5],
    first_col_bold=True,
)

# ---- Rodapé de contato ----
doc.add_paragraph()
tt = doc.add_table(rows=1, cols=1)
tt.style = "Table Grid"
c = tt.rows[0].cells[0]
c.text = ""
p = c.paragraphs[0]
r = p.add_run("Contato")
r.font.bold = True
r.font.color.rgb = WHITE
r.font.size = Pt(11)
p2 = c.add_paragraph()
r2 = p2.add_run("[Sua Empresa] — Parceiro de Go-to-Market & Operação Comercial\n"
                "Responsável: [Nome]  ·  [e-mail]  ·  [telefone]")
r2.font.color.rgb = WHITE
r2.font.size = Pt(9.5)
shade(c, HEADER_FILL)

para("Documento confidencial — preparado para a Rocket Soluções Financeiras (Rocket SF).",
     size=9, italic=True, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER)

import os
os.makedirs("entregaveis", exist_ok=True)
out = "entregaveis/Proposta_Maquina_Vendas_Franquias_Rocket.docx"
doc.save(out)
print("Salvo:", out)
print("Parágrafos:", len(doc.paragraphs), "| Tabelas:", len(doc.tables))
