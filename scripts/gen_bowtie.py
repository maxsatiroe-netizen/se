# -*- coding: utf-8 -*-
"""Gera o diagrama Bowtie (SVG nativo, theme-aware) e injeta na proposta."""

VM = [70, 210, 350, 450, 545, 690, 885, 1035, 1175]  # 9 marcos de valor

# geometria do funil
TL, BL = 110, 430      # topo/base nas extremidades
TN, BN = 225, 315      # topo/base no "pescoço" (centro)
xL0, xLn = VM[0], VM[4]   # funil esquerdo: VM1 -> VM5
xR0, xRn = VM[5], VM[8]   # funil direito: VM6 -> VM9

def ltop(x):  return TL + (TN - TL) * (x - xL0) / (xLn - xL0)
def lbot(x):  return BL + (BN - BL) * (x - xL0) / (xLn - xL0)
def rtop(x):  return TN + (TL - TN) * (x - xR0) / (xRn - xR0)
def rbot(x):  return BN + (BL - BN) * (x - xR0) / (xRn - xR0)

stages_left  = [(140, "Conscientização"), (280, "Educação"), (460, "Seleção")]
stages_right = [(787, "Onboarding"), (960, "Retenção"), (1105, "Expansão")]

# medição
Y_SCOPE = 452
Y_DT    = 500
Y_DTLBL = 489
Y_CR    = 536
Y_CRLBL = 525
Y_VMTOP = 468
Y_VMBOT = 552
Y_VMLBL = 568

s = []
s.append('<svg viewBox="0 0 1245 600" role="img" aria-label="Jornada de expansão de franquias no modelo Bowtie: a aquisição afunila candidatos (Conscientização, Educação, Seleção, Priorização) até o Compromisso Mútuo — a assinatura do franqueado — e depois se abre em Onboarding, Retenção e Expansão. Cada etapa tem marco de valor (VM), taxa de conversão (CR) e tempo de passagem (Δt). A operação comercial cobre a aquisição até o Compromisso Mútuo; o restante é da Rocket após o handoff." xmlns="http://www.w3.org/2000/svg">')

# defs
s.append('<defs>')
s.append('<marker id="ah" markerWidth="9" markerHeight="9" refX="7" refY="3" orient="auto"><path d="M0,0 L7,3 L0,6 Z" fill="currentColor"/></marker>')
s.append('<marker id="ahs" markerWidth="9" markerHeight="9" refX="1" refY="3" orient="auto"><path d="M7,0 L0,3 L7,6 Z" fill="currentColor"/></marker>')
s.append('<pattern id="hatch" width="7" height="7" patternTransform="rotate(45)" patternUnits="userSpaceOnUse"><line x1="0" y1="0" x2="0" y2="7" stroke="currentColor" stroke-width="1.1" opacity="0.55"/></pattern>')
s.append('</defs>')

# cabeçalhos de zona
s.append(f'<text x="300" y="66" text-anchor="middle" font-size="15" font-weight="700" letter-spacing="2.5" fill="currentColor">AQUISIÇÃO</text>')
s.append(f'<text x="935" y="66" text-anchor="middle" font-size="15" font-weight="700" letter-spacing="2.5" fill="currentColor">RETENÇÃO E EXPANSÃO</text>')

# funis (contorno)
s.append(f'<path d="M{VM[0]},{TL} L{xLn},{TN}" fill="none" stroke="currentColor" stroke-width="1.6"/>')
s.append(f'<path d="M{VM[0]},{BL} L{xLn},{BN}" fill="none" stroke="currentColor" stroke-width="1.6"/>')
s.append(f'<line x1="{VM[0]}" y1="{TL}" x2="{VM[0]}" y2="{BL}" stroke="currentColor" stroke-width="1.6"/>')
s.append(f'<path d="M{xR0},{TN} L{VM[8]},{TL}" fill="none" stroke="currentColor" stroke-width="1.6"/>')
s.append(f'<path d="M{xR0},{BN} L{VM[8]},{BL}" fill="none" stroke="currentColor" stroke-width="1.6"/>')
s.append(f'<line x1="{VM[8]}" y1="{TL}" x2="{VM[8]}" y2="{BL}" stroke="currentColor" stroke-width="1.6"/>')

# separadores internos (dotted) nas fronteiras de etapa
for x, top, bot in [(VM[1], ltop, lbot), (VM[2], ltop, lbot), (VM[3], ltop, lbot),
                     (VM[6], rtop, rbot), (VM[7], rtop, rbot)]:
    s.append(f'<line x1="{x}" y1="{top(x):.1f}" x2="{x}" y2="{bot(x):.1f}" stroke="currentColor" stroke-width="1.1" stroke-dasharray="1.5 5" opacity="0.6"/>')

# Priorização (faixa hachurada em VM3)
px, pw = VM[2], 13
s.append(f'<rect x="{px-pw}" y="{ltop(px):.1f}" width="{pw*2}" height="{lbot(px)-ltop(px):.1f}" fill="url(#hatch)" stroke="currentColor" stroke-width="1.1"/>')
s.append(f'<line x1="{px}" y1="{ltop(px):.1f}" x2="{px}" y2="152" stroke="currentColor" stroke-width="1.1"/>')
s.append(f'<circle cx="{px}" cy="{ltop(px):.1f}" r="3" fill="currentColor"/>')
s.append(f'<text x="{px}" y="144" text-anchor="middle" font-size="14" fill="currentColor">Priorização</text>')

# rótulos de etapa
for x, label in stages_left + stages_right:
    s.append(f'<text x="{x}" y="276" text-anchor="middle" font-size="17" fill="currentColor">{label}</text>')

# Compromisso Mútuo (quadrado central)
sq_x0, sq_x1, sq_y0, sq_y1 = VM[4], VM[5], 206, 334
s.append(f'<rect x="{sq_x0}" y="{sq_y0}" width="{sq_x1-sq_x0}" height="{sq_y1-sq_y0}" rx="4" style="fill:var(--pine)"/>')
s.append(f'<text x="{(sq_x0+sq_x1)/2:.0f}" y="264" text-anchor="middle" font-size="14.5" font-weight="700" style="fill:var(--brass-bright)">Compromisso</text>')
s.append(f'<text x="{(sq_x0+sq_x1)/2:.0f}" y="284" text-anchor="middle" font-size="14.5" font-weight="700" style="fill:var(--brass-bright)">Mútuo</text>')

# ciclo retenção<->expansão (seta circular)
s.append('<path d="M998,150 a30,30 0 1 1 -0.5,0" fill="none" stroke="currentColor" stroke-width="1.4" marker-end="url(#ah)"/>')

# barras de escopo
s.append(f'<rect x="{VM[0]}" y="{Y_SCOPE}" width="{VM[5]-VM[0]}" height="7" rx="3.5" style="fill:var(--brass)"/>')
s.append(f'<rect x="{VM[5]}" y="{Y_SCOPE}" width="{VM[8]-VM[5]}" height="7" rx="3.5" style="fill:var(--ink-faint)"/>')
s.append(f'<text x="{(VM[0]+VM[5])/2:.0f}" y="{Y_SCOPE-6}" text-anchor="middle" font-size="12.5" font-weight="600" style="fill:var(--brass)">NOSSA OPERAÇÃO · até o Compromisso Mútuo</text>')
s.append(f'<text x="{(VM[5]+VM[8])/2:.0f}" y="{Y_SCOPE-6}" text-anchor="middle" font-size="12.5" font-weight="600" style="fill:var(--ink-faint)">ROCKET · pós-handoff</text>')

# VM linhas + rótulos
for i, x in enumerate(VM, 1):
    s.append(f'<line x1="{x}" y1="{Y_VMTOP}" x2="{x}" y2="{Y_VMBOT}" stroke="currentColor" stroke-width="1.1" stroke-dasharray="1.5 5" opacity="0.65"/>')
    s.append(f'<text x="{x}" y="{Y_VMLBL}" text-anchor="middle" font-size="12" font-weight="600" fill="currentColor">VM{i}</text>')

# Δt (setas simples) e CR (setas duplas) por segmento
for i in range(8):
    x0, x1 = VM[i] + 8, VM[i+1] - 8
    mid = (VM[i] + VM[i+1]) / 2
    s.append(f'<line x1="{x0}" y1="{Y_DT}" x2="{x1}" y2="{Y_DT}" stroke="currentColor" stroke-width="1.3" marker-end="url(#ah)"/>')
    s.append(f'<text x="{mid:.0f}" y="{Y_DTLBL}" text-anchor="middle" font-size="13" fill="currentColor">&#916;t{i+1}</text>')
    s.append(f'<line x1="{x0}" y1="{Y_CR}" x2="{x1}" y2="{Y_CR}" stroke="currentColor" stroke-width="1.3" marker-start="url(#ahs)" marker-end="url(#ah)"/>')
    s.append(f'<text x="{mid:.0f}" y="{Y_CRLBL}" text-anchor="middle" font-size="13" fill="currentColor">CR{i+1}</text>')

s.append('</svg>')
svg = "".join(s)

figure = (
'<figure class="fig reveal">\n' + svg + '\n'
'<div class="fig-legend">'
'<span class="it"><span class="sw" style="background:var(--brass)"></span>Nossa operação comercial (até o handoff)</span>'
'<span class="it"><span class="sw" style="background:var(--ink-faint)"></span>Rocket · pós-handoff</span>'
'<span class="it"><code>VM</code> marco de valor</span>'
'<span class="it"><code>CR</code> taxa de conversão entre marcos</span>'
'<span class="it"><code>&#916;t</code> tempo de passagem entre marcos</span>'
'</div>\n'
'<figcaption>O modelo Bowtie aplicado à expansão de franquias. A operação que assumimos cobre todo o lado '
'da aquisição — da Conscientização ao Compromisso Mútuo (a assinatura). Do handoff em diante, onboarding, '
'retenção e expansão são da Rocket. Medimos cada passagem por VM, CR e &#916;t, o que expõe exatamente onde '
'otimizar.</figcaption>\n'
'</figure>'
)

path = "entregaveis/proposta_rocket.html"
html = open(path, encoding="utf-8").read()
assert "<!--BOWTIE-->" in html, "placeholder ausente"
html = html.replace("<!--BOWTIE-->", figure)
open(path, "w", encoding="utf-8").write(html)
print("Bowtie injetado. Tamanho SVG:", len(svg), "chars")
