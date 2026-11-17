---
name: apresentacoes-design
description: Guia de estilo padrão (identidade Saldus) para apresentações em PowerPoint (.pptx) — capa, divisores de seção, slides de texto, números em cartões, gráficos nativos, tabelas, comparação, passos, frase de impacto e encerramento, com as cores, fontes, logo e o "S" da marca. Use SEMPRE que for criar, formatar ou "deixar bonita" qualquer apresentação, slides, deck, PowerPoint, pitch, reunião de resultados, treinamento ou apresentação institucional da Saldus — mesmo que o usuário não peça um estilo específico — para que toda apresentação saia com a mesma identidade visual das planilhas e documentos.
---

# Design de Apresentações — Padrão Saldus

Toda apresentação da Saldus segue a mesma identidade das skills `planilhas-design` e
`documentos-design`: as mesmas cores, as mesmas fontes, a mesma logo e o mesmo "S". Siga os valores
abaixo em vez de improvisar. Se o usuário pedir algo diferente, atenda o pedido dele.

**Atalho:** em Python, use `scripts/estilo_pptx.py` (seção 6). Ele monta cada tipo de slide pronto.

## 0. Antes de começar

Diga em uma ou duas frases o que vai montar (quantos slides, quais tipos) e pergunte só o que o
pedido não responde, com opções de clicar:
1. **Para quê:** reunião de resultados, proposta para cliente, treinamento, apresentação institucional.
2. **Público e uso:** apresentada ao vivo (pouco texto, muito visual) ou enviada para leitura (mais texto).
3. **Foto na capa:** sem foto (recomendado) ou com foto (peça a imagem).
4. **Dados:** se houver números, peça a planilha ou os valores. **Nunca invente números.**

Não mude nada que o usuário não pediu. Na dúvida, pergunte.

## 1. Identidade

| Token | Hex | Uso nos slides |
|---|---|---|
| `primaria` (petróleo) | `#12424D` | Fundo da capa, divisores, frase e encerramento; títulos; cartão e coluna em destaque; barras do gráfico |
| `destaque` (âmbar) | `#FFB745` | **Destaque do título** e números sobre o petróleo, marcadores (●), números dos passos. **Nunca texto âmbar sobre fundo branco** |
| `creme` | `#E3DCCE` | Subtítulos sobre o petróleo (mesma cor das letras da logo) |
| `primaria-clara` | `#E7EFF0` | Cartões e colunas comuns, caixa de destaque |
| `texto` | `#1E2B2E` | Texto sobre fundo branco |
| `texto-suave` | `#5F7378` | Sobretítulos, legendas, número do slide |
| Gráficos | `#12424D`, `#FFB745`, `#3E8A96`, `#8CC2C9`, `#D9822B`, `#A7B4B7` | Nesta ordem; uma série só = petróleo |

**Fontes:** títulos em **Sora** (negrito); texto em **Poppins**; a parte fina dos títulos grandes em
**Poppins Light**. Todas grátis no Google Fonts. Precisam estar instaladas em quem abrir o arquivo.

**Logo:** `assets/logo-saldus.svg` (mestre, vetor). Os slides usam `assets/logo.png` (letras creme, sobre
o petróleo) e `assets/logo-escura.png` (letras petróleo, sobre o branco), em alta resolução.

## 2. Estrutura ("sanduíche")

- **Escuros (petróleo):** capa, divisores de seção, frase de impacto e encerramento.
- **Claros (branco):** todo slide de conteúdo.
- Ordem sugerida: capa → (divisor) → resumo com os números → seções com 1 mensagem por slide →
  próximos passos → encerramento. Ao vivo: ~1,5 a 2 minutos por slide.

## 3. Regras de cada slide

- **Título grande em dois pesos** (capa, divisor, frase, encerramento): numa frase só, a parte normal
  em **Poppins Light branco** e o destaque em **Sora negrito âmbar**. Marque o destaque com
  `**palavra**`: "Custos de **tecnologia**". Sem marcação, a 1ª palavra fica fina e o resto em destaque.
- **Capa:** título e subtítulo juntos, centralizados na altura do slide; logo no alto à esquerda.
- **Divisor:** número grande âmbar e, logo abaixo, o título com o subtítulo colado nele.
- **Título de conteúdo:** sempre no mesmo lugar (alto, à esquerda), Sora negrito 30 pt, petróleo. Pode
  ter um **sobretítulo** em MAIÚSCULAS cinza acima (o assunto). Sem ponto final. **Sem linha embaixo.**
- **Texto:** Poppins 17–18 pt, alinhado à esquerda (nunca centralizado), no máximo 4–5 itens, cada um
  com até 2 linhas. Marcador ● âmbar. `**negrito**` sai em petróleo.
- **Números:** cartões com o número em Sora 36–44 pt; **só um** cartão em destaque (fundo petróleo,
  número âmbar). O número sempre numa linha só.
- **Gráficos:** sempre **nativos do PowerPoint** (editáveis), sem grade, sem eixo de valores, com o valor
  escrito na barra, em formato BR. Ao lado, 1 a 3 frases dizendo o que o gráfico mostra. Fonte embaixo.
- **Tabelas:** cabeçalho petróleo com texto branco em Sora, zebra, **sem bordas nem grade**, total em
  petróleo claro e negrito, números à direita. Colunas com largura proporcional ao texto. Até ~8 linhas
  por slide. Pode ter um **painel ao lado** (explicação ou número em destaque) e uma nota no pé.
- **Subtítulo do slide:** uma frase curta, cinza, logo abaixo do título, quando o título precisa de
  contexto ("Quatro problemas do modelo antigo — mapeados pelo próprio time").
- **Refazendo uma apresentação existente:** mantenha todo o texto e os números do original (confira
  palavra por palavra no fim). Mude só a forma: escolha o tipo de slide de cada conteúdo. A logo em texto
  ("S A L D U S") vira a logo oficial.
- **Passos:** círculos petróleo numerados em âmbar, ligados por linha fina, título e texto curto embaixo.
- **Ícones:** use quando ajudam a identificar o assunto de cada item, cartão ou coluna. Não use como
  enfeite. Sempre da mesma família (**Lucide**, traço fino, em `assets/icones/`) e sempre **dentro de
  um círculo**, o mesmo motivo dos números dos passos. Em slide branco, use círculo petróleo claro com
  ícone petróleo. Em cartão ou coluna clara, use círculo branco. Em cartão ou coluna em destaque
  (petróleo), use círculo petróleo escuro com ícone âmbar. Use um ícone por item, sem misturar com o
  marcador ●. Se nenhum ícone combinar com o assunto, não use.
- **"S" da marca:** em contorno bem claro, vazando pela direita, nos slides escuros. Só textura.
- **Rodapé:** logo pequena à esquerda e número do slide à direita, em todo slide menos a capa e o
  encerramento. Sem faixa, sem barra colorida.
- **Fotos:** cortadas pelo centro, sem distorcer. Na capa, a foto ocupa a metade direita.

**Nunca:** linha ou barra decorativa sob títulos, faixas coloridas nas bordas, degradê, sombra, 3D,
texto centralizado em parágrafos, slide só com texto corrido, fonte abaixo de 14 pt no corpo, texto
saindo da caixa, âmbar como texto sobre branco.

## 4. Tipos de slide

| Função | Quando usar |
|---|---|
| `capa` | Abertura (com ou sem foto) |
| `divisor` | Início de cada seção ("01", "02"...) |
| `slide_texto` | Ideias em lista + imagem ou frase de destaque à direita |
| `slide_cartoes` | 2 a 4 números principais |
| `slide_grafico` | Colunas, barras ou linhas + comentário |
| `slide_tabela` | Dados em tabela curta |
| `slide_comparacao` | Antes × depois, opção A × B (a direita é a destacada) |
| `slide_passos` | Processo, cronograma, próximos passos (2 a 5) |
| `slide_frase` | Frase de impacto ou citação |
| `slide_imagem` | Print, foto ou mapa grande com legenda |
| `encerramento` | Obrigado + contatos |

Varie os tipos: não use o mesmo layout em vários slides seguidos.

## 5. Revisar o design antes de entregar (OBRIGATÓRIO)

Nunca entregue sem revisar. Design quebrado (texto estourando a caixa, itens um em cima do outro, número
partido em duas linhas) é o erro mais comum e o usuário vê na hora.

1. **Revisão automática:** `salvar()` roda `revisar(prs)` e mostra cada problema como `⚠ REVISAR: …`
   (texto maior que a caixa, caixas de texto sobrepostas, elemento fora do slide). Corrija **todos** os
   avisos e gere de novo. Só entregue com `✓ revisão: nenhum texto estourando nem sobreposto`.
2. **Revisão visual com as fontes reais:** converta para PDF e olhe **todos** os slides, um por um.
   - O LibreOffice precisa ter Sora e Poppins, senão usa uma fonte mais estreita e o teste mente (o
     texto cabe no teste e estoura no PowerPoint). No Mac, copie as fontes para
     `~/Library/Application Support/LibreOffice/4/user/fonts/` (Linux: `~/.config/libreoffice/4/user/fonts/`).
   - `soffice --headless --convert-to pdf deck.pptx` e veja cada página como imagem.
   - Procure: texto saindo da caixa, sobreposição, títulos em alturas diferentes, espaços vazios grandes,
     número quebrado ("R$" numa linha, "1.800" na outra).
3. **Se o usuário apontar um problema, conserte na skill** (no script), não só naquele arquivo, para
   não acontecer de novo.

O script mede cada texto com a fonte real (Poppins e Sora são largas e têm linhas altas, ~1,4× o tamanho)
e ajusta sozinho a altura das linhas com ícone, das colunas, dos cartões, do painel e dos comentários;
se ainda não couber, diminui a fonte até 14 pt. Abaixo disso, **corte texto ou divida o slide**.

## 6. Script Python (`scripts/estilo_pptx.py`)

Requer `pip install python-pptx pillow pymupdf` (o `pymupdf` converte os ícones SVG em imagem).

```python
import sys; sys.path.insert(0, "<pasta da skill>/scripts")
from estilo_pptx import *

prs = nova_apresentacao("Resultados 2026")
capa(prs, "Resultados de **2026**", "Diretoria · dezembro de 2026", rodape_texto="Saldus · Diretoria")
# capa com foto: capa(prs, "...", "...", foto="foto.jpg")
divisor(prs, 1, "Onde **estamos**", "Subtítulo opcional")
slide_cartoes(prs, "O ano em **números**", [
    {"rotulo": "Receita", "valor": "R$ 4,2 mi", "detalhe": "+18% no ano", "destaque": True},
    {"rotulo": "Clientes", "valor": "1.240", "detalhe": "novos no ano"}], sobretitulo="Resumo")
slide_grafico(prs, "Receita por **trimestre**", ["T1", "T2", "T3", "T4"], {"Receita (R$ mil)": [900, 1000, 1100, 1200]},
              tipo="colunas", formato='"R$" #,##0', comentario=["**+33%** do T1 ao T4."], fonte_dados="ERP Saldus.")
slide_tabela(prs, "Título", ["Item", "Valor"], [["A", "R$ 10,00"]], alinhar_direita=["Valor"], total=["Total", "R$ 10,00"])
# tabela + painel ao lado (explicação ou número em destaque) + nota no pé:
slide_tabela(prs, "Título", cab, linhas, painel={"titulo": "Total do mês", "valor": "R$ 12.500",
             "texto": "**1,63% efetivo**\nlinha 2", "destaque": True, "icone": "lock"},
             nota="Frase curta no pé do slide.", largura_tabela=0.7)   # fração da largura para a tabela
# subtítulo sob o título (qualquer slide de conteúdo): subtitulo="Uma frase curta"
# passos com nota no pé: slide_passos(prs, "Título", passos, nota="...")
slide_texto(prs, "Título", ["Item 1", "Item 2"], destaque="A ideia principal.")   # ou imagem="foto.jpg"
slide_texto(prs, "Título", [("mail", "Item com ícone"), ("calendar", "Outro item")])  # linhas ícone + texto
# cartões com ícone: {"rotulo": ..., "valor": ..., "icone": "wallet"} · comparação: ("Antes", [itens], "clock")
# lista de ícones: icones_disponiveis()  (wallet, piggy-bank, chart-column, trending-up, users, handshake,
#   target, calendar, clock, mail, phone, shield-check, lightbulb, rocket, tractor, sprout, landmark...)
slide_comparacao(prs, "Antes × **depois**", ("Antes", ["..."]), ("Depois", ["..."]))
slide_passos(prs, "Próximos **passos**", [("Passo", "Texto curto."), ("Passo", "Texto curto.")])
slide_frase(prs, "Uma frase de **impacto**", "Autor")
slide_imagem(prs, "Título", "print.png", legenda="Legenda curta")
encerramento(prs, "Obrigado", ["email@saldus.com.br"], "Dúvidas e sugestões")
notas(slide, "Texto para o apresentador")   # anotações (cada função devolve o slide)
salvar(prs, "resultados.pptx")
```

- `python scripts/estilo_pptx.py exemplo.pptx` gera uma apresentação de exemplo.
- **Modelo pronto:** `assets/modelos/modelo-apresentacao.pptx` traz um slide de cada tipo com texto de
  instrução, para quem monta direto no PowerPoint (`--modelo saida.pptx` gera de novo).
- O tema do arquivo já vem com as cores e as fontes da marca: caixas e gráficos novos criados no
  PowerPoint já saem em Sora/Poppins e nas cores Saldus.
