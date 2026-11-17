# apresentacoes-design

Padrão de apresentações PowerPoint da Saldus: mesmas cores (`#12424D` petróleo, `#FFB745` âmbar),
fontes (**Sora**, **Poppins**, **Poppins Light**), logo e "S" da marca das skills
[planilhas-design](https://github.com/JuniorCarlini/planilhas-design) e
[documentos-design](https://github.com/JuniorCarlini/documentos-design).

## Arquivos

| Arquivo | Para quê |
|---|---|
| `SKILL.md` | O guia completo (funciona como skill e como instrução para qualquer IA) |
| `scripts/estilo_pptx.py` | Gera apresentações `.pptx` no padrão (Python + python-pptx) |
| `assets/modelos/modelo-apresentacao.pptx` | Modelo com um slide de cada tipo, para montar direto no PowerPoint |
| `assets/logo-saldus.svg` / `logo-saldus-escura.svg` | Logo em vetor (mestre) |
| `assets/logo.png` / `logo-escura.png` | Logo em PNG de alta resolução, usada nos slides |
| `assets/simbolo-s-contorno.svg` | Contorno do "S" usado como textura nos slides escuros |
| `assets/icones/` | Ícones Lucide (traço fino, licença ISC em `LICENSE-lucide.txt`) usados nos slides |

## Como usar

**Claude Code:** copie a pasta para `~/.claude/skills/apresentacoes-design/`.
**Claude.ai:** compacte a pasta em `.zip` e envie em *Configurações → Capacidades → Skills*.
**ChatGPT / Gemini:** cole o `SKILL.md` nas instruções e anexe `scripts/estilo_pptx.py` e os arquivos de `assets/` (inclusive `icones/`).
**Sem IA:** abra `assets/modelos/modelo-apresentacao.pptx`, duplique os slides que precisar e troque os textos.

```bash
pip install python-pptx pillow pymupdf
python scripts/estilo_pptx.py exemplo.pptx          # apresentação de exemplo
python scripts/estilo_pptx.py --modelo modelo.pptx  # modelo em branco
```
Instale Sora, Poppins e Poppins Light (grátis em fonts.google.com) para o PowerPoint mostrar as fontes certas.
