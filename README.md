# Catálogo de Suculentas — Viveiros Monterosa

App web para gerar o catálogo semanal de disponibilidade em PDF (3 colunas × 4 linhas por
página, com fotos completas — sem cortes — e nomes das variedades).

## O que faz

1. Carregas o Excel semanal (mesmo formato de sempre: nome na coluna B, foto na coluna C)
   **ou** carregas fotos individuais, com o nome da variedade no nome do próprio ficheiro
   (ex: `Aloe_vera_P14.jpg`).
2. Confirmas/corriges os nomes e tamanhos de vaso numa tabela simples.
3. Escolhes o idioma (Português / English).
4. Geras e descarregas o PDF, pronto a enviar.

## Correr localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

Abre em `http://localhost:8501`.

## Publicar no Streamlit Community Cloud (grátis, como o projeto de vasilhame)

1. Cria um repositório no GitHub (ex: `catalogo-suculentas-monterosa`) e envia estes ficheiros:
   - `app.py`
   - `catalog_generator.py`
   - `requirements.txt`
   - `logo.png`
   - pasta `fonts/` (com os `.ttf` — necessários para o PDF ficar bem formatado)
2. Em [share.streamlit.io](https://share.streamlit.io), "New app" → escolhe o repositório →
   ficheiro principal `app.py` → Deploy.
3. Fica com um link fixo (ex: `catalogo-suculentas-monterosa.streamlit.app`) que podes abrir
   no telemóvel ou computador sempre que precisares de gerar o catálogo da semana.

## Estrutura dos ficheiros

```
catalog_app/
├── app.py                 # interface (Streamlit)
├── catalog_generator.py   # lógica de geração do PDF (reutilizável)
├── requirements.txt
├── logo.png
└── fonts/                 # tipos de letra usados no PDF (licença OFL)
```

## Notas

- Se carregares o Excel "à moda antiga", a app reconstrói automaticamente a ordem
  correta foto↔nome a partir da posição real de cada imagem no ficheiro (evita o problema
  de nomes trocados que aconteceu numa versão anterior).
- Se carregares fotos individuais, o nome da variedade e o tamanho do vaso são lidos
  diretamente do nome do ficheiro (`_` e `-` viram espaços; `P14` ou `14cm` no fim vira o
  badge do tamanho do vaso).
