import io
import zipfile
from datetime import date

import streamlit as st
from PIL import Image

import catalog_generator as cg

st.set_page_config(page_title="Catálogo Monterosa", page_icon="🌵", layout="centered")

LOGO_PATH = "logo.png"

st.markdown(
    """
    <style>
    .stApp { background-color: #fdfcf9; }
    h1, h2, h3 { color: #867F45; }
    </style>
    """,
    unsafe_allow_html=True,
)

col_logo, col_title = st.columns([1, 3])
with col_logo:
    st.image(LOGO_PATH, width=140)
with col_title:
    st.title("Catálogo de Suculentas")
    st.caption("Gera o catálogo semanal de disponibilidade em PDF, pronto a enviar.")

st.divider()

mode = st.radio(
    "Como queres enviar as variedades desta semana?",
    ["Ficheiro Excel (nome na coluna B, foto na coluna C)", "Fotos individuais (nome vem do ficheiro)"],
    index=0,
)

lang_label = st.radio("Idioma do catálogo", ["Português", "English"], horizontal=True)
lang = "pt" if lang_label == "Português" else "en"

items = None

if mode.startswith("Ficheiro Excel"):
    xlsx_file = st.file_uploader("Carrega o ficheiro .xlsx", type=["xlsx"])
    if xlsx_file is not None:
        try:
            items = cg.extract_items_from_xlsx(xlsx_file.read())
            st.success(f"{len(items)} variedades encontradas no Excel.")
        except Exception as e:
            st.error(f"Não consegui ler o ficheiro: {e}")

else:
    st.caption(
        "Nomeia cada foto com a variedade (e opcionalmente o tamanho do vaso). "
        "Exemplos: `Aloe_vera_P14.jpg`, `Echeveria-Blue-Prince-14cm.png`"
    )
    photo_files = st.file_uploader(
        "Carrega as fotos (podes selecionar várias de uma vez, ou um .zip)",
        type=["jpg", "jpeg", "png", "zip"],
        accept_multiple_files=True,
    )
    if photo_files:
        raw = []
        for f in photo_files:
            if f.name.lower().endswith(".zip"):
                with zipfile.ZipFile(io.BytesIO(f.read())) as z:
                    for name in z.namelist():
                        if name.lower().endswith((".jpg", ".jpeg", ".png")) and not name.startswith("__MACOSX"):
                            raw.append((name.split("/")[-1], z.read(name)))
            else:
                raw.append((f.name, f.read()))
        try:
            items = cg.build_items_from_photos(raw)
            st.success(f"{len(items)} fotos carregadas.")
        except Exception as e:
            st.error(f"Não consegui processar as fotos: {e}")

if items:
    st.subheader("Confirma os nomes antes de gerar")
    st.caption("Podes corrigir qualquer nome ou tamanho de vaso diretamente na tabela.")

    edited_rows = []
    for i, it in enumerate(items):
        c1, c2, c3 = st.columns([1, 3, 1])
        with c1:
            st.image(it.image, width=70)
        with c2:
            new_name = st.text_input(f"name_{i}", value=it.name, label_visibility="collapsed")
        with c3:
            new_pot = st.text_input(f"pot_{i}", value=it.pot_cm or "", label_visibility="collapsed", placeholder="cm")
        edited_rows.append((new_name, new_pot))

    st.divider()

    if st.button("🌵 Gerar catálogo em PDF", type="primary", use_container_width=True):
        final_items = []
        for it, (name, pot) in zip(items, edited_rows):
            final_items.append(cg.CatalogItem(image=it.image, name=name, pot_cm=pot or None))

        logo = Image.open(LOGO_PATH).convert("RGBA")
        pdf_bytes = cg.generate_catalog_pdf(final_items, logo, lang=lang)

        week = date.today().isocalendar()[1]
        year = date.today().isocalendar()[0]
        filename = f"Monterosa_Catalog_{lang.upper()}_W{week:02d}_{year}.pdf"

        st.success("Catálogo gerado!")
        st.download_button(
            "⬇️ Descarregar PDF",
            data=pdf_bytes,
            file_name=filename,
            mime="application/pdf",
            use_container_width=True,
        )
else:
    st.info("Carrega um ficheiro para começar.")
