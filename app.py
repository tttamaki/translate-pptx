import streamlit as st
from pptx import Presentation
import copy
import io
import time
import asyncio
from datetime import datetime
from googletrans import Translator

async def translate_slide_texts(translator, texts, target_lang='ja', source_lang='auto', retries=3):
    for attempt in range(retries):
        try:
            result = await translator.translate(texts, src=source_lang, dest=target_lang)
            if isinstance(result, list):
                return [item.text for item in result]
            return [result.text]
        except Exception:
            if attempt == retries - 1:
                raise
            await asyncio.sleep(1)
    return texts


async def translate_run_fallback(translator, run, target_lang='ja', source_lang='auto', retries=3):
    for attempt in range(retries):
        try:
            result = await translator.translate(run.text, src=source_lang, dest=target_lang)
            run.text = result.text
            return
        except Exception:
            if attempt == retries - 1:
                return
            await asyncio.sleep(1)


async def translate_pptx_standard_async(input_pptx_file, target_lang='ja', source_lang='auto'):
    prs = Presentation(input_pptx_file)
    new_prs = copy.deepcopy(prs)

    slide_runs = []
    for slide in new_prs.slides:
        current_slide_runs = []
        for shape in slide.shapes:
            text_frame = getattr(shape, "text_frame", None)
            if text_frame is not None:
                for paragraph in text_frame.paragraphs:
                    for run in paragraph.runs:
                        txt = run.text.strip()
                        if len(txt) > 0 and not txt.isdigit() and not txt.isascii():
                            current_slide_runs.append(run)
        slide_runs.append(current_slide_runs)

    total_slides = len(slide_runs)
    progress = st.progress(0)

    async with Translator() as translator:  # type: ignore[attr-defined]
        for idx, runs in enumerate(slide_runs):
            if runs:
                texts = [run.text for run in runs]
                try:
                    translated_texts = await translate_slide_texts(
                        translator,
                        texts,
                        target_lang=target_lang,
                        source_lang=source_lang,
                    )
                    for run, translated in zip(runs, translated_texts):
                        run.text = translated
                except Exception:
                    for run in runs:
                        await translate_run_fallback(
                            translator,
                            run,
                            target_lang=target_lang,
                            source_lang=source_lang,
                        )

            progress.progress((idx + 1) / total_slides if total_slides else 1.0)

    progress.empty()
    output = io.BytesIO()
    new_prs.save(output)
    output.seek(0)
    return output

def translate_pptx_standard(input_pptx_file, target_lang='ja', source_lang='auto'):
    return asyncio.run(
        translate_pptx_standard_async(
            input_pptx_file,
            target_lang=target_lang,
            source_lang=source_lang,
        )
    )

def merge_two_presentations(pptx1_file, pptx2_file, alternate=True):
    prs1 = Presentation(pptx1_file)
    pptx2_file.seek(0)
    prs2 = Presentation(pptx2_file)
    out_prs = Presentation()
    # Remove default blank slide
    if out_prs.slides:
        rId = out_prs.slides._sldIdLst[0].rId
        out_prs.part.drop_rel(rId)
        del out_prs.slides._sldIdLst[0]

    slides1 = list(prs1.slides)
    slides2 = list(prs2.slides)
    total = max(len(slides1), len(slides2)) if alternate else len(slides1) + len(slides2)
    progress = st.progress(0)

    def add_slide_from(src_slide):
        layout = out_prs.slide_layouts[0]
        new_slide = out_prs.slides.add_slide(layout)
        new_slide._element.clear()
        new_slide._element.append(copy.deepcopy(src_slide._element.cSld))

    if alternate:
        for idx in range(total):
            if idx < len(slides1):
                add_slide_from(slides1[idx])
            if idx < len(slides2):
                add_slide_from(slides2[idx])
            progress.progress((idx + 1) / total)
    else:
        for idx, s in enumerate(slides1):
            add_slide_from(s)
            progress.progress((idx + 1) / total)
        for idx, s in enumerate(slides2):
            add_slide_from(s)
            progress.progress((len(slides1) + idx + 1) / total)
    progress.empty()
    output = io.BytesIO()
    out_prs.save(output)
    output.seek(0)
    return output

def get_language_name(lang_code: str) -> str:
    languages = {
        'en': 'English',
        'ja': 'Japanese',
        'es': 'Spanish',
        'fr': 'French',
        'de': 'German',
        'zh': 'Chinese',
        'ko': 'Korean',
        'auto': 'Auto-detect'
    }
    return languages.get(lang_code, lang_code)

# ---- Streamlit App ----
st.set_page_config(page_title="PPTX Tools", page_icon="🌐", layout="wide")
st.title("🌐 PowerPoint PPTX Tools")

mode = st.sidebar.radio(
    "Operation Mode:",
    options=["Translate", "Merge"],
    help="Choose whether to translate a PPTX file or merge two presentations."
)

if mode == "Translate":
    st.header("PPTX Translation")
    source_lang = st.selectbox("Source language:",
        ['auto', 'en', 'ja', 'es', 'fr', 'de', 'zh', 'ko'], format_func=get_language_name)
    target_lang = st.selectbox("Target language:",
        ['ja', 'en', 'es', 'fr', 'de', 'zh', 'ko'], format_func=get_language_name, index=0)
    uploaded = st.file_uploader("Upload PPTX for Translation", type=["pptx"], key="upload-translate")
    st.markdown("""
    - Empty and numeric-only strings are skipped.
    - Text is translated in slide-level batches.
    - Images and formatting are preserved.
    """)
    if uploaded:
        st.success(f"File uploaded: {uploaded.name}")
        if source_lang == target_lang:
            st.error("Source and target languages cannot be the same!")
        elif st.button("🚀 Translate"):
            start = time.time()
            with st.spinner("Translating slides..."):
                translated_bytes = translate_pptx_standard(uploaded, target_lang, source_lang)
            st.success(f"Translated in {time.time()-start:.1f} seconds!")
            filename = f"translated_{get_language_name(source_lang)}_{get_language_name(target_lang)}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
            st.download_button(
                "⬇️ Download Translated PPTX",
                data=translated_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

elif mode == "Merge":
    st.header("Merge Two PPTX Files")
    st.markdown("""
    - **Alternate:** A1, B1, A2, B2, ...
    - **Append:** All slides from first, then all from second
    - All formatting and images are preserved (to the extent python-pptx allows)
    """)
    merge_type = st.radio(
        "Merge style",
        options=["Alternate (A1, B1, ...)", "Append (A1, A2, ..., B1, B2, ...)"],
        index=0
    )
    alternate = merge_type.startswith("Alternate")
    col1, col2 = st.columns(2)
    with col1:
        pptx1 = st.file_uploader("Upload PPTX File 1", type=["pptx"], key="pptx1")
    with col2:
        pptx2 = st.file_uploader("Upload PPTX File 2", type=["pptx"], key="pptx2")
    if pptx1 and pptx2:
        st.success("Both files uploaded!")
        if st.button("🚀 Merge Presentations"):
            start = time.time()
            with st.spinner("Merging presentations..."):
                pptx2.seek(0)
                merged = merge_two_presentations(pptx1, pptx2, alternate=alternate)
            st.success(f"Merged in {time.time()-start:.1f} seconds.")
            st.download_button(
                "⬇️ Download Merged PPTX",
                data=merged,
                file_name=f"merged_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

st.markdown("---")
st.caption("Built with ❤️ using Streamlit and python-pptx.")
