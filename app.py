import asyncio
import copy
import io
import time
from datetime import datetime
from pathlib import Path

import streamlit as st
from googletrans import Translator
from pptx import Presentation

MAX_SINGLE_TEXT_CHARS = 15000
MAX_BATCH_TOTAL_CHARS = 15000


def split_text_by_char_limit(text, char_limit=MAX_SINGLE_TEXT_CHARS):
    if len(text) <= char_limit:
        return [text]

    chunks = []
    start = 0
    while start < len(text):
        end = min(start + char_limit, len(text))
        if end < len(text):
            split_at = text.rfind("\n", start, end)
            if split_at <= start:
                split_at = text.rfind(" ", start, end)
            if split_at <= start:
                split_at = end
            else:
                split_at += 1
        else:
            split_at = end
        chunks.append(text[start:split_at])
        start = split_at
    return chunks


def build_char_limited_batches(segments, char_limit=MAX_BATCH_TOTAL_CHARS):
    batches = []
    current = []
    current_chars = 0

    for segment in segments:
        text = segment[2]
        text_len = len(text)
        if current and current_chars + text_len > char_limit:
            batches.append(current)
            current = []
            current_chars = 0
        current.append(segment)
        current_chars += text_len

    if current:
        batches.append(current)

    return batches


def format_seconds_to_hhmmss(seconds):
    total_seconds = max(0, int(seconds))
    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    secs = total_seconds % 60
    return f"{hours:02d}:{minutes:02d}:{secs:02d}"


class SlideProgressTracker:
    def __init__(self, total_slides):
        self.total_slides = total_slides
        self._progress = st.progress(0)
        self._start_time = time.time()
        self._details = st.empty()

    def is_empty(self):
        return self.total_slides == 0

    def show_empty(self):
        self._progress.progress(1.0, text="slide 0 / 0 | ETA 00:00:00")

    def update(self, completed_slides):
        elapsed = time.time() - self._start_time
        avg_per_slide = elapsed / completed_slides if completed_slides else 0
        remaining_slides = self.total_slides - completed_slides
        eta_seconds = remaining_slides * avg_per_slide
        progress_text = (
            f"slide {completed_slides} / {self.total_slides} "
            f"| ETA {format_seconds_to_hhmmss(eta_seconds)}"
        )
        self._progress.progress(completed_slides / self.total_slides, text=progress_text)

    def finish(self):
        self._progress.empty()

    def update_translation_preview(self, slide_number, translation_pairs, max_items=10):
        if not translation_pairs:
            self._details.code(f"Slide {slide_number}: no translatable text", language="text")
            return

        lines = [f"Slide {slide_number} translation pairs (showing first {max_items})"]
        for idx, (before, after) in enumerate(translation_pairs[:max_items], start=1):
            lines.append(f"[{idx}] {before} --> {after}")
        self._details.code("\n".join(lines).rstrip(), language="text")


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


async def translate_single_text(translator, text, target_lang='ja', source_lang='auto', retries=3):
    for attempt in range(retries):
        try:
            result = await translator.translate(text, src=source_lang, dest=target_lang)
            return result.text
        except Exception:
            if attempt == retries - 1:
                return text
            await asyncio.sleep(1)
    return text


def extract_translatable_paragraphs_by_slide(prs):
    slide_paragraphs = []
    for slide in prs.slides:
        current_slide_paragraphs = []
        for shape in slide.shapes:
            text_frame = getattr(shape, "text_frame", None)
            if text_frame is not None:
                for paragraph in text_frame.paragraphs:
                    txt = paragraph.text.strip()
                    if len(txt) > 0 and not txt.isdigit() and not txt.isascii():
                        current_slide_paragraphs.append(paragraph)
        slide_paragraphs.append(current_slide_paragraphs)
    return slide_paragraphs


async def translate_paragraphs_in_slide(
    translator,
    paragraphs,
    target_lang='ja',
    source_lang='auto',
    para_concurrency=1,
):
    if not paragraphs:
        return []

    original_texts = [paragraph.text for paragraph in paragraphs]
    translated_texts = [paragraph.text for paragraph in paragraphs]

    concurrency = max(1, int(para_concurrency))
    semaphore = asyncio.Semaphore(concurrency)

    async def _translate_single_paragraph(paragraph_idx, paragraph_text):
        if not paragraph_text:
            return paragraph_idx, paragraph_text

        async with semaphore:
            segments = []
            translated_parts = {}
            parts = split_text_by_char_limit(paragraph_text)
            translated_parts[paragraph_idx] = [""] * len(parts)
            for part_idx, part in enumerate(parts):
                segments.append((paragraph_idx, part_idx, part))

            batches = build_char_limited_batches(segments)
            for batch in batches:
                batch_texts = [segment[2] for segment in batch]
                try:
                    batch_translated = await translate_slide_texts(
                        translator,
                        batch_texts,
                        target_lang=target_lang,
                        source_lang=source_lang,
                    )
                except Exception:
                    batch_translated = []
                    for text in batch_texts:
                        translated = await translate_single_text(
                            translator,
                            text,
                            target_lang=target_lang,
                            source_lang=source_lang,
                        )
                        batch_translated.append(translated)

                for segment, translated in zip(batch, batch_translated):
                    _, part_idx, _ = segment
                    translated_parts[paragraph_idx][part_idx] = translated

            return paragraph_idx, "".join(translated_parts[paragraph_idx])

    tasks = [
        _translate_single_paragraph(paragraph_idx, paragraph.text)
        for paragraph_idx, paragraph in enumerate(paragraphs)
    ]
    translated_results = await asyncio.gather(*tasks)
    for paragraph_idx, translated_text in translated_results:
        translated_texts[paragraph_idx] = translated_text

    for paragraph_idx, translated_text in enumerate(translated_texts):
        paragraphs[paragraph_idx].text = translated_text

    return [(original_texts[i], paragraphs[i].text) for i in range(len(paragraphs))]


async def translate_pptx_standard_async(
    input_pptx_file,
    target_lang='ja',
    source_lang='auto',
    preview_limit=10,
    para_concurrency=1,
):
    prs = Presentation(input_pptx_file)
    new_prs = copy.deepcopy(prs)

    slide_paragraphs = extract_translatable_paragraphs_by_slide(new_prs)

    total_slides = len(slide_paragraphs)
    progress_tracker = SlideProgressTracker(total_slides)

    if progress_tracker.is_empty():
        progress_tracker.show_empty()
        progress_tracker.update_translation_preview(0, [], max_items=preview_limit)
        progress_tracker.finish()
        output = io.BytesIO()
        new_prs.save(output)
        output.seek(0)
        return output

    async with Translator() as translator:  # type: ignore[attr-defined]
        for idx, paragraphs in enumerate(slide_paragraphs):
            translation_pairs = await translate_paragraphs_in_slide(
                translator,
                paragraphs,
                target_lang=target_lang,
                source_lang=source_lang,
                para_concurrency=para_concurrency,
            )

            progress_tracker.update(idx + 1)
            progress_tracker.update_translation_preview(
                idx + 1,
                translation_pairs,
                max_items=preview_limit,
            )

    progress_tracker.finish()
    output = io.BytesIO()
    new_prs.save(output)
    output.seek(0)
    return output


def translate_pptx_standard(
    input_pptx_file,
    target_lang='ja',
    source_lang='auto',
    preview_limit=10,
    para_concurrency=1,
):
    return asyncio.run(
        translate_pptx_standard_async(
            input_pptx_file,
            target_lang=target_lang,
            source_lang=source_lang,
            preview_limit=preview_limit,
            para_concurrency=para_concurrency,
        )
    )


async def _translate_one_slide_async(prs_bytes, slide_idx, target_lang, source_lang, para_concurrency=1):
    prs = Presentation(io.BytesIO(prs_bytes))
    slide_paragraphs = extract_translatable_paragraphs_by_slide(prs)
    paragraphs = slide_paragraphs[slide_idx] if slide_idx < len(slide_paragraphs) else []
    pairs = []
    if paragraphs:
        async with Translator() as translator:  # type: ignore[attr-defined]
            pairs = await translate_paragraphs_in_slide(
                translator,
                paragraphs,
                target_lang=target_lang,
                source_lang=source_lang,
                para_concurrency=para_concurrency,
            )
    out = io.BytesIO()
    prs.save(out)
    return out.getvalue(), pairs


def translate_one_slide(prs_bytes, slide_idx, target_lang, source_lang, para_concurrency=1):
    return asyncio.run(
        _translate_one_slide_async(
            prs_bytes,
            slide_idx,
            target_lang,
            source_lang,
            para_concurrency=para_concurrency,
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


def render_translate_mode():
    st.header("PPTX Translation")
    source_lang = st.selectbox(
        "Source language:",
        ['auto', 'en', 'ja', 'es', 'fr', 'de', 'zh', 'ko'],
        format_func=get_language_name,
    )
    target_lang = st.selectbox(
        "Target language:",
        ['ja', 'en', 'es', 'fr', 'de', 'zh', 'ko'],
        format_func=get_language_name,
        index=0,
    )
    uploaded = st.file_uploader("Upload PPTX for Translation", type=["pptx"], key="upload-translate")
    preview_limit = st.number_input(
        "Displayed translation pairs per slide",
        min_value=1,
        max_value=100,
        value=10,
        step=1,
    )
    para_concurrency = st.selectbox(
        "Paragraph translation concurrency",
        options=[1, 2, 4, 8],
        index=0,
        help="Higher values can be faster but may hit API rate limits.",
    )

    for key, default in [
        ('translating', False),
        ('stop_requested', False),
        ('translate_output_bytes', None),
        ('translate_current_slide', 0),
        ('translate_total_slides', 0),
        ('translate_start_time', 0.0),
        ('translate_target_lang', 'ja'),
        ('translate_source_lang', 'auto'),
        ('translate_preview_limit', 10),
        ('translate_para_concurrency', 1),
        ('translate_last_pairs', []),
        ('translate_result', None),
        ('translate_filename', ''),
    ]:
        if key not in st.session_state:
            st.session_state[key] = default

    if uploaded:
        st.success(f"File uploaded: {uploaded.name}")
        if source_lang == target_lang:
            st.error("Source and target languages cannot be the same!")
        else:
            if not st.session_state.translating:
                if st.button("🚀 Translate", key="run-translate"):
                    input_bytes = uploaded.getvalue()
                    prs = Presentation(io.BytesIO(input_bytes))
                    new_prs = copy.deepcopy(prs)
                    buf = io.BytesIO()
                    new_prs.save(buf)
                    uploaded_path = Path(uploaded.name)
                    st.session_state.translate_output_bytes = buf.getvalue()
                    st.session_state.translate_current_slide = 0
                    st.session_state.translate_total_slides = len(list(prs.slides))
                    st.session_state.translate_start_time = time.time()
                    st.session_state.translate_target_lang = target_lang
                    st.session_state.translate_source_lang = source_lang
                    st.session_state.translate_preview_limit = int(preview_limit)
                    st.session_state.translate_para_concurrency = int(para_concurrency)
                    st.session_state.translate_last_pairs = []
                    st.session_state.translate_result = None
                    st.session_state.translate_filename = (
                        f"{uploaded_path.stem}_{target_lang}{uploaded_path.suffix}"
                    )
                    st.session_state.translating = True
                    st.session_state.stop_requested = False
                    st.rerun()
            else:
                if st.button("Translating... (press to stop)", key="run-translate"):
                    st.session_state.stop_requested = True

                idx = st.session_state.translate_current_slide
                total = st.session_state.translate_total_slides
                elapsed = time.time() - st.session_state.translate_start_time
                if total > 0:
                    avg = elapsed / idx if idx > 0 else 0
                    eta = (total - idx) * avg
                    st.progress(
                        idx / total,
                        text=f"slide {idx} / {total} | ETA {format_seconds_to_hhmmss(eta)}",
                    )

                pairs = st.session_state.translate_last_pairs
                lim = st.session_state.translate_preview_limit
                if idx > 0:
                    if pairs:
                        lines = [f"Slide {idx} translation pairs (showing first {lim})"]
                        for i, (before, after) in enumerate(pairs, start=1):
                            lines.append(f"[{i}] {before} --> {after}")
                        st.code("\n".join(lines).rstrip(), language="text")
                    else:
                        st.code(f"Slide {idx}: no translatable text", language="text")

                if st.session_state.stop_requested or idx >= total:
                    stopped = st.session_state.stop_requested
                    st.session_state.translate_result = {
                        'bytes': st.session_state.translate_output_bytes,
                        'filename': st.session_state.translate_filename,
                        'elapsed': elapsed,
                        'stopped': stopped,
                    }
                    st.session_state.translating = False
                    st.rerun()
                else:
                    out_bytes, pairs = translate_one_slide(
                        st.session_state.translate_output_bytes,
                        idx,
                        st.session_state.translate_target_lang,
                        st.session_state.translate_source_lang,
                        para_concurrency=st.session_state.translate_para_concurrency,
                    )
                    st.session_state.translate_output_bytes = out_bytes
                    st.session_state.translate_last_pairs = pairs[:lim]
                    st.session_state.translate_current_slide = idx + 1
                    st.rerun()

        result = st.session_state.get('translate_result')
        if result:
            if result.get('stopped'):
                st.warning(f"Stopped after {result['elapsed']:.1f} seconds. Partial translation saved.")
            else:
                st.success(f"Translated in {result['elapsed']:.1f} seconds!")
            st.download_button(
                "⬇️ Download Translated PPTX",
                data=result['bytes'],
                file_name=result['filename'],
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )


def render_merge_mode():
    st.header("Merge Two PPTX Files")
    st.markdown("""
    - **Alternate:** A1, B1, A2, B2, ...
    - **Append:** All slides from first, then all from second
    - All formatting and images are preserved (to the extent python-pptx allows)
    """)
    merge_type = st.radio(
        "Merge style",
        options=["Alternate (A1, B1, ...)", "Append (A1, A2, ..., B1, B2, ...)"],
        index=0,
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
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )


def main():
    st.set_page_config(page_title="PPTX Tools", page_icon="🌐", layout="wide")
    st.title("🌐 PowerPoint PPTX Tools")

    mode = st.sidebar.radio(
        "Operation Mode:",
        options=["Translate", "Merge"],
        help="Choose whether to translate a PPTX file or merge two presentations."
    )

    if mode == "Translate":
        render_translate_mode()

    elif mode == "Merge":
        render_merge_mode()

    st.markdown("---")
    st.caption("Built with ❤️ using Streamlit and python-pptx.")


if __name__ == "__main__":
    main()
