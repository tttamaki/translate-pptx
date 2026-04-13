# 📑 PPTX English to Japanese Translator (Streamlit App)

This is a simple Streamlit app that translates text inside `.pptx` (PowerPoint) files from **English to Japanese** using Google Translate (via `googletrans`). You can upload a `.pptx` file, get it translated, and download the translated version.

---

## 🚀 Features

- Upload a PowerPoint `.pptx` file.
- Automatically detect and translate text in slide-level batches.
- Split long text and large slide batches by character limits (15k) before translation.
- Reuse a single `Translator` instance per uploaded file.
- Translate to **Japanese** using Google Translate.
- Download the translated `.pptx` file.
- Progress bar to show translation progress.
- Built with **Streamlit**, easy to run locally or deploy.

---

## ⚙️ Installation & Running Locally

### 1. Clone the repository:

```bash
git clone https://github.com/your-username/translate-pptx-streamlit.git
cd translate-pptx-streamlit
```

### 2. Create a virtual environment (recommended):

```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

### 3. Install dependencies:

```bash
pip install -r requirements.txt
```

### 4. Run the app:

```bash
streamlit run app.py
```

---

## 📦 Requirements

- `streamlit`
- `python-pptx`
- `googletrans==4.0.2`

See `requirements.txt` for the full list.

---

## ⚠️ Known Issues / Limitations

- **Googletrans is unofficial**—the app may break if Google changes its backend.
- Google Translate web API has per-request text length limits; very text-heavy slides may require splitting in future improvements.
- Does not handle text embedded inside **images or diagrams**—only text boxes are translated.
- No automatic text resizing—if Japanese translation expands the text size, formatting may need manual adjustments.
- Potential compatibility issues with **Python 3.13+**, as `cgi` module was removed in Python 3.13 which is required by some dependencies like `httpx`.
  **Recommended: Use Python 3.10 or 3.11 for local running.**

---

## 💡 Future Enhancements (Possible)

- Switch to an official translation API (e.g., DeepL, Google Cloud Translate).
- Support for translating PDF files.
- Auto-resizing fonts based on translated text length.

---

## 🤝 Acknowledgements

- [Streamlit](https://streamlit.io/)
- [python-pptx](https://python-pptx.readthedocs.io/en/latest/)
- [googletrans](https://py-googletrans.readthedocs.io/en/latest/)

---

## 📝 License

MIT License. Use freely for personal or educational purposes.
