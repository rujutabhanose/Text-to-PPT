
# 📊 PPT Deck – Auto-Generate a Presentation from Text

**Your Text, Your Style – Turn bulk text or markdown into a polished PowerPoint presentation.**

PPT Deck is a lightweight web app that lets anyone paste long-form text (markdown, prose, notes, reports) and instantly convert it into a styled, ready-to-present PowerPoint deck. Simply upload your own template, add optional guidance, and supply your preferred LLM API key — the app will handle the rest.

---

## ✨ Features

* 📝 **Input Options**: Paste large chunks of text, markdown, or prose.
* 🎯 **Guidance**: Add a one-line instruction for tone or structure (e.g., “make it an investor pitch deck”).
* 🔑 **Bring Your Own API Key**: Supports OpenAI, Anthropic, Gemini, and more (keys are never stored or logged).
* 🎨 **Template Reuse**: Upload your `.pptx` or `.potx` template to apply colors, fonts, and layouts.
* 🖼️ **Image Reuse**: Recycles existing images from your uploaded template.
* 📥 **Instant Download**: Outputs a new `.pptx` file you can download directly.
* ⚡ **Smart Splitting**: Automatically divides your text into a reasonable number of slides.
* 🔒 **Privacy First**: No logging or saving of user text, keys, or files.

---

## 🚀 Quick Start

### 1. Clone the repo

```bash
git clone https://github.com/23f1000805/tds-bonus-project-Auto-PPT-Generator-GyaanSetu-Deck.git
cd tds-bonus-project-Auto-PPT-Generator-PPT-Deck
```

### 2. Install dependencies

```bash
# Backend (Python FastAPI + pptx libraries)
pip install -r requirements.txt

# Frontend (served as static HTML/JS/CSS)
# no build step needed
```

### 3. Run locally

```bash
uvicorn app:app --reload
```

Visit: [http://localhost:8000](http://localhost:8000)

### 4. Deploy (Railway/Render/Vercel/Heroku)

This app works out-of-the-box on cloud platforms. Just connect your repo and deploy.

---

## 🖥️ Usage

1. Paste your text or markdown.
2. (Optional) Add a one-line guidance like *"make it a research summary"*.
3. Paste your LLM API key (OpenAI, Anthropic, Gemini).
4. Upload a `.pptx` or `.potx` template.
5. Click **Generate** → Get your styled PowerPoint deck!

---

## 🛠️ Architecture

* **Frontend**:

  * Responsive HTML + Tailwind UI
  * Handles text input, template upload, and file download
  * Provides toasts, progress feedback, and history of past generations

* **Backend (FastAPI)**:

  * Accepts text, guidance, API key, and template
  * Splits input intelligently into slide sections
  * Maps new content onto the uploaded template’s style, layout, fonts, and images
  * Generates `.pptx` output using `python-pptx`

---

## 🌟 Optional Enhancements

* Auto-generate speaker notes
* Offer prebuilt guidance templates (*sales deck, investor pitch, research summary*)
* Add live slide previews before download
* Retry logic + better error handling for unstable APIs

---

## 📄 License

MIT License – free to use, modify, and share.

