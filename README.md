
# LLM-Project-PPT-Generator-
# 🎓 Paper2PPT 
“Transform any research paper into a professional PowerPoint presentation — instantly.”
 # Overview

Paper2PPT is an LLM-powered agentic system that automatically fetches research papers from Semantic Scholar, arXiv, or CrossRef, summarizes their sections using Ollama (Llama 3.2), and generates a complete PowerPoint presentation with structured slides, visuals, and AI-crafted bullet points — all inside an elegant Streamlit UI.

This project was developed as part of the Agentic AI training to demonstrate how local LLMs and retrieval APIs can work together in a seamless, user-friendly workflow.

# Key Features

- Multi-Source Paper Retrieval

Automatically fetches metadata and abstracts from Semantic Scholar, arXiv, or CrossRef.

- Local LLM Summarization (Ollama Llama 3.2)

Converts abstracts into concise, well-formatted bullet points for each slide section.

- AI-Generated PowerPoint

Automatically builds slides with Introduction, Objectives, Methodology, Results, and Conclusion using python-pptx.

- Elegant Streamlit Interface

Navy-themed, hover-enhanced design with smooth animations and progress feedback.

- Safety & Guardrails

Filters out unwanted words, ignores repetitive model text, and ensures consistent formatting.

- Smart Fallbacks

When Semantic Scholar fails, it retries via CrossRef or arXiv for robust retrieval.

# Tech Stack
| Layer               | Tools / Technologies                          |
| ------------------- | --------------------------------------------- |
| **Frontend / UI**   | Streamlit (custom CSS, animated gradients)    |
| **Backend / Logic** | Python 3.10+, Requests, Ollama API            |
| **LLM**             | Llama 3.2 (via Ollama local model)            |
| **Data Sources**    | Semantic Scholar API, CrossRef API, arXiv API |
| **File Generation** | `python-pptx`, `Pillow`                       |
| **Safety Layer**    | Custom content filter & sanitization pipeline |

# How It Works

User Input: Enter a paper title, DOI, or Semantic Scholar/arXiv link.

Paper Retrieval: The system fetches metadata and abstract via Semantic Scholar or other APIs.

LLM Summarization: Ollama’s Llama 3.2 model generates slide-ready bullet points.

Presentation Generation: python-pptx constructs a PowerPoint with titles, sections, and formatted content.

Download: The user receives a .pptx file ready for submission or presentation.

# Example Flow
```text
┌──────────────┐
│  User enters │
│ Paper title  │
└──────┬───────┘
       ↓
┌──────────────┐
│ System fetches│
│ Academic APIs │
└──────┬───────┘
       ↓
┌──────────────┐
│  LLM creates │
│ Slide content│
└──────┬───────┘
       ↓
┌──────────────┐
│ Auto-builds  │
│ Presentation │
└──────┬───────┘
       ↓
┌──────────────┐
│User downloads│
│ Final slides │
└──────────────┘
```



# Example Sections

Introduction: Introduce about paper.

Objectives: Explain the goal of that research.

Methodology: Describe about the research paper.

Results: gives the paper results.

Conclusion: conclude the research paper.

![alt text](LLM-generated-ppt-photo-1.png)

# Project Architecture
Streamlit UI  →  Paper Retriever  →  Ollama Summarizer  →  PPTX Generator
                     ↑
           Semantic Scholar / CrossRef / arXiv APIs

# Getting Started
1️⃣ Prerequisites

Python ≥ 3.10

Ollama installed locally (ollama pull llama3.2)

Internet connection for fetching paper metadata

2️⃣ Install Dependencies
pip install -r requirements.txt

3️⃣ Run the App
streamlit run app.py

# Output

A fully formatted PowerPoint (____.pptx)

5 slides generated automatically with consistent academic styling

AI-summarized content ready for presentation 
<img width="1908" height="988" alt="LLM_generated_ppt_photo" src="https://github.com/user-attachments/assets/230ce923-8168-4f47-8faf-5026e66380cf" />

# project video
https://drive.google.com/file/d/13KyDblITpiQ_gzKmM3X8PYohWyZTeQEw/view?usp=drive_link


# Conclusion

Paper2PPT showcases the power of Agentic AI systems — combining retrieval, reasoning, and generation into a unified tool that turns any research paper into a presentation in seconds.
It bridges AI automation and academic productivity, aligning perfectly with the Agentic AI curriculum goals.
