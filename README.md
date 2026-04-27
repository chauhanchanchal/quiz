---
title: AI Quiz Platform
emoji: 📝
colorFrom: indigo
colorTo: purple
sdk: docker
app_port: 7860
pinned: false
---

# AI Quiz Platform

A Flask-based AI-powered quiz and exam grading platform.

## Features

- AI Quiz generation (MCQ, Short Answer, Long Answer, Coding)
- Handwritten answer sheet OCR via Google Gemini Vision (with EasyOCR fallback)
- AI-graded exam reports via Groq LLaMA
- Mentor approval workflow
- Student-style answer key generation
- Role-based dashboards

## Tech Stack

- Backend: Flask + Python
- AI: Groq LLaMA 3.3-70B + Google Gemini Vision API
- OCR Fallback: EasyOCR (offline)
- NLP: Sentence-Transformers + spaCy

## Configuration

Required environment variables (set as HF Space Secrets):
- `GEMINI_API_KEY` — from https://aistudio.google.com/apikey
- `GROQ_API_KEY` — from https://console.groq.com/keys

## Project Info

Mini Project — Applied AI, B.Tech Sem II, A.Y. 2025-26
Unitedworld Institute of Technology (UIT), KU
