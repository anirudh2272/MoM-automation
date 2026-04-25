# 📋 Minutes of Meeting Automation System

A full end-to-end AI pipeline that transcribes meeting audio, 
extracts key information, and exports bilingual formatted minutes.

## 🔗 Live App
👉 https://animomautomation.streamlit.app

## 🚀 Pipeline
Audio → Whisper STT → Speaker Diarization → 
NLP Extraction → Translation → Bilingual DOCX

## ✨ Features
- Transcribes audio in any language (50+ supported)
- Auto-detects source language
- Translates output to any target language
- Extracts action items, decisions, and summary
- Speaker-attributed transcript
- Exports formatted bilingual DOCX

## 🤖 Models Used
- faster-whisper — Speech to Text
- facebook/bart-large-cnn — Summarization
- facebook/bart-large-mnli — Zero-Shot Classification
- facebook/mbart-large-50-many-to-many-mmt — Translation

## 🛠️ Tech Stack
Python · HuggingFace Transformers · Streamlit · python-docx
