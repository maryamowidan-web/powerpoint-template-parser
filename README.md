# PowerPoint Template Parser & Helper

A lightweight **TypeScript** helper designed for AI PowerPoint Add-ins. It converts unstructured markdown-like responses from LLM APIs into typed, structured slide objects ready for **Office.js** rendering.

## Features
- Strong TypeScript typing for slide layouts.
- Parses titles and list items dynamically.
- Lightweight with zero dependencies.

## Usage
Import `parseAIToSlide` function and pass raw AI text to receive a clean, structured `SlideContent` object.
