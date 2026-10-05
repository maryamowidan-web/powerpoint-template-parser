# PowerPoint Template Parser & Helper

A production-ready **TypeScript** engine designed for AI-powered PowerPoint Add-ins. It accepts raw markdown responses from LLM APIs (OpenAI, Claude) and parses them into structured, strongly-typed JSON data objects optimized for **Office.js** presentation building.

## Key Features
- **Strongly Typed**: Built with TypeScript interfaces (`Presentation`, `SlideElement`, `SlideType`).
- **Dynamic Layout Assignment**: Automatically categorizes slide layouts (`TITLE`, `BULLET_POINTS`, `TWO_COLUMN`) based on content density.
- **Robust Parsing**: Cleans markdown headings (`#`, `##`) and bullet symbols (`-`, `*`, `•`) seamlessly.
- **Error Handling**: Validates incoming payload and prevents empty slide generation.
- **Zero External Dependencies**: Pure TypeScript core ready for immediate deployment.

## Technical Architecture
```text
Raw LLM Text Payload ➔ PowerPointAIParser ➔ Structured JSON Model ➔ Office.js API Render

