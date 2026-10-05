// Interfaces defining structured PowerPoint Slide elements
export type SlideType = 'TITLE' | 'BULLET_POINTS' | 'TWO_COLUMN' | 'QUOTE';

export interface SlideElement {
  id: string;
  type: SlideType;
  title: string;
  subtitle?: string;
  content: string[];
  notes?: string;
  backgroundColor?: string;
}

export interface Presentation {
  meta: {
    topic: string;
    generatedAt: string;
    totalSlides: number;
    author: string;
  };
  slides: SlideElement[];
}

/**
 * Advanced AI Response Parser for PowerPoint Add-ins
 * Parses raw text outputs from LLM APIs (OpenAI/Claude) into validated Office.js presentation objects.
 */
export class PowerPointAIParser {
  private defaultThemeColor: string;

  constructor(defaultThemeColor: string = '#0078D4') {
    this.defaultThemeColor = defaultThemeColor;
  }

  /**
   * Converts markdown-style LLM output into a structured presentation model
   */
  public parseMarkdownToPresentation(topic: string, rawText: string): Presentation {
    if (!rawText || rawText.trim().length === 0) {
      throw new Error("Invalid AI payload: Raw text content is empty.");
    }

    const rawBlocks = rawText.split(/(?=^#\s+|^##\s+)/m);
    const slides: SlideElement[] = [];

    rawBlocks.forEach((block, index) => {
      const trimmed = block.trim();
      if (!trimmed) return;

      const lines = trimmed.split('\n').filter(l => l.trim().length > 0);
      const headerLine = lines[0] || '';
      
      const cleanTitle = headerLine.replace(/^#+\s*/, '').trim();
      const bodyLines = lines.slice(1);

      const content: string[] = [];
      let subtitle: string | undefined = undefined;

      bodyLines.forEach(line => {
        const lineText = line.trim();
        if (lineText.startsWith('-') || lineText.startsWith('*') || lineText.startsWith('•')) {
          content.push


