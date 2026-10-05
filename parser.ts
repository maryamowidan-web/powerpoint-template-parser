// Interface defining the structure of a PowerPoint Slide
export interface SlideContent {
  title: string;
  bulletPoints: string[];
  themeColor?: string;
}

/**
 * Parses raw AI text output into a structured PowerPoint slide object
 */
export function parseAIToSlide(rawText: string): SlideContent {
  const lines = rawText.split('\n').filter(line => line.trim() !== '');
  
  // Extract first line as Title, rest as Bullet Points
  const title = lines[0] ? lines[0].replace(/^#+\s*/, '') : 'Untitled Slide';
  const bulletPoints = lines.slice(1).map(line => line.replace(/^[-*•]\s*/, '').trim());

  return {
    title,
    bulletPoints: bulletPoints.length > 0 ? bulletPoints : ['No content provided.'],
    themeColor: '#0078D4' // Default Microsoft Office Blue
  };
}

// Example usage and simulation
const samplePromptResult = `AI in Modern Presentations
- Automates slide creation from text
- Enhances visual consistency
- Saves time for users`;

const parsedSlide = parseAIToSlide(samplePromptResult);
console.log("Generated Slide Object:", JSON.stringify(parsedSlide, null, 2));

