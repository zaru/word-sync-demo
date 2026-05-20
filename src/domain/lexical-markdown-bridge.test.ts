import { describe, expect, it } from 'vitest';

import { lexicalStateToMarkdown, markdownToLexicalState } from './lexical-markdown-bridge';

describe('LexicalMarkdownBridge', () => {
  it('preserves 基本Markdown要素 through the editor projection', async () => {
    const markdown = [
      '# 見出し',
      '',
      '段落 with **strong** and *emphasis*, [link](https://example.com), and `inline code`.',
      '',
      '- unordered item',
      '- second item',
      '',
      '1. ordered item',
      '2. second ordered item',
      '',
      '> 引用',
      '',
      '```ts',
      'const value = "code block";',
      '```',
    ].join('\n');

    const editorState = await markdownToLexicalState(markdown);
    const projectedMarkdown = await lexicalStateToMarkdown(editorState);

    expect(projectedMarkdown).toContain('# 見出し');
    expect(projectedMarkdown).toContain('**strong**');
    expect(projectedMarkdown).toContain('*emphasis*');
    expect(projectedMarkdown).toContain('[link](https://example.com)');
    expect(projectedMarkdown).toContain('`inline code`');
    expect(projectedMarkdown).toContain('- unordered item');
    expect(projectedMarkdown).toContain('1. ordered item');
    expect(projectedMarkdown).toContain('> 引用');
    expect(projectedMarkdown).toContain('```ts\nconst value = "code block";\n```');
  });
});
