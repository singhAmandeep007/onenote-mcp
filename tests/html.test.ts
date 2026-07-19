import { describe, it, expect } from 'vitest';
import { htmlToText, escapeHtml } from '../src/html.js';

describe('htmlToText', () => {
  it('strips HTML tags', () => {
    expect(htmlToText('<p>Hello <b>World</b></p>')).toBe('Hello World');
  });

  it('converts <br> to newlines', () => {
    expect(htmlToText('line1<br/>line2<br>line3')).toBe('line1\nline2\nline3');
  });

  it('converts block-closing tags to newlines', () => {
    expect(htmlToText('<p>one</p><p>two</p>')).toBe('one\ntwo');
  });

  it('strips <script> and <style> blocks', () => {
    expect(
      htmlToText('<style>body{color:red}</style><p>hi</p><script>alert(1)</script>'),
    ).toBe('hi');
  });

  it('decodes named entities', () => {
    expect(htmlToText('&amp; &lt; &gt; &quot; &apos;')).toBe('& < > " \'');
  });

  it('decodes numeric entities', () => {
    expect(htmlToText('&#65;&#66;&#67;')).toBe('ABC');
  });

  it('collapses whitespace', () => {
    expect(htmlToText('  lots   of   spaces  ')).toBe('lots of spaces');
  });

  it('collapses multiple blank lines', () => {
    expect(htmlToText('<p>a</p><p></p><p></p><p>b</p>')).toMatch(/a\n{1,2}\n?b/);
  });

  it('handles empty input', () => {
    expect(htmlToText('')).toBe('');
  });
});

describe('escapeHtml', () => {
  it('escapes &, <, >, "', () => {
    expect(escapeHtml('a & b < c > d "e"')).toBe('a &amp; b &lt; c &gt; d &quot;e&quot;');
  });

  it('passes through safe strings unchanged', () => {
    expect(escapeHtml('hello world')).toBe('hello world');
  });
});
