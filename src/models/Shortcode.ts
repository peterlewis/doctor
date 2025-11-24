export interface Shortcode {
  [name: string]: ShortcodeRender;
}

export interface ShortcodeContext {
  currentFilePath?: string;
  frontMatter?: any;
}

export interface ShortcodeRender {
  render: (
    attr: any,
    markup: string,
    context?: ShortcodeContext
  ) => Promise<string> | string;
  beforeMarkdown: boolean;
}
