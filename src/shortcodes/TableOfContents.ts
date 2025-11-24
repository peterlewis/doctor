import { ShortcodeContext, ShortcodeRender } from "@models";

export const TableOfContentsRenderer: ShortcodeRender = {
  render: (
    attrs: { title: string; position: string },
    _markup: string,
    _context?: ShortcodeContext
  ) => {
    if (attrs) {
      return `<div class="doctor__container__toc ${
        attrs.position === "right" ? "doctor__container__toc_right" : ""
      }">
  ${
    attrs.title
      ? `<h2>${attrs.title}</h2>
  
  `
      : ""
  }
  [[toc]]
</div>`;
    }
    return `[[toc]]`;
  },
  beforeMarkdown: true,
};
