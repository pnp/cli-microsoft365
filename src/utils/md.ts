import chalk = require('chalk');
import * as fs from 'fs';
import { EOL } from 'os';
import * as path from 'path';

function convertTitle(md: string): string {
  return md.replace(/^#\s+(.*)/gm, (match, title: string) => {
    return chalk.bold(title.toLocaleUpperCase()) + EOL + Array(title.length + 1).join('=');
  });
}

function convertHeadings(md: string): string {
  return md.replace(/^(#+)\s+(.*)/gm, (match, level, content: string) => {
    return `${EOL}${chalk.bold(content.toLocaleUpperCase())}`;
  });
}

function convertAdmonitions(md: string): string {
  const regex = new RegExp(/^:::(\w+)([\s\S]*?):::$/, 'gm');
  return md.replace(regex, (_, label: string, content: string) => label.toLocaleUpperCase() + EOL + EOL + content.trim());
}

function includeContent(md: string, rootFolder: string): string {
  const mdxImports = [
    { tag: "<Global />", location: "docs/cmd/_global.mdx" },
    { tag: "<CLISettings />", location: "docs/_clisettings.mdx" }
  ];

  mdxImports.forEach(mdxImport => {
    md = md.replace(mdxImport.tag, () =>
      fs.readFileSync(path.join(rootFolder, mdxImport.location), 'utf8')
    ).replace(/(```\r\n)\r\n(```md definition-list\r\n)/g, "$1$2");
  });

  return md;
}

function convertDd(md: string): string {
  return md.replace(/^:\s(.*)/gm, (match, content: string) => {
    return `  ${content}`;
  });
}

function convertHyperlinks(md: string): string {
  return md.replace(/(?!\[1m)(?!\[22m)\[([^\]]+)\]\(([^\)]+)\)/gm, (match, label: string, url: string) => {
    // if the link is the same as the content, return just the link
    if (label === url) {
      return url;
    }

    // if the link is relative, remove it because there's no way to open it
    // from the terminal anyway. In the future, we could convert it to the
    // actual link of the docs.
    if (!url.startsWith('http:') && !url.startsWith('https:')) {
      return label;
    }

    return `${label} (${url})`;
  });
}

function convertContentTabs(md: string): string {
  return md
    .replace(/<TabItem value="([^"]+)">/gm, '$1')
    .replace(/.*<\/?(Tabs|TabItem)>.*\n?/g, '')
    .replace(/```(?:\w+)?\s*([\s\S]*?)\s*```/g, '$1')
    .trim();
}

function convertCodeFences(md: string): string {
  const regex = new RegExp('^```.*?(?:\r\n|\n)(.*?)```(?:\r\n|\n)', 'gms');
  return md.replace(regex, (match, code: string) => {
    return `${code.replace(/^(.+)$/gm, '  $1')}${EOL}`;
  });
}

function removeInlineMarkup(md: string): string {
  // from https://stackoverflow.com/a/70064453
  return md.replace(/(?<marks>[`]|\*{1,3}|_{1,3}|~{2})(?<inmarks>.*?)\1/g, '$<inmarks>$<link_text>');
}

function removeTooManyEmptyLines(md: string): string {
  const regex = new RegExp('(' + EOL + '){4,}', 'g');
  return md.replace(regex, Array(4).join(EOL));
}

function removeFrontmatter(md: string): string {
  return md.replace(/^---[\s\S]*?---/gm, '').trim();
}

function removeImports(md: string): string {
  return md.replace(/^import .+;$/gms, '').trim();
}

function escapeMd(mdString: string | undefined): string | undefined {
  if (!mdString) {
    return mdString;
  }

  return mdString.toString()
    .replace(/([_*~`|])/g, '\\$1')
    .replace(/\n/g, '<br>');
}

const convertFunctions = [
  convertTitle,
  convertHeadings,
  convertAdmonitions,
  convertDd,
  convertHyperlinks,
  convertCodeFences,
  convertContentTabs,
  removeInlineMarkup,
  removeTooManyEmptyLines,
  removeFrontmatter,
  removeImports
];

export const md = {
  md2plain(md: string, rootFolderDocs: string): string {
    md = includeContent(md, rootFolderDocs);
    convertFunctions.forEach(convert => {
      md = convert(md);
    });

    return md;
  },
  escapeMd
};