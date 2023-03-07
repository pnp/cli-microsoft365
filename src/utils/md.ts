import * as fs from 'fs';
import { EOL } from 'os';
import * as path from 'path';

function convertTitle(md: string): string {
  return md.replace(/^#\s+(.*)/gm, (match, title: string) => {
    return '\x1b[1m' + title.toLocaleUpperCase() + EOL + Array(title.length + 1).join('=') + '\x1b[0m';
  });
}

function convertHeadings(md: string): string {
  return md.replace(/^(#+)\s+(.*)/gm, (match, level, content: string) => {
    return `${EOL}\x1b[1m${content.toLocaleUpperCase()}\x1b[0m`;
  });
}

function convertAdmonitions(md: string): string {
  const regex = new RegExp('^!!!\\s(.*)' + EOL + '\\s+', 'gm');
  return md.replace(regex, (match, content: string) => {
    return content.toLocaleUpperCase() + EOL + EOL;
  });
}

function includeContent(md: string, rootFolder: string): string {
  return md.replace(/^--8<-- "([^"]+)"/gm, (match, filePath: string) => {
    return fs.readFileSync(path.join(rootFolder, filePath), 'utf8');
  });
}

function convertDd(md: string): string {
  return md.replace(/^:\s(.*)/gm, (match, content: string) => {
    return `  ${content}`;
  });
}

function convertHyperlinks(md: string): string {
  return md.replace(/(?!\[1m)(?!\[0m)\[([^\]]+)\]\(([^\)]+)\)/gm, (match, label: string, url: string) => {
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
  const regex = new RegExp('^=== "(.+?)"(?:\r\n|\n){2}((?:^    (?:.*?(?:\r\n|\n))?)+)', 'gms');
  return md.replace(regex, (match, title: string, content: string) => {
    return `  ${title}${EOL}${EOL}${content.replace(/^    /gms, '')}`;
  });
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
  convertContentTabs,
  convertCodeFences,
  removeInlineMarkup,
  removeTooManyEmptyLines
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