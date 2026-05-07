import { visit } from 'unist-util-visit';
import { Node, Parent } from 'unist';
import * as fs from 'fs';
import * as path from 'path';
import * as acorn from 'acorn';

const COMPONENT_NAME = 'CommandPlayer';
const ATTR_RESPONSE = 'response';
const ATTR_COMMANDS = 'commands';
const ATTR_RESPONSE_FROM = 'responseFrom';
const ELEMENT_TYPE = 'mdxJsxAttribute';
const PARSER_OPTIONS = { ecmaVersion: 2024, sourceType: 'module' } as const;

interface VFileLike {
  history?: string[];
  message?: (reason: string, node?: Node) => void;
}

interface MdxJsxAttribute {
  type: 'mdxJsxAttribute';
  name: string;
  value: string | MdxJsxAttributeValueExpression | null;
}

interface MdxJsxAttributeValueExpression {
  type: 'mdxJsxAttributeValueExpression';
  value: string;
  data?: Record<string, unknown>;
}

interface MdxJsxFlowElement extends Node {
  type: 'mdxJsxFlowElement';
  name: string | null;
  attributes: MdxJsxAttribute[];
  children: Node[];
}

interface CodeNode extends Node {
  type: 'code';
  lang?: string;
  value: string;
}

interface HeadingNode extends Node {
  type: 'heading';
  depth: number;
  children: Node[];
}

interface TextNode extends Node {
  type: 'text';
  value: string;
}

const plugin = (): ((root: Node, file?: VFileLike) => void) => {
  return (root: Node, file?: VFileLike): void => {
    const currentFilePath = file?.history?.[0];
    const jsonResponse = findFirstJsonInResponseSection(root);

    visit(root, 'mdxJsxFlowElement', (node: MdxJsxFlowElement) => {
      if (node.name !== COMPONENT_NAME) {
        return;
      }

      const hasResponse = hasAttribute(node, ATTR_RESPONSE);
      const hasCommands = hasAttribute(node, ATTR_COMMANDS);
      const hasResponseFrom = hasAttribute(node, ATTR_RESPONSE_FROM);

      if (!hasResponse && !hasCommands && hasResponseFrom) {
        const responseFromAttr = node.attributes.find(
          (a) => a.type === ELEMENT_TYPE && a.name === ATTR_RESPONSE_FROM
        );
        const refPath = typeof responseFromAttr?.value === 'string' ? responseFromAttr.value : null;

        if (refPath) {
          const json = readResponseFromFile(refPath, currentFilePath);

          if (json) {
            node.attributes = node.attributes.filter(
              (a) => !(a.type === ELEMENT_TYPE && a.name === ATTR_RESPONSE_FROM)
            );
            node.attributes.push({
              type: ELEMENT_TYPE,
              name: ATTR_RESPONSE,
              value: json
            });
          }
          else {
            if (file && typeof file.message === 'function') {
              file.message(`[commandPlayer] Could not resolve responseFrom: ${refPath}`, node);
            }
            else {
              console.warn(`[commandPlayer] Could not resolve responseFrom: ${refPath}`);
            }
          }
        }
      }

      if (!hasResponse && !hasCommands && !hasResponseFrom && jsonResponse) {
        node.attributes.push({
          type: ELEMENT_TYPE,
          name: ATTR_RESPONSE,
          value: jsonResponse
        });
      }

      if (hasCommands) {
        resolveCommandsAttribute(node, currentFilePath);
      }
    });
  };
};

function hasAttribute(node: MdxJsxFlowElement, name: string): boolean {
  return node.attributes.some(
    (attr) => attr.type === ELEMENT_TYPE && attr.name === name
  );
}

function resolveCommandsAttribute(node: MdxJsxFlowElement, currentFilePath?: string): void {
  const attr = node.attributes.find(
    (a) => a.type === ELEMENT_TYPE && a.name === ATTR_COMMANDS
  );

  if (!attr?.value || typeof attr.value !== 'object' || attr.value.type !== 'mdxJsxAttributeValueExpression') {
    return;
  }

  const expr = attr.value as MdxJsxAttributeValueExpression;
  const resolved = replaceResponseFromReferences(expr.value, currentFilePath);

  if (resolved === expr.value) {
    return;
  }

  expr.value = resolved;

  try {
    const newEstree = acorn.parse(`(${resolved})`, PARSER_OPTIONS);

    if (!expr.data) {
      expr.data = {};
    }

    expr.data.estree = newEstree;
  }
  catch (err) {
    console.warn('[commandPlayer] Failed to re-parse commands expression:', err);
  }
}

function replaceResponseFromReferences(expression: string, currentFilePath?: string): string {
  const regex = /responseFrom:\s*(['"])((?:(?!\1).)*)\1/g;
  let result = expression;

  const matches: Array<{ full: string; refPath: string }> = [];
  let match: RegExpExecArray | null;

  while ((match = regex.exec(expression)) !== null) {
    matches.push({ full: match[0], refPath: match[2] });
  }

  for (const { full, refPath } of matches) {
    const json = readResponseFromFile(refPath, currentFilePath);

    if (json) {
      const escaped = json
        .replace(/\\/g, '\\\\')
        .replace(/'/g, "\\'")
        .replace(/\n/g, '\\n')
        .replace(/\r/g, '');
      result = result.replace(full, `${ATTR_RESPONSE}: '${escaped}'`);
    }
    else {
      console.warn(`[commandPlayer] Could not resolve responseFrom in commands: ${refPath}`);
    }
  }

  return result;
}

function isPathWithinAllowedRoots(filePath: string, roots: string[]): boolean {
  return roots.some(root => {
    const relativePath = path.relative(root, filePath);
    return !relativePath.startsWith('..') && !path.isAbsolute(relativePath);
  });
}

function readResponseFromFile(refPath: string, currentFilePath?: string): string | null {
  if (path.isAbsolute(refPath)) {
    console.warn(`[commandPlayer] Absolute paths are not allowed in responseFrom: ${refPath}`);
    return null;
  }

  const cleanRef = refPath.replace(/^\.\//, '');
  const docsRoot = path.resolve(__dirname, '../../docs');
  const responsesRoot = path.resolve(__dirname, '..', 'components', 'responses');
  const allowedRoots = [docsRoot, responsesRoot];
  const candidates: string[] = [];

  if (currentFilePath) {
    candidates.push(path.resolve(path.dirname(currentFilePath), refPath));
  }

  candidates.push(path.resolve(docsRoot, cleanRef));
  candidates.push(path.resolve(docsRoot, 'cmd', cleanRef));
  candidates.push(path.resolve(responsesRoot, cleanRef));

  for (const candidate of candidates) {
    if (!isPathWithinAllowedRoots(candidate, allowedRoots)) {
      continue;
    }

    let realPath: string;
    try {
      realPath = fs.realpathSync(candidate);
    }
    catch {
      continue;
    }

    if (!isPathWithinAllowedRoots(realPath, allowedRoots)) {
      continue;
    }

    if (realPath.endsWith('.txt')) {
      try {
        return fs.readFileSync(realPath, 'utf-8');
      }
      catch (err) {
        console.warn(`[commandPlayer] Error reading file: ${realPath}`, err);
        return null;
      }
    }

    return extractJsonFromResponseSection(realPath);
  }

  return null;
}

function extractJsonFromResponseSection(filePath: string): string | null {
  try {
    const content = fs.readFileSync(filePath, 'utf-8');
    const lines = content.split(/\r?\n/);
    let inResponse = false;
    let inJsonBlock = false;
    const jsonLines: string[] = [];

    for (const line of lines) {
      if (/^#{2}\s+Response\b/i.test(line)) {
        inResponse = true;
        continue;
      }

      if (inResponse && !inJsonBlock && /^#{1,2}\s+/.test(line) && !/Response/i.test(line)) {
        break;
      }

      if (inResponse && !inJsonBlock && /^\s*```json/.test(line)) {
        inJsonBlock = true;
        continue;
      }

      if (inJsonBlock && /^\s*```\s*$/.test(line)) {
        break;
      }

      if (inJsonBlock) {
        jsonLines.push(line);
      }
    }

    return jsonLines.length > 0 ? jsonLines.join('\n') : null;
  }
  catch (err) {
    console.warn(`[commandPlayer] Error reading file: ${filePath}`, err);
    return null;
  }
}

function findFirstJsonInResponseSection(root: Node): string | null {
  const children = (root as Parent).children;

  if (!children) {
    return null;
  }

  let inResponseSection = false;

  for (const child of children) {
    if (isHeading(child)) {
      if (child.depth === 2 && getHeadingText(child).toLowerCase() === 'response') {
        inResponseSection = true;
        continue;
      }

      if (inResponseSection && child.depth <= 2) {
        break;
      }
    }

    if (inResponseSection) {
      const jsonValue = findJsonCodeBlock(child);

      if (jsonValue !== null) {
        return jsonValue;
      }
    }
  }

  return null;
}

function findJsonCodeBlock(node: Node): string | null {
  if (isCodeNode(node) && node.lang === 'json') {
    return node.value;
  }

  const children = (node as Parent).children;

  if (children) {
    for (const child of children) {
      const result = findJsonCodeBlock(child);

      if (result !== null) {
        return result;
      }
    }
  }

  return null;
}

function isHeading(node: Node): node is HeadingNode {
  return node.type === 'heading';
}

function isCodeNode(node: Node): node is CodeNode {
  return node.type === 'code';
}

function getHeadingText(heading: HeadingNode): string {
  return heading.children
    .filter((c): c is TextNode => c.type === 'text')
    .map((c) => c.value)
    .join('');
}

export default plugin;
