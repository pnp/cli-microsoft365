import { visit } from 'unist-util-visit';
import { fromMarkdown } from 'mdast-util-from-markdown';
import { Node, Parent } from 'unist';

interface ListItem {
  type: string;
  name: string;
  children: Node[];
}

interface DefinitionList {
  type: string;
  name: string;
  attributes: Array<{ type: string; name: string; value: string }>;
  children: ListItem[];
}

/**
 * Turns a "```md definition-list" code block into a definition list
 */
const plugin = (): ((root: Node) => void) => {
  const transformer = (root: Node): void => {
    visit(root, 'code', (node: any, index: number, parent: Parent) => {
      if (!node.meta?.includes('definition-list')) {
        return;
      }

      const { value } = node;
      const listItems: string[] = value.replace(/\r/g, '').split('\n').filter((x: string) => x);

      const items: ListItem[] = listItems.map((listItem: string, i: number) => {
        const tree = fromMarkdown(i % 2 ? listItem.substring(2, listItem.length) : listItem);

        if (i % 2) {
          return {
            type: 'mdxJsxTextElement',
            name: 'dd',
            children: tree.children
          };
        }
        else {
          return {
            type: 'mdxJsxTextElement',
            name: 'dt',
            children: tree.children
          };
        }
      });

      const definitionList: DefinitionList = {
        type: 'mdxJsxTextElement',
        name: 'dl',
        attributes: [
          {
            type: 'mdxJsxAttribute',
            name: 'class',
            value: 'cli-definitionList'
          }
        ],
        children: [...items]
      };

      parent.children[index] = definitionList;
    });
  };

  return transformer;
};

export default plugin;