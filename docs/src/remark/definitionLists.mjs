import visit from 'unist-util-visit'; // "unist-util-visit": "^2.0.3"
import { micromark } from 'micromark'; // "micromark": "^3.1.0"

/**
 * Turns a "```md definition-list" code block into a definition list
 */

export default function plugin() {
  const transformer = (root) => {
    visit(root, 'code', (node, index, parent) => {
      if (!node.meta?.includes('definition-list')) {
        return;
      }

      const {value} = node;
      const listItems = value.split('\n').filter(x => x);
      const items = listItems.map((listItem, i) => {
        const content = micromark(i % 2 ? listItem.substring(2, listItem.length) : listItem);

        if (i % 2) {
          return `<dd>${content}</dd>`;
        } 
        else {
          return `<dt>${content}</dt>`;
        }
      });

      parent.children[index] = {
        type: 'html',
        value: `<dl>${items.join('')}</dl>`
      };
    });
  };

  return transformer;
}