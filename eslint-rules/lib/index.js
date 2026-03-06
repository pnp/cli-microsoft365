import correctCommandClassName from './rules/correct-command-class-name.js';
import correctCommandName from './rules/correct-command-name.js';
import noByServerRelativeUrlUsage from './rules/no-by-server-relative-url-usage.js';

export const rules = {
  'correct-command-class-name': correctCommandClassName,
  'correct-command-name': correctCommandName,
  'no-by-server-relative-url-usage': noByServerRelativeUrlUsage
};