import correctCommandClassName from './rules/correct-command-class-name.js';
import correctCommandName from './rules/correct-command-name.js';
import noByServerRelativeUrlUsage from './rules/no-by-server-relative-url-usage.js';
import camelcase from './rules/camelcase.js';
import namingConvention from './rules/naming-convention.js';

const plugin = {
  rules: {
    'correct-command-class-name': correctCommandClassName,
    'correct-command-name': correctCommandName,
    'no-by-server-relative-url-usage': noByServerRelativeUrlUsage,
    'camelcase': camelcase,
    'naming-convention': namingConvention
  }
};

export default plugin;
export const rules = plugin.rules;