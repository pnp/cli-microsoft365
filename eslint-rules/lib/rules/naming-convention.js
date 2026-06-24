// Simplified naming-convention rule for method camelCase enforcement.
// Replaces @typescript-eslint/naming-convention with selector: ['method'], format: ['camelCase']

export default {
  meta: {
    type: 'suggestion',
    docs: {
      description: 'Enforce camelCase naming convention for methods'
    },
    messages: {
      notCamelCase: "Method name '{{name}}' is not in camelCase."
    }
  },

  create(context) {
    function isCamelCase(name) {
      // Allow names starting with # (private fields)
      const cleaned = name.startsWith('#') ? name.slice(1) : name;
      // Must start with lowercase
      if (cleaned.length > 0 && cleaned[0] !== cleaned[0].toLowerCase()) {
        return false;
      }
      // Must not contain underscores (except leading _)
      const stripped = cleaned.replace(/^_+/, '');
      return !stripped.includes('_');
    }

    return {
      MethodDefinition(node) {
        if (node.computed) return;

        const key = node.key;
        let name;

        if (key.type === 'Identifier') {
          name = key.name;
        }
        else if (key.type === 'PrivateIdentifier') {
          name = '#' + key.name;
        }
        else {
          return;
        }

        if (!isCamelCase(name)) {
          context.report({
            node: key,
            messageId: 'notCamelCase',
            data: { name }
          });
        }
      }
    };
  }
};
