// Simplified camelcase rule compatible with Oxlint jsPlugins.
// Replicates ESLint's camelcase rule behavior for this project's needs.

export default {
  meta: {
    type: 'suggestion',
    docs: {
      description: 'Enforce camelCase naming convention'
    },
    messages: {
      notCamelCase: "Identifier '{{name}}' is not in camel case."
    },
    schema: [
      {
        type: 'object',
        properties: {
          allow: {
            type: 'array',
            items: { type: 'string' },
            minItems: 0,
            uniqueItems: true
          },
          properties: {
            enum: ['always', 'never']
          },
          ignoreDestructuring: {
            type: 'boolean'
          },
          ignoreImports: {
            type: 'boolean'
          }
        },
        additionalProperties: false
      }
    ]
  },

  create(context) {
    const options = context.options[0] || {};
    const allowPatterns = (options.allow || []).map(p => new RegExp(p, 'u'));
    const checkProperties = options.properties !== 'never';
    const ignoreDestructuring = options.ignoreDestructuring || false;
    const ignoreImports = options.ignoreImports || false;

    const reported = new Set();

    function isUnderscored(name) {
      const stripped = name.replace(/^_+|_+$/gu, '');
      return stripped.includes('_') && stripped !== stripped.toUpperCase();
    }

    function isAllowed(name) {
      return allowPatterns.some(re => name === re.source || re.test(name));
    }

    function isGoodName(name) {
      return !isUnderscored(name) || isAllowed(name);
    }

    function getReportKey(node) {
      if (node.range) return node.range[0];
      if (node.start !== undefined) return node.start;
      return JSON.stringify(node.loc);
    }

    function report(node) {
      const key = getReportKey(node);
      if (reported.has(key)) return;
      reported.add(key);
      context.report({
        node,
        messageId: 'notCamelCase',
        data: { name: node.name }
      });
    }

    function isInDestructuring(node) {
      let current = node.parent;
      while (current) {
        if (current.type === 'ObjectPattern' || current.type === 'ArrayPattern') {
          return true;
        }
        if (current.type === 'Property' && current.parent &&
          (current.parent.type === 'ObjectPattern' || current.parent.type === 'ObjectExpression')) {
          return true;
        }
        if (current.type !== 'Property' && current.type !== 'AssignmentPattern' && current.type !== 'RestElement') {
          break;
        }
        current = current.parent;
      }
      return false;
    }

    function isShorthandProperty(node) {
      return node.parent.type === 'Property' &&
        node.parent.shorthand &&
        node.parent.value === node;
    }

    return {
      Identifier(node) {
        if (isGoodName(node.name)) return;

        const parent = node.parent;

        // Skip call/new expressions (backward compat with ESLint)
        if (parent.type === 'CallExpression' || parent.type === 'NewExpression') {
          return;
        }

        // Skip right side of assignment patterns (default values)
        if (parent.type === 'AssignmentPattern' && parent.right === node) {
          return;
        }

        // Skip TypeScript type contexts
        if (parent.type && (
          parent.type.startsWith('TS') ||
          parent.type === 'TSTypeReference' ||
          parent.type === 'TSTypeAnnotation'
        )) {
          return;
        }

        // Variable declarations
        if (parent.type === 'VariableDeclarator' && parent.id === node) {
          if (ignoreDestructuring && isInDestructuring(node)) return;
          report(node);
          return;
        }

        // Function/class declarations
        if ((parent.type === 'FunctionDeclaration' || parent.type === 'ClassDeclaration') && parent.id === node) {
          report(node);
          return;
        }

        // Function parameters
        if ((parent.type === 'FunctionDeclaration' || parent.type === 'FunctionExpression' ||
          parent.type === 'ArrowFunctionExpression') && parent.params && parent.params.includes(node)) {
          report(node);
          return;
        }

        // Catch clause parameter
        if (parent.type === 'CatchClause' && parent.param === node) {
          report(node);
          return;
        }

        // Object property keys and class members
        if ((parent.type === 'Property' || parent.type === 'MethodDefinition' ||
          parent.type === 'PropertyDefinition') && parent.key === node && !parent.computed) {
          if (!checkProperties) return;
          report(node);
          return;
        }

        // Member expression property (only on assignment targets)
        if (parent.type === 'MemberExpression' && parent.property === node && !parent.computed) {
          if (!checkProperties) return;
          const grandParent = parent.parent;
          if (grandParent &&
            ((grandParent.type === 'AssignmentExpression' && grandParent.left === parent) ||
              (grandParent.type === 'AssignmentPattern' && grandParent.left === parent))) {
            report(node);
          }
          return;
        }

        // Import specifiers
        if (parent.type === 'ImportSpecifier' && parent.local === node) {
          if (ignoreImports) {
            // If the imported name equals the local name, skip
            const importedName = parent.imported.type === 'Identifier' ? parent.imported.name : parent.imported.value;
            if (importedName === node.name) return;
          }
          report(node);
          return;
        }
        if (parent.type === 'ImportDefaultSpecifier') {
          if (ignoreImports) return;
          report(node);
          return;
        }

        // Export specifiers
        if (parent.type === 'ExportSpecifier' && parent.exported === node) {
          report(node);
          return;
        }

        // Labels
        if ((parent.type === 'LabeledStatement' || parent.type === 'BreakStatement' ||
          parent.type === 'ContinueStatement') && parent.label === node) {
          report(node);
          return;
        }

        // Destructuring with shorthand notation
        if (isShorthandProperty(node)) {
          if (ignoreDestructuring && isInDestructuring(node)) return;
          report(node);
          return;
        }

        // Rest element in destructuring
        if (parent.type === 'RestElement' && parent.argument === node) {
          report(node);
          return;
        }

        // Array pattern elements
        if (parent.type === 'ArrayPattern') {
          if (ignoreDestructuring) return;
          report(node);
          return;
        }
      }
    };
  }
};
