function getConstNameFromFilePath(filePath) {
  const pos = filePath.indexOf('/src/m365/');
  if (pos < 0) {
    // not a command file
    return;
  }

  // /src/m365/ = 10
  const relativePath = filePath.substr(pos + 10);
  const segments = relativePath.split('/');
  segments.splice(segments.indexOf('commands'), 1);

  const length = segments.length;
  if (length === 2) {
    // remove service from the command file name
    segments[1] = segments[1].replace(`${segments[0]}-`, '');
  }

  const constName = segments.pop()
    .replace('.ts', '')
    .split('-')
    .map(w => w.toUpperCase())
    .join('_');

  return constName;
}

// unfortunately we can't auto-fix this rule because the
// const needs to be changed where it's defined rather than
// where it's used
module.exports = {
  meta: {
    type: 'problem',
    docs: {
      description: 'Incorrect command name',
      suggestion: true
    },
    messages: {
      invalidName: "'{{ actualConstName }}' is not a valid command name. Expected '{{ expectedConstName }}'"
    }
  },
  create: context => {
    return {
      'MethodDefinition[key.name = "name"] MemberExpression > Identifier[name != "commands"]': function (node) {
        const actualConstName = node.name;
        const expectedConstName = getConstNameFromFilePath(context.getFilename());

        if (!expectedConstName) {
          return;
        }

        if (actualConstName !== expectedConstName) {
          context.report({
            node: node,
            messageId: 'invalidName',
            data: {
              actualConstName,
              expectedConstName
            }
          });
        }
      }
    }
  }
};