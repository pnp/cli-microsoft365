function getClassNameFromFilePath(filePath, dictionary, capitalized) {
  const pos = filePath.indexOf('/src/m365/');
  if (pos < 0) {
    // not a command file
    return;
  }

  // /src/m365/ = 10
  const relativePath = filePath.substr(pos + 10);
  let segments = relativePath.split('/');
  segments.splice(segments.indexOf('commands'), 1);

  // remove command prefix
  const length = segments.length;
  if (length > 1) {
    const commandPrefix = segments[length - 2];
    segments[length - 1] = segments[length - 1].replace(`${commandPrefix}-`, '');
  }

  // replace last element of array with split words
  segments.push(...segments.pop().replace('.ts', '').split('-'));

  const words = segments
    .map(s => breakWords(s, dictionary))
    .flat()
    .map(w => capitalizeWord(w, capitalized))

  const commandName = [
    ...words,
    'Command'
  ].join('');

  return commandName;
}

function capitalizeWord(word, capitalized) {
  const capitalizedWord = capitalized.find(c => c.toLowerCase() === word);
  if (capitalizedWord) {
    return capitalizedWord;
  }

  return word.substr(0, 1).toUpperCase() + word.substr(1).toLowerCase();
}

function breakWords(longWord, dictionary) { 
  const words = [];
  for (let i = 0; i < dictionary.length; i++) {
    if (longWord.indexOf(dictionary[i]) === 0) {
      words.push(dictionary[i]);
      longWord = longWord.replace(dictionary[i], '');
      i = -1;
    }
  }

  if (longWord) {
    words.push(longWord);
  }

  return words;
}

module.exports = {
  // exported for testing
  getClassNameFromFilePath: getClassNameFromFilePath,
  breakWords: breakWords,
  meta: {
    type: 'problem',
    docs: {
      description: 'Incorrect command class name',
      suggestion: true
    },
    fixable: 'code',
    messages: {
      invalidName: "'{{ actualClassName }}' is not a valid command class name. Expected '{{ expectedClassName }}'"
    }
  },
  create: context => {
    return {
      'ClassDeclaration': function (node) {
        if (node.abstract) {
          // command classes are not abstract
          return;
        }

        if (!node.superClass) {
          // class doesn't inherit from another class
          return;
        }

        if (node.superClass.name.indexOf('Command') < 0) {
          // class doesn't inherit from a command class
          return;
        }

        const expectedClassName = getClassNameFromFilePath(context.getFilename(), context.options[0], context.options[1]);
        if (!expectedClassName) {
          return;
        }

        const actualClassName = node.id.name;

        if (actualClassName !== expectedClassName) {
          context.report({
            node: node.id,
            messageId: 'invalidName',
            data: {
              actualClassName,
              expectedClassName
            },
            fix: fixer => fixer.replaceText(node.id, expectedClassName)
          });
        }
      }
    }
  }
};