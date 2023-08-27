function reportIncorrectEndpoint(context, node, urlEndpoint, pathEndpoint, updatedValue) {
  context.report({
    node,
    messageId: 'incorrectEndpoint',
    data: { urlEndpoint, pathEndpoint },
    fix: fixer => fixer.replaceText(node, updatedValue),
  });
}

module.exports = {
  meta: {
    type: 'problem',
    docs: {
      description: "Prevent usage of 'ByServerRelativeUrl' endpoint",
      recommended: true
    },
    fixable: 'code',
    messages: {
      incorrectEndpoint: `Avoid "{{ urlEndpoint }}" endpoint. Instead, use "{{ pathEndpoint }}". Reference issue #5333 for more information.`
    }
  },
  create: context => {
    return {
      TemplateLiteral(node) {
        const sourceCodeText = context.getSourceCode().getText(node);
        const updatedValue = sourceCodeText
          .replace(/GetFileByServerRelativeUrl\(/ig, 'GetFileByServerRelativePath(DecodedUrl=')
          .replace(/GetFolderByServerRelativeUrl\(/ig, 'GetFolderByServerRelativePath(DecodedUrl=');

        if (updatedValue !== sourceCodeText) {
          const templateValue = node.quasis.map(quasi => quasi.value.raw).join('');
          const urlEndpoint = templateValue.match(/GetFileByServerRelativeUrl\(/i) ? "GetFileByServerRelativeUrl('url')" : "GetFolderByServerRelativeUrl('url')";
          const pathEndpoint = urlEndpoint.replace("Url('url')", "Path(DecodedUrl='url')");

          reportIncorrectEndpoint(context, node, urlEndpoint, pathEndpoint, updatedValue);
        }
      },
      VariableDeclarator(node) {
        const { init } = node;
        if (
          init && init.type === 'Literal' &&
          (String(init.value).match(/GetFileByServerRelativeUrl\(/i) || String(init.value).match(/GetFolderByServerRelativeUrl\(/i))
        ) {
          const urlEndpoint = String(init.value).match(/GetFileByServerRelativeUrl\(/i) ? "GetFileByServerRelativeUrl('url')" : "GetFolderByServerRelativeUrl('url')";
          const pathEndpoint = urlEndpoint.replace("Url('url')", "Path(DecodedUrl='url')");
          const updatedValue = String(init.value)
            .replace(/GetFileByServerRelativeUrl\(/i, 'GetFileByServerRelativePath(DecodedUrl=')
            .replace(/GetFolderByServerRelativeUrl\(/i, 'GetFolderByServerRelativePath(DecodedUrl=');

          reportIncorrectEndpoint(context, init, urlEndpoint, pathEndpoint, `'${updatedValue}'`);
        }
      }
    };
  },
};