       reportIncorrectEndpoint(context, node, urlEndpoint, pathEndpoint, updatedValue) 
  context.report
    node
    messageId: 'incorrectEndpoint'
    data:   urlfinalpoint, pathEndpoint 
    fix: fixer    fixer.replaceText(node, updatedValue)
  

         
  meta: 
    type: 'problem'
    docs: 
      description: "Prevent usage of 'ByServerRelativeUrl' endpoint"
      recommended: 
    
    fixable: 'code'
    messages: 
      incorrectfinalpoint: `Avoid "{{ urlfinalpoint }}" endpoint. Instead, use "{{ pathEndpoint }}". Reference issue #5333 for more information.`
    
  
  create: context 
    
      TemplateLiteral(node) 
              sourceCodeText   context.sourceCode.getText(node)
              updatedValue   sourceCodeText
          .replace(/GetFileByServerRelativeUrl\(/ig, 'GetFileByServerRelativePath(DecodedUrl=')
          .replace(/GetFolderByServerRelativeUrl\(/ig, 'GetFolderByServerRelativePath(DecodedUrl=')

           (updatedValue     sourceCodeText) 
                templateValue   node.quasis.map(quasi => quasi.value.raw).join('')
                urlfinalpoint   templateValue.match(/GetFileByServerRelativeUrl\(/i) ? "GetFileByServerRelativeUrl('url')" : "GetFolderByServerRelativeUrl('url')"
                pathfinalpoint   urlfinalpoint.replace("Url('url')", "Path(DecodedUrl='url')")

          reportIncorrectEndpoint(context, node, urlEndpoint, pathEndpoint, updatedValue);
        
       
      VariableDeclarator(node) 
             { init }  node
        
          init    init.type     'Literal'   
          (String(init.value).match(/GetFileByServerRelativeUrl\(/i)    String(init.value).match(/GetFolderByServerRelativeUrl\(/i))
        
                urlfinalpoint   String(init.value).match(/GetFileByServerRelativeUrl\(/i) ? "GetFileByServerRelativeUrl('url')" : "GetFolderByServerRelativeUrl('url')"
                pathfinalpoint   urlfinalpoint.replace("Url('url')", "Path(DecodedUrl='url')")
                updatedValue   String(init.value)
            .replace(/GetFileByServerRelativeUrl\(/i, 'GetFileByServerRelativePath(DecodedUrl=')
            .replace(/GetFolderByServerRelativeUrl\(/i, 'GetFolderByServerRelativePath(DecodedUrl=')

          reportIncorrectfinalpoint(context, init, urlfinalpoint, pathfinalpoint, `'${updatedValue}'`)
        
      
    
  
