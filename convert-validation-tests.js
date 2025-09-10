import fs from 'fs';

const files = [
  'src/m365/entra/commands/app/app-role-add.spec.ts',
  'src/m365/entra/commands/app/app-role-list.spec.ts'
];

files.forEach(filePath => {
  if (!fs.existsSync(filePath)) {
    console.log(`File not found: ${filePath}`);
    return;
  }
  
  let content = fs.readFileSync(filePath, 'utf-8');

  // Replace validation test patterns
  content = content.replace(
    /const actual = await command\.validate\(\{ options: \{([^}]*)\} \}, commandInfo\);\s*assert\.notStrictEqual\(actual, true\);/g,
    'const schema = command.getRefinedSchema(commandOptionsSchema as any);\n    const actual = schema?.safeParse({$1});\n    assert.strictEqual(actual?.success, false);'
  );

  content = content.replace(
    /const actual = await command\.validate\(\{ options: \{([^}]*)\} \}, commandInfo\);\s*assert\.strictEqual\(actual, true\);/g,
    'const schema = command.getRefinedSchema(commandOptionsSchema as any);\n    const actual = schema?.safeParse({$1});\n    assert.strictEqual(actual?.success, true);'
  );

  fs.writeFileSync(filePath, content);
  console.log(`Conversion completed for ${filePath}`);
});
