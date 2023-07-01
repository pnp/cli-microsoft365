import ts, { CreateSourceFileOptions, ScriptKind, ScriptTarget, SourceFile } from 'typescript';

export const tsUtil = {
  // wrapper needed to avoid the
  // "Descriptor for property createSourceFile is non-configurable and non-writable"
  // error in tests
  createSourceFile: (fileName: string, sourceText: string, languageVersionOrOptions: ScriptTarget | CreateSourceFileOptions, setParentNodes?: boolean, scriptKind?: ScriptKind): SourceFile => ts.createSourceFile(fileName, sourceText, languageVersionOrOptions, setParentNodes, scriptKind)
};