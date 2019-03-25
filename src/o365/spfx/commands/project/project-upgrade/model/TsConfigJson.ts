export interface TsConfigJson {
  extends?: string;
  compilerOptions?: {
    lib?: string[];
    module?: string;
    moduleResolution?: string;
    outDir?: string;
    skipLibCheck?: boolean;
    typeRoots?: string[];
    types?: string[];
    experimentalDecorators?: boolean;
    inlineSources?: boolean;
    strictNullChecks?: boolean;
    noUnusedLocals?: boolean;
  };
  exclude?: string[];
  include?: string[];
}