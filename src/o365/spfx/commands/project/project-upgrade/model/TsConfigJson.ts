export interface TsConfigJson {
  compilerOptions?: {
    lib?: string[];
    module?: string;
    moduleResolution?: string;
    outDir?: string;
    skipLibCheck?: boolean;
    typeRoots?: string[];
    types?: string[];
    experimentalDecorators?: boolean;
  };
  exclude?: string[];
  include?: string[];
}