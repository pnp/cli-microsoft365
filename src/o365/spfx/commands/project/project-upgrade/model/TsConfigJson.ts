export interface TsConfigJson {
  compilerOptions: {
    lib?: string[];
    module?: string;
    moduleResolution?: string;
    skipLibCheck?: boolean;
    typeRoots?: string[];
    types?: string[];
  };
}