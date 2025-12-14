export const stringUtil = {
  normalizeLineEndings: (str?: string): string | undefined => str?.replace(/\r\n/g, '\n')
};