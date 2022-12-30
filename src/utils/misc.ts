export const misc = {
  getEnums(en: any): string[] {
    return Object
      .keys(en)
      .filter(k => isNaN(parseInt(k)));
  }
};