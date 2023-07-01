export const basic = {
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  asyncFilter: async <T>(arr: any[], predicate: (value: T, index: number, array: T[]) => unknown, thisArg?: any): Promise<T[]> => {
    const results = await Promise.all(arr.map(predicate));

    return arr.filter((_v, index) => results[index]);
  }
};