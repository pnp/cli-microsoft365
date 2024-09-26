export const misc = {
  deepClone(obj: any): any {
    if (obj === null || typeof obj !== 'object') {
      return obj;
    }

    if (Array.isArray(obj)) {
      return obj.map(item => misc.deepClone(item));
    }

    const clonedObj: any = {};
    for (const key in obj) {
      clonedObj[key] = misc.deepClone(obj[key]);
    }

    return clonedObj;
  }
};