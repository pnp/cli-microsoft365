import * as openpackage from 'open';

export const browserUtil = {
  /* c8 ignore next 4 */
  async open(url: string): Promise<void> {
    const runningOnWindows = process.platform === 'win32';
    await openpackage(url, { wait: runningOnWindows });
  }
};