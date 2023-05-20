export const browserUtil = {
  /* c8 ignore next 5 */
  async open(url: string): Promise<void> {
    const _open = require('open');
    const runningOnWindows = process.platform === 'win32';
    await _open(url, { wait: runningOnWindows });
  }
};