import * as assert from 'assert';
import { fsUtil } from './fsUtil';

describe('utils/fsUtil', () => {
  it('should get safe filename when file\'name.txt', () => {
    const result = fsUtil.getSafeFileName('file\'name.txt');
    assert.strictEqual(result, 'file\'\'name.txt');
  });
});