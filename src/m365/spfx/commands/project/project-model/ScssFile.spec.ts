import assert from 'assert';
import fs from 'fs';
import { sinonUtil } from '../../../../../utils/sinonUtil.js';
import { ScssFile } from './ScssFile.js';

describe('ScssFile', () => {
  afterEach(() => {
    sinonUtil.restore([
      fs.readFileSync
    ]);
  });

  it('doesn\'t fail when reading file contents fails', () => {
    const scssFile = new ScssFile('file.scss');
    assert.strictEqual(scssFile.source, undefined);
  });
});
