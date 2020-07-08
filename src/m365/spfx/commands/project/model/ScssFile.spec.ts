import * as assert from 'assert';
import * as fs from 'fs';
import Utils from '../../../../../Utils';
import { ScssFile } from './ScssFile';

describe('ScssFile', () => {
  afterEach(() => {
    Utils.restore([
      fs.readFileSync
    ]);
  });

  it('doesn\'t fail when reading file contents fails', () => {
    const scssFile = new ScssFile('file.scss');
    assert.strictEqual(scssFile.source, undefined);
  });
});