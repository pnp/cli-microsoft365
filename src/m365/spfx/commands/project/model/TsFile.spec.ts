import * as assert from 'assert';
import * as fs from 'fs';
import * as sinon from 'sinon';
import { TsFile } from ".";
import Utils from '../../../../../Utils';
import * as ts from 'typescript';

describe('TsFile', () => {
  let tsFile: TsFile;

  before(() => {
    tsFile = new TsFile('foo');
  });

  afterEach(() => {
    Utils.restore([
      fs.existsSync,
      ts.createSourceFile
    ]);
    (tsFile as any)._source = undefined;
  });

  it('doesn\'t throw exception if the specified file doesn\'t exist', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    tsFile.source;
    assert(true);
  });

  it('returns undefined source if the specified file doesn\'t exist', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    assert.strictEqual(tsFile.source, undefined);
  });

  it('returns undefined sourceFile if the specified file doesn\'t exist', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    assert.strictEqual(tsFile.sourceFile, undefined);
  });

  it('returns undefined nodes if the specified file doesn\'t exist', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    assert.strictEqual(tsFile.nodes, undefined);
  });

  it('doesn\'t fail when creating TS file fails', () => {
    (tsFile as any)._source = '123';
    sinon.stub(ts, 'createSourceFile').callsFake(() => { throw new Error('An exception has occurred'); })
    assert.strictEqual(tsFile.sourceFile, undefined);
  });
});