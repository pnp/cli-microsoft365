import * as assert from 'assert';
import * as fs from 'fs';
import * as sinon from 'sinon';
import { TsFile } from ".";
import Utils from '../../../../../../Utils';

describe('TsFile', () => {
  let tsFile: TsFile;

  before(() => {
    tsFile = new TsFile('foo');
  });

  afterEach(() => {
    Utils.restore(fs.existsSync);
  });

  it('doesn\'t throw exception if the specified file doesn\'t exist', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    tsFile.source;
    assert(true);
  });

  it('returns undefined source if the specified file doesn\'t exist', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    assert.equal(tsFile.source, undefined);
  });

  it('returns undefined sourceFile if the specified file doesn\'t exist', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    assert.equal(tsFile.sourceFile, undefined);
  });

  it('returns undefined nodes if the specified file doesn\'t exist', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    assert.equal(tsFile.nodes, undefined);
  });
});