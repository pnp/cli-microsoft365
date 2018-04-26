import * as sinon from 'sinon';
import * as assert from 'assert';
import Utils from './Utils';
import * as fs from 'fs';
import { getUserId } from './appInsights';

describe('appInsights', () => {
  afterEach(() => {
    Utils.restore([
      fs.existsSync,
      fs.readFileSync,
      fs.writeFileSync
    ]);
    delete require.cache[require.resolve('./appInsights')];
  });

  it('adds -dev label to version logged in the telemetry when CLI ran locally', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    const i: any = require('./appInsights');
    assert(i.default.commonProperties.version.indexOf('-dev') > -1);
  });

  it('doesn\'t add -dev label to version logged in the telemetry when CLI installed from npm', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const i: any = require('./appInsights');
    assert(i.default.commonProperties.version.indexOf('-dev') === -1);
  });

  it('reads the previously stored user id', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake((path, encoding) => 'abc');
    assert.equal(getUserId(), 'abc');
  });

  it('generates new user id if the user file is empty', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => '');
    assert.notEqual(getUserId(), '');
  });

  it('generates new user id if the user file doesn\'t exist', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    assert.notEqual(getUserId(), '');
  });

  it('writes newly generated user id, when the user file doesn\'t exist', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const writeFileSyncStub = sinon.stub(fs, 'writeFileSync').callsFake(() => { });
    getUserId();
    assert(writeFileSyncStub.called);
  });

  it('writes newly generated user id, when the user file is empty', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => '');
    const writeFileSyncStub = sinon.stub(fs, 'writeFileSync').callsFake(() => { });
    getUserId();
    assert(writeFileSyncStub.called);
  });

  it('doesn\'t write user id, when one is already stored in the user file', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => 'abc');
    const writeFileSyncStub = sinon.stub(fs, 'writeFileSync').callsFake(() => { });
    getUserId();
    assert(writeFileSyncStub.notCalled);
  });

  it('doesn\'t fail when reading the user file throws an exception', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => { throw new Error; });
    assert.notEqual(getUserId(), '');
  });

  it('doesn\'t fail when writing the user file throws an exception', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    sinon.stub(fs, 'writeFileSync').callsFake(() => { throw new Error; });
    assert.notEqual(getUserId(), '');
  });
})