import * as sinon from 'sinon';
import * as assert from 'assert';
import Utils from './Utils';
import * as fs from 'fs';

describe('appInsights', () => {
  afterEach(() => {
    Utils.restore(fs.existsSync);
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
})