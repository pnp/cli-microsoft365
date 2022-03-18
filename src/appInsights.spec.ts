import * as assert from 'assert';
import * as fs from 'fs';
import * as sinon from 'sinon';
import { sinonUtil } from './utils';

const env = Object.assign({}, process.env);

describe('appInsights', () => {

  afterEach(() => {
    sinonUtil.restore(fs.existsSync);
    delete require.cache[require.resolve('./appInsights')];
    process.env = env;
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

  it('sets env logged in the telemetry to \'docker\' when CLI run in CLI docker image', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    process.env.CLIMICROSOFT365_ENV = 'docker';
    const i: any = require('./appInsights');
    assert(i.default.commonProperties.env === 'docker');
  });
});