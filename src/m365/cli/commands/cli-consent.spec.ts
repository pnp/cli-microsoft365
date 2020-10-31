import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../appInsights';
import { Logger } from '../../../cli';
import Command from '../../../Command';
import config from '../../../config';
import Utils from '../../../Utils';
import commands from '../commands';
const command: Command = require('./cli-consent');

describe(commands.CONSENT, () => {
  let log: any[];
  let logger: Logger;
  let loggerSpy: any;
  let originalTenant: string;
  let originalAadAppId: string;

  before(() => {
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    originalTenant = config.tenant;
    originalAadAppId = config.cliAadAppId;
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: any) => {
        log.push(msg);
      }
    };
    loggerSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    config.tenant = originalTenant;
    config.cliAadAppId = originalAadAppId;
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent
    ]);
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CONSENT), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('shows consent URL for yammer permissions for the default multi-tenant app', (done) => {
    command.action(logger, { options: { service: 'yammer' } }, () => {
      try {
        assert(loggerSpy.calledWith(`To consent permissions for executing yammer commands, navigate in your web browser to https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=31359c7f-bd7e-475c-86db-fdb8c937548e&response_type=code&scope=https%3A%2F%2Fapi.yammer.com%2Fuser_impersonation`));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('shows consent URL for yammer permissions for a custom single-tenant app', (done) => {
    config.tenant = 'fb5cb38f-ecdb-4c6a-a93b-b8cfd56b4a89';
    config.cliAadAppId = '2587b55d-a41e-436d-bb1d-6223eb185dd4';
    command.action(logger, { options: { service: 'yammer' } }, () => {
      try {
        assert(loggerSpy.calledWith(`To consent permissions for executing yammer commands, navigate in your web browser to https://login.microsoftonline.com/fb5cb38f-ecdb-4c6a-a93b-b8cfd56b4a89/oauth2/v2.0/authorize?client_id=2587b55d-a41e-436d-bb1d-6223eb185dd4&response_type=code&scope=https%3A%2F%2Fapi.yammer.com%2Fuser_impersonation`));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports specifying service', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--service') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if specified service is invalid ', () => {
    const actual = command.validate({ options: { service: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if service is set to yammer ', () => {
    const actual = command.validate({ options: { service: 'yammer' } });
    assert.strictEqual(actual, true);
  });
});