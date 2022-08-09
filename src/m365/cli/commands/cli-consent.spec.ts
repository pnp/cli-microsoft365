import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../appInsights';
import { Cli, CommandInfo, Logger } from '../../../cli';
import Command from '../../../Command';
import config from '../../../config';
import { sinonUtil } from '../../../utils';
import commands from '../commands';
const command: Command = require('./cli-consent');

describe(commands.CONSENT, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: any;
  let commandInfo: CommandInfo;
  let originalTenant: string;
  let originalAadAppId: string;

  before(() => {
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    originalTenant = config.tenant;
    originalAadAppId = config.cliAadAppId;
    commandInfo = Cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    config.tenant = originalTenant;
    config.cliAadAppId = originalAadAppId;
  });

  after(() => {
    sinonUtil.restore([
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
        assert(loggerLogSpy.calledWith(`To consent permissions for executing yammer commands, navigate in your web browser to https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=31359c7f-bd7e-475c-86db-fdb8c937548e&response_type=code&scope=https%3A%2F%2Fapi.yammer.com%2Fuser_impersonation`));
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
        assert(loggerLogSpy.calledWith(`To consent permissions for executing yammer commands, navigate in your web browser to https://login.microsoftonline.com/fb5cb38f-ecdb-4c6a-a93b-b8cfd56b4a89/oauth2/v2.0/authorize?client_id=2587b55d-a41e-436d-bb1d-6223eb185dd4&response_type=code&scope=https%3A%2F%2Fapi.yammer.com%2Fuser_impersonation`));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports specifying service', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--service') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if specified service is invalid ', async () => {
    const actual = await command.validate({ options: { service: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if service is set to yammer ', async () => {
    const actual = await command.validate({ options: { service: 'yammer' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});