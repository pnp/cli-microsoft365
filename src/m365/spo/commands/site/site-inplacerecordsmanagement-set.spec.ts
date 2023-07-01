import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './site-inplacerecordsmanagement-set.js';

describe(commands.SITE_INPLACERECORDSMANAGEMENT_SET, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SITE_INPLACERECORDSMANAGEMENT_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('correctly handles error when in-place records management already activated', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {

      if ((opts.url as string).indexOf('_api/site/features/add') > -1) {
        throw {
          error: {
            "odata.error": {
              "code": "-1, System.Data.DuplicateNameException",
              "message": {
                "lang": "en-US",
                "value": "Feature 'InPlaceRecords' (ID: da2e115b-07e4-49d9-bb2c-35e93bb9fca9) is already activated at scope 'https://contoso.sharepoint.com/sites/team-a'."
              }
            }
          }
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/team-a', enabled: true } } as any), new CommandError("Feature 'InPlaceRecords' (ID: da2e115b-07e4-49d9-bb2c-35e93bb9fca9) is already activated at scope 'https://contoso.sharepoint.com/sites/team-a'."));
  });

  it('correctly handles error when in-place records management already deactivated', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {

      if ((opts.url as string).indexOf('_api/site/features/remove') > -1) {
        throw {
          error: {
            "odata.error": {
              "code": "-1, System.InvalidOperationException",
              "message": {
                "lang": "en-US",
                "value": "Feature 'da2e115b-07e4-49d9-bb2c-35e93bb9fca9' is not activated at this scope."
              }
            }
          }
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/team-a', enabled: false } } as any), new CommandError("Feature 'da2e115b-07e4-49d9-bb2c-35e93bb9fca9' is not activated at this scope."));
  });

  it('should deactivate in-place records management', async () => {
    const requestStub = sinon.stub(request, 'post').callsFake(async (opts) => {

      if ((opts.url as string).indexOf('_api/site/features/remove') > -1) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, verbose: true, siteUrl: 'https://contoso.sharepoint.com/sites/team-a', enabled: false } });
    assert.strictEqual(requestStub.lastCall.args[0].url, 'https://contoso.sharepoint.com/sites/team-a/_api/site/features/remove');
    assert.strictEqual(requestStub.lastCall.args[0].data.featureId, 'da2e115b-07e4-49d9-bb2c-35e93bb9fca9');
    assert.strictEqual(requestStub.lastCall.args[0].data.force, true);

  });

  it('should activate in-place records management (verbose)', async () => {
    const requestStub = sinon.stub(request, 'post').callsFake(async (opts) => {

      if ((opts.url as string).indexOf('_api/site/features/add') > -1) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, siteUrl: 'https://contoso.sharepoint.com/sites/team-a', enabled: true } });
    assert.strictEqual(requestStub.lastCall.args[0].url, 'https://contoso.sharepoint.com/sites/team-a/_api/site/features/add');
    assert.strictEqual(requestStub.lastCall.args[0].data.featureId, 'da2e115b-07e4-49d9-bb2c-35e93bb9fca9');
    assert.strictEqual(requestStub.lastCall.args[0].data.force, true);
  });

  it('should activate in-place records management', async () => {
    const requestStub = sinon.stub(request, 'post').callsFake(async (opts) => {

      if ((opts.url as string).indexOf('_api/site/features/add') > -1) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/team-a', enabled: true } });
    assert.strictEqual(requestStub.lastCall.args[0].url, 'https://contoso.sharepoint.com/sites/team-a/_api/site/features/add');
    assert.strictEqual(requestStub.lastCall.args[0].data.featureId, 'da2e115b-07e4-49d9-bb2c-35e93bb9fca9');
    assert.strictEqual(requestStub.lastCall.args[0].data.force, true);
  });

  it('supports specifying siteUrl', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--siteUrl') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying enabled', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--enabled') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if siteUrl is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { siteUrl: 'abc', enabled: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the siteUrl is a valid SharePoint URL and enabled set to "true"', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/team-a', enabled: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the siteUrl is a valid SharePoint URL and enabled set to "false"', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/team-a', enabled: false } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});