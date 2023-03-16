import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./homesite-set');

describe(commands.HOMESITE_SET, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
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
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.HOMESITE_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('sets the specified site as the Home Site', async () => {
    const expected = 'The Home site has been set to https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fWork. It may take some time for the change to apply. Check aka.ms/homesites for details.';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_api/SPO.Tenant/SetSPHSite`) {
        return Promise.resolve(
          {
            "value": expected
          }
        );
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        siteUrl: "https://contoso.sharepoint.com/sites/Work"
      }
    } as any);
    assert(loggerLogSpy.calledWith(expected));
  });

  it('sets the specified site as the Home Site and sets the Viva Connections default experience to True', async () => {
    const expected = 'The Home site has been set to https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fWork and the Viva Connections default experience to True. It may take some time for the change to apply. Check aka.ms/homesites for details.';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_api/SPO.Tenant/SetSPHSiteWithConfiguration`) {
        return Promise.resolve({
          "value": expected
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        siteUrl: "https://contoso.sharepoint.com/sites/Work",
        vivaConnectionsDefaultStart: true
      }
    } as any);
    assert(loggerLogSpy.calledWith(expected));
  });

  it('correctly handles error when setting the Home Site', async () => {
    const expected = "[Error ID: 09149788-0a26-4cee-a333-699b81f629d7] The provided site url can't be set as a Home site. Check aka.ms\u002fhomesites for cmdlet requirements.";
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject({
        error: {
          "odata.error": {
            "code": "-2147213238, Microsoft.SharePoint.SPException",
            "message": {
              "lang": "en-US",
              "value": "[Error ID: 09149788-0a26-4cee-a333-699b81f629d7] The provided site url can't be set as a Home site. Check aka.ms/homesites for cmdlet requirements."
            }
          }
        }
      });
    });

    await assert.rejects(command.action(logger, {
      options: {
        siteUrl: "https://contoso.sharepoint.com/sites/Work"
      }
    } as any), new CommandError(expected));
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'post').callsFake(() => Promise.reject('An error has occurred'));

    await assert.rejects(command.action(logger, {
      options: {
        siteUrl: "https://contoso.sharepoint.com/sites/Work"
      }
    } as any), new CommandError(`An error has occurred`));
  });

  it('fails validation if the siteUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { siteUrl: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the siteUrl option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
