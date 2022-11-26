import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./accesstoken-get');

describe(commands.ACCESSTOKEN_GET, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    auth.service.connected = true;
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
      auth.ensureAccessToken
    ]);
    auth.service.accessTokens = {};
    auth.service.spoUrl = undefined;
  });

  after(() => {
    sinonUtil.restore([
      appInsights.trackEvent,
      pid.getProcessName,
      auth.restoreAuth
    ]);
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.ACCESSTOKEN_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves access token for the specified resource', async () => {
    const d: Date = new Date();
    d.setMinutes(d.getMinutes() + 1);
    auth.service.accessTokens['https://graph.microsoft.com'] = {
      expiresOn: d.toString(),
      accessToken: 'ABC'
    };

    await command.action(logger, { options: { debug: false, resource: 'https://graph.microsoft.com' } });
    assert(loggerLogSpy.calledWith('ABC'));
  });

  it('retrieves access token for SharePoint when sharepoint specified as the resource and SPO URL previously retrieved', async () => {
    const d: Date = new Date();
    d.setMinutes(d.getMinutes() + 1);
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
    auth.service.accessTokens['https://contoso.sharepoint.com'] = {
      expiresOn: d.toString(),
      accessToken: 'ABC'
    };

    await command.action(logger, { options: { debug: false, resource: 'sharepoint' } });
    assert(loggerLogSpy.calledWith('ABC'));
  });

  it('correctly handles error when retrieving access token', async () => {
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => Promise.reject('An error has occurred'));

    await assert.rejects(command.action(logger, { options: { debug: false, resource: 'https://graph.microsoft.com' } } as any), new CommandError('An error has occurred'));
  });

  it('returns error when sharepoint specified as resource and SPO URL not available', async () => {
    const d: Date = new Date();
    d.setMinutes(d.getMinutes() + 1);
    auth.service.accessTokens['https://contoso.sharepoint.com'] = {
      expiresOn: d.toString(),
      accessToken: 'ABC'
    };

    await assert.rejects(command.action(logger, { options: { debug: false, resource: 'sharepoint' } } as any), new CommandError(`SharePoint URL undefined. Use the 'm365 spo set --url https://contoso.sharepoint.com' command to set the URL`));
  });
});