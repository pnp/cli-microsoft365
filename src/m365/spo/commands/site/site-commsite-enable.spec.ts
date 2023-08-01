import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './site-commsite-enable.js';

describe(commands.SITE_COMMSITE_ENABLE, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').resolves();
    sinon.stub(session, 'getId').resolves();
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
    assert.strictEqual(command.name, commands.SITE_COMMSITE_ENABLE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('enables communication site features on the specified site (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/sitepages/communicationsite/enable`) {
        return { "odata.null": true };
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: { debug: true, url: 'https://contoso.sharepoint.com' } } as any);
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'post').callsFake(() => Promise.reject('An error has occurred'));
    await assert.rejects(command.action(logger, { options: { debug: true, url: 'https://contoso.sharepoint.com' } } as any), new CommandError('An error has occurred'));
  });

  it('requires site URL', () => {
    const options = command.options;
    assert(options.find(o => o.option.indexOf('<url>') > -1));
  });

  it('supports specifying design package ID', () => {
    const options = command.options;
    assert(options.find(o => o.option.indexOf('[designPackageId]') > -1));
  });

  it('fails validation when no site URL specified', async () => {
    const actual = await command.validate({
      options: {}
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when invalid site URL specified', async () => {
    const actual = await command.validate({
      options: { url: 'http://contoso.sharepoint.com' }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when valid site URL specified', async () => {
    const actual = await command.validate({
      options: { url: 'https://contoso.sharepoint.com' }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when invalid design package ID specified', async () => {
    const actual = await command.validate({
      options: { url: 'https://contoso.sharepoint.com', designPackageId: 'invalid' }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when no design package ID specified', async () => {
    const actual = await command.validate({
      options: { url: 'https://contoso.sharepoint.com' }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when valid design package ID specified', async () => {
    const actual = await command.validate({
      options: { url: 'https://contoso.sharepoint.com', designPackageId: '96c933ac-3698-44c7-9f4a-5fd17d71af9e' }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when invalid design package specified', async () => {
    const actual = await command.validate({
      options: { url: 'https://contoso.sharepoint.com', designPackage: 'invalid' }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when valid design package specified', async () => {
    const actual = await command.validate({
      options: { url: 'https://contoso.sharepoint.com', designPackage: 'Topic' }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when designPackage and designPackageId specified (multiple options)', async () => {
    const actual = await command.validate({
      options: { url: 'https://contoso.sharepoint.com', designPackage: 'Topic', designPackageId: '96c933ac-3698-44c7-9f4a-5fd17d71af9e' }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('enables communication site features with Topic design package on the specified site', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/sitepages/communicationsite/enable`) {
        return { "odata.null": true };
      }

      throw 'Invalid request';
    });
    await assert.doesNotReject(command.action(logger, { options: { designPackage: 'Topic', url: 'https://contoso.sharepoint.com' } } as any));
  });

  it('enables communication site features with Showcase design package on the specified site', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/sitepages/communicationsite/enable`) {
        return { "odata.null": true };
      }

      throw 'Invalid request';
    });
    await assert.doesNotReject(command.action(logger, { options: { designPackage: 'Showcase', url: 'https://contoso.sharepoint.com' } } as any));
  });

  it('enables communication site features with Blank design package on the specified site', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/sitepages/communicationsite/enable`) {
        return { "odata.null": true };
      }

      throw 'Invalid request';
    });
    await assert.doesNotReject(command.action(logger, { options: { designPackage: 'Blank', url: 'https://contoso.sharepoint.com' } } as any));
  });

  it('enables communication site features with design package ID on the specified site', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/sitepages/communicationsite/enable`) {
        return { "odata.null": true };
      }

      throw 'Invalid request';
    });
    await assert.doesNotReject(command.action(logger, { options: { designPackageId: '96c933ac-3698-44c7-9f4a-5fd17d71af9e', url: 'https://contoso.sharepoint.com' } } as any));
  });
});