import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './site-commsite-enable.js';

describe(commands.SITE_COMMSITE_ENABLE, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').resolves();
    sinon.stub(session, 'getId').resolves();
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName: string, defaultValue: any) => {
      if (settingName === 'prompt') {
        return false;
      }

      return defaultValue;
    });
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
      request.post,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
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
    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, url: 'https://contoso.sharepoint.com' }) } as any);
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'post').callsFake(() => Promise.reject('An error has occurred'));
    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ debug: true, url: 'https://contoso.sharepoint.com' }) } as any), new CommandError('An error has occurred'));
  });

  it('fails validation when no site URL specified', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when invalid site URL specified', () => {
    const actual = commandOptionsSchema.safeParse({ url: 'http://contoso.sharepoint.com' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error.issues[0].message, `'http://contoso.sharepoint.com' is not a valid SharePoint Online site URL.`);
  });

  it('passes validation when valid site URL specified', () => {
    const actual = commandOptionsSchema.safeParse({ url: 'https://contoso.sharepoint.com' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation when invalid design package ID specified', () => {
    const actual = commandOptionsSchema.safeParse({ url: 'https://contoso.sharepoint.com', designPackageId: 'invalid' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error.issues[0].message, `invalid is not a valid GUID.`);
  });

  it('passes validation when no design package ID specified', () => {
    const actual = commandOptionsSchema.safeParse({ url: 'https://contoso.sharepoint.com' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when valid design package ID specified', () => {
    const actual = commandOptionsSchema.safeParse({ url: 'https://contoso.sharepoint.com', designPackageId: '96c933ac-3698-44c7-9f4a-5fd17d71af9e' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation when invalid design package specified', () => {
    const actual = commandOptionsSchema.safeParse({ url: 'https://contoso.sharepoint.com', designPackage: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when valid design package specified', () => {
    const actual = commandOptionsSchema.safeParse({ url: 'https://contoso.sharepoint.com', designPackage: 'Topic' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation when designPackage and designPackageId specified (multiple options)', () => {
    const actual = commandOptionsSchema.safeParse({ url: 'https://contoso.sharepoint.com', designPackage: 'Topic', designPackageId: '96c933ac-3698-44c7-9f4a-5fd17d71af9e' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ url: 'https://contoso.sharepoint.com', unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });

  it('enables communication site features with Topic design package on the specified site', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/sitepages/communicationsite/enable`) {
        return { "odata.null": true };
      }

      throw 'Invalid request';
    });
    await assert.doesNotReject(command.action(logger, { options: commandOptionsSchema.parse({ designPackage: 'Topic', url: 'https://contoso.sharepoint.com' }) }));
  });

  it('enables communication site features with Showcase design package on the specified site', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/sitepages/communicationsite/enable`) {
        return { "odata.null": true };
      }

      throw 'Invalid request';
    });
    await assert.doesNotReject(command.action(logger, { options: commandOptionsSchema.parse({ designPackage: 'Showcase', url: 'https://contoso.sharepoint.com' }) }));
  });

  it('enables communication site features with Blank design package on the specified site', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/sitepages/communicationsite/enable`) {
        return { "odata.null": true };
      }

      throw 'Invalid request';
    });
    await assert.doesNotReject(command.action(logger, { options: commandOptionsSchema.parse({ designPackage: 'Blank', url: 'https://contoso.sharepoint.com' }) }));
  });

  it('enables communication site features with design package ID on the specified site', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/sitepages/communicationsite/enable`) {
        return { "odata.null": true };
      }

      throw 'Invalid request';
    });
    await assert.doesNotReject(command.action(logger, { options: commandOptionsSchema.parse({ designPackageId: '96c933ac-3698-44c7-9f4a-5fd17d71af9e', url: 'https://contoso.sharepoint.com' }) }));
  });
});