import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './tenant-recyclebinitem-restore.js';
import { formatting } from '../../../../utils/formatting.js';
import { ListItemListOptions, spoListItem } from '../../../../utils/spoListItem.js';

describe(commands.TENANT_RECYCLEBINITEM_RESTORE, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  const siteUrl = 'https://contoso.sharepoint.com/sites/hr';
  const spoAdminUrl = 'https://contoso-admin.sharepoint.com';
  const siteRestoreUrl = 'https://contoso-admin.sharepoint.com/_api/SPO.Tenant/RestoreDeletedSite';
  const adminSitesListTitle = 'DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS';

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
    commandInfo = cli.getCommandInfo(command);
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
      spoListItem.getListItems
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TENANT_RECYCLEBINITEM_RESTORE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { siteUrl: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { siteUrl: siteUrl } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it(`restores deleted group from a deleted team site`, async () => {
    const groupId = '4b3f5e3f-6e1f-4b1e-8b5f-0f5f5f5f5f5f';
    const groupRestoreUrl = `https://graph.microsoft.com/v1.0/directory/deletedItems/${groupId}/restore`;
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === siteRestoreUrl) {
        return;
      }

      if (opts.url === groupRestoreUrl) {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(spoListItem, 'getListItems').callsFake(async (options: ListItemListOptions) => {
      if (options.webUrl === spoAdminUrl) {
        if (options.listTitle === adminSitesListTitle &&
          options.filter === `SiteUrl eq '${formatting.encodeQueryParameter(siteUrl)}'`
        ) {
          return [{ GroupId: groupId }] as any;
        }
      }

      throw 'Invalid request: ' + JSON.stringify(options);
    });

    await command.action(logger, { options: { siteUrl: siteUrl, verbose: true } });
    assert.strictEqual(postStub.lastCall.args[0].url, groupRestoreUrl);
  });

  it('restores a deleted SharePoint site without group', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === siteRestoreUrl) {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(spoListItem, 'getListItems').callsFake(async (options: ListItemListOptions) => {
      if (options.webUrl === spoAdminUrl) {
        if (options.listTitle === adminSitesListTitle &&
          options.filter === `SiteUrl eq '${formatting.encodeQueryParameter(siteUrl)}'`
        ) {
          return [{ GroupId: '00000000-0000-0000-0000-000000000000' }] as any;
        }
      }

      throw 'Invalid request: ' + JSON.stringify(options);
    });

    await command.action(logger, { options: { siteUrl: siteUrl, verbose: true } });
    assert(postStub.lastCall.args[0].url === siteRestoreUrl);
  });

  it('handles error when the site to restore is not found', async () => {
    const error = {
      error: {
        'odata.error': {
          code: '-2147024809, System.ArgumentException',
          message: {
            lang: 'en-US',
            value: `Unable to find the deleted site: ${siteUrl}`
          }
        }
      }
    };

    sinon.stub(request, 'post').rejects(error);

    await assert.rejects(command.action(logger, { options: { siteUrl: siteUrl, verbose: true } } as any), new CommandError(error.error['odata.error'].message.value));
  });
});