import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command from './site-remove.js';
import { odata } from '../../../../utils/odata.js';
import { settingsNames } from '../../../../settingsNames.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { formatting } from '../../../../utils/formatting.js';
import request from '../../../../request.js';
import { CommandError } from '../../../../Command.js';

describe(commands.SITE_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptIssued: boolean = false;

  const siteUrl = 'https://contoso.sharepoint.com/sites/project-x';
  const adminUrl = 'https://contoso-admin.sharepoint.com';
  const odataUrl = `${adminUrl}/_api/web/lists/GetByTitle('DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS')/items?$filter=SiteUrl eq '${formatting.encodeQueryParameter(siteUrl)}'&$select=GroupId,TimeDeleted,SiteId`;

  const siteDetailsNonGroup = {
    GroupId: '00000000-0000-0000-0000-000000000000',
    SiteId: 'b01dfb5a-ed2d-4f65-8434-f2e51f182dec',
    TimeDeleted: null
  };
  const siteDetailsGroup = {
    GroupId: '8f5ee9a8-7e71-410b-81fd-c661b00d7169',
    SiteId: 'b01dfb5a-ed2d-4f65-8434-f2e51f182dec',
    TimeDeleted: null
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getSpoAdminUrl').resolves(adminUrl);
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    (command as any).pollingInterval = 0;
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

    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName: string, defaultValue: any) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(cli, 'promptForConfirmation').callsFake(async () => {
      promptIssued = true;
      return false;
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      cli.promptForConfirmation,
      cli.getSettingWithDefaultValue,
      odata.getAllItems,
      request.delete,
      request.post,
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SITE_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('deletes a classic site and also immediately deletes it from the recycle bin', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === odataUrl) {
        return [siteDetailsNonGroup];
      }
      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RemoveSite`) {
        return;
      }
      if (opts.url === `${adminUrl}/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RemoveDeletedSite`) {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { url: siteUrl, skipRecycleBin: true, force: true, verbose: true } });
    assert(postStub.calledTwice);
    assert.strictEqual(postStub.firstCall.args[0].data.siteUrl, siteUrl);
    assert.strictEqual(postStub.secondCall.args[0].data.siteUrl, siteUrl);
  });

  it('deletes a group site, deletes the m365 group from entra id', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === odataUrl) {
        return [siteDetailsGroup];
      }
      throw 'Invalid request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/Microsoft.Graph.Group/${siteDetailsGroup.GroupId}?$select=id`) {
        throw {
          response: {
            status: 404
          }
        };
      }
      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_api/GroupSiteManager/Delete?siteUrl='${formatting.encodeQueryParameter(siteUrl)}'`) {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { url: siteUrl, force: true, verbose: true } });
    assert(postStub.calledOnce);
  });

  it('deletes a group site, deletes the m365 group from entra id and immediately deletes it from the recycle bin', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === odataUrl) {
        return [siteDetailsGroup];
      }
      throw 'Invalid request';
    });

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/Microsoft.Graph.Group/${siteDetailsGroup.GroupId}?$select=id`) {
        throw {
          response: {
            status: 404
          }
        };
      }
      throw 'Invalid request';
    });

    getRequestStub.onSecondCall().callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/Microsoft.Graph.Group/${siteDetailsGroup.GroupId}?$select=id`) {
        return { id: siteDetailsGroup.GroupId };
      }
      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_api/GroupSiteManager/Delete?siteUrl='${formatting.encodeQueryParameter(siteUrl)}'`) {
        return;
      }
      if (opts.url === `${adminUrl}/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RemoveDeletedSite`) {
        return;
      }
      throw 'Invalid request';
    });

    const deleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/Microsoft.Graph.Group/${siteDetailsGroup.GroupId}`) {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { url: siteUrl, skipRecycleBin: true, force: true, verbose: true } });
    assert(deleteStub.calledOnce);
  });

  it('deletes a group site, deletes the m365 group from entra id and immediately deletes the site from the recycle bin, but skips deletion of the m365 group when it does not exist in Entra', async () => {
    let amountOfCalls = 0;

    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === odataUrl) {
        return [siteDetailsGroup];
      }
      throw 'Invalid request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/Microsoft.Graph.Group/${siteDetailsGroup.GroupId}?$select=id`) {
        amountOfCalls++;
        throw {
          response: {
            status: 404
          }
        };
      }
      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_api/GroupSiteManager/Delete?siteUrl='${formatting.encodeQueryParameter(siteUrl)}'`) {
        return;
      }
      if (opts.url === `${adminUrl}/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RemoveDeletedSite`) {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { url: siteUrl, skipRecycleBin: true, force: true, verbose: true } });
    assert.strictEqual(amountOfCalls, 24);
  });

  it('deletes a group site from recycle bin and also removes the m365 group from entra id recycle bin', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === odataUrl) {
        return [{ ...siteDetailsGroup, TimeDeleted: new Date().toISOString() }];
      }
      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RemoveDeletedSite`) {
        return;
      }
      throw 'Invalid request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/Microsoft.Graph.Group/${siteDetailsGroup.GroupId}?$select=id`) {
        return { id: siteDetailsGroup.GroupId };
      }
      throw 'Invalid request';
    });

    const deleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/Microsoft.Graph.Group/${siteDetailsGroup.GroupId}`) {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { url: siteUrl, fromRecycleBin: true, force: true, verbose: true } });
    assert(deleteStub.calledOnce);
  });

  it('deletes a group site from recycle bin, removes trailing slash, and skips deletion of the m365 group from entra id recycle bin if it does not exist', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === odataUrl) {
        return [{ ...siteDetailsGroup, TimeDeleted: new Date().toISOString() }];
      }
      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RemoveDeletedSite`) {
        return;
      }
      throw 'Invalid request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/Microsoft.Graph.Group/${siteDetailsGroup.GroupId}?$select=id`) {
        throw {
          response: {
            status: 404
          }
        };
      }
      throw 'Invalid request';
    });

    const deleteStub = sinon.stub(request, 'delete').resolves();

    await command.action(logger, { options: { url: `${siteUrl}/`, fromRecycleBin: true, verbose: true } });
    assert(postStub.calledOnce);
    assert(deleteStub.notCalled);
  });


  it('throws error if site is not found', async () => {
    sinon.stub(odata, 'getAllItems').resolves([]);

    await assert.rejects(command.action(logger, { options: { url: siteUrl, verbose: true, force: true } }),
      new CommandError('Site not found in the tenant.'));
  });

  it('throws error if the endpoint fails when retrieving the deleted group', async () => {
    const errorMessage = 'Error occurred on processing the request.';
    sinon.stub(odata, 'getAllItems').resolves([{ ...siteDetailsGroup, TimeDeleted: new Date().toISOString() }]);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/Microsoft.Graph.Group/${siteDetailsGroup.GroupId}?$select=id`) {
        throw {
          error: {
            code: '-1, InvalidOperationException',
            message: errorMessage
          }
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { url: siteUrl, fromRecycleBin: true, verbose: true, force: true } }),
      new CommandError(errorMessage));
  });

  it('throws error if site has already been deleted when trying to remove it', async () => {
    sinon.stub(odata, 'getAllItems').resolves([{ ...siteDetailsNonGroup, TimeDeleted: new Date().toISOString() }]);

    await assert.rejects(command.action(logger, { options: { url: siteUrl, verbose: true, force: true } }),
      new CommandError('Site is already in the recycle bin. Use --fromRecycleBin to permanently delete it.'));
  });

  it('throws an error when attempting to delete a site that is not present in the recycle bin', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === odataUrl) {
        return [siteDetailsNonGroup];
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { url: siteUrl, fromRecycleBin: true, verbose: true, force: true } }),
      new CommandError('Site is currently not in the recycle bin. Do not specify --fromRecycleBin if you want to try to remove it from the active sites.'));
  });

  it(`correctly shows deprecation warning for option 'wait'`, async () => {
    const chalk = (await import('chalk')).default;
    const loggerErrSpy = sinon.spy(logger, 'logToStderr');

    await command.action(logger, { options: { url: siteUrl, wait: true } });
    assert(loggerErrSpy.calledWith(chalk.yellow(`Option 'wait' is deprecated and will be removed in the next major release.`)));

    sinonUtil.restore(loggerErrSpy);
  });

  it('prompts before removing the site when force option not passed', async () => {
    await command.action(logger, { options: { url: siteUrl, verbose: true } });
    assert(promptIssued);
  });

  it('aborts removing the site when prompt not confirmed', async () => {
    const postStub = sinon.stub(request, 'post').resolves();

    await command.action(logger, { options: { url: siteUrl, verbose: true } });
    assert(postStub.notCalled);
  });

  it('passes validation if siteUrl is a valid url', async () => {
    const actual = await command.validate({ options: { url: siteUrl } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if only fromRecycleBin is specified', async () => {
    const actual = await command.validate({ options: { url: siteUrl, fromRecycleBin: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if only skipRecycleBin is specified', async () => {
    const actual = await command.validate({ options: { url: siteUrl, skipRecycleBin: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if siteUrl is a valid url', async () => {
    const actual = await command.validate({ options: { url: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when trying to remove the root site collection', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both fromRecycleBin and skipRecycleBin is specified is a valid url', async () => {
    const actual = await command.validate({ options: { url: siteUrl, fromRecycleBin: true, skipRecycleBin: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
});