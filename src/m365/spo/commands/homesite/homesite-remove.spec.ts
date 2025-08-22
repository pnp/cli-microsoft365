import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command from './homesite-remove.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { z } from 'zod';

describe(commands.HOMESITE_REMOVE, () => {
  let log: any[];
  let logger: Logger;
  let promptIssued: boolean = false;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;
  const siteId = '00000000-0000-0000-0000-000000000010';

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
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
    sinon.stub(cli, 'promptForConfirmation').callsFake(async () => {
      promptIssued = true;
      return false;
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      cli.promptForConfirmation,
      spo.getSiteAdminPropertiesByUrl
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.HOMESITE_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('prompts before removing the Home Site when force option is not passed', async () => {
    await command.action(logger, { options: { debug: true } } as any);

    assert(promptIssued);
  });

  it('aborts removing Home Site when force option is not passed and prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'post');

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, { options: {} });
    assert(postSpy.notCalled);
  });

  it('fails validation if the url option is not a valid SharePoint site url', async () => {
    const actual = commandOptionsSchema.safeParse({ url: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = commandOptionsSchema.safeParse({ url: 'https://contoso.sharepoint.com' });
    assert.strictEqual(actual.success, true);
  });

  it('removes the Home Site specified by URL', async () => {
    sinon.stub(spo, 'getSiteAdminPropertiesByUrl').resolves({ SiteId: siteId } as any);

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_api/SPO.Tenant/RemoveTargetedSite`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com', verbose: true } });
    assert(postStub.calledOnce);
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { siteId });
  });

  it('removes the Home Site specified by URL by force', async () => {
    sinon.stub(spo, 'getSiteAdminPropertiesByUrl').resolves({ SiteId: siteId } as any);

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_api/SPO.Tenant/RemoveTargetedSite`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com', force: true } });
    assert(postStub.calledOnce);
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { siteId });
  });

  it('correctly handles error when attempting to remove a site that is not a home site or Viva Connections', async () => {
    sinon.stub(request, 'post').rejects({
      error: {
        "odata.error": {
          "code": "-2146232832, Microsoft.SharePoint.SPException",
          "message": {
            "lang": "en-US",
            "value": "[Error ID: 03fc404e-0f70-4607-82e8-8fdb014e8658] The site with ID \"8e4686ed-b00c-4c5f-a0e2-4197081df5d5\" has not been added as a home site or Viva Connections. Check aka.ms/homesites for details."
          }
        }
      }
    });

    await assert.rejects(
      command.action(logger, { options: { debug: true, force: true } } as any),
      new CommandError('[Error ID: 03fc404e-0f70-4607-82e8-8fdb014e8658] The site with ID \"8e4686ed-b00c-4c5f-a0e2-4197081df5d5\" has not been added as a home site or Viva Connections. Check aka.ms/homesites for details.')
    );
  });
});
