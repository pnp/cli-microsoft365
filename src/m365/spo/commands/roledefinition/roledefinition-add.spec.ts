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
import command from './roledefinition-add.js';

describe(commands.ROLEDEFINITION_ADD, () => {
  let log: any[];
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
    assert.strictEqual(command.name, commands.ROLEDEFINITION_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', name: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', name: 'abc' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails if non existing PermissionKind rights specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', name: 'abc', rights: 'abc' } }, commandInfo);
    assert.strictEqual(actual, `Rights option 'abc' is not recognized as valid PermissionKind choice. Please note it is case-sensitive. Allowed values are EmptyMask|ViewListItems|AddListItems|EditListItems|DeleteListItems|ApproveItems|OpenItems|ViewVersions|DeleteVersions|CancelCheckout|ManagePersonalViews|ManageLists|ViewFormPages|AnonymousSearchAccessList|Open|ViewPages|AddAndCustomizePages|ApplyThemeAndBorder|ApplyStyleSheets|ViewUsageData|CreateSSCSite|ManageSubwebs|CreateGroups|ManagePermissions|BrowseDirectories|BrowseUserInfo|AddDelPrivateWebParts|UpdatePersonalWebParts|ManageWeb|AnonymousSearchAccessWebLists|UseClientIntegration|UseRemoteAPIs|ManageAlerts|CreateAlerts|EditMyUserInfo|EnumeratePermissions|FullMask.`);
  });

  it('has correct PermissionKind rights specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', name: 'abc', rights: 'FullMask' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('offers autocomplete for the rights option', () => {
    const options = command.options;
    for (let i = 0; i < options.length; i++) {
      if (options[i].option.indexOf('--rights') > -1) {
        assert(options[i].autocomplete);
        return;
      }
    }
    assert(false);
  });

  it('adds role definition to web with name, description and right', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/roledefinitions') {
        return '';
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        name: 'test',
        description: 'test',
        rights: 'FullMask'
      }
    });
  });

  it('adds role definition to web with name', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/roledefinitions') {
        return '';
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        name: 'test'
      }
    });
  });

  it('handles reject request correctly', async () => {
    const err = 'request rejected';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/roledefinitions') > -1) {
        throw err;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        name: 'test'
      }
    }), new CommandError(err));
  });
});
