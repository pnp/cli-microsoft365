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
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./roledefinition-add');

describe(commands.ROLEDEFINITION_ADD, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    auth.service.connected = true;
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
    assert.strictEqual(command.name.startsWith(commands.ROLEDEFINITION_ADD), true);
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
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/roledefinitions') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
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
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/roledefinitions') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
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
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/roledefinitions') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        name: 'test'
      }
    }), new CommandError(err));
  });
});
