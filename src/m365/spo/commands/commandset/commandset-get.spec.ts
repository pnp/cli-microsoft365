import * as assert from 'assert';
import * as sinon from 'sinon';
import * as os from 'os';
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
import { odata } from '../../../../utils/odata';
import { v4 } from 'uuid';
const command: Command = require('./commandset-get');

describe(commands.COMMANDSET_GET, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/project-z';
  const customActionId = '0a8e82b5-651f-400b-b537-9a739f92d6b4';
  const clientSideComponentId = '2397e6ef-4b89-4508-aea2-e375e312c76d';
  const customActionTitle = 'Alerts';
  const customActionObject = { 'ClientSideComponentId': clientSideComponentId, 'ClientSideComponentProperties': '', 'CommandUIExtension': null, 'Description': null, 'Group': null, 'HostProperties': '', 'Id': customActionId, 'ImageUrl': null, 'Location': 'ClientSideExtension.ListViewCommandSet.CommandBar', 'Name': '{0a8e82b5-651f-400b-b537-9a739f92d6b4}', 'RegistrationId': '119', 'RegistrationType': 1, 'Rights': { 'High': 0, 'Low': 0 }, 'Scope': 3, 'ScriptBlock': null, 'ScriptSrc': null, 'Sequence': 65536, 'Title': customActionTitle, 'Url': null, 'VersionOfUserCustomAction': '1.0.1.0' };
  const customActionResponse: any[] = [customActionObject];

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      odata.getAllItems,
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName,
      session.getId
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.COMMANDSET_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets custom action from specific site by id with scope "Web"', async () => {
    const scope = 'Web';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/${scope}/UserCustomActions(guid'${customActionId}')`) {
        return customActionObject;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, id: customActionId, scope: scope, verbose: true } });
    assert(loggerLogSpy.calledWith(customActionObject));
  });

  it('gets custom action from specific site by title with scope "Site"', async () => {
    const scope = 'Site';
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `${webUrl}/_api/${scope}/UserCustomActions?$filter=startswith(Location,'ClientSideExtension.ListViewCommandSet') and Title eq '${customActionTitle}'`) {
        return customActionResponse;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, title: customActionTitle, scope: scope, verbose: true } });
    assert(loggerLogSpy.calledWith(customActionObject));
  });

  it('gets custom action from specific site by clientSideComponentId without specifying scope', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `${webUrl}/_api/Site/UserCustomActions?$filter=startswith(Location,'ClientSideExtension.ListViewCommandSet') and ClientSideComponentId eq guid'${clientSideComponentId}'`) {
        return customActionResponse;
      }
      if (url === `${webUrl}/_api/Web/UserCustomActions?$filter=startswith(Location,'ClientSideExtension.ListViewCommandSet') and ClientSideComponentId eq guid'${clientSideComponentId}'`) {
        return [];
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, clientSideComponentId: clientSideComponentId, verbose: true } });
    assert(loggerLogSpy.calledWith(customActionObject));
  });

  it('throws error when command set not found by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/Site/UserCustomActions(guid'${customActionId}')`) {
        return { 'odata.null': true };
      }

      if (opts.url === `${webUrl}/_api/Web/UserCustomActions(guid'${customActionId}')`) {
        return { 'odata.null': true };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, id: customActionId, verbose: true } })
      , new CommandError(`Command set with id ${customActionId} can't be found.`));
  });

  it('throws error when command set is found by id but is not of type command set', async () => {
    const customActionObjectClone = { ...customActionObject };
    customActionObjectClone.Location = 'ClientSideExtension.ApplicationCustomizer';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/Site/UserCustomActions(guid'${customActionId}')`) {
        return customActionObjectClone;
      }

      if (opts.url === `${webUrl}/_api/Web/UserCustomActions(guid'${customActionId}')`) {
        return { 'odata.null': true };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, id: customActionId, verbose: true } })
      , new CommandError(`Custom action with id ${customActionId} is not a command set.`));
  });

  it('throws error when command set is not found by clientSideComponentId', async () => {
    const scope = 'Site';
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `${webUrl}/_api/${scope}/UserCustomActions?$filter=startswith(Location,'ClientSideExtension.ListViewCommandSet') and ClientSideComponentId eq guid'${clientSideComponentId}'`) {
        return [];
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, clientSideComponentId: clientSideComponentId, scope: scope, verbose: true } })
      , new CommandError(`No command set with clientSideComponentId '${clientSideComponentId}' found.`));
  });

  it('throws error when command set is not found by title', async () => {
    const scope = 'Web';
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `${webUrl}/_api/${scope}/UserCustomActions?$filter=startswith(Location,'ClientSideExtension.ListViewCommandSet') and Title eq '${customActionTitle}'`) {
        return [];
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, title: customActionTitle, scope: scope, verbose: true } })
      , new CommandError(`No command set with title '${customActionTitle}' found.`));
  });

  it('throws error when multiple command sets are found by title', async () => {
    const commandSetResponseClone = [...customActionResponse];
    const customActionObjectClone = { ...customActionObject };
    const customActionCloneId = v4();
    customActionObjectClone.Id = customActionCloneId;
    commandSetResponseClone.push(customActionObjectClone);
    const scope = 'Web';
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `${webUrl}/_api/${scope}/UserCustomActions?$filter=startswith(Location,'ClientSideExtension.ListViewCommandSet') and Title eq '${customActionTitle}'`) {
        return commandSetResponseClone;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, title: customActionTitle, scope: scope, verbose: true } })
      , new CommandError(`Multiple command sets with title '${customActionTitle}' found. Please disambiguate using IDs: ${os.EOL}${commandSetResponseClone.map(commandSet => `- ${commandSet.Id}`).join(os.EOL)}.`));
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'invalid', id: customActionId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the id option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, id: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the clientSideComponentId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, clientSideComponentId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the scope option is not a valid scope option', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, id: clientSideComponentId, scope: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if options are specified properly with id', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, id: clientSideComponentId, scope: 'All' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if options are specified properly with title', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, title: customActionTitle, scope: 'Web' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if options are specified properly with clientSideComponentId', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, clientSideComponentId: clientSideComponentId, scope: 'Site' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
