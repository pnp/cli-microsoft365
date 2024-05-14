import assert from 'assert';
import sinon from 'sinon';
import { v4 } from 'uuid';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { odata } from '../../../../utils/odata.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './commandset-get.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.COMMANDSET_GET, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/project-z';
  const commandSetId = '0a8e82b5-651f-400b-b537-9a739f92d6b4';
  const clientSideComponentId = '2397e6ef-4b89-4508-aea2-e375e312c76d';
  const commandSetTitle = 'Alerts';
  const commandSetObject = { 'ClientSideComponentId': clientSideComponentId, 'ClientSideComponentProperties': '{"sampleTextOne":"One item is selected in the list.", "sampleTextTwo":"This command is always visible."}', 'CommandUIExtension': null, 'Description': null, 'Group': null, 'HostProperties': '', 'Id': commandSetId, 'ImageUrl': null, 'Location': 'ClientSideExtension.ListViewCommandSet.CommandBar', 'Name': '{0a8e82b5-651f-400b-b537-9a739f92d6b4}', 'RegistrationId': '119', 'RegistrationType': 1, 'Rights': { 'High': 0, 'Low': 0 }, 'Scope': 3, 'ScriptBlock': null, 'ScriptSrc': null, 'Sequence': 65536, 'Title': commandSetTitle, 'Url': null, 'VersionOfUserCustomAction': '1.0.1.0' };
  const commandSetResponse: any[] = [commandSetObject];

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      odata.getAllItems,
      request.get,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.COMMANDSET_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets command set from specific site by id with scope "Web"', async () => {
    const scope = 'Web';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/${scope}/UserCustomActions(guid'${commandSetId}')`) {
        return commandSetObject;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, id: commandSetId, scope: scope, verbose: true } });
    assert(loggerLogSpy.calledOnceWithExactly(commandSetObject));
  });

  it('gets command set from specific site by title with scope "Site"', async () => {
    const scope = 'Site';
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `${webUrl}/_api/${scope}/UserCustomActions?$filter=startswith(Location,'ClientSideExtension.ListViewCommandSet') and Title eq '${commandSetTitle}'`) {
        return commandSetResponse;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, title: commandSetTitle, scope: scope, verbose: true } });
    assert(loggerLogSpy.calledOnceWithExactly(commandSetObject));
  });

  it('gets command set from specific site by clientSideComponentId without specifying scope', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `${webUrl}/_api/Site/UserCustomActions?$filter=startswith(Location,'ClientSideExtension.ListViewCommandSet') and ClientSideComponentId eq guid'${clientSideComponentId}'`) {
        return commandSetResponse;
      }
      if (url === `${webUrl}/_api/Web/UserCustomActions?$filter=startswith(Location,'ClientSideExtension.ListViewCommandSet') and ClientSideComponentId eq guid'${clientSideComponentId}'`) {
        return [];
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, clientSideComponentId: clientSideComponentId, verbose: true } });
    assert(loggerLogSpy.calledOnceWithExactly(commandSetObject));
  });

  it('throws error when command set not found by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/Site/UserCustomActions(guid'${commandSetId}')`) {
        return { 'odata.null': true };
      }

      if (opts.url === `${webUrl}/_api/Web/UserCustomActions(guid'${commandSetId}')`) {
        return { 'odata.null': true };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, id: commandSetId, verbose: true } })
      , new CommandError(`Command set with id ${commandSetId} can't be found.`));
  });

  it('throws error when command set is found by id but is not of type command set', async () => {
    const commandSetObjectClone = { ...commandSetObject };
    commandSetObjectClone.Location = 'ClientSideExtension.ApplicationCustomizer';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/Site/UserCustomActions(guid'${commandSetId}')`) {
        return commandSetObjectClone;
      }

      if (opts.url === `${webUrl}/_api/Web/UserCustomActions(guid'${commandSetId}')`) {
        return { 'odata.null': true };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, id: commandSetId, verbose: true } })
      , new CommandError(`Custom action with id ${commandSetId} is not a command set.`));
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
      if (url === `${webUrl}/_api/${scope}/UserCustomActions?$filter=startswith(Location,'ClientSideExtension.ListViewCommandSet') and Title eq '${commandSetTitle}'`) {
        return [];
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, title: commandSetTitle, scope: scope, verbose: true } })
      , new CommandError(`No command set with title '${commandSetTitle}' found.`));
  });

  it('throws error when multiple command sets are found by title', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const commandSetResponseClone = [...commandSetResponse];
    const commandSetObjectClone = { ...commandSetObject };
    const commandSetCloneId = v4();
    commandSetObjectClone.Id = commandSetCloneId;
    commandSetResponseClone.push(commandSetObjectClone);
    const scope = 'Web';
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `${webUrl}/_api/${scope}/UserCustomActions?$filter=startswith(Location,'ClientSideExtension.ListViewCommandSet') and Title eq '${commandSetTitle}'`) {
        return commandSetResponseClone;
      }

      throw 'Invalid request';
    });


    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, title: commandSetTitle, scope: scope, verbose: true } })
      , new CommandError(`Multiple command sets with title 'Alerts' found. Found: 0a8e82b5-651f-400b-b537-9a739f92d6b4, ${commandSetCloneId}.`));
  });

  it('handles selecting single result when multiple command sets with the specified name found and cli is set to prompt', async () => {
    sinon.stub(cli, 'handleMultipleResultsFound').resolves(commandSetObject);

    const commandSetResponseClone = [...commandSetResponse];
    const commandSetObjectClone = { ...commandSetObject };
    const commandSetCloneId = v4();
    commandSetObjectClone.Id = commandSetCloneId;
    commandSetResponseClone.push(commandSetObjectClone);
    const scope = 'Site';
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `${webUrl}/_api/${scope}/UserCustomActions?$filter=startswith(Location,'ClientSideExtension.ListViewCommandSet') and Title eq '${commandSetTitle}'`) {
        return commandSetResponseClone;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, title: commandSetTitle, scope: scope, verbose: true } });
    assert(loggerLogSpy.calledOnceWithExactly(commandSetObject));
  });

  it('gets client side component properties from a command set by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/Web/UserCustomActions(guid'${commandSetId}')`) {
        return commandSetObject;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, id: commandSetId, clientSideComponentProperties: true } });
    assert(loggerLogSpy.calledOnceWithExactly(JSON.parse(commandSetObject.ClientSideComponentProperties)));
  });

  it('gets malformed client side component properties from a command set by id', async () => {
    const requestResult = { ...commandSetObject, ClientSideComponentProperties: '{"sampleTextOne": One item is selected in the list.}' };
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/Web/UserCustomActions(guid'${commandSetId}')`) {
        return requestResult;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, id: commandSetId, clientSideComponentProperties: true } });
    assert(loggerLogSpy.calledOnceWithExactly(requestResult.ClientSideComponentProperties));
  });

  it('gets undefined client side component properties from a command set by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/Web/UserCustomActions(guid'${commandSetId}')`) {
        return { ...commandSetObject, ClientSideComponentProperties: undefined };
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, id: commandSetId, clientSideComponentProperties: true } });
    assert(loggerLogSpy.calledOnceWithExactly(undefined));
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'invalid', id: commandSetId } }, commandInfo);
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
    const actual = await command.validate({ options: { webUrl: webUrl, title: commandSetTitle, scope: 'Web' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if options are specified properly with clientSideComponentId', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, clientSideComponentId: clientSideComponentId, scope: 'Site' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
