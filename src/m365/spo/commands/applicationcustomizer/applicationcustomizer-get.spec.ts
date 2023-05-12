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
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./applicationcustomizer-get');

describe(commands.APPLICATIONCUSTOMIZER_GET, () => {
  const title = 'Some customizer';
  const id = '14125658-a9bc-4ddf-9c75-1b5767c9a337';
  const clientSideComponentId = '7096cded-b83d-4eab-96f0-df477ed7c0bc';
  const webUrl = 'https://contoso.sharepoint.com/sites/sales';
  const applicationCustomizerGetResponse = {
    "ClientSideComponentId": clientSideComponentId,
    "ClientSideComponentProperties": "",
    "CommandUIExtension": null,
    "Description": null,
    "Group": null,
    "Id": id,
    "ImageUrl": null,
    "Location": "ClientSideExtension.ApplicationCustomizer",
    "Name": title,
    "RegistrationId": null,
    "RegistrationType": 0,
    "Rights": "{\"High\":0,\"Low\":0}",
    "Scope": 3,
    "ScriptBlock": null,
    "ScriptSrc": null,
    "Sequence": 0,
    "Title": title,
    "Url": null,
    "VersionOfUserCustomAction": "16.0.1.0"
  };

  const applicationCustomizerGetMultipleResponse = {
    "value": [
      applicationCustomizerGetResponse
    ]
  };

  const applicationCustomizerGetOutput = {
    ClientSideComponentId: '7096cded-b83d-4eab-96f0-df477ed7c0bc',
    ClientSideComponentProperties: '',
    CommandUIExtension: null,
    Description: null,
    Group: null,
    Id: '14125658-a9bc-4ddf-9c75-1b5767c9a337',
    ImageUrl: null,
    Location: 'ClientSideExtension.ApplicationCustomizer',
    Name: 'Some customizer',
    RegistrationId: null,
    RegistrationType: 0,
    Rights: '"{\\"High\\":0,\\"Low\\":0}"',
    Scope: 'Web',
    ScriptBlock: null,
    ScriptSrc: null,
    Sequence: 0,
    Title: 'Some customizer',
    Url: null,
    VersionOfUserCustomAction: '16.0.1.0'
  };

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    auth.service.connected = true;
    auth.service.spoUrl = webUrl;
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
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APPLICATIONCUSTOMIZER_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'abc', webUrl: webUrl } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the clientSideComponentId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { clientSideComponentId: 'abc', webUrl: webUrl } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({
      options:
      {
        id: id,
        webUrl: 'foo'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when all options are specified', async () => {
    const actual = await command.validate({
      options: {
        title: title,
        id: id,
        clientSideComponentId: clientSideComponentId,
        webUrl: webUrl
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when no options are specified', async () => {
    const actual = await command.validate({
      options: {
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when title and id options are specified', async () => {
    const actual = await command.validate({
      options: {
        title: title,
        id: id,
        webUrl: webUrl
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when title and clientSideComponentId options are specified', async () => {
    const actual = await command.validate({
      options: {
        title: title,
        clientSideComponentId: clientSideComponentId,
        webUrl: webUrl
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when id and clientSideComponentId options are specified', async () => {
    const actual = await command.validate({
      options: {
        id: id,
        clientSideComponentId: clientSideComponentId,
        webUrl: webUrl
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the scope is not a valid application customizer scope', async () => {
    const actual = await command.validate({ options: { id: id, webUrl: webUrl, scope: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if id is a valid GUID', async () => {
    const actual = await command.validate({ options: { id: id, webUrl: webUrl } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passed validation when title specified', async () => {
    const actual = await command.validate({ options: { title: title, webUrl: webUrl } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if clientSideComponentId is a valid GUID', async () => {
    const actual = await command.validate({ options: { clientSideComponentId: clientSideComponentId, webUrl: webUrl } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('humanize scope shows correct value when scope odata is 2', () => {
    const actual = (command as any)["humanizeScope"](2);
    assert(actual === "Site");
  });

  it('humanize scope shows correct value when scope odata is 3', () => {
    const actual = (command as any)["humanizeScope"](3);
    assert(actual === "Web");
  });

  it('humanize scope shows the scope odata value when is different than 2 and 3', () => {
    const actual = (command as any)["humanizeScope"](1);
    assert(actual === "1");
  });

  it('retrieves an application customizer by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/Web/UserCustomActions(guid'14125658-a9bc-4ddf-9c75-1b5767c9a337')`) {
        return applicationCustomizerGetResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        id: id,
        webUrl: webUrl,
        scope: 'Web'
      }
    });

    assert(loggerLogSpy.calledWith(applicationCustomizerGetOutput));
  });

  it('retrieves an application customizer by title', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/Web/UserCustomActions?$filter=Title eq 'Some%20customizer' and Location eq 'ClientSideExtension.ApplicationCustomizer'`) {
        return applicationCustomizerGetMultipleResponse;
      }

      if (opts.url === `${webUrl}/_api/Site/UserCustomActions?$filter=Title eq 'Some%20customizer' and Location eq 'ClientSideExtension.ApplicationCustomizer'`) {
        return {
          value: []
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        title: title,
        webUrl: webUrl,
        debug: true
      }
    });

    assert(loggerLogSpy.calledWith(applicationCustomizerGetOutput));
  });

  it('retrieves an application customizer by clientSideComponentId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/Web/UserCustomActions?$filter=ClientSideComponentId eq guid'7096cded-b83d-4eab-96f0-df477ed7c0bc' and Location eq 'ClientSideExtension.ApplicationCustomizer'`) {
        return applicationCustomizerGetMultipleResponse;
      }

      if (opts.url === `${webUrl}/_api/Site/UserCustomActions?$filter=ClientSideComponentId eq guid'7096cded-b83d-4eab-96f0-df477ed7c0bc' and Location eq 'ClientSideExtension.ApplicationCustomizer'`) {
        return {
          value: []
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        clientSideComponentId: clientSideComponentId,
        webUrl: webUrl,
        debug: true
      }
    });

    assert(loggerLogSpy.calledWith(applicationCustomizerGetOutput));
  });

  it('handles error when no application customizer with the specified id found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/Web/UserCustomActions(guid'14125658-a9bc-4ddf-9c75-1b5767c9a337')`) {
        return {
          'odata.null': true
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        id: id,
        webUrl: webUrl,
        scope: 'Web'
      }
    }), new CommandError(`No application customizer with id '${id}' found`));
  });

  it('handles error when no application customizer with the specified title found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/Web/UserCustomActions?$filter=Title eq 'Some%20customizer' and Location eq 'ClientSideExtension.ApplicationCustomizer'`) {
        return {
          value: [
          ]
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: title,
        webUrl: webUrl,
        scope: 'Web'
      }
    }), new CommandError(`No application customizer with title '${title}' found`));
  });

  it('handles error when no application customizer with the specified clientSideComponentId found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/Web/UserCustomActions?$filter=ClientSideComponentId eq guid'7096cded-b83d-4eab-96f0-df477ed7c0bc' and Location eq 'ClientSideExtension.ApplicationCustomizer'`) {
        return {
          value: [
          ]
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        clientSideComponentId: clientSideComponentId,
        webUrl: webUrl,
        scope: 'Web'
      }
    }), new CommandError(`No application customizer with Client Side Component Id '${clientSideComponentId}' found`));
  });

  it('handles error when multiple application customizers with the specified title found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/Web/UserCustomActions?$filter=Title eq 'Some%20customizer' and Location eq 'ClientSideExtension.ApplicationCustomizer'`) {
        return {
          value: [
            {
              ClientSideComponentId: 'b41916e7-e69d-467f-b37f-ff8ecf8f99f2',
              ClientSideComponentProperties: "'{testMessage:Test message}'",
              CommandUIExtension: null,
              Description: null,
              Group: null,
              HostProperties: '',
              Id: 'a70d8013-3b9f-4601-93a5-0e453ab9a1f3',
              ImageUrl: null,
              Location: 'ClientSideExtension.ApplicationCustomizer',
              Name: 'YourName',
              RegistrationId: null,
              RegistrationType: 0,
              Rights: [Object],
              Scope: 3,
              ScriptBlock: null,
              ScriptSrc: null,
              Sequence: 0,
              Title: title,
              Url: null,
              VersionOfUserCustomAction: '16.0.1.0'
            },
            {
              ClientSideComponentId: 'b41916e7-e69d-467f-b37f-ff8ecf8f99f2',
              ClientSideComponentProperties: "'{testMessage:Test message}'",
              CommandUIExtension: null,
              Description: null,
              Group: null,
              HostProperties: '',
              Id: '63aa745f-b4dd-4055-a4d7-d9032a0cfc59',
              ImageUrl: null,
              Location: 'ClientSideExtension.ApplicationCustomizer',
              Name: 'YourName',
              RegistrationId: null,
              RegistrationType: 0,
              Rights: [Object],
              Scope: 3,
              ScriptBlock: null,
              ScriptSrc: null,
              Sequence: 0,
              Title: title,
              Url: null,
              VersionOfUserCustomAction: '16.0.1.0'
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: title,
        webUrl: webUrl,
        scope: 'Web'
      }
    }), new CommandError(`Multiple application customizers with title '${title}' found. Please disambiguate using IDs: a70d8013-3b9f-4601-93a5-0e453ab9a1f3, 63aa745f-b4dd-4055-a4d7-d9032a0cfc59`));
  });

  it('handles error when multiple application customizers with the specified clientSideComponentId found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/Web/UserCustomActions?$filter=ClientSideComponentId eq guid'7096cded-b83d-4eab-96f0-df477ed7c0bc' and Location eq 'ClientSideExtension.ApplicationCustomizer'`) {
        return {
          value: [
            {
              ClientSideComponentId: clientSideComponentId,
              ClientSideComponentProperties: "'{testMessage:Test message}'",
              CommandUIExtension: null,
              Description: null,
              Group: null,
              HostProperties: '',
              Id: 'a70d8013-3b9f-4601-93a5-0e453ab9a1f3',
              ImageUrl: null,
              Location: 'ClientSideExtension.ApplicationCustomizer',
              Name: 'YourName',
              RegistrationId: null,
              RegistrationType: 0,
              Rights: [Object],
              Scope: 3,
              ScriptBlock: null,
              ScriptSrc: null,
              Sequence: 0,
              Title: 'YourAppCustomizer 1',
              Url: null,
              VersionOfUserCustomAction: '16.0.1.0'
            },
            {
              ClientSideComponentId: clientSideComponentId,
              ClientSideComponentProperties: "'{testMessage:Test message}'",
              CommandUIExtension: null,
              Description: null,
              Group: null,
              HostProperties: '',
              Id: '63aa745f-b4dd-4055-a4d7-d9032a0cfc59',
              ImageUrl: null,
              Location: 'ClientSideExtension.ApplicationCustomizer',
              Name: 'YourName',
              RegistrationId: null,
              RegistrationType: 0,
              Rights: [Object],
              Scope: 3,
              ScriptBlock: null,
              ScriptSrc: null,
              Sequence: 0,
              Title: 'YourAppCustomizer 2',
              Url: null,
              VersionOfUserCustomAction: '16.0.1.0'
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        clientSideComponentId: clientSideComponentId,
        webUrl: webUrl,
        scope: 'Web'
      }
    }), new CommandError(`Multiple application customizers with Client Side Component Id '${clientSideComponentId}' found. Please disambiguate using IDs: a70d8013-3b9f-4601-93a5-0e453ab9a1f3, 63aa745f-b4dd-4055-a4d7-d9032a0cfc59`));
  });

  it('handles error when no valid application customizer with the specified id found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/Web/UserCustomActions(guid'14125658-a9bc-4ddf-9c75-1b5767c9a337')`) {
        return {
          "ClientSideComponentId": clientSideComponentId,
          "ClientSideComponentProperties": "",
          "CommandUIExtension": null,
          "Description": null,
          "Group": null,
          "Id": id,
          "ImageUrl": null,
          "Location": "ClientSideExtension.ListViewCommandSet",
          "Name": title,
          "RegistrationId": null,
          "RegistrationType": 0,
          "Rights": "{\"High\":0,\"Low\":0}",
          "Scope": "3",
          "ScriptBlock": null,
          "ScriptSrc": null,
          "Sequence": 0,
          "Title": title,
          "Url": null,
          "VersionOfUserCustomAction": "16.0.1.0"
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        id: id,
        webUrl: webUrl,
        scope: 'Web'
      }
    }), new CommandError(`No application customizer with id '${id}' found`));
  });
});