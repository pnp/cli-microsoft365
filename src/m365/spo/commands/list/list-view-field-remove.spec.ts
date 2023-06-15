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
const command: Command = require('./list-view-field-remove');

describe(commands.LIST_VIEW_FIELD_REMOVE, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let requests: any[];
  let promptOptions: any;

  const stubAllGetRequests: any = () => {
    return sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/fields/getbyinternalnameortitle') > -1 || (opts.url as string).indexOf('/fields/getbyid') > -1) {
        return {
          "AllowDisplay": true,
          "AllowMultipleValues": false,
          "AutoIndexed": false,
          "CanBeDeleted": false,
          "ClientSideComponentId": "00000000-0000-0000-0000-000000000000",
          "ClientSideComponentProperties": null,
          "CustomFormatter": null,
          "DefaultFormula": null,
          "DefaultValue": null,
          "DependentLookupInternalNames": [],
          "Description": "",
          "Direction": "none",
          "EnforceUniqueValues": false,
          "EntityPropertyName": "Author",
          "FieldTypeKind": 20,
          "Filterable": true,
          "FromBaseType": true,
          "Group": "Custom Columns",
          "Hidden": false,
          "Id": "1df5e554-ec7e-46a6-901d-d85a3881cb18",
          "Indexed": false,
          "InternalName": "Author",
          "IsDependentLookup": false,
          "IsRelationship": false,
          "JSLink": "clienttemplates.js",
          "LookupField": "",
          "LookupList": "{f978b511-305d-45e9-a7e7-f234a67e956d}",
          "LookupWebId": "c0950f14-23ce-4778-977a-9df11b866ede",
          "PinnedToFiltersPane": false,
          "Presence": true,
          "PrimaryFieldId": null,
          "ReadOnlyField": true,
          "RelationshipDeleteBehavior": 0,
          "Required": false,
          "SchemaXml": "<Field ID=\"{1df5e554-ec7e-46a6-901d-d85a3881cb18}\" ColName=\"tp_Author\" RowOrdinal=\"0\" ReadOnly=\"TRUE\" Type=\"User\" List=\"UserInfo\" Name=\"Author\" DisplayName=\"Created By\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Author\" FromBaseType\"TRUE\" />",
          "Scope": "/sites/ninja/Shared Documents",
          "Sealed": false,
          "SelectionGroup": 0,
          "SelectionMode": 1,
          "ShowInFiltersPane": 0,
          "Sortable": true,
          "StaticName": "Author",
          "Title": "Created By",
          "TypeAsString": "User",
          "TypeDisplayName": "Person or Group",
          "TypeShortDescription": "Person or Group",
          "UnlimitedLengthInDocumentLibrary": false,
          "ValidationFormula": null,
          "ValidationMessage": null
        };
      }

      throw 'Invalid request';
    });
  };

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
    requests = [];
    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      Cli.prompt
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.LIST_VIEW_FIELD_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('supports specifying URL', () => {
    const options = command.options;
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<webUrl>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e85', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the listId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '12345', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the viewId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewId: '12345', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the fieldId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', id: '12345' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if valid options are specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before removing field from list view when confirmation argument not passed (list title, view id, field title)', async () => {
    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/ninja', listTitle: 'Documents', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', title: 'Created By' } });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts removing field by title from list view when prompt not confirmed', async () => {
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: false }
    ));
    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/ninja', listTitle: 'Documents', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', title: 'Created By' } });
    assert(requests.length === 0);
  });

  it('removes the field by title from viewId and listTitle when prompt confirmed (debug)', async () => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/views('cc27a922-8224-4296-90a5-ebbc54da2e81')/viewfields/removeviewfield('Author')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));
    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listTitle: 'Documents', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', title: 'Created By' } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web/lists/GetByTitle('Documents')/views('cc27a922-8224-4296-90a5-ebbc54da2e81')/viewfields/removeviewfield('`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes the field by title from viewTitle and listTitle', async () => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/views/GetByTitle('MyView')/viewfields/removeviewfield('Author')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listTitle: 'Documents', viewTitle: 'MyView', title: 'Created By', confirm: true } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web/lists/GetByTitle('Documents')/views/GetByTitle('MyView')/viewfields/removeviewfield('`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes the field by title from viewTitle and listId when prompt confirmed', async () => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'0cd891ef-afce-4e55-b836-fce03286cccf')/views/GetByTitle('MyView')/viewfields/removeviewfield('Author')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/ninja', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewTitle: 'MyView', title: 'Created By' } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web/lists(guid'0cd891ef-afce-4e55-b836-fce03286cccf')/views/GetByTitle('MyView')/viewfields/removeviewfield('`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes the field by title from viewId and listId when prompt confirmed', async () => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'0cd891ef-afce-4e55-b836-fce03286cccf')/views('cc27a922-8224-4296-90a5-ebbc54da2e81')/viewfields/removeviewfield('Author')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/ninja', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', title: 'Created By' } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web/lists(guid'0cd891ef-afce-4e55-b836-fce03286cccf')/views('cc27a922-8224-4296-90a5-ebbc54da2e81')/viewfields/removeviewfield('`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes the field by id from viewId and listTitle when prompt confirmed', async () => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/views('cc27a922-8224-4296-90a5-ebbc54da2e81')/viewfields/removeviewfield('Author')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    await command.action(logger, { options: { verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listTitle: 'Documents', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce' } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web/lists/GetByTitle('Documents')/views('cc27a922-8224-4296-90a5-ebbc54da2e81')/viewfields/removeviewfield('`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes the field by id from viewTitle and listTitle when prompt confirmed', async () => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/views/GetByTitle('MyView')/viewfields/removeviewfield('Author')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/ninja', listTitle: 'Documents', viewTitle: 'MyView', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce' } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web/lists/GetByTitle('Documents')/views/GetByTitle('MyView')/viewfields/removeviewfield('`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes the field by id from viewTitle and listId when prompt confirmed', async () => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'0cd891ef-afce-4e55-b836-fce03286cccf')/views/GetByTitle('MyView')/viewfields/removeviewfield('Author')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/ninja', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewTitle: 'MyView', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce' } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web/lists(guid'0cd891ef-afce-4e55-b836-fce03286cccf')/views/GetByTitle('MyView')/viewfields/removeviewfield('`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes the field by id from viewId and listId when prompt confirmed', async () => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'0cd891ef-afce-4e55-b836-fce03286cccf')/views('cc27a922-8224-4296-90a5-ebbc54da2e81')/viewfields/removeviewfield('Author')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/ninja', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce' } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web/lists(guid'0cd891ef-afce-4e55-b836-fce03286cccf')/views('cc27a922-8224-4296-90a5-ebbc54da2e81')/viewfields/removeviewfield('`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes the field by id from viewId and listUrl when prompt confirmed', async () => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/GetList('%2Fsites%2Fninja%2FShared%20Documents')/views('cc27a922-8224-4296-90a5-ebbc54da2e81')/viewfields/removeviewfield('Author')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/ninja', listUrl: '/sites/ninja/Shared Documents', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce' } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web/GetList('%2Fsites%2Fninja%2FShared%20Documents')/views('cc27a922-8224-4296-90a5-ebbc54da2e81')/viewfields/removeviewfield('`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('correctly handles error when removing the field', async () => {
    const error = {
      error: {
        'odata.error': {
          code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
          message: {
            value: 'An error has occurred'
          }
        }
      }
    };
    sinon.stub(request, 'get').rejects(error);

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        listUrl: '/sites/ninja/Shared Documents',
        viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81',
        id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce',
        confirm: true
      }
    } as any), new CommandError(error.error['odata.error'].message.value));
  });
});
