import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./list-view-field-add');

describe(commands.LIST_VIEW_FIELD_ADD, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let requests: any[];
  const stubAllGetRequests: any = () => {
    return sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/fields/getbyinternalnameortitle') > -1 || (opts.url as string).indexOf('/fields/getbyid') > -1) {
        return Promise.resolve({
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
        });
      }

      return Promise.reject('Invalid request');
    });
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.LIST_VIEW_FIELD_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('add the field by title to viewId and listTitle (debug)', (done) => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/views('cc27a922-8224-4296-90a5-ebbc54da2e81')/viewfields/addviewfield('Author')`) > -1) {
        if (opts.headers &&
          opts.headers.authorization &&
          (opts.headers.authorization as string).indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listTitle: 'Documents', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', title: 'Created By' } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web/lists/GetByTitle('Documents')/views('cc27a922-8224-4296-90a5-ebbc54da2e81')/viewfields/addviewfield('`) > -1 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      try {
        assert(correctRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('add the field by title to viewId and listTitle', (done) => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/views('cc27a922-8224-4296-90a5-ebbc54da2e81')/viewfields/addviewfield('Author')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listTitle: 'Documents', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', title: 'Created By' } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web/lists/GetByTitle('Documents')/views('cc27a922-8224-4296-90a5-ebbc54da2e81')/viewfields/addviewfield('`) > -1 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      try {
        assert(correctRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('add the field by title from viewTitle and listTitle (debug)', (done) => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/views/GetByTitle('All Documents')/viewfields/addviewfield('Author')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listTitle: 'Documents', viewTitle: 'MyView', title: 'Created By' } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web/lists/GetByTitle('Documents')/views/GetByTitle('MyView')/viewfields/addviewfield('`) > -1 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      try {
        assert(correctRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('add the field by title from viewTitle and listTitle', (done) => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/views/GetByTitle('MyView')/viewfields/addviewfield('Author')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listTitle: 'Documents', viewTitle: 'MyView', title: 'Created By' } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web/lists/GetByTitle('Documents')/views/GetByTitle('MyView')/viewfields/addviewfield('`) > -1 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      try {
        assert(correctRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('add the field by title from viewTitle and listId (debug)', (done) => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'0cd891ef-afce-4e55-b836-fce03286cccf')/views/GetByTitle('MyView')/viewfields/addviewfield('Author')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewTitle: 'MyView', title: 'Created By' } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web/lists(guid'0cd891ef-afce-4e55-b836-fce03286cccf')/views/GetByTitle('MyView')/viewfields/addviewfield('`) > -1 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      try {
        assert(correctRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('add the field by title from viewTitle and listId', (done) => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'0cd891ef-afce-4e55-b836-fce03286cccf')/views/GetByTitle('MyView')/viewfields/addviewfield('Author')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewTitle: 'MyView', title: 'Created By' } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web/lists(guid'0cd891ef-afce-4e55-b836-fce03286cccf')/views/GetByTitle('MyView')/viewfields/addviewfield('`) > -1 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      try {
        assert(correctRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('add the field by title to viewId and listId (debug)', (done) => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'0cd891ef-afce-4e55-b836-fce03286cccf')/views('cc27a922-8224-4296-90a5-ebbc54da2e81')/viewfields/addviewfield('Author')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', title: 'Created By' } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web/lists(guid'0cd891ef-afce-4e55-b836-fce03286cccf')/views('cc27a922-8224-4296-90a5-ebbc54da2e81')/viewfields/addviewfield('`) > -1 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      try {
        assert(correctRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('add the field by title to viewId and listId', (done) => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'0cd891ef-afce-4e55-b836-fce03286cccf')/views('cc27a922-8224-4296-90a5-ebbc54da2e81')/viewfields/addviewfield('Author')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', title: 'Created By' } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web/lists(guid'0cd891ef-afce-4e55-b836-fce03286cccf')/views('cc27a922-8224-4296-90a5-ebbc54da2e81')/viewfields/addviewfield('`) > -1 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      try {
        assert(correctRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('add the field by id to viewId and listTitle (debug)', (done) => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/views('cc27a922-8224-4296-90a5-ebbc54da2e81')/viewfields/addviewfield('Author')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listTitle: 'Documents', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce' } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web/lists/GetByTitle('Documents')/views('cc27a922-8224-4296-90a5-ebbc54da2e81')/viewfields/addviewfield('`) > -1 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      try {
        assert(correctRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('add the field by id to viewId and listTitle', (done) => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/views('cc27a922-8224-4296-90a5-ebbc54da2e81')/viewfields/addviewfield('Author')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listTitle: 'Documents', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce' } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web/lists/GetByTitle('Documents')/views('cc27a922-8224-4296-90a5-ebbc54da2e81')/viewfields/addviewfield('`) > -1 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      try {
        assert(correctRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('add the field by id from viewTitle and listTitle (debug)', (done) => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/views/GetByTitle('All Documents')/viewfields/addviewfield('Author')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listTitle: 'Documents', viewTitle: 'MyView', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce' } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web/lists/GetByTitle('Documents')/views/GetByTitle('MyView')/viewfields/addviewfield('`) > -1 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      try {
        assert(correctRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('add the field by id from viewTitle and listTitle', (done) => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/views/GetByTitle('MyView')/viewfields/addviewfield('Author')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listTitle: 'Documents', viewTitle: 'MyView', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce' } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web/lists/GetByTitle('Documents')/views/GetByTitle('MyView')/viewfields/addviewfield('`) > -1 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      try {
        assert(correctRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('add the field by id from viewTitle and listId (debug)', (done) => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'0cd891ef-afce-4e55-b836-fce03286cccf')/views/GetByTitle('MyView')/viewfields/addviewfield('Author')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewTitle: 'MyView', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce' } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web/lists(guid'0cd891ef-afce-4e55-b836-fce03286cccf')/views/GetByTitle('MyView')/viewfields/addviewfield('`) > -1 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      try {
        assert(correctRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('add the field by id from viewTitle and listId', (done) => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'0cd891ef-afce-4e55-b836-fce03286cccf')/views/GetByTitle('MyView')/viewfields/addviewfield('Author')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewTitle: 'MyView', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce' } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web/lists(guid'0cd891ef-afce-4e55-b836-fce03286cccf')/views/GetByTitle('MyView')/viewfields/addviewfield('`) > -1 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      try {
        assert(correctRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('add the field by id to viewId and listId (debug)', (done) => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'0cd891ef-afce-4e55-b836-fce03286cccf')/views('cc27a922-8224-4296-90a5-ebbc54da2e81')/viewfields/addviewfield('Author')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce' } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web/lists(guid'0cd891ef-afce-4e55-b836-fce03286cccf')/views('cc27a922-8224-4296-90a5-ebbc54da2e81')/viewfields/addviewfield('`) > -1 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      try {
        assert(correctRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('add the field by id to viewId and listId', (done) => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'0cd891ef-afce-4e55-b836-fce03286cccf')/views('cc27a922-8224-4296-90a5-ebbc54da2e81')/viewfields/addviewfield('Author')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce' } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web/lists(guid'0cd891ef-afce-4e55-b836-fce03286cccf')/views('cc27a922-8224-4296-90a5-ebbc54da2e81')/viewfields/addviewfield('`) > -1 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      try {
        assert(correctRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('move the field by title to the position index to viewId of listTitle (debug)', (done) => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf(`/viewfields/addviewfield('Author')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listTitle: 'Documents', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', title: 'Created By', position: 1 } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/viewfields/moveviewfieldto`) > -1 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      try {
        assert(correctRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('move the field by title to the position index to viewId of listTitle', (done) => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf(`/viewfields/addviewfield('Author')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listTitle: 'Documents', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', title: 'Created By', position: 1 } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/viewfields/moveviewfieldto`) > -1 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      try {
        assert(correctRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('move the field by title to the position index to viewTitle of listId (debug)', (done) => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf(`/viewfields/addviewfield('Author')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewTitle: 'MyView', title: 'Created By', position: 1 } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/viewfields/moveviewfieldto`) > -1 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      try {
        assert(correctRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('move the field by title to the position index to viewTitle of listId', (done) => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf(`/viewfields/addviewfield('Author')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewTitle: 'MyView', title: 'Created By', position: 1 } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/viewfields/moveviewfieldto`) > -1 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      try {
        assert(correctRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('uses correct API url when list id option is passed', (done) => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid') > -1) {
        return Promise.resolve('Correct Url');
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        viewId: '0cd891ef-afce-4e55-b836-fce03286cccf',
        id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce',
        webUrl: 'https://contoso.sharepoint.com',
        listId: 'cc27a922-8224-4296-90a5-ebbc54da2e81',
        confirm: true
      }
    }, () => {

      try {
        assert(true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('uses correct API url when list title option is passed', (done) => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists/GetByTitle(') > -1) {
        return Promise.resolve('Correct Url');
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        viewId: '0cd891ef-afce-4e55-b836-fce03286cccf',
        id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce',
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'Documents',
        confirm: true
      }
    }, () => {

      try {
        assert(true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
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

  it('fails validation if one of listId or listTitle options are not passed', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e85', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if one of viewId or viewTitle options are not passed', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if one of id or title options are not passed', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e85' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e85', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the listId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '12345', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the viewId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewId: '12345', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the id option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', id: '12345' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the listId option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce' } }, commandInfo);
    assert(actual);
  });

  it('passes validation if the viewId option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce' } }, commandInfo);
    assert(actual);
  });

  it('passes validation if the id option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if both listId and listTitle options are passed', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', listTitle: 'Documents', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both viewId and viewTitle options are passed', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewTitle: 'My view', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both id and title options are passed', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewTitle: 'My view', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce', title: 'Created By' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the position option is defined and is not a valid number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce', position: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the position option is defined and is a valid number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce', position: 1 } }, commandInfo);
    assert(actual);
  });
});