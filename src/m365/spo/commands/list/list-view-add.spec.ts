import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command from '../../../../Command';
import { sinonUtil, urlUtil } from '../../../../utils';
import request from '../../../../request';
import commands from '../../commands';
const command: Command = require('./list-view-add');

describe(commands.LIST_VIEW_ADD, () => {

  const validListTitle = 'List title';
  const validListId = '00000000-0000-0000-0000-000000000000';
  const validListUrl = '/Lists/SampleList';
  const validTitle = 'View title';
  const validWebUrl = 'https://contoso.sharepoint.com/sites/project-x';
  const validFieldsInput = 'Field1,Field2,Field3';

  const viewCreationResponse = {
    DefaultView: false,
    Hidden: false,
    Id: "00000000-0000-0000-0000-000000000000",
    MobileDefaultView: false,
    MobileView: false,
    Paged: true,
    PersonalView: false,
    ViewProjectedFields: null,
    ViewQuery: "",
    RowLimit: 30,
    Scope: 0,
    ServerRelativePath: {
      DecodedUrl: `/sites/project-x/Lists/${validListTitle}/${validTitle}.aspx`
    },
    ServerRelativeUrl: `/sites/project-x/Lists/${validListTitle}/${validTitle}.aspx`,
    Title: validTitle
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
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
    assert.strictEqual(command.name.startsWith(commands.LIST_VIEW_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('has correct option sets', () => {
    assert.deepStrictEqual(command.optionSets(), [['listId', 'listTitle', 'listUrl']]);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', () => {
    const actual = command.validate({ 
      options: { 
        webUrl: 'invalid', 
        listTitle: validListTitle,
        title: validTitle,
        fields: validFieldsInput
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if listId is not a valid GUID', () => {
    const actual = command.validate({ 
      options: { 
        webUrl: validWebUrl, 
        listId: 'invalid',
        title: validTitle,
        fields: validFieldsInput
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if rowLimit is not a number', () => {
    const actual = command.validate({ 
      options: { 
        webUrl: validWebUrl, 
        listId: validListId,
        title: validTitle,
        fields: validFieldsInput,
        rowLimit: 'invalid'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if rowLimit is lower than 1', () => {
    const actual = command.validate({ 
      options: { 
        webUrl: validWebUrl, 
        listId: validListId,
        title: validTitle,
        fields: validFieldsInput,
        rowLimit: 0
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when setting default and personal option', () => {
    const actual = command.validate({ 
      options: { 
        webUrl: validWebUrl, 
        listId: validListId,
        title: validTitle,
        fields: validFieldsInput,
        personal: true,
        default: true
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('correctly validates options', () => {
    const actual = command.validate({ 
      options: { 
        webUrl: validWebUrl, 
        listId: validListId,
        title: validTitle,
        fields: validFieldsInput
      }
    });
    assert.strictEqual(actual, true);
  });

  it('Correctly add view by list title', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `${validWebUrl}/_api/web/lists/getByTitle(\'${encodeURIComponent(validListTitle)}\')/views/add`) {
        return Promise.resolve(viewCreationResponse);
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        webUrl: validWebUrl,
        listTitle: validListTitle,
        title: validTitle,
        fields: validFieldsInput
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith(viewCreationResponse));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('Correctly add view by list id', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `${validWebUrl}/_api/web/lists(guid\'${encodeURIComponent(validListId)}\')/views/add`) {
        return Promise.resolve(viewCreationResponse);
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        webUrl: validWebUrl,
        listId: validListId,
        title: validTitle,
        fields: validFieldsInput
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith(viewCreationResponse));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('Correctly add view by list URL', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `${validWebUrl}/_api/web/GetList(\'${encodeURIComponent(urlUtil.getServerRelativePath(validWebUrl, validListUrl))}\')/views/add`) {
        return Promise.resolve(viewCreationResponse);
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        webUrl: validWebUrl,
        listUrl: validListUrl,
        title: validTitle,
        fields: validFieldsInput,
        rowLimit: 100
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith(viewCreationResponse));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});