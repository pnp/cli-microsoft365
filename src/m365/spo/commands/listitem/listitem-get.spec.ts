import commands from '../../commands';
import Command from '../../../../Command';
import { CommandValidate, CommandOption, CommandError, CommandTypes } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./listitem-get');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.LISTITEM_GET, () => {
  let vorpal: Vorpal;
  let log: any[];
  let cmdInstance: any;

  const expectedTitle = `List Item 1`;

  const expectedId = 147;
  let actualId = 0;

  let getFakes = (opts: any) => {
    if ((opts.url as string).indexOf('/items(') > -1) {
      actualId = opts.url.match(/\/items\((\d+)\)/i)[1];
      return Promise.resolve(
        {
          "Attachments": false,
          "AuthorId": 3,
          "ContentTypeId": "0x0100B21BD271A810EE488B570BE49963EA34",
          "Created": "2018-03-15T10:43:10Z",
          "EditorId": 3,
          "GUID": "ea093c7b-8ae6-4400-8b75-e2d01154dffc",
          "ID": actualId,
          "Modified": "2018-03-15T10:43:10Z",
          "Title": expectedTitle,
        }
      );
    }
    return Promise.reject('Invalid request');
  }
  
  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.LISTITEM_GET), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });

  it('supports specifying URL', () => {
    const options = (command.options() as CommandOption[]);
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<webUrl>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('configures command types', () => {
    assert.notEqual(typeof command.types(), 'undefined', 'command types undefined');
    assert.notEqual((command.types() as CommandTypes).string, 'undefined', 'command string types undefined');
  });

  it('fails validation if listTitle and listId option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: expectedId } });
    assert.notEqual(actual, true);
  });

  it('fails validation if listTitle and listId are specified together', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', id: expectedId } });
    assert.notEqual(actual, true);
  });

  it('fails validation if the webUrl option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { listTitle: 'Demo List', id: expectedId } });
    assert.notEqual(actual, true);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'foo', listTitle: 'Demo List', id: expectedId } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', id: expectedId } });
    assert(actual);
  });

  it('fails validation if the listId option is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: 'foo', id: expectedId } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the listId option is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', id: expectedId } });
    assert(actual);
  });

  it('fails validation if the id option is not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if the specified id is not a number', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', id: 'a' } });
    assert.notEqual(actual, true);
  });

  it('returns listItemInstance object when list item is requested', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);

    command.allowUnknownOptions();

    let options: any = { 
      debug: true, 
      listTitle: 'Demo List', 
      webUrl: 'https://contoso.sharepoint.com/sites/project-x', 
      id: expectedId
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert.equal(actualId, expectedId);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns listItemInstance object when list item is requested with an output type of json, and a list of fields are specified', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);

    command.allowUnknownOptions();

    let options: any = { 
      debug: false, 
      listTitle: 'Demo List', 
      webUrl: 'https://contoso.sharepoint.com/sites/project-x', 
      id: expectedId,
      output: "json",
      fields: "ID,Modified"
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert.equal(actualId, expectedId);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns listItemInstance object when list item is requested with an output type of text, and no list of fields', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);

    command.allowUnknownOptions();

    let options: any = { 
      debug: false, 
      listTitle: 'Demo List', 
      webUrl: 'https://contoso.sharepoint.com/sites/project-x', 
      id: expectedId,
      output: "text"
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert.equal(actualId, expectedId);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns listItemInstance object when list item is requested with an output type of text, and a list of fields specified', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);

    command.allowUnknownOptions();

    let options: any = { 
      debug: false, 
      listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', 
      webUrl: 'https://contoso.sharepoint.com/sites/project-x', 
      id: expectedId,
      output: "json"
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert.equal(actualId, expectedId);
        done();
      }
      catch (e) {
        done(e);
      }
    });    
  });

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'get').callsFake(() => Promise.reject('An error has occurred'));

    let options: any = { 
      debug: false, 
      listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', 
      webUrl: 'https://contoso.sharepoint.com/sites/project-x', 
      id: expectedId,
      output: "json"
    }

    cmdInstance.action({ options: options }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });    
  });

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.LISTITEM_GET));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => { },
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });
});