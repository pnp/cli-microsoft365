import commands from '../../commands';
import Command, {CommandValidate,CommandOption,CommandError} from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./search');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';
import { ResultTableRow } from './datatypes/ResultTableRow';
import { SearchResult } from './datatypes/SearchResult';

describe(commands.SEARCH, () => {
  let vorpal: Vorpal;
  let log: any[];
  let cmdInstance: any;
  let trackEvent: any;
  let telemetry: any;
  let returnArrayLength = 0;
  let stubAuth: any = () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf('/common/oauth2/token') > -1) {
        return Promise.resolve('abc');
      }

      if (opts.url.indexOf('/_api/contextinfo') > -1) {
        return Promise.resolve({
          FormDigestValue: 'abc'
        });
      }

      return Promise.reject('Invalid request');
    });
  }
  let getFakeRows = ():ResultTableRow[] => {
    return [
      {
        "Cells":[
          {"Key":"Rank","Value":"1","ValueType":"Edm.Double"},
          {"Key":"DocId","Value":"1","ValueType":"Edm.Int64"},
          {"Key":"Path","Value":"MyPath-item1","ValueType":"Edm.String"},
          {"Key":"Author","Value":"myAuthor-item1","ValueType":"Edm.String"},
          {"Key":"FileType","Value":"docx","ValueType":"Edm.String"},
          {"Key":"OriginalPath","Value":"myOriginalPath-item1","ValueType":"Edm.String"},
          {"Key":"PartitionId","Value":"00000000-0000-0000-0000-000000000000","ValueType":"Edm.Guid"},
          {"Key":"UrlZone","Value":"0","ValueType":"Edm.Int32"},
          {"Key":"Culture","Value":"en-US","ValueType":"Edm.String"},
          {"Key":"ResultTypeId","Value":"0","ValueType":"Edm.Int32"},
          {"Key":"IsDocument","Value":"true","ValueType":"Edm.Boolean"},
          {"Key":"RenderTemplateId","Value":"~sitecollection/_catalogs/masterpage/Display Templates/Search/Item_Default.js","ValueType":"Edm.String"}
        ]
      },
      {
        "Cells":[
          {"Key":"Rank","Value":"2","ValueType":"Edm.Double"},
          {"Key":"DocId","Value":"2","ValueType":"Edm.Int64"},
          {"Key":"Path","Value":"MyPath-item2","ValueType":"Edm.String"},
          {"Key":"Author","Value":"myAuthor-item2","ValueType":"Edm.String"},
          {"Key":"FileType","Value":"docx","ValueType":"Edm.String"},
          {"Key":"OriginalPath","Value":"myOriginalPath-item2","ValueType":"Edm.String"},
          {"Key":"PartitionId","Value":"00000000-0000-0000-0000-000000000000","ValueType":"Edm.Guid"},
          {"Key":"UrlZone","Value":"0","ValueType":"Edm.Int32"},
          {"Key":"Culture","Value":"en-US","ValueType":"Edm.String"},
          {"Key":"ResultTypeId","Value":"0","ValueType":"Edm.Int32"},
          {"Key":"IsDocument","Value":"true","ValueType":"Edm.Boolean"},
          {"Key":"RenderTemplateId","Value":"~sitecollection/_catalogs/masterpage/Display Templates/Search/Item_Default.js","ValueType":"Edm.String"}
        ]
      },
      {
        "Cells":[
          {"Key":"Rank","Value":"3","ValueType":"Edm.Double"},
          {"Key":"DocId","Value":"3","ValueType":"Edm.Int64"},
          {"Key":"Path","Value":"MyPath-item3","ValueType":"Edm.String"},
          {"Key":"Author","Value":"myAuthor-item3","ValueType":"Edm.String"},
          {"Key":"FileType","Value":"aspx","ValueType":"Edm.String"},
          {"Key":"OriginalPath","Value":"myOriginalPath-item3","ValueType":"Edm.String"},
          {"Key":"PartitionId","Value":"00000000-0000-0000-0000-000000000000","ValueType":"Edm.Guid"},
          {"Key":"UrlZone","Value":"0","ValueType":"Edm.Int32"},
          {"Key":"Culture","Value":"en-US","ValueType":"Edm.String"},
          {"Key":"ResultTypeId","Value":"0","ValueType":"Edm.Int32"},
          {"Key":"IsDocument","Value":"false","ValueType":"Edm.Boolean"},
          {"Key":"RenderTemplateId","Value":"~sitecollection/_catalogs/masterpage/Display Templates/Search/Item_Default.js","ValueType":"Edm.String"}
        ]
      }
    ];
  };
  let fakeRows:ResultTableRow[] = getFakeRows();
  let getQueryResult = (rows:ResultTableRow[],totalRows?:number):SearchResult => {
    returnArrayLength = totalRows ? totalRows : rows.length;

    return {
      "ElapsedTime": 83,
      "PrimaryQueryResult": {
        "CustomResults": [],
        "QueryId": "00000000-0000-0000-0000-000000000000",
        "QueryRuleId": "00000000-0000-0000-0000-000000000000",
        "RefinementResults": null,
        "RelevantResults": {
          "GroupTemplateId": null,
          "ItemTemplateId": null,
          "Properties": [
            {
              "Key": "GenerationId",
              "Value": "9223372036854775806",
              "ValueType": "Edm.Int64"
            }
          ],
          "ResultTitle": null,
          "ResultTitleUrl": null,
          "RowCount": rows.length,
          "Table": {
            "Rows": fakeRows
          },
          "TotalRows": returnArrayLength,
          "TotalRowsIncludingDuplicates": returnArrayLength
        },
        "SpecialTermResults": null
      },
      "Properties": [
        {
          "Key": "RowLimit",
          "Value": "10",
          "ValueType": "Edm.Int32"
        }
      ],
      "SecondaryQueryResults": [],
      "SpellingSuggestion": "",
      "TriggeredRules": []
    };
  }
  let getFakes = (opts:any) => {
    if (opts.url.toUpperCase().indexOf('QUERYTEXT=\'ISDOCUMENT:1\'') > -1) {
      let rows = fakeRows.filter(row => { 
        return row.Cells.filter(cell => { 
          return (cell.Key.toUpperCase() === "ISDOCUMENT" && cell.Value.toUpperCase() === "TRUE");
        }).length > 0; 
      });

      if(opts.url.toUpperCase().indexOf('ROWLIMIT=1') > -1) {
        if(opts.url.toUpperCase().indexOf('STARTROW=0') > -1) {
          return Promise.resolve(getQueryResult([rows[0]],2));
        }
        else if(opts.url.toUpperCase().indexOf('STARTROW=1') > -1) {
          return Promise.resolve(getQueryResult([rows[1]],2));
        }
        else {
          return Promise.resolve(getQueryResult([]));
        }
      }

      return Promise.resolve(getQueryResult(rows));
    }
    if (opts.url.toUpperCase().indexOf('QUERYTEXT=\'*\'') > -1) {
      let rows = fakeRows;
      if(opts.url.toUpperCase().indexOf('ROWLIMIT=1') > -1) {
        return Promise.resolve(getQueryResult([rows[0]]));;
      }
      if(opts.url.toUpperCase().indexOf('SOURCEID=\'6E71030E-5E16-4406-9BFF-9C1829843083\'') > -1) {
        return Promise.resolve(getQueryResult([rows[3]]));
      }
      if(opts.url.toUpperCase().indexOf('TRIMDUPLICATES=TRUE') > -1) {
        return Promise.resolve(getQueryResult([rows[2],rows[3]]));
      }
      return Promise.resolve(getQueryResult(rows));
    }
    returnArrayLength = 0;
    return Promise.reject('Invalid request');
  }

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => { return { FormDigestValue: 'abc' }; });
    trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
    });
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
    log = [];
    cmdInstance = {
      log: (msg: string) => {
        log.push(msg);
      }
    };
    auth.site = new Site();
    telemetry = null;
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.getAccessToken,
      auth.restoreAuth,
      request.get
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.SEARCH), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('calls telemetry', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert(trackEvent.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs correct telemetry event', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert.equal(telemetry.name, commands.SEARCH);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not logged in to a SharePoint site', (done) => {
    auth.site = new Site();
    auth.site.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Log in to a SharePoint Online site first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('executes search request', (done) => {
    stubAuth();

    sinon.stub(request, 'get').callsFake(getFakes);

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        output: 'json',
        debug: true,
        query: '*'
      }
    }, () => {
      try {
        assert.equal(returnArrayLength, 3);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
        Utils.restore(request.post);
      }
    });
  });

  it('executes search request with output option text', (done) => {
    stubAuth();

    sinon.stub(request, 'get').callsFake(getFakes);

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        output: 'text',
        debug: false,
        query: 'IsDocument:1'
      }
    }, () => {
      try {
        assert.equal(returnArrayLength, 2);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
        Utils.restore(request.post);
      }
    });
  });

  it('executes search request with output option text and \'allResults\'', (done) => {
    stubAuth();

    sinon.stub(request, 'get').callsFake(getFakes);

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        output: 'text',
        debug: false,
        query: 'IsDocument:1',
        allResults:true,
        rowLimit:1
      }
    }, () => {
      try {
        assert.equal(returnArrayLength, 2);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
        Utils.restore(request.post);
      }
    });
  });

  it('executes search request with trimDuplicates', (done) => {
    stubAuth();

    sinon.stub(request, 'get').callsFake(getFakes);

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        output: 'text',
        debug: false,
        query: '*',
        trimDuplicates:true
      }
    }, () => {
      try {
        assert.equal(returnArrayLength, 2);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
        Utils.restore(request.post);
      }
    });
  });

  it('executes search request with output option json and \'allResults\'', (done) => {
    stubAuth();

    sinon.stub(request, 'get').callsFake(getFakes);

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        output: 'json',
        debug: false,
        query: 'IsDocument:1',
        allResults:true,
        rowLimit:1
      }
    }, () => {
      try {
        assert.equal(returnArrayLength, 2);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
        Utils.restore(request.post);
      }
    });
  });

  it('executes search request with selectProperties', (done) => {
    stubAuth();

    sinon.stub(request, 'get').callsFake(getFakes);

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        output: 'text',
        debug: false,
        query: 'IsDocument:1',
        selectProperties: 'Path'
      }
    }, () => {
      try {
        assert.equal(returnArrayLength, 2);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
        Utils.restore(request.post);
      }
    });
  });

  it('executes search request with sourceId', (done) => {
    stubAuth();

    sinon.stub(request, 'get').callsFake(getFakes);

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        output: 'text',
        debug: false,
        query: '*',
        sourceId: '6e71030e-5e16-4406-9bff-9c1829843083'
      }
    }, () => {
      try {
        assert.equal(returnArrayLength, 1);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
        Utils.restore(request.post);
      }
    });
  });

  it('executes search request with rowLimits defined', (done) => {
    stubAuth();

    sinon.stub(request, 'get').callsFake(getFakes);

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        output: 'text',
        debug: true,
        query: '*',
        rowLimit: 1
      }
    }, () => {
      try {
        assert.equal(returnArrayLength, 1);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
        Utils.restore(request.post);
      }
    });
  });

  it('fails validation if the sourceId is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { 
        sourceId: '123',
        query:'*'
      } 
    });
    assert.notEqual(actual, true);
  });

  it('passes validation if the sourceId is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { 
        sourceId: '1caf7dcd-7e83-4c3a-94f7-932a1299c844',
        query:'*'
      } 
    });
    assert.equal(actual, true);
  });

  it('command correctly handles reject request', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf('/_api/contextinfo') > -1) {
        return Promise.resolve({
          FormDigestValue: 'abc'
        });
      }

      return Promise.reject('Invalid request');
    });

    const err = 'Invalid request';
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/web/webs') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    cmdInstance.action({
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
      }
    }, (error?: any) => {
      try {
        assert.equal(JSON.stringify(error), JSON.stringify(new CommandError(err)));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          request.post,
          request.get
        ]);
      }
    });
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

  it('supports specifying query', () => {
    const options = (command.options() as CommandOption[]);
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<query>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('fails validation if the query option is not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notEqual(actual, true);
  });

  it('passes validation if all options are provided', () => {
    const actual = (command.validate() as CommandValidate)({ options: { query:'*' } });
    assert.equal(actual, true);
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
    assert(find.calledWith(commands.SEARCH));
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

  it('correctly handles lack of valid access token', (done) => {
    Utils.restore(auth.getAccessToken);
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.reject(new Error('Error getting access token')); });
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        webUrl: "https://contoso.sharepoint.com",
        debug: false
      }
    }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Error getting access token')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
}); 