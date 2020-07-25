import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./field-remove');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.FIELD_REMOVE, () => {
  let log: any[];
  let cmdInstance: any;
  let requests: any[];
  let promptOptions: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      },
      prompt: (options: any, cb: (result: { continue: boolean }) => void) => {
        promptOptions = options;
        cb({ continue: false });
      }
    };
    requests = [];
  });

  afterEach(() => {
    Utils.restore([
      request.post,
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
    assert.strictEqual(command.name.startsWith(commands.FIELD_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('prompts before removing field when confirmation argument not passed (id)', (done) => {
    cmdInstance.action({ options: { debug: false, id: 'b2307a39-e878-458b-bc90-03bc578531d6', webUrl: 'https://contoso.sharepoint.com' } }, () => {
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      try {
        assert(promptIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('prompts before removing field when confirmation argument not passed (title)', (done) => {
    cmdInstance.action({ options: { debug: false, fieldTitle: 'myfield1', webUrl: 'https://contoso.sharepoint.com' } }, () => {
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      try {
        assert(promptIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('prompts before removing list column when confirmation argument not passed', (done) => {
    cmdInstance.action({ options: { debug: false, fieldTitle: 'myfield1', webUrl: 'https://contoso.sharepoint.com', listTitle: 'My List' } }, () => {
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      try {
        assert(promptIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts removing field when prompt not confirmed', (done) => {
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    };
    cmdInstance.action({ options: { debug: false, id: 'b2307a39-e878-458b-bc90-03bc578531d6', webUrl: 'https://contoso.sharepoint.com' } }, () => {
      try {
        assert(requests.length === 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts removing field when prompt not confirmed and passing the group parameter', (done) => {
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    };
    cmdInstance.action({ options: { debug: false, group: 'MyGroup', webUrl: 'https://contoso.sharepoint.com' } }, () => {
      try {
        assert(requests.length === 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the field when prompt confirmed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/fields(guid'`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };
    cmdInstance.action({ options: { debug: false, id: 'b2307a39-e878-458b-bc90-03bc578531d6', webUrl: 'https://contoso.sharepoint.com' } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web/fields/getbyid('`) > -1 &&
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
      finally {
        Utils.restore(request.post);
      }

    });
  });

  it('command correctly handles field get reject request', (done) => {
    const err = 'Invalid request';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/fields/getbyinternalnameortitle(') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    const actionTitle: string = 'field1';

    cmdInstance.action({
      options: {
        debug: true,
        title: actionTitle,
        webUrl: 'https://contoso.sharepoint.com',
        confirm: true
      }
    }, (error?: any) => {
      try {
        assert.strictEqual(JSON.stringify(error), JSON.stringify(new CommandError(err)));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          request.post
        ]);
      }
    });
  });

  it('uses correct API url when id option is passed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/fields/getbyid(\'') > -1) {
        return Promise.resolve('Correct Url')
      }

      return Promise.reject('Invalid request');
    });

    const actionId: string = '0CD891EF-AFCE-4E55-B836-FCE03286CCCF';

    cmdInstance.action({
      options: {
        debug: false,
        id: actionId,
        webUrl: 'https://contoso.sharepoint.com',
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

  it('calls the correct remove url when id and list url specified', (done) => {
    const getStub = sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        return Promise.resolve({
          "Id": "03e45e84-1992-4d42-9116-26f756012634"
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', id: '03e45e84-1992-4d42-9116-26f756012634', listUrl: 'Lists/Events', confirm: true } }, () => {
      try {
        assert.strictEqual(getStub.lastCall.args[0].url, 'https://contoso.sharepoint.com/sites/portal/_api/web/GetList(\'%2Fsites%2Fportal%2FLists%2FEvents\')/fields/getbyid(\'03e45e84-1992-4d42-9116-26f756012634\')');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('calls group and deletes two fields and asks for confirmation', (done) => {
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };

    const getStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web/GetList(\'%2Fsites%2Fportal%2FLists%2FEvents\')/fields`) {
        return Promise.resolve({ "value": [{
          "Id": "03e45e84-1992-4d42-9116-26f756012634",
          "Group": "MyGroup"
        },
        {
          "Id": "03e45e84-1992-4d42-9116-26f756012635",
          "Group": "MyGroup"
        },
        {
          "Id": "03e45e84-1992-4d42-9116-26f756012636",
          "Group": "DifferentGroup"
        }]});
      }
      return Promise.reject('Invalid request');
    });

    const deletion = sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/GetList(\'%2Fsites%2Fportal%2FLists%2FEvents\')/fields/getbyid(\'03e45e84-1992-4d42-9116-26f756012634\')`) > -1) {
        return Promise.resolve({
          "Id": "03e45e84-1992-4d42-9116-26f756012634"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/GetList(\'%2Fsites%2Fportal%2FLists%2FEvents\')/fields/getbyid(\'03e45e84-1992-4d42-9116-26f756012635\')`) > -1) {
        return Promise.resolve({
          "Id": "03e45e84-1992-4d42-9116-26f756012635"
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', group: 'MyGroup', listUrl: '/sites/portal/Lists/Events' } }, () => {
      try {
        assert(getStub.called);
        assert.strictEqual(deletion.firstCall.args[0].url, 'https://contoso.sharepoint.com/sites/portal/_api/web/GetList(\'%2Fsites%2Fportal%2FLists%2FEvents\')/fields/getbyid(\'03e45e84-1992-4d42-9116-26f756012634\')');
        assert.strictEqual(deletion.secondCall.args[0].url, 'https://contoso.sharepoint.com/sites/portal/_api/web/GetList(\'%2Fsites%2Fportal%2FLists%2FEvents\')/fields/getbyid(\'03e45e84-1992-4d42-9116-26f756012635\')');
        assert.strictEqual(deletion.callCount, 2);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('calls group and deletes two fields', (done) => {
    const getStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web/fields`) {
        return Promise.resolve({ "value": [{
          "Id": "03e45e84-1992-4d42-9116-26f756012634",
          "Group": "MyGroup"
        },
        {
          "Id": "03e45e84-1992-4d42-9116-26f756012635",
          "Group": "MyGroup"
        },
        {
          "Id": "03e45e84-1992-4d42-9116-26f756012636",
          "Group": "DifferentGroup"
        }]});
      }
      return Promise.reject('Invalid request');
    });

    const deletion = sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/fields/getbyid(\'03e45e84-1992-4d42-9116-26f756012634\')`) > -1) {
        return Promise.resolve({
          "Id": "03e45e84-1992-4d42-9116-26f756012634"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/fields/getbyid(\'03e45e84-1992-4d42-9116-26f756012635\')`) > -1) {
        return Promise.resolve({
          "Id": "03e45e84-1992-4d42-9116-26f756012635"
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', group: 'MyGroup', confirm: true } }, () => {
      try {
        assert(getStub.called);
        assert.strictEqual(deletion.firstCall.args[0].url, 'https://contoso.sharepoint.com/sites/portal/_api/web/fields/getbyid(\'03e45e84-1992-4d42-9116-26f756012634\')');
        assert.strictEqual(deletion.secondCall.args[0].url, 'https://contoso.sharepoint.com/sites/portal/_api/web/fields/getbyid(\'03e45e84-1992-4d42-9116-26f756012635\')');
        assert.strictEqual(deletion.callCount, 2);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('calls group and deletes no fields', (done) => {
    const getStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web/fields`) {
        return Promise.resolve({ "value": [{
          "Id": "03e45e84-1992-4d42-9116-26f756012634",
          "Group": "MyGroup"
        },
        {
          "Id": "03e45e84-1992-4d42-9116-26f756012635",
          "Group": "MyGroup"
        },
        {
          "Id": "03e45e84-1992-4d42-9116-26f756012636",
          "Group": "DifferentGroup"
        }]});
      }
      return Promise.reject('Invalid request');
    });

    const deletion = sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/fields`) > -1) {
        return Promise.resolve({
          "Id": "03e45e84-1992-4d42-9116-26f756012634"
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', group: 'MyGroup1', confirm: true } }, () => {
      try {
        assert(getStub.called);
        assert(deletion.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles failure when get operation fails', (done) => {
    const err = 'Invalid request';
    
    const getStub = sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject(err);
    });

    const deletion = sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/fields/getbyid(\'03e45e84-1992-4d42-9116-26f756012635\')`) > -1) {
        return Promise.resolve({
          "Id": "03e45e84-1992-4d42-9116-26f756012635"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/fields/getbyid(\'03e45e84-1992-4d42-9116-26f756012634\')`) > -1) {
        return Promise.reject(err);
      }

      return Promise.reject(err);
    });

    cmdInstance.action({ options: { verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', group: 'MyGroup', confirm: true } }, (error?: any) => {
      try {
        assert(getStub.called);
        assert.strictEqual(JSON.stringify(error), JSON.stringify(new CommandError(err)));
        assert(deletion.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles failure when one deletion fails', (done) => {
    const err = 'Invalid request';
    
    const getStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web/fields`) {
        return Promise.resolve({ "value": [{
          "Id": "03e45e84-1992-4d42-9116-26f756012634",
          "Group": "MyGroup"
        },
        {
          "Id": "03e45e84-1992-4d42-9116-26f756012635",
          "Group": "MyGroup"
        },
        {
          "Id": "03e45e84-1992-4d42-9116-26f756012636",
          "Group": "DifferentGroup"
        }]});
      }
      return Promise.reject(err);
    });

    const deletion = sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/fields/getbyid(\'03e45e84-1992-4d42-9116-26f756012635\')`) > -1) {
        return Promise.resolve({
          "Id": "03e45e84-1992-4d42-9116-26f756012635"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/fields/getbyid(\'03e45e84-1992-4d42-9116-26f756012634\')`) > -1) {
        return Promise.reject(err);
      }

      return Promise.reject(err);
    });

    cmdInstance.action({ options: { verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', group: 'MyGroup', confirm: true } }, (error?: any) => {
      try {
        assert(getStub.called);
        assert.strictEqual(deletion.firstCall.args[0].url, 'https://contoso.sharepoint.com/sites/portal/_api/web/fields/getbyid(\'03e45e84-1992-4d42-9116-26f756012634\')');
        assert.strictEqual(deletion.callCount, 2);
        assert.strictEqual(JSON.stringify(error), JSON.stringify(new CommandError(err)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('calls the correct get url when field title and list title specified (verbose)', (done) => {
    const getStub = sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        return Promise.resolve({
          "Id": "03e45e84-1992-4d42-9116-26f756012634"
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', fieldTitle: 'Title', listTitle: 'Documents', confirm: true } }, () => {
      try {
        assert.strictEqual(getStub.lastCall.args[0].url, 'https://contoso.sharepoint.com/sites/portal/_api/web/lists/getByTitle(\'Documents\')/fields/getbyinternalnameortitle(\'Title\')');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('calls the correct get url when field title and list title specified', (done) => {
    const getStub = sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        return Promise.resolve({
          "Id": "03e45e84-1992-4d42-9116-26f756012634"
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', fieldTitle: 'Title', listTitle: 'Documents', confirm: true } }, () => {
      try {
        assert.strictEqual(getStub.lastCall.args[0].url, 'https://contoso.sharepoint.com/sites/portal/_api/web/lists/getByTitle(\'Documents\')/fields/getbyinternalnameortitle(\'Title\')');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('calls the correct get url when field title and list url specified', (done) => {
    const getStub = sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        return Promise.resolve({
          "Id": "03e45e84-1992-4d42-9116-26f756012634"
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', fieldTitle: 'Title', listId: '03e45e84-1992-4d42-9116-26f756012634', confirm: true } }, () => {
      try {
        assert.strictEqual(getStub.lastCall.args[0].url, 'https://contoso.sharepoint.com/sites/portal/_api/web/lists(guid\'03e45e84-1992-4d42-9116-26f756012634\')/fields/getbyinternalnameortitle(\'Title\')');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles site column not found', (done) => {
    const err = 'Invalid request';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/fields/getbyinternalnameortitle(') > -1) {
        return Promise.reject(err);
      }
      return Promise.reject('Invalid request');
    });
    const actionTitle: string = 'field1';

    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', fieldTitle: actionTitle, confirm: true } }, (error?: any) => {
      try {
        assert.strictEqual(JSON.stringify(error), JSON.stringify(new CommandError(err)));
        done();

      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles list column not found', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists/getByTitle('Documents')/fields/getbyid(`) > -1) {
        return Promise.reject({
          error: {
            "odata.error": {
              "code": "-2147024809, System.ArgumentException",
              "message": {
                "lang": "en-US",
                "value": "Invalid field name. {03e45e84-1992-4d42-9116-26f756012634}  /sites/portal/Shared Documents"
              }
            }
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', id: '03e45e84-1992-4d42-9116-26f756012634', listTitle: 'Documents', confirm: true } }, (error?: any) => {
      try {
        assert.strictEqual(JSON.stringify(error), JSON.stringify(new CommandError('Invalid field name. {03e45e84-1992-4d42-9116-26f756012634}  /sites/portal/Shared Documents')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles list not found', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists/getByTitle('Documents')/fields/getbyid(`) > -1) {
        return Promise.reject({
          error: {
            "odata.error": {
              "code": "-1, System.ArgumentException",
              "message": {
                "lang": "en-US",
                "value": "List 'Documents' does not exist at site with URL 'https://contoso.sharepoint.com/sites/portal'."
              }
            }
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', id: '03e45e84-1992-4d42-9116-26f756012634', listTitle: 'Documents', confirm: true } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("List 'Documents' does not exist at site with URL 'https://contoso.sharepoint.com/sites/portal'.")));
        done();
      }
      catch (e) {
        done(e);
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

  it('fails validation if both id and fieldTitle options are not passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', confirm: true, listTitle: 'Documents' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'foo.com', id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } });
    assert(actual);
  });

  it('fails validation if the field ID option is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '12345' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the field ID option is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } });
    assert(actual);
  });

  it('fails validation if the list ID is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listId: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        id: 'BC448D63-484F-49C5-AB8C-96B14AA68D50',
        confirm: true
      }
    });
    assert.strictEqual(actual, true);
  });
});