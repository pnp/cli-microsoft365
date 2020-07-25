import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./sitescript-add');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import auth from '../../../../Auth';

describe(commands.SITESCRIPT_ADD, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({ FormDigestValue: 'ABC' }));
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
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
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
  });

  afterEach(() => {
    Utils.restore([
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      (command as any).getRequestDigest,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });


  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SITESCRIPT_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds new site script', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteScript(Title=@title, Description=@description)?@title='Contoso'&@description='My%20contoso%20script'`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          abc: 'def'
        })) {
        return Promise.resolve({
          "Content": null,
          "Description": "My contoso script",
          "Id": "0f27a016-d277-4bb4-b3c3-b5b040c9559b",
          "Title": "Contoso",
          "Version": 0
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, title: 'Contoso', description: 'My contoso script', content: JSON.stringify({ "abc": "def" }) } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "Content": null,
          "Description": "My contoso script",
          "Id": "0f27a016-d277-4bb4-b3c3-b5b040c9559b",
          "Title": "Contoso",
          "Version": 0
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds new site script (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteScript(Title=@title, Description=@description)?@title='Contoso'&@description='My%20contoso%20script'`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          abc: 'def'
        })) {
        return Promise.resolve({
          "Content": null,
          "Description": "My contoso script",
          "Id": "0f27a016-d277-4bb4-b3c3-b5b040c9559b",
          "Title": "Contoso",
          "Version": 0
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, title: 'Contoso', description: 'My contoso script', content: JSON.stringify({ "abc": "def" }) } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "Content": null,
          "Description": "My contoso script",
          "Id": "0f27a016-d277-4bb4-b3c3-b5b040c9559b",
          "Title": "Contoso",
          "Version": 0
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('doesn\'t fail when description not passed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteScript(Title=@title, Description=@description)?@title='Contoso'&@description=''`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          abc: 'def'
        })) {
        return Promise.resolve({
          "Content": null,
          "Description": "",
          "Id": "0f27a016-d277-4bb4-b3c3-b5b040c9559b",
          "Title": "Contoso",
          "Version": 0
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, title: 'Contoso', description: '', content: JSON.stringify({ "abc": "def" }) } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "Content": null,
          "Description": "",
          "Id": "0f27a016-d277-4bb4-b3c3-b5b040c9559b",
          "Title": "Contoso",
          "Version": 0
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('escapes special characters in user input', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteScript(Title=@title, Description=@description)?@title='Contoso%20script'&@description='My%20contoso%20script'`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          abc: 'def'
        })) {
        return Promise.resolve({
          "Content": null,
          "Description": "My contoso script",
          "Id": "0f27a016-d277-4bb4-b3c3-b5b040c9559b",
          "Title": "Contoso script",
          "Version": 0
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, title: 'Contoso script', description: 'My contoso script', content: JSON.stringify({ "abc": "def" }) } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "Content": null,
          "Description": "My contoso script",
          "Id": "0f27a016-d277-4bb4-b3c3-b5b040c9559b",
          "Title": "Contoso script",
          "Version": 0
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles OData error when creating site script', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });
    });

    cmdInstance.action({ options: { debug: false, title: 'Contoso', content: JSON.stringify({}) } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying title', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--title') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying description', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--description') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying script content', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--content') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if script content is not a valid JSON string', () => {
    const actual = (command.validate() as CommandValidate)({ options: { title: 'Contoso', content: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when title specified and  script content is valid JSON', () => {
    const actual = (command.validate() as CommandValidate)({ options: { title: 'Contoso', content: JSON.stringify({}) } });
    assert.strictEqual(actual, true);
  });
});