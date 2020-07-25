import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./sitescript-get');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import auth from '../../../../Auth';

describe(commands.SITESCRIPT_GET, () => {
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
    assert.strictEqual(command.name.startsWith(commands.SITESCRIPT_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets information about the specified site script', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScriptMetadata`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b'
        })) {
        return Promise.resolve({
          "Content": JSON.stringify({
            "$schema": "schema.json",
            "actions": [
              {
                "verb": "applyTheme",
                "themeName": "Contoso Theme"
              }
            ],
            "bindata": {},
            "version": 1
          }),
          "Description": "My contoso script",
          "Id": "0f27a016-d277-4bb4-b3c3-b5b040c9559b",
          "Title": "Contoso",
          "Version": 1
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "Content": JSON.stringify({
            "$schema": "schema.json",
            "actions": [
              {
                "verb": "applyTheme",
                "themeName": "Contoso Theme"
              }
            ],
            "bindata": {},
            "version": 1
          }),
          "Description": "My contoso script",
          "Id": "0f27a016-d277-4bb4-b3c3-b5b040c9559b",
          "Title": "Contoso",
          "Version": 1
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets information about the specified site script (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScriptMetadata`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b'
        })) {
        return Promise.resolve({
          "Content": JSON.stringify({
            "$schema": "schema.json",
            "actions": [
              {
                "verb": "applyTheme",
                "themeName": "Contoso Theme"
              }
            ],
            "bindata": {},
            "version": 1
          }),
          "Description": "My contoso script",
          "Id": "0f27a016-d277-4bb4-b3c3-b5b040c9559b",
          "Title": "Contoso",
          "Version": 1
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "Content": JSON.stringify({
            "$schema": "schema.json",
            "actions": [
              {
                "verb": "applyTheme",
                "themeName": "Contoso Theme"
              }
            ],
            "bindata": {},
            "version": 1
          }),
          "Description": "My contoso script",
          "Id": "0f27a016-d277-4bb4-b3c3-b5b040c9559b",
          "Title": "Contoso",
          "Version": 1
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when site script not found', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'File Not Found.' } } } });
    });

    cmdInstance.action({ options: { debug: false, id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('File Not Found.')));
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

  it('supports specifying id', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--id') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if the id is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the id is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '2c1ba4c4-cd9b-4417-832f-92a34bc34b2a' } });
    assert.strictEqual(actual, true);
  });
});