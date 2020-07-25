import commands from '../../commands';
import Command, { CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./sitescript-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import auth from '../../../../Auth';

describe(commands.SITESCRIPT_LIST, () => {
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
    assert.strictEqual(command.name.startsWith(commands.SITESCRIPT_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('lists available site scripts', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScripts`) > -1) {
        return Promise.resolve({
          value: [
            {
              Content: null,
              Description: "description",
              Id: "19b0e1b2-e3d1-473f-9394-f08c198ef43e",
              Title: "script1",
              Version: 1
            },
            {
              Content: null,
              Description: "Contoso theme script description",
              Id: "449c0c6d-5380-4df2-b84b-622e0ac8ec24",
              Title: "Contoso theme script",
              Version: 1
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            Content: null,
            Description: "description",
            Id: "19b0e1b2-e3d1-473f-9394-f08c198ef43e",
            Title: "script1",
            Version: 1
          },
          {
            Content: null,
            Description: "Contoso theme script description",
            Id: "449c0c6d-5380-4df2-b84b-622e0ac8ec24",
            Title: "Contoso theme script",
            Version: 1
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists available site scripts (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScripts`) > -1) {
        return Promise.resolve({
          value: [
            {
              Content: null,
              Description: "description",
              Id: "19b0e1b2-e3d1-473f-9394-f08c198ef43e",
              Title: "script1",
              Version: 1
            },
            {
              Content: null,
              Description: "Contoso theme script description",
              Id: "449c0c6d-5380-4df2-b84b-622e0ac8ec24",
              Title: "Contoso theme script",
              Version: 1
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            Content: null,
            Description: "description",
            Id: "19b0e1b2-e3d1-473f-9394-f08c198ef43e",
            Title: "script1",
            Version: 1
          },
          {
            Content: null,
            Description: "Contoso theme script description",
            Id: "449c0c6d-5380-4df2-b84b-622e0ac8ec24",
            Title: "Contoso theme script",
            Version: 1
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no available site scripts', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScripts`) > -1) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
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

    cmdInstance.action({ options: { debug: false } }, (err?: any) => {
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
});