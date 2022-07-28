import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./roledefinition-get');

describe(commands.ROLEDEFINITION_GET, () => {
  let log: any[];
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
      request.get
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
    assert.strictEqual(command.name.startsWith(commands.ROLEDEFINITION_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = command.validate({ options: { webUrl: 'foo', id: 1 } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 1 } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if the id option is not a number', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'aaa' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id option is a number', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 1 } });
    assert.strictEqual(actual, true);
  });

  it('handles reject request correctly', (done) => {
    const err = 'request rejected';
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/roledefinitions(1)') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        id: 1
      }
    }, (error?: any) => {
      try {
        assert.strictEqual(JSON.stringify(error), JSON.stringify(new CommandError(err)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets role definition from web by id', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/roledefinitions(1)') > -1) {
        return Promise.resolve(
          {
            "BasePermissions": {
              "High": "432",
              "Low": "1012866047"
            },
            "Description": "Can view, add, update, delete, approve, and customize.",
            "Hidden": false,
            "Id": 1073741828,
            "Name": "Design",
            "Order": 32,
            "RoleTypeKind": 4
          });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        id: 1
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith(
          {
            "BasePermissions": {
              "High": "432",
              "Low": "1012866047"
            },
            "Description": "Can view, add, update, delete, approve, and customize.",
            "Hidden": false,
            "Id": 1073741828,
            "Name": "Design",
            "Order": 32,
            "RoleTypeKind": 4,
            "BasePermissionsValue": [
              "ViewListItems",
              "AddListItems",
              "EditListItems",
              "DeleteListItems",
              "ApproveItems",
              "OpenItems",
              "ViewVersions",
              "DeleteVersions",
              "CancelCheckout",
              "ManagePersonalViews",
              "ManageLists",
              "ViewFormPages",
              "Open",
              "ViewPages",
              "AddAndCustomizePages",
              "ApplyThemeAndBorder",
              "ApplyStyleSheets",
              "CreateSSCSite",
              "BrowseDirectories",
              "BrowseUserInfo",
              "AddDelPrivateWebParts",
              "UpdatePersonalWebParts",
              "UseClientIntegration",
              "UseRemoteAPIs",
              "CreateAlerts",
              "EditMyUserInfo"
            ],
            "RoleTypeKindValue": "WebDesigner"
          }
        ));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});