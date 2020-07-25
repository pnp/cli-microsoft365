import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./oauth2grant-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.OAUTH2GRANT_LIST, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

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
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
  });

  afterEach(() => {
    Utils.restore([
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
    assert.strictEqual(command.name.startsWith(commands.OAUTH2GRANT_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves OAuth2 permission grants for the specified service principal (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/myorganization/oauth2PermissionGrants?api-version=1.6&$filter=clientId eq '141f7648-0c71-4752-9cdb-c7d5305b7e68'`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            value: [{
              "clientId": "cd4043e7-b749-420b-bd07-aa7c3912ed22",
              "consentType": "AllPrincipals",
              "expiryTime": "9999-12-31T23:59:59.9999999",
              "objectId": "50NAzUm3C0K9B6p8ORLtIhpPRByju_JCmZ9BBsWxwgw",
              "principalId": null,
              "resourceId": "1c444f1a-bba3-42f2-999f-4106c5b1c20c",
              "scope": "Group.ReadWrite.All",
              "startTime": "0001-01-01T00:00:00"
            },
            {
              "clientId": "cd4043e7-b749-420b-bd07-aa7c3912ed22",
              "consentType": "AllPrincipals",
              "expiryTime": "9999-12-31T23:59:59.9999999",
              "objectId": "50NAzUm3C0K9B6p8ORLtIvNe8tzf4ndKg51reFehHHg",
              "principalId": null,
              "resourceId": "dcf25ef3-e2df-4a77-839d-6b7857a11c78",
              "scope": "MyFiles.Read",
              "startTime": "0001-01-01T00:00:00"
            }]
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, clientId: '141f7648-0c71-4752-9cdb-c7d5305b7e68' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            objectId: '50NAzUm3C0K9B6p8ORLtIhpPRByju_JCmZ9BBsWxwgw',
            resourceId: '1c444f1a-bba3-42f2-999f-4106c5b1c20c',
            scope: 'Group.ReadWrite.All'
          },
          {
            objectId: '50NAzUm3C0K9B6p8ORLtIvNe8tzf4ndKg51reFehHHg',
            resourceId: 'dcf25ef3-e2df-4a77-839d-6b7857a11c78',
            scope: 'MyFiles.Read'
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves OAuth2 permission grants for the specified service principal', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/myorganization/oauth2PermissionGrants?api-version=1.6&$filter=clientId eq '141f7648-0c71-4752-9cdb-c7d5305b7e68'`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            value: [{
              "clientId": "cd4043e7-b749-420b-bd07-aa7c3912ed22",
              "consentType": "AllPrincipals",
              "expiryTime": "9999-12-31T23:59:59.9999999",
              "objectId": "50NAzUm3C0K9B6p8ORLtIhpPRByju_JCmZ9BBsWxwgw",
              "principalId": null,
              "resourceId": "1c444f1a-bba3-42f2-999f-4106c5b1c20c",
              "scope": "Group.ReadWrite.All",
              "startTime": "0001-01-01T00:00:00"
            },
            {
              "clientId": "cd4043e7-b749-420b-bd07-aa7c3912ed22",
              "consentType": "AllPrincipals",
              "expiryTime": "9999-12-31T23:59:59.9999999",
              "objectId": "50NAzUm3C0K9B6p8ORLtIvNe8tzf4ndKg51reFehHHg",
              "principalId": null,
              "resourceId": "dcf25ef3-e2df-4a77-839d-6b7857a11c78",
              "scope": "MyFiles.Read",
              "startTime": "0001-01-01T00:00:00"
            }]
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, clientId: '141f7648-0c71-4752-9cdb-c7d5305b7e68' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            objectId: '50NAzUm3C0K9B6p8ORLtIhpPRByju_JCmZ9BBsWxwgw',
            resourceId: '1c444f1a-bba3-42f2-999f-4106c5b1c20c',
            scope: 'Group.ReadWrite.All'
          },
          {
            objectId: '50NAzUm3C0K9B6p8ORLtIvNe8tzf4ndKg51reFehHHg',
            resourceId: 'dcf25ef3-e2df-4a77-839d-6b7857a11c78',
            scope: 'MyFiles.Read'
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('outputs all properties when output is JSON', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/myorganization/oauth2PermissionGrants?api-version=1.6&$filter=clientId eq '141f7648-0c71-4752-9cdb-c7d5305b7e68'`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            value: [{
              "clientId": "cd4043e7-b749-420b-bd07-aa7c3912ed22",
              "consentType": "AllPrincipals",
              "expiryTime": "9999-12-31T23:59:59.9999999",
              "objectId": "50NAzUm3C0K9B6p8ORLtIhpPRByju_JCmZ9BBsWxwgw",
              "principalId": null,
              "resourceId": "1c444f1a-bba3-42f2-999f-4106c5b1c20c",
              "scope": "Group.ReadWrite.All",
              "startTime": "0001-01-01T00:00:00"
            },
            {
              "clientId": "cd4043e7-b749-420b-bd07-aa7c3912ed22",
              "consentType": "AllPrincipals",
              "expiryTime": "9999-12-31T23:59:59.9999999",
              "objectId": "50NAzUm3C0K9B6p8ORLtIvNe8tzf4ndKg51reFehHHg",
              "principalId": null,
              "resourceId": "dcf25ef3-e2df-4a77-839d-6b7857a11c78",
              "scope": "MyFiles.Read",
              "startTime": "0001-01-01T00:00:00"
            }]
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, clientId: '141f7648-0c71-4752-9cdb-c7d5305b7e68', output: 'json' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([{
          "clientId": "cd4043e7-b749-420b-bd07-aa7c3912ed22",
          "consentType": "AllPrincipals",
          "expiryTime": "9999-12-31T23:59:59.9999999",
          "objectId": "50NAzUm3C0K9B6p8ORLtIhpPRByju_JCmZ9BBsWxwgw",
          "principalId": null,
          "resourceId": "1c444f1a-bba3-42f2-999f-4106c5b1c20c",
          "scope": "Group.ReadWrite.All",
          "startTime": "0001-01-01T00:00:00"
        },
        {
          "clientId": "cd4043e7-b749-420b-bd07-aa7c3912ed22",
          "consentType": "AllPrincipals",
          "expiryTime": "9999-12-31T23:59:59.9999999",
          "objectId": "50NAzUm3C0K9B6p8ORLtIvNe8tzf4ndKg51reFehHHg",
          "principalId": null,
          "resourceId": "dcf25ef3-e2df-4a77-839d-6b7857a11c78",
          "scope": "MyFiles.Read",
          "startTime": "0001-01-01T00:00:00"
        }]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no OAuth2 permission grants for the specified service principal found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/myorganization/oauth2PermissionGrants?api-version=1.6&$filter=clientId eq '141f7648-0c71-4752-9cdb-c7d5305b7e68'`) > -1) {
        if (opts.headers &&
          opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            value: []
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, clientId: '141f7648-0c71-4752-9cdb-c7d5305b7e68' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles API OData error', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject({
        error: {
          'odata.error': {
            code: '-1, InvalidOperationException',
            message: {
              value: `Resource '' does not exist or one of its queried reference-property objects are not present`
            }
          }
        }
      });
    });

    cmdInstance.action({ options: { debug: false, clientId: 'b2307a39-e878-458b-bc90-03bc578531d6' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Resource '' does not exist or one of its queried reference-property objects are not present`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if the clientId is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { clientId: '123' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the clientId option specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { clientId: '6a7b1395-d313-4682-8ed4-65a6265a6320' } });
    assert.strictEqual(actual, true);
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

  it('supports specifying clientId', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--clientId') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});