import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./oauth2grant-list');

describe(commands.OAUTH2GRANT_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
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
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.OAUTH2GRANT_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['objectId', 'resourceId', 'scope']);
  });

  it('retrieves OAuth2 permission grants for the specified service principal (debug)', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/oauth2PermissionGrants?$filter=clientId eq '141f7648-0c71-4752-9cdb-c7d5305b7e68'`) > -1) {
        return Promise.resolve({
          value: [{
            "clientId": "cd4043e7-b749-420b-bd07-aa7c3912ed22",
            "consentType": "AllPrincipals",
            "principalId": null,
            "resourceId": "1c444f1a-bba3-42f2-999f-4106c5b1c20c",
            "scope": "Group.ReadWrite.All"
          },
          {
            "clientId": "cd4043e7-b749-420b-bd07-aa7c3912ed22",
            "consentType": "AllPrincipals",
            "principalId": null,
            "resourceId": "dcf25ef3-e2df-4a77-839d-6b7857a11c78",
            "scope": "MyFiles.Read"
          }]
        });

      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: true, spObjectId: '141f7648-0c71-4752-9cdb-c7d5305b7e68' } });
    assert(loggerLogSpy.calledWith([{
      "clientId": "cd4043e7-b749-420b-bd07-aa7c3912ed22",
      "consentType": "AllPrincipals",
      "principalId": null,
      "resourceId": "1c444f1a-bba3-42f2-999f-4106c5b1c20c",
      "scope": "Group.ReadWrite.All"
    },
    {
      "clientId": "cd4043e7-b749-420b-bd07-aa7c3912ed22",
      "consentType": "AllPrincipals",
      "principalId": null,
      "resourceId": "dcf25ef3-e2df-4a77-839d-6b7857a11c78",
      "scope": "MyFiles.Read"
    }]));
  });

  it('retrieves OAuth2 permission grants for the specified service principal', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/oauth2PermissionGrants?$filter=clientId eq '141f7648-0c71-4752-9cdb-c7d5305b7e68'`) > -1) {
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

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { spObjectId: '141f7648-0c71-4752-9cdb-c7d5305b7e68' } });
    assert(loggerLogSpy.calledWith([{
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
  });

  it('outputs all properties when output is JSON', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/oauth2PermissionGrants?$filter=clientId eq '141f7648-0c71-4752-9cdb-c7d5305b7e68'`) > -1) {
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

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { spObjectId: '141f7648-0c71-4752-9cdb-c7d5305b7e68', output: 'json' } });
    assert(loggerLogSpy.calledWith([{
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
  });

  it('correctly handles no OAuth2 permission grants for the specified service principal found', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/oauth2PermissionGrants?$filter=clientId eq '141f7648-0c71-4752-9cdb-c7d5305b7e68'`) > -1) {
        return Promise.resolve({
          value: []
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { spObjectId: '141f7648-0c71-4752-9cdb-c7d5305b7e68' } });
    assert(loggerLogSpy.notCalled);
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'get').callsFake(() => {
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

    await assert.rejects(command.action(logger, { options: { spObjectId: 'b2307a39-e878-458b-bc90-03bc578531d6' } } as any),
      new CommandError(`Resource '' does not exist or one of its queried reference-property objects are not present`));
  });

  it('fails validation if the spObjectId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { spObjectId: '123' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the spObjectId option specified', async () => {
    const actual = await command.validate({ options: { spObjectId: '6a7b1395-d313-4682-8ed4-65a6265a6320' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('supports specifying spObjectId', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--spObjectId') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});
