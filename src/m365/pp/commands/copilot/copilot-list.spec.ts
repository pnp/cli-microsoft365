import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './copilot-list.js';
import { accessToken } from '../../../../utils/accessToken.js';

describe(commands.COPILOT_LIST, () => {
  const envUrl = "https://contoso-dev.api.crm4.dynamics.com";
  const validEnvironment = '4be50206-9576-4237-8b17-38d8aadfaa36';
  const fetchXml: string = `
      <fetch mapping='logical' version='1.0' >
        <entity name='bot'>
          <attribute name='accesscontrolpolicy' alias='accessControlPolicy' />,
          <attribute name='applicationmanifestinformation' alias='applicationManifestInformation' />,
          <attribute name='authenticationmode' alias='authenticationMode' />,
          <attribute name='authenticationtrigger' alias='authenticationTrigger' />,
          <attribute name='authorizedsecuritygroupids' alias='authorizedSecurityGroupIds' />,
          <attribute name='componentidunique' alias='componentIdUnique' />,
          <attribute name='componentstate' alias='componentState' />,
          <attribute name='configuration' alias='configuration' />,
          <attribute name='createdon' alias='createdOn' />,
          <attribute name='importsequencenumber' alias='importSequenceNumber' />,
          <attribute name='ismanaged' alias='isManaged' />,
          <attribute name='language' alias='language' />,
          <attribute name='modifiedon' alias='botModifiedOn' />,
          <attribute name='overriddencreatedon' alias='overriddenCreatedOn' />,
          <attribute name='overwritetime' alias='overwriteTime' />,
          <attribute name='iconbase64' alias='iconBase64' />,
          <attribute name='publishedon' alias='publishedOn' />,
          <attribute name='schemaname' alias='schemaName' />,
          <attribute name='solutionid' alias='solutionId' />,
          <attribute name='statecode' alias='stateCode' />,
          <attribute name='statuscode' alias='statusCode' />,
          <attribute name='timezoneruleversionnumber' alias='timezoneRuleVersionNumber' />,
          <attribute name='utcconversiontimezonecode' alias='utcConversionTimezoneCode' />,
          <attribute name='versionnumber' alias='versionNumber' />,
          <attribute name='name' alias='name' />,
          <attribute name='botid' alias='cdsBotId' />,
          <attribute name='ownerid' alias='ownerId' />,
          <attribute name='synchronizationstatus' alias='synchronizationStatus' />
          <link-entity name='systemuser' to='ownerid' from='systemuserid' link-type='inner' >
            <attribute name='fullname' alias='owner' />
          </link-entity>
          <link-entity name='systemuser' to='modifiedby' from='systemuserid' link-type='inner' >
            <attribute name='fullname' alias='botModifiedBy' />
          </link-entity>
        </entity>
      </fetch>
    `;

  const copilotResponse: any = {
    "value": [
      {
        "language": 1033,
        "botid": "23f5f586-97fd-43d5-95eb-451c9797a53d",
        "authenticationTrigger": 0,
        "stateCode": 0,
        "createdOn": "2022-11-19T10:42:22Z",
        "cdsBotId": "23f5f586-97fd-43d5-95eb-451c9797a53d",
        "schemaName": "new_bot_23f5f58697fd43d595eb451c9797a53d",
        "ownerId": "5fa787c1-1c4d-ed11-bba1-000d3a2caf7f",
        "botModifiedOn": "2022-11-19T20:19:57Z",
        "solutionId": "fd140aae-4df4-11dd-bd17-0019b9312238",
        "isManaged": false,
        "versionNumber": 1429641,
        "timezoneRuleVersionNumber": 0,
        "name": "CLI Copilot",
        "statusCode": 1,
        "owner": "Doe, John",
        "overwriteTime": "1900-01-01T00:00:00Z",
        "componentState": 0,
        "componentIdUnique": "cdcd6496-e25d-4ad1-91cf-3f4d547fdd23",
        "authenticationMode": 1,
        "botModifiedBy": "Doe, John",
        "accessControlPolicy": 0,
        "publishedOn": "2022-11-19T19:19:53Z"
      }
    ]
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(accessToken, 'assertAccessTokenType').returns();
    auth.connection.active = true;
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      powerPlatform.getDynamicsInstanceApiUrl
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.COPILOT_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['name', 'botid', 'publishedOn', 'createdOn', 'botModifiedOn']);
  });

  it('retrieves copilot bots', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/bots?fetchXml=${fetchXml}`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return copilotResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, environmentName: validEnvironment } });
    assert(loggerLogSpy.calledWith(copilotResponse.value));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.1/bots?fetchXml=${fetchXml}`) {
        if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
          throw {
            error: {
              'odata.error': {
                code: '-1, InvalidOperationException',
                message: {
                  value: `Resource '' does not exist or one of its queried reference-property objects are not present`
                }
              }
            }
          };
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { environmentName: validEnvironment } } as any),
      new CommandError(`Resource '' does not exist or one of its queried reference-property objects are not present`));
  });

  it('defines correct alias', () => {
    const alias = command.alias();
    assert.deepStrictEqual(alias, [commands.CHATBOT_LIST]);
  });

});
