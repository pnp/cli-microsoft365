import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
// import { Cli } from '../../../../cli/Cli';
// import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./solution-get');

describe(commands.SOLUTION_GET, () => {
  // let commandInfo: CommandInfo;
  // const validSolutionId = ' ee62fd63-e49e-4c09-80de-8fae1b9a427e';
  // const validSolutionName = 'Solution name';
  // const validEnvironmentId = 'Default-0cac6cda-2e04-4a3d-9c16-9c91470d7022';
  const envResponse: any = { "properties": { "linkedEnvironmentMetadata": { "instanceApiUrl": "https://contoso-dev.api.crm4.dynamics.com" } } };
  const solutionResponse: any = {
    "value": [
      {
        "solutionid": "00000001-0000-0000-0001-00000000009b",
        "uniquename": "Crc00f1",
        "version": "1.0.0.0",
        "installedon": "2021-10-01T21:54:14Z",
        "solutionpackageversion": null,
        "friendlyname": "Common Data Services Default Solution",
        "versionnumber": 860052,
        "publisherid": {
          "friendlyname": "CDS Default Publisher",
          "publisherid": "00000001-0000-0000-0000-00000000005a"
        }
      }
    ]
  };
  const solutionResponseText: any = {
    "uniquename": "Crc00f1",
    "version": "1.0.0.0",
    "publisher": "CDS Default Publisher"
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    auth.service.connected = true;
    // commandInfo = Cli.getCommandInfo(command);
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
      appInsights.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SOLUTION_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['uniquename', 'version', 'publisher']);
  });

  // it('fails validation when environment is used with both id and name', async () => {
  //   const actual = await command.validate({
  //     options: {
  //       environment: validEnvironmentId,
  //       id: validSolutionId,
  //       name: validSolutionName
  //     }
  //   }, commandInfo);
  //   assert.notStrictEqual(actual, false);
  // });

  it('retrieve specific solution from power platform environment with the name parameter', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/4be50206-9576-4237-8b17-38d8aadfaa36?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return envResponse;
        }
      }

      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/solutions?$filter=isvisible eq true and uniquename eq 'Default'&$expand=publisherid($select=friendlyname)&$select=solutionid,uniquename,version,publisherid,installedon,solutionpackageversion,friendlyname,versionnumber&api-version=9.1`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return solutionResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, environment: '4be50206-9576-4237-8b17-38d8aadfaa36', name: 'Default' } });
    assert(loggerLogSpy.calledWith(solutionResponse.value[0]));
  });

  it('retrieve specific solution from power platform environment with the name parameter in format json', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/4be50206-9576-4237-8b17-38d8aadfaa36?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return envResponse;
        }
      }

      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/solutions?$filter=isvisible eq true and uniquename eq 'Default'&$expand=publisherid($select=friendlyname)&$select=solutionid,uniquename,version,publisherid,installedon,solutionpackageversion,friendlyname,versionnumber&api-version=9.1`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return solutionResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, environment: '4be50206-9576-4237-8b17-38d8aadfaa36', name: 'Default', output: 'json' } });
    assert(loggerLogSpy.calledWith(solutionResponse.value[0]));
  });

  it('retrieve specific solution from power platform environment with the name parameter in format json as admin', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/4be50206-9576-4237-8b17-38d8aadfaa36?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return envResponse;
        }
      }

      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/solutions?$filter=isvisible eq true and uniquename eq 'Default'&$expand=publisherid($select=friendlyname)&$select=solutionid,uniquename,version,publisherid,installedon,solutionpackageversion,friendlyname,versionnumber&api-version=9.1`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return solutionResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, environment: '4be50206-9576-4237-8b17-38d8aadfaa36', name: 'Default', asAdmin: true, output: 'json' } });
    assert(loggerLogSpy.calledWith(solutionResponse.value[0]));
  });


  it('retrieve specific solution from power platform environment with name parameter in format text', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/4be50206-9576-4237-8b17-38d8aadfaa36?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return envResponse;
        }
      }

      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/solutions?$filter=isvisible eq true and uniquename eq 'Default'&$expand=publisherid($select=friendlyname)&$select=solutionid,uniquename,version,publisherid,installedon,solutionpackageversion,friendlyname,versionnumber&api-version=9.1`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return solutionResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, environment: '4be50206-9576-4237-8b17-38d8aadfaa36', name: 'Default', output: 'text' } });
    assert(loggerLogSpy.calledWith(solutionResponseText));
  });

  it('retrieve specific solution from power platform environment with the id parameter', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/4be50206-9576-4237-8b17-38d8aadfaa36?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return envResponse;
        }
      }

      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/solutions(ee62fd63-e49e-4c09-80de-8fae1b9a427e)?$expand=publisherid($select=friendlyname)&$select=solutionid,uniquename,version,publisherid,installedon,solutionpackageversion,friendlyname,versionnumber&api-version=9.1`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return solutionResponse.value[0];
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, environment: '4be50206-9576-4237-8b17-38d8aadfaa36', id: 'ee62fd63-e49e-4c09-80de-8fae1b9a427e' } });
    assert(loggerLogSpy.calledWith(solutionResponse.value[0]));
  });

  it('retrieve specific solution from power platform environment with the id parameter in format json', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/4be50206-9576-4237-8b17-38d8aadfaa36?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return envResponse;
        }
      }

      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/solutions(ee62fd63-e49e-4c09-80de-8fae1b9a427e)?$expand=publisherid($select=friendlyname)&$select=solutionid,uniquename,version,publisherid,installedon,solutionpackageversion,friendlyname,versionnumber&api-version=9.1`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return solutionResponse.value[0];
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, environment: '4be50206-9576-4237-8b17-38d8aadfaa36', id: 'ee62fd63-e49e-4c09-80de-8fae1b9a427e', output: 'json' } });
    assert(loggerLogSpy.calledWith(solutionResponse.value[0]));
  });

  it('retrieve specific solution from power platform environment with the id parameter in format json as admin', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/4be50206-9576-4237-8b17-38d8aadfaa36?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return envResponse;
        }
      }

      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/solutions(ee62fd63-e49e-4c09-80de-8fae1b9a427e)?$expand=publisherid($select=friendlyname)&$select=solutionid,uniquename,version,publisherid,installedon,solutionpackageversion,friendlyname,versionnumber&api-version=9.1`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return solutionResponse.value[0];
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, environment: '4be50206-9576-4237-8b17-38d8aadfaa36', id: 'ee62fd63-e49e-4c09-80de-8fae1b9a427e', asAdmin: true, output: 'json' } });
    assert(loggerLogSpy.calledWith(solutionResponse.value[0]));
  });


  it('retrieve specific solution from power platform environment with id parameter in format text', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/4be50206-9576-4237-8b17-38d8aadfaa36?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return envResponse;
        }
      }

      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/solutions(ee62fd63-e49e-4c09-80de-8fae1b9a427e)?$expand=publisherid($select=friendlyname)&$select=solutionid,uniquename,version,publisherid,installedon,solutionpackageversion,friendlyname,versionnumber&api-version=9.1`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return solutionResponse.value[0];
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, environment: '4be50206-9576-4237-8b17-38d8aadfaa36', id: 'ee62fd63-e49e-4c09-80de-8fae1b9a427e', output: 'text' } });
    assert(loggerLogSpy.calledWith(solutionResponseText));
  });

  it('correctly handles no environments', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/providers/Microsoft.BusinessAppPlatform/environments?api-version=2020-10-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return { value: [] };
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { debug: false } }),
      new CommandError(`The environment 'undefined' could not be retrieved. See the inner exception for more details: undefined`));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/4be50206-9576-4237-8b17-38d8aadfaa36?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return envResponse;
        }
      }
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/solutions?$filter=isvisible eq true and uniquename eq 'Default'&$expand=publisherid($select=friendlyname)&$select=solutionid,uniquename,version,publisherid,installedon,solutionpackageversion,friendlyname,versionnumber&api-version=9.1`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
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

    await assert.rejects(command.action(logger, { options: { debug: false, environment: '4be50206-9576-4237-8b17-38d8aadfaa36', name: 'Default' } } as any),
      new CommandError(`Resource '' does not exist or one of its queried reference-property objects are not present`));
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});
