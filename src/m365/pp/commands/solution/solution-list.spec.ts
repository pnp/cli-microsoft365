import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./solution-list');

describe(commands.SOLUTION_LIST, () => {
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
      },
      {
        "solutionid": "fd140aaf-4df4-11dd-bd17-0019b9312238",
        "uniquename": "Default",
        "version": "1.0",
        "installedon": "2021-10-01T21:29:10Z",
        "solutionpackageversion": null,
        "friendlyname": "Default Solution",
        "versionnumber": 860055,
        "publisherid": {
          "friendlyname": "Default Publisher for org6633049a",
          "publisherid": "d21aab71-79e7-11dd-8874-00188b01e34f"
        }
      },
      {
        "solutionid": "d2ba05a9-3e7c-4974-b823-a752e69de519",
        "uniquename": "msdyn_ContextualHelpAnchor",
        "version": "1.0.0.22",
        "installedon": "2021-10-01T23:35:28Z",
        "solutionpackageversion": "9.0",
        "friendlyname": "Contextual Help Base",
        "versionnumber": 860175,
        "publisherid": {
          "friendlyname": "Dynamics 365",
          "publisherid": "858f43d2-6462-4e83-8272-b61d417ee53b"
        }
      }
    ]
  };
  const solutionResponseText: any = [
    {
      "uniquename": "Crc00f1",
      "version": "1.0.0.0",
      "publisher": "CDS Default Publisher"
    },
    {
      "uniquename": "Default",
      "version": "1.0",
      "publisher": "Default Publisher for org6633049a"
    },
    {
      "uniquename": "msdyn_ContextualHelpAnchor",
      "version": "1.0.0.22",
      "publisher": "Dynamics 365"
    }
  ];
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
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
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SOLUTION_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['uniquename', 'version', 'publisher']);
  });

  it('retrieves solutions from power platform environment', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/4be50206-9576-4237-8b17-38d8aadfaa36?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return envResponse;
        }
      }

      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/solutions?$filter=isvisible eq true&$expand=publisherid($select=friendlyname)&$select=solutionid,uniquename,version,publisherid,installedon,solutionpackageversion,friendlyname,versionnumber&api-version=9.1`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return solutionResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, environment: '4be50206-9576-4237-8b17-38d8aadfaa36' } });
    assert(loggerLogSpy.calledWith(solutionResponse.value));
  });

  it('retrieves solutions from power platform environment in format json', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/4be50206-9576-4237-8b17-38d8aadfaa36?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return envResponse;
        }
      }

      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/solutions?$filter=isvisible eq true&$expand=publisherid($select=friendlyname)&$select=solutionid,uniquename,version,publisherid,installedon,solutionpackageversion,friendlyname,versionnumber&api-version=9.1`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return solutionResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, environment: '4be50206-9576-4237-8b17-38d8aadfaa36', output: 'json' } });
    assert(loggerLogSpy.calledWith(solutionResponse.value));
  });

  it('retrieves solutions from power platform environment in format json as admin', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/4be50206-9576-4237-8b17-38d8aadfaa36?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return envResponse;
        }
      }

      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/solutions?$filter=isvisible eq true&$expand=publisherid($select=friendlyname)&$select=solutionid,uniquename,version,publisherid,installedon,solutionpackageversion,friendlyname,versionnumber&api-version=9.1`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return solutionResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, environment: '4be50206-9576-4237-8b17-38d8aadfaa36', asAdmin: true, output: 'json' } });
    assert(loggerLogSpy.calledWith(solutionResponse.value));
  });


  it('retrieves solutions from power platform environment in format text', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/4be50206-9576-4237-8b17-38d8aadfaa36?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return envResponse;
        }
      }

      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/solutions?$filter=isvisible eq true&$expand=publisherid($select=friendlyname)&$select=solutionid,uniquename,version,publisherid,installedon,solutionpackageversion,friendlyname,versionnumber&api-version=9.1`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return solutionResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, environment: '4be50206-9576-4237-8b17-38d8aadfaa36', output: 'text' } });
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
      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.0/solutions?$filter=isvisible eq true&$expand=publisherid($select=friendlyname)&$select=solutionid,uniquename,version,publisherid,installedon,solutionpackageversion,friendlyname,versionnumber&api-version=9.1`)) {
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

    await assert.rejects(command.action(logger, { options: { debug: false, environment: '4be50206-9576-4237-8b17-38d8aadfaa36' } } as any),
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
