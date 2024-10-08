import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './website-webfile-list.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { cli } from '../../../../cli/cli.js';

describe(commands.WEBSITE_WEBFILE_LIST, () => {
  //#region Mocked Responses
  let commandInfo: CommandInfo;
  const validEnvironment = 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c';
  const validWebsiteId = '3bbc8102-8ee7-4dac-afbb-807cc5b6f9c2';
  const validWebsiteName = 'CLI 365 PowerPageSite';

  const envResponse: any = { "properties": { "linkedEnvironmentMetadata": { "instanceApiUrl": "https://contoso-dev.api.crm4.dynamics.com" } } };
  const webfilesResponse: any = {
    "value": [
      {
        "mspp_webfileid": "a1700435-b21a-48f0-83f4-f34411a26cbf",
        "mspp_name": "Site-mockup-1.png",
        "mspp_contentdisposition@OData.Community.Display.V1.FormattedValue": "inline",
        "mspp_contentdisposition": 756150000,
        "mspp_excludefromsearch@OData.Community.Display.V1.FormattedValue": "No",
        "mspp_excludefromsearch": false,
        "mspp_hiddenfromsitemap@OData.Community.Display.V1.FormattedValue": "No",
        "mspp_hiddenfromsitemap": false,
        "_mspp_parentpageid_value@Microsoft.Dynamics.CRM.associatednavigationproperty": "mspp_parentpageid",
        "_mspp_parentpageid_value@Microsoft.Dynamics.CRM.lookuplogicalname": "mspp_webpage",
        "_mspp_parentpageid_value@OData.Community.Display.V1.FormattedValue": "Home",
        "_mspp_parentpageid_value": "a3efb062-0691-4d3e-a9e6-43bfd2daeed6",
        "mspp_partialurl": "Site-mockup-1.png",
        "_mspp_publishingstateid_value@Microsoft.Dynamics.CRM.associatednavigationproperty": "mspp_publishingstateid",
        "_mspp_publishingstateid_value@Microsoft.Dynamics.CRM.lookuplogicalname": "mspp_publishingstate",
        "_mspp_publishingstateid_value@OData.Community.Display.V1.FormattedValue": "Published",
        "_mspp_publishingstateid_value": "0d0fa682-7ed7-419b-bc7a-25e8bc64b6e2",
        "_mspp_websiteid_value@Microsoft.Dynamics.CRM.associatednavigationproperty": "mspp_websiteid",
        "_mspp_websiteid_value@Microsoft.Dynamics.CRM.lookuplogicalname": "mspp_website",
        "_mspp_websiteid_value@OData.Community.Display.V1.FormattedValue": "Website",
        "_mspp_websiteid_value": "3bbc8102-8ee7-4dac-afbb-807cc5b6f9b3",
        "_mspp_createdby_value@Microsoft.Dynamics.CRM.lookuplogicalname": "systemuser",
        "_mspp_createdby_value@OData.Community.Display.V1.FormattedValue": "Shanthakumar",
        "_mspp_createdby_value": "66ece047-0e90-ee11-8179-000d3a37640e",
        "_mspp_modifiedby_value@Microsoft.Dynamics.CRM.lookuplogicalname": "systemuser",
        "_mspp_modifiedby_value@OData.Community.Display.V1.FormattedValue": "Shanthakumar",
        "_mspp_modifiedby_value": "66ece047-0e90-ee11-8179-000d3a37640e",
        "mspp_modifiedon@OData.Community.Display.V1.FormattedValue": "8/27/2024 2:37 PM",
        "mspp_modifiedon": "2024-08-27T09:07:22Z",
        "mspp_createdon@OData.Community.Display.V1.FormattedValue": "8/27/2024 2:37 PM",
        "mspp_createdon": "2024-08-27T09:07:16Z",
        "statecode": 0,
        "statuscode@OData.Community.Display.V1.FormattedValue": "Active",
        "statuscode": 1,
        "mspp_title": null,
        "mspp_displaydate": null,
        "mspp_displayorder": null,
        "mspp_summary": null,
        "mspp_cloudblobaddress": null,
        "mspp_modifiedbyusername": null,
        "mspp_alloworigin": null,
        "mspp_expirationdate": null,
        "mspp_createdbyipaddress": null,
        "mspp_createdbyusername": null,
        "mspp_modifiedbyipaddress": null,
        "mspp_releasedate": null,
        "_mspp_masterwebfileid_value": null
      }
    ]
  };
  //#endregion

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(accessToken, 'assertDelegatedAccessToken').returns();
    commandInfo = cli.getCommandInfo(command);
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
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.WEBSITE_WEBFILE_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if websiteId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        environmentName: validEnvironment,
        websiteId: 'Invalid GUID'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if required option websiteId specified', async () => {
    const actual = await command.validate({ options: { environmentName: validEnvironment, websiteId: validWebsiteId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required option websiteName specified', async () => {
    const actual = await command.validate({ options: { environmentName: validEnvironment, websiteName: validWebsiteName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['mspp_name', 'mspp_webfileid', 'mspp_summary', '_mspp_publishingstateid_value@OData.Community.Display.V1.FormattedValue']);
  });


  it('fails validation on unable to find website based on websiteName', async () => {
    const EmptyWebsiteResponse = {
      value: [
      ]
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/${validEnvironment}?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return envResponse;
        }
      }

      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.2/powerpagesites?$filter=name eq 'Invalid website'&$select=powerpagesiteid`)) {
        return EmptyWebsiteResponse;
      }
      throw `Invalid request`;
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        environmentName: `${validEnvironment}`,
        websiteName: 'Invalid website'
      }
    }), new CommandError(`The specified website 'Invalid website' does not exist.`));
  });


  it('retrieves webfiles', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/${validEnvironment}?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return envResponse;
        }
      }

      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.2/mspp_webfiles?$filter=_mspp_websiteid_value eq '${validWebsiteId}'`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return webfilesResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, environmentName: `${validEnvironment}`, websiteId: `${validWebsiteId}` } });
    assert(loggerLogSpy.calledWith(webfilesResponse.value));

  });

  it('retrieves webfiles based on website name as admin', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/${validEnvironment}?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return envResponse;
        }
      }

      if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.2/powerpagesites?$filter=name eq '${validWebsiteName}'&$select=powerpagesiteid`) {
        return {
          "value": [
            {
              "powerpagesiteid": validWebsiteId
            }
          ]
        };
      }

      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.2/mspp_webfiles?$filter=_mspp_websiteid_value eq '${validWebsiteId}'`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return webfilesResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, environmentName: validEnvironment, websiteName: validWebsiteName, asAdmin: true } });
    assert(loggerLogSpy.calledWith(webfilesResponse.value));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/${validEnvironment}?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return envResponse;
        }
      }

      throw `Resource '' does not exist or one of its queried reference-property objects are not present`;
    });

    await assert.rejects(command.action(logger, { options: { environmentName: validEnvironment } }), new CommandError("Resource '' does not exist or one of its queried reference-property objects are not present"));
  });
});
