import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './website-webfile-get.js';
import { accessToken } from '../../../../utils/accessToken.js';

describe(commands.WEBSITE_WEBFILE_GET, () => {
  let commandInfo: CommandInfo;
  //#region Mocked Responses
  const validEnvironment = '4be50206-9576-4237-8b17-38d8aadfaa36';
  const validId = '3a081d91-5ea8-40a7-8ac9-abbaa3fcb893';
  const validWebsiteId = '3bbc8102-8ee7-4dac-afbb-807cc5b6f9c2';
  const validWebsiteName = 'CLI 365 PowerPageSite';
  const envUrl = "https://contoso-dev.api.crm4.dynamics.com";
  const envResponse: any = { "properties": { "linkedEnvironmentMetadata": { "instanceApiUrl": "https://contoso-dev.api.crm4.dynamics.com" } } };
  const webfileResponse: any = {
    "value": [
      {
        "mspp_webfileid": "3a081d91-5ea8-40a7-8ac9-abbaa3fcb893",
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
        "_mspp_websiteid_value": "3bbc8102-8ee7-4dac-afbb-807cc5b6f9c2",
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
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName: string, defaultValue: any) => {
      if (settingName === 'prompt') {
        return false;
      }
      return defaultValue;
    });
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
      powerPlatform.getDynamicsInstanceApiUrl,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.WEBSITE_WEBFILE_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['mspp_name', 'mspp_webfileid', 'mspp_summary', '_mspp_publishingstateid_value@OData.Community.Display.V1.FormattedValue']);
  });

  it('fails validation if id is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        environmentName: validEnvironment,
        id: 'Invalid GUID',
        websiteId: validWebsiteId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if websiteId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        environmentName: validEnvironment,
        id: validId,
        websiteId: 'Invalid GUID'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if required options specified (id)', async () => {
    const actual = await command.validate({ options: { environmentName: validEnvironment, id: validId, websiteId: validWebsiteId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (websiteName)', async () => {
    const actual = await command.validate({ options: { environmentName: validEnvironment, id: validId, websiteName: validWebsiteName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation on unable to find website based on websiteName', async () => {
    const EmptyWebsiteResponse = {
      value: [
      ]
    };

    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'get').callsFake(async (opts) => {

      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.2/powerpagesites?$filter=name eq 'Invalid website'&$select=powerpagesiteid`)) {
        return EmptyWebsiteResponse;
      }
      throw `Invalid request`;
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        environmentName: `${validEnvironment}`,
        websiteName: 'Invalid website',
        id: validId
      }
    }), new CommandError(`The specified website 'Invalid website' does not exist.`));
  });


  it('return webfile based on websiteName and webfileId', async () => {

    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').callsFake(async () => envUrl);

    sinon.stub(request, 'get').callsFake(async opts => {

      if (opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.2/powerpagesites?$filter=name eq '${validWebsiteName}'&$select=powerpagesiteid`) {
        return {
          "value": [
            {
              "powerpagesiteid": validWebsiteId
            }
          ]
        };
      }

      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.2/mspp_webfiles?$filter=mspp_webfileid eq '${validId}' and _mspp_websiteid_value eq '${validWebsiteId}'`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return webfileResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        environmentName: validEnvironment,
        websiteName: validWebsiteName,
        id: validId
      }
    });
    assert(loggerLogSpy.calledWith(webfileResponse.value[0]));
  });

  it('retrieves webfile based on id and websiteId', async () => {

    sinon.stub(request, 'get').callsFake(async opts => {

      if ((opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/${validEnvironment}?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return envResponse;
        }
      }

      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.2/mspp_webfiles?$filter=mspp_webfileid eq '${validId}' and _mspp_websiteid_value eq '${validWebsiteId}'`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return webfileResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, environmentName: `${validEnvironment}`, websiteId: `${validWebsiteId}`, id: `${validId}` } });
    assert(loggerLogSpy.calledWith(webfileResponse.value[0]));

  });

  it('fails validation on unable to find webfile based on webfileId and websiteId as admin', async () => {

    sinon.stub(request, 'get').callsFake(async opts => {

      if ((opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/${validEnvironment}?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return envResponse;
        }
      }

      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.2/mspp_webfiles?$filter=mspp_webfileid eq 'Invalid id' and _mspp_websiteid_value eq '${validWebsiteId}'`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return { "value": [] };
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        environmentName: `${validEnvironment}`,
        websiteId: `${validWebsiteId}`,
        id: 'Invalid id'
      }
    }), new CommandError(`The specified webfile 'Invalid id' does not exist.`));

  });

});