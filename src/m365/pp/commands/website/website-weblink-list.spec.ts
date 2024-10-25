import assert from 'assert';
import sinon from 'sinon';
import { z } from 'zod';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './website-weblink-list.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { cli } from '../../../../cli/cli.js';

describe(commands.WEBSITE_WEBLINK_LIST, () => {
  //#region Mocked Responses
  let commandInfo: CommandInfo;
  const validEnvironment = 'Default-eff8592e-e14a-4ae8-8771-d96d5c549e1c';
  const validWebsiteId = '3bbc8102-8ee7-4dac-afbb-807cc5b6f9c2';
  const validWebsiteName = 'CLI 365 PowerPageSite';

  const envResponse: any = { "properties": { "linkedEnvironmentMetadata": { "instanceApiUrl": "https://contoso-dev.api.crm4.dynamics.com" } } };
  const weblinksetsResponse: any = {
    "value": [
      {
        "mspp_weblinksetid": "c94de7c8-5474-45b6-9172-c15e8e4ba1e1",
        "mspp_name": "Default",
        "_mspp_websitelanguageid_value": "403f1195-3f13-ef11-9f89-000d3a3755aa",
        "mspp_display_name": "Default",
        "_mspp_publishingstateid_value": "11d71883-3f13-ef11-9f89-000d3a593739",
        "_mspp_websiteid_value": "3bbc8102-8ee7-4dac-afbb-807cc5b6f9c2",
        "_mspp_createdby_value": "5364ffa6-d185-ee11-8179-6045bd0027e0",
        "_mspp_modifiedby_value": "5364ffa6-d185-ee11-8179-6045bd0027e0",
        "mspp_modifiedon": "2024-05-16T04:49:10Z",
        "mspp_createdon": "2024-05-16T04:48:37Z",
        "statecode": 0,
        "statuscode": 1,
        "mspp_title": null,
        "mspp_copy": null
      },
      {
        "mspp_weblinksetid": "28ae3dfa-5939-4f79-9373-1665a839a9b2",
        "mspp_name": "Default",
        "_mspp_websitelanguageid_value": "403f1195-3f13-ef11-9f89-000d3a3755aa",
        "mspp_display_name": "Default",
        "_mspp_publishingstateid_value": "11d71883-3f13-ef11-9f89-000d3a593739",
        "_mspp_websiteid_value": "3bbc8102-8ee7-4dac-afbb-807cc5b6f9c2",
        "_mspp_createdby_value": "5364ffa6-d185-ee11-8179-6045bd0027e0",
        "_mspp_modifiedby_value": "5364ffa6-d185-ee11-8179-6045bd0027e0",
        "mspp_modifiedon": "2024-05-16T04:49:10Z",
        "mspp_createdon": "2024-05-16T04:48:37Z",
        "statecode": 0,
        "statuscode": 1,
        "mspp_title": null,
        "mspp_copy": null
      }
    ]
  };
  const weblinksResponse: any = {
    "value": [
      {
        "mspp_weblinkid": "fccea7a1-a1dc-418e-839e-f39c0ec400a2",
        "mspp_name": "Contact us",
        "mspp_disablepagevalidation": false,
        "mspp_displayimageonly": false,
        "mspp_displayorder": 3,
        "mspp_displaypagechildlinks": false,
        "mspp_openinnewwindow": false,
        "_mspp_pageid_value": "4cdcd042-bb91-4673-ae89-b44bfdb3a751",
        "_mspp_publishingstateid_value": "ffefd269-446e-46aa-9379-92d1b2d323b8",
        "mspp_robotsfollowlink": true,
        "_mspp_weblinksetid_value": "c94de7c8-5474-45b6-9172-c15e8e4ba1e1",
        "_mspp_createdby_value": "5364ffa6-d185-ee11-8179-6045bd0027e0",
        "_mspp_modifiedby_value": "5364ffa6-d185-ee11-8179-6045bd0027e0",
        "mspp_modifiedon": "2024-04-12T05:46:36Z",
        "mspp_createdon": "2024-04-12T05:46:36Z",
        "statecode": 0,
        "statuscode": 1,
        "mspp_description": null,
        "mspp_imagealttext": null,
        "mspp_imageurl": null,
        "mspp_imageheight": null,
        "mspp_modifiedbyusername": null,
        "mspp_createdbyipaddress": null,
        "mspp_createdbyusername": null,
        "mspp_modifiedbyipaddress": null,
        "_mspp_parentweblinkid_value": null,
        "mspp_imagewidth": null,
        "mspp_externalurl": null
      },
      {
        "mspp_weblinkid": "18a5589e-6472-4ed6-90ba-fd51a15325d8",
        "mspp_name": "Subpage 1",
        "mspp_disablepagevalidation": false,
        "mspp_displayimageonly": false,
        "mspp_displayorder": 1,
        "mspp_displaypagechildlinks": false,
        "mspp_openinnewwindow": false,
        "_mspp_pageid_value": "1a4cbb29-223b-4be2-ae9e-bc281de64b12",
        "_mspp_publishingstateid_value": "ffefd269-446e-46aa-9379-92d1b2d323b8",
        "mspp_robotsfollowlink": true,
        "_mspp_weblinksetid_value": "28ae3dfa-5939-4f79-9373-1665a839a9b2",
        "_mspp_createdby_value": "5364ffa6-d185-ee11-8179-6045bd0027e0",
        "_mspp_modifiedby_value": "5364ffa6-d185-ee11-8179-6045bd0027e0",
        "mspp_modifiedon": "2024-04-12T05:46:47Z",
        "mspp_createdon": "2024-04-12T05:46:47Z",
        "statecode": 0,
        "statuscode": 1,
        "mspp_description": null,
        "mspp_imagealttext": null,
        "mspp_imageurl": null,
        "mspp_imageheight": null,
        "mspp_modifiedbyusername": null,
        "mspp_createdbyipaddress": null,
        "mspp_createdbyusername": null,
        "mspp_modifiedbyipaddress": null,
        "_mspp_parentweblinkid_value": "4ebe880e-33fe-48a9-80f6-37f585207382",
        "mspp_imagewidth": null,
        "mspp_externalurl": null
      }
    ]
  };
  //#endregion

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandOptionsSchema: z.ZodTypeAny;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(accessToken, 'assertDelegatedAccessToken').returns();
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
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
    assert.strictEqual(command.name, commands.WEBSITE_WEBLINK_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if websiteId is not a valid guid.', async () => {
    const actual = commandOptionsSchema.safeParse({ environmentName: validEnvironment, websiteId: 'Invalid GUID' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if both websiteId or websiteName is provided', async () => {
    const actual = commandOptionsSchema.safeParse({ environmentName: validEnvironment, websiteId: validWebsiteId, websiteName: validWebsiteName });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if either websiteId or websiteName is not provided', async () => {
    const actual = commandOptionsSchema.safeParse({ environmentName: validEnvironment });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if required option websiteId specified', async () => {
    const actual = commandOptionsSchema.safeParse({ environmentName: validEnvironment, websiteId: validWebsiteId });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if required option websiteName specified', async () => {
    const actual = await command.validate({ options: { environmentName: validEnvironment, websiteName: validWebsiteName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['mspp_name', 'mspp_weblinkid', 'mspp_description', 'statecode']);
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

  it('retrieves weblinks', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/${validEnvironment}?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return envResponse;
        }
      }

      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.2/mspp_weblinksets?$filter=_mspp_websiteid_value eq '${validWebsiteId}'`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return weblinksetsResponse;
        }
      }

      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.2/mspp_weblinks?$filter=Microsoft.Dynamics.CRM.ContainValues(PropertyName=@p1,PropertyValues=@p2)&@p1='mspp_weblinksetid'&@p2=['c94de7c8-5474-45b6-9172-c15e8e4ba1e1','28ae3dfa-5939-4f79-9373-1665a839a9b2']`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return weblinksResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, environmentName: `${validEnvironment}`, websiteId: `${validWebsiteId}` } });
    assert(loggerLogSpy.calledWith(weblinksResponse.value));

  });

  it('failed to fetch weblinksets based on websiteid', async () => {
    const EmptyLinksetsResponse = {
      value: [
      ]
    };

    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/${validEnvironment}?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`)) {
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

      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.2/mspp_weblinksets?$filter=_mspp_websiteid_value eq '${validWebsiteId}'`)) {
        return EmptyLinksetsResponse;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        environmentName: `${validEnvironment}`,
        websiteName: `${validWebsiteName}`
      }
    }), new CommandError(`The specified website '${validWebsiteId}' does not have links.`));
  });

  it('retrieves weblinks based on website name as admin', async () => {
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

      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.2/mspp_weblinksets?$filter=_mspp_websiteid_value eq '${validWebsiteId}'`)) {
        return weblinksetsResponse;
      }


      if ((opts.url === `https://contoso-dev.api.crm4.dynamics.com/api/data/v9.2/mspp_weblinks?$filter=Microsoft.Dynamics.CRM.ContainValues(PropertyName=@p1,PropertyValues=@p2)&@p1='mspp_weblinksetid'&@p2=['c94de7c8-5474-45b6-9172-c15e8e4ba1e1','28ae3dfa-5939-4f79-9373-1665a839a9b2']`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return weblinksResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, environmentName: validEnvironment, websiteName: validWebsiteName, asAdmin: true } });
    assert(loggerLogSpy.calledWith(weblinksResponse.value));
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