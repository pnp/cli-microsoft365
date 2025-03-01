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
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command from './sitedesign-get.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.SITEDESIGN_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
    commandInfo = cli.getCommandInfo(command);
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
      request.post,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SITEDESIGN_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails to get site design when it does not exists', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesigns`) > -1) {
        return { value: [] };
      }
      throw 'The specified site design does not exist';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        title: 'Contoso Site Design'
      }
    } as any), new CommandError('The specified site design does not exist'));
  });

  it('fails when multiple site designs with same title exists', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesigns`) > -1) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams",
          "@odata.count": 2,
          "value": [
            {
              "Description": null,
              "DesignPackageId": "00000000-0000-0000-0000-000000000000",
              "DesignType": "0",
              "IsDefault": false,
              "IsOutOfBoxTemplate": false,
              "IsTenantAdminOnly": false,
              "PreviewImageAltText": null,
              "PreviewImageUrl": null,
              "RequiresGroupConnected": false,
              "RequiresTeamsConnected": false,
              "RequiresYammerConnected": false,
              "SiteScriptIds": [
                "3aff9f82-fe6c-42d3-803f-8951d26ed854"
              ],
              "SupportedWebTemplates": [],
              "TemplateFeatures": [],
              "ThumbnailUrl": null,
              "Title": "Contoso Site Design",
              "WebTemplate": "68",
              "Id": "ca360b7e-1946-4292-b854-e0ad904f1055",
              "Version": 1
            },
            {
              "Description": null,
              "DesignPackageId": "00000000-0000-0000-0000-000000000000",
              "DesignType": "0",
              "IsDefault": false,
              "IsOutOfBoxTemplate": false,
              "IsTenantAdminOnly": false,
              "PreviewImageAltText": null,
              "PreviewImageUrl": null,
              "RequiresGroupConnected": false,
              "RequiresTeamsConnected": false,
              "RequiresYammerConnected": false,
              "SiteScriptIds": [
                "3aff9f82-fe6c-42d3-803f-8951d26ed854"
              ],
              "SupportedWebTemplates": [],
              "TemplateFeatures": [],
              "ThumbnailUrl": null,
              "Title": "Contoso Site Design",
              "WebTemplate": "68",
              "Id": "88ff1405-35d0-4880-909a-97693822d261",
              "Version": 1
            }
          ]
        };
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        title: 'Contoso Site Design'
      }
    } as any), new CommandError(`Multiple site designs with title 'Contoso Site Design' found. Found: ca360b7e-1946-4292-b854-e0ad904f1055, 88ff1405-35d0-4880-909a-97693822d261.`));
  });

  it('handles selecting single result when multiple site designs with the specified title found and cli is set to prompt', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesigns`) > -1) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams",
          "@odata.count": 2,
          "value": [
            {
              "Description": null,
              "DesignPackageId": "00000000-0000-0000-0000-000000000000",
              "DesignType": "0",
              "IsDefault": false,
              "IsOutOfBoxTemplate": false,
              "IsTenantAdminOnly": false,
              "PreviewImageAltText": null,
              "PreviewImageUrl": null,
              "RequiresGroupConnected": false,
              "RequiresTeamsConnected": false,
              "RequiresYammerConnected": false,
              "SiteScriptIds": [
                "3aff9f82-fe6c-42d3-803f-8951d26ed854"
              ],
              "SupportedWebTemplates": [],
              "TemplateFeatures": [],
              "ThumbnailUrl": null,
              "Title": "Contoso Site Design",
              "WebTemplate": "68",
              "Id": "ca360b7e-1946-4292-b854-e0ad904f1055",
              "Version": 1
            },
            {
              "Description": null,
              "DesignPackageId": "00000000-0000-0000-0000-000000000000",
              "DesignType": "0",
              "IsDefault": false,
              "IsOutOfBoxTemplate": false,
              "IsTenantAdminOnly": false,
              "PreviewImageAltText": null,
              "PreviewImageUrl": null,
              "RequiresGroupConnected": false,
              "RequiresTeamsConnected": false,
              "RequiresYammerConnected": false,
              "SiteScriptIds": [
                "3aff9f82-fe6c-42d3-803f-8951d26ed854"
              ],
              "SupportedWebTemplates": [],
              "TemplateFeatures": [],
              "ThumbnailUrl": null,
              "Title": "Contoso Site Design",
              "WebTemplate": "68",
              "Id": "88ff1405-35d0-4880-909a-97693822d261",
              "Version": 1
            }
          ]
        };
      }

      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignMetadata`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          id: 'ca360b7e-1946-4292-b854-e0ad904f1055'
        })) {
        return {
          "Description": null,
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": [
            "449c0c6d-5380-4df2-b84b-622e0ac8ec24"
          ],
          "Title": "Contoso REST",
          "WebTemplate": "64",
          "Id": "ca360b7e-1946-4292-b854-e0ad904f1055",
          "Version": 1
        };
      }

      throw 'Invalid request';
    });


    sinon.stub(cli, 'handleMultipleResultsFound').resolves({ Id: 'ca360b7e-1946-4292-b854-e0ad904f1055' });

    await command.action(logger, { options: { title: 'Contoso Site Design' } });
    assert(loggerLogSpy.calledWith({
      "Description": null,
      "IsDefault": false,
      "PreviewImageAltText": null,
      "PreviewImageUrl": null,
      "SiteScriptIds": [
        "449c0c6d-5380-4df2-b84b-622e0ac8ec24"
      ],
      "Title": "Contoso REST",
      "WebTemplate": "64",
      "Id": "ca360b7e-1946-4292-b854-e0ad904f1055",
      "Version": 1
    }));
  });

  it('gets information about the specified site design by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignMetadata`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          id: 'ca360b7e-1946-4292-b854-e0ad904f1055'
        })) {
        return {
          "Description": null,
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": [
            "449c0c6d-5380-4df2-b84b-622e0ac8ec24"
          ],
          "Title": "Contoso REST",
          "WebTemplate": "64",
          "Id": "ca360b7e-1946-4292-b854-e0ad904f1055",
          "Version": 1
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: 'ca360b7e-1946-4292-b854-e0ad904f1055' } });
    assert(loggerLogSpy.calledWith({
      "Description": null,
      "IsDefault": false,
      "PreviewImageAltText": null,
      "PreviewImageUrl": null,
      "SiteScriptIds": [
        "449c0c6d-5380-4df2-b84b-622e0ac8ec24"
      ],
      "Title": "Contoso REST",
      "WebTemplate": "64",
      "Id": "ca360b7e-1946-4292-b854-e0ad904f1055",
      "Version": 1
    }));
  });

  it('gets information about the specified site design by title', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesigns`) > -1) {
        return {
          "value": [
            {
              "Description": null,
              "DesignPackageId": "00000000-0000-0000-0000-000000000000",
              "DesignType": "0",
              "IsDefault": false,
              "IsOutOfBoxTemplate": false,
              "IsTenantAdminOnly": false,
              "PreviewImageAltText": null,
              "PreviewImageUrl": null,
              "RequiresGroupConnected": false,
              "RequiresTeamsConnected": false,
              "RequiresYammerConnected": false,
              "SiteScriptIds": [
                "3aff9f82-fe6c-42d3-803f-8951d26ed854"
              ],
              "SupportedWebTemplates": [],
              "TemplateFeatures": [],
              "ThumbnailUrl": null,
              "Title": "Contoso Site Design",
              "WebTemplate": "68",
              "Id": "ca360b7e-1946-4292-b854-e0ad904f1055",
              "Version": 1
            }
          ]
        };
      }

      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignMetadata`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          id: 'ca360b7e-1946-4292-b854-e0ad904f1055'
        })) {
        return {
          "Description": null,
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": [
            "3aff9f82-fe6c-42d3-803f-8951d26ed854"
          ],
          "Title": "Contoso Site Design",
          "WebTemplate": "68",
          "Id": "ca360b7e-1946-4292-b854-e0ad904f1055",
          "Version": 1
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { title: 'Contoso Site Design' } });
    assert(loggerLogSpy.calledWith({
      "Description": null,
      "IsDefault": false,
      "PreviewImageAltText": null,
      "PreviewImageUrl": null,
      "SiteScriptIds": [
        "3aff9f82-fe6c-42d3-803f-8951d26ed854"
      ],
      "Title": "Contoso Site Design",
      "WebTemplate": "68",
      "Id": "ca360b7e-1946-4292-b854-e0ad904f1055",
      "Version": 1
    }));
  });

  it('gets information about the specified site design (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignMetadata`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          id: 'ca360b7e-1946-4292-b854-e0ad904f1055'
        })) {
        return {
          "Description": null,
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": [
            "449c0c6d-5380-4df2-b84b-622e0ac8ec24"
          ],
          "Title": "Contoso REST",
          "WebTemplate": "64",
          "Id": "ca360b7e-1946-4292-b854-e0ad904f1055",
          "Version": 1
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, id: 'ca360b7e-1946-4292-b854-e0ad904f1055' } });
    assert(loggerLogSpy.calledWith({
      "Description": null,
      "IsDefault": false,
      "PreviewImageAltText": null,
      "PreviewImageUrl": null,
      "SiteScriptIds": [
        "449c0c6d-5380-4df2-b84b-622e0ac8ec24"
      ],
      "Title": "Contoso REST",
      "WebTemplate": "64",
      "Id": "ca360b7e-1946-4292-b854-e0ad904f1055",
      "Version": 1
    }));
  });

  it('correctly handles error when site design not found', async () => {
    sinon.stub(request, 'post').rejects({ error: { 'odata.error': { message: { value: 'File Not Found.' } } } });

    await assert.rejects(command.action(logger, { options: { id: 'ca360b7e-1946-4292-b854-e0ad904f1055' } } as any), new CommandError('File Not Found.'));
  });

  it('supports specifying id', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--id') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if the id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the id is a valid GUID', async () => {
    const actual = await command.validate({ options: { id: '2c1ba4c4-cd9b-4417-832f-92a34bc34b2a' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
