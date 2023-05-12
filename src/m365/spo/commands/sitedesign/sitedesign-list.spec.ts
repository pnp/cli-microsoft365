import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./sitedesign-list');

describe(commands.SITEDESIGN_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
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
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SITEDESIGN_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['Id', 'IsDefault', 'Title', 'Version', 'WebTemplate']);
  });

  it('lists available site designs', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesigns`) > -1) {
        return Promise.resolve({
          value: [
            {
              "Description": null,
              "IsDefault": false,
              "PreviewImageAltText": null,
              "PreviewImageUrl": null,
              "SiteScriptIds": [
                "449c0c6d-5380-4df2-b84b-622e0ac8ec25"
              ],
              "Title": "Contoso REST",
              "WebTemplate": "64",
              "Id": "9b142c22-037f-4a7f-9017-e9d8c0e34b98",
              "Version": 1
            },
            {
              "Description": null,
              "IsDefault": false,
              "PreviewImageAltText": null,
              "PreviewImageUrl": null,
              "SiteScriptIds": [
                "449c0c6d-5380-4df2-b84b-622e0ac8ec24"
              ],
              "Title": "REST test",
              "WebTemplate": "64",
              "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
              "Version": 1
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith([
      {
        "Description": null,
        "IsDefault": false,
        "PreviewImageAltText": null,
        "PreviewImageUrl": null,
        "SiteScriptIds": [
          "449c0c6d-5380-4df2-b84b-622e0ac8ec25"
        ],
        "Title": "Contoso REST",
        "WebTemplate": "64",
        "Id": "9b142c22-037f-4a7f-9017-e9d8c0e34b98",
        "Version": 1
      },
      {
        "Description": null,
        "IsDefault": false,
        "PreviewImageAltText": null,
        "PreviewImageUrl": null,
        "SiteScriptIds": [
          "449c0c6d-5380-4df2-b84b-622e0ac8ec24"
        ],
        "Title": "REST test",
        "WebTemplate": "64",
        "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
        "Version": 1
      }
    ]));
  });

  it('lists available site designs (debug)', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesigns`) > -1) {
        return Promise.resolve({
          value: [
            {
              "Description": null,
              "IsDefault": false,
              "PreviewImageAltText": null,
              "PreviewImageUrl": null,
              "SiteScriptIds": [
                "449c0c6d-5380-4df2-b84b-622e0ac8ec25"
              ],
              "Title": "Contoso REST",
              "WebTemplate": "64",
              "Id": "9b142c22-037f-4a7f-9017-e9d8c0e34b98",
              "Version": 1
            },
            {
              "Description": null,
              "IsDefault": false,
              "PreviewImageAltText": null,
              "PreviewImageUrl": null,
              "SiteScriptIds": [
                "449c0c6d-5380-4df2-b84b-622e0ac8ec24"
              ],
              "Title": "REST test",
              "WebTemplate": "64",
              "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
              "Version": 1
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: true } });
    assert(loggerLogSpy.calledWith([
      {
        "Description": null,
        "IsDefault": false,
        "PreviewImageAltText": null,
        "PreviewImageUrl": null,
        "SiteScriptIds": [
          "449c0c6d-5380-4df2-b84b-622e0ac8ec25"
        ],
        "Title": "Contoso REST",
        "WebTemplate": "64",
        "Id": "9b142c22-037f-4a7f-9017-e9d8c0e34b98",
        "Version": 1
      },
      {
        "Description": null,
        "IsDefault": false,
        "PreviewImageAltText": null,
        "PreviewImageUrl": null,
        "SiteScriptIds": [
          "449c0c6d-5380-4df2-b84b-622e0ac8ec24"
        ],
        "Title": "REST test",
        "WebTemplate": "64",
        "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
        "Version": 1
      }
    ]));
  });

  it('lists available site designs with all properties for JSON output', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesigns`) > -1) {
        return Promise.resolve({
          value: [
            {
              "Description": null,
              "IsDefault": false,
              "PreviewImageAltText": null,
              "PreviewImageUrl": null,
              "SiteScriptIds": [
                "449c0c6d-5380-4df2-b84b-622e0ac8ec25"
              ],
              "Title": "Contoso REST",
              "WebTemplate": "64",
              "Id": "9b142c22-037f-4a7f-9017-e9d8c0e34b98",
              "Version": 1
            },
            {
              "Description": null,
              "IsDefault": false,
              "PreviewImageAltText": null,
              "PreviewImageUrl": null,
              "SiteScriptIds": [
                "449c0c6d-5380-4df2-b84b-622e0ac8ec24"
              ],
              "Title": "REST test",
              "WebTemplate": "64",
              "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
              "Version": 1
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { output: 'json' } });
    assert(loggerLogSpy.calledWith([
      {
        "Description": null,
        "IsDefault": false,
        "PreviewImageAltText": null,
        "PreviewImageUrl": null,
        "SiteScriptIds": [
          "449c0c6d-5380-4df2-b84b-622e0ac8ec25"
        ],
        "Title": "Contoso REST",
        "WebTemplate": "64",
        "Id": "9b142c22-037f-4a7f-9017-e9d8c0e34b98",
        "Version": 1
      },
      {
        "Description": null,
        "IsDefault": false,
        "PreviewImageAltText": null,
        "PreviewImageUrl": null,
        "SiteScriptIds": [
          "449c0c6d-5380-4df2-b84b-622e0ac8ec24"
        ],
        "Title": "REST test",
        "WebTemplate": "64",
        "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
        "Version": 1
      }
    ]));
  });

  it('correctly handles OData error when retrieving available site designs', async () => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });
});
