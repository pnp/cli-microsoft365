import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./sitedesign-list');

describe(commands.SITEDESIGN_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({ FormDigestValue: 'ABC' }));
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
    Utils.restore([
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      (command as any).getRequestDigest,
      appInsights.trackEvent
    ]);
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

  it('lists available site designs', (done) => {
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

    command.action(logger, { options: { debug: false } }, () => {
      try {
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
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists available site designs (debug)', (done) => {
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

    command.action(logger, { options: { debug: true } }, () => {
      try {
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
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists available site designs with all properties for JSON output', (done) => {
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

    command.action(logger, { options: { debug: false, output: 'json' } }, () => {
      try {
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
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles OData error when retrieving available site designs', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });
    });

    command.action(logger, { options: { debug: false } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});