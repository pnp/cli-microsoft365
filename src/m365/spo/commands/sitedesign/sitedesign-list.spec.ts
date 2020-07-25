import commands from '../../commands';
import Command, { CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./sitedesign-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import auth from '../../../../Auth';

describe(commands.SITEDESIGN_LIST, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({ FormDigestValue: 'ABC' }));
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: any) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
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

    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "Id": "9b142c22-037f-4a7f-9017-e9d8c0e34b98",
            "IsDefault": false,
            "Title": "Contoso REST",
            "Version": 1,
            "WebTemplate": "64"
          },
          {
            "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
            "IsDefault": false,
            "Title": "REST test",
            "Version": 1,
            "WebTemplate": "64"
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

    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "Id": "9b142c22-037f-4a7f-9017-e9d8c0e34b98",
            "IsDefault": false,
            "Title": "Contoso REST",
            "Version": 1,
            "WebTemplate": "64"
          },
          {
            "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
            "IsDefault": false,
            "Title": "REST test",
            "Version": 1,
            "WebTemplate": "64"
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

    cmdInstance.action({ options: { debug: false, output: 'json' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
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

    cmdInstance.action({ options: { debug: false } }, (err?: any) => {
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
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});