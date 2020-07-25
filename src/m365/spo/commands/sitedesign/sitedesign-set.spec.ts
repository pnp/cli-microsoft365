import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./sitedesign-set');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import auth from '../../../../Auth';

describe(commands.SITEDESIGN_SET, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
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
      log: (msg: string) => {
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
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SITEDESIGN_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('updates site design title', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteDesign`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          updateInfo: {
            Id: '2a9f178a-4d1d-449c-9296-df509ab4702c',
            Title: 'New title'
          }
        })) {
        return Promise.resolve({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "New title",
          "Version": 1,
          "WebTemplate": 64
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, id: '2a9f178a-4d1d-449c-9296-df509ab4702c', title: 'New title' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "New title",
          "Version": 1,
          "WebTemplate": 64
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates site design web template to TeamSite', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteDesign`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          updateInfo: {
            Id: '2a9f178a-4d1d-449c-9296-df509ab4702c',
            WebTemplate: '64'
          }
        })) {
        return Promise.resolve({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Title",
          "Version": 1,
          "WebTemplate": 64
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, id: '2a9f178a-4d1d-449c-9296-df509ab4702c', webTemplate: 'TeamSite' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Title",
          "Version": 1,
          "WebTemplate": 64
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates site design web template to CommunicationSite', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteDesign`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          updateInfo: {
            Id: '2a9f178a-4d1d-449c-9296-df509ab4702c',
            WebTemplate: '68'
          }
        })) {
        return Promise.resolve({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Title",
          "Version": 1,
          "WebTemplate": 68
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, id: '2a9f178a-4d1d-449c-9296-df509ab4702c', webTemplate: 'CommunicationSite' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Title",
          "Version": 1,
          "WebTemplate": 68
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates site design site scripts (one script)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteDesign`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          updateInfo: {
            Id: '2a9f178a-4d1d-449c-9296-df509ab4702c',
            SiteScriptIds: ['449c0c6d-5380-4df2-b84b-622e0ac8ec24']
          }
        })) {
        return Promise.resolve({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Title",
          "Version": 1,
          "WebTemplate": 64
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, id: '2a9f178a-4d1d-449c-9296-df509ab4702c', siteScripts: '449c0c6d-5380-4df2-b84b-622e0ac8ec24' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Title",
          "Version": 1,
          "WebTemplate": 64
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates site design site scripts (multiple scripts)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteDesign`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          updateInfo: {
            Id: '2a9f178a-4d1d-449c-9296-df509ab4702c',
            SiteScriptIds: ['449c0c6d-5380-4df2-b84b-622e0ac8ec24', '449c0c6d-5380-4df2-b84b-622e0ac8ec25']
          }
        })) {
        return Promise.resolve({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24", "449c0c6d-5380-4df2-b84b-622e0ac8ec25"],
          "Title": "Title",
          "Version": 1,
          "WebTemplate": 64
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, id: '2a9f178a-4d1d-449c-9296-df509ab4702c', siteScripts: '449c0c6d-5380-4df2-b84b-622e0ac8ec24, 449c0c6d-5380-4df2-b84b-622e0ac8ec25' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24", "449c0c6d-5380-4df2-b84b-622e0ac8ec25"],
          "Title": "Title",
          "Version": 1,
          "WebTemplate": 64
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates site design description', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteDesign`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          updateInfo: {
            Id: '2a9f178a-4d1d-449c-9296-df509ab4702c',
            Description: 'New description'
          }
        })) {
        return Promise.resolve({
          "Description": "New description",
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Title",
          "Version": 1,
          "WebTemplate": 64
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, id: '2a9f178a-4d1d-449c-9296-df509ab4702c', description: 'New description' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "Description": "New description",
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Title",
          "Version": 1,
          "WebTemplate": 64
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates site design previewImageUrl', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteDesign`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          updateInfo: {
            Id: '2a9f178a-4d1d-449c-9296-df509ab4702c',
            PreviewImageUrl: 'https://contoso.com/image.png'
          }
        })) {
        return Promise.resolve({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": "https://contoso.com/image.png",
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Title",
          "Version": 1,
          "WebTemplate": 64
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, id: '2a9f178a-4d1d-449c-9296-df509ab4702c', previewImageUrl: 'https://contoso.com/image.png' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": "https://contoso.com/image.png",
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Title",
          "Version": 1,
          "WebTemplate": 64
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates site design previewImageAltText', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteDesign`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          updateInfo: {
            Id: '2a9f178a-4d1d-449c-9296-df509ab4702c',
            PreviewImageAltText: 'Logo image'
          }
        })) {
        return Promise.resolve({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": "Logo image",
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Title",
          "Version": 1,
          "WebTemplate": 64
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, id: '2a9f178a-4d1d-449c-9296-df509ab4702c', previewImageAltText: 'Logo image' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": "Logo image",
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Title",
          "Version": 1,
          "WebTemplate": 64
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates site design version', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteDesign`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          updateInfo: {
            Id: '2a9f178a-4d1d-449c-9296-df509ab4702c',
            Version: 2
          }
        })) {
        return Promise.resolve({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Title",
          "Version": 2,
          "WebTemplate": 64
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, id: '2a9f178a-4d1d-449c-9296-df509ab4702c', version: 2 } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Title",
          "Version": 2,
          "WebTemplate": 64
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('makes site design default', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteDesign`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          updateInfo: {
            Id: '2a9f178a-4d1d-449c-9296-df509ab4702c',
            IsDefault: true
          }
        })) {
        return Promise.resolve({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": true,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Title",
          "Version": 1,
          "WebTemplate": 64
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, id: '2a9f178a-4d1d-449c-9296-df509ab4702c', isDefault: 'true' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": true,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Title",
          "Version": 1,
          "WebTemplate": 64
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('makes site design not-default (explicit)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteDesign`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          updateInfo: {
            Id: '2a9f178a-4d1d-449c-9296-df509ab4702c',
            IsDefault: false
          }
        })) {
        return Promise.resolve({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Title",
          "Version": 1,
          "WebTemplate": 64
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, id: '2a9f178a-4d1d-449c-9296-df509ab4702c', isDefault: 'false' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Title",
          "Version": 1,
          "WebTemplate": 64
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('makes site design not-default (implicit)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteDesign`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          updateInfo: {
            Id: '2a9f178a-4d1d-449c-9296-df509ab4702c'
          }
        })) {
        return Promise.resolve({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Title",
          "Version": 1,
          "WebTemplate": 64
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, id: '2a9f178a-4d1d-449c-9296-df509ab4702c' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Title",
          "Version": 1,
          "WebTemplate": 64
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates all site design properties (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteDesign`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          "updateInfo": {
            "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
            "Title": "Contoso",
            "Description": "Contoso team site",
            "SiteScriptIds": [
              "449c0c6d-5380-4df2-b84b-622e0ac8ec24"
            ],
            "PreviewImageUrl": "https://contoso.com/assets/team-site-preview.png",
            "PreviewImageAltText": "Contoso team site preview",
            "WebTemplate": "64",
            "Version": 2,
            "IsDefault": true
          }
        })) {
        return Promise.resolve({
          "Description": 'Contoso team site',
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": true,
          "PreviewImageAltText": 'Contoso team site preview',
          "PreviewImageUrl": 'https://contoso.com/assets/team-site-preview.png',
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 2,
          "WebTemplate": 64
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, id: '2a9f178a-4d1d-449c-9296-df509ab4702c', title: 'Contoso', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24", description: 'Contoso team site', previewImageUrl: 'https://contoso.com/assets/team-site-preview.png', previewImageAltText: 'Contoso team site preview', version: 2, isDefault: 'true' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "Description": 'Contoso team site',
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": true,
          "PreviewImageAltText": 'Contoso team site preview',
          "PreviewImageUrl": 'https://contoso.com/assets/team-site-preview.png',
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 2,
          "WebTemplate": 64
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles OData error when updating site design', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });
    });

    cmdInstance.action({ options: { debug: false, id: '9b142c22-037f-4a7f-9017-e9d8c0e34b98', webTemplate: 'TeamSite', siteScripts: '449c0c6d-5380-4df2-b84b-622e0ac8ec24' } }, (err?: any) => {
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

  it('supports specifying id', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--id') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying title', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--title') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying webTemplate', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--webTemplate') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying siteScripts', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--siteScripts') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying description', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--description') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying previewImageUrl', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--previewImageUrl') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying previewImageAltText', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--previewImageAltText') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying version', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--version') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying isDefault', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--isDefault') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if id specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if id is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('passed validation if id is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '9b142c22-037f-4a7f-9017-e9d8c0e34b99' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if specified webTemplate is invalid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '9b142c22-037f-4a7f-9017-e9d8c0e34b99', webTemplate: 'Invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if specified webTemplate is CommunicationSite', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '9b142c22-037f-4a7f-9017-e9d8c0e34b99', webTemplate: 'CommunicationSite' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if specified webTemplate is TeamSite', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '9b142c22-037f-4a7f-9017-e9d8c0e34b99', webTemplate: 'TeamSite' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if specified siteScripts is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '9b142c22-037f-4a7f-9017-e9d8c0e34b99', webTemplate: 'TeamSite', siteScripts: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the second specified siteScriptId is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '9b142c22-037f-4a7f-9017-e9d8c0e34b99', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24,abc" } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if specified siteScriptId is valid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '9b142c22-037f-4a7f-9017-e9d8c0e34b99', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24" } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if all required parameters are valid (multiple siteScripts)', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '9b142c22-037f-4a7f-9017-e9d8c0e34b99', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24,449c0c6d-5380-4df2-b84b-622e0ac8ec25" } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if specified version is not a number', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '9b142c22-037f-4a7f-9017-e9d8c0e34b99', version: 'a' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if specified version is a number', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '9b142c22-037f-4a7f-9017-e9d8c0e34b99', version: 2 } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if specified isDefault value is invalid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '9b142c22-037f-4a7f-9017-e9d8c0e34b99', isDefault: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if specified isDefault value is true', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '9b142c22-037f-4a7f-9017-e9d8c0e34b99', isDefault: 'true' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if specified isDefault value is false', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '9b142c22-037f-4a7f-9017-e9d8c0e34b99', isDefault: 'false' } });
    assert.strictEqual(actual, true);
  });
});