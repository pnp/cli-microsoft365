import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./sitedesign-get');

describe(commands.SITEDESIGN_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
    assert.strictEqual(command.name.startsWith(commands.SITEDESIGN_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if both id and title options are not passed', (done) => {
    const actual = command.validate({
      options: {
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if both id and title options are passed', (done) => {
    const actual = command.validate({
      options: {
        id: '2c1ba4c4-cd9b-4417-832f-92a34bc34b2a',
        title: 'Contoso Site Design'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails to get site design when it does not exists', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesigns`) > -1) {
        return Promise.resolve({ value: [] });
      }
      return Promise.reject('The specified site design does not exist');
    });

    command.action(logger, {
      options: {
        debug: true,
        title: 'Contoso Site Design'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified site design does not exist`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails when multiple site designs with same title exists', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesigns`) > -1) {
        return Promise.resolve({
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
        }
        );
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        title: 'Contoso Site Design'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Multiple site designs with title Contoso Site Design found: ca360b7e-1946-4292-b854-e0ad904f1055, 88ff1405-35d0-4880-909a-97693822d261`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets information about the specified site design by id', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignMetadata`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          id: 'ca360b7e-1946-4292-b854-e0ad904f1055'
        })) {
        return Promise.resolve({
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
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: 'ca360b7e-1946-4292-b854-e0ad904f1055' } }, () => {
      try {
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
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets information about the specified site design by title', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesigns`) > -1) {
        return Promise.resolve({
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
        });
      }

      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignMetadata`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          id: 'ca360b7e-1946-4292-b854-e0ad904f1055'
        })) {
        return Promise.resolve({
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
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, title: 'Contoso Site Design' } }, () => {
      try {
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
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets information about the specified site design (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignMetadata`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          id: 'ca360b7e-1946-4292-b854-e0ad904f1055'
        })) {
        return Promise.resolve({
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
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, id: 'ca360b7e-1946-4292-b854-e0ad904f1055' } }, () => {
      try {
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
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when site design not found', (done) => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'File Not Found.' } } } });
    });

    command.action(logger, { options: { debug: false, id: 'ca360b7e-1946-4292-b854-e0ad904f1055' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('File Not Found.')));
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

  it('supports specifying id', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--id') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if the id is not a valid GUID', () => {
    const actual = command.validate({ options: { id: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the id is a valid GUID', () => {
    const actual = command.validate({ options: { id: '2c1ba4c4-cd9b-4417-832f-92a34bc34b2a' } });
    assert.strictEqual(actual, true);
  });
});