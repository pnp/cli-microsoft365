import commands from '../../commands';
import flowCommands from '../../../flow/commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./connector-export');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import * as fs from 'fs';
import * as path from 'path';

describe(commands.CONNECTOR_EXPORT, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let writeFileSyncStub: sinon.SinonStub;
  let mkdirSyncStub: sinon.SinonStub;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    mkdirSyncStub = sinon.stub(fs, 'mkdirSync').callsFake(() => { });
    writeFileSyncStub = sinon.stub(fs, 'writeFileSync').callsFake(() => { });
    auth.service.connected = true;
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
    mkdirSyncStub.reset();
    writeFileSyncStub.reset();
    Utils.restore([
      request.get,
      fs.existsSync
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      fs.mkdirSync,
      fs.writeFileSync
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CONNECTOR_EXPORT), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines alias', () => {
    const alias = command.alias();
    assert.notStrictEqual(typeof alias, 'undefined');
  });

  it('defines correct alias', () => {
    const alias = command.alias();
    assert.strictEqual((alias && alias.indexOf(flowCommands.CONNECTOR_EXPORT) > -1), true);
  });

  it('exports the custom connectors', (done) => {
    let retrievedConnectorInfo = false;
    let retrievedSwagger = false;
    let retrievedIcon = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.PowerApps/apis/shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa?api-version=2016-11-01&$filter=environment%20eq%20%27Default-5be1aa17-e6cd-4d3d-8355-01af3e607d4b%27%20and%20IsCustomApi%20eq%20%27True%27`) > -1) {
        retrievedConnectorInfo = true;
        return Promise.resolve({ "name": "shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa", "id": "/providers/Microsoft.PowerApps/apis/shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa", "type": "Microsoft.PowerApps/apis", "properties": { "displayName": "Connector 1", "iconUri": "https://az787822.vo.msecnd.net/defaulticons/api-dedicated.png", "iconBrandColor": "#007ee5", "contact": {}, "license": {}, "apiEnvironment": "Shared", "isCustomApi": true, "connectionParameters": {}, "swagger": { "swagger": "2.0", "info": { "title": "Connector 1", "description": "", "version": "1.0" }, "host": "europe-002.azure-apim.net", "basePath": "/apim/connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa", "schemes": ["https"], "consumes": [], "produces": [], "paths": {}, "definitions": {}, "parameters": {}, "responses": {}, "securityDefinitions": {}, "security": [], "tags": [] }, "wadlUrl": "https://pafeblobprodam.blob.core.windows.net:443/apiwadls-6ee8be5d-ee5e-4dfa-b66a-81ef7afbaa1d/shared:2Dconnector:2D201:2D5f20a1f2d8d6777a75:%7C25F161FAF2ED7B7D?sv=2018-03-28&sr=c&sig=PPMiVV%2F%2FmsQ9uE5GI%2B2QSYix1ZVpaXT07MJVVDYIH2Y%3D&se=2020-01-15T21%3A43%3A38Z&sp=rl", "runtimeUrls": ["https://europe-002.azure-apim.net/apim/connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa"], "primaryRuntimeUrl": "https://europe-002.azure-apim.net/apim/connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa", "metadata": { "source": "powerapps-user-defined", "brandColor": "#007ee5", "contact": {}, "license": {}, "publisherUrl": null, "serviceUrl": null, "documentationUrl": null, "environmentName": "Default-5be1aa17-e6cd-4d3d-8355-01af3e607d4b", "xrmConnectorId": null, "almMode": "Environment", "createdBy": "{\"id\":\"9b974388-773f-4966-b27f-2e91c5916b18\",\"displayName\":\"MOD Administrator\",\"email\":\"admin@contoso.OnMicrosoft.com\",\"type\":\"User\",\"tenantId\":\"5be1aa17-e6cd-4d3d-8355-01af3e607d4b\",\"userPrincipalName\":\"admin@contoso.onmicrosoft.com\"}", "modifiedBy": "{\"id\":\"9b974388-773f-4966-b27f-2e91c5916b18\",\"displayName\":\"MOD Administrator\",\"email\":\"admin@contoso.OnMicrosoft.com\",\"type\":\"User\",\"tenantId\":\"5be1aa17-e6cd-4d3d-8355-01af3e607d4b\",\"userPrincipalName\":\"admin@contoso.onmicrosoft.com\"}", "allowSharing": false }, "capabilities": [], "description": "", "apiDefinitions": { "originalSwaggerUrl": "https://paeu2weu8.blob.core.windows.net/api-swagger-files/connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa.json_original?sv=2018-03-28&sr=b&sig=I5b3U5OxbeVYEfjosIU43HJbLqRB7mvZnE1E%2B1Hfeoc%3D&se=2020-01-15T10%3A43%3A38Z&sp=r", "modifiedSwaggerUrl": "https://paeu2weu8.blob.core.windows.net/api-swagger-files/connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa.json?sv=2018-03-28&sr=b&sig=KcB2aqBgcFF1VXnauF9%2B7KOXj8kPQIIayWLa0CtTQ8U%3D&se=2020-01-15T10%3A43%3A38Z&sp=r" }, "createdBy": { "id": "9b974388-773f-4966-b27f-2e91c5916b18", "displayName": "MOD Administrator", "email": "admin@contoso.OnMicrosoft.com", "type": "User", "tenantId": "5be1aa17-e6cd-4d3d-8355-01af3e607d4b", "userPrincipalName": "admin@contoso.onmicrosoft.com" }, "modifiedBy": { "id": "9b974388-773f-4966-b27f-2e91c5916b18", "displayName": "MOD Administrator", "email": "admin@contoso.OnMicrosoft.com", "type": "User", "tenantId": "5be1aa17-e6cd-4d3d-8355-01af3e607d4b", "userPrincipalName": "admin@contoso.onmicrosoft.com" }, "createdTime": "2019-12-18T18:51:32.3316756Z", "changedTime": "2019-12-18T18:51:32.3316756Z", "environment": { "id": "/providers/Microsoft.PowerApps/environments/Default-5be1aa17-e6cd-4d3d-8355-01af3e607d4b", "name": "Default-5be1aa17-e6cd-4d3d-8355-01af3e607d4b" }, "tier": "Standard", "publisher": "MOD Administrator", "almMode": "Environment" } });
      }
      else if (opts.url === 'https://paeu2weu8.blob.core.windows.net/api-swagger-files/connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa.json_original?sv=2018-03-28&sr=b&sig=I5b3U5OxbeVYEfjosIU43HJbLqRB7mvZnE1E%2B1Hfeoc%3D&se=2020-01-15T10%3A43%3A38Z&sp=r') {
        if (opts.headers &&
          opts.headers['x-anonymous'] === true) {
          retrievedSwagger = true;
          return Promise.resolve("{\r\n  \"swagger\": \"2.0\",\r\n  \"info\": {\r\n    \"title\": \"Connector 1\",\r\n    \"description\": \"\",\r\n    \"version\": \"1.0\"\r\n  },\r\n  \"host\": \"api.contoso.com\",\r\n  \"basePath\": \"/\",\r\n  \"schemes\": [\r\n    \"https\"\r\n  ],\r\n  \"consumes\": [],\r\n  \"produces\": [],\r\n  \"paths\": {},\r\n  \"definitions\": {},\r\n  \"parameters\": {},\r\n  \"responses\": {},\r\n  \"securityDefinitions\": {},\r\n  \"security\": [],\r\n  \"tags\": []\r\n}");
        }
        else {
          return Promise.reject('Invalid request');
        }
      }
      else if (opts.url === 'https://az787822.vo.msecnd.net/defaulticons/api-dedicated.png') {
        if (opts.headers &&
          opts.headers['x-anonymous'] === true) {
          retrievedIcon = true;
          return Promise.resolve('123');
        }
        else {
          return Promise.reject('Invalid request');
        }
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(fs, 'existsSync').callsFake(() => false);

    cmdInstance.action({ options: { debug: false, environment: 'Default-5be1aa17-e6cd-4d3d-8355-01af3e607d4b', connector: 'shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa' } }, () => {
      try {
        assert(retrievedConnectorInfo, 'Did not retrieve connector info');
        assert(retrievedSwagger, 'Did not retrieve swagger');
        assert(retrievedIcon, 'Did not retrieve icon');
        const outputFolder = path.resolve('shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa');
        assert(mkdirSyncStub.calledWith(outputFolder), 'Did not create folder in the right location');
        const settings = {
          apiDefinition: "apiDefinition.swagger.json",
          apiProperties: "apiProperties.json",
          connectorId: 'shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa',
          environment: 'Default-5be1aa17-e6cd-4d3d-8355-01af3e607d4b',
          icon: "icon.png",
          powerAppsApiVersion: "2016-11-01",
          powerAppsUrl: "https://api.powerapps.com"
        };
        assert(writeFileSyncStub.calledWithExactly(path.join(outputFolder, 'settings.json'), JSON.stringify(settings, null, 2), 'utf8'), 'Did not create correct settings.json file');
        const apiProperties = {
          properties: {
            "iconBrandColor": "#007ee5",
            "connectionParameters": {},
            "capabilities": []
          }
        };
        assert(writeFileSyncStub.calledWithExactly(path.join(outputFolder, 'apiProperties.json'), JSON.stringify(apiProperties, null, 2), 'utf8'), 'Did not create correct apiProperties.json file');
        const swagger = "{\r\n  \"swagger\": \"2.0\",\r\n  \"info\": {\r\n    \"title\": \"Connector 1\",\r\n    \"description\": \"\",\r\n    \"version\": \"1.0\"\r\n  },\r\n  \"host\": \"api.contoso.com\",\r\n  \"basePath\": \"/\",\r\n  \"schemes\": [\r\n    \"https\"\r\n  ],\r\n  \"consumes\": [],\r\n  \"produces\": [],\r\n  \"paths\": {},\r\n  \"definitions\": {},\r\n  \"parameters\": {},\r\n  \"responses\": {},\r\n  \"securityDefinitions\": {},\r\n  \"security\": [],\r\n  \"tags\": []\r\n}"
        assert(writeFileSyncStub.calledWithExactly(path.join(outputFolder, 'apiDefinition.swagger.json'), swagger, 'utf8'), 'Did not create correct apiDefinition.swagger.json file');
        assert(writeFileSyncStub.calledWith(path.join(outputFolder, 'icon.png')))
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('exports the custom connectors (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.PowerApps/apis/shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa?api-version=2016-11-01&$filter=environment%20eq%20%27Default-5be1aa17-e6cd-4d3d-8355-01af3e607d4b%27%20and%20IsCustomApi%20eq%20%27True%27`) > -1) {
        return Promise.resolve({ "name": "shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa", "id": "/providers/Microsoft.PowerApps/apis/shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa", "type": "Microsoft.PowerApps/apis", "properties": { "displayName": "Connector 1", "iconUri": "https://az787822.vo.msecnd.net/defaulticons/api-dedicated.png", "iconBrandColor": "#007ee5", "contact": {}, "license": {}, "apiEnvironment": "Shared", "isCustomApi": true, "connectionParameters": {}, "swagger": { "swagger": "2.0", "info": { "title": "Connector 1", "description": "", "version": "1.0" }, "host": "europe-002.azure-apim.net", "basePath": "/apim/connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa", "schemes": ["https"], "consumes": [], "produces": [], "paths": {}, "definitions": {}, "parameters": {}, "responses": {}, "securityDefinitions": {}, "security": [], "tags": [] }, "wadlUrl": "https://pafeblobprodam.blob.core.windows.net:443/apiwadls-6ee8be5d-ee5e-4dfa-b66a-81ef7afbaa1d/shared:2Dconnector:2D201:2D5f20a1f2d8d6777a75:%7C25F161FAF2ED7B7D?sv=2018-03-28&sr=c&sig=PPMiVV%2F%2FmsQ9uE5GI%2B2QSYix1ZVpaXT07MJVVDYIH2Y%3D&se=2020-01-15T21%3A43%3A38Z&sp=rl", "runtimeUrls": ["https://europe-002.azure-apim.net/apim/connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa"], "primaryRuntimeUrl": "https://europe-002.azure-apim.net/apim/connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa", "metadata": { "source": "powerapps-user-defined", "brandColor": "#007ee5", "contact": {}, "license": {}, "publisherUrl": null, "serviceUrl": null, "documentationUrl": null, "environmentName": "Default-5be1aa17-e6cd-4d3d-8355-01af3e607d4b", "xrmConnectorId": null, "almMode": "Environment", "createdBy": "{\"id\":\"9b974388-773f-4966-b27f-2e91c5916b18\",\"displayName\":\"MOD Administrator\",\"email\":\"admin@contoso.OnMicrosoft.com\",\"type\":\"User\",\"tenantId\":\"5be1aa17-e6cd-4d3d-8355-01af3e607d4b\",\"userPrincipalName\":\"admin@contoso.onmicrosoft.com\"}", "modifiedBy": "{\"id\":\"9b974388-773f-4966-b27f-2e91c5916b18\",\"displayName\":\"MOD Administrator\",\"email\":\"admin@contoso.OnMicrosoft.com\",\"type\":\"User\",\"tenantId\":\"5be1aa17-e6cd-4d3d-8355-01af3e607d4b\",\"userPrincipalName\":\"admin@contoso.onmicrosoft.com\"}", "allowSharing": false }, "capabilities": [], "description": "", "apiDefinitions": { "originalSwaggerUrl": "https://paeu2weu8.blob.core.windows.net/api-swagger-files/connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa.json_original?sv=2018-03-28&sr=b&sig=I5b3U5OxbeVYEfjosIU43HJbLqRB7mvZnE1E%2B1Hfeoc%3D&se=2020-01-15T10%3A43%3A38Z&sp=r", "modifiedSwaggerUrl": "https://paeu2weu8.blob.core.windows.net/api-swagger-files/connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa.json?sv=2018-03-28&sr=b&sig=KcB2aqBgcFF1VXnauF9%2B7KOXj8kPQIIayWLa0CtTQ8U%3D&se=2020-01-15T10%3A43%3A38Z&sp=r" }, "createdBy": { "id": "9b974388-773f-4966-b27f-2e91c5916b18", "displayName": "MOD Administrator", "email": "admin@contoso.OnMicrosoft.com", "type": "User", "tenantId": "5be1aa17-e6cd-4d3d-8355-01af3e607d4b", "userPrincipalName": "admin@contoso.onmicrosoft.com" }, "modifiedBy": { "id": "9b974388-773f-4966-b27f-2e91c5916b18", "displayName": "MOD Administrator", "email": "admin@contoso.OnMicrosoft.com", "type": "User", "tenantId": "5be1aa17-e6cd-4d3d-8355-01af3e607d4b", "userPrincipalName": "admin@contoso.onmicrosoft.com" }, "createdTime": "2019-12-18T18:51:32.3316756Z", "changedTime": "2019-12-18T18:51:32.3316756Z", "environment": { "id": "/providers/Microsoft.PowerApps/environments/Default-5be1aa17-e6cd-4d3d-8355-01af3e607d4b", "name": "Default-5be1aa17-e6cd-4d3d-8355-01af3e607d4b" }, "tier": "Standard", "publisher": "MOD Administrator", "almMode": "Environment" } });
      }
      else if (opts.url === 'https://paeu2weu8.blob.core.windows.net/api-swagger-files/connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa.json_original?sv=2018-03-28&sr=b&sig=I5b3U5OxbeVYEfjosIU43HJbLqRB7mvZnE1E%2B1Hfeoc%3D&se=2020-01-15T10%3A43%3A38Z&sp=r') {
        if (opts.headers &&
          opts.headers['x-anonymous'] === true) {
          return Promise.resolve("{\r\n  \"swagger\": \"2.0\",\r\n  \"info\": {\r\n    \"title\": \"Connector 1\",\r\n    \"description\": \"\",\r\n    \"version\": \"1.0\"\r\n  },\r\n  \"host\": \"api.contoso.com\",\r\n  \"basePath\": \"/\",\r\n  \"schemes\": [\r\n    \"https\"\r\n  ],\r\n  \"consumes\": [],\r\n  \"produces\": [],\r\n  \"paths\": {},\r\n  \"definitions\": {},\r\n  \"parameters\": {},\r\n  \"responses\": {},\r\n  \"securityDefinitions\": {},\r\n  \"security\": [],\r\n  \"tags\": []\r\n}");
        }
        else {
          return Promise.reject('Invalid request');
        }
      }
      else if (opts.url === 'https://az787822.vo.msecnd.net/defaulticons/api-dedicated.png') {
        if (opts.headers &&
          opts.headers['x-anonymous'] === true) {
          return Promise.resolve('123');
        }
        else {
          return Promise.reject('Invalid request');
        }
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(fs, 'existsSync').callsFake(() => false);

    cmdInstance.action({ options: { debug: true, environment: 'Default-5be1aa17-e6cd-4d3d-8355-01af3e607d4b', connector: 'shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWithExactly('Downloaded swagger'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when connector information misses properties', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.PowerApps/apis/shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa?api-version=2016-11-01&$filter=environment%20eq%20%27Default-5be1aa17-e6cd-4d3d-8355-01af3e607d4b%27%20and%20IsCustomApi%20eq%20%27True%27`) > -1) {
        return Promise.resolve({
          "name": "shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa",
          "id": "/providers/Microsoft.PowerApps/apis/shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa",
          "type": "Microsoft.PowerApps/apis",
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    cmdInstance.action({ options: { debug: false, environment: 'Default-5be1aa17-e6cd-4d3d-8355-01af3e607d4b', connector: 'shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Properties not present in the api registration information.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('skips downloading swagger if the connector information does not contain a swagger reference', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.PowerApps/apis/shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa?api-version=2016-11-01&$filter=environment%20eq%20%27Default-5be1aa17-e6cd-4d3d-8355-01af3e607d4b%27%20and%20IsCustomApi%20eq%20%27True%27`) > -1) {
        return Promise.resolve({
          "name": "shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa",
          "id": "/providers/Microsoft.PowerApps/apis/shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa",
          "type": "Microsoft.PowerApps/apis",
          "properties": {
            "displayName": "Connector 1",
            "iconUri": "https://az787822.vo.msecnd.net/defaulticons/api-dedicated.png",
            "iconBrandColor": "#007ee5",
            "contact": {},
            "license": {},
            "apiEnvironment": "Shared",
            "isCustomApi": true,
            "connectionParameters": {},
            "swagger": {
              "swagger": "2.0",
              "info": {
                "title": "Connector 1",
                "description": "",
                "version": "1.0"
              },
              "host": "europe-002.azure-apim.net",
              "basePath": "/apim/connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa",
              "schemes": [
                "https"
              ],
              "consumes": [],
              "produces": [],
              "paths": {},
              "definitions": {},
              "parameters": {},
              "responses": {},
              "securityDefinitions": {},
              "security": [],
              "tags": []
            },
            "wadlUrl": "https://pafeblobprodam.blob.core.windows.net:443/apiwadls-6ee8be5d-ee5e-4dfa-b66a-81ef7afbaa1d/shared:2Dconnector:2D201:2D5f20a1f2d8d6777a75:%7C25F161FAF2ED7B7D?sv=2018-03-28&sr=c&sig=PPMiVV%2F%2FmsQ9uE5GI%2B2QSYix1ZVpaXT07MJVVDYIH2Y%3D&se=2020-01-15T21%3A43%3A38Z&sp=rl",
            "runtimeUrls": [
              "https://europe-002.azure-apim.net/apim/connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa"
            ],
            "primaryRuntimeUrl": "https://europe-002.azure-apim.net/apim/connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa",
            "metadata": {
              "source": "powerapps-user-defined",
              "brandColor": "#007ee5",
              "contact": {},
              "license": {},
              "publisherUrl": null,
              "serviceUrl": null,
              "documentationUrl": null,
              "environmentName": "Default-5be1aa17-e6cd-4d3d-8355-01af3e607d4b",
              "xrmConnectorId": null,
              "almMode": "Environment",
              "createdBy": "{\"id\":\"9b974388-773f-4966-b27f-2e91c5916b18\",\"displayName\":\"MOD Administrator\",\"email\":\"admin@contoso.OnMicrosoft.com\",\"type\":\"User\",\"tenantId\":\"5be1aa17-e6cd-4d3d-8355-01af3e607d4b\",\"userPrincipalName\":\"admin@contoso.onmicrosoft.com\"}",
              "modifiedBy": "{\"id\":\"9b974388-773f-4966-b27f-2e91c5916b18\",\"displayName\":\"MOD Administrator\",\"email\":\"admin@contoso.OnMicrosoft.com\",\"type\":\"User\",\"tenantId\":\"5be1aa17-e6cd-4d3d-8355-01af3e607d4b\",\"userPrincipalName\":\"admin@contoso.onmicrosoft.com\"}",
              "allowSharing": false
            },
            "capabilities": [],
            "description": "",
            "createdBy": {
              "id": "9b974388-773f-4966-b27f-2e91c5916b18",
              "displayName": "MOD Administrator",
              "email": "admin@contoso.OnMicrosoft.com",
              "type": "User",
              "tenantId": "5be1aa17-e6cd-4d3d-8355-01af3e607d4b",
              "userPrincipalName": "admin@contoso.onmicrosoft.com"
            },
            "modifiedBy": {
              "id": "9b974388-773f-4966-b27f-2e91c5916b18",
              "displayName": "MOD Administrator",
              "email": "admin@contoso.OnMicrosoft.com",
              "type": "User",
              "tenantId": "5be1aa17-e6cd-4d3d-8355-01af3e607d4b",
              "userPrincipalName": "admin@contoso.onmicrosoft.com"
            },
            "createdTime": "2019-12-18T18:51:32.3316756Z",
            "changedTime": "2019-12-18T18:51:32.3316756Z",
            "environment": {
              "id": "/providers/Microsoft.PowerApps/environments/Default-5be1aa17-e6cd-4d3d-8355-01af3e607d4b",
              "name": "Default-5be1aa17-e6cd-4d3d-8355-01af3e607d4b"
            },
            "tier": "Standard",
            "publisher": "MOD Administrator",
            "almMode": "Environment"
          }
        });
      }
      else if (opts.url === 'https://az787822.vo.msecnd.net/defaulticons/api-dedicated.png') {
        if (opts.headers &&
          opts.headers['x-anonymous'] === true) {
          return Promise.resolve('123');
        }
        else {
          return Promise.reject('Invalid request');
        }
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    cmdInstance.action({ options: { debug: false, environment: 'Default-5be1aa17-e6cd-4d3d-8355-01af3e607d4b', connector: 'shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa' } }, (err?: any) => {
      try {
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('skips downloading swagger if the connector information does not contain a swagger reference (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.PowerApps/apis/shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa?api-version=2016-11-01&$filter=environment%20eq%20%27Default-5be1aa17-e6cd-4d3d-8355-01af3e607d4b%27%20and%20IsCustomApi%20eq%20%27True%27`) > -1) {
        return Promise.resolve({
          "name": "shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa",
          "id": "/providers/Microsoft.PowerApps/apis/shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa",
          "type": "Microsoft.PowerApps/apis",
          "properties": {
            "displayName": "Connector 1",
            "iconUri": "https://az787822.vo.msecnd.net/defaulticons/api-dedicated.png",
            "iconBrandColor": "#007ee5",
            "contact": {},
            "license": {},
            "apiEnvironment": "Shared",
            "isCustomApi": true,
            "connectionParameters": {},
            "swagger": {
              "swagger": "2.0",
              "info": {
                "title": "Connector 1",
                "description": "",
                "version": "1.0"
              },
              "host": "europe-002.azure-apim.net",
              "basePath": "/apim/connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa",
              "schemes": [
                "https"
              ],
              "consumes": [],
              "produces": [],
              "paths": {},
              "definitions": {},
              "parameters": {},
              "responses": {},
              "securityDefinitions": {},
              "security": [],
              "tags": []
            },
            "wadlUrl": "https://pafeblobprodam.blob.core.windows.net:443/apiwadls-6ee8be5d-ee5e-4dfa-b66a-81ef7afbaa1d/shared:2Dconnector:2D201:2D5f20a1f2d8d6777a75:%7C25F161FAF2ED7B7D?sv=2018-03-28&sr=c&sig=PPMiVV%2F%2FmsQ9uE5GI%2B2QSYix1ZVpaXT07MJVVDYIH2Y%3D&se=2020-01-15T21%3A43%3A38Z&sp=rl",
            "runtimeUrls": [
              "https://europe-002.azure-apim.net/apim/connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa"
            ],
            "primaryRuntimeUrl": "https://europe-002.azure-apim.net/apim/connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa",
            "metadata": {
              "source": "powerapps-user-defined",
              "brandColor": "#007ee5",
              "contact": {},
              "license": {},
              "publisherUrl": null,
              "serviceUrl": null,
              "documentationUrl": null,
              "environmentName": "Default-5be1aa17-e6cd-4d3d-8355-01af3e607d4b",
              "xrmConnectorId": null,
              "almMode": "Environment",
              "createdBy": "{\"id\":\"9b974388-773f-4966-b27f-2e91c5916b18\",\"displayName\":\"MOD Administrator\",\"email\":\"admin@contoso.OnMicrosoft.com\",\"type\":\"User\",\"tenantId\":\"5be1aa17-e6cd-4d3d-8355-01af3e607d4b\",\"userPrincipalName\":\"admin@contoso.onmicrosoft.com\"}",
              "modifiedBy": "{\"id\":\"9b974388-773f-4966-b27f-2e91c5916b18\",\"displayName\":\"MOD Administrator\",\"email\":\"admin@contoso.OnMicrosoft.com\",\"type\":\"User\",\"tenantId\":\"5be1aa17-e6cd-4d3d-8355-01af3e607d4b\",\"userPrincipalName\":\"admin@contoso.onmicrosoft.com\"}",
              "allowSharing": false
            },
            "capabilities": [],
            "description": "",
            "createdBy": {
              "id": "9b974388-773f-4966-b27f-2e91c5916b18",
              "displayName": "MOD Administrator",
              "email": "admin@contoso.OnMicrosoft.com",
              "type": "User",
              "tenantId": "5be1aa17-e6cd-4d3d-8355-01af3e607d4b",
              "userPrincipalName": "admin@contoso.onmicrosoft.com"
            },
            "modifiedBy": {
              "id": "9b974388-773f-4966-b27f-2e91c5916b18",
              "displayName": "MOD Administrator",
              "email": "admin@contoso.OnMicrosoft.com",
              "type": "User",
              "tenantId": "5be1aa17-e6cd-4d3d-8355-01af3e607d4b",
              "userPrincipalName": "admin@contoso.onmicrosoft.com"
            },
            "createdTime": "2019-12-18T18:51:32.3316756Z",
            "changedTime": "2019-12-18T18:51:32.3316756Z",
            "environment": {
              "id": "/providers/Microsoft.PowerApps/environments/Default-5be1aa17-e6cd-4d3d-8355-01af3e607d4b",
              "name": "Default-5be1aa17-e6cd-4d3d-8355-01af3e607d4b"
            },
            "tier": "Standard",
            "publisher": "MOD Administrator",
            "almMode": "Environment"
          }
        });
      }
      else if (opts.url === 'https://az787822.vo.msecnd.net/defaulticons/api-dedicated.png') {
        if (opts.headers &&
          opts.headers['x-anonymous'] === true) {
          return Promise.resolve('123');
        }
        else {
          return Promise.reject('Invalid request');
        }
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    cmdInstance.action({ options: { debug: true, environment: 'Default-5be1aa17-e6cd-4d3d-8355-01af3e607d4b', connector: 'shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa' } }, (err?: any) => {
      try {
        assert(cmdInstanceLogSpy.calledWith('originalSwaggerUrl not set. Skipping'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('skips downloading icon if the connector information does not contain icon URL', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.PowerApps/apis/shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa?api-version=2016-11-01&$filter=environment%20eq%20%27Default-5be1aa17-e6cd-4d3d-8355-01af3e607d4b%27%20and%20IsCustomApi%20eq%20%27True%27`) > -1) {
        return Promise.resolve({
          "name": "shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa",
          "id": "/providers/Microsoft.PowerApps/apis/shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa",
          "type": "Microsoft.PowerApps/apis",
          "properties": {
            "displayName": "Connector 1",
            "iconBrandColor": "#007ee5",
            "contact": {},
            "license": {},
            "apiEnvironment": "Shared",
            "isCustomApi": true,
            "connectionParameters": {},
            "swagger": {
              "swagger": "2.0",
              "info": {
                "title": "Connector 1",
                "description": "",
                "version": "1.0"
              },
              "host": "europe-002.azure-apim.net",
              "basePath": "/apim/connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa",
              "schemes": [
                "https"
              ],
              "consumes": [],
              "produces": [],
              "paths": {},
              "definitions": {},
              "parameters": {},
              "responses": {},
              "securityDefinitions": {},
              "security": [],
              "tags": []
            },
            "wadlUrl": "https://pafeblobprodam.blob.core.windows.net:443/apiwadls-6ee8be5d-ee5e-4dfa-b66a-81ef7afbaa1d/shared:2Dconnector:2D201:2D5f20a1f2d8d6777a75:%7C25F161FAF2ED7B7D?sv=2018-03-28&sr=c&sig=PPMiVV%2F%2FmsQ9uE5GI%2B2QSYix1ZVpaXT07MJVVDYIH2Y%3D&se=2020-01-15T21%3A43%3A38Z&sp=rl",
            "runtimeUrls": [
              "https://europe-002.azure-apim.net/apim/connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa"
            ],
            "primaryRuntimeUrl": "https://europe-002.azure-apim.net/apim/connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa",
            "metadata": {
              "source": "powerapps-user-defined",
              "brandColor": "#007ee5",
              "contact": {},
              "license": {},
              "publisherUrl": null,
              "serviceUrl": null,
              "documentationUrl": null,
              "environmentName": "Default-5be1aa17-e6cd-4d3d-8355-01af3e607d4b",
              "xrmConnectorId": null,
              "almMode": "Environment",
              "createdBy": "{\"id\":\"9b974388-773f-4966-b27f-2e91c5916b18\",\"displayName\":\"MOD Administrator\",\"email\":\"admin@contoso.OnMicrosoft.com\",\"type\":\"User\",\"tenantId\":\"5be1aa17-e6cd-4d3d-8355-01af3e607d4b\",\"userPrincipalName\":\"admin@contoso.onmicrosoft.com\"}",
              "modifiedBy": "{\"id\":\"9b974388-773f-4966-b27f-2e91c5916b18\",\"displayName\":\"MOD Administrator\",\"email\":\"admin@contoso.OnMicrosoft.com\",\"type\":\"User\",\"tenantId\":\"5be1aa17-e6cd-4d3d-8355-01af3e607d4b\",\"userPrincipalName\":\"admin@contoso.onmicrosoft.com\"}",
              "allowSharing": false
            },
            "capabilities": [],
            "description": "",
            "createdBy": {
              "id": "9b974388-773f-4966-b27f-2e91c5916b18",
              "displayName": "MOD Administrator",
              "email": "admin@contoso.OnMicrosoft.com",
              "type": "User",
              "tenantId": "5be1aa17-e6cd-4d3d-8355-01af3e607d4b",
              "userPrincipalName": "admin@contoso.onmicrosoft.com"
            },
            "modifiedBy": {
              "id": "9b974388-773f-4966-b27f-2e91c5916b18",
              "displayName": "MOD Administrator",
              "email": "admin@contoso.OnMicrosoft.com",
              "type": "User",
              "tenantId": "5be1aa17-e6cd-4d3d-8355-01af3e607d4b",
              "userPrincipalName": "admin@contoso.onmicrosoft.com"
            },
            "createdTime": "2019-12-18T18:51:32.3316756Z",
            "changedTime": "2019-12-18T18:51:32.3316756Z",
            "environment": {
              "id": "/providers/Microsoft.PowerApps/environments/Default-5be1aa17-e6cd-4d3d-8355-01af3e607d4b",
              "name": "Default-5be1aa17-e6cd-4d3d-8355-01af3e607d4b"
            },
            "tier": "Standard",
            "publisher": "MOD Administrator",
            "almMode": "Environment"
          }
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    cmdInstance.action({ options: { debug: false, environment: 'Default-5be1aa17-e6cd-4d3d-8355-01af3e607d4b', connector: 'shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa' } }, (err?: any) => {
      try {
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('skips downloading icon if the connector information does not contain icon URL (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.PowerApps/apis/shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa?api-version=2016-11-01&$filter=environment%20eq%20%27Default-5be1aa17-e6cd-4d3d-8355-01af3e607d4b%27%20and%20IsCustomApi%20eq%20%27True%27`) > -1) {
        return Promise.resolve({
          "name": "shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa",
          "id": "/providers/Microsoft.PowerApps/apis/shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa",
          "type": "Microsoft.PowerApps/apis",
          "properties": {
            "displayName": "Connector 1",
            "iconBrandColor": "#007ee5",
            "contact": {},
            "license": {},
            "apiEnvironment": "Shared",
            "isCustomApi": true,
            "connectionParameters": {},
            "swagger": {
              "swagger": "2.0",
              "info": {
                "title": "Connector 1",
                "description": "",
                "version": "1.0"
              },
              "host": "europe-002.azure-apim.net",
              "basePath": "/apim/connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa",
              "schemes": [
                "https"
              ],
              "consumes": [],
              "produces": [],
              "paths": {},
              "definitions": {},
              "parameters": {},
              "responses": {},
              "securityDefinitions": {},
              "security": [],
              "tags": []
            },
            "wadlUrl": "https://pafeblobprodam.blob.core.windows.net:443/apiwadls-6ee8be5d-ee5e-4dfa-b66a-81ef7afbaa1d/shared:2Dconnector:2D201:2D5f20a1f2d8d6777a75:%7C25F161FAF2ED7B7D?sv=2018-03-28&sr=c&sig=PPMiVV%2F%2FmsQ9uE5GI%2B2QSYix1ZVpaXT07MJVVDYIH2Y%3D&se=2020-01-15T21%3A43%3A38Z&sp=rl",
            "runtimeUrls": [
              "https://europe-002.azure-apim.net/apim/connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa"
            ],
            "primaryRuntimeUrl": "https://europe-002.azure-apim.net/apim/connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa",
            "metadata": {
              "source": "powerapps-user-defined",
              "brandColor": "#007ee5",
              "contact": {},
              "license": {},
              "publisherUrl": null,
              "serviceUrl": null,
              "documentationUrl": null,
              "environmentName": "Default-5be1aa17-e6cd-4d3d-8355-01af3e607d4b",
              "xrmConnectorId": null,
              "almMode": "Environment",
              "createdBy": "{\"id\":\"9b974388-773f-4966-b27f-2e91c5916b18\",\"displayName\":\"MOD Administrator\",\"email\":\"admin@contoso.OnMicrosoft.com\",\"type\":\"User\",\"tenantId\":\"5be1aa17-e6cd-4d3d-8355-01af3e607d4b\",\"userPrincipalName\":\"admin@contoso.onmicrosoft.com\"}",
              "modifiedBy": "{\"id\":\"9b974388-773f-4966-b27f-2e91c5916b18\",\"displayName\":\"MOD Administrator\",\"email\":\"admin@contoso.OnMicrosoft.com\",\"type\":\"User\",\"tenantId\":\"5be1aa17-e6cd-4d3d-8355-01af3e607d4b\",\"userPrincipalName\":\"admin@contoso.onmicrosoft.com\"}",
              "allowSharing": false
            },
            "capabilities": [],
            "description": "",
            "createdBy": {
              "id": "9b974388-773f-4966-b27f-2e91c5916b18",
              "displayName": "MOD Administrator",
              "email": "admin@contoso.OnMicrosoft.com",
              "type": "User",
              "tenantId": "5be1aa17-e6cd-4d3d-8355-01af3e607d4b",
              "userPrincipalName": "admin@contoso.onmicrosoft.com"
            },
            "modifiedBy": {
              "id": "9b974388-773f-4966-b27f-2e91c5916b18",
              "displayName": "MOD Administrator",
              "email": "admin@contoso.OnMicrosoft.com",
              "type": "User",
              "tenantId": "5be1aa17-e6cd-4d3d-8355-01af3e607d4b",
              "userPrincipalName": "admin@contoso.onmicrosoft.com"
            },
            "createdTime": "2019-12-18T18:51:32.3316756Z",
            "changedTime": "2019-12-18T18:51:32.3316756Z",
            "environment": {
              "id": "/providers/Microsoft.PowerApps/environments/Default-5be1aa17-e6cd-4d3d-8355-01af3e607d4b",
              "name": "Default-5be1aa17-e6cd-4d3d-8355-01af3e607d4b"
            },
            "tier": "Standard",
            "publisher": "MOD Administrator",
            "almMode": "Environment"
          }
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    cmdInstance.action({ options: { debug: true, environment: 'Default-5be1aa17-e6cd-4d3d-8355-01af3e607d4b', connector: 'shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa' } }, (err?: any) => {
      try {
        assert(cmdInstanceLogSpy.calledWith('iconUri not set. Skipping'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles environment not found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject({
        "error": {
          "code": "EnvironmentAccessDenied",
          "message": "The environment 'Default-5be1aa17-e6cd-4d3d-8355-01af3e607d4b' could not be found in the tenant '0d645e38-ec52-4a4f-ac58-65f2ac4015f6'."
        }
      });
    });

    cmdInstance.action({ options: { debug: false, environment: 'Default-5be1aa17-e6cd-4d3d-8355-01af3e607d4b', connector: 'shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The environment 'Default-5be1aa17-e6cd-4d3d-8355-01af3e607d4b' could not be found in the tenant '0d645e38-ec52-4a4f-ac58-65f2ac4015f6'.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles connector not found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject({
        "error": {
          "code": "ApiResourceNotFound",
          "message": "Could not find API 'shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfb'."
        }
      });
    });

    cmdInstance.action({ options: { debug: false, environment: 'Default-5be1aa17-e6cd-4d3d-8355-01af3e607d4b', connector: 'shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfb' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Could not find API 'shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfb'.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles OData API error', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject({
        error: {
          'odata.error': {
            code: '-1, InvalidOperationException',
            message: {
              value: 'An error has occurred'
            }
          }
        }
      });
    });

    cmdInstance.action({ options: { debug: false, environment: 'Default-5be1aa17-e6cd-4d3d-8355-01af3e607d4b', connector: 'shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation when the specified output folder does not exist', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const actual = (command.validate() as CommandValidate)({ options: { environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5', connector: 'shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa', outputFolder: '123' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the specified connector folder already exists', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    const actual = (command.validate() as CommandValidate)({ options: { environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5', connector: 'shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when all required options specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5', connector: 'shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when the specified output folder exists', () => {
    sinon.stub(fs, 'existsSync').callsFake((folder) => folder.toString().indexOf('connector') < 0);
    const actual = (command.validate() as CommandValidate)({ options: { environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5', connector: 'shared_connector-201-5f20a1f2d8d6777a75-5fa602f410652f4dfa', outputFolder: '123' } });
    assert.strictEqual(actual, true);
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

  it('supports specifying environment name', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--environment') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});