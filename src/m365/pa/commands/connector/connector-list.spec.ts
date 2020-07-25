import commands from '../../commands';
import flowCommands from '../../../flow/commands';
import Command, { CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./connector-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.CONNECTOR_LIST, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
    Utils.restore([
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CONNECTOR_LIST), true);
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
    assert.strictEqual((alias && alias.indexOf(flowCommands.CONNECTOR_LIST) > -1), true);
  });

  it('retrieves custom connectors (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.PowerApps/apis?api-version=2016-11-01&$filter=environment%20eq%20%27Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6%27%20and%20IsCustomApi%20eq%20%27True%27`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ "value": [{ "name": "shared_my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012", "id": "/providers/Microsoft.PowerApps/apis/shared_my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012", "type": "Microsoft.PowerApps/apis", "properties": { "displayName": "My connector 2", "iconUri": "https://az787822.vo.msecnd.net/defaulticons/api-dedicated.png", "iconBrandColor": "#007ee5", "contact": {}, "license": {}, "apiEnvironment": "Shared", "isCustomApi": true, "connectionParameters": {}, "runtimeUrls": ["https://europe-002.azure-apim.net/apim/my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012"], "primaryRuntimeUrl": "https://europe-002.azure-apim.net/apim/my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012", "metadata": { "source": "powerapps-user-defined", "brandColor": "#007ee5", "contact": {}, "license": {}, "publisherUrl": null, "serviceUrl": null, "documentationUrl": null, "environmentName": "Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "xrmConnectorId": null, "almMode": "Environment", "createdBy": "{\"id\":\"03043611-d01e-4e58-9fbe-1a18ecb861d8\",\"displayName\":\"MOD Administrator\",\"email\":\"admin@contoso.OnMicrosoft.com\",\"type\":\"User\",\"tenantId\":\"0d645e38-ec52-4a4f-ac58-65f2ac4015f6\",\"userPrincipalName\":\"admin@contoso.onmicrosoft.com\"}", "modifiedBy": "{\"id\":\"03043611-d01e-4e58-9fbe-1a18ecb861d8\",\"displayName\":\"MOD Administrator\",\"email\":\"admin@contoso.OnMicrosoft.com\",\"type\":\"User\",\"tenantId\":\"0d645e38-ec52-4a4f-ac58-65f2ac4015f6\",\"userPrincipalName\":\"admin@contoso.onmicrosoft.com\"}", "allowSharing": false }, "capabilities": [], "description": "", "apiDefinitions": { "originalSwaggerUrl": "https://paeu2weu8.blob.core.windows.net/api-swagger-files/my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012.json_original?sv=2018-03-28&sr=b&sig=rwJTmpMb4jb88Fzd9hoz8UbX0ZNbNiz5Cy5yfqTxcjU%3D&se=2019-12-05T19%3A53%3A49Z&sp=r", "modifiedSwaggerUrl": "https://paeu2weu8.blob.core.windows.net/api-swagger-files/my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012.json?sv=2018-03-28&sr=b&sig=eCW8GUjWHkcB8CFFQ%2FSZAGNBCZeAqj4H9ngRbA%2Fa4CI%3D&se=2019-12-05T19%3A53%3A49Z&sp=r" }, "createdBy": { "id": "03043611-d01e-4e58-9fbe-1a18ecb861d8", "displayName": "MOD Administrator", "email": "admin@contoso.OnMicrosoft.com", "type": "User", "tenantId": "0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "userPrincipalName": "admin@contoso.onmicrosoft.com" }, "modifiedBy": { "id": "03043611-d01e-4e58-9fbe-1a18ecb861d8", "displayName": "MOD Administrator", "email": "admin@contoso.OnMicrosoft.com", "type": "User", "tenantId": "0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "userPrincipalName": "admin@contoso.onmicrosoft.com" }, "createdTime": "2019-12-05T18:51:54.3261899Z", "changedTime": "2019-12-05T18:51:54.3261899Z", "environment": { "id": "/providers/Microsoft.PowerApps/environments/Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "name": "Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6" }, "tier": "Standard", "publisher": "MOD Administrator", "almMode": "Environment" } }, { "name": "shared_my-20connector-5f0027f520b23e81c1-5f9888a90360086012", "id": "/providers/Microsoft.PowerApps/apis/shared_my-20connector-5f0027f520b23e81c1-5f9888a90360086012", "type": "Microsoft.PowerApps/apis", "properties": { "displayName": "My connector", "iconUri": "https://az787822.vo.msecnd.net/defaulticons/api-dedicated.png", "iconBrandColor": "#007ee5", "contact": {}, "license": {}, "apiEnvironment": "Shared", "isCustomApi": true, "connectionParameters": {}, "runtimeUrls": ["https://europe-002.azure-apim.net/apim/my-20connector-5f0027f520b23e81c1-5f9888a90360086012"], "primaryRuntimeUrl": "https://europe-002.azure-apim.net/apim/my-20connector-5f0027f520b23e81c1-5f9888a90360086012", "metadata": { "source": "powerapps-user-defined", "brandColor": "#007ee5", "contact": {}, "license": {}, "publisherUrl": null, "serviceUrl": null, "documentationUrl": null, "environmentName": "Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "xrmConnectorId": null, "almMode": "Environment", "createdBy": "{\"id\":\"03043611-d01e-4e58-9fbe-1a18ecb861d8\",\"displayName\":\"MOD Administrator\",\"email\":\"admin@contoso.OnMicrosoft.com\",\"type\":\"User\",\"tenantId\":\"0d645e38-ec52-4a4f-ac58-65f2ac4015f6\",\"userPrincipalName\":\"admin@contoso.onmicrosoft.com\"}", "modifiedBy": "{\"id\":\"03043611-d01e-4e58-9fbe-1a18ecb861d8\",\"displayName\":\"MOD Administrator\",\"email\":\"admin@contoso.OnMicrosoft.com\",\"type\":\"User\",\"tenantId\":\"0d645e38-ec52-4a4f-ac58-65f2ac4015f6\",\"userPrincipalName\":\"admin@contoso.onmicrosoft.com\"}", "allowSharing": false }, "capabilities": [], "description": "", "apiDefinitions": { "originalSwaggerUrl": "https://paeu2weu8.blob.core.windows.net/api-swagger-files/my-20connector-5f0027f520b23e81c1-5f9888a90360086012.json_original?sv=2018-03-28&sr=b&sig=cOkjAecgpr6sSznMpDqiZitUOpVvVDJRCOZfe3VmReU%3D&se=2019-12-05T19%3A53%3A49Z&sp=r", "modifiedSwaggerUrl": "https://paeu2weu8.blob.core.windows.net/api-swagger-files/my-20connector-5f0027f520b23e81c1-5f9888a90360086012.json?sv=2018-03-28&sr=b&sig=rkpKHP8K%2F2yNBIUQcVN%2B0ZPjnP9sECrM%2FfoZMG%2BJZX0%3D&se=2019-12-05T19%3A53%3A49Z&sp=r" }, "createdBy": { "id": "03043611-d01e-4e58-9fbe-1a18ecb861d8", "displayName": "MOD Administrator", "email": "admin@contoso.OnMicrosoft.com", "type": "User", "tenantId": "0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "userPrincipalName": "admin@contoso.onmicrosoft.com" }, "modifiedBy": { "id": "03043611-d01e-4e58-9fbe-1a18ecb861d8", "displayName": "MOD Administrator", "email": "admin@contoso.OnMicrosoft.com", "type": "User", "tenantId": "0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "userPrincipalName": "admin@contoso.onmicrosoft.com" }, "createdTime": "2019-12-05T18:45:03.4615313Z", "changedTime": "2019-12-05T18:45:03.4615313Z", "environment": { "id": "/providers/Microsoft.PowerApps/environments/Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "name": "Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6" }, "tier": "Standard", "publisher": "MOD Administrator", "almMode": "Environment" } }] });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, environment: 'Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            name: 'shared_my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012',
            displayName: 'My connector 2'
          },
          {
            name: 'shared_my-20connector-5f0027f520b23e81c1-5f9888a90360086012',
            displayName: 'My connector'
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves custom connectors', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.PowerApps/apis?api-version=2016-11-01&$filter=environment%20eq%20%27Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6%27%20and%20IsCustomApi%20eq%20%27True%27`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ "value": [{ "name": "shared_my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012", "id": "/providers/Microsoft.PowerApps/apis/shared_my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012", "type": "Microsoft.PowerApps/apis", "properties": { "displayName": "My connector 2", "iconUri": "https://az787822.vo.msecnd.net/defaulticons/api-dedicated.png", "iconBrandColor": "#007ee5", "contact": {}, "license": {}, "apiEnvironment": "Shared", "isCustomApi": true, "connectionParameters": {}, "runtimeUrls": ["https://europe-002.azure-apim.net/apim/my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012"], "primaryRuntimeUrl": "https://europe-002.azure-apim.net/apim/my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012", "metadata": { "source": "powerapps-user-defined", "brandColor": "#007ee5", "contact": {}, "license": {}, "publisherUrl": null, "serviceUrl": null, "documentationUrl": null, "environmentName": "Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "xrmConnectorId": null, "almMode": "Environment", "createdBy": "{\"id\":\"03043611-d01e-4e58-9fbe-1a18ecb861d8\",\"displayName\":\"MOD Administrator\",\"email\":\"admin@contoso.OnMicrosoft.com\",\"type\":\"User\",\"tenantId\":\"0d645e38-ec52-4a4f-ac58-65f2ac4015f6\",\"userPrincipalName\":\"admin@contoso.onmicrosoft.com\"}", "modifiedBy": "{\"id\":\"03043611-d01e-4e58-9fbe-1a18ecb861d8\",\"displayName\":\"MOD Administrator\",\"email\":\"admin@contoso.OnMicrosoft.com\",\"type\":\"User\",\"tenantId\":\"0d645e38-ec52-4a4f-ac58-65f2ac4015f6\",\"userPrincipalName\":\"admin@contoso.onmicrosoft.com\"}", "allowSharing": false }, "capabilities": [], "description": "", "apiDefinitions": { "originalSwaggerUrl": "https://paeu2weu8.blob.core.windows.net/api-swagger-files/my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012.json_original?sv=2018-03-28&sr=b&sig=rwJTmpMb4jb88Fzd9hoz8UbX0ZNbNiz5Cy5yfqTxcjU%3D&se=2019-12-05T19%3A53%3A49Z&sp=r", "modifiedSwaggerUrl": "https://paeu2weu8.blob.core.windows.net/api-swagger-files/my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012.json?sv=2018-03-28&sr=b&sig=eCW8GUjWHkcB8CFFQ%2FSZAGNBCZeAqj4H9ngRbA%2Fa4CI%3D&se=2019-12-05T19%3A53%3A49Z&sp=r" }, "createdBy": { "id": "03043611-d01e-4e58-9fbe-1a18ecb861d8", "displayName": "MOD Administrator", "email": "admin@contoso.OnMicrosoft.com", "type": "User", "tenantId": "0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "userPrincipalName": "admin@contoso.onmicrosoft.com" }, "modifiedBy": { "id": "03043611-d01e-4e58-9fbe-1a18ecb861d8", "displayName": "MOD Administrator", "email": "admin@contoso.OnMicrosoft.com", "type": "User", "tenantId": "0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "userPrincipalName": "admin@contoso.onmicrosoft.com" }, "createdTime": "2019-12-05T18:51:54.3261899Z", "changedTime": "2019-12-05T18:51:54.3261899Z", "environment": { "id": "/providers/Microsoft.PowerApps/environments/Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "name": "Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6" }, "tier": "Standard", "publisher": "MOD Administrator", "almMode": "Environment" } }, { "name": "shared_my-20connector-5f0027f520b23e81c1-5f9888a90360086012", "id": "/providers/Microsoft.PowerApps/apis/shared_my-20connector-5f0027f520b23e81c1-5f9888a90360086012", "type": "Microsoft.PowerApps/apis", "properties": { "displayName": "My connector", "iconUri": "https://az787822.vo.msecnd.net/defaulticons/api-dedicated.png", "iconBrandColor": "#007ee5", "contact": {}, "license": {}, "apiEnvironment": "Shared", "isCustomApi": true, "connectionParameters": {}, "runtimeUrls": ["https://europe-002.azure-apim.net/apim/my-20connector-5f0027f520b23e81c1-5f9888a90360086012"], "primaryRuntimeUrl": "https://europe-002.azure-apim.net/apim/my-20connector-5f0027f520b23e81c1-5f9888a90360086012", "metadata": { "source": "powerapps-user-defined", "brandColor": "#007ee5", "contact": {}, "license": {}, "publisherUrl": null, "serviceUrl": null, "documentationUrl": null, "environmentName": "Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "xrmConnectorId": null, "almMode": "Environment", "createdBy": "{\"id\":\"03043611-d01e-4e58-9fbe-1a18ecb861d8\",\"displayName\":\"MOD Administrator\",\"email\":\"admin@contoso.OnMicrosoft.com\",\"type\":\"User\",\"tenantId\":\"0d645e38-ec52-4a4f-ac58-65f2ac4015f6\",\"userPrincipalName\":\"admin@contoso.onmicrosoft.com\"}", "modifiedBy": "{\"id\":\"03043611-d01e-4e58-9fbe-1a18ecb861d8\",\"displayName\":\"MOD Administrator\",\"email\":\"admin@contoso.OnMicrosoft.com\",\"type\":\"User\",\"tenantId\":\"0d645e38-ec52-4a4f-ac58-65f2ac4015f6\",\"userPrincipalName\":\"admin@contoso.onmicrosoft.com\"}", "allowSharing": false }, "capabilities": [], "description": "", "apiDefinitions": { "originalSwaggerUrl": "https://paeu2weu8.blob.core.windows.net/api-swagger-files/my-20connector-5f0027f520b23e81c1-5f9888a90360086012.json_original?sv=2018-03-28&sr=b&sig=cOkjAecgpr6sSznMpDqiZitUOpVvVDJRCOZfe3VmReU%3D&se=2019-12-05T19%3A53%3A49Z&sp=r", "modifiedSwaggerUrl": "https://paeu2weu8.blob.core.windows.net/api-swagger-files/my-20connector-5f0027f520b23e81c1-5f9888a90360086012.json?sv=2018-03-28&sr=b&sig=rkpKHP8K%2F2yNBIUQcVN%2B0ZPjnP9sECrM%2FfoZMG%2BJZX0%3D&se=2019-12-05T19%3A53%3A49Z&sp=r" }, "createdBy": { "id": "03043611-d01e-4e58-9fbe-1a18ecb861d8", "displayName": "MOD Administrator", "email": "admin@contoso.OnMicrosoft.com", "type": "User", "tenantId": "0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "userPrincipalName": "admin@contoso.onmicrosoft.com" }, "modifiedBy": { "id": "03043611-d01e-4e58-9fbe-1a18ecb861d8", "displayName": "MOD Administrator", "email": "admin@contoso.OnMicrosoft.com", "type": "User", "tenantId": "0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "userPrincipalName": "admin@contoso.onmicrosoft.com" }, "createdTime": "2019-12-05T18:45:03.4615313Z", "changedTime": "2019-12-05T18:45:03.4615313Z", "environment": { "id": "/providers/Microsoft.PowerApps/environments/Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "name": "Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6" }, "tier": "Standard", "publisher": "MOD Administrator", "almMode": "Environment" } }] });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, environment: 'Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            name: 'shared_my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012',
            displayName: 'My connector 2'
          },
          {
            name: 'shared_my-20connector-5f0027f520b23e81c1-5f9888a90360086012',
            displayName: 'My connector'
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves custom connectors in pages', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('skiptoken') === -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            "nextLink": "https://management.azure.com/providers/Microsoft.PowerApps/apis?api-version=2016-11-01&$filter=environment%20eq%20%27Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6%27%20and%20IsCustomApi%20eq%20%27True%27&%24skiptoken=eyJuZXh0TWFya2VyIjoiMjAxOTAyMDRUMTg1NDU2Wi02YTA5NGQwMi02NDFhLTQ4OTEtYjRkZi00NDA1OTRmMjZjODUifQ%3d%3d",
            "value": [
              { "name": "shared_my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012", "id": "/providers/Microsoft.PowerApps/apis/shared_my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012", "type": "Microsoft.PowerApps/apis", "properties": { "displayName": "My connector 2", "iconUri": "https://az787822.vo.msecnd.net/defaulticons/api-dedicated.png", "iconBrandColor": "#007ee5", "contact": {}, "license": {}, "apiEnvironment": "Shared", "isCustomApi": true, "connectionParameters": {}, "runtimeUrls": ["https://europe-002.azure-apim.net/apim/my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012"], "primaryRuntimeUrl": "https://europe-002.azure-apim.net/apim/my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012", "metadata": { "source": "powerapps-user-defined", "brandColor": "#007ee5", "contact": {}, "license": {}, "publisherUrl": null, "serviceUrl": null, "documentationUrl": null, "environmentName": "Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "xrmConnectorId": null, "almMode": "Environment", "createdBy": "{\"id\":\"03043611-d01e-4e58-9fbe-1a18ecb861d8\",\"displayName\":\"MOD Administrator\",\"email\":\"admin@contoso.OnMicrosoft.com\",\"type\":\"User\",\"tenantId\":\"0d645e38-ec52-4a4f-ac58-65f2ac4015f6\",\"userPrincipalName\":\"admin@contoso.onmicrosoft.com\"}", "modifiedBy": "{\"id\":\"03043611-d01e-4e58-9fbe-1a18ecb861d8\",\"displayName\":\"MOD Administrator\",\"email\":\"admin@contoso.OnMicrosoft.com\",\"type\":\"User\",\"tenantId\":\"0d645e38-ec52-4a4f-ac58-65f2ac4015f6\",\"userPrincipalName\":\"admin@contoso.onmicrosoft.com\"}", "allowSharing": false }, "capabilities": [], "description": "", "apiDefinitions": { "originalSwaggerUrl": "https://paeu2weu8.blob.core.windows.net/api-swagger-files/my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012.json_original?sv=2018-03-28&sr=b&sig=rwJTmpMb4jb88Fzd9hoz8UbX0ZNbNiz5Cy5yfqTxcjU%3D&se=2019-12-05T19%3A53%3A49Z&sp=r", "modifiedSwaggerUrl": "https://paeu2weu8.blob.core.windows.net/api-swagger-files/my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012.json?sv=2018-03-28&sr=b&sig=eCW8GUjWHkcB8CFFQ%2FSZAGNBCZeAqj4H9ngRbA%2Fa4CI%3D&se=2019-12-05T19%3A53%3A49Z&sp=r" }, "createdBy": { "id": "03043611-d01e-4e58-9fbe-1a18ecb861d8", "displayName": "MOD Administrator", "email": "admin@contoso.OnMicrosoft.com", "type": "User", "tenantId": "0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "userPrincipalName": "admin@contoso.onmicrosoft.com" }, "modifiedBy": { "id": "03043611-d01e-4e58-9fbe-1a18ecb861d8", "displayName": "MOD Administrator", "email": "admin@contoso.OnMicrosoft.com", "type": "User", "tenantId": "0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "userPrincipalName": "admin@contoso.onmicrosoft.com" }, "createdTime": "2019-12-05T18:51:54.3261899Z", "changedTime": "2019-12-05T18:51:54.3261899Z", "environment": { "id": "/providers/Microsoft.PowerApps/environments/Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "name": "Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6" }, "tier": "Standard", "publisher": "MOD Administrator", "almMode": "Environment" } }
            ]
          });
        }
      }
      else {
        return Promise.resolve({
          "value": [
            { "name": "shared_my-20connector-5f0027f520b23e81c1-5f9888a90360086012", "id": "/providers/Microsoft.PowerApps/apis/shared_my-20connector-5f0027f520b23e81c1-5f9888a90360086012", "type": "Microsoft.PowerApps/apis", "properties": { "displayName": "My connector", "iconUri": "https://az787822.vo.msecnd.net/defaulticons/api-dedicated.png", "iconBrandColor": "#007ee5", "contact": {}, "license": {}, "apiEnvironment": "Shared", "isCustomApi": true, "connectionParameters": {}, "runtimeUrls": ["https://europe-002.azure-apim.net/apim/my-20connector-5f0027f520b23e81c1-5f9888a90360086012"], "primaryRuntimeUrl": "https://europe-002.azure-apim.net/apim/my-20connector-5f0027f520b23e81c1-5f9888a90360086012", "metadata": { "source": "powerapps-user-defined", "brandColor": "#007ee5", "contact": {}, "license": {}, "publisherUrl": null, "serviceUrl": null, "documentationUrl": null, "environmentName": "Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "xrmConnectorId": null, "almMode": "Environment", "createdBy": "{\"id\":\"03043611-d01e-4e58-9fbe-1a18ecb861d8\",\"displayName\":\"MOD Administrator\",\"email\":\"admin@contoso.OnMicrosoft.com\",\"type\":\"User\",\"tenantId\":\"0d645e38-ec52-4a4f-ac58-65f2ac4015f6\",\"userPrincipalName\":\"admin@contoso.onmicrosoft.com\"}", "modifiedBy": "{\"id\":\"03043611-d01e-4e58-9fbe-1a18ecb861d8\",\"displayName\":\"MOD Administrator\",\"email\":\"admin@contoso.OnMicrosoft.com\",\"type\":\"User\",\"tenantId\":\"0d645e38-ec52-4a4f-ac58-65f2ac4015f6\",\"userPrincipalName\":\"admin@contoso.onmicrosoft.com\"}", "allowSharing": false }, "capabilities": [], "description": "", "apiDefinitions": { "originalSwaggerUrl": "https://paeu2weu8.blob.core.windows.net/api-swagger-files/my-20connector-5f0027f520b23e81c1-5f9888a90360086012.json_original?sv=2018-03-28&sr=b&sig=cOkjAecgpr6sSznMpDqiZitUOpVvVDJRCOZfe3VmReU%3D&se=2019-12-05T19%3A53%3A49Z&sp=r", "modifiedSwaggerUrl": "https://paeu2weu8.blob.core.windows.net/api-swagger-files/my-20connector-5f0027f520b23e81c1-5f9888a90360086012.json?sv=2018-03-28&sr=b&sig=rkpKHP8K%2F2yNBIUQcVN%2B0ZPjnP9sECrM%2FfoZMG%2BJZX0%3D&se=2019-12-05T19%3A53%3A49Z&sp=r" }, "createdBy": { "id": "03043611-d01e-4e58-9fbe-1a18ecb861d8", "displayName": "MOD Administrator", "email": "admin@contoso.OnMicrosoft.com", "type": "User", "tenantId": "0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "userPrincipalName": "admin@contoso.onmicrosoft.com" }, "modifiedBy": { "id": "03043611-d01e-4e58-9fbe-1a18ecb861d8", "displayName": "MOD Administrator", "email": "admin@contoso.OnMicrosoft.com", "type": "User", "tenantId": "0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "userPrincipalName": "admin@contoso.onmicrosoft.com" }, "createdTime": "2019-12-05T18:45:03.4615313Z", "changedTime": "2019-12-05T18:45:03.4615313Z", "environment": { "id": "/providers/Microsoft.PowerApps/environments/Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "name": "Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6" }, "tier": "Standard", "publisher": "MOD Administrator", "almMode": "Environment" } }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, environment: 'Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            name: 'shared_my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012',
            displayName: 'My connector 2'
          },
          {
            name: 'shared_my-20connector-5f0027f520b23e81c1-5f9888a90360086012',
            displayName: 'My connector'
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('outputs all properties when output is JSON', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.PowerApps/apis?api-version=2016-11-01&$filter=environment%20eq%20%27Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6%27%20and%20IsCustomApi%20eq%20%27True%27`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ "value": [{ "name": "shared_my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012", "id": "/providers/Microsoft.PowerApps/apis/shared_my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012", "type": "Microsoft.PowerApps/apis", "properties": { "displayName": "My connector 2", "iconUri": "https://az787822.vo.msecnd.net/defaulticons/api-dedicated.png", "iconBrandColor": "#007ee5", "contact": {}, "license": {}, "apiEnvironment": "Shared", "isCustomApi": true, "connectionParameters": {}, "runtimeUrls": ["https://europe-002.azure-apim.net/apim/my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012"], "primaryRuntimeUrl": "https://europe-002.azure-apim.net/apim/my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012", "metadata": { "source": "powerapps-user-defined", "brandColor": "#007ee5", "contact": {}, "license": {}, "publisherUrl": null, "serviceUrl": null, "documentationUrl": null, "environmentName": "Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "xrmConnectorId": null, "almMode": "Environment", "createdBy": "{\"id\":\"03043611-d01e-4e58-9fbe-1a18ecb861d8\",\"displayName\":\"MOD Administrator\",\"email\":\"admin@contoso.OnMicrosoft.com\",\"type\":\"User\",\"tenantId\":\"0d645e38-ec52-4a4f-ac58-65f2ac4015f6\",\"userPrincipalName\":\"admin@contoso.onmicrosoft.com\"}", "modifiedBy": "{\"id\":\"03043611-d01e-4e58-9fbe-1a18ecb861d8\",\"displayName\":\"MOD Administrator\",\"email\":\"admin@contoso.OnMicrosoft.com\",\"type\":\"User\",\"tenantId\":\"0d645e38-ec52-4a4f-ac58-65f2ac4015f6\",\"userPrincipalName\":\"admin@contoso.onmicrosoft.com\"}", "allowSharing": false }, "capabilities": [], "description": "", "apiDefinitions": { "originalSwaggerUrl": "https://paeu2weu8.blob.core.windows.net/api-swagger-files/my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012.json_original?sv=2018-03-28&sr=b&sig=rwJTmpMb4jb88Fzd9hoz8UbX0ZNbNiz5Cy5yfqTxcjU%3D&se=2019-12-05T19%3A53%3A49Z&sp=r", "modifiedSwaggerUrl": "https://paeu2weu8.blob.core.windows.net/api-swagger-files/my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012.json?sv=2018-03-28&sr=b&sig=eCW8GUjWHkcB8CFFQ%2FSZAGNBCZeAqj4H9ngRbA%2Fa4CI%3D&se=2019-12-05T19%3A53%3A49Z&sp=r" }, "createdBy": { "id": "03043611-d01e-4e58-9fbe-1a18ecb861d8", "displayName": "MOD Administrator", "email": "admin@contoso.OnMicrosoft.com", "type": "User", "tenantId": "0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "userPrincipalName": "admin@contoso.onmicrosoft.com" }, "modifiedBy": { "id": "03043611-d01e-4e58-9fbe-1a18ecb861d8", "displayName": "MOD Administrator", "email": "admin@contoso.OnMicrosoft.com", "type": "User", "tenantId": "0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "userPrincipalName": "admin@contoso.onmicrosoft.com" }, "createdTime": "2019-12-05T18:51:54.3261899Z", "changedTime": "2019-12-05T18:51:54.3261899Z", "environment": { "id": "/providers/Microsoft.PowerApps/environments/Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "name": "Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6" }, "tier": "Standard", "publisher": "MOD Administrator", "almMode": "Environment" } }, { "name": "shared_my-20connector-5f0027f520b23e81c1-5f9888a90360086012", "id": "/providers/Microsoft.PowerApps/apis/shared_my-20connector-5f0027f520b23e81c1-5f9888a90360086012", "type": "Microsoft.PowerApps/apis", "properties": { "displayName": "My connector", "iconUri": "https://az787822.vo.msecnd.net/defaulticons/api-dedicated.png", "iconBrandColor": "#007ee5", "contact": {}, "license": {}, "apiEnvironment": "Shared", "isCustomApi": true, "connectionParameters": {}, "runtimeUrls": ["https://europe-002.azure-apim.net/apim/my-20connector-5f0027f520b23e81c1-5f9888a90360086012"], "primaryRuntimeUrl": "https://europe-002.azure-apim.net/apim/my-20connector-5f0027f520b23e81c1-5f9888a90360086012", "metadata": { "source": "powerapps-user-defined", "brandColor": "#007ee5", "contact": {}, "license": {}, "publisherUrl": null, "serviceUrl": null, "documentationUrl": null, "environmentName": "Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "xrmConnectorId": null, "almMode": "Environment", "createdBy": "{\"id\":\"03043611-d01e-4e58-9fbe-1a18ecb861d8\",\"displayName\":\"MOD Administrator\",\"email\":\"admin@contoso.OnMicrosoft.com\",\"type\":\"User\",\"tenantId\":\"0d645e38-ec52-4a4f-ac58-65f2ac4015f6\",\"userPrincipalName\":\"admin@contoso.onmicrosoft.com\"}", "modifiedBy": "{\"id\":\"03043611-d01e-4e58-9fbe-1a18ecb861d8\",\"displayName\":\"MOD Administrator\",\"email\":\"admin@contoso.OnMicrosoft.com\",\"type\":\"User\",\"tenantId\":\"0d645e38-ec52-4a4f-ac58-65f2ac4015f6\",\"userPrincipalName\":\"admin@contoso.onmicrosoft.com\"}", "allowSharing": false }, "capabilities": [], "description": "", "apiDefinitions": { "originalSwaggerUrl": "https://paeu2weu8.blob.core.windows.net/api-swagger-files/my-20connector-5f0027f520b23e81c1-5f9888a90360086012.json_original?sv=2018-03-28&sr=b&sig=cOkjAecgpr6sSznMpDqiZitUOpVvVDJRCOZfe3VmReU%3D&se=2019-12-05T19%3A53%3A49Z&sp=r", "modifiedSwaggerUrl": "https://paeu2weu8.blob.core.windows.net/api-swagger-files/my-20connector-5f0027f520b23e81c1-5f9888a90360086012.json?sv=2018-03-28&sr=b&sig=rkpKHP8K%2F2yNBIUQcVN%2B0ZPjnP9sECrM%2FfoZMG%2BJZX0%3D&se=2019-12-05T19%3A53%3A49Z&sp=r" }, "createdBy": { "id": "03043611-d01e-4e58-9fbe-1a18ecb861d8", "displayName": "MOD Administrator", "email": "admin@contoso.OnMicrosoft.com", "type": "User", "tenantId": "0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "userPrincipalName": "admin@contoso.onmicrosoft.com" }, "modifiedBy": { "id": "03043611-d01e-4e58-9fbe-1a18ecb861d8", "displayName": "MOD Administrator", "email": "admin@contoso.OnMicrosoft.com", "type": "User", "tenantId": "0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "userPrincipalName": "admin@contoso.onmicrosoft.com" }, "createdTime": "2019-12-05T18:45:03.4615313Z", "changedTime": "2019-12-05T18:45:03.4615313Z", "environment": { "id": "/providers/Microsoft.PowerApps/environments/Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "name": "Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6" }, "tier": "Standard", "publisher": "MOD Administrator", "almMode": "Environment" } }] });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, environment: 'Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6', output: 'json' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([{ "name": "shared_my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012", "id": "/providers/Microsoft.PowerApps/apis/shared_my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012", "type": "Microsoft.PowerApps/apis", "properties": { "displayName": "My connector 2", "iconUri": "https://az787822.vo.msecnd.net/defaulticons/api-dedicated.png", "iconBrandColor": "#007ee5", "contact": {}, "license": {}, "apiEnvironment": "Shared", "isCustomApi": true, "connectionParameters": {}, "runtimeUrls": ["https://europe-002.azure-apim.net/apim/my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012"], "primaryRuntimeUrl": "https://europe-002.azure-apim.net/apim/my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012", "metadata": { "source": "powerapps-user-defined", "brandColor": "#007ee5", "contact": {}, "license": {}, "publisherUrl": null, "serviceUrl": null, "documentationUrl": null, "environmentName": "Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "xrmConnectorId": null, "almMode": "Environment", "createdBy": "{\"id\":\"03043611-d01e-4e58-9fbe-1a18ecb861d8\",\"displayName\":\"MOD Administrator\",\"email\":\"admin@contoso.OnMicrosoft.com\",\"type\":\"User\",\"tenantId\":\"0d645e38-ec52-4a4f-ac58-65f2ac4015f6\",\"userPrincipalName\":\"admin@contoso.onmicrosoft.com\"}", "modifiedBy": "{\"id\":\"03043611-d01e-4e58-9fbe-1a18ecb861d8\",\"displayName\":\"MOD Administrator\",\"email\":\"admin@contoso.OnMicrosoft.com\",\"type\":\"User\",\"tenantId\":\"0d645e38-ec52-4a4f-ac58-65f2ac4015f6\",\"userPrincipalName\":\"admin@contoso.onmicrosoft.com\"}", "allowSharing": false }, "capabilities": [], "description": "", "apiDefinitions": { "originalSwaggerUrl": "https://paeu2weu8.blob.core.windows.net/api-swagger-files/my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012.json_original?sv=2018-03-28&sr=b&sig=rwJTmpMb4jb88Fzd9hoz8UbX0ZNbNiz5Cy5yfqTxcjU%3D&se=2019-12-05T19%3A53%3A49Z&sp=r", "modifiedSwaggerUrl": "https://paeu2weu8.blob.core.windows.net/api-swagger-files/my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012.json?sv=2018-03-28&sr=b&sig=eCW8GUjWHkcB8CFFQ%2FSZAGNBCZeAqj4H9ngRbA%2Fa4CI%3D&se=2019-12-05T19%3A53%3A49Z&sp=r" }, "createdBy": { "id": "03043611-d01e-4e58-9fbe-1a18ecb861d8", "displayName": "MOD Administrator", "email": "admin@contoso.OnMicrosoft.com", "type": "User", "tenantId": "0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "userPrincipalName": "admin@contoso.onmicrosoft.com" }, "modifiedBy": { "id": "03043611-d01e-4e58-9fbe-1a18ecb861d8", "displayName": "MOD Administrator", "email": "admin@contoso.OnMicrosoft.com", "type": "User", "tenantId": "0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "userPrincipalName": "admin@contoso.onmicrosoft.com" }, "createdTime": "2019-12-05T18:51:54.3261899Z", "changedTime": "2019-12-05T18:51:54.3261899Z", "environment": { "id": "/providers/Microsoft.PowerApps/environments/Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "name": "Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6" }, "tier": "Standard", "publisher": "MOD Administrator", "almMode": "Environment" } }, { "name": "shared_my-20connector-5f0027f520b23e81c1-5f9888a90360086012", "id": "/providers/Microsoft.PowerApps/apis/shared_my-20connector-5f0027f520b23e81c1-5f9888a90360086012", "type": "Microsoft.PowerApps/apis", "properties": { "displayName": "My connector", "iconUri": "https://az787822.vo.msecnd.net/defaulticons/api-dedicated.png", "iconBrandColor": "#007ee5", "contact": {}, "license": {}, "apiEnvironment": "Shared", "isCustomApi": true, "connectionParameters": {}, "runtimeUrls": ["https://europe-002.azure-apim.net/apim/my-20connector-5f0027f520b23e81c1-5f9888a90360086012"], "primaryRuntimeUrl": "https://europe-002.azure-apim.net/apim/my-20connector-5f0027f520b23e81c1-5f9888a90360086012", "metadata": { "source": "powerapps-user-defined", "brandColor": "#007ee5", "contact": {}, "license": {}, "publisherUrl": null, "serviceUrl": null, "documentationUrl": null, "environmentName": "Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "xrmConnectorId": null, "almMode": "Environment", "createdBy": "{\"id\":\"03043611-d01e-4e58-9fbe-1a18ecb861d8\",\"displayName\":\"MOD Administrator\",\"email\":\"admin@contoso.OnMicrosoft.com\",\"type\":\"User\",\"tenantId\":\"0d645e38-ec52-4a4f-ac58-65f2ac4015f6\",\"userPrincipalName\":\"admin@contoso.onmicrosoft.com\"}", "modifiedBy": "{\"id\":\"03043611-d01e-4e58-9fbe-1a18ecb861d8\",\"displayName\":\"MOD Administrator\",\"email\":\"admin@contoso.OnMicrosoft.com\",\"type\":\"User\",\"tenantId\":\"0d645e38-ec52-4a4f-ac58-65f2ac4015f6\",\"userPrincipalName\":\"admin@contoso.onmicrosoft.com\"}", "allowSharing": false }, "capabilities": [], "description": "", "apiDefinitions": { "originalSwaggerUrl": "https://paeu2weu8.blob.core.windows.net/api-swagger-files/my-20connector-5f0027f520b23e81c1-5f9888a90360086012.json_original?sv=2018-03-28&sr=b&sig=cOkjAecgpr6sSznMpDqiZitUOpVvVDJRCOZfe3VmReU%3D&se=2019-12-05T19%3A53%3A49Z&sp=r", "modifiedSwaggerUrl": "https://paeu2weu8.blob.core.windows.net/api-swagger-files/my-20connector-5f0027f520b23e81c1-5f9888a90360086012.json?sv=2018-03-28&sr=b&sig=rkpKHP8K%2F2yNBIUQcVN%2B0ZPjnP9sECrM%2FfoZMG%2BJZX0%3D&se=2019-12-05T19%3A53%3A49Z&sp=r" }, "createdBy": { "id": "03043611-d01e-4e58-9fbe-1a18ecb861d8", "displayName": "MOD Administrator", "email": "admin@contoso.OnMicrosoft.com", "type": "User", "tenantId": "0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "userPrincipalName": "admin@contoso.onmicrosoft.com" }, "modifiedBy": { "id": "03043611-d01e-4e58-9fbe-1a18ecb861d8", "displayName": "MOD Administrator", "email": "admin@contoso.OnMicrosoft.com", "type": "User", "tenantId": "0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "userPrincipalName": "admin@contoso.onmicrosoft.com" }, "createdTime": "2019-12-05T18:45:03.4615313Z", "changedTime": "2019-12-05T18:45:03.4615313Z", "environment": { "id": "/providers/Microsoft.PowerApps/environments/Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "name": "Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6" }, "tier": "Standard", "publisher": "MOD Administrator", "almMode": "Environment" } }]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no environment found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject({
        "error": {
          "code": "EnvironmentAccessDenied",
          "message": "The environment 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' could not be found in the tenant '0d645e38-ec52-4a4f-ac58-65f2ac4015f6'."
        }
      });
    });

    cmdInstance.action({ options: { debug: false, environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The environment 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' could not be found in the tenant '0d645e38-ec52-4a4f-ac58-65f2ac4015f6'.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no custom connectors found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.resolve({ value: [] });
    });

    cmdInstance.action({ options: { debug: false, environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no custom connectors found (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.resolve({ value: [] });
    });

    cmdInstance.action({ options: { debug: true, environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith('No custom connectors found'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles API OData error', (done) => {
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

    cmdInstance.action({ options: { debug: false, environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when retrieving the second page of data', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('skiptoken') === -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            "nextLink": "https://management.azure.com/providers/Microsoft.PowerApps/apis?api-version=2016-11-01&$filter=environment%20eq%20%27Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6%27%20and%20IsCustomApi%20eq%20%27True%27&%24skiptoken=eyJuZXh0TWFya2VyIjoiMjAxOTAyMDRUMTg1NDU2Wi02YTA5NGQwMi02NDFhLTQ4OTEtYjRkZi00NDA1OTRmMjZjODUifQ%3d%3d",
            "value": [
              { "name": "shared_my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012", "id": "/providers/Microsoft.PowerApps/apis/shared_my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012", "type": "Microsoft.PowerApps/apis", "properties": { "displayName": "My connector 2", "iconUri": "https://az787822.vo.msecnd.net/defaulticons/api-dedicated.png", "iconBrandColor": "#007ee5", "contact": {}, "license": {}, "apiEnvironment": "Shared", "isCustomApi": true, "connectionParameters": {}, "runtimeUrls": ["https://europe-002.azure-apim.net/apim/my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012"], "primaryRuntimeUrl": "https://europe-002.azure-apim.net/apim/my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012", "metadata": { "source": "powerapps-user-defined", "brandColor": "#007ee5", "contact": {}, "license": {}, "publisherUrl": null, "serviceUrl": null, "documentationUrl": null, "environmentName": "Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "xrmConnectorId": null, "almMode": "Environment", "createdBy": "{\"id\":\"03043611-d01e-4e58-9fbe-1a18ecb861d8\",\"displayName\":\"MOD Administrator\",\"email\":\"admin@contoso.OnMicrosoft.com\",\"type\":\"User\",\"tenantId\":\"0d645e38-ec52-4a4f-ac58-65f2ac4015f6\",\"userPrincipalName\":\"admin@contoso.onmicrosoft.com\"}", "modifiedBy": "{\"id\":\"03043611-d01e-4e58-9fbe-1a18ecb861d8\",\"displayName\":\"MOD Administrator\",\"email\":\"admin@contoso.OnMicrosoft.com\",\"type\":\"User\",\"tenantId\":\"0d645e38-ec52-4a4f-ac58-65f2ac4015f6\",\"userPrincipalName\":\"admin@contoso.onmicrosoft.com\"}", "allowSharing": false }, "capabilities": [], "description": "", "apiDefinitions": { "originalSwaggerUrl": "https://paeu2weu8.blob.core.windows.net/api-swagger-files/my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012.json_original?sv=2018-03-28&sr=b&sig=rwJTmpMb4jb88Fzd9hoz8UbX0ZNbNiz5Cy5yfqTxcjU%3D&se=2019-12-05T19%3A53%3A49Z&sp=r", "modifiedSwaggerUrl": "https://paeu2weu8.blob.core.windows.net/api-swagger-files/my-20connector-202-5f0027f520b23e81c1-5f9888a90360086012.json?sv=2018-03-28&sr=b&sig=eCW8GUjWHkcB8CFFQ%2FSZAGNBCZeAqj4H9ngRbA%2Fa4CI%3D&se=2019-12-05T19%3A53%3A49Z&sp=r" }, "createdBy": { "id": "03043611-d01e-4e58-9fbe-1a18ecb861d8", "displayName": "MOD Administrator", "email": "admin@contoso.OnMicrosoft.com", "type": "User", "tenantId": "0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "userPrincipalName": "admin@contoso.onmicrosoft.com" }, "modifiedBy": { "id": "03043611-d01e-4e58-9fbe-1a18ecb861d8", "displayName": "MOD Administrator", "email": "admin@contoso.OnMicrosoft.com", "type": "User", "tenantId": "0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "userPrincipalName": "admin@contoso.onmicrosoft.com" }, "createdTime": "2019-12-05T18:51:54.3261899Z", "changedTime": "2019-12-05T18:51:54.3261899Z", "environment": { "id": "/providers/Microsoft.PowerApps/environments/Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6", "name": "Default-0d645e38-ec52-4a4f-ac58-65f2ac4015f6" }, "tier": "Standard", "publisher": "MOD Administrator", "almMode": "Environment" } }
            ]
          });
        }
      }
      else {
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
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5' } }, (err?: any) => {
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