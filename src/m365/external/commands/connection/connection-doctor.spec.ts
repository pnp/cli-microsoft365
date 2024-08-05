import { ExternalConnectors, SearchResponse } from '@microsoft/microsoft-graph-types';
import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { ODataResponse } from '../../../../utils/odata.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './connection-doctor.js';

describe(commands.CONNECTION_DOCTOR, () => {
  const logger: Logger = {
    log: async () => { },
    logRaw: async () => { },
    logToStderr: async () => { }
  };
  let commandInfo: CommandInfo;
  let loggerLogSpy: sinon.SinonSpy;

  const externalConnection: ExternalConnectors.ExternalConnection = {
    "id": "msgraphdocs",
    "name": "Microsoft Graph documentation",
    "description": "Documentation for Microsoft Graph API which explains what Microsoft Graph is and how to use it.",
    "state": "ready",
    "configuration": {
      "authorizedAppIds": [
        "4f56e2be-b7ae-4b30-9fd0-c42d6bdf8255"
      ]
    },
    "searchSettings": {
      "searchResultTemplates": [
        {
          "id": "msgraphdocs",
          "priority": 1,
          "layout": {
            "type": "AdaptiveCard",
            "version": "1.3",
            "body": [
              {
                "type": "ColumnSet",
                "columns": [
                  {
                    "type": "Column",
                    "width": "auto",
                    "items": [
                      {
                        "type": "Image",
                        "url": "https://raw.githubusercontent.com/waldekmastykarz/img/main/microsoft-graph.png",
                        "altText": "Thumbnail image",
                        "horizontalAlignment": "center",
                        "size": "small"
                      }
                    ],
                    "horizontalAlignment": "center"
                  },
                  {
                    "type": "Column",
                    "width": "stretch",
                    "items": [
                      {
                        "type": "TextBlock",
                        "text": "[${title}](${url})",
                        "weight": "bolder",
                        "color": "accent",
                        "size": "medium",
                        "maxLines": 3
                      },
                      {
                        "type": "TextBlock",
                        "text": "[${url}](${url})",
                        "weight": "bolder",
                        "spacing": "small",
                        "maxLines": 3
                      },
                      {
                        "type": "TextBlock",
                        "text": "${description}",
                        "maxLines": 3,
                        "wrap": true
                      }
                    ],
                    "spacing": "medium"
                  }
                ]
              }
            ],
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "$data": {
              "description": "Marketing team at Contoso.., and looking at the Contoso Marketing documents on the team site. This contains the data from FY20 and will taken over to FY21...Marketing Planning is ongoing for FY20..",
              "url": "https://contoso.com",
              "title": "Contoso Solutions"
            }
          },
          "rules": []
        }
      ]
    },
    "activitySettings": {
      "urlToItemResolvers": [
        {
          "priority": 1,
          "itemId": "auth__{slug}",
          "urlMatchInfo": {
            "baseUrls": [
              "https://learn.microsoft.com"
            ],
            "urlPattern": "/[^/]+/graph/auth/(?<slug>[^/]+)"
          }
        },
        {
          "priority": 2,
          "itemId": "sdks__{slug}",
          "urlMatchInfo": {
            "baseUrls": [
              "https://learn.microsoft.com"
            ],
            "urlPattern": "/[^/]+/graph/sdks/(?<slug>[^/]+)"
          }
        },
        {
          "priority": 3,
          "itemId": "{slug}",
          "urlMatchInfo": {
            "baseUrls": [
              "https://learn.microsoft.com"
            ],
            "urlPattern": "/[^/]+/graph/(?<slug>[^/]+)"
          }
        } as ExternalConnectors.ItemIdResolver
      ]
    }
  };
  const schema: ExternalConnectors.Schema = {
    "baseType": "microsoft.graph.externalItem",
    "properties": [
      {
        "name": "title",
        "type": "string",
        "isSearchable": true,
        "isRetrievable": true,
        "isQueryable": true,
        "labels": [
          "title"
        ],
        "isRefinable": false,
        "aliases": []
      },
      {
        "name": "iconUrl",
        "type": "string",
        "isSearchable": false,
        "isRetrievable": true,
        "isQueryable": false,
        "labels": ["iconUrl"],
        "isRefinable": false,
        "aliases": []
      },
      {
        "name": "url",
        "type": "string",
        "isSearchable": false,
        "isRetrievable": true,
        "isQueryable": false,
        "labels": [
          "url"
        ],
        "isRefinable": false,
        "aliases": []
      }
    ]
  };
  const searchResponse: ODataResponse<SearchResponse> = {
    "value": [
      {
        "searchTerms": [],
        "hitsContainers": [
          {
            "hits": [
              {
                "hitId": "AAMkAGE3YTcwZDFiLTgzMmItNGQxYi1hOWZjLWRiZWQzMjJlN2VlMABGAAAAAACkN3ckZNrUSIhtfaVyEdYWBwB0XsiLKGwtTIwX0CYw7PKyAAAAAAEwAAB0XsiLKGwtTIwX0CYw7PKyAAAAABzKAAA=",
                "contentSource": "msgraphdocs",
                "rank": 1,
                "summary": "<c0>Create</c0> <c0>a</c0> <c0>Microsoft</c0> <c0>Graph</c0> <c0>client</c0> <c0>The</c0> <c0>Microsoft</c0> <c0>Graph</c0> <c0>client</c0> <c0>is</c0> <c0>designed</c0> <c0>to</c0> <c0>make</c0> <c0>it</c0> <c0>simple</c0> <ddd/> <c0>make</c0> <c0>calls</c0> <c0>to</c0> <c0>Microsoft</c0> <c0>Graph</c0>. <c0>You</c0> <c0>can</c0> <c0>use</c0> <c0>a</c0> <c0>single</c0> <c0>client</c0> <c0>instance</c0> <c0>for</c0> <c0>the</c0> <c0>lifetime</c0> <c0>of</c0> <ddd/>",
                "resource": {
                  "@odata.type": "#microsoft.graph.externalConnectors.externalItem",
                  "properties": {
                    "title": "Create a Microsoft Graph client",
                    "url": "https://learn.microsoft.com/graph/sdks/create-client",
                    "description": "Describes how to create a client to use to make calls to Microsoft Graph. Includes how to set up authentication and select a sovereign cloud.",
                    "substrateContentDomainId": "444d0a3470e44d468899fb3dcfcda99c@02b03c8c-a55c-4f23-9285-d5bd8f81979a,sdks__create-client",
                    "substrateLocationId": "APP:GCS_msgraphdocs@02b03c8c-a55c-4f23-9285-d5bd8f81979a",
                    "documentId": 2012,
                    "immutableEntryId": "AAAAAB2EAxGqZhHNm8gAqgAvxFoNAHReyIsobC1MjBfQJjDs8rIAAAAAILIAAA==",
                    "id": "AAMkAGE3YTcwZDFiLTgzMmItNGQxYi1hOWZjLWRiZWQzMjJlN2VlMABGAAAAAACkN3ckZNrUSIhtfaVyEdYWBwB0XsiLKGwtTIwX0CYw7PKyAAAAAAEwAAB0XsiLKGwtTIwX0CYw7PKyAAAAABzKAAA="
                  }
                } as any
              }
            ],
            "total": 15,
            "moreResultsAvailable": true
          }
        ]
      }
    ]
  };
  const externalItem: ExternalConnectors.ExternalItem = {
    "id": "sdks__create-client",
    "acl": [
      {
        "type": "everyone",
        "value": "everyone",
        "accessType": "grant"
      }
    ],
    "properties": {
      "title": "Create a Microsoft Graph client",
      "description": "Describes how to create a client to use to make calls to Microsoft Graph. Includes how to set up authentication and select a sovereign cloud.",
      "url": "https://learn.microsoft.com/graph/sdks/create-client"
    },
    "content": {
      "value": "Create a Microsoft Graph client\nThe Microsoft Graph client is designed to make it simple to make calls to Microsoft Graph. You can use a single client instance for the lifetime of the application. For information about how to add and install the Microsoft Graph client package into your project, see  Install the SDK.\nThe following code examples show how to create an instance of a Microsoft Graph client with an authentication provider in the supported languages. The authentication provider will handle acquiring access tokens for the application. Many different authentication providers are available for each language and platform. The different authentication providers support different client scenarios. For details about which provider and options are appropriate for your scenario, see Choose an Authentication Provider.\n<!-- markdownlint-disable MD025 MD051 -->\nC#\n:::code language=&quot;csharp&quot; source=&quot;./snippets/dotnet/src/SdkSnippets/Snippets/CreateClients.cs&quot; id=&quot;DeviceCodeSnippet&quot;:::\nGo\n:::code language=&quot;go&quot; source=&quot;./snippets/go/src/snippets/create_clients.go&quot; id=&quot;DeviceCodeSnippet&quot;:::\nJava\n:::code language=&quot;java&quot; source=&quot;./snippets/java/app/src/main/java/snippets/CreateClients.java&quot; id=&quot;DeviceCodeSnippet&quot;:::\nPHP\n// PHP client currently doesn't have an authentication provider. You will need to handle\n// getting an access token. The following example demonstrates the client credential\n// OAuth flow and assumes that an administrator has consented to the application.\n$guzzle = new \\GuzzleHttp\\Client();\n$url = 'https://login.microsoftonline.com/' . $tenantId . '/oauth2/token?api-version=1.0';\n$token = json_decode($guzzle-post($url, [\n    'form_params' = [\n        'client_id' = $clientId,\n        'client_secret' = $clientSecret,\n        'resource' = 'https://graph.microsoft.com/',\n        'grant_type' = 'client_credentials',\n    ],\n])-getBody()-getContents());\n$accessToken = $token-access_token;\n\n// Create a new Graph client.\n$graph = new Graph();\n$graph-setAccessToken($accessToken);\n\n// Make a call to /me Graph resource.\n$user = $graph-createRequest(GET, /me)\n              -setReturnType(Model\\User::class)\n              -execute();\nPython\n[!INCLUDE python-sdk-preview]\n:::code language=&quot;python&quot; source=&quot;./snippets/python/src/snippets/create_clients.py&quot; id=&quot;DeviceCodeSnippet&quot;:::\nTypeScript\n:::code language=&quot;typescript&quot; source=&quot;./snippets/typescript/src/snippets/createClients.ts&quot; id=&quot;DeviceCodeSnippet&quot;:::\n",
      "type": "html"
    }
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post
    ]);
    loggerLogSpy.resetHistory();
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CONNECTION_DOCTOR);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('checks connection compatibility for copilot', async () => {
    sinon.stub(request, 'get').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs`) {
        return externalConnection;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs/schema`) {
        return schema;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs/items/sdks__create-client`) {
        return externalItem;
      }
      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/search/query') {
        return searchResponse;
      }
      throw 'Invalid request';
    });
    const options: any = {
      id: 'msgraphdocs',
      ux: 'copilot',
      output: 'json'
    };
    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith([
      {
        "id": "loadExternalConnection",
        "text": "Load connection",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "loadSchema",
        "text": "Load schema",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "copilotRequiredSemanticLabels",
        "text": "Required semantic labels",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "searchableProperties",
        "text": "Searchable properties",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "contentIngested",
        "text": "Items have content ingested",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "enabledForInlineResults",
        "text": "Connection configured for inline results",
        "type": "required",
        "status": "manual"
      },
      {
        "id": "itemsHaveActivities",
        "text": "Items have activities recorded",
        "type": "recommended",
        "status": "manual"
      },
      {
        "id": "meaningfulNameAndDescription",
        "text": "Meaningful connection name and description",
        "type": "required",
        "status": "manual"
      },
      {
        "id": "urlToItemResolver",
        "text": "urlToItemResolver configured",
        "type": "recommended",
        "status": "passed"
      }
    ]));
  });

  it('checks connection compatibility for copilot in debug mode', async () => {
    sinon.stub(request, 'get').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs`) {
        return externalConnection;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs/schema`) {
        return schema;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs/items/sdks__create-client`) {
        return externalItem;
      }
      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/search/query') {
        return searchResponse;
      }
      throw 'Invalid request';
    });
    const options: any = {
      id: 'msgraphdocs',
      ux: 'copilot',
      output: 'json',
      debug: true
    };
    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith([
      {
        "id": "loadExternalConnection",
        "text": "Load connection",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "loadSchema",
        "text": "Load schema",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "copilotRequiredSemanticLabels",
        "text": "Required semantic labels",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "searchableProperties",
        "text": "Searchable properties",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "contentIngested",
        "text": "Items have content ingested",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "enabledForInlineResults",
        "text": "Connection configured for inline results",
        "type": "required",
        "status": "manual"
      },
      {
        "id": "itemsHaveActivities",
        "text": "Items have activities recorded",
        "type": "recommended",
        "status": "manual"
      },
      {
        "id": "meaningfulNameAndDescription",
        "text": "Meaningful connection name and description",
        "type": "required",
        "status": "manual"
      },
      {
        "id": "urlToItemResolver",
        "text": "urlToItemResolver configured",
        "type": "recommended",
        "status": "passed"
      }
    ]));
  });

  it('checks connection compatibility for search', async () => {
    sinon.stub(request, 'get').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs`) {
        return externalConnection;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs/schema`) {
        return schema;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs/items/sdks__create-client`) {
        return externalItem;
      }
      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/search/query') {
        return searchResponse;
      }
      throw 'Invalid request';
    });
    const options: any = {
      id: 'msgraphdocs',
      ux: 'search',
      output: 'json'
    };
    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith([
      {
        "id": "loadExternalConnection",
        "text": "Load connection",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "loadSchema",
        "text": "Load schema",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "semanticLabels",
        "text": "Semantic labels",
        "type": "recommended",
        "status": "passed"
      },
      {
        "id": "searchableProperties",
        "text": "Searchable properties",
        "type": "recommended",
        "status": "passed"
      },
      {
        "id": "resultType",
        "text": "Result type",
        "type": "recommended",
        "status": "passed"
      },
      {
        "id": "contentIngested",
        "text": "Items have content ingested",
        "type": "recommended",
        "status": "passed"
      },
      {
        "id": "itemsHaveActivities",
        "text": "Items have activities recorded",
        "type": "recommended",
        "status": "manual"
      },
      {
        "id": "urlToItemResolver",
        "text": "urlToItemResolver configured",
        "type": "recommended",
        "status": "passed"
      }
    ]));
  });

  it('checks connection compatibility for all ux', async () => {
    sinon.stub(request, 'get').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs`) {
        return externalConnection;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs/schema`) {
        return schema;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs/items/sdks__create-client`) {
        return externalItem;
      }
      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/search/query') {
        return searchResponse;
      }
      throw 'Invalid request';
    });
    const options: any = {
      id: 'msgraphdocs',
      ux: 'all',
      output: 'json'
    };
    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith([
      {
        "id": "loadExternalConnection",
        "text": "Load connection",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "loadSchema",
        "text": "Load schema",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "copilotRequiredSemanticLabels",
        "text": "Required semantic labels",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "searchableProperties",
        "text": "Searchable properties",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "contentIngested",
        "text": "Items have content ingested",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "enabledForInlineResults",
        "text": "Connection configured for inline results",
        "type": "required",
        "status": "manual"
      },
      {
        "id": "itemsHaveActivities",
        "text": "Items have activities recorded",
        "type": "recommended",
        "status": "manual"
      },
      {
        "id": "meaningfulNameAndDescription",
        "text": "Meaningful connection name and description",
        "type": "required",
        "status": "manual"
      },
      {
        "id": "semanticLabels",
        "text": "Semantic labels",
        "type": "recommended",
        "status": "passed"
      },
      {
        "id": "resultType",
        "text": "Result type",
        "type": "recommended",
        "status": "passed"
      },
      {
        "id": "urlToItemResolver",
        "text": "urlToItemResolver configured",
        "type": "recommended",
        "status": "passed"
      }
    ]));
  });

  it('defaults to checking compatibility with all ux when no ux specified', async () => {
    sinon.stub(request, 'get').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs`) {
        return externalConnection;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs/schema`) {
        return schema;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs/items/sdks__create-client`) {
        return externalItem;
      }
      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/search/query') {
        return searchResponse;
      }
      throw 'Invalid request';
    });
    const options: any = {
      id: 'msgraphdocs',
      output: 'json'
    };
    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith([
      {
        "id": "loadExternalConnection",
        "text": "Load connection",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "loadSchema",
        "text": "Load schema",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "copilotRequiredSemanticLabels",
        "text": "Required semantic labels",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "searchableProperties",
        "text": "Searchable properties",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "contentIngested",
        "text": "Items have content ingested",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "enabledForInlineResults",
        "text": "Connection configured for inline results",
        "type": "required",
        "status": "manual"
      },
      {
        "id": "itemsHaveActivities",
        "text": "Items have activities recorded",
        "type": "recommended",
        "status": "manual"
      },
      {
        "id": "meaningfulNameAndDescription",
        "text": "Meaningful connection name and description",
        "type": "required",
        "status": "manual"
      },
      {
        "id": "semanticLabels",
        "text": "Semantic labels",
        "type": "recommended",
        "status": "passed"
      },
      {
        "id": "resultType",
        "text": "Result type",
        "type": "recommended",
        "status": "passed"
      },
      {
        "id": "urlToItemResolver",
        "text": "urlToItemResolver configured",
        "type": "recommended",
        "status": "passed"
      }
    ]));
  });

  it('stops when the specified connection is not found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs`) {
        return Promise.reject({
          response: {
            data: {
              "error": {
                "code": "ItemNotFound",
                "message": "The requested resource does not exist.",
                "innerError": {
                  "code": "ResourceNotFound",
                  "source": "Connection",
                  "message": "'Connection' was not found.",
                  "date": "2023-11-22T18:05:10",
                  "request-id": "a844db09-1ab3-4209-aea8-ee5b8b72949a",
                  "client-request-id": "a844db09-1ab3-4209-aea8-ee5b8b72949a"
                }
              }
            }
          }
        });
      }
      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async () => {
      throw 'Invalid request';
    });
    const options: any = {
      id: 'msgraphdocs',
      ux: 'copilot',
      output: 'json'
    };
    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith([
      {
        "id": "loadExternalConnection",
        "text": "Load connection",
        "type": "required",
        "error": "'Connection' was not found.",
        "errorMessage": "Connection not found",
        "status": "failed"
      }
    ]));
  });

  it(`stops when couldn't retrieve the schema`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs`) {
        return externalConnection;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs/schema`) {
        return Promise.reject({
          response: {
            data: {
              "error": {
                "code": "ItemNotFound",
                "message": "An error has occurred.",
                "innerError": {
                  "code": "ResourceNotFound",
                  "source": "Schema",
                  "message": "An error has occurred.",
                  "date": "2023-11-22T18:05:10",
                  "request-id": "a844db09-1ab3-4209-aea8-ee5b8b72949a",
                  "client-request-id": "a844db09-1ab3-4209-aea8-ee5b8b72949a"
                }
              }
            }
          }
        });
      }
      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async () => {
      throw 'Invalid request';
    });
    const options: any = {
      id: 'msgraphdocs',
      ux: 'copilot',
      output: 'json'
    };
    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith([
      {
        "id": "loadExternalConnection",
        "text": "Load connection",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "loadSchema",
        "text": "Load schema",
        "type": "required",
        "error": "An error has occurred.",
        "errorMessage": "Schema not found",
        "status": "failed"
      }
    ]));
  });

  it(`returns an error when one of the semantic labels required for copilot are missing`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs`) {
        return externalConnection;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs/schema`) {
        return {
          "baseType": "microsoft.graph.externalItem",
          "properties": [
            {
              "name": "title",
              "type": "string",
              "isSearchable": true,
              "isRetrievable": true,
              "isQueryable": true,
              "labels": [
                "title"
              ],
              "isRefinable": false,
              "aliases": []
            },
            {
              "name": "iconUrl",
              "type": "string",
              "isSearchable": false,
              "isRetrievable": true,
              "isQueryable": false,
              "labels": [],
              "isRefinable": false,
              "aliases": []
            },
            {
              "name": "url",
              "type": "string",
              "isSearchable": false,
              "isRetrievable": true,
              "isQueryable": false,
              "labels": [
                "url"
              ],
              "isRefinable": false,
              "aliases": []
            }
          ]
        };
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs/items/sdks__create-client`) {
        return externalItem;
      }
      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/search/query') {
        return searchResponse;
      }
      throw 'Invalid request';
    });
    const options: any = {
      id: 'msgraphdocs',
      ux: 'copilot',
      output: 'json'
    };
    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith([
      {
        "id": "loadExternalConnection",
        "text": "Load connection",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "loadSchema",
        "text": "Load schema",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "copilotRequiredSemanticLabels",
        "text": "Required semantic labels",
        "type": "required",
        "errorMessage": "Missing label iconUrl",
        "status": "failed"
      },
      {
        "id": "searchableProperties",
        "text": "Searchable properties",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "contentIngested",
        "text": "Items have content ingested",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "enabledForInlineResults",
        "text": "Connection configured for inline results",
        "type": "required",
        "status": "manual"
      },
      {
        "id": "itemsHaveActivities",
        "text": "Items have activities recorded",
        "type": "recommended",
        "status": "manual"
      },
      {
        "id": "meaningfulNameAndDescription",
        "text": "Meaningful connection name and description",
        "type": "required",
        "status": "manual"
      },
      {
        "id": "urlToItemResolver",
        "text": "urlToItemResolver configured",
        "type": "recommended",
        "status": "passed"
      }
    ]));
  });

  it(`returns an error when none of the schema properties are searchable`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs`) {
        return externalConnection;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs/schema`) {
        return {
          "baseType": "microsoft.graph.externalItem",
          "properties": [
            {
              "name": "title",
              "type": "string",
              "isSearchable": false,
              "isRetrievable": true,
              "isQueryable": true,
              "labels": [
                "title"
              ],
              "isRefinable": false,
              "aliases": []
            },
            {
              "name": "iconUrl",
              "type": "string",
              "isSearchable": false,
              "isRetrievable": true,
              "isQueryable": false,
              "labels": [
                "iconUrl"
              ],
              "isRefinable": false,
              "aliases": []
            },
            {
              "name": "url",
              "type": "string",
              "isSearchable": false,
              "isRetrievable": true,
              "isQueryable": false,
              "labels": [
                "url"
              ],
              "isRefinable": false,
              "aliases": []
            }
          ]
        };
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs/items/sdks__create-client`) {
        return externalItem;
      }
      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/search/query') {
        return searchResponse;
      }
      throw 'Invalid request';
    });
    const options: any = {
      id: 'msgraphdocs',
      ux: 'copilot',
      output: 'json'
    };
    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith([
      {
        "id": "loadExternalConnection",
        "text": "Load connection",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "loadSchema",
        "text": "Load schema",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "copilotRequiredSemanticLabels",
        "text": "Required semantic labels",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "searchableProperties",
        "text": "Searchable properties",
        "type": "required",
        "errorMessage": "Schema does not have any searchable properties",
        "status": "failed"
      },
      {
        "id": "contentIngested",
        "text": "Items have content ingested",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "enabledForInlineResults",
        "text": "Connection configured for inline results",
        "type": "required",
        "status": "manual"
      },
      {
        "id": "itemsHaveActivities",
        "text": "Items have activities recorded",
        "type": "recommended",
        "status": "manual"
      },
      {
        "id": "meaningfulNameAndDescription",
        "text": "Meaningful connection name and description",
        "type": "required",
        "status": "manual"
      },
      {
        "id": "urlToItemResolver",
        "text": "urlToItemResolver configured",
        "type": "recommended",
        "status": "passed"
      }
    ]));
  });

  it(`returns an error when couldn't find any content belonging to the external connection`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs`) {
        return externalConnection;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs/schema`) {
        return schema;
      }
      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/search/query') {
        return {
          "value": [
            {
              "searchTerms": [],
              "hitsContainers": [
                {
                  "hits": [],
                  "total": 0,
                  "moreResultsAvailable": false
                }
              ]
            }
          ]
        };
      }
      throw 'Invalid request';
    });
    const options: any = {
      id: 'msgraphdocs',
      ux: 'copilot',
      output: 'json'
    };
    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith([
      {
        "id": "loadExternalConnection",
        "text": "Load connection",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "loadSchema",
        "text": "Load schema",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "copilotRequiredSemanticLabels",
        "text": "Required semantic labels",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "searchableProperties",
        "text": "Searchable properties",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "contentIngested",
        "text": "Items have content ingested",
        "type": "required",
        "errorMessage": "No items found that belong to the connection",
        "status": "failed"
      },
      {
        "id": "enabledForInlineResults",
        "text": "Connection configured for inline results",
        "type": "required",
        "status": "manual"
      },
      {
        "id": "itemsHaveActivities",
        "text": "Items have activities recorded",
        "type": "recommended",
        "status": "manual"
      },
      {
        "id": "meaningfulNameAndDescription",
        "text": "Meaningful connection name and description",
        "type": "required",
        "status": "manual"
      },
      {
        "id": "urlToItemResolver",
        "text": "urlToItemResolver configured",
        "type": "recommended",
        "status": "passed"
      }
    ]));
  });

  it(`returns an error when couldn't resolve itemId from a search result`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs`) {
        return externalConnection;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs/schema`) {
        return schema;
      }
      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/search/query') {
        return {
          "value": [
            {
              "searchTerms": [],
              "hitsContainers": [
                {
                  "hits": [
                    {
                      "hitId": "AAMkAGE3YTcwZDFiLTgzMmItNGQxYi1hOWZjLWRiZWQzMjJlN2VlMABGAAAAAACkN3ckZNrUSIhtfaVyEdYWBwB0XsiLKGwtTIwX0CYw7PKyAAAAAAEwAAB0XsiLKGwtTIwX0CYw7PKyAAAAABzKAAA=",
                      "contentSource": "msgraphdocs",
                      "rank": 1,
                      "summary": "<c0>Create</c0> <c0>a</c0> <c0>Microsoft</c0> <c0>Graph</c0> <c0>client</c0> <c0>The</c0> <c0>Microsoft</c0> <c0>Graph</c0> <c0>client</c0> <c0>is</c0> <c0>designed</c0> <c0>to</c0> <c0>make</c0> <c0>it</c0> <c0>simple</c0> <ddd/> <c0>make</c0> <c0>calls</c0> <c0>to</c0> <c0>Microsoft</c0> <c0>Graph</c0>. <c0>You</c0> <c0>can</c0> <c0>use</c0> <c0>a</c0> <c0>single</c0> <c0>client</c0> <c0>instance</c0> <c0>for</c0> <c0>the</c0> <c0>lifetime</c0> <c0>of</c0> <ddd/>",
                      "resource": {
                        "@odata.type": "#microsoft.graph.externalConnectors.externalItem",
                        "properties": {
                          "title": "Create a Microsoft Graph client",
                          "url": "https://learn.microsoft.com/graph/sdks/create-client",
                          "description": "Describes how to create a client to use to make calls to Microsoft Graph. Includes how to set up authentication and select a sovereign cloud.",
                          "substrateContentDomainId": "444d0a3470e44d468899fb3dcfcda99c@02b03c8c-a55c-4f23-9285-d5bd8f81979a",
                          "substrateLocationId": "APP:GCS_msgraphdocs@02b03c8c-a55c-4f23-9285-d5bd8f81979a",
                          "documentId": 2012,
                          "immutableEntryId": "AAAAAB2EAxGqZhHNm8gAqgAvxFoNAHReyIsobC1MjBfQJjDs8rIAAAAAILIAAA==",
                          "id": "AAMkAGE3YTcwZDFiLTgzMmItNGQxYi1hOWZjLWRiZWQzMjJlN2VlMABGAAAAAACkN3ckZNrUSIhtfaVyEdYWBwB0XsiLKGwtTIwX0CYw7PKyAAAAAAEwAAB0XsiLKGwtTIwX0CYw7PKyAAAAABzKAAA="
                        }
                      } as any
                    }
                  ],
                  "total": 15,
                  "moreResultsAvailable": true
                }
              ]
            }
          ]
        };
      }
      throw 'Invalid request';
    });
    const options: any = {
      id: 'msgraphdocs',
      ux: 'copilot',
      output: 'json'
    };
    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith([
      {
        "id": "loadExternalConnection",
        "text": "Load connection",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "loadSchema",
        "text": "Load schema",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "copilotRequiredSemanticLabels",
        "text": "Required semantic labels",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "searchableProperties",
        "text": "Searchable properties",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "contentIngested",
        "text": "Items have content ingested",
        "type": "required",
        "errorMessage": "Item does not have substrateContentDomainId property or the property is invalid",
        "status": "failed"
      },
      {
        "id": "enabledForInlineResults",
        "text": "Connection configured for inline results",
        "type": "required",
        "status": "manual"
      },
      {
        "id": "itemsHaveActivities",
        "text": "Items have activities recorded",
        "type": "recommended",
        "status": "manual"
      },
      {
        "id": "meaningfulNameAndDescription",
        "text": "Meaningful connection name and description",
        "type": "required",
        "status": "manual"
      },
      {
        "id": "urlToItemResolver",
        "text": "urlToItemResolver configured",
        "type": "recommended",
        "status": "passed"
      }
    ]));
  });

  it(`returns an error when external item has no content`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs`) {
        return externalConnection;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs/schema`) {
        return schema;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs/items/sdks__create-client`) {
        return {
          "id": "sdks__create-client",
          "acl": [
            {
              "type": "everyone",
              "value": "everyone",
              "accessType": "grant"
            }
          ],
          "properties": {
            "title": "Create a Microsoft Graph client",
            "description": "Describes how to create a client to use to make calls to Microsoft Graph. Includes how to set up authentication and select a sovereign cloud.",
            "url": "https://learn.microsoft.com/graph/sdks/create-client"
          },
          "content": {
            "value": "",
            "type": "html"
          }
        };
      }
      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/search/query') {
        return searchResponse;
      }
      throw 'Invalid request';
    });
    const options: any = {
      id: 'msgraphdocs',
      ux: 'copilot',
      output: 'json'
    };
    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith([
      {
        "id": "loadExternalConnection",
        "text": "Load connection",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "loadSchema",
        "text": "Load schema",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "copilotRequiredSemanticLabels",
        "text": "Required semantic labels",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "searchableProperties",
        "text": "Searchable properties",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "contentIngested",
        "text": "Items have content ingested",
        "type": "required",
        "errorMessage": "Item does not have content or content is empty",
        "status": "failed"
      },
      {
        "id": "enabledForInlineResults",
        "text": "Connection configured for inline results",
        "type": "required",
        "status": "manual"
      },
      {
        "id": "itemsHaveActivities",
        "text": "Items have activities recorded",
        "type": "recommended",
        "status": "manual"
      },
      {
        "id": "meaningfulNameAndDescription",
        "text": "Meaningful connection name and description",
        "type": "required",
        "status": "manual"
      },
      {
        "id": "urlToItemResolver",
        "text": "urlToItemResolver configured",
        "type": "recommended",
        "status": "passed"
      }
    ]));
  });

  it(`returns an error when searching for external items failed`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs`) {
        return externalConnection;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs/schema`) {
        return schema;
      }
      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/search/query') {
        return Promise.reject({
          response: {
            data: {
              "error": {
                "code": "ItemNotFound",
                "message": "An error has occurred.",
                "innerError": {
                  "code": "ResourceNotFound",
                  "source": "Schema",
                  "message": "An error has occurred.",
                  "date": "2023-11-22T18:05:10",
                  "request-id": "a844db09-1ab3-4209-aea8-ee5b8b72949a",
                  "client-request-id": "a844db09-1ab3-4209-aea8-ee5b8b72949a"
                }
              }
            }
          }
        });
      }
      throw 'Invalid request';
    });
    const options: any = {
      id: 'msgraphdocs',
      ux: 'copilot',
      output: 'json'
    };
    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith([
      {
        "id": "loadExternalConnection",
        "text": "Load connection",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "loadSchema",
        "text": "Load schema",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "copilotRequiredSemanticLabels",
        "text": "Required semantic labels",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "searchableProperties",
        "text": "Searchable properties",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "contentIngested",
        "text": "Items have content ingested",
        "type": "required",
        "error": "An error has occurred.",
        "errorMessage": "Error while checking if content is ingested",
        "status": "failed"
      },
      {
        "id": "enabledForInlineResults",
        "text": "Connection configured for inline results",
        "type": "required",
        "status": "manual"
      },
      {
        "id": "itemsHaveActivities",
        "text": "Items have activities recorded",
        "type": "recommended",
        "status": "manual"
      },
      {
        "id": "meaningfulNameAndDescription",
        "text": "Meaningful connection name and description",
        "type": "required",
        "status": "manual"
      },
      {
        "id": "urlToItemResolver",
        "text": "urlToItemResolver configured",
        "type": "recommended",
        "status": "passed"
      }
    ]));
  });

  it(`returns an error when urlToItemResolver is not configured`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs`) {
        return {
          "id": "msgraphdocs",
          "name": "Microsoft Graph documentation",
          "description": "Documentation for Microsoft Graph API which explains what Microsoft Graph is and how to use it.",
          "state": "ready",
          "configuration": {
            "authorizedAppIds": [
              "4f56e2be-b7ae-4b30-9fd0-c42d6bdf8255"
            ]
          },
          "searchSettings": {
            "searchResultTemplates": [
              {
                "id": "msgraphdocs",
                "priority": 1,
                "layout": {
                  "type": "AdaptiveCard",
                  "version": "1.3",
                  "body": [
                    {
                      "type": "ColumnSet",
                      "columns": [
                        {
                          "type": "Column",
                          "width": "auto",
                          "items": [
                            {
                              "type": "Image",
                              "url": "https://raw.githubusercontent.com/waldekmastykarz/img/main/microsoft-graph.png",
                              "altText": "Thumbnail image",
                              "horizontalAlignment": "center",
                              "size": "small"
                            }
                          ],
                          "horizontalAlignment": "center"
                        },
                        {
                          "type": "Column",
                          "width": "stretch",
                          "items": [
                            {
                              "type": "TextBlock",
                              "text": "[${title}](${url})",
                              "weight": "bolder",
                              "color": "accent",
                              "size": "medium",
                              "maxLines": 3
                            },
                            {
                              "type": "TextBlock",
                              "text": "[${url}](${url})",
                              "weight": "bolder",
                              "spacing": "small",
                              "maxLines": 3
                            },
                            {
                              "type": "TextBlock",
                              "text": "${description}",
                              "maxLines": 3,
                              "wrap": true
                            }
                          ],
                          "spacing": "medium"
                        }
                      ]
                    }
                  ],
                  "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                  "$data": {
                    "description": "Marketing team at Contoso.., and looking at the Contoso Marketing documents on the team site. This contains the data from FY20 and will taken over to FY21...Marketing Planning is ongoing for FY20..",
                    "url": "https://contoso.com",
                    "title": "Contoso Solutions"
                  }
                },
                "rules": []
              }
            ]
          },
          "activitySettings": null
        };
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs/schema`) {
        return schema;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs/items/sdks__create-client`) {
        return externalItem;
      }
      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/search/query') {
        return searchResponse;
      }
      throw 'Invalid request';
    });
    const options: any = {
      id: 'msgraphdocs',
      ux: 'copilot',
      output: 'json'
    };
    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith([
      {
        "id": "loadExternalConnection",
        "text": "Load connection",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "loadSchema",
        "text": "Load schema",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "copilotRequiredSemanticLabels",
        "text": "Required semantic labels",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "searchableProperties",
        "text": "Searchable properties",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "contentIngested",
        "text": "Items have content ingested",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "enabledForInlineResults",
        "text": "Connection configured for inline results",
        "type": "required",
        "status": "manual"
      },
      {
        "id": "itemsHaveActivities",
        "text": "Items have activities recorded",
        "type": "recommended",
        "status": "manual"
      },
      {
        "id": "meaningfulNameAndDescription",
        "text": "Meaningful connection name and description",
        "type": "required",
        "status": "manual"
      },
      {
        "id": "urlToItemResolver",
        "text": "urlToItemResolver configured",
        "type": "recommended",
        "errorMessage": "urlToItemResolver is not configured",
        "status": "failed"
      }
    ]));
  });

  it(`returns an error if the schema has no semantic labels`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs`) {
        return externalConnection;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs/schema`) {
        return {
          "baseType": "microsoft.graph.externalItem",
          "properties": [
            {
              "name": "title",
              "type": "string",
              "isSearchable": true,
              "isRetrievable": true,
              "isQueryable": true,
              "labels": [],
              "isRefinable": false,
              "aliases": []
            },
            {
              "name": "iconUrl",
              "type": "string",
              "isSearchable": false,
              "isRetrievable": true,
              "isQueryable": false,
              "labels": [],
              "isRefinable": false,
              "aliases": []
            },
            {
              "name": "url",
              "type": "string",
              "isSearchable": false,
              "isRetrievable": true,
              "isQueryable": false,
              "labels": [],
              "isRefinable": false,
              "aliases": []
            }
          ]
        };
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs/items/sdks__create-client`) {
        return externalItem;
      }
      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/search/query') {
        return searchResponse;
      }
      throw 'Invalid request';
    });
    const options: any = {
      id: 'msgraphdocs',
      ux: 'search',
      output: 'json'
    };
    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith([
      {
        "id": "loadExternalConnection",
        "text": "Load connection",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "loadSchema",
        "text": "Load schema",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "semanticLabels",
        "text": "Semantic labels",
        "type": "recommended",
        "errorMessage": "Schema does not have semantic labels",
        "status": "failed"
      },
      {
        "id": "searchableProperties",
        "text": "Searchable properties",
        "type": "recommended",
        "status": "passed"
      },
      {
        "id": "resultType",
        "text": "Result type",
        "type": "recommended",
        "status": "passed"
      },
      {
        "id": "contentIngested",
        "text": "Items have content ingested",
        "type": "recommended",
        "status": "passed"
      },
      {
        "id": "itemsHaveActivities",
        "text": "Items have activities recorded",
        "type": "recommended",
        "status": "manual"
      },
      {
        "id": "urlToItemResolver",
        "text": "urlToItemResolver configured",
        "type": "recommended",
        "status": "passed"
      }
    ]));
  });

  it(`returns an error if the connection has no result type`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs`) {
        return {
          "id": "msgraphdocs",
          "name": "Microsoft Graph documentation",
          "description": "Documentation for Microsoft Graph API which explains what Microsoft Graph is and how to use it.",
          "state": "ready",
          "configuration": {
            "authorizedAppIds": [
              "4f56e2be-b7ae-4b30-9fd0-c42d6bdf8255"
            ]
          },
          "searchSettings": null,
          "activitySettings": {
            "urlToItemResolvers": [
              {
                "priority": 1,
                "itemId": "auth__{slug}",
                "urlMatchInfo": {
                  "baseUrls": [
                    "https://learn.microsoft.com"
                  ],
                  "urlPattern": "/[^/]+/graph/auth/(?<slug>[^/]+)"
                }
              },
              {
                "priority": 2,
                "itemId": "sdks__{slug}",
                "urlMatchInfo": {
                  "baseUrls": [
                    "https://learn.microsoft.com"
                  ],
                  "urlPattern": "/[^/]+/graph/sdks/(?<slug>[^/]+)"
                }
              },
              {
                "priority": 3,
                "itemId": "{slug}",
                "urlMatchInfo": {
                  "baseUrls": [
                    "https://learn.microsoft.com"
                  ],
                  "urlPattern": "/[^/]+/graph/(?<slug>[^/]+)"
                }
              } as ExternalConnectors.ItemIdResolver
            ]
          }
        };
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs/schema`) {
        return schema;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs/items/sdks__create-client`) {
        return externalItem;
      }
      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/search/query') {
        return searchResponse;
      }
      throw 'Invalid request';
    });
    const options: any = {
      id: 'msgraphdocs',
      ux: 'search',
      output: 'json'
    };
    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith([
      {
        "id": "loadExternalConnection",
        "text": "Load connection",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "loadSchema",
        "text": "Load schema",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "semanticLabels",
        "text": "Semantic labels",
        "type": "recommended",
        "status": "passed"
      },
      {
        "id": "searchableProperties",
        "text": "Searchable properties",
        "type": "recommended",
        "status": "passed"
      },
      {
        "id": "resultType",
        "text": "Result type",
        "type": "recommended",
        "errorMessage": "Connection has no result types",
        "status": "failed"
      },
      {
        "id": "contentIngested",
        "text": "Items have content ingested",
        "type": "recommended",
        "status": "passed"
      },
      {
        "id": "itemsHaveActivities",
        "text": "Items have activities recorded",
        "type": "recommended",
        "status": "manual"
      },
      {
        "id": "urlToItemResolver",
        "text": "urlToItemResolver configured",
        "type": "recommended",
        "status": "passed"
      }
    ]));
  });

  it('checks connection compatibility for copilot in text output mode', async () => {
    sinon.stub(request, 'get').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs`) {
        return externalConnection;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs/schema`) {
        return schema;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs/items/sdks__create-client`) {
        return externalItem;
      }
      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/search/query') {
        return searchResponse;
      }
      throw 'Invalid request';
    });
    const options: any = {
      id: 'msgraphdocs',
      ux: 'copilot',
      output: 'text'
    };
    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.called);
  });

  it('stops when the specified connection is not found in text output mode', async () => {
    sinon.stub(request, 'get').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs`) {
        return Promise.reject({
          response: {
            data: {
              "error": {
                "code": "ItemNotFound",
                "message": "The requested resource does not exist.",
                "innerError": {
                  "code": "ResourceNotFound",
                  "source": "Connection",
                  "message": "'Connection' was not found.",
                  "date": "2023-11-22T18:05:10",
                  "request-id": "a844db09-1ab3-4209-aea8-ee5b8b72949a",
                  "client-request-id": "a844db09-1ab3-4209-aea8-ee5b8b72949a"
                }
              }
            }
          }
        });
      }
      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async () => {
      throw 'Invalid request';
    });
    const options: any = {
      id: 'msgraphdocs',
      ux: 'copilot',
      output: 'text'
    };
    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.called);
  });

  it(`returns an error when urlToItemResolver is not configured in text output mode`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs`) {
        return {
          "id": "msgraphdocs",
          "name": "Microsoft Graph documentation",
          "description": "Documentation for Microsoft Graph API which explains what Microsoft Graph is and how to use it.",
          "state": "ready",
          "configuration": {
            "authorizedAppIds": [
              "4f56e2be-b7ae-4b30-9fd0-c42d6bdf8255"
            ]
          },
          "searchSettings": {
            "searchResultTemplates": [
              {
                "id": "msgraphdocs",
                "priority": 1,
                "layout": {
                  "type": "AdaptiveCard",
                  "version": "1.3",
                  "body": [
                    {
                      "type": "ColumnSet",
                      "columns": [
                        {
                          "type": "Column",
                          "width": "auto",
                          "items": [
                            {
                              "type": "Image",
                              "url": "https://raw.githubusercontent.com/waldekmastykarz/img/main/microsoft-graph.png",
                              "altText": "Thumbnail image",
                              "horizontalAlignment": "center",
                              "size": "small"
                            }
                          ],
                          "horizontalAlignment": "center"
                        },
                        {
                          "type": "Column",
                          "width": "stretch",
                          "items": [
                            {
                              "type": "TextBlock",
                              "text": "[${title}](${url})",
                              "weight": "bolder",
                              "color": "accent",
                              "size": "medium",
                              "maxLines": 3
                            },
                            {
                              "type": "TextBlock",
                              "text": "[${url}](${url})",
                              "weight": "bolder",
                              "spacing": "small",
                              "maxLines": 3
                            },
                            {
                              "type": "TextBlock",
                              "text": "${description}",
                              "maxLines": 3,
                              "wrap": true
                            }
                          ],
                          "spacing": "medium"
                        }
                      ]
                    }
                  ],
                  "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                  "$data": {
                    "description": "Marketing team at Contoso.., and looking at the Contoso Marketing documents on the team site. This contains the data from FY20 and will taken over to FY21...Marketing Planning is ongoing for FY20..",
                    "url": "https://contoso.com",
                    "title": "Contoso Solutions"
                  }
                },
                "rules": []
              }
            ]
          },
          "activitySettings": null
        };
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs/schema`) {
        return schema;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs/items/sdks__create-client`) {
        return externalItem;
      }
      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/search/query') {
        return searchResponse;
      }
      throw 'Invalid request';
    });
    const options: any = {
      id: 'msgraphdocs',
      ux: 'copilot',
      output: 'text'
    };
    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.called);
  });

  it('checks connection compatibility for copilot in csv output mode', async () => {
    sinon.stub(request, 'get').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs`) {
        return externalConnection;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs/schema`) {
        return schema;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs/items/sdks__create-client`) {
        return externalItem;
      }
      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/search/query') {
        return searchResponse;
      }
      throw 'Invalid request';
    });
    const options: any = {
      id: 'msgraphdocs',
      ux: 'copilot',
      output: 'csv'
    };
    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith([
      {
        "id": "loadExternalConnection",
        "text": "Load connection",
        "type": "required",
        "status": "passed",
        "errorMessage": ""
      },
      {
        "id": "loadSchema",
        "text": "Load schema",
        "type": "required",
        "status": "passed",
        "errorMessage": ""
      },
      {
        "id": "copilotRequiredSemanticLabels",
        "text": "Required semantic labels",
        "type": "required",
        "status": "passed",
        "errorMessage": ""
      },
      {
        "id": "searchableProperties",
        "text": "Searchable properties",
        "type": "required",
        "status": "passed",
        "errorMessage": ""
      },
      {
        "id": "contentIngested",
        "text": "Items have content ingested",
        "type": "required",
        "status": "passed",
        "errorMessage": ""
      },
      {
        "id": "enabledForInlineResults",
        "text": "Connection configured for inline results",
        "type": "required",
        "status": "manual",
        "errorMessage": ""
      },
      {
        "id": "itemsHaveActivities",
        "text": "Items have activities recorded",
        "type": "recommended",
        "status": "manual",
        "errorMessage": ""
      },
      {
        "id": "meaningfulNameAndDescription",
        "text": "Meaningful connection name and description",
        "type": "required",
        "status": "manual",
        "errorMessage": ""
      },
      {
        "id": "urlToItemResolver",
        "text": "urlToItemResolver configured",
        "type": "recommended",
        "status": "passed",
        "errorMessage": ""
      }
    ]));
  });

  it('checks connection compatibility for copilot in md output mode', async () => {
    sinon.stub(request, 'get').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs`) {
        return externalConnection;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs/schema`) {
        return schema;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/msgraphdocs/items/sdks__create-client`) {
        return externalItem;
      }
      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/search/query') {
        return searchResponse;
      }
      throw 'Invalid request';
    });
    const options: any = {
      id: 'msgraphdocs',
      ux: 'copilot',
      output: 'md'
    };
    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith([
      {
        "id": "loadExternalConnection",
        "text": "Load connection",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "loadSchema",
        "text": "Load schema",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "copilotRequiredSemanticLabels",
        "text": "Required semantic labels",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "searchableProperties",
        "text": "Searchable properties",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "contentIngested",
        "text": "Items have content ingested",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "enabledForInlineResults",
        "text": "Connection configured for inline results",
        "type": "required",
        "status": "manual"
      },
      {
        "id": "itemsHaveActivities",
        "text": "Items have activities recorded",
        "type": "recommended",
        "status": "manual"
      },
      {
        "id": "meaningfulNameAndDescription",
        "text": "Meaningful connection name and description",
        "type": "required",
        "status": "manual"
      },
      {
        "id": "urlToItemResolver",
        "text": "urlToItemResolver configured",
        "type": "recommended",
        "status": "passed"
      }
    ]));
  });

  it('prints a user-friendly table for md output', async () => {
    const options: any = {
      id: 'msgraphdocs',
      ux: 'copilot',
      output: 'md'
    };
    const output = command.getMdOutput([[
      {
        "id": "loadExternalConnection",
        "text": "Load connection",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "loadSchema",
        "text": "Load schema",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "copilotRequiredSemanticLabels",
        "text": "Required semantic labels",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "searchableProperties",
        "text": "Searchable properties",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "contentIngested",
        "text": "Items have content ingested",
        "type": "required",
        "status": "passed"
      },
      {
        "id": "enabledForInlineResults",
        "text": "Connection configured for inline results",
        "type": "required",
        "status": "manual"
      },
      {
        "id": "itemsHaveActivities",
        "text": "Items have activities recorded",
        "type": "recommended",
        "status": "manual"
      },
      {
        "id": "meaningfulNameAndDescription",
        "text": "Meaningful connection name and description",
        "type": "required",
        "status": "manual"
      },
      {
        "id": "urlToItemResolver",
        "text": "urlToItemResolver configured",
        "type": "recommended",
        "status": "passed"
      }
    ]], command, options);
    assert(output.indexOf('Check|Type|Status|Error message') > -1);
  });

  it('fails validation if an invalid ux is specified', async () => {
    const actual = await command.validate({
      options: {
        id: 'msgraphdocs',
        ux: 'invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, false);
  });

  it('passes validation for ux copilot', async () => {
    const actual = await command.validate({
      options: {
        id: 'msgraphdocs',
        ux: 'copilot'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation for ux search', async () => {
    const actual = await command.validate({
      options: {
        id: 'msgraphdocs',
        ux: 'search'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation for ux all', async () => {
    const actual = await command.validate({
      options: {
        id: 'msgraphdocs',
        ux: 'all'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when no ux specified', async () => {
    const actual = await command.validate({
      options: {
        id: 'msgraphdocs'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('supports specifying ux', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--ux') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});
