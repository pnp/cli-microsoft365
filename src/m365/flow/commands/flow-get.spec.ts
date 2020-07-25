import commands from '../commands';
import Command, { CommandOption, CommandError } from '../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../appInsights';
import auth from '../../../Auth';
const command: Command = require('./flow-get');
import * as assert from 'assert';
import request from '../../../request';
import Utils from '../../../Utils';

describe(commands.FLOW_GET, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
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
    assert.strictEqual(command.name.startsWith(commands.FLOW_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves information about the specified flow (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/3989cb59-ce1a-4a5c-bb78-257c5c39381d?api-version=2016-11-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            "name": "3989cb59-ce1a-4a5c-bb78-257c5c39381d",
            "id": "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/3989cb59-ce1a-4a5c-bb78-257c5c39381d",
            "type": "Microsoft.ProcessSimple/environments/flows",
            "properties": {
              "apiId": "/providers/Microsoft.PowerApps/apis/shared_logicflows",
              "displayName": "Get a daily digest of the top CNN news",
              "userType": "Owner",
              "definition": {
                "$schema": "https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#",
                "contentVersion": "1.0.0.0",
                "parameters": {
                  "$connections": {
                    "defaultValue": {},
                    "type": "Object"
                  },
                  "$authentication": {
                    "defaultValue": {},
                    "type": "SecureObject"
                  }
                },
                "triggers": {
                  "Every_day": {
                    "recurrence": {
                      "frequency": "Day",
                      "interval": 1
                    },
                    "type": "Recurrence"
                  }
                },
                "actions": {
                  "Check_if_there_were_any_posts_this_week": {
                    "actions": {
                      "Compose_the_links_for_each_blog_post": {
                        "foreach": "@take(body('Filter_array'),15)",
                        "actions": {
                          "Compose": {
                            "runAfter": {},
                            "type": "Compose",
                            "inputs": "<tr><td><h3><a href=\"@{item()?['primaryLink']}\">@{item()['title']}</a></h3></td></tr><tr><td style=\"color: #777777;\">Posted at @{formatDateTime(item()?['publishDate'], 't')} GMT</td></tr><tr><td>@{item()?['summary']}</td></tr>"
                          }
                        },
                        "runAfter": {},
                        "type": "Foreach"
                      },
                      "Send_an_email": {
                        "runAfter": {
                          "Compose_the_links_for_each_blog_post": [
                            "Succeeded"
                          ]
                        },
                        "metadata": {
                          "flowSystemMetadata": {
                            "swaggerOperationId": "SendEmailNotification"
                          }
                        },
                        "type": "ApiConnection",
                        "inputs": {
                          "body": {
                            "notificationBody": "<h2>Check out the latest news from CNN:</h2><table>@{join(outputs('Compose'),'')}</table>",
                            "notificationSubject": "Daily digest from CNN top news"
                          },
                          "host": {
                            "api": {
                              "runtimeUrl": "https://europe-001.azure-apim.net/apim/flowpush"
                            },
                            "connection": {
                              "name": "@parameters('$connections')['shared_flowpush']['connectionId']"
                            }
                          },
                          "method": "post",
                          "path": "/sendEmailNotification",
                          "authentication": "@parameters('$authentication')"
                        }
                      }
                    },
                    "runAfter": {
                      "Filter_array": [
                        "Succeeded"
                      ]
                    },
                    "expression": "@greater(length(body('Filter_array')), 0)",
                    "type": "If"
                  },
                  "Filter_array": {
                    "runAfter": {
                      "List_all_RSS_feed_items": [
                        "Succeeded"
                      ]
                    },
                    "type": "Query",
                    "inputs": {
                      "from": "@body('List_all_RSS_feed_items')",
                      "where": "@greater(item()?['publishDate'], adddays(utcnow(),-2))"
                    }
                  },
                  "List_all_RSS_feed_items": {
                    "runAfter": {},
                    "metadata": {
                      "flowSystemMetadata": {
                        "swaggerOperationId": "ListFeedItems"
                      }
                    },
                    "type": "ApiConnection",
                    "inputs": {
                      "host": {
                        "api": {
                          "runtimeUrl": "https://europe-001.azure-apim.net/apim/rss"
                        },
                        "connection": {
                          "name": "@parameters('$connections')['shared_rss']['connectionId']"
                        }
                      },
                      "method": "get",
                      "path": "/ListFeedItems",
                      "queries": {
                        "feedUrl": "http://rss.cnn.com/rss/cnn_topstories.rss"
                      },
                      "authentication": "@parameters('$authentication')"
                    }
                  }
                },
                "outputs": {},
                "description": "Each day, get an email with a list of all of the top CNN posts from the last day."
              },
              "state": "Started",
              "connectionReferences": {
                "shared_rss": {
                  "connectionName": "shared-rss-6636bfd3-0d29-4842-b835-c5910b6310f6",
                  "apiDefinition": {
                    "name": "shared_rss",
                    "id": "/providers/Microsoft.PowerApps/apis/shared_rss",
                    "type": "/providers/Microsoft.PowerApps/apis",
                    "properties": {
                      "displayName": "RSS",
                      "iconUri": "https://az818438.vo.msecnd.net/icons/rss.png",
                      "purpose": "NotSpecified",
                      "connectionParameters": {},
                      "scopes": {
                        "will": [],
                        "wont": []
                      },
                      "runtimeUrls": [
                        "https://europe-001.azure-apim.net/apim/rss"
                      ],
                      "primaryRuntimeUrl": "https://europe-001.azure-apim.net/apim/rss",
                      "metadata": {
                        "source": "marketplace",
                        "brandColor": "#ff9900"
                      },
                      "capabilities": [
                        "actions"
                      ],
                      "tier": "Standard",
                      "description": "RSS is a popular web syndication format used to publish frequently updated content – like blog entries and news headlines.  Many content publishers provide an RSS feed to allow users to subscribe to it.  Use the RSS connector to retrieve feed information and trigger flows when new items are published in an RSS feed.",
                      "createdTime": "2016-09-30T04:13:14.2871915Z",
                      "changedTime": "2018-01-17T20:20:37.905252Z"
                    }
                  },
                  "source": "Embedded",
                  "id": "/providers/Microsoft.PowerApps/apis/shared_rss",
                  "displayName": "RSS",
                  "iconUri": "https://az818438.vo.msecnd.net/icons/rss.png",
                  "brandColor": "#ff9900",
                  "swagger": {
                    "swagger": "2.0",
                    "info": {
                      "version": "1.0",
                      "title": "RSS",
                      "description": "RSS is a popular web syndication format used to publish frequently updated content – like blog entries and news headlines.  Many content publishers provide an RSS feed to allow users to subscribe to it.  Use the RSS connector to retrieve feed information and trigger flows when new items are published in an RSS feed.",
                      "x-ms-api-annotation": {
                        "status": "Production"
                      }
                    },
                    "host": "europe-001.azure-apim.net",
                    "basePath": "/apim/rss",
                    "schemes": [
                      "https"
                    ],
                    "paths": {
                      "/{connectionId}/OnNewFeed": {
                        "get": {
                          "tags": [
                            "Rss"
                          ],
                          "summary": "When a feed item is published",
                          "description": "This operation triggers a workflow when a new item is published in an RSS feed.",
                          "operationId": "OnNewFeed",
                          "consumes": [],
                          "produces": [
                            "application/json",
                            "text/json",
                            "application/xml",
                            "text/xml"
                          ],
                          "parameters": [
                            {
                              "name": "connectionId",
                              "in": "path",
                              "required": true,
                              "type": "string",
                              "x-ms-visibility": "internal"
                            },
                            {
                              "name": "feedUrl",
                              "in": "query",
                              "description": "The RSS feed URL (Example: http://rss.cnn.com/rss/cnn_topstories.rss).",
                              "required": true,
                              "x-ms-summary": "The RSS feed URL",
                              "x-ms-url-encoding": "double",
                              "type": "string"
                            }
                          ],
                          "responses": {
                            "200": {
                              "description": "OK",
                              "schema": {
                                "$ref": "#/definitions/TriggerBatchResponse[FeedItem]"
                              }
                            },
                            "202": {
                              "description": "Accepted"
                            },
                            "400": {
                              "description": "Bad Request"
                            },
                            "401": {
                              "description": "Unauthorized"
                            },
                            "403": {
                              "description": "Forbidden"
                            },
                            "404": {
                              "description": "Not Found"
                            },
                            "500": {
                              "description": "Internal Server Error. Unknown error occurred"
                            },
                            "default": {
                              "description": "Operation Failed."
                            }
                          },
                          "deprecated": false,
                          "x-ms-visibility": "important",
                          "x-ms-trigger": "batch",
                          "x-ms-trigger-hint": "To see it work now, publish an item to the RSS feed."
                        }
                      },
                      "/{connectionId}/ListFeedItems": {
                        "get": {
                          "tags": [
                            "Rss"
                          ],
                          "summary": "List all RSS feed items",
                          "description": "This operation retrieves all items from an RSS feed.",
                          "operationId": "ListFeedItems",
                          "consumes": [],
                          "produces": [
                            "application/json",
                            "text/json",
                            "application/xml",
                            "text/xml"
                          ],
                          "parameters": [
                            {
                              "name": "connectionId",
                              "in": "path",
                              "required": true,
                              "type": "string",
                              "x-ms-visibility": "internal"
                            },
                            {
                              "name": "feedUrl",
                              "in": "query",
                              "description": "The RSS feed URL (Example: http://rss.cnn.com/rss/cnn_topstories.rss).",
                              "required": true,
                              "x-ms-summary": "The RSS feed URL",
                              "x-ms-url-encoding": "double",
                              "type": "string"
                            }
                          ],
                          "responses": {
                            "200": {
                              "description": "OK",
                              "schema": {
                                "type": "array",
                                "items": {
                                  "$ref": "#/definitions/FeedItem"
                                }
                              }
                            },
                            "202": {
                              "description": "Accepted"
                            },
                            "400": {
                              "description": "Bad Request"
                            },
                            "401": {
                              "description": "Unauthorized"
                            },
                            "403": {
                              "description": "Forbidden"
                            },
                            "404": {
                              "description": "Not Found"
                            },
                            "500": {
                              "description": "Internal Server Error. Unknown error occurred"
                            },
                            "default": {
                              "description": "Operation Failed."
                            }
                          },
                          "deprecated": false,
                          "x-ms-visibility": "important"
                        }
                      }
                    },
                    "definitions": {
                      "TriggerBatchResponse[FeedItem]": {
                        "description": "Represents a wrapper object for batch trigger response",
                        "type": "object",
                        "properties": {
                          "value": {
                            "description": "A list of the response objects",
                            "type": "array",
                            "items": {
                              "$ref": "#/definitions/FeedItem"
                            }
                          }
                        }
                      },
                      "FeedItem": {
                        "description": "Represents an RSS feed item.",
                        "required": [
                          "id",
                          "title"
                        ],
                        "type": "object",
                        "properties": {
                          "id": {
                            "description": "Feed ID",
                            "type": "string",
                            "x-ms-summary": "Feed ID"
                          },
                          "title": {
                            "description": "Feed title",
                            "type": "string",
                            "x-ms-summary": "Feed title"
                          },
                          "primaryLink": {
                            "description": "Primary feed link",
                            "type": "string",
                            "x-ms-summary": "Primary feed link"
                          },
                          "links": {
                            "description": "Feed links",
                            "type": "array",
                            "items": {
                              "type": "string"
                            },
                            "x-ms-summary": "Feed links",
                            "x-ms-visibility": "advanced"
                          },
                          "updatedOn": {
                            "description": "Feed updated on",
                            "type": "string",
                            "x-ms-summary": "Feed updated on"
                          },
                          "publishDate": {
                            "description": "Feed published date",
                            "type": "string",
                            "x-ms-summary": "Feed published on"
                          },
                          "summary": {
                            "description": "Feed item summary",
                            "type": "string",
                            "x-ms-summary": "Feed summary"
                          },
                          "copyright": {
                            "description": "Copyright information",
                            "type": "string",
                            "x-ms-summary": "Feed copyright information"
                          },
                          "categories": {
                            "description": "Feed item categories",
                            "type": "array",
                            "items": {
                              "type": "string"
                            },
                            "x-ms-summary": "Feed categories"
                          }
                        }
                      }
                    }
                  },
                  "tier": "NotSpecified"
                },
                "shared_flowpush": {
                  "connectionName": "shared-flowpush-295e4b80-1a4e-42ec-aa5b-8d72e7c1eb3f",
                  "apiDefinition": {
                    "name": "shared_flowpush",
                    "id": "/providers/Microsoft.PowerApps/apis/shared_flowpush",
                    "type": "/providers/Microsoft.PowerApps/apis",
                    "properties": {
                      "displayName": "Notifications",
                      "iconUri": "https://psux.azureedge.net/Content/Images/Connectors/FlowNotification.svg",
                      "purpose": "NotSpecified",
                      "connectionParameters": {},
                      "runtimeUrls": [
                        "https://europe-001.azure-apim.net/apim/flowpush"
                      ],
                      "primaryRuntimeUrl": "https://europe-001.azure-apim.net/apim/flowpush",
                      "metadata": {
                        "source": "marketplace",
                        "brandColor": "#FF3B30",
                        "connectionLimits": {
                          "*": 1
                        }
                      },
                      "capabilities": [
                        "actions"
                      ],
                      "tier": "NotSpecified",
                      "description": "The notification service enables notifications created by a flow to go to your email account or Microsoft Flow mobile app.",
                      "createdTime": "2016-10-22T06:43:26.1572419Z",
                      "changedTime": "2018-01-25T00:34:52.048009Z"
                    }
                  },
                  "source": "Embedded",
                  "id": "/providers/Microsoft.PowerApps/apis/shared_flowpush",
                  "displayName": "Notifications",
                  "iconUri": "https://psux.azureedge.net/Content/Images/Connectors/FlowNotification.svg",
                  "brandColor": "#FF3B30",
                  "swagger": {
                    "swagger": "2.0",
                    "info": {
                      "version": "1.0",
                      "title": "Notifications",
                      "description": "The notification service enables notifications created by a flow to go to your email account or Microsoft Flow mobile app.",
                      "contact": {
                        "name": "Samuel L. Banina",
                        "email": "saban@microsoft.com"
                      }
                    },
                    "host": "europe-001.azure-apim.net",
                    "basePath": "/apim/flowpush",
                    "schemes": [
                      "https"
                    ],
                    "consumes": [
                      "application/json"
                    ],
                    "produces": [
                      "application/json"
                    ],
                    "x-ms-capabilities": {
                      "buttons": {
                        "flowIosApp": {},
                        "flowAndroidApp": {}
                      }
                    },
                    "definitions": {
                      "NotificationDefinition": {
                        "type": "object",
                        "required": [
                          "notificationText"
                        ],
                        "properties": {
                          "notificationText": {
                            "description": "Create a notification message",
                            "x-ms-summary": "Text",
                            "type": "string"
                          },
                          "notificationLink": {
                            "description": "Custom notification link",
                            "type": "object",
                            "properties": {
                              "uri": {
                                "description": "Include a link in the notification",
                                "x-ms-summary": "Link",
                                "type": "string"
                              },
                              "label": {
                                "description": "The display name for the link",
                                "x-ms-summary": "Link label",
                                "type": "string"
                              }
                            }
                          }
                        }
                      },
                      "NotificationEmailDefinition": {
                        "type": "object",
                        "required": [
                          "notificationSubject",
                          "notificationBody"
                        ],
                        "properties": {
                          "notificationSubject": {
                            "description": "Notification email subject",
                            "x-ms-summary": "Subject",
                            "type": "string"
                          },
                          "notificationBody": {
                            "description": "Notification email body",
                            "x-ms-summary": "Body",
                            "type": "string"
                          }
                        }
                      }
                    },
                    "paths": {
                      "/{connectionId}/sendNotification": {
                        "post": {
                          "description": "Sends a push notification to your Microsoft Flow mobile app.",
                          "summary": "Send me a mobile notification",
                          "operationId": "SendNotification",
                          "x-ms-visibility": "important",
                          "consumes": [
                            "application/json"
                          ],
                          "produces": [
                            "application/json"
                          ],
                          "parameters": [
                            {
                              "name": "connectionId",
                              "in": "path",
                              "required": true,
                              "type": "string",
                              "x-ms-visibility": "internal"
                            },
                            {
                              "name": "NotificationDefinition",
                              "x-ms-summary": "The push notification",
                              "in": "body",
                              "description": "Push notification inputs",
                              "required": true,
                              "schema": {
                                "$ref": "#/definitions/NotificationDefinition"
                              }
                            }
                          ],
                          "responses": {
                            "200": {
                              "description": "OK"
                            },
                            "400": {
                              "description": "Bad Request"
                            },
                            "500": {
                              "description": "Internal Server Error"
                            },
                            "default": {
                              "description": "Operation Failed."
                            }
                          }
                        }
                      },
                      "/{connectionId}/sendEmailNotification": {
                        "post": {
                          "description": "Sends an email notification to the account you signed in to Microsoft Flow with.",
                          "summary": "Send me an email notification",
                          "operationId": "SendEmailNotification",
                          "x-ms-visibility": "important",
                          "consumes": [
                            "application/json"
                          ],
                          "produces": [
                            "application/json"
                          ],
                          "parameters": [
                            {
                              "name": "connectionId",
                              "in": "path",
                              "required": true,
                              "type": "string",
                              "x-ms-visibility": "internal"
                            },
                            {
                              "name": "NotificationEmailDefinition",
                              "x-ms-summary": "The email notification",
                              "in": "body",
                              "description": "Email notification inputs",
                              "required": true,
                              "schema": {
                                "$ref": "#/definitions/NotificationEmailDefinition"
                              }
                            }
                          ],
                          "responses": {
                            "200": {
                              "description": "OK"
                            },
                            "400": {
                              "description": "Bad Request"
                            },
                            "500": {
                              "description": "Internal Server Error"
                            },
                            "default": {
                              "description": "Operation Failed."
                            }
                          }
                        }
                      }
                    }
                  },
                  "tier": "NotSpecified"
                }
              },
              "createdTime": "2018-03-23T17:59:35.4407282Z",
              "lastModifiedTime": "2018-03-23T17:59:37.1164508Z",
              "templateName": "a04de6ce52984b3db0b907f588994bc8",
              "environment": {
                "name": "Default-d87a7535-dd31-4437-bfe1-95340acd55c5",
                "type": "Microsoft.ProcessSimple/environments",
                "id": "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5"
              },
              "definitionSummary": {
                "triggers": [
                  {
                    "type": "Recurrence"
                  }
                ],
                "actions": [
                  {
                    "type": "If"
                  },
                  {
                    "type": "Query"
                  },
                  {
                    "type": "ApiConnection",
                    "swaggerOperationId": "ListFeedItems",
                    "metadata": {
                      "flowSystemMetadata": {
                        "swaggerOperationId": "ListFeedItems"
                      }
                    },
                    "api": {
                      "name": "shared_rss",
                      "id": "/providers/Microsoft.PowerApps/apis/shared_rss",
                      "type": "/providers/Microsoft.PowerApps/apis"
                    }
                  },
                  {
                    "type": "Foreach"
                  },
                  {
                    "type": "ApiConnection",
                    "swaggerOperationId": "SendEmailNotification",
                    "metadata": {
                      "flowSystemMetadata": {
                        "swaggerOperationId": "SendEmailNotification"
                      }
                    },
                    "api": {
                      "name": "shared_flowpush",
                      "id": "/providers/Microsoft.PowerApps/apis/shared_flowpush",
                      "type": "/providers/Microsoft.PowerApps/apis"
                    }
                  },
                  {
                    "type": "Compose"
                  }
                ],
                "description": "Each day, get an email with a list of all of the top CNN posts from the last day."
              },
              "creator": {
                "tenantId": "d87a7535-dd31-4437-bfe1-95340acd55c5",
                "objectId": "da8f7aea-cf43-497f-ad62-c2feae89a194",
                "userId": "da8f7aea-cf43-497f-ad62-c2feae89a194",
                "userType": "ActiveDirectory"
              },
              "flowTriggerUri": "https://management.azure.com:443/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/3989cb59-ce1a-4a5c-bb78-257c5c39381d/triggers/Every_day/run?api-version=2016-11-01",
              "installationStatus": "Installed",
              "provisioningMethod": "FromDefinition",
              "flowFailureAlertSubscribed": true,
              "referencedResources": []
            }
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, name: '3989cb59-ce1a-4a5c-bb78-257c5c39381d', environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          name: '3989cb59-ce1a-4a5c-bb78-257c5c39381d',
          displayName: 'Get a daily digest of the top CNN news',
          description: 'Each day, get an email with a list of all of the top CNN posts from the last day.',
          triggers: 'Every_day',
          actions: 'Check_if_there_were_any_posts_this_week, Filter_array, List_all_RSS_feed_items'
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves information about the specified flow', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/3989cb59-ce1a-4a5c-bb78-257c5c39381d?api-version=2016-11-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            "name": "3989cb59-ce1a-4a5c-bb78-257c5c39381d",
            "id": "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/3989cb59-ce1a-4a5c-bb78-257c5c39381d",
            "type": "Microsoft.ProcessSimple/environments/flows",
            "properties": {
              "apiId": "/providers/Microsoft.PowerApps/apis/shared_logicflows",
              "displayName": "Get a daily digest of the top CNN news",
              "userType": "Owner",
              "definition": {
                "$schema": "https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#",
                "contentVersion": "1.0.0.0",
                "parameters": {
                  "$connections": {
                    "defaultValue": {},
                    "type": "Object"
                  },
                  "$authentication": {
                    "defaultValue": {},
                    "type": "SecureObject"
                  }
                },
                "triggers": {
                  "Every_day": {
                    "recurrence": {
                      "frequency": "Day",
                      "interval": 1
                    },
                    "type": "Recurrence"
                  }
                },
                "actions": {
                  "Check_if_there_were_any_posts_this_week": {
                    "actions": {
                      "Compose_the_links_for_each_blog_post": {
                        "foreach": "@take(body('Filter_array'),15)",
                        "actions": {
                          "Compose": {
                            "runAfter": {},
                            "type": "Compose",
                            "inputs": "<tr><td><h3><a href=\"@{item()?['primaryLink']}\">@{item()['title']}</a></h3></td></tr><tr><td style=\"color: #777777;\">Posted at @{formatDateTime(item()?['publishDate'], 't')} GMT</td></tr><tr><td>@{item()?['summary']}</td></tr>"
                          }
                        },
                        "runAfter": {},
                        "type": "Foreach"
                      },
                      "Send_an_email": {
                        "runAfter": {
                          "Compose_the_links_for_each_blog_post": [
                            "Succeeded"
                          ]
                        },
                        "metadata": {
                          "flowSystemMetadata": {
                            "swaggerOperationId": "SendEmailNotification"
                          }
                        },
                        "type": "ApiConnection",
                        "inputs": {
                          "body": {
                            "notificationBody": "<h2>Check out the latest news from CNN:</h2><table>@{join(outputs('Compose'),'')}</table>",
                            "notificationSubject": "Daily digest from CNN top news"
                          },
                          "host": {
                            "api": {
                              "runtimeUrl": "https://europe-001.azure-apim.net/apim/flowpush"
                            },
                            "connection": {
                              "name": "@parameters('$connections')['shared_flowpush']['connectionId']"
                            }
                          },
                          "method": "post",
                          "path": "/sendEmailNotification",
                          "authentication": "@parameters('$authentication')"
                        }
                      }
                    },
                    "runAfter": {
                      "Filter_array": [
                        "Succeeded"
                      ]
                    },
                    "expression": "@greater(length(body('Filter_array')), 0)",
                    "type": "If"
                  },
                  "Filter_array": {
                    "runAfter": {
                      "List_all_RSS_feed_items": [
                        "Succeeded"
                      ]
                    },
                    "type": "Query",
                    "inputs": {
                      "from": "@body('List_all_RSS_feed_items')",
                      "where": "@greater(item()?['publishDate'], adddays(utcnow(),-2))"
                    }
                  },
                  "List_all_RSS_feed_items": {
                    "runAfter": {},
                    "metadata": {
                      "flowSystemMetadata": {
                        "swaggerOperationId": "ListFeedItems"
                      }
                    },
                    "type": "ApiConnection",
                    "inputs": {
                      "host": {
                        "api": {
                          "runtimeUrl": "https://europe-001.azure-apim.net/apim/rss"
                        },
                        "connection": {
                          "name": "@parameters('$connections')['shared_rss']['connectionId']"
                        }
                      },
                      "method": "get",
                      "path": "/ListFeedItems",
                      "queries": {
                        "feedUrl": "http://rss.cnn.com/rss/cnn_topstories.rss"
                      },
                      "authentication": "@parameters('$authentication')"
                    }
                  }
                },
                "outputs": {},
                "description": "Each day, get an email with a list of all of the top CNN posts from the last day."
              },
              "state": "Started",
              "connectionReferences": {
                "shared_rss": {
                  "connectionName": "shared-rss-6636bfd3-0d29-4842-b835-c5910b6310f6",
                  "apiDefinition": {
                    "name": "shared_rss",
                    "id": "/providers/Microsoft.PowerApps/apis/shared_rss",
                    "type": "/providers/Microsoft.PowerApps/apis",
                    "properties": {
                      "displayName": "RSS",
                      "iconUri": "https://az818438.vo.msecnd.net/icons/rss.png",
                      "purpose": "NotSpecified",
                      "connectionParameters": {},
                      "scopes": {
                        "will": [],
                        "wont": []
                      },
                      "runtimeUrls": [
                        "https://europe-001.azure-apim.net/apim/rss"
                      ],
                      "primaryRuntimeUrl": "https://europe-001.azure-apim.net/apim/rss",
                      "metadata": {
                        "source": "marketplace",
                        "brandColor": "#ff9900"
                      },
                      "capabilities": [
                        "actions"
                      ],
                      "tier": "Standard",
                      "description": "RSS is a popular web syndication format used to publish frequently updated content – like blog entries and news headlines.  Many content publishers provide an RSS feed to allow users to subscribe to it.  Use the RSS connector to retrieve feed information and trigger flows when new items are published in an RSS feed.",
                      "createdTime": "2016-09-30T04:13:14.2871915Z",
                      "changedTime": "2018-01-17T20:20:37.905252Z"
                    }
                  },
                  "source": "Embedded",
                  "id": "/providers/Microsoft.PowerApps/apis/shared_rss",
                  "displayName": "RSS",
                  "iconUri": "https://az818438.vo.msecnd.net/icons/rss.png",
                  "brandColor": "#ff9900",
                  "swagger": {
                    "swagger": "2.0",
                    "info": {
                      "version": "1.0",
                      "title": "RSS",
                      "description": "RSS is a popular web syndication format used to publish frequently updated content – like blog entries and news headlines.  Many content publishers provide an RSS feed to allow users to subscribe to it.  Use the RSS connector to retrieve feed information and trigger flows when new items are published in an RSS feed.",
                      "x-ms-api-annotation": {
                        "status": "Production"
                      }
                    },
                    "host": "europe-001.azure-apim.net",
                    "basePath": "/apim/rss",
                    "schemes": [
                      "https"
                    ],
                    "paths": {
                      "/{connectionId}/OnNewFeed": {
                        "get": {
                          "tags": [
                            "Rss"
                          ],
                          "summary": "When a feed item is published",
                          "description": "This operation triggers a workflow when a new item is published in an RSS feed.",
                          "operationId": "OnNewFeed",
                          "consumes": [],
                          "produces": [
                            "application/json",
                            "text/json",
                            "application/xml",
                            "text/xml"
                          ],
                          "parameters": [
                            {
                              "name": "connectionId",
                              "in": "path",
                              "required": true,
                              "type": "string",
                              "x-ms-visibility": "internal"
                            },
                            {
                              "name": "feedUrl",
                              "in": "query",
                              "description": "The RSS feed URL (Example: http://rss.cnn.com/rss/cnn_topstories.rss).",
                              "required": true,
                              "x-ms-summary": "The RSS feed URL",
                              "x-ms-url-encoding": "double",
                              "type": "string"
                            }
                          ],
                          "responses": {
                            "200": {
                              "description": "OK",
                              "schema": {
                                "$ref": "#/definitions/TriggerBatchResponse[FeedItem]"
                              }
                            },
                            "202": {
                              "description": "Accepted"
                            },
                            "400": {
                              "description": "Bad Request"
                            },
                            "401": {
                              "description": "Unauthorized"
                            },
                            "403": {
                              "description": "Forbidden"
                            },
                            "404": {
                              "description": "Not Found"
                            },
                            "500": {
                              "description": "Internal Server Error. Unknown error occurred"
                            },
                            "default": {
                              "description": "Operation Failed."
                            }
                          },
                          "deprecated": false,
                          "x-ms-visibility": "important",
                          "x-ms-trigger": "batch",
                          "x-ms-trigger-hint": "To see it work now, publish an item to the RSS feed."
                        }
                      },
                      "/{connectionId}/ListFeedItems": {
                        "get": {
                          "tags": [
                            "Rss"
                          ],
                          "summary": "List all RSS feed items",
                          "description": "This operation retrieves all items from an RSS feed.",
                          "operationId": "ListFeedItems",
                          "consumes": [],
                          "produces": [
                            "application/json",
                            "text/json",
                            "application/xml",
                            "text/xml"
                          ],
                          "parameters": [
                            {
                              "name": "connectionId",
                              "in": "path",
                              "required": true,
                              "type": "string",
                              "x-ms-visibility": "internal"
                            },
                            {
                              "name": "feedUrl",
                              "in": "query",
                              "description": "The RSS feed URL (Example: http://rss.cnn.com/rss/cnn_topstories.rss).",
                              "required": true,
                              "x-ms-summary": "The RSS feed URL",
                              "x-ms-url-encoding": "double",
                              "type": "string"
                            }
                          ],
                          "responses": {
                            "200": {
                              "description": "OK",
                              "schema": {
                                "type": "array",
                                "items": {
                                  "$ref": "#/definitions/FeedItem"
                                }
                              }
                            },
                            "202": {
                              "description": "Accepted"
                            },
                            "400": {
                              "description": "Bad Request"
                            },
                            "401": {
                              "description": "Unauthorized"
                            },
                            "403": {
                              "description": "Forbidden"
                            },
                            "404": {
                              "description": "Not Found"
                            },
                            "500": {
                              "description": "Internal Server Error. Unknown error occurred"
                            },
                            "default": {
                              "description": "Operation Failed."
                            }
                          },
                          "deprecated": false,
                          "x-ms-visibility": "important"
                        }
                      }
                    },
                    "definitions": {
                      "TriggerBatchResponse[FeedItem]": {
                        "description": "Represents a wrapper object for batch trigger response",
                        "type": "object",
                        "properties": {
                          "value": {
                            "description": "A list of the response objects",
                            "type": "array",
                            "items": {
                              "$ref": "#/definitions/FeedItem"
                            }
                          }
                        }
                      },
                      "FeedItem": {
                        "description": "Represents an RSS feed item.",
                        "required": [
                          "id",
                          "title"
                        ],
                        "type": "object",
                        "properties": {
                          "id": {
                            "description": "Feed ID",
                            "type": "string",
                            "x-ms-summary": "Feed ID"
                          },
                          "title": {
                            "description": "Feed title",
                            "type": "string",
                            "x-ms-summary": "Feed title"
                          },
                          "primaryLink": {
                            "description": "Primary feed link",
                            "type": "string",
                            "x-ms-summary": "Primary feed link"
                          },
                          "links": {
                            "description": "Feed links",
                            "type": "array",
                            "items": {
                              "type": "string"
                            },
                            "x-ms-summary": "Feed links",
                            "x-ms-visibility": "advanced"
                          },
                          "updatedOn": {
                            "description": "Feed updated on",
                            "type": "string",
                            "x-ms-summary": "Feed updated on"
                          },
                          "publishDate": {
                            "description": "Feed published date",
                            "type": "string",
                            "x-ms-summary": "Feed published on"
                          },
                          "summary": {
                            "description": "Feed item summary",
                            "type": "string",
                            "x-ms-summary": "Feed summary"
                          },
                          "copyright": {
                            "description": "Copyright information",
                            "type": "string",
                            "x-ms-summary": "Feed copyright information"
                          },
                          "categories": {
                            "description": "Feed item categories",
                            "type": "array",
                            "items": {
                              "type": "string"
                            },
                            "x-ms-summary": "Feed categories"
                          }
                        }
                      }
                    }
                  },
                  "tier": "NotSpecified"
                },
                "shared_flowpush": {
                  "connectionName": "shared-flowpush-295e4b80-1a4e-42ec-aa5b-8d72e7c1eb3f",
                  "apiDefinition": {
                    "name": "shared_flowpush",
                    "id": "/providers/Microsoft.PowerApps/apis/shared_flowpush",
                    "type": "/providers/Microsoft.PowerApps/apis",
                    "properties": {
                      "displayName": "Notifications",
                      "iconUri": "https://psux.azureedge.net/Content/Images/Connectors/FlowNotification.svg",
                      "purpose": "NotSpecified",
                      "connectionParameters": {},
                      "runtimeUrls": [
                        "https://europe-001.azure-apim.net/apim/flowpush"
                      ],
                      "primaryRuntimeUrl": "https://europe-001.azure-apim.net/apim/flowpush",
                      "metadata": {
                        "source": "marketplace",
                        "brandColor": "#FF3B30",
                        "connectionLimits": {
                          "*": 1
                        }
                      },
                      "capabilities": [
                        "actions"
                      ],
                      "tier": "NotSpecified",
                      "description": "The notification service enables notifications created by a flow to go to your email account or Microsoft Flow mobile app.",
                      "createdTime": "2016-10-22T06:43:26.1572419Z",
                      "changedTime": "2018-01-25T00:34:52.048009Z"
                    }
                  },
                  "source": "Embedded",
                  "id": "/providers/Microsoft.PowerApps/apis/shared_flowpush",
                  "displayName": "Notifications",
                  "iconUri": "https://psux.azureedge.net/Content/Images/Connectors/FlowNotification.svg",
                  "brandColor": "#FF3B30",
                  "swagger": {
                    "swagger": "2.0",
                    "info": {
                      "version": "1.0",
                      "title": "Notifications",
                      "description": "The notification service enables notifications created by a flow to go to your email account or Microsoft Flow mobile app.",
                      "contact": {
                        "name": "Samuel L. Banina",
                        "email": "saban@microsoft.com"
                      }
                    },
                    "host": "europe-001.azure-apim.net",
                    "basePath": "/apim/flowpush",
                    "schemes": [
                      "https"
                    ],
                    "consumes": [
                      "application/json"
                    ],
                    "produces": [
                      "application/json"
                    ],
                    "x-ms-capabilities": {
                      "buttons": {
                        "flowIosApp": {},
                        "flowAndroidApp": {}
                      }
                    },
                    "definitions": {
                      "NotificationDefinition": {
                        "type": "object",
                        "required": [
                          "notificationText"
                        ],
                        "properties": {
                          "notificationText": {
                            "description": "Create a notification message",
                            "x-ms-summary": "Text",
                            "type": "string"
                          },
                          "notificationLink": {
                            "description": "Custom notification link",
                            "type": "object",
                            "properties": {
                              "uri": {
                                "description": "Include a link in the notification",
                                "x-ms-summary": "Link",
                                "type": "string"
                              },
                              "label": {
                                "description": "The display name for the link",
                                "x-ms-summary": "Link label",
                                "type": "string"
                              }
                            }
                          }
                        }
                      },
                      "NotificationEmailDefinition": {
                        "type": "object",
                        "required": [
                          "notificationSubject",
                          "notificationBody"
                        ],
                        "properties": {
                          "notificationSubject": {
                            "description": "Notification email subject",
                            "x-ms-summary": "Subject",
                            "type": "string"
                          },
                          "notificationBody": {
                            "description": "Notification email body",
                            "x-ms-summary": "Body",
                            "type": "string"
                          }
                        }
                      }
                    },
                    "paths": {
                      "/{connectionId}/sendNotification": {
                        "post": {
                          "description": "Sends a push notification to your Microsoft Flow mobile app.",
                          "summary": "Send me a mobile notification",
                          "operationId": "SendNotification",
                          "x-ms-visibility": "important",
                          "consumes": [
                            "application/json"
                          ],
                          "produces": [
                            "application/json"
                          ],
                          "parameters": [
                            {
                              "name": "connectionId",
                              "in": "path",
                              "required": true,
                              "type": "string",
                              "x-ms-visibility": "internal"
                            },
                            {
                              "name": "NotificationDefinition",
                              "x-ms-summary": "The push notification",
                              "in": "body",
                              "description": "Push notification inputs",
                              "required": true,
                              "schema": {
                                "$ref": "#/definitions/NotificationDefinition"
                              }
                            }
                          ],
                          "responses": {
                            "200": {
                              "description": "OK"
                            },
                            "400": {
                              "description": "Bad Request"
                            },
                            "500": {
                              "description": "Internal Server Error"
                            },
                            "default": {
                              "description": "Operation Failed."
                            }
                          }
                        }
                      },
                      "/{connectionId}/sendEmailNotification": {
                        "post": {
                          "description": "Sends an email notification to the account you signed in to Microsoft Flow with.",
                          "summary": "Send me an email notification",
                          "operationId": "SendEmailNotification",
                          "x-ms-visibility": "important",
                          "consumes": [
                            "application/json"
                          ],
                          "produces": [
                            "application/json"
                          ],
                          "parameters": [
                            {
                              "name": "connectionId",
                              "in": "path",
                              "required": true,
                              "type": "string",
                              "x-ms-visibility": "internal"
                            },
                            {
                              "name": "NotificationEmailDefinition",
                              "x-ms-summary": "The email notification",
                              "in": "body",
                              "description": "Email notification inputs",
                              "required": true,
                              "schema": {
                                "$ref": "#/definitions/NotificationEmailDefinition"
                              }
                            }
                          ],
                          "responses": {
                            "200": {
                              "description": "OK"
                            },
                            "400": {
                              "description": "Bad Request"
                            },
                            "500": {
                              "description": "Internal Server Error"
                            },
                            "default": {
                              "description": "Operation Failed."
                            }
                          }
                        }
                      }
                    }
                  },
                  "tier": "NotSpecified"
                }
              },
              "createdTime": "2018-03-23T17:59:35.4407282Z",
              "lastModifiedTime": "2018-03-23T17:59:37.1164508Z",
              "templateName": "a04de6ce52984b3db0b907f588994bc8",
              "environment": {
                "name": "Default-d87a7535-dd31-4437-bfe1-95340acd55c5",
                "type": "Microsoft.ProcessSimple/environments",
                "id": "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5"
              },
              "definitionSummary": {
                "triggers": [
                  {
                    "type": "Recurrence"
                  }
                ],
                "actions": [
                  {
                    "type": "If"
                  },
                  {
                    "type": "Query"
                  },
                  {
                    "type": "ApiConnection",
                    "swaggerOperationId": "ListFeedItems",
                    "metadata": {
                      "flowSystemMetadata": {
                        "swaggerOperationId": "ListFeedItems"
                      }
                    },
                    "api": {
                      "name": "shared_rss",
                      "id": "/providers/Microsoft.PowerApps/apis/shared_rss",
                      "type": "/providers/Microsoft.PowerApps/apis"
                    }
                  },
                  {
                    "type": "Foreach"
                  },
                  {
                    "type": "ApiConnection",
                    "swaggerOperationId": "SendEmailNotification",
                    "metadata": {
                      "flowSystemMetadata": {
                        "swaggerOperationId": "SendEmailNotification"
                      }
                    },
                    "api": {
                      "name": "shared_flowpush",
                      "id": "/providers/Microsoft.PowerApps/apis/shared_flowpush",
                      "type": "/providers/Microsoft.PowerApps/apis"
                    }
                  },
                  {
                    "type": "Compose"
                  }
                ],
                "description": "Each day, get an email with a list of all of the top CNN posts from the last day."
              },
              "creator": {
                "tenantId": "d87a7535-dd31-4437-bfe1-95340acd55c5",
                "objectId": "da8f7aea-cf43-497f-ad62-c2feae89a194",
                "userId": "da8f7aea-cf43-497f-ad62-c2feae89a194",
                "userType": "ActiveDirectory"
              },
              "flowTriggerUri": "https://management.azure.com:443/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/3989cb59-ce1a-4a5c-bb78-257c5c39381d/triggers/Every_day/run?api-version=2016-11-01",
              "installationStatus": "Installed",
              "provisioningMethod": "FromDefinition",
              "flowFailureAlertSubscribed": true,
              "referencedResources": []
            }
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: '3989cb59-ce1a-4a5c-bb78-257c5c39381d', environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          name: '3989cb59-ce1a-4a5c-bb78-257c5c39381d',
          displayName: 'Get a daily digest of the top CNN news',
          description: 'Each day, get an email with a list of all of the top CNN posts from the last day.',
          triggers: 'Every_day',
          actions: 'Check_if_there_were_any_posts_this_week, Filter_array, List_all_RSS_feed_items'
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves information about the specified flow as admin', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.ProcessSimple/scopes/admin/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/3989cb59-ce1a-4a5c-bb78-257c5c39381d?api-version=2016-11-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            "name": "3989cb59-ce1a-4a5c-bb78-257c5c39381d",
            "id": "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/3989cb59-ce1a-4a5c-bb78-257c5c39381d",
            "type": "Microsoft.ProcessSimple/environments/flows",
            "properties": {
              "apiId": "/providers/Microsoft.PowerApps/apis/shared_logicflows",
              "displayName": "Get a daily digest of the top CNN news",
              "userType": "Owner",
              "definition": {
                "$schema": "https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#",
                "contentVersion": "1.0.0.0",
                "parameters": {
                  "$connections": {
                    "defaultValue": {},
                    "type": "Object"
                  },
                  "$authentication": {
                    "defaultValue": {},
                    "type": "SecureObject"
                  }
                },
                "triggers": {
                  "Every_day": {
                    "recurrence": {
                      "frequency": "Day",
                      "interval": 1
                    },
                    "type": "Recurrence"
                  }
                },
                "actions": {
                  "Check_if_there_were_any_posts_this_week": {
                    "actions": {
                      "Compose_the_links_for_each_blog_post": {
                        "foreach": "@take(body('Filter_array'),15)",
                        "actions": {
                          "Compose": {
                            "runAfter": {},
                            "type": "Compose",
                            "inputs": "<tr><td><h3><a href=\"@{item()?['primaryLink']}\">@{item()['title']}</a></h3></td></tr><tr><td style=\"color: #777777;\">Posted at @{formatDateTime(item()?['publishDate'], 't')} GMT</td></tr><tr><td>@{item()?['summary']}</td></tr>"
                          }
                        },
                        "runAfter": {},
                        "type": "Foreach"
                      },
                      "Send_an_email": {
                        "runAfter": {
                          "Compose_the_links_for_each_blog_post": [
                            "Succeeded"
                          ]
                        },
                        "metadata": {
                          "flowSystemMetadata": {
                            "swaggerOperationId": "SendEmailNotification"
                          }
                        },
                        "type": "ApiConnection",
                        "inputs": {
                          "body": {
                            "notificationBody": "<h2>Check out the latest news from CNN:</h2><table>@{join(outputs('Compose'),'')}</table>",
                            "notificationSubject": "Daily digest from CNN top news"
                          },
                          "host": {
                            "api": {
                              "runtimeUrl": "https://europe-001.azure-apim.net/apim/flowpush"
                            },
                            "connection": {
                              "name": "@parameters('$connections')['shared_flowpush']['connectionId']"
                            }
                          },
                          "method": "post",
                          "path": "/sendEmailNotification",
                          "authentication": "@parameters('$authentication')"
                        }
                      }
                    },
                    "runAfter": {
                      "Filter_array": [
                        "Succeeded"
                      ]
                    },
                    "expression": "@greater(length(body('Filter_array')), 0)",
                    "type": "If"
                  },
                  "Filter_array": {
                    "runAfter": {
                      "List_all_RSS_feed_items": [
                        "Succeeded"
                      ]
                    },
                    "type": "Query",
                    "inputs": {
                      "from": "@body('List_all_RSS_feed_items')",
                      "where": "@greater(item()?['publishDate'], adddays(utcnow(),-2))"
                    }
                  },
                  "List_all_RSS_feed_items": {
                    "runAfter": {},
                    "metadata": {
                      "flowSystemMetadata": {
                        "swaggerOperationId": "ListFeedItems"
                      }
                    },
                    "type": "ApiConnection",
                    "inputs": {
                      "host": {
                        "api": {
                          "runtimeUrl": "https://europe-001.azure-apim.net/apim/rss"
                        },
                        "connection": {
                          "name": "@parameters('$connections')['shared_rss']['connectionId']"
                        }
                      },
                      "method": "get",
                      "path": "/ListFeedItems",
                      "queries": {
                        "feedUrl": "http://rss.cnn.com/rss/cnn_topstories.rss"
                      },
                      "authentication": "@parameters('$authentication')"
                    }
                  }
                },
                "outputs": {},
                "description": "Each day, get an email with a list of all of the top CNN posts from the last day."
              },
              "state": "Started",
              "connectionReferences": {
                "shared_rss": {
                  "connectionName": "shared-rss-6636bfd3-0d29-4842-b835-c5910b6310f6",
                  "apiDefinition": {
                    "name": "shared_rss",
                    "id": "/providers/Microsoft.PowerApps/apis/shared_rss",
                    "type": "/providers/Microsoft.PowerApps/apis",
                    "properties": {
                      "displayName": "RSS",
                      "iconUri": "https://az818438.vo.msecnd.net/icons/rss.png",
                      "purpose": "NotSpecified",
                      "connectionParameters": {},
                      "scopes": {
                        "will": [],
                        "wont": []
                      },
                      "runtimeUrls": [
                        "https://europe-001.azure-apim.net/apim/rss"
                      ],
                      "primaryRuntimeUrl": "https://europe-001.azure-apim.net/apim/rss",
                      "metadata": {
                        "source": "marketplace",
                        "brandColor": "#ff9900"
                      },
                      "capabilities": [
                        "actions"
                      ],
                      "tier": "Standard",
                      "description": "RSS is a popular web syndication format used to publish frequently updated content – like blog entries and news headlines.  Many content publishers provide an RSS feed to allow users to subscribe to it.  Use the RSS connector to retrieve feed information and trigger flows when new items are published in an RSS feed.",
                      "createdTime": "2016-09-30T04:13:14.2871915Z",
                      "changedTime": "2018-01-17T20:20:37.905252Z"
                    }
                  },
                  "source": "Embedded",
                  "id": "/providers/Microsoft.PowerApps/apis/shared_rss",
                  "displayName": "RSS",
                  "iconUri": "https://az818438.vo.msecnd.net/icons/rss.png",
                  "brandColor": "#ff9900",
                  "swagger": {
                    "swagger": "2.0",
                    "info": {
                      "version": "1.0",
                      "title": "RSS",
                      "description": "RSS is a popular web syndication format used to publish frequently updated content – like blog entries and news headlines.  Many content publishers provide an RSS feed to allow users to subscribe to it.  Use the RSS connector to retrieve feed information and trigger flows when new items are published in an RSS feed.",
                      "x-ms-api-annotation": {
                        "status": "Production"
                      }
                    },
                    "host": "europe-001.azure-apim.net",
                    "basePath": "/apim/rss",
                    "schemes": [
                      "https"
                    ],
                    "paths": {
                      "/{connectionId}/OnNewFeed": {
                        "get": {
                          "tags": [
                            "Rss"
                          ],
                          "summary": "When a feed item is published",
                          "description": "This operation triggers a workflow when a new item is published in an RSS feed.",
                          "operationId": "OnNewFeed",
                          "consumes": [],
                          "produces": [
                            "application/json",
                            "text/json",
                            "application/xml",
                            "text/xml"
                          ],
                          "parameters": [
                            {
                              "name": "connectionId",
                              "in": "path",
                              "required": true,
                              "type": "string",
                              "x-ms-visibility": "internal"
                            },
                            {
                              "name": "feedUrl",
                              "in": "query",
                              "description": "The RSS feed URL (Example: http://rss.cnn.com/rss/cnn_topstories.rss).",
                              "required": true,
                              "x-ms-summary": "The RSS feed URL",
                              "x-ms-url-encoding": "double",
                              "type": "string"
                            }
                          ],
                          "responses": {
                            "200": {
                              "description": "OK",
                              "schema": {
                                "$ref": "#/definitions/TriggerBatchResponse[FeedItem]"
                              }
                            },
                            "202": {
                              "description": "Accepted"
                            },
                            "400": {
                              "description": "Bad Request"
                            },
                            "401": {
                              "description": "Unauthorized"
                            },
                            "403": {
                              "description": "Forbidden"
                            },
                            "404": {
                              "description": "Not Found"
                            },
                            "500": {
                              "description": "Internal Server Error. Unknown error occurred"
                            },
                            "default": {
                              "description": "Operation Failed."
                            }
                          },
                          "deprecated": false,
                          "x-ms-visibility": "important",
                          "x-ms-trigger": "batch",
                          "x-ms-trigger-hint": "To see it work now, publish an item to the RSS feed."
                        }
                      },
                      "/{connectionId}/ListFeedItems": {
                        "get": {
                          "tags": [
                            "Rss"
                          ],
                          "summary": "List all RSS feed items",
                          "description": "This operation retrieves all items from an RSS feed.",
                          "operationId": "ListFeedItems",
                          "consumes": [],
                          "produces": [
                            "application/json",
                            "text/json",
                            "application/xml",
                            "text/xml"
                          ],
                          "parameters": [
                            {
                              "name": "connectionId",
                              "in": "path",
                              "required": true,
                              "type": "string",
                              "x-ms-visibility": "internal"
                            },
                            {
                              "name": "feedUrl",
                              "in": "query",
                              "description": "The RSS feed URL (Example: http://rss.cnn.com/rss/cnn_topstories.rss).",
                              "required": true,
                              "x-ms-summary": "The RSS feed URL",
                              "x-ms-url-encoding": "double",
                              "type": "string"
                            }
                          ],
                          "responses": {
                            "200": {
                              "description": "OK",
                              "schema": {
                                "type": "array",
                                "items": {
                                  "$ref": "#/definitions/FeedItem"
                                }
                              }
                            },
                            "202": {
                              "description": "Accepted"
                            },
                            "400": {
                              "description": "Bad Request"
                            },
                            "401": {
                              "description": "Unauthorized"
                            },
                            "403": {
                              "description": "Forbidden"
                            },
                            "404": {
                              "description": "Not Found"
                            },
                            "500": {
                              "description": "Internal Server Error. Unknown error occurred"
                            },
                            "default": {
                              "description": "Operation Failed."
                            }
                          },
                          "deprecated": false,
                          "x-ms-visibility": "important"
                        }
                      }
                    },
                    "definitions": {
                      "TriggerBatchResponse[FeedItem]": {
                        "description": "Represents a wrapper object for batch trigger response",
                        "type": "object",
                        "properties": {
                          "value": {
                            "description": "A list of the response objects",
                            "type": "array",
                            "items": {
                              "$ref": "#/definitions/FeedItem"
                            }
                          }
                        }
                      },
                      "FeedItem": {
                        "description": "Represents an RSS feed item.",
                        "required": [
                          "id",
                          "title"
                        ],
                        "type": "object",
                        "properties": {
                          "id": {
                            "description": "Feed ID",
                            "type": "string",
                            "x-ms-summary": "Feed ID"
                          },
                          "title": {
                            "description": "Feed title",
                            "type": "string",
                            "x-ms-summary": "Feed title"
                          },
                          "primaryLink": {
                            "description": "Primary feed link",
                            "type": "string",
                            "x-ms-summary": "Primary feed link"
                          },
                          "links": {
                            "description": "Feed links",
                            "type": "array",
                            "items": {
                              "type": "string"
                            },
                            "x-ms-summary": "Feed links",
                            "x-ms-visibility": "advanced"
                          },
                          "updatedOn": {
                            "description": "Feed updated on",
                            "type": "string",
                            "x-ms-summary": "Feed updated on"
                          },
                          "publishDate": {
                            "description": "Feed published date",
                            "type": "string",
                            "x-ms-summary": "Feed published on"
                          },
                          "summary": {
                            "description": "Feed item summary",
                            "type": "string",
                            "x-ms-summary": "Feed summary"
                          },
                          "copyright": {
                            "description": "Copyright information",
                            "type": "string",
                            "x-ms-summary": "Feed copyright information"
                          },
                          "categories": {
                            "description": "Feed item categories",
                            "type": "array",
                            "items": {
                              "type": "string"
                            },
                            "x-ms-summary": "Feed categories"
                          }
                        }
                      }
                    }
                  },
                  "tier": "NotSpecified"
                },
                "shared_flowpush": {
                  "connectionName": "shared-flowpush-295e4b80-1a4e-42ec-aa5b-8d72e7c1eb3f",
                  "apiDefinition": {
                    "name": "shared_flowpush",
                    "id": "/providers/Microsoft.PowerApps/apis/shared_flowpush",
                    "type": "/providers/Microsoft.PowerApps/apis",
                    "properties": {
                      "displayName": "Notifications",
                      "iconUri": "https://psux.azureedge.net/Content/Images/Connectors/FlowNotification.svg",
                      "purpose": "NotSpecified",
                      "connectionParameters": {},
                      "runtimeUrls": [
                        "https://europe-001.azure-apim.net/apim/flowpush"
                      ],
                      "primaryRuntimeUrl": "https://europe-001.azure-apim.net/apim/flowpush",
                      "metadata": {
                        "source": "marketplace",
                        "brandColor": "#FF3B30",
                        "connectionLimits": {
                          "*": 1
                        }
                      },
                      "capabilities": [
                        "actions"
                      ],
                      "tier": "NotSpecified",
                      "description": "The notification service enables notifications created by a flow to go to your email account or Microsoft Flow mobile app.",
                      "createdTime": "2016-10-22T06:43:26.1572419Z",
                      "changedTime": "2018-01-25T00:34:52.048009Z"
                    }
                  },
                  "source": "Embedded",
                  "id": "/providers/Microsoft.PowerApps/apis/shared_flowpush",
                  "displayName": "Notifications",
                  "iconUri": "https://psux.azureedge.net/Content/Images/Connectors/FlowNotification.svg",
                  "brandColor": "#FF3B30",
                  "swagger": {
                    "swagger": "2.0",
                    "info": {
                      "version": "1.0",
                      "title": "Notifications",
                      "description": "The notification service enables notifications created by a flow to go to your email account or Microsoft Flow mobile app.",
                      "contact": {
                        "name": "Samuel L. Banina",
                        "email": "saban@microsoft.com"
                      }
                    },
                    "host": "europe-001.azure-apim.net",
                    "basePath": "/apim/flowpush",
                    "schemes": [
                      "https"
                    ],
                    "consumes": [
                      "application/json"
                    ],
                    "produces": [
                      "application/json"
                    ],
                    "x-ms-capabilities": {
                      "buttons": {
                        "flowIosApp": {},
                        "flowAndroidApp": {}
                      }
                    },
                    "definitions": {
                      "NotificationDefinition": {
                        "type": "object",
                        "required": [
                          "notificationText"
                        ],
                        "properties": {
                          "notificationText": {
                            "description": "Create a notification message",
                            "x-ms-summary": "Text",
                            "type": "string"
                          },
                          "notificationLink": {
                            "description": "Custom notification link",
                            "type": "object",
                            "properties": {
                              "uri": {
                                "description": "Include a link in the notification",
                                "x-ms-summary": "Link",
                                "type": "string"
                              },
                              "label": {
                                "description": "The display name for the link",
                                "x-ms-summary": "Link label",
                                "type": "string"
                              }
                            }
                          }
                        }
                      },
                      "NotificationEmailDefinition": {
                        "type": "object",
                        "required": [
                          "notificationSubject",
                          "notificationBody"
                        ],
                        "properties": {
                          "notificationSubject": {
                            "description": "Notification email subject",
                            "x-ms-summary": "Subject",
                            "type": "string"
                          },
                          "notificationBody": {
                            "description": "Notification email body",
                            "x-ms-summary": "Body",
                            "type": "string"
                          }
                        }
                      }
                    },
                    "paths": {
                      "/{connectionId}/sendNotification": {
                        "post": {
                          "description": "Sends a push notification to your Microsoft Flow mobile app.",
                          "summary": "Send me a mobile notification",
                          "operationId": "SendNotification",
                          "x-ms-visibility": "important",
                          "consumes": [
                            "application/json"
                          ],
                          "produces": [
                            "application/json"
                          ],
                          "parameters": [
                            {
                              "name": "connectionId",
                              "in": "path",
                              "required": true,
                              "type": "string",
                              "x-ms-visibility": "internal"
                            },
                            {
                              "name": "NotificationDefinition",
                              "x-ms-summary": "The push notification",
                              "in": "body",
                              "description": "Push notification inputs",
                              "required": true,
                              "schema": {
                                "$ref": "#/definitions/NotificationDefinition"
                              }
                            }
                          ],
                          "responses": {
                            "200": {
                              "description": "OK"
                            },
                            "400": {
                              "description": "Bad Request"
                            },
                            "500": {
                              "description": "Internal Server Error"
                            },
                            "default": {
                              "description": "Operation Failed."
                            }
                          }
                        }
                      },
                      "/{connectionId}/sendEmailNotification": {
                        "post": {
                          "description": "Sends an email notification to the account you signed in to Microsoft Flow with.",
                          "summary": "Send me an email notification",
                          "operationId": "SendEmailNotification",
                          "x-ms-visibility": "important",
                          "consumes": [
                            "application/json"
                          ],
                          "produces": [
                            "application/json"
                          ],
                          "parameters": [
                            {
                              "name": "connectionId",
                              "in": "path",
                              "required": true,
                              "type": "string",
                              "x-ms-visibility": "internal"
                            },
                            {
                              "name": "NotificationEmailDefinition",
                              "x-ms-summary": "The email notification",
                              "in": "body",
                              "description": "Email notification inputs",
                              "required": true,
                              "schema": {
                                "$ref": "#/definitions/NotificationEmailDefinition"
                              }
                            }
                          ],
                          "responses": {
                            "200": {
                              "description": "OK"
                            },
                            "400": {
                              "description": "Bad Request"
                            },
                            "500": {
                              "description": "Internal Server Error"
                            },
                            "default": {
                              "description": "Operation Failed."
                            }
                          }
                        }
                      }
                    }
                  },
                  "tier": "NotSpecified"
                }
              },
              "createdTime": "2018-03-23T17:59:35.4407282Z",
              "lastModifiedTime": "2018-03-23T17:59:37.1164508Z",
              "templateName": "a04de6ce52984b3db0b907f588994bc8",
              "environment": {
                "name": "Default-d87a7535-dd31-4437-bfe1-95340acd55c5",
                "type": "Microsoft.ProcessSimple/environments",
                "id": "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5"
              },
              "definitionSummary": {
                "triggers": [
                  {
                    "type": "Recurrence"
                  }
                ],
                "actions": [
                  {
                    "type": "If"
                  },
                  {
                    "type": "Query"
                  },
                  {
                    "type": "ApiConnection",
                    "swaggerOperationId": "ListFeedItems",
                    "metadata": {
                      "flowSystemMetadata": {
                        "swaggerOperationId": "ListFeedItems"
                      }
                    },
                    "api": {
                      "name": "shared_rss",
                      "id": "/providers/Microsoft.PowerApps/apis/shared_rss",
                      "type": "/providers/Microsoft.PowerApps/apis"
                    }
                  },
                  {
                    "type": "Foreach"
                  },
                  {
                    "type": "ApiConnection",
                    "swaggerOperationId": "SendEmailNotification",
                    "metadata": {
                      "flowSystemMetadata": {
                        "swaggerOperationId": "SendEmailNotification"
                      }
                    },
                    "api": {
                      "name": "shared_flowpush",
                      "id": "/providers/Microsoft.PowerApps/apis/shared_flowpush",
                      "type": "/providers/Microsoft.PowerApps/apis"
                    }
                  },
                  {
                    "type": "Compose"
                  }
                ],
                "description": "Each day, get an email with a list of all of the top CNN posts from the last day."
              },
              "creator": {
                "tenantId": "d87a7535-dd31-4437-bfe1-95340acd55c5",
                "objectId": "da8f7aea-cf43-497f-ad62-c2feae89a194",
                "userId": "da8f7aea-cf43-497f-ad62-c2feae89a194",
                "userType": "ActiveDirectory"
              },
              "flowTriggerUri": "https://management.azure.com:443/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/3989cb59-ce1a-4a5c-bb78-257c5c39381d/triggers/Every_day/run?api-version=2016-11-01",
              "installationStatus": "Installed",
              "provisioningMethod": "FromDefinition",
              "flowFailureAlertSubscribed": true,
              "referencedResources": []
            }
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: '3989cb59-ce1a-4a5c-bb78-257c5c39381d', environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5', asAdmin: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          name: '3989cb59-ce1a-4a5c-bb78-257c5c39381d',
          displayName: 'Get a daily digest of the top CNN news',
          description: 'Each day, get an email with a list of all of the top CNN posts from the last day.',
          triggers: 'Every_day',
          actions: 'Check_if_there_were_any_posts_this_week, Filter_array, List_all_RSS_feed_items'
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns all properties for JSON output', (done) => {
    const flowInfo: any = {
      "name": "3989cb59-ce1a-4a5c-bb78-257c5c39381d",
      "id": "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/3989cb59-ce1a-4a5c-bb78-257c5c39381d",
      "type": "Microsoft.ProcessSimple/environments/flows",
      "properties": {
        "apiId": "/providers/Microsoft.PowerApps/apis/shared_logicflows",
        "displayName": "Get a daily digest of the top CNN news",
        "userType": "Owner",
        "definition": {
          "$schema": "https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#",
          "contentVersion": "1.0.0.0",
          "parameters": {
            "$connections": {
              "defaultValue": {},
              "type": "Object"
            },
            "$authentication": {
              "defaultValue": {},
              "type": "SecureObject"
            }
          },
          "triggers": {
            "Every_day": {
              "recurrence": {
                "frequency": "Day",
                "interval": 1
              },
              "type": "Recurrence"
            }
          },
          "actions": {
            "Check_if_there_were_any_posts_this_week": {
              "actions": {
                "Compose_the_links_for_each_blog_post": {
                  "foreach": "@take(body('Filter_array'),15)",
                  "actions": {
                    "Compose": {
                      "runAfter": {},
                      "type": "Compose",
                      "inputs": "<tr><td><h3><a href=\"@{item()?['primaryLink']}\">@{item()['title']}</a></h3></td></tr><tr><td style=\"color: #777777;\">Posted at @{formatDateTime(item()?['publishDate'], 't')} GMT</td></tr><tr><td>@{item()?['summary']}</td></tr>"
                    }
                  },
                  "runAfter": {},
                  "type": "Foreach"
                },
                "Send_an_email": {
                  "runAfter": {
                    "Compose_the_links_for_each_blog_post": [
                      "Succeeded"
                    ]
                  },
                  "metadata": {
                    "flowSystemMetadata": {
                      "swaggerOperationId": "SendEmailNotification"
                    }
                  },
                  "type": "ApiConnection",
                  "inputs": {
                    "body": {
                      "notificationBody": "<h2>Check out the latest news from CNN:</h2><table>@{join(outputs('Compose'),'')}</table>",
                      "notificationSubject": "Daily digest from CNN top news"
                    },
                    "host": {
                      "api": {
                        "runtimeUrl": "https://europe-001.azure-apim.net/apim/flowpush"
                      },
                      "connection": {
                        "name": "@parameters('$connections')['shared_flowpush']['connectionId']"
                      }
                    },
                    "method": "post",
                    "path": "/sendEmailNotification",
                    "authentication": "@parameters('$authentication')"
                  }
                }
              },
              "runAfter": {
                "Filter_array": [
                  "Succeeded"
                ]
              },
              "expression": "@greater(length(body('Filter_array')), 0)",
              "type": "If"
            },
            "Filter_array": {
              "runAfter": {
                "List_all_RSS_feed_items": [
                  "Succeeded"
                ]
              },
              "type": "Query",
              "inputs": {
                "from": "@body('List_all_RSS_feed_items')",
                "where": "@greater(item()?['publishDate'], adddays(utcnow(),-2))"
              }
            },
            "List_all_RSS_feed_items": {
              "runAfter": {},
              "metadata": {
                "flowSystemMetadata": {
                  "swaggerOperationId": "ListFeedItems"
                }
              },
              "type": "ApiConnection",
              "inputs": {
                "host": {
                  "api": {
                    "runtimeUrl": "https://europe-001.azure-apim.net/apim/rss"
                  },
                  "connection": {
                    "name": "@parameters('$connections')['shared_rss']['connectionId']"
                  }
                },
                "method": "get",
                "path": "/ListFeedItems",
                "queries": {
                  "feedUrl": "http://rss.cnn.com/rss/cnn_topstories.rss"
                },
                "authentication": "@parameters('$authentication')"
              }
            }
          },
          "outputs": {},
          "description": "Each day, get an email with a list of all of the top CNN posts from the last day."
        },
        "state": "Started",
        "connectionReferences": {
          "shared_rss": {
            "connectionName": "shared-rss-6636bfd3-0d29-4842-b835-c5910b6310f6",
            "apiDefinition": {
              "name": "shared_rss",
              "id": "/providers/Microsoft.PowerApps/apis/shared_rss",
              "type": "/providers/Microsoft.PowerApps/apis",
              "properties": {
                "displayName": "RSS",
                "iconUri": "https://az818438.vo.msecnd.net/icons/rss.png",
                "purpose": "NotSpecified",
                "connectionParameters": {},
                "scopes": {
                  "will": [],
                  "wont": []
                },
                "runtimeUrls": [
                  "https://europe-001.azure-apim.net/apim/rss"
                ],
                "primaryRuntimeUrl": "https://europe-001.azure-apim.net/apim/rss",
                "metadata": {
                  "source": "marketplace",
                  "brandColor": "#ff9900"
                },
                "capabilities": [
                  "actions"
                ],
                "tier": "Standard",
                "description": "RSS is a popular web syndication format used to publish frequently updated content – like blog entries and news headlines.  Many content publishers provide an RSS feed to allow users to subscribe to it.  Use the RSS connector to retrieve feed information and trigger flows when new items are published in an RSS feed.",
                "createdTime": "2016-09-30T04:13:14.2871915Z",
                "changedTime": "2018-01-17T20:20:37.905252Z"
              }
            },
            "source": "Embedded",
            "id": "/providers/Microsoft.PowerApps/apis/shared_rss",
            "displayName": "RSS",
            "iconUri": "https://az818438.vo.msecnd.net/icons/rss.png",
            "brandColor": "#ff9900",
            "swagger": {
              "swagger": "2.0",
              "info": {
                "version": "1.0",
                "title": "RSS",
                "description": "RSS is a popular web syndication format used to publish frequently updated content – like blog entries and news headlines.  Many content publishers provide an RSS feed to allow users to subscribe to it.  Use the RSS connector to retrieve feed information and trigger flows when new items are published in an RSS feed.",
                "x-ms-api-annotation": {
                  "status": "Production"
                }
              },
              "host": "europe-001.azure-apim.net",
              "basePath": "/apim/rss",
              "schemes": [
                "https"
              ],
              "paths": {
                "/{connectionId}/OnNewFeed": {
                  "get": {
                    "tags": [
                      "Rss"
                    ],
                    "summary": "When a feed item is published",
                    "description": "This operation triggers a workflow when a new item is published in an RSS feed.",
                    "operationId": "OnNewFeed",
                    "consumes": [],
                    "produces": [
                      "application/json",
                      "text/json",
                      "application/xml",
                      "text/xml"
                    ],
                    "parameters": [
                      {
                        "name": "connectionId",
                        "in": "path",
                        "required": true,
                        "type": "string",
                        "x-ms-visibility": "internal"
                      },
                      {
                        "name": "feedUrl",
                        "in": "query",
                        "description": "The RSS feed URL (Example: http://rss.cnn.com/rss/cnn_topstories.rss).",
                        "required": true,
                        "x-ms-summary": "The RSS feed URL",
                        "x-ms-url-encoding": "double",
                        "type": "string"
                      }
                    ],
                    "responses": {
                      "200": {
                        "description": "OK",
                        "schema": {
                          "$ref": "#/definitions/TriggerBatchResponse[FeedItem]"
                        }
                      },
                      "202": {
                        "description": "Accepted"
                      },
                      "400": {
                        "description": "Bad Request"
                      },
                      "401": {
                        "description": "Unauthorized"
                      },
                      "403": {
                        "description": "Forbidden"
                      },
                      "404": {
                        "description": "Not Found"
                      },
                      "500": {
                        "description": "Internal Server Error. Unknown error occurred"
                      },
                      "default": {
                        "description": "Operation Failed."
                      }
                    },
                    "deprecated": false,
                    "x-ms-visibility": "important",
                    "x-ms-trigger": "batch",
                    "x-ms-trigger-hint": "To see it work now, publish an item to the RSS feed."
                  }
                },
                "/{connectionId}/ListFeedItems": {
                  "get": {
                    "tags": [
                      "Rss"
                    ],
                    "summary": "List all RSS feed items",
                    "description": "This operation retrieves all items from an RSS feed.",
                    "operationId": "ListFeedItems",
                    "consumes": [],
                    "produces": [
                      "application/json",
                      "text/json",
                      "application/xml",
                      "text/xml"
                    ],
                    "parameters": [
                      {
                        "name": "connectionId",
                        "in": "path",
                        "required": true,
                        "type": "string",
                        "x-ms-visibility": "internal"
                      },
                      {
                        "name": "feedUrl",
                        "in": "query",
                        "description": "The RSS feed URL (Example: http://rss.cnn.com/rss/cnn_topstories.rss).",
                        "required": true,
                        "x-ms-summary": "The RSS feed URL",
                        "x-ms-url-encoding": "double",
                        "type": "string"
                      }
                    ],
                    "responses": {
                      "200": {
                        "description": "OK",
                        "schema": {
                          "type": "array",
                          "items": {
                            "$ref": "#/definitions/FeedItem"
                          }
                        }
                      },
                      "202": {
                        "description": "Accepted"
                      },
                      "400": {
                        "description": "Bad Request"
                      },
                      "401": {
                        "description": "Unauthorized"
                      },
                      "403": {
                        "description": "Forbidden"
                      },
                      "404": {
                        "description": "Not Found"
                      },
                      "500": {
                        "description": "Internal Server Error. Unknown error occurred"
                      },
                      "default": {
                        "description": "Operation Failed."
                      }
                    },
                    "deprecated": false,
                    "x-ms-visibility": "important"
                  }
                }
              },
              "definitions": {
                "TriggerBatchResponse[FeedItem]": {
                  "description": "Represents a wrapper object for batch trigger response",
                  "type": "object",
                  "properties": {
                    "value": {
                      "description": "A list of the response objects",
                      "type": "array",
                      "items": {
                        "$ref": "#/definitions/FeedItem"
                      }
                    }
                  }
                },
                "FeedItem": {
                  "description": "Represents an RSS feed item.",
                  "required": [
                    "id",
                    "title"
                  ],
                  "type": "object",
                  "properties": {
                    "id": {
                      "description": "Feed ID",
                      "type": "string",
                      "x-ms-summary": "Feed ID"
                    },
                    "title": {
                      "description": "Feed title",
                      "type": "string",
                      "x-ms-summary": "Feed title"
                    },
                    "primaryLink": {
                      "description": "Primary feed link",
                      "type": "string",
                      "x-ms-summary": "Primary feed link"
                    },
                    "links": {
                      "description": "Feed links",
                      "type": "array",
                      "items": {
                        "type": "string"
                      },
                      "x-ms-summary": "Feed links",
                      "x-ms-visibility": "advanced"
                    },
                    "updatedOn": {
                      "description": "Feed updated on",
                      "type": "string",
                      "x-ms-summary": "Feed updated on"
                    },
                    "publishDate": {
                      "description": "Feed published date",
                      "type": "string",
                      "x-ms-summary": "Feed published on"
                    },
                    "summary": {
                      "description": "Feed item summary",
                      "type": "string",
                      "x-ms-summary": "Feed summary"
                    },
                    "copyright": {
                      "description": "Copyright information",
                      "type": "string",
                      "x-ms-summary": "Feed copyright information"
                    },
                    "categories": {
                      "description": "Feed item categories",
                      "type": "array",
                      "items": {
                        "type": "string"
                      },
                      "x-ms-summary": "Feed categories"
                    }
                  }
                }
              }
            },
            "tier": "NotSpecified"
          },
          "shared_flowpush": {
            "connectionName": "shared-flowpush-295e4b80-1a4e-42ec-aa5b-8d72e7c1eb3f",
            "apiDefinition": {
              "name": "shared_flowpush",
              "id": "/providers/Microsoft.PowerApps/apis/shared_flowpush",
              "type": "/providers/Microsoft.PowerApps/apis",
              "properties": {
                "displayName": "Notifications",
                "iconUri": "https://psux.azureedge.net/Content/Images/Connectors/FlowNotification.svg",
                "purpose": "NotSpecified",
                "connectionParameters": {},
                "runtimeUrls": [
                  "https://europe-001.azure-apim.net/apim/flowpush"
                ],
                "primaryRuntimeUrl": "https://europe-001.azure-apim.net/apim/flowpush",
                "metadata": {
                  "source": "marketplace",
                  "brandColor": "#FF3B30",
                  "connectionLimits": {
                    "*": 1
                  }
                },
                "capabilities": [
                  "actions"
                ],
                "tier": "NotSpecified",
                "description": "The notification service enables notifications created by a flow to go to your email account or Microsoft Flow mobile app.",
                "createdTime": "2016-10-22T06:43:26.1572419Z",
                "changedTime": "2018-01-25T00:34:52.048009Z"
              }
            },
            "source": "Embedded",
            "id": "/providers/Microsoft.PowerApps/apis/shared_flowpush",
            "displayName": "Notifications",
            "iconUri": "https://psux.azureedge.net/Content/Images/Connectors/FlowNotification.svg",
            "brandColor": "#FF3B30",
            "swagger": {
              "swagger": "2.0",
              "info": {
                "version": "1.0",
                "title": "Notifications",
                "description": "The notification service enables notifications created by a flow to go to your email account or Microsoft Flow mobile app.",
                "contact": {
                  "name": "Samuel L. Banina",
                  "email": "saban@microsoft.com"
                }
              },
              "host": "europe-001.azure-apim.net",
              "basePath": "/apim/flowpush",
              "schemes": [
                "https"
              ],
              "consumes": [
                "application/json"
              ],
              "produces": [
                "application/json"
              ],
              "x-ms-capabilities": {
                "buttons": {
                  "flowIosApp": {},
                  "flowAndroidApp": {}
                }
              },
              "definitions": {
                "NotificationDefinition": {
                  "type": "object",
                  "required": [
                    "notificationText"
                  ],
                  "properties": {
                    "notificationText": {
                      "description": "Create a notification message",
                      "x-ms-summary": "Text",
                      "type": "string"
                    },
                    "notificationLink": {
                      "description": "Custom notification link",
                      "type": "object",
                      "properties": {
                        "uri": {
                          "description": "Include a link in the notification",
                          "x-ms-summary": "Link",
                          "type": "string"
                        },
                        "label": {
                          "description": "The display name for the link",
                          "x-ms-summary": "Link label",
                          "type": "string"
                        }
                      }
                    }
                  }
                },
                "NotificationEmailDefinition": {
                  "type": "object",
                  "required": [
                    "notificationSubject",
                    "notificationBody"
                  ],
                  "properties": {
                    "notificationSubject": {
                      "description": "Notification email subject",
                      "x-ms-summary": "Subject",
                      "type": "string"
                    },
                    "notificationBody": {
                      "description": "Notification email body",
                      "x-ms-summary": "Body",
                      "type": "string"
                    }
                  }
                }
              },
              "paths": {
                "/{connectionId}/sendNotification": {
                  "post": {
                    "description": "Sends a push notification to your Microsoft Flow mobile app.",
                    "summary": "Send me a mobile notification",
                    "operationId": "SendNotification",
                    "x-ms-visibility": "important",
                    "consumes": [
                      "application/json"
                    ],
                    "produces": [
                      "application/json"
                    ],
                    "parameters": [
                      {
                        "name": "connectionId",
                        "in": "path",
                        "required": true,
                        "type": "string",
                        "x-ms-visibility": "internal"
                      },
                      {
                        "name": "NotificationDefinition",
                        "x-ms-summary": "The push notification",
                        "in": "body",
                        "description": "Push notification inputs",
                        "required": true,
                        "schema": {
                          "$ref": "#/definitions/NotificationDefinition"
                        }
                      }
                    ],
                    "responses": {
                      "200": {
                        "description": "OK"
                      },
                      "400": {
                        "description": "Bad Request"
                      },
                      "500": {
                        "description": "Internal Server Error"
                      },
                      "default": {
                        "description": "Operation Failed."
                      }
                    }
                  }
                },
                "/{connectionId}/sendEmailNotification": {
                  "post": {
                    "description": "Sends an email notification to the account you signed in to Microsoft Flow with.",
                    "summary": "Send me an email notification",
                    "operationId": "SendEmailNotification",
                    "x-ms-visibility": "important",
                    "consumes": [
                      "application/json"
                    ],
                    "produces": [
                      "application/json"
                    ],
                    "parameters": [
                      {
                        "name": "connectionId",
                        "in": "path",
                        "required": true,
                        "type": "string",
                        "x-ms-visibility": "internal"
                      },
                      {
                        "name": "NotificationEmailDefinition",
                        "x-ms-summary": "The email notification",
                        "in": "body",
                        "description": "Email notification inputs",
                        "required": true,
                        "schema": {
                          "$ref": "#/definitions/NotificationEmailDefinition"
                        }
                      }
                    ],
                    "responses": {
                      "200": {
                        "description": "OK"
                      },
                      "400": {
                        "description": "Bad Request"
                      },
                      "500": {
                        "description": "Internal Server Error"
                      },
                      "default": {
                        "description": "Operation Failed."
                      }
                    }
                  }
                }
              }
            },
            "tier": "NotSpecified"
          }
        },
        "createdTime": "2018-03-23T17:59:35.4407282Z",
        "lastModifiedTime": "2018-03-23T17:59:37.1164508Z",
        "templateName": "a04de6ce52984b3db0b907f588994bc8",
        "environment": {
          "name": "Default-d87a7535-dd31-4437-bfe1-95340acd55c5",
          "type": "Microsoft.ProcessSimple/environments",
          "id": "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5"
        },
        "definitionSummary": {
          "triggers": [
            {
              "type": "Recurrence"
            }
          ],
          "actions": [
            {
              "type": "If"
            },
            {
              "type": "Query"
            },
            {
              "type": "ApiConnection",
              "swaggerOperationId": "ListFeedItems",
              "metadata": {
                "flowSystemMetadata": {
                  "swaggerOperationId": "ListFeedItems"
                }
              },
              "api": {
                "name": "shared_rss",
                "id": "/providers/Microsoft.PowerApps/apis/shared_rss",
                "type": "/providers/Microsoft.PowerApps/apis"
              }
            },
            {
              "type": "Foreach"
            },
            {
              "type": "ApiConnection",
              "swaggerOperationId": "SendEmailNotification",
              "metadata": {
                "flowSystemMetadata": {
                  "swaggerOperationId": "SendEmailNotification"
                }
              },
              "api": {
                "name": "shared_flowpush",
                "id": "/providers/Microsoft.PowerApps/apis/shared_flowpush",
                "type": "/providers/Microsoft.PowerApps/apis"
              }
            },
            {
              "type": "Compose"
            }
          ],
          "description": "Each day, get an email with a list of all of the top CNN posts from the last day."
        },
        "creator": {
          "tenantId": "d87a7535-dd31-4437-bfe1-95340acd55c5",
          "objectId": "da8f7aea-cf43-497f-ad62-c2feae89a194",
          "userId": "da8f7aea-cf43-497f-ad62-c2feae89a194",
          "userType": "ActiveDirectory"
        },
        "flowTriggerUri": "https://management.azure.com:443/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/3989cb59-ce1a-4a5c-bb78-257c5c39381d/triggers/Every_day/run?api-version=2016-11-01",
        "installationStatus": "Installed",
        "provisioningMethod": "FromDefinition",
        "flowFailureAlertSubscribed": true,
        "referencedResources": []
      }
    };

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/3989cb59-ce1a-4a5c-bb78-257c5c39381d?api-version=2016-11-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve(flowInfo);
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: '3989cb59-ce1a-4a5c-bb78-257c5c39381d', environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5', output: 'json' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(flowInfo));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('renders empty string for description, if no description in the Flow specified', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/3989cb59-ce1a-4a5c-bb78-257c5c39381d?api-version=2016-11-01`) > -1) {
        return Promise.resolve({
          "name": "3989cb59-ce1a-4a5c-bb78-257c5c39381d",
          "id": "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/3989cb59-ce1a-4a5c-bb78-257c5c39381d",
          "type": "Microsoft.ProcessSimple/environments/flows",
          "properties": {
            "apiId": "/providers/Microsoft.PowerApps/apis/shared_logicflows",
            "displayName": "Get a daily digest of the top CNN news",
            "userType": "Owner",
            "definition": {
              "$schema": "https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#",
              "contentVersion": "1.0.0.0",
              "parameters": {
                "$connections": {
                  "defaultValue": {},
                  "type": "Object"
                },
                "$authentication": {
                  "defaultValue": {},
                  "type": "SecureObject"
                }
              },
              "triggers": {
                "Every_day": {
                  "recurrence": {
                    "frequency": "Day",
                    "interval": 1
                  },
                  "type": "Recurrence"
                }
              },
              "actions": {
                "Check_if_there_were_any_posts_this_week": {
                  "actions": {
                    "Compose_the_links_for_each_blog_post": {
                      "foreach": "@take(body('Filter_array'),15)",
                      "actions": {
                        "Compose": {
                          "runAfter": {},
                          "type": "Compose",
                          "inputs": "<tr><td><h3><a href=\"@{item()?['primaryLink']}\">@{item()['title']}</a></h3></td></tr><tr><td style=\"color: #777777;\">Posted at @{formatDateTime(item()?['publishDate'], 't')} GMT</td></tr><tr><td>@{item()?['summary']}</td></tr>"
                        }
                      },
                      "runAfter": {},
                      "type": "Foreach"
                    },
                    "Send_an_email": {
                      "runAfter": {
                        "Compose_the_links_for_each_blog_post": [
                          "Succeeded"
                        ]
                      },
                      "metadata": {
                        "flowSystemMetadata": {
                          "swaggerOperationId": "SendEmailNotification"
                        }
                      },
                      "type": "ApiConnection",
                      "inputs": {
                        "body": {
                          "notificationBody": "<h2>Check out the latest news from CNN:</h2><table>@{join(outputs('Compose'),'')}</table>",
                          "notificationSubject": "Daily digest from CNN top news"
                        },
                        "host": {
                          "api": {
                            "runtimeUrl": "https://europe-001.azure-apim.net/apim/flowpush"
                          },
                          "connection": {
                            "name": "@parameters('$connections')['shared_flowpush']['connectionId']"
                          }
                        },
                        "method": "post",
                        "path": "/sendEmailNotification",
                        "authentication": "@parameters('$authentication')"
                      }
                    }
                  },
                  "runAfter": {
                    "Filter_array": [
                      "Succeeded"
                    ]
                  },
                  "expression": "@greater(length(body('Filter_array')), 0)",
                  "type": "If"
                },
                "Filter_array": {
                  "runAfter": {
                    "List_all_RSS_feed_items": [
                      "Succeeded"
                    ]
                  },
                  "type": "Query",
                  "inputs": {
                    "from": "@body('List_all_RSS_feed_items')",
                    "where": "@greater(item()?['publishDate'], adddays(utcnow(),-2))"
                  }
                },
                "List_all_RSS_feed_items": {
                  "runAfter": {},
                  "metadata": {
                    "flowSystemMetadata": {
                      "swaggerOperationId": "ListFeedItems"
                    }
                  },
                  "type": "ApiConnection",
                  "inputs": {
                    "host": {
                      "api": {
                        "runtimeUrl": "https://europe-001.azure-apim.net/apim/rss"
                      },
                      "connection": {
                        "name": "@parameters('$connections')['shared_rss']['connectionId']"
                      }
                    },
                    "method": "get",
                    "path": "/ListFeedItems",
                    "queries": {
                      "feedUrl": "http://rss.cnn.com/rss/cnn_topstories.rss"
                    },
                    "authentication": "@parameters('$authentication')"
                  }
                }
              },
              "outputs": {},
              "description": "Each day, get an email with a list of all of the top CNN posts from the last day."
            },
            "state": "Started",
            "connectionReferences": {
              "shared_rss": {
                "connectionName": "shared-rss-6636bfd3-0d29-4842-b835-c5910b6310f6",
                "apiDefinition": {
                  "name": "shared_rss",
                  "id": "/providers/Microsoft.PowerApps/apis/shared_rss",
                  "type": "/providers/Microsoft.PowerApps/apis",
                  "properties": {
                    "displayName": "RSS",
                    "iconUri": "https://az818438.vo.msecnd.net/icons/rss.png",
                    "purpose": "NotSpecified",
                    "connectionParameters": {},
                    "scopes": {
                      "will": [],
                      "wont": []
                    },
                    "runtimeUrls": [
                      "https://europe-001.azure-apim.net/apim/rss"
                    ],
                    "primaryRuntimeUrl": "https://europe-001.azure-apim.net/apim/rss",
                    "metadata": {
                      "source": "marketplace",
                      "brandColor": "#ff9900"
                    },
                    "capabilities": [
                      "actions"
                    ],
                    "tier": "Standard",
                    "description": "RSS is a popular web syndication format used to publish frequently updated content – like blog entries and news headlines.  Many content publishers provide an RSS feed to allow users to subscribe to it.  Use the RSS connector to retrieve feed information and trigger flows when new items are published in an RSS feed.",
                    "createdTime": "2016-09-30T04:13:14.2871915Z",
                    "changedTime": "2018-01-17T20:20:37.905252Z"
                  }
                },
                "source": "Embedded",
                "id": "/providers/Microsoft.PowerApps/apis/shared_rss",
                "displayName": "RSS",
                "iconUri": "https://az818438.vo.msecnd.net/icons/rss.png",
                "brandColor": "#ff9900",
                "swagger": {
                  "swagger": "2.0",
                  "info": {
                    "version": "1.0",
                    "title": "RSS",
                    "description": "RSS is a popular web syndication format used to publish frequently updated content – like blog entries and news headlines.  Many content publishers provide an RSS feed to allow users to subscribe to it.  Use the RSS connector to retrieve feed information and trigger flows when new items are published in an RSS feed.",
                    "x-ms-api-annotation": {
                      "status": "Production"
                    }
                  },
                  "host": "europe-001.azure-apim.net",
                  "basePath": "/apim/rss",
                  "schemes": [
                    "https"
                  ],
                  "paths": {
                    "/{connectionId}/OnNewFeed": {
                      "get": {
                        "tags": [
                          "Rss"
                        ],
                        "summary": "When a feed item is published",
                        "description": "This operation triggers a workflow when a new item is published in an RSS feed.",
                        "operationId": "OnNewFeed",
                        "consumes": [],
                        "produces": [
                          "application/json",
                          "text/json",
                          "application/xml",
                          "text/xml"
                        ],
                        "parameters": [
                          {
                            "name": "connectionId",
                            "in": "path",
                            "required": true,
                            "type": "string",
                            "x-ms-visibility": "internal"
                          },
                          {
                            "name": "feedUrl",
                            "in": "query",
                            "description": "The RSS feed URL (Example: http://rss.cnn.com/rss/cnn_topstories.rss).",
                            "required": true,
                            "x-ms-summary": "The RSS feed URL",
                            "x-ms-url-encoding": "double",
                            "type": "string"
                          }
                        ],
                        "responses": {
                          "200": {
                            "description": "OK",
                            "schema": {
                              "$ref": "#/definitions/TriggerBatchResponse[FeedItem]"
                            }
                          },
                          "202": {
                            "description": "Accepted"
                          },
                          "400": {
                            "description": "Bad Request"
                          },
                          "401": {
                            "description": "Unauthorized"
                          },
                          "403": {
                            "description": "Forbidden"
                          },
                          "404": {
                            "description": "Not Found"
                          },
                          "500": {
                            "description": "Internal Server Error. Unknown error occurred"
                          },
                          "default": {
                            "description": "Operation Failed."
                          }
                        },
                        "deprecated": false,
                        "x-ms-visibility": "important",
                        "x-ms-trigger": "batch",
                        "x-ms-trigger-hint": "To see it work now, publish an item to the RSS feed."
                      }
                    },
                    "/{connectionId}/ListFeedItems": {
                      "get": {
                        "tags": [
                          "Rss"
                        ],
                        "summary": "List all RSS feed items",
                        "description": "This operation retrieves all items from an RSS feed.",
                        "operationId": "ListFeedItems",
                        "consumes": [],
                        "produces": [
                          "application/json",
                          "text/json",
                          "application/xml",
                          "text/xml"
                        ],
                        "parameters": [
                          {
                            "name": "connectionId",
                            "in": "path",
                            "required": true,
                            "type": "string",
                            "x-ms-visibility": "internal"
                          },
                          {
                            "name": "feedUrl",
                            "in": "query",
                            "description": "The RSS feed URL (Example: http://rss.cnn.com/rss/cnn_topstories.rss).",
                            "required": true,
                            "x-ms-summary": "The RSS feed URL",
                            "x-ms-url-encoding": "double",
                            "type": "string"
                          }
                        ],
                        "responses": {
                          "200": {
                            "description": "OK",
                            "schema": {
                              "type": "array",
                              "items": {
                                "$ref": "#/definitions/FeedItem"
                              }
                            }
                          },
                          "202": {
                            "description": "Accepted"
                          },
                          "400": {
                            "description": "Bad Request"
                          },
                          "401": {
                            "description": "Unauthorized"
                          },
                          "403": {
                            "description": "Forbidden"
                          },
                          "404": {
                            "description": "Not Found"
                          },
                          "500": {
                            "description": "Internal Server Error. Unknown error occurred"
                          },
                          "default": {
                            "description": "Operation Failed."
                          }
                        },
                        "deprecated": false,
                        "x-ms-visibility": "important"
                      }
                    }
                  },
                  "definitions": {
                    "TriggerBatchResponse[FeedItem]": {
                      "description": "Represents a wrapper object for batch trigger response",
                      "type": "object",
                      "properties": {
                        "value": {
                          "description": "A list of the response objects",
                          "type": "array",
                          "items": {
                            "$ref": "#/definitions/FeedItem"
                          }
                        }
                      }
                    },
                    "FeedItem": {
                      "description": "Represents an RSS feed item.",
                      "required": [
                        "id",
                        "title"
                      ],
                      "type": "object",
                      "properties": {
                        "id": {
                          "description": "Feed ID",
                          "type": "string",
                          "x-ms-summary": "Feed ID"
                        },
                        "title": {
                          "description": "Feed title",
                          "type": "string",
                          "x-ms-summary": "Feed title"
                        },
                        "primaryLink": {
                          "description": "Primary feed link",
                          "type": "string",
                          "x-ms-summary": "Primary feed link"
                        },
                        "links": {
                          "description": "Feed links",
                          "type": "array",
                          "items": {
                            "type": "string"
                          },
                          "x-ms-summary": "Feed links",
                          "x-ms-visibility": "advanced"
                        },
                        "updatedOn": {
                          "description": "Feed updated on",
                          "type": "string",
                          "x-ms-summary": "Feed updated on"
                        },
                        "publishDate": {
                          "description": "Feed published date",
                          "type": "string",
                          "x-ms-summary": "Feed published on"
                        },
                        "summary": {
                          "description": "Feed item summary",
                          "type": "string",
                          "x-ms-summary": "Feed summary"
                        },
                        "copyright": {
                          "description": "Copyright information",
                          "type": "string",
                          "x-ms-summary": "Feed copyright information"
                        },
                        "categories": {
                          "description": "Feed item categories",
                          "type": "array",
                          "items": {
                            "type": "string"
                          },
                          "x-ms-summary": "Feed categories"
                        }
                      }
                    }
                  }
                },
                "tier": "NotSpecified"
              },
              "shared_flowpush": {
                "connectionName": "shared-flowpush-295e4b80-1a4e-42ec-aa5b-8d72e7c1eb3f",
                "apiDefinition": {
                  "name": "shared_flowpush",
                  "id": "/providers/Microsoft.PowerApps/apis/shared_flowpush",
                  "type": "/providers/Microsoft.PowerApps/apis",
                  "properties": {
                    "displayName": "Notifications",
                    "iconUri": "https://psux.azureedge.net/Content/Images/Connectors/FlowNotification.svg",
                    "purpose": "NotSpecified",
                    "connectionParameters": {},
                    "runtimeUrls": [
                      "https://europe-001.azure-apim.net/apim/flowpush"
                    ],
                    "primaryRuntimeUrl": "https://europe-001.azure-apim.net/apim/flowpush",
                    "metadata": {
                      "source": "marketplace",
                      "brandColor": "#FF3B30",
                      "connectionLimits": {
                        "*": 1
                      }
                    },
                    "capabilities": [
                      "actions"
                    ],
                    "tier": "NotSpecified",
                    "description": "The notification service enables notifications created by a flow to go to your email account or Microsoft Flow mobile app.",
                    "createdTime": "2016-10-22T06:43:26.1572419Z",
                    "changedTime": "2018-01-25T00:34:52.048009Z"
                  }
                },
                "source": "Embedded",
                "id": "/providers/Microsoft.PowerApps/apis/shared_flowpush",
                "displayName": "Notifications",
                "iconUri": "https://psux.azureedge.net/Content/Images/Connectors/FlowNotification.svg",
                "brandColor": "#FF3B30",
                "swagger": {
                  "swagger": "2.0",
                  "info": {
                    "version": "1.0",
                    "title": "Notifications",
                    "description": "The notification service enables notifications created by a flow to go to your email account or Microsoft Flow mobile app.",
                    "contact": {
                      "name": "Samuel L. Banina",
                      "email": "saban@microsoft.com"
                    }
                  },
                  "host": "europe-001.azure-apim.net",
                  "basePath": "/apim/flowpush",
                  "schemes": [
                    "https"
                  ],
                  "consumes": [
                    "application/json"
                  ],
                  "produces": [
                    "application/json"
                  ],
                  "x-ms-capabilities": {
                    "buttons": {
                      "flowIosApp": {},
                      "flowAndroidApp": {}
                    }
                  },
                  "definitions": {
                    "NotificationDefinition": {
                      "type": "object",
                      "required": [
                        "notificationText"
                      ],
                      "properties": {
                        "notificationText": {
                          "description": "Create a notification message",
                          "x-ms-summary": "Text",
                          "type": "string"
                        },
                        "notificationLink": {
                          "description": "Custom notification link",
                          "type": "object",
                          "properties": {
                            "uri": {
                              "description": "Include a link in the notification",
                              "x-ms-summary": "Link",
                              "type": "string"
                            },
                            "label": {
                              "description": "The display name for the link",
                              "x-ms-summary": "Link label",
                              "type": "string"
                            }
                          }
                        }
                      }
                    },
                    "NotificationEmailDefinition": {
                      "type": "object",
                      "required": [
                        "notificationSubject",
                        "notificationBody"
                      ],
                      "properties": {
                        "notificationSubject": {
                          "description": "Notification email subject",
                          "x-ms-summary": "Subject",
                          "type": "string"
                        },
                        "notificationBody": {
                          "description": "Notification email body",
                          "x-ms-summary": "Body",
                          "type": "string"
                        }
                      }
                    }
                  },
                  "paths": {
                    "/{connectionId}/sendNotification": {
                      "post": {
                        "description": "Sends a push notification to your Microsoft Flow mobile app.",
                        "summary": "Send me a mobile notification",
                        "operationId": "SendNotification",
                        "x-ms-visibility": "important",
                        "consumes": [
                          "application/json"
                        ],
                        "produces": [
                          "application/json"
                        ],
                        "parameters": [
                          {
                            "name": "connectionId",
                            "in": "path",
                            "required": true,
                            "type": "string",
                            "x-ms-visibility": "internal"
                          },
                          {
                            "name": "NotificationDefinition",
                            "x-ms-summary": "The push notification",
                            "in": "body",
                            "description": "Push notification inputs",
                            "required": true,
                            "schema": {
                              "$ref": "#/definitions/NotificationDefinition"
                            }
                          }
                        ],
                        "responses": {
                          "200": {
                            "description": "OK"
                          },
                          "400": {
                            "description": "Bad Request"
                          },
                          "500": {
                            "description": "Internal Server Error"
                          },
                          "default": {
                            "description": "Operation Failed."
                          }
                        }
                      }
                    },
                    "/{connectionId}/sendEmailNotification": {
                      "post": {
                        "description": "Sends an email notification to the account you signed in to Microsoft Flow with.",
                        "summary": "Send me an email notification",
                        "operationId": "SendEmailNotification",
                        "x-ms-visibility": "important",
                        "consumes": [
                          "application/json"
                        ],
                        "produces": [
                          "application/json"
                        ],
                        "parameters": [
                          {
                            "name": "connectionId",
                            "in": "path",
                            "required": true,
                            "type": "string",
                            "x-ms-visibility": "internal"
                          },
                          {
                            "name": "NotificationEmailDefinition",
                            "x-ms-summary": "The email notification",
                            "in": "body",
                            "description": "Email notification inputs",
                            "required": true,
                            "schema": {
                              "$ref": "#/definitions/NotificationEmailDefinition"
                            }
                          }
                        ],
                        "responses": {
                          "200": {
                            "description": "OK"
                          },
                          "400": {
                            "description": "Bad Request"
                          },
                          "500": {
                            "description": "Internal Server Error"
                          },
                          "default": {
                            "description": "Operation Failed."
                          }
                        }
                      }
                    }
                  }
                },
                "tier": "NotSpecified"
              }
            },
            "createdTime": "2018-03-23T17:59:35.4407282Z",
            "lastModifiedTime": "2018-03-23T17:59:37.1164508Z",
            "templateName": "a04de6ce52984b3db0b907f588994bc8",
            "environment": {
              "name": "Default-d87a7535-dd31-4437-bfe1-95340acd55c5",
              "type": "Microsoft.ProcessSimple/environments",
              "id": "/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5"
            },
            "definitionSummary": {
              "triggers": [
                {
                  "type": "Recurrence"
                }
              ],
              "actions": [
                {
                  "type": "If"
                },
                {
                  "type": "Query"
                },
                {
                  "type": "ApiConnection",
                  "swaggerOperationId": "ListFeedItems",
                  "metadata": {
                    "flowSystemMetadata": {
                      "swaggerOperationId": "ListFeedItems"
                    }
                  },
                  "api": {
                    "name": "shared_rss",
                    "id": "/providers/Microsoft.PowerApps/apis/shared_rss",
                    "type": "/providers/Microsoft.PowerApps/apis"
                  }
                },
                {
                  "type": "Foreach"
                },
                {
                  "type": "ApiConnection",
                  "swaggerOperationId": "SendEmailNotification",
                  "metadata": {
                    "flowSystemMetadata": {
                      "swaggerOperationId": "SendEmailNotification"
                    }
                  },
                  "api": {
                    "name": "shared_flowpush",
                    "id": "/providers/Microsoft.PowerApps/apis/shared_flowpush",
                    "type": "/providers/Microsoft.PowerApps/apis"
                  }
                },
                {
                  "type": "Compose"
                }
              ]
            },
            "creator": {
              "tenantId": "d87a7535-dd31-4437-bfe1-95340acd55c5",
              "objectId": "da8f7aea-cf43-497f-ad62-c2feae89a194",
              "userId": "da8f7aea-cf43-497f-ad62-c2feae89a194",
              "userType": "ActiveDirectory"
            },
            "flowTriggerUri": "https://management.azure.com:443/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/3989cb59-ce1a-4a5c-bb78-257c5c39381d/triggers/Every_day/run?api-version=2016-11-01",
            "installationStatus": "Installed",
            "provisioningMethod": "FromDefinition",
            "flowFailureAlertSubscribed": true,
            "referencedResources": []
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: '3989cb59-ce1a-4a5c-bb78-257c5c39381d', environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          name: '3989cb59-ce1a-4a5c-bb78-257c5c39381d',
          displayName: 'Get a daily digest of the top CNN news',
          description: '',
          triggers: 'Every_day',
          actions: 'Check_if_there_were_any_posts_this_week, Filter_array, List_all_RSS_feed_items'
        }));
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
          "message": "Access to the environment 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' is denied."
        }
      });
    });

    cmdInstance.action({ options: { debug: false, environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6', name: '3989cb59-ce1a-4a5c-bb78-257c5c39381d' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Access to the environment 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' is denied.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles Flow not found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject({
        "error": {
          "code": "ConnectionAuthorizationFailed",
          "message": "The caller with object id 'da8f7aea-cf43-497f-ad62-c2feae89a194' does not have permission for connection '1c6ee23a-a835-44bc-a4f5-462b658efc12' under Api 'shared_logicflows'."
        }
      });
    });

    cmdInstance.action({ options: { debug: false, environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6', name: '1c6ee23a-a835-44bc-a4f5-462b658efc12' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The caller with object id 'da8f7aea-cf43-497f-ad62-c2feae89a194' does not have permission for connection '1c6ee23a-a835-44bc-a4f5-462b658efc12' under Api 'shared_logicflows'.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles Flow not found (as admin)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject({
        "error": {
          "code": "FlowNotFound",
          "message": "Could not find flow '1c6ee23a-a835-44bc-a4f5-462b658efc12'."
        }
      });
    });

    cmdInstance.action({ options: { debug: false, environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6', name: '1c6ee23a-a835-44bc-a4f5-462b658efc12', asAdmin: true } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Could not find flow '1c6ee23a-a835-44bc-a4f5-462b658efc12'.`)));
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

    cmdInstance.action({ options: { debug: false, environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5', name: '3989cb59-ce1a-4a5c-bb78-257c5c39381d' } }, (err?: any) => {
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

  it('supports specifying name', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--name') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying environment', () => {
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