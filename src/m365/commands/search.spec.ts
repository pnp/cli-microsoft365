import assert from 'assert';
import sinon from 'sinon';
import auth from '../../Auth.js';
import { cli } from '../../cli/cli.js';
import { CommandInfo } from '../../cli/CommandInfo.js';
import { Logger } from '../../cli/Logger.js';
import { CommandError } from '../../Command.js';
import request from '../../request.js';
import { telemetry } from '../../telemetry.js';
import { misc } from '../../utils/misc.js';
import { MockRequests } from '../../utils/MockRequest.js';
import { pid } from '../../utils/pid.js';
import { session } from '../../utils/session.js';
import { sinonUtil } from '../../utils/sinonUtil.js';
import commands from './commands.js';
import command from './search.js';

const fullSearchResponse = {
  "searchTerms": [
    "contoso"
  ],
  "hitsContainers": [
    {
      "hits": [
        {
          "hitId": "AAMkAGRlM2Y5YTkzLWI2NzAtNDczOS05YWMyLTJhZGY2MGExMGU0MgBGAAAAAABIbfA8TbuRR7JKOZPl5FPxBwB8kpUvTuxuSYh8eqNsOdGBAAAAAAEMAAB8kpUvTuxuSYh8eqNsOdGBAAD30FHPAAA=",
          "rank": 1,
          "summary": "...More.  Your weekly PIM <c0>digest</c0> for MSFT Thanks for using Microsoft Entra Privileged Identity Management (PIM). This weekly <c0>digest</c0> shows your PIM activities over the last seven days:...",
          "resource": {
            "@odata.type": "#microsoft.graph.message",
            "createdDateTime": "2024-03-03T02:27:56Z",
            "lastModifiedDateTime": "2024-03-03T02:27:58Z",
            "receivedDateTime": "2024-03-03T02:27:56Z",
            "sentDateTime": "2024-03-03T02:27:51Z",
            "hasAttachments": false,
            "internetMessageId": "<Q5UVB9PQFMU4.BU5IUGF1B0I93@contoso>",
            "subject": "Your weekly PIM digest for MSFT",
            "bodyPreview": "...More.  Your weekly PIM digest for MSFT Thanks for using Microsoft Entra Privileged Identity Management (PIM). This weekly digest shows your PIM activities over the last seven days:...",
            "importance": "normal",
            "parentFolderId": "AQMkAGRlM2Y5YTkzLWI2NzAtNDczOS05YWMyLTJhZGY2MGExMGU0MgAuAAADSG3wPE27kUeySjmT5eRT8QEAfJKVL07sbkmIfHqjbDnRgQAAAgEMAAAA",
            "conversationId": "AAQkAGRlM2Y5YTkzLWI2NzAtNDczOS05YWMyLTJhZGY2MGExMGU0MgAQAMc0n1gdmdxAgAfGlWKhSm4=",
            "isRead": false,
            "isDraft": false,
            "webLink": "https://outlook.office365.com/owa/?ItemID=AAMkAGRlM2Y5YTkzLWI2NzAtNDczOS05YWMyLTJhZGY2MGExMGU0MgBGAAAAAABIbfA8TbuRR7JKOZPl5FPxBwB8kpUvTuxuSYh8eqNsOdGBAAAAAAEMAAB8kpUvTuxuSYh8eqNsOdGBAAD30FHPAAA%3D&exvsurl=1&viewmodel=ReadMessageItem",
            "inferenceClassification": "focused",
            "replyTo": [
              {
                "emailAddress": {
                  "name": "John Doe"
                }
              }
            ],
            "sender": {
              "emailAddress": {
                "name": "Microsoft Security",
                "address": "MSSecurity-noreply@microsoft.com"
              }
            },
            "from": {
              "emailAddress": {
                "name": "Microsoft Security",
                "address": "MSSecurity-noreply@microsoft.com"
              }
            }
          }
        }
      ],
      "total": 350,
      "moreResultsAvailable": true
    }
  ]
};
const selectedPropertiesSearchResponse = {
  "searchTerms": [
    "contoso"
  ],
  "hitsContainers": [
    {
      "hits": [
        {
          "hitId": "AAMkAGRlM2Y5YTkzLWI2NzAtNDczOS05YWMyLTJhZGY2MGExMGU0MgBGAAAAAABIbfA8TbuRR7JKOZPl5FPxBwB8kpUvTuxuSYh8eqNsOdGBAAAAAAEMAAB8kpUvTuxuSYh8eqNsOdGBAAD30FHPAAA=",
          "rank": 1,
          "summary": "...More.  Your weekly PIM <c0>digest</c0> for MSFT Thanks for using Microsoft Entra Privileged Identity Management (PIM). This weekly <c0>digest</c0> shows your PIM activities over the last seven days:...",
          "resource": {
            "@odata.type": "#microsoft.graph.message",
            "subject": "Your weekly PIM digest for MSFT",
            "importance": "normal"
          }
        }
      ],
      "total": 350,
      "moreResultsAvailable": true
    }
  ]
};
const spellingCorrectionSearchResponse = {
  "searchTerms": [
    "principles"
  ],
  "hitsContainers": [
    {
      "hits": [
        {
          "hitId": "01IOTMML4Q6XCI2NYRXBALIRDODNYA6XLK",
          "rank": 1,
          "summary": "",
          "resource": {
            "@odata.type": "#microsoft.graph.driveItem",
            "size": 0,
            "fileSystemInfo": {
              "createdDateTime": "2023-02-22T06:41:20Z",
              "lastModifiedDateTime": "2023-02-22T06:41:20Z"
            },
            "listItem": {
              "@odata.type": "#microsoft.graph.listItem",
              "fields": {},
              "id": "8dc4f590-1137-40b8-b444-6e1b700f5d6a"
            },
            "id": "01IOTMML4Q6XCI2NYRXBALIRDODNYA6XLK",
            "createdBy": {
              "user": {
                "displayName": "SharePoint App"
              }
            },
            "createdDateTime": "2023-02-22T06:41:20Z",
            "lastModifiedBy": {
              "user": {
                "displayName": "SharePoint App"
              }
            },
            "lastModifiedDateTime": "2023-02-22T06:41:20Z",
            "name": "Contoso Marketing Principles.pptx",
            "parentReference": {
              "driveId": "b!F3G46XHCqU6L-OwKUiU8UObjtItbdTVCk4XtFGm7LW99GmedI39YRpyWEhUHE3Sn",
              "id": "01IOTMML7VRF52SG4UKVBIE6HPXFKUI6U3",
              "sharepointIds": {
                "listId": "9d671a7d-7f23-4658-9c96-1215071374a7",
                "listItemId": "8",
                "listItemUniqueId": "8dc4f590-1137-40b8-b444-6e1b700f5d6a"
              },
              "siteId": "contoso.sharepoint.com,e9b87117-c271-4ea9-8bf8-ec0a52253c50,8bb4e3e6-755b-4235-9385-ed1469bb2d6f"
            },
            "webUrl": "https://contoso.sharepoint.com/sites/Mark8ProjectTeam2/Shared Documents/Digital Assets Web/Contoso Marketing Principles.pptx"
          }
        }
      ],
      "total": 1,
      "moreResultsAvailable": false
    }
  ],
  "queryAlterationResponse": {
    "originalQueryString": "principless",
    "queryAlteration": {
      "alteredQueryString": "principles",
      "alteredHighlightedQueryString": "principles",
      "alteredQueryTokens": [
        {
          "offset": 0,
          "length": 11,
          "suggestion": "principles"
        }
      ]
    },
    "queryAlterationType": "modification"
  }
};

export const mocks = {
  queryNotSpecified: {
    request: {
      url: 'https://graph.microsoft.com/v1.0/search/query',
      method: 'POST',
      bodyFragment: '"queryString": "*"'
    },
    response: {
      body: { "value": [fullSearchResponse] }
    }
  },
  querySpecified: {
    request: {
      url: 'https://graph.microsoft.com/v1.0/search/query',
      method: 'POST',
      bodyFragment: '"queryString": "contoso"'
    },
    response: {
      body: { "value": [fullSearchResponse] }
    }
  },
  size: {
    request: {
      url: 'https://graph.microsoft.com/v1.0/search/query',
      method: 'POST',
      bodyFragment: '"size": 1,'
    },
    response: {
      body: { "value": [fullSearchResponse] }
    }
  },
  from: {
    request: {
      url: 'https://graph.microsoft.com/v1.0/search/query',
      method: 'POST',
      bodyFragment: '"from": 10,'
    },
    response: {
      body: { "value": [fullSearchResponse] }
    }
  },
  fields: {
    request: {
      url: 'https://graph.microsoft.com/v1.0/search/query',
      method: 'POST',
      bodyFragment: '"fields":'
    },
    response: {
      body: { "value": [selectedPropertiesSearchResponse] }
    }
  },
  sort: {
    request: {
      url: 'https://graph.microsoft.com/v1.0/search/query',
      method: 'POST',
      bodyFragment: '"sortProperties": ['
    },
    response: {
      body: { "value": [selectedPropertiesSearchResponse] }
    }
  },
  spelling: {
    request: {
      url: 'https://graph.microsoft.com/v1.0/search/query',
      method: 'POST',
      bodyFragment: '"enableModification": true,'
    },
    response: {
      body: { "value": [spellingCorrectionSearchResponse] }
    }
  }
} satisfies MockRequests;

describe(commands.SEARCH, () => {
  const fullSearchFromIndexResponse = {
    "searchTerms": [
      "contoso"
    ],
    "hitsContainers": [
      {
        "hits": [
          {
            "hitId": "AAMkAGRlM2Y5YTkzLWI2NzAtNDczOS05YWMyLTJhZGY2MGExMGU0MgBGAAAAAABIbfA8TbuRR7JKOZPl5FPxBwB8kpUvTuxuSYh8eqNsOdGBAAAAAAEMAAB8kpUvTuxuSYh8eqNsOdGBAAD30FHPAAA=",
            "rank": 1,
            "summary": "...More.  Your weekly PIM <c0>digest</c0> for MSFT Thanks for using Microsoft Entra Privileged Identity Management (PIM). This weekly <c0>digest</c0> shows your PIM activities over the last seven days:...",
            "resource": {
              "@odata.type": "#microsoft.graph.message",
              "createdDateTime": "2024-03-03T02:27:56Z",
              "lastModifiedDateTime": "2024-03-03T02:27:58Z",
              "receivedDateTime": "2024-03-03T02:27:56Z",
              "sentDateTime": "2024-03-03T02:27:51Z",
              "hasAttachments": false,
              "internetMessageId": "<Q5UVB9PQFMU4.BU5IUGF1B0I93@contoso>",
              "subject": "Your weekly PIM digest for MSFT",
              "bodyPreview": "...More.  Your weekly PIM digest for MSFT Thanks for using Microsoft Entra Privileged Identity Management (PIM). This weekly digest shows your PIM activities over the last seven days:...",
              "importance": "normal",
              "parentFolderId": "AQMkAGRlM2Y5YTkzLWI2NzAtNDczOS05YWMyLTJhZGY2MGExMGU0MgAuAAADSG3wPE27kUeySjmT5eRT8QEAfJKVL07sbkmIfHqjbDnRgQAAAgEMAAAA",
              "conversationId": "AAQkAGRlM2Y5YTkzLWI2NzAtNDczOS05YWMyLTJhZGY2MGExMGU0MgAQAMc0n1gdmdxAgAfGlWKhSm4=",
              "isRead": false,
              "isDraft": false,
              "webLink": "https://outlook.office365.com/owa/?ItemID=AAMkAGRlM2Y5YTkzLWI2NzAtNDczOS05YWMyLTJhZGY2MGExMGU0MgBGAAAAAABIbfA8TbuRR7JKOZPl5FPxBwB8kpUvTuxuSYh8eqNsOdGBAAAAAAEMAAB8kpUvTuxuSYh8eqNsOdGBAAD30FHPAAA%3D&exvsurl=1&viewmodel=ReadMessageItem",
              "inferenceClassification": "focused",
              "replyTo": [
                {
                  "emailAddress": {
                    "name": "John Doe"
                  }
                }
              ],
              "sender": {
                "emailAddress": {
                  "name": "Microsoft Security",
                  "address": "MSSecurity-noreply@microsoft.com"
                }
              },
              "from": {
                "emailAddress": {
                  "name": "Microsoft Security",
                  "address": "MSSecurity-noreply@microsoft.com"
                }
              }
            }
          }
        ],
        "total": 350,
        "moreResultsAvailable": false
      }
    ]
  };

  const allResults = {
    "searchTerms": [
      "contoso"
    ],
    "hitsContainers": [
      {
        "hits": [
          {
            "hitId": "AAMkAGRlM2Y5YTkzLWI2NzAtNDczOS05YWMyLTJhZGY2MGExMGU0MgBGAAAAAABIbfA8TbuRR7JKOZPl5FPxBwB8kpUvTuxuSYh8eqNsOdGBAAAAAAEMAAB8kpUvTuxuSYh8eqNsOdGBAAD30FHPAAA=",
            "rank": 1,
            "summary": "...More.  Your weekly PIM <c0>digest</c0> for MSFT Thanks for using Microsoft Entra Privileged Identity Management (PIM). This weekly <c0>digest</c0> shows your PIM activities over the last seven days:...",
            "resource": {
              "@odata.type": "#microsoft.graph.message",
              "createdDateTime": "2024-03-03T02:27:56Z",
              "lastModifiedDateTime": "2024-03-03T02:27:58Z",
              "receivedDateTime": "2024-03-03T02:27:56Z",
              "sentDateTime": "2024-03-03T02:27:51Z",
              "hasAttachments": false,
              "internetMessageId": "<Q5UVB9PQFMU4.BU5IUGF1B0I93@contoso>",
              "subject": "Your weekly PIM digest for MSFT",
              "bodyPreview": "...More.  Your weekly PIM digest for MSFT Thanks for using Microsoft Entra Privileged Identity Management (PIM). This weekly digest shows your PIM activities over the last seven days:...",
              "importance": "normal",
              "parentFolderId": "AQMkAGRlM2Y5YTkzLWI2NzAtNDczOS05YWMyLTJhZGY2MGExMGU0MgAuAAADSG3wPE27kUeySjmT5eRT8QEAfJKVL07sbkmIfHqjbDnRgQAAAgEMAAAA",
              "conversationId": "AAQkAGRlM2Y5YTkzLWI2NzAtNDczOS05YWMyLTJhZGY2MGExMGU0MgAQAMc0n1gdmdxAgAfGlWKhSm4=",
              "isRead": false,
              "isDraft": false,
              "webLink": "https://outlook.office365.com/owa/?ItemID=AAMkAGRlM2Y5YTkzLWI2NzAtNDczOS05YWMyLTJhZGY2MGExMGU0MgBGAAAAAABIbfA8TbuRR7JKOZPl5FPxBwB8kpUvTuxuSYh8eqNsOdGBAAAAAAEMAAB8kpUvTuxuSYh8eqNsOdGBAAD30FHPAAA%3D&exvsurl=1&viewmodel=ReadMessageItem",
              "inferenceClassification": "focused",
              "replyTo": [
                {
                  "emailAddress": {
                    "name": "John Doe"
                  }
                }
              ],
              "sender": {
                "emailAddress": {
                  "name": "Microsoft Security",
                  "address": "MSSecurity-noreply@microsoft.com"
                }
              },
              "from": {
                "emailAddress": {
                  "name": "Microsoft Security",
                  "address": "MSSecurity-noreply@microsoft.com"
                }
              }
            }
          },
          {
            "hitId": "AAMkAGRlM2Y5YTkzLWI2NzAtNDczOS05YWMyLTJhZGY2MGExMGU0MgBGAAAAAABIbfA8TbuRR7JKOZPl5FPxBwB8kpUvTuxuSYh8eqNsOdGBAAAAAAEMAAB8kpUvTuxuSYh8eqNsOdGBAAD30FHPAAA=",
            "rank": 1,
            "summary": "...More.  Your weekly PIM <c0>digest</c0> for MSFT Thanks for using Microsoft Entra Privileged Identity Management (PIM). This weekly <c0>digest</c0> shows your PIM activities over the last seven days:...",
            "resource": {
              "@odata.type": "#microsoft.graph.message",
              "createdDateTime": "2024-03-03T02:27:56Z",
              "lastModifiedDateTime": "2024-03-03T02:27:58Z",
              "receivedDateTime": "2024-03-03T02:27:56Z",
              "sentDateTime": "2024-03-03T02:27:51Z",
              "hasAttachments": false,
              "internetMessageId": "<Q5UVB9PQFMU4.BU5IUGF1B0I93@contoso>",
              "subject": "Your weekly PIM digest for MSFT",
              "bodyPreview": "...More.  Your weekly PIM digest for MSFT Thanks for using Microsoft Entra Privileged Identity Management (PIM). This weekly digest shows your PIM activities over the last seven days:...",
              "importance": "normal",
              "parentFolderId": "AQMkAGRlM2Y5YTkzLWI2NzAtNDczOS05YWMyLTJhZGY2MGExMGU0MgAuAAADSG3wPE27kUeySjmT5eRT8QEAfJKVL07sbkmIfHqjbDnRgQAAAgEMAAAA",
              "conversationId": "AAQkAGRlM2Y5YTkzLWI2NzAtNDczOS05YWMyLTJhZGY2MGExMGU0MgAQAMc0n1gdmdxAgAfGlWKhSm4=",
              "isRead": false,
              "isDraft": false,
              "webLink": "https://outlook.office365.com/owa/?ItemID=AAMkAGRlM2Y5YTkzLWI2NzAtNDczOS05YWMyLTJhZGY2MGExMGU0MgBGAAAAAABIbfA8TbuRR7JKOZPl5FPxBwB8kpUvTuxuSYh8eqNsOdGBAAAAAAEMAAB8kpUvTuxuSYh8eqNsOdGBAAD30FHPAAA%3D&exvsurl=1&viewmodel=ReadMessageItem",
              "inferenceClassification": "focused",
              "replyTo": [
                {
                  "emailAddress": {
                    "name": "John Doe"
                  }
                }
              ],
              "sender": {
                "emailAddress": {
                  "name": "Microsoft Security",
                  "address": "MSSecurity-noreply@microsoft.com"
                }
              },
              "from": {
                "emailAddress": {
                  "name": "Microsoft Security",
                  "address": "MSSecurity-noreply@microsoft.com"
                }
              }
            }
          }
        ],
        "total": 350,
        "moreResultsAvailable": false
      }
    ]
  };

  const resultsOnlySearchResponse = [
    {
      "hitId": "AAMkAGRlM2Y5YTkzLWI2NzAtNDczOS05YWMyLTJhZGY2MGExMGU0MgBGAAAAAABIbfA8TbuRR7JKOZPl5FPxBwB8kpUvTuxuSYh8eqNsOdGBAAAAAAEMAAB8kpUvTuxuSYh8eqNsOdGBAAD30FHPAAA=",
      "rank": 1,
      "summary": "...More.  Your weekly PIM <c0>digest</c0> for MSFT Thanks for using Microsoft Entra Privileged Identity Management (PIM). This weekly <c0>digest</c0> shows your PIM activities over the last seven days:...",
      "resource": {
        "@odata.type": "#microsoft.graph.message",
        "createdDateTime": "2024-03-03T02:27:56Z",
        "lastModifiedDateTime": "2024-03-03T02:27:58Z",
        "receivedDateTime": "2024-03-03T02:27:56Z",
        "sentDateTime": "2024-03-03T02:27:51Z",
        "hasAttachments": false,
        "internetMessageId": "<Q5UVB9PQFMU4.BU5IUGF1B0I93@contoso>",
        "subject": "Your weekly PIM digest for MSFT",
        "bodyPreview": "...More.  Your weekly PIM digest for MSFT Thanks for using Microsoft Entra Privileged Identity Management (PIM). This weekly digest shows your PIM activities over the last seven days:...",
        "importance": "normal",
        "parentFolderId": "AQMkAGRlM2Y5YTkzLWI2NzAtNDczOS05YWMyLTJhZGY2MGExMGU0MgAuAAADSG3wPE27kUeySjmT5eRT8QEAfJKVL07sbkmIfHqjbDnRgQAAAgEMAAAA",
        "conversationId": "AAQkAGRlM2Y5YTkzLWI2NzAtNDczOS05YWMyLTJhZGY2MGExMGU0MgAQAMc0n1gdmdxAgAfGlWKhSm4=",
        "isRead": false,
        "isDraft": false,
        "webLink": "https://outlook.office365.com/owa/?ItemID=AAMkAGRlM2Y5YTkzLWI2NzAtNDczOS05YWMyLTJhZGY2MGExMGU0MgBGAAAAAABIbfA8TbuRR7JKOZPl5FPxBwB8kpUvTuxuSYh8eqNsOdGBAAAAAAEMAAB8kpUvTuxuSYh8eqNsOdGBAAD30FHPAAA%3D&exvsurl=1&viewmodel=ReadMessageItem",
        "inferenceClassification": "focused",
        "replyTo": [
          {
            "emailAddress": {
              "name": "John Doe"
            }
          }
        ],
        "sender": {
          "emailAddress": {
            "name": "Microsoft Security",
            "address": "MSSecurity-noreply@microsoft.com"
          }
        },
        "from": {
          "emailAddress": {
            "name": "Microsoft Security",
            "address": "MSSecurity-noreply@microsoft.com"
          }
        }
      }
    }
  ];

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
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
    (command as any).pollingInterval = 0;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SEARCH);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation if scopes contains allowed values', async () => {
    const actual = await command.validate({
      options: {
        scopes: 'chatMessage,message,event,drive,driveItem,list,listItem,site,bookmark,acronym,person'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if startIndex equals 0', async () => {
    const actual = await command.validate({
      options: {
        scopes: 'message',
        startIndex: 0
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if startIndex is greater than 0', async () => {
    const actual = await command.validate({
      options: {
        scopes: 'message',
        startIndex: 50
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if pageSize is in allowed range', async () => {
    const actual = await command.validate({
      options: {
        scopes: 'message',
        pageSize: 50
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if enableTopResults is specified for message scope', async () => {
    const actual = await command.validate({
      options: {
        scopes: 'message',
        enableTopResults: true
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if enableTopResults is specified for chatMessage scope', async () => {
    const actual = await command.validate({
      options: {
        scopes: 'chatMessage',
        enableTopResults: true
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if enableTopResults is specified for message and chatMessage scope', async () => {
    const actual = await command.validate({
      options: {
        scopes: 'message, chatMessage',
        enableTopResults: true
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if scopes contains not allowed value', async () => {
    const actual = await command.validate({
      options: {
        scopes: 'chatMessage,message,event,drive,driveItem,list,listItem,site,bookmarks,acronyms,person,foo'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if startIndex is less than 0', async () => {
    const actual = await command.validate({
      options: {
        scopes: 'message',
        startIndex: -1
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if pageSize is less than 1', async () => {
    const actual = await command.validate({
      options: {
        scopes: 'message',
        pageSize: -1
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if pageSize is greater than 500', async () => {
    const actual = await command.validate({
      options: {
        scopes: 'message',
        pageSize: 501
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if enableTopResults is specified together with scope other than message or chatMessage', async () => {
    const actual = await command.validate({
      options: {
        scopes: 'event',
        enableTopResults: true
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if enableTopResults is specified together with multiple scopes, but only one is message', async () => {
    const actual = await command.validate({
      options: {
        scopes: 'message,event',
        enableTopResults: true
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if enableTopResults is specified together with multiple scopes, but only one is chatMessage', async () => {
    const actual = await command.validate({
      options: {
        scopes: 'chatMessage,event',
        enableTopResults: true
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if sortBy is specified for message scope', async () => {
    const actual = await command.validate({
      options: {
        scopes: 'message',
        sortBy: 'subject'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if sortBy is specified for event scope', async () => {
    const actual = await command.validate({
      options: {
        scopes: 'event',
        sortBy: 'subject'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('successfully returns search response when queryText is not specified', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === mocks.queryNotSpecified.request.url &&
        JSON.stringify(opts.data) === JSON.stringify({
          "requests": [
            {
              "entityTypes": [
                "message"
              ],
              "query": {
                "queryString": "*"
              },
              "from": 0,
              "size": 25,
              "queryAlterationOptions": {}
            }
          ]
        })) {
        return misc.deepClone(mocks.queryNotSpecified.response.body);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { scopes: 'message' } });
    assert(loggerLogSpy.calledOnceWithExactly(fullSearchResponse));
  });

  it('successfully returns search response for specified query', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === mocks.querySpecified.request.url &&
        JSON.stringify(opts.data) === JSON.stringify({
          "requests": [
            {
              "entityTypes": [
                "message"
              ],
              "query": {
                "queryString": "contoso"
              },
              "from": 0,
              "size": 25,
              "queryAlterationOptions": {}
            }
          ]
        })) {
        return misc.deepClone(mocks.querySpecified.response.body);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { scopes: 'message', queryText: 'contoso' } });
    assert(loggerLogSpy.calledOnceWithExactly(fullSearchResponse));
  });

  it('successfully returns only search results', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === mocks.querySpecified.request.url &&
        JSON.stringify(opts.data) === JSON.stringify({
          "requests": [
            {
              "entityTypes": [
                "message"
              ],
              "query": {
                "queryString": "contoso"
              },
              "from": 0,
              "size": 25,
              "queryAlterationOptions": {}
            }
          ]
        })) {
        return misc.deepClone(mocks.querySpecified.response.body);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { scopes: 'message', queryText: 'contoso', resultsOnly: true } });
    assert(loggerLogSpy.calledOnceWithExactly(resultsOnlySearchResponse));
  });

  it('successfully returns only specified number of results in search response', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === mocks.size.request.url &&
        JSON.stringify(opts.data) === JSON.stringify({
          "requests": [
            {
              "entityTypes": [
                "message"
              ],
              "query": {
                "queryString": "contoso"
              },
              "from": 0,
              "size": 1,
              "queryAlterationOptions": {}
            }
          ]
        })) {
        return misc.deepClone(mocks.size.response.body);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { scopes: 'message', queryText: 'contoso', pageSize: 1 } });
    assert(loggerLogSpy.calledOnceWithExactly(fullSearchResponse));
  });

  it('successfully returns search response from specified index', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === mocks.from.request.url &&
        JSON.stringify(opts.data) === JSON.stringify({
          "requests": [
            {
              "entityTypes": [
                "message"
              ],
              "query": {
                "queryString": "contoso"
              },
              "from": 10,
              "size": 25,
              "queryAlterationOptions": {}
            }
          ]
        })) {
        return misc.deepClone(mocks.from.response.body);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { scopes: 'message', queryText: 'contoso', startIndex: 10 } });
    assert(loggerLogSpy.calledOnceWithExactly(fullSearchResponse));
  });

  it('successfully returns search response for specified query with selected properties', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === mocks.fields.request.url &&
        JSON.stringify(opts.data) === JSON.stringify({
          "requests": [
            {
              "entityTypes": [
                "message"
              ],
              "query": {
                "queryString": "contoso"
              },
              "from": 0,
              "size": 25,
              "fields": [
                "subject",
                "importance"
              ],
              "queryAlterationOptions": {}
            }
          ]
        })) {
        return misc.deepClone(mocks.fields.response.body);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { scopes: 'message', queryText: 'contoso', select: 'subject,importance' } });
    assert(loggerLogSpy.calledOnceWithExactly(selectedPropertiesSearchResponse));
  });

  it('successfully returns search results sorted by specified properties', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === mocks.sort.request.url &&
        JSON.stringify(opts.data) === JSON.stringify({
          "requests": [
            {
              "entityTypes": [
                "driveItem"
              ],
              "query": {
                "queryString": "contoso"
              },
              "from": 0,
              "size": 25,
              "sortProperties": [
                {
                  "name": "name",
                  "isDescending": true
                },
                {
                  "name": "createdDateTime",
                  "isDescending": false
                }
              ],
              "queryAlterationOptions": {}
            }
          ]
        })) {
        return misc.deepClone(mocks.sort.response.body);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { scopes: 'driveItem', queryText: 'contoso', sortBy: 'name:desc,createdDateTime' } });
    assert(loggerLogSpy.calledOnceWithExactly(selectedPropertiesSearchResponse));
  });

  it('successfully returns search results with spelling corrections', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === mocks.spelling.request.url &&
        JSON.stringify(opts.data) === JSON.stringify({
          "requests": [
            {
              "entityTypes": [
                "driveItem"
              ],
              "query": {
                "queryString": "contoso"
              },
              "from": 0,
              "size": 25,
              "queryAlterationOptions": {
                "enableModification": true,
                "enableSuggestion": true
              }
            }
          ]
        })) {
        return misc.deepClone(mocks.spelling.response.body);
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { scopes: 'driveItem', queryText: 'contoso', enableSpellingModification: true, enableSpellingSuggestion: true } });
    assert(loggerLogSpy.calledOnceWithExactly(spellingCorrectionSearchResponse));
  });

  it('successfully returns all results for specified query', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/search/query' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "requests": [
            {
              "entityTypes": [
                "message"
              ],
              "query": {
                "queryString": "contoso"
              },
              "from": 0,
              "size": 25,
              "queryAlterationOptions": {}
            }
          ]
        })) {
        return { "value": [fullSearchResponse] };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/search/query' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "requests": [
            {
              "entityTypes": [
                "message"
              ],
              "query": {
                "queryString": "contoso"
              },
              "from": 25,
              "size": 25,
              "queryAlterationOptions": {}
            }
          ]
        })) {
        return { "value": [fullSearchFromIndexResponse] };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { scopes: 'message', queryText: 'contoso', allResults: true } });
    assert(loggerLogSpy.calledOnceWithExactly(allResults));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/search/query' &&
        JSON.stringify(opts.data) === JSON.stringify({
          "requests": [
            {
              "entityTypes": [
                "message"
              ],
              "query": {
                "queryString": ""
              },
              "from": 0,
              "size": 25,
              "queryAlterationOptions": {}
            }
          ]
        })) {
        throw {
          "error": {
            "code": "BadRequest",
            "message": "SearchRequest Invalid (EntityRequest Invalid (searchRequest -> query is invalid (queryString required)))",
            "target": "",
            "details": [
              {
                "code": "System.AggregateException",
                "message": "EntityRequest Invalid (searchRequest -> query is invalid (queryString required))",
                "target": "",
                "details": [
                  {
                    "code": "System.AggregateException",
                    "message": "searchRequest -> query is invalid (queryString required)",
                    "target": "",
                    "details": [
                      {
                        "code": "Microsoft.SubstrateSearch.Api.ErrorReporting.ResourceBasedExceptions.BadRequestException",
                        "message": "queryString required",
                        "target": "",
                        "httpCode": 400
                      }
                    ],
                    "httpCode": 400
                  }
                ],
                "httpCode": 400
              }
            ],
            "httpCode": 400
          },
          "Instrumentation": {
            "TraceId": "ab7ab435-2cd2-f820-4949-42b5c0ee0ce3"
          }
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { scopes: 'message', queryText: '' } }), new CommandError('SearchRequest Invalid (EntityRequest Invalid (searchRequest -> query is invalid (queryString required)))'));
  });
});