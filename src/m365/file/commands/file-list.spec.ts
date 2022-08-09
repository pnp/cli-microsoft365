import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../appInsights';
import auth from '../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../cli';
import Command, { CommandError } from '../../../Command';
import request from '../../../request';
import { sinonUtil } from '../../../utils';
import commands from '../commands';
const command: Command = require('./file-list');

describe(commands.LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
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
    (command as any).items = [];
    (command as any).foldersToGetFilesFrom = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
    (command as any).items = undefined;
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines list of properties to display in text output', () => {
    assert.notStrictEqual(command.defaultProperties(), undefined);
  });

  it('loads files from the root site collection without trailing slash, document library without space, root folder, matching case', done => {
    sinon.stub(request, 'get').callsFake(opts => {
      const url: string = opts.url as string;

      switch (url) {
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/?$select=id':
          return Promise.resolve({
            "id": "contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42"
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42/drives?$select=webUrl,id':
          return Promise.resolve({
            "value": [
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYVs-s_Fc6EaRomQ91r_60hi",
                "webUrl": "https://contoso.sharepoint.com/Big%20list"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                "webUrl": "https://contoso.sharepoint.com/DemoDocs"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYUMhRqN81GDSYJHirLtImTh",
                "webUrl": "https://contoso.sharepoint.com/Shared%20Documents"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYWwU-RsQtl_RJxAlcRhJSYH",
                "webUrl": "https://contoso.sharepoint.com/JTDesignDocs"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYX2VXqL8cJvTbmPdTxwV8t4",
                "webUrl": "https://contoso.sharepoint.com/MCASDemoFiles"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYWaXP0JCVAyRZbch81buwYY",
                "webUrl": "https://contoso.sharepoint.com/RMSDemoLib"
              }
            ]
          });
        case 'https://graph.microsoft.com/v1.0/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root?$select=id':
          return Promise.resolve({
            "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ"
          });
        case 'https://graph.microsoft.com/v1.0/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/items/01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ/children':
          return Promise.resolve({
            "value": [
              {
                "createdDateTime": "2021-09-25T13:30:17Z",
                "eTag": "\"{A2331FD8-FAA8-4C41-A8A0-B6F27FB993AF},1\"",
                "id": "01YNDLPYOYD4Z2FKH2IFGKRIFW6J73TE5P",
                "lastModifiedDateTime": "2021-09-25T13:30:17Z",
                "name": "Fo'lde'r",
                "webUrl": "https://contoso.sharepoint.com/DemoDocs/Fo%27lde%27r",
                "cTag": "\"c:{A2331FD8-FAA8-4C41-A8A0-B6F27FB993AF},0\"",
                "size": 151415,
                "createdBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "802d8bad-2ae1-479a-bbbe-48aca058bc26",
                    "displayName": "MOD Administrator"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "802d8bad-2ae1-479a-bbbe-48aca058bc26",
                    "displayName": "MOD Administrator"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-09-25T13:30:17Z",
                  "lastModifiedDateTime": "2021-09-25T13:30:17Z"
                },
                "folder": {
                  "childCount": 2
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=b7b2648f-c50c-4d1e-bb99-9c1dcd5e37ae&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkNKTDB2UUp1SVF3enl1YUVHbDI1WVJpejJDdkM5bDdzRktVUFQzU0F2bTg9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.bmVGTE9nbDJ4c1lvem5sQ3pnTXRmb29HaTRUTWVVTlVaYXBkSVp3QWliRT0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:04Z",
                "eTag": "\"{B7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE},2\"",
                "id": "01YNDLPYMPMSZLODGFDZG3XGM4DXGV4N5O",
                "lastModifiedDateTime": "2021-06-19T11:01:04Z",
                "name": "Blog Post preview.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BB7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE%7D&file=Blog%20Post%20preview.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{B7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE},3\"",
                "size": 134144,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "4wybUm/SrJtzssP/YDDKBwKRYd8="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:04Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:04Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=4a3864b3-db49-4d6f-a72e-7e3269c44e33&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IjlUSENBN0loaWhMMUJhc2EyUjlneVBMYldaOGcxVS93VTYyRTU1NHU4MVE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.VEVUTXJYNk1mWmpBSFhudjJrVWMxTWQ0UzRCVzFad0F3K0QrV2VwcEVoaz0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:13Z",
                "eTag": "\"{4A3864B3-DB49-4D6F-A72E-7E3269C44E33},3\"",
                "id": "01YNDLPYNTMQ4EUSO3N5G2OLT6GJU4ITRT",
                "lastModifiedDateTime": "2021-06-19T11:01:13Z",
                "name": "Contoso Purchasing Permissions.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B4A3864B3-DB49-4D6F-A72E-7E3269C44E33%7D&file=Contoso%20Purchasing%20Permissions.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{4A3864B3-DB49-4D6F-A72E-7E3269C44E33},6\"",
                "dataLossPrevention": {
                  "block": {},
                  "notify": {}
                },
                "size": 27514,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "MTvtw1PC9WtEY0wycgD1aBIIX3A="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:13Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:13Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=1bcaf9f5-6788-42f1-9b77-ee43df61c4dd&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkozNjY5RnEwblpFWmluYTZHNjJSTTl0cGJCd1lxaWJuV2lING5SMVdDU1E9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.ck96RHVlYjdGMDhQWnhhc3UrZXhReHBaNHJnekcxUmVDUjMwdTNpSTJpND0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:23Z",
                "eTag": "\"{1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD},2\"",
                "id": "01YNDLPYPV7HFBXCDH6FBJW57OIPPWDRG5",
                "lastModifiedDateTime": "2021-06-19T11:01:23Z",
                "name": "Credit Cards.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD%7D&file=Credit%20Cards.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD},3\"",
                "size": 25245,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "xOKzXY19rWzvLGnGcjybqNcMtIc="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:23Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:23Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=bc13ae6b-5a75-427d-8915-fcb6e86edd11&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkpiTHU5OHpmTkxCMXhrZ1k1dTlkOUVONXU0UEZjWHpTQ0s0bTg3UU1LcTQ9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.dUFidWp1M1I4N2xwZmhyRDFQRGpZQ0ZQSHBsbVNOcVdFcDAwSjY3U1JGcz0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:32Z",
                "eTag": "\"{BC13AE6B-5A75-427D-8915-FCB6E86EDD11},2\"",
                "id": "01YNDLPYLLVYJ3Y5K2PVBISFP4W3UG5XIR",
                "lastModifiedDateTime": "2021-06-19T11:01:32Z",
                "name": "Customer Accounts.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BBC13AE6B-5A75-427D-8915-FCB6E86EDD11%7D&file=Customer%20Accounts.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{BC13AE6B-5A75-427D-8915-FCB6E86EDD11},3\"",
                "size": 28747,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "7nfr2jjjYOlbSFfYHX72+Ud1o6Y="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:32Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:32Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=5148569c-7459-40a8-a400-e56e4668ad62&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6InphTThqdWw5VW4wZlF3WWVTcnRmd1VWV2cyK2YvdXFVQWlzam9xTkdPTEU9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.K2RqWmd2WU5EWTVaWXprelRhUFBQcjZ0NkwxMzZJbnlPMG1QNDYyM3NvOD0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:00:46Z",
                "eTag": "\"{5148569C-7459-40A8-A400-E56E4668AD62},2\"",
                "id": "01YNDLPYM4KZEFCWLUVBAKIAHFNZDGRLLC",
                "lastModifiedDateTime": "2021-06-19T11:00:46Z",
                "name": "Customer Data.xlsx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B5148569C-7459-40A8-A400-E56E4668AD62%7D&file=Customer%20Data.xlsx&action=default&mobileredirect=true",
                "cTag": "\"c:{5148569C-7459-40A8-A400-E56E4668AD62},3\"",
                "size": 17661,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                  "hashes": {
                    "quickXorHash": "tTrnKAf/Pcq8/NXtvxHuAAUwIUs="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:00:46Z",
                  "lastModifiedDateTime": "2021-06-19T11:00:46Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=dd9ed836-18b0-44e7-b352-da56c300cdcc&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6Ik9XN2pqVTBJWkVxUEVLOU1yUXg0ZzgyQjd6ZnBzRWJ5TXNYY011NlpvbXc9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.eHNsek14b0tsQkZmNE9lT3UxWmlZMnRxZ0lub1Z5RzU4VnRkalRZWXcrQT0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:00:55Z",
                "eTag": "\"{DD9ED836-18B0-44E7-B352-DA56C300CDCC},2\"",
                "id": "01YNDLPYJW3CPN3MAY45CLGUW2K3BQBTOM",
                "lastModifiedDateTime": "2021-06-19T11:00:55Z",
                "name": "Q3 Sales and Marketing Expense Report Audit.pptx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BDD9ED836-18B0-44E7-B352-DA56C300CDCC%7D&file=Q3%20Sales%20and%20Marketing%20Expense%20Report%20Audit.pptx&action=edit&mobileredirect=true",
                "cTag": "\"c:{DD9ED836-18B0-44E7-B352-DA56C300CDCC},3\"",
                "size": 376325,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                  "hashes": {
                    "quickXorHash": "BCE09zeZNS8dnTx14w3stEm9g0g="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:00:55Z",
                  "lastModifiedDateTime": "2021-06-19T11:00:55Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=232b3df5-7f6c-4f06-b811-6029534bd1db&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6ImRNVEtaaGJuWVgyM3g4dURuQ0tXU0RQVFhjMWVSbGR3aWhrOXF6RjYvczA9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.M3ZVcE1NVnAxd0F4WDFiNUR5eHVoa3lZQ2x1eUUyVGwwdnFUZVNXaTJ1UT0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:44Z",
                "eTag": "\"{232B3DF5-7F6C-4F06-B811-6029534BD1DB},1\"",
                "id": "01YNDLPYPVHUVSG3D7AZH3QELAFFJUXUO3",
                "lastModifiedDateTime": "2021-06-19T11:01:44Z",
                "name": "Q3_Product_Strategy.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B232B3DF5-7F6C-4F06-B811-6029534BD1DB%7D&file=Q3_Product_Strategy.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{232B3DF5-7F6C-4F06-B811-6029534BD1DB},2\"",
                "size": 48473,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "54Y/Q4Ykxgg3N7yWL8iTlokMSS0="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:44Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:44Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=fcd62009-cf87-478e-a788-3853a9832f6b&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6ImVLQ0oyWDhpR3JvdGF0TGQ0d2hyOTNnajBKaUpUeEhUUHlmUWl2clJjOWc9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.bllWYkVsb3doSkxZbmFVTTE1TDdhUlU5NjNzWGNHUWRzbUNiOW1MN2s2az0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:52Z",
                "eTag": "\"{FCD62009-CF87-478E-A788-3853A9832F6B},1\"",
                "id": "01YNDLPYIJEDLPZB6PRZD2PCBYKOUYGL3L",
                "lastModifiedDateTime": "2021-06-19T11:01:52Z",
                "name": "Sales Memo.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BFCD62009-CF87-478E-A788-3853A9832F6B%7D&file=Sales%20Memo.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{FCD62009-CF87-478E-A788-3853A9832F6B},2\"",
                "size": 33737,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "dNTTVcFU+d1H7Ou3seZvQI1pGnk="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:52Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:52Z"
                }
              }
            ]
          });
        default:
          return Promise.reject(`Invalid GET request: ${url}`);
      }
    });

    command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: 'DemoDocs'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        assert(loggerLogSpy.calledWith([
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=b7b2648f-c50c-4d1e-bb99-9c1dcd5e37ae&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkNKTDB2UUp1SVF3enl1YUVHbDI1WVJpejJDdkM5bDdzRktVUFQzU0F2bTg9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.bmVGTE9nbDJ4c1lvem5sQ3pnTXRmb29HaTRUTWVVTlVaYXBkSVp3QWliRT0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:04Z",
            "eTag": "\"{B7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE},2\"",
            "id": "01YNDLPYMPMSZLODGFDZG3XGM4DXGV4N5O",
            "lastModifiedDateTime": "2021-06-19T11:01:04Z",
            "name": "Blog Post preview.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BB7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE%7D&file=Blog%20Post%20preview.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{B7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE},3\"",
            "size": 134144,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "4wybUm/SrJtzssP/YDDKBwKRYd8="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:04Z",
              "lastModifiedDateTime": "2021-06-19T11:01:04Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=4a3864b3-db49-4d6f-a72e-7e3269c44e33&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IjlUSENBN0loaWhMMUJhc2EyUjlneVBMYldaOGcxVS93VTYyRTU1NHU4MVE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.VEVUTXJYNk1mWmpBSFhudjJrVWMxTWQ0UzRCVzFad0F3K0QrV2VwcEVoaz0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:13Z",
            "eTag": "\"{4A3864B3-DB49-4D6F-A72E-7E3269C44E33},3\"",
            "id": "01YNDLPYNTMQ4EUSO3N5G2OLT6GJU4ITRT",
            "lastModifiedDateTime": "2021-06-19T11:01:13Z",
            "name": "Contoso Purchasing Permissions.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B4A3864B3-DB49-4D6F-A72E-7E3269C44E33%7D&file=Contoso%20Purchasing%20Permissions.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{4A3864B3-DB49-4D6F-A72E-7E3269C44E33},6\"",
            "dataLossPrevention": {
              "block": {},
              "notify": {}
            },
            "size": 27514,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "MTvtw1PC9WtEY0wycgD1aBIIX3A="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:13Z",
              "lastModifiedDateTime": "2021-06-19T11:01:13Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=1bcaf9f5-6788-42f1-9b77-ee43df61c4dd&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkozNjY5RnEwblpFWmluYTZHNjJSTTl0cGJCd1lxaWJuV2lING5SMVdDU1E9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.ck96RHVlYjdGMDhQWnhhc3UrZXhReHBaNHJnekcxUmVDUjMwdTNpSTJpND0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:23Z",
            "eTag": "\"{1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD},2\"",
            "id": "01YNDLPYPV7HFBXCDH6FBJW57OIPPWDRG5",
            "lastModifiedDateTime": "2021-06-19T11:01:23Z",
            "name": "Credit Cards.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD%7D&file=Credit%20Cards.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD},3\"",
            "size": 25245,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "xOKzXY19rWzvLGnGcjybqNcMtIc="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:23Z",
              "lastModifiedDateTime": "2021-06-19T11:01:23Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=bc13ae6b-5a75-427d-8915-fcb6e86edd11&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkpiTHU5OHpmTkxCMXhrZ1k1dTlkOUVONXU0UEZjWHpTQ0s0bTg3UU1LcTQ9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.dUFidWp1M1I4N2xwZmhyRDFQRGpZQ0ZQSHBsbVNOcVdFcDAwSjY3U1JGcz0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:32Z",
            "eTag": "\"{BC13AE6B-5A75-427D-8915-FCB6E86EDD11},2\"",
            "id": "01YNDLPYLLVYJ3Y5K2PVBISFP4W3UG5XIR",
            "lastModifiedDateTime": "2021-06-19T11:01:32Z",
            "name": "Customer Accounts.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BBC13AE6B-5A75-427D-8915-FCB6E86EDD11%7D&file=Customer%20Accounts.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{BC13AE6B-5A75-427D-8915-FCB6E86EDD11},3\"",
            "size": 28747,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "7nfr2jjjYOlbSFfYHX72+Ud1o6Y="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:32Z",
              "lastModifiedDateTime": "2021-06-19T11:01:32Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=5148569c-7459-40a8-a400-e56e4668ad62&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6InphTThqdWw5VW4wZlF3WWVTcnRmd1VWV2cyK2YvdXFVQWlzam9xTkdPTEU9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.K2RqWmd2WU5EWTVaWXprelRhUFBQcjZ0NkwxMzZJbnlPMG1QNDYyM3NvOD0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:00:46Z",
            "eTag": "\"{5148569C-7459-40A8-A400-E56E4668AD62},2\"",
            "id": "01YNDLPYM4KZEFCWLUVBAKIAHFNZDGRLLC",
            "lastModifiedDateTime": "2021-06-19T11:00:46Z",
            "name": "Customer Data.xlsx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B5148569C-7459-40A8-A400-E56E4668AD62%7D&file=Customer%20Data.xlsx&action=default&mobileredirect=true",
            "cTag": "\"c:{5148569C-7459-40A8-A400-E56E4668AD62},3\"",
            "size": 17661,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              "hashes": {
                "quickXorHash": "tTrnKAf/Pcq8/NXtvxHuAAUwIUs="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:00:46Z",
              "lastModifiedDateTime": "2021-06-19T11:00:46Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=dd9ed836-18b0-44e7-b352-da56c300cdcc&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6Ik9XN2pqVTBJWkVxUEVLOU1yUXg0ZzgyQjd6ZnBzRWJ5TXNYY011NlpvbXc9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.eHNsek14b0tsQkZmNE9lT3UxWmlZMnRxZ0lub1Z5RzU4VnRkalRZWXcrQT0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:00:55Z",
            "eTag": "\"{DD9ED836-18B0-44E7-B352-DA56C300CDCC},2\"",
            "id": "01YNDLPYJW3CPN3MAY45CLGUW2K3BQBTOM",
            "lastModifiedDateTime": "2021-06-19T11:00:55Z",
            "name": "Q3 Sales and Marketing Expense Report Audit.pptx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BDD9ED836-18B0-44E7-B352-DA56C300CDCC%7D&file=Q3%20Sales%20and%20Marketing%20Expense%20Report%20Audit.pptx&action=edit&mobileredirect=true",
            "cTag": "\"c:{DD9ED836-18B0-44E7-B352-DA56C300CDCC},3\"",
            "size": 376325,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
              "hashes": {
                "quickXorHash": "BCE09zeZNS8dnTx14w3stEm9g0g="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:00:55Z",
              "lastModifiedDateTime": "2021-06-19T11:00:55Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=232b3df5-7f6c-4f06-b811-6029534bd1db&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6ImRNVEtaaGJuWVgyM3g4dURuQ0tXU0RQVFhjMWVSbGR3aWhrOXF6RjYvczA9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.M3ZVcE1NVnAxd0F4WDFiNUR5eHVoa3lZQ2x1eUUyVGwwdnFUZVNXaTJ1UT0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:44Z",
            "eTag": "\"{232B3DF5-7F6C-4F06-B811-6029534BD1DB},1\"",
            "id": "01YNDLPYPVHUVSG3D7AZH3QELAFFJUXUO3",
            "lastModifiedDateTime": "2021-06-19T11:01:44Z",
            "name": "Q3_Product_Strategy.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B232B3DF5-7F6C-4F06-B811-6029534BD1DB%7D&file=Q3_Product_Strategy.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{232B3DF5-7F6C-4F06-B811-6029534BD1DB},2\"",
            "size": 48473,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "54Y/Q4Ykxgg3N7yWL8iTlokMSS0="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:44Z",
              "lastModifiedDateTime": "2021-06-19T11:01:44Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=fcd62009-cf87-478e-a788-3853a9832f6b&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6ImVLQ0oyWDhpR3JvdGF0TGQ0d2hyOTNnajBKaUpUeEhUUHlmUWl2clJjOWc9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.bllWYkVsb3doSkxZbmFVTTE1TDdhUlU5NjNzWGNHUWRzbUNiOW1MN2s2az0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:52Z",
            "eTag": "\"{FCD62009-CF87-478E-A788-3853A9832F6B},1\"",
            "id": "01YNDLPYIJEDLPZB6PRZD2PCBYKOUYGL3L",
            "lastModifiedDateTime": "2021-06-19T11:01:52Z",
            "name": "Sales Memo.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BFCD62009-CF87-478E-A788-3853A9832F6B%7D&file=Sales%20Memo.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{FCD62009-CF87-478E-A788-3853A9832F6B},2\"",
            "size": 33737,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "dNTTVcFU+d1H7Ou3seZvQI1pGnk="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:52Z",
              "lastModifiedDateTime": "2021-06-19T11:01:52Z"
            },
            "lastModifiedByUser": "Provisioning User"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('loads files from the root site collection without trailing slash, document library without space, root folder, matching case (debug)', done => {
    sinon.stub(request, 'get').callsFake(opts => {
      const url: string = opts.url as string;

      switch (url) {
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/?$select=id':
          return Promise.resolve({
            "id": "contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42"
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42/drives?$select=webUrl,id':
          return Promise.resolve({
            "value": [
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYVs-s_Fc6EaRomQ91r_60hi",
                "webUrl": "https://contoso.sharepoint.com/Big%20list"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                "webUrl": "https://contoso.sharepoint.com/DemoDocs"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYUMhRqN81GDSYJHirLtImTh",
                "webUrl": "https://contoso.sharepoint.com/Shared%20Documents"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYWwU-RsQtl_RJxAlcRhJSYH",
                "webUrl": "https://contoso.sharepoint.com/JTDesignDocs"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYX2VXqL8cJvTbmPdTxwV8t4",
                "webUrl": "https://contoso.sharepoint.com/MCASDemoFiles"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYWaXP0JCVAyRZbch81buwYY",
                "webUrl": "https://contoso.sharepoint.com/RMSDemoLib"
              }
            ]
          });
        case 'https://graph.microsoft.com/v1.0/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root?$select=id':
          return Promise.resolve({
            "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ"
          });
        case 'https://graph.microsoft.com/v1.0/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/items/01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ/children':
          return Promise.resolve({
            "value": [
              {
                "createdDateTime": "2021-09-25T13:30:17Z",
                "eTag": "\"{A2331FD8-FAA8-4C41-A8A0-B6F27FB993AF},1\"",
                "id": "01YNDLPYOYD4Z2FKH2IFGKRIFW6J73TE5P",
                "lastModifiedDateTime": "2021-09-25T13:30:17Z",
                "name": "Fo'lde'r",
                "webUrl": "https://contoso.sharepoint.com/DemoDocs/Fo%27lde%27r",
                "cTag": "\"c:{A2331FD8-FAA8-4C41-A8A0-B6F27FB993AF},0\"",
                "size": 151415,
                "createdBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "802d8bad-2ae1-479a-bbbe-48aca058bc26",
                    "displayName": "MOD Administrator"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "802d8bad-2ae1-479a-bbbe-48aca058bc26",
                    "displayName": "MOD Administrator"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-09-25T13:30:17Z",
                  "lastModifiedDateTime": "2021-09-25T13:30:17Z"
                },
                "folder": {
                  "childCount": 2
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=b7b2648f-c50c-4d1e-bb99-9c1dcd5e37ae&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkNKTDB2UUp1SVF3enl1YUVHbDI1WVJpejJDdkM5bDdzRktVUFQzU0F2bTg9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.bmVGTE9nbDJ4c1lvem5sQ3pnTXRmb29HaTRUTWVVTlVaYXBkSVp3QWliRT0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:04Z",
                "eTag": "\"{B7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE},2\"",
                "id": "01YNDLPYMPMSZLODGFDZG3XGM4DXGV4N5O",
                "lastModifiedDateTime": "2021-06-19T11:01:04Z",
                "name": "Blog Post preview.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BB7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE%7D&file=Blog%20Post%20preview.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{B7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE},3\"",
                "size": 134144,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "4wybUm/SrJtzssP/YDDKBwKRYd8="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:04Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:04Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=4a3864b3-db49-4d6f-a72e-7e3269c44e33&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IjlUSENBN0loaWhMMUJhc2EyUjlneVBMYldaOGcxVS93VTYyRTU1NHU4MVE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.VEVUTXJYNk1mWmpBSFhudjJrVWMxTWQ0UzRCVzFad0F3K0QrV2VwcEVoaz0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:13Z",
                "eTag": "\"{4A3864B3-DB49-4D6F-A72E-7E3269C44E33},3\"",
                "id": "01YNDLPYNTMQ4EUSO3N5G2OLT6GJU4ITRT",
                "lastModifiedDateTime": "2021-06-19T11:01:13Z",
                "name": "Contoso Purchasing Permissions.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B4A3864B3-DB49-4D6F-A72E-7E3269C44E33%7D&file=Contoso%20Purchasing%20Permissions.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{4A3864B3-DB49-4D6F-A72E-7E3269C44E33},6\"",
                "dataLossPrevention": {
                  "block": {},
                  "notify": {}
                },
                "size": 27514,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "MTvtw1PC9WtEY0wycgD1aBIIX3A="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:13Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:13Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=1bcaf9f5-6788-42f1-9b77-ee43df61c4dd&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkozNjY5RnEwblpFWmluYTZHNjJSTTl0cGJCd1lxaWJuV2lING5SMVdDU1E9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.ck96RHVlYjdGMDhQWnhhc3UrZXhReHBaNHJnekcxUmVDUjMwdTNpSTJpND0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:23Z",
                "eTag": "\"{1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD},2\"",
                "id": "01YNDLPYPV7HFBXCDH6FBJW57OIPPWDRG5",
                "lastModifiedDateTime": "2021-06-19T11:01:23Z",
                "name": "Credit Cards.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD%7D&file=Credit%20Cards.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD},3\"",
                "size": 25245,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "xOKzXY19rWzvLGnGcjybqNcMtIc="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:23Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:23Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=bc13ae6b-5a75-427d-8915-fcb6e86edd11&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkpiTHU5OHpmTkxCMXhrZ1k1dTlkOUVONXU0UEZjWHpTQ0s0bTg3UU1LcTQ9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.dUFidWp1M1I4N2xwZmhyRDFQRGpZQ0ZQSHBsbVNOcVdFcDAwSjY3U1JGcz0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:32Z",
                "eTag": "\"{BC13AE6B-5A75-427D-8915-FCB6E86EDD11},2\"",
                "id": "01YNDLPYLLVYJ3Y5K2PVBISFP4W3UG5XIR",
                "lastModifiedDateTime": "2021-06-19T11:01:32Z",
                "name": "Customer Accounts.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BBC13AE6B-5A75-427D-8915-FCB6E86EDD11%7D&file=Customer%20Accounts.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{BC13AE6B-5A75-427D-8915-FCB6E86EDD11},3\"",
                "size": 28747,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "7nfr2jjjYOlbSFfYHX72+Ud1o6Y="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:32Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:32Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=5148569c-7459-40a8-a400-e56e4668ad62&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6InphTThqdWw5VW4wZlF3WWVTcnRmd1VWV2cyK2YvdXFVQWlzam9xTkdPTEU9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.K2RqWmd2WU5EWTVaWXprelRhUFBQcjZ0NkwxMzZJbnlPMG1QNDYyM3NvOD0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:00:46Z",
                "eTag": "\"{5148569C-7459-40A8-A400-E56E4668AD62},2\"",
                "id": "01YNDLPYM4KZEFCWLUVBAKIAHFNZDGRLLC",
                "lastModifiedDateTime": "2021-06-19T11:00:46Z",
                "name": "Customer Data.xlsx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B5148569C-7459-40A8-A400-E56E4668AD62%7D&file=Customer%20Data.xlsx&action=default&mobileredirect=true",
                "cTag": "\"c:{5148569C-7459-40A8-A400-E56E4668AD62},3\"",
                "size": 17661,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                  "hashes": {
                    "quickXorHash": "tTrnKAf/Pcq8/NXtvxHuAAUwIUs="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:00:46Z",
                  "lastModifiedDateTime": "2021-06-19T11:00:46Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=dd9ed836-18b0-44e7-b352-da56c300cdcc&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6Ik9XN2pqVTBJWkVxUEVLOU1yUXg0ZzgyQjd6ZnBzRWJ5TXNYY011NlpvbXc9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.eHNsek14b0tsQkZmNE9lT3UxWmlZMnRxZ0lub1Z5RzU4VnRkalRZWXcrQT0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:00:55Z",
                "eTag": "\"{DD9ED836-18B0-44E7-B352-DA56C300CDCC},2\"",
                "id": "01YNDLPYJW3CPN3MAY45CLGUW2K3BQBTOM",
                "lastModifiedDateTime": "2021-06-19T11:00:55Z",
                "name": "Q3 Sales and Marketing Expense Report Audit.pptx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BDD9ED836-18B0-44E7-B352-DA56C300CDCC%7D&file=Q3%20Sales%20and%20Marketing%20Expense%20Report%20Audit.pptx&action=edit&mobileredirect=true",
                "cTag": "\"c:{DD9ED836-18B0-44E7-B352-DA56C300CDCC},3\"",
                "size": 376325,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                  "hashes": {
                    "quickXorHash": "BCE09zeZNS8dnTx14w3stEm9g0g="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:00:55Z",
                  "lastModifiedDateTime": "2021-06-19T11:00:55Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=232b3df5-7f6c-4f06-b811-6029534bd1db&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6ImRNVEtaaGJuWVgyM3g4dURuQ0tXU0RQVFhjMWVSbGR3aWhrOXF6RjYvczA9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.M3ZVcE1NVnAxd0F4WDFiNUR5eHVoa3lZQ2x1eUUyVGwwdnFUZVNXaTJ1UT0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:44Z",
                "eTag": "\"{232B3DF5-7F6C-4F06-B811-6029534BD1DB},1\"",
                "id": "01YNDLPYPVHUVSG3D7AZH3QELAFFJUXUO3",
                "lastModifiedDateTime": "2021-06-19T11:01:44Z",
                "name": "Q3_Product_Strategy.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B232B3DF5-7F6C-4F06-B811-6029534BD1DB%7D&file=Q3_Product_Strategy.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{232B3DF5-7F6C-4F06-B811-6029534BD1DB},2\"",
                "size": 48473,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "54Y/Q4Ykxgg3N7yWL8iTlokMSS0="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:44Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:44Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=fcd62009-cf87-478e-a788-3853a9832f6b&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6ImVLQ0oyWDhpR3JvdGF0TGQ0d2hyOTNnajBKaUpUeEhUUHlmUWl2clJjOWc9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.bllWYkVsb3doSkxZbmFVTTE1TDdhUlU5NjNzWGNHUWRzbUNiOW1MN2s2az0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:52Z",
                "eTag": "\"{FCD62009-CF87-478E-A788-3853A9832F6B},1\"",
                "id": "01YNDLPYIJEDLPZB6PRZD2PCBYKOUYGL3L",
                "lastModifiedDateTime": "2021-06-19T11:01:52Z",
                "name": "Sales Memo.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BFCD62009-CF87-478E-A788-3853A9832F6B%7D&file=Sales%20Memo.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{FCD62009-CF87-478E-A788-3853A9832F6B},2\"",
                "size": 33737,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "dNTTVcFU+d1H7Ou3seZvQI1pGnk="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:52Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:52Z"
                }
              }
            ]
          });
        default:
          return Promise.reject(`Invalid GET request: ${url}`);
      }
    });

    command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: 'DemoDocs',
        debug: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        assert(loggerLogSpy.calledWith([
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=b7b2648f-c50c-4d1e-bb99-9c1dcd5e37ae&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkNKTDB2UUp1SVF3enl1YUVHbDI1WVJpejJDdkM5bDdzRktVUFQzU0F2bTg9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.bmVGTE9nbDJ4c1lvem5sQ3pnTXRmb29HaTRUTWVVTlVaYXBkSVp3QWliRT0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:04Z",
            "eTag": "\"{B7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE},2\"",
            "id": "01YNDLPYMPMSZLODGFDZG3XGM4DXGV4N5O",
            "lastModifiedDateTime": "2021-06-19T11:01:04Z",
            "name": "Blog Post preview.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BB7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE%7D&file=Blog%20Post%20preview.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{B7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE},3\"",
            "size": 134144,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "4wybUm/SrJtzssP/YDDKBwKRYd8="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:04Z",
              "lastModifiedDateTime": "2021-06-19T11:01:04Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=4a3864b3-db49-4d6f-a72e-7e3269c44e33&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IjlUSENBN0loaWhMMUJhc2EyUjlneVBMYldaOGcxVS93VTYyRTU1NHU4MVE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.VEVUTXJYNk1mWmpBSFhudjJrVWMxTWQ0UzRCVzFad0F3K0QrV2VwcEVoaz0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:13Z",
            "eTag": "\"{4A3864B3-DB49-4D6F-A72E-7E3269C44E33},3\"",
            "id": "01YNDLPYNTMQ4EUSO3N5G2OLT6GJU4ITRT",
            "lastModifiedDateTime": "2021-06-19T11:01:13Z",
            "name": "Contoso Purchasing Permissions.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B4A3864B3-DB49-4D6F-A72E-7E3269C44E33%7D&file=Contoso%20Purchasing%20Permissions.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{4A3864B3-DB49-4D6F-A72E-7E3269C44E33},6\"",
            "dataLossPrevention": {
              "block": {},
              "notify": {}
            },
            "size": 27514,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "MTvtw1PC9WtEY0wycgD1aBIIX3A="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:13Z",
              "lastModifiedDateTime": "2021-06-19T11:01:13Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=1bcaf9f5-6788-42f1-9b77-ee43df61c4dd&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkozNjY5RnEwblpFWmluYTZHNjJSTTl0cGJCd1lxaWJuV2lING5SMVdDU1E9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.ck96RHVlYjdGMDhQWnhhc3UrZXhReHBaNHJnekcxUmVDUjMwdTNpSTJpND0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:23Z",
            "eTag": "\"{1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD},2\"",
            "id": "01YNDLPYPV7HFBXCDH6FBJW57OIPPWDRG5",
            "lastModifiedDateTime": "2021-06-19T11:01:23Z",
            "name": "Credit Cards.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD%7D&file=Credit%20Cards.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD},3\"",
            "size": 25245,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "xOKzXY19rWzvLGnGcjybqNcMtIc="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:23Z",
              "lastModifiedDateTime": "2021-06-19T11:01:23Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=bc13ae6b-5a75-427d-8915-fcb6e86edd11&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkpiTHU5OHpmTkxCMXhrZ1k1dTlkOUVONXU0UEZjWHpTQ0s0bTg3UU1LcTQ9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.dUFidWp1M1I4N2xwZmhyRDFQRGpZQ0ZQSHBsbVNOcVdFcDAwSjY3U1JGcz0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:32Z",
            "eTag": "\"{BC13AE6B-5A75-427D-8915-FCB6E86EDD11},2\"",
            "id": "01YNDLPYLLVYJ3Y5K2PVBISFP4W3UG5XIR",
            "lastModifiedDateTime": "2021-06-19T11:01:32Z",
            "name": "Customer Accounts.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BBC13AE6B-5A75-427D-8915-FCB6E86EDD11%7D&file=Customer%20Accounts.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{BC13AE6B-5A75-427D-8915-FCB6E86EDD11},3\"",
            "size": 28747,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "7nfr2jjjYOlbSFfYHX72+Ud1o6Y="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:32Z",
              "lastModifiedDateTime": "2021-06-19T11:01:32Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=5148569c-7459-40a8-a400-e56e4668ad62&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6InphTThqdWw5VW4wZlF3WWVTcnRmd1VWV2cyK2YvdXFVQWlzam9xTkdPTEU9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.K2RqWmd2WU5EWTVaWXprelRhUFBQcjZ0NkwxMzZJbnlPMG1QNDYyM3NvOD0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:00:46Z",
            "eTag": "\"{5148569C-7459-40A8-A400-E56E4668AD62},2\"",
            "id": "01YNDLPYM4KZEFCWLUVBAKIAHFNZDGRLLC",
            "lastModifiedDateTime": "2021-06-19T11:00:46Z",
            "name": "Customer Data.xlsx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B5148569C-7459-40A8-A400-E56E4668AD62%7D&file=Customer%20Data.xlsx&action=default&mobileredirect=true",
            "cTag": "\"c:{5148569C-7459-40A8-A400-E56E4668AD62},3\"",
            "size": 17661,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              "hashes": {
                "quickXorHash": "tTrnKAf/Pcq8/NXtvxHuAAUwIUs="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:00:46Z",
              "lastModifiedDateTime": "2021-06-19T11:00:46Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=dd9ed836-18b0-44e7-b352-da56c300cdcc&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6Ik9XN2pqVTBJWkVxUEVLOU1yUXg0ZzgyQjd6ZnBzRWJ5TXNYY011NlpvbXc9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.eHNsek14b0tsQkZmNE9lT3UxWmlZMnRxZ0lub1Z5RzU4VnRkalRZWXcrQT0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:00:55Z",
            "eTag": "\"{DD9ED836-18B0-44E7-B352-DA56C300CDCC},2\"",
            "id": "01YNDLPYJW3CPN3MAY45CLGUW2K3BQBTOM",
            "lastModifiedDateTime": "2021-06-19T11:00:55Z",
            "name": "Q3 Sales and Marketing Expense Report Audit.pptx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BDD9ED836-18B0-44E7-B352-DA56C300CDCC%7D&file=Q3%20Sales%20and%20Marketing%20Expense%20Report%20Audit.pptx&action=edit&mobileredirect=true",
            "cTag": "\"c:{DD9ED836-18B0-44E7-B352-DA56C300CDCC},3\"",
            "size": 376325,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
              "hashes": {
                "quickXorHash": "BCE09zeZNS8dnTx14w3stEm9g0g="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:00:55Z",
              "lastModifiedDateTime": "2021-06-19T11:00:55Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=232b3df5-7f6c-4f06-b811-6029534bd1db&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6ImRNVEtaaGJuWVgyM3g4dURuQ0tXU0RQVFhjMWVSbGR3aWhrOXF6RjYvczA9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.M3ZVcE1NVnAxd0F4WDFiNUR5eHVoa3lZQ2x1eUUyVGwwdnFUZVNXaTJ1UT0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:44Z",
            "eTag": "\"{232B3DF5-7F6C-4F06-B811-6029534BD1DB},1\"",
            "id": "01YNDLPYPVHUVSG3D7AZH3QELAFFJUXUO3",
            "lastModifiedDateTime": "2021-06-19T11:01:44Z",
            "name": "Q3_Product_Strategy.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B232B3DF5-7F6C-4F06-B811-6029534BD1DB%7D&file=Q3_Product_Strategy.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{232B3DF5-7F6C-4F06-B811-6029534BD1DB},2\"",
            "size": 48473,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "54Y/Q4Ykxgg3N7yWL8iTlokMSS0="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:44Z",
              "lastModifiedDateTime": "2021-06-19T11:01:44Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=fcd62009-cf87-478e-a788-3853a9832f6b&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6ImVLQ0oyWDhpR3JvdGF0TGQ0d2hyOTNnajBKaUpUeEhUUHlmUWl2clJjOWc9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.bllWYkVsb3doSkxZbmFVTTE1TDdhUlU5NjNzWGNHUWRzbUNiOW1MN2s2az0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:52Z",
            "eTag": "\"{FCD62009-CF87-478E-A788-3853A9832F6B},1\"",
            "id": "01YNDLPYIJEDLPZB6PRZD2PCBYKOUYGL3L",
            "lastModifiedDateTime": "2021-06-19T11:01:52Z",
            "name": "Sales Memo.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BFCD62009-CF87-478E-A788-3853A9832F6B%7D&file=Sales%20Memo.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{FCD62009-CF87-478E-A788-3853A9832F6B},2\"",
            "size": 33737,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "dNTTVcFU+d1H7Ou3seZvQI1pGnk="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:52Z",
              "lastModifiedDateTime": "2021-06-19T11:01:52Z"
            },
            "lastModifiedByUser": "Provisioning User"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('loads files from the root site collection with trailing slash, document library without space, root folder, matching case', done => {
    sinon.stub(request, 'get').callsFake(opts => {
      const url: string = opts.url as string;

      switch (url) {
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/?$select=id':
          return Promise.resolve({
            "id": "contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42"
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42/drives?$select=webUrl,id':
          return Promise.resolve({
            "value": [
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYVs-s_Fc6EaRomQ91r_60hi",
                "webUrl": "https://contoso.sharepoint.com/Big%20list"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                "webUrl": "https://contoso.sharepoint.com/DemoDocs"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYUMhRqN81GDSYJHirLtImTh",
                "webUrl": "https://contoso.sharepoint.com/Shared%20Documents"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYWwU-RsQtl_RJxAlcRhJSYH",
                "webUrl": "https://contoso.sharepoint.com/JTDesignDocs"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYX2VXqL8cJvTbmPdTxwV8t4",
                "webUrl": "https://contoso.sharepoint.com/MCASDemoFiles"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYWaXP0JCVAyRZbch81buwYY",
                "webUrl": "https://contoso.sharepoint.com/RMSDemoLib"
              }
            ]
          });
        case 'https://graph.microsoft.com/v1.0/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root?$select=id':
          return Promise.resolve({
            "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ"
          });
        case 'https://graph.microsoft.com/v1.0/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/items/01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ/children':
          return Promise.resolve({
            "value": [
              {
                "createdDateTime": "2021-09-25T13:30:17Z",
                "eTag": "\"{A2331FD8-FAA8-4C41-A8A0-B6F27FB993AF},1\"",
                "id": "01YNDLPYOYD4Z2FKH2IFGKRIFW6J73TE5P",
                "lastModifiedDateTime": "2021-09-25T13:30:17Z",
                "name": "Fo'lde'r",
                "webUrl": "https://contoso.sharepoint.com/DemoDocs/Fo%27lde%27r",
                "cTag": "\"c:{A2331FD8-FAA8-4C41-A8A0-B6F27FB993AF},0\"",
                "size": 151415,
                "createdBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "802d8bad-2ae1-479a-bbbe-48aca058bc26",
                    "displayName": "MOD Administrator"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "802d8bad-2ae1-479a-bbbe-48aca058bc26",
                    "displayName": "MOD Administrator"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-09-25T13:30:17Z",
                  "lastModifiedDateTime": "2021-09-25T13:30:17Z"
                },
                "folder": {
                  "childCount": 2
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=b7b2648f-c50c-4d1e-bb99-9c1dcd5e37ae&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkNKTDB2UUp1SVF3enl1YUVHbDI1WVJpejJDdkM5bDdzRktVUFQzU0F2bTg9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.bmVGTE9nbDJ4c1lvem5sQ3pnTXRmb29HaTRUTWVVTlVaYXBkSVp3QWliRT0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:04Z",
                "eTag": "\"{B7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE},2\"",
                "id": "01YNDLPYMPMSZLODGFDZG3XGM4DXGV4N5O",
                "lastModifiedDateTime": "2021-06-19T11:01:04Z",
                "name": "Blog Post preview.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BB7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE%7D&file=Blog%20Post%20preview.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{B7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE},3\"",
                "size": 134144,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "4wybUm/SrJtzssP/YDDKBwKRYd8="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:04Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:04Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=4a3864b3-db49-4d6f-a72e-7e3269c44e33&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IjlUSENBN0loaWhMMUJhc2EyUjlneVBMYldaOGcxVS93VTYyRTU1NHU4MVE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.VEVUTXJYNk1mWmpBSFhudjJrVWMxTWQ0UzRCVzFad0F3K0QrV2VwcEVoaz0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:13Z",
                "eTag": "\"{4A3864B3-DB49-4D6F-A72E-7E3269C44E33},3\"",
                "id": "01YNDLPYNTMQ4EUSO3N5G2OLT6GJU4ITRT",
                "lastModifiedDateTime": "2021-06-19T11:01:13Z",
                "name": "Contoso Purchasing Permissions.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B4A3864B3-DB49-4D6F-A72E-7E3269C44E33%7D&file=Contoso%20Purchasing%20Permissions.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{4A3864B3-DB49-4D6F-A72E-7E3269C44E33},6\"",
                "dataLossPrevention": {
                  "block": {},
                  "notify": {}
                },
                "size": 27514,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "MTvtw1PC9WtEY0wycgD1aBIIX3A="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:13Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:13Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=1bcaf9f5-6788-42f1-9b77-ee43df61c4dd&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkozNjY5RnEwblpFWmluYTZHNjJSTTl0cGJCd1lxaWJuV2lING5SMVdDU1E9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.ck96RHVlYjdGMDhQWnhhc3UrZXhReHBaNHJnekcxUmVDUjMwdTNpSTJpND0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:23Z",
                "eTag": "\"{1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD},2\"",
                "id": "01YNDLPYPV7HFBXCDH6FBJW57OIPPWDRG5",
                "lastModifiedDateTime": "2021-06-19T11:01:23Z",
                "name": "Credit Cards.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD%7D&file=Credit%20Cards.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD},3\"",
                "size": 25245,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "xOKzXY19rWzvLGnGcjybqNcMtIc="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:23Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:23Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=bc13ae6b-5a75-427d-8915-fcb6e86edd11&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkpiTHU5OHpmTkxCMXhrZ1k1dTlkOUVONXU0UEZjWHpTQ0s0bTg3UU1LcTQ9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.dUFidWp1M1I4N2xwZmhyRDFQRGpZQ0ZQSHBsbVNOcVdFcDAwSjY3U1JGcz0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:32Z",
                "eTag": "\"{BC13AE6B-5A75-427D-8915-FCB6E86EDD11},2\"",
                "id": "01YNDLPYLLVYJ3Y5K2PVBISFP4W3UG5XIR",
                "lastModifiedDateTime": "2021-06-19T11:01:32Z",
                "name": "Customer Accounts.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BBC13AE6B-5A75-427D-8915-FCB6E86EDD11%7D&file=Customer%20Accounts.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{BC13AE6B-5A75-427D-8915-FCB6E86EDD11},3\"",
                "size": 28747,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "7nfr2jjjYOlbSFfYHX72+Ud1o6Y="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:32Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:32Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=5148569c-7459-40a8-a400-e56e4668ad62&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6InphTThqdWw5VW4wZlF3WWVTcnRmd1VWV2cyK2YvdXFVQWlzam9xTkdPTEU9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.K2RqWmd2WU5EWTVaWXprelRhUFBQcjZ0NkwxMzZJbnlPMG1QNDYyM3NvOD0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:00:46Z",
                "eTag": "\"{5148569C-7459-40A8-A400-E56E4668AD62},2\"",
                "id": "01YNDLPYM4KZEFCWLUVBAKIAHFNZDGRLLC",
                "lastModifiedDateTime": "2021-06-19T11:00:46Z",
                "name": "Customer Data.xlsx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B5148569C-7459-40A8-A400-E56E4668AD62%7D&file=Customer%20Data.xlsx&action=default&mobileredirect=true",
                "cTag": "\"c:{5148569C-7459-40A8-A400-E56E4668AD62},3\"",
                "size": 17661,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                  "hashes": {
                    "quickXorHash": "tTrnKAf/Pcq8/NXtvxHuAAUwIUs="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:00:46Z",
                  "lastModifiedDateTime": "2021-06-19T11:00:46Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=dd9ed836-18b0-44e7-b352-da56c300cdcc&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6Ik9XN2pqVTBJWkVxUEVLOU1yUXg0ZzgyQjd6ZnBzRWJ5TXNYY011NlpvbXc9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.eHNsek14b0tsQkZmNE9lT3UxWmlZMnRxZ0lub1Z5RzU4VnRkalRZWXcrQT0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:00:55Z",
                "eTag": "\"{DD9ED836-18B0-44E7-B352-DA56C300CDCC},2\"",
                "id": "01YNDLPYJW3CPN3MAY45CLGUW2K3BQBTOM",
                "lastModifiedDateTime": "2021-06-19T11:00:55Z",
                "name": "Q3 Sales and Marketing Expense Report Audit.pptx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BDD9ED836-18B0-44E7-B352-DA56C300CDCC%7D&file=Q3%20Sales%20and%20Marketing%20Expense%20Report%20Audit.pptx&action=edit&mobileredirect=true",
                "cTag": "\"c:{DD9ED836-18B0-44E7-B352-DA56C300CDCC},3\"",
                "size": 376325,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                  "hashes": {
                    "quickXorHash": "BCE09zeZNS8dnTx14w3stEm9g0g="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:00:55Z",
                  "lastModifiedDateTime": "2021-06-19T11:00:55Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=232b3df5-7f6c-4f06-b811-6029534bd1db&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6ImRNVEtaaGJuWVgyM3g4dURuQ0tXU0RQVFhjMWVSbGR3aWhrOXF6RjYvczA9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.M3ZVcE1NVnAxd0F4WDFiNUR5eHVoa3lZQ2x1eUUyVGwwdnFUZVNXaTJ1UT0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:44Z",
                "eTag": "\"{232B3DF5-7F6C-4F06-B811-6029534BD1DB},1\"",
                "id": "01YNDLPYPVHUVSG3D7AZH3QELAFFJUXUO3",
                "lastModifiedDateTime": "2021-06-19T11:01:44Z",
                "name": "Q3_Product_Strategy.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B232B3DF5-7F6C-4F06-B811-6029534BD1DB%7D&file=Q3_Product_Strategy.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{232B3DF5-7F6C-4F06-B811-6029534BD1DB},2\"",
                "size": 48473,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "54Y/Q4Ykxgg3N7yWL8iTlokMSS0="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:44Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:44Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=fcd62009-cf87-478e-a788-3853a9832f6b&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6ImVLQ0oyWDhpR3JvdGF0TGQ0d2hyOTNnajBKaUpUeEhUUHlmUWl2clJjOWc9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.bllWYkVsb3doSkxZbmFVTTE1TDdhUlU5NjNzWGNHUWRzbUNiOW1MN2s2az0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:52Z",
                "eTag": "\"{FCD62009-CF87-478E-A788-3853A9832F6B},1\"",
                "id": "01YNDLPYIJEDLPZB6PRZD2PCBYKOUYGL3L",
                "lastModifiedDateTime": "2021-06-19T11:01:52Z",
                "name": "Sales Memo.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BFCD62009-CF87-478E-A788-3853A9832F6B%7D&file=Sales%20Memo.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{FCD62009-CF87-478E-A788-3853A9832F6B},2\"",
                "size": 33737,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "dNTTVcFU+d1H7Ou3seZvQI1pGnk="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:52Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:52Z"
                }
              }
            ]
          });
        default:
          return Promise.reject(`Invalid GET request: ${url}`);
      }
    });

    command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/',
        folderUrl: 'DemoDocs'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        assert(loggerLogSpy.calledWith([
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=b7b2648f-c50c-4d1e-bb99-9c1dcd5e37ae&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkNKTDB2UUp1SVF3enl1YUVHbDI1WVJpejJDdkM5bDdzRktVUFQzU0F2bTg9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.bmVGTE9nbDJ4c1lvem5sQ3pnTXRmb29HaTRUTWVVTlVaYXBkSVp3QWliRT0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:04Z",
            "eTag": "\"{B7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE},2\"",
            "id": "01YNDLPYMPMSZLODGFDZG3XGM4DXGV4N5O",
            "lastModifiedDateTime": "2021-06-19T11:01:04Z",
            "name": "Blog Post preview.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BB7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE%7D&file=Blog%20Post%20preview.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{B7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE},3\"",
            "size": 134144,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "4wybUm/SrJtzssP/YDDKBwKRYd8="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:04Z",
              "lastModifiedDateTime": "2021-06-19T11:01:04Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=4a3864b3-db49-4d6f-a72e-7e3269c44e33&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IjlUSENBN0loaWhMMUJhc2EyUjlneVBMYldaOGcxVS93VTYyRTU1NHU4MVE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.VEVUTXJYNk1mWmpBSFhudjJrVWMxTWQ0UzRCVzFad0F3K0QrV2VwcEVoaz0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:13Z",
            "eTag": "\"{4A3864B3-DB49-4D6F-A72E-7E3269C44E33},3\"",
            "id": "01YNDLPYNTMQ4EUSO3N5G2OLT6GJU4ITRT",
            "lastModifiedDateTime": "2021-06-19T11:01:13Z",
            "name": "Contoso Purchasing Permissions.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B4A3864B3-DB49-4D6F-A72E-7E3269C44E33%7D&file=Contoso%20Purchasing%20Permissions.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{4A3864B3-DB49-4D6F-A72E-7E3269C44E33},6\"",
            "dataLossPrevention": {
              "block": {},
              "notify": {}
            },
            "size": 27514,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "MTvtw1PC9WtEY0wycgD1aBIIX3A="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:13Z",
              "lastModifiedDateTime": "2021-06-19T11:01:13Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=1bcaf9f5-6788-42f1-9b77-ee43df61c4dd&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkozNjY5RnEwblpFWmluYTZHNjJSTTl0cGJCd1lxaWJuV2lING5SMVdDU1E9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.ck96RHVlYjdGMDhQWnhhc3UrZXhReHBaNHJnekcxUmVDUjMwdTNpSTJpND0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:23Z",
            "eTag": "\"{1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD},2\"",
            "id": "01YNDLPYPV7HFBXCDH6FBJW57OIPPWDRG5",
            "lastModifiedDateTime": "2021-06-19T11:01:23Z",
            "name": "Credit Cards.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD%7D&file=Credit%20Cards.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD},3\"",
            "size": 25245,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "xOKzXY19rWzvLGnGcjybqNcMtIc="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:23Z",
              "lastModifiedDateTime": "2021-06-19T11:01:23Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=bc13ae6b-5a75-427d-8915-fcb6e86edd11&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkpiTHU5OHpmTkxCMXhrZ1k1dTlkOUVONXU0UEZjWHpTQ0s0bTg3UU1LcTQ9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.dUFidWp1M1I4N2xwZmhyRDFQRGpZQ0ZQSHBsbVNOcVdFcDAwSjY3U1JGcz0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:32Z",
            "eTag": "\"{BC13AE6B-5A75-427D-8915-FCB6E86EDD11},2\"",
            "id": "01YNDLPYLLVYJ3Y5K2PVBISFP4W3UG5XIR",
            "lastModifiedDateTime": "2021-06-19T11:01:32Z",
            "name": "Customer Accounts.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BBC13AE6B-5A75-427D-8915-FCB6E86EDD11%7D&file=Customer%20Accounts.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{BC13AE6B-5A75-427D-8915-FCB6E86EDD11},3\"",
            "size": 28747,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "7nfr2jjjYOlbSFfYHX72+Ud1o6Y="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:32Z",
              "lastModifiedDateTime": "2021-06-19T11:01:32Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=5148569c-7459-40a8-a400-e56e4668ad62&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6InphTThqdWw5VW4wZlF3WWVTcnRmd1VWV2cyK2YvdXFVQWlzam9xTkdPTEU9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.K2RqWmd2WU5EWTVaWXprelRhUFBQcjZ0NkwxMzZJbnlPMG1QNDYyM3NvOD0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:00:46Z",
            "eTag": "\"{5148569C-7459-40A8-A400-E56E4668AD62},2\"",
            "id": "01YNDLPYM4KZEFCWLUVBAKIAHFNZDGRLLC",
            "lastModifiedDateTime": "2021-06-19T11:00:46Z",
            "name": "Customer Data.xlsx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B5148569C-7459-40A8-A400-E56E4668AD62%7D&file=Customer%20Data.xlsx&action=default&mobileredirect=true",
            "cTag": "\"c:{5148569C-7459-40A8-A400-E56E4668AD62},3\"",
            "size": 17661,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              "hashes": {
                "quickXorHash": "tTrnKAf/Pcq8/NXtvxHuAAUwIUs="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:00:46Z",
              "lastModifiedDateTime": "2021-06-19T11:00:46Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=dd9ed836-18b0-44e7-b352-da56c300cdcc&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6Ik9XN2pqVTBJWkVxUEVLOU1yUXg0ZzgyQjd6ZnBzRWJ5TXNYY011NlpvbXc9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.eHNsek14b0tsQkZmNE9lT3UxWmlZMnRxZ0lub1Z5RzU4VnRkalRZWXcrQT0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:00:55Z",
            "eTag": "\"{DD9ED836-18B0-44E7-B352-DA56C300CDCC},2\"",
            "id": "01YNDLPYJW3CPN3MAY45CLGUW2K3BQBTOM",
            "lastModifiedDateTime": "2021-06-19T11:00:55Z",
            "name": "Q3 Sales and Marketing Expense Report Audit.pptx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BDD9ED836-18B0-44E7-B352-DA56C300CDCC%7D&file=Q3%20Sales%20and%20Marketing%20Expense%20Report%20Audit.pptx&action=edit&mobileredirect=true",
            "cTag": "\"c:{DD9ED836-18B0-44E7-B352-DA56C300CDCC},3\"",
            "size": 376325,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
              "hashes": {
                "quickXorHash": "BCE09zeZNS8dnTx14w3stEm9g0g="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:00:55Z",
              "lastModifiedDateTime": "2021-06-19T11:00:55Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=232b3df5-7f6c-4f06-b811-6029534bd1db&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6ImRNVEtaaGJuWVgyM3g4dURuQ0tXU0RQVFhjMWVSbGR3aWhrOXF6RjYvczA9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.M3ZVcE1NVnAxd0F4WDFiNUR5eHVoa3lZQ2x1eUUyVGwwdnFUZVNXaTJ1UT0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:44Z",
            "eTag": "\"{232B3DF5-7F6C-4F06-B811-6029534BD1DB},1\"",
            "id": "01YNDLPYPVHUVSG3D7AZH3QELAFFJUXUO3",
            "lastModifiedDateTime": "2021-06-19T11:01:44Z",
            "name": "Q3_Product_Strategy.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B232B3DF5-7F6C-4F06-B811-6029534BD1DB%7D&file=Q3_Product_Strategy.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{232B3DF5-7F6C-4F06-B811-6029534BD1DB},2\"",
            "size": 48473,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "54Y/Q4Ykxgg3N7yWL8iTlokMSS0="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:44Z",
              "lastModifiedDateTime": "2021-06-19T11:01:44Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=fcd62009-cf87-478e-a788-3853a9832f6b&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6ImVLQ0oyWDhpR3JvdGF0TGQ0d2hyOTNnajBKaUpUeEhUUHlmUWl2clJjOWc9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.bllWYkVsb3doSkxZbmFVTTE1TDdhUlU5NjNzWGNHUWRzbUNiOW1MN2s2az0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:52Z",
            "eTag": "\"{FCD62009-CF87-478E-A788-3853A9832F6B},1\"",
            "id": "01YNDLPYIJEDLPZB6PRZD2PCBYKOUYGL3L",
            "lastModifiedDateTime": "2021-06-19T11:01:52Z",
            "name": "Sales Memo.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BFCD62009-CF87-478E-A788-3853A9832F6B%7D&file=Sales%20Memo.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{FCD62009-CF87-478E-A788-3853A9832F6B},2\"",
            "size": 33737,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "dNTTVcFU+d1H7Ou3seZvQI1pGnk="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:52Z",
              "lastModifiedDateTime": "2021-06-19T11:01:52Z"
            },
            "lastModifiedByUser": "Provisioning User"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('loads files from the root site collection with trailing slash, document library without space, leading slash, root folder, matching case', done => {
    sinon.stub(request, 'get').callsFake(opts => {
      const url: string = opts.url as string;

      switch (url) {
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/?$select=id':
          return Promise.resolve({
            "id": "contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42"
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42/drives?$select=webUrl,id':
          return Promise.resolve({
            "value": [
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYVs-s_Fc6EaRomQ91r_60hi",
                "webUrl": "https://contoso.sharepoint.com/Big%20list"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                "webUrl": "https://contoso.sharepoint.com/DemoDocs"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYUMhRqN81GDSYJHirLtImTh",
                "webUrl": "https://contoso.sharepoint.com/Shared%20Documents"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYWwU-RsQtl_RJxAlcRhJSYH",
                "webUrl": "https://contoso.sharepoint.com/JTDesignDocs"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYX2VXqL8cJvTbmPdTxwV8t4",
                "webUrl": "https://contoso.sharepoint.com/MCASDemoFiles"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYWaXP0JCVAyRZbch81buwYY",
                "webUrl": "https://contoso.sharepoint.com/RMSDemoLib"
              }
            ]
          });
        case 'https://graph.microsoft.com/v1.0/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root?$select=id':
          return Promise.resolve({
            "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ"
          });
        case 'https://graph.microsoft.com/v1.0/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/items/01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ/children':
          return Promise.resolve({
            "value": [
              {
                "createdDateTime": "2021-09-25T13:30:17Z",
                "eTag": "\"{A2331FD8-FAA8-4C41-A8A0-B6F27FB993AF},1\"",
                "id": "01YNDLPYOYD4Z2FKH2IFGKRIFW6J73TE5P",
                "lastModifiedDateTime": "2021-09-25T13:30:17Z",
                "name": "Fo'lde'r",
                "webUrl": "https://contoso.sharepoint.com/DemoDocs/Fo%27lde%27r",
                "cTag": "\"c:{A2331FD8-FAA8-4C41-A8A0-B6F27FB993AF},0\"",
                "size": 151415,
                "createdBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "802d8bad-2ae1-479a-bbbe-48aca058bc26",
                    "displayName": "MOD Administrator"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "802d8bad-2ae1-479a-bbbe-48aca058bc26",
                    "displayName": "MOD Administrator"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-09-25T13:30:17Z",
                  "lastModifiedDateTime": "2021-09-25T13:30:17Z"
                },
                "folder": {
                  "childCount": 2
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=b7b2648f-c50c-4d1e-bb99-9c1dcd5e37ae&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkNKTDB2UUp1SVF3enl1YUVHbDI1WVJpejJDdkM5bDdzRktVUFQzU0F2bTg9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.bmVGTE9nbDJ4c1lvem5sQ3pnTXRmb29HaTRUTWVVTlVaYXBkSVp3QWliRT0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:04Z",
                "eTag": "\"{B7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE},2\"",
                "id": "01YNDLPYMPMSZLODGFDZG3XGM4DXGV4N5O",
                "lastModifiedDateTime": "2021-06-19T11:01:04Z",
                "name": "Blog Post preview.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BB7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE%7D&file=Blog%20Post%20preview.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{B7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE},3\"",
                "size": 134144,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "4wybUm/SrJtzssP/YDDKBwKRYd8="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:04Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:04Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=4a3864b3-db49-4d6f-a72e-7e3269c44e33&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IjlUSENBN0loaWhMMUJhc2EyUjlneVBMYldaOGcxVS93VTYyRTU1NHU4MVE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.VEVUTXJYNk1mWmpBSFhudjJrVWMxTWQ0UzRCVzFad0F3K0QrV2VwcEVoaz0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:13Z",
                "eTag": "\"{4A3864B3-DB49-4D6F-A72E-7E3269C44E33},3\"",
                "id": "01YNDLPYNTMQ4EUSO3N5G2OLT6GJU4ITRT",
                "lastModifiedDateTime": "2021-06-19T11:01:13Z",
                "name": "Contoso Purchasing Permissions.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B4A3864B3-DB49-4D6F-A72E-7E3269C44E33%7D&file=Contoso%20Purchasing%20Permissions.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{4A3864B3-DB49-4D6F-A72E-7E3269C44E33},6\"",
                "dataLossPrevention": {
                  "block": {},
                  "notify": {}
                },
                "size": 27514,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "MTvtw1PC9WtEY0wycgD1aBIIX3A="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:13Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:13Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=1bcaf9f5-6788-42f1-9b77-ee43df61c4dd&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkozNjY5RnEwblpFWmluYTZHNjJSTTl0cGJCd1lxaWJuV2lING5SMVdDU1E9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.ck96RHVlYjdGMDhQWnhhc3UrZXhReHBaNHJnekcxUmVDUjMwdTNpSTJpND0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:23Z",
                "eTag": "\"{1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD},2\"",
                "id": "01YNDLPYPV7HFBXCDH6FBJW57OIPPWDRG5",
                "lastModifiedDateTime": "2021-06-19T11:01:23Z",
                "name": "Credit Cards.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD%7D&file=Credit%20Cards.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD},3\"",
                "size": 25245,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "xOKzXY19rWzvLGnGcjybqNcMtIc="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:23Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:23Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=bc13ae6b-5a75-427d-8915-fcb6e86edd11&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkpiTHU5OHpmTkxCMXhrZ1k1dTlkOUVONXU0UEZjWHpTQ0s0bTg3UU1LcTQ9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.dUFidWp1M1I4N2xwZmhyRDFQRGpZQ0ZQSHBsbVNOcVdFcDAwSjY3U1JGcz0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:32Z",
                "eTag": "\"{BC13AE6B-5A75-427D-8915-FCB6E86EDD11},2\"",
                "id": "01YNDLPYLLVYJ3Y5K2PVBISFP4W3UG5XIR",
                "lastModifiedDateTime": "2021-06-19T11:01:32Z",
                "name": "Customer Accounts.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BBC13AE6B-5A75-427D-8915-FCB6E86EDD11%7D&file=Customer%20Accounts.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{BC13AE6B-5A75-427D-8915-FCB6E86EDD11},3\"",
                "size": 28747,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "7nfr2jjjYOlbSFfYHX72+Ud1o6Y="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:32Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:32Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=5148569c-7459-40a8-a400-e56e4668ad62&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6InphTThqdWw5VW4wZlF3WWVTcnRmd1VWV2cyK2YvdXFVQWlzam9xTkdPTEU9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.K2RqWmd2WU5EWTVaWXprelRhUFBQcjZ0NkwxMzZJbnlPMG1QNDYyM3NvOD0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:00:46Z",
                "eTag": "\"{5148569C-7459-40A8-A400-E56E4668AD62},2\"",
                "id": "01YNDLPYM4KZEFCWLUVBAKIAHFNZDGRLLC",
                "lastModifiedDateTime": "2021-06-19T11:00:46Z",
                "name": "Customer Data.xlsx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B5148569C-7459-40A8-A400-E56E4668AD62%7D&file=Customer%20Data.xlsx&action=default&mobileredirect=true",
                "cTag": "\"c:{5148569C-7459-40A8-A400-E56E4668AD62},3\"",
                "size": 17661,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                  "hashes": {
                    "quickXorHash": "tTrnKAf/Pcq8/NXtvxHuAAUwIUs="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:00:46Z",
                  "lastModifiedDateTime": "2021-06-19T11:00:46Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=dd9ed836-18b0-44e7-b352-da56c300cdcc&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6Ik9XN2pqVTBJWkVxUEVLOU1yUXg0ZzgyQjd6ZnBzRWJ5TXNYY011NlpvbXc9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.eHNsek14b0tsQkZmNE9lT3UxWmlZMnRxZ0lub1Z5RzU4VnRkalRZWXcrQT0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:00:55Z",
                "eTag": "\"{DD9ED836-18B0-44E7-B352-DA56C300CDCC},2\"",
                "id": "01YNDLPYJW3CPN3MAY45CLGUW2K3BQBTOM",
                "lastModifiedDateTime": "2021-06-19T11:00:55Z",
                "name": "Q3 Sales and Marketing Expense Report Audit.pptx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BDD9ED836-18B0-44E7-B352-DA56C300CDCC%7D&file=Q3%20Sales%20and%20Marketing%20Expense%20Report%20Audit.pptx&action=edit&mobileredirect=true",
                "cTag": "\"c:{DD9ED836-18B0-44E7-B352-DA56C300CDCC},3\"",
                "size": 376325,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                  "hashes": {
                    "quickXorHash": "BCE09zeZNS8dnTx14w3stEm9g0g="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:00:55Z",
                  "lastModifiedDateTime": "2021-06-19T11:00:55Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=232b3df5-7f6c-4f06-b811-6029534bd1db&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6ImRNVEtaaGJuWVgyM3g4dURuQ0tXU0RQVFhjMWVSbGR3aWhrOXF6RjYvczA9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.M3ZVcE1NVnAxd0F4WDFiNUR5eHVoa3lZQ2x1eUUyVGwwdnFUZVNXaTJ1UT0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:44Z",
                "eTag": "\"{232B3DF5-7F6C-4F06-B811-6029534BD1DB},1\"",
                "id": "01YNDLPYPVHUVSG3D7AZH3QELAFFJUXUO3",
                "lastModifiedDateTime": "2021-06-19T11:01:44Z",
                "name": "Q3_Product_Strategy.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B232B3DF5-7F6C-4F06-B811-6029534BD1DB%7D&file=Q3_Product_Strategy.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{232B3DF5-7F6C-4F06-B811-6029534BD1DB},2\"",
                "size": 48473,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "54Y/Q4Ykxgg3N7yWL8iTlokMSS0="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:44Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:44Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=fcd62009-cf87-478e-a788-3853a9832f6b&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6ImVLQ0oyWDhpR3JvdGF0TGQ0d2hyOTNnajBKaUpUeEhUUHlmUWl2clJjOWc9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.bllWYkVsb3doSkxZbmFVTTE1TDdhUlU5NjNzWGNHUWRzbUNiOW1MN2s2az0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:52Z",
                "eTag": "\"{FCD62009-CF87-478E-A788-3853A9832F6B},1\"",
                "id": "01YNDLPYIJEDLPZB6PRZD2PCBYKOUYGL3L",
                "lastModifiedDateTime": "2021-06-19T11:01:52Z",
                "name": "Sales Memo.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BFCD62009-CF87-478E-A788-3853A9832F6B%7D&file=Sales%20Memo.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{FCD62009-CF87-478E-A788-3853A9832F6B},2\"",
                "size": 33737,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "dNTTVcFU+d1H7Ou3seZvQI1pGnk="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:52Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:52Z"
                }
              }
            ]
          });
        default:
          return Promise.reject(`Invalid GET request: ${url}`);
      }
    });

    command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/',
        folderUrl: '/DemoDocs'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        assert(loggerLogSpy.calledWith([
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=b7b2648f-c50c-4d1e-bb99-9c1dcd5e37ae&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkNKTDB2UUp1SVF3enl1YUVHbDI1WVJpejJDdkM5bDdzRktVUFQzU0F2bTg9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.bmVGTE9nbDJ4c1lvem5sQ3pnTXRmb29HaTRUTWVVTlVaYXBkSVp3QWliRT0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:04Z",
            "eTag": "\"{B7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE},2\"",
            "id": "01YNDLPYMPMSZLODGFDZG3XGM4DXGV4N5O",
            "lastModifiedDateTime": "2021-06-19T11:01:04Z",
            "name": "Blog Post preview.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BB7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE%7D&file=Blog%20Post%20preview.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{B7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE},3\"",
            "size": 134144,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "4wybUm/SrJtzssP/YDDKBwKRYd8="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:04Z",
              "lastModifiedDateTime": "2021-06-19T11:01:04Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=4a3864b3-db49-4d6f-a72e-7e3269c44e33&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IjlUSENBN0loaWhMMUJhc2EyUjlneVBMYldaOGcxVS93VTYyRTU1NHU4MVE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.VEVUTXJYNk1mWmpBSFhudjJrVWMxTWQ0UzRCVzFad0F3K0QrV2VwcEVoaz0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:13Z",
            "eTag": "\"{4A3864B3-DB49-4D6F-A72E-7E3269C44E33},3\"",
            "id": "01YNDLPYNTMQ4EUSO3N5G2OLT6GJU4ITRT",
            "lastModifiedDateTime": "2021-06-19T11:01:13Z",
            "name": "Contoso Purchasing Permissions.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B4A3864B3-DB49-4D6F-A72E-7E3269C44E33%7D&file=Contoso%20Purchasing%20Permissions.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{4A3864B3-DB49-4D6F-A72E-7E3269C44E33},6\"",
            "dataLossPrevention": {
              "block": {},
              "notify": {}
            },
            "size": 27514,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "MTvtw1PC9WtEY0wycgD1aBIIX3A="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:13Z",
              "lastModifiedDateTime": "2021-06-19T11:01:13Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=1bcaf9f5-6788-42f1-9b77-ee43df61c4dd&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkozNjY5RnEwblpFWmluYTZHNjJSTTl0cGJCd1lxaWJuV2lING5SMVdDU1E9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.ck96RHVlYjdGMDhQWnhhc3UrZXhReHBaNHJnekcxUmVDUjMwdTNpSTJpND0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:23Z",
            "eTag": "\"{1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD},2\"",
            "id": "01YNDLPYPV7HFBXCDH6FBJW57OIPPWDRG5",
            "lastModifiedDateTime": "2021-06-19T11:01:23Z",
            "name": "Credit Cards.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD%7D&file=Credit%20Cards.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD},3\"",
            "size": 25245,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "xOKzXY19rWzvLGnGcjybqNcMtIc="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:23Z",
              "lastModifiedDateTime": "2021-06-19T11:01:23Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=bc13ae6b-5a75-427d-8915-fcb6e86edd11&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkpiTHU5OHpmTkxCMXhrZ1k1dTlkOUVONXU0UEZjWHpTQ0s0bTg3UU1LcTQ9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.dUFidWp1M1I4N2xwZmhyRDFQRGpZQ0ZQSHBsbVNOcVdFcDAwSjY3U1JGcz0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:32Z",
            "eTag": "\"{BC13AE6B-5A75-427D-8915-FCB6E86EDD11},2\"",
            "id": "01YNDLPYLLVYJ3Y5K2PVBISFP4W3UG5XIR",
            "lastModifiedDateTime": "2021-06-19T11:01:32Z",
            "name": "Customer Accounts.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BBC13AE6B-5A75-427D-8915-FCB6E86EDD11%7D&file=Customer%20Accounts.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{BC13AE6B-5A75-427D-8915-FCB6E86EDD11},3\"",
            "size": 28747,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "7nfr2jjjYOlbSFfYHX72+Ud1o6Y="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:32Z",
              "lastModifiedDateTime": "2021-06-19T11:01:32Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=5148569c-7459-40a8-a400-e56e4668ad62&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6InphTThqdWw5VW4wZlF3WWVTcnRmd1VWV2cyK2YvdXFVQWlzam9xTkdPTEU9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.K2RqWmd2WU5EWTVaWXprelRhUFBQcjZ0NkwxMzZJbnlPMG1QNDYyM3NvOD0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:00:46Z",
            "eTag": "\"{5148569C-7459-40A8-A400-E56E4668AD62},2\"",
            "id": "01YNDLPYM4KZEFCWLUVBAKIAHFNZDGRLLC",
            "lastModifiedDateTime": "2021-06-19T11:00:46Z",
            "name": "Customer Data.xlsx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B5148569C-7459-40A8-A400-E56E4668AD62%7D&file=Customer%20Data.xlsx&action=default&mobileredirect=true",
            "cTag": "\"c:{5148569C-7459-40A8-A400-E56E4668AD62},3\"",
            "size": 17661,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              "hashes": {
                "quickXorHash": "tTrnKAf/Pcq8/NXtvxHuAAUwIUs="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:00:46Z",
              "lastModifiedDateTime": "2021-06-19T11:00:46Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=dd9ed836-18b0-44e7-b352-da56c300cdcc&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6Ik9XN2pqVTBJWkVxUEVLOU1yUXg0ZzgyQjd6ZnBzRWJ5TXNYY011NlpvbXc9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.eHNsek14b0tsQkZmNE9lT3UxWmlZMnRxZ0lub1Z5RzU4VnRkalRZWXcrQT0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:00:55Z",
            "eTag": "\"{DD9ED836-18B0-44E7-B352-DA56C300CDCC},2\"",
            "id": "01YNDLPYJW3CPN3MAY45CLGUW2K3BQBTOM",
            "lastModifiedDateTime": "2021-06-19T11:00:55Z",
            "name": "Q3 Sales and Marketing Expense Report Audit.pptx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BDD9ED836-18B0-44E7-B352-DA56C300CDCC%7D&file=Q3%20Sales%20and%20Marketing%20Expense%20Report%20Audit.pptx&action=edit&mobileredirect=true",
            "cTag": "\"c:{DD9ED836-18B0-44E7-B352-DA56C300CDCC},3\"",
            "size": 376325,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
              "hashes": {
                "quickXorHash": "BCE09zeZNS8dnTx14w3stEm9g0g="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:00:55Z",
              "lastModifiedDateTime": "2021-06-19T11:00:55Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=232b3df5-7f6c-4f06-b811-6029534bd1db&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6ImRNVEtaaGJuWVgyM3g4dURuQ0tXU0RQVFhjMWVSbGR3aWhrOXF6RjYvczA9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.M3ZVcE1NVnAxd0F4WDFiNUR5eHVoa3lZQ2x1eUUyVGwwdnFUZVNXaTJ1UT0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:44Z",
            "eTag": "\"{232B3DF5-7F6C-4F06-B811-6029534BD1DB},1\"",
            "id": "01YNDLPYPVHUVSG3D7AZH3QELAFFJUXUO3",
            "lastModifiedDateTime": "2021-06-19T11:01:44Z",
            "name": "Q3_Product_Strategy.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B232B3DF5-7F6C-4F06-B811-6029534BD1DB%7D&file=Q3_Product_Strategy.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{232B3DF5-7F6C-4F06-B811-6029534BD1DB},2\"",
            "size": 48473,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "54Y/Q4Ykxgg3N7yWL8iTlokMSS0="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:44Z",
              "lastModifiedDateTime": "2021-06-19T11:01:44Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=fcd62009-cf87-478e-a788-3853a9832f6b&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6ImVLQ0oyWDhpR3JvdGF0TGQ0d2hyOTNnajBKaUpUeEhUUHlmUWl2clJjOWc9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.bllWYkVsb3doSkxZbmFVTTE1TDdhUlU5NjNzWGNHUWRzbUNiOW1MN2s2az0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:52Z",
            "eTag": "\"{FCD62009-CF87-478E-A788-3853A9832F6B},1\"",
            "id": "01YNDLPYIJEDLPZB6PRZD2PCBYKOUYGL3L",
            "lastModifiedDateTime": "2021-06-19T11:01:52Z",
            "name": "Sales Memo.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BFCD62009-CF87-478E-A788-3853A9832F6B%7D&file=Sales%20Memo.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{FCD62009-CF87-478E-A788-3853A9832F6B},2\"",
            "size": 33737,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "dNTTVcFU+d1H7Ou3seZvQI1pGnk="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:52Z",
              "lastModifiedDateTime": "2021-06-19T11:01:52Z"
            },
            "lastModifiedByUser": "Provisioning User"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('loads files from the root site collection without trailing slash, document library without space, root folder, different case', done => {
    sinon.stub(request, 'get').callsFake(opts => {
      const url: string = opts.url as string;

      switch (url) {
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/?$select=id':
          return Promise.resolve({
            "id": "contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42"
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42/drives?$select=webUrl,id':
          return Promise.resolve({
            "value": [
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYVs-s_Fc6EaRomQ91r_60hi",
                "webUrl": "https://contoso.sharepoint.com/Big%20list"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                "webUrl": "https://contoso.sharepoint.com/DemoDocs"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYUMhRqN81GDSYJHirLtImTh",
                "webUrl": "https://contoso.sharepoint.com/Shared%20Documents"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYWwU-RsQtl_RJxAlcRhJSYH",
                "webUrl": "https://contoso.sharepoint.com/JTDesignDocs"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYX2VXqL8cJvTbmPdTxwV8t4",
                "webUrl": "https://contoso.sharepoint.com/MCASDemoFiles"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYWaXP0JCVAyRZbch81buwYY",
                "webUrl": "https://contoso.sharepoint.com/RMSDemoLib"
              }
            ]
          });
        case 'https://graph.microsoft.com/v1.0/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root?$select=id':
          return Promise.resolve({
            "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ"
          });
        case 'https://graph.microsoft.com/v1.0/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/items/01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ/children':
          return Promise.resolve({
            "value": [
              {
                "createdDateTime": "2021-09-25T13:30:17Z",
                "eTag": "\"{A2331FD8-FAA8-4C41-A8A0-B6F27FB993AF},1\"",
                "id": "01YNDLPYOYD4Z2FKH2IFGKRIFW6J73TE5P",
                "lastModifiedDateTime": "2021-09-25T13:30:17Z",
                "name": "Fo'lde'r",
                "webUrl": "https://contoso.sharepoint.com/DemoDocs/Fo%27lde%27r",
                "cTag": "\"c:{A2331FD8-FAA8-4C41-A8A0-B6F27FB993AF},0\"",
                "size": 151415,
                "createdBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "802d8bad-2ae1-479a-bbbe-48aca058bc26",
                    "displayName": "MOD Administrator"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "802d8bad-2ae1-479a-bbbe-48aca058bc26",
                    "displayName": "MOD Administrator"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-09-25T13:30:17Z",
                  "lastModifiedDateTime": "2021-09-25T13:30:17Z"
                },
                "folder": {
                  "childCount": 2
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=b7b2648f-c50c-4d1e-bb99-9c1dcd5e37ae&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkNKTDB2UUp1SVF3enl1YUVHbDI1WVJpejJDdkM5bDdzRktVUFQzU0F2bTg9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.bmVGTE9nbDJ4c1lvem5sQ3pnTXRmb29HaTRUTWVVTlVaYXBkSVp3QWliRT0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:04Z",
                "eTag": "\"{B7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE},2\"",
                "id": "01YNDLPYMPMSZLODGFDZG3XGM4DXGV4N5O",
                "lastModifiedDateTime": "2021-06-19T11:01:04Z",
                "name": "Blog Post preview.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BB7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE%7D&file=Blog%20Post%20preview.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{B7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE},3\"",
                "size": 134144,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "4wybUm/SrJtzssP/YDDKBwKRYd8="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:04Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:04Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=4a3864b3-db49-4d6f-a72e-7e3269c44e33&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IjlUSENBN0loaWhMMUJhc2EyUjlneVBMYldaOGcxVS93VTYyRTU1NHU4MVE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.VEVUTXJYNk1mWmpBSFhudjJrVWMxTWQ0UzRCVzFad0F3K0QrV2VwcEVoaz0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:13Z",
                "eTag": "\"{4A3864B3-DB49-4D6F-A72E-7E3269C44E33},3\"",
                "id": "01YNDLPYNTMQ4EUSO3N5G2OLT6GJU4ITRT",
                "lastModifiedDateTime": "2021-06-19T11:01:13Z",
                "name": "Contoso Purchasing Permissions.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B4A3864B3-DB49-4D6F-A72E-7E3269C44E33%7D&file=Contoso%20Purchasing%20Permissions.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{4A3864B3-DB49-4D6F-A72E-7E3269C44E33},6\"",
                "dataLossPrevention": {
                  "block": {},
                  "notify": {}
                },
                "size": 27514,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "MTvtw1PC9WtEY0wycgD1aBIIX3A="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:13Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:13Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=1bcaf9f5-6788-42f1-9b77-ee43df61c4dd&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkozNjY5RnEwblpFWmluYTZHNjJSTTl0cGJCd1lxaWJuV2lING5SMVdDU1E9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.ck96RHVlYjdGMDhQWnhhc3UrZXhReHBaNHJnekcxUmVDUjMwdTNpSTJpND0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:23Z",
                "eTag": "\"{1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD},2\"",
                "id": "01YNDLPYPV7HFBXCDH6FBJW57OIPPWDRG5",
                "lastModifiedDateTime": "2021-06-19T11:01:23Z",
                "name": "Credit Cards.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD%7D&file=Credit%20Cards.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD},3\"",
                "size": 25245,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "xOKzXY19rWzvLGnGcjybqNcMtIc="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:23Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:23Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=bc13ae6b-5a75-427d-8915-fcb6e86edd11&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkpiTHU5OHpmTkxCMXhrZ1k1dTlkOUVONXU0UEZjWHpTQ0s0bTg3UU1LcTQ9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.dUFidWp1M1I4N2xwZmhyRDFQRGpZQ0ZQSHBsbVNOcVdFcDAwSjY3U1JGcz0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:32Z",
                "eTag": "\"{BC13AE6B-5A75-427D-8915-FCB6E86EDD11},2\"",
                "id": "01YNDLPYLLVYJ3Y5K2PVBISFP4W3UG5XIR",
                "lastModifiedDateTime": "2021-06-19T11:01:32Z",
                "name": "Customer Accounts.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BBC13AE6B-5A75-427D-8915-FCB6E86EDD11%7D&file=Customer%20Accounts.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{BC13AE6B-5A75-427D-8915-FCB6E86EDD11},3\"",
                "size": 28747,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "7nfr2jjjYOlbSFfYHX72+Ud1o6Y="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:32Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:32Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=5148569c-7459-40a8-a400-e56e4668ad62&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6InphTThqdWw5VW4wZlF3WWVTcnRmd1VWV2cyK2YvdXFVQWlzam9xTkdPTEU9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.K2RqWmd2WU5EWTVaWXprelRhUFBQcjZ0NkwxMzZJbnlPMG1QNDYyM3NvOD0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:00:46Z",
                "eTag": "\"{5148569C-7459-40A8-A400-E56E4668AD62},2\"",
                "id": "01YNDLPYM4KZEFCWLUVBAKIAHFNZDGRLLC",
                "lastModifiedDateTime": "2021-06-19T11:00:46Z",
                "name": "Customer Data.xlsx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B5148569C-7459-40A8-A400-E56E4668AD62%7D&file=Customer%20Data.xlsx&action=default&mobileredirect=true",
                "cTag": "\"c:{5148569C-7459-40A8-A400-E56E4668AD62},3\"",
                "size": 17661,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                  "hashes": {
                    "quickXorHash": "tTrnKAf/Pcq8/NXtvxHuAAUwIUs="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:00:46Z",
                  "lastModifiedDateTime": "2021-06-19T11:00:46Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=dd9ed836-18b0-44e7-b352-da56c300cdcc&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6Ik9XN2pqVTBJWkVxUEVLOU1yUXg0ZzgyQjd6ZnBzRWJ5TXNYY011NlpvbXc9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.eHNsek14b0tsQkZmNE9lT3UxWmlZMnRxZ0lub1Z5RzU4VnRkalRZWXcrQT0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:00:55Z",
                "eTag": "\"{DD9ED836-18B0-44E7-B352-DA56C300CDCC},2\"",
                "id": "01YNDLPYJW3CPN3MAY45CLGUW2K3BQBTOM",
                "lastModifiedDateTime": "2021-06-19T11:00:55Z",
                "name": "Q3 Sales and Marketing Expense Report Audit.pptx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BDD9ED836-18B0-44E7-B352-DA56C300CDCC%7D&file=Q3%20Sales%20and%20Marketing%20Expense%20Report%20Audit.pptx&action=edit&mobileredirect=true",
                "cTag": "\"c:{DD9ED836-18B0-44E7-B352-DA56C300CDCC},3\"",
                "size": 376325,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                  "hashes": {
                    "quickXorHash": "BCE09zeZNS8dnTx14w3stEm9g0g="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:00:55Z",
                  "lastModifiedDateTime": "2021-06-19T11:00:55Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=232b3df5-7f6c-4f06-b811-6029534bd1db&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6ImRNVEtaaGJuWVgyM3g4dURuQ0tXU0RQVFhjMWVSbGR3aWhrOXF6RjYvczA9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.M3ZVcE1NVnAxd0F4WDFiNUR5eHVoa3lZQ2x1eUUyVGwwdnFUZVNXaTJ1UT0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:44Z",
                "eTag": "\"{232B3DF5-7F6C-4F06-B811-6029534BD1DB},1\"",
                "id": "01YNDLPYPVHUVSG3D7AZH3QELAFFJUXUO3",
                "lastModifiedDateTime": "2021-06-19T11:01:44Z",
                "name": "Q3_Product_Strategy.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B232B3DF5-7F6C-4F06-B811-6029534BD1DB%7D&file=Q3_Product_Strategy.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{232B3DF5-7F6C-4F06-B811-6029534BD1DB},2\"",
                "size": 48473,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "54Y/Q4Ykxgg3N7yWL8iTlokMSS0="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:44Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:44Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=fcd62009-cf87-478e-a788-3853a9832f6b&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6ImVLQ0oyWDhpR3JvdGF0TGQ0d2hyOTNnajBKaUpUeEhUUHlmUWl2clJjOWc9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.bllWYkVsb3doSkxZbmFVTTE1TDdhUlU5NjNzWGNHUWRzbUNiOW1MN2s2az0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:52Z",
                "eTag": "\"{FCD62009-CF87-478E-A788-3853A9832F6B},1\"",
                "id": "01YNDLPYIJEDLPZB6PRZD2PCBYKOUYGL3L",
                "lastModifiedDateTime": "2021-06-19T11:01:52Z",
                "name": "Sales Memo.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BFCD62009-CF87-478E-A788-3853A9832F6B%7D&file=Sales%20Memo.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{FCD62009-CF87-478E-A788-3853A9832F6B},2\"",
                "size": 33737,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "dNTTVcFU+d1H7Ou3seZvQI1pGnk="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:52Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:52Z"
                }
              }
            ]
          });
        default:
          return Promise.reject(`Invalid GET request: ${url}`);
      }
    });

    command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: 'demodocs'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        assert(loggerLogSpy.calledWith([
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=b7b2648f-c50c-4d1e-bb99-9c1dcd5e37ae&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkNKTDB2UUp1SVF3enl1YUVHbDI1WVJpejJDdkM5bDdzRktVUFQzU0F2bTg9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.bmVGTE9nbDJ4c1lvem5sQ3pnTXRmb29HaTRUTWVVTlVaYXBkSVp3QWliRT0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:04Z",
            "eTag": "\"{B7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE},2\"",
            "id": "01YNDLPYMPMSZLODGFDZG3XGM4DXGV4N5O",
            "lastModifiedDateTime": "2021-06-19T11:01:04Z",
            "name": "Blog Post preview.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BB7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE%7D&file=Blog%20Post%20preview.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{B7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE},3\"",
            "size": 134144,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "4wybUm/SrJtzssP/YDDKBwKRYd8="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:04Z",
              "lastModifiedDateTime": "2021-06-19T11:01:04Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=4a3864b3-db49-4d6f-a72e-7e3269c44e33&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IjlUSENBN0loaWhMMUJhc2EyUjlneVBMYldaOGcxVS93VTYyRTU1NHU4MVE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.VEVUTXJYNk1mWmpBSFhudjJrVWMxTWQ0UzRCVzFad0F3K0QrV2VwcEVoaz0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:13Z",
            "eTag": "\"{4A3864B3-DB49-4D6F-A72E-7E3269C44E33},3\"",
            "id": "01YNDLPYNTMQ4EUSO3N5G2OLT6GJU4ITRT",
            "lastModifiedDateTime": "2021-06-19T11:01:13Z",
            "name": "Contoso Purchasing Permissions.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B4A3864B3-DB49-4D6F-A72E-7E3269C44E33%7D&file=Contoso%20Purchasing%20Permissions.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{4A3864B3-DB49-4D6F-A72E-7E3269C44E33},6\"",
            "dataLossPrevention": {
              "block": {},
              "notify": {}
            },
            "size": 27514,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "MTvtw1PC9WtEY0wycgD1aBIIX3A="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:13Z",
              "lastModifiedDateTime": "2021-06-19T11:01:13Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=1bcaf9f5-6788-42f1-9b77-ee43df61c4dd&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkozNjY5RnEwblpFWmluYTZHNjJSTTl0cGJCd1lxaWJuV2lING5SMVdDU1E9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.ck96RHVlYjdGMDhQWnhhc3UrZXhReHBaNHJnekcxUmVDUjMwdTNpSTJpND0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:23Z",
            "eTag": "\"{1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD},2\"",
            "id": "01YNDLPYPV7HFBXCDH6FBJW57OIPPWDRG5",
            "lastModifiedDateTime": "2021-06-19T11:01:23Z",
            "name": "Credit Cards.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD%7D&file=Credit%20Cards.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD},3\"",
            "size": 25245,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "xOKzXY19rWzvLGnGcjybqNcMtIc="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:23Z",
              "lastModifiedDateTime": "2021-06-19T11:01:23Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=bc13ae6b-5a75-427d-8915-fcb6e86edd11&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkpiTHU5OHpmTkxCMXhrZ1k1dTlkOUVONXU0UEZjWHpTQ0s0bTg3UU1LcTQ9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.dUFidWp1M1I4N2xwZmhyRDFQRGpZQ0ZQSHBsbVNOcVdFcDAwSjY3U1JGcz0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:32Z",
            "eTag": "\"{BC13AE6B-5A75-427D-8915-FCB6E86EDD11},2\"",
            "id": "01YNDLPYLLVYJ3Y5K2PVBISFP4W3UG5XIR",
            "lastModifiedDateTime": "2021-06-19T11:01:32Z",
            "name": "Customer Accounts.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BBC13AE6B-5A75-427D-8915-FCB6E86EDD11%7D&file=Customer%20Accounts.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{BC13AE6B-5A75-427D-8915-FCB6E86EDD11},3\"",
            "size": 28747,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "7nfr2jjjYOlbSFfYHX72+Ud1o6Y="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:32Z",
              "lastModifiedDateTime": "2021-06-19T11:01:32Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=5148569c-7459-40a8-a400-e56e4668ad62&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6InphTThqdWw5VW4wZlF3WWVTcnRmd1VWV2cyK2YvdXFVQWlzam9xTkdPTEU9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.K2RqWmd2WU5EWTVaWXprelRhUFBQcjZ0NkwxMzZJbnlPMG1QNDYyM3NvOD0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:00:46Z",
            "eTag": "\"{5148569C-7459-40A8-A400-E56E4668AD62},2\"",
            "id": "01YNDLPYM4KZEFCWLUVBAKIAHFNZDGRLLC",
            "lastModifiedDateTime": "2021-06-19T11:00:46Z",
            "name": "Customer Data.xlsx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B5148569C-7459-40A8-A400-E56E4668AD62%7D&file=Customer%20Data.xlsx&action=default&mobileredirect=true",
            "cTag": "\"c:{5148569C-7459-40A8-A400-E56E4668AD62},3\"",
            "size": 17661,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              "hashes": {
                "quickXorHash": "tTrnKAf/Pcq8/NXtvxHuAAUwIUs="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:00:46Z",
              "lastModifiedDateTime": "2021-06-19T11:00:46Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=dd9ed836-18b0-44e7-b352-da56c300cdcc&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6Ik9XN2pqVTBJWkVxUEVLOU1yUXg0ZzgyQjd6ZnBzRWJ5TXNYY011NlpvbXc9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.eHNsek14b0tsQkZmNE9lT3UxWmlZMnRxZ0lub1Z5RzU4VnRkalRZWXcrQT0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:00:55Z",
            "eTag": "\"{DD9ED836-18B0-44E7-B352-DA56C300CDCC},2\"",
            "id": "01YNDLPYJW3CPN3MAY45CLGUW2K3BQBTOM",
            "lastModifiedDateTime": "2021-06-19T11:00:55Z",
            "name": "Q3 Sales and Marketing Expense Report Audit.pptx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BDD9ED836-18B0-44E7-B352-DA56C300CDCC%7D&file=Q3%20Sales%20and%20Marketing%20Expense%20Report%20Audit.pptx&action=edit&mobileredirect=true",
            "cTag": "\"c:{DD9ED836-18B0-44E7-B352-DA56C300CDCC},3\"",
            "size": 376325,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
              "hashes": {
                "quickXorHash": "BCE09zeZNS8dnTx14w3stEm9g0g="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:00:55Z",
              "lastModifiedDateTime": "2021-06-19T11:00:55Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=232b3df5-7f6c-4f06-b811-6029534bd1db&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6ImRNVEtaaGJuWVgyM3g4dURuQ0tXU0RQVFhjMWVSbGR3aWhrOXF6RjYvczA9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.M3ZVcE1NVnAxd0F4WDFiNUR5eHVoa3lZQ2x1eUUyVGwwdnFUZVNXaTJ1UT0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:44Z",
            "eTag": "\"{232B3DF5-7F6C-4F06-B811-6029534BD1DB},1\"",
            "id": "01YNDLPYPVHUVSG3D7AZH3QELAFFJUXUO3",
            "lastModifiedDateTime": "2021-06-19T11:01:44Z",
            "name": "Q3_Product_Strategy.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B232B3DF5-7F6C-4F06-B811-6029534BD1DB%7D&file=Q3_Product_Strategy.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{232B3DF5-7F6C-4F06-B811-6029534BD1DB},2\"",
            "size": 48473,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "54Y/Q4Ykxgg3N7yWL8iTlokMSS0="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:44Z",
              "lastModifiedDateTime": "2021-06-19T11:01:44Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=fcd62009-cf87-478e-a788-3853a9832f6b&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6ImVLQ0oyWDhpR3JvdGF0TGQ0d2hyOTNnajBKaUpUeEhUUHlmUWl2clJjOWc9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.bllWYkVsb3doSkxZbmFVTTE1TDdhUlU5NjNzWGNHUWRzbUNiOW1MN2s2az0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:52Z",
            "eTag": "\"{FCD62009-CF87-478E-A788-3853A9832F6B},1\"",
            "id": "01YNDLPYIJEDLPZB6PRZD2PCBYKOUYGL3L",
            "lastModifiedDateTime": "2021-06-19T11:01:52Z",
            "name": "Sales Memo.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BFCD62009-CF87-478E-A788-3853A9832F6B%7D&file=Sales%20Memo.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{FCD62009-CF87-478E-A788-3853A9832F6B},2\"",
            "size": 33737,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "dNTTVcFU+d1H7Ou3seZvQI1pGnk="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:52Z",
              "lastModifiedDateTime": "2021-06-19T11:01:52Z"
            },
            "lastModifiedByUser": "Provisioning User"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('loads files from the root site collection without trailing slash, document library with space, root folder', done => {
    sinon.stub(request, 'get').callsFake(opts => {
      const url: string = opts.url as string;

      switch (url) {
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/?$select=id':
          return Promise.resolve({
            "id": "contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42"
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42/drives?$select=webUrl,id':
          return Promise.resolve({
            "value": [
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYVs-s_Fc6EaRomQ91r_60hi",
                "webUrl": "https://contoso.sharepoint.com/Big%20list"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                "webUrl": "https://contoso.sharepoint.com/Demo%20Docs"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYUMhRqN81GDSYJHirLtImTh",
                "webUrl": "https://contoso.sharepoint.com/Shared%20Documents"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYWwU-RsQtl_RJxAlcRhJSYH",
                "webUrl": "https://contoso.sharepoint.com/JTDesignDocs"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYX2VXqL8cJvTbmPdTxwV8t4",
                "webUrl": "https://contoso.sharepoint.com/MCASDemoFiles"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYWaXP0JCVAyRZbch81buwYY",
                "webUrl": "https://contoso.sharepoint.com/RMSDemoLib"
              }
            ]
          });
        case 'https://graph.microsoft.com/v1.0/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root?$select=id':
          return Promise.resolve({
            "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ"
          });
        case 'https://graph.microsoft.com/v1.0/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/items/01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ/children':
          return Promise.resolve({
            "value": [
              {
                "createdDateTime": "2021-09-25T13:30:17Z",
                "eTag": "\"{A2331FD8-FAA8-4C41-A8A0-B6F27FB993AF},1\"",
                "id": "01YNDLPYOYD4Z2FKH2IFGKRIFW6J73TE5P",
                "lastModifiedDateTime": "2021-09-25T13:30:17Z",
                "name": "Fo'lde'r",
                "webUrl": "https://contoso.sharepoint.com/DemoDocs/Fo%27lde%27r",
                "cTag": "\"c:{A2331FD8-FAA8-4C41-A8A0-B6F27FB993AF},0\"",
                "size": 151415,
                "createdBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "802d8bad-2ae1-479a-bbbe-48aca058bc26",
                    "displayName": "MOD Administrator"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "802d8bad-2ae1-479a-bbbe-48aca058bc26",
                    "displayName": "MOD Administrator"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-09-25T13:30:17Z",
                  "lastModifiedDateTime": "2021-09-25T13:30:17Z"
                },
                "folder": {
                  "childCount": 2
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=b7b2648f-c50c-4d1e-bb99-9c1dcd5e37ae&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkNKTDB2UUp1SVF3enl1YUVHbDI1WVJpejJDdkM5bDdzRktVUFQzU0F2bTg9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.bmVGTE9nbDJ4c1lvem5sQ3pnTXRmb29HaTRUTWVVTlVaYXBkSVp3QWliRT0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:04Z",
                "eTag": "\"{B7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE},2\"",
                "id": "01YNDLPYMPMSZLODGFDZG3XGM4DXGV4N5O",
                "lastModifiedDateTime": "2021-06-19T11:01:04Z",
                "name": "Blog Post preview.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BB7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE%7D&file=Blog%20Post%20preview.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{B7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE},3\"",
                "size": 134144,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "4wybUm/SrJtzssP/YDDKBwKRYd8="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:04Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:04Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=4a3864b3-db49-4d6f-a72e-7e3269c44e33&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IjlUSENBN0loaWhMMUJhc2EyUjlneVBMYldaOGcxVS93VTYyRTU1NHU4MVE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.VEVUTXJYNk1mWmpBSFhudjJrVWMxTWQ0UzRCVzFad0F3K0QrV2VwcEVoaz0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:13Z",
                "eTag": "\"{4A3864B3-DB49-4D6F-A72E-7E3269C44E33},3\"",
                "id": "01YNDLPYNTMQ4EUSO3N5G2OLT6GJU4ITRT",
                "lastModifiedDateTime": "2021-06-19T11:01:13Z",
                "name": "Contoso Purchasing Permissions.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B4A3864B3-DB49-4D6F-A72E-7E3269C44E33%7D&file=Contoso%20Purchasing%20Permissions.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{4A3864B3-DB49-4D6F-A72E-7E3269C44E33},6\"",
                "dataLossPrevention": {
                  "block": {},
                  "notify": {}
                },
                "size": 27514,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "MTvtw1PC9WtEY0wycgD1aBIIX3A="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:13Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:13Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=1bcaf9f5-6788-42f1-9b77-ee43df61c4dd&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkozNjY5RnEwblpFWmluYTZHNjJSTTl0cGJCd1lxaWJuV2lING5SMVdDU1E9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.ck96RHVlYjdGMDhQWnhhc3UrZXhReHBaNHJnekcxUmVDUjMwdTNpSTJpND0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:23Z",
                "eTag": "\"{1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD},2\"",
                "id": "01YNDLPYPV7HFBXCDH6FBJW57OIPPWDRG5",
                "lastModifiedDateTime": "2021-06-19T11:01:23Z",
                "name": "Credit Cards.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD%7D&file=Credit%20Cards.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD},3\"",
                "size": 25245,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "xOKzXY19rWzvLGnGcjybqNcMtIc="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:23Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:23Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=bc13ae6b-5a75-427d-8915-fcb6e86edd11&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkpiTHU5OHpmTkxCMXhrZ1k1dTlkOUVONXU0UEZjWHpTQ0s0bTg3UU1LcTQ9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.dUFidWp1M1I4N2xwZmhyRDFQRGpZQ0ZQSHBsbVNOcVdFcDAwSjY3U1JGcz0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:32Z",
                "eTag": "\"{BC13AE6B-5A75-427D-8915-FCB6E86EDD11},2\"",
                "id": "01YNDLPYLLVYJ3Y5K2PVBISFP4W3UG5XIR",
                "lastModifiedDateTime": "2021-06-19T11:01:32Z",
                "name": "Customer Accounts.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BBC13AE6B-5A75-427D-8915-FCB6E86EDD11%7D&file=Customer%20Accounts.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{BC13AE6B-5A75-427D-8915-FCB6E86EDD11},3\"",
                "size": 28747,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "7nfr2jjjYOlbSFfYHX72+Ud1o6Y="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:32Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:32Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=5148569c-7459-40a8-a400-e56e4668ad62&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6InphTThqdWw5VW4wZlF3WWVTcnRmd1VWV2cyK2YvdXFVQWlzam9xTkdPTEU9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.K2RqWmd2WU5EWTVaWXprelRhUFBQcjZ0NkwxMzZJbnlPMG1QNDYyM3NvOD0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:00:46Z",
                "eTag": "\"{5148569C-7459-40A8-A400-E56E4668AD62},2\"",
                "id": "01YNDLPYM4KZEFCWLUVBAKIAHFNZDGRLLC",
                "lastModifiedDateTime": "2021-06-19T11:00:46Z",
                "name": "Customer Data.xlsx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B5148569C-7459-40A8-A400-E56E4668AD62%7D&file=Customer%20Data.xlsx&action=default&mobileredirect=true",
                "cTag": "\"c:{5148569C-7459-40A8-A400-E56E4668AD62},3\"",
                "size": 17661,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                  "hashes": {
                    "quickXorHash": "tTrnKAf/Pcq8/NXtvxHuAAUwIUs="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:00:46Z",
                  "lastModifiedDateTime": "2021-06-19T11:00:46Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=dd9ed836-18b0-44e7-b352-da56c300cdcc&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6Ik9XN2pqVTBJWkVxUEVLOU1yUXg0ZzgyQjd6ZnBzRWJ5TXNYY011NlpvbXc9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.eHNsek14b0tsQkZmNE9lT3UxWmlZMnRxZ0lub1Z5RzU4VnRkalRZWXcrQT0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:00:55Z",
                "eTag": "\"{DD9ED836-18B0-44E7-B352-DA56C300CDCC},2\"",
                "id": "01YNDLPYJW3CPN3MAY45CLGUW2K3BQBTOM",
                "lastModifiedDateTime": "2021-06-19T11:00:55Z",
                "name": "Q3 Sales and Marketing Expense Report Audit.pptx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BDD9ED836-18B0-44E7-B352-DA56C300CDCC%7D&file=Q3%20Sales%20and%20Marketing%20Expense%20Report%20Audit.pptx&action=edit&mobileredirect=true",
                "cTag": "\"c:{DD9ED836-18B0-44E7-B352-DA56C300CDCC},3\"",
                "size": 376325,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                  "hashes": {
                    "quickXorHash": "BCE09zeZNS8dnTx14w3stEm9g0g="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:00:55Z",
                  "lastModifiedDateTime": "2021-06-19T11:00:55Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=232b3df5-7f6c-4f06-b811-6029534bd1db&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6ImRNVEtaaGJuWVgyM3g4dURuQ0tXU0RQVFhjMWVSbGR3aWhrOXF6RjYvczA9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.M3ZVcE1NVnAxd0F4WDFiNUR5eHVoa3lZQ2x1eUUyVGwwdnFUZVNXaTJ1UT0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:44Z",
                "eTag": "\"{232B3DF5-7F6C-4F06-B811-6029534BD1DB},1\"",
                "id": "01YNDLPYPVHUVSG3D7AZH3QELAFFJUXUO3",
                "lastModifiedDateTime": "2021-06-19T11:01:44Z",
                "name": "Q3_Product_Strategy.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B232B3DF5-7F6C-4F06-B811-6029534BD1DB%7D&file=Q3_Product_Strategy.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{232B3DF5-7F6C-4F06-B811-6029534BD1DB},2\"",
                "size": 48473,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "54Y/Q4Ykxgg3N7yWL8iTlokMSS0="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:44Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:44Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=fcd62009-cf87-478e-a788-3853a9832f6b&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6ImVLQ0oyWDhpR3JvdGF0TGQ0d2hyOTNnajBKaUpUeEhUUHlmUWl2clJjOWc9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.bllWYkVsb3doSkxZbmFVTTE1TDdhUlU5NjNzWGNHUWRzbUNiOW1MN2s2az0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:52Z",
                "eTag": "\"{FCD62009-CF87-478E-A788-3853A9832F6B},1\"",
                "id": "01YNDLPYIJEDLPZB6PRZD2PCBYKOUYGL3L",
                "lastModifiedDateTime": "2021-06-19T11:01:52Z",
                "name": "Sales Memo.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BFCD62009-CF87-478E-A788-3853A9832F6B%7D&file=Sales%20Memo.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{FCD62009-CF87-478E-A788-3853A9832F6B},2\"",
                "size": 33737,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "dNTTVcFU+d1H7Ou3seZvQI1pGnk="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:52Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:52Z"
                }
              }
            ]
          });
        default:
          return Promise.reject(`Invalid GET request: ${url}`);
      }
    });

    command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: 'Demo Docs'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        assert(loggerLogSpy.calledWith([
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=b7b2648f-c50c-4d1e-bb99-9c1dcd5e37ae&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkNKTDB2UUp1SVF3enl1YUVHbDI1WVJpejJDdkM5bDdzRktVUFQzU0F2bTg9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.bmVGTE9nbDJ4c1lvem5sQ3pnTXRmb29HaTRUTWVVTlVaYXBkSVp3QWliRT0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:04Z",
            "eTag": "\"{B7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE},2\"",
            "id": "01YNDLPYMPMSZLODGFDZG3XGM4DXGV4N5O",
            "lastModifiedDateTime": "2021-06-19T11:01:04Z",
            "name": "Blog Post preview.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BB7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE%7D&file=Blog%20Post%20preview.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{B7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE},3\"",
            "size": 134144,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "4wybUm/SrJtzssP/YDDKBwKRYd8="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:04Z",
              "lastModifiedDateTime": "2021-06-19T11:01:04Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=4a3864b3-db49-4d6f-a72e-7e3269c44e33&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IjlUSENBN0loaWhMMUJhc2EyUjlneVBMYldaOGcxVS93VTYyRTU1NHU4MVE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.VEVUTXJYNk1mWmpBSFhudjJrVWMxTWQ0UzRCVzFad0F3K0QrV2VwcEVoaz0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:13Z",
            "eTag": "\"{4A3864B3-DB49-4D6F-A72E-7E3269C44E33},3\"",
            "id": "01YNDLPYNTMQ4EUSO3N5G2OLT6GJU4ITRT",
            "lastModifiedDateTime": "2021-06-19T11:01:13Z",
            "name": "Contoso Purchasing Permissions.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B4A3864B3-DB49-4D6F-A72E-7E3269C44E33%7D&file=Contoso%20Purchasing%20Permissions.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{4A3864B3-DB49-4D6F-A72E-7E3269C44E33},6\"",
            "dataLossPrevention": {
              "block": {},
              "notify": {}
            },
            "size": 27514,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "MTvtw1PC9WtEY0wycgD1aBIIX3A="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:13Z",
              "lastModifiedDateTime": "2021-06-19T11:01:13Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=1bcaf9f5-6788-42f1-9b77-ee43df61c4dd&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkozNjY5RnEwblpFWmluYTZHNjJSTTl0cGJCd1lxaWJuV2lING5SMVdDU1E9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.ck96RHVlYjdGMDhQWnhhc3UrZXhReHBaNHJnekcxUmVDUjMwdTNpSTJpND0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:23Z",
            "eTag": "\"{1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD},2\"",
            "id": "01YNDLPYPV7HFBXCDH6FBJW57OIPPWDRG5",
            "lastModifiedDateTime": "2021-06-19T11:01:23Z",
            "name": "Credit Cards.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD%7D&file=Credit%20Cards.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD},3\"",
            "size": 25245,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "xOKzXY19rWzvLGnGcjybqNcMtIc="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:23Z",
              "lastModifiedDateTime": "2021-06-19T11:01:23Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=bc13ae6b-5a75-427d-8915-fcb6e86edd11&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkpiTHU5OHpmTkxCMXhrZ1k1dTlkOUVONXU0UEZjWHpTQ0s0bTg3UU1LcTQ9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.dUFidWp1M1I4N2xwZmhyRDFQRGpZQ0ZQSHBsbVNOcVdFcDAwSjY3U1JGcz0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:32Z",
            "eTag": "\"{BC13AE6B-5A75-427D-8915-FCB6E86EDD11},2\"",
            "id": "01YNDLPYLLVYJ3Y5K2PVBISFP4W3UG5XIR",
            "lastModifiedDateTime": "2021-06-19T11:01:32Z",
            "name": "Customer Accounts.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BBC13AE6B-5A75-427D-8915-FCB6E86EDD11%7D&file=Customer%20Accounts.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{BC13AE6B-5A75-427D-8915-FCB6E86EDD11},3\"",
            "size": 28747,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "7nfr2jjjYOlbSFfYHX72+Ud1o6Y="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:32Z",
              "lastModifiedDateTime": "2021-06-19T11:01:32Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=5148569c-7459-40a8-a400-e56e4668ad62&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6InphTThqdWw5VW4wZlF3WWVTcnRmd1VWV2cyK2YvdXFVQWlzam9xTkdPTEU9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.K2RqWmd2WU5EWTVaWXprelRhUFBQcjZ0NkwxMzZJbnlPMG1QNDYyM3NvOD0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:00:46Z",
            "eTag": "\"{5148569C-7459-40A8-A400-E56E4668AD62},2\"",
            "id": "01YNDLPYM4KZEFCWLUVBAKIAHFNZDGRLLC",
            "lastModifiedDateTime": "2021-06-19T11:00:46Z",
            "name": "Customer Data.xlsx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B5148569C-7459-40A8-A400-E56E4668AD62%7D&file=Customer%20Data.xlsx&action=default&mobileredirect=true",
            "cTag": "\"c:{5148569C-7459-40A8-A400-E56E4668AD62},3\"",
            "size": 17661,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              "hashes": {
                "quickXorHash": "tTrnKAf/Pcq8/NXtvxHuAAUwIUs="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:00:46Z",
              "lastModifiedDateTime": "2021-06-19T11:00:46Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=dd9ed836-18b0-44e7-b352-da56c300cdcc&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6Ik9XN2pqVTBJWkVxUEVLOU1yUXg0ZzgyQjd6ZnBzRWJ5TXNYY011NlpvbXc9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.eHNsek14b0tsQkZmNE9lT3UxWmlZMnRxZ0lub1Z5RzU4VnRkalRZWXcrQT0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:00:55Z",
            "eTag": "\"{DD9ED836-18B0-44E7-B352-DA56C300CDCC},2\"",
            "id": "01YNDLPYJW3CPN3MAY45CLGUW2K3BQBTOM",
            "lastModifiedDateTime": "2021-06-19T11:00:55Z",
            "name": "Q3 Sales and Marketing Expense Report Audit.pptx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BDD9ED836-18B0-44E7-B352-DA56C300CDCC%7D&file=Q3%20Sales%20and%20Marketing%20Expense%20Report%20Audit.pptx&action=edit&mobileredirect=true",
            "cTag": "\"c:{DD9ED836-18B0-44E7-B352-DA56C300CDCC},3\"",
            "size": 376325,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
              "hashes": {
                "quickXorHash": "BCE09zeZNS8dnTx14w3stEm9g0g="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:00:55Z",
              "lastModifiedDateTime": "2021-06-19T11:00:55Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=232b3df5-7f6c-4f06-b811-6029534bd1db&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6ImRNVEtaaGJuWVgyM3g4dURuQ0tXU0RQVFhjMWVSbGR3aWhrOXF6RjYvczA9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.M3ZVcE1NVnAxd0F4WDFiNUR5eHVoa3lZQ2x1eUUyVGwwdnFUZVNXaTJ1UT0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:44Z",
            "eTag": "\"{232B3DF5-7F6C-4F06-B811-6029534BD1DB},1\"",
            "id": "01YNDLPYPVHUVSG3D7AZH3QELAFFJUXUO3",
            "lastModifiedDateTime": "2021-06-19T11:01:44Z",
            "name": "Q3_Product_Strategy.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B232B3DF5-7F6C-4F06-B811-6029534BD1DB%7D&file=Q3_Product_Strategy.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{232B3DF5-7F6C-4F06-B811-6029534BD1DB},2\"",
            "size": 48473,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "54Y/Q4Ykxgg3N7yWL8iTlokMSS0="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:44Z",
              "lastModifiedDateTime": "2021-06-19T11:01:44Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=fcd62009-cf87-478e-a788-3853a9832f6b&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6ImVLQ0oyWDhpR3JvdGF0TGQ0d2hyOTNnajBKaUpUeEhUUHlmUWl2clJjOWc9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.bllWYkVsb3doSkxZbmFVTTE1TDdhUlU5NjNzWGNHUWRzbUNiOW1MN2s2az0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:52Z",
            "eTag": "\"{FCD62009-CF87-478E-A788-3853A9832F6B},1\"",
            "id": "01YNDLPYIJEDLPZB6PRZD2PCBYKOUYGL3L",
            "lastModifiedDateTime": "2021-06-19T11:01:52Z",
            "name": "Sales Memo.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BFCD62009-CF87-478E-A788-3853A9832F6B%7D&file=Sales%20Memo.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{FCD62009-CF87-478E-A788-3853A9832F6B},2\"",
            "size": 33737,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "dNTTVcFU+d1H7Ou3seZvQI1pGnk="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:52Z",
              "lastModifiedDateTime": "2021-06-19T11:01:52Z"
            },
            "lastModifiedByUser": "Provisioning User"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('loads files from the root site collection without trailing slash, document library without space, subfolder', done => {
    sinon.stub(request, 'get').callsFake(opts => {
      const url: string = opts.url as string;

      switch (url) {
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/?$select=id':
          return Promise.resolve({
            "id": "contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42"
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42/drives?$select=webUrl,id':
          return Promise.resolve({
            "value": [
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYVs-s_Fc6EaRomQ91r_60hi",
                "webUrl": "https://contoso.sharepoint.com/Big%20list"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                "webUrl": "https://contoso.sharepoint.com/DemoDocs"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYUMhRqN81GDSYJHirLtImTh",
                "webUrl": "https://contoso.sharepoint.com/Shared%20Documents"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYWwU-RsQtl_RJxAlcRhJSYH",
                "webUrl": "https://contoso.sharepoint.com/JTDesignDocs"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYX2VXqL8cJvTbmPdTxwV8t4",
                "webUrl": "https://contoso.sharepoint.com/MCASDemoFiles"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYWaXP0JCVAyRZbch81buwYY",
                "webUrl": "https://contoso.sharepoint.com/RMSDemoLib"
              }
            ]
          });
        case 'https://graph.microsoft.com/v1.0/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:/Folder?$select=id':
          return Promise.resolve({
            "id": "01YNDLPYOVCWVNMXAYL5DK3JPGEHBJM6KO"
          });
        case 'https://graph.microsoft.com/v1.0/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/items/01YNDLPYOVCWVNMXAYL5DK3JPGEHBJM6KO/children':
          return Promise.resolve({
            "value": [{
              "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=a05f5fb4-6ac7-4ce2-ba39-47376af92b81&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5Njc4NCIsImV4cCI6IjE2MzYzMDAzODQiLCJlbmRwb2ludHVybCI6ImhaRjR2ZzFOVXZ0cFJ3QmNlUnArMXJZaTVEcVA3SWNUUTVuOHA4aWY2K289IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpHWm1Nakl3WkdRdE5HVXpOeTAwT1RCaExXRm1NVEl0WWpWallXSTJPVEkxWXpBMSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiIyMC4xOTAuMTYwLjE2NCJ9.VzM1N0l1azFQVWhJSVU5MDJncDBDM29RTFY4RmYySGs5VG02cEdRQUw2RT0&ApiVersion=2.0",
              "createdDateTime": "2021-11-07T14:52:39Z",
              "eTag": "\"{A05F5FB4-6AC7-4CE2-BA39-47376AF92B81},2\"",
              "id": "01YNDLPYNUL5P2BR3K4JGLUOKHG5VPSK4B",
              "lastModifiedDateTime": "2021-11-07T14:52:39Z",
              "name": "Blog Post preview.docx",
              "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BA05F5FB4-6AC7-4CE2-BA39-47376AF92B81%7D&file=Blog%20Post%20preview.docx&action=default&mobileredirect=true",
              "cTag": "\"c:{A05F5FB4-6AC7-4CE2-BA39-47376AF92B81},3\"",
              "size": 134144,
              "createdBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "802d8bad-2ae1-479a-bbbe-48aca058bc26",
                  "displayName": "MOD Administrator"
                }
              },
              "lastModifiedBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "802d8bad-2ae1-479a-bbbe-48aca058bc26",
                  "displayName": "MOD Administrator"
                }
              },
              "parentReference": {
                "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                "driveType": "documentLibrary",
                "id": "01YNDLPYOVCWVNMXAYL5DK3JPGEHBJM6KO",
                "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:/Folder"
              },
              "file": {
                "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
              },
              "fileSystemInfo": {
                "createdDateTime": "2021-11-07T14:52:39Z",
                "lastModifiedDateTime": "2021-11-07T14:52:39Z"
              }
            }]
          });
        default:
          return Promise.reject(`Invalid GET request: ${url}`);
      }
    });

    command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: 'DemoDocs/Folder'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        assert(loggerLogSpy.calledWith([{
          "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=a05f5fb4-6ac7-4ce2-ba39-47376af92b81&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5Njc4NCIsImV4cCI6IjE2MzYzMDAzODQiLCJlbmRwb2ludHVybCI6ImhaRjR2ZzFOVXZ0cFJ3QmNlUnArMXJZaTVEcVA3SWNUUTVuOHA4aWY2K289IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpHWm1Nakl3WkdRdE5HVXpOeTAwT1RCaExXRm1NVEl0WWpWallXSTJPVEkxWXpBMSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiIyMC4xOTAuMTYwLjE2NCJ9.VzM1N0l1azFQVWhJSVU5MDJncDBDM29RTFY4RmYySGs5VG02cEdRQUw2RT0&ApiVersion=2.0",
          "createdDateTime": "2021-11-07T14:52:39Z",
          "eTag": "\"{A05F5FB4-6AC7-4CE2-BA39-47376AF92B81},2\"",
          "id": "01YNDLPYNUL5P2BR3K4JGLUOKHG5VPSK4B",
          "lastModifiedDateTime": "2021-11-07T14:52:39Z",
          "name": "Blog Post preview.docx",
          "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BA05F5FB4-6AC7-4CE2-BA39-47376AF92B81%7D&file=Blog%20Post%20preview.docx&action=default&mobileredirect=true",
          "cTag": "\"c:{A05F5FB4-6AC7-4CE2-BA39-47376AF92B81},3\"",
          "size": 134144,
          "createdBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "802d8bad-2ae1-479a-bbbe-48aca058bc26",
              "displayName": "MOD Administrator"
            }
          },
          "lastModifiedBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "802d8bad-2ae1-479a-bbbe-48aca058bc26",
              "displayName": "MOD Administrator"
            }
          },
          "parentReference": {
            "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
            "driveType": "documentLibrary",
            "id": "01YNDLPYOVCWVNMXAYL5DK3JPGEHBJM6KO",
            "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:/Folder"
          },
          "file": {
            "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          },
          "fileSystemInfo": {
            "createdDateTime": "2021-11-07T14:52:39Z",
            "lastModifiedDateTime": "2021-11-07T14:52:39Z"
          },
          "lastModifiedByUser": "MOD Administrator"
        }]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('loads files from the root site collection without trailing slash, document library without space, subfolder with special chars', done => {
    sinon.stub(request, 'get').callsFake(opts => {
      const url: string = opts.url as string;

      switch (url) {
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/?$select=id':
          return Promise.resolve({
            "id": "contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42"
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42/drives?$select=webUrl,id':
          return Promise.resolve({
            "value": [
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYVs-s_Fc6EaRomQ91r_60hi",
                "webUrl": "https://contoso.sharepoint.com/Big%20list"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                "webUrl": "https://contoso.sharepoint.com/DemoDocs"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYUMhRqN81GDSYJHirLtImTh",
                "webUrl": "https://contoso.sharepoint.com/Shared%20Documents"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYWwU-RsQtl_RJxAlcRhJSYH",
                "webUrl": "https://contoso.sharepoint.com/JTDesignDocs"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYX2VXqL8cJvTbmPdTxwV8t4",
                "webUrl": "https://contoso.sharepoint.com/MCASDemoFiles"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYWaXP0JCVAyRZbch81buwYY",
                "webUrl": "https://contoso.sharepoint.com/RMSDemoLib"
              }
            ]
          });
        case `https://graph.microsoft.com/v1.0/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:/Fo'lde'r?$select=id`:
          return Promise.resolve({
            "id": "01YNDLPYOYD4Z2FKH2IFGKRIFW6J73TE5P"
          });
        case 'https://graph.microsoft.com/v1.0/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/items/01YNDLPYOYD4Z2FKH2IFGKRIFW6J73TE5P/children':
          return Promise.resolve({
            "value": [{
              "createdDateTime": "2021-09-28T15:01:45Z",
              "eTag": "\"{5BB5F0F7-1E41-4B48-B27E-75065BF9F32E},1\"",
              "id": "01YNDLPYPX6C2VWQI6JBF3E7TVAZN7T4ZO",
              "lastModifiedDateTime": "2021-09-28T15:01:45Z",
              "name": "Subfolder",
              "webUrl": "https://contoso.sharepoint.com/DemoDocs/Fo%27lde%27r/Subfolder",
              "cTag": "\"c:{5BB5F0F7-1E41-4B48-B27E-75065BF9F32E},0\"",
              "size": 17271,
              "createdBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "802d8bad-2ae1-479a-bbbe-48aca058bc26",
                  "displayName": "MOD Administrator"
                }
              },
              "lastModifiedBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "802d8bad-2ae1-479a-bbbe-48aca058bc26",
                  "displayName": "MOD Administrator"
                }
              },
              "parentReference": {
                "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                "driveType": "documentLibrary",
                "id": "01YNDLPYOYD4Z2FKH2IFGKRIFW6J73TE5P",
                "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:/Fo'lde'r"
              },
              "fileSystemInfo": {
                "createdDateTime": "2021-09-28T15:01:45Z",
                "lastModifiedDateTime": "2021-09-28T15:01:45Z"
              },
              "folder": {
                "childCount": 1
              }
            },
            {
              "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=88fa8bc8-0eca-40e5-84c6-8d3974384803&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NzIzOSIsImV4cCI6IjE2MzYzMDA4MzkiLCJlbmRwb2ludHVybCI6ImJUNlVod0k3TDQ5UEsyY3kxTisvWFlaUm5Ra25ZWGlMeEdxcFhmSk11QzA9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlkyUTBZekE1TlRrdE1HRTVaQzAwT1RNMUxXRmxZek10TjJObU5tTXdOamRtT1dVeCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiIyMC4xOTAuMTYwLjk2In0.T0VCSjFjOU1OS05FR2cvUkNrRXdlenhIckxVWDFiRVFxM2VuT2VhdGhMbz0&ApiVersion=2.0",
              "createdDateTime": "2021-09-25T13:30:33Z",
              "eTag": "\"{88FA8BC8-0ECA-40E5-84C6-8D3974384803},2\"",
              "id": "01YNDLPYOIRP5IRSQO4VAIJRUNHF2DQSAD",
              "lastModifiedDateTime": "2021-09-25T13:30:33Z",
              "name": "Blog Post preview.docx",
              "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B88FA8BC8-0ECA-40E5-84C6-8D3974384803%7D&file=Blog%20Post%20preview.docx&action=default&mobileredirect=true",
              "cTag": "\"c:{88FA8BC8-0ECA-40E5-84C6-8D3974384803},3\"",
              "size": 134144,
              "createdBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "802d8bad-2ae1-479a-bbbe-48aca058bc26",
                  "displayName": "MOD Administrator"
                }
              },
              "lastModifiedBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "802d8bad-2ae1-479a-bbbe-48aca058bc26",
                  "displayName": "MOD Administrator"
                }
              },
              "parentReference": {
                "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                "driveType": "documentLibrary",
                "id": "01YNDLPYOYD4Z2FKH2IFGKRIFW6J73TE5P",
                "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:/Fo'lde'r"
              },
              "file": {
                "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
              },
              "fileSystemInfo": {
                "createdDateTime": "2021-09-25T13:30:33Z",
                "lastModifiedDateTime": "2021-09-25T13:30:33Z"
              }
            }]
          });
        default:
          return Promise.reject(`Invalid GET request: ${url}`);
      }
    });

    command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: `DemoDocs/Fo'lde'r`
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        assert(loggerLogSpy.calledWith([{
          "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=88fa8bc8-0eca-40e5-84c6-8d3974384803&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NzIzOSIsImV4cCI6IjE2MzYzMDA4MzkiLCJlbmRwb2ludHVybCI6ImJUNlVod0k3TDQ5UEsyY3kxTisvWFlaUm5Ra25ZWGlMeEdxcFhmSk11QzA9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlkyUTBZekE1TlRrdE1HRTVaQzAwT1RNMUxXRmxZek10TjJObU5tTXdOamRtT1dVeCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiIyMC4xOTAuMTYwLjk2In0.T0VCSjFjOU1OS05FR2cvUkNrRXdlenhIckxVWDFiRVFxM2VuT2VhdGhMbz0&ApiVersion=2.0",
          "createdDateTime": "2021-09-25T13:30:33Z",
          "eTag": "\"{88FA8BC8-0ECA-40E5-84C6-8D3974384803},2\"",
          "id": "01YNDLPYOIRP5IRSQO4VAIJRUNHF2DQSAD",
          "lastModifiedDateTime": "2021-09-25T13:30:33Z",
          "name": "Blog Post preview.docx",
          "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B88FA8BC8-0ECA-40E5-84C6-8D3974384803%7D&file=Blog%20Post%20preview.docx&action=default&mobileredirect=true",
          "cTag": "\"c:{88FA8BC8-0ECA-40E5-84C6-8D3974384803},3\"",
          "size": 134144,
          "createdBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "802d8bad-2ae1-479a-bbbe-48aca058bc26",
              "displayName": "MOD Administrator"
            }
          },
          "lastModifiedBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "802d8bad-2ae1-479a-bbbe-48aca058bc26",
              "displayName": "MOD Administrator"
            }
          },
          "parentReference": {
            "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
            "driveType": "documentLibrary",
            "id": "01YNDLPYOYD4Z2FKH2IFGKRIFW6J73TE5P",
            "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:/Fo'lde'r"
          },
          "file": {
            "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          },
          "fileSystemInfo": {
            "createdDateTime": "2021-09-25T13:30:33Z",
            "lastModifiedDateTime": "2021-09-25T13:30:33Z"
          },
          "lastModifiedByUser": "MOD Administrator"
        }]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('loads files from non-root site collection without trailing slash, document library without space, root folder', done => {
    sinon.stub(request, 'get').callsFake(opts => {
      const url: string = opts.url as string;

      switch (url) {
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/design?$select=id':
          return Promise.resolve({
            "id": "contoso.sharepoint.com,c88ba08f-4e17-49c1-a34f-ca85d908c24c,fde71d4c-ac91-4464-a60f-632d99b8224d"
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,c88ba08f-4e17-49c1-a34f-ca85d908c24c,fde71d4c-ac91-4464-a60f-632d99b8224d/drives?$select=webUrl,id':
          return Promise.resolve({
            "value": [{
              "id": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
              "webUrl": "https://contoso.sharepoint.com/sites/Design/DemoDocs"
            },
            {
              "id": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik0at4LaajDdTo6njHc5dx7g",
              "webUrl": "https://contoso.sharepoint.com/sites/Design/Shared%20Documents"
            }]
          });
        case 'https://graph.microsoft.com/v1.0/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root?$select=id':
          return Promise.resolve({
            "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ"
          });
        case 'https://graph.microsoft.com/v1.0/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/items/01472SYFF6Y2GOVW7725BZO354PWSELRRZ/children':
          return Promise.resolve({
            "value": [{
              "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=1bc151a6-beb8-4034-be93-6e9a18aa6cdc&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6Ik5PNFErS3pZSFBwMTJpZUhiZjdobmUrQ1E0em0yZVVvbnZCczI4RHdZVGs9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.d0hGUFpuMVhJbGZGSVFzcEhPSjJyS1FIUStXbEtLL3RURW9qVUhwZWNKWT0&ApiVersion=2.0",
              "createdDateTime": "2021-11-13T16:27:04Z",
              "eTag": "\"{1BC151A6-BEB8-4034-BE93-6E9A18AA6CDC},1\"",
              "id": "01472SYFFGKHARXOF6GRAL5E3OTIMKU3G4",
              "lastModifiedDateTime": "2021-11-13T16:27:04Z",
              "name": "Blog Post preview.docx",
              "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B1BC151A6-BEB8-4034-BE93-6E9A18AA6CDC%7D&file=Blog%20Post%20preview.docx&action=default&mobileredirect=true",
              "cTag": "\"c:{1BC151A6-BEB8-4034-BE93-6E9A18AA6CDC},4\"",
              "size": 131072,
              "createdBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "lastModifiedBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "parentReference": {
                "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                "driveType": "documentLibrary",
                "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
              },
              "file": {
                "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "hashes": {
                  "quickXorHash": "fVSMG9+U0AczPgoLQYCHu85LOT4="
                }
              },
              "fileSystemInfo": {
                "createdDateTime": "2021-11-13T16:27:04Z",
                "lastModifiedDateTime": "2021-11-13T16:27:04Z"
              }
            },
            {
              "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=716e4cb5-cefc-4562-97bb-b40a985f3306&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6InlrUjVQQjNjU3NIV1hhVXVwZW9qbGtna05UeW9VcFdweGxiS0FZbURwV1k9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.STJ5ekhaUEswQm1uL05Vb0F3dGhLcVJ4ZzNldkpReTM2YWI0cDlEaXo0VT0&ApiVersion=2.0",
              "createdDateTime": "2021-11-13T16:27:05Z",
              "eTag": "\"{716E4CB5-CEFC-4562-97BB-B40A985F3306},1\"",
              "id": "01472SYFFVJRXHD7GOMJCZPO5UBKMF6MYG",
              "lastModifiedDateTime": "2021-11-13T16:27:05Z",
              "name": "Contoso Purchasing Permissions.docx",
              "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B716E4CB5-CEFC-4562-97BB-B40A985F3306%7D&file=Contoso%20Purchasing%20Permissions.docx&action=default&mobileredirect=true",
              "cTag": "\"c:{716E4CB5-CEFC-4562-97BB-B40A985F3306},4\"",
              "size": 27442,
              "createdBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "lastModifiedBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "parentReference": {
                "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                "driveType": "documentLibrary",
                "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
              },
              "file": {
                "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "hashes": {
                  "quickXorHash": "ksnaYSdMJkWPE/EBL9rJWsK7kHw="
                }
              },
              "fileSystemInfo": {
                "createdDateTime": "2021-11-13T16:27:05Z",
                "lastModifiedDateTime": "2021-11-13T16:27:05Z"
              }
            },
            {
              "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=f123cf01-244e-44d3-b2f7-24f5230937a1&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6IkcxMS8vN2M1Y1RyU1JwVEFZOUJSSG1WVUlYYTRNTXdnKzJRcTVFZE1OZ1k9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.RWRBMjZvRFVUdlRkS3JsZmFWb2xDZnJLREtsZ1hGeVEvYWcrWStxV29rST0&ApiVersion=2.0",
              "createdDateTime": "2021-11-13T16:27:04Z",
              "eTag": "\"{F123CF01-244E-44D3-B2F7-24F5230937A1},1\"",
              "id": "01472SYFABZ4R7CTRE2NCLF5ZE6URQSN5B",
              "lastModifiedDateTime": "2021-11-13T16:27:04Z",
              "name": "Credit Cards.docx",
              "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7BF123CF01-244E-44D3-B2F7-24F5230937A1%7D&file=Credit%20Cards.docx&action=default&mobileredirect=true",
              "cTag": "\"c:{F123CF01-244E-44D3-B2F7-24F5230937A1},4\"",
              "size": 24858,
              "createdBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "lastModifiedBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "parentReference": {
                "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                "driveType": "documentLibrary",
                "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
              },
              "file": {
                "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "hashes": {
                  "quickXorHash": "9LEEpPZjBLjJQmN/AKDXRLmXj/k="
                }
              },
              "fileSystemInfo": {
                "createdDateTime": "2021-11-13T16:27:04Z",
                "lastModifiedDateTime": "2021-11-13T16:27:04Z"
              }
            },
            {
              "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=9b9a62cf-e7c6-4fa4-a261-1b260d7d6cf1&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6ImordllIUXB6d1hhM01HV3lJb3JuZWRzZlpXbGZIQXlsSmlsQTE2NU1HMms9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.ck96aFYwWCtSMnFrMVV4QVFNQklhSkMxYzRUczcwTzdxeWVsQnRHWjN1VT0&ApiVersion=2.0",
              "createdDateTime": "2021-11-13T16:27:05Z",
              "eTag": "\"{9B9A62CF-E7C6-4FA4-A261-1B260D7D6CF1},1\"",
              "id": "01472SYFGPMKNJXRXHURH2EYI3EYGX23HR",
              "lastModifiedDateTime": "2021-11-13T16:27:05Z",
              "name": "Customer Accounts.docx",
              "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B9B9A62CF-E7C6-4FA4-A261-1B260D7D6CF1%7D&file=Customer%20Accounts.docx&action=default&mobileredirect=true",
              "cTag": "\"c:{9B9A62CF-E7C6-4FA4-A261-1B260D7D6CF1},4\"",
              "size": 28360,
              "createdBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "lastModifiedBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "parentReference": {
                "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                "driveType": "documentLibrary",
                "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
              },
              "file": {
                "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "hashes": {
                  "quickXorHash": "PTFegvh8VBE1mJ3eELTSb+LSQxI="
                }
              },
              "fileSystemInfo": {
                "createdDateTime": "2021-11-13T16:27:05Z",
                "lastModifiedDateTime": "2021-11-13T16:27:05Z"
              }
            },
            {
              "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=e2f09ed9-987a-4d26-8513-86878f785e0a&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6IkRORDROcGlaQXAxa3RCR21HckRFa0ZScmdEQWVyeVE0TjgyK1dBNWJDaVE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.c2dqT3krQnBMOUx2QXdUMHlKWGtVbGtxMGZldmUyQWVTSFhPeERUR2pSVT0&ApiVersion=2.0",
              "createdDateTime": "2021-11-13T16:27:04Z",
              "eTag": "\"{E2F09ED9-987A-4D26-8513-86878F785E0A},1\"",
              "id": "01472SYFGZT3YOE6UYEZGYKE4GQ6HXQXQK",
              "lastModifiedDateTime": "2021-11-13T16:27:04Z",
              "name": "Customer Data.xlsx",
              "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7BE2F09ED9-987A-4D26-8513-86878F785E0A%7D&file=Customer%20Data.xlsx&action=default&mobileredirect=true",
              "cTag": "\"c:{E2F09ED9-987A-4D26-8513-86878F785E0A},4\"",
              "size": 17448,
              "createdBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "lastModifiedBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "parentReference": {
                "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                "driveType": "documentLibrary",
                "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
              },
              "file": {
                "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "hashes": {
                  "quickXorHash": "hrMOe12SPOwumxGnhdJiYnw1qoM="
                }
              },
              "fileSystemInfo": {
                "createdDateTime": "2021-11-13T16:27:04Z",
                "lastModifiedDateTime": "2021-11-13T16:27:04Z"
              }
            },
            {
              "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=370c45c8-cca6-4333-9cb4-d33b7bc3fc05&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6InAzSHU2TjRHSjRSTW9sREVJOVRLdE9DVzk0SXUzMllyTWdVak9IbCtlWkU9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.K1JabWc3TDNGWHpWN0Q5anluUTN6eUpYajNQNFY0NEJMUFhaS2VGTTgvTT0&ApiVersion=2.0",
              "createdDateTime": "2021-11-13T16:27:05Z",
              "eTag": "\"{370C45C8-CCA6-4333-9CB4-D33B7BC3FC05},1\"",
              "id": "01472SYFGIIUGDPJWMGNBZZNGTHN54H7AF",
              "lastModifiedDateTime": "2021-11-13T16:27:05Z",
              "name": "Q3 Sales and Marketing Expense Report Audit.pptx",
              "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B370C45C8-CCA6-4333-9CB4-D33B7BC3FC05%7D&file=Q3%20Sales%20and%20Marketing%20Expense%20Report%20Audit.pptx&action=edit&mobileredirect=true",
              "cTag": "\"c:{370C45C8-CCA6-4333-9CB4-D33B7BC3FC05},4\"",
              "size": 375938,
              "createdBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "lastModifiedBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "parentReference": {
                "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                "driveType": "documentLibrary",
                "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
              },
              "file": {
                "mimeType": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                "hashes": {
                  "quickXorHash": "/qGrwC7X0bUey5sxCDlzxzSjTLE="
                }
              },
              "fileSystemInfo": {
                "createdDateTime": "2021-11-13T16:27:05Z",
                "lastModifiedDateTime": "2021-11-13T16:27:05Z"
              }
            },
            {
              "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=2223f837-ce90-4638-a2ca-8fd753c4378d&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6InhuVmpzU1NWUWxTWVhWK3dDNElpVmtpMUIrZ0xTYWFpTldTaEpQemFnaFk9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.a0xuN1JrbXlnUWl2cW1rSHhPY1NNcnA4akhxVkVxSnVzQkNWeStLWGlPdz0&ApiVersion=2.0",
              "createdDateTime": "2021-11-13T16:27:04Z",
              "eTag": "\"{2223F837-CE90-4638-A2CA-8FD753C4378D},1\"",
              "id": "01472SYFBX7ARSFEGOHBDKFSUP25J4IN4N",
              "lastModifiedDateTime": "2021-11-13T16:27:04Z",
              "name": "Q3_Product_Strategy.docx",
              "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B2223F837-CE90-4638-A2CA-8FD753C4378D%7D&file=Q3_Product_Strategy.docx&action=default&mobileredirect=true",
              "cTag": "\"c:{2223F837-CE90-4638-A2CA-8FD753C4378D},4\"",
              "size": 48086,
              "createdBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "lastModifiedBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "parentReference": {
                "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                "driveType": "documentLibrary",
                "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
              },
              "file": {
                "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "hashes": {
                  "quickXorHash": "LWe4j5GTM+Sg3PGOV0xoSSCnTr8="
                }
              },
              "fileSystemInfo": {
                "createdDateTime": "2021-11-13T16:27:04Z",
                "lastModifiedDateTime": "2021-11-13T16:27:04Z"
              }
            },
            {
              "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=bfc3a23e-ed0b-4a59-96f1-b11c075d89b0&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6IjVBMkhoblFOcEozMFlsZUp3NDU2ZENCTUlJNXlXbnRoaFhGVGZZQUZuMXM9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.MGd4U0wzSWR6S2grc3F2ODJHQncyRjY2RkkzR1FMdCtpSUo0Y1lhNjZtcz0&ApiVersion=2.0",
              "createdDateTime": "2021-11-13T16:27:05Z",
              "eTag": "\"{BFC3A23E-ED0B-4A59-96F1-B11C075D89B0},1\"",
              "id": "01472SYFB6ULB36C7NLFFJN4NRDQDV3CNQ",
              "lastModifiedDateTime": "2021-11-13T16:27:05Z",
              "name": "Sales Memo.docx",
              "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7BBFC3A23E-ED0B-4A59-96F1-B11C075D89B0%7D&file=Sales%20Memo.docx&action=default&mobileredirect=true",
              "cTag": "\"c:{BFC3A23E-ED0B-4A59-96F1-B11C075D89B0},4\"",
              "size": 35655,
              "createdBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "lastModifiedBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "parentReference": {
                "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                "driveType": "documentLibrary",
                "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
              },
              "file": {
                "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "hashes": {
                  "quickXorHash": "4qr1R9xcSfoofaFA3nMY6qjQXPQ="
                }
              },
              "fileSystemInfo": {
                "createdDateTime": "2021-11-13T16:27:05Z",
                "lastModifiedDateTime": "2021-11-13T16:27:05Z"
              }
            }]
          });
        default:
          return Promise.reject(`Invalid GET request: ${url}`);
      }
    });

    command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/design',
        folderUrl: 'DemoDocs'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        assert(loggerLogSpy.calledWith([{
          "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=1bc151a6-beb8-4034-be93-6e9a18aa6cdc&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6Ik5PNFErS3pZSFBwMTJpZUhiZjdobmUrQ1E0em0yZVVvbnZCczI4RHdZVGs9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.d0hGUFpuMVhJbGZGSVFzcEhPSjJyS1FIUStXbEtLL3RURW9qVUhwZWNKWT0&ApiVersion=2.0",
          "createdDateTime": "2021-11-13T16:27:04Z",
          "eTag": "\"{1BC151A6-BEB8-4034-BE93-6E9A18AA6CDC},1\"",
          "id": "01472SYFFGKHARXOF6GRAL5E3OTIMKU3G4",
          "lastModifiedDateTime": "2021-11-13T16:27:04Z",
          "name": "Blog Post preview.docx",
          "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B1BC151A6-BEB8-4034-BE93-6E9A18AA6CDC%7D&file=Blog%20Post%20preview.docx&action=default&mobileredirect=true",
          "cTag": "\"c:{1BC151A6-BEB8-4034-BE93-6E9A18AA6CDC},4\"",
          "size": 131072,
          "createdBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "lastModifiedBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "parentReference": {
            "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
            "driveType": "documentLibrary",
            "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
            "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
          },
          "file": {
            "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "hashes": {
              "quickXorHash": "fVSMG9+U0AczPgoLQYCHu85LOT4="
            }
          },
          "fileSystemInfo": {
            "createdDateTime": "2021-11-13T16:27:04Z",
            "lastModifiedDateTime": "2021-11-13T16:27:04Z"
          },
          "lastModifiedByUser": "MOD Administrator"
        },
        {
          "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=716e4cb5-cefc-4562-97bb-b40a985f3306&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6InlrUjVQQjNjU3NIV1hhVXVwZW9qbGtna05UeW9VcFdweGxiS0FZbURwV1k9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.STJ5ekhaUEswQm1uL05Vb0F3dGhLcVJ4ZzNldkpReTM2YWI0cDlEaXo0VT0&ApiVersion=2.0",
          "createdDateTime": "2021-11-13T16:27:05Z",
          "eTag": "\"{716E4CB5-CEFC-4562-97BB-B40A985F3306},1\"",
          "id": "01472SYFFVJRXHD7GOMJCZPO5UBKMF6MYG",
          "lastModifiedDateTime": "2021-11-13T16:27:05Z",
          "name": "Contoso Purchasing Permissions.docx",
          "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B716E4CB5-CEFC-4562-97BB-B40A985F3306%7D&file=Contoso%20Purchasing%20Permissions.docx&action=default&mobileredirect=true",
          "cTag": "\"c:{716E4CB5-CEFC-4562-97BB-B40A985F3306},4\"",
          "size": 27442,
          "createdBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "lastModifiedBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "parentReference": {
            "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
            "driveType": "documentLibrary",
            "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
            "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
          },
          "file": {
            "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "hashes": {
              "quickXorHash": "ksnaYSdMJkWPE/EBL9rJWsK7kHw="
            }
          },
          "fileSystemInfo": {
            "createdDateTime": "2021-11-13T16:27:05Z",
            "lastModifiedDateTime": "2021-11-13T16:27:05Z"
          },
          "lastModifiedByUser": "MOD Administrator"
        },
        {
          "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=f123cf01-244e-44d3-b2f7-24f5230937a1&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6IkcxMS8vN2M1Y1RyU1JwVEFZOUJSSG1WVUlYYTRNTXdnKzJRcTVFZE1OZ1k9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.RWRBMjZvRFVUdlRkS3JsZmFWb2xDZnJLREtsZ1hGeVEvYWcrWStxV29rST0&ApiVersion=2.0",
          "createdDateTime": "2021-11-13T16:27:04Z",
          "eTag": "\"{F123CF01-244E-44D3-B2F7-24F5230937A1},1\"",
          "id": "01472SYFABZ4R7CTRE2NCLF5ZE6URQSN5B",
          "lastModifiedDateTime": "2021-11-13T16:27:04Z",
          "name": "Credit Cards.docx",
          "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7BF123CF01-244E-44D3-B2F7-24F5230937A1%7D&file=Credit%20Cards.docx&action=default&mobileredirect=true",
          "cTag": "\"c:{F123CF01-244E-44D3-B2F7-24F5230937A1},4\"",
          "size": 24858,
          "createdBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "lastModifiedBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "parentReference": {
            "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
            "driveType": "documentLibrary",
            "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
            "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
          },
          "file": {
            "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "hashes": {
              "quickXorHash": "9LEEpPZjBLjJQmN/AKDXRLmXj/k="
            }
          },
          "fileSystemInfo": {
            "createdDateTime": "2021-11-13T16:27:04Z",
            "lastModifiedDateTime": "2021-11-13T16:27:04Z"
          },
          "lastModifiedByUser": "MOD Administrator"
        },
        {
          "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=9b9a62cf-e7c6-4fa4-a261-1b260d7d6cf1&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6ImordllIUXB6d1hhM01HV3lJb3JuZWRzZlpXbGZIQXlsSmlsQTE2NU1HMms9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.ck96aFYwWCtSMnFrMVV4QVFNQklhSkMxYzRUczcwTzdxeWVsQnRHWjN1VT0&ApiVersion=2.0",
          "createdDateTime": "2021-11-13T16:27:05Z",
          "eTag": "\"{9B9A62CF-E7C6-4FA4-A261-1B260D7D6CF1},1\"",
          "id": "01472SYFGPMKNJXRXHURH2EYI3EYGX23HR",
          "lastModifiedDateTime": "2021-11-13T16:27:05Z",
          "name": "Customer Accounts.docx",
          "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B9B9A62CF-E7C6-4FA4-A261-1B260D7D6CF1%7D&file=Customer%20Accounts.docx&action=default&mobileredirect=true",
          "cTag": "\"c:{9B9A62CF-E7C6-4FA4-A261-1B260D7D6CF1},4\"",
          "size": 28360,
          "createdBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "lastModifiedBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "parentReference": {
            "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
            "driveType": "documentLibrary",
            "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
            "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
          },
          "file": {
            "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "hashes": {
              "quickXorHash": "PTFegvh8VBE1mJ3eELTSb+LSQxI="
            }
          },
          "fileSystemInfo": {
            "createdDateTime": "2021-11-13T16:27:05Z",
            "lastModifiedDateTime": "2021-11-13T16:27:05Z"
          },
          "lastModifiedByUser": "MOD Administrator"
        },
        {
          "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=e2f09ed9-987a-4d26-8513-86878f785e0a&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6IkRORDROcGlaQXAxa3RCR21HckRFa0ZScmdEQWVyeVE0TjgyK1dBNWJDaVE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.c2dqT3krQnBMOUx2QXdUMHlKWGtVbGtxMGZldmUyQWVTSFhPeERUR2pSVT0&ApiVersion=2.0",
          "createdDateTime": "2021-11-13T16:27:04Z",
          "eTag": "\"{E2F09ED9-987A-4D26-8513-86878F785E0A},1\"",
          "id": "01472SYFGZT3YOE6UYEZGYKE4GQ6HXQXQK",
          "lastModifiedDateTime": "2021-11-13T16:27:04Z",
          "name": "Customer Data.xlsx",
          "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7BE2F09ED9-987A-4D26-8513-86878F785E0A%7D&file=Customer%20Data.xlsx&action=default&mobileredirect=true",
          "cTag": "\"c:{E2F09ED9-987A-4D26-8513-86878F785E0A},4\"",
          "size": 17448,
          "createdBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "lastModifiedBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "parentReference": {
            "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
            "driveType": "documentLibrary",
            "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
            "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
          },
          "file": {
            "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "hashes": {
              "quickXorHash": "hrMOe12SPOwumxGnhdJiYnw1qoM="
            }
          },
          "fileSystemInfo": {
            "createdDateTime": "2021-11-13T16:27:04Z",
            "lastModifiedDateTime": "2021-11-13T16:27:04Z"
          },
          "lastModifiedByUser": "MOD Administrator"
        },
        {
          "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=370c45c8-cca6-4333-9cb4-d33b7bc3fc05&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6InAzSHU2TjRHSjRSTW9sREVJOVRLdE9DVzk0SXUzMllyTWdVak9IbCtlWkU9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.K1JabWc3TDNGWHpWN0Q5anluUTN6eUpYajNQNFY0NEJMUFhaS2VGTTgvTT0&ApiVersion=2.0",
          "createdDateTime": "2021-11-13T16:27:05Z",
          "eTag": "\"{370C45C8-CCA6-4333-9CB4-D33B7BC3FC05},1\"",
          "id": "01472SYFGIIUGDPJWMGNBZZNGTHN54H7AF",
          "lastModifiedDateTime": "2021-11-13T16:27:05Z",
          "name": "Q3 Sales and Marketing Expense Report Audit.pptx",
          "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B370C45C8-CCA6-4333-9CB4-D33B7BC3FC05%7D&file=Q3%20Sales%20and%20Marketing%20Expense%20Report%20Audit.pptx&action=edit&mobileredirect=true",
          "cTag": "\"c:{370C45C8-CCA6-4333-9CB4-D33B7BC3FC05},4\"",
          "size": 375938,
          "createdBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "lastModifiedBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "parentReference": {
            "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
            "driveType": "documentLibrary",
            "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
            "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
          },
          "file": {
            "mimeType": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            "hashes": {
              "quickXorHash": "/qGrwC7X0bUey5sxCDlzxzSjTLE="
            }
          },
          "fileSystemInfo": {
            "createdDateTime": "2021-11-13T16:27:05Z",
            "lastModifiedDateTime": "2021-11-13T16:27:05Z"
          },
          "lastModifiedByUser": "MOD Administrator"
        },
        {
          "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=2223f837-ce90-4638-a2ca-8fd753c4378d&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6InhuVmpzU1NWUWxTWVhWK3dDNElpVmtpMUIrZ0xTYWFpTldTaEpQemFnaFk9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.a0xuN1JrbXlnUWl2cW1rSHhPY1NNcnA4akhxVkVxSnVzQkNWeStLWGlPdz0&ApiVersion=2.0",
          "createdDateTime": "2021-11-13T16:27:04Z",
          "eTag": "\"{2223F837-CE90-4638-A2CA-8FD753C4378D},1\"",
          "id": "01472SYFBX7ARSFEGOHBDKFSUP25J4IN4N",
          "lastModifiedDateTime": "2021-11-13T16:27:04Z",
          "name": "Q3_Product_Strategy.docx",
          "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B2223F837-CE90-4638-A2CA-8FD753C4378D%7D&file=Q3_Product_Strategy.docx&action=default&mobileredirect=true",
          "cTag": "\"c:{2223F837-CE90-4638-A2CA-8FD753C4378D},4\"",
          "size": 48086,
          "createdBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "lastModifiedBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "parentReference": {
            "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
            "driveType": "documentLibrary",
            "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
            "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
          },
          "file": {
            "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "hashes": {
              "quickXorHash": "LWe4j5GTM+Sg3PGOV0xoSSCnTr8="
            }
          },
          "fileSystemInfo": {
            "createdDateTime": "2021-11-13T16:27:04Z",
            "lastModifiedDateTime": "2021-11-13T16:27:04Z"
          },
          "lastModifiedByUser": "MOD Administrator"
        },
        {
          "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=bfc3a23e-ed0b-4a59-96f1-b11c075d89b0&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6IjVBMkhoblFOcEozMFlsZUp3NDU2ZENCTUlJNXlXbnRoaFhGVGZZQUZuMXM9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.MGd4U0wzSWR6S2grc3F2ODJHQncyRjY2RkkzR1FMdCtpSUo0Y1lhNjZtcz0&ApiVersion=2.0",
          "createdDateTime": "2021-11-13T16:27:05Z",
          "eTag": "\"{BFC3A23E-ED0B-4A59-96F1-B11C075D89B0},1\"",
          "id": "01472SYFB6ULB36C7NLFFJN4NRDQDV3CNQ",
          "lastModifiedDateTime": "2021-11-13T16:27:05Z",
          "name": "Sales Memo.docx",
          "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7BBFC3A23E-ED0B-4A59-96F1-B11C075D89B0%7D&file=Sales%20Memo.docx&action=default&mobileredirect=true",
          "cTag": "\"c:{BFC3A23E-ED0B-4A59-96F1-B11C075D89B0},4\"",
          "size": 35655,
          "createdBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "lastModifiedBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "parentReference": {
            "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
            "driveType": "documentLibrary",
            "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
            "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
          },
          "file": {
            "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "hashes": {
              "quickXorHash": "4qr1R9xcSfoofaFA3nMY6qjQXPQ="
            }
          },
          "fileSystemInfo": {
            "createdDateTime": "2021-11-13T16:27:05Z",
            "lastModifiedDateTime": "2021-11-13T16:27:05Z"
          },
          "lastModifiedByUser": "MOD Administrator"
        }]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('loads files from non-root site collection with trailing slash, document library without space, root folder', done => {
    sinon.stub(request, 'get').callsFake(opts => {
      const url: string = opts.url as string;

      switch (url) {
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/design/?$select=id':
          return Promise.resolve({
            "id": "contoso.sharepoint.com,c88ba08f-4e17-49c1-a34f-ca85d908c24c,fde71d4c-ac91-4464-a60f-632d99b8224d"
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,c88ba08f-4e17-49c1-a34f-ca85d908c24c,fde71d4c-ac91-4464-a60f-632d99b8224d/drives?$select=webUrl,id':
          return Promise.resolve({
            "value": [{
              "id": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
              "webUrl": "https://contoso.sharepoint.com/sites/Design/DemoDocs"
            },
            {
              "id": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik0at4LaajDdTo6njHc5dx7g",
              "webUrl": "https://contoso.sharepoint.com/sites/Design/Shared%20Documents"
            }]
          });
        case 'https://graph.microsoft.com/v1.0/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root?$select=id':
          return Promise.resolve({
            "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ"
          });
        case 'https://graph.microsoft.com/v1.0/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/items/01472SYFF6Y2GOVW7725BZO354PWSELRRZ/children':
          return Promise.resolve({
            "value": [{
              "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=1bc151a6-beb8-4034-be93-6e9a18aa6cdc&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6Ik5PNFErS3pZSFBwMTJpZUhiZjdobmUrQ1E0em0yZVVvbnZCczI4RHdZVGs9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.d0hGUFpuMVhJbGZGSVFzcEhPSjJyS1FIUStXbEtLL3RURW9qVUhwZWNKWT0&ApiVersion=2.0",
              "createdDateTime": "2021-11-13T16:27:04Z",
              "eTag": "\"{1BC151A6-BEB8-4034-BE93-6E9A18AA6CDC},1\"",
              "id": "01472SYFFGKHARXOF6GRAL5E3OTIMKU3G4",
              "lastModifiedDateTime": "2021-11-13T16:27:04Z",
              "name": "Blog Post preview.docx",
              "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B1BC151A6-BEB8-4034-BE93-6E9A18AA6CDC%7D&file=Blog%20Post%20preview.docx&action=default&mobileredirect=true",
              "cTag": "\"c:{1BC151A6-BEB8-4034-BE93-6E9A18AA6CDC},4\"",
              "size": 131072,
              "createdBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "lastModifiedBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "parentReference": {
                "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                "driveType": "documentLibrary",
                "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
              },
              "file": {
                "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "hashes": {
                  "quickXorHash": "fVSMG9+U0AczPgoLQYCHu85LOT4="
                }
              },
              "fileSystemInfo": {
                "createdDateTime": "2021-11-13T16:27:04Z",
                "lastModifiedDateTime": "2021-11-13T16:27:04Z"
              }
            },
            {
              "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=716e4cb5-cefc-4562-97bb-b40a985f3306&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6InlrUjVQQjNjU3NIV1hhVXVwZW9qbGtna05UeW9VcFdweGxiS0FZbURwV1k9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.STJ5ekhaUEswQm1uL05Vb0F3dGhLcVJ4ZzNldkpReTM2YWI0cDlEaXo0VT0&ApiVersion=2.0",
              "createdDateTime": "2021-11-13T16:27:05Z",
              "eTag": "\"{716E4CB5-CEFC-4562-97BB-B40A985F3306},1\"",
              "id": "01472SYFFVJRXHD7GOMJCZPO5UBKMF6MYG",
              "lastModifiedDateTime": "2021-11-13T16:27:05Z",
              "name": "Contoso Purchasing Permissions.docx",
              "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B716E4CB5-CEFC-4562-97BB-B40A985F3306%7D&file=Contoso%20Purchasing%20Permissions.docx&action=default&mobileredirect=true",
              "cTag": "\"c:{716E4CB5-CEFC-4562-97BB-B40A985F3306},4\"",
              "size": 27442,
              "createdBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "lastModifiedBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "parentReference": {
                "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                "driveType": "documentLibrary",
                "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
              },
              "file": {
                "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "hashes": {
                  "quickXorHash": "ksnaYSdMJkWPE/EBL9rJWsK7kHw="
                }
              },
              "fileSystemInfo": {
                "createdDateTime": "2021-11-13T16:27:05Z",
                "lastModifiedDateTime": "2021-11-13T16:27:05Z"
              }
            },
            {
              "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=f123cf01-244e-44d3-b2f7-24f5230937a1&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6IkcxMS8vN2M1Y1RyU1JwVEFZOUJSSG1WVUlYYTRNTXdnKzJRcTVFZE1OZ1k9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.RWRBMjZvRFVUdlRkS3JsZmFWb2xDZnJLREtsZ1hGeVEvYWcrWStxV29rST0&ApiVersion=2.0",
              "createdDateTime": "2021-11-13T16:27:04Z",
              "eTag": "\"{F123CF01-244E-44D3-B2F7-24F5230937A1},1\"",
              "id": "01472SYFABZ4R7CTRE2NCLF5ZE6URQSN5B",
              "lastModifiedDateTime": "2021-11-13T16:27:04Z",
              "name": "Credit Cards.docx",
              "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7BF123CF01-244E-44D3-B2F7-24F5230937A1%7D&file=Credit%20Cards.docx&action=default&mobileredirect=true",
              "cTag": "\"c:{F123CF01-244E-44D3-B2F7-24F5230937A1},4\"",
              "size": 24858,
              "createdBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "lastModifiedBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "parentReference": {
                "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                "driveType": "documentLibrary",
                "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
              },
              "file": {
                "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "hashes": {
                  "quickXorHash": "9LEEpPZjBLjJQmN/AKDXRLmXj/k="
                }
              },
              "fileSystemInfo": {
                "createdDateTime": "2021-11-13T16:27:04Z",
                "lastModifiedDateTime": "2021-11-13T16:27:04Z"
              }
            },
            {
              "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=9b9a62cf-e7c6-4fa4-a261-1b260d7d6cf1&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6ImordllIUXB6d1hhM01HV3lJb3JuZWRzZlpXbGZIQXlsSmlsQTE2NU1HMms9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.ck96aFYwWCtSMnFrMVV4QVFNQklhSkMxYzRUczcwTzdxeWVsQnRHWjN1VT0&ApiVersion=2.0",
              "createdDateTime": "2021-11-13T16:27:05Z",
              "eTag": "\"{9B9A62CF-E7C6-4FA4-A261-1B260D7D6CF1},1\"",
              "id": "01472SYFGPMKNJXRXHURH2EYI3EYGX23HR",
              "lastModifiedDateTime": "2021-11-13T16:27:05Z",
              "name": "Customer Accounts.docx",
              "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B9B9A62CF-E7C6-4FA4-A261-1B260D7D6CF1%7D&file=Customer%20Accounts.docx&action=default&mobileredirect=true",
              "cTag": "\"c:{9B9A62CF-E7C6-4FA4-A261-1B260D7D6CF1},4\"",
              "size": 28360,
              "createdBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "lastModifiedBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "parentReference": {
                "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                "driveType": "documentLibrary",
                "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
              },
              "file": {
                "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "hashes": {
                  "quickXorHash": "PTFegvh8VBE1mJ3eELTSb+LSQxI="
                }
              },
              "fileSystemInfo": {
                "createdDateTime": "2021-11-13T16:27:05Z",
                "lastModifiedDateTime": "2021-11-13T16:27:05Z"
              }
            },
            {
              "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=e2f09ed9-987a-4d26-8513-86878f785e0a&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6IkRORDROcGlaQXAxa3RCR21HckRFa0ZScmdEQWVyeVE0TjgyK1dBNWJDaVE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.c2dqT3krQnBMOUx2QXdUMHlKWGtVbGtxMGZldmUyQWVTSFhPeERUR2pSVT0&ApiVersion=2.0",
              "createdDateTime": "2021-11-13T16:27:04Z",
              "eTag": "\"{E2F09ED9-987A-4D26-8513-86878F785E0A},1\"",
              "id": "01472SYFGZT3YOE6UYEZGYKE4GQ6HXQXQK",
              "lastModifiedDateTime": "2021-11-13T16:27:04Z",
              "name": "Customer Data.xlsx",
              "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7BE2F09ED9-987A-4D26-8513-86878F785E0A%7D&file=Customer%20Data.xlsx&action=default&mobileredirect=true",
              "cTag": "\"c:{E2F09ED9-987A-4D26-8513-86878F785E0A},4\"",
              "size": 17448,
              "createdBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "lastModifiedBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "parentReference": {
                "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                "driveType": "documentLibrary",
                "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
              },
              "file": {
                "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "hashes": {
                  "quickXorHash": "hrMOe12SPOwumxGnhdJiYnw1qoM="
                }
              },
              "fileSystemInfo": {
                "createdDateTime": "2021-11-13T16:27:04Z",
                "lastModifiedDateTime": "2021-11-13T16:27:04Z"
              }
            },
            {
              "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=370c45c8-cca6-4333-9cb4-d33b7bc3fc05&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6InAzSHU2TjRHSjRSTW9sREVJOVRLdE9DVzk0SXUzMllyTWdVak9IbCtlWkU9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.K1JabWc3TDNGWHpWN0Q5anluUTN6eUpYajNQNFY0NEJMUFhaS2VGTTgvTT0&ApiVersion=2.0",
              "createdDateTime": "2021-11-13T16:27:05Z",
              "eTag": "\"{370C45C8-CCA6-4333-9CB4-D33B7BC3FC05},1\"",
              "id": "01472SYFGIIUGDPJWMGNBZZNGTHN54H7AF",
              "lastModifiedDateTime": "2021-11-13T16:27:05Z",
              "name": "Q3 Sales and Marketing Expense Report Audit.pptx",
              "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B370C45C8-CCA6-4333-9CB4-D33B7BC3FC05%7D&file=Q3%20Sales%20and%20Marketing%20Expense%20Report%20Audit.pptx&action=edit&mobileredirect=true",
              "cTag": "\"c:{370C45C8-CCA6-4333-9CB4-D33B7BC3FC05},4\"",
              "size": 375938,
              "createdBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "lastModifiedBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "parentReference": {
                "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                "driveType": "documentLibrary",
                "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
              },
              "file": {
                "mimeType": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                "hashes": {
                  "quickXorHash": "/qGrwC7X0bUey5sxCDlzxzSjTLE="
                }
              },
              "fileSystemInfo": {
                "createdDateTime": "2021-11-13T16:27:05Z",
                "lastModifiedDateTime": "2021-11-13T16:27:05Z"
              }
            },
            {
              "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=2223f837-ce90-4638-a2ca-8fd753c4378d&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6InhuVmpzU1NWUWxTWVhWK3dDNElpVmtpMUIrZ0xTYWFpTldTaEpQemFnaFk9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.a0xuN1JrbXlnUWl2cW1rSHhPY1NNcnA4akhxVkVxSnVzQkNWeStLWGlPdz0&ApiVersion=2.0",
              "createdDateTime": "2021-11-13T16:27:04Z",
              "eTag": "\"{2223F837-CE90-4638-A2CA-8FD753C4378D},1\"",
              "id": "01472SYFBX7ARSFEGOHBDKFSUP25J4IN4N",
              "lastModifiedDateTime": "2021-11-13T16:27:04Z",
              "name": "Q3_Product_Strategy.docx",
              "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B2223F837-CE90-4638-A2CA-8FD753C4378D%7D&file=Q3_Product_Strategy.docx&action=default&mobileredirect=true",
              "cTag": "\"c:{2223F837-CE90-4638-A2CA-8FD753C4378D},4\"",
              "size": 48086,
              "createdBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "lastModifiedBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "parentReference": {
                "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                "driveType": "documentLibrary",
                "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
              },
              "file": {
                "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "hashes": {
                  "quickXorHash": "LWe4j5GTM+Sg3PGOV0xoSSCnTr8="
                }
              },
              "fileSystemInfo": {
                "createdDateTime": "2021-11-13T16:27:04Z",
                "lastModifiedDateTime": "2021-11-13T16:27:04Z"
              }
            },
            {
              "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=bfc3a23e-ed0b-4a59-96f1-b11c075d89b0&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6IjVBMkhoblFOcEozMFlsZUp3NDU2ZENCTUlJNXlXbnRoaFhGVGZZQUZuMXM9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.MGd4U0wzSWR6S2grc3F2ODJHQncyRjY2RkkzR1FMdCtpSUo0Y1lhNjZtcz0&ApiVersion=2.0",
              "createdDateTime": "2021-11-13T16:27:05Z",
              "eTag": "\"{BFC3A23E-ED0B-4A59-96F1-B11C075D89B0},1\"",
              "id": "01472SYFB6ULB36C7NLFFJN4NRDQDV3CNQ",
              "lastModifiedDateTime": "2021-11-13T16:27:05Z",
              "name": "Sales Memo.docx",
              "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7BBFC3A23E-ED0B-4A59-96F1-B11C075D89B0%7D&file=Sales%20Memo.docx&action=default&mobileredirect=true",
              "cTag": "\"c:{BFC3A23E-ED0B-4A59-96F1-B11C075D89B0},4\"",
              "size": 35655,
              "createdBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "lastModifiedBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "parentReference": {
                "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                "driveType": "documentLibrary",
                "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
              },
              "file": {
                "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "hashes": {
                  "quickXorHash": "4qr1R9xcSfoofaFA3nMY6qjQXPQ="
                }
              },
              "fileSystemInfo": {
                "createdDateTime": "2021-11-13T16:27:05Z",
                "lastModifiedDateTime": "2021-11-13T16:27:05Z"
              }
            }]
          });
        default:
          return Promise.reject(`Invalid GET request: ${url}`);
      }
    });

    command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/design/',
        folderUrl: 'DemoDocs'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        assert(loggerLogSpy.calledWith([{
          "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=1bc151a6-beb8-4034-be93-6e9a18aa6cdc&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6Ik5PNFErS3pZSFBwMTJpZUhiZjdobmUrQ1E0em0yZVVvbnZCczI4RHdZVGs9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.d0hGUFpuMVhJbGZGSVFzcEhPSjJyS1FIUStXbEtLL3RURW9qVUhwZWNKWT0&ApiVersion=2.0",
          "createdDateTime": "2021-11-13T16:27:04Z",
          "eTag": "\"{1BC151A6-BEB8-4034-BE93-6E9A18AA6CDC},1\"",
          "id": "01472SYFFGKHARXOF6GRAL5E3OTIMKU3G4",
          "lastModifiedDateTime": "2021-11-13T16:27:04Z",
          "name": "Blog Post preview.docx",
          "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B1BC151A6-BEB8-4034-BE93-6E9A18AA6CDC%7D&file=Blog%20Post%20preview.docx&action=default&mobileredirect=true",
          "cTag": "\"c:{1BC151A6-BEB8-4034-BE93-6E9A18AA6CDC},4\"",
          "size": 131072,
          "createdBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "lastModifiedBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "parentReference": {
            "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
            "driveType": "documentLibrary",
            "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
            "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
          },
          "file": {
            "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "hashes": {
              "quickXorHash": "fVSMG9+U0AczPgoLQYCHu85LOT4="
            }
          },
          "fileSystemInfo": {
            "createdDateTime": "2021-11-13T16:27:04Z",
            "lastModifiedDateTime": "2021-11-13T16:27:04Z"
          },
          "lastModifiedByUser": "MOD Administrator"
        },
        {
          "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=716e4cb5-cefc-4562-97bb-b40a985f3306&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6InlrUjVQQjNjU3NIV1hhVXVwZW9qbGtna05UeW9VcFdweGxiS0FZbURwV1k9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.STJ5ekhaUEswQm1uL05Vb0F3dGhLcVJ4ZzNldkpReTM2YWI0cDlEaXo0VT0&ApiVersion=2.0",
          "createdDateTime": "2021-11-13T16:27:05Z",
          "eTag": "\"{716E4CB5-CEFC-4562-97BB-B40A985F3306},1\"",
          "id": "01472SYFFVJRXHD7GOMJCZPO5UBKMF6MYG",
          "lastModifiedDateTime": "2021-11-13T16:27:05Z",
          "name": "Contoso Purchasing Permissions.docx",
          "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B716E4CB5-CEFC-4562-97BB-B40A985F3306%7D&file=Contoso%20Purchasing%20Permissions.docx&action=default&mobileredirect=true",
          "cTag": "\"c:{716E4CB5-CEFC-4562-97BB-B40A985F3306},4\"",
          "size": 27442,
          "createdBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "lastModifiedBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "parentReference": {
            "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
            "driveType": "documentLibrary",
            "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
            "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
          },
          "file": {
            "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "hashes": {
              "quickXorHash": "ksnaYSdMJkWPE/EBL9rJWsK7kHw="
            }
          },
          "fileSystemInfo": {
            "createdDateTime": "2021-11-13T16:27:05Z",
            "lastModifiedDateTime": "2021-11-13T16:27:05Z"
          },
          "lastModifiedByUser": "MOD Administrator"
        },
        {
          "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=f123cf01-244e-44d3-b2f7-24f5230937a1&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6IkcxMS8vN2M1Y1RyU1JwVEFZOUJSSG1WVUlYYTRNTXdnKzJRcTVFZE1OZ1k9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.RWRBMjZvRFVUdlRkS3JsZmFWb2xDZnJLREtsZ1hGeVEvYWcrWStxV29rST0&ApiVersion=2.0",
          "createdDateTime": "2021-11-13T16:27:04Z",
          "eTag": "\"{F123CF01-244E-44D3-B2F7-24F5230937A1},1\"",
          "id": "01472SYFABZ4R7CTRE2NCLF5ZE6URQSN5B",
          "lastModifiedDateTime": "2021-11-13T16:27:04Z",
          "name": "Credit Cards.docx",
          "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7BF123CF01-244E-44D3-B2F7-24F5230937A1%7D&file=Credit%20Cards.docx&action=default&mobileredirect=true",
          "cTag": "\"c:{F123CF01-244E-44D3-B2F7-24F5230937A1},4\"",
          "size": 24858,
          "createdBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "lastModifiedBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "parentReference": {
            "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
            "driveType": "documentLibrary",
            "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
            "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
          },
          "file": {
            "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "hashes": {
              "quickXorHash": "9LEEpPZjBLjJQmN/AKDXRLmXj/k="
            }
          },
          "fileSystemInfo": {
            "createdDateTime": "2021-11-13T16:27:04Z",
            "lastModifiedDateTime": "2021-11-13T16:27:04Z"
          },
          "lastModifiedByUser": "MOD Administrator"
        },
        {
          "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=9b9a62cf-e7c6-4fa4-a261-1b260d7d6cf1&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6ImordllIUXB6d1hhM01HV3lJb3JuZWRzZlpXbGZIQXlsSmlsQTE2NU1HMms9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.ck96aFYwWCtSMnFrMVV4QVFNQklhSkMxYzRUczcwTzdxeWVsQnRHWjN1VT0&ApiVersion=2.0",
          "createdDateTime": "2021-11-13T16:27:05Z",
          "eTag": "\"{9B9A62CF-E7C6-4FA4-A261-1B260D7D6CF1},1\"",
          "id": "01472SYFGPMKNJXRXHURH2EYI3EYGX23HR",
          "lastModifiedDateTime": "2021-11-13T16:27:05Z",
          "name": "Customer Accounts.docx",
          "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B9B9A62CF-E7C6-4FA4-A261-1B260D7D6CF1%7D&file=Customer%20Accounts.docx&action=default&mobileredirect=true",
          "cTag": "\"c:{9B9A62CF-E7C6-4FA4-A261-1B260D7D6CF1},4\"",
          "size": 28360,
          "createdBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "lastModifiedBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "parentReference": {
            "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
            "driveType": "documentLibrary",
            "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
            "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
          },
          "file": {
            "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "hashes": {
              "quickXorHash": "PTFegvh8VBE1mJ3eELTSb+LSQxI="
            }
          },
          "fileSystemInfo": {
            "createdDateTime": "2021-11-13T16:27:05Z",
            "lastModifiedDateTime": "2021-11-13T16:27:05Z"
          },
          "lastModifiedByUser": "MOD Administrator"
        },
        {
          "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=e2f09ed9-987a-4d26-8513-86878f785e0a&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6IkRORDROcGlaQXAxa3RCR21HckRFa0ZScmdEQWVyeVE0TjgyK1dBNWJDaVE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.c2dqT3krQnBMOUx2QXdUMHlKWGtVbGtxMGZldmUyQWVTSFhPeERUR2pSVT0&ApiVersion=2.0",
          "createdDateTime": "2021-11-13T16:27:04Z",
          "eTag": "\"{E2F09ED9-987A-4D26-8513-86878F785E0A},1\"",
          "id": "01472SYFGZT3YOE6UYEZGYKE4GQ6HXQXQK",
          "lastModifiedDateTime": "2021-11-13T16:27:04Z",
          "name": "Customer Data.xlsx",
          "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7BE2F09ED9-987A-4D26-8513-86878F785E0A%7D&file=Customer%20Data.xlsx&action=default&mobileredirect=true",
          "cTag": "\"c:{E2F09ED9-987A-4D26-8513-86878F785E0A},4\"",
          "size": 17448,
          "createdBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "lastModifiedBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "parentReference": {
            "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
            "driveType": "documentLibrary",
            "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
            "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
          },
          "file": {
            "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "hashes": {
              "quickXorHash": "hrMOe12SPOwumxGnhdJiYnw1qoM="
            }
          },
          "fileSystemInfo": {
            "createdDateTime": "2021-11-13T16:27:04Z",
            "lastModifiedDateTime": "2021-11-13T16:27:04Z"
          },
          "lastModifiedByUser": "MOD Administrator"
        },
        {
          "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=370c45c8-cca6-4333-9cb4-d33b7bc3fc05&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6InAzSHU2TjRHSjRSTW9sREVJOVRLdE9DVzk0SXUzMllyTWdVak9IbCtlWkU9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.K1JabWc3TDNGWHpWN0Q5anluUTN6eUpYajNQNFY0NEJMUFhaS2VGTTgvTT0&ApiVersion=2.0",
          "createdDateTime": "2021-11-13T16:27:05Z",
          "eTag": "\"{370C45C8-CCA6-4333-9CB4-D33B7BC3FC05},1\"",
          "id": "01472SYFGIIUGDPJWMGNBZZNGTHN54H7AF",
          "lastModifiedDateTime": "2021-11-13T16:27:05Z",
          "name": "Q3 Sales and Marketing Expense Report Audit.pptx",
          "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B370C45C8-CCA6-4333-9CB4-D33B7BC3FC05%7D&file=Q3%20Sales%20and%20Marketing%20Expense%20Report%20Audit.pptx&action=edit&mobileredirect=true",
          "cTag": "\"c:{370C45C8-CCA6-4333-9CB4-D33B7BC3FC05},4\"",
          "size": 375938,
          "createdBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "lastModifiedBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "parentReference": {
            "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
            "driveType": "documentLibrary",
            "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
            "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
          },
          "file": {
            "mimeType": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            "hashes": {
              "quickXorHash": "/qGrwC7X0bUey5sxCDlzxzSjTLE="
            }
          },
          "fileSystemInfo": {
            "createdDateTime": "2021-11-13T16:27:05Z",
            "lastModifiedDateTime": "2021-11-13T16:27:05Z"
          },
          "lastModifiedByUser": "MOD Administrator"
        },
        {
          "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=2223f837-ce90-4638-a2ca-8fd753c4378d&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6InhuVmpzU1NWUWxTWVhWK3dDNElpVmtpMUIrZ0xTYWFpTldTaEpQemFnaFk9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.a0xuN1JrbXlnUWl2cW1rSHhPY1NNcnA4akhxVkVxSnVzQkNWeStLWGlPdz0&ApiVersion=2.0",
          "createdDateTime": "2021-11-13T16:27:04Z",
          "eTag": "\"{2223F837-CE90-4638-A2CA-8FD753C4378D},1\"",
          "id": "01472SYFBX7ARSFEGOHBDKFSUP25J4IN4N",
          "lastModifiedDateTime": "2021-11-13T16:27:04Z",
          "name": "Q3_Product_Strategy.docx",
          "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B2223F837-CE90-4638-A2CA-8FD753C4378D%7D&file=Q3_Product_Strategy.docx&action=default&mobileredirect=true",
          "cTag": "\"c:{2223F837-CE90-4638-A2CA-8FD753C4378D},4\"",
          "size": 48086,
          "createdBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "lastModifiedBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "parentReference": {
            "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
            "driveType": "documentLibrary",
            "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
            "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
          },
          "file": {
            "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "hashes": {
              "quickXorHash": "LWe4j5GTM+Sg3PGOV0xoSSCnTr8="
            }
          },
          "fileSystemInfo": {
            "createdDateTime": "2021-11-13T16:27:04Z",
            "lastModifiedDateTime": "2021-11-13T16:27:04Z"
          },
          "lastModifiedByUser": "MOD Administrator"
        },
        {
          "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=bfc3a23e-ed0b-4a59-96f1-b11c075d89b0&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6IjVBMkhoblFOcEozMFlsZUp3NDU2ZENCTUlJNXlXbnRoaFhGVGZZQUZuMXM9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.MGd4U0wzSWR6S2grc3F2ODJHQncyRjY2RkkzR1FMdCtpSUo0Y1lhNjZtcz0&ApiVersion=2.0",
          "createdDateTime": "2021-11-13T16:27:05Z",
          "eTag": "\"{BFC3A23E-ED0B-4A59-96F1-B11C075D89B0},1\"",
          "id": "01472SYFB6ULB36C7NLFFJN4NRDQDV3CNQ",
          "lastModifiedDateTime": "2021-11-13T16:27:05Z",
          "name": "Sales Memo.docx",
          "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7BBFC3A23E-ED0B-4A59-96F1-B11C075D89B0%7D&file=Sales%20Memo.docx&action=default&mobileredirect=true",
          "cTag": "\"c:{BFC3A23E-ED0B-4A59-96F1-B11C075D89B0},4\"",
          "size": 35655,
          "createdBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "lastModifiedBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "parentReference": {
            "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
            "driveType": "documentLibrary",
            "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
            "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
          },
          "file": {
            "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "hashes": {
              "quickXorHash": "4qr1R9xcSfoofaFA3nMY6qjQXPQ="
            }
          },
          "fileSystemInfo": {
            "createdDateTime": "2021-11-13T16:27:05Z",
            "lastModifiedDateTime": "2021-11-13T16:27:05Z"
          },
          "lastModifiedByUser": "MOD Administrator"
        }]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('loads files from non-root site collection without trailing slash, document library without space, subfolder', done => {
    sinon.stub(request, 'get').callsFake(opts => {
      const url: string = opts.url as string;

      switch (url) {
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/design?$select=id':
          return Promise.resolve({
            "id": "contoso.sharepoint.com,c88ba08f-4e17-49c1-a34f-ca85d908c24c,fde71d4c-ac91-4464-a60f-632d99b8224d"
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,c88ba08f-4e17-49c1-a34f-ca85d908c24c,fde71d4c-ac91-4464-a60f-632d99b8224d/drives?$select=webUrl,id':
          return Promise.resolve({
            "value": [{
              "id": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
              "webUrl": "https://contoso.sharepoint.com/sites/Design/DemoDocs"
            },
            {
              "id": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik0at4LaajDdTo6njHc5dx7g",
              "webUrl": "https://contoso.sharepoint.com/sites/Design/Shared%20Documents"
            }]
          });
        case 'https://graph.microsoft.com/v1.0/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:/Folder?$select=id':
          return Promise.resolve({
            "id": "01472SYFERKSBJFD7X35H3TUHUSHXZBRRD"
          });
        case 'https://graph.microsoft.com/v1.0/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/items/01472SYFERKSBJFD7X35H3TUHUSHXZBRRD/children':
          return Promise.resolve({
            "value": [{
              "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=77e5e9f4-4731-478e-82ae-6eece079ae8b&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMTY1MiIsImV4cCI6IjE2MzY4MjUyNTIiLCJlbmRwb2ludHVybCI6IlFWaVJuSEVIRWRHNW5vSHhrOUFwQ3lXS3JKa05uL3pFVFhqT3NGWGpmQkE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1EZzRNV1UyTnpjdE1EUmpOeTAwWVdOaExXRXpZV1V0TWpNeE56WTJaR0UzTkdZeSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.M0xocUpuM3V6RStHbU5NRnBsQTBVK09SU3Jya3o5SmZZU0UyeE9sNTR6cz0&ApiVersion=2.0",
              "createdDateTime": "2021-11-13T16:40:39Z",
              "eTag": "\"{77E5E9F4-4731-478E-82AE-6EECE079AE8B},2\"",
              "id": "01472SYFHU5HSXOMKHRZDYFLTO5TQHTLUL",
              "lastModifiedDateTime": "2021-11-13T16:40:39Z",
              "name": "Blog Post preview.docx",
              "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B77E5E9F4-4731-478E-82AE-6EECE079AE8B%7D&file=Blog%20Post%20preview.docx&action=default&mobileredirect=true",
              "cTag": "\"c:{77E5E9F4-4731-478E-82AE-6EECE079AE8B},4\"",
              "size": 131072,
              "createdBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "lastModifiedBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "parentReference": {
                "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                "driveType": "documentLibrary",
                "id": "01472SYFERKSBJFD7X35H3TUHUSHXZBRRD",
                "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:/Folder"
              },
              "file": {
                "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
              },
              "fileSystemInfo": {
                "createdDateTime": "2021-11-13T16:40:39Z",
                "lastModifiedDateTime": "2021-11-13T16:40:39Z"
              }
            }]
          });
        default:
          return Promise.reject(`Invalid GET request: ${url}`);
      }
    });

    command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/design',
        folderUrl: 'DemoDocs/Folder'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        assert(loggerLogSpy.calledWith([{
          "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=77e5e9f4-4731-478e-82ae-6eece079ae8b&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMTY1MiIsImV4cCI6IjE2MzY4MjUyNTIiLCJlbmRwb2ludHVybCI6IlFWaVJuSEVIRWRHNW5vSHhrOUFwQ3lXS3JKa05uL3pFVFhqT3NGWGpmQkE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1EZzRNV1UyTnpjdE1EUmpOeTAwWVdOaExXRXpZV1V0TWpNeE56WTJaR0UzTkdZeSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.M0xocUpuM3V6RStHbU5NRnBsQTBVK09SU3Jya3o5SmZZU0UyeE9sNTR6cz0&ApiVersion=2.0",
          "createdDateTime": "2021-11-13T16:40:39Z",
          "eTag": "\"{77E5E9F4-4731-478E-82AE-6EECE079AE8B},2\"",
          "id": "01472SYFHU5HSXOMKHRZDYFLTO5TQHTLUL",
          "lastModifiedDateTime": "2021-11-13T16:40:39Z",
          "name": "Blog Post preview.docx",
          "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B77E5E9F4-4731-478E-82AE-6EECE079AE8B%7D&file=Blog%20Post%20preview.docx&action=default&mobileredirect=true",
          "cTag": "\"c:{77E5E9F4-4731-478E-82AE-6EECE079AE8B},4\"",
          "size": 131072,
          "createdBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "lastModifiedBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "parentReference": {
            "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
            "driveType": "documentLibrary",
            "id": "01472SYFERKSBJFD7X35H3TUHUSHXZBRRD",
            "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:/Folder"
          },
          "file": {
            "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          },
          "fileSystemInfo": {
            "createdDateTime": "2021-11-13T16:40:39Z",
            "lastModifiedDateTime": "2021-11-13T16:40:39Z"
          },
          "lastModifiedByUser": "MOD Administrator"
        }]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('loads files recursively from a folder without subfolders', done => {
    sinon.stub(request, 'get').callsFake(opts => {
      const url: string = opts.url as string;

      switch (url) {
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/?$select=id':
          return Promise.resolve({
            "id": "contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42"
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42/drives?$select=webUrl,id':
          return Promise.resolve({
            "value": [
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYVs-s_Fc6EaRomQ91r_60hi",
                "webUrl": "https://contoso.sharepoint.com/Big%20list"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                "webUrl": "https://contoso.sharepoint.com/DemoDocs"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYUMhRqN81GDSYJHirLtImTh",
                "webUrl": "https://contoso.sharepoint.com/Shared%20Documents"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYWwU-RsQtl_RJxAlcRhJSYH",
                "webUrl": "https://contoso.sharepoint.com/JTDesignDocs"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYX2VXqL8cJvTbmPdTxwV8t4",
                "webUrl": "https://contoso.sharepoint.com/MCASDemoFiles"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYWaXP0JCVAyRZbch81buwYY",
                "webUrl": "https://contoso.sharepoint.com/RMSDemoLib"
              }
            ]
          });
        case 'https://graph.microsoft.com/v1.0/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root?$select=id':
          return Promise.resolve({
            "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ"
          });
        case "https://graph.microsoft.com/v1.0/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/items('01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ')/children?$filter=folder ne null&$select=id":
          return Promise.resolve({
            "value": []
          });
        case 'https://graph.microsoft.com/v1.0/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/items/01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ/children':
          return Promise.resolve({
            "value": [
              {
                "createdDateTime": "2021-09-25T13:30:17Z",
                "eTag": "\"{A2331FD8-FAA8-4C41-A8A0-B6F27FB993AF},1\"",
                "id": "01YNDLPYOYD4Z2FKH2IFGKRIFW6J73TE5P",
                "lastModifiedDateTime": "2021-09-25T13:30:17Z",
                "name": "Fo'lde'r",
                "webUrl": "https://contoso.sharepoint.com/DemoDocs/Fo%27lde%27r",
                "cTag": "\"c:{A2331FD8-FAA8-4C41-A8A0-B6F27FB993AF},0\"",
                "size": 151415,
                "createdBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "802d8bad-2ae1-479a-bbbe-48aca058bc26",
                    "displayName": "MOD Administrator"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "802d8bad-2ae1-479a-bbbe-48aca058bc26",
                    "displayName": "MOD Administrator"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-09-25T13:30:17Z",
                  "lastModifiedDateTime": "2021-09-25T13:30:17Z"
                },
                "folder": {
                  "childCount": 2
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=b7b2648f-c50c-4d1e-bb99-9c1dcd5e37ae&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkNKTDB2UUp1SVF3enl1YUVHbDI1WVJpejJDdkM5bDdzRktVUFQzU0F2bTg9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.bmVGTE9nbDJ4c1lvem5sQ3pnTXRmb29HaTRUTWVVTlVaYXBkSVp3QWliRT0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:04Z",
                "eTag": "\"{B7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE},2\"",
                "id": "01YNDLPYMPMSZLODGFDZG3XGM4DXGV4N5O",
                "lastModifiedDateTime": "2021-06-19T11:01:04Z",
                "name": "Blog Post preview.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BB7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE%7D&file=Blog%20Post%20preview.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{B7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE},3\"",
                "size": 134144,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "4wybUm/SrJtzssP/YDDKBwKRYd8="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:04Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:04Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=4a3864b3-db49-4d6f-a72e-7e3269c44e33&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IjlUSENBN0loaWhMMUJhc2EyUjlneVBMYldaOGcxVS93VTYyRTU1NHU4MVE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.VEVUTXJYNk1mWmpBSFhudjJrVWMxTWQ0UzRCVzFad0F3K0QrV2VwcEVoaz0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:13Z",
                "eTag": "\"{4A3864B3-DB49-4D6F-A72E-7E3269C44E33},3\"",
                "id": "01YNDLPYNTMQ4EUSO3N5G2OLT6GJU4ITRT",
                "lastModifiedDateTime": "2021-06-19T11:01:13Z",
                "name": "Contoso Purchasing Permissions.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B4A3864B3-DB49-4D6F-A72E-7E3269C44E33%7D&file=Contoso%20Purchasing%20Permissions.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{4A3864B3-DB49-4D6F-A72E-7E3269C44E33},6\"",
                "dataLossPrevention": {
                  "block": {},
                  "notify": {}
                },
                "size": 27514,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "MTvtw1PC9WtEY0wycgD1aBIIX3A="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:13Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:13Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=1bcaf9f5-6788-42f1-9b77-ee43df61c4dd&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkozNjY5RnEwblpFWmluYTZHNjJSTTl0cGJCd1lxaWJuV2lING5SMVdDU1E9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.ck96RHVlYjdGMDhQWnhhc3UrZXhReHBaNHJnekcxUmVDUjMwdTNpSTJpND0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:23Z",
                "eTag": "\"{1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD},2\"",
                "id": "01YNDLPYPV7HFBXCDH6FBJW57OIPPWDRG5",
                "lastModifiedDateTime": "2021-06-19T11:01:23Z",
                "name": "Credit Cards.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD%7D&file=Credit%20Cards.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD},3\"",
                "size": 25245,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "xOKzXY19rWzvLGnGcjybqNcMtIc="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:23Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:23Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=bc13ae6b-5a75-427d-8915-fcb6e86edd11&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkpiTHU5OHpmTkxCMXhrZ1k1dTlkOUVONXU0UEZjWHpTQ0s0bTg3UU1LcTQ9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.dUFidWp1M1I4N2xwZmhyRDFQRGpZQ0ZQSHBsbVNOcVdFcDAwSjY3U1JGcz0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:32Z",
                "eTag": "\"{BC13AE6B-5A75-427D-8915-FCB6E86EDD11},2\"",
                "id": "01YNDLPYLLVYJ3Y5K2PVBISFP4W3UG5XIR",
                "lastModifiedDateTime": "2021-06-19T11:01:32Z",
                "name": "Customer Accounts.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BBC13AE6B-5A75-427D-8915-FCB6E86EDD11%7D&file=Customer%20Accounts.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{BC13AE6B-5A75-427D-8915-FCB6E86EDD11},3\"",
                "size": 28747,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "7nfr2jjjYOlbSFfYHX72+Ud1o6Y="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:32Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:32Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=5148569c-7459-40a8-a400-e56e4668ad62&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6InphTThqdWw5VW4wZlF3WWVTcnRmd1VWV2cyK2YvdXFVQWlzam9xTkdPTEU9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.K2RqWmd2WU5EWTVaWXprelRhUFBQcjZ0NkwxMzZJbnlPMG1QNDYyM3NvOD0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:00:46Z",
                "eTag": "\"{5148569C-7459-40A8-A400-E56E4668AD62},2\"",
                "id": "01YNDLPYM4KZEFCWLUVBAKIAHFNZDGRLLC",
                "lastModifiedDateTime": "2021-06-19T11:00:46Z",
                "name": "Customer Data.xlsx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B5148569C-7459-40A8-A400-E56E4668AD62%7D&file=Customer%20Data.xlsx&action=default&mobileredirect=true",
                "cTag": "\"c:{5148569C-7459-40A8-A400-E56E4668AD62},3\"",
                "size": 17661,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                  "hashes": {
                    "quickXorHash": "tTrnKAf/Pcq8/NXtvxHuAAUwIUs="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:00:46Z",
                  "lastModifiedDateTime": "2021-06-19T11:00:46Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=dd9ed836-18b0-44e7-b352-da56c300cdcc&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6Ik9XN2pqVTBJWkVxUEVLOU1yUXg0ZzgyQjd6ZnBzRWJ5TXNYY011NlpvbXc9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.eHNsek14b0tsQkZmNE9lT3UxWmlZMnRxZ0lub1Z5RzU4VnRkalRZWXcrQT0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:00:55Z",
                "eTag": "\"{DD9ED836-18B0-44E7-B352-DA56C300CDCC},2\"",
                "id": "01YNDLPYJW3CPN3MAY45CLGUW2K3BQBTOM",
                "lastModifiedDateTime": "2021-06-19T11:00:55Z",
                "name": "Q3 Sales and Marketing Expense Report Audit.pptx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BDD9ED836-18B0-44E7-B352-DA56C300CDCC%7D&file=Q3%20Sales%20and%20Marketing%20Expense%20Report%20Audit.pptx&action=edit&mobileredirect=true",
                "cTag": "\"c:{DD9ED836-18B0-44E7-B352-DA56C300CDCC},3\"",
                "size": 376325,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                  "hashes": {
                    "quickXorHash": "BCE09zeZNS8dnTx14w3stEm9g0g="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:00:55Z",
                  "lastModifiedDateTime": "2021-06-19T11:00:55Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=232b3df5-7f6c-4f06-b811-6029534bd1db&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6ImRNVEtaaGJuWVgyM3g4dURuQ0tXU0RQVFhjMWVSbGR3aWhrOXF6RjYvczA9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.M3ZVcE1NVnAxd0F4WDFiNUR5eHVoa3lZQ2x1eUUyVGwwdnFUZVNXaTJ1UT0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:44Z",
                "eTag": "\"{232B3DF5-7F6C-4F06-B811-6029534BD1DB},1\"",
                "id": "01YNDLPYPVHUVSG3D7AZH3QELAFFJUXUO3",
                "lastModifiedDateTime": "2021-06-19T11:01:44Z",
                "name": "Q3_Product_Strategy.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B232B3DF5-7F6C-4F06-B811-6029534BD1DB%7D&file=Q3_Product_Strategy.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{232B3DF5-7F6C-4F06-B811-6029534BD1DB},2\"",
                "size": 48473,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "54Y/Q4Ykxgg3N7yWL8iTlokMSS0="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:44Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:44Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=fcd62009-cf87-478e-a788-3853a9832f6b&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6ImVLQ0oyWDhpR3JvdGF0TGQ0d2hyOTNnajBKaUpUeEhUUHlmUWl2clJjOWc9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.bllWYkVsb3doSkxZbmFVTTE1TDdhUlU5NjNzWGNHUWRzbUNiOW1MN2s2az0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:52Z",
                "eTag": "\"{FCD62009-CF87-478E-A788-3853A9832F6B},1\"",
                "id": "01YNDLPYIJEDLPZB6PRZD2PCBYKOUYGL3L",
                "lastModifiedDateTime": "2021-06-19T11:01:52Z",
                "name": "Sales Memo.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BFCD62009-CF87-478E-A788-3853A9832F6B%7D&file=Sales%20Memo.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{FCD62009-CF87-478E-A788-3853A9832F6B},2\"",
                "size": 33737,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "dNTTVcFU+d1H7Ou3seZvQI1pGnk="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:52Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:52Z"
                }
              }
            ]
          });
        default:
          return Promise.reject(`Invalid GET request: ${url}`);
      }
    });

    command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/',
        folderUrl: '/DemoDocs',
        recursive: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        assert(loggerLogSpy.calledWith([
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=b7b2648f-c50c-4d1e-bb99-9c1dcd5e37ae&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkNKTDB2UUp1SVF3enl1YUVHbDI1WVJpejJDdkM5bDdzRktVUFQzU0F2bTg9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.bmVGTE9nbDJ4c1lvem5sQ3pnTXRmb29HaTRUTWVVTlVaYXBkSVp3QWliRT0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:04Z",
            "eTag": "\"{B7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE},2\"",
            "id": "01YNDLPYMPMSZLODGFDZG3XGM4DXGV4N5O",
            "lastModifiedDateTime": "2021-06-19T11:01:04Z",
            "name": "Blog Post preview.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BB7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE%7D&file=Blog%20Post%20preview.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{B7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE},3\"",
            "size": 134144,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "4wybUm/SrJtzssP/YDDKBwKRYd8="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:04Z",
              "lastModifiedDateTime": "2021-06-19T11:01:04Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=4a3864b3-db49-4d6f-a72e-7e3269c44e33&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IjlUSENBN0loaWhMMUJhc2EyUjlneVBMYldaOGcxVS93VTYyRTU1NHU4MVE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.VEVUTXJYNk1mWmpBSFhudjJrVWMxTWQ0UzRCVzFad0F3K0QrV2VwcEVoaz0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:13Z",
            "eTag": "\"{4A3864B3-DB49-4D6F-A72E-7E3269C44E33},3\"",
            "id": "01YNDLPYNTMQ4EUSO3N5G2OLT6GJU4ITRT",
            "lastModifiedDateTime": "2021-06-19T11:01:13Z",
            "name": "Contoso Purchasing Permissions.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B4A3864B3-DB49-4D6F-A72E-7E3269C44E33%7D&file=Contoso%20Purchasing%20Permissions.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{4A3864B3-DB49-4D6F-A72E-7E3269C44E33},6\"",
            "dataLossPrevention": {
              "block": {},
              "notify": {}
            },
            "size": 27514,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "MTvtw1PC9WtEY0wycgD1aBIIX3A="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:13Z",
              "lastModifiedDateTime": "2021-06-19T11:01:13Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=1bcaf9f5-6788-42f1-9b77-ee43df61c4dd&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkozNjY5RnEwblpFWmluYTZHNjJSTTl0cGJCd1lxaWJuV2lING5SMVdDU1E9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.ck96RHVlYjdGMDhQWnhhc3UrZXhReHBaNHJnekcxUmVDUjMwdTNpSTJpND0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:23Z",
            "eTag": "\"{1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD},2\"",
            "id": "01YNDLPYPV7HFBXCDH6FBJW57OIPPWDRG5",
            "lastModifiedDateTime": "2021-06-19T11:01:23Z",
            "name": "Credit Cards.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD%7D&file=Credit%20Cards.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{1BCAF9F5-6788-42F1-9B77-EE43DF61C4DD},3\"",
            "size": 25245,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "xOKzXY19rWzvLGnGcjybqNcMtIc="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:23Z",
              "lastModifiedDateTime": "2021-06-19T11:01:23Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=bc13ae6b-5a75-427d-8915-fcb6e86edd11&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkpiTHU5OHpmTkxCMXhrZ1k1dTlkOUVONXU0UEZjWHpTQ0s0bTg3UU1LcTQ9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.dUFidWp1M1I4N2xwZmhyRDFQRGpZQ0ZQSHBsbVNOcVdFcDAwSjY3U1JGcz0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:32Z",
            "eTag": "\"{BC13AE6B-5A75-427D-8915-FCB6E86EDD11},2\"",
            "id": "01YNDLPYLLVYJ3Y5K2PVBISFP4W3UG5XIR",
            "lastModifiedDateTime": "2021-06-19T11:01:32Z",
            "name": "Customer Accounts.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BBC13AE6B-5A75-427D-8915-FCB6E86EDD11%7D&file=Customer%20Accounts.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{BC13AE6B-5A75-427D-8915-FCB6E86EDD11},3\"",
            "size": 28747,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "7nfr2jjjYOlbSFfYHX72+Ud1o6Y="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:32Z",
              "lastModifiedDateTime": "2021-06-19T11:01:32Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=5148569c-7459-40a8-a400-e56e4668ad62&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6InphTThqdWw5VW4wZlF3WWVTcnRmd1VWV2cyK2YvdXFVQWlzam9xTkdPTEU9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.K2RqWmd2WU5EWTVaWXprelRhUFBQcjZ0NkwxMzZJbnlPMG1QNDYyM3NvOD0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:00:46Z",
            "eTag": "\"{5148569C-7459-40A8-A400-E56E4668AD62},2\"",
            "id": "01YNDLPYM4KZEFCWLUVBAKIAHFNZDGRLLC",
            "lastModifiedDateTime": "2021-06-19T11:00:46Z",
            "name": "Customer Data.xlsx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B5148569C-7459-40A8-A400-E56E4668AD62%7D&file=Customer%20Data.xlsx&action=default&mobileredirect=true",
            "cTag": "\"c:{5148569C-7459-40A8-A400-E56E4668AD62},3\"",
            "size": 17661,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              "hashes": {
                "quickXorHash": "tTrnKAf/Pcq8/NXtvxHuAAUwIUs="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:00:46Z",
              "lastModifiedDateTime": "2021-06-19T11:00:46Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=dd9ed836-18b0-44e7-b352-da56c300cdcc&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6Ik9XN2pqVTBJWkVxUEVLOU1yUXg0ZzgyQjd6ZnBzRWJ5TXNYY011NlpvbXc9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.eHNsek14b0tsQkZmNE9lT3UxWmlZMnRxZ0lub1Z5RzU4VnRkalRZWXcrQT0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:00:55Z",
            "eTag": "\"{DD9ED836-18B0-44E7-B352-DA56C300CDCC},2\"",
            "id": "01YNDLPYJW3CPN3MAY45CLGUW2K3BQBTOM",
            "lastModifiedDateTime": "2021-06-19T11:00:55Z",
            "name": "Q3 Sales and Marketing Expense Report Audit.pptx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BDD9ED836-18B0-44E7-B352-DA56C300CDCC%7D&file=Q3%20Sales%20and%20Marketing%20Expense%20Report%20Audit.pptx&action=edit&mobileredirect=true",
            "cTag": "\"c:{DD9ED836-18B0-44E7-B352-DA56C300CDCC},3\"",
            "size": 376325,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
              "hashes": {
                "quickXorHash": "BCE09zeZNS8dnTx14w3stEm9g0g="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:00:55Z",
              "lastModifiedDateTime": "2021-06-19T11:00:55Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=232b3df5-7f6c-4f06-b811-6029534bd1db&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6ImRNVEtaaGJuWVgyM3g4dURuQ0tXU0RQVFhjMWVSbGR3aWhrOXF6RjYvczA9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.M3ZVcE1NVnAxd0F4WDFiNUR5eHVoa3lZQ2x1eUUyVGwwdnFUZVNXaTJ1UT0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:44Z",
            "eTag": "\"{232B3DF5-7F6C-4F06-B811-6029534BD1DB},1\"",
            "id": "01YNDLPYPVHUVSG3D7AZH3QELAFFJUXUO3",
            "lastModifiedDateTime": "2021-06-19T11:01:44Z",
            "name": "Q3_Product_Strategy.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B232B3DF5-7F6C-4F06-B811-6029534BD1DB%7D&file=Q3_Product_Strategy.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{232B3DF5-7F6C-4F06-B811-6029534BD1DB},2\"",
            "size": 48473,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "54Y/Q4Ykxgg3N7yWL8iTlokMSS0="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:44Z",
              "lastModifiedDateTime": "2021-06-19T11:01:44Z"
            },
            "lastModifiedByUser": "Provisioning User"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=fcd62009-cf87-478e-a788-3853a9832f6b&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6ImVLQ0oyWDhpR3JvdGF0TGQ0d2hyOTNnajBKaUpUeEhUUHlmUWl2clJjOWc9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.bllWYkVsb3doSkxZbmFVTTE1TDdhUlU5NjNzWGNHUWRzbUNiOW1MN2s2az0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:52Z",
            "eTag": "\"{FCD62009-CF87-478E-A788-3853A9832F6B},1\"",
            "id": "01YNDLPYIJEDLPZB6PRZD2PCBYKOUYGL3L",
            "lastModifiedDateTime": "2021-06-19T11:01:52Z",
            "name": "Sales Memo.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BFCD62009-CF87-478E-A788-3853A9832F6B%7D&file=Sales%20Memo.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{FCD62009-CF87-478E-A788-3853A9832F6B},2\"",
            "size": 33737,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "lastModifiedBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "dNTTVcFU+d1H7Ou3seZvQI1pGnk="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:52Z",
              "lastModifiedDateTime": "2021-06-19T11:01:52Z"
            },
            "lastModifiedByUser": "Provisioning User"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('loads files recursively from a folder with one level of subfolders', done => {
    sinon.stub(request, 'get').callsFake(opts => {
      const url: string = opts.url as string;

      switch (url) {
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/design?$select=id':
          return Promise.resolve({
            "id": "contoso.sharepoint.com,c88ba08f-4e17-49c1-a34f-ca85d908c24c,fde71d4c-ac91-4464-a60f-632d99b8224d"
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,c88ba08f-4e17-49c1-a34f-ca85d908c24c,fde71d4c-ac91-4464-a60f-632d99b8224d/drives?$select=webUrl,id':
          return Promise.resolve({
            "value": [
              {
                "id": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                "webUrl": "https://contoso.sharepoint.com/sites/Design/DemoDocs"
              },
              {
                "id": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik0at4LaajDdTo6njHc5dx7g",
                "webUrl": "https://contoso.sharepoint.com/sites/Design/Shared%20Documents"
              }
            ]
          });
        case 'https://graph.microsoft.com/v1.0/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root?$select=id':
          return Promise.resolve({
            "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ"
          });
        case "https://graph.microsoft.com/v1.0/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/items('01472SYFF6Y2GOVW7725BZO354PWSELRRZ')/children?$filter=folder ne null&$select=id":
          return Promise.resolve({
            "value": [
              {
                "id": "01472SYFERKSBJFD7X35H3TUHUSHXZBRRD"
              }
            ]
          });
        case "https://graph.microsoft.com/v1.0/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/items('01472SYFERKSBJFD7X35H3TUHUSHXZBRRD')/children?$filter=folder ne null&$select=id":
          return Promise.resolve({
            "value": []
          });
        case 'https://graph.microsoft.com/v1.0/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/items/01472SYFF6Y2GOVW7725BZO354PWSELRRZ/children':
          return Promise.resolve({
            "value": [
              {
                "createdDateTime": "2021-11-13T16:40:31Z",
                "eTag": "\"{92825491-F78F-4FDF-B9D0-F491EF90C623},1\"",
                "id": "01472SYFERKSBJFD7X35H3TUHUSHXZBRRD",
                "lastModifiedDateTime": "2021-11-13T16:40:31Z",
                "name": "Folder",
                "webUrl": "https://contoso.sharepoint.com/sites/Design/DemoDocs/Folder",
                "cTag": "\"c:{92825491-F78F-4FDF-B9D0-F491EF90C623},0\"",
                "size": 131072,
                "createdBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "parentReference": {
                  "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                  "driveType": "documentLibrary",
                  "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-11-13T16:40:31Z",
                  "lastModifiedDateTime": "2021-11-13T16:40:31Z"
                },
                "folder": {
                  "childCount": 1
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=1bc151a6-beb8-4034-be93-6e9a18aa6cdc&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMjc4MSIsImV4cCI6IjE2MzY4MjYzODEiLCJlbmRwb2ludHVybCI6Ik5PNFErS3pZSFBwMTJpZUhiZjdobmUrQ1E0em0yZVVvbnZCczI4RHdZVGs9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik16Rm1Oall4WWpJdE5UQTRaaTAwTlRVNUxXRmtOamN0T0RZeE5ESXlaamd4TURKbCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjQwLjEyNi4zMi45OSJ9.RkRCUmdqa21sd1VKaDdUQU9WcmhnRnd3K01yZDZXaUp4OFJDcXJET2c1MD0&ApiVersion=2.0",
                "createdDateTime": "2021-11-13T16:27:04Z",
                "eTag": "\"{1BC151A6-BEB8-4034-BE93-6E9A18AA6CDC},1\"",
                "id": "01472SYFFGKHARXOF6GRAL5E3OTIMKU3G4",
                "lastModifiedDateTime": "2021-11-13T16:27:04Z",
                "name": "Blog Post preview.docx",
                "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B1BC151A6-BEB8-4034-BE93-6E9A18AA6CDC%7D&file=Blog%20Post%20preview.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{1BC151A6-BEB8-4034-BE93-6E9A18AA6CDC},4\"",
                "size": 131072,
                "createdBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "parentReference": {
                  "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                  "driveType": "documentLibrary",
                  "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "fVSMG9+U0AczPgoLQYCHu85LOT4="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-11-13T16:27:04Z",
                  "lastModifiedDateTime": "2021-11-13T16:27:04Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=716e4cb5-cefc-4562-97bb-b40a985f3306&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMjc4MSIsImV4cCI6IjE2MzY4MjYzODEiLCJlbmRwb2ludHVybCI6InlrUjVQQjNjU3NIV1hhVXVwZW9qbGtna05UeW9VcFdweGxiS0FZbURwV1k9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik16Rm1Oall4WWpJdE5UQTRaaTAwTlRVNUxXRmtOamN0T0RZeE5ESXlaamd4TURKbCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjQwLjEyNi4zMi45OSJ9.NFhRbzBqUHNwRjdnbVphb3I0anF0NnY4NzJRUVpsNTZ2Yk9QeW1Pa2lTQT0&ApiVersion=2.0",
                "createdDateTime": "2021-11-13T16:27:05Z",
                "eTag": "\"{716E4CB5-CEFC-4562-97BB-B40A985F3306},2\"",
                "id": "01472SYFFVJRXHD7GOMJCZPO5UBKMF6MYG",
                "lastModifiedDateTime": "2021-11-13T16:27:05Z",
                "name": "Contoso Purchasing Permissions.docx",
                "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B716E4CB5-CEFC-4562-97BB-B40A985F3306%7D&file=Contoso%20Purchasing%20Permissions.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{716E4CB5-CEFC-4562-97BB-B40A985F3306},5\"",
                "dataLossPrevention": {
                  "block": {},
                  "notify": {}
                },
                "size": 27514,
                "createdBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "parentReference": {
                  "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                  "driveType": "documentLibrary",
                  "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "7PDBFlBace6tv3M/yVpdtz8fm/c="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-11-13T16:27:05Z",
                  "lastModifiedDateTime": "2021-11-13T16:27:05Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=f123cf01-244e-44d3-b2f7-24f5230937a1&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMjc4MSIsImV4cCI6IjE2MzY4MjYzODEiLCJlbmRwb2ludHVybCI6IkcxMS8vN2M1Y1RyU1JwVEFZOUJSSG1WVUlYYTRNTXdnKzJRcTVFZE1OZ1k9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik16Rm1Oall4WWpJdE5UQTRaaTAwTlRVNUxXRmtOamN0T0RZeE5ESXlaamd4TURKbCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjQwLjEyNi4zMi45OSJ9.Y0d5Y1ZuZ05zZDlRRk9TYWliNWV4bXhuL2JhOWxVdnhCZk5iK0txVU9yRT0&ApiVersion=2.0",
                "createdDateTime": "2021-11-13T16:27:04Z",
                "eTag": "\"{F123CF01-244E-44D3-B2F7-24F5230937A1},1\"",
                "id": "01472SYFABZ4R7CTRE2NCLF5ZE6URQSN5B",
                "lastModifiedDateTime": "2021-11-13T16:27:04Z",
                "name": "Credit Cards.docx",
                "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7BF123CF01-244E-44D3-B2F7-24F5230937A1%7D&file=Credit%20Cards.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{F123CF01-244E-44D3-B2F7-24F5230937A1},4\"",
                "size": 24858,
                "createdBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "parentReference": {
                  "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                  "driveType": "documentLibrary",
                  "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "9LEEpPZjBLjJQmN/AKDXRLmXj/k="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-11-13T16:27:04Z",
                  "lastModifiedDateTime": "2021-11-13T16:27:04Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=9b9a62cf-e7c6-4fa4-a261-1b260d7d6cf1&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMjc4MSIsImV4cCI6IjE2MzY4MjYzODEiLCJlbmRwb2ludHVybCI6ImordllIUXB6d1hhM01HV3lJb3JuZWRzZlpXbGZIQXlsSmlsQTE2NU1HMms9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik16Rm1Oall4WWpJdE5UQTRaaTAwTlRVNUxXRmtOamN0T0RZeE5ESXlaamd4TURKbCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjQwLjEyNi4zMi45OSJ9.SmdsZHBMQ1poS3JVdnBSbVN0MzE4R2c4RWJTY29TZktLemtJYlRKVElGaz0&ApiVersion=2.0",
                "createdDateTime": "2021-11-13T16:27:05Z",
                "eTag": "\"{9B9A62CF-E7C6-4FA4-A261-1B260D7D6CF1},1\"",
                "id": "01472SYFGPMKNJXRXHURH2EYI3EYGX23HR",
                "lastModifiedDateTime": "2021-11-13T16:27:05Z",
                "name": "Customer Accounts.docx",
                "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B9B9A62CF-E7C6-4FA4-A261-1B260D7D6CF1%7D&file=Customer%20Accounts.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{9B9A62CF-E7C6-4FA4-A261-1B260D7D6CF1},4\"",
                "size": 28360,
                "createdBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "parentReference": {
                  "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                  "driveType": "documentLibrary",
                  "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "PTFegvh8VBE1mJ3eELTSb+LSQxI="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-11-13T16:27:05Z",
                  "lastModifiedDateTime": "2021-11-13T16:27:05Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=e2f09ed9-987a-4d26-8513-86878f785e0a&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMjc4MSIsImV4cCI6IjE2MzY4MjYzODEiLCJlbmRwb2ludHVybCI6IkRORDROcGlaQXAxa3RCR21HckRFa0ZScmdEQWVyeVE0TjgyK1dBNWJDaVE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik16Rm1Oall4WWpJdE5UQTRaaTAwTlRVNUxXRmtOamN0T0RZeE5ESXlaamd4TURKbCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjQwLjEyNi4zMi45OSJ9.NGM5MkoyWHk5M2xCZTBiN0twVXdXU0c5ZEVlbW1SRUxzZ1hwanp2Wm5FZz0&ApiVersion=2.0",
                "createdDateTime": "2021-11-13T16:27:04Z",
                "eTag": "\"{E2F09ED9-987A-4D26-8513-86878F785E0A},1\"",
                "id": "01472SYFGZT3YOE6UYEZGYKE4GQ6HXQXQK",
                "lastModifiedDateTime": "2021-11-13T16:27:04Z",
                "name": "Customer Data.xlsx",
                "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7BE2F09ED9-987A-4D26-8513-86878F785E0A%7D&file=Customer%20Data.xlsx&action=default&mobileredirect=true",
                "cTag": "\"c:{E2F09ED9-987A-4D26-8513-86878F785E0A},4\"",
                "size": 17448,
                "createdBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "parentReference": {
                  "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                  "driveType": "documentLibrary",
                  "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                  "hashes": {
                    "quickXorHash": "hrMOe12SPOwumxGnhdJiYnw1qoM="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-11-13T16:27:04Z",
                  "lastModifiedDateTime": "2021-11-13T16:27:04Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=370c45c8-cca6-4333-9cb4-d33b7bc3fc05&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMjc4MSIsImV4cCI6IjE2MzY4MjYzODEiLCJlbmRwb2ludHVybCI6InAzSHU2TjRHSjRSTW9sREVJOVRLdE9DVzk0SXUzMllyTWdVak9IbCtlWkU9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik16Rm1Oall4WWpJdE5UQTRaaTAwTlRVNUxXRmtOamN0T0RZeE5ESXlaamd4TURKbCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjQwLjEyNi4zMi45OSJ9.TnVVZkxGWUJ6dVRrNWJuZDJsU3RzZnVLNUR3azRLRGhZRm42RW5tY1ZSTT0&ApiVersion=2.0",
                "createdDateTime": "2021-11-13T16:27:05Z",
                "eTag": "\"{370C45C8-CCA6-4333-9CB4-D33B7BC3FC05},1\"",
                "id": "01472SYFGIIUGDPJWMGNBZZNGTHN54H7AF",
                "lastModifiedDateTime": "2021-11-13T16:27:05Z",
                "name": "Q3 Sales and Marketing Expense Report Audit.pptx",
                "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B370C45C8-CCA6-4333-9CB4-D33B7BC3FC05%7D&file=Q3%20Sales%20and%20Marketing%20Expense%20Report%20Audit.pptx&action=edit&mobileredirect=true",
                "cTag": "\"c:{370C45C8-CCA6-4333-9CB4-D33B7BC3FC05},4\"",
                "size": 375938,
                "createdBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "parentReference": {
                  "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                  "driveType": "documentLibrary",
                  "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                  "hashes": {
                    "quickXorHash": "/qGrwC7X0bUey5sxCDlzxzSjTLE="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-11-13T16:27:05Z",
                  "lastModifiedDateTime": "2021-11-13T16:27:05Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=2223f837-ce90-4638-a2ca-8fd753c4378d&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMjc4MSIsImV4cCI6IjE2MzY4MjYzODEiLCJlbmRwb2ludHVybCI6InhuVmpzU1NWUWxTWVhWK3dDNElpVmtpMUIrZ0xTYWFpTldTaEpQemFnaFk9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik16Rm1Oall4WWpJdE5UQTRaaTAwTlRVNUxXRmtOamN0T0RZeE5ESXlaamd4TURKbCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjQwLjEyNi4zMi45OSJ9.cU1FaENWYW5PWVh0dHErZ2FLWEFPWFBQam12c1dublVEbUc0U0dSWnNTMD0&ApiVersion=2.0",
                "createdDateTime": "2021-11-13T16:27:04Z",
                "eTag": "\"{2223F837-CE90-4638-A2CA-8FD753C4378D},1\"",
                "id": "01472SYFBX7ARSFEGOHBDKFSUP25J4IN4N",
                "lastModifiedDateTime": "2021-11-13T16:27:04Z",
                "name": "Q3_Product_Strategy.docx",
                "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B2223F837-CE90-4638-A2CA-8FD753C4378D%7D&file=Q3_Product_Strategy.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{2223F837-CE90-4638-A2CA-8FD753C4378D},4\"",
                "size": 48086,
                "createdBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "parentReference": {
                  "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                  "driveType": "documentLibrary",
                  "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "LWe4j5GTM+Sg3PGOV0xoSSCnTr8="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-11-13T16:27:04Z",
                  "lastModifiedDateTime": "2021-11-13T16:27:04Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=bfc3a23e-ed0b-4a59-96f1-b11c075d89b0&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMjc4MSIsImV4cCI6IjE2MzY4MjYzODEiLCJlbmRwb2ludHVybCI6IjVBMkhoblFOcEozMFlsZUp3NDU2ZENCTUlJNXlXbnRoaFhGVGZZQUZuMXM9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik16Rm1Oall4WWpJdE5UQTRaaTAwTlRVNUxXRmtOamN0T0RZeE5ESXlaamd4TURKbCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjQwLjEyNi4zMi45OSJ9.N3JrZldpUnVQK09pTnl6RGJSY0pUdWZVbDhSSkl3R0l4ZkJESGJ6a3h3dz0&ApiVersion=2.0",
                "createdDateTime": "2021-11-13T16:27:05Z",
                "eTag": "\"{BFC3A23E-ED0B-4A59-96F1-B11C075D89B0},1\"",
                "id": "01472SYFB6ULB36C7NLFFJN4NRDQDV3CNQ",
                "lastModifiedDateTime": "2021-11-13T16:27:05Z",
                "name": "Sales Memo.docx",
                "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7BBFC3A23E-ED0B-4A59-96F1-B11C075D89B0%7D&file=Sales%20Memo.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{BFC3A23E-ED0B-4A59-96F1-B11C075D89B0},4\"",
                "size": 35655,
                "createdBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "parentReference": {
                  "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                  "driveType": "documentLibrary",
                  "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "4qr1R9xcSfoofaFA3nMY6qjQXPQ="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-11-13T16:27:05Z",
                  "lastModifiedDateTime": "2021-11-13T16:27:05Z"
                }
              }
            ]
          });
        case 'https://graph.microsoft.com/v1.0/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/items/01472SYFERKSBJFD7X35H3TUHUSHXZBRRD/children':
          return Promise.resolve({
            "value": [
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=77e5e9f4-4731-478e-82ae-6eece079ae8b&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMjc4MSIsImV4cCI6IjE2MzY4MjYzODEiLCJlbmRwb2ludHVybCI6IlFWaVJuSEVIRWRHNW5vSHhrOUFwQ3lXS3JKa05uL3pFVFhqT3NGWGpmQkE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik5tSmtPRFppTm1JdE5XWXhaaTAwTW1VNExXRTVOMkV0TXpsaVpqTXhNR0U1TURaaCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjQwLjEyNi4zMi45OSJ9.c0FyM3hCR2pYQzdxQ0xEZ1VEamZJaDhFdHIzWFkxN0tWNGlNV2djcktRRT0&ApiVersion=2.0",
                "createdDateTime": "2021-11-13T16:40:39Z",
                "eTag": "\"{77E5E9F4-4731-478E-82AE-6EECE079AE8B},2\"",
                "id": "01472SYFHU5HSXOMKHRZDYFLTO5TQHTLUL",
                "lastModifiedDateTime": "2021-11-13T16:40:39Z",
                "name": "Blog Post preview.docx",
                "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B77E5E9F4-4731-478E-82AE-6EECE079AE8B%7D&file=Blog%20Post%20preview.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{77E5E9F4-4731-478E-82AE-6EECE079AE8B},4\"",
                "size": 131072,
                "createdBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "parentReference": {
                  "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                  "driveType": "documentLibrary",
                  "id": "01472SYFERKSBJFD7X35H3TUHUSHXZBRRD",
                  "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:/Folder"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-11-13T16:40:39Z",
                  "lastModifiedDateTime": "2021-11-13T16:40:39Z"
                }
              }
            ]
          });
        default:
          return Promise.reject(`Invalid GET request: ${url}`);
      }
    });

    command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/design',
        folderUrl: 'DemoDocs',
        recursive: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        assert(loggerLogSpy.calledWith([
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=1bc151a6-beb8-4034-be93-6e9a18aa6cdc&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMjc4MSIsImV4cCI6IjE2MzY4MjYzODEiLCJlbmRwb2ludHVybCI6Ik5PNFErS3pZSFBwMTJpZUhiZjdobmUrQ1E0em0yZVVvbnZCczI4RHdZVGs9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik16Rm1Oall4WWpJdE5UQTRaaTAwTlRVNUxXRmtOamN0T0RZeE5ESXlaamd4TURKbCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjQwLjEyNi4zMi45OSJ9.RkRCUmdqa21sd1VKaDdUQU9WcmhnRnd3K01yZDZXaUp4OFJDcXJET2c1MD0&ApiVersion=2.0",
            "createdDateTime": "2021-11-13T16:27:04Z",
            "eTag": "\"{1BC151A6-BEB8-4034-BE93-6E9A18AA6CDC},1\"",
            "id": "01472SYFFGKHARXOF6GRAL5E3OTIMKU3G4",
            "lastModifiedDateTime": "2021-11-13T16:27:04Z",
            "name": "Blog Post preview.docx",
            "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B1BC151A6-BEB8-4034-BE93-6E9A18AA6CDC%7D&file=Blog%20Post%20preview.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{1BC151A6-BEB8-4034-BE93-6E9A18AA6CDC},4\"",
            "size": 131072,
            "createdBy": {
              "user": {
                "email": "admin@contoso.OnMicrosoft.com",
                "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                "displayName": "MOD Administrator"
              }
            },
            "lastModifiedBy": {
              "user": {
                "email": "admin@contoso.OnMicrosoft.com",
                "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                "displayName": "MOD Administrator"
              }
            },
            "parentReference": {
              "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
              "driveType": "documentLibrary",
              "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "fVSMG9+U0AczPgoLQYCHu85LOT4="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-11-13T16:27:04Z",
              "lastModifiedDateTime": "2021-11-13T16:27:04Z"
            },
            "lastModifiedByUser": "MOD Administrator"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=716e4cb5-cefc-4562-97bb-b40a985f3306&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMjc4MSIsImV4cCI6IjE2MzY4MjYzODEiLCJlbmRwb2ludHVybCI6InlrUjVQQjNjU3NIV1hhVXVwZW9qbGtna05UeW9VcFdweGxiS0FZbURwV1k9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik16Rm1Oall4WWpJdE5UQTRaaTAwTlRVNUxXRmtOamN0T0RZeE5ESXlaamd4TURKbCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjQwLjEyNi4zMi45OSJ9.NFhRbzBqUHNwRjdnbVphb3I0anF0NnY4NzJRUVpsNTZ2Yk9QeW1Pa2lTQT0&ApiVersion=2.0",
            "createdDateTime": "2021-11-13T16:27:05Z",
            "eTag": "\"{716E4CB5-CEFC-4562-97BB-B40A985F3306},2\"",
            "id": "01472SYFFVJRXHD7GOMJCZPO5UBKMF6MYG",
            "lastModifiedDateTime": "2021-11-13T16:27:05Z",
            "name": "Contoso Purchasing Permissions.docx",
            "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B716E4CB5-CEFC-4562-97BB-B40A985F3306%7D&file=Contoso%20Purchasing%20Permissions.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{716E4CB5-CEFC-4562-97BB-B40A985F3306},5\"",
            "dataLossPrevention": {
              "block": {},
              "notify": {}
            },
            "size": 27514,
            "createdBy": {
              "user": {
                "email": "admin@contoso.OnMicrosoft.com",
                "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                "displayName": "MOD Administrator"
              }
            },
            "lastModifiedBy": {
              "user": {
                "email": "admin@contoso.OnMicrosoft.com",
                "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                "displayName": "MOD Administrator"
              }
            },
            "parentReference": {
              "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
              "driveType": "documentLibrary",
              "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "7PDBFlBace6tv3M/yVpdtz8fm/c="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-11-13T16:27:05Z",
              "lastModifiedDateTime": "2021-11-13T16:27:05Z"
            },
            "lastModifiedByUser": "MOD Administrator"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=f123cf01-244e-44d3-b2f7-24f5230937a1&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMjc4MSIsImV4cCI6IjE2MzY4MjYzODEiLCJlbmRwb2ludHVybCI6IkcxMS8vN2M1Y1RyU1JwVEFZOUJSSG1WVUlYYTRNTXdnKzJRcTVFZE1OZ1k9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik16Rm1Oall4WWpJdE5UQTRaaTAwTlRVNUxXRmtOamN0T0RZeE5ESXlaamd4TURKbCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjQwLjEyNi4zMi45OSJ9.Y0d5Y1ZuZ05zZDlRRk9TYWliNWV4bXhuL2JhOWxVdnhCZk5iK0txVU9yRT0&ApiVersion=2.0",
            "createdDateTime": "2021-11-13T16:27:04Z",
            "eTag": "\"{F123CF01-244E-44D3-B2F7-24F5230937A1},1\"",
            "id": "01472SYFABZ4R7CTRE2NCLF5ZE6URQSN5B",
            "lastModifiedDateTime": "2021-11-13T16:27:04Z",
            "name": "Credit Cards.docx",
            "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7BF123CF01-244E-44D3-B2F7-24F5230937A1%7D&file=Credit%20Cards.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{F123CF01-244E-44D3-B2F7-24F5230937A1},4\"",
            "size": 24858,
            "createdBy": {
              "user": {
                "email": "admin@contoso.OnMicrosoft.com",
                "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                "displayName": "MOD Administrator"
              }
            },
            "lastModifiedBy": {
              "user": {
                "email": "admin@contoso.OnMicrosoft.com",
                "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                "displayName": "MOD Administrator"
              }
            },
            "parentReference": {
              "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
              "driveType": "documentLibrary",
              "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "9LEEpPZjBLjJQmN/AKDXRLmXj/k="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-11-13T16:27:04Z",
              "lastModifiedDateTime": "2021-11-13T16:27:04Z"
            },
            "lastModifiedByUser": "MOD Administrator"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=9b9a62cf-e7c6-4fa4-a261-1b260d7d6cf1&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMjc4MSIsImV4cCI6IjE2MzY4MjYzODEiLCJlbmRwb2ludHVybCI6ImordllIUXB6d1hhM01HV3lJb3JuZWRzZlpXbGZIQXlsSmlsQTE2NU1HMms9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik16Rm1Oall4WWpJdE5UQTRaaTAwTlRVNUxXRmtOamN0T0RZeE5ESXlaamd4TURKbCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjQwLjEyNi4zMi45OSJ9.SmdsZHBMQ1poS3JVdnBSbVN0MzE4R2c4RWJTY29TZktLemtJYlRKVElGaz0&ApiVersion=2.0",
            "createdDateTime": "2021-11-13T16:27:05Z",
            "eTag": "\"{9B9A62CF-E7C6-4FA4-A261-1B260D7D6CF1},1\"",
            "id": "01472SYFGPMKNJXRXHURH2EYI3EYGX23HR",
            "lastModifiedDateTime": "2021-11-13T16:27:05Z",
            "name": "Customer Accounts.docx",
            "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B9B9A62CF-E7C6-4FA4-A261-1B260D7D6CF1%7D&file=Customer%20Accounts.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{9B9A62CF-E7C6-4FA4-A261-1B260D7D6CF1},4\"",
            "size": 28360,
            "createdBy": {
              "user": {
                "email": "admin@contoso.OnMicrosoft.com",
                "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                "displayName": "MOD Administrator"
              }
            },
            "lastModifiedBy": {
              "user": {
                "email": "admin@contoso.OnMicrosoft.com",
                "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                "displayName": "MOD Administrator"
              }
            },
            "parentReference": {
              "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
              "driveType": "documentLibrary",
              "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "PTFegvh8VBE1mJ3eELTSb+LSQxI="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-11-13T16:27:05Z",
              "lastModifiedDateTime": "2021-11-13T16:27:05Z"
            },
            "lastModifiedByUser": "MOD Administrator"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=e2f09ed9-987a-4d26-8513-86878f785e0a&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMjc4MSIsImV4cCI6IjE2MzY4MjYzODEiLCJlbmRwb2ludHVybCI6IkRORDROcGlaQXAxa3RCR21HckRFa0ZScmdEQWVyeVE0TjgyK1dBNWJDaVE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik16Rm1Oall4WWpJdE5UQTRaaTAwTlRVNUxXRmtOamN0T0RZeE5ESXlaamd4TURKbCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjQwLjEyNi4zMi45OSJ9.NGM5MkoyWHk5M2xCZTBiN0twVXdXU0c5ZEVlbW1SRUxzZ1hwanp2Wm5FZz0&ApiVersion=2.0",
            "createdDateTime": "2021-11-13T16:27:04Z",
            "eTag": "\"{E2F09ED9-987A-4D26-8513-86878F785E0A},1\"",
            "id": "01472SYFGZT3YOE6UYEZGYKE4GQ6HXQXQK",
            "lastModifiedDateTime": "2021-11-13T16:27:04Z",
            "name": "Customer Data.xlsx",
            "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7BE2F09ED9-987A-4D26-8513-86878F785E0A%7D&file=Customer%20Data.xlsx&action=default&mobileredirect=true",
            "cTag": "\"c:{E2F09ED9-987A-4D26-8513-86878F785E0A},4\"",
            "size": 17448,
            "createdBy": {
              "user": {
                "email": "admin@contoso.OnMicrosoft.com",
                "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                "displayName": "MOD Administrator"
              }
            },
            "lastModifiedBy": {
              "user": {
                "email": "admin@contoso.OnMicrosoft.com",
                "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                "displayName": "MOD Administrator"
              }
            },
            "parentReference": {
              "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
              "driveType": "documentLibrary",
              "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              "hashes": {
                "quickXorHash": "hrMOe12SPOwumxGnhdJiYnw1qoM="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-11-13T16:27:04Z",
              "lastModifiedDateTime": "2021-11-13T16:27:04Z"
            },
            "lastModifiedByUser": "MOD Administrator"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=370c45c8-cca6-4333-9cb4-d33b7bc3fc05&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMjc4MSIsImV4cCI6IjE2MzY4MjYzODEiLCJlbmRwb2ludHVybCI6InAzSHU2TjRHSjRSTW9sREVJOVRLdE9DVzk0SXUzMllyTWdVak9IbCtlWkU9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik16Rm1Oall4WWpJdE5UQTRaaTAwTlRVNUxXRmtOamN0T0RZeE5ESXlaamd4TURKbCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjQwLjEyNi4zMi45OSJ9.TnVVZkxGWUJ6dVRrNWJuZDJsU3RzZnVLNUR3azRLRGhZRm42RW5tY1ZSTT0&ApiVersion=2.0",
            "createdDateTime": "2021-11-13T16:27:05Z",
            "eTag": "\"{370C45C8-CCA6-4333-9CB4-D33B7BC3FC05},1\"",
            "id": "01472SYFGIIUGDPJWMGNBZZNGTHN54H7AF",
            "lastModifiedDateTime": "2021-11-13T16:27:05Z",
            "name": "Q3 Sales and Marketing Expense Report Audit.pptx",
            "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B370C45C8-CCA6-4333-9CB4-D33B7BC3FC05%7D&file=Q3%20Sales%20and%20Marketing%20Expense%20Report%20Audit.pptx&action=edit&mobileredirect=true",
            "cTag": "\"c:{370C45C8-CCA6-4333-9CB4-D33B7BC3FC05},4\"",
            "size": 375938,
            "createdBy": {
              "user": {
                "email": "admin@contoso.OnMicrosoft.com",
                "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                "displayName": "MOD Administrator"
              }
            },
            "lastModifiedBy": {
              "user": {
                "email": "admin@contoso.OnMicrosoft.com",
                "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                "displayName": "MOD Administrator"
              }
            },
            "parentReference": {
              "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
              "driveType": "documentLibrary",
              "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
              "hashes": {
                "quickXorHash": "/qGrwC7X0bUey5sxCDlzxzSjTLE="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-11-13T16:27:05Z",
              "lastModifiedDateTime": "2021-11-13T16:27:05Z"
            },
            "lastModifiedByUser": "MOD Administrator"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=2223f837-ce90-4638-a2ca-8fd753c4378d&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMjc4MSIsImV4cCI6IjE2MzY4MjYzODEiLCJlbmRwb2ludHVybCI6InhuVmpzU1NWUWxTWVhWK3dDNElpVmtpMUIrZ0xTYWFpTldTaEpQemFnaFk9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik16Rm1Oall4WWpJdE5UQTRaaTAwTlRVNUxXRmtOamN0T0RZeE5ESXlaamd4TURKbCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjQwLjEyNi4zMi45OSJ9.cU1FaENWYW5PWVh0dHErZ2FLWEFPWFBQam12c1dublVEbUc0U0dSWnNTMD0&ApiVersion=2.0",
            "createdDateTime": "2021-11-13T16:27:04Z",
            "eTag": "\"{2223F837-CE90-4638-A2CA-8FD753C4378D},1\"",
            "id": "01472SYFBX7ARSFEGOHBDKFSUP25J4IN4N",
            "lastModifiedDateTime": "2021-11-13T16:27:04Z",
            "name": "Q3_Product_Strategy.docx",
            "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B2223F837-CE90-4638-A2CA-8FD753C4378D%7D&file=Q3_Product_Strategy.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{2223F837-CE90-4638-A2CA-8FD753C4378D},4\"",
            "size": 48086,
            "createdBy": {
              "user": {
                "email": "admin@contoso.OnMicrosoft.com",
                "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                "displayName": "MOD Administrator"
              }
            },
            "lastModifiedBy": {
              "user": {
                "email": "admin@contoso.OnMicrosoft.com",
                "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                "displayName": "MOD Administrator"
              }
            },
            "parentReference": {
              "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
              "driveType": "documentLibrary",
              "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "LWe4j5GTM+Sg3PGOV0xoSSCnTr8="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-11-13T16:27:04Z",
              "lastModifiedDateTime": "2021-11-13T16:27:04Z"
            },
            "lastModifiedByUser": "MOD Administrator"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=bfc3a23e-ed0b-4a59-96f1-b11c075d89b0&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMjc4MSIsImV4cCI6IjE2MzY4MjYzODEiLCJlbmRwb2ludHVybCI6IjVBMkhoblFOcEozMFlsZUp3NDU2ZENCTUlJNXlXbnRoaFhGVGZZQUZuMXM9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik16Rm1Oall4WWpJdE5UQTRaaTAwTlRVNUxXRmtOamN0T0RZeE5ESXlaamd4TURKbCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjQwLjEyNi4zMi45OSJ9.N3JrZldpUnVQK09pTnl6RGJSY0pUdWZVbDhSSkl3R0l4ZkJESGJ6a3h3dz0&ApiVersion=2.0",
            "createdDateTime": "2021-11-13T16:27:05Z",
            "eTag": "\"{BFC3A23E-ED0B-4A59-96F1-B11C075D89B0},1\"",
            "id": "01472SYFB6ULB36C7NLFFJN4NRDQDV3CNQ",
            "lastModifiedDateTime": "2021-11-13T16:27:05Z",
            "name": "Sales Memo.docx",
            "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7BBFC3A23E-ED0B-4A59-96F1-B11C075D89B0%7D&file=Sales%20Memo.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{BFC3A23E-ED0B-4A59-96F1-B11C075D89B0},4\"",
            "size": 35655,
            "createdBy": {
              "user": {
                "email": "admin@contoso.OnMicrosoft.com",
                "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                "displayName": "MOD Administrator"
              }
            },
            "lastModifiedBy": {
              "user": {
                "email": "admin@contoso.OnMicrosoft.com",
                "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                "displayName": "MOD Administrator"
              }
            },
            "parentReference": {
              "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
              "driveType": "documentLibrary",
              "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "4qr1R9xcSfoofaFA3nMY6qjQXPQ="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-11-13T16:27:05Z",
              "lastModifiedDateTime": "2021-11-13T16:27:05Z"
            },
            "lastModifiedByUser": "MOD Administrator"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=77e5e9f4-4731-478e-82ae-6eece079ae8b&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMjc4MSIsImV4cCI6IjE2MzY4MjYzODEiLCJlbmRwb2ludHVybCI6IlFWaVJuSEVIRWRHNW5vSHhrOUFwQ3lXS3JKa05uL3pFVFhqT3NGWGpmQkE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik5tSmtPRFppTm1JdE5XWXhaaTAwTW1VNExXRTVOMkV0TXpsaVpqTXhNR0U1TURaaCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjQwLjEyNi4zMi45OSJ9.c0FyM3hCR2pYQzdxQ0xEZ1VEamZJaDhFdHIzWFkxN0tWNGlNV2djcktRRT0&ApiVersion=2.0",
            "createdDateTime": "2021-11-13T16:40:39Z",
            "eTag": "\"{77E5E9F4-4731-478E-82AE-6EECE079AE8B},2\"",
            "id": "01472SYFHU5HSXOMKHRZDYFLTO5TQHTLUL",
            "lastModifiedDateTime": "2021-11-13T16:40:39Z",
            "name": "Blog Post preview.docx",
            "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B77E5E9F4-4731-478E-82AE-6EECE079AE8B%7D&file=Blog%20Post%20preview.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{77E5E9F4-4731-478E-82AE-6EECE079AE8B},4\"",
            "size": 131072,
            "createdBy": {
              "user": {
                "email": "admin@contoso.OnMicrosoft.com",
                "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                "displayName": "MOD Administrator"
              }
            },
            "lastModifiedBy": {
              "user": {
                "email": "admin@contoso.OnMicrosoft.com",
                "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                "displayName": "MOD Administrator"
              }
            },
            "parentReference": {
              "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
              "driveType": "documentLibrary",
              "id": "01472SYFERKSBJFD7X35H3TUHUSHXZBRRD",
              "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:/Folder"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-11-13T16:40:39Z",
              "lastModifiedDateTime": "2021-11-13T16:40:39Z"
            },
            "lastModifiedByUser": "MOD Administrator"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('loads files recursively from a folder with multiple levels of subfolders', done => {
    sinon.stub(request, 'get').callsFake(opts => {
      const url: string = opts.url as string;

      switch (url) {
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/design?$select=id':
          return Promise.resolve({
            "id": "contoso.sharepoint.com,c88ba08f-4e17-49c1-a34f-ca85d908c24c,fde71d4c-ac91-4464-a60f-632d99b8224d"
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,c88ba08f-4e17-49c1-a34f-ca85d908c24c,fde71d4c-ac91-4464-a60f-632d99b8224d/drives?$select=webUrl,id':
          return Promise.resolve({
            "value": [
              {
                "id": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                "webUrl": "https://contoso.sharepoint.com/sites/Design/DemoDocs"
              },
              {
                "id": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik0at4LaajDdTo6njHc5dx7g",
                "webUrl": "https://contoso.sharepoint.com/sites/Design/Shared%20Documents"
              }
            ]
          });
        case 'https://graph.microsoft.com/v1.0/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root?$select=id':
          return Promise.resolve({
            "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ"
          });
        case "https://graph.microsoft.com/v1.0/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/items('01472SYFF6Y2GOVW7725BZO354PWSELRRZ')/children?$filter=folder ne null&$select=id":
          return Promise.resolve({
            "value": [
              {
                "id": "01472SYFERKSBJFD7X35H3TUHUSHXZBRRD"
              }
            ]
          });
        case "https://graph.microsoft.com/v1.0/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/items('01472SYFERKSBJFD7X35H3TUHUSHXZBRRD')/children?$filter=folder ne null&$select=id":
          return Promise.resolve({
            "value": [
              {
                "id": "01472SYFEVAFHSTTHMVRBZIGMOAWHT2VIG"
              }
            ]
          });
        case "https://graph.microsoft.com/v1.0/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/items('01472SYFEVAFHSTTHMVRBZIGMOAWHT2VIG')/children?$filter=folder ne null&$select=id":
          return Promise.resolve({
            "value": []
          });
        case 'https://graph.microsoft.com/v1.0/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/items/01472SYFF6Y2GOVW7725BZO354PWSELRRZ/children':
          return Promise.resolve({
            "value": [
              {
                "createdDateTime": "2021-11-13T16:40:31Z",
                "eTag": "\"{92825491-F78F-4FDF-B9D0-F491EF90C623},1\"",
                "id": "01472SYFERKSBJFD7X35H3TUHUSHXZBRRD",
                "lastModifiedDateTime": "2021-11-13T16:40:31Z",
                "name": "Folder",
                "webUrl": "https://contoso.sharepoint.com/sites/Design/DemoDocs/Folder",
                "cTag": "\"c:{92825491-F78F-4FDF-B9D0-F491EF90C623},0\"",
                "size": 131072,
                "createdBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "parentReference": {
                  "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                  "driveType": "documentLibrary",
                  "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-11-13T16:40:31Z",
                  "lastModifiedDateTime": "2021-11-13T16:40:31Z"
                },
                "folder": {
                  "childCount": 1
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=1bc151a6-beb8-4034-be93-6e9a18aa6cdc&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMjc4MSIsImV4cCI6IjE2MzY4MjYzODEiLCJlbmRwb2ludHVybCI6Ik5PNFErS3pZSFBwMTJpZUhiZjdobmUrQ1E0em0yZVVvbnZCczI4RHdZVGs9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik16Rm1Oall4WWpJdE5UQTRaaTAwTlRVNUxXRmtOamN0T0RZeE5ESXlaamd4TURKbCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjQwLjEyNi4zMi45OSJ9.RkRCUmdqa21sd1VKaDdUQU9WcmhnRnd3K01yZDZXaUp4OFJDcXJET2c1MD0&ApiVersion=2.0",
                "createdDateTime": "2021-11-13T16:27:04Z",
                "eTag": "\"{1BC151A6-BEB8-4034-BE93-6E9A18AA6CDC},1\"",
                "id": "01472SYFFGKHARXOF6GRAL5E3OTIMKU3G4",
                "lastModifiedDateTime": "2021-11-13T16:27:04Z",
                "name": "Blog Post preview.docx",
                "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B1BC151A6-BEB8-4034-BE93-6E9A18AA6CDC%7D&file=Blog%20Post%20preview.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{1BC151A6-BEB8-4034-BE93-6E9A18AA6CDC},4\"",
                "size": 131072,
                "createdBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "parentReference": {
                  "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                  "driveType": "documentLibrary",
                  "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "fVSMG9+U0AczPgoLQYCHu85LOT4="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-11-13T16:27:04Z",
                  "lastModifiedDateTime": "2021-11-13T16:27:04Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=716e4cb5-cefc-4562-97bb-b40a985f3306&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMjc4MSIsImV4cCI6IjE2MzY4MjYzODEiLCJlbmRwb2ludHVybCI6InlrUjVQQjNjU3NIV1hhVXVwZW9qbGtna05UeW9VcFdweGxiS0FZbURwV1k9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik16Rm1Oall4WWpJdE5UQTRaaTAwTlRVNUxXRmtOamN0T0RZeE5ESXlaamd4TURKbCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjQwLjEyNi4zMi45OSJ9.NFhRbzBqUHNwRjdnbVphb3I0anF0NnY4NzJRUVpsNTZ2Yk9QeW1Pa2lTQT0&ApiVersion=2.0",
                "createdDateTime": "2021-11-13T16:27:05Z",
                "eTag": "\"{716E4CB5-CEFC-4562-97BB-B40A985F3306},2\"",
                "id": "01472SYFFVJRXHD7GOMJCZPO5UBKMF6MYG",
                "lastModifiedDateTime": "2021-11-13T16:27:05Z",
                "name": "Contoso Purchasing Permissions.docx",
                "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B716E4CB5-CEFC-4562-97BB-B40A985F3306%7D&file=Contoso%20Purchasing%20Permissions.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{716E4CB5-CEFC-4562-97BB-B40A985F3306},5\"",
                "dataLossPrevention": {
                  "block": {},
                  "notify": {}
                },
                "size": 27514,
                "createdBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "parentReference": {
                  "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                  "driveType": "documentLibrary",
                  "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "7PDBFlBace6tv3M/yVpdtz8fm/c="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-11-13T16:27:05Z",
                  "lastModifiedDateTime": "2021-11-13T16:27:05Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=f123cf01-244e-44d3-b2f7-24f5230937a1&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMjc4MSIsImV4cCI6IjE2MzY4MjYzODEiLCJlbmRwb2ludHVybCI6IkcxMS8vN2M1Y1RyU1JwVEFZOUJSSG1WVUlYYTRNTXdnKzJRcTVFZE1OZ1k9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik16Rm1Oall4WWpJdE5UQTRaaTAwTlRVNUxXRmtOamN0T0RZeE5ESXlaamd4TURKbCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjQwLjEyNi4zMi45OSJ9.Y0d5Y1ZuZ05zZDlRRk9TYWliNWV4bXhuL2JhOWxVdnhCZk5iK0txVU9yRT0&ApiVersion=2.0",
                "createdDateTime": "2021-11-13T16:27:04Z",
                "eTag": "\"{F123CF01-244E-44D3-B2F7-24F5230937A1},1\"",
                "id": "01472SYFABZ4R7CTRE2NCLF5ZE6URQSN5B",
                "lastModifiedDateTime": "2021-11-13T16:27:04Z",
                "name": "Credit Cards.docx",
                "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7BF123CF01-244E-44D3-B2F7-24F5230937A1%7D&file=Credit%20Cards.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{F123CF01-244E-44D3-B2F7-24F5230937A1},4\"",
                "size": 24858,
                "createdBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "parentReference": {
                  "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                  "driveType": "documentLibrary",
                  "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "9LEEpPZjBLjJQmN/AKDXRLmXj/k="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-11-13T16:27:04Z",
                  "lastModifiedDateTime": "2021-11-13T16:27:04Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=9b9a62cf-e7c6-4fa4-a261-1b260d7d6cf1&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMjc4MSIsImV4cCI6IjE2MzY4MjYzODEiLCJlbmRwb2ludHVybCI6ImordllIUXB6d1hhM01HV3lJb3JuZWRzZlpXbGZIQXlsSmlsQTE2NU1HMms9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik16Rm1Oall4WWpJdE5UQTRaaTAwTlRVNUxXRmtOamN0T0RZeE5ESXlaamd4TURKbCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjQwLjEyNi4zMi45OSJ9.SmdsZHBMQ1poS3JVdnBSbVN0MzE4R2c4RWJTY29TZktLemtJYlRKVElGaz0&ApiVersion=2.0",
                "createdDateTime": "2021-11-13T16:27:05Z",
                "eTag": "\"{9B9A62CF-E7C6-4FA4-A261-1B260D7D6CF1},1\"",
                "id": "01472SYFGPMKNJXRXHURH2EYI3EYGX23HR",
                "lastModifiedDateTime": "2021-11-13T16:27:05Z",
                "name": "Customer Accounts.docx",
                "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B9B9A62CF-E7C6-4FA4-A261-1B260D7D6CF1%7D&file=Customer%20Accounts.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{9B9A62CF-E7C6-4FA4-A261-1B260D7D6CF1},4\"",
                "size": 28360,
                "createdBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "parentReference": {
                  "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                  "driveType": "documentLibrary",
                  "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "PTFegvh8VBE1mJ3eELTSb+LSQxI="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-11-13T16:27:05Z",
                  "lastModifiedDateTime": "2021-11-13T16:27:05Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=e2f09ed9-987a-4d26-8513-86878f785e0a&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMjc4MSIsImV4cCI6IjE2MzY4MjYzODEiLCJlbmRwb2ludHVybCI6IkRORDROcGlaQXAxa3RCR21HckRFa0ZScmdEQWVyeVE0TjgyK1dBNWJDaVE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik16Rm1Oall4WWpJdE5UQTRaaTAwTlRVNUxXRmtOamN0T0RZeE5ESXlaamd4TURKbCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjQwLjEyNi4zMi45OSJ9.NGM5MkoyWHk5M2xCZTBiN0twVXdXU0c5ZEVlbW1SRUxzZ1hwanp2Wm5FZz0&ApiVersion=2.0",
                "createdDateTime": "2021-11-13T16:27:04Z",
                "eTag": "\"{E2F09ED9-987A-4D26-8513-86878F785E0A},1\"",
                "id": "01472SYFGZT3YOE6UYEZGYKE4GQ6HXQXQK",
                "lastModifiedDateTime": "2021-11-13T16:27:04Z",
                "name": "Customer Data.xlsx",
                "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7BE2F09ED9-987A-4D26-8513-86878F785E0A%7D&file=Customer%20Data.xlsx&action=default&mobileredirect=true",
                "cTag": "\"c:{E2F09ED9-987A-4D26-8513-86878F785E0A},4\"",
                "size": 17448,
                "createdBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "parentReference": {
                  "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                  "driveType": "documentLibrary",
                  "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                  "hashes": {
                    "quickXorHash": "hrMOe12SPOwumxGnhdJiYnw1qoM="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-11-13T16:27:04Z",
                  "lastModifiedDateTime": "2021-11-13T16:27:04Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=370c45c8-cca6-4333-9cb4-d33b7bc3fc05&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMjc4MSIsImV4cCI6IjE2MzY4MjYzODEiLCJlbmRwb2ludHVybCI6InAzSHU2TjRHSjRSTW9sREVJOVRLdE9DVzk0SXUzMllyTWdVak9IbCtlWkU9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik16Rm1Oall4WWpJdE5UQTRaaTAwTlRVNUxXRmtOamN0T0RZeE5ESXlaamd4TURKbCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjQwLjEyNi4zMi45OSJ9.TnVVZkxGWUJ6dVRrNWJuZDJsU3RzZnVLNUR3azRLRGhZRm42RW5tY1ZSTT0&ApiVersion=2.0",
                "createdDateTime": "2021-11-13T16:27:05Z",
                "eTag": "\"{370C45C8-CCA6-4333-9CB4-D33B7BC3FC05},1\"",
                "id": "01472SYFGIIUGDPJWMGNBZZNGTHN54H7AF",
                "lastModifiedDateTime": "2021-11-13T16:27:05Z",
                "name": "Q3 Sales and Marketing Expense Report Audit.pptx",
                "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B370C45C8-CCA6-4333-9CB4-D33B7BC3FC05%7D&file=Q3%20Sales%20and%20Marketing%20Expense%20Report%20Audit.pptx&action=edit&mobileredirect=true",
                "cTag": "\"c:{370C45C8-CCA6-4333-9CB4-D33B7BC3FC05},4\"",
                "size": 375938,
                "createdBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "parentReference": {
                  "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                  "driveType": "documentLibrary",
                  "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                  "hashes": {
                    "quickXorHash": "/qGrwC7X0bUey5sxCDlzxzSjTLE="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-11-13T16:27:05Z",
                  "lastModifiedDateTime": "2021-11-13T16:27:05Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=2223f837-ce90-4638-a2ca-8fd753c4378d&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMjc4MSIsImV4cCI6IjE2MzY4MjYzODEiLCJlbmRwb2ludHVybCI6InhuVmpzU1NWUWxTWVhWK3dDNElpVmtpMUIrZ0xTYWFpTldTaEpQemFnaFk9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik16Rm1Oall4WWpJdE5UQTRaaTAwTlRVNUxXRmtOamN0T0RZeE5ESXlaamd4TURKbCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjQwLjEyNi4zMi45OSJ9.cU1FaENWYW5PWVh0dHErZ2FLWEFPWFBQam12c1dublVEbUc0U0dSWnNTMD0&ApiVersion=2.0",
                "createdDateTime": "2021-11-13T16:27:04Z",
                "eTag": "\"{2223F837-CE90-4638-A2CA-8FD753C4378D},1\"",
                "id": "01472SYFBX7ARSFEGOHBDKFSUP25J4IN4N",
                "lastModifiedDateTime": "2021-11-13T16:27:04Z",
                "name": "Q3_Product_Strategy.docx",
                "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B2223F837-CE90-4638-A2CA-8FD753C4378D%7D&file=Q3_Product_Strategy.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{2223F837-CE90-4638-A2CA-8FD753C4378D},4\"",
                "size": 48086,
                "createdBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "parentReference": {
                  "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                  "driveType": "documentLibrary",
                  "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "LWe4j5GTM+Sg3PGOV0xoSSCnTr8="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-11-13T16:27:04Z",
                  "lastModifiedDateTime": "2021-11-13T16:27:04Z"
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=bfc3a23e-ed0b-4a59-96f1-b11c075d89b0&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMjc4MSIsImV4cCI6IjE2MzY4MjYzODEiLCJlbmRwb2ludHVybCI6IjVBMkhoblFOcEozMFlsZUp3NDU2ZENCTUlJNXlXbnRoaFhGVGZZQUZuMXM9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik16Rm1Oall4WWpJdE5UQTRaaTAwTlRVNUxXRmtOamN0T0RZeE5ESXlaamd4TURKbCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjQwLjEyNi4zMi45OSJ9.N3JrZldpUnVQK09pTnl6RGJSY0pUdWZVbDhSSkl3R0l4ZkJESGJ6a3h3dz0&ApiVersion=2.0",
                "createdDateTime": "2021-11-13T16:27:05Z",
                "eTag": "\"{BFC3A23E-ED0B-4A59-96F1-B11C075D89B0},1\"",
                "id": "01472SYFB6ULB36C7NLFFJN4NRDQDV3CNQ",
                "lastModifiedDateTime": "2021-11-13T16:27:05Z",
                "name": "Sales Memo.docx",
                "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7BBFC3A23E-ED0B-4A59-96F1-B11C075D89B0%7D&file=Sales%20Memo.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{BFC3A23E-ED0B-4A59-96F1-B11C075D89B0},4\"",
                "size": 35655,
                "createdBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "parentReference": {
                  "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                  "driveType": "documentLibrary",
                  "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "4qr1R9xcSfoofaFA3nMY6qjQXPQ="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-11-13T16:27:05Z",
                  "lastModifiedDateTime": "2021-11-13T16:27:05Z"
                }
              }
            ]
          });
        case 'https://graph.microsoft.com/v1.0/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/items/01472SYFERKSBJFD7X35H3TUHUSHXZBRRD/children':
          return Promise.resolve({
            "value": [
              {
                "createdDateTime": "2021-11-14T14:13:41Z",
                "eTag": "\"{294F0195-ECCC-43AC-9419-8E058F3D5506},1\"",
                "id": "01472SYFEVAFHSTTHMVRBZIGMOAWHT2VIG",
                "lastModifiedDateTime": "2021-11-14T14:13:41Z",
                "name": "Subfolder",
                "webUrl": "https://contoso.sharepoint.com/sites/Design/DemoDocs/Folder/Subfolder",
                "cTag": "\"c:{294F0195-ECCC-43AC-9419-8E058F3D5506},0\"",
                "size": 0,
                "createdBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "parentReference": {
                  "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                  "driveType": "documentLibrary",
                  "id": "01472SYFERKSBJFD7X35H3TUHUSHXZBRRD",
                  "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:/Folder"
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-11-14T14:13:41Z",
                  "lastModifiedDateTime": "2021-11-14T14:13:41Z"
                },
                "folder": {
                  "childCount": 1
                }
              },
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=77e5e9f4-4731-478e-82ae-6eece079ae8b&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMjc4MSIsImV4cCI6IjE2MzY4MjYzODEiLCJlbmRwb2ludHVybCI6IlFWaVJuSEVIRWRHNW5vSHhrOUFwQ3lXS3JKa05uL3pFVFhqT3NGWGpmQkE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik5tSmtPRFppTm1JdE5XWXhaaTAwTW1VNExXRTVOMkV0TXpsaVpqTXhNR0U1TURaaCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjQwLjEyNi4zMi45OSJ9.c0FyM3hCR2pYQzdxQ0xEZ1VEamZJaDhFdHIzWFkxN0tWNGlNV2djcktRRT0&ApiVersion=2.0",
                "createdDateTime": "2021-11-13T16:40:39Z",
                "eTag": "\"{77E5E9F4-4731-478E-82AE-6EECE079AE8B},2\"",
                "id": "01472SYFHU5HSXOMKHRZDYFLTO5TQHTLUL",
                "lastModifiedDateTime": "2021-11-13T16:40:39Z",
                "name": "Blog Post preview.docx",
                "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B77E5E9F4-4731-478E-82AE-6EECE079AE8B%7D&file=Blog%20Post%20preview.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{77E5E9F4-4731-478E-82AE-6EECE079AE8B},4\"",
                "size": 131072,
                "createdBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "parentReference": {
                  "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                  "driveType": "documentLibrary",
                  "id": "01472SYFERKSBJFD7X35H3TUHUSHXZBRRD",
                  "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:/Folder"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-11-13T16:40:39Z",
                  "lastModifiedDateTime": "2021-11-13T16:40:39Z"
                }
              }
            ]
          });
        case 'https://graph.microsoft.com/v1.0/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/items/01472SYFEVAFHSTTHMVRBZIGMOAWHT2VIG/children':
          return Promise.resolve({
            "value": [
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=4eac39e8-c156-4d0b-a8eb-f86b77f85324&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjg5OTI0NCIsImV4cCI6IjE2MzY5MDI4NDQiLCJlbmRwb2ludHVybCI6IkdsQjZSNnhDRlNub0plMjdKNGJaczVMc3NSdDRvS213Z3I4VlZBZUVWbDA9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1UZGlaakEyWlRjdFlqSmpOQzAwWWpkaExXSTVZVEF0TXpJd01qTmxOR1ZrWXpjNSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjUifQ.Nyt3NXpncGdXdlFhVkxiWVVUY3Z3OXBMVWUyb0FPUUh3NTFEd0J1MXJpWT0&ApiVersion=2.0",
                "createdDateTime": "2021-11-14T14:13:52Z",
                "eTag": "\"{4EAC39E8-C156-4D0B-A8EB-F86B77F85324},2\"",
                "id": "01472SYFHIHGWE4VWBBNG2R27YNN37QUZE",
                "lastModifiedDateTime": "2021-11-14T14:13:52Z",
                "name": "Credit Cards.docx",
                "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B4EAC39E8-C156-4D0B-A8EB-F86B77F85324%7D&file=Credit%20Cards.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{4EAC39E8-C156-4D0B-A8EB-F86B77F85324},4\"",
                "size": 24858,
                "createdBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "lastModifiedBy": {
                  "user": {
                    "email": "admin@contoso.OnMicrosoft.com",
                    "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                    "displayName": "MOD Administrator"
                  }
                },
                "parentReference": {
                  "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                  "driveType": "documentLibrary",
                  "id": "01472SYFEVAFHSTTHMVRBZIGMOAWHT2VIG",
                  "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:/Folder/Subfolder"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-11-14T14:13:52Z",
                  "lastModifiedDateTime": "2021-11-14T14:13:52Z"
                }
              }
            ]
          });
        default:
          return Promise.reject(`Invalid GET request: ${url}`);
      }
    });

    command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/design',
        folderUrl: 'DemoDocs',
        recursive: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        assert(loggerLogSpy.calledWith([
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=1bc151a6-beb8-4034-be93-6e9a18aa6cdc&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMjc4MSIsImV4cCI6IjE2MzY4MjYzODEiLCJlbmRwb2ludHVybCI6Ik5PNFErS3pZSFBwMTJpZUhiZjdobmUrQ1E0em0yZVVvbnZCczI4RHdZVGs9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik16Rm1Oall4WWpJdE5UQTRaaTAwTlRVNUxXRmtOamN0T0RZeE5ESXlaamd4TURKbCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjQwLjEyNi4zMi45OSJ9.RkRCUmdqa21sd1VKaDdUQU9WcmhnRnd3K01yZDZXaUp4OFJDcXJET2c1MD0&ApiVersion=2.0",
            "createdDateTime": "2021-11-13T16:27:04Z",
            "eTag": "\"{1BC151A6-BEB8-4034-BE93-6E9A18AA6CDC},1\"",
            "id": "01472SYFFGKHARXOF6GRAL5E3OTIMKU3G4",
            "lastModifiedDateTime": "2021-11-13T16:27:04Z",
            "name": "Blog Post preview.docx",
            "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B1BC151A6-BEB8-4034-BE93-6E9A18AA6CDC%7D&file=Blog%20Post%20preview.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{1BC151A6-BEB8-4034-BE93-6E9A18AA6CDC},4\"",
            "size": 131072,
            "createdBy": {
              "user": {
                "email": "admin@contoso.OnMicrosoft.com",
                "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                "displayName": "MOD Administrator"
              }
            },
            "lastModifiedBy": {
              "user": {
                "email": "admin@contoso.OnMicrosoft.com",
                "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                "displayName": "MOD Administrator"
              }
            },
            "parentReference": {
              "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
              "driveType": "documentLibrary",
              "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "fVSMG9+U0AczPgoLQYCHu85LOT4="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-11-13T16:27:04Z",
              "lastModifiedDateTime": "2021-11-13T16:27:04Z"
            },
            "lastModifiedByUser": "MOD Administrator"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=716e4cb5-cefc-4562-97bb-b40a985f3306&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMjc4MSIsImV4cCI6IjE2MzY4MjYzODEiLCJlbmRwb2ludHVybCI6InlrUjVQQjNjU3NIV1hhVXVwZW9qbGtna05UeW9VcFdweGxiS0FZbURwV1k9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik16Rm1Oall4WWpJdE5UQTRaaTAwTlRVNUxXRmtOamN0T0RZeE5ESXlaamd4TURKbCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjQwLjEyNi4zMi45OSJ9.NFhRbzBqUHNwRjdnbVphb3I0anF0NnY4NzJRUVpsNTZ2Yk9QeW1Pa2lTQT0&ApiVersion=2.0",
            "createdDateTime": "2021-11-13T16:27:05Z",
            "eTag": "\"{716E4CB5-CEFC-4562-97BB-B40A985F3306},2\"",
            "id": "01472SYFFVJRXHD7GOMJCZPO5UBKMF6MYG",
            "lastModifiedDateTime": "2021-11-13T16:27:05Z",
            "name": "Contoso Purchasing Permissions.docx",
            "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B716E4CB5-CEFC-4562-97BB-B40A985F3306%7D&file=Contoso%20Purchasing%20Permissions.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{716E4CB5-CEFC-4562-97BB-B40A985F3306},5\"",
            "dataLossPrevention": {
              "block": {},
              "notify": {}
            },
            "size": 27514,
            "createdBy": {
              "user": {
                "email": "admin@contoso.OnMicrosoft.com",
                "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                "displayName": "MOD Administrator"
              }
            },
            "lastModifiedBy": {
              "user": {
                "email": "admin@contoso.OnMicrosoft.com",
                "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                "displayName": "MOD Administrator"
              }
            },
            "parentReference": {
              "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
              "driveType": "documentLibrary",
              "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "7PDBFlBace6tv3M/yVpdtz8fm/c="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-11-13T16:27:05Z",
              "lastModifiedDateTime": "2021-11-13T16:27:05Z"
            },
            "lastModifiedByUser": "MOD Administrator"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=f123cf01-244e-44d3-b2f7-24f5230937a1&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMjc4MSIsImV4cCI6IjE2MzY4MjYzODEiLCJlbmRwb2ludHVybCI6IkcxMS8vN2M1Y1RyU1JwVEFZOUJSSG1WVUlYYTRNTXdnKzJRcTVFZE1OZ1k9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik16Rm1Oall4WWpJdE5UQTRaaTAwTlRVNUxXRmtOamN0T0RZeE5ESXlaamd4TURKbCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjQwLjEyNi4zMi45OSJ9.Y0d5Y1ZuZ05zZDlRRk9TYWliNWV4bXhuL2JhOWxVdnhCZk5iK0txVU9yRT0&ApiVersion=2.0",
            "createdDateTime": "2021-11-13T16:27:04Z",
            "eTag": "\"{F123CF01-244E-44D3-B2F7-24F5230937A1},1\"",
            "id": "01472SYFABZ4R7CTRE2NCLF5ZE6URQSN5B",
            "lastModifiedDateTime": "2021-11-13T16:27:04Z",
            "name": "Credit Cards.docx",
            "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7BF123CF01-244E-44D3-B2F7-24F5230937A1%7D&file=Credit%20Cards.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{F123CF01-244E-44D3-B2F7-24F5230937A1},4\"",
            "size": 24858,
            "createdBy": {
              "user": {
                "email": "admin@contoso.OnMicrosoft.com",
                "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                "displayName": "MOD Administrator"
              }
            },
            "lastModifiedBy": {
              "user": {
                "email": "admin@contoso.OnMicrosoft.com",
                "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                "displayName": "MOD Administrator"
              }
            },
            "parentReference": {
              "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
              "driveType": "documentLibrary",
              "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "9LEEpPZjBLjJQmN/AKDXRLmXj/k="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-11-13T16:27:04Z",
              "lastModifiedDateTime": "2021-11-13T16:27:04Z"
            },
            "lastModifiedByUser": "MOD Administrator"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=9b9a62cf-e7c6-4fa4-a261-1b260d7d6cf1&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMjc4MSIsImV4cCI6IjE2MzY4MjYzODEiLCJlbmRwb2ludHVybCI6ImordllIUXB6d1hhM01HV3lJb3JuZWRzZlpXbGZIQXlsSmlsQTE2NU1HMms9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik16Rm1Oall4WWpJdE5UQTRaaTAwTlRVNUxXRmtOamN0T0RZeE5ESXlaamd4TURKbCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjQwLjEyNi4zMi45OSJ9.SmdsZHBMQ1poS3JVdnBSbVN0MzE4R2c4RWJTY29TZktLemtJYlRKVElGaz0&ApiVersion=2.0",
            "createdDateTime": "2021-11-13T16:27:05Z",
            "eTag": "\"{9B9A62CF-E7C6-4FA4-A261-1B260D7D6CF1},1\"",
            "id": "01472SYFGPMKNJXRXHURH2EYI3EYGX23HR",
            "lastModifiedDateTime": "2021-11-13T16:27:05Z",
            "name": "Customer Accounts.docx",
            "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B9B9A62CF-E7C6-4FA4-A261-1B260D7D6CF1%7D&file=Customer%20Accounts.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{9B9A62CF-E7C6-4FA4-A261-1B260D7D6CF1},4\"",
            "size": 28360,
            "createdBy": {
              "user": {
                "email": "admin@contoso.OnMicrosoft.com",
                "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                "displayName": "MOD Administrator"
              }
            },
            "lastModifiedBy": {
              "user": {
                "email": "admin@contoso.OnMicrosoft.com",
                "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                "displayName": "MOD Administrator"
              }
            },
            "parentReference": {
              "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
              "driveType": "documentLibrary",
              "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "PTFegvh8VBE1mJ3eELTSb+LSQxI="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-11-13T16:27:05Z",
              "lastModifiedDateTime": "2021-11-13T16:27:05Z"
            },
            "lastModifiedByUser": "MOD Administrator"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=e2f09ed9-987a-4d26-8513-86878f785e0a&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMjc4MSIsImV4cCI6IjE2MzY4MjYzODEiLCJlbmRwb2ludHVybCI6IkRORDROcGlaQXAxa3RCR21HckRFa0ZScmdEQWVyeVE0TjgyK1dBNWJDaVE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik16Rm1Oall4WWpJdE5UQTRaaTAwTlRVNUxXRmtOamN0T0RZeE5ESXlaamd4TURKbCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjQwLjEyNi4zMi45OSJ9.NGM5MkoyWHk5M2xCZTBiN0twVXdXU0c5ZEVlbW1SRUxzZ1hwanp2Wm5FZz0&ApiVersion=2.0",
            "createdDateTime": "2021-11-13T16:27:04Z",
            "eTag": "\"{E2F09ED9-987A-4D26-8513-86878F785E0A},1\"",
            "id": "01472SYFGZT3YOE6UYEZGYKE4GQ6HXQXQK",
            "lastModifiedDateTime": "2021-11-13T16:27:04Z",
            "name": "Customer Data.xlsx",
            "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7BE2F09ED9-987A-4D26-8513-86878F785E0A%7D&file=Customer%20Data.xlsx&action=default&mobileredirect=true",
            "cTag": "\"c:{E2F09ED9-987A-4D26-8513-86878F785E0A},4\"",
            "size": 17448,
            "createdBy": {
              "user": {
                "email": "admin@contoso.OnMicrosoft.com",
                "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                "displayName": "MOD Administrator"
              }
            },
            "lastModifiedBy": {
              "user": {
                "email": "admin@contoso.OnMicrosoft.com",
                "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                "displayName": "MOD Administrator"
              }
            },
            "parentReference": {
              "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
              "driveType": "documentLibrary",
              "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              "hashes": {
                "quickXorHash": "hrMOe12SPOwumxGnhdJiYnw1qoM="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-11-13T16:27:04Z",
              "lastModifiedDateTime": "2021-11-13T16:27:04Z"
            },
            "lastModifiedByUser": "MOD Administrator"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=370c45c8-cca6-4333-9cb4-d33b7bc3fc05&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMjc4MSIsImV4cCI6IjE2MzY4MjYzODEiLCJlbmRwb2ludHVybCI6InAzSHU2TjRHSjRSTW9sREVJOVRLdE9DVzk0SXUzMllyTWdVak9IbCtlWkU9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik16Rm1Oall4WWpJdE5UQTRaaTAwTlRVNUxXRmtOamN0T0RZeE5ESXlaamd4TURKbCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjQwLjEyNi4zMi45OSJ9.TnVVZkxGWUJ6dVRrNWJuZDJsU3RzZnVLNUR3azRLRGhZRm42RW5tY1ZSTT0&ApiVersion=2.0",
            "createdDateTime": "2021-11-13T16:27:05Z",
            "eTag": "\"{370C45C8-CCA6-4333-9CB4-D33B7BC3FC05},1\"",
            "id": "01472SYFGIIUGDPJWMGNBZZNGTHN54H7AF",
            "lastModifiedDateTime": "2021-11-13T16:27:05Z",
            "name": "Q3 Sales and Marketing Expense Report Audit.pptx",
            "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B370C45C8-CCA6-4333-9CB4-D33B7BC3FC05%7D&file=Q3%20Sales%20and%20Marketing%20Expense%20Report%20Audit.pptx&action=edit&mobileredirect=true",
            "cTag": "\"c:{370C45C8-CCA6-4333-9CB4-D33B7BC3FC05},4\"",
            "size": 375938,
            "createdBy": {
              "user": {
                "email": "admin@contoso.OnMicrosoft.com",
                "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                "displayName": "MOD Administrator"
              }
            },
            "lastModifiedBy": {
              "user": {
                "email": "admin@contoso.OnMicrosoft.com",
                "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                "displayName": "MOD Administrator"
              }
            },
            "parentReference": {
              "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
              "driveType": "documentLibrary",
              "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
              "hashes": {
                "quickXorHash": "/qGrwC7X0bUey5sxCDlzxzSjTLE="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-11-13T16:27:05Z",
              "lastModifiedDateTime": "2021-11-13T16:27:05Z"
            },
            "lastModifiedByUser": "MOD Administrator"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=2223f837-ce90-4638-a2ca-8fd753c4378d&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMjc4MSIsImV4cCI6IjE2MzY4MjYzODEiLCJlbmRwb2ludHVybCI6InhuVmpzU1NWUWxTWVhWK3dDNElpVmtpMUIrZ0xTYWFpTldTaEpQemFnaFk9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik16Rm1Oall4WWpJdE5UQTRaaTAwTlRVNUxXRmtOamN0T0RZeE5ESXlaamd4TURKbCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjQwLjEyNi4zMi45OSJ9.cU1FaENWYW5PWVh0dHErZ2FLWEFPWFBQam12c1dublVEbUc0U0dSWnNTMD0&ApiVersion=2.0",
            "createdDateTime": "2021-11-13T16:27:04Z",
            "eTag": "\"{2223F837-CE90-4638-A2CA-8FD753C4378D},1\"",
            "id": "01472SYFBX7ARSFEGOHBDKFSUP25J4IN4N",
            "lastModifiedDateTime": "2021-11-13T16:27:04Z",
            "name": "Q3_Product_Strategy.docx",
            "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B2223F837-CE90-4638-A2CA-8FD753C4378D%7D&file=Q3_Product_Strategy.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{2223F837-CE90-4638-A2CA-8FD753C4378D},4\"",
            "size": 48086,
            "createdBy": {
              "user": {
                "email": "admin@contoso.OnMicrosoft.com",
                "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                "displayName": "MOD Administrator"
              }
            },
            "lastModifiedBy": {
              "user": {
                "email": "admin@contoso.OnMicrosoft.com",
                "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                "displayName": "MOD Administrator"
              }
            },
            "parentReference": {
              "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
              "driveType": "documentLibrary",
              "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "LWe4j5GTM+Sg3PGOV0xoSSCnTr8="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-11-13T16:27:04Z",
              "lastModifiedDateTime": "2021-11-13T16:27:04Z"
            },
            "lastModifiedByUser": "MOD Administrator"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=bfc3a23e-ed0b-4a59-96f1-b11c075d89b0&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMjc4MSIsImV4cCI6IjE2MzY4MjYzODEiLCJlbmRwb2ludHVybCI6IjVBMkhoblFOcEozMFlsZUp3NDU2ZENCTUlJNXlXbnRoaFhGVGZZQUZuMXM9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik16Rm1Oall4WWpJdE5UQTRaaTAwTlRVNUxXRmtOamN0T0RZeE5ESXlaamd4TURKbCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjQwLjEyNi4zMi45OSJ9.N3JrZldpUnVQK09pTnl6RGJSY0pUdWZVbDhSSkl3R0l4ZkJESGJ6a3h3dz0&ApiVersion=2.0",
            "createdDateTime": "2021-11-13T16:27:05Z",
            "eTag": "\"{BFC3A23E-ED0B-4A59-96F1-B11C075D89B0},1\"",
            "id": "01472SYFB6ULB36C7NLFFJN4NRDQDV3CNQ",
            "lastModifiedDateTime": "2021-11-13T16:27:05Z",
            "name": "Sales Memo.docx",
            "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7BBFC3A23E-ED0B-4A59-96F1-B11C075D89B0%7D&file=Sales%20Memo.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{BFC3A23E-ED0B-4A59-96F1-B11C075D89B0},4\"",
            "size": 35655,
            "createdBy": {
              "user": {
                "email": "admin@contoso.OnMicrosoft.com",
                "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                "displayName": "MOD Administrator"
              }
            },
            "lastModifiedBy": {
              "user": {
                "email": "admin@contoso.OnMicrosoft.com",
                "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                "displayName": "MOD Administrator"
              }
            },
            "parentReference": {
              "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
              "driveType": "documentLibrary",
              "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "4qr1R9xcSfoofaFA3nMY6qjQXPQ="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-11-13T16:27:05Z",
              "lastModifiedDateTime": "2021-11-13T16:27:05Z"
            },
            "lastModifiedByUser": "MOD Administrator"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=77e5e9f4-4731-478e-82ae-6eece079ae8b&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMjc4MSIsImV4cCI6IjE2MzY4MjYzODEiLCJlbmRwb2ludHVybCI6IlFWaVJuSEVIRWRHNW5vSHhrOUFwQ3lXS3JKa05uL3pFVFhqT3NGWGpmQkE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik5tSmtPRFppTm1JdE5XWXhaaTAwTW1VNExXRTVOMkV0TXpsaVpqTXhNR0U1TURaaCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjQwLjEyNi4zMi45OSJ9.c0FyM3hCR2pYQzdxQ0xEZ1VEamZJaDhFdHIzWFkxN0tWNGlNV2djcktRRT0&ApiVersion=2.0",
            "createdDateTime": "2021-11-13T16:40:39Z",
            "eTag": "\"{77E5E9F4-4731-478E-82AE-6EECE079AE8B},2\"",
            "id": "01472SYFHU5HSXOMKHRZDYFLTO5TQHTLUL",
            "lastModifiedDateTime": "2021-11-13T16:40:39Z",
            "name": "Blog Post preview.docx",
            "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B77E5E9F4-4731-478E-82AE-6EECE079AE8B%7D&file=Blog%20Post%20preview.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{77E5E9F4-4731-478E-82AE-6EECE079AE8B},4\"",
            "size": 131072,
            "createdBy": {
              "user": {
                "email": "admin@contoso.OnMicrosoft.com",
                "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                "displayName": "MOD Administrator"
              }
            },
            "lastModifiedBy": {
              "user": {
                "email": "admin@contoso.OnMicrosoft.com",
                "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                "displayName": "MOD Administrator"
              }
            },
            "parentReference": {
              "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
              "driveType": "documentLibrary",
              "id": "01472SYFERKSBJFD7X35H3TUHUSHXZBRRD",
              "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:/Folder"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-11-13T16:40:39Z",
              "lastModifiedDateTime": "2021-11-13T16:40:39Z"
            },
            "lastModifiedByUser": "MOD Administrator"
          },
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=4eac39e8-c156-4d0b-a8eb-f86b77f85324&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjg5OTI0NCIsImV4cCI6IjE2MzY5MDI4NDQiLCJlbmRwb2ludHVybCI6IkdsQjZSNnhDRlNub0plMjdKNGJaczVMc3NSdDRvS213Z3I4VlZBZUVWbDA9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1UZGlaakEyWlRjdFlqSmpOQzAwWWpkaExXSTVZVEF0TXpJd01qTmxOR1ZrWXpjNSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjUifQ.Nyt3NXpncGdXdlFhVkxiWVVUY3Z3OXBMVWUyb0FPUUh3NTFEd0J1MXJpWT0&ApiVersion=2.0",
            "createdDateTime": "2021-11-14T14:13:52Z",
            "eTag": "\"{4EAC39E8-C156-4D0B-A8EB-F86B77F85324},2\"",
            "id": "01472SYFHIHGWE4VWBBNG2R27YNN37QUZE",
            "lastModifiedDateTime": "2021-11-14T14:13:52Z",
            "name": "Credit Cards.docx",
            "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B4EAC39E8-C156-4D0B-A8EB-F86B77F85324%7D&file=Credit%20Cards.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{4EAC39E8-C156-4D0B-A8EB-F86B77F85324},4\"",
            "size": 24858,
            "createdBy": {
              "user": {
                "email": "admin@contoso.OnMicrosoft.com",
                "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                "displayName": "MOD Administrator"
              }
            },
            "lastModifiedBy": {
              "user": {
                "email": "admin@contoso.OnMicrosoft.com",
                "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                "displayName": "MOD Administrator"
              }
            },
            "parentReference": {
              "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
              "driveType": "documentLibrary",
              "id": "01472SYFEVAFHSTTHMVRBZIGMOAWHT2VIG",
              "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:/Folder/Subfolder"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-11-14T14:13:52Z",
              "lastModifiedDateTime": "2021-11-14T14:13:52Z"
            },
            "lastModifiedByUser": "MOD Administrator"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('loads files with paging', done => {
    sinon.stub(request, 'get').callsFake(opts => {
      const url: string = opts.url as string;

      switch (url) {
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/design?$select=id':
          return Promise.resolve({
            "id": "contoso.sharepoint.com,c88ba08f-4e17-49c1-a34f-ca85d908c24c,fde71d4c-ac91-4464-a60f-632d99b8224d"
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,c88ba08f-4e17-49c1-a34f-ca85d908c24c,fde71d4c-ac91-4464-a60f-632d99b8224d/drives?$select=webUrl,id':
          return Promise.resolve({
            "value": [{
              "id": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
              "webUrl": "https://contoso.sharepoint.com/sites/Design/DemoDocs"
            },
            {
              "id": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik0at4LaajDdTo6njHc5dx7g",
              "webUrl": "https://contoso.sharepoint.com/sites/Design/Shared%20Documents"
            }]
          });
        case 'https://graph.microsoft.com/v1.0/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root?$select=id':
          return Promise.resolve({
            "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ"
          });
        case 'https://graph.microsoft.com/v1.0/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/items/01472SYFF6Y2GOVW7725BZO354PWSELRRZ/children':
          return Promise.resolve({
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/items/01472SYFF6Y2GOVW7725BZO354PWSELRRZ/children?$skiptoken=UGFnZWQ9VFJVRSZwX1NvcnRCZWhhdmlvcj0xJnBfRmlsZUxlYWZSZWY9Rm9sZGVyJnBfSUQ9MTc",
            "value": [{
              "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=1bc151a6-beb8-4034-be93-6e9a18aa6cdc&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6Ik5PNFErS3pZSFBwMTJpZUhiZjdobmUrQ1E0em0yZVVvbnZCczI4RHdZVGs9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.d0hGUFpuMVhJbGZGSVFzcEhPSjJyS1FIUStXbEtLL3RURW9qVUhwZWNKWT0&ApiVersion=2.0",
              "createdDateTime": "2021-11-13T16:27:04Z",
              "eTag": "\"{1BC151A6-BEB8-4034-BE93-6E9A18AA6CDC},1\"",
              "id": "01472SYFFGKHARXOF6GRAL5E3OTIMKU3G4",
              "lastModifiedDateTime": "2021-11-13T16:27:04Z",
              "name": "Blog Post preview.docx",
              "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B1BC151A6-BEB8-4034-BE93-6E9A18AA6CDC%7D&file=Blog%20Post%20preview.docx&action=default&mobileredirect=true",
              "cTag": "\"c:{1BC151A6-BEB8-4034-BE93-6E9A18AA6CDC},4\"",
              "size": 131072,
              "createdBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "lastModifiedBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "parentReference": {
                "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                "driveType": "documentLibrary",
                "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
              },
              "file": {
                "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "hashes": {
                  "quickXorHash": "fVSMG9+U0AczPgoLQYCHu85LOT4="
                }
              },
              "fileSystemInfo": {
                "createdDateTime": "2021-11-13T16:27:04Z",
                "lastModifiedDateTime": "2021-11-13T16:27:04Z"
              }
            },
            {
              "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=716e4cb5-cefc-4562-97bb-b40a985f3306&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6InlrUjVQQjNjU3NIV1hhVXVwZW9qbGtna05UeW9VcFdweGxiS0FZbURwV1k9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.STJ5ekhaUEswQm1uL05Vb0F3dGhLcVJ4ZzNldkpReTM2YWI0cDlEaXo0VT0&ApiVersion=2.0",
              "createdDateTime": "2021-11-13T16:27:05Z",
              "eTag": "\"{716E4CB5-CEFC-4562-97BB-B40A985F3306},1\"",
              "id": "01472SYFFVJRXHD7GOMJCZPO5UBKMF6MYG",
              "lastModifiedDateTime": "2021-11-13T16:27:05Z",
              "name": "Contoso Purchasing Permissions.docx",
              "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B716E4CB5-CEFC-4562-97BB-B40A985F3306%7D&file=Contoso%20Purchasing%20Permissions.docx&action=default&mobileredirect=true",
              "cTag": "\"c:{716E4CB5-CEFC-4562-97BB-B40A985F3306},4\"",
              "size": 27442,
              "createdBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "lastModifiedBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "parentReference": {
                "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                "driveType": "documentLibrary",
                "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
              },
              "file": {
                "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "hashes": {
                  "quickXorHash": "ksnaYSdMJkWPE/EBL9rJWsK7kHw="
                }
              },
              "fileSystemInfo": {
                "createdDateTime": "2021-11-13T16:27:05Z",
                "lastModifiedDateTime": "2021-11-13T16:27:05Z"
              }
            },
            {
              "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=f123cf01-244e-44d3-b2f7-24f5230937a1&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6IkcxMS8vN2M1Y1RyU1JwVEFZOUJSSG1WVUlYYTRNTXdnKzJRcTVFZE1OZ1k9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.RWRBMjZvRFVUdlRkS3JsZmFWb2xDZnJLREtsZ1hGeVEvYWcrWStxV29rST0&ApiVersion=2.0",
              "createdDateTime": "2021-11-13T16:27:04Z",
              "eTag": "\"{F123CF01-244E-44D3-B2F7-24F5230937A1},1\"",
              "id": "01472SYFABZ4R7CTRE2NCLF5ZE6URQSN5B",
              "lastModifiedDateTime": "2021-11-13T16:27:04Z",
              "name": "Credit Cards.docx",
              "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7BF123CF01-244E-44D3-B2F7-24F5230937A1%7D&file=Credit%20Cards.docx&action=default&mobileredirect=true",
              "cTag": "\"c:{F123CF01-244E-44D3-B2F7-24F5230937A1},4\"",
              "size": 24858,
              "createdBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "lastModifiedBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "parentReference": {
                "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                "driveType": "documentLibrary",
                "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
              },
              "file": {
                "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "hashes": {
                  "quickXorHash": "9LEEpPZjBLjJQmN/AKDXRLmXj/k="
                }
              },
              "fileSystemInfo": {
                "createdDateTime": "2021-11-13T16:27:04Z",
                "lastModifiedDateTime": "2021-11-13T16:27:04Z"
              }
            },
            {
              "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=9b9a62cf-e7c6-4fa4-a261-1b260d7d6cf1&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6ImordllIUXB6d1hhM01HV3lJb3JuZWRzZlpXbGZIQXlsSmlsQTE2NU1HMms9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.ck96aFYwWCtSMnFrMVV4QVFNQklhSkMxYzRUczcwTzdxeWVsQnRHWjN1VT0&ApiVersion=2.0",
              "createdDateTime": "2021-11-13T16:27:05Z",
              "eTag": "\"{9B9A62CF-E7C6-4FA4-A261-1B260D7D6CF1},1\"",
              "id": "01472SYFGPMKNJXRXHURH2EYI3EYGX23HR",
              "lastModifiedDateTime": "2021-11-13T16:27:05Z",
              "name": "Customer Accounts.docx",
              "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B9B9A62CF-E7C6-4FA4-A261-1B260D7D6CF1%7D&file=Customer%20Accounts.docx&action=default&mobileredirect=true",
              "cTag": "\"c:{9B9A62CF-E7C6-4FA4-A261-1B260D7D6CF1},4\"",
              "size": 28360,
              "createdBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "lastModifiedBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "parentReference": {
                "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                "driveType": "documentLibrary",
                "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
              },
              "file": {
                "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "hashes": {
                  "quickXorHash": "PTFegvh8VBE1mJ3eELTSb+LSQxI="
                }
              },
              "fileSystemInfo": {
                "createdDateTime": "2021-11-13T16:27:05Z",
                "lastModifiedDateTime": "2021-11-13T16:27:05Z"
              }
            }]
          });
        case 'https://graph.microsoft.com/v1.0/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/items/01472SYFF6Y2GOVW7725BZO354PWSELRRZ/children?$skiptoken=UGFnZWQ9VFJVRSZwX1NvcnRCZWhhdmlvcj0xJnBfRmlsZUxlYWZSZWY9Rm9sZGVyJnBfSUQ9MTc':
          return Promise.resolve({
            "value": [{
              "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=e2f09ed9-987a-4d26-8513-86878f785e0a&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6IkRORDROcGlaQXAxa3RCR21HckRFa0ZScmdEQWVyeVE0TjgyK1dBNWJDaVE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.c2dqT3krQnBMOUx2QXdUMHlKWGtVbGtxMGZldmUyQWVTSFhPeERUR2pSVT0&ApiVersion=2.0",
              "createdDateTime": "2021-11-13T16:27:04Z",
              "eTag": "\"{E2F09ED9-987A-4D26-8513-86878F785E0A},1\"",
              "id": "01472SYFGZT3YOE6UYEZGYKE4GQ6HXQXQK",
              "lastModifiedDateTime": "2021-11-13T16:27:04Z",
              "name": "Customer Data.xlsx",
              "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7BE2F09ED9-987A-4D26-8513-86878F785E0A%7D&file=Customer%20Data.xlsx&action=default&mobileredirect=true",
              "cTag": "\"c:{E2F09ED9-987A-4D26-8513-86878F785E0A},4\"",
              "size": 17448,
              "createdBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "lastModifiedBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "parentReference": {
                "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                "driveType": "documentLibrary",
                "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
              },
              "file": {
                "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "hashes": {
                  "quickXorHash": "hrMOe12SPOwumxGnhdJiYnw1qoM="
                }
              },
              "fileSystemInfo": {
                "createdDateTime": "2021-11-13T16:27:04Z",
                "lastModifiedDateTime": "2021-11-13T16:27:04Z"
              }
            },
            {
              "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=370c45c8-cca6-4333-9cb4-d33b7bc3fc05&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6InAzSHU2TjRHSjRSTW9sREVJOVRLdE9DVzk0SXUzMllyTWdVak9IbCtlWkU9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.K1JabWc3TDNGWHpWN0Q5anluUTN6eUpYajNQNFY0NEJMUFhaS2VGTTgvTT0&ApiVersion=2.0",
              "createdDateTime": "2021-11-13T16:27:05Z",
              "eTag": "\"{370C45C8-CCA6-4333-9CB4-D33B7BC3FC05},1\"",
              "id": "01472SYFGIIUGDPJWMGNBZZNGTHN54H7AF",
              "lastModifiedDateTime": "2021-11-13T16:27:05Z",
              "name": "Q3 Sales and Marketing Expense Report Audit.pptx",
              "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B370C45C8-CCA6-4333-9CB4-D33B7BC3FC05%7D&file=Q3%20Sales%20and%20Marketing%20Expense%20Report%20Audit.pptx&action=edit&mobileredirect=true",
              "cTag": "\"c:{370C45C8-CCA6-4333-9CB4-D33B7BC3FC05},4\"",
              "size": 375938,
              "createdBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "lastModifiedBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "parentReference": {
                "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                "driveType": "documentLibrary",
                "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
              },
              "file": {
                "mimeType": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                "hashes": {
                  "quickXorHash": "/qGrwC7X0bUey5sxCDlzxzSjTLE="
                }
              },
              "fileSystemInfo": {
                "createdDateTime": "2021-11-13T16:27:05Z",
                "lastModifiedDateTime": "2021-11-13T16:27:05Z"
              }
            },
            {
              "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=2223f837-ce90-4638-a2ca-8fd753c4378d&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6InhuVmpzU1NWUWxTWVhWK3dDNElpVmtpMUIrZ0xTYWFpTldTaEpQemFnaFk9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.a0xuN1JrbXlnUWl2cW1rSHhPY1NNcnA4akhxVkVxSnVzQkNWeStLWGlPdz0&ApiVersion=2.0",
              "createdDateTime": "2021-11-13T16:27:04Z",
              "eTag": "\"{2223F837-CE90-4638-A2CA-8FD753C4378D},1\"",
              "id": "01472SYFBX7ARSFEGOHBDKFSUP25J4IN4N",
              "lastModifiedDateTime": "2021-11-13T16:27:04Z",
              "name": "Q3_Product_Strategy.docx",
              "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B2223F837-CE90-4638-A2CA-8FD753C4378D%7D&file=Q3_Product_Strategy.docx&action=default&mobileredirect=true",
              "cTag": "\"c:{2223F837-CE90-4638-A2CA-8FD753C4378D},4\"",
              "size": 48086,
              "createdBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "lastModifiedBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "parentReference": {
                "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                "driveType": "documentLibrary",
                "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
              },
              "file": {
                "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "hashes": {
                  "quickXorHash": "LWe4j5GTM+Sg3PGOV0xoSSCnTr8="
                }
              },
              "fileSystemInfo": {
                "createdDateTime": "2021-11-13T16:27:04Z",
                "lastModifiedDateTime": "2021-11-13T16:27:04Z"
              }
            },
            {
              "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=bfc3a23e-ed0b-4a59-96f1-b11c075d89b0&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6IjVBMkhoblFOcEozMFlsZUp3NDU2ZENCTUlJNXlXbnRoaFhGVGZZQUZuMXM9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.MGd4U0wzSWR6S2grc3F2ODJHQncyRjY2RkkzR1FMdCtpSUo0Y1lhNjZtcz0&ApiVersion=2.0",
              "createdDateTime": "2021-11-13T16:27:05Z",
              "eTag": "\"{BFC3A23E-ED0B-4A59-96F1-B11C075D89B0},1\"",
              "id": "01472SYFB6ULB36C7NLFFJN4NRDQDV3CNQ",
              "lastModifiedDateTime": "2021-11-13T16:27:05Z",
              "name": "Sales Memo.docx",
              "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7BBFC3A23E-ED0B-4A59-96F1-B11C075D89B0%7D&file=Sales%20Memo.docx&action=default&mobileredirect=true",
              "cTag": "\"c:{BFC3A23E-ED0B-4A59-96F1-B11C075D89B0},4\"",
              "size": 35655,
              "createdBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "lastModifiedBy": {
                "user": {
                  "email": "admin@contoso.OnMicrosoft.com",
                  "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
                  "displayName": "MOD Administrator"
                }
              },
              "parentReference": {
                "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
                "driveType": "documentLibrary",
                "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
                "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
              },
              "file": {
                "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "hashes": {
                  "quickXorHash": "4qr1R9xcSfoofaFA3nMY6qjQXPQ="
                }
              },
              "fileSystemInfo": {
                "createdDateTime": "2021-11-13T16:27:05Z",
                "lastModifiedDateTime": "2021-11-13T16:27:05Z"
              }
            }]
          });
        default:
          return Promise.reject(`Invalid GET request: ${url}`);
      }
    });

    command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/design',
        folderUrl: 'DemoDocs'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        assert(loggerLogSpy.calledWith([{
          "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=1bc151a6-beb8-4034-be93-6e9a18aa6cdc&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6Ik5PNFErS3pZSFBwMTJpZUhiZjdobmUrQ1E0em0yZVVvbnZCczI4RHdZVGs9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.d0hGUFpuMVhJbGZGSVFzcEhPSjJyS1FIUStXbEtLL3RURW9qVUhwZWNKWT0&ApiVersion=2.0",
          "createdDateTime": "2021-11-13T16:27:04Z",
          "eTag": "\"{1BC151A6-BEB8-4034-BE93-6E9A18AA6CDC},1\"",
          "id": "01472SYFFGKHARXOF6GRAL5E3OTIMKU3G4",
          "lastModifiedDateTime": "2021-11-13T16:27:04Z",
          "name": "Blog Post preview.docx",
          "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B1BC151A6-BEB8-4034-BE93-6E9A18AA6CDC%7D&file=Blog%20Post%20preview.docx&action=default&mobileredirect=true",
          "cTag": "\"c:{1BC151A6-BEB8-4034-BE93-6E9A18AA6CDC},4\"",
          "size": 131072,
          "createdBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "lastModifiedBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "parentReference": {
            "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
            "driveType": "documentLibrary",
            "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
            "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
          },
          "file": {
            "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "hashes": {
              "quickXorHash": "fVSMG9+U0AczPgoLQYCHu85LOT4="
            }
          },
          "fileSystemInfo": {
            "createdDateTime": "2021-11-13T16:27:04Z",
            "lastModifiedDateTime": "2021-11-13T16:27:04Z"
          },
          "lastModifiedByUser": "MOD Administrator"
        },
        {
          "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=716e4cb5-cefc-4562-97bb-b40a985f3306&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6InlrUjVQQjNjU3NIV1hhVXVwZW9qbGtna05UeW9VcFdweGxiS0FZbURwV1k9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.STJ5ekhaUEswQm1uL05Vb0F3dGhLcVJ4ZzNldkpReTM2YWI0cDlEaXo0VT0&ApiVersion=2.0",
          "createdDateTime": "2021-11-13T16:27:05Z",
          "eTag": "\"{716E4CB5-CEFC-4562-97BB-B40A985F3306},1\"",
          "id": "01472SYFFVJRXHD7GOMJCZPO5UBKMF6MYG",
          "lastModifiedDateTime": "2021-11-13T16:27:05Z",
          "name": "Contoso Purchasing Permissions.docx",
          "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B716E4CB5-CEFC-4562-97BB-B40A985F3306%7D&file=Contoso%20Purchasing%20Permissions.docx&action=default&mobileredirect=true",
          "cTag": "\"c:{716E4CB5-CEFC-4562-97BB-B40A985F3306},4\"",
          "size": 27442,
          "createdBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "lastModifiedBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "parentReference": {
            "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
            "driveType": "documentLibrary",
            "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
            "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
          },
          "file": {
            "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "hashes": {
              "quickXorHash": "ksnaYSdMJkWPE/EBL9rJWsK7kHw="
            }
          },
          "fileSystemInfo": {
            "createdDateTime": "2021-11-13T16:27:05Z",
            "lastModifiedDateTime": "2021-11-13T16:27:05Z"
          },
          "lastModifiedByUser": "MOD Administrator"
        },
        {
          "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=f123cf01-244e-44d3-b2f7-24f5230937a1&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6IkcxMS8vN2M1Y1RyU1JwVEFZOUJSSG1WVUlYYTRNTXdnKzJRcTVFZE1OZ1k9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.RWRBMjZvRFVUdlRkS3JsZmFWb2xDZnJLREtsZ1hGeVEvYWcrWStxV29rST0&ApiVersion=2.0",
          "createdDateTime": "2021-11-13T16:27:04Z",
          "eTag": "\"{F123CF01-244E-44D3-B2F7-24F5230937A1},1\"",
          "id": "01472SYFABZ4R7CTRE2NCLF5ZE6URQSN5B",
          "lastModifiedDateTime": "2021-11-13T16:27:04Z",
          "name": "Credit Cards.docx",
          "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7BF123CF01-244E-44D3-B2F7-24F5230937A1%7D&file=Credit%20Cards.docx&action=default&mobileredirect=true",
          "cTag": "\"c:{F123CF01-244E-44D3-B2F7-24F5230937A1},4\"",
          "size": 24858,
          "createdBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "lastModifiedBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "parentReference": {
            "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
            "driveType": "documentLibrary",
            "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
            "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
          },
          "file": {
            "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "hashes": {
              "quickXorHash": "9LEEpPZjBLjJQmN/AKDXRLmXj/k="
            }
          },
          "fileSystemInfo": {
            "createdDateTime": "2021-11-13T16:27:04Z",
            "lastModifiedDateTime": "2021-11-13T16:27:04Z"
          },
          "lastModifiedByUser": "MOD Administrator"
        },
        {
          "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=9b9a62cf-e7c6-4fa4-a261-1b260d7d6cf1&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6ImordllIUXB6d1hhM01HV3lJb3JuZWRzZlpXbGZIQXlsSmlsQTE2NU1HMms9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.ck96aFYwWCtSMnFrMVV4QVFNQklhSkMxYzRUczcwTzdxeWVsQnRHWjN1VT0&ApiVersion=2.0",
          "createdDateTime": "2021-11-13T16:27:05Z",
          "eTag": "\"{9B9A62CF-E7C6-4FA4-A261-1B260D7D6CF1},1\"",
          "id": "01472SYFGPMKNJXRXHURH2EYI3EYGX23HR",
          "lastModifiedDateTime": "2021-11-13T16:27:05Z",
          "name": "Customer Accounts.docx",
          "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B9B9A62CF-E7C6-4FA4-A261-1B260D7D6CF1%7D&file=Customer%20Accounts.docx&action=default&mobileredirect=true",
          "cTag": "\"c:{9B9A62CF-E7C6-4FA4-A261-1B260D7D6CF1},4\"",
          "size": 28360,
          "createdBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "lastModifiedBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "parentReference": {
            "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
            "driveType": "documentLibrary",
            "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
            "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
          },
          "file": {
            "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "hashes": {
              "quickXorHash": "PTFegvh8VBE1mJ3eELTSb+LSQxI="
            }
          },
          "fileSystemInfo": {
            "createdDateTime": "2021-11-13T16:27:05Z",
            "lastModifiedDateTime": "2021-11-13T16:27:05Z"
          },
          "lastModifiedByUser": "MOD Administrator"
        },
        {
          "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=e2f09ed9-987a-4d26-8513-86878f785e0a&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6IkRORDROcGlaQXAxa3RCR21HckRFa0ZScmdEQWVyeVE0TjgyK1dBNWJDaVE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.c2dqT3krQnBMOUx2QXdUMHlKWGtVbGtxMGZldmUyQWVTSFhPeERUR2pSVT0&ApiVersion=2.0",
          "createdDateTime": "2021-11-13T16:27:04Z",
          "eTag": "\"{E2F09ED9-987A-4D26-8513-86878F785E0A},1\"",
          "id": "01472SYFGZT3YOE6UYEZGYKE4GQ6HXQXQK",
          "lastModifiedDateTime": "2021-11-13T16:27:04Z",
          "name": "Customer Data.xlsx",
          "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7BE2F09ED9-987A-4D26-8513-86878F785E0A%7D&file=Customer%20Data.xlsx&action=default&mobileredirect=true",
          "cTag": "\"c:{E2F09ED9-987A-4D26-8513-86878F785E0A},4\"",
          "size": 17448,
          "createdBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "lastModifiedBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "parentReference": {
            "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
            "driveType": "documentLibrary",
            "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
            "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
          },
          "file": {
            "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "hashes": {
              "quickXorHash": "hrMOe12SPOwumxGnhdJiYnw1qoM="
            }
          },
          "fileSystemInfo": {
            "createdDateTime": "2021-11-13T16:27:04Z",
            "lastModifiedDateTime": "2021-11-13T16:27:04Z"
          },
          "lastModifiedByUser": "MOD Administrator"
        },
        {
          "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=370c45c8-cca6-4333-9cb4-d33b7bc3fc05&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6InAzSHU2TjRHSjRSTW9sREVJOVRLdE9DVzk0SXUzMllyTWdVak9IbCtlWkU9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.K1JabWc3TDNGWHpWN0Q5anluUTN6eUpYajNQNFY0NEJMUFhaS2VGTTgvTT0&ApiVersion=2.0",
          "createdDateTime": "2021-11-13T16:27:05Z",
          "eTag": "\"{370C45C8-CCA6-4333-9CB4-D33B7BC3FC05},1\"",
          "id": "01472SYFGIIUGDPJWMGNBZZNGTHN54H7AF",
          "lastModifiedDateTime": "2021-11-13T16:27:05Z",
          "name": "Q3 Sales and Marketing Expense Report Audit.pptx",
          "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B370C45C8-CCA6-4333-9CB4-D33B7BC3FC05%7D&file=Q3%20Sales%20and%20Marketing%20Expense%20Report%20Audit.pptx&action=edit&mobileredirect=true",
          "cTag": "\"c:{370C45C8-CCA6-4333-9CB4-D33B7BC3FC05},4\"",
          "size": 375938,
          "createdBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "lastModifiedBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "parentReference": {
            "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
            "driveType": "documentLibrary",
            "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
            "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
          },
          "file": {
            "mimeType": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            "hashes": {
              "quickXorHash": "/qGrwC7X0bUey5sxCDlzxzSjTLE="
            }
          },
          "fileSystemInfo": {
            "createdDateTime": "2021-11-13T16:27:05Z",
            "lastModifiedDateTime": "2021-11-13T16:27:05Z"
          },
          "lastModifiedByUser": "MOD Administrator"
        },
        {
          "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=2223f837-ce90-4638-a2ca-8fd753c4378d&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6InhuVmpzU1NWUWxTWVhWK3dDNElpVmtpMUIrZ0xTYWFpTldTaEpQemFnaFk9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.a0xuN1JrbXlnUWl2cW1rSHhPY1NNcnA4akhxVkVxSnVzQkNWeStLWGlPdz0&ApiVersion=2.0",
          "createdDateTime": "2021-11-13T16:27:04Z",
          "eTag": "\"{2223F837-CE90-4638-A2CA-8FD753C4378D},1\"",
          "id": "01472SYFBX7ARSFEGOHBDKFSUP25J4IN4N",
          "lastModifiedDateTime": "2021-11-13T16:27:04Z",
          "name": "Q3_Product_Strategy.docx",
          "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7B2223F837-CE90-4638-A2CA-8FD753C4378D%7D&file=Q3_Product_Strategy.docx&action=default&mobileredirect=true",
          "cTag": "\"c:{2223F837-CE90-4638-A2CA-8FD753C4378D},4\"",
          "size": 48086,
          "createdBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "lastModifiedBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "parentReference": {
            "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
            "driveType": "documentLibrary",
            "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
            "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
          },
          "file": {
            "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "hashes": {
              "quickXorHash": "LWe4j5GTM+Sg3PGOV0xoSSCnTr8="
            }
          },
          "fileSystemInfo": {
            "createdDateTime": "2021-11-13T16:27:04Z",
            "lastModifiedDateTime": "2021-11-13T16:27:04Z"
          },
          "lastModifiedByUser": "MOD Administrator"
        },
        {
          "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/download.aspx?UniqueId=bfc3a23e-ed0b-4a59-96f1-b11c075d89b0&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NjcwNTIuc2hhcmVwb2ludC5jb21ANzRhNTJiNGEtYTc5Ny00OGFkLTlkMGItMjQxYzliNjk1ZDFlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjgyMDgzOSIsImV4cCI6IjE2MzY4MjQ0MzkiLCJlbmRwb2ludHVybCI6IjVBMkhoblFOcEozMFlsZUp3NDU2ZENCTUlJNXlXbnRoaFhGVGZZQUZuMXM9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMzUiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IlpEWmlZamszT0RFdFpEZ3lOeTAwWXpjd0xXSXlabUl0TjJVeE9XRmtOemszTTJOaSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJZemc0WW1Fd09HWXROR1V4TnkwME9XTXhMV0V6TkdZdFkyRTROV1E1TURoak1qUmoiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJhcHBpZCI6IjMxMzU5YzdmLWJkN2UtNDc1Yy04NmRiLWZkYjhjOTM3NTQ4ZSIsInRpZCI6Ijc0YTUyYjRhLWE3OTctNDhhZC05ZDBiLTI0MWM5YjY5NWQxZSIsInVwbiI6ImFkbWluQG0zNjV4OTY3MDUyLm9ubWljcm9zb2Z0LmNvbSIsInB1aWQiOiIxMDAzMjAwMUEyQ0I4QjA1IiwiY2FjaGVrZXkiOiIwaC5mfG1lbWJlcnNoaXB8MTAwMzIwMDFhMmNiOGIwNUBsaXZlLmNvbSIsInNjcCI6ImFsbHNpdGVzLmZ1bGxjb250cm9sIGdyb3VwLndyaXRlIGFsbHByb2ZpbGVzLndyaXRlIHRlcm1zdG9yZS53cml0ZSIsInR0IjoiMiIsInVzZVBlcnNpc3RlbnRDb29raWUiOm51bGwsImlwYWRkciI6IjIwLjE5MC4xNjAuMjQifQ.MGd4U0wzSWR6S2grc3F2ODJHQncyRjY2RkkzR1FMdCtpSUo0Y1lhNjZtcz0&ApiVersion=2.0",
          "createdDateTime": "2021-11-13T16:27:05Z",
          "eTag": "\"{BFC3A23E-ED0B-4A59-96F1-B11C075D89B0},1\"",
          "id": "01472SYFB6ULB36C7NLFFJN4NRDQDV3CNQ",
          "lastModifiedDateTime": "2021-11-13T16:27:05Z",
          "name": "Sales Memo.docx",
          "webUrl": "https://contoso.sharepoint.com/sites/Design/_layouts/15/Doc.aspx?sourcedoc=%7BBFC3A23E-ED0B-4A59-96F1-B11C075D89B0%7D&file=Sales%20Memo.docx&action=default&mobileredirect=true",
          "cTag": "\"c:{BFC3A23E-ED0B-4A59-96F1-B11C075D89B0},4\"",
          "size": 35655,
          "createdBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "lastModifiedBy": {
            "user": {
              "email": "admin@contoso.OnMicrosoft.com",
              "id": "030dce8a-ab44-4cec-8f12-fa058b57209b",
              "displayName": "MOD Administrator"
            }
          },
          "parentReference": {
            "driveId": "b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3",
            "driveType": "documentLibrary",
            "id": "01472SYFF6Y2GOVW7725BZO354PWSELRRZ",
            "path": "/drives/b!j6CLyBdOwUmjT8qF2QjCTEwd5_2RrGREpg9jLZm4Ik3YB-53a_xKSazP19oz8Zw3/root:"
          },
          "file": {
            "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "hashes": {
              "quickXorHash": "4qr1R9xcSfoofaFA3nMY6qjQXPQ="
            }
          },
          "fileSystemInfo": {
            "createdDateTime": "2021-11-13T16:27:05Z",
            "lastModifiedDateTime": "2021-11-13T16:27:05Z"
          },
          "lastModifiedByUser": "MOD Administrator"
        }]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles file without last modified info', done => {
    sinon.stub(request, 'get').callsFake(opts => {
      const url: string = opts.url as string;

      switch (url) {
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/?$select=id':
          return Promise.resolve({
            "id": "contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42"
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42/drives?$select=webUrl,id':
          return Promise.resolve({
            "value": [
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYVs-s_Fc6EaRomQ91r_60hi",
                "webUrl": "https://contoso.sharepoint.com/Big%20list"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                "webUrl": "https://contoso.sharepoint.com/DemoDocs"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYUMhRqN81GDSYJHirLtImTh",
                "webUrl": "https://contoso.sharepoint.com/Shared%20Documents"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYWwU-RsQtl_RJxAlcRhJSYH",
                "webUrl": "https://contoso.sharepoint.com/JTDesignDocs"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYX2VXqL8cJvTbmPdTxwV8t4",
                "webUrl": "https://contoso.sharepoint.com/MCASDemoFiles"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYWaXP0JCVAyRZbch81buwYY",
                "webUrl": "https://contoso.sharepoint.com/RMSDemoLib"
              }
            ]
          });
        case 'https://graph.microsoft.com/v1.0/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root?$select=id':
          return Promise.resolve({
            "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ"
          });
        case 'https://graph.microsoft.com/v1.0/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/items/01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ/children':
          return Promise.resolve({
            "value": [
              {
                "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=b7b2648f-c50c-4d1e-bb99-9c1dcd5e37ae&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkNKTDB2UUp1SVF3enl1YUVHbDI1WVJpejJDdkM5bDdzRktVUFQzU0F2bTg9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.bmVGTE9nbDJ4c1lvem5sQ3pnTXRmb29HaTRUTWVVTlVaYXBkSVp3QWliRT0&ApiVersion=2.0",
                "createdDateTime": "2021-06-19T11:01:04Z",
                "eTag": "\"{B7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE},2\"",
                "id": "01YNDLPYMPMSZLODGFDZG3XGM4DXGV4N5O",
                "lastModifiedDateTime": "2021-06-19T11:01:04Z",
                "name": "Blog Post preview.docx",
                "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BB7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE%7D&file=Blog%20Post%20preview.docx&action=default&mobileredirect=true",
                "cTag": "\"c:{B7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE},3\"",
                "size": 134144,
                "createdBy": {
                  "user": {
                    "displayName": "Provisioning User"
                  }
                },
                "parentReference": {
                  "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                  "driveType": "documentLibrary",
                  "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
                  "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
                },
                "file": {
                  "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                  "hashes": {
                    "quickXorHash": "4wybUm/SrJtzssP/YDDKBwKRYd8="
                  }
                },
                "fileSystemInfo": {
                  "createdDateTime": "2021-06-19T11:01:04Z",
                  "lastModifiedDateTime": "2021-06-19T11:01:04Z"
                }
              }
            ]
          });
        default:
          return Promise.reject(`Invalid GET request: ${url}`);
      }
    });

    command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: 'DemoDocs'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        assert(loggerLogSpy.calledWith([
          {
            "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/_layouts/15/download.aspx?UniqueId=b7b2648f-c50c-4d1e-bb99-9c1dcd5e37ae&Translate=false&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXg5NTQ4MTAuc2hhcmVwb2ludC5jb21AMWIxMWY1MDItOWViMC00MDFhLWIxNjQtNjg5MzNlNmU5NDQzIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYzNjI5NDE4OCIsImV4cCI6IjE2MzYyOTc3ODgiLCJlbmRwb2ludHVybCI6IkNKTDB2UUp1SVF3enl1YUVHbDI1WVJpejJDdkM5bDdzRktVUFQzU0F2bTg9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjIiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik1ESXhZbVV4Tm1JdFpUYzVOaTAwTURKaExXRmtaalF0TXpJNU5UZGtZelJsTWpZMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJORFkxWXprM05UZ3RNREpsWXkwME9UbGxMVGhrTXpJdFptTmhPV0V6TURSaE1UazAiLCJhcHBfZGlzcGxheW5hbWUiOiJQblAgTWFuYWdlbWVudCBTaGVsbCIsImdpdmVuX25hbWUiOiJNT0QiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwiYXBwaWQiOiIzMTM1OWM3Zi1iZDdlLTQ3NWMtODZkYi1mZGI4YzkzNzU0OGUiLCJ0aWQiOiIxYjExZjUwMi05ZWIwLTQwMWEtYjE2NC02ODkzM2U2ZTk0NDMiLCJ1cG4iOiJhZG1pbkBtMzY1eDk1NDgxMC5vbm1pY3Jvc29mdC5jb20iLCJwdWlkIjoiMTAwMzIwMDE1M0Y5NjFBMiIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxNTNmOTYxYTJAbGl2ZS5jb20iLCJzY3AiOiJhbGxzaXRlcy5mdWxsY29udHJvbCBncm91cC53cml0ZSBhbGxwcm9maWxlcy53cml0ZSB0ZXJtc3RvcmUud3JpdGUiLCJ0dCI6IjIiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI0MC4xMjYuMzIuOTkifQ.bmVGTE9nbDJ4c1lvem5sQ3pnTXRmb29HaTRUTWVVTlVaYXBkSVp3QWliRT0&ApiVersion=2.0",
            "createdDateTime": "2021-06-19T11:01:04Z",
            "eTag": "\"{B7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE},2\"",
            "id": "01YNDLPYMPMSZLODGFDZG3XGM4DXGV4N5O",
            "lastModifiedDateTime": "2021-06-19T11:01:04Z",
            "name": "Blog Post preview.docx",
            "webUrl": "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BB7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE%7D&file=Blog%20Post%20preview.docx&action=default&mobileredirect=true",
            "cTag": "\"c:{B7B2648F-C50C-4D1E-BB99-9C1DCD5E37AE},3\"",
            "size": 134144,
            "createdBy": {
              "user": {
                "displayName": "Provisioning User"
              }
            },
            "parentReference": {
              "driveId": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
              "driveType": "documentLibrary",
              "id": "01YNDLPYN6Y2GOVW7725BZO354PWSELRRZ",
              "path": "/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:"
            },
            "file": {
              "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              "hashes": {
                "quickXorHash": "4wybUm/SrJtzssP/YDDKBwKRYd8="
              }
            },
            "fileSystemInfo": {
              "createdDateTime": "2021-06-19T11:01:04Z",
              "lastModifiedDateTime": "2021-06-19T11:01:04Z"
            },
            "lastModifiedByUser": undefined
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when site not found', done => {
    sinon.stub(request, 'get').callsFake(opts => {
      const url: string = opts.url as string;

      switch (url) {
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/design?$select=id':
          return Promise.reject({
            "error": {
              "code": "itemNotFound",
              "message": "Requested site could not be found",
              "innerError": {
                "date": "2021-11-14T14:38:46",
                "request-id": "81170202-c798-4ecc-ba13-bab07c0b27b7",
                "client-request-id": "81170202-c798-4ecc-ba13-bab07c0b27b7"
              }
            }
          });
        default:
          return Promise.reject(`Invalid GET request: ${url}`);
      }
    });

    command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/design',
        folderUrl: 'DemoDocs'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Requested site could not be found')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when document library not found in the root site without trailing slash', done => {
    sinon.stub(request, 'get').callsFake(opts => {
      const url: string = opts.url as string;

      switch (url) {
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/?$select=id':
          return Promise.resolve({
            "id": "contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42"
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42/drives?$select=webUrl,id':
          return Promise.resolve({
            "value": [
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYVs-s_Fc6EaRomQ91r_60hi",
                "webUrl": "https://contoso.sharepoint.com/Big%20list"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                "webUrl": "https://contoso.sharepoint.com/DemoDocs1"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYUMhRqN81GDSYJHirLtImTh",
                "webUrl": "https://contoso.sharepoint.com/Shared%20Documents"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYWwU-RsQtl_RJxAlcRhJSYH",
                "webUrl": "https://contoso.sharepoint.com/JTDesignDocs"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYX2VXqL8cJvTbmPdTxwV8t4",
                "webUrl": "https://contoso.sharepoint.com/MCASDemoFiles"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYWaXP0JCVAyRZbch81buwYY",
                "webUrl": "https://contoso.sharepoint.com/RMSDemoLib"
              }
            ]
          });
        default:
          return Promise.reject(`Invalid GET request: ${url}`);
      }
    });

    command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: 'DemoDocs'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Document library 'DemoDocs' not found`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when document library not found in the root site with trailing slash', done => {
    sinon.stub(request, 'get').callsFake(opts => {
      const url: string = opts.url as string;

      switch (url) {
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/?$select=id':
          return Promise.resolve({
            "id": "contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42"
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42/drives?$select=webUrl,id':
          return Promise.resolve({
            "value": [
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYVs-s_Fc6EaRomQ91r_60hi",
                "webUrl": "https://contoso.sharepoint.com/Big%20list"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                "webUrl": "https://contoso.sharepoint.com/DemoDocs1"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYUMhRqN81GDSYJHirLtImTh",
                "webUrl": "https://contoso.sharepoint.com/Shared%20Documents"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYWwU-RsQtl_RJxAlcRhJSYH",
                "webUrl": "https://contoso.sharepoint.com/JTDesignDocs"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYX2VXqL8cJvTbmPdTxwV8t4",
                "webUrl": "https://contoso.sharepoint.com/MCASDemoFiles"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYWaXP0JCVAyRZbch81buwYY",
                "webUrl": "https://contoso.sharepoint.com/RMSDemoLib"
              }
            ]
          });
        default:
          return Promise.reject(`Invalid GET request: ${url}`);
      }
    });

    command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/',
        folderUrl: 'DemoDocs'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Document library 'DemoDocs' not found`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when folder not found', done => {
    sinon.stub(request, 'get').callsFake(opts => {
      const url: string = opts.url as string;

      switch (url) {
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/?$select=id':
          return Promise.resolve({
            "id": "contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42"
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42/drives?$select=webUrl,id':
          return Promise.resolve({
            "value": [
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYVs-s_Fc6EaRomQ91r_60hi",
                "webUrl": "https://contoso.sharepoint.com/Big%20list"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                "webUrl": "https://contoso.sharepoint.com/DemoDocs"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYUMhRqN81GDSYJHirLtImTh",
                "webUrl": "https://contoso.sharepoint.com/Shared%20Documents"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYWwU-RsQtl_RJxAlcRhJSYH",
                "webUrl": "https://contoso.sharepoint.com/JTDesignDocs"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYX2VXqL8cJvTbmPdTxwV8t4",
                "webUrl": "https://contoso.sharepoint.com/MCASDemoFiles"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYWaXP0JCVAyRZbch81buwYY",
                "webUrl": "https://contoso.sharepoint.com/RMSDemoLib"
              }
            ]
          });
        case 'https://graph.microsoft.com/v1.0/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:/Fodler?$select=id':
          return Promise.reject({
            "error": {
              "code": "itemNotFound",
              "message": "The resource could not be found.",
              "innerError": {
                "date": "2021-11-14T14:58:14",
                "request-id": "d7d20190-ae3c-4d13-b096-5cc08640b3bf",
                "client-request-id": "d7d20190-ae3c-4d13-b096-5cc08640b3bf"
              }
            }
          });
        default:
          return Promise.reject(`Invalid GET request: ${url}`);
      }
    });

    command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: 'DemoDocs/Fodler'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('The resource could not be found.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when subfolder not found', done => {
    sinon.stub(request, 'get').callsFake(opts => {
      const url: string = opts.url as string;

      switch (url) {
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/?$select=id':
          return Promise.resolve({
            "id": "contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42"
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42/drives?$select=webUrl,id':
          return Promise.resolve({
            "value": [
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYVs-s_Fc6EaRomQ91r_60hi",
                "webUrl": "https://contoso.sharepoint.com/Big%20list"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT",
                "webUrl": "https://contoso.sharepoint.com/DemoDocs"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYUMhRqN81GDSYJHirLtImTh",
                "webUrl": "https://contoso.sharepoint.com/Shared%20Documents"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYWwU-RsQtl_RJxAlcRhJSYH",
                "webUrl": "https://contoso.sharepoint.com/JTDesignDocs"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYX2VXqL8cJvTbmPdTxwV8t4",
                "webUrl": "https://contoso.sharepoint.com/MCASDemoFiles"
              },
              {
                "id": "b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYWaXP0JCVAyRZbch81buwYY",
                "webUrl": "https://contoso.sharepoint.com/RMSDemoLib"
              }
            ]
          });
        case 'https://graph.microsoft.com/v1.0/drives/b!WJdcRuwCnkmNMvypowShlJAOO7sb8BNGi5bd40SvsYXCJjiTCgSgSq19j0OM3YgT/root:/Folder/Subfolder?$select=id':
          return Promise.reject({
            "error": {
              "code": "itemNotFound",
              "message": "The resource could not be found.",
              "innerError": {
                "date": "2021-11-14T14:58:14",
                "request-id": "d7d20190-ae3c-4d13-b096-5cc08640b3bf",
                "client-request-id": "d7d20190-ae3c-4d13-b096-5cc08640b3bf"
              }
            }
          });
        default:
          return Promise.reject(`Invalid GET request: ${url}`);
      }
    });

    command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: 'DemoDocs/Folder/Subfolder'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('The resource could not be found.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`fails validation if the specified webUrl is invalid`, async () => {
    const actual = await command.validate({
      options: {
        folderUrl: '/Shared Documents',
        webUrl: '/'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it(`passes validation if the target file is a URL`, async () => {
    const actual = await command.validate({
      options: {
        folderUrl: 'Shared Documents',
        webUrl: 'https://contoso.sharepoint.com/Shared Documents'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});