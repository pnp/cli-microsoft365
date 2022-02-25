import * as assert from 'assert';
import * as fs from 'fs';
import * as sinon from 'sinon';
import appInsights from '../../../appInsights';
import auth from '../../../Auth';
import { Logger } from '../../../cli';
import Command, { CommandError } from '../../../Command';
import request from '../../../request';
import { sinonUtil } from '../../../utils';
import commands from '../commands';
const command: Command = require('./file-add');

describe(commands.ADD, () => {
  let log: string[];
  let logger: Logger;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(fs, 'readFileSync').callsFake(() => 'abc');
    auth.service.connected = true;
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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      request.put,
      fs.existsSync,
      fs.readFileSync
    ]);
    (command as any).sourceFileGraphUrl = undefined;
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      fs.readFileSync
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('uploads file to the root site collection, root site, default document library, root folder', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      const url: string = opts.url as string;

      switch (url) {
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/?$select=id':
          return Promise.resolve({
            "id": "contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42"
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/Shared%20Documents?$select=id':
          return Promise.reject({
            "error": {
              "code": "itemNotFound",
              "message": "The provided path does not exist, or does not represent a site",
              "innerError": {
                "date": "2020-12-27T18:08:36",
                "request-id": "e44abe41-7f08-43d2-92b6-089e3335b47c",
                "client-request-id": "e44abe41-7f08-43d2-92b6-089e3335b47c"
              }
            }
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42/drives?$select=webUrl,id':
          return Promise.resolve({
            "value": [
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KAXP5NlvNnQLH92D7KrxA5",
                "webUrl": "https://contoso.sharepoint.com/DemoDocs"
              },
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU",
                "webUrl": "https://contoso.sharepoint.com/Shared%20Documents"
              },
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KCswD4M9qeR6qB9K5J5Kvp",
                "webUrl": "https://contoso.sharepoint.com/JTDesignDocs"
              },
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0LCxmZShRH-S4chwRsWoq23",
                "webUrl": "https://contoso.sharepoint.com/MCASDemoFiles"
              },
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0LxywkjzYwYSqUtcpywFv6S",
                "webUrl": "https://contoso.sharepoint.com/RMSDemoLib"
              }
            ]
          });
        default:
          return Promise.reject(`Invalid GET request: ${url}`);
      }
    });

    sinon.stub(fs, 'readFileSync').callsFake(() => 'abc');

    sinon.stub(request, 'post').callsFake(opts => {
      const url: string = opts.url as string;

      if (url === 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU/root:/file.pdf:/createUploadSession') {
        return Promise.resolve({
          "expirationDateTime": "2020-12-27T18:23:37.078Z",
          "nextExpectedRanges": [
            "0-"
          ],
          "uploadUrl": "https://contoso.sharepoint.com/_api/v2.0/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU/items/01AH65SIN6Y2GOVW7725BZO354PWSELRRZ/uploadSession?guid='19a2b995-5b72-4460-980a-a564ff63108c'&path='~tmpEF_file.pdf'&overwrite=True&rename=False&dc=0&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXgyNzE1MzQuc2hhcmVwb2ludC5jb21AZjczMjIzODAtZjIwMy00MmZmLTkzZTgtNjZlMjY2ZjZkMmU0IiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYwOTA5MjUxNyIsImV4cCI6IjE2MDkxNzg5MTciLCJlbmRwb2ludHVybCI6Ild5dUNlVWluMHBaQmUvTGI1WXQ1SDY2RGQzSDVzOFhZWUF6eU1KZ0VJcFE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIyNzMiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik4ySmlObUkyWldRdE9ETXhOQzAwTnpaaExXRmlPVEF0TVRjNVpHVTFZemxoWlRFMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJaV0UwT1dFek9UTXRaVE5sTmkwME56WXdMV0V4WWpJdFpUazJOVE01WlRFMU16Y3kiLCJhcHBfZGlzcGxheW5hbWUiOiJDTEkgdGVzdCIsIm5hbWVpZCI6IjgxYzZkODNhLWViYzYtNDM5Ni1hZTYwLTk1NDhiMmRlZTQ2ZEBmNzMyMjM4MC1mMjAzLTQyZmYtOTNlOC02NmUyNjZmNmQyZTQiLCJyb2xlcyI6ImFsbGZpbGVzLndyaXRlIiwidHQiOiIxIiwidXNlUGVyc2lzdGVudENvb2tpZSI6bnVsbH0.N2d0Tll4WFlqVWJmNWxnMHZTMjBaaEdJVXpUWC9NaDBrM1NRNlNYTXZzWT0"
        });
      }

      return Promise.reject(`Invalid POST request: ${url}`);
    });
    sinon.stub(request, 'put').callsFake(opts => {
      if (opts.url === `https://contoso.sharepoint.com/_api/v2.0/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU/items/01AH65SIN6Y2GOVW7725BZO354PWSELRRZ/uploadSession?guid='19a2b995-5b72-4460-980a-a564ff63108c'&path='~tmpEF_file.pdf'&overwrite=True&rename=False&dc=0&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXgyNzE1MzQuc2hhcmVwb2ludC5jb21AZjczMjIzODAtZjIwMy00MmZmLTkzZTgtNjZlMjY2ZjZkMmU0IiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYwOTA5MjUxNyIsImV4cCI6IjE2MDkxNzg5MTciLCJlbmRwb2ludHVybCI6Ild5dUNlVWluMHBaQmUvTGI1WXQ1SDY2RGQzSDVzOFhZWUF6eU1KZ0VJcFE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIyNzMiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik4ySmlObUkyWldRdE9ETXhOQzAwTnpaaExXRmlPVEF0TVRjNVpHVTFZemxoWlRFMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJaV0UwT1dFek9UTXRaVE5sTmkwME56WXdMV0V4WWpJdFpUazJOVE01WlRFMU16Y3kiLCJhcHBfZGlzcGxheW5hbWUiOiJDTEkgdGVzdCIsIm5hbWVpZCI6IjgxYzZkODNhLWViYzYtNDM5Ni1hZTYwLTk1NDhiMmRlZTQ2ZEBmNzMyMjM4MC1mMjAzLTQyZmYtOTNlOC02NmUyNjZmNmQyZTQiLCJyb2xlcyI6ImFsbGZpbGVzLndyaXRlIiwidHQiOiIxIiwidXNlUGVyc2lzdGVudENvb2tpZSI6bnVsbH0.N2d0Tll4WFlqVWJmNWxnMHZTMjBaaEdJVXpUWC9NaDBrM1NRNlNYTXZzWT0`) {
        return Promise.resolve({
          webUrl: "https://contoso.sharepoint.com/Shared%20Documents/file.pdf"
        });
      }

      return Promise.reject(`Invalid PUT request: ${opts}`);
    });

    command.action(logger, {
      options: {
        debug: true,
        filePath: 'file.pdf',
        folderUrl: 'https://contoso.sharepoint.com/Shared Documents'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('uploads empty file to the root site collection, root site, default document library, root folder', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      const url: string = opts.url as string;

      switch (url) {
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/?$select=id':
          return Promise.resolve({
            "id": "contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42"
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/Shared%20Documents?$select=id':
          return Promise.reject({
            "error": {
              "code": "itemNotFound",
              "message": "The provided path does not exist, or does not represent a site",
              "innerError": {
                "date": "2020-12-27T18:08:36",
                "request-id": "e44abe41-7f08-43d2-92b6-089e3335b47c",
                "client-request-id": "e44abe41-7f08-43d2-92b6-089e3335b47c"
              }
            }
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42/drives?$select=webUrl,id':
          return Promise.resolve({
            "value": [
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KAXP5NlvNnQLH92D7KrxA5",
                "webUrl": "https://contoso.sharepoint.com/DemoDocs"
              },
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU",
                "webUrl": "https://contoso.sharepoint.com/Shared%20Documents"
              },
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KCswD4M9qeR6qB9K5J5Kvp",
                "webUrl": "https://contoso.sharepoint.com/JTDesignDocs"
              },
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0LCxmZShRH-S4chwRsWoq23",
                "webUrl": "https://contoso.sharepoint.com/MCASDemoFiles"
              },
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0LxywkjzYwYSqUtcpywFv6S",
                "webUrl": "https://contoso.sharepoint.com/RMSDemoLib"
              }
            ]
          });
        default:
          return Promise.reject(`Invalid GET request: ${url}`);
      }
    });

    sinon.stub(fs, 'readFileSync').callsFake(() => '');

    sinon.stub(request, 'put').callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU/root:/file.pdf:/content`) {
        return Promise.resolve({
          webUrl: "https://contoso.sharepoint.com/Shared%20Documents/file.pdf"
        });
      }

      return Promise.reject(`Invalid PUT request: ${opts}`);
    });

    command.action(logger, {
      options: {
        debug: true,
        filePath: 'file.pdf',
        folderUrl: 'https://contoso.sharepoint.com/Shared Documents'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('uploads file to the root site collection, root site, default document library, root folder with trailing slash', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      const url: string = opts.url as string;

      switch (url) {
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/?$select=id':
          return Promise.resolve({
            "id": "contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42"
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/Shared%20Documents?$select=id':
          return Promise.reject({
            "error": {
              "code": "itemNotFound",
              "message": "The provided path does not exist, or does not represent a site",
              "innerError": {
                "date": "2020-12-27T18:08:36",
                "request-id": "e44abe41-7f08-43d2-92b6-089e3335b47c",
                "client-request-id": "e44abe41-7f08-43d2-92b6-089e3335b47c"
              }
            }
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42/drives?$select=webUrl,id':
          return Promise.resolve({
            "value": [
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KAXP5NlvNnQLH92D7KrxA5",
                "webUrl": "https://contoso.sharepoint.com/DemoDocs"
              },
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU",
                "webUrl": "https://contoso.sharepoint.com/Shared%20Documents"
              },
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KCswD4M9qeR6qB9K5J5Kvp",
                "webUrl": "https://contoso.sharepoint.com/JTDesignDocs"
              },
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0LCxmZShRH-S4chwRsWoq23",
                "webUrl": "https://contoso.sharepoint.com/MCASDemoFiles"
              },
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0LxywkjzYwYSqUtcpywFv6S",
                "webUrl": "https://contoso.sharepoint.com/RMSDemoLib"
              }
            ]
          });
        default:
          return Promise.reject(`Invalid GET request: ${url}`);
      }
    });

    sinon.stub(fs, 'readFileSync').callsFake(() => 'abc');

    sinon.stub(request, 'post').callsFake(opts => {
      const url: string = opts.url as string;

      if (url === 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU/root:/file.pdf:/createUploadSession') {
        return Promise.resolve({
          "expirationDateTime": "2020-12-27T18:23:37.078Z",
          "nextExpectedRanges": [
            "0-"
          ],
          "uploadUrl": "https://contoso.sharepoint.com/_api/v2.0/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU/items/01AH65SIN6Y2GOVW7725BZO354PWSELRRZ/uploadSession?guid='19a2b995-5b72-4460-980a-a564ff63108c'&path='~tmpEF_file.pdf'&overwrite=True&rename=False&dc=0&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXgyNzE1MzQuc2hhcmVwb2ludC5jb21AZjczMjIzODAtZjIwMy00MmZmLTkzZTgtNjZlMjY2ZjZkMmU0IiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYwOTA5MjUxNyIsImV4cCI6IjE2MDkxNzg5MTciLCJlbmRwb2ludHVybCI6Ild5dUNlVWluMHBaQmUvTGI1WXQ1SDY2RGQzSDVzOFhZWUF6eU1KZ0VJcFE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIyNzMiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik4ySmlObUkyWldRdE9ETXhOQzAwTnpaaExXRmlPVEF0TVRjNVpHVTFZemxoWlRFMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJaV0UwT1dFek9UTXRaVE5sTmkwME56WXdMV0V4WWpJdFpUazJOVE01WlRFMU16Y3kiLCJhcHBfZGlzcGxheW5hbWUiOiJDTEkgdGVzdCIsIm5hbWVpZCI6IjgxYzZkODNhLWViYzYtNDM5Ni1hZTYwLTk1NDhiMmRlZTQ2ZEBmNzMyMjM4MC1mMjAzLTQyZmYtOTNlOC02NmUyNjZmNmQyZTQiLCJyb2xlcyI6ImFsbGZpbGVzLndyaXRlIiwidHQiOiIxIiwidXNlUGVyc2lzdGVudENvb2tpZSI6bnVsbH0.N2d0Tll4WFlqVWJmNWxnMHZTMjBaaEdJVXpUWC9NaDBrM1NRNlNYTXZzWT0"
        });
      }

      return Promise.reject(`Invalid POST request: ${url}`);
    });
    sinon.stub(request, 'put').callsFake(opts => {
      if (opts.url === `https://contoso.sharepoint.com/_api/v2.0/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU/items/01AH65SIN6Y2GOVW7725BZO354PWSELRRZ/uploadSession?guid='19a2b995-5b72-4460-980a-a564ff63108c'&path='~tmpEF_file.pdf'&overwrite=True&rename=False&dc=0&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXgyNzE1MzQuc2hhcmVwb2ludC5jb21AZjczMjIzODAtZjIwMy00MmZmLTkzZTgtNjZlMjY2ZjZkMmU0IiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYwOTA5MjUxNyIsImV4cCI6IjE2MDkxNzg5MTciLCJlbmRwb2ludHVybCI6Ild5dUNlVWluMHBaQmUvTGI1WXQ1SDY2RGQzSDVzOFhZWUF6eU1KZ0VJcFE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIyNzMiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik4ySmlObUkyWldRdE9ETXhOQzAwTnpaaExXRmlPVEF0TVRjNVpHVTFZemxoWlRFMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJaV0UwT1dFek9UTXRaVE5sTmkwME56WXdMV0V4WWpJdFpUazJOVE01WlRFMU16Y3kiLCJhcHBfZGlzcGxheW5hbWUiOiJDTEkgdGVzdCIsIm5hbWVpZCI6IjgxYzZkODNhLWViYzYtNDM5Ni1hZTYwLTk1NDhiMmRlZTQ2ZEBmNzMyMjM4MC1mMjAzLTQyZmYtOTNlOC02NmUyNjZmNmQyZTQiLCJyb2xlcyI6ImFsbGZpbGVzLndyaXRlIiwidHQiOiIxIiwidXNlUGVyc2lzdGVudENvb2tpZSI6bnVsbH0.N2d0Tll4WFlqVWJmNWxnMHZTMjBaaEdJVXpUWC9NaDBrM1NRNlNYTXZzWT0`) {
        return Promise.resolve({
          webUrl: "https://contoso.sharepoint.com/Shared%20Documents/file.pdf"
        });
      }

      return Promise.reject(`Invalid PUT request: ${opts}`);
    });

    command.action(logger, {
      options: {
        debug: true,
        filePath: 'file.pdf',
        folderUrl: 'https://contoso.sharepoint.com/Shared Documents/'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('uploads file to the root site collection, root site, default document library, sub folder', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      const url: string = opts.url as string;

      switch (url) {
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/?$select=id':
          return Promise.resolve({
            "id": "contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42"
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/Shared%20Documents?$select=id':
          return Promise.reject({
            "error": {
              "code": "itemNotFound",
              "message": "The provided path does not exist, or does not represent a site",
              "innerError": {
                "date": "2020-12-27T18:08:36",
                "request-id": "e44abe41-7f08-43d2-92b6-089e3335b47c",
                "client-request-id": "e44abe41-7f08-43d2-92b6-089e3335b47c"
              }
            }
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42/drives?$select=webUrl,id':
          return Promise.resolve({
            "value": [
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KAXP5NlvNnQLH92D7KrxA5",
                "webUrl": "https://contoso.sharepoint.com/DemoDocs"
              },
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU",
                "webUrl": "https://contoso.sharepoint.com/Shared%20Documents"
              },
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KCswD4M9qeR6qB9K5J5Kvp",
                "webUrl": "https://contoso.sharepoint.com/JTDesignDocs"
              },
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0LCxmZShRH-S4chwRsWoq23",
                "webUrl": "https://contoso.sharepoint.com/MCASDemoFiles"
              },
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0LxywkjzYwYSqUtcpywFv6S",
                "webUrl": "https://contoso.sharepoint.com/RMSDemoLib"
              }
            ]
          });
        default:
          return Promise.reject(`Invalid GET request: ${url}`);
      }
    });

    sinon.stub(fs, 'readFileSync').callsFake(() => 'abc');

    sinon.stub(request, 'post').callsFake(opts => {
      const url: string = opts.url as string;

      if (url === 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU/root:/Folder/file.pdf:/createUploadSession') {
        return Promise.resolve({
          "expirationDateTime": "2020-12-27T18:23:37.078Z",
          "nextExpectedRanges": [
            "0-"
          ],
          "uploadUrl": "https://contoso.sharepoint.com/_api/v2.0/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU/items/01AH65SIN6Y2GOVW7725BZO354PWSELRRZ/uploadSession?guid='19a2b995-5b72-4460-980a-a564ff63108c'&path='~tmpEF_file.pdf'&overwrite=True&rename=False&dc=0&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXgyNzE1MzQuc2hhcmVwb2ludC5jb21AZjczMjIzODAtZjIwMy00MmZmLTkzZTgtNjZlMjY2ZjZkMmU0IiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYwOTA5MjUxNyIsImV4cCI6IjE2MDkxNzg5MTciLCJlbmRwb2ludHVybCI6Ild5dUNlVWluMHBaQmUvTGI1WXQ1SDY2RGQzSDVzOFhZWUF6eU1KZ0VJcFE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIyNzMiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik4ySmlObUkyWldRdE9ETXhOQzAwTnpaaExXRmlPVEF0TVRjNVpHVTFZemxoWlRFMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJaV0UwT1dFek9UTXRaVE5sTmkwME56WXdMV0V4WWpJdFpUazJOVE01WlRFMU16Y3kiLCJhcHBfZGlzcGxheW5hbWUiOiJDTEkgdGVzdCIsIm5hbWVpZCI6IjgxYzZkODNhLWViYzYtNDM5Ni1hZTYwLTk1NDhiMmRlZTQ2ZEBmNzMyMjM4MC1mMjAzLTQyZmYtOTNlOC02NmUyNjZmNmQyZTQiLCJyb2xlcyI6ImFsbGZpbGVzLndyaXRlIiwidHQiOiIxIiwidXNlUGVyc2lzdGVudENvb2tpZSI6bnVsbH0.N2d0Tll4WFlqVWJmNWxnMHZTMjBaaEdJVXpUWC9NaDBrM1NRNlNYTXZzWT0"
        });
      }

      return Promise.reject(`Invalid POST request: ${url}`);
    });
    sinon.stub(request, 'put').callsFake(opts => {
      if (opts.url === `https://contoso.sharepoint.com/_api/v2.0/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU/items/01AH65SIN6Y2GOVW7725BZO354PWSELRRZ/uploadSession?guid='19a2b995-5b72-4460-980a-a564ff63108c'&path='~tmpEF_file.pdf'&overwrite=True&rename=False&dc=0&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXgyNzE1MzQuc2hhcmVwb2ludC5jb21AZjczMjIzODAtZjIwMy00MmZmLTkzZTgtNjZlMjY2ZjZkMmU0IiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYwOTA5MjUxNyIsImV4cCI6IjE2MDkxNzg5MTciLCJlbmRwb2ludHVybCI6Ild5dUNlVWluMHBaQmUvTGI1WXQ1SDY2RGQzSDVzOFhZWUF6eU1KZ0VJcFE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIyNzMiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik4ySmlObUkyWldRdE9ETXhOQzAwTnpaaExXRmlPVEF0TVRjNVpHVTFZemxoWlRFMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJaV0UwT1dFek9UTXRaVE5sTmkwME56WXdMV0V4WWpJdFpUazJOVE01WlRFMU16Y3kiLCJhcHBfZGlzcGxheW5hbWUiOiJDTEkgdGVzdCIsIm5hbWVpZCI6IjgxYzZkODNhLWViYzYtNDM5Ni1hZTYwLTk1NDhiMmRlZTQ2ZEBmNzMyMjM4MC1mMjAzLTQyZmYtOTNlOC02NmUyNjZmNmQyZTQiLCJyb2xlcyI6ImFsbGZpbGVzLndyaXRlIiwidHQiOiIxIiwidXNlUGVyc2lzdGVudENvb2tpZSI6bnVsbH0.N2d0Tll4WFlqVWJmNWxnMHZTMjBaaEdJVXpUWC9NaDBrM1NRNlNYTXZzWT0`) {
        return Promise.resolve({
          webUrl: "https://contoso.sharepoint.com/Shared%20Documents/Folder/file.pdf"
        });
      }

      return Promise.reject(`Invalid PUT request: ${opts}`);
    });

    command.action(logger, {
      options: {
        debug: true,
        filePath: 'file.pdf',
        folderUrl: 'https://contoso.sharepoint.com/Shared Documents/Folder'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('uploads file to the root site collection, root site, custom document library, root folder', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      const url: string = opts.url as string;

      switch (url) {
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/?$select=id':
          return Promise.resolve({
            "id": "contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42"
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/DemoDocs?$select=id':
          return Promise.reject({
            "error": {
              "code": "itemNotFound",
              "message": "The provided path does not exist, or does not represent a site",
              "innerError": {
                "date": "2020-12-27T18:08:36",
                "request-id": "e44abe41-7f08-43d2-92b6-089e3335b47c",
                "client-request-id": "e44abe41-7f08-43d2-92b6-089e3335b47c"
              }
            }
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42/drives?$select=webUrl,id':
          return Promise.resolve({
            "value": [
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KAXP5NlvNnQLH92D7KrxA5",
                "webUrl": "https://contoso.sharepoint.com/DemoDocs"
              },
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU",
                "webUrl": "https://contoso.sharepoint.com/Shared%20Documents"
              },
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KCswD4M9qeR6qB9K5J5Kvp",
                "webUrl": "https://contoso.sharepoint.com/JTDesignDocs"
              },
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0LCxmZShRH-S4chwRsWoq23",
                "webUrl": "https://contoso.sharepoint.com/MCASDemoFiles"
              },
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0LxywkjzYwYSqUtcpywFv6S",
                "webUrl": "https://contoso.sharepoint.com/RMSDemoLib"
              }
            ]
          });
        default:
          return Promise.reject(`Invalid GET request: ${url}`);
      }
    });

    sinon.stub(fs, 'readFileSync').callsFake(() => 'abc');

    sinon.stub(request, 'post').callsFake(opts => {
      const url: string = opts.url as string;

      if (url === 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KAXP5NlvNnQLH92D7KrxA5/root:/file.pdf:/createUploadSession') {
        return Promise.resolve({
          "expirationDateTime": "2020-12-27T18:23:37.078Z",
          "nextExpectedRanges": [
            "0-"
          ],
          "uploadUrl": "https://contoso.sharepoint.com/_api/v2.0/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KAXP5NlvNnQLH92D7KrxA5/items/01AH65SIN6Y2GOVW7725BZO354PWSELRRZ/uploadSession?guid='19a2b995-5b72-4460-980a-a564ff63108c'&path='~tmpEF_file.pdf'&overwrite=True&rename=False&dc=0&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXgyNzE1MzQuc2hhcmVwb2ludC5jb21AZjczMjIzODAtZjIwMy00MmZmLTkzZTgtNjZlMjY2ZjZkMmU0IiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYwOTA5MjUxNyIsImV4cCI6IjE2MDkxNzg5MTciLCJlbmRwb2ludHVybCI6Ild5dUNlVWluMHBaQmUvTGI1WXQ1SDY2RGQzSDVzOFhZWUF6eU1KZ0VJcFE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIyNzMiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik4ySmlObUkyWldRdE9ETXhOQzAwTnpaaExXRmlPVEF0TVRjNVpHVTFZemxoWlRFMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJaV0UwT1dFek9UTXRaVE5sTmkwME56WXdMV0V4WWpJdFpUazJOVE01WlRFMU16Y3kiLCJhcHBfZGlzcGxheW5hbWUiOiJDTEkgdGVzdCIsIm5hbWVpZCI6IjgxYzZkODNhLWViYzYtNDM5Ni1hZTYwLTk1NDhiMmRlZTQ2ZEBmNzMyMjM4MC1mMjAzLTQyZmYtOTNlOC02NmUyNjZmNmQyZTQiLCJyb2xlcyI6ImFsbGZpbGVzLndyaXRlIiwidHQiOiIxIiwidXNlUGVyc2lzdGVudENvb2tpZSI6bnVsbH0.N2d0Tll4WFlqVWJmNWxnMHZTMjBaaEdJVXpUWC9NaDBrM1NRNlNYTXZzWT0"
        });
      }

      return Promise.reject(`Invalid POST request: ${url}`);
    });
    sinon.stub(request, 'put').callsFake(opts => {
      if (opts.url === `https://contoso.sharepoint.com/_api/v2.0/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KAXP5NlvNnQLH92D7KrxA5/items/01AH65SIN6Y2GOVW7725BZO354PWSELRRZ/uploadSession?guid='19a2b995-5b72-4460-980a-a564ff63108c'&path='~tmpEF_file.pdf'&overwrite=True&rename=False&dc=0&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXgyNzE1MzQuc2hhcmVwb2ludC5jb21AZjczMjIzODAtZjIwMy00MmZmLTkzZTgtNjZlMjY2ZjZkMmU0IiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYwOTA5MjUxNyIsImV4cCI6IjE2MDkxNzg5MTciLCJlbmRwb2ludHVybCI6Ild5dUNlVWluMHBaQmUvTGI1WXQ1SDY2RGQzSDVzOFhZWUF6eU1KZ0VJcFE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIyNzMiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik4ySmlObUkyWldRdE9ETXhOQzAwTnpaaExXRmlPVEF0TVRjNVpHVTFZemxoWlRFMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJaV0UwT1dFek9UTXRaVE5sTmkwME56WXdMV0V4WWpJdFpUazJOVE01WlRFMU16Y3kiLCJhcHBfZGlzcGxheW5hbWUiOiJDTEkgdGVzdCIsIm5hbWVpZCI6IjgxYzZkODNhLWViYzYtNDM5Ni1hZTYwLTk1NDhiMmRlZTQ2ZEBmNzMyMjM4MC1mMjAzLTQyZmYtOTNlOC02NmUyNjZmNmQyZTQiLCJyb2xlcyI6ImFsbGZpbGVzLndyaXRlIiwidHQiOiIxIiwidXNlUGVyc2lzdGVudENvb2tpZSI6bnVsbH0.N2d0Tll4WFlqVWJmNWxnMHZTMjBaaEdJVXpUWC9NaDBrM1NRNlNYTXZzWT0`) {
        return Promise.resolve({
          webUrl: "https://contoso.sharepoint.com/DemoDocs/file.pdf"
        });
      }

      return Promise.reject(`Invalid PUT request: ${opts}`);
    });

    command.action(logger, {
      options: {
        debug: true,
        filePath: 'file.pdf',
        folderUrl: 'https://contoso.sharepoint.com/DemoDocs'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('uploads file to One Drive for Business, default doc lib, root folder', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      const url: string = opts.url as string;

      switch (url) {
        case 'https://graph.microsoft.com/v1.0/sites/contoso-my.sharepoint.com:/?$select=id':
          return Promise.resolve({
            "id": "contoso-my.sharepoint.com,0c452457-5819-46d5-b676-422b0d77ef13,250cd3fe-13b2-43a8-aa6c-c706122adf88"
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso-my.sharepoint.com:/personal/steve_contoso_com?$select=id':
          return Promise.resolve({
            "id": "contoso-my.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,7c130ce6-4b35-47cc-a07c-e251175696ef"
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso-my.sharepoint.com:/personal/steve_contoso_com/Documents?$select=id':
          return Promise.reject({
            "error": {
              "code": "itemNotFound",
              "message": "The provided path does not exist, or does not represent a site",
              "innerError": {
                "date": "2020-12-25T14:38:23",
                "request-id": "3bb31418-e1d4-48a8-9abe-f74ce9d1fff6",
                "client-request-id": "3bb31418-e1d4-48a8-9abe-f74ce9d1fff6"
              }
            }
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso-my.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,7c130ce6-4b35-47cc-a07c-e251175696ef/drives?$select=webUrl,id':
          return Promise.resolve({
            "value": [
              {
                "id": "b!910I76DSwUGO4gdQ5LIwxA-_eGhZ0MhHqzcnffK9MY7oZnn6NbBJT7qm_AaWHNyv",
                "webUrl": "https://contoso-my.sharepoint.com/personal/steve_contoso_com/Documents"
              }
            ]
          });
        default:
          return Promise.reject(`Invalid GET request: ${url}`);
      }
    });

    sinon.stub(fs, 'readFileSync').callsFake(() => 'abc');

    sinon.stub(request, 'post').callsFake(opts => {
      const url: string = opts.url as string;

      if (url === 'https://graph.microsoft.com/v1.0/sites/contoso-my.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,7c130ce6-4b35-47cc-a07c-e251175696ef/drives/b!910I76DSwUGO4gdQ5LIwxA-_eGhZ0MhHqzcnffK9MY7oZnn6NbBJT7qm_AaWHNyv/root:/file.pdf:/createUploadSession') {
        return Promise.resolve({
          "expirationDateTime": "2020-12-27T18:23:37.078Z",
          "nextExpectedRanges": [
            "0-"
          ],
          "uploadUrl": "https://contoso-my.sharepoint.com/_api/v2.0/drives/b!910I76DSwUGO4gdQ5LIwxA-_eGhZ0MhHqzcnffK9MY7oZnn6NbBJT7qm_AaWHNyv/items/01AH65SIN6Y2GOVW7725BZO354PWSELRRZ/uploadSession?guid='19a2b995-5b72-4460-980a-a564ff63108c'&path='~tmpEF_file.pdf'&overwrite=True&rename=False&dc=0&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXgyNzE1MzQuc2hhcmVwb2ludC5jb21AZjczMjIzODAtZjIwMy00MmZmLTkzZTgtNjZlMjY2ZjZkMmU0IiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYwOTA5MjUxNyIsImV4cCI6IjE2MDkxNzg5MTciLCJlbmRwb2ludHVybCI6Ild5dUNlVWluMHBaQmUvTGI1WXQ1SDY2RGQzSDVzOFhZWUF6eU1KZ0VJcFE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIyNzMiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik4ySmlObUkyWldRdE9ETXhOQzAwTnpaaExXRmlPVEF0TVRjNVpHVTFZemxoWlRFMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJaV0UwT1dFek9UTXRaVE5sTmkwME56WXdMV0V4WWpJdFpUazJOVE01WlRFMU16Y3kiLCJhcHBfZGlzcGxheW5hbWUiOiJDTEkgdGVzdCIsIm5hbWVpZCI6IjgxYzZkODNhLWViYzYtNDM5Ni1hZTYwLTk1NDhiMmRlZTQ2ZEBmNzMyMjM4MC1mMjAzLTQyZmYtOTNlOC02NmUyNjZmNmQyZTQiLCJyb2xlcyI6ImFsbGZpbGVzLndyaXRlIiwidHQiOiIxIiwidXNlUGVyc2lzdGVudENvb2tpZSI6bnVsbH0.N2d0Tll4WFlqVWJmNWxnMHZTMjBaaEdJVXpUWC9NaDBrM1NRNlNYTXZzWT0"
        });
      }

      return Promise.reject(`Invalid POST request: ${url}`);
    });
    sinon.stub(request, 'put').callsFake(opts => {
      if (opts.url === `https://contoso-my.sharepoint.com/_api/v2.0/drives/b!910I76DSwUGO4gdQ5LIwxA-_eGhZ0MhHqzcnffK9MY7oZnn6NbBJT7qm_AaWHNyv/items/01AH65SIN6Y2GOVW7725BZO354PWSELRRZ/uploadSession?guid='19a2b995-5b72-4460-980a-a564ff63108c'&path='~tmpEF_file.pdf'&overwrite=True&rename=False&dc=0&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXgyNzE1MzQuc2hhcmVwb2ludC5jb21AZjczMjIzODAtZjIwMy00MmZmLTkzZTgtNjZlMjY2ZjZkMmU0IiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYwOTA5MjUxNyIsImV4cCI6IjE2MDkxNzg5MTciLCJlbmRwb2ludHVybCI6Ild5dUNlVWluMHBaQmUvTGI1WXQ1SDY2RGQzSDVzOFhZWUF6eU1KZ0VJcFE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIyNzMiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik4ySmlObUkyWldRdE9ETXhOQzAwTnpaaExXRmlPVEF0TVRjNVpHVTFZemxoWlRFMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJaV0UwT1dFek9UTXRaVE5sTmkwME56WXdMV0V4WWpJdFpUazJOVE01WlRFMU16Y3kiLCJhcHBfZGlzcGxheW5hbWUiOiJDTEkgdGVzdCIsIm5hbWVpZCI6IjgxYzZkODNhLWViYzYtNDM5Ni1hZTYwLTk1NDhiMmRlZTQ2ZEBmNzMyMjM4MC1mMjAzLTQyZmYtOTNlOC02NmUyNjZmNmQyZTQiLCJyb2xlcyI6ImFsbGZpbGVzLndyaXRlIiwidHQiOiIxIiwidXNlUGVyc2lzdGVudENvb2tpZSI6bnVsbH0.N2d0Tll4WFlqVWJmNWxnMHZTMjBaaEdJVXpUWC9NaDBrM1NRNlNYTXZzWT0`) {
        return Promise.resolve({
          webUrl: "https://contoso-my.sharepoint.com/personal/steve_contoso_com/Documents/file.pdf"
        });
      }

      return Promise.reject(`Invalid PUT request: ${opts}`);
    });

    command.action(logger, {
      options: {
        debug: false,
        filePath: 'file.pdf',
        folderUrl: 'https://contoso-my.sharepoint.com/personal/steve_contoso_com/Documents'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('uploads file to a non-root site collection, doc lib, root folder', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      const url: string = opts.url as string;

      switch (url) {
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/?$select=id':
          return Promise.resolve({
            "id": "contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42"
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/Contoso?$select=id':
          return Promise.resolve({
            "id": "contoso.sharepoint.com,9d1b2174-9906-43ec-8c9e-f8589de047af,bf674ab6-4b20-4368-8516-d71e6002d4b9"
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/Contoso/Shared%20Documents?$select=id':
          return Promise.reject({
            "error": {
              "code": "itemNotFound",
              "message": "The provided path does not exist, or does not represent a site",
              "innerError": {
                "date": "2020-12-25T18:21:34",
                "request-id": "6b24a926-4018-4279-a66a-f5a1ab7f8181",
                "client-request-id": "6b24a926-4018-4279-a66a-f5a1ab7f8181"
              }
            }
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,9d1b2174-9906-43ec-8c9e-f8589de047af,bf674ab6-4b20-4368-8516-d71e6002d4b9/drives?$select=webUrl,id':
          return Promise.resolve({
            "value": [
              {
                "id": "b!dCEbnQaZ7EOMnvhYneBHr7ZKZ78gS2hDhRbXHmAC1LnkVKXD20dsSYInKHJxx08q",
                "webUrl": "https://contoso.sharepoint.com/sites/Contoso/Shared%20Documents"
              }
            ]
          });
        default:
          return Promise.reject(`Invalid GET request: ${url}`);
      }
    });

    sinon.stub(fs, 'readFileSync').callsFake(() => 'abc');

    sinon.stub(request, 'post').callsFake(opts => {
      const url: string = opts.url as string;

      if (url === 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,9d1b2174-9906-43ec-8c9e-f8589de047af,bf674ab6-4b20-4368-8516-d71e6002d4b9/drives/b!dCEbnQaZ7EOMnvhYneBHr7ZKZ78gS2hDhRbXHmAC1LnkVKXD20dsSYInKHJxx08q/root:/file.pdf:/createUploadSession') {
        return Promise.resolve({
          "expirationDateTime": "2020-12-27T18:23:37.078Z",
          "nextExpectedRanges": [
            "0-"
          ],
          "uploadUrl": "https://contoso.sharepoint.com/_api/v2.0/drives/b!dCEbnQaZ7EOMnvhYneBHr7ZKZ78gS2hDhRbXHmAC1LnkVKXD20dsSYInKHJxx08q/items/01AH65SIN6Y2GOVW7725BZO354PWSELRRZ/uploadSession?guid='19a2b995-5b72-4460-980a-a564ff63108c'&path='~tmpEF_file.pdf'&overwrite=True&rename=False&dc=0&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXgyNzE1MzQuc2hhcmVwb2ludC5jb21AZjczMjIzODAtZjIwMy00MmZmLTkzZTgtNjZlMjY2ZjZkMmU0IiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYwOTA5MjUxNyIsImV4cCI6IjE2MDkxNzg5MTciLCJlbmRwb2ludHVybCI6Ild5dUNlVWluMHBaQmUvTGI1WXQ1SDY2RGQzSDVzOFhZWUF6eU1KZ0VJcFE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIyNzMiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik4ySmlObUkyWldRdE9ETXhOQzAwTnpaaExXRmlPVEF0TVRjNVpHVTFZemxoWlRFMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJaV0UwT1dFek9UTXRaVE5sTmkwME56WXdMV0V4WWpJdFpUazJOVE01WlRFMU16Y3kiLCJhcHBfZGlzcGxheW5hbWUiOiJDTEkgdGVzdCIsIm5hbWVpZCI6IjgxYzZkODNhLWViYzYtNDM5Ni1hZTYwLTk1NDhiMmRlZTQ2ZEBmNzMyMjM4MC1mMjAzLTQyZmYtOTNlOC02NmUyNjZmNmQyZTQiLCJyb2xlcyI6ImFsbGZpbGVzLndyaXRlIiwidHQiOiIxIiwidXNlUGVyc2lzdGVudENvb2tpZSI6bnVsbH0.N2d0Tll4WFlqVWJmNWxnMHZTMjBaaEdJVXpUWC9NaDBrM1NRNlNYTXZzWT0"
        });
      }

      return Promise.reject(`Invalid POST request: ${url}`);
    });
    sinon.stub(request, 'put').callsFake(opts => {
      if (opts.url === `https://contoso.sharepoint.com/_api/v2.0/drives/b!dCEbnQaZ7EOMnvhYneBHr7ZKZ78gS2hDhRbXHmAC1LnkVKXD20dsSYInKHJxx08q/items/01AH65SIN6Y2GOVW7725BZO354PWSELRRZ/uploadSession?guid='19a2b995-5b72-4460-980a-a564ff63108c'&path='~tmpEF_file.pdf'&overwrite=True&rename=False&dc=0&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXgyNzE1MzQuc2hhcmVwb2ludC5jb21AZjczMjIzODAtZjIwMy00MmZmLTkzZTgtNjZlMjY2ZjZkMmU0IiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYwOTA5MjUxNyIsImV4cCI6IjE2MDkxNzg5MTciLCJlbmRwb2ludHVybCI6Ild5dUNlVWluMHBaQmUvTGI1WXQ1SDY2RGQzSDVzOFhZWUF6eU1KZ0VJcFE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIyNzMiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik4ySmlObUkyWldRdE9ETXhOQzAwTnpaaExXRmlPVEF0TVRjNVpHVTFZemxoWlRFMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJaV0UwT1dFek9UTXRaVE5sTmkwME56WXdMV0V4WWpJdFpUazJOVE01WlRFMU16Y3kiLCJhcHBfZGlzcGxheW5hbWUiOiJDTEkgdGVzdCIsIm5hbWVpZCI6IjgxYzZkODNhLWViYzYtNDM5Ni1hZTYwLTk1NDhiMmRlZTQ2ZEBmNzMyMjM4MC1mMjAzLTQyZmYtOTNlOC02NmUyNjZmNmQyZTQiLCJyb2xlcyI6ImFsbGZpbGVzLndyaXRlIiwidHQiOiIxIiwidXNlUGVyc2lzdGVudENvb2tpZSI6bnVsbH0.N2d0Tll4WFlqVWJmNWxnMHZTMjBaaEdJVXpUWC9NaDBrM1NRNlNYTXZzWT0`) {
        return Promise.resolve({
          webUrl: "https://contoso.sharepoint.com/sites/Contoso/Shared Documents/file.pdf"
        });
      }

      return Promise.reject(`Invalid PUT request: ${opts}`);
    });

    command.action(logger, {
      options: {
        debug: false,
        filePath: 'file.pdf',
        folderUrl: 'https://contoso.sharepoint.com/sites/Contoso/Shared Documents'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', JSON.stringify(err));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('uploads file to a non-root site collection, doc lib, root folder without site lookup with siteUrl specified', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      const url: string = opts.url as string;

      switch (url) {
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/Contoso?$select=id':
          return Promise.resolve({
            "id": "contoso.sharepoint.com,9d1b2174-9906-43ec-8c9e-f8589de047af,bf674ab6-4b20-4368-8516-d71e6002d4b9"
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,9d1b2174-9906-43ec-8c9e-f8589de047af,bf674ab6-4b20-4368-8516-d71e6002d4b9/drives?$select=webUrl,id':
          return Promise.resolve({
            "value": [
              {
                "id": "b!dCEbnQaZ7EOMnvhYneBHr7ZKZ78gS2hDhRbXHmAC1LnkVKXD20dsSYInKHJxx08q",
                "webUrl": "https://contoso.sharepoint.com/sites/Contoso/Shared%20Documents"
              }
            ]
          });
        default:
          return Promise.reject(`Invalid GET request: ${url}`);
      }
    });

    sinon.stub(fs, 'readFileSync').callsFake(() => 'abc');

    sinon.stub(request, 'post').callsFake(opts => {
      const url: string = opts.url as string;

      if (url === 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,9d1b2174-9906-43ec-8c9e-f8589de047af,bf674ab6-4b20-4368-8516-d71e6002d4b9/drives/b!dCEbnQaZ7EOMnvhYneBHr7ZKZ78gS2hDhRbXHmAC1LnkVKXD20dsSYInKHJxx08q/root:/file.pdf:/createUploadSession') {
        return Promise.resolve({
          "expirationDateTime": "2020-12-27T18:23:37.078Z",
          "nextExpectedRanges": [
            "0-"
          ],
          "uploadUrl": "https://contoso.sharepoint.com/_api/v2.0/drives/b!dCEbnQaZ7EOMnvhYneBHr7ZKZ78gS2hDhRbXHmAC1LnkVKXD20dsSYInKHJxx08q/items/01AH65SIN6Y2GOVW7725BZO354PWSELRRZ/uploadSession?guid='19a2b995-5b72-4460-980a-a564ff63108c'&path='~tmpEF_file.pdf'&overwrite=True&rename=False&dc=0&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXgyNzE1MzQuc2hhcmVwb2ludC5jb21AZjczMjIzODAtZjIwMy00MmZmLTkzZTgtNjZlMjY2ZjZkMmU0IiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYwOTA5MjUxNyIsImV4cCI6IjE2MDkxNzg5MTciLCJlbmRwb2ludHVybCI6Ild5dUNlVWluMHBaQmUvTGI1WXQ1SDY2RGQzSDVzOFhZWUF6eU1KZ0VJcFE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIyNzMiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik4ySmlObUkyWldRdE9ETXhOQzAwTnpaaExXRmlPVEF0TVRjNVpHVTFZemxoWlRFMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJaV0UwT1dFek9UTXRaVE5sTmkwME56WXdMV0V4WWpJdFpUazJOVE01WlRFMU16Y3kiLCJhcHBfZGlzcGxheW5hbWUiOiJDTEkgdGVzdCIsIm5hbWVpZCI6IjgxYzZkODNhLWViYzYtNDM5Ni1hZTYwLTk1NDhiMmRlZTQ2ZEBmNzMyMjM4MC1mMjAzLTQyZmYtOTNlOC02NmUyNjZmNmQyZTQiLCJyb2xlcyI6ImFsbGZpbGVzLndyaXRlIiwidHQiOiIxIiwidXNlUGVyc2lzdGVudENvb2tpZSI6bnVsbH0.N2d0Tll4WFlqVWJmNWxnMHZTMjBaaEdJVXpUWC9NaDBrM1NRNlNYTXZzWT0"
        });
      }

      return Promise.reject(`Invalid POST request: ${url}`);
    });
    sinon.stub(request, 'put').callsFake(opts => {
      if (opts.url === `https://contoso.sharepoint.com/_api/v2.0/drives/b!dCEbnQaZ7EOMnvhYneBHr7ZKZ78gS2hDhRbXHmAC1LnkVKXD20dsSYInKHJxx08q/items/01AH65SIN6Y2GOVW7725BZO354PWSELRRZ/uploadSession?guid='19a2b995-5b72-4460-980a-a564ff63108c'&path='~tmpEF_file.pdf'&overwrite=True&rename=False&dc=0&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXgyNzE1MzQuc2hhcmVwb2ludC5jb21AZjczMjIzODAtZjIwMy00MmZmLTkzZTgtNjZlMjY2ZjZkMmU0IiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYwOTA5MjUxNyIsImV4cCI6IjE2MDkxNzg5MTciLCJlbmRwb2ludHVybCI6Ild5dUNlVWluMHBaQmUvTGI1WXQ1SDY2RGQzSDVzOFhZWUF6eU1KZ0VJcFE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIyNzMiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik4ySmlObUkyWldRdE9ETXhOQzAwTnpaaExXRmlPVEF0TVRjNVpHVTFZemxoWlRFMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJaV0UwT1dFek9UTXRaVE5sTmkwME56WXdMV0V4WWpJdFpUazJOVE01WlRFMU16Y3kiLCJhcHBfZGlzcGxheW5hbWUiOiJDTEkgdGVzdCIsIm5hbWVpZCI6IjgxYzZkODNhLWViYzYtNDM5Ni1hZTYwLTk1NDhiMmRlZTQ2ZEBmNzMyMjM4MC1mMjAzLTQyZmYtOTNlOC02NmUyNjZmNmQyZTQiLCJyb2xlcyI6ImFsbGZpbGVzLndyaXRlIiwidHQiOiIxIiwidXNlUGVyc2lzdGVudENvb2tpZSI6bnVsbH0.N2d0Tll4WFlqVWJmNWxnMHZTMjBaaEdJVXpUWC9NaDBrM1NRNlNYTXZzWT0`) {
        return Promise.resolve({
          webUrl: "https://contoso.sharepoint.com/sites/Contoso/Shared Documents/file.pdf"
        });
      }

      return Promise.reject(`Invalid PUT request: ${opts}`);
    });

    command.action(logger, {
      options: {
        debug: false,
        filePath: 'file.pdf',
        folderUrl: 'https://contoso.sharepoint.com/sites/Contoso/Shared Documents',
        siteUrl: 'https://contoso.sharepoint.com/sites/Contoso'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', JSON.stringify(err));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`returns error when the specified document library doesn't exist`, (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      const url: string = opts.url as string;

      switch (url) {
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/?$select=id':
          return Promise.resolve({
            "id": "contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42"
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/Shared%20Documents?$select=id':
          return Promise.reject({
            "error": {
              "code": "itemNotFound",
              "message": "The provided path does not exist, or does not represent a site",
              "innerError": {
                "date": "2020-12-25T14:38:23",
                "request-id": "3bb31418-e1d4-48a8-9abe-f74ce9d1fff6",
                "client-request-id": "3bb31418-e1d4-48a8-9abe-f74ce9d1fff6"
              }
            }
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42/drives?$select=webUrl,id':
          return Promise.resolve({
            "value": [
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KAXP5NlvNnQLH92D7KrxA5",
                "webUrl": "https://contoso.sharepoint.com/DemoDocs"
              },
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU",
                "webUrl": "https://contoso.sharepoint.com/Shared%20Documents"
              },
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KCswD4M9qeR6qB9K5J5Kvp",
                "webUrl": "https://contoso.sharepoint.com/JTDesignDocs"
              },
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0LCxmZShRH-S4chwRsWoq23",
                "webUrl": "https://contoso.sharepoint.com/MCASDemoFiles"
              },
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0LxywkjzYwYSqUtcpywFv6S",
                "webUrl": "https://contoso.sharepoint.com/RMSDemoLib"
              }
            ]
          });
        default:
          return Promise.reject(`Invalid GET request: ${url}`);
      }
    });
    sinon.stub(fs, 'readFileSync').callsFake(() => 'abc');
    sinon.stub(request, 'post').callsFake(_ => Promise.reject('Issued POST request'));
    sinon.stub(request, 'put').callsFake(_ => Promise.reject('Issued PUT request'));

    command.action(logger, {
      options: {
        debug: true,
        filePath: 'file.pdf',
        folderUrl: 'https://contoso.sharepoint.com/Docs'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Drive not found')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns error when resolving Graph URL for the file to be uploaded failed', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      const url: string = opts.url as string;

      switch (url) {
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/?$select=id':
          return Promise.reject({
            "error": {
              "message": "An error has occurred"
            }
          });
        default:
          return Promise.reject(`Invalid GET request: ${url}`);
      }
    });
    sinon.stub(fs, 'readFileSync').callsFake(() => 'abc');
    sinon.stub(request, 'post').callsFake(_ => Promise.reject('Issued POST request'));
    sinon.stub(request, 'put').callsFake(_ => Promise.reject('Issued PUT request'));

    command.action(logger, {
      options: {
        debug: false,
        filePath: 'file.pdf',
        folderUrl: 'https://contoso.sharepoint.com/Shared Documents'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns error when creating Graph upload session for the file to be uploaded failed', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      const url: string = opts.url as string;

      switch (url) {
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/?$select=id':
          return Promise.resolve({
            "id": "contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42"
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/Shared%20Documents?$select=id':
          return Promise.reject({
            "error": {
              "code": "itemNotFound",
              "message": "The provided path does not exist, or does not represent a site",
              "innerError": {
                "date": "2020-12-25T14:38:23",
                "request-id": "3bb31418-e1d4-48a8-9abe-f74ce9d1fff6",
                "client-request-id": "3bb31418-e1d4-48a8-9abe-f74ce9d1fff6"
              }
            }
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42/drives?$select=webUrl,id':
          return Promise.resolve({
            "value": [
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KAXP5NlvNnQLH92D7KrxA5",
                "webUrl": "https://contoso.sharepoint.com/DemoDocs"
              },
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU",
                "webUrl": "https://contoso.sharepoint.com/Shared%20Documents"
              },
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KCswD4M9qeR6qB9K5J5Kvp",
                "webUrl": "https://contoso.sharepoint.com/JTDesignDocs"
              },
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0LCxmZShRH-S4chwRsWoq23",
                "webUrl": "https://contoso.sharepoint.com/MCASDemoFiles"
              },
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0LxywkjzYwYSqUtcpywFv6S",
                "webUrl": "https://contoso.sharepoint.com/RMSDemoLib"
              }
            ]
          });
        default:
          return Promise.reject(`Invalid GET request: ${url}`);
      }
    });
    sinon.stub(fs, 'readFileSync').callsFake(() => 'abc');
    sinon.stub(request, 'post').callsFake(_ => Promise.reject({
      "error": {
        "message": "An error has occurred"
      }
    }));
    sinon.stub(request, 'put').callsFake(_ => Promise.reject('Issued PUT request'));

    command.action(logger, {
      options: {
        debug: false,
        filePath: 'https://contoso.sharepoint.com/Shared Documents/file.pdf',
        folderUrl: 'https://contoso.sharepoint.com/Shared Documents/file.pdf'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns error when uploading the file failed', (done) => {
    sinon.stub(request, 'get').callsFake(opts => {
      const url: string = opts.url as string;

      switch (url) {
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/?$select=id':
          return Promise.resolve({
            "id": "contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42"
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/Shared%20Documents?$select=id':
          return Promise.reject({
            "error": {
              "code": "itemNotFound",
              "message": "The provided path does not exist, or does not represent a site",
              "innerError": {
                "date": "2020-12-25T14:38:23",
                "request-id": "3bb31418-e1d4-48a8-9abe-f74ce9d1fff6",
                "client-request-id": "3bb31418-e1d4-48a8-9abe-f74ce9d1fff6"
              }
            }
          });
        case 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42/drives?$select=webUrl,id':
          return Promise.resolve({
            "value": [
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KAXP5NlvNnQLH92D7KrxA5",
                "webUrl": "https://contoso.sharepoint.com/DemoDocs"
              },
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU",
                "webUrl": "https://contoso.sharepoint.com/Shared%20Documents"
              },
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KCswD4M9qeR6qB9K5J5Kvp",
                "webUrl": "https://contoso.sharepoint.com/JTDesignDocs"
              },
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0LCxmZShRH-S4chwRsWoq23",
                "webUrl": "https://contoso.sharepoint.com/MCASDemoFiles"
              },
              {
                "id": "b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0LxywkjzYwYSqUtcpywFv6S",
                "webUrl": "https://contoso.sharepoint.com/RMSDemoLib"
              }
            ]
          });
        default:
          return Promise.reject(`Invalid GET request: ${url}`);
      }
    });
    sinon.stub(fs, 'readFileSync').callsFake(() => 'abc');
    sinon.stub(request, 'post').callsFake(opts => {
      const url: string = opts.url as string;
      if (url.startsWith('https://graph.microsoft.com/v1.0/drive/root:/') &&
        url.endsWith(':/createUploadSession')) {
        return Promise.resolve({
          "expirationDateTime": "2020-12-27T13:36:41.895Z",
          "nextExpectedRanges": [
            "0-"
          ],
          "uploadUrl": "https://contoso.sharepoint.com/_api/v2.0/drive/items/01AH65SIN6Y2GOVW7725BZO354PWSELRRZ/uploadSession?guid='12da75d9-3bb5-45b3-9145-3587993b1b34'&path='~tmp66_7f17e1b7-b40c-4259-9859-c94e7f2bd1c8.docx'&overwrite=True&rename=False&dc=0&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXgyNzE1MzQuc2hhcmVwb2ludC5jb21AZjczMjIzODAtZjIwMy00MmZmLTkzZTgtNjZlMjY2ZjZkMmU0IiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYwOTA3NTMwMiIsImV4cCI6IjE2MDkxNjE3MDIiLCJlbmRwb2ludHVybCI6IjVhRjUvSWxwOTFKTkhFVHhvOWU3ekJHcmw0a1hRZ1lEbmdpR0dubDVVRlU9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIyMzgiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik56ZzRPV1k0WkdNdE5UZ3lOeTAwTm1GbUxUZzBNMlF0WmpnMk1HVXpZelJrTXpFeiIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJaV0UwT1dFek9UTXRaVE5sTmkwME56WXdMV0V4WWpJdFpUazJOVE01WlRFMU16Y3kiLCJhcHBfZGlzcGxheW5hbWUiOiJDTEkgdGVzdCIsIm5hbWVpZCI6IjgxYzZkODNhLWViYzYtNDM5Ni1hZTYwLTk1NDhiMmRlZTQ2ZEBmNzMyMjM4MC1mMjAzLTQyZmYtOTNlOC02NmUyNjZmNmQyZTQiLCJyb2xlcyI6ImFsbGZpbGVzLndyaXRlIiwidHQiOiIxIiwidXNlUGVyc2lzdGVudENvb2tpZSI6bnVsbH0.cUhHUFVOSHNZR2lFellqalpJQ2R2SUpOMjl5d3RpY0g0WHZYQXVUMmtIaz0"
        });
      }

      if (url === 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU/root:/file.pdf:/createUploadSession') {
        return Promise.resolve({
          "expirationDateTime": "2020-12-27T18:23:37.078Z",
          "nextExpectedRanges": [
            "0-"
          ],
          "uploadUrl": "https://contoso.sharepoint.com/_api/v2.0/drives/b!k6NJ6ubjYEehsullOeFTchyG4mbZlhhEp1wO0bymi0KkhVdx52mJQ5y68EfLYQYU/items/01AH65SIN6Y2GOVW7725BZO354PWSELRRZ/uploadSession?guid='19a2b995-5b72-4460-980a-a564ff63108c'&path='~tmpEF_file.pdf'&overwrite=True&rename=False&dc=0&tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvbTM2NXgyNzE1MzQuc2hhcmVwb2ludC5jb21AZjczMjIzODAtZjIwMy00MmZmLTkzZTgtNjZlMjY2ZjZkMmU0IiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTYwOTA5MjUxNyIsImV4cCI6IjE2MDkxNzg5MTciLCJlbmRwb2ludHVybCI6Ild5dUNlVWluMHBaQmUvTGI1WXQ1SDY2RGQzSDVzOFhZWUF6eU1KZ0VJcFE9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIyNzMiLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6Ik4ySmlObUkyWldRdE9ETXhOQzAwTnpaaExXRmlPVEF0TVRjNVpHVTFZemxoWlRFMCIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJaV0UwT1dFek9UTXRaVE5sTmkwME56WXdMV0V4WWpJdFpUazJOVE01WlRFMU16Y3kiLCJhcHBfZGlzcGxheW5hbWUiOiJDTEkgdGVzdCIsIm5hbWVpZCI6IjgxYzZkODNhLWViYzYtNDM5Ni1hZTYwLTk1NDhiMmRlZTQ2ZEBmNzMyMjM4MC1mMjAzLTQyZmYtOTNlOC02NmUyNjZmNmQyZTQiLCJyb2xlcyI6ImFsbGZpbGVzLndyaXRlIiwidHQiOiIxIiwidXNlUGVyc2lzdGVudENvb2tpZSI6bnVsbH0.N2d0Tll4WFlqVWJmNWxnMHZTMjBaaEdJVXpUWC9NaDBrM1NRNlNYTXZzWT0"
        });
      }

      return Promise.reject(`Invalid POST request: ${url}`);
    });
    sinon.stub(request, 'put').callsFake(_ => Promise.reject({
      "error": {
        "message": "An error has occurred"
      }
    }));

    command.action(logger, {
      options: {
        debug: false,
        filePath: 'file.pdf',
        folderUrl: 'https://contoso.sharepoint.com/Shared Documents'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`fails validation if the specified local source file doesn't exist`, () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const actual = command.validate({ options: { filePath: 'file.pdf', folderUrl: 'https://contoso.sharepoint.com/Shared Documents' } });
    assert.notStrictEqual(actual, true);
  });

  it(`fails validation if the specified siteUrl is invalid`, () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    const actual = command.validate({ options: {
      filePath: 'file.pdf',
      folderUrl: 'https://contoso.sharepoint.com/Shared Documents',
      siteUrl: '/'
    } });
    assert.notStrictEqual(actual, true);
  });

  it(`passes validation if the target file is a URL`, () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    const actual = command.validate({ options: { filePath: 'file.pdf', folderUrl: 'https://contoso.sharepoint.com/Shared Documents' } });
    assert.strictEqual(actual, true);
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
