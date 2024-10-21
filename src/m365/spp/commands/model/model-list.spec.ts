import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './model-list.js';

describe(commands.MODEL_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  const models = [
    {
      "AIBuilderHybridModelType": null,
      "AzureCognitivePrebuiltModelName": null,
      "BaseContentTypeName": null,
      "ConfidenceScore": "{\"trainingStatus\":{\"kind\":\"original\",\"ClassifierStatus\":{\"TrainingStatus\":\"success\",\"TimeStamp\":1716547860981},\"ExtractorsStatus\":[{\"TimeStamp\":1716547860173,\"ExtractorName\":\"Name\",\"TrainingStatus\":\"success\"}]},\"modelAccuracy\":{\"Classifier\":1,\"Extractors\":{\"Name\":0.333333343}},\"perSampleAccuracy\":{\"1\":{\"Classifier\":1,\"Extractors\":{\"Name\":1}},\"2\":{\"Classifier\":1,\"Extractors\":{\"Name\":0}},\"3\":{\"Classifier\":1,\"Extractors\":{\"Name\":0}},\"4\":{\"Classifier\":1,\"Extractors\":{\"Name\":0}},\"5\":{\"Classifier\":1,\"Extractors\":{\"Name\":0}},\"6\":{\"Classifier\":1,\"Extractors\":{\"Name\":1}}},\"perSamplePrediction\":{\"1\":{\"Extractors\":{\"Name\":[\"Michał\"]}},\"2\":{\"Extractors\":{\"Name\":[]}},\"3\":{\"Extractors\":{\"Name\":[]}},\"4\":{\"Extractors\":{\"Name\":[]}},\"5\":{\"Extractors\":{\"Name\":[]}},\"6\":{\"Extractors\":{\"Name\":[]}}},\"trainingFailures\":{}}",
      "ContentTypeGroup": "Intelligent Document Content Types",
      "ContentTypeId": "0x010100A5C3671D1FB1A64D9F280C628D041692",
      "ContentTypeName": "TeachingModel",
      "Created": "2024-05-23T16:51:29Z",
      "CreatedBy": "i:0#.f|membership|user@contoso.onmicrosoft.com",
      "DriveId": "b!qTVLltt1P02LUgejOy6O_1amoFeu1EJBlawH83UtYbQs_H3KVKAcQpuQOpNLl646",
      "Explanations": "{\"Classifier\":[{\"id\":\"950f39cd-5e72-4442-ae8a-ea4bae32962c\",\"kind\":\"regexFeature\",\"name\":\"Email address\",\"active\":true,\"pattern\":\"[A-Za-z0-9._%-]+@[A-Za-z0-9.-]+.[A-Za-z]{2,6}\"},{\"id\":\"d3f2940d-1df1-4ba8-975a-db3a4d626d5c\",\"kind\":\"dictionaryFeature\",\"name\":\"FirstName\",\"active\":true,\"nGrams\":[\"Michał\"],\"caseSensitive\":false,\"ignoreDigitIdentity\":false,\"ignoreLetterIdentity\":false},{\"id\":\"077966e1-73be-44c9-855f-a4eade6a280b\",\"kind\":\"modelFeature\",\"name\":\"Name\",\"active\":true,\"modelReference\":\"Name\",\"conceptId\":\"309b64e5-acd5-4538-a5b6-c6bfcdc1ffbf\"}],\"Extractors\":{\"Name\":[{\"id\":\"69e412bc-e5bc-4657-b378-34a01966bb92\",\"kind\":\"dictionaryFeature\",\"name\":\"Before label\",\"active\":true,\"nGrams\":[\"Test\",\"'Surname\",\"Test (\"],\"caseSensitive\":false,\"ignoreDigitIdentity\":false,\"ignoreLetterIdentity\":false}]}}",
      "ID": 1,
      "LastTrained": "2024-05-24T17:51:16Z",
      "ListID": "ca7dfc2c-a054-421c-9b90-3a934b97ae3a",
      "ModelSettings": "{\"ModelOrigin\":{\"LibraryId\":\"8a6027ab-c584-4394-ba9c-3dc4dd152b65\",\"Published\":false,\"SelecteFileUniqueIds\":[]}}",
      "ModelType": 2,
      "Modified": "2024-05-24T17:51:02Z",
      "ModifiedBy": "i:0#.f|membership|user@contoso.onmicrosoft.com",
      "ObjectId": "01HQDCWVGIEBDRN3RVK5A3UJW3M4TCMT45",
      "PublicationType": 0,
      "Schemas": "{\"Version\":2,\"Extractors\":{\"Name\":{\"concepts\":{\"309b64e5-acd5-4538-a5b6-c6bfcdc1ffbf\":{\"name\":\"Name\"}},\"relationships\":[],\"id\":\"Name\"}}}",
      "SourceSiteUrl": "https://contoso.sharepoint.com/sites/SyntexTest",
      "SourceUrl": null,
      "SourceWebServerRelativeUrl": "/sites/SyntexTest",
      "UniqueId": "164720c8-35ee-4157-ba26-db6726264f9d"
    }
  ];

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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.MODEL_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['AIBuilderHybridModelType', 'ContentTypeName', 'LastTrained', 'UniqueId']);
  });

  it('passes validation when required parameters are valid', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/sales' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when siteUrl is not valid', async () => {
    const actual = await command.validate({ options: { siteUrl: 'invalidUrl' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('correctly handles site is not Content Site', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web?$select=WebTemplateConfiguration`) {
        return {
          WebTemplateConfiguration: 'SITEPAGEPUBLISHING#0'
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { verbose: true, siteUrl: 'https://contoso.sharepoint.com/sites/portal' } }),
      new CommandError('https://contoso.sharepoint.com/sites/portal is not a content site.'));
  });


  it('correctly handles an access denied error', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web?$select=WebTemplateConfiguration`) {
        throw {
          error: {
            "odata.error": {
              message: {
                lang: "en-US",
                value: "Attempted to perform an unauthorized operation."
              }
            }
          }
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { verbose: true, siteUrl: 'https://contoso.sharepoint.com/sites/portal' } }),
      new CommandError('Attempted to perform an unauthorized operation.'));
  });


  it('retrieves all site models', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web?$select=WebTemplateConfiguration`) {
        return {
          WebTemplateConfiguration: 'CONTENTCTR#0'
        };
      }

      if (opts.url === 'https://contoso.sharepoint.com/sites/portal/_api/machinelearning/models') {
        return { value: models };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/portal' } });
    assert(loggerLogSpy.calledOnceWithExactly(models));
  });

  it('gets correct model list when the site URL has a trailing slash', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/portal/_api/web?$select=WebTemplateConfiguration`) {
        return {
          WebTemplateConfiguration: 'CONTENTCTR#0'
        };
      }

      if (opts.url === 'https://contoso.sharepoint.com/sites/portal/_api/machinelearning/models') {
        return { value: models };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/portal/' } });
    assert(loggerLogSpy.calledOnceWithExactly(models));
  });
});
